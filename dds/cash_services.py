from decimal import Decimal, ROUND_HALF_UP
from django.core.exceptions import ValidationError
from django.db import transaction
from django.utils import timezone

from .models import CashRegister, CashMovement, CashTransfer, Point

FIELD_MAP = {
    CashMovement.ACC_CASH: "cash_balance",
    CashMovement.ACC_MKASSA: "mkassa_balance",
    CashMovement.ACC_ZADATOK: "zadatok_balance",
    CashMovement.ACC_OPTIMA: "optima_balance",
}


class CashTransferError(Exception):
    pass


def _to_decimal(x) -> Decimal:
    try:
        return Decimal(str(x))
    except Exception:
        raise ValidationError("Некорректная сумма.")


@transaction.atomic
def apply_cash_movement(
    *,
    hotel: Point,
    account: str,
    direction: str,
    amount,
    created_by,
    happened_at=None,
    comment="",
    dds_operation=None,
    incasso=None,
    transfer=None,
):
    """
    Создаёт CashMovement и обновляет CashRegister.
    ВАЖНО: обновляем через объект reg (а не через .update(F())),
    чтобы обновлялся updated_at и не было “визуально не меняется”.
    """
    happened_at = happened_at or timezone.now()
    amount = _to_decimal(amount)

    if amount <= 0:
        raise ValidationError("Сумма должна быть больше 0.")

    field = FIELD_MAP.get(account)
    if not field:
        raise ValidationError(f"Неизвестный счет: {account}")

    # берём/создаём кассу
    reg, _ = CashRegister.objects.get_or_create(hotel=hotel)
    reg = CashRegister.objects.select_for_update().get(pk=reg.pk)

    current = getattr(reg, field) or Decimal("0.00")

    # расчёт нового баланса
    if direction == CashMovement.IN:
        new_value = current + amount
    elif direction == CashMovement.OUT:
        if amount > current:
            raise ValidationError(f"Недостаточно средств на счете {account}. Доступно: {current}")
        new_value = current - amount
    else:
        raise ValidationError("Неверное направление движения.")

    # ✅ создаём движение
    move = CashMovement.objects.create(
        register=reg,
        hotel=hotel,
        account=account,
        direction=direction,
        amount=amount,
        happened_at=happened_at,
        comment=comment or "",
        created_by=created_by,
        dds_operation=dds_operation,
        incasso=incasso,
        transfer=transfer,
    )

    # ✅ обновляем баланс + updated_at
    setattr(reg, field, new_value)
    reg.save(update_fields=[field, "updated_at"])

    # ✅ обновляем общую кассу
    from .models import GlobalCashRegister
    gcr = GlobalCashRegister.objects.select_for_update().get_or_create(pk=1)[0]
    gcr_current = getattr(gcr, field) or Decimal("0")
    if direction == CashMovement.IN:
        setattr(gcr, field, gcr_current + amount)
    else:
        setattr(gcr, field, max(Decimal("0"), gcr_current - amount))
    gcr.save(update_fields=[field, "updated_at"])

    return move


@transaction.atomic
def transfer_between_accounts(
    *,
    hotel: Point,
    from_account: str,
    to_account: str,
    amount,
    user,
    happened_at=None,
    comment="",
) -> CashTransfer:
    """
    Внутренний перевод между счетами одного отеля:
    OPTIMA -> CASH, MKASSA -> OPTIMA и т.д.
    Создаёт CashTransfer + 2 CashMovement (OUT/IN) и обновляет кассу.
    """
    happened_at = happened_at or timezone.now()
    amount = _to_decimal(amount)

    if from_account == to_account:
        raise CashTransferError("Нельзя переводить на тот же самый счет.")
    if amount <= 0:
        raise CashTransferError("Сумма должна быть больше 0.")

    # касса под блокировкой
    reg, _ = CashRegister.objects.get_or_create(hotel=hotel)
    reg = CashRegister.objects.select_for_update().get(pk=reg.pk)

    from_field = FIELD_MAP.get(from_account)
    to_field = FIELD_MAP.get(to_account)
    if not from_field or not to_field:
        raise CashTransferError("Неверный счет.")

    from_balance = getattr(reg, from_field) or Decimal("0.00")
    if from_balance < amount:
        raise CashTransferError(
            f"Недостаточно средств на счете '{from_account}'. Баланс: {from_balance}"
        )

    # создаём перевод
    transfer = CashTransfer.objects.create(
        hotel=hotel,
        register=reg,
        from_account=from_account,
        to_account=to_account,
        amount=amount,
        happened_at=happened_at,
        comment=comment or "",
        created_by=user,
    )

    # ✅ делаем 2 движения через общий сервис (и касса обновится корректно)
    apply_cash_movement(
        hotel=hotel,
        account=from_account,
        direction=CashMovement.OUT,
        amount=amount,
        created_by=user,
        happened_at=happened_at,
        comment=f"Перевод на {to_account}. {comment}".strip(),
        transfer=transfer,
    )

    apply_cash_movement(
        hotel=hotel,
        account=to_account,
        direction=CashMovement.IN,
        amount=amount,
        created_by=user,
        happened_at=happened_at,
        comment=f"Перевод с {from_account}. {comment}".strip(),
        transfer=transfer,
    )

    return transfer


def _distribute_amount(total: Decimal, points_balances: list) -> list:
    """
    Делит total среди точек пропорционально равными долями.
    Если у точки не хватает — берётся всё что есть, дефицит
    перераспределяется между оставшимися точками итеративно.

    points_balances: [(point, balance), ...]
    Возвращает: [(point, amount, note), ...]
    """
    result = {}   # point.id -> Decimal
    notes  = {}   # point.id -> str

    remaining = total
    pool = list(points_balances)   # [(point, available_balance)]

    while remaining > Decimal("0.01") and pool:
        per = (remaining / len(pool)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        next_pool = []
        covered = Decimal("0")

        for point, avail in pool:
            pay = min(per, avail)
            result[point.id] = result.get(point.id, Decimal("0")) + pay
            covered += pay
            leftover = avail - pay
            if leftover > Decimal("0"):
                next_pool.append((point, leftover))
            else:
                if pay < per:
                    notes[point.id] = "недостаточно средств"

        remaining = remaining - covered
        if not next_pool:
            break
        pool = next_pool

    return [
        (p, result.get(p.id, Decimal("0")), notes.get(p.id, ""))
        for p, _ in points_balances
    ]


@transaction.atomic
def global_cash_income(*, account: str, amount, comment: str, created_by, happened_at=None):
    """Пополнение общей кассы."""
    from .models import GlobalCashRegister, GlobalCashOperation

    amount = _to_decimal(amount)
    if amount <= 0:
        raise ValidationError("Сумма должна быть больше 0.")

    happened_at = happened_at or timezone.now()
    field = FIELD_MAP.get(account)
    if not field:
        raise ValidationError(f"Неизвестный счёт: {account}")

    gcr = GlobalCashRegister.objects.select_for_update().get_or_create(pk=1)[0]
    setattr(gcr, field, (getattr(gcr, field) or Decimal("0")) + amount)
    gcr.save(update_fields=[field, "updated_at"])

    op = GlobalCashOperation.objects.create(
        direction=GlobalCashOperation.IN,
        account=account,
        amount=amount,
        happened_at=happened_at,
        comment=comment,
        created_by=created_by,
    )
    return op


@transaction.atomic
def global_cash_expense(*, account: str, amount, comment: str, created_by, happened_at=None, points_qs=None):
    """
    Расход из общей кассы с распределением по точкам.
    Деньги списываются с GlobalCashRegister и с касс точек.
    Если у точки не хватает — дефицит перекладывается на другие.
    """
    from .models import GlobalCashRegister, GlobalCashOperation, GlobalCashDistribution

    amount = _to_decimal(amount)
    if amount <= 0:
        raise ValidationError("Сумма должна быть больше 0.")

    happened_at = happened_at or timezone.now()
    field = FIELD_MAP.get(account)
    if not field:
        raise ValidationError(f"Неизвестный счёт: {account}")

    gcr = GlobalCashRegister.objects.select_for_update().get_or_create(pk=1)[0]
    gcr_balance = getattr(gcr, field) or Decimal("0")
    if amount > gcr_balance:
        raise ValidationError(f"Недостаточно средств в общей кассе. Доступно: {gcr_balance}")

    # Собираем активные точки с их балансами по нужному счёту
    if points_qs is None:
        points_qs = Point.objects.filter(is_active=True)

    points_list = list(points_qs)
    if not points_list:
        raise ValidationError("Нет активных точек для распределения.")

    # Получаем кассы точек под блокировкой
    regs = {
        r.hotel_id: r
        for r in CashRegister.objects.select_for_update().filter(hotel__in=points_list)
    }
    # у кого нет кассы — создаём
    for p in points_list:
        if p.id not in regs:
            r, _ = CashRegister.objects.get_or_create(hotel=p)
            regs[p.id] = r

    points_balances = [
        (p, getattr(regs[p.id], field) or Decimal("0"))
        for p in points_list
    ]

    distributions = _distribute_amount(amount, points_balances)

    # Списываем с общей кассы
    setattr(gcr, field, gcr_balance - amount)
    gcr.save(update_fields=[field, "updated_at"])

    # Создаём операцию
    op = GlobalCashOperation.objects.create(
        direction=GlobalCashOperation.OUT,
        account=account,
        amount=amount,
        happened_at=happened_at,
        comment=comment,
        created_by=created_by,
    )

    # Списываем с касс точек и сохраняем распределение
    for point, dist_amount, note in distributions:
        if dist_amount <= 0:
            continue
        reg = regs[point.id]
        current = getattr(reg, field) or Decimal("0")
        setattr(reg, field, current - dist_amount)
        reg.save(update_fields=[field, "updated_at"])

        GlobalCashDistribution.objects.create(
            operation=op,
            point=point,
            amount=dist_amount,
            note=note,
        )

    return op
