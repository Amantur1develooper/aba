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


ACCOUNT_LABELS = {
    "cash": "Наличные", "mkassa": "Банк1", "zadatok": "Задаток", "optima": "Банк2",
}


class CashTransferError(Exception):
    pass


def _notify_global(*, icon, label, account, amount, comment, created_by,
                   happened_at, article_name=None, distributions=None):
    """Send Telegram notification for global cash operations (non-blocking)."""
    try:
        from .telegram import notify_transaction
        from accounts.models import Profile
        chat_ids = list(Profile.objects.filter(tg_chat_id__gt="").values_list("tg_chat_id", flat=True))
        if not chat_ids:
            return
        lines = [
            f"{icon} <b>{label}</b>",
            f"💰 {float(amount):,.0f} сом  |  {ACCOUNT_LABELS.get(account, account)}",
        ]
        if article_name:
            lines.append(f"📋 {article_name}")
        if distributions:
            pts = "  ".join(f"{p.name}: {float(a):,.0f}" for p, a in distributions)
            lines.append(f"📊 {pts}")
        if comment:
            lines.append(f"💬 {comment}")
        lines.append(f"👤 {created_by.get_full_name() or created_by.username}")
        lines.append(f"🕐 {happened_at.strftime('%d.%m.%Y %H:%M')}")
        try:
            from .models import GlobalCashRegister
            gcr = GlobalCashRegister.objects.get(pk=1)
            lines += ["─────────────────", f"🏛 <b>Общая касса:</b> {float(gcr.total):,.0f} сом"]
        except Exception:
            pass
        notify_transaction(chat_ids, "\n".join(lines))
    except Exception:
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

    _notify_global(
        icon="🟢", label="Пополнение общей кассы",
        account=account, amount=amount, comment=comment,
        created_by=created_by, happened_at=happened_at,
    )
    return op


@transaction.atomic
def global_cash_expense(*, account: str, amount, comment: str, created_by,
                        happened_at=None, points_qs=None, article=None):
    """
    Расход из общей кассы с равномерным распределением по точкам.
    Если article передан — создаёт DDSOperation для каждой точки.
    Если у точки не хватает — дефицит покрывают другие точки.
    """
    from .models import GlobalCashRegister, GlobalCashOperation, GlobalCashDistribution, DDSOperation

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

    if points_qs is None:
        points_qs = Point.objects.filter(is_active=True)
    points_list = list(points_qs)
    if not points_list:
        raise ValidationError("Нет активных точек для распределения.")

    # кассы под блокировкой
    regs = {r.hotel_id: r for r in CashRegister.objects.select_for_update().filter(hotel__in=points_list)}
    for p in points_list:
        if p.id not in regs:
            r, _ = CashRegister.objects.get_or_create(hotel=p)
            regs[p.id] = r

    points_balances = [(p, getattr(regs[p.id], field) or Decimal("0")) for p in points_list]
    distributions = _distribute_amount(amount, points_balances)

    # списываем с общей кассы
    setattr(gcr, field, gcr_balance - amount)
    gcr.save(update_fields=[field, "updated_at"])

    op = GlobalCashOperation.objects.create(
        direction=GlobalCashOperation.OUT,
        account=account,
        amount=amount,
        happened_at=happened_at,
        comment=comment,
        article=article,
        created_by=created_by,
    )

    for point, dist_amount, note in distributions:
        if dist_amount <= 0:
            continue

        # создаём DDSOperation для точки (если указана статья)
        if article:
            dds_op = DDSOperation.objects.create(
                hotel=point,
                article=article,
                amount=dist_amount,
                happened_at=happened_at,
                method=account,
                comment=comment or f"Общая касса — распределение",
                source="global_cash",
                created_by=created_by,
            )
        else:
            dds_op = None

        # обновляем кассу точки напрямую (минуя apply_cash_movement чтобы не задвоить GCR)
        reg = regs[point.id]
        current = getattr(reg, field) or Decimal("0")
        setattr(reg, field, max(Decimal("0"), current - dist_amount))
        reg.save(update_fields=[field, "updated_at"])

        GlobalCashDistribution.objects.create(
            operation=op,
            point=point,
            amount=dist_amount,
            note=note,
        )

    article_name = article.name if article else None
    _notify_global(
        icon="🔴", label="Расход из общей кассы",
        account=account, amount=amount, comment=comment,
        created_by=created_by, happened_at=happened_at,
        article_name=article_name,
        distributions=[(p, a) for p, a, _ in distributions if a > 0],
    )
    return op
