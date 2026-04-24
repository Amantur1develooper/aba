# pms/services.py
# from __future__ import annotations

# from dataclasses import dataclass
from decimal import Decimal
from datetime import datetime, time, timedelta
from typing import Optional, Iterable

from django.db import transaction
from django.db.models import Q, F
from django.utils import timezone

from .models import Stay, Room, CompanyFolio, CompanyFolioItem
from dds.models import CashRegister, CashMovement, DDSOperation, DDSArticle, DDSCategory


class PMSConflictError(Exception):
    """Пересечение броней/проживаний для одного номера."""
    pass


def _money(v) -> Decimal:
    if v is None:
        return Decimal("0.00")
    if isinstance(v, Decimal):
        return v
    return Decimal(str(v))


def _period_overlap_q(start_dt, end_dt):
    # пересечение: A.start < B.end AND A.end > B.start
    return Q(check_in__lt=end_dt) & Q(check_out__gt=start_dt)


def assert_no_overlap(*, room: Room, start_dt, end_dt, exclude_stay_id: Optional[int] = None):
    qs = Stay.objects.filter(room=room).exclude(status__in=[Stay.CANCELED, Stay.NO_SHOW])
    qs = qs.filter(_period_overlap_q(start_dt, end_dt))
    if exclude_stay_id:
        qs = qs.exclude(id=exclude_stay_id)
    if qs.exists():
        raise PMSConflictError(f"Номер {room.number} уже занят в выбранный период.")


def ensure_cash_register(hotel) -> CashRegister:
    register, _ = CashRegister.objects.get_or_create(hotel=hotel)
    return register


def _register_field_for_method(method: str) -> str:
    # CashRegister поля: cash_balance/mkassa_balance/zadatok_balance/optima_balance
    mapping = {
        DDSOperation.CASH: "cash_balance",
        DDSOperation.MKASSA: "mkassa_balance",
        DDSOperation.ZADATOK: "zadatok_balance",
        DDSOperation.OPTIMA: "optima_balance",
    }
    return mapping.get(method, "cash_balance")


def _cash_account_for_method(method: str) -> str:
    mapping = {
        DDSOperation.CASH: CashMovement.ACC_CASH,
        DDSOperation.MKASSA: CashMovement.ACC_MKASSA,
        DDSOperation.ZADATOK: CashMovement.ACC_ZADATOK,
        DDSOperation.OPTIMA: CashMovement.ACC_OPTIMA,
    }
    return mapping.get(method, CashMovement.ACC_CASH)


def ensure_default_stay_income_article() -> DDSArticle:
    """
    Создаем (если нет) категорию/статью для оплаты проживания.
    Можно потом заменить на выбор статьи из формы.
    """
    cat, _ = DDSCategory.objects.get_or_create(
        kind=DDSCategory.INCOME,
        parent=None,
        name="Проживание",
        defaults={"is_active": True},
    )
    art, _ = DDSArticle.objects.get_or_create(
        kind=DDSArticle.INCOME,
        category=cat,
        name="Оплата проживания",
        defaults={"is_active": True},
    )
    return art


@transaction.atomic
def apply_cash_in(
    *,
    hotel,
    user,
    method: str,
    amount: Decimal,
    happened_at,
    comment: str = "",
    dds_operation: Optional[DDSOperation] = None,
) -> CashMovement:
    """
    Приход денег (увеличиваем баланс нужного счета + пишем CashMovement).
    """
    amount = _money(amount)
    if amount <= 0:
        raise ValueError("Сумма должна быть больше 0.")

    reg = ensure_cash_register(hotel)
    field = _register_field_for_method(method)
    account = _cash_account_for_method(method)

    # Обновляем баланс атомарно
    CashRegister.objects.filter(id=reg.id).update(**{field: F(field) + amount})

    move = CashMovement.objects.create(
        register=reg,
        hotel=hotel,
        direction=CashMovement.IN,
        account=account,
        amount=amount,
        happened_at=happened_at,
        comment=comment,
        dds_operation=dds_operation,
        created_by=user,
    )
    return move


@transaction.atomic
def make_dds_income(
    *,
    hotel,
    user,
    article: DDSArticle,
    method: str,
    amount: Decimal,
    happened_at,
    counterparty: str = "",
    source: str = "",
    comment: str = "",
) -> DDSOperation:
    amount = _money(amount)
    if amount <= 0:
        raise ValueError("Сумма должна быть больше 0.")

    op = DDSOperation.objects.create(
        hotel=hotel,
        article=article,
        amount=amount,
        happened_at=happened_at,
        method=method,
        counterparty=counterparty or "",
        source=source or "",
        comment=comment or "",
        created_by=user,
    )
    return op


@transaction.atomic
def folio_charge_for_stay(*, stay: Stay, user, description: str = "") -> CompanyFolioItem:
    """
    Начисление в фолио (для корпоративных оплат по счету/раз в неделю).
    В ДДС/кассу пока НЕ пишем (ДДС = движение денег).
    """
    if not stay.company:
        raise ValueError("Для фолио нужна компания.")

    folio, _ = CompanyFolio.objects.get_or_create(hotel=stay.hotel, company=stay.company)

    amt = _money(stay.total_to_pay)
    item = CompanyFolioItem.objects.create(
        folio=folio,
        kind=CompanyFolioItem.CHARGE,
        happened_at=timezone.now(),
        description=description or f"Начисление проживания (Stay #{stay.id})",
        amount=amt,
        stay=stay,
        created_by=user,
    )
    return item


@transaction.atomic
def check_in_stay(
    *,
    stay: Stay,
    user,
    pay_now: bool,
    method: str,
    paid_amount: Optional[Decimal] = None,
    dds_article: Optional[DDSArticle] = None,
):
    """
    Заселение:
    - всегда ставим статус IN
    - если корпоративный и pay_now=False (по счету/раз в неделю) -> начисляем в фолио, без кассы/ДДС
    - если pay_now=True -> пишем CashMovement(IN) + DDSOperation(INCOME)
    """
    if stay.status in [Stay.CANCELED, Stay.NO_SHOW]:
        raise ValueError("Нельзя заселить отменённую/No Show запись.")

    stay.status = Stay.IN
    stay.save(update_fields=["status"])

    # Корпоративное "в долг" -> фолио (без денег)
    if stay.stay_type == Stay.CORPORATE and stay.company and not pay_now:
        folio_charge_for_stay(stay=stay, user=user)
        return

    # Оплата сейчас -> касса + ДДС
    amount_to_take = _money(paid_amount) if paid_amount is not None else _money(stay.total_to_pay)
    if amount_to_take <= 0:
        return

    if dds_article is None:
        dds_article = ensure_default_stay_income_article()

    counterparty = stay.company.name if (stay.company and stay.stay_type == Stay.CORPORATE) else (stay.guest_name or "")
    source = f"pms:stay:{stay.id}"

    dds_op = make_dds_income(
        hotel=stay.hotel,
        user=user,
        article=dds_article,
        method=method,
        amount=amount_to_take,
        happened_at=stay.check_in,
        counterparty=counterparty,
        source=source,
        comment="Оплата проживания",
    )

    cash_move = apply_cash_in(
        hotel=stay.hotel,
        user=user,
        method=method,
        amount=amount_to_take,
        happened_at=stay.check_in,
        comment=f"Оплата проживания (Stay #{stay.id})",
        dds_operation=dds_op,
    )

    stay.dds_operation = dds_op
    stay.cash_movement = cash_move
    stay.save(update_fields=["dds_operation", "cash_movement"])


@transaction.atomic
def check_out_stay(*, stay: Stay, user):
    """
    Выезд:
    - status = OUT
    - room.clean_status = DIRTY
    """
    stay.status = Stay.OUT
    stay.save(update_fields=["status"])

    room = stay.room
    room.clean_status = Room.DIRTY
    room.save(update_fields=["clean_status"])


@transaction.atomic
def cancel_stay(*, stay: Stay, user, reason: str = "Отмена"):
    """
    MVP: ставим статус отмены.
    Если были ДДС-операции — сторнируем DDSOperation (у тебя есть void()).
    С кассой лучше делать возврат отдельной операцией (можно добавить позже).
    """
    stay.status = Stay.CANCELED
    stay.save(update_fields=["status"])

    if stay.dds_operation and not stay.dds_operation.is_voided:
        stay.dds_operation.void(user=user, reason=reason)


from decimal import Decimal
from django.db import transaction
from django.db.models import F
from django.utils import timezone

from dds.models import DDSArticle, DDSOperation, CashRegister, CashMovement
from .models import CompanyFolio, CompanyFolioItem


ACCOUNT_FIELD = {
    CashMovement.ACC_CASH: "cash_balance",
    CashMovement.ACC_MKASSA: "mkassa_balance",
    CashMovement.ACC_ZADATOK: "zadatok_balance",
    CashMovement.ACC_OPTIMA: "optima_balance",
}


def _get_default_income_article():
    # максимально безопасный дефолт — первый INCOME
    art = DDSArticle.objects.filter(kind=DDSArticle.INCOME, is_active=True).order_by("id").first()
    return art


@transaction.atomic
def folio_add_payment(*, folio: CompanyFolio, user, amount: Decimal, method: str, article=None, comment: str = ""):
    """
    Создаёт:
    - DDSOperation (income)
    - CashMovement (IN)
    - обновляет CashRegister
    - CompanyFolioItem (payment)
    """
    amount = (amount or Decimal("0.00"))
    if amount <= 0:
        raise ValueError("Сумма должна быть больше 0")

    hotel = folio.hotel
    company = folio.company

    if article is None:
        article = _get_default_income_article()
    if article is None:
        raise ValueError("Нет статьи дохода (DDSArticle INCOME). Создай хотя бы одну.")

    reg, _ = CashRegister.objects.get_or_create(hotel=hotel)

    # 1) DDS
    dds_op = DDSOperation.objects.create(
        hotel=hotel,
        article=article,
        amount=amount,
        happened_at=timezone.now(),
        method=method,
        counterparty=company.name,
        comment=(comment or f"Оплата по фолио компании: {company.name}"),
        source="company_folio",
        created_by=user,
    )

    # 2) CashMovement
    cash_mv = CashMovement.objects.create(
        register=reg,
        hotel=hotel,
        direction=CashMovement.IN,
        account=method,  # совпадает с choices cash/mkassa/zadatok/optima
        amount=amount,
        happened_at=timezone.now(),
        comment=f"Оплата по фолио: {company.name}",
        dds_operation=dds_op,
        created_by=user,
    )

    # 3) обновляем баланс кассы (быстро и безопасно)
    field = ACCOUNT_FIELD.get(method)
    if field:
        CashRegister.objects.filter(pk=reg.pk).update(**{field: F(field) + amount})

    # 4) Folio item
    item = CompanyFolioItem.objects.create(
        folio=folio,
        item_type=CompanyFolioItem.PAYMENT,
        happened_at=timezone.now(),
        description=f"Оплата ({method})",
        amount=amount,
        signed_amount=CompanyFolioItem.make_signed(CompanyFolioItem.PAYMENT, amount),
        dds_operation=dds_op,
        cash_movement=cash_mv,
        created_by=user,
    )

    # 5) закрыть долг если погашено
    folio.refresh_closed_flag()

    return item
