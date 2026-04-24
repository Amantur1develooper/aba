# dds/services.py
from decimal import Decimal
from django.db import transaction
from django.db.models import F
from django.utils import timezone

from .models import CashTransfer, CashMovement, CashRegister

def _balance_field(account: str) -> str:
    return {
        CashMovement.ACC_CASH: "cash_balance",
        CashMovement.ACC_MKASSA: "mkassa_balance",
        CashMovement.ACC_ZADATOK: "zadatok_balance",
        CashMovement.ACC_OPTIMA: "optima_balance",
    }[account]

@transaction.atomic
def create_cash_transfer(*, hotel, user, from_account: str, to_account: str, amount: Decimal, happened_at=None, comment="") -> CashTransfer:
    if not happened_at:
        happened_at = timezone.now()

    # блокируем кассу отеля
    register = CashRegister.objects.select_for_update().get(hotel=hotel)

    from_field = _balance_field(from_account)
    to_field = _balance_field(to_account)

    current_from = getattr(register, from_field) or Decimal("0.00")
    if current_from < amount:
        raise ValueError(f"Недостаточно средств на счёте {from_account}. Баланс={current_from}")

    transfer = CashTransfer.objects.create(
        hotel=hotel,
        register=register,
        from_account=from_account,
        to_account=to_account,
        amount=amount,
        happened_at=happened_at,
        comment=comment,
        created_by=user,
    )

    # движения денег
    CashMovement.objects.create(
        register=register,
        hotel=hotel,
        direction=CashMovement.OUT,
        account=from_account,
        amount=amount,
        happened_at=happened_at,
        comment=f"Перевод {from_account} → {to_account}. {comment}".strip(),
        created_by=user,
        transfer=transfer,
    )
    CashMovement.objects.create(
        register=register,
        hotel=hotel,
        direction=CashMovement.IN,
        account=to_account,
        amount=amount,
        happened_at=happened_at,
        comment=f"Перевод {from_account} → {to_account}. {comment}".strip(),
        created_by=user,
        transfer=transfer,
    )

    # обновляем балансы атомарно
    CashRegister.objects.filter(pk=register.pk).update(
        **{
            from_field: F(from_field) - amount,
            to_field: F(to_field) + amount,
        }
    )

    return transfer
