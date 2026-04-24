from django.db.models.signals import post_save
from django.db.models import Q
from django.dispatch import receiver
from .models import Point, CashRegister, DDSOperation, CashIncasso, CashTransfer


@receiver(post_save, sender=Point)
def ensure_cash_register(sender, instance, created, **kwargs):
    if created:
        CashRegister.objects.get_or_create(hotel=instance)


def _chat_ids_for_hotel(hotel):
    from accounts.models import Profile
    profiles = Profile.objects.filter(tg_chat_id__gt="").filter(
        Q(is_finance_admin=True) | Q(hotel=hotel)
    )
    return [p.tg_chat_id for p in profiles]


def _fmt_amount(amount):
    try:
        return f"{float(amount):,.0f}".replace(",", " ")
    except Exception:
        return str(amount)


@receiver(post_save, sender=DDSOperation)
def notify_dds_operation(sender, instance, created, **kwargs):
    if not created or instance.is_voided:
        return
    from .telegram import notify_transaction
    kind_icon = "🟢" if instance.article.kind == "income" else "🔴"
    kind_label = "Доход" if instance.article.kind == "income" else "Расход"
    lines = [
        f"{kind_icon} <b>{kind_label}</b> — {instance.hotel.name}",
        f"💰 {_fmt_amount(instance.amount)} сом",
        f"📋 {instance.article.name}",
        f"💳 {instance.get_method_display()}",
    ]
    if instance.counterparty:
        lines.append(f"🤝 {instance.counterparty}")
    if instance.comment:
        lines.append(f"💬 {instance.comment}")
    lines.append(f"👤 {instance.created_by.get_full_name() or instance.created_by.username}")
    lines.append(f"🕐 {instance.happened_at.strftime('%d.%m.%Y %H:%M')}")
    text = "\n".join(lines)
    notify_transaction(_chat_ids_for_hotel(instance.hotel), text)


@receiver(post_save, sender=CashIncasso)
def notify_incasso(sender, instance, created, **kwargs):
    if not created:
        return
    from .telegram import notify_transaction
    lines = [
        f"🏦 <b>Инкассация</b> — {instance.hotel.name}",
        f"💰 {_fmt_amount(instance.amount)} сом",
        f"💳 {instance.get_method_display()}",
    ]
    if instance.comment:
        lines.append(f"💬 {instance.comment}")
    lines.append(f"👤 {instance.created_by.get_full_name() or instance.created_by.username}")
    lines.append(f"🕐 {instance.happened_at.strftime('%d.%m.%Y %H:%M')}")
    notify_transaction(_chat_ids_for_hotel(instance.hotel), "\n".join(lines))


@receiver(post_save, sender=CashTransfer)
def notify_transfer(sender, instance, created, **kwargs):
    if not created or instance.is_voided:
        return
    from .telegram import notify_transaction
    from .models import CashMovement
    labels = dict(CashMovement.ACCOUNT_CHOICES)
    lines = [
        f"🔄 <b>Перевод между счетами</b> — {instance.hotel.name}",
        f"💰 {_fmt_amount(instance.amount)} сом",
        f"📤 {labels.get(instance.from_account, instance.from_account)} → 📥 {labels.get(instance.to_account, instance.to_account)}",
    ]
    if instance.comment:
        lines.append(f"💬 {instance.comment}")
    lines.append(f"👤 {instance.created_by.get_full_name() or instance.created_by.username}")
    lines.append(f"🕐 {instance.happened_at.strftime('%d.%m.%Y %H:%M')}")
    notify_transaction(_chat_ids_for_hotel(instance.hotel), "\n".join(lines))
