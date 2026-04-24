import json
import logging
from django.http import HttpResponse
from django.views.decorators.csrf import csrf_exempt
from django.views.decorators.http import require_POST

logger = logging.getLogger(__name__)


@csrf_exempt
@require_POST
def tg_webhook(request):
    try:
        data = json.loads(request.body)
    except Exception:
        return HttpResponse("bad json", status=400)

    message = data.get("message") or data.get("edited_message")
    if not message:
        return HttpResponse("ok")

    chat_id = str(message.get("chat", {}).get("id", ""))
    text = (message.get("text") or "").strip()

    if not chat_id or not text.startswith("/start"):
        return HttpResponse("ok")

    parts = text.split(maxsplit=1)
    token = parts[1].strip() if len(parts) > 1 else ""

    from accounts.models import Profile
    from .telegram import send_tg

    if token:
        try:
            profile = Profile.objects.get(tg_link_token=token)
            profile.tg_chat_id = chat_id
            profile.tg_link_token = ""
            profile.save(update_fields=["tg_chat_id", "tg_link_token"])
            name = profile.user.get_full_name() or profile.user.username
            send_tg(chat_id, f"✅ Готово, {name}! Теперь вы будете получать уведомления о транзакциях.")
        except Profile.DoesNotExist:
            send_tg(chat_id, "❌ Код не найден или уже использован. Зайдите в профиль и получите новый код.")
    else:
        send_tg(chat_id, f"👋 Привет! Чтобы подключить уведомления, зайдите в свой профиль и нажмите «Подключить через Telegram».")

    return HttpResponse("ok")
