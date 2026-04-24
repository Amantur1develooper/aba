import json
import logging
import threading
import urllib.parse
import urllib.request
from django.conf import settings

logger = logging.getLogger(__name__)


def _send(chat_id: str, text: str) -> None:
    token = getattr(settings, "TELEGRAM_BOT_TOKEN", "")
    if not token or not chat_id:
        return
    url = f"https://api.telegram.org/bot{token}/sendMessage"
    data = json.dumps({"chat_id": chat_id, "text": text, "parse_mode": "HTML"}).encode()
    req = urllib.request.Request(url, data=data, headers={"Content-Type": "application/json"})
    try:
        urllib.request.urlopen(req, timeout=5)
    except Exception as e:
        logger.warning("Telegram send failed for chat_id=%s: %s", chat_id, e)


def send_tg(chat_id: str, text: str) -> None:
    """Send a Telegram message in a background thread so it never blocks a request."""
    threading.Thread(target=_send, args=(chat_id, text), daemon=True).start()


def notify_transaction(chat_ids: list[str], text: str) -> None:
    for cid in set(chat_ids):
        if cid:
            send_tg(cid, text)
