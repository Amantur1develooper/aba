"""
Fetches recent /start messages from the bot via getUpdates
and auto-links chat_ids to profiles by token.

Run once after users send /start <token> to the bot:
    python manage.py sync_tg
"""
import json
import urllib.request
from django.core.management.base import BaseCommand
from django.conf import settings


class Command(BaseCommand):
    help = "Sync Telegram chat IDs from recent bot messages"

    def handle(self, *args, **options):
        token = settings.TELEGRAM_BOT_TOKEN
        if not token:
            self.stderr.write("TELEGRAM_BOT_TOKEN not set")
            return

        url = f"https://api.telegram.org/bot{token}/getUpdates?limit=100"
        with urllib.request.urlopen(url, timeout=10) as r:
            data = json.loads(r.read())

        if not data.get("ok"):
            self.stderr.write(f"Error: {data}")
            return

        from accounts.models import Profile

        linked = 0
        for update in data.get("result", []):
            msg = update.get("message") or update.get("edited_message")
            if not msg:
                continue
            text = (msg.get("text") or "").strip()
            if not text.startswith("/start"):
                continue
            parts = text.split(maxsplit=1)
            link_token = parts[1].strip() if len(parts) > 1 else ""
            chat_id = str(msg.get("chat", {}).get("id", ""))
            if not chat_id or not link_token:
                continue
            try:
                profile = Profile.objects.get(tg_link_token=link_token)
                profile.tg_chat_id = chat_id
                profile.tg_link_token = ""
                profile.save(update_fields=["tg_chat_id", "tg_link_token"])
                self.stdout.write(self.style.SUCCESS(
                    f"Linked {profile.user.username} -> chat_id {chat_id}"
                ))
                linked += 1
            except Profile.DoesNotExist:
                pass

        if linked == 0:
            self.stdout.write("No new links found.")
        else:
            self.stdout.write(self.style.SUCCESS(f"Done: {linked} user(s) linked."))
