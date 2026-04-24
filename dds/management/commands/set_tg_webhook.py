import json
import urllib.request
from django.core.management.base import BaseCommand
from django.conf import settings


class Command(BaseCommand):
    help = "Register Telegram webhook URL"

    def add_arguments(self, parser):
        parser.add_argument("site_url", help="e.g. https://yourdomain.com")

    def handle(self, *args, **options):
        token = settings.TELEGRAM_BOT_TOKEN
        if not token:
            self.stderr.write("TELEGRAM_BOT_TOKEN not set")
            return

        site = options["site_url"].rstrip("/")
        webhook_url = f"{site}/tg/webhook/"

        api_url = f"https://api.telegram.org/bot{token}/setWebhook"
        data = json.dumps({"url": webhook_url}).encode()
        req = urllib.request.Request(api_url, data=data, headers={"Content-Type": "application/json"})
        with urllib.request.urlopen(req, timeout=10) as resp:
            result = json.loads(resp.read())

        if result.get("ok"):
            self.stdout.write(self.style.SUCCESS(f"Webhook set: {webhook_url}"))
        else:
            self.stderr.write(f"Error: {result}")
