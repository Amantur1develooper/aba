from django.db import models
from django.conf import settings
from dds.models import Point


class Profile(models.Model):
    user = models.OneToOneField(settings.AUTH_USER_MODEL, on_delete=models.CASCADE, related_name="profile")
    hotel = models.ForeignKey(Point, on_delete=models.PROTECT, null=True, blank=True)
    is_finance_admin = models.BooleanField(default=False)
    tg_chat_id = models.CharField(
        max_length=50, blank=True,
        verbose_name="Telegram Chat ID",
        help_text="Получите у бота командой /start",
    )

    def __str__(self):
        return f"Profile: {self.user}"
