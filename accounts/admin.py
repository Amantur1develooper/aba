from django.contrib import admin
from django.contrib import messages
from .models import Profile


@admin.register(Profile)
class ProfileAdmin(admin.ModelAdmin):
    list_display = ("user", "hotel", "is_finance_admin", "tg_chat_id", "tg_link_token")
    list_filter = ("hotel", "is_finance_admin")
    search_fields = ("user__username", "user__first_name", "user__last_name")
    fields = ("user", "hotel", "is_finance_admin", "tg_chat_id", "tg_link_token")
    actions = ["send_test_message", "clear_tg"]

    @admin.action(description="Отправить тест в Telegram")
    def send_test_message(self, request, queryset):
        from dds.telegram import send_tg
        sent = 0
        for profile in queryset:
            if profile.tg_chat_id:
                send_tg(profile.tg_chat_id, "✅ Тест: бот работает! Уведомления о транзакциях активны.")
                sent += 1
        self.message_user(request, f"Отправлено {sent} сообщений.", messages.SUCCESS)

    @admin.action(description="Очистить Telegram привязку")
    def clear_tg(self, request, queryset):
        queryset.update(tg_chat_id="", tg_link_token="")
        self.message_user(request, "Привязка удалена.", messages.WARNING)
