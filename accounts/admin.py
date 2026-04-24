from django.contrib import admin

# Register your models here.
from django.contrib import admin
from .models import Profile

@admin.register(Profile)
class ProfileAdmin(admin.ModelAdmin):
    list_display = ("user", "hotel", "is_finance_admin", "tg_chat_id")
    list_filter = ("hotel", "is_finance_admin")
    search_fields = ("user__username", "user__first_name", "user__last_name")
