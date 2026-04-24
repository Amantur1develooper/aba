from django.contrib import admin

# Register your models here.
from django.contrib import admin
from django.db.models import Sum

from .models import (
    HotelPMSSettings,
    RoomType, Room,
    Company,
    Booking, Stay, Guest, StayGuest,
    CompanyFolio, CompanyFolioItem,
    # Warehouse, Supplier, StockItem,
    # PurchaseReceipt, PurchaseLine,
    # Dish, WriteOff, WriteOffLine,
    # ExtraService, ExtraServiceSale,
    # Staff, StaffShift,
    # InventorySession, InventoryLine,
)


@admin.register(HotelPMSSettings)
class HotelPMSSettingsAdmin(admin.ModelAdmin):
    list_display = ("hotel", "is_enabled", "mode", "check_in_time", "check_out_time")
    list_filter = ("is_enabled", "mode")
    search_fields = ("hotel__name",)


@admin.register(RoomType)
class RoomTypeAdmin(admin.ModelAdmin):
    list_display = ("hotel", "name", "is_active")
    list_filter = ("hotel", "is_active")
    search_fields = ("name", "hotel__name")


@admin.register(Room)
class RoomAdmin(admin.ModelAdmin):
    list_display = ("hotel", "number", "floor", "room_type", "capacity", "is_active", "is_out_of_service", "clean_status")
    list_filter = ("hotel", "floor", "room_type", "is_active", "is_out_of_service", "clean_status")
    search_fields = ("number", "hotel__name", "room_type__name")
    list_editable = ("clean_status", "is_out_of_service", "is_active")


@admin.register(Company)
class CompanyAdmin(admin.ModelAdmin):
    list_display = ("name", "pay_terms", "contact_name", "contact_phone", "is_active")
    list_filter = ("pay_terms", "is_active")
    search_fields = ("name", "contact_name", "contact_phone")


@admin.register(Booking)
class BookingAdmin(admin.ModelAdmin):
    list_display = ("hotel", "booking_number", "booked_at", "guest_name", "check_in", "check_out", "status", "payment_status", "channel")
    list_filter = ("hotel", "status", "payment_status", "channel", "room_type")
    search_fields = ("booking_number", "guest_name", "hotel__name")
    raw_id_fields = ("room", "created_by")


class StayGuestInline(admin.TabularInline):
    model = StayGuest
    extra = 0
    autocomplete_fields = ("guest",)


@admin.register(Stay)
class StayAdmin(admin.ModelAdmin):
    list_display = ("hotel", "room", "check_in", "check_out", "status", "company", "guest_name", "guests_count", "tourist_tax_total")
    list_filter = ("hotel", "status", "company")
    search_fields = ("guest_name", "company__name", "room__number", "hotel__name")
    raw_id_fields = ("room", "booking", "created_by")
    inlines = [StayGuestInline]


@admin.register(Guest)
class GuestAdmin(admin.ModelAdmin):
    list_display = ("hotel", "full_name", "inn", "nationality", "is_foreigner")
    list_filter = ("hotel", "is_foreigner", "nationality")
    search_fields = ("full_name", "inn")


class CompanyFolioItemInline(admin.TabularInline):
    model = CompanyFolioItem
    extra = 0
    readonly_fields = ("created_at", "created_by")
    raw_id_fields = ("dds_operation", "cash_movement")
    fields = ("happened_at", "item_type", "description", "amount", "signed_amount", "dds_operation", "cash_movement", "created_by")


@admin.register(CompanyFolio)
class CompanyFolioAdmin(admin.ModelAdmin):
    list_display = ("hotel", "company", "is_closed", "balance_value", "created_at")
    list_filter = ("hotel", "is_closed")
    search_fields = ("company__name", "hotel__name")
    inlines = [CompanyFolioItemInline]

    def balance_value(self, obj):
        return obj.balance
    balance_value.short_description = "Баланс"


@admin.register(CompanyFolioItem)
class CompanyFolioItemAdmin(admin.ModelAdmin):
    list_display = ("folio", "happened_at", "item_type", "amount", "signed_amount", "created_by")
    list_filter = ("item_type", "folio__hotel")
    search_fields = ("folio__company__name", "description")
    raw_id_fields = ("dds_operation", "cash_movement", "created_by")
