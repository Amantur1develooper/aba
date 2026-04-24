# dds/admin.py
from django.contrib import admin
from django.db.models import Sum
from django.utils.html import format_html
from django.utils import timezone

from .models import (
    Point,
    PointContact,
    DDSCategory,
    DDSArticle,
    DDSOperation,
    CashRegister,
    CashMovement,
    CashIncasso,
    CashTransfer,
)

# ----------------------------
# Helpers
# ----------------------------

def money(v):
    try:
        return f"{v:,.2f}".replace(",", " ")
    except Exception:
        return v


# ----------------------------
# Inlines
# ----------------------------

class PointContactInline(admin.TabularInline):
    model = PointContact
    extra = 1
    fields = ("name", "position", "phone", "whatsapp")


class CashRegisterInline(admin.StackedInline):
    model = CashRegister
    extra = 0
    max_num = 1
    can_delete = False
    readonly_fields = ("noncash_total", "total", "updated_at")
    fieldsets = (
        ("Баланс", {
            "fields": (
                ("cash_balance",),
                ("mkassa_balance", "zadatok_balance", "optima_balance"),
                ("noncash_total", "total"),
                "updated_at",
            )
        }),
    )


class DDSOperationInline(admin.TabularInline):
    model = DDSOperation
    extra = 0
    fields = ("happened_at", "article", "method", "amount", "counterparty", "source", "is_voided")
    readonly_fields = ("happened_at", "article", "method", "amount", "counterparty", "source", "is_voided")
    show_change_link = True
    ordering = ("-happened_at",)
    can_delete = False


class CashMovementInline(admin.TabularInline):
    model = CashMovement
    extra = 0
    fields = ("happened_at", "direction", "account", "amount", "dds_operation", "incasso", "transfer")
    readonly_fields = ("happened_at", "direction", "account", "amount", "dds_operation", "incasso", "transfer")
    show_change_link = True
    ordering = ("-happened_at",)
    can_delete = False


class CashTransferInline(admin.TabularInline):
    model = CashTransfer
    extra = 0
    fields = ("happened_at", "from_account", "to_account", "amount", "is_voided", "created_by")
    readonly_fields = ("happened_at", "from_account", "to_account", "amount", "is_voided", "created_by")
    show_change_link = True
    ordering = ("-happened_at",)
    can_delete = False


class CashIncassoInline(admin.TabularInline):
    model = CashIncasso
    extra = 0
    fields = ("happened_at", "method", "amount", "created_by")
    readonly_fields = ("happened_at", "method", "amount", "created_by")
    show_change_link = True
    ordering = ("-happened_at",)
    can_delete = False


# ----------------------------
# Admin: Point
# ----------------------------

@admin.register(Point)
class PointAdmin(admin.ModelAdmin):
    list_display = (
        "name",
        "is_active",
        "cash_balance",
        "mkassa_balance",
        "zadatok_balance",
        "optima_balance",
        "noncash_total",
        "total_balance",
        "open_register_link",
    )
    list_filter = ("is_active",)
    search_fields = ("name",)
    ordering = ("name",)

    inlines = (
        PointContactInline,
        CashRegisterInline,
        DDSOperationInline,
        CashMovementInline,
        CashTransferInline,
        CashIncassoInline,
    )

    def get_queryset(self, request):
        qs = super().get_queryset(request)
        return qs.select_related("cash_register")

    def _reg(self, obj):
        return getattr(obj, "cash_register", None)

    @admin.display(description="Нал", ordering="cash_register__cash_balance")
    def cash_balance(self, obj):
        reg = self._reg(obj)
        return money(reg.cash_balance) if reg else "—"

    @admin.display(description="Mkassa", ordering="cash_register__mkassa_balance")
    def mkassa_balance(self, obj):
        reg = self._reg(obj)
        return money(reg.mkassa_balance) if reg else "—"

    @admin.display(description="Задаток", ordering="cash_register__zadatok_balance")
    def zadatok_balance(self, obj):
        reg = self._reg(obj)
        return money(reg.zadatok_balance) if reg else "—"

    @admin.display(description="Оптима", ordering="cash_register__optima_balance")
    def optima_balance(self, obj):
        reg = self._reg(obj)
        return money(reg.optima_balance) if reg else "—"

    @admin.display(description="Безнал (итого)")
    def noncash_total(self, obj):
        reg = self._reg(obj)
        return money(reg.noncash_total) if reg else "—"

    @admin.display(description="Итого")
    def total_balance(self, obj):
        reg = self._reg(obj)
        return money(reg.total) if reg else "—"

    @admin.display(description="Касса")
    def open_register_link(self, obj):
        reg = getattr(obj, "cash_register", None)
        if not reg:
            return format_html("<span style='color:#999'>нет кассы</span>")
        return format_html(
            "<a class='button' href='/admin/dds/cashregister/{}/change/'>Открыть</a>",
            reg.id
        )


# ----------------------------
# Admin: CashRegister
# ----------------------------

@admin.register(CashRegister)
class CashRegisterAdmin(admin.ModelAdmin):
    list_display = (
        "hotel",
        "cash_balance",
        "mkassa_balance",
        "zadatok_balance",
        "optima_balance",
        "noncash_total",
        "total",
        "updated_at",
    )
    list_select_related = ("hotel",)
    search_fields = ("hotel__name",)
    ordering = ("hotel__name",)
    readonly_fields = ("noncash_total", "total", "updated_at")
    fieldsets = (
        ("Точка", {"fields": ("hotel",)}),
        ("Баланс", {
            "fields": (
                ("cash_balance",),
                ("mkassa_balance", "zadatok_balance", "optima_balance"),
                ("noncash_total", "total"),
                "updated_at",
            )
        }),
    )


# ----------------------------
# Admin: DDSCategory
# ----------------------------

@admin.register(DDSCategory)
class DDSCategoryAdmin(admin.ModelAdmin):
    list_display = ("name", "kind", "parent", "is_active")
    list_filter = ("kind", "is_active")
    search_fields = ("name", "parent__name")
    ordering = ("kind", "parent_id", "name")
    autocomplete_fields = ("parent",)


# ----------------------------
# Admin: DDSArticle
# ----------------------------

from django.contrib import admin
from .models import Point, PointContact, DDSCategory, DDSArticle, DDSOperation, CashRegister, CashMovement, CashIncasso, CashTransfer

@admin.register(DDSArticle)
class DDSArticleAdmin(admin.ModelAdmin):
    list_display = ("id", "kind", "name", "category", "is_active", "hotels_list")
    list_filter = ("kind", "is_active", "category")
    search_fields = ("name",)
    filter_horizontal = ("hotels",)  # ✅ удобно выбирать точки

    def hotels_list(self, obj):
        qs = obj.hotels.all()[:5]
        if not obj.hotels.exists():
            return "Все точки"
        names = [h.name for h in qs]
        more = obj.hotels.count() - len(names)
        return ", ".join(names) + (f" (+{more})" if more > 0 else "")
    hotels_list.short_description = "Точки"

# ----------------------------
# Admin: DDSOperation
# ----------------------------

@admin.register(DDSOperation)
class DDSOperationAdmin(admin.ModelAdmin):
    date_hierarchy = "happened_at"
    list_display = (
        "happened_at",
        "hotel",
        "kind_badge",
        "category_name",
        "article",
        "method",
        "amount",
        "counterparty",
        "source",
        "is_voided",
        "created_by",
    )
    list_filter = (
        "hotel",
        "method",
        "is_voided",
        "article__kind",
        "article__category",
        "article",
    )
    search_fields = (
        "hotel__name",
        "article__name",
        "counterparty",
        "source",
        "comment",
    )
    ordering = ("-happened_at", "-id")
    autocomplete_fields = ("hotel", "article", "created_by", "voided_by")
    readonly_fields = ("created_at",)

    fieldsets = (
        ("Основное", {
            "fields": ("hotel", "article", "amount", "method", "happened_at")
        }),
        ("Дополнительно", {
            "fields": ("counterparty", "source", "comment", "created_by", "created_at")
        }),
        ("Сторно", {
            "fields": ("is_voided", "void_reason", "voided_at", "voided_by")
        }),
    )

    def get_queryset(self, request):
        qs = super().get_queryset(request)
        return qs.select_related(
            "hotel", "article", "article__category", "article__category__parent", "created_by", "voided_by"
        )
    @admin.display(description="Вид")
    def kind_badge(self, obj):
        kind = obj.article.kind
        if kind == DDSArticle.INCOME:
            return format_html(
            "<span style='padding:2px 8px;border-radius:10px;background:#d1fae5;color:#065f46;'>{}</span>",
            "Доход"
        )
        return format_html(
        "<span style='padding:2px 8px;border-radius:10px;background:#fee2e2;color:#991b1b;'>{}</span>",
        "Расход"
    )

    # @admin.display(description="Вид")
    # def kind_badge(self, obj):
    #     kind = obj.article.kind
    #     if kind == DDSArticle.INCOME:
    #         return format_html("<span style='padding:2px 8px;border-radius:10px;background:#d1fae5;color:#065f46;'>Доход</span>")
    #     return format_html("<span style='padding:2px 8px;border-radius:10px;background:#fee2e2;color:#991b1b;'>Расход</span>")

    @admin.display(description="Категория")
    def category_name(self, obj):
        cat = obj.article.category
        if not cat:
            return "—"
        if cat.parent:
            return f"{cat.parent.name} → {cat.name}"
        return cat.name

    def save_model(self, request, obj, form, change):
        # если админ создаёт вручную — подставим created_by
        if not change and not obj.created_by_id:
            obj.created_by = request.user
        super().save_model(request, obj, form, change)


# ----------------------------
# Admin: CashMovement
# ----------------------------

@admin.register(CashMovement)
class CashMovementAdmin(admin.ModelAdmin):
    date_hierarchy = "happened_at"
    list_display = (
        "happened_at",
        "hotel",
        "direction",
        "account",
        "amount",
        "dds_operation",
        "incasso",
        "transfer",
        "created_by",
    )
    list_filter = ("hotel", "direction", "account")
    search_fields = ("hotel__name", "comment", "dds_operation__source")
    ordering = ("-happened_at", "-id")
    autocomplete_fields = ("hotel", "register", "dds_operation", "incasso", "transfer", "created_by")
    readonly_fields = ("created_at",)

    def get_queryset(self, request):
        return super().get_queryset(request).select_related(
            "hotel", "register", "dds_operation", "incasso", "transfer", "created_by"
        )


# ----------------------------
# Admin: CashIncasso
# ----------------------------

@admin.register(CashIncasso)
class CashIncassoAdmin(admin.ModelAdmin):
    date_hierarchy = "happened_at"
    list_display = ("happened_at", "hotel", "method", "amount", "created_by")
    list_filter = ("hotel", "method")
    search_fields = ("hotel__name", "comment")
    ordering = ("-happened_at", "-id")
    autocomplete_fields = ("hotel", "created_by")
    readonly_fields = ("created_at",)


# ----------------------------
# Admin: CashTransfer
# ----------------------------

@admin.register(CashTransfer)
class CashTransferAdmin(admin.ModelAdmin):
    date_hierarchy = "happened_at"
    list_display = (
        "happened_at",
        "hotel",
        "from_account",
        "to_account",
        "amount",
        "is_voided",
        "created_by",
    )
    list_filter = ("hotel", "from_account", "to_account", "is_voided")
    search_fields = ("hotel__name", "comment")
    ordering = ("-happened_at", "-id")
    autocomplete_fields = ("hotel", "register", "created_by", "voided_by")
    readonly_fields = ("created_at", "voided_at")
