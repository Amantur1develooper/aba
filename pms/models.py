from decimal import Decimal
from collections import defaultdict
from django.conf import settings
from django.db import models
from django.db.models import Q, Sum
from django.utils import timezone


class HotelPMSSettings(models.Model):
    DAILY = "daily"
    HOURLY = "hourly"
    MODE_CHOICES = ((DAILY, "Посуточно"), (HOURLY, "Почасово"))

    hotel = models.OneToOneField("dds.Point", on_delete=models.CASCADE, related_name="pms_settings")
    is_enabled = models.BooleanField(default=False, verbose_name="Шахматка включена")

    check_in_time = models.TimeField(
        default=timezone.datetime(2000, 1, 1, 14, 0).time(),
        verbose_name="Время заезда"
    )
    check_out_time = models.TimeField(
        default=timezone.datetime(2000, 1, 1, 12, 0).time(),
        verbose_name="Время выезда"
    )
    mode = models.CharField(max_length=10, choices=MODE_CHOICES, default=DAILY, verbose_name="Режим")

    class Meta:
        verbose_name = "Настройки PMS"
        verbose_name_plural = "Настройки PMS"

    def __str__(self):
        return f"PMS: {self.hotel}"


class RoomType(models.Model):
    hotel = models.ForeignKey("dds.Point", on_delete=models.PROTECT, related_name="room_types", verbose_name="Точка")
    name = models.CharField(max_length=120, verbose_name="Тип номера")
    default_capacity = models.PositiveSmallIntegerField(default=2, verbose_name="Вместимость по умолчанию")
    default_day_rate = models.DecimalField(max_digits=12, decimal_places=2, default=0, verbose_name="Тариф/сутки по умолчанию")
    default_hour_rate = models.DecimalField(max_digits=12, decimal_places=2, default=0, verbose_name="Тариф/час по умолчанию")
    is_active = models.BooleanField(default=True, verbose_name="Активен")

    class Meta:
        verbose_name = "Тип номера"
        verbose_name_plural = "Типы номеров"
        ordering = ["hotel_id", "name"]
        constraints = [
            models.UniqueConstraint(fields=["hotel", "name"], name="uniq_roomtype_hotel_name")
        ]

    def __str__(self):
        return f"{self.hotel} • {self.name}"


class Room(models.Model):
    CLEAN = "clean"
    DIRTY = "dirty"
    IN_PROGRESS = "progress"
    CLEAN_CHOICES = (
        (CLEAN, "Убран"),
        (DIRTY, "Не убран"),
        (IN_PROGRESS, "Уборка"),
    )

    hotel = models.ForeignKey("dds.Point", on_delete=models.PROTECT, related_name="rooms", verbose_name="Точка")
    number = models.CharField(max_length=20, verbose_name="Номер/название")  # 101, A-1
    floor = models.IntegerField(default=1, verbose_name="Этаж")

    room_type = models.ForeignKey(RoomType, on_delete=models.PROTECT, related_name="rooms", verbose_name="Тип номера")
    capacity = models.PositiveSmallIntegerField(null=True, blank=True, verbose_name="Вместимость (если отличается)")

    day_rate = models.DecimalField(max_digits=12, decimal_places=2, default=0, verbose_name="Тариф/сутки")
    hour_rate = models.DecimalField(max_digits=12, decimal_places=2, default=0, verbose_name="Тариф/час")

    is_active = models.BooleanField(default=True, verbose_name="Активен")
    is_out_of_service = models.BooleanField(default=False, verbose_name="На ремонте")

    clean_status = models.CharField(max_length=10, choices=CLEAN_CHOICES, default=CLEAN, verbose_name="Уборка")

    class Meta:
        verbose_name = "Номер"
        verbose_name_plural = "Номера"
        ordering = ["hotel_id", "floor", "number"]
        constraints = [
            models.UniqueConstraint(fields=["hotel", "number"], name="uniq_room_hotel_number"),
        ]
        indexes = [
            models.Index(fields=["hotel", "is_active", "is_out_of_service"]),
            models.Index(fields=["hotel", "floor"]),
        ]

    def __str__(self):
        return f"{self.hotel} — {self.number}"

    @property
    def effective_capacity(self) -> int:
        return self.capacity or self.room_type.default_capacity

    @property
    def effective_day_rate(self) -> Decimal:
        return self.day_rate or self.room_type.default_day_rate

    @property
    def effective_hour_rate(self) -> Decimal:
        return self.hour_rate or self.room_type.default_hour_rate


class Company(models.Model):
    PAY_NOW = "now"
    PAY_WEEKLY = "weekly"
    PAY_INVOICE = "invoice"
    PAY_CHOICES = (
        (PAY_NOW, "Оплата сразу"),
        (PAY_WEEKLY, "Раз в неделю"),
        (PAY_INVOICE, "По счету"),
    )

    name = models.CharField(max_length=200, verbose_name="Компания")
    inn = models.CharField(max_length=14, blank=True, verbose_name="ИНН (опционально)")
    contact_name = models.CharField(max_length=200, blank=True, verbose_name="Контактное лицо")
    contact_phone = models.CharField(max_length=50, blank=True, verbose_name="Телефон")
    pay_terms = models.CharField(max_length=10, choices=PAY_CHOICES, default=PAY_NOW, verbose_name="Условия оплаты")
    is_active = models.BooleanField(default=True, verbose_name="Активна")

    class Meta:
        verbose_name = "Компания"
        verbose_name_plural = "Компании"
        ordering = ["name"]

    def __str__(self):
        return self.name

class CompanyFolio(models.Model):
    hotel = models.ForeignKey("dds.Point", on_delete=models.PROTECT, related_name="company_folios", verbose_name="Точка")
    company = models.ForeignKey(Company, on_delete=models.PROTECT, related_name="folios", verbose_name="Компания")

    is_closed = models.BooleanField(default=False, verbose_name="Закрыто")
    closed_at = models.DateTimeField(null=True, blank=True, verbose_name="Дата закрытия")

    created_at = models.DateTimeField(auto_now_add=True, verbose_name="Создано")

    class Meta:
        verbose_name = "Фолио компании"
        verbose_name_plural = "Фолио компаний"
        ordering = ["-created_at", "-id"]
        constraints = [
            models.UniqueConstraint(fields=["hotel", "company"], name="uniq_folio_hotel_company")
        ]
        indexes = [
            models.Index(fields=["hotel", "is_closed"]),
            models.Index(fields=["company"]),
        ]

    def __str__(self):
        return f"{self.hotel} • {self.company}"

    @property
    def balance(self) -> Decimal:
        agg = self.items.aggregate(s=Sum("signed_amount"))
        return agg["s"] or Decimal("0.00")

    def refresh_closed_flag(self):
        bal = self.balance
        if bal <= 0 and not self.is_closed:
            self.is_closed = True
            self.closed_at = timezone.now()
            self.save(update_fields=["is_closed", "closed_at"])
        elif bal > 0 and self.is_closed:
            self.is_closed = False
            self.closed_at = None
            self.save(update_fields=["is_closed", "closed_at"])


class CompanyFolioItem(models.Model):
    CHARGE = "charge"
    PAYMENT = "payment"
    ADJUST = "adjust"
    TYPE_CHOICES = (
        (CHARGE, "Начисление"),
        (PAYMENT, "Оплата"),
        (ADJUST, "Корректировка"),
    )

    folio = models.ForeignKey(CompanyFolio, on_delete=models.CASCADE, related_name="items", verbose_name="Фолио")
    item_type = models.CharField(max_length=12, null=True, blank=True, choices=TYPE_CHOICES, verbose_name="Тип")

    happened_at = models.DateTimeField(default=timezone.now, verbose_name="Дата")
    description = models.CharField(max_length=255, blank=True, verbose_name="Описание")

    amount = models.DecimalField(max_digits=12, decimal_places=2, verbose_name="Сумма")  # всегда +
    # signed_amount: charge/adjust = +, payment = -
    signed_amount = models.DecimalField(max_digits=12, blank=True, null=True, decimal_places=2, verbose_name="Сумма (со знаком)")

    stay_id = models.IntegerField(null=True, blank=True, verbose_name="ID проживания (если нужно)")  # MVP без FK
    dds_operation = models.ForeignKey("dds.DDSOperation", on_delete=models.SET_NULL, null=True, blank=True, related_name="folio_items")
    cash_movement = models.ForeignKey("dds.CashMovement", on_delete=models.SET_NULL, null=True, blank=True, related_name="folio_items")

    created_by = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.PROTECT, related_name="folio_items_created")
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name = "Строка фолио"
        verbose_name_plural = "Строки фолио"
        ordering = ["-happened_at", "-id"]
        indexes = [
            models.Index(fields=["folio", "happened_at"]),
            models.Index(fields=["item_type"]),
        ]

    def __str__(self):
        return f"{self.folio} • {self.get_item_type_display()} • {self.amount}"

    @staticmethod
    def make_signed(item_type: str, amount: Decimal) -> Decimal:
        a = amount or Decimal("0.00")
        return -a if item_type == CompanyFolioItem.PAYMENT else a


class Booking(models.Model):
    """
    Бронь (для УНО/отчетов). Из нее можно создать Stay (заселение).
    """
    OP_BOOK = "book"
    OP_CANCEL = "cancel"
    OP_EDIT = "edit"
    OP_CHOICES = (
        (OP_BOOK, "Бронь"),
        (OP_CANCEL, "Отмена"),
        (OP_EDIT, "Изменение"),
    )

    PAY_UNPAID = "unpaid"
    PAY_PARTIAL = "partial"
    PAY_PAID = "paid"
    PAY_CHOICES = (
        (PAY_UNPAID, "Не оплачено"),
        (PAY_PARTIAL, "Частично"),
        (PAY_PAID, "Оплачено"),
    )

    STATUS_NEW = "new"
    STATUS_CONFIRMED = "confirmed"
    STATUS_NO_SHOW = "no_show"
    STATUS_CANCELLED = "cancelled"
    STATUS_CHOICES = (
        (STATUS_NEW, "Новая"),
        (STATUS_CONFIRMED, "Подтверждена"),
        (STATUS_NO_SHOW, "Не заехал"),
        (STATUS_CANCELLED, "Отменена"),
    )

    hotel = models.ForeignKey("dds.Point", on_delete=models.PROTECT, related_name="bookings", verbose_name="Точка")

    shift_no = models.CharField(max_length=32, blank=True, verbose_name="№ смены")
    booked_at = models.DateTimeField(default=timezone.now, verbose_name="Дата брони")
    booking_number = models.CharField(max_length=64, verbose_name="Номер брони")

    stay_type = models.CharField(
        max_length=10,
        choices=(("private", "Частный"), ("corporate", "Корпоративный")),
        default="private",
        verbose_name="Тип"
    )
    company = models.ForeignKey(Company, on_delete=models.PROTECT, null=True, blank=True, related_name="bookings", verbose_name="Компания")

    guest_name = models.CharField(max_length=180, verbose_name="Гость (имя)")
    guest_phone = models.CharField(max_length=60, blank=True, verbose_name="Телефон")
    operation_type = models.CharField(max_length=10, choices=OP_CHOICES, default=OP_BOOK, verbose_name="Тип операции")

    guests_count = models.PositiveSmallIntegerField(default=1, verbose_name="Количество гостей")
    room_type = models.ForeignKey(RoomType, on_delete=models.PROTECT, null=True, blank=True, verbose_name="Тип номера")
    room = models.ForeignKey(Room, on_delete=models.PROTECT, null=True, blank=True, verbose_name="Номер")

    check_in = models.DateField(verbose_name="Дата заезда")
    check_out = models.DateField(verbose_name="Дата выезда")

    price_per_night = models.DecimalField(max_digits=12, decimal_places=2, default=0, verbose_name="Сумма за ночь")
    commission_amount = models.DecimalField(max_digits=12, decimal_places=2, default=0, verbose_name="Комиссия за бронь")
    gross_amount = models.DecimalField(max_digits=12, decimal_places=2, default=0, verbose_name="Общая сумма (до комиссии)")

    payment_status = models.CharField(max_length=10, choices=PAY_CHOICES, default=PAY_UNPAID, verbose_name="Статус оплаты")
    channel = models.CharField(max_length=80, blank=True, verbose_name="Канал брони")

    status = models.CharField(max_length=12, choices=STATUS_CHOICES, default=STATUS_NEW, verbose_name="Статус")

    created_by = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.PROTECT, related_name="bookings_created", verbose_name="Создал")
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name = "Бронирование"
        verbose_name_plural = "Бронирования"
        ordering = ["-booked_at", "-id"]
        constraints = [
            models.UniqueConstraint(fields=["hotel", "booking_number"], name="uniq_booking_hotel_number")
        ]
        indexes = [
            models.Index(fields=["hotel", "check_in", "check_out"]),
            models.Index(fields=["hotel", "booked_at"]),
        ]

    @property
    def nights(self) -> int:
        return max((self.check_out - self.check_in).days, 0)

    @property
    def net_amount(self) -> Decimal:
        return (self.gross_amount or Decimal("0.00")) - (self.commission_amount or Decimal("0.00"))

    def __str__(self):
        return f"{self.hotel} • бронь {self.booking_number}"


class Stay(models.Model):
    """
    Проживание (шахматка): бронь/заселение/выезд.
    Здесь держим связи с кассой и ДДС — как в ТЗ.
    """
    BOOKED = "booked"
    IN = "in"
    OUT = "out"
    CANCELED = "canceled"
    NO_SHOW = "no_show"
    STATUS_CHOICES = (
        (BOOKED, "Бронь"),
        (IN, "Проживает"),
        (OUT, "Выехал"),
        (CANCELED, "Отменено"),
        (NO_SHOW, "Не заехал"),
    )

    PRIVATE = "private"
    CORPORATE = "corporate"
    TYPE_CHOICES = ((PRIVATE, "Частный"), (CORPORATE, "Корпоративный"))

    hotel = models.ForeignKey("dds.Point", on_delete=models.PROTECT, related_name="stays", verbose_name="Точка")
    room = models.ForeignKey(Room, on_delete=models.PROTECT, related_name="stays", verbose_name="Номер")

    booking = models.ForeignKey(Booking, on_delete=models.SET_NULL, null=True, blank=True, related_name="stays", verbose_name="Бронь (источник)")

    stay_type = models.CharField(max_length=10, choices=TYPE_CHOICES, default=PRIVATE, verbose_name="Тип")
    company = models.ForeignKey(Company, on_delete=models.PROTECT, null=True, blank=True, related_name="stays", verbose_name="Компания")

    guest_name = models.CharField(max_length=200, blank=True, verbose_name="Гость (ФИО)")
    guest_phone = models.CharField(max_length=60, blank=True, verbose_name="Телефон")

    check_in = models.DateTimeField(verbose_name="Заезд")
    check_out = models.DateTimeField(verbose_name="Выезд")

    guests_count = models.PositiveSmallIntegerField(default=1, verbose_name="Кол-во гостей")
    channel = models.CharField(max_length=80, blank=True, verbose_name="Канал (для УНО)")

    amount = models.DecimalField(max_digits=12, decimal_places=2, default=0, verbose_name="Сумма")
    discount = models.DecimalField(max_digits=12, decimal_places=2, default=0, verbose_name="Скидка")
    tourist_tax_total = models.DecimalField(max_digits=12, decimal_places=2, default=0, verbose_name="Тур. сбор (итого)")
    comment = models.TextField(blank=True, verbose_name="Комментарий")

    status = models.CharField(max_length=12, choices=STATUS_CHOICES, default=BOOKED, verbose_name="Статус")

    # связи с финансами
    dds_operation = models.OneToOneField("dds.DDSOperation", on_delete=models.SET_NULL, null=True, blank=True, verbose_name="Запись ДДС")
    cash_movement = models.OneToOneField("dds.CashMovement", on_delete=models.SET_NULL, null=True, blank=True, verbose_name="Кассовое движение")

    created_by = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.PROTECT, related_name="stays_created", verbose_name="Создал")
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name = "Проживание (шахматка)"
        verbose_name_plural = "Проживания (шахматка)"
        ordering = ["-check_in", "-id"]
        indexes = [
            models.Index(fields=["hotel", "status"]),
            models.Index(fields=["hotel", "room", "check_in", "check_out"]),
            models.Index(fields=["company", "check_in"]),
        ]

    def __str__(self):
        return f"{self.hotel} {self.room.number} #{self.id}"

    @property
    def total_to_pay(self) -> Decimal:
        return (self.amount or Decimal("0.00")) - (self.discount or Decimal("0.00"))


class Guest(models.Model):
    hotel = models.ForeignKey("dds.Point", on_delete=models.PROTECT, related_name="guests", verbose_name="Точка")
    full_name = models.CharField(max_length=180, verbose_name="ФИО")
    inn = models.CharField(max_length=14, blank=True, verbose_name="ИНН (если есть)")
    nationality = models.CharField(max_length=80, blank=True, verbose_name="Национальность")
    is_foreigner = models.BooleanField(default=False, verbose_name="Иностранец")
    doc_number = models.CharField(max_length=80, blank=True, verbose_name="Документ (паспорт/ID)")

    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name = "Гость"
        verbose_name_plural = "Гости"
        ordering = ["hotel_id", "full_name"]
        indexes = [
            models.Index(fields=["hotel", "is_foreigner"]),
        ]

    def __str__(self):
        return self.full_name


class StayGuest(models.Model):
    stay = models.ForeignKey(Stay, on_delete=models.CASCADE, related_name="stay_guests", verbose_name="Проживание")
    guest = models.ForeignKey(Guest, on_delete=models.PROTECT, related_name="guest_stays", verbose_name="Гость")
    tourist_tax_amount = models.DecimalField(max_digits=12, decimal_places=2, default=0, verbose_name="Тур. налог (по гостю)")

    class Meta:
        verbose_name = "Гость в проживании"
        verbose_name_plural = "Гости в проживании"
        constraints = [
            models.UniqueConstraint(fields=["stay", "guest"], name="uniq_stay_guest")
        ]
