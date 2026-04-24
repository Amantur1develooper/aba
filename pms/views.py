# pms/views.py
# from __future__ import annotations
# from __future__ import annotations
# grep -R "from __future__ import annotations" -n .

from collections import defaultdict
from datetime import date, datetime, time, timedelta
from django.urls import reverse
from django import forms
from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.shortcuts import get_object_or_404, redirect, render
from django.utils import timezone

from .models import Room, RoomType, Stay, Booking
from .services import (
    PMSConflictError,
    assert_no_overlap,
    check_in_stay,
    check_out_stay,
    cancel_stay,
)
from dds.models import DDSOperation, DDSArticle


# если у тебя есть ограничение "какие отели доступны пользователю" — используй свой helper
try:
    from dds.utils import _user_hotels_qs  # если у тебя так
except Exception:
    from dds.models import Point
    def _user_hotels_qs(user):
        return Point.objects.filter(is_active=True)

from typing import Optional
def _parse_date(s: str) -> Optional[date]:

# def _parse_date(s: str) -> date | None:
    try:
        return datetime.strptime(s, "%Y-%m-%d").date()
    except Exception:
        return None


def _month_range(d: date):
    start = d.replace(day=1)
    # следующий месяц
    if start.month == 12:
        end = start.replace(year=start.year + 1, month=1, day=1)
    else:
        end = start.replace(month=start.month + 1, day=1)
    return start, end


def _week_range(d: date):
    # понедельник..воскресенье
    start = d - timedelta(days=d.weekday())
    end = start + timedelta(days=7)
    return start, end


def _daterange(start: date, end: date):
    cur = start
    while cur < end:
        yield cur
        cur += timedelta(days=1)


class StayCreateForm(forms.ModelForm):
    class Meta:
        model = Stay
        fields = [
            "hotel", "room",
            "stay_type", "company",
            "guest_name", "guest_phone",
            "check_in", "check_out",
            "guests_count", "channel",
            "amount", "discount", "tourist_tax_total",
            "comment",
            "status",
        ]
        widgets = {
            "check_in": forms.DateTimeInput(attrs={"type": "datetime-local"}),
            "check_out": forms.DateTimeInput(attrs={"type": "datetime-local"}),
        }


class CheckInForm(forms.Form):
    pay_now = forms.BooleanField(required=False, initial=True, label="Оплата сейчас?")
    method = forms.ChoiceField(choices=DDSOperation.METHOD_CHOICES, label="Способ оплаты")
    paid_amount = forms.DecimalField(required=False, min_value=0, decimal_places=2, max_digits=12, label="Сумма оплаты (если частично)")
    article = forms.ModelChoiceField(
        queryset=DDSArticle.objects.filter(is_active=True, kind=DDSArticle.INCOME),
        required=False,
        label="Статья ДДС (доход)",
    )


from datetime import datetime, time, timedelta
from django.db.models import Q
from django.utils import timezone
from django.shortcuts import get_object_or_404, render
from django.contrib.auth.decorators import login_required

@login_required
def board(request):
    hotels = _user_hotels_qs(request.user)

    hotel_id = request.GET.get("hotel") or ""
    if hotel_id:
        selected_hotel = get_object_or_404(hotels, id=hotel_id)
    else:
        selected_hotel = hotels.first()

    if not selected_hotel:
        return render(request, "pms/board.html", {"hotels": hotels, "selected_hotel": None})

    view_mode = request.GET.get("view") or "week"   # week|month
    start_param = request.GET.get("start") or ""
    start_date = _parse_date(start_param) or timezone.localdate()

    floor = (request.GET.get("floor") or "").strip()
    room_type_id = request.GET.get("room_type") or ""

    # период
    if view_mode == "month":
        period_start, period_end = _month_range(start_date)  # end = первый день след. месяца (exclusive)
    else:
        period_start, period_end = _week_range(start_date)   # end = +7 дней (exclusive)

    days = list(_daterange(period_start, period_end))

    # главное: month -> делим на недели по 7 дней (без горизонтального скролла)
    if view_mode == "month":
        day_chunks = [days[i:i+7] for i in range(0, len(days), 7)]
    else:
        day_chunks = [days]

    tz = timezone.get_current_timezone()
    period_start_dt = timezone.make_aware(datetime.combine(period_start, time.min), timezone=tz)
    period_end_dt = timezone.make_aware(datetime.combine(period_end, time.min), timezone=tz)

    # номера
    rooms_qs = (
        Room.objects
        .filter(hotel=selected_hotel, is_active=True)
        .select_related("room_type")
        .order_by("floor", "number")
    )
    if floor:
        try:
            rooms_qs = rooms_qs.filter(floor=int(floor))
        except ValueError:
            pass
    if room_type_id:
        rooms_qs = rooms_qs.filter(room_type_id=room_type_id)

    rooms = list(rooms_qs)

    # статусы, которые НЕ показываем на шахматке
    # (чтобы не зависеть от твоих констант в модели)
    hidden_statuses = {"canceled", "cancelled", "no_show"}

    # проживание/брони, которые пересекаются с периодом
    stays_qs = (
        Stay.objects
        .filter(hotel=selected_hotel, room__in=rooms_qs)
        .exclude(status__in=hidden_statuses)
        .filter(check_in__lt=period_end_dt, check_out__gt=period_start_dt)
        .select_related("room", "company", "booking")
        .order_by("check_in")
    )
    stays = list(stays_qs)

    # cell_map: ключ "room_id:YYYY-MM-DD" -> stay
    cell_map = {}
    for st in stays:
        s = max(st.check_in, period_start_dt)
        e = min(st.check_out, period_end_dt)

        start_day = timezone.localtime(s, tz).date()
        end_day = timezone.localtime(e - timedelta(seconds=1), tz).date()

        for d in days:
            if start_day <= d <= end_day:
                k = f"{st.room_id}:{d.isoformat()}"
                # если вдруг пересечение (не должно быть) — оставляем первую запись
                cell_map.setdefault(k, st)

    room_types = RoomType.objects.filter(hotel=selected_hotel, is_active=True).order_by("name")

    context = {
        "hotels": hotels,
        "selected_hotel": selected_hotel,

        "view_mode": view_mode,
        "period_start": period_start,
        "period_end": period_end,
        "days": days,
        "day_chunks": day_chunks,   # ВАЖНО для шаблона без горизонтального скролла

        "floor": floor,
        "room_type_id": room_type_id,
        "room_types": room_types,

        "rooms": rooms,
        "cell_map": cell_map,
    }
    return render(request, "pms/board.html", context)



@login_required
def stay_create(request):
    hotels = _user_hotels_qs(request.user)
    hotel_id = request.GET.get("hotel") or ""
    room_id = request.GET.get("room") or ""
    day_s = request.GET.get("day") or ""

    initial = {}
    if hotel_id:
        initial["hotel"] = get_object_or_404(hotels, id=hotel_id)
    if room_id:
        initial["room"] = get_object_or_404(Room, id=room_id)
        initial["hotel"] = initial["room"].hotel

    # если кликнули на конкретный день — подставим чек-ин/чек-аут
    d = _parse_date(day_s)
    if d and "hotel" in initial:
        # ориентируемся на настройки отеля (если есть)
        settings = getattr(initial["hotel"], "pms_settings", None)
        ci_t = settings.check_in_time if settings else time(14, 0)
        co_t = settings.check_out_time if settings else time(12, 0)
        initial["check_in"] = timezone.make_aware(datetime.combine(d, ci_t))
        initial["check_out"] = timezone.make_aware(datetime.combine(d + timedelta(days=1), co_t))
        initial["status"] = Stay.BOOKED

    if request.method == "POST":
        form = StayCreateForm(request.POST)
        if form.is_valid():
            stay = form.save(commit=False)
            stay.created_by = request.user

            try:
                assert_no_overlap(room=stay.room, start_dt=stay.check_in, end_dt=stay.check_out, exclude_stay_id=None)
            except PMSConflictError as e:
                form.add_error(None, str(e))
                return render(request, "pms/stay_form.html", {"form": form})

            stay.save()
            messages.success(request, "Запись создана.")
            url = reverse("pms:board")
            return redirect(f"{url}?hotel={stay.hotel_id}")
            # return redirect("pms:board") + f"?hotel={stay.hotel_id}"
    else:
        form = StayCreateForm(initial=initial)

    return render(request, "pms/stay_form.html", {"form": form})


@login_required
def stay_edit(request, pk: int):
    stay = get_object_or_404(Stay, pk=pk)
    hotels = _user_hotels_qs(request.user)
    if stay.hotel_id not in hotels.values_list("id", flat=True):
        return redirect("pms:board")

    if request.method == "POST":
        form = StayCreateForm(request.POST, instance=stay)
        if form.is_valid():
            st = form.save(commit=False)

            try:
                assert_no_overlap(room=st.room, start_dt=st.check_in, end_dt=st.check_out, exclude_stay_id=st.id)
            except PMSConflictError as e:
                form.add_error(None, str(e))
                return render(request, "pms/stay_form.html", {"form": form, "stay": stay})

            st.save()
            messages.success(request, "Сохранено.")
            return redirect("pms:board") + f"?hotel={st.hotel_id}"
    else:
        form = StayCreateForm(instance=stay)

    return render(request, "pms/stay_form.html", {"form": form, "stay": stay})


@login_required
def stay_checkin(request, pk: int):
    stay = get_object_or_404(Stay, pk=pk)
    hotels = _user_hotels_qs(request.user)
    if stay.hotel_id not in hotels.values_list("id", flat=True):
        return redirect("pms:board")

    if request.method == "POST":
        form = CheckInForm(request.POST)
        if form.is_valid():
            pay_now = form.cleaned_data["pay_now"]
            method = form.cleaned_data["method"]
            paid_amount = form.cleaned_data["paid_amount"]
            article = form.cleaned_data["article"]

            try:
                check_in_stay(
                    stay=stay,
                    user=request.user,
                    pay_now=pay_now,
                    method=method,
                    paid_amount=paid_amount,
                    dds_article=article,
                )
                messages.success(request, "Заселение выполнено.")
                return redirect("pms:board") + f"?hotel={stay.hotel_id}"
            except Exception as e:
                messages.error(request, str(e))
    else:
        form = CheckInForm(initial={"method": DDSOperation.CASH, "pay_now": True})

    return render(request, "pms/stay_checkin.html", {"stay": stay, "form": form})


@login_required
def stay_checkout(request, pk: int):
    stay = get_object_or_404(Stay, pk=pk)
    hotels = _user_hotels_qs(request.user)
    if stay.hotel_id not in hotels.values_list("id", flat=True):
        return redirect("pms:board")

    try:
        check_out_stay(stay=stay, user=request.user)
        messages.success(request, "Выезд выполнен. Номер помечен как 'Не убран'.")
    except Exception as e:
        messages.error(request, str(e))

    return redirect("pms:board") + f"?hotel={stay.hotel_id}"


@login_required
def stay_cancel(request, pk: int):
    stay = get_object_or_404(Stay, pk=pk)
    hotels = _user_hotels_qs(request.user)
    if stay.hotel_id not in hotels.values_list("id", flat=True):
        return redirect("pms:board")

    try:
        cancel_stay(stay=stay, user=request.user, reason="Отмена из шахматки")
        messages.success(request, "Отменено.")
    except Exception as e:
        messages.error(request, str(e))

    return redirect("pms:board") + f"?hotel={stay.hotel_id}"
