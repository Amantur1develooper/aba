from django.contrib.auth.decorators import login_required
from django.db.models import Q
from django.shortcuts import get_object_or_404, render, redirect
from django.urls import reverse

from dds.views import _user_hotels_qs  # или твоя функция user_hotels_qs
from .models import CompanyFolio
from .forms import FolioPaymentForm
from .services import folio_add_payment


@login_required
def folio_list(request):
    hotels_qs = _user_hotels_qs(request.user)

    hotel_id = request.GET.get("hotel") or ""
    q = (request.GET.get("q") or "").strip()
    status = request.GET.get("status") or "open"  # open/closed/all

    qs = CompanyFolio.objects.select_related("hotel", "company").filter(hotel__in=hotels_qs)

    if hotel_id:
        qs = qs.filter(hotel_id=hotel_id)
    if q:
        qs = qs.filter(company__name__icontains=q)

    if status == "open":
        qs = qs.filter(is_closed=False)
    elif status == "closed":
        qs = qs.filter(is_closed=True)

    # баланс посчитаем на странице через property (MVP). Если много данных — оптимизируем позже.
    return render(request, "pms/folio_list.html", {
        "hotels": hotels_qs,
        "rows": qs.order_by("-id")[:500],
        "hotel_id": hotel_id,
        "q": q,
        "status": status,
    })


@login_required
def folio_detail(request, pk: int):
    hotels_qs = _user_hotels_qs(request.user)

    folio = get_object_or_404(
        CompanyFolio.objects.select_related("hotel", "company"),
        pk=pk,
        hotel__in=hotels_qs,
    )

    items = folio.items.select_related("dds_operation", "cash_movement").order_by("-happened_at", "-id")[:300]

    return render(request, "pms/folio_detail.html", {
        "folio": folio,
        "items": items,
        "balance": folio.balance,
    })


@login_required
def folio_payment(request, pk: int):
    hotels_qs = _user_hotels_qs(request.user)

    folio = get_object_or_404(
        CompanyFolio.objects.select_related("hotel", "company"),
        pk=pk,
        hotel__in=hotels_qs,
    )

    if request.method == "POST":
        form = FolioPaymentForm(request.POST)
        if form.is_valid():
            item = folio_add_payment(
                folio=folio,
                user=request.user,
                amount=form.cleaned_data["pay_amount"],
                method=form.cleaned_data["method"],
                article=form.cleaned_data.get("article"),
                comment=form.cleaned_data.get("comment") or "",
            )
            return redirect(reverse("pms:folio_detail", args=[folio.id]))
    else:
        form = FolioPaymentForm()

    return render(request, "pms/folio_payment.html", {
        "folio": folio,
        "form": form,
        "balance": folio.balance,
    })
