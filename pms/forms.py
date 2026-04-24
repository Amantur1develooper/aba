from decimal import Decimal
from django import forms
from dds.models import DDSArticle, DDSOperation


class FolioPaymentForm(forms.Form):
    pay_amount = forms.DecimalField(label="Сумма оплаты", max_digits=12, decimal_places=2, min_value=Decimal("0.01"))

    method = forms.ChoiceField(
        label="Счёт/способ",
        choices=DDSOperation.METHOD_CHOICES,
    )

    article = forms.ModelChoiceField(
        label="Статья дохода (ДДС)",
        queryset=DDSArticle.objects.filter(kind=DDSArticle.INCOME, is_active=True).order_by("name"),
        required=False,
        empty_label="(авто: первая доходная статья)",
    )

    comment = forms.CharField(label="Комментарий", required=False, widget=forms.Textarea(attrs={"rows": 2}))
