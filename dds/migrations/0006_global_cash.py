from django.conf import settings
from django.db import migrations, models
import django.db.models.deletion
import django.utils.timezone


class Migration(migrations.Migration):

    dependencies = [
        ("dds", "0005_point_dates"),
        migrations.swappable_dependency(settings.AUTH_USER_MODEL),
    ]

    operations = [
        migrations.CreateModel(
            name="GlobalCashRegister",
            fields=[
                ("id", models.BigAutoField(auto_created=True, primary_key=True, serialize=False)),
                ("cash_balance",    models.DecimalField(decimal_places=2, default=0, max_digits=14, verbose_name="Наличные")),
                ("mkassa_balance",  models.DecimalField(decimal_places=2, default=0, max_digits=14, verbose_name="Банк1")),
                ("zadatok_balance", models.DecimalField(decimal_places=2, default=0, max_digits=14, verbose_name="Задаток")),
                ("optima_balance",  models.DecimalField(decimal_places=2, default=0, max_digits=14, verbose_name="Банк2")),
                ("updated_at", models.DateTimeField(auto_now=True)),
            ],
            options={"verbose_name": "Общая касса", "verbose_name_plural": "Общая касса"},
        ),
        migrations.CreateModel(
            name="GlobalCashOperation",
            fields=[
                ("id", models.BigAutoField(auto_created=True, primary_key=True, serialize=False)),
                ("direction",   models.CharField(choices=[("in", "Пополнение"), ("out", "Расход / Распределение")], max_length=3, verbose_name="Тип")),
                ("account",     models.CharField(choices=[("cash", "Наличные"), ("mkassa", "Банк1"), ("zadatok", "Задаток"), ("optima", "Банк2")], max_length=10, verbose_name="Счёт")),
                ("amount",      models.DecimalField(decimal_places=2, max_digits=14, verbose_name="Сумма")),
                ("happened_at", models.DateTimeField(default=django.utils.timezone.now, verbose_name="Дата")),
                ("comment",     models.TextField(blank=True, verbose_name="Комментарий")),
                ("created_at",  models.DateTimeField(auto_now_add=True)),
                ("created_by",  models.ForeignKey(on_delete=django.db.models.deletion.PROTECT, related_name="global_cash_ops", to=settings.AUTH_USER_MODEL)),
            ],
            options={"verbose_name": "Операция общей кассы", "verbose_name_plural": "Операции общей кассы", "ordering": ["-happened_at", "-id"]},
        ),
        migrations.CreateModel(
            name="GlobalCashDistribution",
            fields=[
                ("id", models.BigAutoField(auto_created=True, primary_key=True, serialize=False)),
                ("amount", models.DecimalField(decimal_places=2, max_digits=14, verbose_name="Сумма")),
                ("note",   models.CharField(blank=True, max_length=255, verbose_name="Примечание")),
                ("operation", models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, related_name="distributions", to="dds.globalcashoperation")),
                ("point",     models.ForeignKey(on_delete=django.db.models.deletion.PROTECT, related_name="global_distributions", to="dds.point")),
            ],
            options={"verbose_name": "Распределение", "verbose_name_plural": "Распределения"},
        ),
    ]
