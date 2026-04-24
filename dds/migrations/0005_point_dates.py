from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ("dds", "0004_point_links"),
    ]

    operations = [
        migrations.AddField(
            model_name="point",
            name="launch_date",
            field=models.DateField(blank=True, null=True, verbose_name="Дата запуска"),
        ),
        migrations.AddField(
            model_name="point",
            name="payment_date",
            field=models.DateField(blank=True, null=True, verbose_name="Дата оплаты"),
        ),
    ]
