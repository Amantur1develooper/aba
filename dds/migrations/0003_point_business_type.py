from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ("dds", "0002_add_point_contact"),
    ]

    operations = [
        migrations.AddField(
            model_name="point",
            name="business_type",
            field=models.CharField(
                choices=[
                    ("hotel", "Отель"),
                    ("shop", "Магазин"),
                    ("restaurant", "Ресторан"),
                    ("construction", "Строительная компания"),
                    ("other", "Другое"),
                ],
                default="other",
                max_length=20,
                verbose_name="Вид бизнеса",
            ),
        ),
    ]
