from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ("dds", "0003_point_business_type"),
    ]

    operations = [
        migrations.AddField(
            model_name="point",
            name="website",
            field=models.URLField(blank=True, verbose_name="Сайт"),
        ),
        migrations.AddField(
            model_name="point",
            name="app_store_url",
            field=models.URLField(blank=True, verbose_name="App Store"),
        ),
        migrations.AddField(
            model_name="point",
            name="play_market_url",
            field=models.URLField(blank=True, verbose_name="Google Play"),
        ),
    ]
