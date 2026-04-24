from .models import Point

def user_hotels_qs(user):
    """
    - superuser видит всё
    - finance_admin (в профиле) видит всё
    - обычный пользователь видит только свой отель
    """
    if user.is_superuser:
        return Point.objects.filter(is_active=True)

    profile = getattr(user, "profile", None)
    if profile and profile.is_finance_admin:
        return Point.objects.filter(is_active=True)

    if profile and profile.hotel_id:
        return Point.objects.filter(is_active=True, id=profile.hotel_id)

    return Point.objects.none()
