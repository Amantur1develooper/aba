from .models import Point

def user_hotels_qs(user):
    return Point.objects.filter(is_active=True)
