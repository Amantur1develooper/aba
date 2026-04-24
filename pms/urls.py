# pms/urls.py
from django.urls import path
from . import views, views_folio

app_name = "pms"

urlpatterns = [
    path("", views.board, name="board"),
    path("stay/add/", views.stay_create, name="stay_create"),
    path("stay/<int:pk>/edit/", views.stay_edit, name="stay_edit"),
    path("stay/<int:pk>/checkin/", views.stay_checkin, name="stay_checkin"),
    path("stay/<int:pk>/checkout/", views.stay_checkout, name="stay_checkout"),
    path("stay/<int:pk>/cancel/", views.stay_cancel, name="stay_cancel"),
    
    
    path("folios/", views_folio.folio_list, name="folio_list"),
    path("folios/<int:pk>/", views_folio.folio_detail, name="folio_detail"),
    path("folios/<int:pk>/pay/", views_folio.folio_payment, name="folio_payment"),
]
