from django.urls import path
from . import views
from .views_cash import cash_transfer_create, dds_op_add, dds_op_add, dds_articles_json


app_name = "dds"

urlpatterns = [
    path("", views.dds_dashboard, name="dds_dashboard"),
    path("list/", views.dds_list, name="dds_list"),
    path("create/", views.dds_create, name="dds_create"),
    path("void/<int:pk>/", views.dds_void, name="dds_void"),
    path("articles/", views.dds_articles, name="dds_articles"),
    
    path("hotels/", views.hotel_catalog, name="hotel_catalog"),
    path("hotels2/", views.hotel_list, name="hotel_list"),
    path("hotels/<int:pk>/", views.hotel_detail, name="hotel_detail"),

    path("hotels/<int:pk>/export/excel/", views.hotel_detail_export_excel, name="hotel_export_excel"),
    path("report/export/excel/", views.unified_report_export_excel, name="unified_report_excel"),

    path("report/", views.unified_report, name="unified_report"),

    path("cash-transfer/add/<int:hotel_id>/", cash_transfer_create, name="cash_transfer_create"),
    # path("cash/transfer/add/", transfer_create, name="transfer_create"),
    
    
    path("hotel/<int:hotel_id>/income/add/", dds_op_add, {"kind": "income"}, name="income_add"),
    path("hotel/<int:hotel_id>/expense/add/", dds_op_add, {"kind": "expense"}, name="expense_add"),
    # path("hotel/<int:hotel_id>/income/add/", dds_op_add, {"kind": "income"}, name="income_add"),
    # path("hotel/<int:hotel_id>/expense/add/", dds_op_add, {"kind": "expense"}, name="expense_add"),
    #  path("hotel/<int:hotel_id>/income/add/", views.dds_op_add, {"kind": "income"}, name="income_add"),
    # path("hotel/<int:hotel_id>/expense/add/", views.dds_op_add, {"kind": "expense"}, name="expense_add"),
    # path("articles/", dds_articles_json, name="articles_json"),
    # Ajax: статьи по категории
    path("articles/", dds_articles_json, name="articles_json"),
    

    
    
    
    path("global-cash/", views.global_cash_view, name="global_cash"),

    path("hotels/<int:hotel_id>/contacts/add/", views.contact_add, name="contact_add"),
    path("hotels/<int:hotel_id>/contacts/<int:contact_id>/delete/", views.contact_delete, name="contact_delete"),

    path("hotels/<int:pk>/incasso/", views.incasso_create, name="incasso_create"),
    path("accounting/", views.accounting, name="accounting"),
    path("accounting/export/excel/", views.accounting_export_excel, name="accounting_excel"),

]
