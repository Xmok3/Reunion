from django.urls import path
from . import views

urlpatterns = [
    path('', views.register, name="register"),
    path('success/<str:code>/', views.success, name="success"),
    path('dashboard/', views.dashboard, name="dashboard"),
    path('export/csv/', views.export_csv, name="export_csv"),
    path('export/xlsx/', views.export_xlsx, name="export_xlsx"),
    path('import/', views.import_guests, name="import_guests"),
]
