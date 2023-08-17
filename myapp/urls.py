from django.urls import path

from . import views

app_name = "myapp"

urlpatterns = [
    path('StudentsList', views.index, name='index'),
path('', views.CheckExceldata, name='CheckExceldata'),
]
