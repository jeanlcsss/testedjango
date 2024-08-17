from django.contrib import admin
from django.urls import path
from . import views

urlpatterns = [
    path('', views.automacao, name='automacao'),
    path('success', views.success, name='success'),   
]