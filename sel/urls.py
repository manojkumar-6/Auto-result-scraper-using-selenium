from django.contrib import admin
from django.urls import path
from . import views
urlpatterns = [
    path('', views.h,name="hello"),
    path("r",views.r,name="r"),
    path("z",views.z,name="z"),
]