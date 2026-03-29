from django.contrib import admin
from django.urls import path
from merger import views

urlpatterns = [
    path('admin/', admin.site.urls),  # ✅ fixed এখানে
    path('', views.index, name='index'),
]