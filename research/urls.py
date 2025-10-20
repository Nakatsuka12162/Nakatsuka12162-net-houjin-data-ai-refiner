from django.urls import path
from . import views

urlpatterns = [
    path('', views.home, name='home'),
    path('scrape/', views.run_scraping, name='scrape'),
    path('health/', views.health_check, name='health'),
]
