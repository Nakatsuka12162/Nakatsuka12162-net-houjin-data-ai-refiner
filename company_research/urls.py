from django.urls import path, include

urlpatterns = [
    path('admin/', include('research.admin_urls')),
    path('', include('research.urls')),
]
