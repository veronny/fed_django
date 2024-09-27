from django.contrib import admin
from django.urls import path, include
from django.conf.urls.static import static
from django.conf import settings

urlpatterns = [
    path('admin/', admin.site.urls),
    path('', include('base.urls')),
    path('', include('discapacidad.urls')),
    path('', include('discapacidad.padron_urls')),
    path('', include('paquete_neonatal.urls')),
    path('', include('paquete_gestante.urls')),
    path('', include('paquete_nino.urls')),
    path('', include('s1_gestante_anemia.urls')),
]

if settings.DEBUG:
    urlpatterns += static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)
    urlpatterns += static(settings.STATIC_URL, document_root=settings.STATIC_ROOT)
    
    