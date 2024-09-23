from django.urls import path
from .views import index_paquete_neonatal, get_redes_paquete_neonatal, RptPaqueteNeonatalRed


urlpatterns = [

    path('paquete_neonatal/', index_paquete_neonatal, name='index_paquete_neonatal'),
    
    ### SEGUIMIENTO
    # redes
    path('get_redes_paquete_neonatal/<int:redes_id>/', get_redes_paquete_neonatal, name='get_redes_paquete_neonatal'),
    #-- redes excel
    path('rpt_paquete_neonatal_red_excel/', RptPaqueteNeonatalRed.as_view(), name = 'rpt_paquete_neonatal_red_xls'),
    
    #microredes
    # path('get_microredes/<int:microredes_id>/', views.get_microredes, name='get_microredes'),
    # path('p_microredes/', views.p_microredes, name='p_microredes'),
    # #-- microredes excel
    # path('rpt_operacional_microred_excel/', RptOperacinalMicroRed.as_view(), name = 'rpt_operacional_microred_xls'),
    # 
    # # establecimientos
    # path('get_establecimientos/<int:establecimiento_id>/', views.get_establecimientos, name='get_establecimientos'),
    # path('p_microredes_establec/', views.p_microredes_establec, name='p_microredes_establec'),
    # path('p_establecimiento/', views.p_establecimientos, name='p_establecimientos'),    
]