from django.urls import path
from .views import index_paquete_nino, get_redes_paquete_nino, RptPaqueteNinoRed, RptCoberturaPaqueteNino


urlpatterns = [

    path('paquete_nino/', index_paquete_nino, name='index_paquete_nino'),
    
    ### SEGUIMIENTO
    # redes
    path('get_redes_paquete_nino/<int:redes_id>/', get_redes_paquete_nino, name='get_redes_paquete_nino'),
    #-- redes excel
    path('rpt_paquete_nino_red_excel/', RptPaqueteNinoRed.as_view(), name = 'rpt_paquete_nino_red_xls'),
    
    ### COBERTURA
    path('rpt_cobertura_nino_red_excel/', RptCoberturaPaqueteNino.as_view(), name = 'rpt_cobertura_paquete_nino_red_xls'),
    
    
    #microredes
    # path('get_microredes/<int:microredes_id>/', views.get_microredes, name='get_microredes'),
    # path('p_microredes/', views.p_microredes, name='p_microredes'),
    # #-- microredes excel
    # path('rpt_operacional_microred_excel/', RptOperacinalMicroRed.as_view(), name = 'rpt_operacional_microred_xls'),
    # 
    #establecimientos
    # path('get_establecimientos/<int:establecimiento_id>/', views.get_establecimientos, name='get_establecimientos'),
    # path('p_microredes_establec/', views.p_microredes_establec, name='p_microredes_establec'),
    # path('p_establecimiento/', views.p_establecimientos, name='p_establecimientos'),    
]