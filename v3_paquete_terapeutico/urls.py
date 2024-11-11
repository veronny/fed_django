from django.urls import path
from .views import index_v3_paquete_terapeutico, get_redes_v3_paquete_terapeutico, RptV1CondicionPreviaRed, RptCoberturaV3PaqueteTerapeutico


urlpatterns = [

    path('v3_paquete_terapeutico/', index_v3_paquete_terapeutico, name='index_v3_paquete_terapeutico'),
    
    ### SEGUIMIENTO
    # redes
    path('get_redes_v3_paquete_terapeutico/<int:redes_id>/', get_redes_v3_paquete_terapeutico, name='get_redes_v3_paquete_terapeutico'),
    #-- redes excel
    path('rpt_v3_paquete_terapeutico_excel/', RptV1CondicionPreviaRed.as_view(), name = 'rpt_v3_paquete_terapeutico_red_xls'),
    
    ### COBERTURA
    path('rpt_cobertura_v3_paquete_terapeutico_excel/', RptCoberturaV3PaqueteTerapeutico.as_view(), name = 'rpt_cobertura_v3_paquete_terapeutico_xls'),
    
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