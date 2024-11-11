from django.urls import path
from .views import index_s1_gestante_anemia, get_redes_s1_gestante_anemia, RptS1GestanteAnemiaRed, RptCoberturaS1GestanteAnemia


urlpatterns = [

    path('s1_gestante_anemia/', index_s1_gestante_anemia, name='index_s1_gestante_anemia'),
    
    ### SEGUIMIENTO
    # redes
    path('get_redes_s1_gestante_anemia/<int:redes_id>/', get_redes_s1_gestante_anemia, name='get_redes_s1_gestante_anemia'),
    #-- redes excel
    path('rpt_s1_gestante_anemia_red_excel/', RptS1GestanteAnemiaRed.as_view(), name = 'rpt_s1_gestante_anemia_red_xls'),
    
    
    ### COBERTURA
    path('rpt_cobertura_s1_gestante_anemia_excel/', RptCoberturaS1GestanteAnemia.as_view(), name = 'rpt_cobertura_s1_gestante_anemia_xls'),
    
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