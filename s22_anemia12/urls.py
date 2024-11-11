from django.urls import path
from .views import index_s22_anemia12, get_redes_s22_anemia12, RptS22Anemia12Red, RptCoberturaS22Anemia12


urlpatterns = [

    path('s22_anemia12/', index_s22_anemia12, name='index_s22_anemia12'),
    
    ### SEGUIMIENTO
    # redes
    path('get_redes_s22_anemia12/<int:redes_id>/', get_redes_s22_anemia12, name='get_redes_s22_anemia12'),
    #-- redes excel
    path('rpt_s22_anemia12_excel/', RptS22Anemia12Red.as_view(), name = 'rpt_s22_anemia12_red_xls'),
    
    ### COBERTURA
    path('rpt_cobertura_s22_anemia12_excel/', RptCoberturaS22Anemia12.as_view(), name = 'rpt_cobertura_s22_anemia12_xls'),
    
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