from django.urls import path
from .views import index_s21_suplementacion6, get_redes_s21_suplementacion6, RptS21Suplementacion6Red, RptCoberturaS21Suplementacion6


urlpatterns = [

    path('s21_suplementacion6/', index_s21_suplementacion6, name='index_s21_suplementacion6'),
    
    ### SEGUIMIENTO
    # redes
    path('get_redes_s21_suplementacion6/<int:redes_id>/', get_redes_s21_suplementacion6, name='get_redes_s21_suplementacion6'),
    #-- redes excel
    path('rpt_s21_suplementacion6_excel/', RptS21Suplementacion6Red.as_view(), name = 'rpt_s21_suplementacion6_red_xls'),
    
    ### COBERTURA
    path('rpt_cobertura_s21_suplementacion6_excel/', RptCoberturaS21Suplementacion6.as_view(), name = 'rpt_cobertura_s21_suplementacion6_xls'),
    
    
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