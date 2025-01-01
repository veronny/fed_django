from django.urls import path
from .views import index_s21_suplementacion6, get_redes_s21_suplementacion6, RptS21Suplementacion6Red, RptCoberturaS21Suplementacion6
from .views import get_microredes_s21_suplementacion6, p_microredes_s21_suplementacion6, RptS21Suplementacion6MicroRed
from .views import get_establecimientos_s21_suplementacion6, p_microredes_establec_s21_suplementacion6, p_establecimientos_s21_suplementacion6, RptS21Suplementacion6Establec

urlpatterns = [

    path('s21_suplementacion6/', index_s21_suplementacion6, name='index_s21_suplementacion6'),
    
    ### SEGUIMIENTO
    # redes
    path('get_redes_s21_suplementacion6/<int:redes_id>/', get_redes_s21_suplementacion6, name='get_redes_s21_suplementacion6'),
    #-- redes excel
    path('rpt_s21_suplementacion6_excel/', RptS21Suplementacion6Red.as_view(), name = 'rpt_s21_suplementacion6_red_xls'),
    
    
    # microredes
    path('get_microredes_s21_suplementacion6/<int:microredes_id>/', get_microredes_s21_suplementacion6, name='get_microredes_s21_suplementacion6'),
    path('p_microredes_s21_suplementacion6/', p_microredes_s21_suplementacion6, name='p_microredes_s21_suplementacion6'),
    #-- microredes excel
    path('rpt_s21_suplementacion6_microred_excel/', RptS21Suplementacion6MicroRed.as_view(), name = 'rpt_s21_suplementacion6_red_xls'),
    
    # establecimientos
    path('get_establecimientos_s21_suplementacion6/<int:establecimiento_id>/', get_establecimientos_s21_suplementacion6, name='get_establecimientos_s21_suplementacion6'),
    path('p_microredes_establec_s21_suplementacion6/', p_microredes_establec_s21_suplementacion6, name='p_microredes_establec_s21_suplementacion6'),
    path('p_establecimiento_s21_suplementacion6/', p_establecimientos_s21_suplementacion6, name='p_establecimientos_s21_suplementacion6'),       
    #-- estableccimiento excel
    path('rpt_s21_suplementacion6_establec_excel/', RptS21Suplementacion6Establec.as_view(), name = 'rpt_s21_suplementacion6_red_xls'),
    
    
    
    ### COBERTURA
    path('rpt_cobertura_s21_suplementacion6_excel/', RptCoberturaS21Suplementacion6.as_view(), name = 'rpt_cobertura_s21_suplementacion6_xls'),
    

]