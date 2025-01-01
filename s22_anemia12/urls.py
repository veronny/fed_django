from django.urls import path
from .views import index_s22_anemia12, get_redes_s22_anemia12, RptS22Anemia12Red, RptCoberturaS22Anemia12
from .views import get_microredes_s22_anemia12, p_microredes_s22_anemia12, RptS22Anemia12MicroRed
from .views import get_establecimientos_s22_anemia12, p_microredes_establec_s22_anemia12, p_establecimientos_s22_anemia12, RptS22Anemia12Establec

urlpatterns = [

    path('s22_anemia12/', index_s22_anemia12, name='index_s22_anemia12'),
    
    ### SEGUIMIENTO
    # redes
    path('get_redes_s22_anemia12/<int:redes_id>/', get_redes_s22_anemia12, name='get_redes_s22_anemia12'),
    #-- redes excel
    path('rpt_s22_anemia12_excel/', RptS22Anemia12Red.as_view(), name = 'rpt_s22_anemia12_red_xls'),
    
    
    
    # microredes
    path('get_microredes_s22_anemia12/<int:microredes_id>/', get_microredes_s22_anemia12, name='get_microredes_s22_anemia12'),
    path('p_microredes_s22_anemia12/', p_microredes_s22_anemia12, name='p_microredes_s22_anemia12'),
    #-- microredes excel
    path('rpt_s22_anemia12_microred_excel/', RptS22Anemia12MicroRed.as_view(), name = 'rpt_s22_anemia12_red_xls'),
    
    # establecimientos
    path('get_establecimientos_s22_anemia12/<int:establecimiento_id>/', get_establecimientos_s22_anemia12, name='get_establecimientos_s22_anemia12'),
    path('p_microredes_establec_s22_anemia12/', p_microredes_establec_s22_anemia12, name='p_microredes_establec_s22_anemia12'),
    path('p_establecimiento_s22_anemia12/', p_establecimientos_s22_anemia12, name='p_establecimientos_s22_anemia12'),       
    #-- estableccimiento excel
    path('rpt_s22_anemia12_establec_excel/', RptS22Anemia12Establec.as_view(), name = 'rpt_s22_anemia12_red_xls'),
    
    
    
    ### COBERTURA
    path('rpt_cobertura_s22_anemia12_excel/', RptCoberturaS22Anemia12.as_view(), name = 'rpt_cobertura_s22_anemia12_xls'),

]