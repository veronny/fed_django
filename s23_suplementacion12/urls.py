from django.urls import path
from .views import index_s23_suplementacion12, get_redes_s23_suplementacion12, RptS23Suplementacion12Red, RptCoberturaS23Suplementacion12
from .views import get_microredes_s23_suplementacion12, p_microredes_s23_suplementacion12, RptS23Suplementacion12MicroRed
from .views import get_establecimientos_s23_suplementacion12, p_microredes_establec_s23_suplementacion12, p_establecimientos_s23_suplementacion12, RptS23Suplementacion12Establec

urlpatterns = [

    path('s23_suplementacion12/', index_s23_suplementacion12, name='index_s23_suplementacion12'),
    
    ### SEGUIMIENTO
    # redes
    path('get_redes_s23_suplementacion12/<int:redes_id>/', get_redes_s23_suplementacion12, name='get_redes_s23_suplementacion12'),
    #-- redes excel
    path('rpt_s23_suplementacion12_excel/', RptS23Suplementacion12Red.as_view(), name = 'rpt_s23_suplementacion12_red_xls'),
    
    # microredes
    path('get_microredes_s23_suplementacion12/<int:microredes_id>/', get_microredes_s23_suplementacion12, name='get_microredes_s23_suplementacion12'),
    path('p_microredes_s23_suplementacion12/', p_microredes_s23_suplementacion12, name='p_microredes_s23_suplementacion12'),
    #-- microredes excel
    path('rpt_s23_suplementacion12_microred_excel/', RptS23Suplementacion12MicroRed.as_view(), name = 'rpt_s23_suplementacion12_red_xls'),
    
    # establecimientos
    path('get_establecimientos_s23_suplementacion12/<int:establecimiento_id>/', get_establecimientos_s23_suplementacion12, name='get_establecimientos_s23_suplementacion12'),
    path('p_microredes_establec_s23_suplementacion12/', p_microredes_establec_s23_suplementacion12, name='p_microredes_establec_s23_suplementacion12'),
    path('p_establecimiento_s23_suplementacion12/', p_establecimientos_s23_suplementacion12, name='p_establecimientos_s23_suplementacion12'),       
    #-- estableccimiento excel
    path('rpt_s23_suplementacion12_establec_excel/', RptS23Suplementacion12Establec.as_view(), name = 'rpt_s23_suplementacion12_red_xls'),
    
    ### COBERTURA
    path('rpt_cobertura_s23_suplementacion12_excel/', RptCoberturaS23Suplementacion12.as_view(), name = 'rpt_cobertura_s23_suplementacion12_xls'),

]