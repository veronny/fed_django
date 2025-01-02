from django.urls import path
from .views import index_v3_paquete_terapeutico, get_redes_v3_paquete_terapeutico, RptV3PaqueteTerapeuticoRed, RptCoberturaV3PaqueteTerapeutico
from .views import get_microredes_v3_paquete_terapeutico, p_microredes_v3_paquete_terapeutico, RptV3PaqueteTerapeuticoMicroRed
from .views import get_establecimientos_v3_paquete_terapeutico, p_microredes_establec_v3_paquete_terapeutico, p_establecimientos_v3_paquete_terapeutico, RptV3PaqueteTerapeuticoEstablec

urlpatterns = [

    path('v3_paquete_terapeutico/', index_v3_paquete_terapeutico, name='index_v3_paquete_terapeutico'),
    
    ### SEGUIMIENTO
    # redes
    path('get_redes_v3_paquete_terapeutico/<int:redes_id>/', get_redes_v3_paquete_terapeutico, name='get_redes_v3_paquete_terapeutico'),
    #-- redes excel
    path('rpt_v3_paquete_terapeutico_excel/', RptV3PaqueteTerapeuticoRed.as_view(), name = 'rpt_v3_paquete_terapeutico_red_xls'),
    
    # microredes
    path('get_microredes_v3_paquete_terapeutico/<int:microredes_id>/', get_microredes_v3_paquete_terapeutico, name='get_microredes_v3_paquete_terapeutico'),
    path('p_microredes_v3_paquete_terapeutico/', p_microredes_v3_paquete_terapeutico, name='p_microredes_v3_paquete_terapeutico'),
    #-- microredes excel
    path('rpt_v3_paquete_terapeutico_microred_excel/', RptV3PaqueteTerapeuticoMicroRed.as_view(), name = 'rpt_v3_paquete_terapeutico_red_xls'),
    
    # establecimientos
    path('get_establecimientos_v3_paquete_terapeutico/<int:establecimiento_id>/', get_establecimientos_v3_paquete_terapeutico, name='get_establecimientos_v3_paquete_terapeutico'),
    path('p_microredes_establec_v3_paquete_terapeutico/', p_microredes_establec_v3_paquete_terapeutico, name='p_microredes_establec_v3_paquete_terapeutico'),
    path('p_establecimiento_v3_paquete_terapeutico/', p_establecimientos_v3_paquete_terapeutico, name='p_establecimientos_v3_paquete_terapeutico'),       
    #-- estableccimiento excel
    path('rpt_v3_paquete_terapeutico_establec_excel/', RptV3PaqueteTerapeuticoEstablec.as_view(), name = 'rpt_v3_paquete_terapeutico_red_xls'),
    
    
    ### COBERTURA
    path('rpt_cobertura_v3_paquete_terapeutico_excel/', RptCoberturaV3PaqueteTerapeutico.as_view(), name = 'rpt_cobertura_v3_paquete_terapeutico_xls'),
    
]