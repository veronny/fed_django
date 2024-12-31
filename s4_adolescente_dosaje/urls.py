from django.urls import path
from .views import index_s4_adolescente_dosaje, get_redes_s4_adolescente_dosaje, RptS4AdolescenteDosajeRed, RptCoberturaS4AdolescenteDosaje
from .views import get_microredes_s4_adolescente_dosaje, p_microredes_s4_adolescente_dosaje, RptS4AdolescenteDosajeMicroRed
from .views import get_establecimientos_s4_adolescente_dosaje, p_microredes_establec_s4_adolescente_dosaje, p_establecimientos_s4_adolescente_dosaje, RptS4AdolescenteDosajeEstablec

urlpatterns = [

    path('s4_adolescente_dosaje/', index_s4_adolescente_dosaje, name='index_s4_adolescente_dosaje'),
    
    ### SEGUIMIENTO
    # redes
    path('get_redes_s4_adolescente_dosaje/<int:redes_id>/', get_redes_s4_adolescente_dosaje, name='get_redes_s4_adolescente_dosaje'),
    #-- redes excel
    path('rpt_s4_adolescente_dosaje_excel/', RptS4AdolescenteDosajeRed.as_view(), name = 'rpt_s4_adolescente_dosaje_red_xls'),
    
    # microredes
    path('get_microredes_s4_adolescente_dosaje/<int:microredes_id>/', get_microredes_s4_adolescente_dosaje, name='get_microredes_s4_adolescente_dosaje'),
    path('p_microredes_s4_adolescente_dosaje/', p_microredes_s4_adolescente_dosaje, name='p_microredes_s4_adolescente_dosaje'),
    #-- microredes excel
    path('rpt_s4_adolescente_dosaje_microred_excel/', RptS4AdolescenteDosajeMicroRed.as_view(), name = 'rpt_s4_adolescente_dosaje_red_xls'),
    
    # establecimientos
    path('get_establecimientos_s4_adolescente_dosaje/<int:establecimiento_id>/', get_establecimientos_s4_adolescente_dosaje, name='get_establecimientos_s4_adolescente_dosaje'),
    path('p_microredes_establec_s4_adolescente_dosaje/', p_microredes_establec_s4_adolescente_dosaje, name='p_microredes_establec_s4_adolescente_dosaje'),
    path('p_establecimiento_s4_adolescente_dosaje/', p_establecimientos_s4_adolescente_dosaje, name='p_establecimientos_s4_adolescente_dosaje'),       
    #-- estableccimiento excel
    path('rpt_s4_adolescente_dosaje_establec_excel/', RptS4AdolescenteDosajeEstablec.as_view(), name = 'rpt_s4_adolescente_dosaje_red_xls'),
    
    
    
    ### COBERTURA
    path('rpt_cobertura_s4_adolescente_dosaje_excel/', RptCoberturaS4AdolescenteDosaje.as_view(), name = 'rpt_cobertura_s4_adolescente_dosaje_xls'),
    
    
]