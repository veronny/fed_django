from django.urls import path
from .views import index_v1_condicion_previa, get_redes_v1_condicion_previa, RptV1CondicionPreviaRed, RptCoberturaV1CondicionPrevia
from .views import get_microredes_v1_condicion_previa, p_microredes_v1_condicion_previa, RptV1CondicionPreviaMicroRed
from .views import get_establecimientos_v1_condicion_previa, p_microredes_establec_v1_condicion_previa, p_establecimientos_v1_condicion_previa, RptV1CondicionPreviaEstablec

urlpatterns = [

    path('v1_condicion_previa/', index_v1_condicion_previa, name='index_v1_condicion_previa'),
    
    ### SEGUIMIENTO
    # redes
    path('get_redes_v1_condicion_previa/<int:redes_id>/', get_redes_v1_condicion_previa, name='get_redes_v1_condicion_previa'),
    #-- redes excel
    path('rpt_v1_condicion_previa_excel/', RptV1CondicionPreviaRed.as_view(), name = 'rpt_v1_condicion_previa_red_xls'),
    
    
    # microredes
    path('get_microredes_v1_condicion_previa/<int:microredes_id>/', get_microredes_v1_condicion_previa, name='get_microredes_v1_condicion_previa'),
    path('p_microredes_v1_condicion_previa/', p_microredes_v1_condicion_previa, name='p_microredes_v1_condicion_previa'),
    #-- microredes excel
    path('rpt_v1_condicion_previa_microred_excel/', RptV1CondicionPreviaMicroRed.as_view(), name = 'rpt_v1_condicion_previa_red_xls'),
    
    # establecimientos
    path('get_establecimientos_v1_condicion_previa/<int:establecimiento_id>/', get_establecimientos_v1_condicion_previa, name='get_establecimientos_v1_condicion_previa'),
    path('p_microredes_establec_v1_condicion_previa/', p_microredes_establec_v1_condicion_previa, name='p_microredes_establec_v1_condicion_previa'),
    path('p_establecimiento_v1_condicion_previa/', p_establecimientos_v1_condicion_previa, name='p_establecimientos_v1_condicion_previa'),       
    #-- estableccimiento excel
    path('rpt_v1_condicion_previa_establec_excel/', RptV1CondicionPreviaEstablec.as_view(), name = 'rpt_v1_condicion_previa_red_xls'),
    
    
    
    ### COBERTURA
    path('rpt_cobertura_v1_condicion_previa_excel/', RptCoberturaV1CondicionPrevia.as_view(), name = 'rpt_cobertura_v1_condicion_previa_xls'),
    
 
]