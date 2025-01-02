from django.urls import path
from .views import index_v2_tamizaje_violencia, get_redes_v2_tamizaje_violencia, RptV2TamizajeViolenciaRed, RptCoberturaV2TamizajeViolencia
from .views import get_microredes_v2_tamizaje_violencia, p_microredes_v2_tamizaje_violencia, RptV2TamizajeViolenciaMicroRed
from .views import get_establecimientos_v2_tamizaje_violencia, p_microredes_establec_v2_tamizaje_violencia, p_establecimientos_v2_tamizaje_violencia, RptV2TamizajeViolenciaEstablec

urlpatterns = [

    path('v2_tamizaje_violencia/', index_v2_tamizaje_violencia, name='index_v2_tamizaje_violencia'),
    
    ### SEGUIMIENTO
    # redes
    path('get_redes_v2_tamizaje_violencia/<int:redes_id>/', get_redes_v2_tamizaje_violencia, name='get_redes_v2_tamizaje_violencia'),
    #-- redes excel
    path('rpt_v2_tamizaje_violencia_excel/', RptV2TamizajeViolenciaRed.as_view(), name = 'rpt_v2_tamizaje_violencia_red_xls'),
    
    # microredes
    path('get_microredes_v2_tamizaje_violencia/<int:microredes_id>/', get_microredes_v2_tamizaje_violencia, name='get_microredes_v2_tamizaje_violencia'),
    path('p_microredes_v2_tamizaje_violencia/', p_microredes_v2_tamizaje_violencia, name='p_microredes_v2_tamizaje_violencia'),
    #-- microredes excel
    path('rpt_v2_tamizaje_violencia_microred_excel/', RptV2TamizajeViolenciaMicroRed.as_view(), name = 'rpt_v2_tamizaje_violencia_red_xls'),
    
    # establecimientos
    path('get_establecimientos_v2_tamizaje_violencia/<int:establecimiento_id>/', get_establecimientos_v2_tamizaje_violencia, name='get_establecimientos_v2_tamizaje_violencia'),
    path('p_microredes_establec_v2_tamizaje_violencia/', p_microredes_establec_v2_tamizaje_violencia, name='p_microredes_establec_v2_tamizaje_violencia'),
    path('p_establecimiento_v2_tamizaje_violencia/', p_establecimientos_v2_tamizaje_violencia, name='p_establecimientos_v2_tamizaje_violencia'),       
    #-- estableccimiento excel
    path('rpt_v2_tamizaje_violencia_establec_excel/', RptV2TamizajeViolenciaEstablec.as_view(), name = 'rpt_v2_tamizaje_violencia_red_xls'),
    
    ### COBERTURA
    path('rpt_cobertura_v2_tamizaje_violencia_excel/', RptCoberturaV2TamizajeViolencia.as_view(), name = 'rpt_cobertura_v2_tamizaje_violencia_xls'),
    

]