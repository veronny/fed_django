from django.urls import path
from .views import index_paquete_nino, get_redes_paquete_nino, RptPaqueteNinoRed, RptCoberturaPaqueteNino
from .views import get_microredes_paquete_nino, p_microredes_paquete_nino, RptPaqueteNinoMicroRed
from .views import get_establecimientos_paquete_nino, p_microredes_establec_paquete_nino, p_establecimientos_paquete_nino, RptPaqueteNinoEstablec

urlpatterns = [

    path('paquete_nino/', index_paquete_nino, name='index_paquete_nino'),
    
    ### SEGUIMIENTO
    # redes
    path('get_redes_paquete_nino/<int:redes_id>/', get_redes_paquete_nino, name='get_redes_paquete_nino'),
    #-- redes excel
    path('rpt_paquete_nino_red_excel/', RptPaqueteNinoRed.as_view(), name = 'rpt_paquete_nino_red_xls'),
    
    # microredes
    path('get_microredes_paquete_nino/<int:microredes_id>/', get_microredes_paquete_nino, name='get_microredes_paquete_nino'),
    path('p_microredes_paquete_nino/', p_microredes_paquete_nino, name='p_microredes_paquete_nino'),
    #-- microredes excel
    path('rpt_paquete_nino_microred_excel/', RptPaqueteNinoMicroRed.as_view(), name = 'rpt_paquete_nino_red_xls'),
    
    # establecimientos
    path('get_establecimientos_paquete_nino/<int:establecimiento_id>/', get_establecimientos_paquete_nino, name='get_establecimientos_paquete_nino'),
    path('p_microredes_establec_paquete_nino/', p_microredes_establec_paquete_nino, name='p_microredes_establec_paquete_nino'),
    path('p_establecimiento_paquete_nino/', p_establecimientos_paquete_nino, name='p_establecimientos_paquete_nino'),       
    #-- estableccimiento excel
    path('rpt_paquete_nino_establec_excel/', RptPaqueteNinoEstablec.as_view(), name = 'rpt_paquete_nino_red_xls'),
    
    ### COBERTURA
    path('rpt_cobertura_nino_red_excel/', RptCoberturaPaqueteNino.as_view(), name = 'rpt_cobertura_paquete_nino_red_xls'),
]