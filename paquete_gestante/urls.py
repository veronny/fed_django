from django.urls import path
from .views import index_paquete_gestante, get_redes_paquete_gestante, RptPaqueteGestanteRed, RptCoberturaPaqueteGestante
from .views import get_microredes_paquete_gestante, p_microredes_paquete_gestante, RptPaqueteGestanteMicroRed
from .views import get_establecimientos_paquete_gestante, p_microredes_establec_paquete_gestante, p_establecimientos_paquete_gestante, RptPaqueteGestanteEstablec

urlpatterns = [

    path('paquete_gestante/', index_paquete_gestante, name='index_paquete_gestante'),
    
    ### SEGUIMIENTO
    # redes
    path('get_redes_paquete_gestante/<int:redes_id>/', get_redes_paquete_gestante, name='get_redes_paquete_gestante'),
    #-- redes excel
    path('rpt_paquete_gestante_red_excel/', RptPaqueteGestanteRed.as_view(), name = 'rpt_paquete_gestante_red_xls'),
    
    # microredes
    path('get_microredes_paquete_gestante/<int:microredes_id>/', get_microredes_paquete_gestante, name='get_microredes_paquete_gestante'),
    path('p_microredes_paquete_gestante/', p_microredes_paquete_gestante, name='p_microredes_paquete_gestante'),
    #-- microredes excel
    path('rpt_paquete_gestante_microred_excel/', RptPaqueteGestanteMicroRed.as_view(), name = 'rpt_paquete_gestante_red_xls'),
    
    # establecimientos
    path('get_establecimientos_paquete_gestante/<int:establecimiento_id>/', get_establecimientos_paquete_gestante, name='get_establecimientos_paquete_gestante'),
    path('p_microredes_establec_paquete_gestante/', p_microredes_establec_paquete_gestante, name='p_microredes_establec_paquete_gestante'),
    path('p_establecimiento_paquete_gestante/', p_establecimientos_paquete_gestante, name='p_establecimientos_paquete_gestante'),       
    #-- estableccimiento excel
    path('rpt_paquete_gestante_establec_excel/', RptPaqueteGestanteEstablec.as_view(), name = 'rpt_paquete_gestante_red_xls'),
    
    
    ### COBERTURA
    path('rpt_cobertura_gestante_red_excel/', RptCoberturaPaqueteGestante.as_view(), name = 'rpt_cobertura_paquete_gestante_red_xls'),
    
]