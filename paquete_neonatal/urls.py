from django.urls import path
from .views import index_paquete_neonatal, get_redes_paquete_neonatal, RptPaqueteNeonatalRed, RptCoberturaPaqueteNeonatal
from .views import get_microredes_paquete_neonatal, p_microredes_paquete_neonatal, RptPaqueteNeonatalMicroRed
from .views import get_establecimientos_paquete_neonatal, p_microredes_establec_paquete_neonatal, p_establecimientos_paquete_neonatal, RptPaqueteNeonatalEstablec


urlpatterns = [

    path('paquete_neonatal/', index_paquete_neonatal, name='index_paquete_neonatal'),
    
    ### SEGUIMIENTO
    # redes
    path('get_redes_paquete_neonatal/<int:redes_id>/', get_redes_paquete_neonatal, name='get_redes_paquete_neonatal'),
    ### EXCEL
    path('rpt_paquete_neonatal_red_excel/', RptPaqueteNeonatalRed.as_view(), name = 'rpt_paquete_neonatal_red_xls'),
        
    # microredes
    path('get_microredes_paquete_neonatal/<int:microredes_id>/', get_microredes_paquete_neonatal, name='get_microredes_paquete_neonatal'),
    path('p_microredes_paquete_neonatal/', p_microredes_paquete_neonatal, name='p_microredes_paquete_neonatal'),
    #-- microredes excel
    path('rpt_paquete_neonatal_microred_excel/', RptPaqueteNeonatalMicroRed.as_view(), name = 'rpt_paquete_neonatal_red_xls'),
    
    # establecimientos
    path('get_establecimientos_paquete_neonatal/<int:establecimiento_id>/', get_establecimientos_paquete_neonatal, name='get_establecimientos_paquete_neonatal'),
    path('p_microredes_establec_paquete_neonatal/', p_microredes_establec_paquete_neonatal, name='p_microredes_establec_paquete_neonatal'),
    path('p_establecimiento_paquete_neonatal/', p_establecimientos_paquete_neonatal, name='p_establecimientos_paquete_neonatal'),       
    #-- estableccimiento excel
    path('rpt_paquete_neonatal_establec_excel/', RptPaqueteNeonatalEstablec.as_view(), name = 'rpt_paquete_neonatal_red_xls'),
    
        ### COBERTURA
    path('rpt_cobertura_neonatal_red_excel/', RptCoberturaPaqueteNeonatal.as_view(), name = 'rpt_cobertura_paquete_neonatal_red_xls'),
]