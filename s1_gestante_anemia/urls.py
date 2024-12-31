from django.urls import path
from .views import index_s1_gestante_anemia, get_redes_s1_gestante_anemia, RptS1GestanteAnemiaRed, RptCoberturaS1GestanteAnemia
from .views import get_microredes_s1_gestante_anemia, p_microredes_s1_gestante_anemia, RpS1GestanteAnemiaMicroRed
from .views import get_establecimientos_s1_gestante_anemia, p_microredes_establec_s1_gestante_anemia, p_establecimientos_s1_gestante_anemia, RptS1GestanteAnemiaEstablec


urlpatterns = [

    path('s1_gestante_anemia/', index_s1_gestante_anemia, name='index_s1_gestante_anemia'),
    
    ### SEGUIMIENTO
    # redes
    path('get_redes_s1_gestante_anemia/<int:redes_id>/', get_redes_s1_gestante_anemia, name='get_redes_s1_gestante_anemia'),
    #-- redes excel
    path('rpt_s1_gestante_anemia_red_excel/', RptS1GestanteAnemiaRed.as_view(), name = 'rpt_s1_gestante_anemia_red_xls'),
    
    
    # microredes
    path('get_microredes_s1_gestante_anemia/<int:microredes_id>/', get_microredes_s1_gestante_anemia, name='get_microredes_s1_gestante_anemia'),
    path('p_microredes_s1_gestante_anemia/', p_microredes_s1_gestante_anemia, name='p_microredes_s1_gestante_anemia'),
    #-- microredes excel
    path('rpt_s1_gestante_anemia_microred_excel/', RpS1GestanteAnemiaMicroRed.as_view(), name = 'rpt_s1_gestante_anemia_red_xls'),
    
    # establecimientos
    path('get_establecimientos_s1_gestante_anemia/<int:establecimiento_id>/', get_establecimientos_s1_gestante_anemia, name='get_establecimientos_s1_gestante_anemia'),
    path('p_microredes_establec_s1_gestante_anemia/', p_microredes_establec_s1_gestante_anemia, name='p_microredes_establec_s1_gestante_anemia'),
    path('p_establecimiento_s1_gestante_anemia/', p_establecimientos_s1_gestante_anemia, name='p_establecimientos_s1_gestante_anemia'),       
    #-- estableccimiento excel
    path('rpt_s1_gestante_anemia_establec_excel/', RptS1GestanteAnemiaEstablec.as_view(), name = 'rpt_s1_gestante_anemia_red_xls'),
    
    
    ### COBERTURA
    path('rpt_cobertura_s1_gestante_anemia_excel/', RptCoberturaS1GestanteAnemia.as_view(), name = 'rpt_cobertura_s1_gestante_anemia_xls'),
    

]