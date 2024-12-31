from django.urls import path
from .views import index_s3_cred12, get_redes_s3_cred12, RptS3CredRed, RptCoberturaS3CredRed
from .views import get_microredes_s3_cred12, p_microredes_s3_cred12, RptS3CredMicroRed
from .views import get_establecimientos_s3_cred12, p_microredes_establec_s3_cred12, p_establecimientos_s3_cred12, RptS3CredEstablec


urlpatterns = [

    path('s3_cred12/', index_s3_cred12, name='index_s3_cred12'),
    
    ### SEGUIMIENTO
    # redes
    path('get_redes_s3_cred12/<int:redes_id>/', get_redes_s3_cred12, name='get_redes_s3_cred12'),
    #-- redes excel
    path('rpt_s3_cred12_excel/', RptS3CredRed.as_view(), name = 'rpt_s3_cred12_red_xls'),
    
    
    # microredes
    path('get_microredes_s3_cred12/<int:microredes_id>/', get_microredes_s3_cred12, name='get_microredes_s3_cred12'),
    path('p_microredes_s3_cred12/', p_microredes_s3_cred12, name='p_microredes_s3_cred12'),
    #-- microredes excel
    path('rpt_s3_cred12_microred_excel/', RptS3CredMicroRed.as_view(), name = 'rpt_s3_cred12_red_xls'),
    
    # establecimientos
    path('get_establecimientos_s3_cred12/<int:establecimiento_id>/', get_establecimientos_s3_cred12, name='get_establecimientos_s3_cred12'),
    path('p_microredes_establec_s3_cred12/', p_microredes_establec_s3_cred12, name='p_microredes_establec_s3_cred12'),
    path('p_establecimiento_s3_cred12/', p_establecimientos_s3_cred12, name='p_establecimientos_s3_cred12'),       
    #-- estableccimiento excel
    path('rpt_s3_cred12_establec_excel/', RptS3CredEstablec.as_view(), name = 'rpt_s3_cred12_red_xls'),
        
    ### COBERTURA
    path('rpt_cobertura_s3_cred12_excel/', RptCoberturaS3CredRed.as_view(), name = 'rpt_cobertura_s3_cred12_xls'),
    

]