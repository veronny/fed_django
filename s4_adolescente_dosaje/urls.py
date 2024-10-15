from django.urls import path
from .views import index_s4_adolescente_dosaje, get_redes_s4_adolescente_dosaje, RptS4AdolescenteDosajeRed


urlpatterns = [

    path('s4_adolescente_dosaje/', index_s4_adolescente_dosaje, name='index_s4_adolescente_dosaje'),
    
    ### SEGUIMIENTO
    # redes
    path('get_redes_s4_adolescente_dosaje/<int:redes_id>/', get_redes_s4_adolescente_dosaje, name='get_redes_s4_adolescente_dosaje'),
    #-- redes excel
    path('rpt_s4_adolescente_dosaje_excel/', RptS4AdolescenteDosajeRed.as_view(), name = 'rpt_s4_adolescente_dosaje_red_xls'),
    
    #microredes
    # path('get_microredes/<int:microredes_id>/', views.get_microredes, name='get_microredes'),
    # path('p_microredes/', views.p_microredes, name='p_microredes'),
    # #-- microredes excel
    # path('rpt_operacional_microred_excel/', RptOperacinalMicroRed.as_view(), name = 'rpt_operacional_microred_xls'),
    # 
    #establecimientos
    # path('get_establecimientos/<int:establecimiento_id>/', views.get_establecimientos, name='get_establecimientos'),
    # path('p_microredes_establec/', views.p_microredes_establec, name='p_microredes_establec'),
    # path('p_establecimiento/', views.p_establecimientos, name='p_establecimientos'),    
]