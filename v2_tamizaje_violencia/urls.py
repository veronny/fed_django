from django.urls import path
from .views import index_v2_tamizaje_violencia, get_redes_v2_tamizaje_violencia, RptV1CondicionPreviaRed


urlpatterns = [

    path('v2_tamizaje_violencia/', index_v2_tamizaje_violencia, name='index_v2_tamizaje_violencia'),
    
    ### SEGUIMIENTO
    # redes
    path('get_redes_v2_tamizaje_violencia/<int:redes_id>/', get_redes_v2_tamizaje_violencia, name='get_redes_v2_tamizaje_violencia'),
    #-- redes excel
    path('rpt_v2_tamizaje_violencia_excel/', RptV1CondicionPreviaRed.as_view(), name = 'rpt_v2_tamizaje_violencia_red_xls'),
    
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