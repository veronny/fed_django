from django.shortcuts import render

# TABLERO PAQUETE NEONATAL 
from django.db import connection
from django.http import JsonResponse
from base.models import MAESTRO_HIS_ESTABLECIMIENTO, DimPeriodo
from django.db.models.functions import Substr
import logging

# report excel
from django.http.response import HttpResponse
from django.views.generic.base import TemplateView
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import openpyxl
from openpyxl.utils import get_column_letter

from django.db.models.functions import Substr

from datetime import datetime
import locale

from django.db.models import IntegerField  # Importar IntegerField
from django.db.models.functions import Cast, Substr  # Importar Cast y Substr


logger = logging.getLogger(__name__)
# Create your views here.
def obtener_distritos(provincia):
    distritos = MAESTRO_HIS_ESTABLECIMIENTO.objects.filter(Provincia=provincia).values('Distrito').distinct().order_by('Distrito')
    return list(distritos)

def obtener_avance_paquete_nino(red):
    with connection.cursor() as cursor:
        cursor.execute(
            "SELECT * FROM public.obtener_avance_paquete_nino(%s)",
            [red]
        )
        return cursor.fetchall()

def obtener_ranking_paquete_nino(anio, mes):
    with connection.cursor() as cursor:
        cursor.execute(
            "SELECT * FROM public.obtener_ranking_paquete_nino(%s, %s)",
            [anio, mes]
        )
        result = cursor.fetchall()
        return result
    
def index_paquete_nino(request):
    # RANKING 
    anio = request.GET.get('anio')  # Valor predeterminado# Valor predeterminado
    mes_seleccionado = request.GET.get('mes')
    # GRAFICO
    red_seleccionada = request.GET.get('red')
    red = MAESTRO_HIS_ESTABLECIMIENTO.objects.filter(Disa='JUNIN').values_list('Red', flat=True).distinct().order_by('Red')
    # Si la solicitud es AJAX
    if request.headers.get('x-requested-with') == 'XMLHttpRequest':
        try:
            # Obtener datos de RANKING 
            resultados_ranking_paquete_nino = obtener_ranking_paquete_nino(anio,mes_seleccionado)
            # Obtener datos de AVANCE GRAFICO MESES
            resultados_avance_paquete_nino = obtener_avance_paquete_nino(red_seleccionada)
            
            # Procesar los resultados
            if any(len(row) < 4 for row in resultados_ranking_paquete_nino):
                raise ValueError("Algunas filas del ranking no tienen suficientes elementos")
            
            data = {               
                #ranking
                'red': [],
                'num_r': [],
                'den_r': [],
                'avance_r': [],
                
                #avance meses
                'mes': [],
                'num': [],
                'den': [],
                'avance': [],
            }
            #RANKING
            for index, row in enumerate(resultados_ranking_paquete_nino):
                try:
                    # Verifica que la tupla tenga exactamente 4 elementos
                    if len(row) != 4:
                        raise ValueError(f"La fila {index} no tiene 4 elementos: {row}")

                    red_value = row[0] if row[0] is not None else ''
                    num_r_value = float(row[1]) if row[1] is not None else 0.0
                    den_r_value = float(row[2]) if row[2] is not None else 0.0
                    avance_r_value = float(row[3]) if row[3] is not None else 0.0

                    data['red'].append(red_value)
                    data['num_r'].append(num_r_value)
                    data['den_r'].append(den_r_value)
                    data['avance_r'].append(avance_r_value)

                except Exception as e:
                    logger.error(f"Error procesando la fila {index}: {str(e)}")
            
            #AVANCE GRAFICO MESES
            for index, row in enumerate(resultados_avance_paquete_nino):
                try:
                    # Verifica que la tupla tenga exactamente 4 elementos
                    if len(row) != 5:
                        raise ValueError(f"La fila {index} no tiene 5 elementos: {row}")

                    mes_value = row[1] if row[1] is not None else ''
                    num_value = float(row[2]) if row[2] is not None else 0.0
                    den_value = float(row[3]) if row[3] is not None else 0.0
                    avance_value = float(row[4]) if row[4] is not None else 0.0

                    data['mes'].append(mes_value)
                    data['num'].append(num_value)
                    data['den'].append(den_value)
                    data['avance'].append(avance_value)

                except Exception as e:
                    logger.error(f"Error procesando la fila {index}: {str(e)}")
            
            return JsonResponse(data)

        except Exception as e:
            logger.error(f"Error al obtener datos: {str(e)}")

    # Si no es una solicitud AJAX, renderiza la página principal
    return render(request, 'paquete_nino/index_paquete_nino.html', {
        'red': red,
        'mes_seleccionado': mes_seleccionado,
    })

## SEGUIMIENTO
def get_redes_paquete_nino(request,redes_id):
    redes = (
            MAESTRO_HIS_ESTABLECIMIENTO
            .objects.filter(Descripcion_Sector='GOBIERNO REGIONAL',Departamento='JUNIN')
            .annotate(codigo_red_filtrado=Substr('Codigo_Red', 1, 4))
            .values('Red','codigo_red_filtrado')
            .distinct()
            .order_by('Red')
    )
    mes_inicio = (
                DimPeriodo
                .objects.filter(Anio='2024')
                .annotate(nro_mes=Cast('NroMes', IntegerField())) 
                .values('Mes','nro_mes')
                .order_by('NroMes')
                .distinct()
    ) 
    mes_fin = (
                DimPeriodo
                .objects.filter(Anio='2024')
                .annotate(nro_mes=Cast('NroMes', IntegerField())) 
                .values('Mes','nro_mes')
                .order_by('NroMes')
                .distinct()
    ) 
    context = {
                'redes': redes,
                'mes_inicio':mes_inicio,
                'mes_fin':mes_fin,
    }
    
    return render(request, 'paquete_nino/redes.html', context)

def obtener_seguimiento_redes_paquete_nino(p_red,p_inicio,p_fin):
    with connection.cursor() as cursor:
        cursor.execute(
            "SELECT DISTINCT * FROM public.fn_seguimiento_paquete_nino(%s, %s, %s)",
            [p_red, p_inicio, p_fin]
        )
        return cursor.fetchall()

class RptPaqueteNinoRed(TemplateView):
    def get(self, request, *args, **kwargs):
        # Variables ingresadas
        p_red = request.GET.get('red')
        p_inicio = request.GET.get('fecha_inicio')
        p_fin = request.GET.get('fecha_fin')
        # Creación de la consulta
        resultado_seguimiento = obtener_seguimiento_redes_paquete_nino(p_red, p_inicio, p_fin)
        
        wb = Workbook()
        
        consultas = [
                ('Seguimiento', resultado_seguimiento)
        ]
        
        for index, (sheet_name, results) in enumerate(consultas):
            if index == 0:
                ws = wb.active
                ws.title = sheet_name
            else:
                ws = wb.create_sheet(title=sheet_name)
        
            fill_worksheet(ws, results)
        ##########################################################################          
        # Establecer el nombre del archivo
        nombre_archivo = "rpt_paquete_nino_red.xlsx"
        # Definir el tipo de respuesta que se va a dar
        response = HttpResponse(content_type="application/ms-excel")
        contenido = "attachment; filename={}".format(nombre_archivo)
        response["Content-Disposition"] = contenido
        wb.save(response)

        return response

def fill_worksheet(ws, results): 
    # cambia el alto de la columna
    ws.row_dimensions[1].height = 14
    ws.row_dimensions[2].height = 14
    ws.row_dimensions[3].height = 3
    ws.row_dimensions[4].height = 25
    ws.row_dimensions[5].height = 3
    ws.row_dimensions[7].height = 3
    ws.row_dimensions[8].height = 25
    # cambia el ancho de la columna
    ws.column_dimensions['A'].width = 2
    ws.column_dimensions['B'].width = 5
    ws.column_dimensions['C'].width = 9
    ws.column_dimensions['D'].width = 9
    ws.column_dimensions['E'].width = 5
    ws.column_dimensions['F'].width = 9
    ws.column_dimensions['G'].width = 5
    ws.column_dimensions['H'].width = 5
    ws.column_dimensions['I'].width = 5
    ws.column_dimensions['J'].width = 6
    ws.column_dimensions['K'].width = 5
    ws.column_dimensions['L'].width = 5
    ws.column_dimensions['M'].width = 5
    ws.column_dimensions['N'].width = 5
    ws.column_dimensions['O'].width = 9
    ws.column_dimensions['P'].width = 3
    ws.column_dimensions['Q'].width = 9
    ws.column_dimensions['R'].width = 9
    ws.column_dimensions['S'].width = 9
    ws.column_dimensions['T'].width = 3
    ws.column_dimensions['U'].width = 9
    ws.column_dimensions['V'].width = 3
    ws.column_dimensions['W'].width = 9
    ws.column_dimensions['X'].width = 3
    ws.column_dimensions['Y'].width = 9
    ws.column_dimensions['Z'].width = 3
    ws.column_dimensions['AA'].width = 9
    ws.column_dimensions['AB'].width = 9
    ws.column_dimensions['AC'].width = 3
    ws.column_dimensions['AD'].width = 9
    ws.column_dimensions['AE'].width = 3
    ws.column_dimensions['AF'].width = 9
    ws.column_dimensions['AG'].width = 3
    ws.column_dimensions['AH'].width = 9
    ws.column_dimensions['AI'].width = 3
    ws.column_dimensions['AJ'].width = 9
    ws.column_dimensions['AK'].width = 3
    ws.column_dimensions['AL'].width = 9
    ws.column_dimensions['AM'].width = 3
    ws.column_dimensions['AN'].width = 9
    ws.column_dimensions['AO'].width = 3
    ws.column_dimensions['AP'].width = 9
    ws.column_dimensions['AQ'].width = 3
    ws.column_dimensions['AR'].width = 9
    ws.column_dimensions['AS'].width = 3
    ws.column_dimensions['AT'].width = 9
    ws.column_dimensions['AU'].width = 3
    ws.column_dimensions['AV'].width = 9
    ws.column_dimensions['AW'].width = 3
    ws.column_dimensions['AX'].width = 9
    ws.column_dimensions['AY'].width = 9
    ws.column_dimensions['AZ'].width = 9
    ws.column_dimensions['BA'].width = 3
    ws.column_dimensions['BB'].width = 9
    ws.column_dimensions['BC'].width = 3
    ws.column_dimensions['BD'].width = 9
    ws.column_dimensions['BE'].width = 9
    ws.column_dimensions['BF'].width = 3
    ws.column_dimensions['BG'].width = 9
    ws.column_dimensions['BH'].width = 3
    ws.column_dimensions['BI'].width = 9
    ws.column_dimensions['BJ'].width = 3
    ws.column_dimensions['BK'].width = 9
    ws.column_dimensions['BL'].width = 9
    ws.column_dimensions['BM'].width = 3
    ws.column_dimensions['BN'].width = 9
    ws.column_dimensions['BO'].width = 3
    ws.column_dimensions['BP'].width = 9
    ws.column_dimensions['BQ'].width = 3
    ws.column_dimensions['BR'].width = 9
    ws.column_dimensions['BS'].width = 9
    ws.column_dimensions['BT'].width = 3
    ws.column_dimensions['BU'].width = 9
    ws.column_dimensions['BV'].width = 3
    ws.column_dimensions['BW'].width = 9
    ws.column_dimensions['BX'].width = 9
    ws.column_dimensions['BY'].width = 9
    ws.column_dimensions['BZ'].width = 3
    ws.column_dimensions['CA'].width = 9
    ws.column_dimensions['CB'].width = 9
    ws.column_dimensions['CC'].width = 9
    ws.column_dimensions['CD'].width = 3
    ws.column_dimensions['CE'].width = 9
    ws.column_dimensions['CF'].width = 3
    ws.column_dimensions['CG'].width = 9
    ws.column_dimensions['CH'].width = 9
    ws.column_dimensions['CI'].width = 3
    ws.column_dimensions['CJ'].width = 9
    ws.column_dimensions['CK'].width = 3
    ws.column_dimensions['CL'].width = 9
    ws.column_dimensions['CM'].width = 3
    ws.column_dimensions['CN'].width = 9
    ws.column_dimensions['CO'].width = 9
    ws.column_dimensions['CP'].width = 3
    ws.column_dimensions['CQ'].width = 9
    ws.column_dimensions['CR'].width = 3
    ws.column_dimensions['CS'].width = 9
    ws.column_dimensions['CT'].width = 3
    ws.column_dimensions['CU'].width = 9
    ws.column_dimensions['CV'].width = 3
    ws.column_dimensions['CW'].width = 9
    ws.column_dimensions['CX'].width = 3
    ws.column_dimensions['CY'].width = 9
    ws.column_dimensions['CZ'].width = 3
    ws.column_dimensions['DA'].width = 9
    ws.column_dimensions['DB'].width = 9
    ws.column_dimensions['DC'].width = 3
    ws.column_dimensions['DD'].width = 9
    ws.column_dimensions['DE'].width = 10
    ws.column_dimensions['DF'].width = 6
    ws.column_dimensions['DG'].width = 6
    ws.column_dimensions['DH'].width = 10
    ws.column_dimensions['DI'].width = 10
    ws.column_dimensions['DJ'].width = 16
    ws.column_dimensions['DK'].width = 20
    ws.column_dimensions['DL'].width = 20
    ws.column_dimensions['DM'].width = 6
    ws.column_dimensions['DN'].width = 25
    ws.column_dimensions['DO'].width = 9
    ws.column_dimensions['DP'].width = 33
    
    # linea de division
    ws.freeze_panes = 'Q9'
    # Configuración del fondo y el borde
    # Definir el color usando formato aRGB (opacidad completa 'FF' + color RGB)
    fill = PatternFill(start_color='FF60D7E0', end_color='FF60D7E0', fill_type='solid')
    # Definir el color anaranjado usando formato aRGB
    orange_fill = PatternFill(start_color='FFE0A960', end_color='FFE0A960', fill_type='solid')
    # Definir los estilos para gris
    gray_fill = PatternFill(start_color='FFD3D3D3', end_color='FFD3D3D3', fill_type='solid')
    # Definir el estilo de color verde
    green_fill = PatternFill(start_color='FF60E0B3', end_color='FF60E0B3', fill_type='solid')
    # Definir el estilo de color amarillo
    yellow_fill = PatternFill(start_color='FFE0DE60', end_color='FFE0DE60', fill_type='solid')
    # Definir el estilo de color azul
    blue_fill = PatternFill(start_color='FF60A2E0', end_color='FF60A2E0', fill_type='solid')
    # Definir el estilo de color verde 2
    green_fill_2 = PatternFill(start_color='FF60E07E', end_color='FF60E07E', fill_type='solid')   
    
    green_font = Font(name='Arial', size=8, color='00FF00')  # Verde
    red_font = Font(name='Arial', size=8, color='FF0000')    # Rojo
    
    border = Border(left=Side(style='thin', color='00B0F0'),
                    right=Side(style='thin', color='00B0F0'),
                    top=Side(style='thin', color='00B0F0'),
                    bottom=Side(style='thin', color='00B0F0'))
    borde_plomo = Border(left=Side(style='thin', color='A9A9A9'), # Plomo
                    right=Side(style='thin', color='A9A9A9'), # Plomo
                    top=Side(style='thin', color='A9A9A9'), # Plomo
                    bottom=Side(style='thin', color='A9A9A9')) # Plomo
    
    ## crea titulo del reporte
    ws['B1'].alignment = Alignment(horizontal= "left", vertical="center")
    ws['B1'].font = Font(name = 'Arial', size= 7, bold = True)
    ws['B1'] = 'OFICINA DE TECNOLOGIAS DE LA INFORMACION'
    
    ws['B2'].alignment = Alignment(horizontal= "left", vertical="center")
    ws['B2'].font = Font(name = 'Arial', size= 7, bold = True)
    ws['B2'] = 'DIRECCION REGIONAL DE SALUD JUNIN'
    
    ws['B4'].alignment = Alignment(horizontal= "left", vertical="center")
    ws['B4'].font = Font(name = 'Arial', size= 12, bold = True)
    ws['B4'] = 'SEGUIMIENTO NOMINAL DEL INDICADOR MC-02. NIÑAS Y NIÑOS MENORES DE 12 MESES DE EDAD PROCEDENTES DE LOS QUINTILES 1 Y 2 DE POBREZA DEPARTAMENTAL QUE RECIBEN EL PAQUETE INTEGRADO DE SERVICIOS'
    
    ws['B6'].alignment = Alignment(horizontal= "left", vertical="center")
    ws['B6'].font = Font(name = 'Arial', size= 7, bold = True, color='0000CC')
    ws['B6'] ='El usuario se compromete a mantener la confidencialidad de los datos personales que conozca como resultado del reporte realizado, cumpliendo con lo establecido en la Ley N° 29733 - Ley de Protección de Datos Personales y sus normas complementarias.'
        
    ws['B8'].alignment = Alignment(horizontal= "center", vertical="center")
    ws['B8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['B8'].fill = fill
    ws['B8'].border = border
    ws['B8'] = 'TD'
    
    ws['C8'].alignment = Alignment(horizontal= "center", vertical="center")
    ws['C8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['C8'].fill = fill
    ws['C8'].border = border
    ws['C8'] = 'NUM DOC'
    
    ws['D8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['D8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['D8'].fill = fill
    ws['D8'].border = border
    ws['D8'] = 'FECHA NAC'      
    
    ws['E8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['E8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['E8'].fill = fill
    ws['E8'].border = border
    ws['E8'] = 'SEXO' 
    
    ws['F8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['F8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['F8'].fill = fill
    ws['F8'].border = border
    ws['F8'] = 'SEGURO'     
    
    ws['G8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['G8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['G8'].fill = fill
    ws['G8'].border = border
    ws['G8'] = 'ED DIAS'    
    
    ws['H8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['H8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['H8'].fill = fill
    ws['H8'].border = border
    ws['H8'] = 'ED MES'    
    
    ws['I8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['I8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['I8'].fill = fill
    ws['I8'].border = border
    ws['I8'] = 'CNV'    
    
    ws['J8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['J8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['J8'].fill = fill
    ws['J8'].border = border
    ws['J8'] = 'PESO'  
    
    ws['K8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['K8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['K8'].fill = fill
    ws['K8'].border = border
    ws['K8'] = 'BPN'  
    
    ws['L8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['L8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['L8'].fill = fill
    ws['L8'].border = border
    ws['L8'] = 'SEM GEST'  
    
    ws['M8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['M8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['M8'].fill = fill
    ws['M8'].border = border
    ws['M8'] = 'PREM'  
    
    ws['N8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['N8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['N8'].fill = fill
    ws['N8'].border = border
    ws['N8'] = 'BPN PREM'  
    
    ws['O8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['O8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['O8'].fill = fill
    ws['O8'].border = border
    ws['O8'] = 'DENOMINADOR'  
    
    ws['P8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['P8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['P8'].fill = fill
    ws['P8'].border = border
    ws['P8'] = 'SIN DNI'  
    
    ws['Q8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['Q8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['Q8'].fill = green_fill_2
    ws['Q8'].border = border
    ws['Q8'] = 'CRED'    
    
    ws['R8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['R8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['R8'].fill = green_fill_2
    ws['R8'].border = border
    ws['R8'] = 'CRED RN' 
    
    ws['S8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['S8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['S8'].fill = green_fill_2
    ws['S8'].border = border
    ws['S8'] = '1° CRED RN' 
    
    ws['T8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['T8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['T8'].fill = green_fill_2
    ws['T8'].border = border
    ws['T8'] = 'V' 
    
    ws['U8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['U8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['U8'].fill = green_fill_2
    ws['U8'].border = border
    ws['U8'] = '2° CRED RN' 
    
    ws['V8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['V8'].font = Font(name = 'Arial', size= 7, bold = True, color='000000')
    ws['V8'].fill = green_fill_2
    ws['V8'].border = border
    ws['V8'] = 'V' 
    
    ws['W8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['W8'].font = Font(name = 'Arial', size= 7, bold = True, color='000000')
    ws['W8'].fill = green_fill_2
    ws['W8'].border = border
    ws['W8'] = '3° CRED RN'   
    
    ws['X8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['X8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['X8'].fill = green_fill_2
    ws['X8'].border = border
    ws['X8'] = 'V' 
    
    ws['Y8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['Y8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['Y8'].fill = green_fill_2
    ws['Y8'].border = border
    ws['Y8'] = '4° CRED RN' 
    
    ws['Z8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['Z8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['Z8'].fill = green_fill_2
    ws['Z8'].border = border
    ws['Z8'] = 'V' 
    
    ws['AA8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AA8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AA8'].fill = green_fill
    ws['AA8'].border = border
    ws['AA8'] = 'CRED MES' 
    
    ws['AB8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AB8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AB8'].fill = green_fill
    ws['AB8'].border = border
    ws['AB8'] = '1° CRED'     
    
    ws['AC8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AC8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AC8'].fill = green_fill
    ws['AC8'].border = border
    ws['AC8'] = 'V' 
    
    ws['AD8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AD8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AD8'].fill = green_fill
    ws['AD8'].border = border
    ws['AD8'] = '2° CRED' 
    
    ws['AE8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AE8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AE8'].fill = green_fill
    ws['AE8'].border = border
    ws['AE8'] = 'V' 
    
    ws['AF8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AF8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AF8'].fill = green_fill
    ws['AF8'].border = border
    ws['AF8'] = '3° CRED' 
    
    ws['AG8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AG8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AG8'].fill = green_fill
    ws['AG8'].border = border
    ws['AG8'] = 'V' 
    
    ws['AH8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AH8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AH8'].fill = green_fill
    ws['AH8'].border = border
    ws['AH8'] = '4° CRED' 
    
    ws['AI8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AI8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AI8'].fill = green_fill
    ws['AI8'].border = border
    ws['AI8'] = 'V' 
    
    ws['AJ8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AJ8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AJ8'].fill = green_fill
    ws['AJ8'].border = border
    ws['AJ8'] = '5° CRED' 
    
    ws['AK8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AK8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AK8'].fill = green_fill
    ws['AK8'].border = border
    ws['AK8'] = 'V' 
    
    ws['AL8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AL8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AL8'].fill = green_fill
    ws['AL8'].border = border
    ws['AL8'] = '6° CRED' 
    
    ws['AM8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AM8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AM8'].fill = green_fill
    ws['AM8'].border = border
    ws['AM8'] = 'V' 
    
    ws['AN8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AN8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AN8'].fill = green_fill
    ws['AN8'].border = border
    ws['AN8'] = '7° CRED' 
    
    ws['AO8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AO8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AO8'].fill = green_fill
    ws['AO8'].border = border
    ws['AO8'] = 'V' 
    
    ws['AP8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AP8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AP8'].fill = green_fill
    ws['AP8'].border = border
    ws['AP8'] = '8° CRED' 
    
    ws['AQ8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AQ8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AQ8'].fill = green_fill
    ws['AQ8'].border = border
    ws['AQ8'] = 'V' 
    
    ws['AR8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AR8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AR8'].fill = green_fill
    ws['AR8'].border = border
    ws['AR8'] = '9° CRED' 
    
    ws['AS8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AS8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AS8'].fill = green_fill
    ws['AS8'].border = border
    ws['AS8'] = 'V' 
    
    ws['AT8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AT8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AT8'].fill = green_fill
    ws['AT8'].border = border
    ws['AT8'] = '10° CRED' 
    
    ws['AU8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AU8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AU8'].fill = green_fill
    ws['AU8'].border = border
    ws['AU8'] = 'V' 
    
    ws['AV8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AV8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AV8'].fill = green_fill
    ws['AV8'].border = border
    ws['AV8'] = '11° CRED' 
    
    ws['AW8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AW8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AW8'].fill = green_fill
    ws['AW8'].border = border
    ws['AW8'] = 'V' 
    
    ws['AX8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AX8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AX8'].fill = yellow_fill
    ws['AX8'].border = border
    ws['AX8'] = 'NUM VAC' 
    
    ws['AY8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AY8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AY8'].fill = yellow_fill
    ws['AY8'].border = border
    ws['AY8'] = 'NUM NEUMO' 
    
    ws['AZ8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AZ8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AZ8'].fill = yellow_fill
    ws['AZ8'].border = border
    ws['AZ8'] = '1° NEUMO' 
    
    ws['BA8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BA8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BA8'].fill = yellow_fill
    ws['BA8'].border = border
    ws['BA8'] = 'V' 
    
    ws['BB8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BB8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BB8'].fill = yellow_fill
    ws['BB8'].border = border
    ws['BB8'] = '2° NEUMO' 
    
    ws['BC8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BC8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BC8'].fill = yellow_fill
    ws['BC8'].border = border
    ws['BC8'] = 'V'     
    
    ws['BD8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BD8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BD8'].fill = yellow_fill
    ws['BD8'].border = border
    ws['BD8'] = 'NUM POLIO' 
    
    ws['BE8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BE8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BE8'].fill = yellow_fill
    ws['BE8'].border = border
    ws['BE8'] = '1° POLIO' 
    
    ws['BF8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BF8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BF8'].fill = yellow_fill
    ws['BF8'].border = border
    ws['BF8'] = 'V' 
    
    ws['BG8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BG8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BG8'].fill = yellow_fill
    ws['BG8'].border = border
    ws['BG8'] = '2° POLIO' 
    
    ws['BH8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BH8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BH8'].fill = yellow_fill
    ws['BH8'].border = border
    ws['BH8'] = 'V' 
    
    ws['BI8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BI8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BI8'].fill = yellow_fill
    ws['BI8'].border = border
    ws['BI8'] = '3° POLIO' 
    
    ws['BJ8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BJ8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BJ8'].fill = yellow_fill
    ws['BJ8'].border = border
    ws['BJ8'] = 'V' 
    
    ws['BK8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BK8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BK8'].fill = yellow_fill
    ws['BK8'].border = border
    ws['BK8'] = 'NUM PENTA' 
    
    ws['BL8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BL8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BL8'].fill = yellow_fill
    ws['BL8'].border = border
    ws['BL8'] = '1° PENTA' 
    
    ws['BM8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BM8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BM8'].fill = yellow_fill
    ws['BM8'].border = border
    ws['BM8'] = 'V' 
    
    ws['BN8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BN8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BN8'].fill = yellow_fill
    ws['BN8'].border = border
    ws['BN8'] = '2° PENTA' 
    
    ws['BO8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BO8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BO8'].fill = yellow_fill
    ws['BO8'].border = border
    ws['BO8'] = 'V' 
    
    ws['BP8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BP8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BP8'].fill = yellow_fill
    ws['BP8'].border = border
    ws['BP8'] = '3° PENTA' 
    
    ws['BQ8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BQ8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BQ8'].fill = yellow_fill
    ws['BQ8'].border = border
    ws['BQ8'] = 'V'  
    
    ws['BR8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BR8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BR8'].fill = yellow_fill
    ws['BR8'].border = border
    ws['BR8'] = 'NUM ROTA' 
    
    ws['BS8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BS8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BS8'].fill = yellow_fill
    ws['BS8'].border = border
    ws['BS8'] = '1° ROTA' 
    
    ws['BT8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BT8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BT8'].fill = yellow_fill
    ws['BT8'].border = border
    ws['BT8'] = 'V' 
    
    ws['BU8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BU8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BU8'].fill = yellow_fill
    ws['BU8'].border = border
    ws['BU8'] = '2° ROTA' 
    
    ws['BV8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BV8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BV8'].fill = yellow_fill
    ws['BV8'].border = border
    ws['BV8'] = 'V' 
    
    ws['BW8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BW8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BW8'].fill = blue_fill
    ws['BW8'].border = border
    ws['BW8'] = 'NUM ESQ'
    
    ws['BX8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BX8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BX8'].fill = blue_fill
    ws['BX8'].border = border
    ws['BX8'] = 'ESQ 4M' 
    
    ws['BY8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BY8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['BY8'].fill = blue_fill
    ws['BY8'].border = border
    ws['BY8'] = 'SUP 4M' 
    
    ws['BZ8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BZ8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['BZ8'].fill = blue_fill
    ws['BZ8'].border = border
    ws['BZ8'] = 'V' 
    
    ws['CA8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['CA8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['CA8'].fill = blue_fill
    ws['CA8'].border = border
    ws['CA8'] = 'ESQ 6M' 
    
    ws['CB8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['CB8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['CB8'].fill = blue_fill
    ws['CB8'].border = border
    ws['CB8'] = 'NUM SUP 6M' 
    
    ws['CC8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['CC8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['CC8'].fill = blue_fill
    ws['CC8'].border = border
    ws['CC8'] = '1° SUP 6M' 
    
    ws['CD8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['CD8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['CD8'].fill = blue_fill
    ws['CD8'].border = border
    ws['CD8'] = 'V' 
    
    ws['CE8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['CE8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['CE8'].fill = blue_fill
    ws['CE8'].border = border
    ws['CE8'] = '2° SUP 6M' 
    
    ws['CF8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['CF8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['CF8'].fill = blue_fill
    ws['CF8'].border = border
    ws['CF8'] = 'V' 
    
    ws['CG8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['CG8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['CG8'].fill = blue_fill
    ws['CG8'].border = border
    ws['CG8'] = 'NUM TOO 6M' 
    
    ws['CH8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['CH8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['CH8'].fill = blue_fill
    ws['CH8'].border = border
    ws['CH8'] = '1° TTO 6M' 
    
    ws['CI8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['CI8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['CI8'].fill = blue_fill
    ws['CI8'].border = border
    ws['CI8'] = 'V' 
    
    ws['CJ8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['CJ8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['CJ8'].fill = blue_fill
    ws['CJ8'].border = border
    ws['CJ8'] = '2° TTO 6M' 
    
    ws['CK8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['CK8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['CK8'].fill = blue_fill
    ws['CK8'].border = border
    ws['CK8'] = 'V' 
    
    ws['CL8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['CL8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['CL8'].fill = blue_fill
    ws['CL8'].border = border
    ws['CL8'] = '3° TTO 6M' 
    
    ws['CM8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['CM8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['CM8'].fill = blue_fill
    ws['CM8'].border = border
    ws['CM8'] = 'V' 
    
    ws['CN8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['CN8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['CN8'].fill = blue_fill
    ws['CN8'].border = border
    ws['CN8'] = 'NUM MULT 6M' 
    
    ws['CO8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['CO8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['CO8'].fill = blue_fill
    ws['CO8'].border = border
    ws['CO8'] = '1° MULTI 6M' 
    
    ws['CP8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['CP8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['CP8'].fill = blue_fill
    ws['CP8'].border = border
    ws['CP8'] = 'V' 
    
    ws['CQ8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['CQ8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['CQ8'].fill = blue_fill
    ws['CQ8'].border = border
    ws['CQ8'] = '2° MULTI 6M' 
    
    ws['CR8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['CR8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['CR8'].fill = blue_fill
    ws['CR8'].border = border
    ws['CR8'] = 'V' 
    
    ws['CS8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['CS8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['CS8'].fill = blue_fill
    ws['CS8'].border = border
    ws['CS8'] = '3° MULTI 6M' 
    
    ws['CT8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['CT8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['CT8'].fill = blue_fill
    ws['CT8'].border = border
    ws['CT8'] = 'V' 
    
    ws['CU8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['CU8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['CU8'].fill = blue_fill
    ws['CU8'].border = border
    ws['CU8'] = '4° MULTI 6M' 
    
    ws['CV8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['CV8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['CV8'].fill = blue_fill
    ws['CV8'].border = border
    ws['CV8'] = 'V' 
    
    ws['CW8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['CW8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['CW8'].fill = blue_fill
    ws['CW8'].border = border
    ws['CW8'] = '5° MULTI 6M' 
    
    ws['CX8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['CX8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['CX8'].fill = blue_fill
    ws['CX8'].border = border
    ws['CX8'] = 'V' 
    
    ws['CY8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['CY8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['CY8'].fill = blue_fill
    ws['CY8'].border = border
    ws['CY8'] = '6° MULTI 6M' 
    
    ws['CZ8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['CZ8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['CZ8'].fill = blue_fill
    ws['CZ8'].border = border
    ws['CZ8'] = 'V' 
    
    ws['DA8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['DA8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['DA8'].fill = blue_fill
    ws['DA8'].border = border
    ws['DA8'] = 'NUM DOSAJE HB' 
    
    ws['DB8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['DB8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['DB8'].fill = blue_fill
    ws['DB8'].border = border
    ws['DB8'] = 'DOSAJE HB' 
    
    ws['DC8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['DC8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['DC8'].fill = blue_fill
    ws['DC8'].border = border
    ws['DC8'] = 'NUM HB' 
    
    ws['DD8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['DD8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['DD8'].fill = gray_fill
    ws['DD8'].border = border
    ws['DD8'] = 'NUM DNI EMISION' 
    
    ws['DE8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['DE8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['DE8'].fill = gray_fill
    ws['DE8'].border = border
    ws['DE8'] = 'EMISION' 
    
    ws['DF8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['DF8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['DF8'].fill = gray_fill
    ws['DF8'].border = border
    ws['DF8'] = 'DNI 30D' 
    
    ws['DG8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['DG8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['DG8'].fill = gray_fill
    ws['DG8'].border = border
    ws['DG8'] = 'DNI 60D' 
        
    ws['DH8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['DH8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['DH8'].fill = fill
    ws['DH8'].border = border
    ws['DH8'] = 'MES' 
    
    ws['DI8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['DI8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['DI8'].fill = gray_fill
    ws['DI8'].border = border
    ws['DI8'] = 'IND' 
    
    ws['DJ8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['DJ8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['DJ8'].fill = orange_fill
    ws['DJ8'].border = border
    ws['DJ8'] = 'UBIGEO'  
    
    ws['DK8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['DK8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['DK8'].fill = orange_fill
    ws['DK8'].border = border
    ws['DK8'] = 'PROVINCIA'       
    
    ws['DL8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['DL8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['DL8'].fill = orange_fill
    ws['DL8'].border = border
    ws['DL8'] = 'DISTRITO' 
    
    ws['DM8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['DM8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['DM8'].fill = orange_fill
    ws['DM8'].border = border
    ws['DM8'] = 'RED'  
    
    ws['DN8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['DN8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['DN8'].fill = orange_fill
    ws['DN8'].border = border
    ws['DN8'] = 'MICRORED'  
    
    ws['DO8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['DO8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['DO8'].fill = orange_fill
    ws['DO8'].border = border
    ws['DO8'] = 'COD EST'  
    
    ws['DP8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['DP8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['DP8'].fill = orange_fill
    ws['DP8'].border = border
    ws['DP8'] = 'ESTABLECIMIENTO'  
    
    # Definir estilos
    header_font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    centered_alignment = Alignment(horizontal='center')
    border = Border(left=Side(style='thin', color='A9A9A9'),
            right=Side(style='thin', color='A9A9A9'),
            top=Side(style='thin', color='A9A9A9'),
            bottom=Side(style='thin', color='A9A9A9'))
    header_fill = PatternFill(patternType='solid', fgColor='00B0F0')
    
    
    # Definir los caracteres especiales de check y X
    check_mark = '✓'  # Unicode para check
    x_mark = '✗'  # Unicode para X
    sub_cumple = 'CUMPLE'
    sub_no_cumple = 'NO CUMPLE'
    
    # Escribir datos
    for row, record in enumerate(results, start=9):
        for col, value in enumerate(record, start=2):
            cell = ws.cell(row=row, column=col, value=value)

            # Alinear a la izquierda solo en las columnas 6,14,15,16
            if col in [116, 117, 118, 120]:
                cell.alignment = Alignment(horizontal='left')
            else:
                cell.alignment = Alignment(horizontal='center')

            # Aplicar color en la columna 27
            if col == 113:
                if isinstance(value, str):
                    value_upper = value.strip().upper()
                    if value_upper == "NO CUMPLE":
                        cell.fill = PatternFill(patternType='solid', fgColor='FF0000')  # Fondo rojo
                        cell.font = Font(name='Arial', size=7,  bold = True, color='FFFFFF')  # Letra blanca
                    elif value_upper == "CUMPLE":
                        cell.fill = PatternFill(patternType='solid', fgColor='00FF00')  # Fondo verde
                        cell.font = Font(name='Arial', size=7,  bold = True, color='FFFFFF')  # Letra blanca
                    else:
                        cell.font = Font(name='Arial', size=7)
                else:
                    cell.font = Font(name='Arial', size=8)
            
            # Aplicar color de letra SUB INDICADORES
            elif col in [15, 18, 27, 51, 56, 63, 70, 76, 79, 80, 85, 92]:
                if value == 0:
                    cell.value = sub_no_cumple  # Insertar check
                    cell.font = Font(name='Arial', size=7, color="FF0000") 
                elif value == 1:
                    cell.value = sub_cumple # Insertar check
                    cell.font = Font(name='Arial', size=7, color="00B050")
                else:
                    cell.font = Font(name='Arial', size=7)
            # Fuente normal para otras columnas
            # Aplicar color de letra SUB GENERALIDADES
            elif col in [17, 50, 75, 105, 108]:
                if value == 0:
                    cell.value = sub_no_cumple  # Insertar check
                    cell.font = Font(name='Arial', size=7, color="FF0000") 
                    cell.fill = gray_fill # Letra roja
                elif value == 1:
                    cell.value = sub_cumple # Insertar check
                    cell.font = Font(name='Arial', size=7, color="00B050")
                    cell.fill = gray_fill# Letra verde
                else:
                    cell.font = Font(name='Arial', size=7)
            # Fuente normal para otras columnas
            else:
                cell.font = Font(name='Arial', size=8)  # Fuente normal para otras columnas
            
            # Aplicar caracteres especiales check y X
            if col in [9, 11, 13, 14, 16, 20, 22, 24, 26, 29, 31, 33, 35, 37, 39, 41, 43, 45, 47, 49, 53, 55, 58, 60, 62, 65, 67, 69, 72, 74, 78, 82, 84, 87, 89, 91, 94, 96, 98, 100, 102, 104, 107, 110, 111]:
                if value == 1:
                    cell.value = check_mark  # Insertar check
                    cell.font = Font(name='Arial', size=10, color='00B050')  # Letra verde
                elif value == 0:
                    cell.value = x_mark  # Insertar X
                    cell.font = Font(name='Arial', size=10, color='FF0000')  # Letra roja
                else:
                    cell.font = Font(name='Arial', size=8)  # Fuente normal si no es 1 o 0
            
            cell.border = border

###########################################################################################
# -- COBERTURA PAQUETE NEONATAL
def obtener_cobertura_paquete_nino():
    with connection.cursor() as cursor:
        cursor.execute(
            'SELECT * FROM public."Cobertura_MC02_PaquetenNino" ORDER BY "Red", "MicroRed", "Nombre_Establecimiento";'
        )
        return cursor.fetchall()

class RptCoberturaPaqueteNino(TemplateView):
    def get(self, request, *args, **kwargs):
        # Variables ingresadas
                
        # Creación de la consulta
        resultado_cobertura = obtener_cobertura_paquete_nino()
        
        wb = Workbook()
        
        consultas = [
                ('Cobertura', resultado_cobertura)
        ]
        
        for index, (sheet_name, results) in enumerate(consultas):
            if index == 0:
                ws = wb.active
                ws.title = sheet_name
            else:
                ws = wb.create_sheet(title=sheet_name)
        
            fill_worksheet_cobertura_paquete_nino(ws, results)
        
        ##########################################################################          
        # Establecer el nombre del archivo
        nombre_archivo = "rpt_cobertura_paquete_nino.xlsx"
        # Definir el tipo de respuesta que se va a dar
        response = HttpResponse(content_type="application/ms-excel")
        contenido = "attachment; filename={}".format(nombre_archivo)
        response["Content-Disposition"] = contenido
        wb.save(response)

        return response

def fill_worksheet_cobertura_paquete_nino(ws, results): 
    # cambia el alto de la columna
    ws.row_dimensions[1].height = 14
    ws.row_dimensions[2].height = 14
    ws.row_dimensions[3].height = 3
    ws.row_dimensions[4].height = 25
    ws.row_dimensions[5].height = 3
    ws.row_dimensions[7].height = 20
    ws.row_dimensions[8].height = 130
    # cambia el ancho de la columna
    ws.column_dimensions['A'].width = 2
    ws.column_dimensions['B'].width = 25
    ws.column_dimensions['C'].width = 28
    ws.column_dimensions['D'].width = 30
    ws.column_dimensions['E'].width = 15
    ws.column_dimensions['F'].width = 15
    ws.column_dimensions['G'].width = 8
    ws.column_dimensions['H'].width = 15
    ws.column_dimensions['I'].width = 15
    ws.column_dimensions['J'].width = 8
    ws.column_dimensions['K'].width = 15
    ws.column_dimensions['L'].width = 15
    ws.column_dimensions['M'].width = 8
    ws.column_dimensions['N'].width = 15
    ws.column_dimensions['O'].width = 15
    ws.column_dimensions['P'].width = 8
    ws.column_dimensions['Q'].width = 15
    ws.column_dimensions['R'].width = 15
    ws.column_dimensions['S'].width = 8
    ws.column_dimensions['T'].width = 15
    ws.column_dimensions['U'].width = 15
    ws.column_dimensions['V'].width = 8
    ws.column_dimensions['W'].width = 15
    ws.column_dimensions['X'].width = 15
    ws.column_dimensions['Y'].width = 8
    ws.column_dimensions['Z'].width = 15    
    ws.column_dimensions['AA'].width = 15
    ws.column_dimensions['AB'].width = 8       
    ws.column_dimensions['AC'].width = 15
    ws.column_dimensions['AD'].width = 15    
    ws.column_dimensions['AE'].width = 8
    ws.column_dimensions['AF'].width = 15    
    ws.column_dimensions['AG'].width = 15
    ws.column_dimensions['AH'].width = 8    
    ws.column_dimensions['AI'].width = 15
    ws.column_dimensions['AJ'].width = 15    
    ws.column_dimensions['AK'].width = 8
    ws.column_dimensions['AL'].width = 15    
    ws.column_dimensions['AM'].width = 15
    ws.column_dimensions['AN'].width = 8    
    ws.column_dimensions['AO'].width = 15
    ws.column_dimensions['AP'].width = 15    
    ws.column_dimensions['AQ'].width = 8    
    
    # linea de division
    ws.freeze_panes = 'H9'
    # Configuración del fondo y el borde
    # Definir el color usando formato aRGB (opacidad completa 'FF' + color RGB)
    fill = PatternFill(start_color='FF60D7E0', end_color='FF60D7E0', fill_type='solid')
    # Definir el color anaranjado usando formato aRGB
    orange_fill = PatternFill(start_color='FFE0A960', end_color='FFE0A960', fill_type='solid')
    # Definir los estilos para gris
    gray_fill = PatternFill(start_color='FFD3D3D3', end_color='FFD3D3D3', fill_type='solid')
    # Definir el estilo de color verde
    green_fill = PatternFill(start_color='FF60E0B3', end_color='FF60E0B3', fill_type='solid')
    # Definir el estilo de color amarillo
    yellow_fill = PatternFill(start_color='FFE0DE60', end_color='FFE0DE60', fill_type='solid')
    # Definir el estilo de color azul
    blue_fill = PatternFill(start_color='FF60A2E0', end_color='FF60A2E0', fill_type='solid')
    # Definir el estilo de color verde 2
    green_fill_2 = PatternFill(start_color='FF60E07E', end_color='FF60E07E', fill_type='solid')
    
    green_font = Font(name='Arial', size=8, color='00FF00')  # Verde
    red_font = Font(name='Arial', size=8, color='FF0000')    # Rojo
    
    
    border = Border(left=Side(style='thin', color='00B0F0'),
                    right=Side(style='thin', color='00B0F0'),
                    top=Side(style='thin', color='00B0F0'),
                    bottom=Side(style='thin', color='00B0F0'))
    
    borde_plomo = Border(left=Side(style='thin', color='A9A9A9'), # Plomo
                    right=Side(style='thin', color='A9A9A9'), # Plomo
                    top=Side(style='thin', color='A9A9A9'), # Plomo
                    bottom=Side(style='thin', color='A9A9A9')) # Plomo
    
    border_negro = Border(left=Side(style='thin', color='000000'), # negro
                    right=Side(style='thin', color='000000'),
                    top=Side(style='thin', color='000000'), 
                    bottom=Side(style='thin', color='000000')) 
    
    # Merge cells 
    ws.merge_cells('E7:G7') 
    ws.merge_cells('H7:J7')
    ws.merge_cells('K7:M7')
    ws.merge_cells('N7:P7')
    ws.merge_cells('Q7:S7')
    ws.merge_cells('T7:V7')
    ws.merge_cells('W7:Y7')
    ws.merge_cells('Z7:AB7')
    ws.merge_cells('AC7:AE7')
    ws.merge_cells('AF7:AH7')
    ws.merge_cells('AI7:AK7')
    ws.merge_cells('AL7:AN7')
    ws.merge_cells('AO7:AQ7')

    # Set the value for the merged cell
    ws['E7'] = 'TOTAL'
    ws['H7'] = 'ENERO'
    ws['K7'] = 'FEBRERO'
    ws['N7'] = 'MARZO'
    ws['Q7'] = 'ABRIL'
    ws['T7'] = 'MAYO'
    ws['W7'] = 'JUNIO'
    ws['Z7'] = 'JULIO'
    ws['AC7'] = 'AGOSTO'
    ws['AF7'] = 'SETIEMBRE'
    ws['AI7'] = 'OCTUBRE'
    ws['AL7'] = 'NOVIEMBRE'
    ws['AO7'] = 'DICIEMBRE'

    # Definir el rango desde B3 hasta AA3
    inicio_columna = 'E'
    fin_columna = 'AQ'
    fila = 7
    from openpyxl.utils import column_index_from_string
    # Convertir letras de columna a índices numéricos
    indice_inicio = column_index_from_string(inicio_columna)
    indice_fin = column_index_from_string(fin_columna)

    # Iterar sobre las columnas en el rango especificado
    for col in range(indice_inicio, indice_fin + 1):
        celda = ws.cell(row=fila, column=col)
        celda.border = border_negro
    
    # Definir el rango desde B3 hasta AA3
    inicio_columna_cab = 'B'
    fin_columna_cab = 'AQ'
    fila = 8
    
    # Convertir letras de columna a índices numéricos
    indice_inicio_cab = column_index_from_string(inicio_columna_cab)
    indice_fin_cab = column_index_from_string(fin_columna_cab)

    # Iterar sobre las columnas en el rango especificado
    for col in range(indice_inicio_cab, indice_fin_cab + 1):
        celda = ws.cell(row=fila, column=col)
        celda.border = border_negro
    
    # Apply formatting to the merged cell
    ws['E7'].alignment = Alignment(horizontal="center", vertical="center")
    ws['E7'].font = Font(name='Arial', size=8, bold=True, color='000000')
    ws['E7'].fill = yellow_fill  # Assuming yellow_fill is predefined
    ws['E7'].border = border_negro     # Assuming border is predefined

    ws['H7'].alignment = Alignment(horizontal="center", vertical="center")
    ws['H7'].font = Font(name='Arial', size=8, bold=True, color='000000')
    ws['H7'].fill = blue_fill  # Assuming yellow_fill is predefined
    ws['H7'].border = border_negro 
    
    ws['K7'].alignment = Alignment(horizontal="center", vertical="center")
    ws['K7'].font = Font(name='Arial', size=8, bold=True, color='000000')
    ws['K7'].fill = blue_fill  # Assuming yellow_fill is predefined
    ws['K7'].border = border_negro 
    
    ws['N7'].alignment = Alignment(horizontal="center", vertical="center")
    ws['N7'].font = Font(name='Arial', size=8, bold=True, color='000000')
    ws['N7'].fill = blue_fill  # Assuming yellow_fill is predefined
    ws['N7'].border = border_negro 
    
    ws['Q7'].alignment = Alignment(horizontal="center", vertical="center")
    ws['Q7'].font = Font(name='Arial', size=8, bold=True, color='000000')
    ws['Q7'].fill = blue_fill  # Assuming yellow_fill is predefined
    ws['Q7'].border = border_negro 
    
    ws['T7'].alignment = Alignment(horizontal="center", vertical="center")
    ws['T7'].font = Font(name='Arial', size=8, bold=True, color='000000')
    ws['T7'].fill = blue_fill  # Assuming yellow_fill is predefined
    ws['T7'].border = border_negro 
    
    ws['W7'].alignment = Alignment(horizontal="center", vertical="center")
    ws['W7'].font = Font(name='Arial', size=8, bold=True, color='000000')
    ws['W7'].fill = blue_fill  # Assuming yellow_fill is predefined
    ws['W7'].border = border_negro 
    
    ws['Z7'].alignment = Alignment(horizontal="center", vertical="center")
    ws['Z7'].font = Font(name='Arial', size=8, bold=True, color='000000')
    ws['Z7'].fill = blue_fill  # Assuming yellow_fill is predefined
    ws['Z7'].border = border_negro 
    
    ws['AC7'].alignment = Alignment(horizontal="center", vertical="center")
    ws['AC7'].font = Font(name='Arial', size=8, bold=True, color='000000')
    ws['AC7'].fill = blue_fill  # Assuming yellow_fill is predefined
    ws['AC7'].border = border_negro 
    
    ws['AF7'].alignment = Alignment(horizontal="center", vertical="center")
    ws['AF7'].font = Font(name='Arial', size=8, bold=True, color='000000')
    ws['AF7'].fill = blue_fill  # Assuming yellow_fill is predefined
    ws['AF7'].border = border_negro 
    
    ws['AI7'].alignment = Alignment(horizontal="center", vertical="center")
    ws['AI7'].font = Font(name='Arial', size=8, bold=True, color='000000')
    ws['AI7'].fill = blue_fill  # Assuming yellow_fill is predefined
    ws['AI7'].border = border_negro 
    
    ws['AL7'].alignment = Alignment(horizontal="center", vertical="center")
    ws['AL7'].font = Font(name='Arial', size=8, bold=True, color='000000')
    ws['AL7'].fill = blue_fill  # Assuming yellow_fill is predefined
    ws['AL7'].border = border_negro 
    
    ws['AO7'].alignment = Alignment(horizontal="center", vertical="center")
    ws['AO7'].font = Font(name='Arial', size=8, bold=True, color='000000')
    ws['AO7'].fill = blue_fill  # Assuming yellow_fill is predefined
    ws['AO7'].border = border_negro 
    
    ## crea titulo del reporte
    ws['B1'].alignment = Alignment(horizontal= "left", vertical="center")
    ws['B1'].font = Font(name = 'Arial', size= 7, bold = True)
    ws['B1'] = 'OFICINA DE TECNOLOGIAS DE LA INFORMACION'
    
    ws['B2'].alignment = Alignment(horizontal= "left", vertical="center")
    ws['B2'].font = Font(name = 'Arial', size= 7, bold = True)
    ws['B2'] = 'DIRECCION REGIONAL DE SALUD JUNIN'
    
    ws['B4'].alignment = Alignment(horizontal= "left", vertical="center")
    ws['B4'].font = Font(name = 'Arial', size= 12, bold = True)
    ws['B4'] = 'COBERURA DEL INDICADOR MC-02. NIÑAS Y NIÑOS MENORES DE 12 MESES DE EDAD PROCEDENTES DE LOS QUINTILES 1 Y 2 DE POBREZA DEPARTAMENTAL QUE RECIBEN EL PAQUETE INTEGRADO DE SERVICIOS'
    
    ws['B6'].alignment = Alignment(horizontal= "left", vertical="center")
    ws['B6'].font = Font(name = 'Arial', size= 7, bold = True, color='0000CC')
    ws['B6'] ='NOTAS: Primera columna "NUMERADOR", segunda columna "DENOMINADOR", tercera columna "PORCENTAJE AVANCE"'
        
    ws['B8'].alignment = Alignment(horizontal= "center", vertical="center")
    ws['B8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['B8'].fill = yellow_fill
    ws['B8'].border = border_negro
    ws['B8'] = 'RED'
    
    ws['C8'].alignment = Alignment(horizontal= "center", vertical="center")
    ws['C8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['C8'].fill = yellow_fill
    ws['C8'].border = border_negro
    ws['C8'] = 'MICRORED'
    
    ws['D8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['D8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['D8'].fill = yellow_fill
    ws['D8'].border = border_negro
    ws['D8'] = 'ESTABLECIMIENTO'      
    
    ws['E8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['E8'].font = Font(name = 'Arial', size= 7, bold = True, color='000000')
    ws['E8'].fill = yellow_fill
    ws['E8'].border = border_negro
    ws['E8'] = 'N° de niñas y niños del denominador que reciben el paquete integrado de servicios según edad, y que han sido registrados en HIS y cuentan con DNI emitido'
    
    ws['F8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['F8'].font = Font(name = 'Arial', size= 7, bold = True, color='000000')
    ws['F8'].fill = yellow_fill
    ws['F8'].border = border_negro
    ws['F8'] = 'N° de niñas y niños menores de 12 meses de edad en el mes de medición, procedentes de distritos de quintiles 1 y 2 de pobreza departamental, registrados en el padrón nominal con DNI o CNV, con tipo de seguro MINSA'
    
    ws['G8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['G8'].font = Font(name = 'Arial', size= 7, bold = True, color='000000')
    ws['G8'].fill = yellow_fill
    ws['G8'].border = border_negro
    ws['G8'] = '% Avance (Num/Den)'    
    
    ws['H8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['H8'].font = Font(name = 'Arial', size= 7, bold = True, color='000000')
    ws['H8'].fill = blue_fill
    ws['H8'].border = border_negro
    ws['H8'] = 'N° de niñas y niños del denominador que reciben el paquete integrado de servicios según edad, y que han sido registrados en HIS y cuentan con DNI emitido'    
    
    ws['I8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['I8'].font = Font(name = 'Arial', size= 7, bold = True, color='000000')
    ws['I8'].fill = blue_fill
    ws['I8'].border = border_negro
    ws['I8'] = 'N° de niñas y niños menores de 12 meses de edad en el mes de medición, procedentes de distritos de quintiles 1 y 2 de pobreza departamental, registrados en el padrón nominal con DNI o CNV, con tipo de seguro MINSA' 
    
    ws['J8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['J8'].font = Font(name = 'Arial', size= 7, bold = True, color='000000')
    ws['J8'].fill = gray_fill
    ws['J8'].border = border_negro
    ws['J8'] = '% Avance (Num/Den)'    
    
    ws['K8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['K8'].font = Font(name = 'Arial', size= 7, bold = True, color='000000')
    ws['K8'].fill = blue_fill
    ws['K8'].border = border_negro
    ws['K8'] = 'N° de niñas y niños del denominador que reciben el paquete integrado de servicios según edad, y que han sido registrados en HIS y cuentan con DNI emitido'     
    
    ws['L8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['L8'].font = Font(name = 'Arial', size= 7, bold = True, color='000000')
    ws['L8'].fill = blue_fill
    ws['L8'].border = border_negro
    ws['L8'] = 'N° de niñas y niños menores de 12 meses de edad en el mes de medición, procedentes de distritos de quintiles 1 y 2 de pobreza departamental, registrados en el padrón nominal con DNI o CNV, con tipo de seguro MINSA' 
    
    ws['M8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['M8'].font = Font(name = 'Arial', size= 7, bold = True, color='000000')
    ws['M8'].fill = gray_fill
    ws['M8'].border = border_negro
    ws['M8'] = '% Avance (Num/Den)'
    
    ws['N8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['N8'].font = Font(name = 'Arial', size= 7, bold = True, color='000000')
    ws['N8'].fill = blue_fill
    ws['N8'].border = border_negro
    ws['N8'] = 'N° de niñas y niños del denominador que reciben el paquete integrado de servicios según edad, y que han sido registrados en HIS y cuentan con DNI emitido'   
    
    ws['O8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['O8'].font = Font(name = 'Arial', size= 7, bold = True, color='000000')
    ws['O8'].fill = blue_fill
    ws['O8'].border = border_negro
    ws['O8'] = 'N° de niñas y niños menores de 12 meses de edad en el mes de medición, procedentes de distritos de quintiles 1 y 2 de pobreza departamental, registrados en el padrón nominal con DNI o CNV, con tipo de seguro MINSA'
    
    ws['P8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['P8'].font = Font(name = 'Arial', size= 7, bold = True, color='000000')
    ws['P8'].fill = gray_fill
    ws['P8'].border = border_negro
    ws['P8'] = '% Avance (Num/Den)'     
    
    ws['Q8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['Q8'].font = Font(name = 'Arial', size= 7, bold = True, color='000000')
    ws['Q8'].fill = blue_fill
    ws['Q8'].border = border_negro
    ws['Q8'] = 'N° de niñas y niños del denominador que reciben el paquete integrado de servicios según edad, y que han sido registrados en HIS y cuentan con DNI emitido'   
    
    ws['R8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['R8'].font = Font(name = 'Arial', size= 7, bold = True, color='000000')
    ws['R8'].fill = blue_fill
    ws['R8'].border = border_negro
    ws['R8'] = 'N° de niñas y niños menores de 12 meses de edad en el mes de medición, procedentes de distritos de quintiles 1 y 2 de pobreza departamental, registrados en el padrón nominal con DNI o CNV, con tipo de seguro MINSA' 
    
    ws['S8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['S8'].font = Font(name = 'Arial', size= 7, bold = True, color='000000')
    ws['S8'].fill = gray_fill
    ws['S8'].border = border_negro
    ws['S8'] = '% Avance (Num/Den)'    
    
    ws['T8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['T8'].font = Font(name = 'Arial', size= 7, bold = True, color='000000')
    ws['T8'].fill = blue_fill
    ws['T8'].border = border_negro
    ws['T8'] = 'N° de niñas y niños del denominador que reciben el paquete integrado de servicios según edad, y que han sido registrados en HIS y cuentan con DNI emitido'    
    
    ws['U8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['U8'].font = Font(name = 'Arial', size= 7, bold = True, color='000000')
    ws['U8'].fill = blue_fill
    ws['U8'].border = border_negro
    ws['U8'] = 'N° de niñas y niños menores de 12 meses de edad en el mes de medición, procedentes de distritos de quintiles 1 y 2 de pobreza departamental, registrados en el padrón nominal con DNI o CNV, con tipo de seguro MINSA'
    
    ws['V8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['V8'].font = Font(name = 'Arial', size= 7, bold = True, color='000000')
    ws['V8'].fill = gray_fill
    ws['V8'].border = border_negro
    ws['V8'] = '% Avance (Num/Den)'    
    
    ws['W8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['W8'].font = Font(name = 'Arial', size= 7, bold = True, color='000000')
    ws['W8'].fill = blue_fill
    ws['W8'].border = border_negro
    ws['W8'] = 'N° de niñas y niños del denominador que reciben el paquete integrado de servicios según edad, y que han sido registrados en HIS y cuentan con DNI emitido'   
        
    ws['X8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['X8'].font = Font(name = 'Arial', size= 7, bold = True, color='000000')
    ws['X8'].fill = blue_fill
    ws['X8'].border = border_negro
    ws['X8'] = 'N° de niñas y niños menores de 12 meses de edad en el mes de medición, procedentes de distritos de quintiles 1 y 2 de pobreza departamental, registrados en el padrón nominal con DNI o CNV, con tipo de seguro MINSA'

    ws['Y8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['Y8'].font = Font(name = 'Arial', size= 7, bold = True, color='000000')
    ws['Y8'].fill = gray_fill
    ws['Y8'].border = border_negro
    ws['Y8'] = '% Avance (Num/Den)'    
    
    ws['Z8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['Z8'].font = Font(name = 'Arial', size= 7, bold = True, color='000000')
    ws['Z8'].fill = blue_fill
    ws['Z8'].border = border_negro
    ws['Z8'] = 'N° de niñas y niños del denominador que reciben el paquete integrado de servicios según edad, y que han sido registrados en HIS y cuentan con DNI emitido'   

    ws['AA8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AA8'].font = Font(name = 'Arial', size= 7, bold = True, color='000000')
    ws['AA8'].fill = blue_fill
    ws['AA8'].border = border_negro
    ws['AA8'] = 'N° de niñas y niños menores de 12 meses de edad en el mes de medición, procedentes de distritos de quintiles 1 y 2 de pobreza departamental, registrados en el padrón nominal con DNI o CNV, con tipo de seguro MINSA'
    
    ws['AB8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AB8'].font = Font(name = 'Arial', size= 7, bold = True, color='000000')
    ws['AB8'].fill = gray_fill
    ws['AB8'].border = border_negro
    ws['AB8'] = '% Avance (Num/Den)'    
    
    ws['AC8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AC8'].font = Font(name = 'Arial', size= 7, bold = True, color='000000')
    ws['AC8'].fill = blue_fill
    ws['AC8'].border = border_negro
    ws['AC8'] = 'N° de niñas y niños del denominador que reciben el paquete integrado de servicios según edad, y que han sido registrados en HIS y cuentan con DNI emitido'   
    
    ws['AD8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AD8'].font = Font(name = 'Arial', size= 7, bold = True, color='000000')
    ws['AD8'].fill = blue_fill
    ws['AD8'].border = border_negro
    ws['AD8'] = 'N° de niñas y niños menores de 12 meses de edad en el mes de medición, procedentes de distritos de quintiles 1 y 2 de pobreza departamental, registrados en el padrón nominal con DNI o CNV, con tipo de seguro MINSA'
    
    ws['AE8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AE8'].font = Font(name = 'Arial', size= 7, bold = True, color='000000')
    ws['AE8'].fill = gray_fill
    ws['AE8'].border = border_negro
    ws['AE8'] = '% Avance (Num/Den)'    
    
    ws['AF8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AF8'].font = Font(name = 'Arial', size= 7, bold = True, color='000000')
    ws['AF8'].fill = blue_fill
    ws['AF8'].border = border_negro
    ws['AF8'] = 'N° de niñas y niños del denominador que reciben el paquete integrado de servicios según edad, y que han sido registrados en HIS y cuentan con DNI emitido'   
    
    ws['AG8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AG8'].font = Font(name = 'Arial', size= 7, bold = True, color='000000')
    ws['AG8'].fill = blue_fill
    ws['AG8'].border = border_negro
    ws['AG8'] = 'N° de niñas y niños menores de 12 meses de edad en el mes de medición, procedentes de distritos de quintiles 1 y 2 de pobreza departamental, registrados en el padrón nominal con DNI o CNV, con tipo de seguro MINSA' 
    
    ws['AH8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AH8'].font = Font(name = 'Arial', size= 7, bold = True, color='000000')
    ws['AH8'].fill = gray_fill
    ws['AH8'].border = border_negro
    ws['AH8'] = '% Avance (Num/Den)'    
    
    ws['AI8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AI8'].font = Font(name = 'Arial', size= 7, bold = True, color='000000')
    ws['AI8'].fill = blue_fill
    ws['AI8'].border = border_negro
    ws['AI8'] = 'N° de niñas y niños del denominador que reciben el paquete integrado de servicios según edad, y que han sido registrados en HIS y cuentan con DNI emitido'    
    
    ws['AJ8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AJ8'].font = Font(name = 'Arial', size= 7, bold = True, color='000000')
    ws['AJ8'].fill = blue_fill
    ws['AJ8'].border = border_negro
    ws['AJ8'] = 'N° de niñas y niños menores de 12 meses de edad en el mes de medición, procedentes de distritos de quintiles 1 y 2 de pobreza departamental, registrados en el padrón nominal con DNI o CNV, con tipo de seguro MINSA'
    
    ws['AK8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AK8'].font = Font(name = 'Arial', size= 7, bold = True, color='000000')
    ws['AK8'].fill = gray_fill
    ws['AK8'].border = border_negro
    ws['AK8'] = '% Avance (Num/Den)'    
    
    ws['AL8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AL8'].font = Font(name = 'Arial', size= 7, bold = True, color='000000')
    ws['AL8'].fill = blue_fill
    ws['AL8'].border = border_negro
    ws['AL8'] = 'N° de niñas y niños del denominador que reciben el paquete integrado de servicios según edad, y que han sido registrados en HIS y cuentan con DNI emitido'    
    
    ws['AM8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AM8'].font = Font(name = 'Arial', size= 7, bold = True, color='000000')
    ws['AM8'].fill = blue_fill
    ws['AM8'].border = border_negro
    ws['AM8'] = 'N° de niñas y niños menores de 12 meses de edad en el mes de medición, procedentes de distritos de quintiles 1 y 2 de pobreza departamental, registrados en el padrón nominal con DNI o CNV, con tipo de seguro MINSA'
    
    ws['AN8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AN8'].font = Font(name = 'Arial', size= 7, bold = True, color='000000')
    ws['AN8'].fill = gray_fill
    ws['AN8'].border = border_negro
    ws['AN8'] = '% Avance (Num/Den)'    
    
    ws['AO8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AO8'].font = Font(name = 'Arial', size= 7, bold = True, color='000000')
    ws['AO8'].fill = blue_fill
    ws['AO8'].border = border_negro
    ws['AO8'] = 'N° de niñas y niños del denominador que reciben el paquete integrado de servicios según edad, y que han sido registrados en HIS y cuentan con DNI emitido'   
    
    ws['AP8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AP8'].font = Font(name = 'Arial', size= 7, bold = True, color='000000')
    ws['AP8'].fill = blue_fill
    ws['AP8'].border = border_negro
    ws['AP8'] = 'N° de niñas y niños menores de 12 meses de edad en el mes de medición, procedentes de distritos de quintiles 1 y 2 de pobreza departamental, registrados en el padrón nominal con DNI o CNV, con tipo de seguro MINSA' 
    
    ws['AQ8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AQ8'].font = Font(name = 'Arial', size= 7, bold = True, color='000000')
    ws['AQ8'].fill = gray_fill
    ws['AQ8'].border = border_negro
    ws['AQ8'] = '% Avance (Num/Den)'    
    
    # Definir estilos
    header_font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    centered_alignment = Alignment(horizontal='center')
    border = Border(left=Side(style='thin', color='A9A9A9'),
            right=Side(style='thin', color='A9A9A9'),
            top=Side(style='thin', color='A9A9A9'),
            bottom=Side(style='thin', color='A9A9A9'))
    header_fill = PatternFill(patternType='solid', fgColor='00B0F0')
    
    # Definir los caracteres especiales de check y X
    
    # Escribir datos
    for row, record in enumerate(results, start=9):
        for col, value in enumerate(record, start=2):
            cell = ws.cell(row=row, column=col, value=value)

            # Alinear a la izquierda solo en las columnas 6,14,15,16
            if col in [2, 3, 4]:
                cell.alignment = Alignment(horizontal='left')
            else:
                cell.alignment = Alignment(horizontal='center')

            # Aplicar formato basado en el valor para columnas específicas
            if col in [7, 10, 13, 16, 19, 22, 25, 28, 31, 34, 37, 40, 43]:
                try:
                    value_float = float(value)
                except (ValueError, TypeError):
                    # Si el valor no es numérico, aplicar formato por defecto
                    cell.font = Font(name='Arial', size=7)
                else:
                    # Si los valores están entre 0 y 100, dividimos entre 100
                    if value_float > 1:
                        value_float = value_float / 100
                        cell.value = value_float

                    # Establecer el formato de número a porcentaje
                    cell.number_format = '0.0%'

                    if value_float >= 0.75:
                        # Colorear la celda de verde
                        cell.fill = PatternFill(patternType='solid', fgColor='00B050')  # Fondo verde
                        cell.font = Font(name='Arial', size=7, color='000000')  # Letra negra
                    else:
                        # Colorear la celda de rojo con letras blancas
                        cell.fill = PatternFill(patternType='solid', fgColor='FF0000')  # Fondo rojo
                        cell.font = Font(name='Arial', size=7, color='FFFFFF')  # Letra blanca
            # Fuente normal para otras columnas
            else:
                cell.font = Font(name='Arial', size=8)
            
            cell.border = border