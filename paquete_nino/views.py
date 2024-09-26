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
    ws.column_dimensions['D'].width = 4
    ws.column_dimensions['E'].width = 5
    ws.column_dimensions['F'].width = 33
    ws.column_dimensions['G'].width = 9
    ws.column_dimensions['H'].width = 5
    ws.column_dimensions['I'].width = 5
    ws.column_dimensions['J'].width = 5
    ws.column_dimensions['K'].width = 5
    ws.column_dimensions['L'].width = 9
    ws.column_dimensions['M'].width = 9
    ws.column_dimensions['N'].width = 9
    ws.column_dimensions['O'].width = 3
    ws.column_dimensions['P'].width = 9
    ws.column_dimensions['Q'].width = 3
    ws.column_dimensions['R'].width = 9
    ws.column_dimensions['S'].width = 3
    ws.column_dimensions['T'].width = 9
    ws.column_dimensions['U'].width = 3
    ws.column_dimensions['V'].width = 9
    ws.column_dimensions['W'].width = 9
    ws.column_dimensions['X'].width = 3
    ws.column_dimensions['Y'].width = 9
    ws.column_dimensions['Z'].width = 3
    ws.column_dimensions['AA'].width = 9
    ws.column_dimensions['AB'].width = 3
    ws.column_dimensions['AC'].width = 9
    ws.column_dimensions['AD'].width = 3
    ws.column_dimensions['AE'].width = 9
    ws.column_dimensions['AF'].width = 3
    ws.column_dimensions['AG'].width = 9
    ws.column_dimensions['AH'].width = 3
    ws.column_dimensions['AI'].width = 9
    ws.column_dimensions['AJ'].width = 3
    ws.column_dimensions['AK'].width = 9
    ws.column_dimensions['AL'].width = 3
    ws.column_dimensions['AM'].width = 9
    ws.column_dimensions['AN'].width = 3
    ws.column_dimensions['AO'].width = 9
    ws.column_dimensions['AP'].width = 3
    ws.column_dimensions['AQ'].width = 9
    ws.column_dimensions['AR'].width = 3
    ws.column_dimensions['AS'].width = 9
    ws.column_dimensions['AT'].width = 9
    ws.column_dimensions['AU'].width = 9
    ws.column_dimensions['AV'].width = 9
    ws.column_dimensions['AW'].width = 3
    ws.column_dimensions['AX'].width = 9
    ws.column_dimensions['AY'].width = 9
    ws.column_dimensions['AZ'].width = 3
    ws.column_dimensions['BA'].width = 9
    ws.column_dimensions['BB'].width = 9
    ws.column_dimensions['BC'].width = 9
    ws.column_dimensions['BD'].width = 3
    ws.column_dimensions['BE'].width = 9
    ws.column_dimensions['BF'].width = 9
    ws.column_dimensions['BG'].width = 9
    ws.column_dimensions['BH'].width = 3
    ws.column_dimensions['BI'].width = 9
    ws.column_dimensions['BJ'].width = 9
    ws.column_dimensions['BK'].width = 9
    ws.column_dimensions['BL'].width = 9
    ws.column_dimensions['BM'].width = 9
    ws.column_dimensions['BN'].width = 3
    ws.column_dimensions['BO'].width = 9
    ws.column_dimensions['BP'].width = 9
    ws.column_dimensions['BQ'].width = 9
    ws.column_dimensions['BR'].width = 9
    ws.column_dimensions['BS'].width = 9
    ws.column_dimensions['BT'].width = 9
    ws.column_dimensions['BU'].width = 3
    ws.column_dimensions['BV'].width = 9
    ws.column_dimensions['BW'].width = 9
    ws.column_dimensions['BX'].width = 13
    ws.column_dimensions['BY'].width = 10
    ws.column_dimensions['BZ'].width = 9
    ws.column_dimensions['CA'].width = 9
    ws.column_dimensions['CB'].width = 16
    ws.column_dimensions['CC'].width = 20
    ws.column_dimensions['CD'].width = 20
    ws.column_dimensions['CE'].width = 6
    ws.column_dimensions['CF'].width = 25
    ws.column_dimensions['CG'].width = 9
    ws.column_dimensions['CH'].width = 33
    ws.column_dimensions['CI'].width = 9
    ws.column_dimensions['CJ'].width = 9
    ws.column_dimensions['CK'].width = 33
    
    # linea de division
    ws.freeze_panes = 'N9'
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
    ws['D8'].fill = gray_fill
    ws['D8'].border = border
    ws['D8'] = 'VAL 30D'      
    
    ws['E8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['E8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['E8'].fill = gray_fill
    ws['E8'].border = border
    ws['E8'] = 'VAL 60D' 
    
    ws['F8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['F8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['F8'].fill = fill
    ws['F8'].border = border
    ws['F8'] = 'NOMBRE COMPLETO DE NIÑO/A'     
    
    ws['G8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['G8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['G8'].fill = fill
    ws['G8'].border = border
    ws['G8'] = 'FECHA NAC'    
    
    ws['H8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['H8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['H8'].fill = fill
    ws['H8'].border = border
    ws['H8'] = 'EDAD DIAS'    
    
    ws['I8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['I8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['I8'].fill = fill
    ws['I8'].border = border
    ws['I8'] = 'ED AÑO'    
    
    ws['J8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['J8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['J8'].fill = fill
    ws['J8'].border = border
    ws['J8'] = 'ED MES'  
    
    ws['K8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['K8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['K8'].fill = fill
    ws['K8'].border = border
    ws['K8'] = 'ED DIA'  
    
    ws['L8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['L8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['L8'].fill = fill
    ws['L8'].border = border
    ws['L8'] = 'FECHA INICIO'  
    
    ws['M8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['M8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['M8'].fill = fill
    ws['M8'].border = border
    ws['M8'] = 'SEGURO'  
    
    ws['N8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['N8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['N8'].fill = green_fill_2
    ws['N8'].border = border
    ws['N8'] = 'CRED RN1'  
    
    ws['O8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['O8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['O8'].fill = green_fill_2
    ws['O8'].border = border
    ws['O8'] = 'V'  
    
    ws['P8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['P8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['P8'].fill = green_fill_2
    ws['P8'].border = border
    ws['P8'] = 'CRED RN2'  
    
    ws['Q8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['Q8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['Q8'].fill = green_fill_2
    ws['Q8'].border = border
    ws['Q8'] = 'V'    
    
    ws['R8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['R8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['R8'].fill = green_fill_2
    ws['R8'].border = border
    ws['R8'] = 'CRED RN3' 
    
    ws['S8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['S8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['S8'].fill = green_fill_2
    ws['S8'].border = border
    ws['S8'] = 'V' 
    
    ws['T8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['T8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['T8'].fill = green_fill_2
    ws['T8'].border = border
    ws['T8'] = 'CRED RN4' 
    
    ws['U8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['U8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['U8'].fill = green_fill_2
    ws['U8'].border = border
    ws['U8'] = 'V' 
    
    ws['V8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['V8'].font = Font(name = 'Arial', size= 7, bold = True, color='000000')
    ws['V8'].fill = gray_fill
    ws['V8'].border = border
    ws['V8'] = 'VAL CRED RN' 
    
    ws['W8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['W8'].font = Font(name = 'Arial', size= 7, bold = True, color='000000')
    ws['W8'].fill = green_fill
    ws['W8'].border = border
    ws['W8'] = 'CRED 1M'   
    
    ws['X8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['X8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['X8'].fill = green_fill
    ws['X8'].border = border
    ws['X8'] = 'V' 
    
    ws['Y8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['Y8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['Y8'].fill = green_fill
    ws['Y8'].border = border
    ws['Y8'] = 'CRED 2M' 
    
    ws['Z8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['Z8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['Z8'].fill = green_fill
    ws['Z8'].border = border
    ws['Z8'] = 'V' 
    
    ws['AA8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AA8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AA8'].fill = green_fill
    ws['AA8'].border = border
    ws['AA8'] = 'CRED 3M' 
    
    ws['AB8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AB8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AB8'].fill = green_fill
    ws['AB8'].border = border
    ws['AB8'] = 'V'     
    
    ws['AC8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AC8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AC8'].fill = green_fill
    ws['AC8'].border = border
    ws['AC8'] = 'CRED 4M' 
    
    ws['AD8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AD8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AD8'].fill = green_fill
    ws['AD8'].border = border
    ws['AD8'] = 'V' 
    
    ws['AE8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AE8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AE8'].fill = green_fill
    ws['AE8'].border = border
    ws['AE8'] = 'CRED 5M' 
    
    ws['AF8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AF8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AF8'].fill = green_fill
    ws['AF8'].border = border
    ws['AF8'] = 'V' 
    
    ws['AG8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AG8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AG8'].fill = green_fill
    ws['AG8'].border = border
    ws['AG8'] = 'CRED 6M' 
    
    ws['AH8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AH8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AH8'].fill = green_fill
    ws['AH8'].border = border
    ws['AH8'] = 'V' 
    
    ws['AI8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AI8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AI8'].fill = green_fill
    ws['AI8'].border = border
    ws['AI8'] = 'CRED 7M' 
    
    ws['AJ8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AJ8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AJ8'].fill = green_fill
    ws['AJ8'].border = border
    ws['AJ8'] = 'V' 
    
    ws['AK8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AK8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AK8'].fill = green_fill
    ws['AK8'].border = border
    ws['AK8'] = 'CRED 8M' 
    
    ws['AL8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AL8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AL8'].fill = green_fill
    ws['AL8'].border = border
    ws['AL8'] = 'V' 
    
    ws['AM8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AM8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AM8'].fill = green_fill
    ws['AM8'].border = border
    ws['AM8'] = 'CRED 9M' 
    
    ws['AN8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AN8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AN8'].fill = green_fill
    ws['AN8'].border = border
    ws['AN8'] = 'V' 
    
    ws['AO8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AO8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AO8'].fill = green_fill
    ws['AO8'].border = border
    ws['AO8'] = 'CRED 10M' 
    
    ws['AP8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AP8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AP8'].fill = green_fill
    ws['AP8'].border = border
    ws['AP8'] = 'V' 
    
    ws['AQ8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AQ8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AQ8'].fill = green_fill
    ws['AQ8'].border = border
    ws['AQ8'] = 'CRED 11M' 
    
    ws['AR8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AR8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AR8'].fill = green_fill
    ws['AR8'].border = border
    ws['AR8'] = 'V' 
    
    ws['AS8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AS8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AS8'].fill = gray_fill
    ws['AS8'].border = border
    ws['AS8'] = 'VAL CRED' 
    
    ws['AT8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AT8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AT8'].fill = gray_fill
    ws['AT8'].border = border
    ws['AT8'] = 'VAL EDAD CRED' 
    
    ws['AU8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AU8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AU8'].fill = yellow_fill
    ws['AU8'].border = border
    ws['AU8'] = '1° NEUMO' 
    
    ws['AV8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AV8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AV8'].fill = yellow_fill
    ws['AV8'].border = border
    ws['AV8'] = '2° NEUMO' 
    
    ws['AW8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AW8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AW8'].fill = yellow_fill
    ws['AW8'].border = border
    ws['AW8'] = 'V' 
    
    ws['AX8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AX8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AX8'].fill = yellow_fill
    ws['AX8'].border = border
    ws['AX8'] = '1° ROTA' 
    
    ws['AY8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AY8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AY8'].fill = yellow_fill
    ws['AY8'].border = border
    ws['AY8'] = '2° ROTA' 
    
    ws['AZ8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['AZ8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['AZ8'].fill = yellow_fill
    ws['AZ8'].border = border
    ws['AZ8'] = 'V' 
    
    ws['BA8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BA8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BA8'].fill = yellow_fill
    ws['BA8'].border = border
    ws['BA8'] = '1° POLIO' 
    
    ws['BB8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BB8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BB8'].fill = yellow_fill
    ws['BB8'].border = border
    ws['BB8'] = '2° POLIO'     
    
    ws['BC8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BC8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BC8'].fill = yellow_fill
    ws['BC8'].border = border
    ws['BC8'] = '3° POLIO' 
    
    ws['BD8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BD8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BD8'].fill = yellow_fill
    ws['BD8'].border = border
    ws['BD8'] = 'V' 
    
    ws['BE8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BE8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BE8'].fill = yellow_fill
    ws['BE8'].border = border
    ws['BE8'] = '1° PENTA' 
    
    ws['BF8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BF8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BF8'].fill = yellow_fill
    ws['BF8'].border = border
    ws['BF8'] = '2° PENTA' 
    
    ws['BG8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BG8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BG8'].fill = yellow_fill
    ws['BG8'].border = border
    ws['BG8'] = '3° PENTA' 
    
    ws['BH8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BH8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BH8'].fill = yellow_fill
    ws['BH8'].border = border
    ws['BH8'] = 'V' 
    
    ws['BI8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BI8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BI8'].fill = gray_fill
    ws['BI8'].border = border
    ws['BI8'] = 'VAL VACUNA' 
    
    ws['BJ8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BJ8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BJ8'].fill = blue_fill
    ws['BJ8'].border = border
    ws['BJ8'] = 'DOSAJE' 
    
    ws['BK8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BK8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BK8'].fill = gray_fill
    ws['BK8'].border = border
    ws['BK8'] = 'VAL DOSAJE' 
    
    ws['BL8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BL8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BL8'].fill = blue_fill
    ws['BL8'].border = border
    ws['BL8'] = '1° SUPLE 4M' 
    
    ws['BM8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BM8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BM8'].fill = blue_fill
    ws['BM8'].border = border
    ws['BM8'] = '2° SUPLE 4M' 
    
    ws['BN8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BN8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BN8'].fill = blue_fill
    ws['BN8'].border = border
    ws['BN8'] = 'V' 
    
    ws['BO8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BO8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BO8'].fill = blue_fill
    ws['BO8'].border = border
    ws['BO8'] = '1° SUPLE 6M'  
    
    ws['BP8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BP8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BP8'].fill = blue_fill
    ws['BP8'].border = border
    ws['BP8'] = '2° SUPLE 6M' 
    
    ws['BQ8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BQ8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BQ8'].fill = blue_fill
    ws['BQ8'].border = border
    ws['BQ8'] = '3° SUPLE 6M' 
    
    ws['BR8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BR8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BR8'].fill = blue_fill
    ws['BR8'].border = border
    ws['BR8'] = '4° SUPLE 6M' 
    
    ws['BS8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BS8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BS8'].fill = blue_fill
    ws['BS8'].border = border
    ws['BS8'] = '5° SUPLE 6M' 
    
    ws['BT8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BT8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BT8'].fill = blue_fill
    ws['BT8'].border = border
    ws['BT8'] = '6° SUPLE 6M' 
    
    ws['BU8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BU8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BU8'].fill = blue_fill
    ws['BU8'].border = border
    ws['BU8'] = 'V'
    
    ws['BV8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BV8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BV8'].fill = gray_fill
    ws['BV8'].border = border
    ws['BV8'] = 'VAL SUPLE' 
    
    ws['BW8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BW8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['BW8'].fill = fill
    ws['BW8'].border = border
    ws['BW8'] = 'FECHA FIN' 
    
    ws['BX8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BX8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['BX8'].fill = fill
    ws['BX8'].border = border
    ws['BX8'] = 'MES' 
    
    ws['BY8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BY8'].font = Font(name = 'Arial', size= 8, bold = True, color='000000')
    ws['BY8'].fill = gray_fill
    ws['BY8'].border = border
    ws['BY8'] = 'IND' 
    
    ws['BZ8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['BZ8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['BZ8'].fill = orange_fill
    ws['BZ8'].border = border
    ws['BZ8'] = 'UBIGEO'  
    
    ws['CA8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['CA8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['CA8'].fill = orange_fill
    ws['CA8'].border = border
    ws['CA8'] = 'PROVINCIA'       
    
    ws['CB8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['CB8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['CB8'].fill = orange_fill
    ws['CB8'].border = border
    ws['CB8'] = 'DISTRITO' 
    
    ws['CC8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['CC8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['CC8'].fill = orange_fill
    ws['CC8'].border = border
    ws['CC8'] = 'RED'  
    
    ws['CD8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['CD8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['CD8'].fill = orange_fill
    ws['CD8'].border = border
    ws['CD8'] = 'MICRORED'  
    
    ws['CE8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['CE8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['CE8'].fill = orange_fill
    ws['CE8'].border = border
    ws['CE8'] = 'COD EST'  
    
    ws['CF8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['CF8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['CF8'].fill = orange_fill
    ws['CF8'].border = border
    ws['CF8'] = 'ESTABLECIMIENTO'  
    
    ws['CG8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['CG8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['CG8'].fill = orange_fill
    ws['CG8'].border = border
    ws['CG8'] = 'DNI MADRE'  
    
    ws['CH8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['CH8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['CH8'].fill = orange_fill
    ws['CH8'].border = border
    ws['CH8'] = 'NOMBRE DE MADRE'  
    
    ws['CI8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['CI8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['CI8'].fill = orange_fill
    ws['CI8'].border = border
    ws['CI8'] = 'NUM CELULAR'  
    
    ws['CJ8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['CJ8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['CJ8'].fill = orange_fill
    ws['CJ8'].border = border
    ws['CJ8'] = 'CONVENIO'  
    
    ws['CK8'].alignment = Alignment(horizontal= "center", vertical="center", wrap_text=True)
    ws['CK8'].font = Font(name = 'Arial', size= 8, bold = True, color='FFFFFF')
    ws['CK8'].fill = orange_fill
    ws['CK8'].border = border
    ws['CK8'] = 'OBSERVACIONES'  
    
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
    
    
    # Escribir datos
    for row, record in enumerate(results, start=9):
        for col, value in enumerate(record, start=2):
            cell = ws.cell(row=row, column=col, value=value)

            # Alinear a la izquierda solo en las columnas 6,14,15,16
            if col in [6, 80, 81, 82, 84, 86]:
                cell.alignment = Alignment(horizontal='left')
            else:
                cell.alignment = Alignment(horizontal='center')

            # Aplicar color en la columna 27
            if col == 77:
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
            
            # Aplicar color de letra en las columnas 7 y 17
            elif col in [22, 45, 46, 61, 63, 74]:
                if isinstance(value, str):
                    value_upper = value.strip().upper()
                    if value_upper == "NO CUMPLE":
                        cell.font = Font(name='Arial', size=7, color="FF0000")  # Letra roja
                    elif value_upper == "CUMPLE":
                        cell.font = Font(name='Arial', size=7, color="00B050")  # Letra verde
                    else:
                        cell.font = Font(name='Arial', size=7)
                else:
                    cell.font = Font(name='Arial', size=7)
            # Fuente normal para otras columnas
            else:
                cell.font = Font(name='Arial', size=8)  # Fuente normal para otras columnas

            # Aplicar caracteres especiales check y X
            if col in [4, 5, 15, 17, 19, 21, 24, 26, 28, 30, 32, 34, 36, 38, 40, 42, 44, 49, 52, 56, 60, 66, 73]:
                if value == 1:
                    cell.value = check_mark  # Insertar check
                    cell.font = Font(name='Arial', size=10, color='00B050')  # Letra verde
                elif value == 0:
                    cell.value = x_mark  # Insertar X
                    cell.font = Font(name='Arial', size=10, color='FF0000')  # Letra roja
                else:
                    cell.font = Font(name='Arial', size=8)  # Fuente normal si no es 1 o 0
            
                        
            cell.border = border