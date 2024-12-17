# app_name/services.py

from django.db import connection
from base.models import MAESTRO_HIS_ESTABLECIMIENTO, DimPeriodo


def obtener_distritos(provincia):
    """Devuelve una lista de distritos únicos filtrados por provincia."""
    return list(
        MAESTRO_HIS_ESTABLECIMIENTO.objects
        .filter(Provincia=provincia)
        .values('Distrito')
        .distinct()
        .order_by('Distrito')
    )


def obtener_avance_paquete_nino(red):
    """Ejecuta una función en la BD para obtener el avance del paquete nino."""
    with connection.cursor() as cursor:
        cursor.execute("SELECT * FROM public.obtener_avance_paquete_nino(%s)", [red])
        return cursor.fetchall()


def obtener_ranking_paquete_nino(anio, mes):
    """Ejecuta una función en la BD para obtener el ranking del paquete nino."""
    with connection.cursor() as cursor:
        cursor.execute("SELECT * FROM public.obtener_ranking_paquete_nino(%s, %s)", [anio, mes])
        return cursor.fetchall()


def obtener_seguimiento_redes_paquete_nino(red, inicio, fin):
    """Obtiene el seguimiento nominal del indicador MC-02 en un rango dado."""
    with connection.cursor() as cursor:
        cursor.execute("SELECT * FROM public.fn_seguimiento_paquete_nino(%s, %s, %s)", [red, inicio, fin])
        return cursor.fetchall()


def obtener_cobertura_paquete_nino():
    """Obtiene la cobertura del paquete nino desde una vista/tabla en la BD."""
    with connection.cursor() as cursor:
        cursor.execute(
            'SELECT * FROM public."Cobertura_MC02_PaquetenNino" '
            'ORDER BY "Red", "MicroRed", "Nombre_Establecimiento";'
        )
        return cursor.fetchall()


def obtener_avance_regional():
    """
    Obtiene el avance regional del gestante de anemia
    Supone que hay una función o campo que devuelve un valor entre 0 y 100.
    """
    with connection.cursor() as cursor:
        cursor.execute(
            'SELECT SUM(numerador_anual) AS NUM, SUM(denominador_anual) AS DEN, '
            'ROUND((SUM(numerador_anual)::NUMERIC / NULLIF(SUM(denominador_anual), 0)) * 100, 2) AS COB FROM public."Cobertura_SI_01_AnemiaGestante"; '
        )
        return cursor.fetchall()
