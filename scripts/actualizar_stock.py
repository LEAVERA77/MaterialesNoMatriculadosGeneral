#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Actualiza el Excel de stock a partir de una operación de retiro / devolución.

Entrada: variable de entorno DATOS con el JSON que envía la aplicación web.
Salida:  listado_materiales_ENERSA.xlsx modificado.

Columnas de la hoja "Historial":
    A Fecha | B Usuario | C Operacion | D Material | E Cantidad | F Uso
    G Legajo | H Dispositivo | I Huella

Columnas de la hoja "Dispositivos" (se crea automáticamente):
    A ID | B Huella | C Detalle | D Ultimo usuario | E Ultimo legajo | F Ultima vez | G Operaciones

Las columnas nuevas (G, H, I) se agregan al final para no romper compatibilidad
con registros ni lecturas anteriores.

Hoja "Empleados" (PIN autogestionado por el empleado):
    A Legajo | B Apellido y nombre | C PIN (formato "SAL$HASH_SHA256") | D Activo
    E PIN actualizado | F Notas

Acciones de PIN (llegan con "accion" en el payload):
    pin_establecer  → el propio empleado define su PIN en el primer ingreso
    pin_resetear    → solo lo puede pedir un administrador de ADMIN_PINES

Acciones de legajos (solo administradores de ADMIN_PINES):
    empleado_alta | empleado_baja | empleado_activar | empleado_eliminar

Ingreso de mercadería (accion/operacion = "ingreso", SOLO administradores):
    Suma cantidades a los materiales existentes y da de alta los materiales
    nuevos al final de la hoja "Materiales". Deja constancia en "Historial"
    con la Operacion "INGRESO" y el remito/proveedor en la columna "Uso".

Hoja "Pines" (auditoría de altas y restablecimientos):
    A Fecha | B Accion | C Legajo | D Empleado | E Admin | F Dispositivo | G Huella | H Hash (8)
"""

import json
import os
import re
import sys
import unicodedata
from copy import copy
from datetime import datetime

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ARCHIVO = "listado_materiales_ENERSA.xlsx"

ENCABEZADO_HISTORIAL = [
    "Fecha", "Usuario", "Operacion", "Material", "Cantidad", "Uso",
    "Legajo", "Dispositivo", "Huella",
]
ENCABEZADO_EMPLEADOS = [
    "LEGAJO", "APELLIDO Y NOMBRE", "PIN", "ACTIVO", "PIN ACTUALIZADO", "NOTAS",
]
ENCABEZADO_PINES = [
    "Fecha", "Accion", "Legajo", "Empleado", "Administrador",
    "Dispositivo", "Huella", "Hash (primeros 8)",
]
# Administrador único de este despliegue (Sector General).
#   "todos"  → puede restablecer cualquier PIN y gestionar legajos.
#   "excepto"→ puede restablecer, salvo los legajos listados en "protegidos".
# Los legajos de esta lista son los únicos autorizados a:
#   - restablecer PINs (pin_resetear)
#   - dar de alta / editar / dar de baja / eliminar legajos (empleado_*)
ADMIN_PINES = {
    "10246": {"nombre": "VERA, LEANDRO FEDERICO", "alcance": "todos", "protegidos": ()},
}
ADMIN_LEGAJOS = tuple(ADMIN_PINES.keys())


def puede_resetear(admin, legajo):
    """Misma regla que aplica la web, pero acá no se puede saltear."""
    info = ADMIN_PINES.get(admin)
    if not info:
        return False
    if info["alcance"] == "todos":
        return True
    return legajo not in info["protegidos"]
ENCABEZADO_DISPOSITIVOS = [
    "ID Dispositivo", "Huella", "Detalle", "Ultimo usuario", "Ultimo legajo",
    "Ultima operacion", "Operaciones",
]


def log(msg):
    print(msg, flush=True)


def fatal(msg):
    log("ERROR FATAL: " + msg)
    sys.exit(1)


def valor(datos, clave, default=""):
    """Lectura tolerante: los payloads viejos no traen legajo ni dispositivo."""
    v = datos.get(clave, default)
    return default if v is None else v


# GitHub limita el client_payload del repository_dispatch a 10 propiedades,
# por eso los datos del equipo pueden venir planos (dispositivo_id) o anidados
# en un objeto "dispositivo" ({id, huella, resumen, ua, zona, timestamp}).
CAMPOS_DISPOSITIVO = {
    "id": ("dispositivo_id", "id"),
    "huella": ("dispositivo_huella", "huella"),
    "resumen": ("dispositivo_resumen", "resumen"),
    "ua": ("dispositivo_ua", "ua"),
    "zona": ("zona_horaria", "zona"),
    "timestamp": ("timestamp_cliente", "timestamp"),
}


def dato_dispositivo(datos, campo, default=""):
    plano, anidado = CAMPOS_DISPOSITIVO[campo]
    v = datos.get(plano)
    if v not in (None, ""):
        return v
    disp = datos.get("dispositivo")
    if isinstance(disp, dict):
        v = disp.get(anidado)
        if v not in (None, ""):
            return v
    return default


def cargar_datos():
    raw = os.environ.get("DATOS", "").strip()
    if not raw:
        fatal("No se recibieron datos (variable DATOS vacía).")
    try:
        return json.loads(raw)
    except json.JSONDecodeError as e:
        fatal(f"JSON inválido: {e}")


def validar_legajo(wb, datos):
    """Si existe la hoja 'Empleados', el legajo debe existir y estar activo.
    De esta forma la validación no depende solo del navegador."""
    if "Empleados" not in wb.sheetnames:
        log("Aviso: no existe la hoja 'Empleados'; no se valida el legajo en el servidor.")
        return

    legajo = str(valor(datos, "legajo", "")).strip()
    if not legajo:
        # Operación registrada antes de la migración: se acepta, queda sin legajo.
        log("Aviso: la operación no informa legajo (cliente antiguo). Se registra igual.")
        return

    ws = wb["Empleados"]
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row and row[0] is not None and str(row[0]).strip() == legajo:
            activo = str(row[3]).strip().upper() if len(row) > 3 and row[3] is not None else "SI"
            if activo in ("NO", "FALSE", "0", "INACTIVO"):
                fatal(f"El legajo {legajo} está inactivo en la hoja Empleados.")
            return
    fatal(f"El legajo {legajo} no existe en la hoja Empleados.")


def validar_stock(ws, datos):
    if datos.get("operacion") != "retiro":
        return
    solicitado = {}
    for item in datos.get("items", []):
        nombre = str(item.get("material", "")).strip().lower()
        solicitado[nombre] = solicitado.get(nombre, 0.0) + float(item.get("cantidad", 0))

    for nombre, cantidad in solicitado.items():
        encontrado = False
        for row in ws.iter_rows(min_row=3):
            nombre_excel = str(row[0].value).strip().lower() if row[0].value else ""
            if nombre_excel == nombre:
                actual = float(row[3].value) if row[3].value not in (None, "") else 0.0
                if cantidad > actual:
                    fatal(
                        f"Stock insuficiente para {nombre}: "
                        f"solicitado {cantidad}, disponible {actual}"
                    )
                encontrado = True
                break
        if not encontrado:
            fatal(f"Material no encontrado en el Excel: {nombre}")


def aplicar_movimientos(ws, datos):
    operacion = datos.get("operacion")
    for item in datos.get("items", []):
        nombre = str(item.get("material", "")).strip().lower()
        cantidad = float(item.get("cantidad", 0))
        encontrado = False
        for row in ws.iter_rows(min_row=3):
            nombre_excel = str(row[0].value).strip().lower() if row[0].value else ""
            if nombre_excel == nombre:
                actual = float(row[3].value) if row[3].value not in (None, "") else 0.0
                row[3].value = actual - cantidad if operacion == "retiro" else actual + cantidad
                encontrado = True
                break
        if encontrado:
            log(f"  {operacion}: {item.get('material')} x {cantidad}")
        else:
            log(f"  Aviso: no se encontró en el stock: {item.get('material')}")



# ==========================================================
# INGRESO DE MERCADERÍA (solo administradores)
# ==========================================================
# La operación "ingreso" es la única que puede AGREGAR materiales nuevos a la
# planilla y la única que suma cantidades. Por eso exige un administrador de
# ADMIN_PINES, igual que la gestión de legajos.
MAX_LINEAS_INGRESO = 50
MAX_NOMBRE_INGRESO = 120
MAX_CANTIDAD_INGRESO = 1_000_000


def es_ingreso(datos):
    return (str(datos.get("accion", "")).strip().lower() == "ingreso" or
            str(datos.get("operacion", "")).strip().lower() == "ingreso")


def validar_admin_ingreso(datos, wb=None):
    """Solo un legajo de ADMIN_PINES puede cargar mercadería.

    Se exige que el legajo que opera (el de la sesión) sea administrador. Si
    además llega admin_legajo, tiene que coincidir: así nadie puede enviar un
    legajo común declarando el de otro para saltear el control.
    """
    actor = str(datos.get("legajo") or "").strip()
    admin = str(datos.get("admin_legajo") or "").strip()
    if admin and actor and admin != actor:
        fatal("Payload inconsistente: el legajo que opera no coincide con el "
              "administrador informado. No se registró nada.")
    quien = actor or admin
    autorizado = es_admin_en_wb(wb, quien) if wb is not None else quien in ADMIN_LEGAJOS
    if not autorizado:
        fatal("Solo un administrador autorizado puede registrar ingresos de mercadería "
              f"(legajo informado: {quien or 'ninguno'}).")
    return quien


def normalizar_material(nombre):
    """Clave de comparación tolerante: sin mayúsculas, acentos ni espacios de más.
    Evita que "Alambre  de Fardo" cree un material duplicado."""
    s = unicodedata.normalize("NFKD", str(nombre or ""))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", s).strip().lower()


def validar_items_ingreso(datos):
    """Normaliza y valida las líneas ANTES de tocar el Excel: si algo está mal,
    no se escribe nada (todo o nada)."""
    items = datos.get("items") or []
    if not items:
        fatal("El ingreso no tiene materiales.")
    if len(items) > MAX_LINEAS_INGRESO:
        fatal(f"El ingreso tiene {len(items)} líneas y el máximo es {MAX_LINEAS_INGRESO}.")

    limpios = []
    for n, item in enumerate(items, start=1):
        if not isinstance(item, dict):
            fatal(f"Línea {n}: formato inválido.")
        nombre = " ".join(str(item.get("material", "")).split())
        if len(nombre) < 2:
            fatal(f"Línea {n}: falta la descripción del material.")
        if len(nombre) > MAX_NOMBRE_INGRESO:
            fatal(f"Línea {n}: la descripción supera los {MAX_NOMBRE_INGRESO} caracteres.")
        try:
            cantidad = float(str(item.get("cantidad", "")).replace(",", "."))
        except (TypeError, ValueError):
            fatal(f"Línea {n}: la cantidad no es un número válido.")
        if cantidad <= 0:
            fatal(f"Línea {n}: la cantidad debe ser mayor que cero.")
        if cantidad > MAX_CANTIDAD_INGRESO:
            fatal(f"Línea {n}: la cantidad supera el máximo permitido ({MAX_CANTIDAD_INGRESO}).")

        limpios.append({
            "material": nombre,
            "especificacion": " ".join(str(item.get("especificacion", "") or "").split()),
            "unidad": str(item.get("unidad", "") or "").strip() or "u.",
            "cantidad": cantidad,
        })
    return limpios


def hoja_materiales(wb):
    return wb["Materiales"] if "Materiales" in wb.sheetnames else wb.worksheets[0]


def indice_materiales(ws):
    """{clave normalizada: fila} para ubicar cada material una sola vez."""
    indice = {}
    for fila in range(3, ws.max_row + 1):
        valor_a = ws.cell(row=fila, column=1).value
        clave = normalizar_material(valor_a)
        if clave and clave not in indice:
            indice[clave] = fila
    return indice


def copiar_formato_fila(ws, fila_origen, fila_destino):
    """La fila nueva hereda el formato de la última fila cargada (bordes,
    alineación, tipo de letra y formato de número)."""
    for columna in range(1, 5):
        origen = ws.cell(row=fila_origen, column=columna)
        destino = ws.cell(row=fila_destino, column=columna)
        if origen.has_style:
            destino._style = copy(origen._style)


def aplicar_ingresos(wb, items):
    ws = hoja_materiales(wb)
    indice = indice_materiales(ws)
    ultima_fila_datos = ws.max_row

    sumados = 0
    nuevos = 0
    for item in items:
        clave = normalizar_material(item["material"])
        if clave in indice:
            celda = ws.cell(row=indice[clave], column=4)
            anterior = float(celda.value) if celda.value not in (None, "") else 0.0
            celda.value = anterior + item["cantidad"]
            sumados += 1
            log(f"  ingreso: {item['material']} +{item['cantidad']:g} → {celda.value:g} "
                f"(antes {anterior:g})")
            continue

        # Material nuevo: se agrega al final de la planilla.
        fila = ws.max_row + 1
        ws.cell(row=fila, column=1, value=item["material"])
        ws.cell(row=fila, column=2, value=item["especificacion"] or None)
        ws.cell(row=fila, column=3, value=item["unidad"])
        ws.cell(row=fila, column=4, value=item["cantidad"])
        copiar_formato_fila(ws, ultima_fila_datos, fila)
        indice[clave] = fila
        ultima_fila_datos = fila
        nuevos += 1
        log(f"  ingreso: NUEVO {item['material']} · {item['unidad']} · {item['cantidad']:g}")

    return {"nuevos": nuevos, "sumados": sumados, "total": len(items)}


def obtener_historial(wb):
    if "Historial" in wb.sheetnames:
        ws = wb["Historial"]
    else:
        ws = wb.create_sheet("Historial")
        ws.append(ENCABEZADO_HISTORIAL)
        return ws

    # Completar/crear encabezados nuevos sin borrar los datos existentes
    fila1 = [c.value for c in ws[1]] if ws.max_row >= 1 else []
    if not any(fila1):
        for i, titulo in enumerate(ENCABEZADO_HISTORIAL, start=1):
            ws.cell(row=1, column=i, value=titulo)
    else:
        for i, titulo in enumerate(ENCABEZADO_HISTORIAL, start=1):
            actual = ws.cell(row=1, column=i).value
            if actual in (None, ""):
                ws.cell(row=1, column=i, value=titulo)
    return ws


def registrar_historial(wb, datos):
    ws = obtener_historial(wb)
    fecha = valor(datos, "fecha", datetime.now().strftime("%Y-%m-%d %H:%M"))
    usuario = valor(datos, "usuario", "")
    operacion = str(valor(datos, "operacion", "")).upper()
    uso = valor(datos, "uso", "")
    legajo = str(valor(datos, "legajo", "")).strip()
    dispositivo = dato_dispositivo(datos, "id", "")
    huella = dato_dispositivo(datos, "huella", "")

    for item in datos.get("items", []):
        ws.append([
            fecha, usuario, operacion, item.get("material", ""),
            item.get("cantidad", 0), uso, legajo, dispositivo, huella,
        ])


def registrar_dispositivo(wb, datos):
    """Mantiene una planilla de equipos vistos, para auditar la trazabilidad."""
    dispositivo = str(dato_dispositivo(datos, "id", "")).strip()
    if not dispositivo:
        return

    if "Dispositivos" in wb.sheetnames:
        ws = wb["Dispositivos"]
    else:
        ws = wb.create_sheet("Dispositivos")
        ws.append(ENCABEZADO_DISPOSITIVOS)
        for i in range(1, len(ENCABEZADO_DISPOSITIVOS) + 1):
            ws.column_dimensions[get_column_letter(i)].width = 22
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="0B3B2E")

    fila = None
    for r in ws.iter_rows(min_row=2):
        if r[0].value and str(r[0].value).strip() == dispositivo:
            fila = r
            break

    huella = dato_dispositivo(datos, "huella", "")
    detalle = dato_dispositivo(datos, "resumen", "") or dato_dispositivo(datos, "ua", "")
    usuario = valor(datos, "usuario", "")
    legajo = str(valor(datos, "legajo", "")).strip()
    operacion = str(valor(datos, "operacion", "")).upper()
    fecha = dato_dispositivo(datos, "timestamp", "") or valor(datos, "fecha", "")

    if fila is None:
        ws.append([dispositivo, huella, detalle, usuario, legajo, f"{fecha} · {operacion}", 1])
    else:
        r = fila[0].row
        operaciones = ws.cell(row=r, column=7).value or 0
        try:
            operaciones = int(operaciones) + 1
        except (TypeError, ValueError):
            operaciones = 1
        ws.cell(row=r, column=1, value=dispositivo)
        if huella:
            ws.cell(row=r, column=2, value=huella)
        if detalle:
            ws.cell(row=r, column=3, value=detalle)
        ws.cell(row=r, column=4, value=usuario)
        ws.cell(row=r, column=5, value=legajo)
        ws.cell(row=r, column=6, value=f"{fecha} · {operacion}")
        ws.cell(row=r, column=7, value=operaciones)


def es_accion_pin(datos):
    return str(datos.get("accion", "")).strip().lower() in ("pin_establecer", "pin_resetear")


def hoja_empleados(wb):
    if "Empleados" not in wb.sheetnames:
        fatal("No existe la hoja 'Empleados' en el Excel.")
    ws = wb["Empleados"]
    for i, titulo in enumerate(ENCABEZADO_EMPLEADOS, start=1):
        if ws.cell(row=1, column=i).value in (None, ""):
            ws.cell(row=1, column=i, value=titulo)
    for c in ws[1]:
        if c.value:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="0B3B2E")
    return ws


def fila_empleado(ws, legajo):
    for row in ws.iter_rows(min_row=2, values_only=False):
        if row and row[0].value is not None and str(row[0].value).strip() == legajo:
            return row
    return None


def hoja_pines(wb):
    if "Pines" in wb.sheetnames:
        return wb["Pines"]
    ws = wb.create_sheet("Pines")
    ws.append(ENCABEZADO_PINES)
    for i in range(1, len(ENCABEZADO_PINES) + 1):
        ws.column_dimensions[get_column_letter(i)].width = 20
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="0B3B2E")
    return ws


def auditar_pin(wb, datos, legajo, nombre, accion, hash_pin):
    ws = hoja_pines(wb)
    ws.append([
        valor(datos, "fecha", datetime.now().strftime("%Y-%m-%d %H:%M")),
        accion,
        legajo,
        nombre,
        str(valor(datos, "admin_legajo", "")).strip() or "—",
        str(dato_dispositivo(datos, "id", "")).strip() or "—",
        str(dato_dispositivo(datos, "huella", "")).strip() or "—",
        (hash_pin or "")[:8],
    ])


def establecer_pin(wb, datos):
    """El empleado define su PIN: se guarda SAL$HASH (nunca el PIN en claro)."""
    legajo = str(valor(datos, "legajo", "")).strip()
    sal = str(valor(datos, "sal", "")).strip()
    hash_pin = str(valor(datos, "pin_hash", "")).strip().lower()

    if not legajo:
        fatal("El alta de PIN no informa legajo.")
    if not re.fullmatch(r"[0-9a-f]{8,32}", sal):
        fatal(f"Sal inválida para el legajo {legajo}.")
    if not re.fullmatch(r"[0-9a-f]{64}", hash_pin):
        fatal(f"Hash de PIN inválido para el legajo {legajo}.")

    ws = hoja_empleados(wb)
    fila = fila_empleado(ws, legajo)
    if fila is None:
        fatal(f"El legajo {legajo} no existe en la hoja Empleados.")

    activo = str(fila[3].value or "SI").strip().upper()
    if activo in ("NO", "FALSE", "0", "INACTIVO"):
        fatal(f"El legajo {legajo} está inactivo.")

    fila[2].value = f"{sal}${hash_pin}"          # C: PIN (sal + hash)
    fila[4].value = valor(datos, "fecha", datetime.now().strftime("%Y-%m-%d %H:%M"))
    nombre = str(fila[1].value or "").strip()

    auditar_pin(wb, datos, legajo, nombre, "pin_establecer", hash_pin)
    log(f"  PIN establecido para {legajo} · {nombre}")


def resetear_pin(wb, datos):
    """Solo un administrador puede dejar el PIN vacío para que el empleado cree otro."""
    admin = str(valor(datos, "admin_legajo", "")).strip()
    legajo = str(valor(datos, "legajo", "")).strip()

    if not es_admin_en_wb(wb, admin):
        fatal(f"El legajo {admin or '(vacío)'} no está autorizado para restablecer PIN.")
    if not legajo:
        fatal("El restablecimiento de PIN no informa el legajo del empleado.")
    if admin in ADMIN_PINES and not puede_resetear(admin, legajo):
        fatal(
            f"El administrador {admin} ({ADMIN_PINES[admin]['nombre']}) no tiene alcance "
            f"para restablecer el PIN del legajo {legajo}."
        )
    if legajo == admin:
        log("  Aviso: el administrador restablece su propio PIN.")

    ws = hoja_empleados(wb)
    fila = fila_empleado(ws, legajo)
    if fila is None:
        fatal(f"El legajo {legajo} no existe en la hoja Empleados.")

    fila[2].value = ""                            # C: sin PIN → deberá crear uno nuevo
    fila[4].value = f"{valor(datos, 'fecha', datetime.now().strftime('%Y-%m-%d %H:%M'))} · reset por {admin}"
    nombre = str(fila[1].value or "").strip()

    auditar_pin(wb, datos, legajo, nombre, "pin_resetear", "")
    log(f"  PIN restablecido por {admin} para {legajo} · {nombre}")


# ----------------------------------------------------------------------
# Gestión de legajos autorizados (solo administradores de ADMIN_LEGAJOS)
#   empleado_alta      → crea el legajo o actualiza el apellido y nombre
#   empleado_baja      → ACTIVO = NO (deja de poder operar y de crear PIN)
#   empleado_activar   → ACTIVO = SI
#   empleado_eliminar  → borra la fila de la hoja Empleados
# El PIN nunca se toca acá: se crea con pin_establecer o se limpia con pin_resetear.
# ----------------------------------------------------------------------
ACCIONES_EMPLEADO = (
    "empleado_alta", "empleado_baja", "empleado_activar", "empleado_eliminar",
    "empleado_admin_agregar", "empleado_admin_quitar",
)


def es_accion_empleado(datos):
    return str(datos.get("accion", "")).strip().lower() in ACCIONES_EMPLEADO


def legajo_valido(legajo):
    return bool(re.fullmatch(r"[0-9A-Za-z]{2,12}", legajo or ""))




ADMIN_ROLE_HEADER = "ADMINISTRADOR"

def indice_columna_admin(ws):
    for c in range(1, ws.max_column + 1):
        if str(ws.cell(1, c).value or "").strip().upper() == ADMIN_ROLE_HEADER:
            return c
    return None

def es_admin_en_wb(wb, legajo):
    legajo = str(legajo or "").strip()
    if "Empleados" not in wb.sheetnames:
        return legajo in ADMIN_LEGAJOS
    ws = wb["Empleados"]
    col = indice_columna_admin(ws)
    if col is None:
        return legajo in ADMIN_LEGAJOS  # compatibilidad inicial, antes de activar roles dinámicos
    fila = fila_empleado(ws, legajo)
    if fila is None:
        return False
    activo = str(fila[3].value or "SI").strip().upper()
    rol = str(fila[col-1].value or "").strip().upper()
    return activo not in ("NO", "FALSE", "0", "INACTIVO") and rol in ("SI", "SÍ", "ADMIN", "ADMINISTRADOR", "TRUE", "1")

def asegurar_columna_admin(ws):
    col = indice_columna_admin(ws)
    if col is not None:
        return col
    col = ws.max_column + 1
    ws.cell(1, col, ADMIN_ROLE_HEADER)
    ws.cell(1, col)._style = ws.cell(1, max(1, col-1))._style
    # Migración segura: conserva el administrador inicial como admin hasta que
    # se retire expresamente después de promover a otra persona.
    for legajo in ADMIN_LEGAJOS:
        fila = fila_empleado(ws, legajo)
        if fila is not None:
            ws.cell(fila[0].row, col, "SI")
    return col

def admins_activos(wb):
    if "Empleados" not in wb.sheetnames:
        return list(ADMIN_LEGAJOS)
    ws = wb["Empleados"]
    col = indice_columna_admin(ws)
    if col is None:
        return [a for a in ADMIN_LEGAJOS if fila_empleado(ws, a) is not None]
    return [str(ws.cell(r,1).value).strip() for r in range(2,ws.max_row+1)
            if es_admin_en_wb(wb, str(ws.cell(r,1).value or "").strip())]


def gestionar_empleado(wb, datos):
    accion = str(datos.get("accion", "")).strip().lower()
    admin = str(valor(datos, "admin_legajo", "")).strip()
    legajo = str(valor(datos, "legajo", "")).strip()
    nombre = str(valor(datos, "nombre", "")).strip().upper()
    fecha = valor(datos, "fecha", datetime.now().strftime("%Y-%m-%d %H:%M"))

    if not es_admin_en_wb(wb, admin):
        fatal(f"El legajo {admin or '(vacío)'} no está autorizado para administrar legajos.")
    if not legajo_valido(legajo):
        fatal(f"Legajo inválido: {legajo!r} (se esperan entre 2 y 12 caracteres alfanuméricos).")

    ws = hoja_empleados(wb)
    fila = fila_empleado(ws, legajo)

    if accion == "empleado_alta":
        if not nombre:
            fatal("El alta de legajo necesita el apellido y nombre.")
        if fila is None:
            ws.append([legajo, nombre, "", "SI", fecha, f"Alta por {admin}"])
            log(f"  Legajo {legajo} · {nombre} creado por {admin}")
        else:
            anterior = str(fila[1].value or "").strip()
            fila[1].value = nombre
            if not str(fila[3].value or "").strip():
                fila[3].value = "SI"
            log(f"  Legajo {legajo} actualizado por {admin}: {anterior!r} → {nombre!r}")

    elif accion in ("empleado_admin_agregar", "empleado_admin_quitar"):
        if fila is None:
            fatal(f"El legajo {legajo} no existe en la hoja Empleados.")
        col = asegurar_columna_admin(ws)
        activo = str(fila[3].value or "SI").strip().upper()
        if accion == "empleado_admin_agregar":
            if activo in ("NO", "FALSE", "0", "INACTIVO"):
                fatal("No se puede promover a un empleado inactivo; reactivá el legajo primero.")
            ws.cell(fila[0].row, col).value = "SI"
        else:
            actuales = admins_activos(wb)
            if legajo not in actuales:
                fatal(f"El legajo {legajo} no tiene rol de administrador.")
            if len(actuales) <= 1:
                fatal("No se puede quitar el último administrador activo. Promové a otra persona primero.")
            ws.cell(fila[0].row, col).value = "NO"
        nombre = str(fila[1].value or "").strip()
        log(f"  Rol administrador {'asignado' if accion == 'empleado_admin_agregar' else 'quitado'}: {legajo} · {nombre} (por {admin})")

    elif accion in ("empleado_baja", "empleado_activar"):
        if fila is None:
            fatal(f"El legajo {legajo} no existe en la hoja Empleados.")
        if legajo == admin and accion == "empleado_baja":
            fatal("Un administrador no puede darse de baja a sí mismo.")
        if es_admin_en_wb(wb, legajo) and accion == "empleado_baja":
            fatal(f"El legajo {legajo} es administrador: quitale primero el rol de administrador.")
        fila[3].value = "NO" if accion == "empleado_baja" else "SI"
        nombre = str(fila[1].value or "").strip()   # para que la auditoría lo registre
        log(f"  Legajo {legajo} → {'INACTIVO' if accion == 'empleado_baja' else 'ACTIVO'} (por {admin})")

    elif accion == "empleado_eliminar":
        if fila is None:
            fatal(f"El legajo {legajo} no existe en la hoja Empleados.")
        if es_admin_en_wb(wb, legajo):
            fatal(f"El legajo {legajo} es administrador: quitale primero el rol de administrador.")
        nombre = str(fila[1].value or "").strip()
        ws.delete_rows(fila[0].row, 1)
        log(f"  Legajo {legajo} · {nombre} eliminado por {admin}")

    auditar_pin(wb, datos, legajo, nombre, accion, "")


def main():
    datos = cargar_datos()
    log("--- DATOS RECIBIDOS ---")
    log(json.dumps(datos, indent=2, ensure_ascii=False))
    log("-----------------------")

    if not os.path.exists(ARCHIVO):
        fatal(f"No se encontró {ARCHIVO}")

    # --- Gestión de legajos (no toca el stock ni los PIN) ---
    if es_accion_empleado(datos):
        wb = load_workbook(ARCHIVO)
        gestionar_empleado(wb, datos)
        wb.save(ARCHIVO)
        log("Archivo Excel guardado exitosamente (gestión de legajo).")
        return

    # --- Acciones de PIN (no tocan el stock) ---
    if es_accion_pin(datos):
        wb = load_workbook(ARCHIVO)
        accion = str(datos.get("accion", "")).strip().lower()
        if accion == "pin_establecer":
            establecer_pin(wb, datos)
        else:
            resetear_pin(wb, datos)
        wb.save(ARCHIVO)
        log("Archivo Excel guardado exitosamente (acción de PIN).")
        return

    # --- Ingreso de mercadería (solo administradores; suma o da de alta) ---
    if es_ingreso(datos):
        wb = load_workbook(ARCHIVO)
        validar_legajo(wb, datos)
        validar_admin_ingreso(datos, wb)
        items = validar_items_ingreso(datos)
        # El historial registra los nombres/cantidades ya normalizados.
        datos["items"] = items
        resumen = aplicar_ingresos(wb, items)
        registrar_historial(wb, datos)
        registrar_dispositivo(wb, datos)
        wb.save(ARCHIVO)
        log(f"Ingreso aplicado: {resumen['sumados']} material(es) sumado(s), "
            f"{resumen['nuevos']} material(es) nuevo(s), {resumen['total']} línea(s).")
        log("Archivo Excel guardado exitosamente (ingreso de mercadería).")
        return

    if str(datos.get("operacion", "")).lower() not in ("retiro", "devolucion"):
        fatal("Operación inválida (se esperaba 'retiro', 'devolucion' o 'ingreso').")
    if not datos.get("items"):
        fatal("La operación no tiene materiales.")

    wb = load_workbook(ARCHIVO)
    ws = wb["Materiales"] if "Materiales" in wb.sheetnames else wb.worksheets[0]

    validar_legajo(wb, datos)
    validar_stock(ws, datos)
    aplicar_movimientos(ws, datos)
    registrar_historial(wb, datos)
    registrar_dispositivo(wb, datos)

    wb.save(ARCHIVO)
    log("Archivo Excel guardado exitosamente.")


if __name__ == "__main__":
    main()
