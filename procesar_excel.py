#!/usr/bin/env python3
"""
Sistema de procesamiento de archivos Excel para Picking Center.

Toma 2 archivos Excel:
  - Archivo 1: PROGRAMACIÓN DE PLACAS (contiene mapeo de placas genéricas a placas reales)
  - Archivo 2: PROGRAMACION PC (archivo de rutas a enriquecer)

Produce como salida el Archivo 2 modificado con:
  - Hoja RUTA: columnas Q (PLACA) y R (CONDUCTOR/OPL) rellenadas
  - Hoja1: columna K agregada (Recuento distinto de CLIENTE)
  - Hoja3: nueva hoja resumen por HORA/RUTA PC
  - Hoja8: nueva hoja pivot por HORA/PLACA/DISTRITO
"""

import sys
import os
import re
import glob
from collections import defaultdict, OrderedDict
from copy import copy
from datetime import time as dt_time

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill


def detectar_archivos(ruta_carpeta):
    """Detecta automáticamente cuál archivo es el 1 y cuál es el 2."""
    archivos_xlsx = glob.glob(os.path.join(ruta_carpeta, "*.xlsx"))
    archivos_xlsx = [f for f in archivos_xlsx if not os.path.basename(f).startswith("~$")]

    if len(archivos_xlsx) < 2:
        print("Error: Se necesitan al menos 2 archivos .xlsx en la carpeta.")
        sys.exit(1)

    archivo1 = None
    archivo2 = None

    for f in archivos_xlsx:
        nombre = os.path.basename(f).upper()
        if "RESULTADO" in nombre or "OUTPUT" in nombre:
            continue
        try:
            wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
            hojas = wb.sheetnames
            wb.close()
            if any("PROGRAMACIÓN" in h or "PROGRAMACION" in h for h in hojas):
                if "PLANIFICACION" in hojas or any("PLANIFICACION" in h for h in hojas):
                    archivo1 = f
                elif "RUTA" in hojas:
                    archivo2 = f
                else:
                    if archivo1 is None:
                        archivo1 = f
            elif "RUTA" in hojas:
                archivo2 = f
        except Exception:
            continue

    if archivo1 is None or archivo2 is None:
        print("No se pudieron identificar los archivos automáticamente.")
        print("Archivos encontrados:")
        for f in archivos_xlsx:
            print(f"  - {os.path.basename(f)}")
        sys.exit(1)

    return archivo1, archivo2


def extraer_mapeo_archivo1(ruta_archivo1):
    """
    Extrae el mapeo de PLACA GENERICA -> PLACA REAL desde el Archivo 1.
    Usa la hoja PROGRAMACIÓN, tomando solo filas donde CARGO = 'Conductor'.
    Retorna dict: {N_TRANSPORTE: PLACA_REAL}
    """
    wb = openpyxl.load_workbook(ruta_archivo1, data_only=True)

    # Buscar la hoja PROGRAMACIÓN (puede tener espacio al final)
    hoja_prog = None
    for nombre in wb.sheetnames:
        if "PROGRAMACI" in nombre.upper():
            hoja_prog = wb[nombre]
            break

    if hoja_prog is None:
        print("Error: No se encontró la hoja PROGRAMACIÓN en el Archivo 1.")
        sys.exit(1)

    # Encontrar la fila de encabezados
    header_row = None
    for r in range(1, min(10, hoja_prog.max_row + 1)):
        for c in range(1, hoja_prog.max_column + 1):
            val = hoja_prog.cell(row=r, column=c).value
            if val and "TRANSPORTE" in str(val).upper():
                header_row = r
                break
        if header_row:
            break

    if header_row is None:
        header_row = 3  # fallback

    # Mapear columnas por nombre
    col_map = {}
    for c in range(1, hoja_prog.max_column + 1):
        val = hoja_prog.cell(row=header_row, column=c).value
        if val:
            val_upper = str(val).upper().strip()
            if "TRANSPORTE" in val_upper:
                col_map["n_transporte"] = c
            elif "PLACA" in val_upper:
                col_map["placa"] = c
            elif "CARGO" in val_upper:
                col_map["cargo"] = c

    if "n_transporte" not in col_map or "placa" not in col_map:
        print("Error: No se encontraron las columnas necesarias en PROGRAMACIÓN.")
        print(f"Columnas encontradas: {col_map}")
        sys.exit(1)

    # Extraer mapeo
    mapeo = {}
    for r in range(header_row + 1, hoja_prog.max_row + 1):
        n_trans = hoja_prog.cell(row=r, column=col_map["n_transporte"]).value
        placa = hoja_prog.cell(row=r, column=col_map["placa"]).value
        cargo = hoja_prog.cell(row=r, column=col_map.get("cargo", 0)).value if "cargo" in col_map else "Conductor"

        if n_trans and placa and str(cargo).strip().lower() == "conductor":
            mapeo[str(n_trans).strip()] = str(placa).strip()

    wb.close()
    return mapeo


def derivar_opl(placa_generica):
    """Deriva el nombre corto OPL quitando dígitos finales de la PLACA GENERICA."""
    if not placa_generica:
        return ""
    return re.sub(r'\d+$', '', str(placa_generica).strip())


def reparar_pivot_cache(ruta_original, ruta_salida):
    """
    Repara el archivo de salida copiando archivos de pivot cache
    faltantes desde el archivo original. openpyxl puede perder estos
    archivos al guardar, causando errores al abrir en Excel.
    """
    import zipfile
    import shutil
    import tempfile

    with zipfile.ZipFile(ruta_original, 'r') as z_orig:
        orig_files = set(z_orig.namelist())

    with zipfile.ZipFile(ruta_salida, 'r') as z_out:
        out_files = set(z_out.namelist())

    # Encontrar archivos de pivot que faltan
    pivot_missing = [f for f in orig_files - out_files
                     if 'pivot' in f.lower()]

    if not pivot_missing:
        return

    # Copiar archivos faltantes al zip de salida
    tmp_path = ruta_salida + '.tmp'
    shutil.copy2(ruta_salida, tmp_path)

    with zipfile.ZipFile(tmp_path, 'a') as z_tmp:
        with zipfile.ZipFile(ruta_original, 'r') as z_orig:
            for fname in pivot_missing:
                z_tmp.writestr(fname, z_orig.read(fname))

    shutil.move(tmp_path, ruta_salida)
    print(f"  Reparados {len(pivot_missing)} archivos de pivot cache.")


def procesar(ruta_archivo1, ruta_archivo2, ruta_salida=None):
    """Proceso principal de transformación."""

    print(f"Archivo 1 (Placas): {os.path.basename(ruta_archivo1)}")
    print(f"Archivo 2 (Rutas):  {os.path.basename(ruta_archivo2)}")

    # 1. Extraer mapeo del Archivo 1
    mapeo_placas = extraer_mapeo_archivo1(ruta_archivo1)
    print(f"\nMapeo de placas extraído ({len(mapeo_placas)} entradas):")
    for pg, pr in sorted(mapeo_placas.items()):
        print(f"  {pg} -> {pr}")

    # 2. Cargar Archivo 2
    wb2 = openpyxl.load_workbook(ruta_archivo2)

    # =========================================================================
    # 3. Modificar hoja RUTA: llenar columnas Q (PLACA) y R (CONDUCTOR/OPL)
    # =========================================================================
    ws_ruta = wb2["RUTA"]

    # Encontrar columnas por encabezado
    col_indices = {}
    for c in range(1, ws_ruta.max_column + 1):
        val = ws_ruta.cell(row=1, column=c).value
        if val:
            col_indices[str(val).strip().upper()] = c

    col_orden = col_indices.get("ORDEN", 1)
    col_cliente = col_indices.get("CLIENTE", 5)
    col_distrito1 = col_indices.get("DISTRITO 1", 9)
    col_ruta_pc = col_indices.get("RUTA PC", 13)
    col_placa_gen = col_indices.get("PLACA GENERICA", 14)
    col_hora = col_indices.get("HORA", 15)
    col_tipo_und = col_indices.get("TIPO UND", 16)
    col_placa = col_indices.get("PLACA", 17)
    col_conductor = col_indices.get("CONDUCTOR", 18)

    # Rellenar Q y R, y recopilar datos para agregaciones
    datos_por_placa_gen = defaultdict(lambda: {
        "orders": [], "clients": set(), "hora": None, "ruta_pc": None,
        "tipo": None, "placa_real": None, "opl": None,
        "por_distrito": defaultdict(lambda: {"orders": [], "clients": set()})
    })

    for r in range(2, ws_ruta.max_row + 1):
        placa_gen = ws_ruta.cell(row=r, column=col_placa_gen).value
        if not placa_gen or str(placa_gen).strip() == "":
            continue

        placa_gen_str = str(placa_gen).strip()
        placa_real = mapeo_placas.get(placa_gen_str)
        opl = derivar_opl(placa_gen_str)

        # Rellenar columnas Q y R
        if placa_real:
            ws_ruta.cell(row=r, column=col_placa, value=placa_real)
            ws_ruta.cell(row=r, column=col_conductor, value=opl)

        # Recopilar datos para agregaciones
        orden = ws_ruta.cell(row=r, column=col_orden).value
        cliente = ws_ruta.cell(row=r, column=col_cliente).value
        hora = ws_ruta.cell(row=r, column=col_hora).value
        ruta_pc = ws_ruta.cell(row=r, column=col_ruta_pc).value
        tipo = ws_ruta.cell(row=r, column=col_tipo_und).value
        distrito = ws_ruta.cell(row=r, column=col_distrito1).value

        grupo = datos_por_placa_gen[placa_gen_str]
        grupo["orders"].append(orden)
        if cliente:
            grupo["clients"].add(cliente)
        if hora is not None:
            grupo["hora"] = hora
        if ruta_pc:
            grupo["ruta_pc"] = ruta_pc
        if tipo:
            grupo["tipo"] = tipo
        grupo["placa_real"] = placa_real or str(placa_gen_str)
        grupo["opl"] = opl

        if distrito:
            d_grupo = grupo["por_distrito"][str(distrito).strip()]
            d_grupo["orders"].append(orden)
            if cliente:
                d_grupo["clients"].add(cliente)

    print(f"\nDatos procesados: {len(datos_por_placa_gen)} grupos de placa genérica")

    # =========================================================================
    # 4. Modificar Hoja1: agregar columna K y encabezados
    # =========================================================================
    ws_hoja1 = wb2["Hoja1"]

    # Agregar encabezados de filtro
    ws_hoja1.cell(row=1, column=6, value="CONDUCTOR")
    ws_hoja1.cell(row=1, column=7, value="All")

    # Actualizar encabezado J3 y agregar K3
    ws_hoja1.cell(row=3, column=10, value="Recuento de ORDEN")
    ws_hoja1.cell(row=3, column=11, value="Recuento distinto de CLIENTE")

    # Llenar columna K para cada fila de datos
    total_clientes = 0
    for r in range(4, ws_hoja1.max_row + 1):
        placa_gen = ws_hoja1.cell(row=r, column=8).value  # H = PLACA GENERICA
        hora_label = ws_hoja1.cell(row=r, column=6).value  # F = HORA o "Total general"

        if placa_gen and str(placa_gen).strip() in datos_por_placa_gen:
            n_clientes = len(datos_por_placa_gen[str(placa_gen).strip()]["clients"])
            ws_hoja1.cell(row=r, column=11, value=n_clientes)
            total_clientes += n_clientes
        elif hora_label and "total" in str(hora_label).lower():
            ws_hoja1.cell(row=r, column=11, value=total_clientes)

    # =========================================================================
    # 5. Crear Hoja3: resumen por HORA/RUTA PC
    # =========================================================================
    # Eliminar si ya existe
    if "Hoja3" in wb2.sheetnames:
        del wb2["Hoja3"]

    ws3 = wb2.create_sheet("Hoja3", wb2.sheetnames.index("RUTA"))

    # Encabezados
    headers_h3 = ["HORA", "RUTA PC", "PLACA GENERICA", "TIPO UND",
                   "Recuento de ORDEN", "Recuento distinto de CLIENTE", "placa", "OPL"]
    for c, h in enumerate(headers_h3, 1):
        cell = ws3.cell(row=1, column=c, value=h)
        cell.font = Font(bold=True)

    # Ordenar por HORA, luego por la posición en Hoja1
    orden_hoja1 = []
    for r in range(4, ws_hoja1.max_row + 1):
        pg = ws_hoja1.cell(row=r, column=8).value
        if pg and str(pg).strip() in datos_por_placa_gen:
            orden_hoja1.append(str(pg).strip())

    # Escribir datos agrupados por HORA
    fila = 2
    hora_actual = None
    total_ordenes = 0
    total_clientes = 0

    for pg in orden_hoja1:
        grupo = datos_por_placa_gen[pg]
        hora = grupo["hora"]

        # Solo mostrar HORA cuando cambia
        if hora != hora_actual:
            ws3.cell(row=fila, column=1, value=hora)
            hora_actual = hora

        ws3.cell(row=fila, column=2, value=grupo["ruta_pc"])
        ws3.cell(row=fila, column=3, value=pg)
        ws3.cell(row=fila, column=4, value=grupo["tipo"])

        n_ord = len(grupo["orders"])
        n_cli = len(grupo["clients"])
        ws3.cell(row=fila, column=5, value=n_ord)
        ws3.cell(row=fila, column=6, value=n_cli)
        ws3.cell(row=fila, column=7, value=grupo["placa_real"])
        ws3.cell(row=fila, column=8, value=grupo["opl"])

        total_ordenes += n_ord
        total_clientes += n_cli
        fila += 1

    # Fila total
    ws3.cell(row=fila, column=1, value="Total general")
    ws3.cell(row=fila, column=1).font = Font(bold=True)
    ws3.cell(row=fila, column=5, value=total_ordenes)
    ws3.cell(row=fila, column=6, value=total_clientes)

    # Ajustar anchos
    for c in range(1, 9):
        ws3.column_dimensions[get_column_letter(c)].width = 22

    # =========================================================================
    # 6. Crear Hoja8: pivot por HORA > PLACA > DISTRITO
    # =========================================================================
    if "Hoja8" in wb2.sheetnames:
        del wb2["Hoja8"]

    ws8 = wb2.create_sheet("Hoja8", wb2.sheetnames.index("RUTA"))

    # Encabezados de filtro
    ws8.cell(row=1, column=1, value="CONDUCTOR")
    ws8.cell(row=1, column=1).font = Font(bold=True)
    ws8.cell(row=1, column=2, value="All")

    # Encabezados de tabla
    headers_h8 = ["HORA", "PLACA", "DISTRITO 1", "Recuento de ORDEN", "Recuento distinto de CLIENTE"]
    for c, h in enumerate(headers_h8, 1):
        cell = ws8.cell(row=3, column=c, value=h)
        cell.font = Font(bold=True)

    # Construir estructura: HORA -> PLACA_REAL -> DISTRITO -> (ordenes, clientes)
    pivot_data = defaultdict(lambda: defaultdict(lambda: defaultdict(
        lambda: {"orders": [], "clients": set()}
    )))

    for pg, grupo in datos_por_placa_gen.items():
        hora = grupo["hora"]
        placa_real = grupo["placa_real"]
        for distrito, d_data in grupo["por_distrito"].items():
            pivot_data[hora][placa_real][distrito]["orders"].extend(d_data["orders"])
            pivot_data[hora][placa_real][distrito]["clients"].update(d_data["clients"])

    # Función para ordenar horas (manejar None y tipos mixtos)
    def sort_hora(h):
        if h is None:
            return (999, 0)
        if isinstance(h, dt_time):
            return (h.hour, h.minute)
        try:
            return (int(h), 0)
        except (ValueError, TypeError):
            return (999, 0)

    fila = 4
    gran_total_ord = 0
    gran_total_cli = 0

    for hora in sorted(pivot_data.keys(), key=sort_hora):
        hora_total_ord = 0
        hora_total_cli = 0
        first_in_hora = True

        for placa in sorted(pivot_data[hora].keys()):
            placa_total_ord = 0
            placa_total_cli = 0
            first_in_placa = True

            for distrito in sorted(pivot_data[hora][placa].keys()):
                d_data = pivot_data[hora][placa][distrito]
                n_ord = len(d_data["orders"])
                n_cli = len(d_data["clients"])

                if first_in_hora:
                    ws8.cell(row=fila, column=1, value=hora)
                    first_in_hora = False

                if first_in_placa:
                    ws8.cell(row=fila, column=2, value=placa)
                    first_in_placa = False

                ws8.cell(row=fila, column=3, value=distrito)
                ws8.cell(row=fila, column=4, value=n_ord)
                ws8.cell(row=fila, column=5, value=n_cli)

                placa_total_ord += n_ord
                placa_total_cli += n_cli
                fila += 1

            # Subtotal por placa
            ws8.cell(row=fila, column=2, value=f"Total {placa}")
            ws8.cell(row=fila, column=2).font = Font(bold=True)
            ws8.cell(row=fila, column=4, value=placa_total_ord)
            ws8.cell(row=fila, column=5, value=placa_total_cli)
            hora_total_ord += placa_total_ord
            hora_total_cli += placa_total_cli
            fila += 1

        # Subtotal por hora
        hora_str = str(hora) if hora else ""
        if isinstance(hora, dt_time):
            hora_str = hora.strftime("%H:%M:%S")
        ws8.cell(row=fila, column=1, value=f"Total {hora_str}")
        ws8.cell(row=fila, column=1).font = Font(bold=True)
        ws8.cell(row=fila, column=4, value=hora_total_ord)
        ws8.cell(row=fila, column=5, value=hora_total_cli)
        gran_total_ord += hora_total_ord
        gran_total_cli += hora_total_cli
        fila += 1

    # Gran total
    ws8.cell(row=fila, column=1, value="Total general")
    ws8.cell(row=fila, column=1).font = Font(bold=True)
    ws8.cell(row=fila, column=4, value=gran_total_ord)
    ws8.cell(row=fila, column=5, value=gran_total_cli)

    # Ajustar anchos
    for c in range(1, 6):
        ws8.column_dimensions[get_column_letter(c)].width = 28

    # =========================================================================
    # 7. Guardar resultado y reparar pivot cache
    # =========================================================================
    if ruta_salida is None:
        base, ext = os.path.splitext(ruta_archivo2)
        ruta_salida = f"{base} - RESULTADO{ext}"

    wb2.save(ruta_salida)

    # Reparar pivot cache: openpyxl puede perder archivos de pivot tables
    # Copiamos los archivos faltantes desde el original
    reparar_pivot_cache(ruta_archivo2, ruta_salida)

    print(f"\nArchivo guardado: {ruta_salida}")
    return ruta_salida


def main():
    if len(sys.argv) == 3:
        # Modo: script archivo1.xlsx archivo2.xlsx
        archivo1 = sys.argv[1]
        archivo2 = sys.argv[2]
    elif len(sys.argv) == 2:
        # Modo: script carpeta/
        carpeta = sys.argv[1]
        archivo1, archivo2 = detectar_archivos(carpeta)
    else:
        # Modo: detectar en carpeta actual
        archivo1, archivo2 = detectar_archivos(".")

    print("=" * 60)
    print("SISTEMA DE PROCESAMIENTO EXCEL - PICKING CENTER")
    print("=" * 60)

    procesar(archivo1, archivo2)

    print("\n¡Proceso completado exitosamente!")


if __name__ == "__main__":
    main()
