from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse, FileResponse
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime, time
import io
import os
from fastapi.staticfiles import StaticFiles

app = FastAPI(title="Sistema de Control Logístico de Transportes")
app.mount("/static", StaticFiles(directory="static"), name="static")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 1. Ruta para el Procesador de Excel Principal
@app.get("/")
async def serve_index():
    return FileResponse("index.html")

# 2. Ruta para el Panel Administrativo de Reportes
@app.get("/reportes")
@app.get("/reportes.html")
async def serve_reportes():
    return FileResponse("reportes.html")

# 3. Ruta para el Formulario Móvil QR
@app.get("/evaluacion")
@app.get("/evaluacion.html")
async def serve_evaluacion():
    return FileResponse("evaluacion.html")

# 4. Ruta para generar el QR
@app.get("/qr_generator")
@app.get("/qr_generator.html")
async def serve_qr_generator():
    return FileResponse("qr_generator.html")

# 5. Inicio Sesión
@app.get("/login")
def read_login():
    return FileResponse("login.html")

def parse_hora(hora_val):
    if pd.isna(hora_val):
        return None
    if isinstance(hora_val, time):
        return hora_val
    if isinstance(hora_val, datetime):
        return hora_val.time()
    
    hora_str = str(hora_val).strip()
    for fmt in ("%H:%M:%S", "%H:%M", "%I:%M:%S %p", "%I:%M %p"):
        try:
            return datetime.strptime(hora_str, fmt).time()
        except ValueError:
            pass
    return None

# --- CLASIFICACIÓN 1: TABLAS SUPERIORES (10 HORARIOS ORIGINALES) ---
def clasificar_turno_superior(hora_val):
    h = parse_hora(hora_val)
    if h is None:
        return "OTRO TURNO / FUERA DE RANGO"

    sec = h.hour * 3600 + h.minute * 60 + h.second

    if sec <= 21599:             # Hasta 05:59:59
        return "ENTRADA PRIMER TURNO"
    elif sec <= 23399:           # 06:00:00 a 06:29:59
        return "SALIDA TERCER TURNO"
    elif 23400 <= sec <= 28799:  # 06:30:00 a 07:59:59
        return "ENTRADA TURNO GENERAL"
    elif 45000 <= sec <= 50399:  # 12:30:00 a 13:59:59
        return "ENTRADA SEGUNDO TURNO"
    elif 50400 <= sec <= 59399:  # 14:00:00 a 16:29:59
        return "SALIDA PRIMER TURNO"
    elif 59400 <= sec <= 62999:  # 16:30:00 a 17:29:59
        return "ENTRADA TURNO LARGO"
    elif 63000 <= sec <= 66599:  # 17:30:00 a 18:29:59
        return "SALIDA TURNO LARGO"
    elif 66600 <= sec <= 73199:  # 18:30:00 a 20:19:59
        return "SALIDA TURNO GENERAL"
    elif 73200 <= sec <= 79199:  # 20:20:00 a 21:59:59
        return "ENTRADA TERCER TURNO"
    elif 79200 <= sec <= 84600:  # 22:00:00 a 23:30:00
        return "SALIDA SEGUNDO TURNO"
    else:
        return "OTRO TURNO / FUERA DE RANGO"

# --- CLASIFICACIÓN 2: TABLAS INFERIORES (4 HORARIOS NUEVOS INDEPENDIENTES) ---
def clasificar_turno_inferior(hora_val):
    h = parse_hora(hora_val)
    if h is None:
        return "OTRO TURNO / FUERA DE RANGO"

    sec = h.hour * 3600 + h.minute * 60 + h.second

    if 22800 <= sec <= 28799:       # 06:20:00 a 07:59:59
        return "ENTRADA TURNO GENERAL"
    elif 45000 <= sec <= 50399:    # 12:30:00 a 13:59:59
        return "ENTRADA SEGUNDO TURNO"
    elif 63000 <= sec <= 70200:    # 17:30:00 a 19:30:00
        return "SALIDA TURNO GENERAL"
    elif 77400 <= sec <= 84600:    # 21:30:00 a 23:30:00
        return "SALIDA SEGUNDO TURNO"
    else:
        return "OTRO TURNO / FUERA DE RANGO"

def limpiar_valor(val):
    if pd.isna(val):
        return ""
    if isinstance(val, float):
        if val.is_integer():
            return str(int(val))
        return str(val)
    return str(val)

@app.post("/procesar-reporte/")
@app.post("/procesar_excel")
async def procesar_reporte(file: UploadFile = File(...)):
    contents = await file.read()
    
    df = None
    try:
        df = pd.read_excel(io.BytesIO(contents), engine='openpyxl')
    except Exception:
        try:
            df = pd.read_csv(io.BytesIO(contents), sep='\t')
        except Exception:
            try:
                df = pd.read_csv(io.BytesIO(contents))
            except Exception:
                raise HTTPException(status_code=400, detail="No se pudo interpretar la estructura del archivo cargado.")

    if df is None or df.empty:
        raise HTTPException(status_code=400, detail="El archivo cargado está vacío o no contiene datos válidos.")

    df.columns = [str(col).strip() for col in df.columns]

    col_fecha = "Fecha" if "Fecha" in df.columns else next((c for c in df.columns if str(c).lower() == 'fecha'), None)
    if not col_fecha:
        col_fecha = next((c for c in df.columns if 'fecha' in str(c).lower() and 'inicio' not in str(c).lower() and 'fin' not in str(c).lower()), None)

    col_hora = next((c for c in df.columns if 'hora' in str(c).lower()), None)
    col_camion = next((c for c in df.columns if 'clavecamion' in str(c).lower() or 'camion' in str(c).lower() or 'unidad' in str(c).lower()), None)
    col_cc = next((c for c in df.columns if 'centrodecostos' in str(c).lower() or 'centro_costos' in str(c).lower()), None)
    col_total = next((c for c in df.columns if str(c).lower() == 'total'), None)

    if not col_fecha or not col_hora or not col_camion:
        raise HTTPException(
            status_code=400, 
            detail="No se encontraron las columnas requeridas (Fecha, Hora, ClaveCamion)."
        )

    col_linea = next((c for c in df.columns if 'linea' in str(c).lower() and 'clave' not in str(c).lower() and 'id' not in str(c).lower()), None)
    if not col_linea:
        col_linea = next((c for c in df.columns if 'linea' in str(c).lower()), None)

    col_poblacion = next((c for c in df.columns if 'poblaci' in str(c).lower()), None)

    if col_cc:
        df['Es_Bionatur'] = df[col_cc].astype(str).str.strip().str.endswith('D')
    else:
        df['Es_Bionatur'] = False

    df['Fecha_Clean'] = df[col_fecha].astype(str).str.split().str[0]
    df['Fecha_Parsed'] = pd.to_datetime(df['Fecha_Clean'], dayfirst=True, errors='coerce')
    
    if df['Fecha_Parsed'].dropna().empty:
        df['Fecha_Parsed'] = pd.to_datetime(df['Fecha_Clean'], errors='coerce')

    # Aplicamos ambas clasificaciones de forma independiente
    df['Turno_Superior'] = df[col_hora].apply(clasificar_turno_superior)
    df['Turno_Inferior'] = df[col_hora].apply(clasificar_turno_inferior)
    
    df = df.sort_values(by=[col_camion, 'Fecha_Parsed', col_hora], ascending=[True, True, True])

    original_columns = [c for c in df.columns if c not in ['Es_Bionatur', 'Fecha_Clean', 'Fecha_Parsed', 'Turno_Superior', 'Turno_Inferior']]

    wb = Workbook()
    wb.remove(wb.active)

    turno_color_hex = {
        "ENTRADA PRIMER TURNO": "E3F2FD",
        "SALIDA TERCER TURNO": "F3E5F5",
        "ENTRADA TURNO GENERAL": "E8F5E9",
        "ENTRADA SEGUNDO TURNO": "FFF3E0",
        "SALIDA PRIMER TURNO": "FFEBEE",
        "ENTRADA TURNO LARGO": "E0F7FA",
        "SALIDA TURNO LARGO": "FCE4EC",
        "SALIDA TURNO GENERAL": "FFFDE7",
        "ENTRADA TERCER TURNO": "E8EAF6",
        "SALIDA SEGUNDO TURNO": "EFEBE9",
        "OTRO TURNO / FUERA DE RANGO": "F5F5F5",
    }

    turno_fills = {t: PatternFill(start_color=hex_val, end_color=hex_val, fill_type="solid") for t, hex_val in turno_color_hex.items()}

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    summary_header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    bionatur_header_fill = PatternFill(start_color="1E4620", end_color="1E4620", fill_type="solid")
    global_fill = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    total_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    bionatur_total_fill = PatternFill(start_color="A5D6A7", end_color="A5D6A7", fill_type="solid")
    
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_bold = Font(name="Calibri", size=11, bold=True)
    font_regular = Font(name="Calibri", size=11)
    
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    align_center = Alignment(horizontal='center', vertical='center')
    align_left = Alignment(horizontal='left', vertical='center')
    align_right = Alignment(horizontal='right', vertical='center')

    # Lista de los 10 turnos originales para las tablas superiores
    turnos_orden_superior = [
        "ENTRADA PRIMER TURNO", "SALIDA TERCER TURNO", "ENTRADA TURNO GENERAL",
        "ENTRADA SEGUNDO TURNO", "SALIDA PRIMER TURNO", "ENTRADA TURNO LARGO",
        "SALIDA TURNO LARGO", "SALIDA TURNO GENERAL", "ENTRADA TERCER TURNO", 
        "SALIDA SEGUNDO TURNO"
    ]

    # Lista de los 4 turnos nuevos independientes para las tablas inferiores
    turnos_orden_inferior = [
        "ENTRADA TURNO GENERAL", "ENTRADA SEGUNDO TURNO",
        "SALIDA TURNO GENERAL", "SALIDA SEGUNDO TURNO"
    ]

    # --- 1. HOJA INICIAL DE RESUMEN GLOBAL ---
    ws_global = wb.create_sheet(title="RESUMEN GLOBAL", index=0)
    ws_global.views.sheetView[0].showGridLines = True

    ws_global.merge_cells("A1:F1")
    title_cell = ws_global["A1"]
    title_cell.value = "CONSOLIDADO LOGÍSTICO DE REPORTES PROCESADOS"
    title_cell.font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    title_cell.fill = global_fill
    title_cell.alignment = align_center

    global_headers = ["CAMIÓN / UNIDAD", "TOTAL REGISTROS", "TOTAL BIONATUR", "LÍNEA PREDOMINANTE", "POBLACIÓN", "DÍAS OPERATIVOS"]
    ws_global.append([])
    ws_global.append(global_headers)

    for col_idx in range(1, 7):
        c = ws_global.cell(row=3, column=col_idx)
        c.fill = summary_header_fill
        c.font = font_header
        c.alignment = align_center

    g_row = 4
    total_general_registros = 0
    total_general_bionatur = 0

    for camion_val, df_camion in df.groupby(col_camion, sort=True):
        camion_str = limpiar_valor(camion_val)
        cnt_total = len(df_camion)
        cnt_bionatur = int(df_camion['Es_Bionatur'].sum())

        total_general_registros += cnt_total
        total_general_bionatur += cnt_bionatur

        linea_str = ""
        if col_linea and not df_camion[col_linea].dropna().empty:
            linea_str = limpiar_valor(df_camion[col_linea].mode()[0])

        poblacion_str = ""
        if col_poblacion and not df_camion[col_poblacion].dropna().empty:
            poblacion_str = limpiar_valor(df_camion[col_poblacion].mode()[0])

        dias_ops = df_camion['Fecha_Parsed'].dropna().dt.date.nunique()

        ws_global.cell(row=g_row, column=1, value=camion_str).alignment = align_center
        ws_global.cell(row=g_row, column=2, value=cnt_total).alignment = align_right
        ws_global.cell(row=g_row, column=3, value=cnt_bionatur).alignment = align_right
        ws_global.cell(row=g_row, column=4, value=linea_str).alignment = align_left
        ws_global.cell(row=g_row, column=5, value=poblacion_str).alignment = align_left
        ws_global.cell(row=g_row, column=6, value=dias_ops).alignment = align_center

        for c_i in range(1, 7):
            cell = ws_global.cell(row=g_row, column=c_i)
            cell.font = font_regular
            cell.border = thin_border

        g_row += 1

    ws_global.cell(row=g_row, column=1, value="TOTAL GENERAL").alignment = align_center
    ws_global.cell(row=g_row, column=2, value=total_general_registros).alignment = align_right
    ws_global.cell(row=g_row, column=3, value=total_general_bionatur).alignment = align_right
    
    for c_i in range(1, 7):
        cell = ws_global.cell(row=g_row, column=c_i)
        cell.fill = total_fill
        cell.font = font_bold
        cell.border = thin_border

    # --- 2. GENERACIÓN DE HOJAS POR CAMIÓN ---
    for camion_val, df_camion in df.groupby(col_camion, sort=True):
        camion_str = limpiar_valor(camion_val)
        ws = wb.create_sheet(title=f"CAMION {camion_str}")
        ws.views.sheetView[0].showGridLines = True
        
        ws.append(original_columns)
        for col_num in range(1, len(original_columns) + 1):
            cell = ws.cell(row=1, column=col_num)
            cell.fill = header_fill
            cell.font = font_header
            cell.alignment = align_center

        row_pointer = 3
        fechas_unicas = sorted(df_camion['Fecha_Parsed'].dropna().unique())
        summary_col_start = len(original_columns) + 2
        bionatur_col_start = summary_col_start + 3

        for fecha_p in fechas_unicas:
            df_fecha = df_camion[df_camion['Fecha_Parsed'] == fecha_p].copy()
            fecha_actual = fecha_p.date()
            
            start_block_row = row_pointer
            current_left_row = start_block_row

            rows_data = []
            fills_data = []
            for _, row in df_fecha.iterrows():
                es_bio = row['Es_Bionatur']
                turno_nombre = row['Turno_Superior']
                row_vals = []
                for col_name in original_columns:
                    val = row[col_name]
                    if col_name == col_fecha and pd.notna(val):
                        row_vals.append(fecha_actual.strftime('%Y-%m-%d'))
                    else:
                        row_vals.append(limpiar_valor(val))
                rows_data.append(row_vals)
                fills_data.append(turno_fills.get(turno_nombre) if es_bio else None)

            for r_idx, r_vals in enumerate(rows_data):
                target_row = current_left_row + r_idx
                active_fill = fills_data[r_idx]
                for col_idx, val_str in enumerate(r_vals, 1):
                    cell = ws.cell(row=target_row, column=col_idx, value=val_str)
                    cell.font = font_regular
                    cell.border = thin_border
                    cell.alignment = align_center if original_columns[col_idx-1] in [col_fecha, col_hora, col_camion] else align_left
                    if active_fill:
                        cell.fill = active_fill

            current_left_row += len(rows_data)

            linea_nombre = ""
            if col_linea and not df_fecha[col_linea].dropna().empty:
                linea_nombre = limpiar_valor(df_fecha[col_linea].dropna().iloc[0])

            poblacion_valor = ""
            if col_poblacion and not df_fecha[col_poblacion].dropna().empty:
                poblacion_valor = limpiar_valor(df_fecha[col_poblacion].dropna().iloc[0])

            df_fecha_bionatur = df_fecha[df_fecha['Es_Bionatur'] == True].copy()
            
            if col_total:
                df_fecha[col_total] = pd.to_numeric(df_fecha[col_total], errors='coerce').fillna(0)
                df_fecha_bionatur[col_total] = pd.to_numeric(df_fecha_bionatur[col_total], errors='coerce').fillna(0)

            # --- TABLAS SUPERIORES (10 TURNOS ORIGINALES) ---
            conteo_sup_gen = df_fecha['Turno_Superior'].value_counts()
            conteo_sup_bio = df_fecha_bionatur['Turno_Superior'].value_counts()
            
            hay_fuera_sup = (
                conteo_sup_gen.get("OTRO TURNO / FUERA DE RANGO", 0) > 0 or
                conteo_sup_bio.get("OTRO TURNO / FUERA DE RANGO", 0) > 0
            )

            lista_turnos_superiores = list(turnos_orden_superior)
            if hay_fuera_sup:
                lista_turnos_superiores.append("OTRO TURNO / FUERA DE RANGO")

            c_hdr1 = ws.cell(row=start_block_row, column=summary_col_start, value="CONCEPTO / TURNO GENERAL")
            v_hdr1 = ws.cell(row=start_block_row, column=summary_col_start + 1, value="CONTEO")
            for h_cell in [c_hdr1, v_hdr1]:
                h_cell.fill = summary_header_fill
                h_cell.font = font_header
                h_cell.alignment = align_center

            c_hdr2 = ws.cell(row=start_block_row, column=bionatur_col_start, value="CONCEPTO / TURNO BIONATUR")
            v_hdr2 = ws.cell(row=start_block_row, column=bionatur_col_start + 1, value="CONTEO")
            for h_cell in [c_hdr2, v_hdr2]:
                h_cell.fill = bionatur_header_fill
                h_cell.font = font_header
                h_cell.alignment = align_center

            summary_items_gen = [
                ("LINEA", linea_nombre),
                ("FECHA", fecha_actual.strftime('%Y-%m-%d')),
                ("POBLACIÓN", poblacion_valor),
                ("CAMIÓN", camion_str)
            ]
            summary_items_bio = [
                ("LINEA", linea_nombre),
                ("FECHA", fecha_actual.strftime('%Y-%m-%d')),
                ("POBLACIÓN", poblacion_valor),
                ("CAMIÓN", camion_str)
            ]

            total_gen = 0
            total_bio = 0
            for t in lista_turnos_superiores:
                cnt_g = int(conteo_sup_gen.get(t, 0))
                cnt_b = int(conteo_sup_bio.get(t, 0))
                summary_items_gen.append((t, cnt_g))
                summary_items_bio.append((t, cnt_b))
                total_gen += cnt_g
                total_bio += cnt_b

            summary_items_gen.append(("TOTAL REGISTROS", total_gen))
            summary_items_bio.append(("TOTAL BIONATUR", total_bio))

            s_row = start_block_row + 1
            for concept, val in summary_items_gen:
                c_cell = ws.cell(row=s_row, column=summary_col_start, value=concept)
                v_cell = ws.cell(row=s_row, column=summary_col_start + 1, value=val)
                c_cell.border = thin_border
                v_cell.border = thin_border
                if concept in turno_fills:
                    c_cell.fill = turno_fills[concept]
                    v_cell.fill = turno_fills[concept]
                if concept == "TOTAL REGISTROS":
                    c_cell.fill = total_fill
                    v_cell.fill = total_fill
                    c_cell.font = font_bold
                    v_cell.font = font_bold
                    c_cell.alignment = align_right
                else:
                    c_cell.font = font_bold if s_row <= start_block_row + 4 else font_regular
                    v_cell.font = font_regular
                    c_cell.alignment = align_left if isinstance(val, str) else align_right
                    v_cell.alignment = align_right
                s_row += 1

            s_row_b = start_block_row + 1
            for concept, val in summary_items_bio:
                c_cell = ws.cell(row=s_row_b, column=bionatur_col_start, value=concept)
                v_cell = ws.cell(row=s_row_b, column=bionatur_col_start + 1, value=val)
                c_cell.border = thin_border
                v_cell.border = thin_border
                if concept in turno_fills:
                    c_cell.fill = turno_fills[concept]
                    v_cell.fill = turno_fills[concept]
                if concept == "TOTAL BIONATUR":
                    c_cell.fill = bionatur_total_fill
                    v_cell.fill = bionatur_total_fill
                    c_cell.font = font_bold
                    v_cell.font = font_bold
                    c_cell.alignment = align_right
                else:
                    c_cell.font = font_bold if s_row_b <= start_block_row + 4 else font_regular
                    v_cell.font = font_regular
                    c_cell.alignment = align_left if isinstance(val, str) else align_right
                    v_cell.alignment = align_right
                s_row_b += 1

            # --- TABLAS INFERIORES ADICIONALES (4 HORARIOS NUEVOS INDEPENDIENTES) ---
            pop_row_start = max(s_row, s_row_b) + 1
            bionatur_pop_col_start = summary_col_start + 4

            conteo_inf_gen = df_fecha['Turno_Inferior'].value_counts()
            conteo_inf_bio = df_fecha_bionatur['Turno_Inferior'].value_counts()

            hay_fuera_inf = (
                conteo_inf_gen.get("OTRO TURNO / FUERA DE RANGO", 0) > 0 or
                conteo_inf_bio.get("OTRO TURNO / FUERA DE RANGO", 0) > 0
            )

            lista_turnos_inferiores = list(turnos_orden_inferior)
            if hay_fuera_inf:
                lista_turnos_inferiores.append("OTRO TURNO / FUERA DE RANGO")

            # Encabezados Tabla General
            ws.cell(row=pop_row_start, column=summary_col_start, value="TURNO / POBLACIÓN (GENERAL)").fill = summary_header_fill
            ws.cell(row=pop_row_start, column=summary_col_start + 1, value="PERSONAS").fill = summary_header_fill
            ws.cell(row=pop_row_start, column=summary_col_start + 2, value="SUMA TOTAL").fill = summary_header_fill
            for col_i in range(summary_col_start, summary_col_start + 3):
                cell = ws.cell(row=pop_row_start, column=col_i)
                cell.font = font_header
                cell.alignment = align_center

            # Encabezados Tabla Bionatur
            ws.cell(row=pop_row_start, column=bionatur_pop_col_start, value="TURNO / POBLACIÓN (BIONATUR)").fill = bionatur_header_fill
            ws.cell(row=pop_row_start, column=bionatur_pop_col_start + 1, value="PERSONAS").fill = bionatur_header_fill
            ws.cell(row=pop_row_start, column=bionatur_pop_col_start + 2, value="SUMA TOTAL").fill = bionatur_header_fill
            for col_i in range(bionatur_pop_col_start, bionatur_pop_col_start + 3):
                cell = ws.cell(row=pop_row_start, column=col_i)
                cell.font = font_header
                cell.alignment = align_center

            r_p_gen = pop_row_start + 1
            suma_gran_total_gen = 0
            personas_gran_total_gen = 0

            for t in lista_turnos_inferiores:
                df_t = df_fecha[df_fecha['Turno_Inferior'] == t]
                if not df_t.empty:
                    c_tur = ws.cell(row=r_p_gen, column=summary_col_start, value=t)
                    ws.cell(row=r_p_gen, column=summary_col_start + 1, value="")
                    ws.cell(row=r_p_gen, column=summary_col_start + 2, value="")
                    for col_i in range(summary_col_start, summary_col_start + 3):
                        cell = ws.cell(row=r_p_gen, column=col_i)
                        cell.border = thin_border
                        if t in turno_fills:
                            cell.fill = turno_fills[t]
                    c_tur.font = font_bold
                    r_p_gen += 1

                    sub_personas = 0
                    sub_suma = 0
                    if col_poblacion and col_total:
                        grp_gen = df_t.groupby(col_poblacion).agg(
                            cnt=(col_total, 'count'),
                            tot=(col_total, 'sum')
                        ).reset_index()

                        for _, r_data in grp_gen.iterrows():
                            c1 = ws.cell(row=r_p_gen, column=summary_col_start, value=f"   {r_data[col_poblacion]}")
                            c2 = ws.cell(row=r_p_gen, column=summary_col_start + 1, value=int(r_data['cnt']))
                            c3 = ws.cell(row=r_p_gen, column=summary_col_start + 2, value=float(r_data['tot']))
                            for cell in [c1, c2, c3]:
                                cell.border = thin_border
                                cell.font = font_regular
                            c1.alignment = align_left
                            c2.alignment = align_right
                            c3.alignment = align_right
                            sub_personas += int(r_data['cnt'])
                            sub_suma += float(r_data['tot'])
                            r_p_gen += 1

                    c_sub_lbl = ws.cell(row=r_p_gen, column=summary_col_start, value=f"   SUBTOTAL {t}")
                    c_sub_cnt = ws.cell(row=r_p_gen, column=summary_col_start + 1, value=sub_personas)
                    c_sub_tot = ws.cell(row=r_p_gen, column=summary_col_start + 2, value=sub_suma)
                    for cell in [c_sub_lbl, c_sub_cnt, c_sub_tot]:
                        cell.border = thin_border
                        cell.font = font_bold
                        cell.fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
                    c_sub_lbl.alignment = align_left
                    c_sub_cnt.alignment = align_right
                    c_sub_tot.alignment = align_right
                    
                    personas_gran_total_gen += sub_personas
                    suma_gran_total_gen += sub_suma
                    r_p_gen += 1

            c1_tot = ws.cell(row=r_p_gen, column=summary_col_start, value="TOTAL GENERAL")
            c2_tot = ws.cell(row=r_p_gen, column=summary_col_start + 1, value=personas_gran_total_gen)
            c3_tot = ws.cell(row=r_p_gen, column=summary_col_start + 2, value=suma_gran_total_gen)
            for cell in [c1_tot, c2_tot, c3_tot]:
                cell.fill = total_fill
                cell.font = font_bold
                cell.border = thin_border
            c1_tot.alignment = align_right
            c2_tot.alignment = align_right
            c3_tot.alignment = align_right
            r_p_gen += 1

            r_p_bio = pop_row_start + 1
            suma_gran_total_bio = 0
            personas_gran_total_bio = 0

            for t in lista_turnos_inferiores:
                df_t_bio = df_fecha_bionatur[df_fecha_bionatur['Turno_Inferior'] == t]
                if not df_t_bio.empty:
                    c_tur = ws.cell(row=r_p_bio, column=bionatur_pop_col_start, value=t)
                    ws.cell(row=r_p_bio, column=bionatur_pop_col_start + 1, value="")
                    ws.cell(row=r_p_bio, column=bionatur_pop_col_start + 2, value="")
                    for col_i in range(bionatur_pop_col_start, bionatur_pop_col_start + 3):
                        cell = ws.cell(row=r_p_bio, column=col_i)
                        cell.border = thin_border
                        if t in turno_fills:
                            cell.fill = turno_fills[t]
                    c_tur.font = font_bold
                    r_p_bio += 1

                    sub_personas_b = 0
                    sub_suma_b = 0
                    if col_poblacion and col_total:
                        grp_bio = df_t_bio.groupby(col_poblacion).agg(
                            cnt=(col_total, 'count'),
                            tot=(col_total, 'sum')
                        ).reset_index()

                        for _, r_data in grp_bio.iterrows():
                            c1 = ws.cell(row=r_p_bio, column=bionatur_pop_col_start, value=f"   {r_data[col_poblacion]}")
                            c2 = ws.cell(row=r_p_bio, column=bionatur_pop_col_start + 1, value=int(r_data['cnt']))
                            c3 = ws.cell(row=r_p_bio, column=bionatur_pop_col_start + 2, value=float(r_data['tot']))
                            for cell in [c1, c2, c3]:
                                cell.border = thin_border
                                cell.font = font_regular
                            c1.alignment = align_left
                            c2.alignment = align_right
                            c3.alignment = align_right
                            sub_personas_b += int(r_data['cnt'])
                            sub_suma_b += float(r_data['tot'])
                            r_p_bio += 1

                    c_sub_lbl = ws.cell(row=r_p_bio, column=bionatur_pop_col_start, value=f"   SUBTOTAL {t}")
                    c_sub_cnt = ws.cell(row=r_p_bio, column=bionatur_pop_col_start + 1, value=sub_personas_b)
                    c_sub_tot = ws.cell(row=r_p_bio, column=bionatur_pop_col_start + 2, value=sub_suma_b)
                    for cell in [c_sub_lbl, c_sub_cnt, c_sub_tot]:
                        cell.border = thin_border
                        cell.font = font_bold
                        cell.fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
                    c_sub_lbl.alignment = align_left
                    c_sub_cnt.alignment = align_right
                    c_sub_tot.alignment = align_right

                    personas_gran_total_bio += sub_personas_b
                    suma_gran_total_bio += sub_suma_b
                    r_p_bio += 1

            c1_tot_b = ws.cell(row=r_p_bio, column=bionatur_pop_col_start, value="TOTAL BIONATUR")
            c2_tot_b = ws.cell(row=r_p_bio, column=bionatur_pop_col_start + 1, value=personas_gran_total_bio)
            c3_tot_b = ws.cell(row=r_p_bio, column=bionatur_pop_col_start + 2, value=suma_gran_total_bio)
            for cell in [c1_tot_b, c2_tot_b, c3_tot_b]:
                cell.fill = bionatur_total_fill
                cell.font = font_bold
                cell.border = thin_border
            c1_tot_b.alignment = align_right
            c2_tot_b.alignment = align_right
            c3_tot_b.alignment = align_right
            r_p_bio += 1

            max_pop_end = max(r_p_gen, r_p_bio)
            row_pointer = max(current_left_row, max_pop_end) + 4

        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    for col in ws_global.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws_global.column_dimensions[col_letter].width = max(max_len + 4, 15)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename=Reporte_Procesado_{datetime.now().strftime('%Y%m%d')}.xlsx"}
    )

if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)