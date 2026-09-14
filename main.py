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

app = FastAPI(title="Sistema de Control Logístico de Transportes")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/")
async def serve_index():
    return FileResponse("index.html")

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

def clasificar_turno_por_hora(hora_val):
    h = parse_hora(hora_val)
    if h is None:
        return "ENTRADA TURNO GENERAL"

    sec = h.hour * 3600 + h.minute * 60 + h.second

    def to_sec(h_str):
        t = datetime.strptime(h_str, "%H:%M:%S").time()
        return t.hour * 3600 + t.minute * 60 + t.second

    if sec <= to_sec("05:59:59"):
        return "ENTRADA PRIMER TURNO"
    elif sec <= to_sec("06:29:59"):
        return "SALIDA TERCER TURNO"
    elif sec <= to_sec("12:29:59"):
        return "ENTRADA TURNO GENERAL"
    elif sec <= to_sec("13:59:59"):
        return "ENTRADA SEGUNDO TURNO"
    elif sec <= to_sec("16:29:59"):
        return "SALIDA PRIMER TURNO"
    elif sec <= to_sec("17:29:59"):
        return "ENTRADA TURNO LARGO"
    elif sec <= to_sec("18:29:59"):
        return "SALIDA TURNO LARGO"
    elif sec <= to_sec("20:19:59"):
        return "SALIDA TURNO GENERAL"
    elif sec <= to_sec("21:59:59"):
        return "ENTRADA TERCER TURNO"
    else:
        return "SALIDA SEGUNDO TURNO"

def limpiar_valor(val):
    if pd.isna(val):
        return ""
    if isinstance(val, float):
        if val.is_integer():
            return str(int(val))
        return str(val)
    return str(val)

@app.post("/procesar-reporte/")
async def procesar_reporte(file: UploadFile = File(...)):
    if not file.filename.endswith('.xlsx'):
        raise HTTPException(status_code=400, detail="El archivo enviado no tiene la extensión .xlsx requerida.")

    contents = await file.read()
    try:
        df = pd.read_excel(io.BytesIO(contents))
    except Exception as e:
        raise HTTPException(status_code=400, detail="No se pudo leer el archivo Excel. Asegúrate de que no esté corrupto.")

    if df.empty:
        raise HTTPException(status_code=400, detail="El archivo subido está completamente vacío.")

    original_columns = list(df.columns)

    # Identificación inteligente de columnas
    col_fecha = next((c for c in df.columns if 'fecha' in str(c).lower()), None)
    col_hora = next((c for c in df.columns if 'hora' in str(c).lower()), None)
    col_camion = next((c for c in df.columns if 'camion' in str(c).lower() or 'unidad' in str(c).lower()), None)

    # Validaciones sobre la validez del archivo
    if not col_fecha or not col_hora or not col_camion:
        raise HTTPException(
            status_code=400, 
            detail="Estructura de archivo no válida. Faltan columnas fundamentales (Fecha, Hora o Camión/Unidad). Verifique que no esté subiendo un archivo ya procesado o incompleto."
        )

    col_linea = next((c for c in df.columns if 'linea' in str(c).lower() and 'clave' not in str(c).lower() and 'id' not in str(c).lower()), None)
    if not col_linea:
        col_linea = next((c for c in df.columns if 'linea' in str(c).lower()), None)

    col_poblacion = next((c for c in df.columns if 'poblaci' in str(c).lower()), None)

    # Procesar Fecha y Clasificación de Turnos
    df['Fecha_Parsed'] = pd.to_datetime(df[col_fecha], dayfirst=True, errors='coerce')
    
    if df['Fecha_Parsed'].dropna().empty:
        raise HTTPException(status_code=400, detail="La columna de fechas no contiene valores con formato de fecha válidos.")

    df['Turno_Clasificado'] = df[col_hora].apply(clasificar_turno_por_hora)
    df = df.sort_values(by=[col_camion, 'Fecha_Parsed', col_hora], ascending=[True, True, True])

    wb = Workbook()
    wb.remove(wb.active)

    # Estilos
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    summary_header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    global_fill = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    total_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    
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

    turnos_orden = [
        "ENTRADA PRIMER TURNO", "SALIDA TERCER TURNO", "ENTRADA TURNO GENERAL",
        "ENTRADA SEGUNDO TURNO", "SALIDA PRIMER TURNO", "ENTRADA TURNO LARGO",
        "SALIDA TURNO LARGO", "SALIDA TURNO GENERAL", "ENTRADA TERCER TURNO", 
        "SALIDA SEGUNDO TURNO"
    ]

    # --- 1. HOJA INICIAL DE RESUMEN GLOBAL ---
    ws_global = wb.create_sheet(title="RESUMEN GLOBAL", index=0)
    ws_global.views.sheetView[0].showGridLines = True

    ws_global.merge_cells("A1:E1")
    title_cell = ws_global["A1"]
    title_cell.value = "CONSOLIDADO LOGÍSTICO DE REPORTES PROCESADOS"
    title_cell.font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    title_cell.fill = global_fill
    title_cell.alignment = align_center

    global_headers = ["CAMIÓN / UNIDAD", "TOTAL REGISTROS", "LÍNEA PREDOMINANTE", "POBLACIÓN", "DÍAS OPERATIVOS"]
    ws_global.append([])
    ws_global.append(global_headers)

    for col_idx in range(1, 6):
        c = ws_global.cell(row=3, column=col_idx)
        c.fill = summary_header_fill
        c.font = font_header
        c.alignment = align_center

    g_row = 4
    total_general_registros = 0

    for camion_val, df_camion in df.groupby(col_camion, sort=True):
        camion_str = limpiar_valor(camion_val)
        cnt_total = len(df_camion)
        total_general_registros += cnt_total

        linea_str = ""
        if col_linea and not df_camion[col_linea].dropna().empty:
            linea_str = limpiar_valor(df_camion[col_linea].mode()[0])

        poblacion_str = ""
        if col_poblacion and not df_camion[col_poblacion].dropna().empty:
            poblacion_str = limpiar_valor(df_camion[col_poblacion].mode()[0])

        dias_ops = df_camion['Fecha_Parsed'].dropna().dt.date.nunique()

        ws_global.cell(row=g_row, column=1, value=camion_str).alignment = align_center
        ws_global.cell(row=g_row, column=2, value=cnt_total).alignment = align_right
        ws_global.cell(row=g_row, column=3, value=linea_str).alignment = align_left
        ws_global.cell(row=g_row, column=4, value=poblacion_str).alignment = align_left
        ws_global.cell(row=g_row, column=5, value=dias_ops).alignment = align_center

        for c_i in range(1, 6):
            cell = ws_global.cell(row=g_row, column=c_i)
            cell.font = font_regular
            cell.border = thin_border

        g_row += 1

    # Fila de Total Global
    ws_global.cell(row=g_row, column=1, value="TOTAL GENERAL").alignment = align_center
    ws_global.cell(row=g_row, column=2, value=total_general_registros).alignment = align_right
    
    for c_i in range(1, 6):
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

        for fecha_p in fechas_unicas:
            df_fecha = df_camion[df_camion['Fecha_Parsed'] == fecha_p]
            fecha_actual = fecha_p.date()
            
            start_block_row = row_pointer
            current_left_row = start_block_row

            for _, row in df_fecha.iterrows():
                for col_idx, col_name in enumerate(original_columns, 1):
                    val = row[col_name]
                    if col_name == col_fecha and pd.notna(val):
                        val_str = fecha_actual.strftime('%Y-%m-%d')
                    else:
                        val_str = limpiar_valor(val)

                    cell = ws.cell(row=current_left_row, column=col_idx, value=val_str)
                    cell.font = font_regular
                    cell.border = thin_border
                    cell.alignment = align_center if col_name in [col_fecha, col_hora, col_camion] else align_left
                
                current_left_row += 1

            linea_nombre = ""
            if col_linea and not df_fecha[col_linea].dropna().empty:
                linea_nombre = limpiar_valor(df_fecha[col_linea].dropna().iloc[0])

            poblacion_valor = ""
            if col_poblacion and not df_fecha[col_poblacion].dropna().empty:
                poblacion_valor = limpiar_valor(df_fecha[col_poblacion].dropna().iloc[0])

            conteo_turnos = df_fecha['Turno_Clasificado'].value_counts()

            c_hdr = ws.cell(row=start_block_row, column=summary_col_start, value="CONCEPTO / TURNO")
            v_hdr = ws.cell(row=start_block_row, column=summary_col_start + 1, value="VALOR / CONTEO")
            
            for h_cell in [c_hdr, v_hdr]:
                h_cell.fill = summary_header_fill
                h_cell.font = font_header
                h_cell.alignment = align_center

            summary_items = [
                ("LINEA", linea_nombre),
                ("FECHA", fecha_actual.strftime('%Y-%m-%d')),
                ("POBLACIÓN", poblacion_valor),
                ("CAMIÓN", camion_str)
            ]

            total_pasajeros = 0
            for t in turnos_orden:
                cnt = int(conteo_turnos.get(t, 0))
                summary_items.append((t, cnt))
                total_pasajeros += cnt

            summary_items.append(("TOTAL REGISTROS", total_pasajeros))

            s_row = start_block_row + 1
            for concept, val in summary_items:
                c_cell = ws.cell(row=s_row, column=summary_col_start, value=concept)
                v_cell = ws.cell(row=s_row, column=summary_col_start + 1, value=val)

                c_cell.border = thin_border
                v_cell.border = thin_border

                if concept == "TOTAL REGISTROS":
                    c_cell.fill = total_fill
                    v_cell.fill = total_fill
                    c_cell.font = font_bold
                    v_cell.font = font_bold
                    v_cell.alignment = align_right
                else:
                    c_cell.font = font_bold if s_row <= start_block_row + 4 else font_regular
                    v_cell.font = font_regular
                    v_cell.alignment = align_left if isinstance(val, str) else align_right

                s_row += 1

            max_rows = max(len(df_fecha), len(summary_items) + 1)
            row_pointer = start_block_row + max_rows + 3

        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    # Ajuste de columnas de la hoja global
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
