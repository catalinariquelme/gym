"""
Genera un archivo Excel (.xlsm) con macros VBA para rastrear
entrenamientos en casa: rutinas por día, series, repeticiones y peso.
"""

import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from copy import copy

# ── Colores y estilos ─────────────────────────────────────────────
AZUL_OSCURO  = "1B2A4A"
AZUL_MEDIO   = "2E5090"
AZUL_CLARO   = "D6E4F0"
VERDE        = "27AE60"
VERDE_CLARO  = "D5F5E3"
NARANJA      = "E67E22"
NARANJA_CL   = "FDEBD0"
GRIS_CLARO   = "F2F2F2"
BLANCO       = "FFFFFF"
ROJO         = "E74C3C"

font_titulo   = Font(name="Calibri", size=16, bold=True, color=BLANCO)
font_subtit   = Font(name="Calibri", size=12, bold=True, color=BLANCO)
font_header   = Font(name="Calibri", size=11, bold=True, color=BLANCO)
font_normal   = Font(name="Calibri", size=11, color="333333")
font_bold     = Font(name="Calibri", size=11, bold=True, color="333333")
font_small    = Font(name="Calibri", size=10, color="666666")
font_btn      = Font(name="Calibri", size=11, bold=True, color=BLANCO)

fill_titulo   = PatternFill("solid", fgColor=AZUL_OSCURO)
fill_header   = PatternFill("solid", fgColor=AZUL_MEDIO)
fill_alt1     = PatternFill("solid", fgColor=BLANCO)
fill_alt2     = PatternFill("solid", fgColor=GRIS_CLARO)
fill_input    = PatternFill("solid", fgColor=AZUL_CLARO)
fill_verde    = PatternFill("solid", fgColor=VERDE)
fill_verde_cl = PatternFill("solid", fgColor=VERDE_CLARO)
fill_naranja  = PatternFill("solid", fgColor=NARANJA)
fill_naranja_cl = PatternFill("solid", fgColor=NARANJA_CL)
fill_rojo     = PatternFill("solid", fgColor=ROJO)

align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
align_left   = Alignment(horizontal="left", vertical="center", wrap_text=True)

thin_border = Border(
    left=Side(style="thin", color="CCCCCC"),
    right=Side(style="thin", color="CCCCCC"),
    top=Side(style="thin", color="CCCCCC"),
    bottom=Side(style="thin", color="CCCCCC"),
)


def set_cell(ws, row, col, value, font=font_normal, fill=None, alignment=align_center, border=thin_border):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = font
    if fill:
        cell.fill = fill
    cell.alignment = alignment
    cell.border = border
    return cell


def merge_and_set(ws, start_row, start_col, end_row, end_col, value, font, fill, alignment=align_center):
    ws.merge_cells(
        start_row=start_row, start_column=start_col,
        end_row=end_row, end_column=end_col,
    )
    cell = ws.cell(row=start_row, column=start_col, value=value)
    cell.font = font
    cell.fill = fill
    cell.alignment = alignment
    cell.border = thin_border
    # Apply border/fill to all merged cells
    for r in range(start_row, end_row + 1):
        for c in range(start_col, end_col + 1):
            ws.cell(row=r, column=c).border = thin_border
            ws.cell(row=r, column=c).fill = fill


# ── Rutinas predefinidas para casa ────────────────────────────────
RUTINAS = {
    "Lunes - Pecho y Tríceps": [
        "Flexiones clásicas",
        "Flexiones diamante",
        "Flexiones declinadas",
        "Fondos en silla",
        "Flexiones abiertas",
        "Extensión tríceps con mancuerna",
    ],
    "Martes - Espalda y Bíceps": [
        "Remo con mancuernas",
        "Dominadas (o banda elástica)",
        "Remo invertido (mesa)",
        "Curl bíceps con mancuerna",
        "Curl martillo",
        "Superman",
    ],
    "Miércoles - Piernas": [
        "Sentadillas",
        "Sentadilla búlgara",
        "Zancadas",
        "Peso muerto rumano",
        "Elevación de talones",
        "Puente de glúteos",
    ],
    "Jueves - Hombros y Core": [
        "Press militar con mancuernas",
        "Elevaciones laterales",
        "Elevaciones frontales",
        "Plancha frontal (seg)",
        "Plancha lateral (seg)",
        "Crunch abdominal",
    ],
    "Viernes - Full Body": [
        "Burpees",
        "Sentadilla con salto",
        "Flexiones",
        "Remo con mancuernas",
        "Zancadas con salto",
        "Mountain climbers",
    ],
}

DIAS_SEMANA = list(RUTINAS.keys())


def crear_hoja_registro(wb):
    """Hoja principal donde se registra cada sesión de entrenamiento."""
    ws = wb.active
    ws.title = "Registro"
    ws.sheet_properties.tabColor = AZUL_MEDIO

    # Column widths
    col_widths = {1: 14, 2: 30, 3: 30, 4: 10, 5: 10, 6: 12, 7: 18, 8: 30}
    for c, w in col_widths.items():
        ws.column_dimensions[get_column_letter(c)].width = w

    # ── Título ────────────────────────────────────────────────────
    merge_and_set(ws, 1, 1, 1, 8,
                  "REGISTRO DE ENTRENAMIENTO EN CASA",
                  font_titulo, fill_titulo)

    merge_and_set(ws, 2, 1, 2, 8,
                  "Ingresa tus datos de cada sesión. Usa los botones o escribe directamente.",
                  font_small, PatternFill("solid", fgColor=AZUL_CLARO))

    # ── Zona de entrada rápida ────────────────────────────────────
    merge_and_set(ws, 4, 1, 4, 8,
                  "ENTRADA RÁPIDA",
                  font_subtit, fill_header)

    labels = ["Fecha:", "Día/Rutina:", "Ejercicio:", "Series:", "Reps:", "Peso (kg):", "Descanso (seg):", "Notas:"]
    for i, label in enumerate(labels, start=1):
        set_cell(ws, 5, i, label, font_bold, fill_naranja_cl)

    # Input row (row 6)
    for i in range(1, 9):
        set_cell(ws, 6, i, "", font_normal, fill_input)

    # Default date formula
    ws.cell(row=6, column=1).value = '=TODAY()'
    ws.cell(row=6, column=1).number_format = "DD/MM/YYYY"

    # ── Validación desplegable para Día/Rutina ────────────────────
    dv_dia = DataValidation(
        type="list",
        formula1='"' + ",".join(DIAS_SEMANA) + '"',
        allow_blank=True,
    )
    dv_dia.error = "Selecciona un día válido"
    dv_dia.errorTitle = "Día inválido"
    dv_dia.prompt = "Selecciona el día de rutina"
    dv_dia.promptTitle = "Día"
    ws.add_data_validation(dv_dia)
    dv_dia.add(ws["B6"])

    # ── Botones de acción (celdas con texto + macro asignada via VBA) ──
    merge_and_set(ws, 8, 2, 8, 3,
                  "▶ REGISTRAR ENTRADA",
                  font_btn, fill_verde)
    merge_and_set(ws, 8, 5, 8, 6,
                  "✕ LIMPIAR CAMPOS",
                  font_btn, fill_naranja)
    merge_and_set(ws, 8, 7, 8, 8,
                  "⟳ DESHACER ÚLTIMO",
                  font_btn, fill_rojo)

    # ── Encabezados del historial ─────────────────────────────────
    merge_and_set(ws, 10, 1, 10, 8,
                  "HISTORIAL DE ENTRENAMIENTOS",
                  font_subtit, fill_titulo)

    headers = ["Fecha", "Día / Rutina", "Ejercicio", "Series", "Reps", "Peso (kg)", "Descanso (s)", "Notas"]
    for i, h in enumerate(headers, start=1):
        set_cell(ws, 11, i, h, font_header, fill_header)

    # Rows 12-200 pre-formatted for data
    for r in range(12, 201):
        fill = fill_alt1 if r % 2 == 0 else fill_alt2
        for c in range(1, 9):
            cell = set_cell(ws, r, c, "", font_normal, fill)
            if c == 1:
                cell.number_format = "DD/MM/YYYY"

    # ── Freeze panes ──────────────────────────────────────────────
    ws.freeze_panes = "A12"

    # ── Auto-filter ───────────────────────────────────────────────
    ws.auto_filter.ref = "A11:H200"

    return ws


def crear_hoja_rutinas(wb):
    """Hoja con las rutinas predefinidas por día."""
    ws = wb.create_sheet("Rutinas")
    ws.sheet_properties.tabColor = VERDE

    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 35
    ws.column_dimensions["C"].width = 35
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 14
    ws.column_dimensions["F"].width = 14

    merge_and_set(ws, 1, 1, 1, 6,
                  "RUTINAS SEMANALES - ENTRENAMIENTO EN CASA",
                  font_titulo, fill_titulo)

    merge_and_set(ws, 2, 1, 2, 6,
                  "Personaliza tus rutinas aquí. Los ejercicios aparecerán en la hoja de Registro.",
                  font_small, PatternFill("solid", fgColor=VERDE_CLARO))

    row = 4
    for dia, ejercicios in RUTINAS.items():
        merge_and_set(ws, row, 1, row, 6, dia, font_subtit, fill_header)
        row += 1

        sub_headers = ["#", "Ejercicio", "Grupo Muscular", "Series Obj.", "Reps Obj.", "Peso Obj. (kg)"]
        for i, h in enumerate(sub_headers, start=1):
            set_cell(ws, row, i, h, font_header, fill_verde)
        row += 1

        for idx, ej in enumerate(ejercicios, start=1):
            fill = fill_alt1 if idx % 2 == 0 else fill_alt2
            set_cell(ws, row, 1, idx, font_normal, fill)
            set_cell(ws, row, 2, ej, font_normal, fill, align_left)
            for c in range(3, 7):
                set_cell(ws, row, c, "", font_normal, fill)
            row += 1

        row += 1  # espacio entre días

    return ws


def crear_hoja_progreso(wb):
    """Hoja de resumen/progreso con fórmulas."""
    ws = wb.create_sheet("Progreso")
    ws.sheet_properties.tabColor = NARANJA

    ws.column_dimensions["A"].width = 25
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 18
    ws.column_dimensions["E"].width = 18

    merge_and_set(ws, 1, 1, 1, 5,
                  "RESUMEN DE PROGRESO",
                  font_titulo, fill_titulo)

    merge_and_set(ws, 2, 1, 2, 5,
                  "Estadísticas calculadas automáticamente desde tu historial.",
                  font_small, PatternFill("solid", fgColor=NARANJA_CL))

    # ── Estadísticas generales ────────────────────────────────────
    merge_and_set(ws, 4, 1, 4, 5,
                  "ESTADÍSTICAS GENERALES",
                  font_subtit, fill_header)

    stats = [
        ("Total sesiones registradas",  '=COUNTA(Registro!A12:A200)'),
        ("Último entrenamiento",        '=MAX(Registro!A12:A200)'),
        ("Peso máximo levantado (kg)",  '=MAX(Registro!F12:F200)'),
        ("Promedio de repeticiones",    '=IFERROR(AVERAGE(Registro!E12:E200),"")'),
        ("Ejercicio más frecuente",     '=IFERROR(INDEX(Registro!C12:C200,MATCH(MAX(COUNTIF(Registro!C12:C200,Registro!C12:C200)),COUNTIF(Registro!C12:C200,Registro!C12:C200),0)),"")'),
    ]

    for i, (label, formula) in enumerate(stats):
        r = 5 + i
        set_cell(ws, r, 1, label, font_bold, fill_naranja_cl, align_left)
        merge_and_set(ws, r, 2, r, 3, "", font_normal, fill_input)
        ws.cell(row=r, column=2).value = formula
        if i == 1:
            ws.cell(row=r, column=2).number_format = "DD/MM/YYYY"

    # ── Volumen por día de rutina ─────────────────────────────────
    merge_and_set(ws, 12, 1, 12, 5,
                  "VOLUMEN POR DÍA DE RUTINA",
                  font_subtit, fill_header)

    vol_headers = ["Día / Rutina", "Sesiones", "Total Series", "Total Reps", "Peso Prom. (kg)"]
    for i, h in enumerate(vol_headers, start=1):
        set_cell(ws, 13, i, h, font_header, fill_verde)

    for i, dia in enumerate(DIAS_SEMANA):
        r = 14 + i
        fill = fill_alt1 if i % 2 == 0 else fill_alt2
        set_cell(ws, r, 1, dia, font_normal, fill, align_left)
        set_cell(ws, r, 2, f'=COUNTIF(Registro!B12:B200,A{r})', font_normal, fill)
        set_cell(ws, r, 3, f'=SUMIF(Registro!B12:B200,A{r},Registro!D12:D200)', font_normal, fill)
        set_cell(ws, r, 4, f'=SUMIF(Registro!B12:B200,A{r},Registro!E12:E200)', font_normal, fill)
        set_cell(ws, r, 5, f'=IFERROR(AVERAGEIF(Registro!B12:B200,A{r},Registro!F12:F200),"")', font_normal, fill)

    # ── Últimos 10 entrenamientos ─────────────────────────────────
    merge_and_set(ws, 21, 1, 21, 5,
                  "ÚLTIMOS 10 REGISTROS",
                  font_subtit, fill_header)

    last_headers = ["Fecha", "Rutina", "Ejercicio", "Reps", "Peso (kg)"]
    for i, h in enumerate(last_headers, start=1):
        set_cell(ws, 22, i, h, font_header, fill_verde)

    for i in range(10):
        r = 23 + i
        fill = fill_alt1 if i % 2 == 0 else fill_alt2
        data_row = f'=IFERROR(INDEX(Registro!A$12:A$200,MATCH(LARGE(Registro!A$12:A$200,{i+1}),Registro!A$12:A$200,0)),"")'
        for c in range(1, 6):
            set_cell(ws, r, c, "", font_normal, fill)
        # Simplified: just reference the latest rows based on COUNTA
        src_row_formula = f'=IFERROR(COUNTA(Registro!A$12:A$200)-{i}+11,"")'
        cols_map = {1: "A", 2: "B", 3: "C", 4: "E", 5: "F"}
        for c, col_letter in cols_map.items():
            formula = f'=IFERROR(INDEX(Registro!{col_letter}$12:{col_letter}$200,COUNTA(Registro!A$12:A$200)-{i}),"")'
            ws.cell(row=r, column=c).value = formula
            if c == 1:
                ws.cell(row=r, column=c).number_format = "DD/MM/YYYY"

    return ws


def crear_hoja_instrucciones(wb):
    """Hoja con instrucciones de uso."""
    ws = wb.create_sheet("Instrucciones")
    ws.sheet_properties.tabColor = "95A5A6"

    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 80

    merge_and_set(ws, 1, 1, 1, 2,
                  "INSTRUCCIONES DE USO",
                  font_titulo, fill_titulo)

    instrucciones = [
        ("HOJA 'REGISTRO'", [
            "Esta es tu hoja principal. Aquí registras cada ejercicio que hagas.",
            "1. La FECHA se llena automáticamente con el día de hoy.",
            "2. Selecciona el DÍA/RUTINA del desplegable (Lunes-Pecho, Martes-Espalda, etc.).",
            "3. Escribe el nombre del EJERCICIO que realizaste.",
            "4. Ingresa el número de SERIES, REPETICIONES y PESO usado.",
            "5. Opcionalmente agrega el tiempo de DESCANSO y NOTAS.",
            "6. Haz clic en '▶ REGISTRAR ENTRADA' para mover los datos al historial.",
            "7. Usa '✕ LIMPIAR CAMPOS' para borrar la zona de entrada.",
            "8. Usa '⟳ DESHACER ÚLTIMO' para eliminar el último registro.",
        ]),
        ("HOJA 'RUTINAS'", [
            "Aquí están las rutinas predefinidas para 5 días de la semana.",
            "Puedes editar los ejercicios, agregar objetivos de series/reps/peso.",
            "Los nombres de los ejercicios te sirven de guía al llenar el Registro.",
        ]),
        ("HOJA 'PROGRESO'", [
            "Muestra estadísticas automáticas calculadas desde tu historial.",
            "Total de sesiones, peso máximo, promedios y volumen por día.",
            "Los últimos 10 registros se muestran al final.",
        ]),
        ("TIPS", [
            "Registra CADA serie por separado para un seguimiento más detallado.",
            "O registra el total de series/reps por ejercicio si prefieres algo rápido.",
            "Revisa la hoja de Progreso regularmente para ver tu avance.",
            "¡La constancia es la clave! Intenta entrenar al menos 3-4 días por semana.",
        ]),
    ]

    row = 3
    for titulo, lineas in instrucciones:
        merge_and_set(ws, row, 1, row, 2, titulo, font_subtit, fill_header)
        row += 1
        for linea in lineas:
            set_cell(ws, row, 1, "•", font_normal, fill_alt1)
            set_cell(ws, row, 2, linea, font_normal, fill_alt1, align_left)
            row += 1
        row += 1

    return ws


# ── Código VBA para los botones ───────────────────────────────────
VBA_CODE = '''
Attribute VB_Name = "ModEntrenamiento"

' ============================================================
' MACRO: Registrar una entrada desde la zona de entrada rápida
' ============================================================
Sub RegistrarEntrada()
    Dim wsReg As Worksheet
    Set wsReg = ThisWorkbook.Sheets("Registro")

    ' Buscar la siguiente fila vacía en el historial (desde fila 12)
    Dim nextRow As Long
    nextRow = 12
    Do While wsReg.Cells(nextRow, 1).Value <> ""
        nextRow = nextRow + 1
        If nextRow > 200 Then
            MsgBox "El historial esta lleno (200 registros). " & _
                   "Copia los datos a otro archivo para continuar.", _
                   vbExclamation, "Historial lleno"
            Exit Sub
        End If
    Loop

    ' Validar que haya al menos un ejercicio
    If Trim(wsReg.Cells(6, 3).Value) = "" Then
        MsgBox "Por favor ingresa al menos el nombre del ejercicio.", _
               vbExclamation, "Falta informacion"
        Exit Sub
    End If

    ' Copiar datos de la fila 6 (entrada) a la fila del historial
    Dim col As Integer
    For col = 1 To 8
        wsReg.Cells(nextRow, col).Value = wsReg.Cells(6, col).Value
    Next col

    ' Formatear la fecha
    wsReg.Cells(nextRow, 1).NumberFormat = "DD/MM/YYYY"

    ' Limpiar campos de entrada (excepto fecha y dia)
    wsReg.Cells(6, 3).Value = ""  ' Ejercicio
    wsReg.Cells(6, 4).Value = ""  ' Series
    wsReg.Cells(6, 5).Value = ""  ' Reps
    wsReg.Cells(6, 6).Value = ""  ' Peso
    wsReg.Cells(6, 7).Value = ""  ' Descanso
    wsReg.Cells(6, 8).Value = ""  ' Notas

    MsgBox "Entrada registrada en la fila " & nextRow & ".", _
           vbInformation, "Registrado"
End Sub


' ============================================================
' MACRO: Limpiar todos los campos de entrada rápida
' ============================================================
Sub LimpiarCampos()
    Dim wsReg As Worksheet
    Set wsReg = ThisWorkbook.Sheets("Registro")

    wsReg.Cells(6, 1).Value = Date
    wsReg.Cells(6, 1).NumberFormat = "DD/MM/YYYY"
    wsReg.Cells(6, 2).Value = ""
    wsReg.Cells(6, 3).Value = ""
    wsReg.Cells(6, 4).Value = ""
    wsReg.Cells(6, 5).Value = ""
    wsReg.Cells(6, 6).Value = ""
    wsReg.Cells(6, 7).Value = ""
    wsReg.Cells(6, 8).Value = ""

    MsgBox "Campos limpiados.", vbInformation, "Listo"
End Sub


' ============================================================
' MACRO: Deshacer el último registro (eliminar última fila con datos)
' ============================================================
Sub DeshacerUltimo()
    Dim wsReg As Worksheet
    Set wsReg = ThisWorkbook.Sheets("Registro")

    ' Encontrar la última fila con datos
    Dim lastRow As Long
    lastRow = 11
    Do While wsReg.Cells(lastRow + 1, 1).Value <> ""
        lastRow = lastRow + 1
    Loop

    If lastRow < 12 Then
        MsgBox "No hay registros para deshacer.", _
               vbExclamation, "Sin registros"
        Exit Sub
    End If

    Dim resp As VbMsgBoxResult
    resp = MsgBox("Eliminar el ultimo registro?" & vbCrLf & _
                  "Fecha: " & wsReg.Cells(lastRow, 1).Value & vbCrLf & _
                  "Ejercicio: " & wsReg.Cells(lastRow, 3).Value, _
                  vbYesNo + vbQuestion, "Confirmar")

    If resp = vbYes Then
        Dim col As Integer
        For col = 1 To 8
            wsReg.Cells(lastRow, col).Value = ""
        Next col
        MsgBox "Ultimo registro eliminado.", vbInformation, "Deshecho"
    End If
End Sub


' ============================================================
' MACRO: Cargar ejercicios de la rutina seleccionada
' ============================================================
Sub CargarEjercicios()
    Dim wsReg As Worksheet, wsRut As Worksheet
    Set wsReg = ThisWorkbook.Sheets("Registro")
    Set wsRut = ThisWorkbook.Sheets("Rutinas")

    Dim diaSeleccionado As String
    diaSeleccionado = Trim(wsReg.Cells(6, 2).Value)

    If diaSeleccionado = "" Then
        MsgBox "Primero selecciona un Dia/Rutina en la celda B6.", _
               vbExclamation, "Selecciona un dia"
        Exit Sub
    End If

    ' Buscar el día en la hoja Rutinas
    Dim r As Long
    Dim found As Boolean
    found = False

    For r = 1 To 100
        If InStr(1, wsRut.Cells(r, 1).Value, diaSeleccionado, vbTextCompare) > 0 Then
            found = True
            Exit For
        End If
    Next r

    If Not found Then
        MsgBox "No se encontro la rutina: " & diaSeleccionado, _
               vbExclamation, "No encontrado"
        Exit Sub
    End If

    ' Leer ejercicios (están 2 filas abajo del título del día)
    Dim ejercicios As String
    Dim startRow As Long
    startRow = r + 2  ' Saltar fila de sub-headers

    ejercicios = "Ejercicios para " & diaSeleccionado & ":" & vbCrLf & vbCrLf

    Dim ej As Long
    ej = 0
    Do While wsRut.Cells(startRow + ej, 2).Value <> "" And ej < 20
        ejercicios = ejercicios & (ej + 1) & ". " & wsRut.Cells(startRow + ej, 2).Value & vbCrLf
        ej = ej + 1
    Loop

    MsgBox ejercicios, vbInformation, "Rutina del dia"
End Sub
'''


def main():
    wb = openpyxl.Workbook()

    # Crear hojas
    crear_hoja_registro(wb)
    crear_hoja_rutinas(wb)
    crear_hoja_progreso(wb)
    crear_hoja_instrucciones(wb)

    # Mover Instrucciones al principio? No, dejarlo al final.
    # Asegurar que Registro es la hoja activa
    wb.active = wb.sheetnames.index("Registro")

    # ── Guardar como .xlsm con VBA ───────────────────────────────
    # openpyxl no soporta VBA nativamente en archivos nuevos.
    # Guardamos como .xlsx y creamos un archivo .bas con las macros.
    output_xlsx = "/home/user/gym/Entrenamiento_Casa.xlsx"
    output_vba  = "/home/user/gym/macros_entrenamiento.bas"

    wb.save(output_xlsx)

    # Guardar el código VBA como archivo .bas importable
    with open(output_vba, "w", encoding="utf-8") as f:
        f.write(VBA_CODE)

    print(f"Archivo Excel creado: {output_xlsx}")
    print(f"Archivo de macros VBA: {output_vba}")
    print()
    print("INSTRUCCIONES PARA ACTIVAR LAS MACROS:")
    print("=" * 50)
    print("1. Abre el archivo .xlsx en Excel")
    print("2. Presiona Alt+F11 para abrir el Editor de VBA")
    print("3. Ve a Archivo > Importar archivo...")
    print("4. Selecciona 'macros_entrenamiento.bas'")
    print("5. Cierra el editor VBA")
    print("6. Guarda el archivo como .xlsm (Libro habilitado para macros)")
    print()
    print("ASIGNAR MACROS A LOS BOTONES:")
    print("=" * 50)
    print("1. Clic derecho sobre '▶ REGISTRAR ENTRADA' > Asignar macro > RegistrarEntrada")
    print("2. Clic derecho sobre '✕ LIMPIAR CAMPOS' > Asignar macro > LimpiarCampos")
    print("3. Clic derecho sobre '⟳ DESHACER ÚLTIMO' > Asignar macro > DeshacerUltimo")
    print()
    print("ALTERNATIVA RÁPIDA (sin importar .bas):")
    print("=" * 50)
    print("En la hoja de Registro, usa Alt+F8 para ver las macros disponibles")
    print("y ejecutarlas manualmente.")


if __name__ == "__main__":
    main()
