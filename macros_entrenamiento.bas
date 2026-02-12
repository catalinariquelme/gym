
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
