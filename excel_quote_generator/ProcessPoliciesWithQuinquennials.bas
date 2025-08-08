Option Explicit

Public Sub ProcesarPolizasConQuinquenios()
    On Error GoTo fallo

    ' ===== Rendimiento / entorno =====
    Dim prevCalc As XlCalculation
    prevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    ' ===== Selección de archivos =====
    MsgBox "Selecciona el archivo del cotizador (.xlsm).", vbInformation
    Dim rutaCotizador As Variant
    rutaCotizador = Application.GetOpenFilename("Excel Macro-Enabled (*.xlsm), *.xlsm", , "Select cotizador file")
    If rutaCotizador = False Then GoTo salida

    MsgBox "Selecciona el archivo de quinquenios (.xlsx).", vbInformation
    Dim rutaQuinquenios As Variant
    rutaQuinquenios = Application.GetOpenFilename("Excel Workbook (*.xlsx), *.xlsx", , "Select quinquennials file")
    If rutaQuinquenios = False Then GoTo salida

    ' ===== Abrir libros =====
    Dim wbCot As Workbook, wbQ As Workbook
    Set wbCot = Workbooks.Open(CStr(rutaCotizador), ReadOnly:=False)
    Set wbQ = Workbooks.Open(CStr(rutaQuinquenios), ReadOnly:=True)

    ' ===== Constantes configurables =====
    Const SH_POLIZARIO As String = "POLIZARIO"
    Const COL_POLIZA As Long = 2                ' Columna B
    Const FILA_INICIO As Long = 9
    Const SH_PROPUESTA As String = "PROPUESTA DE RENOVACIÓN"
    Const CELDA_QUINQUENIO As String = "D15"
    Const RANGO_Q_LOOKUP As String = "A:B"      ' Col A: póliza, Col B: quinquenio
    Dim passwords() As Variant
    passwords = Array("PoolMercer2021", "Pool2023Despachosr", "")

    ' ===== Validaciones de hojas =====
    Dim shPol As Worksheet, shProp As Worksheet, shQ As Worksheet
    Set shPol = GetSheet(wbCot, SH_POLIZARIO)
    Set shProp = GetSheet(wbCot, SH_PROPUESTA)
    Set shQ = wbQ.Sheets(1) ' Ajustar si corresponde

    If shPol Is Nothing Or shProp Is Nothing Then
        MsgBox "No se encontraron hojas requeridas en el cotizador.", vbCritical
        GoTo salida
    End If

    ' ===== Desproteger hojas del cotizador (si aplica) =====
    Dim sh As Worksheet
    For Each sh In wbCot.Worksheets
        If sh.ProtectContents Then
            If Not TryUnprotect(sh, passwords) Then
                MsgBox "No se pudo desproteger la hoja: " & sh.Name, vbCritical
                GoTo salida
            End If
        End If
    Next sh

    ' ===== Carpeta de salida =====
    Dim rutaSalida As String
    rutaSalida = Environ$("USERPROFILE") & "\Documents\" & Replace(wbCot.Name, ".xlsm", "") & "_Processed"
    EnsureFolder rutaSalida

    ' ===== Preparar índice de quinquenios (Match + lectura) =====
    Dim lastQ As Long
    lastQ = shQ.Cells(shQ.Rows.Count, 1).End(xlUp).Row
    Dim rangoQ As Range
    Set rangoQ = shQ.Range("A1").Resize(lastQ, 2)

    ' ===== Recorrido de pólizas =====
    Dim lastRow As Long, fila As Long
    lastRow = shPol.Cells(shPol.Rows.Count, COL_POLIZA).End(xlUp).Row

    For fila = FILA_INICIO To lastRow
        Dim nombrePoliza As String
        nombrePoliza = Trim(CStr(shPol.Cells(fila, COL_POLIZA).Value))
        If Len(nombrePoliza) = 0 Then GoTo siguiente

        ' Buscar posición con Match
        Dim pos As Variant
        pos = Application.Match(nombrePoliza, rangoQ.Columns(1), 0)

        If Not IsError(pos) And Not IsEmpty(pos) Then
            ' Leer quinquenio de la columna 2
            Dim quinquenio As Variant
            quinquenio = rangoQ.Cells(CLng(pos), 2).Value
            shProp.Range(CELDA_QUINQUENIO).Value = quinquenio
        Else
            ' Puede ser informativo; no interrumpir el proceso por una no-coincidencia
            Debug.Print "Quinquenio no encontrado para: " & nombrePoliza
        End If

        ' ===== Ejecutar macros dependientes (si existen) =====
        RunIfExists wbCot, "subgrupos"
        RunIfExists wbCot, "Tarifas_enlace"
        RunIfExists wbCot, "Tarifa_Modificaciones"
        RunIfExists wbCot, "resumen"

        ' ===== Copiar hojas y guardar nuevo libro =====
        Dim nuevo As Workbook
        wbCot.Sheets(Array(SH_PROPUESTA, "Textos", "Endosos")).Copy
        Set nuevo = ActiveWorkbook

        Dim nombreFinal As String
        nombreFinal = nombrePoliza & "_" & Format(Now, "yyyymmdd_HHMMss") & ".xlsm"
        nuevo.SaveAs Filename:=rutaSalida & "\" & nombreFinal, FileFormat:=xlOpenXMLWorkbookMacroEnabled
        nuevo.Close SaveChanges:=False

siguiente:
    Next fila

    MsgBox "Proceso completado.", vbInformation

salida:
    On Error Resume Next
    If Not wbCot Is Nothing Then wbCot.Close SaveChanges:=False
    If Not wbQ Is Nothing Then wbQ.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Application.Calculation = prevCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    Exit Sub

fallo:
    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical
    Resume salida
End Sub

' === Utilidades ===

Private Function GetSheet(wb As Workbook, ByVal name As String) As Worksheet
    On Error Resume Next
    Set GetSheet = wb.Sheets(name)
    On Error GoTo 0
End Function

Private Function TryUnprotect(ws As Worksheet, ByRef passwords As Variant) As Boolean
    Dim i As Long
    On Error Resume Next
    ws.Unprotect Password:=""
    If Err.Number = 0 And Not ws.ProtectContents Then
        TryUnprotect = True
        Exit Function
    End If
    Err.Clear
    For i = LBound(passwords) To UBound(passwords)
        ws.Unprotect Password:=passwords(i)
        If Err.Number = 0 And Not ws.ProtectContents Then
            TryUnprotect = True
            Exit Function
        End If
        Err.Clear
    Next i
    On Error GoTo 0
End Function

Private Sub EnsureFolder(ByVal path As String)
    If Len(Dir$(path, vbDirectory)) = 0 Then MkDir path
End Sub

Private Sub RunIfExists(wb As Workbook, procName As String)
    On Error Resume Next
    Application.Run "'" & wb.Name & "'!" & procName
    Err.Clear
    On Error GoTo 0
End Sub
