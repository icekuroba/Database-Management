Attribute VB_Name = "App"
Public Const nombreHoja As String = "POLIZARIO"
Public Const propuesta As String = "PROPUESTA"
Public Const modificaciones As String = "MODIFICACIONES"
' Arrays globales
Public polizas() As String          ' nombre de cada póliza
Public filas() As Long              ' filas donde está cada póliza
Public nPolizas As Long             ' cantidad total de pólizas
Option Explicit
Sub MostrarCotizador()
    frmCotizador.Show
End Sub
'-----------------------------------------------
' Buscar en archivo y procesar datos
'-----------------------------------------------
Public Sub BuscarEnArchivo(rutaCotizador As String, rutaQuinquenios As String)
    Dim libroOrigen As Workbook, libroCopia As Workbook, nuevoLibro As Workbook
    Dim hojaDestino As Worksheet
    Dim archivoRuta As Variant, rutaCopia As String, rutaFinal As String
    Dim carpetaSalida As String, docsPath As String     ' Rutas
    
    ' Control
    Dim i As Long, combo As OLEObject, nombreCotizador As String
    Dim tipoCotizador As String, mesAbrev As String, poolId As String
    Dim X As Long, limpiarLibro As Worksheet
    Dim token As String, anio As String
    
    ' Rendimiento
    Dim prevCalc As XlCalculation, prevScr As Boolean, prevEvt As Boolean, prevDisp As Boolean
    
    ' Quinquenios
    Dim libroQ As Workbook, hojaQ As Worksheet, archivoQ As String, rutaQ As String, archivoRutaQ As Variant
    
    ' Parámetros
    Dim lista As Worksheet, ultimaFila As Long, j As Long
    Dim agente As String, partes() As String, partesQ() As String
    Dim y2 As String, y4 As String, falla As String, nomQ As String
    Dim P As Long, mes As String
    Dim nombreQuinquenios As String
    On Error GoTo msj
    '-----------------------------------------------
    ' 1) Abre los archivos y ordena las pólizas por quinquenios
    '-----------------------------------------------
    archivoRuta = rutaCotizador
    archivoRutaQ = rutaQuinquenios
    Set libroOrigen = Workbooks.Open(archivoRuta)
    Set libroQ = Workbooks.Open(archivoRutaQ)

    ' Ordena la columna B del libro de quinquenios
    Set hojaQ = libroQ.Sheets(1)
    ultimaFila = hojaQ.Cells(hojaQ.Rows.Count, "B").End(xlUp).Row

    With hojaQ.Sort
        .SortFields.Clear
        .SortFields.Add Key:=hojaQ.Range("B2:B" & ultimaFila), _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange hojaQ.Range("A1:G" & ultimaFila)
        .Header = xlYes
        .Apply
    End With
    Debug.Print "El archivo de quinquenios ya está ordenado"

    '-----------------------------------------------
    ' 2) Desbloquear libro y limpiar datos
    '-----------------------------------------------
    Call Desbloquear(libroOrigen)

    On Error Resume Next
    Set limpiarLibro = libroOrigen.Sheets(propuesta)
    On Error GoTo 0

    If Not limpiarLibro Is Nothing Then
        limpiarLibro.Range("E37:F50").ClearContents '("E95:F112").ClearContents
        Debug.Print "Censo limpiado con éxito"
    Else
        Debug.Print "No se limpió censo porque no se encontró la hoja de polizario"
    End If

    libroOrigen.Save

    ' Guardar nombre del archivo
    nombreCotizador = nombresSinExtension(libroOrigen.name)
    partes = Split(nombreCotizador, "_")
    tipoCotizador = ""
    poolId = ""
    mesAbrev = ""
    anio = ""

    '3) Extraer datos de quinquenios
    Set lista = libroContrasena.Sheets(1)
    ultimaFila = lista.Cells(lista.Rows.Count, "A").End(xlUp).Row
    For X = LBound(partes) To UBound(partes)
        token = UCase(Trim(partes(X)))
        If tipoCotizador = "" Then
            For j = 1 To ultimaFila
                agente = UCase(Trim(lista.Cells(j, 1).Value))
                If token Like "*" & agente & "*" Then
                    tipoCotizador = lista.Cells(j, 2).Value
                    poolId = lista.Cells(j, 2).Value
                    Exit For
                End If
            Next j
        End If
        
    If mesAbrev = "" Then mesAbrev = ObtenerMes(token)
    If anio = "" And IsNumeric(Right(token, 2)) Then
        anio = Right(token, 2)
    End If
    If tipoCotizador <> "" And mesAbrev <> "" And anio <> "" Then Exit For
    Next X
    
    '4) Validacion para verificar si coincide el archivo de cotizador con el de quinquenios
    nomQ = UCase$(nombresSinExtension(libroQ.name))
    mes = UCase$(mesAbrev)
    y2 = anio
    y4 = "20" & anio
    If InStr(1, nomQ, CStr(poolId)) = 0 Then falla = "El quinquenio no contiene el pool " & poolId
    If falla = "" And mes <> "" Then
        If InStr(1, nomQ, mes, vbTextCompare) = 0 Then falla = "Mes distinto (cotizador: " & mesAbrev & ")"
    End If
    If falla = "" And InStr(1, tipoCotizador, "A", vbTextCompare) > 0 Then
        If y2 <> "" And (InStr(1, nomQ, y2) = 0 And InStr(1, nomQ, y4) = 0) Then
            falla = "Año distinto (cotizador: " & y4 & ")"
        End If
    End If
    If Len(falla) > 0 Then
        MsgBox "El archivo de quinquenios no corresponde al del cotizador " & "Detalle: " & falla & vbCrLf & "y " & nomQ, vbCritical, "Verifica los archivos."
        On Error Resume Next: libroOrigen.Close savechanges:=False: On Error GoTo 0
        Exit Sub
    End If
    
    '5) Verificar si la hoja POLIZARIO y la celda NPOLIZA existe
    On Error Resume Next
    Set hojaDestino = libroOrigen.Sheets(nombreHoja)
    On Error GoTo msj
    If hojaDestino Is Nothing Then
        MsgBox "La hoja " & nombreHoja & " no existe.", vbCritical
        libroOrigen.Close savechanges:=False
        Exit Sub
    End If
    
    If Trim$(UCase$(hojaDestino.Range("b8").Value)) <> "NPOLIZA" Then
        MsgBox "No se encuentra la celda NPOLIZA. No se puede continuar.", vbCritical
        libroOrigen.Close False
        Exit Sub
    End If
    Debug.Print "Archivo y hoja validados correctamente"
    
    '6) Copia del cotizador que se este ejecutando
    rutaCopia = Environ("TEMP") & "\" & Replace(libroOrigen.name, ".xlsm", "") & "_COPIA.xlsm"
    'rutaCopia = libroOrigen.path & "\" & Replace(libroOrigen.name, ".xlsm", "") & "_COPIA.xlsm"
    libroOrigen.SaveCopyAs rutaCopia
    libroOrigen.Close savechanges:=True
    
    '7) Carpeta de salida, verificando en OneDrive despues en la carpeta local
    docsPath = rutaDocumentos()
    MsgBox "Se guardara en: " & vbCrLf & docsPath, vbInformation
    rutaFinal = docsPath & "\" & Replace(nombresSinExtension(rutaCopia), "_COPIA", "") & Format(Now(), "YYMMDD")
    If Dir(rutaFinal, vbDirectory) = vbNullString Then MkDir rutaFinal
    archivoQ = libroQ.name
    Debug.Print "archivo de quinquenio: " & archivoQ
    
    '8) Afinando entorno
    prevCalc = Application.Calculation
    prevScr = Application.ScreenUpdating
    prevEvt = Application.EnableEvents
    prevDisp = Application.DisplayAlerts
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True 'debe quedar en true para que no se rompa en lista() de la macro
    Application.DisplayAlerts = False
    Application.WindowState = xlNormal
    
    '9) Se abre copia para trabajar
    Application.EnableEvents = False
    Set libroCopia = Workbooks.Open(rutaCopia) 'Call Desbloquear(libroCopia)
    Application.EnableEvents = True 'Debe estar en true para que se vea reflejado el cambio del desbloqueo
    If libroCopia.Sheets(propuesta).ProtectContents = True Then
        Debug.Print "La copia sigue bloqueada"
    Else
        Debug.Print "La copia esta desbloqueada"
    End If
    
    '10) Leer poliza
    If Not LeerPoliza(libroCopia.Sheets(nombreHoja)) Then
        Debug.Print "No hay polizas por procesar"
        MsgBox "No hay polizas por procesar", vbExclamation
        libroCopia.Close False
        GoTo Auxiliar
    End If
    Debug.Print "Cantidad de polizas detectadas: " & nPolizas
    Call IniciarRegistros
    
    '11) Procesar polizas
    For i = 1 To nPolizas
        Set combo = libroCopia.Sheets(propuesta).OLEObjects("ComboBox1")
        
        'Activar libro y hoja base
        Debug.Print "Libro activo: " & ActiveWorkbook.name
        Debug.Print "Existe hoja 5 " & ExisteHoja("MODIFICACIONES", libroCopia)
        
        libroCopia.Sheets(modificaciones).Select
        On Error Resume Next
        combo.Object.Value = polizas(i)
        On Error GoTo 0
        Debug.Print "Procesando en: " & polizas(i)
        
        '--- funciones de la macro ---
        Application.Run "'" & libroCopia.name & "'!subgrupo"
        Application.CalculateFullRebuild
        Call Desbloquear(libroCopia)
        libroContrasena.Close savechanges:=False
        
        If archivoQ <> "" Then
            numPolizaGlobal = polizas(i)
            Call Quinquenios(libroCopia, archivoQ)
            If Not libroCopia Is Nothing Then
                libroCopia.Activate
                Debug.Print "Hojas del libro: " & libroCopia.Worksheets.Count
                Debug.Print "Libro activo para Tarifas_enlace: " & ActiveWorkbook.name
            Else
                Debug.Print "no se encontro libroCopia antes de Tarifa", vbCritical
                Exit Sub
            End If
            Application.Run "'" & libroCopia.name & "'!Tarifa"
        End If
        Application.CalculateFullRebuild 'Recalcula todo
        libroCopia.Close False
        Set nuevoLibro = ActiveWorkbook
        Debug.Print nuevoLibro.name
        
        '12) Se guarda por libros cada poliza en la carpeta "Documentos de OneDrive"
        carpetaSalida = rutaFinal & "\" & LimpiarArchivo(polizas(i)) & ".xlsx"
        nuevoLibro.SaveAs Filename:=carpetaSalida, FileFormat:=51
        nuevoLibro.Close
        Call RegistrarPolizas(CStr(polizas(i)), carpetaSalida)
        Debug.Print "Archivo guardado : " & carpetaSalida
        On Error GoTo 0
        Set libroCopia = Workbooks.Open(rutaCopia)
    Next i
    
    MsgBox "Todas las polizas fueron procesadas correctamente.", vbInformation
    libroCopia.Close False
    libroQ.Close True
    Call EnviarCorreo
    
Auxiliar:
    On Error Resume Next
    Application.Calculation = prevCalc
    Application.ScreenUpdating = prevScr
    Application.EnableEvents = prevEvt
    Application.DisplayAlerts = prevDisp
    Application.StatusBar = False
    Exit Sub
    
msj:
    Debug.Print "Error: " & Err.Description
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Auxiliar
End Sub




