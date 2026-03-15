'================================================================================================================================
'                                   Cotizaciones - Modulo principal
'================================================================================================================================
Public Const nombreHoja As String = "POLIZARIO"
Public Const propuesta As String = "PROPUESTA"
Public Const modificaciones As String = "MODIFICACIONES"
'   =====    Arrays globales    ======
Public polizas() As String 'nombre de cada poliza
Public filas() As Long  'filas donde esta cada poliza
Public nPolizas As Long  'cantidad total de polizas
Option Explicit
Sub MostrarCotizador()
    frmCotizad.Show
End Sub
'--------------------------------------------------------------------------------------------------------------------------------
Public Sub BuscarEnArchivo(rutaCotizador As String, rutaQuinquenios As String)
    Dim libroOrigen As Workbook, libroCopia As Workbook, nuevoLibro As Workbook
    Dim hojaDestino As Worksheet
    Dim archivoRuta As Variant, rutaCopia As String, rutaFinal As String, carpetaSalida As String, docsPath As String '-- Rutas
    ' ---  Control  ---
    Dim i As Long, combo As OLEObject, nombreCotizador As String, tipoCotizador As String
    Dim token As String, anio As String, mesAbrev As String, poolId As String, X As Long, limpiarLibro As Worksheet
    Dim ajuste As Long, inicioFila As Long, finFila As Long

    ' --  Rendimiento  ---
    Dim prevCalc As XlCalculation, prevScr As Boolean, prevEvt As Boolean, prevDisp As Boolean, tipoEncontrado As Boolean
    ' -- Quinquenios --
    Dim libroQ As Workbook, hojaQ As Worksheet, archivoQ As String, rutaQ As String, archivoRutaQ As Variant
    ' --  Parametros  --
    Dim lista As Worksheet, ultimaFila As Long, j As Long, agente As String, partes() As String, partesQ() As String
    Dim y2 As String, y4 As String, falla As String, nomQ As String, p As Long, mes As String, mesQ As String, poolIDQ As String
    On Error GoTo msj
' --------------------------------------------------------------------------------------------------------------------------------
    '-- 1) Abre los archivos y ordena de menor a mayor las polizas en quinquenios
    archivoRuta = rutaCotizador
    archivoRutaQ = rutaQuinquenios
    Set libroOrigen = Workbooks.Open(archivoRuta)
    Set libroQ = Workbooks.Open(archivoRutaQ)
    
    Set hojaQ = libroQ.Sheets(1) ' == Ordena la columna B del libro de quinquenios
    ultimaFila = hojaQ.Cells(hojaQ.Rows.Count, "B").End(xlUp).Row
    With hojaQ.Sort
        .SortFields.Clear
        .SortFields.Add Key:=hojaQ.Range("B2:B" & ultimaFila), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange hojaQ.Range("A1:G" & ultimaFila)
        .Header = xlYes
        .Apply
    End With
    Debug.Print "El archivo de quinquenios ya esta ordenado"
    
    '-- 2)Desbloquear libro y limpiar datos de la hoja propuesta de renovación
    Call Desbloquear(libroOrigen)
    
    On Error Resume Next
    Set limpiarLibro = libroOrigen.Sheets(propuesta)
    On Error GoTo 0
    
    If limpiarLibro Is Nothing Then
        Debug.Print "No se encontró la hoja: " & propuesta
    Else
        Debug.Print "Hoja propuesta encontrada: " & propuesta
    End If

    ' == Guardar nombre del archivo
    nombreCotizador = nombreSinExtension(libroOrigen.name)
    partes = Split(nombreCotizador, "_")
    tipoCotizador = ""
    poolId = ""
    mesAbrev = ""
    anio = ""
    
    '-- 3) Extraer datos de cotizador
    Set lista = libroContrasena.Sheets(1)
    ultimaFila = lista.Cells(lista.Rows.Count, "A").End(xlUp).Row
    nombreCotizador = nombreSinExtension(libroOrigen.name)
    partes = Split(nombreCotizador, "_")
    tipoCotizador = ""
    poolId = ""
    mesAbrev = ""
    anio = ""
    
    For X = LBound(partes) To UBound(partes)
        token = UCase$(Trim(partes(X)))
    
        ' === A) Buscar en la hoja "lista" con texto normalizado ===
        If tipoCotizador = "" Then
            Dim tokenNorm As String: tokenNorm = NormalizarTexto(token)
            For j = 1 To ultimaFila
                agente = CStr(lista.Cells(j, 1).Value)
                If Len(Trim$(agente)) > 0 Then
                    Dim agenteNorm As String: agenteNorm = NormalizarTexto(agente)
                    If InStr(1, tokenNorm, agenteNorm, vbTextCompare) > 0 Then
                        tipoCotizador = lista.Cells(j, 1).Value
                        poolId = CStr(lista.Cells(j, 2).Value)
                        Exit For
                    End If
                End If
            Next j
        End If
    
        ' === B) Mes y año ===
        If mesAbrev = "" Then mesAbrev = ObtenerMes(token)
        If anio = "" And IsNumeric(Right(token, 2)) Then anio = Right(token, 2)
        If tipoCotizador <> "" And mesAbrev <> "" And anio <> "" Then Exit For
    Next X
    
    ' === C) Fallback: si no lo halló en "lista", detectar por nombre
    If tipoCotizador = "" Then
        tipoCotizador = DetectarTipoDesdeNombre(libroOrigen.name)
    End If
    
    ' === 3.5) LIMPIEZA del CENSO (E:F)
    On Error Resume Next
    Set limpiarLibro = libroOrigen.Sheets(propuesta)
    On Error GoTo 0
    
    If Not limpiarLibro Is Nothing Then
        Dim rngCenso As Range
        ' Busca la última tabla cuyo título contenga "Tabla" Inicio: 2 filas debajo del encabezado "CENSO" | Fin: 2 filas arriba de "SUBTOTAL"
        Set rngCenso = ObtenerRangoCensoPorTitulo(limpiarLibro, tituloRenov:="Tabla", offsetInicio:=2, offsetFin:=2, guardarNombre:="RANGO_CENSO")
        If Not rngCenso Is Nothing Then
            rngCenso.ClearContents
            Debug.Print "Censo limpiado por título. Rango: " & rngCenso.Address(False, False)
        Else
            Debug.Print "No se detectó el rango de censo por título/encabezados; no se limpió."
        End If
    Else
        Debug.Print "No se limpió censo: hoja propuesta no existe."
    End If
    
    '-- 4) Validacion para verificar si coincide el archivo de cotizador con el de quinquenios
    If InStr(1, tipoCotizador, "EMPR1", vbTextCompare) > 0 Then
        If mesAbrev = "" Or anio = "" Then
            Debug.Print "No coincide el mes o año"
            MsgBox "No se encontro mes/año en el nombre: " & nombreCotizador, vbCritical: Exit Sub
        End If
    Else
        If mesAbrev = "" Then
            Debug.Print "No coincide el mes" & tipoCotizador
            MsgBox "No se encontro mes en el nombre: " & nombreCotizador, vbCritical: Exit Sub
        End If
    End If
    Debug.Print "Detectado -> Tipo: " & tipoCotizador & " | Pool: " & poolId & " | Mes: " & mesAbrev & " | Año: 20" & anio
    
    ' Validacion del mes en el archivo de quinquenios
    nomQ = UCase$(nombreSinExtension(libroQ.name))
    mesQ = ""
    If InStr(1, nomQ, "JAN", vbTextCompare) > 0 Or InStr(1, nomQ, "ENE", vbTextCompare) > 0 Then mesQ = "Ene"
    If InStr(1, nomQ, "FEB", vbTextCompare) > 0 Then mesQ = "Feb"
    If InStr(1, nomQ, "MAR", vbTextCompare) > 0 Then mesQ = "Mar"
    If InStr(1, nomQ, "APR", vbTextCompare) > 0 Or InStr(1, nomQ, "ABR", vbTextCompare) > 0 Then mesQ = "Abr"
    If InStr(1, nomQ, "MAY", vbTextCompare) > 0 Then mesQ = "May"
    If InStr(1, nomQ, "JUN", vbTextCompare) > 0 Then mesQ = "Jun"
    If InStr(1, nomQ, "JUL", vbTextCompare) > 0 Then mesQ = "Jul"
    If InStr(1, nomQ, "AUG", vbTextCompare) > 0 Or InStr(1, nomQ, "AGO", vbTextCompare) > 0 Then mesQ = "Ago"
    If InStr(1, nomQ, "SEP", vbTextCompare) > 0 Then mesQ = "Sep"
    If InStr(1, nomQ, "OCT", vbTextCompare) > 0 Then mesQ = "Oct"
    If InStr(1, nomQ, "NOV", vbTextCompare) > 0 Then mesQ = "Nov"
    If InStr(1, nomQ, "DEC", vbTextCompare) > 0 Or InStr(1, nomQ, "DIC", vbTextCompare) > 0 Then mesQ = "Dic"
    If mesQ <> "" Then mes = mesQ
    y2 = anio
    y4 = "20" & anio
    poolIDQ = Trim(Split(Split(UCase(nomQ), "POOL")(1), "-")(0))
    
    '============= Validación entre el archivo de cotizador con el archivo de quinquenios ===============
    If Val(poolIDQ) <> Val(poolId) Then
        Debug.Print "Los archivos no coinciden." & vbCrLf & "El archivo de cotizador " & poolId & " no corresponde al archivo de quinquenios " & poolIDQ
        MsgBox "Los archivos no coinciden." & vbCrLf & "El archivo de cotizador " & poolId & " no corresponde al archivo de quinquenios " & poolIDQ & ". Verifica que ambos archivos sean de la misma agencia.", vbCritical
        Exit Sub
    End If
    
    If falla = "" And InStr(1, tipoCotizador, "EMPR1", vbTextCompare) > 0 Then
        If y2 <> "" And InStr(1, nomQ, y2) = 0 And InStr(1, nomQ, y4) = 0 Then
        MsgBox "Hay una falla con el año del cotizador " & y4 & " con el de quinquenios. Verifica que ambos archivos coincidan en el año.", vbCritical
        Exit Sub
        End If
    End If
    
    If mes <> mesAbrev Then
        MsgBox "Los archivos no coinciden en el mes." & vbCrLf & "El cotizador es de " & mesAbrev & " y el de quinquenios es de " & mes & ". Verifica que ambos archivos sean del mismo mes.", vbCritical
        On Error Resume Next: libroOrigen.Close savechanges:=False: On Error GoTo 0: Exit Sub
    End If
    
    '-- 5)Verificar si la hoja "POLIZARIO"  y la celda "NPOLIZA" existe
    On Error Resume Next
    Set hojaDestino = libroOrigen.Sheets(nombreHoja)
    On Error GoTo msj
    If hojaDestino Is Nothing Then
        MsgBox "La hoja ' " & nombreHoja & " ' no existe.", vbCritical
        libroOrigen.Close savechanges:=False: Exit Sub
    End If
    
    If Trim$(UCase$(hojaDestino.Range("B8").Value)) <> "NPOLIZA" Then
        MsgBox "No se encunetra la celda NPOLIZA. No se puede continuar", vbCritical
        libroOrigen.Close False: Exit Sub
    End If
    Debug.Print "Archivo y hoja validados correctamente"
    
    '-- 6)Copia del cotizador que se este ejecutando
    rutaCopia = Environ("TEMP") & "\" & Replace(libroOrigen.name, ".xlsm", "") & "_COPIA.xlsm"
    libroOrigen.SaveCopyAs rutaCopia
    libroOrigen.Close savechanges:=True

    '-- 7) Carpeta de salida, verificando en OneDrive despues en la carpeta local
    docsPath = rutaDocumentos()
        MsgBox "Se guardara en: " & vbCrLf & docsPath, vbInformation
    rutaFinal = docsPath & "\" & Replace(nombreSinExtension(rutaCopia), "_COPIA", "") & Format(Now(), "YYMMDD")
    If Dir(rutaFinal, vbDirectory) = vbNullString Then MkDir rutaFinal
        archivoQ = libroQ.name
    Debug.Print "archivo de quinquenio: " & archivoQ
    
    '-- 8) Afinando entorno
    prevCalc = Application.Calculation
    prevScr = Application.ScreenUpdating
    prevEvt = Application.EnableEvents
    prevDisp = Application.DisplayAlerts
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True 'debe quedar en true para que no se rompa en lista() de la macro
    Application.DisplayAlerts = False
    Application.WindowState = xlNormal
    
    '-- 9) Se abre copia para trabajar
    Application.EnableEvents = False
    Set libroCopia = Workbooks.Open(rutaCopia) 'Call Desbloquear(libroCopia)
    Application.EnableEvents = True 'Debe estar en true para que se vea reflejado el cambio del desbloqueo de la macro
    If libroCopia.Sheets(propuesta).ProtectContents = True Then
        Debug.Print "La copia sigue bloqueada"
    Else
        Debug.Print "La copia esta desbloqueada"
    End If
    
    '-- 10) Leer poliza
    If Not LeerPoliza(libroCopia.Sheets(nombreHoja)) Then
        Debug.Print "No hay polizas por procesar"
        MsgBox "No hay polizas por procesar", vbExclamation
        libroCopia.Close False
        GoTo Auxiliar
    End If
    Debug.Print "Cantidad de polizas detectadas: " & nPolizas
    Call IniciarRegistros
    
    '-- 11) Procesar polizas
    For i = 1 To nPolizas
    Set combo = libroCopia.Sheets(propuesta).OLEObjects("ComboBox1")
        '  Activar libro y la hoja base
        Debug.Print "Libro activo: " & ActiveWorkbook.name
        Debug.Print "Existe hoja 5 " & ExisteHoja("MODIFICACIONES", libroCopia)
        '  Selección de polizas
    libroCopia.Sheets(modificaciones).Select
    On Error Resume Next
    combo.Object.Value = polizas(i)
    On Error GoTo 0
        Debug.Print "Procesando en: " & polizas(i)
        ' ===  funciones de la macro ===
    Application.Run "'" & libroCopia.name & "'!subgrupo"
    Application.CalculateFullRebuild
    Call Desbloquear(libroCopia)
    libroContrasena.Close savechanges:=False
    If archivoQ <> "" Then
        numPolizaGlobal = polizas(i)
        If Not Quinquenios(libroCopia, archivoQ) Then                       '++++++ Función de quinquenios +++
            Debug.Print "Poliza omitida (no existe en quinquenios): " & polizas(i)
            libroCopia.Close False
            Set libroCopia = Workbooks.Open(rutaCopia)
            GoTo siguiente
        End If
        
        If Not libroCopia Is Nothing Then
            libroCopia.Activate
            Debug.Print "Hojas del libro: " & libroCopia.Worksheets.Count
            Debug.Print "libro activo para Tarifa: " & ActiveWorkbook.name
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
                            
    '-- 12) Se guarda por libros cada poliza en la carpeta "Documentos de OneDrive"
    carpetaSalida = rutaFinal & "\" & LimpiarArchivo(polizas(i)) & ".xlsx"
    On Error Resume Next
    nuevoLibro.SaveAs fileName:=carpetaSalida, FileFormat:=51
    nuevoLibro.Close
    Call RegistrarPolizas(CStr(polizas(i)), carpetaSalida)
    Debug.Print "Registrada poliza #" & i & "->" & polizas(i)
    On Error GoTo 0
    Set libroCopia = Workbooks.Open(rutaCopia)

siguiente:
    Next i
    libroCopia.Close False
    libroQ.Close True
    'Debug.Print String(70, "-") ' codigo de verificación de que si lee las polizas
    'Debug.Print "ContadorPolizas = "; contadorPolizas
    'Dim key As Long
    'For key = 1 To contadorPolizas
        'Debug.Print "Indice " & key & " | Polizas=[" & polizasProcesadas(key) & "] | Rutas=[" & rutasProcesadas(key) & "]"
    'Next key
    'Debug.Print String(70, "-")
    MsgBox "Todas las polizas fueron procesadas correctamente.", vbInformation
    Call EnviarCorreo(tipoCotizador)
    
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
