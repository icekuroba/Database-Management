'================================================================================================================================
'                                 Funcion para el libro de Quinquenios
'================================================================================================================================
Option Explicit
Public Const propuesta As String = "PROPUESTA"
Public Const datos As String = "DATOS"
'   =====    Variable global    ======
Public numPolizaGlobal As String 'nombre de cada poliza
'--------------------------------------------------------------------------------------------------------------------------------
Public Function Quinquenios(libroCopia As Workbook, archivoQ As String) As Boolean
    Quinquenios = False
    If libroCopia Is Nothing Then Exit Function
    If archivoQ = "" Then Exit Function
    
    Dim libroQ As Workbook, hojaPropuesta As Worksheet, hojaDatos As Worksheet, hojaQ As Worksheet
    ' contadores y control
    Dim i As Long, j As Long, listaFiltrada() As Variant, k As Long, datosQ As Variant
    Dim ultimaFila As Long, nSub As Long, clave As Variant, contarSub As Object
    Dim filaOut As Long, totalCols As Long, totalFilas As Long, valorCelda As String, fila As Long
    Dim numPolizaActual As String, grupos As Object, llaves As Variant, subg As Long, pos As Long
    'oldCalc As XlCalculation
' ----------------------------------------------------------------------------------------------------
    'oldCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    
    Set hojaPropuesta = libroCopia.Sheets(propuesta)
    Set hojaDatos = libroCopia.Sheets(datos)
    
    '-- 1) Cargar datos en el arreglo de la hoja del libro de quinquenios
    Set libroQ = Workbooks.Open(archivoQ, ReadOnly:=True)
    Set hojaQ = libroQ.Sheets(1)
    ultimaFila = hojaQ.Cells(hojaQ.Rows.Count, "B").End(xlUp).Row
    If ultimaFila < 2 Then
        Debug.Print "El archivo no tiene datos"
        libroQ.Close False
        GoTo Salir
    End If
    datosQ = hojaQ.Range("A2:G" & ultimaFila).Value

    '-- 2) Filtrar solo la poliza actual
    numPolizaActual = numPolizaGlobal
    totalFilas = UBound(datosQ, 1)
    totalCols = UBound(datosQ, 2)
    filaOut = 0
    For i = 1 To totalFilas
        valorCelda = Trim$(CStr(datosQ(i, 2)))
        If CLng(Val(valorCelda)) = CLng(Val(numPolizaActual)) Then
            filaOut = filaOut + 1
        End If
    Next i
        
    If filaOut = 0 Then
        Debug.Print "No hay registros en quinquenios para la poliza" & numPolizaActual
        Exit Function 'GoTo salir
    End If
    
    '-- 3) Redimensiona y copia coincidencias
    ReDim listaFiltrada(1 To filaOut, 1 To totalCols)
    fila = 0
    For i = 1 To totalFilas
        valorCelda = Trim$(CStr(datosQ(i, 2)))
        If CLng(Val(valorCelda)) = CLng(Val(numPolizaActual)) Then
            fila = fila + 1
            For j = 1 To totalCols
                listaFiltrada(fila, j) = datosQ(i, j)
            Next j
        End If
    Next i
    Debug.Print filaOut; " registros encontrados en la poliza " & numPolizaActual

    '-- 4) Contar subgrupo unicos en lista filtrada
    Set grupos = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(listaFiltrada, 1)
        clave = CLng(Val(listaFiltrada(i, 3)))
        If Not grupos.Exists(clave) Then Set grupos(clave) = New Collection
        grupos(clave).Add i
    Next i
    Debug.Print "Subgrupo detectados: " & clave
    
    '-- 5) Ordenar claves de subgrupo y llamar a LlenarCenso
    llaves = grupos.Keys
    Call SortVariantArray(llaves)
    pos = 0
    For k = LBound(llaves) To UBound(llaves)
        subg = CLng(llaves(k))
        pos = pos + 1
        Debug.Print " Llenando subgrupo: "; subg; " en la posicion "; pos
        Call LlenarCenso(libroCopia, hojaPropuesta, pos, listaFiltrada, grupos(subg))
    Next k
Salir:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = False
    Quinquenios = True
End Function

' *******************************************************************************************************************************
'                                      Subfunciones de la función de quinquenios
' *******************************************************************************************************************************
Private Sub SortVariantArray(arr As Variant)
    Dim i As Long, j As Long, tmp As Variant
    For i = LBound(arr) + 1 To UBound(arr)
        tmp = arr(i): j = i - 1
        Do While j >= LBound(arr) And CLng(arr(j)) > CLng(tmp)
            arr(j + 1) = arr(j)
            j = j - 1
        Loop
        arr(j + 1) = tmp
    Next i
End Sub
'-----------------------------------------------------------------------------------------------------------------------------
'                                           Llenar censo en propuesta
Private Sub LlenarCenso(libroCopia As Workbook, hojaPropuesta As Worksheet, posicion As Long, datosQ As Variant, filas As Collection)
    Dim hojaDatos As Worksheet, baseCol As Long, filaPoliza As Range, tipoArchivo As String, llave As String
    Dim i As Long, datoBruto As String, index As Long, edadesD As Object, rangoEdades As Range, filaEdad As Long
    Dim sexo As String, edad As Long, cantidad As Long, k As Variant
    Dim resultado As Long, incremento As Double, filaBase As Long, colIn As Long, colFin As Long, colIncremento As Long
    Dim valorCelda As String, colOrigen As Long
    
    ' === Intentar detectar el bloque dinámico del censo (E:F, filas) ===
    Dim rngCenso As Range, inicioFila As Long, finFila As Long, bins As Long
    Set rngCenso = ObtenerRangoCensoPorTitulo(hojaPropuesta, "Tabla", 2, 2, "RANGO_CENSO")
    If rngCenso Is Nothing Then
        Set rngCenso = RangoCensoPorAgencia(hojaPropuesta, libroCopia)
    End If
    
    If Not rngCenso Is Nothing Then
        inicioFila = rngCenso.Row
        finFila = rngCenso.Row + rngCenso.Rows.Count - 1
        bins = rngCenso.Rows.Count
        Debug.Print "Usando rango dinámico: "; rngCenso.Address(False, False); " | bins=" & bins
    Else
        ' Fallback final: tu lógica original fija con "ajuste"
        Dim ajuste As Long
        If Trim(hojaPropuesta.Cells(37, 2).Value) = "" Then
            ajuste = 1
        Else
            ajuste = 0
        End If
        inicioFila = 37 + ajuste
        finFila = 55 + ajuste
        bins = finFila - inicioFila + 1
        Debug.Print "Fallback original 37..55 (ajuste=" & ajuste & ") | bins=" & bins
    End If
    
    '-- 1) Base de columnas por subgrupo, anclamos a su primera columna (debe ser E=5); si no, E=5.
    If Not rngCenso Is Nothing Then
        baseCol = rngCenso.Columns(1).Column + (8 * (posicion - 1))
    Else
        baseCol = 5 + (8 * (posicion - 1)) ' E = 5
    End If
    Debug.Print "Subgrupo: "; posicion; " | BaseCol: "; baseCol; " | Filas: "; inicioFila; "-"; finFila
    
    '-- 2) Crear mapa de edades (lee etiquetas en Columna B entre inicioFila y finFila)
    Set edadesD = CreateObject("Scripting.Dictionary")
    Set rangoEdades = hojaPropuesta.Range("B" & inicioFila & ":B" & finFila)
    
    For i = 1 To rangoEdades.Rows.Count
        valorCelda = Trim$(CStr(rangoEdades.Cells(i, 1).Value))
        If Len(valorCelda) > 0 And InStr(1, valorCelda, "-", vbTextCompare) > 0 Then
            ' Usa el límite inferior como llave ("00","05","10"...)
            llave = Format$(CLng(Val(Split(valorCelda, "-")(0))), "00")
            If Not edadesD.Exists(llave) Then edadesD(llave) = rangoEdades.Cells(i, 1).Row
        End If
    Next i
    Debug.Print "Mapa de edades cargado: "; edadesD.Count; " | (esperado 16 o 18)"
        
    '-- 3) Limpiar la sección del subgrupo (solo E:F del bloque dinámico)
    hojaPropuesta.Range(hojaPropuesta.Cells(inicioFila, baseCol), hojaPropuesta.Cells(finFila, baseCol + 1)).ClearContents
    
    'Recorrer filas filtradas para este subgrupo y sumar en E (M) / F (F)
    For i = 1 To filas.Count
        index = filas(i)
        datoBruto = UCase$(Replace(CStr(datosQ(index, 4)), " ", "")) ' sexo en col D (termina en M/F)
        sexo = Right$(datoBruto, 1)
        edad = CLng(Val(datosQ(index, 5)))                           ' edad en col E
        cantidad = CLng(Val(datosQ(index, 7)))                       ' cantidad en col G
        llave = Format$(edad, "00")
        ' Buscar edad en el diccionario (por límite inferior)
        If edadesD.Exists(llave) Then
            filaEdad = edadesD(llave)
            If sexo = "M" Then
                hojaPropuesta.Cells(filaEdad, baseCol).Value = hojaPropuesta.Cells(filaEdad, baseCol).Value + cantidad
            ElseIf sexo = "F" Then
                hojaPropuesta.Cells(filaEdad, baseCol + 1).Value = hojaPropuesta.Cells(filaEdad, baseCol + 1).Value + cantidad
            Else
                Debug.Print "sexo no reconocido: "; datoBruto
            End If
        Else
            Debug.Print "edad no mapeada: "; edad
        End If
    Next i
    Debug.Print "Censo completado para subgrupo: " & posicion & " | bins=" & bins
    
    '-- 4) Colocar incremento maximo (tu lógica, sin cambios)
    Set hojaDatos = libroCopia.Sheets(datos)
    tipoArchivo = UCase(libroCopia.name)
    If InStr(1, tipoArchivo, "EMPR1") > 0 Then
        filaBase = 30         ' fila donde se limpia y coloca el incremento
        colOrigen = 3
        colIncremento = 7
        colIn = 6
        colFin = 7
        resultado = 8 * (posicion - 1)
        ' Limpiar las celdas desde G:30
        hojaPropuesta.Range(hojaPropuesta.Cells(filaBase, colFin + resultado), hojaPropuesta.Cells(filaBase, colFin + resultado)).ClearContents
        hojaPropuesta.Range(hojaPropuesta.Cells(filaBase, colIn + resultado), hojaPropuesta.Cells(filaBase, colIn + resultado)).ClearContents
    End If
End Sub
'-----------------------------------------------------------------------------------------------------------------------------
Private Function RangoCensoPorAgencia(ByVal hoja As Worksheet, ByVal wb As Workbook) As Range
    On Error GoTo falla
    Dim tipo As String, anio As Long, rIni As Long, rFin As Long
    tipo = DetectarTipoDesdeNombre(wb.name)
    anio = ExtraerAnio4DeNombre(wb.name)
    Select Case UCase$(tipo)
        Case "EMPR1":          rIni = 37:  rFin = 55   ' 18
        Case "EMPR2":           rIni = 38:  rFin = 53   ' 16
        Case "EMPR3":           rIni = 29: rFin = 47   ' 18
        Case "EMPR4":           rIni = 34: rFin = 50   ' 16
            rIni = 0: rFin = 0
        Case Else
            rIni = 0: rFin = 0
    End Select
    
    If rIni > 0 And rFin >= rIni Then
        Set RangoCensoPorAgencia = hoja.Range("E" & rIni & ":F" & rFin)
        Debug.Print "[Fallback Agencia] "; tipo; " Año=" & IIf(anio = 0, "N/D", anio) & _
                    " -> " & RangoCensoPorAgencia.Address(False, False)
    Else
        Set RangoCensoPorAgencia = Nothing
    End If
    Exit Function
falla:
    Set RangoCensoPorAgencia = Nothing
End Function
'-----------------------------------------------------------------------------------------------------------------------------
Private Function ExtraerAnio4DeNombre(ByVal nombre As String) As Long
    Dim i As Long, t As String, n As Long
    For i = 1 To Len(nombre) - 3
        t = Mid$(nombre, i, 4)
        If IsNumeric(t) Then
            n = CLng(t)
            If n >= 2020 And n <= 2035 Then
                ExtraerAnio4DeNombre = n
                Exit Function
            End If
        End If
    Next i
    ExtraerAnio4DeNombre = 0
End Function


