Attribute VB_Name = "Quinquenio"
'================================================================================================
'            Función para el libro de Quinquenios (ya abierto)
'================================================================================================
Public Const propuesta As String = "PROPUESTA"
Public Const datos As String = "DATOS"
Public numPolizaGlobal As String 'nombre de cada poliza

Public Sub Quinquenios(libroCopia As Workbook, archivoQ As String)
    Dim libroQ As Workbook, hojaPropuesta As Worksheet, hojaDatos As Worksheet, hojaQ As Worksheet
    Dim i As Long, j As Long, listaFiltrada() As Variant, datosQ As Variant
    Dim ultimaFila As Long, filaOut As Long, totalCols As Long, totalFilas As Long
    Dim valorCelda As String, fila As Long, numPolizaActual As String
    Dim grupos As Object, llaves As Variant, clave As Variant, subg As Variant, pos As Long

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    Set hojaPropuesta = libroCopia.Sheets(propuesta)
    Set hojaDatos = libroCopia.Sheets(datos)

    ' === 1) Cargar datos del archivo de quinquenios ===
    Set libroQ = Workbooks.Open(archivoQ, ReadOnly:=True)
    Set hojaQ = libroQ.Sheets(1)
    ultimaFila = hojaQ.Cells(hojaQ.Rows.Count, "B").End(xlUp).Row
    If ultimaFila < 2 Then
        Debug.Print "El archivo no tiene datos"
        libroQ.Close False
        GoTo Salir
    End If
    datosQ = hojaQ.Range("A2:G" & ultimaFila).Value

    ' === 2) Filtrar solo la póliza actual ===
    numPolizaActual = numPolizaGlobal
    totalFilas = UBound(datosQ, 1)
    totalCols = UBound(datosQ, 2)
    filaOut = 0

    ' Contar coincidencias
    For i = 1 To totalFilas
        valorCelda = Trim$(CStr(datosQ(i, 2)))
        If CLng(Val(valorCelda)) = CLng(Val(numPolizaActual)) Then
            filaOut = filaOut + 1
        End If
    Next i

    If filaOut = 0 Then
        Debug.Print "No hay registros en quinquenios para la póliza ¯\_(¬_¬)_/¯ " & numPolizaActual
        GoTo Salir
    End If

    ' === 3) Copiar coincidencias ===
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
    Debug.Print filaOut; " registros encontrados en la póliza " & numPolizaActual

    ' === 4) Contar subgrupos ===
    Set grupos = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(listaFiltrada, 1)
        clave = CLng(Val(listaFiltrada(i, 3))) ' Columna C = NSUB
        If Not grupos.exists(clave) Then
            Set grupos(clave) = New Collection
        End If
        grupos(clave).Add i ' <- Guarda el índice del arreglo, no el valor (CORRECCIÓN CLAVE)
    Next i
    Debug.Print "Subgrupos detectados: "; grupos.Count

    ' === 5) Procesar cada subgrupo ===
    llaves = grupos.Keys
    Call SortVariantArray(llaves)
    pos = 0

    For Each subg In llaves
        pos = pos + 1
        Debug.Print "Llenando subgrupo " & subg & " en posición " & pos
        Call LlenarCenso(libroCopia, hojaPropuesta, pos, listaFiltrada, grupos(subg))
    Next subg

Salir:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
End Sub


'************************************************************************************************************
'                 Subfunciones de la función de quinquenios
'************************************************************************************************************
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


'******************************************************************************************************************
'                           Llenar censo en propuesta (versión sin error)
'******************************************************************************************************************
Private Sub LlenarCenso(libroCopia As Workbook, hojaPropuesta As Worksheet, posicion As Long, datosQ As Variant, filas As Collection)
    Dim hojaDatos As Worksheet, baseCol As Long, tipoArchivo As String, llave As String
    Dim i As Long, index As Long, datoBruto As String, filaEdad As Long
    Dim edadesD As Object, rangoEdades As Range, valorCelda As String
    Dim sexo As String, edad As Long, cantidad As Long
    Dim incremento As Double, filaBase As Long, colOrigen As Long, colIncremento As Long, resultado As Long

    ' === 1. Determinar columna base según el subgrupo ===
    baseCol = 5 + (8 * (posicion - 1)) ' E = 5
    Debug.Print "Subgrupo:"; posicion; " BaseCol:"; baseCol

    ' === 2. Crear mapa de edades ===
    Dim inicioFila As Long, finFila As Long
    inicioFila = 37
    finFila = 50
    Set edadesD = CreateObject("Scripting.Dictionary")
    Set rangoEdades = hojaPropuesta.Range("B" & inicioFila & ":B" & finFila)

    For i = 1 To rangoEdades.Rows.Count
        valorCelda = Trim$(CStr(rangoEdades.Cells(i, 1).Value))
        If Len(valorCelda) > 0 And InStr(valorCelda, "-") > 0 Then
            llave = Format$(CLng(Val(Split(valorCelda, "-")(0))), "00")
            If Not edadesD.exists(llave) Then edadesD(llave) = rangoEdades.Cells(i, 1).Row
        End If
    Next i
    Debug.Print "Mapa de edades cargado:"; edadesD.Count

    ' === 3. Limpiar celdas previas de ese bloque ===
    hojaPropuesta.Range(hojaPropuesta.Cells(inicioFila, baseCol), hojaPropuesta.Cells(finFila, baseCol + 1)).ClearContents

    ' === 4. Llenar los datos del censo ===
    For i = 1 To filas.Count
        index = filas(i) ' <- ahora es un índice válido (ej. 1, 2, 3)
        datoBruto = UCase$(Replace(CStr(datosQ(index, 4)), " ", "")) ' Col D = código tipo
        sexo = Right$(datoBruto, 1)
        edad = CLng(Val(datosQ(index, 5))) ' Col E = edad
        cantidad = CLng(Val(datosQ(index, 7))) ' Col G = cantidad
        llave = Format$(edad, "00")

        If edadesD.exists(llave) Then
            filaEdad = edadesD(llave)
            If sexo = "M" Then
                hojaPropuesta.Cells(filaEdad, baseCol).Value = hojaPropuesta.Cells(filaEdad, baseCol).Value + cantidad
            ElseIf sexo = "F" Then
                hojaPropuesta.Cells(filaEdad, baseCol + 1).Value = hojaPropuesta.Cells(filaEdad, baseCol + 1).Value + cantidad
            End If
        Else
            Debug.Print "Edad no mapeada:"; edad
        End If
    Next i

    Debug.Print "Censo completado para subgrupo:"; posicion

    ' === 5. Colocar incremento máximo (si aplica) ===
    Set hojaDatos = libroCopia.Sheets(datos)
    tipoArchivo = UCase$(libroCopia.name)

    If InStr(1, tipoArchivo, "A", vbTextCompare) > 0 Then
        filaBase = 30
        colOrigen = 3
        colIncremento = 7
        resultado = 8 * (posicion - 1)

        incremento = hojaPropuesta.Cells(filaBase, colOrigen + resultado).Value
        If IsNumeric(incremento) Then
            incremento = WorksheetFunction.Round(incremento * 100, 0) / 100
            hojaPropuesta.Cells(filaBase, colIncremento + resultado).Value = incremento
            hojaPropuesta.Cells(filaBase, colIncremento + resultado).NumberFormat = "0.00%"
            Debug.Print "Incremento max asignado (" & numPolizaGlobal & "): " & incremento
        Else
            Debug.Print "Incremento no numérico en subgrupo"; posicion
        End If
    End If
End Sub

