'================================================================================================================================
'                                           Funciones auxiliares
'================================================================================================================================
Public libroContrasena As Workbook
Option Explicit
Function ObtenerExcelLocal() As String
    Dim fso As Object, rutaLocal As String, carpetaApp As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    carpetaApp = Environ("LOCALAPPDATA") & "\Quotix"
    
    If Not fso.FolderExists(carpetaApp) Then
        fso.CreateFolder carpetaApp
    End If
    
    rutaLocal = carpetaApp & "\" & ThisWorkbook.name
    
    If InStr(1, ThisWorkbook.path, "http", vbTextCompare) > 0 Or Not fso.FileExists(ThisWorkbook.FullName) Then
        ThisWorkbook.SaveCopyAs rutaLocal
        ObtenerExcelLocal = rutaLocal
    Else
        ObtenerExcelLocal = ThisWorkbook.FullName
    End If
End Function
Public Sub crearAccesoDirecto(Optional dummy As Boolean)
    Dim libroShell As Object, fso As Object, acceso As Object, iconoDisponible As Boolean
    Dim rutaExcel As String, carpetaBase As String, carpetaImg As String, rutaOrigen As String
    Dim rutaIcono As String, rutaAcceso As String, escritorio As String, nombreAcceso As String
    
    On Error GoTo manejoErrores
    Set libroShell = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")
    rutaOrigen = ThisWorkbook.path
    iconoDisponible = False
    
    'Ruta del archivo actual (para que se cree en la carpeta donde se guarda COTIA)
    rutaExcel = ObtenerExcelLocal
    carpetaBase = fso.GetParentFolderName(rutaExcel)
    
    'Ruta para crear el acceso directo en escritorio
    escritorio = libroShell.SpecialFolders("Desktop")
    nombreAcceso = fso.GetBaseName(rutaExcel)
    rutaAcceso = escritorio & "\" & nombreAcceso & ".lnk"  'rutaAcceso = escritorio & "\COTIA.lnk"
    
    'Rutas automáticas
    carpetaImg = carpetaBase & "\img"
    rutaIcono = carpetaImg & "\icono.ico"
    
    If Not fso.FolderExists(carpetaImg) Then
        fso.CreateFolder carpetaImg
    End If
    
    'Copiar icono si no existe
    If Not fso.FileExists(rutaIcono) Then
        If fso.FileExists(carpetaBase & "\icono.ico") Then
            fso.copyfile carpetaBase & "\icono.ico", rutaIcono, True
            iconoDisponible = True
        Else
            iconoDisponible = False
        End If
    End If
    
    'Crea acceso directo
    If fso.FileExists(rutaAcceso) Then
        Debug.Print "Ya existe el acceso directo: "; rutaAcceso & "en la ruta del archivo" & rutaExcel
        Set acceso = libroShell.CreateShortcut(rutaAcceso)
        If LCase(acceso.TargetPath) = LCase(rutaExcel) Then
            Exit Sub
        End If
    End If
        
    Set acceso = libroShell.CreateShortcut(rutaAcceso)
    With acceso
        .TargetPath = rutaExcel
        .WorkingDirectory = carpetaBase
        .Description = "Acceso directo al sistema Quotix" '.IconLocation = rutaIcono
        If Dir(rutaIcono) <> "" Then .IconLocation = rutaIcono & ",0"
        .Save
    End With
    MsgBox "Acceso directo creado en el escritorio.", vbInformation
    Exit Sub
    
manejoErrores:
    Debug.Print "error"; Err.Description
    MsgBox "error" & Err.Description, vbCritical
End Sub
'=================    Desbloquar el libro de cotizador (la macro) por hojas y libro ========================================
Public Sub Desbloquear(ByVal libroOrigen As Workbook)
    Dim ruta As String, lista As Worksheet
    Dim hojaIgnorar As Variant, ws As Worksheet
    ' -- control de flujo
    Dim contr As String, i As Long, ultimaFila As Long, desbloqueado As Boolean
    
    hojaIgnorar = Array("CAT", "Cat_Textos")
    desbloqueado = False
    ruta = "https://comunidadunammx.sharepoint.com/:x:/r/sites/KUROBACORPORATION/Proyectos/Cotizador/Parametros.xlsx"
        '   Abre libro de contraseña
    Set libroContrasena = Workbooks.Open(ruta, UpdateLinks:=0, ReadOnly:=True)
    Set lista = libroContrasena.Sheets(3)
    ultimaFila = lista.Cells(lista.Rows.Count, "B").End(xlUp).Row
    
        '   Probar cada contraseña
    For i = 1 To ultimaFila
        contr = Trim(CStr(lista.Cells(i, 2).Value))
        If contr <> "" Then
            Debug.Print "Probando contraseña en fila" & i & ": " & contr
        On Error Resume Next
        libroOrigen.Unprotect Password:=contr
        '   Desbloquear cada hoja
        For Each ws In libroOrigen.Worksheets
            If Not EstaEnArray(ws.name, hojaIgnorar) Then
                ws.Unprotect Password:=contr
                ws.Visible = xlSheetVisible
            End If
        Next ws
        On Error GoTo 0
        '   Valida si se desbloqueo
        If Not libroOrigen.ProtectStructure Then
            'MsgBox "Libro desbloqueado", vbInformation
            Debug.Print "libro desbloqueado con: " & contr
            desbloqueado = True
            Exit For
            End If
        End If
        Next i
        
    If Not desbloqueado Then
            Debug.Print "No se puede desbloquear el libro"
        'MsgBox "No se puede desbloquear porque la contraseÃ±a no se encuentra", vbInformation
    End If
End Sub
Private Function EstaEnArray(valor As String, arr As Variant) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If StrComp(valor, arr(i), vbTextCompare) = 0 Then
            EstaEnArray = True
            Exit Function
        End If
    Next i
    EstaEnArray = False
End Function
'   =======    La ruta donde se guardara la carpeta con las polizas   =======
Public Function rutaDocumentos() As String
    Dim base As String, up As String, f As Object, fs As Object
    up = Environ$("USERPROFILE")
    base = Environ$("OneDriveCommercial")
    If Len(base) = 0 Then base = Environ$("OneDrive")
        If Len(base) > 0 Then
            If Dir(base & "\Documentos", vbDirectory) <> "" Then
                rutaDocumentos = base & "\Documentos": Exit Function
            ElseIf Dir(base & "\Documents", vbDirectory) <> "" Then
                rutaDocumentos = base & "\Documents": Exit Function
        End If
    End If
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FolderExists(up) Then
        For Each f In fs.GetFolder(up).SubFolders
            If LCase$(Left$(f.name, 10)) = LCase$("OneDrive -") Then
                If fs.FolderExists(f.path & "\Documentos") Then
                    rutaDocumentos = f.path & "\Documentos": Exit Function
                ElseIf fs.FolderExists(f.path & "\Documents") Then
                    rutaDocumentos = f.path & "\Documents": Exit Function
            End If
        End If
        Next
    End If
        
    If Dir(up & "\Documents", vbDirectory) <> "" Then
        rutaDocumentos = up & "\Documents"
    Else
        rutaDocumentos = up
        Debug.Print "libroContraseña cargado: " & libroContrasena.name
    End If
End Function
'   ====================            Carga de arrays               ==============================
Public Function LeerPoliza(ws As Worksheet, Optional col As Long = 2, Optional filaIn As Long = 9) As Boolean
    Dim ult As Long, i As Long, t As String, n As Long
    
    LeerPoliza = False
    If ws Is Nothing Then Exit Function
    ult = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
        Debug.Print "Depuracion de leer polizas"
        Debug.Print "columna: " & col & " (" & ws.Cells(1, col).Address & ")"
        Debug.Print "fila :" & filaIn
        Debug.Print "Ultima fila: " & ult
    For i = filaIn To ult
        Debug.Print "Fila " & i & ": [" & ws.Cells(i, col).Value & "]"
    Next i
    
    '1- Contar
    n = 0
    For i = filaIn To ult
        t = Normaliza(ws.Cells(i, col).Value)
        If Len(t) = 0 Then Exit For
        n = n + 1
    Next i
    If n = 0 Then
        Debug.Print "No pues no funciona"
        Exit Function
    End If
    
    '2- Redimensionar
    ReDim polizas(1 To n)
    ReDim filas(1 To n)
        
    '3- Llenar
    n = 0
    For i = filaIn To ult
    t = Normaliza(ws.Cells(i, col).Value)
    If Len(t) = 0 Then Exit For
        n = n + 1
        polizas(n) = t
        filas(n) = i
        Debug.Print "polizas detectadas"
    Next
        nPolizas = n
        LeerPoliza = True
        Debug.Print "Total de polizas encontradas: " & nPolizas
End Function
Public Function nombreSinExtension(ByVal fullPath As String) As String
    Dim nm As String
    nm = Mid$(fullPath, InStrRev(fullPath, "\") + 1)
    If InStrRev(nm, ".") > 0 Then nm = Left$(nm, InStrRev(nm, ".") - 1)
        nombreSinExtension = nm
End Function

'***************  Helper para restaurar interfaz y ocultar la hoja de excel de fondo  *************
Public Sub ToggleRibbon(ByVal Mostrar As Boolean)
    On Error Resume Next
    If Mostrar Then
        Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", True)"
    Else
        Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", False)"
    End If
    On Error GoTo 0
End Sub
Public Sub RestaurarInterfaz(Optional dummy As Boolean)
    Application.DisplayFormulaBar = True
    Application.DisplayStatusBar = True
    Call ToggleRibbon(True)
    Application.Visible = True
End Sub
' ===========================       Funciones para Pool GMM        =======================
Public Function ExisteHoja(wsName As String, wb As Workbook) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(wsName)
    ExisteHoja = Not ws Is Nothing
    On Error GoTo 0
End Function
Public Function LimpiarArchivo(ByVal s As String) As String
    Dim bad As Variant, a As Long
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "!")
    For a = LBound(bad) To UBound(bad)
        s = Replace$(s, bad(a), "_")
    Next a
    LimpiarArchivo = Trim$(s)
End Function
Private Function Normaliza(ByVal s As String) As String
    Dim t As String
    t = CStr(s)
    t = Trim$(t)
    ' Colapsa espacios mÃºltiples en uno
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    ' Opcional: elimina espacios no separables
    t = Replace(t, Chr(160), " ")
    Normaliza = t
End Function
Public Function ObtenerMes(ByVal token As String) As String
    Dim abreviatura As String
    abreviatura = UCase(Left(Trim(token), 3))
        Select Case abreviatura
            Case "ENE", "JAN": ObtenerMes = "Ene"
            Case "FEB": ObtenerMes = "Feb"
            Case "MAR": ObtenerMes = "Mar"
            Case "ABR", "APR": ObtenerMes = "Abr"
            Case "MAY": ObtenerMes = "May"
            Case "JUN": ObtenerMes = "Jun"
            Case "JUL": ObtenerMes = "Jul"
            Case "AGO", "AUG": ObtenerMes = "Ago"
            Case "SEP": ObtenerMes = "Sep"
            Case "OCT": ObtenerMes = "Oct"
            Case "NOV": ObtenerMes = "Nov"
            Case "DIC", "DEC": ObtenerMes = "Dic"
            Case Else: ObtenerMes = ""
        End Select
End Function
'Normalizacion y Deteccion de Tipo
Public Function NormalizarTexto(ByVal s As String) As String
    Dim t As String
    t = UCase$(Trim$(s))
    t = Replace(t, vbTab, " ")
    t = Replace(t, "-", " ")
    t = Replace(t, "_", " ")
    t = Replace(t, ".", " ")
    t = Replace(t, "/", " ")
    t = Replace(t, "\", " ")
    t = Replace(t, "Ã", "A")
    t = Replace(t, "Ã‰", "E")
    t = Replace(t, "Ã", "I")
    t = Replace(t, "Ã“", "O")
    t = Replace(t, "Ãš", "U")
    t = Replace(t, "Ãœ", "U")
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    NormalizarTexto = t
End Function

' Detecta la clave canonica
Public Function DetectarTipoDesdeNombre(ByVal nombreArchivo As String) As String
    Dim base As String: base = NormalizarTexto(nombreSinExtension(nombreArchivo))
    Dim d As Object: Set d = DiccionarioAlias()
    Dim pr As Variant: pr = PrioridadBrokers()
    Dim candidatos As Object: Set candidatos = CreateObject("Scripting.Dictionary")
    candidatos.CompareMode = 1
    ' Tokens
    Dim tokens() As String, t As Variant
    tokens = Split(base, " ")
    For Each t In tokens
        If d.Exists(t) Then candidatos(d(t)) = True
    Next t
    ' Subcadenas (para los nombres cortos de agencias con espacios)
    Dim k As Variant
    For Each k In d.Keys
        If InStr(1, base, k, vbTextCompare) > 0 Then
            candidatos(d(k)) = True
        End If
    Next k
    Dim i As Long
    For i = LBound(pr) To UBound(pr)
        If candidatos.Exists(pr(i)) Then
            DetectarTipoDesdeNombre = pr(i)
            Exit Function
        End If
    Next i
    DetectarTipoDesdeNombre = ""
End Function

' Diccionarios de las agencias
Public Function PrioridadBrokers() As Variant
    PrioridadBrokers = Array("EMPR1", "EMPR2", "EMPR2", "EMPR3", "EMPR4")
End Function
' Nombres de las agencias para que esten en clave canonica
Public Function DiccionarioAlias() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    ' empresa1
    d("EMPRESA1") = "EMPR1"
    d("EMPR1") = "EMPR1"
    d("empresa1") = "EMPR1"
    ' empresa2
    d("EMPRESA2") = "EMPR2"
    d("EMPR2") = "EMPR2"
    d("empresa2") = "EMPR2"
    ' empresa3
    d("EMPRESA3") = "EMPR3"
    d("EMPR3") = "EMPR3"
    d("empresa3") = "EMPR3"
    ' empresa4
    d("EMPRESA4") = "EMPR4"
    d("EMPR4") = "EMPR4"
    d("empresa4") = "EMPR4"
    Set DiccionarioAlias = d
End Function

Public Function ObtenerRangoCensoPorTitulo(ByVal hoja As Worksheet, ByVal tituloRenov As String, Optional ByVal offsetInicio As Long = 2, _
    Optional ByVal offsetFin As Long = 2, Optional ByVal guardarNombre As String = vbNullString) As Range
    On Error GoTo Salir
    Dim fTitulo As Range, fCenso As Range, fSub As Range, area As Range, filaIni As Long, filaFin As Long
    Dim rng As Range, lastRow As Long, lastCol As Long
    
    If hoja Is Nothing Then Exit Function
    'Área usada
    lastRow = hoja.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    lastCol = hoja.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    
    '-- 1) Buscar título
    Set fTitulo = hoja.Cells.Find(What:=tituloRenov, LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, _
                    SearchDirection:=xlNext, MatchCase:=False)
                    
    If fTitulo Is Nothing Then
        Debug.Print "No se encontró título: "; tituloRenov
        Exit Function
    End If
    Debug.Print "Titulo encontrado en fila:", fTitulo.Row
    'área desde el título hacia abajo
    Set area = hoja.Range(hoja.Cells(fTitulo.Row, 1), hoja.Cells(lastRow, lastCol))
    
    '-- 2) Buscar CENSO
    Set fCenso = area.Find(What:="CENSO", LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                    MatchCase:=False)
    If fCenso Is Nothing Then
        Debug.Print "No se encontró 'CENSO'"
        Exit Function
    End If
    Debug.Print "CENSO encontrado en fila:", fCenso.Row
    
    '-- 4) Buscar SUBTOTAL-
    Set fSub = area.Find(What:="SUBTOTAL", LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, _
                    SearchDirection:=xlNext, MatchCase:=False)
    If fSub Is Nothing Then
        Set fSub = area.Find("SUB TOTAL", LookIn:=xlValues, LookAt:=xlPart)
    End If
    
    If fSub Is Nothing Then
        Debug.Print "No se encontró SUBTOTAL"
        Exit Function
    End If
    Debug.Print "SUBTOTAL encontrado en fila:", fSub.Row
    
    If fSub.Row <= fCenso.Row Then
        Debug.Print "SUBTOTAL está antes de CENSO"
        Exit Function
    End If
    
    '-- 5) calcular filas
    filaIni = fCenso.Row + offsetInicio
    filaFin = fSub.Row - offsetFin
    
    If filaFin < filaIni Then
        Debug.Print "Rango inválido"
        Exit Function
    End If
    
    '-- 6) rango de limpieza
    Set rng = hoja.Range("E" & filaIni & ":F" & filaFin)
    Debug.Print "Rango detectado:", rng.Address
    Set ObtenerRangoCensoPorTitulo = rng
    
    '-- 7) nombre opcional
    If guardarNombre <> "" Then
        On Error Resume Next
        hoja.Parent.Names(guardarNombre).Delete
        On Error GoTo 0
        hoja.Parent.Names.Add name:=guardarNombre, RefersTo:=rng
    End If
Salir:
    If Err.Number <> 0 Then
        Debug.Print "Error:", Err.Description
    End If
End Function

' ===========================       Funciones para Pool VIDA        =======================
Public Function ObtenerMesV(ByVal token As String) As String
    Dim abreviatura As String
    abreviatura = UCase(Left(Trim(token), 3))
        Select Case abreviatura
            Case "ENE", "JAN": ObtenerMesV = "enero"
            Case "FEB": ObtenerMesV = "febrero"
            Case "MAR": ObtenerMesV = "marzo"
            Case "ABR", "APR": ObtenerMesV = "abril"
            Case "MAY": ObtenerMesV = "mayo"
            Case "JUN": ObtenerMesV = "junio"
            Case "JUL": ObtenerMesV = "julio"
            Case "AGO", "AUG": ObtenerMesV = "agosto"
            Case "SEP": ObtenerMesV = "septiembre"
            Case "OCT": ObtenerMesV = "octubre"
            Case "NOV": ObtenerMesV = "noviembre"
            Case "DIC", "DEC": ObtenerMesV = "diciembre"
            Case Else: ObtenerMesV = ""
        End Select
End Function

Public Function TokenizarNombre(ByVal nombre As String) As Collection
    Dim col As New Collection, partes() As String, p As Variant, subPartes() As String, s As Variant
    nombre = Replace(nombre, vbTab, " ")
    nombre = Trim$(nombre)

    '-- 1) Primero por underscore
    partes = Split(nombre, "_")
    For Each p In partes
        Dim pieza As String
        pieza = Trim$(CStr(p))
        If Len(pieza) > 0 Then
            ' Guardar la pieza completa (para buscar "contains" en lista)
            col.Add pieza

    '-- 2) Expandir ademÃ¡s por espacios (para mes/aÃ±o/version)
            subPartes = Split(WorksheetFunction.Trim(pieza), " ")
            For Each s In subPartes
                If Len(Trim$(CStr(s))) > 0 Then col.Add Trim$(CStr(s))
            Next s
        End If
    Next p
    Set TokenizarNombre = col
End Function
Public Function nombreSinExtensionV(ByVal fileName As String) As String
    Dim pos As Long
    pos = InStrRev(fileName, ".")
    If pos > 0 Then
        nombreSinExtensionV = Left$(fileName, pos - 1)
    Else
        nombreSinExtensionV = fileName
    End If
End Function
Public Function ExtraerAnio(ByVal token As String) As String
    Dim t As String
    t = NormalizarTexto(token)
    Dim re As Object, m As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True

    '-- 1) Buscar 4 digitos
    re.Pattern = "(19\d{2}|20\d{2})"
    If re.Test(t) Then
        Set m = re.Execute(t)(0)
        ExtraerAnio = Right$(m.Value, 2)
        Exit Function
    End If

    '-- 2) Si no hay 4 digitos, buscar 2 diitos aislados
    re.Pattern = "(^|\D)(\d{2})(\D|$)"
    If re.Test(t) Then
        Set m = re.Execute(t)(0)
        ExtraerAnio = m.SubMatches(1)
        Exit Function
    End If
    ExtraerAnio = ""
End Function
Public Function DetectarRamo(ByVal token As String) As String
    Dim t As String
    t = NormalizarTexto(token)
    Select Case t
        Case "VG", "VIDA", "VIDAS", "LIFE"
            DetectarRamo = "VG"
        Case Else
            DetectarRamo = ""
    End Select
End Function
Public Function CoincideTokenAgente(ByVal tokenNorm As String, ByVal agenteNorm As String) As Boolean
    Dim t As String, a As String
    ' Asegurar espacios a los lados para match por palabra
    t = " " & Trim$(tokenNorm) & " "
    a = " " & Trim$(agenteNorm) & " "

    If Len(Trim$(agenteNorm)) <= 4 Then
        CoincideTokenAgente = (InStr(1, t, a, vbTextCompare) > 0) ' Codigos cortos: match por palabra completa
    Else
        CoincideTokenAgente = (InStr(1, tokenNorm, agenteNorm, vbTextCompare) > 0) ' Nombres largos: match por subcadena
    End If
End Function

Public Function SeleccionarCarpeta(Optional ByVal titulo As String = "Selecciona una carpeta") As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)

    With fd
        .Title = titulo
        .AllowMultiSelect = False
        If .Show = -1 Then
            SeleccionarCarpeta = .SelectedItems(1)
        Else
            SeleccionarCarpeta = ""
        End If
    End With
End Function

Public Function ContarExcelsEnCarpeta(ByVal carpeta As String) As Long
    Dim fso As Object, folder As Object, file As Object
    Dim ext As String, n As Long
    ContarExcelsEnCarpeta = 0
    If Len(Trim$(carpeta)) = 0 Then Exit Function

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(carpeta) Then Exit Function
    Set folder = fso.GetFolder(carpeta)
    For Each file In folder.Files
        ' Ignorar temporales de Excel
        If Left$(file.name, 2) <> "~$" Then
            ext = LCase$(fso.GetExtensionName(file.name))
            If ext = "xls" Or ext = "xlsx" Or ext = "xlsm" Then
                n = n + 1
            End If
        End If
    Next file
    ContarExcelsEnCarpeta = n
End Function
' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' ========= Opcion 2: Indexar carpeta por poliza =========
Public Function IndexarCarpetaPorPoliza(ByVal carpeta As String) As Object
    Dim fso As Object, folder As Object, file As Object
    Dim dict As Object, re As Object, ms As Object, m As Object
    Dim base As String, clave As String, al As Object

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(carpeta) Then
        Set IndexarCarpetaPorPoliza = dict
        Exit Function
    End If

    Set folder = fso.GetFolder(carpeta)

    'cualquier bloque de digitos. (Evita capturar "100" sueltos)
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "(\d{5,})"

    Dim fechaClave As String
    For Each file In folder.Files
        If LCase$(fso.GetExtensionName(file.name)) Like "xls*" Then
            If Left$(file.name, 2) <> "~$" Then
                base = UCase$(file.name)
                If re.Test(base) Then
                    Set ms = re.Execute(base)
                    For Each m In ms
                        clave = CStr(m.Value) ' cada grupo de digitos (NPOLIZA, Id, etc.)
                        If Not dict.Exists(clave) Then
                            dict.Add clave, CreateObject("System.Collections.ArrayList")
                        End If
                        fechaClave = Format$(file.DateLastModified, "yyyymmddhhnnss")
                        dict(clave).Add fechaClave & "|" & file.path
                    Next m
                End If
            End If
        End If
    Next

    ' Ordenar cada lista (descendente por fecha)
    Dim k As Variant
    For Each k In dict.Keys
        Set al = dict(k)
        If al.Count > 1 Then
            al.Sort
            al.Reverse
        End If
    Next

    Set IndexarCarpetaPorPoliza = dict
End Function
Public Function GetHoja(ByVal wb As Workbook, ByVal nombreHoja As String, Optional ByVal FallbackPrimera As Boolean = True) As Worksheet
    On Error Resume Next
    Set GetHoja = wb.Sheets(nombreHoja)
    On Error GoTo 0
    If GetHoja Is Nothing And FallbackPrimera Then
        Set GetHoja = wb.Sheets(1)
    End If
End Function
Public Function NombreContienePoliza(ByVal nombreArchivo As String, ByVal poliza As String) As Boolean
    Dim re As Object
    poliza = Trim$(CStr(poliza))
    If Len(poliza) = 0 Then Exit Function

    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.Pattern = "(^|\D)" & poliza & "(\D|$)"
    NombreContienePoliza = re.Test(UCase$(nombreArchivo))
End Function
