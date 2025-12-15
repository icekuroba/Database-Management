Attribute VB_Name = "Configuraciones"
'==========================================================================================
'                         Funciones auxiliares
'==========================================================================================
Public libroContrasena As Workbook
Public Sub Desbloquear(ByVal libroOrigen As Workbook)
    Dim ruta As String, lista As Worksheet
    Dim hojaIgnorar As Variant, ws As Worksheet
    'control de flujo
    Dim contr As String, j As Long, ultimaFila As Long, desbloqueado As Boolean

    'hojaIgnorar = Array("CAT", "Cat_Texto9s")
    desbloqueado = False
    ruta = "https://comunidadunammx.sharepoint.com/:x:/r/sites/KUROBACORPORATION/Proyectos/Cotizador/Parametros.xlsx"
    'ruta = "C:\Users\SSD\Desktop\Corporativo\Cotizador\Parametros.xlsx"

    'Abre libro de contraseñas
    Set libroContrasena = Workbooks.Open(ruta, UpdateLinks:=0, ReadOnly:=True)
    Set lista = libroContrasena.Sheets(2)
    ultimaFila = lista.Cells(lista.Rows.Count, "B").End(xlUp).Row

    'Probar cada contraseña
    For i = 1 To ultimaFila
        contr = Trim(CStr(lista.Cells(i, 2).Value))
        If contr <> "" Then
            Debug.Print "Probando contraseña en fila" & i & ": " & contr
            On Error Resume Next
            libroOrigen.Unprotect Password:=contr
            'Desbloquear cada hoja
            For Each ws In libroOrigen.Worksheets
                If Not EstaEnArray(ws.name, hojaIgnorar) Then
                    ws.Unprotect Password:=contr
                    ws.Visible = xlSheetVisible
                End If
            Next ws
            On Error GoTo 0
            'Valida si se desbloqueo
            If Not libroOrigen.ProtectStructure Then
                'MsgBox "Libro desbloqueado", vbInformation
                Debug.Print "libro desbloqueado con " & contr
                desbloqueado = True
                Exit For
            End If
        End If
    Next i

    If Not desbloqueado Then
        Debug.Print "No se puede desbloquear el libro"
        'MsgBox "No se puede desbloquear porque la contraseña no se encuentra", vbInformation
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
'===========================================================================================
'                La ruta donde se guardara la carpeta con las polizas
'===========================================================================================
Public Function rutaDocumentos() As String
    Dim base As String, up As String
    up = Environ("USERPROFILE")
    base = Environ("OneDrive")
    If Len(base) = 0 Then base = Environ("OneDriveEmpresa")
    If Len(base) > 0 Then
        If Dir(base & "\Documentos", vbDirectory) <> "" Then
            rutaDocumentos = base & "\Documentos": Exit Function
        ElseIf Dir(base & "\Documents", vbDirectory) <> "" Then
            rutaDocumentos = base & "\Documents": Exit Function
        End If
    End If

    Dim f As Object, fs As Object
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
        Next f
    End If

    If Dir(up & "\Documents", vbDirectory) <> "" Then
        rutaDocumentos = up & "\Documents"
    Else
        rutaDocumentos = up
    End If
    Debug.Print "libroContraseña cargado: " & libroContrasena.name
End Function
'===============================================================================================================
'                          Carga de arrays
'===============================================================================================================
Public Function LeerPoliza(ws As Worksheet, Optional col As Long = 2, Optional filaIn As Long = 9) As Boolean
    Dim ult As Long, i As Long, t As String, n As Long

    LeerPoliza = False
    If ws Is Nothing Then Exit Function
    ult = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
        Debug.Print "Depuracion de leer polizas"
        Debug.Print "columna:" & col & "(" & ws.Cells(1, col).Address & ")"
        Debug.Print "fila:" & filaIn
        Debug.Print "Ultima fila:" & ult
    For i = filaIn To ult
        Debug.Print "Fila" & i; ": [" & ws.Cells(i, col).Value & "]"
    Next i
    
    '1.- Contar
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

    '2.- Redimensionar
    ReDim polizas(1 To n)
    ReDim filas(1 To n)

    '3.- Llenar
    n = 0
    For i = filaIn To ult
        t = Normaliza(ws.Cells(i, col).Value)
        If Len(t) = 0 Then Exit For
        n = n + 1
        polizas(n) = t
        filas(n) = i
        Debug.Print "polizas detectadas"
    Next i
    nPolizas = n
    LeerPoliza = True
    Debug.Print "Total de polizas encontradas: " & nPolizas
End Function
'===============================================================================================================
'                  Funciones aún más auxiliares
'===============================================================================================================
Public Function ExisteHoja(wsName As String, wb As Workbook) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(wsName)
    ExisteHoja = Not ws Is Nothing
End Function
Public Function nombresSinExtension(ByVal fullPath As String) As String
    Dim nm As String
    nm = Mid$(fullPath, InStrRev(fullPath, "\") + 1)
    If InStrRev(nm, ".") > 0 Then nm = Left$(nm, InStrRev(nm, ".") - 1)
    nombresSinExtension = nm
End Function
Public Function LimpiarArchivo(ByVal s As String) As String
    Dim bad As Variant, a As Long
    bad = Array("/", "\", ":", "*", "?", """", "<", ">", "|", "!")
    For a = LBound(bad) To UBound(bad)
        s = Replace(s, bad(a), "_")
    Next a
    LimpiarArchivo = Trim$(s)
End Function
Private Function Normaliza(ByVal s As String) As String
    Dim t As String
    t = CStr(s)
    t = Trim$(s)
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    Normaliza = t
End Function
Public Function ObtenerMes(ByVal token As String) As String
    Dim abreviatura As String
    abreviatura = UCase(Left(Trim(token), 3))
    Select Case abreviatura
        Case "ENE", "ENE", "JAN": ObtenerMes = "Ene"
        Case "FEB": ObtenerMes = "Feb"
        Case "MAR": ObtenerMes = "Mar"
        Case "ABR": ObtenerMes = "Abr"
        Case "MAY": ObtenerMes = "May"
        Case "JUN": ObtenerMes = "Jun"
        Case "JUL": ObtenerMes = "Jul"
        Case "AGO": ObtenerMes = "Ago"
        Case "SEP": ObtenerMes = "Sep"
        Case "OCT": ObtenerMes = "Oct"
        Case "NOV": ObtenerMes = "Nov"
        Case "DIC", "DEC": ObtenerMes = "Dic"
        Case Else: ObtenerMes = ""
    End Select
End Function
Public Sub ToggleRibbon(ByVal Mostrar As Boolean)
    On Error Resume Next
    If Mostrar Then
        Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", True)"
    Else
        Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", False)"
    End If
    On Error GoTo 0
End Sub
Public Sub RestaurarInterfaz()
    Application.DisplayFormulaBar = True
    Application.DisplayStatusBar = True
    Call ToggleRibbon(True)
    Application.Visible = True
End Sub

