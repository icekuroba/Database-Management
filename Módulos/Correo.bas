Attribute VB_Name = "Correo"
'===========================================
'  Módulo: Envío automático de pólizas por correo
'  Versión robusta con control de índices y depuración
'===========================================

Public Const correos As String = "TablaCorreos"
Public polizasProcesadas() As String
Public rutasProcesadas() As String
Public contadorPolizas As Long

'-------------------------------------------
' Inicializa los arreglos de registro
'-------------------------------------------
Public Sub IniciarRegistros(Optional dummy As Boolean)
    ReDim polizasProcesadas(1 To 1)
    ReDim rutasProcesadas(1 To 1)
    contadorPolizas = 0
    Debug.Print "Registro de pólizas iniciado"
End Sub

'-------------------------------------------
' Registra las pólizas procesadas
'-------------------------------------------
Public Sub RegistrarPolizas(ByVal numPoliza As String, ByVal rutaArchivo As String)
    If Len(numPoliza) = 0 Or Len(rutaArchivo) = 0 Then Exit Sub

    contadorPolizas = contadorPolizas + 1
    ReDim Preserve polizasProcesadas(1 To contadorPolizas)
    ReDim Preserve rutasProcesadas(1 To contadorPolizas)

    polizasProcesadas(contadorPolizas) = Trim(numPoliza)
    rutasProcesadas(contadorPolizas) = rutaArchivo

    Debug.Print "Registrada #" & contadorPolizas & " -> " & numPoliza & " | Ruta: " & rutaArchivo
End Sub

'-------------------------------------------
' Enviar correos automáticos desde Outlook
'-------------------------------------------
Public Sub EnviarCorreo(Optional dummy As Boolean)
    On Error GoTo ManejarError

    Dim base As Worksheet, i As Long, ultimaFila As Long
    Dim numPoliza As String, correoD As String
    Dim ejecutivo As String, gerente As String
    Dim idxPoliza As Object, oApp As Object, oMail As Object
    Dim archivo As String, cuenta As Object, encontrada As Boolean
    Dim nombreCuenta As String, correoGerente As String, ccList As String
    Dim datos() As String, j As Long

    '1?? Validar si hay pólizas registradas
    If contadorPolizas = 0 Then
        MsgBox "No hay pólizas registradas para enviar.", vbExclamation
        Exit Sub
    End If

    '2?? Validar integridad de los arreglos
    If Not IsArray(polizasProcesadas) Or Not IsArray(rutasProcesadas) Then
        MsgBox "Error: los arreglos de pólizas no están inicializados.", vbCritical
        Exit Sub
    End If
    If contadorPolizas > UBound(polizasProcesadas) Then
        Debug.Print "? Ajustando tamaño de arreglos (contador=" & contadorPolizas & _
                    ", límite=" & UBound(polizasProcesadas) & ")"
        ReDim Preserve polizasProcesadas(1 To contadorPolizas)
        ReDim Preserve rutasProcesadas(1 To contadorPolizas)
    End If

    '3?? Cargar la tabla de correos
    If Not HojaExiste(correos, ThisWorkbook) Then
        MsgBox "No se encontró la hoja '" & correos & "'.", vbCritical
        Exit Sub
    End If

    Set base = ThisWorkbook.Sheets(correos)
    ultimaFila = base.Cells(base.Rows.Count, "A").End(xlUp).Row
    Set idxPoliza = CreateObject("Scripting.Dictionary")

    For i = 2 To ultimaFila
        Dim pol As String, ej As String, cor As String, ger As String
        pol = Trim(CStr(base.Cells(i, "A").Value))
        ej = Trim(CStr(base.Cells(i, "B").Value))
        cor = Trim(CStr(base.Cells(i, "C").Value))
        ger = Trim(CStr(base.Cells(i, "D").Value))
        If Len(pol) > 0 And Len(cor) > 0 Then
            idxPoliza(pol) = ej & "|" & cor & "|" & ger
        End If
    Next i

    Debug.Print "Total de registros cargados en diccionario: " & idxPoliza.Count

    '4?? Crear instancia de Outlook
    On Error Resume Next
    Set oApp = GetObject(, "Outlook.Application")
    If oApp Is Nothing Then Set oApp = CreateObject("Outlook.Application")
    On Error GoTo ManejarError

    '5?? Procesar y enviar
    For i = 1 To contadorPolizas
        On Error Resume Next
        numPoliza = Trim(polizasProcesadas(i))
        archivo = Trim(rutasProcesadas(i))
        On Error GoTo ManejarError

        Debug.Print "? Procesando póliza índice " & i & " / " & contadorPolizas & ": " & numPoliza

        If Len(numPoliza) = 0 Then GoTo siguiente
        If Not idxPoliza.exists(numPoliza) Then
            Debug.Print "? No se encontró la póliza en tabla de correos: " & numPoliza
            GoTo siguiente
        End If

        datos = Split(idxPoliza(numPoliza), "|")
        ejecutivo = ""
        correoD = ""
        gerente = ""

        If UBound(datos) >= 0 Then ejecutivo = Trim(datos(0))
        If UBound(datos) >= 1 Then correoD = Trim(datos(1))
        If UBound(datos) >= 2 Then gerente = Trim(datos(2))

        If Len(correoD) = 0 Then
            Debug.Print "? Sin correo asociado para póliza " & numPoliza
            GoTo siguiente
        End If

        ' Buscar correo del gerente
        correoGerente = ""
        For j = 2 To ultimaFila
            If UCase(Trim(base.Cells(j, "B").Value)) = UCase(gerente) Then
                correoGerente = Trim(base.Cells(j, "C").Value)
                Exit For
            End If
        Next j

        ccList = "ice.kuroba@gmail.com"
        If Len(correoGerente) > 0 Then ccList = ccList & "," & correoGerente

        ' Crear correo
        Set oMail = oApp.CreateItem(0)

        ' Seleccionar cuenta de envío
        nombreCuenta = "malinallisag.3@gmail.com"
        encontrada = False
        For Each cuenta In oApp.Session.Accounts
            If LCase(cuenta.SmtpAddress) = LCase(nombreCuenta) Then
                Set oMail.SendUsingAccount = cuenta
                Debug.Print "Usando cuenta: " & cuenta.SmtpAddress
                encontrada = True
                Exit For
            End If
        Next cuenta

        If Not encontrada Then
            On Error Resume Next
            Set oMail.SendUsingAccount = oApp.Session.Accounts.Item(1)
            Debug.Print "Cuenta predeterminada: " & oMail.SendUsingAccount.SmtpAddress
            On Error GoTo ManejarError
        End If

        '6?? Redactar correo
        With oMail
            .To = correoD
            .CC = ccList
            .Subject = "Propuesta de Renovación"
            .Body = "Estimad(a)," & vbCrLf & vbCrLf & _
                    "Adjunto la propuesta bajo los mismos términos y condiciones para la póliza correspondiente." & vbCrLf & vbCrLf & _
                    "Saludos," & vbCrLf & "Equipo KUROBA"
            If Len(archivo) > 0 Then
                If Dir(archivo) <> "" Then
                    .Attachments.Add archivo
                    Debug.Print "Adjunto: " & archivo
                Else
                    Debug.Print "? No se encontró archivo para adjuntar: " & archivo
                End If
            End If
            .Display
        End With

        Debug.Print "? Correo preparado para " & correoD & " (" & numPoliza & ")"

siguiente:
    Next i

    MsgBox "Correos preparados correctamente.", vbInformation
    Exit Sub

ManejarError:
    Debug.Print "? Error general en índice " & i & " (" & numPoliza & "): " & Err.Description
    MsgBox "Error al enviar correo: " & Err.Description, vbCritical
End Sub

'-------------------------------------------
' Verifica si existe una hoja
'-------------------------------------------
Private Function HojaExiste(nombreHoja As String, wb As Workbook) As Boolean
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = wb.Sheets(nombreHoja)
    HojaExiste = Not sh Is Nothing
    On Error GoTo 0
End Function


