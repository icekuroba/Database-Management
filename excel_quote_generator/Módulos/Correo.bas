'================================================================================================================================
'                                   Función para envio por correo
'================================================================================================================================
Public Const correos As String = "TablaCorreos"
Public polizasProcesadas() As String
Public rutasProcesadas() As String
Public contadorPolizas As Long
Public Sub IniciarRegistros(Optional dummy As Boolean)
    ReDim polizasProcesadas(1 To 1)
    ReDim rutasProcesadas(1 To 1)
    contadorPolizas = 0
    Debug.Print "Registro de polizas iniciando"
End Sub
Public Sub RegistrarPolizas(ByVal numPoliza As String, ByVal rutaArchivo As String)
    If Len(numPoliza) = 0 Or Len(rutaArchivo) = 0 Then Exit Sub
    contadorPolizas = contadorPolizas + 1
    ReDim Preserve polizasProcesadas(1 To contadorPolizas)
    ReDim Preserve rutasProcesadas(1 To contadorPolizas)
    polizasProcesadas(contadorPolizas) = Trim(numPoliza)
    rutasProcesadas(contadorPolizas) = rutaArchivo
    Debug.Print "Registrada #" & contadorPolizas & "->" & numPoliza & "| Ruta: " & rutaArchivo
End Sub
Public Sub EnviarCorreo(tipoCotizador As String)
    On Error GoTo ManejarError
    Dim base As Worksheet, i As Long, ultimaFila As Long, numPoliza As String
    Dim archivo As String, idxPoliza As Object, encontrada As Boolean
    Dim olApp As Object, olMail As Object, cuenta As Object
    Dim ej As String, gerente As String, correoE As String, correoG As String
    Dim datos As Variant, paraTo As String, paraCC As String
    Dim nombreCuenta As String

    '-- 1) Validar si hay polizas
    If contadorPolizas = 0 Then
        MsgBox "No hay polizas registradas para enviar", vbExclamation
        Exit Sub
    End If

    Set base = ThisWorkbook.Sheets(correos) ' "TablaCorreos"
    ultimaFila = base.Cells(base.Rows.Count, "A").End(xlUp).Row

    Set idxPoliza = CreateObject("Scripting.Dictionary")
    idxPoliza.CompareMode = 1

    '-- 2) Definir la hoja base y cargar tabla en array
    Dim pol As String, corE As String, corG_Idx As String
    For i = 2 To ultimaFila
        pol = Trim$(CStr(base.Cells(i, "A").Value)) ' Póliza
        ej = Trim$(CStr(base.Cells(i, "B").Value))  ' Ejecutivo
        corE = Trim$(CStr(base.Cells(i, "C").Value)) ' Correo Ejecutivo
        gerente = Trim$(CStr(base.Cells(i, "D").Value)) ' Gerencia
        corG_Idx = Trim$(CStr(base.Cells(i, "E").Value)) ' Correo Gerencia
        'siempre indexa la póliza, aunque el correo de ejecutivo esté vacío
        If Len(pol) > 0 Then '  "And Len(cor) > 0 Then"  -> esto era para que saltara la poliza
            idxPoliza(pol) = ej & "|" & corE & "|" & gerente & "|" & corG_Idx
        End If
    Next i

    '-- 3) Crear instancia de outlook
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then Set olApp = CreateObject("Outlook.Application")
    On Error GoTo ManejarError

    '-- 4) Busca la poliza en el array para el correo destino y con copia
    For i = 1 To contadorPolizas
        numPoliza = Trim$(CStr(polizasProcesadas(i)))
        archivo = Trim$(Replace(rutasProcesadas(i), """", ""))

        If Len(numPoliza) = 0 Then GoTo siguiente
        If Not idxPoliza.Exists(numPoliza) Then
            Debug.Print "No se encontró la póliza [" & numPoliza & "] en " & correos
            GoTo siguiente
        End If

        datos = Split(CStr(idxPoliza(numPoliza)), "|")
        ej = Trim$(CStr(datos(0)))
        correoE = Trim$(CStr(datos(1)))
        gerente = Trim$(CStr(datos(2)))
        correoG = ""
        If UBound(datos) >= 3 Then correoG = Trim$(CStr(datos(3)))

        ' === A) Selección de destinatarios (fallback a Gerencia) ==
        paraTo = ""
        paraCC = ""
        If Len(correoE) > 0 Then
            paraTo = correoE
            If Len(correoG) > 0 Then paraCC = paraCC & "; " & correoG
        Else
            If Len(correoG) > 0 Then
                paraTo = correoG
            Else
                Debug.Print "Sin correo de ejecutivo NI gerencia para póliza " & numPoliza & ". Se omite."
                GoTo siguiente
            End If
        End If

    '-- 5) Crear correos
        On Error Resume Next
        If olApp Is Nothing Then
            Set olApp = GetObject(, "Outlook.Application")
            If olApp Is Nothing Then Set olApp = CreateObject("Outlook.Application")
        End If
        On Error GoTo ManejarError

        Set olMail = olApp.CreateItem(0)

    '-- 6) Busca la cuenta de remitente desde Outlook
        nombreCuenta = ""
        Dim acc As Object
        For Each acc In olApp.Session.Accounts
            If LCase$(acc.SmtpAddress) = LCase$(nombreCuenta) Then
                Set olMail.SendUsingAccount = acc
                Exit For
            End If
        Next acc
        If olMail.SendUsingAccount Is Nothing Then
            On Error Resume Next
            Set olMail.SendUsingAccount = olApp.Session.Accounts.item(1)
            On Error GoTo ManejarError
        End If

        ' Cuerpo
        Dim htmlBody As String
        htmlBody = "<html>" & "<body style='margin:0; padding:0; background-color:#ffffff;'>" & _
        "<table width='100%' cellpadding='0' cellspacing='0' style='font-family:Segoe UI, Arial, sans-serif;'><tr><td align='center'>" & _
        "<table width='720' cellpadding='24' cellspacing='0' style='background-color:#ffffff;'>" & _
        "<!-- Línea superior -->" & "<tr><td>" & _
        "  <table width='100%' cellpadding='0' cellspacing='0' role='presentation'>" & _
        "    <tr>" & _
        "      <td width='33.33%' style='height:2px; line-height:2px; font-size:0; background-color:#0090DA;'>&nbsp;</td>" & _
        "      <td width='33.33%' style='height:2px; line-height:2px; font-size:0; background-color:#4B0082;'>&nbsp;</td>" & _
        "      <td width='33.33%' style='height:2px; line-height:2px; font-size:0; background-color:#5A5AE6;'>&nbsp;</td>" & _
        "    </tr>" & "  </table>" & "</td></tr>" & _
        "<tr><td style='font-size:18px; line-height:1.8; color:#333333;'>" & _
        "<p style='margin:0 0 16px 0;'>Buen día,</p>" & _
        "<p style='margin:0 0 18px 0;'>Adjunto la propuesta bajo mismos términos y condiciones para la póliza indicada en el nombre correspondiente.</p>" & _
        "<p style='margin:0;'>Saludos.</p>" & _
        "</td></tr>" & _
        "<tr><td style='padding-top:24px; text-align:center;'>" & _
        "<img src='https://raw.githubusercontent.com/icekuroba/Database-Management/refs/heads/main/excel_quote_generator/image/5.jpg' alt='Logo' width='140' style='display:inline-block;html>"
        '-------------------------------------------------------------------------------------------------------------------------------------------
        With olMail
            .To = paraTo
            .CC = paraCC
            .Subject = "Propuesta de Renovación " & tipoCotizador & " - Póliza " & numPoliza
            .htmlBody = htmlBody
            On Error Resume Next
            If Len(archivo) > 0 Then
                If Dir(archivo) <> "" Then
                .Attachments.Add archivo
                Debug.Print "Archivo adjunto: " & archivo
            Else
                Debug.Print "No se encontro archivo: " & archivo
                End If
            End If
            On Error GoTo ManejarError
            .Display
            '************.Send
        End With
        Debug.Print "Correo preparado -> TO: " & paraTo & " | CC: " & paraCC & " | Póliza: " & numPoliza
 
siguiente:
    Next i
    Exit Sub

ManejarError:
    Debug.Print "Error al enviar el correo: " & Err.Description
    MsgBox "Error al enviar el correo: " & Err.Description, vbCritical
End Sub


