VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSeleccionArchivosG 
   ClientHeight    =   3040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4170
   OleObjectBlob   =   "frmSeleccionArchivosG.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmSeleccionArchivosG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAceptar_Click()
    If Me.TxtCotizador.Value = "" Or Me.TxtQuinquenios.Value = "" Then
        MsgBox "Selecciona los dos archivos antes de continuar", vbExclamation, "Archivos faltantes"
        Exit Sub
    End If
    ' Oculta el formulario
    Me.Hide
    Debug.Print Me.TxtCotizador.Value
    Call AppMod.BuscarEnArchivo(Me.TxtCotizador.Value, Me.TxtQuinquenios.Value)
    'MsgBox "Rutas seleccionadas:" & vbCrLf & "Cotizador:" & Me.txtCotizador.Value & vbCrLf & "Quinquenios:" & Me.txtQuinquenios.Value, vbInformation, "Verificación"
    'Se enlazara con el módulo principal "AppMódulo"
End Sub
Private Sub btnCancelar_Click()
    Unload Me
End Sub
Private Sub ExaminarC_Click()
    Dim ruta As Variant
    ruta = Application.GetOpenFilename("Archivos Excel (*.xlsm), *.xlsm", , "Selecciona archivo de cotizacion")
    Debug.Print ruta
    If ruta = False Then Exit Sub
    Me.TxtCotizador.Value = ruta
    'Debug.Print Me.txtCotizador.Value
    'Debug.Print Me.txtCotizador.Text
End Sub
Private Sub ExaminarQ_Click()
    Dim ruta As Variant
    ruta = Application.GetOpenFilename("Archivos Excel (*.xlsx), *.xlsx", , "Selecciona archivo de quinquenios")
    Debug.Print ruta
    If ruta = False Then Exit Sub
    Me.TxtQuinquenios.Value = ruta
    'Debug.Print Me.txtQuinquenios.Value
    'Debug.Print Me.txtQuinquenios.Text
End Sub


