VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    On Error Resume Next
    Call crearAccesoDirecto
    On Error GoTo Restaurar
    
    'Oculta el archivo de Excel
    Application.Visible = False
    Application.DisplayFormulaBar = False
    Application.DisplayStatusBar = False
    Call ToggleRibbon(False)
    frmCotizad.Show vbModal

Restaurar:
    Call RestaurarInterfaz
End Sub

