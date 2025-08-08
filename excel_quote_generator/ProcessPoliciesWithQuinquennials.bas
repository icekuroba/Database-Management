Option Explicit

Public Sub ProcessPoliciesWithQuinquennials()
    On Error GoTo fail

    ' ===== Performance =====
    Dim prevCalc As XlCalculation
    prevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    ' ===== File selection =====
    MsgBox "Select the quote workbook (.xlsm).", vbInformation
    Dim pathQuote As Variant
    pathQuote = Application.GetOpenFilename("Excel Macro-Enabled (*.xlsm), *.xlsm", , "Select quote file")
    If pathQuote = False Then GoTo cleanup

    MsgBox "Select the quinquennials workbook (.xlsx).", vbInformation
    Dim pathQuinq As Variant
    pathQuinq = Application.GetOpenFilename("Excel Workbook (*.xlsx), *.xlsx", , "Select quinquennials file")
    If pathQuinq = False Then GoTo cleanup

    ' ===== Open workbooks =====
    Dim wbQuote As Workbook, wbQuinq As Workbook
    Set wbQuote = Workbooks.Open(CStr(pathQuote), ReadOnly:=False)
    Set wbQuinq = Workbooks.Open(CStr(pathQuinq), ReadOnly:=True)

    ' ===== Configurable constants (generic names) =====
    Const SH_POLICIES As String = "POLICIES"              ' was: POLIZARIO
    Const COL_POLICY As Long = 2                           ' column B
    Const ROW_START As Long = 9
    Const SH_PROPOSAL As String = "RENEWAL_PROPOSAL"       ' was: PROPUESTA DE RENOVACIÓN
    Const CELL_QUINQ As String = "D15"

    ' Optional dependent procedures (adjust or leave empty)
    Dim dependentProcs() As Variant
    dependentProcs = Array("subgrupos", "Tarifas_enlace", "Tarifa_Modificaciones", "resumen")

    ' ===== Sheet validation =====
    Dim shPol As Worksheet, shProp As Worksheet, shQuinq As Worksheet
    Set shPol = GetSheet(wbQuote, SH_POLICIES)
    Set shProp = GetSheet(wbQuote, SH_PROPOSAL)
    Set shQuinq = wbQuinq.Sheets(1) ' adjust if needed

    If shPol Is Nothing Or shProp Is Nothing Then
        MsgBox "Required sheets not found in the quote workbook.", vbCritical
        GoTo cleanup
    End If

    ' ===== Unprotect (public repo: no real passwords) =====
    Dim passwords() As Variant
    passwords = Array() ' public: keep empty; load locally via InputBox/CONFIG if needed

    Dim ws As Worksheet
    For Each ws In wbQuote.Worksheets
        If ws.ProtectContents Then
            If Not TryUnprotect(ws, passwords) Then
                MsgBox "A protected sheet could not be unprotected: " & ws.Name, vbExclamation
                GoTo cleanup
            End If
        End If
    Next ws

    ' ===== Output folder =====
    Dim outPath As String
    outPath = Environ$("USERPROFILE") & "\Documents\" & Replace(wbQuote.Name, ".xlsm", "") & "_Processed"
    EnsureFolder outPath

    ' ===== Build quinquennial index =====
    Dim lastQ As Long
    lastQ = shQuinq.Cells(shQuinq.Rows.Count, 1).End(xlUp).Row
    Dim rngQuinq As Range
    Set rngQuinq = shQuinq.Range("A1").Resize(lastQ, 2)

    ' ===== Iterate policies =====
    Dim lastRow As Long, r As Long
    lastRow = shPol.Cells(shPol.Rows.Count, COL_POLICY).End(xlUp).Row

    For r = ROW_START To lastRow
        Dim policyName As String
        policyName = Trim(CStr(shPol.Cells(r, COL_POLICY).Value))
        If Len(policyName) = 0 Then GoTo nextRow

        ' Lookup position
        Dim pos As Variant
        pos = Application.Match(policyName, rngQuinq.Columns(1), 0)

        If Not IsError(pos) And Not IsEmpty(pos) Then
            Dim quinquennial As Variant
            quinquennial = rngQuinq.Cells(CLng(pos), 2).Value
            shProp.Range(CELL_QUINQ).Value = quinquennial
        Else
            Debug.Print "Quinquennial not found for policy: " & policyName
        End If

        ' Run dependent procedures if available
        Dim i As Long
        For i = LBound(dependentProcs) To UBound(dependentProcs)
            RunIfExists wbQuote, CStr(dependentProcs(i))
        Next i

        ' Copy available target sheets and save new workbook
        Dim toCopy As Collection: Set toCopy = New Collection
        If Not GetSheet(wbQuote, SH_PROPOSAL) Is Nothing Then toCopy.Add SH_PROPOSAL
        If Not GetSheet(wbQuote, "TEXTS") Is Nothing Then toCopy.Add "TEXTS"
        If Not GetSheet(wbQuote, "ENDORSEMENTS") Is Nothing Then toCopy.Add "ENDORSEMENTS"

        If toCopy.Count > 0 Then
            Dim arr(): ReDim arr(0 To toCopy.Count - 1)
            Dim k As Long
            For k = 1 To toCopy.Count: arr(k - 1) = toCopy(k): Next
            wbQuote.Sheets(arr).Copy

            Dim wbNew As Workbook
            Set wbNew = ActiveWorkbook

            Dim outName As String
            outName = SanitizeFileName(policyName & "_" & Format(Now, "yyyymmdd_HHMMss") & ".xlsm")
            wbNew.SaveAs Filename:=outPath & "\" & outName, FileFormat:=xlOpenXMLWorkbookMacroEnabled
            wbNew.Close SaveChanges:=False
        Else
            Debug.Print "No target sheets available to copy for policy: " & policyName
        End If

nextRow:
    Next r

    MsgBox "Process completed.", vbInformation

cleanup:
    On Error Resume Next
    If Not wbQuote Is Nothing Then wbQuote.Close SaveChanges:=False
    If Not wbQuinq Is Nothing Then wbQuinq.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Application.Calculation = prevCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    Exit Sub

fail:
    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical
    Resume cleanup
End Sub

' === Utilities ===

Private Function GetSheet(wb As Workbook, ByVal name As String) As Worksheet
    On Error Resume Next
    Set GetSheet = wb.Sheets(name)
    On Error GoTo 0
End Function

Private Function TryUnprotect(ws As Worksheet, ByRef passwords As Variant) As Boolean
    Dim i As Long
    On Error Resume Next
    ws.Unprotect Password:=""
    If Err.Number = 0 And Not ws.ProtectContents Then
        TryUnprotect = True
        Exit Function
    End If
    Err.Clear
    For i = LBound(passwords) To UBound(passwords)
        ws.Unprotect Password:=passwords(i)
        If Err.Number = 0 And Not ws.ProtectContents Then
            TryUnprotect = True
            Exit Function
        End If
        Err.Clear
    Next i
    On Error GoTo 0
End Function

Private Sub EnsureFolder(ByVal path As String)
    If Len(Dir$(path, vbDirectory)) = 0 Then MkDir path
End Sub

Private Sub RunIfExists(wb As Workbook, procName As String)
    On Error Resume Next
    Application.Run "'" & wb.Name & "'!" & procName
    Err.Clear
    On Error GoTo 0
End Sub

Private Function SanitizeFileName(ByVal s As String) As String
    ' replace invalid file name characters and trim length
    Dim badChars As Variant: badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    Dim i As Long
    For i = LBound(badChars) To UBound(badChars)
        s = Replace$(s, CStr(badChars(i)), "_")
    Next i
    If Len(s) > 120 Then s = Left$(s, 120)
    SanitizeFileName = s
End Function
