Option Explicit:      DefLng A-Z

Public Sub Barcodes()
    On Error GoTo ErrorExit
    Dim lLastRow As Long
    Dim Cell     As Range
    Dim sCol     As String
    Dim xlSheet1 As Worksheet
    Dim xlSheet2 As Worksheet

    If ActiveWorkbook.ActiveSheet.Range("A1").Value2 = vbNullString Then
        MsgBox "Please open a .csv file first!", vbCritical + vbOKOnly
    End If

    LudicrousMode True

    Set xlSheet1 = ActiveWorkbook.ActiveSheet
    Set xlSheet2 = ActiveWorkbook.Worksheets.Add

    xlSheet1.Rows(2).EntireRow.Delete
    lLastRow = xlSheet1.UsedRange.Rows.Count

    For Each Cell In xlSheet1.Range("A1:CA1")
        If Cell.Value2 = "From Location" Then
            sCol = Split(Cells(1, Cell.Column).Address, "$")(1)
            xlSheet2.Range("A1:A" & lLastRow).Value2 = xlSheet1.Range(sCol & "1:" & sCol & lLastRow).Value2
        End If

        If Cell.Value2 = "Move Time" Then
            sCol = Split(Cells(1, Cell.Column).Address, "$")(1)
            xlSheet2.Range("B1:B" & lLastRow).Value2 = xlSheet1.Range(sCol & "1:" & sCol & lLastRow).Value2
        End If

        If Cell.Value2 = "Pallet" Then
            sCol = Split(Cells(1, Cell.Column).Address, "$")(1)
            xlSheet2.Range("C1:C" & lLastRow).Value2 = xlSheet1.Range(sCol & "1:" & sCol & lLastRow).Value2
        End If

        If Cell.Value2 = "Consignment" Then
            sCol = Split(Cells(1, Cell.Column).Address, "$")(1)
            xlSheet2.Range("D1:D" & lLastRow).Value2 = xlSheet1.Range(sCol & "1:" & sCol & lLastRow).Value2
        End If

        If Cell.Value2 = "Stage Route ID" Then
            sCol = Split(Cells(1, Cell.Column).Address, "$")(1)
            xlSheet2.Range("F1:F" & lLastRow).Value2 = xlSheet1.Range(sCol & "1:" & sCol & lLastRow).Value2
        End If
    Next Cell
    xlSheet2.Range("E1").Value2 = "Barcode"
    For Each Cell In xlSheet2.Range("C2:C" & lLastRow)
        Cell.Offset(0, 2).Value2 = "*" & Cell.Value2 & "*"
    Next Cell

    xlSheet2.Range("B2:B" & lLastRow).NumberFormat = "hh:mm"
    xlSheet2.Range("A1:F1").Font.Bold = True
    xlSheet2.Range("A1:F1").AutoFilter
    xlSheet2.Range("A:F").VerticalAlignment = xlCenter
    xlSheet2.Range("A:F").HorizontalAlignment = xlCenter
    xlSheet2.Range("E2:E" & lLastRow).Font.Name = "Free 3 of 9"
    xlSheet2.Range("E2:E" & lLastRow).Font.Size = 18
    xlSheet2.Range("A1:F" & lLastRow).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    xlSheet2.Columns("A:F").AutoFit

    Excel.Application.EnableEvents = False
    xlSheet1.Delete
ErrorExit:

    If Err.Number <> 0 Then MsgBox Err.Number & vbTab$ & Err.Description, vbCritical + vbOKOnly
    LudicrousMode False
    On Error GoTo 0
End Sub

Private Sub LudicrousMode(ByVal Toggle As Boolean)
    Application.ScreenUpdating = Not Toggle
    Application.EnableEvents = Not Toggle
    Application.DisplayAlerts = Not Toggle
    Application.EnableAnimations = Not Toggle
    Application.DisplayStatusBar = Not Toggle
    Application.PrintCommunication = Not Toggle
    Application.Calculation = IIf(Toggle, xlCalculationManual, xlCalculationAutomatic)
End Sub
