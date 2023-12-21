Attribute VB_Name = "Module1"

Sub ProcessAllSheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "Summary" Then
            ProcessSheet ws
        End If
    Next ws
End Sub

Sub ProcessSheet(ws As Worksheet)
    Dim lastRow As Long
    Dim i As Long
    Dim uniqueValues As Collection
    Dim currentValue As Variant
    Dim firstC As Double
    Dim lastF As Double
    Dim percentChange As Double
    Dim sumColumnG As Double
    Dim firstRowDict As Object
    Dim lastRowDict As Object
    Dim maxColumnK As Double
    Dim minColumnK As Double
    Dim maxColumnL As Double
    Dim maxColumnKRow As Long
    Dim minColumnKRow As Long
    Dim maxColumnLRow As Long
    Dim overallMaxColumnK As Double
    Dim overallMinColumnK As Double
    Dim overallMaxColumnL As Double
    Dim overallMaxColumnKRow As Long
    Dim overallMinColumnKRow As Long
    Dim overallMaxColumnLRow As Long
    Dim maxColumnKSheet As String
    Dim minColumnKSheet As String
    Dim maxColumnLSheet As String

    overallMaxColumnK = -1E+30
    overallMinColumnK = 1E+30
    overallMaxColumnL = -1E+30

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Set uniqueValues = New Collection

    On Error Resume Next
    For i = 2 To lastRow
        currentValue = ws.Cells(i, 1).Value
        If Not IsError(currentValue) And Len(currentValue) > 0 Then
            uniqueValues.Add currentValue, CStr(currentValue)
        End If
    Next i
    On Error GoTo 0

    Set firstRowDict = CreateObject("Scripting.Dictionary")
    Set lastRowDict = CreateObject("Scripting.Dictionary")

    For i = 2 To lastRow
        currentValue = ws.Cells(i, 1).Value
        If Not firstRowDict.Exists(currentValue) Then
            firstRowDict.Add currentValue, i
        End If
        lastRowDict(currentValue) = i
    Next i

    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

    For i = 1 To uniqueValues.Count
        ws.Cells(i + 1, 9).Value = uniqueValues(i)
        firstC = ws.Cells(firstRowDict(uniqueValues(i)), "C").Value
        lastF = ws.Cells(lastRowDict(uniqueValues(i)), "F").Value
        ws.Cells(i + 1, 10).Value = lastF - firstC
        If firstC <> 0 Then
            percentChange = ((lastF - firstC) / Abs(firstC))
            ws.Cells(i + 1, 11).Value = percentChange
        Else
            ws.Cells(i + 1, 11).Value = 0
        End If
        ws.Cells(i + 1, 11).NumberFormat = "0.00%"
        sumColumnG = Application.WorksheetFunction.SumIf(ws.Range("A:A"), uniqueValues(i), ws.Range("G:G"))
        ws.Cells(i + 1, 12).Value = sumColumnG
    Next i

    maxColumnK = Application.WorksheetFunction.Max(ws.Range(ws.Cells(2, 11), ws.Cells(lastRow, 11)))
    minColumnK = Application.WorksheetFunction.Min(ws.Range(ws.Cells(2, 11), ws.Cells(lastRow, 11)))
    maxColumnL = Application.WorksheetFunction.Max(ws.Range(ws.Cells(2, 12), ws.Cells(lastRow, 12)))

    maxColumnKRow = WorksheetFunction.Match(maxColumnK, ws.Range(ws.Cells(2, 11), ws.Cells(lastRow, 11)), 0) + 1
    minColumnKRow = WorksheetFunction.Match(minColumnK, ws.Range(ws.Cells(2, 11), ws.Cells(lastRow, 11)), 0) + 1
    maxColumnLRow = WorksheetFunction.Match(maxColumnL, ws.Range(ws.Cells(2, 12), ws.Cells(lastRow, 12)), 0) + 1

    overallMaxColumnK = maxColumnK
    overallMinColumnK = minColumnK
    overallMaxColumnL = maxColumnL
    overallMaxColumnKRow = maxColumnKRow
    overallMinColumnKRow = minColumnKRow
    overallMaxColumnLRow = maxColumnLRow
    maxColumnKSheet = ws.Name
    minColumnKSheet = ws.Name
    maxColumnLSheet = ws.Name

    ws.Cells(2, "Q").Value = overallMaxColumnK
    ws.Cells(2, "Q").NumberFormat = "0.00%"
    ws.Cells(3, "Q").Value = overallMinColumnK
    ws.Cells(3, "Q").NumberFormat = "0.00%"
    ws.Cells(4, "Q").Value = overallMaxColumnL
    ws.Cells(4, "Q").NumberFormat = "General"

    ws.Cells(2, "P").Value = ws.Cells(overallMaxColumnKRow, "I").Value
    ws.Cells(3, "P").Value = ws.Cells(overallMinColumnKRow, "I").Value
    ws.Cells(4, "P").Value = ws.Cells(overallMaxColumnLRow, "I").Value

    ApplyConditionalFormatting ws
End Sub

Sub ApplyConditionalFormatting(ws As Worksheet)
    Dim rng As Range

    Set rng = ws.Range(ws.Cells(2, 10), ws.Cells(ws.Cells(ws.Rows.Count, 10).End(xlUp).Row, 10))

    rng.FormatConditions.Delete
    rng.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
    rng.FormatConditions(1).Interior.Color = RGB(255, 0, 0)
    rng.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
    rng.FormatConditions(2).Interior.Color = RGB(0, 255, 0)

    ws.Range(ws.Cells(2, 12), ws.Cells(ws.Cells(ws.Rows.Count, 12).End(xlUp).Row, 12)).FormatConditions.Delete
End Sub

