Sub Wolf()

For Each ws In Worksheets

'Set variables
Dim Ticker As String
Dim YearlyChange As Double
Dim PercentChange As Double
Dim SummaryTableRow As Long
SummaryTableRow = 2
Dim YearlyOpen As Double
Dim YearlyClose As Double
Dim PreviousOpen As Long
PreviousOpen = 2

'SummaryTable
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

'Set last row for data
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To LastRow
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        Ticker = ws.Cells(i, 1).Value
        TotalVolume = TotalVolume + ws.Cells(i, 7).Value
        YearlyOpen = ws.Range("C" & PreviousOpen)
        YearlyClose = ws.Range("F" & i)
        YearlyChange = YearlyClose - YearlyOpen

    If YearlyOpen = 0 Then
        PercentChange = 0
    Else
        PercentChange = YearlyChange / YearlyOpen
    End If

    ws.Range("I" & SummaryTableRow).Value = Ticker
    ws.Range("J" & SummaryTableRow).Value = YearlyChange
    ws.Range("K" & SummaryTableRow).Value = PercentChange
    ws.Range("K" & SummaryTableRow).NumberFormat = "0.00%"
    ws.Range("L" & SummaryTableRow).Value = TotalVolume

'Conditional formatting
    If ws.Range("J" & SummaryTableRow).Value >= 0 Then
        ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
    Else
        ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
    End If

'Go next row
SummaryTableRow = SummaryTableRow + 1

PreviousOpen = i + 1

'Reset for next ticker
TotalVolume = 0
Else
TotalVolume = TotalVolume + ws.Cells(i, 7).Value
End If

Next i
Next ws


'Adjust worksheet column width

With Worksheets("2018").Columns("I:L")
 .ColumnWidth = .ColumnWidth * 2
End With

With Worksheets("2019").Columns("I:L")
 .ColumnWidth = .ColumnWidth * 2
End With

With Worksheets("2020").Columns("I:L")
 .ColumnWidth = .ColumnWidth * 2
End With

End Sub