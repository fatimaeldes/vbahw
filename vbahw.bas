Attribute VB_Name = "Module1"
Sub stockTest()

Dim ws As Worksheet

For Each ws In Worksheets

    Dim tickerSym As String
    Dim totalVol As Double

    Dim yearOpen As Double
    Dim yearClose As Double

    Dim summaryRow As Integer
    summaryRow = 2

    ws.Cells(1, 10).Value = "Ticker"
    ws.Cells(1, 11).Value = "Total Volume"
    ws.Cells(1, 12).Value = "Yearly Change"
    ws.Cells(1, 13).Value = "Percent Change"

    Dim lastRow As Double
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        If (ws.Cells(i, 3).Value = 0) Then
            If (ws.Cells(i + 1).Value <> ws.Cells(i, 1).Value) Then
                tickerSym = ws.Cells(i, 1).Value
            End If
        ElseIf (ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value) Then
            totalVol = totalVol + ws.Cells(i, 7).Value
            If (ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value) Then
                yearOpen = ws.Cells(i, 3).Value
            End If
        Else
            tickerSym = ws.Cells(i, 1).Value
            totalVol = totalVol + ws.Cells(i, 7).Value
            yearClose = ws.Cells(i, 6).Value
            ws.Cells(summaryRow, 10).Value = tickerSym
            ws.Cells(summaryRow, 11).Value = totalVol
            If (totalVol > 0) Then
                ws.Cells(summaryRow, 12).Value = yearClose - yearOpen
                    If (ws.Cells(summaryRow, 12).Value > 0) Then
                        ws.Cells(summaryRow, 12).Interior.ColorIndex = 4
                    Else
                        ws.Cells(summaryRow, 12).Interior.ColorIndex = 3
                    End If
                ws.Cells(summaryRow, 13).Value = ws.Cells(summaryRow, 12).Value / yearOpen
            Else
                ws.Cells(summaryRow, 12).Value = 0
                ws.Cells(summaryRow, 13).Value = 0
            End If
            ws.Cells(summaryRow, 13).Style = "percent"
            totalVol = 0
            summaryRow = summaryRow + 1
        End If
        
     Next i
    
    Next ws

End Sub


