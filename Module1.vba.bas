Attribute VB_Name = "Module1"
Sub MultipleQuarterStockData()
    Dim ws As Worksheet
    Dim WorksheetName As String
    Dim i As Long
    Dim j As Long
    Dim TickCount As Long
    Dim LastRowA As Long
    Dim LastRowI As Long
    Dim PerChange As Double
    Dim GreatIncr As Double
    Dim GreatDecr As Double
    Dim GreatVol As Double
    Dim GreatIncrTicker As String
    Dim GreatDecrTicker As String
    Dim GreatVolTicker As String

    For Each ws In Worksheets
        WorksheetName = ws.Name

        ' Creating column headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"

        TickCount = 2
        j = 2
        LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row

        For i = 2 To LastRowA
            ' Checking if the quarter changes or the ticker changes
            If DatePart("q", ws.Cells(i + 1, 2).Value) <> DatePart("q", ws.Cells(i, 2).Value) Or ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ws.Cells(TickCount, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(TickCount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value

                ' Conditional formatting for Quarterly Change
                If ws.Cells(TickCount, 10).Value < 0 Then
                    ws.Cells(TickCount, 10).Interior.ColorIndex = 3 ' Red
                ElseIf ws.Cells(TickCount, 10).Value > 0 Then
                    ws.Cells(TickCount, 10).Interior.ColorIndex = 4 ' Green
                Else
                    ws.Cells(TickCount, 10).Interior.ColorIndex = xlNone ' No color for 0
                End If

                ' Calculating and writing percent change in column K (#11)
                If ws.Cells(j, 3).Value <> 0 Then
                    PerChange = (ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value
                Else
                    PerChange = 0
                End If
                ws.Cells(TickCount, 11).Value = Format(PerChange, "0.00%")

                ' Conditional formatting for Percent Change
                If PerChange < 0 Then
                    ws.Cells(TickCount, 11).Interior.ColorIndex = 3 ' Red
                ElseIf PerChange > 0 Then
                    ws.Cells(TickCount, 11).Interior.ColorIndex = 4 ' Green
                Else
                    ws.Cells(TickCount, 11).Interior.ColorIndex = xlNone ' No color for 0
                End If

                ' Calculate and write total volume in column L (#12)
                ws.Cells(TickCount, 12).Value = WorksheetFunction.Sum(ws.Range(ws.Cells(j, 7), ws.Cells(i, 7)))

                TickCount = TickCount + 1
                j = i + 1
            End If
        Next i

        LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row

        GreatVol = ws.Cells(2, 12).Value
        GreatIncr = ws.Cells(2, 11).Value
        GreatDecr = ws.Cells(2, 11).Value
        GreatIncrTicker = ws.Cells(2, 9).Value
        GreatDecrTicker = ws.Cells(2, 9).Value
        GreatVolTicker = ws.Cells(2, 9).Value

        For i = 2 To LastRowI
            If ws.Cells(i, 12).Value > GreatVol Then
                GreatVol = ws.Cells(i, 12).Value
                GreatVolTicker = ws.Cells(i, 9).Value
            End If

            If ws.Cells(i, 11).Value > GreatIncr Then
                GreatIncr = ws.Cells(i, 11).Value
                GreatIncrTicker = ws.Cells(i, 9).Value
            End If

            If ws.Cells(i, 11).Value < GreatDecr Then
                GreatDecr = ws.Cells(i, 11).Value
                GreatDecrTicker = ws.Cells(i, 9).Value
            End If
        Next i

        ws.Cells(2, 16).Value = GreatIncrTicker
        ws.Cells(3, 16).Value = GreatDecrTicker
        ws.Cells(4, 16).Value = GreatVolTicker

        ws.Cells(2, 17).Value = Format(GreatIncr, "0.00%")
        ws.Cells(3, 17).Value = Format(GreatDecr, "0.00%")
        ws.Cells(4, 17).Value = Format(GreatVol, "0.00E+00")

        ws.Columns("A:Z").AutoFit
    Next ws
End Sub

