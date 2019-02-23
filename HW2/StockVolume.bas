Attribute VB_Name = "Module1"
Sub StockVolume()
For Each ws In Worksheets
    
    Dim row As Integer
    Dim openValue As Double
    Dim closeValue As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim Ticker As String
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).row

    row = 1
    totalVolume = 0
    openValue = ws.Cells(2, 3).Value
    closeValue = 0
        
        
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            ws.Cells(row + 1, 9).Value = ws.Cells(i, 1).Value
            row = row + 1
                If i > 2 And openValue <> 0 Then
                    closeValue = ws.Cells(i - 1, 6).Value
                    yearlyChange = closeValue - openValue
                    ws.Cells(row - 1, 10).Value = yearlyChange
                    percentChange = yearlyChange / openValue
                    ws.Cells(row - 1, 11).Value = percentChange
                    ws.Cells(row - 1, 12).Value = totalVolume
                End If
            If ws.Cells(i, 3).Value <> 0 Then
                openValue = ws.Cells(i, 3).Value
            Else
                For j = 0 To 500 'arbitrary upper bond, iteration through same Ticker until first non-zero value is find
                    If ws.Cells(i + j, 3).Value <> 0 Then
                        openValue = ws.Cells(i + j, 3).Value
                        Exit For
                    End If
                Next j
            End If
            totalVolume = ws.Cells(i, 7).Value
        Else
            totalVolume = totalVolume + ws.Cells(i, 7).Value
        End If
    Next i


    'Hard part

    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
    
    maxIncrease = WorksheetFunction.Max(ws.Range("K1:K8000"))
    ws.Range("P2").Value = maxIncrease
    ws.Range("P2").NumberFormat = "0.00%"

    minIncrease = WorksheetFunction.Min(ws.Range("K1:K8000"))
    ws.Range("P3").Value = minIncrease
    ws.Range("P3").NumberFormat = "0.00%"
    
    maxVolume = WorksheetFunction.Max(ws.Range("L1:L8000"))
    ws.Range("P4").Value = maxVolume

    lastTickerRow = ws.Cells(Rows.Count, 9).End(xlUp).row
    
    For i = 2 To lastTickerRow
        If ws.Cells(i, 12).Value = maxVolume Then
            Ticker = ws.Cells(i, 9).Value
            ws.Range("O4").Value = Ticker
        End If
        If ws.Cells(i, 11).Value = maxIncrease Then
            Ticker = ws.Cells(i, 9).Value
            ws.Range("O2").Value = Ticker
        End If
        If ws.Cells(i, 11).Value = minIncrease Then
            Ticker = ws.Cells(i, 9).Value
            ws.Range("O3").Value = Ticker
        End If
        If ws.Cells(i, 10).Value >= 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 4
        Else
            ws.Cells(i, 10).Interior.ColorIndex = 3
        End If
        ws.Cells(i, 11).NumberFormat = "0.00%"
    Next i
    
Next ws

End Sub

