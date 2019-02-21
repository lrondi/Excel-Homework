Attribute VB_Name = "StockVolume"
Sub StockVolume()
Dim row As Integer
Dim openValue As Double
Dim closeValue As Double
Dim yearlyChange As Double
Dim percentChange As Double

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

row = 1
totalVolume = 0
openValue = Cells(2, 3).Value
closeValue = 0

For i = 2 To 800000
    If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
        Cells(row + 1, 9).Value = Cells(i, 1).Value
        row = row + 1
        totalVolume = Cells(i, 7).Value
            If i > 2 And openValue <> 0 Then
                closeValue = Cells(i - 1, 6).Value
                yearlyChange = closeValue - openValue
                Cells(row - 1, 10).Value = yearlyChange
                percentChange = yearlyChange / openValue
                Cells(row - 1, 11).Value = percentChange
            End If
        If Cells(i, 3).Value <> 0 Then
            openValue = Cells(i, 3).Value
        Else
            For j = 0 To 1000
                If Cells(i + j, 3).Value <> 0 Then
                    openValue = Cells(i + j, 3).Value
                    Exit For
                    End If
            Next j
        End If
    Else
        totalVolume = totalVolume + Cells(i, 7).Value
        Cells(row, 12).Value = totalVolume
    End If
Next i

'Hard part
Dim Ticker As String

Range("N2").Value = "Greatest % Increase"
Range("N3").Value = "Greatest % Decrease"
Range("N4").Value = "Greatest Total Volume"
Range("O1").Value = "Ticker"
Range("P1").Value = "Value"

maxIncrease = WorksheetFunction.Max(Range("K1:K800000"))
Range("P2").Value = maxIncrease

minIncrease = WorksheetFunction.Min(Range("K1:K800000"))
Range("P3").Value = minIncrease

maxVolume = WorksheetFunction.Max(Range("L1:L800000"))
Range("P4").Value = maxVolume

For i = 1 To 3500
    If Cells(i, 12).Value = maxVolume Then
        Ticker = Cells(i, 9).Value
        Range("O4").Value = Ticker
    End If
    If Cells(i, 11).Value = maxIncrease Then
        Ticker = Cells(i, 9).Value
        Range("O2").Value = Ticker
    End If
    If Cells(i, 11).Value = minIncrease Then
        Ticker = Cells(i, 9).Value
        Range("O3").Value = Ticker
    End If
Next i
End Sub
