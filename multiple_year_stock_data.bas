Attribute VB_Name = "Module1"
Sub StockDataAnalysis()
    Dim WS As Worksheet
    Dim LastRow As Long, i As Long
    Dim OpenPrice As Double, ClosePrice As Double
    Dim VolSum As Double, PercentChange As Double
    Dim Ticker As String
    
    'variables to store the best and worst performance
    Dim MaxIncrease As Double, MaxDecrease As Double, MaxVolumn As Double
    Dim MaxIncreaseTicker As String, MaxDecreaseTicker As String, MaxVolumnTicker As String
    
    'loop through each worksheet named Q1,Q2,Q3,Q4
    For Each WS In ThisWorkbook.Worksheets
        If WS.Name Like "Q?" Then ' check of the workshee's name is Q1,Q2,Q3, or Q4
            LastRow = WS.Cells(WS.Rows.Count, "A").End(xlUp).Row
            If LastRow > 1 Then
                Ticker = WS.Cells(2, 1).Value
                OpenPrice = WS.Cells(2, 3).Value
                VolSum = 0
                
                MaxIncrease = -1E+308
                MaxDecrease = 1E+308
                MaxVolume = 0
                
                'Setup headers for new data
                WS.Cells(1, 9).Value = "Ticker"
                WS.Cells(1, 10).Value = "Quarterly change"
                WS.Cells(1, 11).Value = "Percent change"
                WS.Cells(1, 12).Value = "Total stock volume"
    
                Dim Outputrowindex As Integer
                Outputrowindex = 2
                
                'loop through all rows in the sheet
                For i = 2 To LastRow
                   If WS.Cells(i, 1).Value <> Ticker Or i = LastRow Then
                       If i = LastRow Then
                           VolSum = VolSum + WS.Cells(i, 6).Value ' Ensure last row volume is added
                       End If
                       'Record data for the previous ticker
                       ClosePrice = WS.Cells(i - 1, 5).Value
                       WS.Cells(Outputrowindex, 9).Value = Ticker
                       WS.Cells(Outputrowindex, 10).Value = ClosePrice - OpenPrice
                       WS.Cells(Outputrowindex, 11).Value = Format((ClosePrice - OpenPrice) / OpenPrice, "Percent")
                       WS.Cells(Outputrowindex, 12).Value = VolSum
                       
                       'update maximums and minimums calculations
                       If PercentChange > MaxIncrease Then
                           MaxIncrease = PercentChange
                           MaxIncreaseTicker = Ticker
                       End If
                       If PencentChange < MaxDecrease Then
                           MaxDecrease = PercentChange
                           MaxDecreaseTicker = Ticker
                       End If
                       If VolSum > MaxVolume Then
                           MaxVolume = VolSum
                           MaxVolume = Ticker
                       End If
                       
                       Outputrowindex = Outputrowindex + 1
                       'reset for the next ticker
                       If i < LastRow Then
                           Ticker = WS.Cells(i, 1).Value
                           OpenPrice = WS.Cells(i, 3).Value
                           VolSum = 0
                       End If
                    End If
                    If i < LastRow Then
                    VolSum = VolSum + WS.Cells(i, 6).Value
                    End If
                Next i
                'Output the greatest metrics at specific positions
                WS.Cells(2, 15).Value = "Greatest % increase"
                WS.Cells(2, 15).Value = MaxIncreaseTicker
                WS.Cells(2, 15).Value = Format(MaxIncrease, "Percent")
                WS.Cells(2, 15).Value = "Greatest % decrease"
                WS.Cells(2, 15).Value = MaxIncreaseTicker
                WS.Cells(2, 15).Value = Format(MaxIncrease, "Percent")
                WS.Cells(2, 15).Value = "Greatest total volume"
                WS.Cells(2, 15).Value = MaxVolumnTicker
                WS.Cells(2, 15).Value = MaxVolume
            End If
        End If
    Next WS
End Sub
