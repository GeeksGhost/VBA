Sub Stocks()
    Dim ws As Worksheet
    Dim Ticker As Long
    Dim NewTicker As Long
    Dim TotalStock As Double
    Dim FirstOpenPrice As Double
    Dim LastClosePrice As Double
    Dim YearlyChange As Double
    Dim MaxIncreaseTicker As String
    Dim MaxDecreaseTicker As String
    Dim MaxVolumeTicker As String
    Dim MaxIncrease As Double
    Dim MaxDecrease As Double
    Dim MaxVolume As Double

    For Each ws In ThisWorkbook.Worksheets
        NewTicker = 2
        TotalStock = 0
        FirstOpenPrice = 0
        LastClosePrice = 0
        MaxIncrease = 0
        MaxDecrease = 0
        MaxVolume = 0

        ' Loop through each row in column A
        For Ticker = 2 To ws.Range("A" & ws.Rows.Count).End(xlUp).Row
            
            ' When the next ticker is different, update values one more time and post totals
            If ws.Cells(Ticker + 1, 1).Value <> ws.Cells(Ticker, 1).Value Then
                ' Write Ticker to Column I
                ws.Cells(NewTicker, "I").Value = ws.Cells(Ticker, 1).Value
                ' Write TotalStock to Column L
                TotalStock = TotalStock + ws.Cells(Ticker, 7).Value
                ws.Cells(NewTicker, "L").Value = TotalStock
                
                ' Calculate YearlyChange
                LastClosePrice = ws.Cells(Ticker, 6).Value
                YearlyChange = LastClosePrice - FirstOpenPrice
                ws.Cells(NewTicker, "J").Value = YearlyChange
                
                ' Calculate the percentage and put it in column K
                If FirstOpenPrice <> 0 Then
                    ws.Cells(NewTicker, "K").Value = YearlyChange / FirstOpenPrice
                Else
                    ws.Cells(NewTicker, "K").Value = 0 ' Avoid division by zero error
                End If
                
                ' Set cell color if values are < 0
                If YearlyChange < 0 Then
                    ' Set cell color as red
                    ws.Cells(NewTicker, "J").Interior.ColorIndex = 3
                Else
                    ' Set cell color as green
                    ws.Cells(NewTicker, "J").Interior.ColorIndex = 4
                End If
                
                ' Check for the greatest % increase, % decrease, and total volume
                If ws.Cells(NewTicker, "K").Value > MaxIncrease Then
                    MaxIncrease = ws.Cells(NewTicker, "K").Value
                    MaxIncreaseTicker = ws.Cells(NewTicker, "I").Value
                ElseIf ws.Cells(NewTicker, "K").Value < MaxDecrease Then
                    MaxDecrease = ws.Cells(NewTicker, "K").Value
                    MaxDecreaseTicker = ws.Cells(NewTicker, "I").Value
                End If
                If ws.Cells(NewTicker, "L").Value > MaxVolume Then
                    MaxVolume = ws.Cells(NewTicker, "L").Value
                    MaxVolumeTicker = ws.Cells(NewTicker, "I").Value
                End If
             
                ' Move to the next row for NewTicker
                NewTicker = NewTicker + 1
                ' Reset TotalStock, FirstOpenPrice, and LastClosePrice for the next ticker
                TotalStock = 0
                FirstOpenPrice = 0
                LastClosePrice = 0
            
            ' When the ticker is the same, add values
            Else
                ' Accumulate the total stock value
                TotalStock = TotalStock + ws.Cells(Ticker, 7).Value
                ' Update FirstOpenPrice only if it's not set yet
                If FirstOpenPrice = 0 Then
                    FirstOpenPrice = ws.Cells(Ticker, 3).Value
                End If
                ' Always update LastClosePrice
                LastClosePrice = ws.Cells(Ticker, 6).Value
            End If
        
        Next Ticker
        
        ' Display the results for the greatest % increase, % decrease, and total volume
        ws.Cells(1, "I").Value = "NewTIcker"
        ws.Cells(1, "J").Value = "Yearly Change"
        ws.Cells(1, "K").Value = "Percent Change"
        ws.Cells(1, "L").Value = "Total Volume"
        ws.Cells(1, "O").Value = "Ticker"
        ws.Cells(1, "P").Value = "Value"
        ws.Cells(2, "N").Value = "Greatest % Increase"
        ws.Cells(3, "N").Value = "Greatest % Decrease"
        ws.Cells(4, "N").Value = "Greatest Total Volume"
        ws.Cells(2, "O").Value = MaxIncreaseTicker
        ws.Cells(3, "O").Value = MaxDecreaseTicker
        ws.Cells(4, "O").Value = MaxVolumeTicker
        ws.Cells(2, "P").Value = MaxIncrease
        ws.Cells(3, "P").Value = MaxDecrease
        ws.Cells(4, "P").Value = MaxVolume
        
    Next ws
End Sub


