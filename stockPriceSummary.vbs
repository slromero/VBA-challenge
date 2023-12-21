Sub stockPriceSummary()
    For Each ws In Worksheets

        Dim openingPrice As Double
        Dim closingPrice As Double
        Dim ticker As String
        openingPrice = ws.Cells(2, 3).Value
        closingPrice = ws.Cells(2, 6).Value
        ticker = ws.Cells(2, 1).Value
            
        Dim openingPriceRow As Long
        Dim closingPriceRow As Long
        Dim tickerRow As Long
        openingPriceRow = 2
        closingPriceRow = 2
        tickerRow = 2
        
        Dim annualPriceChange As Double
        Dim annualPercentChange As Double
        Dim totalStockVolume As LongLong
        Dim currentSummaryRow As Integer
        currentSummaryRow = 2
        
        Dim lastrow As Long
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        Dim MaxIncreaseTicker As String
        Dim MaxDecreaseTicker As String
        Dim MaxTotalVolTicker As String
        Dim MaxIncrease As Double
        Dim MaxDecrease As Double
        Dim MaxTotalVol As Double
        MaxIncrease = 0
        MaxDecrease = 0
        MaxTotalVol = 0
            
            
        For i = 2 To lastrow
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then 'ticker does not match cell below
                
                ticker = ws.Cells(tickerRow, 1) 'capture ticker value
                openingPrice = ws.Cells(openingPriceRow, 3).Value 'capture opening price
                closingPriceRow = i 'capture closingPrice row
                closingPrice = ws.Cells(i, 6).Value 'set closing price
                            
                annualPriceChange = closingPrice - openingPrice
                
                If openingPrice = 0 Then
                    annualPercentChange = 0
                Else
                annualPercentChange = annualPriceChange / openingPrice
                End If
                
                totalStockVolume = totalStockVolume + Cells(i, 7).Value
                
                ws.Cells(currentSummaryRow, 9).Value = ticker
                ws.Cells(currentSummaryRow, 10).Value = annualPriceChange
                ws.Cells(currentSummaryRow, 11).Value = annualPercentChange
                ws.Cells(currentSummaryRow, 12).Value = totalStockVolume
            
                If annualPercentChange > MaxIncrease Then
                    MaxIncrease = annualPercentChange
                    MaxIncreaseTicker = ws.Cells(currentSummaryRow, 9).Value
                End If
                
                If annualPercentChange < MaxDecrease Then
                    MaxDecrease = annualPercentChange
                    MaxDecreaseTicker = ws.Cells(currentSummaryRow, 9).Value
                End If
                
                If totalStockVolume > MaxTotalVol Then
                    MaxTotalVol = totalStockVolume
                    MaxTotalVolTicker = ws.Cells(currentSummaryRow, 9).Value
                End If
                
                ws.Cells(2, 16).Value = MaxIncreaseTicker
                ws.Cells(2, 17).Value = MaxIncrease
                ws.Cells(3, 16).Value = MaxDecreaseTicker
                ws.Cells(3, 17).Value = MaxDecrease
                ws.Cells(4, 16).Value = MaxTotalVolTicker
                ws.Cells(4, 17).Value = MaxTotalVol
            
                openingPriceRow = i + 1
                tickerRow = i + 1
                currentSummaryRow = currentSummaryRow + 1
                totalStockVolume = 0
    
            Else
                ticker = ws.Cells(tickerRow, 1) 'capture ticker value
                openingPrice = ws.Cells(openingPriceRow, 3).Value
                closingPriceRow = i
                closingPrice = ws.Cells(i, 6).Value
                totalStockVolume = totalStockVolume + ws.Cells(i, 7).Value
            End If
        Next i
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    Next ws
End Sub
