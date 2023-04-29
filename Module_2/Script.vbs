Sub StockAnalysis()


' Loop through all worksheets in the workbook
    For Each ws In Worksheets
        
        ' Set initial variables for yearly change, percent change, and total volume
        Dim ticker As String
        Dim openingPrice As Double
        Dim closingPrice As Double
        Dim yearlyChange As Double
        Dim percentChange As Double
        Dim totalVolume As Double
        Dim lastRow As Long
        Dim summaryRow As Long
        
        Dim maxPercentIncrease As Double
        Dim maxPercentDecrease As Double
        Dim maxTotalVolume As Double
        Dim maxPercentIncreaseTicker As String
        Dim maxPercentDecreaseTicker As String
        Dim maxTotalVolumeTicker As String
        
        maxPercentIncrease = 0
        maxPercentDecrease = 0
        maxTotalVolume = 0
        
        ' Set headers for summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Find the last row of data for the worksheet
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Loop through each row of data for the worksheet
        For i = 2 To lastRow
        
            ' Check if we're still on the same ticker symbol, if not, assign the new ticker symbol and opening price
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                openingPrice = ws.Cells(i, 3).Value
            End If
            
            ' Add to the total volume of the current ticker symbol
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            
            ' Check if we're on the last row of data for the current ticker symbol
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                closingPrice = ws.Cells(i, 6).Value
                
                ' Calculate the yearly change and percent change for the current ticker symbol
                yearlyChange = closingPrice - openingPrice
                
                If openingPrice <> 0 Then
                    percentChange = yearlyChange / openingPrice
                Else
                    percentChange = 0
            End If
            
                
            ' Output the ticker symbol, yearly change, percent change, and total volume to the summary table
                summaryRow = ws.Cells(Rows.Count, 9).End(xlUp).Row + 1
                ws.Cells(summaryRow, 9).Value = ticker
                ws.Cells(summaryRow, 10).Value = yearlyChange
                ws.Cells(summaryRow, 11).Value = percentChange
                ws.Cells(summaryRow, 11).NumberFormat = "0.00%"
                ws.Cells(summaryRow, 12).Value = totalVolume
                
            'Apply the format for the yearly change
                If ws.Cells(summaryRow, 10).Value < 0 Then
                    ws.Cells(summaryRow, 10).Interior.ColorIndex = 3
                    Else
                    ws.Cells(summaryRow, 10).Interior.ColorIndex = 4
                End If
                
            'Apply the format for the percentage change
                If ws.Cells(summaryRow, 11).Value < 0 Then
                    ws.Cells(summaryRow, 11).Interior.ColorIndex = 3
                    Else
                    ws.Cells(summaryRow, 11).Interior.ColorIndex = 4
                End If
            
            
                
                
                
                ' Check if this stock has the greatest percent increase, greatest percent decrease, or greatest total volume so far
                If percentChange > maxPercentIncrease Then
                    maxPercentIncrease = percentChange
                    maxPercentIncreaseTicker = ticker
                End If
                
                If percentChange < maxPercentDecrease Then
                    maxPercentDecrease = percentChange
                    maxPercentDecreaseTicker = ticker
                End If
                
                If totalVolume > maxTotalVolume Then
                    maxTotalVolume = totalVolume
                    maxTotalVolumeTicker = ticker
                End If
                ' Reset total volume for the next ticker symbol
                totalVolume = 0
             End If
                
            Next i
        
        
             ' Output the stock with the greatest percent increase, greatest percent decrease, and greatest total volume
            ws.Cells(1, 16).Value = "Ticker"
            ws.Cells(1, 17).Value = "Value"
            ws.Cells(2, 15).Value = "Greatest % Increase"
            ws.Cells(3, 15).Value = "Greatest % Decrease"
            ws.Cells(4, 15).Value = "Greatest Total Volume"
            ws.Cells(2, 16).Value = maxPercentIncreaseTicker
            ws.Cells(3, 16).Value = maxPercentDecreaseTicker
            ws.Cells(4, 16).Value = maxTotalVolumeTicker
            ws.Cells(2, 17).Value = maxPercentIncrease
            ws.Cells(2, 17).NumberFormat = "0.00%"
            ws.Cells(3, 17).Value = maxPercentDecrease
            ws.Cells(3, 17).NumberFormat = "0.00%"
            ws.Cells(4, 17).Value = maxTotalVolume
            ws.Cells(4, 17).EntireColumn.AutoFit
            
            'Entire Column Fit
            ws.Columns("A:Q").EntireColumn.AutoFit
        
        
    Next ws
    
    
    
End Sub
