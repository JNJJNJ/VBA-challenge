Sub Button1_Click()
    For Each ws In Worksheets
        
    'Define variables
    Dim tickerSymbol As String
    Dim lastRow As Long
    Dim outputRow As Integer
    Dim currentTicker As String
    Dim PriceOpen As Double
    Dim PriceClose As Double
    Dim totalStock As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim GreatestTickerIncrease As String
    Dim GreatestValueIncrease As Double
    Dim GreatestTickerDecrease As String
    Dim GreatestValueDecrease As Double
    Dim GreatestTotalTicker As String
    Dim GreatestTotalValue As Double

    'Label Header Columns
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
        
    'Initialize variables
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    outputRow = 2
    PriceOpen = ws.Range("C2").Value
    currentTicker = ws.Range("A2").Value
        
    'Begin loop
    For i = 2 To lastRow
        If currentTicker <> ws.Cells(i, 1).Value Then
            If i > 2 And i < lastRow Then
                    
                PriceClose = ws.Cells(i - 1, 6).Value
                    
                yearlyChange = PriceClose - PriceOpen
                percentChange = yearlyChange / PriceOpen
                    
                    
                'check for greatest increase, decrease, and volume
                If percentChange > GreatestValueIncrease Then
                    GreatestTickerIncrease = currentTicker
                    GreatestValueIncrease = percentChange
                ElseIf percentChange < GreatestValueDecrease Then
                    GreatestTickerDecrease = currentTicker
                    GreatestValueDecrease = percentChange
                End If
                    
                If totalStock > GreatestTotalValue Then
                    GreatestTotalTicker = currentTicker
                    GreatestTotalValue = totalStock
                End If
                    
                ' Print yearlyChange, current Ticker
                ws.Cells(outputRow, 9).Value = currentTicker
                ws.Cells(outputRow, 10).Value = yearlyChange
                    
                ' Check for conditional formatting for positive or negative change
                If yearlyChange > 0 Then
                    ws.Cells(outputRow, 10).Interior.ColorIndex = 4
                Else
                ws.Cells(outputRow, 10).Interior.ColorIndex = 3
                End If
                    
                ws.Cells(outputRow, 11).Value = percentChange
                ws.Cells(outputRow, 11).NumberFormat = "0.00%"
                    
                ws.Cells(outputRow, 12).Value = totalStock
                outputRow = outputRow + 1
                totalStock = 0
                End If
                    
                PriceOpen = ws.Cells(i, 3).Value
                currentTicker = ws.Cells(i, 1).Value
                
                ElseIf i = lastRow Then
                
                totalStock = totalStock + ws.Cells(i, 7)
                
                PriceClose = ws.Cells(i, 6).Value
                
                yearlyChange = PriceClose - PriceOpen
                percentChange = yearlyChange / PriceOpen
                
                'check for greatest increase, decrease, and volume
                If percentChange > GreatestValueIncrease Then
                    GreatestTickerIncrease = currentTicker
                    GreatestValueIncrease = percentChange
                ElseIf percentChange < GreatestValueDecrease Then
                    GreatestTickerDecrease = currentTicker
                    GreatestValueDecrease = percentChange
                End If
                
                If totalStock > GreatestTotalValue Then
                    GreatestTotalTicker = currentTicker
                    GreatestTotalValue = totalStock
                End If
                    
                ' Print yearlyChange, current Ticker
                ws.Cells(outputRow, 9).Value = currentTicker
                ws.Cells(outputRow, 10).Value = yearlyChange
                
                ' Check for conditional formatting for positive or negative change
                If yearlyChange > 0 Then
                    ws.Cells(outputRow, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(outputRow, 10).Interior.ColorIndex = 3
                End If
                
                ws.Cells(outputRow, 11).Value = percentChange
                ws.Cells(outputRow, 11).NumberFormat = "0.00%"
                ws.Cells(outputRow, 12).Value = totalStock
                
            End If
            
            totalStock = totalStock + ws.Cells(i, 7)
        Next i
        
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        ws.Range("P2").Value = GreatestTickerIncrease
        ws.Range("P3").Value = GreatestTickerDecrease
        ws.Range("P4").Value = GreatestTotalTicker
            
        ws.Range("Q2").Value = GreatestValueIncrease
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").Value = GreatestValueDecrease
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("Q4").Value = GreatestTotalValue
        
Next

End Sub