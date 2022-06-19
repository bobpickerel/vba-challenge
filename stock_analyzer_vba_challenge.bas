Attribute VB_Name = "Module1"
Sub stock_analyzer():
'loop through data and bring back ticker symbol,
'total stock volume (sum volume),
'Yearly change (Last price - first price) and
'percent change (Last Price - first price)/first price) * 100

    Dim sheet As Worksheet
    Dim maxChange As Double
    Dim maxNegChange As Double
    Dim maxVolume As Double
    Dim maxTicker As String
    Dim maxNegTicker As String
    Dim maxVolTicker As String
    
    stockTick = ""
    
    totalVol = 0
    
    priceOpen = 0
    
    priceClose = 0
    
    summaryTableRow = 2
    
    

    
    For Each sheet In ThisWorkbook.Worksheets
    
        'MsgBox (sheet)
        'Add column headers
        sheet.Cells(summaryTableRow - 1, 9).Value = "Ticker"
        sheet.Cells(summaryTableRow - 1, 10).Value = "Yearly Change"
        sheet.Cells(summaryTableRow - 1, 11).Value = "Percent Change"
        sheet.Cells(summaryTableRow - 1, 12).Value = "Total Stock Volume"
      
        lastRow = sheet.Cells(Rows.Count, 1).End(xlUp).Row
        
        For Row = 2 To lastRow
            'Check to see if the stockTick changes
            If sheet.Cells(Row + 1, 1).Value <> sheet.Cells(Row, 1).Value Then
                'if the stockTick changes...
                'set the stockTick
                
                stockTick = sheet.Cells(Row, 1).Value
                
                priceClose = sheet.Cells(Row, 6).Value
                
                'add the last volume from the row
                totalVol = totalVol + sheet.Cells(Row, 7).Value
                
                'add the stockTick to summary table
                sheet.Cells(summaryTableRow, 9) = stockTick
                
                'add the total volume to the summary table
                sheet.Cells(summaryTableRow, 12) = totalVol
                
                'format volume for readability
                sheet.Cells(summaryTableRow, 12).NumberFormat = "#,###"
                
                'Calculate the yearly change
                sheet.Cells(summaryTableRow, 10) = priceClose - priceOpen
                
                 'Calculate the percent change
                percentChange = ((priceClose - priceOpen) / priceOpen)
                
                'Calculate the persent change
                sheet.Cells(summaryTableRow, 11) = percentChange
                
                'format percent change for readability
                sheet.Cells(summaryTableRow, 11).NumberFormat = "0.00%"
                
                'Add conditional formating to Yearly Change
                If percentChange > 0 Then
                    sheet.Cells(summaryTableRow, 10).Interior.ColorIndex = 4
                ElseIf percentChange = 0 Then
                    sheet.Cells(summaryTableRow, 10).Interior.ColorIndex = 6
                Else
                    sheet.Cells(summaryTableRow, 10).Interior.ColorIndex = 3
                End If
                
                
                'For testing
                'sheet.Cells(summaryTableRow, 13) = priceOpen
                      
                'For testing
                'sheet.Cells(summaryTableRow, 14) = priceClose
                
                'go to the next sumarry table row
                summaryTableRow = summaryTableRow + 1
                
                'reset the total volume for next stock
                totalVol = 0
                
                'reset the priceOpen
                priceOpen = 0
                
                'reset the priceClose
                priceClose = 0
            
            
                    
            Else
                'if the stockTick stays the same
                'add on the total volume from the G column
                totalVol = totalVol + sheet.Cells(Row, 7).Value
                
                    'Get the open price
                    If priceOpen = 0 Then
                        priceOpen = sheet.Cells(Row, 3).Value
                    
                    Else
                        'do nothing
                    End If
                    
            End If
            
        
        
        
        
        
        Next Row
        
        'reset summaryTableRow for next sheet
        summaryTableRow = 2
        
        
        'Column Headers for Max Table
        sheet.Cells(2, 16).Value = "Greatest % Increase"
        sheet.Cells(3, 16).Value = "Greatest % Decrease"
        sheet.Cells(4, 16).Value = "Greatest % Volume"
        sheet.Cells(1, 17).Value = "Ticker"
        sheet.Cells(1, 18).Value = "Value"
        
        'LastRow for Summary Table
        LastRowBonus = sheet.Cells(Rows.Count, 9).End(xlUp).Row
        
        maxChange = 0
        maxNegChange = 0
        maxVolume = 0
        maxTicker = ""
        maxNegTicker = ""
        maxVolTicker = ""
        
        'Find Max Increase, Max Decrease, and Max Volume from Summary Table
        For rowBonus = 2 To LastRowBonus
            If sheet.Cells(rowBonus, 11).Value > maxChange Then
                maxChange = sheet.Cells(rowBonus, 11).Value
                maxTicker = sheet.Cells(rowBonus, 9).Value
            ElseIf sheet.Cells(rowBonus, 11).Value < maxNegChange Then
                maxNegChange = sheet.Cells(rowBonus, 11).Value
                maxNegTicker = sheet.Cells(rowBonus, 9).Value
            Else
                'do nothing
            End If
                
            If sheet.Cells(rowBonus, 12).Value > maxVolume Then
                maxVolume = sheet.Cells(rowBonus, 12).Value
                maxVolTicker = sheet.Cells(rowBonus, 9).Value
                'And maxVolTicker = sheet.Cells(Row, 9)
            Else
            
            End If
            
        Next rowBonus
    
        'Populate Max values into Max Table
        sheet.Cells(2, 17).Value = maxTicker
        sheet.Cells(2, 18).Value = maxChange
        sheet.Cells(2, 18).NumberFormat = "0.00%"
        sheet.Cells(3, 17).Value = maxNegTicker
        sheet.Cells(3, 18).Value = maxNegChange
        sheet.Cells(3, 18).NumberFormat = "0.00%"
        sheet.Cells(4, 17).Value = maxVolTicker
        sheet.Cells(4, 18).Value = maxVolume
        sheet.Cells(4, 18).NumberFormat = "#,###"

    
    Next sheet

    
End Sub
