Attribute VB_Name = "Module1"
Sub stockinfo():

'_______________________________________________________

'Output ticker, open price from beginning year, close price from end of year, total volume
'_______________________________________________________

For Each ws In Worksheets
    
    Dim i As Double
    
    Dim ticker As String
    
    Dim beginYrPrice As Double
    beginYrPrice = 0
    
    Dim endYrPrice As Double
    endYrPrice = 0
    
    Dim yearlyChange As Double
    yearlyChange = 0
    
    Dim percentChange As Double
    percentChange = 0
    
    Dim totalVol As Double
    totalVol = 0
    
    Dim summaryTableRow As Integer
    summaryTableRow = 1
    
    'Write out summary table header starting in column 9
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Volume"
    
    'count how many rows in worksheet
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
    For i = 2 To LastRow
            
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
                'current ticker
                ticker = ws.Cells(i, 1).Value
                
                'total volume
                totalVol = totalVol + ws.Cells(i, 7).Value
                
                'closing price at end of year
                endYrPrice = ws.Cells(i, 6).Value
                
                yearlyChange = endYrPrice - beginYrPrice
                
                If endYrPrice <> 0 Then
                    percentChange = (yearlyChange / endYrPrice) * 100
                Else
                    percentChange = 0
                End If
                
                'fill in values in summary table for this particular stock
                summaryTableRow = summaryTableRow + 1
                ws.Cells(summaryTableRow, 9).Value = ticker
                ws.Cells(summaryTableRow, 10).Value = yearlyChange
                ws.Cells(summaryTableRow, 11).Value = percentChange
                ws.Cells(summaryTableRow, 12).Value = totalVol
                
                'Format cells green if positive change, red if negative change
                If yearlyChange > 0 Then
                    ws.Cells(summaryTableRow, 10).Interior.ColorIndex = 4 'green
                ElseIf yearlyChange < 0 Then
                    ws.Cells(summaryTableRow, 10).Interior.ColorIndex = 3 'red
                End If

                'reset total volume sum and open price
                totalVol = 0
                beginYrPrice = 0
                endYrPrice = 0
                percentChange = 0

            Else
                
                'opening price at beginning of year
                If beginYrPrice = 0 Then
                    beginYrPrice = ws.Cells(i, 3).Value
                End If
                    
                'total volume
                totalVol = totalVol + ws.Cells(i, 7).Value

            End If
            
        Next i
        
        '__________
        'bonus
        '___________
        
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        
        'Greatest % increase
        inc = WorksheetFunction.Max(ws.Range("K:K"))
        incRowNum = Application.WorksheetFunction.Match(inc, ws.Range("K:K"), 0)
        incTicker = ws.Cells(incRowNum, 9)
        ws.Cells(2, 15).Value = incTicker
        ws.Cells(2, 16).Value = inc
        
        'Greatest % decrease
        dec = Application.WorksheetFunction.Min(ws.Range("K:K"))
        decRowNum = Application.WorksheetFunction.Match(dec, ws.Range("K:K"), 0)
        decTicker = ws.Cells(decRowNum, 9)
        ws.Cells(3, 15).Value = decTicker
        ws.Cells(3, 16).Value = dec
        
        'Greatest total volume
        grtsVol = Application.WorksheetFunction.Max(ws.Range("L:L"))
        volRowNum = Application.WorksheetFunction.Match(grtsVol, ws.Range("L:L"), 0)
        volTicker = ws.Cells(volRowNum, 9)
        ws.Cells(4, 15).Value = volTicker
        ws.Cells(4, 16).Value = grtsVol
        
    Next ws

End Sub

