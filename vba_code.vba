Sub stock_data()

    'create variables
   Dim i As Long
   Dim ws As Worksheet
   
    'variants for the first table
   Dim ticker As String
   Dim qchange As Double
   Dim percentChange As Double
   Dim stockTotal As Double
   

    'variants for the second table
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim value As Double
    Dim greatestTotal As Double
    
    'summary_table1
    Dim summary_table As Integer

    'loop through all sheets
For Each ws In ThisWorkbook.Worksheets

    'set values
   qchange = 0
   stockTotal = 0
   summary_table = 2

   'lastRow
   Dim lastrow As Long
   lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
   
  'print headers on first table
    ws.Cells(1, 9).value = "Ticker"
    ws.Cells(1, 10).value = "Quarterly Change"
    ws.Cells(1, 11).value = "Percentage Change"
    ws.Cells(1, 12).value = "Total Stock Volume"
    
       
    'print headers second table
    
    ws.Cells(1, 16).value = "Ticker"
    ws.Cells(1, 17).value = "Value"
    ws.Cells(2, 15).value = "Greatest Increase"
    ws.Cells(3, 15).value = "Greatest Decrease"
    ws.Cells(4, 15).value = "Greatest Total"
            
            ' Set to a very low number
           greatestIncrease = -1E+30
          
          ' Set to a very high number
            greatestDecrease = 1E+30
            
            greatestTotal = 0
   

   'loop
   For i = 2 To lastrow
        'conditional
        If i = lastrow Or ws.Cells(i + 1, 1).value <> ws.Cells(i, 1).value Then
            ticker = ws.Cells(i, 1).value
            
            
            qchange = ws.Cells(i, 6).value - ws.Cells(i, 3).value
                
            stockTotal = stockTotal + ws.Cells(i, 7).value

            'conditional2
            If ws.Cells(i, 3).value <> 0 Then
                percentChange = ((ws.Cells(i, 6).value - ws.Cells(i, 3).value) / ws.Cells(i, 3).value)
                                   
            Else
                percentChange = 0
            End If
            
        'fill summary table
            
            ws.Range("I" & summary_table).value = ticker
        
            ws.Range("J" & summary_table).value = qchange
            
            'conditionalFormatingcolors
            If qchange > 0 Then
                ws.Range("J" & summary_table).Interior.ColorIndex = 4
            ElseIf qchange < 0 Then
                ws.Range("J" & summary_table).Interior.ColorIndex = 3
            Else
                qchange = 0
                ws.Range("j" & summary_table).Interior.ColorIndex = 2
            End If
        
            ws.Range("K" & summary_table).value = percentChange
        'Format the PercentChange as percentage
            ws.Range("K" & summary_table).NumberFormat = "0.00%"
        
            ws.Range("L" & summary_table).value = stockTotal
        
            summary_table = summary_table + 1
        
         'condition greatest increase/decrease/total
                Increase = percentChange
                Decrease = percentChange
                Total = stockTotal
                
                If Increase > greatestIncrease Then
                    greatestIncrease = Increase
                    greatestIncreaseTicker = ticker
                End If
                
                If Decrease < greatestDecrease Then
                    greatestDecrease = Decrease
                    greatestDecreaseTicker = ticker
                End If
                
                If Total > greatestTotal Then
                    greatestTotal = Total
                    greatestTotalTicker = ticker
                End If
                
                ' Reset stock total for the next ticker
                stockTotal = 0

        Else
        ' Add to stock total for each ticker
            stockTotal = stockTotal + ws.Cells(i, 7).value
        
       End If
       
             
    Next i
    
    ' Print the greatest values in the summarytable2
        'ticker
        ws.Cells(2, 16).value = greatestIncreaseTicker
        ws.Cells(3, 16).value = greatestDecreaseTicker
        ws.Cells(4, 16).value = greatestTotalTicker
       
        'value
        ws.Cells(2, 17).value = greatestIncrease
        ws.Cells(3, 17).value = greatestDecrease
        ws.Cells(4, 17).value = greatestTotal
         
        'formating percentage
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).NumberFormat = "0.00%"
Next ws
        
End Sub
