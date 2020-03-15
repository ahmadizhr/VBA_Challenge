Attribute VB_Name = "Module1"


Sub stockmarcket()
For Each ws In Worksheets
   'calculating the last row in each worksheets
   last_row = ws.Cells(Rows.Count, "A").End(xlUp).Row
   'counter c for counting the result's rows
   c = 3
   
   'initiate the header
   ws.Cells(1, 9).Value = "Ticker"
   ws.Cells(1, 10).Value = "Yearly Change"
   ws.Cells(1, 11).Value = "Percent Change"
   ws.Cells(1, 12).Value = "Total Stock Valume"
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   'The Ticker Symbols
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
   'initiate the first item of the result
   ws.Cells(2, 9).Value = ws.Cells(2, 1).Value
   'loop for searching tickers column
   For i = 2 To last_row
       'if the value of the cells changed then write it in the result ticker column
       If ws.Cells((i + 1), 1).Value <> ws.Cells(i, 1).Value Then
          ws.Cells(c, 9).Value = ws.Cells((i + 1), 1).Value
          c = c + 1
       End If
       
   Next i
   
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Yearly change and Percent change from opening price at the beginning of a given year to the closing price at the end of that year
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim yearly_change As Double
    Dim opening_price As Double
    Dim closing_price As Double
    Dim percent_change As Double
    c = 2
    opening_price = ws.Cells(2, 3).Value
    'loop for counting rows in each worksheets
    For i = 2 To last_row
        'condition statement for calculating parameters for all of each stocks seperately
        If ws.Cells(i, 1).Value <> ws.Cells((i + 1), 1).Value Then
            closing_price = ws.Cells(i, 6).Value
            yearly_change = closing_price - opening_price
            ws.Cells(c, 10).Value = yearly_change
            'checking the opening price,divide by 0, its not defined
            If opening_price <> 0 Then
                percent_change = (yearly_change / opening_price)
                ws.Cells(c, 11).Value = percent_change
            Else
                ws.Cells(c, 11).Value = 0
            End If
            
         'counting variable forresult's rows
            c = c + 1
            opening_price = ws.Cells((i + 1), 3).Value
           
        End If
        
    Next i
    
    'select format for percent_change's column
    For i = 2 To last_row
        'ws.Cells(i, 11).Style = "percent"
        ws.Cells(i, 11).NumberFormat = "0.00%"
    Next i
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'The total stock volume of the stock
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Dim Total As Double
    
    c = 2
    Total = ws.Cells(2, 7).Value
    For i = 2 To last_row
        If ws.Cells(i, 1).Value = ws.Cells((i + 1), 1).Value Then
           Total = Total + ws.Cells((i + 1), 7).Value
        Else
           ws.Cells(c, 12).Value = Total
           Total = ws.Cells((i + 1), 7).Value
           c = c + 1
        End If
    Next i
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'color set adjustments for column "yearly change"
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For i = 2 To last_row
        
        If ws.Cells(i, 10).Value < 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 3
        ElseIf ws.Cells(i, 10).Value > 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 4
        End If
        
    Next i
    
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   'find the Greatest % increase and Greatest % decrease of price and Greatest Total Volume
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim Greatest_increase As Variant
    Dim Greatest_decrease As Variant
    Dim Greatest_Total As Double
    Dim row_count As Integer
    'find greatest increase of price and also get it's index to find it's Ticker
    Greatest_increase = 0
    For i = 2 To last_row
        If ws.Cells(i, 11).Value > Greatest_increase Then
                Greatest_increase = ws.Cells(i, 11).Value
                row_count = i
        End If
    Next i
    ws.Cells(2, 16).Value = ws.Cells(row_count, 9).Value
    ws.Cells(2, 17).Value = ws.Cells(row_count, 11).Value
    ws.Cells(2, 17).NumberFormat = "0.00%"
        
    
    'find greatest decrease of price and also get it's index to find it's Ticker
    
    Greatest_decrease = 0
    For i = 2 To last_row
        If ws.Cells(i, 11) < 0 Then
            If ws.Cells(i, 11).Value < Greatest_decrease Then
                Greatest_decrease = ws.Cells(i, 11).Value
                row_count = i
            End If
        End If
    Next i
     
     ws.Cells(3, 16).Value = ws.Cells(row_count, 9).Value
     ws.Cells(3, 17).Value = ws.Cells(row_count, 11).Value
     ws.Cells(3, 17).NumberFormat = "0.00%"
     
        
    'find Greatest Total Volume and also get it's index to find it's Ticker
    
    Greatest_Total = Application.WorksheetFunction.Max(ws.Range("L:L"))
    For i = 2 To last_row
        If ws.Cells(i, 12).Value = Greatest_Total Then
            ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
            ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
        End If
    Next i
    
    

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'design and initiate the header result
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 16).Interior.ColorIndex = 6
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(1, 17).Interior.ColorIndex = 6
    ws.Cells(2, 15).Value = "Greatest%Increase"
    ws.Cells(2, 15).Interior.ColorIndex = 6
    ws.Cells(3, 15).Value = "Greatest%Decrease"
    ws.Cells(3, 15).Interior.ColorIndex = 6
    ws.Cells(4, 15).Value = "Greatest Total Valume"
    ws.Cells(4, 15).Interior.ColorIndex = 6
    
Next ws

End Sub


