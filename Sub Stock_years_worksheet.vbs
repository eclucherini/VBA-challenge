Sub Stocks_years_worksheets():

'Get variables for every worksheet in the workbook
For Each ws In Worksheets
    
    'lastrow formula to avoid having to count number of rows
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Column titles for outputs
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    'Define variables
    Dim ticker As String
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim stock_volume As LongLong
    Dim year_open As Double
    Dim year_close As Double
    
    'Define loop format
    Dim i As LongLong
    
    'Set calculations to zero and summary_row to 2 to start
    yearly_change = 0
    stock_volume = 0
    summary_row = 2
    
    'Pull each new ticker
    For i = 2 To lastrow
    
       If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ticker = ws.Cells(i, 1).Value
            
            ws.Cells(summary_row, 9).Value = ticker
    
            ticker = ""
            summary_row = summary_row + 1
    
        End If
    
    Next i
    
    'Reset summary_row back to the top at row 2
    summary_row = 2
    
    'Pull each yearly change and percent change
    For i = 2 To lastrow
    
        If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
            year_open = ws.Cells(i, 3).Value
        
        ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            year_close = ws.Cells(i, 6).Value
        
            yearly_change = year_close - year_open
            percent_change = yearly_change / year_open
            
            ws.Cells(summary_row, 10).Value = yearly_change
            ws.Cells(summary_row, 11).Value = percent_change
            
            ws.Cells(summary_row, 10).NumberFormat = "$#,##0.00"
            ws.Cells(summary_row, 11).NumberFormat = "0.00%"
            
            year_open = 0
            year_close = 0
            summary_row = summary_row + 1
        
        End If
    
    Next i
    
    'Reset summary_row back to the top at row 2
    summary_row = 2
    
    'Pull stock volume sum for each ticker
    For i = 2 To lastrow
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            stock_volume = stock_volume + ws.Cells(i, 7).Value
            
            ws.Cells(summary_row, 12).Value = stock_volume
            
            stock_volume = 0
            summary_row = summary_row + 1
        
        Else
            stock_volume = stock_volume + ws.Cells(i, 7).Value
            
        End If
    
    Next i
    
    'Conditional formatting for yearly change
    For i = 2 To summary_row
        
        If ws.Cells(i, 10).Value = "" Then
            Exit For
        ElseIf ws.Cells(i, 10).Value > 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 4
        ElseIf ws.Cells(i, 10).Value < 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 3
        Else
            ws.Cells(i, 10).Interior.ColorIndex = 2
        End If
    
    Next i
    
    'Greatest % increase, % decrease, and total volume
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    'Format output of max and min percent changes to percentages
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    
    'WorksheetFunctions found via Google
    max_percent = WorksheetFunction.Max(ws.Columns("K"))
    min_percent = WorksheetFunction.Min(ws.Columns("K"))
    max_volume = WorksheetFunction.Max(ws.Columns("L"))
    
    'Loop to find max percent change, min percent change, and max stock volume
    For i = 2 To summary_row - 1
        
        If max_percent = ws.Cells(i, 11).Value Then
            ws.Range("Q2").Value = ws.Cells(i, 11).Value
            ws.Range("P2").Value = ws.Cells(i, 9).Value
    
        ElseIf min_percent = ws.Cells(i, 11).Value Then
            ws.Range("Q3").Value = ws.Cells(i, 11).Value
            ws.Range("P3").Value = ws.Cells(i, 9).Value
        
        ElseIf max_volume = ws.Cells(i, 12).Value Then
            ws.Range("Q4").Value = ws.Cells(i, 12).Value
            ws.Range("P4").Value = ws.Cells(i, 9).Value
        
        End If
     
    Next i
    
    'Format columns to autofit width
    ws.Columns("I:Q").AutoFit

Next ws

MsgBox ("finished")

End Sub