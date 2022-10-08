Attribute VB_Name = "Module1"
Sub SummariseStocks()
    For Each ws In Worksheets

        Dim column As Integer
        Dim lastRow As Long
        Dim open_price As Double
        Dim close_price As Double
        Dim yearly_change As Double
        Dim pct_change As Double
        Dim stock_vol As Double
        Dim stock_name As String
        Dim Summary_table_row As Long
        Dim open_date As String
    
        column = 1
        lastRow = Cells(Rows.Count, 1).End(xlUp).Row
        
        Summary_table_row = 2
        'Create headers for the summary table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        open_date = ws.Cells(2, 2).Value
        stock_vol = 0
        open_price = 0
        
        'Create a summary table
        For i = 2 To lastRow
            If ws.Cells(i + 1, column).Value <> ws.Cells(i, column).Value Then
                'Record stock name in the summary table
                stock_name = ws.Cells(i, 1).Value
                ws.Range("I" & Summary_table_row) = stock_name
                
                'Record close price of that stock
                close_price = ws.Cells(i, 6).Value
                
                'Calculate the yearly change in stock price and record it in the summary table
                yearly_change = close_price - open_price
                ws.Range("J" & Summary_table_row) = yearly_change
                
                'Conditional formatting the yearly change cell
                If yearly_change < 0 Then
                    ws.Range("J" & Summary_table_row).Interior.ColorIndex = 3
                Else
                    ws.Range("J" & Summary_table_row).Interior.ColorIndex = 4
                End If
                
                'Calculate the percentage change and record it in the summary table
                pct_change = yearly_change / open_price
                ws.Range("K" & Summary_table_row) = pct_change
                ws.Range("K" & Summary_table_row).NumberFormat = "0.00%"
                
                'Add the stock volume and record the total stock volume
                stock_vol = stock_vol + ws.Cells(i, 7).Value
                ws.Range("L" & Summary_table_row) = stock_vol
                
                'move to the next row in the summary table and reset the total volume
                Summary_table_row = Summary_table_row + 1
                stock_vol = 0
                open_price = 0
            Else
                'record the opening price at the first openning day of the year
                If ws.Cells(i, 2).Value = open_date Then
                    open_price = ws.Cells(i, 3).Value
                End If
                'Adding up stock volume
                stock_vol = stock_vol + ws.Cells(i, 7).Value
    
            End If
        Next i
        
        'Bonus
        'Create headers for the summary table
        ws.Range("O2").Value = "Greatest % increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest total volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        Dim summary_column As Integer
        Dim summary_table_lastRow As Long
        Dim stock_name_max_pct_change As String
        Dim stock_name_min_pct_change As String
        Dim stock_name_max_vol As String
        summary_column = 9
        summary_table_lastRow = Cells(Rows.Count, 9).End(xlUp).Row
    
        'Find stock with the greatest % increase
        ws.Range("Q2").Value = WorksheetFunction.Max(ws.Range("K2:K" & summary_table_lastRow))
        'Find stock with the greatest % decrease
        ws.Range("Q3").Value = WorksheetFunction.Min(ws.Range("K2:K" & summary_table_lastRow))
        'Pct format
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        'Find stock with the greatest vol
        ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L2:L" & summary_table_lastRow))
        
        'Loop through the summary table to find stock names that has the required values
        For i = 9 To summary_table_lastRow
            If ws.Cells(i, summary_column + 2) = ws.Range("Q2").Value Then
                stock_name_max_pct_change = ws.Cells(i, summary_column).Value
                ws.Range("P2").Value = stock_name_max_pct_change
            ElseIf ws.Cells(i, summary_column + 2) = ws.Range("Q3").Value Then
                stock_name_min_pct_change = ws.Cells(i, summary_column).Value
                ws.Range("P3").Value = stock_name_min_pct_change
            ElseIf ws.Cells(i, summary_column + 3) = ws.Range("Q4").Value Then
                stock_name_max_vol = ws.Cells(i, summary_column).Value
                ws.Range("P4").Value = stock_name_max_vol
            End If
         Next i
        ws.Range("O1").EntireColumn.AutoFit
    Next ws
End Sub


