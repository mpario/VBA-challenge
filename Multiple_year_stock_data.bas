Attribute VB_Name = "Module1"
'Written by: Mehul Parekh

Sub stockmarketchange()

For Each ws In Worksheets

Dim ticker_last_row As String
Dim summary_table_row As Long
Dim start_year_open As Double
Dim end_year_close As Double
Dim year_change As Double
Dim percent_change As Double
Dim start_year_volume As Long
Dim end_year_volume As Long
Dim total_stock_volume As Double

Dim j As Long

'Header columns
    ws.Cells(1, 9) = "Ticker"
    ws.Cells(1, 10) = "Yearly Change"
    ws.Cells(1, 11) = "Percent Change"
    ws.Cells(1, 12) = "Total Stock Volume"
    
'Last row finder
    ticker_last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
'Initialize counters
    summary_table_row = 2
    start_year_open = ws.Cells(2, 3).Value
    total_stock_volume = ws.Cells(2, 7).Value
    
'Loop through ticker values
    For j = 2 To ticker_last_row
        current_row_ticker = ws.Cells(j, 1).Value
        next_row_ticker = ws.Cells(j + 1, 1).Value
            
        'Runs on the last day of the year for the current stock
        If current_row_ticker <> next_row_ticker Then
            
            'Input
            end_year_close = ws.Cells(j, 6).Value
            total_stock_volume = 0
            
            'Calculation
            year_change = Round(end_year_close - start_year_open, 2)
            
            'Output
            ws.Cells(summary_table_row, 9).Value = current_row_ticker
            ws.Cells(summary_table_row, 10).Value = year_change
            'ws.Cells(summary_table_row, 14).Value = start_year_open
            'ws.Cells(summary_table_row, 15).Value = end_year_close
    
            'Percent Increase
            If start_year_open = 0 Then
                percent_change = 0
                ws.Cells(summary_table_row, 11).Value = FormatPercent(percent_change, 2)
            ElseIf year_change > 0 Then
                percent_change = year_change / start_year_open
                ws.Cells(summary_table_row, 11).Value = FormatPercent(percent_change, 2)
            'Percent Decrease
            ElseIf year_change < 0 Then
                percent_change = (start_year_open - end_year_close) / start_year_open
                ws.Cells(summary_table_row, 11).Value = FormatPercent(percent_change, 2)
            'Percent No Change
            Else
                percent_change = 0
                ws.Cells(summary_table_row, 11).Value = FormatPercent(percent_change, 2)
            End If
            
            'Setup
            start_year_open = ws.Cells(j + 1, 3).Value
            
            'Coloring Column J
            If year_change < 0 Then
                ws.Cells(summary_table_row, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(summary_table_row, 10).Interior.ColorIndex = 4
            End If
            
        summary_table_row = summary_table_row + 1
    
    Else
        'Total Stock Volume Counter
        total_stock_volume = total_stock_volume + ws.Cells(j + 1, 7).Value
        ws.Cells(summary_table_row, 12).Value = total_stock_volume
        
    End If
    
    Next j

Next ws

End Sub
