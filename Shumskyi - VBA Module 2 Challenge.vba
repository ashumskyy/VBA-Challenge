Sub multiple_year_stock_data()
    
    
    'Declaring a variable to count existing worksheets
    Dim current_WS As Integer
    current_WS = Application.Worksheets.Count
    
    'FORlooping through existing worksheets
    For I = 1 To current_WS
    'Activating the current worksheet
    Worksheets(I).Activate
    
    
    'Giving the names for needed cells
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    
    
    'Declaring a variable and assigning the counter for rows in our Summary table
    Dim summary_table_row As Long
    summary_table_row = 2
    
    'Declaring needed variables. Using Double data type for large numbers.
    Dim ticker_name As String
    Dim yearly_changes As Double
    Dim percent_changes As Double
    Dim open_year As Double
    Dim close_year As Double
    'Total volume starts with 0
    Dim total_volume As Double
    total_volume = 0
    
    'Declaring a variable and assigning it to the last row finder
    Dim last_row As Long
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Declaring variables for the Greatest values table and assigning the counters
    Dim gr_perc_incr As Double
    gr_perc_incr = 0
    Dim gr_perc_incr_ticker As String
    
    Dim gr_perc_decr As Double
    gr_perc_decr = 0
    Dim gr_perc_decr_ticker As String
    
    Dim gr_ttl_vol As Double
    gr_ttl_vol = 0
    Dim gr_ttl_vol_ticker As String
 
    'LOGIC
    
    'FORlooping through every row on the current worksheet
    For current_row = 2 To last_row
        
        'By using VBA "not equal <>" functionality
        'Finding the opening price at the beginning of the year and assigning it
        If Cells(current_row - 1, 1).Value <> Cells(current_row, 1).Value Then
            open_year = Cells(current_row, 3).Value
        End If

        'By using VBA "not equal <>" functionality on each row we can:
        If Cells(current_row + 1, 1).Value <> Cells(current_row, 1).Value Then
        
            'find the unique ticker name
            ticker_name = Cells(current_row, 1).Value
            
            'add total volume to our counter (for the last row for this ticker)
            total_volume = total_volume + Cells(current_row, 7).Value
            
            'find the closing price at the end of the year and assign it
            close_year = Cells(current_row, 6).Value
            
            'enter the found values in the desired cells in the Summary table
            Range("I" & summary_table_row).Value = ticker_name
            
            Range("L" & summary_table_row).Value = total_volume
            
            'calculate yearly changes of a price
            Range("J" & summary_table_row).Value = close_year - open_year
            'calculate the yearly percentage changes of a price
            Range("K" & summary_table_row).Value = ((close_year - open_year) / open_year)
            
            'Using an If statement to condition format desired cells. Here we add the colors
            If Range("J" & summary_table_row).Value > 0 Then
                Range("J" & summary_table_row).Interior.ColorIndex = 4
            Else
                Range("J" & summary_table_row).Interior.ColorIndex = 3
            End If
            
            'Using If statements to find the values for the Greatest values table
            If Range("K" & summary_table_row).Value > gr_perc_incr Then
                gr_perc_incr = Range("K" & summary_table_row).Value
                gr_perc_incr_ticker = ticker_name
            End If
            
             If Range("K" & summary_table_row).Value < gr_perc_decr Then
                gr_perc_decr = Range("K" & summary_table_row).Value
                gr_perc_decr_ticker = ticker_name
            End If
                
            If Range("L" & summary_table_row).Value > gr_ttl_vol Then
                gr_ttl_vol = Range("L" & summary_table_row).Value
                gr_ttl_vol_ticker = ticker_name
            End If
            
            'Adding 1 to our summary row counter to jump to the next row
            summary_table_row = summary_table_row + 1
            
            'Resetting total volume for the next ticker
            total_volume = 0
        
        'Adds total volume to our counter and keeps updating it
        Else
            total_volume = total_volume + Cells(current_row, 7).Value
            
        End If
        
    'To the next row
    Next current_row
    
    'Enter the found values in the desired cells for the Greatest values table
    Range("P" & 2).Value = gr_perc_incr_ticker
    Range("Q" & 2).Value = gr_perc_incr
    
    Range("P" & 3).Value = gr_perc_decr_ticker
    Range("Q" & 3).Value = gr_perc_decr
    
    Range("P" & 4).Value = gr_ttl_vol_ticker
    Range("Q" & 4).Value = gr_ttl_vol

    'Formatting % column and cells
    Range("K2:K" & last_row).NumberFormat = "0.00%"    
    Range("Q2:Q3").NumberFormat = "0.00%"
    
	'Auto-fit for columns
    Cells.Select
    Cells.EntireColumn.AutoFit
    Cells(1, 1).Select
    
    'To the next worksheet
    Next I
    
    'Returns to the first cell in the first worksheet
    Worksheets(1).Activate
    Worksheets(1).Cells(1, 1).Select
            
End Sub