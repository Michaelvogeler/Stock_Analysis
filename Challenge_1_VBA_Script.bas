Attribute VB_Name = "Module1"
Sub ticker()

For Each ws In Worksheets

' declare variables
Dim yearly_change As Double
Dim yearly_change_open As Double
Dim yearly_change_close As Double
Dim ticker As String
Dim first_row_flag As Boolean
Dim current_output_table As Integer
Dim total_stock_volume As LongLong
Dim percent_change As Double
Dim greatest_percent_increase_ticker As String
Dim greatest_percent_decrease_ticker As String
Dim greatest_total_volume_ticker As String
Dim greatest_percent_increase_value As Double
Dim greatest_percent_decrease_value As Double
Dim greatest_total_volume_value As Double

'initialize variables
first_row_flag = True
current_output_table_row = 2
total_stock_volume = 0

'start iterating data table rows from first row to last row
'where i = row number
For I = 2 To 753001
    'running total for stock volume
    'add value from volume column to toal for each row
    total_stock_volume = total_stock_volume + ws.Cells(I, 7).Value
    
    'check if first row of ticker (only runs in first row of ticker)
    If first_row_flag = True Then
        'get value from open column of first ticker
        yearly_change_open = ws.Cells(I, 3).Value
        'set first row flag to false
        first_row_flag = False
    End If
    
    'Check if last row of Ticker (only runs if last row of ticker)
    If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
        ticker = ws.Cells(I, 1).Value
        yearly_change_close = ws.Cells(I, 6).Value
        yearly_change = yearly_change_close - yearly_change_open
        percent_change = ((yearly_change_close - yearly_change_open) / yearly_change_open)
        
        'Output to table
        ws.Cells(current_output_table_row, 9).Value = ticker
        ws.Cells(current_output_table_row, 10).Value = yearly_change
        ws.Cells(current_output_table_row, 12).Value = total_stock_volume
        ws.Cells(current_output_table_row, 11).Value = percent_change
        
        'format cells to percentage
        ws.Cells(current_output_table_row, 11).NumberFormat = "0.00%"
        
        'format cell colors
        If yearly_change > 0 Then
            ws.Cells(current_output_table_row, 10).Interior.ColorIndex = 4
        ElseIf yearly_change < 0 Then
           ws.Cells(current_output_table_row, 10).Interior.ColorIndex = 3
        Else
            ws.Cells(current_output_table_row, 10).Interior.ColorIndex = 2
        End If
        
        'increase output table iterator
        current_output_table_row = current_output_table_row + 1
        
        'reset first row flag
        first_row_flag = True
        'reset total volume
        total_stock_volume = 0
    End If
    
Next I

For I = 2 To 2977
    If I = 2 Then
        greatest_percent_increase_ticker = ws.Cells(I, 9).Value
        greatest_percent_decrease_ticker = ws.Cells(I, 9).Value
        greatest_total_volume_ticker = ws.Cells(I, 9).Value
        greatest_percent_increase_value = ws.Cells(I, 11).Value
        greatest_percent_decrease_value = ws.Cells(I, 11).Value
        greatest_total_volume_value = ws.Cells(I, 12).Value
    Else
        If Cells(I, 11).Value > greatest_percent_increase_value Then
            greatest_percent_increase_value = ws.Cells(I, 11).Value
            'New
            greatest_percent_increase_ticker = ws.Cells(I, 9).Value
        End If
        If ws.Cells(I, 11).Value < greatest_percent_decrease_value Then
            greatest_percent_decrease_value = ws.Cells(I, 11).Value
            greatest_percent_decrease_ticker = ws.Cells(I, 9).Value
        End If
        If greatest_total_volume_value > ws.Cells(I, 12).Value Then
            greatest_total_volume_value = ws.Cells(I, 12).Value
            greatest_total_volume_ticker = ws.Cells(I, 9).Value
        End If
    
    End If

Next I

ws.Cells(2, 15).Value = greatest_percent_increase_ticker
ws.Cells(3, 15).Value = greatest_percent_decrease_ticker
ws.Cells(4, 15).Value = greatest_total_volume_ticker
ws.Cells(2, 16).Value = greatest_percent_increase_value
ws.Cells(3, 16).Value = greatest_percent_decrease_value
ws.Cells(4, 16).Value = greatest_total_volume_value

ws.Cells(2, 16).NumberFormat = "0.00%"
ws.Cells(3, 16).NumberFormat = "0.00%"



Next ws

End Sub
