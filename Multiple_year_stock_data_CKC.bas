Attribute VB_Name = "Multiple_year_stock_data_CKC"
Sub VBAHomework_MultiYear()

' Loop through all sheets
For Each ws In Worksheets

' Declare variables
Dim i As Double
Dim j As Double
Dim ticker As String
Dim volume As Double
Dim year_open As Double
Dim year_close As Double
Dim last_row As Double
Dim summary_table_row As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim great_inc As Double
Dim great_dec As Double
Dim great_vol As Double
Dim great_inc_ticker As String
Dim great_dec_ticker As String
Dim great_vol_ticker As String
Dim last_summary_row As Double
   
' Set summary table headers at the top of each worksheet
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Volume"

' Set bonus table headers at the top of each worksheet
ws.Cells(2, 16).Value = "Greatest % Increase"
ws.Cells(3, 16).Value = "Greatest % Decrease"
ws.Cells(4, 16).Value = "Greatest Total Volume"
ws.Cells(1, 17).Value = "Ticker"
ws.Cells(1, 18).Value = "Value"

    'Find the last row of raw data columns1-7 for use in the summary table
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Keep track of the location for printing values in the summary table, 3rd pointer
    summary_table_row = 2
    
    'Define 2nd pointer to track open price
    open_price_row = 2
    
    'Find the last row of summary table results in columns 9-12 for use in the bonus table
    last_summary_row = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'Initiate i loop to run from row 2 to last row in the summary table, 1st pointer
    For i = 2 To last_row
    
        'Test whether current cell is is equal to the cell in the next row
        If ws.Cells(i + 1, "A").Value <> ws.Cells(i, "A").Value Then
        
            'Stores the row the last ticker is on before new ticker appears
            ticker = ws.Cells(i, "A").Value
            
            'Add volume
            volume = volume + ws.Cells(i, "G").Value
            
            'find closing stock
            year_close = ws.Cells(i, "F").Value
            
            'find opening stock
            year_open = ws.Cells(open_price_row, "C").Value
                                
            'Display the ticker in the summary table
            ws.Cells(summary_table_row, "I").Value = ticker
                        
            'Display the volume in the summary table
            ws.Cells(summary_table_row, "L").Value = volume
            
            'calculate yearly change
            yearly_change = year_close - year_open
            
            'write yearly change
            ws.Cells(summary_table_row, "J").Value = yearly_change
            
                If yearly_change >= 0 Then
                    
                        ws.Cells(summary_table_row, "J").Interior.ColorIndex = 4
                
                Else
                        
                        ws.Cells(summary_table_row, "J").Interior.ColorIndex = 3
                
                End If
    
            'calculate percent change
            percent_change = (year_close - year_open) / year_open * 100
                        
            'write percent change
            ws.Cells(summary_table_row, "K").Value = "%" & percent_change
                                               
            'Add one to the summary table row
            summary_table_row = summary_table_row + 1
            
            'Add one to point to the next unique ticker
            open_price_row = i + 1
            
            'reset the total volume
            volume = 0
              
        'If the cell in the next row is the same ticker
        Else

            'Add to the volume and print
            volume = volume + ws.Cells(i, "G").Value
            
        'close if condition
        End If

    'close for i loop
    Next i

'set counter to 0 for bonus metrics
great_inc = 0
great_dec = 0
great_vol = 0

'set empty cell for ticker names for bonus metrics
great_inc_ticker = ""
great_dec_ticker = ""
great_vol_ticker = ""

    'initiate j loop to run from row 2 to last row of summary table
    For j = 2 To last_summary_row
 
        'If the current row's Percent Change is greater than 0
        If ws.Cells(j, "K").Value > great_inc Then
        
            ' then the greatest Percent Change increase is the current row
            great_inc = ws.Cells(j, "K").Value
            
            ' and the greatest Percent Change increase Ticker is the current row
            great_inc_ticker = ws.Cells(j, "I").Value
            
        ' close if loop
        End If
            
        ' If the current row's Percent Change is less than 0
        If ws.Cells(j, "K").Value < great_dec Then
        
            ' then the greatest Percent Change decrease is the current row
            great_dec = ws.Cells(j, "K").Value
            
            ' and the greatest Perccent Change decrease Ticker is the current row
            great_dec_ticker = ws.Cells(j, "I").Value
            
        ' close if loop
        End If
        
        ' If the current row's Volume is greater than 0
        If ws.Cells(j, "L").Value > great_vol Then
        
            ' then the greatest Volume is the current row
            great_vol = ws.Cells(j, "L").Value
            
            ' and the greatest Volume Ticker is the current row
            great_vol_ticker = ws.Cells(j, "I").Value
            
        ' close if loop
        End If
            
    ' close for j loop
    Next j
            
    ' once the j loop has finished going through each summary table row, print the final results
    ws.Cells(2, "Q").Value = great_inc_ticker
    ws.Cells(3, "Q").Value = great_dec_ticker
    ws.Cells(4, "Q").Value = great_vol_ticker

    ws.Cells(2, "R").Value = Format(great_inc, "0.00%")
    ws.Cells(3, "R").Value = Format(great_dec, "0.00%")
    ws.Cells(4, "R").Value = great_vol
                
' close for ws loop and move onto the next Worksheet
Next ws
 
End Sub



Sub ResetBoard()

' Loop through all sheets
For Each ws In Worksheets

' Declare variable for cells to be cleared
Dim rng As Range

' Set what columns contain the cells to be cleared
Set rng = ws.Range("I:R")

    ' Clear the text from the range
    rng.Value = ""
    
    ' Clear color formatting from the range
    rng.Interior.ColorIndex = 0
    
    ' Clear percentages and data type formatting from the range
    rng.Clear

' close for ws loop and move onto the next Worksheet
Next ws

End Sub


