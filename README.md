Sub Stock_Market():
'Create loop for worksheets

Dim ws As Worksheet
For Each ws In ActiveWorkbook.Worksheets
ws.Activate

'Declare variables

Dim Ticker As String
Dim open_price As Double
Dim close_price As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim total_stock_volume As Double
Dim summary_table_row As Double

total_stock_volume = 0
summary_table_row = 2

'Creat columns

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'Start Loop for calculations

For i = 2 To lastrow

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    Ticker = Cells(i, 1).Value
    
    total_stock_volume = total_stock_volume + Cells(i, 7).Value
    
    Range("I" & summary_table_row).Value = Ticker
    Range("L" & summary_table_row).Value = total_stock_volume
    
    close_price = Cells(i, 6).Value
    
    yearly_change = close_price - open_price
    
    If open_price = 0 Then
        percent_change = 0
    Else
        percent_change = (yearly_change / open_price)
    End If
    
    Range("J" & summary_table_row).Value = yearly_change
    Range("K" & summary_table_row).Value = percent_change
    Range("K" & summary_table_row).NumberFormat = "0.00%"
    
    summary_table_row = summary_table_row + 1
    total_stock_volume = 0
    open_price = Cells(i + 1, 3).Value
    
Else
    total_stock_volume = total_stock_volume + Cells(i, 7).Value
    
    End If

Next i
'Color code


summary_table_last_row = Cells(Rows.Count, 9).End(xlUp).Row

    For i = 2 To summary_table_last_row
        If Cells(i, 10).Value > 0 Then
            Cells(i, 10).Interior.ColorIndex = 4
        Else
            Cells(i, 10).Interior.ColorIndex = 3
            
        End If
    
    Next i
    
Next ws


End Sub

