Attribute VB_Name = "Módulo1"
Sub wallStreet()
 
 Dim ticker As String
 Dim last_row As Long
 Dim single_ticker_row_index As Integer
 Dim year_open As Double
 Dim year_close As Double
 Dim yearly_change As Double
 Dim volumen As Double
 Dim total_volumen As Double
 
 last_row = Cells(Rows.Count, 1).End(xlUp).Row
 single_ticker_row_index = 2
 year_open = Cells(2, 3)
 volumen = Cells(2, 7).Value
 total_volumen = 0
 
 ' Iterate over the cells to get the requested info
 For i = 2 To last_row
    
    ticker = Cells(i, 1).Value
    next_ticker = Cells(i + 1, 1).Value
    total_volumen = total_volumen + volumen
    
    ' Evaluate the actual cell with the following one to get just one of each ticker
    If (ticker <> next_ticker) Then
        Cells(single_ticker_row_index, 9).Value = ticker
        
        ' Get the last year close for each ticker
        year_close = Cells(i, 6).Value
        
        ' Get the yearly change
        yearly_change = year_close - year_open
        
        If (year_open > 0) Then
            percent_change = (year_open - year_close) / year_open
        Else
            percent_change = 0
        End If
        
        ' Set the yearly change in the corresponding row
        Cells(single_ticker_row_index, 10).Value = yearly_change
        
        ' Set the percent change in the corresponding row
        Cells(single_ticker_row_index, 11).Value = percent_change
        Range("K" & single_ticker_row_index).NumberFormat = "0.00%"
        
        ' Set the total volumen
        Cells(single_ticker_row_index, 12).Value = total_volumen
        
        ' Get the following first year open for each ticker
        year_open = Cells(i + 1, 3)
        volumen = Cells(i + 1, 7)
        total_volumen = 0
        
        ' Increase the row number
        single_ticker_row_index = single_ticker_row_index + 1
    
    End If
        
    ' If i = 600 Then
    '    Exit For
    ' End If
    
 Next i
 
End Sub
