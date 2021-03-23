Sub wallStreet()
 
    Dim ticker As String
    Dim last_row As Long
    Dim single_ticker_row_index As Integer
    Dim year_open As Double
    Dim year_close As Double
    Dim yearly_change As Double
    Dim volumen As Double
    Dim total_volumen As Double
    Dim colorIndex As Integer
    
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    single_ticker_row_index = 2
    year_open = Cells(2, 3)
    total_volume = 0
    total_tickers = 0
 
    ' -------------------------------------------------------------------
    ' Set the headers
    ' -------------------------------------------------------------------
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    ' Autofit the columns
    Columns("A:Q").AutoFit
 
    ' -------------------------------------------------------------------
    ' Iterate over the cells to get the requested info
    ' -------------------------------------------------------------------
    For i = 2 To last_row
       
       ticker = Cells(i, 1).Value
       next_ticker = Cells(i + 1, 1).Value
       volume = Cells(i, 7).Value
       total_volume = total_volume + volume
       colorIndex = 4 ' Green color
       percent_change = 0
           
       ' Evaluate the actual cell with the following one to get just one of each ticker
       If (ticker <> next_ticker) Then
           ' Sets the single ticker value
           Cells(single_ticker_row_index, 9).Value = ticker
           
           ' -------------------------------------------------------------------
           ' Yearly change logic
           ' -------------------------------------------------------------------
           ' Get the last year close for each ticker
           year_close = Cells(i, 6).Value
           
           ' Get the yearly change
           yearly_change = year_close - year_open
           
           ' Set the yearly change in the corresponding row
           Cells(single_ticker_row_index, 10).Value = yearly_change
           
           ' Evaluates the yearly change to set the correct color
           If (yearly_change < 0) Then
               colorIndex = 3
           End If
           
           ' Set the correct color
           Range("J" & single_ticker_row_index).Interior.colorIndex = colorIndex
           
           ' -------------------------------------------------------------------
           ' Percent change logic
           ' -------------------------------------------------------------------
           ' Condition to avoid divided by zero error
           If (year_open > 0) Then
               percent_change = (year_close - year_open) / year_open
           End If
           
           ' Set the percent change in the corresponding row
           Cells(single_ticker_row_index, 11).Value = percent_change
           
           ' Apply the percentage format
           Range("K" & single_ticker_row_index).NumberFormat = "0.00%"
                   
           ' -------------------------------------------------------------------
           ' Volume logic
           ' -------------------------------------------------------------------
           ' Set the total volume
           Cells(single_ticker_row_index, 12).Value = total_volume
           
           ' -------------------------------------------------------------------
           ' Increasing and reset logic
           ' -------------------------------------------------------------------
           ' Get the following first year open for each ticker
           year_open = Cells(i + 1, 3)
           ' Reset the total volume
           total_volume = 0
           ' Increase the row number
           single_ticker_row_index = single_ticker_row_index + 1
       
       End If
       
    Next i
    
End Sub
