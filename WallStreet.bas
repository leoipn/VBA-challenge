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
 Dim colorIndex As Integer
 Dim greatest_total_volume As Double
 Dim greatest_total_volume_ticker As Double
 
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
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
 
 ' Iterate over the cells to get the requested info
 For i = 2 To last_row
    
    ticker = Cells(i, 1).Value
    next_ticker = Cells(i + 1, 1).Value
    volume = Cells(i, 7).Value
    total_volume = total_volume + volume
    colorIndex = 4 ' Green color
    percent_change = 0
        
    ' Evaluate the actual cell with the following one to get just one of each ticker
    If (ticker <> next_ticker) Then
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
        If (yearly_change < 0) Then
            colorIndex = 3
        End If
        Range("J" & single_ticker_row_index).Interior.colorIndex = colorIndex
        
        ' -------------------------------------------------------------------
        ' Percent change logic
        ' -------------------------------------------------------------------
        ' Condition to avoid divided by zero error
        If (year_open > 0) Then
            percent_change = (year_open - year_close) / year_open
        End If
        
        ' Set the percent change in the corresponding row
        Cells(single_ticker_row_index, 11).Value = percent_change
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
        total_tickers = total_tickers + 1
    
    End If
        
     'If i = 1000 Then
     '   Exit For
     'End If
    
 Next i
 
 greatest_total_volume = WorksheetFunction.Max(Range("L1:L" & total_tickers))
 greatest_total_volume_ticker = WorksheetFunction.Match(greatest_total_volume, Range("L:L"), 0)
 
 Range("P4").Value = Range("I" & greatest_total_volume_ticker).Value
 Range("Q4").Value = greatest_total_volume

End Sub
