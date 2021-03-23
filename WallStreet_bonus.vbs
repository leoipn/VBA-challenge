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
    Dim ws As Worksheet

    For Each ws In Worksheets
    
         last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
         single_ticker_row_index = 2
         year_open = ws.Cells(2, 3)
         total_volume = 0
         total_tickers = 0
         
        ' -------------------------------------------------------------------
        ' Set the headers
        ' -------------------------------------------------------------------
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        ' Autofit the columns
        ws.Columns("A:Q").AutoFit
         
        ' -------------------------------------------------------------------
        ' Iterate over the cells to get the requested info
        ' -------------------------------------------------------------------
         For i = 2 To last_row
            
            ticker = ws.Cells(i, 1).Value
            next_ticker = ws.Cells(i + 1, 1).Value
            volume = ws.Cells(i, 7).Value
            total_volume = total_volume + volume
            colorIndex = 4 ' Green color
            percent_change = 0
                
            ' Evaluate the actual cell with the following one to get just one of each ticker
            If (ticker <> next_ticker) Then
                ws.Cells(single_ticker_row_index, 9).Value = ticker
                
                
                ' -------------------------------------------------------------------
                ' Yearly change logic
                ' -------------------------------------------------------------------
                ' Get the last year close for each ticker
                year_close = ws.Cells(i, 6).Value
                
                ' Get the yearly change
                yearly_change = year_close - year_open
                
                ' Set the yearly change in the corresponding row
                ws.Cells(single_ticker_row_index, 10).Value = yearly_change
                
                ' Evaluates the yearly change to set the correct color
                If (yearly_change < 0) Then
                    colorIndex = 3
                End If
                
                ' Set the correct color
                ws.Range("J" & single_ticker_row_index).Interior.colorIndex = colorIndex
                
                ' -------------------------------------------------------------------
                ' Percent change logic
                ' -------------------------------------------------------------------
                ' Condition to avoid divided by zero error
                If (year_open > 0) Then
                    percent_change = (year_close - year_open) / year_open
                End If
                
                ' Set the percent change in the corresponding row
                ws.Cells(single_ticker_row_index, 11).Value = percent_change
                
                ' Apply the percentage format
                ws.Range("K" & single_ticker_row_index).NumberFormat = "0.00%"
                        
                ' -------------------------------------------------------------------
                ' Volume logic
                ' -------------------------------------------------------------------
                ' Set the total volume
                ws.Cells(single_ticker_row_index, 12).Value = total_volume
                
                ' -------------------------------------------------------------------
                ' Increasing and reset logic
                ' -------------------------------------------------------------------
                ' Get the following first year open for each ticker
                year_open = ws.Cells(i + 1, 3)
                ' Reset the total volume
                total_volume = 0
                ' Increase the row number
                single_ticker_row_index = single_ticker_row_index + 1
                ' Count the total of tickers of a single group
                total_tickers = total_tickers + 1
            
            End If
                
         Next i
         
        ' -------------------------------------------------------------------
        ' Greatest Percentage Increase
        ' -------------------------------------------------------------------
         greatest_percent_increase = WorksheetFunction.Max(ws.Range("K1:K" & total_tickers))
         greatest_percent_increase_ticker = WorksheetFunction.Match(greatest_percent_increase, ws.Range("K:K"), 0)
         
         ws.Range("P2").Value = ws.Range("I" & greatest_percent_increase_ticker).Value
         ws.Range("Q2").Value = greatest_percent_increase
         ws.Range("Q2").NumberFormat = "0.00%"
         
        ' -------------------------------------------------------------------
        ' Greatest Percentage Decrease
        ' -------------------------------------------------------------------
         greatest_percent_decrease = WorksheetFunction.Min(ws.Range("K1:K" & total_tickers))
         greatest_percent_decrease_ticker = WorksheetFunction.Match(greatest_percent_decrease, ws.Range("K:K"), 0)
         
         ws.Range("P3").Value = ws.Range("I" & greatest_percent_decrease_ticker).Value
         ws.Range("Q3").Value = greatest_percent_decrease
         ws.Range("Q3").NumberFormat = "0.00%"
         
        ' -------------------------------------------------------------------
        ' Greatest Total Volume
        ' -------------------------------------------------------------------
         greatest_total_volume = WorksheetFunction.Max(ws.Range("L1:L" & total_tickers))
         greatest_total_volume_ticker = WorksheetFunction.Match(greatest_total_volume, ws.Range("L:L"), 0)
         
         ws.Range("P4").Value = ws.Range("I" & greatest_total_volume_ticker).Value
         ws.Range("Q4").Value = greatest_total_volume
    
    Next ws

End Sub
