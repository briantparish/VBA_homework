Sub Calculate_stocks()
'Define variables
    Dim ticker As String
    Dim volume As Double
    Dim ticker_index As Long
    Dim start_price As Double
    Dim daily_price As Double
    Dim end_price As Double
    Dim difference As Double
    Dim percent_change As Double
    Dim biggest_winner_percent As Double
    Dim biggest_winnder_ticker As String
    Dim biggest_loser_percent As Double
    Dim biggest_loser_ticker As String
    Dim biggest_mover_volume As Double
    Dim biggest_mover_ticker As String

    'Initialize first stock
    ticker_index = 2
    volume = 0
    start_price = Cells(2, 3)
    daily_price = Cells(2, 3)
    Cells(ticker_index, 10) = ticker
    Cells(ticker_index, 3) = start_price
    
    'Initialize record holders
    biggest_winner_percent = 0
    biggest_loser_percent = 0
    biggest_mover_volume = 0
    
    'label columns
    Cells(1, 10) = "Ticker"
    Cells(1, 11) = "Yearly Change"
    Cells(1, 12) = "Percent Change"
    Cells(1, 13) = "Total Stock Volume"
    
    ticker = Cells(ticker_index, 1)
    Cells(ticker_index, 10) = ticker
    
    'Loop through every row in the spreadsheet
    Dim i As Long
    For i = 3 To Rows.Count
        If ticker = Cells(i, 1) Then 'Same ticker as previous, not much to do
            volume = volume + Cells(i, 7) 'Increase volume
            daily_price = Cells(i, 6) 'Set end of day price
        Else
            'New ticker! Lots to do! Calculate annual price changes and print values for previous ticker
            Cells(ticker_index, 13) = volume
            end_price = daily_price 'daily price is from previous ticker end of year price
            difference = end_price - start_price
            Cells(ticker_index, 11) = difference
            If start_price > 0 Then   'make sure start price wasn't 0 to avoid divide by 0
                percentage_change = (difference / start_price)
                Cells(ticker_index, 12) = percentage_change
                Cells(ticker_index, 12).NumberFormat = "0.00%" 'make it pretty
            Else
                Cells(ticker_index, 12) = "N/A" 'Start price was 0, can't calculate % change
            End If
            
            'add color scheme
            If difference > 0 Then
                Cells(ticker_index, 11).Interior.ColorIndex = 4 'green
            ElseIf difference < 0 Then
                Cells(ticker_index, 11).Interior.ColorIndex = 3 'red
            End If
            
            'Check to see if it holds a record
            If percentage_change > biggest_winner_percent Then
                biggest_winner_ticker = ticker
                biggest_winner_percent = percentage_change
            End If
            
            If percentage_change < biggest_loser_percent Then
                biggest_loser_ticker = ticker
                biggest_loser_percent = percentage_change
            End If
            
            If volume > biggest_mover_volume Then
                biggest_mover_ticker = ticker
                biggest_mover_volume = volume
            End If
            'End of checking for record holder
            
            'Reset Values to get ready for next ticker
            ticker_index = ticker_index + 1
            ticker = Cells(i, 1)
            Cells(ticker_index, 10) = Cells(i, 1) 'label ticker in spreadsheet
            volume = 0
            start_price = Cells(i, 3)
            daily_price = Cells(i, 6)
        End If
    Next i
    
    'Populate record holders
    Cells(2, 15) = "Greatest % Increase"
    Cells(2, 16) = biggest_winner_ticker
    Cells(2, 17) = biggest_winner_percent
    Cells(2, 17).NumberFormat = "0.00%"
    
    Cells(3, 15) = "Greatest % decrease"
    Cells(3, 16) = biggest_loser_ticker
    Cells(3, 17) = biggest_loser_percent
    Cells(3, 17).NumberFormat = "0.00%"
    
    Cells(4, 15) = "Greatest Total Volume"
    Cells(4, 16) = biggest_mover_ticker
    Cells(4, 17) = biggest_mover_volume
    'done
End Sub
