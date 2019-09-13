Attribute VB_Name = "Mod_ticker_totals"

Sub Stock_Analysis_tickers()

    Dim ws As Worksheet

    For Each ws In Worksheets
    
        'activate worksheet
        ws.Activate
    
        'set variables
        Dim TickerID As String
        Dim Volume_total As Double
        Dim Ticker_counter As Long
        Dim open_price As Double
        Dim close_price As Double
        Dim price_change As Double
        Dim percent_change As Double
        
        
        'initialize variables
        Ticker_counter = 2
        Volume_total = 0
        open_price = 0
        close_price = 0
        
        
        
        'set headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
    
        'Find end of rows
        Dim lastrow As Long
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
            'loop thru ticker data
            For i = 2 To lastrow
            
                Volume_total = Volume_total + ws.Cells(i, 7).Value
            
                'obtain opening price
                If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then

                open_price = ws.Cells(i, 3).Value

                End If
            
                If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                
                    ws.Cells(Ticker_counter, 9).Value = ws.Cells(i, 1).Value
                    
                    ws.Cells(Ticker_counter, 12).Value = Volume_total
                
                    'obtain closing price
                    close_price = ws.Cells(i, 6).Value
                    
                    'calculating open to closing price change
                    price_change = close_price - open_price
                    
                    'calculate percent change from open to close price
                    If open_price <> 0 Then
                    
                    percent_change = price_change / open_price
                    
                    'Else
                    
                        
                    
                    End If
                    
                    ws.Cells(Ticker_counter, 10).Value = price_change
                    ws.Cells(Ticker_counter, 11).Value = percent_change
                    
                    ws.Cells(Ticker_counter, 11).NumberFormat = "0.00%"
                    
                    
                    'formatting cells that are greater then 0 = green
                    'and less then 0 = red
                    If price_change >= 0 Then
                        ws.Cells(Ticker_counter, 10).Interior.Color = vbGreen
                    Else
                        ws.Cells(Ticker_counter, 10).Interior.Color = vbRed
                    End If
                    
                
                    Ticker_counter = Ticker_counter + 1
                
                    Volume_total = 0
                    'price_change = 0
                    'percent_change = 0
                
                End If
                
            Next i
        
        ' Performance table setup
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"

        'Find end of rows for performance summary table
        lastrow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
        'set variables
        Dim great_percent_inc As Double
        Dim great_percent_dec As Double
        Dim great_total_volume As Double
        Dim best_tickerID As String
        Dim worst_tickerID As String
        Dim great_total_tickerID As String
        
        'initalizing first stock value
        great_percent_inc = ws.Cells(2, 11).Value
        great_percent_dec = ws.Cells(2, 11).Value
        great_total_volume = ws.Cells(2, 12).Value
        
        For j = 2 To lastrow2

            'Determine best performer
            If ws.Cells(j, 11).Value > great_percent_inc Then
                great_percent_inc = ws.Cells(j, 11).Value
                best_tickerID = ws.Cells(j, 9).Value
                
            End If

            'Determine worst performer
            If ws.Cells(j, 11).Value < great_percent_dec Then
                great_percent_dec = ws.Cells(j, 11).Value
                worst_tickerID = ws.Cells(j, 9).Value
                
            End If

            'Determine stock with the greatest volume traded
            If ws.Cells(j, 12).Value > great_total_volume Then
                great_total_volume = ws.Cells(j, 12).Value
                great_total_tickerID = ws.Cells(j, 9).Value
                
            End If

        Next j
    
        'printing data to performance table
        ws.Cells(2, 16).Value = best_tickerID
        ws.Cells(2, 17).Value = great_percent_inc
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 16).Value = worst_tickerID
        ws.Cells(3, 17).Value = great_percent_dec
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(4, 16).Value = great_total_tickerID
        ws.Cells(4, 17).Value = great_total_volume
    
        'Autoformatting columns
        ws.Columns("I:L").EntireColumn.AutoFit
        ws.Columns("O:Q").EntireColumn.AutoFit
    
    Next ws
    

End Sub

