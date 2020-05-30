Attribute VB_Name = "Module1"
Sub stockMarketAnalyst():

'Loop through all ws's
Dim ws As Worksheet
For Each ws In worksheets
    
    'Title new columns
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    'Challenge
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    'Formatting
    ws.Range("k:k").NumberFormat = "0.00%"
    ws.Range("I1:L1").Font.Bold = True
    ws.Range("I1:L1").HorizontalAlignment = xlCenter
    ws.Range("O2:O4").Font.Bold = True
    ws.Range("O2:O4").HorizontalAlignment = xlCenter
    ws.Range("P1:Q1").Font.Bold = True
    ws.Range("P1:Q1").HorizontalAlignment = xlCenter
    
    'Number of Rows & Columns
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    last_column = ws.Cells(1, Columns.Count).End(xlToLeft).Column
    
    'Initiate Variables
    ticker_control = ws.Cells(2, 1).Value
    ws.Cells(2, 9).Value = "A"
    opening_price = ws.Cells(2, 3).Value
    Count = 3
    volume_sum = 0
    greatest_percent_increase = 0
    greatest_percent_decrease = 0
    greatest_total_volume = 0
    
    'Loop through all rows
    For i = 2 To last_row
    
        ticker_current = ws.Cells(i, 1).Value
        
        ' if cells(i,1).value <> cells(i+1,1) then
        If ticker_current <> ticker_control Then
        
            'List new ticker
            ws.Cells(Count, 9).Value = ticker_current
            
            'Opening vs closing price
            closing_price = ws.Cells(i - 1, 6).Value
            price_change = closing_price - opening_price
            ws.Cells(Count - 1, 10).Value = price_change
            If price_change < 0 Then
                ws.Cells(Count - 1, 10).Interior.Color = vbRed
            ElseIf price_change > 0 Then
                ws.Cells(Count - 1, 10).Interior.Color = vbGreen
            End If
            
            'Percent Change
            If opening_price = 0 Or price_change = 0 Then
                Percent_change = 0
            Else
                Percent_change = price_change / opening_price
            End If
                
            ws.Cells(Count - 1, 11).Value = Percent_change
            
            If Percent_change > greatest_percent_increase Then
            
                greatest_percent_increase = Percent_change
                greatest_percent_increase_ticker = ticker_control
                
            ElseIf Percent_change < greatest_percent_decrease Then
            
                greatest_percent_decrease = Percent_change
                greatest_percent_decrease_ticker = ticker_control
                
            End If
            
            'Volume
            ws.Cells(Count - 1, 12).Value = volume_sum
            
            'Update ticker, count, opening price, volume
            ticker_control = ticker_current
            Count = Count + 1
            opening_price = ws.Cells(i, 3).Value
            volume_sum = 0
            
        End If
          
        'Volume
        volume = ws.Cells(i, 7).Value
        volume_sum = volume_sum + volume
        If volume_sum > greatest_total_volume Then
        
            greatest_total_volume = volume_sum
            greatest_total_volume_ticker = ticker_control
            
        End If
        
    Next i

    'Change in price Last value
    last_closing_price = ws.Cells(i - 1, 6).Value
    last_price_change = last_closing_price - opening_price
    ws.Cells(Count - 1, 10).Value = last_price_change
    If last_price_change < 0 Then
        ws.Cells(Count - 1, 10).Interior.Color = vbRed
    ElseIf price_change > 0 Then
        ws.Cells(Count - 1, 10).Interior.Color = vbGreen
    End If

    'Percent Change last value
    last_percent_change = last_price_change / opening_price
    ws.Cells(Count - 1, 11).Value = last_percent_change
    If last_percent_change > greatest_percent_increase Then
            
        greatest_percent_increase = last_percent_change
        greatest_percent_increase_ticker = ticker_control
                
    ElseIf last_percent_change < greatest_percent_decrease Then
            
        greatest_percent_decrease = last_percent_change
        greatest_percent_decrease_ticker = ticker_control

    End If

    'Challenge Input values
    ws.Range("Q2").Value = greatest_percent_increase
    ws.Range("Q3").Value = greatest_percent_decrease
    ws.Range("Q4").Value = greatest_total_volume
    ws.Range("P2").Value = greatest_percent_increase_ticker
    ws.Range("P3").Value = greatest_percent_decrease_ticker
    ws.Range("P4").Value = greatest_total_volume_ticker

    'Challenge Formatting
    ws.Range("Q2:Q3").NumberFormat = "0.00%"


    'Volume last value
    ws.Cells(Count - 1, 12).Value = volume_sum

    'Autofit
    ws.Range("J:L").Columns.AutoFit
    ws.Range("N:O").Columns.AutoFit

Next ws

End Sub

