Attribute VB_Name = "Module1"
Sub StockAccumulator()
    Dim current As Worksheet
        
    Dim greatest_increase_ticker As String
    Dim greatest_increase As Double
    greatest_increase = 0
    
    Dim greatest_decrease_ticker As String
    Dim greatest_decrease As Double
    greatest_decrease = 0
    
    Dim greatest_volume_ticker As String
    Dim greatest_volume As Double
    greatest_volume = 0
    
    For Each current In Worksheets
        current.Range("I1").Value = "Ticker"
        current.Range("J1").Value = "Yearly Change"
        current.Range("K1").Value = "Percent Change"
        current.Range("L1").Value = "Total Stock Volume"
        
        ' the row we are currently evaluating
        Dim stock_day As Range
        'the stored ticker
        Dim ticker As String
        'the stored volume
        Dim volume As Double
        'the stored opening price
        Dim opening_price As Double
        'the stored closing price
        Dim closing_price As Double
        
        'remove header from worksheet
        Dim data_range As Range
        Set data_range = current.UsedRange.rows.Offset(2)
        
        Dim output_row As Integer
        output_row = 2
        
        'get initial row
        ticker = current.UsedRange.rows.Cells(2, 1)
        volume = 0
        opening_price = current.UsedRange.rows.Cells(2, 3)
        
        
        'loop through the rest of the rows
        For Each stock_row In data_range
            'the stock we are evalutating
            Dim stock_ticker As String
            stock_ticker = stock_row.rows.Cells(1)
                        
            'are we still on the same stock?
            If stock_ticker <> ticker Then
                'no, we have hit a different stock so we have to calculate the needed data  and copy to the current sheet
                Dim percent_change As Double
                percent_change = 0
                'can't divide by 0
                If opening_price > 0 Then
                    percent_change = (closing_price / opening_price - 1)
                End If
                
                Dim total_change As Double
                total_change = closing_price - opening_price
                
                'copy data to primary table
                current.Cells(output_row, 9).Value = ticker
                current.Cells(output_row, 10).Value = total_change
                If total_change < 0 Then current.Cells(output_row, 10).Interior.Color = RGB(205, 50, 50)
                If total_change > 0 Then current.Cells(output_row, 10).Interior.Color = RGB(50, 205, 50)
                current.Cells(output_row, 11).Value = percent_change
                current.Cells(output_row, 11).NumberFormat = "0.00%"
                current.Cells(output_row, 12).Value = volume
                output_row = output_row + 1
                
                'check for greatest stocks
                If percent_change > greatest_increase Then
                    greatest_increase = percent_change
                    greatest_increase_ticker = ticker
                End If
                
                If percent_change < greatest_decrease Then
                    greatest_decrease = percent_change
                    greatest_decrease_ticker = ticker
                End If
                
                If volume > greatest_volume Then
                    greatest_volume = volume
                    greatest_volume_ticker = ticker
                End If
                
                'reset stock information for next stock
                ticker = stock_ticker
                volume = 0
                opening_price = stock_row.rows.Cells(3)
            Else 'yes, we are still on the same stock so increase volume and update closing price
                volume = volume + stock_row.rows.Cells(7)
                closing_price = stock_row.rows.Cells(6)
            End If
        Next
    Next
    
    'greatest are found after all loops are finished
    'copy data to primary table
    Set current = Worksheets(1)
    current.Range("n2").Value = "Greatest % Increase"
    current.Range("n3").Value = "Greatest % Decrease"
    current.Range("n4").Value = "Greatest Total Volume"
    
    current.Range("o1").Value = "Ticker"
    current.Range("o2").Value = greatest_increase_ticker
    current.Range("o3").Value = greatest_decrease_ticker
    current.Range("o4").Value = greatest_volume_ticker
    
    current.Range("p1").Value = "Value"
    current.Range("p2").Value = greatest_increase
    current.Range("p2").NumberFormat = "0.00%"
    current.Range("p3").Value = greatest_decrease
    current.Range("p3").NumberFormat = "0.00%"
    current.Range("p4").Value = greatest_volume
    
    Debug.Print "All done."
End Sub

