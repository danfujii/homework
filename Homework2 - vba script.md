Sub stock_ticker()

    Dim tickername As String
    Dim stocktotal As Double
    stocktotal = 0
    Dim summary_table_row As Integer
    summary_table_row = 2
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Cells(1, 10).Value = "Ticker"
    Cells(1, 11).Value = "Total Stock Volume"
    
    
    
    For i = 2 To LastRow
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            tickername = Cells(i, 1).Value
            stocktotal = stocktotal + Cells(i, 7).Value
            Range("J" & summary_table_row).Value = tickername
            Range("K" & summary_table_row).Value = stocktotal
            summary_table_row = summary_table_row + 1
            stock_total = 0
            
        Else
            stocktotal = stocktotal + Cells(i, 7).Value
            
        End If
        
    Next i
    
End Sub
