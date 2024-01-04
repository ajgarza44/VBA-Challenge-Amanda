Sub tickerloop():

    Dim tickername As String
    
    Dim tickervolume As Double
    tickervolume = 0
    
    Dim summary_ticker_row As Integer
    summary_ticker_row = 2
    
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Total Stock Volume"
    
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To LastRow
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            tickername = Cells(i, 1).Value
            
            tickervolume = tickervolume + Cells(i, 7).Value
            
            Range("I" & summary_ticker_row).Value = tickername
            
            Range("J" & summary_ticker_row).Value = tickervolume
            
            summary_ticker_row = summary_ticker_row + 1
            
            tickervolume = 0
            
        Else
            
            tickervolume = tickervolume + Cells(i, 7).Value
            
        End If
        
    Next i
    
End Sub
