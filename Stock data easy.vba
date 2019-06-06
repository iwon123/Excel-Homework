Sub Easy()

Dim ws As Worksheet

For Each ws In Worksheets

    Dim ticker As String
    Dim total As Double
    total = 0
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Total stock Vol"
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Summary = 2
    
    For i = 2 To 1000
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            ticker = ws.Cells(i, 1).Value
            total = total + ws.Cells(i, 7).Value
            
            ws.Range("I" & Summary).Value = ticker
            ws.Range("J" & Summary).Value = total
            
            Summary = Summary + 1
            
            total = 0
            
        Else
            total = total + ws.Cells(i, 7).Value
            
        End If
        
    Next i
    
Next
    
    
End Sub
