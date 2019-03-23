Sub StockData_Easy()

For Each ws In ActiveWorkbook.Worksheets

Dim Ticker_Value As String
Dim Total_Volume As Double
Total_Volume = 0

Dim Summary_Table_Row As Long
Summary_Table_Row = 2

Dim Lastrow As Long
Lastrow = Range("A" & Rows.Count).End(xlUp).Row

    For i = 2 To Lastrow

        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
    
        Total_Volume = Total_Volume + ws.Cells(i, 7).Value
    
        Ticker_Value = ws.Cells(i, 1).Value
        
        
        ws.Range("K1").Value = "Ticker Value"
        ws.Range("L1").Value = "Total Volume"
        ws.Range("K" & Summary_Table_Row).Value = Ticker_Value
        ws.Range("L" & Summary_Table_Row).Value = Total_Volume
    
        Summary_Table_Row = Summary_Table_Row + 1
    
        Total_Volume = 0

    Else

    Total_Volume = Total_Volume + ws.Cells(i, 7).Value
    
    End If

Next i

Next ws

End Sub

