Attribute VB_Name = "Module1"
Sub challange2()
For Each ws In Worksheets
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        Dim table As Integer
        Dim ticker As String
        Dim tickertotal As Double
        ws.Cells(1, 10).Value = "Ticker"
        ws.Cells(1, 11).Value = "Yearly change"
        ws.Cells(1, 12).Value = "Percentage Change"
        ws.Cells(1, 13).Value = "total stock volume"
        tablerow = 2
        stockvolume = 0
        
        For i = 2 To lastrow
            Dim ispricecaptured As Boolean
            Dim openingprice As Double
    
            If ispricecaptured = False Then
                openingprice = ws.Cells(i, 3).Value
                ispricecaptured = True
            End If
            
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                tickertotal = ws.Cells(i, 6).Value - openingprice
                ws.Range("K" & tablerow).Value = tickertotal
                ws.Range("J" & tablerow).Value = ticker
                stockvolume = stockvolume + ws.Cells(i, 7).Value
                ws.Range("M" & tablerow).Value = stockvolume
                ispricecaptured = False
                stockvolume = 0
                Dim change As Double
                Dim price As Double
                change = tickertotal
                price = openingprice
                ws.Range("L" & tablerow).Value = tickertotal / openingprice
                ws.Range("L" & tablerow).Style = "Percent"
                If ws.Range("k" & tablerow).Value < 0 Then
                    ws.Range("k" & tablerow).Interior.ColorIndex = 3
                    Else
                    ws.Range("k" & tablerow).Interior.ColorIndex = 4
                End If
                tablerow = tablerow + 1
                Else
                stockvolume = stockvolume + ws.Cells(i, 7).Value
            End If
           
                        
        Next i
    Next ws
                
            
        

End Sub
