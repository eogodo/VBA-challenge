Attribute VB_Name = "Module1"
Sub stock()
    For Each ws In Worksheets
    
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        Dim Srow As Long
        Dim currentrow As String
        Dim nextrow As String
        Dim volume As Single
        Dim openprice As Double
        Dim closeprice As Double
        Dim yearlychange As Double
        Dim percentchange As Double
        Dim lastrow As Long
        
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        sumrow = 2
        openprice = ws.Cells(2, 3).Value
        volume = 0
        
        For Srow = 2 To lastrow
            currentrow = ws.Cells(Srow, 1).Value
            nextrow = ws.Cells(Srow + 1, 1).Value
            volume = volume + ws.Cells(Srow, 7).Value
            
            If currentrow <> nextrow Then
                closeprice = ws.Cells(Srow, 6).Value
                
                ws.Cells(sumrow, 9).Value = currentrow
                
                yearlychange = closeprice - openprice
                ws.Cells(sumrow, 10).Value = yearlychange
                
                If yearlychange < 0 Then
                    ws.Cells(sumrow, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(sumrow, 10).Interior.ColorIndex = 4
                End If
                
                If openprice <> 0 Then
                    percentchange = ((closeprice - openprice) / openprice)
                Else
                    percentchange = 0
                End If
                
                ws.Cells(sumrow, 11).Value = percentchange
                ws.Cells(sumrow, 11).NumberFormat = "0.00%"
                
                ws.Cells(sumrow, 12).Value = volume
                sumrow = sumrow + 1
                volume = 0
                openprice = ws.Cells(Srow + 1, 3).Value
                
            End If
            
        Next Srow
        
        sumrow = sumrow - 1
        
        Dim i As Integer
        Dim gpinc, gpdec As Double
        Dim greatestvolume As Single
        
        gpinc = WorksheetFunction.Max(ws.Range("K2:K" & (sumrow)))
        gpdec = WorksheetFunction.Min(ws.Range("K2:K" & (sumrow)))
        greatestvolume = WorksheetFunction.Max(ws.Range("L2:L" & (sumrow)))
        
        For i = 2 To sumrow
        
            If gpinc = ws.Cells(i, 11).Value Then
                
                'greatest % increase
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(2, 17).Value = gpinc
                ws.Cells(2, 17).NumberFormat = "0.00%"
                
            End If
            
            If gpdec = ws.Cells(i, 11).Value Then
                'greatest % decrease
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(3, 17).Value = gpdec
                ws.Cells(3, 17).NumberFormat = "0.00%"
                
            End If
            
            'greatest total volume
            If greatestvolume = ws.Cells(i, 12).Value Then
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(4, 17).Value = greatestvolume
                
            End If
            
        Next i
        
    Next ws
End Sub

