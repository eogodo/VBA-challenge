Attribute VB_Name = "Module1"
Sub stock()

    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    Dim Srow As Long
    Dim currentrow As String
    Dim nextrow As String
    Dim volume As Single
    Dim openprice As Double
    Dim closeprice As Double
    Dim yearlychange As Double
    Dim percentchange As Double
    Dim lastrow As Long
    
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    sumrow = 2
    openprice = Cells(2, 3).Value
    volume = 0
    
    For Srow = 2 To lastrow
        currentrow = Cells(Srow, 1).Value
        nextrow = Cells(Srow + 1, 1).Value
        volume = volume + Cells(Srow, 7).Value
        
        If currentrow <> nextrow Then
            closeprice = Cells(Srow, 6).Value
            
            Cells(sumrow, 9).Value = currentrow
            
            yearlychange = closeprice - openprice
            Cells(sumrow, 10).Value = yearlychange
            
            If yearlychange < 0 Then
                Cells(sumrow, 10).Interior.ColorIndex = 3
            Else
                Cells(sumrow, 10).Interior.ColorIndex = 4
            End If
            
            If openprice <> 0 Then
                percentchange = ((closeprice - openprice) / openprice)
            Else
                percentchange = 0
            End If
            
            Cells(sumrow, 11).Value = percentchange
            Cells(sumrow, 11).NumberFormat = "0.00%"
            
            Cells(sumrow, 12).Value = volume
            sumrow = sumrow + 1
            volume = 0
            openprice = Cells(Srow + 1, 3).Value
            
        End If
        
    Next Srow
        
    
End Sub

