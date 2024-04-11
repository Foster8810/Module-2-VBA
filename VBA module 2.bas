Attribute VB_Name = "Module1"
Sub stocktracker()

    For Each ws In Worksheets
    
        outputrow = 2
        yearlychange = 0
        percentchange = 0
        totalvol = 0
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For row_ = 2 To lastrow
        
            If ws.Cells(row_ + 1, "a").Value <> ws.Cells(row_, "a").Value Then
                
                ticker = ws.Cells(row_, "a").Value
                endprice = ws.Cells(row_, "f").Value
                yearlychange = endprice - startprice
                percentchange = (yearlychange / startprice) * 100
                totalvol = totalvol + ws.Cells(row_, "g").Value
                
                ws.Cells(outputrow, "i").Value = ticker
                ws.Cells(outputrow, "j").Value = yearlychange
                ws.Cells(outputrow, "k").Value = percentchange & "%"
                ws.Cells(outputrow, "l").Value = totalvol
                
                    If yearlychange < 0 Then
                    ws.Cells(outputrow, "j").Interior.ColorIndex = 3
                    
                    Else
                    ws.Cells(outputrow, "j").Interior.ColorIndex = 4
                    
                    End If
                    
                outputrow = outputrow + 1
                totalvol = 0
                starprice = 0
                endprice = 0
                totaldays = 0
                
            Else
                totalvol = totalvol + ws.Cells(row_, "g").Value
                startprice = ws.Cells(row_ - totaldays, "c").Value
                totaldays = totaldays + 1
                
            End If
            
        Next row_
        
'----------
           
        greatestincrease = 0
        greatestdecrease = 0
        greatestvolume = 0
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For row_ = 2 To lastrow
            If greatestincrease < ws.Cells(row_, "k").Value Then
            greatestincrease = ws.Cells(row_, "k").Value
            greatestticker = ws.Cells(row_, "i").Value
            ws.Cells(2, "o").Value = greatestticker
            ws.Cells(2, "p").Value = greatestincrease * 100 & "%"
            
            Else
            ws.Cells(2, "o").Value = greatestticker
            ws.Cells(2, "p").Value = greatestincrease * 100 & "%"
            
            
            End If
            
        Next row_
        
        For row_ = 2 To lastrow
            If greatestdecrease > ws.Cells(row_, "k").Value Then
            greatestdecrease = ws.Cells(row_, "k").Value
            smallestticker = ws.Cells(row_, "i").Value
            ws.Cells(3, "o").Value = smallestticker
            ws.Cells(3, "p").Value = greatestdecrease * 100 & "%"
            
            Else
            ws.Cells(3, "o").Value = smallestticker
            ws.Cells(3, "p").Value = greatestdecrease * 100 & "%"
            
            End If
            
        Next row_
        
            
        For row_ = 2 To lastrow
            If greatestvolume < ws.Cells(row_, "l").Value Then
            greatestvolume = ws.Cells(row_, "l").Value
            largestticker = ws.Cells(row_, "i").Value
            ws.Cells(4, "o").Value = largestticker
            ws.Cells(4, "p").Value = greatestvolume
            
            Else
            ws.Cells(4, "o").Value = largestticker
            ws.Cells(4, "p").Value = greatestvolume
            
            End If
        
        Next row_
        
    Next ws
                
End Sub


