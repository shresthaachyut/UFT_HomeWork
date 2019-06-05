Sub getValues()
    getTotalVolByTicker
End Sub

Sub getTotalVolByTicker()
    Dim totalVol, summary_row, LastRow, yearlyChange, percentageChange As Double
    Dim ticker, maxIncTicker, maxDecTicker, gtreatestTotalVolTicker As String
    Dim maxPercentInc, maxPercentDec, greatestTotalVol

    totalVol = 0
    summary_row = 2
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    Dim firstOpen As Double
    firstOpen = Cells(2, 3).Value

    For i = 2 To LastRow
        totalVol = totalVol + Cells(i, 7).Value
        

        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            ticker = Cells(i, 1).Value
            lastClose = Cells(i, 6).Value
            yearlyChange = lastClose - firstOpen
            
            If firstOpen <> 0 Then
               percentageChange = yearlyChange / firstOpen
            Else
                percentageChange = 0
            End If
            Cells(summary_row, 9).Value = ticker
            Cells(summary_row, 10).Value = yearlyChange
            Cells(summary_row, 10).NumberFormat = "0.000000000"
            Cells(summary_row, 11).Value = percentageChange
            Cells(summary_row, 11).NumberFormat = "0.00%"
            Cells(summary_row, 12).Value = totalVol
            
            If percentageChange > maxPercentInc Then
                maxPercentInc = percentageChange
                maxIncTicker = ticker
            End If

            
            If percentageChange < maxPercentDec Then
                maxPercentDec = percentageChange
                maxDecTicker = ticker
            End If
            
            If totalVol > greatestTotalVol Then
                greatestTotalVol = totalVol
                gtreatestTotalVolTicker = ticker
            End If
            

            
            'format the cells
            If (Cells(summary_row, 10).Value < 0) Then
                Cells(summary_row, 10).Interior.ColorIndex = 3
            Else
                Cells(summary_row, 10).Interior.ColorIndex = 4
            End If
    
            
            firstOpen = Cells(i + 1, 3).Value
            totalVol = 0
            summary_row = summary_row + 1
            
        End If
        
    Next i

Cells(2, 15).Value = maxIncTicker
Cells(3, 15).Value = maxDecTicker
Cells(4, 15).Value = gtreatestTotalVolTicker


Cells(2, 16).Value = maxPercentInc
Cells(3, 16).Value = maxPercentDec
Cells(4, 16).Value = greatestTotalVol

Cells(2, 16).NumberFormat = "0.00%"
Cells(3, 16).NumberFormat = "0.00%"
End Sub

Sub Clear_All()

Dim LastRow As String

LastRow = CStr(Cells(Rows.Count, 9).End(xlUp).Row)

Range("I2:I" + LastRow).Clear
Range("J2:J" + LastRow).Clear
Range("K2:K" + LastRow).Clear
Range("L2:L" + LastRow).Clear
Range("O2:O" + LastRow).Clear
Range("P2:P" + LastRow).Clear
End Sub