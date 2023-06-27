Sub Alphabetical_Test()
'Create Variable for the Worksheet
    Dim ws As Worksheet
    Dim tickerSymbol As String
    Dim priceChange As Double
    Dim percentageChange As Double
    Dim totalStockVolume As Double
    Dim RowNumber As Double
    Dim yearlyChange As Double
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestVolume As Double
    
'Loop through the enitre Worksheet
    For Each ws In Worksheets
    ws.Activate
'Count the number of Rows
        Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
'Create Header for all the addtional Columns
        Range("I1").Value = "Ticker"
        Range("J1").Value = "yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        Range("N2").Value = "Greatest % Increase"
        Range("N3").Value = "Greatest % Decrease"
        RowNumber = 2
        yearlyChange = 0
        Opening = Cells(2, 3).Value
        GreatestIncrease = 0
        GreatestVolume = 0
        GreatestDecrease = 999999999
'Create For loop to Lastrow
    For i = 2 To Lastrow

        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ticker = Cells(i, 1).Value
            Cells(RowNumber, "I").Value = ticker
            Cells(RowNumber, "L").Value = totalStockVolume
            Cells(RowNumber, "K").Value = percentChange
            Closing = Cells(i, "F").Value
            yearlyChange = Closing - Opening
            Cells(RowNumber, "J").Value = yearlyChange
    
'percentChange
       percentageChange = yearlyChange / Closing
       Cells(RowNumber, "K").Value = percentageChange
       
'if statement to see if the current percenatge change is larger  than the largest percentage change
    If percentageChange > GreatestIncrease Then
        GreatestIncrease = percentageChange
        Range("P2").Value = GreatestIncrease
    End If
'if statement to see if the current percenatge change is less than the smallest percentage change
    If percentageChange < GreatestDecrease Then
        GreatestDecrease = percentageChange
        Range("P3").Value = GreatestDecrease
    End If
 'if statement for the total stock  to see if the current total stockvolume is larger than the largest stockvolume
    If totalStockVolume > GreatestIncrease Then
        GreatestIncrease = totalStockVolume
        Range("P4").Value = GreatestIncrease
    End If
  'Move the totalStockVolumn further down to aviod "0"
        totalStockVolume = 0
       Cells(RowNumber, "K").NumberFormat = "0.00%"
            
'Opening Vaule for the next Ticker
            Opening = Cells(i + 1, 3).Value
            RowNumber = RowNumber + 1

        Else
            totalStockVolume = Cells(i, 7).Value + totalStockVolume
         
        End If
    Next i
Next ws
    
End Sub