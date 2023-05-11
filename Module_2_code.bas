Attribute VB_Name = "Module1"
Sub stockAnalysis()
    
    Dim lastRow As Long
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Declare variables
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    
    ' Declare variables for tracking greatest values
    Dim greatestPercentIncrease As Double
    Dim greatestPercentDecrease As Double
    Dim greatestTotalVolume As Double
    Dim greatestPercentIncreaseTicker As String
    Dim greatestPercentDecreaseTicker As String
    Dim greatestTotalVolumeTicker As String
    
    ' Initialize tracking variables
    greatestPercentIncrease = 0
    greatestPercentDecrease = 0
    greatestTotalVolume = 0
    
    ' Loop through all rows
    For i = 2 To lastRow
        
        ' Check if we're still looking at the same ticker
        If Cells(i, 1).Value <> ticker Then
            
            ' Set the ticker
            ticker = Cells(i, 1).Value
            
            ' Set the opening price
            openingPrice = Cells(i, 3).Value
            
            ' Reset the total volume counter
            totalVolume = 0
            
        End If
        
        ' Add to the total volume
        totalVolume = totalVolume + Cells(i, 7).Value
        
        ' Check if we're at the end of the current ticker
        If Cells(i + 1, 1).Value <> ticker Then
            
            ' Set the closing price
            closingPrice = Cells(i, 6).Value
            
            ' Calculate the yearly change and percent change
            yearlyChange = closingPrice - openingPrice
            percentChange = yearlyChange / openingPrice
            
            ' Output the results
            Cells(1, 9).Value = "Ticker"
            Cells(1, 10).Value = "Yearly Change"
            Cells(1, 11).Value = "Percent Change"
            Cells(1, 12).Value = "Total Volume"
            
            Cells(Rows.Count, 9).End(xlUp).Offset(1, 0).Value = ticker
            Cells(Rows.Count, 10).End(xlUp).Offset(1, 0).Value = yearlyChange
            Cells(Rows.Count, 11).End(xlUp).Offset(1, 0).Value = percentChange
            Cells(Rows.Count, 12).End(xlUp).Offset(1, 0).Value = totalVolume
            
            ' Check for new greatest values
            If percentChange > greatestPercentIncrease Then
                greatestPercentIncrease = percentChange
                greatestPercentIncreaseTicker = ticker
            ElseIf percentChange < greatestPercentDecrease Then
                greatestPercentDecrease = percentChange
                greatestPercentDecreaseTicker = ticker
            End If
            
            If totalVolume > greatestTotalVolume Then
                greatestTotalVolume = totalVolume
                greatestTotalVolumeTicker = ticker
            End If
            
            ' Reset variables for next ticker
            openingPrice = 0
            closingPrice = 0
            yearlyChange = 0
            percentChange = 0
            totalVolume = 0
            
        End If
        
    Next i
    
    ' Output the greatest values
    Cells(2, 16).Value = "Greatest % Increase"
    Cells(3, 16).Value = "Greatest % Decrease"
    Cells(4, 16).Value = "Greatest Total Volume"
    
    End Sub
    

