Attribute VB_Name = "Module1"
Sub StockMarket()
    Dim currentSheet As Worksheet
    Dim sheetNo As Integer
    For sheetNo = 1 To ThisWorkbook.Worksheets.Count
        Set currentSheet = ThisWorkbook.Worksheets(sheetNo)
        
        Dim rowRange As Range
        Dim LastRow As Long
        LastRow = currentSheet.Cells(currentSheet.Rows.Count, "A").End(xlUp).Row
    
        Set rowRange = currentSheet.Range("A3:A" & LastRow)
        
        Dim tickerSymbol As String
        tickerSymbol = currentSheet.Cells(2, 1).Text
        
        Dim openPrice As Double
        openPrice = currentSheet.Cells(2, 3).Value
        
        Dim closePrice As Double
        closePrice = currentSheet.Cells(2, 6).Value
        
        Dim yearlyChange As Double
        Dim percentageChange As Double
        
        Dim dayVol As Double
        dayVol = currentSheet.Cells(2, 7).Value
        
        currentSheet.Cells(1, 9).Value = "Ticker"
        currentSheet.Cells(1, 10).Value = "Yearly Change"
        currentSheet.Cells(1, 11).Value = "Percentage Change"
        currentSheet.Cells(1, 12).Value = "Total Stock Value"
        
        currentSheet.Cells(2, 15).Value = "Greatest % Increase"
        currentSheet.Cells(3, 15).Value = "Greatest % Decrease"
        currentSheet.Cells(4, 15).Value = "Greatest Total Volume"
        
        currentSheet.Cells(1, 16).Value = "Ticker"
        currentSheet.Cells(1, 17).Value = "Value"
        
        Dim outputRow As Long
        outputRow = 2
        
        Dim greatestIncreaseTicker As String
        Dim greatestIncrease As Double
        greatestIncrease = 0
        
        Dim greatestDecreaseTicker As String
        Dim greatestDecrease As Double
        greatestDecrease = 0
        
        Dim greatestVolumeTicker As String
        Dim greatestVolume As Double
        greatestVolume = 0
        
        For Each currentRow In rowRange
            If StrComp(currentSheet.Cells(currentRow.Row, 1).Text, tickerSymbol, vbTextCompare) = 0 Then
                closePrice = currentSheet.Cells(currentRow.Row, 6).Value
                dayVol = dayVol + currentSheet.Cells(currentRow.Row, 7).Value
            Else
                yearlyChange = closePrice - openPrice
                percentageChange = yearlyChange / openPrice
                
                currentSheet.Cells(outputRow, 9).Value = tickerSymbol
                currentSheet.Cells(outputRow, 10).Value = yearlyChange
                currentSheet.Cells(outputRow, 11).Value = percentageChange
                currentSheet.Cells(outputRow, 12).Value = dayVol
                
                If yearlyChange < 0 Then
                    currentSheet.Cells(outputRow, 10).Interior.ColorIndex = 3
                Else
                    currentSheet.Cells(outputRow, 10).Interior.ColorIndex = 4
                End If
                
                If percentageChange < greatestDecrease Then
                    greatestDecrease = percentageChange
                    greatestDecreaseTicker = tickerSymbol
                End If
                 
                If percentageChange > greatestIncrease Then
                    greatestIncrease = percentageChange
                    greatestIncreaseTicker = tickerSymbol
                End If
                
                If dayVol > greatestVolume Then
                    greatestVolume = dayVol
                    greatestVolumeTicker = tickerSymbol
                End If
                
                ' Setup for new symbol
                tickerSymbol = currentSheet.Cells(currentRow.Row, 1).Text
                openPrice = currentSheet.Cells(currentRow.Row, 3).Value
                closePrice = currentSheet.Cells(currentRow.Row, 6).Value
                dayVol = currentSheet.Cells(currentRow.Row, 7).Value
                outputRow = outputRow + 1
            End If
        
        Next currentRow
        yearlyChange = closePrice - openPrice
        percentageChange = yearlyChange / openPrice
        
        currentSheet.Cells(outputRow, 9).Value = tickerSymbol
        currentSheet.Cells(outputRow, 10).Value = yearlyChange
        currentSheet.Cells(outputRow, 11).Value = percentageChange
        currentSheet.Cells(outputRow, 12).Value = dayVol
        
        If yearlyChange < 0 Then
            currentSheet.Cells(outputRow, 10).Interior.ColorIndex = 3
        Else
            currentSheet.Cells(outputRow, 10).Interior.ColorIndex = 4
        End If
        
        If percentageChange < greatestDecrease Then
            greatestDecrease = percentageChange
            greatestDecrease = tickerSymbol
        End If
         
        If percentageChange > greatestIncrease Then
            greatestIncrease = percentageChange
            greatestIncreaseTicker = tickerSymbol
        End If
        
        If dayVol > greatestVolume Then
            greatestVolume = dayVol
            greatestVolumeTicker = tickerSymbol
        End If
        
        currentSheet.Cells(2, 16).Value = greatestIncreaseTicker
        currentSheet.Cells(2, 17).Value = greatestIncrease
        
        currentSheet.Cells(3, 16).Value = greatestDecreaseTicker
        currentSheet.Cells(3, 17).Value = greatestDecrease
        
        currentSheet.Cells(4, 16).Value = greatestVolumeTicker
        currentSheet.Cells(4, 17).Value = greatestVolume
    
    Next sheetNo
End Sub

