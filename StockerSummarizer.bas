Attribute VB_Name = "StockerSummarizer"
Option Explicit

' <Chan Feng> 2018-02-08 Ready to check in

Sub SummarizeAllSheets()
    ' Run through all sheets
    Dim current As Worksheet
    
    For Each current In ActiveWorkbook.Worksheets
        current.Activate
        Application.StatusBar = "Processing " & current.Name
        Call SummerizeTicker
    Next
End Sub

Sub SummerizeTicker()
    Dim ticker As String
    Dim prevTicker As String
    
    Dim volumne As Long
    Dim totalVolume As Double
    
    Dim row As Long
    Dim summaryRow As Long
    
    Dim openPrice As Currency
    Dim closePrice As Currency
    Dim yearlyChange As Double
    Dim yearlyPercChange As Double
    
    Dim greatestPercIncreaseTicker As String
    Dim greatestPercIncrease As Double
    
    Dim greatestPercDecreaseTicker As String
    Dim greatestPercDecrease As Double
    
    Dim greatestTotalVolumeTicker As String
    Dim greatestTotalVolume As Double
                
    ' Set header
    Range("I1") = "Ticker"
    Range("J1") = "Yearly Change"
    Range("K1") = "Percent Change"
    Range("L1") = "Total Stock Volumne"
    Range("O2") = "Greatest % Increase"
    Range("O3") = "Greatest % Decrease"
    Range("O4") = "Greatest Total Volume"
    Range("P1") = "Ticker"
    Range("Q1") = "Values"
    summaryRow = 2
    
    ' Init values
    row = 2
    ticker = Range("A" & row)
    openPrice = Range("C" & row)
    totalVolume = 0
    greatestPercIncrease = 0#
    greatestPercDecrease = 0#
    greatestTotalVolume = 0
    
    While ticker <> ""
        prevTicker = ticker
       
        closePrice = Range("F" & row)
        totalVolume = totalVolume + Range("G" & row)

        row = row + 1
        ticker = Range("A" & row)

        If ticker <> prevTicker Then 'Passed all data for one ticker, now summarize
            Application.StatusBar = "Done with ticker: " + prevTicker
            yearlyChange = closePrice - openPrice
            Range("I" & summaryRow) = prevTicker
            Range("J" & summaryRow) = yearlyChange
            Range("J" & summaryRow).NumberFormat = "#.#0"
    
            ' Optimized to use conditional format
            With Range("J" & summaryRow).FormatConditions
                .Delete
                .Add(xlCellValue, xlGreater, "0").Interior.Color = RGB(0, 255, 0)
                .Add(xlCellValue, xlLessEqual, "0").Interior.Color = RGB(255, 0, 0)
            End With
            
            ' Handle exceptions
            If openPrice = 0 Then
                yearlyPercChange = 0
            Else
                yearlyPercChange = yearlyChange / openPrice
            End If
    
            Range("K" & summaryRow) = yearlyPercChange
            Range("K" & summaryRow).NumberFormat = "#.#0%"
            Range("L" & summaryRow) = totalVolume
            Range("L" & summaryRow).NumberFormat = "#,###"
            
            ' Set the greatest tickers
            If yearlyPercChange > greatestPercIncrease Then
                greatestPercIncrease = yearlyPercChange
                greatestPercIncreaseTicker = prevTicker
            End If
            
            If yearlyPercChange < greatestPercDecrease Then
                greatestPercDecrease = yearlyPercChange
                greatestPercDecreaseTicker = prevTicker
            End If
            
            If totalVolume > greatestTotalVolume Then
                greatestTotalVolume = totalVolume
                greatestTotalVolumeTicker = prevTicker
            End If
            
            openPrice = Range("C" & row)
            summaryRow = summaryRow + 1
            totalVolume = 0
         End If
    Wend
    
    Range("P2") = greatestPercIncreaseTicker
    Range("Q2") = greatestPercIncrease
    Range("Q2").NumberFormat = "#.#0%"
    Range("P3") = greatestPercDecreaseTicker
    Range("Q3").NumberFormat = "#.#0%"
    Range("Q3") = greatestPercDecrease
    Range("P4") = greatestTotalVolumeTicker
    Range("Q4") = greatestTotalVolume
    Range("Q4").NumberFormat = "#,###"
    
    Application.StatusBar = "Done with Sheet " & ActiveSheet.Name
    
End Sub
