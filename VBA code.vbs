Sub AnalyzeStockData()
    ' Initialize variables for analysis
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim TickerCounter As Long
    Dim StartRow As Long
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestTotalVolume As Double

    For Each ws In Worksheets
        ' Find the last row in column A
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ' Set column headers for analysis results
        With ws.Range("I1:L1")
            .Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
            .Font.Bold = True
        End With
        
        ' Set column headers for summary results
        With ws.Range("P1:Q1")
            .Value = Array("Ticker", "Value")
            .Font.Bold = True
        End With
       
        ' Set summary result titles
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        ' Initialize counters and variables
        TickerCounter = 2
        StartRow = 2
        GreatestIncrease = 0
        GreatestDecrease = 0
        GreatestTotalVolume = 0

        For i = 2 To LastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Extract stock data
                Dim Ticker As String
                Dim YearlyChange As Double
                Dim PercentChange As Double
                Dim TotalVolume As Double
                
                Ticker = ws.Cells(i, 1).Value
                YearlyChange = ws.Cells(i, 6).Value - ws.Cells(StartRow, 3).Value
                PercentChange = IIf(ws.Cells(StartRow, 3).Value <> 0, YearlyChange / ws.Cells(StartRow, 3).Value, 0)
                TotalVolume = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(StartRow, 7), ws.Cells(i, 7)))
                
                ' Update the analysis results in the worksheet
                ws.Cells(TickerCounter, 9).Value = Ticker
                ws.Cells(TickerCounter, 10).Value = YearlyChange
                ws.Cells(TickerCounter, 11).Value = Format(PercentChange, "Percent")
                ws.Cells(TickerCounter, 12).Value = TotalVolume
                ws.Cells(TickerCounter, 10).Interior.ColorIndex = IIf(YearlyChange < 0, 3, 4)
                
                ' Update the greatest values
                If TotalVolume > GreatestTotalVolume Then
                    GreatestTotalVolume = TotalVolume
                    ws.Cells(4, 16).Value = Ticker
                End If
                
                If PercentChange > GreatestIncrease Then
                    GreatestIncrease = PercentChange
                    ws.Cells(2, 16).Value = Ticker
                End If
                
                If PercentChange < GreatestDecrease Then
                    GreatestDecrease = PercentChange
                    ws.Cells(3, 16).Value = Ticker
                End If
                
                ' Move to the next ticker
                TickerCounter = TickerCounter + 1
                StartRow = i + 1
            End If
        Next i

        ' Update the summary results
        With ws
            .Cells(2, 17).Value = Format(GreatestIncrease, "Percent")
            .Cells(3, 17).Value = Format(GreatestDecrease, "Percent")
            .Cells(4, 17).Value = Format(GreatestTotalVolume, "Scientific")
            .Columns("A:Z").AutoFit
        End With
    Next ws
End Sub

