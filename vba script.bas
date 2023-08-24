Attribute VB_Name = "Module1"
Sub CalculateStockData()
    Dim LastRow As Long
    Dim Ticker As String
    Dim YearlyOpenPrice As Double
    Dim YearlyClosePrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    Dim SummaryRow As Long
    
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    SummaryRow = 2
    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Volume"
    
    For i = 2 To LastRow
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Ticker = Cells(i, 1).Value
            YearlyOpenPrice = Cells(i, 3).Value
            YearlyClosePrice = Cells(i, 6).Value
            YearlyChange = YearlyClosePrice - YearlyOpenPrice
            If YearlyOpenPrice <> 0 Then
                PercentChange = (YearlyChange / YearlyOpenPrice) * 100
            Else
                PercentChange = 0
            End If
            TotalVolume = Application.WorksheetFunction.Sum(Range(Cells(SummaryRow, 7), Cells(i, 7)))
            
            Range("I" & SummaryRow).Value = Ticker
            Range("J" & SummaryRow).Value = YearlyChange
            Range("K" & SummaryRow).Value = PercentChange
            Range("L" & SummaryRow).Value = TotalVolume
            
            SummaryRow = SummaryRow + 1
        End If
    Next i
    
    Range("K2:K" & SummaryRow - 1).NumberFormat = "0.00%"
    
    
    Dim MaxIncrease As Double
    Dim MaxDecrease As Double
    Dim MaxVolume As Double
    Dim MaxIncreaseTicker As String
    Dim MaxDecreaseTicker As String
    Dim MaxVolumeTicker As String
    
    MaxIncrease = WorksheetFunction.Max(Range("D2:D" & SummaryRow - 1))
    MaxDecrease = WorksheetFunction.Min(Range("E2:E" & SummaryRow - 1))
    MaxVolume = WorksheetFunction.Max(Range("G2:G" & SummaryRow - 1))
    
    MaxIncreaseTicker = Cells(Application.WorksheetFunction.Match(MaxIncrease, Range("D2:d" & SummaryRow - 1), 0) + 1, 8).Value
    MaxDecreaseTicker = Cells(Application.WorksheetFunction.Match(MaxDecrease, Range("E2:E" & SummaryRow - 1), 0) + 1, 8).Value
    MaxVolumeTicker = Cells(Application.WorksheetFunction.Match(MaxVolume, Range("G2:G" & SummaryRow - 1), 0) + 1, 8).Value
    
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    
    Range("P2").Value = MaxIncreaseTicker
    Range("P3").Value = MaxDecreaseTicker
    Range("P4").Value = MaxVolumeTicker
    Range("Q2").Value = MaxIncrease
    Range("Q3").Value = MaxDecrease
    Range("Q4").Value = MaxVolume
    
    Range("Q2:Q3").NumberFormat = "0.00%"
    
End Sub

