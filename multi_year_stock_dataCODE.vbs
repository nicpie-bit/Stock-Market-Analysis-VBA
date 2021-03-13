VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub MultiYearStockData():
    Dim CurrentWs As Worksheet
    For Each CurrentWs In Worksheets
    
        Dim YearlyChange As Long
        Dim PercentChange As Long
        Dim TotalVolumn As Double
        Dim LastRow As Long
        Dim OpeningPrice As Long
        Dim ClosingPrice As Long
        Dim SummaryRow As Long
        
        Dim MaxTicker As String
        Dim MinTicker As String
        Dim MaxTickerVol As String
        Dim MaxPercent As Long
        Dim MinPercent As Long
        Dim MaxVol As Double
        
        'Set initial Values
        CurrentWs.Range("I1").Value = "Ticker"
        CurrentWs.Range("J1").Value = "Yearly Change"
        CurrentWs.Range("K1").Value = "Percent Change"
        CurrentWs.Range("L1").Value = "Total Stock Volumn"
        
        CurrentWs.Range("O1").Value = "Ticker"
        CurrentWs.Range("P1").Value = "Value"
        CurrentWs.Range("N2").Value = "Greatest % Increase"
        CurrentWs.Range("N3").Value = "Greatest % Decrease"
        CurrentWs.Range("N4").Value = "Greatest Total Volumn"
        
        'Set initial values
        TotalVolumn = 0
        SummaryRow = 1
        MaxPercent = 0
        MinPercent = 0
        MaxVol = 0
        
        'Last Row
        LastRow = CurrentWs.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Loop through tickers
        Dim i As Long
        
        
        For i = 2 To LastRow
            'If next value /= to current value
            If CurrentWs.Cells(i - 1, 1).Value <> CurrentWs.Cells(i, 1).Value Then
                SummaryRow = SummaryRow + 1
                CurrentWs.Cells(SummaryRow, 9).Value = CurrentWs.Cells(i, 1).Value
                
                OpeningPrice = CurrentWs.Cells(i, 3).Value
                
                TotalVolumn = 0
            'If next row is same ticker
            ElseIf CurrentWs.Cells(i, 1).Value <> CurrentWs.Cells(i + 1, 1).Value Then
                ClosingPrice = CurrentWs.Cells(i, 6).Value
                YearlyChange = ClosingPrice - OpeningPrice
                
                If OpeningPrice = 0 Then
                    PercentChange = NA
                Else
                    PercentChange = Round(((YearlyChange / OpeningPrice) * 100), 2)
                End If
                
                CurrentWs.Cells(SummaryRow, 10).Value = YearlyChange
                CurrentWs.Cells(SummaryRow, 11).Value = (CStr(PercentChange) & "%")
                
                If (YearlyChange > 0) Then
                    CurrentWs.Range("J" & SummaryRow).Interior.ColorIndex = 4
                ElseIf (YearlyChange <= 0) Then
                    CurrentWs.Range("J" & SummaryRow).Interior.ColorIndex = 3
                End If
             
             'Bonus
                If PercentChange > MaxPercent Then
                    MaxPercent = PercentChange
                    MaxTicker = CurrentWs.Cells(i, 1).Value
                ElseIf PercentChange < MinPercent Then
                    MinPercent = PercentChange
                    MinTicker = CurrentWs.Cells(i, 1).Value
                End If
                
                If TotalVolumn > MaxVol Then
                    MaxVol = TotalVolumn
                    MaxTickerVol = CurrentWs.Cells(i, 1).Value
                End If
                
                CurrentWs.Range("O2").Value = MaxTicker
                CurrentWs.Range("O3").Value = MinTicker
                CurrentWs.Range("O4").Value = MaxTickerVol
                CurrentWs.Range("P2").Value = (CStr(MaxPercent) & "%")
                CurrentWs.Range("P3").Value = (CStr(MinPercent) & "%")
                CurrentWs.Range("P4").Value = MaxVol
            
            End If
            
            If OpeningPrice = 0 And CurrentWs.Cells(i, 3).Value <> 0 Then
                OpeningPrice = CurrentWs.Cells(i, 3).Value
            End If
            
            TotalVolumn = TotalVolumn + CurrentWs.Cells(i, 7).Value
            CurrentWs.Cells(SummaryRow, 12).Value = TotalVolumn
            
            
        Next i
    Next CurrentWs
                
            
    
End Sub

