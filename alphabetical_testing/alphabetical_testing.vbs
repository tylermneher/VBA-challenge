Sub AlphabeticalTesting():

    Dim ws As Worksheet
    Dim TickerSymbol As String
    Dim YearOpen As Double
    Dim YearClose As Double
    Dim PercentChange As Double
    
    Cells.Range("I1").Value = "Ticker"
    Cells.Range("J1").Value = "Total Volume for Year"
    Cells.Range("K1").Value = "Percent Change from Year Open to Year Close"
    
    Dim TotalVolume As Double
    TotalVolume = 0
    
    Dim TrackerRow As Integer
    TrackerRow = 2

    For Each ws In ActiveWorkbook.Worksheets
        For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        
        Cells(i, 3).Value = YearOpen
        
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

                TickerSymbol = Cells(i, 1).Value
                
                YearClose = Cells(i, 6).Value
                PercentChange = (YearClose - YearOpen) / YearOpen
                
                Range("I" & TrackerRow).Value = TickerSymbol
                Range("J" & TrackerRow).Value = TotalVolume
                Range("K" & TrackerRow).Value = PercentChange
                
                TrackerRow = TrackerRow + 1
                TotalVolume = 0
                PercentChange = 0

            Else
            
                TotalVolume = TotalVolume + Cells(i, 7)
                
            End If

        Next i
    Next ws
End Sub

