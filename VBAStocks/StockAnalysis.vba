VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub StockAnalysis():
    
    ' Loop through all the worksheets.
    For Each ws In Worksheets
    
    ' Print column headers to cells.
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    
    ' Declare and initialize variables.
    Dim Ticker As String
    Dim StockVolume As Double
    StockVolume = 0
    
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    Dim MaxIncrease, MinDecrease, MaxVolume As Double
    MaxIncrease = 0
    MaxDecrease = 0
    MaxVolume = 0
    
    ' Get the location of the last row.
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Get the opening value of the first stock.
    OpeningValue = ws.Cells(2, 3).Value
        
    
    ' Loop through the rows.
    For i = 2 To LastRow
        
        
        ' Take these actions at the transition rows from one stock to the next.
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            ' Record the ticker of the current stock and write it to the summary table.
            Ticker = ws.Cells(i, 1).Value
            ws.Range("I" & Summary_Table_Row).Value = Ticker
            
            ' Get the closing value of the current stock.
            ClosingValue = ws.Cells(i, 6).Value
            
            ' Calculate the yearly change, write it to the summary table, and format the cell.
            YearlyChange = ClosingValue - OpeningValue
            ws.Range("J" & Summary_Table_Row).Value = YearlyChange
            If YearlyChange > 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            Else
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            End If
                        
            ' Calculate percent change, write it to the summary table, and format the cell.
            If OpeningValue <> 0 Then
                PercentChange = (ClosingValue - OpeningValue) / OpeningValue
                ws.Range("K" & Summary_Table_Row).Value = PercentChange
                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                
            Else
                ws.Range("K" & Summary_Table_Row).Value = "Error"
            End If
            
            ' Get the new opening value of the next stock.
            OpeningValue = ws.Cells(i + 1, 3).Value
            
            ' Update the total stock volume and write it to the summary table.
            StockVolume = StockVolume + ws.Cells(i, 7).Value
            ws.Range("L" & Summary_Table_Row).Value = StockVolume
            
            ' Update Max Increase/Decrease if applicable
                If PercentChange > MaxIncrease Then
                    MaxIncrease = PercentChange
                    MaxIncreaseTicker = Ticker
                End If
                
                If PercentChange < MaxDecrease Then
                    MaxDecrease = PercentChange
                    MaxDecreaseTicker = Ticker
                End If
            
            ' Update the Max Total Volume if applicable
            If StockVolume > MaxVolume Then
                MaxVolume = StockVolume
                MaxVolumeTicker = Ticker
            End If
            
            ' Increment the summary table row and reset the stock volume.
            Summary_Table_Row = Summary_Table_Row + 1
            StockVolume = 0
            
        Else
            
            ' Update the total stock volume.
            StockVolume = StockVolume + ws.Cells(i, 7).Value
            
        
        End If
    
    
    
    Next i
    
    ' Populate Max Increase/Decrease/Volume Table, and format.
    ws.Range("P2").Value = MaxIncreaseTicker
    ws.Range("P3").Value = MaxDecreaseTicker
    ws.Range("P4").Value = MaxVolumeTicker
    
    ws.Range("Q2").Value = MaxIncrease
    ws.Range("Q3").Value = MaxDecrease
    ws.Range("Q4").Value = MaxVolume
    
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    
    ' Autofit the columns
    ws.Range("A1:Q4").Columns.AutoFit
    
    Next ws
    
End Sub



