Attribute VB_Name = "Module1"
Sub Stock():

For Each ws In Worksheets

    Dim Ticker As String
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalStockVolume As Double
    TotalStockVolume = 0
    Dim AnalysisRow As Integer
    AnalysisRow = 2
    Dim FirstRow As Long
    FirstRow = 2
    Dim LastRow As Long
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        
        For i = 2 To LastRow
                
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        Ticker = ws.Cells(i, 1).Value
        TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
            If ws.Cells(FirstRow, 6) = 0 Then
            YearlyChange = 0
            PercentChange = 0
            Else
            YearlyChange = ws.Cells(i, 6).Value - ws.Cells(FirstRow, 3)
            PercentChange = YearlyChange / ws.Cells(FirstRow, 3)
            End If
        
        ws.Range("I" & AnalysisRow).Value = Ticker
        ws.Range("J" & AnalysisRow).Value = YearlyChange
        ws.Range("K" & AnalysisRow).Value = PercentChange
        ws.Range("L" & AnalysisRow).Value = TotalStockVolume
        
        ws.Cells(AnalysisRow, 11).NumberFormat = "0.00%"
        
            If YearlyChange < 0 Then
            ws.Range("J" & AnalysisRow).Interior.ColorIndex = 3
            Else
            ws.Range("J" & AnalysisRow).Interior.ColorIndex = 4
            End If
            
        AnalysisRow = AnalysisRow + 1
        TotalStockVolume = 0
        FirstRow = ws.Cells(i + 1, 1).Row
        
        Else

        TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
       
        
        End If
        
        Next i
           
    Dim MaxIncrease As Double
    Dim MinIncrease As Double
    Dim MaxTotalVolume As Double
          
    ws.Range("O2").Value = "Greatest % increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest total volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    MaxIncrease = Application.WorksheetFunction.Max(ws.Range("K:K"))
    MinIncrease = Application.WorksheetFunction.Min(ws.Range("K:K"))
    MaxTotalVolume = Application.WorksheetFunction.Max(ws.Range("L:L"))
        
    ws.Range("Q2").Value = MaxIncrease
    ws.Range("Q3").Value = MinIncrease
    ws.Range("Q4").Value = MaxTotalVolume
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
        
    Dim LastRow2 As Long
        
    LastRow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        For i = 2 To LastRow2

        If ws.Cells(i, 11).Value = MaxIncrease Then
        ws.Range("P2").Value = ws.Cells(i, 9).Value
        
        ElseIf ws.Cells(i, 11).Value = MinIncrease Then
        ws.Range("P3").Value = ws.Cells(i, 9).Value
        
        ElseIf ws.Cells(i, 12).Value = MaxTotalVolume Then
        ws.Range("P4").Value = ws.Cells(i, 9).Value
        
        End If
        
        Next i
                
Next ws


End Sub
