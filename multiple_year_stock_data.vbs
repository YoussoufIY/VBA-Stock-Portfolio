Sub VBAofWallStreet2()

Dim WS_Count As Integer
Dim j As Double
WS_Count = ActiveWorkbook.Worksheets.Count

For j = 1 To WS_Count
    ActiveWorkbook.Worksheets(j).Activate
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Columns(10).AutoFit
    Columns(11).AutoFit
    Columns(12).AutoFit

    Dim ticker As String
    
    Dim No_Of_Rows As Double
    No_Of_Rows = Range("A2").End(xlDown).Row
    Dim openprice As Double
    openprice = Cells(2, 3).Value
    Dim ticker_summary_table As Double
    ticker_summary_table = 2
    Dim closeprice As Double
    Dim change_summary_table As Double
    change_summary_table = 2
    
    Dim TotalStockVolume As Double
    TotalStockVolume = 0
    
    For i = 2 To No_Of_Rows
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ticker = Cells(i, 1).Value
            TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
            Range("I" & ticker_summary_table).Value = ticker
            Range("L" & ticker_summary_table).Value = TotalStockVolume
            ticker_summary_table = ticker_summary_table + 1
            TotalStockVolume = 0
        Else
            TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
            
        End If
        
    Next i
        For i = 2 To No_Of_Rows
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            closeprice = Cells(i, 6).Value
            Range("J" & change_summary_table).Value = closeprice - openprice
            Range("K" & change_summary_table).Value = (closeprice - openprice) / openprice
            change_summary_table = change_summary_table + 1
            Range("K:K").NumberFormat = "0.00%"
            openprice = Cells(i + 1, 3).Value
            End If
        Next i
    
        For i = 2 To No_Of_Rows
        If Cells(i, 10).Value > 0 Then
            Cells(i, 10).Interior.ColorIndex = 4
        ElseIf Cells(i, 10).Value = 0 Then
            Cells(i, 10).Interior.ColorIndex = 0
        Else
            Cells(i, 10).Interior.ColorIndex = 3
            
            End If
         Next i
        Dim mycell As Range
        Dim myrange As Range
        Dim myvolumerange As Range
        Dim myvolume As Range
        
        
        Set myrange = Range("K:K")
        Set myvolumerange = Range("L:L")
        lowestnumber = Application.WorksheetFunction.Min(myrange)
        highestnumber = Application.WorksheetFunction.Max(myrange)
        highestvolume = Application.WorksheetFunction.Max(myvolumerange)
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        Range("O3").Value = "Greatest % Decrease"
        Range("O2").Value = "Greatest % Increase"
        Range("O4").Value = "Greatest Total Volume"
        Columns(15).AutoFit
        Columns(16).AutoFit
        Columns(17).AutoFit
        
        For Each mycell In myrange
            If mycell.Value = lowestnumber Then
                        Range("Q3").Value = lowestnumber
            Range("Q2:Q3").NumberFormat = "0.00%"
            Cells(3, 16).Value = Application.WorksheetFunction.XLookup(Cells(3, 17).Value, Range("K:K"), Range("I:I"))
            
        End If
        
        If mycell.Value = highestnumber Then
                Range("Q2").Value = highestnumber
        Cells(2, 16).Value = Application.WorksheetFunction.XLookup(Cells(2, 17).Value, Range("K:K"), Range("I:I"))
        End If
        Next mycell
        
        For Each myvolume In myvolumerange
        If myvolume.Value = highestvolume Then
                Range("Q4").Value = highestvolume
        Cells(4, 16).Value = Application.WorksheetFunction.XLookup(Cells(4, 17).Value, Range("L:L"), Range("I:I"))
        End If
        Next myvolume
    Next j
End Sub
