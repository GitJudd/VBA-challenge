Attribute VB_Name = "Module1"
'Sub clearSample()
'
' Testing Environment Use - Remove for main project
'
'  sheets("A").Columns("I:R").EntireColumn.Delete
'  sheets("B").Columns("I:R").EntireColumn.Delete
'  sheets("C").Columns("I:R").EntireColumn.Delete
'  sheets("D").Columns("I:R").EntireColumn.Delete
'  sheets("E").Columns("I:R").EntireColumn.Delete
'  sheets("F").Columns("I:R").EntireColumn.Delete
'  sheets("P").Columns("I:R").EntireColumn.Delete
'
'End Sub

Sub clear()

  sheets("2014").Columns("I:R").EntireColumn.Delete
  sheets("2015").Columns("I:R").EntireColumn.Delete
  sheets("2016").Columns("I:R").EntireColumn.Delete

End Sub

Sub multisheet()
    Dim sheets As Worksheet
        Application.ScreenUpdating = False
            For Each sheets In Worksheets
            sheets.Select
            Call pleasework
        Next
        Application.ScreenUpdating = True
End Sub

Sub pleasework()

    Dim findlast As Long
    Dim ticker As String
    Dim yr_change As Double
    Dim yr_percent As Double
    Dim ws As Worksheet
    Dim Summary_Table_Row As Integer
    Dim yr_open As Double
    Dim yr_close As Double
    Dim vol As Double

    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Yearly Percentage"
    Cells(1, 12).Value = "Total Stock Vol"
    
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Columns("J:L").ColumnWidth = 15
    Columns("M:N").ColumnWidth = 5
    Columns("O").ColumnWidth = 20
    
    Summary_Table_Row = 2
    
    findlast = ActiveSheet.UsedRange.Rows(ActiveSheet.UsedRange.Rows.Count).Row
    
    For i = 2 To findlast

        If yr_open = 0 Then
            yr_open = Cells(i, 3).Value
        End If
  
        vol = vol + Cells(i, 7).Value

        If Cells(i - 1, 1) = Cells(i, 1) And Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ticker = Cells(i, 1).Value
            yr_close = Cells(i, 6).Value
            yr_change = yr_close - yr_open
        
                If yr_open = 0 Then
                    yr_percent = 0
                Else
                    yr_percent = yr_change / yr_open
                End If
            
            Range("I" & Summary_Table_Row).Value = ticker
            Range("J" & Summary_Table_Row).Value = yr_change
            Range("K2:K" & findlast).NumberFormat = "0.00%"
            Range("K" & Summary_Table_Row).Value = yr_percent
            Range("L" & Summary_Table_Row).Value = vol
            Summary_Table_Row = Summary_Table_Row + 1

            vol = 0
            yr_open = 0
      
        End If
    
        
    Next i
        gMax = Application.WorksheetFunction.Max(Columns("K"))
        gMaxTickerRow = Application.WorksheetFunction.Match(gMax, Columns("K"), 0)
        Cells(2, 16).Value = Cells(gMaxTickerRow, 9).Value
        
        gMin = Application.WorksheetFunction.Min(Columns("K"))
        gMinTickerRow = Application.WorksheetFunction.Match(gMin, Columns("K"), 0)
        Cells(3, 16).Value = Cells(gMinTickerRow, 9).Value
            
        vMax = Application.WorksheetFunction.Max(Columns("L"))
        vMaxTickerRow = Application.WorksheetFunction.Match(vMax, Columns("L"), 0)
        Cells(4, 16).Value = Cells(vMaxTickerRow, 9).Value
        
        Cells(2, 17).Value = gMax
        Cells(2, 17).NumberFormat = "0.00%"
        Cells(3, 17).Value = gMin
        Cells(3, 17).NumberFormat = "0.00%"
        Cells(4, 17).Value = vMax
        
    'Conditional Formatting for J Column
    Dim rng As Range
    Dim condition1 As FormatCondition
    Dim condition2 As FormatCondition

    Set rng = Range("J2", "J9000")
    rng.FormatConditions.Delete
    Set condition1 = rng.FormatConditions.Add(xlCellValue, xlGreater, "=0")
    Set condition2 = rng.FormatConditions.Add(xlCellValue, xlLess, "=0")
    With condition1
    .Interior.Color = vbGreen
    End With
    With condition2
     .Interior.Color = vbRed
   End With


End Sub

