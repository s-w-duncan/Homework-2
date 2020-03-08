Attribute VB_Name = "TestLoop"
Sub StockAnalysis()
    On Error Resume Next

    Dim Worksheet As Integer
    Dim Worksheet_Count As Integer
    Dim r As Long
    Dim a As Integer
    Dim b As Integer
    Dim c As Integer
    Dim d As Integer
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Total_Volume As Double
    
    b = 2
    Worksheet_Count = ActiveWorkbook.Worksheets.Count

    For Worksheet = 1 To Worksheet_Count
        ActiveWorkbook.Worksheets(Worksheet).Cells(1, 9).Value = "Ticker"
        ActiveWorkbook.Worksheets(Worksheet).Cells(1, 10).Value = "Total Stock Volume"
        ActiveWorkbook.Worksheets(Worksheet).Cells(1, 11).Value = "Yearly Change"
        ActiveWorkbook.Worksheets(Worksheet).Cells(1, 12).Value = "Percent Change"
        ActiveWorkbook.Worksheets(Worksheet).Cells(1, 16).Value = "Ticker"
        ActiveWorkbook.Worksheets(Worksheet).Cells(1, 17).Value = "Value"
        ActiveWorkbook.Worksheets(Worksheet).Cells(2, 15).Value = "Greatest % Increase"
        ActiveWorkbook.Worksheets(Worksheet).Cells(3, 15).Value = "Greatest % Decrease"
        ActiveWorkbook.Worksheets(Worksheet).Cells(4, 15).Value = "Greatest Total Volume"
        
        Total_Volume = 0
        
        For r = 2 To ActiveWorkbook.Worksheets(Worksheet).Cells.SpecialCells(xlCellTypeLastCell).Row
            If ActiveWorkbook.Worksheets(Worksheet).Cells(r, 1).Value <> ActiveWorkbook.Worksheets(Worksheet).Cells(r + 1, 1).Value Then
                
                Close_Price = ActiveWorkbook.Worksheets(Worksheet).Cells(r, 6).Value
                
                ActiveWorkbook.Worksheets(Worksheet).Cells(b, 11).Value = Close_Price - Open_Price
                ActiveWorkbook.Worksheets(Worksheet).Cells(b, 12).Value = (Close_Price - Open_Price) / Open_Price
                
                Open_Price = 0
                Close_Price = 0
                
                ActiveWorkbook.Worksheets(Worksheet).Cells(b, 9).Value = ActiveWorkbook.Worksheets(Worksheet).Cells(r, 1).Value
                
                Total_Volume = Total_Volume + ActiveWorkbook.Worksheets(Worksheet).Cells(r, 7).Value
                ActiveWorkbook.Worksheets(Worksheet).Cells(b, 10).Value = Total_Volume
                
                b = b + 1
                Total_Volume = 0
                
            ElseIf ActiveWorkbook.Worksheets(Worksheet).Cells(r - 1, 1).Value <> ActiveWorkbook.Worksheets(Worksheet).Cells(r, 1).Value Then
                
                Open_Price = ActiveWorkbook.Worksheets(Worksheet).Cells(r, 3).Value
                Total_Volume = Total_Volume + ActiveWorkbook.Worksheets(Worksheet).Cells(r, 7).Value
                
            Else
                Total_Volume = Total_Volume + ActiveWorkbook.Worksheets(Worksheet).Cells(r, 7).Value
            
            End If
        Next r

        For c = 2 To ActiveWorkbook.Worksheets(Worksheet).Range("K1").CurrentRegion.Rows.Count
            ActiveWorkbook.Worksheets(Worksheet).Cells(c, 12).Style = "Percent"
            
            If ActiveWorkbook.Worksheets(Worksheet).Cells(c, 11).Value > 0 Then
                ActiveWorkbook.Worksheets(Worksheet).Cells(c, 11).Interior.ColorIndex = 4
                
            Else
                ActiveWorkbook.Worksheets(Worksheet).Cells(c, 11).Interior.ColorIndex = 3
                    
            End If
        Next c
        
        ActiveWorkbook.Worksheets(Worksheet).Cells(2, 17).Value = WorksheetFunction.Max(Worksheets(Worksheet).Range("L2:L" & ActiveWorkbook.Worksheets(Worksheet).Range("K1").CurrentRegion.Rows.Count))
        ActiveWorkbook.Worksheets(Worksheet).Cells(3, 17).Value = WorksheetFunction.Min(Worksheets(Worksheet).Range("L2:L" & ActiveWorkbook.Worksheets(Worksheet).Range("K1").CurrentRegion.Rows.Count))
        ActiveWorkbook.Worksheets(Worksheet).Cells(4, 17).Value = WorksheetFunction.Max(Worksheets(Worksheet).Range("J2:J" & ActiveWorkbook.Worksheets(Worksheet).Range("K1").CurrentRegion.Rows.Count))
        
        ActiveWorkbook.Worksheets(Worksheet).Cells(2, 17).Style = "Percent"
        ActiveWorkbook.Worksheets(Worksheet).Cells(3, 17).Style = "Percent"
        
        For d = 2 To ActiveWorkbook.Worksheets(Worksheet).Range("K1").CurrentRegion.Rows.Count
            If ActiveWorkbook.Worksheets(Worksheet).Cells(d, 12).Value = ActiveWorkbook.Worksheets(Worksheet).Cells(2, 17).Value Then
            ActiveWorkbook.Worksheets(Worksheet).Cells(2, 16).Value = ActiveWorkbook.Worksheets(Worksheet).Cells(d, 9).Value
            ElseIf ActiveWorkbook.Worksheets(Worksheet).Cells(d, 12).Value = ActiveWorkbook.Worksheets(Worksheet).Cells(3, 17).Value Then
            ActiveWorkbook.Worksheets(Worksheet).Cells(3, 16).Value = ActiveWorkbook.Worksheets(Worksheet).Cells(d, 9).Value
            ElseIf ActiveWorkbook.Worksheets(Worksheet).Cells(d, 10).Value = ActiveWorkbook.Worksheets(Worksheet).Cells(4, 17).Value Then
            ActiveWorkbook.Worksheets(Worksheet).Cells(4, 16).Value = ActiveWorkbook.Worksheets(Worksheet).Cells(d, 9).Value
            End If
        Next d
        b = 2
        
    Next Worksheet
    
    For Worksheet = 1 To Worksheet_Count
    Worksheets(Worksheet).Columns("A:Z").AutoFit
    Next Worksheet
    
End Sub
