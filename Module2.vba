'VBA code
Sub stock()

For Each ws In Worksheets
    Dim Ticker As String
    '_____
    Dim Volume As Double
    Volume = 0
    '____
    Dim Table_Row As Double
    Table_Row = 2
    '____
    Dim finaltrade As Double
    Dim firsttrade As Double
    
    '____
    Dim WorksheetName As String
    '____
    Dim Yearchange As Double
    Dim Percentchange As Double
    '____
    'Sheets("2018").Range("I1:L1").Copy
    Dim max As Double
    Dim min As Double
    Dim maxvolume As Double
    
    Dim maxrow, minrow, mvrow As String
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    LastRow2 = ws.Cells(Rows.Count, 11).End(xlUp).Row
    firsttrade = ws.Cells(2, 3)
    WorksheetName = ws.Name
    
    For i = 2 To LastRow
            
        If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
            Ticker = ws.Cells(i, 1)
            
            finaltrade = ws.Cells(i, 6)
            Yearchange = finaltrade - firsttrade
            
            Percentchange = (Yearchange / firsttrade)
            
            Volume = Volume + ws.Cells(i, 7)
            
            'Insert data
            ws.Range("I" & Table_Row) = Ticker
            ws.Range("L" & Table_Row) = Volume
            ws.Range("J" & Table_Row) = Yearchange
            ws.Range("K" & Table_Row) = Percentchange
            'Format Percent for Percent Change
            ws.Range("K" & Table_Row).NumberFormat = "00.00%"
            'Color base on changes
            If Yearchange > 0 Then
                ws.Range("J" & Table_Row).Interior.ColorIndex = 4
            Else
                ws.Range("J" & Table_Row).Interior.ColorIndex = 3
            End If
            
            
            Table_Row = Table_Row + 1
            Volume = 0
            firsttrade = ws.Cells(i + 1, 3)
        Else
            Volume = Volume + ws.Cells(i, 7)
        End If
    Next i
    'Sheets(WorksheetName).Activate
    'ws.Range("I1:L1").Select
    'ActiveSheet.Paste
    'Application.CutCopyMode = False
    'Sheets(WorksheetName).Columns("I:L").AutoFit
    max = ws.Cells(2, 11)
    min = ws.Cells(2, 11)
    maxvolume = ws.Cells(2, 12)
    For j = 2 To LastRow2
        If ws.Cells(j + 1, 11) > max Then
            max = ws.Cells(j + 1, 11)
            maxrow = ws.Cells(j + 1, 9)
        End If
        
        If ws.Cells(j + 1, 11) < min Then
            min = ws.Cells(j + 1, 11)
            minrow = ws.Cells(j + 1, 9)
        End If
        
        If ws.Cells(j + 1, 12) > maxvolume Then
            maxvolume = ws.Cells(j + 1, 12)
            mvrow = ws.Cells(j + 1, 9)
        End If
        
    Next j
    ws.Range("O2") = maxrow
    ws.Range("P2") = max
    ws.Range("P2").Interior.ColorIndex = 4
    ws.Range("P2").NumberFormat = "00.00%"
    
    ws.Range("O3") = minrow
    ws.Range("P3") = min
    ws.Range("P3").Interior.ColorIndex = 3
    ws.Range("P3").NumberFormat = "00.00%"
    
    ws.Range("O4") = mvrow
    ws.Range("P4") = maxvolume
Next ws

End Sub

