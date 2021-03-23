Attribute VB_Name = "Module1"
Sub StockData():

    '-----------------------------------------------
    'Loop through all sheets
    '-----------------------------------------------
    'Dim ws As Worksheet
    For Each ws In Worksheets
    
        ' Defined dimensions of variables
        Dim LastRow As Long
        Dim Ticker As Integer
        Dim AnalysisRow As Integer
        
        
        ' Find Last Row
        'LastRow = ws.Cells(Rows.Count, "A").End(x1Up).Row
        'MsgBox (LastRow)
        
        'Define other variables
        Ticker = 1
        AnalysisRow = 2

        ' Add titles to analysis cells and set to Autofit
        ws.Range("I1").Value = "Ticker"
        
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("J1").EntireColumn.AutoFit
        
        ws.Range("K1").Value = "Percent Change"
        ws.Range("K1").EntireColumn.AutoFit
        
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("L1").EntireColumn.AutoFit
        
            ' Loop through ticker symbols and output one of each to Ticker column
            For i = 2 To LastRow
                If ws.Cells(i + 1, Ticker).Value <> ws.Cells(i, Ticker).Value Then
                    ws.Cells(AnalysisRow, 9).Value = ws.Cells(i, Ticker).Value
        
                    AnalysisRow = AnalysisRow + 1
            End If
        
        Next i
            
        
    Next ws

End Sub
