Attribute VB_Name = "Module1"
Sub StockData():

    '-----------------------------------------------
    'Looped through all sheets
    '-----------------------------------------------

    For Each ws In Worksheets
    
        ' Defined dimensions of gloabl variables
        Dim LastRow As Long
        Dim Ticker As Integer
        Dim AnalysisRow As Integer
        Dim OpenValue As Integer
        Dim ClosingValue As Integer
        Dim VolumeColumn As Integer
        Dim StartNextTicker As Long
        Dim ClosingPrice As Double
        Dim OpeningPrice2 As Double
        Dim ClosingPrice2 As Double
        Dim FirstRowOpen As Double
        'Dim FirstRowClose As Double
        
        ' Found Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
       
        
        'Defined other variables
        Ticker = 1
        AnalysisRow = 3
        OpeningValueColumn = 3
        ClosingValueColumn = 6
        VolumeColumn = 7
        StartNextTicker = 0
        ClosingPrice = 0
        OpeningPrice2 = 0
        ClosingPrice2 = 0
        FirstRowOpen = 0

        ' Added titles to analysis cells and set to Autofit
        ws.Range("I1").Value = "Ticker"
        
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("J1").EntireColumn.AutoFit
        
        ws.Range("K1").Value = "Percent Change"
        ws.Range("K1").EntireColumn.AutoFit
        ws.Range("K1").NumberFormat = "0.00%"
        
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("L1").EntireColumn.AutoFit
        
        
        ' Process first row of summary table
        'ws.Cells(2, 9).Value = ws.Cells(2, Ticker)
        FirstRowVolume = ws.Cells(2, VolumeColumn).Value
               
        
            ' Looped through ticker symbols and output one of each to Ticker column
            
            ' Looped through first ticker symbol to process first row of summary table
            'For i = 2 To 526
                For j = 2 To LastRow
                    'Found where the ticker changed
                    If ws.Cells(j + 1, Ticker).Value <> ws.Cells(j, Ticker).Value Then
                        ws.Cells(2, 9).Value = ws.Cells(j, Ticker).Value
                        ' Set closing/opening prices for first Ticker value
                        FirstRowOpen = ws.Cells(2, OpeningValueColumn).Value
                        ClosingPrice = ws.Cells(j, ClosingValueColumn).Value
                        ' Processed first row of summary table
                        ws.Cells(2, 10).Value = ClosingPrice - FirstRowOpen
                        ws.Cells(2, 11).Value = ((ClosingPrice - FirstRowOpen) / FirstRowOpen)
                        ws.Range("L2") = WorksheetFunction.Sum(Range(ws.Cells(2, VolumeColumn), ws.Cells(j, VolumeColumn)))
                        ' Set global variable to use in next conditional
                        OpeningPrice2 = ws.Cells(j + 1, OpeningValueColumn).Value
                        StartNextTicker = j + 1
                        ' Exited after finding first difference
                        Exit For
                    
                    'End first conditional
                    End If
                
                
                Next j
                 
                For k = (j + 1) To LastRow
                'Start conditional for the rest of the rows
                    If (ws.Cells(k + 1, Ticker).Value <> ws.Cells(k, Ticker).Value And OpeningPrice2 <> 0) Then
                        ws.Cells(AnalysisRow, 9).Value = ws.Cells(k, Ticker).Value
                    
                        
                        ClosingPrice = ws.Cells(k, ClosingValueColumn).Value
                        ws.Cells(AnalysisRow, 10).Value = ClosingPrice - OpeningPrice2
                        ws.Cells(AnalysisRow, 11).Value = ((ClosingPrice - OpeningPrice2) / OpeningPrice2)
                        ws.Cells(AnalysisRow, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(StartNextTicker, VolumeColumn), ws.Cells(k, VolumeColumn)))
                        
                        OpeningPrice2 = ws.Cells(k + 1, OpeningValueColumn).Value
                        StartNextTicker = k + 1
                      
                    'ClosingPrice2 = ws.Cells(i, ClosingValueColumn).Value
                
                    'ws.Cells(AnalysisRow, 10).Value = ClosingPrice - OpeningPrice
        
                        AnalysisRow = AnalysisRow + 1
                        
                    ElseIf (ws.Cells(k + 1, Ticker).Value <> ws.Cells(k, Ticker).Value And OpeningPrice2 = 0) Then
                        ws.Cells(AnalysisRow, 9).Value = ws.Cells(k, Ticker).Value
                        
                        ClosingPrice = ws.Cells(k, ClosingValueColumn).Value
                        ws.Cells(AnalysisRow, 10).Value = ClosingPrice - OpeningPrice2
                        ws.Cells(AnalysisRow, 11).Value = "Undefined"
                        ws.Cells(AnalysisRow, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(StartNextTicker, VolumeColumn), ws.Cells(k, VolumeColumn)))
                        
                        OpeningPrice2 = ws.Cells(k + 1, OpeningValueColumn).Value
                        StartNextTicker = k + 1
                        
                        AnalysisRow = AnalysisRow + 1
                    End If
        
                Next k
            
        
    Next ws

End Sub
