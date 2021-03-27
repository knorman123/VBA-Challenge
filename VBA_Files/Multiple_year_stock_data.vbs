Attribute VB_Name = "Module1"
Sub StockData():

    '-----------------------------------------------
    'Loop through all sheets
    '-----------------------------------------------

    For Each ws In Worksheets
    
        ' Define dimensions of global variables
        Dim LastRow As Long
        Dim Ticker As Integer
        Dim AnalysisRow As Integer
        Dim OpeningValueColumn As Integer
        Dim ClosingValueColumn As Integer
        Dim VolumeColumn As Integer
        Dim StartNextTicker As Long
        Dim ClosingPrice As Double
        Dim OpeningPrice2 As Double
        Dim FirstRowOpen As Double
        
        
        ' Found Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
       
        
        ' Set global variables
        Ticker = 1
        AnalysisRow = 3
        OpeningValueColumn = 3
        ClosingValueColumn = 6
        VolumeColumn = 7
        StartNextTicker = 0
        ClosingPrice = 0
        OpeningPrice2 = 0
        FirstRowOpen = 0

        ' Add titles to analysis cells and set to Autofit
        ws.Range("I1").Value = "Ticker"
        
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("J1").EntireColumn.AutoFit
        
        ws.Range("K1").Value = "Percent Change"
        ws.Range("K1").EntireColumn.AutoFit
        ws.Range("K1").NumberFormat = "0.00%"
        
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("L1").EntireColumn.AutoFit
           
        
            '-----------------------------------------------------------------------------------
            ' Loop through first ticker symbol to process first row of summary table
            '-----------------------------------------------------------------------------------
            For j = 2 To LastRow
                ' Find where the ticker changes, output name to Ticker column
                If ws.Cells(j + 1, Ticker).Value <> ws.Cells(j, Ticker).Value Then
                    ws.Cells(2, 9).Value = ws.Cells(j, Ticker).Value
                    
                    ' Set closing/opening prices for first Ticker value
                    FirstRowOpen = ws.Cells(2, OpeningValueColumn).Value
                    ClosingPrice = ws.Cells(j, ClosingValueColumn).Value
                    
                    ' Process first row of summary table
                    ws.Cells(2, 10).Value = ClosingPrice - FirstRowOpen
                    ws.Cells(2, 11).Value = ((ClosingPrice - FirstRowOpen) / FirstRowOpen)
                    ws.Range("L2") = WorksheetFunction.Sum(Range(ws.Cells(2, VolumeColumn), ws.Cells(j, VolumeColumn)))
                    
                    ' Set global variables to use in next conditional
                    OpeningPrice2 = ws.Cells(j + 1, OpeningValueColumn).Value
                    StartNextTicker = j + 1
                    
                    ' Exit after finding last row of first ticker
                    Exit For
                    
                ' End first conditional
                End If
                
                
            Next j
                 
            '---------------------------------------------------------------------------------
            ' Loop through the rest of the tickers
            '---------------------------------------------------------------------------------
            For k = (j + 1) To LastRow
                ' Start conditional for the rest of the rows to find row where ticker changes and output name to Ticker Column
                ' Make sure the opening price doesn't equal 0
                If (ws.Cells(k + 1, Ticker).Value <> ws.Cells(k, Ticker).Value And OpeningPrice2 <> 0) Then
                    ws.Cells(AnalysisRow, 9).Value = ws.Cells(k, Ticker).Value
                    
                    ' Grab Opening and Closing Prices, calculate yearly change, percent change, total stock volume
                    ClosingPrice = ws.Cells(k, ClosingValueColumn).Value
                    ws.Cells(AnalysisRow, 10).Value = ClosingPrice - OpeningPrice2
                    ws.Cells(AnalysisRow, 11).Value = ((ClosingPrice - OpeningPrice2) / OpeningPrice2)
                    ws.Cells(AnalysisRow, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(StartNextTicker, VolumeColumn), ws.Cells(k, VolumeColumn)))
                        
                    ' Reset variables for next loop
                    OpeningPrice2 = ws.Cells(k + 1, OpeningValueColumn).Value
                    StartNextTicker = k + 1
                    AnalysisRow = AnalysisRow + 1
                        
                ' Same conditional, run if the opening price is equal to 0
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
            
        ' Define dimension of columns to format
        Dim YearlyChange As Range
        Dim PercentChange As Range
        
        ' Set Columns as variables
        Set YearlyChange = ws.Range("J2:J" & LastRow)
        Set PercentChange = ws.Range("K:K")
        
        ' Clear any previous formatting
        YearlyChange.FormatConditions.Delete
        PercentChange.FormatConditions.Delete
        
        ' Set conditional formatting for positive/negative yearly change
        YearlyChange.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
            Formula1:="=0"
        YearlyChange.FormatConditions(1).Interior.Color = RGB(0, 255, 0)
        
        YearlyChange.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
            Formula1:="=0"
        YearlyChange.FormatConditions(2).Interior.Color = RGB(255, 0, 0)
        
        ' Set percent change column to percent number formatting
        PercentChange.NumberFormat = "0.00%"
        
    Next ws

End Sub


