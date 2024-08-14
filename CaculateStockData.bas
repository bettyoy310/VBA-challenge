VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub caculateStockData()

    'Create a variable to hold the counter
    Dim ws As Worksheet
    Dim ticker As String
    Dim QuarterlyChange As Double
    Dim PercentChange As Double
    Dim TotalStockVolume As Double
    Dim LastRow As Long
    Dim CurrentRow As Long
    Dim startOpen As Double
    Dim endClose As Double
    Dim outputRow As Long
    
    Dim MaxPercentIncrease As Double
    Dim MaxPercentDecrease As Double
    Dim MaxTotalVolum As Double
    Dim MaxIncreaseTicker As String
    Dim MaxDecreaseTicker As String
    Dim MaxVolumeTicker As String
    
    MaxPercentIncrease = -999999
    MaxPercentDecrease = 999999
    MaxVolume = 0
   
    'set the header and LastRow
    For Each ws In Worksheets
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        'set an initial variable for each ticker in the summary table
        outputRow = 2
        ticker = ws.Cells(2, 1).Value
        startOpen = ws.Cells(2, 3).Value
        TotalStockVolume = ws.Cells(2, 7).Value
        
        'Loop through all tickers
        For CurrentRow = 3 To LastRow
            TotalStockVolume = TotalStockVolume + ws.Cells(CurrentRow, 7).Value
            'check if we are still within the same ticker,if it is not...
            If ws.Cells(CurrentRow + 1, 1).Value <> ticker Or CurrentRow = LastRow Then
                endClose = ws.Cells(CurrentRow, 6).Value
                QuarterlyChange = endClose - startOpen
                If startOpen <> 0 Then
                    PercentChange = QuarterlyChange / startOpen
                Else
                    PercentChange = 0
                End If

                ws.Cells(outputRow, 9).Value = ticker
                ws.Cells(outputRow, 10).Value = QuarterlyChange
                ws.Cells(outputRow, 11).Value = PercentChange
                ws.Cells(outputRow, 11).NumberFormat = "0.00%"
                ws.Cells(outputRow, 12).Value = TotalStockVolume
                
                'Apply colors based on QuarterlyChange
                If QuarterlyChange < 0 Then
                    ws.Cells(outputRow, 10).Interior.ColorIndex = 3
                ElseIf QuarterlyChange > 0 Then
                    ws.Cells(outputRow, 10).Interior.ColorIndex = 4
                End If
                
                'caculate the greatest values
                If PercentChange > MaxPercentIncrease Then
                    MaxPercentIncrease = PercentChange
                    MaxIncreaseTicker = ticker
                End If
                
                
                If PercentChange < MaxPercentDecrease Then
                    MaxPercentDecrease = PercentChange
                    MaxDecreaseTicker = ticker
                End If
                
                If TotalStockVolume > MaxVolume Then
                    MaxVolume = TotalStockVolume
                    MaxVolumeTicker = ticker
                End If
                
                'add one to the outputrow
                outputRow = outputRow + 1
                ticker = ws.Cells(CurrentRow + 1, 1).Value
                startOpen = ws.Cells(CurrentRow + 1, 3).Value
                TotalStockVolume = 0
            End If
                
                              
        Next CurrentRow

        'print the information in the summary of greatest values
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        ws.Cells(2, 15).Value = "Greatest%Increase"
        ws.Cells(2, 16).Value = MaxIncreaseTicker
        ws.Cells(2, 17).Value = MaxPercentIncrease
        ws.Cells(2, 17).NumberFormat = "0.00%"
        
        ws.Cells(3, 15).Value = "Greatest%Decrease"
        ws.Cells(3, 16).Value = MaxDecreaseTicker
        ws.Cells(3, 17).Value = MaxPercentDecrease
        ws.Cells(3, 17).NumberFormat = "0.00%"
        
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(4, 16).Value = MaxVolumeTicker
        ws.Cells(4, 17).Value = MaxVolume
        
    Next ws
       
End Sub






