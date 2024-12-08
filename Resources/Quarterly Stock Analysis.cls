VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub QuarterlyStockAnalysis()
    Dim ws As Worksheet
    Dim Ticker As String
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim QuarterlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestVolume As Double
    Dim GreatestIncreaseTicker As String
    Dim GreatestDecreaseTicker As String
    Dim GreatestVolumeTicker As String
    Dim LastRow As Long
    Dim i As Long
    Dim StartRow As Long
    
    ' Initialize variables for tracking the greatest changes and volumes
    GreatestIncrease = 0
    GreatestDecrease = 0
    GreatestVolume = 0
    
    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Setting headers for output
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Volume"
        
        ' Initialize variables for looping through stocks
        TotalVolume = 0
        StartRow = 2
        i = 2
        Dim OutputRow As Integer
        OutputRow = 2
        
        ' Loop through all rows in the worksheet
        While i <= LastRow
            ' Check if we are still on the same ticker
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
                OpeningPrice = ws.Cells(StartRow, 3).Value
                ClosingPrice = ws.Cells(i, 6).Value
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                
                ' Calculate Quarterly and Percent Change
                QuarterlyChange = ClosingPrice - OpeningPrice
                If OpeningPrice <> 0 Then
                    PercentChange = (QuarterlyChange / OpeningPrice)
                Else
                    PercentChange = 0
                End If
                
                ' Output results to the worksheet
                ws.Cells(OutputRow, 9).Value = Ticker
                ws.Cells(OutputRow, 10).Value = QuarterlyChange
                ws.Cells(OutputRow, 11).Value = PercentChange
                ws.Cells(OutputRow, 12).Value = TotalVolume
                
                ' Apply color formatting to Quarterly Change
                If QuarterlyChange > 0 Then
                    ws.Cells(OutputRow, 10).Interior.Color = RGB(144, 238, 144) ' Light green for positive
                ElseIf QuarterlyChange < 0 Then
                    ws.Cells(OutputRow, 10).Interior.Color = RGB(255, 99, 71) ' Light red for negative
                End If
                
                ' Format Percent Change as a percentage
                ws.Cells(OutputRow, 11).NumberFormat = "0.00%"
                
                ' Check for greatest increase, decrease, and volume
                If PercentChange > GreatestIncrease Then
                    GreatestIncrease = PercentChange
                    GreatestIncreaseTicker = Ticker
                End If
                If PercentChange < GreatestDecrease Then
                    GreatestDecrease = PercentChange
                    GreatestDecreaseTicker = Ticker
                End If
                If TotalVolume > GreatestVolume Then
                    GreatestVolume = TotalVolume
                    GreatestVolumeTicker = Ticker
                End If
                
                ' Reset variables for the next ticker
                TotalVolume = 0
                OutputRow = OutputRow + 1
                StartRow = i + 1
            Else
                ' Accumulate volume if same ticker
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            End If
            i = i + 1
        Wend
        
        ' Display results of greatest changes and volumes
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        ws.Range("P2").Value = GreatestIncreaseTicker
        ws.Range("Q2").Value = GreatestIncrease
        ws.Range("P3").Value = GreatestDecreaseTicker
        ws.Range("Q3").Value = GreatestDecrease
        ws.Range("P4").Value = GreatestVolumeTicker
        ws.Range("Q4").Value = GreatestVolume
        
        ' Format Greatest % Increase and Greatest % Decrease as percentages
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
    Next ws
End Sub

Private Sub Worksheet_Activate()

End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub
