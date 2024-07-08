VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub mod2()
    Dim q1 As Worksheet
    Dim lastRow As Long
    Dim tickerRow As Long
    Dim i As Long
    Dim ticker As String
    Dim startPrice As Double
    Dim endPrice As Double
    Dim totalVolume As Variant
    Dim quarterlyChange As Double
    Dim percentChange As Double

    Set q1 = ThisWorkbook.Sheets("Q2")
    lastRow = q1.Cells(q1.Rows.Count, "A").End(xlUp).Row

    q1.Cells(1, 9).Value = "Ticker"
    q1.Cells(1, 10).Value = "Quarterly Change"
    q1.Cells(1, 11).Value = "Percent Change"
    q1.Cells(1, 12).Value = "Total Stock Volume"

    tickerRow = 2
    
    maxIncrease = -100000
    maxDecrease = 100000
    maxVolume = 0

    For i = 2 To lastRow
        ticker = q1.Cells(i, 1).Value
        startPrice = q1.Cells(i, 3).Value
        totalVolume = 0

        Do While q1.Cells(i, 1).Value = ticker
            endPrice = q1.Cells(i, 6).Value
            totalVolume = totalVolume + q1.Cells(i, 7).Value
            i = i + 1
        Loop
        i = i - 1

        quarterlyChange = endPrice - startPrice
        If startPrice <> 0 Then
            percentChange = (quarterlyChange / (startPrice)) * 1
        Else
            percentChange = 0
        End If

        q1.Cells(tickerRow, 9).Value = ticker
        q1.Cells(tickerRow, 10).Value = quarterlyChange
        q1.Cells(tickerRow, 11).Value = percentChange
        q1.Cells(tickerRow, 11).NumberFormat = "0.00%"
        q1.Cells(tickerRow, 12).Value = totalVolume

        If quarterlyChange > 0 Then
            q1.Cells(tickerRow, 10).Interior.Color = RGB(0, 255, 0)
        Else
            q1.Cells(tickerRow, 10).Interior.Color = RGB(255, 0, 0)
        End If
        
        If percentChange > maxIncrease Then
            maxIncrease = percentChange
            maxIncreaseTicker = ticker
        End If

        If percentChange < maxDecrease Then
            maxDecrease = percentChange
            maxDecreaseTicker = ticker
        End If

        If totalVolume > maxVolume Then
            maxVolume = totalVolume
            maxVolumeTicker = ticker
        End If

        tickerRow = tickerRow + 1
    Next i
    
    q1.Cells(2, 15).Value = "Greatest % Increase"
    q1.Cells(2, 16).Value = maxIncreaseTicker
    q1.Cells(2, 17).Value = maxIncrease
    q1.Cells(2, 17).NumberFormat = "0.00%"

    q1.Cells(3, 15).Value = "Greatest % Decrease"
    q1.Cells(3, 16).Value = maxDecreaseTicker
    q1.Cells(3, 17).Value = maxDecrease
    q1.Cells(3, 17).NumberFormat = "0.00%"

    q1.Cells(4, 15).Value = "Greatest Total Volume"
    q1.Cells(4, 16).Value = maxVolumeTicker
    q1.Cells(4, 17).Value = maxVolume
    
    q1.Columns("I:Q").AutoFit
End Sub

