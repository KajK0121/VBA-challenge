VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub TickersStock()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentageChange As Double
    Dim totalVolume As Double
    Dim P As Long
    
    For Each ws In ThisWorkbook.Worksheets
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
        totalVolume = 0
        openingPrice = ws.Cells(2, 3).Value
        P = 1
        
        ws.Cells(P, 9).Value = "Ticker Symbol"
        ws.Cells(P, 10).Value = "Yearly Change"
        ws.Cells(P, 11).Value = "Percentage Change"
        ws.Cells(P, 12).Value = "Total Stock Volume"
            
        For i = 2 To lastRow
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                closingPrice = ws.Cells(i, 6).Value
                yearlyChange = closingPrice - openingPrice
                If openingPrice <> 0 Then
                    percentageChange = (yearlyChange / openingPrice) * 100
                Else
                    percentageChange = 0
                End If
        
                P = P + 1
                ws.Cells(P, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(P, 10).Value = yearlyChange
                ws.Cells(P, 11).Value = percentageChange / 100
                ws.Cells(P, 12).Value = totalVolume
      
                totalVolume = totalVolume + ws.Cells(i, 7).Value
               
                openingPrice = ws.Cells(i + 1, 3).Value
                totalVolume = 0
            Else
                totalVolume = totalVolume + ws.Cells(i, 7).Value
            End If
        Next i
        
        With ws.Range("J2:J" & P)
            .FormatConditions.Delete
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
            .FormatConditions(1).Interior.ColorIndex = 3
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
            .FormatConditions(2).Interior.ColorIndex = 4
        End With
        
        Dim percentageChangeAnalyse As Double
        Dim totalVolumeAnalyse As Double
        Dim greatestIncrease As Double
        Dim greatestDecrease As Double
        Dim greatestVolume As Double
        Dim tickerGreatestIncrease As String
        Dim tickerGreatestDecrease As String
        Dim tickerGreatestVolume As String
        
        lastRowAnalyse = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
        
        greatestIncrease = -1
        greatestDecrease = 1
        greatestVolume = 0
        
        For i = 2 To lastRowAnalyse
            percentageChangeAnalyse = ws.Cells(i, 11).Value
            totalVolumeAnalyse = ws.Cells(i, 12).Value
            tickerAnalyse = ws.Cells(i, 9).Value
            
            If percentageChangeAnalyse > greatestIncrease Then
                greatestIncrease = percentageChangeAnalyse
                tickerGreatestIncrease = tickerAnalyse
            End If
            
            If percentageChangeAnalyse < greatestDecrease Then
                greatestDecrease = percentageChangeAnalyse
                tickerGreatestDecrease = tickerAnalyse
            End If
            
            If totalVolumeAnalyse > greatestVolume Then
                greatestVolume = totalVolumeAnalyse
                tickerGreatestVolume = tickerAnalyse
            End If
        Next i
        
        ws.Cells(1, 15).Value = "Ticker Symbol"
        ws.Cells(1, 16).Value = "Value"
        
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(2, 15).Value = tickerGreatestIncrease
        ws.Cells(2, 16).Value = greatestIncrease
        
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(3, 15).Value = tickerGreatestDecrease
        ws.Cells(3, 16).Value = greatestDecrease
        
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(4, 15).Value = tickerGreatestVolume
        ws.Cells(4, 16).Value = greatestVolume
        
        For i = 2 To lastRowAnalyse
            ws.Cells(i, 11).NumberFormat = "0.00%"
        Next i
    Next ws
End Sub

