Attribute VB_Name = "Module1"
Sub Stonks()

For Each ws In Worksheets

    Dim Ticker As String
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalStockVolume As Long
    Dim lastrow As Long
    Dim openprice As Double
    Dim closeprice As Double
    Dim SummaryRowTable As Long
    
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"
    
    On Error Resume Next
    
    SummaryRowTable = 2
    TotalStockVolume = ws.Range("G2").Value
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    openprice = ws.Cells(2, 3)
    
For i = 2 To lastrow
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        Ticker = ws.Cells(i, 1).Value
        TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
        
        
        closeprice = ws.Cells(i, 6).Value
        YearlyChange = closeprice - openprice
    
        
        If openprice = 0 Then
        PercentChange = 0
        Else
        PercentChange = (YearlyChange / openprice)
        End If
    
    ws.Range("I" & SummaryRowTable).Value = Ticker
    ws.Range("J" & SummaryRowTable).Value = YearlyChange
    ws.Range("K" & SummaryRowTable).Value = PercentChange
    ws.Range("L" & SummaryRowTable).Value = CLng(TotalStockVolume)
    ws.Range("J" & SummaryRowTable).NumberFormat = "0.00"
    ws.Range("K" & SummaryRowTable).NumberFormat = "0.00%"
       
        openprice = ws.Cells(i + 1, 3).Value
        SummaryRowTable = SummaryRowTable + 1
        TotalStockVolume = 0
    Else
        TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
    
    End If
    
            
        If ws.Cells(i, 10).Value < 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 3
        ElseIf ws.Cells(i, 10).Value >= 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 4
        ElseIf IsEmpty(ws.Cells(i, 10).Value) Then
            ws.Cells(i, 10).Interior.ColorIndex = 5
        End If
        
        If ws.Cells(i, 11).Value < 0 Then
            ws.Cells(i, 11).Interior.ColorIndex = 3
        ElseIf ws.Cells(i, 11).Value >= 0 Then
            ws.Cells(i, 11).Interior.ColorIndex = 4
        ElseIf IsEmpty(ws.Cells(i, 11).Value) Then
            ws.Cells(i, 11).Interior.ColorIndex = 5
        End If
    
Next i
Next ws
End Sub

