Attribute VB_Name = "Module1"
Sub StockMkt()

    On Error Resume Next
    
    Dim TickerName As String
    Dim SummaryRow As Integer
    Dim LastRow As Double
    Dim YearOpen As Double
    Dim YearClose As Double
    Dim Volume As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim WS As Worksheet
    
    
    For Each WS In Worksheets
        
        SummaryRow = 2
            
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
        
        WS.Cells(1, 9).Value = "Ticker Symbol"
        WS.Cells(1, 10).Value = "Yearly Change"
        WS.Cells(1, 11).Value = "Percent Change"
        WS.Cells(1, 12).Value = "Total Stock Volume"
        
               
        For i = 2 To LastRow
        
            If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then
        
            
            TickerName = WS.Cells(i, 1).Value
            YearOpen = WS.Cells(i, 3).Value
            YearClose = WS.Cells(i, 6).Value
            Volume = WS.Cells(i, 7).Value
            
                    
            YearlyChange = YearClose - YearOpen
            PercentChange = (YearClose - YearOpen) / YearClose

                
            WS.Cells(SummaryRow, 9).Value = TickerName
            WS.Cells(SummaryRow, 10).Value = YearlyChange
            WS.Cells(SummaryRow, 11).Value = PercentChange
            WS.Cells(SummaryRow, 12).Value = Volume
            
            SummaryRow = SummaryRow + 1
            
            Volume = 0
      
            End If
    
        Next i
        
        WS.Cells(i, 10).NumberFormat = "0.00%"
        
    
        For i = 2 To LastRow
            If WS.Cells(i, 10).Value >= 0 Then
            WS.Cells(i, 10).Interior.ColorIndex = 4
    
            Else
            WS.Cells(i, 10).Interior.ColorIndex = 3
    
        End If
        Next i
        MsgBox WS.Name
    Next
    
End Sub

