Attribute VB_Name = "Module2"
Sub Module_2_Challenge_Part_1()

End Sub

For Each ws In Worksheets

'Define Part 1 Variables
Dim i As Long
Dim StockWorksheet As String
Dim LastRow As Long
LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
StockWorksheet = ws.Name
Dim TickName As String

Dim FirstValue As Double
Dim LastValue As Double

Dim YearlyChange As Double
YearlyChange = 0
Dim PercentChange As Double
PercentChange = 0
Dim RowCount As Integer
RowCount = 2
Dim TotalStockVolume As Double
TotalStockVolume = 0


'Part 2 Variables
Dim PercentMax As Double
Dim PercentMin As Double
Dim VolumeMax As Double



'TotalStockVolume

'Adding Titles to Columns
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"


For i = 2 To LastRow
    'If cell above is not equal
    If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        FirstValue = ws.Cells(i, 3).Value
    End If
    
    'If cell below is not equal
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        TickName = ws.Cells(i, 1).Value
        LastValue = ws.Cells(i, 6).Value
        YearlyChange = LastValue - FirstValue
        
    'If first value doesn't equal 0
        If FirstValue <> 0 Then
            PercentChange = YearlyChange / FirstValue
        Else
            PercentChange = 0
    End If
        
        ws.Cells(RowCount, 9).Value = TickName
        ws.Cells(RowCount, 10).Value = YearlyChange
        ws.Cells(RowCount, 11).Value = PercentChange
        
        ' Move down the row counter
        RowCount = RowCount + 1
        
        
        FirstValue = 0
        
    End If
Next i

    RowCount = 2

For i = 2 To LastRow
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
        ws.Cells(RowCount, 12).Value = TotalStockVolume
        
        RowCount = RowCount + 1
        
        TotalStockVolume = 0
        
    Else
        
        TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
    End If
    
Next i



    For i = 2 To LastRow
        ws.Cells(i, 11).NumberFormat = "0.00%"
        
    Next i
    
Next ws

End Sub
