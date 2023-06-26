Attribute VB_Name = "Module1"
Sub Module_2_Challenge_Part_1_and_Part_2()

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
    
    If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        FirstValue = ws.Cells(i, 3).Value
    End If
    
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        TickName = ws.Cells(i, 1).Value
        LastValue = ws.Cells(i, 6).Value
        YearlyChange = LastValue - FirstValue
        
        
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
    
    '-----------------------PART 2---------------------------








    'Create the table titles
    ws.Cells(2, 15).Value = "Greatst % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"

    PercentMax = WorksheetFunction.Max(ws.Range("K2:K" & LastRow))
    PercentMin = WorksheetFunction.Min(ws.Range("K2:K" & LastRow))
    VolumeMax = WorksheetFunction.Max(ws.Range("L2:L" & LastRow))
    
    For i = 2 To LastRow
    'Define Loop Variables
   
        'Find the maximum percent
        If ws.Cells(i, 11).Value = PercentMax Then
            
            ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
            ws.Cells(2, 17).Value = PercentMax
        'Find the minimum percent
        ElseIf ws.Cells(i, 11).Value = PercentMin Then
        
            ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
            ws.Cells(3, 17).Value = PercentMin
        'Find the maximum Volume
        ElseIf ws.Cells(i, 12).Value = VolumeMax Then
        
            ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
            ws.Cells(4, 17).Value = VolumeMax
            
        End If
        
    Next i
    
    'Set the formatting for the table
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 17).NumberFormat = "0.00%"
    
    'Set Conditional Formatting for the Yearly Differences
    For i = 2 To LastRow
    
        If ws.Cells(i, 10).Value <= 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 3
        Else
            ws.Cells(i, 10).Interior.ColorIndex = 4
        End If
    Next i

  
Next ws

End Sub

