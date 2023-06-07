Sub Ticker():
   'Fill ticker Column
   Dim i As Long
   Dim j As Long
   Dim Total As Double
   Dim Start As Long
   Dim First As Double
   Dim Percent_Change As Double
   Dim Max As Integer
   Dim Min As Integer
   Dim Greatest As Double
    Dim Increase As Long
    Dim Answer As Double
    Dim ws As Worksheet
    
    
   'Start Worksheet Loop
   'https://excelchamps.com/vba/loop-sheets/
For Each ws In ThisWorkbook.Worksheets
   
   Total = 0
   j = 0
   Start = 2
   RowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row


   First = ws.Cells(2, 3).Value
   
   
    For i = 2 To RowCount
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Total = Total + ws.Cells(i, 7).Value
            Change = ws.Cells(i, 6).Value - First
            Percent_Change = Change / First
            
            First = ws.Cells(i + 1, 3).Value
            
            ws.Range("i" & 2 + j).Value = ws.Cells(i, 1).Value
            ws.Range("j" & 2 + j).Value = Change
            ws.Range("k" & 2 + j).Value = Percent_Change
            ws.Range("l" & 2 + j).Value = Total
            
           'Format Output
            ws.Range("j" & 2 + j).NumberFormat = "0.00"
            ws.Range("k" & 2 + j).NumberFormat = "0.00%"
            
            Select Case Change
                Case Is > 0
                    ws.Range("j" & 2 + j).Interior.ColorIndex = 4
                Case Is < 0
                    ws.Range("j" & 2 + j).Interior.ColorIndex = 3
                Case Else
                    ws.Range("j" & 2 + j).Interior.ColorIndex = 0
            End Select
            
            
            
            
            
            j = j + 1
            Total = 0
            Change = ws.Cells(i + 1, 3).Value
    
            
            
        End If
        Total = Total + ws.Cells(i, 7).Value
        
    Next i
    
    
    
    
    
    
    'Second Table (the small table to the right)
        
    
        
        
    ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & RowCount)) * 100
    ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & RowCount)) * 100
    ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & RowCount))
    
    'Formatting For Second Table
    ws.Cells(17, 4).NumberFormat = "0.00"
    
    
    Increase = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & RowCount)), ws.Range("K2:K" & RowCount), 0)
    Decrease = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & RowCount)), ws.Range("K2:K" & RowCount), 0)
    V_Number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & RowCount)), ws.Range("L2:L" & RowCount), 0)
    
    
    ws.Range("P2") = ws.Cells(Increase + 1, 9)
    ws.Range("P3") = ws.Cells(Decrease + 1, 9)
    ws.Range("P4") = ws.Cells(V_Number + 1, 9)
    
    Next ws
    
End Sub




