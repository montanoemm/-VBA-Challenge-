Attribute VB_Name = "Module1"

Sub MaxValues()

Dim LastRow As Long
Dim PercentInc As Double
Dim PercenDec As Double
Dim GreatTotV As Double



'loop through all worksheets
For Each ws In ActiveWorkbook.Worksheets

'determine last row
    LastRow = ws.Range("A" & Rows.Count).End(xlUp).row


    'loop max increase
    
    maxvalue = WorksheetFunction.Max(ws.Range("K2:K" & LastRow))
    
    maxindex = WorksheetFunction.Match(maxvalue, ws.Range("K2:K" & LastRow), 0)
    
    ws.Range("Q2") = "%" & maxvalue * 100
    
    ws.Range("P2") = Cells(maxindex + 1, 9)
    
    '2nd loop max decrease
    
   
    minvalue = WorksheetFunction.Min(ws.Range("K2:K" & LastRow))
    
    minindex = WorksheetFunction.Match(minvalue, ws.Range("K2:K" & LastRow), 0)
    
    ws.Range("Q3") = "%" & minvalue * 100
    
    ws.Range("P3") = Cells(minindex + 1, 9)
    
    
    '3rd loop volume
        
    
    maxvolvalue = WorksheetFunction.Max(ws.Range("L2:L" & LastRow))
    
    maxvolumeindex = WorksheetFunction.Match(maxvolvalue, ws.Range("L2:L" & LastRow), 0)
    
    ws.Range("Q4") = maxvolvalue
    
    ws.Range("P4") = Cells(maxvolumeindex + 1, 9)
         
    
    
Next ws
    
    

End Sub

