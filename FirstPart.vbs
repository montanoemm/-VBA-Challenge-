Attribute VB_Name = "Module3"
Sub test3()


'setting variables

'Dim ws As Worksheet
Dim Ticker As String
Dim OpenValue As Double
Dim CloseValue As Double
Dim YearlyChange As Double
Dim PercentChange As Double
Dim Volume As Long
Dim volume_total As Double
Dim SummaryTbrow As Integer
Dim LastRow As Long
'Dim STblrow As Range
Dim start As Long


For Each ws In ActiveWorkbook.Worksheets


' set value to variable
volume_total = 0
SummaryTbrow = 2
start = 2


    
    'create headers for columns I, J, K, L
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
        'create columns for rows O2,O3,O4
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest & Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        'create headers for columns P,Q
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        
    
    'determine last row
    LastRow = ws.Range("A" & Rows.Count).End(xlUp).row
    
    
    
    'loop through all tickers
    For i = 2 To LastRow
    
        If i = LastRow + 1 Then
            End If
            
    
    'Set ticker name by checking we are in the same ticker firs
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        Ticker = ws.Cells(i, 1).Value
    
        'adding total volume
          volume_total = volume_total + ws.Cells(i, 7).Value
        
        'opening value
         OpenValue = ws.Cells(start, 3).Value
         
        'closing value
        CloseValue = ws.Cells(i, 6).Value
         
        'Yearly Change calc
        YearlyChange = CloseValue - OpenValue
        'Percent Change
        PercentChange = YearlyChange / OpenValue
        start = i + 1
        
         
         'print ticker name, yearly change, to summary table
         ws.Cells(SummaryTbrow, 9).Value = Ticker
         ws.Cells(SummaryTbrow, 10).Value = YearlyChange
         ws.Cells(SummaryTbrow, 11).Value = FormatPercent(PercentChange, 2)
         ws.Cells(SummaryTbrow, 12).Value = volume_total
         
            'format Yearly Change
            If YearlyChange >= 0 Then
            
            ws.Cells(SummaryTbrow, 10).Interior.ColorIndex = 4
            
            Else
            
            ws.Cells(SummaryTbrow, 10).Interior.ColorIndex = 3
            
            End If
            
    
         'add one to the summary table row
         SummaryTbrow = SummaryTbrow + 1
    
        'reset volume
        volume_total = 0
    
        'if immediate cells is the same ticker ...
        Else
    
        'add to the total volume
        volume_total = volume_total + ws.Cells(i, 7).Value
        
            
        End If
        
        'find max values and print them in cells(2,16) to cells (4,16)
        
            
        
        Next i
        
       

    
'set variale to 0
     Ticker = ""
     volume_total = 0




'nex worksheet
Next ws




End Sub






