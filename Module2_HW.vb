Sub tickertotaler_moderate()


For Each ws In ThisWorkbook.Worksheets

ws.Activate

'define everything
'Dim ws As Worksheet
Dim ticker As String
Dim Vol As Double
Vol = 0
Dim year_open As Double
Dim year_close As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim Summary_Table_Row As Integer


'this prevents my overflow error
'On Error Resume Next

'run through each worksheet
'For Each ws In ThisWorkbook.Worksheets
    'set headers
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    'setup integers for loop
    Summary_Table_Row = 2
    
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    'loop
        For i = 2 To lastrow
        
        
            If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
            
            
            'find all the values
            ticker = Cells(i, 1).Value
            Cells(Summary_Table_Row, 9).Value = ticker
            
            
            
            year_open = Cells(i, 3).Value
            'Cells(Summary_Table_Row, 14).Value = year_open
            
            
            ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            
            year_close = Cells(i, 6).Value
            'Cells(Summary_Table_Row, 15).Value = year_close
                        
            yearly_change = year_close - year_open
            percent_change = yearly_change / year_open
            Cells(Summary_Table_Row, 11).Value = percent_change
            
            Cells(Summary_Table_Row, 10).Value = yearly_change
            Summary_Table_Row = Summary_Table_Row + 1
            
           
           
             
             
                          Vol = 0
         
         Else
         
         
            
         

        End If
        
        
    'finish loop
        Next i
       Summary_Table_Row = 2
       
      For D = 2 To lastrow
        
        
            If Cells(D + 1, 1).Value <> Cells(D, 1).Value Then
               
            Vol = Vol + Cells(D, 7).Value
            
            Cells(Summary_Table_Row, 12).Value = Vol
               
               
               
            Summary_Table_Row = Summary_Table_Row + 1
               
              Vol = 0
               
         End If
         Vol = Vol + Cells(D, 7).Value
         
         
         Next D
               
               
               
               
               
    Columns("K").NumberFormat = "0.00%"
    
    
    Dim SecondTable As Long
    
    SecondTable = Cells(Rows.Count, "I").End(xlUp).Row
    
        For F = 2 To SecondTable
        
            If Cells(F, 10).Value <= 0 Then
                Cells(F, 10).Interior.ColorIndex = 3
            Else
                Cells(F, 10).Interior.ColorIndex = 4
            End If
        Next F
        
   Summary_Table_Row = 2
   
    Dim Maximum As Double
    Dim Minimum As Double
    Dim VolMax As Double
    Maximum = WorksheetFunction.Max(Range("K2:K" & SecondTable))
       Cells(Summary_Table_Row, 17).Value = Maximum
        
    Minimum = WorksheetFunction.Min(Range("K2:K" & SecondTable))
       Cells(Summary_Table_Row + 1, 17).Value = Minimum
        
    VolMax = WorksheetFunction.Max(Range("L2:L" & SecondTable))
       Cells(Summary_Table_Row + 2, 17).Value = VolMax
            
    Range("Q2:Q3").NumberFormat = "0.00%"
        
        
  Cells(2, 15).Value = "Greatest % Increase"
  Cells(3, 15).Value = "Greatest % Decrease"
  Cells(4, 15).Value = "Greatest Total Volume"
  Cells(1, 16).Value = "Ticker"
  Cells(1, 17).Value = "Value"
       
       
  'loop
  For M = 2 To SecondTable
       
    If Maximum = Cells(M, 11).Value Then
    
      TickerTwo = Cells(M, 9).Value
      
      Cells(2, 16).Value = TickerTwo
      
    'Request
    ElseIf Minimum = Cells(M, 11).Value Then
    'Answer
    TickerTwo = Cells(M, 9).Value
    'Show me the Answer
    Cells(3, 16).Value = TickerTwo
    
    
    ElseIf VolMax = Cells(M, 12).Value Then
    
    TickerTwo = Cells(M, 9).Value
    
    Cells(4, 16).Value = TickerTwo
    
    End If
    
    Next M
    
    
    
    
       
       
       
       
       
    'move to next worksheet
Next ws
            
            
End Sub

