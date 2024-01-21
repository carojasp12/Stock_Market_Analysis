Sub Multiple_ys()

' Loop through all sheets
For Each ws In Worksheets

' Set an initial variables for holding the ticker symbol
Dim ticker As String

' Set an initial variables for holding the last row count
Dim LR As Long

' Set an initial variables for holding position to print the results
Dim row As Integer

' Set an initial variables for holding yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
Dim yearly_change As Double

' Set an initial variables for holding the percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
Dim percentage_change As String

' Set an initial variables for holding the total stock volume of the stock
Dim ts_volume As LongLong

' Set an initial variables for holding the stock with the greatest % increase
Dim max_change As Double

' Set an initial variables for holding the stock with the greatest % decrease
Dim min_change As Double

' Set an initial variables for holding the stock with the greatest total volume
Dim Gt_volume As LongLong

' look up for the last row of ticker column
LR = ws.Cells(Rows.Count, 1).End(xlUp).row

row = 2
j = 2

    
' loop through all
For i = 2 To LR

    ' Check if we are still within the ticker symbol
    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                         
      'set the ticker symbol
      ticker = ws.Cells(i, 1).Value
      
      'Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
      yearly_change = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
      
      ' sum the total stock volume of each year
      ts_volume = ts_volume + ws.Cells(i, 7).Value
                 
      'print the ticker symbol in the table
      ws.Range("I" & row).Value = ticker
      
      'print the yearly change in the table
      ws.Range("J" & row).Value = yearly_change
      
        ' Check if the yearly change is positive or negative
        If yearly_change > 0 Then
            
            ' print a highlight positive change in green
            ws.Range("J" & row).Interior.ColorIndex = 4
            
        Else
            ' print a highlight negative change in red
            ws.Range("J" & row).Interior.ColorIndex = 3
        
        End If
            
      
      'print the total stock volume
      ws.Range("l" & row).Value = ts_volume
      
      ' reset the total volume
      ts_volume = 0
      
      
      ' check for any zero values in the opening price at the beginning of the year to prevent a zero denominator
        If ws.Cells(j, 3).Value <> 0 Then
      
        'calculate the percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
        percentage_change = yearly_change / ws.Cells(j, 3).Value
        
            
        'print the percentage change in the table
        ws.Range("k" & row).Value = Format(percentage_change, "percent")
       
        End If
                    
      ' add one to the table
      row = row + 1
      
      ' reset the sum of opening and closing price
      yearly_change = 0
        
      ' reset the position of opening and closing price for a new ticker symbol
      j = i + 1
      
      Else
      ' sum the total stock volume of each year
      ts_volume = ts_volume + ws.Cells(i, 7).Value
     
     End If
     
      

Next i


' print the names of each new tittle in the new tables created with the results
ws.Range("I1") = "Ticker"
ws.Range("J1") = "Yearly Change"
ws.Range("K1") = "Percentage Change"
ws.Range("L1") = "Total Stock Volume"
ws.Range("L1") = "Total Stock Volume"
ws.Range("N2") = "Greatest % Increase"
ws.Range("N3") = "Greatest % Decrease"
ws.Range("N4") = "Greatest Total Volume"
ws.Range("O1") = "Ticker"
ws.Range("P1") = "Value"



' look for and print the stock with the greatest % increase
max_change = WorksheetFunction.Max(ws.Range("K:K"))
ws.Cells(2, 16).Value = Format(max_change, "percent")

' look for and print the stock with the greatest % decrease
min_change = WorksheetFunction.Min(ws.Range("K:K"))
ws.Cells(3, 16).Value = Format(min_change, "percent")

' look for and print the stock with the greatest total volume
Gt_volume = WorksheetFunction.Max(ws.Range("L:L"))
ws.Cells(4, 16).Value = Gt_volume

' look for and print the ticker symbol of the stock with the greatest % increase
ws.Range("O2").Value = "=xlookup(P2,K:K,I:I)"

' look for and print the ticker symbol of the stock with the greatest % decrease
ws.Range("O3").Value = "=xlookup(P3,K:K,I:I)"

' look for and print the ticker symbol of the stock with the greatest total volume
ws.Range("O4").Value = "=xlookup(P4,L:L,I:I)"

'fit the colums for the results size
ws.Columns("I:P").AutoFit


Next ws

End Sub



