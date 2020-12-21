Attribute VB_Name = "Module1"
Sub wallstreet_stock()
'set ws as variable for code to loop through every worksheets
Dim ws As Worksheet
  
'loop through every worksheet
For Each ws In Worksheets


'set variable to hold ticker symble
  Dim Ticker As String
 
 'set variable to keep track of the ticker row in summary
  Dim rowcount As Integer
  rowcount = 2
 

 'set variable to hold stock volume
  Dim Stock_total As Double
  Stock_total = 0
 
 'set variable to hold yearly change
  Dim open_stock As Double
  open_stock = 0
  
  Dim close_stock As Double
  close_stock = 0
  
  Dim yearly_change As Double
  yearly_change = 0
  
  Dim percent_change As Double
  percent_change = 0
 
 'set last row for the loop
  Dim lastrow As Long
  lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
'Print header on the summary column
  ws.Cells(1, 9).Value = "Ticker"
  ws.Cells(1, 10).Value = "Yearly Change"
  ws.Cells(1, 11).Value = "Percent Change"
  ws.Cells(1, 12).Value = "Total Stock Volume"
  
'format the column to fit
 ws.Columns("I:L").EntireColumn.AutoFit
  
 'loop though ticker symbols
   For i = 2 To lastrow
   
    'set conditions sorting stocks by ticker symbols
    'set open stock
     If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
   
        open_stock = ws.Cells(i, 3).Value
     End If
   
    'define stock total
     Stock_total = Stock_total + ws.Cells(i, 7).Value
     
    'sort out stock by ticker symbol
     If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
     
       'print ticker to summary table
        ws.Cells(rowcount, 9).Value = ws.Cells(i, 1).Value
     
       'print total to summary
        ws.Cells(rowcount, 12).Value = Stock_total
     
       'set year close price
        close_stock = ws.Cells(i, 6).Value
      
       'work out year difference
        yearly_change = close_stock - open_stock
     
        ws.Cells(rowcount, 10).Value = yearly_change
        
     
       'set the color changing with the yearly change
        If yearly_change >= 0 Then
          ws.Cells(rowcount, 10).Interior.ColorIndex = 4
          
        Else
          ws.Cells(rowcount, 10).Interior.ColorIndex = 3
          
        End If
       
      'calculate percentage change
      'zero can not be divided and hence needs to be excluedded to avoid error
        If open_stock = 0 Then
        
           percent_change = 0
        
        Else
           percent_change = yearly_change / open_stock
        
     
       'print percent change to summary
        ws.Cells(rowcount, 11).Value = percent_change
       
       'format the column for comparison on the challenge question
        ws.Cells(rowcount, 11).NumberFormat = "0.00%"
      
        End If
  
   
    'go to next new ticker
    rowcount = rowcount + 1
    
    'reset open_stock
    open_stock = 0
    
    'rest close stock
    close_stock = 0
    
    'reset stock total
    Stock_total = 0
    
    'reset yearly change
    yearly_change = 0
    
    'rest precent change
    percent_change = 0
    
    
    End If
  
  
 Next i
  
 'new loop for the chellange part home work
 For i = 2 To last_summary_row
 'greatest increase and decrease table
  ws.Cells(2, 15).Value = "Greatest % Increase"
  ws.Cells(3, 15).Value = "Greatest % Decrease"
  ws.Cells(4, 15).Value = "Greatest Total Volume"
  ws.Cells(1, 16).Value = "Ticker"
  ws.Cells(1, 17).Value = "Value"
  
 'locate the last row on the summary form
  last_summary_row = ws.Cells(Rows.Count, 9).End(xlUp).Row

 'the largest number in percent change is the greatest in increase and the lowest number is the greatest in decrease.
 'excel function to be used to workout the answer and bring the corresponding ticker sybol the summary
 'loop through the summary table

  

  'find out the greatest in increase
   If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(Range("K2:K" & last_summary_row)) Then
        ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
        ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
        ws.Cells(2, 17).NumberFormat = "0.00%"
   
  'find ou the greatest in decrease which is the smallest number in percent change
   ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(Range("K2:K" & last_summary_row)) Then
        ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
        ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
        ws.Cells(3, 17).NumberFormat = "0.00%"
        
 'find out the highest volume using excel function
  ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(Range("L2:L" & last_summary_row)) Then
         ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
         ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
        
  End If
 
        
 
  Next i

 
'format column to fit
 ws.Columns("O:Q").EntireColumn.AutoFit


Next ws



End Sub
