# VBA-challenge
Sub Stockdata()


 ' data is in multiple sheets of the workbook
 For Each ws In Worksheets

 ' Identify where to put answers in sheet
  
  'Create column headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
 
  
' 1. multiple variables to be idnetifies; target (and where it begins from), i, LastRow,

  Dim TickerName As Integer
  TickerName = 2
  
  Dim i As Long
  
  
  ' need to assign rows as 2 different letters due to second question
  Dim x As Long
  x = 2
    
  Dim Lastrow As Long
  Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
   
   ' highlighting boxes based on pos or neg value
 
 
  ' begin writing funciton/question: looping in this case
  
  For i = 2 To Lastrow
  
   ' ticker change formula
   
  If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
  
  
            ' identify output location and value
             ws.Cells(TickerName, 9).Value = ws.Cells(i, 1).Value
  
  ' yearly change from openning to closing F2 - C2 (F = i and C = x)
  
    ws.Cells(TickerName, 10).Value = ws.Cells(i, 6).Value - ws.Cells(x, 3).Value
  
  
  
       ' highlight pos as green and neg as red
    
   If ws.Cells(TickerName, 10).Value > 0 Then
    
       ws.Cells(TickerName, 10).Interior.Color = vbGreen
       
       Else
       ws.Cells(TickerName, 10).Interior.Color = vbRed
       
       End If
  
  

  '% change from the opening price beginning to closing price end of year
   ' equation: (final value-initial value)/initial value
   ws.Cells(TickerName, 11).Value = ws.Cells(TickerName, 10).Value / ws.Cells(x, 3).Value
   
   ' need to have Range "K" in percent format
   ws.Cells(TickerName, 11).Value = FormatPercent(ws.Cells(TickerName, 11).Value)
  

       
   'calculate sum of stock volume located in Range "G"
   ws.Cells(TickerName, 12).Value = Application.WorksheetFunction.Sum(Range(ws.Cells(i, 7), ws.Cells(x, 7)))
  
    ' change to whole number format
    ws.Cells(TickerName, 12).NumberFormat = 0
  
  
  
  
   TickerName = TickerName + 1
    
    
    x = i + 1
 
  ' extract Ticker with "Greatest % increase", "Greatest % decrease", and "Greatest total volume"
  
  ' identify where to place answers

 ws.Range("O2").Value = "Greatest % increase"
  
 ws.Range("O3").Value = "Greatest % decrease"
 
 ws.Range("O4").Value = "Greatest total volume"
  
  End If
  
  Next i

  Next ws

 End Sub
 
