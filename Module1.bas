Attribute VB_Name = "Module1"

Sub stock_calculation()
Dim ws As Worksheet
For Each ws In ActiveWorkbook.Worksheets
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stcok Volume"
ws.Range("O2").Value = "Greatest % Incr"
ws.Range("O3").Value = "Greatest % Decr"
ws.Range("O4").Value = "Greatest Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

    Dim ticker As String
    Dim output_index, last_row, i As Long
    Dim stock_volume, yearly_change, percent_change, open_price, close_price As Double
    
   
    output_index = 2
    stock_volume = 0
    yearly_change = 0
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    j = 2
    
 For i = 2 To last_row
 stock_volume = stock_volume + ws.Cells(i, 7).Value
 open_price = ws.Cells(j, 3).Value
 
  If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    ticker = ws.Cells(i, 1).Value
     close_price = ws.Cells(i, 6).Value
     yearly_change = close_price - open_price
    If open_price <> 0 Then
     percent_change = (yearly_change / open_price) * 100
    'Else:
     'ws.Cells(output_index, 11).Value = "NULL"
    End If
    ws.Cells(output_index, 9).Value = ticker
    ws.Cells(output_index, 10).Value = yearly_change
    If ws.Cells(output_index, 10) < 0 Then
     ws.Cells(output_index, 10).Interior.ColorIndex = 3
      Else
     ws.Cells(output_index, 10).Interior.ColorIndex = 4
      End If
    ws.Cells(output_index, 11).Value = Format(percent_change, "0.00\%")
    ws.Cells(output_index, 12).Value = stock_volume
    
    output_index = output_index + 1
    stock_volume = 0
    yearly_change = 0
    percent_change = 0
    j = i + 1

    End If
     
 Next i
  
  last_row = ws.Cells(Rows.Count, 9).End(xlUp).Row
  Maximum = ws.Cells(2, 11).Value
  Minimum = ws.Cells(2, 11).Value
  Maximum_Index = 2
  Minimum_Index = 2
  Maximum_volume = ws.Cells(2, 12).Value
  Maximum_volIndex = 2
  
For i = 2 To last_row

If ws.Cells(i + 1, 11).Value > Maximum Then
 Maximum = ws.Cells(i + 1, 11).Value
 Maximum_Index = i + 1
 
 ElseIf ws.Cells(i + 1, 11).Value < Minimum Then
 Minimum = ws.Cells(i + 1, 11).Value
 Minimum_Index = i + 1
 
 End If
 If ws.Cells(i + 1, 12).Value > Maximum_volume Then
 Maximum_volume = ws.Cells(i + 1, 12).Value
 Maximum_volIndex = i + 1
 End If
   Next i
 
  ws.Cells(2, 16).Value = ws.Cells(Maximum_Index, 9).Value
  ws.Cells(3, 16).Value = ws.Cells(Minimum_Index, 9).Value
  ws.Cells(2, 17).Value = Maximum
  ws.Cells(3, 17).Value = Minimum
  ws.Cells(2, 17).Value = Format(Maximum, "Percent")
  ws.Cells(3, 17).Value = Format(Minimum, "Percent")
   
   ws.Cells(4, 17).Value = Maximum_volume
   ws.Cells(4, 16).Value = ws.Cells(Maximum_volIndex, 9).Value
   
  
  Next ws
   
End Sub


