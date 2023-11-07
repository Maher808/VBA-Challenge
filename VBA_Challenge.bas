Attribute VB_Name = "Module1"
Sub Stock_analysis():

Dim ticker_name As String
Dim total_volume As Double
Dim summary_row As Integer
Dim yearly_change As Double
Dim last_row As Long
Dim last_row_2 As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim open_first_date As Double
Dim close_last_date As Double
Dim Percetage_Change As Double
Dim TheMax As Double
Dim TheMin As Double
Dim Max_volume As Double

For Each ws In Worksheets


'/Writing column headings
 ws.Range("P1") = "Ticker"
 ws.Range("I1") = "Ticker"
 ws.Range("Q1") = "Value"
 ws.Range("J1") = "Yearly Change"
 ws.Range("K1") = "Percentage Change"
 ws.Range("L1") = "Total Stock Volue"
 ws.Range("O2") = "Greatest % Increase"
 ws.Range("O3") = "Greatest % Decrease"
 ws.Range("O4") = "Greatest Total Volume"
 
 
 '/build ticker and volume columns
 total_volume = 0
 summary_row = 2

 last_row = ws.Range("A1").End(xlDown).Row
 
 '/start the loop to calculate Yearly change and persentage change
 For i = 2 To last_row
 
 If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
 total_volume = total_volume + ws.Cells(i, 7).Value
 
 Else
 total_volume = total_volume + ws.Cells(i, 7).Value
 
 ws.Cells(summary_row, 9).Value = ws.Cells(i, 1).Value
 ws.Cells(summary_row, 12).Value = total_volume
 
 total_volume = 0

 End If
 
 '\ calculating Yearly change and persentage change
  If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
  open_first_date = ws.Cells(i, 3)

 ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
 close_last_date = ws.Cells(i, 6)

yearly_change = ((close_last_date) - (open_first_date))
Percentage_Change = ((yearly_change) / (open_first_date))

ws.Cells(summary_row, 10).Value = yearly_change
ws.Cells(summary_row, 11).Value = Percentage_Change
summary_row = summary_row + 1

End If

'\ To color The cells

If ws.Cells(i, 10) > 0 Then
ws.Cells(i, 10).Interior.ColorIndex = 4

ElseIf ws.Cells(i, 10) < 0 Then
ws.Cells(i, 10).Interior.ColorIndex = 3
End If


 '\ percentage format
ws.Cells(i, 11).NumberFormat = "0.00%"
 Next i
 

 '\ calculating the Min & Max
 last_row_2 = ws.Range("K1").End(xlDown).Row
 
 TheMax = WorksheetFunction.Max(ws.Range("k2:k" & last_row_2))
 'MsgBox (TheMax)
 TheMin = WorksheetFunction.Min(ws.Range("k2:k" & last_row_2))
 'MsgBox (thein)
 Max_volume = WorksheetFunction.Max(ws.Range("L2:L" & last_row_2))
 'MsgBox (Max_volume)
 
ws.Range("Q2:Q3").NumberFormat = "0.00%"
   
'\ Printing the result
 For j = 2 To last_row_2
 
 If ws.Cells(j, 11) = TheMax Then
 ws.Range("P2") = ws.Cells(j, 1)
 ws.Range("Q2") = ws.Cells(j, 11)
 
 ElseIf ws.Cells(j, 11) = TheMin Then
 ws.Range("P3") = ws.Cells(j, 1)
 ws.Range("Q3") = ws.Cells(j, 11)
 
End If
Next j

For k = 2 To last_row_2
 
 If ws.Cells(k, 12) = Max_volume Then
 ws.Range("P4") = ws.Cells(k, 1)
 ws.Range("Q4") = ws.Cells(k, 12)
End If
Next
 
Next ws

End Sub












