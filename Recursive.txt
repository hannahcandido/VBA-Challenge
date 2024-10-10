Sub Ticker()

Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
ws.Activate  ' Activate the worksheet

'ticker, openvalue, closevalue, Quaterly_change, Percent_change and Total_Stock_volume name

Dim Ticker As String
Dim Openvalue As Double
Dim closevalue As Double
Dim Quarterly_Change As Double
Dim Percent_Change As Double
Dim Total_Stock_volume As Double


Quaterly_Change = 0
Openvalue = 0
closevalue = 0

Total_Stock_volume = 0

Dim summary As Integer
summary = 2

'determine the last row
Dim lastrow As Double
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'setting the loop
 For i = 2 To lastrow
 
 'check if its the same Ticker symbol
 If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
 Ticker = Cells(i, 1).Value
 Openvalue = Cells(i - 61, 3).Value
 closevalue = Cells(i, 6).Value
 
 'Display Ticker name
 Range("i" & summary).Value = Ticker
 
 'Display open amount
 
 Range("j" & summary).Value = Openvalue
 
 'Display close amount
 
 Range("k" & summary).Value = closevalue
 
 Quarterly_Change = closevalue - Openvalue
 
 Range("l" & summary).Value = Quarterly_Change
 
 Percent_Change = Quarterly_Change / Openvalue

  Range("m" & summary).Value = Percent_Change
 
 Total_Stock_volume = Total_Stock_volume + Cells(i, 7).Value
 Range("N" & summary).Value = Total_Stock_volume
 
 summary = summary + 1
 
 Openvalue = 0
 
 closevalue = 0
 Quarterly_Change = 0
 Percent_Change = 0
 Total_Stock_volume = 0
 
 Else
 
 Openvalue = Cells(i + 61, 3).Value
 closevalue = Cells(i + 61, 6).Value

 Total_Stock_volume = Total_Stock_volume + Cells(i, 7).Value
 
 
 End If
 
 Next i
 
  
 
 For i = 2 To lastrow
 If Cells(i, 12).Value > 0 Then
 
 Cells(i, 12).Interior.ColorIndex = 4
 
 ElseIf (Cells(i, 12).Value < O) Then
 
 Cells(i, 12).Interior.ColorIndex = 3
 
 Else
 
  Cells(i, 12).Interior.ColorIndex = 0
 
 
 End If
 
Next i

 For i = 2 To lastrow
 If Cells(i, 13).Value > 0 Then
 
 Cells(i, 13).Interior.ColorIndex = 4
 
 ElseIf (Cells(i, 13).Value < O) Then
 
 Cells(i, 13).Interior.ColorIndex = 3
 
 Else
 
  Cells(i, 13).Interior.ColorIndex = 0
 
 
 End If
 
Next i

Dim Greatest_Total_volume As Double
Dim Greatest_percent_increase As Double
Dim Greatest_percent_Decrease As Double

Greatest_percent_increase = Application.WorksheetFunction.Max(Range("M:M").Value)
Cells(2, 20).Value = Greatest_percent_increase
Cells(2, 20).NumberFormat = "#0.00%"

Greatest_percent_Decrease = Application.WorksheetFunction.Min(Range("M:M").Value)
Cells(3, 20).Value = Greatest_percent_Decrease
Cells(3, 20).NumberFormat = "#0.00%"

Greatest_Total_volume = Application.WorksheetFunction.Max(Range("N:N").Value)
Cells(4, 20).Value = Greatest_Total_volume


Greatest_percent_increase = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("M2:M" & lastrow)), ws.Range("M2:M" & lastrow), 0)
Greatest_percent_Decrease = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("M2:M" & lastrow)), ws.Range("M2:M" & lastrow), 0)
Greatest_Total_volume = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("N2:N" & lastrow)), ws.Range("N2:N" & lastrow), 0)

ws.Range("S2") = ws.Cells(Greatest_percent_increase + 1, 9)
ws.Range("S3") = ws.Cells(Greatest_percent_Decrease + 1, 9)
ws.Range("S4") = ws.Cells(Greatest_Total_volume + 1, 9)


Next ws

End Sub