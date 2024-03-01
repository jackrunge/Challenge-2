Attribute VB_Name = "Module1"
Sub Challenge2()
Dim first_opening As Double
Dim last_closing As Double
Dim ticker As String
Dim tracker As Integer
Dim volume As Double
Dim max As Double
Dim max_row As Integer
Dim min As Double
Dim min_row As Integer
Dim vol_max As Double
Dim vol_row As Integer
Dim rg As Range
Dim rg2 As Range


For Each ws In Worksheets
ws.Activate

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "($)Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"

tracker = 2
ticker = ""
volume = 0
max = 0



LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 1 To LastRow

    If ticker = Cells(i + 1, 1).Value And Cells(i + 1, 1).Value = Cells(i + 2, 1).Value Then
    volume = volume + Cells(i + 1, 7).Value
    
    ElseIf Cells(i + 1, 1).Value = Cells(i + 2, 1).Value Then
    ticker = Cells(i + 1, 1).Value
    first_opening = Cells(i + 1, 3).Value
    volume = volume + Cells(i + 1, 7).Value
    
    Else
    last_closing = Cells(i + 1, 6).Value
    Cells(tracker, 10).Value = last_closing - first_opening
    Cells(tracker, 11).Value = Cells(tracker, 10) / first_opening
    Cells(tracker, 12).Value = volume
    Cells(tracker, 9).Value = ticker
    volume = 0
    first_opening = 0
    last_closing = 0
    tracker = tracker + 1
    
    End If
    
    Next i
    
    Last_Entry = ws.Cells(Rows.Count, 11).End(xlUp).Row

 For Each Cell In Range("K2:K" & Last_Entry)
 If Cell.Value > max Then
 max = Cell.Value
 max_row = Cell.Row
 
 End If
 If Cell.Value < min Then
 min = Cell.Value
 min_row = Cell.Row
 End If
 Next Cell
 
 For Each Cell In Range("L2:L" & Last_Entry)
 If Cell.Value > vol_max Then
 vol_max = Cell.Value
 vol_row = Cell.Row
 End If
 Next Cell
 
 Cells(2, 17).Value = max
 Cells(3, 17).Value = min
 Cells(4, 17).Value = vol_max
 Cells(2, 16).Value = Cells(max_row, 9).Value
 Cells(3, 16).Value = Cells(min_row, 9).Value
 Cells(4, 16).Value = Cells(vol_row, 9).Value
 
 ws.Range("K:K").Style = "Percent"
 ws.Range("Q2:Q3").Style = "Percent"

 
 
 Set rg = Range("J2:J" & Last_Entry)
 Set rg2 = Range("K2:K" & Last_Entry)
 

 
 With rg.FormatConditions
 .Delete
  .Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
    .Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
 End With
 
 rg.FormatConditions(1).Interior.Color = vbGreen
 rg.FormatConditions(2).Interior.Color = vbRed
 
  
 With rg2.FormatConditions
 .Delete
  .Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
    .Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
 End With
 
 rg2.FormatConditions(1).Interior.Color = vbGreen
 rg2.FormatConditions(2).Interior.Color = vbRed

Next ws
End Sub


