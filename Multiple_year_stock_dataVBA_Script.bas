Attribute VB_Name = "Module1"
'Stock Analysis
 Sub Dosomething()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call RunCode
    Next
    Application.ScreenUpdating = True
End Sub

Sub RunCode()
'initialise variable

 Dim i As Long
Dim cc_row As Integer
Dim count As Integer
Dim stockvol As Long
Dim high As Long
Dim Last_row As Long
Dim j As Integer
Dim k As Integer
Dim l As Integer
Dim m As Integer

cc_row = 0
count = 0
Last_row = ActiveSheet.Cells(Rows.count, 1).End(xlUp).Row
MsgBox Last_row
Cells(cc_row + 2, 13).Value = 0
 
'Input the headers
With ActiveSheet
Range("J1").Value = "Ticker"
Range("K1").Value = "Yearly Change"
Range("L1").Value = "% Change"
Range("M1").Value = "Total Stock"
End With




'Loop through the rows to compare two adjacent ticker or stock names
For i = 2 To Last_row

'If the ticker names are the same add the stock value and place it in Total stock value,put the ticker name in a cell and
'the counter is used to get the number of same stock or ticker names found

   If Cells(i + 1, 1).Value = Cells(i, 1).Value Then
     Cells(cc_row + 2, 10).Value = Cells(i, 1).Value
     count = count + 1
     
     Cells(cc_row + 2, 13).Value = Cells(i, 7).Value + Cells(cc_row + 2, 13).Value
     
     'if the adjacent ticker is different,add the stock value of the first stock to the existing total stock value,finf the difference between
     'the final stock price of the stock and the start stockprice of the stock in a year and also its percentage change
     'and place the second stock name in a new cell and the corresponding total stock value is intialized to zer0 to start with.
     
     
    Else
     Cells(cc_row + 2, 13).Value = Cells(i, 7).Value + Cells(cc_row + 2, 13).Value
     Cells(cc_row + 2, 11).Value = Cells(i, 6).Value - Cells(i - count, 3).Value
    If Cells(i - count, 3).Value <> 0 And Cells(i, 6).Value <> 0 Then
     Cells(cc_row + 2, 12).Value = Cells(cc_row + 2, 11).Value / Cells(i - count, 3).Value
    ElseIf Cells(i - count, 3).Value = 0 Or Cells(i, 6).Value = 0 Then
     Cells(cc_row + 2, 12).Value = Cells(cc_row + 2, 11).Value
     End If

' row is increased by 1 so as to get the next stock to be compared

  cc_row = cc_row + 1
   Cells(cc_row + 2, 13).Value = 0
  count = 0
  End If
  

Next i
'Close the loop
MsgBox ("cc_row is" + Str(cc_row))

 ' change the change to percentage format
 With ActiveSheet
     Columns("L").NumberFormat = "0.00%"
End With

' find the greatest % change and corresponding ticker name to display it in in 2,16 and 2,17
'the first stock value is used to compare to start with.
Cells(2, 17).Value = Cells(2, 12).Value

Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % increase"
Cells(3, 15).Value = "Smallest % increase"
Cells(4, 15).Value = "Greatest stock volume"
'loop through to compare with each of the rows

For j = 3 To cc_row + 1
'when the first value is less than the compared value, the compared value replaces the the first value

    If Cells(2, 17).Value < Cells(j, 12).Value Then
       Cells(2, 16).Value = Cells(j, 10).Value
       Cells(2, 17).Value = Cells(j, 12).Value
'If the first value is greater no change
    Else
'row is increased by one every time to compare the next value
    'cc_row = cc_row +1

    End If

Next j

'Loop through the rows to get the smallest % change and the corresponding ticker name in the same way as the earlier loop
'and display it
Cells(3, 17).Value = Cells(2, 12).Value

For k = 3 To cc_row + 1
    If Cells(3, 17).Value > Cells(k, 12).Value Then
       Cells(3, 16).Value = Cells(k, 10).Value
       Cells(3, 17).Value = Cells(k, 12).Value

    Else

    'cc_row = cc_row + 1

     End If

Next k

'change to percentage format
With ActiveSheet
  Range("Q2:Q3").NumberFormat = "0.00% "
End With


Cells(4, 17).Value = Cells(2, 13).Value

'Loop through the rows to get the highest total stock volume and the corresponding ticker name as the earlier loop
For l = 3 To cc_row + 1
     If Cells(4, 17).Value < Cells(l, 13).Value Then
         Cells(4, 16).Value = Cells(l, 10).Value
        Cells(4, 17).Value = Cells(l, 13).Value

     Else
     
     'cc_row = cc_row + 1
     End If

Next l

'For positive %change values color the cells green and for negative values color red
For m = 2 To cc_row + 1
    Set r = ActiveSheet.Cells(m, "K") ' starts at K2
    v = r.Value

    If v >= 0 Then
        r.Interior.ColorIndex = 4
        ElseIf v < 0 Then
        r.Interior.ColorIndex = 3
        
   'If ((ActiveCell.Value) >= 0) Then
       
       ' ActiveCell.Interior.Color = 4
    'ElseIf ((ActiveCell.Value) < 0) Then
        'ActiveCell.Interior.Color = 3
    End If
    
 Next m

End Sub

Sub clear()
Dim Answer As VbMsgBoxResult
With ActiveSheet
  Answer = MsgBox(" are you sure", vbYesNo + vbQuestion, " Clear Cells")
  
  If Answer = vbYes Then
  
  ActiveSheet.Range("J2:M291").clear
 ActiveSheet.Range("O1: Q4").clear
  Else
  
  
  Exit Sub
  End If
  End With
  
End Sub
