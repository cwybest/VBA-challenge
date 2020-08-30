Sub stock():
Dim ws As Worksheet
For Each ws In Worksheets
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
lastcolumn = Cells(1, Columns.Count).End(xlToLeft).Column
Dim open_price As Double
Dim ticker As String
Dim counter As Integer
Dim volume As Double
Dim max As Double
Dim min As Double
Dim greatest As Double
Dim percentage_range As Range
open_price = Cells(2, 3)
counter = 2
volume = 0
Range("i1") = "ticker"
Range("j1") = "yearly change"
Range("k1") = "Percentage Change"
Range("l1") = "Total Stock Volume"
Range("p1") = "ticker"
Range("q1") = "Value"
Range("o2") = "Greatest % increase"
Range("o3") = "Greatest % decrease"
Range("o4") = "Greatest Total Volume"
For i = 2 To lastrow
    If Cells(i, 1) = Cells(i + 1, 1) Then
        volume = volume + Cells(i, 7)
    Else
        ticker = Cells(i, 1)
        Row = Cells(i, 6).Row
        row2 = Cells(i + 1, 3).Row
        closed_price = Cells(Row, 6)
        Range("i" & counter) = ticker
        volume = volume + Cells(i, 7)
        Range("l" & counter) = volume
        Range("j" & counter) = closed_price - open_price
        open_price = Cells(row2, 3)
        counter = counter + 1
        volume = 0
        Range("j" & counter) = closed_price - open_price
    End If
Next i
counter = 2

lastrow2 = Cells(Rows.Count, 10).End(xlUp).Row
Range("j" & lastrow2).Delete
lastrow2 = Cells(Rows.Count, 10).End(xlUp).Row
For i = 1 To lastrow - 1
open_price = Cells(2, 3)
    If Cells(i, 1) <> Cells(i + 1, 1) And open_price <> 0 Then
        row3 = Cells(i + 1, 3).Row
        open_price = Cells(row3, 3)
        Range("k" & counter) = Range("j" & counter) / open_price
        counter = counter + 1
    End If
    If Cells(i, 1) <> Cells(i + 1, 1) And open_price = 0 Then
        Range("k" & counter) = Range("j" & counter) / 1
        End If
Next i
Range("k2:k" & lastrow2).Style = "Percent"
For i = 2 To lastrow2
If Range("k" & i).Value < 0 Then
    Range("k" & i).Interior.ColorIndex = 3
ElseIf Range("k" & i).Value < 0 Then
    Range("k" & i).Interior.ColorIndex = 1
Else
    Range("k" & i).Interior.ColorIndex = 4
End If
Next i
Set percentage_range = Range("k1:k" & lastrow2)
max = WorksheetFunction.max(percentage_range)
min = WorksheetFunction.min(percentage_range)
volume_range = Range("l1").EntireColumn
greatest = WorksheetFunction.max(volume_range)
Range("q2") = max
Range("q3") = min
Range("q4") = greatest
Range("q2").Style = "Percent"
Range("q3").Style = "Percent"
For i = 2 To lastrow2
If max = Range("k" & i) Then
    Range("p2") = Range("I" & i)
ElseIf min = Range("k" & i) Then
    Range("p3") = Range("I" & i)
End If
If greatest = Range("l" & i) Then
    Range("p4") = Range("I" & i)
End If
Next i
Next
End Sub







