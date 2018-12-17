'first bit is from mailchimpvba Rental list builder
'loops through the whole column for each filled cell in column, then fills in the column for each duplicate entry in column
Sub bookaggregate()
Dim first As Integer, last As Integer, blankcol As Integer
first = Cells(2, "A").Row
last = Cells(2, "A").End(xlDown).Row
For i = 1 To last
    For j = 1 To last
        If (Cells(i, "A").Value = Cells(j, "A")) Then
            blankcol = Cells(i, "A").End(xlToRight).Column + 1
            Cells(i, blankcol) = Cells(j, "C").Value
        End If
    Next j
Next i
End Sub


'to clean up reports with junk inbetween data blocks
Sub prepdatlatefees()
Dim last As Integer, blankcol As Integer, selex As String
last = Cells(2, "A").End(xlDown).Row
For i = 1 To last
    If InStr(1, Cells(i, "A").Value, "Report Run") Then
        selex = i & ":" & (i + 2)
        Rows(selex).Select
        Selection.Delete Shift:=xlUp
    End If
Next i
End Sub
