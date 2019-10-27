Sub fixdata()
Dim ender, jender, iender, datasheet, metalsheet, insurersheet
Set datasheet = Worksheets(1)
Set metalsheet = Worksheets(2)
Set insurersheet = Worksheets(3)
datasheet.Activate
ender = datasheet.Cells(1, "A").End(xlDown).Row
jender = insurersheet.Cells(1, "A").End(xlDown).Row
iender = metalsheet.Cells(1, "A").End(xlDown).Row
For i = 1 To ender
    For j = 1 To jender
        If InStr(1, datasheet.Cells(i, "E").Value, insurersheet.Cells(j, "C")) Then
            datasheet.Cells(i, "E").Value = insurersheet.Cells(j, "A")
        End If
    Next j
    For j = 1 To iender
        If InStr(1, datasheet.Cells(i, "F").Value, metalsheet.Cells(j, "C")) Then
            datasheet.Cells(i, "F").Value = insurersheet.Cells(j, "C")
        End If
    Next j
Next i


End Sub
