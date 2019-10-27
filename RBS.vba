Sub fixdata()
Dim ender, jender, iender, datasheet, metalsheet, insurersheet, var1, var2, var3, var4
Set datasheet = Worksheets("healthplan")
Set metalsheet = Worksheets("metal")
Set insurersheet = Worksheets("insurer")
datasheet.Activate
ender = datasheet.Cells(1, "A").End(xlDown).Row
jender = insurersheet.Cells(1, "A").End(xlDown).Row
iender = metalsheet.Cells(1, "A").End(xlDown).Row
For i = 1 To ender
    For j = 1 To jender
        If InStr(1, datasheet.Cells(i, "E").Value, insurersheet.Cells(j, "C"), vbTextCompare) Or _
        InStr(1, insurersheet.Cells(j, "C"), datasheet.Cells(i, "E").Value, vbTextCompare) Then
            datasheet.Cells(i, "E").Value = insurersheet.Cells(j, "A")
        End If
    Next j
    For j = 1 To iender
            var1 = datasheet.Cells(i, "F").Value
            var2 = metalsheet.Cells(j, "B")
        If InStr(1, datasheet.Cells(i, "F").Value, metalsheet.Cells(j, "B"), vbTextCompare) Or _
        InStr(1, metalsheet.Cells(j, "B"), datasheet.Cells(i, "F").Value, vbTextCompare) Then
            datasheet.Cells(i, "F").Value = insurersheet.Cells(j, "A")
        End If
    Next j
Next i


End Sub
