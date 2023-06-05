Sub ExtractPercentage()
    Dim rng As Range
    Dim cell As Range
    Dim start As Integer, length As Integer

    ' Set your range here
    Set rng = ThisWorkbook.Sheets("Sheet1").Range("B2:C10")

    For Each cell In rng
        If InStr(cell.Value, "(") > 0 Then
            start = InStr(cell.Value, "(") + 1
            length = InStr(cell.Value, "%") - start
            cell.Value = Mid(cell.Value, start, length) & "%"
        End If
    Next cell
End Sub
