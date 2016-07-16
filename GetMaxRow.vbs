Public Function GetMaxRow(ByVal sht As Worksheet, ByVal nMaxColumn As Integer) As Long
    Dim i As Long, j As Integer
    Dim sCellValue As String
    Dim lMaxRow As Long

    lMaxRow = sht.UsedRange.Rows.Count
    For i = lMaxRow To 1 Step -1
        For j = 1 To nMaxColumn
            sCellValue = sht.Cells(i, j).Value
            sCellValue = Trim$(sCellValue)
            If sCellValue <> "" Then Exit For
        Next
        If j <= nMaxColumn Then
            Exit For
        End If
    Next
    lMaxRow = i
    GetMaxRow = lMaxRow
End Function
