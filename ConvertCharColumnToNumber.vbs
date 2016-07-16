Public Function ConvertCharColumnToNumber(col As Integer) As String
    If col < 0 Or col > 256 Then
        ConvertCharColumnToNumber = "Invalid Input"
        Exit Function
    End If
    Dim t_col As String
    If col < 27 And col > 0 Then
        t_col = Chr(64 + col)
    End If
    If col > 26 Then
        t_col = Chr(64 + (col \ 26)) & Chr((col Mod 26) + 64)
    End If
    ConvertCharColumnToNumber = t_col
End Function
