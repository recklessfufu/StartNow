# StartNow
Public Sub CloseFile(fid As Integer)
    Close fid
End Sub

Public Function OpenFile(path As String) As Integer
'    Kill path
    Dim fid As Integer
    Set fileSystemObj = CreateObject("Scripting.FileSystemObject")
    If (fileSystemObj.FileExists(path) = True) Then
        Kill (path)
    End If
    
    fid = FreeFile()
    Open path For Binary As fid
    OpenFile = fid
End Function

Public Function WriteFile(fid As Integer, S As String) As Integer
    Dim bs() As Byte, buf(0 To 6) As Byte
    'Dim c As Integer, l As Integer, i As Integer, ch As Long
    Dim c As Long, l As Integer, i As Long, ch As Long
    bs = S
    c = UBound(bs)
    For i = 0 To c Step 2
        ch = bs(i + 1) * 2 ^ 8 Or bs(i)
        l = UnicodeToUTF8(ch, buf)
        Call WriteBytes(buf, l, fid)
    Next i
End Function

Private Sub WriteBytes(buf() As Byte, cnt As Integer, fid As Integer)
    Dim i As Integer
    Dim b(0 To 0) As Byte
    For i = 0 To cnt
        b(0) = buf(i)
        Put fid, , b
    Next i
End Sub
Private Function UnicodeToUTF8(ch As Long, buf() As Byte) As Integer
    Dim i As Integer: i = 0
    If (ch < &H80) Then
        buf(i) = ch
    ElseIf (ch < &H800) Then
        buf(i) = &HC0 Or ((ch And &H7C00) / 2 ^ 6): i = i + 1
        buf(i) = &H80 Or (ch And &H3F)
    ElseIf (ch < &H10000) Then
        buf(i) = &HE0 Or ((ch And &HF000) / 2 ^ 12): i = i + 1
        buf(i) = &H80 Or ((ch And &HFC0) / 2 ^ 6): i = i + 1
        buf(i) = &H80 Or (ch And &H3F)
    End If
    UnicodeToUTF8 = i
End Function
