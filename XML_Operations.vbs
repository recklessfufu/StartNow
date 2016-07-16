Public Function XML_APPEND(old As String, data As String) As String
    XML_APPEND = old + data + Chr(13) + Chr(10)
End Function

Public Function XML_Tag(old As String, TagName As String) As String
    XML_Tag = "<" + TagName + ">" + old + "</" + TagName + ">" + Chr(13) + Chr(10)
End Function

Public Function checkXMLSpecialCharacter(text As String) As String    '×?·?×aò?
    Dim tmp As String
    If (InStr(text, "&") + InStr(text, ">") + InStr(text, "<")) > 0 Then
        tmp = Replace(text, "&", "&amp;")
        tmp = Replace(tmp, "", "&quot;")
        tmp = Replace(tmp, ">", "&gt;")
        tmp = Replace(tmp, "<", "&lt;")
        checkXMLSpecialCharacter = tmp
    Else
        checkXMLSpecialCharacter = text
    End If
    
End Function
