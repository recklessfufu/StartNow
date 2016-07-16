Public Function CreateDitamap()
'（5）根据Result页签，生成评审用的Ditamap文件。
    Dim folderName As String
    Dim tableID As String
    Dim resultPageName As String
    Dim fileName As String
    Dim language As String
    
    '从配置文件页签获取信息
    folderName = Worksheets("配置文件").Cells(7, 2)
    resultPageName = Worksheets("配置文件").Cells(5, 2)
    language = Worksheets("配置文件").Cells(6, 2)

    '设置表单
    Set resultSheet = Worksheets(resultPageName)
    
    If language = "英文" Then
        CreateDir (Application.ActiveWorkbook.path & "\EN\" & folderName & "\")
    ElseIf language = "中文" Then
        CreateDir (Application.ActiveWorkbook.path & "\CN\" & folderName & "\")
    Else
        MsgBox "请先在配置文件页签设置语言类型!"
        Exit Function
    End If
    
    Dim lastrow As Integer
    lastrow = resultSheet.Range("B65536").End(xlUp).Row

    Dim sEnXmlContent As String
    Dim sCommonXmlContent As String

    Dim sEnFileContent As String
    sEnXmlContent = XML_APPEND(sEnXmlContent, "<?xml version=""1.0"" encoding=""UTF-8""?>")
    sEnXmlContent = XML_APPEND(sEnXmlContent, "<!DOCTYPE map PUBLIC ""-//OASIS//DTD DITA Map//EN"" ""map.dtd"">")
    
    If language = "英文" Then
        sEnXmlContent = XML_APPEND(sEnXmlContent, "<map id=""map"" title="" WebService Error Codes"" xml:lang=""en-us"">")
    ElseIf language = "中文" Then
        sEnXmlContent = XML_APPEND(sEnXmlContent, "<map id=""map"" title="" WebService Error Codes"" xml:lang=""zh-cn"">")
    End If
    
    Dim tempStr As String
    Dim lastInterface As String
    tempStr = ""
    lastInterface = ""

    For i = 2 To lastrow
        tempStr = resultSheet.Cells(i, 1)
        If lastInterface <> tempStr Then
            sEnXmlContent = XML_APPEND(sEnXmlContent, "<topicref href=""" & tempStr & ".xml""></topicref>")
            lastInterface = tempStr
        End If
    Next i

    sEnXmlContent = XML_APPEND(sEnXmlContent, "</map>")
                
    '处理文件路径
    If language = "英文" Then
        sEnFileContent = Application.ActiveWorkbook.path & "\EN\" & folderName & "\map.ditamap"
    ElseIf language = "中文" Then
        sEnFileContent = Application.ActiveWorkbook.path & "\CN\" & folderName & "\map.ditamap"
    End If
    
    '写入文件
    Dim file As Integer
    file = FreeFile()
    Open sEnFileContent For Binary As file
    Call WriteFile(file, sEnXmlContent)
    CloseFile (file)
    
End Function
