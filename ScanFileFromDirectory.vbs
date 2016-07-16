Public Function ScanFiles()

   '清空记录
   Set mysheet1 = Application.ActiveSheet
   mysheet1.Range("E1:E100").ClearContents


    Dim sFileName As String
    Dim ntx As Integer
    
    ntx = 0
    
    '查找文件
    sFileName = Dir(ThisWorkbook.path & "\", vbHidden Or vbDirectory Or vbReadOnly Or vbSystem)
    
    Do While sFileName <> ""
    
       If sFileName <> "." And sFileName <> ".." And InStr(1, sFileName, "批量发送附件") <= 0 Then
            ntx = ntx + 1
          
            mysheet1.Cells(ntx, 5) = sFileName

        End If
       sFileName = Dir
    Loop

End Function
