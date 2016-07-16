    Dim EmailEXE As Object
    Dim EmailSub As Object
    Dim BodyStr As String
    Set mysheet = Application.ActiveSheet
    
    Set EmailEXE = CreateObject("Outlook.Application")
        
    
    If InStr(1, mysheet.Cells(1, 2).Text, "@") > 0 And InStr(1, mysheet.Cells(2, 2).Text, "@") > 0 Then
    '主送和抄送都为有效email地址
    Else
        MsgBox "请先在A2，B2两个单元格填写邮件发送地址。"
        Exit Sub
    End If
    
    '扫描当前文件夹下所有压缩文件
    
    Call ScanFiles
       
    '生成邮件
        
    Dim n As Integer
    n = mysheet.Range("E65535").End(xlUp).Row
       

    For i = 1 To n
        If mysheet.Cells(i, 5).Text <> "" Then
        
            Set EmailSub = EmailEXE.CreateItem(olMailItem)
            With EmailSub
                .To = mysheet.Cells(1, 2)    '邮件主送地址
                .CC = mysheet.Cells(2, 2)    '邮件抄送地址
                .Subject = "请收附件 共拆分成" & n & "个邮件发出。"    '邮件标题

                BodyStr = "请获取所有压缩包后，解压。" & vbNewLine & vbNewLine

                .Body = BodyStr
                .Attachments.Add ThisWorkbook.path & "\" & mysheet.Cells(i, 5).Text  'mysheet.Cells(i, 5).Text 存储文件名
                .Save
            End With
        End If
    Next
    
    mysheet.Range("E1:E100").ClearContents
    
    MsgBox "邮件已成功生成，请打开Outlook草稿箱完成发送。"
