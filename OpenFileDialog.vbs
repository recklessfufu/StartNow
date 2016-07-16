Public Function OpenFileDialog() As String
    
    Dim fd As FileDialog
    Dim vrtSelectedItem As Variant
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
   
    With fd                                      
    '设置fd的对话框参数
        .Title = "Please select file!"
        .ButtonName = "确定"
        .AllowMultiSelect = False
        .Filters.Add "Excel Files", "*.xls; *.xlsm; *.xlsx", 1
    End With

    If fd.Show = -1 Then
        For Each vrtSelectedItem In fd.SelectedItems
            OpenFileDialog = vrtSelectedItem
        Next vrtSelectedItem
    End If
    
    vrtSelectedItem = 0
    Set fd = Nothing
    
End Function
