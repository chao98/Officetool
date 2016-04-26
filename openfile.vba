Sub Button1_Click()
    'Put a button into one worksheet
    'Connect button with this macro
    
    Dim filename As String
    Dim fd As Office.FileDialog
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .AllowMultiSelect = False
        .Title = "Please select the file."
        .Filters.Clear
        
        If .Show = True Then
            filename = Dir(.SelectedItems(1))
        End If
    End With
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    If filename <> "" Then
        Worksheets("Control").Range("D2").Value = ThisWorkbook.Path & "\" & filename
    Else
        MsgBox "no file selected"
        Range("D2").Clear
        
    End If
    
End Sub
