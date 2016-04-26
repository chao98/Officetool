Option Explicit

Sub Button1_Click()
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
        'Worksheets("Control").Range("D2").Value = filename
            
    Else
        MsgBox "no file selected"
        Range("D2").Clear
        
    End If
    
End Sub

Sub preCopyCheck(ByVal srcsht, tgtsht As String)
    'MsgBox "start data copy"
    
    If srcsht = "" Or tgtsht = "" Then
        MsgBox "Error, src/tgt sheet to be specified"
        Exit Sub
    Else
        'MsgBox "Src sheet: " & srcsht & Chr(13) _
        '& "Tgt sheet: " & tgtsht
    End If
End Sub

Sub Button2_Click()
    Dim srcsht, tgtsht As String
    srcsht = ActiveSheet.Range("D11").Value
    tgtsht = ActiveSheet.Range("D12").Value
    
    Call preCopyCheck(srcsht, tgtsht)
    
    Dim tgtwb, srcwb As Workbook
    Set tgtwb = ThisWorkbook
    'MsgBox "tgtwb is: " & tgtwb.Name
    Dim srcwbname As String
    srcwbname = ActiveSheet.Range("D2").Value
    Set srcwb = Workbooks.Open(srcwbname)
    
    Dim tgtRng As Range
    Set tgtRng = tgtwb.Worksheets(tgtsht).Range("A1").Offset(1, 0)
    'srcwb.Worksheets(srcsht).Range("B2:DM250").Copy tgtRng
    Dim srcR, srcC As Integer
    srcR = srcwb.Worksheets(srcsht).UsedRange.Rows.Count - 1
    srcC = srcwb.Worksheets(srcsht).UsedRange.Columns.Count
    'MsgBox "srcR : srcC - " & srcR & " : " & srcC
    srcwb.Worksheets(srcsht).Range("B2").Resize(srcR, srcC).Copy tgtRng
    
    tgtwb.Worksheets(tgtsht).Activate
    ActiveSheet.Range("A2").Select
    srcwb.Close False
    
End Sub

Sub Button3_Click()
    Dim shtname As String
    shtname = ActiveSheet.Range("D7").Value
    If shtname = "" Then
        MsgBox "Error: no sheet specified to clean!"
        Exit Sub
    Else
        'MsgBox "Error: sheet name is not correct!"
        'Exit Sub
    End If
    
    'Dim sht As Worksheet
    'Set sht = Application.Worksheets(shtname)
    'If sht Is Nothing Then
    '    MsgBox "Error: sheet name is not correct!"
    '    Exit Sub
    'End If
    
    Dim usedR, usedC As Integer
    usedR = Worksheets(shtname).UsedRange.Rows.Count - 1
    usedC = Worksheets(shtname).UsedRange.Columns.Count
    MsgBox "usedR : usedC: " & usedR & " : " & usedC
    
    'Dim rng As String
    'rng = "2:" & usedR
    'MsgBox "range to clean: R" & rng
    Worksheets(shtname).Range("A2").Resize(usedR, usedC).Clear
    
End Sub
