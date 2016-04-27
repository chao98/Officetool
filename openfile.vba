Option Explicit

Sub Button1_Click()
    Dim filename As String
    Dim fd As Office.FileDialog
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .AllowMultiSelect = False
        .title = "Please select the file."
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
    
    Dim srcwb, tgtwb As Workbook
    Dim srcfn As String
    srcfn = Worksheets("Control").Range("D2").Value
    Set srcwb = Workbooks.Open(srcfn)
    Set tgtwb = ThisWorkbook
    
    Dim srcsht, tgtsht As Worksheet
    Dim srcshtname, tgtshtname As String
    tgtshtname = tgtwb.Worksheets("Control").Range("D12").Value
    Set tgtsht = tgtwb.Worksheets(tgtshtname)
    srcshtname = tgtwb.Worksheets("Control").Range("D11").Value
    Set srcsht = srcwb.Worksheets(srcshtname)
    
    Dim tgtRng, tgtElem As Range
    Dim srcRng, srcElem As Range
    Dim tgtUsedC, srcUsedC As Integer
    tgtUsedC = tgtsht.UsedRange.Columns.Count + 1
    Set tgtRng = tgtsht.Range("A1").Resize(1, tgtUsedC)
    srcUsedC = srcsht.UsedRange.Columns.Count + 1
    Set srcRng = srcsht.Range("A1").Resize(1, srcUsedC)
    
    'tgtsht.Activate
    For Each tgtElem In tgtRng
        'tgtElem.Select
        Dim srcTitle, tgtTitle As String
        tgtTitle = tgtElem.Value
        For Each srcElem In srcRng
            srcTitle = srcElem.Value
            If srcTitle = tgtTitle Then
                srcElem.Offset(1, 0).Resize(250, 1).Copy tgtElem.Offset(1, 0)
            ElseIf srcTitle = "CSR Number - Key" Then
                srcElem.Offset(1, 0).Resize(250, 1).Copy tgtsht.Range("A1").Offset(1, 0)
            End If
        Next
    Next
    tgtsht.Activate
    tgtsht.Range("A2").Select
    srcwb.Close
    
End Sub

Sub Button2_Click_old()
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
    Dim srcR, srcC As Integer
    srcR = srcwb.Worksheets(srcsht).UsedRange.Rows.Count - 1
    srcC = srcwb.Worksheets(srcsht).UsedRange.Columns.Count
    'MsgBox "srcR : srcC - " & srcR & " : " & srcC
    'srcwb.Worksheets(srcsht).Range("B2").Resize(srcR, srcC).Copy tgtRng
    
    Call copydata
    
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
    'MsgBox "usedR : usedC: " & usedR & " : " & usedC
    
    'Dim rng As String
    'rng = "2:" & usedR
    'MsgBox "range to clean: R" & rng
    Worksheets(shtname).Range("A2").Resize(usedR, usedC).Clear
    
End Sub

Sub Button4_Click()
    
    Dim shtname, rawname As String
    shtname = Worksheets("Control").Range("D15").Value
    rawname = Worksheets("Control").Range("D12").Value
    Dim wkyear As Integer
    wkyear = Worksheets("Control").Range("D16").Value
    Worksheets(rawname).Activate
    
    Dim sht, raw As Worksheet
    Dim rng As Range
    Set sht = Worksheets(shtname)
    Set raw = Worksheets(rawname)
    Set rng = sht.Range("B2:B24")
    Dim item As Range
    For Each item In rng
        Select Case item.Value
            Case Is = "BDC In"
                Call bdcin(sht, raw, wkyear)
            Case Is = "BDC T2 In"
                'MsgBox "BDC T2 In"
            Case Else
                'MsgBox "nothing"
            End Select
    Next
End Sub

Private Sub bdcin(ByVal sht As Worksheet, raw As Worksheet, wkyear As Integer)
    Dim created As Integer
    Dim usedR, usedC As Integer
    usedR = raw.UsedRange.Rows.Count - 1
    usedC = raw.UsedRange.Columns.Count - 1
    
    created = 1
    Do While created <= usedC And Cells(1, created).Value <> "Month: Created"
        created = created + 1
    Loop
    
    Dim r As Integer
    Dim inflowMonth As String
    Dim dateRng, tmpRng As Range
    Set dateRng = sht.Range("C1:N1")
    For Each tmpRng In dateRng
        inflowMonth = tmpRng.Value & ", " & wkyear
        Dim cnt As Integer
        cnt = 0
        For r = 2 To usedR
            Dim tmpdate As String
            tmpdate = raw.Rows(r).Cells(1, created).Value
            If Year(tmpdate) = Year(inflowMonth) Then
                If Month(tmpdate) = Month(inflowMonth) Then
                    cnt = cnt + 1
                End If
            End If
        Next
        tmpRng.Offset(1, 0).Value = cnt
    Next
End Sub
