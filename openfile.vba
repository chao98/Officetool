Option Explicit

Sub open_Click()
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

Sub copy_Click()
    
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
    Dim tgtUsedC, srcUsedC, srcUsedR As Integer
    tgtUsedC = tgtsht.UsedRange.Columns.Count + 1
    Set tgtRng = tgtsht.Range("A1").Resize(1, tgtUsedC)
    srcUsedC = srcsht.UsedRange.Columns.Count + 1
    srcUsedR = srcsht.UsedRange.Rows.Count + 1
    Set srcRng = srcsht.Range("A1").Resize(1, srcUsedC)
    
    'tgtsht.Activate
    For Each tgtElem In tgtRng
        'tgtElem.Select
        Dim srcTitle, tgtTitle As String
        tgtTitle = tgtElem.Value
        For Each srcElem In srcRng
            srcTitle = srcElem.Value
            If srcTitle = tgtTitle Then
                srcElem.Offset(1, 0).Resize(srcUsedR, 1).Copy tgtElem.Offset(1, 0)
            ElseIf srcTitle = "CSR Number - Key" Then
                srcElem.Offset(1, 0).Resize(srcUsedR, 1).Copy tgtsht.Range("A1").Offset(1, 0)
            'ElseIf srcTitle = "Date: Last GS->LS" Then
            '    srcElem.Offset(1, 0).Resize(srcUsedR, 1).Copy tgtsht.Range("AB1").Offset(1, 0)
            End If
        Next
        
        'If tgtTitle = "Date: Last GS->LS" Then
        '    Dim tmpsht As Worksheet
        '    Set tmpsht = srcwb.Worksheets(1)
        '    Dim i As Integer
        '    For i = 1 To srcUsedC
        '        If tmpsht.Cells(1, i).Value = tgtTitle Then
        '            Exit For
        '        End If
        '    Next
        '    tmpsht.Cells(1, i).Offset(1, 0).Resize(srcUsedR, 1).Copy tgtsht.Range("AB1").Offset(1, 0)
        'End If
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

Sub Clean_Click()
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

Sub inflow_Click()
    
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
    'Set rng = sht.Range("B2:B24")
    'Dim item As Range
    Dim i As Integer
    'For Each item In rng
    For i = 2 To 24
        Select Case sht.Cells(i, "B").Value
            Case Is = "eMBMS In"
                Call cntin(shtname, rawname, wkyear, "Month: Created", i - 1)
            Case Is = "T2 In"
                Call cntin(shtname, rawname, wkyear, "Date: First LS->GS", i - 1)
            Case Is = "T2 Out"
                Call cntin(shtname, rawname, wkyear, "Date: Last GS->LS", i - 1)
            Case Is = "T2 Open"
                Call bdct2in(shtname, rawname, wkyear, "Date: Last GS->LS", "", i - 1)
            Case Is = "BDC T2 In"
                Call bdct2in(shtname, rawname, wkyear, "Node Type", "LINUX BROADCAST DELIVERY CENTER", i - 1)
            Case Is = "BMC T2 In"
                Call bdct2in(shtname, rawname, wkyear, "Node Type", "LINUX BROADCAST MANAGEMENT CENTER", i - 1)
            Case Is = "T2 Low"
                Call bdct2in(shtname, rawname, wkyear, "Severity", "Low", i - 1)
            Case Is = "T2 Medium"
                Call bdct2in(shtname, rawname, wkyear, "Severity", "Medium", i - 1)
            Case Is = "T2 High"
                Call bdct2in(shtname, rawname, wkyear, "Severity", "High", i - 1)
            Case Is = "T2 Hot"
            Case Is = "T2 Emergency"
                Call bdct2in(shtname, rawname, wkyear, "Severity", "Emergency", i - 1)
            Case Is = "T2 Consultant"
                Call bdct2in(shtname, rawname, wkyear, "CSR Type", "Consultant", i - 1)
            Case Is = "T2 Internal"
                Call bdct2in(shtname, rawname, wkyear, "CSR Type", "Internal", i - 1)
            Case Is = "T2 Problem"
                Call bdct2in(shtname, rawname, wkyear, "CSR Type", "Problem", i - 1)
            Case Is = "T2 Project"
                Call bdct2in(shtname, rawname, wkyear, "CSR Type", "Project", i - 1)
            Case Else
                'MsgBox "nothing"
            End Select
    Next
End Sub

Private Function findTitle(ByVal raw As Worksheet, title As String)
    Dim pos, rawCols As Byte
    
    rawCols = raw.UsedRange.Columns.Count - 1
    
    For pos = 1 To rawCols
        If raw.Cells(1, pos).Value = title Then
            Exit For
        End If
    Next

    findTitle = pos
    
End Function

Private Function cntRowsCondition(ByVal raw As Worksheet, ByVal pos1 As Integer, ByVal cond1 As String, ByVal pos2 As Integer, ByVal cond2 As String)
    Dim rawRows As Integer
    rawRows = raw.UsedRange.Rows.Count - 1
    
    Dim i, cnt As Integer
    cnt = 0
    For i = 2 To rawRows
        Dim cell1, cell2 As Range
        Set cell1 = raw.Rows(i).Cells(1, pos1)
        'MsgBox i & " : " & cell1
        If cell1.Value <> "" Then
            If Year(cond1) = Year(cell1.Value) And Month(cond1) = Month(cell1.Value) Then
                Set cell2 = raw.Rows(i).Cells(1, pos2)
                If cell2.Value = cond2 Then
                    cnt = cnt + 1
                End If
            End If
        End If
    Next
    
    cntRowsCondition = cnt
    
End Function

Private Sub bdct2in(ByVal shtname, rawname As String, ByVal wkyear As Integer, ByVal chktitle As String, ByVal chkcond As String, ByVal Oset As Integer)
    Dim raw, inflow As Worksheet
    Set raw = ThisWorkbook.Worksheets(rawname)
    Set inflow = ThisWorkbook.Worksheets(shtname)
    
    
    Dim rawCols, rawRows As Byte
    rawRows = raw.UsedRange.Rows.Count - 1
    rawCols = raw.UsedRange.Columns.Count - 1
    'MsgBox rawRows & " : " & rawCols
    
    Dim title As String

    Dim pos1, pos2 As Byte
    title = "Date: First LS->GS"
    pos1 = findTitle(raw, title)
    'MsgBox pos1
    title = chktitle
    pos2 = findTitle(raw, title)
    'MsgBox pos2
    
    Dim cond1, cond2 As String
    'cond1 = "feb 2016"
    cond2 = chkcond
    
    Dim rng, tmpRng As Range
    Set rng = inflow.Range("C1:N1")
    
    For Each tmpRng In rng
        cond1 = tmpRng.Value & " " & wkyear
        Dim cnt As Integer
        cnt = cntRowsCondition(raw, pos1, cond1, pos2, cond2)
        tmpRng.Offset(Oset, 0).Value = cnt
    Next

    'MsgBox cnt
    
End Sub



Private Sub cntin(ByVal shtname As String, rawname As String, ByVal wkyear As Integer, str As String, Oset As Integer)
    Dim usedR, usedC As Integer
    Dim sht, raw As Worksheet
    'MsgBox shtname & " : " & rawname
    Set sht = Worksheets(shtname)
    Set raw = Worksheets(rawname)
    
    usedR = raw.UsedRange.Rows.Count - 1
    usedC = raw.UsedRange.Columns.Count - 1
    sht.Activate
    
    Dim pos As Integer
    pos = 1
    Do While pos <= usedC And raw.Cells(1, pos).Value <> str
        pos = pos + 1
    Loop
    
    Dim r As Integer
    Dim inflowDate As String
    Dim dateRng, tmpRng As Range
    Set dateRng = sht.Range("C1:N1")
    For Each tmpRng In dateRng
        inflowDate = tmpRng.Value & ", " & wkyear
        Dim cnt As Integer
        cnt = 0
        For r = 2 To usedR
            Dim tmpdate As String
            tmpdate = raw.Rows(r).Cells(1, pos).Value
            'MsgBox tmpdate & " : " & inflowDate
            If tmpdate <> "" Then
                If Year(tmpdate) = Year(inflowDate) Then
                    If Month(tmpdate) = Month(inflowDate) Then
                        cnt = cnt + 1
                    End If
                End If
            End If
        Next
        tmpRng.Offset(Oset, 0).Value = cnt
    Next
End Sub
