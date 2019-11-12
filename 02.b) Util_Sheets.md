

<pre>
' November 17th, 2018
'
'

Sub InitSheet(SHName)
    Call DeleteSheet(SHName)
    Call MakeSheet(SHName)
End Sub

Sub DeleteSheet(SHName, Optional WBName)
Dim t As String

    If IsMissing(WBName) Then WBName = ActiveWorkbook.Name

    If IsNumeric(SHName) Then
        t = """" & SHName & """"
    Else
        t = SHName
    End If
    
    On Error Resume Next
    Call AlertsOff
        Workbooks(WBName).Sheets(t).Delete
    Call AlertsOn
End Sub

Function BaseSheet()
    BaseSheet = Worksheets(Worksheets.Count).Name
End Function

Sub CopySheetHeader(SHFrom, SHTo)
    Worksheets(SHFrom).Rows(1).Copy Destination:=Worksheets(SHTo).Cells(1, 1)
    Call ClearClipboard
End Sub
Sub CheckSheetExists(SHName, Optional WBUse)

    If IsMissing(WBUse) Then WBUse = ActiveWorkbook.Name
    
    SHOrig = ActiveSheet.Name
    If (SheetExists(SHName, WBUse)) Then
        retCode = MsgBox("Reset " & SHName & "?", vbYesNoCancel)
        If (retCode = vbCancel) Then Exit Sub
        If (retCode = vbYes) Then
            Call AlertsOff
            Worksheets(SHName).Delete
            Call AlertsOn
        End If
    End If
    
    Sheets.Add
    ActiveSheet.Name = SHName
    Workbooks(WBUse).Worksheets(SHOrig).Activate
End Sub
Function SheetExists(Optional SHName, Optional WBName)
    
    If IsMissing(WBName) Then WBName = ActiveWorkbook.Name
    If IsMissing(SHName) Then SHName = ActiveSheet.Name
    On Error GoTo NoSheet
        n = Workbooks(WBName).Worksheets(SHName).Cells(1, 1)
        SheetExists = True
        Exit Function
NoSheet:
        SheetExists = False
        On Error GoTo 0

End Function

Function NextSheetName(Optional SHName, Optional WBName) As String
    
    If IsMissing(WBName) Then WBName = ActiveWorkbook.Name
    If IsMissing(SHName) Then SHName = ActiveSheet.Name
    k = 0
    On Error GoTo NoSheet
        t = Workbooks(WBName).Sheets(SHName).Name
        For i = 1 To Workbooks(WBName).Sheets.Count
            If InStr(Workbooks(WBName).Worksheets(i).Name, SHName) Then k = k + 1
        Next i
    On Error GoTo 0
    NextSheetName = SHName & " " & Chr(65 + k)
    Exit Function
NoSheet:
        NextSheetName = SHName
        On Error GoTo 0

End Function

'
' Returns the name of the new sheet
'
Function MakeSheet(Optional copyHeader, Optional SHName)
    SHCurrent = ActiveSheet.Name
    If IsMissing(SHName) Then SHName = "Sheet"
    nSheets = SheetExists(SHName)
    If (nSheets > 0) Then
        If Not (SHName = "Sheet") Then
            retCode = MsgBox("Make new " & SHName & "?", vbYesNoCancel)
        End If
        If (retCode = vbCancel) Then Exit Function
        If (retCode = vbNo) Then
            MakeSheet = SHName
            Exit Function
        End If
        If (retCode = vbYes) Then SHName = SHName & "-" & (nSheets + 1)
    End If
    
    If Not SheetExists(SHName) Then
        Set SHNew = ActiveWorkbook.Sheets.Add
        SHNew.Name = SHName
    End If
    Sheets(SHCurrent).Activate
    If Not IsMissing(copyHeader) Then
        If (copyHeader) Then Call HeaderToSheet(SHCurrent, SHName)
        
    Else
        retCode = MsgBox("Copy Headers To Scratch?", vbYesNoCancel)
        If (retCode = vbCancel) Then Exit Function
        If (retCode = vbYes) Then Call HeaderToSheet(SHCurrent, SHName)
    End If
    
    MakeSheet = SHName
End Function

Sub AddColumnToSheet(fromCol, SHFrom, SHTo, Optional colWidth, Optional colAlign)
    If (IsNumeric(fromCol)) Then
        useCol = fromCol
    Else
        useCol = FindColumnHeader(fromCol, SHFrom)
    End If
    nextCol = NextColumn(SHTo)
    If useCol > 0 Then
        If IsMissing(colWidth) Then colWidth = 8.43
        Call CopyColumnToSheet2(useCol, SHFrom, nextCol, SHTo, colWidth, colAlign)
    Else
        Worksheets(SHTo).Cells(1, nextCol) = fromCol
    End If
End Sub

Sub CopyColumnToSheet2(fromCol, SHFrom, toCol, SHTo, Optional colWidth, Optional colAlign)
    botRow = ColumnLastRow(fromCol, SHFrom)
    
    Set fromRange = Range(Worksheets(SHFrom).Cells(1, fromCol), Worksheets(SHFrom).Cells(botRow, fromCol))
    
    Set toRange = Range(Worksheets(SHTo).Cells(1, toCol), Worksheets(SHTo).Cells(1, toCol))
    
    fromRange.Copy Destination:=toRange
    
    If Not IsMissing(colWidth) Then Worksheets(SHTo).Columns(toCol).ColumnWidth = colWidth
    If Not IsMissing(colAlign) Then Worksheets(SHTo).Columns(toCol).Alignment = colAlign
    
    Call ClearClipboard
    
End Sub
</pre>
