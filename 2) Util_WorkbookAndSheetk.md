<h2>SetActiveWorkbook / GetActiveWorkbook</h2>

<pre>
' Row locations on TopSheet
Public Const PASSWORDROW = 1
Public Const USERNAMEROW = 2
Public Const ACTIVEWORKBOOKROW = 3
Public Const ACTIVESHEETROW = 4

Sub test()
Dim t() As String
    t() = Split("A1,$a$6,zoo", ",")
    
    For i = 0 To UBound(t)
        Debug.Print t(i)
        
    Next i
    For Each R In Selection
    Debug.Print R.Column
    Next R
End Sub

Function GetActiveWorkbook(Optional WBName) As String

    If Not IsMissing(WBName) Then WBName = ThisWorkbook.Sheets("TopSheet").Cells(ACTIVEWORKBOOKROW, 1)
    GetActiveWorkbook = ThisWorkbook.Sheets("TopSheet").Cells(ACTIVEWORKBOOKROW, 1)
    
End Function

Function SetActiveWorkbook(Optional WBName) As String

    'WBOld = GetActiveWorkbook
    If Not IsMissing(WBName) Then
        t = Workbooks(WBName).Name '= WBName  ' this is a read only property, need to .saveas to change name
        ThisWorkbook.Sheets("TopSheet").Cells(ACTIVEWORKBOOKROW, 1) = WBName
    End If
    SetActiveWorkbook = WBName
    
End Function

Function GetActiveSheet(Optional SHName) As String

    If Not IsMissing(SHName) Then SHName = ThisWorkbook.Sheets("TopSheet").Cells(ACTIVESHEETROW, 1)
    GetActiveSheet = ThisWorkbook.Sheets("TopSheet").Cells(ACTIVESHEETROW, 1)
    
End Function

Function SetActiveSheet(Optional SHName) As String

    'WBName = GetActiveWorkbook
    If Not IsMissing(SHName) Then
        'SHOld = GetActiveSheet
        'If Not Workbooks(WBName).Sheets(SHOld).Name = SHName Then Workbooks(WBName).Sheets(SHOld).Name = SHName
        ThisWorkbook.Sheets("TopSheet").Cells(ACTIVESHEETROW, 1) = SHName
    End If
    SetActiveSheet = SHName
    
End Function

Function SetActiveWorkbookAndSheet(Optional WBName, Optional SHName)

    If IsMissing(WBName) Then WBName = ""
    If IsMissing(SHName) Then SHName = ""
    
    If WBName = "" Or SHName = "" Then
        Set selectRange = Nothing
        On Error Resume Next
        Set selectRange = Application.InputBox("Select Workbook and Sheet", Type:=8)
        If selectRange Is Nothing Then Exit Function
        
        WBName = selectRange.Parent.Parent.Name
        SHName = selectRange.Parent.Name
    End If
    
    If Not IsMissing(WBName) Then Call SetActiveWorkbook(WBName)
    If Not IsMissing(SHName) Then Call SetActiveSheet(SHName)

End Function

Function GetActiveWorkbookAndSheet(WBName, Optional SHName)

    WBName = GetActiveWorkbook
    If Not IsMissing(SHName) Then SHName = GetActiveSheet
    
End Function

Function ActivateActiveWorkbookAndSheet(Optional WBName, Optional SHName)

    On Error GoTo NoWorkbookErr
    If IsMissing(WBName) Or IsEmpty(WBName) Then WBName = GetActiveWorkbook
    Workbooks(WBName).Activate
    
    On Error GoTo NoSheetErr
    If IsMissing(SHName) Or IsEmpty(SHName) Then
        SHName = GetActiveSheet
    Else
         Call SetActiveSheet(SHName)
    End If
    
    On Error GoTo GeneralErr
    Workbooks(WBName).Sheets(SHName).Activate
    On Error GoTo 0
    
    Exit Function
    
NoSheetErr:
    MsgBox ("No Sheet ~" & SHName & "~")
    Stop
    Exit Function

NoWorkbookErr:
    MsgBox ("No workbook ~" & WBName & "~")
    Stop
    Exit Function
    
GeneralErr:
    MsgBox ("General Error")
    Stop
    Exit Function
End Function

Function NewActiveWorkbookandSheet(Optional WBName, Optional SHName)
    ' currently ignores WBName
    WBName = Workbooks.Add.Name
    SetActiveWorkbook (WBName)
    SetActiveSheet (ActiveSheet.Name)
    SHOld = GetActiveSheet
    Workbooks(WBName).Sheets.Add.Name = SHName
    Call SetActiveSheet(SHName)
End Function

Function NewActiveSheet(Optional SHName)
    WBName = GetActiveWorkbook
    Workbooks(WBName).Sheets.Add
    
    If Not IsMissing(SHName) Then
        ActiveSheet.Name = SHName
    Else
        SHName = ActiveSheet.Name
    End If

    Call SetActiveSheet(SHName)
End Function

Function DeleteActiveWorkbook()
    Workbooks(GetActiveWorkbook).Close (False)
    Call SetActiveWorkbook("")
    Call SetActiveSheet("")
End Function
</pre>
