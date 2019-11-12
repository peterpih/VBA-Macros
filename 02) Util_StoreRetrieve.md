<h2>Util_StoreRetrieve</h2>

<pre>
'
' November 12, 2019
' Routines for saving fetching Username and Password
'
' Row locations on TopSheet
Public Const PASSWORDROW = 1
Public Const USERNAMEROW = 2
Public Const ACTIVEWORKBOOKROW = 3
Public Const ACTIVESHEETROW = 4
'
Function StorePassword(Password)
    With ThisWorkbook.Sheets("TopSheet")
        .Cells(1, 1).Font.ThemeColor = xlThemeColorDark1
        .Cells(1, 1).Font.TintAndShade = 0
        .Cells(1, 1) = Password
    End With
End Function

Function RetrievePassword(Optional Password) As String
    Password = ThisWorkbook.Sheets("TopSheet").Cells(1, 1)
    RetrievePassword = Password
End Function

Function StoreUserName(UserName)
    With ThisWorkbook.Sheets("TopSheet")
        .Cells(2, 1).Font.ThemeColor = xlThemeColorDark1
        .Cells(2, 1).Font.TintAndShade = 0
        .Cells(2, 1) = UserName
    End With
End Function

Function RetrieveUserName(Optional UserName) As String
    UserName = ThisWorkbook.Sheets("TopSheet").Cells(2, 1)
    RetrieveUserName = UserName
End Function
'------------------------------------------------------------------------------------------
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

Function CheckTopSheet() As Boolean
    WBActive = ActiveWorkbook.Name
    SHActive = ActiveSheet.Name
    On Error GoTo NameTopSheet
        t = ThisWorkbook.Sheets("TopSheet").Name
        CheckTopSheet = True
        GoTo EndReturn
NameTopSheet:
    If ThisWorkbook.Sheets.Count >= 1 Then
        ThisWorkbook.Sheets(1).Name = "TopSheet"
    End If
    CheckTopSheet = True
EndReturn:
    On Error GoTo 0
    'Workbooks(WBActive).Sheets(SHActive).Activate
End Function

Function GetActiveWorkbook(Optional WBName) As String

    If Not IsMissing(WBName) Then WBName = ThisWorkbook.Sheets("TopSheet").Cells(ACTIVEWORKBOOKROW, 1)
    GetActiveWorkbook = ThisWorkbook.Sheets("TopSheet").Cells(ACTIVEWORKBOOKROW, 1)
    
End Function

Function SetActiveWorkbook(Optional WBName) As String

    Call CheckTopSheet
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

    Call CheckTopSheet
    If Not IsMissing(SHName) Then
        ThisWorkbook.Sheets("TopSheet").Cells(ACTIVESHEETROW, 1) = SHName
    End If
    SetActiveSheet = SHName
    
End Function
</pre>
