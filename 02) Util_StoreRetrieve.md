<h2>Util_StoreRetrieve</h2>

<pre>
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
Function StoreCurrentWorkbook(WBName)
    With ThisWorkbook.Sheets("TopSheet")
        .Cells(3, 1).Font.ThemeColor = xlThemeColorDark1
        .Cells(3, 1).Font.TintAndShade = 0
        .Cells(3, 1) = WBName
    End With
End Function

Function RetrieveCurrentWorkbook(Optional WBName) As String
    WBName = ThisWorkbook.Sheets("TopSheet").Cells(3, 1)
    RetrieveCurrentWorkbook = WBName
End Function

Function StoreCurrentSheet(SHName)
    With ThisWorkbook.Sheets("TopSheet")
        .Cells(4, 1).Font.ThemeColor = xlThemeColorDark1
        .Cells(4, 1).Font.TintAndShade = 0
        .Cells(4, 1) = SHName
    End With
End Function

Function RetrieveCurrentSheet(Optional SHName) As String
    SHName = ThisWorkbook.Sheets("TopSheet").Cells(4, 1)
    RetrieveCurrentSheet = SHName
End Function

Function StoreActiveWorkbookSheet()
    Call StoreCurrentWorkbook(ActiveWorkbook.Name)
    Call StoreCurrentSheet(ActiveSheet.Name)
End Function
</pre>
