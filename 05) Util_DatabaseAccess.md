<h2>Util_DatabaseAccess</h2>

<pre>
Public DBCn As ADODB.Connection
Public DBRs As ADODB.Recordset
Public DBGlbConnection As ADODB.Connection  ' for caching
Public GLBUsername As String
Public GLBPassword As String

Sub TestDBCheckConnection()

    Set DBGlbConnection = Nothing
    Set DBCn = DBCheckConnection(DBCn)
    MsgBox ("Connection Opened")
End Sub

'
' This routine is a workhorse
' It checks to see if the provided object is connected
' if not it checks if the global object is connected
' if so, it uses the global connection
' otherwise, it opens a new connection and saves to the global connection
'
Function DBCheckConnection(Optional DBConn) As ADODB.Connection
Dim haderror As Boolean

TryAgain:  ' loop to recover from login failure
    Call StatusbarDisplay("DBCheckConnection: Check is Nothing.")
    haderror = False
    If Not DBConn Is Nothing Then
        Set DBCheckConnection = DBConn
        Exit Function
    End If
    If DBGlbConnection Is Nothing Then
        Call StatusbarDisplay("DBCheckConnection: Allocate New.")
        Set DBConn = New ADODB.Connection
    Else
        Set DBConn = DBGlbConnection
    End If
    
    Call StatusbarDisplay("DBCheckConnection: Check Open or Closed")
    If DBConn.State = adStateClosed Then
        UserName = RetrieveUserName()
        Password = RetrievePassword()
        If (Len(UserName) = 0 Or Len(Password) = 0) Or Password = "" Then
            Load LoginForm
            LoginForm.txtUsername = UserName
            LoginForm.Show
            If loginCancel Then
                Set DBCheckConnection = Nothing
                Unload LoginForm
                Exit Function
            Else
                UserName = LoginForm.txtUsername
                Password = LoginForm.txtPassword
                Unload LoginForm
            End If
        End If
        
        Call StatusbarDisplay("DBCheckConnection: Opening...")
        
        loginString = "DSN=OGE;Databasename=dbc;Uid=" & UserName & ";PWD=" & Password & ";Authentication Mechanism=LDAP;"
        
        'loginString = "DSN=Saratoga;Description=Saratoga Client;UID=dardanvp;Trusted_Connection=Yes;APP=Microsoft Office 2013;WSID=CORPLT400925;DATABASE=saratoga;"
        
        On Error GoTo LoginError
        DBConn.ConnectionTimeout = 0 'To wait till the query finishes without generating error
        
        '-------------------------------------------------------------------------
        Call StatusbarDisplay("DBCheckConnection: Logging In")
        DBConn.Open loginString
        '-------------------------------------------------------------------------
        If DBConn.State = adStateOpen Then
            '
            ' Save Username and Password
            '
            Call StatusbarDisplay("DBCheckConnection: Logged In")
            Call StoreUserName(UserName)
            Call StorePassword(Password)
            Set DBCheckConnection = DBConn
            'retCode = MsgBox("Connection Successful", vbOKOnly, "Connect To Teradata")
        Else
            MsgBox ("Login Failure")
            Call StatusbarDisplay("Connection Failed")
            Exit Function
        End If
        Application.ODBCTimeout = 900
        DBConn.CommandTimeout = 1200
    End If
    
    Call StatusbarDisplay("DBCheckConnection: Opened")
    Set DBGlbConnection = DBConn
    Set DBCheckConnection = DBConn
    Exit Function
    
OverAndOut:
    DBCheckConnection (DBConn)
    Set DBCheckConnection = Nothing
    DBConn.Close
    Set DBConn = Nothing
    Exit Function
LoginError:
    MsgBox "DBCheckConnection: " & vbNewLine & Err.Description & vbNewLine & vbNewLine & loginString, Title:="Login Error"
    ThisWorkbook.Sheets("TopSheet").Cells(1, 1) = "" ' only way to correct an incorrect Password
    haderror = True
    On Error GoTo 0
    GoTo TryAgain
    
End Function

Function DBCloseConnection(Optional DBConn)
    If IsMissing(DBConn) Then Set DBConn = DBGlbConnection
    If Not DBConn Is Nothing Then
        If DBConn.State <> 0 Then DBConn.Close
        Set DBConn = Nothing
    End If
    If Not DBGlbConnection Is Nothing Then
        If DBGlbConnection.State <> 0 Then DBGlbConnection.Close
        Set DBGlbConnection = Nothing
        On Error GoTo 0
    End If
    'MsgBox "Database Connection Reset", Title:="DBCloseConnection"
End Function

Function DBCheckRecordset(DBRecordset)
    Call StatusbarDisplay("DBCheckRecordset: Check for Nothing.")
    If DBRecordset Is Nothing Then
        Call StatusbarDisplay("DBCheckRecordset: Allocate New.")
        Set DBCheckRecordset = New ADODB.Recordset
    Else
        Set DBCheckRecordset = DBRecordset
    End If
    Call StatusbarDisplay("DBCheckRecordset: Return.")
End Function

Function DBCloseRecordset(DBRecordset)
    If Not DBRecordset Is Nothing Then
        If DBRecordset.State <> 0 Then DBRecordset.Close
        Set DBRecordset = Nothing
    End If
End Function
</pre>
