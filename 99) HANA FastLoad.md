

<pre>
' April 8th, 2019 - deal with ":" in column header
' Novemeber 19th, 2018
'
Public GLBTableName As String
'
'
'
' This function replaces invalid characters in table names
'
Function MyTrim(t) As String
Dim s As String
    s = Trim(t)
    s = Replace(s, Chr(160), "")  ' this is from using copy dfrom web and fixed column width conversion
    s2 = left(s, 1)
    If IsNumeric(s2) Then s = "Z" & Right(s, Len(s) - 1)
    MyTrim = s
End Function
Function TrimReplace(t)
Dim s As String
    s = MyTrim(t)
    s = Replace(s, ".", "")
    s = Replace(s, " ", "_")
    s = Replace(s, "(", "_")
    s = Replace(s, ")", "_")
    s = Replace(s, "/", "_")
    s = Replace(s, "-", "_")
    s = Replace(s, "?", "_")
    s = Replace(s, ">", "_")
    s = Replace(s, "<", "_")
    s = Replace(s, "'", "")
    s = Replace(s, ":", "")
    s = Replace(s, """", "")
    s = Replace(s, "%", "pct")
    t1 = Len(s)
    s = Replace(s, "__", "_")
    If Len(s) <> t1 Then s = Replace(s, "__", "_")
    t1 = Len(s)
    If Len(s) <> t1 Then s = Replace(s, "__", "_")
    t1 = Len(s)
    If Len(s) <> t1 Then s = Replace(s, "__", "_")
    TrimReplace = s
End Function
'
' Returns:
'   0 - error
'   1 - Success
'   2 - Column Names error
'   3 - Stop
Function HANAFastLoad(Optional tableName, Optional SHSource, Optional WBSource) As Integer
Dim charWidth As Integer
Dim wsh As Object
Dim userTableName As String     ' from user
Dim newTableName As String      ' table name after TrimReplace, user may have used invalid characters
Dim fullTableName As String     ' database table name

Dim waitOnReturn As Boolean: waitOnReturn = True
Dim windowStyle As Integer: windowStyle = 1 ' or whatever suits you best
Dim emptyColumnCount As Integer: emptyColumnCount = 1
Dim errorCode As Integer
Dim mergeTable As Boolean       ' whether a merge is necessary because the table already exists


On Error GoTo gotError

    On Error Resume Next

    WBOrig = RetrieveCurrentWorkbook()
    SHOrig = RetrieveCurrentSheet
    
    If IsMissing(tableName) Then
        On Error Resume Next
        HANATableNameForm.Show
        If formCancel Then
            Call ProgressFormClose
            Unload HANATableNameForm
            FastLoad = 3
            Exit Function
        Else
            tableName = HANATableNameForm.txtTableName
            Unload HANATableNameForm
        End If
    End If

    Call UsageTracker("FastLoad", "Start: " & tableName)
    Call ProgressFormAdd("FastLoad - Start")
    Set wsh = CreateObject("WScript.Shell")
    
    GLBTableName = userTableName
    newTableName = TrimReplace(userTableName)
    If newTableName <> userTableName Then
        Call ProgressFormAdd(userTableName & "  ->  " & newTableName)
        retCode = MsgBox("Modifying Table Name:" & vbNewLine & vbNewLine & userTableName & "  ->  " & newTableName, vbOKCancel, Title:="FastLoad")
        If retCode = vbCancel Then Exit Function
    End If
    '
    '
    DatabaseName = RetrieveUserName()
    
    fullTableName = DatabaseName & "." & tableName
    Call UsageTracker("FastLoad", fullTableName)
    '
    ' Get the Column Formats
    '
    Workbooks.Add
    WBFmt = ActiveWorkbook.Name
    SHFmt = ActiveSheet.Name
    Call FastLoad_ColumnNamesFormat(SHOrig, WBOrig, SHFmt, WBFmt)
    
    If TDTableExists(fullTableName) Then
        Call ProgressFormAdd("FastLoad - Table Exists")
        '
        '-- Check Table and FastLoad Columns ------------------------------------------------------------
        '
        If Not FastLoad_CheckColumnNames(tableName, SHOrig, WBOrig, SHFmt, WBFmt) Then
            Call ProgressFormAdd("FastLoad - Error Column Name")
            Exit Function
        End If

        Call ProgressFormAdd("FastLoad - Already Loaded?")
        If FastLoad_Check(fullTableName, "event_log_id", SHOrig, WBOrig) Then
            MsgBox ("Already Loaded")
            FastLoad = 0
            Application.DisplayAlerts = False
            Workbooks(WBFmt).Close (False)
            Application.DisplayAlerts = True
            Exit Function
        End If
        Call ProgressFormAdd("FastLoad - Will Append Table")
        If HANA_LoadData(fullTableName, SHOrig, WBOrig, SHFmt, WBFmt) = True Then
            MsgBox ("HANA - " & fullTableName & " loaded.")
            Exit Function
        End If

        'Call Table_ColumnFormat
       
        mergeTable = True
        newTableName = newTableName & "_up"
        fullTableName = DatabaseName & "." & tableName
        Call ProgressFormAdd("FastLoad - Appending Table")
    Else
        'If useMenus Then
        '    retCode = MsgBox("Creating Table: " & newTableName, vbOKCancel, Title:="Fastload")
        '    If retCode = vbCancel Then Exit Function
        'Else
            Call ProgressFormAdd("FastLoad - Creating Table")
            
            If HANA_CreateTable(fullTableName, SHFmt, WBFmt) = True Then
                If HANA_LoadData(fullTableName, SHOrig, WBOrig, SHFmt, WBFmt) = True Then
                    MsgBox ("HANA - " & fullTableName & " loaded.")
                    Exit Function
                End If
            End If
        'End If
        mergeTable = False
    End If
    

    '
    filePath = "C:\oge\fastload\" & newTableName & ".fl"
    On Error Resume Next
    Kill filePath
    
    On Error GoTo gotError
    Call StatusbarDisplay("Fastload: Setup")
    Call ProgressFormAdd("FastLoad - Setup")
    
    Call FastLoadWrite(filePath, "LOGMECH LDAP;")
    UserName = LCase(Environ$("Username"))
    Call FastLoadWrite(filePath, "LOGON TD1/" & RetrieveUserName & "," & RetrievePassword & ";")
    Call FastLoadWrite(filePath, "DATABASE dl_oge_analytics;")
    '
    ' DROP TABLES ------------------------------------------------------------------------------
    '
    Call FastLoadWrite(filePath, "DROP TABLE " & fullTableName & ";")
    Call FastLoadWrite(filePath, "DROP TABLE " & fullTableName & "_ET;")
    Call FastLoadWrite(filePath, "DROP TABLE " & fullTableName & "_UV;")
    '
    ' CREATE TABLE -----------------------------------------------------------------------------
    '
    Call StatusbarDisplay("Fastload: Create Table")
    Call FastLoadWrite(filePath, "")
    Call FastLoadWrite(filePath, "CREATE MULTISET TABLE " & fullTableName & ",")
    Call FastLoadWrite(filePath, "NO FALLBACK,")
    Call FastLoadWrite(filePath, "NO BEFORE JOURNAL,")
    Call FastLoadWrite(filePath, "NO AFTER JOURNAL,")
    Call FastLoadWrite(filePath, "CHECKSUM = DEFAULT,")
    Call FastLoadWrite(filePath, "DEFAULT MERGEBLOCKRATIO")
    Call FastLoadWrite(filePath, "(")
    rightCol = RowLastColumn(1, SHOrig, WBOrig)
    '
    ' COLUMN NAMES -----------------------------------------------------------------------------
    '
    Call FastLoadWrite(filePath, "LoadDate DATE,") ' load date column

    '=============================================================================================================================================
    
    '-- Get the Column Format from The Format Sheet for DATE and TIMESTAMP ---------------------------------------------------
    For i = 1 To rightCol
        t = Workbooks(WBFmt).Worksheets(SHFmt).Cells(1, i) ' Column Name
        c = ","
        If i = rightCol Then c = ")"

        m = Workbooks(WBFmt).Worksheets(SHFmt).Cells(2, i) ' column format
        Call FastLoadWrite(filePath, t & " " & m & c)

    Next i

    't = CheckReservedWord(Workbooks(WBFmt).Worksheets(SHFmt).Cells(1, 1))
    t = Workbooks(WBFmt).Worksheets(SHFmt).Cells(1, 1)
    Call FastLoadWrite(filePath, "PRIMARY INDEX(" & t & ");") 'set first column as primary index to spread processing
    '
    '
    Call FastLoadWrite(filePath, "")
    Call FastLoadWrite(filePath, "BEGIN LOADING " & fullTableName)
    Call FastLoadWrite(filePath, "ERRORFILES " & newTableName & "_ET, " & newTableName & "_UV;")
    Call FastLoadWrite(filePath, "SET RECORD VARTEXT delimiter " & "'|' QUOTE YES " & "'" & """" & "'" & ";")
    '
    ' DEFINE -------------------------------------------------------------------------------------
    '
    Call StatusbarDisplay("Fastload: Define")
    Call FastLoadWrite(filePath, "")
    Call FastLoadWrite(filePath, "DEFINE")
    Call FastLoadWrite(filePath, "in_LoadDate (varchar(20)),")
    For i = 1 To rightCol
        t = Workbooks(WBFmt).Worksheets(SHFmt).Cells(1, i) ' formatted header
        t = "in_" & TrimReplace(t)
        c = ","
        If i = rightCol Then c = ""
        Call FastLoadWrite(filePath, t & " (varchar(300))" & c)
    Next i
    Call FastLoadWrite(filePath, "FILE= " & newTableName & ".txt;")
    '
    ' INSERT --------------------------------------------------------------------------------------
    '
    Call FastLoadWrite(filePath, "")
    Call FastLoadWrite(filePath, "INSERT INTO " & fullTableName & " (")
    Call FastLoadWrite(filePath, "LoadDate,")
    For i = 1 To rightCol
        t = Workbooks(WBFmt).Worksheets(SHFmt).Cells(1, i) ' formatted header
        t = CheckReservedWord(t)
        c = ","
        If i = rightCol Then c = ")"
        Call FastLoadWrite(filePath, t & c)
    Next i
    '
    ' VALUES --------------------------------------------------------------------------------------
    '
    Call FastLoadWrite(filePath, "")
    Call FastLoadWrite(filePath, "VALUES (")
    Call FastLoadWrite(filePath, ": in_LoadDate,")
    For i = 1 To rightCol
        t = Workbooks(WBFmt).Worksheets(SHFmt).Cells(1, i) ' formatted header
        t = ": in_" & TrimReplace(t)
        c = ","
        If i = rightCol Then c = ");"
        Call FastLoadWrite(filePath, t & c)
    Next i
    '
    ' END LOADING ---------------------------------------------------------------------------------
    '
    Call FastLoadWrite(filePath, "END LOADING;")
    Call FastLoadWrite(filePath, "LOGOFF;")
    '
    ' DATA FILE -----------------------------------------------------------------------------------
    '
    Call StatusbarDisplay("Fastload: Create Data File")
    filePath = "C:\OGE\fastload\" & newTableName & ".txt"
    On Error Resume Next
    Kill filePath
    
    
    Set foundRange = FindRangeErrors(Workbooks(WBOrig).Sheets(SHOrig).Cells)
    foundRange.Value = ""  ' clear all #N/As
    
    On Error GoTo gotError
    botRow = ColumnLastRow(1, SHOrig, WBOrig)
    For j = 2 To botRow
        aline = """" & Format(Now(), "yyyy-mm-dd") & """" & "|"  ' for LoadDate
        'aline = ""
        For i = 1 To rightCol
            'If j = botRow + 10 Then
            '    ' get the format frm the first line of data, header may only be "General"
            '    aline = aline & """" & Workbooks(WBOrig).Worksheets(SHOrig).Cells(j + 1, i).NumberFormat & """"
            'Else
            '----------
            '--If Workbooks(WBFmt).Worksheets(SHFmt).Cells(2, i) = "DATE" Then
            '--    t = Trim(Format(Workbooks(WBOrig).Worksheets(SHOrig).Cells(j, i).text, "yyyy-mm-dd"))
            '--ElseIf Workbooks(WBFmt).Worksheets(SHFmt).Cells(2, i) = "TIME" Then
            '--    t = Trim(Format(Workbooks(WBOrig).Worksheets(SHOrig).Cells(j, i).text, "hh:mm:ss"))
            '--ElseIf Workbooks(WBFmt).Worksheets(SHFmt).Cells(2, i) = "TIMESTAMP" Then
            '--    t = Trim(Workbooks(WBOrig).Worksheets(SHOrig).Cells(j, i).Value)
            '--Else
            '--    t = Trim(Workbooks(WBOrig).Worksheets(SHOrig).Cells(j, i).Value)
            '--End If
            
            If Workbooks(WBFmt).Worksheets(SHFmt).Cells(2, i) = "DATE" Then
                t = reformatDate(Trim(Workbooks(WBOrig).Worksheets(SHOrig).Cells(j, i).text))
            ElseIf Workbooks(WBFmt).Worksheets(SHFmt).Cells(2, i) = "TIME" Then
                t = reformatTime(Trim(Workbooks(WBOrig).Worksheets(SHOrig).Cells(j, i).text))
            ElseIf Workbooks(WBFmt).Worksheets(SHFmt).Cells(2, i) = "TIMESTAMP" Then
                t = reformatTimestamp(Trim(Workbooks(WBOrig).Worksheets(SHOrig).Cells(j, i).Value))
            Else
                t = Trim(Workbooks(WBOrig).Worksheets(SHOrig).Cells(j, i).Value)
            End If
            
            t = Replace(t, """", "")
            aline = aline & """" & t & """"

            c = "|"
            If i = rightCol Then c = ""
            aline = aline & c
        Next i
        Call FastLoadWrite(filePath, aline)
    Next j
    Call StatusbarDisplay("Fastload: Shell Run")
    '
    ' Shell DOS command ---------------------------------------------------------------------------
    '
    Call ProgressFormAdd("FastLoad - Running")
    t = "cmd.exe /c cd /d C:\oge\fastload && fastload < " & newTableName & ".fl"
    output = ShellRun("cmd.exe /c cd /d C:\oge\fastload && fastload < " & newTableName & ".fl")
    
    filePath = "C:\OGE\fastload\" & newTableName & ".log"
    On Error Resume Next
    Kill filePath
    
    On Error GoTo gotError
    Call FastLoadWrite(filePath, output)
    '
    ' Need to MERGE?
    If mergeTable Then
        'Call DBCloseConnection(DBCn)
        Call FastloadMerge(fullTableName)
    End If
    '
    ' Extract return code
    '
    If showmenus Then
        Load TextForm
        TextForm.txtBody = output
    End If
    
    '-------------------------------------------------------------------------------------------------------
    ' Check the return Stats from FastLoad
    '
     k = InStr(1, output, "END LOADING COMPLETE")
     t = Mid(output, k)

     k = InStr(1, t, "Total Records Read")
     k2 = InStr(k, t, vbNewLine)
     t2 = Mid(t, k, k2 - k)
     m = InStr(1, t2, "=") + 1
     records_read = Trim(Mid(t2, m))
     k = InStr(1, t, "Total Inserts Applied")
     k2 = InStr(k, t, vbNewLine)
     t2 = Mid(t, k, k2 - k)
     m = InStr(1, t2, "=") + 1
     inserts_applied = Trim(Mid(t2, m))
    If records_read = inserts_applied Then
        TextForm.txtHeader = "Success"
        Call ProgressFormAdd("FastLoad - Success")
    Else
        TextForm.txtHeader = "Failed"
        Call ProgressFormAdd("FastLoad - Failed")
    End If
    If showmenus Then
        TextForm.Show
        Unload TextForm
    End If

    Call UsageTracker("FastLoad", "Finished")
    
    Call DBCloseConnection(DBCn)
    Call DBCloseRecordset(DBRs)
    
    Workbooks(WBFmt).Close (False)
    
    FastLoad = 1
    Exit Function
    
gotError:
    t = Err.Number & " " & Err.Description & vbNewLine & vbNewLine & "Error on line: " & Erl
    MsgBox t, Title:="Fastload"
    Call UsageTracker("FastLoad", t)
    Stop
    Resume Next
    
End Function

Sub FastloadMerge(fullTableName)
    
Dim DBCn As ADODB.Connection
Dim DBRs As ADODB.Recordset

    Call ProgressFormAdd("FastLoad - Merge")

On Error GoTo gotError
     Call DBCloseConnection(DBCn)
10    qq = "INSERT INTO " & left(fullTableName, Len(fullTableName) - 3) & " SELECT * from " & fullTableName

20    Set DBCn = DBCheckConnection(DBCn)
30    Set DBRs = DBCheckRecordset(DBRs)

40    With DBRs
50        .CursorLocation = adUseServer ' adUseClient
60        .CursorType = adOpenDynamic ' adUseClient ' adOpenStatic ' adOpenDynamic ' adOpenForwardOnly
70        .LockType = adLockOptimistic  ' adLockReadOnly
80        Set .ActiveConnection = DBCn
90    End With

100   DBRs.Open qq, DBCn

110   Exit Sub
    
gotError:
    MsgBox Err.Number & " " & Err.Description & vbNewLine & vbNewLine & "Error on line: " & Erl, Title:="FastloadMerge"
    Stop
    Call DBCloseConnection(DBCn)
    Call DBCloseRecordset(DBRs)
    Resume Next
End Sub

Sub GetDate()
    
    On Error GoTo gotError

    output = Shell("dir", vbNormalFocus)
    
    Exit Sub
gotError:
    MsgBox Err.Number & " " & Err.Description & vbNewLine & vbNewLine & "Error on line: " & Erl, Title:="Fastload"
    Stop
    Resume Next
End Sub

Function CheckReservedWord(word)

    word = TrimReplace(word)
    ThisWorkbook.Worksheets("SQLReservedWords").Cells(1, 1) = word
    If Not IsError(ThisWorkbook.Worksheets("SQLReservedWords").Cells(2, 1)) Then
        CheckReservedWord = "a_" & word
    Else
        CheckReservedWord = word
    End If
    
End Function

Function IsReservedWord(word)

    ThisWorkbook.Worksheets("SQLReservedWords").Cells(1, 1) = word
    If Not IsError(ThisWorkbook.Worksheets("SQLReservedWords").Cells(2, 1)) Then
        IsReservedWord = True
    Else
        IsReservedWord = False
    End If
    
End Function

Sub FastLoadWrite(filePath, str)
Dim fso As Object
Dim oFile As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    '
    ' 1 - readonly
    ' 2 - writing
    ' 8 - append
    '
    ' 0 - Ascii format
    Set oFile = fso.OpenTextFile(filePath, 8, True, 0)
    
    oFile.WriteLine str
    oFile.Close

    Set fso = Nothing  ' for garbage collector
    Set oFile = Nothing

End Sub

Function CheckFastLoadTable()
Dim DBCn As ADODB.Connection
Dim DBRs As ADODB.Recordset

    Set DBCn = DBCheckConnection(DBCn)
    Set DBRs = DBCheckRecordset(DBRs)
    With DBRs
        .CursorLocation = adUseClient ' adUseServer, adUseClient
        .CursorType = adUseClient ' adUseClient ' adOpenStatic ' adOpenDynamic ' adOpenForwardOnly
        .LockType = adLockOptimistic ' adLockReadOnly
        Set .ActiveConnection = DBCn
    End With
    
    On Error GoTo gotError
        Debug_Print GLBUserQuery
        Set DBCn = DBCheckConnection(DBCn)
        If DBCn Is Nothing Then Exit Function
        Set DBRs = DBCheckRecordset(DBRs)
        
        'useQuery = "select count(_fl_id) from dl_oge_analytics.delete_me"
        'DBRs.Open useQuery
        
        'fieldCount = DBRs.Fields.count
        'For i = 1 To fieldCount
            Debug_Print DBRs.DataSource
gotError:
    MsgBox "DBQuery Error (" & Err.Number & "): " & Err.Description, vbOKOnly, Title:="DBQuery ERROR"
    Set DBCn = DBCheckConnection(DBCn)
    Set DBRs = DBCheckRecordset(DBRs)
End Function
'

Function FastLoad_Check(fullTableName, columnName, SHUse, WBUse)
'
' This function checks if the first line in the spreadsheet has already been loaded
' This function assumes the identifier used is UNIQUE
' This function needs to be customized for each use.

On Error GoTo gotError
    useCol = FindColumnHeader(columnName, SHUse, WBUse)
    With Workbooks(WBUse).Sheets(SHUse)
        useQuery = "select * from " & fullTableName & " where " & .Cells(1, useCol) & " = " & .Cells(2, useCol)
    End With
    
    Set DBCn = DBCheckConnection(DBCn)
    Set DBRs = DBCheckRecordset(DBRs)
    With DBRs
        .CursorLocation = adUseClient ' adUseServer, adUseClient
        .CursorType = adUseClient ' adUseClient ' adOpenStatic ' adOpenDynamic ' adOpenForwardOnly
        .LockType = adLockOptimistic ' adLockReadOnly
        Set .ActiveConnection = DBCn
    End With
    
    DBRs.Open useQuery
    
    If DBRs.State = 0 Then
        MsgBox ("ERROR FastLoad_Check - RecordSet")
        Call DBCloseConnection(DBCn)
        Call DBCloseRecordset(DBRs)
        Exit Function
    End If
    If DBRs.recordCount > 0 Then
        FastLoad_Check = True
    Else
        FastLoad_Check = False
    End If
    
    Call DBCloseConnection(DBCn)
    Call DBCloseRecordset(DBRs)
    Exit Function
    
gotError:
    MsgBox Err.Number & " " & Err.Description & vbNewLine & vbNewLine & "Error on line: " & Erl, Title:="Fastload_Check ERROR"
    'Set DBCn = DBCheckConnection(DBCn)
    'Set DBRs = DBCheckRecordset(DBRs)
    Resume
    Stop
End Function

Function FastLoad_ColumnNamesFormat(SHOrig, WBOrig, SHFmt, WBFmt) As Boolean

    FastLoad_ColumnNamesFormat = True

    rightCol = RowLastColumn(1, SHOrig, WBOrig)
    '
    ' Get Column Headers from spreadsheet
    '
    For i = 1 To rightCol
        Workbooks(WBFmt).Worksheets(SHFmt).Cells(1, i) = MyTrim(Workbooks(WBOrig).Worksheets(SHOrig).Cells(1, i))
        Workbooks(WBFmt).Worksheets(SHFmt).Cells(2, i) = Workbooks(WBOrig).Worksheets(SHOrig).Cells(2, i)
    Next i
    
    '-- Check For Reserved Words In Column Names --------------------------------------------------------------------------------
    For i = 1 To rightCol
        t = Workbooks(WBFmt).Worksheets(SHFmt).Cells(1, i)
        If t = "" Then
            emptyColumnCount = emptyColumnCount + 1
            Workbooks(WBFmt).Worksheets(SHFmt).Cells(1, i) = "EmptyColumn" & emptyColumnCount
        Else
            Workbooks(WBFmt).Worksheets(SHFmt).Cells(1, i) = CheckReservedWord(t)
        End If
    Next i
    
    '-- Check For Duplicate Names -----------------------------------------------------------------------
    For i = 1 To rightCol - 1
        t = Workbooks(WBFmt).Worksheets(SHFmt).Cells(1, i)
        Set sRange = Range(Workbooks(WBFmt).Worksheets(SHFmt).Cells(1, i), Workbooks(WBFmt).Worksheets(SHFmt).Cells(1, rightCol))
        Set fRange = FindInRangeExact(t, sRange)
        If Not fRange Is Nothing Then
            If fRange.Count > 1 Then
                Count = 0
                For Each f In fRange
                    Count = Count + 1
                    f.Value = f.Value & "_" & Count
                Next f
            End If
        End If
    Next i
        
    '-- Column Width MAX Total Length Is 64K -----------------------------------------------------------------------------
    If rightCol < 50 Then
        charWidth = 300
    Else
        charWidth = 100
    End If
    
    '-- Figure Out Column Format ----------------------------------------------------------------------------
    On Error Resume Next
    For j = 1 To rightCol
        useRow = 2 ' ColumnFirstRow(j, SHOrig, WBOrig, 2)
        t = Workbooks(WBFmt).Worksheets(SHFmt).Cells(useRow, j).text  ' get value
        useFmt = ""
        
        useFmt = ColumnDataType(t)
        
        If useFmt = "VARCHAR" Then useFmt = "VARCHAR(" & charWidth & ")"
        
        Workbooks(WBFmt).Worksheets(SHFmt).Cells(3, j) = useFmt
    Next j
    On Error GoTo 0

End Function
    '=============================================================================================================================================

'
' The routine will check the column names from the existing table against the spreadsheet to be uploaded
'
Function FastLoad_CheckColumnNames(tableName, SHOrig, WBOrig, SHFmt, WBFmt) As Boolean
'End Function

'Function Table_ColumnFormat(tableName, Optional SHUse, Optional WBUse)

On Error GoTo gotError
    ' tableName = "ppih_ssn"
    t = "select * from dbc.columnsv where databasename = 'dl_oge_analytics' and tablename = '" & tableName & "' order by columnid"
    
    t = "select TOP 1 * from pihpj." & tableName
    
    Call RunMyQuery(t, Range(Workbooks(WBFmt).Sheets(SHFmt).Cells(5, 1), Workbooks(WBFmt).Sheets(SHFmt).Cells(5, 1)), True)
    
    rightCol = RowLastColumn(5, SHFmt, WBFmt)
    
    If rightCol < 50 Then
        charWidth = 300
    Else
        charWidth = 100
    End If
    
    On Error Resume Next
    For j = 1 To rightCol
        useRow = 6 ' ColumnFirstRow(j, SHOrig, WBOrig, 2)
        t = Workbooks(WBFmt).Worksheets(SHFmt).Cells(useRow, j).text  ' get value
        useFmt = ""
        
        useFmt = ColumnDataType(t)
        
        If useFmt = "VARCHAR" Then useFmt = "VARCHAR(" & charWidth & ")"
        
        Workbooks(WBFmt).Worksheets(SHFmt).Cells(useRow + 1, j) = useFmt
    Next j
    
    FastLoad_CheckColumnNames = True
    
    For j = 1 To rightCol
        useRow = 6 ' ColumnFirstRow(j, SHOrig, WBOrig, 2)
        If UCase(Workbooks(WBFmt).Worksheets(SHFmt).Cells(useRow, j).text) <> _
            UCase(Workbooks(WBFmt).Worksheets(SHFmt).Cells(useRow, j).text) Then
                
                FastLoad_CheckColumnNames = False
        End If
    Next j
    On Error GoTo 0
    
    Exit Function
'------------------------------------------------------------------------------------
    Set DBCn = DBCheckConnection(DBCn)
    Set DBRs = DBCheckRecordset(DBRs)
    With DBRs
        .CursorLocation = adUseClient ' adUseServer, adUseClient
        .CursorType = adUseClient ' adUseClient ' adOpenStatic ' adOpenDynamic ' adOpenForwardOnly
        .LockType = adLockOptimistic ' adLockReadOnly
        Set .ActiveConnection = DBCn
    End With
    
    DBRs.Open t, DBCn
    
    If DBRs.State = 0 Then
        MsgBox ("ERROR FastLoad_Check - RecordSet")
        Call DBCloseConnection(DBCn)
        Call DBCloseRecordset(DBRs)
        Exit Function
    End If
    If DBRs.recordCount > 0 Then
        nCols = DBRs.recordCount
        DBRs.MoveNext  ' skip over LoadDate column
        mismatch = ""
        
        With Workbooks(WBFmt).Worksheets(SHFmt)
            For i = 1 To nCols - 1
                .Cells(4, i) = Trim(DBRs.Fields(2).Value)
                .Cells(5, i) = DBRs.Fields(3).Value
                If (.Cells(1, i) <> Workbooks(WBFmt).Sheets(SHFmt).Cells(4, i)) Then
                    t = Split(Columns(i).Address(False, False), ":")
                    mismatch = mismatch & "column " & t(0) & ":   " & ActiveSheet.Cells(1, i) & vbNewLine
                End If
                DBRs.MoveNext
            Next i
        End With
        
    Else
        'FastLoad_Check = False
    End If
    
    Call DBCloseConnection(DBCn)
    Call DBCloseRecordset(DBRs)
    
    'Workbooks(WBFmt).Close (False)
    
    If Len(mismatch) > 0 Then
        InputBox ("Mismatched Column Names:" & vbNewLine & vbNewLine & mismatch)
        Call ProgressFormAdd("FastLoad - Mismatch Column: " & mismatch)
        FastLoad_CheckColumnNames = False
        Exit Function
    End If
    
    FastLoad_CheckColumnNames = True
    Exit Function
    
gotError:
    MsgBox Err.Number & " " & Err.Description & vbNewLine & vbNewLine & "Error on line: " & Erl, Title:="Fastload_Check ERROR"
    'Set DBCn = DBCheckConnection(DBCn)
    'Set DBRs = DBCheckRecordset(DBRs)
    Resume
    Stop
    
End Function

Function FastLoad_CheckColumnNames2(fullTableName, SHOrig, WBOrig, SHFmt, WBFmt) As Boolean
' This function will check the EXISTING loaded table, without using system tables

'End Function

'Function Table_ColumnFormat(tableName, Optional SHUse, Optional WBUse)

On Error GoTo gotError
    ' tableName = "ppih_ssn"
    t = "select TOP 1 * from " & fullTableName 'dbc.columnsv where databasename = 'dl_oge_analytics' and tablename = '" & tableName & "' order by columnid"
    
    SHDest = "Destination"
    Workbooks(WBFmt).Sheets.Add.Name = SHDest
    Set outRange = Range(Workbooks(WBFmt).Sheets(SHDest).Cells(1, 1), Workbooks(WBFmt).Sheets(SHDest).Cells(1, 1))
    Call RunMyQuery(t, outRange, True)
    
    Call FastLoad_ColumnNamesFormat(SHDest, WBFmt, SHFmt, WBFmt)
    
    Set DBCn = DBCheckConnection(DBCn)
    Set DBRs = DBCheckRecordset(DBRs)
    With DBRs
        .CursorLocation = adUseClient ' adUseServer, adUseClient
        .CursorType = adUseClient ' adUseClient ' adOpenStatic ' adOpenDynamic ' adOpenForwardOnly
        .LockType = adLockOptimistic ' adLockReadOnly
        Set .ActiveConnection = DBCn
    End With
    
    DBRs.Open t, DBCn
    
    If DBRs.State = 0 Then
        MsgBox ("ERROR FastLoad_Check - RecordSet")
        Call DBCloseConnection(DBCn)
        Call DBCloseRecordset(DBRs)
        Exit Function
    End If
    If DBRs.recordCount > 0 Then
        nCols = DBRs.recordCount
        DBRs.MoveNext  ' skip over LoadDate column
        mismatch = ""
        
        With Workbooks(WBFmt).Worksheets(SHFmt)
            For i = 1 To nCols - 1
                .Cells(4, i) = Trim(DBRs.Fields(2).Value)
                .Cells(5, i) = DBRs.Fields(3).Value
                If (.Cells(1, i) <> ActiveSheet.Cells(4, i)) Then
                    t = Split(Columns(i).Address(False, False), ":")
                    mismatch = mismatch & "column " & t(0) & ":   " & ActiveSheet.Cells(1, i) & vbNewLine
                End If
                DBRs.MoveNext
            Next i
        End With
        
    Else
        'FastLoad_Check = False
    End If
    
    Call DBCloseConnection(DBCn)
    Call DBCloseRecordset(DBRs)
    
    'Workbooks(WBFmt).Close (False)
    
    If Len(mismatch) > 0 Then
        InputBox ("Mismatched Column Names:" & vbNewLine & vbNewLine & mismatch)
        Call ProgressFormAdd("FastLoad - Mismatch Column: " & mismatch)
        FastLoad_CheckColumnNames2 = False
        Exit Function
    End If
    
    FastLoad_CheckColumnNames2 = True
    Exit Function
    
gotError:
    MsgBox Err.Number & " " & Err.Description & vbNewLine & vbNewLine & "Error on line: " & Erl, Title:="Fastload_Check ERROR"
    'Set DBCn = DBCheckConnection(DBCn)
    'Set DBRs = DBCheckRecordset(DBRs)
    Resume
    Stop
    
End Function
</pre>
