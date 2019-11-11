

<pre>
Function HANA_CreateTable(fullTableName, SHFmt, WBFmt) As Boolean
Dim DBCn As ADODB.Connection
Dim DBRs As ADODB.Recordset

    Call ProgressFormAdd("FastLoad - Create Table")

On Error GoTo gotError
     Call DBCloseConnection(DBCn)
     
     rightCol = RowLastColumn(1, SHFmt, WBFmt)
     botRow = ColumnLastRow(1, SHFmt, WBFmt)
     
     useQuery = "create column table " & fullTableName & "( LoadDate date, "
     q = ""
     With Workbooks(WBFmt).Sheets(SHFmt)
        For i = 1 To rightCol
            q = q & .Cells(1, i) & " " & .Cells(3, i)
            If i < rightCol Then q = q & ", "
        Next i
    End With
    q = q & ")"
    useQuery = useQuery & q
    
20    Set DBCn = DBCheckConnection(DBCn)
30    Set DBRs = DBCheckRecordset(DBRs)

40    With DBRs
50        .CursorLocation = adUseServer ' adUseClient
60        .CursorType = adOpenDynamic ' adUseClient ' adOpenStatic ' adOpenDynamic ' adOpenForwardOnly
70        .LockType = adLockOptimistic  ' adLockReadOnly
80        Set .ActiveConnection = DBCn
90    End With

100   DBRs.Open useQuery, DBCn

        If DBRs.State = 0 Then
            HANA_CreateTable = True
        Else
            HANA_CreateTable = False
        End If
110   Exit Function
    
gotError:
    MsgBox Err.Number & " " & Err.Description & vbNewLine & vbNewLine & "Error on line: " & Erl, Title:="FastloadMerge"
    Stop
    Call DBCloseConnection(DBCn)
    Call DBCloseRecordset(DBRs)
    Resume Next
End Function


Function HANA_LoadData(fullTableName, SHData, WBData, SHFmt, WBFmt) As Boolean
Dim DBCn As ADODB.Connection
Dim DBRs As ADODB.Recordset

    Call ProgressFormAdd("FastLoad - Load Data")

On Error GoTo gotError
     Call DBCloseConnection(DBCn)
    
20    Set DBCn = DBCheckConnection(DBCn)
30    Set DBRs = DBCheckRecordset(DBRs)

40    With DBRs
50        .CursorLocation = adUseServer ' adUseClient
60        .CursorType = adOpenDynamic ' adUseClient ' adOpenStatic ' adOpenDynamic ' adOpenForwardOnly
70        .LockType = adLockOptimistic  ' adLockReadOnly
80        Set .ActiveConnection = DBCn
90    End With

     rightCol = RowLastColumn(1, SHData, WBData)
     botRow = ColumnLastRow(1, SHData, WBData)
     
     Debug.Print "botRow: " & botRow
     
     With Workbooks(WBData).Sheets(SHData)
        For i = 2 To botRow
            qq = "insert into " & fullTableName & " values( current_date, "
            For j = 1 To rightCol   ' a row of data
                t = Replace(.Cells(i, j), "'", "''")  ' for HANA handle single quotes
                If Workbooks(WBFmt).Sheets(SHFmt).Cells(3, j) = "TIMESTAMP" Then t = Mid(t, 1, InStr(t, ".") - 1)  ' modify the timestamp
                qq = qq & "'" & t & "'"
                If j < rightCol Then qq = qq & ", "
            Next j
            qq = qq & ");"
            
100         DBRs.Open qq, DBCn
        Next i
    End With
        HANA_LoadData = True
110   Exit Function
    
gotError:
    MsgBox Err.Number & " " & Err.Description & vbNewLine & vbNewLine & "Error on line: " & Erl, Title:="FastloadMerge"
    Stop
    Call DBCloseConnection(DBCn)
    Call DBCloseRecordset(DBRs)
    Resume Next
End Function
</pre>
