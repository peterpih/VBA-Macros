

<pre>
' November 19th, 2018
'
'
Function TDTableExists(tableName) As Boolean
Dim DBCn As ADODB.Connection
Dim DBRs As ADODB.Recordset

On Error GoTo gotError
10    qq = "SELECT TOP 1 * from " & tableName
20    Set DBCn = DBCheckConnection(DBCn)
30    Set DBRs = DBCheckRecordset(DBRs)

40    With DBRs
50        .CursorLocation = adUseClient ' adUseServer
60        .CursorType = adOpenDynamic ' adUseClient ' adOpenStatic ' adOpenDynamic ' adOpenForwardOnly
70        .LockType = adLockReadOnly ' adLockOptimistic
80        Set .ActiveConnection = DBCn
90    End With

On Error Resume Next

100   DBRs.Open qq, DBCn
    
    If DBRs.State = 0 Then
        TDTableExists = False
    Else
'110   fieldCount = DBRs.Fields.Count
      TDTableExists = True
120   Exit Function
    End If
    
    Exit Function

gotError:
    k = InStr(Err.Description, "does not exist")
    If k > 0 Then
        TDtablexists = False
        Exit Function
    End If
    MsgBox Err.Number & " " & Err.Description & vbNewLine & vbNewLine & "Error on line: " & Erl, Title:=" "
    Stop
    Resume Next
End Function

Function labelFillSpaces(s)
    labelFillSpaces = Replace(s, " ", "_")
End Function
'
' Lookup a Device Type for a Meter
'
Function TeradataLookup(meterId)
Dim DBCn As ADODB.Connection
Dim DBRs As ADODB.Recordset

    t = "Select EQUIP_MATERIAL_CODE from putlvw.EUL_POS_METERS_D WHERE EQUIP_MFG_SERIAL_NUMBER ='"
    t = t & meterId & "'"
    'debug_print t
    
    Set DBCn = DBCheckConnection(DBCn)
    Set DBRs = DBCheckRecordset(DBRs)
    
    With DBRs
        .CursorLocation = adUseServer ', adUseClient
        .CursorType = adOpenDynamic ' adUseClient ' adOpenStatic ' adOpenDynamic ' adOpenForwardOnly
        .LockType = adLockOptimistic ' adLockReadOnly
        Set .ActiveConnection = DBCn
    End With
    
    On Error GoTo gotError
    
    DBRs.Open t, DBCn
    
    fieldCount = DBRs.Fields.Count
    t = DBRs.Fields(0).Value
    TeradataLookup = Right(t, 7)
    Exit Function
    
gotError:
    MsgBox "DBQuery Error (" & Err.Number & "): " & Err.Description, vbOKOnly, Title:="DBQuery ERROR"
    DBGlbAdodbError = True
    Resume Next
    
End Function

'
' Lookup a Device Type for a Meter
'
Function TeradataLookup_new(useQuery, selectField, where1Field, Optional where2Field, Optional where3Field)
Dim DBCn As ADODB.Connection
Dim DBRs As ADODB.Recordset
Dim targetCol As Integer
Dim q As String, qq As String
Dim where1Col As Integer, where2Col As Integer, where3Col As Integer
    '
    ' rename old column used for SELECT
    '
    selectCol = FindColumnHeader(selectField)
    If selectCol > 0 Then
        Cells(1, selectCol) = Cells(1, selectCol) & "_old"
        Cells(1, selectCol).Interior.color = BLUE
    End If
    
    where1Col = FindColumnHeader(where1Field)
    newCol = ColumnInsertRight(where1Col)
    
    With Cells(1, newCol)
        .Value = selectField
        .Interior.color = LIGHTBLUE
        .Font.Bold = True
    End With
    
    If Not IsMissing(where2Field) Then where2Col = FindColumnHeader(where2Field)
    If Not IsMissing(where3Field) Then where3Col = FindColumnHeader(where3Field)
    
    botRow = ColumnLastRow(where1Col)
    
    On Error GoTo gotError
    For i = 2 To botRow
        
        Set DBCn = DBCheckConnection(DBCn)
        Set DBRs = DBCheckRecordset(DBRs)
    
        With DBRs
            .CursorLocation = adUseServer ', adUseClient
            .CursorType = adOpenDynamic ' adUseClient ' adOpenStatic ' adOpenDynamic ' adOpenForwardOnly
            .LockType = adLockOptimistic ' adLockReadOnly
            Set .ActiveConnection = DBCn
        End With
        
        qq = "SELECT " & selectField & " " & useQuery & " WHERE " & where1Field & " = " & Cells(i, where1Col)
        If where2Col > 0 Then qq = qq & " AND " & where2Field & " = " & Cells(i, where2Col)
        Debug_Print qq
        DBRs.Open qq, DBCn
    
        fieldCount = DBRs.Fields.Count
        If Not (DBRs.EOF And DBRs.BOF) Then
            t = DBRs.Fields(0).Value
            Cells(i, targetCol) = t
        Else
            Cells(i, targetCol) = "#N/A"
        End If
    Next i
    
gotError:
    MsgBox "TeradataLookup Error (" & Err.Number & "): " & Err.Description, vbOKOnly, Title:="TeradataLookup ERROR"
    DBGlbAdodbError = True
    Resume Next
    
End Function

Sub NewMoveOutDate()

    selectField = "move_out_date"
    where1Field = "INSTALLATION_NUMBER"
    where2Field = "METER_SERIAL_NUM"
    useQuery = "from putlvw.EUL_ACCOUNT_D"

    Call TeradataLookup_new(useQuery, selectField, where1Field, where2Field, where3Field)

End Sub
</pre>
