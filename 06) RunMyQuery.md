
<pre>
Function RunMyQuery(useQuery, outputRange, Optional showHeaders) As Boolean
Dim DBCn As ADODB.Connection
Dim DBRs As ADODB.Recordset

    If IsMissing(showHeaders) Then showHeaders = True
    
On Error GoTo gotError
20    Set DBCn = DBCheckConnection(DBCn)
30    Set DBRs = DBCheckRecordset(DBRs)

40    With DBRs
50        .CursorLocation = adUseClient ' adUseServer
60        .CursorType = adOpenDynamic ' adUseClient ' adOpenStatic ' adOpenDynamic ' adOpenForwardOnly
70        .LockType = adLockReadOnly ' adLockOptimistic
80        Set .ActiveConnection = DBCn
90    End With

100   DBRs.Open useQuery, DBCn
    
        If DBRs.State = 0 Then   ' closed
            If InStr(UCase(useQuery), "INSERT") > 0 Or InStr(UCase(useQuery), "UPDATE") > 0 Then
                RunMyQuery = True
                Exit Function
            Else
                GoTo gotError
            End If
        End If
        
        Offset = 0
        If showHeaders Then Offset = 1
        recordCount = DBRs.recordCount
110     fieldCount = DBRs.Fields.Count

        If showHeaders Then
            For j = 0 To fieldCount - 1
                If (i = 0) And showHeaders Then outputRange.Offset(0, j) = DBRs.Fields(j).Name
            Next j
        End If
        
        For i = 0 To recordCount - 1
            For j = 0 To fieldCount - 1
                'If (i = 0) And showHeaders Then outputRange.Offset(i, j) = DBRs.Fields(j).Name
                 outputRange.Offset(i + Offset, j) = DBRs.Fields(j).Value
            Next j
            DBRs.MoveNext
        Next i

        Call DBCloseRecordset(DBRs)
        Call DBCloseConnection(DBCn)
120   Exit Function

gotError:
    k = InStr(Err.Description, "does not exist")
    If k > 0 Then
        TDtablexists = False
        Exit Function
    End If
    MsgBox Err.Number & " " & Err.Description & vbNewLine & vbNewLine & "Error on line: " & Erl, Title:="RunMyQuery"
    Stop
    Resume Next

End Function
</pre>
