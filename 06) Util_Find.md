

<pre>
' Novemeber 17th, 2018
'

Function FindInRangeExact(searchValue, searchRange)
    Set FindInRangeExact = FindInRange(searchValue, searchRange, xlWhole)
End Function

Function FindInRange(searchValue, searchRange, Optional useLookAt) As Range
Dim startRange As Range, resultRange As Range
Dim findRange As Range
Dim SHSearch As String, WBSearch As String

On Error GoTo gotError

10  If (IsMissing(useLookAt)) Then useLookAt = xlPart

20  SHSearch = searchRange.Parent.Name
30  WBSearch = searchRange.Parent.Parent.Name
    
    ' following line incase searchValue is in the first cell of the searchRange
40  Set startRange = searchRange.Cells(searchRange.Rows.Count, searchRange.Columns.Count)
    '
    '  LookAt:= xlWhole, xlPart
    '  SearchOrder:= xlByRows, xlByColumns
    '  SearchDirection:= xlNext, xlPrevious
    '  searchformat:=  True, False
    '
50  Set resultRange = searchRange.Find(What:=searchValue, _
                    After:=startRange, _
                    LookAt:=useLookAt, _
                    LookIn:=xlValues, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlNext, _
                    searchformat:=False, _
                    MatchCase:=False)
60  If Not (resultRange Is Nothing) Then
70      startAddress = resultRange.Address(True, True, xlA1)
80      Set findRange = resultRange
90      Do
            'debug_print findRange.Address
         'Set findRange = Workbooks(WBSearch).Worksheets(SHSearch).FindNext(after:=findRange)
100         Set findRange = searchRange.FindNext(After:=findRange)

            t = findRange.Address(True, True, xlA1)
110            If (findRange.Address(True, True, xlA1) = startAddress) Then Exit Do
            'findRange.Copy
120          Set resultRange = MyUnion(resultRange, findRange)
130      Loop
140  End If
150 Set FindInRange = resultRange
160 Set resultRange = Nothing
170 Set startRange = Nothing
    Exit Function
gotError:
     MsgBox Err.Number & " " & Err.Description & vbNewLine & vbNewLine & "Error on line: " & Erl, Title:="FindInRange"
    Stop
    Resume Next
End Function

Function FindColumnHeader(columnName, Optional SHUse, Optional WBUse, Optional headerRow) As Long
Dim sRange As Range, fRange As Range

    If IsMissing(SHUse) Then SHUse = ActiveSheet.Name
    If IsMissing(WBUse) Then WBUse = ActiveWorkbook.Name
    If IsMissing(headerRow) Then headerRow = 1
    '
    ' expanded since top row may have the FullTableName
    '
    With Workbooks(WBUse).Worksheets(SHUse)
        Set sRange = Range(.Rows(headerRow), .Rows(headerRow))
        'sRange.Copy
        Set fRange = FindInRange(columnName, sRange)
        If fRange Is Nothing Then       ' may be 2 row header
            Set sRange = Range(.Rows(headerRow), .Rows(headerRow + 1))
            'sRange.Copy
            Set fRange = FindInRange(columnName, sRange)
        End If
    End With
    
    If Not fRange Is Nothing Then
        FindColumnHeader = fRange(1).Column  ' get the first one
        Exit Function
    Else
        FindColumnHeader = -1
    End If
    
    Set sRange = Nothing
    Set fRange = Nothing
    
End Function

Function FindRangeErrors(useRange As Range) As Range
Dim aRange As Range, bRange As Range

    On Error Resume Next
    Set aRange = Nothing
    Set aRange = useRange.SpecialCells(xlCellTypeConstants, xlErrors)
    Set bRange = Nothing
    Set bRange = useRange.SpecialCells(xlCellTypeFormulas, xlErrors)
    Set FindRangeErrors = MyUnion(aRange, bRange)
    If IsEmpty(FindRangeErrors) Then Set FindRangeErrors = Nothing
    
    Set aRange = Nothing
    Set bRange = Nothing
    
End Function

Function FindRangeNonEmpty(searchRange)
    
    On Error Resume Next
    Set aRange = Nothing
    Set aRange = searchRange.SpecialCells(xlCellTypeConstants)
    Set bRange = Nothing
    Set bRange = searchRange.SpecialCells(xlCellTypeFormulas)
    Set FindRangeNonEmpty = MyUnion(aRange, bRange)
    If IsEmpty(FindRangeNonEmpty) Then Set FindRangeNonEmpty = Nothing
    
    Set aRange = Nothing
    Set bRange = Nothing
End Function
</pre>
