<h2>Util_Columns</h2>

<pre>
Public Const EXCELMAXROW = 1048576

Function <b>ColumnAppend</b>(Optional fromRange, Optional toRange, Optional SHUse, Optional WBUse) As Boolean
Dim aRange As Range
Dim botRowFrom As Long, nextRowTo As Long
Dim fromCol As Integer, toCol As Integer

    'If IsMissing(SHUse) Then SHUse


    On Error Resume Next
    If IsMissing(fromRange) Then
        Set fromRange = Nothing
        Set fromRange = Application.InputBox("Select Column To Append", Default:=Selection.Address, Type:=8)
        If fromRange Is Nothing Then Exit Function
    End If
    
    If IsMissing(toRange) Then
        Set toRange = Nothing
        Set toRange = Application.InputBox("Select Column To Append To", Default:=Selection.Address, Type:=8)
        If toRange Is Nothing Then Exit Function
    End If
    
    
    WBFrom = fromRange.Parent.Parent.Name
    SHFrom = fromRange.Parent.Name
    fromCol = fromRange.Column
    
    WBTo = toRange.Parent.Parent.Name
    SHTo = toRange.Parent.Name
    toCol = toRange.Column
    
    botRowFrom = ColumnLastRow(fromRange.Column, SHFrom, WBFrom)
    nextRowTo = ColumnNextRow(toRange.Column, SHTo, WBTo)
    
    Set aRange = Range(Workbooks(WBFrom).Worksheets(SHFrom).Cells(2, fromCol), _
                  Workbooks(WBFrom).Worksheets(SHFrom).Cells(botRowFrom, fromCol))
    aRange.Copy Destination:=Workbooks(WBTo).Worksheets(SHTo).Cells(nextRowTo, toCol)
    
    ColumnAppend = True
End Function

Function <b>ColumnsSetDefaultWidth</b>(Optional useWidth, Optional SHUse, Optional WBUse) As Boolean

    If IsMissing(SHUse) Then SHUse = GetActiveSheet
    If IsMissing(WBUse) Then WBUse = GetActiveWorkbook

    If IsMissing(useWidth) Then useWidth = 8.43
    Call ColumnAutoWidthMax(maxWidth:=useWidth)

    ColumnsSetDefaultWidth = True
End Function

Function <b>ColumnAutoWidthMax</b>(Optional useCol, Optional maxWidth, Optional SHUse, Optional WBUse) As Boolean

    If IsMissing(SHUse) Then SHUse = GetActiveSheet
    If IsMissing(WBUse) Then SHUse = GetActiveWorkbook
    
    If IsMissing(maxWidth) Then maxWidth = 30
    
    If IsMissing(useCol) Then ' assume it's for the entire sheet
        Workbooks(WBUse).Worksheets(SHUse).Columns.AutoFit
        rightCol = LastColumn(SHUse, WBUse)
        For i = 1 To rightCol
            qw = Workbooks(SBUse).Worksheets(SHUse).Columns(i).ColumnWidth
            If qw > maxWidth Then Workbooks(WBUse).Worksheets(SHUse).Columns(i).ColumnWidth = maxWidth
        Next i
    Else
        Workbooks(WBUse).Worksheets(SHUse).Columns(useCol).AutoFit
        qw = Workbooks(WBUse).Worksheets(SHUse).Columns(useCol).ColumnWidth
        If qw > maxWidth Then Workbooks(WBUse).Worksheets(SHUse).Columns(useCol).ColumnWidth = maxWidth
    End If
    
    ColumnAutoWidthMax = True
End Function

Function <b>ColumnSetAlignment</b>(headerName, Optional useAlign, Optional SHUse, Optional WBUse) As Boolean
    
    If IsMissing(SHUse) Then SHUse = GetActiveSheet
    If IsMissing(WBUse) Then SHUse = GetActiveWorkbook
    
    If IsMissing(useAlign) Then useAlign = xlRight
    
    useCol = FindColumnHeader(headerName)
    Worksheets(SHUse).Columns(useCol).HorizontalAlignment = useAlign
    
    ColumnSetAlignment = True
End Function

Function <b>ColumnInsertRight</b>(useCol, Optional SHUse, Optional WBUse) As Integer

    If IsMissing(SHUse) Then SHUse = GetActiveSheet
    If IsMissing(WBUse) Then WBUse = GetActiveWorkbook
    
    Call ClearClipboard
    Call CalculationOff
    Workbooks(WBUse).Worksheets(SHUse).Columns(useCol + 1).Insert Shift:=xlToRight
    Columns(useCol + 1).NumberFormat = "General"
    ColumnInsertRight = useCol + 1
    Call CalculationOn
    
    ColumnInsertRight = useCol + 1
End Function

Function <b>ColumnInsertLeft</b>(useCol, Optional SHUse, Optional WBUse) As Integer

    If IsMissing(SHUse) Then SHUse = GetActiveSheet
    If IsMissing(WBUse) Then WBUse = GetActiveWorkbook
    
    Call ClearClipboard
    Call CalculationOff
    Workbooks(WBUse).Worksheets(SHUse).Columns(useCol).Insert Shift:=xlToRight
    Columns(useCol).NumberFormat = "General"
    ColumnInsertLeft = useCol
    Call CalculationOn
End Function

Function ColumnLastRow_test()
    Call NewActiveWorkbookandSheet
    
    ActiveSheet.Range("A1") = 1
    ActiveSheet.Range("A2") = 2
    ActiveSheet.Range("A3") = 3
    ActiveSheet.Range("A4") = ""
    
    ActiveSheet.Range("B1") = ""
    ActiveSheet.Range("B2") = 2
    ActiveSheet.Range("B3") = 3
    ActiveSheet.Range("B4") = ""
    
    t = ColumnLastRow(1)
    t = ColumnLastRow(2)
    
    Call DeleteActiveWorkbook

End Function

Function <b>ColumnLastRow</b>(Optional useCol, Optional SHUse, Optional WBUse) As Long

10  On Error GoTo gotError

20  If IsMissing(SHUse) Then SHUse = GetActiveSheet
30  If IsMissing(WBUse) Then WBUse = GetActiveWorkbook

40  If IsMissing(useCol) Then useCol = 1
    
50  topRow = ColumnFirstRow(useCol, SHUse, WBUse)

    If Not IsEmpty(Workbooks(WBUse).Worksheets(SHUse).Cells(topRow + 1, useCol)) Then
    
60      ColumnLastRow = Workbooks(WBUse).Worksheets(SHUse).Cells(topRow, useCol).End(xlDown).Row
    End If

70  k = LastRow(SHUse, WBUse)
    
80  With Workbooks(WBUse).Worksheets(SHUse)
.Cells(k + 1, useCol).Copy
100      ColumnLastRow = .Cells(k + 1, useCol).End(xlUp).Row
110      If IsEmpty(.Cells(ColumnLastRow, useCol)) Then ColumnLastRow = 0
120  End With

130 Exit Function

gotError:
    MsgBox Err.Number & " " & Err.Description & vbNewLine & vbNewLine & "Error on line: " & Erl, Title:="ColumnLastRow"
    Stop
    Resume Next
End Function

Function <b>ColumnNextRow</b>(Optional useCol, Optional SHUse, Optional WBUse) As Long
    ColumnNextRow = ColumnLastRow(useCol, SHUse, WBUse) + 1
End Function

Function <b>ColumnFirstRow</b>(Optional useCol, Optional SHUse, Optional WBUse, Optional startRow) As Long
Dim useRow, botRow As Long

    If IsMissing(SHUse) Then SHUse = GetActiveSheet
    If IsMissing(WBUse) Then WBUse = GetActiveWorkbook
    
    If IsMissing(startRow) Then startRow = 2 ' assumes header row
    
    If IsEmpty(Workbooks(WBUse).Sheets(SHUse).Cells(startRow, useCol)) Then
        useRow = Workbooks(WBUse).Sheets(SHUse).Cells(startRow, useCol).End(xlDown).Row
        If useRow = EXCELMAXROW Or startRow = 2 Then useRow = 2
    Else
        useRow = startRow
    End If
    
    ColumnFirstRow = useRow
End Function

Function <b>ColumnDataRange</b>(useCol, Optional SHUse, Optional WBUse) As Range
Dim botRow As Long

    Call GetActiveWorkbookAndSheet(SHUse, WBUse)
    
    botRow = ColumnLastRow(useCol, SHUse, WBUse)
    Set ColumnDataRange = Range(Cells(2, useCol), Cells(botRow, useCol))
End Function

Function <b>ColumnCountA</b>(useCol)
    Set aRange = Range(Cells(2, useCol), Cells(ColumnLastRow(useCol), useCol))
    ColumnCountA = WorksheetFunction.CountA(aRange)
    ColumnCountA = ActiveSheet.Columns(useCol).Cells.SpecialCells(xlCellTypeConstants).Count
End Function
'
' similar to filterMultipleWorkOrders
'
Sub <b>ColumnCountValues</b>(Optional useCol)
Dim colRange As Range
Dim botRow As Long, i As Long, k As Long

    If IsMissing(useCol) Then useCol = ActiveCell.Column
    Set colRange = Range(Cells(1, useCol), Cells(1, useCol))
    
    Call SortSheetUp(colRange.Column)
    
    countCol = ColumnInsertLeft(colRange.Column)
    Cells(1, countCol) = "Count of " & Cells(1, colRange.Column)
    
    botRow = ColumnLastRow(colRange.Column) + 1
    
    i = 2
    k = 1
    While i < botRow
        While (Cells(i, colRange.Column) = Cells(i + 1, colRange.Column))
            k = k + 1
            Rows(i + 1).Delete
            botRow = botRow - 1
        Wend
        Cells(i, countCol) = k
        
        i = i + 1
        k = 1
    Wend

End Sub

' assumes there is a header row at top
Function <b>ExtractColumnToSheet</b>(SHTo, COLNAME, Optional SHFrom)

    If IsMissing(SHFrom) Then SHFrom = ActiveSheet.Name
    
    ' find column
    ''Set headerRange = Worksheets(SHFrom).Rows(1) ' for finding columns by name
    
    ''Set aRange = FindInRangeExact(colName, headerRange)
    fromCol = HeaderToColumnNum(COLNAME, SHFrom)
    
    toCol = NextColumn(SHTo)
    Set fromRange = Worksheets(SHFrom).Columns(fromCol)
    fromRange.Copy
    
    Set toRange = Worksheets(SHTo).Cells(1, toCol)
    toRange.PasteSpecial Paste:=xlAll
    
End Function

Function <b>HeaderToColumnNum</b>(useHeader, Optional SHUse, Optional WBUse)
On Error GoTo None:

    If IsMissing(SHUse) Then SHUse = GetActiveSheet
    If IsMissing(WBUse) Then WBUse = GetActiveWorkbook
    
    Set headerRange = Workbooks(WBUse).Worksheets(SHUse).Rows(1) ' for finding columns by name
    Set aRange = FindInRangeExact(useHeader, headerRange)
    HeaderToColumnNum = aRange.Column
    Exit Function
    
None:
    HeaderToColumnNum = -1
    
End Function

Function <b>ColumnNumToLetter</b>(iCol) As String
   Dim iAlpha As Integer
   Dim iRemainder As Integer
   
   t = Cells(1, iCol).Address(False, False, xlA1)
   ColumnNumToLetter = Left(t, Len(t) - 1)

End Function

Function IsColumnEmpty(colNum, Optional SHUse, Optional WBUse) As Boolean

    On Error GoTo gotError
    
10  If IsMissing(SHUse) Then SHUse = GetActiveSheet
20  If IsMissing(WBUse) Then WBUse = GetActiveWorkbook

30  IsColumnEmpty = False
40  If ColumnLastRow(colNum, SHUse, WBUse) <= 1 Then IsColumnEmpty = True

    Exit Function
gotError:
    MsgBox Err.Number & " " & Err.Description & vbNewLine & vbNewLine & "Error on line: " & Erl, Title:="IsColumnEmpty"
    Stop
    Resume Next
End Function

Function <b>NextColumn</b>(Optional SHUse, Optional WBUse)
On Error GoTo Err1:
    NextColumn = LastColumn(SHUse, WBUse) + 1
    Exit Function
Err1:
    LastColumn = 0
    
End Function

Function <b>LastColumn</b>(Optional SHUse, Optional WBUse)
On Error GoTo Err1:

    If IsMissing(SHUse) Then SHUse = ActiveSheet.Name
    If IsMissing(WBUse) Then WBUse = ActiveWorkbook.Name
    
    LastColumn = Workbooks(WBUse).Sheets(SHUse).Cells.Find(What:="*", _
                    After:=Workbooks(WBUse).Worksheets(SHUse).Range("A1"), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByColumns, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False).Column
    Exit Function
Err1:
    LastColumn = 0
End Function

Function <b>ThisColumnRange</b>()
Dim colRange As Range

    useCol = ActiveCell.Column
    Set colRange = Range(Cells(1, useCol), Cells(ColumnLastRow, useCol))
    Set ThisColumnRange = colRange
End Function

Sub <b>ColumnConCat</b>(Optional colRange, Optional toRange)
Dim aRange As Range
Dim i As Integer, j As Integer
Dim fff As String
Dim botRow As Long

    If IsMissing(colRange) Then
        On Error Resume Next
        Set colRange = Nothing
        Set colRange = Application.InputBox("Select Columns To Concatenate", Default:=Selection.Address, Title:="ColConCat", Type:=8)
        If colRange Is Nothing Then Exit Sub
    End If
    
    If IsMissing(toRange) Then
        Set toRange = Nothing
        Set toRange = Application.InputBox("Select Destination", Title:="ColConCat", Type:=8)
        If toRange Is Nothing Then Exit Sub
    End If
        
    botRow = LastRow()
    
    fff = "="
    For i = 1 To colRange.Areas.Count
        For j = 1 To colRange.Areas(i).Columns.Count
            fff = fff & Left(colRange.Areas(i).Columns(j).End(xlUp).Address(False, False), 1) & "2" & "&"
        Next j
    Next i
    
    fff = Left(fff, Len(fff) - 1)
    
    useCol = toRange.Column
    newCol = ColumnInsertLeft(toRange.Column)
    
    Set aRange = Range(Cells(2, newCol), Cells(botRow, newCol))
    
    aRange.Formula = fff
    
    Call RangeToValues(aRange)
    
            
End Sub

Function ColumnNumToLet_test()
    t1 = ColumnNumToLet(44)
    t2 = ColumnNumToLet(3)
    If t1 = "AR" And t2 = "C" Then Debug.Print "Success"
End Function
Function <b>ColumnNumToLet(n)</b> As String
    t = Cells(1, n).Address
    k = InStr(2, t, "$")
    ColumnNumToLet = Mid(t, 2, k - 2)
End Function
</pre>
