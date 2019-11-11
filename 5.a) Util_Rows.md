

<pre>
' November 17th, 2018
'
'
Option Explicit

Sub CopyRangeRowsToSheet(rowRange As Range, SHTo As String)
Dim copyRange As Range, rw As Range
Dim SHFrom As String
Dim i As Integer, rightCol As Integer

    'rowRange.Copy
    'Worksheets(SHTo).Activate
    
    'Cells(2, 1).PasteSpecial xlPasteAll
    
    SHFrom = rowRange.Parent.Name
    
    i = 2
    For Each rw In rowRange
        rightCol = RowLastColumn(rw.Row, SHFrom)
        Set copyRange = Range(Worksheets(SHFrom).Cells(rw.Row, 1), Worksheets(SHFrom).Cells(rw.Row, rightCol))
        copyRange.Copy Destination:=Worksheets(SHTo).Cells(i, 1)
        i = i + 1
    Next rw
End Sub

Function AddRowNumbers(Optional useCol, Optional SHUse, Optional WBUse) As Integer
Dim useRange, numberRange As Range
Dim botRow As Long
Dim useHeader As String
Dim newCol As Long

    If IsMissing(SHUse) Then SHUse = ActiveSheet.Name
    If IsMissing(WBUse) Then WBUse = ActiveWorkbook.Name
    
    Call ScreenOff
    If IsMissing(useCol) Then
        On Error Resume Next
            Set numberRange = Nothing
            Set numberRange = Application.InputBox("Add Row Numbers for which column?", Title:="AddRowNumbers", Type:=8)
            On Error GoTo 0
            If numberRange Is Nothing Then Exit Function
        useCol = numberRange.Column
    End If
    
    With Workbooks(WBUse).Worksheets(SHUse)
        ' ordering of the next 3 steps is important
        botRow = ColumnLastRow(useCol, SHUse, WBUse)
        useHeader = .Cells(1, useCol).Value
    ' If (useCol = 1) Then
    '     useCol = 0
    '     useHeader = "ThisSheet"
    ' End If
    
        newCol = ColumnInsertRight(useCol, SHUse, WBUse)
    
        Set useRange = Range(.Cells(DATASTARTROW, newCol), .Cells(botRow, newCol))
        useRange.NumberFormat = "General"
        useRange.Formula = "=ROW()"
        Call RangeToValues(useRange)
        useRange.NumberFormat = "0"
        Workbooks(WBUse).Worksheets(SHUse).Cells(1, newCol).Value = useHeader & "-RowIndex"
        Call ColorRange(Workbooks(WBUse).Worksheets(SHUse).Cells(1, newCol), LIGHTGREEN)
    End With
    Call ScreenOn
    AddRowNumbers = newCol
End Function

Function NextRow(Optional SHUse, Optional WBUse) As Long
    If IsMissing(SHUse) Then SHUse = ActiveSheet.Name
    If IsMissing(WBUse) Then WBUse = ActiveWorkbook.Name
    
    NextRow = LastRow(SHUse, WBUse) + 1
End Function
'
' This function will return 0 if nothing there
'
Function LastRow2(Optional SHUse, Optional WBUse) As Long
Dim useRow As Long, useCol As Long
Dim t As Variant

On Error GoTo Err1:

    If (IsMissing(SHUse)) Then SHUse = ActiveSheet.Name
    If (IsMissing(WBUse)) Then WBUse = ActiveWorkbook.Name
    'LastRow = Workbooks(WBUse).Worksheets(SHUse).Cells.Find(What:="*", _
    '                after:=Workbooks(WBUse).Worksheets(SHUse).Cells(1, 1), _
    '                LookAt:=xlPart, _
    '                LookIn:=xlFormulas, _
    '                SearchOrder:=xlByRows, _
    '                SearchDirection:=xlPrevious, _
    '                MatchCase:=False).Row
    
    useRow = Workbooks(WBUse).Worksheets(SHUse).Cells.SpecialCells(xlCellTypeLastCell).Row
    useCol = Workbooks(WBUse).Worksheets(SHUse).Cells.SpecialCells(xlCellTypeLastCell).Column
    LastRow = useRow
    If useRow = 1 Then
        useCol = Workbooks(WBUse).Worksheets(SHUse).Cells.SpecialCells(xlCellTypeLastCell).Column
        If IsEmpty(Cells(useRow, useCol)) Then LastRow = 0
    End If
    Exit Function

Err1:
    LastRow = 0
    
End Function

Function LastRow(Optional SHUse, Optional WBUse) As Long

    If (IsMissing(SHUse)) Then SHUse = ActiveSheet.Name
    If (IsMissing(WBUse)) Then WBUse = ActiveWorkbook.Name

    If WorksheetFunction.CountA(Workbooks(WBUse).Sheets(SHUse).Cells) > 0 Then

        'Search for any entry, by searching backwards by Rows.

        LastRow = Workbooks(WBUse).Sheets(SHUse).Cells.Find(What:="*", After:=[A1], _
              SearchOrder:=xlByRows, _
              SearchDirection:=xlPrevious).Row
              
        'LastRow = Cells.SpecialCells(xlCellTypeLastCell).Row
    End If

End Function

Function RowLastColumn(Optional useRow, Optional SHUse, Optional WBUse) As Long

    If IsMissing(SHUse) Then SHUse = ActiveSheet.Name
    If IsMissing(WBUse) Then WBUse = ActiveWorkbook.Name
    
    With Workbooks(WBUse).Worksheets(SHUse)
        RowLastColumn = .Cells(useRow, .Columns.Count).End(xlToLeft).Column
        If (IsEmpty(.Cells(useRow, RowLastColumn))) Then RowLastColumn = 0
    End With
End Function

Function RowNextColumn(useRow As Long, Optional SHUse, Optional WBUse) As Long
    RowNextColumn = RowLastColumn(useRow, SHUse, WBUse) + 1
End Function
</pre>
