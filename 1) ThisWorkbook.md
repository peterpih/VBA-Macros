<h2>ThisWorkbook</h2>

This code gets placed in ThisWorkbook  
Dependency: Store Retrieve rountines

<pre>
'
' Novemeber 11, 2019
'
'-----------------------------------------------------------------------------
' Update target Workbook / Worksheet
'
Private Sub Workbook_Activate()
    On Error Resume Next
    WBUse = Application.Windows(2).Caption
    Call SetActiveWorkbook(Application.Windows(2).Caption)
    Call SetActiveSheet(Workbooks(WBUse).ActiveSheet.Name)
End Sub

Private Sub Workbook_Deactivate()
    On Error Resume Next
    Call SetActiveWorkbook(ActiveWorkbook.Name)
    Call SetActiveSheet(Workbooks(ActiveWorkbook.Name).ActiveSheet.Name)
End Sub
'------------------------------------------------------------------------------
' Erase password on Save / Close
'
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ThisWorkbook.Sheets("TopSheet").Cells(1, 1) = "" ' erase password
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    ThisWorkbook.Sheets("TopSheet").Cells(1, 1) = "" ' erase password
End Sub
</pre>
