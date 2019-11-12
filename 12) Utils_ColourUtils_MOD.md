<h2ColorUtils_MOD</h2>

<pre>
' Standard Colors
'
Public currentColor As Long
'
'
Public Const RUST = 192
Public Const RED = 255
Public Const HILITERED = 393372
Public Const ORANGE = 49407
Public Const YELLOW = 65535
Public Const LIGHTGREEN = 5296274
Public Const GREEN = 5287936
Public Const LIGHTBLUE = 15773696
Public Const BLUE = 12611584
Public Const BLACK = 10
Public Const DARKBLUE = 6299648
Public Const PURPLE = 10498160
Public Const PINK = 13395711
Public Const LIGHTRUST = 6724095

Public Const NOCOLOR = 16777215
Public Const LIGHTPINK = 13421823 ' 0.599993896298105
Public Const HILITEPINK = 13551615

Public Const GREY = 9868950
Public Const LIGHTGREY = 14540253
Public Const GREYSPECKLE = 3
'
' Color handling for highlighting cells()
'
Sub ColorRange(useRange, Optional useColor)

    If useColor = APTCOLLAPSE Then
        With useRange.Interior
            .Pattern = xlPatternLinearGradient
            .Gradient.Degree = 270
            .Gradient.ColorStops.Clear
        End With
        With useRange.Interior.Gradient.ColorStops.Add(0)
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
        End With
        With useRange.Interior.Gradient.ColorStops.Add(1)
            .ThemeColor = xlThemeColorAccent1
            .TintAndShade = 0
        End With
    ElseIf useColor = GREYSPECKLE Then
        With useRange.Interior
            .Pattern = xlGray16
            .PatternColorIndex = xlAutomatic
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    ElseIf useColor = LIGHTGREY Then
        With useRange.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .color = useColor
            .ThemeColor = xlThemeColorDark2
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    Else '------------------------------------- regular color
        With useRange.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .color = useColor
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    End If
End Sub

Sub ColorHeader(headerName, useColor)
    Set sRange = Rows(1)
    Set fRange = FindInRange(headerName, sRange)
    Call ColorRange(fRange, useColor)
End Sub
'
' Return the next color in a sequence to ensure same color ot used consecutively
'
Function ChooseColor(color)
    Select Case color
        Case ORANGE
            ChooseColor = YELLOW
        Case YELLOW
            ChooseColor = LIGHTGREEN
        Case LIGHTGREEN
            ChooseColor = LIGHTBLUE
        Case LIGHTBLUE
            ChooseColor = ORANGE
        Case Else
            ChooseColor = YELLOW
    End Select
End Function

Function WhatColor()
    If ActiveCell.Interior.color = NOCOLOR Then
        WhatColor = ChooseColor(currentColor)
    Else
        WhatColor = ActiveCell.Interior.color
    End If
End Function

Sub LightUp()
    useColor = WhatColor()
    
    useCol = ActiveCell.Column
    botRow = ColumnLastRow(useCol)
    useValue = ActiveCell.Value
    Call ColorRange(Cells(ActiveCell.Row, useCol), useColor)
    For i = 2 To botRow
        If Cells(i, useCol) = useValue Then Call ColorRange(Cells(i, useCol), useColor)
    Next i
End Sub

Sub MakePallette()
    SHPallette = "Pallette"
    For j = 1 To 10
        For i = 1 To 10
            Call ColorRange(ThisWorkbook.Worksheets("Pallette").Cells(10 + j, i * j), i * j)
        Next i
    Next j
End Sub

Sub ColorWheel()
    useColor = PickNewColor(ActiveCell.Interior.color)
    Debug_Print "Color: " & useColor
    MsgBox "Color: " & useColor
    
End Sub

Sub ColorValue()
Dim vRange As Range, sRange As Range, fRange As Range
Dim useValue As String
Dim useColor As Long, botRow As Long

    On Error Resume Next
    Set vRange = Nothing
    Set vRange = Application.InputBox("Select Value", Title:="Color Value", Default:=Selection.Address, Type:=8)
    If vRange Is Nothing Then Exit Sub
    
    useValue = vRange.Text
    
    botRow = ColumnLastRow(vRange.Column)
    
    Set sRange = Range(Cells(2, vRange.Column), Cells(botRow, vRange.Column))
    
    useColor = PickNewColor(ActiveCell.Interior.color)
    Set fRange = ActiveCell
    fRange.Interior.color = useColor
    
    Call ScreenOff
    
    Set fRange = FindInRange(useValue, sRange)
    
    Call ScreenOn
    
    fRange.Interior.color = useColor
    
    MsgBox "Found " & fRange.Count & " with value of " & useValue & "."
    
    Set sRange = Nothing
    Set fRange = Nothing
End Sub

Function PickNewColor(Optional i_OldColor As Double = xlNone) As Double
Const BGColor As Long = 13160660  'background color of dialogue
Const ColorIndexLast As Long = 32 'index of last custom color in palette

Dim myOrgColor As Double          'original color of color index 32
Dim myNewColor As Double          'color that was picked in the dialogue
Dim myRGB_R As Integer            'RGB values of the color that will be
Dim myRGB_G As Integer            'displayed in the dialogue as
Dim myRGB_B As Integer            '"Current" color
  
  'save original palette color, because we don't really want to change it
  myOrgColor = ActiveWorkbook.Colors(ColorIndexLast)
  
  If i_OldColor = xlNone Then
    'get RGB values of background color, so the "Current" color looks empty
    Color2RGB BGColor, myRGB_R, myRGB_G, myRGB_B
  Else
    'get RGB values of i_OldColor
    Color2RGB i_OldColor, myRGB_R, myRGB_G, myRGB_B
  End If
  
  'call the color picker dialogue
  If Application.Dialogs(xlDialogEditColor).Show(ColorIndexLast, _
     myRGB_R, myRGB_G, myRGB_B) = True Then
    '"OK" was pressed, so Excel automatically changed the palette
    'read the new color from the palette
    PickNewColor = ActiveWorkbook.Colors(ColorIndexLast)
    'reset palette color to its original value
    ActiveWorkbook.Colors(ColorIndexLast) = myOrgColor
  Else
    '"Cancel" was pressed, palette wasn't changed
    'return old color (or xlNone if no color was passed to the function)
    PickNewColor = i_OldColor
  End If
End Function

'Converts a color to RGB values
Sub Color2RGB(ByVal i_Color As Long, _
              o_R As Integer, o_G As Integer, o_B As Integer)
  o_R = i_Color Mod 256
  i_Color = i_Color \ 256
  o_G = i_Color Mod 256
  i_Color = i_Color \ 256
  o_B = i_Color Mod 256
End Sub
</pre>
