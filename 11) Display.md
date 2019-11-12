

<pre>
' June 1, 2019 - added AlertsOff, AlertsOn
' November 17th, 2018 - created
'
'
Function AlertsOff()
    Application.DisplayAlerts = False
End Function
Function AlertsOn()
    Application.DisplayAlerts = True
End Function
Function ScreenOff_del()
    Application.ScreenUpdating = False
End Function

Function ScreenOn_del()
    Application.ScreenUpdating = True
End Function

Function CalculationOff()
    Application.Calculation = xlManual
End Function

Function CalculationOn()
    Application.Calculation = xlAutomatic
End Function

Function StatusbarDisplay(Optional s)
    Application.DisplayStatusBar = True
    If IsMissing(s) Then s = "testing..."
        Application.StatusBar = s
        'DoEvents
End Function
</pre>
