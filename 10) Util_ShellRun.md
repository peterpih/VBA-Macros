

<pre>

Sub currDir()
    MsgBox ShellRun("cmd.exe /c cd")
End Sub

Public Function ShellRun(sCmd As String) As String
    
    On Error GoTo gotError
    'Run a shell command, returning the output as a string'

      Dim oShell As Object
10    Set oShell = CreateObject("WScript.Shell")

    'run command'
      Dim oExec As Object
      Dim oOutput As Object
20    Set oExec = oShell.Exec(sCmd)
30    Set oOutput = oExec.StdOut

    'handle the results as they are written to and read from the StdOut object'
      Dim s As String
      Dim sLine As String
40    While Not oOutput.AtEndOfStream
50        sLine = oOutput.ReadLine
60        If sLine <> "" Then s = s & sLine & vbCrLf
70    Wend

80    ShellRun = s
      
      Set oExec = Nothing
      Set oShell = Nothing

      Exit Function
      
gotError:
    MsgBox Err.Number & " " & Err.Description & vbNewLine & vbNewLine & "Error on line: " & Erl, Title:="ShellRun"
    Stop
    Resume Next
      
End Function
</pre>
