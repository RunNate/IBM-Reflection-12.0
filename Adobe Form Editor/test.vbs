Dim objShell, strCommand, objExec, strOutput

' Create a shell object
Set objShell = CreateObject("WScript.Shell")

' Define the command to execute
strCommand = "test.exe printFieldNames adl.pdf"

' Execute the command and capture the output
Set objExec = objShell.Exec(strCommand)
strOutput = objExec.StdOut.ReadAll()

' Display the output in a message box
MsgBox strOutput, vbInformation, "Field Names"

' Clean up objects
Set objShell = Nothing
Set objExec = Nothing