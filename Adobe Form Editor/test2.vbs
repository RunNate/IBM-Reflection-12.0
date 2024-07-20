Dim objShell, strCommand, objExec, strOutput
Dim fieldNames, fieldValues

' Define fieldNames array
fieldNames = Array("First name", "Last name", "Date of birth", "Street name and number", _
                   "Suburb", "State", "Postcode", "Worker phone number", _
                   "Worker email address", "Key contact worker", "Key contact number", _
                   "Relationship of key contact", "Claim number", "Date of injury", _
                   "Interpreter required", "Contact made with worker")

' Define fieldValues array
fieldValues = Array("NATHAN", "DUNCOMBE", "3", "4", "5", "6", "7", "8", "9", "10", _
                    "11", "12", "080000000000", "11/11/1111", "$no", "$yes")

' Create a shell object
Set objShell = CreateObject("WScript.Shell")

' Define the command to execute
strCommand = "test.exe updateFields adl.pdf output.pdf "

' Append field names and values to the command
For i = 0 To UBound(fieldNames)
    strCommand = strCommand & Chr(34) & fieldNames(i) & Chr(34) & " "
Next

For i = 0 To UBound(fieldValues)
    strCommand = strCommand & Chr(34) & fieldValues(i) & Chr(34) & " "
Next

' Execute the command and capture the output
Set objExec = objShell.Exec(strCommand)
strOutput = objExec.StdOut.ReadAll()

' Display the output in a message box
MsgBox strOutput, vbInformation, "Update Fields Result"

' Clean up objects
Set objShell = Nothing
Set objExec = Nothing
