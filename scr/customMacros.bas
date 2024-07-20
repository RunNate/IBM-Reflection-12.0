Attribute VB_Name = "customMacros"
Sub openFormGenerator()
    NathansFormGenerator.Show
End Sub

Sub pasteClaimNumber()
    Dim screen, text, objHTML
    
    Set objHTML = CreateObject("htmlfile")
    text = Replace(Replace(Trim(objHTML.ParentWindow.ClipboardData.GetData("text")), " ", ""), vbCrLf, "")

    With Session
        ECL01
        screen = .GetDisplayText(2, 12, 5)
        If screen = "ECL01" Then
            Numeric = True
            For i = 1 To Len(text)
                If Not IsNumeric(Mid(text, i, 1)) Then
                    Numeric = False
                    Exit For
                End If
            Next i
        
            If Len(text) = 11 And Numeric = True Then
                .TransmitTerminalKey rcIBMHomeKey
                .TransmitANSI Mid(text, 3, 2)
                .TransmitANSI Mid(text, 5, 7)
                .TransmitANSI Mid(text, 1, 2)
                .TransmitTerminalKey rcIBMEnterKey
                .WaitForEvent rcKbdEnabled, "30", "0", 1, 1
                .TransmitTerminalKey rcIBMBacktabKey
            End If
        End If
    End With
    
    Set objHTML = Nothing
End Sub

Sub copyClaimNumber()
    Dim screen, claim, temp, tempx, count, objData, text
    Set objData = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    
    With Session
        ECL01
        screen = .GetDisplayText(2, 12, 5)
        If screen = "ECL01" Then
            temp = .GetDisplayText(5, 20, 7)
            For tempx = 1 To Len(.GetDisplayText(5, 20, 7))
                If Mid(.GetDisplayText(5, 20, 7), tempx, 1) = "_" Then
                    count = count + 1
                    temp = "0" & temp
                End If
            Next
            temp = Replace(temp, "_", "")
            claim = .GetDisplayText(5, 28, 2) & .GetDisplayText(5, 17, 2) & temp
            objData.SetText claim
            objData.PutInClipboard
        End If
        .TransmitTerminalKey rcIBMHomeKey
        .TransmitTerminalKey rcIBMBacktabKey
    End With
    
    Set objData = Nothing
End Sub

Sub createTaxiRequestForm()
    Dim workerDetails, cmDetails
    workerDetails = getWorkerDetails()
    cmDetails = getCMDetails()
    
    If Not IsEmpty(workerDetails) Then
        replaceFields = Array( _
            Array("{todayDate}", formatDate(Now)), _
            Array("{workerName}", StrConv(workerDetails(0) & " " & workerDetails(1), vbProperCase)), _
            Array("{claimNumber}", workerDetails(8)), _
            Array("{workerPhone}", workerDetails(7)), _
            Array("{addressLine1}", workerDetails(3)), _
            Array("{addressLine2}", workerDetails(4) & " " & workerDetails(5) & " " & workerDetails(6)), _
            Array("{caseManagerName}", cmDetails(1)), _
            Array("{caseManagerPhone}", cmDetails(2)), _
            Array("{caseManagerEmail}", cmDetails(3)) _
        )
        
        Set fso = CreateObject("Scripting.FileSystemObject")
        filePath = "G:\Macros\letters\13CABS Taxi Travel Request Template.docm"
        savePath = "H:\13CABS Taxi Travel Request " & workerDetails(8) & ".docm"
        Call wordDocumentFindReplace(replaceFields, filePath, savePath)
        Set fso = Nothing
    Else
        MsgBox "Error locating Information, please ensure you are logged in on a valid screen."
    End If
End Sub

Sub createADLForm()
    Dim workerDetails, cmDetails
    workerDetails = getWorkerDetails()
    cmDetails = getCMDetails()
    
    If Not IsEmpty(workerDetails) Then
        fieldNames = Array("First name", "Last name", "Date of birth", "Street name and number", "Suburb", "State", "Postcode", "Worker phone number", "Worker email address", "Key contact worker", "Key contact number", "Relationship of key contact", "Claim number", "Date of injury", "Contact made with worker", "Accepted compensable injuries", "Case manager", "Case manager number", "Case manager email")
        fieldValues = Array(workerDetails(0), workerDetails(1), workerDetails(2), workerDetails(3), workerDetails(4), workerDetails(5), workerDetails(6), workerDetails(7), workerDetails(9), "N/A", "N/A", "N/A", workerDetails(8), workerDetails(11), "$yes", workerDetails(12), cmDetails(1), cmDetails(2), cmDetails(3))
        strCommand = """G:\Macros\Adobe Form Editor\test.exe"" updateFields ""G:\Macros\Adobe Form Editor\adl.pdf"" ""H:\ADL Request Form " & workerDetails(8) & ".pdf"" "
            
        For i = 0 To UBound(fieldNames)
            strCommand = strCommand & Chr(34) & fieldNames(i) & Chr(34) & " "
        Next
        For i = 0 To UBound(fieldValues)
            strCommand = strCommand & Chr(34) & fieldValues(i) & Chr(34) & " "
        Next
        
        Set objShell = CreateObject("WScript.Shell")
        Set objExec = objShell.Exec(strCommand)
        strOutput = objExec.StdOut.ReadAll()
        
        If InStr(strOutput, "updated successfully") Then
            objShell.Run """H:\ADL Request Form " & workerDetails(8) & ".pdf""", 1, False
        Else
            MsgBox strOutput, vbInformation, "Update Fields Result"
        End If
            
        Set objShell = Nothing
        Set objExec = Nothing
    Else
        MsgBox "Error locating Information, please ensure you are logged in on a valid screen."
    End If
End Sub

Function wordDocumentFindReplace(replaceFields, filePath, saveLocation)
    On Error Resume Next
    Set wordApp = GetObject(, "Word.Application")

    If Err.Number <> 0 Then
        Err.Clear
        Set wordApp = CreateObject("Word.Application")
        wordApp.Visible = False
    End If

    Set wordDoc = wordApp.Documents.Open(filePath)
    If Err.Number <> 0 Then
        MsgBox "Failed to open document: " & filePath & vbCrLf & "Error: " & Err.Description, vbCritical
        wordApp.Quit
        Set wordApp = Nothing
        WScript.Quit
    End If
    On Error GoTo 0
    
    For i = 0 To UBound(replaceFields)
        With wordDoc.Content.Find
            .Execute replaceFields(i)(0), False, False, False, False, False, True, 1, False, replaceFields(i)(1), 2
        End With
    Next

    wordApp.Visible = True

    If saveLocation <> "" Then
        wordDoc.SaveAs saveLocation
    End If
    
    wordApp.Activate

    Set wordDoc = Nothing
    Set wordApp = Nothing
End Function

Function getWorkerDetails()
    With Session
        ECL01
        screen = .GetDisplayText(2, 12, 5)
        If screen = "ECL01" Then
            temp = .GetDisplayText(5, 20, 7)
            For tempx = 1 To Len(.GetDisplayText(5, 20, 7))
                If Mid(.GetDisplayText(5, 20, 7), tempx, 1) = "_" Then
                    count = count + 1
                    temp = "0" & temp
                End If
            Next
            temp = Replace(temp, "_", "")
            claim = .GetDisplayText(5, 28, 2) & .GetDisplayText(5, 17, 2) & temp 'index 8
            InjuryText = Trim(.GetDisplayText(13, 29, 51)) 'index 12
            InjuryDate = .GetDisplayText(11, 17, 10) 'index 11
            
            If Len(claim) = 11 Then
                .TransmitTerminalKey rcIBMHomeKey
                .TransmitTerminalKey rcIBMBacktabKey
                .TransmitANSI "ECM02"
                .TransmitTerminalKey rcIBMPf11Key
                .TransmitTerminalKey rcIBMEnterKey
                .WaitForEvent rcKbdEnabled, "30", "0", 1, 1
                .WaitForEvent rcEnterPos, "30", "0", 5, 22
                
                'residental address
                ClaimantAdd17 = .GetDisplayText(13, 22, 30)
                claimantAdd1 = Trim(StrConv(ClaimantAdd17, vbProperCase))
                ClaimantAdd27 = .GetDisplayText(14, 22, 30)
                claimantAdd2 = Trim(StrConv(ClaimantAdd27, vbProperCase))
                ClaimantAdd37 = .GetDisplayText(15, 22, 30)
                claimantAdd3 = Trim(StrConv(ClaimantAdd37, vbProperCase))
                    
                If claimantAdd3 = "" Then
                    address1 = StrConv(claimantAdd1, vbProperCase) 'index 3
                    address2 = splitAddress(Split(claimantAdd2, " "))
                Else
                    address1 = StrConv(claimantAdd1 & " " & claimantAdd2, vbProperCase) 'index 3
                    address2 = splitAddress(Split(claimantAdd3, " "))
                End If
                
                Suburb = StrConv(address2(0), vbProperCase) 'index 4
                State = address2(1) 'index 5
                Postcode = address2(2) 'index 6
                FirstName = StrConv(Trim(.GetDisplayText(7, 22, 16)), vbProperCase) 'index 0
                SurName = StrConv(Split(Trim(.GetDisplayText(6, 22, 26)), " ")(0), vbProperCase) 'index 1
                Mobile = Trim(.GetDisplayText(22, 54, 12)) 'index 7
                DateOfBirth = .GetDisplayText(5, 45, 10) 'index 2
                digitalConsent = .GetDisplayText(20, 80, 1) 'index 10
                If digitalConsent = "Y" Then
                    email = StrConv(.GetDisplayText(21, 22, 30), vbLowerCase)
                Else
                    email = "N/A" 'index 9
                End If
                
                ECL01
                getWorkerDetails = Array(FirstName, SurName, DateOfBirth, address1, Suburb, State, Postcode, Mobile, claim, email, digitalConsent, InjuryDate, InjuryText)
                Exit Function
            Else
                MsgBox "Couldn't find a valid claim number."
            End If
        End If
    End With
End Function

Function getCMDetails()
    With Session
        ECL01
        
        On Error Resume Next
        caseManagerFirstName = Trim(Split(.GetDisplayText(23, 52, 29), ",")(1))
        caseManagerSurname = Trim(Split(.GetDisplayText(23, 52, 29), ",")(0))
        If Err.Number = 0 Then
            .TransmitTerminalKey rcIBMHomeKey
            .TransmitTerminalKey rcIBMBacktabKey
            .TransmitANSI "ECX07"
            .TransmitTerminalKey rcIBMPf11Key
            .WaitForEvent rcKbdEnabled, "2", "0", 1, 1
            .WaitForEvent rcEnterPos, "2", "0", 7, 18
            .TransmitTerminalKey rcIBMNewLineKey
            .TransmitANSI caseManagerSurname
            .TransmitTerminalKey rcIBMEnterKey
            .WaitForEvent rcKbdEnabled, "2", "0", 1, 1
            .WaitForEvent rcEnterPos, "2", "0", 7, 18
            
            For cmRowNum = 14 To 23
                screenCMSurname = Trim(.GetDisplayText(cmRowNum, 21, 26))
                screenCMFirstName = Trim(.GetDisplayText(cmRowNum, 48, 15))
                If screenCMSurname = UCase(caseManagerSurname) And screenCMFirstName = UCase(caseManagerFirstName) Then
                    .SetMousePos cmRowNum, 4
                    .TerminalMouse rcLeftClick, rcMouseRow, rcMouseCol
                    .GraphicsMouse rcLeftClick, rcCurrentGraphicsCursorX, rcCurrentGraphicsCursorY
                    .TransmitANSI "S"
                    .TransmitTerminalKey rcIBMEnterKey
                    .WaitForEvent rcKbdEnabled, "2", "0", 1, 1
                    .WaitForEvent rcEnterPos, "2", "0", 24, 72
                    cmUserID = Trim(.GetDisplayText(9, 26, 7))
                    cmPositionID = Trim(.GetDisplayText(20, 26, 5))
                    foundCM = True
                    Exit For
                End If
            Next cmRowNum
            .TransmitTerminalKey rcIBMPf5Key
        Else
            Err.Clear
            foundCM = True
            cmUserID = .GetDisplayText(3, 12, 7)
        End If
        On Error GoTo 0
            
        If foundCM = True Then
            myFile = "G:\Shared\macros\ACCtionMacroCSVfiles\DHRX0001A - Team Structure.txt"
            Open myFile For Input As #1
            Do Until EOF(1)
                Line Input #1, currentLine
                currentLine = Replace(currentLine, Chr(34), "")
                userArray = Split(currentLine, ",")
                userID = userArray(4) 'index 0
                CMteamManager = userArray(2)
                CMteam = userArray(3) 'index 4
                CMFullName = userArray(5) 'index 1
                CMphone = "03 " & userArray(6) 'index 2
                CMemail = userArray(11) 'index 3
                
                If userID = cmUserID Then
                    ECL01
                    getCMDetails = Array(userID, CMFullName, CMphone, CMemail, CMteam)
                    Close #1
                    Exit Function
                End If
            Loop
            Close #1
        End If
    End With
End Function

Function containsNumbers(str)
    Dim regEx, match
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = "[0-9]"
    regEx.IgnoreCase = True
    regEx.Global = True
    Set match = regEx.Execute(str)
    containsNumbers = (match.count > 0)
End Function

Function containsLetters(str)
    Dim regEx, match
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = "[a-zA-Z]"
    regEx.IgnoreCase = True
    regEx.Global = True
    Set match = regEx.Execute(str)
    containsLetters = (match.count > 0)
End Function

Function IsStringInArray(str, arr)
    Dim i
    IsStringInArray = False
    For i = 0 To UBound(arr)
        If arr(i) = str Then
            IsStringInArray = True
            Exit For
        End If
    Next
End Function

Function splitAddress(arr)
    States = Array("VIC", "NSW", "QLD", "SA", "WA", "ACT", "TAS", "NT")
    State = ""
    For i = 0 To UBound(arr)
        If IsStringInArray(arr(i), States) Then
            State = arr(i)
        ElseIf containsLetters(arr(i)) Then
            If Not Suburb = "" Then
                Suburb = Suburb & " " & arr(i)
            Else
                Suburb = arr(i)
            End If
        Else
            If Not arr(i) = "" Then
                Postcode = arr(i)
            End If
        End If
    Next
    If State = "" Then
        State = getStateFromPostcode(Postcode)
    End If
    splitAddress = Array(Suburb, State, Postcode)
End Function

Function getStateFromPostcode(Postcode)
    If Postcode >= 3000 And Postcode <= 3999 Then
        State = "VIC"
    ElseIf Postcode >= 1000 And Postcode <= 2599 Then
        State = "NSW"
    ElseIf Postcode >= 2619 And Postcode <= 2899 Then
        State = "NSW"
    ElseIf Postcode >= 2921 And Postcode <= 2999 Then
        State = "NSW"
    ElseIf Postcode >= 4000 And Postcode <= 4999 Then
        State = "QLD"
    ElseIf Postcode >= 5000 And Postcode <= 5999 Then
        State = "SA"
    ElseIf Postcode >= 6000 And Postcode <= 6999 Then
        State = "WA"
    ElseIf Postcode >= 7000 And Postcode <= 7999 Then
        State = "TAS"
    ElseIf Postcode >= 800 And Postcode <= 999 Then
        State = "NT"
    ElseIf Postcode >= 200 And Postcode <= 299 Then
        State = "ACT"
    ElseIf Postcode >= 2600 And Postcode <= 2618 Then
        State = "ACT"
    ElseIf Postcode >= 2900 And Postcode <= 2920 Then
        State = "ACT"
    End If
    getStateFromPostcode = State
End Function

Function formatDate(xdate)
    dayPart = Day(xdate)
    monthPart = Month(xdate)
    yearPart = Year(xdate)
    
    If dayPart < 10 Then
        dayPart = "0" & dayPart
    End If
    If monthPart < 10 Then
        monthPart = "0" & monthPart
    End If
    
    formatDate = dayPart & "/" & monthPart & "/" & yearPart
End Function

Sub ECL01()
    With Session
        screen = .GetDisplayText(2, 12, 5)
        If screen <> "ECL01" Then
            .TransmitTerminalKey rcIBMHomeKey
            .TransmitTerminalKey rcIBMBacktabKey
            .TransmitANSI "ECL01"
            .TransmitTerminalKey rcIBMPf11Key
            .WaitForEvent rcKbdEnabled, "30", "0", 1, 1
            If .GetDisplayText(24, 55, 1) = "Y" Then
                .TransmitTerminalKey rcIBMHomeKey
                .TransmitTerminalKey rcIBMBacktabKey
                .TransmitTerminalKey rcIBMBacktabKey
                .TransmitANSI "N"
            End If
            .TransmitTerminalKey rcIBMEnterKey
            .WaitForEvent rcKbdEnabled, "30", "0", 1, 1
        End If
    End With
End Sub
