On Error Resume Next
Set Session = GetObject(, "ReflectionIBM.Session")
Set Reflection = CreateObject("ReflectionIBM")
    
If Err.Number <> 0 Then
    Err.Clear
    CreateObject("WScript.Shell").Run("H:\Acction\acction.rsf")
    Set Session = GetObject(, "ReflectionIBM.Session")
    While Err.Number <> 0
        WScript.Sleep 100
        Err.Clear
        Set Session = GetObject(, "ReflectionIBM.Session")
    Wend
End If
On Error GoTo 0

Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.OpenTextFile(fso.GetAbsolutePathName(".") & "\scr\currentUpdate.csv", 1)
updateContent = file.ReadAll
file.Close
updateVersion = Split(Split(updateContent, vbCrLf)(0), ",")(1)
updatePath = "H:/Acction/update.csv"
upToDate = False
If fso.FileExists(updatePath) Then
    Set file = fso.OpenTextFile(updatePath, 1)
    csvContent = file.ReadAll
    file.Close
    updateData = Split(Split(csvContent, vbCrLf)(0), ",")  
    If updateData(1) = updateVersion Then
        upToDate = True
    End If
Else
    Set file = fso.CreateTextFile(updatePath, True)
    file.WriteLine("UpdateVersion," & updateVersion)
    file.WriteLine("DateUpdated," & Now)
    file.Close
End If

If not upToDate Then
    moduleExists = False
    Set project = Session.VBProject
    For Each importedModule In project.VBComponents
        If InStr(importedModule.Name, "customMacros") Then
            moduleExists = True
        End If
    Next
        
    If moduleExists = True Then 
        Session.VBProject.VBComponents.Remove Session.VBProject.VBComponents.Item("customMacros")
        Session.VBProject.VBComponents.Remove Session.VBProject.VBComponents.Item("NathansFormGenerator")
    End If
    project.VBComponents.Import(fso.GetAbsolutePathName(".") & "\scr\customMacros.bas")
    project.VBComponents.Import(fso.GetAbsolutePathName(".") & "\scr\NathansFormGenerator.frm")

    tempUpdatePath = "H:/Acction/tempupdate.csv" 
    Set file = fso.OpenTextFile(updatePath, 1) 
    
    On Error Resume Next
    fileText = file.ReadAll
    If Err.Number <> 0 Then
        lines = Array(("UpdateVersion," & updateVersion), ("DateUpdated," & Now))
        Err.Clear
    Else
        lines = Split(fileText, vbCrLf)
        lines(0) = "UpdateVersion," & updateVersion
        lines(1) = "DateUpdated," & Now
    End If
    file.Close
    On Error GoTo 0
    
    Set file = fso.CreateTextFile(tempUpdatePath, True)
    For i = 0 To UBound(lines)
        file.WriteLine lines(i)
    Next
    file.Close
    
    fso.DeleteFile (updatePath)
    fso.MoveFile tempUpdatePath, updatePath

    session.SetKeyMap 0, "KpMinus", "RunMacro ""customMacros.openFormGenerator"", """""
    session.SetKeyMap 4, "C", "RunMacro ""customMacros.copyClaimNumber"", """""
    session.SetKeyMap 4, "V", "RunMacro ""customMacros.pasteClaimNumber"", """""

    MsgBox("update complete!")
Else
    MsgBox "Already up to date."
End If