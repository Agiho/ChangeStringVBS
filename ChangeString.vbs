Dim Path, oldString, newString, objFSO, objFolder, objFile, strText, fileExtension

Path = WScript.Arguments(0)
oldString = WScript.Arguments(1)
newString = WScript.Arguments(2)

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(Path)


Sub searchFolders(objFolder)
    For Each objFile In objFolder.Files
    On Error Resume Next
        fileExtension = objFSO.GetExtensionName(objFile)
        If (fileExtension <> "MSP") And (fileExtension <> "MST") Then
            strText = objFSO.OpenTextFile(objFile).ReadAll
            WScript.Echo objFile
            If Err.Number = 0 Then
                strText = Replace(strText, oldString, newString)
                objFSO.OpenTextFile(objFile, 2).Write strText
            End If
            On Error GoTo 0
        End If
    Next

    For Each objSubFolder In objFolder.SubFolders
        searchFolders objSubFolder
    Next

End Sub

searchFolders objFolder

WScript.Echo "Replace operation completed."