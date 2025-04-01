On Error Resume Next
Dim objFSO, objShell, strUSBDrive, objDummyFolder, objFolder, item, strMessage
Dim recoveredFiles, recoveredFolders, recoveredFilesList, recoveryLogFile
Dim currentProgress, totalItems, progressBar, logFilePath, userChoice

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

recoveredFiles = 0
recoveredFolders = 0
recoveredFilesList = ""
currentProgress = 0
totalItems = 0
progressBar = ""

MsgBox "Flash Drive Recovery Tool" & vbCrLf & vbCrLf & _
       "This tool will help you recover files hidden by malware on your flash drive." & vbCrLf & _
       "1. Your original files will be restored to a folder named 'Recovered'" & vbCrLf & _
       "2. A detailed report will be created with the list of all recovered items" & vbCrLf & _
       "3. Malicious files can be automatically removed" & vbCrLf & vbCrLf & _
       "Click OK to continue.", vbInformation, "Flash Drive Recovery"

strUSBDrive = InputBox("Enter flash drive letter (e.g., F:)" & vbCrLf & _
                       "Available drives: " & GetAvailableDrives(), "Flash Drive Recovery")

If strUSBDrive = "" Then
    MsgBox "Operation canceled by user.", vbExclamation, "Canceled"
    WScript.Quit
End If

If Right(strUSBDrive, 1) <> ":" Then
    strUSBDrive = strUSBDrive & ":"
End If

If Not objFSO.DriveExists(strUSBDrive) Then
    MsgBox "Drive " & strUSBDrive & " does not exist." & vbCrLf & _
           "Available drives: " & GetAvailableDrives(), vbExclamation, "Drive Not Found"
    WScript.Quit
End If

logFilePath = strUSBDrive & "\RecoveryReport.txt"

If objFSO.FolderExists(strUSBDrive & "\_") Then
    Set objDummyFolder = objFSO.GetFolder(strUSBDrive & "\_")

    CountItems objDummyFolder
    
    objDummyFolder.Attributes = objDummyFolder.Attributes And Not 2
    
    If Not objFSO.FolderExists(strUSBDrive & "\Recovered") Then
        objFSO.CreateFolder(strUSBDrive & "\Recovered")
    End If

    MsgBox "Found " & totalItems & " items to recover." & vbCrLf & _
           "Recovery process will now begin.", vbInformation, "Starting Recovery"
    
    RecoverItems objDummyFolder, strUSBDrive & "\Recovered"
    
    CreateRecoveryLog
    
    If MsgBox("Do you want to remove malicious files from the flash drive?" & vbCrLf & _
              "This will delete the virus files and shortcuts.", _
              vbYesNo + vbQuestion, "Remove Malicious Files") = vbYes Then
        RemoveMaliciousFiles strUSBDrive
    End If
    
    If MsgBox("Recovery complete!" & vbCrLf & _
              "Files recovered: " & recoveredFiles & vbCrLf & _
              "Folders recovered: " & recoveredFolders & vbCrLf & vbCrLf & _
              "A detailed report has been saved to: " & logFilePath & vbCrLf & vbCrLf & _
              "Do you want to open the recovered files folder?", _
              vbYesNo + vbQuestion, "Recovery Complete") = vbYes Then
        objShell.Run "explorer.exe """ & strUSBDrive & "\Recovered""", 1, False
    End If
    
    If MsgBox("Do you want to view the recovery log?", _
              vbYesNo + vbQuestion, "View Log") = vbYes Then
        objShell.Run "notepad.exe """ & logFilePath & """", 1, False
    End If
Else
    userChoice = MsgBox("Hidden folder with data not found on drive " & strUSBDrive & "." & vbCrLf & _
                       "Would you like to scan the entire drive for hidden files?", _
                       vbYesNo + vbQuestion, "Hidden Folder Not Found")
    
    If userChoice = vbYes Then
        MsgBox "Scanning entire drive. This may take some time...", vbInformation, "Scanning"
        ScanDriveForHiddenItems strUSBDrive
    Else
        MsgBox "Files may have been moved elsewhere or deleted." & vbCrLf & _
               "You may need to use specialized data recovery software.", _
               vbExclamation, "Recovery Failed"
    End If
End If

Sub CountItems(folder)
    Dim subFolder, file
    
    For Each file In folder.Files
        totalItems = totalItems + 1
    Next
    
    For Each subFolder In folder.SubFolders
        totalItems = totalItems + 1
        CountItems subFolder
    Next
End Sub

Sub RecoverItems(sourceFolder, destPath)
    Dim subFolder, file, newFolder, fileSize
    
    For Each file In sourceFolder.Files
        On Error Resume Next
        objFSO.CopyFile file.Path, destPath & "\" & file.Name, True
        
        If Err.Number = 0 Then
            recoveredFiles = recoveredFiles + 1
            
            fileSize = FormatFileSize(file.Size)
            
            recoveredFilesList = recoveredFilesList & file.Name & " (" & fileSize & ")" & vbCrLf
            
            currentProgress = currentProgress + 1
            UpdateProgress
        Else
            recoveredFilesList = recoveredFilesList & "ERROR: " & file.Name & " - " & Err.Description & vbCrLf
        End If
        On Error GoTo 0
    Next
    
    For Each subFolder In sourceFolder.SubFolders
        newFolder = destPath & "\" & subFolder.Name
        If Not objFSO.FolderExists(newFolder) Then
            objFSO.CreateFolder newFolder
            recoveredFolders = recoveredFolders + 1
            
            recoveredFilesList = recoveredFilesList & "[FOLDER] " & subFolder.Name & vbCrLf
            
            currentProgress = currentProgress + 1
            UpdateProgress
        End If
        
        RecoverItems subFolder, newFolder
    Next
End Sub

Sub CreateRecoveryLog
    Dim logFile
    
    Set logFile = objFSO.CreateTextFile(logFilePath, True)
    
    logFile.WriteLine "======== FLASH DRIVE RECOVERY REPORT ========" 
    logFile.WriteLine "Date and Time: " & Now
    logFile.WriteLine "Drive: " & strUSBDrive
    logFile.WriteLine "Total files recovered: " & recoveredFiles
    logFile.WriteLine "Total folders recovered: " & recoveredFolders
    logFile.WriteLine "=============================================="
    logFile.WriteLine ""
    logFile.WriteLine "RECOVERED ITEMS:"
    logFile.WriteLine "----------------"
    logFile.WriteLine recoveredFilesList
    logFile.WriteLine ""
    logFile.WriteLine "All recovered items are located in: " & strUSBDrive & "\Recovered"
    logFile.WriteLine "=============================================="
    
    logFile.Close
End Sub

Sub RemoveMaliciousFiles(drivePath)
    On Error Resume Next
    Dim filesRemoved, foldersRemoved
    
    filesRemoved = 0
    foldersRemoved = 0
    
    If objFSO.FileExists(drivePath & "\*.lnk") Then
        Set objFolder = objFSO.GetFolder(drivePath)
        For Each item In objFolder.Files
            If LCase(Right(item.Name, 4)) = ".lnk" Then
                objFSO.DeleteFile item.Path, True
                filesRemoved = filesRemoved + 1
            End If
        Next
    End If
    
    If objFSO.FolderExists(drivePath & "\WindowsServices") Then
        objFSO.DeleteFolder drivePath & "\WindowsServices", True
        foldersRemoved = foldersRemoved + 1
    End If
    
    If objFSO.FolderExists(drivePath & "\_") Then
        If MsgBox("All files have been recovered to the 'Recovered' folder." & vbCrLf & _
                  "Do you want to remove the original hidden folder (_)?" & vbCrLf & _
                  "(This is safe if all files were successfully recovered)", _
                  vbYesNo + vbQuestion, "Remove Original Files") = vbYes Then
            objFSO.DeleteFolder drivePath & "\_", True
            foldersRemoved = foldersRemoved + 1
        End If
    End If
    
    MsgBox "Cleanup complete!" & vbCrLf & _
           "Removed " & filesRemoved & " malicious files" & vbCrLf & _
           "Removed " & foldersRemoved & " malicious folders", _
           vbInformation, "Cleanup Complete"
    
    On Error GoTo 0
End Sub

Sub ScanDriveForHiddenItems(drivePath)
    Dim rootFolder, hiddenItems, foundItems
    foundItems = 0
    
    Set rootFolder = objFSO.GetFolder(drivePath & "\")
    
    If Not objFSO.FolderExists(drivePath & "\Recovered") Then
        objFSO.CreateFolder drivePath & "\Recovered"
    End If
    
    MsgBox "Scanning for hidden items. This may take a while...", vbInformation, "Scanning"
    
    SearchHiddenItems rootFolder, drivePath & "\Recovered", foundItems
    
    If foundItems > 0 Then
        MsgBox "Found and recovered " & foundItems & " hidden items!", vbInformation, "Scan Complete"
    Else
        MsgBox "No hidden items found on this drive." & vbCrLf & _
               "You may need specialized data recovery software.", vbExclamation, "Scan Complete"
    End If
End Sub

Sub SearchHiddenItems(folder, destPath, foundItems)
    Dim subFolder, file, attribs
    
    On Error Resume Next
    
    For Each file In folder.Files
        attribs = file.Attributes
        If (attribs And 2) = 2 Then
            objFSO.CopyFile file.Path, destPath & "\" & file.Name, True
            foundItems = foundItems + 1
        End If
    Next

    For Each subFolder In folder.SubFolders
        attribs = subFolder.Attributes
        
        If subFolder.Name <> "System Volume Information" And subFolder.Name <> "$RECYCLE.BIN" Then
            If (attribs And 2) = 2 Then
                If Not objFSO.FolderExists(destPath & "\" & subFolder.Name) Then
                    objFSO.CreateFolder destPath & "\" & subFolder.Name
                End If
                
                CopyFolderContents subFolder, destPath & "\" & subFolder.Name, foundItems
                foundItems = foundItems + 1
            End If
            
            SearchHiddenItems subFolder, destPath, foundItems
        End If
    Next
    
    On Error GoTo 0
End Sub

Sub CopyFolderContents(sourceFolder, destFolder, foundItems)
    Dim subFolder, file
    
    On Error Resume Next
    
    For Each file In sourceFolder.Files
        objFSO.CopyFile file.Path, destFolder & "\" & file.Name, True
        foundItems = foundItems + 1
    Next
    
    For Each subFolder In sourceFolder.SubFolders
        If Not objFSO.FolderExists(destFolder & "\" & subFolder.Name) Then
            objFSO.CreateFolder destFolder & "\" & subFolder.Name
        End If
        CopyFolderContents subFolder, destFolder & "\" & subFolder.Name, foundItems
    Next
    
    On Error GoTo 0
End Sub

Function FormatFileSize(byteSize)
    If byteSize < 1024 Then
        FormatFileSize = byteSize & " B"
    ElseIf byteSize < 1048576 Then
        FormatFileSize = FormatNumber(byteSize / 1024, 2) & " KB"
    ElseIf byteSize < 1073741824 Then
        FormatFileSize = FormatNumber(byteSize / 1048576, 2) & " MB"
    Else
        FormatFileSize = FormatNumber(byteSize / 1073741824, 2) & " GB"
    End If
End Function

Function GetAvailableDrives()
    Dim drives, drive, driveList
    driveList = ""
    
    Set drives = objFSO.Drives
    
    For Each drive in drives
        If drive.DriveType = 1 Then
            driveList = driveList & drive.DriveLetter & ": (" & drive.VolumeName & "), "
        End If
    Next
    
    If driveList = "" Then
        driveList = "No removable drives found"
    Else
        driveList = Left(driveList, Len(driveList) - 2)
    End If
    
    GetAvailableDrives = driveList
End Function

Sub UpdateProgress()
    Dim percentage, progressMsg
    
    If totalItems > 0 Then
        percentage = Int((currentProgress / totalItems) * 100)
       
        progressBar = String(percentage \ 5, "|")
        progressMsg = "Recovering files: " & percentage & "% complete" & vbCrLf & _
                      "[" & progressBar & String(20 - (percentage \ 5), " ") & "]" & vbCrLf & _
                      "Files: " & recoveredFiles & " | Folders: " & recoveredFolders & vbCrLf & _
                      "Currently processing: " & currentProgress & " of " & totalItems & " items"
                      
        WScript.Echo progressMsg
    End If
End Sub
