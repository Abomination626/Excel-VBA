Attribute VB_Name = "Create_Batchfile_VBA"
Sub Main()
'******************************************
'THIS CODE CAN HELP EXECUTE UTIL SCRIPTS
'USING CMD OR ANY TERMINAL APPLICATION
'MODIFY AS NEEDED
'
'******************************************

    '  Modify this to desired name
    BatchFileName = "AnyName.bat"
    
    '  CREATE BATCH FILE
    Call CreateBatchFile(BatchFileName)
     
    '  COMMENCE EXECUTION
    ChDir ThisWorkbook.Path
    Shell vbShellExec & BatchFileName
    
    ' ADD WAITING TIME FOR BATCH FILE CREATION 3SECONDS TO PREVENT ERROR
    Application.Wait Now + TimeValue("0:00:03")
    
    ' Optional: Display confirmation message
    MsgBox "File " & BatchFileName & " executed."
    
    'DELETE BATCH FILE - CLEANUP
    Kill ThisWorkbook.Path & "\" & BatchFileName
    
    ' Optional: Display confirmation message
    MsgBox "Done!", vbOKOnly
    
End Sub

Sub CreateBatchFile(BatchFileName)
    ' VARIABLE DECLARATION
    Dim filePath As String
    Dim batchContent As String
    Dim fso As Object
    Dim file As Object
  
    filePath = ThisWorkbook.Path & "\" & BatchFileName
    
    ' Write the contents of batch file line-by-line
    batchContent = batchContent & "@echo off" & vbCrLf      'suppress cmd line display
    batchContent = batchContent & "setlocal" & vbCrLf       'create local env for variable
    batchContent = batchContent & "ipconfig/all" & vbCrLf   'any desired commands
    batchContent = batchContent & "pause" & vbCrLf
    batchContent = batchContent & "endlocal" & vbCrLf       'end local env
    
    'batchContent = batchContent & "cmd.exe /c other_batch_file.bat" & vbCrLf       'execute other .bat
    'batchContent = batchContent & "cmd.exe /k my_app.exe" & vbCrLf                 'execute other .exe
    'batchContent = batchContent & "cmd.exe /k /D "D:/Folder" my_app.exe" & vbCrLf  'execute from other folder
    
    ' Create FileSystemObject and open file
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.CreateTextFile(filePath, True)
    
    ' Write content and close file
    file.WriteLine batchContent
    file.Close
    
    ' Optional: Display confirmation message
    MsgBox "Batch file created successfully!", vbInformation
    
    ' CLEAR VARIABLES
    Set fso = Nothing
    Set file = Nothing
    
End Sub

