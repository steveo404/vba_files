Attribute VB_Name = "Module1"
Option Compare Database

Sub CopyFile()

    'Script copies the batch file and it's source file from the Audit drive to the users' local drive
    'Files are copied and placed on the local drive in the root c: directory

    Dim obj As Object
    Dim scriptFile As String
    Dim scriptSourceFile As String
    Dim endFolder As String
    Dim deskfile As String
    
    deskfile = "LogonIDs.FXD"
    
    Name "C:\Users\soneal\Desktop\LogonIDs.FXD" As "C:\Users\soneal\Desktop\LogonIDs.txt"
    
    scriptFile = "I:\Continuous Auditing\Manual Controls Testing\Terminated Employees\UserLookUp_count_1.1.bat"
    scriptSourceFile = "I:\Continuous Auditing\Manual Controls Testing\Terminated Employees\UsersForScript.txt"
    endFolder = "c:\"
    
    If Len(Dir(scriptFile)) <> 0 Then
        Set obj = CreateObject("Scripting.FileSystemObject")
        obj.CopyFile scriptFile, endFolder, True
    End If
    If Len(Dir(scriptSourceFile)) <> 0 Then
        Set obj = CreateObject("Scripting.FileSystemObject")
        obj.CopyFile scriptSourceFile, endFolder, True
    End If
End Sub
