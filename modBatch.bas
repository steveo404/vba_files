Attribute VB_Name = "modBatch"
Option Compare Database

Sub CmdTest()
    Dim strBatch As String
    strBatch = "c:\Scripts\testUserLookUp.bat"

    Shell "cmd /k """ & strBatch & """,vbNormalFocus"
    
        
    Dim i As Integer
      
    i = 0
    
    Do
        If Len(Dir("c:\Trigger\helloworld.txt")) <> 0 Then
            MsgBox ("Script Complete")
            Exit Do
        End If
        'i = i + 1
    Loop
End Sub
