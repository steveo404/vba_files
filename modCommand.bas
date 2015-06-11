Attribute VB_Name = "modCommand"
Option Compare Database

Sub Cmd()
    
    ' Start a copy of the Calculator.

    Shell "cmd.exe", 1
    
    
    Dim PauseTime, Start, Finish, TotalTime
    
    'If (MsgBox("Press Yes to pause for 5 seconds", 4)) = vbYes Then
        PauseTime = 5    ' Set duration.
        Start = Timer    ' Set start time.
        Do While Timer < Start + PauseTime
            DoEvents    ' Yield to other processes.
        Loop
        Finish = Timer    ' Set end time.
        TotalTime = Finish - Start    ' Calculate total time.
        'MsgBox "Paused for " & TotalTime & " seconds"
    'Else
    '    End
    'End If
 



    ' Tell the user that the Calculator is ready.

    'MsgBox "Command line is started!"

    ' Make sure the application is in the foreground.

    AppActivate "Administrator: C:\Windows\system32\cmd.exe"

    ' Send a message to the application.

    SendKeys "echo 'Hello World'"


End Sub
