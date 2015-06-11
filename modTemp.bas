Attribute VB_Name = "modTemp"


Sub BatchTest()

    Dim i As Integer
    Dim rslt As Double
    

    
    
    'For i = 1 To 100
     Do
        i = i + 10
        If i Mod 100 = 0 Then
            DoCmd.Echo True, i
            Application.Echo EchoOn:=False, bstrStatusBarText:="Your Message Here"
            
            Dim varReturn As Variant
            varReturn = SysCmd(acSysCmdSetStatus, "Text to write on the Status Bar!")
 

        End If
    
        'If (i Mod 3) = 0 Then
        '    If (i Mod 5) = 0 Then
        '        MsgBox ("FizzBuzz")
        '    Else
        '        MsgBox ("Buzz")
        '    End If
        'Else
        '    MsgBox (i)
        'End If
    Loop
    'Next i
    
   
End Sub

