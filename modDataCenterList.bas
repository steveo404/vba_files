Attribute VB_Name = "modDataCenterList"
Option Compare Database

Public Sub DataCenterList()

    Dim serverRoomList As String
    Dim entry As String
    Dim entryResult As String
    Dim triggerOut As Integer
    Dim userName As String
    Dim door As String
    Dim location As String
    Dim extension As String
    Dim Anything As String
    Dim outFile As String
    Dim trigger As String
      
    
    
    'serverRoomList = GetFileName()
    serverRoomList = "c:\server_room.txt"
    
    outFile = "C:\serverroomformat.txt"
    Open outFile For Append As #3
    Close #3
    Kill outFile
    

    Open serverRoomList For Input As #9
    
    Do Until EOF(9)
        Input #9, entry
        entryResult = (Trim(entry))
        
        If Len(entryResult) = 0 Then
            trigger = "off"
        Else
            trigger = "on"
        End If
        If trigger = "on" Then
            If triggerOut < 5 Then
                Select Case triggerOut
                    Case 0
                        userName = entryResult
                    Case 1
                        door = entryResult
                        door = Trim(door)
                    Case 2
                        location = entryResult
                    Case 3
                        extension = entryResult
                        triggerOut = -1
                        trigger = "off"
                        Open outFile For Append As #4
                        Print #4, userName, door, location, extension
                        Close #4
                End Select

                triggerOut = triggerOut + 1
            End If
        End If

    Loop
    
    
    
    
    MsgBox ("ServerRoomList Complete")

End Sub


