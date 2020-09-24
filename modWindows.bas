Attribute VB_Name = "modWindows"
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Type POINTAPI
        x As Long
        y As Long
End Type

Function GetWindowIndex(strCaption As String)
    Dim i As Integer
    For i = 1 To WindowCount
        If GetWindowTitle(i) = strCaption Then
            GetWindowIndex = i
            Exit Function
        End If
    Next i
    GetWindowIndex = -1
End Function

Function GetWindowTitle(intWhich As Integer) As String
    Dim i As Integer, cnt As Integer
    If intWhich = 1 Then GetWindowTitle = "Status": Exit Function
    If intWhich = 2 Then GetWindowTitle = "Friend Tracker": Exit Function
    
    cnt = 3
    For i = 1 To intChannels
        If cnt = intWhich Then
            GetWindowTitle = Channels(i).strName
            Exit Function
        End If
        cnt = cnt + 1
    Next i
    
    For i = 1 To intQueries
        If cnt = intWhich Then
            GetWindowTitle = Queries(i).strNick
            Exit Function
        End If
        cnt = cnt + 1
    Next i
    
    GoTo final
    
    For i = 1 To intDCCChats
        If cnt = intWhich Then
            'GetWindowTitle = DCCChats(i).Caption
            Exit Function
        End If
        cnt = cnt + 1
    Next i
    
    For i = 1 To intDCCSends
        If cnt = intWhich Then
            'GetWindowTitle = DCCSends(i).Caption
            Exit Function
        End If
        cnt = cnt + 1
    Next i
final:
    GetWindowTitle = ""
End Function
Sub SetWinFocus(intWhich As Integer)
    Dim i As Integer, cnt As Integer

    If intWhich = 1 Then
        If Status.Visible = False Then Status.Visible = True
        Status.SetFocus
        Exit Sub
    End If
    If intWhich = 2 Then
        If BuddyList.Visible = fase Then BuddyList.Visible = True
        BuddyList.SetFocus
        Exit Sub
    End If
    
    cnt = 3
    For i = 1 To intChannels
        If Channels(i).strName <> "" Then
            If cnt = intWhich Then
                Channels(i).SetFocus
                Exit Sub
            End If
            cnt = cnt + 1
        End If
    Next i
    
    For i = 1 To intQueries
        If Queries(i).strNick <> "" Then
            If cnt = intWhich Then
                Queries(i).SetFocus
                Exit Sub
            End If
            cnt = cnt + 1
        End If
    Next i
    
    
    'WindowCount = cnt + 1 'add 1 for status window

End Sub

Function TaskCenter(intActual As Integer, strText As String) As Integer
    TaskCenter = (intActual - Client.picTask.TextWidth(strText)) / 2
End Function


Function TaskText(intWidth As Integer, strText As String) As String
    'MsgBox intWidth & ".." & Client.picTask.TextWidth(strText) & ".."
    Dim lastWidth As Integer, i As Integer, strBuf As String
    Dim intTemp As Integer
    
    For i = 1 To Len(strText)
        strBuf = Left(strText, i) & "..."
        intTemp = Client.picTask.TextWidth(strBuf) + 8
        
        If intTemp >= intWidth Then
            TaskText = Left(strText, i - 1) & "..."
            Exit Function
        End If
    Next i
    TaskText = strText
        
End Function

Function WindowCount() As Integer
    Dim cnt As Integer, i As Integer
    
    For i = 1 To intChannels
        If Channels(i).strName <> "" Then cnt = cnt + 1
    Next i
    
    For i = 1 To intQueries
        If Queries(i).strNick <> "" Then cnt = cnt + 1
    Next i
    
    
    WindowCount = cnt + 2 'add 2 for status window and buddy list
    
    '* Add DCC and stuff here

End Function


Function WindowNewBuffer(intWhich As Integer) As String
    Dim i As Integer, cnt As Integer
    If intWhich = 1 Then
        If Status.newBuffer = True Then
            WindowNewBuffer = True
        Else
            WindowNewBuffer = False
        End If
        Exit Function
    End If
    If intWhich = 1 Then
        If BuddyList.newBuffer = True Then
            WindowNewBuffer = True
        Else
            WindowNewBuffer = False
        End If
        Exit Function
    End If
    
    cnt = 3
    For i = 1 To intChannels
        If cnt = intWhich Then
            If Channels(i).newBuffer = True Then
                WindowNewBuffer = True
                Exit Function
            Else
                WindowNewBuffer = False
                Exit Function
            End If
        End If
        cnt = cnt + 1
    Next i
    
    
    For i = 1 To intQueries
        If cnt = intWhich Then
            If Queries(i).newBuffer = True Then
                WindowNewBuffer = True
                Exit Function
            Else
                WindowNewBuffer = False
                Exit Function
            End If
        End If
        cnt = cnt + 1
    Next i
    
    GoTo final
    
    For i = 1 To intDCCChats
        If cnt = intWhich Then
'            If dccchats(i).newBuffer = True Then
                WindowNewBuffer = True
                Exit Function
'            Else
                WindowNewBuffer = False
                Exit Function
'            End If
        End If
        cnt = cnt + 1
    Next i
    
    For i = 1 To intDCCSends
        If cnt = intWhich Then
'            If dccsends(i).newBuffer = True Then
                WindowNewBuffer = True
                Exit Function
'            Else
                WindowNewBuffer = False
                Exit Function
'            End If
        End If
        cnt = cnt + 1
    Next i
final:
    WindowNewBuffer = False ' ""

End Function


