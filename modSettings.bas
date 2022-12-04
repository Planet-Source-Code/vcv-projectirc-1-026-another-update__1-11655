Attribute VB_Name = "modSettings"
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Global INI As String

'* Server settings
Global strServer    As String
Global strMyNick    As String
Global strOtherNick As String
Global strFullName  As String
Global strMyIdent   As String
Global lngPort      As Long
Global bConOnLoad   As Boolean
Global bReconnect   As Boolean
Global bInvisible   As Boolean
Global bRetry       As Boolean
Global intRetry     As Integer

'* Font...
Global strFont As String
Global strFontSize As Integer
Public Function TF(bVal As Boolean) As Integer
    If bVal Then
        TF = 1
    Else
        TF = 0
    End If
End Function

Function FileExists(FullFileName As String) As Boolean
    On Error GoTo MakeF
        'If file does Not exist, there will be an Error
        Open FullFileName For Input As #1
        Close #1
        'no error, file exists
        FileExists = True
    Exit Function
MakeF:
        'error, file does Not exist
        FileExists = False
    Exit Function
End Function
Function ReadINI(strSection As String, strSetting As String, strDefault As String)
    Dim lngReturn As Long, strReturn As String, lngSize As Long
    lngSize = 255
    strReturn = String(lngSize, 0)
    lngReturn = GetPrivateProfileString(strSection, strSetting, strDefault, strReturn, lngSize, path & "settings.ini")
    If strReturn = "" Then
        ReadINI = strDefault
        WriteINI strSection, strSetting, strDefault
    Else
        ReadINI = strReturn
    End If
End Function


Sub WriteINI(strSection As String, strLValue As String, strRValue As String)
    Dim lngReturn As Long
    lngReturn = WritePrivateProfileString(strSection, strLValue, strRValue, path & "settings.ini")
    'MsgBox lngReturn & "..ini"
End Sub


