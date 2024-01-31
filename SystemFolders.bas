Attribute VB_Name = "SystemFolders"
Option Explicit

'******************** Code Start **************************
' This code was originally written by Dev Ashish.
' It is not to be altered or distributed,
' except as part of an application.
' You are free to use it in any application,
' provided the copyright notice is left unchanged.
'
' Code Courtesy of
' Dev Ashish
'
Private Const MAX_PATH As Integer = 255
Private Declare Function apiGetSystemDirectory& Lib "kernel32" _
        Alias "GetSystemDirectoryA" _
        (ByVal lpBuffer As String, ByVal nSize As Long)

Private Declare Function apiGetWindowsDirectory& Lib "kernel32" _
        Alias "GetWindowsDirectoryA" _
        (ByVal lpBuffer As String, ByVal nSize As Long)

Private Declare Function apiGetTempDir Lib "kernel32" _
        Alias "GetTempPathA" (ByVal nBufferLength As Long, _
        ByVal lpBuffer As String) As Long
Function fReturnTempDir()
'Returns Temp Folder Name
Dim strTempDir As String
Dim lngx As Long
    strTempDir = String$(MAX_PATH, 0)
    lngx = apiGetTempDir(MAX_PATH, strTempDir)
    If lngx <> 0 Then
        fReturnTempDir = Left$(strTempDir, lngx)
    Else
        fReturnTempDir = ""
    End If
End Function
Function fReturnSysDir()
'Returns System Folder Name (C:\WinNT\System32)
Dim strSysDirName As String
Dim lngx As Long
    strSysDirName = String$(MAX_PATH, 0)
    lngx = apiGetSystemDirectory(strSysDirName, MAX_PATH)
    If lngx <> 0 Then
        fReturnSysDir = Left$(strSysDirName, lngx)
    Else
        fReturnSysDir = ""
    End If
End Function
Function fReturnWinDir()
'Returns OS Folder (C:\Win95)
Dim strWinDirName As String
Dim lngx As Long
    strWinDirName = String$(MAX_PATH, 0)
    lngx = apiGetWindowsDirectory(strWinDirName, MAX_PATH)
    If lngx <> 0 Then
        fReturnWinDir = Left$(strWinDirName, lngx)
    Else
        fReturnWinDir = ""
    End If
End Function
'******************** Code End**************************

