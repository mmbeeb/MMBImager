Attribute VB_Name = "CommonFile"
Option Explicit

' Note: This code is from the Access 95 Developer's Handbook,
' by Paul Litwin, Ken Getz, Mike Gilbert, and Greg Reddick.
' (c) 1995 by Sybex.
' Used with permission


Type tagOPENFILENAME
     lStructSize As Long
     hwndOwner As Long
     hInstance As Long
     strFilter As String
     strCustomFilter As String
     nMaxCustFilter As Long
     nFilterIndex As Long
     strFile As String
     nMaxFile As Long
     strFileTitle As String
     nMaxFileTitle As Long
     strInitialDir As String
     strTitle As String
     Flags As Long
     nFileOffset As Integer
     nFileExtension As Integer
     strDefExt As String
     lCustData As Long
     lpfnHook As Long
     lpTemplateName As String
End Type

Declare Function glr_apiGetOpenFileName Lib "comdlg32.dll" _
 Alias "GetOpenFileNameA" (ofn As tagOPENFILENAME) As Boolean
Declare Function glr_apiGetSaveFileName Lib "comdlg32.dll" _
 Alias "GetSaveFileNameA" (ofn As tagOPENFILENAME) As Boolean
Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long

Public Const glrOFN_READONLY = &H1
Public Const glrOFN_OVERWRITEPROMPT = &H2
Public Const glrOFN_HIDEREADONLY = &H4
Public Const glrOFN_NOCHANGEDIR = &H8
Public Const glrOFN_SHOWHELP = &H10
Public Const glrOFN_NOVALIDATE = &H100
Public Const glrOFN_ALLOWMULTISELECT = &H200
Public Const glrOFN_EXTENSIONDIFFERENT = &H400
Public Const glrOFN_PATHMUSTEXIST = &H800
Public Const glrOFN_FILEMUSTEXIST = &H1000
Public Const glrOFN_CREATEPROMPT = &H2000
Public Const glrOFN_SHAREAWARE = &H4000
Public Const glrOFN_NOREADONLYRETURN = &H8000
Public Const glrOFN_NOTESTFILECREATE = &H10000
Public Const glrOFN_NONETWORKBUTTON = &H20000
Public Const glrOFN_NOLONGNAMES = &H40000
Public Const glrOFN_EXPLORER = &H80000
Public Const glrOFN_NODEREFERENCELINKS = &H100000
Public Const glrOFN_LONGNAMES = &H200000

Function glrCommonFileOpenSave( _
 Optional ByRef Flags As Variant, _
 Optional ByVal InitialDir As Variant, _
 Optional ByVal Filter As Variant, _
 Optional ByVal FilterIndex As Variant, _
 Optional ByVal DefaultExt As Variant, _
 Optional ByVal FileName As Variant, _
 Optional ByVal DialogTitle As Variant, _
 Optional ByVal OpenFile As Variant, _
 Optional ByVal Hwnd As Variant) As Variant
 
    ' This is the entry point you'll use to call the common
    ' file open/save dialog. The parameters are listed
    ' below, and all are optional.
    '
    ' In:
    '    Flags: one or more of the glrOFN_* constants, OR'd together.
    '    InitialDir: the directory in which to first look
    '    Filter: a set of file filters, set up by calling
    '            AddFilterItem.  See examples.
    '    FilterIndex: 1-based integer indicating which filter
    '            set to use, by default (1 if unspecified)
    '    DefaultExt: Extension to use if the user doesn't enter one.
    '            Only useful on file saves.
    '    FileName: Default value for the file name text box.
    '    DialogTitle: Title for the dialog.
    '    OpenFile: Boolean(True=Open File/False=Save As)
    '    Handle of window to act as parent to the dialog.
    ' Out:
    '    Return Value: Either Null or the selected filename

    Dim ofn As tagOPENFILENAME
    Dim strFileName As String
    Dim strFileTitle As String
    Dim fResult As Boolean

    ' Give the dialog a caption title.
    If IsMissing(InitialDir) Then InitialDir = ""
    If IsMissing(Filter) Then Filter = ""
    If IsMissing(FilterIndex) Then FilterIndex = 1
    If IsMissing(Flags) Then Flags = 0&
    If IsMissing(DefaultExt) Then DefaultExt = ""
    If IsMissing(FileName) Then FileName = ""
    If IsMissing(DialogTitle) Then DialogTitle = ""
    If IsMissing(OpenFile) Then OpenFile = True
    'If IsMissing(Hwnd) Then Hwnd = Application.hWndAccessApp
    
    ' Allocate string space for the returned strings.
    strFileName = Left(FileName & String(256, 0), 256)
    strFileTitle = String(256, 0)

    ' Set up the data structure before you call the function
    With ofn
        .lStructSize = Len(ofn)
        .hwndOwner = Hwnd
        .strFilter = Filter
        .nFilterIndex = FilterIndex
        .strFile = strFileName
        .nMaxFile = Len(strFileName)
        .strFileTitle = strFileTitle
        .nMaxFileTitle = Len(strFileTitle)
        .strTitle = DialogTitle
        .Flags = Flags
        .strDefExt = DefaultExt
        .strInitialDir = CurDir

        ' Didn't think most people would want to deal with
        ' these options.
        .hInstance = 0
        .strCustomFilter = String(255, 0)
        .nMaxCustFilter = 255
        .lpfnHook = 0
    End With

    ' This will pass the desired data structure to the
    ' Windows API, which will in turn it uses to display
    ' the Open/Save As Dialog.

    If OpenFile Then
        fResult = glr_apiGetOpenFileName(ofn)
    Else
        fResult = glr_apiGetSaveFileName(ofn)
    End If

    ' The function call filled in the strFileTitle member
    ' of the structure. You'll have to write special code
    ' to retrieve that if you're interested.

    If fResult Then
        ' You might care to check the Flags member of the
        ' structure to get information about the chosen file.
        ' In this example, if you bothered to pass in a
        ' value for Flags, we'll fill it in with the outgoing
        ' Flags value.
        If Not IsMissing(Flags) Then Flags = ofn.Flags
        glrCommonFileOpenSave = glrTrimNull(ofn.strFile)
    Else
        glrCommonFileOpenSave = Null
    End If
End Function

Function glrAddFilterItem(strFilter As String, _
 strDescription As String, Optional varItem As Variant) As String

    ' Tack a new chunk onto the file filter.
    ' That is, take the old value, stick onto it the description,
    ' (like "Databases"), a null character, the skeleton
    ' (like "*.mdb;*.mda") and a final null character.

    If IsMissing(varItem) Then varItem = "*.*"
    glrAddFilterItem = strFilter & _
     strDescription & vbNullChar & _
     varItem & vbNullChar
End Function

Function glrTrimNull(ByVal strItem As String) As String
    Dim intPos As Integer
    intPos = InStr(strItem, vbNullChar)
    If intPos > 0 Then
        glrTrimNull = Left(strItem, intPos - 1)
    Else
        glrTrimNull = strItem
    End If
End Function


