VERSION 5.00
Begin VB.Form NewImageForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create New Image"
   ClientHeight    =   3510
   ClientLeft      =   4935
   ClientTop       =   4290
   ClientWidth     =   5385
   Icon            =   "NewImageForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   465
      Left            =   2880
      TabIndex        =   7
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   315
      Left            =   1320
      TabIndex        =   5
      Text            =   "511"
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   465
      Left            =   1320
      TabIndex        =   4
      Top             =   2880
      Width           =   1245
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   1320
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   2100
      Width           =   3795
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   1320
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   1650
      Width           =   3795
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1200
      Width           =   3795
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   2355
   End
   Begin VB.Label Label5 
      Caption         =   "Size of Image:"
      Height          =   315
      Left            =   120
      TabIndex        =   11
      Top             =   2160
      Width           =   1155
   End
   Begin VB.Label Label4 
      Caption         =   "No. of Disks::"
      Height          =   315
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   1155
   End
   Begin VB.Label Label2 
      Caption         =   "Free Space:"
      Height          =   315
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   915
   End
   Begin VB.Label Label3 
      Caption         =   "Destination:"
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "Max Disks:"
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   915
   End
End
Attribute VB_Name = "NewImageForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' NewImageForm
' Written by Martin Mather 2005
' http://mmbeeb.mysite.wanadoo-members.co.uk/

Option Explicit
    
Private Sub cmdOK_Click()
    ' Create image
    Dim strFile As String
    
    strFile = CreateMMBFile
    If strFile <> "" Then
        Unload Me
        ' If successful open image
        ImageForm.OpenImageFile strFile
    End If
End Sub

Private Sub cmdCancel_Click()
    ' Close form
    Unload Me
End Sub

Private Sub Drive1_Change()
    ' Show Info
    Dim f As Long
    
    f = FreeSpace(Me.Drive1.Drive)
    Me.Text1.Text = ShowSize(f) & " free"
    Me.Text2.Text = dfsDiskCount(f) & " X " & ShowSize(DiskSize) & " disks"
    Me.Text3.Text = ShowSize(MMBSize(f)) & " bytes"
End Sub

Private Function dfsDiskCount(z As Long) As Long
    ' Calc number of dfs disks
On Error GoTo err_
    Dim m As Long
    
    m = Me.Text4.Text
    If m > MaxDisks Then m = MaxDisks
    
    dfsDiskCount = (z - Disk1Offset) \ DiskSize
    If dfsDiskCount > m Then dfsDiskCount = m
    Exit Function
    
err_:
    dfsDiskCount = 0
End Function

Private Function MMBSize(z As Long) As Long
    ' Calc size of MMB image file
    Dim i As Long
    
    i = dfsDiskCount(z)
    If i > 0 Then
        MMBSize = Disk1Offset + i * DiskSize
    End If
End Function

Private Function FreeSpace(drvPath As String) As Long
On Error GoTo err_
    ' Return free space on drive
    Dim fs, d
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set d = fs.GetDrive(fs.GetDriveName(drvPath))
    If d.FreeSpace > MaxDisks * DiskSize + Disk1Offset Then
        FreeSpace = MaxDisks * DiskSize + Disk1Offset
    Else
        FreeSpace = d.FreeSpace
    End If
exit_:
    Exit Function
err_:
    If Err.Number <> 429 Then
        eBox "FreeSpace"
    End If
    ' If error occurs allow maximum size
    FreeSpace = Disk1Offset + MaxDisks * DiskSize
    Resume exit_
End Function

Private Function CreateMMBFile() As String
    ' Create MMB image file
On Error GoTo err_
    Dim n As String
    'Dim z As Long
    Dim i As Long
    Dim p As Long
    Dim x As Long
    Dim f As Long
    Dim y As Long
    Dim bs(0 To DiskTableSize - 1) As Byte
    Dim ds(0 To DiskSize - 1) As Byte
    Dim varFile As Variant
    Dim strFile As String
    
    CreateMMBFile = ""
    varFile = glrCommonFileOpenSave(glrOFN_OVERWRITEPROMPT, _
                    OpenFile:=False, _
                    Filter:=glrAddFilterItem("", "MMB Image", "*.mmb"), _
                    Hwnd:=Me.Hwnd)
    
    If Not IsNull(varFile) Then
        n = varFile
    
        Me.MousePointer = vbHourglass
    
        f = FreeSpace(Me.Drive1.Drive)
        
        i = dfsDiskCount(f)
        'z = MMBSize(f) / MMCSecSize ' MMC sectors
    
        If i > 0 Then
            Debug.Print "Create : "; n
            Debug.Print , i & " disks"
            'Debug.Print , z & " MMC sectors"
            
            If Dir(n, vbNormal) <> "" Then
                Kill n ' already prompted in save dialog
            End If
        
            ' "Boot Sector"
            
            ' Disk table
            For x = 0 To i - 1
                y = (x + 1) * 16
                bs(y + 15) = DiskUnformatted
            Next
            
            For x = i To MaxDisks - 1
                y = (x + 1) * 16
                bs(y + 15) = DiskInvalid
            Next
            
            ' default disks (on boot)
            For x = 0 To 3
                bs(x) = x
            Next
            
            ' Open file
            f = FreeFile
            Open n For Binary Access Write As f
            ' Write "boot sector"
            Put f, , bs
            ' Blank sectors
            For x = 0 To i - 1
                Put f, , ds
            Next
            ' Close file
            Close f
            
            xBox "Image created!"
            CreateMMBFile = n
        End If
    End If
exit_:
On Error Resume Next
    Close f
    Me.MousePointer = vbDefault
    Exit Function
err_:
    eBox "Create Image"
    Resume exit_
End Function

Private Sub Form_Load()
    Drive1_Change
    Me.Text4.Text = MaxDisks
End Sub

Private Sub Text4_Change()
    Drive1_Change
End Sub
