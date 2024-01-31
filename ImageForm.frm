VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form ImageForm 
   Caption         =   "MMB Explorer"
   ClientHeight    =   4965
   ClientLeft      =   2415
   ClientTop       =   2100
   ClientWidth     =   8460
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ImageForm.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   4965
   ScaleWidth      =   8460
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Enabled         =   0   'False
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   915
   End
   Begin VB.CommandButton cmdRestoreDisk 
      Caption         =   "Restore"
      Enabled         =   0   'False
      Height          =   315
      Left            =   7560
      TabIndex        =   5
      Top             =   0
      Width           =   795
   End
   Begin VB.CommandButton cmdUnlockDisk 
      Caption         =   "Unlock"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5280
      TabIndex        =   3
      Top             =   0
      Width           =   675
   End
   Begin VB.CommandButton cmdLockDisk 
      Caption         =   "Lock"
      Enabled         =   0   'False
      Height          =   315
      Left            =   4680
      TabIndex        =   2
      Top             =   0
      Width           =   555
   End
   Begin VB.CommandButton cmdSelectAllDisks 
      Caption         =   "Select All"
      Enabled         =   0   'False
      Height          =   315
      Left            =   2760
      TabIndex        =   1
      Top             =   0
      Width           =   915
   End
   Begin VB.CommandButton cmdKillDisk 
      Caption         =   "Kill"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6960
      TabIndex        =   4
      Top             =   0
      Width           =   555
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   4590
      Width           =   8460
      _ExtentX        =   14923
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageListIcons 
      Left            =   1500
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ImageForm.frx":030A
            Key             =   "UnformattedDisk"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ImageForm.frx":0624
            Key             =   "Disk"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ImageForm.frx":093E
            Key             =   "LockedDisk"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageListSmallIcons 
      Left            =   2100
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ImageForm.frx":0C58
            Key             =   "UnformattedDisk"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ImageForm.frx":0F72
            Key             =   "LockedDisk"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ImageForm.frx":128C
            Key             =   "Disk"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListViewDisks 
      Height          =   4215
      Left            =   120
      TabIndex        =   7
      Top             =   360
      Width           =   8280
      _ExtentX        =   14605
      _ExtentY        =   7435
      View            =   1
      Arrange         =   1
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDragMode     =   1
      OLEDropMode     =   1
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageListIcons"
      SmallIcons      =   "ImageListSmallIcons"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Disk"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "File Count"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Menu FileMenu 
      Caption         =   "&File"
      Begin VB.Menu OpenImage 
         Caption         =   "&Open Image"
      End
      Begin VB.Menu NewImage 
         Caption         =   "&New Image"
      End
      Begin VB.Menu CloseImage 
         Caption         =   "&Close Image"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu DiskTableMenu 
      Caption         =   "Disk &Table"
      Begin VB.Menu RebuildTable 
         Caption         =   "&Rebuild Disk Table"
         Enabled         =   0   'False
      End
      Begin VB.Menu DiskTableSmallIcons 
         Caption         =   "&Small Icons"
         Checked         =   -1  'True
      End
      Begin VB.Menu DiskTableLargeIcons 
         Caption         =   "&Large Icons"
         Checked         =   -1  'True
      End
      Begin VB.Menu HideUnformattedDisks 
         Caption         =   "&Hide Unformatted Disks"
      End
      Begin VB.Menu SaveFullImage 
         Caption         =   "&Extract Full Images (200 KB)"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu ProtectionMenu 
      Caption         =   "&Protection"
      Begin VB.Menu ProtectionDisabled 
         Caption         =   "&Disabled"
      End
   End
   Begin VB.Menu HelpMenu 
      Caption         =   "Help"
      Begin VB.Menu AboutMe 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "ImageForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ImageForm
' Written by Martin Mather 2005
' http://mmbeeb.mysite.wanadoo-members.co.uk/

Option Explicit

Private mDiskTable As disktable_type

Private mstrImagePathname As String
Private mboolHideUnformattedDisks As Boolean
Private mboolProtection As Boolean
Private mboolResize As Boolean

Private mintDisksSelected As Integer
Private mboolDraggingDisk As Boolean
Private mintMoveDisk As Integer
Private mboolMoveDisk As Boolean
Private mintTargetDisk As Integer
Private mstrInfo As String

Private mboolDisksFullSize As Boolean

' Disk actions
Private Const DiskAct_Kill = 1
Private Const DiskAct_Restore = 2
Private Const DiskAct_Lock = 3
Private Const DiskAct_Unlock = 4

Private m_iButton As Integer ' BUG Workaround - MS Article ID : 240946

Private Sub DiskActButtons()
    ' Enable disk activity buttons
    Dim e As Boolean
    Dim k As Boolean
    Dim x As Long
    Dim intDiskNo As Integer
    
    e = mintDisksSelected > 0
    k = e
    If k And mboolProtection Then
        ' Check if any disks in selection locked
        ' If yes then disable 'Kill' button
        With ListViewDisks
            For x = 1 To .ListItems.Count
                If .ListItems(x).Selected Then
                    intDiskNo = LVDiskNo(x)
                    k = k And Not mDiskTable.Disk(intDiskNo).ReadOnly
                End If
            Next
        End With
    End If
    
    cmdKillDisk.Enabled = k
    cmdRestoreDisk.Enabled = e
    cmdLockDisk.Enabled = e
    cmdUnlockDisk.Enabled = e
End Sub

Private Sub AboutMe_Click()
    ' Open 'About' form
    AboutForm.Show 1, Me
End Sub

Private Sub CloseImage_Click()
    ' Close MMB Image
    mstrImagePathname = ""
    mDiskTable.ImageName = ""
    Caption = "MMB Imager"
    ListViewDisks.ListItems.Clear
    ListViewDisks.Enabled = False
    mintDisksSelected = 0
    DiskActButtons
    CloseImage.Enabled = False
    RebuildTable.Enabled = False
    cmdSelectAllDisks.Enabled = False
    StatusBar1.SimpleText = ""
    cmdRefresh.Enabled = False
End Sub

Private Sub cmdKillDisk_Click()
    ' Disks: Kill disk
    DiskActions DiskAct_Kill
End Sub

Private Sub cmdLockDisk_Click()
    ' Disks: Lock disk
    DiskActions DiskAct_Lock
End Sub

Private Sub cmdRefresh_Click()
    ' Refresh
    OpenImageFile mstrImagePathname
End Sub

Private Sub cmdRestoreDisk_Click()
    ' Disks: Restore disk
    DiskActions DiskAct_Restore
End Sub

Private Sub cmdSelectAllDisks_Click()
    ' Disks: Select all
    LVSelectAll ListViewDisks
    ListViewDisks_Click
    ListViewDisks.SetFocus
End Sub

Private Sub cmdUnlockDisk_Click()
    ' Disks: Unlock disk
    DiskActions DiskAct_Unlock
End Sub

Private Sub DiskTableLargeIcons_Click()
    ' Disks: Large icons
    DiskTableLargeIcons.Checked = True
    DiskTableSmallIcons.Checked = False
    With ListViewDisks
        .View = lvwIcon
        .Sorted = True
        .Arrange = lvwAutoLeft
    End With
End Sub

Private Sub DiskTableSmallIcons_Click()
    ' Disks: Small icons
    DiskTableLargeIcons.Checked = False
    DiskTableSmallIcons.Checked = True
    With ListViewDisks
        .View = lvwSmallIcon
        .Sorted = True
        .Arrange = lvwAutoLeft
    End With
End Sub

Private Sub Form_Load()
    ' Initialise form
    Dim strOpenFile As String
    
    DiskTableSmallIcons_Click
    ListViewDisks.Enabled = False
    ListViewDisks.SortKey = 0
    ListViewDisks.Sorted = True
    
    mboolProtection = True
    mboolHideUnformattedDisks = False
    mboolDisksFullSize = False
    SaveFullImage.Checked = False
    cmdRefresh.Enabled = False
    
    strOpenFile = Trim(Command()) ' Command line argument
    If strOpenFile <> "" Then
        ' Strip marks
        If Left(strOpenFile, 1) = """" Then
            strOpenFile = Mid(strOpenFile, 2)
        End If
        If Right(strOpenFile, 1) = """" Then
            strOpenFile = Left(strOpenFile, Len(strOpenFile) - 1)
        End If
    End If
    
    If strOpenFile <> "" Then
        Debug.Print "Open: "; strOpenFile
        OpenImageFile strOpenFile
    End If
End Sub

Private Sub Form_Resize()
    ' Resize form
On Error Resume Next
    With Me.ListViewDisks
        .Width = Me.Width - 350
        .Height = Me.Height - .Top - 1100
    End With
End Sub

Private Sub HideUnformattedDisks_Click()
    ' Toggle "hide unformatted disks"
    Dim x As Integer
    
    With HideUnformattedDisks
        .Checked = Not .Checked
        mboolHideUnformattedDisks = .Checked
    End With
    
    For x = 0 To MaxDisks - 1
        If Not mDiskTable.Disk(x).Formatted Then
            DItem x
        End If
    Next
    DList False
End Sub

Private Sub ListViewDisks_Click()
    ' Select disk(s)
    mintDisksSelected = LVSelectCount(ListViewDisks)
    DiskActButtons
    UpdateInfoBox
End Sub

Private Sub ListViewDisks_DblClick()
    ' Open disk using 'DFSImager'
    Dim intDiskNo As Integer
    Dim strKey As String
    
    If mintDisksSelected = 1 Then
        strKey = ListViewDisks.SelectedItem.Key
        If strKey <> "" Then
            intDiskNo = CInt(Mid(strKey, 2))
            If mDiskTable.Disk(intDiskNo).Formatted Then
                DFSImager mDiskTable.ImageName, intDiskNo
            End If
        End If
    End If
End Sub

Private Sub ListViewDisks_KeyUp(KeyCode As Integer, Shift As Integer)
    ListViewDisks_Click
End Sub

Private Sub ListViewDisks_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button And (ListViewDisks.HitTest(x, y) Is Nothing) Then
        m_iButton = Button
        LVSelectNothing ListViewDisks
    End If
End Sub

Private Sub ListViewDisks_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If (Button = 0) And m_iButton Then
        ListViewDisks_Click
        m_iButton = 0
    End If
End Sub

Private Sub ListViewDisks_OLECompleteDrag(Effect As Long)
    ' Disk: Complete drag from image
    
    If mboolDraggingDisk And mboolMoveDisk Then
        ' Kill source disk if "moved"
        ' i.e. a copy has been created so delete the original
    
        With mDiskTable
            ' Debug.Print "Kill "; mintMoveDisk
            ' Set disk table entry of copy to that of source
            .Disk(mintTargetDisk) = .Disk(mintMoveDisk)
            .Disk(mintTargetDisk).DiskNo = mintTargetDisk
            
            ' Kill source
            .Disk(mintMoveDisk).Formatted = False
        End With
        
        ' Update disk table
        ModifyDiskTable mDiskTable, mintTargetDisk, True
        ModifyDiskTable mDiskTable, mintMoveDisk, False
        
        LVSelectNothing ListViewDisks
        
        DItem mintTargetDisk
        DItem mintMoveDisk
        DList False
    End If
    
    mboolDraggingDisk = False
End Sub

Private Sub ListViewDisks_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    ' Disk: Object dropped on image disk
    Dim lngIndex As Long
    Dim varFilename
    Dim intDiskNo As Integer
    
    ''Debug.Print "lvd dd "; mboolDraggingDisk
    
    lngIndex = LVIndexUnderPoint(ListViewDisks, x, y)
    If lngIndex > 0 Then
        ' Target is an unformatted/empty disk being pointed at
        intDiskNo = LVDiskNo(lngIndex)
        mintTargetDisk = intDiskNo ' used when moving disk within image
    Else
        ' Target is first unformatted disk
        intDiskNo = -1
    End If
    
    If Data.GetFormat(vbCFFiles) Then
        Me.MousePointer = vbHourglass

        For Each varFilename In Data.Files
            intDiskNo = ImportDFSImage(intDiskNo, mDiskTable, _
                            CStr(varFilename))
            If intDiskNo < 0 Then
                ' Prevent "source" being killed
                ' in ListViewDisks_OLECompleteDrag
                mboolMoveDisk = False
                ' Stop if error occurs
                Exit For
            Else
                DItem intDiskNo
            End If
        Next
        
        If Not mboolMoveDisk Then DList False
        
        Me.MousePointer = vbDefault
    End If
End Sub

Private Sub ListViewDisks_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    ' Object being dragged over ListViewDisks
    ' Check and return whether it can be 'dropped'
    Dim lngIndex As Long
    Dim intDiskNo As Integer
    Dim bytSide As Byte
    Dim strInf As String
    
    ''Debug.Print "LVD : do", mboolDraggingDisk, Shift
    mboolMoveDisk = False
    strInf = mstrInfo
    
    If mboolDraggingDisk And Data.Files.Count > 1 Then
        ' Can't drag more than on disk within
        ' image on to ListViewDisks
        Effect = vbDropEffectNone
        
    ElseIf Data.GetFormat(vbCFFiles) Then
    
        If State = vbOver Then
            ' Index of disk being pointed at
            lngIndex = LVIndexUnderPoint(ListViewDisks, x, y)
        End If
        
        If lngIndex > 0 Then
            ' Drag on to disk
            intDiskNo = LVDiskNo(lngIndex)
            
            If Not mDiskTable.Disk(intDiskNo).Formatted _
                    And Data.Files.Count = 1 Then
                ' Target is an unformatted disk
                ' (only 1 disk being dragged)
                
                ' Move disk if CTRL is not pressed
                ' (when drag from within ListViewFiles)
                mboolMoveDisk = Shift <> 2 And mboolDraggingDisk 'True
                Effect = vbDropEffectCopy And Effect
                
                ' Give user feedback
                If mboolDraggingDisk Then
                    If mboolMoveDisk Then
                        strInf = "Move disk " & mintMoveDisk & _
                                    " to disk " & intDiskNo & _
                                    " (Press CTRL to copy)"
                    Else
                        strInf = "Copy disk " & mintMoveDisk & _
                                    " to disk " & intDiskNo
                    End If
                Else
                    strInf = "Copy to disk " & intDiskNo
                End If
                
                With ListViewDisks.ListItems(lngIndex)
                    .Selected = True
                    .Selected = False
                End With
            Else
                ' Target disk is formatted and contains
                ' files - drop not allowed
                
                Effect = vbDropEffectNone
            End If
        Else
            ' Drag to "empty space"
            ' Target will be first unformatted disk(s)
            
            Effect = vbDropEffectCopy And Effect
        End If
    End If
    
    StatusBar1.SimpleText = strInf
End Sub

Private Sub ListViewDisks_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
    ' Disk: Drag from ListViewDisks started
    ' Note: Will leave disk image files in the temporary folder
    '       whether drag successful or not
    Dim intDiskNo As Integer
    Dim bytSide As Byte
    Dim x As Long
    Dim srcf As Long
    Dim strFile As String
    
    mboolDraggingDisk = True
    mboolMoveDisk = False
    
    ' Save files from image to temporary directory
    Me.MousePointer = vbHourglass
    
    With ListViewDisks
        srcf = FreeFile
        Open mDiskTable.ImageName For Binary Access Read As srcf
        
        For x = 1 To .ListItems.Count
            If .ListItems(x).Selected Then
                ' Ignore unformatted disks
                intDiskNo = LVDiskNo(x)
                        
                mintMoveDisk = intDiskNo
                If mDiskTable.Disk(intDiskNo).Formatted Then
                    strFile = ExportDFSDisk(srcf, _
                                mDiskTable.Disk(intDiskNo), _
                                mboolDisksFullSize)
                    If strFile = "" Then
                        Exit For ' Stop on error
                    Else
                        Data.Files.Add strFile
                        Debug.Print ">> "; strFile; " == disk "; intDiskNo
                    End If
                End If
            End If
        Next
        
        Close srcf
    End With
    
    If Data.Files.Count > 0 Then
        Data.SetData , vbCFFiles
        AllowedEffects = vbDropEffectCopy ' Or vbDropEffectMove
    Else
        AllowedEffects = vbDropEffectNone
    End If
    
    Me.MousePointer = vbDefault
End Sub

Private Sub NewImage_Click()
    ' Create new Image
    NewImageForm.Show 1, Me
End Sub

Private Sub OpenImage_Click()
    ' Find and open MMB Image
    Dim varFile As Variant
    Dim strFile As String
    
    varFile = glrCommonFileOpenSave( _
            Filter:=glrAddFilterItem("", "MMB Images", "*.mmb"), Hwnd:=Me.Hwnd)
    
    If Not IsNull(varFile) Then
        strFile = varFile
        OpenImageFile strFile
    End If
End Sub

Public Sub OpenImageFile(strFile As String)
    ' Open MMB Image
    Me.MousePointer = vbHourglass

    mboolResize = True
    mstrImagePathname = strFile
    Caption = strFile
    mDiskTable = ReadDiskTable(strFile)
    ListViewDisks.ListItems.Clear
    DList True
    ListViewDisks.Enabled = True
    CloseImage.Enabled = True
    RebuildTable.Enabled = True
    cmdRefresh.Enabled = True

    Me.MousePointer = vbDefault
End Sub

Private Sub UpdateInfoBox()
    ' Status: Show info about selected disks on status bar
    Dim x As Integer
    Dim intDiskCount As Integer
    Dim intUnformattedCount As Integer
    Dim strInf As String
    Dim intDiskNo As Integer
    Dim boolAll As Boolean

    boolAll = LVSelectCount(ListViewDisks) = 0
    
    With ListViewDisks
        For x = 0 To .ListItems.Count - 1
            If .ListItems(x + 1).Selected Or boolAll Then
                intDiskNo = LVDiskNo(x + 1)
                
                With mDiskTable.Disk(intDiskNo)
                    If .ValidDisk Then
                    
                        intDiskCount = intDiskCount + 1
                        If Not .Formatted Then
                            intUnformattedCount = intUnformattedCount + 1
                        End If
                    End If
                End With
                
            End If
        Next
        
        strInf = ShowInt(intDiskCount, " disk")
        
        If intUnformattedCount > 0 Then
            strInf = strInf & ", " & _
                ShowInt(intUnformattedCount, " disk") & " unformatted"
        End If
    End With
    
    mstrInfo = strInf
    StatusBar1.SimpleText = strInf
End Sub

Private Sub DList(boolRefresh As Boolean)
    ' Disks: Rebuild ListViewDisks
    Dim x As Integer
    Dim lngView As Long
    
'    Debug.Print "DList start"
    
    If mboolResize Then
        ListViewDisks.View = lvwList ' Force column with resize
        mboolResize = False
    End If
    
    If boolRefresh Then
        For x = 0 To MaxDisks - 1
            DItem x
        Next
    End If
    
    cmdSelectAllDisks.Enabled = ListViewDisks.ListItems.Count > 0
    
'    Debug.Print "DList mid"
    
    UpdateInfoBox
    
    lngView = IIf(DiskTableSmallIcons.Checked, lvwSmallIcon, lvwIcon)
    With ListViewDisks
        If .View <> lngView Then
            .View = lngView
            .Arrange = lvwAutoLeft
            .SortKey = 0
            .Sorted = True
        End If
    End With
    
'    Debug.Print "DList end"
End Sub

Private Sub DItem(intDiskNo As Integer)
    ' Disks: Add disk to ListViewDisks
    Dim strIcon As String
    Dim strName As String
    Dim d As ListItem
    Dim i As Integer
    
    i = DIndex(intDiskNo)
    
    With mDiskTable.Disk(intDiskNo)
        If Not .Formatted And mboolHideUnformattedDisks Then
        
            If i >= 0 Then
                ListViewDisks.ListItems.Remove i
            End If
            
        ElseIf .ValidDisk Then
        
            If Not .Formatted Then
                strIcon = "UnformattedDisk"
            ElseIf .ReadOnly Then
                strIcon = "LockedDisk"
            Else
                strIcon = "Disk"
            End If
            
            strName = intDiskNo & ":"
            
            ' spaces required for correct sorting
            If intDiskNo < 10 Then strName = " " & strName
            If intDiskNo < 100 Then strName = " " & strName
            
            If .Formatted Then
                strName = strName & " " & .DiskTitle
            End If
           
            If i = -1 Then
                'i = x * 2 + s + 1
                Set d = ListViewDisks.ListItems.Add '(i)
                'Debug.Print "ADL "; d.Index, i
                d.Key = DiskKey(intDiskNo)
                d.Icon = strIcon
                d.SmallIcon = strIcon
                d.Text = strName
            Else
                Set d = ListViewDisks.ListItems(i)
                
                If d.Icon <> strIcon Or d.Text <> strName Then
                    d.Icon = strIcon
                    d.SmallIcon = strIcon
                    d.Text = strName
                End If
            End If
            
            Set d = Nothing
        End If
    End With
End Sub

Private Function DIndex(intDiskNo As Integer) As Integer
    ' Return ListViewDisks index of DiskNo
    Dim x As Long
    Dim k As String
    
    DIndex = -1
    k = DiskKey(intDiskNo)
    For x = 1 To ListViewDisks.ListItems.Count
        If ListViewDisks.ListItems(x).Key = k Then
            DIndex = x
            Exit Function
        End If
    Next
End Function

Private Sub ProtectionDisabled_Click()
    ' Enable/disable protection
    mboolProtection = Not mboolProtection
    ProtectionDisabled.Checked = Not mboolProtection
    DiskActButtons
End Sub

Private Sub RebuildTable_Click()
    ' Rebuild disk table
    Me.MousePointer = vbHourglass
    RebuildDiskTable mDiskTable
    Me.MousePointer = vbDefault
    cmdRefresh_Click
End Sub

Private Sub SaveFullImage_Click()
    ' Toggle whether to save full 200Kb disk images,
    ' or minimum size dependent on sectors used
    With SaveFullImage
        .Checked = Not .Checked
         mboolDisksFullSize = .Checked
    End With
End Sub

Private Sub DiskActions(intAction As Integer)
    ' Do action on selected disks
    Dim x As Long
    Dim intDiskNo As Integer
    
    With ListViewDisks
        For x = 1 To .ListItems.Count
            If .ListItems(x).Selected Then
                intDiskNo = LVDiskNo(x)
                        
                With mDiskTable.Disk(intDiskNo)
                    Select Case intAction
                        Case DiskAct_Kill
                            .Formatted = False
                        Case DiskAct_Restore
                            .Formatted = True
                        Case DiskAct_Lock
                            .ReadOnly = True
                        Case DiskAct_Unlock
                            .ReadOnly = False
                    End Select
                End With
                
                ModifyDiskTable mDiskTable, intDiskNo, False
                DItem intDiskNo
            End If
        Next
    End With
    
    DList False
    DiskActButtons
    ListViewDisks.SetFocus
End Sub

Private Function DiskKey(intDiskNo As Integer)
    ' Create 'ListViewDisks' key
    DiskKey = "D" & Format(intDiskNo, "000")
End Function

Private Function LVDiskNo(lngIndex As Long) As Integer
    ' Extract Disk No from 'ListViewDisks' key
    LVDiskNo = CInt(Mid(ListViewDisks.ListItems(lngIndex).Key, 2))
End Function
