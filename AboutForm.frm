VERSION 5.00
Begin VB.Form AboutForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DFS Imager"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   Icon            =   "AboutForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton GoSite 
      Caption         =   "http://mmbeeb.mysite.wanadoo-members.co.uk/"
      Height          =   375
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1080
      Width           =   3735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   615
      Left            =   2520
      TabIndex        =   0
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "MMBeeb image manipulation, including disk image dragging to and from Windows Explorer."
      Height          =   615
      Left            =   720
      TabIndex        =   3
      Top             =   1680
      Width           =   4695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Created by Martin Mather"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   6135
   End
   Begin VB.Label LabelVer 
      Alignment       =   2  'Center
      Caption         =   "MMB Imager Version X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "AboutForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' AboutForm
' Written by Martin Mather 2005
' http://mmbeeb.mysite.wanadoo-members.co.uk/

Option Explicit

Private Const SW_SHOWNORMAL As Long = 1
Private Const SW_SHOWMAXIMIZED As Long = 3
Private Const SW_SHOWDEFAULT As Long = 10

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function ShellExecute Lib "shell32" _
   Alias "ShellExecuteA" _
  (ByVal Hwnd As Long, _
   ByVal lpOperation As String, _
   ByVal lpFile As String, _
   ByVal lpParameters As String, _
   ByVal lpDirectory As String, _
   ByVal nShowCmd As Long) As Long

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    LabelVer.Caption = "MMB Imager Version " & ProgVersion
End Sub

Private Sub GoSite_Click()
    Dim h As String
    
    h = GoSite.Caption
    ShellExecute Me.Hwnd, "open", h, 0, 0, SW_SHOWNORMAL
End Sub
