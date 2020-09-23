VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Zipfile Scanner"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   6615
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   6615
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraInfo 
      Caption         =   "Information"
      Height          =   1695
      Left            =   120
      TabIndex        =   4
      Top             =   5640
      Width           =   6375
      Begin VB.TextBox txtPath 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1200
         Width           =   5415
      End
      Begin VB.CommandButton cmdFileInfo 
         Caption         =   "Local File Info"
         Height          =   375
         Left            =   4320
         TabIndex        =   14
         Top             =   720
         Width           =   1815
      End
      Begin VB.CommandButton cmdComment 
         Caption         =   "Zip Comment"
         Height          =   375
         Left            =   4320
         TabIndex        =   13
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtRatio 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtCrc 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtCompressed 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtSize 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblPath 
         Caption         =   "Path:"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label lblRatio 
         Caption         =   "Ratio:"
         Height          =   255
         Left            =   2520
         TabIndex        =   11
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblCRC 
         Caption         =   "CRC:"
         Height          =   255
         Left            =   2520
         TabIndex        =   9
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lblCompressed 
         Caption         =   "Compressed:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblSize 
         Caption         =   "Size:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   375
      End
   End
   Begin ZipfileScanner.usrTreeView usrZipFiles 
      Height          =   2295
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   4048
   End
   Begin VB.FileListBox filFiles 
      Height          =   3015
      Left            =   3360
      Pattern         =   "*.zip"
      TabIndex        =   2
      Top             =   120
      Width           =   3135
   End
   Begin VB.DirListBox dirDirectories 
      Height          =   2565
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   3135
   End
   Begin VB.DriveListBox drvDrives 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuComment 
         Caption         =   "&Zip comment..."
      End
      Begin VB.Menu mnuLocalInfo 
         Caption         =   "&Local file info..."
      End
      Begin VB.Menu mnuHash0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FileSel As Long
Dim OldDrive As String

Private Sub cmdComment_Click()
  ShowComment
End Sub

Private Sub cmdFileInfo_Click()
  If FileSel = -1 Then
    MsgBox "First select a file", vbInformation
  Else
    ShowFileInfo FileSel
  End If
End Sub

Private Sub dirDirectories_Change()
  EraseInfo
  filFiles.Path = dirDirectories.Path
End Sub

Private Sub drvDrives_Change()
  EraseInfo
  On Error Resume Next
  dirDirectories.Path = drvDrives.Drive
  If Err.Number Then
    If MsgBox("Device is not ready", vbCritical + vbRetryCancel) = vbRetry Then
      drvDrives_Change
    Else
      drvDrives.Drive = OldDrive
    End If
  End If
End Sub

Private Sub filFiles_Click()
  EraseInfo
  LoadFiles filFiles.Path + IIf(Right(filFiles.Path, 1) <> "\", "\", "") + filFiles.List(filFiles.ListIndex)
End Sub

Private Sub Form_Load()
  OldDrive = drvDrives.Drive
  EraseInfo
  usrZipFiles.RaiseInSelection = True
End Sub

Private Sub mnuAbout_Click()
  frmAbout.Show vbModal
End Sub

Private Sub mnuComment_Click()
  cmdComment_Click
End Sub

Private Sub mnuExit_Click()
  End
End Sub

Private Sub EraseInfo()
  FileSel = -1
  usrZipFiles.Clear
  txtCompressed.Text = ""
  txtSize.Text = ""
  txtRatio.Text = ""
  txtCrc.Text = ""
  txtPath.Text = ""
End Sub

Private Sub mnuLocalInfo_Click()
  cmdFileInfo_Click
End Sub

Private Sub usrZipFiles_SelectItem(ByVal Changed As Boolean, ByVal ItemId As Long, ByVal ItemText As String, ByVal Hierarchy As Long)
  FileSel = ItemId
  ShowInfo ItemId
End Sub
