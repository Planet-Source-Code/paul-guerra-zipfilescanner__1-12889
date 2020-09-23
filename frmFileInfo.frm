VERSION 5.00
Begin VB.Form frmFileInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File Information"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtExtraField 
      BackColor       =   &H8000000F&
      Height          =   1575
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   17
      Top             =   3960
      Width           =   4695
   End
   Begin VB.TextBox txtPath 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   1320
      Width           =   2295
   End
   Begin VB.TextBox txtComment 
      BackColor       =   &H8000000F&
      Height          =   1575
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   14
      Top             =   2040
      Width           =   4695
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   5760
      Width           =   1455
   End
   Begin VB.TextBox txtTime 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox txtDate 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox txtSize 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox txtCompressed 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox txtCrc 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtRatio 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lblExtraField 
      Caption         =   "Extra field:"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   3720
      Width           =   735
   End
   Begin VB.Label lblPath 
      Caption         =   "Path:"
      Height          =   255
      Left            =   2640
      TabIndex        =   15
      Top             =   960
      Width           =   495
   End
   Begin VB.Label lblComment 
      Caption         =   "Local comment:"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label lblTime 
      Caption         =   "Time modified:"
      Height          =   255
      Left            =   2640
      TabIndex        =   11
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblDate 
      Caption         =   "Date modified:"
      Height          =   255
      Left            =   2640
      TabIndex        =   9
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lblSize 
      Caption         =   "Size:"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   375
   End
   Begin VB.Label lblCompressed 
      Caption         =   "Compressed:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblCRC 
      Caption         =   "CRC:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   375
   End
   Begin VB.Label lblRatio 
      Caption         =   "Ratio:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   495
   End
End
Attribute VB_Name = "frmFileInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
  Unload Me
End Sub
