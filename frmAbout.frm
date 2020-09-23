VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About ZipFile Scanner"
   ClientHeight    =   2625
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5160
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2625
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3720
      TabIndex        =   0
      Top             =   2040
      Width           =   1245
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   1110
      TabIndex        =   4
      Top             =   480
      Width           =   3030
   End
   Begin VB.Line Linea 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      X1              =   105
      X2              =   5055
      Y1              =   1695
      Y2              =   1695
   End
   Begin VB.Label lblDescription 
      Alignment       =   2  'Center
      Caption         =   "Zip files analyzer. Copyright © 2001 by Paul Guerra. All rights reserved."
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   600
      TabIndex        =   1
      Top             =   960
      Width           =   3750
   End
   Begin VB.Label lblName 
      Caption         =   "ZipFile Scanner"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1065
      TabIndex        =   3
      Top             =   240
      Width           =   3030
   End
   Begin VB.Line Linea2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   120
      X2              =   5055
      Y1              =   1710
      Y2              =   1710
   End
   Begin VB.Label lblWarning 
      Caption         =   "Warning: This program is copyrighted. Its unauthorized reproduction may result in criminal charges."
      ForeColor       =   &H00000000&
      Height          =   705
      Left            =   135
      TabIndex        =   2
      Top             =   1920
      Width           =   3390
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  lblVersion.Caption = App.Major & "." & App.Minor & "." & App.Revision
End Sub
