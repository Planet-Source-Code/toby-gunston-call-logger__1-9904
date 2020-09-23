VERSION 5.00
Begin VB.Form frmInstall 
   BackColor       =   &H80000001&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Choose directory to install to:"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7290
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   7290
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtInstallTo 
      Height          =   375
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1680
      Width           =   3375
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H80000001&
      Caption         =   "&Next       >"
      Height          =   375
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Click here to continue"
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H80000001&
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Click here to cancel the installation"
      Top             =   2760
      Width           =   1095
   End
   Begin VB.DirListBox dir1 
      Height          =   1890
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   3255
   End
   Begin VB.DriveListBox drv1 
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label lblInstallTo 
      BackColor       =   &H80000003&
      BackStyle       =   0  'Transparent
      Caption         =   "Install to:"
      BeginProperty Font 
         Name            =   "@Batang"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label lblInstall 
      BackStyle       =   0  'Transparent
      Caption         =   "Please choose the directory you wish to install to:"
      BeginProperty Font 
         Name            =   "@Batang"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
   End
End
Attribute VB_Name = "frmInstall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
Cancel
End Sub

Private Sub cmdNext_Click()
Me.Hide
frmInstalling.Show
End Sub


Private Sub dir1_Change()
txtInstallTo = dir1.Path
End Sub

Private Sub drv1_Change()
On Error GoTo error

dir1.Path = drv1.Drive
Exit Sub

error:
If drv1.Drive = "a:" Then
 MsgBox "No disk in drive", vbCritical, "Error"
Else
 MsgBox "No data in drive", vbCritical, "Error"
End If
End Sub



