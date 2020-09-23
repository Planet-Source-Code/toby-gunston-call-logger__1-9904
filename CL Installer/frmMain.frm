VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H80000001&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CALL LOGGER    setup program"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7290
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   7290
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H80000001&
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Click here to cancel the installation"
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H80000001&
      Caption         =   "&Next       >"
      Height          =   375
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label lblWelcome 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMain.frx":030A
      BeginProperty Font 
         Name            =   "@Batang"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
Cancel
End Sub

Private Sub cmdNext_Click()
Me.Hide
frmInstall.Show
End Sub


