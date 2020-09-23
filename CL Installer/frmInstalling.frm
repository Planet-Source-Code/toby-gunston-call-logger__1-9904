VERSION 5.00
Begin VB.Form frmInstalling 
   BackColor       =   &H80000001&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Installing...........................  Please wait"
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
   Begin VB.CommandButton cmdFinish 
      BackColor       =   &H80000001&
      Caption         =   "&Finish"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Click here to finish the installation"
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H80000001&
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Click here to cancel the installation"
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Timer tmr4 
      Interval        =   4000
      Left            =   6720
      Top             =   1560
   End
   Begin VB.Timer tmr6 
      Interval        =   6000
      Left            =   6720
      Top             =   2520
   End
   Begin VB.Timer tmr2 
      Interval        =   2000
      Left            =   6720
      Top             =   600
   End
   Begin VB.Timer tmr3 
      Interval        =   3000
      Left            =   6720
      Top             =   1080
   End
   Begin VB.Timer tmr5 
      Interval        =   5000
      Left            =   6720
      Top             =   2040
   End
   Begin VB.Timer tmr1 
      Interval        =   1000
      Left            =   6720
      Top             =   120
   End
   Begin VB.Label lblDone 
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "Done..."
      BeginProperty Font 
         Name            =   "@Batang"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Label lbl5 
      BackColor       =   &H80000001&
      Caption         =   "n"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lbl4 
      BackColor       =   &H80000001&
      Caption         =   "n"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lbl3 
      BackColor       =   &H80000001&
      Caption         =   "n"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lbl2 
      BackColor       =   &H80000001&
      Caption         =   "n"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lbl1 
      BackColor       =   &H80000001&
      Caption         =   "n"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblInstalling 
      BackStyle       =   0  'Transparent
      Caption         =   "Installing"
      BeginProperty Font 
         Name            =   "@Batang"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "frmInstalling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
Cancel
End Sub

Private Sub cmdFinish_Click()
End
End Sub

Private Sub Form_Load()
On Error GoTo error

FileCopy (App.Path & "\CL.dat"), frmInstall.dir1.Path & "\Call Logger.exe"
Exit Sub

error:
MsgBox "There was an error installing the program, it DID NOT install successfully", vbCritical, "Error"
End
End Sub

Private Sub tmr1_Timer()
lbl1.Visible = True
tmr1.Enabled = False
End Sub

Private Sub tmr2_Timer()
lbl2.Visible = True
tmr2.Enabled = False
End Sub

Private Sub tmr3_Timer()
lbl3.Visible = True
tmr3.Enabled = False
End Sub

Private Sub tmr4_Timer()
lbl4.Visible = True
tmr4.Enabled = False
End Sub

Private Sub tmr5_Timer()
lbl5.Visible = True
tmr5.Enabled = False
End Sub

Private Sub tmr6_Timer()
lblDone.Visible = True
lblDone.BackStyle = 1
tmr6.Enabled = False
cmdFinish.Enabled = True
cmdCancel.Enabled = False
frmInstalling.Caption = "Installation complete, have a nice day"
End Sub


