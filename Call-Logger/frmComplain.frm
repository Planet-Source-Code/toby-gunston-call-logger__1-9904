VERSION 5.00
Begin VB.Form frmComplain 
   Caption         =   "Enter your complaint"
   ClientHeight    =   2895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2895
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBack 
      Caption         =   "<  &Back"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit complaint"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   975
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   720
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "Complaint:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "frmComplain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBack_Click()
Unload frmComplain
End Sub

Private Sub cmdSubmit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If cmdSubmit.Left = 1560 Then 'If centre
 cmdSubmit.Left = 120         'move left
ElseIf cmdSubmit.Left = 120 Then  'If left
 cmdSubmit.Left = 3000        'move right
ElseIf cmdSubmit.Left = 3000 Then 'If right
 cmdSubmit.Left = 1560        'move centre
End If
End Sub


