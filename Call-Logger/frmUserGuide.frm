VERSION 5.00
Begin VB.Form frmUserGuide 
   Caption         =   "User guide"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7920
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6870
   ScaleWidth      =   7920
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstHelp 
      Height          =   3180
      ItemData        =   "frmUserGuide.frx":0000
      Left            =   120
      List            =   "frmUserGuide.frx":000D
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtHelpScript 
      Height          =   3255
      Left            =   1440
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<  &Back"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3360
      Width           =   975
   End
   Begin VB.Image imgLoading 
      Height          =   3255
      Left            =   4920
      Top             =   120
      Width           =   2895
   End
   Begin VB.Image imgStartingOff 
      Height          =   3255
      Left            =   4920
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmUserGuide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBack_Click()
Unload frmUserGuide
End Sub

Private Sub lstHelp_Click()
If lstHelp.ListIndex = 0 Then ' "Starting off"
 txtHelpScript = "First of all, enter some data into the text boxes on the main form. After filling in all the fields you wish to fill in click 'Submit' to submit your data. A box saying 'Data saved' will then appear, the default directory is in the directory for Visual Basic 6. This file will contain only data about phonecalls to the help desk only"
ElseIf lstHelp.ListIndex = 1 Then ' "Saving"
 txtHelpScript = "After entering the data you wish to submit on the main form titled 'Call Logger' Click 'Submit'. This will save the data you entered for you to view later."
ElseIf lstHelp.ListIndex = 2 Then ' "Loading"
 txtHelpScript = "Click on the 'View records' button on the main form. This will take you to the form to view all the records that have been previously entered. The caption of the form will display what record you are viewing out of how many records there are. You may press 'Next  >' to view the next record or '<  Back' to view the previous record."
ElseIf lstHelp.ListIndex = 3 Then '

End If
End Sub
