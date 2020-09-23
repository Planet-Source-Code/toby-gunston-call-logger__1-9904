VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H8000000A&
   Caption         =   "Call logger"
   ClientHeight    =   5835
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6210
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5835
   ScaleWidth      =   6210
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox cboCompany 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   600
      Width           =   2775
   End
   Begin VB.Timer tmrTimeMoveOut 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5760
      Top             =   1080
   End
   Begin VB.Timer tmrTimeMoveIn 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5760
      Top             =   600
   End
   Begin VB.Timer tmrTimeNow 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5760
      Top             =   120
   End
   Begin VB.TextBox txtTimeNow 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox cboCall 
      Height          =   315
      ItemData        =   "frmMain.frx":030A
      Left            =   1080
      List            =   "frmMain.frx":032F
      TabIndex        =   7
      Text            =   "(Select a name)"
      Top             =   4560
      Width           =   2055
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   2760
      TabIndex        =   9
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmdViewRecords 
      Caption         =   "&View records"
      Height          =   375
      Left            =   3960
      TabIndex        =   10
      Top             =   5280
      Width           =   1695
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "&Submit"
      Height          =   375
      Left            =   1080
      TabIndex        =   8
      Top             =   5280
      Width           =   1695
   End
   Begin VB.TextBox txtAction 
      Height          =   735
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   3720
      Width           =   4575
   End
   Begin VB.TextBox txtResponse 
      Height          =   735
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   2880
      Width           =   4575
   End
   Begin VB.TextBox txtProblem 
      Height          =   735
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   2040
      Width           =   4575
   End
   Begin VB.TextBox txtDate 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox txtTime 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox txtCaller 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label lblCall 
      Caption         =   "Call taken by:"
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label lblAction 
      Caption         =   "Action:"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label lblResponse 
      Caption         =   "Response:"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label lblProblem 
      Caption         =   "Problem:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label lblDate 
      Caption         =   "Date:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label lblTime 
      Caption         =   "Time:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label lblPerson 
      Caption         =   "Caller:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lblCompany 
      Caption         =   "Company:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      ToolTipText     =   "nowt"
      Top             =   600
      Width           =   735
   End
   Begin VB.Menu FileMenu 
      Caption         =   "&File"
      Begin VB.Menu ExitCommand 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu EditCommand 
      Caption         =   "&Edit"
      Begin VB.Menu ShowTime 
         Caption         =   "Show time"
      End
      Begin VB.Menu FontSizeCommand 
         Caption         =   "&Font size"
         Begin VB.Menu FontSize8 
            Caption         =   "8"
         End
         Begin VB.Menu FontSize10 
            Caption         =   "10"
         End
         Begin VB.Menu FontSize12 
            Caption         =   "12"
         End
         Begin VB.Menu FontSize14 
            Caption         =   "14"
         End
      End
      Begin VB.Menu NormalSizeCommand 
         Caption         =   "&Make form normal size"
      End
   End
   Begin VB.Menu HelpMenu 
      Caption         =   "&Help"
      Begin VB.Menu AboutCommand 
         Caption         =   "About"
      End
      Begin VB.Menu ComplainCommand 
         Caption         =   "Complain"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Record As PhoneRecord
Dim RecordLen As Long
Dim CurrentRecord As Long
Dim LastRecord As Long

Public Sub SaveCurrentRecord()
Open "PrevCallers.dat" For Random As #1 Len = RecordLen
 Record.Company = cboCompany.Text
 Record.Caller = txtCaller.Text
 Record.Time = txtTime.Text
 Record.Date = txtDate.Text
 Record.Problem = txtProblem.Text
 Record.Response = txtResponse.Text
 Record.Action = txtAction.Text
 Record.Call = cboCall.Text
 Put #1, CurrentRecord, Record
Close #1
End Sub

Private Sub AboutCommand_Click()
frmAbout.Show
End Sub

Private Sub cboCall_Change()
'Basic search options for the combo box.
'Typing in a single letter into the combo box will find
'the first entry with that letter and display it
If cboCall.Text = "s" Then
   cboCall.Text = "Stacey Cutting"
End If
End Sub

Private Sub cmdClear_Click()
'This is the code to clear all the textboxes on
'frmRecords
cboCompany = "(Select a company)"
txtCaller = ""
txtTime = ""
txtDate = ""
txtProblem = ""
txtResponse = ""
txtAction = ""
cboCall = "(Select a name)"
End Sub

Private Sub cmdSubmit_Click()
On Error GoTo SaveError

'Calculate the length of a record
RecordLen = Len(Record)
'Create the file.
'If it already exists it doesn't matter, the file will
'just be opened then closed
Open "PrevCallers.dat" For Random As #1 Len = RecordLen
Close #1
'Update gCurrentRecord
CurrentRecord = 1
'Find what is the last record number of the file
LastRecord = FileLen("PrevCallers.dat") / RecordLen
'If the file was just created (I.E. LastRecord=0) then update
'gLastRecord to 1
If LastRecord = 0 Then
 LastRecord = 1
End If
'Display the current record
CurrentRecord = LastRecord
ShowCurrentRecord
SaveCurrentRecord

'Add a new blank record. This is where the next entry will
'go. The user will never be able to access this record.
'It is restriced from being viewed as there is no record
'of it even existing in frmRecords
'the number
LastRecord = LastRecord + 1
Record.Company = ""
Record.Caller = ""
Record.Time = ""
Record.Date = ""
Record.Problem = ""
Record.Response = ""
Record.Action = ""
Record.Call = ""
Open "PrevCallers.dat" For Random As #1 Len = RecordLen
 Put #1, LastRecord, Record
Close #1
'Update current record
CurrentRecord = LastRecord
'Display record just created
ShowCurrentRecord
'Give the focus to the cboCompany field
cboCompany.SetFocus
MsgBox "Data saved successfully"
Exit Sub

SaveError:
MsgBox "Error writing data to that file, the data wasn't saved!" & vbCrLf & vbCrLf & "Possible reason:   File is write protected", vbCritical, "Error writing to file"
End Sub

Private Sub cmdViewRecords_Click()
frmRecords.Show
Unload frmRecords
frmRecords.Show
End Sub

Private Sub ComplainCommand_Click()
frmComplain.Show
End Sub

Private Sub ExitCommand_Click()
End
End Sub

Private Sub FontSize10_Click()
cboCompany.Font.Size = 10
txtCaller.Font.Size = 10
txtTime.Font.Size = 10
txtDate.Font.Size = 10
txtProblem.Font.Size = 10
txtProblem.Font.Size = 10
txtResponse.Font.Size = 10
txtAction.Font.Size = 10
End Sub

Private Sub FontSize12_Click()
cboCompany.Font.Size = 12
txtCaller.Font.Size = 12
txtTime.Font.Size = 12
txtDate.Font.Size = 12
txtProblem.Font.Size = 12
txtProblem.Font.Size = 12
txtResponse.Font.Size = 12
txtAction.Font.Size = 12
End Sub

Private Sub FontSize14_Click()
cboCompany.Font.Size = 14
txtCaller.Font.Size = 14
txtTime.Font.Size = 14
txtDate.Font.Size = 14
txtProblem.Font.Size = 14
txtProblem.Font.Size = 14
txtResponse.Font.Size = 14
txtAction.Font.Size = 14
End Sub

Private Sub FontSize16_Click()
cboCompany.Font.Size = 16
txtCaller.Font.Size = 16
txtTime.Font.Size = 16
txtDate.Font.Size = 16
txtProblem.Font.Size = 16
txtProblem.Font.Size = 16
txtResponse.Font.Size = 16
txtAction.Font.Size = 16
End Sub

Private Sub FontSize8_Click()
cboCompany.Font.Size = 8
txtCaller.Font.Size = 8
txtTime.Font.Size = 8
txtDate.Font.Size = 8
txtProblem.Font.Size = 8
txtProblem.Font.Size = 8
txtResponse.Font.Size = 8
txtAction.Font.Size = 8
End Sub


Private Sub Form_Load() 'MAIN
'Create the CL Credits.txt file for the about screen.
'Its created here everytime the program starts up
'so if anyone deletes the text file
'not knowing its needed for the program or if
'someone deletes it on purpose so as to try and
'pass the program off as theirs it wont matter as the
'text file will be created next time the program
'starts.
Open "CL Credits.txt" For Output As #1
Print #1, "CALL LOGGER"
Print #1,
Print #1,
Print #1,
Print #1,
Print #1, "CODED BY"
Print #1,
Print #1, "Toby Gunston"
Print #1,
Print #1,
Print #1, "June-July 2000"
Print #1,
Close #1
End Sub

Private Sub NormalSizeCommand_Click()
frmMain.WindowState = 0
frmMain.Height = 6525
frmMain.Width = 6330
End Sub

Private Sub ShowTime_Click()
If txtTimeNow.Left = 6240 Then 'Check to see if time textbox is out of the screen
    tmrTimeMoveIn.Enabled = True 'Enable the timer to animate moving in the time textbox
ElseIf txtTimeNow.Left = 4200 Then 'Check to see if time textbox is on the screen
    tmrTimeMoveOut.Enabled = True 'Enable the timer to animate moving out the time textbox
End If
End Sub

Private Sub tmrTimeMoveIn_Timer()
txtTimeNow.Visible = True
While txtTimeNow.Left > 4200
    txtTimeNow.Left = txtTimeNow.Left - 10
Wend
If txtTimeNow.Left = 4200 Then
    tmrTimeNow.Enabled = True
    tmrTimeMoveIn.Enabled = False
End If
End Sub

Private Sub tmrTimeMoveOut_Timer()
While txtTimeNow.Left < 6240
    txtTimeNow.Left = txtTimeNow.Left + 10
Wend
If txtTimeNow.Left = 6240 Then
    tmrTimeNow.Enabled = False
    tmrTimeMoveOut.Enabled = False
    txtTimeNow.Visible = False
End If
End Sub

Private Sub tmrTimeNow_Timer()
txtTimeNow = Time
End Sub

Private Sub txtDate_GotFocus()
txtDate = Date
End Sub

Private Sub txtTime_GotFocus()
txtTime = Time
End Sub



Public Sub ShowCurrentRecord()
'Leave in to make real pro work
End Sub
