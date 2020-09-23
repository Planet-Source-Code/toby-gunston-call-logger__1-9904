VERSION 5.00
Begin VB.Form frmRecords 
   Caption         =   "View records"
   ClientHeight    =   7110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   ScaleHeight     =   7110
   ScaleWidth      =   6315
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   1800
      TabIndex        =   18
      Top             =   6600
      Width           =   2655
   End
   Begin VB.CommandButton cmdDeleteAll 
      Caption         =   "&Delete all records"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   6240
      Width           =   2655
   End
   Begin VB.TextBox cboCall 
      Height          =   315
      Left            =   1080
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5400
      Width           =   1815
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next      >"
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<      &Back"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   5880
      Width           =   1335
   End
   Begin VB.TextBox txtAction 
      Height          =   1125
      Left            =   1080
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4200
      Width           =   4215
   End
   Begin VB.TextBox txtResponse 
      Height          =   1125
      Left            =   1080
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3000
      Width           =   4215
   End
   Begin VB.TextBox txtProblem 
      Height          =   1125
      Left            =   1080
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1800
      Width           =   4215
   End
   Begin VB.TextBox txtDate 
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox txtTime 
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox txtCaller 
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   720
      Width           =   2775
   End
   Begin VB.TextBox cboCompany 
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label Label7 
      Caption         =   "Action:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "Response:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Problem:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Date:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Time:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Caller:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Company:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "frmRecords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Record As PhoneRecord
Dim RecordLen As Long
Dim CurrentRecord As Long
Dim LastRecord As Long

Private Sub cmdExit_Click()
'Close the file
Close #1
End
End Sub

Private Sub cmdBack_Click()
'If the current record is the first record, beep and display
'an error message. Otherwise, save the current record
'and go to the previous record
If CurrentRecord = 1 Then
 MsgBox "Beginning of file!", vbExclamation
Else
 SaveCurrentRecord
 CurrentRecord = CurrentRecord - 1
 ShowCurrentRecord
End If
'Give the focus to the txtName field
cboCompany.SetFocus
End Sub

Private Sub cmdClose_Click()
Unload frmRecords
End Sub

Private Sub cmdDeleteAll_Click()
Dim Message As String
Dim ButtonsAndIcon As Integer
Dim Title As String
Dim Response As Integer
On Error GoTo Error:

Message = "Are you sure you want to delete all the records?"
ButtonsAndIcon = vbYesNo + vbQuestion
Title = "Confirm file delete"
Response = MsgBox(Message, ButtonsAndIcon, Title)

If Response = vbYes Then
    Kill "PrevCallers.dat"
    Unload frmRecords
    frmRecords.Show
ElseIf Response = vbNo Then
End If
Exit Sub

Error:
MsgBox "Error deleting records", vbCritical, "Error"
End Sub

Private Sub cmdNext_Click()
'If the current record is the last record
'beep and display an error message. Otherwise,
'save the current record and skip to the next record.
If CurrentRecord = LastRecord - 1 Then
 MsgBox "End of file!", vbExclamation, "Error"
Else
 SaveCurrentRecord
 CurrentRecord = CurrentRecord + 1
 ShowCurrentRecord
End If
'Give the focus to the txtName field
cboCompany.SetFocus
End Sub

Private Sub cmdSubmit_Click()
On Error GoTo SaveError:
SaveCurrentRecord
MsgBox "Saved"
Exit Sub

SaveError:
MsgBox "Not saved"
End Sub

Private Sub Form_Load() 'RECORDS
'Calculate the length of a record
RecordLen = Len(Record)
'Open the file for random access. If the file
'does not exist then it is created.
Open "PrevCallers.dat" For Random As #1 Len = RecordLen
'Update gCurrentRecord
CurrentRecord = 1
'Find what is the last record number of the file
LastRecord = FileLen("PrevCallers.dat") / RecordLen
'End If
'Display the current record
ShowCurrentRecord

'If the file was just created (I.E. LastRecord=0) then update
'LastRecord to 1
If LastRecord = 0 Then
 frmRecords.Caption = "View records - there are no records to view"
 cmdNext.Enabled = False
 cmdBack.Enabled = False
 cmdDeleteAll.Enabled = False
End If

Close #1

End Sub

Public Sub SaveCurrentRecord()
Open "PrevCallers.dat" For Random As #1 Len = RecordLen
 Record.Company = cboCompany
 Record.Caller = txtCaller
 Record.Time = txtTime
 Record.Date = txtDate
 Record.Problem = txtProblem
 Record.Response = txtResponse
 Record.Action = txtAction
 Record.Call = cboCall
 Put #1, CurrentRecord, Record
Close #1
End Sub

Public Sub ShowCurrentRecord()
'Fill Person with data of current record
 Close #1
 Open "PrevCallers.dat" For Random As #1 Len = RecordLen
 Get #1, CurrentRecord, Record
 'Display RECORD
 cboCompany.Text = Trim(Record.Company)
 txtCaller.Text = Trim(Record.Caller)
 txtTime.Text = Trim(Record.Time)
 txtDate.Text = Trim(Record.Date)
 txtProblem.Text = Trim(Record.Problem)
 txtResponse.Text = Trim(Record.Response)
 txtAction.Text = Trim(Record.Action)
 cboCall.Text = Trim(Record.Call)
 'Display the current record number in the caption of the form
 frmRecords.Caption = "Record " + Str(CurrentRecord) + " of " + _
                                  Str(LastRecord - 1)
 Close #1
End Sub

