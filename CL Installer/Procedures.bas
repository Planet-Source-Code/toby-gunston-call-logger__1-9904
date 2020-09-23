Attribute VB_Name = "Procedures"
Public Sub Cancel()
Dim Message As String
Dim ButtonsAndIcon As Integer
Dim Title As String
Dim Response As Integer

Message = "Are you sure you want to quit?"
ButtonsAndIcon = vbYesNo + vbQuestion
Title = "Confirm exit"
Response = MsgBox(Message, ButtonsAndIcon, Title)

If Response = vbYes Then
    End
ElseIf Response = vbNo Then
End If
End Sub
