Attribute VB_Name = "DataCleaner"

Public Function DCLogin(Data As String)
With frmLogIn.Hidden
.Text = Data
.Text = RTrim(.Text)
DCLogin = .Text
.Text = ""
End With
End Function

Public Function DCPersnal(Data As String)
With frmPersnal.Hidden
.Text = Data
.Text = RTrim(.Text)
DCPersnal = .Text
.Text = ""
End With
End Function

Public Function DCQnAResult(Data As String)
With frmQnAResult.Hidden
.Text = Data
.Text = RTrim(.Text)
DCQnAResult = .Text
.Text = ""
End With
End Function
