VERSION 5.00
Begin VB.Form frmQnAmain2 
   BorderStyle     =   1  '단일 고정
   Caption         =   "QnA Prog"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmQnAmain2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmQnAmain2.frx":038A
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.TextBox txtUser 
      Appearance      =   0  '평면
      BackColor       =   &H00E9C9A9&
      BorderStyle     =   0  '없음
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1680
      Width           =   3135
   End
   Begin VB.Label lblSize 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   2775
   End
   Begin VB.Line Line1 
      BorderStyle     =   4  '대시-점
      X1              =   0
      X2              =   4440
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line2 
      X1              =   3000
      X2              =   3000
      Y1              =   1560
      Y2              =   3120
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   5  '대시-점-점
      X1              =   4080
      X2              =   4080
      Y1              =   1920
      Y2              =   3000
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00404000&
      BorderStyle     =   3  '점
      X1              =   3360
      X2              =   4320
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line7 
      X1              =   3360
      X2              =   3360
      Y1              =   1800
      Y2              =   2880
   End
   Begin VB.Line Line6 
      X1              =   3000
      X2              =   3360
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line5 
      X1              =   3000
      X2              =   3360
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line4 
      X1              =   3000
      X2              =   3360
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line3 
      X1              =   3000
      X2              =   3360
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label lblNext 
      BackStyle       =   0  '투명
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00B07C4C&
      Height          =   615
      Left            =   3480
      TabIndex        =   2
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label lblDesc 
      BackStyle       =   0  '투명
      Height          =   1095
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "frmQnAmain2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim i As Byte
Dim qData As QnaGeneral
Dim Picnum As Byte

If QnA_Num = 1 Then
ReDim nRndNum(1 To 10)
For i = 1 To 10
nRndNum(i) = RndNumProg(0, 9, 10)
Next
End If
Picnum = nRndNum(QnA_Num)
Me.Picture = LoadPicture(App.Path & "\Data\" & "QnA" & Picnum & ".jpg")

Open FilePath & FileName0 For Random As #1 Len = Len(qData)
Get #1, QnA_Num, qData
With qData
Me.Caption = .Title
lblDesc.Caption = .Desc
End With
Close #1

Me.Caption = "문항" & QnA_Num & ": " & RTrim(Me.Caption)
lblDesc.Caption = RTrim(lblDesc.Caption)
End Sub

Private Sub lblNext_Click()
Dim qType As Byte
Dim qData As QnaGeneral

'///사용자가 선택문을 선택하였는지 알아냄.
If txtUser.Text = "" Then
MsgBox "의견을 쓰지 않았습니다..", vbExclamation, "주관식이예요 'ㅁ'a"
Exit Sub
End If
'///끝 =_=;;

FreeAnswer = txtUser.Text
ResultTemp = ResultTemp & "0000"

QnA_Num = QnA_Num + 1
If QnA_Num > 10 Then
Unload Me
frmQnAResult.Show
Exit Sub
End If

Open FilePath & FileName0 For Random As #1 Len = Len(qData)
Get #1, QnA_Num, qData
qType = Val(qData.qType)
Close #1
Unload Me
Select Case qType
Case 0
frmQnAmain0.Show
Case 1
frmQnAmain1.Show
Case 2
frmQnAmain2.Show
End Select
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblNext.ForeColor = &HB07C4C
End Sub
Private Sub lblNext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblNext.ForeColor = &H0
End Sub

Private Sub txtUser_Change()
Dim size As Integer
size = Len(txtUser.Text)
lblSize.Caption = size & " / " & "64"
lblSize.ForeColor = &H0
If size > 64 Then
lblSize.Caption = "위험! 글자가 짤릴 수도 있습니다" & lblSize.Caption
lblSize.ForeColor = &HFF
Else
End If
End Sub
