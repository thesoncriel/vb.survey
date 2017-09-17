VERSION 5.00
Begin VB.Form frmQnAmain0 
   BorderStyle     =   1  '단일 고정
   Caption         =   "QnA Prog"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmQnAmain0.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmQnAmain0.frx":038A
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.OptionButton Rad 
      BackColor       =   &H00E9C9A9&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   3135
   End
   Begin VB.OptionButton Rad 
      BackColor       =   &H00E9C9A9&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   3135
   End
   Begin VB.OptionButton Rad 
      BackColor       =   &H00E9C9A9&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   3135
   End
   Begin VB.OptionButton Rad 
      BackColor       =   &H00E9C9A9&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   3135
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
      TabIndex        =   5
      Top             =   2400
      Width           =   975
   End
   Begin VB.Line Line2 
      X1              =   3000
      X2              =   3000
      Y1              =   1560
      Y2              =   3240
   End
   Begin VB.Line Line1 
      BorderStyle     =   4  '대시-점
      X1              =   0
      X2              =   4440
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label lblDesc 
      BackStyle       =   0  '투명
      Height          =   1095
      Left            =   360
      TabIndex        =   4
      Top             =   360
      Width           =   3975
   End
End
Attribute VB_Name = "frmQnAmain0"
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
For i = 0 To 3
Rad(i).Caption = .Sel(i)
Next
End With
Close #1

Me.Caption = "문항" & QnA_Num & ": " & RTrim(Me.Caption)
lblDesc.Caption = RTrim(lblDesc.Caption)
For i = 0 To 3
Rad(i).Caption = (i + 1) & ".  " & RTrim(Rad(i).Caption)
Next

If Len(Rad(2).Caption) = 4 Then Rad(2).Visible = False
If Len(Rad(3).Caption) = 4 Then Rad(3).Visible = False
End Sub

Private Sub lblNext_Click()
Dim qType As Byte, i As Byte
Dim qData As QnaGeneral
Dim User As String * 4

'///사용자가 선택문을 선택하였는지 알아냄.
For i = 0 To 3
If Rad(i).Value Then
User = Format(1000 / 10 ^ i, "000#")
End If
Next
If Val(User) = 0 Then
MsgBox "문장을 선택하지 않으셨습니다.", vbExclamation, "문장을 선택하세요~!"
Exit Sub
End If
'///끝 =_=;;

ResultTemp = ResultTemp & User

QnA_Num = QnA_Num + 1
'///질문 횟수가 10번을 초과할 경우 결과 창을 출력.
If QnA_Num > 10 Then
Unload Me
frmQnAResult.Show
Exit Sub
End If
'///끝 ^^


'///다음 질문을 하기전에 미리 DB에서
'///다음 질문에 해당되는 형태를 알아오고,
'///그 값을 전역 변수인 qType에 대입 시킴.
'///0이면 중복 선택 불가(Radio Button)
'///1이면 중복 선택 가능(Check Button)
'///2이면 주관형
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
'///끝;;
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblNext.ForeColor = &HB07C4C
End Sub
Private Sub lblNext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblNext.ForeColor = &H0
End Sub
