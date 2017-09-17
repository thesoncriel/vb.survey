VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton btnStart2 
      Caption         =   "설문조사2: 심리테스트"
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   720
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "설문조사 편집"
      Height          =   1095
      Left            =   1320
      TabIndex        =   1
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton btnStart1 
      Caption         =   "설문조사1: 온라인게임"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnStart1_Click()
Dim qType As Byte, i As Byte
Dim qData As QnaGeneral
Dim Access As QnaGeneralResult, LineNum As Byte

FileName0 = "OnlineGame.gdb"
FileName1 = Left(FileName0, Len(FileName0) - 4) & ".rdb"
Open FilePath & FileName1 For Random As #2 Len = Len(Access)
LineNum = LOF(2) / Len(Access)
If LineNum = 0 Then
LineNum = 1
End If

For i = 1 To LineNum
Get #2, i, Access
If UserID = DataCleaner(Access.ID) Then
    If MsgBox("이미 같은 ID로 설문조사를 끝내셨습니다." & vbCrLf & "다시 하시겠습니까?", vbYesNo, "이미 한번하셨네용 -ㅁ-") = vbNo Then
    Close #2
    Exit Sub
    Else
    '///이미 한번 설문조사를 마쳤다면 그 사용자의 결과 데이터를 재수정 함.
    '///전역 변수 FixLineNum는 사용자가 설문조사를 했던 데이터가 있는 위치를 저장.
    FixLineNum = i
    Exit For
    End If
End If
Next
Close #2

QnA_Num = QnA_Num + 1

Open FilePath & FileName0 For Random As #2 Len = Len(qData)
Get #2, QnA_Num, qData
qType = Val(qData.qType)

Close #2
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

Private Sub btnStart2_Click()
Dim qType As Byte, i As Byte
Dim qData As QnaGeneral
Dim Access As QnaGeneralResult, LineNum As Byte

FileName0 = "MindTest.gdb"
FileName1 = Left(FileName0, Len(FileName0) - 4) & ".rdb"
Open FilePath & FileName1 For Random As #2 Len = Len(Access)
LineNum = LOF(2) / Len(Access)
If LineNum = 0 Then
LineNum = 1
End If
UserID = InputBox("사용자 ID를 적어주세요.", "ID가 필요합니다.")
    If UserID = "" Then
    MsgBox "프로그램을 실행할 수 없습니다~!", , "뭐하삼 -_-;;"
    Close #2
    Exit Sub
    End If
For i = 1 To LineNum
Get #2, i, Access
If UserID = RTrim(Access.ID) Then
    If MsgBox("이미 같은 ID로 설문조사를 끝내셨습니다." & vbCrLf & "다시 하시겠습니까?", vbYesNo, "이미 한번하셨네용 -ㅁ-") = vbNo Then
    Close #2
    Exit Sub
    Else
    FixLineNum = i
    Exit For
    End If
End If
Next
Close #2

QnA_Num = QnA_Num + 1

Open FilePath & FileName0 For Random As #2 Len = Len(qData)
Get #2, QnA_Num, qData
qType = Val(qData.qType)

Close #2
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

Private Sub Command2_Click()
frmQnaEdit.Show
End Sub

Private Sub Form_Load()
FilePath = App.Path & "\Data\"
End Sub
