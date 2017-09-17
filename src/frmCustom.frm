VERSION 5.00
Begin VB.Form frmQnACustom 
   BorderStyle     =   1  '단일 고정
   Caption         =   "사용자 정의형 설문조사"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2775
   Icon            =   "frmCustom.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmCustom.frx":038A
   ScaleHeight     =   3225
   ScaleWidth      =   2775
   Begin VB.FileListBox File1 
      Appearance      =   0  '평면
      BackColor       =   &H00E9C9A9&
      Height          =   1470
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label lblGo 
      BackStyle       =   0  '투명
      Caption         =   "Open"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H006E543B&
      Height          =   615
      Left            =   720
      TabIndex        =   1
      Top             =   1800
      Width           =   1095
   End
End
Attribute VB_Name = "frmQnACustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblGo.ForeColor = &H6E543B
End Sub

Private Sub lblGo_Click()
Dim Filename As String
Filename = File1.List(File1.ListIndex)
LoadFile (Filename)
End Sub

Private Sub File1_DblClick()
Dim Filename As String
Filename = File1.List(File1.ListIndex)
LoadFile (Filename)
End Sub

Private Sub Form_Load()
File1.Path = App.Path & "\data\"
File1.Pattern = "*.gdb"
End Sub

'///파일을 불러내는 프로시저
Private Sub LoadFile(Filename As String)
Dim qType As Byte, i As Byte
Dim qData As QnaGeneral
Dim Access As QnaGeneralResult, LineNum As Byte
Dim IDs As String * 12

If QnA_Num > 0 Then
    Beep
    If MsgBox("이미 설문조사를 실행하고 계십니다." & vbCrLf & "다시 하시겠습니까?", vbYesNo + vbQuestion, "이미 실행중...") = vbYes Then
    Unload frmQnAmain0
    Unload frmQnAmain1
    Unload frmQnAmain2
    End If
End If

IDs = UserID
FileName0 = Filename
FileName1 = Left(FileName0, Len(FileName0) - 4) & ".rdb"
Open FilePath & FileName1 For Random As #2 Len = Len(Access)
LineNum = LOF(2) / Len(Access)
If LineNum = 0 Then
LineNum = 1
End If

If UserID <> "Guest" Then
For i = 1 To LineNum
Get #2, i, Access
If IDs = Access.ID Then
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
End If
Close #2

QnA_Num = 1
FreeAnswer = ""
ResultTemp = ""

Open FilePath & FileName0 For Random As #2 Len = Len(qData)
Get #2, QnA_Num, qData
qType = Val(qData.qType)

Close #2
Select Case qType
Case 0
frmQnAmain0.Show
Case 1
frmQnAmain1.Show
Case 2
frmQnAmain2.Show
End Select
End Sub

Private Sub lblGo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblGo.ForeColor = &H0
End Sub
