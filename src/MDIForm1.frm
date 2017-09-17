VERSION 5.00
Begin VB.MDIForm MDIForm1 
   Appearance      =   0  '평면
   BackColor       =   &H00FF8080&
   ClientHeight    =   7395
   ClientLeft      =   885
   ClientTop       =   450
   ClientWidth     =   10425
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   ScrollBars      =   0   'False
   WindowState     =   2  '최대화
   Begin VB.Menu File 
      Caption         =   "파일"
      NegotiatePosition=   1  '왼쪽
      Begin VB.Menu LogIn 
         Caption         =   "새 로그인"
      End
      Begin VB.Menu PersnalJoin 
         Caption         =   "사용자 등록 -_-;;"
      End
      Begin VB.Menu Modify 
         Caption         =   "개인정보 변경"
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu QnA 
         Caption         =   "설문조사"
         Begin VB.Menu QnA1 
            Caption         =   "온라인 게임"
         End
         Begin VB.Menu QnA2 
            Caption         =   "심리 테스트"
         End
         Begin VB.Menu Weekend 
            Caption         =   "주말 레저 활동 조사"
         End
         Begin VB.Menu Trafic 
            Caption         =   "교통 문제"
         End
         Begin VB.Menu Kyoyang 
            Caption         =   "교양 수준 조사"
         End
      End
      Begin VB.Menu Custom 
         Caption         =   "사용자 정의"
      End
      Begin VB.Menu QnAEdit 
         Caption         =   "설문조사 에디터"
      End
      Begin VB.Menu line2 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "나가기"
      End
   End
   Begin VB.Menu NumBase 
      Caption         =   "숫자야구게임"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type ProgSet
Font As String * 6
First As String * 3
Linef As String * 2
End Type

Private Sub Custom_Click()
frmQnACustom.Show
End Sub

Private Sub MDIForm_Load()
Dim Picnum As Byte
Dim sdata As ProgSet

Me.Caption = MyInfo
FilePath = App.Path & "\Data\"

Randomize
Picnum = Int(7 * Rnd + 1)
MDIForm1.Picture = LoadPicture(FilePath & Picnum & ".jpg")

Open FilePath & ProgramSetting For Random As #1 Len = Len(sdata)
If LOF(1) = 0 Then
MsgBox ProgramSetting & " 파일의 데이터를 찾을 수 없습니다!!", vbCritical, "파일이 없어요 ㅠ.ㅠ 흑흑 일부러 지웠죠;;"
End
End If

Get #1, 1, sdata
If sdata.First = "000" Then
    If MsgBox("처음 시작 하시는군요 ^^" & vbCrLf & "Guest로 시작 하시겠습니까?.", vbYesNo + vbInformation, "첫 스타트 (~^^)~") = vbYes Then
    UserID = "Guest"
    LogInOK = 1
    MDIForm1.Caption = MyInfo & UserID
    End If
ElseIf sdata.First = "100" Then
sdata.Font = "Open= "
sdata.First = "001"
sdata.Linef = vbCrLf
Put #1, 1, sdata
Close #1
Exit Sub
ElseIf sdata.First = "999" Then
MsgBox "이거 파일 수정해서 보고 계시다는거 다 압니다 -_-;;", vbExclamation, "이, 이런 장난은 젭알 자제를 ㅠ.ㅠ;;"
sdata.Font = "Open= "
sdata.First = "000"
sdata.Linef = vbCrLf
Put #1, 1, sdata
Close #1
End
End If

sdata.Font = "Open= "
sdata.First = Format(Str(Val(sdata.First) + 1), "00#")
sdata.Linef = vbCrLf
Put #1, 1, sdata
Close #1

End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If LogInOK = 0 Then
frmLogIn.Show 1
LogInOK = LogInOK + 1
End If
End Sub



'///파일메뉴 클릭시 이벤트들
Private Sub LogIn_Click()
frmLogIn.Show 1
End Sub
Private Sub PersnalJoin_Click()
frmPersnal.Show
End Sub
Private Sub Modify_Click()
If UserID = "Guest" Or UserID = "" Then
MsgBox "Guest님은 사용자 정보 변경이 불가능합니다.", vbInformation, "손님에겐 기록된 개인정보가 없습니다."
Exit Sub
End If
EditMode = 1
frmPersnal.Show
End Sub
Private Sub QnA1_Click()
LoadFile ("OnlineGame.gdb")
End Sub
Private Sub QnA2_Click()
LoadFile ("MindTest.gdb")
End Sub
Private Sub Trafic_Click()
LoadFile ("Kyotong.gdb")
End Sub
Private Sub Weekend_Click()
LoadFile ("Weekend.gdb")
End Sub
Private Sub Kyoyang_Click()
LoadFile ("Kyoyang.gdb")
End Sub
Private Sub QnAEdit_Click()
frmQnaEdit.Show
End Sub
Private Sub Exit_Click()
End
End Sub

Private Sub NumBase_Click()
frmBaseball.Show
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


