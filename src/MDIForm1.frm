VERSION 5.00
Begin VB.MDIForm MDIForm1 
   Appearance      =   0  '���
   BackColor       =   &H00FF8080&
   ClientHeight    =   7395
   ClientLeft      =   885
   ClientTop       =   450
   ClientWidth     =   10425
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   ScrollBars      =   0   'False
   WindowState     =   2  '�ִ�ȭ
   Begin VB.Menu File 
      Caption         =   "����"
      NegotiatePosition=   1  '����
      Begin VB.Menu LogIn 
         Caption         =   "�� �α���"
      End
      Begin VB.Menu PersnalJoin 
         Caption         =   "����� ��� -_-;;"
      End
      Begin VB.Menu Modify 
         Caption         =   "�������� ����"
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu QnA 
         Caption         =   "��������"
         Begin VB.Menu QnA1 
            Caption         =   "�¶��� ����"
         End
         Begin VB.Menu QnA2 
            Caption         =   "�ɸ� �׽�Ʈ"
         End
         Begin VB.Menu Weekend 
            Caption         =   "�ָ� ���� Ȱ�� ����"
         End
         Begin VB.Menu Trafic 
            Caption         =   "���� ����"
         End
         Begin VB.Menu Kyoyang 
            Caption         =   "���� ���� ����"
         End
      End
      Begin VB.Menu Custom 
         Caption         =   "����� ����"
      End
      Begin VB.Menu QnAEdit 
         Caption         =   "�������� ������"
      End
      Begin VB.Menu line2 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "������"
      End
   End
   Begin VB.Menu NumBase 
      Caption         =   "���ھ߱�����"
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
MsgBox ProgramSetting & " ������ �����͸� ã�� �� �����ϴ�!!", vbCritical, "������ ����� ��.�� ���� �Ϻη� ������;;"
End
End If

Get #1, 1, sdata
If sdata.First = "000" Then
    If MsgBox("ó�� ���� �Ͻô±��� ^^" & vbCrLf & "Guest�� ���� �Ͻðڽ��ϱ�?.", vbYesNo + vbInformation, "ù ��ŸƮ (~^^)~") = vbYes Then
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
MsgBox "�̰� ���� �����ؼ� ���� ��ôٴ°� �� �дϴ� -_-;;", vbExclamation, "��, �̷� �峭�� ���� ������ ��.��;;"
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



'///���ϸ޴� Ŭ���� �̺�Ʈ��
Private Sub LogIn_Click()
frmLogIn.Show 1
End Sub
Private Sub PersnalJoin_Click()
frmPersnal.Show
End Sub
Private Sub Modify_Click()
If UserID = "Guest" Or UserID = "" Then
MsgBox "Guest���� ����� ���� ������ �Ұ����մϴ�.", vbInformation, "�մԿ��� ��ϵ� ���������� �����ϴ�."
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








'///������ �ҷ����� ���ν���
Private Sub LoadFile(Filename As String)
Dim qType As Byte, i As Byte
Dim qData As QnaGeneral
Dim Access As QnaGeneralResult, LineNum As Byte
Dim IDs As String * 12

If QnA_Num > 0 Then
    Beep
    If MsgBox("�̹� �������縦 �����ϰ� ��ʴϴ�." & vbCrLf & "�ٽ� �Ͻðڽ��ϱ�?", vbYesNo + vbQuestion, "�̹� ������...") = vbYes Then
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
    If MsgBox("�̹� ���� ID�� �������縦 �����̽��ϴ�." & vbCrLf & "�ٽ� �Ͻðڽ��ϱ�?", vbYesNo, "�̹� �ѹ��ϼ̳׿� -��-") = vbNo Then
    Close #2
    Exit Sub
    Else
    '///�̹� �ѹ� �������縦 ���ƴٸ� �� ������� ��� �����͸� ����� ��.
    '///���� ���� FixLineNum�� ����ڰ� �������縦 �ߴ� �����Ͱ� �ִ� ��ġ�� ����.
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


