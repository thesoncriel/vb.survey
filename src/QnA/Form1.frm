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
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton btnStart2 
      Caption         =   "��������2: �ɸ��׽�Ʈ"
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   720
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�������� ����"
      Height          =   1095
      Left            =   1320
      TabIndex        =   1
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton btnStart1 
      Caption         =   "��������1: �¶��ΰ���"
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
UserID = InputBox("����� ID�� �����ּ���.", "ID�� �ʿ��մϴ�.")
    If UserID = "" Then
    MsgBox "���α׷��� ������ �� �����ϴ�~!", , "���ϻ� -_-;;"
    Close #2
    Exit Sub
    End If
For i = 1 To LineNum
Get #2, i, Access
If UserID = RTrim(Access.ID) Then
    If MsgBox("�̹� ���� ID�� �������縦 �����̽��ϴ�." & vbCrLf & "�ٽ� �Ͻðڽ��ϱ�?", vbYesNo, "�̹� �ѹ��ϼ̳׿� -��-") = vbNo Then
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
