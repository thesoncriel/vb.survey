VERSION 5.00
Begin VB.Form frmQnAmain1 
   BorderStyle     =   1  '���� ����
   Caption         =   "QnA Prog"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmQnAmain1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmQnAmain1.frx":038A
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.CheckBox Chk 
      BackColor       =   &H00E9C9A9&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   3255
   End
   Begin VB.CheckBox Chk 
      BackColor       =   &H00E9C9A9&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   3255
   End
   Begin VB.CheckBox Chk 
      BackColor       =   &H00E9C9A9&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   3255
   End
   Begin VB.CheckBox Chk 
      BackColor       =   &H00E9C9A9&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Label lblNext 
      BackStyle       =   0  '����
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
      Left            =   3600
      TabIndex        =   5
      Top             =   2400
      Width           =   975
   End
   Begin VB.Line Line3 
      X1              =   3120
      X2              =   3480
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line4 
      X1              =   3120
      X2              =   3480
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line5 
      X1              =   3120
      X2              =   3480
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line6 
      X1              =   3120
      X2              =   3480
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line7 
      X1              =   3480
      X2              =   3480
      Y1              =   1800
      Y2              =   2880
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00404000&
      BorderStyle     =   3  '��
      X1              =   3480
      X2              =   4440
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   5  '���-��-��
      X1              =   4200
      X2              =   4200
      Y1              =   1920
      Y2              =   3000
   End
   Begin VB.Line Line2 
      X1              =   3120
      X2              =   3120
      Y1              =   1560
      Y2              =   3120
   End
   Begin VB.Line Line1 
      BorderStyle     =   4  '���-��
      X1              =   120
      X2              =   4560
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label lblDesc 
      BackStyle       =   0  '����
      Height          =   1095
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "frmQnAmain1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim i As Byte
Dim qData As QnaGeneral

If QnA_Num = 1 Then
ReDim nRndNum(1 To 10)
For i = 1 To 10
nRndNum(i) = RndNumProg(0, 9, 10)
Next
End If
Picnum = nRndNum(QnA_Num)
Me.Picture = LoadPicture(App.Path & "\Data\" & "QnA" & Picnum & ".jpg")

For i = 0 To 3
Chk(i).Caption = (i + 1) & ". "
Next

Open FilePath & FileName0 For Random As #1 Len = Len(qData)
Get #1, QnA_Num, qData
With qData
Me.Caption = .Title
lblDesc.Caption = .Desc
For i = 0 To 3
Chk(i).Caption = .Sel(i)
Next
End With
Close #1

Me.Caption = "����" & QnA_Num & ": " & RTrim(Me.Caption)
lblDesc.Caption = RTrim(lblDesc.Caption)
For i = 0 To 3
Chk(i).Caption = i + 1 & ":  " & RTrim(Chk(i).Caption)
Next

If Len(Chk(2).Caption) = 4 Then Chk(2).Visible = False
If Len(Chk(3).Caption) = 4 Then Chk(3).Visible = False
End Sub

Private Sub lblNext_Click()
Dim qType As Byte
Dim qData As QnaGeneral
Dim User As String, i As Byte

'///����ڰ� ���ù��� �����Ͽ����� �˾Ƴ�.
If Chk(0).Value = 1 Then
User = "2"
Else
User = "0"
End If
For i = 1 To 3
If Chk(i).Value = 1 Then
User = User & "2"
Else
User = User & "0"
End If
Next
If User = "0000" Then
MsgBox "������ �������� �����̽��ϴ�.", vbExclamation, "������ �����ϼ���~!"
Exit Sub
End If
'///�� =_=;;

'///����ڰ� ������ ������ ResultTemp�� �ӽ� �����Ų��.
ResultTemp = ResultTemp & User

QnA_Num = QnA_Num + 1
'///���� Ƚ���� 10���� �ʰ��� ��� ��� â�� ���.
If QnA_Num > 10 Then
Unload Me
frmQnAResult.Show
Exit Sub
End If
'///�� ^^


'///���� ������ �ϱ����� �̸� DB����
'///���� ������ �ش�Ǵ� ���¸� �˾ƿ���,
'///�� ���� ���� ������ qType�� ���� ��Ŵ.
'///0�̸� �ߺ� ���� �Ұ�(Radio Button)
'///1�̸� �ߺ� ���� ����(Check Box)
'///2�̸� �ְ���
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
'///��;;
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblNext.ForeColor = &HB07C4C
End Sub
Private Sub lblNext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblNext.ForeColor = &H0
End Sub
