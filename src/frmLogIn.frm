VERSION 5.00
Begin VB.Form frmLogIn 
   Appearance      =   0  '���
   BackColor       =   &H80000005&
   BorderStyle     =   0  '����
   Caption         =   "�α��� ^^;;"
   ClientHeight    =   1950
   ClientLeft      =   6495
   ClientTop       =   6750
   ClientWidth     =   3420
   Icon            =   "frmLogIn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmLogIn.frx":1CCA
   ScaleHeight     =   1950
   ScaleWidth      =   3420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.TextBox Hidden 
      Height          =   270
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtPass 
      Appearance      =   0  '���
      Height          =   270
      IMEMode         =   3  '��� ����
      Left            =   840
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox txtID 
      Appearance      =   0  '���
      Height          =   270
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label lblAcc 
      Alignment       =   2  '��� ����
      Appearance      =   0  '���
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "�Ϸ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "::::Log-In::::"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   30
      Width           =   3375
   End
   Begin VB.Label lblLogin 
      Alignment       =   1  '������ ����
      BackStyle       =   0  '����
      Caption         =   "Pass"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1020
      Width           =   615
   End
   Begin VB.Label lblLogin 
      Alignment       =   1  '������ ����
      BackStyle       =   0  '����
      Caption         =   "ID"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   660
      Width           =   615
   End
End
Attribute VB_Name = "frmLogIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Label1_Click()
Beep
If MsgBox("������ ��ư�� �����̽��ϴ�." & vbCrLf & "Guest ���� �Ͻðڽ��ϱ�?", vbYesNo + vbQuestion, "���̵� ������ �׳� ���ó׿� -��-;;") = vbYes Then
UserID = "Guest"
UserInfo = "�湮��"
MDIForm1.Caption = MyInfo & UserInfo & "(" & UserID & ")"
LogInOK = 1
Unload Me
End If
End Sub

Private Sub lblAcc_Click()
Dim sdata() As PersnalData
Dim LineNum As Byte, i As Byte

Open FilePath & PersnalDataFile For Random As #1 Len = Len(sdata(1))

If LOF(1) = 0 Then
Beep
    If MsgBox("�ƹ��� �Էµ� ����ڰ� �����ϴ� ��.��" & vbCrLf & "���� ��� �Ͻðڽ��ϱ�?", vbYesNo + vbQuestion, "�̷� ����ڰ� �Ѹ� ���׿� OTL") = vbYes Then
    Close #1
    LogInOK = 2
    Unload Me
    frmPersnal.Show
    Exit Sub
    Else
    Close #1
    Exit Sub
    End If
End If



LineNum = LOF(1) / Len(sdata(1))
ReDim sdata(1 To LineNum)

For i = 1 To LineNum
Get #1, i, sdata(i)
If LCase(txtID.Text) = DCLogin(sdata(i).ID) Then
    If txtPass.Text = DCLogin(sdata(i).PW) Then
    UserInfo = DCLogin(sdata(i).Name)
    UserID = DCLogin(sdata(i).ID)
    MDIForm1.Caption = MyInfo & UserInfo & "(" & UserID & ")"
    Close #1
    LogInOK = 1
    Unload Me
    frmLogin2.Show 1
    Exit Sub
    Else
    Close #1
    MsgBox "��й�ȣ�� Ʋ�Ƚ��ϴ�.", , "���� ��ŷ �� �͵� ����ϴ� ��_��"
    Exit Sub
    End If
End If
Next
Beep
If MsgBox("ID�� ã�� �� �����ϴ�." & vbCrLf & "Guest ���� �Ͻðڽ��ϱ�?", vbYesNo + vbQuestion, "���̵� �����ϴ�~") = vbYes Then
UserID = "Guest"
UserInfo = "�湮��"
MDIForm1.Caption = MyInfo & UserInfo & "(" & UserID & ")"
LogInOK = 1
Unload Me
End If

Close #1
End Sub


'///Ȯ�� ��ư�� ���콺 ���� ȿ��
Private Sub lblAcc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblAcc.BackStyle = 1
lblAcc.BackColor = &HE0E0E0
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblAcc.BackStyle = 0
End Sub
'///���콺 ���� ȿ�� ��

'///����Ű �̺�Ʈ (����� ���Ǹ� ���ؼ� ��.,��;;)
Private Sub txtID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtPass.SetFocus
txtPass.SelStart = Len(txtPass.Text)
End If
End Sub
Private Sub txtPass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
lblAcc_Click
End If
End Sub
'///�� -_-)b
