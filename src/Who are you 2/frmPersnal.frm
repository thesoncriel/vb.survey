VERSION 5.00
Begin VB.Form frmPersnal 
   Appearance      =   0  '���
   BackColor       =   &H00A3B7D4&
   BorderStyle     =   1  '���� ����
   Caption         =   "�������� �Է¶�"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7500
   Icon            =   "frmPersnal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   MousePointer    =   4  '������
   ScaleHeight     =   4815
   ScaleWidth      =   7500
   Begin VB.TextBox Hidden 
      Height          =   270
      Left            =   240
      TabIndex        =   21
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox lstBloodType 
      Appearance      =   0  '���
      Height          =   1470
      ItemData        =   "frmPersnal.frx":038A
      Left            =   1080
      List            =   "frmPersnal.frx":03A6
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtBloodType 
      Appearance      =   0  '���
      BackColor       =   &H00D4E4F4&
      Height          =   270
      Left            =   960
      TabIndex        =   6
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox txtSex 
      Appearance      =   0  '���
      BackColor       =   &H00D4E4F4&
      Height          =   270
      Left            =   960
      MaxLength       =   1
      TabIndex        =   5
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox txtCivilCodeR 
      Appearance      =   0  '���
      BackColor       =   &H00D4E4F4&
      Height          =   270
      IMEMode         =   3  '��� ����
      Left            =   1800
      MaxLength       =   7
      PasswordChar    =   "*"
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox txtAddress 
      Appearance      =   0  '���
      BackColor       =   &H00D4E4F4&
      Height          =   270
      Left            =   960
      TabIndex        =   8
      Top             =   3240
      Width           =   3135
   End
   Begin VB.TextBox txtCivilCodeL 
      Appearance      =   0  '���
      BackColor       =   &H00D4E4F4&
      Height          =   270
      Left            =   960
      MaxLength       =   6
      TabIndex        =   3
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox txtPW 
      Appearance      =   0  '���
      BackColor       =   &H00D4E4F4&
      Height          =   270
      IMEMode         =   3  '��� ����
      Left            =   960
      MaxLength       =   16
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox txtID 
      Appearance      =   0  '���
      BackColor       =   &H00D4E4F4&
      Height          =   270
      Left            =   960
      MaxLength       =   12
      TabIndex        =   1
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  '���
      BackColor       =   &H00D4E4F4&
      Height          =   270
      Left            =   960
      MaxLength       =   15
      TabIndex        =   0
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label lblDesc 
      BackStyle       =   0  '����
      Height          =   1575
      Left            =   4680
      TabIndex        =   20
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00404080&
      X1              =   240
      X2              =   2760
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line4 
      X1              =   360
      X2              =   360
      Y1              =   3960
      Y2              =   4560
   End
   Begin VB.Label lblComplete 
      Alignment       =   2  '��� ����
      Appearance      =   0  '���
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "Input Complete"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080A0&
      Height          =   735
      Left            =   0
      TabIndex        =   19
      Top             =   4080
      Width           =   3015
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00008080&
      BorderStyle     =   5  '���-��-��
      X1              =   240
      X2              =   4680
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Label lblPersnal 
      Alignment       =   1  '������ ����
      Appearance      =   0  '���
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "�ּ�"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   0
      TabIndex        =   18
      Top             =   3290
      Width           =   855
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000080&
      BorderStyle     =   3  '��
      X1              =   4320
      X2              =   4320
      Y1              =   840
      Y2              =   4680
   End
   Begin VB.Line Line1 
      X1              =   7320
      X2              =   240
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label lblPersnal 
      Alignment       =   1  '������ ����
      Appearance      =   0  '���
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "������"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   0
      TabIndex        =   17
      Top             =   2925
      Width           =   855
   End
   Begin VB.Label Label2 
      Appearance      =   0  '���
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      Caption         =   "-"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1620
      TabIndex        =   16
      Top             =   2205
      Width           =   135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '����
      Caption         =   "Persnal Information"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   3600
      TabIndex        =   15
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '����
      Caption         =   "Persnal Information"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080A0&
      Height          =   615
      Index           =   0
      Left            =   3705
      TabIndex        =   14
      Top             =   315
      Width           =   3615
   End
   Begin VB.Label lblPersnal 
      Alignment       =   1  '������ ����
      BackStyle       =   0  '����
      Caption         =   "����"
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   13
      Top             =   2560
      Width           =   855
   End
   Begin VB.Label lblPersnal 
      Alignment       =   1  '������ ����
      BackStyle       =   0  '����
      Caption         =   "�ֹι�ȣ"
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   12
      Top             =   2205
      Width           =   855
   End
   Begin VB.Label lblPersnal 
      Alignment       =   1  '������ ����
      BackStyle       =   0  '����
      Caption         =   "��й�ȣ"
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   11
      Top             =   1845
      Width           =   855
   End
   Begin VB.Label lblPersnal 
      Alignment       =   1  '������ ����
      BackStyle       =   0  '����
      Caption         =   "ID"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   10
      Top             =   1500
      Width           =   855
   End
   Begin VB.Label lblPersnal 
      Alignment       =   1  '������ ����
      BackStyle       =   0  '����
      Caption         =   "�̸�"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   9
      Top             =   1125
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   2445
      Left            =   4200
      Picture         =   "frmPersnal.frx":03E4
      Top             =   2400
      Width           =   3345
   End
End
Attribute VB_Name = "frmPersnal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
If EditMode = 1 Then
Dim sdata() As PersnalData
Dim LineNum As Byte, i As Byte

Open FilePath & PersnalDataFile For Random As #1 Len = Len(sdata(1))
LineNum = LOF(1) / Len(sdata(1))

ReDim sdata(1 To LineNum)

For i = 1 To LineNum
Get #1, i, sdata(i)
'///Hidden.Text�� �� ������ �ѱ� �������� ���鸦 ���ֱ� ���� �� �ϳ��� ���� �Դϴ�.
'///�α��� ���� �����ؼ� �������� DB ������ �ҷ����� �� ���� ������(?)�ϰ� ���Դϴ�.
'///���� �����ʹ� �̷� ���� ���� RTrim�� �ٷ� ����˴ϴ�. (�� ��� ���ٴ� -_-;;)
'///HIdden.Text�� DC~ �ø����� ����� �Լ��� ���Դϴ�.
    If UserID = DCPersnal(sdata(i).ID) Then
    '///�� �ؽ�Ʈ �ڽ��� �����͸� �Է�
    With sdata(i)
    txtName.Text = DCPersnal(.Name)
    txtID.Text = UserID
    txtPW.Text = DCPersnal(.PW)
    txtCivilCodeL.Text = RTrim(Left(.Civilcode, 6))
    txtCivilCodeR.Text = RTrim(Right(.Civilcode, 7))
    txtSex.Text = .Sex
    txtBloodType.Text = RTrim(.BloodType)
    txtAddress.Text = DCPersnal(.Address)
    End With
    Close #1
'///frmPersnal.Tag�� ���� ����� ���� ��������� ���,
'///������� ������ ��ġ�� ����ϴ� ������ �þҽ��ϴ�. ^^
'///������ �� ���� ������....
'///�׳� -_-;;;;
    frmPersnal.Tag = i
    Exit Sub
    End If
Next

End If
Close #1
End Sub

'///�Է¶��� ���콺 ���� ȿ��
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Const TxtBackColor As Long = &HD4E4F4
lblComplete.ForeColor = &H8080A0
lblDesc.Caption = ""
txtName.BackColor = TxtBackColor
txtID.BackColor = TxtBackColor
txtPW.BackColor = TxtBackColor
txtCivilCodeL.BackColor = TxtBackColor
txtCivilCodeR.BackColor = TxtBackColor
txtSex.BackColor = TxtBackColor
txtBloodType.BackColor = TxtBackColor
txtAddress.BackColor = TxtBackColor
End Sub
Private Sub lblComplete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblDesc.Caption = ""
lblComplete.ForeColor = &H10101
End Sub
Private Sub txtName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblDesc.Caption = ""
txtName.BackColor = &HFFFFFF
End Sub
Private Sub txtID_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblDesc.Caption = ""
txtID.BackColor = &HFFFFFF
End Sub
Private Sub txtPW_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblDesc.Caption = ""
txtPW.BackColor = &HFFFFFF
End Sub
Private Sub txtCivilCodeL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblDesc.Caption = ""
txtCivilCodeL.BackColor = &HFFFFFF
txtCivilCodeR.BackColor = &HFFFFFF
End Sub
Private Sub txtCivilCodeR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblDesc.Caption = ""
txtCivilCodeL.BackColor = &HFFFFFF
txtCivilCodeR.BackColor = &HFFFFFF
End Sub
Private Sub txtSex_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblDesc.Caption = ""
txtSex.BackColor = &HFFFFFF
End Sub
Private Sub txtBloodType_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblDesc.Caption = ""
txtBloodType.BackColor = &HFFFFFF
End Sub
Private Sub txtAddress_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblDesc.Caption = ""
txtAddress.BackColor = &HFFFFFF
End Sub
'///�Է¶� ���콺 ���� ȿ�� ��




'///������ �Է¶��� �޺��ڽ� ��������
Private Sub lstBloodType_Click()
txtBloodType.Text = lstBloodType.Text
End Sub
Private Sub lstBloodType_DblClick()
lstBloodType.Visible = False
txtAddress.SetFocus
End Sub
Private Sub lstBloodType_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtAddress.SetFocus
End Sub
Private Sub lstBloodType_LostFocus()
lstBloodType.Visible = False
End Sub
Private Sub txtBloodType_GotFocus()
lstBloodType.Visible = True
lstBloodType.SetFocus
End Sub
'///������ �Է¶� �޺��ڽ� ��



'///�ֹι�ȣ �Է¶� ȿ��
Private Sub txtCivilCodeL_Change()
If EditMode = 0 Then
If Len(txtCivilCodeL.Text) = 6 Then txtCivilCodeR.SetFocus
End If
End Sub
Private Sub txtCivilCodeR_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 47 To 57
Exit Sub
Case 8
If txtCivilCodeR.Text = "" Then
txtCivilCodeL.SetFocus
txtCivilCodeL.SelStart = Len(txtCivilCodeL.Text)
End If
Exit Sub
Case Else
Beep
End Select
End Sub
'///�ֹι�ȣ �Է¶� ȿ�� ��




Private Sub txtCivilCodeR_LostFocus()
Dim SexNum As Byte
If Len(txtCivilCodeR.Text) = 7 And IsNumeric(txtCivilCodeR.Text) Then
SexNum = Left(txtCivilCodeR.Text, 1)
Select Case SexNum
Case 1, 3
txtSex.Text = "��"
Case 2, 4
txtSex.Text = "��"
End Select
txtSex.Enabled = False
End If

End Sub

Private Sub txtSex_Click()
Dim SexWW As String
SexWW = txtSex.Text
Select Case SexWW
Case "", "��"
txtSex.Text = "��"
Case "��"
txtSex.Text = "��"
End Select
End Sub

Private Sub lblComplete_Click()
Complete
End Sub



'///�Է¶��鿡 ���� �������� ����
Private Sub Complete()
Dim MsgResult As Integer, Civilcode As String
Dim PDCheck As PersnalData, RepeatID As String
Dim LineNum As Integer, FileSize As Long

Civilcode = txtCivilCodeL.Text & txtCivilCodeR.Text


If txtName.Text = "" Then
MsgBox "�̸��� �����ּ���.", , "�̸� ����"
Exit Sub
ElseIf txtID.Text = "" Then
MsgBox "���̵� �����ּ���.", , "���̵� ����#1"
Exit Sub
ElseIf Trim(LCase(txtID.Text)) = "guest" Then
MsgBox "Guest�� ���̵� �� �� �����ϴ�.", , "���̵� ����#2 : �峭�Ͻô°� ����? -_-;;"
Exit Sub
ElseIf txtPW.Text = "" Then
MsgBox "��й�ȣ�� �����ּ���.", , "��й�ȣ ����"
Exit Sub
ElseIf 0 < Len(Civilcode) And Len(Civilcode) < 13 Then
MsgBox "�ֹι�ȣ�� �ٽ� �����ּ���.", , "�ֹι�ȣ ����#1"
Exit Sub
ElseIf IsNumeric(Civilcode) = False Then
    If Len(Civilcode) > 0 Then
    MsgBox "�ֹι�ȣ�� �ٽ� �����ּ���.", , "�ֹι�ȣ ����#2"
    Exit Sub
    End If
End If

If EditMode = 0 Then
Open PersnalDataFile For Random As #1 Len = Len(PDCheck)
LineNum = LOF(1) / Len(PDCheck)
For i = 1 To LineNum
Get #1, i, PDCheck
If txtID.Text = RTrim(PDCheck.ID) Then
MsgBox "�̹� ���� ID�� ���� �մϴ�.", , "���̵� ����#3"
Close #1
Exit Sub
ElseIf Civilcode = PDCheck.Civilcode Then
MsgBox "�̹� ���� �ֹι�ȣ�� ���� �մϴ�.", , "�ֹι�ȣ ����#3"
Close #1
Exit Sub
End If
Next

Close #1
End If
frmPersnalResult.Show 1
End Sub
'///�Է¶� ���� ���� ��
