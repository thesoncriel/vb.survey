VERSION 5.00
Begin VB.Form frmBaseballHelp 
   BorderStyle     =   1  '���� ����
   Caption         =   "Form1"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   7035
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      Height          =   1095
      Left            =   4680
      MultiLine       =   -1  'True
      ScrollBars      =   2  '����
      TabIndex        =   12
      Text            =   "frmBseballHelp.frx":0000
      Top             =   3480
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFC0C0&
      Enabled         =   0   'False
      Height          =   735
      Left            =   6120
      TabIndex        =   10
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFC0C0&
      Enabled         =   0   'False
      Height          =   735
      Left            =   5400
      TabIndex        =   9
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFC0C0&
      Enabled         =   0   'False
      Height          =   735
      Left            =   4680
      TabIndex        =   8
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "?"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4680
      TabIndex        =   6
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "?"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5400
      TabIndex        =   5
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "?"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6120
      TabIndex        =   4
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Fight!"
      Height          =   735
      Left            =   6120
      TabIndex        =   2
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Wating for Fight"
      Height          =   735
      Left            =   6120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   855
      Left            =   120
      TabIndex        =   15
      Top             =   4800
      Width           =   3735
   End
   Begin VB.Label Label6 
      Caption         =   "������ ���� �߱�����"
      Height          =   735
      Left            =   120
      TabIndex        =   14
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label5 
      Alignment       =   1  '������ ����
      Height          =   975
      Left            =   120
      TabIndex        =   13
      Top             =   3600
      Width           =   4455
   End
   Begin VB.Label Label4 
      Alignment       =   1  '������ ����
      Caption         =   "�Ķ����� �� �ڽ��ȿ� ����Ǵ� 3���� ���ڸ� �����ð� 'Enter""Ű�� �����ø�"
      Height          =   735
      Left            =   2040
      TabIndex        =   11
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label Label3 
      Alignment       =   1  '������ ����
      Caption         =   "����ǥ �ڽ��� 3���� ������ ���ڰ� �Ҵ�˴ϴ�."
      Height          =   615
      Left            =   2040
      TabIndex        =   7
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   1  '������ ����
      Caption         =   "���۵Ǹ� ��ư�� ���ڰ� 'Fight!""�� �ٲ�� "
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   1080
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   1  '������ ����
      Caption         =   "������ ���� ��ư�� ������ �����մϴ�."
      Height          =   615
      Left            =   2280
      TabIndex        =   1
      Top             =   240
      Width           =   3735
   End
   Begin VB.Image Image1 
      Appearance      =   0  '���
      BorderStyle     =   1  '���� ����
      Height          =   600
      Left            =   120
      Picture         =   "frmBseballHelp.frx":0037
      Stretch         =   -1  'True
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmBaseballHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Label5.Caption = "�����ʰ� ���� ��� ȭ�鿡 ���� ��Ȳ�� ��µ˴ϴ�." & vbCrLf & "'��Ʈ����ũ'�� ��ȣ�� ��ġ�� ��ġ�Ѵٴ� ��." & vbCrLf & "'��'�� ��ȣ�� ��ġ�ϵ�, ��ġ�� ��ġ���� ������," & vbCrLf & "Out�� �� �� ���� ���� ��츦 ���մϴ�."
Label6.Caption = "������ ���� �߱�����" & vbCrLf & "����: ������"
Label7.Caption = " - ����Ű -" & vbCrLf & "��, �� : ����, �Ǵ� ������ĭ ���� �̵�." & vbCrLf & "Enter : ����, �Ǵ� ���� ��ġ �Է�"
End Sub

