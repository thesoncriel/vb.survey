VERSION 5.00
Begin VB.Form frmBaseballHelp 
   BorderStyle     =   1  '단일 고정
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
   StartUpPosition =   3  'Windows 기본값
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      Height          =   1095
      Left            =   4680
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
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
         Name            =   "굴림"
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
         Name            =   "굴림"
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
         Name            =   "굴림"
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
      Caption         =   "왕허접 숫자 야구게임"
      Height          =   735
      Left            =   120
      TabIndex        =   14
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label5 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   975
      Left            =   120
      TabIndex        =   13
      Top             =   3600
      Width           =   4455
   End
   Begin VB.Label Label4 
      Alignment       =   1  '오른쪽 맞춤
      Caption         =   "파란색의 빈 박스안에 예상되는 3가지 숫자를 적으시고 'Enter""키를 누르시면"
      Height          =   735
      Left            =   2040
      TabIndex        =   11
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label Label3 
      Alignment       =   1  '오른쪽 맞춤
      Caption         =   "물음표 박스에 3개의 무작위 숫자가 할당됩니다."
      Height          =   615
      Left            =   2040
      TabIndex        =   7
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   1  '오른쪽 맞춤
      Caption         =   "시작되면 버튼의 글자가 'Fight!""로 바뀌며 "
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   1080
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      Caption         =   "오른쪽 같은 버튼을 누르면 시작합니다."
      Height          =   615
      Left            =   2280
      TabIndex        =   1
      Top             =   240
      Width           =   3735
   End
   Begin VB.Image Image1 
      Appearance      =   0  '평면
      BorderStyle     =   1  '단일 고정
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
Label5.Caption = "오른쪽과 같은 녹색 화면에 현재 상황이 출력됩니다." & vbCrLf & "'스트라이크'는 번호와 위치가 일치한다는 것." & vbCrLf & "'볼'은 번호는 일치하되, 위치는 일치하지 않음을," & vbCrLf & "Out은 둘 다 맞지 않을 경우를 말합니다."
Label6.Caption = "왕허접 숫자 야구게임" & vbCrLf & "제작: 손준현"
Label7.Caption = " - 단축키 -" & vbCrLf & "←, → : 왼쪽, 또는 오른쪽칸 으로 이동." & vbCrLf & "Enter : 시작, 또는 현재 수치 입력"
End Sub

