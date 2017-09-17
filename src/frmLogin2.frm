VERSION 5.00
Begin VB.Form frmLogin2 
   BorderStyle     =   0  '없음
   ClientHeight    =   4935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin2.frx":0000
   ScaleHeight     =   4935
   ScaleWidth      =   3870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.Timer Timer1 
      Interval        =   600
      Left            =   1560
      Top             =   3120
   End
   Begin VB.Label lblMyinfo 
      BackStyle       =   0  '투명
      Caption         =   "만든이 : 손준현"
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Label lblinfo 
      BackStyle       =   0  '투명
      Caption         =   "Name"
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   4
      Top             =   2640
      Width           =   2895
   End
   Begin VB.Label lblinfo 
      BackStyle       =   0  '투명
      Caption         =   "ID"
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   3
      Top             =   2400
      Width           =   2895
   End
   Begin VB.Label lblWelcome 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "Thank you for use"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Label lblWelcome 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "Thank you for use"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A0A0A0&
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   1760
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "::::Now LogIn::::"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   3855
   End
End
Attribute VB_Name = "frmLogin2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Private Sub Form_Load()
lblinfo(0).Caption = "Your ID: " & UserID
lblinfo(1).Caption = "Your Name: " & UserInfo
End Sub
Private Sub Timer1_Timer()
i = i + 1
If i Mod 2 = 0 Then Unload Me
End Sub
