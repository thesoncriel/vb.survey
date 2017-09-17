VERSION 5.00
Begin VB.Form frmPersnal 
   Appearance      =   0  '평면
   BackColor       =   &H80000005&
   Caption         =   "개인정보 입력란"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   5250
   StartUpPosition =   3  'Windows 기본값
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   3000
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblPersnal 
      Caption         =   "ID"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblPersnal 
      Caption         =   "비밀번호"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblPersnal 
      Caption         =   "ID"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblPersnal 
      Caption         =   "이름"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmPersnal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
