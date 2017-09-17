VERSION 5.00
Begin VB.Form frmBaseball 
   BorderStyle     =   1  '단일 고정
   Caption         =   "숫자 야구 게임"
   ClientHeight    =   3225
   ClientLeft      =   1935
   ClientTop       =   1995
   ClientWidth     =   5475
   Icon            =   "frmBaseball.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   5475
   Begin VB.TextBox txtResult 
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      Height          =   3015
      Left            =   3240
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton btnUserProg 
      Caption         =   "Wating for Fight"
      Height          =   735
      Left            =   2400
      TabIndex        =   3
      Tag             =   "off"
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox txtUseNum3 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   27.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   710
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   2
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox txtUseNum2 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   27.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   710
      Left            =   840
      MaxLength       =   1
      TabIndex        =   1
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox txtUseNum1 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   27.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   710
      Left            =   120
      MaxLength       =   1
      TabIndex        =   0
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton btnComNum 
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
      Index           =   3
      Left            =   1560
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton btnComNum 
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
      Index           =   2
      Left            =   840
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton btnComNum 
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
      Index           =   1
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   735
   End
   Begin VB.Menu NewGame 
      Caption         =   "새로운 게임"
      Index           =   1
      NegotiatePosition=   1  '왼쪽
   End
   Begin VB.Menu Help 
      Caption         =   "도움말"
      Index           =   2
      NegotiatePosition=   1  '왼쪽
   End
End
Attribute VB_Name = "frmBaseball"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ComNum(1 To 3) As Byte

Private Sub btnUserProg_Click()
Dim UseNum1 As Byte, UseNum2 As Byte, UseNum3 As Byte
Dim Strike As Byte, Ball As Byte, i As Byte
Static Cnum As Integer
Cnum = Cnum + 1

If Cnum = 1 Then
ReDim nRndNum(1 To 3)
For i = 1 To 3
nRndNum(i) = RndNumProg(0, 9, 3)
Next
End If

For i = 1 To 3
ComNum(i) = nRndNum(i)
Next

UseNum1 = Val(txtUseNum1.Text)
UseNum2 = Val(txtUseNum2.Text)
UseNum3 = Val(txtUseNum3.Text)

If btnUserProg.Tag = "off" Then
btnUserProg.Tag = "on"
btnUserProg.Caption = "Attack!"
txtUseNum1.SetFocus
Exit Sub
ElseIf (UseNum1 = UseNum2) Or (UseNum2 = UseNum3) Or (UseNum3 = UseNum1) Then
MsgBox "당신의 숫자가 2개이상 같거나 빈칸입니다."
Exit Sub
End If

If ComNum(1) = UseNum1 Then
Strike = Strike + 1
ElseIf ComNum(1) = UseNum2 Then
Ball = Ball + 1
ElseIf ComNum(1) = UseNum3 Then
Ball = Ball + 1
End If

If ComNum(2) = UseNum1 Then
Ball = Ball + 1
ElseIf ComNum(2) = UseNum2 Then
Strike = Strike + 1
ElseIf ComNum(2) = UseNum3 Then
Ball = Ball + 1
End If

If ComNum(3) = UseNum1 Then
Ball = Ball + 1
ElseIf ComNum(3) = UseNum2 Then
Ball = Ball + 1
ElseIf ComNum(3) = UseNum3 Then
Strike = Strike + 1
End If

With txtResult
If (Strike = 0) And (Ball = 0) Then
.Text = .Text & "Out~!   -_-)p" & vbCrLf
ElseIf Strike < 3 Then
.Text = .Text & Strike & "스트라이크, " & Ball & "볼~!" & vbCrLf
Else
.Text = .Text & "H.O.M.E - R.U.N - ! !" & vbCrLf
btnComNum(1).Caption = ComNum(1)
btnComNum(2).Caption = ComNum(2)
btnComNum(3).Caption = ComNum(3)

MsgBox "홈런입니다~! (~^^)~" & vbCrLf & "적의 숫자는 " & ComNum(1) & ":" & ComNum(2) & ":" & ComNum(3) & " 이었습니다.", , "Congratulation~!"
Beep
    If MsgBox("또 하시겠습니까?", vbQuestion + vbYesNo, "Replay?") = vbYes Then
    For i = 1 To 3
    ComNum(i) = RndNumProg(0, 9, 3)
    Next
    btnComNum(1).Caption = "?"
    btnComNum(2).Caption = "?"
    btnComNum(3).Caption = "?"
    txtUseNum1.SetFocus
    Else
    btnUserProg.Tag = "off"
    btnUserProg.Caption = "Wating for Fight"
    End If
txtUseNum1.Text = ""
txtUseNum2.Text = ""
txtUseNum3.Text = ""
txtResult.Text = ""
Cnum = 0
End If
End With

End Sub

Private Sub Form_Load()
Me.Caption = Me.Caption & " - 사용자: " & UserInfo & "(" & UserID & ")"
End Sub

Private Sub Help_Click(Index As Integer)
frmBaseballHelp.Show 1
End Sub

Private Sub NewGame_Click(Index As Integer)
    For i = 1 To 3
    ComNum(i) = RndNumProg(0, 9, 3)
    Next
    btnComNum(1).Caption = "?"
    btnComNum(2).Caption = "?"
    btnComNum(3).Caption = "?"
    txtUseNum1.SetFocus
    btnUserProg.Tag = "off"
    btnUserProg.Caption = "Wating for Fight"
txtUseNum1.Text = ""
txtUseNum2.Text = ""
txtUseNum3.Text = ""
txtResult.Text = ""
Cnum = 0
End Sub

Private Sub txtUseNum1_Change()
With txtUseNum1
If IsNumeric(.Text) Then
txtUseNum2.SetFocus
End If
End With
End Sub

Private Sub txtUseNum2_Change()
With txtUseNum2
If IsNumeric(.Text) Then
txtUseNum3.SetFocus
End If
End With
End Sub

Private Sub txtUseNum1_GotFocus()
With txtUseNum1
.BackColor = &HFFC0C0
.SelStart = 0
.SelLength = 1
End With
txtUseNum2.BackColor = &HFFA0A0
txtUseNum3.BackColor = &HFFA0A0
End Sub

Private Sub txtUseNum2_GotFocus()
With txtUseNum2
.BackColor = &HFFC0C0
.SelStart = 0
.SelLength = 1
End With
txtUseNum1.BackColor = &HFFA0A0
txtUseNum3.BackColor = &HFFA0A0
End Sub

Private Sub txtUseNum3_GotFocus()
With txtUseNum3
.BackColor = &HFFC0C0
.SelStart = 0
.SelLength = 1
End With
txtUseNum1.BackColor = &HFFA0A0
txtUseNum2.BackColor = &HFFA0A0
End Sub

Private Sub txtUseNum1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
btnUserProg.SetFocus
Call btnUserProg_Click
End If
End Sub

Private Sub txtUseNum2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
btnUserProg.SetFocus
Call btnUserProg_Click
ElseIf KeyAscii = 8 Then
txtUseNum1.SetFocus
End If
End Sub

Private Sub txtUseNum3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
btnUserProg.SetFocus
Call btnUserProg_Click
ElseIf KeyAscii = 8 Then
txtUseNum2.SetFocus
End If
End Sub

Private Sub btnUserProg_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
txtUseNum3.SetFocus
End If
End Sub

Private Sub txtUseNum1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyRight Then txtUseNum2.SetFocus
End Sub
Private Sub txtUseNum2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyLeft Then
txtUseNum1.SetFocus
ElseIf KeyCode = vbKeyRight Then
txtUseNum3.SetFocus
End If
End Sub
Private Sub txtUseNum3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyLeft Then txtUseNum2.SetFocus
End Sub
