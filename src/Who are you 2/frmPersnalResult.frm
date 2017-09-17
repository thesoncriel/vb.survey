VERSION 5.00
Begin VB.Form frmPersnalResult 
   BackColor       =   &H00A3B7D4&
   BorderStyle     =   1  '단일 고정
   Caption         =   "개인정보 입력 완료"
   ClientHeight    =   5235
   ClientLeft      =   2520
   ClientTop       =   2280
   ClientWidth     =   4680
   Icon            =   "frmPersnalResult.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   4680
   Begin VB.Label lblNo 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080A0&
      Height          =   375
      Left            =   3840
      TabIndex        =   16
      Top             =   4620
      Width           =   615
   End
   Begin VB.Label lblYes 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080A0&
      Height          =   375
      Left            =   3120
      TabIndex        =   15
      Top             =   4620
      Width           =   615
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00004040&
      BorderStyle     =   5  '대시-점-점
      X1              =   120
      X2              =   2520
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00004080&
      X1              =   1440
      X2              =   4440
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line3 
      X1              =   2040
      X2              =   2040
      Y1              =   4440
      Y2              =   5040
   End
   Begin VB.Label lblResult 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "맞습니까?"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   1800
      TabIndex        =   14
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label lblAddress 
      BackStyle       =   0  '투명
      Caption         =   "Empty"
      Height          =   495
      Left            =   1080
      TabIndex        =   13
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label lblBloodType 
      BackStyle       =   0  '투명
      Caption         =   "Empty"
      Height          =   255
      Left            =   1080
      TabIndex        =   12
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label lblSex 
      BackStyle       =   0  '투명
      Caption         =   "Empty"
      Height          =   255
      Left            =   1080
      TabIndex        =   11
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label lblCivilCode 
      BackStyle       =   0  '투명
      Caption         =   "Empty"
      Height          =   255
      Left            =   1080
      TabIndex        =   10
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label lblID 
      BackStyle       =   0  '투명
      Caption         =   "Empty"
      Height          =   255
      Left            =   1080
      TabIndex        =   9
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label lblName 
      BackStyle       =   0  '투명
      Caption         =   "Empty"
      Height          =   255
      Left            =   1080
      TabIndex        =   8
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label lblResult 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "주소"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   7
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label lblResult 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "혈액형"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   6
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label lblResult 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "성별"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label lblResult 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "주민번호"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label lblResult 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "ID"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label lblResult 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "이름"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label lblResultTitle 
      BackStyle       =   0  '투명
      Caption         =   "Result"
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
      Left            =   3240
      TabIndex        =   0
      Top             =   960
      Width           =   1455
   End
   Begin VB.Line Line1 
      X1              =   4560
      X2              =   240
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000080&
      BorderStyle     =   3  '점
      X1              =   3000
      X2              =   3000
      Y1              =   1560
      Y2              =   3960
   End
   Begin VB.Label lblResultTilteS 
      BackStyle       =   0  '투명
      Caption         =   "Result"
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
      Left            =   3345
      TabIndex        =   1
      Top             =   1035
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   1875
      Left            =   -120
      Picture         =   "frmPersnalResult.frx":038A
      Top             =   0
      Width           =   3000
   End
   Begin VB.Image imgRealyBG 
      Height          =   420
      Left            =   1920
      Picture         =   "frmPersnalResult.frx":128C6
      Top             =   4560
      Width           =   1170
   End
End
Attribute VB_Name = "frmPersnalResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim Name As String, ID As String, Civilcode As String, Sex As String, BloodType As String
Dim Address As String
Dim PW As String

'///이전 폼인 frmPersnal의 텍스트 박스에 있던 데이터를
'///저장시킬 Type형 변수로 대입 시킵니다.
With frmPersnal
Name = .txtName.Text
ID = LCase(.txtID.Text)
PW = .txtPW.Text
Civilcode = .txtCivilCodeL.Text & "-" & .txtCivilCodeR.Text
Sex = .txtSex.Text
BloodType = .txtBloodType.Text
Address = .txtAddress.Text
End With


'///데이터를 폼의 레이블에 출력시킵니다.
'///필수 입력요소를 제외한 나머지가 데이터가 비었을 경우
'///Empty라는 값을 출력하며 글자색상을 옅게 만듭니다.
lblName.Caption = Name
lblID.Caption = ID

If Civilcode = "-" Then
Civilcode = "Empty"
lblCivilCode.ForeColor = &HA0A0C0
Else
lblCivilCode.Caption = Left(Civilcode, 6) & "-" & String(7, "*")
End If

If Sex = "" Then
lblSex.Caption = "Empty"
lblSex.ForeColor = &HA0A0C0
Else
lblSex.Caption = Sex
End If

If BloodType = "" Then
lblBloodType.Caption = "Empty"
lblBloodType.ForeColor = &HA0A0C0
Else
lblBloodType.Caption = BloodType
End If

If Address = "" Then
lblAddress.Caption = "Empty"
lblAddress.ForeColor = &HA0A0C0
Else
lblAddress.Caption = Address
End If

End Sub


Private Sub lblNo_Click()
Unload Me
End Sub

Private Sub lblYes_Click()

Dim PersnalResult As PersnalData
Dim LineNum As Integer, FileSize As Long
Dim i As Integer

With PersnalResult
.ID = frmPersnal.txtID.Text
.PW = frmPersnal.txtPW.Text
.Name = frmPersnal.txtName.Text
.Civilcode = frmPersnal.txtCivilCodeL.Text & frmPersnal.txtCivilCodeR.Text
.Sex = frmPersnal.txtSex.Text
.BloodType = frmPersnal.txtBloodType.Text
.Address = frmPersnal.txtAddress.Text
.Linef = vbCrLf
End With

Open FilePath & PersnalDataFile For Random As #1 Len = Len(PersnalResult)

LineNum = LOF(1) / Len(PersnalResult)
If EditMode = 0 Then
Put #1, (LineNum + 1), PersnalResult
Else
Put #1, Val(frmPersnal.Tag), PersnalResult
EditMode = 0
frmPersnal.Tag = ""
End If

Close #1
Unload Me
MsgBox "사용자 등록이 성공적으로 끝났습니다 ^^;;", vbInformation, "오오~ 감사합니다 ^^)r"
UserID = RTrim(PersnalResult.ID)
UserInfo = RTrim(PersnalResult.Name)
MDIForm1.Caption = MyInfo & UserInfo & "(" & UserID & ")"
Unload frmPersnal
End Sub


'///Yes, NO 버튼의 마우스 오버 효과
Private Sub lblYes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblYes.ForeColor = &H0
End Sub
Private Sub lblNo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblNo.ForeColor = &H0
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblYes.ForeColor = &H8080A0
lblNo.ForeColor = &H8080A0
End Sub
'///끝
