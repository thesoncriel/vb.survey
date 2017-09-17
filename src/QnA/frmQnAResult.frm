VERSION 5.00
Begin VB.Form frmQnAResult 
   Appearance      =   0  '평면
   BackColor       =   &H00D0A070&
   BorderStyle     =   1  '단일 고정
   Caption         =   "::::감사합니다::::"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4185
   FillColor       =   &H00D0A070&
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   8.25
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmQnAResult.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmQnAResult.frx":038A
   ScaleHeight     =   5145
   ScaleWidth      =   4185
   Begin VB.TextBox Hidden 
      Height          =   255
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtResult 
      Appearance      =   0  '평면
      BackColor       =   &H00CBBBAB&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2295
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   3
      Top             =   2760
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Label lblViewResult 
      BackStyle       =   0  '투명
      Caption         =   "결과 보기"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   2280
      Width           =   975
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FF8080&
      BorderStyle     =   5  '대시-점-점
      X1              =   1680
      X2              =   3240
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00606000&
      BorderStyle     =   3  '점
      X1              =   3000
      X2              =   3000
      Y1              =   840
      Y2              =   3000
   End
   Begin VB.Line Line1 
      BorderStyle     =   4  '대시-점
      X1              =   4200
      X2              =   120
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label lblDesc 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   2940
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  '투명
      Caption         =   "QnA Result"
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
      Index           =   0
      Left            =   1920
      TabIndex        =   0
      Top             =   300
      Width           =   2415
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  '투명
      Caption         =   "QnA Result"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00B09068&
      Height          =   615
      Index           =   1
      Left            =   2040
      TabIndex        =   2
      Top             =   360
      Width           =   2415
   End
End
Attribute VB_Name = "frmQnAResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const Desc As String = "설문에 응해 주셔서 감사합니다." & vbCrLf & _
vbCrLf & "늘 좋은 하루 되세요~ ^^*"

Private Sub Form_Load()
Dim ID As String * 12
Dim Temp As QnaResultTemp
Dim LineNum As Byte

ID = UserID

With Temp
.ID = ID
.QN = ResultTemp
.DD = Format(Date, "yyddmm") & Format(Time, "hhmm")
.QF = FreeAnswer
.Linef = vbCrLf
End With
Open FilePath & FileName1 For Random As #1 Len = Len(Temp)
LineNum = LOF(1) / Len(Temp) + 1

'///설문조사를 한번 더 이행하는 중이었는지 검사
If FixLineNum = 0 Then
Put #1, LineNum, Temp
ElseIf FixLineNum > 0 Then
Put #1, FixLineNum, Temp
End If
'///검사 끝


Close #1
lblDesc.Caption = Desc
QnA_Num = 0
FreeAnswer = ""
ResultTemp = ""
End Sub

Private Sub lblViewResult_Click()
Dim LineNum1 As Byte, LineNum2 As Byte, i As Byte, j As Byte
Dim Result As QnaGeneralResult
Dim QnAData As QnaGeneral

txtResult.Visible = True

Open FilePath & FileName0 For Random As #1 Len = Len(QnAData)
Open FilePath & FileName1 For Random As #2 Len = Len(Result)
LineNum2 = LOF(2) / Len(Result)
If FixLineNum = 0 Then
Get #2, LineNum2, Result
Else
Get #2, FixLineNum, Result
End If
txtResult.Text = "사용자: " & DCQnAResult(Result.ID) & vbCrLf & vbCrLf

With txtResult
For i = 1 To 10
Get #1, i, QnAData
.Text = .Text & "문항" & i & "." & vbCrLf & DCQnAResult(QnAData.Desc) & vbCrLf
    If QnAData.qType = "0" Then
        For j = 0 To 3
            If Result.Q(i) = Format(1000 / 10 ^ j, "000#") Then
            .Text = .Text & "  " & j + 1 & ".  " & DCQnAResult(QnAData.Sel(j)) & vbCrLf
            End If
        Next
    ElseIf QnAData.qType = "1" Then
        For j = 0 To 3
            If Mid(Result.Q(i), j + 1, 1) = 2 Then
            .Text = .Text & "  * " & DCQnAResult(QnAData.Sel(j)) & vbCrLf
            End If
        Next
    Else
        .Text = .Text & DCQnAResult(Result.QF) & vbCrLf
    End If
    .Text = .Text & vbCrLf
Next
End With
Close #1
Close #2
End Sub
