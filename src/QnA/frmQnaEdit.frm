VERSION 5.00
Begin VB.Form frmQnaEdit 
   BorderStyle     =   1  '단일 고정
   Caption         =   "QnA Edit"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8490
   Icon            =   "frmQnaEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   8490
   Begin VB.CommandButton btnSave 
      Caption         =   "새 파일"
      Height          =   375
      Left            =   600
      TabIndex        =   25
      Top             =   3480
      Width           =   1215
   End
   Begin VB.ComboBox cmbFileType 
      Height          =   300
      ItemData        =   "frmQnaEdit.frx":038A
      Left            =   6120
      List            =   "frmQnaEdit.frx":0391
      TabIndex        =   23
      Text            =   "*.gdb - General DB"
      Top             =   3600
      Width           =   2295
   End
   Begin VB.TextBox txtFileName 
      Height          =   270
      Left            =   3000
      TabIndex        =   22
      Top             =   3600
      Width           =   3015
   End
   Begin VB.FileListBox File1 
      Height          =   1530
      Left            =   6600
      TabIndex        =   21
      Top             =   1920
      Width           =   1815
   End
   Begin VB.OptionButton Rad 
      Caption         =   "수동(주관식)"
      Height          =   255
      Index           =   2
      Left            =   5040
      TabIndex        =   20
      Top             =   2400
      Width           =   1455
   End
   Begin VB.OptionButton Rad 
      Caption         =   "체크 버튼"
      Height          =   255
      Index           =   1
      Left            =   5040
      TabIndex        =   19
      Top             =   2160
      Width           =   1455
   End
   Begin VB.OptionButton Rad 
      Caption         =   "라디오 버튼"
      Height          =   255
      Index           =   0
      Left            =   5040
      TabIndex        =   18
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox txtHidden 
      Height          =   270
      Left            =   0
      TabIndex        =   17
      Top             =   840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton btnInitial 
      Caption         =   "텍스트 초기화"
      Height          =   615
      Left            =   4080
      TabIndex        =   10
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton btnFix 
      Caption         =   "고치기"
      Height          =   615
      Left            =   4800
      TabIndex        =   9
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton btnOpen 
      Caption         =   "열기"
      Enabled         =   0   'False
      Height          =   615
      Left            =   5520
      TabIndex        =   8
      Top             =   2760
      Width           =   975
   End
   Begin VB.ListBox lstTitle 
      Height          =   1680
      Left            =   5280
      TabIndex        =   7
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox txtTitle 
      Height          =   270
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.CommandButton btnCreate 
      Caption         =   " 새로  만들기"
      Enabled         =   0   'False
      Height          =   615
      Left            =   4080
      TabIndex        =   6
      Top             =   2760
      Width           =   735
   End
   Begin VB.TextBox txtSel 
      Height          =   270
      Index           =   3
      Left            =   720
      TabIndex        =   5
      Top             =   3000
      Width           =   3255
   End
   Begin VB.TextBox txtSel 
      Height          =   270
      Index           =   2
      Left            =   720
      TabIndex        =   4
      Top             =   2640
      Width           =   3255
   End
   Begin VB.TextBox txtSel 
      Height          =   270
      Index           =   1
      Left            =   720
      TabIndex        =   3
      Top             =   2280
      Width           =   3255
   End
   Begin VB.TextBox txtSel 
      Height          =   270
      Index           =   0
      Left            =   720
      TabIndex        =   2
      Top             =   1920
      Width           =   3255
   End
   Begin VB.TextBox txtDesc 
      Height          =   1335
      Left            =   720
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   4455
   End
   Begin VB.Label lblname 
      Caption         =   "File Name:"
      Height          =   255
      Left            =   1920
      TabIndex        =   24
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label lblEdit 
      Caption         =   "선택3"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   16
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label lblEdit 
      Caption         =   "선택2"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   15
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label lblEdit 
      Caption         =   "선택1"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label lblEdit 
      Caption         =   "선택0"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   13
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label lblEdit 
      Caption         =   "설명"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   480
      Width           =   495
   End
   Begin VB.Label lblEdit 
      Caption         =   "제목"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmQnaEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnCreate_Click()
Dim qData As QnaGeneral
Dim LineNum As Byte, i As Byte

If txtTitle.Text = "" Then
MsgBox "제목을 적지 않았습니다 : 필수~!", vbCritical, "오류! ^^;;"
Exit Sub
End If



Open FilePath & FileName0 For Random As #1 Len = Len(qData)
LineNum = LOF(1) / Len(qData)

'///같은 제목(Title)이 있는지 검색
For i = 1 To LineNum
Get #1, i, qData
txtHidden.Text = qData.Title
txtHidden.Text = RTrim(txtHidden.Text)
If txtTitle.Text = txtHidden.Text Then
MsgBox "이미 같은 제목의 데이터가 있습니다.", vbCritical, "제목(Title)은 key Data 입니다."
Close #1
Exit Sub
End If
Next
'///제목 중목 검색 끝

With qData
.Title = txtTitle.Text
For i = 0 To 3
.Sel(i) = txtSel(i).Text
Next
.Desc = txtDesc.Text
.Linef = vbCrLf
For i = 0 To 2
If Rad(i).Value Then .qType = i
Next
End With

Put #1, (LineNum + 1), qData
Close #1
lstTitle.AddItem (LineNum + 1) & " : " & qData.Title
txtTitle.Text = ""
txtDesc.Text = ""
For i = 0 To 3
txtSel(i).Text = ""
Next
End Sub

Private Sub btnFix_Click()
Dim qData As QnaGeneral
Dim RecNum As Byte, i As Byte

If txtTitle.Text = "" Then
MsgBox "제목을 적지 않았습니다 : 필수~!", vbCritical, "오류! ^^;;"
Exit Sub
End If

With qData
.Title = txtTitle.Text
For i = 0 To 3
.Sel(i) = txtSel(i).Text
Next
.Desc = txtDesc.Text
.Linef = vbCrLf
For i = 0 To 2
If Rad(i).Value Then .qType = i
Next
End With



Open FilePath & FileName0 For Random As #1 Len = Len(qData)
RecNum = lstTitle.ListIndex + 1
If RecNum = 0 Then
RecNum = LOF(1) / Len(qData)
End If
Put #1, RecNum, qData
Close #1
lstTitle.RemoveItem RecNum - 1 '리스트박스의 Index는 0부터 시작하므로
lstTitle.AddItem RecNum & " : " & qData.Title, RecNum - 1
End Sub

Private Sub btnInitial_Click()
txtTitle.Text = ""
txtDesc.Text = ""
txtSel(0).Text = ""
txtSel(1).Text = ""
txtSel(2).Text = ""
txtSel(3).Text = ""
End Sub

Private Sub btnOpen_Click()
Dim qData As QnaGeneral
Dim LineNum As Byte, i As Byte
Dim File_Path As String
btnCreate.Enabled = True
i = File1.ListIndex
File_Path = App.Path & "\Data\" & File1.List(i)
lstTitle.Clear
Open File_Path For Random As #1 Len = Len(qData)
LineNum = LOF(1) / Len(qData)


For i = 1 To LineNum
Get #1, i, qData
lstTitle.AddItem i & " : " & qData.Title
Next

With qData
txtTitle.Text = .Title
txtDesc.Text = .Desc
For i = 0 To 3
txtSel(i).Text = .Sel(i)
Next
End With
txtTitle.Text = RTrim(txtTitle.Text)
txtDesc.Text = RTrim(txtDesc.Text)
txtSel(0).Text = RTrim(txtSel(0).Text)
txtSel(1).Text = RTrim(txtSel(1).Text)
txtSel(2).Text = RTrim(txtSel(2).Text)
txtSel(3).Text = RTrim(txtSel(3).Text)
Close #1

Select Case qData.qType
Case 0
Rad(0).Value = True
Case 1
Rad(1).Value = True
Case 2
Rad(2).Value = True
End Select
End Sub

Private Sub btnSave_Click()
Dim haha As Boolean
haha = txtFileName.Text Like Left(cmbFileType.Text, 5)
If Not haha Then
txtFileName.Text = txtFileName.Text & Mid(cmbFileType.Text, 2, 4)
End If
FileName0 = txtFileName.Text
Open FilePath & FileName0 For Random As #1
Close #1
File1.Refresh
End Sub

Private Sub File1_Click()
FileName0 = File1.List(File1.ListIndex)
btnOpen.Enabled = True
End Sub

Private Sub Form_Load()
File1.Path = App.Path & "\data\"
File1.Pattern = Left(cmbFileType, 5)
End Sub

Private Sub lstTitle_Click()
Dim RecNum As Byte, i As Byte
Dim qData As QnaGeneral

RecNum = lstTitle.ListIndex + 1
Open FilePath & FileName0 For Random As #1 Len = Len(qData)
Get #1, RecNum, qData

With qData
txtTitle.Text = .Title
txtDesc.Text = .Desc
For i = 0 To 3
txtSel(i).Text = .Sel(i)
Next
End With
txtTitle.Text = RTrim(txtTitle.Text)
txtDesc.Text = RTrim(txtDesc.Text)
txtSel(0).Text = RTrim(txtSel(0).Text)
txtSel(1).Text = RTrim(txtSel(1).Text)
txtSel(2).Text = RTrim(txtSel(2).Text)
txtSel(3).Text = RTrim(txtSel(3).Text)
Close #1

Select Case qData.qType
Case 0
Rad(0).Value = True
Case 1
Rad(1).Value = True
Case 2
Rad(2).Value = True
End Select
End Sub

Private Sub Rad_Click(Index As Integer)
If Index = 2 Then
For i = 0 To 3
txtSel(i).Enabled = False
txtSel(i).BackColor = &HC0C0C0
txtSel(i).Text = ""
Next
Else
For i = 0 To 3
txtSel(i).Enabled = True
txtSel(i).BackColor = &HFFFFFF
Next

End If
End Sub
