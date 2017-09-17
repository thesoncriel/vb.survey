Attribute VB_Name = "FileDataPath"
Public FilePath As String
Public Const PersnalDataFile As String = "Persnal_Data.pdb"
Public Const ProgramSetting As String = "ProgSet.txt"

Public Const MyInfo As String = "컴퓨터 프로그래밍 02반 ::산림환경자원학 2003012500 손준현 - 사용자: "
Public LogInOK As Byte '0이면 처음시작, 1이면 로그인 성공, 2이면 로그인 안함
Public EditMode As Byte
Public UserInfo As String, UserID As String

Public Const pDataField As Byte = 8

Public QnA_Num As Integer
Public FileName0 As String, FileName1 As String
Public ResultTemp As String, FreeAnswer As String
Public FixLineNum As Byte

Public Type PersnalData
ID As String * 12
PW As String * 16
Name As String * 16
Civilcode As String * 13
Sex As String * 2
BloodType As String * 6
Address As String * 64
Linef As String * 2
End Type

Public Type QnaGeneral
Title As String * 12
qType As String * 1
Sel(0 To 3) As String * 24
Desc As String * 128
Linef As String * 2
End Type

Public Type QnaGeneralResult
'///Qn값: 0000-주관식, 1000~0001-라디오버튼, 2000~0002-체크버튼
ID As String * 12
Q(1 To 10) As String * 4
DD As String * 10
QF As String * 128
Linef As String * 2
End Type
Public Type QnaResultTemp
ID As String * 12
QN As String * 40
DD As String * 10
QF As String * 128
Linef As String * 2
End Type
