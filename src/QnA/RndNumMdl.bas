Attribute VB_Name = "RndNumMdl"
Public Function RndNumProg(X As Integer, Y As Integer, Z As Integer, A As Integer, P As Boolean) As Integer
Dim iRnd As Integer, jRnd As Integer
Static RndNum() As Integer, PowerNum As Integer
'X -> Min
'Y -> Max
'Z -> LoopNum
'A -> ArrayNum
'P -> Power On/Off (True/False)
PowerNum = PowerNum + 1

If P = False Then
ReDim RndNum(0 To 0)
PowerNum = 0
Exit Function
End If

If PowerNum > 1 Then
RndNumProg = RndNum(A)
Exit Function
Else
ReDim RndNum(1 To Z)
End If

Randomize
RndNum(1) = Int((Y - X + 1) * Rnd + X)

For iRnd = 2 To Z
Reprog:
    Randomize
    RndNum(iRnd) = Int((Y - X + 1) * Rnd + X)
        For jRnd = 1 To (iRnd - 1)
        If RndNum(iRnd) = RndNum(jRnd) Then
        GoTo Reprog
        End If
    Next
Next

RndNumProg = RndNum(A)

End Function
