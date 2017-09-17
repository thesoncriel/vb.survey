Attribute VB_Name = "RndNumMdl"
Public nRndNum() As Integer

Public Function RndNumProg(X As Integer, Y As Integer, Z As Integer) As Integer
Dim iRnd As Integer, jRnd As Integer
Static RndNum() As Integer, PowerNum As Integer
'X -> Min
'Y -> Max
'Z -> LoopNum
PowerNum = PowerNum + 1

If PowerNum = 1 Then
    ReDim RndNum(1 To Z)
ElseIf PowerNum = Z Then
    RndNumProg = RndNum(PowerNum)
    ReDim RndNum(0 To 1)
    PowerNum = 0
    Exit Function
ElseIf PowerNum > 1 Then
    RndNumProg = RndNum(PowerNum)
    Exit Function
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

RndNumProg = RndNum(1)

End Function
