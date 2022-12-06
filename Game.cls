VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim isGameRunning As Boolean

Private Sub CommandButton1_Click()
    If isGameRunning Then isGameRunning = False
    isGameRunning = True
    Call StartRound
End Sub

Sub StartRound()
    Call Worksheet_Activate
    Call SetScoreBoard
    
    [ballx] = 0: [bally] = 0
    [balldirx] = Array(1, -1)(Int(WorksheetFunction.RandBetween(0, 1))): [balldiry] = Array(1, -1)(Int(WorksheetFunction.RandBetween(0, 1)))
    Call ResetPaddlePos
    Call SetScoreBoard
    t = Timer
    
    Do While isGameRunning
        If Timer - t >= 0.001 Then
            Call MoveBall
            Call ProcessInput
            Call CheckBallPosition
            DoEvents
        End If
    Loop
End Sub

Sub ResetPaddlePos()
    [k2] = -40:   [L2] = -20
    [N2] = 40:   [O2] = -20
End Sub


Sub CheckBallPosition()
    If [ballx] < 0 Then
        If [ballx] < [K4] And Not ([bally] >= [L5] And [bally] <= [L4]) Then
            isGameRunning = False
            [player2Score] = [player2Score] + 1
            Call SetScoreBoard
        End If
    End If
    
    If [ballx] > 0 Then
        If [ballx] > [N4] And Not ([bally] >= [O5] And [bally] <= [O4]) Then
            isGameRunning = False
            [player1Score] = [player1Score] + 1
            Call SetScoreBoard
        End If
    End If
    

End Sub

Sub SetScoreBoard()
    On Error GoTo e
    For Each shp In Sheet4.Shapes
        If shp.Name Like "_*" Then shp.Delete
    Next
    Application.CutCopyMode = True
    Sheet3.Shapes("_" & [player1Score]).Copy
    Sheet4.Paste Sheet4.Range("E2")
    
    Sheet3.Shapes("_" & [player2Score]).Copy
    Sheet4.Paste Sheet4.Range("I2")
    Application.CutCopyMode = False
    Sheet4.Range("N11").Select
    Exit Sub
e:
    Call ClearScore
End Sub

Sub ClearScore()
    [player1Score] = 0: [player2Score] = 0
    Call SetScoreBoard
End Sub

Sub MoveBall()
    If [bally] <= -50 Or [bally] >= 50 Then
        [balldiry] = [balldiry] * -1
    End If
    
    If [ballx] <= -50 Or [ballx] >= 50 Then
        [balldirx] = [balldirx] * -1
    End If
    
    
    '' check player1
    
    If [ballx] < 0 Then
        If [ballx] <= [K4] And [bally] >= [L5] And [bally] <= [L4] Then
            [balldirx] = [balldirx] * -1
        End If
    End If
    
    
    '' check player2
    If [ballx] > 0 Then
        If [ballx] >= [N4] And [bally] >= [O5] And [bally] <= [O4] Then
            [balldirx] = [balldirx] * -1
        End If
    End If
    
    [bally] = [bally] + [balldiry]
    [ballx] = [ballx] + [balldirx]
End Sub

Sub ProcessInput()
    If GetAsyncKeyState(vbKeyPageUp) Then
        [player2Y] = WorksheetFunction.Min([player2Y] + 1, 10)
    ElseIf GetAsyncKeyState(vbKeyPageDown) Then
        [player2Y] = WorksheetFunction.Max([player2Y] - 1, -50)
    ElseIf GetAsyncKeyState(vbKeyUp) Then
        [player1Y] = WorksheetFunction.Min([player1Y] + 1, 10)
    ElseIf GetAsyncKeyState(vbKeyDown) Then
        [player1Y] = WorksheetFunction.Max([player1Y] - 1, -50)
    End If
End Sub


Public Sub CommandButton2_Click()
    Call ClearScore
    Call ResetPaddlePos
    [ballx] = 0: [bally] = 0
End Sub

Private Sub Worksheet_Activate()
    Application.OnKey "{PGDN}", ""
    Application.OnKey "{PGUP}", ""
    Application.OnKey "{UP}", ""
    Application.OnKey "{DOWN}", ""
End Sub

Private Sub Worksheet_Deactivate()
    Application.OnKey "{PGDN}"
    Application.OnKey "{PGUP}"
    Application.OnKey "{UP}"
    Application.OnKey "{DOWN}"
End Sub

