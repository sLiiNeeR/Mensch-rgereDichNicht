Option Explicit

    
    Dim NächsterSpieler As Boolean
    Dim i, j As Integer
    Dim k As Integer
    Dim diceValue As Integer
    'Route
    Dim btnRoad(40) As CommandButton
    'Safe Areas
    Dim btnRoadSafe(3, 3) As CommandButton
    'Start Points
    Dim btnHomes(3) As CommandButton
    'Figures (User/Group, id)
    Dim btnFigures(3, 3) As CommandButton
    Dim isFigureInSafeHome(3, 3) As Boolean
    'Figures Properties (position, color index)
    Dim FiguresP(3, 3) As Integer
    Dim FiguresC(3, 3) As Integer
    'Inital Positions
    Dim FiguresInitL(3, 3) As Integer
    Dim FiguresInitT(3, 3) As Integer
    Dim FiguresInitW(3, 3) As Integer
    Dim FiguresInitH(3, 3) As Integer
    
    Dim FiguresHomePosition(3) As Integer
    Dim FiguresStopPosition(3) As Integer
    Dim playerIndex As Integer
    Dim Images(5) As Image
    'At 39th position from start, follow/enter safe area
    'At 44th is Figure win


Private Sub btnRestHome_Click()

End Sub

Private Sub cmdHide_Click()
    HideAllOpponentsFigures
End Sub

Private Sub HideAllOpponentsFigures()
    For i = 0 To 3
        For j = 0 To 3
            If (i <> playerIndex) And FiguresP(i, j) >= 0 And FiguresP(i, j) <= 39 Then
                btnFigures(i, j).Visible = False
            End If
        Next
    Next
End Sub

Private Sub HideCollidedOpponentsFigures()
    For i = 0 To 3
        For j = 0 To 3
            If (i <> playerIndex) And FiguresP(i, j) >= 0 And FiguresP(i, j) <= 39 Then
                For k = 0 To 3
                    If FiguresP(i, j) = FiguresP(playerIndex, k) Then
                        btnFigures(i, j).Visible = False
                    End If
                Next
            End If
        Next
    Next
End Sub

Private Sub ShowAllAvailableFigures()
    For i = 0 To 3
        For j = 0 To 3
            If FiguresP(i, j) <= 39 + 4 Then
                btnFigures(i, j).Visible = True
            End If
        Next
    Next
End Sub

Private Sub cmdReset_Click()
    loadGameData
    
    
    For i = 0 To 3
        For j = 0 To 3
            btnFigures(i, j).Left = FiguresInitL(i, j)
            btnFigures(i, j).Top = FiguresInitT(i, j)
            btnFigures(i, j).Width = FiguresInitW(i, j)
            btnFigures(i, j).Height = FiguresInitH(i, j)
        Next
    Next
    
    
    For i = 0 To UBound(btnRoad) - 1
        btnRoad(i).Visible = True
    Next
    
    For j = 0 To 3
        For k = 0 To 3
            btnRoadSafe(j, k).Visible = True
        Next
    Next
End Sub

'Red Button Click
Private Sub R1_Click()
    Play 0, 0
End Sub

Private Sub R2_Click()
    Play 0, 1
End Sub

Private Sub R3_Click()
    Play 0, 2
End Sub

Private Sub R4_Click()
    Play 0, 3
End Sub

'Green Button Click
Private Sub G1_Click()
    Play 1, 0
End Sub

Private Sub G2_Click()
    Play 1, 1
End Sub

Private Sub G3_Click()
    Play 1, 2
End Sub

Private Sub G4_Click()
    Play 1, 3
End Sub


'Blue Button Click
Private Sub B1_Click()
    Play 2, 0
End Sub

Private Sub B2_Click()
    Play 2, 1
End Sub

Private Sub B3_Click()
    Play 2, 2
End Sub

Private Sub B4_Click()
    Play 2, 3
End Sub



Private Sub SB4_Click()

End Sub

'Yellow Button Click
Private Sub Y1_Click()
    Play 3, 0
End Sub

Private Sub Y2_Click()
    Play 3, 1
End Sub

Private Sub Y3_Click()
    Play 3, 2
End Sub

Private Sub Y4_Click()
    Play 3, 3
End Sub


Private Sub UserForm_Activate()
        
        loadGameData
        
        'Copy default locations
        For i = 0 To 3
            For j = 0 To 3
                FiguresInitL(i, j) = btnFigures(i, j).Left
                FiguresInitT(i, j) = btnFigures(i, j).Top
                FiguresInitW(i, j) = btnFigures(i, j).Width
                FiguresInitH(i, j) = btnFigures(i, j).Height
            Next
        Next
    End Sub

    Sub loadGameData()
        'Dice Images
        Set Images(0) = W1
        Set Images(1) = W2
        Set Images(2) = W3
        Set Images(3) = W4
        Set Images(4) = W5
        Set Images(5) = W6
        

        'Red Figures
        Set btnFigures(0, 0) = R1
        Set btnFigures(0, 1) = R2
        Set btnFigures(0, 2) = R3
        Set btnFigures(0, 3) = R4
        'Green Figures
        Set btnFigures(1, 0) = G1
        Set btnFigures(1, 1) = G2
        Set btnFigures(1, 2) = G3
        Set btnFigures(1, 3) = G4
        'Blue Figures
        Set btnFigures(2, 0) = B1
        Set btnFigures(2, 1) = B2
        Set btnFigures(2, 2) = B3
        Set btnFigures(2, 3) = B4
        'Yellow Figures
        Set btnFigures(3, 0) = Y1
        Set btnFigures(3, 1) = Y2
        Set btnFigures(3, 2) = Y3
        Set btnFigures(3, 3) = Y4

        'Safe Areas Red
        Set btnRoadSafe(0, 0) = SR1
        Set btnRoadSafe(0, 1) = SR2
        Set btnRoadSafe(0, 2) = SR3
        Set btnRoadSafe(0, 3) = SR4
        'Safe Areas Green
        Set btnRoadSafe(1, 0) = SG1
        Set btnRoadSafe(1, 1) = SG2
        Set btnRoadSafe(1, 2) = SG3
        Set btnRoadSafe(1, 3) = SG4
        'Safe Areas Blue
        Set btnRoadSafe(2, 0) = SB1
        Set btnRoadSafe(2, 1) = SB2
        Set btnRoadSafe(2, 2) = SB3
        Set btnRoadSafe(2, 3) = SB4
        'Safe Areas Yellow
        Set btnRoadSafe(3, 0) = SY1
        Set btnRoadSafe(3, 1) = SY2
        Set btnRoadSafe(3, 2) = SY3
        Set btnRoadSafe(3, 3) = SY4

        'Home Positions
        Set btnHomes(0) = H1
        Set btnHomes(1) = H2
        Set btnHomes(2) = H3
        Set btnHomes(3) = H4

        'Reset Positions and colors
        For i = 0 To 3
            For j = 0 To 3
                FiguresP(i, j) = -1
                FiguresC(i, j) = -1
                isFigureInSafeHome(i, j) = False
            Next
        Next

        Set btnRoad(0) = F1
        
        Set btnRoad(1) = H1
        FiguresHomePosition(0) = 1
        
        Set btnRoad(2) = F3
        Set btnRoad(3) = F4
        Set btnRoad(4) = F5
        Set btnRoad(5) = F6
        Set btnRoad(6) = F7
        Set btnRoad(7) = F8
        Set btnRoad(8) = F9
        Set btnRoad(9) = F10
        Set btnRoad(10) = F11
        
        Set btnRoad(11) = H2
        FiguresHomePosition(1) = 11
        
        Set btnRoad(12) = F13
        Set btnRoad(13) = F14
        Set btnRoad(14) = F15
        Set btnRoad(15) = F16
        Set btnRoad(16) = F17
        Set btnRoad(17) = F18
        Set btnRoad(18) = F19
        Set btnRoad(19) = F20
        Set btnRoad(20) = F21
        
        Set btnRoad(21) = H3
        FiguresHomePosition(2) = 21
        
        Set btnRoad(22) = F23
        Set btnRoad(23) = F24
        Set btnRoad(24) = F25
        Set btnRoad(25) = F26
        Set btnRoad(26) = F27
        Set btnRoad(27) = F28
        Set btnRoad(28) = F29
        Set btnRoad(29) = F30
        Set btnRoad(30) = F31
        
        Set btnRoad(31) = H4
        FiguresHomePosition(3) = 31
        
        Set btnRoad(32) = F33
        Set btnRoad(33) = F34
        Set btnRoad(34) = F35
        Set btnRoad(35) = F36
        Set btnRoad(36) = F37
        Set btnRoad(37) = F38
        Set btnRoad(38) = F39
        Set btnRoad(39) = F40
        
        FiguresStopPosition(0) = 39
        FiguresStopPosition(1) = 9
        FiguresStopPosition(2) = 19
        FiguresStopPosition(3) = 29
        
        DisableFigures
        'Red starts Game
        'playerIndex = 1
        'FiguresP(1, 0) = 35
        'Randomly starts Game
        Randomize
        playerIndex = Rnd() * (2 + 1)
        txtCurrentPlayer.Text = "Player " & (playerIndex + 1)
        
        
                'FiguresP(0, 0) = 39
                'FiguresP(1, 0) = 9
                'FiguresP(2, 0) = 19
                'FiguresP(3, 0) = 29
    End Sub





Private Sub HideWonFigures()
    For i = 0 To 3
        For j = 0 To 3
            If btnFigures(i, j).Left = btnRestHome.Left And btnFigures(i, j).Top = btnRestHome.Top Then
                btnFigures(i, j).Visible = False
            End If
        Next
    Next
End Sub

Private Sub DisableFigures()
    For i = 0 To 3
        For j = 0 To 3
            btnFigures(i, j).Enabled = False
        Next
    Next
    Würfel.Enabled = True
    txtInstruction.Text = "Throw Dice!"
End Sub

Private Sub EnableFigures()
    For j = 0 To 3
        If CanPlay(playerIndex, j) Then
            btnFigures(playerIndex, j).Enabled = True
        End If
    Next
    Würfel.Enabled = False
    txtInstruction.Text = "Move Figure!"
    HideCollidedOpponentsFigures
End Sub

Public Sub Würfel_Click()
    ShowAllAvailableFigures
    Randomize
    diceValue = Int(1 + (Rnd() * 6))
    'diceValue = 1
    TextBox1 = diceValue
    'Set Dice Image
    For i = 0 To 5
        If diceValue = i + 1 Then
            Images(i).Visible = True
        Else
            Images(i).Visible = False
        End If
    Next
    Dim canP As Boolean
    For j = 0 To 3
        If CanPlay(playerIndex, j) Then
            canP = True
            Exit For
        End If
    Next
    If canP Then
        EnableFigures
    Else
        SetNextPlayer
    End If
End Sub
   
Public Sub SetNextPlayer()
    If playerIndex < 3 Then
        playerIndex = playerIndex + 1
    Else
        playerIndex = 0
    End If
    
    txtCurrentPlayer.Text = "Player " & (playerIndex + 1)
End Sub

Public Sub Play(playerIndex As Integer, FigureIndex As Integer)
    If Not CanPlay(playerIndex, FigureIndex) Then
    Else
        Dim sumValue, x As Integer
        sumValue = FiguresP(playerIndex, FigureIndex) + diceValue
    
        Dim v, l, w, h, t As Integer
        Dim btn As CommandButton
        'Bring out new Figure
        If diceValue = 6 And FiguresP(playerIndex, FigureIndex) = -1 Then
            v = FiguresHomePosition(playerIndex)
            Set btn = btnRoad(v)
        'Can Move Figure
        ElseIf FiguresP(playerIndex, FigureIndex) >= 0 Then
            If playerIndex = 0 Then
                'Can Move
                If sumValue <= 39 Then
                    v = sumValue
                    Set btn = btnRoad(v)
                'Entered Safe region
                ElseIf sumValue <= 39 + 4 Or (isFigureInSafeHome(playerIndex, FigureIndex) = True And _
                    sumValue - 39 <= 4) Then
                    v = sumValue
                    x = sumValue - 39 - 1
                    Set btn = btnRoadSafe(playerIndex, x)
                    isFigureInSafeHome(playerIndex, FigureIndex) = True
                'Figure won
                ElseIf sumValue = 39 + 5 Then
                    isFigureInSafeHome(playerIndex, FigureIndex) = True
                    v = FiguresP(playerIndex, FigureIndex) + diceValue
                    'btnFigures(playerIndex, FigureIndex).Visible = False
                    isFigureInSafeHome(playerIndex, FigureIndex) = True
                    Set btn = btnRestHome
                    HideWonFigures
                End If
            Else
                If sumValue > 39 Then
                    sumValue = sumValue - 39 - 1
                Else
                    'sumValue = sumValue
                End If
                
                'Can Move
                If (FiguresP(playerIndex, FigureIndex) > FiguresStopPosition(playerIndex) Or _
                    sumValue <= FiguresStopPosition(playerIndex)) And _
                    isFigureInSafeHome(playerIndex, FigureIndex) = False Then
                    v = sumValue
                    Set btn = btnRoad(v)
                'Entered Safe region
                ElseIf ((FiguresP(playerIndex, FigureIndex) <= FiguresStopPosition(playerIndex) And _
                    sumValue >= FiguresStopPosition(playerIndex)) Or _
                    isFigureInSafeHome(playerIndex, FigureIndex) = True) And _
                    sumValue <= FiguresStopPosition(playerIndex) + 4 Then
                    v = sumValue
                    x = sumValue - FiguresStopPosition(playerIndex) - 1
                    Set btn = btnRoadSafe(playerIndex, x)
                    isFigureInSafeHome(playerIndex, FigureIndex) = True
                'Figure won
                ElseIf sumValue = FiguresStopPosition(playerIndex) + 5 Then
                    v = sumValue
                    isFigureInSafeHome(playerIndex, FigureIndex) = True
                    Set btn = btnRestHome
                    HideWonFigures
                End If
            
            End If
        'Can Not Move Figure
        Else
            MsgBox "Not supposed to reach here"
            DisableFigures
            Exit Sub
        End If
        FiguresP(playerIndex, FigureIndex) = v
        l = btn.Left
        t = btn.Top
        w = btn.Width
        h = btn.Height
        'Move Figure
        btnFigures(playerIndex, FigureIndex).Move l, t, w, h
        'btn.Visible = False
        
        'Hide all road buttons containing Figures and
        'Display all road buttons not containing Figures
        'Dim hasFigure As Boolean
        For i = 0 To UBound(btnRoad) - 1
            btnRoad(i).Visible = True
            For j = 0 To 3
                For k = 0 To 3
                    'If FiguresP(j, k) = i And FiguresP(j, k) > 0 And isFigureInSafeHome(j, k) = False Then
                    If FiguresP(j, k) = i And FiguresP(j, k) >= 0 And isFigureInSafeHome(j, k) = False Then
                        btnRoad(i).Visible = False
                    End If
                Next
            Next
        Next
        
        'Display all home buttons not containing Figures
        'Dim hasFigure As Boolean
        For j = 0 To 3
            For k = 0 To 3
                btnRoadSafe(j, k).Visible = True
            Next
        Next
        
        'Hide all home buttons containing Figures and
        For j = 0 To 3
            For k = 0 To 3
                For x = 0 To 4
                    If j = 0 Then
                        If FiguresP(j, k) > FiguresStopPosition(j) And isFigureInSafeHome(j, k) = True And _
                            FiguresP(j, k) <= FiguresStopPosition(j) + 4 And _
                             FiguresP(j, k) - FiguresStopPosition(j) - 1 = x Then
                            btnRoadSafe(j, x).Visible = False
                        End If
                    Else
                        If FiguresP(j, k) > FiguresStopPosition(j) And isFigureInSafeHome(j, k) = True And _
                            FiguresP(j, k) <= FiguresStopPosition(j) + 4 And _
                             FiguresP(j, k) - FiguresStopPosition(j) - 1 = x Then
                            btnRoadSafe(j, x).Visible = False
                        End If
                    End If
                Next
            Next
        Next
        
        DisableFigures
        If diceValue <> 6 Then
            SetNextPlayer
        End If
    End If
End Sub
      
Public Function CanPlay(playerIndex As Integer, FigureIndex As Integer)
    Dim sumValue, x As Integer
    sumValue = FiguresP(playerIndex, FigureIndex) + diceValue
    CanPlay = False
    'Can bring out new Figure
    If diceValue = 6 And FiguresP(playerIndex, FigureIndex) = -1 Then
        CanPlay = True
    'Can Move Figure
    ElseIf FiguresP(playerIndex, FigureIndex) >= 0 Then
        If playerIndex = 0 Then
            'Can Move
            If sumValue <= 39 Then
                CanPlay = True
            'Entered Safe region
            ElseIf sumValue <= 39 + 4 Or (isFigureInSafeHome(playerIndex, FigureIndex) = True And _
                sumValue - 39 <= 4) Then
                CanPlay = True
            'Figure won
            ElseIf sumValue = 39 + 5 Then
                CanPlay = True
            End If
        Else
            If sumValue > 39 Then
                sumValue = sumValue - 39 - 1
            Else
                'sumValue = sumValue
            End If
            
            'Can Move
            If (FiguresP(playerIndex, FigureIndex) > FiguresStopPosition(playerIndex) Or _
                sumValue <= FiguresStopPosition(playerIndex)) And _
                isFigureInSafeHome(playerIndex, FigureIndex) = False Then
                CanPlay = True
            'Entered Safe region
            ElseIf ((FiguresP(playerIndex, FigureIndex) <= FiguresStopPosition(playerIndex) And _
                sumValue >= FiguresStopPosition(playerIndex)) Or _
                isFigureInSafeHome(playerIndex, FigureIndex) = True) And _
                sumValue <= FiguresStopPosition(playerIndex) + 4 Then
                CanPlay = True
                'Figure won
            ElseIf sumValue = FiguresStopPosition(playerIndex) + 5 Then
                CanPlay = True
            End If
        
        End If
    'Can Not Move Figure
    Else
        CanPlay = False
    End If
    
End Function


 

