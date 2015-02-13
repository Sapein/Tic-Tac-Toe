Public Class Form1
    'These are the primary Variables these are used to decide on certain things
    Dim intPriority() As Integer = {0, 0, 0, 0, 0, 0, 0, 0, 0}
    Dim intBoard() As Integer = {0, 1, 2, 3, 4, 5, 6, 7, 8}
    Dim boolComputerTurn As Boolean = False

    Private Sub picTopLeft_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles picTopLeft.Click
        'What this does is it checks to see if it's the computer's turn, if it is...allow the player to 
        'Move, otherwise do the computer's turn and check to see if the game is at a draw
        If boolComputerTurn = False Then
            picTopLeft.BackgroundImage = System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\..\..\..\images\X.bmp")
            intPriority(0) = 4
            PriorityCheck(1, 3, 4)
            boolComputerTurn = True
            WinCheck(0, 1, 2)
            WinCheck(0, 4, 8)
            WinCheck(0, 3, 6)
        End If
        Computer()
        DrawCheck()
    End Sub

    Private Sub picTopCenter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles picTopCenter.Click
        If boolComputerTurn = False Then
            picTopCenter.BackgroundImage = System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\..\..\..\images\X.bmp")
            intPriority(1) = 4
            PriorityCheck(0, 2, 4)
            WinCheck(0, 1, 2)
            WinCheck(1, 4, 7)
            boolComputerTurn = True
        End If
        Computer()
        DrawCheck()
    End Sub

    Private Sub picTopRight_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles picTopRight.Click
        If boolComputerTurn = False Then
            picTopRight.BackgroundImage = System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\..\..\..\images\X.bmp")
            intPriority(2) = 4
            PriorityCheck(4, 1, 5)
            WinCheck(0, 1, 2)
            WinCheck(2, 5, 8)
            WinCheck(2, 4, 6)
            boolComputerTurn = True
        End If
        Computer()
        DrawCheck()
    End Sub

    Private Sub picCenterLeft_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles picCenterLeft.Click
        If boolComputerTurn = False Then
            picCenterLeft.BackgroundImage = System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\..\..\..\images\X.bmp")
            intPriority(3) = 4
            PriorityCheck(0, 4, 6)
            boolComputerTurn = True
            WinCheck(0, 3, 6)
            WinCheck(3, 4, 5)
        End If
        Computer()
        DrawCheck()
    End Sub

    Private Sub picCenter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles picCenter.Click
        If boolComputerTurn = False Then
            picCenter.BackgroundImage = System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\..\..\..\images\X.bmp")
            intPriority(4) = 4
            Dim x As Integer = 0
            'This doesn't use PriorityCheck(x, y, z) because of the fact that it needs to modify ALL 
            'squares
            While x <= 8
                If intPriority(x) > 4 And intPriority(x) < 0 Then
                    intPriority(x) = intPriority(x) + 1
                End If
                x = x + 1
            End While
            boolComputerTurn = True
            WinCheck(3, 4, 5)
            WinCheck(1, 4, 7)
            WinCheck(0, 4, 8)
            WinCheck(2, 4, 6)
        End If
        Computer()
        DrawCheck()
    End Sub

    Private Sub picCenterRight_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles picCenterRight.Click
        If boolComputerTurn = False Then
            picCenterRight.BackgroundImage = System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\..\..\..\images\X.bmp")
            intPriority(5) = 4
            PriorityCheck(4, 8, 2)
            boolComputerTurn = True
            WinCheck(2, 5, 8)
            WinCheck(3, 4, 5)
        End If
        Computer()
        DrawCheck()
    End Sub

    Private Sub picBottomLeft_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles picBottomLeft.Click
        If boolComputerTurn = False Then
            picBottomLeft.BackgroundImage = System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\..\..\..\images\X.bmp")
            intPriority(6) = 4
            PriorityCheck(3, 4, 7)
            boolComputerTurn = True
            WinCheck(6, 7, 8)
            WinCheck(0, 3, 6)
            WinCheck(2, 4, 6)
        End If
        Computer()
        DrawCheck()
    End Sub

    Private Sub picBottomCenter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles picBottomCenter.Click
        If boolComputerTurn = False Then
            picBottomCenter.BackgroundImage = System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\..\..\..\images\X.bmp")
            intPriority(7) = 4
            PriorityCheck(6, 8, 4)
            boolComputerTurn = True
            WinCheck(6, 7, 8)
            WinCheck(1, 4, 7)
        End If
        Computer()
        DrawCheck()
    End Sub

    Private Sub picBottomRight_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles picBottomRight.Click
        If boolComputerTurn = False Then
            picBottomRight.BackgroundImage = System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\..\..\..\images\X.bmp")
            intPriority(8) = 4
            PriorityCheck(4, 5, 7)
            boolComputerTurn = True
            WinCheck(6, 7, 8)
            WinCheck(2, 5, 8)
            WinCheck(0, 4, 8)
        End If
        Computer()
        DrawCheck()
    End Sub

    Private Sub Reset()
        'This resets the game, it resets all priority to 0 and moves again, if it didn't
        'The Checks would mess up and both players wouldn't be able to move
        picBottomRight.BackgroundImage = System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\..\..\..\images\Blank.bmp")
        picBottomCenter.BackgroundImage = System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\..\..\..\images\Blank.bmp")
        picBottomLeft.BackgroundImage = System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\..\..\..\images\Blank.bmp")
        picCenterRight.BackgroundImage = System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\..\..\..\images\Blank.bmp")
        picCenter.BackgroundImage = System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\..\..\..\images\Blank.bmp")
        picCenterLeft.BackgroundImage = System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\..\..\..\images\Blank.bmp")
        picTopRight.BackgroundImage = System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\..\..\..\images\Blank.bmp")
        picTopCenter.BackgroundImage = System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\..\..\..\images\Blank.bmp")
        picTopLeft.BackgroundImage = System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\..\..\..\images\Blank.bmp")
        intPriority(0) = 0
        intPriority(1) = 0
        intPriority(2) = 0
        intPriority(3) = 0
        intPriority(4) = 0
        intPriority(5) = 0
        intPriority(6) = 0
        intPriority(7) = 0
        intPriority(8) = 0
    End Sub

    Private Sub PriorityCheck(ByVal increment1, ByVal increment2, ByVal increment3)
        'This checks to make sure it doesn't overwrite the other square's Priority
        If intPriority(increment2) <> 4 Then
            If intPriority(increment2) = 0 Then
                intPriority(increment2) = 1
            ElseIf intPriority(increment2) > 0 And intPriority(increment2) < 4 Then
                intPriority(increment2) = intPriority(increment2) + 1
            End If
        End If
        If intPriority(increment3) <> 4 Then
            If intPriority(increment3) = 0 Then
                intPriority(increment3) = 1
            ElseIf intPriority(increment3) > 0 And intPriority(increment3) < 4 Then
                intPriority(increment3) = intPriority(increment3) + 1
            End If
        End If
        If intPriority(increment1) <> 4 Then
            If intPriority(increment1) = 0 Then
                intPriority(increment1) = 1

            ElseIf intPriority(increment1) > 0 And intPriority(increment1) < 4 Then
                intPriority(increment1) = intPriority(increment1) + 1
            End If
        End If
    End Sub

    Private Sub WinCheck(ByVal check1, ByVal check2, ByVal check3)
        'This checks to see if the person has won
        If intPriority(check1) = 4 And intPriority(check2) = 4 And intPriority(check3) = 4 Then
            boolComputerTurn = False
            MessageBox.Show("You Win!")
            Reset()
        End If
    End Sub
    Private Sub LossCheck(ByVal check1, ByVal check2, ByVal check3)
        'Checks to see if the Computer Has Won
        If intPriority(check1) = -1 And intPriority(check2) = -1 And intPriority(check3) = -1 Then
            MessageBox.Show("You Loose!")
            Reset()
            boolComputerTurn = False
        End If
    End Sub
    Private Sub Computer()
        'This is the computer's movements, this is how the computer works
        '(See Priority and Computer Documentation for more indepth details)
        Dim x As Integer = 0
        Dim y As Integer = 0
        Dim z As Integer = 1
        Dim highestNumber As Integer = 0
        Dim randomSelection As Integer
        DrawCheck()
        WinCheck(3, 4, 5)
        WinCheck(1, 4, 7)
        WinCheck(0, 4, 8)
        WinCheck(2, 4, 6)
        WinCheck(0, 4, 8)
        WinCheck(6, 7, 8)
        WinCheck(0, 1, 2)
        WinCheck(2, 5, 8)
        If boolComputerTurn = True Then
            While x <= 8
                If intPriority(x) > highestNumber And intPriority(x) <> 4 Then
                    highestNumber = intPriority(x)
                End If
                x = x + 1
            End While
            x = 0
            While x <= 8
                If intPriority(x) = highestNumber Then
                    y = y + 1
                End If
                x = x + 1
            End While
            Randomize()
            randomSelection = CInt(Int((y * Rnd()) + 1))
            x = 0
            While x <= 8
                If intPriority(x) = highestNumber Then
                    If z = randomSelection Then
                        If intBoard(x) = 0 Then
                            picTopLeft.BackgroundImage = System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\..\..\..\images\O.bmp")
                            intPriority(x) = -1
                        ElseIf intBoard(x) = 1 Then
                            picTopCenter.BackgroundImage = System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\..\..\..\images\O.bmp")
                            intPriority(x) = -1
                        ElseIf intBoard(x) = 2 Then
                            picTopRight.BackgroundImage = System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\..\..\..\images\O.bmp")
                            intPriority(x) = -1
                        ElseIf intBoard(x) = 3 Then
                            picCenterLeft.BackgroundImage = System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\..\..\..\images\O.bmp")
                            intPriority(x) = -1
                        ElseIf intBoard(x) = 4 Then
                            picCenter.BackgroundImage = System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\..\..\..\images\O.bmp")
                            intPriority(x) = -1
                        ElseIf intBoard(x) = 5 Then
                            picCenterRight.BackgroundImage = System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\..\..\..\images\O.bmp")
                            intPriority(x) = -1
                        ElseIf intBoard(x) = 6 Then
                            picBottomLeft.BackgroundImage = System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\..\..\..\images\O.bmp")
                            intPriority(x) = -1
                        ElseIf intBoard(x) = 7 Then
                            picBottomCenter.BackgroundImage = System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\..\..\..\images\O.bmp")
                            intPriority(x) = -1
                        ElseIf intBoard(x) = 8 Then
                            picBottomRight.BackgroundImage = System.Drawing.Bitmap.FromFile(My.Application.Info.DirectoryPath & "\..\..\..\images\O.bmp")
                            intPriority(x) = -1
                        End If
                    End If
                    z = z + 1
                End If
                x = x + 1
            End While
        End If
        LossCheck(3, 4, 5)
        LossCheck(1, 4, 7)
        LossCheck(0, 4, 8)
        LossCheck(2, 4, 6)
        LossCheck(0, 4, 8)
        LossCheck(6, 7, 8)
        LossCheck(0, 1, 2)
        LossCheck(2, 5, 8)
        boolComputerTurn = False
    End Sub
    Private Sub DrawCheck()
        'This checks for Draws, if neither the player, nor the comptuer, can move then it ends here.
        Dim x As Integer
        Dim y As Integer = 0
        While x <= 8
            If intPriority(x) = 4 Or intPriority(x) = -1 Then
                y = y + 1
            End If
            x = x + 1
        End While
        If y = 9 Then
            MessageBox.Show("DRAW!")
            boolComputerTurn = False
            Reset()
        End If
    End Sub
End Class
