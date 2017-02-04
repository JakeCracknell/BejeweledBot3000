Imports System.Drawing.Imaging
Imports System.Threading

Public Class Form1


    Dim OscillationDetector As New OscillationDetector
    Dim BejeweledScreenReader As New BejeweledScreenReader

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        dgvBoard.RowCount = TileCount
        dgvBoard.ColumnCount = TileCount
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Console.Beep()
        Thread.Sleep(200)
        Console.Beep()
        Thread.Sleep(200)
        Console.Beep()
        Thread.Sleep(200)

        GetMoves().ForEach(Sub(m) PerformMoveUsingMouse(m))
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Dim bejeweledLocation As Rect = BejeweledScreenReader.GetBejeweledWindowRect()
        Me.Location = New Point(bejeweledLocation.Right, bejeweledLocation.Top)
        Dim moves = GetMoves()
        If cbWaitForStaticBoard.Checked AndAlso BejeweledBoardIsStatic() Then
            PerformMovesUsingMouseIfCapsLock(moves)
            TryClickOnPlayAgainButtonIfCapsLock()
        Else
            LOG("Waiting for static board")
        End If

    End Sub

    Private Function BejeweledBoardIsStatic() As Boolean
        Dim bejeweledBoard1 As BejeweledBoard = BejeweledScreenReader.DrawAndGetBejeweledBoardFromScreen_VeryQuick()
        Dim bejeweledBoard2 As BejeweledBoard = BejeweledScreenReader.DrawAndGetBejeweledBoardFromScreen_VeryQuick()
        Return OscillationDetector.GetDifference(bejeweledBoard1, bejeweledBoard2) <= 3
    End Function

    Function GetMoves() As List(Of BejeweledMove)
        LOG("Reading from screen")
        Dim bejeweledBoard As BejeweledBoard = BejeweledScreenReader.DrawAndGetBejeweledBoardFromScreen()
        DrawBoardOnForm(bejeweledBoard)
        If (OscillationDetector.LogBoardAndReturnTrueIfOscillationDetected(bejeweledBoard)) Then
            LOG("Oscillation detected - sleeping")
            Thread.Sleep(500) 'Wait half a second for board to calm down.
            Return New List(Of BejeweledMove)
        Else
            Dim moves As List(Of BejeweledMove) = bejeweledBoard.FindMoves()
            Me.Text = bejeweledBoard.GetUniqueTileCount & " unique tiles - moves " & moves.Count &
                             IIf(bejeweledBoard.IsValidBoard, "", " - Board looks invalid")
            Return moves
        End If
    End Function

    Sub DrawBoardOnForm(bejeweledBoard As BejeweledBoard)
        PictureBox1.Image = BejeweledScreenReader.LatestBitmap
        For x = 0 To TileCount - 1
            For y = 0 To TileCount - 1
                Dim bejeweledTile As BejeweledTile = bejeweledBoard.GetTile(x, y)
                dgvBoard.Rows(y).Cells(x).Style.BackColor = Color.FromArgb(bejeweledTile.TileCode)
                dgvBoard.Rows(y).Cells(x).Value = bejeweledTile.NormalisedTileCode
            Next
        Next
    End Sub

    Sub PerformMovesUsingMouseIfCapsLock(bejeweledMoves As List(Of BejeweledMove))
        If GetKeyState(VK_CAPSLOCK) = 1 Then
            LOG("Performing " & bejeweledMoves.Count & " moves")
            For Each bejeweledMove In bejeweledMoves
                PerformMoveUsingMouse(bejeweledMove)
            Next
        End If
    End Sub


    Sub PerformMoveUsingMouse(move As BejeweledMove)
        Dim basePosition = BejeweledScreenReader.GetTopLeftOfBejeweledBoard()
        Cursor.Position = New Point(basePosition.X + (TileSize * move.X), basePosition.Y + (TileSize * move.Y))
        Call apimouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        Select Case move.Direction
            Case ArrowDirection.Down
                Cursor.Position = New Point(Cursor.Position.X, Cursor.Position.Y - TileSize)
            Case ArrowDirection.Up
                Cursor.Position = New Point(Cursor.Position.X, Cursor.Position.Y + TileSize)
            Case ArrowDirection.Left
                Cursor.Position = New Point(Cursor.Position.X - TileSize, Cursor.Position.Y)
            Case ArrowDirection.Right
                Cursor.Position = New Point(Cursor.Position.X + TileSize, Cursor.Position.Y)
        End Select
        Thread.Sleep(25)
        Call apimouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    End Sub

    Private Sub TryClickOnPlayAgainButtonIfCapsLock()
        If GetKeyState(VK_CAPSLOCK) = 1 Then
            TryClickOnPlayAgainButton()
        End If
    End Sub

    Private Sub TryClickOnPlayAgainButton()
        Dim oldPosition = Cursor.Position
        Dim b As Rect = BejeweledScreenReader.GetBejeweledWindowRect()
        Cursor.Position = New Point(b.Left + (b.Right - b.Left) / 2, b.Top + 440)
        Call apimouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        Call apimouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        Cursor.Position = oldPosition
    End Sub

    Sub LOG(str As String)
        lblStatus.Text = str.ToUpper
    End Sub
End Class
