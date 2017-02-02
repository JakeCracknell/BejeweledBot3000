Imports System.Drawing.Imaging
Imports System.Threading

Public Class Form1
    Declare Function apimouse_event Lib "user32.dll" Alias "mouse_event" (ByVal dwFlags As Int32, ByVal dX As Int32, ByVal dY As Int32, ByVal cButtons As Int32, ByVal dwExtraInfo As Int32) As Boolean
    Public Const MOUSEEVENTF_LEFTDOWN = &H2
    Public Const MOUSEEVENTF_LEFTUP = &H4
    Declare Function GetWindowRect Lib "user32.dll" Alias "GetWindowRect" (hwnd As IntPtr, ByRef rectangle As Rect) As IntPtr

    Structure Rect
        Public Left As Int32
        Public Top As Int32
        Public Right As Int32
        Public Bottom As Int32
    End Structure

    Dim TileSize As Integer = 50
    Dim TileCount As Integer = 8

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

        Cursor.Position = GetTopLeftOfBejeweledBoard()
        Dim move = GetMove()
        If move IsNot Nothing Then
            PerformMoveUsingMouse(move)
        End If
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Dim move = GetMove()
        If move IsNot Nothing Then
            PerformMoveUsingMouse(move)
        End If
    End Sub

    Function GetMove() As BejeweledMove
        Dim bFull As New Bitmap(TileCount * TileSize, TileCount * TileSize)
        Dim gFull As Graphics = Graphics.FromImage(bFull)
        gFull.CopyFromScreen(GetTopLeftOfBejeweledBoard, New Point(0, 0), bFull.Size)
        'gFull.InterpolationMode = Drawing2D.InterpolationMode.NearestNeighbor
        Dim bSmall As New Bitmap(TileCount, TileCount)
        Dim gSmall As Graphics = Graphics.FromImage(bSmall)
        gSmall.DrawImage(bFull, 0, 0, TileCount, TileCount)
        bSmall = bSmall.Clone(New Rectangle(0, 0, bSmall.Width, bSmall.Height), PixelFormat.Format8bppIndexed)

        PictureBox1.Image = bSmall

        Dim BejeweledBoard As New BejeweledBoard(TileCount)
        For x = 0 To TileCount - 1
            For y = 0 To TileCount - 1
                Dim tileCode = bSmall.GetPixel(x, y).ToArgb
                BejeweledBoard.SetTile(x, y, tileCode)
            Next
        Next
        BejeweledBoard.Normalise()
        For x = 0 To TileCount - 1
            For y = 0 To TileCount - 1
                dgvBoard.Rows(y).Cells(x).Value = BejeweledBoard.GetTile(x, y)
            Next
        Next

        Me.Text = BejeweledBoard.GetUniqueTileCount
        If BejeweledBoard.IsValidBoard Then
            Dim move As BejeweledMove = BejeweledBoard.FindMove()
            Button1.Text = move.ToString
            Return move
        Else
            Button1.Text = BejeweledBoard.GetScore
            Return Nothing
        End If
    End Function

    Sub PerformMoveUsingMouse(move As BejeweledMove)
        Dim oldPosition = Cursor.Position
        Cursor.Position = GetTopLeftOfBejeweledBoard()
        Cursor.Position = New Point(Cursor.Position.X + (TileSize * move.X), Cursor.Position.Y + (TileSize * move.Y))
        Thread.Sleep(200)
        'Call apimouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        Select Case move.Direction
            Case ArrowDirection.Down
                Cursor.Position = New Point(Cursor.Position.X, Cursor.Position.Y - TileSize)
            Case ArrowDirection.Up
                Cursor.Position = New Point(Cursor.Position.X, Cursor.Position.Y + TileSize)
            Case ArrowDirection.Left
                Cursor.Position = New Point(Cursor.Position.X - TileSize, Cursor.Position.Y)
            Case Else 'ArrowDirection.Right
                Cursor.Position = New Point(Cursor.Position.X + TileSize, Cursor.Position.Y)
        End Select
        Thread.Sleep(200)

        Cursor.Position = oldPosition
        'Call apimouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    End Sub

    Function GetTopLeftOfBejeweledBoard()
        Dim handle As IntPtr = Process.GetProcessesByName("Bejeweled3")(0).MainWindowHandle
        Dim rect As New Rect
        GetWindowRect(handle, rect)
        Return New Point(rect.Left + 7 + 4.15 * TileSize, rect.Top + 25 + 0.65 * TileSize)
    End Function


End Class
