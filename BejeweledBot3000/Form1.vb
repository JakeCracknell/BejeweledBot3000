Imports System.Drawing.Imaging
Imports System.Threading

Public Class Form1
    Declare Function apimouse_event Lib "user32.dll" Alias "mouse_event" (ByVal dwFlags As Int32, ByVal dX As Int32, ByVal dY As Int32, ByVal cButtons As Int32, ByVal dwExtraInfo As Int32) As Boolean
    Public Const MOUSEEVENTF_LEFTDOWN = &H2
    Public Const MOUSEEVENTF_LEFTUP = &H4
    Declare Function GetWindowRect Lib "user32.dll" Alias "GetWindowRect" (hwnd As IntPtr, ByRef rectangle As Rect) As IntPtr
    Declare Function GetKeyState Lib "user32" Alias "GetKeyState" (ByVal ByValnVirtKey As Int32) As Int16
    Private Const VK_CAPSLOCK = &H14
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
            If GetKeyState(VK_CAPSLOCK) = 1 Then
                PerformMoveUsingMouse(move)
            End If
        End If
    End Sub

    Function GetMove() As BejeweledMove
        Dim bFull As New Bitmap(TileCount * TileSize, TileCount * TileSize)
        Dim gFull As Graphics = Graphics.FromImage(bFull)
        gFull.CopyFromScreen(GetTopLeftOfBejeweledBoard, New Point(0, 0), bFull.Size)
        bFull = bFull.Clone(New Rectangle(0, 0, bFull.Width, bFull.Height), PixelFormat.Format8bppIndexed)
        bFull = bFull.Clone(New Rectangle(0, 0, bFull.Width, bFull.Height), PixelFormat.Format16bppRgb555)

        gFull = Graphics.FromImage(bFull)


        Dim BejeweledBoard As New BejeweledBoard(TileCount)
        For x = 0 To TileCount - 1
            For y = 0 To TileCount - 1
                Dim m As Double = 0.4
                Dim sampleRectangle = New Rectangle((TileSize * m) + (x * TileSize), (TileSize * m) + (y * TileSize),
                                                    TileSize * (1 - m * 2), TileSize * (1 - m * 2))
                Dim colorBucket As New ColorBucket()
                For i = sampleRectangle.Left To sampleRectangle.Right
                    For j = sampleRectangle.Top To sampleRectangle.Bottom
                        colorBucket.Colors.Add(bFull.GetPixel(i, j))
                    Next
                Next
                Dim tileCode = colorBucket.GetColorCode
                gFull.DrawRectangle(New Pen(Color.FromArgb(tileCode), 2), sampleRectangle)

                BejeweledBoard.SetTile(x, y, tileCode)
            Next
        Next

        PictureBox1.Image = bFull

        For x = 0 To TileCount - 1
            For y = 0 To TileCount - 1
                dgvBoard.Rows(y).Cells(x).Style.BackColor = Color.FromArgb(BejeweledBoard.GetTile(x, y))
            Next
        Next
        BejeweledBoard.Normalise()
        For x = 0 To TileCount - 1
            For y = 0 To TileCount - 1
                dgvBoard.Rows(y).Cells(x).Value = BejeweledBoard.GetTile(x, y)
            Next
        Next

        If BejeweledBoard.IsValidBoard Then
            Dim move As BejeweledMove = BejeweledBoard.FindMove()
            Me.Text = BejeweledBoard.GetUniqueTileCount & " unique tiles on VALID board - best move " & move.ToString
            Return move
        Else
            Me.Text = BejeweledBoard.GetScore & " - board not valid as it already contains matches"
            Return Nothing
        End If
    End Function

    Sub PerformMoveUsingMouse(move As BejeweledMove)
        Cursor.Position = GetTopLeftOfBejeweledBoard()
        Cursor.Position = New Point(Cursor.Position.X + (TileSize * move.X), Cursor.Position.Y + (TileSize * move.Y))
        Thread.Sleep(200)
        Call apimouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
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
        Call apimouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    End Sub

    Function GetTopLeftOfBejeweledBoard()
        Dim handle As IntPtr = Process.GetProcessesByName("Bejeweled3")(0).MainWindowHandle
        Dim rect As New Rect
        GetWindowRect(handle, rect)
        Return New Point(rect.Left + 7 + 4.15 * TileSize, rect.Top + 25 + 0.65 * TileSize)
    End Function


End Class
