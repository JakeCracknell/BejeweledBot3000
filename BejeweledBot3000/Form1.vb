﻿Imports System.Drawing.Imaging
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

        GetMoves().ForEach(Sub(m) PerformMoveUsingMouse(m))
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Dim bejeweledLocation As Rect = GetBejeweledWindowRect()
        Me.Location = New Point(bejeweledLocation.Right, bejeweledLocation.Top)
        GetMoves().ForEach(Sub(m) PerformMoveUsingMouseIfCapsLock(m))
        TryClickOnPlayAgainButtonIfCapsLock()
    End Sub


    Function GetMoves() As List(Of BejeweledMove)
        Dim bFull As New Bitmap(TileCount * TileSize, TileCount * TileSize)
        Dim gFull As Graphics = Graphics.FromImage(bFull)
        gFull.CopyFromScreen(GetTopLeftOfBejeweledBoard, New Point(0, 0), bFull.Size)
        bFull = bFull.Clone(New Rectangle(0, 0, bFull.Width, bFull.Height), PixelFormat.Format8bppIndexed)
        bFull = bFull.Clone(New Rectangle(0, 0, bFull.Width, bFull.Height), PixelFormat.Format16bppRgb555)

        gFull = Graphics.FromImage(bFull)


        Dim BejeweledBoard As New BejeweledBoard(TileCount)
        For x = 0 To TileCount - 1
            For y = 0 To TileCount - 1
                Dim m As Double = 0.25
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
                dgvBoard.Rows(y).Cells(x).Style.BackColor = Color.FromArgb(BejeweledBoard.GetTile(x, y).TileCode)
            Next
        Next
        BejeweledBoard.Normalise()
        For x = 0 To TileCount - 1
            For y = 0 To TileCount - 1
                dgvBoard.Rows(y).Cells(x).Value = BejeweledBoard.GetTile(x, y).TileCode
            Next
        Next

        If BejeweledBoard.IsValidBoard Then
            Dim moves As List(Of BejeweledMove) = BejeweledBoard.FindMoves()
            Me.Text = BejeweledBoard.GetUniqueTileCount & " unique tiles on VALID board - moves " & moves.Count
            Return moves
        Else
            Me.Text = BejeweledBoard.GetScore & " - board not valid as it already contains matches"
            Return New List(Of BejeweledMove)
        End If
    End Function

    Sub PerformMoveUsingMouseIfCapsLock(move As BejeweledMove)
        If GetKeyState(VK_CAPSLOCK) = 1 Then
            PerformMoveUsingMouse(move)
        End If
    End Sub


    Sub PerformMoveUsingMouse(move As BejeweledMove)
        Dim basePosition = GetTopLeftOfBejeweledBoard()
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
        Dim b As Rect = GetBejeweledWindowRect()
        Cursor.Position = New Point(b.Left + (b.Right - b.Left) / 2, b.Top + 440)
        Call apimouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        Call apimouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        Cursor.Position = oldPosition
    End Sub

    Function GetTopLeftOfBejeweledBoard()
        Dim b As Rect = GetBejeweledWindowRect()
        Return New Point(b.Left + 7 + 4.15 * TileSize, b.Top + 25 + 0.65 * TileSize)
    End Function

    Function GetBejeweledWindowRect() As Rect
        Dim handle As IntPtr = Process.GetProcessesByName("Bejeweled3")(0).MainWindowHandle
        Dim rect As New Rect
        GetWindowRect(handle, rect)
        Return rect
    End Function


End Class
