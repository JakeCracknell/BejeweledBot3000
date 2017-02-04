Imports System.Drawing.Imaging

Public Class BejeweledScreenReader
    Public LatestBitmap As New Bitmap(1, 1)

    Function DrawAndGetBejeweledBoardFromScreen() As BejeweledBoard
        Return DrawAndGetBejeweledBoardFromScreen(0.25, False)
    End Function

    Function DrawAndGetBejeweledBoardFromScreen(m As Double, useMode As Boolean) As BejeweledBoard
        Dim bFull As New Bitmap(TileCount * TileSize, TileCount * TileSize)
        Dim gFull As Graphics = Graphics.FromImage(bFull)
        gFull.CopyFromScreen(GetTopLeftOfBejeweledBoard, New Point(0, 0), bFull.Size)
        bFull = bFull.Clone(New Rectangle(0, 0, bFull.Width, bFull.Height), PixelFormat.Format8bppIndexed)
        bFull = bFull.Clone(New Rectangle(0, 0, bFull.Width, bFull.Height), PixelFormat.Format16bppRgb555)
        gFull = Graphics.FromImage(bFull)
        LatestBitmap = bFull

        Dim BejeweledBoard As New BejeweledBoard(TileCount)
        For x = 0 To TileCount - 1
            For y = 0 To TileCount - 1
                Dim sampleRectangle = New Rectangle((TileSize * m) + (x * TileSize), (TileSize * m) + (y * TileSize),
                                                    TileSize * (1 - m * 2), TileSize * (1 - m * 2))
                Dim colorBucket As New ColorBucket()
                For i = sampleRectangle.Left To sampleRectangle.Right
                    For j = sampleRectangle.Top To sampleRectangle.Bottom
                        colorBucket.Colors.Add(bFull.GetPixel(i, j))
                    Next
                Next
                Dim tileCode
                If useMode Then
                    tileCode = colorBucket.GetColorCode_MODE
                Else
                    tileCode = colorBucket.GetColorCode()
                End If
                gFull.DrawRectangle(New Pen(Color.FromArgb(tileCode), 2), sampleRectangle)
                BejeweledBoard.SetTile(x, y, tileCode)
            Next
        Next
        BejeweledBoard.Normalise()
        Return BejeweledBoard
    End Function

    Function DrawAndGetBejeweledBoardFromScreen_VeryQuick() As BejeweledBoard
        Return DrawAndGetBejeweledBoardFromScreen(0.49, True)
    End Function

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
