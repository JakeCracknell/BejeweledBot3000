Public Class OscillationDetector
    Private boards(2) As BejeweledBoard
    Private nextInsertionIndex As Byte = 0

    Function LogBoardAndReturnTrueIfOscillationDetected(bejeweledBoard As BejeweledBoard)
        Dim board2Ago As BejeweledBoard = boards(nextInsertionIndex)
        boards(nextInsertionIndex) = bejeweledBoard
        nextInsertionIndex = (nextInsertionIndex + 1) Mod 2
        Dim board1Ago As BejeweledBoard = boards(nextInsertionIndex)

        Dim differencesTo1Ago As Integer = 0
        Dim differencesTo2Ago As Integer = 0
        If board2Ago IsNot Nothing Then
            For x = 0 To TileCount - 1
                For y = 0 To TileCount - 1
                    If board1Ago.GetTile(x, y).TileCode <> bejeweledBoard.GetTile(x, y).TileCode Then
                        differencesTo1Ago += 1
                    End If
                    If board2Ago.GetTile(x, y).TileCode <> bejeweledBoard.GetTile(x, y).TileCode Then
                        differencesTo2Ago += 1
                    End If
                Next
            Next
        End If
        Return differencesTo1Ago > 0 And differencesTo2Ago < 3
    End Function
End Class
