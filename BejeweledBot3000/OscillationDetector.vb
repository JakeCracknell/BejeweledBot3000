Imports BejeweledBot3000

Public Class OscillationDetector
    Private boards(2) As BejeweledBoard
    Private nextInsertionIndex As Byte = 0

    Function LogBoardAndReturnTrueIfOscillationDetected(bejeweledBoard As BejeweledBoard)
        Dim board2Ago As BejeweledBoard = boards(nextInsertionIndex)
        boards(nextInsertionIndex) = bejeweledBoard
        nextInsertionIndex = (nextInsertionIndex + 1) Mod 2
        Dim board1Ago As BejeweledBoard = boards(nextInsertionIndex)

        Dim differencesTo1Ago As Integer = GetDifference(bejeweledBoard, board1Ago)
        Dim differencesTo2Ago As Integer = GetDifference(bejeweledBoard, board2Ago)
        Return differencesTo1Ago > 0 And differencesTo2Ago < 3
    End Function

    Friend Shared Function GetDifference(bejeweledBoard1 As BejeweledBoard, bejeweledBoard2 As BejeweledBoard) As Integer
        Dim differences As Integer = 0
        If bejeweledBoard1 IsNot Nothing And bejeweledBoard2 IsNot Nothing Then
            For x = 0 To TileCount - 1
                For y = 0 To TileCount - 1
                    If bejeweledBoard1.GetTile(x, y).TileCode <> bejeweledBoard2.GetTile(x, y).TileCode Then
                        differences += 1
                    End If
                Next
            Next
        End If
        Return differences
    End Function
End Class
