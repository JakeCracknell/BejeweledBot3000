Imports BejeweledBot3000

Public Class OscillationDetector
    Private moveListHashcodes(6) As Integer
    Private nextInsertionIndex As Byte = 0

    Function LogMovesAndReturnTrueIfOscillationDetected(moves As List(Of BejeweledMove))
        Dim hashcode As Integer = moves.Select(Function(m) m.GetHashCode).Sum
        Try
            Return hashcode <> 0 AndAlso moveListHashcodes.Where(Function(h) h = hashcode).Count > 2
        Finally
            moveListHashcodes(nextInsertionIndex) = hashcode
            nextInsertionIndex = (nextInsertionIndex + 1) Mod moveListHashcodes.Length
        End Try
    End Function

    Friend Shared Function GetDifferenceInBoards(bejeweledBoard1 As BejeweledBoard, bejeweledBoard2 As BejeweledBoard) As Integer
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
