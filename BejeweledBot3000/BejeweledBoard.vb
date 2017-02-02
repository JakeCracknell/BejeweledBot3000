Imports BejeweledBot3000

Public Class BejeweledBoard
    Private ReadOnly tileCount As Integer
    Private squares As Integer(,)

    Public Sub New(tileCount As Integer)
        Me.tileCount = tileCount
        ReDim squares(tileCount - 1, tileCount - 1)
    End Sub

    Private Sub New(squares As Integer(,))
        Me.New(squares.GetLength(0))
        Array.Copy(squares, Me.squares, squares.Length)
    End Sub

    Friend Sub SetTile(x As Integer, y As Integer, tileCode As Integer)
        squares(x, y) = tileCode
    End Sub

    Private Sub SwapTiles(x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer)
        Dim tmp As Integer = squares(x1, y1)
        squares(x1, y1) = squares(x2, y2)
        squares(x2, y2) = tmp
    End Sub

    Private Function TrySwapTiles(x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer) As Boolean
        For Each index In {x1, y1, x2, y2}
            If index < 0 Or index >= tileCount Then
                Return False
            End If
        Next
        SwapTiles(x1, y1, x2, y2)
        Return True
    End Function

    Function FindMove() As BejeweledMove
        Dim possibleMoves As New List(Of BejeweledMove)
        For x = 0 To tileCount - 1
            For y = 0 To tileCount - 1
                For Each direction In [Enum].GetValues(GetType(ArrowDirection))
                    possibleMoves.Add(New BejeweledMove(x, y, direction))
                Next
            Next
        Next
        Console.WriteLine(String.Join(vbNewLine, possibleMoves.OrderByDescending(Function(m) GetScoreForMove(m))))
        Return possibleMoves.OrderByDescending(Function(m) GetScoreForMove(m)).First
    End Function

    Function IsValidBoard() As Boolean
        Return GetScore() <= 2
    End Function

    Function GetUniqueTileCount() As Integer
        Dim codes As New HashSet(Of Integer)
        For x = 0 To tileCount - 1
            For y = 0 To tileCount - 1
                codes.Add(squares(x, y))
            Next
        Next
        Return codes.Count
    End Function

    Sub Normalise()
        Dim nextNorm As Integer = 0
        Dim dic As New Dictionary(Of Integer, Integer)
        For x = 0 To tileCount - 1
            For y = 0 To tileCount - 1
                Dim preNorm As Integer = squares(x, y)
                If Not dic.ContainsKey(preNorm) Then
                    dic.Add(preNorm, nextNorm)
                    nextNorm += 1
                End If
                squares(x, y) = dic(preNorm)
            Next
        Next
    End Sub

    Friend Function GetTile(x As Integer, y As Integer) As Integer
        Return squares(x, y)
    End Function

    Private Function GetScoreForMove(move As BejeweledMove) As Object
        Dim cloneBoard As New BejeweledBoard(squares)
        Return cloneBoard.PerformMoveAndGetScore(move)
    End Function

    Private Function PerformMoveAndGetScore(move As BejeweledMove) As Integer
        If PerformMoveAndGetValidity(move) Then
            Return GetScore()
        Else
            Return -1
        End If
    End Function

    Public Function GetScore() As Integer
        Dim score As Integer = 0
        For x = 0 To tileCount - 1
            Dim currentMatchingCode As Integer = Nothing
            Dim matches As Integer = 1
            For y = 0 To tileCount - 1
                If currentMatchingCode = squares(x, y) Then
                    matches += 1
                Else
                    currentMatchingCode = squares(x, y)
                    matches = 1
                End If
                score = Math.Max(matches, score)
            Next
        Next
        For y = 0 To tileCount - 1
            Dim currentMatchingCode As Integer = Nothing
            Dim matches As Integer = 1
            For x = 0 To tileCount - 1
                If currentMatchingCode = squares(x, y) Then
                    matches += 1
                Else
                    currentMatchingCode = squares(x, y)
                    matches = 1
                End If
                score = Math.Max(matches, score)
            Next
        Next
        Return score
    End Function

    Private Function PerformMoveAndGetValidity(move As BejeweledMove) As Integer
        Select Case move.Direction
            Case ArrowDirection.Down
                Return TrySwapTiles(move.X, move.Y, move.X, move.Y - 1)
            Case ArrowDirection.Up
                Return TrySwapTiles(move.X, move.Y, move.X, move.Y + 1)
            Case ArrowDirection.Left
                Return TrySwapTiles(move.X, move.Y, move.X - 1, move.Y)
            Case Else 'ArrowDirection.Right
                Return TrySwapTiles(move.X, move.Y, move.X + 1, move.Y)
        End Select
    End Function
End Class
