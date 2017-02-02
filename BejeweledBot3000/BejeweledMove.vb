Public Class BejeweledMove
    Public Direction As ArrowDirection
    Public X As Integer
    Public Y As Integer
    Public Score As Integer

    Public Sub New(x As Integer, y As Integer, direction As ArrowDirection)
        Me.X = x
        Me.Y = y
        Me.Direction = direction
    End Sub

    Public Overrides Function ToString() As String
        Return X & "," & Y & "," & Direction.ToString
    End Function
End Class
