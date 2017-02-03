Public Class ColorBucket
    Public Colors As New List(Of Color)
    Private White As Color = Color.FromArgb(255, 255, 255)
    Private knownColors As New HashSet(Of Color) From {
        Color.FromArgb(255, 0, 49),
        Color.FromArgb(255, 0, 255),
        Color.FromArgb(0, 156, 255),
        Color.FromArgb(255, 255, 49),
        Color.FromArgb(255, 99, 0),
        Color.FromArgb(49, 255, 99),
        White
    }


    Public Function GetColorCode_MODE() As Integer
        Return Colors.GroupBy(Function(n) n).OrderByDescending(Function(g) g.Count).
                                Select(Function(g) g.Key).First.ToArgb
    End Function

    Public Function GetColorCode() As Integer
        Dim colorsByFrequency = Colors.GroupBy(Function(n) n).OrderByDescending(Function(g) g.Count).
                                       Select(Function(g) g.Key).Where(Function(c) knownColors.Contains(c))

        If colorsByFrequency.Count > 0 Then
            If colorsByFrequency.First = White Then
                If colorsByFrequency.Count = 1 Then
                    Return White.ToArgb
                Else
                    Return colorsByFrequency(1).ToArgb
                End If
            Else
                Return colorsByFrequency.First.ToArgb
            End If
        Else
            Return Color.Black.ToArgb
        End If
    End Function

End Class
