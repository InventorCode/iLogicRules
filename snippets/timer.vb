        Dim timer As New Timer("Testing")
        timer.StartTimer()

        'Do Stuff

        timer.StopTimer()
        timer.Print()



Public Class Timer

    Private _totalTimer As Stopwatch = New Stopwatch()
    Private _totalTs As TimeSpan
    Private _totalTime As String
    Public Message As String = ""

    Public Sub New()
    End Sub

    Public Sub New(value As String)
        me.Message = value
    End Sub

    Public Sub StartTimer()
            _totalTimer.Start()
    End Sub

    Public Sub StopTimer()
            _totalTimer.Stop()
            _totalTs = _totalTimer.Elapsed
            _totalTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}", _totalTs.Hours, _totalTs.Minutes, _totalTs.Seconds, _totalTs.Milliseconds / 10)
    End Sub

    Public Sub Print()
        MsgBox(me.Message & " Total Time: " & _totalTime.ToString)
    End Sub

    Public Function Value()
        return me.Message & " Total Time: " & _totalTime.ToString
    End Function

End Class