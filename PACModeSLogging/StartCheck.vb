Public Class StartCheck
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Exit Sub
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        End
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        PACModeSLogging.Timer1.Stop()
        Dim MyConfig As New Config
        MyConfig.Show()
    End Sub
End Class