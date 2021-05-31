Public Class PickReg
    Public TrueReg As Boolean
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TrueReg = True
        Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TrueReg = False
        Close()
    End Sub

    Private Sub PickReg_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Button1.Text = PACModeSLogging.TrueReg
        Button2.Text = PACModeSLogging.ToLogReg
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        TrueReg = True
        Me.Close()
    End Sub
End Class