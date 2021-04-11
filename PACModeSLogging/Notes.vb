Public Class Notes

    Public Passvalue As String
    Public Canx As Integer = 0

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Canx = 1
        'Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'RaiseEvent Passvalue(TextBox1.Text)
        Passvalue = TextBox1.Text
        Me.Close()
    End Sub

    Private Sub Notes_Load() Handles MyBase.Load
        Dim ToLogReg As String = PACModeSLogging.ToLogReg
        Dim ToLogHex As String = PACModeSLogging.ToLogHex
        Label1.Text = "You are logging " & ToLogReg & " (ModeS code " & ToLogHex & ")"
    End Sub
End Class