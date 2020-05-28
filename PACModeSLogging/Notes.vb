Public Class Notes

    Public Passvalue As String
    Public Canx As Integer = 0

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Canx = 1
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'RaiseEvent Passvalue(TextBox1.Text)
        Passvalue = TextBox1.Text
        Me.Close()
    End Sub

    Private Sub Notes_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label1.Text = "You are logging " & PACModeSLogging.ToLogReg & " (ModeS code " & PACModeSLogging.ToLogHex & ")"
    End Sub
End Class