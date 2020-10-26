Imports System.IO

Public Class Config

    Private Sub Config_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.CenterToParent()

        If My.Settings.InterestedButton = False Then
            RadioButton2.Checked = True
        ElseIf My.Settings.InterestedButton = True Then
            RadioButton1.Checked = True
        End If
        Label4.Text = "V" & My.Application.Info.Version.ToString
        TextBox1.Text = My.Settings.Location
        TextBox2.Text = My.Settings.BSLoc
        NumericUpDown1.Value = My.Settings.SampleRate
        If My.Settings.Sounds Then
            PACModeSLogging.Button2.Tag = "Sound"
        End If
        If My.Settings.AlwaysOnTop = True Then
            CheckBox1.CheckState = CheckState.Checked
        Else
            CheckBox1.CheckState = CheckState.Unchecked
        End If
        If My.Settings.RemIntFlag = True Then
            CheckBox2.CheckState = CheckState.Checked
        Else
            CheckBox2.CheckState = CheckState.Unchecked
        End If
    End Sub

    Private Sub RadioButton2_clicked(sender As Object, e As EventArgs) Handles RadioButton2.Click
        My.Settings.InterestedButton = False
        My.Settings.Save()
    End Sub

    Private Sub RadioButton1_clicked(sender As Object, e As EventArgs) Handles RadioButton1.Click
        My.Settings.InterestedButton = True
        My.Settings.Save()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        My.Settings.Location = TextBox1.Text
        My.Settings.BSLoc = TextBox2.Text
        My.Settings.SampleRate = NumericUpDown1.Value
        My.Settings.Save()
        Me.Close()
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.CheckState = CheckState.Checked Then
            My.Settings.AlwaysOnTop = True
        Else
            My.Settings.AlwaysOnTop = False
        End If
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim BSLocchk As String
        Using dialog As New FolderBrowserDialog
            If dialog.ShowDialog() <> DialogResult.OK Then Return
            'BSLoc = dialog.SelectedPath
            TextBox2.Text = dialog.SelectedPath
            BSLocchk = Path.Combine(dialog.SelectedPath, "basestation.sqb")

            If File.Exists(BSLocchk) = True Then
                My.Settings.BSLoc = TextBox2.Text.ToString
                Button1.Enabled = True
            Else
                Button1.Enabled = False
                MsgBox("There is no basestation.sqb file in this directory", MsgBoxStyle.OkOnly)
            End If
        End Using
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            My.Settings.RemIntFlag = True
        Else
            My.Settings.RemIntFlag = False
        End If
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged

    End Sub
End Class