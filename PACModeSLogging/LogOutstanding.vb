Imports System.Data.OleDb
Public Class LogOutstanding

    Private Sub LogOutstanding_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'LoggedDataSet.PRO_tbloperator' table. You can move, or remove it, as needed.
        Me.PRO_tbloperatorTableAdapter.Fill(Me.LoggedDataSet.PRO_tbloperator)
        'TODO: This line of code loads data into the 'LoggedDataSet.TypeList' table. You can move, or remove it, as needed.
        Me.TypeListTableAdapter.Fill(Me.LoggedDataSet.TypeList)
        PACModeSLogging.ToLogReg = PACModeSLogging.ComboBox1.SelectedItem.ToString
        PACModeSLogging.ToLogReg = PACModeSLogging.ToLogReg.Remove(PACModeSLogging.ToLogReg.Length - 9)

        TextBox1.Text = PACModeSLogging.ToLogReg
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim ToLogType As String
        Dim ToLogOperator As String
        Dim ToLogMilUnit As String
        Dim ToLogAcCode As String
        Dim ToLogName As String
        Dim ToLogMarks As String
        Dim cmd2 As New OleDbCommand


        ToLogType = ComboBox1.SelectedValue.ToString()
        ToLogOperator = ComboBox2.SelectedValue.ToString()
        ToLogMil = 501
        MDPO = "O"

        logged_SQL = "INSERT INTO logllp (ID, Registration, Aircraft, Operator, [Where], MDPO, [When], Lockk, Source) " &
        " Values(3335 , " & Chr(34) & PACModeSLogging.ToLogReg & Chr(34) & ", " & Chr(34) & ToLogType & Chr(34) & ", " & Chr(34) & ToLogOperator & Chr(34) &
        ", " & Chr(34) & Where & Chr(34) & ", " & Chr(34) & MDPO & Chr(34) &
        ", " & Chr(34) & ToLogdate & Chr(34) & ", 0, " & Chr(34) & Source & Chr(34) & ")"
        cmd2 = New OleDb.OleDbCommand(logged_SQL, Logged_con)
        cmd2.ExecuteNonQuery()

        If String.IsNullOrEmpty(TextBox4.Text) Then
            ' "Contains Empty value or Null Value" 
        Else
            ToLogMilUnit = TextBox4.Text
            logged_SQL = "UPDATE logllp SET logllp.logunit = " & Chr(34) & ToLogMilUnit & Chr(34) &
                        " WHERE ID = 3335 and Registration = " & Chr(34) & PACModeSLogging.ToLogReg & Chr(34)
            cmd2 = New OleDb.OleDbCommand(logged_SQL, Logged_con)
            cmd2.ExecuteNonQuery()
        End If

        If String.IsNullOrEmpty(TextBox6.Text) Then
            ' "Contains Empty value or Null Value" 
        Else
            ToLogAcCode = TextBox6.Text
            logged_SQL = "UPDATE logllp SET logllp.[loga/c code] = " & Chr(34) & ToLogAcCode & Chr(34) &
                        " WHERE ID = 3335 and Registration = " & Chr(34) & PACModeSLogging.ToLogReg & Chr(34)
            cmd2 = New OleDb.OleDbCommand(logged_SQL, Logged_con)
            cmd2.ExecuteNonQuery()
        End If

        If String.IsNullOrEmpty(TextBox7.Text) Then
            ' "Contains Empty value or Null Value" 
        Else
            ToLogName = TextBox7.Text
            logged_SQL = "UPDATE logllp SET logllp.acName = " & Chr(34) & ToLogName & Chr(34) &
                        " WHERE ID = 3335 and Registration = " & Chr(34) & PACModeSLogging.ToLogReg & Chr(34)
            cmd2 = New OleDb.OleDbCommand(logged_SQL, Logged_con)
            cmd2.ExecuteNonQuery()
        End If

        If String.IsNullOrEmpty(TextBox8.Text) Then
            ' "Contains Empty value or Null Value" 
        Else
            ToLogMarks = TextBox8.Text
            logged_SQL = "UPDATE logllp SET logllp.[other] = " & Chr(34) & ToLogMarks & Chr(34) &
                        " WHERE ID = 3335 and Registration = " & Chr(34) & PACModeSLogging.ToLogReg & Chr(34)
            cmd2 = New OleDb.OleDbCommand(logged_SQL, Logged_con)
            cmd2.ExecuteNonQuery()
        End If

        '****Test Data
        'ToLogMilUnit = "TEST"
        'ToLogSubUnit = "TEST"
        'ToLogAcCode = "TEST"
        'ToLogName = "TEST"
        'ToLogMarks = "TEST"
        '****Test Data

        Close()
        Exit Sub

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Close()
        PACModeSLogging.Timer1.Start()
        PACModeSLogging.Timer1_Tick(Nothing, Nothing)
        Exit Sub
    End Sub
End Class