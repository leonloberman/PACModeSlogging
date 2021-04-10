Imports System
Imports System.Data
Imports System.Data.OleDb

Public Class Log_data

    Private Sub No_reg_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim log_dbname As String = "C:\ModeS\logged.mdb"
        Dim dt As New DataTable("dt")
        Dim dt2 As New DataTable("dt2")
        Dim no_reg_SQL As String = "SELECT tblManufacturer.UID, tblManufacturer.Builder FROM tblManufacturer where Builder <> '-' OR Builder <> ''
                                    ORDER BY tblManufacturer.Builder;"
        'If PACModeSLogging.ComboBox1.SelectedIndex <> 0 Then
        '    PACModeSLogging.ToLogReg = PACModeSLogging.ComboBox1.SelectedItem.ToString
        '    PACModeSLogging.ToLogReg = PACModeSLogging.ToLogReg.Remove(PACModeSLogging.ToLogReg.Length - 9)

        '    TextBox1.Text = PACModeSLogging.ToLogReg
        'End If


        Using no_reg_con As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & log_dbname & "")
            Using no_reg_cmd As New OleDbCommand(no_reg_SQL, no_reg_con)
                Using no_reg_adapter As New OleDbDataAdapter(no_reg_cmd)
                    no_reg_con.Open()
                    no_reg_adapter.Fill(dt)
                    no_reg_con.Close()
                End Using
            End Using
        End Using


        ComboBox1.ValueMember = dt.Columns("UID").ToString
        ComboBox1.DisplayMember = dt.Columns("Builder").ToString
        ComboBox1.DataSource = dt

        no_reg_SQL = "SELECT PRO_tbloperator.fkoperator, PRO_tbloperator.Operator FROM PRO_tbloperator
                        WHERE PRO_tbloperator.AH = 'A' ORDER BY PRO_tbloperator.Operator;"
        Using no_reg_con As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & log_dbname & "")
            Using no_reg_cmd As New OleDbCommand(no_reg_SQL, no_reg_con)
                Using no_reg_adapter As New OleDbDataAdapter(no_reg_cmd)
                    no_reg_con.Open()
                    no_reg_adapter.Fill(dt2)
                    no_reg_con.Close()
                End Using
            End Using
        End Using

        ComboBox4.ValueMember = dt2.Columns("fkoperator").ToString
        ComboBox4.DisplayMember = dt2.Columns("Operator").ToString
        ComboBox4.DataSource = dt2

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

        Dim log_dbname As String = "C:\ModeS\logged.mdb"
        Dim dt2 As New DataTable("dt2")
        Dim UID = ComboBox1.SelectedValue.ToString
        Dim no_reg_SQL As String = "SELECT tblModel.fkmodel, tblModel.model FROM tblModel where tblModel.UID = " & UID & " ORDER BY tblModel.model;"
        Using no_reg_con As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & log_dbname & "")
            Using no_reg_cmd As New OleDbCommand(no_reg_SQL, no_reg_con)
                Using no_reg_adapter As New OleDbDataAdapter(no_reg_cmd)
                    no_reg_con.Open()
                    no_reg_adapter.Fill(dt2)
                    no_reg_con.Close()
                End Using
            End Using
        End Using


        ComboBox2.ValueMember = dt2.Columns("fkmodel").ToString
        ComboBox2.DisplayMember = dt2.Columns("Model").ToString
        ComboBox2.DataSource = dt2

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

        Dim log_dbname As String = "C:\ModeS\logged.mdb"
        Dim dt3 As New DataTable("dt3")
        Dim fkmodel = ComboBox2.SelectedValue.ToString
        Dim no_reg_SQL As String = "SELECT tblSeries.FKseries, tblVariant.Variant
                                    FROM tblmodel INNER JOIN (tblSeries INNER JOIN tblVariant ON tblSeries.FKseries = tblVariant.FKseries) ON tblmodel.FKmodel = tblSeries.FKModel
                                    WHERE tblSeries.FKmodel = " & fkmodel & " ORDER BY tblSeries.Series;"
        Using no_reg_con As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & log_dbname & "")
            Using no_reg_cmd As New OleDbCommand(no_reg_SQL, no_reg_con)
                Using no_reg_adapter As New OleDbDataAdapter(no_reg_cmd)
                    no_reg_con.Open()
                    no_reg_adapter.Fill(dt3)
                    no_reg_con.Close()
                End Using
            End Using
        End Using


        'ComboBox3.ValueMember = dt3.Columns("fkmodel").ToString
        ComboBox3.DisplayMember = dt3.Columns("Variant").ToString
        ComboBox3.DataSource = dt3


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim SQL As String
        Dim DBname As String = "C:\DataAir\MyLogs\privatelogs.mdb"
        Dim Aircraft As String
        Dim tologdate = DateAndTime.Now.ToShortDateString

        Aircraft = ComboBox1.Text & ComboBox2.Text & ComboBox3.Text & " [" & TextBox2.Text & "]"

        If TextBox1.TextLength = 0 Then
            MsgBox("You must enter the registration you are logging!!", vbExclamation, "Registration Check")
        End If

        If ComboBox1.SelectedValue.ToString = "-" Then
            MsgBox("You must select a manufacturer!!", vbExclamation, "Manufacturer Check")
        End If

        Dim logging_con As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBname & "")

        logging_con.Open()

        'Insert into logllp with Notes field

        SQL = "INSERT INTO logLLp ( ID, [when], Registration, Aircraft, Operator, [Where], flag, MDPO, LOCKK, Notes ) " &
                "VALUES (3335, " & Chr(34) & tologdate & Chr(34) & ",'" & TextBox1.Text & "','" & Aircraft & "','" &
                ComboBox4.Text & "','" & My.Settings.Location & "',True,'O', False, '" & TextBox3.Text & "'" & ");"



        Dim logging_cmd As New OleDbCommand(SQL, logging_con)
        logging_cmd.ExecuteNonQuery()
        logging_con.Close()
        Me.Close()
        'AccessSQL(SQL, DBname)


    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

End Class