Imports System.Data.OleDb
Module GFIAGetData

    Public Logged_con As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & log_dbname & "")
    Public logged_SQL As String = "Select ID, registration from logllp;"
    ReadOnly log_dbname As String = "C:\ModeS\logged.mdb"
    Dim logged_cmd As New OleDbCommand
    Public Tologid As String
    Public ToLogMil As Integer


    Public Sub GetGFIAdata(MPDO As String, ToLogReg As String, ToLogHex As String)

        Try
            Logged_con = New OleDbConnection
            Logged_con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & log_dbname & ""

        Catch OledbConnection As Exception
        End Try


        ToLogReg = PACModeSLogging.ComboBox1.SelectedItem.ToString
        ToLogHex = ToLogReg.Substring(Math.Max(0, (ToLogReg.Length - 6)))
        ToLogReg = ToLogReg.Remove(ToLogReg.Length - 9)

        Try
            If Logged_con.State = ConnectionState.Open Then Logged_con.Close()
            If Logged_con.State = ConnectionState.Closed Then Logged_con.Open()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Access Connection Error")
        End Try

        'Try
        '    Logged_con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & log_dbname & ""
        'Catch OledbConnection As Exception
        'End Try

        logged_SQL = "SELECT ID, Hex, FKcmxo FROM tbldataset where Registration ="
        logged_SQL = logged_SQL & Chr(34) & ToLogReg & Chr(34) & " And Hex ="
        logged_SQL = logged_SQL & Chr(34) & ToLogHex & Chr(34) & Chr(59)

        '**** Test Record ****
        'logged_SQL = "SELECT ID, Hex, FKcmxo FROM tbldataset where Registration = '03-3119' And Hex = 'AE119C'"
        'ToLogReg = "03-3119"
        '******

        '**** Test Record ****
        'logged_SQL = "SELECT ID, Hex, FKcmxo FROM tbldataset where Registration = 'LXN90459' And Hex = '4D03D0'"
        'ToLogReg = "LXN90459"
        '******

        logged_cmd = New OleDbCommand(logged_SQL, Logged_con)
        Dim Logged_rdr As OleDbDataReader = logged_cmd.ExecuteReader()
        Logged_rdr.Read()
        If Logged_rdr.HasRows = False Then
            Dim response As DialogResult
            response = MsgBox("The registration you are trying to log (" & ToLogReg & ") does not match the one in GFIA - do you wish to continue?", vbYesNo)
            If response = DialogResult.Yes Then
                Logged_rdr.Close()

                Dim LogOutstanding As New LogOutstanding
                LogOutstanding.ShowDialog()
                UpdateBS(ToLogHex, ToLogReg)
                Exit Sub
            ElseIf response = DialogResult.No Then
                Exit Sub
            End If
        ElseIf Logged_rdr.HasRows = True Then
            Tologid = Logged_rdr(0)
            ToLogHex = Logged_rdr(1)
            ToLogMil = Logged_rdr(2)
            Logged_rdr.Close()
            UpdateGFIA(ToLogReg, ToLogHex, Tologid, ToLogMil, MDPO)
        End If
        Logged_con.Close()
        Logged_con.Dispose()

    End Sub
End Module
