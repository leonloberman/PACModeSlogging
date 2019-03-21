Imports System.Data.OleDb


Module SQLforAccess
    Public Sub AccessSQL(SQL As String, DBName As String)
        Dim con As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBName & "")
        con = New OleDbConnection
        con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBName & ""
        Dim cmd2 As New OleDb.OleDbCommand(SQL, con)

        Try
            'Write logfile
            'Using writer As StreamWriter = New StreamWriter(logfile, True)
            '    writer.WriteLine(Label1.Text & "-" & Date.Now)
            'End Using
            'SQL = "DELETE allhex.* FROM allhex;"
            con.Open()
            cmd2.ExecuteNonQuery()
            con.Dispose()
            con.Close()

        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try
    End Sub
End Module