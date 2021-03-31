Imports System.Data.SQLite
Module BSGetData
    Public BS_Con As SQLiteConnection
    Public BS_Con_cs As String = "Provider=System.Data.SQLite;Data Source=" & BSLoc & ";Pooling=False;Max Pool Size=100;"
    Public BS_Cmd As New SQLiteCommand(BS_SQL, BS_Con)
    Public BS_rdr As SQLiteDataReader


    Public BSLoc = My.Settings.BSLoc & "\basestation.sqb"
    Public BS_SQL As String = ""
    Public ToLogType As String
    Public LoggedTag As String
    ReadOnly SymbolCode As String = ""
    Public MDPO As String = ""



    Public Sub GetBSdata(ToLogHex As String, ToLogreg As String)
        Try
            BS_Con = New SQLiteConnection
            BS_Con.ConnectionString = "Provider=System.Data.SQLite;Data Source=" & BSLoc & "" & ";PRAGMA cache_size = -10000;"

        Catch SqliteConnection As Exception

        End Try

        'Open connection to BaseStation

        Dim daBSCommand = New SQLiteCommand
        Try
            If BS_Con.State = ConnectionState.Open Then BS_Con.Close()
            If BS_Con.State = ConnectionState.Closed Then BS_Con.Open()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Basestation Connection Error")
        End Try
        'BS_Con.Open()

        PACModeSLogging.Timer1.Stop()
        ToLogreg = PACModeSLogging.ComboBox1.SelectedItem.ToString
        ToLogHex = ToLogreg.Substring(Math.Max(0, (ToLogreg.Length - 6)))
        ToLogreg = ToLogreg.Remove(ToLogreg.Length - 9)


        ' **** Test Record ****
        'BSstr = "Select UserTag, UserString1 from Aircraft WHERE Modes = 'AE119C'"
        ' ******


        BS_SQL = "Select UserTag, UserString1 from Aircraft WHERE Modes = " & Chr(34) & ToLogHex & Chr(34) & Chr(59)

        BS_Cmd = New SQLiteCommand(BS_SQL, BS_Con)
        BS_rdr = BS_Cmd.ExecuteReader()
        BS_rdr.Read()
        ToLogType = BS_rdr(0)
        If IsDBNull(BS_rdr(1)) Then
            LoggedTag = "LOG"
        Else
            'SymbolCode = BS_rdr(1)
            LoggedTag = "LOG" & BS_rdr(1)
        End If
        BS_rdr.Close()
        BS_Con.Close()
        tologdate = DateAndTime.Now.ToShortDateString
        Where = My.Settings.Location
        If ToLogType.Contains("RQ") Then
            MDPO = "M"
        ElseIf ToLogType.Contains("Ps") Then
            MDPO = "P"
        End If

        ' **** Test Record ****
        'MDPO = "M"
        ' ******#

        GetGFIAdata(MDPO, ToLogreg, ToLogHex)


    End Sub
End Module
