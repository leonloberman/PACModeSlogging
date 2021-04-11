Imports System.Data.OleDb
Imports System.Data.SQLite
Imports AutoUpdaterDotNET

Public Class PACModeSLogging

    'Declare the variables
    Dim drag As Boolean

    Dim mousex As Integer
    Dim mousey As Integer

    Public MyObject As Object
    Dim UT As String
    Dim i As Int16 = 0
    Dim PPHex As String
    Dim PPInt As String
    ReadOnly TempStr(3) As String
    ReadOnly RQ As Int16 = 0
    ReadOnly Interested As Int16 = 0
    ReadOnly PPType As String
    ReadOnly PPCallsign As String
    Dim PPAll As String

    Dim response As DialogResult

    Dim Reg As String
    Dim ListRec As String

    'Loggings definitions


    Dim Sound As New System.Media.SoundPlayer()

    Public ToLogReg As String
    Public ToLogHex As String

    Dim LResponse As Integer

    Public ToLogRec As Integer

    Dim cmd2 As New OleDbCommand
    Dim cmd3 As New OleDbCommand
    Dim logged_cmd As New OleDbCommand
    Dim logged_cmd2 As New OleDbCommand

    ReadOnly log_dbname As String = "C:\ModeS\logged.mdb"
    Dim CheckBusy As Boolean = False

#Disable Warning IDE0044 ' Add readonly modifier
    Dim BS_Cmd As New SQLiteCommand(BS_SQL, BS_Con)
#Enable Warning IDE0044 ' Add readonly modifier

    ReadOnly oInput As String



    Private Sub Form1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseDown, MyBase.MouseClick
        If e.Button = Windows.Forms.MouseButtons.Left Then
            drag = True 'Sets the variable drag to true.
            mousex = Windows.Forms.Cursor.Position.X - Me.Left 'Sets variable mousex
            mousey = Windows.Forms.Cursor.Position.Y - Me.Top 'Sets variable mousey
        End If
    End Sub

    Private Sub Form1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseMove
        'If drag is set to true then move the form accordingly.
        If drag Then
            Me.Top = Windows.Forms.Cursor.Position.Y - mousey
            Me.Left = Windows.Forms.Cursor.Position.X - mousex
        End If
    End Sub

    Private Sub Form1_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseUp
        drag = False 'Sets drag to false, so the form does not move according to the code in MouseMove
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

#Disable Warning BC42025 ' Access of shared member, constant member, enum member or nested type through an instance
        If My.Settings.Default.UpgradeRequired Then
#Enable Warning BC42025 ' Access of shared member, constant member, enum member or nested type through an instance
#Disable Warning BC42025 ' Access of shared member, constant member, enum member or nested type through an instance
            My.Settings.Default.Upgrade()
#Enable Warning BC42025 ' Access of shared member, constant member, enum member or nested type through an instance
#Disable Warning BC42025 ' Access of shared member, constant member, enum member or nested type through an instance
            My.Settings.Default.UpgradeRequired = False
#Enable Warning BC42025 ' Access of shared member, constant member, enum member or nested type through an instance
#Disable Warning BC42025 ' Access of shared member, constant member, enum member or nested type through an instance
            My.Settings.Default.Save()
#Enable Warning BC42025 ' Access of shared member, constant member, enum member or nested type through an instance
        End If

        If My.Computer.Network.IsAvailable Then
            If My.Computer.Network.Ping("www.google.com") Then
                'Application Upgrade Check
                'Dim BasicAuthentication As BasicAuthentication = New BasicAuthentication("pacmodes2020", "FkNrELRx")
                Dim BasicAuthentication As BasicAuthentication = New BasicAuthentication("pad", "Blackmrs99")
                AutoUpdater.BasicAuthXML = BasicAuthentication
                AutoUpdater.ReportErrors = True
                AutoUpdater.ShowSkipButton = False
                'AutoUpdater.Mandatory = True
                'AutoUpdater.Synchronous = True
                AutoUpdater.Start("https://www.gfiapac.org/ModeSVersions/PACModeSLoggingVersion.xml")
                'AutoUpdater.Start("https://www.gfiapac.org/ModeSVersions/PACModeSTestLoggingVersion.xml")

                AddHandler AutoUpdater.CheckForUpdateEvent, AddressOf AutoUpdaterOnCheckForUpdateEvent
            End If
        Else
            'MsgBox("Computer is not connected to the internet.")
        End If

        If My.Settings.Location = "<enter your location for logging>" Then
            Config.Show()
        End If
        If My.Settings.Sounds Then
            Button2.BackgroundImage = My.Resources.Sound_on
            Button2.Tag = "Sound"
        Else
            Button2.BackgroundImage = My.Resources.Sound_off
            Button2.Tag = "NoSound"
        End If

        If My.Settings.Autostart = True Then
            Button3.Visible = False
            Button1.Visible = True

            Try
                If Process.GetProcessesByName("PlanePlotter").Length = 0 Then
                    response = MsgBox("Waiting for PlanePlotter to start", vbOKCancel)
                    If response = DialogResult.Cancel Then
                        Close()
                        End
                    Else
                        Do Until Process.GetProcessesByName("PlanePlotter").Length > 0
                            System.Threading.Thread.Sleep(5000)
                            response = MsgBox("Waiting for PlanePlotter to start", vbOKCancel)
                            If response = DialogResult.Cancel Then
                                Close()
                                End
                            End If
                        Loop
                    End If
                End If

            Catch ex As Exception
                Timer1.Stop()
                Button1.Visible = True
                MsgBox("PlanePlotter Not Running", MsgBoxStyle.OkOnly, "PlanePlotter Check")

                Exit Sub
            End Try

            GetPPdata()
        End If
        'RunProcess()

    End Sub

    Private Sub AutoUpdaterOnCheckForUpdateEvent(ByVal args As UpdateInfoEventArgs)
        If args IsNot Nothing Then

            If args.IsUpdateAvailable Then

                AutoUpdater.ShowUpdateForm(args)
            Else
                'MessageBox.Show("There is no update available please try again later.", "No update available", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If My.Computer.Network.Ping("www.google.com") Then
                    'UpgradeCheck("C:\ModeS\ICAOCodes.mdb")
                Else
                    MsgBox("Computer is not connected to the internet.")
                End If
            End If
        Else
            MessageBox.Show("There is a problem reaching update server please check your internet connection and try again later.", "Update check failed", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Timer1.Stop()
        Button1.Visible = False
        Button3.Visible = True

    End Sub

    Public Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If My.Settings.BSLoc = " " Or My.Settings.Location = "<enter your location for logging>" Then
            MsgBox("Please configure the application parameters for basestation.sqb and your logging location", vbExclamation, "Settings Check")
            Dim MyConfig As New Config
            MyConfig.Show()
            Button3.Visible = True
        End If

        Button3.Visible = False
        Button1.Visible = True
        If My.Settings.AlwaysOnTop Then
            Me.TopMost = True
        Else
            Me.TopMost = False
        End If
        Me.StartPosition = FormStartPosition.CenterScreen
        'Me.Show()
        GetPPdata()

    End Sub

    Public Sub GetPPdata()

Start:
        MyObject = GetObject(, "PlanePlotter.Document")

        Dim RunMode As String = "Live"

        Timer1.Interval = CInt(My.Settings.SampleRate * 1000)
        Timer1.Start()
        Dim cancel As Boolean = False

        i = 0

        Try
            If RunMode = "Live" Then
                While i < MyObject.GetallPlaneCount()
                    PPHex = String.Empty
                    Reg = String.Empty
                    ListRec = String.Empty
                    PPAll = String.Empty
                    PPAll = MyObject.getallPlaneData(i, 98)
                    Dim PPbits() As String = Split(PPAll, ";")
                    PPHex = PPbits(0)
                    Reg = PPbits(1)
                    PPInt = PPbits(21)
                    UT = PPbits(22)
                    If Reg = "<gnd>" Or Reg = "<ground>" Or UT.Contains("LOG") Then GoTo EmptyStep

                    ListRec = Reg + " - " + PPHex

                    If My.Settings.InterestedButton = False Then
                        If UT.Contains("RQ") Then
                            If ComboBox1.Items.IndexOf(ListRec) = -1 Then
                                ComboBox1.Items.Add(ListRec)
                                If Button2.Tag = "Sound" Then
                                    Sound = New System.Media.SoundPlayer(My.Resources.RQ)
                                    Sound.Play()
                                End If
                            End If
                        ElseIf UT.Contains("Ps") Then
                            If ComboBox1.Items.IndexOf(ListRec) = -1 Then
                                ComboBox1.Items.Add(ListRec)
                                If Button2.Tag = "Sound" Then
                                    Sound = New System.Media.SoundPlayer(My.Resources.Ps)
                                    Sound.Play()
                                End If
                            End If
                            'ElseIf UT = "new" Then
                            '    Try
                            '        If Logged_con.State = ConnectionState.Open Then Logged_con.Close()
                            '        If Logged_con.State = ConnectionState.Closed Then Logged_con.Open()
                            '    Catch ex As Exception
                            '        MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Access Connection Error")
                            '    End Try

                            '    'Logged_con.Open()
                            '    logged_SQL = "SELECT * From Unknowns where ModeS ="
                            '    logged_SQL = logged_SQL & Chr(34) & PPHex & Chr(34) & Chr(59)
                            '    logged_cmd = New OleDbCommand(logged_SQL, Logged_con)
                            '    Dim logged_rdr As OleDbDataReader = logged_cmd.ExecuteReader()
                            '    logged_rdr.Read()
                            '    If logged_rdr.HasRows = False Then
                            '        'Write hex to Unknowns table
                            '        logged_SQL = "INSERT INTO Unknowns Values (" & Chr(34) & PPHex & Chr(34) & ", " & Chr(34) & Reg & Chr(34) & ", " & Chr(34) & PPCallsign & Chr(34) & ", now());"
                            '        logged_cmd = New OleDbCommand(logged_SQL, Logged_con)
                            '        logged_cmd.ExecuteNonQuery()
                            '    End If
                            '    logged_rdr.Close()
                            '    logged_cmd.Dispose()
                            '    Logged_con.Close()
                            'End If
                        End If
                        PPHex = String.Empty
                        Reg = String.Empty
                        ListRec = String.Empty
                    ElseIf My.Settings.InterestedButton = True Then
                        If PPInt = 1 Then
                            If ComboBox1.Items.IndexOf(ListRec) = -1 Then
                                ComboBox1.Items.Add(ListRec)
                                If Button2.Tag = "Sound" Then
                                End If
                            End If
                            'ElseIf UT = "new" Then
                            '    Try
                            '        If Logged_con.State = ConnectionState.Open Then Logged_con.Close()
                            '        If Logged_con.State = ConnectionState.Closed Then Logged_con.Open()
                            '    Catch ex As Exception
                            '        MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Connection Error")
                            '    End Try
                            '    'Logged_con.Open()
                            '    logged_SQL = "SELECT * From Unknowns where ModeS ="
                            '    logged_SQL = logged_SQL & Chr(34) & PPHex & Chr(34) & Chr(59)
                            '    logged_cmd = New OleDbCommand(logged_SQL, Logged_con)
                            '    Dim logged_rdr As OleDbDataReader = logged_cmd.ExecuteReader()
                            '    logged_rdr.Read()
                            '    If logged_rdr.HasRows = False Then
                            '        'Write hex to Unknowns table
                            '        logged_SQL = "INSERT INTO Unknowns Values (" & Chr(34) & PPHex & Chr(34) & ", " & Chr(34) & Reg & Chr(34) & ", now());"
                            '        logged_cmd = New OleDbCommand(logged_SQL, Logged_con)
                            '        logged_cmd.ExecuteNonQuery()
                            '    End If
                            '    logged_rdr.Close()
                            '    logged_cmd.Dispose()
                            '    Logged_con.Close()
                            'End If
                        End If
                    End If

                    PPHex = String.Empty
                    Reg = String.Empty
                    ListRec = String.Empty

EmptyStep:

                    PPHex = String.Empty
                    Reg = String.Empty
                    ListRec = String.Empty
                    i += 1
                End While
            Else
                '**** Test Data *****
                Reg = "LXN9059"
                PPHex = "4D03D0"
                ListRec = Reg + " - " + PPHex
                ComboBox1.Items.Add(ListRec)
                ''**** Test Data *****
            End If
        Catch ex As Exception
            Timer1.Stop()
                Button1.Visible = True
            MsgBox(ex.ToString)

        End Try



    End Sub


    Public Sub Combobox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged

        Timer1.Stop()

        ToLogRec = ComboBox1.SelectedIndex()

        'GetBSdata(ToLogHex, ToLogReg)

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

        Timer1.Stop()
        ToLogReg = ComboBox1.SelectedItem.ToString
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
        ToLogdate = DateAndTime.Now.ToShortDateString
        Where = My.Settings.Location
        If ToLogType.Contains("RQ") Then
            MDPO = "M"
        ElseIf ToLogType.Contains("Ps") Then
            MDPO = "P"
        End If

        ' **** Test Record ****
        'MDPO = "M"
        ' ******#

        'GetGFIAdata(MDPO, ToLogreg, ToLogHex)
        Try
            Logged_con = New OleDbConnection
            Logged_con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & log_dbname & ""

        Catch OledbConnection As Exception
        End Try


        ToLogReg = ComboBox1.SelectedItem.ToString
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

        Dim response As DialogResult
        logged_cmd = New OleDbCommand(logged_SQL, Logged_con)
        Dim Logged_rdr As OleDbDataReader = logged_cmd.ExecuteReader()
        Logged_rdr.Read()
        If Logged_rdr.HasRows = False Then
            logged_SQL = "SELECT * From logllp where ID = 3335 AND Registration = " & Chr(34) & ToLogReg & Chr(34)
            logged_cmd2 = New OleDbCommand(logged_SQL, Logged_con)
            Dim Logged_rdr2 As OleDbDataReader = logged_cmd2.ExecuteReader()
            Logged_rdr2.Read()
            If Logged_rdr2.HasRows = True Then
                response = MsgBox("You have already logged " & ToLogReg & " as an Outstanding record", vbOKOnly)
                Logged_rdr2.Close()
                UpdateBS(ToLogHex, ToLogReg)
                Exit Sub
            Else
                response = MsgBox("The registration you are trying to log (" & ToLogReg & ") does not match the one in GFIA - do you wish to continue?", vbYesNo)
                If response = DialogResult.Yes Then
                    Logged_rdr.Close()
                    Dim Log_Data As New Log_data
                    Log_Data.ShowDialog()
                    UpdateBS(ToLogHex, ToLogReg)
                    Exit Sub
                ElseIf response = DialogResult.No Then
                    Exit Sub
                End If
            End If

        ElseIf Logged_rdr.HasRows = True Then
            Tologid = Logged_rdr(0)
            ToLogHex = Logged_rdr(1)
            ToLogMil = Logged_rdr(2)
            Logged_rdr.Close()
            'UpdateGFIA(ToLogReg, ToLogHex, Tologid, ToLogMil, MDPO)
            Dim LogNote As New Notes
            'LogNote.ShowDialog()

            If LogNote.ShowDialog = DialogResult.OK Then

                LogText = LogNote.Passvalue

                'Get Model Prefix
                Dim Prefix As String = ""
                Dim Prefix_SQL As String = ""
                Prefix_SQL = "SELECT tbldataset.ID, tbldataset.FKsp2 as PrefixText " &
                "FROM tbldataset WHERE (((tbldataset.ID)=" & Tologid & "))"
                Dim cmd4 As OleDbCommand
                cmd4 = New OleDb.OleDbCommand(Prefix_SQL, Logged_con)
                Dim cmd4_rdr As OleDbDataReader
                cmd4_rdr = cmd4.ExecuteReader

                'Get Typnames
                Dim typename As String = ""
                Dim TypeName_SQL As String = ""
                TypeName_SQL = "SELECT tbldataset.ID, tblVariant.FKname AS TypeNameKey, tblNames.TypeName " &
                                " FROM (tblVariant INNER JOIN tbldataset ON tblVariant.FKvariant = tbldataset.FKvariant)" &
                                " INNER JOIN tblNames ON tblVariant.FKname = tblNames.FKname WHERE (((tbldataset.ID)=" & Tologid & "))"
                cmd3 = New OleDb.OleDbCommand(TypeName_SQL, Logged_con)
                Dim cmd3_rdr As OleDbDataReader
                cmd3_rdr = cmd3.ExecuteReader()

                If cmd3_rdr.HasRows = False Then

                    'Insert into logllp no typename no prefix
                    logged_SQL = "INSERT INTO logllp SELECT tbldataset.ID AS ID, ([tblManufacturer].[Builder]+' '+[tblmodel].[Model]+RIGHT([tblvariant].[variant], LEN([tblvariant].[variant]) - 1)" &
                                    "+' '+'['+[tbldataset].[cn]+']') AS Aircraft, tbldataset.Registration AS Registration, PRO_tbloperator.operator AS operator, " & Chr(34) & Where & Chr(34) & " AS [Where]," & Chr(34) & MDPO & Chr(34) & " AS MDPO, " &
                                    Chr(34) & ToLogdate & Chr(34) & " As [When], 0 As Lockk, " & Chr(34) & LogText & Chr(34) & " AS [Notes], " & Chr(34) & Source & Chr(34) & " AS [Source] " &
                                    " FROM tblVariant INNER JOIN ((tblmodel INNER JOIN tblManufacturer ON tblmodel.UID = tblManufacturer.UID) " &
                                    " INNER JOIN (PRO_tbloperator RIGHT JOIN tbldataset ON PRO_tbloperator.FKoperator = tbldataset.FKoperator) ON " &
                                    " (tblManufacturer.UID = tbldataset.UID) AND (tblmodel.FKmodel = tbldataset.FKmodel)) " &
                                    "ON tblVariant.FKvariant = tbldataset.FKvariant " & "WHERE (((tbldataset.ID)=" & Tologid & "));"
                    cmd2 = New OleDb.OleDbCommand(logged_SQL, Logged_con)
                    cmd2.ExecuteNonQuery()

                ElseIf cmd3_rdr.HasRows = True Then
                    cmd3_rdr.Close()
                    cmd3.ExecuteNonQuery()
                    cmd3_rdr = cmd3.ExecuteReader()
                    cmd3_rdr.Read()
                    typename = cmd3_rdr(2)

                    'Insert into logllp with typename no prefix
                    logged_SQL = "INSERT INTO logllp SELECT tbldataset.ID AS ID, ([tblManufacturer].[Builder]+' '+[tblmodel].[Model]+RIGHT([tblvariant].[variant], LEN([tblvariant].[variant]) - 1)" &
                                    "+' " & typename & "'+' ['+[tbldataset].[cn]+']') AS Aircraft, tbldataset.Registration AS Registration, PRO_tbloperator.operator AS operator, " & Chr(34) & Where & Chr(34) & " AS [Where]," & Chr(34) & MDPO & Chr(34) & " AS MDPO, " &
                                    Chr(34) & ToLogdate & Chr(34) & " As [When], 0 As Lockk, " & Chr(34) & LogText & Chr(34) & " AS [Notes], " & Chr(34) & Source & Chr(34) & " AS [Source] " &
                                    " FROM (tblVariant INNER JOIN ((tblmodel INNER JOIN tblManufacturer ON tblmodel.UID = tblManufacturer.UID) " &
                                    " INNER JOIN (PRO_tbloperator RIGHT JOIN tbldataset ON PRO_tbloperator.FKoperator = tbldataset.FKoperator) ON " &
                                    "(tblmodel.FKmodel = tbldataset.FKmodel) And (tblManufacturer.UID = tbldataset.UID)) " &
                                    "ON tblVariant.FKvariant = tbldataset.FKvariant) INNER JOIN tblNames ON tblVariant.FKname = tblNames.FKname " & "WHERE (((tbldataset.ID)=" & Tologid & "));"
                    cmd2 = New OleDb.OleDbCommand(logged_SQL, Logged_con)
                    cmd2.ExecuteNonQuery()

                ElseIf cmd3_rdr.HasRows = False And cmd4_rdr(1) = "" Then

                    'Insert into logllp no typename with no prefix
                    logged_SQL = "INSERT INTO logllp SELECT tbldataset.ID AS ID, ([tblManufacturer].[Builder]+' '+[tblmodel].[Model]+RIGHT([tblvariant].[variant], LEN([tblvariant].[variant]) - 1)" &
                                    "+' '+' ['+[tbldataset].[cn]+']') AS Aircraft, tbldataset.Registration AS Registration, PRO_tbloperator.operator AS operator, " & Chr(34) & Where & Chr(34) & " AS [Where]," & Chr(34) & MDPO & Chr(34) & " AS MDPO, " &
                                    Chr(34) & ToLogdate & Chr(34) & " As [When], 0 As Lockk, " & Chr(34) & LogText & Chr(34) & " AS [Notes], " & Chr(34) & Source & Chr(34) & " AS [Source] " &
                                    " FROM tblVariant INNER JOIN ((tblmodel INNER JOIN tblManufacturer ON tblmodel.UID = tblManufacturer.UID) " &
                                    " INNER JOIN (PRO_tbloperator RIGHT JOIN tbldataset ON PRO_tbloperator.FKoperator = tbldataset.FKoperator) ON " &
                                    " (tblManufacturer.UID = tbldataset.UID) AND (tblmodel.FKmodel = tbldataset.FKmodel)) " &
                                    "ON tblVariant.FKvariant = tbldataset.FKvariant " & "WHERE (((tbldataset.ID)=" & Tologid & "));"
                    cmd2 = New OleDb.OleDbCommand(logged_SQL, Logged_con)
                    cmd2.ExecuteNonQuery()

                ElseIf cmd3_rdr.HasRows = True And cmd4_rdr.HasRows = True Then
                    cmd3_rdr.Close()
                    cmd3.ExecuteNonQuery()
                    cmd3_rdr = cmd3.ExecuteReader()
                    cmd3_rdr.Read()
                    typename = cmd3_rdr(2)

                    cmd4_rdr.Close()
                    cmd4.ExecuteNonQuery()
                    cmd4_rdr = cmd4.ExecuteReader()
                    cmd4_rdr.Read()
                    Prefix = cmd4_rdr(1)

                    'Insert into logllp with typename with prefix
                    logged_SQL = "INSERT INTO logllp SELECT tbldataset.ID AS ID, ([tblManufacturer].[Builder]+' '+" & Chr(34) & Prefix & Chr(34) & "+[tblmodel].[Model]+RIGHT([tblvariant].[variant], LEN([tblvariant].[variant]) - 1)" &
                        "+' " & typename & "'+' ['+[tbldataset].[cn]+']') AS Aircraft, tbldataset.Registration AS Registration, PRO_tbloperator.operator AS operator, " & Chr(34) & Where & Chr(34) & " AS [Where]," & Chr(34) & MDPO & Chr(34) & " AS MDPO, " &
                        Chr(34) & ToLogdate & Chr(34) & " As [When], 0 As Lockk, " & Chr(34) & LogText & Chr(34) & " AS [Notes], " & Chr(34) & Source & Chr(34) & " AS [Source] " &
                        " FROM (tblVariant INNER JOIN ((tblmodel INNER JOIN tblManufacturer ON tblmodel.UID = tblManufacturer.UID) " &
                        " INNER JOIN (PRO_tbloperator RIGHT JOIN tbldataset ON PRO_tbloperator.FKoperator = tbldataset.FKoperator) ON " &
                        "(tblmodel.FKmodel = tbldataset.FKmodel) And (tblManufacturer.UID = tbldataset.UID)) " &
                        "ON tblVariant.FKvariant = tbldataset.FKvariant) INNER JOIN tblNames ON tblVariant.FKname = tblNames.FKname " & "WHERE (((tbldataset.ID)=" & Tologid & "));"
                    cmd2 = New OleDb.OleDbCommand(logged_SQL, Logged_con)
                    cmd2.ExecuteNonQuery()
                End If

                'Check for Fleetname and add
                If ToLogMil <> 502 Then
                    logged_SQL = " SELECT tbldataset.ID, LEFT([PRO-Marks].FleetName,50)" &
                            " FROM tbldataset LEFT JOIN [PRO-Marks] ON tbldataset.ID = [PRO-Marks].ID " &
                            " WHERE (tbldataset.ID)= " & Tologid
                    cmd2 = New OleDb.OleDbCommand(logged_SQL, Logged_con)
                    Dim cmd2_rdr As OleDbDataReader
                    cmd2_rdr = cmd2.ExecuteReader()

                    If cmd2_rdr.HasRows Then
                        cmd2_rdr.Close()
                        cmd2.ExecuteNonQuery()
                        cmd2_rdr = cmd2.ExecuteReader()
                        cmd2_rdr.Read()

                        If Not IsDBNull(cmd2_rdr(1)) Then
                            If (cmd2_rdr(1) <> " ") Then
                                ToLogacName = cmd2_rdr(1).Replace(Chr(34), "'")
                            End If
                            'Insert into logllp
                            logged_SQL = "UPDATE logLLp SET logllp.acName = " & Chr(34) & ToLogacName & Chr(34) & " WHERE ID = " & Tologid & ";"
                            cmd2 = New OleDb.OleDbCommand(logged_SQL, Logged_con)
                            cmd2.ExecuteNonQuery()
                            ToLogacName = String.Empty
                        Else
                            'Do Nothing
                        End If

                    End If

                End If

                If ToLogMil = 502 Then

                    'Check for mil records
                    logged_SQL = " SELECT tbldataset.ID, tblUnits.unit&' '&tblChildUnit.FamilyUnit AS logunit, [PRO-Marks].aCcode, LEFT([PRO-Marks].FleetName,50), tblChildUnit.FKtail" &
                                " FROM ((tbldataset LEFT JOIN [PRO-Marks] ON tbldataset.ID = [PRO-Marks].ID) LEFT JOIN tblUnits ON tbldataset.FKParent = tblUnits.FKUnits) LEFT JOIN tblChildUnit" &
                                " ON (tbldataset.FKChild = tblChildUnit.FKChild) AND (tbldataset.FKBaby = tblChildUnit.SubUnit)" &
                                " WHERE (((tbldataset.ID)=" & Tologid & ") AND (tbldataset.FKChild) > 0);"
                    cmd2 = New OleDb.OleDbCommand(logged_SQL, Logged_con)
                    Dim cmd2_rdr As OleDbDataReader
                    cmd2_rdr = cmd2.ExecuteReader()

                    If cmd2_rdr.HasRows Then
                        cmd2_rdr.Close()
                        cmd2.ExecuteNonQuery()
                        cmd2_rdr = cmd2.ExecuteReader()
                        cmd2_rdr.Read()

                        If Not IsDBNull(cmd2_rdr(1)) Then
                            If (cmd2_rdr(1) <> " ") Then
                                ToLogUnit = cmd2_rdr(1)
                            End If
                        Else
                            'Do Nothing
                        End If
                        If Not IsDBNull(cmd2_rdr(2)) Then
                            If (cmd2_rdr(2) <> " ") Then
                                ToLogaCcode = cmd2_rdr(2)
                            End If
                        Else
                            'Do nothing
                        End If
                        If Not IsDBNull(cmd2_rdr(3)) Then
                            If (cmd2_rdr(3) <> " ") Then

                                ToLogacName = cmd2_rdr(3).Replace(Chr(34), "'")
                                ToLogNotes = cmd2_rdr(3).Replace(Chr(34), "'")
                            Else
                                'Do Nothing
                            End If
                            If Not IsDBNull(cmd2_rdr(4)) Then
                                If (cmd2_rdr(4) <> " ") Then
                                    ToLogOther = cmd2_rdr(4).Replace(Chr(34), "'")
                                Else
                                    'Do nothing
                                End If
                            End If
                            'Insert into logllp
                            logged_SQL = "UPDATE logLLp SET logllp.logunit = " & Chr(34) & ToLogUnit & Chr(34) & ", logllp.[loga/c code] = " & Chr(34) & ToLogaCcode & Chr(34) &
                                    " , logllp.acName = " & Chr(34) & ToLogacName & Chr(34) & ", logllp.other = " & Chr(34) & ToLogOther & Chr(34) & " WHERE ID = " & Tologid & ";"
                            cmd2 = New OleDb.OleDbCommand(logged_SQL, Logged_con)
                            cmd2.ExecuteNonQuery()
                            ToLogacName = String.Empty
                            ToLogaCcode = String.Empty
                            ToLogUnit = String.Empty
                            ToLogOther = String.Empty
                        End If

                    End If
                End If
                'UpdateBS(ToLogHex, ToLogReg)

                Try
                    If BS_Con.State = ConnectionState.Open Then BS_Con.Close()
                    If BS_Con.State = ConnectionState.Closed Then BS_Con.Open()
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Basestation Connection Error")
                End Try
                'BS_Con.Open()
UpdBS2:         CheckBusy = False
                BS_Cmd = New SQLiteCommand(BS_SQL, BS_Con)
                BS_SQL = "SELECT * FROM sqlite_master"
                Try
                    BS_Cmd.ExecuteNonQuery()
                Catch SQLiteexception As Exception
                    If SQLiteErrorCode.Locked Then
                        CheckBusy = True
                    Else
                        CheckBusy = False
                    End If
                End Try
                If CheckBusy = True Then
                    GoTo UpdBS2
                Else
                    Try
                        Dim BS_SQL2 As String
                        If My.Settings.RemIntFlag = True Then
                            BS_SQL2 = "UPDATE AIRCRAFT SET UserTag = " & Chr(39) & LoggedTag & Chr(39) & ", LastModified = DATETIME(" & Chr(39) & "now" & Chr(39) & "," &
                            Chr(39) & "localtime" & Chr(39) & "), Interested = FALSE WHERE ModeS = " & Chr(39) & ToLogHex & Chr(39) & ";"
                        Else
                            BS_SQL2 = "UPDATE AIRCRAFT SET UserTag = " & Chr(39) & LoggedTag & Chr(39) & ", LastModified = DATETIME(" & Chr(39) & "now" & Chr(39) & "," &
                            Chr(39) & "localtime" & Chr(39) & ") WHERE ModeS = " & Chr(39) & ToLogHex & Chr(39) & ";"
                        End If

                        BS_Cmd = New SQLiteCommand(BS_SQL2, BS_Con)
                        BS_Cmd.ExecuteNonQuery()
                    Catch SQLITEexception As Exception
                        If SQLiteErrorCode.Locked Then
                            CheckBusy = True
                            GoTo UpdBS2
                        Else
                            CheckBusy = False
                        End If
                    End Try

                    BS_Con.Close()
                    BS_Con.Dispose()
                    MyObject.RefreshDatabaseInfo()
                    ComboBox1.Items.RemoveAt(ComboBox1.SelectedIndex)
                    Timer1.Start()
                    Timer1_Tick(Nothing, Nothing)
                End If

            Else
                RemoveHandler ComboBox1.SelectedIndexChanged, AddressOf Combobox1_SelectedIndexChanged
                ComboBox1.SelectedIndex = -1
                AddHandler ComboBox1.SelectedIndexChanged, AddressOf Combobox1_SelectedIndexChanged
                Timer1.Start()
                Timer1_Tick(Nothing, Nothing)

            End If
            Logged_con.Close()
            Logged_con.Dispose()

            If Process.GetProcessesByName("BaseStation.exe").Length >= 1 Then
                MyObject.RefreshDatabaseInfo()
            End If


        End If

    End Sub

    Public Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        MyObject.RefreshDatabaseInfo()
        'ComboBox1.Items.Clear()
        GetPPdata()
    End Sub

    Private Sub ComboBox1_selected() Handles ComboBox1.DropDown
        Timer1.Stop()
    End Sub

    Private Sub ComboBox1_deselected() Handles ComboBox1.DropDownClosed
        ComboBox1.Refresh()
        Timer1.Start()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Button2.Tag = "Sound" Then
            Button2.BackgroundImage = My.Resources.Sound_off
            Button2.Tag = "NoSound"
            My.Settings.Sounds = False
        Else
            Button2.BackgroundImage = My.Resources.Sound_on
            Button2.Tag = "Sound"
            My.Settings.Sounds = True
        End If
        My.Settings.Save()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Timer1.Stop()
        Button3.Visible = True
        Dim MyConfig As New Config
        MyConfig.Show()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Me.Close()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Timer1.Stop()
        ComboBox1.Items.Clear()
        Timer1.Start()
        Timer1_Tick(Nothing, Nothing)
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Timer1.Stop()
        Button3.Visible = True
        Dim LogData As New Log_data
        LogData.Show()
        ComboBox1.Items.Clear()
        Timer1.Start()
    End Sub

End Class