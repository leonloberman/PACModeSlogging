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

                MyObject = GetObject(, "PlanePlotter.Document")
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


        Timer1.Interval = CInt(My.Settings.SampleRate * 1000)
        Timer1.Start()
        Dim cancel As Boolean = False

        i = 0
        Try
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

                '**** Test Data *****
                'Reg = "LXN9059"
                'PPHex = "4D03D0"
                'ListRec = Reg + " - " + PPHex
                'ComboBox1.Items.Add(ListRec)
                '**** Test Data *****


                PPHex = String.Empty
                Reg = String.Empty
            ListRec = String.Empty

EmptyStep:

            PPHex = String.Empty
            Reg = String.Empty
            ListRec = String.Empty
            i += 1
            End While
        Catch ex As Exception
            Timer1.Stop()
            Button1.Visible = True
            MsgBox(ex.ToString)
        End Try

    End Sub


    Public Sub Combobox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged

        Timer1.Stop()

        ToLogRec = ComboBox1.SelectedIndex()

        GetBSdata(ToLogHex, ToLogReg)

        If Process.GetProcessesByName("BaseStation.exe").Length >= 1 Then
            MyObject.RefreshDatabaseInfo()
        End If

    End Sub

    Public Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        MyObject.RefreshDatabaseInfo()
        ComboBox1.Items.Clear()
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