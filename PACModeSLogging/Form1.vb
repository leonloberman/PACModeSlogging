Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SQLite


Public Class Form1
    'Declare the variables
    Dim drag As Boolean
    Dim mousex As Integer
    Dim mousey As Integer

    Dim MyObject
    Dim UT As String
    Dim i As Int16 = 0
    Dim PPHex As String
    Dim PPInt As String
    Dim TempStr(3) As String
    Dim RQ As Int16 = 0
    Dim Interested As Int16 = 0
    Dim PPType As String
    Dim PPCallsign As String
    Dim PPAll As String
    Dim Sharers As String

    Dim Reg As String
    Dim ListRec As String


    'Loggings definitions
    Dim Logged_con As OleDbConnection

    Dim logged_SQL As String = "Select ID, registration from logllp;"
    Dim Logllp_dbname As String = "C:\DataAir\Mylogs\privatelogs.mdb"
    'Dim Logllp_dbname As String = "C:\ModeS\privatelogs.mdb"
    Dim log_dbname As String = "C:\ModeS\logged.mdb"

    Dim logged_cmd As New OleDbCommand(logged_SQL, Logged_con)
    Dim logged2_cmd As New OleDbCommand(logged_SQL, Logged_con)
    Dim logged2_SQL As String


    'GFIA database definitions
    Dim dtset_con As OleDbConnection
    Dim dtset_sql As String = "SELECT ID FROM tbldataset where hex ="
    Dim dtset As String = "C:\DataAir\dtset.mdb"

    Dim dtset_cmd As New OleDbCommand(dtset_sql, dtset_con)

    Dim Sound As New System.Media.SoundPlayer()

    Dim ToLogReg As String
    Dim ToLogHex As String
    Dim ToLogType As String
    Dim Tologid As String
    Dim tologdate As String
    Dim ToLogMil As Integer

    Dim MDPO As String = ""
    Dim Where As String
    Dim LResponse As Integer


    'BaseStation definitions
    Dim BS_Con As SQLiteConnection
    'Dim BSLoc As String = My.Settings.BSLoc + "/basestation.sqb"
    Dim BSLoc As String = "C:/ModeS/basestation.sqb"
    Dim BS_SQL As String = "Update Aircraft set UserTag = " & """LOG""" & " WHERE Modes = "

    Dim BS_Cmd As New SQLiteCommand(BS_SQL, BS_Con)

    Dim oInput As String
    Dim BS_id As String
    Dim BS_Hex As String
    Dim BS_CtryName As String
    Dim BS_Reg As String
    Dim BS_ICAOTypeCode As String
    Dim BS_CN As String
    Dim BS_FKCMXO As String
    Dim BS_Operator As String
    Dim BS_OperatorFlagCode As String
    Dim BS_UserTagv3 As String
    Dim BS_UserTagV1 As String
    Dim BS_Interested As String
    Dim BS_UserInt1 As String


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

        'If My.Settings.Location.Length = 0 Then
        'Me.Visible = False
        'MyConfig.Show()
        'End If
        If My.Settings.Sounds Then
            Button2.BackgroundImage = My.Resources.Sound_on
            Button2.Tag = "Sound"
        Else
            Button2.BackgroundImage = My.Resources.Sound_off
            Button2.Tag = "NoSound"
        End If

        RunProcess()

    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Timer1.Stop()
        Button1.Visible = False
        Button3.Visible = True

    End Sub



    Public Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        If My.Settings.Location.Length = 0 Then
            MsgBox("You must enter the location you are logging at!!", vbExclamation, "Location Check")
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

        RunProcess()

    End Sub


    Public Sub RunProcess()

        Dim BS_Con_cs As String = "Provider=System.Data.SQLite;Data Source=" & BSLoc & ";Pooling=False;Max Pool Size=100;"
        Dim BS_Con As New SQLiteConnection(BS_Con_cs)
        Dim BS_Cmd As New SQLiteCommand(BS_Con)
        Dim BS_rdr As SQLiteDataReader

        Logged_con = New OleDbConnection
        Logged_con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & log_dbname & ""
        dtset_con = New OleDbConnection
        dtset_con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dtset & ""

Start:

        Try
            MyObject = GetObject(, "PlanePlotter.Document")
        Catch ex As Exception
            Timer1.Stop()
            Button1.Visible = True
            MsgBox("PlanePlotter Not Running", MsgBoxStyle.OkOnly, "PlanePlotter Check")

            Exit Sub
        End Try

        Timer1.Interval = My.Settings.SampleRate * 1000
        Timer1.Start()
        Dim cancel As Boolean = False

        i = 0
        Try
            While i < MyObject.GetallPlaneCount()
                i = i + 1
                PPHex = String.Empty
                Reg = String.Empty
                ListRec = String.Empty
                PPAll = String.Empty
                PPAll = MyObject.getallPlaneData(i, 98)
                Dim PPbits() As String = Split(PPAll, ";")
                PPHex = PPbits(0)
                Reg = PPbits(1)
                PPInt = PPbits(21)


                Dim CheckBusy As Boolean = False
GetPPHex:       CheckBusy = False

                If BS_Con.State = ConnectionState.Closed Then BS_Con.Open()

                Dim BSstr As String = "Select * from Aircraft WHERE Modes = " & Chr(34) & PPHex & Chr(34) & ";"

                BS_Cmd.Connection = BS_Con
                BS_Cmd.CommandText = BSstr
                BS_rdr = BS_Cmd.ExecuteReader()
                BS_rdr.Read()
                BS_rdr.Close()
                If CInt(BS_Cmd.ExecuteScalar) = 0 Then

                    'Logged_con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source=" & log_dbname & ""
                    Dim logged_cmd As New OleDbCommand


                    dtset_con = New OleDbConnection
                        dtset_con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dtset & ""


                    'dtset_sql = "SELECT tbldataset.ID, tbldataset.Hex, tblCountry.CountryName, tbldataset.Registration, tblSeries.code, tbldataset.CN, tbldataset.FKCMXO, PRO_tbloperator.Operator,  Str([PRO_tbloperator].[FKoperator]) AS OperatorFlagCode
                    '   FROM PRO_tbloperator RIGHT JOIN (tblSeries INNER JOIN (tblCountry RIGHT JOIN tbldataset ON tblCountry.FKcountry = tbldataset.FKcountry) ON tblSeries.FKseries = tbldataset.FKseries) ON PRO_tbloperator.FKoperator = tbldataset.FKoperator
                    '   where hex = '" & PPHex & "'; "

                    logged_SQL = "Select tbldataset.ID, tbldataset.Hex, tblCountry.CountryName, tbldataset.Registration, tblSeries.code, tbldataset.CN, tbldataset.FKCMXO, PRO_tbloperator.Operator, Str([PRO_tbloperator].[FKoperator]) As OperatorFlagCode, PP_SymbolsByType.UserString1_v3, PP_SymbolsByType.UserString1_v1
                        FROM(PRO_tbloperator RIGHT JOIN (tblSeries INNER JOIN (tblCountry RIGHT JOIN tbldataset On tblCountry.FKcountry = tbldataset.FKcountry) On tblSeries.FKseries = tbldataset.FKseries) On PRO_tbloperator.FKoperator = tbldataset.FKoperator) INNER JOIN PP_SymbolsByType On tblSeries.code = PP_SymbolsByType.ICAOTypeCode
                        Where (tbldataset.[hex]) = '" & PPHex & "'; "

                    If Logged_con.State = ConnectionState.Closed Then Logged_con.Open()
                    logged_cmd = New OleDbCommand(logged_SQL, Logged_con)
                    Dim logged_rdr As OleDbDataReader = logged_cmd.ExecuteReader()
                    logged_rdr.Read()
                    If logged_rdr.HasRows = False Then Continue While

                    BS_id = logged_rdr(0)
                    BS_Hex = logged_rdr(1)
                    BS_CtryName = logged_rdr(2)
                    BS_Reg = logged_rdr(3)

                    If Not IsDBNull(logged_rdr.Item(4)) Then
                        BS_ICAOTypeCode = logged_rdr(4)
                    Else BS_ICAOTypeCode = "????"
                    End If

                    BS_CN = logged_rdr(5)
                    BS_FKCMXO = logged_rdr(6)
                    BS_Operator = logged_rdr(7)
                    BS_OperatorFlagCode = logged_rdr(8)
                    If Not IsDBNull(logged_rdr(9)) Then
                        BS_UserTagv3 = logged_rdr(9)
                    Else BS_UserTagv3 = ""
                    End If

                    If Not IsDBNull(logged_rdr(9)) Then
                        BS_UserInt1 = logged_rdr(9)
                    Else BS_UserInt1 = ""
                    End If

                    'BS_UserTagV1 = logged_rdr(10)

                    'BS_UserTagv3 = "RQ" & BS_UserTagv3

                    BS_OperatorFlagCode = "x" & LTrim([BS_OperatorFlagCode])

                    If BS_FKCMXO = "502" Then
                        BS_Interested = 1
                    Else BS_Interested = 0
                    End If

                    'logged2_SQL = "SELECT DISTINCT logLLp.Registration, tbldataset.ID, tbldataset.Hex
                    '                FROM (logLLp INNER JOIN tblOperatorHistory ON logLLp.ID = tblOperatorHistory.ID)
                    'INNER JOIN tbldataset ON (logLLp.ID = tbldataset.ID)
                    'Where (tbldataset.[hex]) = '" & PPHex & "'; "

                    logged2_SQL = "Select DISTINCT logLLp.Registration,  logLLp.ID As AircraftID, tbldataset.Hex
                      FROM ((logLLp INNER JOIN tblOperatorHistory On logLLp.ID = tblOperatorHistory.ID) INNER JOIN tbldataset On logLLp.ID = tbldataset.ID)
                                             WHERE logLLp.Registration In (Select tblOperatorHistory.previous from tblOperatorHistory) 
                        AND (tbldataset.[hex]) = '" & PPHex & "';"

                    logged2_cmd = New OleDbCommand(logged2_SQL, Logged_con)
                    Dim logged2_rdr As OleDbDataReader = logged2_cmd.ExecuteReader()
                    logged2_rdr.Read()
                    If logged2_rdr.HasRows = True Then
                        BS_UserTagv3 = "Ps" & BS_UserTagv3
                    End If

                    logged2_SQL = "Select DISTINCT logLLp.Registration,  logLLp.ID As AircraftID, tbldataset.Hex
                      FROM ((logLLp INNER JOIN tblOperatorHistory On logLLp.ID = tblOperatorHistory.ID) INNER JOIN tbldataset On logLLp.ID = tbldataset.ID)
                                             WHERE logLLp.Registration In (Select tblOperatorHistory.previous from tblOperatorHistory) 
                        AND (tbldataset.[hex]) = '" & PPHex & "';"

                    logged2_cmd = New OleDbCommand(logged2_SQL, Logged_con)
                    logged2_rdr.Read()
                    If logged2_rdr.HasRows = False Then
                        BS_UserTagv3 = "RQ" & BS_UserTagv3
                    End If

                    logged2_SQL = "Select DISTINCT logLLp.Registration,  logLLp.ID As AircraftID, tbldataset.Hex
                      FROM logLLp INNER JOIN tbldataset On logLLp.ID = tbldataset.ID
                        AND (tbldataset.[hex]) = '" & PPHex & "';"

                    'If Logged_con.State = ConnectionState.Closed Then Logged_con.Open()
                    logged2_cmd = New OleDbCommand(logged2_SQL, Logged_con)
                    logged2_rdr.Read()
                    If logged2_rdr.HasRows = False Then
                        BS_UserTagv3 = logged_rdr(9)
                    End If
                    logged_rdr.Close()
                    logged2_rdr.Close()

                    'Write to basestation.sqb
                    BS_SQL = "INSERT INTO Aircraft (RowID, AircraftID, ModeS, ModeSCountry, Registration, ICAOTypeCode, SerialNo, RegisteredOwners, OperatorFlagCode, UserTag, Interested, UserInt1, UserString1, FirstCreated, LastModified)" &
                     "Values (" & BS_id & ", " & BS_id & ", " & Chr(39) & BS_Hex & Chr(39) & ", " & Chr(39) & BS_CtryName & Chr(39) & ", " & Chr(39) & BS_Reg & Chr(39) & ", " &
                      Chr(39) & BS_ICAOTypeCode & Chr(39) & ", " & Chr(39) & BS_CN & Chr(39) & ", " & Chr(39) & BS_Operator & Chr(39) & ", " & Chr(39) & BS_OperatorFlagCode & Chr(39) & ", " &
                      Chr(39) & BS_UserTagv3 & Chr(39) & ", " & Chr(39) & BS_Interested & Chr(39) & ", " & Chr(39) & BS_FKCMXO & Chr(39) & ", " & Chr(39) & BS_UserInt1 & Chr(39) & ", " & "DateTime('now'), DateTime('now'));"
                    BS_Cmd = New SQLiteCommand(BS_SQL, BS_Con)

BS_execute1:        Try
                        BS_Cmd.ExecuteNonQuery()
                    Catch SQLiteexception As Exception
                        If SQLiteErrorCode.Locked Then
                                CheckBusy = True
                            Else
                                CheckBusy = False
                            End If
                        End Try
                        If CheckBusy = True Then
                        GoTo BS_execute1

                    End If
                    Else Continue While
                    End If


                If Reg = "<gnd>" Or Reg = "<ground>" Or UT = "Log" Then Continue While

                ListRec = Reg + " - " + PPHex

                If My.Settings.InterestedButton = False Then
                    If BS_UserTagv3.Contains("RQ") Then
                        If ComboBox1.Items.IndexOf(ListRec) = -1 Then
                            ComboBox1.Items.Add(ListRec)
                            If Button2.Tag = "Sound" Then
                                Sound = New System.Media.SoundPlayer(My.Resources.RQ)
                                Sound.Play()
                            End If
                        End If
                    ElseIf BS_UserTagv3.Contains("Ps") Then
                        If ComboBox1.Items.IndexOf(ListRec) = -1 Then
                            ComboBox1.Items.Add(ListRec)
                            If Button2.Tag = "Sound" Then
                                Sound = New System.Media.SoundPlayer(My.Resources.Ps)
                                Sound.Play()
                            End If
                        End If
                        'ElseIf BS_UserTagv3 = "new" Then
                        '    Logged_con.Open()
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
                    End If
                End If


                PPHex = String.Empty
            Reg = String.Empty
            ListRec = String.Empty
            If My.Settings.InterestedButton = True Then
                If PPInt = 1 Then
                    If ComboBox1.Items.IndexOf(ListRec) = -1 Then
                        ComboBox1.Items.Add(ListRec)
                        If Button2.Tag = "Sound" Then
                        End If
                    End If
                        'ElseIf BS_UserTagv3 = "new" Then
                        '    Logged_con.Open()
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
                    End If
            End If
            PPHex = String.Empty
            Reg = String.Empty
                ListRec = String.Empty




EmptyStep:

            PPHex = String.Empty
            Reg = String.Empty
            ListRec = String.Empty

            End While

            BS_Con.Close()
            BS_Con.Dispose()
        Catch ex As Exception
            Timer1.Stop()
            Button1.Visible = True
            MsgBox(ex.ToString)
        End Try

        Logged_con.Close()
        dtset_con.Close()


    End Sub


    Public Sub Combobox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        
        Logged_con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & log_dbname & ""
        Dim logged_cmd As New OleDbCommand


        dtset_con = New OleDbConnection
        dtset_con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dtset & ""
       
        Timer1.Stop()
        ToLogReg = ComboBox1.SelectedItem.ToString
        ToLogHex = ToLogReg.Substring(Math.Max(0, (ToLogReg.Length - 6)))
        ToLogReg = ToLogReg.Remove(ToLogReg.Length - 9)
        LResponse = MsgBox("Do you want to log " & ToLogReg & Chr(32) & "?", vbOKCancel)
        If LResponse = 1 Then
            If Logged_con.State = ConnectionState.Closed Then Logged_con.Open()
            If dtset_con.State = ConnectionState.Closed Then dtset_con.Open()
            dtset_sql = "SELECT ID, Hex, FKcmxo FROM tbldataset where Registration ="
            dtset_sql = dtset_sql & Chr(34) & ToLogReg & Chr(34) & " and Hex ="
            dtset_sql = dtset_sql & Chr(34) & ToLogHex & Chr(34) & Chr(59)
            dtset_cmd = New OleDbCommand(dtset_sql, dtset_con)
            Dim dtset_rdr As OleDbDataReader = dtset_cmd.ExecuteReader()
            dtset_rdr.Read()
            Tologid = dtset_rdr(0)
            ToLogHex = dtset_rdr(1)
            ToLogMil = dtset_rdr(2)
            dtset_rdr.Close()
            dtset_con.Close()
            dtset_con.Dispose()

            Dim BSstr As String = "Select UserTag from Aircraft WHERE Modes = " & Chr(34) & ToLogHex & Chr(34) & Chr(59)
            Dim BS_Con_cs As String = "Provider=System.Data.SQLite;Data Source=" & BSLoc & ";Pooling=False;Max Pool Size=100;"
            Dim BS_Con As New SQLiteConnection(BS_Con_cs)
            Dim BS_Cmd As New SQLiteCommand(BS_Con)
            Dim BS_rdr As SQLiteDataReader
            BS_Con.Open()

            BS_Cmd.Connection = BS_Con
            BS_Cmd.CommandText = BSstr
            BS_rdr = BS_Cmd.ExecuteReader()
            BS_rdr.Read()
            ToLogType = BS_rdr(0)
            BS_rdr.Close()
            BS_Con.Close()
            tologdate = DateAndTime.Now.ToShortDateString
            Where = My.Settings.Location
            If ToLogType.Contains("RQ") Then MDPO = "M"
            If ToLogType.Contains("Ps") Then MDPO = "P"
            'Insert No-Reg into logllp as O'


            'Insert into logllp
            logged_SQL = "INSERT INTO logllp SELECT tbldataset.ID AS ID, ([tblManufacturer].[Builder]+' '+[tblmodel].[Model]+RIGHT([tblvariant].[variant], LEN([tblvariant].[variant]) - 1)+' '" & _
                "+'['+[tbldataset].[cn]+']') AS Aircraft, tbldataset.Registration AS Registration, PRO_tbloperator.operator AS operator, " & Chr(34) & Where & Chr(34) & " AS [Where]," & Chr(34) & MDPO & Chr(34) & " AS MDPO, " & Chr(34) & tologdate & Chr(34) & " As [When], 0 As Lockk" & _
                " FROM PRO_tbloperator RIGHT JOIN (tblVariant INNER JOIN (tblmodel INNER JOIN (tblManufacturer INNER JOIN tbldataset ON tblManufacturer.UID = tbldataset.UID) ON (tblManufacturer.UID = tblmodel.UID)" & _
                " AND (tblmodel.FKmodel = tbldataset.FKmodel)) ON tblVariant.FKvariant = tbldataset.FKvariant) ON PRO_tbloperator.FKoperator = tbldataset.FKoperator WHERE (((tbldataset.ID)=" & Tologid & "));"
            Dim cmd2 As New OleDb.OleDbCommand(logged_SQL, Logged_con)
            cmd2.ExecuteNonQuery()
            If ToLogMil = 502 Then
                'Insert into loglls
                logged_SQL = "INSERT INTO logLLs ( ID, [when], logunit, [loga/c code], notes, other )" & _
                            " SELECT tbldataset.ID, " & Chr(34) & tologdate & Chr(34) & ", tblUnits.unit+" & Chr(34) & Chr(32) & Chr(34) & _
                            " &tblChildUnit.FamilyUnit AS logunit, [PRO-Marks].aCcode, LEFT([PRO-Marks].FleetName,25), tblChildUnit.FKtail" & _
                            " FROM ((tbldataset LEFT JOIN [PRO-Marks] ON tbldataset.ID = [PRO-Marks].ID) LEFT JOIN tblUnits ON tbldataset.FKParent = tblUnits.FKUnits) LEFT JOIN tblChildUnit" & _
                            " ON (tbldataset.FKChild = tblChildUnit.FKChild) AND (tbldataset.FKBaby = tblChildUnit.SubUnit)" & _
                            " WHERE (((tbldataset.ID)=" & Tologid & ") AND (tbldataset.FKChild) > 0);"
                cmd2 = New OleDb.OleDbCommand(logged_SQL, Logged_con)
                cmd2.ExecuteNonQuery()

            End If
            Logged_con.Close()
            Logged_con.Dispose()

            'Update BaseStation UserTag
            BS_Con.Open()
            Dim CheckBusy As Boolean = False
UpdBS:      CheckBusy = False
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
                GoTo UpdBS
            Else
                Try
                    Dim BS_SQL2 As String
                    BS_SQL2 = "UPDATE AIRCRAFT SET UserTag = ifnull(UserString1," & Chr(34) & Chr(32) & Chr(34) & "), LastModified = DATETIME(" & Chr(39) & "now" & Chr(39) & "," &
                    Chr(39) & "localtime" & Chr(39) & ") WHERE Modes = " & Chr(34) & ToLogHex & Chr(34) & Chr(59)
                    BS_Cmd = New SQLiteCommand(BS_SQL2, BS_Con)
                    BS_Cmd.ExecuteNonQuery()
                Catch SQLITEexception As Exception
                    If SQLiteErrorCode.Locked Then
                        CheckBusy = True
                        GoTo UpdBS
                    Else
                        CheckBusy = False
                    End If
                End Try
                BS_Con.Close()
                BS_Con.Dispose()
                ComboBox1.Items.Remove(ComboBox1.SelectedItem)
                ComboBox1.Refresh()
            End If
        End If
        Timer1.Start()
        If Process.GetProcessesByName("BaseStation.exe").Length >= 1 Then
            For Each ObjProcess As Process In Process.GetProcessesByName("BaseStation.exe")
                AppActivate(ObjProcess.Id)
                SendKeys.SendWait("{F5}")
            Next
            For Each ObjProcess As Process In Process.GetProcessesByName("PlanePlotter.exe")
                AppActivate(ObjProcess.Id)
                SendKeys.Send("^(Q)")
            Next
        End If

    End Sub

    Public Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        RunProcess()
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
        Timer1.Start()
    End Sub
End Class

