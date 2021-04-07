Imports System.Data.OleDb

Module GFIAUpdate

    Public LogText As String
    Dim logged_cmd As New OleDbCommand
    Dim cmd2 As New OleDbCommand
    Dim cmd3 As New OleDbCommand
    Public ToLogdate As String
    Public Where As String
    Public Source As String = "PacModes v" & My.Application.Info.Version.ToString
    Public ToLogUnit As String = " "
    Public ToLogaCcode As String = " "
    Public ToLogacName As String = " "
    Public ToLogOther As String = " "
    Public ToLogNotes As String = " "




    Public Sub UpdateGFIA(ToLogReg As String, ToLogHex As String, ToLogId As String, ToLogMil As Integer, MDPO As String)


        Dim LogNote As New Notes
        LogNote.ShowDialog()

        If LogNote.Canx = 0 Then

            LogText = LogNote.Passvalue

            'Get Model Prefix
            Dim Prefix As String = ""
            Dim Prefix_SQL As String = ""
            Prefix_SQL = "SELECT tbldataset.ID, tbldataset.FKsp2 as PrefixText " &
            "FROM tbldataset WHERE (((tbldataset.ID)=" & ToLogId & "))"
            Dim cmd4 As OleDbCommand
            cmd4 = New OleDb.OleDbCommand(Prefix_SQL, Logged_con)
            Dim cmd4_rdr As OleDbDataReader
            cmd4_rdr = cmd4.ExecuteReader

            'Get Typnames
            Dim typename As String = ""
            Dim TypeName_SQL As String = ""
            TypeName_SQL = "SELECT tbldataset.ID, tblVariant.FKname AS TypeNameKey, tblNames.TypeName " &
                            " FROM (tblVariant INNER JOIN tbldataset ON tblVariant.FKvariant = tbldataset.FKvariant)" &
                            " INNER JOIN tblNames ON tblVariant.FKname = tblNames.FKname WHERE (((tbldataset.ID)=" & ToLogId & "))"
            cmd3 = New OleDb.OleDbCommand(TypeName_SQL, Logged_con)
            Dim cmd3_rdr As OleDbDataReader
            cmd3_rdr = cmd3.ExecuteReader()

            If cmd3_rdr.HasRows = False Then

                'Insert into logllp no typename no prefix
                logged_SQL = "INSERT INTO logllp SELECT tbldataset.ID AS ID, ([tblManufacturer].[Builder]+' '+[tblmodel].[Model]+RIGHT([tblvariant].[variant], LEN([tblvariant].[variant]) - 1)" &
                                "+' '+'['+[tbldataset].[cn]+']') AS Aircraft, tbldataset.Registration AS Registration, PRO_tbloperator.operator AS operator, " & Chr(34) & Where & Chr(34) & " AS [Where]," & Chr(34) & MDPO & Chr(34) & " AS MDPO, " &
                                Chr(34) & tologdate & Chr(34) & " As [When], 0 As Lockk, " & Chr(34) & LogText & Chr(34) & " AS [Notes], " & Chr(34) & Source & Chr(34) & " AS [Source] " &
                                " FROM tblVariant INNER JOIN ((tblmodel INNER JOIN tblManufacturer ON tblmodel.UID = tblManufacturer.UID) " &
                                " INNER JOIN (PRO_tbloperator RIGHT JOIN tbldataset ON PRO_tbloperator.FKoperator = tbldataset.FKoperator) ON " &
                                " (tblManufacturer.UID = tbldataset.UID) AND (tblmodel.FKmodel = tbldataset.FKmodel)) " &
                                "ON tblVariant.FKvariant = tbldataset.FKvariant " & "WHERE (((tbldataset.ID)=" & ToLogId & "));"
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
                                Chr(34) & tologdate & Chr(34) & " As [When], 0 As Lockk, " & Chr(34) & LogText & Chr(34) & " AS [Notes], " & Chr(34) & Source & Chr(34) & " AS [Source] " &
                                " FROM (tblVariant INNER JOIN ((tblmodel INNER JOIN tblManufacturer ON tblmodel.UID = tblManufacturer.UID) " &
                                " INNER JOIN (PRO_tbloperator RIGHT JOIN tbldataset ON PRO_tbloperator.FKoperator = tbldataset.FKoperator) ON " &
                                "(tblmodel.FKmodel = tbldataset.FKmodel) And (tblManufacturer.UID = tbldataset.UID)) " &
                                "ON tblVariant.FKvariant = tbldataset.FKvariant) INNER JOIN tblNames ON tblVariant.FKname = tblNames.FKname " & "WHERE (((tbldataset.ID)=" & ToLogId & "));"
                cmd2 = New OleDb.OleDbCommand(logged_SQL, Logged_con)
                cmd2.ExecuteNonQuery()

            ElseIf cmd3_rdr.HasRows = False And cmd4_rdr(1) = "" Then

                'Insert into logllp no typename with no prefix
                logged_SQL = "INSERT INTO logllp SELECT tbldataset.ID AS ID, ([tblManufacturer].[Builder]+' '+[tblmodel].[Model]+RIGHT([tblvariant].[variant], LEN([tblvariant].[variant]) - 1)" &
                                "+' '+' ['+[tbldataset].[cn]+']') AS Aircraft, tbldataset.Registration AS Registration, PRO_tbloperator.operator AS operator, " & Chr(34) & Where & Chr(34) & " AS [Where]," & Chr(34) & MDPO & Chr(34) & " AS MDPO, " &
                                Chr(34) & tologdate & Chr(34) & " As [When], 0 As Lockk, " & Chr(34) & LogText & Chr(34) & " AS [Notes], " & Chr(34) & Source & Chr(34) & " AS [Source] " &
                                " FROM tblVariant INNER JOIN ((tblmodel INNER JOIN tblManufacturer ON tblmodel.UID = tblManufacturer.UID) " &
                                " INNER JOIN (PRO_tbloperator RIGHT JOIN tbldataset ON PRO_tbloperator.FKoperator = tbldataset.FKoperator) ON " &
                                " (tblManufacturer.UID = tbldataset.UID) AND (tblmodel.FKmodel = tbldataset.FKmodel)) " &
                                "ON tblVariant.FKvariant = tbldataset.FKvariant " & "WHERE (((tbldataset.ID)=" & ToLogId & "));"
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
                    Chr(34) & tologdate & Chr(34) & " As [When], 0 As Lockk, " & Chr(34) & LogText & Chr(34) & " AS [Notes], " & Chr(34) & Source & Chr(34) & " AS [Source] " &
                    " FROM (tblVariant INNER JOIN ((tblmodel INNER JOIN tblManufacturer ON tblmodel.UID = tblManufacturer.UID) " &
                    " INNER JOIN (PRO_tbloperator RIGHT JOIN tbldataset ON PRO_tbloperator.FKoperator = tbldataset.FKoperator) ON " &
                    "(tblmodel.FKmodel = tbldataset.FKmodel) And (tblManufacturer.UID = tbldataset.UID)) " &
                    "ON tblVariant.FKvariant = tbldataset.FKvariant) INNER JOIN tblNames ON tblVariant.FKname = tblNames.FKname " & "WHERE (((tbldataset.ID)=" & ToLogId & "));"
                cmd2 = New OleDb.OleDbCommand(logged_SQL, Logged_con)
                cmd2.ExecuteNonQuery()
            End If

            'Check for Fleetname and add
            If ToLogMil <> 502 Then
                logged_SQL = " SELECT tbldataset.ID, LEFT([PRO-Marks].FleetName,50)" &
                        " FROM tbldataset LEFT JOIN [PRO-Marks] ON tbldataset.ID = [PRO-Marks].ID " &
                        " WHERE (tbldataset.ID)= " & ToLogId
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
                        logged_SQL = "UPDATE logLLp SET logllp.acName = " & Chr(34) & ToLogacName & Chr(34) & " WHERE ID = " & ToLogId & ";"
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
                            " WHERE (((tbldataset.ID)=" & ToLogId & ") AND (tbldataset.FKChild) > 0);"
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
                                " , logllp.acName = " & Chr(34) & ToLogacName & Chr(34) & ", logllp.other = " & Chr(34) & ToLogOther & Chr(34) & " WHERE ID = " & ToLogId & ";"
                        cmd2 = New OleDb.OleDbCommand(logged_SQL, Logged_con)
                        cmd2.ExecuteNonQuery()
                        ToLogacName = String.Empty
                        ToLogaCcode = String.Empty
                        ToLogUnit = String.Empty
                        ToLogOther = String.Empty
                    End If

                End If
            End If
        End If
        UpdateBS(ToLogHex, ToLogReg)
    End Sub

End Module