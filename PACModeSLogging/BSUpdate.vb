Imports System.Data.SQLite
Module BSUpdate
    Dim CheckBusy As Boolean = False



    Public Sub UpdateBS(ToLogHex As String)
        'Update BaseStation UserTag


        Try
            If BS_Con.State = ConnectionState.Open Then BS_Con.Close()
            If BS_Con.State = ConnectionState.Closed Then BS_Con.Open()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Basestation Connection Error")
        End Try
        'BS_Con.Open()
UpdBS2: CheckBusy = False
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
                PACModeSLogging.ComboBox1.Items.Remove(PACModeSLogging.ComboBox1.SelectedItem)
                PACModeSLogging.ComboBox1.Refresh()
                MsgBox("do i get here", vbOKOnly)
                PACModeSLogging.Timer1.Start()
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
        End If

    End Sub
End Module
