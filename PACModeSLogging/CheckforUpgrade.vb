Imports System.Data.OleDb
Imports System.IO
Imports System.Net

Module CheckforUpgrades

    Friend Currentloggedversion As Integer
    Friend CurrentICAOversion As Integer
    Friend Newestloggedversion As String
    Friend Newestloggedversionsplit As String()
    Friend NewestICAOversion As String
    Friend NewestICAOversionsplit As String()
    Dim Versions As String()
    Friend Newestappversion As String
    Friend Currentappversion As String = Application.ProductVersion
    Dim Lines1 As String()
    Dim Product As String

    Friend Newestloggedversionvalue As Integer
    Friend NewestICAOversionvalue As Integer





    Public Sub UpgradeCheck()
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
        Dim Request As HttpWebRequest = System.Net.HttpWebRequest.Create("https://www.gfiapac.org/ModeSVersions/PACModeS2020version.txt")
        Dim Response As HttpWebResponse = Request.GetResponse
        Dim Stream As Stream = Response.GetResponseStream()

        Using SR As StreamReader = New StreamReader(Response.GetResponseStream)

            Dim StreamString As String = SR.ReadToEnd

            If StreamString.Length > 0 Then
                Lines1 = StreamString.Split(vbLf)

                'App version check
                Dim result As String() = Array.FindAll(Lines1, Function(s) s.Contains(Application.ProductName))
                Product = result(0)
                Versions = Product.Split(":")
                Newestappversion = Versions(1)

                Dim UpgradeText As String = "Version " & Newestappversion & " is available - do you want to upgrade now?"
                Dim UpgradeFile As String = Replace(Newestappversion, ".", "")
                Dim UpgradeURL As String = "https://www.gfiapac.org/members/ModeS/PACModeSLogging_v" & UpgradeFile & "_install.exe"


                'Messaging section
                If Newestappversion = Currentappversion Then
                    'Carry on
                ElseIf (Newestappversion) <> (Currentappversion) Then
                    'MsgBox(UpgradeText, vbYesNo, "Upgrade check")
                    Select Case MsgBox(UpgradeText, MsgBoxStyle.YesNo, "PACModeSLogging Upgrade check")
                        Case MsgBoxResult.Yes
                            Dim sInfo As New ProcessStartInfo(UpgradeURL)
                            Process.Start(sInfo)
                            End
                        Case MsgBoxResult.No
                            'Carry on
                    End Select
                End If
            End If
        End Using

    End Sub
End Module
