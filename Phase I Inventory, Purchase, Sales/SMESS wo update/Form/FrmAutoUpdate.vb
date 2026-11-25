Option Explicit On
Option Strict Off

Imports System.IO
Imports System.Net

Public Class FrmAutoUpdate
    Dim ObjUpdate As New Core_Module.ClsTransaksi_SMESSAutoUpdate
    Private WithEvents httpclient As WebClient
    Dim ObjDB As New Core_Module.ClsDBVersion

    Private Sub FrmAutoUpdate_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.ProgressBar1.Style = ProgressBarStyle.Marquee
            Me.Timer1.Interval = Me.ObjUpdate.Suspend
            Me.Timer1.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Dim webclient As New WebClient
        httpclient = New WebClient
        Dim URL As String = "http://arthajina.com/downloads/SMESS/CMSI/version.txt"
        Dim Source As New System.Text.StringBuilder
        Dim webresponse As Net.WebResponse = Net.WebRequest.Create(URL).GetResponse
        Dim StreamReader As New IO.StreamReader(webresponse.GetResponseStream)
        Dim isAvailable As Boolean

        Try
            Me.Timer1.Enabled = False

            isAvailable = My.Computer.Network.IsAvailable
            If isAvailable Then
                ProgressBar1.Value = 0
                ProgressBar1.Maximum = 100

                Do
                    Source.Append(StreamReader.ReadLine)
                Loop Until StreamReader.ReadLine Is Nothing
                StreamReader.Close()


                If webclient.DownloadString("http://arthajina.com/downloads/SMESS/CMSI/version.txt") <> ObjDB.GetDBVersion Then
                    Me.Timer1.Enabled = False
                    'httpclient.DownloadFileAsync(New Uri("http://arthajina.com/downloads/SMESS/CMSI/Setup SMESS.msi"), _
                    '"Setup SMESS.msi")
                    httpclient.DownloadFileAsync(New Uri("\\ROSADIN-PC\Users\Public\Setup SMESS.msi"), _
                    "Setup SMESS1.msi")
                Else
                    Me.Timer1.Enabled = False
                    MessageBox.Show("You're already using latest SMESS version", Source.ToString)
                    Me.Close()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Me.Timer1.Enabled = False
        End Try
    End Sub

    Private Sub httpclient_DownloadFileCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.AsyncCompletedEventArgs) Handles httpclient.DownloadFileCompleted
        Try
            If e.Cancelled Then
            ElseIf Not e.Error Is Nothing Then
                MsgBox("(1) There was a problem downloading the update,  please try again later.")
                Me.Close()
            Else
                'Lets see if the test.msi file or what ever you wanted to call it exist on the drive
                'if so we gonna remove it first.

                If My.Computer.FileSystem.FileExists(Application.StartupPath & "\Setup SMESS.msi") Then
                    My.Computer.FileSystem.DeleteFile(Application.StartupPath & "\Setup SMESS.msi", FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
                End If


                'After the dowload has finished we gonna rename the test.msi.zip to test.msi or what name you used.

                If My.Computer.FileSystem.FileExists(Application.StartupPath & "\Setup SMESS1.msi") Then
                    My.Computer.FileSystem.RenameFile(Application.StartupPath & "\Setup SMESS1.msi", "Setup SMESS.msi")
                End If

                If My.Computer.FileSystem.FileExists(Application.StartupPath & "\Setup SMESS.msi") Then
                    'Now the update has really been finsihed to lets try to run the test.msi at once before we exit the update.
                    If ProgressBar1.Value = ProgressBar1.Maximum Then
                        Dim p As Process = New Process()
                        p.StartInfo.FileName = "Setup SMESS.msi"
                        p.Start()

                        FrmLogin.Close()
                        Me.Close()
                    End If
                Else
                    MsgBox("(2) There was a problem downloading the update,  please try again later.")
                    Me.Close()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub httpclient_DownloadProgressChanged(ByVal sender As Object, ByVal e As System.Net.DownloadProgressChangedEventArgs) Handles httpclient.DownloadProgressChanged
        Try
            Me.ProgressBar1.Style = ProgressBarStyle.Continuous
            Me.ProgressBar1.Maximum = e.TotalBytesToReceive
            Me.ProgressBar1.Value = e.BytesReceived
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
