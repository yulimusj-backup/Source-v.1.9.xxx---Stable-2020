Option Explicit On
Option Strict Off

Imports System.Net
Imports System.IO
Imports System.Environment

Public Class FrmDownload3
    Public vShowMsg As Boolean = False

    Dim vServerPathDontExist As Boolean = False
    Dim NoUpdateAvailable As Boolean = False
    Dim vRetValConfirmUlang As Boolean = False

    Dim ObjAutoUpdate As New ClsAutoUpdate
    Dim ObjUpdate As New ClsAutoUpdate
    Dim ObjURL As New CustomSMESS.ClsUpdate
    Dim ObjUpdate1 As New ClsUpdateProgram
    Dim Counter As Integer = 3
    Dim RetVal As Boolean = False
    Dim DefaultVersion As String = "1.6.40"

    Private Sub FrmDownload3_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Try
        '    MsgBox("")
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'End Try
    End Sub

    Private Sub FrmCekUpdate_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            'Me.Timer1.Interval = 3000
            Me.btnClose.Height = 25
            Me.Timer1.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim ObjAuto As New ClsUpdateProgram
        Dim Obj1 As New ClsAutoUpdate
        Dim ObjAutoUpdate As New ClsAutoUpdate
        Dim Frm As New FrmDownload3

        Try
            If System.IO.Directory.Exists(Obj1.GetServerPath) = False Then
                Me.vServerPathDontExist = True
                Me.RetVal = False
            Else
                If Me.RunProsesDownloadUpdate(ClsUpdateProgram.EnumFungsi._UpdateViaServer, vShowMsg) = True Then
                    Me.RetVal = True
                Else
                    Me.RetVal = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjAuto = Nothing
            Obj1 = Nothing
        End Try
    End Sub

    Public Function RunProsesDownloadUpdate(ByVal vEnumFungsi As ClsUpdateProgram.EnumFungsi, ByVal vShowMsg As Boolean) As Boolean
        Dim NewVersionTxt As String = "0.0.00"
        Dim LastVersionTxt As String = "0.0.00"
        Dim ServerPath As String = Me.ObjAutoUpdate.GetServerPath()
        Dim ObjVersion As New ClsVersionControl
        Dim ObjIdentifier As New CustomSMESS.ClsIdentifier

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            If vEnumFungsi = ClsUpdateProgram.EnumFungsi._download Then
                If Me.ObjUpdate1.ConnectionInternet(vEnumFungsi, vShowMsg) = False Then
                    Return False
                End If
                'download
                Me.ObjUpdate1.Download(Me.ObjURL.GetFileVersionTxt_SMESS, Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "SMESS_Version.smess")
            ElseIf vEnumFungsi = ClsUpdateProgram.EnumFungsi._UpdateViaServer Then
                If ServerPath = "" Then
                    Me.vServerPathDontExist = True
                    Return False
                End If
                'copy File via server
                If System.IO.File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "SMESS_Version_Current" + "_" + ObjIdentifier.GetIdentifier + ".smess") = True Then
                    System.IO.File.SetAttributes(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "SMESS_Version_Current" + "_" + ObjIdentifier.GetIdentifier + ".smess", FileAttributes.Normal)
                End If
                If System.IO.File.Exists(ServerPath & "\" & "SMESS_Version.smess") = True Then
                    System.IO.File.Copy(ServerPath & "\" & "SMESS_Version.smess" _
                                                     , Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "SMESS_Version.smess", True)
                Else
                    Me.NoUpdateAvailable = True
                    Return False
                End If
            End If

            'decrypt file
            Me.ObjUpdate1.DecryptFile(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "SMESS_Version.smess")

            'bandingin version
            NewVersionTxt = My.Computer.FileSystem.ReadAllText(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "SMESS_Version.txt")
            ''NewVersionTxt = "1.7.00"
            LastVersionTxt = ObjVersion.GetVersion(Me.DefaultVersion)
            'MsgBox("Please cek file")
            If Me.ObjUpdate1.ApakahVersionBerbeda(LastVersionTxt, NewVersionTxt) = False Then
                System.IO.File.Delete(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "SMESS_Version.txt")
                System.IO.File.Delete(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\" & "SMESS_Version.smess")
                If vShowMsg = True Then
                    Me.BackgroundWorker1.ReportProgress(1)
                    If Me.vRetValConfirmUlang = True Then
                        Return True
                    Else
                        Me.NoUpdateAvailable = True
                        Return False
                    End If
                End If
                Me.NoUpdateAvailable = True
                Return False
            Else
                Return True
            End If
            Windows.Forms.Cursor.Current = Cursors.Arrow
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            ObjVersion = Nothing
        Finally
            ObjIdentifier = Nothing
        End Try
    End Function

    Private Sub BackgroundWorker1_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        Try
            If MsgBox("We detect your smess is up to date, update anyway?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.vRetValConfirmUlang = True
            Else
                Me.vRetValConfirmUlang = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            If Me.RetVal = False Then
                If Me.vServerPathDontExist = True Then
                    'If vShowMsg = True Then
                    Me.LblTypeFile.Text = "Server path doesn't exists."
                    'End If
                ElseIf Me.NoUpdateAvailable = True Then
                    'If vShowMsg = True Then
                    Me.LblTypeFile.Text = "There is no update available."
                    'End If
                End If
                Me.Timer2.Enabled = True
                Beep()
            Else
                Me.DialogResult = Windows.Forms.DialogResult.OK
                Me.Close()
            End If
            'If Me.DialogResult = Windows.Forms.DialogResult.OK Then

            'End If
            'Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Try
            Control.CheckForIllegalCrossThreadCalls = False
            Me.BackgroundWorker1.RunWorkerAsync()
            Me.Timer1.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Timer2_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        Try
            If Me.Counter = 3 Then
                Me.btnClose.Visible = True
                Me.btnClose.Focus()
            End If
            Me.btnClose.Text = "Close (" + Me.Counter.ToString + ")"
            Counter -= 1
            If Me.Counter <= 0 Then
                Me.DialogResult = Windows.Forms.DialogResult.Cancel
                Me.Timer2.Enabled = False
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.DialogResult = Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub
End Class