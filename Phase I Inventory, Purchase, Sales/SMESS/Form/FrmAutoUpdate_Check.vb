Public Class FrmAutoUpdate_Check
    'Public Download_SMESS As Integer = 0
    'Public Download_ClientUpdate As Integer = 0

    Private Sub FrmAutoUpdate_Check_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.Text = "Please wait, while we check for updates."
            Control.CheckForIllegalCrossThreadCalls = False
            Me.BackgroundWorker1.RunWorkerAsync()
            'Me.CheckForNewUpdate()
            'Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Private Sub CheckForNewUpdate_ClientUpdate()
    '    Dim ObjAuto As New ClsUpdateProgram

    '    Try
    '        Threading.Thread.Sleep(1000)
    '        If ObjAuto.ConnectionInternet_ClientUpdate(ClsUpdateProgram.EnumFungsi._download, False) = True Then
    '            If ObjAuto.RunProsesDownloadUpdate_ClientUpdate(ClsUpdateProgram.EnumFungsi._download, False) = True Then
    '                Me.DialogResult = Windows.Forms.DialogResult.OK
    '                Exit Sub
    '            End If
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        ObjAuto = Nothing
    '    End Try
    'End Sub

    Private Sub CheckForNewUpdate()
        Dim ObjAuto As New ClsUpdateProgram

        Try
            Threading.Thread.Sleep(1000)
            If ObjAuto.ConnectionInternet(ClsUpdateProgram.EnumFungsi._download, True) = True Then
                If ObjAuto.RunProsesDownloadUpdate(ClsUpdateProgram.EnumFungsi._download, True) = True Then
                    Me.DialogResult = Windows.Forms.DialogResult.OK
                    'Me.Download_SMESS = 1
                    Exit Sub
                End If
            End If

            Me.DialogResult = Windows.Forms.DialogResult.Cancel
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjAuto = Nothing
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Me.CheckForNewUpdate()
            'Me.CheckForNewUpdate_ClientUpdate()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            'If Me.Download_SMESS = 0 And Me.Download_ClientUpdate = 0 Then
            'Me.DialogResult = Windows.Forms.DialogResult.Cancel
            'MsgBox("Application is up to date", MsgBoxStyle.Information, "Information")
            'Else
            'Me.DialogResult = Windows.Forms.DialogResult.OK
            'End If
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class