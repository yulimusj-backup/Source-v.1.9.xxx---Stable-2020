Public Class FrmTransaksiMonitor
    Dim ObjMonitor As New ClsTransaksiMonitor
    Dim ObjAntrian As New Core_Module.ClsAntrian
    Dim ObjUser As New Core_Module.ClsUser
    Dim TipeTrans As ClsTransaksiMonitor.EnumTipeTransaksi = ClsTransaksiMonitor.EnumTipeTransaksi.SMESS_License

    Public BackColorAll As Color = Color.AliceBlue
    Public UserCode As String = ""
    Public FromMainMenu As Boolean = False

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub FrmTransaksiMonitor_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Try
            Me.ObjMonitor.ShowMonitor(Me.dgvMonitor, Me, TipeTrans)
            Me.ProgressBar1.Value = 0

            Me.ObjAntrian.ShowQueue(Me.dgvQueue, Me)
            Me.prgQueue.Value = 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiMonitor_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim JmlLicense As Double = -1

        Try
            Me.dgvMonitor.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvQueue.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            Me.ObjMonitor.ShowMonitor(Me.dgvMonitor, Me, Me.TipeTrans)
            Me.ProgressBar1.Value = 0
            Me.ProgressBar1.Maximum = 1000
            Me.Timer2.Enabled = True

            JmlLicense = Me.ObjMonitor.GetSMESSActiveLicense
            If JmlLicense <> -1 Then
                Me.Label1.Text = "Active SMESS License : " + JmlLicense.ToString + " user(s)."
            Else
                Me.Label1.Text = ""
                Me.dgvMonitor.Height = Me.ProgressBar1.Top - 3 - Me.dgvMonitor.Location.Y
            End If

            Me.ObjAntrian.ShowQueue(Me.dgvQueue, Me)
            Me.prgQueue.Value = 0
            Me.prgQueue.Maximum = 1000
            Me.Timer1.Enabled = True

            If Me.FromMainMenu = False Then
                Me.btnQueue_Clear.Enabled = False
            End If
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransactionQueue") = False Then
                Me.btnQueue_Clear.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        Try
            Me.ProgressBar1.Value += 1
            If Me.ProgressBar1.Value = 1000 Then
                Me.ProgressBar1.Value = 0
                Me.ObjMonitor.ShowMonitor(Me.dgvMonitor, Me, Me.TipeTrans)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefreshMonitor.Click
        Try
            Me.ObjMonitor.ShowMonitor(Me.dgvMonitor, Me, Me.TipeTrans)
            Me.ProgressBar1.Value = 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnQueue_Refresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnQueue_Refresh.Click
        Try
            Me.ObjAntrian.ShowQueue(Me.dgvQueue, Me)
            Me.prgQueue.Value = 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Try
            Me.prgQueue.Value += 1
            If Me.prgQueue.Value = 1000 Then
                Me.prgQueue.Value = 0
                Me.ObjAntrian.ShowQueue(Me.dgvQueue, Me)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnQueue_Clear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnQueue_Clear.Click
        Try
            Me.btnRefreshMonitor.PerformClick()
            'If Me.dgvMonitor.RowCount > 1 Then
            '    MsgBox("To clear queue data, please make sure only 1 user is online.", MsgBoxStyle.Information, "Information")
            '    Exit Sub
            'End If
            Me.ObjAntrian.HapusAntrian_HanyaYgLogOff()
            btnQueue_Refresh.PerformClick()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class