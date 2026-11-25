Option Explicit On
Option Strict Off

Public Class FrmTrans_ManufacturingOrder_SetFinishWOStatus

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            If MsgBox("Are you sure?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.ControlBox = False
                Me.Button1.Enabled = False
                Me.Cursor = Cursors.WaitCursor
                FrmTrans_ManufacturingOrder_SetFinishWOStatus.CheckForIllegalCrossThreadCalls = False
                Me.BackgroundWorker1.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Err")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim Dta As New DseTrans_StockBundlingTableAdapters.tbltrans_mo_item_wo2TableAdapter
        Dim ObjCd As New Core_Module.ClsCekData
        Try
            ObjCd.UpdateThreading()
            Me.ProgressBar1.Enabled = True
            Me.ProgressBar1.Value = 5
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            Dta.UpdateQuery_SetFinishWO_Status(ClsTrans_ManufacturingOrder.EnumPilihStatusWO.Finish)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Err")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Me.Cursor = Cursors.Default
            Me.Button1.Enabled = True
            Me.ControlBox = True
            Me.ProgressBar1.Value = 0
            Me.ProgressBar1.Enabled = False
            MsgBox("Finished.", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Err")
        End Try
    End Sub
End Class