Public Class FrmTrans_ProductionSchedule_Progress2
    Public ObjSch As ClsTrans_Schedule
    Public vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable
    Public vIdentifierJO As String
    Public vTBLItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1DataTable
    Public vTBLItemWO_ItemSch As DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable
    Public vTargetFinishDate As Date
    Public vPrimaryPredefinedWU As Double
    Public vTBLWU As DseTrans_Schedule.tblmaster_product_item_workunit1DataTable
    Public vTBLSchBarKosong As DseTrans_Schedule.tbl_sch_barDataTable
    Public vTBLSchBarTanpaWO As DseTrans_Schedule.tbl_sch_barDataTable
    Dim ObjCD As New Core_Module.ClsCekData

    Private Sub FrmTrans_ProductionSchedule_Progress2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Control.CheckForIllegalCrossThreadCalls = False

            Me.LblDescription.Text = ""
            Me.LblProgress.Text = ""
            Me.bgw1.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgw1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgw1.DoWork
        Try
            ObjCD.UpdateThreading()
            Me.ObjSch.HitungJadwalBaru(Me.vTBLSchBar, Me.vIdentifierJO, Me.vTBLItemWO, Me.vTBLItemWO_ItemSch, Me.vTargetFinishDate, Me.vPrimaryPredefinedWU, _
            Me.vTBLWU, Me.vTBLSchBarKosong, Me.vTBLSchBarTanpaWO, Me)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgw1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgw1.RunWorkerCompleted
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class