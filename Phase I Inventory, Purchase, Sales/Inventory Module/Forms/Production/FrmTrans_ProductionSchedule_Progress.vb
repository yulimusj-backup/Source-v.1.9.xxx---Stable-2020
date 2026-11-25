Option Explicit On
Option Strict Off

Public Class FrmTrans_ProductionSchedule_Progress
    Public vIdentifierJO As String = ""
    Public vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable
    Public vTBLItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1DataTable
    Public vTBLItemWO_ItemSch As DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable
    Public a As Integer = 0
    Public r As Integer
    Public g As Integer
    Public b As Integer
    Public vTBLSchBarKosong As DseTrans_Schedule.tbl_sch_barDataTable
    Public vTBLSchBarTanpaWO As DseTrans_Schedule.tbl_sch_barDataTable
    Public vPrimaryWU_Acuan As Double = -1
    Public vIdentifierWO_Acuan As String = ""
    Public ObjSch As ClsTrans_Schedule
    Public vStartTimeAcuan As Date
    Public vPrimaryWUConfig As Double = -1
    Dim ObjCD As New Core_Module.ClsCekData

    Private Sub FrmTrans_ProductionSchedule_Reschedule_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Control.CheckForIllegalCrossThreadCalls = False

            Me.LblDescription.Text = ""
            Me.LblProgress.Text = ""
            Me.bgw1.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgw1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgw1.DoWork
        Try
            ObjCD.UpdateThreading()
            Me.ObjSch.Reschedule_JadwalBaru1(Me.vIdentifierJO, Me.vTBLSchBar, Me.vTBLItemWO, Me.vTBLItemWO_ItemSch, Me.a, Me.r, Me.g, Me.b, Me.vTBLSchBarKosong, _
            Me.vTBLSchBarTanpaWO, Me, vStartTimeAcuan, vPrimaryWUConfig, Me.vPrimaryWU_Acuan, Me.vIdentifierWO_Acuan)
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