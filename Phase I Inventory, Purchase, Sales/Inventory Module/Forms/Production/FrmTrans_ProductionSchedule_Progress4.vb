Public Class FrmTrans_ProductionSchedule_Progress4
    Public PrimaryWU As Double
    Public DurasiGeser As Double
    Public TBLSchBarKosong As DseTrans_Schedule.tbl_sch_barDataTable
    Public IdentifierWOAcuan As String
    Public StartDateAcuan As Date
    Public IdentifierJOAcuan As String
    Public ObjSch As ClsTrans_Schedule
    Dim ObjCD As New Core_Module.ClsCekData

    Private Sub FrmTrans_ProductionSchedule_Progress3_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
            Me.ObjSch.GeserJadwal(Me.PrimaryWU, Me.DurasiGeser, Me.TBLSchBarKosong, Me.IdentifierWOAcuan, Me.StartDateAcuan, Me.IdentifierJOAcuan, Me)
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