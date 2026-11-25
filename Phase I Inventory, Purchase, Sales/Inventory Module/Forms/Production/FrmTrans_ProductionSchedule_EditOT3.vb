Public Class FrmTrans_ProductionSchedule_EditOT3

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub FrmTrans_ProductionSchedule_EditOT3_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.dgvCek.AlternatingRowsDefaultCellStyle.BackColor = Core_Module.GetSMESSBackColorAll

            Beep()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class