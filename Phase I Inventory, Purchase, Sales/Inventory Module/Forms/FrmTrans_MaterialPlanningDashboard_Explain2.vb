Public Class FrmTrans_MaterialPlanningDashboard_Explain2
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub FrmTrans_MaterialPlanningDashboard_Explain2_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try
            Me.dgvData.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvDataSR.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class