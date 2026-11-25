Public Class FrmTrans_MaterialPlanningDashboard_AuthorizeDetail
    Public backColorAll As Color = Core_Module.GetSMESSBackColorAll
    Dim objCd As New Core_Module.ClsCekData

    Private Sub FrmTrans_MaterialPlanningDashboard_AuthorizeDetail_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.backColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvItem.DataError

    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
End Class