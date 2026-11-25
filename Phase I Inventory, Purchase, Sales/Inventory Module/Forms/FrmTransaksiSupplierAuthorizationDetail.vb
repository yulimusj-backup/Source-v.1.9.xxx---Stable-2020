Public Class FrmTransaksiSupplierAuthorizationDetail
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub FrmTransaksiSupplierAuthorizationDetail_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.dgvDetail.AlternatingRowsDefaultCellStyle.BackColor = Core_Module.GetSMESSBackColorAll
    End Sub
End Class