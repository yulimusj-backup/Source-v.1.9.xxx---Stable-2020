Public Class FrmUserLocationPrivilegeSetting2

    Private Sub FrmUserLocationPrivilegeSetting2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.dgvLocation.AlternatingRowsDefaultCellStyle.BackColor = Core_Module.GetSMESSBackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class