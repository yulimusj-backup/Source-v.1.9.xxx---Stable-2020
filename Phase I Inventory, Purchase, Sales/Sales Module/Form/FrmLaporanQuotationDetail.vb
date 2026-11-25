Public Class FrmLaporanQuotationDetail
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public UserCode As String = ""

    Private Sub FrmLaporanQuotationDetail_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.CboFormat.SelectedIndex = 0
            Me.cboRealization.SelectedIndex = 0
            Me.cboAuthorization.SelectedIndex = 1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class