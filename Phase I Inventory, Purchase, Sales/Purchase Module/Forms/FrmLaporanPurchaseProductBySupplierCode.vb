Option Explicit On
Option Strict Off

Public Class FrmLaporanPurchaseProductBySupplierCode
    Public UserCode As String = ""
    Public BackColorAll As Color = BackColorAll

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub txtSupplierCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSupplierCode.GotFocus
        txtSupplierCode.BackColor = Me.BackColorAll
    End Sub

    Private Sub txtSupplierCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSupplierCode.LostFocus
        txtSupplierCode.BackColor = Me.txtPrimariSupplier.BackColor
    End Sub

    Private Sub btnFindSupplier_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindSupplier.Click

    End Sub
End Class