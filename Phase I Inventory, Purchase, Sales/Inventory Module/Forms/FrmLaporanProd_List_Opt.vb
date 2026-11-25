Public Class FrmLaporanProd_List_Opt
    Public BAckColorALL As Color = Core_Module.GetSMESSBackColorAll

    Private Sub BtnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnOK.Click
        Me.DialogResult = Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub TxtDesc_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDesc.GotFocus
        Me.TxtDesc.BackColor = Me.BAckColorALL
    End Sub

    Private Sub TxtDesc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDesc.LostFocus
        Me.TxtDesc.BackColor = Me.TextKosong.BackColor
    End Sub

    Private Sub FrmLaporanProd_List_Opt_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.TxtDesc.Focus()
    End Sub

    Private Sub btnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
End Class