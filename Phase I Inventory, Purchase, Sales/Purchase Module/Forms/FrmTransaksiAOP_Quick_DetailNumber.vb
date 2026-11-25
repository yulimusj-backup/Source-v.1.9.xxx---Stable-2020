Public Class FrmTransaksiAOP_Quick_DetailNumber
    Public PrimaryAOPQuickAcuan As Double = -1
    Public NoPO_Lama As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim ObjPO As New ClsTransaksiPO

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click

        Try
            Me.TxtPONumber.Text = Me.TxtPONumber.Text.Trim
            If Me.TxtPONumber.Text = "" Then
                MsgBox("Please specify PO Number.", MsgBoxStyle.Information, "Information")
                Me.TxtPONumber.Focus()
                Exit Sub
            End If

            If Me.ObjPO.ApakahSudahAda(Me.TxtPONumber.Text, Me.PrimaryAOPQuickAcuan) = True Then
                MsgBox("PO Number duplication, please specify different PO number.", MsgBoxStyle.Information, "Infomation")
                Me.TxtPONumber.Focus()
                Exit Sub
            End If

            Me.DialogResult = Windows.Forms.DialogResult.OK
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            Me.DialogResult = Windows.Forms.DialogResult.Cancel
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPONumber_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPONumber.GotFocus
        Me.TxtPONumber.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtPONumber_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPONumber.LostFocus
        Me.TxtPONumber.BackColor = Me.TextBox1.BackColor
    End Sub

    Private Sub FrmTransaksiAOP_Quick_DetailNumber_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.TxtPONumber.Text = Me.NoPO_Lama
        Me.TxtPONumber.Focus()
        Me.ShowIcon = False
    End Sub
End Class