Option Explicit On
Option Strict Off

Public Class FrmTransaksiSO_SetDeliveryDate
    Private Sub RboSetDeliveryDate_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RboSetDeliveryDate.CheckedChanged
        Try
            If Me.RboSetDeliveryDate.Checked = True Then
                Me.DtpDate.Enabled = True
            Else
                Me.DtpDate.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnOK.Click
        Try
            Me.DialogResult = Windows.Forms.DialogResult.OK
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class