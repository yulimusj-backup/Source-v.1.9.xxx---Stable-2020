Public Class FrmTransferBeginingBalance_Error
    Private Sub btnContinue_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnContinue.Click
        Try
            If MsgBox("Are you sure to continue and skip some transaction above?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.DialogResult = Windows.Forms.DialogResult.OK
                Me.Close()
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btnCancel2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel2.Click
        Try
            If MsgBox("Are you sure to cancel?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.DialogResult = Windows.Forms.DialogResult.Cancel
                Me.Close()
            End If
        Catch ex As Exception

        End Try
    End Sub
End Class