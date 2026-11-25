Public Class FrmTrans_SPL_Cetak_Option

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Try

            Me.DialogResult = Windows.Forms.DialogResult.OK

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub
End Class