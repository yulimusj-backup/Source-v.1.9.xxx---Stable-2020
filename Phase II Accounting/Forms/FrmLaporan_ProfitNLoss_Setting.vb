Public Class FrmLaporan_ProfitNLoss_Setting
    Private Sub chkShowLvl2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowLvl2.CheckedChanged
        Try
            If Me.chkShowLvl2.Checked = True Then
                Me.chkShowLvl3.Enabled = True
                Me.chkShowLvl4.Enabled = False
                Me.chkShowLvl5.Enabled = False
                Me.chkShowLvl6.Enabled = False
                Me.chkShowLvl7.Enabled = False
                Me.chkShowLvl8.Enabled = False
            Else
                Me.chkShowLvl3.Enabled = False
                Me.chkShowLvl4.Enabled = False
                Me.chkShowLvl3.Checked = False
                Me.chkShowLvl4.Checked = False
                Me.chkShowLvl5.Enabled = False
                Me.chkShowLvl5.Checked = False
                Me.chkShowLvl6.Enabled = False
                Me.chkShowLvl6.Checked = False
                Me.chkShowLvl7.Enabled = False
                Me.chkShowLvl7.Checked = False
                Me.chkShowLvl8.Enabled = False
                Me.chkShowLvl8.Checked = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkShowLvl3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowLvl3.CheckedChanged
        Try
            If Me.chkShowLvl3.Checked = True Then
                Me.chkShowLvl4.Enabled = True
                Me.chkShowLvl5.Enabled = False
                Me.chkShowLvl6.Enabled = False
                Me.chkShowLvl7.Enabled = False
                Me.chkShowLvl8.Enabled = False
            Else
                Me.chkShowLvl4.Enabled = False
                Me.chkShowLvl4.Checked = False
                Me.chkShowLvl5.Enabled = False
                Me.chkShowLvl5.Checked = False
                Me.chkShowLvl6.Enabled = False
                Me.chkShowLvl6.Checked = False
                Me.chkShowLvl7.Enabled = False
                Me.chkShowLvl7.Checked = False
                Me.chkShowLvl8.Enabled = False
                Me.chkShowLvl8.Checked = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkShowLvl4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowLvl4.CheckedChanged
        Try
            If Me.chkShowLvl4.Checked = True Then
                Me.chkShowLvl5.Enabled = True
                Me.chkShowLvl6.Enabled = False
                Me.chkShowLvl7.Enabled = False
                Me.chkShowLvl8.Enabled = False
            Else
                Me.chkShowLvl5.Enabled = False
                Me.chkShowLvl5.Checked = False
                Me.chkShowLvl6.Enabled = False
                Me.chkShowLvl6.Checked = False
                Me.chkShowLvl7.Enabled = False
                Me.chkShowLvl7.Checked = False
                Me.chkShowLvl8.Enabled = False
                Me.chkShowLvl8.Checked = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkShowLvl6_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowLvl6.CheckedChanged
        Try
            If Me.chkShowLvl6.Checked = True Then
                Me.chkShowLvl7.Enabled = True
                Me.chkShowLvl8.Enabled = False
            Else
                Me.chkShowLvl7.Enabled = False
                Me.chkShowLvl7.Checked = False
                Me.chkShowLvl8.Enabled = False
                Me.chkShowLvl8.Checked = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkShowLvl5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowLvl5.CheckedChanged
        Try
            If Me.chkShowLvl5.Checked = True Then
                Me.chkShowLvl6.Enabled = True
                Me.chkShowLvl7.Enabled = False
                Me.chkShowLvl8.Enabled = False
            Else
                Me.chkShowLvl6.Enabled = False
                Me.chkShowLvl6.Checked = False
                Me.chkShowLvl7.Enabled = False
                Me.chkShowLvl7.Checked = False
                Me.chkShowLvl8.Enabled = False
                Me.chkShowLvl8.Checked = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkShowLvl7_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowLvl7.CheckedChanged
        Try
            If Me.chkShowLvl7.Checked = True Then
                Me.chkShowLvl8.Enabled = True
            Else
                Me.chkShowLvl8.Enabled = False
                Me.chkShowLvl8.Checked = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click
        Me.DialogResult = Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.DialogResult = Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub
End Class