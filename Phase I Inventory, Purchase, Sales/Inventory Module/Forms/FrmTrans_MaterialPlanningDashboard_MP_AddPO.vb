Public Class FrmTrans_MaterialPlanningDashboard_MP_AddPO
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim LangsungTutup As Boolean = False

    Private Sub FrmTrans_MaterialPlanningDashboard_MP_AddPO_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.LangsungTutup = False Then
                If MsgBox("Cancel po creation?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Me.DialogResult = Windows.Forms.DialogResult.Cancel
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_MaterialPlanningDashboard_MP_AddPO_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            'Me.dgvDetail.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            'Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            Me.DialogResult = Windows.Forms.DialogResult.OK
            Me.LangsungTutup = True
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class