Public Class FrmPilihCetakLoan
    Public vFromBeginingBalance As Boolean = False

    Private Sub BtnShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnShow.Click
        Try
            Me.DialogResult = Windows.Forms.DialogResult.OK
            Me.Hide()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmPilihCetakLoan_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Me.vFromBeginingBalance = True Then
                Me.RbtJournal.Enabled = False
                Me.RbtDetail.Checked = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class