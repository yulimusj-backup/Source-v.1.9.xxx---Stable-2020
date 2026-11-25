Public Class FrmMasterKaryawan_DetailNama
    Public frm As Core_Module.FrmMasterKaryawan
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click

        Try
            frm.TxtPrefix.Text = Me.TxtPrefix.Text
            frm.TxtFirstName.Text = Me.TxtFirstName.Text
            frm.TxtMiddleName.Text = Me.TxtMiddleName.Text
            frm.TxtLastName.Text = Me.TxtLastName.Text
            frm.TxtSuffix.Text = Me.TxtSuffix.Text
            Me.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub TxtFirstName_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtFirstName.GotFocus
        Me.TxtFirstName.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtFirstName_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtFirstName.LostFocus
        Me.TxtFirstName.BackColor = Color.White
    End Sub

    Private Sub TxtLastName_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLastName.GotFocus
        Me.TxtLastName.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtLastName_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLastName.LostFocus
        Me.TxtLastName.BackColor = Color.White
    End Sub

    Private Sub TxtMiddleName_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMiddleName.GotFocus
        Me.TxtMiddleName.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtMiddleName_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMiddleName.LostFocus
        Me.TxtMiddleName.BackColor = Color.White
    End Sub

    Private Sub TxtPrefix_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPrefix.GotFocus
        Me.TxtPrefix.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtPrefix_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPrefix.LostFocus
        Me.TxtPrefix.BackColor = Color.White
    End Sub

    Private Sub TxtSuffix_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSuffix.GotFocus
        Me.TxtSuffix.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtSuffix_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSuffix.LostFocus
        Me.TxtSuffix.BackColor = Color.White
    End Sub

    Private Sub FrmMasterKaryawan_DetailNama_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.ShowIcon = False
    End Sub
End Class