Public Class FrmGantiPassword
    Public KodeUser As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Private Sub FrmGantiPassword_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'set kode user
        Me.tbKodeUser.Text = Me.KodeUser
    End Sub

    Private Sub chkShowPassword_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowPassword.CheckedChanged
        Try
            If Me.chkShowPassword.Checked = True Then
                'tunjukkan password
                Me.tbKonfirmasiPassword.UseSystemPasswordChar = False
                Me.tbPasswordBaru.UseSystemPasswordChar = False
            Else
                'gunakan sistem password proc
                Me.tbKonfirmasiPassword.UseSystemPasswordChar = True
                Me.tbPasswordBaru.UseSystemPasswordChar = True
            End If
            Me.ShowIcon = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim Password As String = Me.tbPasswordLama.Text
        Me.tbKonfirmasiPassword.Text = Me.tbKonfirmasiPassword.Text.Trim
        Me.tbPasswordBaru.Text = Me.tbPasswordBaru.Text.Trim
        If ModUser.GantiPassword(Me.tbKodeUser, Me.tbPasswordLama, Me.tbPasswordBaru, Me.tbKonfirmasiPassword, Me.tbKodeUser.Text) = True Then
            'konfirmasi
            MsgBox("Password changed", MsgBoxStyle.Information, "Information")
            Me.Close()
        Else
            Me.tbPasswordLama.Text = Password
        End If
    End Sub

    Private Sub btnKeluar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnKeluar.Click
        Me.Close()
    End Sub

    Private Sub tbPasswordLama_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbPasswordLama.GotFocus
        Try
            Me.tbPasswordLama.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub tbPasswordLama_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbPasswordLama.LostFocus
        Try
            Me.tbPasswordLama.BackColor = Me.TxtBackColor.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub tbKonfirmasiPassword_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbKonfirmasiPassword.GotFocus
        Try
            Me.tbKonfirmasiPassword.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub tbKonfirmasiPassword_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbKonfirmasiPassword.LostFocus
        Try
            Me.tbKonfirmasiPassword.BackColor = Me.TxtBackColor.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub tbPasswordBaru_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbPasswordBaru.GotFocus
        Try
            Me.tbPasswordBaru.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub tbPasswordBaru_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbPasswordBaru.LostFocus
        Try
            Me.tbPasswordBaru.BackColor = Me.TxtBackColor.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class