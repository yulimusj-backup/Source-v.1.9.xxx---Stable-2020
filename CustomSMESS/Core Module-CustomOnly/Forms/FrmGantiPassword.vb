Public Class FrmGantiPassword
    Public KodeUser As String = ""

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
            MsgBox("Password Anda sudah diganti", MsgBoxStyle.Information, "Konfirmasi")
            Me.Close()
        Else
            Me.tbPasswordLama.Text = Password
        End If
    End Sub

    Private Sub btnKeluar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnKeluar.Click
        Me.Close()
    End Sub
End Class