Option Explicit On
Option Strict Off

Public Class FrmResetPassword
    Dim obj As New ClsUser

    Private Sub FrmResetPassword_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            'load semua data kode customer
            obj.IsiComboBox(Me.cboUser)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboUser_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboUser.SelectedIndexChanged
        Try
            If Me.cboUser.SelectedIndex <> -1 Then
                'cari data nama user
                Me.txtNamaUser.Text = obj.GetNama(Me.cboUser.SelectedItem)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            'Simpan perubahan data password
            If MsgBox("Anda yakin akan menyimpan data ini?", MsgBoxStyle.YesNo, "Konfirmasi") = MsgBoxResult.Yes Then
                Windows.Forms.Cursor.Current = Cursors.WaitCursor

                If obj.SetPassword(Me.cboUser.SelectedItem, Me.txtPasswordBaru, Me.txtKonfirmasiPassword) Then
                    Me.NonAktifkanDetail()
                    Me.KosongkanField()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub chkShowPassword_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowPassword.CheckedChanged
        Try
            If Me.chkShowPassword.Checked = True Then
                'tunjukkan password
                Me.txtKonfirmasiPassword.UseSystemPasswordChar = False
                Me.txtPasswordBaru.UseSystemPasswordChar = False
            Else
                'gunakan sistem password proc
                Me.txtKonfirmasiPassword.UseSystemPasswordChar = True
                Me.txtPasswordBaru.UseSystemPasswordChar = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AktifkanDetail()
        Try
            Me.chkShowPassword.Enabled = True
            ModCekData.KendalikanControl(Me.GroupBox1.Controls, True, True, True, True, False, False, False, False)
            ModCekData.KendalikanControl(Me.Controls, True, False, True, False, False, False, True, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            Me.chkShowPassword.Enabled = False
            ModCekData.KendalikanControl(Me.GroupBox1.Controls, True, False, True, False, False, False, False, False)
            ModCekData.KendalikanControl(Me.Controls, True, True, True, True, False, False, True, True)
            Me.CancelButton = Me.btnKeluar
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            Me.chkShowPassword.Checked = False
            Me.txtPasswordBaru.Clear()
            Me.txtKonfirmasiPassword.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnGanti_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGanti.Click
        Try
            If Me.cboUser.SelectedIndex <> -1 Then
                Me.AktifkanDetail()
                Me.KosongkanField()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnBatal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBatal.Click
        Try
            If MsgBox("Anda yakin akan batal menyimpan data ini?", MsgBoxStyle.YesNo, "Konfirmasi") = MsgBoxResult.Yes Then
                Me.NonAktifkanDetail()
                Me.KosongkanField()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnKeluar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnKeluar.Click
        Me.Close()
    End Sub
End Class