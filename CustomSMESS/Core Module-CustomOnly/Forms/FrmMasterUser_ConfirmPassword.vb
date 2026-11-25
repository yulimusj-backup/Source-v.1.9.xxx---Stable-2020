Public Class FrmMasterUser_ConfirmPassword
    Dim ObjUser As New Core_Module.ClsUser
    Dim Dummy As String
    Dim counter As Double = 0

    Private Sub FrmTransaksiSO_DialogInvoiceNunggak_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            'load orang - orang yang memiliki akses untuk melakukan oveeride invoice nunggka
            'orang-orang yang punya akses terhadap form payment
            Me.ObjUser.IsiComboBox(Me.cboUser)
            While counter < Me.cboUser.Items.Count
                If Me.ObjUser.ApakahPunyaHakAkses(Me.cboUser.Items(counter), "frmTransaksiPayment") = False Then
                    Me.cboUser.Items.RemoveAt(counter)
                Else
                    counter += 1
                End If
            End While
            If Me.cboUser.Items.Count > 0 Then
                Me.cboUser.SelectedIndex = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            'cek password
            If Core_Module.ModUser.Login(Me.TxtUserName, Me.TxtPassword) = True Then
                'jika benar maka dialog result = ok
                Me.DialogResult = Windows.Forms.DialogResult.OK
                Me.Close()
            Else
                Me.TxtPassword.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboUser_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboUser.SelectedIndexChanged
        Try
            Me.TxtUserName.Text = Me.cboUser.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPassword_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPassword.KeyDown
        Try
            If e.KeyValue = 13 Then
                Me.btnOK.PerformClick()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class