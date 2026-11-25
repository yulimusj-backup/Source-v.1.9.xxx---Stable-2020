Public Class FrmMasterUser_Otorisasi
    Public AccessName As String = ""

    Dim ObjUser As New Core_Module.ClsUser
    Dim BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            If Me.ObjUser.Login(Me.cboUserCode.Text, Me.TxtPassword.Text) = True Then
                Me.DialogResult = Windows.Forms.DialogResult.OK
                Me.Close()
            Else
                Me.TxtPassword.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.DialogResult = Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub FrmMasterUser_Otorisasi_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            'isikan cbo sesuai dengan kode hak akses
            Me.ObjUser.IsiComboBox_ByHakAkses(Me.cboUserCode, Me.AccessName)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPassword_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPassword.GotFocus, TxtReason.GotFocus
        Dim Txt As TextBox

        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                Txt.BackColor = Me.BackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPassword_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPassword.LostFocus, TxtReason.LostFocus
        Dim Txt As TextBox

        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                Txt.BackColor = Me.TextBox1.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class