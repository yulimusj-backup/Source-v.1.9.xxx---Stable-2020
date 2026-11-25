Public Class FrmMaster_EditNPWPLama
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Private Sub BtnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnOK.Click
        Dim Dialog As MsgBoxResult
        Try
            Dialog = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirmation")
            If Dialog = MsgBoxResult.Yes Then
                If Me.TxtNPWPNo.Text = "" Then
                    MsgBox("NPWP No cannot empty string.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
                DialogResult = Windows.Forms.DialogResult.OK
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnCancel.Click
        Try
            DialogResult = Windows.Forms.DialogResult.Cancel
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMaster_EditNPWPLama_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function BlokSimbol(ByVal obj As TextBox, ByVal e As Windows.Forms.KeyPressEventArgs) As Boolean
        Try
            If e.KeyChar = "'" Or _
            e.KeyChar = "%" Or _
            e.KeyChar = "*" Or _
            e.KeyChar = "#" Then
                obj.ReadOnly = True
                obj.BackColor = Me.BackColorAll
            Else
                obj.ReadOnly = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub TxtNPWPNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNPWPNo.GotFocus
        TxtNPWPNo.BackColor = BackColorAll
    End Sub

    Private Sub TxtNPWPNo_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtNPWPNo.KeyPress
        Me.BlokSimbol(Me.TxtNPWPNo, e)
    End Sub

    Private Sub TxtNPWPNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNPWPNo.LostFocus
        TxtNPWPNo.BackColor = Txt.BackColor
    End Sub
End Class