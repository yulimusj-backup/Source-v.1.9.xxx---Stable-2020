Public Class FrmSetting_QuickInputAdj
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Dim ObjCD As New Core_Module.ClsCekData

    Public Function apakahDataBenar() As Boolean
        Try
            If Me.txtBatchNo.Text = "" Then
                MsgBox("Batch no cannot empty string.", MsgBoxStyle.Information, "Information")
                Me.txtBatchNo.Focus()
                Return False
            End If

            If ObjCD.CekBilanganPositif(Me.txtLimit.Text) = False Then
                MsgBox("Please enter positive number for Limit.", MsgBoxStyle.Information, "Information")
                Me.txtLimit.Focus()
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim Rs As MsgBoxResult
        Try
            If Me.apakahDataBenar = True Then
                Rs = MsgBox("Quick input to datagrid from batch no : " & Me.txtBatchNo.Text & " and limit " & Me.txtLimit.Text & " ?", MsgBoxStyle.YesNo, "Information")
                If Rs = MsgBoxResult.Yes Then
                    Me.DialogResult = Windows.Forms.DialogResult.Yes
                Else
                    Me.DialogResult = Windows.Forms.DialogResult.No
                End If
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtBatchNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBatchNo.GotFocus
        txtBatchNo.BackColor = BackColorAll
    End Sub

    Private Sub txtBatchNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBatchNo.LostFocus
        txtBatchNo.BackColor = Me.TextBox1.BackColor
    End Sub

    Private Sub txtLimit_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLimit.GotFocus
        txtLimit.BackColor = BackColorAll
    End Sub

    Private Sub txtLimit_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtLimit.KeyDown
        Try
            ObjCD.Number_KeyDown(Me.txtLimit, sender, e, 0)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtLimit_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLimit.LostFocus
        txtLimit.BackColor = Me.TextBox1.BackColor
    End Sub
End Class