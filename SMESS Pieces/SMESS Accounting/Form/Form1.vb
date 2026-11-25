Public Class Form1
    Dim ObjCD As New Core_Module.ClsCekData

    Private Sub TextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown
        Me.ObjCD.Number_KeyDown(Me.TextBox1, sender, e)
    End Sub

    Private Sub TextBox1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyUp
        Me.ObjCD.Number_KeyUp(Me.TextBox1, sender, e, False)
    End Sub

    Private Sub TextBox1_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.Leave
        Dim ij As Core_Module.ClsCekData.EnumTipeRounding = Core_Module.ClsCekData.EnumTipeRounding.Down

        Try
            Me.TextBox2.Text = Me.ObjCD.GetRoundingValue(Me.TextBox1.Text, 0, ij, 50, ObjCD)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
      
    End Sub
End Class