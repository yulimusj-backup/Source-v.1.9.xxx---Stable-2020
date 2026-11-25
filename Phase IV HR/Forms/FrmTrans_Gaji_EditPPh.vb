Public Class FrmTrans_Gaji_EditPPh
    Dim ObjCD As New Core_Module.ClsCekData
    Dim BackColorAll As Drawing.Color = Core_Module.GetSMESSBackColorAll

    Private Sub TxtPPHNonRutin_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPPHNonRutin.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtPPHNonRutin, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPPHNonRutin_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPPHNonRutin.GotFocus
        Try
            Me.TxtPPHNonRutin.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPPHNonRutin_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPPHNonRutin.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtPPHNonRutin, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPPHNonRutin_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPPHNonRutin.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtPPHNonRutin, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPPHNonRutin_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPPHNonRutin.Leave
        Try
            Me.ObjCD.Number_Leave(Me.TxtPPHNonRutin, sender, e, 0, Me.ObjCD)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPPHNonRutin_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPPHNonRutin.LostFocus
        Try
            Me.TxtPPHNonRutin.BackColor = Me.TxtRef.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub TxtPPHRutin_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPPhRutin.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtPPhRutin, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPPHRutin_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPPhRutin.GotFocus
        Try
            Me.TxtPPhRutin.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPPHRutin_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPPhRutin.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtPPhRutin, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPPHRutin_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPPhRutin.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtPPhRutin, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPPHRutin_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPPhRutin.Leave
        Try
            Me.ObjCD.Number_Leave(Me.TxtPPhRutin, sender, e, 0, Me.ObjCD)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPPHRutin_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPPhRutin.LostFocus
        Try
            Me.TxtPPhRutin.BackColor = Me.TxtRef.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            Me.DialogResult = Windows.Forms.DialogResult.OK
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_Gaji_EditPPh_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            '#sinkron2018_pphcls
            Me.ShowIcon = False
            '#sinkron2018_pphcls-end
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class