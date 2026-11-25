Option Explicit On
Option Strict Off

Public Class FrmTrans_MaterialPlanningDashboard_MP_DetailTolerance
    Public vQtyOrder As Decimal = 0
    Public vQtyUpperTolerance As Decimal = 0
    Public vQtyLowerTolerance As Decimal = 0
    Public vUnitCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public vDefPrec As Decimal = 0

    Dim buffer As String = ""
    Dim ObjCD As New Core_Module.ClsCekData
    Dim MaxTolerance As Decimal = 0
    Dim vAllowClose As Boolean = False

    Private Sub TxtPerct_UpperTolerance_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPerct_UpperTolerance.Enter
        Try
            Me.buffer = Me.TxtPerct_UpperTolerance.Text
            Me.ObjCD.Number_Enter(Me.TxtPerct_UpperTolerance, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPerct_UpperTolerance_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPerct_UpperTolerance.GotFocus
        Try
            Me.TxtPerct_UpperTolerance.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPerct_UpperTolerance_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPerct_UpperTolerance.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtPerct_UpperTolerance, sender, e, 50)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPerct_UpperTolerance_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPerct_UpperTolerance.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtPerct_UpperTolerance, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPerct_UpperTolerance_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPerct_UpperTolerance.Leave
        Try
            If CDec(Me.TxtPerct_UpperTolerance.Text) <> CDec(Me.buffer) Then
                Me.ObjCD.Number_Leave(Me.TxtPerct_UpperTolerance, sender, e, 2, Me.ObjCD)
                Me.TxtQty_UpperTolerance.Text = Me.ObjCD.DigitGrouping((CDec(Me.TxtQty_Order.Text) * CDec(Me.TxtPerct_UpperTolerance.Text)) / 100, Me.vDefPrec)
            Else
                Me.TxtPerct_UpperTolerance.Text = Me.buffer
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPerct_UpperTolerance_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPerct_UpperTolerance.LostFocus
        Try
            Me.TxtPerct_UpperTolerance.BackColor = Me.TxtPrimary.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPerct_LowerTolerance_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPerct_LowerTolerance.Enter
        Try
            Me.buffer = Me.TxtPerct_LowerTolerance.Text
            Me.ObjCD.Number_Enter(Me.TxtPerct_LowerTolerance, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPerct_LowerTolerance_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPerct_LowerTolerance.GotFocus
        Try
            Me.TxtPerct_LowerTolerance.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPerct_LowerTolerance_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPerct_LowerTolerance.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtPerct_LowerTolerance, sender, e, 50)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPerct_LowerTolerance_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtPerct_LowerTolerance.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtPerct_LowerTolerance, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPerct_LowerTolerance_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPerct_LowerTolerance.Leave
        Try
            If CDec(Me.TxtPerct_LowerTolerance.Text) <> CDec(Me.buffer) Then
                Me.ObjCD.Number_Leave(Me.TxtPerct_LowerTolerance, sender, e, 2, Me.ObjCD)
                Me.TxtQty_LowerTolerance.Text = Me.ObjCD.DigitGrouping((CDec(Me.TxtQty_Order.Text) * CDec(Me.TxtPerct_LowerTolerance.Text)) / 100, Me.vDefPrec)
            Else
                Me.TxtPerct_LowerTolerance.Text = Me.buffer
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPerct_LowerTolerance_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPerct_LowerTolerance.LostFocus
        Try
            Me.TxtPerct_LowerTolerance.BackColor = Me.TxtPrimary.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_MaterialPlanningDashboard_MP_DetailTolerance_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.vAllowClose = False Then
                If MsgBox("Cancel data changes?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Me.DialogResult = Windows.Forms.DialogResult.Cancel
                    Me.vAllowClose = True
                    Me.Close()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_MaterialPlanningDashboard_MP_DetailTolerance_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.ShowIcon = False
            Me.TxtQty_Order.Text = Me.ObjCD.DigitGrouping(Me.vQtyOrder, Me.vDefPrec)
            Me.LblUnit.Text = Me.vUnitCode
            Me.LblUnit1.Text = Me.LblUnit.Text
            Me.LblUnit2.Text = Me.LblUnit.Text
            Me.TxtQty_UpperTolerance.Text = Me.ObjCD.DigitGrouping(Me.vQtyUpperTolerance, Me.vDefPrec)
            Me.TxtQty_LowerTolerance.Text = Me.ObjCD.DigitGrouping(Me.vQtyLowerTolerance, Me.vDefPrec)
            Me.MaxTolerance = Math.Round(CDec(0.5 * Me.TxtQty_Order.Text), CInt(Me.vDefPrec), MidpointRounding.AwayFromZero)
            If CDec(Me.vQtyOrder) = 0 Then
                Me.TxtPerct_UpperTolerance.Text = Me.ObjCD.DigitGrouping(0, 2)
                Me.TxtPerct_LowerTolerance.Text = Me.ObjCD.DigitGrouping(0, 2)
            Else
                Me.TxtPerct_UpperTolerance.Text = Me.ObjCD.DigitGrouping((Me.vQtyUpperTolerance * 100) / Me.vQtyOrder, 2)
                Me.TxtPerct_LowerTolerance.Text = Me.ObjCD.DigitGrouping((Me.vQtyLowerTolerance * 100) / Me.vQtyOrder, 2)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQty_UpperTolerance_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQty_UpperTolerance.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtQty_UpperTolerance, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQty_UpperTolerance_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQty_UpperTolerance.GotFocus
        Try
            Me.TxtQty_UpperTolerance.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQty_UpperTolerance_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQty_UpperTolerance.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtQty_UpperTolerance, sender, e, MaxTolerance)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQty_UpperTolerance_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQty_UpperTolerance.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtQty_UpperTolerance, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQty_UpperTolerance_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQty_UpperTolerance.Leave
        Try
            Me.ObjCD.Number_Leave(Me.TxtQty_UpperTolerance, sender, e, Me.vDefPrec, Me.ObjCD)
            Me.TxtPerct_UpperTolerance.Text = Me.ObjCD.DigitGrouping((CDec(Me.TxtQty_UpperTolerance.Text) * 100) / CDec(Me.TxtQty_Order.Text), Me.vDefPrec)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQty_UpperTolerance_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQty_UpperTolerance.LostFocus
        Try
            Me.TxtQty_UpperTolerance.BackColor = Me.TxtPrimary.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQty_LowerTolerance_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQty_LowerTolerance.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtQty_LowerTolerance, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQty_LowerTolerance_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQty_LowerTolerance.GotFocus
        Try
            Me.TxtQty_LowerTolerance.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQty_LowerTolerance_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQty_LowerTolerance.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtQty_LowerTolerance, sender, e, MaxTolerance)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQty_LowerTolerance_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQty_LowerTolerance.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtQty_LowerTolerance, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQty_LowerTolerance_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQty_LowerTolerance.Leave
        Try
            Me.ObjCD.Number_Leave(Me.TxtQty_LowerTolerance, sender, e, Me.vDefPrec, Me.ObjCD)
            Me.TxtPerct_LowerTolerance.Text = Me.ObjCD.DigitGrouping((CDec(Me.TxtQty_LowerTolerance.Text) * 100) / CDec(Me.TxtQty_Order.Text), Me.vDefPrec)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQty_LowerTolerance_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQty_LowerTolerance.LostFocus
        Try
            Me.TxtQty_LowerTolerance.BackColor = Me.TxtPrimary.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            Me.DialogResult = Windows.Forms.DialogResult.OK
            Me.vAllowClose = True
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            If MsgBox("Cancel data changes?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.DialogResult = Windows.Forms.DialogResult.Cancel
                Me.vAllowClose = True
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPerct_UpperTolerance_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtPerct_UpperTolerance.TextChanged

    End Sub

    Private Sub TxtPerct_LowerTolerance_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtPerct_LowerTolerance.TextChanged

    End Sub
End Class