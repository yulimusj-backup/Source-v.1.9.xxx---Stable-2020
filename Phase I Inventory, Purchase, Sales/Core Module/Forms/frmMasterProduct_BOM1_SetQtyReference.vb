Option Explicit On
Option Strict Off

Public Class frmMasterProduct_BOM1_SetQtyReference

    Public QtyReference As Decimal = 0
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public Decimal_ As Integer = 2
    Dim objCD As New Core_Module.ClsCekData

    Private Sub txtQtyReference_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtQtyReference.Enter
        Try
            Me.objCD.Number_Enter(Me.txtQtyReference, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtQtyReference_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtQtyReference.GotFocus
        Try
            Me.txtQtyReference.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtQtyReference_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtQtyReference.KeyDown
        Try
            Me.objCD.Number_KeyDown(Me.txtQtyReference, sender, e, 0)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtQtyReference_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtQtyReference.KeyUp
        Try
            Me.objCD.Number_KeyUp(Me.txtQtyReference, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtQtyReference_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtQtyReference.Leave
        Try
            Me.objCD.Number_Leave(Me.txtQtyReference, sender, e, Decimal_, Me.objCD)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtQtyReference_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtQtyReference.LostFocus
        Try
            Me.txtQtyReference.BackColor = Color.White
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            If CDec(Me.txtQtyReference.Text) <= 0 Then
                If Me.lblUnit.Text <> "%" Then
                    MsgBox(lblQtyReference.Text & " must be a positive number.", MsgBoxStyle.Information, "Information")
                Else
                    MsgBox(lblQtyReference.Text & " must be a positive number.", MsgBoxStyle.Information, "Information")
                End If
                Me.txtQtyReference.Focus()
                Exit Sub
            End If
            If Me.lblUnit.Text = "%" And CDec(Me.txtQtyReference.Text) > 100 Then
                MsgBox("Maximum good percentage is 100 %.", MsgBoxStyle.Information, "Information")
                Me.txtQtyReference.Focus()
                Exit Sub
            End If
            Me.QtyReference = Me.txtQtyReference.Text
            Me.DialogResult = Windows.Forms.DialogResult.OK
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub frmMasterProduct_BOM1_SetQtyReference_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.DialogResult = Windows.Forms.DialogResult.None
            Me.txtQtyReference.Text = Me.objCD.DigitGrouping(0, Me.Decimal_)
            Me.txtQtyReference.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            Me.DialogResult = Windows.Forms.DialogResult.Cancel
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class