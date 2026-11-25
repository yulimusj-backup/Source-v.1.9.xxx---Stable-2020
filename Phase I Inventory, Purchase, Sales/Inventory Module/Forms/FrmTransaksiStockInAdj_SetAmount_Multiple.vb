Option Explicit On
Option Strict Off

Public Class FrmTransaksiStockInAdj_SetAmount_Multiple

    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public DefPrec As Integer = 0

    Dim objCD As New Core_Module.ClsCekData

    Private Sub Txt_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitPrice.GotFocus
        Dim Tx As Windows.Forms.TextBox
        Try
            Tx = TryCast(sender, Windows.Forms.TextBox)
            If IsNothing(Tx) = False Then
                Tx.BackColor = Me.BackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Txt_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitPrice.LostFocus
        Dim Tx As Windows.Forms.TextBox

        Try
            Tx = TryCast(sender, Windows.Forms.TextBox)
            If IsNothing(Tx) = False Then
                Tx.BackColor = Color.White
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Txt_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitPrice.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtUnitPrice, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUnitPrice_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtUnitPrice.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtUnitPrice, sender, e, 0)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUnitPrice_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtUnitPrice.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtUnitPrice, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUnitPrice_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitPrice.Leave
        Try
            Me.ObjCD.Number_Leave(Me.TxtUnitPrice, sender, e, Me.DefPrec, Me.ObjCD)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnCancel.Click
        Try
            Me.DialogResult = Windows.Forms.DialogResult.Cancel
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiStockInAdj_SetAmount_Multiple_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.DialogResult = Windows.Forms.DialogResult.None
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnOK.Click
        Try
            If MsgBox("Are you sure?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                Me.DialogResult = Windows.Forms.DialogResult.OK
                Me.Close()
            Else
                Me.DialogResult = Windows.Forms.DialogResult.None
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class