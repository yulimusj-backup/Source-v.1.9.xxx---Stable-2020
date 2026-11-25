Option Explicit On
Option Strict Off
Imports System.Windows.Forms, System.Drawing

Public Class FrmAddReplace
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjCurr As New Core_Module.clsMasterCurrency
    Dim Def_Prec As Double = Me.ObjCurr.GetDecimal_DefaultCurrency()

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Try
            If Me.ApakahDataBenar() = True Then
                Me.DialogResult = System.Windows.Forms.DialogResult.OK
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Try
            If Me.ApakahDataBenar() = True Then
                Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            Me.DialogResult = Windows.Forms.DialogResult.No
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function ApakahDataBenar() As Boolean
        Try
            If Me.rboFixedValue.Checked = True Then
                If Me.ObjCD.CekBilanganPositif(Me.TxtOverrideAmount.Text) = False Then
                    MsgBox("Please input positive number.", MsgBoxStyle.Information, "Information")
                    Me.TxtOverrideAmount.Focus()
                    Me.DialogResult = Windows.Forms.DialogResult.No
                    Return False
                End If
            ElseIf Me.rboByBS.Checked = True Then
                If Me.ObjCD.CekBilanganPositif(Me.TxtBS.Text) = False Then
                    MsgBox("Please input positive number.", MsgBoxStyle.Information, "Information")
                    Me.DialogResult = Windows.Forms.DialogResult.No
                    Me.TxtBS.Focus()
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub TxtOverrideAmount_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtOverrideAmount.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtOverrideAmount, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtOverrideAmount_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtOverrideAmount.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtOverrideAmount, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtOverrideAmount_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtOverrideAmount.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtOverrideAmount, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtOverrideAmount_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtOverrideAmount.Leave
        Try
            Me.ObjCD.Number_Leave(Me.TxtOverrideAmount, sender, e, Me.Def_Prec, Me.ObjCD)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtOverrideAmount_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtOverrideAmount.GotFocus
        Try
            If Me.TxtRemarksItem.Enabled = True Then
                If Me.TxtOverrideAmount.Enabled = True Then
                    If Me.TxtOverrideAmount.ReadOnly = False Then
                        TxtOverrideAmount.BackColor = Me.BackColorAll
                    Else
                        TxtOverrideAmount.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtOverrideAmount_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtOverrideAmount.LostFocus
        Try
            If Me.TxtRemarksItem.Enabled = True Then
                If Me.TxtOverrideAmount.Enabled = True Then
                    If Me.TxtOverrideAmount.ReadOnly = False Then
                        TxtOverrideAmount.BackColor = Me.TxtPrimaryLoanPayment.BackColor
                    Else
                        TxtOverrideAmount.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarksItem_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarksItem.GotFocus
        Try
            If Me.TxtRemarksItem.Enabled = True Then
                If Me.TxtRemarksItem.Enabled = True Then
                    If Me.TxtRemarksItem.ReadOnly = False Then
                        TxtRemarksItem.BackColor = Me.BackColorAll
                    Else
                        TxtRemarksItem.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarksItem_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarksItem.LostFocus
        Try
            If Me.TxtRemarksItem.Enabled = True Then
                If Me.TxtRemarksItem.Enabled = True Then
                    If Me.TxtRemarksItem.ReadOnly = False Then
                        TxtRemarksItem.BackColor = Me.TxtPrimaryLoanPayment.BackColor
                    Else
                        TxtRemarksItem.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmAddReplace_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub rboFixedValue_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rboFixedValue.CheckedChanged
        Try
            If Me.rboFixedValue.Checked = True Then
                Me.TxtOverrideAmount.Enabled = True
                Me.TxtBS.Enabled = False
                Me.TxtBS.Text = "1.00"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub rboByBS_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rboByBS.CheckedChanged
        Try
            If Me.rboByBS.Checked = True Then
                Me.TxtBS.Enabled = True
                Me.TxtOverrideAmount.Text = "0"
                Me.TxtOverrideAmount.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBS_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtBS.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtBS, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBS_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtBS.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtBS, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBS_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBS.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtBS, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBS_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBS.Leave
        Try
            Me.ObjCD.Number_Leave(Me.TxtBS, sender, e, 2, Me.ObjCD)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBS_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBS.GotFocus
        Try
            If Me.TxtRemarksItem.Enabled = True Then
                If Me.TxtBS.Enabled = True Then
                    If Me.TxtBS.ReadOnly = False Then
                        TxtBS.BackColor = Me.BackColorAll
                    Else
                        TxtBS.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBS_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBS.LostFocus
        Try
            If Me.TxtRemarksItem.Enabled = True Then
                If Me.TxtBS.Enabled = True Then
                    If Me.TxtBS.ReadOnly = False Then
                        TxtBS.BackColor = Me.TxtPrimaryLoanPayment.BackColor
                    Else
                        TxtBS.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
