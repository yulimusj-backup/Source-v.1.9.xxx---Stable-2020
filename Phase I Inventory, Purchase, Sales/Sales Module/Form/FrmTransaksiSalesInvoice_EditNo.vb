Option Explicit On
Option Strict Off

Imports System.Windows.Forms

Public Class FrmTransaksiSalesInvoice_EditNo
    Public MinLength As Integer = 0
    Public Aktif As Boolean = True
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    'update ory
    Public TanggalTransaksi As Date
    Public exPrimary As Double

    'update ory
    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Try
            If Me.mtb_Number.Text.Length <> Me.MinLength Then
                MsgBox("Must a number in " + Me.MinLength.ToString + " character.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If Not ApakahNomorBenar() Then
                Exit Sub
            End If

            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub mtb_Number_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mtb_Number.GotFocus
        Try
            Me.mtb_Number.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub mtb_Number_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mtb_Number.LostFocus
        Try
            Me.mtb_Number.BackColor = Me.MaskedTextBox1.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRevision_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRevision.GotFocus
        Try
            Me.TxtRevision.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRevision_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRevision.LostFocus
        Try
            Me.TxtRevision.BackColor = Me.MaskedTextBox1.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'update ory
    Private Function ApakahNomorBenar() As Boolean
        Dim objInv As New Sales_Module.ClsTransaksiSalesInvoice
        Try
            If objInv.apakahNoUrutSudahAda(exPrimary, CInt(mtb_Number.Text), TanggalTransaksi) Then
                MsgBox("This number have been used by other Sales Invoice Transaction on this numbering " + _
                    "reset period (Global Setting). Please enter another number.", _
                    MsgBoxStyle.OkOnly, "Information")
                mtb_Number.Focus()
                mtb_Number.SelectAll()
                Exit Function
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            objInv = Nothing
        End Try
    End Function
End Class
