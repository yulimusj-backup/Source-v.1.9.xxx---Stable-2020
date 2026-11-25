Public Class FrmTransaksiSalesInvoiceQuick_DetailNumber
    Public PrimaryInvQuickAcuan As Double = -1
    Public NoDO_Lama As String = ""
    Public NoSO_Lama As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim ObjDO As New ClsTransaksiDO
    Dim ObjSO As New ClsTransaksiSalesOrder

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            Me.TxtDONumber.Text = Me.TxtDONumber.Text.Trim
            If Me.TxtDONumber.Text = "" Then
                MsgBox("Please specify DO Number.", MsgBoxStyle.Information, "Information")
                Me.TxtDONumber.Focus()
                Exit Sub
            End If

            If Me.ObjDO.ApakahSudahAda(Me.TxtDONumber.Text, Me.PrimaryInvQuickAcuan) = True Then
                MsgBox("DO Number duplication, please specify different DO number.", MsgBoxStyle.Information, "Infomation")
                Me.TxtDONumber.Focus()
                Exit Sub
            End If

            Me.TxtSONumber.Text = Me.TxtSONumber.Text.Trim

            Me.DialogResult = Windows.Forms.DialogResult.OK
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            Me.DialogResult = Windows.Forms.DialogResult.Cancel
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPONumber_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDONumber.GotFocus, TxtSONumber.GotFocus
        Dim ij As TextBox

        Try
            ij = TryCast(sender, TextBox)
            If IsNothing(ij) = False Then
                ij.BackColor = Me.BackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try


        'Me.TxtDONumber.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtPONumber_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDONumber.LostFocus, TxtSONumber.LostFocus
        Dim ij As TextBox

        Try
            ij = TryCast(sender, TextBox)
            If IsNothing(ij) = False Then
                ij.BackColor = Me.TextBox1.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
        'Me.TxtDONumber.BackColor = Me.TextBox1.BackColor
    End Sub

    Private Sub FrmTransaksiAOP_Quick_DetailNumber_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.TxtDONumber.Text = Me.NoDO_Lama
        Me.TxtSONumber.Text = Me.NoSO_Lama
        Me.TxtDONumber.Focus()
        Me.ShowIcon = False
    End Sub
End Class