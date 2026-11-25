Option Explicit On
Option Strict Off
Public Class FrmTransaksiQuotation_SetQtyRatio

    Public vRatioMulti As Double = 0
    Public vDefPrecMulti As Integer = 0
    Public vUnitMulti As String = "Unit"
    Public vRatioDefault As Double = 0
    Public vDefPrecDefault As Integer = 0
    Public vUnitDefault As String = "Unit"

    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim objCd As New Core_Module.ClsCekData
    Dim Buffer As String = ""

    Private Sub FrmTransaksiPO_SetQtyRatio_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.TxtRatioMulti.Text = Me.objCd.DigitGrouping(vRatioMulti, vDefPrecMulti)
            Me.LlbUnitMulti.Text = vUnitMulti
            Me.TxtRatioDefault.Text = Me.objCd.DigitGrouping(vRatioDefault, vDefPrecDefault)
            Me.LlbUnitDefault.Text = vUnitDefault
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnCancel.Click
        Me.DialogResult = Windows.Forms.DialogResult.Cancel
    End Sub

    Private Sub BtnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnOK.Click
        Me.DialogResult = Windows.Forms.DialogResult.OK
    End Sub

    Private Sub TxtRatioMulti_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRatioMulti.Enter
        Try
            Buffer = Me.TxtRatioMulti.Text
            Me.objCd.Number_Enter(Me.TxtRatioMulti, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRatioMulti_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRatioMulti.GotFocus
        Try
            TxtRatioMulti.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRatioMulti_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtRatioMulti.KeyDown
        Try
            Me.objCd.Number_KeyDown(Me.TxtRatioMulti, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRatioMulti_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtRatioMulti.KeyUp
        Try
            Me.objCd.Number_KeyUp(Me.TxtRatioMulti, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRatioMulti_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRatioMulti.Leave
        Try
            If Me.objCd.CekBilanganPositif(TxtRatioMulti.Text) = False Then
                MsgBox("Please enter positive number for ratio data.", MsgBoxStyle.Information, "Information")
                TxtRatioMulti.Text = Me.Buffer
                Exit Sub
            End If
            Me.TxtRatioMulti.Text = Me.objCd.DigitGrouping(Me.TxtRatioMulti.Text, vDefPrecMulti)
            If CDec(Me.TxtRatioMulti.Text) = 0 Then
                MsgBox("Please input positive number for ratio data.", MsgBoxStyle.Information, "Information")
                TxtRatioMulti.Text = Me.Buffer
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    '------------------------------------------------------------------
    Private Sub TxtRatioDefault_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRatioDefault.Enter
        Try
            Buffer = Me.TxtRatioDefault.Text
            Me.objCd.Number_Enter(Me.TxtRatioDefault, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRatioDefault_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRatioDefault.GotFocus
        Try
            TxtRatioDefault.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRatioDefault_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtRatioDefault.KeyDown
        Try
            Me.objCd.Number_KeyDown(Me.TxtRatioDefault, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRatioDefault_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtRatioDefault.KeyUp
        Try
            Me.objCd.Number_KeyUp(Me.TxtRatioDefault, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRatioDefault_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRatioDefault.Leave
        Try
            If Me.objCd.CekBilanganPositif(TxtRatioDefault.Text) = False Then
                MsgBox("Please enter positive number for ratio data.", MsgBoxStyle.Information, "Information")
                TxtRatioDefault.Text = Me.Buffer
                Exit Sub
            End If
            Me.TxtRatioDefault.Text = Me.objCd.DigitGrouping(Me.TxtRatioDefault.Text, vDefPrecDefault)
            If CDec(Me.TxtRatioDefault.Text) = 0 Then
                MsgBox("Please input positive number for ratio data.", MsgBoxStyle.Information, "Information")
                TxtRatioMulti.Text = Me.Buffer
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRatioDefault_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRatioDefault.LostFocus
        Try
            TxtRatioDefault.BackColor = Me.TxtPrimaryCodeNo.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRatioMulti_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRatioMulti.LostFocus
        Try
            TxtRatioMulti.BackColor = Me.TxtPrimaryCodeNo.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class