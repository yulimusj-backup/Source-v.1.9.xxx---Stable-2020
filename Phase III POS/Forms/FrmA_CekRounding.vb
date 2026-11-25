Option Explicit On
Option Strict Off
Imports Core_Module

Public Class FrmA_CekRounding
    Public UserCode As String = ""
    Public BackColorAll As Color = Color.LightGoldenrodYellow
    Public Bahasa As Core_Module.ModGlobalVariable.EnumBahasa = Core_Module.EnumBahasa.Inggris
    Dim ObjCD As New ClsCekData
    Dim Buffer As String
    Dim vLoad As Boolean = False

    Private Sub FrmA_CekRounding_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            vLoad = True
            Me.cboRoundingTipe.SelectedIndex = 0
            Me.cboRoundingValue.SelectedIndex = 6
            Me.TxtBaseAmount.Text = "0"
            Me.TxtDecimal.Text = "0"
            Me.TxtDiscount.Text = "0"
            Me.TxtDiscountAmount.Text = "0"
            Me.TxtNetAmount.Text = "0"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscount_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscount.Enter
        Try
            If CDec(Me.Buffer) > 0 Then
                Me.Buffer = Me.TxtDiscount.Text
            Else
                Me.Buffer = "0"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscount_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscount.Leave
        Try
            ObjCD.Disc_Leave(Me.TxtDiscount, Me.TxtDiscount.Text, sender, e, Me.TxtDecimal.Text, ObjCD, Me.Buffer)
            Me.TxtDiscountAmount.Text = ObjCD.Get_Disc_Amount(Me.TxtDiscount.Text, Me.TxtBaseAmount.Text, _
            Me.TxtDecimal.Text, Me.cboRoundingTipe.SelectedIndex, CDec(Me.cboRoundingValue.Text), ObjCD)
            If Me.cboRoundingTipe.SelectedIndex = 0 Then
                Me.TxtNetAmount.Text = ObjCD.DigitGrouping(Me.TxtBaseAmount.Text - Me.TxtDiscountAmount.Text, Me.TxtDecimal.Text, Me.cboRoundingTipe.SelectedIndex)
            ElseIf Me.cboRoundingTipe.SelectedIndex = 1 Then
                Me.TxtNetAmount.Text = ObjCD.DigitGrouping(Me.TxtBaseAmount.Text - Me.TxtDiscountAmount.Text, Me.TxtDecimal.Text, Me.cboRoundingTipe.SelectedIndex)
            ElseIf Me.cboRoundingTipe.SelectedIndex = 2 Then
                Me.TxtNetAmount.Text = ObjCD.DigitGrouping(Me.TxtBaseAmount.Text - Me.TxtDiscountAmount.Text, Me.TxtDecimal.Text)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscount_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDiscount.KeyDown
        Try
            ObjCD.Disc_KeyDown(TxtDiscount, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscount_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDiscount.KeyUp
        Try
            ObjCD.Disc_KeyUp(Me.TxtDiscount, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    '------
    Private Sub TxtBaseAmount_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBaseAmount.Enter
        Try
            ObjCD.Number_Enter(Me.TxtBaseAmount, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBaseAmount_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtBaseAmount.KeyDown
        Try
            ObjCD.Number_KeyDown(Me.TxtBaseAmount, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBaseAmount_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtBaseAmount.KeyUp
        Try
            ObjCD.Number_KeyUp(Me.TxtBaseAmount, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtBaseAmount_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtBaseAmount.Leave
        Try
            ObjCD.Number_Leave(Me.TxtBaseAmount, sender, e, Me.TxtDecimal.Text, ObjCD)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDecimal_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDecimal.KeyDown
        Try
            ObjCD.Number_KeyDown(Me.TxtDecimal, sender, e, 6)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDecimal_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDecimal.KeyUp
        Try
            ObjCD.Number_KeyUp(Me.TxtDecimal, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDecimal_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDecimal.Leave
        Try
            Me.TxtBaseAmount.Text = ObjCD.DigitGrouping(Me.TxtBaseAmount.Text, Me.TxtDecimal.Text)
            Me.TxtDiscountAmount.Text = ObjCD.Get_Disc_Amount(Me.TxtDiscount.Text, Me.TxtBaseAmount.Text, _
            Me.TxtDecimal.Text, Me.cboRoundingTipe.SelectedIndex, CDec(Me.cboRoundingValue.Text), ObjCD)
            If Me.cboRoundingTipe.SelectedIndex = 0 Then
                Me.TxtNetAmount.Text = ObjCD.DigitGrouping(Me.TxtBaseAmount.Text - Me.TxtDiscountAmount.Text, Me.TxtDecimal.Text, Me.cboRoundingTipe.SelectedIndex)
            ElseIf Me.cboRoundingTipe.SelectedIndex = 1 Then
                Me.TxtNetAmount.Text = ObjCD.DigitGrouping(Me.TxtBaseAmount.Text - Me.TxtDiscountAmount.Text, Me.TxtDecimal.Text, Me.cboRoundingTipe.SelectedIndex)
            ElseIf Me.cboRoundingTipe.SelectedIndex = 2 Then
                Me.TxtNetAmount.Text = ObjCD.DigitGrouping(Me.TxtBaseAmount.Text - Me.TxtDiscountAmount.Text, Me.TxtDecimal.Text)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboRoundingValue_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboRoundingValue.SelectedIndexChanged
        Try
            Me.TxtDiscountAmount.Text = ObjCD.Get_Disc_Amount(Me.TxtDiscount.Text, Me.TxtBaseAmount.Text, _
            Me.TxtDecimal.Text, Me.cboRoundingTipe.SelectedIndex, CDec(Me.cboRoundingValue.Text), ObjCD)
            If Me.cboRoundingTipe.SelectedIndex = 0 Then
                Me.TxtNetAmount.Text = ObjCD.DigitGrouping(Me.TxtBaseAmount.Text - Me.TxtDiscountAmount.Text, Me.TxtDecimal.Text, Me.cboRoundingTipe.SelectedIndex)
            ElseIf Me.cboRoundingTipe.SelectedIndex = 1 Then
                Me.TxtNetAmount.Text = ObjCD.DigitGrouping(Me.TxtBaseAmount.Text - Me.TxtDiscountAmount.Text, Me.TxtDecimal.Text, Me.cboRoundingTipe.SelectedIndex)
            ElseIf Me.cboRoundingTipe.SelectedIndex = 2 Then
                Me.TxtNetAmount.Text = ObjCD.DigitGrouping(Me.TxtBaseAmount.Text - Me.TxtDiscountAmount.Text, Me.TxtDecimal.Text)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboRoundingTipe_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboRoundingTipe.SelectedIndexChanged
        Try
            If vLoad = True Then
                Me.cboRoundingValue.SelectedIndex = 6
                vLoad = False
            Else
                Me.TxtDiscountAmount.Text = ObjCD.Get_Disc_Amount(Me.TxtDiscount.Text, Me.TxtBaseAmount.Text, _
                Me.TxtDecimal.Text, Me.cboRoundingTipe.SelectedIndex, CDec(Me.cboRoundingValue.Text), ObjCD)
                If Me.cboRoundingTipe.SelectedIndex = 0 Then
                    Me.TxtNetAmount.Text = ObjCD.DigitGrouping(Me.TxtBaseAmount.Text - Me.TxtDiscountAmount.Text, Me.TxtDecimal.Text, Me.cboRoundingTipe.SelectedIndex)
                ElseIf Me.cboRoundingTipe.SelectedIndex = 1 Then
                    Me.TxtNetAmount.Text = ObjCD.DigitGrouping(Me.TxtBaseAmount.Text - Me.TxtDiscountAmount.Text, Me.TxtDecimal.Text, Me.cboRoundingTipe.SelectedIndex)
                ElseIf Me.cboRoundingTipe.SelectedIndex = 2 Then
                    Me.TxtNetAmount.Text = ObjCD.DigitGrouping(Me.TxtBaseAmount.Text - Me.TxtDiscountAmount.Text, Me.TxtDecimal.Text)
                End If
                vLoad = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Cetak_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cetak.Click
        Dim Obj As New ClsTrans_POS
        Dim ij As New FrmA_Cetak
        Try
            Obj.CetakCobaCoba(ij.CRV1)
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Obj = Nothing
            ij.Dispose()
        End Try
    End Sub
End Class