Option Explicit On
Option Strict Off

Public Class FrmTrans_ManufacturingOrder_Exp_RatioReplace
    Public vPrimaryProductReplaced As Double = -1
    Public vPrimaryProductReplacement As Double = -1
    Public vUnitCodeReplaced, vUnitCodeReplacement As String
    Public vQtyReplaced, vQtyReplacement, vRatioReplaced, vRatioReplacement As Decimal
    Public vDefPrecReplaced, vDefPrecReplacement As Integer
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim ObjCD As New Core_Module.ClsCekData
    Dim LangsungTutup As Boolean = False
    Dim ObjProduct As New Core_Module.clsMasterProduct 

    Private Sub FrmTrans_ManufacturingOrder_Exp_RatioReplace_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If LangsungTutup = False Then
                If MsgBox("Close form and discard data changes?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_ManufacturingOrder_Exp_RatioReplace_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            'RefLeftPosition = Me.TxtRatioReplaced.Left

            Me.TxtQtyReplaced.Text = Me.ObjCD.DigitGrouping(Me.vQtyReplaced, Me.vDefPrecReplaced)
            Me.TxtQtyReplacement.Text = Me.ObjCD.DigitGrouping(Me.vQtyReplacement, Me.vDefPrecReplacement)
            Me.TxtRatioReplaced.Text = Me.ObjCD.DigitGrouping(Me.vRatioReplaced, Me.vDefPrecReplaced)
            Me.TxtRatioReplacement.Text = Me.ObjCD.DigitGrouping(Me.vRatioReplacement, Me.vDefPrecReplacement)
            Me.LblRatio.Text = Me.vUnitCodeReplaced + " " + Me.ObjProduct.GetProductCode(Me.vPrimaryProductReplaced) + " replaced by "
            Me.LblUnitReplacementRatio.Text = Me.vUnitCodeReplacement + " " + Me.ObjProduct.GetProductCode(Me.vPrimaryProductReplacement)
            Me.TxtRatioReplacement.Left = 5 + Me.LblRatio.Width + Me.LblRatio.Left
            'Me.TxtRatioReplacement.Left = Me.RefLeftPosition
            'Me.TxtRatioReplaced.Left = 5 + Me.LblRatio.Width + Me.LblRatio.Left

            Me.LblUnitReplacementRatio.Left = 5 + Me.TxtRatioReplacement.Left + Me.TxtRatioReplacement.Width
            'Me.LblUnitReplacementRatio.Left = 5 + Me.TxtRatioReplaced.Left + Me.TxtRatioReplaced.Width

            Me.LblUnitReplaced.Text = Me.vUnitCodeReplaced
            Me.LblUnitReplacement.Text = Me.vUnitCodeReplacement
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRatioReplaced_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyReplaced.Enter, _
    TxtQtyReplacement.Enter, TxtRatioReplaced.Enter, TxtRatioReplacement.Enter
        Dim TxtBox As Windows.Forms.TextBox

        Try
            TxtBox = TryCast(sender, Windows.Forms.TextBox)
            If IsNothing(TxtBox) = False Then
                Me.ObjCD.Number_Enter(TxtBox, sender, e)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRatioReplaced_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyReplaced.GotFocus, _
    TxtQtyReplacement.GotFocus, TxtRatioReplaced.GotFocus, TxtRatioReplacement.GotFocus
        Dim TxtBox As Windows.Forms.TextBox

        Try
            TxtBox = TryCast(sender, Windows.Forms.TextBox)
            If IsNothing(TxtBox) = False Then
                TxtBox.BackColor = Me.BackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRatioReplaced_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyReplaced.KeyDown, _
    TxtQtyReplacement.KeyDown, TxtRatioReplaced.KeyDown, TxtRatioReplacement.KeyDown
        Dim TxtBox As Windows.Forms.TextBox

        Try
            TxtBox = TryCast(sender, Windows.Forms.TextBox)
            If IsNothing(TxtBox) = False Then
                Me.ObjCD.Number_KeyDown(TxtBox, sender, e)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRatioReplaced_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyReplaced.KeyUp, _
    TxtQtyReplacement.KeyUp, TxtRatioReplaced.KeyUp, TxtRatioReplacement.KeyUp
        Dim TxtBox As Windows.Forms.TextBox

        Try
            TxtBox = TryCast(sender, Windows.Forms.TextBox)
            If IsNothing(TxtBox) = False Then
                Me.ObjCD.Number_KeyUp(TxtBox, sender, e, False)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyReplacement_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyReplacement.Leave
        Try
            Me.ObjCD.Number_Leave(Me.TxtQtyReplacement, sender, e, Me.vDefPrecReplacement, Me.ObjCD)
            'hitung ulang qty replaced
            Me.TxtQtyReplaced.Text = Me.ObjCD.DigitGrouping(Math.Round((CDec(Me.TxtQtyReplacement.Text) * _
            CDec(Me.TxtRatioReplaced.Text) / CDec(Me.TxtRatioReplacement.Text)), _
            Me.vDefPrecReplaced, MidpointRounding.AwayFromZero), Me.vDefPrecReplaced)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRatioReplaced_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRatioReplaced.Leave
        Try
            If CDec(Me.TxtRatioReplaced.Text) <= 0 Then
                Me.TxtRatioReplaced.Text = 1
            End If
            Me.ObjCD.Number_Leave(Me.TxtRatioReplaced, sender, e, Me.vDefPrecReplaced, Me.ObjCD)
            'hitung ulang qty replacement
            Me.TxtQtyReplacement.Text = Me.ObjCD.DigitGrouping(Math.Round((CDec(Me.TxtQtyReplaced.Text) * _
            CDec(Me.TxtRatioReplacement.Text) / CDec(Me.TxtRatioReplaced.Text)), _
            Me.vDefPrecReplacement, MidpointRounding.AwayFromZero), Me.vDefPrecReplacement)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRatioReplacement_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRatioReplacement.Leave
        Try
            If CDec(Me.TxtRatioReplacement.Text) <= 0 Then
                Me.TxtRatioReplacement.Text = 1
            End If
            Me.ObjCD.Number_Leave(Me.TxtRatioReplacement, sender, e, Me.vDefPrecReplacement, Me.ObjCD)
            'hitung ulang qty replacement
            Me.TxtQtyReplacement.Text = Me.ObjCD.DigitGrouping(Math.Round((CDec(Me.TxtQtyReplaced.Text) * _
            CDec(Me.TxtRatioReplacement.Text) / CDec(Me.TxtRatioReplaced.Text)), _
            Me.vDefPrecReplacement, MidpointRounding.AwayFromZero), Me.vDefPrecReplacement)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRatioReplaced_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyReplaced.LostFocus, _
    TxtQtyReplacement.LostFocus, TxtRatioReplaced.LostFocus, TxtRatioReplacement.LostFocus
        Dim TxtBox As Windows.Forms.TextBox

        Try
            TxtBox = TryCast(sender, Windows.Forms.TextBox)
            If IsNothing(TxtBox) = False Then
                TxtBox.BackColor = Me.TxtRef.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyReplaced_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyReplaced.Leave
        Try
            Me.ObjCD.Number_Leave(Me.TxtQtyReplaced, sender, e, Me.vDefPrecReplaced, Me.ObjCD)
            'hitung ulang qty replaced
            Me.TxtQtyReplacement.Text = Me.ObjCD.DigitGrouping(Math.Round((CDec(Me.TxtQtyReplaced.Text) * _
             CDec(Me.TxtRatioReplacement.Text) / CDec(Me.TxtRatioReplaced.Text)), _
            Me.vDefPrecReplacement, MidpointRounding.AwayFromZero), Me.vDefPrecReplacement)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            If MsgBox("Accept data changes?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.LangsungTutup = True
                Me.DialogResult = Windows.Forms.DialogResult.OK
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class