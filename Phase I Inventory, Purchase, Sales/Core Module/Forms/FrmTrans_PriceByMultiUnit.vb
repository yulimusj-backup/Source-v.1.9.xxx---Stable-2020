Option Explicit On
Option Strict Off

Public Class FrmTrans_PriceByMultiUnit
    Public UnitMulti As String = ""
    Public UnitDefault As String = ""
    Public Buffer_PriceByMultiUnit As String = ""

    Private Sub FrmTrans_PriceByMultiUnit_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.cboPriceByMultiUnit.Items.Clear()
            Me.cboPriceByMultiUnit.Items.Add(Me.UnitDefault)
            Me.cboPriceByMultiUnit.Items.Add(Me.UnitMulti)
            Me.cboPriceByMultiUnit.SelectedIndex = Me.Buffer_PriceByMultiUnit
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            If (Me.Buffer_PriceByMultiUnit = "1" And Me.cboPriceByMultiUnit.Text = Me.UnitDefault) Or _
            (Me.Buffer_PriceByMultiUnit = "0" And Me.cboPriceByMultiUnit.Text = Me.UnitMulti) Then
                If MsgBox("Change unit for price calculation?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Me.DialogResult = Windows.Forms.DialogResult.OK
                    Me.Close()
                End If
            Else
                Me.DialogResult = Windows.Forms.DialogResult.Cancel
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class