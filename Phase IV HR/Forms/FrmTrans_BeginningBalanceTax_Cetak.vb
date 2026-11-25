Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Public Class FrmTrans_Tax_Beginning_Cetak

    Public mPrimary As Double = -1
    Public mDef_Prec As Double = 0
    Dim Obj As New ClsTrans_Tax_Beginning_Balance

    Private Sub btnShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub FrmTrans_Gaji_Cetak_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.Obj.CetakTaxBalance(Me.mPrimary, Me.crv1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


End Class