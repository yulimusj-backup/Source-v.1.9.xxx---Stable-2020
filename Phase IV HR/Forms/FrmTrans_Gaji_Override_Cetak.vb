Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Public Class FrmTrans_Gaji_Override_Cetak

    Public mPrimary As Double = -1
    Public mDef_Prec As Double = 0
    Public vTipe As ClsTransaksiSalary_Override.EnumType = ClsTransaksiSalary_Override.EnumType.OtherPayment

    Dim Obj As New ClsTransaksiSalary_Override

    Private Sub btnShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub FrmTrans_Gaji_Cetak_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.Obj.Cetakan(Me.mPrimary, Me.mDef_Prec, Me.crv1, Me.vTipe)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


End Class