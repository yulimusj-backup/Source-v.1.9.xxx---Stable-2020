Option Explicit On
Option Strict Off

Public Class FrmUjiCetak

    Private Sub btnCetakInvoice_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCetakInvoice.Click
        Dim ObjInv As New ClsTransaksiSalesInvoice

        ObjInv.CetakInvoiceCustom(Me.CrystalReportViewer1, Me.TxtPrimary.Text, #1/1/1950#, 2, 1, 1, 1, 2, 1, 1, 1, 1)
        ObjInv = Nothing
    End Sub

    Private Sub FrmUjiCetak_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class