Option Explicit On
Option Strict Off

Imports System.Data
Imports Microsoft.VisualBasic.FileIO
Imports System.IO
Imports System.Text

Public Class FrmExportBeginningData
    Public UserCode As String = "USER"
    Public BackColorAll As Color = Color.LightSeaGreen

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub FrmExportBeginningData_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Control.CheckForIllegalCrossThreadCalls = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnExportProduct_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExportProduct.Click

    End Sub

    Private Sub btnExportAP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExportAP.Click

    End Sub

    Private Sub btnExportAR_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExportAR.Click

    End Sub

    Private Sub btnExportSalesReturn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExportSalesReturn.Click

    End Sub

    Private Sub btnExportPurchaseReturn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExportPurchaseReturn.Click

    End Sub

    Private Sub btnExportSalesDP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExportSalesDP.Click

    End Sub

    Private Sub btnExportPurchaseDP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExportPurchaseDP.Click

    End Sub

    Private Sub btnExportSO_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExportSO.Click

    End Sub

    Private Sub btnExportPO_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExportPO.Click

    End Sub

    Private Sub btnExportCreditNote_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExportCreditNote.Click

    End Sub

    Private Sub btnExportDebitNote_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExportDebitNote.Click

    End Sub
End Class