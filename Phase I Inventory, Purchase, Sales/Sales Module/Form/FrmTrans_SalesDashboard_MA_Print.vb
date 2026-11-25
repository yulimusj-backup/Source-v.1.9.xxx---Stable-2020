Option Explicit On
Option Strict Off

Public Class FrmTrans_SalesDashboard_MA_Print
    Public UserName As String = ""
    Public vFilter As String = ""
    Public vTBLItemSOB As New Inventory_Module.DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryDataTable
    Public vEnumPrint As Inventory_Module.ClsTransaksiShippingForecast.EnumPrintAvailability = Inventory_Module.ClsTransaksiShippingForecast.EnumPrintAvailability.All
    Public vAvDateChecked As Boolean = False
    Public vAvDate As Date = Today
    Dim ObjSF As New Inventory_Module.ClsTransaksiShippingForecast

    Private Sub FrmTrans_MaterialPlanningDashboard_SF_Print_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim ij As New Inventory_Module.FrmTrans_MaterialPlanningDashboard_SF_PrintDetail

        Try
            ij.UserName = Me.UserName
            If ij.ShowDialog = Windows.Forms.DialogResult.OK Then
                ObjSF.Cetak(Me.CrvLaporan, Me.vTBLItemSOB, ij.cboGroupBy.SelectedIndex, ij.cboOrderNo.SelectedIndex, Me.vFilter, _
                Me.vEnumPrint, Me.vAvDate, Me.vAvDateChecked)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
End Class