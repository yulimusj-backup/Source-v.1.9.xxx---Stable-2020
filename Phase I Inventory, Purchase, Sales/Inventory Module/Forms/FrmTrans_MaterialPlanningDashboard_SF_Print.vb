Option Explicit On
Option Strict Off

Public Class FrmTrans_MaterialPlanningDashboard_SF_Print
    Public vPrimaryTransaksi As Double = -1
    Public UserName As String = ""
    Public vFilter As String = ""
    Public vLateOnly As Boolean
    Public vEnumPrint As ClsTransaksiShippingForecast.EnumPrintAvailability = ClsTransaksiShippingForecast.EnumPrintAvailability.All
    Public vAvDateChecked As Boolean = False
    Public vAvDate As Date = Today

    Dim ObjSF As New ClsTransaksiShippingForecast

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub FrmTrans_MaterialPlanningDashboard_SF_Print_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim ij As New FrmTrans_MaterialPlanningDashboard_SF_PrintDetail

        Try
            ij.UserName = Me.UserName
            If ij.ShowDialog = Windows.Forms.DialogResult.OK Then
                ObjSF.Cetak(Me.CrvLaporan, Me.vPrimaryTransaksi, ij.cboGroupBy.SelectedIndex, ij.cboOrderNo.SelectedIndex, Me.vFilter, Me.vLateOnly, Me.vEnumPrint, _
                Me.vAvDate, Me.vAvDateChecked)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub
End Class