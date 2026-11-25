Option Explicit On
Option Strict Off

Public Class frmLaporanStockTakingTeam
    Public Primary As Double
    Public Description As String
    Public StartDate As Date
    Public StockTakingNo As String
    Public TotalTeam As Int32
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub frmLaporanStockTakingTeam_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.Cetak()
    End Sub

    Private Sub Cetak()
        Dim DTA As New dseTransaksiStockTakingTableAdapters.tblCetakTableAdapter
        Dim DS As New dseTransaksiStockTaking
        Dim report As New crpLaporanStockTakingTeam
        Try
            DTA.FillBy_PrimaryStockTaking(DS.tblCetak, Primary)
            DTA.Dispose()
            report.SetDataSource(DS)
            report.SetParameterValue(report.Parameter_Description.ParameterFieldName, Description)
            report.SetParameterValue(report.Parameter_StartDate.ParameterFieldName, StartDate)
            report.SetParameterValue(report.Parameter_StockTakingNo.ParameterFieldName, StockTakingNo)
            report.SetParameterValue(report.Parameter_TotalTeam.ParameterFieldName, TotalTeam)
            Me.CrvLaporan.ReportSource = report
            Me.CrvLaporan.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
        End Try
    End Sub
End Class