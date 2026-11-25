Option Explicit On
Option Strict Off

Public Class FrmLaporanStockTaking_LastEDCheck
    Public DS As dseLaporanStockTaking

    Private report As New crpLaporanStockTaking_ED_LastCheck

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    ''' <summary>
    ''' load data ke dalam laporan
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub FrmLaporanStockTaking_LastEDCheck_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            report.SetDataSource(DS)
            report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, 3)
            report.SetParameterValue(report.Parameter_FormatYears.ParameterFieldName, 1)
            report.SetParameterValue(report.Parameter_FormatDays.ParameterFieldName, 1)

            Me.CrvLaporan.ReportSource = report
            Me.CrvLaporan.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class