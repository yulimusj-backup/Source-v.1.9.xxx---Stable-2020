Option Explicit On
Option Strict Off

Public Class FrmLaporanHistoricalCostImport
    Public BackColorAll As Color = BackColorAll

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub

    Private Sub CetakSemua()
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim report As New CrpHistoricalCostImport
        Dim DTA As New dseLaporanHistoricalCostImportTableAdapters.TblCetakTableAdapter
        Dim ds As New dseLaporanHistoricalCostImport
        Try
            Dim objCD As New Core_Module.ClsCekData
            Dim TglDari As Date = objCD.GetFirstDateofTheMonth(Me.DtpDate.Value)
            Dim TglSampai As Date = objCD.GetLasDateOfTheMonth(Me.DtpDate.Value)
            objCD = Nothing
            DTA.FillBy_TglDariSampai(ds.TblCetak, TglDari, TglSampai)
            Dim baris As Integer = ds.TblCetak.Rows.Count
            report.SetDataSource(ds)
            report.SetParameterValue(report.Parameter_TanggalDari.ParameterFieldName, TglDari)
            report.SetParameterValue(report.Parameter_TanggalSampai.ParameterFieldName, TglSampai)
            Me.CrystalReportViewer1.ReportSource = report
            Me.CrystalReportViewer1.Show()
            report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, 1)
            report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, 3)
            report.SetParameterValue(report.Parameter_FormatYear.ParameterFieldName, 1)
            report.SetParameterValue(report.Parameter_FormatQty.ParameterFieldName, 2)
            report.SetParameterValue(report.Parameter_FormatPrice.ParameterFieldName, 2)
            Windows.Forms.Cursor.Current = Cursors.Arrow

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ds.Dispose()
        End Try
    End Sub
    Private Sub CetakByproductCode()
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim report As New CrpHistoricalCostImport
        Dim DTA As New dseLaporanHistoricalCostImportTableAdapters.TblCetakTableAdapter
        Dim ds As New dseLaporanHistoricalCostImport
        Try
            Dim objCD As New Core_Module.ClsCekData
            Dim TglDari As Date = objCD.GetFirstDateofTheMonth(Me.DtpDate.Value)
            Dim TglSampai As Date = objCD.GetLasDateOfTheMonth(Me.DtpDate.Value)
            objCD = Nothing

            If Me.TxtCodeProduct.Text = "" Then
                MsgBox("Please Insert Brand Code", MsgBoxStyle.Information, "Information")
            Else
                DTA.FillBy_TglDariSampaiNPrimaryProduct(ds.TblCetak, TglDari, TglSampai, TxtPrimaryProduct.Text)
                Dim baris As Integer = ds.TblCetak.Rows.Count
                report.SetDataSource(ds)
                report.SetParameterValue(report.Parameter_TanggalDari.ParameterFieldName, TglDari)
                report.SetParameterValue(report.Parameter_TanggalSampai.ParameterFieldName, TglSampai)
                Me.CrystalReportViewer1.ReportSource = report
                Me.CrystalReportViewer1.Show()
                report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, 1)
                report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, 3)
                report.SetParameterValue(report.Parameter_FormatYear.ParameterFieldName, 1)
                report.SetParameterValue(report.Parameter_FormatQty.ParameterFieldName, 2)
                report.SetParameterValue(report.Parameter_FormatPrice.ParameterFieldName, 2)
                Windows.Forms.Cursor.Current = Cursors.Arrow
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ds.Dispose()
        End Try
    End Sub

    Private Sub BtnFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFind.Click
        Try
            Dim i As New Core_Module.frmCariMasterProduct
            i.Show()
            If Me.TxtCodeProduct.Text <> "" Then
                i.txtCari.Text = Me.TxtCodeProduct.Text
            End If
            i.Hide()
            i.Focus()
            i.ShowDialog()
            If Not CStr(i.PrimaryProduct) = "" Then
                Me.TxtPrimaryProduct.Text = CStr(i.PrimaryProduct)
                Dim BRg As New Core_Module.clsMasterProduct
                Me.TxtCodeProduct.Text = BRg.GetProductCode(i.PrimaryProduct)
                BRg = Nothing
                Me.DtpDate.Focus()
            End If
            i.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub rbFillAll_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbFillAll.CheckedChanged
        Me.SesuaikanMunculAtauHilang()
    End Sub
    Private Sub SesuaikanMunculAtauHilang()
        Try
            If Me.rbFillAll.Checked = True Then
                Me.Label1.Visible = False
                Me.TxtCodeProduct.Visible = False
                Me.BtnFind.Visible = False
            ElseIf Me.rbFillByProductCode.Checked = True Then
                Me.Label1.Visible = True
                Me.TxtCodeProduct.Visible = True
                Me.BtnFind.Visible = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub rbFillByProductCode_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbFillByProductCode.CheckedChanged
        Me.SesuaikanMunculAtauHilang()
    End Sub
    Private Sub BtnCetak_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnCetak.Click
        Try
            If Me.rbFillAll.Checked = True Then
                Me.CetakSemua()
            ElseIf Me.rbFillByProductCode.Checked = True Then
                Me.CetakByproductCode()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub TxtCodeProduct_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeProduct.GotFocus
        TxtCodeProduct.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtCodeProduct_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeProduct.LostFocus
        TxtCodeProduct.BackColor = Me.TxtPrimaryProduct.BackColor
    End Sub
End Class