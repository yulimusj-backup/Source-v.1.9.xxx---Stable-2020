Option Explicit On
Option Strict Off

Public Class FrmLaporanPurchaseStatisticBySupplier
    Public UserCode As String = ""
    Dim FormatDay As Integer = 1
    Dim FormatMonth As Integer = 2
    Dim FormatYears As Integer = 1
    Dim FormatDecimal As Integer = 0
    Dim ObjLicense As New Core_Module.ClsLicense
    Dim ObjCurr As New Core_Module.clsMasterCurrency

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter.Click
        Dim EndLicenseDate As Date = #1/1/1950#
        Try
            EndLicenseDate = Me.ObjLicense.GetExpiryDate(My.Settings.b)
            If Me.dtpFrom.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpFrom.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            Me.Cetak()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Cetak()
        Dim SupplierType As String = ""
        Dim DS As New DseLaporanPurchaseStatisticBySupplier
        Dim report As New CrpLaporanPurchaseStatisticBySupplier
        Dim objCD As New Core_Module.ClsCekData
        Dim TglDari As Date = "#1/1/" + Me.dtpFrom.Value.Year.ToString + "#"
        Dim TglSampai As Date = objCD.GetLasDateOfTheMonth(Me.dtpFrom.Value)
        objCD = Nothing
        Dim counter As Double = 0
        Dim Filter As String = ""
        Dim DTA As New DseLaporanPurchaseStatisticBySupplierTableAdapters.DataTable1TableAdapter

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            'tanggal dari dibatasi sejak tanggal start implementasi CMSI
            If TglDari.Date < #4/1/2010# Then
                TglDari = #4/1/2010#
            End If

            If Me.rboForeignSupplier.Checked = True Then
                DTA.Fill(DS.DataTable1, TglSampai, TglDari, 1)
                SupplierType = "Foreign"
            ElseIf Me.rboLocalSupplier.Checked = True Then
                DTA.Fill(DS.DataTable1, TglSampai, TglDari, 0)
                SupplierType = "Local"
            Else
                DTA.FillBy_All(DS.DataTable1, TglSampai, TglDari)
                SupplierType = "All"
            End If

            Dim row As DseLaporanPurchaseStatisticBySupplier.DataTable1Row
            Dim RowCetak As DseLaporanPurchaseStatisticBySupplier.Purchase_Statistic_DataTableRow
            For Each row In DS.DataTable1.Rows
                Filter = "[Supplier Code] ='" + row.Supplier_Code + "'"
                DS.Purchase_Statistic_DataTable.DefaultView.RowFilter = Filter
                If DS.Purchase_Statistic_DataTable.DefaultView.Count > 0 Then
                    RowCetak = DS.Purchase_Statistic_DataTable.DefaultView(0).Row
                    If row.Tanggal_Transaksi.Month = 1 Then
                        RowCetak.January += row.Price
                    ElseIf row.Tanggal_Transaksi.Month = 2 Then
                        RowCetak.February += row.Price
                    ElseIf row.Tanggal_Transaksi.Month = 3 Then
                        RowCetak.March += row.Price
                    ElseIf row.Tanggal_Transaksi.Month = 4 Then
                        RowCetak.April += row.Price
                    ElseIf row.Tanggal_Transaksi.Month = 5 Then
                        RowCetak.May += row.Price
                    ElseIf row.Tanggal_Transaksi.Month = 6 Then
                        RowCetak.June += row.Price
                    ElseIf row.Tanggal_Transaksi.Month = 7 Then
                        RowCetak.July += row.Price
                    ElseIf row.Tanggal_Transaksi.Month = 8 Then
                        RowCetak.August += row.Price
                    ElseIf row.Tanggal_Transaksi.Month = 9 Then
                        RowCetak.September += row.Price
                    ElseIf row.Tanggal_Transaksi.Month = 10 Then
                        RowCetak.October += row.Price
                    ElseIf row.Tanggal_Transaksi.Month = 11 Then
                        RowCetak.November += row.Price
                    ElseIf row.Tanggal_Transaksi.Month = 12 Then
                        RowCetak.December += row.Price
                    End If
                Else
                    If row.Tanggal_Transaksi.Month = 1 Then
                        DS.Purchase_Statistic_DataTable.Rows.Add(row.Supplier_Code, _
                        row.Supplier_Name, row.Price, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
                    ElseIf row.Tanggal_Transaksi.Month = 2 Then
                        DS.Purchase_Statistic_DataTable.Rows.Add(row.Supplier_Code, _
                        row.Supplier_Name, 0, row.Price, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
                    ElseIf row.Tanggal_Transaksi.Month = 3 Then
                        DS.Purchase_Statistic_DataTable.Rows.Add(row.Supplier_Code, _
                        row.Supplier_Name, 0, 0, row.Price, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
                    ElseIf row.Tanggal_Transaksi.Month = 4 Then
                        DS.Purchase_Statistic_DataTable.Rows.Add(row.Supplier_Code, _
                        row.Supplier_Name, 0, 0, 0, row.Price, 0, 0, 0, 0, 0, 0, 0, 0, 0)
                    ElseIf row.Tanggal_Transaksi.Month = 5 Then
                        DS.Purchase_Statistic_DataTable.Rows.Add(row.Supplier_Code, _
                        row.Supplier_Name, 0, 0, 0, 0, row.Price, 0, 0, 0, 0, 0, 0, 0, 0)
                    ElseIf row.Tanggal_Transaksi.Month = 6 Then
                        DS.Purchase_Statistic_DataTable.Rows.Add(row.Supplier_Code, _
                        row.Supplier_Name, 0, 0, 0, 0, 0, row.Price, 0, 0, 0, 0, 0, 0, 0)
                    ElseIf row.Tanggal_Transaksi.Month = 7 Then
                        DS.Purchase_Statistic_DataTable.Rows.Add(row.Supplier_Code, _
                        row.Supplier_Name, 0, 0, 0, 0, 0, 0, row.Price, 0, 0, 0, 0, 0, 0)
                    ElseIf row.Tanggal_Transaksi.Month = 8 Then
                        DS.Purchase_Statistic_DataTable.Rows.Add(row.Supplier_Code, _
                        row.Supplier_Name, 0, 0, 0, 0, 0, 0, 0, row.Price, 0, 0, 0, 0, 0)
                    ElseIf row.Tanggal_Transaksi.Month = 9 Then
                        DS.Purchase_Statistic_DataTable.Rows.Add(row.Supplier_Code, _
                        row.Supplier_Name, 0, 0, 0, 0, 0, 0, 0, 0, row.Price, 0, 0, 0, 0)
                    ElseIf row.Tanggal_Transaksi.Month = 10 Then
                        DS.Purchase_Statistic_DataTable.Rows.Add(row.Supplier_Code, _
                        row.Supplier_Name, 0, 0, 0, 0, 0, 0, 0, 0, 0, row.Price, 0, 0, 0)
                    ElseIf row.Tanggal_Transaksi.Month = 11 Then
                        DS.Purchase_Statistic_DataTable.Rows.Add(row.Supplier_Code, _
                        row.Supplier_Name, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, row.Price, 0, 0)
                    ElseIf row.Tanggal_Transaksi.Month = 12 Then
                        DS.Purchase_Statistic_DataTable.Rows.Add(row.Supplier_Code, _
                        row.Supplier_Name, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, row.Price, 0)
                    End If
                End If
            Next

            For Each RowCetak In DS.Purchase_Statistic_DataTable.Rows
                RowCetak.Total_Of_Purchase = RowCetak.January + RowCetak.February + _
                RowCetak.March + RowCetak.April + RowCetak.May + RowCetak.June + _
                RowCetak.July + RowCetak.August + RowCetak.September + _
                RowCetak.October + RowCetak.November + RowCetak.December
            Next

            FormatDecimal = Me.ObjCurr.GetDecimal_DefaultCurrency
            report.SetDataSource(DS)
            report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, Me.FormatDay)
            report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, Me.FormatDecimal)
            report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, Me.FormatMonth)
            report.SetParameterValue(report.Parameter_FormatYear.ParameterFieldName, Me.FormatYears)
            report.SetParameterValue(report.Parameter_SupplierType.ParameterFieldName, SupplierType)
            report.SetParameterValue(report.Parameter_Tgl.ParameterFieldName, Me.dtpFrom.Value)
            Me.crvReport.ReportSource = report
            Me.crvReport.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTA.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub FrmLaporanPurchaseStatisticBySupplier_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            Me.dtpFrom.Value = ObjCD.GetFirstDateofTheMonth(Me.dtpFrom.Value)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCD = Nothing
        End Try
    End Sub
End Class