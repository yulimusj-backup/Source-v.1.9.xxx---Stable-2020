Option Explicit On
Option Strict Off

Public Class FrmLaporanInternalPurchaseOrder
    Dim TBLSup As New Core_Module.dseMasterSupplier.tblmaster_supplier2DataTable
    Dim TBLBrand As New Core_Module.dseMasterProductGroup.tblmaster_product_group2DataTable

    Public UserCode As String = ""

    Private Sub FrmLaporanInternalPurchaseOrderBySupplier_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim Obj As New Core_Module.clsMasterSupplier
        Dim OBjBrand As New Core_Module.ClsMasterProductGroup
        Try
            Obj.IsiCheckListBox(Me.clbSupplier, Me.TBLSup)
            OBjBrand.IsiListBoxBrand(Me.clbBrand, Me.TBLBrand)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Obj = Nothing
            OBjBrand = Nothing
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Try
            For i As Integer = 0 To Me.clbBrand.Items.Count - 1
                Me.clbBrand.SetItemCheckState(i, CheckState.Checked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Try
            For i As Integer = 0 To Me.clbBrand.Items.Count - 1
                Me.clbBrand.SetItemCheckState(i, CheckState.Unchecked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CetakLaporanInternalPurchaseOrder()
        Dim DS As New DseLaporanInternalPurchaseOrder
        Dim DTA As New DseLaporanInternalPurchaseOrderTableAdapters.internal_purchase_order_report_viewTableAdapter
        '     Dim row As DseLaporanInternalPurchaseOrder.internal_purchase_order_report_viewRow
        Dim report As New CrpLaporanInternalPurchaseOrder
        Try
            Dim objCD As New Core_Module.ClsCekData
            Dim TglDari As Date = objCD.GetFirstDateofTheMonth(Me.dtpFrom.Value)
            Dim TglSampai As Date = objCD.GetLasDateOfTheMonth(Me.dtpTo.Value)
            objCD = Nothing
            DTA.FillBy_tglIPODariSampai(DS.internal_purchase_order_report_view, TglDari, TglSampai)

            report.SetDataSource(DS)
            report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, Me.txtFormatDay.Text)
            report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, Me.txtFormatMonth.Text)
            report.SetParameterValue(report.Parameter_FormatYear.ParameterFieldName, Me.txtFormatYears.Text)
            report.SetParameterValue(report.Parameter_TglDari.ParameterFieldName, TglDari)
            report.SetParameterValue(report.Parameter_TglSampai.ParameterFieldName, TglSampai)

            Me.crvReport.ReportSource = report
            Me.crvReport.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTA.Dispose()
        End Try
    End Sub

    Private Sub CetakLaporanInternalPurchaseOrderByBrand()
        Dim DS As New DseLaporanInternalPurchaseOrderByBrand
        Dim DTA As New DseLaporanInternalPurchaseOrderByBrandTableAdapters.internal_purchase_order_report_viewTableAdapter
        Dim row As DseLaporanInternalPurchaseOrderByBrand.internal_purchase_order_report_viewRow
        Dim report As New CrpLaporanInternalPurchaseOrderByBrand
        Try
            Dim Brand As String = ""
            Dim objCD As New Core_Module.ClsCekData
            Dim TglDari As Date = objCD.GetFirstDateofTheMonth(Me.dtpFrom.Value)
            Dim TglSampai As Date = objCD.GetLasDateOfTheMonth(Me.dtpTo.Value)
            objCD = Nothing
            Dim BrandCode As String = ""
            Dim Jumlah As Integer = 0
            Dim DR As DataRowView
            For Each DR In Me.clbBrand.CheckedItems
                BrandCode = DR.Item(1)
                DTA.ClearBeforeFill = False
                DTA.FillBy_TglIPODariSampaiNBrandCode(DS.internal_purchase_order_report_view, TglDari, TglSampai, BrandCode.ToString)
                Jumlah += DS.internal_purchase_order_report_view.Rows.Count
            Next
            If TBLSup.Rows.Count > 0 Then
                Dim KumpulanGroup As New List(Of String)
                For Each row In DS.internal_purchase_order_report_view.Rows
                    If KumpulanGroup.Contains(row.Brand_Code) = False Then
                        KumpulanGroup.Add(row.Brand_Code)
                    End If
                Next
                Dim i As String
                For Each i In KumpulanGroup
                    If Brand = "" Then
                        Brand = i
                    Else
                        Brand = CStr(Brand) + ", " + i
                    End If
                Next
            End If
            report.SetDataSource(DS)
            report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, Me.txtFormatDay.Text)
            report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, Me.txtFormatMonth.Text)
            report.SetParameterValue(report.Parameter_FormatYear.ParameterFieldName, Me.txtFormatYears.Text)
            report.SetParameterValue(report.Parameter_TglDari.ParameterFieldName, TglDari)
            report.SetParameterValue(report.Parameter_TglSampai.ParameterFieldName, TglSampai)
            report.SetParameterValue(report.Parameter_Brand.ParameterFieldName, Brand)

            Me.crvReport.ReportSource = report
            Me.crvReport.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTA.Dispose()
        End Try
    End Sub

    Private Sub CetakLaporanInternalPurchaseOrderBySupplier()
        Dim DS As New DseLaporanInternalPurchaseOrderBySupplier
        Dim DTA As New DseLaporanInternalPurchaseOrderBySupplierTableAdapters.internal_purchase_order_report_viewTableAdapter
        Dim row As DseLaporanInternalPurchaseOrderBySupplier.internal_purchase_order_report_viewRow
        Dim report As New CrpLaporanInternalPurchaseOrderBySupplier
        Try
            Dim Supplier As String = ""
            Dim objCD As New Core_Module.ClsCekData
            Dim TglDari As Date = objCD.GetFirstDateofTheMonth(Me.dtpFrom.Value)
            Dim TglSampai As Date = objCD.GetLasDateOfTheMonth(Me.dtpTo.Value)
            objCD = Nothing
            Dim SupplierCode As String = ""
            Dim Jumlah As Integer = 0
            Dim DR As DataRowView
            For Each DR In Me.clbSupplier.CheckedItems
                SupplierCode = DR.Item(1)
                DTA.ClearBeforeFill = False
                DTA.FillBy_TglIPODariSampaiNSupplierCode(DS.internal_purchase_order_report_view, TglDari, TglSampai, SupplierCode.ToString)
                Jumlah += DS.internal_purchase_order_report_view.Rows.Count
            Next

            If TBLSup.Rows.Count > 0 Then
                Dim KumpulanGroup As New List(Of String)
                For Each row In DS.internal_purchase_order_report_view.Rows
                    If KumpulanGroup.Contains(row.Supplier_Code) = False Then
                        KumpulanGroup.Add(row.Supplier_Code)
                    End If
                Next
                Dim i As String
                For Each i In KumpulanGroup
                    If Supplier = "" Then
                        Supplier = i
                    Else
                        Supplier = CStr(Supplier) + ", " + i
                    End If
                Next
            End If
            report.SetDataSource(DS)
            report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, Me.txtFormatDay.Text)
            report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, Me.txtFormatMonth.Text)
            report.SetParameterValue(report.Parameter_FormatYear.ParameterFieldName, Me.txtFormatYears.Text)
            report.SetParameterValue(report.Parameter_TglDari.ParameterFieldName, TglDari)
            report.SetParameterValue(report.Parameter_TglSampai.ParameterFieldName, TglSampai)
            report.SetParameterValue(report.Parameter_Supplier.ParameterFieldName, Supplier)

            Me.crvReport.ReportSource = report
            Me.crvReport.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTA.Dispose()
        End Try
    End Sub

    Private Sub LinkLabel4_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel4.LinkClicked
        Try
            For i As Integer = 0 To Me.clbSupplier.Items.Count - 1
                Me.clbSupplier.SetItemCheckState(i, CheckState.Checked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel3_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel3.LinkClicked
        Try
            For i As Integer = 0 To Me.clbSupplier.Items.Count - 1
                Me.clbSupplier.SetItemCheckState(i, CheckState.Unchecked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter.Click
        Try
            If Me.rboNoGrouping.Checked = True Then
                Me.CetakLaporanInternalPurchaseOrder()
            ElseIf Me.rboGroupByBrand.Checked = True Then
                Me.CetakLaporanInternalPurchaseOrderByBrand()
            ElseIf Me.rboSupplier.Checked = True Then
                Me.CetakLaporanInternalPurchaseOrderBySupplier()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class