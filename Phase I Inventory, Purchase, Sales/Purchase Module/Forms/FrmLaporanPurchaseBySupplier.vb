Option Explicit On
Option Strict Off

Public Class FrmLaporanPurchaseBySupplier
    Public UserCode As String = ""
    Dim FormatDay As Integer = 1
    Dim FormatMonth As Integer = 2
    Dim FormatYears As Integer = 1
    Dim FormatDecimal As Integer = 0
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim ObjLicense As New Core_Module.ClsLicense
    Dim TBL As New Core_Module.dseMasterSupplier.tblmaster_supplier2DataTable

    Private Sub FrmLaporanInternalPurchaseOrderBySupplier_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim Obj As New Core_Module.clsMasterSupplier
        Try
            Me.FormatDecimal = Me.ObjCur.GetDecimal_DefaultCurrency
            Obj.IsiCheckListBox(Me.clbSupplier, Me.TBL)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Obj = Nothing
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Try
            For i As Integer = 0 To Me.clbSupplier.Items.Count - 1
                Me.clbSupplier.SetItemCheckState(i, CheckState.Checked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Try
            For i As Integer = 0 To Me.clbSupplier.Items.Count - 1
                Me.clbSupplier.SetItemCheckState(i, CheckState.Unchecked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
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
            If Me.dtpTo.Enabled = True Then
                If EndLicenseDate.Date <= Me.dtpTo.Value.Date Then
                    MsgBox("This aplication has expire date on: " + Format(EndLicenseDate, "dd MMM yyyy") + " please contact your software vendor to renew your license.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            If Me.dtpFrom.Value.Date > Me.dtpTo.Value.Date Then
                MsgBox("Wrong date filter", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If Me.clbSupplier.CheckedItems.Count = 0 Then
                MsgBox("Please specify at least 1 supplier", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            Me.Cetak()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Cetak()
        Dim DS As New DseLaporanPurchasebySupplier
        Dim report As New CrpLaporanPurchaseBySupplier
        Dim DTA As New DseLaporanPurchasebySupplierTableAdapters.purchase_by_supplier_report_viewTableAdapter
        Dim row As DseLaporanPurchasebySupplier.purchase_by_supplier_report_viewRow

        Dim DTA_AOP As New DseLaporanPurchasebySupplierTableAdapters.tbltrans_aop1TableAdapter
        Dim TBL_AOP As New DseLaporanPurchasebySupplier.tbltrans_aop1DataTable
        Dim row_AOP As DseLaporanPurchasebySupplier.tbltrans_aop1Row
        Dim DummyDate As Date = Date.MaxValue
        Dim Filter As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA_AOP.Connection.ConnectionString = DTA.Connection.ConnectionString

            DTA.FillBy(DS.purchase_by_supplier_report_view, Me.dtpTo.Value, Me.dtpFrom.Value)
            DTA.Dispose()

            Dim DV As DataRowView
            Dim KodeSupplier As String = ""
            Dim Ketemu As Boolean = True
            For i As Integer = 0 To DS.purchase_by_supplier_report_view.Rows.Count - 1
                Ketemu = False
                If i = DS.purchase_by_supplier_report_view.Rows.Count Then
                    Exit For
                End If
                For Each DV In Me.clbSupplier.CheckedItems
                    If DV.Item(1).ToString = DS.purchase_by_supplier_report_view.Rows(i).Item("Supplier Code") Then
                        Ketemu = True
                        Exit For
                    End If
                Next
                If Ketemu = False Then
                    DS.purchase_by_supplier_report_view.Rows.RemoveAt(i)
                    i -= 1
                Else
                    row = DS.purchase_by_supplier_report_view.Rows(i)
                    row.At_Price = row.At_Amount_Total
                    If row.At_Price <> row.Purchase_IDR + row.Total_Cost Then
                        row.Selisih = 1
                    Else
                        row.Selisih = 0
                    End If
                End If
            Next

            KodeSupplier = ""
            If Me.clbSupplier.CheckedItems.Count = Me.clbSupplier.Items.Count Then
                KodeSupplier = "All"
            Else
                For Each DV In Me.clbSupplier.CheckedItems
                    If KodeSupplier = "" Then
                        KodeSupplier = DV.Item(1)
                    Else
                        KodeSupplier += ", " + DV.Item(1)
                    End If
                Next
            End If

            If Me.CheckBox1.Checked = True And DS.purchase_by_supplier_report_view.Rows.Count <> 0 Then
                For Each row In DS.purchase_by_supplier_report_view
                    If DummyDate > row._Date Then
                        DummyDate = row._Date
                    End If
                Next

                If DummyDate < Date.MaxValue Then
                    DTA_AOP.Fill(TBL_AOP, DummyDate)

                    For Each row In DS.purchase_by_supplier_report_view.Rows
                        Filter = "[No Transaksi] = '" + row.AOP_No + "'"

                        TBL_AOP.DefaultView.RowFilter = Filter
                        If TBL_AOP.DefaultView.Count <> 0 Then
                            row_AOP = TBL_AOP.DefaultView(0).Row
                            row.AOP_No = row_AOP.Supplier_Invoice_No
                        End If
                    Next
                End If
            End If

            report.SetDataSource(DS)
            report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, Me.FormatDay)
            report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, Me.FormatDecimal)
            report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, Me.FormatMonth)
            report.SetParameterValue(report.Parameter_FormatYear.ParameterFieldName, Me.FormatYears)
            report.SetParameterValue(report.Parameter_SupplierCode.ParameterFieldName, KodeSupplier)
            report.SetParameterValue(report.Parameter_TglDari.ParameterFieldName, Me.dtpFrom.Value)
            report.SetParameterValue(report.Parameter_TglSampai.ParameterFieldName, Me.dtpTo.Value)
            Me.crvReport.ReportSource = report
            Me.crvReport.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTA.Dispose()
            DTA_AOP.Dispose()
            TBL_AOP.Dispose()
        End Try
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged

    End Sub
End Class