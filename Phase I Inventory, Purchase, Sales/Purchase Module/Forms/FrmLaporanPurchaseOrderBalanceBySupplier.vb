Option Explicit On
Option Strict Off

Imports Core_Module

Public Class FrmLaporanPurchaseOrderBalanceBySupplier
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Dim ObjSupplier As New clsMasterSupplier

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub txtSupplierCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSupplierCode.KeyDown
        If Me.txtSupplierCode.Enabled = False Or Me.txtSupplierCode.ReadOnly = True Then
            Exit Sub
        End If

        Dim i As New FrmCariMasterSupplier
        i.BackColorAll = Me.BackColorAll

        Try
            If e.KeyValue = 13 Then
                If Me.ObjSupplier.ApakahSudahAda(Me.txtSupplierCode.Text) = True Then
                    Me.txtPrimarySupplier.Text = Me.ObjSupplier.GetPrimary(Me.txtSupplierCode.Text)
                    Me.txtSupplierName.Text = Me.ObjSupplier.GetSupplierName(Me.txtPrimarySupplier.Text)
                    Me.btnFindSupplier.Focus()
                Else
                    'ada 2 Customer dengan kode yang sama atau tidak ada Customer yang dimaksud
                    If Me.txtSupplierCode.Text <> "" Then
                        i.txtCari.Text = Me.txtSupplierCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimarySupplier > 0 Then
                        Me.txtPrimarySupplier.Text = i.PrimarySupplier
                        Me.txtSupplierCode.Text = Me.ObjSupplier.GetSupplierCode(Me.txtPrimarySupplier.Text)
                        Me.txtSupplierName.Text = Me.ObjSupplier.GetSupplierName(Me.txtPrimarySupplier.Text)
                        Me.btnFindSupplier.Focus()
                    Else
                        Me.txtSupplierCode.Clear()
                        Me.txtSupplierName.Clear()
                        Me.txtPrimarySupplier.Text = "-1"
                        Me.txtSupplierCode.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtSupplierCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSupplierCode.Leave
        If Me.txtSupplierCode.Enabled = False Or Me.txtSupplierCode.ReadOnly = True Then
            Exit Sub
        End If

        Dim i As New FrmCariMasterSupplier
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.txtSupplierCode.Text <> "" Then
                If Me.ObjSupplier.ApakahSudahAda(Me.txtSupplierCode.Text) Then
                    Dim PrimaryBaru As Double = Me.ObjSupplier.GetPrimary(Me.txtSupplierCode.Text)
                    If PrimaryBaru <> Me.txtPrimarySupplier.Text Then
                        Me.txtPrimarySupplier.Text = Me.ObjSupplier.GetPrimary(Me.txtSupplierCode.Text)
                        Me.txtSupplierName.Text = Me.ObjSupplier.GetSupplierName(Me.txtPrimarySupplier.Text)
                        Me.btnFindSupplier.Focus()
                    End If
                Else
                    If Me.txtSupplierCode.Text <> "" Then
                        i.txtCari.Text = Me.txtSupplierCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimarySupplier > 0 Then
                        Me.txtPrimarySupplier.Text = i.PrimarySupplier
                        Me.txtSupplierCode.Text = Me.ObjSupplier.GetSupplierCode(Me.txtPrimarySupplier.Text)
                        Me.txtSupplierName.Text = Me.ObjSupplier.GetSupplierName(Me.txtPrimarySupplier.Text)
                        Me.btnFindSupplier.Focus()
                    Else
                        Me.txtSupplierCode.Clear()
                        Me.txtSupplierName.Clear()
                        Me.txtPrimarySupplier.Text = "-1"
                        Me.txtSupplierCode.Focus()
                    End If
                End If
            Else
                Me.txtSupplierCode.Clear()
                Me.txtSupplierName.Clear()
                Me.txtPrimarySupplier.Text = "-1"
                Me.btnFindSupplier.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindSupplier_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindSupplier.Click
        Dim i As New FrmCariMasterSupplier
        i.BackColorAll = Me.BackColorAll
        Try
            i.cboCriteria.SelectedIndex = 0
            If Me.txtSupplierCode.Text <> "" Then
                i.txtCari.Text = Me.txtSupplierCode.Text
            End If

            i.ShowDialog()
            If i.PrimarySupplier > 0 Then
                Me.txtPrimarySupplier.Text = i.PrimarySupplier
                Me.txtSupplierCode.Text = Me.ObjSupplier.GetSupplierCode(Me.txtPrimarySupplier.Text)
                Me.txtSupplierName.Text = Me.ObjSupplier.GetSupplierName(Me.txtPrimarySupplier.Text)
                Me.btnFindSupplier.Focus()
            Else
                Me.txtSupplierCode.Clear()
                Me.txtSupplierName.Clear()
                Me.txtPrimarySupplier.Text = "-1"
                Me.txtSupplierCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtSupplierCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSupplierCode.GotFocus
        txtSupplierCode.BackColor = Me.BackColorAll
    End Sub

    Private Sub btnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter.Click
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Try
            If Me.txtSupplierCode.Text <> "" Then
                If Me.dtpFrom.Value.Date > Me.dtpTo.Value.Date Then
                    MsgBox("Wrong date filter", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                Me.Cetak2()
            Else
                MsgBox("Please enter supplier code first", MsgBoxStyle.Information, "Information")
                Me.txtSupplierCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub Cetak_Obselete()
        'fungsi ini sudah obselete
        'dulu setiap kali generate laporan ini, maka software akan menghitung ulang semuanya, yang baru nilai
        'pob disimpan pada item

        Dim DS As New DseLaporanPurchaseOrderBalance
        Dim report As New CrpLaporanOrderBalanceReportBySupplier
        Try
            'Fillby Tgl Untuk orderbalance_report_po_item_view
            Dim DTA As New DseLaporanPurchaseOrderBalanceTableAdapters.orderbalance_report_po_item_viewTableAdapter
            DTA.Fill(DS.orderbalance_report_po_item_view, CDbl(Me.txtPrimarySupplier.Text))
            DTA.Dispose()

            'Lakukan Pengurangan dengan POC
            If DS.orderbalance_report_po_item_view.Rows.Count > 0 Then
                'Lakukan Penambahan dengan P.Return
                Me.LakukanPenambahanDenganPurchaseReturn1(DS)
                'Lakukan Pengurangan dengan POC
                Me.LakukanPenguranganDenganPOC1(DS)
                'Lakukan Pengurangan dengan GIT
                Me.LakukanPenguranganDenganGIT1(DS)
                'Lakukan Pengurangan dengan GRH
                Me.LakukanPenguranganDenganGRN1(DS)
                'Lakukan Perhitungan Kembali Amount
                Me.HitungKembaliNilaiTotalAmount(DS)
                'Qty Yang Nol Nga usah di munculin 
                Me.EliminasiNol(DS)
            End If

            report.SetDataSource(DS)
            report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, Me.txtFormatDay.Text)
            report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, Me.txtFormatDecimal.Text)
            report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, Me.txtFormatMonth.Text)
            report.SetParameterValue(report.Parameter_FormatYear.ParameterFieldName, Me.txtFormatYears.Text)
            report.SetParameterValue(report.Parameter_SupplierCode.ParameterFieldName, Me.txtSupplierCode.Text)
            report.SetParameterValue(report.Parameter_SupplierName.ParameterFieldName, Me.txtSupplierName.Text)
            Me.crvReport.ReportSource = report
            Me.crvReport.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
        End Try
    End Sub

    Private Sub Cetak2()
        Dim DS As New DseLaporanPurchaseOrderBalance
        Dim report As New CrpLaporanOrderBalanceReportBySupplier
        Dim DTA As New DseLaporanPurchaseOrderBalanceTableAdapters.orderbalance_report_po_item_viewTableAdapter

        Try
            If Me.chkPeriod.Checked = True Then
                DTA.FillBy_TanggalSampaiDari(DS.orderbalance_report_po_item_view, CDbl(Me.txtPrimarySupplier.Text), Me.dtpTo.Value, Me.dtpFrom.Value)
            Else
                DTA.Fill(DS.orderbalance_report_po_item_view, CDbl(Me.txtPrimarySupplier.Text))
            End If

            report.SetDataSource(DS)
            report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, Me.txtFormatDay.Text)
            report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, Me.txtFormatDecimal.Text)
            report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, Me.txtFormatMonth.Text)
            report.SetParameterValue(report.Parameter_FormatYear.ParameterFieldName, Me.txtFormatYears.Text)
            report.SetParameterValue(report.Parameter_SupplierCode.ParameterFieldName, Me.txtSupplierCode.Text)
            report.SetParameterValue(report.Parameter_SupplierName.ParameterFieldName, Me.txtSupplierName.Text)
            Me.crvReport.ReportSource = report
            Me.crvReport.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTA.Dispose()
        End Try
    End Sub

    Private Sub EliminasiNol(ByRef DS As DseLaporanPurchaseOrderBalance)
        Try
            For i As Integer = 0 To DS.orderbalance_report_po_item_view.Rows.Count - 1
                If i > DS.orderbalance_report_po_item_view.Rows.Count - 1 Then
                    Exit For
                End If
                If DS.orderbalance_report_po_item_view.Rows(i).Item("Total Qty") = 0 Then
                    DS.orderbalance_report_po_item_view.Rows.RemoveAt(i)
                    i -= 1
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LakukanPenguranganDenganPOC1(ByRef DS As DseLaporanPurchaseOrderBalance)
        Dim rowPO As DseLaporanPurchaseOrderBalance.orderbalance_report_po_item_viewRow
        Dim rowPOC As DseLaporanPurchaseOrderBalance.orderbalance_report_item_poc_viewRow
        Try
            Dim DTAPOC As New DseLaporanPurchaseOrderBalanceTableAdapters.orderbalance_report_item_poc_viewTableAdapter
            DTAPOC.FillBy_TglDariSampai(DS.orderbalance_report_item_poc_view, Me.dtpFrom.Value, Me.dtpTo.Value)
            DTAPOC.Dispose()
            If DS.orderbalance_report_item_poc_view.Rows.Count > 0 Then
                For Each rowPOC In DS.orderbalance_report_item_poc_view.Rows
                    Dim TBL As DataTable = DS.orderbalance_report_po_item_view
                    Dim View As DataView = TBL.DefaultView
                    Dim cmd As String = "[Primary PO] =" + CStr(rowPOC.Primary_PO) + _
                    " and [Discount] =" + CStr(rowPOC.Discount) + _
                    " and [Primary Product] =" + CStr(rowPOC.Primary_Product) + _
                    " and [Unit Price] =" + CStr(rowPOC.Unit_Price) + _
                    " and [Nett Price] =" + CStr(rowPOC.Nett_Price) + _
                    " and [Primary IPO] > -1"
                    Dim rowview As DataRowView

                    View.RowFilter = cmd
                    View.Sort = "[Primary IPO] ASC"

                    For Each rowview In View
                        rowPO = rowview.Row
                        If rowPOC.Total_Qty = 0 Then
                            Exit For
                        End If
                        If rowPO.Total_Qty >= rowPOC.Total_Qty Then
                            rowPO.Total_Qty -= rowPOC.Total_Qty
                            rowPOC.Total_Qty -= rowPOC.Total_Qty
                        Else
                            'Jika Qty PO Lebih Kecil maka Kurangi dengan diri sendiri agar menjadi Nol Kemudian biarkan 
                            'Agar Item Selanjutnya yang di kurangkan
                            rowPOC.Total_Qty -= rowPO.Total_Qty
                            rowPO.Total_Qty -= rowPO.Total_Qty
                        End If
                    Next

                    'View Kedua
                    cmd = "[Primary PO] =" + CStr(rowPOC.Primary_PO) + _
                    " and Discount =" + CStr(rowPOC.Discount) + _
                    " and [Primary Product] =" + CStr(rowPOC.Primary_Product) + _
                    " and [Unit Price] =" + CStr(rowPOC.Unit_Price) + _
                    " and [Nett Price] =" + CStr(rowPOC.Nett_Price) + _
                    " and [Primary IPO] =-1"
                    View.RowFilter = cmd
                    For Each rowview In View
                        rowPO = rowview.Row
                        If rowPOC.Total_Qty = 0 Then
                            Exit For
                        End If
                        If rowPO.Total_Qty >= rowPOC.Total_Qty Then
                            rowPO.Total_Qty -= rowPOC.Total_Qty
                            rowPOC.Total_Qty -= rowPOC.Total_Qty
                        Else
                            'Jika Qty PO Lebih Kecil maka Kurangi dengan diri sendiri agar menjadi Nol Kemudian biarkan 
                            'Agar Item Selanjutnya yang di kurangkan
                            rowPOC.Total_Qty -= rowPO.Total_Qty
                            rowPO.Total_Qty -= rowPO.Total_Qty
                        End If
                    Next
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LakukanPenguranganDenganPOC(ByRef DS As DseLaporanPurchaseOrderBalance)
        Dim rowPO As DseLaporanPurchaseOrderBalance.orderbalance_report_po_item_viewRow
        Dim rowPOC As DseLaporanPurchaseOrderBalance.orderbalance_report_item_poc_viewRow
        Try
            Dim DTAPOC As New DseLaporanPurchaseOrderBalanceTableAdapters.orderbalance_report_item_poc_viewTableAdapter
            DTAPOC.FillBy_TglDariSampai(DS.orderbalance_report_item_poc_view, Me.dtpFrom.Value, Me.dtpTo.Value)
            DTAPOC.Dispose()
            If DS.orderbalance_report_item_poc_view.Rows.Count > 0 Then
                For Each rowPO In DS.orderbalance_report_po_item_view.Rows
                    For Each rowPOC In DS.orderbalance_report_item_poc_view.Rows
                        If rowPO.Primary_PO = rowPOC.Primary_PO And rowPO.Discount = rowPOC.Discount And _
                        rowPO.Primary_Product = rowPOC.Primary_Product And rowPO.Unit_Price = rowPOC.Unit_Price And _
                        rowPO.Nett_Price = rowPOC.Nett_Price Then


                            If rowPO.Total_Qty >= rowPOC.Total_Qty Then
                                rowPO.Total_Qty -= rowPOC.Total_Qty
                                rowPOC.Total_Qty -= rowPOC.Total_Qty
                            Else
                                'Jika Qty PO Lebih Kecil maka Kurangi dengan diri sendiri agar menjadi Nol Kemudian biarkan 
                                'Agar Item Selanjutnya yang di kurangkan
                                rowPO.Total_Qty -= rowPO.Total_Qty
                                rowPOC.Total_Qty -= rowPO.Total_Qty
                            End If
                        End If
                    Next
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LakukanPenguranganDenganGIT1(ByRef DS As DseLaporanPurchaseOrderBalance)
        Dim rowPO As DseLaporanPurchaseOrderBalance.orderbalance_report_po_item_viewRow
        Dim rowGIT As DseLaporanPurchaseOrderBalance.orderbalance_report_goods_in_transit_viewRow
        Try
            Dim DTAGIT As New DseLaporanPurchaseOrderBalanceTableAdapters.orderbalance_report_goods_in_transit_viewTableAdapter
            DTAGIT.FillBy_TglDariSampai(DS.orderbalance_report_goods_in_transit_view, Me.dtpFrom.Value, Me.dtpTo.Value)
            DTAGIT.Dispose()
            If DS.orderbalance_report_goods_in_transit_view.Rows.Count > 0 Then
                For Each rowGIT In DS.orderbalance_report_goods_in_transit_view.Rows
                    Dim TBL As DataTable = DS.orderbalance_report_po_item_view
                    Dim View As DataView = TBL.DefaultView
                    Dim cmd As String = "[Primary PO] =" + CStr(rowGIT.Primary_PO) + _
                    " and [Discount] =" + CStr(rowGIT.Discount) + _
                    " and [Primary Product] =" + CStr(rowGIT.Primary_Product) + _
                    " and [Unit Price] =" + CStr(rowGIT.Unit_Price) + _
                    " and [Nett Price] =" + CStr(rowGIT.Nett_Price) + _
                    " and [Primary IPO] > -1"
                    Dim rowview As DataRowView

                    View.RowFilter = cmd
                    View.Sort = "[Primary IPO] ASC"

                    For Each rowview In View
                        rowPO = rowview.Row
                        If rowGIT.Total_Qty = 0 Then
                            Exit For
                        End If
                        If rowPO.Total_Qty >= rowGIT.Total_Qty Then
                            rowPO.Total_Qty -= rowGIT.Total_Qty
                            rowGIT.Total_Qty -= rowGIT.Total_Qty
                        Else
                            'Jika Qty PO Lebih Kecil maka Kurangi dengan diri sendiri agar menjadi Nol Kemudian biarkan 
                            'Agar Item Selanjutnya yang di kurangkan
                            rowGIT.Total_Qty -= rowPO.Total_Qty
                            rowPO.Total_Qty -= rowPO.Total_Qty
                        End If
                    Next

                    'View Kedua
                    cmd = "[Primary PO] =" + CStr(rowGIT.Primary_PO) + _
                    " and Discount =" + CStr(rowGIT.Discount) + _
                    " and [Primary Product] =" + CStr(rowGIT.Primary_Product) + _
                    " and [Unit Price] =" + CStr(rowGIT.Unit_Price) + _
                    " and [Nett Price] =" + CStr(rowGIT.Nett_Price) + _
                    " and [Primary IPO] =-1"
                    View.RowFilter = cmd
                    For Each rowview In View
                        rowPO = rowview.Row
                        If rowGIT.Total_Qty = 0 Then
                            Exit For
                        End If
                        If rowPO.Total_Qty >= rowGIT.Total_Qty Then
                            rowPO.Total_Qty -= rowGIT.Total_Qty
                            rowGIT.Total_Qty -= rowGIT.Total_Qty
                        Else
                            'Jika Qty PO Lebih Kecil maka Kurangi dengan diri sendiri agar menjadi Nol Kemudian biarkan 
                            'Agar Item Selanjutnya yang di kurangkan
                            rowGIT.Total_Qty -= rowPO.Total_Qty
                            rowPO.Total_Qty -= rowPO.Total_Qty
                        End If
                    Next
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LakukanPenguranganDenganGIT(ByRef DS As DseLaporanPurchaseOrderBalance)
        Dim rowPO As DseLaporanPurchaseOrderBalance.orderbalance_report_po_item_viewRow
        Dim rowGIT As DseLaporanPurchaseOrderBalance.orderbalance_report_goods_in_transit_viewRow
        Try
            Dim DTAGIT As New DseLaporanPurchaseOrderBalanceTableAdapters.orderbalance_report_goods_in_transit_viewTableAdapter
            DTAGIT.FillBy_TglDariSampai(DS.orderbalance_report_goods_in_transit_view, Me.dtpFrom.Value, Me.dtpTo.Value)
            DTAGIT.Dispose()
            If DS.orderbalance_report_goods_in_transit_view.Rows.Count > 0 Then
                For Each rowPO In DS.orderbalance_report_po_item_view.Rows
                    For Each rowGIT In DS.orderbalance_report_goods_in_transit_view.Rows
                        If rowPO.Primary_PO = rowGIT.Primary_PO And rowPO.Discount = rowGIT.Discount And _
                        rowPO.Primary_Product = rowGIT.Primary_Product And rowPO.Unit_Price = rowGIT.Unit_Price And _
                        rowPO.Nett_Price = rowGIT.Nett_Price Then
                            If rowPO.Total_Qty >= rowGIT.Total_Qty Then
                                rowPO.Total_Qty -= rowGIT.Total_Qty
                                rowGIT.Total_Qty -= rowGIT.Total_Qty
                            Else
                                'Jika Qty PO Lebih Kecil maka Kurangi dengan diri sendiri agar menjadi Nol Kemudian biarkan 
                                'Agar Item Selanjutnya yang di kurangkan
                                rowPO.Total_Qty -= rowPO.Total_Qty
                                rowGIT.Total_Qty -= rowPO.Total_Qty
                            End If
                        End If
                    Next
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LakukanPenguranganDenganGRN1(ByRef DS As DseLaporanPurchaseOrderBalance)
        Dim rowPO As DseLaporanPurchaseOrderBalance.orderbalance_report_po_item_viewRow
        Dim rowGRN As DseLaporanPurchaseOrderBalance.orderbalance_report_grnitem_viewRow
        Try
            Dim DTAGRN As New DseLaporanPurchaseOrderBalanceTableAdapters.orderbalance_report_grnitem_viewTableAdapter
            DTAGRN.FillBy_TanggalSampaiDari(DS.orderbalance_report_grnitem_view, Me.dtpTo.Value, Me.dtpFrom.Value)
            DTAGRN.Dispose()
            If DS.orderbalance_report_grnitem_view.Rows.Count > 0 Then
                For Each rowGRN In DS.orderbalance_report_grnitem_view.Rows
                    Dim TBL As DataTable = DS.orderbalance_report_po_item_view
                    Dim View As DataView = TBL.DefaultView
                    Dim cmd As String = "[Primary PO] =" + CStr(rowGRN.Primary_PO) + _
                    " and [Discount] =" + CStr(rowGRN.Discount) + _
                    " and [Primary Product] =" + CStr(rowGRN.Primary_Product) + _
                    " and [Unit Price] =" + CStr(rowGRN.Unit_Price) + _
                    " and [Nett Price] =" + CStr(rowGRN.Nett_Price) + _
                    " and [Primary IPO] > -1"
                    Dim rowview As DataRowView

                    View.RowFilter = cmd
                    View.Sort = "[Primary IPO] ASC"

                    For Each rowview In View
                        rowPO = rowview.Row
                        If rowGRN.Total_Qty = 0 Then
                            Exit For
                        End If
                        If rowPO.Total_Qty >= rowGRN.Total_Qty Then
                            rowPO.Total_Qty -= rowGRN.Total_Qty
                            rowGRN.Total_Qty -= rowGRN.Total_Qty
                        Else
                            'Jika Qty PO Lebih Kecil maka Kurangi dengan diri sendiri agar menjadi Nol Kemudian biarkan 
                            'Agar Item Selanjutnya yang di kurangkan
                            rowGRN.Total_Qty -= rowPO.Total_Qty
                            rowPO.Total_Qty -= rowPO.Total_Qty
                        End If
                    Next

                    'View Kedua
                    cmd = "[Primary PO] =" + CStr(rowGRN.Primary_PO) + _
                    " and Discount =" + CStr(rowGRN.Discount) + _
                    " and [Primary Product] =" + CStr(rowGRN.Primary_Product) + _
                    " and [Unit Price] =" + CStr(rowGRN.Unit_Price) + _
                    " and [Nett Price] =" + CStr(rowGRN.Nett_Price) + _
                    " and [Primary IPO] =-1"
                    View.RowFilter = cmd
                    For Each rowview In View
                        rowPO = rowview.Row
                        If rowGRN.Total_Qty = 0 Then
                            Exit For
                        End If
                        If rowPO.Total_Qty >= rowGRN.Total_Qty Then
                            rowPO.Total_Qty -= rowGRN.Total_Qty
                            rowGRN.Total_Qty -= rowGRN.Total_Qty
                        Else
                            'Jika Qty PO Lebih Kecil maka Kurangi dengan diri sendiri agar menjadi Nol Kemudian biarkan 
                            'Agar Item Selanjutnya yang di kurangkan
                            rowGRN.Total_Qty -= rowPO.Total_Qty
                            rowPO.Total_Qty -= rowPO.Total_Qty
                        End If
                    Next
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LakukanPenguranganDenganGRN(ByRef DS As DseLaporanPurchaseOrderBalance)
        Dim rowPO As DseLaporanPurchaseOrderBalance.orderbalance_report_po_item_viewRow
        Dim rowGRN As DseLaporanPurchaseOrderBalance.orderbalance_report_grnitem_viewRow
        Try
            Dim DTAGRN As New DseLaporanPurchaseOrderBalanceTableAdapters.orderbalance_report_grnitem_viewTableAdapter
            DTAGRN.FillBy_TanggalSampaiDari(DS.orderbalance_report_grnitem_view, Me.dtpTo.Value, Me.dtpFrom.Value)
            DTAGRN.Dispose()
            If DS.orderbalance_report_grnitem_view.Rows.Count > 0 Then
                For Each rowPO In DS.orderbalance_report_po_item_view.Rows
                    For Each rowGRN In DS.orderbalance_report_grnitem_view.Rows
                        If rowPO.Primary_PO = rowGRN.Primary_PO And rowPO.Discount = rowGRN.Discount And _
                        rowPO.Primary_Product = rowGRN.Primary_Product And rowPO.Unit_Price = rowGRN.Unit_Price And _
                        rowPO.Nett_Price = rowGRN.Nett_Price Then
                            If rowPO.Total_Qty >= rowGRN.Total_Qty Then
                                rowPO.Total_Qty -= rowGRN.Total_Qty
                                rowGRN.Total_Qty -= rowGRN.Total_Qty
                            Else
                                'Jika Qty PO Lebih Kecil maka Kurangi dengan diri sendiri agar menjadi Nol Kemudian biarkan 
                                'Agar Item Selanjutnya yang di kurangkan
                                rowPO.Total_Qty -= rowPO.Total_Qty
                                rowGRN.Total_Qty -= rowPO.Total_Qty
                            End If
                        End If
                    Next
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LakukanPenambahanDenganPurchaseReturn1(ByRef DS As DseLaporanPurchaseOrderBalance)
        Dim rowPO As DseLaporanPurchaseOrderBalance.orderbalance_report_po_item_viewRow
        Dim rowPurchaseReturn As DseLaporanPurchaseOrderBalance.orderbalance_report_return_item_viewRow
        Try
            Dim DTAPurchaseReturn As New DseLaporanPurchaseOrderBalanceTableAdapters.orderbalance_report_return_item_viewTableAdapter
            DTAPurchaseReturn.FillBy_TglDariSampai(DS.orderbalance_report_return_item_view, Me.dtpFrom.Value, Me.dtpTo.Value)
            DTAPurchaseReturn.Dispose()
            If DS.orderbalance_report_return_item_view.Rows.Count > 0 Then
                For Each rowPurchaseReturn In DS.orderbalance_report_return_item_view.Rows
                    Dim TBL As DataTable = DS.orderbalance_report_po_item_view
                    Dim View As DataView = TBL.DefaultView
                    Dim cmd As String = ""
                    cmd = "[Primary PO] =" + CStr(rowPurchaseReturn.Primary_PO) + _
                    " and Discount =" + CStr(rowPurchaseReturn.Discount) + _
                    " and [Primary Product] =" + CStr(rowPurchaseReturn.Primary_Product) + _
                    " and [Unit Price] =" + CStr(rowPurchaseReturn.Unit_Price) + _
                    " and [Nett Price] =" + CStr(rowPurchaseReturn.Nett_Price) + _
                    " and [Primary IPO] =-1"
                    Dim rowview As DataRowView
                    View.RowFilter = cmd

                    For Each rowview In View
                        rowPO = rowview.Row
                        If rowPurchaseReturn.Total_Qty = 0 Then
                            Exit For
                        End If
                        rowPO.Total_Qty += rowPurchaseReturn.Total_Qty
                        rowPurchaseReturn.Total_Qty -= rowPurchaseReturn.Total_Qty
                    Next

                    'View Kedua
                    cmd = "[Primary PO] =" + CStr(rowPurchaseReturn.Primary_PO) + _
                    " and [Discount] =" + CStr(rowPurchaseReturn.Discount) + _
                    " and [Primary Product] =" + CStr(rowPurchaseReturn.Primary_Product) + _
                    " and [Unit Price] =" + CStr(rowPurchaseReturn.Unit_Price) + _
                    " and [Nett Price] =" + CStr(rowPurchaseReturn.Nett_Price) + _
                    " and [Primary IPO] > -1"
                    View.RowFilter = cmd
                    View.Sort = "[Primary IPO] DESC"
                    For Each rowview In View
                        rowPO = rowview.Row
                        If rowPurchaseReturn.Total_Qty = 0 Then
                            Exit For
                        End If
                        rowPO.Total_Qty += rowPurchaseReturn.Total_Qty
                        rowPurchaseReturn.Total_Qty -= rowPurchaseReturn.Total_Qty
                    Next
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LakukanPenambahanDenganPurchaseReturn(ByRef DS As DseLaporanPurchaseOrderBalance)
        Dim rowPO As DseLaporanPurchaseOrderBalance.orderbalance_report_po_item_viewRow
        Dim rowPurchaseReturn As DseLaporanPurchaseOrderBalance.orderbalance_report_return_item_viewRow
        Try
            Dim DTAPurchaseReturn As New DseLaporanPurchaseOrderBalanceTableAdapters.orderbalance_report_return_item_viewTableAdapter
            DTAPurchaseReturn.FillBy_TglDariSampai(DS.orderbalance_report_return_item_view, Me.dtpFrom.Value, Me.dtpTo.Value)
            DTAPurchaseReturn.Dispose()
            If DS.orderbalance_report_return_item_view.Rows.Count > 0 Then
                For Each rowPO In DS.orderbalance_report_po_item_view.Rows
                    For Each rowPurchaseReturn In DS.orderbalance_report_return_item_view.Rows
                        If rowPO.Primary_PO = rowPurchaseReturn.Primary_PO And rowPO.Discount = rowPurchaseReturn.Discount And _
                        rowPO.Primary_Product = rowPurchaseReturn.Primary_Product And rowPO.Unit_Price = rowPurchaseReturn.Unit_Price And _
                        rowPO.Nett_Price = rowPurchaseReturn.Nett_Price Then
                            rowPO.Total_Qty += rowPurchaseReturn.Total_Qty
                            rowPurchaseReturn.Total_Qty -= rowPurchaseReturn.Total_Qty
                        End If
                    Next
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HitungKembaliNilaiTotalAmount(ByRef DS As DseLaporanPurchaseOrderBalance)
        Dim rowPO As DseLaporanPurchaseOrderBalance.orderbalance_report_po_item_viewRow
        Try
            For Each rowPO In DS.orderbalance_report_po_item_view.Rows
                rowPO.Amount = rowPO.Total_Qty * rowPO.Nett_Price
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSupplierCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSupplierCode.LostFocus
        txtSupplierCode.BackColor = Me.txtPrimarySupplier.BackColor
    End Sub

    Private Sub chkPeriod_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkPeriod.CheckedChanged
        Try
            If Me.chkPeriod.Checked = True Then
                Me.dtpFrom.Enabled = True
                Me.dtpTo.Enabled = True
            Else
                Me.dtpFrom.Enabled = False
                Me.dtpTo.Enabled = False
                Me.dtpFrom.Value = Today
                Me.dtpTo.Value = Today
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class