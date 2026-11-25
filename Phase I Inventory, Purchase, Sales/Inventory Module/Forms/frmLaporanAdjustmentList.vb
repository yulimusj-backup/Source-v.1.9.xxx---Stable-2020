Option Explicit On
Option Strict Off

Public Class frmLaporanAdjustmentList
    Dim TBLBrand As New Core_Module.dseMasterProductGroup.tblmaster_product_group2DataTable
    Dim TBLSetting As New Inventory_Module.dseTransaksiStockInOut.tbl_setting_adjustmentDataTable
    Dim rowSetting As Inventory_Module.dseTransaksiStockInOut.tbl_setting_adjustmentRow
    Dim FormatDay As Integer = 1
    Dim FormatMonth As Integer = 2
    Dim FormatYears As Integer = 1
    Dim FormatDecimal As Integer = 2
    Dim FormatQty As Integer = 0
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim ObjAdjustmentType As New clsTransaksiStockInOut

    Public UserCode As String = ""

    Dim DTAClosing1 As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchase_closingTableAdapter

    Dim DTAClosing As New dsLaporanAdjustmentListTableAdapters.DataTable1TableAdapter
    Dim TBLClosing As New dsLaporanAdjustmentList.DataTable1DataTable
    Dim TanggalClosing As Date = #1/1/1950#
    Dim DTA_CL As New DseTransaksiMonthlyClosingTableAdapters.DataTable25TableAdapter
    Dim TBL_CL As New DseTransaksiMonthlyClosing.DataTable25DataTable
    'Dim TBLAdj As New dsLaporanAdjustmentList.tbltrans_purchaseclosing_item5DataTable
    'Dim DTAAdj As New dsLaporanAdjustmentListTableAdapters.tbltrans_purchaseclosing_item5TableAdapter
    Dim closingPeriod As Date = #1/1/1950#
    Dim ObjLicense As New Core_Module.ClsLicense
    Public DSCekOut As dsLaporanAdjustmentList

    Public Enum EnumTipeDokumenAcuan As Integer
        StockInAdjust
        StockTakingInAdjustment
        StockBundlingAdjustment
    End Enum

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Function ApakahBolehCetak() As Boolean
        Try
            If Me.clbBrand.CheckedItems.Count <= 0 Then
                MsgBox("Please select at least 1 Brand.", MsgBoxStyle.Information, "Information")
                Me.clbBrand.Focus()
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

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

            If Me.ApakahBolehCetak = True Then
                If Me.cboTipeAdjustment.SelectedIndex = 0 Then
                    Me.Cetak_InAdjust()
                Else
                    Me.Cetak_Out_Adjust()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetTanggalTransaksi(ByVal vPrimaryDocumentAcuan As Double, _
    ByVal vTipeDocumenAcuan As EnumTipeDokumenAcuan) As Date
        Try
            If vTipeDocumenAcuan = EnumTipeDokumenAcuan.StockBundlingAdjustment Then
                Dim ObjBDL As New ClsTransaksiStockBundling
                Return ObjBDL.GetTanggalTransaksi(vPrimaryDocumentAcuan)
                ObjBDL = Nothing
            ElseIf vTipeDocumenAcuan = EnumTipeDokumenAcuan.StockInAdjust Then
                Dim ObjStockIn As New clsTransaksiStockInOut
                Return ObjStockIn.GetTanggalTransaksi(vPrimaryDocumentAcuan)
                ObjStockIn = Nothing
            ElseIf vTipeDocumenAcuan = EnumTipeDokumenAcuan.StockTakingInAdjustment Then
                Dim ObjStockTaking As New clsTransaksiStockTaking
                Return ObjStockTaking.GetStartDate(vPrimaryDocumentAcuan)
                ObjStockTaking = Nothing
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub Cetak_InAdjust()
        Dim DS As New dsLaporanAdjustmentList
        Dim DTA As New dsLaporanAdjustmentListTableAdapters.tbltrans_purchase_closing_item_detailTableAdapter
        Dim report As New CRPLaporanAdjustmentList
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim row As dsLaporanAdjustmentList.laporan_AdjustmentList_view0Row
        'Dim row1 As dsLaporanAdjustmentList.tbltrans_purchase_closing_item_detailRow
        Dim DTACLosing As New dsLaporanAdjustmentListTableAdapters.tbltrans_purchase_closingTableAdapter
        Dim ObjCD As New Core_Module.ClsCekData
        Dim DV As Data.DataRowView
        Dim primaryBrand As Double = 0
        Dim PrimaryClosing As Double = -1
        Dim ClosingPeriod As Date = ObjCD.GetFirstDateofTheMonth(Me.dtpFrom.Value)
        Dim ObjClosing_Acc As New Core_Module.ClsTranskasiAccClosing
        Dim TipeTransaksiAsal As ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust
        Dim TipeTransaksiAsal_StockTaking As ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InStockTakeAdjust
        Dim TipeItemDetail As ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
        Dim TBlInStockTaking As New dsLaporanAdjustmentList.tbltrans_purchase_closing_item_detailDataTable
        Dim rowInStockTaking As dsLaporanAdjustmentList.tbltrans_purchase_closing_item_detailRow
        Dim RowInAdjust As dsLaporanAdjustmentList.tbltrans_purchase_closing_item_detailRow
        Dim TblDummy As New dsLaporanAdjustmentList.laporan_AdjustmentList_view0DataTable
        Dim rowDummy As dsLaporanAdjustmentList.laporan_AdjustmentList_view0Row
        Dim ListPrimaryAdjustmentMasukLaporan As New List(Of Double)

        Try
            For Each DV In Me.clbAdjustmentType.CheckedItems
                Me.rowSetting = DV.Row
                ListPrimaryAdjustmentMasukLaporan.Add(Me.rowSetting.Primary)
            Next

            DTACLosing.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Connection.ConnectionString = DTACLosing.Connection.ConnectionString

            If DTACLosing.ScalarQuery_GetPrimaryClosing(ClosingPeriod).HasValue = True Then
                PrimaryClosing = DTACLosing.ScalarQuery_GetPrimaryClosing(ClosingPeriod).Value
            End If
            If PrimaryClosing = -1 Then
                MsgBox("In adjustment not available, please do purchase closing for " + Format(ClosingPeriod, "MMMMM yyyy") + " first.", _
                MsgBoxStyle.Information, _
                "Information")
                Exit Sub
            End If

            'stock in adjust
            DS.laporan_AdjustmentList_view0.Rows.Clear()
            DTA.ClearBeforeFill = True

            If Me.chkAdjust.Checked = True Then
                For Each DV In Me.clbBrand.CheckedItems
                    primaryBrand = DV.Item(TBLBrand.PrimaryColumn.ColumnName)
                    DTA.Fill(DS.tbltrans_purchase_closing_item_detail, PrimaryClosing, TipeTransaksiAsal, TipeItemDetail, primaryBrand)

                    'tambahkan dengan adjustment stock taking --> hati2 saat edit, walau menggunakan datatable yg sama, tapi query-nya beda
                    'karena no dokumen dan tanggal dokumennya diambil dari stock taking
                    DTA.FillBy_StockTaking(TBlInStockTaking, PrimaryClosing, TipeTransaksiAsal_StockTaking, TipeItemDetail, primaryBrand)
                    For Each rowInStockTaking In TBlInStockTaking.Rows
                        RowInAdjust = DS.tbltrans_purchase_closing_item_detail.Newtbltrans_purchase_closing_item_detailRow
                        For ColCount As Integer = 0 To TBlInStockTaking.Columns.Count - 1
                            RowInAdjust.Item(ColCount) = rowInStockTaking.Item(ColCount)
                        Next
                        DS.tbltrans_purchase_closing_item_detail.Rows.Add(RowInAdjust)
                    Next

                    For Each RowInAdjust In DS.tbltrans_purchase_closing_item_detail.Rows
                        'hanya masukkan yang adjustment type nya tertentu saja
                        If ListPrimaryAdjustmentMasukLaporan.Contains(RowInAdjust.primary_adjustment_type) = True Then
                            row = DS.laporan_AdjustmentList_view0.Newlaporan_AdjustmentList_view0Row
                            With RowInAdjust
                                row.Brand = .Brand
                                row.Document_No = .No_Dokumen_Acuan
                                row.Primary = .Primary
                                row.Primary_Product = .Primary_Code_No
                                row.PrimaryProductGroup = .Primary_Brand
                                row.Product_Code = .Product_Code
                                row.Quantity = .Qty_In_Adjust
                                If .Adjustment_Code.Trim = "" Then
                                    row.Remarks = .Remarks_Header
                                Else
                                    If .Remarks_Header.Trim <> "" Then
                                        row.Remarks = .Adjustment_Code + ": " + .Remarks_Header
                                    Else
                                        row.Remarks = .Adjustment_Code
                                    End If
                                End If
                                row.Unit_Code = .Unit_Code
                                row._Date = .Tanggal_Transaksi
                                row.Cost_Amount = .amount
                                row._Decimal = ._Decimal
                                row.Adjustment_Type_Code = .Adjustment_Code
                            End With
                            DS.laporan_AdjustmentList_view0.Rows.Add(row)
                        End If
                    Next
                Next
            End If

            'masukkan data dari bundling
            If Me.chkBundling.Checked = True Then
                Me.IsiDgvBundling(DS.laporan_AdjustmentList_view0)
            End If

            'urutkan ulang berdasarkan Brand, Tipe Adjustment, Tgl Adjustment
            DS.laporan_AdjustmentList_view0.DefaultView.RowFilter = ""
            DS.laporan_AdjustmentList_view0.DefaultView.Sort = "Brand, [Adjustment Type Code], Date, [Document No], [Product Code]"
            For Each DV In DS.laporan_AdjustmentList_view0.DefaultView
                row = DV.Row
                rowDummy = TblDummy.Newlaporan_AdjustmentList_view0Row
                For wj As Integer = 0 To TblDummy.Columns.Count - 1
                    rowDummy.Item(wj) = row.Item(wj)
                Next
                TblDummy.Rows.Add(rowDummy)
            Next
            DS.laporan_AdjustmentList_view0.Rows.Clear()
            For Each rowDummy In TblDummy.Rows
                row = DS.laporan_AdjustmentList_view0.Newlaporan_AdjustmentList_view0Row
                For wj As Integer = 0 To TblDummy.Columns.Count - 1
                    row.Item(wj) = rowDummy.Item(wj)
                Next
                DS.laporan_AdjustmentList_view0.Rows.Add(row)
            Next

            Dim Brand As String = ""
            If Me.clbBrand.CheckedItems.Count = Me.clbBrand.Items.Count Then
                Brand = "All"
            Else
                For Each DV In Me.clbBrand.CheckedItems
                    If Brand = "" Then
                        Brand = DV.Item(1)
                    Else
                        Brand += ", " + DV.Item(1)
                    End If
                Next
            End If

            report.SetDataSource(DS)
            report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, Me.FormatDay)
            report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, Me.FormatMonth)
            report.SetParameterValue(report.Parameter_FormatYears.ParameterFieldName, Me.FormatYears)
            report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, Me.FormatDecimal)
            report.SetParameterValue(report.Parameter_FormatQTY.ParameterFieldName, Me.FormatQty)
            report.SetParameterValue(report.Parameter_Tanggal.ParameterFieldName, Me.dtpFrom.Value)
            report.SetParameterValue(report.Parameter_FilterBYBrand.ParameterFieldName, Brand)
            report.SetParameterValue(report.Parameter_Tipe.ParameterFieldName, "IN")

            Me.CrystalReportViewer1.ReportSource = report
            Me.CrystalReportViewer1.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTA.Dispose()
            DTACLosing.Dispose()
            ObjCD = Nothing
            ObjClosing_Acc = Nothing
            TBlInStockTaking.Dispose()
            TblDummy.Dispose()
            ListPrimaryAdjustmentMasukLaporan = Nothing
        End Try
    End Sub

    Private Sub IsiDgvBundling(ByVal TBL As dsLaporanAdjustmentList.laporan_AdjustmentList_view0DataTable)
        Dim Filter As String = ""
        Dim TanggalSampai As Date = Me.ObjCD.GetLasDateOfTheMonth(Me.dtpFrom.Value.Date)
        Dim TanggalDari As Date = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom.Value.Date)
        Dim rowAdjust As dsLaporanAdjustmentList.laporan_AdjustmentList_view0Row
        Dim PrimaryClosing As Double = -1

        Dim Filter2 As String = ""
        Dim counter As Double = 0
        Dim DV As Data.DataRowView
        Dim primaryBrand As Double = -1

        Dim DTA_After1 As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchase_closing_item_afterTableAdapter
        Dim TBL_After1 As New DseTransaksiMonthlyClosing.tbltrans_purchase_closing_item_afterDataTable
        Dim row_After1, row_After2 As DseTransaksiMonthlyClosing.tbltrans_purchase_closing_item_afterRow
        Dim TBL_After1_Dummy As New DseTransaksiMonthlyClosing.tbltrans_purchase_closing_item_afterDataTable
        Dim PeriodClosing As Date = Today

        Try
            PeriodClosing = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom.Value.Date)
            If Me.DTAClosing1.ScalarQuery_GetPrimary(PeriodClosing).HasValue = True Then
                PrimaryClosing = Me.DTAClosing1.ScalarQuery_GetPrimary(PeriodClosing).Value
            End If
            If PrimaryClosing = -1 Then
                Exit Sub
            End If

            'isikan semua item bundling
            DTA_After1.Connection.ConnectionString = Core_Module.GetConstring
            TBL_After1.Rows.Clear()
            DTA_After1.Fill(TBL_After1_Dummy, TanggalSampai, TanggalDari, CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling), _
            CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem), PrimaryClosing)

            For Each DV In Me.clbBrand.CheckedItems
                primaryBrand = DV.Item(TBLBrand.PrimaryColumn.ColumnName)
                For Each row_After1 In TBL_After1_Dummy.Rows
                    If row_After1.Primary_Brand = primaryBrand Then
                        row_After2 = TBL_After1.Newtbltrans_purchase_closing_item_afterRow

                        For k As Integer = 0 To TBL_After1.Columns.Count - 1
                            row_After2.Item(k) = row_After1.Item(k)
                        Next

                        TBL_After1.Rows.Add(row_After2)
                    End If
                Next
            Next

            'jumlahkan (unit price * faktor pemberat) untuk setiap item after yang ada
            For Each row_After1 In TBL_After1.Rows
                With row_After1
                    rowAdjust = TBL.Newlaporan_AdjustmentList_view0Row

                    rowAdjust.Brand = .Brand
                    'unit price disini adalah unit price sejumlah qty yang ada (bukan untuk setiap unitnya)
                    rowAdjust.Cost_Amount = .amount
                    rowAdjust.Document_No = .No_Bundling
                    rowAdjust.Primary = .primary_transaksi_bundling
                    rowAdjust.Primary_Product = .primary_code_no
                    rowAdjust.PrimaryProductGroup = .Primary_Brand
                    rowAdjust.Product_Code = .Product_Code
                    rowAdjust.Quantity = .qty
                    rowAdjust.Remarks = .Remarks_Header
                    rowAdjust.Unit_Code = .Unit_Code
                    rowAdjust._Date = .tgl_transaksi_bundling
                    rowAdjust._Decimal = ._Decimal
                    rowAdjust.Adjustment_Type_Code = " Bundling"

                    TBL.Rows.Add(rowAdjust)
                End With
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_After1.Dispose()
            TBL_After1.Dispose()
            TBL_After1_Dummy.Dispose()
        End Try
    End Sub

    ''' <summary>
    ''' out adjustment menggabungkan out adjust, stock out bundling, penyesuaian minus stock taking
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Cetak_Out_Adjust()
        Dim DS As New dsLaporanAdjustmentList
        Dim DTA As New dsLaporanAdjustmentListTableAdapters.tbltrans_purchase_closing_item_detailTableAdapter
        Dim report As New CRPLaporanAdjustmentList
        Dim DTABundling1 As New dsLaporanAdjustmentListTableAdapters.Tbl_Out_Bundling1TableAdapter
        Dim TBLBundling1Dummy As New dsLaporanAdjustmentList.Tbl_Out_Bundling1DataTable
        Dim rowBundling1 As dsLaporanAdjustmentList.Tbl_Out_Bundling1Row
        Dim PrimaryClosing As Double = -1
        Dim PeriodClosing As Date = Today
        Dim TipeItemClosing As ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
        Dim TipeTransaksiAsal As ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutAdjust
        Dim TipeTransaksiAsal1 As ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutStockTakeAdjust
        Dim DV As Data.DataRowView
        Dim primaryBrand As Double = 0
        Dim TipeItemDetail As ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
        Dim TBlInStockTaking As New dsLaporanAdjustmentList.tbltrans_purchase_closing_item_detailDataTable
        Dim rowOutStockTaking As dsLaporanAdjustmentList.tbltrans_purchase_closing_item_detailRow
        Dim RowOutAdjust As dsLaporanAdjustmentList.tbltrans_purchase_closing_item_detailRow
        Dim Row As dsLaporanAdjustmentList.laporan_AdjustmentList_view0Row
        Dim TglDari As Date = ObjCD.GetFirstDateofTheMonth(Me.dtpFrom.Value)
        Dim TglSampai As Date = ObjCD.GetLasDateOfTheMonth(Me.dtpFrom.Value)
        Dim TblDummy As New dsLaporanAdjustmentList.laporan_AdjustmentList_view0DataTable
        Dim rowDummy As dsLaporanAdjustmentList.laporan_AdjustmentList_view0Row
        Dim ListPrimaryAdjustmentMasukLaporan As New List(Of Double)

        Try
            For Each DV In Me.clbAdjustmentType.CheckedItems
                Me.rowSetting = DV.Row
                ListPrimaryAdjustmentMasukLaporan.Add(Me.rowSetting.Primary)
            Next

            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            'con string
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'period
            PeriodClosing = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom.Value.Date)
            If Me.DTAClosing1.ScalarQuery_GetPrimary(PeriodClosing).HasValue = True Then
                PrimaryClosing = Me.DTAClosing1.ScalarQuery_GetPrimary(PeriodClosing).Value
            End If
            If PrimaryClosing = -1 Then
                Exit Sub
            End If

            'stock out adjust
            If Me.chkAdjust.Checked = True Then
                For Each DV In Me.clbBrand.CheckedItems
                    primaryBrand = DV.Item(TBLBrand.PrimaryColumn.ColumnName)
                    DTA.Fill(DS.tbltrans_purchase_closing_item_detail, PrimaryClosing, TipeTransaksiAsal, TipeItemDetail, primaryBrand)

                    'tambahkan dengan adjustment stock taking --> hati2 saat edit, walau menggunakan datatable yg sama, tapi query-nya beda
                    'karena no dokumen dan tanggal dokumennya diambil dari stock taking
                    DTA.FillBy_StockTaking(TBlInStockTaking, PrimaryClosing, TipeTransaksiAsal1, TipeItemDetail, primaryBrand)
                    For Each rowOutStockTaking In TBlInStockTaking.Rows
                        RowOutAdjust = DS.tbltrans_purchase_closing_item_detail.Newtbltrans_purchase_closing_item_detailRow
                        For ColCount As Integer = 0 To TBlInStockTaking.Columns.Count - 1
                            RowOutAdjust.Item(ColCount) = rowOutStockTaking.Item(ColCount)
                        Next
                        DS.tbltrans_purchase_closing_item_detail.Rows.Add(RowOutAdjust)
                    Next

                    For Each RowOutAdjust In DS.tbltrans_purchase_closing_item_detail.Rows
                        'hanya memasukkan yang tipe adjustmennya sama

                        Row = DS.laporan_AdjustmentList_view0.Newlaporan_AdjustmentList_view0Row
                        If ListPrimaryAdjustmentMasukLaporan.Contains(RowOutAdjust.primary_adjustment_type) = True Then
                            With RowOutAdjust
                                Row.Brand = .Brand
                                Row.Document_No = .No_Dokumen_Acuan
                                Row.Primary = .Primary
                                Row.Primary_Product = .Primary_Code_No
                                Row.PrimaryProductGroup = .Primary_Brand
                                Row.Product_Code = .Product_Code
                                Row.Quantity = .Qty_In_Adjust
                                If .Adjustment_Code.Trim = "" Then
                                    Row.Remarks = .Remarks_Header
                                Else
                                    If .Remarks_Header.Trim <> "" Then
                                        Row.Remarks = .Adjustment_Code + ": " + .Remarks_Header
                                    Else
                                        Row.Remarks = .Adjustment_Code
                                    End If
                                End If
                                Row.Unit_Code = .Unit_Code
                                Row._Date = .Tanggal_Transaksi
                                Row.Cost_Amount = .amount
                                Row._Decimal = ._Decimal
                                Row.Adjustment_Type_Code = .Adjustment_Code

                            End With
                            DS.laporan_AdjustmentList_view0.Rows.Add(Row)
                        End If
                    Next
                Next
            End If

            'stock out bundling
            If Me.chkBundling.Checked = True Then
                DTABundling1.Fill(TBLBundling1Dummy, TglSampai, TglDari, PrimaryClosing, _
                CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem), _
                CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling))

                For Each DV In Me.clbBrand.CheckedItems
                    primaryBrand = DV.Item(TBLBrand.PrimaryColumn.ColumnName)

                    For Each rowBundling1 In TBLBundling1Dummy.Rows
                        If rowBundling1.Primary_Brand = primaryBrand Then
                            Row = DS.laporan_AdjustmentList_view0.Newlaporan_AdjustmentList_view0Row
                            With rowBundling1
                                Row.Primary = .primary
                                Row.Brand = .Brand
                                Row.Cost_Amount = .amount
                                Row.Document_No = .No_Bundling
                                Row.Primary_Product = .primary_code_no
                                Row.PrimaryProductGroup = .Primary_Brand
                                Row.Product_Code = .Product_Code
                                Row.Quantity = .qty
                                 Row.Remarks = .Remarks_Header
                                Row.Unit_Code = .Unit_Code
                                Row._Date = .tgl_transaksi_bundling
                                Row._Decimal = ._Decimal
                                Row.Adjustment_Type_Code = " Bundling"
                            End With

                            DS.laporan_AdjustmentList_view0.Rows.Add(Row)
                        End If
                    Next
                Next
            End If

            'urutkan ulang berdasarkan Brand, Tipe Adjustment, Tgl Adjustment
            DS.laporan_AdjustmentList_view0.DefaultView.RowFilter = ""
            DS.laporan_AdjustmentList_view0.DefaultView.Sort = "Brand, [Adjustment Type Code], Date, [Document No], [Product Code]"
            For Each DV In DS.laporan_AdjustmentList_view0.DefaultView
                Row = DV.Row
                rowDummy = TblDummy.Newlaporan_AdjustmentList_view0Row
                For wj As Integer = 0 To TblDummy.Columns.Count - 1
                    rowDummy.Item(wj) = Row.Item(wj)
                Next
                TblDummy.Rows.Add(rowDummy)
            Next
            DS.laporan_AdjustmentList_view0.Rows.Clear()
            For Each rowDummy In TblDummy.Rows
                Row = DS.laporan_AdjustmentList_view0.Newlaporan_AdjustmentList_view0Row
                For wj As Integer = 0 To TblDummy.Columns.Count - 1
                    Row.Item(wj) = rowDummy.Item(wj)
                Next
                DS.laporan_AdjustmentList_view0.Rows.Add(Row)
            Next

            Dim Brand As String = ""
            If Me.clbBrand.CheckedItems.Count = Me.clbBrand.Items.Count Then
                Brand = "All"
            Else
                For Each DV In Me.clbBrand.CheckedItems
                    If Brand = "" Then
                        Brand = DV.Item(1)
                    Else
                        Brand += ", " + DV.Item(1)
                    End If
                Next
            End If

            report.SetDataSource(DS)
            report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, Me.FormatDay)
            report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, Me.FormatMonth)
            report.SetParameterValue(report.Parameter_FormatYears.ParameterFieldName, Me.FormatYears)
            report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, Me.FormatDecimal)
            report.SetParameterValue(report.Parameter_FormatQTY.ParameterFieldName, Me.FormatQty)
            report.SetParameterValue(report.Parameter_Tanggal.ParameterFieldName, Me.dtpFrom.Value)
            report.SetParameterValue(report.Parameter_FilterBYBrand.ParameterFieldName, Brand)
            report.SetParameterValue(report.Parameter_Tipe.ParameterFieldName, "OUT")

            Me.CrystalReportViewer1.ReportSource = report
            Me.CrystalReportViewer1.Show()
        Catch ex As Exception
            MsgBox(ex.Message + " : primary brand " + primaryBrand.ToString, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            TblDummy.Dispose()
            ListPrimaryAdjustmentMasukLaporan = Nothing
        End Try
    End Sub

    Private Sub CheckAllBrand()
        Try
            For i As Integer = 0 To Me.clbBrand.Items.Count - 1
                Me.clbBrand.SetItemCheckState(i, CheckState.Checked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Me.CheckAllBrand()
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

    Private Sub frmLaporanAdjustmentList_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim OBjBrand As New Core_Module.ClsMasterProductGroup
        Try
            Me.FormatDecimal = Me.ObjCur.GetDecimal_DefaultCurrency
            OBjBrand.IsiListBoxBrand(Me.clbBrand, Me.TBLBrand)
            Me.cboTipeAdjustment.SelectedIndex = 0
            Me.dtpFrom.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpFrom.Value)
            Me.ObjAdjustmentType.IsiClbSetting(Me.TBLSetting, Me.clbAdjustmentType)
            Me.CheckAllBrand()
            Me.CheckAllAdjustmentType()
            Me.Text = "Adjustment List"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            OBjBrand = Nothing
        End Try
    End Sub

    Private Sub CheckAllAdjustmentType()
        Try
            For i As Integer = 0 To Me.clbAdjustmentType.Items.Count - 1
                Me.clbAdjustmentType.SetItemCheckState(i, CheckState.Checked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel4_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel4.LinkClicked
        Me.CheckAllAdjustmentType()
    End Sub

    Private Sub LinkLabel3_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel3.LinkClicked
        Try
            For i As Integer = 0 To Me.clbAdjustmentType.Items.Count - 1
                Me.clbAdjustmentType.SetItemCheckState(i, CheckState.Unchecked)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class