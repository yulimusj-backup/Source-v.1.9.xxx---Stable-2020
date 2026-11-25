Option Explicit On
Option Strict Off

Public Class ClsTransaksiMonthlyClosingItem
    Dim ObjCD As New Core_Module.ClsCekData

    Public Enum EnumTipeDokumenAcuan As Integer
        StockInAdjust
        StockTakingInAdjustment
        StockBundlingAdjustment
        StockInSalesReturnPeriodeSama
    End Enum

    Public Enum EnumTipeItemDetail As Integer
        BeginingBalanceItem
        InItem
        OutItem
        EndingBalanceItem
    End Enum

    Public Enum EnumTipeTranskasiAsal As Integer
        InPurchase '0
        InSR '1
        InAdjust '2
        InStockTakeAdjust '3
        OutBundling '4
        OutSales '5
        OutPR '6
        OutStockTakeAdjust '7
        OutAdjust '8
        InBundling '9
        OverheadCost '10
    End Enum

    Public Enum EnumCostingMethod As Integer
        Average
        FIFO
        LIFO
    End Enum

    Public Sub Get_BeginAmount_BeginQty_InQty_IAmount(ByVal vPrimaryProduct As Double, ByRef vBeginAmount As Decimal, _
     ByRef vBeginQty As Decimal, ByRef vInQty As Decimal, ByVal vTanggalSampai As Date, _
     ByVal vTanggalDari As Date, ByVal ObjCD As Core_Module.ClsCekData, ByVal TBL1 As DseTransaksiMonthlyClosing.DataTable16DataTable, _
     ByVal TBL2 As DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item2DataTable)
        Dim row1 As DseTransaksiMonthlyClosing.DataTable16Row
        Dim Filter As String = ""
        Dim row2 As DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item2Row

        Try
            Filter = "[Primary Product] = " + vPrimaryProduct.ToString
            TBL2.DefaultView.RowFilter = Filter
            If TBL2.DefaultView.Count > 0 Then
                row2 = TBL2.DefaultView(0).Row
                vInQty = row2.In_Qty
            Else
                vInQty = 0
            End If

            'mencari data begin amount n begin qty dari bulan sebelumnya
            TBL1.DefaultView.RowFilter = Filter
            If TBL1.DefaultView.Count = 0 Then
                vBeginQty = 0
                vBeginAmount = 0
            Else
                row1 = TBL1.DefaultView(0).Row
                vBeginQty = row1.Ending_Qty
                vBeginAmount = row1.Ending_Amount
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub Get_BeginAmount_BeginQty_InQty_IAmount(ByVal vPrimaryProduct As Double, ByRef vBeginAmount As Decimal, _
    ByRef vBeginQty As Decimal, ByRef vInQty As Decimal, ByVal vTanggalSampai As Date, _
    ByVal vTanggalDari As Date, ByVal ObjCD As Core_Module.ClsCekData)
        Dim DTA As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchaseclosing_item2TableAdapter
        Dim DTA1 As New DseTransaksiMonthlyClosingTableAdapters.DataTable16TableAdapter
        Dim TBL1 As New DseTransaksiMonthlyClosing.DataTable16DataTable
        Dim row1 As DseTransaksiMonthlyClosing.DataTable16Row

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString

            If DTA.ScalarQuery(vPrimaryProduct, vTanggalSampai, vTanggalDari).HasValue = True Then
                vInQty = DTA.ScalarQuery(vPrimaryProduct, vTanggalSampai, vTanggalDari).Value
            Else
                vInQty = 0
            End If

            'mencari data begin amount n begin qty dari bulan sebelumnya
            vTanggalSampai = vTanggalDari - TimeSpan.FromDays(1)
            vTanggalDari = ObjCD.GetFirstDateofTheMonth(vTanggalSampai)
            If DTA1.FillBy_ClosingPeriod(TBL1, vTanggalDari, vPrimaryProduct) > 0 Then
                row1 = TBL1.Rows(0)
                vBeginQty = row1.Ending_Qty
                vBeginAmount = row1.Ending_Amount
            Else
                vBeginQty = 0
                vBeginAmount = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
            TBL1.Dispose()
        End Try
    End Sub

    Public Function GetReference(ByVal vPrimaryTipeAdjustment As Double, ByVal vTipeTransaksiAsal As EnumTipeTranskasiAsal, _
    ByVal vTBL As DseTransaksiMonthlyClosing2.tbl_setting_adjustmentDataTable) As String
        Dim TipeAdjustment As String = ""
        Dim rowSetting As DseTransaksiMonthlyClosing2.tbl_setting_adjustmentRow

        Try
            If vTipeTransaksiAsal = EnumTipeTranskasiAsal.InAdjust Then
                vTBL.DefaultView.RowFilter = "Primary = " + vPrimaryTipeAdjustment.ToString
                If vTBL.DefaultView.Count = 0 Then
                    TipeAdjustment = ""
                Else
                    rowSetting = vTBL.DefaultView(0).Row
                    TipeAdjustment = rowSetting.Desc
                End If

                If TipeAdjustment = "" Then
                    Return "In Adjust"
                Else
                    Return "In Adjust - " + TipeAdjustment
                End If
            ElseIf vTipeTransaksiAsal = EnumTipeTranskasiAsal.InBundling Then
                Return "In Bundling"
            ElseIf vTipeTransaksiAsal = EnumTipeTranskasiAsal.InPurchase Then
                Return "In Purchase"
            ElseIf vTipeTransaksiAsal = EnumTipeTranskasiAsal.InSR Then
                Return "In Sales Return"
            ElseIf vTipeTransaksiAsal = EnumTipeTranskasiAsal.InStockTakeAdjust Then
                Return "In Stock Take"
            ElseIf vTipeTransaksiAsal = EnumTipeTranskasiAsal.OutAdjust Then
                vTBL.DefaultView.RowFilter = "Primary = " + vPrimaryTipeAdjustment.ToString
                If vTBL.DefaultView.Count = 0 Then
                    TipeAdjustment = ""
                Else
                    rowSetting = vTBL.DefaultView(0).Row
                    TipeAdjustment = rowSetting.Desc
                End If
                If TipeAdjustment = "" Then
                    Return "Out Adjust"
                Else
                    Return "Out Adjust - " + TipeAdjustment
                End If
            ElseIf vTipeTransaksiAsal = EnumTipeTranskasiAsal.OutBundling Then
                Return "Out Bundling"
            ElseIf vTipeTransaksiAsal = EnumTipeTranskasiAsal.OutPR Then
                Return "Out Purchase Return"
            ElseIf vTipeTransaksiAsal = EnumTipeTranskasiAsal.OutSales Then
                Return "Out Sales"
            ElseIf vTipeTransaksiAsal = EnumTipeTranskasiAsal.OutStockTakeAdjust Then
                Return "Out Stock Take"
            ElseIf vTipeTransaksiAsal = EnumTipeTranskasiAsal.OverheadCost Then
                Return "Overhead Cost"
            End If

            Return "n/a"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "n/a"
        End Try
    End Function

    Public Function IsiDGV(ByVal vDGVItem As DataGridView, ByVal vDGVItemAdjust As DataGridView, _
    ByVal vPrimaryTransaksi As Double, _
    ByVal vTBLDetail As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailDataTable, _
    ByVal vTBLBundling As DseTransaksiMonthlyClosing2.TblBundling1DataTable, _
    ByVal i As FrmTransaksiMonthlyClosing2, _
    ByVal vDGVAkun As DataGridView, ByVal vTBLSettingItem As DseTransaksiMonthlyClosing2.tbltrans_purchase_closing_item_tipe_adjustmentDataTable, _
    ByVal vDGVCostGroup As DataGridView, ByVal vTBLCostGroup_Account As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_accountDataTable, _
    ByVal vTBLCostGroup_Scope As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_scopeDataTable, _
    ByVal vTBLCostGroup_CodeNo As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_code_noDataTable, _
    ByVal vFrm As FrmTransaksiMonthlyClosing2, _
    ByVal vTBLCostGroup_CodeNo2 As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_code_no2DataTable, _
    ByVal vTblProductQty As Inventory_Module.DseLaporanStockDetail2.TblCetakDataTable) As Boolean
        Dim DTA As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_puchaseclosing_item0TableAdapter
        Dim TBL As New DseTransaksiMonthlyClosing.tbltrans_puchaseclosing_item0DataTable
        Dim row As DseTransaksiMonthlyClosing.tbltrans_puchaseclosing_item0Row
        Dim DTA1 As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchaseclosing_item5TableAdapter
        Dim TBL1 As New DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item5DataTable
        Dim row1 As DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item5Row
        Dim ObjIO As New Inventory_Module.clsTransaksiStockInOut
        Dim rowDGV As DataGridViewRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim DefPrec As Integer = 0

        Dim DTADetail As New DseMonthlyClosingNewTableAdapters.tbltrans_purchase_closing_item_detailTableAdapter
        Dim vRowDetail As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailRow
        Dim DTABundling As New DseTransaksiMonthlyClosing2TableAdapters.TblBundling1TableAdapter
        Dim DTAAfter As New DseTransaksiMonthlyClosing2TableAdapters.TblBundlingAfterTableAdapter
        Dim DTABefore As New DseTransaksiMonthlyClosing2TableAdapters.TblBundlingBeforeTableAdapter

        Dim DTAakun As New DseTransaksiMonthlyClosing2TableAdapters.tbltrans_purchase_closing_item_akun1TableAdapter
        Dim TBLAkun As New DseTransaksiMonthlyClosing2.tbltrans_purchase_closing_item_akun1DataTable
        Dim rowAkun As DseTransaksiMonthlyClosing2.tbltrans_purchase_closing_item_akun1Row

        Dim DTASetting As New DseTransaksiMonthlyClosing2TableAdapters.tbl_setting_adjustmentTableAdapter
        Dim TBLSetting As New DseTransaksiMonthlyClosing2.tbl_setting_adjustmentDataTable

        Dim DTASettingItem As New DseTransaksiMonthlyClosing2TableAdapters.tbltrans_purchase_closing_item_tipe_adjustmentTableAdapter
        'Dim rowSettingItem As DseTransaksiMonthlyClosing2.tbltrans_purchase_closing_item_tipe_adjustmentRow

        Try
            'vTBLBundling.Rows.Clear()
            'vTBLDetail.Rows.Clear()

            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTASetting.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTADetail.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTABundling.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTAAfter.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTABefore.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTAakun.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTASettingItem.Connection.ConnectionString = DTA.Connection.ConnectionString

            DefPrec = ObjCur.GetDecimal_DefaultCurrency
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            vDGVItem.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            vDGVItem.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            vDGVItemAdjust.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItemAdjust.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            vDGVItemAdjust.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            vDGVItemAdjust.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            vDGVAkun.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVAkun.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            vDGVAkun.Rows.Clear()

            'kosongkan data tabel
            vTBLDetail.Rows.Clear()
            vTBLBundling.Rows.Clear()
            vTBLSettingItem.Rows.Clear()

            'kosongkan dgv
            vDGVItem.Rows.Clear()
            vDGVItemAdjust.Rows.Clear()
            DefPrec = ObjCur.GetDecimal_DefaultCurrency

            'isikan tabel
            If DTA.Fill(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    With row
                        vDGVItem.Rows.Add((vDGVItem.Rows.Count + 1).ToString)
                        rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                        rowDGV.Cells(0).Value = CStr(vDGVItem.Rows.Count)
                        rowDGV.Cells("ColCodeNo").Value = .Product_Code
                        rowDGV.Cells("ColDescription").Value = .Description
                        rowDGV.Cells("ColUnits").Value = .Unit_Code
                        rowDGV.Cells("ColBeginQty").Value = ObjCD.DigitGrouping(.Begin_Qty, ._Decimal)
                        rowDGV.Cells("ColBeginAmount").Value = ObjCD.DigitGrouping(.Begin_Amount, DefPrec)
                        rowDGV.Cells("ColInQty").Value = ObjCD.DigitGrouping(.In_Qty, ._Decimal)
                        rowDGV.Cells("ColInAmount").Value = ObjCD.DigitGrouping(.In_Amount, DefPrec)
                        rowDGV.Cells("ColOutQty").Value = ObjCD.DigitGrouping(.Out_Qty, ._Decimal)
                        rowDGV.Cells("ColOutAmount").Value = ObjCD.DigitGrouping(.Out_Amount, DefPrec)
                        rowDGV.Cells("ColEndQty").Value = ObjCD.DigitGrouping(.Ending_Qty, ._Decimal)
                        rowDGV.Cells("ColEndAmount").Value = ObjCD.DigitGrouping(.Ending_Amount, DefPrec)
                        rowDGV.Cells("ColRemarks").Value = .Remarks
                        rowDGV.Cells("ColDetailCalculation").Value = "Show"
                        rowDGV.Cells("ColPrimaryUnits").Value = .Primary_Unit
                        rowDGV.Cells("ColPrimaryCodeNo").Value = .Primary_Product
                        rowDGV.Cells(i.ColItem_PrimaryAdjustmentAccount.Index).Value = .Primary_Adjustment_Account
                        rowDGV.Cells(i.ColItem_PrimaryCOGSAccount.Index).Value = .Primary_COGS_Account
                        rowDGV.Cells(i.ColItem_PrimarySalesAccount.Index).Value = .Primary_Sales_Account
                        rowDGV.Cells(i.ColItem_PrimaryInventoryAccount.Index).Value = .Primary_Inventory_Account
                        rowDGV.Cells(i.ColItem_PrimarySalesReturnAccount.Index).Value = .Primary_Sales_Return_Account
                        rowDGV.Cells(i.ColItem_PrimaryStockTakingAccount.Index).Value = .Primary_Stock_Taking_Account
                        rowDGV.Cells(i.ColItem_Decimal.Index).Value = ._Decimal
                    End With
                Next
            End If

            DTABundling.FillBy_PrimaryTransaksi(vTBLBundling, vPrimaryTransaksi)
            DTADetail.FillBy(vTBLDetail, vPrimaryTransaksi)
            DTASettingItem.FillBy_PrimaryTransaksi(vTBLSettingItem, vPrimaryTransaksi)

            If DTA1.FillBy_PrimaryTransaksi(TBL1, CDec(vPrimaryTransaksi)) > 0 Then
                vFrm.TblA.Rows.Clear()

                For Each row1 In TBL1.Rows
                    With row1
                        vDGVItemAdjust.Rows.Add((vDGVItemAdjust.Rows.Count + 1).ToString)
                        rowDGV = vDGVItemAdjust.Rows(vDGVItemAdjust.Rows.Count - 1)
                        rowDGV.Cells(0).Value = CStr(vDGVItemAdjust.Rows.Count)
                        rowDGV.Cells("ColCodeNoAdjust").Value = .Product_Code
                        rowDGV.Cells("ColJenisDokumen").Value = .Tipe_Dokumen_Acuan
                        If .Tipe_Dokumen_Acuan = ClsTransaksiMonthlyClosingItem.EnumTipeDokumenAcuan.StockInAdjust Then
                            rowDGV.Cells("ColJenisDocumentAdjustStr").Value = "Stock In Adjust"
                        ElseIf .Tipe_Dokumen_Acuan = ClsTransaksiMonthlyClosingItem.EnumTipeDokumenAcuan.StockTakingInAdjustment Then
                            rowDGV.Cells("ColJenisDocumentAdjustStr").Value = "Stock Taking In Adjustment"
                        ElseIf .Tipe_Dokumen_Acuan = ClsTransaksiMonthlyClosingItem.EnumTipeDokumenAcuan.StockInSalesReturnPeriodeSama Then
                            rowDGV.Cells("ColJenisDocumentAdjustStr").Value = "Stock In Sales Return"
                            .No_Dokumen_Acuan = ObjIO.GetNoTransaksi(.Primary_Dokumen_Acuan)
                        ElseIf .Tipe_Dokumen_Acuan = ClsTransaksiMonthlyClosingItem.EnumTipeDokumenAcuan.StockBundlingAdjustment Then
                            rowDGV.Cells("ColJenisDocumentAdjustStr").Value = "Stock Bundling Item Before"
                        End If
                        rowDGV.Cells("ColDocumentNo").Value = .No_Dokumen_Acuan
                        rowDGV.Cells("ColDescAdjust").Value = .Description
                        rowDGV.Cells("ColQty").Value = ObjCD.DigitGrouping(.Qty_In_Adjust, ._Decimal)
                        rowDGV.Cells("ColUnitAdjust").Value = .Unit_Code
                        rowDGV.Cells("ColUnitPriceAdjust").Value = ObjCD.DigitGrouping(.Unit_Price, DefPrec)
                        rowDGV.Cells("ColRemarksAdjust").Value = .Remarks
                        rowDGV.Cells("ColPrimaryStockInAdjust").Value = .Primary_Dokumen_Acuan
                        rowDGV.Cells("ColPrimaryCodeNoAdjust").Value = .Primary_Code_No
                        rowDGV.Cells("ColPrimaryUnitAdjust").Value = .Primary_Units

                        'isikan juga baris2 dalam tabelnya
                        vFrm.rowA = vFrm.TblA.NewAdjustment_CorrectionRow
                        rowDGV = vFrm.dgvItemAdjust.Rows(vFrm.dgvItemAdjust.Rows.Count - 1)
                        With rowDGV
                            vFrm.rowA.Document = .Cells(vFrm.ColJenisDocumentAdjustStr.Index).Value
                            vFrm.rowA.Document_No = .Cells(vFrm.ColDocumentNo.Index).Value
                            vFrm.rowA.Code_No = .Cells(vFrm.ColCodeNoAdjust.Index).Value
                            vFrm.rowA.Description = .Cells(vFrm.ColDescAdjust.Index).Value
                            vFrm.rowA.Qty_In_Adjust = .Cells(vFrm.ColQty.Index).Value
                            vFrm.rowA.Units = .Cells(vFrm.ColUnitAdjust.Index).Value
                            vFrm.rowA.Unit_Price = CDec(.Cells(vFrm.ColUnitPriceAdjust.Index).Value)
                            vFrm.rowA.Primary_Code_No = .Cells(vFrm.ColPrimaryCodeNoAdjust.Index).Value
                            vFrm.rowA.Primary_Unit = .Cells(vFrm.ColPrimaryUnitAdjust.Index).Value
                            vFrm.rowA.Primary_Documen_Acuan = .Cells(vFrm.ColPrimaryStockInAdjust.Index).Value
                            vFrm.rowA.Jenis_Dokumen = .Cells(vFrm.ColJenisDokumen.Index).Value
                            vFrm.rowA.Remarks = .Cells(vFrm.ColRemarksAdjust.Index).Value

                            vFrm.rowA.Primary = -1
                            vFrm.rowA.Primary_Transaksi = -1
                            vFrm.rowA.Tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem

                            If vFrm.rowA.Jenis_Dokumen = ClsTransaksiMonthlyClosingItem.EnumTipeDokumenAcuan.StockInAdjust Then
                                vFrm.rowA.Tipe_Transaksi_Asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust
                            ElseIf vFrm.rowA.Jenis_Dokumen = ClsTransaksiMonthlyClosingItem.EnumTipeDokumenAcuan.StockTakingInAdjustment Then
                                vFrm.rowA.Tipe_Transaksi_Asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InStockTakeAdjust
                            ElseIf vFrm.rowA.Jenis_Dokumen = ClsTransaksiMonthlyClosingItem.EnumTipeDokumenAcuan.StockInSalesReturnPeriodeSama Then
                                vFrm.rowA.Tipe_Transaksi_Asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR
                            ElseIf vFrm.rowA.Jenis_Dokumen = ClsTransaksiMonthlyClosingItem.EnumTipeDokumenAcuan.StockBundlingAdjustment Then
                                vFrm.rowA.Tipe_Transaksi_Asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling
                            End If

                            vTBLDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem).ToString + _
                                " AND [tipe transaksi asal] = " + vFrm.rowA.Tipe_Transaksi_Asal.ToString
                            vRowDetail = vTBLDetail.DefaultView(0).Row

                            vFrm.rowA.Tanggal_Documen_Acuan = vRowDetail.tgl_transaksi_asal
                            vFrm.rowA.Primary_Adjustment_type = vRowDetail.primary_adjustment_type
                            vFrm.rowA.Primary_Gain_Account = vRowDetail.primary_gain_account
                            vFrm.rowA.Primary_Loss_Account = vRowDetail.primary_loss_account
                            If vFrm.rowA.Primary_Adjustment_type <> -1 Then
                                vFrm.rowA.Adjust_Type_Desc = vFrm.ObjStockInOut.GetDescriptionSettingAdjustment(vFrm.rowA.Primary_Adjustment_type)
                            Else
                                vFrm.rowA.Adjust_Type_Desc = ""
                            End If
                            vFrm.rowA.remarks_adjust = ""
                            vFrm.rowA.DeffPrec = row1._Decimal 'vFrm.ObjProduct.GetAngkaDecimal(vFrm.rowA.Primary_Code_No)
                        End With
                        vFrm.TblA.Rows.Add(vFrm.rowA)
                    End With
                Next
            End If

            DTAakun.Fill(TBLAkun, vPrimaryTransaksi)
            DTASetting.Fill(TBLSetting)
            For Each rowAkun In TBLAkun.Rows
                vDGVAkun.Rows.Add()
                rowDGV = vDGVAkun.Rows(vDGVAkun.Rows.Count - 1)

                With rowDGV
                    .Cells(i.ColJournal_Account.Index).Value = rowAkun.Account_Code

                    .Cells(i.ColJournal_Credit_OriginalCurrency.Index).Value = ObjCD.DigitGrouping(rowAkun.Credit_Original, DefPrec)
                    .Cells(i.ColJournal_Debet.Index).Value = ObjCD.DigitGrouping(rowAkun.Debet, DefPrec)
                    .Cells(i.ColJournal_Debet_OriginalCurr.Index).Value = ObjCD.DigitGrouping(rowAkun.Debet_Original, DefPrec)
                    .Cells(i.ColJournal_Kredit.Index).Value = ObjCD.DigitGrouping(rowAkun.Kredit, DefPrec)

                    .Cells(i.ColJournal_Dept.Index).Value = rowAkun.Dept_Code
                    .Cells(i.ColJournal_DescAccount.Index).Value = rowAkun.Account_Desc
                    .Cells(i.ColJournal_ExchangeRate.Index).Value = ObjCD.DigitGrouping(rowAkun.Ratio, Core_Module.DefPrecExchangeRate)
                    .Cells(i.ColJournal_Job.Index).Value = rowAkun.Job_Code
                    .Cells(i.ColJournal_OriginalCurrency.Index).Value = rowAkun.Currency_Code
                    .Cells(i.ColJournal_PrimaryAkun.Index).Value = rowAkun.Primary_Akun
                    .Cells(i.ColJournal_PrimaryCurrency.Index).Value = rowAkun.primary_currency
                    .Cells(i.ColJournal_PrimaryDept.Index).Value = rowAkun.primary_dept
                    .Cells(i.ColJournal_PrimaryJob.Index).Value = rowAkun.primary_job
                    .Cells(i.ColJournal_Remarks.Index).Value = rowAkun.Remarks
                    .Cells(i.ColJournal_TipeTransaksiAsalInt.Index).Value = rowAkun.tipe_transaksi_asal
                    .Cells(i.ColJournal_PrimaryAdjustmentType.Index).Value = rowAkun.primary_reference_adjustment
                    .Cells(i.ColJournal_TipeTransaksiAsalStr.Index).Value = Me.GetReference(rowAkun.primary_reference_adjustment, rowAkun.tipe_transaksi_asal, TBLSetting)
                End With
            Next

            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVItemAdjust.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVAkun.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells


            'terkait cost group
            Me.IsiDGV_Overhead(vPrimaryTransaksi, vDGVCostGroup, vTBLCostGroup_Account, vTBLCostGroup_Scope, _
            vTBLCostGroup_CodeNo, vFrm.ColCostGroup_AllocationStyle.Index, vFrm.ColCostGroup_GroupDesc.Index, _
            vFrm.ColCostGroup_Identifier.Index, vFrm.ColCostGroup_OverheadAmount.Index, vFrm.ColCostGroup_PrimaryRefUnit.Index, _
            vFrm.ColCostGroup_Remarks.Index, vFrm.ColCostGroup_TotalQty.Index, vFrm.ColCostGroup_RefUnit.Index, _
            vFrm.ObjCD, DefPrec, DTA.Connection.ConnectionString, vTBLCostGroup_CodeNo2)

            Me.IsiTBL_ProductQty(vPrimaryTransaksi, vTblProductQty)

            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
            ObjCD = Nothing
            ObjCur = Nothing
            DTASetting.Dispose()
            TBLSetting.Dispose()
            DTA1.Dispose()
            DTADetail.Dispose()
            DTABundling.Dispose()
            DTAakun.Dispose()
            TBL1.Dispose()
            DTASettingItem.Dispose()
        End Try
    End Function

    Public Sub IsiTBL_ProductQty(ByVal vPrimaryTransaksi As Double, ByVal vTblProductQty As Inventory_Module.DseLaporanStockDetail2.TblCetakDataTable)
        Dim Dta As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchase_closing_item_productqtyTableAdapter
        Dim Tbl As New DseTransaksiMonthlyClosing.tbltrans_purchase_closing_item_productqtyDataTable
        Dim Row As DseTransaksiMonthlyClosing.tbltrans_purchase_closing_item_productqtyRow
        Dim vRow As Inventory_Module.DseLaporanStockDetail2.TblCetakRow
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            Dta.Fill(Tbl, vPrimaryTransaksi)
            vTblProductQty.Rows.Clear()
            For Each Row In Tbl
                vRow = vTblProductQty.NewTblCetakRow
                vRow.Primary_Product = Row.Primary_Product
                vRow.Lot_No = Row.Lot_No
                vRow.Expiry_Date_Date = Row.Expiry_Date
                vRow.Primary_Location = Row.Primary_Storage_Area
                vRow.QtyDec = Row.Qty
                vTblProductQty.Rows.Add(vRow)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub IsiDGV_Overhead(ByVal vPrimaryTransaksi As Double, _
    ByVal vDGVCostGroup As DataGridView, ByVal vTBLCostGroup_Account As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_accountDataTable, _
    ByVal vTBLCostGroup_Scope As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_scopeDataTable, _
    ByVal vTBLCostGroup_CodeNo As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_code_noDataTable, _
    ByVal vIdx_ColCostGroup_AllocationStyle As Integer, ByVal vIdx_ColCostGroup_GroupDesc As Integer, ByVal vIdx_ColCostGroup_Identifier As Integer, _
    ByVal vIdx_ColCostGroup_OverheadAmount As Integer, ByVal vIdx_ColCostGroup_PrimaryRefUnit As Integer, ByVal vIdx_ColCostGroup_Remarks As Integer, _
    ByVal vIdx_ColCostGroup_TotalQty As Integer, ByVal vIdx_ColCostGroup_RefUnit As Integer, ByVal vObjCD As Core_Module.ClsCekData, _
    ByVal DefPrec As Integer, ByVal vConstring As String, _
    ByVal vTBLCostGroup_CodeNo2 As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_code_no2DataTable, Optional ByVal vIsikanCodeNo As Boolean = True)
        Dim DTA_CostGroup As New DseTransaksiMonthlyClosing_OverheadTableAdapters.tbltrans_purchase_closing_item_cost_group1TableAdapter
        Dim TBL_CostGroup As New DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group1DataTable
        Dim row_CostGroup As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group1Row
        Dim CboDgv As DataGridViewComboBoxCell
        Dim DTA_CostScope As New DseTransaksiMonthlyClosing_OverheadTableAdapters.tbltrans_purchase_closing_item_cost_group_scopeTableAdapter
        Dim DTA_CostAcc As New DseTransaksiMonthlyClosing_OverheadTableAdapters.tbltrans_purchase_closing_item_cost_group_accountTableAdapter
        Dim DTA_CostCodeNo As New DseTransaksiMonthlyClosing_OverheadTableAdapters.tbltrans_purchase_closing_item_cost_group_code_noTableAdapter
        Dim DTA_CostCodeNo2 As New DseTransaksiMonthlyClosing_OverheadTableAdapters.tbltrans_purchase_closing_item_cost_group_code_no2TableAdapter
        Dim Dummy As Boolean = False
        Dim rowDGV As DataGridViewRow

        Try
            DTA_CostGroup.Connection.ConnectionString = vConstring
            DTA_CostScope.Connection.ConnectionString = vConstring
            DTA_CostAcc.Connection.ConnectionString = vConstring
            DTA_CostCodeNo.Connection.ConnectionString = vConstring
            DTA_CostCodeNo2.Connection.ConnectionString = vConstring

            Dummy = vDGVCostGroup.AllowUserToAddRows
            vDGVCostGroup.AllowUserToAddRows = False
            DTA_CostGroup.Fill(TBL_CostGroup, vPrimaryTransaksi)
            vDGVCostGroup.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVCostGroup.Rows.Clear()
            For Each row_CostGroup In TBL_CostGroup.Rows
                vDGVCostGroup.Rows.Add()
                rowDGV = vDGVCostGroup.Rows(vDGVCostGroup.Rows.Count - 1)
                With row_CostGroup
                    CboDgv = rowDGV.Cells(vIdx_ColCostGroup_AllocationStyle)
                    CboDgv.Value = CboDgv.Items.Item(.allocation_style)
                    rowDGV.Cells(vIdx_ColCostGroup_GroupDesc).Value = .desc
                    rowDGV.Cells(vIdx_ColCostGroup_Identifier).Value = .identifier
                    rowDGV.Cells(vIdx_ColCostGroup_OverheadAmount).Value = vObjCD.DigitGrouping(.overhead_amount, DefPrec)
                    rowDGV.Cells(vIdx_ColCostGroup_PrimaryRefUnit).Value = .primary_ref_unit
                    rowDGV.Cells(vIdx_ColCostGroup_Remarks).Value = .remarks
                    rowDGV.Cells(vIdx_ColCostGroup_TotalQty).Value = .total_qty
                    rowDGV.Cells(vIdx_ColCostGroup_RefUnit).Value = .Unit_Code
                End With
            Next

            vDGVCostGroup.AllowUserToAddRows = Dummy
            vDGVCostGroup.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vTBLCostGroup_Account.Rows.Clear()
            vTBLCostGroup_CodeNo.Rows.Clear()
            vTBLCostGroup_Scope.Rows.Clear()
            If vIsikanCodeNo = True Then
                DTA_CostCodeNo.FillBy_PrimaryTransaksi(vTBLCostGroup_CodeNo, vPrimaryTransaksi)
                DTA_CostCodeNo2.FillBy_PrimaryTransaksi(vTBLCostGroup_CodeNo2, vPrimaryTransaksi)
            End If
            DTA_CostAcc.FillBy_PrimaryTransaksi(vTBLCostGroup_Account, vPrimaryTransaksi)
            DTA_CostScope.FillBy_Primary(vTBLCostGroup_Scope, vPrimaryTransaksi)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_CostGroup.Dispose()
            TBL_CostGroup.Dispose()
            DTA_CostScope.Dispose()
            DTA_CostAcc.Dispose()
            DTA_CostCodeNo.Dispose()
            DTA_CostCodeNo2.Dispose()
        End Try
    End Sub

    Public Function Simpan(ByVal vdgvItem As DataGridView, ByVal vDGVItemAdjust As DataGridView, _
    ByVal vPrimaryTransaksi As Double, ByVal vFRM As FrmTransaksiMonthlyClosing2, _
    ByVal vTBLDetail As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailDataTable, _
    ByVal vTBLBundling As DseTransaksiMonthlyClosing2.TblBundling1DataTable, _
    ByVal vDGVItemAkun As DataGridView, _
    ByVal vTBLSetting As DseTransaksiMonthlyClosing2.tbltrans_purchase_closing_item_tipe_adjustmentDataTable, _
    ByVal vTBLPool As DseMonthlyClosingNew.tblItemDetailPoolDataTable, _
    ByVal vDGVCostGroup As DataGridView, ByVal vTBLCostGroup_Account As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_accountDataTable, _
    ByVal vTBLCostGroup_Scope As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_scopeDataTable, _
    ByVal vTBLCostGroup_CodeNo As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_code_noDataTable, _
    ByVal vTBLCostGroup_CodeNo2 As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_code_no2DataTable, _
    ByVal vTblProductQty As Inventory_Module.DseLaporanStockDetail2.TblCetakDataTable) As Boolean
        Dim DTA As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchase_closing_itemTableAdapter
        Dim DTA1 As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchase_closing_item_inadjustmentTableAdapter
        Dim DTA2 As New DseTransaksiMonthlyClosing2TableAdapters.tbltrans_purchase_closing_item_akunTableAdapter
        Dim DTA3 As New DseTransaksiMonthlyClosing2TableAdapters.tbltrans_purchase_closing_item_tipe_adjustmentTableAdapter
        Dim rowDGV As DataGridViewRow
        Dim rowSetting As DseTransaksiMonthlyClosing2.tbltrans_purchase_closing_item_tipe_adjustmentRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA2.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA3.Connection.ConnectionString = DTA.Connection.ConnectionString

            'simpan item
            Me.Simpan_UsingScript(vdgvItem, vPrimaryTransaksi, vFRM)
            'simpan item adjustment
            Me.Simpan_UsingScript_ItemAdjustment(vDGVItemAdjust, vPrimaryTransaksi, vFRM)
            'simpan item detail
            Me.SimpanUsingScript_ItemDetail(vTBLDetail, vPrimaryTransaksi)
            'simpan item bundling
            Me.SimpanUsingScript_ItemBundling(vTBLBundling, vPrimaryTransaksi)
            'simpan item bundling before
            'Me.SimpanUsingScript_ItemBundlingBefore(vTBLBefore, vPrimaryTransaksi)
            'simpan item bundling after
            'Me.SimpanUsingScript_ItemBundlingAfter(vTBLAfter, vPrimaryTransaksi)
            'simpan item pool
            Me.SimpanUsingScript_ItemPool(vTBLPool, vPrimaryTransaksi)
            'simpan item akun
            For Each rowDGV In vDGVItemAkun.Rows
                With rowDGV
                    DTA2.Insert(vPrimaryTransaksi, .Cells(vFRM.ColJournal_PrimaryAkun.Index).Value, .Cells(vFRM.ColJournal_Debet.Index).Value, _
                    .Cells(vFRM.ColJournal_Kredit.Index).Value, .Cells(vFRM.ColJournal_Remarks.Index).Value, _
                    .Cells(vFRM.ColJournal_Debet_OriginalCurr.Index).Value, .Cells(vFRM.ColJournal_Credit_OriginalCurrency.Index).Value, _
                    .Cells(vFRM.ColJournal_ExchangeRate.Index).Value, .Cells(vFRM.ColJournal_PrimaryDept.Index).Value, _
                    .Cells(vFRM.ColJournal_PrimaryJob.Index).Value, .Cells(vFRM.ColJournal_PrimaryAdjustmentType.Index).Value, _
                    .Cells(vFRM.ColJournal_TipeTransaksiAsalInt.Index).Value)
                End With
            Next
            'simpan item setting
            For Each rowSetting In vTBLSetting.Rows
                DTA3.InsertQuery(rowSetting.PrimaryGain, rowSetting.PrimaryLoss, vPrimaryTransaksi, rowSetting.Primary_Adjustment)
                rowSetting.Primary_Transaksi = vPrimaryTransaksi
            Next

            'cost group
            Me.Simpan_UsingScript_ItemCostGroup(vDGVCostGroup, vPrimaryTransaksi, vFRM)
            Me.SimpanUsingScript_ItemCostGroup_Account(vTBLCostGroup_Account, vPrimaryTransaksi)
            Me.SimpanUsingScript_ItemCostGroup_Scope(vTBLCostGroup_Scope, vPrimaryTransaksi)
            Me.SimpanUsingScript_ItemCostGroup_CodeNo(vTBLCostGroup_CodeNo, vPrimaryTransaksi)
            Me.SimpanUsingScript_ItemCostGroup_CodeNo2(vTBLCostGroup_CodeNo2, vPrimaryTransaksi)
            Me.SimpanUsingScript_ItemProductQty(vTblProductQty, vPrimaryTransaksi)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
            DTA2.Dispose()
            DTA3.Dispose()
        End Try
    End Function

    Private Function SimpanUsingScript_ItemProductQty(ByVal vTBL As Inventory_Module.DseLaporanStockDetail2.TblCetakDataTable, _
            ByVal vPrimaryTransaksi As Double) As Boolean
        Dim conStr As String = My.Settings.core_dbConnectionString
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 50
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < vTBL.Rows.Count
                sql = Me.GetInsertString_ItemProductQty(StartRowIndex, BatchQty, vTBL, vPrimaryTransaksi)
                script.Query = sql
                'execute script
                script.Execute()
                StartRowIndex += BatchQty
            End While

            Conn.Close()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Private Function GetInsertString_ItemProductQty(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, ByVal vTBLItemDetail As Inventory_Module.DseLaporanStockDetail2.TblCetakDataTable, _
            ByVal vPrimaryTransaksi As Double) As String
        Dim retval As String = ""
        Dim Delimiter As String = ","
        Dim row As Inventory_Module.DseLaporanStockDetail2.TblCetakRow

        Try
            retval = "INSERT INTO `tbltrans_purchase_closing_item_productqty` (`Primary Transaksi`,`Primary Product`,`Qty`,`Lot No`,`Expiry Date`,`Primary Storage Area`) VALUES "

            For i As Double = 0 To BatchQty - 1 Step 1
                If vTBLItemDetail.Rows.Count > vStartRowIndex + i Then
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTBLItemDetail.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    row = vTBLItemDetail.Rows(vStartRowIndex + i)
                    If row.Lot_No <> "" Then
                        If row.Lot_No = "-" Then
                            row.Lot_No = ""
                        End If
                        If row.Lot_No.Contains("\") = True Then
                            row.Lot_No = row.Lot_No.Replace("\", "\\")
                        End If
                        If row.Lot_No.Contains(vbCrLf) = True Then
                            row.Lot_No = row.Lot_No.Replace(vbCrLf, "\r\n")
                        End If
                        If row.Lot_No.Contains("'") = True Then
                            row.Lot_No = row.Lot_No.Replace("'", "''")
                        End If
                    End If
                    retval = retval & vbCrLf & "(" & vPrimaryTransaksi.ToString & "," _
                    & CDbl(row.Primary_Product).ToString & "," _
                    & CDec(row.QtyDec).ToString & "," _
                    & "'" & row.Lot_No & "'," _
                    & "'" & Format(row.Expiry_Date_Date, "yyyy-MM-dd") & "'" & "," _
                    & CDbl(row.Primary_Location).ToString & ")" & Delimiter
                Else
                    Exit For
                End If
            Next
            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function SimpanUsingScript_ItemCostGroup_Account(ByVal vTBL As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_accountDataTable, _
ByVal vPrimaryTransaksi As Double) As Boolean
        Dim conStr As String = My.Settings.core_dbConnectionString
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 50
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < vTBL.Rows.Count
                sql = Me.GetInsertString_ItemCostGroup_Account(StartRowIndex, BatchQty, vTBL, vPrimaryTransaksi)
                script.Query = sql
                'execute script
                script.Execute()
                StartRowIndex += BatchQty
            End While

            Conn.Close()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Private Function SimpanUsingScript_ItemCostGroup_Scope(ByVal vTBL As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_scopeDataTable, _
    ByVal vPrimaryTransaksi As Double) As Boolean
        Dim conStr As String = My.Settings.core_dbConnectionString
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 50
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < vTBL.Rows.Count
                sql = Me.GetInsertString_ItemCostGroup_Scope(StartRowIndex, BatchQty, vTBL, vPrimaryTransaksi)
                script.Query = sql
                'execute script
                script.Execute()
                StartRowIndex += BatchQty
            End While

            Conn.Close()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Private Function GetInsertString_ItemCostGroup_Scope(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, ByVal vTBLItemDetail As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_scopeDataTable, _
    ByVal vPrimaryTransaksi As Double) As String
        Dim retval As String = ""
        Dim Delimiter As String = ","
        Dim row As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_scopeRow

        Try
            retval = "INSERT INTO `tbltrans_purchase_closing_item_cost_group_scope` (`primary transaksi`,`primary process`,`identifier`,`remarks`) VALUES "

            For i As Double = 0 To BatchQty - 1 Step 1
                If vTBLItemDetail.Rows.Count > vStartRowIndex + i Then
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTBLItemDetail.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    row = vTBLItemDetail.Rows(vStartRowIndex + i)
                    retval = retval + vbCrLf + "(" + vPrimaryTransaksi.ToString + "," _
                    + row.primary_process.ToString + "," _
                    + "'" + row.identifier + "'," _
                    + "'" + row.remarks.ToString + "'" _
                    + ")" + Delimiter
                Else
                    Exit For
                End If
            Next
            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function SimpanUsingScript_ItemCostGroup_CodeNo(ByVal vTBL As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_code_noDataTable, _
    ByVal vPrimaryTransaksi As Double) As Boolean
        Dim conStr As String = My.Settings.core_dbConnectionString
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 50
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < vTBL.Rows.Count
                sql = Me.GetInsertString_ItemCostGroup_CodeNo(StartRowIndex, BatchQty, vTBL, vPrimaryTransaksi)
                script.Query = sql
                'execute script
                script.Execute()
                StartRowIndex += BatchQty
            End While

            Conn.Close()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Private Function GetInsertString_ItemCostGroup_CodeNo(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, ByVal vTBLItemDetail As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_code_noDataTable, _
    ByVal vPrimaryTransaksi As Double) As String
        Dim retval As String = ""
        Dim Delimiter As String = ","
        Dim row As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_code_noRow

        Try
            retval = "INSERT INTO `tbltrans_purchase_closing_item_cost_group_code_no` (`primary transaksi`,`primary code no`,`qty ref`,`primary unit`,`identifier`,`remark`,`allocated amount`,`primary bundling`,`decimal qty`,`primary process`,`cost factor`) VALUES "

            For i As Double = 0 To BatchQty - 1 Step 1
                If vTBLItemDetail.Rows.Count > vStartRowIndex + i Then
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTBLItemDetail.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    row = vTBLItemDetail.Rows(vStartRowIndex + i)
                    retval = retval + vbCrLf + "(" + vPrimaryTransaksi.ToString + "," _
                    + row.primary_code_no.ToString + "," _
                    + row.qty_ref.ToString + "," _
                    + row.primary_unit.ToString + "," _
                    + "'" + row.identifier + "'," _
                    + "'" + row.remark.ToString + "'," _
                    + row.allocated_amount.ToString + "," _
                    + row.primary_bundling.ToString + "," _
                    + row.decimal_qty.ToString + "," _
                    + row.primary_process.ToString + "," _
                    + row.cost_factor.ToString + "" _
                    + ")" + Delimiter
                Else
                    Exit For
                End If
            Next
            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function SimpanUsingScript_ItemCostGroup_CodeNo2(ByVal vTBL As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_code_no2DataTable, _
   ByVal vPrimaryTransaksi As Double) As Boolean
        Dim conStr As String = My.Settings.core_dbConnectionString
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 50
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < vTBL.Rows.Count
                sql = Me.GetInsertString_ItemCostGroup_CodeNo2(StartRowIndex, BatchQty, vTBL, vPrimaryTransaksi)
                script.Query = sql
                'execute script
                script.Execute()
                StartRowIndex += BatchQty
            End While

            Conn.Close()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Private Function GetInsertString_ItemCostGroup_CodeNo2(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, _
    ByVal vTBLItemDetail As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_code_no2DataTable, _
    ByVal vPrimaryTransaksi As Double) As String
        Dim retval As String = ""
        Dim Delimiter As String = ","
        Dim row As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_code_no2Row

        Try
            retval = "INSERT INTO `tbltrans_purchase_closing_item_cost_group_code_no2` (`primary transaksi`,`allocated amount`,`primary account`,`primary process`) VALUES "

            For i As Double = 0 To BatchQty - 1 Step 1
                If vTBLItemDetail.Rows.Count > vStartRowIndex + i Then
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTBLItemDetail.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    row = vTBLItemDetail.Rows(vStartRowIndex + i)
                    retval = retval + vbCrLf + "(" + vPrimaryTransaksi.ToString + "," _
                    + CDec(row.allocated_amount).ToString + "," _
                    + CDbl(row.primary_account).ToString + "," _
                    + CDbl(row.primary_process).ToString + "" _
                    + ")" + Delimiter
                Else
                    Exit For
                End If
            Next
            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function GetInsertString_ItemCostGroup_Account(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, ByVal vTBLItemDetail As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_accountDataTable, _
    ByVal vPrimaryTransaksi As Double) As String
        Dim retval As String = ""
        Dim Delimiter As String = ","
        Dim row As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_accountRow

        Try
            retval = "INSERT INTO `tbltrans_purchase_closing_item_cost_group_account` (`primary transaksi`,`identifier`,`balance`,`perc allocated`,`allocated amount`,`remarks`,`primary akun`,`ignore negative`) VALUES "

            For i As Double = 0 To BatchQty - 1 Step 1
                If vTBLItemDetail.Rows.Count > vStartRowIndex + i Then
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTBLItemDetail.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    row = vTBLItemDetail.Rows(vStartRowIndex + i)
                    retval = retval + vbCrLf + "(" + vPrimaryTransaksi.ToString + ",'" + row.identifier + "'," _
                    + row.balance.ToString + "," _
                    + row.perc_allocated.ToString + "," _
                    + row.allocated_amount.ToString + "," _
                    + "'" + row.remarks.ToString + "'," _
                    + row.primary_akun.ToString + "," _
                    + row.ignore_negative.ToString + "" _
                    + ")" + Delimiter
                Else
                    Exit For
                End If
            Next
            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function Simpan_UsingScript_ItemCostGroup(ByVal vdgvItem_CostGroup As DataGridView, _
  ByVal vPrimaryTransaksi As Double, ByVal i As FrmTransaksiMonthlyClosing2) As Boolean
        Dim conStr As String = My.Settings.core_dbConnectionString
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 50
        Dim StartRowIndex As Double = 0
        Dim Dummy As Boolean = False

        Try
            'open connection
            Conn.Open()
            Dummy = vdgvItem_CostGroup.AllowUserToAddRows
            vdgvItem_CostGroup.AllowUserToAddRows = False

            While StartRowIndex < vdgvItem_CostGroup.Rows.Count
                sql = Me.GetInsertString_ItemCostGroup(StartRowIndex, BatchQty, vdgvItem_CostGroup, vPrimaryTransaksi, i)
                script.Query = sql
                'execute script
                script.Execute()
                StartRowIndex += BatchQty
            End While

            Conn.Close()
            vdgvItem_CostGroup.AllowUserToAddRows = Dummy

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Private Function GetInsertString_ItemCostGroup(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, ByVal vdgvItem_CostGroup As DataGridView, _
    ByVal vPrimaryTransaksi As Double, ByVal ijk As FrmTransaksiMonthlyClosing2) As String
        Dim retval As String = ""
        Dim row As DataGridViewRow
        Dim Delimiter As String = ","
        Dim CboDgv As DataGridViewComboBoxCell
        Dim AllocationStyle_Int As Integer = 0

        Try
            retval = "INSERT INTO `tbltrans_purchase_closing_item_cost_group` (`primary transaksi`,`desc`,`identifier`,`remarks`,`allocation style`,`primary ref unit`,`total qty`,`overhead amount`) VALUES "
            For i As Double = 0 To BatchQty - 1 Step 1
                If vdgvItem_CostGroup.Rows.Count > vStartRowIndex + i Then
                    row = vdgvItem_CostGroup.Rows(vStartRowIndex + i)
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vdgvItem_CostGroup.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    CboDgv = row.Cells(ijk.ColCostGroup_AllocationStyle.Index)
                    For kl As Integer = 0 To CboDgv.Items.Count - 1
                        If CboDgv.Items.Item(kl) = CboDgv.Value Then
                            AllocationStyle_Int = kl
                            Exit For
                        End If
                    Next

                    If IsNothing(row.Cells(ijk.ColCostGroup_Remarks.Index).Value) = True Then
                        row.Cells(ijk.ColCostGroup_Remarks.Index).Value = ""
                    End If
                    retval = retval + vbCrLf + "(" + vPrimaryTransaksi.ToString + ",'" + row.Cells(ijk.ColCostGroup_GroupDesc.Index).Value.ToString + "'," _
                    + "'" + row.Cells(ijk.ColCostGroup_Identifier.Index).Value.ToString + "'," _
                    + "'" + row.Cells(ijk.ColCostGroup_Remarks.Index).Value.ToString + "'," _
                    + AllocationStyle_Int.ToString + "," _
                    + CDbl(row.Cells(ijk.ColCostGroup_PrimaryRefUnit.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.ColCostGroup_TotalQty.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.ColCostGroup_OverheadAmount.Index).Value).ToString + "" _
                    + ")" + Delimiter
                Else
                    Exit For
                End If
            Next
            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function SimpanUsingScript_ItemDetail(ByVal vTBL As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailDataTable, _
    ByVal vPrimaryTransaksi As Double) As Boolean
        Dim conStr As String = My.Settings.core_dbConnectionString
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 50
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < vTBL.Rows.Count
                sql = Me.GetInsertString_ItemDetail(StartRowIndex, BatchQty, vTBL, vPrimaryTransaksi)
                script.Query = sql
                'execute script
                script.Execute()
                StartRowIndex += BatchQty
            End While

            Conn.Close()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Private Function SimpanUsingScript_ItemPool(ByVal vTBL As DseMonthlyClosingNew.tblItemDetailPoolDataTable, _
   ByVal vPrimaryTransaksi As Double) As Boolean
        Dim conStr As String = My.Settings.core_dbConnectionString
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 50
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < vTBL.Rows.Count
                sql = Me.GetInsertString_ItemPool(StartRowIndex, BatchQty, vTBL, vPrimaryTransaksi)
                script.Query = sql
                'execute script
                script.Execute()
                StartRowIndex += BatchQty
            End While

            Conn.Close()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Private Function SimpanUsingScript_ItemBundling(ByVal vTBL As DseTransaksiMonthlyClosing2.TblBundling1DataTable, _
    ByVal vPrimaryTransaksi As Double) As Boolean
        Dim conStr As String = My.Settings.core_dbConnectionString
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 50
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < vTBL.Rows.Count
                sql = Me.GetInsertString_ItemBundling(StartRowIndex, BatchQty, vTBL, vPrimaryTransaksi)
                script.Query = sql
                'execute script
                script.Execute()
                StartRowIndex += BatchQty
            End While

            Conn.Close()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Private Function Simpan_UsingScript(ByVal vdgvItem As DataGridView, _
    ByVal vPrimaryTransaksi As Double, ByVal i As FrmTransaksiMonthlyClosing2) As Boolean
        Dim conStr As String = My.Settings.core_dbConnectionString
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 50
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < vdgvItem.Rows.Count
                sql = Me.GetInsertString(StartRowIndex, BatchQty, vdgvItem, vPrimaryTransaksi, i)
                script.Query = sql
                'execute script
                script.Execute()
                StartRowIndex += BatchQty
            End While

            Conn.Close()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Private Function GetInsertString(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, ByVal vdgvItem As DataGridView, _
    ByVal vPrimaryTransaksi As Double, ByVal ijk As FrmTransaksiMonthlyClosing2) As String
        Dim retval As String = ""
        Dim row As DataGridViewRow
        Dim Delimiter As String = ","

        Try
            retval = "INSERT INTO `tbltrans_purchase_closing_item` (`Primary Transaksi`, `Primary Product`, " & _
            "`Primary Unit`, `Begin Qty`, `Begin Amount`, `In Qty`, `In Amount`, `Out Qty`, " & _
            "`Out Amount`, `Ending Qty`, `Ending Amount`, `Remarks`, `Primary Inventory Account`, " & _
            "`Primary Adjustment Account`, `Primary Stock Taking Account`, `Primary COGS Account`, " & _
            "`Primary Sales Account`, `Primary Sales Return Account`, `Decimal`) VALUES "
            'retval = "INSERT INTO `tbltrans_purchase_closing_item` (`Primary Transaksi`, `Primary Product`, `Primary Unit`, `Begin Qty`, `Begin Amount`, `In Qty`, `In Amount`, `Out Qty`, `Out Amount`, `Ending Qty`, `Ending Amount`, `Remarks`) VALUES "
            For i As Double = 0 To BatchQty - 1 Step 1
                If vdgvItem.Rows.Count > vStartRowIndex + i Then
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vdgvItem.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    row = vdgvItem.Rows(vStartRowIndex + i)
                    retval = retval + vbCrLf + "(" + vPrimaryTransaksi.ToString + "," + row.Cells("ColPrimaryCodeNo").Value.ToString + "," _
                    + row.Cells("ColPrimaryUnits").Value.ToString + "," _
                    + CDec(row.Cells("ColBeginQty").Value).ToString + "," _
                    + CDec(row.Cells("ColBeginAmount").Value).ToString + "," _
                    + CDec(row.Cells("ColInQty").Value).ToString + "," _
                    + CDec(row.Cells("ColInAmount").Value).ToString + "," _
                    + CDec(row.Cells("ColOutQty").Value).ToString + "," _
                    + CDec(row.Cells("ColOutAmount").Value).ToString + "," _
                    + CDec(row.Cells("ColEndQty").Value).ToString + "," _
                    + CDec(row.Cells("ColEndAmount").Value).ToString + "," _
                    + "'" + row.Cells("ColRemarks").Value.ToString + "'," _
                    + CDec(row.Cells(ijk.ColItem_PrimaryInventoryAccount.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.ColItem_PrimaryAdjustmentAccount.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.ColItem_PrimaryStockTakingAccount.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.ColItem_PrimaryCOGSAccount.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.ColItem_PrimarySalesAccount.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.ColItem_PrimarySalesReturnAccount.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.ColItem_Decimal.Index).Value).ToString + "" _
                    + ")" + Delimiter
                Else
                    Exit For
                End If
            Next
            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function GetInsertString_ItemDetail(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, ByVal vTBLItemDetail As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailDataTable, _
    ByVal vPrimaryTransaksi As Double) As String
        Dim retval As String = ""
        Dim Delimiter As String = ","
        Dim row As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailRow

        Try
            retval = "INSERT INTO `tbltrans_purchase_closing_item_detail` " & _
            "(`primary transaksi`, `primary transaksi asal`, `tipe transaksi asal`, `tgl transaksi asal`, `primary code no`, `tipe item detail`, `unit price`, `amount`, `qty`, `return price`, `primary gain account`,`primary loss account`,`primary adjustment type`,`remarks`, `Overhead`, `Cost Perct Bundling`) VALUES "
            'retval = "INSERT INTO `tbltrans_purchase_closing_item` (`Primary Transaksi`, `Primary Product`, `Primary Unit`, `Begin Qty`, `Begin Amount`, `In Qty`, `In Amount`, `Out Qty`, `Out Amount`, `Ending Qty`, `Ending Amount`, `Remarks`) VALUES "
            For i As Double = 0 To BatchQty - 1 Step 1
                If vTBLItemDetail.Rows.Count > vStartRowIndex + i Then
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTBLItemDetail.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    row = vTBLItemDetail.Rows(vStartRowIndex + i)
                    retval = retval + vbCrLf + "(" + vPrimaryTransaksi.ToString + "," + row.primary_transaksi_asal.ToString + "," _
                    + row.tipe_transaksi_asal.ToString + "," _
                    + "'" + Format(row.tgl_transaksi_asal, "yyyy-MM-dd") + "'" + "," _
                    + row.primary_code_no.ToString + "," _
                    + row.tipe_item_detail.ToString + "," _
                    + row.unit_price.ToString + "," _
                    + row.amount.ToString + "," _
                    + row.qty.ToString + "," _
                    + row.return_price.ToString + "," _
                    + row.primary_gain_account.ToString + "," _
                    + row.primary_loss_account.ToString + "," _
                    + row.primary_adjustment_type.ToString + "," _
                    + "'" + row.remarks.ToString + "'," _
                    + row.Overhead.ToString + "," _
                    + row.Cost_Perct_Bundling.ToString + _
                    ")" + Delimiter
                Else
                    Exit For
                End If
            Next
            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function GetInsertString_ItemPool(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, _
    ByVal vTBLItemPool As DseMonthlyClosingNew.tblItemDetailPoolDataTable, _
    ByVal vPrimaryTransaksi As Double) As String
        Dim retval As String = ""
        Dim Delimiter As String = ","
        Dim row As DseMonthlyClosingNew.tblItemDetailPoolRow

        Try
            retval = "INSERT INTO `tbltrans_purchase_closing_item_detail_pool` " & _
            "(`primary transaksi`, `primary transaksi asal`, " & _
            "`tipe transaksi asal`, `tgl transasi asal`, " & _
            "`primary code no`, `tipe item detail`, " & _
            "`unit price`, `primary gain account`, " & _
            "`primary loss account`, `primary adjustment type`, " & _
            "`remarks`, `overhead`, " & _
            "`Cost Perct Bundling`) VALUES "

            For i As Double = 0 To BatchQty - 1 Step 1
                If vTBLItemPool.Rows.Count > vStartRowIndex + i Then
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTBLItemPool.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    row = vTBLItemPool.Rows(vStartRowIndex + i)
                    retval = retval + vbCrLf + "(" + vPrimaryTransaksi.ToString + "," + row.Primary_Transaksi_Asal.ToString + "," _
                    + row.tipe_transaksi_asal.ToString + "," _
                    + "'" + Format(row.tgl_transaksi_asal, "yyyy-MM-dd") + "'" + "," _
                    + row.primary_code_no.ToString + "," _
                    + row.tipe_item_detail.ToString + "," _
                    + row.unit_price.ToString + "," _
                    + row.amount.ToString + "," _
                    + row.qty.ToString + "," _
                    + row.return_price.ToString + "," _
                    + row.primary_gain_account.ToString + "," _
                    + row.primary_loss_account.ToString + "," _
                    + row.primary_adjustment_type.ToString + "," _
                    + "'" + row.remarks.ToString + "'," _
                    + row.Overhead.ToString + "," _
                    + row.Cost_Perct_Bundling.ToString + _
                    +")" + Delimiter


                Else
                    Exit For
                End If
            Next
            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function GetInsertString_ItemBundling(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, _
    ByVal vTBLItemDetail As DseTransaksiMonthlyClosing2.TblBundling1DataTable, _
    ByVal vPrimaryTransaksi As Double) As String
        Dim retval As String = ""
        Dim Delimiter As String = ","
        Dim row As DseTransaksiMonthlyClosing2.TblBundling1Row

        Try
            retval = "INSERT INTO `tbltrans_purchase_closing_item_bundling` (`primary transaksi`, `primary code no before`, `primary code no after`) VALUES "

            For i As Double = 0 To BatchQty - 1 Step 1
                If vTBLItemDetail.Rows.Count > vStartRowIndex + i Then
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTBLItemDetail.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    row = vTBLItemDetail.Rows(vStartRowIndex + i)
                    retval = retval + vbCrLf + "(" + vPrimaryTransaksi.ToString + "," + row.Primary_Product_Before.ToString + "," _
                    + row.Primary_Product_After.ToString + "" _
                    + ")" + Delimiter
                Else
                    Exit For
                End If
            Next
            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function Simpan_UsingScript_ItemAdjustment(ByVal vdgvItem As DataGridView, _
    ByVal vPrimaryTransaksi As Double, ByVal vFRM As FrmTransaksiMonthlyClosing2) As Boolean
        Dim conStr As String = My.Settings.core_dbConnectionString
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 50
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < vdgvItem.Rows.Count
                sql = Me.GetInsertString_ItemAdjustment(StartRowIndex, BatchQty, vdgvItem, vPrimaryTransaksi, vFRM)
                script.Query = sql
                'execute script
                script.Execute()
                StartRowIndex += BatchQty
            End While

            Conn.Close()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message + " - Start Row Index: " + StartRowIndex.ToString, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Private Function GetInsertString_ItemAdjustment(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, ByVal vdgvItem As DataGridView, _
    ByVal vPrimaryTransaksi As Double, ByVal vFrm As FrmTransaksiMonthlyClosing2) As String
        Dim retval As String = ""
        Dim row As DseMonthlyClosingNew.Adjustment_CorrectionRow
        Dim Delimiter As String = ","

        Try
            retval = "INSERT INTO `tbltrans_purchase_closing_item_inadjustment` " & _
            "(`Primary Transaksi`, `Primary Code No`, " & _
            "`Primary Units`, `Primary Dokumen Acuan`, " & _
            "`Tipe Dokumen Acuan`, `Qty In Adjust`, " & _
            "`Unit Price`, `Remarks`, " & _
            "`Tipe item detail`, `Tipe transaksi asal`, " & _
            "`Primary Adjustment Type`, `Primary Gain Account`, " & _
            "`Primary Loss Account`, `Remarks Adjust`, `Deff prec`) VALUES "

            For i As Double = 0 To BatchQty - 1 Step 1
                If vFrm.TblA.Rows.Count > vStartRowIndex + i Then
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vFrm.TblA.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    row = vFrm.TblA.Rows(vStartRowIndex + i)
                    row.Primary_Transaksi = vPrimaryTransaksi
                    retval = retval + vbCrLf + "(" + row.Primary_Transaksi.ToString + "," _
                    + row.Primary_Code_No.ToString + "," _
                    + row.Primary_Unit.ToString + "," _
                    + row.Primary_Documen_Acuan.ToString + "," _
                    + row.Jenis_Dokumen.ToString + "," _
                    + row.Qty_In_Adjust.ToString + "," _
                    + row.Unit_Price.ToString + "," _
                    + "'" + ObjCD.FormatSpecialChar_MySQL(row.Remarks.ToString) + "'," _
                    + row.Tipe_item_detail.ToString + "," _
                    + row.Tipe_Transaksi_Asal.ToString + "," _
                    + row.Primary_Adjustment_type.ToString + "," _
                    + row.Primary_Gain_Account.ToString + "," _
                    + row.Primary_Loss_Account.ToString + "," _
                    + "'" + row.remarks_adjust.ToString + "'," _
                    + row.DeffPrec.ToString + _
                    ")" + Delimiter
                Else
                    Exit For
                End If
            Next
            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function Hapus(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchase_closing_itemTableAdapter
        Dim DTA1 As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchase_closing_item_inadjustmentTableAdapter
        Dim DTA2 As New DseTransaksiMonthlyClosing2TableAdapters.TblBundling1TableAdapter
        Dim DTA3 As New DseTransaksiMonthlyClosing2TableAdapters.TblBundlingBeforeTableAdapter
        Dim DTA4 As New DseTransaksiMonthlyClosing2TableAdapters.TblBundlingAfterTableAdapter
        Dim DTA5 As New DseTransaksiMonthlyClosing2TableAdapters.tbltrans_purchase_closing_item_akunTableAdapter
        Dim DTA6 As New DseMonthlyClosingNewTableAdapters.tbltrans_purchase_closing_item_detailTableAdapter
        Dim DTA7 As New DseTransaksiMonthlyClosing2TableAdapters.tbltrans_purchase_closing_item_tipe_adjustmentTableAdapter
        Dim DTA81 As New DseMonthlyClosingNewTableAdapters.tbltrans_purchase_closing_item_detail_poolTableAdapter
        Dim DTA8 As New DseTransaksiMonthlyClosing_OverheadTableAdapters.tbltrans_purchase_closing_item_cost_group_accountTableAdapter
        Dim DTA9 As New DseTransaksiMonthlyClosing_OverheadTableAdapters.tbltrans_purchase_closing_item_cost_group_code_noTableAdapter
        Dim DTA10 As New DseTransaksiMonthlyClosing_OverheadTableAdapters.tbltrans_purchase_closing_item_cost_group_scopeTableAdapter
        Dim DTA11 As New DseTransaksiMonthlyClosing_OverheadTableAdapters.tbltrans_purchase_closing_item_cost_groupTableAdapter
        Dim DTA12 As New DseTransaksiMonthlyClosing_OverheadTableAdapters.tbltrans_purchase_closing_item_cost_group_code_no2TableAdapter
        Dim DtaProductQty As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchase_closing_item_productqtyTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA2.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA3.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA4.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA5.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA6.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA7.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA81.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA8.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA9.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA10.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA11.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA12.Connection.ConnectionString = DTA.Connection.ConnectionString
            DtaProductQty.Connection.ConnectionString = Core_Module.GetConstring

            'hapus data item
            DTA.Delete(vPrimaryTransaksi)
            DTA1.Delete(vPrimaryTransaksi)
            DTA2.DeleteQuery(vPrimaryTransaksi)
            DTA3.DeleteQuery(vPrimaryTransaksi)
            DTA4.DeleteQuery(vPrimaryTransaksi)
            DTA5.Delete(vPrimaryTransaksi)
            DTA6.Delete(vPrimaryTransaksi)
            DTA7.DeleteQuery(vPrimaryTransaksi)
            DTA81.Delete(vPrimaryTransaksi)
            DTA8.Delete(vPrimaryTransaksi)
            DTA9.Delete(vPrimaryTransaksi)
            DTA10.Delete(vPrimaryTransaksi)
            DTA11.Delete(vPrimaryTransaksi)
            DTA12.DeleteQuery(vPrimaryTransaksi)
            DtaProductQty.DeleteQuery(vPrimaryTransaksi)

            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
            DTA2.Dispose()
            DTA3.Dispose()
            DTA4.Dispose()
            DTA5.Dispose()
            DTA6.Dispose()
            DTA7.Dispose()
            DTA81.Dispose()
            DTA8.Dispose()
            DTA9.Dispose()
            DTA10.Dispose()
            DTA11.Dispose()
            DTA12.Dispose()
        End Try
    End Function

    Public Function GetInAmount_Purchase_SalesReturnBedaPeriode2(ByVal vPrimaryCodeNo As Double, _
      ByVal vTanggalSampai As Date, ByVal vTanggalDari As Date, _
      ByVal TBL As DseTransaksiMonthlyClosing.DataTable25DataTable, ByVal ObjCD As Core_Module.ClsCekData, ByVal vDef_Prec As Integer) As Decimal
        Dim DTA1 As New DseTransaksiMonthlyClosingTableAdapters.DataTable4TableAdapter
        Dim TBL1 As New DseTransaksiMonthlyClosing.DataTable4DataTable
        Dim row As DseTransaksiMonthlyClosing.DataTable4Row
        Dim Dummy As Double = 0
        Dim TBLSR As New DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item9_1DataTable
        Dim DTASR As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchaseclosing_item9_1TableAdapter
        Dim rowSR As DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item9_1Row

        Try
            'in purchase
            DTA1.Connection.ConnectionString = Core_Module.GetConstring
            DTASR.Connection.ConnectionString = DTA1.Connection.ConnectionString

            DTA1.FillBy(TBL1, vTanggalSampai, vTanggalDari, vPrimaryCodeNo)
            If TBL1.Rows.Count > 0 Then
                row = TBL1.Rows(0)
                Dummy += row.Purchase_Amount
            End If

            'in sales return dari do yang beda periode sehingga bisa langsung diambil price at closingnya
            DTASR.Fill(TBLSR, vPrimaryCodeNo, vTanggalSampai, vTanggalDari)
            For Each rowSR In TBLSR.Rows
                If rowSR.DO_Date >= vTanggalDari And rowSR.DO_Date <= vTanggalSampai Then
                    rowSR.Unit_Price = 0
                ElseIf rowSR.DO_Date < vTanggalDari And rowSR.Unit_Price = 0 Then
                    rowSR.Unit_Price = Me.GetUnitPrice(vPrimaryCodeNo, ObjCD.GetFirstDateofTheMonth(rowSR.DO_Date), _
                    TBL, vDef_Prec)
                End If
                Dummy += Math.Round(rowSR.Unit_Price * rowSR.Total_Qty, vDef_Prec, MidpointRounding.AwayFromZero)
            Next

            Return Dummy
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA1.Dispose()
            TBL1.Dispose()
            DTASR.Dispose()
            TBLSR.Dispose()
        End Try
    End Function

    Public Function GetInAmount_Purchase_SalesReturnBedaPeriode(ByVal vPrimaryCodeNo As Double, _
    ByVal vTanggalSampai As Date, ByVal vTanggalDari As Date, _
    ByVal TBL As DseTransaksiMonthlyClosing.DataTable25DataTable, ByVal ObjCD As Core_Module.ClsCekData, _
    ByVal TBLPurchase As DseTransaksiMonthlyClosing.DataTable4DataTable, _
    ByVal TBLSR As DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item9_1DataTable, ByVal vDef_Prec As Integer) As Double
        Dim row As DseTransaksiMonthlyClosing.DataTable4Row
        Dim Dummy As Double = 0
        Dim rowSR As DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item9_1Row
        Dim Filter As String = ""
        Dim RowView As Data.DataRowView

        Try
            'in purchase
            Filter = "[Primary Product] = " + vPrimaryCodeNo.ToString
            TBLPurchase.DefaultView.RowFilter = Filter
            If TBLPurchase.DefaultView.Count > 0 Then
                row = TBLPurchase.DefaultView(0).Row
                Dummy += row.Purchase_Amount
            End If

            'in sales return dari do yang beda periode sehingga bisa langsung diambil price at closingnya
            TBLSR.DefaultView.RowFilter = Filter
            For Each RowView In TBLSR.DefaultView
                rowSR = RowView.Row
                If rowSR.DO_Date >= vTanggalDari And rowSR.DO_Date <= vTanggalSampai Then
                    rowSR.Unit_Price = 0
                ElseIf rowSR.DO_Date < vTanggalDari And rowSR.Unit_Price = 0 Then
                    rowSR.Unit_Price = Me.GetUnitPrice(vPrimaryCodeNo, ObjCD.GetFirstDateofTheMonth(rowSR.DO_Date), _
                    TBL, vDef_Prec)
                End If
                Dummy += Math.Round(rowSR.Unit_Price * rowSR.Total_Qty, vDef_Prec, MidpointRounding.AwayFromZero)
            Next

            Return Dummy
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetUnitPrice(ByVal vPrimaryProduct As Double, ByVal vClosingPeriode As Date, _
    ByVal TBL As DseTransaksiMonthlyClosing.DataTable25DataTable, ByVal vDef_Prec As Integer) As Decimal
        Dim rowClosing As DseTransaksiMonthlyClosing.DataTable25Row
        Dim Filter As String = "[Primary Product] = " + vPrimaryProduct.ToString

        Try
            TBL.DefaultView.RowFilter = Filter
            If TBL.DefaultView.Count = 0 Then
                Return 0
            Else
                rowClosing = TBL.DefaultView(0).Row
                If rowClosing.Ending_Qty = 0 Then
                    If rowClosing.Out_Qty = 0 Then
                        Return 0
                    Else
                        Return Math.Round(rowClosing.Out_Amount / rowClosing.Out_Qty, vDef_Prec, MidpointRounding.AwayFromZero)
                    End If
                Else
                    Return Math.Round(rowClosing.Ending_Amount / rowClosing.Ending_Qty, vDef_Prec, MidpointRounding.AwayFromZero)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        End Try
    End Function

    Private Function GetUnitPrice2(ByVal vPrimaryProduct As Double, ByVal vClosingPeriode As Date, _
    ByVal DTA As DseTransaksiMonthlyClosingTableAdapters.DataTable25TableAdapter, _
    ByVal TBL As DseTransaksiMonthlyClosing.DataTable25DataTable, ByVal vDef_Prec As Integer) As Double
        Dim rowClosing As DseTransaksiMonthlyClosing.DataTable25Row

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            DTA.Fill(TBL, vClosingPeriode, vPrimaryProduct)

            If TBL.Rows.Count = 0 Then
                Return 0
            Else
                rowClosing = TBL.Rows(0)
                If rowClosing.Ending_Qty = 0 Then
                    If rowClosing.Out_Qty = 0 Then
                        Return 0
                    Else
                        Return Math.Round(rowClosing.Out_Amount / rowClosing.Out_Qty, vDef_Prec, MidpointRounding.AwayFromZero)
                    End If
                Else
                    Return Math.Round(rowClosing.Ending_Amount / rowClosing.Ending_Qty, vDef_Prec, MidpointRounding.AwayFromZero)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        End Try
    End Function

    Public Function GetJmlItemIn(ByVal vPrimaryProduct As Double) As Double
        Dim DTA As New DseMonthlyClosingNewTableAdapters.tbltrans_purchase_closing_item_detailTableAdapter
        Dim RetVal As Nullable(Of Decimal)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetJmlByTipeItem(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem, vPrimaryProduct)

            If RetVal.HasValue = True Then
                Return RetVal.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        End Try
    End Function
End Class
