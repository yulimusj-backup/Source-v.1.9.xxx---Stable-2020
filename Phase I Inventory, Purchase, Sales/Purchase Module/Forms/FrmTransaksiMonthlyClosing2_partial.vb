Option Explicit On
Option Strict Off

Partial Public Class FrmTransaksiMonthlyClosing2
    Dim DTAProduct As New DseTransaksiMonthlyClosing2TableAdapters.tblmaster_productTableAdapter
    Dim TBLProduct As New DseTransaksiMonthlyClosing2.tblmaster_productDataTable
    Dim rowProduct As DseTransaksiMonthlyClosing2.tblmaster_productRow
    Dim TBLInAdjust As New DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item6DataTable
    Dim DTAAdjustIn As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchaseclosing_item6TableAdapter
    Dim rowAdjust, rowAdjust1 As DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item6Row
    Dim TipeTransaksiStockInOut As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut = _
    Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust
    Dim DTATaking1 As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_stock_takingTableAdapter
    Dim TBLTaking1 As New DseTransaksiMonthlyClosing.tbltrans_stock_takingDataTable
    Dim rowTaking1 As DseTransaksiMonthlyClosing.tbltrans_stock_takingRow
    Dim TBL_In_SR As New DseTransaksiMonthlyClosing.DataTable20DataTable
    Dim row_In_SR As DseTransaksiMonthlyClosing.DataTable20Row
    Dim DTA_In_SR As New DseTransaksiMonthlyClosingTableAdapters.DataTable20TableAdapter
    Dim DTAPurchase As New DseTransaksiMonthlyClosing2TableAdapters.tbltrans_aop_itemTableAdapter
    Dim TBLPurchase As New DseTransaksiMonthlyClosing2.tbltrans_aop_itemDataTable
    Dim rowPurchase As DseTransaksiMonthlyClosing2.tbltrans_aop_itemRow
    Dim DTAPR_Rev_AOP As New DseTransaksiMonthlyClosing2TableAdapters.tbltrans_aop_item_returnTableAdapter
    Dim TBLPR_Rev_AOP As New DseTransaksiMonthlyClosing2.tbltrans_aop_item_returnDataTable
    Dim rowPR_Rev_AOP As DseTransaksiMonthlyClosing2.tbltrans_aop_item_returnRow
    Dim DTA_HIstoricalPrice As New DseTransaksiMonthlyClosingTableAdapters.DataTable29TableAdapter
    Dim TBL_HistoricalPrice As New DseTransaksiMonthlyClosing.DataTable29DataTable
    Dim row_HistoricalPrice As DseTransaksiMonthlyClosing.DataTable29Row
    Dim DTATaking2 As New DseTransaksiMonthlyClosingTableAdapters.laporan_stock_taking_view2TableAdapter
    Dim TBLTaking2 As New DseTransaksiMonthlyClosing.laporan_stock_taking_view2DataTable
    Dim rowTaking2 As DseTransaksiMonthlyClosing.laporan_stock_taking_view2Row
    Dim DTASales As New DseTransaksiMonthlyClosing2TableAdapters.tbltrans_salesinvoice_itemTableAdapter
    Dim TBLSales As New DseTransaksiMonthlyClosing2.tbltrans_salesinvoice_itemDataTable
    Dim rowSales As DseTransaksiMonthlyClosing2.tbltrans_salesinvoice_itemRow
    Dim DTASalesReturn As New DseTransaksiMonthlyClosing2TableAdapters.tbltrans_salesinvoice_item_returnTableAdapter
    Dim TBLSalesReturn As New DseTransaksiMonthlyClosing2.tbltrans_salesinvoice_item_returnDataTable
    Dim rowSalesReturn As DseTransaksiMonthlyClosing2.tbltrans_salesinvoice_item_returnRow
    Dim DTAOutTaking As New DseTransaksiMonthlyClosing2TableAdapters.stock_taking_out_adjustmentTableAdapter
    Dim TBLOutTaking As New DseTransaksiMonthlyClosing2.stock_taking_out_adjustmentDataTable
    Dim rowOutTaking As DseTransaksiMonthlyClosing2.stock_taking_out_adjustmentRow
    Dim DTA_PR As New DseTransaksiMonthlyClosing2TableAdapters.tbltrans_purchase_return_itemTableAdapter
    Dim TBL_PR As New DseTransaksiMonthlyClosing2.tbltrans_purchase_return_itemDataTable
    Dim row_PR As DseTransaksiMonthlyClosing2.tbltrans_purchase_return_itemRow
    Dim DTA_OutAdjust As New DseTransaksiMonthlyClosing2TableAdapters.tbltrans_out_adjustTableAdapter
    Dim TBL_OutAdjust As New DseTransaksiMonthlyClosing2.tbltrans_out_adjustDataTable
    Dim row_OutAdjust As DseTransaksiMonthlyClosing2.tbltrans_out_adjustRow

    Dim TBLBundling1 As New DseTransaksiMonthlyClosing2.TblBundling1DataTable
    Dim DTABundling1 As New DseTransaksiMonthlyClosing2TableAdapters.TblBundling1TableAdapter
    Dim rowBundling1 As DseTransaksiMonthlyClosing2.TblBundling1Row

    Dim DTABeginingBalance As New DseTransaksiMonthlyClosing2TableAdapters.TBLBeginingBalanceTableAdapter
    Dim TBLBeginingBalance As New DseTransaksiMonthlyClosing2.TBLBeginingBalanceDataTable
    Dim rowBeginingBalance As DseTransaksiMonthlyClosing2.TBLBeginingBalanceRow
    Dim DTABeginingBalance1 As New DseTransaksiMonthlyClosing2TableAdapters.TBLBeginingBalance1TableAdapter
    Dim TBLBeginingBalance1 As New DseTransaksiMonthlyClosing2.TBLBeginingBalance1DataTable
    Dim rowBeginingBalance1 As DseTransaksiMonthlyClosing2.TBLBeginingBalance1Row
    Dim ObjTaking As New Inventory_Module.clsTransaksiStockTaking

    Public TBLItemDetail As New DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailDataTable
    Public rowItemDetail As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailRow

    'tabel untuk menampung data yang sudah beres 100% ketika hitung looping menyelesaikan siklik bundling
    Public TBLInvPool_Dummy As New DseMonthlyClosingNew.tblItemDetailPoolDataTable
    Public TBLItemDetail_Dummy As New DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailDataTable

    Dim ObjSetting As New Core_Module.ClsTransaksiSetting

    Dim ObjAkun As New AccModule.ClsMasterAkun
    Public ObjStockInOut As New Inventory_Module.clsTransaksiStockInOut

    Dim TBL_In_SR1 As New DseTransaksiMonthlyClosing.DataTable1DataTable
    Dim row_In_SR1 As DseTransaksiMonthlyClosing.DataTable1Row
    Dim DTA_In_SR1 As New DseTransaksiMonthlyClosingTableAdapters.DataTable1TableAdapter

    Dim DTA_SR2 As New DseTransaksiMonthlyClosing2TableAdapters.TBLSRTableAdapter
    Dim TBL_SR2 As New DseTransaksiMonthlyClosing2.TBLSRDataTable
    Dim row_SR2 As DseTransaksiMonthlyClosing2.TBLSRRow
    Dim DTA_SR2_Dummy As New DseTransaksiMonthlyClosing2TableAdapters.TBLSRTableAdapter
    Dim TBL_SR2_Dummy As New DseTransaksiMonthlyClosing2.TBLSRDataTable
    Dim row_SR2_Dummy As DseTransaksiMonthlyClosing2.TBLSRRow
    Dim DTA_SR2_Cost As New DseTransaksiMonthlyClosing2TableAdapters.TBLSR_CostTableAdapter
    Dim TBL_SR2_Cost As New DseTransaksiMonthlyClosing2.TBLSR_CostDataTable
    Dim row_SR2_Cost As DseTransaksiMonthlyClosing2.TBLSR_CostRow
    Dim TipeSR_CostInvoice_TipeTranskasiAsal As ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutSales
    Dim TipeSR_CostInvoice_TipeItemDetail As ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
    Dim TglDOMinimum As Date = Date.MaxValue

    Dim DTASetting As New DseTransaksiMonthlyClosing2TableAdapters.tbl_setting_adjustmentTableAdapter
    Dim TBLSetting As New DseTransaksiMonthlyClosing2.tbl_setting_adjustmentDataTable


    Dim rowSetting As DseTransaksiMonthlyClosing2.tbl_setting_adjustmentRow
    Dim TBLSettingItem As New DseTransaksiMonthlyClosing2.tbltrans_purchase_closing_item_tipe_adjustmentDataTable
    Dim rowSettingItem As DseTransaksiMonthlyClosing2.tbltrans_purchase_closing_item_tipe_adjustmentRow
    Dim tblPool As New DseMonthlyClosingNew.tblItemDetailPoolDataTable
    Dim rowPool As DseMonthlyClosingNew.tblItemDetailPoolRow

    Dim Calc_TotalStep As Integer = 0
    Dim Calc_CurrentStep As Integer = 0
    Dim TBL_Pending1 As New DseMonthlyClosingNew.tblItemDetailPoolDataTable
    Dim TBL_Pending2 As New DseMonthlyClosingNew.tblItemDetailPoolDataTable
    Dim RecalculateNeeded As Boolean = False

    Dim DTAItemClosing As New DseMonthlyClosingNewTableAdapters.DataTable7TableAdapter
    Dim TBLItemClosing As New DseMonthlyClosingNew.DataTable7DataTable
    Dim HitungOtomatis As Boolean = True
    Dim ObjIdentifier As New CustomSMESS.ClsIdentifier
    Dim custom_module As String = ""

    'Private Sub LoadItemDetail()
    '    'Dim StartDate, EndDate As Date
    '    'Dim Filter As String = ""
    '    'Dim OldText As String = "Monthly Inventory Closing"
    '    'Dim TotalAdjustment As Double = 0
    '    'Dim rowDGV As DataGridViewRow
    '    'Dim cmd As String = ""
    '    'Dim counter As Double = 0
    '    'Dim DefPrec As Integer = 0
    '    'Dim counter1 As Double = 0
    '    'Dim TipePR As ClsTransaksiPurchaseReturn.EnumTipe_PurchaseReturn = ClsTransaksiPurchaseReturn.EnumTipe_PurchaseReturn.Nota_Retur
    '    'Dim TipeStockOut As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust
    '    'Dim ListProduct As New List(Of Double)
    '    'Dim ListProductReject As New List(Of Double)
    '    'Dim Lvl As Integer = 0
    '    'Dim TipeStockInOut As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust
    '    'Dim PrimaryClosingSebelum As Double = -1
    '    'Dim TipeItem As ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.EndingBalanceItem
    '    'Dim PrimaryStockTaking As Double = -1
    '    'Dim TglTaking As Date = #1/1/1950#
    '    'Dim IndexDGV As Double = 0
    '    'Dim AmountDummy As Decimal = 0
    '    'Dim QtyDummy As Decimal = 0
    '    'Dim CostingMethod As ClsTransaksiMonthlyClosingItem.EnumCostingMethod = ClsTransaksiMonthlyClosingItem.EnumCostingMethod.Average
    '    'Dim DummyCostingMethod As String = ""
    '    'Dim DetailAmountDummy As Decimal = 0
    '    'Dim ObjAccClosing As New AccModule.ClsTransaksiClosing
    '    'Dim QtySRDummy As Decimal = 0
    '    'Dim QtyCost As Decimal = 0
    '    'Dim row_In_SR2 As DseTransaksiMonthlyClosing.DataTable1Row
    '    'Dim row_In_SR3 As DseTransaksiMonthlyClosing.DataTable20Row
    '    'Dim InterimString As String = ":"

    '    'Try
    '    '    Me.TglDOMinimum = Date.MaxValue
    '    '    DummyCostingMethod = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
    '    '    Core_Module.ClsTransaksiSetting.EnumNamaParameter.CostingMethod)
    '    '    If DummyCostingMethod = "" Then
    '    '        CostingMethod = ClsTransaksiMonthlyClosingItem.EnumCostingMethod.Average
    '    '    Else
    '    '        CostingMethod = DummyCostingMethod
    '    '    End If
    '    '    DefPrec = Me.objCurrency.GetDecimal_DefaultCurrency

    '    '    Windows.Forms.Cursor.Current = Cursors.WaitCursor
    '    '    StartDate = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value)
    '    '    EndDate = Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value)

    '    '    'reset data untuk persiapan perhitungan

    '    '    Me.dgvItem2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
    '    '    Me.dgvItem2.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
    '    '    Me.dgvItem2.Rows.Clear()
    '    '    Me.dgvItemAdjust.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
    '    '    Me.dgvItemAdjust.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
    '    '    Me.dgvItemAdjust.Rows.Clear()
    '    '    Me.TBLItemDetail.Rows.Clear()
    '    '    Me.TBL_In_SR.Rows.Clear()
    '    '    Me.TBL_OutAdjust.Rows.Clear()
    '    '    Me.TBLInAdjust.Rows.Clear()
    '    '    Me.TBLTaking1.Rows.Clear()
    '    '    Me.TBLTaking2.Rows.Clear()
    '    '    Me.TBLPurchase.Rows.Clear()
    '    '    Me.TBLPR_Rev_AOP.Rows.Clear()
    '    '    Me.TBL_HistoricalPrice.Rows.Clear()
    '    '    Me.TBLSales.Rows.Clear()
    '    '    Me.TBLSalesReturn.Rows.Clear()
    '    '    Me.TBLOutTaking.Rows.Clear()
    '    '    Me.TBL_PR.Rows.Clear()
    '    '    Me.TBL_OutAdjust.Rows.Clear()
    '    '    Me.TBLBundling1.Rows.Clear()
    '    '    Me.TBLBefore.Rows.Clear()
    '    '    Me.TBLAfter.Rows.Clear()
    '    '    Me.TBLBeginingBalance.Rows.Clear()
    '    '    Me.TBLBeginingBalance1.Rows.Clear()
    '    '    Me.TBLItemDetail.Rows.Clear()

    '    '    '0. A. load data product
    '    '    Me.TBLProduct.Rows.Clear()
    '    '    Me.DTAProduct.Fill(Me.TBLProduct)

    '    '    '0. B. Load Data beginning balance
    '    '    If (Me.ObjClosing.GetJml = 0 And Me.baru = True) Or (Me.ObjClosing.GetJml = 1 And Me.baru = False) Then
    '    '        '(i). Jika closing pertama maka ambil datanya dari begining balance(stock in adjust begining balance)
    '    '        'tidak perlu isikan disini

    '    '        'Me.DTABeginingBalance.Fill(Me.TBLBeginingBalance, TipeStockInOut)
    '    '        'For Each Me.rowBeginingBalance In Me.TBLBeginingBalance.Rows
    '    '        'Me.rowItemDetail = Me.TBLItemDetail.Newtbltrans_purchase_closing_item_detailRow
    '    '        'With rowItemDetail
    '    '        '.amount = Me.rowBeginingBalance.Amount
    '    '        '.primary = -1
    '    '        '.primary_code_no = Me.rowBeginingBalance.Primary_Product
    '    '        '.primary_transaksi = -1
    '    '        '.primary_transaksi_asal = Me.rowBeginingBalance.Primary_Transaksi
    '    '        '.qty = Me.rowBeginingBalance.Qty
    '    '        '.tgl_transaksi_asal = Me.rowBeginingBalance.Tanggal_Transaksi
    '    '        '.tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.BeginingBalanceItem
    '    '        '.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust
    '    '        '.unit_price = Me.ObjCD.DigitGrouping(Me.rowBeginingBalance.Amount / Me.rowBeginingBalance.Qty, DefPrec)
    '    '        'End With
    '    '        'Me.TBLItemDetail.Rows.Add(Me.rowItemDetail)
    '    '        'Next
    '    '    Else
    '    '        '(ii). jika closing selanjutnya, maka ambil begining balance dari ending balance transaksi sebelumnya
    '    '        PrimaryClosingSebelum = Me.ObjClosing.GetPrimary(Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value - TimeSpan.FromDays(5)))
    '    '        Me.DTABeginingBalance1.Fill(Me.TBLBeginingBalance1, PrimaryClosingSebelum, TipeItem)
    '    '        For Each Me.rowBeginingBalance1 In Me.TBLBeginingBalance1.Rows
    '    '            Me.rowItemDetail = Me.TBLItemDetail.Newtbltrans_purchase_closing_item_detailRow
    '    '            With rowItemDetail
    '    '                .amount = Me.rowBeginingBalance1.amount
    '    '                .primary = -1
    '    '                .primary_code_no = Me.rowBeginingBalance1.primary_code_no
    '    '                .primary_transaksi = -1
    '    '                .primary_transaksi_asal = Me.rowBeginingBalance1.primary_transaksi_asal
    '    '                .qty = Me.rowBeginingBalance1.qty
    '    '                .tgl_transaksi_asal = Me.rowBeginingBalance1.tgl_transaksi_asal
    '    '                .tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.BeginingBalanceItem
    '    '                .tipe_transaksi_asal = Me.rowBeginingBalance1.tipe_transaksi_asal
    '    '                .unit_price = Me.ObjCD.DigitGrouping(Me.rowBeginingBalance1.unit_price, DefPrec)
    '    '                .return_price = 0
    '    '                .primary_gain_account = -1
    '    '                .primary_loss_account = -1
    '    '                .primary_adjustment_type = -1
    '    '                .remarks = ""
    '    '            End With
    '    '            Me.TBLItemDetail.Rows.Add(Me.rowItemDetail)
    '    '        Next
    '    '    End If

    '    '    '1. load data adjustment, stock taking, sales return
    '    '    'A. Adjust
    '    '    DTA_HIstoricalPrice.Fill(TBL_HistoricalPrice)
    '    '    TotalAdjustment = DTAAdjustIn.Fill(TBLInAdjust, TipeTransaksiStockInOut, EndDate, StartDate)
    '    '    DetailAmountDummy = 0
    '    '    For Each rowAdjust In TBLInAdjust.Rows
    '    '        Me.Text = OldText + "- Loading adjustment data (" + CStr(Me.dgvItemAdjust.Rows.Count) _
    '    '        + " of " + TotalAdjustment.ToString + ")"

    '    '        Me.dgvItemAdjust.Rows.Add("")
    '    '        rowDGV = Me.dgvItemAdjust.Rows(Me.dgvItemAdjust.Rows.Count - 1)
    '    '        With rowDGV
    '    '            .Cells(Me.ColNoAdjust.Index).Value = Me.dgvItemAdjust.Rows.Count
    '    '            .Cells(Me.ColJenisDocumentAdjustStr.Index).Value = "Stock In Adjust"
    '    '            .Cells(Me.ColDocumentNo.Index).Value = rowAdjust.No_Transaksi_Stock_In_Adjust
    '    '            .Cells(Me.ColCodeNoAdjust.Index).Value = rowAdjust.Product_Code
    '    '            .Cells(Me.ColDescAdjust.Index).Value = rowAdjust.Description
    '    '            .Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(rowAdjust.Qty, rowAdjust._Decimal)
    '    '            .Cells(Me.ColUnitAdjust.Index).Value = rowAdjust.Unit_Code

    '    '            If (Me.ObjClosing.GetJml = 0 And Me.baru = True) Or (Me.ObjClosing.GetJml = 1 And Me.baru = False) Then
    '    '                'closing pertama ambil data dari begining balance
    '    '                .Cells(Me.ColUnitPriceAdjust.Index).Value = Me.ObjCD.DigitGrouping(rowAdjust.Amount, Me.DefaultPrecision)
    '    '                .Cells(Me.ColRemarksAdjust.Index).Value = "Begining Balance"
    '    '                rowAdjust.Unit_Price = Me.ObjCD.DigitGrouping(rowAdjust.Amount / rowAdjust.Qty, Me.DefaultPrecision)
    '    '            Else
    '    '                'jika closing normal, ambil data dari po
    '    '                cmd = "[Primary Product] =" + rowAdjust.Primary_Product.ToString
    '    '                TBL_HistoricalPrice.DefaultView.RowFilter = cmd
    '    '                If TBL_HistoricalPrice.DefaultView.Count > 0 Then
    '    '                    row_HistoricalPrice = TBL_HistoricalPrice.DefaultView(0).Row
    '    '                    If Me.ObjCD.GetFirstDateofTheMonth(row_HistoricalPrice.Tanggal_Transaksi) = #2/1/2014# Then
    '    '                        ''#sementara, hanya untuk closing lmb, seharusnya nilainya 0
    '    '                        '.Cells(Me.ColUnitPriceAdjust.Index).Value = Me.ObjCD.DigitGrouping(1000, Me.DefaultPrecision)
    '    '                        '.Cells(Me.ColRemarksAdjust.Index).Value = ""
    '    '                        'rowAdjust.Unit_Price = 1000
    '    '                        ''#
    '    '                        .Cells(Me.ColUnitPriceAdjust.Index).Value = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
    '    '                        .Cells(Me.ColRemarksAdjust.Index).Value = ""
    '    '                        rowAdjust.Unit_Price = 0
    '    '                    Else
    '    '                        .Cells(Me.ColRemarksAdjust.Index).Value = "refer to purchase price from: " _
    '    '                        + row_HistoricalPrice.Supplier_Name + "(" + row_HistoricalPrice.Supplier_Code + _
    '    '                        ") with aop number: " + row_HistoricalPrice.No_Transaksi + " on " + Format(row_HistoricalPrice.Tanggal_Transaksi, _
    '    '                        "dd MMM yyyy")
    '    '                        rowAdjust.Unit_Price = row_HistoricalPrice.At_Price
    '    '                        .Cells(Me.ColUnitPriceAdjust.Index).Value = Me.ObjCD.DigitGrouping(rowAdjust.Unit_Price, Me.DefaultPrecision)
    '    '                    End If
    '    '                Else
    '    '                    ''#sementara, hanya untuk closing lmb, seharusnya nilainya 0
    '    '                    '.Cells(Me.ColUnitPriceAdjust.Index).Value = Me.ObjCD.DigitGrouping(1000, Me.DefaultPrecision)
    '    '                    '.Cells(Me.ColRemarksAdjust.Index).Value = ""
    '    '                    ''rowAdjust.Unit_Price = 0
    '    '                    'rowAdjust.Unit_Price = 1000
    '    '                    ''#
    '    '                    .Cells(Me.ColUnitPriceAdjust.Index).Value = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
    '    '                    .Cells(Me.ColRemarksAdjust.Index).Value = ""
    '    '                    rowAdjust.Unit_Price = 0

    '    '                End If
    '    '                rowAdjust.Amount = Me.ObjCD.DigitGrouping(rowAdjust.Unit_Price * rowAdjust.Qty, DefPrec)
    '    '            End If

    '    '            .Cells(Me.ColPrimaryCodeNoAdjust.Index).Value = rowAdjust.Primary_Product
    '    '            .Cells(Me.ColPrimaryUnitAdjust.Index).Value = rowAdjust.Primary_Units
    '    '            .Cells(Me.ColPrimaryStockInAdjust.Index).Value = rowAdjust.Primary_Stock_In_Adjust
    '    '            .Cells(Me.ColJenisDokumen.Index).Value = ClsTransaksiMonthlyClosingItem.EnumTipeDokumenAcuan.StockInAdjust
    '    '        End With

    '    '        Filter = "[primary transaksi asal] = " + rowAdjust.Primary_Stock_In_Adjust.ToString + _
    '    '        " AND [primary code no] = " + rowAdjust.Primary_Product.ToString + _
    '    '        " AND [tipe transaksi asal] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust).ToString + _
    '    '        " AND [tipe item detail] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem).ToString
    '    '        Me.TBLItemDetail.DefaultView.RowFilter = Filter
    '    '        If Me.TBLItemDetail.DefaultView.Count = 0 Then
    '    '            Me.rowItemDetail = Me.TBLItemDetail.Newtbltrans_purchase_closing_item_detailRow
    '    '            With rowItemDetail
    '    '                If (Me.ObjClosing.GetJml = 0 And Me.baru = True) Or (Me.ObjClosing.GetJml = 1 And Me.baru = False) Then
    '    '                    .amount = rowAdjust.Amount
    '    '                    .unit_price = Me.ObjCD.DigitGrouping(rowAdjust.Unit_Price, DefPrec)
    '    '                Else
    '    '                    .amount = Me.ObjCD.DigitGrouping(rowAdjust.Unit_Price * rowAdjust.Qty, DefPrec)
    '    '                    .unit_price = Me.ObjCD.DigitGrouping(rowAdjust.Unit_Price, DefPrec)
    '    '                End If
    '    '                .primary = -1
    '    '                .primary_code_no = rowAdjust.Primary_Product
    '    '                .primary_transaksi = -1
    '    '                .primary_transaksi_asal = rowAdjust.Primary_Stock_In_Adjust
    '    '                .qty = rowAdjust.Qty
    '    '                .tgl_transaksi_asal = rowAdjust.Tanggal_Transaksi
    '    '                .tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
    '    '                .tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust
    '    '                .return_price = 0
    '    '                .primary_gain_account = rowAdjust.Primary_Loss_Account
    '    '                .primary_loss_account = rowAdjust.Primary_Gain_Account
    '    '                .primary_adjustment_type = rowAdjust.PrimaryAdjType
    '    '                rowAdjust.Remarks = rowAdjust.Remarks.Trim
    '    '                If rowAdjust.Remarks <> "" Then
    '    '                    InterimString = " : "
    '    '                Else
    '    '                    InterimString = ""
    '    '                End If
    '    '                If rowAdjust.Adjusment_Type_Code <> "" Then
    '    '                    If (rowAdjust.Adjustment_Type_Desc + InterimString + rowAdjust.Remarks).Length > 500 Then
    '    '                        .remarks = (rowAdjust.Adjustment_Type_Desc + InterimString + rowAdjust.Remarks).Substring(0, 500)
    '    '                    Else
    '    '                        .remarks = rowAdjust.Adjustment_Type_Desc + InterimString + rowAdjust.Remarks
    '    '                    End If
    '    '                Else
    '    '                    If (rowAdjust.Remarks).Length > 500 Then
    '    '                        .remarks = (rowAdjust.Remarks).Substring(0, 500)
    '    '                    Else
    '    '                        .remarks = rowAdjust.Remarks
    '    '                    End If
    '    '                End If
    '    '            End With
    '    '            Me.TBLItemDetail.Rows.Add(Me.rowItemDetail)
    '    '        Else
    '    '            Me.rowItemDetail = Me.TBLItemDetail.DefaultView(0).Row
    '    '            Me.TBLInAdjust.DefaultView.RowFilter = "[Primary Stock In Adjust] = " + rowAdjust.Primary_Stock_In_Adjust.ToString + _
    '    '            " AND [Primary Product] = " + rowAdjust.Primary_Product.ToString
    '    '            AmountDummy = 0
    '    '            QtyDummy = 0
    '    '            For Each dv As DataRowView In Me.TBLInAdjust.DefaultView
    '    '                Me.rowAdjust1 = dv.Row
    '    '                AmountDummy += Me.rowAdjust1.Amount
    '    '                QtyDummy += Me.rowAdjust1.Qty
    '    '            Next
    '    '            Me.rowItemDetail.qty = QtyDummy
    '    '            Me.rowItemDetail.amount = AmountDummy
    '    '            Me.rowItemDetail.unit_price = Me.ObjCD.DigitGrouping(Me.rowItemDetail.amount / Me.rowItemDetail.qty, DefPrec)
    '    '            If (Me.rowItemDetail.remarks + " - " + rowAdjust.Remarks).Length > 500 Then
    '    '                Me.rowItemDetail.remarks = (Me.rowItemDetail.remarks + " - " + rowAdjust.Remarks).Substring(0, 500)
    '    '            Else
    '    '                Me.rowItemDetail.remarks = Me.rowItemDetail.remarks + " - " + rowAdjust.Remarks
    '    '            End If
    '    '        End If
    '    '    Next

    '    '    'B. Stock Taking
    '    '    If Me.ObjCD.GetFirstDateofTheMonth(Me.ObjAccClosing.GetConvertionDate) <= Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value) Then
    '    '        DTATaking1.Fill(TBLTaking1, EndDate, StartDate)
    '    '        TotalAdjustment = 0
    '    '        For Each rowTaking1 In TBLTaking1.Rows
    '    '            TotalAdjustment += DTATaking2.ScalarQuery_GetJmlByPrimaryTaking(rowTaking1.Primary_Stock_Taking).Value
    '    '        Next
    '    '        'masukkan ke dalam item adjustment
    '    '        counter = 0
    '    '        For Each rowTaking1 In TBLTaking1.Rows
    '    '            DTATaking2.Fill(TBLTaking2, rowTaking1.Primary_Stock_Taking)
    '    '            For Each rowTaking2 In TBLTaking2.Rows
    '    '                counter += 1
    '    '                Me.Text = OldText + "- Loading adjustment data (" + CStr(counter) _
    '    '                + " of " + TotalAdjustment.ToString + ")"
    '    '                Me.dgvItemAdjust.Rows.Add("")
    '    '                rowDGV = Me.dgvItemAdjust.Rows(Me.dgvItemAdjust.Rows.Count - 1)
    '    '                With rowDGV
    '    '                    .Cells(Me.ColNoAdjust.Index).Value = Me.dgvItemAdjust.Rows.Count
    '    '                    .Cells(Me.ColJenisDocumentAdjustStr.Index).Value = "Stock Taking In Adjustment"
    '    '                    .Cells(Me.ColDocumentNo.Index).Value = rowTaking1.Stock_Taking_Number
    '    '                    .Cells(Me.ColCodeNoAdjust.Index).Value = rowTaking2.Product_Code
    '    '                    .Cells(Me.ColDescAdjust.Index).Value = rowTaking2.Description
    '    '                    .Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(rowTaking2.Qty, rowTaking2._Decimal)
    '    '                    .Cells(Me.ColUnitAdjust.Index).Value = rowTaking2.Unit_Code
    '    '                    'If rowTaking2.PO_Number = "" Then
    '    '                    '.Cells(Me.ColRemarksAdjust.Index).Value = ""
    '    '                    'Else
    '    '                    '.Cells(Me.ColRemarksAdjust.Index).Value = "refer to purchase price from: " _
    '    '                    '+ rowTaking2.Supplier_Name + "(" + rowTaking2.Supplier_Code + _
    '    '                    '") with po number: " + rowTaking2.PO_Number + " on " + Format(rowTaking2.PO_Date, _
    '    '                    '"dd MMM yyyy")
    '    '                    'End If

    '    '                    cmd = "[Primary Product] =" + rowTaking2.Primary_Product.ToString
    '    '                    TBL_HistoricalPrice.DefaultView.RowFilter = cmd
    '    '                    If TBL_HistoricalPrice.DefaultView.Count > 0 Then
    '    '                        row_HistoricalPrice = TBL_HistoricalPrice.DefaultView(0).Row
    '    '                        If Me.ObjCD.GetFirstDateofTheMonth(row_HistoricalPrice.Tanggal_Transaksi) = #2/1/2014# Then
    '    '                            .Cells(Me.ColUnitPriceAdjust.Index).Value = Me.ObjCD.DigitGrouping(0, DefPrec)
    '    '                            .Cells(Me.ColRemarksAdjust.Index).Value = ""
    '    '                            rowTaking2.Unit_Price = 0
    '    '                        Else
    '    '                            .Cells(Me.ColRemarksAdjust.Index).Value = "refer to purchase price from: " _
    '    '                            + row_HistoricalPrice.Supplier_Name + "(" + row_HistoricalPrice.Supplier_Code + _
    '    '                            ") with aop number: " + row_HistoricalPrice.No_Transaksi + " on " + Format(row_HistoricalPrice.Tanggal_Transaksi, _
    '    '                            "dd MMM yyyy")
    '    '                            rowTaking2.Unit_Price = Me.ObjCD.DigitGrouping(row_HistoricalPrice.At_Price, DefPrec)
    '    '                            .Cells(Me.ColUnitPriceAdjust.Index).Value = Me.ObjCD.DigitGrouping(rowTaking2.Unit_Price, DefPrec)
    '    '                        End If
    '    '                    Else
    '    '                        rowTaking2.Unit_Price = 0
    '    '                        .Cells(Me.ColUnitPriceAdjust.Index).Value = Me.ObjCD.DigitGrouping(0, DefPrec)
    '    '                        .Cells(Me.ColRemarksAdjust.Index).Value = ""
    '    '                    End If

    '    '                    .Cells(Me.ColPrimaryCodeNoAdjust.Index).Value = rowTaking2.Primary_Product
    '    '                    .Cells(Me.ColPrimaryUnitAdjust.Index).Value = rowTaking2.Primary_Unit
    '    '                    .Cells(Me.ColPrimaryStockInAdjust.Index).Value = rowTaking2.Primary_Stock_Taking
    '    '                    .Cells(Me.ColJenisDokumen.Index).Value = ClsTransaksiMonthlyClosingItem.EnumTipeDokumenAcuan.StockTakingInAdjustment
    '    '                End With

    '    '                Filter = "[primary transaksi asal] = " + rowTaking2.Primary_Stock_Taking.ToString + _
    '    '                " AND [primary code no] = " + rowTaking2.Primary_Product.ToString + _
    '    '                " AND [tipe transaksi asal] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InStockTakeAdjust).ToString + _
    '    '                " AND [tipe item detail] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem).ToString
    '    '                Me.TBLItemDetail.DefaultView.RowFilter = Filter
    '    '                If Me.TBLItemDetail.DefaultView.Count = 0 Then
    '    '                    Me.rowItemDetail = Me.TBLItemDetail.Newtbltrans_purchase_closing_item_detailRow
    '    '                    With rowItemDetail
    '    '                        .amount = Me.ObjCD.DigitGrouping(rowTaking2.Unit_Price * rowTaking2.Qty, DefPrec)
    '    '                        .unit_price = Me.ObjCD.DigitGrouping(rowTaking2.Unit_Price, DefPrec)
    '    '                        .primary = -1
    '    '                        .primary_code_no = rowTaking2.Primary_Product
    '    '                        .primary_transaksi = -1
    '    '                        If PrimaryStockTaking <> rowTaking2.Primary_Stock_Taking Then
    '    '                            TglTaking = ObjTaking.GetStartDate(rowTaking2.Primary_Stock_Taking)
    '    '                            PrimaryStockTaking = rowTaking2.Primary_Stock_Taking
    '    '                        End If

    '    '                        .primary_transaksi_asal = rowTaking2.Primary_Stock_Taking
    '    '                        .qty = rowTaking2.Qty
    '    '                        .tgl_transaksi_asal = TglTaking
    '    '                        .tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
    '    '                        .tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InStockTakeAdjust
    '    '                        .return_price = 0
    '    '                        .primary_gain_account = -1
    '    '                        .primary_loss_account = -1
    '    '                        .primary_adjustment_type = -1
    '    '                        .remarks = "stock taking"
    '    '                    End With
    '    '                    Me.TBLItemDetail.Rows.Add(Me.rowItemDetail)
    '    '                Else
    '    '                    Me.rowItemDetail = Me.TBLItemDetail.DefaultView(0).Row
    '    '                    Me.TBLTaking2.DefaultView.RowFilter = "[Primary Stock Taking] = " + rowTaking2.Primary_Stock_Taking.ToString + _
    '    '                    " AND [Primary Product] = " + rowTaking2.Primary_Product.ToString
    '    '                    AmountDummy = 0
    '    '                    QtyDummy = 0
    '    '                    For Each dv As DataRowView In Me.TBLTaking2.DefaultView
    '    '                        Me.rowTaking2 = dv.Row
    '    '                        AmountDummy += Me.ObjCD.DigitGrouping(rowTaking2.Unit_Price * rowTaking2.Qty, DefPrec)
    '    '                        QtyDummy += Me.rowTaking2.Qty
    '    '                    Next
    '    '                    Me.rowItemDetail.qty = QtyDummy
    '    '                    Me.rowItemDetail.amount = AmountDummy
    '    '                    Me.rowItemDetail.unit_price = Me.ObjCD.DigitGrouping(Me.rowItemDetail.amount / Me.rowItemDetail.qty, DefPrec)
    '    '                End If
    '    '            Next
    '    '        Next

    '    '        'C. Stock In Sales Return yang di jual pada bulan yang sama
    '    '        TotalAdjustment = DTA_In_SR.Fill(TBL_In_SR, EndDate, StartDate)
    '    '        counter = 0
    '    '        For Each row_In_SR In TBL_In_SR.Rows
    '    '            'rekomendasikan closing price bulan kemarin
    '    '            counter1 = Me.GetUnitPrice(row_In_SR.Primary_Product, Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value.Date _
    '    '            - TimeSpan.FromDays(1)))
    '    '            row_In_SR.Unit_Price = counter1
    '    '        Next
    '    '        For Each row_In_SR In TBL_In_SR.Rows
    '    '            counter += 1
    '    '            Me.Text = OldText + "- Loading in sales return data (" + CStr(counter) _
    '    '            + " of " + TotalAdjustment.ToString + ")"
    '    '            Me.dgvItemAdjust.Rows.Add("")
    '    '            rowDGV = Me.dgvItemAdjust.Rows(Me.dgvItemAdjust.Rows.Count - 1)
    '    '            With rowDGV
    '    '                .Cells(Me.ColNoAdjust.Index).Value = Me.dgvItemAdjust.Rows.Count
    '    '                .Cells(Me.ColJenisDocumentAdjustStr.Index).Value = "Stock In Sales Return"
    '    '                .Cells(Me.ColDocumentNo.Index).Value = row_In_SR.Stock_In_Sales_Return_Number
    '    '                .Cells(Me.ColCodeNoAdjust.Index).Value = row_In_SR.Product_Code
    '    '                .Cells(Me.ColDescAdjust.Index).Value = row_In_SR.Description
    '    '                .Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(row_In_SR.Qty, row_In_SR._Decimal)
    '    '                .Cells(Me.ColUnitAdjust.Index).Value = row_In_SR.Unit_Code
    '    '                .Cells(Me.ColUnitPriceAdjust.Index).Value = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
    '    '                .Cells(Me.ColPrimaryCodeNoAdjust.Index).Value = row_In_SR.Primary_Product
    '    '                .Cells(Me.ColPrimaryUnitAdjust.Index).Value = row_In_SR.Primary_Units
    '    '                .Cells(Me.ColPrimaryStockInAdjust.Index).Value = row_In_SR.Primary
    '    '                .Cells(Me.ColJenisDokumen.Index).Value = ClsTransaksiMonthlyClosingItem.EnumTipeDokumenAcuan.StockInSalesReturnPeriodeSama
    '    '                counter1 = row_In_SR.Unit_Price

    '    '                If counter1 = 0 Then
    '    '                    .Cells(Me.ColRemarksAdjust.Index).Value = ""
    '    '                Else
    '    '                    .Cells(Me.ColUnitPriceAdjust.Index).Value = Me.ObjCD.DigitGrouping(counter1, DefPrec)
    '    '                    .Cells(Me.ColRemarksAdjust.Index).Value = "refer to unit price on " + _
    '    '                    Format(Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value.Date - TimeSpan.FromDays(1)), _
    '    '                    "MMM yyyy")
    '    '                End If
    '    '            End With

    '    '            Filter = "[primary transaksi asal] = " + row_In_SR.Primary.ToString + _
    '    '            " AND [primary code no] = " + row_In_SR.Primary_Product.ToString + _
    '    '            " AND [tipe transaksi asal] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR).ToString + _
    '    '            " AND [tipe item detail] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem).ToString
    '    '            Me.TBLItemDetail.DefaultView.RowFilter = Filter
    '    '            If Me.TBLItemDetail.DefaultView.Count = 0 Then
    '    '                Me.rowItemDetail = Me.TBLItemDetail.Newtbltrans_purchase_closing_item_detailRow
    '    '                With rowItemDetail
    '    '                    .amount = Me.ObjCD.DigitGrouping(row_In_SR.Unit_Price * row_In_SR.Qty, DefPrec)
    '    '                    .unit_price = Me.ObjCD.DigitGrouping(row_In_SR.Unit_Price, DefPrec)
    '    '                    .primary = -1
    '    '                    .primary_code_no = row_In_SR.Primary_Product
    '    '                    .primary_transaksi = -1
    '    '                    .primary_transaksi_asal = row_In_SR.Primary
    '    '                    .qty = row_In_SR.Qty
    '    '                    .tgl_transaksi_asal = row_In_SR.Stock_In_Date
    '    '                    .tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
    '    '                    .tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR
    '    '                    .return_price = 0
    '    '                    .primary_gain_account = -1
    '    '                    .primary_loss_account = -1
    '    '                    .primary_adjustment_type = -1
    '    '                    .remarks = ""
    '    '                End With
    '    '                Me.TBLItemDetail.Rows.Add(Me.rowItemDetail)
    '    '            Else
    '    '                Me.rowItemDetail = Me.TBLItemDetail.DefaultView(0).Row
    '    '                Me.TBL_In_SR.DefaultView.RowFilter = "Primary = " + row_In_SR.Primary.ToString + _
    '    '                " AND [Primary Product] = " + row_In_SR.Primary_Product.ToString
    '    '                AmountDummy = 0
    '    '                QtyDummy = 0
    '    '                For Each dv As DataRowView In Me.TBL_In_SR.DefaultView
    '    '                    row_In_SR3 = dv.Row
    '    '                    AmountDummy += Me.ObjCD.DigitGrouping(row_In_SR3.Unit_Price * row_In_SR3.Qty, DefPrec)
    '    '                    QtyDummy += row_In_SR3.Qty
    '    '                Next
    '    '                Me.rowItemDetail.qty = QtyDummy
    '    '                Me.rowItemDetail.amount = AmountDummy
    '    '                Me.rowItemDetail.unit_price = Me.ObjCD.DigitGrouping(Me.rowItemDetail.amount / Me.rowItemDetail.qty, DefPrec)
    '    '            End If
    '    '        Next

    '    '        'C.1. Sales return atas transaksi yang dibuat sebelum convertion date
    '    '        TotalAdjustment = DTA_In_SR1.Fill(TBL_In_SR1, EndDate, StartDate, ObjAccClosing.GetConvertionDate - TimeSpan.FromDays(1))
    '    '        counter = 0
    '    '        For Each row_In_SR1 In TBL_In_SR1.Rows
    '    '            'set unit price nol
    '    '            row_In_SR1.Unit_Price = 0
    '    '        Next
    '    '        For Each row_In_SR1 In TBL_In_SR1.Rows
    '    '            counter += 1
    '    '            Me.Text = OldText + "- Loading in sales return data (" + CStr(counter) _
    '    '            + " of " + TotalAdjustment.ToString + ")"
    '    '            Me.dgvItemAdjust.Rows.Add("")
    '    '            rowDGV = Me.dgvItemAdjust.Rows(Me.dgvItemAdjust.Rows.Count - 1)
    '    '            With rowDGV
    '    '                .Cells(Me.ColNoAdjust.Index).Value = Me.dgvItemAdjust.Rows.Count
    '    '                .Cells(Me.ColJenisDocumentAdjustStr.Index).Value = "Stock In Sales Return"
    '    '                .Cells(Me.ColDocumentNo.Index).Value = row_In_SR1.Stock_In_Sales_Return_Number
    '    '                .Cells(Me.ColCodeNoAdjust.Index).Value = row_In_SR1.Product_Code
    '    '                .Cells(Me.ColDescAdjust.Index).Value = row_In_SR1.Description
    '    '                .Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(row_In_SR1.Qty, row_In_SR1._Decimal)
    '    '                .Cells(Me.ColUnitAdjust.Index).Value = row_In_SR1.Unit_Code
    '    '                .Cells(Me.ColUnitPriceAdjust.Index).Value = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
    '    '                .Cells(Me.ColPrimaryCodeNoAdjust.Index).Value = row_In_SR1.Primary_Product
    '    '                .Cells(Me.ColPrimaryUnitAdjust.Index).Value = row_In_SR1.Primary_Unit
    '    '                .Cells(Me.ColPrimaryStockInAdjust.Index).Value = row_In_SR1.Primary_Stock_In_Out
    '    '                .Cells(Me.ColJenisDokumen.Index).Value = ClsTransaksiMonthlyClosingItem.EnumTipeDokumenAcuan.StockInSalesReturnPeriodeSama
    '    '                .Cells(Me.ColRemarksAdjust.Index).Value = ""
    '    '            End With

    '    '            Filter = "[primary transaksi asal] = " + row_In_SR1.Primary_Stock_In_Out.ToString + _
    '    '            " AND [primary code no] = " + row_In_SR1.Primary_Product.ToString + _
    '    '            " AND [tipe transaksi asal] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR).ToString + _
    '    '            " AND [tipe item detail] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem).ToString
    '    '            Me.TBLItemDetail.DefaultView.RowFilter = Filter
    '    '            If Me.TBLItemDetail.DefaultView.Count = 0 Then
    '    '                Me.rowItemDetail = Me.TBLItemDetail.Newtbltrans_purchase_closing_item_detailRow
    '    '                With rowItemDetail
    '    '                    .amount = Me.ObjCD.DigitGrouping(row_In_SR1.Unit_Price * row_In_SR1.Qty, DefPrec)
    '    '                    .unit_price = Me.ObjCD.DigitGrouping(row_In_SR1.Unit_Price, DefPrec)
    '    '                    .primary = -1
    '    '                    .primary_code_no = row_In_SR1.Primary_Product
    '    '                    .primary_transaksi = -1
    '    '                    .primary_transaksi_asal = row_In_SR1.Primary_Stock_In_Out
    '    '                    .qty = row_In_SR1.Qty
    '    '                    .tgl_transaksi_asal = row_In_SR1.Stock_In_Date
    '    '                    .tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
    '    '                    .tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR
    '    '                    .return_price = 0
    '    '                    .primary_gain_account = -1
    '    '                    .primary_loss_account = -1
    '    '                    .primary_adjustment_type = -1
    '    '                    .remarks = ""
    '    '                End With
    '    '                Me.TBLItemDetail.Rows.Add(Me.rowItemDetail)
    '    '            Else
    '    '                Me.rowItemDetail = Me.TBLItemDetail.DefaultView(0).Row
    '    '                Me.TBL_In_SR1.DefaultView.RowFilter = "[Primary Stock In Out] = " + row_In_SR1.Primary_Stock_In_Out.ToString + _
    '    '                " AND [Primary Product] = " + row_In_SR1.Primary_Product.ToString
    '    '                AmountDummy = 0
    '    '                QtyDummy = 0
    '    '                For Each dv As DataRowView In Me.TBL_In_SR1.DefaultView
    '    '                    row_In_SR2 = dv.Row
    '    '                    AmountDummy += Me.ObjCD.DigitGrouping(row_In_SR2.Unit_Price * row_In_SR2.Qty, DefPrec)
    '    '                    QtyDummy += row_In_SR2.Qty
    '    '                Next
    '    '                Me.rowItemDetail.qty = QtyDummy
    '    '                Me.rowItemDetail.amount = AmountDummy
    '    '                Me.rowItemDetail.unit_price = Me.ObjCD.DigitGrouping(Me.rowItemDetail.amount / Me.rowItemDetail.qty, DefPrec)
    '    '            End If
    '    '        Next

    '    '        'C.2. Quick Sales return 
    '    '        TBL_In_SR1.Rows.Clear()
    '    '        TotalAdjustment = DTA_In_SR1.FillBy_QuickSR(TBL_In_SR1, EndDate, StartDate)
    '    '        counter = 0
    '    '        For Each row_In_SR1 In TBL_In_SR1.Rows
    '    '            'set unit price nol
    '    '            row_In_SR1.Unit_Price = 0
    '    '        Next
    '    '        For Each row_In_SR1 In TBL_In_SR1.Rows
    '    '            counter += 1
    '    '            Me.Text = OldText + "- Loading in sales return data (" + CStr(counter) _
    '    '            + " of " + TotalAdjustment.ToString + ")"
    '    '            Me.dgvItemAdjust.Rows.Add("")
    '    '            rowDGV = Me.dgvItemAdjust.Rows(Me.dgvItemAdjust.Rows.Count - 1)
    '    '            With rowDGV
    '    '                .Cells(Me.ColNoAdjust.Index).Value = Me.dgvItemAdjust.Rows.Count
    '    '                .Cells(Me.ColJenisDocumentAdjustStr.Index).Value = "Stock In Sales Return"
    '    '                .Cells(Me.ColDocumentNo.Index).Value = row_In_SR1.Stock_In_Sales_Return_Number
    '    '                .Cells(Me.ColCodeNoAdjust.Index).Value = row_In_SR1.Product_Code
    '    '                .Cells(Me.ColDescAdjust.Index).Value = row_In_SR1.Description
    '    '                .Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(row_In_SR1.Qty, row_In_SR1._Decimal)
    '    '                .Cells(Me.ColUnitAdjust.Index).Value = row_In_SR1.Unit_Code
    '    '                .Cells(Me.ColUnitPriceAdjust.Index).Value = Me.ObjCD.DigitGrouping(0, Me.DefaultPrecision)
    '    '                .Cells(Me.ColPrimaryCodeNoAdjust.Index).Value = row_In_SR1.Primary_Product
    '    '                .Cells(Me.ColPrimaryUnitAdjust.Index).Value = row_In_SR1.Primary_Unit
    '    '                .Cells(Me.ColPrimaryStockInAdjust.Index).Value = row_In_SR1.Primary_Stock_In_Out
    '    '                .Cells(Me.ColJenisDokumen.Index).Value = ClsTransaksiMonthlyClosingItem.EnumTipeDokumenAcuan.StockInSalesReturnPeriodeSama
    '    '                .Cells(Me.ColRemarksAdjust.Index).Value = ""
    '    '            End With

    '    '            Filter = "[primary transaksi asal] = " + row_In_SR1.Primary_Stock_In_Out.ToString + _
    '    '            " AND [primary code no] = " + row_In_SR1.Primary_Product.ToString + _
    '    '            " AND [tipe transaksi asal] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR).ToString + _
    '    '            " AND [tipe item detail] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem).ToString
    '    '            Me.TBLItemDetail.DefaultView.RowFilter = Filter
    '    '            If Me.TBLItemDetail.DefaultView.Count = 0 Then
    '    '                Me.rowItemDetail = Me.TBLItemDetail.Newtbltrans_purchase_closing_item_detailRow
    '    '                With rowItemDetail
    '    '                    .amount = Me.ObjCD.DigitGrouping(row_In_SR1.Unit_Price * row_In_SR1.Qty, DefPrec)
    '    '                    .unit_price = Me.ObjCD.DigitGrouping(row_In_SR1.Unit_Price, DefPrec)
    '    '                    .primary = -1
    '    '                    .primary_code_no = row_In_SR1.Primary_Product
    '    '                    .primary_transaksi = -1
    '    '                    .primary_transaksi_asal = row_In_SR1.Primary_Stock_In_Out
    '    '                    .qty = row_In_SR1.Qty
    '    '                    .tgl_transaksi_asal = row_In_SR1.Stock_In_Date
    '    '                    .tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
    '    '                    .tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR
    '    '                    .return_price = 0
    '    '                    .primary_gain_account = -1
    '    '                    .primary_loss_account = -1
    '    '                    .primary_adjustment_type = -1
    '    '                    .remarks = ""
    '    '                End With
    '    '                Me.TBLItemDetail.Rows.Add(Me.rowItemDetail)
    '    '            Else
    '    '                Me.rowItemDetail = Me.TBLItemDetail.DefaultView(0).Row
    '    '                Me.TBL_In_SR1.DefaultView.RowFilter = "[Primary Stock In Out] = " + row_In_SR1.Primary_Stock_In_Out.ToString + _
    '    '                " AND [Primary Product] = " + row_In_SR1.Primary_Product.ToString
    '    '                AmountDummy = 0
    '    '                QtyDummy = 0
    '    '                For Each dv As DataRowView In Me.TBL_In_SR1.DefaultView
    '    '                    row_In_SR2 = dv.Row
    '    '                    AmountDummy += Me.ObjCD.DigitGrouping(row_In_SR2.Unit_Price * row_In_SR2.Qty, DefPrec)
    '    '                    QtyDummy += row_In_SR2.Qty
    '    '                Next
    '    '                Me.rowItemDetail.qty = QtyDummy
    '    '                Me.rowItemDetail.amount = AmountDummy
    '    '                Me.rowItemDetail.unit_price = Me.ObjCD.DigitGrouping(Me.rowItemDetail.amount / Me.rowItemDetail.qty, DefPrec)
    '    '            End If
    '    '        Next

    '    '        'D. Stock bundling tanpa batas tidak lagi berlaku
    '    '        'E. Sales return nota return
    '    '        'E.1. Sales return yang terjadi pada bulan ybs
    '    '        Me.TglDOMinimum = Date.MaxValue
    '    '        Me.DTA_SR2.Connection.ConnectionString = Core_Module.GetConstring
    '    '        Me.DTA_SR2_Dummy.Connection.ConnectionString = Core_Module.GetConstring
    '    '        Me.DTA_SR2_Cost.Connection.ConnectionString = Core_Module.GetConstring
    '    '        Me.DTA_SR2.Fill(Me.TBL_SR2, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value), _
    '    '        Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value), Me.ObjAccClosing.GetConvertionDate - TimeSpan.FromDays(1))

    '    '        'E.2. Hitung cost - load data cost invoice
    '    '        If Me.TBL_SR2.Rows.Count > 0 Then
    '    '            'tgl do minimum
    '    '            For Each Me.row_SR2 In Me.TBL_SR2.Rows
    '    '                If Me.row_SR2.Do_Date < Me.TglDOMinimum Then
    '    '                    Me.TglDOMinimum = Me.row_SR2.Do_Date
    '    '                End If
    '    '            Next
    '    '            Me.TglDOMinimum = Me.ObjCD.GetFirstDateofTheMonth(Me.TglDOMinimum)

    '    '            'load item cost invoice yang sudah terjadi
    '    '            Me.DTA_SR2_Cost.Fill(Me.TBL_SR2_Cost, CInt(Me.TipeSR_CostInvoice_TipeTranskasiAsal), _
    '    '            CInt(Me.TipeSR_CostInvoice_TipeItemDetail), Me.TglDOMinimum, Me.dtpPeriod.Value.Date)

    '    '            'E.3. Kurangkan data cost dengan return yang sudah terjadi sebelumnya, sehinga item invoice yang terjadi adalah setelah pengurangan
    '    '            'dan memastikan cost yang diambil tidak tumpang tindih, terutama jika FIFO/LIFO
    '    '            While Me.TglDOMinimum >= Me.ObjAccClosing.GetConvertionDate
    '    '                'load item return untuk bulan ybs
    '    '                Me.DTA_SR2_Dummy.Fill(Me.TBL_SR2_Dummy, Me.ObjCD.GetLasDateOfTheMonth(Me.TglDOMinimum), _
    '    '                Me.ObjCD.GetFirstDateofTheMonth(Me.TglDOMinimum), Me.ObjAccClosing.GetConvertionDate - TimeSpan.FromDays(1))

    '    '                'kurangkan item invoice dengan item return
    '    '                For Each Me.row_SR2_Dummy In Me.TBL_SR2_Dummy.Rows
    '    '                    Me.TBL_SR2_Cost.DefaultView.RowFilter = "[Primary Invoice] = " + Me.row_SR2_Dummy.Primary_Invoice_Sales.ToString + _
    '    '                    " AND [primary code no] = " + Me.row_SR2_Dummy.Primary_Product.ToString
    '    '                    While Me.row_SR2_Dummy.Qty > 0 And Me.TBL_SR2_Cost.DefaultView.Count > 0
    '    '                        If Me.TBL_SR2_Cost.DefaultView.Count > 0 Then
    '    '                            Me.row_SR2_Cost = Me.TBL_SR2_Cost.DefaultView(0).Row
    '    '                            If Me.row_SR2_Dummy.Qty >= Me.row_SR2_Cost.qty Then
    '    '                                Me.row_SR2_Dummy.Qty -= Me.row_SR2_Cost.qty
    '    '                                Me.TBL_SR2_Cost.Rows.Remove(Me.row_SR2_Cost)
    '    '                            Else
    '    '                                Me.row_SR2_Cost.qty -= Me.row_SR2_Dummy.Qty
    '    '                                AmountDummy = Math.Round((Me.row_SR2_Cost.amount * Me.row_SR2_Dummy.Qty) / Me.row_SR2_Cost.qty, _
    '    '                                DefPrec, MidpointRounding.AwayFromZero)
    '    '                                Me.row_SR2_Cost.amount -= AmountDummy
    '    '                                Me.row_SR2_Dummy.Qty = 0
    '    '                            End If
    '    '                        Else
    '    '                            Exit While
    '    '                        End If
    '    '                    End While
    '    '                Next

    '    '                'mundur 1 bulan
    '    '                Me.TglDOMinimum = Me.ObjCD.GetFirstDateofTheMonth(Me.TglDOMinimum - TimeSpan.FromDays(1))
    '    '            End While

    '    '            'E.4. Hitung cost aktual yang terjadi
    '    '            For Each Me.row_SR2 In Me.TBL_SR2.Rows
    '    '                Me.TBL_SR2_Cost.DefaultView.RowFilter = "[Primary Invoice] = " + Me.row_SR2.Primary_Invoice_Sales.ToString + _
    '    '                " AND [primary code no] = " + Me.row_SR2.Primary_Product.ToString
    '    '                QtySRDummy = Me.row_SR2.Qty
    '    '                QtyCost = 0
    '    '                For Each dv As DataRowView In Me.TBL_SR2_Cost.DefaultView
    '    '                    Me.row_SR2_Cost = dv.Row
    '    '                    QtyCost += Me.row_SR2_Cost.qty
    '    '                Next
    '    '                If QtySRDummy <= QtyCost Then
    '    '                    Me.row_SR2.Amount = 0
    '    '                    While QtySRDummy > 0
    '    '                        Me.row_SR2_Cost = Me.TBL_SR2_Cost.DefaultView(0).Row
    '    '                        If QtySRDummy >= Me.row_SR2_Cost.qty Then
    '    '                            Me.row_SR2.Amount += Me.row_SR2_Cost.amount
    '    '                            QtySRDummy -= Me.row_SR2_Cost.qty
    '    '                            Me.TBL_SR2_Cost.Rows.Remove(Me.row_SR2_Cost)
    '    '                        Else
    '    '                            AmountDummy = Math.Round((Me.row_SR2_Cost.amount * QtySRDummy) / Me.row_SR2_Cost.qty, DefPrec, MidpointRounding.AwayFromZero)
    '    '                            Me.row_SR2.Amount += AmountDummy
    '    '                            Me.row_SR2_Cost.amount -= AmountDummy
    '    '                            Me.row_SR2_Cost.qty -= QtySRDummy
    '    '                            QtySRDummy = 0
    '    '                        End If
    '    '                    End While
    '    '                Else
    '    '                    'kemungkinan lainnya adalah, return nota return pajak, atas invoice dibulan yang sama, sehingga datanya tidak ada
    '    '                    'Me.row_SR2.Primary_Stock_In_Sales_Return
    '    '                    If Me.ApakahAdaItemAdjustment(ClsTransaksiMonthlyClosingItem.EnumTipeDokumenAcuan.StockInSalesReturnPeriodeSama, Me.row_SR2.Primary_Stock_In_Sales_Return) = True Then
    '    '                        Me.row_SR2.Amount = 0
    '    '                        Me.row_SR2.Amount += Math.Round(Me.row_SR2.Qty * _
    '    '                        Me.GetCostPerUnit_Adjustment(ClsTransaksiMonthlyClosingItem.EnumTipeDokumenAcuan.StockInSalesReturnPeriodeSama, _
    '    '                        Me.row_SR2.Primary_Stock_In_Sales_Return), DefPrec, MidpointRounding.AwayFromZero)
    '    '                    Else
    '    '                        MsgBox("Logic error please contact your vendor.", MsgBoxStyle.Information, "Information")
    '    '                        Exit Sub
    '    '                    End If
    '    '                End If
    '    '            Next

    '    '            'E.5. Masukan ke dalam item detail
    '    '            For Each Me.row_SR2 In Me.TBL_SR2.Rows
    '    '                Filter = "[primary transaksi asal] = " + Me.row_SR2.Primary_Stock_In_Sales_Return.ToString + _
    '    '                " AND [primary code no] = " + row_SR2.Primary_Product.ToString + _
    '    '                " AND [tipe transaksi asal] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR).ToString + _
    '    '                " AND [tipe item detail] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem).ToString

    '    '                Me.TBLItemDetail.DefaultView.RowFilter = Filter
    '    '                If Me.TBLItemDetail.DefaultView.Count = 0 Then
    '    '                    Me.rowItemDetail = Me.TBLItemDetail.Newtbltrans_purchase_closing_item_detailRow
    '    '                    With rowItemDetail
    '    '                        .amount = Me.ObjCD.DigitGrouping(Me.row_SR2.Amount, DefPrec)
    '    '                        .unit_price = Me.ObjCD.DigitGrouping(Me.row_SR2.Amount / Me.row_SR2.Qty, DefPrec)
    '    '                        .primary = -1
    '    '                        .primary_code_no = row_SR2.Primary_Product
    '    '                        .primary_transaksi = -1
    '    '                        .primary_transaksi_asal = row_SR2.Primary_Stock_In_Sales_Return
    '    '                        .qty = row_SR2.Qty
    '    '                        .tgl_transaksi_asal = row_SR2.SR_Date
    '    '                        .tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
    '    '                        .tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR
    '    '                        .return_price = 0
    '    '                        .primary_gain_account = -1
    '    '                        .primary_loss_account = -1
    '    '                        .primary_adjustment_type = -1
    '    '                        .remarks = ""
    '    '                    End With
    '    '                    Me.TBLItemDetail.Rows.Add(Me.rowItemDetail)
    '    '                    DetailAmountDummy += rowItemDetail.amount
    '    '                Else
    '    '                    Me.rowItemDetail = Me.TBLItemDetail.DefaultView(0).Row
    '    '                    Me.TBL_SR2.DefaultView.RowFilter = "[Primary Stock In Sales Return] = " + row_SR2.Primary_Stock_In_Sales_Return.ToString + _
    '    '                    " AND [Primary Product] = " + row_SR2.Primary_Product.ToString
    '    '                    AmountDummy = 0
    '    '                    QtyDummy = 0
    '    '                    For Each dv As DataRowView In Me.TBL_SR2.DefaultView
    '    '                        Me.row_SR2 = dv.Row
    '    '                        AmountDummy += Me.ObjCD.DigitGrouping(Me.row_SR2.Amount, DefPrec)
    '    '                        QtyDummy += Me.row_SR2.Qty
    '    '                    Next
    '    '                    Me.rowItemDetail.qty = QtyDummy
    '    '                    Me.rowItemDetail.amount = AmountDummy
    '    '                    Me.rowItemDetail.unit_price = Me.ObjCD.DigitGrouping(Me.rowItemDetail.amount / Me.rowItemDetail.qty, DefPrec)
    '    '                    DetailAmountDummy += AmountDummy
    '    '                End If
    '    '            Next
    '    '        End If

    '    '        '-
    '    '        '2. load semua data purchase
    '    '        Me.DTAPurchase.Fill(Me.TBLPurchase, EndDate, _
    '    '        StartDate)
    '    '        DetailAmountDummy = 0
    '    '        '- adjust data qty, dikurangkan dengan item return revisi aop
    '    '        Me.DTAPR_Rev_AOP.Fill(Me.TBLPR_Rev_AOP, EndDate, StartDate)
    '    '        For Each Me.rowPR_Rev_AOP In Me.TBLPR_Rev_AOP.Rows
    '    '            Filter = "[Primary AOP] = " + Me.rowPR_Rev_AOP.Primary_AOP.ToString + _
    '    '            " AND [Primary Product] = " + Me.rowPR_Rev_AOP.Primary_Product.ToString + _
    '    '            " AND BatchNumber = '" + Me.rowPR_Rev_AOP.Batch_No + "'"

    '    '            Me.TBLPurchase.DefaultView.RowFilter = Filter
    '    '            If Me.TBLPurchase.DefaultView.Count = 0 Then
    '    '                MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
    '    '                Exit Sub
    '    '            Else
    '    '                Me.rowPurchase = Me.TBLPurchase.DefaultView(0).Row
    '    '                Me.rowPurchase.Total_Qty -= Me.rowPR_Rev_AOP.Total_Qty
    '    '            End If

    '    '            If Me.rowPurchase.Total_Qty = 0 Then
    '    '                Me.TBLPurchase.Rows.Remove(Me.rowPurchase)
    '    '            ElseIf Me.rowPurchase.Total_Qty < 0 Then
    '    '                MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
    '    '                Exit Sub
    '    '            End If
    '    '        Next
    '    '        For Each Me.rowPurchase In Me.TBLPurchase.Rows
    '    '            Filter = "[primary transaksi asal] = " + rowPurchase.Primary_AOP.ToString + _
    '    '           " AND [primary code no] = " + rowPurchase.Primary_Product.ToString + _
    '    '           " AND [tipe transaksi asal] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InPurchase).ToString + _
    '    '           " AND [tipe item detail] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem).ToString
    '    '            Me.TBLItemDetail.DefaultView.RowFilter = Filter
    '    '            If Me.TBLItemDetail.DefaultView.Count = 0 Then
    '    '                Me.rowItemDetail = Me.TBLItemDetail.Newtbltrans_purchase_closing_item_detailRow
    '    '                With rowItemDetail
    '    '                    .amount = Me.ObjCD.DigitGrouping(Me.rowPurchase.At_Amount, DefPrec)
    '    '                    .unit_price = Me.ObjCD.DigitGrouping(Me.rowPurchase.At_Price, DefPrec)
    '    '                    .primary = -1
    '    '                    .primary_code_no = rowPurchase.Primary_Product
    '    '                    .primary_transaksi = -1
    '    '                    .primary_transaksi_asal = rowPurchase.Primary_AOP
    '    '                    .qty = rowPurchase.Total_Qty
    '    '                    .tgl_transaksi_asal = rowPurchase.Tanggal_Transaksi
    '    '                    .tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
    '    '                    .tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InPurchase
    '    '                    .return_price = 0
    '    '                    .primary_gain_account = -1
    '    '                    .primary_loss_account = -1
    '    '                    .primary_adjustment_type = -1
    '    '                    .remarks = ""
    '    '                End With
    '    '                Me.TBLItemDetail.Rows.Add(Me.rowItemDetail)
    '    '                DetailAmountDummy += rowItemDetail.amount
    '    '            Else
    '    '                Me.rowItemDetail = Me.TBLItemDetail.DefaultView(0).Row
    '    '                Me.TBLPurchase.DefaultView.RowFilter = "[Primary AOP] = " + rowPurchase.Primary_AOP.ToString + _
    '    '                " AND [Primary Product] = " + rowPurchase.Primary_Product.ToString
    '    '                AmountDummy = 0
    '    '                QtyDummy = 0
    '    '                For Each dv As DataRowView In Me.TBLPurchase.DefaultView
    '    '                    Me.rowPurchase = dv.Row
    '    '                    AmountDummy += Me.ObjCD.DigitGrouping(Me.rowPurchase.At_Amount, DefPrec)
    '    '                    QtyDummy += Me.rowPurchase.Total_Qty
    '    '                Next
    '    '                Me.rowItemDetail.qty = QtyDummy
    '    '                Me.rowItemDetail.amount = AmountDummy
    '    '                Me.rowItemDetail.unit_price = Me.ObjCD.DigitGrouping(Me.rowItemDetail.amount / Me.rowItemDetail.qty, DefPrec)
    '    '                DetailAmountDummy += AmountDummy
    '    '            End If
    '    '        Next

    '    '        '3. load semua data out: sales, stock taking, purchase return, adjust --> mulai pada bulan convertion date
    '    '        '3.a sales, kurangkan qty dengan stock sales return rev invoice
    '    '        Me.DTASales.Fill(Me.TBLSales, EndDate, StartDate)
    '    '        Me.DTASalesReturn.Fill(Me.TBLSalesReturn, EndDate, StartDate)
    '    '        For Each Me.rowSalesReturn In Me.TBLSalesReturn.Rows
    '    '            Filter = "[Primary Product] = " + Me.rowSalesReturn.Primary_Product.ToString + _
    '    '            " AND [Batch Number] = '" + Me.rowSalesReturn.Batch_No + "' AND [Primary Transaksi] = " + Me.rowSalesReturn.Primary_Sales_Invoice.ToString

    '    '            Me.TBLSales.DefaultView.RowFilter = Filter
    '    '            If Me.TBLSales.DefaultView.Count = 0 Then
    '    '                MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
    '    '                Exit Sub
    '    '            Else
    '    '                Me.rowSales = Me.TBLSales.DefaultView(0).Row
    '    '                Me.rowSales.Total_Qty -= Me.rowSalesReturn.Total_Qty
    '    '            End If

    '    '            If Me.rowSales.Total_Qty = 0 Then
    '    '                Me.TBLSales.Rows.Remove(Me.rowSales)
    '    '            ElseIf Me.rowSales.Total_Qty < 0 Then
    '    '                MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
    '    '                Exit Sub
    '    '            End If
    '    '        Next
    '    '        For Each Me.rowSales In Me.TBLSales.Rows
    '    '            Filter = "[primary transaksi asal] = " + rowSales.Primary_Transaksi.ToString + _
    '    '            " AND [primary code no] = " + rowSales.Primary_Product.ToString + _
    '    '            " AND [tipe transaksi asal] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutSales).ToString + _
    '    '            " AND [tipe item detail] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem).ToString
    '    '            Me.TBLItemDetail.DefaultView.RowFilter = Filter
    '    '            If Me.TBLItemDetail.DefaultView.Count = 0 Then
    '    '                Me.rowItemDetail = Me.TBLItemDetail.Newtbltrans_purchase_closing_item_detailRow
    '    '                With rowItemDetail
    '    '                    .amount = 0
    '    '                    .unit_price = 0
    '    '                    .primary = -1
    '    '                    .primary_code_no = rowSales.Primary_Product
    '    '                    .primary_transaksi = -1
    '    '                    .primary_transaksi_asal = rowSales.Primary_Transaksi
    '    '                    .qty = rowSales.Total_Qty
    '    '                    .tgl_transaksi_asal = rowSales.Tanggal_Transaksi
    '    '                    .tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
    '    '                    .tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutSales
    '    '                    .return_price = 0
    '    '                    .primary_gain_account = -1
    '    '                    .primary_loss_account = -1
    '    '                    .primary_adjustment_type = -1
    '    '                    .remarks = ""
    '    '                End With
    '    '                Me.TBLItemDetail.Rows.Add(Me.rowItemDetail)
    '    '            Else
    '    '                Me.rowItemDetail = Me.TBLItemDetail.DefaultView(0).Row
    '    '                Me.TBLSales.DefaultView.RowFilter = "[Primary Transaksi] = " + rowSales.Primary_Transaksi.ToString + _
    '    '                " AND [Primary Product] = " + rowSales.Primary_Product.ToString
    '    '                QtyDummy = 0
    '    '                For Each dv As DataRowView In Me.TBLSales.DefaultView
    '    '                    Me.rowSales = dv.Row
    '    '                    QtyDummy += Me.rowSales.Total_Qty
    '    '                Next
    '    '                Me.rowItemDetail.qty = QtyDummy
    '    '                Me.rowItemDetail.unit_price = 0
    '    '            End If
    '    '        Next
    '    '        '3.b stock taking
    '    '        Me.DTAOutTaking.Fill(Me.TBLOutTaking, EndDate, StartDate)
    '    '        For Each Me.rowOutTaking In Me.TBLOutTaking.Rows
    '    '            'karena dari query sudah di group per product maka langsung tambahkan
    '    '            Me.rowItemDetail = Me.TBLItemDetail.Newtbltrans_purchase_closing_item_detailRow
    '    '            With rowItemDetail
    '    '                .amount = 0
    '    '                .unit_price = 0
    '    '                .primary = -1
    '    '                .primary_code_no = rowOutTaking.Primary_Product
    '    '                .primary_transaksi = -1
    '    '                .primary_transaksi_asal = rowOutTaking.Primary_Stock_Taking
    '    '                .qty = Math.Abs(rowOutTaking.Qty)
    '    '                .tgl_transaksi_asal = rowOutTaking.Start_Date
    '    '                .tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
    '    '                .tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutStockTakeAdjust
    '    '                .return_price = 0
    '    '                .primary_gain_account = -1
    '    '                .primary_loss_account = -1
    '    '                .primary_adjustment_type = -1
    '    '                .remarks = "stock taking"
    '    '            End With
    '    '            Me.TBLItemDetail.Rows.Add(Me.rowItemDetail)
    '    '        Next
    '    '        '3.c purchase return
    '    '        Me.DTA_PR.Fill(Me.TBL_PR, EndDate, TipePR, StartDate)
    '    '        For Each Me.row_PR In Me.TBL_PR.Rows
    '    '            Me.rowItemDetail = Me.TBLItemDetail.Newtbltrans_purchase_closing_item_detailRow
    '    '            With rowItemDetail
    '    '                .amount = 0
    '    '                .unit_price = 0
    '    '                .primary = -1
    '    '                .primary_code_no = row_PR.Primary_Product
    '    '                .primary_transaksi = -1
    '    '                .primary_transaksi_asal = row_PR.Primary_Transaksi
    '    '                .qty = row_PR.Total_Qty
    '    '                .tgl_transaksi_asal = row_PR.Tanggal_Transaksi
    '    '                .tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
    '    '                .tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutPR
    '    '                .return_price = 0
    '    '                .primary_gain_account = -1
    '    '                .primary_loss_account = -1
    '    '                .primary_adjustment_type = -1
    '    '                .remarks = ""
    '    '            End With
    '    '            Me.TBLItemDetail.Rows.Add(Me.rowItemDetail)
    '    '        Next
    '    '        '3.d out adjustment
    '    '        Me.DTA_OutAdjust.Fill(Me.TBL_OutAdjust, TipeStockOut, EndDate, StartDate)
    '    '        For Each Me.row_OutAdjust In Me.TBL_OutAdjust.Rows
    '    '            Me.rowItemDetail = Me.TBLItemDetail.Newtbltrans_purchase_closing_item_detailRow
    '    '            With rowItemDetail
    '    '                .amount = 0
    '    '                .unit_price = 0
    '    '                .primary = -1
    '    '                .primary_code_no = row_OutAdjust.Primary_Product
    '    '                .primary_transaksi = -1
    '    '                .primary_transaksi_asal = row_OutAdjust.Primary_Transaksi
    '    '                .qty = row_OutAdjust.Qty
    '    '                .tgl_transaksi_asal = row_OutAdjust.Tanggal_Transaksi
    '    '                .tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
    '    '                .tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutAdjust
    '    '                .return_price = 0
    '    '                .primary_gain_account = row_OutAdjust.Primary_Loss_Account
    '    '                .primary_loss_account = row_OutAdjust.Primary_Gain_Account
    '    '                .primary_adjustment_type = row_OutAdjust.PrimaryAdjType
    '    '                row_OutAdjust.Remarks = row_OutAdjust.Remarks.Trim
    '    '                If row_OutAdjust.Remarks <> "" Then
    '    '                    InterimString = " : "
    '    '                Else
    '    '                    InterimString = ""
    '    '                End If
    '    '                If row_OutAdjust.Adjustment_Type_Code <> "" Then
    '    '                    If (row_OutAdjust.Adjustment_Type_Desc + InterimString + row_OutAdjust.Remarks).Length > 500 Then
    '    '                        .remarks = (row_OutAdjust.Adjustment_Type_Desc + InterimString + row_OutAdjust.Remarks).Substring(0, 500)
    '    '                    Else
    '    '                        .remarks = row_OutAdjust.Adjustment_Type_Desc + InterimString + row_OutAdjust.Remarks
    '    '                    End If
    '    '                Else
    '    '                    If (row_OutAdjust.Remarks).Length > 500 Then
    '    '                        .remarks = (row_OutAdjust.Remarks).Substring(0, 500)
    '    '                    Else
    '    '                        .remarks = row_OutAdjust.Remarks
    '    '                    End If
    '    '                End If
    '    '            End With
    '    '            Me.TBLItemDetail.Rows.Add(Me.rowItemDetail)
    '    '        Next
    '    '        '4. load semua data bundling:
    '    '        '   A. Hubungan Before Item
    '    '        Me.DTABundling1.Fill(Me.TBLBundling1, EndDate, StartDate)
    '    '        '   B. Item Before
    '    '        Me.DTABefore.Fill(Me.TBLBefore, EndDate, StartDate)
    '    '        For Each Me.rowBefore In Me.TBLBefore.Rows
    '    '            Me.rowItemDetail = Me.TBLItemDetail.Newtbltrans_purchase_closing_item_detailRow
    '    '            With rowItemDetail
    '    '                .amount = 0
    '    '                .unit_price = 0
    '    '                .primary = -1
    '    '                .primary_code_no = rowBefore.Primary_Product
    '    '                .primary_transaksi = -1
    '    '                .primary_transaksi_asal = rowBefore.Primary_Transaksi
    '    '                .qty = rowBefore.Qty
    '    '                .tgl_transaksi_asal = rowBefore.Tanggal_Transaksi
    '    '                .tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
    '    '                .tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling
    '    '                .return_price = 0
    '    '                .primary_gain_account = -1
    '    '                .primary_loss_account = -1
    '    '                .primary_adjustment_type = -1
    '    '                .remarks = ""
    '    '            End With
    '    '            Me.TBLItemDetail.Rows.Add(Me.rowItemDetail)
    '    '        Next
    '    '        '   C. Item After
    '    '        Me.DTAAfter.Fill(Me.TBLAfter, EndDate, StartDate)
    '    '        'GabungkanBarisTBLAfterBundling(Me.TBLAfter)
    '    '        For Each Me.rowAfter In Me.TBLAfter.Rows
    '    '            Filter = "[primary transaksi asal] = " + rowAfter.Primary_Transaksi.ToString + _
    '    '           " AND [primary code no] = " + rowAfter.Primary_Product.ToString + _
    '    '           " AND [tipe transaksi asal] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling).ToString + _
    '    '           " AND [tipe item detail] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem).ToString
    '    '            Me.TBLItemDetail.DefaultView.RowFilter = Filter
    '    '            If Me.TBLItemDetail.DefaultView.Count = 0 Then
    '    '                Me.rowItemDetail = Me.TBLItemDetail.Newtbltrans_purchase_closing_item_detailRow
    '    '                With rowItemDetail
    '    '                    .amount = 0
    '    '                    .unit_price = 0
    '    '                    .primary = -1
    '    '                    .primary_code_no = rowAfter.Primary_Product
    '    '                    .primary_transaksi = -1
    '    '                    .primary_transaksi_asal = rowAfter.Primary_Transaksi
    '    '                    .qty = rowAfter.Qty
    '    '                    .tgl_transaksi_asal = rowAfter.Tanggal_Transaksi
    '    '                    .tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
    '    '                    .tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling
    '    '                    .return_price = 0
    '    '                    .primary_gain_account = -1
    '    '                    .primary_loss_account = -1
    '    '                    .primary_adjustment_type = -1
    '    '                    .remarks = ""
    '    '                End With
    '    '                Me.TBLItemDetail.Rows.Add(Me.rowItemDetail)
    '    '            Else
    '    '                Me.rowItemDetail = Me.TBLItemDetail.DefaultView(0).Row
    '    '                Me.TBLAfter.DefaultView.RowFilter = "[Primary Transaksi] = " + rowAfter.Primary_Transaksi.ToString + _
    '    '                " AND [Primary Product] = " + rowAfter.Primary_Product.ToString
    '    '                QtyDummy = 0
    '    '                For Each dv As DataRowView In Me.TBLAfter.DefaultView
    '    '                    Me.rowAfter = dv.Row
    '    '                    QtyDummy += Me.rowAfter.Qty
    '    '                Next
    '    '                Me.rowItemDetail.qty = QtyDummy
    '    '                Me.rowItemDetail.unit_price = 0
    '    '            End If
    '    '        Next

    '    '        '5. Hitung cost untuk semua product item A, mulai dr level 1.
    '    '        Lvl = 0
    '    '        ListProductReject.Clear()
    '    '        ListProduct.Clear()
    '    '        While Me.GetProductBundling(Lvl, ListProduct, Me.TBLBundling1, ListProductReject) = True
    '    '            'hitung unit price
    '    '            For Each PrimaryCodeNo As Double In ListProduct
    '    '                Me.HitungCost(PrimaryCodeNo, DefPrec, CostingMethod)
    '    '            Next
    '    '            'masukkan ke dalam list primary product reject
    '    '            For Each PrimaryCodeNo As Double In ListProduct
    '    '                ListProductReject.Add(PrimaryCodeNo)
    '    '            Next
    '    '            'tambahkan level
    '    '            Lvl += 1
    '    '            ListProduct.Clear()
    '    '        End While
    '    '    End If


    '    '    '6. Hitung semua cost untuk item yang lain, yang sudah dihitung, tidak perlu lagi
    '    '    For Each Me.rowProduct In Me.TBLProduct.Rows 'saring yang sudah pernah dihitung
    '    '        If ListProductReject.Contains(Me.rowProduct.Primary) = False Then
    '    '            ListProduct.Add(Me.rowProduct.Primary)
    '    '        End If
    '    '    Next
    '    '    'hitung unit price
    '    '    For Each PrimaryCodeNo As Double In ListProduct
    '    '        Me.HitungCost(PrimaryCodeNo, DefPrec, CostingMethod)
    '    '    Next


    '    '    'cek kembali
    '    '    Me.UpdateKeDGVItemUtama()
    '    '    'Me.CekKonsistensi() --> dicancel karena sudah diblok di button ok

    '    '    'buat jurnal
    '    '    Me.UpdateAkunDGV()
    '    '    If (Me.ObjClosing.GetJml = 0 And Me.baru = True) Or (Me.ObjClosing.GetJml = 1 And Me.baru = False) Then
    '    '    Else
    '    '        'jurnal hanya dibuat selain untuk begining balance
    '    '        Me.BuatJurnal(False)
    '    '    End If
    '    '    Me.dgvItemAdjust.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
    '    'Catch ex As Exception
    '    '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    '    MsgBox(counter.ToString)
    '    'Finally
    '    '    Windows.Forms.Cursor.Current = Cursors.Arrow
    '    '    ObjAccClosing = Nothing
    '    '    ListProduct = Nothing
    '    '    ListProductReject = Nothing
    '    'End Try
    'End Sub

    'mengisikan item detail, persiapan untuk perhitungan
    Private Sub LoadItemDetail2()
        Dim StartDate, EndDate As Date
        Dim Filter As String = ""
        'Dim OldText As String = "Monthly Inventory Closing"
        Dim TotalAdjustment As Double = 0
        Dim cmd As String = ""
        Dim counter As Double = 0
        Dim DefPrec As Integer = 0
        Dim counter1 As Double = 0
        Dim TipePR As ClsTransaksiPurchaseReturn.EnumTipe_PurchaseReturn = ClsTransaksiPurchaseReturn.EnumTipe_PurchaseReturn.Nota_Retur
        Dim TipeStockOut As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust
        Dim ListProduct As New List(Of Double)
        Dim ListProductReject As New List(Of Double)
        Dim Lvl As Integer = 0
        Dim TipeStockInOut As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust
        Dim PrimaryClosingSebelum As Double = -1
        Dim TipeItem As ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.EndingBalanceItem
        Dim PrimaryStockTaking As Double = -1
        Dim TglTaking As Date = #1/1/1950#
        Dim IndexDGV As Double = 0
        Dim AmountDummy As Decimal = 0
        Dim QtyDummy As Decimal = 0
        Dim DetailAmountDummy As Decimal = 0
        Dim ObjAccClosing As New AccModule.ClsTransaksiClosing
        Dim QtySRDummy As Decimal = 0
        Dim QtyCost As Decimal = 0
        Dim InterimString As String = ":"
        Dim DTAI As New DseMonthlyClosingNewTableAdapters.tbltrans_purchase_closing_item_detailTableAdapter
        Dim tblI As New DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailDataTable
        Dim rowI As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailRow
        Dim constring As String = ""
        Dim ApakahHitungUlang As Boolean = True
        Dim DTABefore As New DseTransaksiMonthlyClosing2TableAdapters.TblBundlingBeforeTableAdapter
        Dim TBLBefore As New DseTransaksiMonthlyClosing2.TblBundlingBeforeDataTable
        Dim rowBefore As DseTransaksiMonthlyClosing2.TblBundlingBeforeRow
        Dim DTAAfter As New DseTransaksiMonthlyClosing2TableAdapters.TblBundlingAfterTableAdapter
        Dim TBLAfter As New DseTransaksiMonthlyClosing2.TblBundlingAfterDataTable
        Dim rowAfter As DseTransaksiMonthlyClosing2.TblBundlingAfterRow
        Dim ClosingPertama As Boolean = False
        Dim ObjIdentifier As New CustomSMESS.ClsIdentifier
        Dim ObjStock As New Inventory_Module.FrmLaporanStockDetail2
        Dim ObjCustomIdentifier As New CustomSMESS.ClsIdentifier
        Try
            Me.TblProductQty.Rows.Clear()
            ObjStock.Scope_Index = Inventory_Module.FrmLaporanStockDetail2.EnumScope.Both
            ObjStock.LoadForMonthlyClosing = True
            If ObjCustomIdentifier.GetIdentifier <> "smess_lmb" Then
                ObjStock.FromCutOff_BeginingBalance(Me.TblProductQty, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value.Date))
            End If

            ClosingPertama = Me.ObjClosing.ApakahClosingPalingAwal(Me.TxtPrimaryClosing.Text)
            constring = Core_Module.GetConstring
            DTAI.Connection.ConnectionString = constring
            Me.TglDOMinimum = Date.MaxValue
            DefPrec = Me.objCurrency.GetDecimal_DefaultCurrency

            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            StartDate = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value)
            EndDate = Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value)


            'If Me.ApakahAdaPerubahanAdj(Me.TblA, Me.TblA2) = True Then
            'reset data untuk persiapan perhitungan
            Me.dgvItem2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem2.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.dgvItem2.Rows.Clear()

            Me.TBLItemDetail.Rows.Clear()
            Me.TBL_In_SR.Rows.Clear()
            Me.TBL_OutAdjust.Rows.Clear()
            Me.TBLInAdjust.Rows.Clear()
            Me.TBLTaking1.Rows.Clear()
            Me.TBLTaking2.Rows.Clear()
            Me.TBLPurchase.Rows.Clear()
            Me.TBLPR_Rev_AOP.Rows.Clear()
            Me.TBL_HistoricalPrice.Rows.Clear()
            Me.TBLSales.Rows.Clear()
            Me.TBLSalesReturn.Rows.Clear()
            Me.TBLOutTaking.Rows.Clear()
            Me.TBL_PR.Rows.Clear()
            Me.TBL_OutAdjust.Rows.Clear()
            Me.TBLBundling1.Rows.Clear()
            TBLBefore.Rows.Clear()
            TBLAfter.Rows.Clear()
            Me.TBLBeginingBalance.Rows.Clear()
            Me.TBLBeginingBalance1.Rows.Clear()
            Me.TBLItemDetail.Rows.Clear()
            tblI.Rows.Clear()

            '0. A. load data product
            Me.TBLProduct.Rows.Clear()
            Me.DTAProduct.Fill(Me.TBLProduct)
            If Me.baru = False Then
                ApakahHitungUlang = False
                DTAI.FillBy(tblI, Me.TxtPrimaryClosing.Text)
                For Each rowI In tblI.Rows
                    Me.rowItemDetail = Me.TBLItemDetail.Newtbltrans_purchase_closing_item_detailRow
                    With rowItemDetail
                        .primary = rowI.primary
                        .primary_code_no = rowI.primary_code_no
                        .primary_transaksi = Me.TxtPrimaryClosing.Text
                        .primary_transaksi_asal = rowI.primary_transaksi_asal
                        .qty = rowI.qty
                        .tgl_transaksi_asal = rowI.tgl_transaksi_asal
                        .tipe_item_detail = rowI.tipe_item_detail
                        .tipe_transaksi_asal = rowI.tipe_transaksi_asal
                        .unit_price = Me.ObjCD.DigitGrouping(rowI.unit_price, DefPrec)
                        .return_price = rowI.unit_price
                        .primary_gain_account = rowI.primary_gain_account
                        .primary_loss_account = rowI.primary_loss_account
                        .primary_adjustment_type = rowI.primary_adjustment_type
                        .remarks = rowI.remarks
                        .amount = Me.ObjCD.DigitGrouping(rowI.amount, DefPrec)
                        .Cost_Perct_Bundling = rowI.Cost_Perct_Bundling
                        .Overhead = rowI.Overhead
                    End With
                    Me.TBLItemDetail.Rows.Add(Me.rowItemDetail)
                Next
            Else
                '0. B. Load Data beginning balance
                If ClosingPertama = True Then
                    '(i). Jika closing pertama maka ambil datanya dari begining balance(stock in adjust begining balance)
                    'tidak perlu isikan disini, sudah di ambil alih fungsi isi item adjustment
                Else
                    '(ii). jika closing selanjutnya, maka ambil begining balance dari ending balance transaksi sebelumnya
                    PrimaryClosingSebelum = Me.ObjClosing.GetPrimary(Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value - TimeSpan.FromDays(5)))
                    Me.DTABeginingBalance1.Fill(Me.TBLBeginingBalance1, PrimaryClosingSebelum, TipeItem)
                    For Each Me.rowBeginingBalance1 In Me.TBLBeginingBalance1.Rows
                        Me.rowItemDetail = Me.TBLItemDetail.Newtbltrans_purchase_closing_item_detailRow
                        With rowItemDetail
                            .amount = Me.rowBeginingBalance1.amount
                            .primary = -1
                            .primary_code_no = Me.rowBeginingBalance1.primary_code_no
                            .primary_transaksi = -1
                            .primary_transaksi_asal = Me.rowBeginingBalance1.primary_transaksi_asal
                            .qty = Me.rowBeginingBalance1.qty
                            .tgl_transaksi_asal = Me.rowBeginingBalance1.tgl_transaksi_asal
                            .tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.BeginingBalanceItem
                            .tipe_transaksi_asal = Me.rowBeginingBalance1.tipe_transaksi_asal
                            .unit_price = Me.ObjCD.DigitGrouping(Me.rowBeginingBalance1.unit_price, DefPrec)
                            .return_price = 0
                            .primary_gain_account = -1
                            .primary_loss_account = -1
                            .primary_adjustment_type = -1
                            .remarks = ""
                            .Cost_Perct_Bundling = 0
                            .Overhead = 0
                        End With
                        Me.TBLItemDetail.Rows.Add(Me.rowItemDetail)
                    Next
                End If

                '1. load data adjustment, stock taking, sales return
                DetailAmountDummy = 0
                Me.Calc_CurrentStep += 1
                Me.ProgressBar1.Value = 0
                Me.LblProgress.Text = ""
                Me.ProgressBar1.Maximum = Me.TblA.Rows.Count
                Me.LblStep.Text = "Step " + Me.Calc_CurrentStep.ToString + " of " + Me.Calc_TotalStep.ToString + " : collecting in adjustment data"
                'Me.LblProgress.Visible = True

                For Each Me.rowA In Me.TblA.Rows
                    Me.ProgressBar1.Value += 1

                    If Math.IEEERemainder(Me.ProgressBar1.Value, 50) = 0 Then
                        Me.LblProgress.Text = "Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString
                    End If

                    'rowA.Amount = Math.Round(rowA.Unit_Price * rowA.Qty_In_Adjust, DefPrec, MidpointRounding.AwayFromZero)
                    If rowA.Tipe_Transaksi_Asal = CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust) Then
                        'A. Adjust
                        Filter = "[primary transaksi asal] = " + rowA.Primary_Documen_Acuan.ToString + _
                                 " AND [primary code no] = " + rowA.Primary_Code_No.ToString + _
                                 " AND [tipe transaksi asal] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust).ToString + _
                                 " AND [tipe item detail] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem).ToString
                        Me.TBLItemDetail.DefaultView.RowFilter = Nothing
                        Me.TBLItemDetail.DefaultView.RowFilter = Filter
                        If Me.TBLItemDetail.DefaultView.Count = 0 Then
                            Me.rowItemDetail = Me.TBLItemDetail.Newtbltrans_purchase_closing_item_detailRow
                            With rowItemDetail
                                If ClosingPertama = True Then
                                    .amount = rowA.Amount
                                    .unit_price = Me.ObjCD.DigitGrouping(rowA.Unit_Price, DefPrec)
                                Else
                                    If ObjIdentifier.GetIdentifier = "smess_lmb" And _
                                    Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value) = #12/1/2017# And _
                                    rowA.Document_No = "RESET2017-5" Then
                                        If rowA.Amount <> 0 Then
                                            .amount = rowA.Amount
                                            .unit_price = Me.ObjCD.DigitGrouping(rowA.Unit_Price, DefPrec)
                                        Else
                                            .amount = Me.ObjCD.DigitGrouping(rowA.Unit_Price * rowA.Qty_In_Adjust, DefPrec)
                                            .unit_price = Me.ObjCD.DigitGrouping(rowA.Unit_Price, DefPrec)
                                        End If
                                    Else
                                        .amount = Me.ObjCD.DigitGrouping(rowA.Unit_Price * rowA.Qty_In_Adjust, DefPrec)
                                        .unit_price = Me.ObjCD.DigitGrouping(rowA.Unit_Price, DefPrec)
                                    End If
                                End If
                                .primary = -1
                                .primary_code_no = rowA.Primary_Code_No
                                .primary_transaksi = -1
                                .primary_transaksi_asal = rowA.Primary_Documen_Acuan
                                .qty = rowA.Qty_In_Adjust
                                .tgl_transaksi_asal = rowA.Tanggal_Documen_Acuan
                                .tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
                                .tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust
                                .return_price = 0
                                .primary_gain_account = rowA.Primary_Loss_Account
                                .primary_loss_account = rowA.Primary_Gain_Account
                                .primary_adjustment_type = rowA.Primary_Adjustment_type
                                .Cost_Perct_Bundling = 0
                                .Overhead = 0
                                rowA.Remarks = rowA.Remarks.Trim
                                If rowA.Remarks <> "" Then
                                    InterimString = " : "
                                Else
                                    InterimString = ""
                                End If
                                If rowA.Adjust_Type_Desc <> "" Then
                                    If (rowA.Adjust_Type_Desc + InterimString + rowA.remarks_adjust).Length > 250 Then
                                        .remarks = (rowA.Adjust_Type_Desc + InterimString + rowA.remarks_adjust).Substring(0, 250)
                                    Else
                                        .remarks = rowA.Adjust_Type_Desc + InterimString + rowA.remarks_adjust
                                    End If
                                Else
                                    If (rowA.Remarks).Length > 250 Then
                                        .remarks = (rowA.remarks_adjust).Substring(0, 250)
                                    Else
                                        .remarks = rowA.remarks_adjust
                                    End If
                                End If
                                If .remarks.Length > 250 Then
                                    .remarks = .remarks.Substring(0, 250)
                                End If
                            End With
                            Me.TBLItemDetail.Rows.Add(Me.rowItemDetail)
                        Else
                            Me.rowItemDetail = Me.TBLItemDetail.DefaultView(0).Row

                            If ClosingPertama = True Then
                                Me.rowItemDetail.amount += rowA.Amount
                            Else
                                'Me.rowItemDetail.amount += Math.Round(rowA.Unit_Price * rowA.Qty_In_Adjust, DefPrec, MidpointRounding.AwayFromZero)

                                If ObjIdentifier.GetIdentifier = "smess_lmb" And _
                                   Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value) = #12/1/2017# And _
                                   rowA.Document_No = "RESET2017-5" Then
                                    If rowA.Amount <> 0 Then
                                        Me.rowItemDetail.amount += rowA.Amount
                                    Else
                                        Me.rowItemDetail.amount += Me.ObjCD.DigitGrouping(rowA.Unit_Price * rowA.Qty_In_Adjust, DefPrec)
                                    End If
                                Else
                                    Me.rowItemDetail.amount += Me.ObjCD.DigitGrouping(rowA.Unit_Price * rowA.Qty_In_Adjust, DefPrec)
                                End If
                            End If
                            Me.rowItemDetail.qty += rowA.Qty_In_Adjust
                            Me.rowItemDetail.unit_price = Me.ObjCD.DigitGrouping(Me.rowItemDetail.amount / Me.rowItemDetail.qty, DefPrec)
                            If (Me.rowItemDetail.remarks + " - " + rowA.Remarks).Length > 250 Then
                                Me.rowItemDetail.remarks = (Me.rowItemDetail.remarks + " - " + rowA.Remarks).Substring(0, 250)
                            Else
                                Me.rowItemDetail.remarks = Me.rowItemDetail.remarks + " - " + rowA.Remarks
                            End If

                            If Me.rowItemDetail.remarks.Length > 250 Then
                                Me.rowItemDetail.remarks = Me.rowItemDetail.remarks.Substring(0, 250)
                            End If
                        End If
                    ElseIf rowA.Tipe_Transaksi_Asal = CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InStockTakeAdjust) Then
                        'B. Stock Taking
                        If Me.ObjCD.GetFirstDateofTheMonth(Me.ObjAccClosing.GetConvertionDate) <= Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value) Then
                            Filter = "[primary transaksi asal] = " + rowA.Primary_Documen_Acuan.ToString + _
                                     " AND [primary code no] = " + rowA.Primary_Code_No.ToString + _
                                     " AND [tipe transaksi asal] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InStockTakeAdjust).ToString + _
                                     " AND [tipe item detail] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem).ToString
                            Me.TBLItemDetail.DefaultView.RowFilter = Nothing
                            Me.TBLItemDetail.DefaultView.RowFilter = Filter
                            If Me.TBLItemDetail.DefaultView.Count = 0 Then
                                Me.rowItemDetail = Me.TBLItemDetail.Newtbltrans_purchase_closing_item_detailRow
                                With rowItemDetail
                                    .amount = Me.ObjCD.DigitGrouping(rowA.Unit_Price * rowA.Qty_In_Adjust, DefPrec)
                                    .unit_price = Me.ObjCD.DigitGrouping(rowA.Unit_Price, DefPrec)
                                    .primary = -1
                                    .primary_code_no = rowA.Primary_Code_No
                                    .primary_transaksi = -1
                                    If PrimaryStockTaking <> rowA.Primary_Documen_Acuan Then
                                        TglTaking = ObjTaking.GetStartDate(rowA.Primary_Documen_Acuan)
                                        PrimaryStockTaking = rowA.Primary_Documen_Acuan
                                    End If

                                    .primary_transaksi_asal = rowA.Primary_Documen_Acuan
                                    .qty = rowA.Qty_In_Adjust
                                    .tgl_transaksi_asal = TglTaking
                                    .tipe_item_detail = rowA.Tipe_item_detail
                                    .tipe_transaksi_asal = rowA.Tipe_Transaksi_Asal
                                    .return_price = 0
                                    .primary_gain_account = -1
                                    .primary_loss_account = -1
                                    .primary_adjustment_type = -1
                                    .remarks = "stock taking"
                                    .Cost_Perct_Bundling = 0
                                    .Overhead = 0
                                    If .remarks.Length > 250 Then
                                        .remarks = .remarks.Substring(0, 250)
                                    End If
                                End With
                                Me.TBLItemDetail.Rows.Add(Me.rowItemDetail)
                            Else
                                Filter = "[Primary Documen Acuan] = " + rowA.Primary_Documen_Acuan.ToString + _
                                   " AND [primary code no] = " + rowA.Primary_Code_No.ToString + _
                                   " AND [tipe transaksi asal] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InStockTakeAdjust).ToString + _
                                   " AND [tipe item detail] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem).ToString
                                Me.rowItemDetail = Me.TBLItemDetail.DefaultView(0).Row
                                Me.TblA.DefaultView.RowFilter = Nothing
                                Me.TblA.DefaultView.RowFilter = Filter
                                AmountDummy = 0
                                QtyDummy = 0
                                For Each dv As DataRowView In Me.TblA.DefaultView
                                    Me.rowA = dv.Row
                                    AmountDummy += Me.ObjCD.DigitGrouping(rowA.Unit_Price * rowA.Qty_In_Adjust, DefPrec)
                                    QtyDummy += Me.rowA.Qty_In_Adjust
                                Next
                                Me.rowItemDetail.qty = QtyDummy
                                Me.rowItemDetail.amount = AmountDummy
                                If Me.rowItemDetail.remarks.Length > 250 Then
                                    Me.rowItemDetail.remarks = Me.rowItemDetail.remarks.Substring(0, 250)
                                End If
                                Me.rowItemDetail.unit_price = Me.ObjCD.DigitGrouping(Me.rowItemDetail.amount / Me.rowItemDetail.qty, DefPrec)
                            End If
                        End If
                    ElseIf rowA.Tipe_Transaksi_Asal = CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR) Then
                        'C. Stock In Sales Return yang di jual pada bulan yang sama 
                        'C.1. Sales return atas transaksi yang dibuat sebelum convertion date
                        'C.2. Quick Sales return 
                        '-> pada cara perhitungan ini, tidak berlaku lagi sebagai adjustment, tapi tetap dihitung in qty
                        'karena menggunakan cara perhitungan yang sekarang, angka ini bisa didapat
                        counter = 0
                        Filter = "[primary transaksi asal] = " + rowA.Primary_Documen_Acuan.ToString + _
                            " AND [primary code no] = " + rowA.Primary_Code_No.ToString + _
                            " AND [tipe transaksi asal] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR).ToString + _
                            " AND [tipe item detail] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem).ToString
                        Me.TBLItemDetail.DefaultView.RowFilter = Nothing
                        Me.TBLItemDetail.DefaultView.RowFilter = Filter
                        If Me.TBLItemDetail.DefaultView.Count = 0 Then
                            Me.rowItemDetail = Me.TBLItemDetail.Newtbltrans_purchase_closing_item_detailRow
                            With rowItemDetail
                                .amount = Me.ObjCD.DigitGrouping(rowA.Unit_Price * rowA.Qty_In_Adjust, DefPrec)
                                .unit_price = Me.ObjCD.DigitGrouping(rowA.Unit_Price, DefPrec)
                                .primary = -1
                                .primary_code_no = rowA.Primary_Code_No
                                .primary_transaksi = -1
                                .primary_transaksi_asal = rowA.Primary_Documen_Acuan
                                .qty = rowA.Qty_In_Adjust
                                .tgl_transaksi_asal = rowA.Tanggal_Documen_Acuan
                                .tipe_item_detail = rowA.Tipe_item_detail
                                .tipe_transaksi_asal = rowA.Tipe_Transaksi_Asal
                                .return_price = 0
                                .primary_gain_account = -1
                                .primary_loss_account = -1
                                .primary_adjustment_type = -1
                                .remarks = ""
                                .Cost_Perct_Bundling = 0
                                .Overhead = 0
                                If .remarks.Length > 250 Then
                                    .remarks = .remarks.Substring(0, 250)
                                End If
                            End With
                            Me.TBLItemDetail.Rows.Add(Me.rowItemDetail)
                        Else
                            Filter = "[Primary Documen Acuan] = " + rowA.Primary_Documen_Acuan.ToString + _
                          " AND [primary code no] = " + rowA.Primary_Code_No.ToString + _
                          " AND [tipe transaksi asal] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR).ToString + _
                          " AND [tipe item detail] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem).ToString
                            Me.rowItemDetail = Me.TBLItemDetail.DefaultView(0).Row
                            Me.TblA.DefaultView.RowFilter = Nothing
                            Me.TblA.DefaultView.RowFilter = Filter
                            AmountDummy = 0
                            QtyDummy = 0
                            For Each dv As DataRowView In Me.TblA.DefaultView
                                rowA = dv.Row
                                AmountDummy += Me.ObjCD.DigitGrouping(rowA.Unit_Price * rowA.Qty_In_Adjust, DefPrec)
                                QtyDummy += rowA.Qty_In_Adjust
                            Next

                            If Me.rowItemDetail.remarks.Length > 250 Then
                                Me.rowItemDetail.remarks = Me.rowItemDetail.remarks.Substring(0, 250)
                            End If
                            Me.rowItemDetail.qty = QtyDummy
                            Me.rowItemDetail.amount = AmountDummy
                            Me.rowItemDetail.unit_price = Me.ObjCD.DigitGrouping(Me.rowItemDetail.amount / Me.rowItemDetail.qty, DefPrec)
                        End If
                    End If
                Next

                If Me.ObjCD.GetFirstDateofTheMonth(Me.ObjAccClosing.GetConvertionDate) <= Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value) Then
                    'D. Stock bundling tanpa batas tidak lagi berlaku
                    'E. Sales return nota return
                    'E.1. Sales return yang terjadi pada bulan ybs, ambil data cost berdasarkan cost invoice pada bulan sebelumnya yang dirujuk
                    Me.TglDOMinimum = Date.MaxValue
                    Me.DTA_SR2.Connection.ConnectionString = Core_Module.GetConstring
                    Me.DTA_SR2_Dummy.Connection.ConnectionString = Core_Module.GetConstring
                    Me.DTA_SR2_Cost.Connection.ConnectionString = Core_Module.GetConstring
                    Me.DTA_SR2.Fill(Me.TBL_SR2, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value), _
                    Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value), Me.ObjAccClosing.GetConvertionDate - TimeSpan.FromDays(1))

                    'E.2. Hitung cost - load data cost invoice
                    Me.Calc_CurrentStep += 1
                    Me.LblStep.Text = "Step " + Me.Calc_CurrentStep.ToString + " of " + Me.Calc_TotalStep.ToString + " : collecting sales return data"
                    Me.ProgressBar1.Value = 0
                    Me.LblProgress.Text = ""
                    'Me.LblProgress.Visible = True

                    If Me.TBL_SR2.Rows.Count > 0 Then
                        'tgl do minimum
                        For Each Me.row_SR2 In Me.TBL_SR2.Rows
                            If Me.row_SR2.Do_Date < Me.TglDOMinimum Then
                                Me.TglDOMinimum = Me.row_SR2.Do_Date
                            End If
                        Next
                        Me.TglDOMinimum = Me.ObjCD.GetFirstDateofTheMonth(Me.TglDOMinimum)

                        'load item cost invoice yang sudah terjadi
                        Me.DTA_SR2_Cost.Fill(Me.TBL_SR2_Cost, CInt(Me.TipeSR_CostInvoice_TipeTranskasiAsal), _
                        CInt(Me.TipeSR_CostInvoice_TipeItemDetail), Me.TglDOMinimum, Me.dtpPeriod.Value.Date)

                        'E.3. Kurangkan data cost dengan return yang sudah terjadi sebelumnya, sehinga item invoice yang terjadi adalah setelah pengurangan
                        'dan memastikan cost yang diambil tidak tumpang tindih, terutama jika FIFO/LIFO
                        While Me.TglDOMinimum >= Me.ObjAccClosing.GetConvertionDate
                            'load item return untuk bulan ybs
                            Me.DTA_SR2_Dummy.Fill(Me.TBL_SR2_Dummy, Me.ObjCD.GetLasDateOfTheMonth(Me.TglDOMinimum), _
                            Me.ObjCD.GetFirstDateofTheMonth(Me.TglDOMinimum), Me.ObjAccClosing.GetConvertionDate - TimeSpan.FromDays(1))

                            'kurangkan item invoice dengan item return
                            For Each Me.row_SR2_Dummy In Me.TBL_SR2_Dummy.Rows
                                Me.TBL_SR2_Cost.DefaultView.RowFilter = "[Primary Invoice] = " + Me.row_SR2_Dummy.Primary_Invoice_Sales.ToString + _
                                " AND [primary code no] = " + Me.row_SR2_Dummy.Primary_Product.ToString
                                While Me.row_SR2_Dummy.Qty > 0 And Me.TBL_SR2_Cost.DefaultView.Count > 0
                                    If Me.TBL_SR2_Cost.DefaultView.Count > 0 Then
                                        Me.row_SR2_Cost = Me.TBL_SR2_Cost.DefaultView(0).Row
                                        If Me.row_SR2_Dummy.Qty >= Me.row_SR2_Cost.qty Then
                                            Me.row_SR2_Dummy.Qty -= Me.row_SR2_Cost.qty
                                            Me.TBL_SR2_Cost.Rows.Remove(Me.row_SR2_Cost)
                                        Else
                                            Me.row_SR2_Cost.qty -= Me.row_SR2_Dummy.Qty
                                            AmountDummy = Math.Round((Me.row_SR2_Cost.amount * Me.row_SR2_Dummy.Qty) / Me.row_SR2_Cost.qty, _
                                            DefPrec, MidpointRounding.AwayFromZero)
                                            Me.row_SR2_Cost.amount -= AmountDummy
                                            Me.row_SR2_Dummy.Qty = 0
                                        End If
                                    Else
                                        Exit While
                                    End If
                                End While
                            Next

                            'mundur 1 bulan
                            Me.TglDOMinimum = Me.ObjCD.GetFirstDateofTheMonth(Me.TglDOMinimum - TimeSpan.FromDays(1))
                        End While

                        'E.4. Hitung cost aktual yang terjadi
                        For Each Me.row_SR2 In Me.TBL_SR2.Rows
                            Me.TBL_SR2_Cost.DefaultView.RowFilter = "[Primary Invoice] = " + Me.row_SR2.Primary_Invoice_Sales.ToString + _
                            " AND [primary code no] = " + Me.row_SR2.Primary_Product.ToString
                            QtySRDummy = Me.row_SR2.Qty
                            QtyCost = 0
                            For Each dv As DataRowView In Me.TBL_SR2_Cost.DefaultView
                                Me.row_SR2_Cost = dv.Row
                                QtyCost += Me.row_SR2_Cost.qty
                            Next
                            If QtySRDummy <= QtyCost Then
                                Me.row_SR2.Amount = 0
                                While QtySRDummy > 0
                                    Me.row_SR2_Cost = Me.TBL_SR2_Cost.DefaultView(0).Row
                                    If QtySRDummy >= Me.row_SR2_Cost.qty Then
                                        Me.row_SR2.Amount += Me.row_SR2_Cost.amount
                                        QtySRDummy -= Me.row_SR2_Cost.qty
                                        Me.TBL_SR2_Cost.Rows.Remove(Me.row_SR2_Cost)
                                    Else
                                        AmountDummy = Math.Round((Me.row_SR2_Cost.amount * QtySRDummy) / Me.row_SR2_Cost.qty, DefPrec, MidpointRounding.AwayFromZero)
                                        Me.row_SR2.Amount += AmountDummy
                                        Me.row_SR2_Cost.amount -= AmountDummy
                                        Me.row_SR2_Cost.qty -= QtySRDummy
                                        QtySRDummy = 0
                                    End If
                                End While
                            Else
                                'kemungkinan lainnya adalah, return nota return pajak, atas invoice dibulan yang sama, sehingga datanya tidak ada
                                'Me.row_SR2.Primary_Stock_In_Sales_Return
                                If Me.ApakahAdaItemAdjustment(ClsTransaksiMonthlyClosingItem.EnumTipeDokumenAcuan.StockInSalesReturnPeriodeSama, Me.row_SR2.Primary_Stock_In_Sales_Return) = True Then
                                    Me.row_SR2.Amount = 0
                                    Me.row_SR2.Amount += Math.Round(Me.row_SR2.Qty * _
                                    Me.GetCostPerUnit_Adjustment(ClsTransaksiMonthlyClosingItem.EnumTipeDokumenAcuan.StockInSalesReturnPeriodeSama, _
                                    Me.row_SR2.Primary_Stock_In_Sales_Return), DefPrec, MidpointRounding.AwayFromZero)
                                Else
                                    MsgBox("Logic error please contact your vendor.", MsgBoxStyle.Information, "Information")
                                    Exit Sub
                                End If
                            End If
                        Next

                        'E.5. Masukan ke dalam item detail
                        Me.ProgressBar1.Maximum = Me.TBL_SR2.Rows.Count
                        For Each Me.row_SR2 In Me.TBL_SR2.Rows
                            Me.ProgressBar1.Value += 1
                            If Math.IEEERemainder(Me.ProgressBar1.Value, 50) = 0 Then
                                Me.LblProgress.Text = "Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString
                            End If

                            Filter = "[primary transaksi asal] = " + Me.row_SR2.Primary_Stock_In_Sales_Return.ToString + _
                            " AND [primary code no] = " + row_SR2.Primary_Product.ToString + _
                            " AND [tipe transaksi asal] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR).ToString + _
                            " AND [tipe item detail] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem).ToString

                            Me.TBLItemDetail.DefaultView.RowFilter = Filter
                            If Me.TBLItemDetail.DefaultView.Count = 0 Then
                                Me.rowItemDetail = Me.TBLItemDetail.Newtbltrans_purchase_closing_item_detailRow
                                With rowItemDetail
                                    .amount = Me.ObjCD.DigitGrouping(Me.row_SR2.Amount, DefPrec)
                                    .unit_price = Me.ObjCD.DigitGrouping(Me.row_SR2.Amount / Me.row_SR2.Qty, DefPrec)
                                    .primary = -1
                                    .primary_code_no = row_SR2.Primary_Product
                                    .primary_transaksi = -1
                                    .primary_transaksi_asal = row_SR2.Primary_Stock_In_Sales_Return
                                    .qty = row_SR2.Qty
                                    .tgl_transaksi_asal = row_SR2.SR_Date
                                    .tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
                                    .tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR
                                    .return_price = 0
                                    .primary_gain_account = -1
                                    .primary_loss_account = -1
                                    .primary_adjustment_type = -1
                                    .remarks = ""
                                    .Cost_Perct_Bundling = 0
                                    .Overhead = 0
                                End With
                                If Me.rowItemDetail.remarks.Length > 250 Then
                                    Me.rowItemDetail.remarks = Me.rowItemDetail.remarks.Substring(0, 250)
                                End If
                                Me.TBLItemDetail.Rows.Add(Me.rowItemDetail)
                                DetailAmountDummy += rowItemDetail.amount
                            Else
                                Me.rowItemDetail = Me.TBLItemDetail.DefaultView(0).Row
                                Me.TBL_SR2.DefaultView.RowFilter = "[Primary Stock In Sales Return] = " + row_SR2.Primary_Stock_In_Sales_Return.ToString + _
                                " AND [Primary Product] = " + row_SR2.Primary_Product.ToString
                                AmountDummy = 0
                                QtyDummy = 0
                                For Each dv As DataRowView In Me.TBL_SR2.DefaultView
                                    Me.row_SR2 = dv.Row
                                    AmountDummy += Me.ObjCD.DigitGrouping(Me.row_SR2.Amount, DefPrec)
                                    QtyDummy += Me.row_SR2.Qty
                                Next
                                If Me.rowItemDetail.remarks.Length > 250 Then
                                    Me.rowItemDetail.remarks = Me.rowItemDetail.remarks.Substring(0, 250)
                                End If
                                Me.rowItemDetail.qty = QtyDummy
                                Me.rowItemDetail.amount = AmountDummy
                                Me.rowItemDetail.unit_price = Me.ObjCD.DigitGrouping(Me.rowItemDetail.amount / Me.rowItemDetail.qty, DefPrec)
                                DetailAmountDummy += AmountDummy
                            End If
                        Next
                    End If

                    '-
                    '2. load semua data purchase
                    Me.Calc_CurrentStep += 1
                    Me.LblStep.Text = "Step " + Me.Calc_CurrentStep.ToString + " of " + Me.Calc_TotalStep.ToString + " : collecting purchase data"
                    Me.ProgressBar1.Value = 0
                    Me.LblProgress.Text = ""
                    'Me.LblProgress.Visible = True

                    Me.DTAPurchase.Fill(Me.TBLPurchase, EndDate, StartDate)
                    DetailAmountDummy = 0
                    '- adjust data qty, dikurangkan dengan item return revisi aop
                    Me.DTAPR_Rev_AOP.Fill(Me.TBLPR_Rev_AOP, EndDate, StartDate)
                    For Each Me.rowPR_Rev_AOP In Me.TBLPR_Rev_AOP.Rows
                        Filter = "[Primary AOP] = " + Me.rowPR_Rev_AOP.Primary_AOP.ToString + _
                        " AND [Primary Product] = " + Me.rowPR_Rev_AOP.Primary_Product.ToString + _
                        " AND BatchNumber = '" + Me.rowPR_Rev_AOP.Batch_No + "'" + _
                        " AND [Primary PO] = " + Me.rowPR_Rev_AOP.Primary_PO.ToString

                        Me.TBLPurchase.DefaultView.RowFilter = Filter
                        If Me.TBLPurchase.DefaultView.Count = 0 Then
                            MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                            Exit Sub
                        Else
                            Me.rowPurchase = Me.TBLPurchase.DefaultView(0).Row
                            Me.rowPurchase.Total_Qty -= Me.rowPR_Rev_AOP.Total_Qty
                        End If

                        If Me.rowPurchase.Total_Qty = 0 Then
                            Me.TBLPurchase.Rows.Remove(Me.rowPurchase)
                        ElseIf Me.rowPurchase.Total_Qty < 0 Then
                            MsgBox("Logic error, please contact your vendor if you see this message." + _
                            vbCrLf + "Purchase Qty: " + Me.rowPurchase.Total_Qty.ToString + " - Filter: " + Filter, MsgBoxStyle.Information, "Information")
                            Exit Sub
                        End If
                    Next

                    Me.ProgressBar1.Maximum = Me.TBLPurchase.Rows.Count
                    For Each Me.rowPurchase In Me.TBLPurchase.Rows
                        Me.ProgressBar1.Value += 1
                        If Math.IEEERemainder(Me.ProgressBar1.Value, 50) = 0 Then
                            Me.LblProgress.Text = "Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString
                        End If

                        Filter = "[primary transaksi asal] = " + rowPurchase.Primary_AOP.ToString + _
                       " AND [primary code no] = " + rowPurchase.Primary_Product.ToString + _
                       " AND [tipe transaksi asal] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InPurchase).ToString + _
                       " AND [tipe item detail] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem).ToString
                        Me.TBLItemDetail.DefaultView.RowFilter = Filter
                        If Me.TBLItemDetail.DefaultView.Count = 0 Then
                            Me.rowItemDetail = Me.TBLItemDetail.Newtbltrans_purchase_closing_item_detailRow
                            With rowItemDetail
                                .amount = Me.ObjCD.DigitGrouping(Me.rowPurchase.At_Amount, DefPrec)
                                .unit_price = Me.ObjCD.DigitGrouping(Me.rowPurchase.At_Price, DefPrec)
                                .primary = -1
                                .primary_code_no = rowPurchase.Primary_Product
                                .primary_transaksi = -1
                                .primary_transaksi_asal = rowPurchase.Primary_AOP
                                .qty = rowPurchase.Total_Qty
                                .tgl_transaksi_asal = rowPurchase.Tanggal_Transaksi
                                .tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
                                .tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InPurchase
                                .return_price = 0
                                .primary_gain_account = -1
                                .primary_loss_account = -1
                                .primary_adjustment_type = -1
                                .remarks = ""
                                .Cost_Perct_Bundling = 0
                                .Overhead = 0
                                If Me.rowItemDetail.remarks.Length > 250 Then
                                    Me.rowItemDetail.remarks = Me.rowItemDetail.remarks.Substring(0, 250)
                                End If
                            End With
                            Me.TBLItemDetail.Rows.Add(Me.rowItemDetail)
                            DetailAmountDummy += rowItemDetail.amount
                        Else
                            Me.rowItemDetail = Me.TBLItemDetail.DefaultView(0).Row
                            Me.TBLPurchase.DefaultView.RowFilter = "[Primary AOP] = " + rowPurchase.Primary_AOP.ToString + _
                            " AND [Primary Product] = " + rowPurchase.Primary_Product.ToString
                            AmountDummy = 0
                            QtyDummy = 0
                            For Each dv As DataRowView In Me.TBLPurchase.DefaultView
                                Me.rowPurchase = dv.Row
                                AmountDummy += Me.ObjCD.DigitGrouping(Me.rowPurchase.At_Amount, DefPrec)
                                QtyDummy += Me.rowPurchase.Total_Qty
                            Next
                            Me.rowItemDetail.qty = QtyDummy
                            Me.rowItemDetail.amount = AmountDummy
                            Me.rowItemDetail.unit_price = Me.ObjCD.DigitGrouping(Me.rowItemDetail.amount / Me.rowItemDetail.qty, DefPrec)
                            DetailAmountDummy += AmountDummy
                        End If
                    Next

                    '3. load semua data out: sales, stock taking, purchase return, adjust --> mulai pada bulan convertion date
                    '3.a sales, kurangkan qty dengan stock sales return rev invoice
                    Me.Calc_CurrentStep += 1
                    Me.LblStep.Text = "Step " + Me.Calc_CurrentStep.ToString + " of " + Me.Calc_TotalStep.ToString + " : collecting sales data"
                    Me.ProgressBar1.Value = 0
                    Me.LblProgress.Text = ""
                    'Me.LblProgress.Visible = True

                    Me.DTASales.Fill(Me.TBLSales, EndDate, StartDate)
                    Me.DTASalesReturn.Fill(Me.TBLSalesReturn, EndDate, StartDate)
                    For Each Me.rowSalesReturn In Me.TBLSalesReturn.Rows
                        Filter = "[Primary Product] = " + Me.rowSalesReturn.Primary_Product.ToString + _
                        " AND [Batch Number] = '" + Me.rowSalesReturn.Batch_No + "' AND [Primary Transaksi] = " + Me.rowSalesReturn.Primary_Sales_Invoice.ToString

                        Me.TBLSales.DefaultView.RowFilter = Filter
                        If Me.TBLSales.DefaultView.Count = 0 Then
                            MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                            Exit Sub
                        Else
                            Me.rowSales = Me.TBLSales.DefaultView(0).Row
                            Me.rowSales.Total_Qty -= Me.rowSalesReturn.Total_Qty
                        End If

                        If Me.rowSales.Total_Qty = 0 Then
                            Me.TBLSales.Rows.Remove(Me.rowSales)
                        ElseIf Me.rowSales.Total_Qty < 0 Then
                            MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                            Exit Sub
                        End If
                    Next

                    Me.ProgressBar1.Maximum = Me.TBLSales.Rows.Count
                    For Each Me.rowSales In Me.TBLSales.Rows
                        Me.ProgressBar1.Value += 1
                        If Math.IEEERemainder(Me.ProgressBar1.Value, 50) = 0 Then
                            Me.LblProgress.Text = "Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString
                        End If

                        Filter = "[primary transaksi asal] = " + rowSales.Primary_Transaksi.ToString + _
                        " AND [primary code no] = " + rowSales.Primary_Product.ToString + _
                        " AND [tipe transaksi asal] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutSales).ToString + _
                        " AND [tipe item detail] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem).ToString
                        Me.TBLItemDetail.DefaultView.RowFilter = Filter
                        If Me.TBLItemDetail.DefaultView.Count = 0 Then
                            Me.rowItemDetail = Me.TBLItemDetail.Newtbltrans_purchase_closing_item_detailRow
                            With rowItemDetail
                                .amount = 0
                                .unit_price = 0
                                .primary = -1
                                .primary_code_no = rowSales.Primary_Product
                                .primary_transaksi = -1
                                .primary_transaksi_asal = rowSales.Primary_Transaksi
                                .qty = rowSales.Total_Qty
                                .tgl_transaksi_asal = rowSales.Tanggal_Transaksi
                                .tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
                                .tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutSales
                                .return_price = 0
                                .primary_gain_account = -1
                                .primary_loss_account = -1
                                .primary_adjustment_type = -1
                                .remarks = ""
                                .Cost_Perct_Bundling = 0
                                .Overhead = 0
                                If Me.rowItemDetail.remarks.Length > 250 Then
                                    Me.rowItemDetail.remarks = Me.rowItemDetail.remarks.Substring(0, 250)
                                End If
                            End With
                            Me.TBLItemDetail.Rows.Add(Me.rowItemDetail)
                        Else
                            Me.rowItemDetail = Me.TBLItemDetail.DefaultView(0).Row
                            Me.TBLSales.DefaultView.RowFilter = "[Primary Transaksi] = " + rowSales.Primary_Transaksi.ToString + _
                            " AND [Primary Product] = " + rowSales.Primary_Product.ToString
                            QtyDummy = 0
                            For Each dv As DataRowView In Me.TBLSales.DefaultView
                                Me.rowSales = dv.Row
                                QtyDummy += Me.rowSales.Total_Qty
                            Next
                            Me.rowItemDetail.qty = QtyDummy
                            Me.rowItemDetail.unit_price = 0
                        End If
                    Next

                    '3.b stock taking
                    Me.Calc_CurrentStep += 1
                    Me.LblStep.Text = "Step " + Me.Calc_CurrentStep.ToString + " of " + Me.Calc_TotalStep.ToString + " : collecting stock taking data"
                    Me.ProgressBar1.Value = 0
                    Me.LblProgress.Text = ""
                    'Me.LblProgress.Visible = True

                    Me.DTAOutTaking.Fill(Me.TBLOutTaking, EndDate, StartDate)
                    Me.ProgressBar1.Maximum = Me.TBLSales.Rows.Count
                    For Each Me.rowOutTaking In Me.TBLOutTaking.Rows
                        Me.ProgressBar1.Value += 1
                        If Math.IEEERemainder(Me.ProgressBar1.Value, 50) = 0 Then
                            Me.LblProgress.Text = "Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString
                        End If

                        'karena dari query sudah di group per product maka langsung tambahkan
                        Me.rowItemDetail = Me.TBLItemDetail.Newtbltrans_purchase_closing_item_detailRow
                        With rowItemDetail
                            .amount = 0
                            .unit_price = 0
                            .primary = -1
                            .primary_code_no = rowOutTaking.Primary_Product
                            .primary_transaksi = -1
                            .primary_transaksi_asal = rowOutTaking.Primary_Stock_Taking
                            .qty = Math.Abs(rowOutTaking.Qty)
                            .tgl_transaksi_asal = rowOutTaking.Start_Date
                            .tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
                            .tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutStockTakeAdjust
                            .return_price = 0
                            .primary_gain_account = -1
                            .primary_loss_account = -1
                            .primary_adjustment_type = -1
                            .remarks = "stock taking"
                            .Cost_Perct_Bundling = 0
                            .Overhead = 0
                            If Me.rowItemDetail.remarks.Length > 250 Then
                                Me.rowItemDetail.remarks = Me.rowItemDetail.remarks.Substring(0, 250)
                            End If
                        End With
                        Me.TBLItemDetail.Rows.Add(Me.rowItemDetail)
                    Next

                    '3.c purchase return
                    Me.Calc_CurrentStep += 1
                    Me.LblStep.Text = "Step " + Me.Calc_CurrentStep.ToString + " of " + Me.Calc_TotalStep.ToString + " : collecting purchase return data"
                    Me.ProgressBar1.Value = 0
                    Me.LblProgress.Text = ""
                    'Me.LblProgress.Visible = True

                    Me.DTA_PR.Fill(Me.TBL_PR, EndDate, TipePR, StartDate)
                    Me.ProgressBar1.Maximum = Me.TBL_PR.Rows.Count
                    For Each Me.row_PR In Me.TBL_PR.Rows
                        Me.ProgressBar1.Value += 1
                        If Math.IEEERemainder(Me.ProgressBar1.Value, 50) = 0 Then
                            Me.LblProgress.Text = "Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString
                        End If

                        Me.rowItemDetail = Me.TBLItemDetail.Newtbltrans_purchase_closing_item_detailRow
                        With rowItemDetail
                            .amount = 0
                            .unit_price = 0
                            .primary = -1
                            .primary_code_no = row_PR.Primary_Product
                            .primary_transaksi = -1
                            .primary_transaksi_asal = row_PR.Primary_Transaksi
                            .qty = row_PR.Total_Qty
                            .tgl_transaksi_asal = row_PR.Tanggal_Transaksi
                            .tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
                            .tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutPR
                            .return_price = 0
                            .primary_gain_account = -1
                            .primary_loss_account = -1
                            .primary_adjustment_type = -1
                            .remarks = ""
                            .Cost_Perct_Bundling = 0
                            .Overhead = 0
                            If Me.rowItemDetail.remarks.Length > 250 Then
                                Me.rowItemDetail.remarks = Me.rowItemDetail.remarks.Substring(0, 250)
                            End If
                        End With
                        Me.TBLItemDetail.Rows.Add(Me.rowItemDetail)
                    Next

                    '3.d out adjustment
                    Me.Calc_CurrentStep += 1
                    Me.LblStep.Text = "Step " + Me.Calc_CurrentStep.ToString + " of " + Me.Calc_TotalStep.ToString + " : collecting out adjustment data"
                    Me.ProgressBar1.Value = 0
                    Me.LblProgress.Text = ""
                    'Me.LblProgress.Visible = True

                    Me.DTA_OutAdjust.Fill(Me.TBL_OutAdjust, TipeStockOut, EndDate, StartDate)
                    Me.ProgressBar1.Maximum = Me.TBL_OutAdjust.Rows.Count
                    For Each Me.row_OutAdjust In Me.TBL_OutAdjust.Rows
                        Me.ProgressBar1.Value += 1
                        If Math.IEEERemainder(Me.ProgressBar1.Value, 50) = 0 Then
                            Me.LblProgress.Text = "Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString
                        End If

                        Me.rowItemDetail = Me.TBLItemDetail.Newtbltrans_purchase_closing_item_detailRow
                        With rowItemDetail
                            .amount = 0
                            .unit_price = 0
                            .primary = -1
                            .primary_code_no = row_OutAdjust.Primary_Product
                            .primary_transaksi = -1
                            .primary_transaksi_asal = row_OutAdjust.Primary_Transaksi
                            .qty = row_OutAdjust.Qty
                            .tgl_transaksi_asal = row_OutAdjust.Tanggal_Transaksi
                            .tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
                            .tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutAdjust
                            .return_price = 0
                            .primary_gain_account = row_OutAdjust.Primary_Loss_Account
                            .primary_loss_account = row_OutAdjust.Primary_Gain_Account
                            .primary_adjustment_type = row_OutAdjust.PrimaryAdjType
                            .Cost_Perct_Bundling = 0
                            .Overhead = 0
                            row_OutAdjust.Remarks = row_OutAdjust.Remarks.Trim
                            If row_OutAdjust.Remarks <> "" Then
                                InterimString = " : "
                            Else
                                InterimString = ""
                            End If
                            If row_OutAdjust.Adjustment_Type_Code <> "" Then
                                If (row_OutAdjust.Adjustment_Type_Desc + InterimString + row_OutAdjust.Remarks).Length > 250 Then
                                    .remarks = (row_OutAdjust.Adjustment_Type_Desc + InterimString + row_OutAdjust.Remarks).Substring(0, 250)
                                Else
                                    .remarks = row_OutAdjust.Adjustment_Type_Desc + InterimString + row_OutAdjust.Remarks
                                End If
                            Else
                                If (row_OutAdjust.Remarks).Length > 250 Then
                                    .remarks = (row_OutAdjust.Remarks).Substring(0, 250)
                                Else
                                    .remarks = row_OutAdjust.Remarks
                                End If
                            End If
                            If Me.rowItemDetail.remarks.Length > 250 Then
                                Me.rowItemDetail.remarks = Me.rowItemDetail.remarks.Substring(0, 250)
                            End If
                        End With

                        Me.TBLItemDetail.Rows.Add(Me.rowItemDetail)
                    Next

                    '4. load semua data bundling:
                    '   A. Hubungan Before Item
                    ' Me.DTABundling1.Fill(Me.TBLBundling1, EndDate, StartDate)
                    '   B. Item Before
                    Me.Calc_CurrentStep += 1
                    Me.LblStep.Text = "Step " + Me.Calc_CurrentStep.ToString + " of " + Me.Calc_TotalStep.ToString + " : collecting out bundling data"
                    Me.ProgressBar1.Value = 0
                    Me.LblProgress.Text = ""
                    'Me.LblProgress.Visible = True

                    DTABefore.Fill(TBLBefore, EndDate, StartDate)
                    Me.ProgressBar1.Maximum = TBLBefore.Rows.Count
                    For Each rowBefore In TBLBefore.Rows
                        Me.ProgressBar1.Value += 1
                        If Math.IEEERemainder(Me.ProgressBar1.Value, 50) = 0 Then
                            Me.LblProgress.Text = "Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString
                        End If

                        Me.rowItemDetail = Me.TBLItemDetail.Newtbltrans_purchase_closing_item_detailRow
                        With rowItemDetail
                            .amount = 0
                            .unit_price = 0
                            .primary = -1
                            .primary_code_no = rowBefore.Primary_Product
                            .primary_transaksi = -1
                            .primary_transaksi_asal = rowBefore.Primary_Transaksi
                            .qty = rowBefore.Qty
                            .tgl_transaksi_asal = rowBefore.Tanggal_Transaksi
                            .tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
                            .tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling
                            .return_price = 0
                            .primary_gain_account = -1
                            .primary_loss_account = -1
                            .primary_adjustment_type = -1
                            .remarks = ""
                            .Cost_Perct_Bundling = 0
                            .Overhead = 0
                            If Me.rowItemDetail.remarks.Length > 250 Then
                                Me.rowItemDetail.remarks = Me.rowItemDetail.remarks.Substring(0, 250)
                            End If
                        End With
                        Me.TBLItemDetail.Rows.Add(Me.rowItemDetail)
                    Next

                    '   C. Item After
                    Me.Calc_CurrentStep += 1
                    Me.LblStep.Text = "Step " + Me.Calc_CurrentStep.ToString + " of " + Me.Calc_TotalStep.ToString + " : collecting in bundling data"
                    Me.ProgressBar1.Value = 0
                    Me.LblProgress.Text = ""
                    'Me.LblProgress.Visible = True

                    DTAAfter.Fill1(TBLAfter, StartDate, EndDate)
                    Me.ProgressBar1.Maximum = TBLAfter.Rows.Count
                    For Each rowAfter In TBLAfter.Rows
                        Me.ProgressBar1.Value += 1
                        If Math.IEEERemainder(Me.ProgressBar1.Value, 50) = 0 Then
                            Me.LblProgress.Text = "Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString
                        End If

                        Me.rowItemDetail = Me.TBLItemDetail.Newtbltrans_purchase_closing_item_detailRow
                        With rowItemDetail
                            .amount = 0
                            .unit_price = 0
                            .primary = -1
                            .primary_code_no = rowAfter.Primary_Product
                            .primary_transaksi = -1
                            .primary_transaksi_asal = rowAfter.Primary_Transaksi
                            .qty = rowAfter.Qty
                            .tgl_transaksi_asal = rowAfter.Tanggal_Transaksi
                            .tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
                            .tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling
                            .return_price = 0
                            .primary_gain_account = -1
                            .primary_loss_account = -1
                            .primary_adjustment_type = -1
                            .remarks = ""
                            .Cost_Perct_Bundling = rowAfter.Faktor_Pemberat
                            .Overhead = 0
                            If Me.rowItemDetail.remarks.Length > 250 Then
                                Me.rowItemDetail.remarks = Me.rowItemDetail.remarks.Substring(0, 250)
                            End If
                        End With
                        Me.TBLItemDetail.Rows.Add(Me.rowItemDetail)
                    Next
                End If
            End If
            'End If

            Me.dgvItemAdjust.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            MsgBox(counter.ToString)
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
            ObjAccClosing = Nothing
            ListProduct = Nothing
            ListProductReject = Nothing
            DTABefore.Dispose()
            TBLBefore.Dispose()
            DTAAfter.Dispose()
            TBLAfter.Dispose()
        End Try
    End Sub

    Public Function ApakahAdaPerubahanAdj(ByVal vtblA As DseMonthlyClosingNew.Adjustment_CorrectionDataTable, ByVal vtblA2 As DseMonthlyClosingNew.Adjustment_Correction2DataTable) As Boolean
        Dim vrowA As DseMonthlyClosingNew.Adjustment_CorrectionRow
        Dim vrowA2 As DseMonthlyClosingNew.Adjustment_Correction2Row

        Try
            If vtblA2.Rows.Count > 0 Then
                For Each vrowA2 In vtblA2.Rows
                    vtblA.DefaultView.RowFilter = Nothing
                    vtblA.DefaultView.RowFilter = "[Primary] = " & vrowA2.Primary & _
                    " And [Primary Documen Acuan] = " & vrowA2.Primary_Documen_Acuan & _
                    " And [Tipe Transaksi Asal] = " & vrowA2.Tipe_Transaksi_Asal & _
                    " And [tipe item detail] = " & vrowA2.Tipe_item_detail & _
                    " And [Primary Code No] = " & vrowA2.Primary_Code_No
                    If vtblA.DefaultView.Count > 0 Then
                        vrowA = vtblA.DefaultView(0).Row
                        If CDec(vrowA.Unit_Price) <> CDec(vrowA2.Unit_Price) Then
                            Return True
                        End If
                    Else
                        Return True
                    End If
                Next
            Else
                Return True
            End If


            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub BgwAktifkanDetail_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwAktifkanDetail.DoWork
        Try
            Me.custom_module = Me.ObjIdentifier.GetIdentifier
            Me.ObjCD.UpdateThreading()
            Me.BgwAktifkanDetail.ReportProgress(1)

            If Me.baru = True Then
                'sort dari awal saja -> 2018-02-02
                Me.TBLItemDetail.DefaultView.Sort = "[tgl transaksi asal], [tipe transaksi asal], [primary transaksi asal], [primary code no]"

                Me.Calc_CurrentStep = 0
                Me.Calc_TotalStep = 25

                Me.GroupBox2.BringToFront()
                'alokasikan overhead
                Me.RecalculateOverhead()
                Me.BgwAktifkanDetail.ReportProgress(5)

                Me.LoadItemDetail2()
                Me.HitungKeDGVUtama1()
                If Me.TBL_Pending1.Rows.Count <> 0 Then
                    Exit Sub
                End If

                'jurnal hanya dibuat selain untuk begining balance
                Me.Calc_CurrentStep += 1
                Me.LblStep.Text = "Step " + Me.Calc_CurrentStep.ToString + " of " + Me.Calc_TotalStep.ToString + " : creating item data summary"
                Me.ProgressBar1.Value = 0
                Me.LblProgress.Text = ""
                'persiapan untuk isi data dgv
                Me.BgwAktifkanDetail_Progress_2_a()

                Me.ThreadBuffer = "Start"
                Me.LblProgress.Text = "Process 3 of 3: populating data grid view"
                Threading.Thread.Sleep(2000)
                Me.BgwAktifkanDetail.ReportProgress(2)

                While Me.ThreadBuffer = "Start"
                    Threading.Thread.Sleep(20000)
                End While

                If Me.ObjClosing.ApakahClosingPalingAwal(Me.TxtPrimaryClosing.Text) = False Then
                    Me.BuatJurnal(False)
                End If
            End If

            Me.BgwAktifkanDetail_Progress_3()
            'MsgBox(Me.Calc_CurrentStep.ToString)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AktifkanDetail_Edit()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryClosing.BackColor, False)
            Me.ObjCD.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, _
            True, True, True, True, True)

            ' Dim nt As New Threading.Thread(AddressOf Me.LoadClosingData)
            'nt.Start()
            'Me.LoadClosingData()

            Me.dgvItem.ReadOnly = False
            Me.dgvItemAdjust.ReadOnly = False
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None

            Me.TBLSetting.Rows.Clear()
            Me.DTASetting.Fill(Me.TBLSetting)
            Me.RecalculateNeeded = False

            Me.AktifkanDetail_Overhead()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwAktifkanDetail_Progress_1()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryClosing.BackColor, False)
            Me.ObjCD.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, _
            True, True, True, True, True)

            Me.dtpPeriod.Enabled = False
            Me.btnOK.Enabled = False
            Me.btnCancel.Enabled = False
            Me.TabControl1.Enabled = False
            Me.btnCekData.Enabled = False
            Me.btnFindRow.Enabled = False
            Me.ControlBox = False
            Me.btnRecalculate.Enabled = False
            Me.Button3.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwAktifkanDetail_Progress_4()
        Try
            'MsgBox("Cyclic bundling detected, please input unit price for the following data.", MsgBoxStyle.Information, "Information")

            Dim mn As New FrmTransaksiMonthlyClosing_CyclicBundlingExplain
            mn.BackColorAll = Me.BAckColorAll
            mn.TBL_Pending = TBL_Pending1
            mn.TBL_Pending2 = Me.TBL_Pending2
            mn.StartDate = Me.dtpPeriod.Value.Date
            mn.EndDate = Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value.Date)
            mn.ClosingPeriod = Me.dtpPeriod.Value.Date
            mn.chkAutomatic.Checked = Me.HitungOtomatis
            mn.TBLItemClosing = Me.TBLItemClosing
            mn.TBLItemDetail = Me.TBLItemDetail

            If mn.ShowDialog() = Windows.Forms.DialogResult.OK Then
                Me.HitungOtomatis = mn.chkAutomatic.Checked
            Else
                'Me.TBLItemDetail.Rows.Clear() -> dipindahkan ke fungsi hitung cost, sehingga gampang kalo mau melacak error ketika mau closing
            End If

            mn = Nothing
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwAktifkanDetail_Progress_3()
        Try
            If Me.ObjClosing.ApakahClosingPalingAwal(Me.TxtPrimaryClosing.Text) = True Then
                Me.ColInAmount.Visible = False
                Me.ColInQty.Visible = False
                Me.ColBeginQty.Visible = False
                Me.ColBeginAmount.Visible = False
                Me.ColOutQty.Visible = False
                Me.ColOutAmount.Visible = False
            Else
                Me.ColInAmount.Visible = True
                Me.ColInQty.Visible = True
                Me.ColBeginQty.Visible = True
                Me.ColBeginAmount.Visible = True
                Me.ColOutQty.Visible = True
                Me.ColOutAmount.Visible = True
            End If

            Me.dgvItem.ReadOnly = False
            'Me.dgvItemAdjust.ReadOnly = False
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None

            Me.TBLSetting.Rows.Clear()
            Me.DTASetting.Fill(Me.TBLSetting)

            Me.AktifkanDetail_Overhead()
            Me.btnOK.Enabled = True
            Me.btnCancel.Enabled = True
            Me.btnRecalculate.Enabled = True
            Me.TabControl1.Enabled = True
            Me.btnCekData.Enabled = True
            Me.btnFindRow.Enabled = True
            Me.ControlBox = True

            Me.GroupBox2.SendToBack()
            Me.ProgressBar1.Value = 0
            Me.LblProgress.Text = ""
            Me.LblStep.Text = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Private Sub BgwAktifkanDetail_Progress_2_a()
    '    Dim ListPrimaryCodeNo As New List(Of Double)
    '    Dim decQtyBegin As Decimal = 0
    '    Dim decAmountBegin As Decimal = 0
    '    Dim decQtyIn As Decimal = 0
    '    Dim decAmountIn As Decimal = 0
    '    Dim decQtyOut As Decimal = 0
    '    Dim decAmountOut As Decimal = 0
    '    Dim decQtyEnd As Decimal = 0
    '    Dim decAmountEnd As Decimal = 0
    '    Dim Decimalx As Integer = 0
    '    'Dim rowDgv As DataGridViewRow
    '    Dim TBLItemDetail_Load As New DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailDataTable
    '    Dim rowItemDetail_Load As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailRow
    '    Dim DV As DataRowView
    '    Dim PrimaryBefore As Double = -1

    '    Try
    '        TBLItemDetail_Load = Me.TBLItemDetail.Copy
    '        TBLItemDetail_Load.DefaultView.Sort = "[primary code no]"
    '        ListPrimaryCodeNo.Clear()
    '        Me.LblProgress.Text = "Process 1 of 4: Collecting code no"
    '        Me.ProgressBar1.Value = 0
    '        Me.ProgressBar1.Maximum = TBLItemDetail_Load.Count

    '        For Each DV In TBLItemDetail_Load.DefaultView
    '            Me.ProgressBar1.Value += 1
    '            If Me.ProgressBar1.Value > 25 Then
    '                If Math.IEEERemainder(Me.ProgressBar1.Value, 100) = 0 Then
    '                    Me.LblProgress.Text = "Process 1 of 4: Collecting code no - data: " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString
    '                ElseIf Me.ProgressBar1.Value = Me.ProgressBar1.Maximum Then
    '                    Me.LblProgress.Text = "Process 1 of 4: Collecting code no - data: " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString
    '                End If
    '            Else
    '                Me.LblProgress.Text = "Process 1 of 4: Collecting code no - data: " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString
    '            End If

    '            Me.rowItemDetail = DV.Row
    '            If PrimaryBefore <> Me.rowItemDetail.primary_code_no Then
    '                ListPrimaryCodeNo.Add(rowItemDetail.primary_code_no)
    '                PrimaryBefore = rowItemDetail.primary_code_no
    '            End If
    '        Next

    '        '4. salin ke dgv utama
    '        Me.TBLProduct.Rows.Clear()
    '        Me.ProgressBar1.Value = 0
    '        Me.LblProgress.Text = "Process 2 of 4: collecting detail description"
    '        Me.DTAProduct.Fill(TBLProduct)

    '        Me.ProgressBar1.Value = 0
    '        Me.ProgressBar1.Maximum = ListPrimaryCodeNo.Count
    '        Me.LblProgress.Text = "Process 3 of 4: calculating detail item qty and amount"
    '        For Each PrimaryCodeNo As Double In ListPrimaryCodeNo
    '            Me.ProgressBar1.Value += 1
    '            If Me.ProgressBar1.Value > 25 Then
    '                If Math.IEEERemainder(Me.ProgressBar1.Value, 100) = 0 Then
    '                    Me.LblProgress.Text = "Process 3 of 4: calculating detail item qty and amount - data: " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString
    '                ElseIf Me.ProgressBar1.Value = Me.ProgressBar1.Maximum Then
    '                    Me.LblProgress.Text = "Process 3 of 4: calculating detail item qty and amount - data: " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString
    '                End If
    '            Else
    '                Me.LblProgress.Text = "Process 3 of 4: calculating detail item qty and amount - data: " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString
    '            End If

    '            Me.TBLProduct.DefaultView.RowFilter = "[Primary] = " & PrimaryCodeNo.ToString

    '            If Me.TBLProduct.DefaultView.Count > 0 Then
    '                Me.rowProduct = Me.TBLProduct.DefaultView(0).Row
    '                Decimalx = Me.rowProduct._Decimal

    '                decQtyBegin = 0
    '                decAmountBegin = 0
    '                decQtyIn = 0
    '                decAmountIn = 0
    '                decQtyOut = 0
    '                decAmountOut = 0
    '                decQtyEnd = 0
    '                decAmountEnd = 0

    '                TBLItemDetail_Load.DefaultView.RowFilter = "[primary code no] = " & PrimaryCodeNo.ToString

    '                If TBLItemDetail_Load.DefaultView.Count > 0 Then
    '                    While TBLItemDetail_Load.DefaultView.Count > 0
    '                        Me.rowItemDetail = TBLItemDetail_Load.DefaultView(0).Row
    '                        With Me.rowItemDetail
    '                            If .tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.BeginingBalanceItem Then
    '                                If IsNumeric(.qty) Then
    '                                    decQtyBegin += CDec(.qty)
    '                                    decAmountBegin += CDec(.amount)
    '                                End If
    '                            ElseIf .tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem Then
    '                                If IsNumeric(.qty) Then
    '                                    decQtyIn += CDec(.qty)
    '                                    decAmountIn += CDec(.amount)
    '                                End If
    '                            ElseIf .tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem Then
    '                                If IsNumeric(.qty) Then
    '                                    decQtyOut += CDec(.qty)
    '                                    decAmountOut += CDec(.amount)
    '                                End If
    '                            ElseIf .tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.EndingBalanceItem Then
    '                                If IsNumeric(.qty) Then
    '                                    decQtyEnd += CDec(.qty)
    '                                    decAmountEnd += CDec(.amount)
    '                                End If
    '                            End If
    '                        End With
    '                        TBLItemDetail_Load.Rows.Remove(Me.rowItemDetail)
    '                    End While
    '                End If

    '                Me.rowProduct.Qty_BB = Me.ObjCD.DigitGrouping(decQtyBegin, Decimalx)
    '                Me.rowProduct.Amount_BB = Me.ObjCD.DigitGrouping(decAmountBegin, Me.DefaultPrecision)
    '                Me.rowProduct.Qty_In = Me.ObjCD.DigitGrouping(decQtyIn, Decimalx)
    '                Me.rowProduct.Amount_In = Me.ObjCD.DigitGrouping(decAmountIn, Me.DefaultPrecision)
    '                Me.rowProduct.Qty_Out = Me.ObjCD.DigitGrouping(decQtyOut, Decimalx)
    '                Me.rowProduct.Amount_Out = Me.ObjCD.DigitGrouping(decAmountOut, Me.DefaultPrecision)
    '                Me.rowProduct.Qty_Ending = Me.ObjCD.DigitGrouping(decQtyEnd, Decimalx)
    '                Me.rowProduct.Amount_Ending = Me.ObjCD.DigitGrouping(decAmountEnd, Me.DefaultPrecision)
    '            End If
    '        Next
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        ListPrimaryCodeNo = Nothing
    '        TBLItemDetail_Load.Rows.Clear()
    '        TBLItemDetail_Load.Dispose()
    '        ListPrimaryCodeNo = Nothing
    '    End Try
    'End Sub

    Private Sub BgwAktifkanDetail_Progress_2_a()
        Dim decQtyBegin As Decimal = 0
        Dim decAmountBegin As Decimal = 0
        Dim decQtyIn As Decimal = 0
        Dim decAmountIn As Decimal = 0
        Dim decQtyOut As Decimal = 0
        Dim decAmountOut As Decimal = 0
        Dim decQtyEnd As Decimal = 0
        Dim decAmountEnd As Decimal = 0
        Dim Decimalx As Integer = 0
        Dim DVProduct As DataRowView
        Dim StartIdx As Double = 0

        Try
            If Me.TBLItemDetail.DefaultView.RowFilter <> "" Then
                TBLItemDetail.DefaultView.RowFilter = ""
            End If
            If TBLItemDetail.DefaultView.Sort <> "[Primary Code No]" Then
                TBLItemDetail.DefaultView.Sort = "[Primary Code No]"
            End If

            '4. salin ke dgv utama
            Me.TBLProduct.Rows.Clear()
            Me.ProgressBar1.Value = 0
            Me.LblProgress.Text = "Process 1 of 3: collecting detail description"
            Me.DTAProduct.Fill(TBLProduct)
            Me.TBLProduct.DefaultView.Sort = "Primary"
            Me.TBLProduct.DefaultView.RowFilter = ""

            Me.ProgressBar1.Value = 0
            Me.ProgressBar1.Maximum = TBLProduct.DefaultView.Count
            Me.LblProgress.Text = "Process 2 of 3: calculating detail item qty and amount"

            For Each DVProduct In TBLProduct.DefaultView
                Me.rowProduct = DVProduct.Row
                Decimalx = Me.rowProduct._Decimal

                decQtyBegin = 0
                decAmountBegin = 0
                decQtyIn = 0
                decAmountIn = 0
                decQtyOut = 0
                decAmountOut = 0
                decQtyEnd = 0
                decAmountEnd = 0

                While StartIdx <= Me.TBLItemDetail.DefaultView.Count - 1
                    Me.rowItemDetail = Me.TBLItemDetail.DefaultView(StartIdx).Row
                    If Me.rowItemDetail.primary_code_no = Me.rowProduct.Primary Then
                        With Me.rowItemDetail
                            If .tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.BeginingBalanceItem Then
                                If IsNumeric(.qty) Then
                                    decQtyBegin += CDec(.qty)
                                    decAmountBegin += CDec(.amount)
                                End If
                            ElseIf .tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem Then
                                If IsNumeric(.qty) Then
                                    decQtyIn += CDec(.qty)
                                    decAmountIn += CDec(.amount)
                                End If
                            ElseIf .tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem Then
                                If IsNumeric(.qty) Then
                                    decQtyOut += CDec(.qty)
                                    decAmountOut += CDec(.amount)
                                End If
                            ElseIf .tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.EndingBalanceItem Then
                                If IsNumeric(.qty) Then
                                    decQtyEnd += CDec(.qty)
                                    decAmountEnd += CDec(.amount)
                                End If
                            End If
                        End With
                        StartIdx += 1
                    ElseIf Me.rowItemDetail.primary_code_no > Me.rowProduct.Primary Then
                        If StartIdx > 0 Then
                            StartIdx -= 1
                        End If
                        Exit While
                    Else
                        StartIdx += 1
                    End If
                End While

                Me.rowProduct.Qty_BB = Me.ObjCD.DigitGrouping(decQtyBegin, Decimalx)
                Me.rowProduct.Amount_BB = Me.ObjCD.DigitGrouping(decAmountBegin, Me.DefaultPrecision)
                Me.rowProduct.Qty_In = Me.ObjCD.DigitGrouping(decQtyIn, Decimalx)
                Me.rowProduct.Amount_In = Me.ObjCD.DigitGrouping(decAmountIn, Me.DefaultPrecision)
                Me.rowProduct.Qty_Out = Me.ObjCD.DigitGrouping(decQtyOut, Decimalx)
                Me.rowProduct.Amount_Out = Me.ObjCD.DigitGrouping(decAmountOut, Me.DefaultPrecision)
                Me.rowProduct.Qty_Ending = Me.ObjCD.DigitGrouping(decQtyEnd, Decimalx)
                Me.rowProduct.Amount_Ending = Me.ObjCD.DigitGrouping(decAmountEnd, Me.DefaultPrecision)
            Next
        Catch ex As Exception
            MsgBox(ex.Message + "Calc. Process: " + Me.Calc_CurrentStep.ToString + " - BgwAktifkanDetail_Progress_2_a: Primary Product " + Me.rowProduct.Primary.ToString, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwAktifkanDetail_Progress_2_b()
        Dim Decimalx As Integer = 0
        Dim rowDgv As DataGridViewRow

        Try
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            For Each rowProduct In TBLProduct.Rows
                With rowProduct
                    If rowProduct.Qty_BB <> 0 Or rowProduct.Qty_In <> 0 Or .Qty_Out <> 0 Or .Qty_Ending <> 0 _
                    Or .Amount_BB <> 0 Or .Amount_In <> 0 Or .Amount_In <> 0 Or .Amount_Ending <> 0 Then
                        Decimalx = Me.rowProduct._Decimal
                        Me.dgvItem.Rows.Add()

                        rowDgv = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                        rowDgv.Cells(Me.ColNo.Index).Value = Me.dgvItem.Rows.Count
                        rowDgv.Cells(Me.ColPrimaryCodeNo.Index).Value = Me.rowProduct.Primary
                        rowDgv.Cells(Me.ColCodeNo.Index).Value = Me.rowProduct.Product_Code
                        rowDgv.Cells(Me.ColDescription.Index).Value = Me.rowProduct.Description
                        rowDgv.Cells(Me.ColPrimaryUnits.Index).Value = Me.rowProduct.Primary_Unit
                        rowDgv.Cells(Me.ColUnits.Index).Value = Me.rowProduct.Unit_Code
                        rowDgv.Cells(Me.ColItem_PrimaryAdjustmentAccount.Index).Value = Me.rowProduct.Primary_Adjustment_Account
                        rowDgv.Cells(Me.ColItem_PrimaryCOGSAccount.Index).Value = Me.rowProduct.Primary_COGS_Account
                        rowDgv.Cells(Me.ColItem_PrimaryInventoryAccount.Index).Value = Me.rowProduct.Primary_Inventory_Account
                        rowDgv.Cells(Me.ColItem_PrimarySalesAccount.Index).Value = Me.rowProduct.Primary_Sales_Account
                        rowDgv.Cells(Me.ColItem_PrimarySalesReturnAccount.Index).Value = Me.rowProduct.Primary_Sales_Return_Account
                        rowDgv.Cells(Me.ColItem_PrimaryStockTakingAccount.Index).Value = Me.rowProduct.Primary_Stock_Taking_Account
                        rowDgv.Cells(Me.ColDetailCalculation.Index).Value = "Show"
                        rowDgv.Cells(Me.ColBeginQty.Index).Value = Me.ObjCD.DigitGrouping(.Qty_BB, Decimalx)
                        rowDgv.Cells(Me.ColBeginAmount.Index).Value = Me.ObjCD.DigitGrouping(.Amount_BB, Me.DefaultPrecision)
                        rowDgv.Cells(Me.ColInQty.Index).Value = Me.ObjCD.DigitGrouping(.Qty_In, Decimalx)
                        rowDgv.Cells(Me.ColInAmount.Index).Value = Me.ObjCD.DigitGrouping(.Amount_In, Me.DefaultPrecision)
                        rowDgv.Cells(Me.ColOutQty.Index).Value = Me.ObjCD.DigitGrouping(.Qty_Out, Decimalx)
                        rowDgv.Cells(Me.ColOutAmount.Index).Value = Me.ObjCD.DigitGrouping(.Amount_Out, Me.DefaultPrecision)
                        rowDgv.Cells(Me.ColEndQty.Index).Value = Me.ObjCD.DigitGrouping(.Qty_Ending, Decimalx)
                        rowDgv.Cells(Me.ColEndAmount.Index).Value = Me.ObjCD.DigitGrouping(.Amount_Ending, Me.DefaultPrecision)
                        rowDgv.Cells(Me.ColRemarks.Index).Value = ""
                    End If
                End With
            Next

            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.TBLProduct.Rows.Clear()
            'tandai proses sudah selesai
            Me.ThreadBuffer = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Private Sub BgwAktifkanDetail_Progress_2()
    '    Dim ListPrimaryCodeNo As New List(Of Double)
    '    Dim decQtyBegin As Decimal = 0
    '    Dim decAmountBegin As Decimal = 0
    '    Dim decQtyIn As Decimal = 0
    '    Dim decAmountIn As Decimal = 0
    '    Dim decQtyOut As Decimal = 0
    '    Dim decAmountOut As Decimal = 0
    '    Dim decQtyEnd As Decimal = 0
    '    Dim decAmountEnd As Decimal = 0
    '    Dim Decimalx As Integer = 0
    '    Dim rowDgv As DataGridViewRow

    '    Try
    '        '4. salin ke dgv utama
    '        ListPrimaryCodeNo.Clear()
    '        For Each Me.rowItemDetail In Me.TBLItemDetail.Rows
    '            If ListPrimaryCodeNo.Contains(Me.rowItemDetail.primary_code_no) = False Then
    '                ListPrimaryCodeNo.Add(Me.rowItemDetail.primary_code_no)
    '            End If
    '        Next
    '        Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
    '        Me.DTAProduct.Fill(TBLProduct)

    '        For Each PrimaryCodeNo As Double In ListPrimaryCodeNo
    '            Me.TBLProduct.DefaultView.RowFilter = "[Primary] = " & PrimaryCodeNo.ToString

    '            If Me.TBLProduct.DefaultView.Count > 0 Then
    '                Me.rowProduct = Me.TBLProduct.DefaultView(0).Row
    '                Decimalx = Me.rowProduct._Decimal
    '                Me.dgvItem.Rows.Add()

    '                rowDgv = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
    '                rowDgv.Cells(Me.ColPrimaryCodeNo.Index).Value = Me.rowProduct.Primary
    '                rowDgv.Cells(Me.ColCodeNo.Index).Value = Me.rowProduct.Product_Code
    '                rowDgv.Cells(Me.ColDescription.Index).Value = Me.rowProduct.Description
    '                rowDgv.Cells(Me.ColPrimaryUnits.Index).Value = Me.rowProduct.Primary_Unit
    '                rowDgv.Cells(Me.ColUnits.Index).Value = Me.rowProduct.Unit_Code
    '                rowDgv.Cells(Me.ColItem_PrimaryAdjustmentAccount.Index).Value = Me.rowProduct.Primary_Adjustment_Account
    '                rowDgv.Cells(Me.ColItem_PrimaryCOGSAccount.Index).Value = Me.rowProduct.Primary_COGS_Account
    '                rowDgv.Cells(Me.ColItem_PrimaryInventoryAccount.Index).Value = Me.rowProduct.Primary_Inventory_Account
    '                rowDgv.Cells(Me.ColItem_PrimarySalesAccount.Index).Value = Me.rowProduct.Primary_Sales_Account
    '                rowDgv.Cells(Me.ColItem_PrimarySalesReturnAccount.Index).Value = Me.rowProduct.Primary_Sales_Return_Account
    '                rowDgv.Cells(Me.ColItem_PrimaryStockTakingAccount.Index).Value = Me.rowProduct.Primary_Stock_Taking_Account
    '                rowDgv.Cells(Me.ColDetailCalculation.Index).Value = "Show"

    '                decQtyBegin = 0
    '                decAmountBegin = 0
    '                decQtyIn = 0
    '                decAmountIn = 0
    '                decQtyOut = 0
    '                decAmountOut = 0
    '                decQtyEnd = 0
    '                decAmountEnd = 0

    '                Me.TBLItemDetail.DefaultView.RowFilter = "[primary code no] = " & PrimaryCodeNo.ToString

    '                If Me.TBLItemDetail.DefaultView.Count > 0 Then
    '                    For Each dvx As Data.DataRowView In Me.TBLItemDetail.DefaultView
    '                        Me.rowItemDetail = dvx.Row
    '                        With Me.rowItemDetail
    '                            If .tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.BeginingBalanceItem Then
    '                                If IsNumeric(.qty) Then
    '                                    decQtyBegin += CDec(.qty)
    '                                    decAmountBegin += CDec(.amount)
    '                                End If
    '                            ElseIf .tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem Then
    '                                If IsNumeric(.qty) Then
    '                                    decQtyIn += CDec(.qty)
    '                                    decAmountIn += CDec(.amount)
    '                                End If
    '                            ElseIf .tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem Then
    '                                If IsNumeric(.qty) Then
    '                                    decQtyOut += CDec(.qty)
    '                                    decAmountOut += CDec(.amount)
    '                                End If
    '                            ElseIf .tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.EndingBalanceItem Then
    '                                If IsNumeric(.qty) Then
    '                                    decQtyEnd += CDec(.qty)
    '                                    decAmountEnd += CDec(.amount)
    '                                End If
    '                            End If
    '                        End With
    '                    Next
    '                End If

    '                rowDgv.Cells(Me.ColBeginQty.Index).Value = Me.ObjCD.DigitGrouping(decQtyBegin, Decimalx)
    '                rowDgv.Cells(Me.ColBeginAmount.Index).Value = Me.ObjCD.DigitGrouping(decAmountBegin, Me.DefaultPrecision)
    '                rowDgv.Cells(Me.ColInQty.Index).Value = Me.ObjCD.DigitGrouping(decQtyIn, Decimalx)
    '                rowDgv.Cells(Me.ColInAmount.Index).Value = Me.ObjCD.DigitGrouping(decAmountIn, Me.DefaultPrecision)
    '                rowDgv.Cells(Me.ColOutQty.Index).Value = Me.ObjCD.DigitGrouping(decQtyOut, Decimalx)
    '                rowDgv.Cells(Me.ColOutAmount.Index).Value = Me.ObjCD.DigitGrouping(decAmountOut, Me.DefaultPrecision)
    '                rowDgv.Cells(Me.ColEndQty.Index).Value = Me.ObjCD.DigitGrouping(decQtyEnd, Decimalx)
    '                rowDgv.Cells(Me.ColEndAmount.Index).Value = Me.ObjCD.DigitGrouping(decAmountEnd, Me.DefaultPrecision)
    '                rowDgv.Cells(Me.ColRemarks.Index).Value = ""
    '            End If
    '        Next

    '        Me.dgvItem.Sort(Me.ColCodeNo, System.ComponentModel.ListSortDirection.Ascending)
    '        Me.ObjCD.SetNumber(Me.dgvItem)

    '        Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

    '        'tandai proses sudah selesai
    '        Me.ThreadBuffer = ""
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        ListPrimaryCodeNo = Nothing
    '    End Try
    'End Sub

    Private Sub BgwAktifkanDetail_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BgwAktifkanDetail.ProgressChanged
        Try
            If e.ProgressPercentage = 1 Then
                Me.BgwAktifkanDetail_Progress_1()
            ElseIf e.ProgressPercentage = 2 Then
                Me.BgwAktifkanDetail_Progress_2_b()
            ElseIf e.ProgressPercentage = 3 Then
                Me.BgwAktifkanDetail_Progress_3()
            ElseIf e.ProgressPercentage = 5 Then
                Me.BgwAktifkanDetail_Progress_5()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwAktifkanDetail_Progress_5()
        Try
            Me.TampilkanOverheadDetail()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwAktifkanDetail_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwAktifkanDetail.RunWorkerCompleted
        Try
            If Me.TBL_Pending1.Rows.Count = 0 Then
                Me.RecalculateNeeded = False
                Me.btnOK.PerformClick()
            Else
                'Me.BgwAktifkanDetail_Progress_4()
                Me.dgvJurnal.Rows.Clear()
                Me.dgvItem.Rows.Clear()
                BgwAktifkanDetail_Progress_3()
                Me.RecalculateNeeded = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HitungKeDGVUtama1()
        Dim TBLInvPool As New DseMonthlyClosingNew.tblItemDetailPoolDataTable
        Dim CostingMethod As ClsTransaksiMonthlyClosingItem.EnumCostingMethod = ClsTransaksiMonthlyClosingItem.EnumCostingMethod.Average
        Dim DummyCostingMethod As String = ""


        Try
            DefaultPrecision = Me.objCurrency.GetDecimal_DefaultCurrency

            '1. buat inventory pool berdasarkan data ending bulan sebelumnya
            Me.IsikanDataAwalInventoryPool(TBLInvPool)

            '2. hitung item detail urut tanggal dari 1-31, hitung in dulu baru out
            DummyCostingMethod = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.CostingMethod)
            If DummyCostingMethod = "" Then
                CostingMethod = ClsTransaksiMonthlyClosingItem.EnumCostingMethod.Average
            Else
                CostingMethod = DummyCostingMethod
            End If
            Me.HitungCost(TBLInvPool, CostingMethod)

            '3. kembalikan data dari inventory pool menjadi end balance
            Me.IsikanDataEndBalance(TBLInvPool)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TBLInvPool.Dispose()
        End Try
    End Sub

    Private Sub IsikanDataEndBalance(ByVal TBLInvPool As DseMonthlyClosingNew.tblItemDetailPoolDataTable)
        Dim rowInvPool As DseMonthlyClosingNew.tblItemDetailPoolRow
        Dim rowItemDetail_D As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailRow

        Try
            Me.Calc_CurrentStep += 1
            Me.LblStep.Text = "Step " + Me.Calc_CurrentStep.ToString + " of " + Me.Calc_TotalStep.ToString + " : calculating end balance data"
            Me.ProgressBar1.Value = 0
            Me.LblProgress.Text = ""
            'Me.LblProgress.Visible = True
            Me.ProgressBar1.Maximum = TBLInvPool.Rows.Count

            For Each rowInvPool In TBLInvPool.Rows
                Me.ProgressBar1.Value += 1
                If Math.IEEERemainder(Me.ProgressBar1.Value, 500) = 0 Then
                    Me.LblProgress.Text = "Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString
                End If

                rowItemDetail_D = Me.TBLItemDetail.Newtbltrans_purchase_closing_item_detailRow
                With rowItemDetail_D
                    .amount = rowInvPool.amount
                    .Cost_Perct_Bundling = rowInvPool.Cost_Perct_Bundling
                    .Overhead = rowInvPool.Overhead
                    .primary = -1
                    .primary_adjustment_type = rowInvPool.primary_adjustment_type
                    .primary_code_no = rowInvPool.primary_code_no
                    .primary_gain_account = rowInvPool.primary_gain_account
                    .primary_loss_account = rowInvPool.primary_loss_account
                    .primary_transaksi = Me.TxtPrimaryClosing.Text
                    .primary_transaksi_asal = rowInvPool.Primary_Transaksi_Asal
                    .qty = rowInvPool.qty
                    .remarks = rowInvPool.remarks
                    .return_price = rowInvPool.return_price
                    .tgl_transaksi_asal = rowInvPool.tgl_transaksi_asal
                    .tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.EndingBalanceItem
                    .tipe_transaksi_asal = rowInvPool.tipe_transaksi_asal
                    .unit_price = rowInvPool.unit_price
                    If .remarks.Length > 250 Then
                        .remarks = .remarks.Substring(0, 250)
                    End If
                End With
                Me.TBLItemDetail.Rows.Add(rowItemDetail_D)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub MasukkanOverheadAmount()
        Dim Dummy As Decimal = 0
        Dim DV1 As DataRowView
        Dim row1 As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_code_noRow

        Try
            'kosongkan nilai overhead, pada setiap item in bundling
            Me.Calc_CurrentStep += 1
            Me.LblStep.Text = "Step " + Me.Calc_CurrentStep.ToString + " of " + Me.Calc_TotalStep.ToString + " : inserting overhead amount data"
            Me.ProgressBar1.Value = 0
            Me.LblProgress.Text = ""
            'Me.LblProgress.Visible = True

            'untuk setiap item in bundling, cari nilai overheadnya
            Me.TBLItemDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem).ToString + _
            " AND [tipe transaksi asal] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling).ToString

            Me.ProgressBar1.Maximum = Me.TBLItemDetail.DefaultView.Count
            For Each DV As DataRowView In Me.TBLItemDetail.DefaultView
                Me.ProgressBar1.Value += 1
                If Math.IEEERemainder(Me.ProgressBar1.Value, 50) = 0 Then
                    Me.LblProgress.Text = "Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString
                End If

                Me.rowItemDetail = DV.Row
                Dummy = 0

                Me.vTBLCostGroup_CodeNo.DefaultView.RowFilter = "[primary bundling] = " + Me.rowItemDetail.primary_transaksi_asal.ToString + " AND [primary code no] = " + Me.rowItemDetail.primary_code_no.ToString
                For Each DV1 In Me.vTBLCostGroup_CodeNo.DefaultView
                    row1 = DV1.Row
                    Dummy += row1.allocated_amount
                Next

                Me.rowItemDetail.Overhead = Dummy
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HitungCost(ByVal TBLInvPool As DseMonthlyClosingNew.tblItemDetailPoolDataTable, _
    ByVal CostingMethod As ClsTransaksiMonthlyClosingItem.EnumCostingMethod)
        Dim CalcDate As Date = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value)
        Dim EndDate As Date = Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value)
        Dim TBL_Pending As New DseMonthlyClosingNew.tblItemDetailPoolDataTable
        Dim row_Pending, row_Pending11, row_Pending2 As DseMonthlyClosingNew.tblItemDetailPoolRow
        Dim DummyFilter As String = ""
        Dim DummySort As String = ""
        Dim ListPrimaryBundling As New List(Of Double)
        Dim rowItemDetail_D1 As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailRow

        Dim counter As Double = 0
        Dim DefPrec As Integer = 0
        Dim LoopingPendingWithoutProgress As Integer
        Dim QtyPending1, QtyPending2 As Decimal
        Dim HitungUlang As Boolean = False
        Dim JmlLoop As Integer = 0
        Dim ObjIdentifier As New CustomSMESS.ClsIdentifier 
        Dim JmlAwal As Double
        Dim SudahPindah As Boolean = False
        Dim SudahPindah1 As Boolean = False
        Dim SudahPindah2 As Boolean = False 

        Try
            If CostingMethod = ClsTransaksiMonthlyClosingItem.EnumCostingMethod.FIFO Then
                TBLInvPool.DefaultView.Sort = "[tgl transaksi asal], [tipe transaksi asal], [primary transaksi asal], [primary code no]"
            Else
                TBLInvPool.DefaultView.Sort = "[tgl transaksi asal] DESC, [tipe transaksi asal], [primary transaksi asal] DESC, [primary code no]"
            End If

            Me.TBLInvPool_Dummy.Rows.Clear()
            Me.TBLItemDetail_Dummy.Rows.Clear()

            'CalcDate = #2/28/2016#
            Me.TBL_Pending1.Rows.Clear()
            DefPrec = Me.objCurrency.GetDecimal_DefaultCurrency

            'buang item detail
            Me.TBLItemDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.EndingBalanceItem).ToString
            While Me.TBLItemDetail.DefaultView.Count > 0
                Me.TBLItemDetail.Rows.Remove(Me.TBLItemDetail.DefaultView(0).Row)
            End While

            'set nol amount untuk item out 
            Me.TBLItemDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem).ToString
            For Each dv As DataRowView In Me.TBLItemDetail.DefaultView
                Me.rowItemDetail = dv.Row
                Me.rowItemDetail.amount = 0
            Next
            'set nol untuk item in bundling: overhead + amount 
            Me.TBLItemDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem).ToString + _
            " AND [tipe transaksi asal] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling).ToString
            For Each dv As DataRowView In Me.TBLItemDetail.DefaultView
                Me.rowItemDetail = dv.Row
                Me.rowItemDetail.amount = 0
                Me.rowItemDetail.Overhead = 0
            Next

            'masukkan nilai overhead
            Me.MasukkanOverheadAmount()

            'sort item detail berdasarkan tanggal transaksi asal, tipe transaksi asal (in dulu baru out), primary product
            'revisi: 2018-02-02: sort 1x saja
            'Me.''TBLItemDetail.DefaultView.Sort = "[tgl transaksi asal], [tipe transaksi asal], [primary transaksi asal], [primary code no]"
            Me.Calc_CurrentStep += 1
            Me.LblStep.Text = "Step " + Me.Calc_CurrentStep.ToString + " of " + Me.Calc_TotalStep.ToString + " : compacting product with no transaction"
            Me.ProgressBar1.Value = 0
            Me.LblProgress.Text = ""
            Me.MasukkanKeDummy1(TBLInvPool) 

            'filter yang ditampilkan hanya yang tipenya in + out saja
            Me.Calc_CurrentStep += 1
            Me.LblStep.Text = "Step " + Me.Calc_CurrentStep.ToString + " of " + Me.Calc_TotalStep.ToString + " : calculating cost data"
            Me.ProgressBar1.Value = 0
            Me.LblProgress.Text = ""
            'Me.LblProgress.Visible = True

            Me.ProgressBar1.Maximum = (EndDate.Date - CalcDate.Date).TotalDays + 1
            While CalcDate.Date <= EndDate.Date
                Me.ProgressBar1.Value += 1
                Me.LblProgress.Text = "Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString + " : " + Format(CalcDate, "dd MMM yyyy")
                HitungUlang = False

                'hitung ulang item detail untuk tanggal yang bersangkutan +update inventory pool
                Me.TBLItemDetail.DefaultView.RowFilter = "([tipe item detail] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem).ToString + _
                " OR [tipe item detail] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem).ToString + _
                ") AND [tgl transaksi asal] = #" + Format(CalcDate, "MM/dd/yyyy") + "#"
                counter = 0
                'tandain, boleh hitung transaksi bundling itu lagi jika pada hari sebelumnya tidak berhasil hitung
                ListPrimaryBundling.Clear()

                'If CalcDate = #1/31/2016# Then
                For counter = 0 To Me.TBLItemDetail.DefaultView.Count - 1
                    If Math.IEEERemainder((counter + 1), 50) = 0 Then
                        Me.LblProgress.Text = "Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString + " : " + Format(CalcDate, "dd MMM yyyy") + " : item " + (counter + 1).ToString + " of " + Me.TBLItemDetail.DefaultView.Count.ToString
                    End If

                    'counter = 549
                    Me.rowItemDetail = Me.TBLItemDetail.DefaultView(counter).Row
                    If Me.rowItemDetail.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust Or _
                    Me.rowItemDetail.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InPurchase Or _
                    Me.rowItemDetail.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR Or _
                    Me.rowItemDetail.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InStockTakeAdjust Then
                        'tambahkan inventory pool
                        'asumsi karena tidak mungkin ada item yang double, maka langsung tambahkan saja
                        'tidak ditandai hitung ulang true, karena sudah hitung ulang dibagian ini
                        If CostingMethod = ClsTransaksiMonthlyClosingItem.EnumCostingMethod.Average Then
                            TambahItem_InventoryPool_Avg(Me.rowItemDetail, TBLInvPool, DefPrec)
                        Else
                            Me.TambahItem_InventoryPool_FIFOLIFO1(Me.rowItemDetail, TBLInvPool, DefPrec, CostingMethod)
                        End If

                        'cek apakah ada item yang masih gantung, untuk code no ybs, hitung item-nya
                        TBL_Pending.DefaultView.RowFilter = "[primary code no] = " + Me.rowItemDetail.primary_code_no.ToString + _
                        " AND [tipe item detail] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem).ToString
                        If TBL_Pending.DefaultView.Count > 0 Then
                            TBL_Pending.DefaultView.Sort = "[tgl transaksi asal], [tipe transaksi asal], [primary transaksi asal], [primary code no]"
                            DummyFilter = Me.TBLItemDetail.DefaultView.RowFilter
                            'revisi: 2018-02-02: sort 1x saja
                            'DummySort = Me.'TBLItemDetail.DefaultView.Sort
                            Me.HitungUlangItemPending(TBL_Pending, TBL_Pending.DefaultView.RowFilter, TBLInvPool, CostingMethod, DefPrec, Nothing)
                            Me.TBLItemDetail.DefaultView.RowFilter = DummyFilter
                            'revisi: 2018-02-02: sort 1x saja
                            'Me.'TBLItemDetail.DefaultView.Sort = DummySort
                        End If
                    ElseIf Me.rowItemDetail.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutAdjust Or _
                    Me.rowItemDetail.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutPR Or _
                    Me.rowItemDetail.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutSales Or _
                    Me.rowItemDetail.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutStockTakeAdjust Or _
                    Me.rowItemDetail.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling Then
                        Me.HitungOutItem(TBLInvPool, CostingMethod, TBL_Pending, Me.rowItemDetail, DefPrec, True, Nothing)
                        If Me.rowItemDetail.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling Then
                            'jika tipenya item out bundling dan perhitungannya berhasil, maka coba hitung item in bundling-nya
                            'walaupun tidak berhasil hitung, tetap saja melakukan perhitungan in, supaya item in-nya masuk ke dalam item pending
                            'sehingga urutan FIFO/LIFO tetap terjaga
                            DummyFilter = Me.TBLItemDetail.DefaultView.RowFilter
                            'revisi: 2018-02-02: sort 1x saja
                            'DummySort = Me.'TBLItemDetail.DefaultView.Sort

                            Me.TBLItemDetail.DefaultView.RowFilter = "[tipe transaksi asal] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling).ToString + _
                            " AND [primary transaksi asal] = " + Me.rowItemDetail.primary_transaksi_asal.ToString + _
                            " AND [tipe item detail] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem).ToString

                            If Me.TBLItemDetail.DefaultView.Count > 0 Then
                                rowItemDetail_D1 = Me.TBLItemDetail.DefaultView(0).Row
                                Me.HitungBundling(TBLInvPool, CostingMethod, TBL_Pending, rowItemDetail_D1, DefPrec, True)
                            Else
                                MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                                Exit Sub
                            End If

                            If rowItemDetail_D1.amount <> 0 Then
                                'jika berhasil hitung bundling, maka tandai hari itu boleh hitung ulang
                                HitungUlang = True
                            End If
                            Me.TBLItemDetail.DefaultView.RowFilter = DummyFilter
                            'revisi: 2018-02-02: sort 1x saja
                            'Me.''TBLItemDetail.DefaultView.Sort = DummySort
                        End If
                    End If
                Next

                'hitung ulang semua yang ada di item pending, hanya jika bukan tanggal 1, klo tanggal 1, belum ada pendingan dari hari sebelumnya
                'yang dihitung ulang, hanya pendingan dari hari sebelumnya
                'karena klo dari atas sudah pending, pasti transaksi hari itu pending, ga usah susah2 itung ulang
                If HitungUlang = True Then
                    'hanya hitung ulang pada hari kelipatan 3
                    If Math.IEEERemainder(CalcDate.Date.Day, 4) <> 0 Then
                        HitungUlang = False
                    End If
                End If
                If CalcDate.Date = EndDate.Date Then
                    'jika akhir bulan, pastikan hitung ulang jika tbl_pending ada itemnya
                    HitungUlang = True
                End If
                If TBL_Pending.Rows.Count > 0 And CalcDate.Date <> Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value).Date And HitungUlang = True Then

                    DummyFilter = Me.TBLItemDetail.DefaultView.RowFilter
                    'revisi: 2018-02-02: sort 1x saja
                    'DummySort = Me.'TBLItemDetail.DefaultView.Sort
                    QtyPending1 = TBL_Pending.Rows.Count
                    JmlLoop = 0

                    While TBL_Pending.Rows.Count > 0
                        Me.LblProgress.Text = "Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString + " : " + Format(CalcDate, "dd MMM yyyy") + " : calculating " + TBL_Pending.Rows.Count.ToString + " pending item"
                        JmlLoop += 1
                        Me.HitungUlangItemPending(TBL_Pending, CalcDate, TBLInvPool, CostingMethod, DefPrec)

                        If TBL_Pending.Rows.Count = 0 Then
                            Exit While
                        ElseIf TBL_Pending.Rows.Count = QtyPending1 Then
                            Exit While
                        ElseIf TBL_Pending.Rows.Count > 2000 Then
                            Exit While
                        ElseIf JmlLoop > 2 Then
                            'maksimal lopp 3x, karena dihitung setiap 4 hari
                            Exit While
                        End If
                        QtyPending1 = TBL_Pending.Rows.Count
                    End While
                    Me.TBLItemDetail.DefaultView.RowFilter = DummyFilter
                    'revisi: 2018-02-02: sort 1x saja
                    'Me.''TBLItemDetail.DefaultView.Sort = DummySort
                End If
                'End If

                CalcDate = CalcDate + TimeSpan.FromDays(1)
            End While

            'habiskan item pending
            LoopingPendingWithoutProgress = 0
            QtyPending1 = 0
            QtyPending2 = QtyPending1
            Me.Calc_CurrentStep += 1
            Me.LblStep.Text = "Step " + Me.Calc_CurrentStep.ToString + " of " + Me.Calc_TotalStep.ToString + " : calculating pending data"
            Me.ProgressBar1.Value = 0
            Me.LblProgress.Text = ""
            'Me.LblProgress.Visible = True
            Me.ProgressBar1.Maximum = TBL_Pending.Rows.Count

            jmlawal = Me.TBL_Pending1.Rows.Count
            While LoopingPendingWithoutProgress <= 0 And TBL_Pending.Rows.Count > 0
                Me.ProgressBar1.Value = Me.ProgressBar1.Maximum - TBL_Pending.Rows.Count + 1
                Me.LblProgress.Text = "Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString

                QtyPending1 = TBL_Pending.Rows.Count
                Me.HitungUlangItemPending(TBL_Pending, Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value + TimeSpan.FromDays(32)), TBLInvPool, CostingMethod, DefPrec)
                If QtyPending1 = TBL_Pending.Rows.Count Then
                    LoopingPendingWithoutProgress += 1
                End If
                If TBL_Pending.Rows.Count > 0 Then
                    If SudahPindah = False Then
                        SudahPindah = True
                        Me.LblProgress.Text = "Compacting data, moving to dummy data table..."
                        Me.MasukkanKeDummy(TBLInvPool, TBL_Pending)
                    ElseIf SudahPindah1 = False And TBL_Pending.Rows.Count <= jmlawal * 0.67 Then
                        'pindah lagi pada point 2/3 selesai
                        SudahPindah1 = True
                        Me.LblProgress.Text = "Compacting data, moving to dummy data table..."
                        Me.MasukkanKeDummy(TBLInvPool, TBL_Pending)
                    ElseIf SudahPindah2 = False And TBL_Pending.Rows.Count <= jmlawal * 0.33 Then
                        'pindah lagi pada point 1/3 selesai
                        SudahPindah2 = True
                        Me.LblProgress.Text = "Compacting data, moving to dummy data table..."
                        Me.MasukkanKeDummy(TBLInvPool, TBL_Pending)
                    End If
                End If
            End While

            If TBL_Pending.Rows.Count > 0 Then
                SudahPindah = False
                SudahPindah1 = False
                SudahPindah2 = False
                JmlAwal = TBL_Pending.Rows.Count

                TBLItemClosing.Rows.Clear()
                DTAItemClosing.Fill(TBLItemClosing, Me.dtpPeriod.Value.Date, CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem))

                Me.TBL_Pending1.Rows.Clear()
                For Each row_Pending In TBL_Pending.Rows
                    row_Pending11 = TBL_Pending1.NewtblItemDetailPoolRow
                    For mn As Integer = 0 To TBL_Pending.Columns.Count - 1
                        row_Pending11.Item(mn) = row_Pending.Item(mn)
                    Next
                    TBL_Pending1.Rows.Add(row_Pending11)
                Next
                Me.BgwAktifkanDetail_Progress_4()

                While Me.TBL_Pending2.Rows.Count > 0
                    If TBL_Pending.Rows.Count > 0 Then
                        If SudahPindah = False Then
                            SudahPindah = True
                            Me.LblProgress.Text = "Compacting data, moving to dummy data table..."
                            Me.MasukkanKeDummy(TBLInvPool, TBL_Pending)
                        ElseIf SudahPindah1 = False And TBL_Pending.Rows.Count <= JmlAwal * 0.67 Then
                            'pindah lagi pada point 2/3 selesai
                            SudahPindah1 = True
                            Me.LblProgress.Text = "Compacting data, moving to dummy data table..."
                            Me.MasukkanKeDummy(TBLInvPool, TBL_Pending)
                        ElseIf SudahPindah2 = False And TBL_Pending.Rows.Count <= JmlAwal * 0.33 Then
                            'pindah lagi pada point 1/3 selesai
                            SudahPindah2 = True
                            Me.LblProgress.Text = "Compacting data, moving to dummy data table..."
                            Me.MasukkanKeDummy(TBLInvPool, TBL_Pending)
                        End If
                    End If

                    'hitung item pending yang diberikan unit costnya
                    TBL_Pending.DefaultView.Sort = "[tgl transaksi asal], [tipe transaksi asal], [primary transaksi asal], [primary code no]"

                    For Each row_Pending2 In Me.TBL_Pending2.Rows
                        TBL_Pending.DefaultView.RowFilter = "[primary code no] = " + row_Pending2.primary_code_no.ToString + _
                        " AND [tipe item detail] = " + CInt(row_Pending2.tipe_item_detail).ToString + _
                        " AND [tipe transaksi asal] = " + CInt(row_Pending2.tipe_transaksi_asal).ToString

                        If TBL_Pending.DefaultView.Count > 0 Then
                            Me.HitungUlangItemPending(TBL_Pending, TBL_Pending.DefaultView.RowFilter, TBLInvPool, CostingMethod, DefPrec, row_Pending2)
                        End If
                    Next
                    For ij As Integer = Me.TBL_Pending2.Rows.Count - 1 To 0 Step -1
                        row_Pending2 = Me.TBL_Pending2.Rows(ij)

                        TBL_Pending.DefaultView.RowFilter = "[primary code no] = " + row_Pending2.primary_code_no.ToString + _
                        " AND [tipe item detail] = " + CInt(row_Pending2.tipe_item_detail).ToString + _
                        " AND [tipe transaksi asal] = " + CInt(row_Pending2.tipe_transaksi_asal).ToString + _
                        " AND [Primary Transaksi Asal] = " + row_Pending2.Primary_Transaksi_Asal.ToString

                        If TBL_Pending.DefaultView.Count = 0 Then
                            Me.TBL_Pending2.Rows.Remove(row_Pending2)
                        End If
                    Next

                    LoopingPendingWithoutProgress = 0
                    While LoopingPendingWithoutProgress <= 0 And TBL_Pending.Rows.Count > 0
                        'Me.ProgressBar1.Value = Me.ProgressBar1.Maximum - TBL_Pending.Rows.Count + 1
                        Me.LblProgress.Text = "Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString

                        QtyPending1 = TBL_Pending.Rows.Count
                        Me.HitungUlangItemPending(TBL_Pending, Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value.Date + TimeSpan.FromDays(32)), TBLInvPool, CostingMethod, DefPrec)
                        If QtyPending1 = TBL_Pending.Rows.Count Then
                            LoopingPendingWithoutProgress += 1
                        End If
                    End While

                    Me.TBL_Pending1.Rows.Clear()

                    For Each row_Pending In TBL_Pending.Rows
                        row_Pending11 = TBL_Pending1.NewtblItemDetailPoolRow
                        For mn As Integer = 0 To TBL_Pending.Columns.Count - 1
                            row_Pending11.Item(mn) = row_Pending.Item(mn)
                        Next
                        TBL_Pending1.Rows.Add(row_Pending11)
                    Next
                    If TBL_Pending1.Rows.Count > 0 Then
                        Me.BgwAktifkanDetail_Progress_4()
                    End If
                End While

                If Me.TBLItemDetail_Dummy.Rows.Count > 0 Or Me.TBLInvPool_Dummy.Rows.Count > 0 Then
                    'Me.LblProgress.Text = "Restoring data, moving to source data table..."
                    'StartTime = Now
                    Me.MasukkanKeTabelAsal(TBLInvPool)
                    'EndTime = Now
                    'MsgBox("Restoring from dummy: " + Me.ObjCD.DigitGrouping((EndTime - StartTime).TotalMinutes, 1) + " minutes")
                End If

                If TBL_Pending.Rows.Count > 0 Then
                    Me.TBL_Pending1.Rows.Clear()
                    Me.TBLItemDetail.Rows.Clear() '-> pindahan dari fungsi Me.BgwAktifkanDetail_Progress_4, dipindahkan supaya mudah kalo mau melacak error siklik bundling ketika mau closing

                    For Each row_Pending In TBL_Pending.Rows
                        row_Pending11 = TBL_Pending1.NewtblItemDetailPoolRow
                        For mn As Integer = 0 To TBL_Pending.Columns.Count - 1
                            row_Pending11.Item(mn) = row_Pending.Item(mn)
                        Next
                        TBL_Pending1.Rows.Add(row_Pending11)
                    Next
                End If
            Else
                If Me.TBLItemDetail_Dummy.Rows.Count > 0 Or Me.TBLInvPool_Dummy.Rows.Count > 0 Then
                    'Me.LblProgress.Text = "Restoring data, moving to source data table..."
                    'StartTime = Now
                    Me.MasukkanKeTabelAsal(TBLInvPool)
                    'EndTime = Now
                    'MsgBox("Restoring from dummy: " + Me.ObjCD.DigitGrouping((EndTime - StartTime).TotalMinutes, 1) + " minutes") 
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message + " " + Format(CalcDate, "dd MMM yyyy"), MsgBoxStyle.Critical, "Error")
        Finally
            TBL_Pending = Nothing
            ListPrimaryBundling = Nothing
            ObjIdentifier = Nothing
        End Try
    End Sub

    Private Sub MasukkanKeDummy1(ByVal TBLInvPool As DseMonthlyClosingNew.tblItemDetailPoolDataTable)
        Dim RowInvPool, RowInvPool_D As DseMonthlyClosingNew.tblItemDetailPoolRow
        Dim rowItemDetail_Dummy, rowItemDetail As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailRow
        Dim DTA As New DseTransaksiMonthlyClosing3TableAdapters.DataTable2TableAdapter
        Dim TBL As New DseTransaksiMonthlyClosing3.DataTable2DataTable
        Dim row As DseTransaksiMonthlyClosing3.DataTable2Row
        Dim JmlAwal, JmlReduce As Decimal

        Dim DTA1 As New DseTransaksiMonthlyClosing3TableAdapters.tbltrans_purchase_closing_item_detailTableAdapter
        Dim TBL1 As New DseTransaksiMonthlyClosing3.tbltrans_purchase_closing_item_detailDataTable
        Dim row1 As DseTransaksiMonthlyClosing3.tbltrans_purchase_closing_item_detailRow
        Dim PrimaryClosingSebelum As Double = -1

        Try
            'persiapan data step 2
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString
            Me.LblProgress.Text = "Preparing data"
            DTA.FillBy(TBL, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value.Date), Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value.Date))
            PrimaryClosingSebelum = Me.ObjClosing.GetPrimary(Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value.Date - TimeSpan.FromDays(2)))
            DTA1.Fill(TBL1, PrimaryClosingSebelum, ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.EndingBalanceItem)
            JmlAwal = Me.TBLItemDetail.Rows.Count + TBLInvPool.Rows.Count
            Me.LblProgress.Text = "Preparing data"

            'step 1: pindahkan semua begining balance
            Me.ProgressBar1.Value = 0
            Me.ProgressBar1.Maximum = Me.TBLItemDetail.Rows.Count
            For sd As Double = Me.TBLItemDetail.Rows.Count - 1 To 0 Step -1
                Me.ProgressBar1.Value += 1
                If Me.ProgressBar1.Value > 25 Then
                    If Math.IEEERemainder(Me.ProgressBar1.Value, 100) = 0 Then
                        Me.LblProgress.Text = "Step 1 of 2: begining balance - data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString + " - " + _
                        Me.ObjCD.DigitGrouping(JmlReduce, 0) + " (" + Me.ObjCD.DigitGrouping((JmlReduce * 100) / JmlAwal, 2) + "%)" + " moved"
                    ElseIf Me.ProgressBar1.Value = Me.ProgressBar1.Maximum Then
                        Me.LblProgress.Text = "Step 1 of 2: begining balance - data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString + " - " + _
                        Me.ObjCD.DigitGrouping(JmlReduce, 0) + " (" + Me.ObjCD.DigitGrouping((JmlReduce * 100) / JmlAwal, 2) + "%)" + " moved"
                    End If
                Else
                    Me.LblProgress.Text = "Step 1 of 2: begining balance - data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString + " - " + _
                        Me.ObjCD.DigitGrouping(JmlReduce, 0) + " (" + Me.ObjCD.DigitGrouping((JmlReduce * 100) / JmlAwal, 2) + "%)" + " moved"
                End If

                rowItemDetail = Me.TBLItemDetail.Rows(sd)
                If rowItemDetail.tipe_item_detail = CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.BeginingBalanceItem) Then
                    rowItemDetail_Dummy = Me.TBLItemDetail_Dummy.Newtbltrans_purchase_closing_item_detailRow
                    For gh As Integer = 0 To Me.TBLItemDetail.Columns.Count - 1
                        rowItemDetail_Dummy.Item(gh) = rowItemDetail.Item(gh)
                    Next
                    Me.TBLItemDetail_Dummy.Rows.Add(rowItemDetail_Dummy)
                    Me.TBLItemDetail.Rows.Remove(rowItemDetail)
                    JmlReduce += 1
                End If
            Next

            'step 2: pindahkan semua yang ga ada transaksi pada bulan ybs
            'syarat: ada di end item detail bulan lalu dan bulan ini tidak ada stock card sama sekali
            Me.ProgressBar1.Value = 0
            Me.ProgressBar1.Maximum = TBL1.Rows.Count
            For Each row1 In TBL1.Rows
                Me.ProgressBar1.Value += 1
                TBL.DefaultView.RowFilter = "[Primary Product] = " + row1.primary_code_no.ToString

                If TBL.DefaultView.Count > 0 Then
                    row = TBL.DefaultView(0).Row
                    If Me.ProgressBar1.Value > 25 Then
                        If Math.IEEERemainder(Me.ProgressBar1.Value, 100) = 0 Then
                            Me.LblProgress.Text = "Step 2 of 2: inactive code no - data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString + " - " + _
                            Me.ObjCD.DigitGrouping(JmlReduce, 0) + " (" + Me.ObjCD.DigitGrouping((JmlReduce * 100) / JmlAwal, 2) + "%)" + " moved"
                        ElseIf Me.ProgressBar1.Value = Me.ProgressBar1.Maximum Then
                            Me.LblProgress.Text = "Step 2 of 2: inactive code no - data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString + " - " + _
                            Me.ObjCD.DigitGrouping(JmlReduce, 0) + " (" + Me.ObjCD.DigitGrouping((JmlReduce * 100) / JmlAwal, 2) + "%)" + " moved"
                        End If
                    Else
                        Me.LblProgress.Text = "Step 2 of 2: inactive code no - data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString + " - " + _
                            Me.ObjCD.DigitGrouping(JmlReduce, 0) + " (" + Me.ObjCD.DigitGrouping((JmlReduce * 100) / JmlAwal, 2) + "%)" + " moved"
                    End If

                    '1. item detail
                    Me.TBLItemDetail.DefaultView.RowFilter = "[primary code no] = " + row.Primary_Product.ToString
                    JmlReduce += Me.TBLItemDetail.DefaultView.Count
                    While Me.TBLItemDetail.DefaultView.Count > 0
                        rowItemDetail = Me.TBLItemDetail.DefaultView(0).Row
                        rowItemDetail_Dummy = Me.TBLItemDetail_Dummy.Newtbltrans_purchase_closing_item_detailRow
                        For gh As Integer = 0 To Me.TBLItemDetail.Columns.Count - 1
                            rowItemDetail_Dummy.Item(gh) = rowItemDetail.Item(gh)
                        Next
                        Me.TBLItemDetail_Dummy.Rows.Add(rowItemDetail_Dummy)
                        Me.TBLItemDetail.Rows.Remove(rowItemDetail)
                    End While

                    '2. inventory pool
                    TBLInvPool.DefaultView.RowFilter = "[primary code no] = " + row.Primary_Product.ToString
                    JmlReduce += TBLInvPool.DefaultView.Count
                    While TBLInvPool.DefaultView.Count > 0
                        RowInvPool = TBLInvPool.DefaultView(0).Row
                        RowInvPool_D = Me.TBLInvPool_Dummy.NewtblItemDetailPoolRow
                        For gh As Integer = 0 To Me.TBLInvPool_Dummy.Columns.Count - 1
                            RowInvPool_D.Item(gh) = RowInvPool.Item(gh)
                        Next
                        Me.TBLInvPool_Dummy.Rows.Add(RowInvPool_D)
                        TBLInvPool.Rows.Remove(RowInvPool)
                    End While
                End If
            Next 
            Me.LblProgress.Text = ""
            Me.ProgressBar1.Value = 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Rows.Clear()
            TBL.Dispose()
            DTA1.Dispose()
            TBL1.Rows.Clear()
            TBL1.Dispose()
        End Try
    End Sub

    Private Sub MasukkanKeDummy(ByVal TBLInvPool As DseMonthlyClosingNew.tblItemDetailPoolDataTable, ByVal TBL_Pending As DseMonthlyClosingNew.tblItemDetailPoolDataTable)
        Dim ListProductPindah As New List(Of Double)
        Dim ListProductPending As New List(Of Double)
        Dim ListProductSudahCek As New List(Of Double)
        Dim ListPrimaryBundling As New List(Of Double)
        Dim RowInvPool, RowInvPool_D, RowPending As DseMonthlyClosingNew.tblItemDetailPoolRow
        Dim rowItemDetail_Dummy, rowItemDetail As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailRow
        Dim PrimaryMax, PrimaryMin As Double
        Dim DTAItem1 As New DseTransaksiMonthlyClosing3TableAdapters.tbltransaksi_stockbundling_item_afterTableAdapter
        Dim TBLItem1 As New DseTransaksiMonthlyClosing3.tbltransaksi_stockbundling_item_afterDataTable
        Dim rowItem1 As DseTransaksiMonthlyClosing3.tbltransaksi_stockbundling_item_afterRow
        Dim Jml, JmlTotal As Integer
        Dim JmlAwal, JmlReduce As Decimal

        Try
            DTAItem1.Connection.ConnectionString = Core_Module.GetConstring

            '(nanti) 1. simpan tabel inventory pool, tbl item detail, dan tbl pending ke dalam db -> ini supaya checking data selanjutnya lebih mudah

            '2. untuk setiap row pending, jika tipenya:
            PrimaryMin = Double.MaxValue
            PrimaryMax = Double.MinValue
            Me.LblProgress.Text = "Preparing data"
            For Each RowPending In TBL_Pending.Rows
                If RowPending.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling Or _
                RowPending.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling Then
                    If PrimaryMax < RowPending.Primary_Transaksi_Asal Then
                        PrimaryMax = RowPending.Primary_Transaksi_Asal
                    End If
                    If PrimaryMin > RowPending.Primary_Transaksi_Asal Then
                        PrimaryMin = RowPending.Primary_Transaksi_Asal
                    End If
                End If
            Next
            'PrimaryMin = PrimaryMax
            If PrimaryMax <> Double.MinValue Then
                DTAItem1.Fill(TBLItem1, PrimaryMax, PrimaryMin, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value.Date), Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value))
            End If

            ListPrimaryBundling.Clear()


            Me.LblProgress.Text = "Collecting calculated code no"
            Jml = 0
            JmlTotal = TBL_Pending.Rows.Count
            For Each RowPending In TBL_Pending.Rows
                Jml += 1
                If Jml > 25 Then
                    If Math.IEEERemainder(Jml, 100) = 0 Then
                        Me.LblProgress.Text = "Step 1 of 3: item detail - data " + Jml.ToString + " of " + JmlTotal.ToString
                    ElseIf Jml = JmlTotal Then
                        Me.LblProgress.Text = "Step 1 of 3: item detail - data " + Jml.ToString + " of " + JmlTotal.ToString
                    End If
                Else
                    Me.LblProgress.Text = "Step 1 of 3: item detail - data " + Jml.ToString + " of " + JmlTotal.ToString
                End If

                If RowPending.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling Or _
                RowPending.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling Then
                    If ListPrimaryBundling.Contains(RowPending.Primary_Transaksi_Asal) = False Then
                        ListPrimaryBundling.Add(RowPending.Primary_Transaksi_Asal)
                        '2.1. in bundling + out bundling: 
                        '   a. kumpulkan semua product dibawah bunding tersebut
                        '   b. pastikan setiap product tersebut ada di list product pending
                        TBLItem1.DefaultView.RowFilter = "[Primary Transaksi] = " + RowPending.Primary_Transaksi_Asal.ToString
                        While TBLItem1.DefaultView.Count > 0
                            rowItem1 = TBLItem1.DefaultView(0).Row
                            If ListProductSudahCek.Contains(rowItem1.Primary_Product) = False Then
                                ListProductSudahCek.Add(rowItem1.Primary_Product)
                                ListProductPending.Add(rowItem1.Primary_Product)
                            End If
                            TBLItem1.Rows.Remove(rowItem1)
                        End While
                    End If
                Else
                    '2.2. lainnya: pastikan product tersebut ada di list product pending
                    If ListProductSudahCek.Contains(RowPending.primary_code_no) = False Then
                        ListProductSudahCek.Add(RowPending.primary_code_no)
                        ListProductPending.Add(RowPending.primary_code_no)
                    End If
                End If
            Next

            '3. untuk setiap row inventory pool pastikan primary productnya belum di cek, lalu bandingkan dengan primary product pending
            '   kalo tidak ada di 2 list tersebut, maka masukkan ke dalam list pindah
            Me.LblProgress.Text = "Collecting calculated code no"
            Jml = 0
            JmlTotal = TBLInvPool.Rows.Count
            For Each RowInvPool In TBLInvPool.Rows
                Jml += 1
                If Jml > 25 Then
                    If Math.IEEERemainder(Jml, 100) = 0 Then
                        Me.LblProgress.Text = "Step 2 of 3: inventory pool - data " + Jml.ToString + " of " + JmlTotal.ToString
                    ElseIf Jml = JmlTotal Then
                        Me.LblProgress.Text = "Step 2 of 3: inventory pool - data " + Jml.ToString + " of " + JmlTotal.ToString
                    End If
                Else
                    Me.LblProgress.Text = "Step 2 of 3: inventory pool - data " + Jml.ToString + " of " + JmlTotal.ToString
                End If

                If ListProductSudahCek.Contains(RowInvPool.primary_code_no) = False Then
                    ListProductSudahCek.Add(RowInvPool.primary_code_no)
                    If ListProductPending.Contains(RowInvPool.primary_code_no) = False Then
                        ListProductPindah.Add(RowInvPool.primary_code_no)
                    Else
                        MsgBox("")
                    End If
                End If
            Next

            '4. untuk setiap baris pada list pindah, pindahkan ke tabel dummy
            'untuk sisa primary product, pindahkan datanya ke dalam dummy
            Me.LblProgress.Text = "Moving data item"
            Jml = 0
            JmlTotal = ListProductPindah.Count
            JmlAwal = Me.TBLItemDetail.Rows.Count + TBLInvPool.Rows.Count
            For Each PrimaryProduct_D As Double In ListProductPindah
                Jml += 1
                If Jml > 25 Then
                    If Math.IEEERemainder(Jml, 100) = 0 Then
                        Me.LblProgress.Text = "Step 3 of 3: moving - data " + Jml.ToString + " of " + JmlTotal.ToString + " - " + _
                        Me.ObjCD.DigitGrouping(JmlReduce, 0) + " (" + Me.ObjCD.DigitGrouping((JmlReduce * 100) / JmlAwal, 2) + "%)" + " moved"
                    ElseIf Jml = JmlTotal Then
                        Me.LblProgress.Text = "Step 3 of 3: moving - data " + Jml.ToString + " of " + JmlTotal.ToString + " - " + _
                        Me.ObjCD.DigitGrouping(JmlReduce, 0) + " (" + Me.ObjCD.DigitGrouping((JmlReduce * 100) / JmlAwal, 2) + "%)" + " moved"
                    End If
                Else
                    Me.LblProgress.Text = "Step 3 of 3: moving - data " + Jml.ToString + " of " + JmlTotal.ToString + " - " + _
                        Me.ObjCD.DigitGrouping(JmlReduce, 0) + " (" + Me.ObjCD.DigitGrouping((JmlReduce * 100) / JmlAwal, 2) + "%)" + " moved"
                End If

                '1. item detail
                Me.TBLItemDetail.DefaultView.RowFilter = "[primary code no] = " + PrimaryProduct_D.ToString
                While Me.TBLItemDetail.DefaultView.Count > 0
                    JmlReduce += 1
                    rowItemDetail = Me.TBLItemDetail.DefaultView(0).Row
                    rowItemDetail_Dummy = Me.TBLItemDetail_Dummy.Newtbltrans_purchase_closing_item_detailRow
                    For gh As Integer = 0 To Me.TBLItemDetail.Columns.Count - 1
                        rowItemDetail_Dummy.Item(gh) = rowItemDetail.Item(gh)
                    Next
                    Me.TBLItemDetail_Dummy.Rows.Add(rowItemDetail_Dummy)
                    Me.TBLItemDetail.Rows.Remove(rowItemDetail)
                End While

                '2. inventory pool
                TBLInvPool.DefaultView.RowFilter = "[primary code no] = " + PrimaryProduct_D.ToString
                While TBLInvPool.DefaultView.Count > 0
                    JmlReduce += 1
                    RowInvPool = TBLInvPool.DefaultView(0).Row
                    RowInvPool_D = Me.TBLInvPool_Dummy.NewtblItemDetailPoolRow
                    For gh As Integer = 0 To Me.TBLInvPool_Dummy.Columns.Count - 1
                        RowInvPool_D.Item(gh) = RowInvPool.Item(gh)
                    Next
                    Me.TBLInvPool_Dummy.Rows.Add(RowInvPool_D)
                    TBLInvPool.Rows.Remove(RowInvPool)
                End While
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListProductPindah.Clear()
            ListProductPindah = Nothing
            ListProductPending.Clear()
            ListProductPending = Nothing
            ListProductSudahCek.Clear()
            ListProductSudahCek = Nothing
            ListPrimaryBundling.Clear()
            ListPrimaryBundling = Nothing
            DTAItem1.Dispose()
            TBLItem1.Rows.Clear()
            TBLItem1.Dispose()
        End Try
    End Sub

    'Private Sub MasukkanKeDummy1(ByVal TBLInvPool As DseMonthlyClosingNew.tblItemDetailPoolDataTable, ByVal TBL_Pending As DseMonthlyClosingNew.tblItemDetailPoolDataTable)
    '    Dim ListPrimary1 As New List(Of Double)
    '    Dim RowInvPool, RowInvPool_D, RowPending As DseMonthlyClosingNew.tblItemDetailPoolRow
    '    Dim rowItemDetail_Dummy, rowItemDetail As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailRow

    '    Try
    '        'daftar semua primary product yang ada di inv. pool
    '        For Each RowInvPool In TBLInvPool.Rows
    '            If ListPrimary1.Contains(RowInvPool.primary_code_no) = False Then
    '                ListPrimary1.Add(RowInvPool.primary_code_no)
    '            End If
    '        Next

    '        'buang primary product yang ada di pending dan di inv. pool
    '        For Each RowPending In TBL_Pending.Rows
    '            If ListPrimary1.Contains(RowPending.primary_code_no) = True Then
    '                ListPrimary1.Remove(RowPending.primary_code_no)
    '            End If
    '        Next

    '        'untuk sisa primary product, pindahkan datanya ke dalam dummy
    '        For Each PrimaryProduct_D As Double In ListPrimary1
    '            '1. item detail
    '            Me.TBLItemDetail.DefaultView.RowFilter = "[primary code no] = " + PrimaryProduct_D.ToString
    '            While Me.TBLItemDetail.DefaultView.Count > 0
    '                rowItemDetail = Me.TBLItemDetail.DefaultView(0).Row
    '                rowItemDetail_Dummy = Me.TBLItemDetail_Dummy.Newtbltrans_purchase_closing_item_detailRow
    '                For gh As Integer = 0 To Me.TBLItemDetail.Columns.Count - 1
    '                    rowItemDetail_Dummy.Item(gh) = rowItemDetail.Item(gh)
    '                Next
    '                Me.TBLItemDetail_Dummy.Rows.Add(rowItemDetail_Dummy)
    '                Me.TBLItemDetail.Rows.Remove(rowItemDetail)
    '            End While

    '            '2. inventory pool
    '            TBLInvPool.DefaultView.RowFilter = "[primary code no] = " + PrimaryProduct_D.ToString
    '            While TBLInvPool.DefaultView.Count > 0
    '                RowInvPool = TBLInvPool.DefaultView(0).Row
    '                RowInvPool_D = Me.TBLInvPool_Dummy.NewtblItemDetailPoolRow
    '                For gh As Integer = 0 To Me.TBLInvPool_Dummy.Columns.Count - 1
    '                    RowInvPool_D.Item(gh) = RowInvPool.Item(gh)
    '                Next
    '                Me.TBLInvPool_Dummy.Rows.Add(RowInvPool_D)
    '                TBLInvPool.Rows.Remove(RowInvPool)
    '            End While
    '        Next
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        ListPrimary1 = Nothing
    '    End Try
    'End Sub

    Private Sub MasukkanKeTabelAsal(ByVal TBLInvPool As DseMonthlyClosingNew.tblItemDetailPoolDataTable)
        Dim RowInvPool, RowInvPool_D As DseMonthlyClosingNew.tblItemDetailPoolRow
        Dim rowItemDetail_Dummy, rowItemDetail As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailRow

        Try
            'item detail
            For Each RowInvPool_D In Me.TBLInvPool_Dummy.Rows
                RowInvPool = TBLInvPool.NewtblItemDetailPoolRow
                For gh As Integer = 0 To Me.TBLInvPool_Dummy.Columns.Count - 1
                    RowInvPool.Item(gh) = RowInvPool_D.Item(gh)
                Next
                TBLInvPool.Rows.Add(RowInvPool)
            Next

            'inv pool
            For Each rowItemDetail_Dummy In Me.TBLItemDetail_Dummy.Rows
                rowItemDetail = Me.TBLItemDetail.Newtbltrans_purchase_closing_item_detailRow
                For gh As Integer = 0 To Me.TBLItemDetail.Columns.Count - 1
                    rowItemDetail.Item(gh) = rowItemDetail_Dummy.Item(gh)
                Next
                Me.TBLItemDetail.Rows.Add(rowItemDetail)
            Next

            Me.TBLItemDetail_Dummy.Rows.Clear()
            Me.TBLInvPool_Dummy.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HitungUlangItemPending(ByVal TBL_Pending As DseMonthlyClosingNew.tblItemDetailPoolDataTable, ByVal vRefDate As Date, _
    ByVal TBLInvPool As DseMonthlyClosingNew.tblItemDetailPoolDataTable, _
    ByVal CostingMethod As ClsTransaksiMonthlyClosingItem.EnumCostingMethod, ByVal vDefPrec As Integer)
        Dim TblCek As New DseMonthlyClosingNew.DataTable6DataTable
        Dim rowCek As DseMonthlyClosingNew.DataTable6Row
        Dim row_Pending As DseMonthlyClosingNew.tblItemDetailPoolRow
        Dim Counter As Double = 0

        Try
            TblCek.Rows.Clear()

            Counter = 0
            For Each row_Pending In TBL_Pending.Rows
                Counter += 1
                If Math.IEEERemainder(Counter, 50) = 0 Then
                    Me.LblProgress.Text = "Collecting pending data: " + Counter.ToString + " of " + TBL_Pending.Rows.Count.ToString
                End If
                'yang dihitung ulang hanya transaksi tanggal sebelumnya
                If row_Pending.tgl_transaksi_asal.Date < vRefDate.Date Then
                    TblCek.DefaultView.RowFilter = "[Primary Code No] = " + row_Pending.primary_code_no.ToString + _
                    " AND [Tipe Item Detail] = " + row_Pending.tipe_item_detail.ToString
                    If TblCek.DefaultView.Count = 0 Then
                        rowCek = TblCek.NewDataTable6Row
                        rowCek.Primary_Code_No = row_Pending.primary_code_no
                        rowCek.Tipe_Item_Detail = row_Pending.tipe_item_detail
                        TblCek.Rows.Add(rowCek)
                    End If
                End If
            Next

            TBL_Pending.DefaultView.Sort = "[tgl transaksi asal], [tipe transaksi asal], [primary transaksi asal], [primary code no]"
            Counter = 0
            For Each rowCek In TblCek.Rows
                Counter += 1
                If Math.IEEERemainder(Counter, 50) = 0 Then
                    Me.LblProgress.Text = "Calculating data: " + Counter.ToString + " of " + TblCek.Rows.Count.ToString + _
                    " - Remaining Pending: " + TBL_Pending.Rows.Count.ToString
                End If

                TBL_Pending.DefaultView.RowFilter = "[primary code no] = " + rowCek.Primary_Code_No.ToString + _
                " AND [tipe item detail] = " + CInt(rowCek.Tipe_Item_Detail).ToString

                If TBL_Pending.DefaultView.Count > 0 Then
                    Me.HitungUlangItemPending(TBL_Pending, TBL_Pending.DefaultView.RowFilter, TBLInvPool, CostingMethod, vDefPrec, Nothing)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TblCek = Nothing
        End Try
    End Sub

    Private Sub HitungUlangItemPending1(ByVal TBL_Pending As DseMonthlyClosingNew.tblItemDetailPoolDataTable, ByVal vRefDate As Date, _
   ByVal TBLInvPool As DseMonthlyClosingNew.tblItemDetailPoolDataTable, _
   ByVal CostingMethod As ClsTransaksiMonthlyClosingItem.EnumCostingMethod, ByVal vDefPrec As Integer)
        Dim TblCek As New DseMonthlyClosingNew.DataTable6DataTable
        Dim rowCek As DseMonthlyClosingNew.DataTable6Row
        Dim row_Pending As DseMonthlyClosingNew.tblItemDetailPoolRow
        Dim Counter As Double = 0

        Try
            TblCek.Rows.Clear()

            Counter = 0
            For Each row_Pending In TBL_Pending.Rows
                Counter += 1
                If Math.IEEERemainder(Counter, 50) = 0 Then
                    Me.LblProgress.Text = "Collecting pending data: " + Counter.ToString + " of " + TBL_Pending.Rows.Count.ToString
                End If
                'yang dihitung ulang hanya transaksi tanggal sebelumnya
                If row_Pending.tgl_transaksi_asal.Date < vRefDate.Date Then
                    TblCek.DefaultView.RowFilter = "[Primary Code No] = " + row_Pending.primary_code_no.ToString + _
                    " AND [Tipe Item Detail] = " + row_Pending.tipe_item_detail.ToString
                    If TblCek.DefaultView.Count = 0 Then
                        rowCek = TblCek.NewDataTable6Row
                        rowCek.Primary_Code_No = row_Pending.primary_code_no
                        rowCek.Tipe_Item_Detail = row_Pending.tipe_item_detail
                        TblCek.Rows.Add(rowCek)
                    End If
                End If
            Next

            TBL_Pending.DefaultView.Sort = "[tgl transaksi asal], [tipe transaksi asal], [primary transaksi asal], [primary code no]"
            Counter = 0
            For Each rowCek In TblCek.Rows
                Counter += 1
                If Math.IEEERemainder(Counter, 50) = 0 Then
                    Me.LblProgress.Text = "Calculating data: " + Counter.ToString + " of " + TblCek.Rows.Count.ToString + _
                    " - Remaining Pending: " + TBL_Pending.Rows.Count.ToString
                End If

                TBL_Pending.DefaultView.RowFilter = "[primary code no] = " + rowCek.Primary_Code_No.ToString + _
                " AND [tipe item detail] = " + CInt(rowCek.Tipe_Item_Detail).ToString

                If TBL_Pending.DefaultView.Count > 0 Then
                    Me.HitungUlangItemPending1(TBL_Pending, TBL_Pending.DefaultView.RowFilter, TBLInvPool, CostingMethod, vDefPrec)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TblCek = Nothing
        End Try
    End Sub

    Private Sub HitungUlangItemPending(ByVal TBL_Pending As DseMonthlyClosingNew.tblItemDetailPoolDataTable, _
    ByVal vFilterPending As String, _
    ByVal TBLInvPool As DseMonthlyClosingNew.tblItemDetailPoolDataTable, _
    ByVal CostingMethod As ClsTransaksiMonthlyClosingItem.EnumCostingMethod, ByVal vDefPrec As Integer, _
    ByVal vRow_Pending2 As DseMonthlyClosingNew.tblItemDetailPoolRow)
        'klo manggil fungsi ini, sebelumnya harus mencatat dummy sort dan dummy row filter untuk me.tblitemdetail.defaultview
        Dim rowPending As DseMonthlyClosingNew.tblItemDetailPoolRow
        Dim rowItemDetail_D As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailRow
        Dim rowItemDetail_D1 As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailRow
        'Dim rowItemDetail_D2 As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailRow
        Dim counter As Integer = 0
        Dim RefNumber As Integer = 0

        Try
            While TBL_Pending.DefaultView.Count > counter
                RefNumber = TBL_Pending.Rows.Count

                rowPending = TBL_Pending.DefaultView(counter).Row
                'cari padanan row item detailnya
                Me.TBLItemDetail.DefaultView.RowFilter = "[tipe transaksi asal] = " + rowPending.tipe_transaksi_asal.ToString + _
                " AND [primary transaksi asal] = " + rowPending.Primary_Transaksi_Asal.ToString + _
                " AND [primary code no] = " + rowPending.primary_code_no.ToString
                If Me.TBLItemDetail.DefaultView.Count = 0 Then
                    MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
                rowItemDetail_D = Me.TBLItemDetail.DefaultView(0).Row
                If rowItemDetail_D.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutAdjust Or _
                rowItemDetail_D.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutPR Or _
                rowItemDetail_D.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutSales Or _
                rowItemDetail_D.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutStockTakeAdjust Or _
                rowItemDetail_D.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling Then
                    'hitung rowitem detailnya, pastikan tidak menambahkan lagi ke item pending, karena ini memang
                    'perhitungan yang berasal dari item pending
                    Me.HitungOutItem(TBLInvPool, CostingMethod, TBL_Pending, rowItemDetail_D, vDefPrec, False, vRow_Pending2)
                    'jika tipe itemnya adalah out bundling, maka coba hitung in bundlingnya
                    'walaupun tidak berhasil hitung, tetap saja melakukan perhitungan in, supaya item in-nya masuk ke dalam item pending
                    'sehingga urutan FIFO/LIFO tetap terjaga
                    'berbeda dengan ketika pada fungsi hitung cost yang berlaku seperti di atas,
                    'pada fungsi ini menghitung in bundling, hanya jika hitung outnya berhasil
                    'karena pada point ini yang terjadi adalah hitung ulang sehingga urutan item FIFO-nya sudah bener
                    If rowItemDetail_D.amount <> 0 Then
                        If rowItemDetail_D.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling Then
                            Me.TBLItemDetail.DefaultView.RowFilter = "[tipe transaksi asal] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling).ToString + _
                            " AND [primary transaksi asal] = " + rowItemDetail_D.primary_transaksi_asal.ToString + _
                            " AND [tipe item detail] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem).ToString
                            If Me.TBLItemDetail.DefaultView.Count > 0 Then
                                rowItemDetail_D1 = Me.TBLItemDetail.DefaultView(0).Row
                                Me.HitungBundling(TBLInvPool, CostingMethod, TBL_Pending, rowItemDetail_D1, vDefPrec, False)
                            Else
                                MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                                Exit Sub
                            End If
                        End If
                        'Else
                        'MsgBox("")
                    End If
                Else
                    'hitung item bundling, pastikan tidak menambahkan lagi ke item pending, karena ini memang
                    'perhitungan yang berasal dari item pending
                    Me.HitungBundling(TBLInvPool, CostingMethod, TBL_Pending, rowItemDetail_D, vDefPrec, False)
                End If

                'jika perhitungan berhasil, coba hitung item pending selanjutnya, jika ada
                If rowItemDetail_D.amount = 0 Then
                    If rowPending.tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem Then
                        'amount =0, artinya di pending kembali, klo out item, maka bawahnya yang out item juga
                        'pasti juga ga kebagian, maka langsung exit saja
                        Exit While
                    Else
                        If RefNumber = TBL_Pending.Rows.Count Then
                            counter += 1
                        End If
                    End If
                Else
                    TBL_Pending.DefaultView.RowFilter = vFilterPending
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HitungUlangItemPending1(ByVal TBL_Pending As DseMonthlyClosingNew.tblItemDetailPoolDataTable, _
    ByVal vFilterPending As String, _
    ByVal TBLInvPool As DseMonthlyClosingNew.tblItemDetailPoolDataTable, _
    ByVal CostingMethod As ClsTransaksiMonthlyClosingItem.EnumCostingMethod, ByVal vDefPrec As Integer)
        'klo manggil fungsi ini, sebelumnya harus mencatat dummy sort dan dummy row filter untuk me.tblitemdetail.defaultview
        Dim rowPending As DseMonthlyClosingNew.tblItemDetailPoolRow
        Dim rowItemDetail_D As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailRow
        Dim rowItemDetail_D1 As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailRow
        'Dim rowItemDetail_D2 As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailRow
        Dim counter As Integer = 0
        Dim RefNumber As Integer = 0

        Try
            While TBL_Pending.DefaultView.Count > counter
                RefNumber = TBL_Pending.Rows.Count

                rowPending = TBL_Pending.DefaultView(counter).Row
                'cari padanan row item detailnya
                Me.TBLItemDetail.DefaultView.RowFilter = "[tipe transaksi asal] = " + rowPending.tipe_transaksi_asal.ToString + _
                " AND [primary transaksi asal] = " + rowPending.Primary_Transaksi_Asal.ToString + _
                " AND [primary code no] = " + rowPending.primary_code_no.ToString
                If Me.TBLItemDetail.DefaultView.Count = 0 Then
                    MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
                rowItemDetail_D = Me.TBLItemDetail.DefaultView(0).Row
                If rowItemDetail_D.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutAdjust Or _
                rowItemDetail_D.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutPR Or _
                rowItemDetail_D.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutSales Or _
                rowItemDetail_D.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutStockTakeAdjust Or _
                rowItemDetail_D.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling Then
                    'hitung rowitem detailnya, pastikan tidak menambahkan lagi ke item pending, karena ini memang
                    'perhitungan yang berasal dari item pending
                    Me.HitungOutItem(TBLInvPool, CostingMethod, TBL_Pending, rowItemDetail_D, vDefPrec, False, Nothing)
                    'jika tipe itemnya adalah out bundling, maka coba hitung in bundlingnya
                    'walaupun tidak berhasil hitung, tetap saja melakukan perhitungan in, supaya item in-nya masuk ke dalam item pending
                    'sehingga urutan FIFO/LIFO tetap terjaga
                    'berbeda dengan ketika pada fungsi hitung cost yang berlaku seperti di atas,
                    'pada fungsi ini menghitung in bundling, hanya jika hitung outnya berhasil
                    'karena pada point ini yang terjadi adalah hitung ulang sehingga urutan item FIFO-nya sudah bener
                    If rowItemDetail_D.amount <> 0 Then
                        If rowItemDetail_D.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling Then
                            Me.TBLItemDetail.DefaultView.RowFilter = "[tipe transaksi asal] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling).ToString + _
                            " AND [primary transaksi asal] = " + rowItemDetail_D.primary_transaksi_asal.ToString + _
                            " AND [tipe item detail] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem).ToString
                            If Me.TBLItemDetail.DefaultView.Count > 0 Then
                                rowItemDetail_D1 = Me.TBLItemDetail.DefaultView(0).Row
                                Me.HitungBundling(TBLInvPool, CostingMethod, TBL_Pending, rowItemDetail_D1, vDefPrec, False)
                            Else
                                MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                                Exit Sub
                            End If
                        End If
                        'Else
                        'MsgBox("")
                    End If
                Else
                    'hitung item bundling, pastikan tidak menambahkan lagi ke item pending, karena ini memang
                    'perhitungan yang berasal dari item pending
                    Me.HitungBundling1(TBLInvPool, CostingMethod, TBL_Pending, rowItemDetail_D, vDefPrec, False)
                End If

                'jika perhitungan berhasil, coba hitung item pending selanjutnya, jika ada
                If rowItemDetail_D.amount = 0 Then
                    If rowPending.tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem Then
                        'amount =0, artinya di pending kembali, klo out item, maka bawahnya yang out item juga
                        'pasti juga ga kebagian, maka langsung exit saja
                        Exit While
                    Else
                        If RefNumber = TBL_Pending.Rows.Count Then
                            counter += 1
                        End If
                    End If
                Else
                    TBL_Pending.DefaultView.RowFilter = vFilterPending
                    'hanya khusus untuk fungsi ini saja, untuk fungsi standarnya tidak ada
                    counter += 1
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function HitungBundling(ByVal TBLInvPool As DseMonthlyClosingNew.tblItemDetailPoolDataTable, _
   ByVal CostingMethod As ClsTransaksiMonthlyClosingItem.EnumCostingMethod, _
   ByVal TBL_Pending As DseMonthlyClosingNew.tblItemDetailPoolDataTable, _
   ByVal vRowRef As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailRow, ByVal vDefPrec As Integer, _
   ByVal vAddToPending As Boolean) As Boolean
        Dim DummyFilter As String = ""
        Dim DummySort As String = ""

        Dim PrimaryTransaksiBundling As Double = -1
        Dim QtyInventoryPool As Decimal = 0
        Dim QtyPending As Decimal = 0
        Dim DV As DataRowView
        Dim Lanjut As Boolean = True
        Dim RowOutBundling As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailRow
        Dim row_Pending As DseMonthlyClosingNew.tblItemDetailPoolRow
        Dim QtyOut As Decimal = 0
        Dim QtyAllocated As Decimal = 0
        Dim Amount As Decimal = 0
        Dim AmountAllocated As Decimal = 0
        Dim BundlingAmount As Decimal = 0
        Dim RowInBundling As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailRow
        Dim counter As Integer = 0

        Try
            'pastikan semua item beforenya sudah dihitung
            DummyFilter = Me.TBLItemDetail.DefaultView.RowFilter
            'revisi: 2018-02-02: sort 1x saja
            'DummySort = Me.'TBLItemDetail.DefaultView.Sort
            PrimaryTransaksiBundling = vRowRef.primary_transaksi_asal

            Lanjut = True
            'revisi: 2018-02-02: sort 1x saja
            'Me.''TBLItemDetail.DefaultView.Sort = "[tgl transaksi asal], [tipe transaksi asal], [primary transaksi asal], [primary code no]"

            Me.TBLItemDetail.DefaultView.RowFilter = "[primary transaksi asal] = " + PrimaryTransaksiBundling.ToString + _
            " AND [tipe transaksi asal] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling).ToString
            '" AND [tipe item detail] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem).ToString + _

            For Each DV In Me.TBLItemDetail.DefaultView
                RowOutBundling = DV.Row

                If RowOutBundling.amount = 0 Then
                    Lanjut = False
                    Exit For
                End If
            Next

            If Lanjut = False Then
                'tambahkan item after ke dalam table pending, jika sebelumnya tidak ada
                If vAddToPending = True Then
                    TBL_Pending.DefaultView.RowFilter = "[primary transaksi asal] = " + PrimaryTransaksiBundling.ToString + _
                    " AND [tipe item detail] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem).ToString + _
                    " AND [tipe transaksi asal] = " + CInt(vRowRef.tipe_transaksi_asal).ToString

                    If TBL_Pending.DefaultView.Count = 0 Then
                        'hanya tambahkan ke item pending, bila sebelumnya di item pending tidak ada
                        Me.TBLItemDetail.DefaultView.RowFilter = "[primary transaksi asal] = " + PrimaryTransaksiBundling.ToString + _
                        " AND [tipe item detail] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem).ToString + _
                        " AND [tipe transaksi asal] = " + CInt(vRowRef.tipe_transaksi_asal).ToString

                        For Each DV In Me.TBLItemDetail.DefaultView
                            RowInBundling = DV.Row
                            With RowInBundling
                                row_Pending = TBL_Pending.NewtblItemDetailPoolRow

                                row_Pending.amount = .amount
                                row_Pending.Cost_Perct_Bundling = .Cost_Perct_Bundling
                                row_Pending.Overhead = .Overhead
                                row_Pending.Primary = Me.GetPrimaryMaxInventoryPool(TBL_Pending) + 1
                                row_Pending.primary_adjustment_type = .primary_adjustment_type
                                row_Pending.primary_code_no = .primary_code_no
                                row_Pending.primary_gain_account = .primary_gain_account
                                row_Pending.primary_loss_account = .primary_loss_account
                                row_Pending.Primary_transaksi = Me.TxtPrimaryClosing.Text
                                row_Pending.Primary_Transaksi_Asal = .primary_transaksi_asal
                                row_Pending.qty = .qty
                                row_Pending.remarks = .remarks
                                row_Pending.return_price = .return_price
                                row_Pending.tgl_transaksi_asal = .tgl_transaksi_asal
                                row_Pending.tipe_item_detail = .tipe_item_detail
                                row_Pending.tipe_transaksi_asal = .tipe_transaksi_asal
                                row_Pending.unit_price = .unit_price
                                If row_Pending.remarks.Length > 250 Then
                                    row_Pending.remarks = row_Pending.remarks.Substring(0, 250)
                                End If
                                TBL_Pending.Rows.Add(row_Pending)
                            End With
                        Next
                    End If
                End If
            Else
                'hitung amount bundling berdasarkan item before
                BundlingAmount = 0
                'revisi: 2018-02-02: sort 1x saja
                'Me.''TBLItemDetail.DefaultView.Sort = "[tgl transaksi asal], [tipe transaksi asal], [primary transaksi asal], [primary code no]"
                Me.TBLItemDetail.DefaultView.RowFilter = "[primary transaksi asal] = " + PrimaryTransaksiBundling.ToString + _
                " AND [tipe item detail] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem).ToString + _
                " AND [tipe transaksi asal] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling).ToString
                For Each DV In Me.TBLItemDetail.DefaultView
                    RowOutBundling = DV.Row
                    BundlingAmount += RowOutBundling.amount
                Next

                'start perhitungan item after
                Me.TBLItemDetail.DefaultView.RowFilter = "[primary transaksi asal] = " + PrimaryTransaksiBundling.ToString + _
                " AND [tipe item detail] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem).ToString + _
                " AND [tipe transaksi asal] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling).ToString
                'revisi: 2018-02-02: sort 1x saja
                'Me.''TBLItemDetail.DefaultView.Sort = "[tgl transaksi asal], [tipe transaksi asal], [primary transaksi asal], [primary code no]"

                AmountAllocated = 0
                For counter = 0 To Me.TBLItemDetail.DefaultView.Count - 1
                    RowInBundling = Me.TBLItemDetail.DefaultView(counter).Row
                    If counter = Me.TBLItemDetail.DefaultView.Count - 1 Then
                        'sisanya
                        RowInBundling.amount = RowInBundling.Overhead + (BundlingAmount - AmountAllocated)
                    Else
                        'hitung berdasarkan cost factor
                        Amount = Math.Round((BundlingAmount * RowInBundling.Cost_Perct_Bundling) / 100, _
                        vDefPrec, MidpointRounding.AwayFromZero)
                        RowInBundling.amount = RowInBundling.Overhead + Amount
                        AmountAllocated += Amount
                    End If
                    RowInBundling.unit_price = Math.Round(RowInBundling.amount / RowInBundling.qty, vDefPrec, MidpointRounding.AwayFromZero)

                    'tambahkan inventory pool sesuai item after
                    If CostingMethod = ClsTransaksiMonthlyClosingItem.EnumCostingMethod.Average Then
                        TambahItem_InventoryPool_Avg(RowInBundling, TBLInvPool, vDefPrec)
                    Else
                        Me.TambahItem_InventoryPool_FIFOLIFO1(RowInBundling, TBLInvPool, vDefPrec, CostingMethod)
                    End If
                Next

                'buang item pending jika ada
                TBL_Pending.DefaultView.RowFilter = "[tipe transaksi asal] = " + vRowRef.tipe_transaksi_asal.ToString + _
                " AND [primary transaksi asal] = " + vRowRef.primary_transaksi_asal.ToString + _
                " AND [tipe item detail] = " + vRowRef.tipe_item_detail.ToString

                While TBL_Pending.DefaultView.Count > 0
                    TBL_Pending.Rows.Remove(TBL_Pending.DefaultView(0).Row)
                End While
            End If

            Me.TBLItemDetail.DefaultView.RowFilter = DummyFilter
            'revisi: 2018-02-02: sort 1x saja
            'Me.''TBLItemDetail.DefaultView.Sort = DummySort

            'mengembalikan nilai true jika berhasil hitung transaksi bundling
            Return Lanjut
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function HitungBundling1(ByVal TBLInvPool As DseMonthlyClosingNew.tblItemDetailPoolDataTable, _
  ByVal CostingMethod As ClsTransaksiMonthlyClosingItem.EnumCostingMethod, _
  ByVal TBL_Pending As DseMonthlyClosingNew.tblItemDetailPoolDataTable, _
  ByVal vRowRef As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailRow, ByVal vDefPrec As Integer, _
  ByVal vAddToPending As Boolean) As Boolean
        Dim DummyFilter As String = ""
        Dim DummySort As String = ""

        Dim PrimaryTransaksiBundling As Double = -1
        Dim QtyInventoryPool As Decimal = 0
        Dim QtyPending As Decimal = 0
        Dim DV As DataRowView
        Dim Lanjut As Boolean = True
        Dim RowOutBundling As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailRow
        Dim row_Pending As DseMonthlyClosingNew.tblItemDetailPoolRow
        Dim QtyOut As Decimal = 0
        Dim QtyAllocated As Decimal = 0
        Dim Amount As Decimal = 0
        Dim AmountAllocated As Decimal = 0
        Dim BundlingAmount As Decimal = 0
        Dim RowInBundling As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailRow
        Dim counter As Integer = 0

        Try
            'pastikan semua item beforenya sudah dihitung
            DummyFilter = Me.TBLItemDetail.DefaultView.RowFilter
            'revisi: 2018-02-02: sort 1x saja
            'DummySort = Me.''TBLItemDetail.DefaultView.Sort
            PrimaryTransaksiBundling = vRowRef.primary_transaksi_asal

            Lanjut = True
            'revisi: 2018-02-02: sort 1x saja
            'Me.'TBLItemDetail.DefaultView.Sort = "[tgl transaksi asal], [tipe transaksi asal], [primary transaksi asal], [primary code no]"

            Me.TBLItemDetail.DefaultView.RowFilter = "[primary transaksi asal] = " + PrimaryTransaksiBundling.ToString + _
            " AND [tipe transaksi asal] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling).ToString
            '" AND [tipe item detail] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem).ToString + _

            For Each DV In Me.TBLItemDetail.DefaultView
                RowOutBundling = DV.Row

                If RowOutBundling.amount = 0 Then
                    RowOutBundling.amount = (RowOutBundling.qty * 200) + RowOutBundling.Overhead
                    RowOutBundling.unit_price = Math.Round(RowOutBundling.amount / RowOutBundling.qty, 0, MidpointRounding.AwayFromZero)

                    'Lanjut = False
                    'Exit For
                End If
            Next

            If Lanjut = False Then
                'tambahkan item after ke dalam table pending, jika sebelumnya tidak ada
                If vAddToPending = True Then
                    TBL_Pending.DefaultView.RowFilter = "[primary transaksi asal] = " + PrimaryTransaksiBundling.ToString + _
                    " AND [tipe item detail] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem).ToString + _
                    " AND [tipe transaksi asal] = " + CInt(vRowRef.tipe_transaksi_asal).ToString

                    If TBL_Pending.DefaultView.Count = 0 Then
                        'hanya tambahkan ke item pending, bila sebelumnya di item pending tidak ada
                        Me.TBLItemDetail.DefaultView.RowFilter = "[primary transaksi asal] = " + PrimaryTransaksiBundling.ToString + _
                        " AND [tipe item detail] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem).ToString + _
                        " AND [tipe transaksi asal] = " + CInt(vRowRef.tipe_transaksi_asal).ToString

                        For Each DV In Me.TBLItemDetail.DefaultView
                            RowInBundling = DV.Row
                            With RowInBundling
                                row_Pending = TBL_Pending.NewtblItemDetailPoolRow

                                row_Pending.amount = .amount
                                row_Pending.Cost_Perct_Bundling = .Cost_Perct_Bundling
                                row_Pending.Overhead = .Overhead
                                row_Pending.Primary = Me.GetPrimaryMaxInventoryPool(TBL_Pending) + 1
                                row_Pending.primary_adjustment_type = .primary_adjustment_type
                                row_Pending.primary_code_no = .primary_code_no
                                row_Pending.primary_gain_account = .primary_gain_account
                                row_Pending.primary_loss_account = .primary_loss_account
                                row_Pending.Primary_transaksi = Me.TxtPrimaryClosing.Text
                                row_Pending.Primary_Transaksi_Asal = .primary_transaksi_asal
                                row_Pending.qty = .qty
                                row_Pending.remarks = .remarks
                                row_Pending.return_price = .return_price
                                row_Pending.tgl_transaksi_asal = .tgl_transaksi_asal
                                row_Pending.tipe_item_detail = .tipe_item_detail
                                row_Pending.tipe_transaksi_asal = .tipe_transaksi_asal
                                row_Pending.unit_price = .unit_price
                                If row_Pending.remarks.Length > 250 Then
                                    row_Pending.remarks = row_Pending.remarks.Substring(0, 250)
                                End If
                                TBL_Pending.Rows.Add(row_Pending)
                            End With
                        Next
                    End If
                End If
            Else
                'hitung amount bundling berdasarkan item before
                BundlingAmount = 0
                'revisi: 2018-02-02: sort 1x saja
                'Me.'TBLItemDetail.DefaultView.Sort = "[tgl transaksi asal], [tipe transaksi asal], [primary transaksi asal], [primary code no]"
                Me.TBLItemDetail.DefaultView.RowFilter = "[primary transaksi asal] = " + PrimaryTransaksiBundling.ToString + _
                " AND [tipe item detail] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem).ToString + _
                " AND [tipe transaksi asal] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling).ToString
                For Each DV In Me.TBLItemDetail.DefaultView
                    RowOutBundling = DV.Row
                    BundlingAmount += RowOutBundling.amount
                Next

                'start perhitungan item after
                Me.TBLItemDetail.DefaultView.RowFilter = "[primary transaksi asal] = " + PrimaryTransaksiBundling.ToString + _
                " AND [tipe item detail] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem).ToString + _
                " AND [tipe transaksi asal] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling).ToString
                'revisi: 2018-02-02: sort 1x saja
                'Me.'TBLItemDetail.DefaultView.Sort = "[tgl transaksi asal], [tipe transaksi asal], [primary transaksi asal], [primary code no]"

                AmountAllocated = 0
                For counter = 0 To Me.TBLItemDetail.DefaultView.Count - 1
                    RowInBundling = Me.TBLItemDetail.DefaultView(counter).Row
                    If counter = Me.TBLItemDetail.DefaultView.Count - 1 Then
                        'sisanya
                        RowInBundling.amount = RowInBundling.Overhead + (BundlingAmount - AmountAllocated)
                    Else
                        'hitung berdasarkan cost factor
                        Amount = Math.Round((BundlingAmount * RowInBundling.Cost_Perct_Bundling) / 100, _
                        vDefPrec, MidpointRounding.AwayFromZero)
                        RowInBundling.amount = RowInBundling.Overhead + Amount
                        AmountAllocated += Amount
                    End If
                    RowInBundling.unit_price = Math.Round(RowInBundling.amount / RowInBundling.qty, vDefPrec, MidpointRounding.AwayFromZero)

                    'tambahkan inventory pool sesuai item after
                    If CostingMethod = ClsTransaksiMonthlyClosingItem.EnumCostingMethod.Average Then
                        TambahItem_InventoryPool_Avg(RowInBundling, TBLInvPool, vDefPrec)
                    Else
                        Me.TambahItem_InventoryPool_FIFOLIFO1(RowInBundling, TBLInvPool, vDefPrec, CostingMethod)
                    End If
                Next

                'buang item pending jika ada
                TBL_Pending.DefaultView.RowFilter = "[tipe transaksi asal] = " + vRowRef.tipe_transaksi_asal.ToString + _
                " AND [primary transaksi asal] = " + vRowRef.primary_transaksi_asal.ToString + _
                " AND [tipe item detail] = " + vRowRef.tipe_item_detail.ToString

                While TBL_Pending.DefaultView.Count > 0
                    TBL_Pending.Rows.Remove(TBL_Pending.DefaultView(0).Row)
                End While
            End If

            Me.TBLItemDetail.DefaultView.RowFilter = DummyFilter
            'revisi: 2018-02-02: sort 1x saja
            'Me.'TBLItemDetail.DefaultView.Sort = DummySort

            'mengembalikan nilai true jika berhasil hitung transaksi bundling
            Return Lanjut
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub TambahItem_InventoryPool_FIFOLIFO1(ByVal vrowItemDetail As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailRow, _
       ByVal TBLInvPool As DseMonthlyClosingNew.tblItemDetailPoolDataTable, ByVal vDefPrec As Integer, _
       ByVal vCostingMethod As ClsTransaksiMonthlyClosingItem.EnumCostingMethod)
        Dim rowInvPool As DseMonthlyClosingNew.tblItemDetailPoolRow
        Dim AdaQtyNegative As Boolean = False
        Dim QtyTutupMinus As Decimal = 0
        Dim AmountTutupMinus As Decimal = 0

        Try
            TBLInvPool.DefaultView.RowFilter = "[primary code no] = " + vrowItemDetail.primary_code_no.ToString + " AND qty <= 0"
            If TBLInvPool.DefaultView.Count > 0 Then
                'kurangkan minusnya dahulu
                rowInvPool = TBLInvPool.DefaultView(0).Row
                If rowInvPool.qty <> 0 Then
                    If rowInvPool.qty + vrowItemDetail.qty >= 0 Then
                        QtyTutupMinus = rowInvPool.qty * -1
                        AmountTutupMinus = rowInvPool.amount * -1

                        'rowInvPool.amount += vrowItemDetail.amount
                        'rowInvPool.qty += vrowItemDetail.qty

                        If rowInvPool.amount + vrowItemDetail.amount = 0 And (rowInvPool.qty + vrowItemDetail.qty) = 0 Then
                            'jika amount sisa =0 dan qty sisa = 0 -> hapus baris inventory pool, ga usah tambah baris baru
                            TBLInvPool.Rows.Remove(rowInvPool)
                            Exit Sub
                        ElseIf (rowInvPool.qty + vrowItemDetail.qty) = 0 Or rowInvPool.amount + vrowItemDetail.amount < 0 Then
                            'jika qty sisa = 0 dan amount sisa <> 0 -> tambahan nilai qty dan amount pd inventory pool, tidak usah tambah baris baru
                            rowInvPool.amount += vrowItemDetail.amount
                            rowInvPool.qty += vrowItemDetail.qty

                            If rowInvPool.qty = 0 Then
                                rowInvPool.unit_price = 0
                            Else
                                rowInvPool.unit_price = Math.Round((rowInvPool.amount / rowInvPool.qty), vDefPrec, MidpointRounding.AwayFromZero)
                            End If

                            Exit Sub
                        ElseIf rowInvPool.amount + vrowItemDetail.amount = 0 Then
                            'jika qty sisa <> 0 and amount sisa = 0 -> remove row inv pool, kemudian ditambahkan lagi row inventory pool pada bagian bawahnya
                            TBLInvPool.Rows.Remove(rowInvPool)
                        ElseIf rowInvPool.amount + vrowItemDetail.amount > 0 And (rowInvPool.qty + vrowItemDetail.qty) > 0 Then
                            TBLInvPool.Rows.Remove(rowInvPool)
                        End If
                    Else
                        'hanya kurangkan sebagian
                        rowInvPool.amount += vrowItemDetail.amount
                        rowInvPool.qty += vrowItemDetail.qty
                        If rowInvPool.qty = 0 Then
                            rowInvPool.unit_price = 0
                        Else
                            rowInvPool.unit_price = Math.Abs(Math.Round((rowInvPool.amount / rowInvPool.qty), vDefPrec, MidpointRounding.AwayFromZero))
                        End If
                        QtyTutupMinus = vrowItemDetail.qty
                    End If
                Else
                    AmountTutupMinus = rowInvPool.amount * -1
                    TBLInvPool.Rows.Remove(rowInvPool)
                End If
            Else
                'cek kasus khusus dimana qty > 0 tapi amount <= 0
                TBLInvPool.DefaultView.RowFilter = "[primary code no] = " + vrowItemDetail.primary_code_no.ToString + " AND amount <= 0"
                If TBLInvPool.DefaultView.Count > 0 Then
                    rowInvPool = TBLInvPool.DefaultView(0).Row
                    AmountTutupMinus = rowInvPool.amount * -1
                    If rowInvPool.qty > 0 Then
                        QtyTutupMinus = rowInvPool.qty * -1
                        TBLInvPool.Rows.Remove(rowInvPool)
                    ElseIf rowInvPool.qty = 0 Then
                        AmountTutupMinus = rowInvPool.amount * -1
                        TBLInvPool.Rows.Remove(rowInvPool)
                    Else
                        MsgBox("Logic error, there should be no qty < 0, on this spot.", MsgBoxStyle.Information, "Information")
                    End If
                End If
            End If


            If vrowItemDetail.qty - QtyTutupMinus > 0 Then
                With vrowItemDetail
                    rowInvPool = TBLInvPool.NewtblItemDetailPoolRow

                    rowInvPool.amount = .amount - AmountTutupMinus
                    rowInvPool.Cost_Perct_Bundling = .Cost_Perct_Bundling
                    rowInvPool.Overhead = .Overhead
                    rowInvPool.Primary = Me.GetPrimaryMaxInventoryPool(TBLInvPool) + 1
                    rowInvPool.primary_adjustment_type = .primary_adjustment_type
                    rowInvPool.primary_code_no = .primary_code_no
                    rowInvPool.primary_gain_account = .primary_gain_account
                    rowInvPool.primary_loss_account = .primary_loss_account
                    rowInvPool.Primary_transaksi = Me.TxtPrimaryClosing.Text
                    rowInvPool.Primary_Transaksi_Asal = .primary_transaksi_asal
                    rowInvPool.qty = .qty - QtyTutupMinus
                    rowInvPool.remarks = .remarks
                    rowInvPool.return_price = .return_price
                    rowInvPool.tgl_transaksi_asal = .tgl_transaksi_asal
                    rowInvPool.tipe_item_detail = .tipe_item_detail
                    rowInvPool.tipe_transaksi_asal = .tipe_transaksi_asal
                    rowInvPool.unit_price = .unit_price

                    If rowInvPool.remarks.Length > 250 Then
                        rowInvPool.remarks = rowInvPool.remarks.Substring(0, 250)
                    End If
                    TBLInvPool.Rows.Add(rowInvPool)
                End With
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TambahItem_InventoryPool_Avg(ByVal vrowItemDetail As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailRow, _
    ByVal TBLInvPool As DseMonthlyClosingNew.tblItemDetailPoolDataTable, ByVal vDefPrec As Integer)
        Dim rowInvPool As DseMonthlyClosingNew.tblItemDetailPoolRow

        Try
            'cari data diinventory pool, yang sesuai primary code no-nya
            TBLInvPool.DefaultView.RowFilter = "[primary code no] = " + vrowItemDetail.primary_code_no.ToString
            If TBLInvPool.DefaultView.Count = 0 Then
                With vrowItemDetail
                    rowInvPool = TBLInvPool.NewtblItemDetailPoolRow

                    rowInvPool.amount = .amount
                    rowInvPool.Cost_Perct_Bundling = 0
                    rowInvPool.Overhead = 0
                    rowInvPool.Primary = Me.GetPrimaryMaxInventoryPool(TBLInvPool) + 1
                    rowInvPool.primary_adjustment_type = -1
                    rowInvPool.primary_code_no = .primary_code_no
                    rowInvPool.primary_gain_account = -1
                    rowInvPool.primary_loss_account = -1
                    rowInvPool.Primary_transaksi = Me.TxtPrimaryClosing.Text
                    rowInvPool.Primary_Transaksi_Asal = -1
                    rowInvPool.qty = .qty
                    rowInvPool.remarks = ""
                    rowInvPool.return_price = .return_price
                    rowInvPool.tgl_transaksi_asal = #1/1/1950#
                    rowInvPool.tipe_item_detail = -1
                    rowInvPool.tipe_transaksi_asal = -1
                    rowInvPool.unit_price = .unit_price

                    TBLInvPool.Rows.Add(rowInvPool)
                End With
            Else
                'buang barang minus terlebih dahulu
                rowInvPool = TBLInvPool.DefaultView(0).Row

                With vrowItemDetail
                    rowInvPool.qty += .qty
                    rowInvPool.amount += .amount

                    If rowInvPool.qty = 0 And rowInvPool.amount = 0 Then
                        TBLInvPool.Rows.Remove(rowInvPool)
                    Else
                        If rowInvPool.qty = 0 Then
                            rowInvPool.unit_price = 0
                        Else
                            rowInvPool.unit_price = Math.Round(rowInvPool.amount / rowInvPool.qty, vDefPrec, MidpointRounding.AwayFromZero)
                        End If
                    End If
                End With
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HitungOutItem(ByVal TBLInvPool As DseMonthlyClosingNew.tblItemDetailPoolDataTable, _
  ByVal CostingMethod As ClsTransaksiMonthlyClosingItem.EnumCostingMethod, _
  ByVal TBL_Pending As DseMonthlyClosingNew.tblItemDetailPoolDataTable, _
  ByVal vRowRef As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailRow, ByVal vDefPrec As Integer, _
  ByVal vAddToPending As Boolean, _
  ByVal vRow_Pending2 As DseMonthlyClosingNew.tblItemDetailPoolRow)
        Dim QtyInventoryPool As Decimal = 0
        Dim QtyPending As Decimal = 0
        Dim DV As DataRowView
        Dim row_Pending As DseMonthlyClosingNew.tblItemDetailPoolRow
        Dim rowInvPool As DseMonthlyClosingNew.tblItemDetailPoolRow
        Dim QtyOut As Decimal = 0
        Dim QtyAllocated As Decimal = 0
        Dim Amount As Decimal = 0
        Dim AmountAllocated As Decimal = 0
        Dim AdaInPending As Boolean = False
        Dim DummyString As String = ""
        Dim Syarat1, Syarat2, Syarat3 As Boolean
        Dim DataBetul As Boolean = True
        Dim CaraHitung As Integer = 0
        Dim JmlPendingSblm As Decimal = 0

        Try
            'cek apakah ada item bundling yang dipending, klo ada, maka jangan itung dulu karena kemungkinan akan mengganggu urutan kerja
            TBL_Pending.DefaultView.RowFilter = "[primary code no] = " + vRowRef.primary_code_no.ToString + _
            " AND [tipe item detail] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem).ToString
            If TBL_Pending.DefaultView.Count > 0 Then
                'yang bisa menunda pending adalah jika item in pending, bukan berasal dari transaksi yang sama
                'dengan baris yang akan dihitung
                If vRowRef.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling Then
                    'jika tipe out-nya bundling, maka pastikan yang ngeblok bukan dari transaksi bundling yang sama
                    'pastikan juga tgl transaksi yang ngeblok lebih kecil atau sama dengan tgl transaksi yg lg mau dihitung
                    'pastikan juga primary transaksi lebih kecil dari item yang mau dihitung
                    If CostingMethod = ClsTransaksiMonthlyClosingItem.EnumCostingMethod.Average Or _
                    CostingMethod = ClsTransaksiMonthlyClosingItem.EnumCostingMethod.LIFO Then
                        For Each DV In TBL_Pending.DefaultView
                            row_Pending = DV.Row
                            If row_Pending.Primary_Transaksi_Asal <> vRowRef.primary_transaksi_asal Then
                                If row_Pending.tgl_transaksi_asal.Date <= vRowRef.tgl_transaksi_asal Then
                                    If row_Pending.Primary_Transaksi_Asal < vRowRef.primary_transaksi_asal Then
                                        'jika 
                                        AdaInPending = True
                                        Exit For
                                    End If
                                End If
                            End If
                        Next
                    Else
                        'khusus untuk FIFO ada kasus khusus dimana:
                        'item out tetap bisa dihitung, walaupun ada item pending, jika memenuhi 3 syarat:
                        '1. semua transaksi pending yang tanggalnya lebih kecil dari transaksi acuan adalah in item -> syarat1
                        '2. semua transaksi pending yang tanggalnya sama dengan transaksi acuan dan primary transaksinya lebih kecil adalah in item -> syarat2
                        '3. jml inventory pool - jml pending in sebelumnya yang belum ada harganya >= jmlref -> syarat3 : 
                        'revisi 2018-01-21: tadinya: jml inventory pool >= jmlref, menjadi
                        'jml inventory pool - jml pending in sebelumnya yang belum ada harganya >= jmlref 
                        'jika tidak direvisi, maka urutan pengambilan amount, akan loncat, karena ada kemungkinan inventory pol stocknya cukup
                        'dari in item tanggal selanjutnya, yang in adjust misalnya, padahal seharusnya ambil dari in bundling yang masih pending
                        'akibat revisi ini, syarat 1,2,3 harus tetap di cek, ga boleh lompat2 syarat
                        'syarat1
                        Syarat1 = True
                        Syarat2 = True
                        Syarat3 = True
                        JmlPendingSblm = 0
                        'syarat 1
                        For Each DV In TBL_Pending.DefaultView
                            row_Pending = DV.Row
                            If row_Pending.tgl_transaksi_asal.Date < vRowRef.tgl_transaksi_asal.Date Then
                                If row_Pending.tipe_item_detail <> CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem) Then
                                    Syarat1 = False
                                ElseIf row_Pending.tipe_item_detail = CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem) Then
                                    JmlPendingSblm += row_Pending.qty
                                End If
                            End If
                        Next
                        'syarat 2
                        For Each DV In TBL_Pending.DefaultView
                            row_Pending = DV.Row
                            If row_Pending.tgl_transaksi_asal.Date = vRowRef.tgl_transaksi_asal.Date Then
                                If row_Pending.Primary_Transaksi_Asal < vRowRef.primary_transaksi_asal Then
                                    If row_Pending.tipe_item_detail <> CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem) Then
                                        Syarat2 = False
                                    ElseIf row_Pending.tipe_item_detail = CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem) Then
                                        JmlPendingSblm += row_Pending.qty
                                    End If
                                End If
                            End If
                        Next
                        'syarat 3
                        TBLInvPool.DefaultView.RowFilter = "[primary code no] = " + vRowRef.primary_code_no.ToString
                        QtyInventoryPool = 0
                        For Each DV In TBLInvPool.DefaultView
                            rowInvPool = DV.Row
                            QtyInventoryPool += rowInvPool.qty
                        Next
                        If Me.custom_module = "smess_lmb" Then
                            JmlPendingSblm = 0
                        End If
                        If QtyInventoryPool - JmlPendingSblm < vRowRef.qty Then
                            Syarat3 = False
                        End If

                        If Syarat3 = False And Me.custom_module <> "smess_lmb" Then
                            AdaInPending = True
                        Else
                            If Syarat1 = False Or Syarat2 = False Or Syarat3 = False Then
                                'cek normal seperti avg dan lifo
                                For Each DV In TBL_Pending.DefaultView
                                    row_Pending = DV.Row
                                    If row_Pending.Primary_Transaksi_Asal <> vRowRef.primary_transaksi_asal Then
                                        If row_Pending.tgl_transaksi_asal.Date <= vRowRef.tgl_transaksi_asal Then
                                            If row_Pending.Primary_Transaksi_Asal < vRowRef.primary_transaksi_asal Then
                                                'jika 
                                                AdaInPending = True
                                                Exit For
                                            End If
                                        End If
                                    End If
                                Next
                            End If
                        End If
                    End If
                Else
                    If CostingMethod = ClsTransaksiMonthlyClosingItem.EnumCostingMethod.Average Or _
                    CostingMethod = ClsTransaksiMonthlyClosingItem.EnumCostingMethod.LIFO Then
                        AdaInPending = True
                    Else
                        'item out tetap bisa dihitung, walaupun ada item pending, jika memenuhi 3 syarat:
                        '1. semua transaksi pending yang tanggalnya lebih kecil dari transaksi acuan adalah in item -> syarat1
                        '2. semua transaksi pending yang tanggalnya sama dengan transaksi acuan dan primary transaksinya lebih kecil adalah in item -> syarat2
                        '3. jml inventory pool >= jmlref -> syarat3 -> revisi 2018-01-21: lihat point di atas, untuk perhitungan out bundling
                        Syarat1 = True
                        Syarat2 = True
                        Syarat3 = True
                        JmlPendingSblm = 0
                        For Each DV In TBL_Pending.DefaultView
                            row_Pending = DV.Row
                            If row_Pending.tgl_transaksi_asal.Date < vRowRef.tgl_transaksi_asal.Date Then
                                If row_Pending.tipe_item_detail <> CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem) Then
                                    Syarat1 = False
                                ElseIf row_Pending.tipe_item_detail = CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem) Then
                                    JmlPendingSblm += row_Pending.qty
                                End If
                            End If
                        Next
                        For Each DV In TBL_Pending.DefaultView
                            row_Pending = DV.Row
                            If row_Pending.tgl_transaksi_asal.Date = vRowRef.tgl_transaksi_asal.Date Then
                                If row_Pending.Primary_Transaksi_Asal < vRowRef.primary_transaksi_asal Then
                                    If row_Pending.tipe_item_detail <> CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem) Then
                                        Syarat2 = False
                                    ElseIf row_Pending.tipe_item_detail = CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem) Then
                                        JmlPendingSblm += row_Pending.qty
                                    End If
                                End If
                            End If
                        Next
                        TBLInvPool.DefaultView.RowFilter = "[primary code no] = " + vRowRef.primary_code_no.ToString
                        QtyInventoryPool = 0
                        For Each DV In TBLInvPool.DefaultView
                            rowInvPool = DV.Row
                            QtyInventoryPool += rowInvPool.qty
                        Next
                        If Me.custom_module = "smess_lmb" Then
                            JmlPendingSblm = 0
                        End If
                        If QtyInventoryPool - JmlPendingSblm < vRowRef.qty Then
                            Syarat3 = False
                        End If
                        If Syarat1 = False Or Syarat2 = False Or Syarat3 = False Then
                            'jika tipe out-nya bukan bundling, maka langsung hitung aj
                            AdaInPending = True
                        End If
                    End If
                End If
            Else
                AdaInPending = False
            End If
            'pastikan qty-nya cukup, setelah memperhitungkan inventory pool dan item pending out
            'item pending in tidak diperhitungkan
            If AdaInPending = False Then
                TBL_Pending.DefaultView.RowFilter = "[primary code no] = " + vRowRef.primary_code_no.ToString + _
                            " AND [tipe item detail] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem).ToString
                QtyPending = 0
                For Each DV In TBL_Pending.DefaultView
                    row_Pending = DV.Row
                    If row_Pending.tgl_transaksi_asal.Date.Date < vRowRef.tgl_transaksi_asal.Date Then
                        'jika ada item pending yang tanggalnya lebih kecil, pasti termasuk pending
                        QtyPending += row_Pending.qty
                    ElseIf row_Pending.tgl_transaksi_asal.Date = vRowRef.tgl_transaksi_asal.Date Then
                        If row_Pending.tgl_transaksi_asal < vRowRef.tgl_transaksi_asal Then
                            'jika tanggal yang sama, maka tentukan berdasarkan urutan hitung
                            'urutan hitung berdasarkan tipe transaksi asal item detail, yang lebih kecil dihitung duluan
                            'maka yang lebih kecil dihitung sebagai qty pending, yang lebih besar tidak
                            QtyPending += row_Pending.qty
                        ElseIf row_Pending.tgl_transaksi_asal = vRowRef.tgl_transaksi_asal Then
                            'jika tanggal sama, tipe transaksi sama
                            'maka dihitung pending jika primarynya lebih kecil
                            If row_Pending.Primary_Transaksi_Asal < vRowRef.primary_transaksi_asal Then
                                QtyPending += row_Pending.qty
                                'Else
                                'MsgBox("")
                            End If
                            'Else
                            'MsgBox("")
                        End If
                        'Else
                        'MsgBox("")
                    End If
                Next
                TBLInvPool.DefaultView.RowFilter = "[primary code no] = " + vRowRef.primary_code_no.ToString
                QtyInventoryPool = 0
                For Each DV In TBLInvPool.DefaultView
                    rowInvPool = DV.Row
                    QtyInventoryPool += rowInvPool.qty
                Next
            End If

            If IsNothing(vRow_Pending2) = False Then
                If vRow_Pending2.Primary_Transaksi_Asal = vRowRef.primary_transaksi_asal And _
                vRow_Pending2.primary_code_no = vRowRef.primary_code_no And _
                vRow_Pending2.tipe_transaksi_asal = vRowRef.tipe_transaksi_asal Then 'And QtyInventoryPool >= 0 Then
                    DataBetul = True
                ElseIf vRowRef.qty > QtyInventoryPool - QtyPending Or AdaInPending = True Then
                    DataBetul = False
                Else
                    DataBetul = True
                End If
            ElseIf vRowRef.qty > QtyInventoryPool - QtyPending Or AdaInPending = True Then
                DataBetul = False
            Else
                DataBetul = True
            End If

            If DataBetul = False Then
                'tambahkan ke dalam tabel pending, hanya jika vRowRef bukan dihitung berdasarkan item pending
                If vAddToPending = True Then
                    With vRowRef
                        row_Pending = TBL_Pending.NewtblItemDetailPoolRow

                        row_Pending.amount = .amount
                        row_Pending.Cost_Perct_Bundling = .Cost_Perct_Bundling
                        row_Pending.Overhead = .Overhead
                        row_Pending.Primary = Me.GetPrimaryMaxInventoryPool(TBL_Pending) + 1
                        row_Pending.primary_adjustment_type = .primary_adjustment_type
                        row_Pending.primary_code_no = .primary_code_no
                        row_Pending.primary_gain_account = .primary_gain_account
                        row_Pending.primary_loss_account = .primary_loss_account
                        row_Pending.Primary_transaksi = Me.TxtPrimaryClosing.Text
                        row_Pending.Primary_Transaksi_Asal = .primary_transaksi_asal
                        row_Pending.qty = .qty
                        row_Pending.remarks = .remarks
                        row_Pending.return_price = .return_price
                        row_Pending.tgl_transaksi_asal = .tgl_transaksi_asal
                        row_Pending.tipe_item_detail = .tipe_item_detail
                        row_Pending.tipe_transaksi_asal = .tipe_transaksi_asal
                        row_Pending.unit_price = .unit_price
                        If row_Pending.remarks.Length > 250 Then
                            row_Pending.remarks = row_Pending.remarks.Substring(0, 250)
                        End If
                        TBL_Pending.Rows.Add(row_Pending)
                    End With
                End If
            Else
                'mulai perhitungan nilai out
                If IsNothing(vRow_Pending2) = True Then
                    CaraHitung = 0
                Else
                    If vRow_Pending2.Primary_Transaksi_Asal = vRowRef.primary_transaksi_asal Then
                        CaraHitung = 1
                    Else
                        CaraHitung = 0
                    End If
                End If

                If CaraHitung = 0 Then
                    TBLInvPool.DefaultView.RowFilter = "[primary code no] = " + vRowRef.primary_code_no.ToString
                    'If CostingMethod = ClsTransaksiMonthlyClosingItem.EnumCostingMethod.FIFO Then
                    '    TBLInvPool.DefaultView.Sort = "[tgl transaksi asal], [tipe transaksi asal], [primary transaksi asal], [primary code no]"
                    'Else
                    '    TBLInvPool.DefaultView.Sort = "[tgl transaksi asal] DESC, [tipe transaksi asal], [primary transaksi asal] DESC, [primary code no]"
                    'End If

                    QtyOut = vRowRef.qty
                    QtyAllocated = 0
                    Amount = 0
                    DummyString = "" '--> format untuk dummy string: tipe transaksi asal, primary transaksi asal, qty, amount, dipisahkan oleh ';'
                    While QtyOut > QtyAllocated
                        rowInvPool = TBLInvPool.DefaultView(0).Row
                        If rowInvPool.qty < 0 Then
                            MsgBox("Logic error: no precedure for minus qty in tbl inventory pool", _
                            MsgBoxStyle.Information, "Information")
                            Exit Sub
                        End If

                        If rowInvPool.qty > QtyOut - QtyAllocated Then
                            'ambil sebagian saja: qty out - qty allocated
                            'update amount
                            AmountAllocated = Math.Round((((QtyOut - QtyAllocated) * rowInvPool.amount) / rowInvPool.qty), vDefPrec, MidpointRounding.AwayFromZero)
                            Amount += AmountAllocated
                            'update qty + amount pada inventory pool
                            rowInvPool.amount -= AmountAllocated
                            rowInvPool.qty -= (QtyOut - QtyAllocated)
                            'update dummy string
                            If CostingMethod <> ClsTransaksiMonthlyClosingItem.EnumCostingMethod.Average Then
                                If DummyString = "" Then
                                    DummyString = rowInvPool.tipe_transaksi_asal.ToString + "," + rowInvPool.Primary_Transaksi_Asal.ToString + "," _
                                    + Me.ObjCD.BuangNolDibelakangDecimal((QtyOut - QtyAllocated).ToString) + "," + Me.ObjCD.BuangNolDibelakangDecimal(AmountAllocated.ToString)
                                Else
                                    DummyString = DummyString + ";" + rowInvPool.tipe_transaksi_asal.ToString + "," + rowInvPool.Primary_Transaksi_Asal.ToString + "," _
                                    + Me.ObjCD.BuangNolDibelakangDecimal((QtyOut - QtyAllocated).ToString) + "," + Me.ObjCD.BuangNolDibelakangDecimal(AmountAllocated.ToString)
                                End If
                            End If
                            'update qty allocated
                            QtyAllocated = QtyOut
                        Else
                            'ambil 1 baris pada inventory pool secara keseluruhan
                            'update qty allocated + amount
                            QtyAllocated += rowInvPool.qty
                            Amount += rowInvPool.amount
                            'update dummy string
                            If CostingMethod <> ClsTransaksiMonthlyClosingItem.EnumCostingMethod.Average Then
                                If DummyString = "" Then
                                    DummyString = rowInvPool.tipe_transaksi_asal.ToString + "," + rowInvPool.Primary_Transaksi_Asal.ToString + "," _
                                    + Me.ObjCD.BuangNolDibelakangDecimal(rowInvPool.qty.ToString) + "," + Me.ObjCD.BuangNolDibelakangDecimal(rowInvPool.amount.ToString)
                                Else
                                    DummyString = DummyString + ";" + rowInvPool.tipe_transaksi_asal.ToString + "," + rowInvPool.Primary_Transaksi_Asal.ToString + "," _
                                    + Me.ObjCD.BuangNolDibelakangDecimal(rowInvPool.qty.ToString) + "," + Me.ObjCD.BuangNolDibelakangDecimal(rowInvPool.amount.ToString)
                                End If
                            End If
                            'buang baris inventory pool
                            TBLInvPool.Rows.Remove(rowInvPool)
                        End If
                    End While
                    vRowRef.amount = Amount
                    vRowRef.unit_price = Math.Round(vRowRef.amount / vRowRef.qty, vDefPrec, MidpointRounding.AwayFromZero)
                    If DummyString.Length > 20000 Then
                        DummyString = DummyString.Substring(0, 20000)
                    End If
                    vRowRef.remarks = DummyString
                Else
                    TBLInvPool.DefaultView.RowFilter = "[primary code no] = " + vRowRef.primary_code_no.ToString
                    'If CostingMethod = ClsTransaksiMonthlyClosingItem.EnumCostingMethod.FIFO Then
                    '    TBLInvPool.DefaultView.Sort = "[tgl transaksi asal], [tipe transaksi asal], [primary transaksi asal], [primary code no]"
                    'Else
                    '    TBLInvPool.DefaultView.Sort = "[tgl transaksi asal] DESC, [tipe transaksi asal], [primary transaksi asal] DESC, [primary code no]"
                    'End If

                    QtyOut = vRowRef.qty
                    QtyAllocated = 0
                    Amount = 0
                    DummyString = "" '--> format untuk dummy string: tipe transaksi asal, primary transaksi asal, qty, amount, dipisahkan oleh ';'

                    While QtyOut > QtyAllocated
                        If TBLInvPool.DefaultView.Count > 0 Then
                            rowInvPool = TBLInvPool.DefaultView(0).Row

                            If rowInvPool.qty < 0 Then
                                AmountAllocated = Math.Round((QtyOut - QtyAllocated) * vRow_Pending2.unit_price, vDefPrec, MidpointRounding.AwayFromZero)
                                Amount += AmountAllocated
                                rowInvPool.amount -= AmountAllocated
                                rowInvPool.qty -= (QtyOut - QtyAllocated)
                                QtyAllocated = QtyOut
                            Else
                                If rowInvPool.qty > QtyOut - QtyAllocated Then
                                    'ambil sebagian saja: qty out - qty allocated
                                    'update amount
                                    AmountAllocated = Math.Round((((QtyOut - QtyAllocated) * rowInvPool.amount) / rowInvPool.qty), vDefPrec, MidpointRounding.AwayFromZero)
                                    Amount += AmountAllocated
                                    'update qty + amount pada inventory pool
                                    rowInvPool.amount -= AmountAllocated
                                    rowInvPool.qty -= (QtyOut - QtyAllocated)
                                    'update qty allocated
                                    QtyAllocated = QtyOut

                                    If CostingMethod <> ClsTransaksiMonthlyClosingItem.EnumCostingMethod.Average Then
                                        If DummyString = "" Then
                                            DummyString = rowInvPool.tipe_transaksi_asal.ToString + "," + rowInvPool.Primary_Transaksi_Asal.ToString + "," _
                                            + Me.ObjCD.BuangNolDibelakangDecimal((QtyOut - QtyAllocated).ToString) + "," + Me.ObjCD.BuangNolDibelakangDecimal(AmountAllocated.ToString)
                                        Else
                                            DummyString = DummyString + ";" + rowInvPool.tipe_transaksi_asal.ToString + "," + rowInvPool.Primary_Transaksi_Asal.ToString + "," _
                                            + Me.ObjCD.BuangNolDibelakangDecimal((QtyOut - QtyAllocated).ToString) + "," + Me.ObjCD.BuangNolDibelakangDecimal(AmountAllocated.ToString)
                                        End If
                                    End If
                                Else
                                    'ambil 1 baris pada inventory pool secara keseluruhan
                                    'update qty allocated + amount
                                    QtyAllocated += rowInvPool.qty
                                    Amount += rowInvPool.amount
                                    'buang baris inventory pool

                                    If CostingMethod <> ClsTransaksiMonthlyClosingItem.EnumCostingMethod.Average Then
                                        If DummyString = "" Then
                                            DummyString = rowInvPool.tipe_transaksi_asal.ToString + "," + rowInvPool.Primary_Transaksi_Asal.ToString + "," _
                                            + Me.ObjCD.BuangNolDibelakangDecimal((QtyOut - QtyAllocated).ToString) + "," + Me.ObjCD.BuangNolDibelakangDecimal(AmountAllocated.ToString)
                                        Else
                                            DummyString = DummyString + ";" + rowInvPool.tipe_transaksi_asal.ToString + "," + rowInvPool.Primary_Transaksi_Asal.ToString + "," _
                                            + Me.ObjCD.BuangNolDibelakangDecimal((QtyOut - QtyAllocated).ToString) + "," + Me.ObjCD.BuangNolDibelakangDecimal(AmountAllocated.ToString)
                                        End If
                                    End If

                                    TBLInvPool.Rows.Remove(rowInvPool)
                                End If
                            End If
                        Else
                            'tambahkan baris inventory pool dengan nilai minus
                            rowInvPool = TBLInvPool.NewtblItemDetailPoolRow
                            With rowInvPool
                                .amount = Math.Round(vRow_Pending2.unit_price * (QtyOut - QtyAllocated), vDefPrec, MidpointRounding.AwayFromZero)
                                .amount = .amount * (-1)
                                .Cost_Perct_Bundling = 0
                                .Overhead = 0
                                .Primary = Me.GetPrimaryMaxInventoryPool(TBLInvPool) + 1
                                .primary_adjustment_type = -1
                                .primary_code_no = vRowRef.primary_code_no
                                .primary_gain_account = -1
                                .primary_loss_account = -1
                                .Primary_transaksi = Me.TxtPrimaryClosing.Text
                                .Primary_Transaksi_Asal = -1
                                .qty = (QtyOut - QtyAllocated)
                                .qty = .qty * (-1)
                                .remarks = ""
                                .return_price = 0
                                .tgl_transaksi_asal = #1/1/1950#
                                .tipe_item_detail = 0
                                .tipe_transaksi_asal = 0
                                .unit_price = vRow_Pending2.unit_price
                            End With
                            TBLInvPool.Rows.Add(rowInvPool)

                            'amount allocated
                            Amount += Math.Abs(rowInvPool.amount)
                            'update qty allocated
                            QtyAllocated = QtyOut
                        End If
                    End While

                    vRowRef.amount = Amount
                    vRowRef.unit_price = Math.Round(vRowRef.amount / vRowRef.qty, vDefPrec, MidpointRounding.AwayFromZero)
                    If DummyString.Length > 20000 Then
                        DummyString = DummyString.Substring(0, 20000)
                    End If
                    vRowRef.remarks = DummyString
                End If

                'buang barisnya dari tbl pending jika ada
                TBL_Pending.DefaultView.RowFilter = "[tipe transaksi asal] = " + vRowRef.tipe_transaksi_asal.ToString + _
                " AND [primary transaksi asal] = " + vRowRef.primary_transaksi_asal.ToString + _
                " AND [primary code no] = " + vRowRef.primary_code_no.ToString + _
                " AND [tipe item detail] = " + vRowRef.tipe_item_detail.ToString

                While TBL_Pending.DefaultView.Count > 0
                    TBL_Pending.Rows.Remove(TBL_Pending.DefaultView(0).Row)
                End While
            End If
        Catch ex As Exception
            MsgBox(ex.Message + " - Dummy string length: " + DummyString.Length.ToString, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanDataAwalInventoryPool(ByVal TBLInvPool As DseMonthlyClosingNew.tblItemDetailPoolDataTable)
        Dim rowInvPool As DseMonthlyClosingNew.tblItemDetailPoolRow

        Try
            Me.Calc_CurrentStep += 1
            Me.LblStep.Text = "Step " + Me.Calc_CurrentStep.ToString + " of " + Me.Calc_TotalStep.ToString + " : creating inventory pool data"
            Me.ProgressBar1.Value = 0
            Me.LblProgress.Text = ""
            'Me.LblProgress.Visible = True

            TBLInvPool.Rows.Clear()

            Me.ProgressBar1.Maximum = Me.TBLItemDetail.Rows.Count
            For Each Me.rowItemDetail In Me.TBLItemDetail.Rows
                Me.ProgressBar1.Value += 1
                If Math.IEEERemainder(Me.ProgressBar1.Value, 500) = 0 Then
                    Me.LblProgress.Text = "Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString
                End If

                If Me.rowItemDetail.tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.BeginingBalanceItem Then
                    With Me.rowItemDetail
                        rowInvPool = TBLInvPool.NewtblItemDetailPoolRow

                        rowInvPool.amount = .amount
                        rowInvPool.Cost_Perct_Bundling = .Cost_Perct_Bundling
                        rowInvPool.Overhead = .Overhead
                        rowInvPool.Primary = Me.GetPrimaryMaxInventoryPool(TBLInvPool) + 1
                        rowInvPool.primary_adjustment_type = .primary_adjustment_type
                        rowInvPool.primary_code_no = .primary_code_no
                        rowInvPool.primary_gain_account = .primary_gain_account
                        rowInvPool.primary_loss_account = .primary_loss_account
                        rowInvPool.Primary_transaksi = Me.TxtPrimaryClosing.Text
                        rowInvPool.Primary_Transaksi_Asal = .primary_transaksi_asal
                        rowInvPool.qty = .qty
                        rowInvPool.remarks = .remarks
                        rowInvPool.return_price = .return_price
                        rowInvPool.tgl_transaksi_asal = .tgl_transaksi_asal
                        rowInvPool.tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.BeginingBalanceItem
                        rowInvPool.tipe_transaksi_asal = .tipe_transaksi_asal
                        rowInvPool.unit_price = .unit_price
                        If rowInvPool.remarks.Length > 250 Then
                            rowInvPool.remarks = rowInvPool.remarks.Substring(0, 250)
                        End If

                        TBLInvPool.Rows.Add(rowInvPool)
                    End With
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetPrimaryMaxInventoryPool(ByVal TBLInvPool As DseMonthlyClosingNew.tblItemDetailPoolDataTable) As Double
        Dim PrimaryMax As Double = 0
        Dim rowInvPool1 As DseMonthlyClosingNew.tblItemDetailPoolRow

        Try
            'primary paling besar, pasti pada baris yang paling bawah
            If TBLInvPool.Rows.Count > 0 Then
                rowInvPool1 = TBLInvPool.Rows(TBLInvPool.Rows.Count - 1)
                PrimaryMax = rowInvPool1.Primary
            End If

            Return PrimaryMax
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    'Private Sub HitungKeDGVUtama_Obs(ByVal ApakahHitungUlang As Boolean)
    '    Dim ObjSetting As New Core_Module.ClsTransaksiSetting
    '    Dim CostingMethod As Integer = 0
    '    CostingMethod = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
    '                                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.CostingMethod)

    '    Dim tblID As New DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailDataTable
    '    Dim rowID As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailRow

    '    Dim tblIDX As New DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailDataTable
    '    Dim rowIDX As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailRow

    '    Dim DTAJ As New DseMonthlyClosingNewTableAdapters.tbltrans_purchase_closing_item_detail_poolTableAdapter
    '    Dim tblJ As New DseMonthlyClosingNew.tbltrans_purchase_closing_item_detail_poolDataTable
    '    Dim rowJ As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detail_poolRow

    '    Dim DTAB As New DseMonthlyClosingNewTableAdapters.BundlingTableAdapter
    '    Dim tblB As New DseMonthlyClosingNew.BundlingDataTable
    '    Dim rowB As DseMonthlyClosingNew.BundlingRow

    '    Dim DTAPx As New DseMonthlyClosingNewTableAdapters.tbltrans_purchase_closing_item_detail_poolTableAdapter
    '    Dim tblPx As New DseMonthlyClosingNew.tblItemDetailPoolDataTable
    '    Dim rowPX As DseMonthlyClosingNew.tblItemDetailPoolRow


    '    Dim startDate As Date
    '    Dim EndDate As Date

    '    Try
    '        If Me.ApakahAdaPerubahanAdj(Me.TblA, Me.TblA2) = True Then
    '            tblID.Rows.Clear()
    '            tblIDX.Rows.Clear()
    '            tblPool.Rows.Clear()
    '            tblJ.Rows.Clear()
    '            tblB.Rows.Clear()
    '            tblPx.Rows.Clear()
    '            DTAJ.Connection.ConnectionString = Core_Module.GetConstring
    '            DTAB.Connection.ConnectionString = Core_Module.GetConstring

    '            startDate = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value)
    '            EndDate = Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value)
    '            DTAB.Fill(tblB, EndDate, startDate)


    '            Me.TBLItemDetail.DefaultView.RowFilter = Nothing
    '            'Dim x1y As New FormTEST
    '            'x1y.DataGridView1.DataSource = Nothing
    '            'x1y.DataGridView1.DataSource = Me.TBLItemDetail
    '            'x1y.Text = "Me.TBLItemDetail"
    '            'x1y.ShowDialog()

    '            If ApakahHitungUlang = False Then
    '                DTAJ.Fill(tblJ, Me.TxtPrimaryClosing.Text)
    '                For Each rowJ In tblJ.Rows
    '                    With rowJ
    '                        rowPool = tblPool.NewtblItemDetailPoolRow
    '                        rowPool.Primary = .primary
    '                        rowPool.Primary_transaksi = .primary_transaksi
    '                        rowPool.Primary_Transaksi_Asal = .primary_transaksi_asal
    '                        rowPool.tipe_transaksi_asal = .tipe_transaksi_asal
    '                        rowPool.tgl_transaksi_asal = .tgl_transaksi_asal
    '                        rowPool.primary_code_no = .primary_code_no
    '                        rowPool.tipe_item_detail = .tipe_item_detail
    '                        rowPool.unit_price = .unit_price
    '                        rowPool.amount = .amount
    '                        rowPool.qty = .qty
    '                        rowPool.return_price = .return_price
    '                        rowPool.primary_gain_account = .primary_gain_account
    '                        rowPool.primary_loss_account = .primary_loss_account
    '                        rowPool.primary_adjustment_type = .primary_adjustment_type
    '                        rowPool.remarks = .remarks
    '                        rowPool.Cost_Perct_Bundling = .Cost_Perct_Bundling
    '                        rowPool.Overhead = .Overhead
    '                        tblPool.Rows.Add(rowPool)
    '                    End With
    '                Next


    '                Dim ListX As New List(Of Double)
    '                For Each Me.rowItemDetail In Me.TBLItemDetail.Rows
    '                    With Me.rowItemDetail
    '                        If ListX.Contains(.primary_code_no) = False Then
    '                            ListX.Add(.primary_code_no)
    '                        End If
    '                    End With
    '                Next

    '                Dim rowDgv As New DataGridViewRow
    '                If ListX.Count > 0 Then
    '                    For intX As Integer = 0 To ListX.Count - 1
    '                        Me.dgvItem.Rows.Add()
    '                        rowDgv = Me.dgvItem.Rows(intX)
    '                        Me.TBLProduct.DefaultView.RowFilter = Nothing
    '                        Me.TBLProduct.DefaultView.RowFilter = "[Primary] = " & ListX.Item(intX)
    '                        If Me.TBLProduct.DefaultView.Count > 0 Then
    '                            Me.rowProduct = Me.TBLProduct.DefaultView(0).Row
    '                            Dim Decimalx As Integer = 0
    '                            Decimalx = Me.rowProduct._Decimal
    '                            rowDgv.Cells(Me.ColPrimaryCodeNo.Index).Value = Me.rowProduct.Primary
    '                            rowDgv.Cells(Me.ColCodeNo.Index).Value = Me.rowProduct.Product_Code
    '                            rowDgv.Cells(Me.ColDescription.Index).Value = Me.rowProduct.Description
    '                            rowDgv.Cells(Me.ColPrimaryUnits.Index).Value = Me.rowProduct.Primary_Unit
    '                            rowDgv.Cells(Me.ColUnits.Index).Value = Me.rowProduct.Unit_Code
    '                            rowDgv.Cells(Me.ColItem_PrimaryAdjustmentAccount.Index).Value = Me.rowProduct.Primary_Adjustment_Account
    '                            rowDgv.Cells(Me.ColItem_PrimaryCOGSAccount.Index).Value = Me.rowProduct.Primary_COGS_Account
    '                            rowDgv.Cells(Me.ColItem_PrimaryInventoryAccount.Index).Value = Me.rowProduct.Primary_Inventory_Account
    '                            rowDgv.Cells(Me.ColItem_PrimarySalesAccount.Index).Value = Me.rowProduct.Primary_Sales_Account
    '                            rowDgv.Cells(Me.ColItem_PrimarySalesReturnAccount.Index).Value = Me.rowProduct.Primary_Sales_Return_Account
    '                            rowDgv.Cells(Me.ColItem_PrimaryStockTakingAccount.Index).Value = Me.rowProduct.Primary_Stock_Taking_Account

    '                            Dim decQtyBegin As Decimal = 0
    '                            Dim decAmountBegin As Decimal = 0
    '                            Dim decQtyIn As Decimal = 0
    '                            Dim decAmountIn As Decimal = 0
    '                            Dim decQtyOut As Decimal = 0
    '                            Dim decAmountOut As Decimal = 0
    '                            Dim decQtyEnd As Decimal = 0
    '                            Dim decAmountEnd As Decimal = 0
    '                            Me.TBLItemDetail.DefaultView.RowFilter = Nothing
    '                            Me.TBLItemDetail.DefaultView.RowFilter = "[Primary Code No] = " & ListX.Item(intX)
    '                            Me.'TBLItemDetail.DefaultView.Sort = "[Primary Code No] "
    '                            If Me.TBLItemDetail.DefaultView.Count > 0 Then
    '                                For Each dvx As Data.DataRowView In Me.TBLItemDetail.DefaultView
    '                                    Me.rowItemDetail = dvx.Row
    '                                    With rowItemDetail
    '                                        If .tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.BeginingBalanceItem Then
    '                                            If IsNumeric(.qty) Then
    '                                                decQtyBegin += CDec(.qty)
    '                                                decAmountBegin += CDec(.amount)
    '                                            End If
    '                                        ElseIf .tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem Then
    '                                            If IsNumeric(.qty) Then
    '                                                decQtyIn += CDec(.qty)
    '                                                decAmountIn += CDec(.amount)
    '                                            End If
    '                                        ElseIf .tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem Then
    '                                            If IsNumeric(.qty) Then
    '                                                decQtyOut += CDec(.qty)
    '                                                decAmountOut += CDec(.amount)
    '                                            End If
    '                                        ElseIf .tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.EndingBalanceItem Then
    '                                            If IsNumeric(.qty) Then
    '                                                decQtyEnd += CDec(.qty)
    '                                                decAmountEnd += CDec(.amount)
    '                                            End If
    '                                        End If
    '                                    End With
    '                                Next
    '                            End If
    '                            rowDgv.Cells(Me.ColBeginQty.Index).Value = Me.ObjCD.DigitGrouping(decQtyBegin, Decimalx)
    '                            rowDgv.Cells(Me.ColBeginAmount.Index).Value = Me.ObjCD.DigitGrouping(decAmountBegin, Me.DefaultPrecision)
    '                            rowDgv.Cells(Me.ColInQty.Index).Value = Me.ObjCD.DigitGrouping(decQtyIn, Decimalx)
    '                            rowDgv.Cells(Me.ColInAmount.Index).Value = Me.ObjCD.DigitGrouping(decAmountIn, Me.DefaultPrecision)
    '                            rowDgv.Cells(Me.ColOutQty.Index).Value = Me.ObjCD.DigitGrouping(decQtyOut, Decimalx)
    '                            rowDgv.Cells(Me.ColOutAmount.Index).Value = Me.ObjCD.DigitGrouping(decAmountOut, Me.DefaultPrecision)
    '                            rowDgv.Cells(Me.ColEndQty.Index).Value = Me.ObjCD.DigitGrouping(decQtyEnd, Decimalx)
    '                            rowDgv.Cells(Me.ColEndAmount.Index).Value = Me.ObjCD.DigitGrouping(decAmountEnd, Me.DefaultPrecision)
    '                        End If
    '                    Next
    '                End If

    '            Else
    '                Me.TBLItemDetail.DefaultView.RowFilter = Nothing
    '                'Me.TBLItemDetail.DefaultView.RowFilter = " [Primary Code No] = 692"
    '                Me.'TBLItemDetail.DefaultView.Sort = " [tgl transaksi asal] "

    '                'Dim C As New FormTEST
    '                'C.DataGridView1.DataSource = Nothing
    '                'C.DataGridView1.DataSource = Me.TBLBefore
    '                'C.Text = "tblbefore"
    '                'C.ShowDialog()

    '                If Me.TBLItemDetail.DefaultView.Count > 0 Then
    '                    Dim Index As Integer = 0
    '                    For Each dv As Data.DataRowView In Me.TBLItemDetail.DefaultView
    '                        Me.rowItemDetail = dv.Row
    '                        With Me.rowItemDetail
    '                            Dim decQty As Decimal = 0
    '                            decQty = .qty
    '                            If (rowItemDetail.tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.BeginingBalanceItem _
    '                            OrElse rowItemDetail.tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem) AndAlso _
    '                            rowItemDetail.tipe_transaksi_asal <> ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling Then
    '                                '------------------------------------------------------------------------------------------------
    '                                '-----------------> Fill dulu qty in
    '                                '------------------------------------------------------------------------------------------------
    '                                rowPool = tblPool.NewtblItemDetailPoolRow
    '                                rowPool.Primary = .primary
    '                                rowPool.Primary_transaksi = .primary_transaksi
    '                                rowPool.Primary_Transaksi_Asal = .primary_transaksi_asal
    '                                rowPool.tipe_transaksi_asal = .tipe_transaksi_asal
    '                                rowPool.tgl_transaksi_asal = .tgl_transaksi_asal
    '                                rowPool.primary_code_no = .primary_code_no
    '                                rowPool.tipe_item_detail = .tipe_item_detail
    '                                rowPool.unit_price = .unit_price
    '                                rowPool.amount = .amount
    '                                rowPool.qty = .qty
    '                                rowPool.return_price = .return_price
    '                                rowPool.primary_gain_account = .primary_gain_account
    '                                rowPool.primary_loss_account = .primary_loss_account
    '                                rowPool.primary_adjustment_type = .primary_adjustment_type
    '                                rowPool.remarks = .remarks
    '                                rowPool.Cost_Perct_Bundling = .Cost_Perct_Bundling
    '                                rowPool.Overhead = .Overhead
    '                                Index += 1
    '                                tblPool.Rows.Add(rowPool)

    '                                rowID = tblID.Newtbltrans_purchase_closing_item_detailRow
    '                                rowID.primary = .primary
    '                                rowID.primary_transaksi = .primary_transaksi
    '                                rowID.primary_transaksi_asal = .primary_transaksi_asal
    '                                rowID.tipe_transaksi_asal = .tipe_transaksi_asal
    '                                rowID.tgl_transaksi_asal = .tgl_transaksi_asal
    '                                rowID.primary_code_no = .primary_code_no
    '                                rowID.tipe_item_detail = .tipe_item_detail
    '                                rowID.unit_price = .unit_price
    '                                rowID.amount = .amount
    '                                rowID.qty = .qty
    '                                rowID.return_price = .return_price
    '                                rowID.primary_gain_account = .primary_gain_account
    '                                rowID.primary_loss_account = .primary_loss_account
    '                                rowID.primary_adjustment_type = .primary_adjustment_type
    '                                rowID.remarks = .remarks
    '                                rowID.Cost_Perct_Bundling = .Cost_Perct_Bundling
    '                                rowID.Overhead = .Overhead
    '                                tblID.Rows.Add(rowID)

    '                                rowIDX = tblIDX.Newtbltrans_purchase_closing_item_detailRow
    '                                rowIDX.primary = .primary
    '                                rowIDX.primary_transaksi = .primary_transaksi
    '                                rowIDX.primary_transaksi_asal = .primary_transaksi_asal
    '                                rowIDX.tipe_transaksi_asal = .tipe_transaksi_asal
    '                                rowIDX.tgl_transaksi_asal = .tgl_transaksi_asal
    '                                rowIDX.primary_code_no = .primary_code_no
    '                                rowIDX.tipe_item_detail = .tipe_item_detail
    '                                rowIDX.unit_price = .unit_price
    '                                rowIDX.amount = .amount
    '                                rowIDX.qty = .qty
    '                                rowIDX.return_price = .return_price
    '                                rowIDX.primary_gain_account = .primary_gain_account
    '                                rowIDX.primary_loss_account = .primary_loss_account
    '                                rowIDX.primary_adjustment_type = .primary_adjustment_type
    '                                rowIDX.remarks = .remarks
    '                                rowIDX.Cost_Perct_Bundling = .Cost_Perct_Bundling
    '                                rowIDX.Overhead = .Overhead
    '                                tblIDX.Rows.Add(rowIDX)

    '                                '------------------------------------------------------------------------------------------------
    '                                '-----------------> cek dulu apakah ada price = 0
    '                                '------------------------------------------------------------------------------------------------
    '                                tblIDX.DefaultView.RowFilter = Nothing
    '                                tblIDX.DefaultView.RowFilter = " [Primary Code No] = " & .primary_code_no & _
    '                                                              " And [tipe item detail] = " & ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem & _
    '                                                              " And [unit price] = 0 "
    '                                If tblIDX.DefaultView.Count > 0 Then
    '                                    'cek terlebih dahulu apakah masih ada pencatatan belum selesai di transaksi sebelumnya
    '                                    ' hitung previous data yg belum ada nilainya
    '                                    Me.CalculatePrevDataNol(tblPool, tblPx, tblIDX, tblID, tblB, .primary_code_no, CostingMethod)
    '                                End If
    '                            ElseIf rowItemDetail.tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem Then
    '                                '------------------------------------------------------------------------------------------------
    '                                '-----------------> Perhitungkan out normal
    '                                '------------------------------------------------------------------------------------------------
    '                                tblPool.DefaultView.RowFilter = Nothing
    '                                tblPool.DefaultView.RowFilter = " [primary code no] = " & .primary_code_no
    '                                If CostingMethod = ClsTransaksiMonthlyClosingItem.EnumCostingMethod.LIFO Then
    '                                    tblPool.DefaultView.Sort = " [tgl transaksi asal] desc "
    '                                Else
    '                                    tblPool.DefaultView.Sort = " [tgl transaksi asal] asc "
    '                                End If
    '                                If Me.ApakahAkanMinus(tblPool, .primary_code_no, decQty) = True Then
    '                                    rowIDX = tblIDX.Newtbltrans_purchase_closing_item_detailRow
    '                                    rowIDX.primary = .primary
    '                                    rowIDX.primary_transaksi = .primary_transaksi
    '                                    rowIDX.primary_transaksi_asal = .primary_transaksi_asal
    '                                    rowIDX.tipe_transaksi_asal = .tipe_transaksi_asal
    '                                    rowIDX.tgl_transaksi_asal = .tgl_transaksi_asal
    '                                    rowIDX.primary_code_no = .primary_code_no
    '                                    rowIDX.tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
    '                                    rowIDX.unit_price = 0
    '                                    rowIDX.amount = 0
    '                                    rowIDX.qty = .qty
    '                                    rowIDX.return_price = .return_price
    '                                    rowIDX.primary_gain_account = .primary_gain_account
    '                                    rowIDX.primary_loss_account = .primary_loss_account
    '                                    rowIDX.primary_adjustment_type = .primary_adjustment_type
    '                                    rowIDX.remarks = .remarks
    '                                    rowIDX.Overhead = .Overhead
    '                                    rowIDX.Cost_Perct_Bundling = .Cost_Perct_Bundling
    '                                    tblIDX.Rows.Add(rowIDX)
    '                                Else
    '                                    tblPool.DefaultView.RowFilter = Nothing
    '                                    tblPool.DefaultView.RowFilter = " [primary code no] = " & .primary_code_no
    '                                    If CostingMethod = ClsTransaksiMonthlyClosingItem.EnumCostingMethod.LIFO Then
    '                                        tblPool.DefaultView.Sort = " [tgl transaksi asal] desc "
    '                                    Else
    '                                        tblPool.DefaultView.Sort = " [tgl transaksi asal] asc "
    '                                    End If
    '                                    If tblPool.DefaultView.Count > 0 Then
    '                                        If CostingMethod = ClsTransaksiMonthlyClosingItem.EnumCostingMethod.Average Then
    '                                            Dim decPrice As Decimal = 0
    '                                            decPrice = Me.GetNilaiAVG(tblPool, .primary_code_no)
    '                                            rowID = tblID.Newtbltrans_purchase_closing_item_detailRow
    '                                            rowID.primary = .primary
    '                                            rowID.primary_transaksi = .primary_transaksi
    '                                            rowID.primary_transaksi_asal = .primary_transaksi_asal
    '                                            rowID.tipe_transaksi_asal = .tipe_transaksi_asal
    '                                            rowID.tgl_transaksi_asal = .tgl_transaksi_asal
    '                                            rowID.primary_code_no = .primary_code_no
    '                                            rowID.tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
    '                                            rowID.qty = decQty
    '                                            rowID.unit_price = decPrice
    '                                            rowID.amount = rowID.qty * rowID.unit_price
    '                                            rowID.return_price = .return_price
    '                                            rowID.primary_gain_account = .primary_gain_account
    '                                            rowID.primary_loss_account = .primary_loss_account
    '                                            rowID.primary_adjustment_type = .primary_adjustment_type
    '                                            rowID.remarks = .remarks
    '                                            rowID.Cost_Perct_Bundling = .Cost_Perct_Bundling
    '                                            rowID.Overhead = .Overhead
    '                                            tblID.Rows.Add(rowID)

    '                                            While decQty > 0
    '                                                tblPool.DefaultView.RowFilter = Nothing
    '                                                tblPool.DefaultView.RowFilter = " [primary code no] = " & .primary_code_no
    '                                                tblPool.DefaultView.Sort = " [tgl transaksi asal] asc "

    '                                                For intA As Integer = 0 To tblPool.DefaultView.Count - 1
    '                                                    rowPool = tblPool.DefaultView(intA).Row
    '                                                    If rowPool.qty > decQty Then
    '                                                        rowPool.qty -= decQty
    '                                                        rowPool.unit_price = decPrice
    '                                                        rowPool.amount = rowPool.unit_price * rowPool.qty
    '                                                        decQty = 0
    '                                                        Exit While
    '                                                    Else
    '                                                        decQty -= rowPool.qty
    '                                                        rowPool.qty = 0
    '                                                        rowPool.unit_price = decPrice
    '                                                        rowPool.amount = rowPool.unit_price * rowPool.qty
    '                                                    End If

    '                                                    If decQty = 0 Then
    '                                                        Exit While
    '                                                    End If
    '                                                Next
    '                                                Me.HapustblPool(tblPool, .primary_code_no, CostingMethod)
    '                                            End While
    '                                            Me.HapustblPool(tblPool, .primary_code_no, CostingMethod)
    '                                        Else
    '                                            While decQty > 0
    '                                                tblPool.DefaultView.RowFilter = Nothing
    '                                                tblPool.DefaultView.RowFilter = " [primary code no] = " & .primary_code_no
    '                                                If CostingMethod = ClsTransaksiMonthlyClosingItem.EnumCostingMethod.LIFO Then
    '                                                    tblPool.DefaultView.Sort = " [tgl transaksi asal] desc "
    '                                                Else
    '                                                    tblPool.DefaultView.Sort = " [tgl transaksi asal] asc "
    '                                                End If

    '                                                If tblPool.DefaultView.Count > 0 Then
    '                                                    For intA As Integer = 0 To tblPool.DefaultView.Count - 1
    '                                                        rowPool = tblPool.DefaultView(intA).Row
    '                                                        If rowPool.qty > decQty Then
    '                                                            rowID = tblID.Newtbltrans_purchase_closing_item_detailRow
    '                                                            rowID.primary = .primary
    '                                                            rowID.primary_transaksi = .primary_transaksi
    '                                                            rowID.primary_transaksi_asal = .primary_transaksi_asal
    '                                                            rowID.tipe_transaksi_asal = .tipe_transaksi_asal
    '                                                            rowID.tgl_transaksi_asal = .tgl_transaksi_asal
    '                                                            rowID.primary_code_no = .primary_code_no
    '                                                            rowID.tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
    '                                                            rowID.qty = decQty
    '                                                            rowID.unit_price = rowPool.unit_price
    '                                                            rowID.amount = decQty * rowID.unit_price
    '                                                            rowID.return_price = .return_price
    '                                                            rowID.primary_gain_account = .primary_gain_account
    '                                                            rowID.primary_loss_account = .primary_loss_account
    '                                                            rowID.primary_adjustment_type = .primary_adjustment_type
    '                                                            rowID.remarks = .remarks
    '                                                            rowID.Cost_Perct_Bundling = .Cost_Perct_Bundling
    '                                                            rowID.Overhead = .Overhead
    '                                                            tblID.Rows.Add(rowID)
    '                                                            'kurangi qty pool 
    '                                                            rowPool.qty -= decQty
    '                                                            rowPool.amount = rowPool.qty * rowPool.unit_price
    '                                                            'decQty habis dipakai
    '                                                            decQty = 0
    '                                                        Else
    '                                                            rowID = tblID.Newtbltrans_purchase_closing_item_detailRow
    '                                                            rowID.primary = .primary
    '                                                            rowID.primary_transaksi = .primary_transaksi
    '                                                            rowID.primary_transaksi_asal = .primary_transaksi_asal
    '                                                            rowID.tipe_transaksi_asal = .tipe_transaksi_asal
    '                                                            rowID.tgl_transaksi_asal = .tgl_transaksi_asal
    '                                                            rowID.primary_code_no = .primary_code_no
    '                                                            rowID.tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
    '                                                            rowID.qty = rowPool.qty
    '                                                            rowID.unit_price = rowPool.unit_price
    '                                                            rowID.amount = rowID.qty * rowID.unit_price
    '                                                            rowID.return_price = .return_price
    '                                                            rowID.primary_gain_account = .primary_gain_account
    '                                                            rowID.primary_loss_account = .primary_loss_account
    '                                                            rowID.primary_adjustment_type = .primary_adjustment_type
    '                                                            rowID.remarks = .remarks
    '                                                            rowID.Cost_Perct_Bundling = .Cost_Perct_Bundling
    '                                                            rowID.Overhead = .Overhead
    '                                                            tblID.Rows.Add(rowID)
    '                                                            'decQtyID kurangi qty pool
    '                                                            decQty -= rowPool.qty
    '                                                            'qty pool habis
    '                                                            rowPool.qty = 0
    '                                                            rowPool.amount = 0
    '                                                        End If

    '                                                        If decQty = 0 Then
    '                                                            Exit While
    '                                                        End If
    '                                                    Next
    '                                                    Me.HapustblPool(tblPool, .primary_code_no, CostingMethod)
    '                                                End If
    '                                            End While
    '                                            Me.HapustblPool(tblPool, .primary_code_no, CostingMethod)
    '                                        End If
    '                                    End If
    '                                    If .tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling Then
    '                                        Me.UpdatePriceOutBundling(.primary_transaksi_asal, .primary_code_no, tblB, tblID)
    '                                        If Me.ApakahOutBundlingLengkap(.primary_transaksi_asal, tblB) = True Then
    '                                            Me.calculatePriceInBundling(.primary_transaksi_asal, tblB, tblPool, tblPx, tblID, tblIDX, CostingMethod)
    '                                        End If
    '                                    End If
    '                                End If
    '                            End If
    '                        End With
    '                    Next
    '                End If

    '                tblPool.DefaultView.RowFilter = Nothing
    '                For Each rowPool In tblPool.Rows
    '                    With rowPool
    '                        rowID = tblID.Newtbltrans_purchase_closing_item_detailRow
    '                        rowID.primary = .Primary
    '                        rowID.primary_transaksi = .Primary_transaksi
    '                        rowID.primary_transaksi_asal = .Primary_Transaksi_Asal
    '                        rowID.tipe_transaksi_asal = .tipe_transaksi_asal
    '                        rowID.tgl_transaksi_asal = rowPool.tgl_transaksi_asal
    '                        rowID.primary_code_no = .primary_code_no
    '                        rowID.tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.EndingBalanceItem
    '                        rowID.qty = rowPool.qty
    '                        rowID.unit_price = rowPool.unit_price
    '                        rowID.amount = rowID.qty * rowID.unit_price
    '                        rowID.return_price = .return_price
    '                        rowID.primary_gain_account = .primary_gain_account
    '                        rowID.primary_loss_account = .primary_loss_account
    '                        rowID.primary_adjustment_type = .primary_adjustment_type
    '                        rowID.Cost_Perct_Bundling = .Cost_Perct_Bundling
    '                        rowID.Overhead = .Overhead
    '                        rowID.remarks = .remarks
    '                        tblID.Rows.Add(rowID)
    '                    End With
    '                Next

    '                Me.TBLItemDetail.Rows.Clear()
    '                Dim ListX As New List(Of Double)
    '                tblID.DefaultView.RowFilter = Nothing
    '                If tblID.DefaultView.Count > 0 Then
    '                    For Each rowID In tblID.Rows
    '                        With rowID
    '                            Me.rowItemDetail = Me.TBLItemDetail.Newtbltrans_purchase_closing_item_detailRow
    '                            Me.rowItemDetail.primary = .primary
    '                            Me.rowItemDetail.primary_transaksi = .primary_transaksi
    '                            Me.rowItemDetail.primary_transaksi_asal = .primary_transaksi_asal
    '                            Me.rowItemDetail.tipe_transaksi_asal = .tipe_transaksi_asal
    '                            Me.rowItemDetail.tgl_transaksi_asal = .tgl_transaksi_asal
    '                            Me.rowItemDetail.primary_code_no = .primary_code_no
    '                            Me.rowItemDetail.tipe_item_detail = .tipe_item_detail
    '                            Me.rowItemDetail.qty = .qty
    '                            Me.rowItemDetail.unit_price = .unit_price
    '                            Me.rowItemDetail.amount = .amount
    '                            Me.rowItemDetail.return_price = .return_price
    '                            Me.rowItemDetail.primary_gain_account = .primary_gain_account
    '                            Me.rowItemDetail.primary_loss_account = .primary_loss_account
    '                            Me.rowItemDetail.primary_adjustment_type = .primary_adjustment_type
    '                            Me.rowItemDetail.remarks = .remarks
    '                            Me.rowItemDetail.Cost_Perct_Bundling = .Cost_Perct_Bundling
    '                            Me.rowItemDetail.Overhead = .Overhead
    '                            Me.TBLItemDetail.Rows.Add(Me.rowItemDetail)
    '                            If .primary_code_no <> -1 Then
    '                                If ListX.Contains(.primary_code_no) = False Then
    '                                    ListX.Add(.primary_code_no)
    '                                End If
    '                            End If
    '                        End With
    '                    Next
    '                End If

    '                Dim rowDgv As New DataGridViewRow
    '                If ListX.Count > 0 Then
    '                    Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
    '                    Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

    '                    For intX As Integer = 0 To ListX.Count - 1
    '                        Me.TBLProduct.DefaultView.RowFilter = Nothing
    '                        Me.TBLProduct.DefaultView.RowFilter = "[Primary] = " & ListX.Item(intX)
    '                        If Me.TBLProduct.DefaultView.Count > 0 Then
    '                            Me.rowProduct = Me.TBLProduct.DefaultView(0).Row
    '                            Dim Decimalx As Integer = 0
    '                            Decimalx = Me.rowProduct._Decimal
    '                            Me.dgvItem.Rows.Add()
    '                            rowDgv = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
    '                            rowDgv.Cells(Me.ColPrimaryCodeNo.Index).Value = Me.rowProduct.Primary
    '                            rowDgv.Cells(Me.ColCodeNo.Index).Value = Me.rowProduct.Product_Code
    '                            rowDgv.Cells(Me.ColDescription.Index).Value = Me.rowProduct.Description
    '                            rowDgv.Cells(Me.ColPrimaryUnits.Index).Value = Me.rowProduct.Primary_Unit
    '                            rowDgv.Cells(Me.ColUnits.Index).Value = Me.rowProduct.Unit_Code
    '                            rowDgv.Cells(Me.ColItem_PrimaryAdjustmentAccount.Index).Value = Me.rowProduct.Primary_Adjustment_Account
    '                            rowDgv.Cells(Me.ColItem_PrimaryCOGSAccount.Index).Value = Me.rowProduct.Primary_COGS_Account
    '                            rowDgv.Cells(Me.ColItem_PrimaryInventoryAccount.Index).Value = Me.rowProduct.Primary_Inventory_Account
    '                            rowDgv.Cells(Me.ColItem_PrimarySalesAccount.Index).Value = Me.rowProduct.Primary_Sales_Account
    '                            rowDgv.Cells(Me.ColItem_PrimarySalesReturnAccount.Index).Value = Me.rowProduct.Primary_Sales_Return_Account
    '                            rowDgv.Cells(Me.ColItem_PrimaryStockTakingAccount.Index).Value = Me.rowProduct.Primary_Stock_Taking_Account

    '                            Dim decQtyBegin As Decimal = 0
    '                            Dim decAmountBegin As Decimal = 0
    '                            Dim decQtyIn As Decimal = 0
    '                            Dim decAmountIn As Decimal = 0
    '                            Dim decQtyOut As Decimal = 0
    '                            Dim decAmountOut As Decimal = 0
    '                            Dim decQtyEnd As Decimal = 0
    '                            Dim decAmountEnd As Decimal = 0
    '                            Me.TBLItemDetail.DefaultView.RowFilter = Nothing
    '                            Me.TBLItemDetail.DefaultView.RowFilter = "[Primary Code No] = " & ListX.Item(intX)
    '                            Me.'TBLItemDetail.DefaultView.Sort = "[Primary Code No] "
    '                            If Me.TBLItemDetail.DefaultView.Count > 0 Then
    '                                For Each dvx As Data.DataRowView In Me.TBLItemDetail.DefaultView
    '                                    Me.rowItemDetail = dvx.Row
    '                                    With rowItemDetail
    '                                        If .tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.BeginingBalanceItem Then
    '                                            If IsNumeric(.qty) Then
    '                                                decQtyBegin += CDec(.qty)
    '                                                decAmountBegin += CDec(.amount)
    '                                            End If
    '                                        ElseIf .tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem Then
    '                                            If IsNumeric(.qty) Then
    '                                                decQtyIn += CDec(.qty)
    '                                                decAmountIn += CDec(.amount)
    '                                            End If
    '                                        ElseIf .tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem Then
    '                                            If IsNumeric(.qty) Then
    '                                                decQtyOut += CDec(.qty)
    '                                                decAmountOut += CDec(.amount)
    '                                            End If
    '                                        ElseIf .tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.EndingBalanceItem Then
    '                                            If IsNumeric(.qty) Then
    '                                                decQtyEnd += CDec(.qty)
    '                                                decAmountEnd += CDec(.amount)
    '                                            End If
    '                                        End If
    '                                    End With
    '                                Next
    '                            End If
    '                            rowDgv.Cells(Me.ColBeginQty.Index).Value = Me.ObjCD.DigitGrouping(decQtyBegin, Decimalx)
    '                            rowDgv.Cells(Me.ColBeginAmount.Index).Value = Me.ObjCD.DigitGrouping(decAmountBegin, Me.DefaultPrecision)
    '                            rowDgv.Cells(Me.ColInQty.Index).Value = Me.ObjCD.DigitGrouping(decQtyIn, Decimalx)
    '                            rowDgv.Cells(Me.ColInAmount.Index).Value = Me.ObjCD.DigitGrouping(decAmountIn, Me.DefaultPrecision)
    '                            rowDgv.Cells(Me.ColOutQty.Index).Value = Me.ObjCD.DigitGrouping(decQtyOut, Decimalx)
    '                            rowDgv.Cells(Me.ColOutAmount.Index).Value = Me.ObjCD.DigitGrouping(decAmountOut, Me.DefaultPrecision)
    '                            rowDgv.Cells(Me.ColEndQty.Index).Value = Me.ObjCD.DigitGrouping(decQtyEnd, Decimalx)
    '                            rowDgv.Cells(Me.ColEndAmount.Index).Value = Me.ObjCD.DigitGrouping(decAmountEnd, Me.DefaultPrecision)
    '                        End If
    '                    Next
    '                    Me.ObjCD.SetNumber(Me.dgvItem)

    '                    Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
    '                    Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

    '                End If
    '            End If
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        tblIDX.Dispose()
    '        tblID.Dispose()
    '    End Try
    'End Sub

    Private Sub UpdatePriceOutBundling(ByVal vprimaryTransaksiAsal As Double, ByVal vprimaryProductCode As Double, _
    ByRef vtblB As DseMonthlyClosingNew.BundlingDataTable, ByVal vtblID As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailDataTable)
        Dim vrowB As DseMonthlyClosingNew.BundlingRow
        Dim vrowID As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailRow

        Try
            vtblB.DefaultView.RowFilter = Nothing
            vtblB.DefaultView.RowFilter = " [Primary Transaksi] = " & vprimaryTransaksiAsal
            If vtblB.DefaultView.Count > 0 Then
                For Each DVB As Data.DataRowView In vtblB.DefaultView
                    vrowB = DVB.Row
                    vtblID.DefaultView.RowFilter = Nothing
                    vtblID.DefaultView.RowFilter = "[Primary Transaksi asal] = " & vrowB.Primary_Transaksi & " And [Primary Code No] = " & vrowB.Primary_Code_B4
                    Dim decQty As Decimal = 0
                    Dim decPrice As Decimal = 0
                    Dim decAmount As Decimal = 0
                    If vtblID.DefaultView.Count > 0 Then
                        For Each dvID As Data.DataRowView In vtblID.DefaultView
                            vrowID = dvID.Row
                            decQty += vrowID.qty
                            decAmount += vrowID.amount
                        Next
                        If decQty <> 0 Then
                            decPrice = Me.ObjCD.DigitGrouping(decAmount / decQty, Me.DefaultPrecision)
                        End If
                    End If
                    vrowB.Price_B4 = Me.ObjCD.DigitGrouping(decPrice, Me.DefaultPrecision)
                    vrowB.Amount_B4 = vrowB.Price_B4 * vrowB.Qty_B4
                Next
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Private Sub CalculatePrevDataNol(ByVal vtblPool As DseMonthlyClosingNew.tblItemDetailPoolDataTable, _
    '                                 ByVal vtblPx As DseMonthlyClosingNew.tblItemDetailPoolDataTable, _
    '                                 ByRef vtblIDx As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailDataTable, _
    '                                 ByRef vtblID As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailDataTable, _
    '                                 ByRef vtblB As DseMonthlyClosingNew.BundlingDataTable, _
    '                                 ByVal vPrimaryCodeNo As Double, ByVal CostingMethod As Integer)
    '    Dim vrowIDX As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailRow
    '    Dim vrowID As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailRow
    '    Dim vrowP As DseMonthlyClosingNew.tblItemDetailPoolRow
    '    Dim vrowPX As DseMonthlyClosingNew.tblItemDetailPoolRow
    '    Try


    '        vtblIDx.DefaultView.RowFilter = Nothing
    '        vtblIDx.DefaultView.RowFilter = " [Primary Code No] = " & vPrimaryCodeNo & _
    '                                      " And [tipe item detail] = " & ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem & _
    '                                      " And [unit price] = 0 "

    '        If vtblIDx.DefaultView.Count > 0 Then
    '            For intIDx As Integer = 0 To vtblIDx.DefaultView.Count - 1
    '                vrowIDX = vtblIDx.DefaultView(intIDx).Row
    '                Dim decPrice As Decimal = 0
    '                Dim decQtyID As Decimal = 0
    '                decQtyID = vrowIDX.qty

    '                vtblPool.DefaultView.RowFilter = Nothing
    '                vtblPool.DefaultView.RowFilter = " [primary code no] = " & vrowIDX.primary_code_no
    '                If CostingMethod = ClsTransaksiMonthlyClosingItem.EnumCostingMethod.LIFO Then
    '                    vtblPool.DefaultView.Sort = " [tgl transaksi asal] desc "
    '                Else
    '                    vtblPool.DefaultView.Sort = " [tgl transaksi asal] asc "
    '                End If

    '                If Me.ApakahAkanMinus(vtblPool, vrowIDX.primary_code_no, decQtyID) = False Then
    '                    vtblPool.DefaultView.RowFilter = Nothing
    '                    vtblPool.DefaultView.RowFilter = " [primary code no] = " & vrowIDX.primary_code_no
    '                    If CostingMethod = ClsTransaksiMonthlyClosingItem.EnumCostingMethod.LIFO Then
    '                        vtblPool.DefaultView.Sort = " [tgl transaksi asal] desc "
    '                    Else
    '                        vtblPool.DefaultView.Sort = " [tgl transaksi asal] asc "
    '                    End If
    '                    If vtblPool.DefaultView.Count > 0 Then
    '                        If CostingMethod = ClsTransaksiMonthlyClosingItem.EnumCostingMethod.Average Then
    '                            decPrice = Me.GetNilaiAVG(vtblPool, vrowIDX.primary_code_no)
    '                            vrowID = vtblID.Newtbltrans_purchase_closing_item_detailRow
    '                            vrowID.primary = vrowIDX.primary
    '                            vrowID.primary_transaksi = vrowIDX.primary_transaksi
    '                            vrowID.primary_transaksi_asal = vrowIDX.primary_transaksi_asal
    '                            vrowID.tipe_transaksi_asal = vrowIDX.tipe_transaksi_asal
    '                            vrowID.tgl_transaksi_asal = vrowIDX.tgl_transaksi_asal
    '                            vrowID.primary_code_no = vrowIDX.primary_code_no
    '                            vrowID.tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
    '                            vrowID.qty = decQtyID
    '                            vrowID.unit_price = decPrice
    '                            vrowID.amount = vrowID.qty * vrowID.unit_price
    '                            vrowID.return_price = vrowIDX.return_price
    '                            vrowID.primary_gain_account = vrowIDX.primary_gain_account
    '                            vrowID.primary_loss_account = vrowIDX.primary_loss_account
    '                            vrowID.primary_adjustment_type = vrowIDX.primary_adjustment_type
    '                            vrowID.remarks = vrowIDX.remarks
    '                            vrowID.Overhead = vrowIDX.Overhead
    '                            vrowID.Cost_Perct_Bundling = vrowIDX.Cost_Perct_Bundling
    '                            vtblID.Rows.Add(vrowID)
    '                            vrowIDX.unit_price = vrowID.unit_price
    '                            vrowIDX.amount = vrowID.amount

    '                            While decQtyID > 0
    '                                vtblPool.DefaultView.RowFilter = Nothing
    '                                vtblPool.DefaultView.RowFilter = " [primary code no] = " & vrowIDX.primary_code_no
    '                                If CostingMethod = ClsTransaksiMonthlyClosingItem.EnumCostingMethod.LIFO Then
    '                                    vtblPool.DefaultView.Sort = " [tgl transaksi asal] desc "
    '                                Else
    '                                    vtblPool.DefaultView.Sort = " [tgl transaksi asal] asc "
    '                                End If
    '                                For intA As Integer = 0 To vtblPool.DefaultView.Count - 1
    '                                    vrowP = vtblPool.DefaultView(intA).Row
    '                                    If vrowP.qty > decQtyID Then
    '                                        vrowP.qty -= decQtyID
    '                                        vrowP.unit_price = decPrice
    '                                        vrowP.amount = vrowP.unit_price * vrowP.qty
    '                                        decQtyID = 0
    '                                    Else
    '                                        decQtyID -= vrowP.qty
    '                                        vrowP.qty = 0
    '                                    End If
    '                                    If decQtyID = 0 Then
    '                                        Exit While
    '                                    End If
    '                                Next
    '                                Me.HapustblPool(vtblPool, vPrimaryCodeNo, CostingMethod)
    '                            End While
    '                            Me.HapustblPool(vtblPool, vPrimaryCodeNo, CostingMethod)
    '                        Else
    '                            While decQtyID > 0
    '                                vtblPool.DefaultView.RowFilter = Nothing
    '                                vtblPool.DefaultView.RowFilter = " [primary code no] = " & vrowIDX.primary_code_no
    '                                If CostingMethod = ClsTransaksiMonthlyClosingItem.EnumCostingMethod.LIFO Then
    '                                    vtblPool.DefaultView.Sort = " [tgl transaksi asal] desc "
    '                                Else
    '                                    vtblPool.DefaultView.Sort = " [tgl transaksi asal] asc "
    '                                End If
    '                                For intA As Integer = 0 To vtblPool.DefaultView.Count - 1
    '                                    vrowP = vtblPool.DefaultView(intA).Row
    '                                    If vrowP.qty > decQtyID Then
    '                                        vrowID = vtblID.Newtbltrans_purchase_closing_item_detailRow
    '                                        vrowID.primary = vrowIDX.primary
    '                                        vrowID.primary_transaksi = vrowIDX.primary_transaksi
    '                                        vrowID.primary_transaksi_asal = vrowIDX.primary_transaksi_asal
    '                                        vrowID.tipe_transaksi_asal = vrowIDX.tipe_transaksi_asal
    '                                        vrowID.tgl_transaksi_asal = vrowIDX.tgl_transaksi_asal
    '                                        vrowID.primary_code_no = vrowIDX.primary_code_no
    '                                        vrowID.tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
    '                                        vrowID.qty = decQtyID
    '                                        vrowID.unit_price = vrowP.unit_price
    '                                        vrowID.amount = decQtyID * vrowID.unit_price
    '                                        vrowID.return_price = vrowIDX.return_price
    '                                        vrowID.primary_gain_account = vrowIDX.primary_gain_account
    '                                        vrowID.primary_loss_account = vrowIDX.primary_loss_account
    '                                        vrowID.primary_adjustment_type = vrowIDX.primary_adjustment_type
    '                                        vrowID.remarks = vrowIDX.remarks
    '                                        vrowID.Overhead = vrowIDX.Overhead
    '                                        vrowID.Cost_Perct_Bundling = vrowIDX.Cost_Perct_Bundling
    '                                        vtblID.Rows.Add(vrowID)
    '                                        'kurangi qty pool 
    '                                        vrowP.qty -= decQtyID
    '                                        'decQty habis dipakai
    '                                        decQtyID = 0
    '                                        vrowIDX.unit_price = vrowID.unit_price
    '                                        vrowIDX.amount = vrowID.amount
    '                                        vrowIDX.qty = vrowID.qty
    '                                    Else
    '                                        vrowID = vtblID.Newtbltrans_purchase_closing_item_detailRow
    '                                        vrowID.primary = vrowIDX.primary
    '                                        vrowID.primary_transaksi = vrowIDX.primary_transaksi
    '                                        vrowID.primary_transaksi_asal = vrowIDX.primary_transaksi_asal
    '                                        vrowID.tipe_transaksi_asal = vrowIDX.tipe_transaksi_asal
    '                                        vrowID.tgl_transaksi_asal = vrowIDX.tgl_transaksi_asal
    '                                        vrowID.primary_code_no = vrowIDX.primary_code_no
    '                                        vrowID.tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
    '                                        vrowID.qty = vrowP.qty
    '                                        vrowID.unit_price = vrowP.unit_price
    '                                        vrowID.amount = vrowID.qty * vrowID.unit_price
    '                                        vrowID.return_price = vrowIDX.return_price
    '                                        vrowID.primary_gain_account = vrowIDX.primary_gain_account
    '                                        vrowID.primary_loss_account = vrowIDX.primary_loss_account
    '                                        vrowID.primary_adjustment_type = vrowIDX.primary_adjustment_type
    '                                        vrowID.remarks = vrowIDX.remarks
    '                                        vrowID.Overhead = vrowIDX.Overhead
    '                                        vrowID.Cost_Perct_Bundling = vrowIDX.Cost_Perct_Bundling
    '                                        vrowIDX.unit_price = vrowID.unit_price
    '                                        vrowIDX.amount = vrowID.amount
    '                                        vrowIDX.qty = vrowID.qty
    '                                        vtblID.Rows.Add(vrowID)
    '                                        'decQtyID kurangi qty pool
    '                                        decQtyID -= vrowP.qty
    '                                        'qty pool habis
    '                                        vrowP.qty = 0
    '                                    End If

    '                                    If decQtyID = 0 Then
    '                                        Exit While
    '                                    End If
    '                                Next
    '                                Me.HapustblPool(vtblPool, vPrimaryCodeNo, CostingMethod)
    '                            End While
    '                            Me.HapustblPool(vtblPool, vPrimaryCodeNo, CostingMethod)
    '                        End If
    '                    End If

    '                    vtblB.DefaultView.RowFilter = Nothing
    '                    vtblID.DefaultView.RowFilter = Nothing
    '                    If vrowIDX.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling Then
    '                        Me.UpdatePriceOutBundling(vrowIDX.primary_transaksi_asal, vrowIDX.primary_code_no, vtblB, vtblID)
    '                        If Me.ApakahOutBundlingLengkap(vrowIDX.primary_transaksi_asal, vtblB) = True Then
    '                            Me.calculatePriceInBundling(vrowIDX.primary_transaksi_asal, vtblB, vtblPool, vtblPx, vtblID, vtblIDx, CostingMethod)
    '                        Else
    '                            vrowPX = vtblPx.NewtblItemDetailPoolRow
    '                            vrowPX.Primary = vrowIDX.primary
    '                            vrowPX.Primary_transaksi = vrowIDX.primary_transaksi
    '                            vrowPX.Primary_Transaksi_Asal = vrowIDX.primary_transaksi_asal
    '                            vrowPX.tipe_transaksi_asal = vrowIDX.tipe_transaksi_asal
    '                            vrowPX.tgl_transaksi_asal = vrowIDX.tgl_transaksi_asal
    '                            vrowPX.primary_code_no = vrowIDX.primary_code_no
    '                            vrowPX.tipe_item_detail = vrowIDX.tipe_item_detail
    '                            vrowPX.unit_price = 0
    '                            vrowPX.amount = 0
    '                            vrowPX.qty = vrowIDX.qty
    '                            vrowPX.return_price = vrowIDX.return_price
    '                            vrowPX.primary_gain_account = vrowIDX.primary_gain_account
    '                            vrowPX.primary_loss_account = vrowIDX.primary_loss_account
    '                            vrowPX.primary_adjustment_type = vrowIDX.primary_adjustment_type
    '                            vrowPX.remarks = vrowIDX.remarks
    '                            vrowPX.Cost_Perct_Bundling = vrowIDX.Cost_Perct_Bundling
    '                            vrowPX.Overhead = vrowIDX.Overhead
    '                            vtblPx.Rows.Add(vrowPX)
    '                        End If
    '                    End If
    '                Else
    '                    Exit Sub
    '                End If
    '            Next
    '        Else
    '            Exit Sub
    '        End If

    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    'Private Function ApakahOutBundlingLengkap(ByVal vPrimaryTransaksiAsal As Double, _
    '                                          ByVal vtblB As DseMonthlyClosingNew.BundlingDataTable) As Boolean
    '    Dim vrowB As DseMonthlyClosingNew.BundlingRow

    '    Try
    '        vtblB.DefaultView.RowFilter = Nothing
    '        vtblB.DefaultView.RowFilter = "[Primary Transaksi] = " & vPrimaryTransaksiAsal
    '        If vtblB.DefaultView.Count > 0 Then
    '            For intB As Integer = 0 To vtblB.DefaultView.Count - 1
    '                vrowB = vtblB.DefaultView(intB).Row
    '                If vrowB.Price_B4 = 0 Then
    '                    Return False
    '                End If
    '            Next
    '        End If

    '        Return True

    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Function

    'Private Sub calculatePriceInBundling(ByVal vPrimaryTransaksiAsal As Double, _
    '                                     ByRef vtblB As DseMonthlyClosingNew.BundlingDataTable, _
    '                                     ByRef vtblPool As DseMonthlyClosingNew.tblItemDetailPoolDataTable, _
    '                                     ByRef vtblPx As DseMonthlyClosingNew.tblItemDetailPoolDataTable, _
    '                                     ByRef vtblID As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailDataTable, _
    '                                     ByRef vtblIDX As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailDataTable, ByVal CostingMethod As Integer)
    '    Dim vrowB As DseMonthlyClosingNew.BundlingRow
    '    Dim vrowPool As DseMonthlyClosingNew.tblItemDetailPoolRow
    '    Dim vrowPx As DseMonthlyClosingNew.tblItemDetailPoolRow
    '    Dim vrowID As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailRow
    '    Dim vrowIDx As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailRow
    '    Dim decOverhead As Decimal = 0
    '    Dim vtblB2 As New DseMonthlyClosingNew.BundlingDataTable
    '    Dim vrowB2 As DseMonthlyClosingNew.BundlingRow


    '    Try
    '        vtblB2.Rows.Clear()
    '        vtblB.DefaultView.RowFilter = Nothing
    '        vtblB.DefaultView.RowFilter = "[Primary Transaksi] = " & vPrimaryTransaksiAsal
    '        If vtblB.DefaultView.Count > 0 Then
    '            Dim ListPrimaryCodeAfter As New List(Of Double)
    '            For intB As Integer = 0 To vtblB.DefaultView.Count - 1
    '                vrowB = vtblB.DefaultView(intB).Row
    '                vtblB2.DefaultView.RowFilter = Nothing
    '                vtblB2.DefaultView.RowFilter = " [Primary Code B4] = " & vrowB.Primary_Code_B4
    '                If vtblB2.DefaultView.Count = 0 Then
    '                    vrowB2 = vtblB2.NewBundlingRow
    '                    vrowB2.Primary_Transaksi = vrowB.Primary_Transaksi
    '                    vrowB2.Primary_Code_B4 = vrowB.Primary_Code_B4
    '                    vrowB2.Product_Code_B4 = vrowB.Product_Code_B4
    '                    vrowB2.Qty_B4 = vrowB.Qty_B4
    '                    vrowB2.Price_B4 = vrowB.Price_B4
    '                    vrowB2.Amount_B4 = vrowB.Amount_B4
    '                    vrowB2.Primary_Unit_B4 = vrowB.Primary_Unit_B4
    '                    vrowB2.Primary_Code_After = vrowB.Primary_Code_After
    '                    vrowB2.Product_Code_After = vrowB.Product_Code_After
    '                    vrowB2.qty_After = vrowB.qty_After
    '                    vrowB2.Price_After = vrowB.Price_After
    '                    vrowB2.Amount_After = vrowB.Amount_After
    '                    vrowB2.Primary_Unit_After = vrowB.Primary_Unit_After
    '                    vrowB2.Faktor_pemberat = vrowB.Faktor_pemberat
    '                    vrowB2.Tanggal_Transaksi = vrowB.Tanggal_Transaksi
    '                    vrowB2.No_Transaksi = vrowB.No_Transaksi
    '                    vtblB2.Rows.Add(vrowB2)
    '                End If
    '                If ListPrimaryCodeAfter.Contains(vrowB.Primary_Code_After) = False Then
    '                    ListPrimaryCodeAfter.Add(vrowB.Primary_Code_After)
    '                End If
    '            Next

    '            Dim decTotalX As Decimal = 0
    '            For intP As Integer = 0 To ListPrimaryCodeAfter.Count - 1
    '                vtblB.DefaultView.RowFilter = Nothing
    '                vtblB.DefaultView.RowFilter = "[Primary Transaksi] = " & vPrimaryTransaksiAsal & _
    '                " and [Primary code after] = " & ListPrimaryCodeAfter.Item(intP)
    '                If vtblB.DefaultView.Count > 0 Then
    '                    For intB As Integer = 0 To vtblB.DefaultView.Count - 1
    '                        vrowB = vtblB.DefaultView(intB).Row
    '                        vtblB2.DefaultView.RowFilter = Nothing
    '                        vtblB2.DefaultView.RowFilter = "[Primary Transaksi] = " & vPrimaryTransaksiAsal
    '                        Dim decTotal As Decimal = 0
    '                        If vtblB2.DefaultView.Count > 0 Then
    '                            For Each dvB2 As Data.DataRowView In vtblB2.DefaultView
    '                                vrowB2 = dvB2.Row
    '                                decTotal += vrowB2.Amount_B4
    '                            Next
    '                        End If
    '                        If intP <> ListPrimaryCodeAfter.Count - 1 Then
    '                            vrowB.Amount_After = Me.ObjCD.DigitGrouping(CDec(decTotal * CDec(vrowB.Faktor_pemberat) / 100), Me.DefaultPrecision)
    '                            decTotalX = Me.ObjCD.DigitGrouping(CDec(decTotal * CDec(vrowB.Faktor_pemberat) / 100), Me.DefaultPrecision)
    '                        Else
    '                            vrowB.Amount_After = Me.ObjCD.DigitGrouping(decTotal - decTotalX, Me.DefaultPrecision)
    '                        End If
    '                        ' vrowB.Price_After = vrowB.Amount_After / vrowB.qty_After
    '                    Next
    '                End If
    '            Next
    '        End If


    '        Me.TBLItemDetail.DefaultView.RowFilter = Nothing
    '        Me.TBLItemDetail.DefaultView.RowFilter = " [Primary transaksi asal] = " & vPrimaryTransaksiAsal & _
    '                                                 " And [tipe transaksi asal] = " & ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling
    '        If Me.TBLItemDetail.DefaultView.Count > 0 Then
    '            For intDetail As Integer = 0 To Me.TBLItemDetail.DefaultView.Count - 1
    '                Me.rowItemDetail = Me.TBLItemDetail.DefaultView(intDetail).Row
    '                With Me.rowItemDetail
    '                    Dim AmountAfter As Decimal = 0
    '                    Dim QtyAfter As Decimal = 0
    '                    vtblB.DefaultView.RowFilter = Nothing
    '                    vtblB.DefaultView.RowFilter = "[Primary transaksi] = " & vPrimaryTransaksiAsal & _
    '                                                  " And [Primary Code After] = " & .primary_code_no
    '                    If vtblB.DefaultView.Count > 0 Then
    '                        vrowB = vtblB.DefaultView(0).Row
    '                        AmountAfter = vrowB.Amount_After + .Overhead
    '                        QtyAfter = vrowB.qty_After
    '                    End If
    '                    vrowPool = vtblPool.NewtblItemDetailPoolRow
    '                    vrowPool.Primary = .primary
    '                    vrowPool.Primary_transaksi = .primary_transaksi
    '                    vrowPool.Primary_Transaksi_Asal = .primary_transaksi_asal
    '                    vrowPool.tipe_transaksi_asal = .tipe_transaksi_asal
    '                    vrowPool.tgl_transaksi_asal = .tgl_transaksi_asal
    '                    vrowPool.primary_code_no = .primary_code_no
    '                    vrowPool.tipe_item_detail = .tipe_item_detail
    '                    vrowPool.amount = AmountAfter
    '                    vrowPool.qty = .qty
    '                    vrowPool.unit_price = vrowPool.amount / vrowPool.qty
    '                    vrowPool.return_price = .return_price
    '                    vrowPool.primary_gain_account = .primary_gain_account
    '                    vrowPool.primary_loss_account = .primary_loss_account
    '                    vrowPool.primary_adjustment_type = .primary_adjustment_type
    '                    vrowPool.remarks = .remarks
    '                    vrowPool.Cost_Perct_Bundling = .Cost_Perct_Bundling
    '                    vrowPool.Overhead = .Overhead
    '                    vtblPool.Rows.Add(vrowPool)

    '                    vrowID = vtblID.Newtbltrans_purchase_closing_item_detailRow
    '                    vrowID.primary = .primary
    '                    vrowID.primary_transaksi = .primary_transaksi
    '                    vrowID.primary_transaksi_asal = .primary_transaksi_asal
    '                    vrowID.tipe_transaksi_asal = .tipe_transaksi_asal
    '                    vrowID.tgl_transaksi_asal = .tgl_transaksi_asal
    '                    vrowID.primary_code_no = .primary_code_no
    '                    vrowID.tipe_item_detail = .tipe_item_detail
    '                    vrowID.unit_price = .unit_price
    '                    vrowID.amount = AmountAfter + .Overhead
    '                    vrowID.qty = .qty
    '                    vrowID.unit_price = vrowID.amount / vrowID.qty
    '                    vrowID.return_price = .return_price
    '                    vrowID.primary_gain_account = .primary_gain_account
    '                    vrowID.primary_loss_account = .primary_loss_account
    '                    vrowID.primary_adjustment_type = .primary_adjustment_type
    '                    vrowID.remarks = .remarks
    '                    vrowID.Cost_Perct_Bundling = .Cost_Perct_Bundling
    '                    vrowID.Overhead = .Overhead
    '                    vtblID.Rows.Add(vrowID)

    '                    vrowIDx = vtblIDX.Newtbltrans_purchase_closing_item_detailRow
    '                    vrowIDx.primary = .primary
    '                    vrowIDx.primary_transaksi = .primary_transaksi
    '                    vrowIDx.primary_transaksi_asal = .primary_transaksi_asal
    '                    vrowIDx.tipe_transaksi_asal = .tipe_transaksi_asal
    '                    vrowIDx.tgl_transaksi_asal = .tgl_transaksi_asal
    '                    vrowIDx.primary_code_no = .primary_code_no
    '                    vrowIDx.tipe_item_detail = .tipe_item_detail
    '                    vrowIDx.amount = AmountAfter + .Overhead
    '                    vrowIDx.qty = .qty
    '                    vrowIDx.unit_price = vrowIDx.amount / vrowIDx.qty
    '                    vrowIDx.return_price = .return_price
    '                    vrowIDx.primary_gain_account = .primary_gain_account
    '                    vrowIDx.primary_loss_account = .primary_loss_account
    '                    vrowIDx.primary_adjustment_type = .primary_adjustment_type
    '                    vrowIDx.remarks = .remarks
    '                    vrowIDx.Cost_Perct_Bundling = .Cost_Perct_Bundling
    '                    vrowIDx.Overhead = .Overhead
    '                    vtblIDX.Rows.Add(vrowIDx)
    '                    '------------------------------------------------------------------------------------------------
    '                    '-----------------> cek dulu apakah ada price = 0
    '                    '------------------------------------------------------------------------------------------------
    '                    vtblIDX.DefaultView.RowFilter = Nothing
    '                    vtblIDX.DefaultView.RowFilter = " [Primary Code No] = " & .primary_code_no & _
    '                                                  " And [tipe item detail] = " & ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem & _
    '                                                  " And [unit price] = 0 "
    '                    If vtblIDX.DefaultView.Count > 0 Then
    '                        'cek terlebih dahulu apakah masih ada pencatatan belum selesai di transaksi sebelumnya
    '                        ' hitung previous data yg belum ada nilainya
    '                        Me.CalculatePrevDataNol(vtblPool, vtblPx, vtblIDX, vtblID, vtblB, .primary_code_no, CostingMethod)
    '                    End If
    '                End With
    '            Next
    '        End If

    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Private Sub HapustblPool(ByRef vtblPool As DseMonthlyClosingNew.tblItemDetailPoolDataTable, ByVal vPrimaryCodeNo As Double, ByVal vcostingMethod As Integer)
        Dim vrowPool As DseMonthlyClosingNew.tblItemDetailPoolRow
        Try
            vtblPool.DefaultView.RowFilter = Nothing
            vtblPool.DefaultView.RowFilter = " [primary code no] = " & vPrimaryCodeNo
            If vcostingMethod = ClsTransaksiMonthlyClosingItem.EnumCostingMethod.LIFO Then
                vtblPool.DefaultView.Sort = " [tgl transaksi asal] desc "
            Else
                vtblPool.DefaultView.Sort = " [tgl transaksi asal] asc "
            End If
            If vtblPool.Rows.Count > 0 Then
                Dim IntA As Integer = 0
                IntA = vtblPool.Rows.Count - 1
                While IntA >= 0
                    vrowPool = vtblPool.Rows(IntA)
                    If vrowPool.qty = 0 Then
                        vtblPool.Rows.RemoveAt(IntA)
                    End If
                    IntA -= 1
                End While
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Function GetNilaiAVG(ByVal vtblPool As DseMonthlyClosingNew.tblItemDetailPoolDataTable, ByVal vPrimaryCodeNo As Double) As Decimal
        Dim rowPool As DseMonthlyClosingNew.tblItemDetailPoolRow
        Try
            Dim DecAmount As Decimal = 0
            Dim decQty As Decimal = 0

            vtblPool.DefaultView.RowFilter = Nothing
            vtblPool.DefaultView.RowFilter = "[Primary Code No] = " & vPrimaryCodeNo
            vtblPool.DefaultView.Sort = "[tgl transaksi asal] "
            If vtblPool.DefaultView.Count > 0 Then
                For Each dv As Data.DataRowView In vtblPool.DefaultView
                    rowPool = dv.Row
                    If IsNumeric(rowPool.qty) Then
                        decQty += rowPool.qty
                    End If

                    If IsNumeric(rowPool.amount) Then
                        DecAmount += rowPool.amount
                    End If
                Next
            End If


            Dim decPrice As Decimal = 0
            If decQty > 0 Then
                decPrice = DecAmount / decQty
            End If

            Return decPrice

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function ApakahAkanMinus(ByVal vtblP As DseMonthlyClosingNew.tblItemDetailPoolDataTable, ByVal vPrimaryCodeNo As Double, ByVal decQtyOut As Decimal) As Boolean
        Dim vrowP As DseMonthlyClosingNew.tblItemDetailPoolRow

        Try
            vtblP.DefaultView.RowFilter = Nothing
            vtblP.DefaultView.RowFilter = "[Primary Code No] = " & vPrimaryCodeNo
            vtblP.DefaultView.Sort = "[tgl transaksi asal] "
            Dim decQtyIn As Decimal = 0
            If vtblP.DefaultView.Count > 0 Then
                For Each dvP As Data.DataRowView In vtblP.DefaultView
                    vrowP = dvP.Row
                    If IsNumeric(vrowP.qty) Then
                        decQtyIn += vrowP.qty
                    End If
                Next
            End If


            If decQtyIn - decQtyOut < 0 Then
                Return True
            Else
                Return False
            End If


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function


    Private Function GetCostPerUnit_Adjustment(ByVal vTipeItem As ClsTransaksiMonthlyClosingItem.EnumTipeDokumenAcuan, ByVal vPrimaryDocAcuan As Double) As Decimal
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvItemAdjust.Rows
                If CInt(rowDGV.Cells(Me.ColJenisDokumen.Index).Value) = CInt(vTipeItem) And _
                rowDGV.Cells(Me.ColPrimaryStockInAdjust.Index).Value = vPrimaryDocAcuan Then
                    Return rowDGV.Cells(Me.ColUnitPriceAdjust.Index).Value
                End If
            Next

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function ApakahAdaItemAdjustment(ByVal vTipeItem As ClsTransaksiMonthlyClosingItem.EnumTipeDokumenAcuan, ByVal vPrimaryDocAcuan As Double) As Boolean
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvItemAdjust.Rows
                If CInt(rowDGV.Cells(Me.ColJenisDokumen.Index).Value) = CInt(vTipeItem) And _
                rowDGV.Cells(Me.ColPrimaryStockInAdjust.Index).Value = vPrimaryDocAcuan Then
                    Return True
                End If
            Next

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    '''update data dari item detail ke dgv utama
    'Private Sub UpdateKeDGVItemUtama(Optional ByVal vPrimaryProduct As Double = -1)
    '    Dim rowDGV As DataGridViewRow
    '    Dim Filter As String = ""
    '    Dim TipeItem As ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.BeginingBalanceItem
    '    Dim Amount As Decimal = 0
    '    Dim vDefPrec As Integer = 0
    '    Dim DummyBool As DataGridViewAutoSizeColumnMode = DataGridViewAutoSizeColumnMode.None
    '    Dim vCek As Boolean = False

    '    Try
    '        DummyBool = Me.dgvItem.AutoSizeColumnsMode
    '        If vPrimaryProduct = -1 Then
    '            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
    '            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
    '        End If

    '        vDefPrec = Me.objCurrency.GetDecimal_DefaultCurrency
    '        For Each rowDGV In Me.dgvItem.Rows
    '            With rowDGV
    '                vCek = False
    '                If vPrimaryProduct = -1 Then
    '                    vCek = True
    '                Else
    '                    If rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct Then
    '                        vCek = True
    '                    Else
    '                        vCek = False
    '                    End If
    '                End If
    '                If vCek = True Then
    '                    'begining balance
    '                    Amount = 0
    '                    TipeItem = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.BeginingBalanceItem
    '                    Filter = "[primary code no] = " + rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value.ToString + _
    '                    " AND [tipe item detail] = " + CInt(TipeItem).ToString
    '                    Me.TBLItemDetail.DefaultView.RowFilter = Filter
    '                    For Each dv As DataRowView In Me.TBLItemDetail.DefaultView
    '                        Me.rowItemDetail = dv.Row
    '                        Amount += Me.rowItemDetail.amount
    '                    Next
    '                    rowDGV.Cells(Me.ColBeginAmount.Index).Value = Me.ObjCD.DigitGrouping(Amount, vDefPrec)

    '                    'in
    '                    Amount = 0
    '                    TipeItem = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
    '                    Filter = "[primary code no] = " + rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value.ToString + _
    '                    " AND [tipe item detail] = " + CInt(TipeItem).ToString
    '                    Me.TBLItemDetail.DefaultView.RowFilter = Filter
    '                    For Each dv As DataRowView In Me.TBLItemDetail.DefaultView
    '                        Me.rowItemDetail = dv.Row
    '                        Amount += Me.rowItemDetail.amount
    '                    Next
    '                    rowDGV.Cells(Me.ColInAmount.Index).Value = Me.ObjCD.DigitGrouping(Amount, vDefPrec)

    '                    'out
    '                    Amount = 0
    '                    TipeItem = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
    '                    Filter = "[primary code no] = " + rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value.ToString + _
    '                    " AND [tipe item detail] = " + CInt(TipeItem).ToString
    '                    Me.TBLItemDetail.DefaultView.RowFilter = Filter
    '                    For Each dv As DataRowView In Me.TBLItemDetail.DefaultView
    '                        Me.rowItemDetail = dv.Row
    '                        Amount += Me.rowItemDetail.amount
    '                    Next
    '                    rowDGV.Cells(Me.ColOutAmount.Index).Value = Me.ObjCD.DigitGrouping(Amount, vDefPrec)

    '                    'hitung end balance
    '                    Amount = CDec(rowDGV.Cells(Me.ColBeginAmount.Index).Value) + _
    '                    CDec(rowDGV.Cells(Me.ColInAmount.Index).Value) - CDec(rowDGV.Cells(Me.ColOutAmount.Index).Value)
    '                    rowDGV.Cells(Me.ColEndAmount.Index).Value = Me.ObjCD.DigitGrouping(Amount, vDefPrec)

    '                    If vPrimaryProduct <> -1 Then
    '                        Exit Sub
    '                    End If
    '                End If
    '            End With
    '        Next

    '        Me.dgvItem.AutoSizeColumnsMode = DummyBool
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    'sementara
    'Private Sub BuangItemDGV_YangTidakTermasukCategory()
    '    Try

    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Private Sub UpdateTBLProduct()
        Dim rowDGV As DataGridViewRow

        Try
            Me.TBLProduct.Rows.Clear()

            For Each rowDGV In Me.dgvItem.Rows
                Me.rowProduct = Me.TBLProduct.Newtblmaster_productRow

                Me.rowProduct.Description = rowDGV.Cells(Me.ColDescription.Index).Value
                Me.rowProduct.Primary = rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value
                Me.rowProduct.Primary_Adjustment_Account = rowDGV.Cells(Me.ColItem_PrimaryAdjustmentAccount.Index).Value
                Me.rowProduct.Primary_COGS_Account = rowDGV.Cells(Me.ColItem_PrimaryCOGSAccount.Index).Value
                Me.rowProduct.Primary_Inventory_Account = rowDGV.Cells(Me.ColItem_PrimaryInventoryAccount.Index).Value
                Me.rowProduct.Primary_Sales_Account = rowDGV.Cells(Me.ColItem_PrimarySalesAccount.Index).Value
                Me.rowProduct.Primary_Sales_Return_Account = rowDGV.Cells(Me.ColItem_PrimarySalesReturnAccount.Index).Value
                Me.rowProduct.Primary_Stock_Taking_Account = rowDGV.Cells(Me.ColItem_PrimaryStockTakingAccount.Index).Value
                Me.rowProduct.Primary_Unit = rowDGV.Cells(Me.ColPrimaryUnits.Index).Value
                Me.rowProduct.Product_Code = rowDGV.Cells(Me.ColCodeNo.Index).Value
                Me.rowProduct.Unit_Code = rowDGV.Cells(Me.ColUnits.Index).Value
                Me.rowProduct._Decimal = rowDGV.Cells(Me.ColItem_Decimal.Index).Value

                Me.TBLProduct.Rows.Add(Me.rowProduct)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Private Sub UpdateAkunDGV()
    '    Dim rowDGV As DataGridViewRow

    '    Try
    '        For Each rowDGV In Me.dgvItem.Rows
    '            If IsNumeric(rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value) AndAlso rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value > -1 Then
    '                Me.TBLProduct.DefaultView.RowFilter = Nothing
    '                Me.TBLProduct.DefaultView.RowFilter = "[Primary] = " + rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value.ToString

    '                Me.rowProduct = Me.TBLProduct.DefaultView(0).Row
    '                With rowDGV
    '                    .Cells(Me.ColItem_Decimal.Index).Value = Me.rowProduct._Decimal
    '                    .Cells(Me.ColItem_PrimaryAdjustmentAccount.Index).Value = Me.rowProduct.Primary_Adjustment_Account
    '                    .Cells(Me.ColItem_PrimaryCOGSAccount.Index).Value = Me.rowProduct.Primary_COGS_Account
    '                    .Cells(Me.ColItem_PrimaryInventoryAccount.Index).Value = Me.rowProduct.Primary_Inventory_Account
    '                    .Cells(Me.ColItem_PrimarySalesAccount.Index).Value = Me.rowProduct.Primary_Sales_Account
    '                    .Cells(Me.ColItem_PrimarySalesReturnAccount.Index).Value = Me.rowProduct.Primary_Sales_Return_Account
    '                    .Cells(Me.ColItem_PrimaryStockTakingAccount.Index).Value = Me.rowProduct.Primary_Stock_Taking_Account
    '                End With
    '            End If
    '        Next
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Private Sub BuatJurnal(Optional ByVal vCekUnbalance As Boolean = True)
        Dim DefPrec As Integer = 0
        Dim rowDGV As DataGridViewRow
        Dim AmountDebet, AmountCredit As Decimal
        Dim StartDate, EndDate As Date

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'kumpulkan data setting
            DTASetting.Fill(Me.TBLSetting)
            Me.TBLSettingItem.Rows.Clear()
            For Each rowSetting In Me.TBLSetting.Rows
                Me.rowSettingItem = Me.TBLSettingItem.Newtbltrans_purchase_closing_item_tipe_adjustmentRow
                With Me.rowSettingItem
                    .Primary = Me.TBLSettingItem.Rows.Count + 1
                    .Primary_Transaksi = Me.TxtPrimaryClosing.Text
                    .PrimaryGain = rowSetting.PrimaryGain
                    .PrimaryLoss = rowSetting.PrimaryLoss
                    .Primary_Adjustment = rowSetting.Primary
                End With
                TBLSettingItem.Rows.Add(Me.rowSettingItem)
            Next

            DefPrec = Me.objCurrency.GetDecimal_DefaultCurrency
            StartDate = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value)
            EndDate = Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value)

            'kosongkan dgv jurnal
            Me.dgvJurnal.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvJurnal.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.dgvJurnal.Rows.Clear()

            'buat jurnal bundling
            If Me.ApakahAdaBundling = False Then
                Me.Calc_CurrentStep += 1
                Me.LblStep.Text = "Step " + Me.Calc_CurrentStep.ToString + " of " + Me.Calc_TotalStep.ToString + " : creating bundling journal"
                Me.ProgressBar1.Value = 0
                Me.LblProgress.Text = ""
                Me.ProgressBar1.Maximum = Me.dgvItem.RowCount
                'Me.LblProgress.Visible = True
            Else
                Me.BuatJurnalBundling(DefPrec)
            End If

            'buat jurnal overhead
            Me.BuatJurnalOverhead(DefPrec)

            'buat jurnal sales
            Me.BuatJurnalCOGS(DefPrec, StartDate, EndDate)

            'buat jurnal in out adjust
            Me.BuatJurnalAdjustment1(DefPrec, StartDate, EndDate)

            'buat jurnal in out stock take adjust
            If Me.ObjTaking.GetJml(StartDate, EndDate) > 0 Then
                Me.BuatJurnalStockTakeAdjustment(DefPrec)
            Else
                Me.Calc_CurrentStep += 1
                Me.LblStep.Text = "Step " + Me.Calc_CurrentStep.ToString + " of " + Me.Calc_TotalStep.ToString + " : creating stock taking journal"
            End If

            'buat jurnal sales return
            Me.BuatJurnalSalesReturn(DefPrec, StartDate, EndDate)

            'buat jurnal purchase return 
            Me.BuatJurnalPR(DefPrec)

            'set autosize jurnal
            Me.dgvJurnal.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            Me.dgvJurnal.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            'cek klo un balance
            If vCekUnbalance = True Then
                For Each rowDGV In Me.dgvJurnal.Rows
                    AmountDebet += CDec(rowDGV.Cells(Me.ColJournal_Debet.Index).Value)
                    AmountCredit += CDec(rowDGV.Cells(Me.ColJournal_Kredit.Index).Value)
                Next
                If AmountDebet <> AmountCredit Then
                    MsgBox("Unbalance amount", MsgBoxStyle.Information, "Information")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub BuatJurnalOverhead(ByVal vDefPrec As Integer)
        Dim Dummy As Boolean = False
        Dim rowDGV As DataGridViewRow
        Dim rowAccount As DseTransaksiMonthlyClosing_Overhead.tbltrans_purchase_closing_item_cost_group_accountRow
        Dim SaldoNormal As Double
        Dim Amount1 As Decimal = 0
        Dim IsDebet As Boolean

        Try
            Me.Calc_CurrentStep += 1
            Me.LblStep.Text = "Step " + Me.Calc_CurrentStep.ToString + " of " + Me.Calc_TotalStep.ToString + " : creating overhead cost journal"
            Me.ProgressBar1.Value = 0
            Me.LblProgress.Text = ""
            'Me.LblProgress.Visible = True
            Me.ProgressBar1.Maximum = Me.dgvItemCost_CostGroup.RowCount

            Dummy = Me.dgvItemCost_CostGroup.AllowUserToAddRows
            Me.dgvItemCost_CostGroup.AllowUserToAddRows = False

            'untuk setiap baris cost group,
            For Each rowDGV In Me.dgvItemCost_CostGroup.Rows
                Me.ProgressBar1.Value += 1
                If Me.ProgressBar1.Value > 25 Then
                    If Math.IEEERemainder(Me.ProgressBar1.Value, 100) = 0 Then
                        Me.LblProgress.Text = "Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString + " : " + rowDGV.Cells(Me.ColCostGroup_GroupDesc.Index).Value.ToString
                    ElseIf Me.ProgressBar1.Value = Me.ProgressBar1.Maximum Then
                        Me.LblProgress.Text = "Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString + " : " + rowDGV.Cells(Me.ColCostGroup_GroupDesc.Index).Value.ToString
                    End If
                Else
                    Me.LblProgress.Text = "Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString + " : " + rowDGV.Cells(Me.ColCostGroup_GroupDesc.Index).Value.ToString
                End If

                '1. cari apakah ada product yang di override
                Me.vTBLCostGroup_CodeNo.DefaultView.RowFilter = "identifier = '" + rowDGV.Cells(Me.ColCostGroup_Identifier.Index).Value.ToString + "'"
                If Me.vTBLCostGroup_CodeNo.DefaultView.Count > 0 Then
                    '2. jika ada, maka setiap akun pada cost group tersebut, dibuatkan jurnalnya sejumlah allocated amount.
                    Me.vTBLCostGroup_Account.DefaultView.RowFilter = "identifier = '" + rowDGV.Cells(Me.ColCostGroup_Identifier.Index).Value.ToString + "'"
                    For Each DV As DataRowView In Me.vTBLCostGroup_Account.DefaultView
                        rowAccount = DV.Row
                        If rowAccount.allocated_amount <> 0 Then
                            If rowAccount.allocated_amount > 0 Then
                                'a. jika allocated amount +, jurnal pada kebalikan saldo normal
                                SaldoNormal = Me.ObjAkun.GetSaldoNormal(rowAccount.primary_akun)

                                If SaldoNormal = 0 Then
                                    'jurnal pada kredit
                                    IsDebet = False
                                Else
                                    'jurnal pada debet
                                    IsDebet = True
                                End If
                            ElseIf rowAccount.allocated_amount < 0 Then
                                'b. jika allocated amount -, jurnal pada saldo normal
                                SaldoNormal = Me.ObjAkun.GetSaldoNormal(rowAccount.primary_akun)
                                If SaldoNormal = 0 Then
                                    'jurnal pada debet
                                    IsDebet = True
                                Else
                                    'jurnal pada kredit
                                    IsDebet = False
                                End If
                            End If

                            'jurnal
                            If Me.GetIndexJurnal(rowAccount.primary_akun, IsDebet, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OverheadCost, -1) = -1 Then
                                'tambahkan baris baru
                                Me.BuatRowJurnal(rowAccount.primary_akun, IsDebet, Math.Abs(rowAccount.allocated_amount), vDefPrec, _
                                ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OverheadCost, -1)
                            Else
                                Amount1 = Me.GetAmountRowJurnal(rowAccount.primary_akun, IsDebet, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OverheadCost, -1)
                                Me.SetAmountRowJurnal(rowAccount.primary_akun, IsDebet, Math.Abs(rowAccount.allocated_amount) + Amount1, vDefPrec, _
                                ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OverheadCost, -1)
                            End If
                        End If
                    Next
                End If
            Next

            Me.dgvItemCost_CostGroup.AllowUserToAddRows = Dummy
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BuatRowJurnal(ByVal vPrimaryAkun As Double, ByVal vDebet As Boolean, ByVal vAmount As Decimal, ByVal vDefPrec As Decimal, _
    ByVal vTipeTransaksiAsal As ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal, _
    ByVal vPrimaryAdjustmentReference As Double)
        Dim rowDGV As DataGridViewRow

        Try
            Me.dgvJurnal.Rows.Add()
            rowDGV = Me.dgvJurnal.Rows(Me.dgvJurnal.Rows.Count - 1)

            With rowDGV
                .Cells(Me.ColJournal_Account.Index).Value = Me.ObjAkun.GetAccountCode(vPrimaryAkun)

                If vDebet = True Then
                    .Cells(Me.ColJournal_Credit_OriginalCurrency.Index).Value = Me.ObjCD.DigitGrouping(0, vDefPrec)
                    .Cells(Me.ColJournal_Kredit.Index).Value = Me.ObjCD.DigitGrouping(0, vDefPrec)
                    .Cells(Me.ColJournal_Debet.Index).Value = Me.ObjCD.DigitGrouping(vAmount, vDefPrec)
                    .Cells(Me.ColJournal_Debet_OriginalCurr.Index).Value = Me.ObjCD.DigitGrouping(vAmount, vDefPrec)
                Else
                    .Cells(Me.ColJournal_Debet.Index).Value = Me.ObjCD.DigitGrouping(0, vDefPrec)
                    .Cells(Me.ColJournal_Debet_OriginalCurr.Index).Value = Me.ObjCD.DigitGrouping(0, vDefPrec)
                    .Cells(Me.ColJournal_Credit_OriginalCurrency.Index).Value = Me.ObjCD.DigitGrouping(vAmount, vDefPrec)
                    .Cells(Me.ColJournal_Kredit.Index).Value = Me.ObjCD.DigitGrouping(vAmount, vDefPrec)
                End If
                .Cells(Me.ColJournal_ExchangeRate.Index).Value = 1

                .Cells(Me.ColJournal_Dept.Index).Value = ""
                .Cells(Me.ColJournal_Job.Index).Value = ""
                .Cells(Me.ColJournal_DescAccount.Index).Value = Me.ObjAkun.GetDescription_ByPrimaryAcc(vPrimaryAkun)
                .Cells(Me.ColJournal_OriginalCurrency.Index).Value = Me.objCurrency.GetCurrencyCode(Core_Module.PrimaryBaseCurrency)
                .Cells(Me.ColJournal_PrimaryAkun.Index).Value = vPrimaryAkun
                .Cells(Me.ColJournal_PrimaryCurrency.Index).Value = Core_Module.PrimaryBaseCurrency
                .Cells(Me.ColJournal_PrimaryDept.Index).Value = -1
                .Cells(Me.ColJournal_PrimaryJob.Index).Value = -1
                If vTipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust Then
                    .Cells(Me.ColJournal_Remarks.Index).Value = "In Adjust Cost"
                ElseIf vTipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling Then
                    .Cells(Me.ColJournal_Remarks.Index).Value = "In Bundling Cost"
                ElseIf vTipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InPurchase Then
                    .Cells(Me.ColJournal_Remarks.Index).Value = "In Purchase Cost"
                ElseIf vTipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR Then
                    .Cells(Me.ColJournal_Remarks.Index).Value = "In Sales Return Cost"
                ElseIf vTipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InStockTakeAdjust Then
                    .Cells(Me.ColJournal_Remarks.Index).Value = "In Stock Take Cost"
                ElseIf vTipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutAdjust Then
                    .Cells(Me.ColJournal_Remarks.Index).Value = "Out Adjust Cost"
                ElseIf vTipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling Then
                    .Cells(Me.ColJournal_Remarks.Index).Value = "Out Bundling Cost"
                ElseIf vTipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutPR Then
                    .Cells(Me.ColJournal_Remarks.Index).Value = "Out Purchase Return Cost"
                ElseIf vTipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutSales Then
                    .Cells(Me.ColJournal_Remarks.Index).Value = "Out Sales Cost"
                ElseIf vTipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutStockTakeAdjust Then
                    .Cells(Me.ColJournal_Remarks.Index).Value = "Out Stock Take Adjust Cost"
                ElseIf vTipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OverheadCost Then
                    .Cells(Me.ColJournal_Remarks.Index).Value = "Overhead Cost"
                End If

                .Cells(Me.ColJournal_TipeTransaksiAsalInt.Index).Value = vTipeTransaksiAsal
                .Cells(Me.ColJournal_TipeTransaksiAsalStr.Index).Value = Me.ObjCLosingItem.GetReference(vPrimaryAdjustmentReference, vTipeTransaksiAsal, Me.TBLSetting)
                .Cells(Me.ColJournal_PrimaryAdjustmentType.Index).Value = vPrimaryAdjustmentReference
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetMulticurrency()
        Dim rowDGV As DataGridViewRow
        Dim Multicurrency As Boolean = False

        Try
            For Each rowDGV In Me.dgvItem.Rows
                If rowDGV.Index = Me.dgvItem.RowCount - 1 Then
                    If Me.dgvItem.AllowUserToAddRows = True Then
                        'tidak dicek
                    Else
                        'dicek
                        If rowDGV.Cells(Me.ColJournal_PrimaryCurrency.Index).Value <> Core_Module.PrimaryBaseCurrency Then
                            Multicurrency = True
                            Exit For
                        End If
                    End If
                Else
                    If rowDGV.Cells(Me.ColJournal_PrimaryCurrency.Index).Value <> Core_Module.PrimaryBaseCurrency Then
                        Multicurrency = True
                        Exit For
                    End If
                End If
            Next

            If Multicurrency = True Then
                Me.ColJournal_ExchangeRate.Visible = True
                Me.ColJournal_Debet_OriginalCurr.Visible = True
                Me.ColJournal_Credit_OriginalCurrency.Visible = True
                Me.ColJournal_OriginalCurrency.Visible = True
            Else
                Me.ColJournal_OriginalCurrency.Visible = False
                Me.ColJournal_ExchangeRate.Visible = False
                Me.ColJournal_Debet_OriginalCurr.Visible = False
                Me.ColJournal_Credit_OriginalCurrency.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetAmountRowJurnal(ByVal vPrimaryAkun As Double, ByVal vDebet As Boolean, ByVal vAmount As Decimal, ByVal vDefPrec As Decimal, _
    ByVal vTipeTransaksiAsal As ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal, _
    ByVal vPrimaryAdjustmentReference As Double)
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvJurnal.Rows
                If rowDGV.Cells(Me.ColJournal_PrimaryAkun.Index).Value = vPrimaryAkun And _
                rowDGV.Cells(Me.ColJournal_PrimaryAdjustmentType.Index).Value = vPrimaryAdjustmentReference And _
                rowDGV.Cells(Me.ColJournal_TipeTransaksiAsalInt.Index).Value = vTipeTransaksiAsal Then
                    If (vDebet = True And CDec(rowDGV.Cells(Me.ColJournal_Debet.Index).Value > 0)) Or _
                    (vDebet = False And CDec(rowDGV.Cells(Me.ColJournal_Kredit.Index).Value > 0)) Then
                        With rowDGV
                            If vDebet = True Then
                                .Cells(Me.ColJournal_Credit_OriginalCurrency.Index).Value = Me.ObjCD.DigitGrouping(0, vDefPrec)
                                .Cells(Me.ColJournal_Kredit.Index).Value = Me.ObjCD.DigitGrouping(0, vDefPrec)
                                .Cells(Me.ColJournal_Debet.Index).Value = Me.ObjCD.DigitGrouping(vAmount, vDefPrec)
                                .Cells(Me.ColJournal_Debet_OriginalCurr.Index).Value = Me.ObjCD.DigitGrouping(vAmount, vDefPrec)
                            Else
                                .Cells(Me.ColJournal_Debet.Index).Value = Me.ObjCD.DigitGrouping(0, vDefPrec)
                                .Cells(Me.ColJournal_Debet_OriginalCurr.Index).Value = Me.ObjCD.DigitGrouping(0, vDefPrec)
                                .Cells(Me.ColJournal_Credit_OriginalCurrency.Index).Value = Me.ObjCD.DigitGrouping(vAmount, vDefPrec)
                                .Cells(Me.ColJournal_Kredit.Index).Value = Me.ObjCD.DigitGrouping(vAmount, vDefPrec)
                            End If
                        End With
                    End If
                End If
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetAmountRowJurnal(ByVal vPrimaryAkun As Double, ByVal vDebet As Boolean, ByVal vTipeTransaksiAsal As ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal, _
    ByVal vPrimaryAdjustmentReference As Double) As Decimal
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvJurnal.Rows
                If rowDGV.Cells(Me.ColJournal_PrimaryAkun.Index).Value = vPrimaryAkun And _
                rowDGV.Cells(Me.ColJournal_PrimaryAdjustmentType.Index).Value = vPrimaryAdjustmentReference And _
                rowDGV.Cells(Me.ColJournal_TipeTransaksiAsalInt.Index).Value = vTipeTransaksiAsal Then
                    If (vDebet = True And CDec(rowDGV.Cells(Me.ColJournal_Debet.Index).Value > 0)) Or _
                    (vDebet = False And CDec(rowDGV.Cells(Me.ColJournal_Kredit.Index).Value > 0)) Then
                        With rowDGV
                            If vDebet = True Then
                                Return .Cells(Me.ColJournal_Debet.Index).Value
                            Else
                                Return .Cells(Me.ColJournal_Kredit.Index).Value
                            End If
                        End With
                    End If
                End If
            Next

            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetIndexJurnal(ByVal vPrimaryAkun As Double, ByVal vDebet As Boolean, ByVal vTipeTransaksiAsal As ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal, _
    ByVal vPrimaryAdjustmentReference As Double) As Double
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvJurnal.Rows
                If rowDGV.Cells(Me.ColJournal_PrimaryAkun.Index).Value = vPrimaryAkun And _
                rowDGV.Cells(Me.ColJournal_PrimaryAdjustmentType.Index).Value = vPrimaryAdjustmentReference And _
                rowDGV.Cells(Me.ColJournal_TipeTransaksiAsalInt.Index).Value = vTipeTransaksiAsal Then
                    If (vDebet = True And CDec(rowDGV.Cells(Me.ColJournal_Debet.Index).Value > 0)) Or _
                    (vDebet = False And CDec(rowDGV.Cells(Me.ColJournal_Kredit.Index).Value > 0)) Then
                        With rowDGV
                            If vDebet = True Then
                                Return rowDGV.Index
                            Else
                                Return rowDGV.Index
                            End If
                        End With
                    End If
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub BuatJurnalBundling(ByVal vDefPrec As Integer) 
        Dim Amount1 As Decimal = 0
        Dim PrimaryAkunInv As Double = -1
        Dim AmountIn As Decimal = 0
        Dim AmountOut As Decimal = 0
        Dim DTA_B As New DseTransaksiMonthlyClosing3TableAdapters.DataTable1TableAdapter
        Dim TBL_B As New DseTransaksiMonthlyClosing3.DataTable1DataTable
        Dim row_B As DseTransaksiMonthlyClosing3.DataTable1Row

        Try
            DTA_B.Fill(TBL_B, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value), Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value))
            Me.Calc_CurrentStep += 1
            Me.LblStep.Text = "Step " + Me.Calc_CurrentStep.ToString + " of " + Me.Calc_TotalStep.ToString + " : creating bundling journal"
            Me.ProgressBar1.Value = 0
            Me.LblProgress.Text = ""
            Me.ProgressBar1.Maximum = TBL_B.Rows.Count

            For Each row_B In TBL_B.Rows
                Me.ProgressBar1.Value += 1

                If Me.ProgressBar1.Value > 25 Then
                    If Math.IEEERemainder(Me.ProgressBar1.Value, 100) = 0 Then
                        Me.LblProgress.Text = "Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString + " : " + row_B.Product_Code
                    ElseIf Me.ProgressBar1.Value = Me.ProgressBar1.Maximum Then
                        Me.LblProgress.Text = "Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString + " : " + row_B.Product_Code
                    End If
                Else
                    Me.LblProgress.Text = "Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString + " : " + row_B.Product_Code
                End If

                '1. untuk setiap item after, masukkan nilai inventory pada debet
                If IsNumeric(row_B.Primary_Product) AndAlso CDbl(row_B.Primary_Product) > -1 Then
                    Me.GetAmount(Me.TBLItemDetail, row_B.Primary_Product, AmountIn, AmountOut)
                    If AmountIn <> 0 Then
                        PrimaryAkunInv = row_B.Primary_Inventory_Account
                        If Me.GetIndexJurnal(PrimaryAkunInv, True, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling, -1) = -1 Then
                            'tambahkan baris baru
                            Me.BuatRowJurnal(PrimaryAkunInv, True, AmountIn, vDefPrec, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling, -1)
                        Else
                            Amount1 = Me.GetAmountRowJurnal(PrimaryAkunInv, True, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling, -1)
                            Me.SetAmountRowJurnal(PrimaryAkunInv, True, AmountIn + Amount1, vDefPrec, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling, -1)
                        End If
                    End If

                    '2. untuk setiap item before, masukkan nilai inventory pada credit
                    If AmountOut <> 0 Then
                        PrimaryAkunInv = row_B.Primary_Inventory_Account

                        If Me.GetIndexJurnal(PrimaryAkunInv, False, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling, -1) = -1 Then
                            'tambahkan baris baru
                            Me.BuatRowJurnal(PrimaryAkunInv, False, AmountOut, vDefPrec, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling, -1)
                        Else
                            Amount1 = Me.GetAmountRowJurnal(PrimaryAkunInv, False, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling, -1)
                            Me.SetAmountRowJurnal(PrimaryAkunInv, False, AmountOut + Amount1, vDefPrec, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling, -1)
                        End If
                    End If
                End If
            Next

            TBL_B.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TBL_B.Dispose()
            DTA_B.Dispose()
        End Try
    End Sub

    Private Sub BuatJurnalStockTakeAdjustment(ByVal vDefPrec As Integer)
        Dim Filter As String = ""
        Dim rowDGV As DataGridViewRow
        Dim DV As DataRowView
        Dim Amount As Decimal = 0
        Dim PrimaryAkunInv As Double = -1
        Dim PrimaryAkunAdjExpense As Double = -1
        Dim Amount1 As Decimal = 0
        Dim TipeTransaksiAsal As ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust
        Dim TipeTransaksi As ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
        Dim ReferenceIndex1 As Integer = -1
        Dim ReferenceIndex2 As Integer = -1

        Try
            Me.Calc_CurrentStep += 1
            Me.LblStep.Text = "Step " + Me.Calc_CurrentStep.ToString + " of " + Me.Calc_TotalStep.ToString + " : creating stock taking journal"
            Me.ProgressBar1.Value = 0
            Me.LblProgress.Text = ""
            Me.ProgressBar1.Maximum = Me.dgvItem.RowCount
            'Me.LblProgress.Visible = True

            For Each rowDGV In Me.dgvItem.Rows
                Me.ProgressBar1.Value += 1
                If Me.ProgressBar1.Value > 25 Then
                    If Math.IEEERemainder(Me.ProgressBar1.Value, 100) = 0 Then
                        Me.LblProgress.Text = "Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString
                    ElseIf Me.ProgressBar1.Value = Me.ProgressBar1.Maximum Then
                        Me.LblProgress.Text = "Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString
                    End If
                Else
                    Me.LblProgress.Text = "Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString
                End If

                If IsNumeric(rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value) AndAlso CDbl(rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value) > -1 Then
                    '1. in adjust
                    TipeTransaksi = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
                    TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InStockTakeAdjust
                    Me.TBLItemDetail.DefaultView.RowFilter = "[primary code no] = " + rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value.ToString + _
                    " AND [tipe transaksi asal] = " + CInt(TipeTransaksiAsal).ToString + " AND [tipe item detail] = " + CInt(TipeTransaksi).ToString

                    Amount = 0
                    For Each DV In Me.TBLItemDetail.DefaultView
                        Me.rowItemDetail = DV.Row
                        Amount += Me.rowItemDetail.amount
                    Next

                    If Amount <> 0 Then
                        'inventory debet
                        PrimaryAkunInv = rowDGV.Cells(Me.ColItem_PrimaryInventoryAccount.Index).Value
                        ReferenceIndex1 = Me.GetIndexJurnal(PrimaryAkunInv, True, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InStockTakeAdjust, -1)
                        If ReferenceIndex1 = -1 Then
                            'tambahkan baris baru
                            Me.BuatRowJurnal(PrimaryAkunInv, True, Amount, vDefPrec, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InStockTakeAdjust, -1)
                        Else
                            Amount1 = Me.GetAmountRowJurnal(PrimaryAkunInv, True, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InStockTakeAdjust, -1)
                            Me.SetAmountRowJurnal(PrimaryAkunInv, True, Amount + Amount1, vDefPrec, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InStockTakeAdjust, -1)
                        End If

                        'akun adjust kredit
                        PrimaryAkunAdjExpense = rowDGV.Cells(Me.ColItem_PrimaryStockTakingAccount.Index).Value
                        ReferenceIndex2 = Me.GetIndexJurnal(PrimaryAkunAdjExpense, False, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InStockTakeAdjust, -1)
                        If ReferenceIndex2 = -1 Then
                            'tambahkan baris baru
                            Me.BuatRowJurnal(PrimaryAkunAdjExpense, False, Amount, vDefPrec, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InStockTakeAdjust, -1)
                        Else
                            Amount1 = Me.GetAmountRowJurnal(PrimaryAkunAdjExpense, False, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InStockTakeAdjust, -1)
                            Me.SetAmountRowJurnal(PrimaryAkunAdjExpense, False, Amount + Amount1, vDefPrec, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InStockTakeAdjust, -1)
                        End If

                        If ReferenceIndex1 = -1 And ReferenceIndex2 <> -1 Then
                            'debet buat baris baru
                            'kredit insert amount
                            '-> pindahkan debetnya
                            Me.PindahkanRowJurnal(True, ReferenceIndex2, Me.dgvJurnal.RowCount - 1)
                        ElseIf ReferenceIndex1 <> -1 And ReferenceIndex2 = -1 Then
                            'debet inset amount
                            'kredit buat baris baru
                            'pindahkan kreditnya
                            Me.PindahkanRowJurnal(False, ReferenceIndex1, Me.dgvJurnal.RowCount - 1)
                        End If
                    End If

                    '2. out adjust
                    TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutStockTakeAdjust
                    TipeTransaksi = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
                    Me.TBLItemDetail.DefaultView.RowFilter = "[primary code no] = " + rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value.ToString + _
                     " AND [tipe transaksi asal] = " + CInt(TipeTransaksiAsal).ToString + " AND [tipe item detail] = " + CInt(TipeTransaksi).ToString

                    Amount = 0
                    For Each DV In Me.TBLItemDetail.DefaultView
                        Me.rowItemDetail = DV.Row
                        Amount += Me.rowItemDetail.amount
                    Next

                    If Amount <> 0 Then
                        'akun adjust debet
                        PrimaryAkunAdjExpense = rowDGV.Cells(Me.ColItem_PrimaryStockTakingAccount.Index).Value
                        ReferenceIndex1 = Me.GetIndexJurnal(PrimaryAkunAdjExpense, True, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutStockTakeAdjust, -1)
                        If ReferenceIndex1 = -1 Then
                            'tambahkan baris baru
                            Me.BuatRowJurnal(PrimaryAkunAdjExpense, True, Amount, vDefPrec, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutStockTakeAdjust, -1)
                        Else
                            Amount1 = Me.GetAmountRowJurnal(PrimaryAkunAdjExpense, True, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutStockTakeAdjust, -1)
                            Me.SetAmountRowJurnal(PrimaryAkunAdjExpense, True, Amount + Amount1, vDefPrec, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutStockTakeAdjust, -1)
                        End If
                        'inventory kredit
                        PrimaryAkunInv = rowDGV.Cells(Me.ColItem_PrimaryInventoryAccount.Index).Value
                        ReferenceIndex2 = Me.GetIndexJurnal(PrimaryAkunInv, False, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutStockTakeAdjust, -1)
                        If ReferenceIndex2 = -1 Then
                            'tambahkan baris baru
                            Me.BuatRowJurnal(PrimaryAkunInv, False, Amount, vDefPrec, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutStockTakeAdjust, -1)
                        Else
                            Amount1 = Me.GetAmountRowJurnal(PrimaryAkunInv, False, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutStockTakeAdjust, -1)
                            Me.SetAmountRowJurnal(PrimaryAkunInv, False, Amount + Amount1, vDefPrec, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutStockTakeAdjust, -1)
                        End If

                        If ReferenceIndex1 = -1 And ReferenceIndex2 <> -1 Then
                            'debet buat baris baru
                            'kredit insert amount
                            '-> pindahkan debetnya
                            Me.PindahkanRowJurnal(True, ReferenceIndex2, Me.dgvJurnal.RowCount - 1)
                        ElseIf ReferenceIndex1 <> -1 And ReferenceIndex2 = -1 Then
                            'debet inset amount
                            'kredit buat baris baru
                            'pindahkan kreditnya
                            Me.PindahkanRowJurnal(False, ReferenceIndex1, Me.dgvJurnal.RowCount - 1)
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub PindahkanRowJurnal(ByVal vRemovingDebet As Boolean, ByVal vReferenceIndex As Integer, ByVal vFromIndex As Integer)
        Dim rowDGVTarget, rowDGVBefore, rowDGVAfter As DataGridViewRow
        Dim rowCopy, rowReference As DataGridViewRow
        Dim DummyReference As Integer = vReferenceIndex
        Try
            'Exit Sub
            rowDGVTarget = Me.dgvJurnal.Rows(vReferenceIndex)

            If vRemovingDebet = True Then
                'pindah ke atas reference index, pastikan baris yang atasnya reference index sudah bernilai debet
                'jika tidak naik turun lagi, maksimal sampai 0
                vReferenceIndex = DummyReference
                rowDGVBefore = Me.dgvJurnal.Rows(vReferenceIndex - 1)
                While rowDGVBefore.Cells(Me.ColJournal_Debet.Index).Value = 0 And vReferenceIndex - 1 > 0
                    vReferenceIndex -= 1
                    rowDGVBefore = Me.dgvJurnal.Rows(vReferenceIndex - 1)
                End While
                Me.dgvJurnal.Rows.InsertCopy(vFromIndex, vReferenceIndex)
                rowCopy = Me.dgvJurnal.Rows(vReferenceIndex)
                rowReference = Me.dgvJurnal.Rows(vFromIndex + 1)
                For ij As Integer = 0 To Me.dgvJurnal.ColumnCount - 1
                    rowCopy.Cells(ij).Value = rowReference.Cells(ij).Value
                Next
                Me.dgvJurnal.Rows.RemoveAt(vFromIndex + 1)
            Else
                'pindah ke bawahnya reference index
                vReferenceIndex += 1
                'pindah ke bawah reference index, pastikan baris bawahnya reference index sudah bernilai credit
                'jika tidak turun lagi, maksimal sampai fromindex-1 (atasnya baris from)
                rowDGVAfter = Me.dgvJurnal.Rows(vReferenceIndex + 1)
                While rowDGVAfter.Cells(Me.ColJournal_Kredit.Index).Value = 0 And vReferenceIndex + 1 < vFromIndex - 1
                    vReferenceIndex += 1
                    rowDGVAfter = Me.dgvJurnal.Rows(vReferenceIndex + 1)
                End While
                If vReferenceIndex <> vFromIndex And vReferenceIndex < vFromIndex - 1 Then
                    Me.dgvJurnal.Rows.InsertCopy(vFromIndex, vReferenceIndex)
                    rowCopy = Me.dgvJurnal.Rows(vReferenceIndex)
                    rowReference = Me.dgvJurnal.Rows(vFromIndex + 1)
                    For ij As Integer = 0 To Me.dgvJurnal.ColumnCount - 1
                        rowCopy.Cells(ij).Value = rowReference.Cells(ij).Value
                    Next
                    Me.dgvJurnal.Rows.RemoveAt(vFromIndex + 1)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BuatJurnalAdjustment1(ByVal vDefPrec As Integer, ByVal vStartDate As Date, ByVal vEndDate As Date)
        Dim Filter As String = ""
        Dim DV As DataRowView
        Dim Amount As Decimal = 0
        Dim PrimaryAkunInv As Double = -1
        Dim PrimaryAkunAdjExpense As Double = -1
        Dim Amount1 As Decimal = 0
        Dim TipeTransaksiAsal As ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust
        Dim TipeTransaksi As ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
        Dim TBLDummy As New DseTransaksiMonthlyClosing2.TblOutAdjustDataTable
        Dim rowDummy As DseTransaksiMonthlyClosing2.TblOutAdjustRow
        Dim PrimaryAsset As Double = -1
        Dim TipeOutAdjust As Inventory_Module.clsTransaksiStockInOut.EnumInOutAdjust = Inventory_Module.clsTransaksiStockInOut.EnumInOutAdjust.NoTipe
        Dim ObjAsset As New Core_Module.ClsMasterAsset
        Dim DTA_Adjustment As New DseTransaksiMonthlyClosing2TableAdapters.tbltransaksi_stockinout_item1TableAdapter
        Dim TBL_Adjustment As New DseTransaksiMonthlyClosing2.tbltransaksi_stockinout_item1DataTable
        Dim row_Adjustment As DseTransaksiMonthlyClosing2.tbltransaksi_stockinout_item1Row
        Dim TBLAmount As New DseTransaksiMonthlyClosing2.DataTable3DataTable
        Dim rowAmount As DseTransaksiMonthlyClosing2.DataTable3Row
        'Dim DVAmount As DataRowView
        Dim ReferenceIndex1 As Integer = -1
        Dim ReferenceIndex2 As Integer = -1

        Try
            Me.Calc_CurrentStep += 1
            Me.LblStep.Text = "Step " + Me.Calc_CurrentStep.ToString + " of " + Me.Calc_TotalStep.ToString + " : creating in adjustment journal"
            Me.ProgressBar1.Value = 0
            Me.LblProgress.Text = ""
            'Me.LblProgress.Visible = True

            Me.LblProgress.Text = "Collecting from database"
            DTA_Adjustment.Fill(TBL_Adjustment, CInt(Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust), _
            CInt(Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust), vEndDate, vStartDate)
            Me.ProgressBar1.Maximum = TBL_Adjustment.Rows.Count

            For Each row_Adjustment In TBL_Adjustment.Rows
                Me.ProgressBar1.Value += 1
                If Me.ProgressBar1.Value > 25 Then
                    If Math.IEEERemainder(Me.ProgressBar1.Value, 100) = 0 Then
                        Me.LblProgress.Text = "Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString
                    ElseIf Me.ProgressBar1.Value = Me.ProgressBar1.Maximum Then
                        Me.LblProgress.Text = "Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString
                    End If
                Else
                    Me.LblProgress.Text = "Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString
                End If


                '1. in adjust
                TipeTransaksi = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
                TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust
                Me.TBLItemDetail.DefaultView.RowFilter = "[primary code no] = " + row_Adjustment.Primary_Product.ToString + _
                " AND [tipe transaksi asal] = " + CInt(TipeTransaksiAsal).ToString + " AND [tipe item detail] = " + CInt(TipeTransaksi).ToString


                TBLAmount.Rows.Clear()
                For Each DV In Me.TBLItemDetail.DefaultView
                    Me.rowItemDetail = DV.Row
                    TBLAmount.DefaultView.RowFilter = "[Primary Adjustment Type] = " + Me.rowItemDetail.primary_adjustment_type.ToString
                    If TBLAmount.DefaultView.Count > 0 Then
                        'tambahkan amount
                        rowAmount = TBLAmount.DefaultView(0).Row
                        rowAmount.Amount += Me.rowItemDetail.amount
                    Else
                        'buat baris baru
                        rowAmount = TBLAmount.NewDataTable3Row
                        rowAmount.Primary_Adjustment_Type = Me.rowItemDetail.primary_adjustment_type
                        If Me.rowItemDetail.primary_adjustment_type = -1 Then
                            rowAmount.Primary_Gain_Account = row_Adjustment.Primary_Adjustment_Account
                            rowAmount.Primary_Loss_Account = row_Adjustment.Primary_Adjustment_Account
                        Else
                            rowAmount.Primary_Gain_Account = Me.rowItemDetail.primary_gain_account
                            rowAmount.Primary_Loss_Account = Me.rowItemDetail.primary_loss_account
                        End If
                        rowAmount.Amount = Me.rowItemDetail.amount
                        TBLAmount.Rows.Add(rowAmount)
                    End If
                Next

                If TBLAmount.Rows.Count <> 0 Then
                    For Each rowAmount In TBLAmount.Rows
                        'inventory debet
                        PrimaryAkunInv = row_Adjustment.Primary_Inventory_Account
                        ReferenceIndex1 = Me.GetIndexJurnal(PrimaryAkunInv, True, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust, rowAmount.Primary_Adjustment_Type)
                        If ReferenceIndex1 = -1 Then
                            'tambahkan baris baru
                            Me.BuatRowJurnal(PrimaryAkunInv, True, rowAmount.Amount, vDefPrec, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust, _
                            rowAmount.Primary_Adjustment_Type)
                        Else
                            Amount1 = Me.GetAmountRowJurnal(PrimaryAkunInv, True, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust, rowAmount.Primary_Adjustment_Type)
                            Me.SetAmountRowJurnal(PrimaryAkunInv, True, rowAmount.Amount + Amount1, vDefPrec, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust, _
                            rowAmount.Primary_Adjustment_Type)
                        End If

                        'akun adjust kredit --> gain account karena in adjust
                        PrimaryAkunAdjExpense = rowAmount.Primary_Gain_Account
                        ReferenceIndex2 = Me.GetIndexJurnal(PrimaryAkunAdjExpense, False, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust, rowAmount.Primary_Adjustment_Type)
                        If ReferenceIndex2 = -1 Then
                            'tambahkan baris baru
                            Me.BuatRowJurnal(PrimaryAkunAdjExpense, False, rowAmount.Amount, vDefPrec, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust, _
                            rowAmount.Primary_Adjustment_Type)
                        Else
                            Amount1 = Me.GetAmountRowJurnal(PrimaryAkunAdjExpense, False, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust, rowAmount.Primary_Adjustment_Type)
                            Me.SetAmountRowJurnal(PrimaryAkunAdjExpense, False, rowAmount.Amount + Amount1, vDefPrec, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust, _
                            rowAmount.Primary_Adjustment_Type)
                        End If
                        If ReferenceIndex1 = -1 And ReferenceIndex2 <> -1 Then
                            'debet buat baris baru
                            'kredit insert amount
                            '-> pindahkan debetnya
                            Me.PindahkanRowJurnal(True, ReferenceIndex2, Me.dgvJurnal.RowCount - 1)
                        ElseIf ReferenceIndex1 <> -1 And ReferenceIndex2 = -1 Then
                            'debet inset amount
                            'kredit buat baris baru
                            'pindahkan kreditnya
                            Me.PindahkanRowJurnal(False, ReferenceIndex1, Me.dgvJurnal.RowCount - 1)
                        End If
                    Next
                End If
            Next

            '2. out adjust
            Me.Calc_CurrentStep += 1
            Me.LblStep.Text = "Step " + Me.Calc_CurrentStep.ToString + " of " + Me.Calc_TotalStep.ToString + " : creating out adjustment journal"
            Me.ProgressBar1.Value = 0
            Me.LblProgress.Text = ""
            'Me.LblProgress.Visible = True

            Me.ObjClosing.GetTBLDummy_OutAdjust_MaterialUsingAsset(Me.dgvItem, TBLDummy, Me.TBLItemDetail, Me, Me.TBLSetting, Me.LblProgress, Me.ProgressBar1)
            Me.ProgressBar1.Value = 0
            Me.LblProgress.Text = ""
            Me.ProgressBar1.Maximum = TBLDummy.Rows.Count

            For Each rowDummy In TBLDummy.Rows
                Me.ProgressBar1.Value += 1
                If Me.ProgressBar1.Value > 25 Then
                    If Math.IEEERemainder(Me.ProgressBar1.Value, 100) = 0 Then
                        Me.LblProgress.Text = "Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString
                    ElseIf Me.ProgressBar1.Value = Me.ProgressBar1.Maximum Then
                        Me.LblProgress.Text = "Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString
                    End If
                Else
                    Me.LblProgress.Text = "Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString
                End If

                Amount = rowDummy.Amount

                'akun adjust debet
                PrimaryAkunAdjExpense = rowDummy.Primary_Account_Expense
                ReferenceIndex1 = Me.GetIndexJurnal(PrimaryAkunAdjExpense, True, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutAdjust, _
                rowDummy.Primary_Adjustment_Type)
                If ReferenceIndex1 = -1 Then
                    'tambahkan baris baru
                    Me.BuatRowJurnal(PrimaryAkunAdjExpense, True, Amount, vDefPrec, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutAdjust, rowDummy.Primary_Adjustment_Type)
                Else
                    Amount1 = Me.GetAmountRowJurnal(PrimaryAkunAdjExpense, True, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutAdjust, rowDummy.Primary_Adjustment_Type)
                    Me.SetAmountRowJurnal(PrimaryAkunAdjExpense, True, Amount + Amount1, vDefPrec, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutAdjust, _
                    rowDummy.Primary_Adjustment_Type)
                End If

                'inventory kredit
                PrimaryAkunInv = rowDummy.Primary_Account_Inventory
                ReferenceIndex2 = Me.GetIndexJurnal(PrimaryAkunInv, False, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutAdjust, _
                rowDummy.Primary_Adjustment_Type)
                If ReferenceIndex2 = -1 Then
                    'tambahkan baris baru
                    Me.BuatRowJurnal(PrimaryAkunInv, False, Amount, vDefPrec, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutAdjust, rowDummy.Primary_Adjustment_Type)
                Else
                    Amount1 = Me.GetAmountRowJurnal(PrimaryAkunInv, False, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutAdjust, rowDummy.Primary_Adjustment_Type)
                    Me.SetAmountRowJurnal(PrimaryAkunInv, False, Amount + Amount1, vDefPrec, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutAdjust, _
                    rowDummy.Primary_Adjustment_Type)
                End If
                If ReferenceIndex1 = -1 And ReferenceIndex2 <> -1 Then
                    'debet buat baris baru
                    'kredit insert amount
                    '-> pindahkan debetnya
                    Me.PindahkanRowJurnal(True, ReferenceIndex2, Me.dgvJurnal.RowCount - 1)
                ElseIf ReferenceIndex1 <> -1 And ReferenceIndex2 = -1 Then
                    'debet inset amount
                    'kredit buat baris baru
                    'pindahkan kreditnya
                    Me.PindahkanRowJurnal(False, ReferenceIndex1, Me.dgvJurnal.RowCount - 1)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TBLDummy.Dispose()
            ObjAsset = Nothing
            DTA_Adjustment.Dispose()
            TBL_Adjustment.Dispose()
            TBLAmount.Dispose()
        End Try
    End Sub

    Private Sub BuatJurnalPR(ByVal vDefPrec As Integer)
        Dim Filter As String = ""
        'Dim rowDGV As DataGridViewRow
        Dim DV As DataRowView
        Dim Amount As Decimal = 0
        Dim PrimaryAkunInv As Double = -1
        Dim Amount1 As Decimal = 0
        Dim PrimaryAkunCOGS As Double = -1
        Dim TipeTransaksiPR As ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutPR
        Dim DTAAmount As New DseTransaksiMonthlyClosingTableAdapters.TBLPR_AmountTableAdapter
        Dim TBLAmount As New DseTransaksiMonthlyClosing.TBLPR_AmountDataTable
        Dim rowAmount As DseTransaksiMonthlyClosing.TBLPR_AmountRow
        Dim TBLAmount1 As New DseTransaksiMonthlyClosing.TBLPR_Amount1DataTable
        Dim rowAmount1 As DseTransaksiMonthlyClosing.TBLPR_Amount1Row

        Dim StartDate, EndDate As Date
        Dim TipePR As ClsTransaksiPurchaseReturn.EnumTipe_PurchaseReturn = ClsTransaksiPurchaseReturn.EnumTipe_PurchaseReturn.Nota_Retur
        'Dim rowItemDetail1 As DseTransaksiMonthlyClosing2.tbltrans_purchase_closing_item_detailRow
        Dim AmountPurchase As Decimal = 0
        Dim RoundingIDR As Integer = 0

        Dim TBL_PR As New DseTransaksiMonthlyClosing2.tbltransaksi_stockinout_item1DataTable
        Dim row_PR As DseTransaksiMonthlyClosing2.tbltransaksi_stockinout_item1Row

        Dim ReferenceIndex1 As Integer = -1
        Dim ReferenceIndex2 As Integer = -1

        Try
            Me.Calc_CurrentStep += 1
            Me.LblStep.Text = "Step " + Me.Calc_CurrentStep.ToString + " of " + Me.Calc_TotalStep.ToString + " : creating purchase return journal"
            Me.ProgressBar1.Value = 0
            Me.LblProgress.Text = ""
            'Me.LblProgress.Visible = True

            StartDate = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value)
            EndDate = Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value)
            DTAAmount.Fill(TBLAmount, TipePR, EndDate, StartDate)
            RoundingIDR = Me.objCurrency.GetDecimal_DefaultCurrency

            '1. kumpulkan item purchase return amount
            Me.ProgressBar1.Value = 0
            Me.ProgressBar1.Maximum = TBLAmount.Rows.Count
            For Each rowAmount In TBLAmount.Rows
                Me.ProgressBar1.Value += 1
                If Me.ProgressBar1.Value > 25 Then
                    If Math.IEEERemainder(Me.ProgressBar1.Value, 100) = 0 Then
                        Me.LblProgress.Text = "Data preparation 1 of 2 - Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString
                    ElseIf Me.ProgressBar1.Value = Me.ProgressBar1.Maximum Then
                        Me.LblProgress.Text = "Data preparation 1 of 2 - Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString
                    End If
                Else
                    Me.LblProgress.Text = "Data preparation 1 of 2 - Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString
                End If

                If IsNumeric(rowAmount.Primary_Inventory_Account) AndAlso CDbl(rowAmount.Primary_Inventory_Account) > -1 Then
                    TBLAmount1.DefaultView.RowFilter = Nothing
                    TBLAmount1.DefaultView.RowFilter = "[Primary Inventory Account] = " + rowAmount.Primary_Inventory_Account.ToString
                    AmountPurchase = Math.Round(rowAmount.Purchase_Price_In_IDR, RoundingIDR, MidpointRounding.AwayFromZero)
                    If TBLAmount1.DefaultView.Count = 0 Then
                        'tambahkan baris baru
                        rowAmount1 = TBLAmount1.NewTBLPR_Amount1Row
                        rowAmount1.Primary_Inventory_Account = rowAmount.Primary_Inventory_Account
                        rowAmount1.Amount = AmountPurchase
                        rowAmount1.Cost = 0
                        rowAmount1.Primary_COGS_Amount = rowAmount.Primary_COGS_Account
                        TBLAmount1.Rows.Add(rowAmount1)
                    Else
                        'tambahkan amount return 
                        rowAmount1 = TBLAmount1.DefaultView(0).Row
                        rowAmount1.Amount += AmountPurchase
                    End If

                    'update nilai purchase priceny
                    If IsNumeric(rowAmount.Primary_Product) AndAlso CDbl(rowAmount.Primary_Product) > -1 Then
                        Me.TBLItemDetail.DefaultView.RowFilter = Nothing
                        Me.TBLItemDetail.DefaultView.RowFilter = "[primary code no] = " + rowAmount.Primary_Product.ToString + _
                        " AND [tipe transaksi asal] = " + CInt(TipeTransaksiPR).ToString + " AND [primary transaksi asal] = " + rowAmount.Primary_Transaksi.ToString
                        If Me.TBLItemDetail.DefaultView.Count > 0 Then
                            Me.rowItemDetail = Me.TBLItemDetail.DefaultView(0).Row
                            Me.rowItemDetail.return_price = AmountPurchase
                        End If

                        'masukkan ke data tabel untuk melakukan pengecekan cost
                        TBL_PR.DefaultView.RowFilter = Nothing
                        TBL_PR.DefaultView.RowFilter = "[Primary Product] = " & rowAmount.Primary_Product.ToString
                        If TBL_PR.DefaultView.Count = 0 Then
                            row_PR = TBL_PR.Newtbltransaksi_stockinout_item1Row
                            row_PR.Primary_Adjustment_Account = -1
                            row_PR.Primary_Inventory_Account = rowAmount.Primary_Inventory_Account
                            row_PR.Primary_Product = rowAmount.Primary_Product
                            TBL_PR.Rows.Add(row_PR)
                        End If
                    End If

                End If
            Next

            '2. kumpulkan item cost
            Me.ProgressBar1.Value = 0
            Me.ProgressBar1.Maximum = TBL_PR.Rows.Count
            For Each row_PR In TBL_PR.Rows
                Me.ProgressBar1.Value += 1
                If Me.ProgressBar1.Value > 25 Then
                    If Math.IEEERemainder(Me.ProgressBar1.Value, 100) = 0 Then
                        Me.LblProgress.Text = "Data preparation 2 of 2 - Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString
                    ElseIf Me.ProgressBar1.Value = Me.ProgressBar1.Maximum Then
                        Me.LblProgress.Text = "Data preparation 2 of 2 - Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString
                    End If
                Else
                    Me.LblProgress.Text = "Data preparation 2 of 2 - Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString
                End If

                If IsNumeric(row_PR.Primary_Product) AndAlso CDbl(row_PR.Primary_Product) > -1 Then
                    Me.TBLItemDetail.DefaultView.RowFilter = Nothing
                    Me.TBLItemDetail.DefaultView.RowFilter = "[primary code no] = " + row_PR.Primary_Product.ToString + _
                                    " AND [tipe transaksi asal] = " + CInt(TipeTransaksiPR).ToString
                    Amount = 0
                    If Me.TBLItemDetail.DefaultView.Count > 0 Then
                        For Each DV In Me.TBLItemDetail.DefaultView
                            Me.rowItemDetail = DV.Row
                            Amount += Me.rowItemDetail.amount
                        Next
                    End If


                    If Amount <> 0 Then
                        'tambahkan ke item cost
                        PrimaryAkunInv = row_PR.Primary_Inventory_Account
                        TBLAmount1.DefaultView.RowFilter = "[Primary Inventory Account] = " + PrimaryAkunInv.ToString
                        If TBLAmount1.DefaultView.Count = 0 Then
                            'ga mungkin, ada cost, tapi tidak ada purchase return
                            MsgBox("Please contact your vendor if you see this message", MsgBoxStyle.Information, "Information")
                        Else
                            'tambahkan amount return 
                            rowAmount1 = TBLAmount1.DefaultView(0).Row
                            rowAmount1.Cost += Amount
                        End If
                    End If
                End If
            Next

            '3. selisihkan per akun, lalu jurnal
            Me.ProgressBar1.Value = 0
            Me.ProgressBar1.Maximum = TBLAmount1.Rows.Count
            For Each rowAmount1 In TBLAmount1.Rows
                Me.ProgressBar1.Value += 1
                If Me.ProgressBar1.Value > 25 Then
                    If Math.IEEERemainder(Me.ProgressBar1.Value, 100) = 0 Then
                        Me.LblProgress.Text = "Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString
                    ElseIf Me.ProgressBar1.Value = Me.ProgressBar1.Maximum Then
                        Me.LblProgress.Text = "Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString
                    End If
                Else
                    Me.LblProgress.Text = "Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString
                End If

                With rowAmount1
                    rowAmount1.Selisih = rowAmount1.Amount - rowAmount1.Cost
                    If .Selisih > 0 Then
                        'debet inventory
                        ReferenceIndex1 = Me.GetIndexJurnal(.Primary_Inventory_Account, True, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutPR, -1)
                        If ReferenceIndex1 = -1 Then
                            'tambahkan baris baru
                            Me.BuatRowJurnal(.Primary_Inventory_Account, True, Math.Abs(.Selisih), vDefPrec, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutPR, -1)
                        Else
                            Amount1 = Me.GetAmountRowJurnal(.Primary_Inventory_Account, True, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutPR, -1)
                            Me.SetAmountRowJurnal(.Primary_Inventory_Account, True, Math.Abs(.Selisih) + Amount1, vDefPrec, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutPR, -1)
                        End If
                        'kredit cogs
                        ReferenceIndex2 = Me.GetIndexJurnal(.Primary_COGS_Amount, False, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutPR, -1)
                        If ReferenceIndex2 = -1 Then
                            'tambahkan baris baru
                            Me.BuatRowJurnal(.Primary_COGS_Amount, False, Math.Abs(.Selisih), vDefPrec, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutPR, -1)
                        Else
                            Amount1 = Me.GetAmountRowJurnal(.Primary_COGS_Amount, False, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutPR, -1)
                            Me.SetAmountRowJurnal(.Primary_COGS_Amount, False, Math.Abs(.Selisih) + Amount1, vDefPrec, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutPR, -1)
                        End If
                        If ReferenceIndex1 = -1 And ReferenceIndex2 <> -1 Then
                            'debet buat baris baru
                            'kredit insert amount
                            '-> pindahkan debetnya
                            Me.PindahkanRowJurnal(True, ReferenceIndex2, Me.dgvJurnal.RowCount - 1)
                        ElseIf ReferenceIndex1 <> -1 And ReferenceIndex2 = -1 Then
                            'debet inset amount
                            'kredit buat baris baru
                            'pindahkan kreditnya
                            Me.PindahkanRowJurnal(False, ReferenceIndex1, Me.dgvJurnal.RowCount - 1)
                        End If
                    ElseIf .Selisih < 0 Then
                        'debet cogs
                        ReferenceIndex1 = Me.GetIndexJurnal(.Primary_COGS_Amount, True, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutPR, -1)
                        If ReferenceIndex1 = -1 Then
                            'tambahkan baris baru
                            Me.BuatRowJurnal(.Primary_COGS_Amount, True, Math.Abs(.Selisih), vDefPrec, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutPR, -1)
                        Else
                            Amount1 = Me.GetAmountRowJurnal(.Primary_COGS_Amount, True, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutPR, -1)
                            Me.SetAmountRowJurnal(.Primary_COGS_Amount, True, Math.Abs(.Selisih) + Amount1, vDefPrec, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutPR, -1)
                        End If
                        'kredit inventory
                        ReferenceIndex2 = Me.GetIndexJurnal(.Primary_Inventory_Account, False, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutPR, -1)
                        If ReferenceIndex2 = -1 Then
                            'tambahkan baris baru
                            Me.BuatRowJurnal(.Primary_Inventory_Account, False, Math.Abs(.Selisih), vDefPrec, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutPR, -1)
                        Else
                            Amount1 = Me.GetAmountRowJurnal(.Primary_Inventory_Account, False, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutPR, -1)
                            Me.SetAmountRowJurnal(.Primary_Inventory_Account, False, Math.Abs(.Selisih) + Amount1, vDefPrec, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutPR, -1)
                        End If
                        If ReferenceIndex1 = -1 And ReferenceIndex2 <> -1 Then
                            'debet buat baris baru
                            'kredit insert amount
                            '-> pindahkan debetnya
                            Me.PindahkanRowJurnal(True, ReferenceIndex2, Me.dgvJurnal.RowCount - 1)
                        ElseIf ReferenceIndex1 <> -1 And ReferenceIndex2 = -1 Then
                            'debet inset amount
                            'kredit buat baris baru
                            'pindahkan kreditnya
                            Me.PindahkanRowJurnal(False, ReferenceIndex1, Me.dgvJurnal.RowCount - 1)
                        End If
                    End If
                End With
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAAmount.Dispose()
            TBLAmount.Dispose()
            TBLAmount1.Dispose()
            TBL_PR.Dispose()
        End Try
    End Sub

    Private Sub BuatJurnalSalesReturn(ByVal vDefPrec As Integer, ByVal vStartDate As Date, ByVal vEndDate As Date)
        Dim Filter As String = ""
        Dim DV As DataRowView
        Dim Amount As Decimal = 0
        Dim PrimaryAkunInv As Double = -1
        Dim Amount1 As Decimal = 0
        Dim PrimaryAkunSales As Double = -1
        Dim TipeTransaksiSalesR As ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR
        Dim DTA_SalesR As New DseTransaksiMonthlyClosing2TableAdapters.tbltrans_salesreturn_item1TableAdapter
        Dim TBL_SalesR As New DseTransaksiMonthlyClosing2.tbltrans_salesreturn_item1DataTable
        Dim row_SalesR As DseTransaksiMonthlyClosing2.tbltrans_salesreturn_item1Row
        Dim TipeItemDetail As ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
        Dim ReferenceIndex1 As Integer = -1
        Dim ReferenceIndex2 As Integer = -1

        Try
            Me.Calc_CurrentStep += 1
            Me.LblStep.Text = "Step " + Me.Calc_CurrentStep.ToString + " of " + Me.Calc_TotalStep.ToString + " : creating sales return journal"
            Me.ProgressBar1.Value = 0
            Me.LblProgress.Text = ""
            'Me.LblProgress.Visible = True

            DTA_SalesR.Fill(TBL_SalesR, vEndDate, vStartDate)
            Me.ProgressBar1.Maximum = TBL_SalesR.Rows.Count

            For Each row_SalesR In TBL_SalesR.Rows
                Me.ProgressBar1.Value += 1
                If Me.ProgressBar1.Value > 25 Then
                    If Math.IEEERemainder(Me.ProgressBar1.Value, 100) = 0 Then
                        Me.LblProgress.Text = "Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString
                    ElseIf Me.ProgressBar1.Value = Me.ProgressBar1.Maximum Then
                        Me.LblProgress.Text = "Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString
                    End If
                Else
                    Me.LblProgress.Text = "Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString
                End If

                '1. untuk setiap item yang dijual, masukkan nilai inventory pada credit
                Me.TBLItemDetail.DefaultView.RowFilter = Nothing
                Me.TBLItemDetail.DefaultView.RowFilter = "[primary code no] = " + row_SalesR.Primary_Product.ToString + _
                " AND [tipe transaksi asal] = " + CInt(TipeTransaksiSalesR).ToString + " AND [tipe item detail] = " + CInt(TipeItemDetail).ToString

                Amount = 0
                If Me.TBLItemDetail.DefaultView.Count > 0 Then
                    For Each DV In Me.TBLItemDetail.DefaultView
                        Me.rowItemDetail = DV.Row
                        Amount += Me.rowItemDetail.amount
                    Next
                End If


                If Amount <> 0 Then
                    'inventory debet
                    PrimaryAkunInv = row_SalesR.Primary_Inventory_Account
                    ReferenceIndex1 = Me.GetIndexJurnal(PrimaryAkunInv, True, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR, -1)
                    If ReferenceIndex1 = -1 Then
                        'tambahkan baris baru
                        Me.BuatRowJurnal(PrimaryAkunInv, True, Amount, vDefPrec, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR, -1)
                    Else
                        Amount1 = Me.GetAmountRowJurnal(PrimaryAkunInv, True, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR, -1)
                        Me.SetAmountRowJurnal(PrimaryAkunInv, True, Amount + Amount1, vDefPrec, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR, -1)
                    End If

                    'cogs kredit
                    PrimaryAkunSales = row_SalesR.Primary_COGS_Account
                    ReferenceIndex2 = Me.GetIndexJurnal(PrimaryAkunSales, False, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR, -1)
                    If ReferenceIndex2 = -1 Then
                        'tambahkan baris baru
                        Me.BuatRowJurnal(PrimaryAkunSales, False, Amount, vDefPrec, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR, -1)
                    Else
                        Amount1 = Me.GetAmountRowJurnal(PrimaryAkunSales, False, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR, -1)
                        Me.SetAmountRowJurnal(PrimaryAkunSales, False, Amount + Amount1, vDefPrec, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR, -1)
                    End If

                    If ReferenceIndex1 = -1 And ReferenceIndex2 <> -1 Then
                        'debet buat baris baru
                        'kredit insert amount
                        '-> pindahkan debetnya
                        Me.PindahkanRowJurnal(True, ReferenceIndex2, Me.dgvJurnal.RowCount - 1)
                    ElseIf ReferenceIndex1 <> -1 And ReferenceIndex2 = -1 Then
                        'debet inset amount
                        'kredit buat baris baru
                        'pindahkan kreditnya
                        Me.PindahkanRowJurnal(False, ReferenceIndex1, Me.dgvJurnal.RowCount - 1)
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_SalesR.Dispose()
            TBL_SalesR.Dispose()
        End Try
    End Sub

    Private Sub BuatJurnalCOGS(ByVal vDefPrec As Integer, ByVal vStartDate As Date, ByVal vEndDate As Date)
        Dim Filter As String = ""
        Dim DV As DataRowView
        Dim Amount As Decimal = 0
        Dim PrimaryAkunInv As Double = -1
        Dim Amount1 As Decimal = 0
        Dim PrimaryAkunSales As Double = -1
        Dim TipeTransaksiSales As ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutSales
        Dim DTA_Sales As New DseTransaksiMonthlyClosing2TableAdapters.tbltrans_salesinvoice_item2TableAdapter
        Dim TBL_Sales As New DseTransaksiMonthlyClosing2.tbltrans_salesinvoice_item2DataTable
        Dim row_Sales As DseTransaksiMonthlyClosing2.tbltrans_salesinvoice_item2Row
        Dim ReferenceIndex1 As Integer = -1
        Dim ReferenceIndex2 As Integer = -1

        Try
            Me.Calc_CurrentStep += 1
            Me.LblStep.Text = "Step " + Me.Calc_CurrentStep.ToString + " of " + Me.Calc_TotalStep.ToString + " : creating cogs journal"
            Me.ProgressBar1.Value = 0
            Me.LblProgress.Text = ""
            'Me.LblProgress.Visible = True

            DTA_Sales.Fill(TBL_Sales, vEndDate, vStartDate)
            Me.ProgressBar1.Maximum = TBL_Sales.Rows.Count

            For Each row_Sales In TBL_Sales.Rows
                Me.ProgressBar1.Value += 1
                If Me.ProgressBar1.Value > 25 Then
                    If Math.IEEERemainder(Me.ProgressBar1.Value, 100) = 0 Then
                        Me.LblProgress.Text = "Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString
                    ElseIf Me.ProgressBar1.Value = Me.ProgressBar1.Maximum Then
                        Me.LblProgress.Text = "Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString
                    End If
                Else
                    Me.LblProgress.Text = "Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString
                End If

                'Me.LblProgress.Text = "Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString

                '1. untuk setiap item yang dijual, masukkan nilai inventory pada credit
                Me.TBLItemDetail.DefaultView.RowFilter = Nothing
                Me.TBLItemDetail.DefaultView.RowFilter = "[primary code no] = " + row_Sales.Primary_Product.ToString + _
                " AND [tipe transaksi asal] = " + CInt(TipeTransaksiSales).ToString

                Amount = 0
                If Me.TBLItemDetail.DefaultView.Count > 0 Then
                    For Each DV In Me.TBLItemDetail.DefaultView
                        Me.rowItemDetail = DV.Row
                        Amount += Me.rowItemDetail.amount
                    Next
                End If


                If Amount <> 0 Then
                    'cogs debet
                    PrimaryAkunSales = row_Sales.Primary_COGS_Account
                    ReferenceIndex1 = Me.GetIndexJurnal(PrimaryAkunSales, True, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutSales, -1)
                    If ReferenceIndex1 = -1 Then
                        'tambahkan baris baru
                        Me.BuatRowJurnal(PrimaryAkunSales, True, Amount, vDefPrec, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutSales, -1)
                    Else
                        Amount1 = Me.GetAmountRowJurnal(PrimaryAkunSales, True, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutSales, -1)
                        Me.SetAmountRowJurnal(PrimaryAkunSales, True, Amount + Amount1, vDefPrec, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutSales, -1)
                    End If

                    'inventory kredit
                    PrimaryAkunInv = row_Sales.Primary_Inventory_Account
                    ReferenceIndex2 = Me.GetIndexJurnal(PrimaryAkunInv, False, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutSales, -1)
                    If ReferenceIndex2 = -1 Then
                        'tambahkan baris baru
                        Me.BuatRowJurnal(PrimaryAkunInv, False, Amount, vDefPrec, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutSales, -1)
                    Else
                        Amount1 = Me.GetAmountRowJurnal(PrimaryAkunInv, False, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutSales, -1)
                        Me.SetAmountRowJurnal(PrimaryAkunInv, False, Amount + Amount1, vDefPrec, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutSales, -1)
                    End If
                    If ReferenceIndex1 = -1 And ReferenceIndex2 <> -1 Then
                        'debet buat baris baru
                        'kredit insert amount
                        '-> pindahkan debetnya
                        Me.PindahkanRowJurnal(True, ReferenceIndex2, Me.dgvJurnal.RowCount - 1)
                    ElseIf ReferenceIndex1 <> -1 And ReferenceIndex2 = -1 Then
                        'debet inset amount
                        'kredit buat baris baru
                        'pindahkan kreditnya
                        Me.PindahkanRowJurnal(False, ReferenceIndex1, Me.dgvJurnal.RowCount - 1)
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_Sales.Dispose()
            TBL_Sales.Dispose()
        End Try
    End Sub

    Private Function GetProductBundling(ByVal vLvl As Integer, ByVal vListPrimaryProduct As List(Of Double), _
    ByVal vTBLBundling As DseTransaksiMonthlyClosing2.TblBundling1DataTable, ByVal vRejecttedList As List(Of Double)) As Boolean
        Dim ListDummy As New List(Of Double)
        Dim Filter As String = ""
        Dim row As DseTransaksiMonthlyClosing2.TblBundling1Row

        Try
            vListPrimaryProduct.Clear()

            If vLvl = 0 Then
                'cari product dr product before, pastikan dia tidak ada di after
                For Each row In vTBLBundling.Rows
                    If ListDummy.Contains(row.Primary_Product_Before) = False Then
                        ListDummy.Add(row.Primary_Product_Before)
                    End If
                Next

                For Each PrimaryCodeNo As Double In ListDummy
                    Filter = "[Primary Product After] = " + PrimaryCodeNo.ToString
                    vTBLBundling.DefaultView.RowFilter = Nothing
                    vTBLBundling.DefaultView.RowFilter = Filter
                    If vTBLBundling.DefaultView.Count = 0 Then
                        If vRejecttedList.Contains(PrimaryCodeNo) = False And vListPrimaryProduct.Contains(PrimaryCodeNo) = False Then
                            vListPrimaryProduct.Add(PrimaryCodeNo)
                        End If
                    End If
                Next
            Else
                'cari di primary product di product after, kemudian hitung levelnya
                For Each row In vTBLBundling.Rows
                    If ListDummy.Contains(row.Primary_Product_After) = False Then
                        ListDummy.Add(row.Primary_Product_After)
                    End If
                Next

                For Each PrimaryCodeNo As Double In ListDummy
                    If vLvl = Me.GetLvlMax(PrimaryCodeNo, vTBLBundling, 0) Then
                        If vRejecttedList.Contains(PrimaryCodeNo) = False And vListPrimaryProduct.Contains(PrimaryCodeNo) = False Then
                            vListPrimaryProduct.Add(PrimaryCodeNo)
                        End If
                    End If
                Next
            End If

            If vListPrimaryProduct.Count > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetLvlMax(ByVal vPrimaryProduct As Double, ByVal vTBLBundling As DseTransaksiMonthlyClosing2.TblBundling1DataTable, _
    ByVal vAcuan As Integer) As Integer
        Dim Filter As String = ""
        Dim ListItemBefore As New List(Of Double)
        Dim RetVal As Integer = 0
        Dim row As DseTransaksiMonthlyClosing2.TblBundling1Row
        Dim Dummy As Integer = 0
        Dim PrimaryDummy As Double = 0

        Try
            ''#sementara untuk mengahalangi siklik bundling
            'Return 1
            ''#

            Filter = "[Primary Product After] = " + vPrimaryProduct.ToString
            vTBLBundling.DefaultView.RowFilter = Nothing
            vTBLBundling.DefaultView.RowFilter = Filter
            If vTBLBundling.DefaultView.Count = 0 Then
                Return vAcuan
            Else
                For Each dv As DataRowView In vTBLBundling.DefaultView
                    row = dv.Row
                    If ListItemBefore.Contains(row.Primary_Product_Before) = False Then
                        ListItemBefore.Add(row.Primary_Product_Before)
                    End If
                Next
                For Each PrimaryDummy In ListItemBefore
                    Dummy = Me.GetLvlMax(PrimaryDummy, vTBLBundling, vAcuan + 1)
                    ''#sementara untuk menghalangi siklik bundling
                    'If vAcuan > 10 Then
                    '    Return 10
                    'End If
                    ''#
                    If Dummy > RetVal Then
                        RetVal = Dummy
                    End If
                Next
                Return RetVal
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListItemBefore = Nothing
        End Try
    End Function

    'Private Sub HitungCost(ByVal vPrimaryProduct As Double, ByVal vDefPrec As Integer, ByVal vCalcMethod As ClsTransaksiMonthlyClosingItem.EnumCostingMethod)
    '    'Average
    '    Dim Index As Integer = 0
    '    Dim Filter As String = ""
    '    Dim TipeItemDetail1 As ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.EndingBalanceItem
    '    Dim TipeItemDetail2 As ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
    '    Dim QtyEnd As Decimal = 0
    '    Dim QtyBB, QtyIn, QtyOut As Decimal
    '    Dim DV As DataRowView
    '    Dim Amount_BB_In As Decimal = 0
    '    Dim Qty_BB_In As Decimal = 0
    '    Dim AmountOut_EB As Decimal = 0
    '    Dim UpdateBundlingAfter As Boolean = False

    '    Try
    '        QtyBB = 0
    '        QtyIn = 0
    '        QtyOut = 0

    '        '0. buang item ending pada tbl detail
    '        Filter = "[tipe item detail] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.EndingBalanceItem).ToString + _
    '        " AND [primary code no] = " + vPrimaryProduct.ToString
    '        Me.TBLItemDetail.DefaultView.RowFilter = Nothing
    '        Me.TBLItemDetail.DefaultView.RowFilter = Filter
    '        While Me.TBLItemDetail.DefaultView.Count > 0
    '            Me.TBLItemDetail.Rows.Remove(Me.TBLItemDetail.DefaultView(0).Row)
    '        End While

    '        'masuk ke fungsi yang sesuai dengan perhitungan
    '        If vCalcMethod = ClsTransaksiMonthlyClosingItem.EnumCostingMethod.FIFO Then
    '            Me.HitungCostFIFO_LIFO(vPrimaryProduct, vDefPrec, True)
    '            Exit Sub
    '        ElseIf vCalcMethod = ClsTransaksiMonthlyClosingItem.EnumCostingMethod.LIFO Then
    '            Me.HitungCostFIFO_LIFO(vPrimaryProduct, vDefPrec, False)
    '            Exit Sub
    '        End If

    '        '1. item pindahkan data qty dr data tabel
    '        Filter = "[primary code no] = " + vPrimaryProduct.ToString
    '        Me.TBLItemDetail.DefaultView.RowFilter = Nothing
    '        Me.TBLItemDetail.DefaultView.RowFilter = Filter
    '        For Each DV In Me.TBLItemDetail.DefaultView
    '            Me.rowItemDetail = DV.Row
    '            If Me.rowItemDetail.tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.BeginingBalanceItem Then
    '                QtyBB += Me.rowItemDetail.qty
    '                Amount_BB_In += Me.rowItemDetail.amount
    '                Qty_BB_In += Me.rowItemDetail.qty
    '            End If
    '            If Me.rowItemDetail.tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem Then
    '                QtyIn += Me.rowItemDetail.qty
    '                Amount_BB_In += Me.rowItemDetail.amount
    '                Qty_BB_In += Me.rowItemDetail.qty
    '            End If
    '            If Me.rowItemDetail.tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem Then
    '                QtyOut += Me.rowItemDetail.qty
    '            End If
    '        Next
    '        QtyEnd = QtyBB + QtyIn - QtyOut

    '        '2. hitung nilai out + ending berdasarkan average
    '        Filter = "([tipe item detail] = " + CInt(TipeItemDetail1).ToString + _
    '        " OR [tipe item detail] = " + CInt(TipeItemDetail2).ToString + _
    '        ") AND [primary code no] = " + vPrimaryProduct.ToString
    '        Me.TBLItemDetail.DefaultView.RowFilter = Nothing
    '        Me.TBLItemDetail.DefaultView.RowFilter = Filter
    '        Me.'TBLItemDetail.DefaultView.Sort = "[tipe item detail] ASC, [tgl transaksi asal] ASC"

    '        If QtyEnd = 0 Then
    '            'tidak ada sisa saldo ending balance
    '            Index = 0
    '            If Me.TBLItemDetail.DefaultView.Count > 0 Then
    '                For Each DV In Me.TBLItemDetail.DefaultView
    '                    Me.rowItemDetail = DV.Row
    '                    If Index = Me.TBLItemDetail.DefaultView.Count - 1 Then
    '                        'kurangkan amount agar tidak terjadi selisih perhitungan
    '                        Me.rowItemDetail.amount = Amount_BB_In - AmountOut_EB
    '                    Else
    '                        'hitung berdasarkan qty
    '                        Me.rowItemDetail.amount = Me.ObjCD.DigitGrouping((Me.rowItemDetail.qty * Amount_BB_In) / Qty_BB_In, vDefPrec)
    '                        AmountOut_EB += Me.rowItemDetail.amount
    '                    End If
    '                    Me.rowItemDetail.unit_price = Me.ObjCD.DigitGrouping(Me.rowItemDetail.amount / Me.rowItemDetail.qty, vDefPrec)
    '                    If Me.rowItemDetail.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling Then
    '                        Me.SetAmountOutBundling(Me.rowItemDetail.primary_code_no, _
    '                        Me.rowItemDetail.primary_transaksi_asal, Me.rowItemDetail.amount)
    '                        UpdateBundlingAfter = True
    '                    End If
    '                    Index += 1
    '                Next
    '            End If

    '        Else
    '            If Me.TBLItemDetail.DefaultView.Count > 0 Then
    '                For Each DV In Me.TBLItemDetail.DefaultView
    '                    Me.rowItemDetail = DV.Row
    '                    'hitung berdasarkan qty
    '                    If Qty_BB_In = 0 Then
    '                        Me.rowItemDetail.amount = 0
    '                    Else
    '                        Me.rowItemDetail.amount = Me.ObjCD.DigitGrouping((Me.rowItemDetail.qty * Amount_BB_In) / Qty_BB_In, vDefPrec)
    '                    End If
    '                    AmountOut_EB += Me.rowItemDetail.amount
    '                    Me.rowItemDetail.unit_price = Me.ObjCD.DigitGrouping(Me.rowItemDetail.amount / Me.rowItemDetail.qty, vDefPrec)

    '                    If Me.rowItemDetail.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling Then
    '                        Me.SetAmountOutBundling(Me.rowItemDetail.primary_code_no, _
    '                        Me.rowItemDetail.primary_transaksi_asal, Me.rowItemDetail.amount)
    '                        UpdateBundlingAfter = True
    '                    End If
    '                Next
    '            End If


    '            'ada sisa saldo ending balance
    '            Me.rowItemDetail = Me.TBLItemDetail.Newtbltrans_purchase_closing_item_detailRow

    '            With Me.rowItemDetail
    '                .amount = Amount_BB_In - AmountOut_EB
    '                .primary = -1
    '                .primary_code_no = vPrimaryProduct
    '                .primary_transaksi = -1
    '                .primary_transaksi_asal = -1
    '                .qty = QtyEnd
    '                .tgl_transaksi_asal = Me.dtpPeriod.Value
    '                .tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.EndingBalanceItem
    '                .tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust
    '                .unit_price = Me.ObjCD.DigitGrouping(.amount / .qty, vDefPrec)
    '                .return_price = 0
    '                .primary_gain_account = -1
    '                .primary_loss_account = -1
    '                .primary_adjustment_type = -1
    '                .remarks = ""
    '            End With

    '            Me.TBLItemDetail.Rows.Add(Me.rowItemDetail)
    '        End If

    '        '3. hitung nilai in bundling dari item2 yang menggunakan code no in sebagai item before 
    '        '--> update nilai ini ke dalam dgv bundling after, dgv detail
    '        If UpdateBundlingAfter = True Then
    '            Me.HitungNilaiInBundling(vPrimaryProduct, vDefPrec)
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    'Private Sub HitungCostFIFO_LIFO(ByVal vPrimaryProduct As Double, ByVal vDefPrec As Integer, ByVal vASCENDING As Boolean)
    '    Dim Filter As String = ""
    '    Dim TipeItemDetail1 As ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.EndingBalanceItem
    '    Dim TipeItemDetail2 As ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
    '    Dim DV As DataRowView
    '    Dim UpdateBundlingAfter As Boolean = False
    '    Dim TBLDummy As New DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailDataTable
    '    Dim rowDummy As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailRow

    '    Try
    '        '1. item pindahkan data qty dr data tabel
    '        TipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.BeginingBalanceItem
    '        TipeItemDetail2 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
    '        Filter = "[primary code no] = " + vPrimaryProduct.ToString + " AND ([tipe item detail] = " + CInt(TipeItemDetail1).ToString + _
    '        " OR [tipe item detail] = " + CInt(TipeItemDetail2).ToString + ")"
    '        Me.TBLItemDetail.DefaultView.RowFilter = Filter
    '        For Each DV In Me.TBLItemDetail.DefaultView
    '            Me.rowItemDetail = DV.Row
    '            rowDummy = TBLDummy.Newtbltrans_purchase_closing_item_detailRow

    '            For ijk As Integer = 0 To TBLDummy.Columns.Count - 1
    '                rowDummy.Item(ijk) = Me.rowItemDetail.Item(ijk)
    '            Next

    '            TBLDummy.Rows.Add(rowDummy)
    '        Next

    '        '2. hitung nilai out + ending berdasarkan FIFO/LIFO
    '        TipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.EndingBalanceItem
    '        TipeItemDetail2 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
    '        Filter = "([tipe item detail] = " + CInt(TipeItemDetail1).ToString + _
    '        " OR [tipe item detail] = " + CInt(TipeItemDetail2).ToString + _
    '        ") AND [primary code no] = " + vPrimaryProduct.ToString
    '        Me.TBLItemDetail.DefaultView.RowFilter = Filter
    '        Me.'TBLItemDetail.DefaultView.Sort = "[tipe item detail], [tgl transaksi asal], [tipe transaksi asal], [primary transaksi]"
    '        For Each DV In Me.TBLItemDetail.DefaultView
    '            Me.rowItemDetail = DV.Row
    '            Me.rowItemDetail.amount = Me.GetAmount(TBLDummy, vASCENDING, Me.rowItemDetail.qty, vDefPrec)
    '            Me.rowItemDetail.unit_price = Me.ObjCD.DigitGrouping(Me.rowItemDetail.amount / Me.rowItemDetail.qty, vDefPrec)

    '            If Me.rowItemDetail.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling Then
    '                Me.SetAmountOutBundling(Me.rowItemDetail.primary_code_no, _
    '                Me.rowItemDetail.primary_transaksi_asal, Me.rowItemDetail.amount)
    '                UpdateBundlingAfter = True
    '            End If
    '        Next

    '        '3. jika ada sisa baris pada tabel dummy, maka itu adalah ending balance
    '        TBLDummy.DefaultView.RowFilter = ""
    '        For Each DV In TBLDummy.DefaultView
    '            Me.rowItemDetail = DV.Row
    '            rowDummy = Me.TBLItemDetail.Newtbltrans_purchase_closing_item_detailRow

    '            For ijk As Integer = 0 To TBLDummy.Columns.Count - 1
    '                rowDummy.Item(ijk) = Me.rowItemDetail.Item(ijk)
    '            Next
    '            rowDummy.tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.EndingBalanceItem

    '            TBLItemDetail.Rows.Add(rowDummy)
    '        Next

    '        '4. hitung nilai in bundling dari item2 yang menggunakan code no in sebagai item before 
    '        '--> update nilai ini ke dalam dgv bundling after, dgv detail
    '        If UpdateBundlingAfter = True Then
    '            Me.HitungNilaiInBundling(vPrimaryProduct, vDefPrec)
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        TBLDummy = Nothing
    '    End Try
    'End Sub

    Private Function GetAmount(ByVal vTBLDummy As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailDataTable, _
    ByVal vASCENDING As Boolean, ByVal Qty As Decimal, ByVal vDefPrec As Decimal) As Decimal
        Dim rowDummy As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailRow
        Dim Retval As Decimal = 0
        Dim PartialAmount As Decimal = 0

        Try
            If vASCENDING = True Then
                vTBLDummy.DefaultView.Sort = "[tipe item detail], [tgl transaksi asal], [tipe transaksi asal], [primary transaksi]"
            Else
                vTBLDummy.DefaultView.Sort = "[tipe item detail], [tgl transaksi asal] DESC, [tipe transaksi asal], [primary transaksi] DESC"
            End If

            While Qty > 0
                If vTBLDummy.DefaultView.Count > 0 Then
                    rowDummy = vTBLDummy.DefaultView(0).Row

                    If Qty < rowDummy.qty Then
                        PartialAmount = Me.ObjCD.DigitGrouping((Qty * rowDummy.amount) / rowDummy.qty, vDefPrec)
                        Retval += PartialAmount
                        rowDummy.qty -= Qty
                        rowDummy.amount -= PartialAmount
                        Return Retval
                    Else
                        Retval += rowDummy.amount
                        Qty -= rowDummy.qty
                        vTBLDummy.Rows.Remove(rowDummy)
                    End If
                Else
                    '#sementara, hanya untuk bundling LMB, msgbox() tidak keluar
                    MsgBox("Logic error, contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                    Return Retval
                End If
            End While

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub dgvJurnal_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvJurnal.CellContentClick
        Dim DV As DataRowView
        Dim rowItemDetailDummy As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailRow
        Dim TipeItemInt As ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
        Dim TipeAsalInt As ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust

        Try
            If e.RowIndex <> -1 Then
                If e.ColumnIndex = Me.ColJournal_TipeTransaksiAsalStr.Index Then
                    If Me.dgvJurnal.Item(Me.ColJournal_TipeTransaksiAsalStr.Index, e.RowIndex).Value <> "n/a" And _
                    Me.dgvJurnal.Item(Me.ColJournal_TipeTransaksiAsalStr.Index, e.RowIndex).Value <> "Overhead Cost" Then
                        Dim ij As New FrmTransaksiMonthlyClosing_DetailAkun
                        ij.vDGVItem = Me.dgvItem
                        ij.vFrmAcuan = Me
                        ij.vPeriod = Me.dtpPeriod.Value
                        ij.vTipeTrasaksiAsal = Me.dgvJurnal.Item(Me.ColJournal_TipeTransaksiAsalInt.Index, e.RowIndex).Value
                        ij.vTblSettingItem = Me.TBLSettingItem

                        ij.vUserCode = Me.UserCode
                        If CDec(Me.dgvJurnal.Item(Me.ColJournal_Debet.Index, e.RowIndex).Value) > 0 Then
                            ij.IsDebet = True
                        Else
                            ij.IsDebet = False
                        End If
                        ij.vPrimaryAkun = Me.dgvJurnal.Item(Me.ColJournal_PrimaryAkun.Index, e.RowIndex).Value
                        ij.vPrimaryTipeAdjustment = Me.dgvJurnal.Item(Me.ColJournal_PrimaryAdjustmentType.Index, e.RowIndex).Value

                        TipeAsalInt = Me.dgvJurnal.Item(Me.ColJournal_TipeTransaksiAsalInt.Index, e.RowIndex).Value
                        If TipeAsalInt = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust Or _
                        TipeAsalInt = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling Or _
                        TipeAsalInt = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InPurchase Or _
                        TipeAsalInt = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR Or _
                        TipeAsalInt = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InStockTakeAdjust Then
                            TipeItemInt = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
                        Else
                            TipeItemInt = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
                        End If

                        Me.TBLItemDetail.DefaultView.RowFilter = "[tipe transaksi asal] = " + CInt(Me.dgvJurnal.Item(Me.ColJournal_TipeTransaksiAsalInt.Index, e.RowIndex).Value).ToString + _
                        " AND [tipe item detail] = " + CInt(TipeItemInt).ToString
                        ij.TBLItemDetail.Rows.Clear()

                        For Each DV In Me.TBLItemDetail.DefaultView
                            Me.rowItemDetail = DV.Row
                            rowItemDetailDummy = ij.TBLItemDetail.Newtbltrans_purchase_closing_item_detailRow
                            For jk As Integer = 0 To ij.TBLItemDetail.Columns.Count - 1
                                rowItemDetailDummy.Item(jk) = Me.rowItemDetail.Item(jk)
                            Next
                            ij.TBLItemDetail.Rows.Add(rowItemDetailDummy)
                        Next
                        ij.ShowDialog()

                        ij.Dispose()
                    ElseIf Me.dgvJurnal.Item(Me.ColJournal_TipeTransaksiAsalStr.Index, e.RowIndex).Value = "Overhead Cost" Then
                        Me.TabControl1.SelectedTab = Me.TabPage4
                        Me.TabControl2.SelectedTab = Me.TabPage8
                        Me.rboGroup_Account.Checked = True
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemAdjust_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItemAdjust.RowsAdded
        Try
            'MsgBox("")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahAdaBundling() As Boolean
        Try
            Me.TBLItemDetail.DefaultView.RowFilter = "[tipe transaksi asal] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling).ToString + _
            " AND [tipe item detail] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem).ToString

            If Me.TBLItemDetail.DefaultView.Count > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
End Class
