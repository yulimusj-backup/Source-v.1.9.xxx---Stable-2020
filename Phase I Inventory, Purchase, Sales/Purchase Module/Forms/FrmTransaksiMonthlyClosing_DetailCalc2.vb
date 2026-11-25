Option Explicit On
Option Strict Off

Public Class FrmTransaksiMonthlyClosing_DetailCalc2
    Public TBLItemDetail As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailDataTable
    Public vPrimaryProduct As Double
    Public Period As Date = Today
    Public UserCode As String = ""

    Dim CostingMethod As ClsTransaksiMonthlyClosingItem.EnumCostingMethod = ClsTransaksiMonthlyClosingItem.EnumCostingMethod.Average
    Dim rowItemDetail As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailRow
    Dim DS As New DseTransaksiMonthlyClosing2
    Dim TBL As DseTransaksiMonthlyClosing2.TblDetailCalc1DataTable = DS.TblDetailCalc1
    Dim row, rowBefore As DseTransaksiMonthlyClosing2.TblDetailCalc1Row
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim ObjProd As New Core_Module.clsMasterProduct
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting
    Dim ObjSettingUser As New Core_Module.ClsTransaksiSettingUser
    Dim Report1 As New CrpTransaksiMontlyClosing_DetailCalc
    Dim Aktif As Boolean = False
    Dim TBLRekap As DseTransaksiMonthlyClosing2.TblDetailCalcRekapDataTable = DS.TblDetailCalcRekap
    Dim rowRekap, rowRekapBefore As DseTransaksiMonthlyClosing2.TblDetailCalcRekapRow
    Dim DTAProduct As New DseMonthlyClosingNewTableAdapters.tblmaster_product1TableAdapter
    Dim TBLProduct As New DseMonthlyClosingNew.tblmaster_product1DataTable
    Dim rowProduct As DseMonthlyClosingNew.tblmaster_product1Row
    Dim DefPrec As Integer = 0

    Private Sub IsikanDataRekap()
        Dim TipeItemDetail1 As ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.EndingBalanceItem
        Dim TipeTransaksiAsal As ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust
        Dim Qty, Amount As Decimal
        Dim DV As DataRowView

        Try
            TBLRekap.Rows.Clear()

            'begining balance
            TipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.BeginingBalanceItem
            Qty = 0
            Amount = 0
            Me.TBLItemDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(TipeItemDetail1).ToString + " AND [primary code no] = " + Me.vPrimaryProduct.ToString
            For Each DV In Me.TBLItemDetail.DefaultView
                Me.rowItemDetail = DV.Row
                Qty += Me.rowItemDetail.qty
                Amount += Me.rowItemDetail.amount
            Next
            Me.rowRekap = TBLRekap.NewRow
            With rowRekap
                .Amount = 0
                .Amount_Balance = Amount
                .Qty = 0
                .Qty_Balance = Qty
                .Tipe_Item_Str = "Begining Balance"
            End With
            TBLRekap.Rows.Add(Me.rowRekap)
            'InPurchase 
            TipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
            TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InPurchase
            Qty = 0
            Amount = 0
            Me.TBLItemDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(TipeItemDetail1).ToString + " AND [primary code no] = " + Me.vPrimaryProduct.ToString + _
            " AND [tipe transaksi asal] = " + CInt(TipeTransaksiAsal).ToString
            For Each DV In Me.TBLItemDetail.DefaultView
                Me.rowItemDetail = DV.Row
                Qty += Me.rowItemDetail.qty
                Amount += Me.rowItemDetail.amount
            Next
            If Qty > 0 Then
                Me.rowRekap = TBLRekap.NewRow
                Me.rowRekapBefore = TBLRekap.Rows(Me.TBLRekap.Rows.Count - 1)
                With rowRekap
                    .Amount = Amount
                    .Amount_Balance = Me.rowRekapBefore.Amount_Balance + .Amount
                    .Qty = Qty
                    .Qty_Balance = Me.rowRekapBefore.Qty_Balance + .Qty
                    .Tipe_Item_Str = "In Purchase"
                End With
                TBLRekap.Rows.Add(Me.rowRekap)
            End If
            'InSR 
            TipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
            TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR
            Qty = 0
            Amount = 0
            Me.TBLItemDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(TipeItemDetail1).ToString + " AND [primary code no] = " + Me.vPrimaryProduct.ToString + _
            " AND [tipe transaksi asal] = " + CInt(TipeTransaksiAsal).ToString
            For Each DV In Me.TBLItemDetail.DefaultView
                Me.rowItemDetail = DV.Row
                Qty += Me.rowItemDetail.qty
                Amount += Me.rowItemDetail.amount
            Next
            If Qty > 0 Then
                Me.rowRekap = TBLRekap.NewRow
                Me.rowRekapBefore = TBLRekap.Rows(Me.TBLRekap.Rows.Count - 1)
                With rowRekap
                    .Amount = Amount
                    .Amount_Balance = Me.rowRekapBefore.Amount_Balance + .Amount
                    .Qty = Qty
                    .Qty_Balance = Me.rowRekapBefore.Qty_Balance + .Qty
                    .Tipe_Item_Str = "In Sales Return"
                End With
                TBLRekap.Rows.Add(Me.rowRekap)
            End If
            'InBundling
            TipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
            TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling
            Qty = 0
            Amount = 0
            Me.TBLItemDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(TipeItemDetail1).ToString + " AND [primary code no] = " + Me.vPrimaryProduct.ToString + _
            " AND [tipe transaksi asal] = " + CInt(TipeTransaksiAsal).ToString
            For Each DV In Me.TBLItemDetail.DefaultView
                Me.rowItemDetail = DV.Row
                Qty += Me.rowItemDetail.qty
                Amount += Me.rowItemDetail.amount
            Next
            If Qty > 0 Then
                Me.rowRekap = TBLRekap.NewRow
                Me.rowRekapBefore = TBLRekap.Rows(Me.TBLRekap.Rows.Count - 1)
                With rowRekap
                    .Amount = Amount
                    .Amount_Balance = Me.rowRekapBefore.Amount_Balance + .Amount
                    .Qty = Qty
                    .Qty_Balance = Me.rowRekapBefore.Qty_Balance + .Qty
                    .Tipe_Item_Str = "In Bundling"
                End With
                TBLRekap.Rows.Add(Me.rowRekap)
            End If
            'InAdjust
            TipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
            TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust
            Qty = 0
            Amount = 0
            Me.TBLItemDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(TipeItemDetail1).ToString + " AND [primary code no] = " + Me.vPrimaryProduct.ToString + _
            " AND [tipe transaksi asal] = " + CInt(TipeTransaksiAsal).ToString
            For Each DV In Me.TBLItemDetail.DefaultView
                Me.rowItemDetail = DV.Row
                Qty += Me.rowItemDetail.qty
                Amount += Me.rowItemDetail.amount
            Next
            If Qty > 0 Then
                Me.rowRekap = TBLRekap.NewRow
                Me.rowRekapBefore = TBLRekap.Rows(Me.TBLRekap.Rows.Count - 1)
                With rowRekap
                    .Amount = Amount
                    .Amount_Balance = Me.rowRekapBefore.Amount_Balance + .Amount
                    .Qty = Qty
                    .Qty_Balance = Me.rowRekapBefore.Qty_Balance + .Qty
                    .Tipe_Item_Str = "In Adjust"
                End With
                TBLRekap.Rows.Add(Me.rowRekap)
            End If
            'InStockTakeAdjust
            TipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
            TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InStockTakeAdjust
            Qty = 0
            Amount = 0
            Me.TBLItemDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(TipeItemDetail1).ToString + " AND [primary code no] = " + Me.vPrimaryProduct.ToString + _
            " AND [tipe transaksi asal] = " + CInt(TipeTransaksiAsal).ToString
            For Each DV In Me.TBLItemDetail.DefaultView
                Me.rowItemDetail = DV.Row
                Qty += Me.rowItemDetail.qty
                Amount += Me.rowItemDetail.amount
            Next
            If Qty > 0 Then
                Me.rowRekap = TBLRekap.NewRow
                Me.rowRekapBefore = TBLRekap.Rows(Me.TBLRekap.Rows.Count - 1)
                With rowRekap
                    .Amount = Amount
                    .Amount_Balance = Me.rowRekapBefore.Amount_Balance + .Amount
                    .Qty = Qty
                    .Qty_Balance = Me.rowRekapBefore.Qty_Balance + .Qty
                    .Tipe_Item_Str = "In Stock Take"
                End With
                TBLRekap.Rows.Add(Me.rowRekap)
            End If
            'OutSales 
            TipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
            TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutSales
            Qty = 0
            Amount = 0
            Me.TBLItemDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(TipeItemDetail1).ToString + " AND [primary code no] = " + Me.vPrimaryProduct.ToString + _
            " AND [tipe transaksi asal] = " + CInt(TipeTransaksiAsal).ToString
            For Each DV In Me.TBLItemDetail.DefaultView
                Me.rowItemDetail = DV.Row
                Qty += Me.rowItemDetail.qty
                Amount += Me.rowItemDetail.amount
            Next
            If Qty > 0 Then
                Me.rowRekap = TBLRekap.NewRow
                Me.rowRekapBefore = TBLRekap.Rows(Me.TBLRekap.Rows.Count - 1)
                With rowRekap
                    .Amount = Amount * -1
                    .Amount_Balance = Me.rowRekapBefore.Amount_Balance + .Amount
                    .Qty = Qty * -1
                    .Qty_Balance = Me.rowRekapBefore.Qty_Balance + .Qty
                    .Tipe_Item_Str = "Out Sales"
                End With
                TBLRekap.Rows.Add(Me.rowRekap)
            End If
            'OutBundling
            TipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
            TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling
            Qty = 0
            Amount = 0
            Me.TBLItemDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(TipeItemDetail1).ToString + " AND [primary code no] = " + Me.vPrimaryProduct.ToString + _
            " AND [tipe transaksi asal] = " + CInt(TipeTransaksiAsal).ToString
            For Each DV In Me.TBLItemDetail.DefaultView
                Me.rowItemDetail = DV.Row
                Qty += Me.rowItemDetail.qty
                Amount += Me.rowItemDetail.amount
            Next
            If Qty > 0 Then
                Me.rowRekap = TBLRekap.NewRow
                Me.rowRekapBefore = TBLRekap.Rows(Me.TBLRekap.Rows.Count - 1)
                With rowRekap
                    .Amount = Amount * -1
                    .Amount_Balance = Me.rowRekapBefore.Amount_Balance + .Amount
                    .Qty = Qty * -1
                    .Qty_Balance = Me.rowRekapBefore.Qty_Balance + .Qty
                    .Tipe_Item_Str = "Out Bundling"
                End With
                TBLRekap.Rows.Add(Me.rowRekap)
            End If
            'OutPR 
            TipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
            TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutPR
            Qty = 0
            Amount = 0
            Me.TBLItemDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(TipeItemDetail1).ToString + " AND [primary code no] = " + Me.vPrimaryProduct.ToString + _
            " AND [tipe transaksi asal] = " + CInt(TipeTransaksiAsal).ToString
            For Each DV In Me.TBLItemDetail.DefaultView
                Me.rowItemDetail = DV.Row
                Qty += Me.rowItemDetail.qty
                Amount += Me.rowItemDetail.amount
            Next
            If Qty > 0 Then
                Me.rowRekap = TBLRekap.NewRow
                Me.rowRekapBefore = TBLRekap.Rows(Me.TBLRekap.Rows.Count - 1)
                With rowRekap
                    .Amount = Amount * -1
                    .Amount_Balance = Me.rowRekapBefore.Amount_Balance + .Amount
                    .Qty = Qty * -1
                    .Qty_Balance = Me.rowRekapBefore.Qty_Balance + .Qty
                    .Tipe_Item_Str = "Out Purchase Return"
                End With
                TBLRekap.Rows.Add(Me.rowRekap)
            End If
            'OutStockTakeAdjust
            TipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
            TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutStockTakeAdjust
            Qty = 0
            Amount = 0
            Me.TBLItemDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(TipeItemDetail1).ToString + " AND [primary code no] = " + Me.vPrimaryProduct.ToString + _
            " AND [tipe transaksi asal] = " + CInt(TipeTransaksiAsal).ToString
            For Each DV In Me.TBLItemDetail.DefaultView
                Me.rowItemDetail = DV.Row
                Qty += Me.rowItemDetail.qty
                Amount += Me.rowItemDetail.amount
            Next
            If Qty > 0 Then
                Me.rowRekap = TBLRekap.NewRow
                Me.rowRekapBefore = TBLRekap.Rows(Me.TBLRekap.Rows.Count - 1)
                With rowRekap
                    .Amount = Amount * -1
                    .Amount_Balance = Me.rowRekapBefore.Amount_Balance + .Amount
                    .Qty = Qty * -1
                    .Qty_Balance = Me.rowRekapBefore.Qty_Balance + .Qty
                    .Tipe_Item_Str = "Out Stock Take Adjust"
                End With
                TBLRekap.Rows.Add(Me.rowRekap)
            End If
            'OutAdjust
            TipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
            TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutAdjust
            Qty = 0
            Amount = 0
            Me.TBLItemDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(TipeItemDetail1).ToString + " AND [primary code no] = " + Me.vPrimaryProduct.ToString + _
            " AND [tipe transaksi asal] = " + CInt(TipeTransaksiAsal).ToString
            For Each DV In Me.TBLItemDetail.DefaultView
                Me.rowItemDetail = DV.Row
                Qty += Me.rowItemDetail.qty
                Amount += Me.rowItemDetail.amount
            Next
            If Qty > 0 Then
                Me.rowRekap = TBLRekap.NewRow
                Me.rowRekapBefore = TBLRekap.Rows(Me.TBLRekap.Rows.Count - 1)
                With rowRekap
                    .Amount = Amount * -1
                    .Amount_Balance = Me.rowRekapBefore.Amount_Balance + .Amount
                    .Qty = Qty * -1
                    .Qty_Balance = Me.rowRekapBefore.Qty_Balance + .Qty
                    .Tipe_Item_Str = "Out Adjust"
                End With
                TBLRekap.Rows.Add(Me.rowRekap)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ShowDetailCalc(ByVal vCostingMethod1 As ClsTransaksiMonthlyClosingItem.EnumCostingMethod)
        Dim TipeItemDetail1 As ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.EndingBalanceItem
        Dim TipeTransaksiAsal As ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust
        Dim Qty, Amount As Decimal
        Dim DV As DataRowView
        Dim DTAAOP As New DseTransaksiMonthlyClosing2TableAdapters.TblTransAOPTableAdapter
        Dim TBLAOP As New DseTransaksiMonthlyClosing2.TblTransAOPDataTable
        Dim rowAOP As DseTransaksiMonthlyClosing2.TblTransAOPRow
        Dim TglSampai, TglDari As Date
        Dim DTASR As New DseTransaksiMonthlyClosing2TableAdapters.TblSR2TableAdapter
        Dim TBLSR As New DseTransaksiMonthlyClosing2.TblSR2DataTable
        Dim rowSR As DseTransaksiMonthlyClosing2.TblSR2Row
        Dim DTABundling As New DseTransaksiMonthlyClosing2TableAdapters.tbltransaksi_stockbundlingTableAdapter
        Dim TBLBundling As New DseTransaksiMonthlyClosing2.tbltransaksi_stockbundlingDataTable
        Dim rowBundling As DseTransaksiMonthlyClosing2.tbltransaksi_stockbundlingRow
        Dim DTAStockInOut As New DseTransaksiMonthlyClosing2TableAdapters.tbltransaksi_stockinoutTableAdapter
        Dim TBLStockInOut As New DseTransaksiMonthlyClosing2.tbltransaksi_stockinoutDataTable
        Dim rowStockInOut As DseTransaksiMonthlyClosing2.tbltransaksi_stockinoutRow
        Dim TipeStockInOut As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust
        Dim DTAStockTake As New DseTransaksiMonthlyClosing2TableAdapters.tbltrans_stock_takingTableAdapter
        Dim TBLStockTake As New DseTransaksiMonthlyClosing2.tbltrans_stock_takingDataTable
        Dim rowStockTake As DseTransaksiMonthlyClosing2.tbltrans_stock_takingRow
        Dim DTAInvoice As New DseTransaksiMonthlyClosing2TableAdapters.TblInvoiceTableAdapter
        Dim TBLInvoice As New DseTransaksiMonthlyClosing2.TblInvoiceDataTable
        Dim rowInvoice As DseTransaksiMonthlyClosing2.TblInvoiceRow
        Dim DTAPR As New DseTransaksiMonthlyClosing2TableAdapters.TblPurchaseReturnTableAdapter
        Dim TBLPR As New DseTransaksiMonthlyClosing2.TblPurchaseReturnDataTable
        Dim rowPR As DseTransaksiMonthlyClosing2.TblPurchaseReturnRow
        Dim TipePR As ClsTransaksiPurchaseReturn.EnumTipe_PurchaseReturn = ClsTransaksiPurchaseReturn.EnumTipe_PurchaseReturn.Nota_Retur
        Dim Negatif As Integer = 0
        Dim TBLDummy As New DseTransaksiMonthlyClosing2.TblDetailCalc1DataTable
        Dim rowDummy As DseTransaksiMonthlyClosing2.TblDetailCalc1Row
        Dim TglDariPurchase, TglDariInSR, TglDariInAdjust, TglDariInBundling As Date
        Dim DocumentNoAsal As String = ""
        Dim TipeTransaksiAsalSTR As String = ""

        Try
            Me.IsikanDataRekap()

            Me.chkShowDetailBegining.Enabled = False
            Me.chkUseSupplierInvoiceNumber.Enabled = False

            DTAAOP.Connection.ConnectionString = Core_Module.GetConstring
            DTASR.Connection.ConnectionString = DTAAOP.Connection.ConnectionString
            DTABundling.Connection.ConnectionString = DTAAOP.Connection.ConnectionString
            DTAStockInOut.Connection.ConnectionString = DTAAOP.Connection.ConnectionString
            DTAStockTake.Connection.ConnectionString = DTAAOP.Connection.ConnectionString
            DTAInvoice.Connection.ConnectionString = DTAAOP.Connection.ConnectionString
            DTAPR.Connection.ConnectionString = DTAAOP.Connection.ConnectionString

            TglSampai = Me.ObjCD.GetLasDateOfTheMonth(Me.Period)
            TglDari = Me.ObjCD.GetFirstDateofTheMonth(Me.Period)

            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.TBL.Rows.Clear()

            'isikan data awal untuk menentukan no transaksi
            TglDariPurchase = Date.MaxValue
            TglDariInAdjust = Date.MaxValue
            TglDariInSR = Date.MaxValue
            TglDariInBundling = Date.MaxValue
            Me.TBLItemDetail.DefaultView.RowFilter = "[primary code no] = " + Me.vPrimaryProduct.ToString

            For Each DV In Me.TBLItemDetail.DefaultView
                Me.rowItemDetail = DV.Row
                TipeTransaksiAsal = Me.rowItemDetail.tipe_transaksi_asal
                If TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InPurchase Then
                    'in purchase
                    If TglDariPurchase > Me.rowItemDetail.tgl_transaksi_asal Then
                        TglDariPurchase = Me.rowItemDetail.tgl_transaksi_asal
                    End If
                ElseIf TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling Or _
                TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling Then
                    'in bundling
                    If TglDariInBundling > Me.rowItemDetail.tgl_transaksi_asal Then
                        TglDariInBundling = Me.rowItemDetail.tgl_transaksi_asal
                    End If
                ElseIf TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR Then
                    'in sr
                    If TglDariInSR > Me.rowItemDetail.tgl_transaksi_asal Then
                        TglDariInSR = Me.rowItemDetail.tgl_transaksi_asal
                    End If
                ElseIf TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust Or _
                TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutAdjust Then
                    'in adjust
                    If TglDariInAdjust > Me.rowItemDetail.tgl_transaksi_asal Then
                        TglDariInAdjust = Me.rowItemDetail.tgl_transaksi_asal
                    End If
                End If
            Next
            If TglDariPurchase <> Date.MaxValue Then
                DTAAOP.Fill(TBLAOP, TglSampai, TglDariPurchase)
            End If
            If TglDariInAdjust <> Date.MaxValue Then
                TipeStockInOut = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust
                DTAStockInOut.Fill(TBLStockInOut, TglSampai, TglDariInAdjust)
            End If
            If TglDariInSR <> Date.MaxValue Then
                DTASR.Fill(TBLSR, TglSampai, TglDariInSR)
            End If
            If TglDariInBundling <> Date.MaxValue Then
                DTABundling.Fill(TBLBundling, TglSampai, TglDariInBundling)
            End If
            DTAStockTake.Fill(TBLStockTake)

            'isikan data begining balance
            Qty = 0
            Amount = 0
            TipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.BeginingBalanceItem
            Me.TBLItemDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(TipeItemDetail1).ToString + " AND [primary code no] = " + Me.vPrimaryProduct.ToString
            Me.TBLItemDetail.DefaultView.Sort = "[tgl transaksi asal], [tipe transaksi asal], [primary transaksi asal]"
            If Me.chkShowDetailBegining.Checked = False Then
                For Each DV In Me.TBLItemDetail.DefaultView
                    Me.rowItemDetail = DV.Row
                    Qty += Me.rowItemDetail.qty
                    Amount += Me.rowItemDetail.amount
                Next
                Me.row = TBL.NewRow
                With row
                    .Amount = 0
                    .Amount_Balance = Amount
                    .No_Dokumen_Asal = "Begining Balance"
                    .Qty = 0
                    .Qty_Balance = Qty
                    .Tanggal = Me.ObjCD.GetFirstDateofTheMonth(Me.Period)
                    .Tipe_Item_Int = CInt(TipeItemDetail1)
                    .Tipe_Transaksi_Asal_Str = ""
                    .Tipe_Transaksi_Int = 0
                    .Tipe_Item_Str = "Begining Balance"
                    .Overhead = 0
                    .Remarks = ""
                End With
                TBL.Rows.Add(Me.row)
            Else
                If Me.TBLItemDetail.DefaultView.Count = 0 Then
                    Me.row = TBL.NewRow
                    With row
                        .Amount = 0
                        .Amount_Balance = Amount
                        .No_Dokumen_Asal = "Begining Balance"
                        .Qty = 0
                        .Qty_Balance = Qty
                        .Tanggal = Me.ObjCD.GetFirstDateofTheMonth(Me.Period)
                        .Tipe_Item_Int = CInt(TipeItemDetail1)
                        .Tipe_Transaksi_Asal_Str = ""
                        .Tipe_Transaksi_Int = 0
                        .Tipe_Item_Str = "Begining Balance"
                        .Overhead = 0
                        .Remarks = ""
                    End With
                    TBL.Rows.Add(Me.row)
                Else
                    For Each DV In Me.TBLItemDetail.DefaultView
                        Me.rowItemDetail = DV.Row
                        TipeTransaksiAsal = Me.rowItemDetail.tipe_transaksi_asal
                        If TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InPurchase Then
                            'in purchase
                            TipeTransaksiAsalSTR = "In Purchase"
                            TBLAOP.DefaultView.RowFilter = "Primary = " + Me.rowItemDetail.primary_transaksi_asal.ToString
                            rowAOP = TBLAOP.DefaultView(0).Row
                            If Me.chkUseSupplierInvoiceNumber.Checked = True Then
                                If rowAOP.Supplier_Invoice_No.Trim <> "" Then
                                    DocumentNoAsal = rowAOP.Supplier_Invoice_No + vbCrLf + rowAOP.Supplier_Name
                                Else
                                    DocumentNoAsal = rowAOP.AOP_Number + vbCrLf + rowAOP.Supplier_Name
                                End If
                            Else
                                DocumentNoAsal = rowAOP.AOP_Number + vbCrLf + rowAOP.Supplier_Name
                            End If
                        ElseIf TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling Then
                            'in bundling
                            TipeTransaksiAsalSTR = "In Bundling"
                            TBLBundling.DefaultView.RowFilter = "Primary = " + Me.rowItemDetail.primary_transaksi_asal.ToString
                            rowBundling = TBLBundling.DefaultView(0).Row
                            DocumentNoAsal = rowBundling.No_Transaksi
                        ElseIf TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR Then
                            'in sr
                            TipeTransaksiAsalSTR = "In Sales Return"
                            TBLSR.DefaultView.RowFilter = "[Primary Stock In Out] = " + Me.rowItemDetail.primary_transaksi_asal.ToString
                            rowSR = TBLSR.DefaultView(0).Row
                            DocumentNoAsal = rowSR.No_SR + vbCrLf + rowSR.Customer_Name
                        ElseIf TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust Then
                            'in adjust
                            TipeTransaksiAsalSTR = "In Adjust"
                            TBLStockInOut.DefaultView.RowFilter = "Primary = " + Me.rowItemDetail.primary_transaksi_asal.ToString
                            rowStockInOut = TBLStockInOut.DefaultView(0).Row
                            If rowStockInOut.Tipe_Adjustment <> "" Then
                                DocumentNoAsal = rowStockInOut.No_Transaksi + " (" + rowStockInOut.Tipe_Adjustment + ")"
                            Else
                                DocumentNoAsal = rowStockInOut.No_Transaksi
                            End If
                        ElseIf TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InStockTakeAdjust Then
                            TipeTransaksiAsalSTR = "In Stock Take"
                            TBLStockTake.DefaultView.RowFilter = "Primary = " + Me.rowItemDetail.primary_transaksi_asal.ToString
                            rowStockTake = TBLStockTake.DefaultView(0).Row
                            DocumentNoAsal = rowStockTake.No_Transaksi
                        End If

                        If TBL.Rows.Count > 0 Then
                            rowBefore = TBL.Rows(TBL.Rows.Count - 1)
                        End If
                        Me.row = TBL.NewRow
                        With row
                            .Amount = Me.rowItemDetail.amount
                            .No_Dokumen_Asal = DocumentNoAsal
                            .Qty = Me.rowItemDetail.qty
                            .Tanggal = Me.rowItemDetail.tgl_transaksi_asal
                            .Tipe_Item_Int = Me.rowItemDetail.tipe_item_detail
                            .Tipe_Transaksi_Asal_Str = TipeTransaksiAsalSTR
                            .Tipe_Transaksi_Int = 0
                            .Tipe_Item_Str = "Begining Balance"
                            .Overhead = Me.rowItemDetail.Overhead
                            .Remarks = Me.rowItemDetail.remarks

                            If TBL.Rows.Count > 0 Then
                                .Amount_Balance = rowBefore.Amount + .Amount
                                .Qty_Balance = rowBefore.Qty + .Qty
                            Else
                                .Amount_Balance = .Amount
                                .Qty_Balance = .Qty
                            End If
                        End With
                        TBL.Rows.Add(Me.row)
                    Next
                End If
            End If

            'isikan data in
            'InPurchase
            TglSampai = Me.ObjCD.GetLasDateOfTheMonth(Me.Period)
            TglDari = Me.ObjCD.GetFirstDateofTheMonth(Me.Period)
            'DTAAOP.Fill(TBLAOP, TglSampai, TglDari)
            TipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
            TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InPurchase
            Me.TBLItemDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(TipeItemDetail1).ToString + _
            " AND [primary code no] = " + Me.vPrimaryProduct.ToString + _
            " AND [tipe transaksi asal] = " + CInt(TipeTransaksiAsal).ToString
            Me.TBLItemDetail.DefaultView.Sort = "[tgl transaksi asal], [tipe transaksi asal], [primary transaksi asal]"
            For Each DV In Me.TBLItemDetail.DefaultView
                Me.rowItemDetail = DV.Row
                Me.rowBefore = Me.TBL.Rows(Me.TBL.Rows.Count - 1)
                TBLAOP.DefaultView.RowFilter = "Primary = " + Me.rowItemDetail.primary_transaksi_asal.ToString
                rowAOP = TBLAOP.DefaultView(0).Row

                Me.row = TBL.NewRow
                With row
                    .Amount = Me.rowItemDetail.amount
                    .Amount_Balance = 0
                    If Me.chkUseSupplierInvoiceNumber.Checked = True Then
                        If rowAOP.Supplier_Invoice_No.Trim <> "" Then
                            .No_Dokumen_Asal = rowAOP.Supplier_Invoice_No + vbCrLf + rowAOP.Supplier_Name
                        Else
                            .No_Dokumen_Asal = rowAOP.AOP_Number + vbCrLf + rowAOP.Supplier_Name
                        End If
                    Else
                        .No_Dokumen_Asal = rowAOP.AOP_Number + vbCrLf + rowAOP.Supplier_Name
                    End If
                    .Qty = Me.rowItemDetail.qty
                    .Qty_Balance = 0
                    .Tanggal = Me.rowItemDetail.tgl_transaksi_asal
                    .Tipe_Item_Int = CInt(TipeItemDetail1)
                    .Tipe_Transaksi_Asal_Str = "In Purchase"
                    .Tipe_Transaksi_Int = CInt(TipeTransaksiAsal)
                    .Tipe_Item_Str = ""
                    .Overhead = Me.rowItemDetail.Overhead
                    .Remarks = Me.rowItemDetail.remarks
                End With
                TBL.Rows.Add(Me.row)
            Next

            'InSR
            'DTASR.Fill(TBLSR, TglSampai, TglDari)
            TipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
            TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR
            Me.TBLItemDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(TipeItemDetail1).ToString + _
            " AND [primary code no] = " + Me.vPrimaryProduct.ToString + _
            " AND [tipe transaksi asal] = " + CInt(TipeTransaksiAsal).ToString
            Me.TBLItemDetail.DefaultView.Sort = "[tgl transaksi asal], [tipe transaksi asal], [primary transaksi asal]"
            For Each DV In Me.TBLItemDetail.DefaultView
                Me.rowItemDetail = DV.Row
                Me.rowBefore = Me.TBL.Rows(Me.TBL.Rows.Count - 1)
                TBLSR.DefaultView.RowFilter = "[Primary Stock In Out] = " + Me.rowItemDetail.primary_transaksi_asal.ToString
                rowSR = TBLSR.DefaultView(0).Row

                Me.row = TBL.NewRow
                With row
                    .Amount = Me.rowItemDetail.amount
                    .Amount_Balance = 0
                    .No_Dokumen_Asal = rowSR.No_SR + vbCrLf + rowSR.Customer_Name
                    .Qty = Me.rowItemDetail.qty
                    .Qty_Balance = 0
                    .Tanggal = Me.rowItemDetail.tgl_transaksi_asal
                    .Tipe_Item_Int = CInt(TipeItemDetail1)
                    .Tipe_Transaksi_Asal_Str = "In Sales Return"
                    .Tipe_Transaksi_Int = CInt(TipeTransaksiAsal)
                    .Tipe_Item_Str = ""
                    .Overhead = Me.rowItemDetail.Overhead
                    .Remarks = Me.rowItemDetail.remarks
                End With
                TBL.Rows.Add(Me.row)
            Next
            'InBundling
            'DTABundling.Fill(TBLBundling, TglSampai, TglDari)
            TipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
            TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling
            Me.TBLItemDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(TipeItemDetail1).ToString + _
            " AND [primary code no] = " + Me.vPrimaryProduct.ToString + _
            " AND [tipe transaksi asal] = " + CInt(TipeTransaksiAsal).ToString
            Me.TBLItemDetail.DefaultView.Sort = "[tgl transaksi asal], [tipe transaksi asal], [primary transaksi asal]"
            For Each DV In Me.TBLItemDetail.DefaultView
                Me.rowItemDetail = DV.Row
                Me.rowBefore = Me.TBL.Rows(Me.TBL.Rows.Count - 1)
                TBLBundling.DefaultView.RowFilter = "Primary = " + Me.rowItemDetail.primary_transaksi_asal.ToString
                rowBundling = TBLBundling.DefaultView(0).Row

                Me.row = TBL.NewRow
                With row
                    .Amount = Me.rowItemDetail.amount
                    .Amount_Balance = 0
                    .No_Dokumen_Asal = rowBundling.No_Transaksi
                    .Qty = Me.rowItemDetail.qty
                    .Qty_Balance = 0
                    .Tanggal = Me.rowItemDetail.tgl_transaksi_asal
                    .Tipe_Item_Int = CInt(TipeItemDetail1)
                    .Tipe_Transaksi_Asal_Str = "In Bundling"
                    .Tipe_Transaksi_Int = CInt(TipeTransaksiAsal)
                    .Tipe_Item_Str = ""
                    .Overhead = Me.rowItemDetail.Overhead
                    .Remarks = Me.rowItemDetail.remarks
                End With
                TBL.Rows.Add(Me.row)
            Next
            'InAdjust
            TipeStockInOut = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust
            'DTAStockInOut.Fill(TBLStockInOut, TipeStockInOut, TglSampai, TglDari)
            TipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
            TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust
            Me.TBLItemDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(TipeItemDetail1).ToString + _
            " AND [primary code no] = " + Me.vPrimaryProduct.ToString + _
            " AND [tipe transaksi asal] = " + CInt(TipeTransaksiAsal).ToString
            Me.TBLItemDetail.DefaultView.Sort = "[tgl transaksi asal], [tipe transaksi asal], [primary transaksi asal]"
            For Each DV In Me.TBLItemDetail.DefaultView
                Me.rowItemDetail = DV.Row
                Me.rowBefore = Me.TBL.Rows(Me.TBL.Rows.Count - 1)
                TBLStockInOut.DefaultView.RowFilter = "Primary = " + Me.rowItemDetail.primary_transaksi_asal.ToString
                rowStockInOut = TBLStockInOut.DefaultView(0).Row

                Me.row = TBL.NewRow
                With row
                    .Amount = Me.rowItemDetail.amount
                    .Amount_Balance = 0
                    If rowStockInOut.Tipe_Adjustment <> "" Then
                        .No_Dokumen_Asal = rowStockInOut.No_Transaksi + " (" + rowStockInOut.Tipe_Adjustment + ")"
                    Else
                        .No_Dokumen_Asal = rowStockInOut.No_Transaksi
                    End If
                    .Qty = Me.rowItemDetail.qty
                    .Qty_Balance = 0
                    .Tanggal = Me.rowItemDetail.tgl_transaksi_asal
                    .Tipe_Item_Int = CInt(TipeItemDetail1)
                    .Tipe_Transaksi_Asal_Str = "In Adjust"
                    .Tipe_Transaksi_Int = CInt(TipeTransaksiAsal)
                    .Tipe_Item_Str = ""
                    .Overhead = Me.rowItemDetail.Overhead
                    .Remarks = Me.rowItemDetail.remarks
                End With
                TBL.Rows.Add(Me.row)
            Next
            'InStockTakeAdjust
            DTAStockTake.Fill(TBLStockTake)
            TipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
            TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InStockTakeAdjust
            Me.TBLItemDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(TipeItemDetail1).ToString + _
            " AND [primary code no] = " + Me.vPrimaryProduct.ToString + _
            " AND [tipe transaksi asal] = " + CInt(TipeTransaksiAsal).ToString
            Me.TBLItemDetail.DefaultView.Sort = "[tgl transaksi asal], [tipe transaksi asal], [primary transaksi asal]"
            For Each DV In Me.TBLItemDetail.DefaultView
                Me.rowItemDetail = DV.Row
                Me.rowBefore = Me.TBL.Rows(Me.TBL.Rows.Count - 1)
                TBLStockTake.DefaultView.RowFilter = "Primary = " + Me.rowItemDetail.primary_transaksi_asal.ToString
                rowStockTake = TBLStockTake.DefaultView(0).Row

                Me.row = TBL.NewRow
                With row
                    .Amount = Me.rowItemDetail.amount
                    .Amount_Balance = 0
                    .No_Dokumen_Asal = rowStockTake.No_Transaksi
                    .Qty = Me.rowItemDetail.qty
                    .Qty_Balance = 0
                    .Tanggal = Me.rowItemDetail.tgl_transaksi_asal
                    .Tipe_Item_Int = CInt(TipeItemDetail1)
                    .Tipe_Transaksi_Asal_Str = "In Stock Take"
                    .Tipe_Transaksi_Int = CInt(TipeTransaksiAsal)
                    .Tipe_Item_Str = ""
                    .Overhead = Me.rowItemDetail.Overhead
                    .Remarks = Me.rowItemDetail.remarks
                End With
                TBL.Rows.Add(Me.row)
            Next

            'isikan data out
            'OutSales
            DTAInvoice.Fill(TBLInvoice, TglSampai, TglDari)
            TipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
            TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutSales
            Me.TBLItemDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(TipeItemDetail1).ToString + _
            " AND [primary code no] = " + Me.vPrimaryProduct.ToString + _
            " AND [tipe transaksi asal] = " + CInt(TipeTransaksiAsal).ToString
            Me.TBLItemDetail.DefaultView.Sort = "[tgl transaksi asal], [tipe transaksi asal], [primary transaksi asal]"
            'dianggap in lagi supaya tampilan pada laporan lebih frienly
            'If vCostingMethod <> ClsTransaksiMonthlyClosingItem.EnumCostingMethod.Average Then
            TipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
            'End If
            For Each DV In Me.TBLItemDetail.DefaultView
                Me.rowItemDetail = DV.Row
                Me.rowBefore = Me.TBL.Rows(Me.TBL.Rows.Count - 1)
                TBLInvoice.DefaultView.RowFilter = "Primary = " + Me.rowItemDetail.primary_transaksi_asal.ToString
                rowInvoice = TBLInvoice.DefaultView(0).Row

                Me.row = TBL.NewRow
                With row
                    .Amount = Me.rowItemDetail.amount * -1
                    .Amount_Balance = 0
                    .No_Dokumen_Asal = rowInvoice.No_Transaksi + vbCrLf + rowInvoice.Customer_Name
                    .Qty = Me.rowItemDetail.qty * -1
                    .Qty_Balance = 0
                    .Tanggal = Me.rowItemDetail.tgl_transaksi_asal
                    .Tipe_Item_Int = CInt(TipeItemDetail1)
                    .Tipe_Transaksi_Asal_Str = "Out Sales"
                    .Tipe_Item_Str = ""
                    .Tipe_Transaksi_Int = CInt(TipeTransaksiAsal)
                    .Overhead = Me.rowItemDetail.Overhead * -1
                    If Me.rowItemDetail.remarks <> "" Then
                        .Remarks = Me.TranslateRemarks(Me.rowItemDetail.remarks, TBLAOP, TBLSR, TBLBundling, TBLStockInOut, TBLStockTake)
                    Else
                        .Remarks = Me.rowItemDetail.remarks
                    End If
                End With
                TBL.Rows.Add(Me.row)
            Next
            'OutBundling
            TipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
            TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling
            Me.TBLItemDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(TipeItemDetail1).ToString + _
            " AND [primary code no] = " + Me.vPrimaryProduct.ToString + _
            " AND [tipe transaksi asal] = " + CInt(TipeTransaksiAsal).ToString
            Me.TBLItemDetail.DefaultView.Sort = "[tgl transaksi asal], [tipe transaksi asal], [primary transaksi asal]"
            'dianggap in lagi supaya tampilan pada laporan lebih frienly
            'If vCostingMethod <> ClsTransaksiMonthlyClosingItem.EnumCostingMethod.Average Then
            TipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
            'End If
            For Each DV In Me.TBLItemDetail.DefaultView
                Me.rowItemDetail = DV.Row
                Me.rowBefore = Me.TBL.Rows(Me.TBL.Rows.Count - 1)
                TBLBundling.DefaultView.RowFilter = "Primary = " + Me.rowItemDetail.primary_transaksi_asal.ToString
                rowBundling = TBLBundling.DefaultView(0).Row

                Me.row = TBL.NewRow
                With row
                    .Amount = Me.rowItemDetail.amount * -1
                    .Amount_Balance = 0
                    .No_Dokumen_Asal = rowBundling.No_Transaksi
                    .Qty = Me.rowItemDetail.qty * -1
                    .Qty_Balance = 0
                    .Tanggal = Me.rowItemDetail.tgl_transaksi_asal
                    .Tipe_Item_Int = CInt(TipeItemDetail1)
                    .Tipe_Transaksi_Asal_Str = "Out Bundling"
                    .Tipe_Transaksi_Int = CInt(TipeTransaksiAsal)
                    .Tipe_Item_Str = ""
                    .Overhead = Me.rowItemDetail.Overhead * -1
                    If Me.rowItemDetail.remarks <> "" Then
                        .Remarks = Me.TranslateRemarks(Me.rowItemDetail.remarks, TBLAOP, TBLSR, TBLBundling, TBLStockInOut, TBLStockTake)
                    Else
                        .Remarks = Me.rowItemDetail.remarks
                    End If
                End With
                TBL.Rows.Add(Me.row)
            Next
            'OutPR
            TipePR = ClsTransaksiPurchaseReturn.EnumTipe_PurchaseReturn.Nota_Retur
            DTAPR.Fill(TBLPR, TglSampai, TipePR, TglDari)
            TipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
            TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutPR
            Me.TBLItemDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(TipeItemDetail1).ToString + _
            " AND [primary code no] = " + Me.vPrimaryProduct.ToString + _
            " AND [tipe transaksi asal] = " + CInt(TipeTransaksiAsal).ToString
            Me.TBLItemDetail.DefaultView.Sort = "[tgl transaksi asal], [tipe transaksi asal], [primary transaksi asal]"
            'dianggap in lagi supaya tampilan pada laporan lebih frienly
            'If vCostingMethod <> ClsTransaksiMonthlyClosingItem.EnumCostingMethod.Average Then
            TipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
            'End If
            For Each DV In Me.TBLItemDetail.DefaultView
                Me.rowItemDetail = DV.Row
                Me.rowBefore = Me.TBL.Rows(Me.TBL.Rows.Count - 1)
                TBLPR.DefaultView.RowFilter = "[Primary Transaksi] = " + Me.rowItemDetail.primary_transaksi_asal.ToString
                rowPR = TBLPR.DefaultView(0).Row

                Me.row = TBL.NewRow
                With row
                    .Amount = Me.rowItemDetail.amount * -1
                    .Amount_Balance = 0
                    .No_Dokumen_Asal = rowPR.No_Transaksi + vbCrLf + rowPR.Supplier_Name
                    .Qty = Me.rowItemDetail.qty * -1
                    .Qty_Balance = 0
                    .Tanggal = Me.rowItemDetail.tgl_transaksi_asal
                    .Tipe_Item_Int = CInt(TipeItemDetail1)
                    .Tipe_Transaksi_Asal_Str = "Out Purchase Return"
                    .Tipe_Transaksi_Int = CInt(TipeTransaksiAsal)
                    .Tipe_Item_Str = ""
                    .Overhead = Me.rowItemDetail.Overhead * -1
                    If Me.rowItemDetail.remarks <> "" Then
                        .Remarks = Me.TranslateRemarks(Me.rowItemDetail.remarks, TBLAOP, TBLSR, TBLBundling, TBLStockInOut, TBLStockTake)
                    Else
                        .Remarks = Me.rowItemDetail.remarks
                    End If
                End With
                TBL.Rows.Add(Me.row)
            Next
            'OutStockTakeAdjust
            TipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
            TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutStockTakeAdjust
            Me.TBLItemDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(TipeItemDetail1).ToString + _
            " AND [primary code no] = " + Me.vPrimaryProduct.ToString + _
            " AND [tipe transaksi asal] = " + CInt(TipeTransaksiAsal).ToString
            Me.TBLItemDetail.DefaultView.Sort = "[tgl transaksi asal], [tipe transaksi asal], [primary transaksi asal]"
            'dianggap in lagi supaya tampilan pada laporan lebih frienly
            'If vCostingMethod <> ClsTransaksiMonthlyClosingItem.EnumCostingMethod.Average Then
            TipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
            'End If
            For Each DV In Me.TBLItemDetail.DefaultView
                Me.rowItemDetail = DV.Row
                Me.rowBefore = Me.TBL.Rows(Me.TBL.Rows.Count - 1)
                TBLStockTake.DefaultView.RowFilter = "Primary = " + Me.rowItemDetail.primary_transaksi_asal.ToString
                rowStockTake = TBLStockTake.DefaultView(0).Row

                Me.row = TBL.NewRow
                With row
                    .Amount = Me.rowItemDetail.amount * -1
                    .Amount_Balance = 0
                    .No_Dokumen_Asal = rowStockTake.No_Transaksi
                    .Qty = Me.rowItemDetail.qty * -1
                    .Qty_Balance = 0
                    .Tanggal = Me.rowItemDetail.tgl_transaksi_asal
                    .Tipe_Item_Int = CInt(TipeItemDetail1)
                    .Tipe_Transaksi_Asal_Str = "Out Stock Take"
                    .Tipe_Transaksi_Int = CInt(TipeTransaksiAsal)
                    .Tipe_Item_Str = ""
                    .Overhead = Me.rowItemDetail.Overhead * -1
                    If Me.rowItemDetail.remarks <> "" Then
                        .Remarks = Me.TranslateRemarks(Me.rowItemDetail.remarks, TBLAOP, TBLSR, TBLBundling, TBLStockInOut, TBLStockTake)
                    Else
                        .Remarks = Me.rowItemDetail.remarks
                    End If
                End With
                TBL.Rows.Add(Me.row)
            Next
            'OutAdjust
            TipeStockInOut = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust
            'DTAStockInOut.Fill(TBLStockInOut, TipeStockInOut, TglSampai, TglDari)
            TipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
            TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutAdjust
            Me.TBLItemDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(TipeItemDetail1).ToString + _
            " AND [primary code no] = " + Me.vPrimaryProduct.ToString + _
            " AND [tipe transaksi asal] = " + CInt(TipeTransaksiAsal).ToString
            Me.TBLItemDetail.DefaultView.Sort = "[tgl transaksi asal], [tipe transaksi asal], [primary transaksi asal]"
            'dianggap in lagi supaya tampilan pada laporan lebih frienly
            'If vCostingMethod <> ClsTransaksiMonthlyClosingItem.EnumCostingMethod.Average Then
            TipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
            'End If
            For Each DV In Me.TBLItemDetail.DefaultView
                Me.rowItemDetail = DV.Row
                Me.rowBefore = Me.TBL.Rows(Me.TBL.Rows.Count - 1)
                TBLStockInOut.DefaultView.RowFilter = "Primary = " + Me.rowItemDetail.primary_transaksi_asal.ToString
                rowStockInOut = TBLStockInOut.DefaultView(0).Row

                Me.row = TBL.NewRow
                With row
                    .Amount = Me.rowItemDetail.amount * -1
                    .Amount_Balance = rowBefore.Amount_Balance + .Amount
                    .No_Dokumen_Asal = rowStockInOut.No_Transaksi + " (" + rowStockInOut.Tipe_Adjustment + ")"
                    .Qty = Me.rowItemDetail.qty * -1
                    .Qty_Balance = rowBefore.Qty_Balance + .Qty
                    .Tanggal = Me.rowItemDetail.tgl_transaksi_asal
                    .Tipe_Item_Int = CInt(TipeItemDetail1)
                    .Tipe_Transaksi_Asal_Str = "Out Adjust"
                    .Tipe_Item_Str = ""
                    .Tipe_Transaksi_Int = CInt(TipeTransaksiAsal)
                    .Overhead = Me.rowItemDetail.Overhead * -1
                    If Me.rowItemDetail.remarks <> "" Then
                        .Remarks = Me.TranslateRemarks(Me.rowItemDetail.remarks, TBLAOP, TBLSR, TBLBundling, TBLStockInOut, TBLStockTake)
                    Else
                        .Remarks = Me.rowItemDetail.remarks
                    End If
                End With
                TBL.Rows.Add(Me.row)
            Next

            If Me.chkShowDetailBegining.Checked = True Then
                TipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.EndingBalanceItem
                Me.TBLItemDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(TipeItemDetail1).ToString + " AND [primary code no] = " + Me.vPrimaryProduct.ToString
                Me.TBLItemDetail.DefaultView.Sort = "[tgl transaksi asal], [tipe transaksi asal], [primary transaksi asal]"
                If Me.TBLItemDetail.DefaultView.Count > 0 Then
                    ''isikan data awal untuk menentukan no transaksi
                    'TglDariPurchase = Date.MaxValue
                    'TglDariInAdjust = Date.MaxValue
                    'TglDariInSR = Date.MaxValue
                    'TglDariInBundling = Date.MaxValue
                    'For Each DV In Me.TBLItemDetail.DefaultView
                    '    Me.rowItemDetail = DV.Row
                    '    TipeTransaksiAsal = Me.rowItemDetail.tipe_transaksi_asal
                    '    If TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InPurchase Then
                    '        'in purchase
                    '        If TglDariPurchase > Me.rowItemDetail.tgl_transaksi_asal Then
                    '            TglDariPurchase = Me.rowItemDetail.tgl_transaksi_asal
                    '        End If
                    '    ElseIf TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling Then
                    '        'in bundling
                    '        If TglDariInBundling > Me.rowItemDetail.tgl_transaksi_asal Then
                    '            TglDariInBundling = Me.rowItemDetail.tgl_transaksi_asal
                    '        End If
                    '    ElseIf TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR Then
                    '        'in sr
                    '        If TglDariInSR > Me.rowItemDetail.tgl_transaksi_asal Then
                    '            TglDariInSR = Me.rowItemDetail.tgl_transaksi_asal
                    '        End If
                    '    ElseIf TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust Then
                    '        'in adjust
                    '        If TglDariInAdjust > Me.rowItemDetail.tgl_transaksi_asal Then
                    '            TglDariInAdjust = Me.rowItemDetail.tgl_transaksi_asal
                    '        End If
                    '    End If
                    'Next
                    'TBLAOP.Rows.Clear()
                    'If TglDariPurchase <> Date.MaxValue Then
                    '    DTAAOP.Fill(TBLAOP, TglSampai, TglDariPurchase)
                    'End If
                    'TBLStockInOut.Rows.Clear()
                    'If TglDariInAdjust <> Date.MaxValue Then
                    '    TipeStockInOut = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust
                    '    DTAStockInOut.Fill(TBLStockInOut, TipeStockInOut, TglSampai, TglDariInAdjust)
                    'End If
                    'TBLSR.Rows.Clear()
                    'If TglDariInSR <> Date.MaxValue Then
                    '    DTASR.Fill(TBLSR, TglSampai, TglDariInSR)
                    'End If
                    'TBLBundling.Rows.Clear()
                    'If TglDariInBundling <> Date.MaxValue Then
                    '    DTABundling.Fill(TBLBundling, TglSampai, TglDariInBundling)
                    'End If
                    'TBLStockTake.Rows.Clear()
                    'DTAStockTake.Fill(TBLStockTake)
                    For Each DV In Me.TBLItemDetail.DefaultView
                        Me.rowItemDetail = DV.Row
                        TipeTransaksiAsal = Me.rowItemDetail.tipe_transaksi_asal
                        If TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InPurchase Then
                            'in purchase
                            TipeTransaksiAsalSTR = "In Purchase"
                            TBLAOP.DefaultView.RowFilter = "Primary = " + Me.rowItemDetail.primary_transaksi_asal.ToString
                            If TBLAOP.DefaultView.Count > 0 Then
                                rowAOP = TBLAOP.DefaultView(0).Row
                                If Me.chkUseSupplierInvoiceNumber.Checked = True Then
                                    If rowAOP.Supplier_Invoice_No.Trim <> "" Then
                                        DocumentNoAsal = rowAOP.Supplier_Invoice_No + vbCrLf + rowAOP.Supplier_Name
                                    Else
                                        DocumentNoAsal = rowAOP.AOP_Number + vbCrLf + rowAOP.Supplier_Name
                                    End If
                                Else
                                    DocumentNoAsal = rowAOP.AOP_Number + vbCrLf + rowAOP.Supplier_Name
                                End If
                            Else
                                DocumentNoAsal = ""
                            End If
                        ElseIf TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling Then
                            'in bundling
                            TipeTransaksiAsalSTR = "In Bundling"
                            TBLBundling.DefaultView.RowFilter = "Primary = " + Me.rowItemDetail.primary_transaksi_asal.ToString
                            If TBLBundling.DefaultView.Count > 0 Then
                                rowBundling = TBLBundling.DefaultView(0).Row
                                DocumentNoAsal = rowBundling.No_Transaksi
                            Else
                                DocumentNoAsal = ""
                            End If
                        ElseIf TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR Then
                            'in sr
                            TipeTransaksiAsalSTR = "In Sales Return"
                            TBLSR.DefaultView.RowFilter = "[Primary Stock In Out] = " + Me.rowItemDetail.primary_transaksi_asal.ToString
                            If TBLSR.DefaultView.Count > 0 Then
                                rowSR = TBLSR.DefaultView(0).Row
                                DocumentNoAsal = rowSR.No_SR + vbCrLf + rowSR.Customer_Name
                            Else
                                DocumentNoAsal = ""
                            End If
                        ElseIf TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust Then
                            'in adjust
                            TipeTransaksiAsalSTR = "In Adjust"
                            TBLStockInOut.DefaultView.RowFilter = "Primary = " + Me.rowItemDetail.primary_transaksi_asal.ToString
                            If TBLStockInOut.DefaultView.Count > 0 Then
                                rowStockInOut = TBLStockInOut.DefaultView(0).Row
                                If rowStockInOut.Tipe_Adjustment <> "" Then
                                    DocumentNoAsal = rowStockInOut.No_Transaksi + " (" + rowStockInOut.Tipe_Adjustment + ")"
                                Else
                                    DocumentNoAsal = rowStockInOut.No_Transaksi
                                End If
                            Else
                                DocumentNoAsal = ""
                            End If
                        ElseIf TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InStockTakeAdjust Then
                            TipeTransaksiAsalSTR = "In Stock Take"
                            TBLStockTake.DefaultView.RowFilter = "Primary = " + Me.rowItemDetail.primary_transaksi_asal.ToString
                            If TBLStockTake.DefaultView.Count > 0 Then
                                rowStockTake = TBLStockTake.DefaultView(0).Row
                                DocumentNoAsal = rowStockTake.No_Transaksi
                            Else
                                DocumentNoAsal = ""
                            End If
                        End If

                        If TBL.Rows.Count > 0 Then
                            rowBefore = TBL.Rows(TBL.Rows.Count - 1)
                        End If
                        Me.row = TBL.NewRow
                        With row
                            .Amount = Me.rowItemDetail.amount * -1
                            .No_Dokumen_Asal = DocumentNoAsal
                            .Qty = Me.rowItemDetail.qty * -1
                            .Tanggal = Me.rowItemDetail.tgl_transaksi_asal
                            .Tipe_Item_Int = Me.rowItemDetail.tipe_item_detail
                            .Tipe_Transaksi_Asal_Str = TipeTransaksiAsalSTR
                            .Tipe_Transaksi_Int = TipeTransaksiAsal
                            .Tipe_Item_Str = "Ending Balance"
                            .Overhead = Me.rowItemDetail.Overhead * -1
                            .Remarks = Me.rowItemDetail.remarks

                            If TBL.Rows.Count > 0 Then
                                .Amount_Balance = rowBefore.Amount + .Amount
                                .Qty_Balance = rowBefore.Qty + .Qty
                            Else
                                .Amount_Balance = .Amount
                                .Qty_Balance = .Qty
                            End If
                        End With
                        TBL.Rows.Add(Me.row)
                    Next
                End If
            End If

            'urutkan ulang
            TBLDummy.Rows.Clear()
            For Each row In TBL.Rows
                rowDummy = TBLDummy.NewTblDetailCalc1Row
                For ij As Integer = 0 To TBLDummy.Columns.Count - 1
                    rowDummy.Item(ij) = row.Item(ij)
                Next
                TBLDummy.Rows.Add(rowDummy)
            Next
            TBLDummy.DefaultView.Sort = "[Tipe Item Int], Tanggal, [Tipe Transaksi Int]"
            TBLDummy.DefaultView.RowFilter = ""
            TBL.Rows.Clear()
            For Each DV In TBLDummy.DefaultView
                rowDummy = DV.Row
                row = TBL.NewTblDetailCalc1Row
                For ij As Integer = 0 To TBLDummy.Columns.Count - 1
                    row.Item(ij) = rowDummy.Item(ij)
                Next
                TBL.Rows.Add(row)
            Next
            'hitung ulang balance
            If TBL.Rows.Count > 1 Then
                For ij As Double = 1 To Me.TBL.Rows.Count - 1
                    rowBefore = Me.TBL.Rows(ij - 1)
                    row = TBL.Rows(ij)
                    With row
                        .Qty_Balance = rowBefore.Qty_Balance + .Qty
                        .Amount_Balance = rowBefore.Amount_Balance + .Amount
                        .Amount = .Amount - .Overhead
                    End With
                Next
            End If

            'show report
            Me.Report1.SetDataSource(Me.DS)
            Me.Report1.SetParameterValue(Me.Report1.Parameter_DefPrecAmount.ParameterFieldName, Me.ObjCur.GetDecimal_DefaultCurrency)
            Me.Report1.SetParameterValue(Me.Report1.Parameter_DefPrecQty.ParameterFieldName, Me.ObjProd.GetAngkaDecimal(Me.vPrimaryProduct))
            Me.Report1.SetParameterValue(Me.Report1.Parameter_Period.ParameterFieldName, Me.Period)
            Me.Report1.SetParameterValue(Me.Report1.Parameter_Product.ParameterFieldName, Me.ObjProd.GetProductCode(Me.vPrimaryProduct) + " - " + Me.ObjProd.GetDescription(Me.vPrimaryProduct))
            Me.Report1.SetParameterValue(Me.Report1.Parameter_Unit.ParameterFieldName, Me.ObjProd.GetUnitCode(Me.vPrimaryProduct))
            If ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, _
                           Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets) = "" Then
                Negatif = 1
            Else
                Negatif = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets)
            End If
            Me.Report1.SetParameterValue(Me.Report1.Parameter_UseBracket.ParameterFieldName, Negatif)
            Me.Report1.SetParameterValue(Me.Report1.Parameter_CrpTransaksiMonthly_DetailCalc_Rekap_DP_Qty.ParameterFieldName, Me.ObjProd.GetAngkaDecimal(Me.vPrimaryProduct))
            Me.Report1.SetParameterValue(Me.Report1.Parameter_CrpTransaksiMonthly_DetailCalc_Rekap_DP_Amount.ParameterFieldName, Me.ObjCur.GetDecimal_DefaultCurrency)
            Me.Report1.SetParameterValue(Me.Report1.Parameter_CrpTransaksiMonthly_DetailCalc_Rekap_SubReport_Negative.ParameterFieldName, Negatif)

            Me.crv1.ReportSource = Me.Report1
            Me.crv1.DisplayGroupTree = False
            Me.crv1.DisplayStatusBar = True
            Me.crv1.DisplayToolbar = True
            Me.crv1.ShowGroupTreeButton = False
            Me.crv1.ShowExportButton = True
            Me.crv1.ShowRefreshButton = False
            Me.crv1.Show()

            Me.chkShowDetailBegining.Enabled = True
            Me.chkUseSupplierInvoiceNumber.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
            DTAAOP.Dispose()
            TBLAOP.Dispose()
            DTASR.Dispose()
            TBLSR.Dispose()
            DTABundling.Dispose()
            TBLBundling.Dispose()
            DTAStockInOut.Dispose()
            TBLStockInOut.Dispose()
            DTAStockTake.Dispose()
            TBLStockTake.Dispose()
            DTAInvoice.Dispose()
            TBLInvoice.Dispose()
            DTAPR.Dispose()
            TBLPR.Dispose()
        End Try
    End Sub

    Private Function TranslateRemarks(ByVal vRemarks As String, ByVal TBLAOP As DseTransaksiMonthlyClosing2.TblTransAOPDataTable, _
    ByVal TBLSR As DseTransaksiMonthlyClosing2.TblSR2DataTable, ByVal TBLBundling As DseTransaksiMonthlyClosing2.tbltransaksi_stockbundlingDataTable, _
    ByVal TBLStockInOut As DseTransaksiMonthlyClosing2.tbltransaksi_stockinoutDataTable, _
    ByVal TBLStockTake As DseTransaksiMonthlyClosing2.tbltrans_stock_takingDataTable) As String
        Dim rowAOP As DseTransaksiMonthlyClosing2.TblTransAOPRow
        Dim rowSR As DseTransaksiMonthlyClosing2.TblSR2Row
        Dim rowBundling As DseTransaksiMonthlyClosing2.tbltransaksi_stockbundlingRow
        Dim rowStockInOut As DseTransaksiMonthlyClosing2.tbltransaksi_stockinoutRow
        Dim rowStockTake As DseTransaksiMonthlyClosing2.tbltrans_stock_takingRow
        Dim KeteranganAsal As String() = vRemarks.Split(";")
        Dim Detail As String()
        Dim TipeTransaksiAsal As ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal
        Dim Qty As Decimal = 0
        Dim PrimaryTransaksiAsal As Double = -1
        Dim Amount As Decimal = 0
        Dim RetVal As String = ""
        Dim DocumentNoAsal As String = ""

        Try
            If vRemarks = "" Then
                Return ""
            Else
                If Me.TBLProduct.Rows.Count = 0 Then
                    Me.DTAProduct.Connection.ConnectionString = Core_Module.GetConstring
                    Me.DTAProduct.Fill(Me.TBLProduct)
                    Me.DefPrec = Me.ObjCur.GetDecimal_DefaultCurrency
                End If
                Me.TBLProduct.DefaultView.RowFilter = "primary = " + vPrimaryProduct.ToString
                rowProduct = Me.TBLProduct.DefaultView(0).Row

                For Each Keterangan As String In KeteranganAsal
                    Detail = Keterangan.Split(",")
                    TipeTransaksiAsal = Detail(0)
                    PrimaryTransaksiAsal = Detail(1)
                    Qty = Detail(2)
                    Amount = Detail(3)

                    If TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InPurchase Then
                        'in purchase
                        TBLAOP.DefaultView.RowFilter = "Primary = " + PrimaryTransaksiAsal.ToString
                        rowAOP = TBLAOP.DefaultView(0).Row
                        If Me.chkUseSupplierInvoiceNumber.Checked = True Then
                            If rowAOP.Supplier_Invoice_No.Trim <> "" Then
                                DocumentNoAsal = rowAOP.Supplier_Invoice_No
                            Else
                                DocumentNoAsal = rowAOP.AOP_Number
                            End If
                        Else
                            DocumentNoAsal = rowAOP.AOP_Number
                        End If
                    ElseIf TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling Then
                        'in bundling
                        TBLBundling.DefaultView.RowFilter = "Primary = " + PrimaryTransaksiAsal.ToString
                        rowBundling = TBLBundling.DefaultView(0).Row
                        DocumentNoAsal = rowBundling.No_Transaksi
                    ElseIf TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR Then
                        'in sr
                        TBLSR.DefaultView.RowFilter = "[Primary Stock In Out] = " + PrimaryTransaksiAsal.ToString
                        rowSR = TBLSR.DefaultView(0).Row
                        DocumentNoAsal = rowSR.No_SR + vbCrLf + rowSR.Customer_Name
                    ElseIf TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust Then
                        'in adjust
                        TBLStockInOut.DefaultView.RowFilter = "Primary = " + PrimaryTransaksiAsal.ToString
                        rowStockInOut = TBLStockInOut.DefaultView(0).Row
                        If rowStockInOut.Tipe_Adjustment <> "" Then
                            DocumentNoAsal = rowStockInOut.No_Transaksi + " (" + rowStockInOut.Tipe_Adjustment + ")"
                        Else
                            DocumentNoAsal = rowStockInOut.No_Transaksi
                        End If
                    ElseIf TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InStockTakeAdjust Then
                        TBLStockTake.DefaultView.RowFilter = "Primary = " + PrimaryTransaksiAsal.ToString
                        rowStockTake = TBLStockTake.DefaultView(0).Row
                        DocumentNoAsal = rowStockTake.No_Transaksi
                    End If

                    'While DocumentNoAsal.StartsWith("0")
                    '    DocumentNoAsal = DocumentNoAsal.Substring(1)
                    'End While
                    If RetVal = "" Then
                        RetVal = Me.ObjCD.DigitGrouping(Amount, DefPrec) + vbCrLf + "(" + DocumentNoAsal + " - " + Me.ObjCD.DigitGrouping(Qty, rowProduct._Decimal) + " " + rowProduct.Unit_Code + ")"
                    Else
                        RetVal = RetVal + " + " + vbCrLf + Me.ObjCD.DigitGrouping(Amount, DefPrec) + vbCrLf + "(" + DocumentNoAsal + " - " + Me.ObjCD.DigitGrouping(Qty, rowProduct._Decimal) + " " + rowProduct.Unit_Code + ")"
                    End If
                Next
            End If

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    'Private Sub ShowDetailCalc_Obs(ByVal vCostingMethod As ClsTransaksiMonthlyClosingItem.EnumCostingMethod)
    '    Dim TipeItemDetail1 As ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.EndingBalanceItem
    '    Dim TipeTransaksiAsal As ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust
    '    Dim Qty, Amount As Decimal
    '    Dim DV As DataRowView
    '    Dim DTAAOP As New DseTransaksiMonthlyClosing2TableAdapters.TblTransAOPTableAdapter
    '    Dim TBLAOP As New DseTransaksiMonthlyClosing2.TblTransAOPDataTable
    '    Dim rowAOP As DseTransaksiMonthlyClosing2.TblTransAOPRow
    '    Dim TglSampai, TglDari As Date
    '    Dim DTASR As New DseTransaksiMonthlyClosing2TableAdapters.TblSR2TableAdapter
    '    Dim TBLSR As New DseTransaksiMonthlyClosing2.TblSR2DataTable
    '    Dim rowSR As DseTransaksiMonthlyClosing2.TblSR2Row
    '    Dim DTABundling As New DseTransaksiMonthlyClosing2TableAdapters.tbltransaksi_stockbundlingTableAdapter
    '    Dim TBLBundling As New DseTransaksiMonthlyClosing2.tbltransaksi_stockbundlingDataTable
    '    Dim rowBundling As DseTransaksiMonthlyClosing2.tbltransaksi_stockbundlingRow
    '    Dim DTAStockInOut As New DseTransaksiMonthlyClosing2TableAdapters.tbltransaksi_stockinoutTableAdapter
    '    Dim TBLStockInOut As New DseTransaksiMonthlyClosing2.tbltransaksi_stockinoutDataTable
    '    Dim rowStockInOut As DseTransaksiMonthlyClosing2.tbltransaksi_stockinoutRow
    '    Dim TipeStockInOut As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust
    '    Dim DTAStockTake As New DseTransaksiMonthlyClosing2TableAdapters.tbltrans_stock_takingTableAdapter
    '    Dim TBLStockTake As New DseTransaksiMonthlyClosing2.tbltrans_stock_takingDataTable
    '    Dim rowStockTake As DseTransaksiMonthlyClosing2.tbltrans_stock_takingRow
    '    Dim DTAInvoice As New DseTransaksiMonthlyClosing2TableAdapters.TblInvoiceTableAdapter
    '    Dim TBLInvoice As New DseTransaksiMonthlyClosing2.TblInvoiceDataTable
    '    Dim rowInvoice As DseTransaksiMonthlyClosing2.TblInvoiceRow
    '    Dim DTAPR As New DseTransaksiMonthlyClosing2TableAdapters.TblPurchaseReturnTableAdapter
    '    Dim TBLPR As New DseTransaksiMonthlyClosing2.TblPurchaseReturnDataTable
    '    Dim rowPR As DseTransaksiMonthlyClosing2.TblPurchaseReturnRow
    '    Dim TipePR As ClsTransaksiPurchaseReturn.EnumTipe_PurchaseReturn = ClsTransaksiPurchaseReturn.EnumTipe_PurchaseReturn.Nota_Retur
    '    Dim Negatif As Integer = 0
    '    Dim TBLDummy As New DseTransaksiMonthlyClosing2.TblDetailCalc1DataTable
    '    Dim rowDummy As DseTransaksiMonthlyClosing2.TblDetailCalc1Row
    '    Dim TglDariPurchase, TglDariInSR, TglDariInAdjust, TglDariInBundling As Date
    '    Dim DocumentNoAsal As String = ""
    '    Dim TipeTransaksiAsalSTR As String = ""

    '    Try
    '        Me.IsikanDataRekap()

    '        Me.chkShowDetailBegining.Enabled = False
    '        Me.chkUseSupplierInvoiceNumber.Enabled = False

    '        DTAAOP.Connection.ConnectionString = Core_Module.GetConstring
    '        DTASR.Connection.ConnectionString = DTAAOP.Connection.ConnectionString
    '        DTABundling.Connection.ConnectionString = DTAAOP.Connection.ConnectionString
    '        DTAStockInOut.Connection.ConnectionString = DTAAOP.Connection.ConnectionString
    '        DTAStockTake.Connection.ConnectionString = DTAAOP.Connection.ConnectionString
    '        DTAInvoice.Connection.ConnectionString = DTAAOP.Connection.ConnectionString
    '        DTAPR.Connection.ConnectionString = DTAAOP.Connection.ConnectionString

    '        TglSampai = Me.ObjCD.GetLasDateOfTheMonth(Me.Period)
    '        TglDari = Me.ObjCD.GetFirstDateofTheMonth(Me.Period)

    '        Windows.Forms.Cursor.Current = Cursors.WaitCursor
    '        Me.TBL.Rows.Clear()

    '        'isikan data begining balance
    '        Qty = 0
    '        Amount = 0
    '        TipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.BeginingBalanceItem
    '        Me.TBLItemDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(TipeItemDetail1).ToString + " AND [primary code no] = " + Me.vPrimaryProduct.ToString
    '        Me.TBLItemDetail.DefaultView.Sort = "[tgl transaksi asal], [tipe transaksi asal], [primary transaksi asal]"
    '        If Me.chkShowDetailBegining.Checked = False Then
    '            For Each DV In Me.TBLItemDetail.DefaultView
    '                Me.rowItemDetail = DV.Row
    '                Qty += Me.rowItemDetail.qty
    '                Amount += Me.rowItemDetail.amount
    '            Next
    '            Me.row = TBL.NewRow
    '            With row
    '                .Amount = 0
    '                .Amount_Balance = Amount
    '                .No_Dokumen_Asal = "Begining Balance"
    '                .Qty = 0
    '                .Qty_Balance = Qty
    '                .Tanggal = Me.ObjCD.GetFirstDateofTheMonth(Me.Period)
    '                .Tipe_Item_Int = CInt(TipeItemDetail1)
    '                .Tipe_Transaksi_Asal_Str = ""
    '                .Tipe_Transaksi_Int = 0
    '                .Tipe_Item_Str = "Begining Balance"
    '            End With
    '            TBL.Rows.Add(Me.row)
    '        Else
    '            If Me.TBLItemDetail.DefaultView.Count = 0 Then
    '                Me.row = TBL.NewRow
    '                With row
    '                    .Amount = 0
    '                    .Amount_Balance = Amount
    '                    .No_Dokumen_Asal = "Begining Balance"
    '                    .Qty = 0
    '                    .Qty_Balance = Qty
    '                    .Tanggal = Me.ObjCD.GetFirstDateofTheMonth(Me.Period)
    '                    .Tipe_Item_Int = CInt(TipeItemDetail1)
    '                    .Tipe_Transaksi_Asal_Str = ""
    '                    .Tipe_Transaksi_Int = 0
    '                    .Tipe_Item_Str = "Begining Balance"
    '                End With
    '                TBL.Rows.Add(Me.row)
    '            Else
    '                'isikan data awal untuk menentukan no transaksi
    '                TglDariPurchase = Date.MaxValue
    '                TglDariInAdjust = Date.MaxValue
    '                TglDariInSR = Date.MaxValue
    '                TglDariInBundling = Date.MaxValue
    '                For Each DV In Me.TBLItemDetail.DefaultView
    '                    Me.rowItemDetail = DV.Row
    '                    TipeTransaksiAsal = Me.rowItemDetail.tipe_transaksi_asal
    '                    If TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InPurchase Then
    '                        'in purchase
    '                        If TglDariPurchase > Me.rowItemDetail.tgl_transaksi_asal Then
    '                            TglDariPurchase = Me.rowItemDetail.tgl_transaksi_asal
    '                        End If
    '                    ElseIf TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling Then
    '                        'in bundling
    '                        If TglDariInBundling > Me.rowItemDetail.tgl_transaksi_asal Then
    '                            TglDariInBundling = Me.rowItemDetail.tgl_transaksi_asal
    '                        End If
    '                    ElseIf TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR Then
    '                        'in sr
    '                        If TglDariInSR > Me.rowItemDetail.tgl_transaksi_asal Then
    '                            TglDariInSR = Me.rowItemDetail.tgl_transaksi_asal
    '                        End If
    '                    ElseIf TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust Then
    '                        'in adjust
    '                        If TglDariInAdjust > Me.rowItemDetail.tgl_transaksi_asal Then
    '                            TglDariInAdjust = Me.rowItemDetail.tgl_transaksi_asal
    '                        End If
    '                    End If
    '                Next
    '                If TglDariPurchase <> Date.MaxValue Then
    '                    DTAAOP.Fill(TBLAOP, TglSampai, TglDariPurchase)
    '                End If
    '                If TglDariInAdjust <> Date.MaxValue Then
    '                    TipeStockInOut = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust
    '                    DTAStockInOut.Fill(TBLStockInOut, TipeStockInOut, TglSampai, TglDariInAdjust)
    '                End If
    '                If TglDariInSR <> Date.MaxValue Then
    '                    DTASR.Fill(TBLSR, TglSampai, TglDariInSR)
    '                End If
    '                If TglDariInBundling <> Date.MaxValue Then
    '                    DTABundling.Fill(TBLBundling, TglSampai, TglDariInBundling)
    '                End If
    '                DTAStockTake.Fill(TBLStockTake)
    '                For Each DV In Me.TBLItemDetail.DefaultView
    '                    Me.rowItemDetail = DV.Row
    '                    TipeTransaksiAsal = Me.rowItemDetail.tipe_transaksi_asal
    '                    If TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InPurchase Then
    '                        'in purchase
    '                        TipeTransaksiAsalSTR = "In Purchase"
    '                        TBLAOP.DefaultView.RowFilter = "Primary = " + Me.rowItemDetail.primary_transaksi_asal.ToString
    '                        rowAOP = TBLAOP.DefaultView(0).Row
    '                        If Me.chkUseSupplierInvoiceNumber.Checked = True Then
    '                            If rowAOP.Supplier_Invoice_No.Trim <> "" Then
    '                                DocumentNoAsal = rowAOP.Supplier_Invoice_No + vbCrLf + rowAOP.Supplier_Name
    '                            Else
    '                                DocumentNoAsal = rowAOP.AOP_Number + vbCrLf + rowAOP.Supplier_Name
    '                            End If
    '                        Else
    '                            DocumentNoAsal = rowAOP.AOP_Number + vbCrLf + rowAOP.Supplier_Name
    '                        End If
    '                    ElseIf TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling Then
    '                        'in bundling
    '                        TipeTransaksiAsalSTR = "In Bundling"
    '                        TBLBundling.DefaultView.RowFilter = "Primary = " + Me.rowItemDetail.primary_transaksi_asal.ToString
    '                        rowBundling = TBLBundling.DefaultView(0).Row
    '                        DocumentNoAsal = rowBundling.No_Transaksi
    '                    ElseIf TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR Then
    '                        'in sr
    '                        TipeTransaksiAsalSTR = "In Sales Return"
    '                        TBLSR.DefaultView.RowFilter = "[Primary Stock In Out] = " + Me.rowItemDetail.primary_transaksi_asal.ToString
    '                        rowSR = TBLSR.DefaultView(0).Row
    '                        DocumentNoAsal = rowSR.No_SR + vbCrLf + rowSR.Customer_Name
    '                    ElseIf TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust Then
    '                        'in adjust
    '                        TipeTransaksiAsalSTR = "In Adjust"
    '                        TBLStockInOut.DefaultView.RowFilter = "Primary = " + Me.rowItemDetail.primary_transaksi_asal.ToString
    '                        rowStockInOut = TBLStockInOut.DefaultView(0).Row
    '                        DocumentNoAsal = rowStockInOut.No_Transaksi + " (" + rowStockInOut.Tipe_Adjustment + ")"
    '                    ElseIf TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InStockTakeAdjust Then
    '                        TipeTransaksiAsalSTR = "In Stock Take"
    '                        TBLStockTake.DefaultView.RowFilter = "Primary = " + Me.rowItemDetail.primary_transaksi_asal.ToString
    '                        rowStockTake = TBLStockTake.DefaultView(0).Row
    '                        DocumentNoAsal = rowStockTake.No_Transaksi
    '                    End If

    '                    If TBL.Rows.Count > 0 Then
    '                        rowBefore = TBL.Rows(TBL.Rows.Count - 1)
    '                    End If
    '                    Me.row = TBL.NewRow
    '                    With row
    '                        .Amount = Me.rowItemDetail.amount
    '                        .No_Dokumen_Asal = DocumentNoAsal
    '                        .Qty = Me.rowItemDetail.qty
    '                        .Tanggal = Me.rowItemDetail.tgl_transaksi_asal
    '                        .Tipe_Item_Int = Me.rowItemDetail.tipe_item_detail
    '                        .Tipe_Transaksi_Asal_Str = TipeTransaksiAsalSTR
    '                        .Tipe_Transaksi_Int = 0
    '                        .Tipe_Item_Str = "Begining Balance"

    '                        If TBL.Rows.Count > 0 Then
    '                            .Amount_Balance = rowBefore.Amount + .Amount
    '                            .Qty_Balance = rowBefore.Qty + .Qty
    '                        Else
    '                            .Amount_Balance = .Amount
    '                            .Qty_Balance = .Qty
    '                        End If
    '                    End With
    '                    TBL.Rows.Add(Me.row)
    '                Next
    '            End If
    '        End If

    '        'isikan data in
    '        'InPurchase
    '        TglSampai = Me.ObjCD.GetLasDateOfTheMonth(Me.Period)
    '        TglDari = Me.ObjCD.GetFirstDateofTheMonth(Me.Period)
    '        DTAAOP.Fill(TBLAOP, TglSampai, TglDari)
    '        TipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
    '        TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InPurchase
    '        Me.TBLItemDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(TipeItemDetail1).ToString + _
    '        " AND [primary code no] = " + Me.vPrimaryProduct.ToString + _
    '        " AND [tipe transaksi asal] = " + CInt(TipeTransaksiAsal).ToString
    '        Me.TBLItemDetail.DefaultView.Sort = "[tgl transaksi asal], [tipe transaksi asal], [primary transaksi asal]"
    '        For Each DV In Me.TBLItemDetail.DefaultView
    '            Me.rowItemDetail = DV.Row
    '            Me.rowBefore = Me.TBL.Rows(Me.TBL.Rows.Count - 1)
    '            TBLAOP.DefaultView.RowFilter = "Primary = " + Me.rowItemDetail.primary_transaksi_asal.ToString
    '            rowAOP = TBLAOP.DefaultView(0).Row

    '            Me.row = TBL.NewRow
    '            With row
    '                .Amount = Me.rowItemDetail.amount
    '                .Amount_Balance = 0
    '                If Me.chkUseSupplierInvoiceNumber.Checked = True Then
    '                    If rowAOP.Supplier_Invoice_No.Trim <> "" Then
    '                        .No_Dokumen_Asal = rowAOP.Supplier_Invoice_No + vbCrLf + rowAOP.Supplier_Name
    '                    Else
    '                        .No_Dokumen_Asal = rowAOP.AOP_Number + vbCrLf + rowAOP.Supplier_Name
    '                    End If
    '                Else
    '                    .No_Dokumen_Asal = rowAOP.AOP_Number + vbCrLf + rowAOP.Supplier_Name
    '                End If
    '                .Qty = Me.rowItemDetail.qty
    '                .Qty_Balance = 0
    '                .Tanggal = Me.rowItemDetail.tgl_transaksi_asal
    '                .Tipe_Item_Int = CInt(TipeItemDetail1)
    '                .Tipe_Transaksi_Asal_Str = "In Purchase"
    '                .Tipe_Transaksi_Int = CInt(TipeTransaksiAsal)
    '                .Tipe_Item_Str = ""
    '            End With
    '            TBL.Rows.Add(Me.row)
    '        Next

    '        'InSR
    '        DTASR.Fill(TBLSR, TglSampai, TglDari)
    '        TipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
    '        TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR
    '        Me.TBLItemDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(TipeItemDetail1).ToString + _
    '        " AND [primary code no] = " + Me.vPrimaryProduct.ToString + _
    '        " AND [tipe transaksi asal] = " + CInt(TipeTransaksiAsal).ToString
    '        Me.TBLItemDetail.DefaultView.Sort = "[tgl transaksi asal], [tipe transaksi asal], [primary transaksi asal]"
    '        For Each DV In Me.TBLItemDetail.DefaultView
    '            Me.rowItemDetail = DV.Row
    '            Me.rowBefore = Me.TBL.Rows(Me.TBL.Rows.Count - 1)
    '            TBLSR.DefaultView.RowFilter = "[Primary Stock In Out] = " + Me.rowItemDetail.primary_transaksi_asal.ToString
    '            rowSR = TBLSR.DefaultView(0).Row

    '            Me.row = TBL.NewRow
    '            With row
    '                .Amount = Me.rowItemDetail.amount
    '                .Amount_Balance = 0
    '                .No_Dokumen_Asal = rowSR.No_SR + vbCrLf + rowSR.Customer_Name
    '                .Qty = Me.rowItemDetail.qty
    '                .Qty_Balance = 0
    '                .Tanggal = Me.rowItemDetail.tgl_transaksi_asal
    '                .Tipe_Item_Int = CInt(TipeItemDetail1)
    '                .Tipe_Transaksi_Asal_Str = "In Sales Return"
    '                .Tipe_Transaksi_Int = CInt(TipeTransaksiAsal)
    '                .Tipe_Item_Str = ""
    '            End With
    '            TBL.Rows.Add(Me.row)
    '        Next
    '        'InBundling
    '        DTABundling.Fill(TBLBundling, TglSampai, TglDari)
    '        TipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
    '        TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling
    '        Me.TBLItemDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(TipeItemDetail1).ToString + _
    '        " AND [primary code no] = " + Me.vPrimaryProduct.ToString + _
    '        " AND [tipe transaksi asal] = " + CInt(TipeTransaksiAsal).ToString
    '        Me.TBLItemDetail.DefaultView.Sort = "[tgl transaksi asal], [tipe transaksi asal], [primary transaksi asal]"
    '        For Each DV In Me.TBLItemDetail.DefaultView
    '            Me.rowItemDetail = DV.Row
    '            Me.rowBefore = Me.TBL.Rows(Me.TBL.Rows.Count - 1)
    '            TBLBundling.DefaultView.RowFilter = "Primary = " + Me.rowItemDetail.primary_transaksi_asal.ToString
    '            rowBundling = TBLBundling.DefaultView(0).Row

    '            Me.row = TBL.NewRow
    '            With row
    '                .Amount = Me.rowItemDetail.amount
    '                .Amount_Balance = 0
    '                .No_Dokumen_Asal = rowBundling.No_Transaksi
    '                .Qty = Me.rowItemDetail.qty
    '                .Qty_Balance = 0
    '                .Tanggal = Me.rowItemDetail.tgl_transaksi_asal
    '                .Tipe_Item_Int = CInt(TipeItemDetail1)
    '                .Tipe_Transaksi_Asal_Str = "In Bundling"
    '                .Tipe_Transaksi_Int = CInt(TipeTransaksiAsal)
    '                .Tipe_Item_Str = ""
    '            End With
    '            TBL.Rows.Add(Me.row)
    '        Next
    '        'InAdjust
    '        TipeStockInOut = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust
    '        DTAStockInOut.Fill(TBLStockInOut, TipeStockInOut, TglSampai, TglDari)
    '        TipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
    '        TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust
    '        Me.TBLItemDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(TipeItemDetail1).ToString + _
    '        " AND [primary code no] = " + Me.vPrimaryProduct.ToString + _
    '        " AND [tipe transaksi asal] = " + CInt(TipeTransaksiAsal).ToString
    '        Me.TBLItemDetail.DefaultView.Sort = "[tgl transaksi asal], [tipe transaksi asal], [primary transaksi asal]"
    '        For Each DV In Me.TBLItemDetail.DefaultView
    '            Me.rowItemDetail = DV.Row
    '            Me.rowBefore = Me.TBL.Rows(Me.TBL.Rows.Count - 1)
    '            TBLStockInOut.DefaultView.RowFilter = "Primary = " + Me.rowItemDetail.primary_transaksi_asal.ToString
    '            rowStockInOut = TBLStockInOut.DefaultView(0).Row

    '            Me.row = TBL.NewRow
    '            With row
    '                .Amount = Me.rowItemDetail.amount
    '                .Amount_Balance = 0
    '                .No_Dokumen_Asal = rowStockInOut.No_Transaksi + " (" + rowStockInOut.Tipe_Adjustment + ")"
    '                .Qty = Me.rowItemDetail.qty
    '                .Qty_Balance = 0
    '                .Tanggal = Me.rowItemDetail.tgl_transaksi_asal
    '                .Tipe_Item_Int = CInt(TipeItemDetail1)
    '                .Tipe_Transaksi_Asal_Str = "In Adjust"
    '                .Tipe_Transaksi_Int = CInt(TipeTransaksiAsal)
    '                .Tipe_Item_Str = ""
    '            End With
    '            TBL.Rows.Add(Me.row)
    '        Next
    '        'InStockTakeAdjust
    '        DTAStockTake.Fill(TBLStockTake)
    '        TipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
    '        TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InStockTakeAdjust
    '        Me.TBLItemDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(TipeItemDetail1).ToString + _
    '        " AND [primary code no] = " + Me.vPrimaryProduct.ToString + _
    '        " AND [tipe transaksi asal] = " + CInt(TipeTransaksiAsal).ToString
    '        Me.TBLItemDetail.DefaultView.Sort = "[tgl transaksi asal], [tipe transaksi asal], [primary transaksi asal]"
    '        For Each DV In Me.TBLItemDetail.DefaultView
    '            Me.rowItemDetail = DV.Row
    '            Me.rowBefore = Me.TBL.Rows(Me.TBL.Rows.Count - 1)
    '            TBLStockTake.DefaultView.RowFilter = "Primary = " + Me.rowItemDetail.primary_transaksi_asal.ToString
    '            rowStockTake = TBLStockTake.DefaultView(0).Row

    '            Me.row = TBL.NewRow
    '            With row
    '                .Amount = Me.rowItemDetail.amount
    '                .Amount_Balance = 0
    '                .No_Dokumen_Asal = rowStockTake.No_Transaksi
    '                .Qty = Me.rowItemDetail.qty
    '                .Qty_Balance = 0
    '                .Tanggal = Me.rowItemDetail.tgl_transaksi_asal
    '                .Tipe_Item_Int = CInt(TipeItemDetail1)
    '                .Tipe_Transaksi_Asal_Str = "In Stock Take"
    '                .Tipe_Transaksi_Int = CInt(TipeTransaksiAsal)
    '                .Tipe_Item_Str = ""
    '            End With
    '            TBL.Rows.Add(Me.row)
    '        Next

    '        'isikan data out
    '        'OutSales
    '        DTAInvoice.Fill(TBLInvoice, TglSampai, TglDari)
    '        TipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
    '        TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutSales
    '        Me.TBLItemDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(TipeItemDetail1).ToString + _
    '        " AND [primary code no] = " + Me.vPrimaryProduct.ToString + _
    '        " AND [tipe transaksi asal] = " + CInt(TipeTransaksiAsal).ToString
    '        Me.TBLItemDetail.DefaultView.Sort = "[tgl transaksi asal], [tipe transaksi asal], [primary transaksi asal]"
    '        'dianggap in lagi supaya tampilan pada laporan lebih frienly
    '        If vCostingMethod <> ClsTransaksiMonthlyClosingItem.EnumCostingMethod.Average Then
    '            TipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
    '        End If
    '        For Each DV In Me.TBLItemDetail.DefaultView
    '            Me.rowItemDetail = DV.Row
    '            Me.rowBefore = Me.TBL.Rows(Me.TBL.Rows.Count - 1)
    '            TBLInvoice.DefaultView.RowFilter = "Primary = " + Me.rowItemDetail.primary_transaksi_asal.ToString
    '            rowInvoice = TBLInvoice.DefaultView(0).Row

    '            Me.row = TBL.NewRow
    '            With row
    '                .Amount = Me.rowItemDetail.amount * -1
    '                .Amount_Balance = 0
    '                .No_Dokumen_Asal = rowInvoice.No_Transaksi + vbCrLf + rowInvoice.Customer_Name
    '                .Qty = Me.rowItemDetail.qty * -1
    '                .Qty_Balance = 0
    '                .Tanggal = Me.rowItemDetail.tgl_transaksi_asal
    '                .Tipe_Item_Int = CInt(TipeItemDetail1)
    '                .Tipe_Transaksi_Asal_Str = "Out Sales"
    '                .Tipe_Item_Str = ""
    '                .Tipe_Transaksi_Int = CInt(TipeTransaksiAsal)
    '            End With
    '            TBL.Rows.Add(Me.row)
    '        Next
    '        'OutBundling
    '        TipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
    '        TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling
    '        Me.TBLItemDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(TipeItemDetail1).ToString + _
    '        " AND [primary code no] = " + Me.vPrimaryProduct.ToString + _
    '        " AND [tipe transaksi asal] = " + CInt(TipeTransaksiAsal).ToString
    '        Me.TBLItemDetail.DefaultView.Sort = "[tgl transaksi asal], [tipe transaksi asal], [primary transaksi asal]"
    '        'dianggap in lagi supaya tampilan pada laporan lebih frienly
    '        If vCostingMethod <> ClsTransaksiMonthlyClosingItem.EnumCostingMethod.Average Then
    '            TipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
    '        End If
    '        For Each DV In Me.TBLItemDetail.DefaultView
    '            Me.rowItemDetail = DV.Row
    '            Me.rowBefore = Me.TBL.Rows(Me.TBL.Rows.Count - 1)
    '            TBLBundling.DefaultView.RowFilter = "Primary = " + Me.rowItemDetail.primary_transaksi_asal.ToString
    '            rowBundling = TBLBundling.DefaultView(0).Row

    '            Me.row = TBL.NewRow
    '            With row
    '                .Amount = Me.rowItemDetail.amount * -1
    '                .Amount_Balance = 0
    '                .No_Dokumen_Asal = rowBundling.No_Transaksi
    '                .Qty = Me.rowItemDetail.qty * -1
    '                .Qty_Balance = 0
    '                .Tanggal = Me.rowItemDetail.tgl_transaksi_asal
    '                .Tipe_Item_Int = CInt(TipeItemDetail1)
    '                .Tipe_Transaksi_Asal_Str = "Out Bundling"
    '                .Tipe_Transaksi_Int = CInt(TipeTransaksiAsal)
    '                .Tipe_Item_Str = ""
    '            End With
    '            TBL.Rows.Add(Me.row)
    '        Next
    '        'OutPR
    '        TipePR = ClsTransaksiPurchaseReturn.EnumTipe_PurchaseReturn.Nota_Retur
    '        DTAPR.Fill(TBLPR, TglSampai, TipePR, TglDari)
    '        TipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
    '        TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutPR
    '        Me.TBLItemDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(TipeItemDetail1).ToString + _
    '        " AND [primary code no] = " + Me.vPrimaryProduct.ToString + _
    '        " AND [tipe transaksi asal] = " + CInt(TipeTransaksiAsal).ToString
    '        Me.TBLItemDetail.DefaultView.Sort = "[tgl transaksi asal], [tipe transaksi asal], [primary transaksi asal]"
    '        'dianggap in lagi supaya tampilan pada laporan lebih frienly
    '        If vCostingMethod <> ClsTransaksiMonthlyClosingItem.EnumCostingMethod.Average Then
    '            TipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
    '        End If
    '        For Each DV In Me.TBLItemDetail.DefaultView
    '            Me.rowItemDetail = DV.Row
    '            Me.rowBefore = Me.TBL.Rows(Me.TBL.Rows.Count - 1)
    '            TBLPR.DefaultView.RowFilter = "[Primary Transaksi] = " + Me.rowItemDetail.primary_transaksi_asal.ToString
    '            rowPR = TBLPR.DefaultView(0).Row

    '            Me.row = TBL.NewRow
    '            With row
    '                .Amount = Me.rowItemDetail.amount * -1
    '                .Amount_Balance = 0
    '                .No_Dokumen_Asal = rowPR.No_Transaksi + vbCrLf + rowPR.Supplier_Name
    '                .Qty = Me.rowItemDetail.qty * -1
    '                .Qty_Balance = 0
    '                .Tanggal = Me.rowItemDetail.tgl_transaksi_asal
    '                .Tipe_Item_Int = CInt(TipeItemDetail1)
    '                .Tipe_Transaksi_Asal_Str = "Out Purchase Return"
    '                .Tipe_Transaksi_Int = CInt(TipeTransaksiAsal)
    '                .Tipe_Item_Str = ""
    '            End With
    '            TBL.Rows.Add(Me.row)
    '        Next
    '        'OutStockTakeAdjust
    '        TipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
    '        TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutStockTakeAdjust
    '        Me.TBLItemDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(TipeItemDetail1).ToString + _
    '        " AND [primary code no] = " + Me.vPrimaryProduct.ToString + _
    '        " AND [tipe transaksi asal] = " + CInt(TipeTransaksiAsal).ToString
    '        Me.TBLItemDetail.DefaultView.Sort = "[tgl transaksi asal], [tipe transaksi asal], [primary transaksi asal]"
    '        'dianggap in lagi supaya tampilan pada laporan lebih frienly
    '        If vCostingMethod <> ClsTransaksiMonthlyClosingItem.EnumCostingMethod.Average Then
    '            TipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
    '        End If
    '        For Each DV In Me.TBLItemDetail.DefaultView
    '            Me.rowItemDetail = DV.Row
    '            Me.rowBefore = Me.TBL.Rows(Me.TBL.Rows.Count - 1)
    '            TBLStockTake.DefaultView.RowFilter = "Primary = " + Me.rowItemDetail.primary_transaksi_asal.ToString
    '            rowStockTake = TBLStockTake.DefaultView(0).Row

    '            Me.row = TBL.NewRow
    '            With row
    '                .Amount = Me.rowItemDetail.amount * -1
    '                .Amount_Balance = 0
    '                .No_Dokumen_Asal = rowStockTake.No_Transaksi
    '                .Qty = Me.rowItemDetail.qty * -1
    '                .Qty_Balance = 0
    '                .Tanggal = Me.rowItemDetail.tgl_transaksi_asal
    '                .Tipe_Item_Int = CInt(TipeItemDetail1)
    '                .Tipe_Transaksi_Asal_Str = "Out Stock Take"
    '                .Tipe_Transaksi_Int = CInt(TipeTransaksiAsal)
    '                .Tipe_Item_Str = ""
    '            End With
    '            TBL.Rows.Add(Me.row)
    '        Next
    '        'OutAdjust
    '        TipeStockInOut = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust
    '        DTAStockInOut.Fill(TBLStockInOut, TipeStockInOut, TglSampai, TglDari)
    '        TipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
    '        TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutAdjust
    '        Me.TBLItemDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(TipeItemDetail1).ToString + _
    '        " AND [primary code no] = " + Me.vPrimaryProduct.ToString + _
    '        " AND [tipe transaksi asal] = " + CInt(TipeTransaksiAsal).ToString
    '        Me.TBLItemDetail.DefaultView.Sort = "[tgl transaksi asal], [tipe transaksi asal], [primary transaksi asal]"
    '        'dianggap in lagi supaya tampilan pada laporan lebih frienly
    '        If vCostingMethod <> ClsTransaksiMonthlyClosingItem.EnumCostingMethod.Average Then
    '            TipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
    '        End If
    '        For Each DV In Me.TBLItemDetail.DefaultView
    '            Me.rowItemDetail = DV.Row
    '            Me.rowBefore = Me.TBL.Rows(Me.TBL.Rows.Count - 1)
    '            TBLStockInOut.DefaultView.RowFilter = "Primary = " + Me.rowItemDetail.primary_transaksi_asal.ToString
    '            rowStockInOut = TBLStockInOut.DefaultView(0).Row

    '            Me.row = TBL.NewRow
    '            With row
    '                .Amount = Me.rowItemDetail.amount * -1
    '                .Amount_Balance = rowBefore.Amount_Balance + .Amount
    '                .No_Dokumen_Asal = rowStockInOut.No_Transaksi + " (" + rowStockInOut.Tipe_Adjustment + ")"
    '                .Qty = Me.rowItemDetail.qty * -1
    '                .Qty_Balance = rowBefore.Qty_Balance + .Qty
    '                .Tanggal = Me.rowItemDetail.tgl_transaksi_asal
    '                .Tipe_Item_Int = CInt(TipeItemDetail1)
    '                .Tipe_Transaksi_Asal_Str = "Out Adjust"
    '                .Tipe_Item_Str = ""
    '                .Tipe_Transaksi_Int = CInt(TipeTransaksiAsal)
    '            End With
    '            TBL.Rows.Add(Me.row)
    '        Next

    '        If Me.chkShowDetailBegining.Checked = True Then
    '            TipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.EndingBalanceItem
    '            Me.TBLItemDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(TipeItemDetail1).ToString + " AND [primary code no] = " + Me.vPrimaryProduct.ToString
    '            Me.TBLItemDetail.DefaultView.Sort = "[tgl transaksi asal], [tipe transaksi asal], [primary transaksi asal]"
    '            If Me.TBLItemDetail.DefaultView.Count > 0 Then
    '                'isikan data awal untuk menentukan no transaksi
    '                TglDariPurchase = Date.MaxValue
    '                TglDariInAdjust = Date.MaxValue
    '                TglDariInSR = Date.MaxValue
    '                TglDariInBundling = Date.MaxValue
    '                For Each DV In Me.TBLItemDetail.DefaultView
    '                    Me.rowItemDetail = DV.Row
    '                    TipeTransaksiAsal = Me.rowItemDetail.tipe_transaksi_asal
    '                    If TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InPurchase Then
    '                        'in purchase
    '                        If TglDariPurchase > Me.rowItemDetail.tgl_transaksi_asal Then
    '                            TglDariPurchase = Me.rowItemDetail.tgl_transaksi_asal
    '                        End If
    '                    ElseIf TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling Then
    '                        'in bundling
    '                        If TglDariInBundling > Me.rowItemDetail.tgl_transaksi_asal Then
    '                            TglDariInBundling = Me.rowItemDetail.tgl_transaksi_asal
    '                        End If
    '                    ElseIf TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR Then
    '                        'in sr
    '                        If TglDariInSR > Me.rowItemDetail.tgl_transaksi_asal Then
    '                            TglDariInSR = Me.rowItemDetail.tgl_transaksi_asal
    '                        End If
    '                    ElseIf TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust Then
    '                        'in adjust
    '                        If TglDariInAdjust > Me.rowItemDetail.tgl_transaksi_asal Then
    '                            TglDariInAdjust = Me.rowItemDetail.tgl_transaksi_asal
    '                        End If
    '                    End If
    '                Next
    '                TBLAOP.Rows.Clear()
    '                If TglDariPurchase <> Date.MaxValue Then
    '                    DTAAOP.Fill(TBLAOP, TglSampai, TglDariPurchase)
    '                End If
    '                TBLStockInOut.Rows.Clear()
    '                If TglDariInAdjust <> Date.MaxValue Then
    '                    TipeStockInOut = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust
    '                    DTAStockInOut.Fill(TBLStockInOut, TipeStockInOut, TglSampai, TglDariInAdjust)
    '                End If
    '                TBLSR.Rows.Clear()
    '                If TglDariInSR <> Date.MaxValue Then
    '                    DTASR.Fill(TBLSR, TglSampai, TglDariInSR)
    '                End If
    '                TBLBundling.Rows.Clear()
    '                If TglDariInBundling <> Date.MaxValue Then
    '                    DTABundling.Fill(TBLBundling, TglSampai, TglDariInBundling)
    '                End If
    '                TBLStockTake.Rows.Clear()
    '                DTAStockTake.Fill(TBLStockTake)
    '                For Each DV In Me.TBLItemDetail.DefaultView
    '                    Me.rowItemDetail = DV.Row
    '                    TipeTransaksiAsal = Me.rowItemDetail.tipe_transaksi_asal
    '                    If TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InPurchase Then
    '                        'in purchase
    '                        TipeTransaksiAsalSTR = "In Purchase"
    '                        TBLAOP.DefaultView.RowFilter = "Primary = " + Me.rowItemDetail.primary_transaksi_asal.ToString
    '                        rowAOP = TBLAOP.DefaultView(0).Row
    '                        If Me.chkUseSupplierInvoiceNumber.Checked = True Then
    '                            If rowAOP.Supplier_Invoice_No.Trim <> "" Then
    '                                DocumentNoAsal = rowAOP.Supplier_Invoice_No + vbCrLf + rowAOP.Supplier_Name
    '                            Else
    '                                DocumentNoAsal = rowAOP.AOP_Number + vbCrLf + rowAOP.Supplier_Name
    '                            End If
    '                        Else
    '                            DocumentNoAsal = rowAOP.AOP_Number + vbCrLf + rowAOP.Supplier_Name
    '                        End If
    '                    ElseIf TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling Then
    '                        'in bundling
    '                        TipeTransaksiAsalSTR = "In Bundling"
    '                        TBLBundling.DefaultView.RowFilter = "Primary = " + Me.rowItemDetail.primary_transaksi_asal.ToString
    '                        rowBundling = TBLBundling.DefaultView(0).Row
    '                        DocumentNoAsal = rowBundling.No_Transaksi
    '                    ElseIf TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR Then
    '                        'in sr
    '                        TipeTransaksiAsalSTR = "In Sales Return"
    '                        TBLSR.DefaultView.RowFilter = "[Primary Stock In Out] = " + Me.rowItemDetail.primary_transaksi_asal.ToString
    '                        rowSR = TBLSR.DefaultView(0).Row
    '                        DocumentNoAsal = rowSR.No_SR + vbCrLf + rowSR.Customer_Name
    '                    ElseIf TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust Then
    '                        'in adjust
    '                        TipeTransaksiAsalSTR = "In Adjust"
    '                        TBLStockInOut.DefaultView.RowFilter = "Primary = " + Me.rowItemDetail.primary_transaksi_asal.ToString
    '                        rowStockInOut = TBLStockInOut.DefaultView(0).Row
    '                        DocumentNoAsal = rowStockInOut.No_Transaksi + " (" + rowStockInOut.Tipe_Adjustment + ")"
    '                    ElseIf TipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InStockTakeAdjust Then
    '                        TipeTransaksiAsalSTR = "In Stock Take"
    '                        TBLStockTake.DefaultView.RowFilter = "Primary = " + Me.rowItemDetail.primary_transaksi_asal.ToString
    '                        rowStockTake = TBLStockTake.DefaultView(0).Row
    '                        DocumentNoAsal = rowStockTake.No_Transaksi
    '                    End If

    '                    If TBL.Rows.Count > 0 Then
    '                        rowBefore = TBL.Rows(TBL.Rows.Count - 1)
    '                    End If
    '                    Me.row = TBL.NewRow
    '                    With row
    '                        .Amount = Me.rowItemDetail.amount * -1
    '                        .No_Dokumen_Asal = DocumentNoAsal
    '                        .Qty = Me.rowItemDetail.qty * -1
    '                        .Tanggal = Me.rowItemDetail.tgl_transaksi_asal
    '                        .Tipe_Item_Int = Me.rowItemDetail.tipe_item_detail
    '                        .Tipe_Transaksi_Asal_Str = TipeTransaksiAsalSTR
    '                        .Tipe_Transaksi_Int = TipeTransaksiAsal
    '                        .Tipe_Item_Str = "Ending Balance"

    '                        If TBL.Rows.Count > 0 Then
    '                            .Amount_Balance = rowBefore.Amount + .Amount
    '                            .Qty_Balance = rowBefore.Qty + .Qty
    '                        Else
    '                            .Amount_Balance = .Amount
    '                            .Qty_Balance = .Qty
    '                        End If
    '                    End With
    '                    TBL.Rows.Add(Me.row)
    '                Next
    '            End If
    '        End If

    '        'urutkan ulang
    '        TBLDummy.Rows.Clear()
    '        For Each row In TBL.Rows
    '            rowDummy = TBLDummy.NewTblDetailCalc1Row
    '            For ij As Integer = 0 To TBLDummy.Columns.Count - 1
    '                rowDummy.Item(ij) = row.Item(ij)
    '            Next
    '            TBLDummy.Rows.Add(rowDummy)
    '        Next
    '        TBLDummy.DefaultView.Sort = "[Tipe Item Int], Tanggal, [Tipe Transaksi Int]"
    '        TBLDummy.DefaultView.RowFilter = ""
    '        TBL.Rows.Clear()
    '        For Each DV In TBLDummy.DefaultView
    '            rowDummy = DV.Row
    '            row = TBL.NewTblDetailCalc1Row
    '            For ij As Integer = 0 To TBLDummy.Columns.Count - 1
    '                row.Item(ij) = rowDummy.Item(ij)
    '            Next
    '            TBL.Rows.Add(row)
    '        Next
    '        'hitung ulang balance
    '        If TBL.Rows.Count > 1 Then
    '            For ij As Double = 1 To Me.TBL.Rows.Count - 1
    '                rowBefore = Me.TBL.Rows(ij - 1)
    '                row = TBL.Rows(ij)
    '                With row
    '                    .Qty_Balance = rowBefore.Qty_Balance + .Qty
    '                    .Amount_Balance = rowBefore.Amount_Balance + .Amount
    '                End With
    '            Next
    '        End If

    '        'show report
    '        Me.Report1.SetDataSource(Me.DS)
    '        Me.Report1.SetParameterValue(Me.Report1.Parameter_DefPrecAmount.ParameterFieldName, Me.ObjCur.GetDecimal_DefaultCurrency)
    '        Me.Report1.SetParameterValue(Me.Report1.Parameter_DefPrecQty.ParameterFieldName, Me.ObjProd.GetAngkaDecimal(Me.vPrimaryProduct))
    '        Me.Report1.SetParameterValue(Me.Report1.Parameter_Period.ParameterFieldName, Me.Period)
    '        Me.Report1.SetParameterValue(Me.Report1.Parameter_Product.ParameterFieldName, Me.ObjProd.GetProductCode(Me.vPrimaryProduct) + " - " + Me.ObjProd.GetDescription(Me.vPrimaryProduct))
    '        Me.Report1.SetParameterValue(Me.Report1.Parameter_Unit.ParameterFieldName, Me.ObjProd.GetUnitCode(Me.vPrimaryProduct))
    '        If ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, _
    '                       Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets) = "" Then
    '            Negatif = 1
    '        Else
    '            Negatif = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, _
    '            Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets)
    '        End If
    '        Me.Report1.SetParameterValue(Me.Report1.Parameter_UseBracket.ParameterFieldName, Negatif)
    '        Me.Report1.SetParameterValue(Me.Report1.Parameter_CrpTransaksiMonthly_DetailCalc_Rekap_DP_Qty.ParameterFieldName, Me.ObjProd.GetAngkaDecimal(Me.vPrimaryProduct))
    '        Me.Report1.SetParameterValue(Me.Report1.Parameter_CrpTransaksiMonthly_DetailCalc_Rekap_DP_Amount.ParameterFieldName, Me.ObjCur.GetDecimal_DefaultCurrency)
    '        Me.Report1.SetParameterValue(Me.Report1.Parameter_CrpTransaksiMonthly_DetailCalc_Rekap_SubReport_Negative.ParameterFieldName, Negatif)

    '        Me.crv1.ReportSource = Me.Report1
    '        Me.crv1.DisplayGroupTree = False
    '        Me.crv1.DisplayStatusBar = True
    '        Me.crv1.DisplayToolbar = True
    '        Me.crv1.ShowGroupTreeButton = False
    '        Me.crv1.ShowExportButton = True
    '        Me.crv1.ShowRefreshButton = False
    '        Me.crv1.Show()

    '        Me.chkShowDetailBegining.Enabled = True
    '        Me.chkUseSupplierInvoiceNumber.Enabled = True
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        Windows.Forms.Cursor.Current = Cursors.Arrow
    '        DTAAOP.Dispose()
    '        TBLAOP.Dispose()
    '        DTASR.Dispose()
    '        TBLSR.Dispose()
    '        DTABundling.Dispose()
    '        TBLBundling.Dispose()
    '        DTAStockInOut.Dispose()
    '        TBLStockInOut.Dispose()
    '        DTAStockTake.Dispose()
    '        TBLStockTake.Dispose()
    '        DTAInvoice.Dispose()
    '        TBLInvoice.Dispose()
    '        DTAPR.Dispose()
    '        TBLPR.Dispose()
    '    End Try
    'End Sub

    Private Sub FrmTransaksiMonthlyClosing_DetailCalc2_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.chkUseSupplierInvoiceNumber.Checked = True Then
                Me.ObjSettingUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.MonthlyInventoryClosing, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.MonthlyInventoryClosing_ShowSupplierInvoiceNumber, _
                "1", Me.UserCode)
            Else
                Me.ObjSettingUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.MonthlyInventoryClosing, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.MonthlyInventoryClosing_ShowSupplierInvoiceNumber, _
                "0", Me.UserCode)
            End If
            If Me.chkShowDetailBegining.Checked = True Then
                Me.ObjSettingUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.MonthlyInventoryClosing, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.MonthlyInventoryClosing_ShowDetailBegining, _
                "1", Me.UserCode)
            Else
                Me.ObjSettingUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.MonthlyInventoryClosing, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.MonthlyInventoryClosing_ShowDetailBegining, _
                "0", Me.UserCode)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiMonthlyClosing_DetailCalc2_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim DummyCostingMethod As String = ""
        Dim ShowSupplierInvStr As String = ""
        Dim ShowDetailStr As String = ""

        Try
            ShowSupplierInvStr = Me.ObjSettingUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.MonthlyInventoryClosing, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.MonthlyInventoryClosing_ShowSupplierInvoiceNumber, Me.UserCode)
            If ShowSupplierInvStr = "" Then
                Me.chkUseSupplierInvoiceNumber.Checked = False
            Else
                If ShowSupplierInvStr = "1" Then
                    Me.chkUseSupplierInvoiceNumber.Checked = True
                Else
                    Me.chkUseSupplierInvoiceNumber.Checked = False
                End If
            End If
            ShowDetailStr = Me.ObjSettingUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.MonthlyInventoryClosing, _
             Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.MonthlyInventoryClosing_ShowDetailBegining, Me.UserCode)
            If ShowDetailStr = "" Then
                Me.chkShowDetailBegining.Checked = False
            Else
                If ShowDetailStr = "1" Then
                    Me.chkShowDetailBegining.Checked = True
                Else
                    Me.chkShowDetailBegining.Checked = False
                End If
            End If
            Me.crv1.EnableDrillDown = False
            DummyCostingMethod = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
          Core_Module.ClsTransaksiSetting.EnumNamaParameter.CostingMethod)
            If DummyCostingMethod = "" Then
                CostingMethod = ClsTransaksiMonthlyClosingItem.EnumCostingMethod.Average
            Else
                CostingMethod = DummyCostingMethod
            End If

            'Me.ShowInTaskbar = False
            Me.Icon = My.Resources.desktop_icon_3_fw
            Me.ShowDetailCalc(CostingMethod)
            If CostingMethod = ClsTransaksiMonthlyClosingItem.EnumCostingMethod.Average Then
                Me.chkShowDetailBegining.Visible = False
                Me.chkShowDetailBegining.TabStop = False
            Else
                Me.chkShowDetailBegining.Visible = True
                Me.chkShowDetailBegining.TabStop = True
            End If
            Aktif = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkUseSupplierInvoiceNumber_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkUseSupplierInvoiceNumber.CheckedChanged
        Try
            If Me.Aktif = True Then
                Me.ShowDetailCalc(Me.CostingMethod)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkShowDetailBegining_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowDetailBegining.CheckedChanged
        Try
            If Me.Aktif = True Then
                Me.ShowDetailCalc(Me.CostingMethod)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class