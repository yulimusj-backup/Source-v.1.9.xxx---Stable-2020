Option Explicit On
Option Strict Off

Public Class ClsMonthlyClosing01
    Public ClosingPeriod As Date = #12/1/2015#
    Public SourceConstring As String = "server=192.168.1.152;user id=root;Password=%168JCbcmsi7%;persist security info=True;database=core_db;port=3307;default command timeout=19000000"
    Public TargetConstring As String = "server=192.168.1.152;user id=root;Password=%168JCbcmsi7%;persist security info=True;database=core_db;port=3307;default command timeout=19000000"
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjCur As New Core_Module.clsMasterCurrency

    Public Sub StartCutOffClosingPertama(ByVal vClosingPeriod As Date, ByVal vTargetConstring As String, ByVal vSumberConstring As String, ByVal vPrg As ProgressBar, ByVal vLbl As Label)
        Dim PrimaryClosingTarget As Double = -1
        Dim DTAHeaderTarget As New DseClosingCutOff01TableAdapters.tbltrans_purchase_closingTableAdapter
        Dim DTAItemTarget As New DseClosingCutOff01TableAdapters.tbltrans_purchase_closing_itemTableAdapter
        Dim DTAItemAkun As New DseClosingCutOff01TableAdapters.tbltrans_purchase_closing_item_akunTableAdapter
        Dim DTAItemAdjust As New DseClosingCutOff01TableAdapters.tbltrans_purchase_closing_item_inadjustmentTableAdapter
        Dim DTAItemDetail As New DseClosingCutOff01TableAdapters.tbltrans_purchase_closing_item_detailTableAdapter
        Dim TBLItemDetailAdjust As New DseClosingCutOff01.tbltrans_purchase_closing_item_inadjustmentDataTable
        Dim TBLStockInAdjust As New DseClosingCutOff01.tbltransaksi_stock_inoutDataTable

        Try
            'tentukan closing
            Me.ClosingPeriod = vClosingPeriod
            Me.TargetConstring = vTargetConstring
            Me.SourceConstring = vSumberConstring

            vPrg.Value = 0
            vLbl.Text = ""

            DTAHeaderTarget.Connection.ConnectionString = Me.TargetConstring
            DTAItemTarget.Connection.ConnectionString = Me.TargetConstring
            DTAItemAkun.Connection.ConnectionString = Me.TargetConstring
            DTAItemAdjust.Connection.ConnectionString = Me.TargetConstring
            DTAItemDetail.Connection.ConnectionString = Me.TargetConstring


            'If MsgBox("Re-Transfer monthly closing?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
            '    'hapus di database target
            '    Me.BersihkanTabelTarget()

            '    'transfer dari source ke target
            '    Me.TransferKeTarget()
            '    'Exit Sub
            'End If

            'get primary closing target
            vLbl.Text = "Preparing for first closing data modification..."
            PrimaryClosingTarget = DTAHeaderTarget.ScalarQuery_GetPrimaryTransaksi2(Me.ClosingPeriod).Value

            'update primary jurnal
            DTAHeaderTarget.UpdateQuery_SetPrimaryJurnalMinus(PrimaryClosingTarget)

            'beresin item 
            '1. qty begining & qty out =0; amount begining & amount out =0
            DTAItemTarget.UpdateQuery_SetNol(PrimaryClosingTarget)
            '2. qty in = qty out; amount in = amount end
            DTAItemTarget.UpdateQuery_SetQtyAmountIn(PrimaryClosingTarget)

            'hapus item akun
            DTAItemAkun.Delete(PrimaryClosingTarget)

            'hapus item in adjust
            DTAItemAdjust.Delete(PrimaryClosingTarget)

            'hapus item detail (begining balance, out, in)
            DTAItemDetail.DeleteQuery1(PrimaryClosingTarget, CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.BeginingBalanceItem))
            DTAItemDetail.DeleteQuery1(PrimaryClosingTarget, CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem))
            DTAItemDetail.DeleteQuery1(PrimaryClosingTarget, CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem))

            'isikan tabel stock in adjust, item in adjust
            Me.IsikanItem_DetailIn_Adjust(TBLItemDetailAdjust, TBLStockInAdjust, PrimaryClosingTarget, vPrg, vLbl)

            vPrg.Value = 0
            vLbl.Text = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub StartCutOffClosingPertama_Obselete(ByVal vClosingPeriod As Date, ByVal vTargetConstring As String, ByVal vSumberConstring As String, ByVal vPrg As ProgressBar, ByVal vLbl As Label)
        Dim PrimaryClosingTarget As Double = -1
        Dim DTAHeaderTarget As New DseClosingCutOff01TableAdapters.tbltrans_purchase_closingTableAdapter
        Dim DTAItemTarget As New DseClosingCutOff01TableAdapters.tbltrans_purchase_closing_itemTableAdapter
        Dim DTAItemAkun As New DseClosingCutOff01TableAdapters.tbltrans_purchase_closing_item_akunTableAdapter
        Dim DTAItemAdjust As New DseClosingCutOff01TableAdapters.tbltrans_purchase_closing_item_inadjustmentTableAdapter
        Dim DTAItemDetail As New DseClosingCutOff01TableAdapters.tbltrans_purchase_closing_item_detailTableAdapter
        Dim TBLItemDetailAdjust As New DseClosingCutOff01.tbltrans_purchase_closing_item_inadjustmentDataTable
        Dim TBLStockInAdjust As New DseClosingCutOff01.tbltransaksi_stock_inoutDataTable

        Try
            vPrg.Value = 0
            vLbl.Text = ""

            DTAHeaderTarget.Connection.ConnectionString = Me.TargetConstring
            DTAItemTarget.Connection.ConnectionString = Me.TargetConstring
            DTAItemAkun.Connection.ConnectionString = Me.TargetConstring
            DTAItemAdjust.Connection.ConnectionString = Me.TargetConstring
            DTAItemDetail.Connection.ConnectionString = Me.TargetConstring

            'tentukan closing
            Me.ClosingPeriod = vClosingPeriod
            Me.TargetConstring = vTargetConstring
            Me.SourceConstring = vSumberConstring

            If MsgBox("Re-Transfer monthly closing?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                'hapus di database target
                Me.BersihkanTabelTarget()

                'transfer dari source ke target
                Me.TransferKeTarget()
                'Exit Sub
            End If

            'get primary closing target
            vLbl.Text = "Preparing for first closing data modification..."
            PrimaryClosingTarget = DTAHeaderTarget.ScalarQuery_GetPrimaryTransaksi2(Me.ClosingPeriod).Value

            'update primary jurnal
            DTAHeaderTarget.UpdateQuery_SetPrimaryJurnalMinus(PrimaryClosingTarget)

            'beresin item 
            '1. qty begining & qty out =0; amount begining & amount out =0
            DTAItemTarget.UpdateQuery_SetNol(PrimaryClosingTarget)
            '2. qty in = qty out; amount in = amount end
            DTAItemTarget.UpdateQuery_SetQtyAmountIn(PrimaryClosingTarget)

            'hapus item akun
            DTAItemAkun.Delete(PrimaryClosingTarget)

            'hapus item in adjust
            DTAItemAdjust.Delete(PrimaryClosingTarget)

            'hapus item detail (begining balance, out, in)
            DTAItemDetail.DeleteQuery1(PrimaryClosingTarget, CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.BeginingBalanceItem))
            DTAItemDetail.DeleteQuery1(PrimaryClosingTarget, CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem))
            DTAItemDetail.DeleteQuery1(PrimaryClosingTarget, CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem))

            'isikan tabel stock in adjust, item in adjust
            Me.IsikanItem_DetailIn_Adjust(TBLItemDetailAdjust, TBLStockInAdjust, PrimaryClosingTarget, vPrg, vLbl)

            vPrg.Value = 0
            vLbl.Text = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanItem_DetailIn_Adjust(ByVal TBLItemDetailAdjust As DseClosingCutOff01.tbltrans_purchase_closing_item_inadjustmentDataTable, _
    ByVal TBLStockInAdjust As DseClosingCutOff01.tbltransaksi_stock_inoutDataTable, ByVal vPrimaryClosingTarget As Double, ByVal vPrg As ProgressBar, ByVal vLbl As Label)
        Dim DTAStockInAdjust As New DseClosingCutOff01TableAdapters.tbltransaksi_stock_inoutTableAdapter
        Dim rowStockInAdjust As DseClosingCutOff01.tbltransaksi_stock_inoutRow
        Dim ListPrimaryProduct As New List(Of Double)

        Dim TBLItemDetail As New DseClosingCutOff01.tbltrans_purchase_closing_item_detailDataTable
        Dim rowItemDetail As DseClosingCutOff01.tbltrans_purchase_closing_item_detailRow
        Dim DTAItemDetailTarget As New DseClosingCutOff01TableAdapters.tbltrans_purchase_closing_item_detailTableAdapter
        Dim QtyStockInAdjust, QtyEndClosing As Decimal

        Dim RowItemDetailAdjust As DseClosingCutOff01.tbltrans_purchase_closing_item_inadjustmentRow
        Dim TotalAmountIn As Decimal = 0
        Dim DummyInAmount As Decimal = 0
        Dim DefPrec As Integer = 0

        Dim vDTAItemDetailIn As New DseClosingCutOff01TableAdapters.tbltrans_purchase_closing_item_detailTableAdapter
        Dim vTBLItemDetailIn As New DseClosingCutOff01.tbltrans_purchase_closing_item_detailDataTable
        Dim vRowItemDetailIn As DseClosingCutOff01.tbltrans_purchase_closing_item_detailRow
        Dim DTAItemDetailAdjust As New DseClosingCutOff01TableAdapters.tbltrans_purchase_closing_item_inadjustmentTableAdapter

        Try
            vPrg.Value = 0

            vDTAItemDetailIn.Connection.ConnectionString = Me.TargetConstring
            DTAItemDetailAdjust.Connection.ConnectionString = Me.TargetConstring
            DTAStockInAdjust.Connection.ConnectionString = Me.TargetConstring
            DTAItemDetailTarget.Connection.ConnectionString = Me.TargetConstring

            'hapus item in adjust + item detail in
            DTAItemDetailAdjust.Delete(vPrimaryClosingTarget)
            vDTAItemDetailIn.DeleteQuery1(vPrimaryClosingTarget, CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem))

            DefPrec = Me.ObjCur.GetDecimal_DefaultCurrency
            TBLStockInAdjust.Rows.Clear()
            TBLItemDetail.Rows.Clear()

            DTAStockInAdjust.Connection.ConnectionString = Me.TargetConstring
            DTAItemDetailTarget.Connection.ConnectionString = Me.TargetConstring

            DTAStockInAdjust.Fill(TBLStockInAdjust, _
            CInt(Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust), _
            Me.ObjCD.GetLasDateOfTheMonth(Me.ClosingPeriod))

            vPrg.Value = 0
            vPrg.Maximum = TBLStockInAdjust.Rows.Count
            For Each rowStockInAdjust In TBLStockInAdjust.Rows
                vPrg.Value += 1
                vLbl.Text = "Collecting begining balance - data " + vPrg.Value.ToString + " of " + vPrg.Maximum.ToString

                If ListPrimaryProduct.Contains(rowStockInAdjust.Primary_Product) = False Then
                    ListPrimaryProduct.Add(rowStockInAdjust.Primary_Product)
                End If
            Next

            DTAItemDetailTarget.FillBy_PrimaryTransaksi(TBLItemDetail, vPrimaryClosingTarget, CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.EndingBalanceItem))
            vPrg.Value = 0
            vPrg.Maximum = TBLItemDetail.Rows.Count
            For Each rowItemDetail In TBLItemDetail.Rows
                vPrg.Value += 1
                vLbl.Text = "Collecting begining balance - data " + vPrg.Value.ToString + " of " + vPrg.Maximum.ToString

                If ListPrimaryProduct.Contains(rowItemDetail.primary_code_no) = False Then
                    MsgBox("Logic error, inconsistent product, primary product: " + _
                    rowItemDetail.primary_code_no.ToString, MsgBoxStyle.Information, "Information")
                End If
            Next

            'pastikan qty sama
            vPrg.Value = 0
            vPrg.Maximum = ListPrimaryProduct.Count
            For Each PrimaryCodeNo As Double In ListPrimaryProduct
                vPrg.Value += 1
                vLbl.Text = "Checking begining balance - data " + vPrg.Value.ToString + " of " + vPrg.Maximum.ToString

                QtyEndClosing = 0
                QtyStockInAdjust = 0

                TBLItemDetail.DefaultView.RowFilter = "[primary code no] = " + PrimaryCodeNo.ToString
                For Each DV As DataRowView In TBLItemDetail.DefaultView
                    rowItemDetail = DV.Row
                    QtyEndClosing += rowItemDetail.qty
                Next

                TBLStockInAdjust.DefaultView.RowFilter = "[Primary Product] = " + PrimaryCodeNo.ToString
                For Each DV As DataRowView In TBLStockInAdjust.DefaultView
                    rowStockInAdjust = DV.Row
                    QtyStockInAdjust += rowStockInAdjust.Qty
                Next

                If QtyEndClosing <> QtyStockInAdjust Then
                    MsgBox("Logic error, inconsistent product, primary product: " + _
                   PrimaryCodeNo.ToString, MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            Next

            'tentukan nilai amount untuk setiap item in adjust
            vPrg.Value = 0
            vPrg.Maximum = ListPrimaryProduct.Count
            For Each PrimaryCodeNo As Double In ListPrimaryProduct
                vPrg.Value += 1
                vLbl.Text = "Updating begining balance amount - data " + vPrg.Value.ToString + " of " + vPrg.Maximum.ToString

                TotalAmountIn = 0
                QtyEndClosing = 0
                DummyInAmount = 0

                TBLItemDetail.DefaultView.RowFilter = "[primary code no] = " + PrimaryCodeNo.ToString
                For Each DV As DataRowView In TBLItemDetail.DefaultView
                    rowItemDetail = DV.Row
                    TotalAmountIn += rowItemDetail.amount
                    QtyEndClosing += rowItemDetail.qty
                Next

                TBLStockInAdjust.DefaultView.RowFilter = "[Primary Product] = " + PrimaryCodeNo.ToString
                For ij As Double = 0 To TBLStockInAdjust.DefaultView.Count - 1
                    rowStockInAdjust = TBLStockInAdjust.DefaultView(ij).Row
                    If ij = TBLStockInAdjust.DefaultView.Count - 1 Then
                        rowStockInAdjust.Amount = TotalAmountIn - DummyInAmount
                    Else
                        rowStockInAdjust.Amount = Math.Round((rowStockInAdjust.Qty * TotalAmountIn) / QtyEndClosing, DefPrec, MidpointRounding.AwayFromZero)
                        DummyInAmount += rowStockInAdjust.Amount
                    End If
                    DTAStockInAdjust.UpdateQuery_SetAmount(rowStockInAdjust.Amount, rowStockInAdjust.Primary)
                Next
            Next

            'pindahkan ke item adjustment
            vPrg.Value = 0
            vPrg.Maximum = TBLStockInAdjust.Rows.Count
            TBLItemDetailAdjust.Rows.Clear()
            For Each rowStockInAdjust In TBLStockInAdjust.Rows
                vPrg.Value += 1
                vLbl.Text = "Updating item in adjustment - data " + vPrg.Value.ToString + " of " + vPrg.Maximum.ToString

                RowItemDetailAdjust = TBLItemDetailAdjust.Newtbltrans_purchase_closing_item_inadjustmentRow
                With RowItemDetailAdjust
                    .Deff_prec = rowStockInAdjust._Decimal
                    .Primary = -1
                    .Primary_Adjustment_Type = -1
                    .Primary_Code_No = rowStockInAdjust.Primary_Product
                    .Primary_Dokumen_Acuan = rowStockInAdjust.Primary_Transaksi
                    .Primary_Gain_Account = -1
                    .Primary_Loss_Account = -1
                    .Primary_Transaksi = vPrimaryClosingTarget
                    .Primary_Units = rowStockInAdjust.Primary_Units
                    .Qty_In_Adjust = rowStockInAdjust.Qty
                    .Remarks = ""
                    .Remarks_Adjust = ""
                    .Tipe_Dokumen_Acuan = ClsTransaksiMonthlyClosingItem.EnumTipeDokumenAcuan.StockInAdjust
                    .Tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
                    .Tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust
                    .Unit_Price = rowStockInAdjust.Amount
                    TBLItemDetailAdjust.Rows.Add(RowItemDetailAdjust)

                    DTAItemDetailAdjust.Insert(.Primary_Transaksi, .Primary_Code_No, .Primary_Units, .Primary_Dokumen_Acuan, .Tipe_Dokumen_Acuan, _
                    .Qty_In_Adjust, .Unit_Price, .Remarks, .Tipe_item_detail, .Tipe_transaksi_asal, .Primary_Adjustment_Type, .Primary_Gain_Account, _
                    .Primary_Loss_Account, .Remarks_Adjust, .Deff_prec)
                End With
            Next

            'pindahkan ke item detail in
            vPrg.Value = 0
            vPrg.Maximum = TBLStockInAdjust.Rows.Count
            vTBLItemDetailIn.Rows.Clear()
            For Each rowStockInAdjust In TBLStockInAdjust.Rows
                vPrg.Value += 1
                vLbl.Text = "Updating item detail closing - data " + vPrg.Value.ToString + " of " + vPrg.Maximum.ToString

                vRowItemDetailIn = vTBLItemDetailIn.Newtbltrans_purchase_closing_item_detailRow
                With vRowItemDetailIn
                    .amount = rowStockInAdjust.Amount
                    .Cost_Perct_Bundling = 0
                    .Overhead = 0
                    .primary = -1
                    .primary_adjustment_type = -1
                    .primary_code_no = rowStockInAdjust.Primary_Product
                    .primary_gain_account = -1
                    .primary_loss_account = -1
                    .primary_transaksi = vPrimaryClosingTarget
                    .primary_transaksi_asal = rowStockInAdjust.Primary_Transaksi
                    .qty = rowStockInAdjust.Qty
                    .remarks = ""
                    .return_price = 0
                    .tgl_transaksi_asal = rowStockInAdjust.Tanggal_Transaksi
                    .tipe_item_detail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
                    .tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust
                    .unit_price = Math.Round(.amount / .qty, DefPrec, MidpointRounding.AwayFromZero)
                    vTBLItemDetailIn.Rows.Add(vRowItemDetailIn)

                    vDTAItemDetailIn.Insert(.primary_transaksi, .primary_transaksi_asal, .tipe_transaksi_asal, .tgl_transaksi_asal, _
                    .primary_code_no, .tipe_item_detail, .unit_price, .amount, .qty, .return_price, .primary_gain_account, _
                    .primary_loss_account, .primary_adjustment_type, .remarks, .Overhead, .Cost_Perct_Bundling)
                End With
            Next

            'MsgBox("Start isikan amount")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TransferKeTarget()
        Dim PrimaryTransaksiSumber As Double = -1
        Dim PrimaryTransaksiTarget As Double = -1

        Dim TblHeader As New DseClosingCutOff01.tbltrans_purchase_closingDataTable
        Dim rowHeader As DseClosingCutOff01.tbltrans_purchase_closingRow
        Dim DTAHeaderSumber As New DseClosingCutOff01TableAdapters.tbltrans_purchase_closingTableAdapter
        Dim DTAHeaderTarget As New DseClosingCutOff01TableAdapters.tbltrans_purchase_closingTableAdapter

        Dim DTAItemSumber As New DseClosingCutOff01TableAdapters.tbltrans_purchase_closing_itemTableAdapter
        Dim DTAItemTarget As New DseClosingCutOff01TableAdapters.tbltrans_purchase_closing_itemTableAdapter
        Dim TBLItem As New DseClosingCutOff01.tbltrans_purchase_closing_itemDataTable
        Dim rowItem As DseClosingCutOff01.tbltrans_purchase_closing_itemRow

        Dim TBLItemDetail As New DseClosingCutOff01.tbltrans_purchase_closing_item_detailDataTable
        Dim rowItemDetail As DseClosingCutOff01.tbltrans_purchase_closing_item_detailRow
        Dim DTAItemDetailSumber As New DseClosingCutOff01TableAdapters.tbltrans_purchase_closing_item_detailTableAdapter
        Dim DTAItemDetailTarget As New DseClosingCutOff01TableAdapters.tbltrans_purchase_closing_item_detailTableAdapter

        Try
            DTAHeaderSumber.Connection.ConnectionString = Me.SourceConstring
            DTAHeaderTarget.Connection.ConnectionString = Me.TargetConstring
            DTAItemSumber.Connection.ConnectionString = Me.SourceConstring
            DTAItemTarget.Connection.ConnectionString = Me.TargetConstring
            DTAItemDetailSumber.Connection.ConnectionString = Me.SourceConstring
            DTAItemDetailTarget.Connection.ConnectionString = Me.TargetConstring

            'tansfer header
            DTAHeaderSumber.Fill(TblHeader, Me.ClosingPeriod.Date)
            rowHeader = TblHeader.Rows(0)
            PrimaryTransaksiSumber = rowHeader.Primary
            With rowHeader
                DTAHeaderTarget.Insert(.No_Transaksi, .Closing_Date, .Closing_Period, .Remark, .Created_By, .Created_Date, .Modified_By, .Modified_Date, .Timestamp, -1)
            End With
            PrimaryTransaksiTarget = DTAHeaderTarget.ScalarQuery_GetPrimaryTransaksi(rowHeader.No_Transaksi).Value

            'transfer item 
            DTAItemSumber.Fill(TBLItem, PrimaryTransaksiSumber)
            For Each rowItem In TBLItem.Rows
                With rowItem
                    DTAItemTarget.Insert(PrimaryTransaksiTarget, .Primary_Product, .Primary_Unit, 0, 0, .Ending_Qty, .Ending_Amount, 0, 0, .Ending_Qty, .Ending_Amount, .Remarks, _
                    .Primary_Inventory_Account, .Primary_Adjustment_Account, .Primary_Stock_Taking_Account, .Primary_COGS_Account, .Primary_Sales_Account, _
                    .Primary_Sales_Return_Account, ._Decimal)
                End With
            Next

            'transfer item detail
            DTAItemDetailSumber.FillBy_PrimaryTransaksi(TBLItemDetail, PrimaryTransaksiSumber, _
            CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.EndingBalanceItem))
            For Each rowItemDetail In TBLItemDetail.Rows
                With rowItemDetail
                    DTAItemDetailTarget.Insert(PrimaryTransaksiTarget, .primary_transaksi_asal, .tipe_transaksi_asal, .tgl_transaksi_asal, .primary_code_no, _
                    .tipe_item_detail, .unit_price, .amount, .qty, .return_price, .primary_gain_account, .primary_loss_account, .primary_adjustment_type, _
                    .remarks, .Overhead, .Cost_Perct_Bundling)
                End With
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BersihkanTabelTarget()
        Dim DTA As New DseClosingCutOff01TableAdapters.tbltrans_purchase_closingTableAdapter

        Try
            DTA.Connection.ConnectionString = Me.TargetConstring
            DTA.DeleteQuery_ByClosingPeriod(Me.ClosingPeriod)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
