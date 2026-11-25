Option Explicit On
Option Strict Off

Public Class FrmTransaksiMonthlyClosing_CyclicBundlingExplain1
    Public mTBLItemDetail As New DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailDataTable
    Public mPrimaryCodeNo As Double = -1
    Public mBackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public mRefPeriod As Date = Today
    Public mDefPrecQty As Integer
    Public mDefPrecCur As Integer

    Dim rowItemDetail As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailRow
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjStockInOut As New Inventory_Module.clsTransaksiStockInOut
    Dim ObjStockBundling As New Inventory_Module.ClsTransaksiStockBundling
    Dim ObjAOP As New ClsTransaksiAOP
    Dim ObjStockTake As New Inventory_Module.clsTransaksiStockTaking
    Dim ObjInv As New Sales_Module.ClsTransaksiSalesInvoice
    Dim ObjPR As New ClsTransaksiPurchaseReturn
    Dim ObjProd As New Core_Module.clsMasterProduct

    Private Sub FrmTransaksiMonthlyClosing_CyclicBundlingExplain1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim TipeItemDetail1 As ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.BeginingBalanceItem
        Dim TipeItemDetail2 As ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.EndingBalanceItem
        Dim TipeItemDetail3 As ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.EndingBalanceItem
        Dim TipeTransaksiAsal As ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust
        Dim OldSort As String = ""
        Dim OldFilter As String = ""
        Dim Jml As Decimal = 0
        Dim DV As DataRowView
        Dim Amount As Decimal = 0

        Try
            Windows.Forms.Cursor.Current = Windows.Forms.Cursors.WaitCursor

            Me.dgvData1.AlternatingRowsDefaultCellStyle.BackColor = Me.mBackColorAll
            Me.dgvData1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            OldSort = Me.mTBLItemDetail.DefaultView.Sort
            OldFilter = Me.mTBLItemDetail.DefaultView.RowFilter

            'load data + sort
            '1. begining balance
            TipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.BeginingBalanceItem
            Me.mTBLItemDetail.DefaultView.Sort = "[tgl transaksi asal], [tipe transaksi asal], [primary transaksi asal], [primary code no]"
            Me.mTBLItemDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(TipeItemDetail1).ToString + " AND [primary code no] = " + Me.mPrimaryCodeNo.ToString

            Jml = 0
            For Each DV In Me.mTBLItemDetail.DefaultView
                Me.rowItemDetail = DV.Row
                Jml += Me.rowItemDetail.qty
                Amount += Me.rowItemDetail.amount
            Next
            Me.AddRowDGV(-1, ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OverheadCost, Jml, Me.ObjCD.GetFirstDateofTheMonth(Me.mRefPeriod), Amount, TipeItemDetail1)

            '2. mulai per tanggal
            TipeItemDetail3 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
            Me.mTBLItemDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(TipeItemDetail3).ToString + " AND [primary code no] = " + Me.mPrimaryCodeNo.ToString
            For Each DV In Me.mTBLItemDetail.DefaultView
                Me.rowItemDetail = DV.Row
                If Me.rowItemDetail.tipe_item_detail <> TipeItemDetail1 And Me.rowItemDetail.tipe_item_detail <> TipeItemDetail2 Then
                    Me.AddRowDGV(Me.rowItemDetail.primary_transaksi_asal, Me.rowItemDetail.tipe_transaksi_asal, Me.rowItemDetail.qty, Me.rowItemDetail.tgl_transaksi_asal, _
                    Me.rowItemDetail.amount, Me.rowItemDetail.tipe_item_detail)
                End If
            Next
            TipeItemDetail3 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem
            Me.mTBLItemDetail.DefaultView.RowFilter = "[tipe item detail] = " + CInt(TipeItemDetail3).ToString + " AND [primary code no] = " + Me.mPrimaryCodeNo.ToString
            For Each DV In Me.mTBLItemDetail.DefaultView
                Me.rowItemDetail = DV.Row
                If Me.rowItemDetail.tipe_item_detail <> TipeItemDetail1 And Me.rowItemDetail.tipe_item_detail <> TipeItemDetail2 Then
                    Me.AddRowDGV(Me.rowItemDetail.primary_transaksi_asal, Me.rowItemDetail.tipe_transaksi_asal, Me.rowItemDetail.qty, Me.rowItemDetail.tgl_transaksi_asal, _
                    Me.rowItemDetail.amount, Me.rowItemDetail.tipe_item_detail)
                End If
            Next

            'tampilkan ke dgv
            Me.mTBLItemDetail.DefaultView.Sort = OldSort
            Me.mTBLItemDetail.DefaultView.RowFilter = OldFilter

            Me.ColData1_QtyIn.DefaultCellStyle.Format = "N" + Me.mDefPrecQty.ToString
            Me.ColData1_QtyOut.DefaultCellStyle.Format = "N" + Me.mDefPrecQty.ToString
            Me.ColData1_Balance.DefaultCellStyle.Format = "N" + Me.mDefPrecQty.ToString
            Me.ColData1_Amount.DefaultCellStyle.Format = "N" + Me.mDefPrecCur.ToString
            Me.ColData1_Date.DefaultCellStyle.Format = "dd MMM yyyy"

            Me.dgvData1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            Me.Text = Me.Text + ": " + Me.ObjProd.GetProductCode(Me.mPrimaryCodeNo) + " - " + Me.ObjProd.GetDescription(Me.mPrimaryCodeNo)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error") 
        End Try
    End Sub

    Private Sub AddRowDGV(ByVal vPrimaryDocRef As Double, ByVal vTipeDocRef As ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal, ByVal vQty As Decimal, _
    ByVal vDate As Date, ByVal vAmount As Decimal, _
    ByVal vTipeItemDetail1 As ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail)
        Dim rowDGV As DataGridViewRow
        Dim rowDGVB4 As DataGridViewRow
        Dim JmlBalance As Decimal = 0

        Try
            If Me.dgvData1.Rows.Count > 0 Then
                rowDGVB4 = Me.dgvData1.Rows(Me.dgvData1.RowCount - 1)
                JmlBalance = CDec(rowDGVB4.Cells(Me.ColData1_Balance.Index).Value)
            End If

            Me.dgvData1.Rows.Add()
            rowDGV = Me.dgvData1.Rows(Me.dgvData1.RowCount - 1)

            rowDGV.Cells(Me.ColData1_Amount.Index).Value = vAmount

            If vTipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.BeginingBalanceItem Then
                rowDGV.Cells(Me.ColData1_Desc.Index).Value = "Begining Balance"
            ElseIf vTipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem Or _
            vTipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem Then
                If vTipeDocRef = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust Then
                    rowDGV.Cells(Me.ColData1_Desc.Index).Value = "Stock In Adjust"
                ElseIf vTipeDocRef = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling Then
                    rowDGV.Cells(Me.ColData1_Desc.Index).Value = "In Bundling"
                ElseIf vTipeDocRef = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InPurchase Then
                    rowDGV.Cells(Me.ColData1_Desc.Index).Value = "In Purchase"
                ElseIf vTipeDocRef = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR Then
                    rowDGV.Cells(Me.ColData1_Desc.Index).Value = "In SR"
                ElseIf vTipeDocRef = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InStockTakeAdjust Then
                    rowDGV.Cells(Me.ColData1_Desc.Index).Value = "In Stock Take"
                ElseIf vTipeDocRef = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutAdjust Then
                    rowDGV.Cells(Me.ColData1_Desc.Index).Value = "Out Adjust"
                ElseIf vTipeDocRef = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling Then
                    rowDGV.Cells(Me.ColData1_Desc.Index).Value = "Out Bundling"
                ElseIf vTipeDocRef = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutPR Then
                    rowDGV.Cells(Me.ColData1_Desc.Index).Value = "Out PR"
                ElseIf vTipeDocRef = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutSales Then
                    rowDGV.Cells(Me.ColData1_Desc.Index).Value = "Out Sales"
                ElseIf vTipeDocRef = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutStockTakeAdjust Then
                    rowDGV.Cells(Me.ColData1_Desc.Index).Value = "Out Stock Take"
                End If
            Else
                MsgBox("Logic error: no more code here.")
                Exit Sub
            End If
            rowDGV.Cells(Me.ColData1_DocNo.Index).Value = vPrimaryDocRef
            rowDGV.Cells(Me.ColData1_PrimaryDocRef.Index).Value = vPrimaryDocRef

            If vTipeDocRef = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust Or _
            vTipeDocRef = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling Or _
            vTipeDocRef = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InPurchase Or _
            vTipeDocRef = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR Or _
            vTipeDocRef = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InStockTakeAdjust Then
                rowDGV.Cells(Me.ColData1_Balance.Index).Value = JmlBalance + vQty
                rowDGV.Cells(Me.ColData1_QtyIn.Index).Value = vQty
                rowDGV.Cells(Me.ColData1_QtyOut.Index).Value = 0
            Else
                rowDGV.Cells(Me.ColData1_Balance.Index).Value = JmlBalance - vQty
                rowDGV.Cells(Me.ColData1_QtyIn.Index).Value = 0
                rowDGV.Cells(Me.ColData1_QtyOut.Index).Value = vQty
            End If
            rowDGV.Cells(Me.ColData1_TipeItemDetail.Index).Value = vTipeItemDetail1
            rowDGV.Cells(Me.ColData1_TipeDocRef.Index).Value = vTipeDocRef
            rowDGV.Cells(Me.ColData1_Date.Index).Value = vDate
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Private Sub AddRow(ByVal vPrimaryDocRef As Double, ByVal vTipeDocRef As ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal, ByVal vQty As Decimal, _
    'ByVal vDate As Date, ByVal vAmount As Decimal, ByVal vTBL As DseTransaksiMonthlyClosing3.TblCodeNoDetailDataTable, _
    'ByVal vTipeItemDetail1 As ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail)
    '    Dim row As DseTransaksiMonthlyClosing3.TblCodeNoDetailRow
    '    Dim JmlBalance As Decimal = 0

    '    Try
    '        If vTBL.Rows.Count > 0 Then
    '            row = vTBL.Rows(vTBL.Rows.Count - 1)
    '            JmlBalance = row.Qty_Balance
    '        End If
    '        row = vTBL.NewTblCodeNoDetailRow

    '        row.Amount = vAmount
    '        If vTipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.BeginingBalanceItem Then
    '            row.Desc = "Begining Balance"
    '        ElseIf vTipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem Or _
    '        vTipeItemDetail1 = ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem Then
    '            If vTipeDocRef = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust Then
    '                row.Desc = "Stock In Adjust"
    '            ElseIf vTipeDocRef = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling Then
    '                row.Desc = "In Bundling"
    '            ElseIf vTipeDocRef = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InPurchase Then
    '                row.Desc = "In Purchase"
    '            ElseIf vTipeDocRef = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR Then
    '                row.Desc = "In SR"
    '            ElseIf vTipeDocRef = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InStockTakeAdjust Then
    '                row.Desc = "In Stock Take"
    '            ElseIf vTipeDocRef = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutAdjust Then
    '                row.Desc = "Out Adjust"
    '            ElseIf vTipeDocRef = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling Then
    '                row.Desc = "Out Bundling"
    '            ElseIf vTipeDocRef = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutPR Then
    '                row.Desc = "Out PR"
    '            ElseIf vTipeDocRef = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutSales Then
    '                row.Desc = "Out Sales"
    '            ElseIf vTipeDocRef = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutStockTakeAdjust Then
    '                row.Desc = "Out Stock Take"
    '            End If
    '        Else
    '            MsgBox("Logic error: no more code here.")
    '            Exit Sub
    '        End If
    '        row.Doc_No = ""
    '        row.Primary_Ref_Doc = vPrimaryDocRef

    '        If vTipeDocRef = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust Or _
    '        vTipeDocRef = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling Or _
    '        vTipeDocRef = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InPurchase Or _
    '        vTipeDocRef = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR Or _
    '        vTipeDocRef = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InStockTakeAdjust Then
    '            row.Qty_Balance = JmlBalance + vQty
    '            row.Qty_In = vQty
    '            row.Qty_Out = 0
    '        Else
    '            row.Qty_Balance = JmlBalance - vQty
    '            row.Qty_In = 0
    '            row.Qty_Out = vQty
    '        End If
    '        row.Tipe_Item_Detail = vTipeItemDetail1
    '        row.Tipe_Ref_Doc = vTipeDocRef
    '        row._Date = vDate

    '        vTBL.Rows.Add(row)
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Private Sub dgvData1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvData1.CellClick 

        Try
            If e.RowIndex <> -1 Then
                If e.ColumnIndex = Me.ColData1_ShowDocNo.Index Then
                    Me.IsiKanDocNo(e)
                ElseIf e.ColumnIndex = Me.ColData1_DocNo.Index Then
                    Dim ij As New FrmTransaksiMonthlyClosing_CyclicBundlingExplain2
                    ij.mPrimaryRefDoc = Me.dgvData1.Item(Me.ColData1_PrimaryDocRef.Index, e.RowIndex).Value
                    ij.mTBLItemDetail = Me.mTBLItemDetail
                    ij.mRefPeriod = Me.mRefPeriod
                    ij.mTipRefDoc = Me.dgvData1.Item(Me.ColData1_TipeDocRef.Index, e.RowIndex).Value
                    ij.mBackColorAll = Me.mBackColorAll
                    ij.mDefPrecCur = Me.mDefPrecCur
                    Me.IsiKanDocNo(e)
                    ij.mDocNo = Me.dgvData1.Item(Me.ColData1_DocNo.Index, e.RowIndex).Value
                    Me.dgvData1.ClearSelection()

                    ij.ShowDialog()

                    ij = Nothing
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiKanDocNo(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim TipeDocRef As ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal

        Try
            TipeDocRef = Me.dgvData1.Item(Me.ColData1_TipeDocRef.Index, e.RowIndex).Value
            If TipeDocRef = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust Or _
            TipeDocRef = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR Or _
            TipeDocRef = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutAdjust Then
                Me.dgvData1.Item(Me.ColData1_DocNo.Index, e.RowIndex).Value = Me.ObjStockInOut.GetNoTransaksi(Me.dgvData1.Item(Me.ColData1_PrimaryDocRef.Index, e.RowIndex).Value)
            ElseIf TipeDocRef = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InPurchase Then
                Me.dgvData1.Item(Me.ColData1_DocNo.Index, e.RowIndex).Value = Me.ObjAOP.GetNoTransaksi(Me.dgvData1.Item(Me.ColData1_PrimaryDocRef.Index, e.RowIndex).Value)
            ElseIf TipeDocRef = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling Or _
            TipeDocRef = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling Then
                Me.dgvData1.Item(Me.ColData1_DocNo.Index, e.RowIndex).Value = Me.ObjStockBundling.GetNoTransaksi(Me.dgvData1.Item(Me.ColData1_PrimaryDocRef.Index, e.RowIndex).Value)
            ElseIf TipeDocRef = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InStockTakeAdjust Or _
            TipeDocRef = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutStockTakeAdjust Then
                Me.dgvData1.Item(Me.ColData1_DocNo.Index, e.RowIndex).Value = _
                Me.ObjStockTake.GetNoTransaksi(Me.dgvData1.Item(Me.ColData1_PrimaryDocRef.Index, e.RowIndex).Value)
            ElseIf TipeDocRef = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutSales Then
                Me.dgvData1.Item(Me.ColData1_DocNo.Index, e.RowIndex).Value = _
                Me.ObjInv.GetNoInvoice(Me.dgvData1.Item(Me.ColData1_PrimaryDocRef.Index, e.RowIndex).Value)
            ElseIf TipeDocRef = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutPR Then
                Me.dgvData1.Item(Me.ColData1_DocNo.Index, e.RowIndex).Value = _
                Me.ObjPR.GetNoTransaksi(Me.dgvData1.Item(Me.ColData1_PrimaryDocRef.Index, e.RowIndex).Value)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class