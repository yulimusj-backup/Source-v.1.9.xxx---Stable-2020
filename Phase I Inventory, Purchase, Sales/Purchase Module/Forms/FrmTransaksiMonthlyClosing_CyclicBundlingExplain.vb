Option Explicit On
Option Strict Off

Public Class FrmTransaksiMonthlyClosing_CyclicBundlingExplain
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public TBL_Pending As New DseMonthlyClosingNew.tblItemDetailPoolDataTable
    Public StartDate, EndDate As Date
    Public ClosingPeriod As Date
    Public TBL_Pending2 As DseMonthlyClosingNew.tblItemDetailPoolDataTable
    Public TBLItemClosing As New DseMonthlyClosingNew.DataTable7DataTable
    Public TBLItemDetail As DseMonthlyClosingNew.tbltrans_purchase_closing_item_detailDataTable

    Dim row_Pending As DseMonthlyClosingNew.tblItemDetailPoolRow
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim ObjUnit As New Core_Module.ClsMasterUnit
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim ObjBundling As New Inventory_Module.ClsTransaksiStockBundling
    Dim ObjClosing As New ClsTransaksiMonthlyClosingItem
    Dim Buffer As String = ""
    Dim LangsungTutup As Boolean = False
    Dim DefPrec1 As Integer = 0

    Dim DTAStockInOutAdjust As New DseMonthlyClosingNewTableAdapters.tbltransaksi_stockinoutTableAdapter
    Dim TBLStockInOutAdjust As New DseMonthlyClosingNew.tbltransaksi_stockinoutDataTable
    Dim rowStockInOut As DseMonthlyClosingNew.tbltransaksi_stockinoutRow
    Dim DTABundling As New DseMonthlyClosingNewTableAdapters.tbltransaksi_stockbundlingTableAdapter
    Dim TblBundling As New DseMonthlyClosingNew.tbltransaksi_stockbundlingDataTable
    Dim rowBundling As DseMonthlyClosingNew.tbltransaksi_stockbundlingRow
    Dim DTAAOP As New DseMonthlyClosingNewTableAdapters.tbltrans_aopTableAdapter
    Dim TBLAOP As New DseMonthlyClosingNew.tbltrans_aopDataTable
    Dim rowAOP As DseMonthlyClosingNew.tbltrans_aopRow
    Dim DTAInv As New DseMonthlyClosingNewTableAdapters.tbltrans_salesinvoiceTableAdapter
    Dim TBLInv As New DseMonthlyClosingNew.tbltrans_salesinvoiceDataTable
    Dim rowInv As DseMonthlyClosingNew.tbltrans_salesinvoiceRow
    Dim DTA_PR As New DseMonthlyClosingNewTableAdapters.tbltrans_purchasereturnTableAdapter
    Dim TBL_PR As New DseMonthlyClosingNew.tbltrans_purchasereturnDataTable
    Dim row_PR As DseMonthlyClosingNew.tbltrans_purchasereturnRow
    Dim DTA_StockTaking As New DseMonthlyClosingNewTableAdapters.tbltrans_stock_takingTableAdapter
    Dim TBL_StockTaking As New DseMonthlyClosingNew.tbltrans_stock_takingDataTable
    Dim row_StockTaking As DseMonthlyClosingNew.tbltrans_stock_takingRow
    Dim row_Pending2 As DseMonthlyClosingNew.tblItemDetailPoolRow

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim ListPrimaryBundling As New List(Of Double)
        Dim rowDGV As DataGridViewRow

        Try
            'minimal ada 1 transaksi out bundling yang lengkap yang diisikan unit price-nya
            '1. catat semua primary transaksi out bundling
            For Each rowDGV In Me.dgvItem.Rows
                If CDec(rowDGV.Cells(Me.ColUnitPrice.Index).Value) <> 0 Then
                    If rowDGV.Cells(Me.ColTipeTransaksiAsal.Index).Value = _
                    CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling) Then
                        If ListPrimaryBundling.Contains(rowDGV.Cells(Me.ColPrimaryDocAcuan.Index).Value) = False Then
                            ListPrimaryBundling.Add(rowDGV.Cells(Me.ColPrimaryDocAcuan.Index).Value)
                        End If
                    End If
                End If
            Next
            If ListPrimaryBundling.Count = 0 Then
                MsgBox("Please specify unit price at least for 1 bundling transaction.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            '2. pastikan semua out bundling untuk setiap transaksi tersebut ada nilainya
            For Each PrimaryBundling As Double In ListPrimaryBundling
                For Each rowDGV In Me.dgvItem.Rows
                    If rowDGV.Cells(Me.ColTipeTransaksiAsal.Index).Value = _
                    CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling) Then
                        If PrimaryBundling = rowDGV.Cells(Me.ColPrimaryDocAcuan.Index).Value Then
                            If CDec(rowDGV.Cells(Me.ColUnitPrice.Index).Value) = 0 Then
                                MsgBox("Please make sure every out bundling item on transaction: " + _
                                Me.ObjBundling.GetNoTransaksi(PrimaryBundling) + " has a positive value.", MsgBoxStyle.Information, "Information")
                                Exit Sub
                            End If
                        End If
                    End If
                Next
            Next

            'jika ok, tutup dengan baik
            Me.DialogResult = Windows.Forms.DialogResult.OK
            Me.LangsungTutup = True
            Me.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadData()
        Dim vTipeTransaksiAsal As ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust
        Dim Lanjut As Boolean = False
        Dim ZeroVal As String = ""
        Dim rowDGV As DataGridViewRow

        Try
            ZeroVal = Me.ObjCD.DigitGrouping(0, DefPrec1)
            Me.TBLAOP.Rows.Clear()
            Me.TblBundling.Rows.Clear()
            Me.TBLInv.Rows.Clear()
            Me.TBLStockInOutAdjust.Rows.Clear()
            Me.TBL_PR.Rows.Clear()
            Me.TBL_StockTaking.Rows.Clear()

            Me.DTAAOP.Fill(Me.TBLAOP, EndDate, StartDate)
            Me.DTABundling.Fill(Me.TblBundling, EndDate, StartDate)
            Me.DTAInv.Fill(Me.TBLInv, EndDate, StartDate)
            Me.DTAStockInOutAdjust.Fill(Me.TBLStockInOutAdjust, EndDate, StartDate)
            Me.DTA_PR.Fill(Me.TBL_PR, StartDate, EndDate)
            Me.DTA_StockTaking.Fill(Me.TBL_StockTaking)

            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.Rows.Clear()

            For Each row_Pending In Me.TBL_Pending.Rows
                vTipeTransaksiAsal = row_Pending.tipe_transaksi_asal
                Lanjut = True

                If vTipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust And Me.chkInAdjust.Checked = False Then
                    Lanjut = False
                ElseIf vTipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling And Me.chkInBundling.Checked = False Then
                    Lanjut = False
                ElseIf vTipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InPurchase And Me.chkInPurchase.Checked = False Then
                    Lanjut = False
                ElseIf vTipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR And Me.chkInSalesReturn.Checked = False Then
                    Lanjut = False
                ElseIf vTipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InStockTakeAdjust And Me.chkInAdjust.Checked = False Then
                    Lanjut = False
                ElseIf vTipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutAdjust And Me.chkOutAdjust.Checked = False Then
                    Lanjut = False
                ElseIf vTipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling And Me.chkOutBundling.Checked = False Then
                    Lanjut = False
                ElseIf vTipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutPR And Me.chkOutPurchaseReturn.Checked = False Then
                    Lanjut = False
                ElseIf vTipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutSales And Me.chkOutSales.Checked = False Then
                    Lanjut = False
                ElseIf vTipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutStockTakeAdjust And Me.chkOutAdjust.Checked = False Then
                    Lanjut = False
                Else
                    'MsgBox("Logic error: invalid transaction type. Please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                    'Exit For
                End If

                If Lanjut = True Then
                    Me.dgvItem.Rows.Add()

                    rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                    rowDGV.Cells(Me.ColCodeNoAdjust.Index).Value = Me.ObjProduct.GetProductCode(row_Pending.primary_code_no)
                    rowDGV.Cells(Me.ColDeffPrec.Index).Value = Me.ObjProduct.GetAngkaDecimal(row_Pending.primary_code_no)
                    rowDGV.Cells(Me.ColDescAdjust.Index).Value = Me.ObjProduct.GetDescription(row_Pending.primary_code_no)

                    rowDGV.Cells(Me.ColDocumentNo.Index).Value = Me.GetDocumentNo(row_Pending.Primary_Transaksi_Asal, row_Pending.tipe_transaksi_asal)

                    rowDGV.Cells(Me.ColJenisDocumentAdjustStr.Index).Value = GetTipeDocument_TransaksiAsal(row_Pending.tipe_transaksi_asal)
                    rowDGV.Cells(Me.ColTglDocAcuan.Index).Value = row_Pending.tgl_transaksi_asal
                    rowDGV.Cells(Me.ColNoAdjust.Index).Value = rowDGV.Index + 1
                    rowDGV.Cells(Me.ColPrimaryCodeNoAdjust.Index).Value = row_Pending.primary_code_no
                    rowDGV.Cells(Me.ColPrimaryDocAcuan.Index).Value = row_Pending.Primary_Transaksi_Asal
                    rowDGV.Cells(Me.ColUnitAdjust.Index).Value = Me.ObjUnit.GetUnitCode(Me.ObjProduct.GetPrimaryUnit(row_Pending.primary_code_no))
                    rowDGV.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(row_Pending.qty, rowDGV.Cells(Me.ColDeffPrec.Index).Value)

                    Me.TBL_Pending2.DefaultView.RowFilter = "[Primary Transaksi Asal] = " + row_Pending.Primary_Transaksi_Asal.ToString + _
                        " AND [tipe transaksi asal] = " + row_Pending.tipe_transaksi_asal.ToString + _
                        " AND [primary code no] = " + row_Pending.primary_code_no.ToString
                    If Me.TBL_Pending2.DefaultView.Count = 0 Then
                        rowDGV.Cells(Me.ColUnitPrice.Index).Value = ZeroVal
                    Else
                        Me.row_Pending = Me.TBL_Pending2.DefaultView(0).Row
                        rowDGV.Cells(Me.ColUnitPrice.Index).Value = Me.ObjCD.DigitGrouping(Me.row_Pending.unit_price, DefPrec1)
                    End If
                    rowDGV.Cells(Me.ColTipeTransaksiAsal.Index).Value = row_Pending.tipe_transaksi_asal
                    rowDGV.Cells(Me.ColTipeItemDetail.Index).Value = row_Pending.tipe_item_detail
                End If
            Next

            Me.ColTglDocAcuan.DefaultCellStyle.Format = "dd MMM yyyy"
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiMonthlyClosing_CyclicBundlingExplain_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.LangsungTutup = False Then
                If MsgBox("Close and discard data changes?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Me.DialogResult = Windows.Forms.DialogResult.Cancel
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiMonthlyClosing_CyclicBundlingExplain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.DefPrec1 = Me.ObjCur.GetDecimal_DefaultCurrency

            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.chkInBundling.Checked = False

            Me.DTAAOP.Connection.ConnectionString = Core_Module.GetConstring
            Me.DTABundling.Connection.ConnectionString = Core_Module.GetConstring
            Me.DTAInv.Connection.ConnectionString = Core_Module.GetConstring
            Me.DTAStockInOutAdjust.Connection.ConnectionString = Core_Module.GetConstring
            Me.DTA_PR.Connection.ConnectionString = Core_Module.GetConstring
            DTA_StockTaking.Connection.ConnectionString = Core_Module.GetConstring

            If Me.chkAutomatic.Checked = True Then
                If Me.IsikanTabelPending2_HargaBaru = True Then
                    'langsung tutup
                    Me.DialogResult = Windows.Forms.DialogResult.OK
                    Me.LangsungTutup = True
                    Me.Close()
                Else
                    Me.LoadData()
                End If
            Else
                Me.LoadData()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function IsikanTabelPending2_HargaBaru() As Boolean
        Dim ListPrimary As New List(Of Double)
        Dim PrimaryAcuan As Double = -1
        Dim TBL_Pending3 As New DseMonthlyClosingNew.tblItemDetailPoolDataTable
        Dim row_Pending11, row_Pending3 As DseMonthlyClosingNew.tblItemDetailPoolRow
        Dim Lengkap As Boolean = False
        Dim rowItemClosing As DseMonthlyClosingNew.DataTable7Row

        Try

            For Each row_Pending In Me.TBL_Pending.Rows
                If row_Pending.tipe_transaksi_asal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling Then
                    If ListPrimary.Contains(row_Pending.Primary_Transaksi_Asal) = False Then
                        ListPrimary.Add(row_Pending.Primary_Transaksi_Asal)

                        Me.TBL_Pending2.DefaultView.RowFilter = "[Primary Transaksi Asal] = " + row_Pending.Primary_Transaksi_Asal.ToString
                        If Me.TBL_Pending2.DefaultView.Count = 0 Then
                            TBL_Pending3.Rows.Clear()
                            Lengkap = True

                            'copy data ke tbl pending 3
                            Me.TBL_Pending.DefaultView.RowFilter = "[Primary Transaksi Asal] = " + row_Pending.Primary_Transaksi_Asal.ToString + _
                            " AND [tipe transaksi asal] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling).ToString
                            For Each dv As Data.DataRowView In Me.TBL_Pending.DefaultView
                                row_Pending11 = dv.Row
                                row_Pending3 = TBL_Pending3.NewtblItemDetailPoolRow

                                For ij As Integer = 0 To Me.TBL_Pending.Columns.Count - 1
                                    row_Pending3.Item(ij) = row_Pending11(ij)
                                Next
                                TBL_Pending3.Rows.Add(row_Pending3)
                            Next

                            'cari data ref price
                            TBLItemClosing.DefaultView.Sort = "[tgl transaksi asal] DESC, [primary transaksi asal] DESC"
                            For Each row_Pending3 In TBL_Pending3.Rows
                                TBLItemClosing.DefaultView.RowFilter = "[primary code no] = " + row_Pending3.primary_code_no.ToString
                                If TBLItemClosing.DefaultView.Count > 0 Then
                                    rowItemClosing = TBLItemClosing.DefaultView(0).Row
                                    row_Pending3.unit_price = rowItemClosing.unit_price
                                Else
                                    row_Pending3.unit_price = 0
                                    Lengkap = False
                                End If
                            Next

                            'jika semua item lengkap, maka pindahkan ke tabel pending 2, isikan data item pending baru berhasil
                            'return true
                            If Lengkap = True Then
                                'tambahkan ke tabel pending 2
                                For Each row_Pending3 In TBL_Pending3.Rows
                                    MasukkanUnitPrice(row_Pending3.Primary_Transaksi_Asal, row_Pending3.primary_code_no, row_Pending3.tipe_transaksi_asal, _
                                    row_Pending3.unit_price, row_Pending3.qty, row_Pending3.tgl_transaksi_asal, row_Pending3.tipe_item_detail)
                                Next
                                Return True
                            End If
                        End If
                    End If
                End If
            Next

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetTipeDocument_TransaksiAsal(ByVal vTipeTransaksiAsal As ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal) As String
        Try
            If vTipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust Then
                Return "In Adjustment"
            ElseIf vTipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling Then
                Return "In Bundling"
            ElseIf vTipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InPurchase Then
                Return "In Purchase"
            ElseIf vTipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR Then
                Return "In Sales Return"
            ElseIf vTipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InStockTakeAdjust Then
                Return "In Stock Take"
            ElseIf vTipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutAdjust Then
                Return "Out Adjustment"
            ElseIf vTipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling Then
                Return "Out Bundling"
            ElseIf vTipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutPR Then
                Return "Out Purchase Return"
            ElseIf vTipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutSales Then
                Return "Out Sales"
            ElseIf vTipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutStockTakeAdjust Then
                Return "Out Stock Take"
            End If

            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function GetDocumentNo(ByVal vPrimaryTransksiAsal As Double, ByVal vTipeTransaksiAsal As ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal) As String
        Try
            If vTipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust Or vTipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutAdjust Then
                Me.TBLStockInOutAdjust.DefaultView.RowFilter = "Primary = " + vPrimaryTransksiAsal.ToString
                If Me.TBLStockInOutAdjust.DefaultView.Count > 0 Then
                    Me.rowStockInOut = Me.TBLStockInOutAdjust.DefaultView(0).Row
                    Return Me.rowStockInOut.No_Transaksi
                Else
                    Return ""
                End If
            ElseIf vTipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling Or vTipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling Then
                Me.TblBundling.DefaultView.RowFilter = "Primary = " + vPrimaryTransksiAsal.ToString
                If Me.TblBundling.DefaultView.Count > 0 Then
                    Me.rowBundling = Me.TblBundling.DefaultView(0).Row
                    Return Me.rowBundling.No_Transaksi
                Else
                    Return ""
                End If
            ElseIf vTipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InPurchase Then
                Me.TBLAOP.DefaultView.RowFilter = "Primary = " + vPrimaryTransksiAsal.ToString
                If Me.TBLAOP.DefaultView.Count > 0 Then
                    Me.rowAOP = Me.TBLAOP.DefaultView(0).Row
                    Return Me.rowAOP.No_Transaksi
                Else
                    Return ""
                End If
            ElseIf vTipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR Then
                Me.TBLStockInOutAdjust.DefaultView.RowFilter = "Primary = " + vPrimaryTransksiAsal.ToString
                If Me.TBLStockInOutAdjust.DefaultView.Count > 0 Then
                    Me.rowStockInOut = Me.TBLStockInOutAdjust.DefaultView(0).Row
                    Return Me.rowStockInOut.No_Transaksi
                Else
                    Return ""
                End If
            ElseIf vTipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InStockTakeAdjust Or _
             vTipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutStockTakeAdjust Then
                Me.TBL_StockTaking.DefaultView.RowFilter = "Primary = " + vPrimaryTransksiAsal.ToString
                If Me.TBL_StockTaking.DefaultView.Count > 0 Then
                    Me.row_StockTaking = Me.TBL_StockTaking.DefaultView(0).Row
                    Return Me.row_StockTaking.No_Transaksi
                Else
                    Return ""
                End If
            ElseIf vTipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutPR Then
                Me.TBL_PR.DefaultView.RowFilter = "Primary = " + vPrimaryTransksiAsal.ToString
                If Me.TBL_PR.DefaultView.Count > 0 Then
                    Me.row_PR = Me.TBL_PR.DefaultView(0).Row
                    Return Me.row_PR.No_Transaksi
                Else
                    Return ""
                End If
            ElseIf vTipeTransaksiAsal = ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutSales Then
                Me.TBLInv.DefaultView.RowFilter = "Primary = " + vPrimaryTransksiAsal.ToString
                If Me.TBLInv.DefaultView.Count > 0 Then
                    Me.rowInv = Me.TBLInv.DefaultView(0).Row
                    Return Me.rowInv.No_Transaksi
                Else
                    Return ""
                End If
            End If

            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Sub btnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter.Click
        Try
            Me.LoadData()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColUnitPrice.Index Then
                If Me.dgvItem.Item(Me.ColTipeTransaksiAsal.Index, e.RowIndex).Value = CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling) Then
                    Me.Buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.BuangNolDibelakangDecimal(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                Else
                    e.Cancel = True
                End If
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub 

    Private Sub dgvItem_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellClick
        Dim ij As New FrmTransaksiMonthlyClosing_GetReferencePrice

        Try
            If e.RowIndex <> -1 Then
                If e.ColumnIndex = Me.ColSetUnitPrice.Index Then
                    'get jml reference
                    If Me.dgvItem.Item(Me.ColTipeTransaksiAsal.Index, e.RowIndex).Value <> CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling) Then
                        'MsgBox("No reference detected", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    Else
                        If Me.ObjClosing.GetJmlItemIn(Me.dgvItem.Item(Me.ColPrimaryCodeNoAdjust.Index, e.RowIndex).Value) > 0 Then
                            ij.PrimaryProduct = Me.dgvItem.Item(Me.ColPrimaryCodeNoAdjust.Index, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ClosingPeriod = Me.ClosingPeriod

                            If ij.ShowDialog = Windows.Forms.DialogResult.OK Then
                                Me.dgvItem.Item(Me.ColUnitPrice.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(ij.UnitPrice, Me.DefPrec1)

                                Me.MasukkanUnitPrice(Me.dgvItem.Item(Me.ColPrimaryDocAcuan.Index, e.RowIndex).Value, _
                                Me.dgvItem.Item(Me.ColPrimaryCodeNoAdjust.Index, e.RowIndex).Value, _
                                Me.dgvItem.Item(Me.ColTipeTransaksiAsal.Index, e.RowIndex).Value, Me.dgvItem.Item(Me.ColUnitPrice.Index, e.RowIndex).Value, _
                                Me.dgvItem.Item(Me.ColQty.Index, e.RowIndex).Value, Me.dgvItem.Item(Me.ColTglDocAcuan.Index, e.RowIndex).Value, _
                                Me.dgvItem.Item(Me.ColTipeItemDetail.Index, e.RowIndex).Value)
                            End If
                        Else
                            MsgBox("No reference detected", MsgBoxStyle.Information, "Information")
                            Exit Sub
                        End If
                    End If
                ElseIf e.ColumnIndex = Me.ColCodeNoAdjust.Index Then
                    Dim Frm1 As New FrmTransaksiMonthlyClosing_CyclicBundlingExplain1

                    Frm1.mTBLItemDetail = Me.TBLItemDetail
                    Frm1.mRefPeriod = Me.ClosingPeriod
                    Frm1.mPrimaryCodeNo = Me.dgvItem.Item(Me.ColPrimaryCodeNoAdjust.Index, e.RowIndex).Value
                    Frm1.mDefPrecQty = Me.dgvItem.Item(Me.ColDeffPrec.Index, e.RowIndex).Value
                    Frm1.mDefPrecCur = Me.DefPrec1

                    Frm1.ShowDialog()
                    Frm1.Dispose()
                ElseIf e.ColumnIndex = Me.ColDocumentNo.Index Then
                    Dim mn As New FrmTransaksiMonthlyClosing_CyclicBundlingExplain2
                    mn.mPrimaryRefDoc = Me.dgvItem.Item(Me.ColPrimaryDocAcuan.Index, e.RowIndex).Value
                    mn.mTBLItemDetail = Me.TBLItemDetail
                    mn.mRefPeriod = Me.ClosingPeriod
                    mn.mTipRefDoc = Me.dgvItem.Item(Me.ColTipeTransaksiAsal.Index, e.RowIndex).Value
                    mn.mBackColorAll = Me.BackColorAll
                    mn.mDefPrecCur = Me.DefPrec1
                    mn.mDocNo = Me.dgvItem.Item(Me.ColDocumentNo.Index, e.RowIndex).Value

                    mn.ShowDialog()

                    ij = Nothing
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub MasukkanUnitPrice(ByVal vPrimaryTransaksiAsal As Double, ByVal vPrimaryCodeNo As Double, _
    ByVal vTipeTransaksiAsal As ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal, ByVal vUnitPrice As Decimal, _
    ByVal vQty As Decimal, ByVal vTglTransaksiAcuan As Date, ByVal vTipeItemDetail As Integer)

        Try
            'masukkan ke tabel bundling
            Me.TBL_Pending2.DefaultView.RowFilter = "[Primary Transaksi Asal] = " + vPrimaryTransaksiAsal.ToString + _
            " AND [tipe transaksi asal] = " + CInt(vTipeTransaksiAsal).ToString + _
            " AND [primary code no] = " + vPrimaryCodeNo.ToString

            If Me.TBL_Pending2.DefaultView.Count = 1 Then
                Me.row_Pending2 = Me.TBL_Pending2.DefaultView(0).Row
                Me.row_Pending2.unit_price = CDec(vUnitPrice)
                Me.row_Pending2.amount = Math.Round(Me.row_Pending2.unit_price * Me.row_Pending2.qty, DefPrec1, MidpointRounding.AwayFromZero)

            ElseIf Me.TBL_Pending2.DefaultView.Count = 0 Then
                Me.row_Pending2 = Me.TBL_Pending2.NewtblItemDetailPoolRow

                Me.row_Pending2.unit_price = CDec(vUnitPrice)
                Me.row_Pending2.qty = CDec(vQty)
                Me.row_Pending2.amount = Math.Round(Me.row_Pending2.unit_price * Me.row_Pending2.qty, DefPrec1, MidpointRounding.AwayFromZero)
                Me.row_Pending2.Cost_Perct_Bundling = 0
                Me.row_Pending2.Overhead = 0
                Me.row_Pending2.Primary = -1
                Me.row_Pending2.primary_adjustment_type = -1
                Me.row_Pending2.primary_code_no = vPrimaryCodeNo
                Me.row_Pending2.primary_gain_account = -1
                Me.row_Pending2.primary_loss_account = -1
                Me.row_Pending2.Primary_transaksi = -1
                Me.row_Pending2.Primary_Transaksi_Asal = vPrimaryTransaksiAsal
                Me.row_Pending2.qty = vQty
                Me.row_Pending2.remarks = ""
                Me.row_Pending2.return_price = 0
                Me.row_Pending2.tgl_transaksi_asal = vTglTransaksiAcuan
                Me.row_Pending2.tipe_item_detail = vTipeItemDetail
                Me.row_Pending2.tipe_transaksi_asal = vTipeTransaksiAsal


                Me.TBL_Pending2.Rows.Add(Me.row_Pending2)
            Else
                MsgBox("Logic error: invalide tbl_pending2.defaultview.count", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Private Sub MasukkanUnitPrice_TblPending3(ByVal vPrimaryTransaksiAsal As Double, ByVal vPrimaryCodeNo As Double, _
    'ByVal vTipeTransaksiAsal As ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal, ByVal vUnitPrice As Decimal, _
    'ByVal vQty As Decimal, ByVal vTglTransaksiAcuan As Date, ByVal vTipeItemDetail As Integer, ByVal vRemarks As String)

    '    Try
    '        'masukkan ke tabel bundling
    '        Me.TBL_Pending3.DefaultView.RowFilter = "[Primary Transaksi Asal] = " + vPrimaryTransaksiAsal.ToString + _
    '        " AND [tipe transaksi asal] = " + CInt(vTipeTransaksiAsal).ToString + _
    '        " AND [primary code no] = " + vPrimaryCodeNo.ToString

    '        If Me.TBL_Pending3.DefaultView.Count = 1 Then
    '            Me.row_Pending2 = Me.TBL_Pending3.DefaultView(0).Row
    '            Me.row_Pending2.unit_price = CDec(vUnitPrice)
    '            Me.row_Pending2.amount = Math.Round(Me.row_Pending2.unit_price * Me.row_Pending2.qty, DefPrec1, MidpointRounding.AwayFromZero)
    '            Me.row_Pending2.remarks = vRemarks
    '        ElseIf Me.TBL_Pending3.DefaultView.Count = 0 Then
    '            Me.row_Pending2 = Me.TBL_Pending3.NewtblItemDetailPoolRow

    '            Me.row_Pending2.unit_price = CDec(vUnitPrice)
    '            Me.row_Pending2.qty = CDec(vQty)
    '            Me.row_Pending2.amount = Math.Round(Me.row_Pending2.unit_price * Me.row_Pending2.qty, DefPrec1, MidpointRounding.AwayFromZero)
    '            Me.row_Pending2.Cost_Perct_Bundling = 0
    '            Me.row_Pending2.Overhead = 0
    '            Me.row_Pending2.Primary = -1
    '            Me.row_Pending2.primary_adjustment_type = -1
    '            Me.row_Pending2.primary_code_no = vPrimaryCodeNo
    '            Me.row_Pending2.primary_gain_account = -1
    '            Me.row_Pending2.primary_loss_account = -1
    '            Me.row_Pending2.Primary_transaksi = -1
    '            Me.row_Pending2.Primary_Transaksi_Asal = vPrimaryTransaksiAsal
    '            Me.row_Pending2.qty = vQty
    '            Me.row_Pending2.remarks = vRemarks
    '            Me.row_Pending2.return_price = 0
    '            Me.row_Pending2.tgl_transaksi_asal = vTglTransaksiAcuan
    '            Me.row_Pending2.tipe_item_detail = vTipeItemDetail
    '            Me.row_Pending2.tipe_transaksi_asal = vTipeTransaksiAsal


    '            Me.TBL_Pending3.Rows.Add(Me.row_Pending2)
    '        Else
    '            MsgBox("Logic error: invalide TBL_Pending3.defaultview.count", MsgBoxStyle.Information, "Information")
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Dim DefPrec As Integer

        Try
            DefPrec = Me.ObjCur.GetDecimal_DefaultCurrency

            If e.ColumnIndex = Me.ColUnitPrice.Index Then
                If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If Me.ObjCD.CekNonNegatif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    Else
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, DefPrec)
                        Me.MasukkanUnitPrice(Me.dgvItem.Item(Me.ColPrimaryDocAcuan.Index, e.RowIndex).Value, Me.dgvItem.Item(Me.ColPrimaryCodeNoAdjust.Index, e.RowIndex).Value, _
                        Me.dgvItem.Item(Me.ColTipeTransaksiAsal.Index, e.RowIndex).Value, Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value, _
                         Me.dgvItem.Item(Me.ColQty.Index, e.RowIndex).Value, Me.dgvItem.Item(Me.ColTglDocAcuan.Index, e.RowIndex).Value, _
                         Me.dgvItem.Item(Me.ColTipeItemDetail.Index, e.RowIndex).Value)
                    End If
                Else
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.LangsungTutup = False
        Me.Close()
    End Sub

    Private Sub FrmTransaksiMonthlyClosing_CyclicBundlingExplain_LocationChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.LocationChanged

    End Sub
End Class