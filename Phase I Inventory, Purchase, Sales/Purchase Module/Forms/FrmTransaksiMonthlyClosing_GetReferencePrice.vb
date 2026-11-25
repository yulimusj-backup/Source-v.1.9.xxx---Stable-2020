Option Explicit On
Option Strict Off

Public Class FrmTransaksiMonthlyClosing_GetReferencePrice
    Public BackColorAll As Color = Color.LightGoldenrodYellow
    Public PrimaryProduct As Double = -1
    Public ClosingPeriod As Date = #1/1/1950#
    Public UnitPrice As Decimal = 0

    Dim LangsungClose As Boolean = False
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
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjCur As New Core_Module.clsMasterCurrency

    Private Sub FrmTransaksiMonthlyClosing_GetReferencePrice_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim rowDGV As DataGridViewRow

        Try
            rowDGV = Me.dgv1.Rows(Me.dgv1.SelectedCells(0).RowIndex)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTransaksiMonthlyClosing_GetReferencePrice_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim rowDGV As DataGridViewRow
        Dim DTA As New DseMonthlyClosingNewTableAdapters.DataTable7TableAdapter
        Dim TBL As New DseMonthlyClosingNew.DataTable7DataTable
        Dim row As DseMonthlyClosingNew.DataTable7Row
        Dim StartDate, EndDate As Date
        Dim DefPrec As Integer = 0

        Try
            DefPrec = Me.ObjCur.GetDecimal_DefaultCurrency
            Me.DTAAOP.Connection.ConnectionString = Core_Module.GetConstring
            Me.DTABundling.Connection.ConnectionString = Core_Module.GetConstring
            Me.DTAInv.Connection.ConnectionString = Core_Module.GetConstring
            Me.DTAStockInOutAdjust.Connection.ConnectionString = Core_Module.GetConstring
            Me.DTA_PR.Connection.ConnectionString = Core_Module.GetConstring
            DTA_StockTaking.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            DTA.FillBy_PrimaryProduct(TBL, Me.ClosingPeriod, CInt(ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem), Me.PrimaryProduct)
            StartDate = Date.MinValue
            EndDate = Date.MaxValue
            For Each row In TBL.Rows
                If StartDate.Date > row.tgl_transaksi_asal.Date Then
                    StartDate = row.tgl_transaksi_asal
                End If
                If EndDate.Date < row.tgl_transaksi_asal.Date Then
                    EndDate = row.tgl_transaksi_asal
                End If
            Next

            Me.DTAAOP.Fill(Me.TBLAOP, EndDate, StartDate)
            Me.DTABundling.Fill(Me.TblBundling, EndDate, StartDate)
            Me.DTAInv.Fill(Me.TBLInv, EndDate, StartDate)
            Me.DTAStockInOutAdjust.Fill(Me.TBLStockInOutAdjust, EndDate, StartDate)
            Me.DTA_PR.Fill(Me.TBL_PR, StartDate, EndDate)
            Me.DTA_StockTaking.Fill(Me.TBL_StockTaking)

            Me.dgv1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgv1.Rows.Clear()
            For Each row In TBL.Rows
                Me.dgv1.Rows.Add()
                rowDGV = Me.dgv1.Rows(Me.dgv1.RowCount - 1)
                rowDGV.Cells(Me.ColDate.Index).Value = Format(row.tgl_transaksi_asal, "dd MMM yyyy")
                rowDGV.Cells(Me.ColDocumentNo.Index).Value = Me.GetDocumentNo(row.primary_transaksi_asal, row.tipe_transaksi_asal)
                rowDGV.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(row.qty, row._Decimal)

                If row.tipe_transaksi_asal = CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InPurchase) Then
                    rowDGV.Cells(Me.ColReference.Index).Value = Me.GetSupplier(row.primary_transaksi_asal)
                ElseIf row.tipe_transaksi_asal = CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR) Then
                    rowDGV.Cells(Me.ColReference.Index).Value = Me.GetCustomer(row.primary_transaksi_asal)
                Else
                    rowDGV.Cells(Me.ColReference.Index).Value = ""
                End If

                rowDGV.Cells(Me.ColTipeDocument.Index).Value = Me.GetTipeDocument_TransaksiAsal(row.tipe_transaksi_asal)
                rowDGV.Cells(Me.ColUnitPrice.Index).Value = Me.ObjCD.DigitGrouping(row.unit_price, DefPrec)
            Next
            Me.dgv1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            Me.Timer1.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetSupplier(ByVal vPrimaryTransksiAsal As Double) As String
        Try
            Me.TBLAOP.DefaultView.RowFilter = "Primary = " + vPrimaryTransksiAsal.ToString
            If Me.TBLAOP.DefaultView.Count > 0 Then
                Me.rowAOP = Me.TBLAOP.DefaultView(0).Row
                Return Me.rowAOP.Supplier_Code + " - " + rowAOP.Supplier_Name
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function GetCustomer(ByVal vPrimaryTransksiAsal As Double) As String
        Try
            Me.TBLStockInOutAdjust.DefaultView.RowFilter = "Primary = " + vPrimaryTransksiAsal.ToString
            If Me.TBLStockInOutAdjust.DefaultView.Count > 0 Then
                Me.rowStockInOut = Me.TBLStockInOutAdjust.DefaultView(0).Row
                Return Me.rowStockInOut.No_Transaksi
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
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

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.DialogResult = Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim rowDGV As DataGridViewRow

        Try
            rowDGV = Me.dgv1.SelectedRows(0)
            If MsgBox("Use " + rowDGV.Cells(Me.ColUnitPrice.Index).Value.ToString + " as unit price?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.UnitPrice = CDec(rowDGV.Cells(Me.ColUnitPrice.Index).Value)
                Me.LangsungClose = True
                Me.DialogResult = Windows.Forms.DialogResult.OK
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Try
            Me.dgv1.ClearSelection()
            If Me.dgv1.RowCount > 0 Then
                Me.dgv1.Rows(0).Selected = True
            End If
            Me.Timer1.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class