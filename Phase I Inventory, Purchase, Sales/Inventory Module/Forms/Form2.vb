Option Explicit On
Option Strict Off

Public Class Form2
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim TipeProgress As EnumTipeProgress = EnumTipeProgress.AddDGVStock
    Dim DS As New DseLaporanStockDetail2
    Dim TBL As DseLaporanStockDetail2.DataTable1DataTable = DS.DataTable1
    Dim PrimaryProductTdkCocok As Double = -1
    Dim SudahLoadDGV As Boolean = False
    Dim TipePrg As Windows.Forms.ProgressBarStyle = ProgressBarStyle.Marquee
    Dim PrgBar As Windows.Forms.ProgressBar
    Dim TBLPositive1 As New DataSet2.tbltrans_custom_stockDataTable
    Dim TBLPositive2 As New DataSet2.tbltrans_custom_stockDataTable
    Dim TBLNegative As New DataSet2.tbltrans_custom_stockDataTable
    Dim objStockInOut As New clsTransaksiStockInOut

    Private Sub ChangeProgressBar2()
        Try
            Me.PrgBar.Style = Me.TipePrg
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChangeProgressBar()
        Try
            If Me.prb_Step1_2.Style = ProgressBarStyle.Continuous Then
                Me.prb_Step1_2.Style = ProgressBarStyle.Marquee
            Else
                Me.prb_Step1_2.Style = ProgressBarStyle.Continuous
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Enum EnumTipeProgress As Integer
        AddDGVStock
        AddStockTidakCocok
        ChangePrgStyle
        ChangePrgStyle2
    End Enum

    Private Sub btnStep1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStep1.Click
        Try
            If MsgBox("Move data from " + Format(Me.dtpStep1.Value, "dd MMM yyyy") + " to " + _
            Format(Me.dtpStep1.Value + TimeSpan.FromDays(-1), "dd MMM yyyy") + "?", MsgBoxStyle.YesNo, "confirm") = MsgBoxResult.Yes Then
                Control.CheckForIllegalCrossThreadCalls = False
                Me.bgwStep1.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwStep1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgwStep1.DoWork
        Dim DTA1 As New DataSet2TableAdapters.tbltransaksi_stockinoutTableAdapter
        Dim TBL1 As New DataSet2.tbltransaksi_stockinoutDataTable
        Dim row1 As DataSet2.tbltransaksi_stockinoutRow
        Dim DTA2 As New DataSet2TableAdapters.tbltransaksi_stockbundlingTableAdapter
        Dim TBL2 As New DataSet2.tbltransaksi_stockbundlingDataTable
        Dim row2 As DataSet2.tbltransaksi_stockbundlingRow
        Dim DTA3 As New DataSet2TableAdapters.tbltransaksi_stocktransferTableAdapter
        Dim TBL3 As New DataSet2.tbltransaksi_stocktransferDataTable
        Dim row3 As DataSet2.tbltransaksi_stocktransferRow
        Dim TipeSC As Inventory_Module.ClsTransaksiStockCard.EnumJenisTransaksiAsal
        Dim TipeStockInOut As clsTransaksiStockInOut.EnumTipeStockInOut
        Dim DTASC As New DataSet2TableAdapters.tbltransaksi_stock_cardTableAdapter
        Dim TBLSC As New DataSet2.tbltransaksi_stock_cardDataTable
        Dim TglAsal, TglTujuan As Date
        Dim DTA4 As New DataSet2_CekSelisihQtyTableAdapters.tbltrans_salesinvoiceTableAdapter

        Try
            TglAsal = Me.dtpStep1.Value.Date
            TglTujuan = Me.dtpStep1.Value.Date - TimeSpan.FromDays(1)

            '1.1. load data stock in out 
            Me.ObjCD.UpdateThreading()
            DTA1.Fill(TBL1, TglAsal)
            Me.prgStep1.Value = 0
            Me.prgStep1.Maximum = TBL1.Rows.Count
            If TBL1.Rows.Count > 0 Then
                For Each row1 In TBL1.Rows
                    'MsgBox("stock in out:" + TBL1.Rows.Count.ToString)
                    'Exit For

                    Me.prgStep1.Value += 1
                    Me.LblPrg_Step1.Text = "Fase 1 - Stock In Out - Data " + Me.prgStep1.Value.ToString + " of " + Me.prgStep1.Maximum.ToString

                    '1.2. pindahkan stock card
                    TipeStockInOut = row1.Tipe_Transaksi
                    If TipeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust Then
                        TipeSC = ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockIn_Adjust
                    ElseIf TipeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_BackProcess Then
                        MsgBox("Invalid data")
                    ElseIf TipeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Process Then
                        MsgBox("Invalid data")
                    ElseIf TipeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase Then
                        TipeSC = ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockIn_Purchase
                    ElseIf TipeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn Then
                        TipeSC = ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockIn_SalesReturn
                    ElseIf TipeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust Then
                        TipeSC = ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockOut_Adjust
                    ElseIf TipeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Process Then
                        MsgBox("Invalid data")
                    ElseIf TipeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn Then
                        TipeSC = ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockOut_PurchaseReturn
                    ElseIf TipeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Sales Then
                        TipeSC = ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockOut_Sales
                    End If
                    DTASC.UpdateQuery(TglTujuan, row1.Primary, TipeSC)
                    '1.3. pindahkan tanggal
                    DTA1.UpdateQuery_SetTgl(TglTujuan, row1.Primary)
                Next

            End If

            '2.1. load data stock bundling
            DTA2.Fill(TBL2, TglAsal)
            Me.prgStep1.Value = 0
            Me.prgStep1.Maximum = TBL2.Rows.Count
            If TBL2.Rows.Count > 0 Then
                For Each row2 In TBL2.Rows
                    'MsgBox("stock bundling:" + TBL2.Rows.Count.ToString)
                    'Exit For

                    Me.prgStep1.Value += 1
                    Me.LblPrg_Step1.Text = "Fase 2 - Stock Bundling - Data " + Me.prgStep1.Value.ToString + " of " + Me.prgStep1.Maximum.ToString

                    '2.2. pindahkan stock card
                    TipeSC = ClsTransaksiStockCard.EnumJenisTransaksiAsal.Stock_Bundling
                    DTASC.UpdateQuery(TglTujuan, row2.Primary, TipeSC)
                    '2.3. pindahkan tanggal
                    DTA2.UpdateQuery_SetTgl(TglTujuan, row2.Primary)
                Next
            End If

            '3.1. load data stock transfer
            DTA3.Fill(TBL3, TglAsal)
            Me.prgStep1.Value = 0
            Me.prgStep1.Maximum = TBL3.Rows.Count
            If TBL3.Rows.Count > 0 Then
                For Each row3 In TBL3.Rows
                    'MsgBox("stock transfer:" + TBL3.Rows.Count.ToString)
                    'Exit For

                    Me.prgStep1.Value += 1
                    Me.LblPrg_Step1.Text = "Fase 3 - Stock Transfer - Data " + Me.prgStep1.Value.ToString + " of " + Me.prgStep1.Maximum.ToString
                    '3.2. pindahkan stock card - ga ada
                    '3.3. pindahkan tanggal
                    DTA3.UpdateQuery_SetTgl(TglTujuan, row3.Primary)
                Next
            End If

            'final data check
            DTASC.FillBy_Tgl(TBLSC, TglAsal)
            If TBLSC.Rows.Count > 0 Then
                MsgBox("Logic error, please check again.", MsgBoxStyle.Information, "Information")
            End If

            'pindahkan semua data sales invoice
            DTA4.UpdateQuery_SetTglInv(TglTujuan, TglAsal)
            'pindahkan semua data sales return
            DTA4.UpdateQuery_SetTglSR(TglTujuan, TglAsal)
            'pindahkan semua data aop
            DTA4.UpdateQuery_SetTglAOP(TglTujuan, TglAsal)
            'pindahkan semua data purchase return
            DTA4.UpdateQuery_SetTglPR(TglTujuan, TglAsal)
            'pindahakan semua data kartu ap/ar
            DTA4.UpdateQuery_SetTglKartuAPAR(TglTujuan, TglAsal)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwStep1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgwStep1.RunWorkerCompleted
        Try
            MsgBox("Data updated", MsgBoxStyle.Information, "Information")
            Me.LblPrg_Step1.Text = "Done"
            Me.prgStep1.Value = 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnShowOverhead_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShowOverhead.Click
        Dim PrimaryTrans As Double = -1
        Dim DTA As New DataSet2TableAdapters.tbltrans_purchase_closing_item_cost_group_accountTableAdapter
        Dim TBL As New DataSet2.tbltrans_purchase_closing_item_cost_group_accountDataTable
        Dim row As DataSet2.tbltrans_purchase_closing_item_cost_group_accountRow
        Dim DTA1 As New DataSet2TableAdapters.tbltrans_purchase_closing_item_cost_groupTableAdapter
        Dim TBL1 As New DataSet2.tbltrans_purchase_closing_item_cost_groupDataTable
        Dim row1 As DataSet2.tbltrans_purchase_closing_item_cost_groupRow
        Dim rowDGV As DataGridViewRow

        Try
            'konfirmasi ulang
            If MsgBox("reload dgv?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Windows.Forms.Cursor.Current = Cursors.WaitCursor

                Me.dgvOverhead.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.dgvOverhead.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                'kosongkan dgv
                Me.dgvOverhead.Rows.Clear()

                'get primary sesuai period
                If DTA1.ScalarQuery_GetPrimary(Me.dtpPeriodOverhead.Value.Date).HasValue = True Then
                    PrimaryTrans = DTA1.ScalarQuery_GetPrimary(Me.dtpPeriodOverhead.Value.Date).Value

                    'load data sesuai primary
                    DTA1.Fill(TBL1, PrimaryTrans)
                    DTA.Fill(TBL, PrimaryTrans)
                    'tampilkan ke dalam dgv
                    For Each row1 In TBL1.Rows
                        TBL.DefaultView.RowFilter = "identifier = '" + row1.identifier + "'"
                        If TBL.DefaultView.Count = 0 Then
                            'tambahkan 1 baris kosong saja
                            Me.dgvOverhead.Rows.Add()
                            rowDGV = Me.dgvOverhead.Rows(Me.dgvOverhead.Rows.Count - 1)
                            rowDGV.Cells(Me.ColOv_AccCode.Index).Value = ""
                            rowDGV.Cells(Me.ColOv_AccDesc.Index).Value = ""
                            rowDGV.Cells(Me.ColOv_AllocatedPerc.Index).Value = Me.ObjCD.DigitGrouping(0, 2)
                            rowDGV.Cells(Me.ColOv_Balance.Index).Value = Me.ObjCD.DigitGrouping(0, 2)
                            rowDGV.Cells(Me.ColOv_Desc.Index).Value = row1.desc
                            rowDGV.Cells(Me.ColOv_No.Index).Value = Me.dgvOverhead.Rows.Count
                            rowDGV.Cells(Me.ColOv_RefUnit.Index).Value = row1.Unit_Code
                        Else
                            'tambah baris sejumlah baris yang kosong
                            For Each DV As Data.DataRowView In TBL.DefaultView
                                row = DV.Row
                                Me.dgvOverhead.Rows.Add()
                                rowDGV = Me.dgvOverhead.Rows(Me.dgvOverhead.Rows.Count - 1)
                                rowDGV.Cells(Me.ColOv_AccCode.Index).Value = row.Account_Code
                                rowDGV.Cells(Me.ColOv_AccDesc.Index).Value = row.Account_Desc
                                rowDGV.Cells(Me.ColOv_AllocatedPerc.Index).Value = Me.ObjCD.DigitGrouping(row.perc_allocated, 2)
                                rowDGV.Cells(Me.ColOv_Balance.Index).Value = Me.ObjCD.DigitGrouping(0, 2)
                                rowDGV.Cells(Me.ColOv_Desc.Index).Value = row1.desc
                                rowDGV.Cells(Me.ColOv_No.Index).Value = Me.dgvOverhead.Rows.Count
                                rowDGV.Cells(Me.ColOv_RefUnit.Index).Value = row1.Unit_Code
                            Next
                        End If
                    Next
                End If

                Me.dgvOverhead.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                Me.dgvOverhead.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Form2_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.dtpPeriodOverhead.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriodOverhead.Value)
            Me.dtpStep14.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpStep14.Value)
            Me.prb_Step1_2.Style = ProgressBarStyle.Continuous
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnStep1_2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStep1_2.Click
        Try
            If MsgBox("Start collecting stock data on " + Format(Me.dtpResetStock.Value.Date, "dd MMM yyyy") + "?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Control.CheckForIllegalCrossThreadCalls = False
                Me.bgwStep1_2.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Enum EnumScope
        Goods_Only
        Service_Only
        Both
    End Enum

    Public Function CetakanStockDetail() As Boolean
        Dim DTA1 As New DseLaporanStockDetail2TableAdapters.tblmaster_productTableAdapter
        Dim TBL1 As New DseLaporanStockDetail2.tblmaster_productDataTable
        Dim row1, row2 As DseLaporanStockDetail2.tblmaster_productRow
        Dim DTA2 As New DataSet2TableAdapters.tbltrans_purchase_closing_item_cost_groupTableAdapter
        Dim TglAcuan As Date = #11/30/2017#
        Dim CurIdx, TargetIdx, BatchQty, FillIdx1, FillIdx2 As Integer
        Dim dtaSC As New DseLaporanStockDetail2TableAdapters.DataTable5TableAdapter
        Dim tblSC As New DseLaporanStockDetail2.DataTable5DataTable
        Dim tblSC1 As New DseLaporanStockDetail2.DataTable5DataTable
        Dim rowSC, rowSC1 As DseLaporanStockDetail2.DataTable5Row
        Dim dtaLTF As New DseLaporanStockDetail2TableAdapters.DataTable6TableAdapter
        Dim tblLTF As New DseLaporanStockDetail2.DataTable6DataTable
        Dim tblLTF1 As New DseLaporanStockDetail2.DataTable6DataTable
        Dim rowLTF, rowLTF1 As DseLaporanStockDetail2.DataTable6Row
        Dim DTA_ST As New DseLaporanStockDetail2TableAdapters.DataTable7TableAdapter
        Dim TBL_ST As New DseLaporanStockDetail2.DataTable7DataTable
        Dim TBL_ST1 As New DseLaporanStockDetail2.DataTable7DataTable
        Dim row_ST, row_ST1 As DseLaporanStockDetail2.DataTable7Row
        Dim PrimaryMax, PrimaryMin As Double
        Dim DTA As New DseLaporanStockDetail2TableAdapters.DataTable1TableAdapter
        Dim TBL_Prod As New DseLaporanStockDetail2.DataTable1DataTable
        Dim TBL_Prod1 As New DseLaporanStockDetail2.DataTable1DataTable
        Dim row_Prod, row_Prod1 As DseLaporanStockDetail2.DataTable1Row

        Try
            FillIdx1 = -1
            FillIdx2 = -1

            Me.ObjCD.UpdateThreading()
            Me.TBL.Rows.Clear()
            CurIdx = 0
            BatchQty = 99
            TglAcuan = Me.dtpResetStock.Value.Date

            DTA.Connection.ConnectionString = Core_Module.GetConstring
            dtaSC.Connection.ConnectionString = Core_Module.GetConstring
            dtaLTF.Connection.ConnectionString = Core_Module.GetConstring
            DTA_ST.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = Core_Module.GetConstring

            Me.SudahLoadDGV = False
            DTA1.Fill(TBL1)
            Me.prb_Step1_2a.Value = 0
            Me.prb_Step1_2a.Maximum = TBL1.Rows.Count
            Me.lblPrgStep1_2a.Text = ""
            Me.TBL.Rows.Clear()

            'isikan per product code
            While CurIdx <= TBL1.Rows.Count - 1
                If CurIdx + BatchQty <= TBL1.Rows.Count - 1 Then
                    TargetIdx = CurIdx + BatchQty
                Else
                    TargetIdx = TBL1.Rows.Count - 1
                End If
                If CurIdx > FillIdx2 Then
                    FillIdx1 = CurIdx
                    FillIdx2 = (BatchQty * 10) + 9 + FillIdx1
                    If FillIdx2 > TBL1.Rows.Count - 1 Then
                        FillIdx2 = TBL1.Rows.Count - 1
                    End If
                    row1 = TBL1.Rows(FillIdx1)
                    row2 = TBL1.Rows(FillIdx2)

                    Me.lblPrgStep1_2a.Text = "Load detail product qty"
                    TBL_Prod.Rows.Clear()
                    DTA.FillBy_AllGoods2(TBL_Prod, row2.Primary, row1.Primary)
                    Me.lblPrgStep1_2a.Text = "Load detail stock card"
                    tblSC.Rows.Clear()
                    dtaSC.FillBy_AllGoods2(tblSC, row2.Primary, row1.Primary, TglAcuan.Date.AddDays(1), CInt(ClsTransaksiStockCard.EnumJenisTransaksiAsal.Stock_Taking))
                    Me.lblPrgStep1_2a.Text = "Load detail stock transfer"
                    tblLTF.Rows.Clear()
                    dtaLTF.FillBy_AllGoods2(tblLTF, row2.Primary, row1.Primary, TglAcuan.Date.AddDays(1))
                    Me.lblPrgStep1_2a.Text = "Load detail stock taking"
                    TBL_ST.Rows.Clear()
                    DTA_ST.FillBy_AllGoods2(TBL_ST, TglAcuan.Date.AddDays(1), row2.Primary, row1.Primary)
                End If

                Me.prb_Step1_2a.Value += (TargetIdx - CurIdx)
                row1 = TBL1.Rows(CurIdx)
                row2 = TBL1.Rows(TargetIdx)
                TBL_Prod1.Rows.Clear()
                TBL_ST1.Rows.Clear()
                tblLTF1.Rows.Clear()
                tblSC1.Rows.Clear()
               

                Me.lblPrgStep1_2a.Text = "Data " + (TargetIdx + 1).ToString + " of " + Me.prb_Step1_2a.Maximum.ToString + " - " + row1.Product_Code

                Me.lblPrgStep1_2a.Text = "transfer detail product qty"
                TBL_Prod.DefaultView.RowFilter = "[Primary Product] <= " + row2.Primary.ToString + " AND [Primary Product] >= " + row1.Primary.ToString
                Me.prb_Step1_2.Value = 0
                Me.prb_Step1_2.Maximum = TBL_Prod.DefaultView.Count
                While TBL_Prod.DefaultView.Count > 0
                    Me.prb_Step1_2.Value += 1
                    If Me.prb_Step1_2.Value = Me.prb_Step1_2.Maximum Or Math.IEEERemainder(Me.prb_Step1_2.Value, 100) = 0 Then
                        lblPrgStep1_2.Text = "Data. " & Me.prb_Step1_2.Value.ToString & " of " & Me.prb_Step1_2.Maximum.ToString & "."
                    End If
                    row_Prod = TBL_Prod.DefaultView(0).Row
                    row_Prod1 = TBL_Prod1.NewDataTable1Row
                    For gh As Integer = 0 To TBL_Prod.Columns.Count - 1
                        row_Prod1.Item(gh) = row_Prod.Item(gh)
                    Next
                    TBL_Prod1.Rows.Add(row_Prod1)
                    TBL_Prod.Rows.Remove(row_Prod)
                End While
                Me.lblPrgStep1_2a.Text = "transfer detail stock taking"
                TBL_ST.DefaultView.RowFilter = "[Primary Product] <= " + row2.Primary.ToString + " AND [Primary Product] >= " + row1.Primary.ToString
                Me.prb_Step1_2.Value = 0
                Me.prb_Step1_2.Maximum = TBL_ST.DefaultView.Count
                While TBL_ST.DefaultView.Count > 0
                    Me.prb_Step1_2.Value += 1
                    If Me.prb_Step1_2.Value = Me.prb_Step1_2.Maximum Or Math.IEEERemainder(Me.prb_Step1_2.Value, 100) = 0 Then
                        lblPrgStep1_2.Text = "Data. " & Me.prb_Step1_2.Value.ToString & " of " & Me.prb_Step1_2.Maximum.ToString & "."
                    End If
                    row_ST = TBL_ST.DefaultView(0).Row
                    row_ST1 = TBL_ST1.NewDataTable7Row
                    For gh As Integer = 0 To TBL_ST.Columns.Count - 1
                        row_ST1.Item(gh) = row_ST.Item(gh)
                    Next
                    TBL_ST1.Rows.Add(row_ST1)
                    TBL_ST.Rows.Remove(row_ST)
                End While
                Me.lblPrgStep1_2a.Text = "transfer detail stock transfer"
                tblLTF.DefaultView.RowFilter = "[Primary Product] <= " + row2.Primary.ToString + " AND [Primary Product] >= " + row1.Primary.ToString
                Me.prb_Step1_2.Value = 0
                Me.prb_Step1_2.Maximum = tblLTF.DefaultView.Count
                While tblLTF.DefaultView.Count > 0
                    Me.prb_Step1_2.Value += 1
                    If Me.prb_Step1_2.Value = Me.prb_Step1_2.Maximum Or Math.IEEERemainder(Me.prb_Step1_2.Value, 100) = 0 Then
                        lblPrgStep1_2.Text = "Data. " & Me.prb_Step1_2.Value.ToString & " of " & Me.prb_Step1_2.Maximum.ToString & "."
                    End If
                    rowLTF = tblLTF.DefaultView(0).Row
                    rowLTF1 = tblLTF1.NewDataTable6Row
                    For gh As Integer = 0 To tblLTF.Columns.Count - 1
                        rowLTF1.Item(gh) = rowLTF.Item(gh)
                    Next
                    tblLTF1.Rows.Add(rowLTF1)
                    tblLTF.Rows.Remove(rowLTF)
                End While

                Me.lblPrgStep1_2a.Text = "transfer detail stock card"
                tblSC.DefaultView.RowFilter = "[Primary Product] <= " + row2.Primary.ToString + " AND [Primary Product] >= " + row1.Primary.ToString
                Me.prb_Step1_2.Value = 0
                Me.prb_Step1_2.Maximum = tblSC.DefaultView.Count
                While tblSC.DefaultView.Count > 0
                    Me.prb_Step1_2.Value += 1
                    If Me.prb_Step1_2.Value = Me.prb_Step1_2.Maximum Or Math.IEEERemainder(Me.prb_Step1_2.Value, 100) = 0 Then
                        lblPrgStep1_2.Text = "Data. " & Me.prb_Step1_2.Value.ToString & " of " & Me.prb_Step1_2.Maximum.ToString & "."
                    End If
                    rowSC = tblSC.DefaultView(0).Row
                    rowSC1 = tblSC1.NewDataTable5Row
                    For gh As Integer = 0 To tblSC.Columns.Count - 1
                        rowSC1.Item(gh) = rowSC.Item(gh)
                    Next
                    tblSC1.Rows.Add(rowSC1)
                    tblSC.Rows.Remove(rowSC)
                End While

                Me.lblPrgStep1_2a.Text = "Data " + (TargetIdx + 1).ToString + " of " + Me.prb_Step1_2a.Maximum.ToString + " - " + row1.Product_Code
                Me.CetakanStockDetail2(TBL_Prod1, TglAcuan, tblSC1, tblLTF1, TBL_ST1)
                Me.prb_Step1_2.Value = 0
                Me.prb_Step1_2.Maximum = TBL_Prod1.Rows.Count
                For Each row_Prod In TBL_Prod1.Rows
                    Me.prb_Step1_2.Value += 1
                    If Me.prb_Step1_2.Value = Me.prb_Step1_2.Maximum Or Math.IEEERemainder(Me.prb_Step1_2.Value, 100) = 0 Then
                        lblPrgStep1_2.Text = "Moving result. " & Me.prb_Step1_2.Value.ToString & " of " & Me.prb_Step1_2.Maximum.ToString & "."
                    End If
                    row_Prod1 = Me.TBL.NewDataTable1Row
                    For jk As Integer = 0 To Me.TBL.Columns.Count - 1
                        row_Prod1.Item(jk) = row_Prod.Item(jk)
                    Next
                    Me.TBL.Rows.Add(row_Prod1)
                Next
                CurIdx = TargetIdx + 1
            End While

            'load ke dgv
            Me.dgvStock.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvStock.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.dgvStockTidakCocok.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.dgvStockTidakCocok.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvStockMinus.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.dgvStockMinus.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvStock.Rows.Clear()
            Me.dgvStockTidakCocok.Rows.Clear()
            Me.dgvStockMinus.Rows.Clear()

            'Me.TipeProgress = EnumTipeProgress.ChangePrgStyle
            'Me.bgwStep1_2.ReportProgress(0)

            Me.TipeProgress = EnumTipeProgress.AddDGVStock
            Me.bgwStep1_2.ReportProgress(0)
            lblPrgStep1_2.Text = "Load data to DGV..."
            While Me.SudahLoadDGV = False
                'lblPrgStep1_2.Text = "Load data to DGV - data " + Me.dgvStock.Rows.Count.ToString + " of " + TBL.Rows.Count.ToString
                Threading.Thread.Sleep(60000)
            End While

            'Me.TipeProgress = EnumTipeProgress.ChangePrgStyle
            'Me.bgwStep1_2.ReportProgress(0)
            'pengecekan stock tidak cocok
            'If MsgBox("Check unmatched data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
            Me.CekTidakCocok(TBL, TglAcuan)
            'End If

            'simpan ke db
            'If MsgBox("Simpan ke db?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
            'hapus db
            Me.prb_Step1_2.Value = 0
            Me.lblPrgStep1_2.Text = "Deleting database on tbl item stock"
            DTA2.DeleteQuery()
            'simpan item
            Me.Simpan_Item_UsingScript(TBL)
            'End If

            Me.lblPrgStep1_2.Text = "Finished"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function CetakanStockDetail2(ByVal TBL_Prod As DseLaporanStockDetail2.DataTable1DataTable, ByVal TglAcuan As Date, _
 ByVal tblSC As DseLaporanStockDetail2.DataTable5DataTable, ByVal tblLTF As DseLaporanStockDetail2.DataTable6DataTable, _
 ByVal TBL_ST As DseLaporanStockDetail2.DataTable7DataTable) As Boolean
        Dim Row As DseLaporanStockDetail2.DataTable1Row
        Dim rowSC As DseLaporanStockDetail2.DataTable5Row
        Dim rowLTF As DseLaporanStockDetail2.DataTable6Row
        Dim row_ST As DseLaporanStockDetail2.DataTable7Row

        Try
            Me.prb_Step1_2.Value = 0
            Me.prb_Step1_2.Maximum = tblSC.Rows.Count
            While tblSC.Rows.Count > 0
                rowSC = tblSC.Rows(0)
                Me.prb_Step1_2.Value += 1
                If Me.prb_Step1_2.Value = Me.prb_Step1_2.Maximum Or Math.IEEERemainder(Me.prb_Step1_2.Value, 100) = 0 Then
                    lblPrgStep1_2.Text = "Stock Card calculation... " & Me.prb_Step1_2.Value.ToString & " of " & Me.prb_Step1_2.Maximum.ToString & "."
                End If
                TBL_Prod.DefaultView.RowFilter = "[Primary Product] = " & rowSC.Primary_Product.ToString & " AND [Lot No] = '" & Me.ObjCD.FormatSpecialChar_DataFilter(rowSC.Lot_No) & "' AND [Primary Storage Area] = " & rowSC.Primary_Storage_Area.ToString
                If TBL_Prod.DefaultView.Count > 0 Then
                    Row = TBL_Prod.DefaultView(0).Row
                    Row.Qty = Row.Qty + rowSC._Total__Out_ - rowSC.Total_In
                    If Row.Qty = 0 Then
                        TBL_Prod.Rows.Remove(Row)
                    End If
                Else
                    Row = TBL_Prod.NewDataTable1Row
                    Row.Primary_Product = rowSC.Primary_Product
                    Row.Product_Code = rowSC.Product_Code
                    Row.Description = rowSC.Description
                    Row._Decimal = rowSC._Decimal
                    Row.Batch_No_Controlled = rowSC.Batch_No_Controlled
                    Row.Location_Controlled = rowSC.Location_Controlled
                    Row.Expiry_Date_Controlled = rowSC.Expiry_Date_Controlled
                    Row.Type = rowSC.Type
                    Row.Qty = rowSC._Total__Out_ - rowSC.Total_In
                    Row.Primary_Unit = rowSC.Primary_Unit
                    Row.Unit_Code = rowSC.Unit_Code
                    Row.Category_Primary = rowSC.Category_Primary
                    Row.Categorey_Code = rowSC.Categorey_Code
                    Row.Category_Desc = rowSC.Categrory_Desc
                    Row.Sub_Brand_Primary = rowSC.Sub_Brand_Primary
                    Row.Sub_Brand_Code = rowSC.Sub_Brand_Code
                    Row.Sub_Brand_Desc = rowSC.Sub_Brand_Desc
                    Row.Brand_Primary = rowSC.Brand_Primary
                    Row.Brand_Code = rowSC.Brand_Code
                    Row.Brand_Desc = rowSC.Brand_Desc
                    Row.Group_Primary = rowSC.Group_Primary
                    Row.Group_Code = rowSC.Group_Code
                    Row.Group_Desc = rowSC.Group_Desc
                    Row.Expiry_Date = rowSC.Expiry_Date
                    Row.Primary_Storage_Area = rowSC.Primary_Storage_Area
                    Row.Storage_Area_Code = rowSC.Storage_Area_Code
                    Row.Storage_Area_Desc = rowSC.Storage_Area_Desc
                    Row.Lot_No = rowSC.Lot_No
                    Row.Qty_Secondary = -1
                    Row.Warehouse_Code = rowSC.Warehouse_Code
                    Row.Warehouse_Desc = rowSC.Warehouse_Desc

                    If Row.Qty <> 0 Then
                        TBL_Prod.Rows.Add(Row)
                    End If
                End If
                tblSC.Rows.Remove(rowSC)
            End While

            lblPrgStep1_2.Text = "Preparing for stock transfer reversal...."
            Me.prb_Step1_2.Value = 0
            Me.prb_Step1_2.Maximum = tblLTF.Rows.Count
            While tblLTF.Rows.Count > 0
                Me.prb_Step1_2.Value += 1
                If Me.prb_Step1_2.Value = Me.prb_Step1_2.Maximum Or Math.IEEERemainder(Me.prb_Step1_2.Value, 100) = 0 Then
                    lblPrgStep1_2.Text = "Stock Transfer Calculation... " & Me.prb_Step1_2.Value.ToString & " of " & Me.prb_Step1_2.Maximum.ToString & "."
                End If

                rowLTF = tblLTF.Rows(0)
                TBL_Prod.DefaultView.RowFilter = "[Primary Product] = " & rowLTF.Primary_Product.ToString & " AND [Lot No] = '" & Me.ObjCD.FormatSpecialChar_DataFilter(rowLTF.Lot_No) & "' AND [Primary Storage Area] = " & rowLTF.Primary_Storage_Area_II.ToString & ""
                If TBL_Prod.DefaultView.Count > 0 Then
                    Row = TBL_Prod.DefaultView(0).Row
                    Row.Qty = Row.Qty - rowLTF.Qty
                    If Row.Qty = 0 Then
                        TBL_Prod.Rows.Remove(Row)
                    End If
                Else
                    If rowLTF.Qty > 0 Then
                        Row = TBL_Prod.NewDataTable1Row
                        Row.Primary_Product = rowLTF.Primary_Product
                        Row.Product_Code = rowLTF.Product_Code
                        Row.Description = rowLTF.Description
                        Row._Decimal = rowLTF._Decimal
                        Row.Batch_No_Controlled = rowLTF.Batch_No_Controlled
                        Row.Location_Controlled = rowLTF.Location_Controlled
                        Row.Expiry_Date_Controlled = rowLTF.Expiry_Date_Controlled
                        Row.Type = rowLTF.Type
                        Row.Qty = rowLTF.Qty * -1
                        Row.Primary_Unit = rowLTF.Primary_Unit
                        Row.Unit_Code = rowLTF.Unit_Code
                        Row.Category_Primary = rowLTF.Category_Primary
                        Row.Categorey_Code = rowLTF.Categorey_Code
                        Row.Category_Desc = rowLTF.Category_Desc
                        Row.Sub_Brand_Primary = rowLTF.Sub_Brand_Primary
                        Row.Sub_Brand_Code = rowLTF.Sub_Brand_Code
                        Row.Sub_Brand_Desc = rowLTF.Sub_Brand_Desc
                        Row.Brand_Primary = rowLTF.Brand_Primary
                        Row.Brand_Code = rowLTF.Brand_Code
                        Row.Brand_Desc = rowLTF.Brand_Desc
                        Row.Group_Primary = rowLTF.Group_Primary
                        Row.Group_Code = rowLTF.Group_Code
                        Row.Group_Desc = rowLTF.Group_Desc
                        Row.Expiry_Date = rowLTF.Expiry_Date
                        Row.Primary_Storage_Area = rowLTF.Primary_Storage_Area_II
                        Row.Storage_Area_Code = rowLTF.Storage_Area_Code_II
                        Row.Storage_Area_Desc = rowLTF.Storage_Area_Desc_II
                        Row.Lot_No = rowLTF.Lot_No
                        Row.Qty_Secondary = -1
                        Row.Warehouse_Code = rowLTF.Warehouse_Code_II
                        Row.Warehouse_Desc = rowLTF.Warehouse_Desc_II
                        If Row.Qty <> 0 Then
                            TBL_Prod.Rows.Add(Row)
                        End If
                    End If
                End If

                TBL_Prod.DefaultView.RowFilter = "[Primary Product] = " & rowLTF.Primary_Product.ToString & " AND [Lot No] = '" & Me.ObjCD.FormatSpecialChar_DataFilter(rowLTF.Lot_No) & "' AND [Primary Storage Area] = " & rowLTF.Primary_Storage_Area.ToString & ""
                If TBL_Prod.DefaultView.Count > 0 Then
                    Row = TBL_Prod.DefaultView(0).Row
                    Row.Qty = Row.Qty + rowLTF.Qty
                    If Row.Qty = 0 Then
                        TBL_Prod.Rows.Remove(Row)
                    End If
                Else
                    If rowLTF.Qty > 0 Then
                        Row = TBL_Prod.NewDataTable1Row
                        Row.Primary_Product = rowLTF.Primary_Product
                        Row.Product_Code = rowLTF.Product_Code
                        Row.Description = rowLTF.Description
                        Row._Decimal = rowLTF._Decimal
                        Row.Batch_No_Controlled = rowLTF.Batch_No_Controlled
                        Row.Location_Controlled = rowLTF.Location_Controlled
                        Row.Expiry_Date_Controlled = rowLTF.Expiry_Date_Controlled
                        Row.Type = rowLTF.Type
                        Row.Qty = rowLTF.Qty
                        Row.Primary_Unit = rowLTF.Primary_Unit
                        Row.Unit_Code = rowLTF.Unit_Code
                        Row.Category_Primary = rowLTF.Category_Primary
                        Row.Categorey_Code = rowLTF.Categorey_Code
                        Row.Category_Desc = rowLTF.Categorey_Code
                        Row.Sub_Brand_Primary = rowLTF.Sub_Brand_Primary
                        Row.Sub_Brand_Code = rowLTF.Sub_Brand_Code
                        Row.Sub_Brand_Desc = rowLTF.Sub_Brand_Desc
                        Row.Brand_Primary = rowLTF.Brand_Primary
                        Row.Brand_Code = rowLTF.Brand_Code
                        Row.Brand_Desc = rowLTF.Brand_Desc
                        Row.Group_Primary = rowLTF.Group_Primary
                        Row.Group_Code = rowLTF.Group_Code
                        Row.Group_Desc = rowLTF.Group_Desc
                        Row.Expiry_Date = rowLTF.Expiry_Date
                        Row.Primary_Storage_Area = rowLTF.Primary_Storage_Area
                        Row.Storage_Area_Code = rowLTF.Storage_Area_Code
                        Row.Storage_Area_Desc = rowLTF.Storage_Area_Desc
                        Row.Lot_No = rowLTF.Lot_No
                        Row.Qty_Secondary = -1
                        Row.Warehouse_Code = rowLTF.Warehouse_Code
                        Row.Warehouse_Desc = rowLTF.Warehouse_Desc
                        If Row.Qty <> 0 Then
                            TBL_Prod.Rows.Add(Row)
                        End If
                    End If
                End If
                tblLTF.Rows.Remove(rowLTF)
            End While

            Me.prb_Step1_2.Value = 0
            Me.prb_Step1_2.Maximum = TBL_ST.Rows.Count
            While TBL_ST.Rows.Count > 0
                row_ST = TBL_ST.Rows(0)
                Me.prb_Step1_2.Value += 1
                If Me.prb_Step1_2.Value = Me.prb_Step1_2.Maximum Or Math.IEEERemainder(Me.prb_Step1_2.Value, 100) = 0 Then
                    lblPrgStep1_2.Text = "Stock Taking calculation... " & Me.prb_Step1_2.Value.ToString & " of " & Me.prb_Step1_2.Maximum.ToString & "."
                End If
                TBL_Prod.DefaultView.RowFilter = "[Primary Product] = " & row_ST.Primary_Product.ToString & " AND [Lot No] = '" & Me.ObjCD.FormatSpecialChar_DataFilter(row_ST.Batch_No) & "' AND [Primary Storage Area] = " & row_ST.Primary_Location.ToString
                If TBL_Prod.DefaultView.Count > 0 Then
                    Row = TBL_Prod.DefaultView(0).Row
                    Row.Qty = Row.Qty - row_ST.Qty
                    If Row.Qty = 0 Then
                        TBL_Prod.Rows.Remove(Row)
                    End If
                Else
                    Row = TBL_Prod.NewDataTable1Row
                    Row.Primary_Product = row_ST.Primary_Product
                    Row.Product_Code = row_ST.Product_Code
                    Row.Description = row_ST.Description
                    Row._Decimal = row_ST._Decimal
                    Row.Batch_No_Controlled = row_ST.Batch_No_Controlled
                    Row.Location_Controlled = row_ST.Location_Controlled
                    Row.Expiry_Date_Controlled = row_ST.Expiry_Date_Controlled
                    Row.Type = row_ST.Type
                    Row.Qty = row_ST.Qty * (-1)
                    Row.Primary_Unit = row_ST.Primary_Unit
                    Row.Unit_Code = row_ST.Unit_Code
                    Row.Category_Primary = row_ST.Category_Primary
                    Row.Categorey_Code = row_ST.Category_Code
                    Row.Category_Desc = row_ST.Category_Desc
                    Row.Sub_Brand_Primary = row_ST.Sub_Brand_Primary
                    Row.Sub_Brand_Code = row_ST.Sub_Brand_Code
                    Row.Sub_Brand_Desc = row_ST.Sub_Brand_Desc
                    Row.Brand_Primary = row_ST.Brand_Primary
                    Row.Brand_Code = row_ST.Brand_Code
                    Row.Brand_Desc = row_ST.Brand_Desc
                    Row.Group_Primary = row_ST.Group_Primary
                    Row.Group_Code = row_ST.Product_Group_Code
                    Row.Group_Desc = row_ST.Product_Group_Desc
                    Row.Expiry_Date = row_ST.Expiry_Date
                    Row.Primary_Storage_Area = row_ST.Primary_Location
                    Row.Storage_Area_Code = row_ST.Storage_Area_Code
                    Row.Storage_Area_Desc = row_ST.Storage_Area_Desc
                    Row.Lot_No = row_ST.Batch_No
                    Row.Qty_Secondary = -1
                    Row.Warehouse_Code = row_ST.Warehouse_Code
                    Row.Warehouse_Desc = row_ST.Warehouse_Desc
                    If Row.Qty <> 0 Then
                        TBL_Prod.Rows.Add(Row)
                    End If
                End If
                TBL_ST.Rows.Remove(row_ST)
            End While

            lblPrgStep1_2.Text = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            tblSC.Rows.Clear()
            tblLTF.Rows.Clear()
            TBL_ST.Rows.Clear()
        End Try
    End Function

    Private Sub CekTidakCocok(ByVal TBL As DseLaporanStockDetail2.DataTable1DataTable, ByVal vTglAcuan As Date)
        Dim DTASC As New DataSet2TableAdapters.DataTable1TableAdapter
        Dim TBLSC As New DataSet2.DataTable1DataTable
        Dim rowSC As DataSet2.DataTable1Row
        Dim RetVal As Decimal = 0
        Dim row As DseLaporanStockDetail2.DataTable1Row

        Try
            Me.lblPrgStep1_2.Text = "Loading stock card data..."

            Me.prb_Step1_2.Value = 0
            'load data sc
            DTASC.Fill(TBLSC, vTglAcuan)
            Me.prb_Step1_2.Maximum = TBLSC.Rows.Count

            'banding dengan data tbl
            For Each rowSC In TBLSC.Rows
                Me.prb_Step1_2.Value += 1
                If Me.prb_Step1_2.Value = Me.prb_Step1_2.Maximum Or Math.IEEERemainder(Me.prb_Step1_2.Value, 10) = 0 Then
                    Me.lblPrgStep1_2.Text = "Checking stock tidak cocok - data " + Me.prb_Step1_2.Value.ToString + " of " + Me.prb_Step1_2.Maximum.ToString
                End If

                RetVal = 0
                TBL.DefaultView.RowFilter = "[Primary Product] = " + rowSC.Primary.ToString
                For Each DV As Data.DataRowView In TBL.DefaultView
                    row = DV.Row
                    RetVal += row.Qty
                Next

                If rowSC.IsQtyNull = True Then
                    rowSC.Qty = 0
                End If
                If rowSC.Qty <> RetVal Then
                    Me.TipeProgress = EnumTipeProgress.AddStockTidakCocok
                    Me.PrimaryProductTdkCocok = rowSC.Primary
                    Me.bgwStep1_2.ReportProgress(0)
                End If
            Next
            'langsung tampilkan klo ada yang tidak cocok

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AddDGVTidakCocok()
        Dim rowDGV As DataGridViewRow

        Try
            Me.dgvStockTidakCocok.Rows.Add()
            rowDGV = Me.dgvStockTidakCocok.Rows(Me.dgvStockTidakCocok.Rows.Count - 1)
            rowDGV.Cells(Me.ColTdkCocok_CodeNo.Index).Value = Me.PrimaryProductTdkCocok
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetStock(ByVal vPrimaryProduct As Double, ByVal TBL As DseLaporanStockDetail2.DataTable1DataTable) As Decimal
        Dim row As DseLaporanStockDetail2.DataTable1Row
        Dim RetVal As Decimal = 0

        Try
            TBL.DefaultView.RowFilter = "[Primary Product] = " + vPrimaryProduct.ToString
            For Each DV As Data.DataRowView In TBL.DefaultView
                row = DV.Row
                RetVal += row.Qty
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub LoadDGVTidakCocok()
        Dim rowDGV As DataGridViewRow

        Try
            Me.dgvStockTidakCocok.Rows.Add()
            rowDGV = Me.dgvStockTidakCocok.Rows(Me.dgvStockTidakCocok.Rows.Count - 1)
            rowDGV.Cells(Me.ColTdkCocok_CodeNo.Index).Value = Me.ObjProduct.GetProductCode(Me.PrimaryProductTdkCocok)
            rowDGV.Cells(Me.ColTdkCocok_Desc.Index).Value = Me.ObjProduct.GetDescription(Me.PrimaryProductTdkCocok)
            rowDGV.Cells(Me.ColTdkCocok_No.Index).Value = Me.dgvStockTidakCocok.Rows.Count
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    '1. load ke dgv item detail
    '2. load dgv stock minus
    Private Sub LoadDGVStock()
        Dim DV1 As Data.DataRowView
        Dim Row As DseLaporanStockDetail2.DataTable1Row
        Dim rowDGV, rowDGV1 As DataGridViewRow
        Dim ListMinus As New List(Of Double)
        Dim ListKodeMinus As New List(Of String) 

        Try
            TBL.DefaultView.RowFilter = ""
            TBL.DefaultView.Sort = "[Product Code], [Lot No]"
            Me.dgvStock.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvStock.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.dgvStock.Rows.Clear()

            Me.dgvStockMinus.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvStockMinus.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.dgvStockMinus.Rows.Clear()

            TBL.DefaultView.RowFilter = ""
            For Each DV1 In TBL.DefaultView
                Row = DV1.Row
                Me.dgvStock.Rows.Add()
                rowDGV = Me.dgvStock.Rows(Me.dgvStock.Rows.Count - 1)
                rowDGV.Cells(Me.ColStock_Batch.Index).Value = Row.Lot_No
                rowDGV.Cells(Me.ColStock_CodeNo.Index).Value = Row.Product_Code
                rowDGV.Cells(Me.ColStock_Desc.Index).Value = Row.Description
                rowDGV.Cells(Me.ColStock_Expiry.Index).Value = Row.Expiry_Date
                rowDGV.Cells(Me.ColStock_No.Index).Value = Me.dgvStock.Rows.Count
                rowDGV.Cells(Me.ColStock_Qty.Index).Value = Me.ObjCD.DigitGrouping(Row.Qty, Row._Decimal)
                rowDGV.Cells(Me.ColStock_Unit.Index).Value = Row.Unit_Code

                If Row.Qty < 0 Then
                    Me.dgvStockMinus.Rows.Add()
                    rowDGV1 = Me.dgvStockMinus.Rows(Me.dgvStockMinus.Rows.Count - 1)

                    rowDGV1.Cells(Me.ColDetailMinus_Batch.Index).Value = Row.Lot_No
                    rowDGV1.Cells(Me.ColDetailMinus_CodeNo.Index).Value = Row.Product_Code
                    rowDGV1.Cells(Me.ColDetailMinus_Desc.Index).Value = Row.Description
                    rowDGV1.Cells(Me.ColDetailMinus_Expiry.Index).Value = Row.Expiry_Date
                    rowDGV1.Cells(Me.ColDetailMinus_No.Index).Value = Me.dgvStockMinus.Rows.Count
                    rowDGV1.Cells(Me.ColDetailMinus_Qty.Index).Value = Me.ObjCD.DigitGrouping(Row.Qty, Row._Decimal)
                    rowDGV1.Cells(Me.ColDetailMinus_Unit.Index).Value = Row.Unit_Code
                End If
            Next
            Me.ColStock_Expiry.DefaultCellStyle.Format = "dd MMM yyyy"
            Me.dgvStock.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvStock.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            Me.ColDetailMinus_Expiry.DefaultCellStyle.Format = "dd MMM yyyy"
            Me.dgvStockMinus.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            Me.dgvStockMinus.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            Me.SudahLoadDGV = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function Simpan_Item_UsingScript(ByVal vTBL As DseLaporanStockDetail2.DataTable1DataTable) As Boolean
        Dim conStr As String = Core_Module.GetConstring
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()
            Me.prb_Step1_2.Value = 0
            Me.prb_Step1_2.Maximum = vTBL.Rows.Count

            While StartRowIndex < vTBL.Rows.Count
                Me.prb_Step1_2.Value = StartRowIndex
                Me.lblPrgStep1_2.Text = "Saving on database - data: " + StartRowIndex.ToString + " of " + vTBL.Rows.Count.ToString

                sql = Me.GetInsertString_Item(StartRowIndex, BatchQty, vTBL)
                script.Query = sql
                'execute script
                script.Execute()
                StartRowIndex += BatchQty
            End While

            Conn.Close()

            Me.lblPrgStep1_2.Text = ""
            Me.prb_Step1_2.Value = 0

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Private Function GetInsertString_Item(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, _
    ByVal vTBL As DseLaporanStockDetail2.DataTable1DataTable) As String
        Dim retval As String = ""
        Dim row As DseLaporanStockDetail2.DataTable1Row
        Dim Delimiter As String = ","
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            retval = "INSERT INTO `tbltrans_custom_stock` (`Primary Product`,`Qty`,`Qty Secondary`,`Lot No`,`Expiry Date`,`Primary Storage Area`) VALUES "
            For i As Double = 0 To BatchQty - 1 Step 1
                If vTBL.Rows.Count > vStartRowIndex + i Then
                    row = vTBL.Rows(i + vStartRowIndex)
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTBL.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    With row
                        retval = retval + vbCrLf + "(" + CDbl(.Primary_Product).ToString + "," _
                        + CDec(.Qty).ToString + "," + CDec(.Qty_Secondary).ToString + ",'" + Me.ObjCD.FormatSpecialChar_MySQL(.Lot_No) + "'," _
                        + "'" + Format(CDate(.Expiry_Date), "yyyy-MM-dd") + "'," + CDbl(.Primary_Storage_Area).ToString + ")" + Delimiter
                    End With
                Else
                    Exit For
                End If
            Next
            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            ObjCD = Nothing
        End Try
    End Function

    Private Sub bgwStep1_2_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgwStep1_2.DoWork
        Try
            Me.ObjCD.UpdateThreading()
            Me.CetakanStockDetail()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwStep1_2_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles bgwStep1_2.ProgressChanged
        Try
            If Me.TipeProgress = EnumTipeProgress.AddDGVStock Then
                Me.LoadDGVStock()
            ElseIf Me.TipeProgress = EnumTipeProgress.AddStockTidakCocok Then
                Me.LoadDGVTidakCocok()
            ElseIf Me.TipeProgress = EnumTipeProgress.ChangePrgStyle Then
                Me.ChangeProgressBar()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwStep1_2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgwStep1_2.RunWorkerCompleted
        Try
            Me.dgvStock.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvStock.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            Me.dgvStockTidakCocok.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            Me.dgvStockTidakCocok.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvStockMinus.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            Me.dgvStockMinus.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            MsgBox("Data updated", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnStep2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStep2.Click
        Try
            'buat stock out adjust untuk stock yang positive -> buat pada dtp
            If MsgBox("Start creating stock out adjust on " + Format(Me.dtpResetStock.Value, "dd MMM yyyy") + "?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Control.CheckForIllegalCrossThreadCalls = False
                Me.bgwStep2.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnStep3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStep3.Click
        Try
            'buat stock in adjust untuk stock yang minus -> buat pada dtp 1
            If MsgBox("Start creating stock in adjust on " + Format(Me.dtpResetStock.Value, "dd MMM yyyy") + "?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Control.CheckForIllegalCrossThreadCalls = False
                Me.bgwStep3.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnStep4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStep4.Click
        Try
            'buat stock in adjust, sejumlah stock negative -> buat pada dtp
            If MsgBox("Start creating stock in adjust on " + Format(Me.dtpResetStock.Value, "dd MMM yyyy") + "?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Control.CheckForIllegalCrossThreadCalls = False
                Me.bgwStep4.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnStep5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            'buat stock in adjust pada stock yang positif -> buat pada dtp +1

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnStep6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            'hitung amount untuk setiap item-nya


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwStep2_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgwStep2.DoWork
        Dim DocNo As String = "RESET2017-1"
        Dim TipeStockInOut As clsTransaksiStockInOut.EnumTipeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust
        Dim TipeSC As ClsTransaksiStockCard.EnumJenisTransaksiAsal = ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockOut_Adjust
        Dim InOut As ClsTransaksiStockCard.EnumJenisStockCard = ClsTransaksiStockCard.EnumJenisStockCard.Stock_Out
        Dim PrimaryLama As Double = -1
        Dim Tgl As Date = Me.dtpResetStock.Value.Date
        Dim DTAHeader As New DataSet2TableAdapters.tbltransaksi_stockinout1TableAdapter
        Dim DTASC As New DataSet2TableAdapters.tbltransaksi_stock_card1TableAdapter
        Dim DTACustom As New DataSet2TableAdapters.tbltrans_custom_stockTableAdapter
        Dim TBLCustom As New DataSet2.tbltrans_custom_stockDataTable

        Try
            Me.ObjCD.UpdateThreading()

            Me.prgStep2.Value = 0
            Me.PrgBar = Me.prgStep2
            Me.LblPrgStep2.Text = "Initializing..."
            Me.TipePrg = ProgressBarStyle.Marquee
            Me.bgwStep2.ReportProgress(0)

            'buat stock out adjust untuk stock yang positive -> buat pada dtp
            'cari primary yang lama, hapus
            Me.LblPrgStep2.Text = "Deleting old data..."
            If DTAHeader.ScalarQuery_GetPrimary(DocNo).HasValue = True Then
                PrimaryLama = DTAHeader.ScalarQuery_GetPrimary(DocNo).Value
                'hapus header + item
                DTAHeader.Delete(PrimaryLama)
                'hapus stock card
                DTASC.Delete(PrimaryLama, TipeSC)
            End If

            'load data dari stock custom, hanya yang positive saja
            Me.LblPrgStep2.Text = "Loading custom stock data..."
            DTACustom.FillBy_PositiveOnly(TBLCustom)

            'masukkan data header
            Me.LblPrgStep2.Text = "Inserting header data..."
            DTAHeader.Insert(DocNo, 0, CInt(TipeStockInOut), Tgl, -1, -1, -1, -1, 0, -1, "", "USER", Now, "USER", Now, Now, "", -1, 0, -1, 0, 0, -1)
            PrimaryLama = DTAHeader.ScalarQuery_GetPrimary(DocNo).Value

            'masukkan data item -> qty harus abs
            Me.LblPrgStep2.Text = "Inserting item data..."
            Me.TipePrg = ProgressBarStyle.Continuous
            Me.bgwStep2.ReportProgress(0)
            Me.prgStep2.Value = 0
            Me.prgStep2.Maximum = TBLCustom.Rows.Count
            Me.Simpan_ItemInOut_UsingScript(TBLCustom, PrimaryLama, Me.prgStep2, Me.LblPrgStep2)
            'masukkan data stock card
            Me.prgStep2.Value = 0
            Me.prgStep2.Maximum = TBLCustom.Rows.Count
            Me.Simpan_ItemSC_UsingScript(TBLCustom, PrimaryLama, TipeSC, DocNo, InOut, Tgl, Me.prgStep2, Me.LblPrgStep2)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function Simpan_ItemSC_UsingScript(ByVal vTBL As DataSet2.tbltrans_custom_stockDataTable, ByVal vPrimaryTrans As Double, _
    ByVal vTipeSC As ClsTransaksiStockCard.EnumJenisTransaksiAsal, ByVal vDocNo As String, _
    ByVal vInOut As ClsTransaksiStockCard.EnumJenisStockCard, ByVal vTglTrans As Date, ByVal vPrg As ProgressBar, ByVal vLbl As Label) As Boolean
        Dim conStr As String = Core_Module.GetConstring
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()
            vPrg.Value = 0
            vPrg.Maximum = vTBL.Rows.Count

            While StartRowIndex < vTBL.Rows.Count
                vPrg.Value = StartRowIndex
                vLbl.Text = "Saving on stock card item table - data: " + StartRowIndex.ToString + " of " + vTBL.Rows.Count.ToString

                sql = Me.GetInsertString_ItemSC(StartRowIndex, BatchQty, vTBL, vPrimaryTrans, vTipeSC, vDocNo, vInOut, vTglTrans)
                script.Query = sql
                'execute script
                script.Execute()
                StartRowIndex += BatchQty
            End While

            Conn.Close()

            vLbl.Text = ""
            vPrg.Value = 0

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Private Function GetInsertString_ItemSC(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, _
    ByVal vTBL As DataSet2.tbltrans_custom_stockDataTable, ByVal vPrimaryTrans As Double, ByVal vTipeSC As ClsTransaksiStockCard.EnumJenisTransaksiAsal, _
    ByVal vDocNo As String, ByVal vInOut As ClsTransaksiStockCard.EnumJenisStockCard, ByVal vTglTrans As Date) As String
        Dim retval As String = ""
        Dim row As DataSet2.tbltrans_custom_stockRow
        Dim Delimiter As String = ","
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            retval = "INSERT INTO `tbltransaksi_stock_card` (`Primary Transaksi Asal`,`Type Transaksi Asal`,`Primary Product`,`Document No`,`In`,`Out`,`Description`,`InOut`" + _
            ",`Date`,`Expiry`,`primary location`,`Batch Number`,`Remarks`) VALUES "
            For i As Double = 0 To BatchQty - 1 Step 1
                If vTBL.Rows.Count > vStartRowIndex + i Then
                    row = vTBL.Rows(i + vStartRowIndex)
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTBL.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    With row
                        If vTipeSC = ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockIn_Adjust Then
                            retval = retval + vbCrLf + "(" + CDbl(vPrimaryTrans).ToString + "," + CInt(vTipeSC).ToString + "," + CDbl(.Primary_Product).ToString + ",'" + vDocNo + "'," _
                            + CDec(Math.Abs(.Qty)).ToString + ",0,'Stock In Adjust'," + CInt(vInOut).ToString + ",'" + Format(vTglTrans, "yyyy-MM-dd") + "'," + "'" + _
                            Format(CDate(.Expiry_Date), "yyyy-MM-dd") + "'," + CDbl(.Primary_Storage_Area).ToString + "," + "'" + Me.ObjCD.FormatSpecialChar_MySQL(.Lot_No) + "'," _
                            + "''" + ")" + Delimiter
                        ElseIf vTipeSC = ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockOut_Adjust Then
                            retval = retval + vbCrLf + "(" + CDbl(vPrimaryTrans).ToString + "," + CInt(vTipeSC).ToString + "," + CDbl(.Primary_Product).ToString + ",'" + vDocNo + "'," _
                            + "0," + CDec(Math.Abs(.Qty)).ToString + ",'Stock Out Adjust'," + CInt(vInOut).ToString + ",'" + Format(vTglTrans, "yyyy-MM-dd") + "'," + "'" + _
                            Format(CDate(.Expiry_Date), "yyyy-MM-dd") + "'," + CDbl(.Primary_Storage_Area).ToString + "," + "'" + Me.ObjCD.FormatSpecialChar_MySQL(.Lot_No) + "'," _
                            + "''" + ")" + Delimiter
                        Else
                            MsgBox("Logic error")
                        End If
                    End With
                Else
                    Exit For
                End If
            Next
            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            ObjCD = Nothing
        End Try
    End Function

    Private Function Simpan_ItemInOut_UsingScript(ByVal vTBL As DataSet2.tbltrans_custom_stockDataTable, ByVal vPrimaryTrans As Double, ByVal vPrg As ProgressBar, ByVal vLbl As Label) As Boolean
        Dim conStr As String = Core_Module.GetConstring
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()
            vPrg.Value = 0
            vPrg.Maximum = vTBL.Rows.Count

            While StartRowIndex < vTBL.Rows.Count
                vPrg.Value = StartRowIndex
                vLbl.Text = "Saving on stock in out item table - data: " + StartRowIndex.ToString + " of " + vTBL.Rows.Count.ToString

                sql = Me.GetInsertString_ItemInOut(StartRowIndex, BatchQty, vTBL, vPrimaryTrans)
                script.Query = sql
                'execute script
                script.Execute()
                StartRowIndex += BatchQty
            End While

            Conn.Close()

            vLbl.Text = ""
            vPrg.Value = 0

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Private Function GetInsertString_ItemInOut(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, _
    ByVal vTBL As DataSet2.tbltrans_custom_stockDataTable, ByVal vPrimaryTrans As Double) As String
        Dim retval As String = ""
        Dim row As DataSet2.tbltrans_custom_stockRow
        Dim Delimiter As String = ","
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            retval = "INSERT INTO `tbltransaksi_stockinout_item` (`Primary Transaksi`,`Primary Product`,`Qty`,`Primary Units`,`Secondary Qty`,`Primary Secondary Units`," + _
            "`Batch No`,`Expiry Date`,`Primary Storage Area`,`Primary PO Purchase`,`Primary PdO`,`Supp. DO No`,`Remarks`,`Amount`,`Primary Multi Units`," + _
            "`Qty Multi Unit`,`Rasio`,`Type Sizing`,`Ratio Default`) VALUES "
            For i As Double = 0 To BatchQty - 1 Step 1
                If vTBL.Rows.Count > vStartRowIndex + i Then
                    row = vTBL.Rows(i + vStartRowIndex)
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTBL.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    With row
                        retval = retval + vbCrLf + "(" + CDbl(vPrimaryTrans).ToString + "," + CDbl(.Primary_Product).ToString + "," _
                        + CDec(Math.Abs(.Qty)).ToString + "," + CDbl(.Primary_Unit).ToString + ",-1,-1," + "'" + Me.ObjCD.FormatSpecialChar_MySQL(.Lot_No) + "'," _
                        + "'" + Format(CDate(.Expiry_Date), "yyyy-MM-dd") + "'," + CDbl(.Primary_Storage_Area).ToString + ",-1,-1,'','',0," + _
                        CDbl(.Primary_Unit).ToString + "," + CDec(.Qty).ToString + ",1,0,1" + ")" + Delimiter
                    End With
                Else
                    Exit For
                End If
            Next
            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            ObjCD = Nothing
        End Try
    End Function

    Private Sub bgwStep2_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles bgwStep2.ProgressChanged
        Try
            If Me.TipeProgress = EnumTipeProgress.ChangePrgStyle2 Then
                Me.ChangeProgressBar2()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwStep2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgwStep2.RunWorkerCompleted
        Try
            Me.prgStep2.Value = 0
            Me.LblPrgStep2.Text = "Finished"
            MsgBox("Finished", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwStep3_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgwStep3.DoWork
        Dim DocNo As String = "RESET2017-2"
        Dim TipeStockInOut As clsTransaksiStockInOut.EnumTipeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust
        Dim TipeSC As ClsTransaksiStockCard.EnumJenisTransaksiAsal = ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockIn_Adjust
        Dim InOut As ClsTransaksiStockCard.EnumJenisStockCard = ClsTransaksiStockCard.EnumJenisStockCard.Stock_In
        Dim PrimaryLama As Double = -1
        Dim Tgl As Date = Me.dtpResetStock.Value.Date
        Dim DTAHeader As New DataSet2TableAdapters.tbltransaksi_stockinout1TableAdapter
        Dim DTASC As New DataSet2TableAdapters.tbltransaksi_stock_card1TableAdapter
        Dim DTACustom As New DataSet2TableAdapters.tbltrans_custom_stockTableAdapter
        Dim TBLCustom As New DataSet2.tbltrans_custom_stockDataTable

        Try
            Me.ObjCD.UpdateThreading()

            Me.PrgStep3.Value = 0
            Me.LblPrgStep3.Text = "Initializing..."
            Me.PrgBar = Me.PrgStep3
            Me.TipePrg = ProgressBarStyle.Marquee
            Me.bgwStep3.ReportProgress(0)

            'buat stock out adjust untuk stock yang positive -> buat pada dtp
            'cari primary yang lama, hapus
            Me.LblPrgStep3.Text = "Deleting old data..."
            If DTAHeader.ScalarQuery_GetPrimary(DocNo).HasValue = True Then
                PrimaryLama = DTAHeader.ScalarQuery_GetPrimary(DocNo).Value
                'hapus header + item
                DTAHeader.Delete(PrimaryLama)
                'hapus stock card
                DTASC.Delete(PrimaryLama, TipeSC)
            End If

            'load data dari stock custom, hanya yang positive saja
            Me.LblPrgStep3.Text = "Loading custom stock data..."
            DTACustom.FillBy_NegativeOnly(TBLCustom)

            'masukkan data header
            Me.LblPrgStep3.Text = "Inserting header data..."
            DTAHeader.Insert(DocNo, 0, CInt(TipeStockInOut), Tgl, -1, -1, -1, -1, 0, -1, "", "USER", Now, "USER", Now, Now, "", -1, 0, -1, 0, 0, -1)
            PrimaryLama = DTAHeader.ScalarQuery_GetPrimary(DocNo).Value

            'masukkan data item -> qty harus abs
            Me.LblPrgStep3.Text = "Inserting item data..."
            Me.TipePrg = ProgressBarStyle.Continuous
            Me.bgwStep3.ReportProgress(0)
            Me.PrgStep3.Value = 0
            Me.PrgStep3.Maximum = TBLCustom.Rows.Count
            Me.Simpan_ItemInOut_UsingScript(TBLCustom, PrimaryLama, Me.PrgStep3, Me.LblPrgStep3)
            'masukkan data stock card
            Me.PrgStep3.Value = 0
            Me.PrgStep3.Maximum = TBLCustom.Rows.Count
            Me.Simpan_ItemSC_UsingScript(TBLCustom, PrimaryLama, TipeSC, DocNo, InOut, Tgl, Me.PrgStep3, Me.LblPrgStep3)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwStep3_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles bgwStep3.ProgressChanged
        Try
            If Me.TipeProgress = EnumTipeProgress.ChangePrgStyle2 Then
                Me.ChangeProgressBar2()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwStep3_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgwStep3.RunWorkerCompleted
        Try
            Me.PrgStep3.Value = 0
            Me.LblPrgStep3.Text = "Finished"
            MsgBox("Finished", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnStep7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStep7.Click
        Dim Primary1, Primary2, Primary3, Primary4, Primary5 As Double
        Dim DTASC As New DataSet2TableAdapters.tbltransaksi_stockinout1TableAdapter
        Dim DTASC1 As New DataSet2TableAdapters.DataTable2TableAdapter
        Dim TBLSC1 As New DataSet2.DataTable2DataTable
        'Dim rowSC As DataSet2.DataTable2Row
        Dim rowDGV As DataGridViewRow

        Dim DTASC2 As New DataSet2TableAdapters.DataTable3TableAdapter
        Dim TBLSC2 As New DataSet2.DataTable3DataTable
        Dim rowSC2 As DataSet2.DataTable3Row
        Dim PrimaryTest As Double = -1

        Try
            PrimaryTest = Me.objStockInOut.GetPrimary("RESET1")
            If PrimaryTest = -1 Then
                MsgBox("Can't find 'RESET1', please re-check coding", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            PrimaryTest = Me.objStockInOut.GetPrimary("RESET2")
            If PrimaryTest = -1 Then
                MsgBox("Can't find 'RESET2', please re-check coding", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            PrimaryTest = Me.objStockInOut.GetPrimary("RESET3")
            If PrimaryTest = -1 Then
                MsgBox("Can't find 'RESET3', please re-check coding", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            PrimaryTest = Me.objStockInOut.GetPrimary("RESET4")
            If PrimaryTest = -1 Then
                MsgBox("Can't find 'RESET4', please re-check coding", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            PrimaryTest = Me.objStockInOut.GetPrimary("RESET5")
            If PrimaryTest = -1 Then
                MsgBox("Can't find 'RESET5', please re-check coding", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            Primary1 = DTASC.ScalarQuery_GetPrimary("RESET2017-1").Value
            Primary2 = DTASC.ScalarQuery_GetPrimary("RESET2017-2").Value
            Primary3 = DTASC.ScalarQuery_GetPrimary("RESET2017-3").Value
            Primary4 = DTASC.ScalarQuery_GetPrimary("RESET2017-4").Value
            Primary5 = DTASC.ScalarQuery_GetPrimary("RESET2017-5").Value

            'pastikan total sc = 0
            DTASC2.Fill(TBLSC2, Primary1, Primary2, Primary3, Primary4, Primary5)
            If TBLSC2.Rows.Count > 0 Then
                For Each rowSC2 In TBLSC2.Rows
                    Me.dgvCekSC.Rows.Add()
                    rowDGV = Me.dgvCekSC.Rows(Me.dgvCekSC.Rows.Count - 1)
                    rowDGV.Cells(Me.ColCekSC_BatchNo.Index).Value = ""
                    rowDGV.Cells(Me.ColCekSC_CodeNo.Index).Value = rowSC2.Product_Code
                    rowDGV.Cells(Me.ColCekSC_Desc.Index).Value = rowSC2.Description
                    rowDGV.Cells(Me.ColCekSC_Location.Index).Value = ""
                    rowDGV.Cells(Me.ColCekSC_No.Index).Value = Me.dgvCekSC.Rows.Count
                    rowDGV.Cells(Me.ColCekSC_Qty.Index).Value = Me.ObjCD.BuangNolDibelakangDecimal(rowSC2.Qty)
                    rowDGV.Cells(Me.ColCekSC_Unit.Index).Value = rowSC2.Unit_Code
                Next
                MsgBox("Overall stock card is not valid", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            'pastikan total stock in adjust = total negative
            Dim DTACustom As New DataSet2TableAdapters.tbltrans_custom_stockTableAdapter
            Dim TBLCustom As New DataSet2.tbltrans_custom_stockDataTable
            Dim RowCustom As DataSet2.tbltrans_custom_stockRow
            Dim TBLSC_N_InAdjust As New DataSet2.DataTable2DataTable
            Dim TBLSC_N_OutAdjust As New DataSet2.DataTable2DataTable
            Dim TBLSC_P_InAdjust1 As New DataSet2.DataTable2DataTable
            Dim TBLSC_P_InAdjust2 As New DataSet2.DataTable2DataTable
            Dim TBLSC_P_OutAdjust As New DataSet2.DataTable2DataTable
            Dim row1, row2, row3 As DataSet2.DataTable2Row

            '-. load data negative
            DTACustom.FillBy_NegativeOnly(TBLCustom)
            '-. load semua stock in/out adjust dari negative
            DTASC1.FillBy1(TBLSC_N_InAdjust, Primary2)
            DTASC1.FillBy1(TBLSC_N_OutAdjust, Primary4)
            '-. total stock in/out adjust harus sama dengan 0
            For Each row1 In TBLSC_N_InAdjust.Rows
                If row1.Qty < 0 Then
                    row1.Qty = Math.Abs(row1.Qty)
                End If
            Next
            For Each row2 In TBLSC_N_OutAdjust.Rows
                If row2.Qty < 0 Then
                    row2.Qty = Math.Abs(row2.Qty)
                End If
            Next
            For ij As Double = TBLSC_N_InAdjust.Rows.Count - 1 To 0 Step -1
                row1 = TBLSC_N_InAdjust.Rows(ij)
                For jk As Double = TBLSC_N_OutAdjust.Rows.Count - 1 To 0 Step -1
                    row2 = TBLSC_N_OutAdjust.Rows(jk)
                    If row1.primary_location = row2.primary_location And row1.Primary_Product = row2.Primary_Product And _
                    row1.Batch_Number = row2.Batch_Number Then
                        If row1.Qty > row2.Qty Then
                            row1.Qty -= row2.Qty
                            TBLSC_N_OutAdjust.Rows.Remove(row2)
                        Else
                            row2.Qty -= row1.Qty
                            row1.Qty = 0
                            Exit For
                        End If
                        If row1.Qty = 0 Then
                            Exit For
                        End If
                    End If
                Next
                If row1.Qty = 0 Then
                    TBLSC_N_InAdjust.Rows.Remove(row1)
                End If
            Next
            If TBLSC_N_InAdjust.Rows.Count <> 0 Then
                MsgBox("Negative stock is not balance", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            '-. total stock in adjust harus sama dengan negativenya
            TBLSC_N_InAdjust.Rows.Clear()
            DTASC1.FillBy1(TBLSC_N_InAdjust, Primary2)
            For Each row1 In TBLSC_N_InAdjust.Rows
                If row1.Qty < 0 Then
                    row1.Qty = Math.Abs(row1.Qty)
                End If
            Next
            For Each RowCustom In TBLCustom.Rows
                If RowCustom.Qty < 0 Then
                    RowCustom.Qty = Math.Abs(RowCustom.Qty)
                End If
            Next
            For ij As Double = TBLSC_N_InAdjust.Rows.Count - 1 To 0 Step -1
                row1 = TBLSC_N_InAdjust.Rows(ij)
                For jk As Double = TBLCustom.Rows.Count - 1 To 0 Step -1
                    RowCustom = TBLCustom.Rows(jk)
                    If row1.primary_location = RowCustom.Primary_Storage_Area And row1.Primary_Product = RowCustom.Primary_Product And _
                    row1.Batch_Number = RowCustom.Lot_No Then
                        If row1.Qty > RowCustom.Qty Then
                            row1.Qty -= RowCustom.Qty
                            TBLCustom.Rows.Remove(RowCustom)
                        Else
                            RowCustom.Qty -= row1.Qty
                            row1.Qty = 0
                            Exit For
                        End If
                        If row1.Qty = 0 Then
                            Exit For
                        End If
                    End If
                Next
                If row1.Qty = 0 Then
                    TBLSC_N_InAdjust.Rows.Remove(row1)
                End If
            Next
            If TBLSC_N_InAdjust.Rows.Count <> 0 Then
                MsgBox("Negative stock is not balance2", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            'pastikan todal stock out adjust = total positive
            '-. load data positive
            TBLCustom.Rows.Clear()
            DTACustom.FillBy_PositiveOnly(TBLCustom)
            '-. load semua stock in/out adjust dari positive
            DTASC1.FillBy1(TBLSC_P_OutAdjust, Primary1)
            DTASC1.FillBy1(TBLSC_P_InAdjust1, Primary3)
            DTASC1.FillBy1(TBLSC_P_InAdjust2, Primary5)
            '-. total stock in/out adjust harus sama dengan 0
            For Each row1 In TBLSC_P_InAdjust1.Rows
                If row1.Qty < 0 Then
                    row1.Qty = Math.Abs(row1.Qty)
                End If
            Next
            For Each row1 In TBLSC_P_InAdjust2.Rows
                If row1.Qty < 0 Then
                    row1.Qty = Math.Abs(row1.Qty)
                End If
            Next
            For Each row1 In TBLSC_P_OutAdjust.Rows
                If row1.Qty < 0 Then
                    row1.Qty = Math.Abs(row1.Qty)
                End If
            Next
            For ij As Double = TBLSC_P_OutAdjust.Rows.Count - 1 To 0 Step -1
                row1 = TBLSC_P_OutAdjust.Rows(ij)
                For jk As Double = TBLSC_P_InAdjust1.Rows.Count - 1 To 0 Step -1
                    row2 = TBLSC_P_InAdjust1.Rows(jk)
                    If row1.primary_location = row2.primary_location And row1.Primary_Product = row2.Primary_Product And _
                    row1.Batch_Number = row2.Batch_Number Then
                        If row1.Qty > row2.Qty Then
                            row1.Qty -= row2.Qty
                            TBLSC_P_InAdjust1.Rows.Remove(row2)
                        Else
                            row2.Qty -= row1.Qty
                            row1.Qty = 0
                            Exit For
                        End If
                        If row1.Qty = 0 Then
                            Exit For
                        End If
                    End If
                Next
                If row1.Qty = 0 Then
                    TBLSC_P_OutAdjust.Rows.Remove(row1)
                End If
            Next
            For ij As Double = TBLSC_P_OutAdjust.Rows.Count - 1 To 0 Step -1
                row1 = TBLSC_P_OutAdjust.Rows(ij)
                For jk As Double = TBLSC_P_InAdjust2.Rows.Count - 1 To 0 Step -1
                    row2 = TBLSC_P_InAdjust2.Rows(jk)
                    If row1.primary_location = row2.primary_location And row1.Primary_Product = row2.Primary_Product And _
                    row1.Batch_Number = row2.Batch_Number Then
                        If row1.Qty > row2.Qty Then
                            row1.Qty -= row2.Qty
                            TBLSC_P_InAdjust2.Rows.Remove(row2)
                        Else
                            row2.Qty -= row1.Qty
                            row1.Qty = 0
                            Exit For
                        End If
                        If row1.Qty = 0 Then
                            Exit For
                        End If
                    End If
                Next
                If row1.Qty = 0 Then
                    TBLSC_P_OutAdjust.Rows.Remove(row1)
                End If
            Next
            If TBLSC_P_OutAdjust.Rows.Count <> 0 Then
                MsgBox("Positive stock is not balance", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            '-. total stock out adjust harus sama dengan positivenya
            TBLSC_P_OutAdjust.Rows.Clear()
            DTASC1.FillBy1(TBLSC_P_OutAdjust, Primary1)
            For ij As Double = TBLSC_P_OutAdjust.Rows.Count - 1 To 0 Step -1
                row1 = TBLSC_P_OutAdjust.Rows(ij)
                For jk As Double = TBLCustom.Rows.Count - 1 To 0 Step -1
                    RowCustom = TBLCustom.Rows(jk)
                    If row1.primary_location = RowCustom.Primary_Storage_Area And row1.Primary_Product = RowCustom.Primary_Product And _
                    row1.Batch_Number = RowCustom.Lot_No Then
                        If row1.Qty > RowCustom.Qty Then
                            row1.Qty -= RowCustom.Qty
                            TBLCustom.Rows.Remove(RowCustom)
                        Else
                            RowCustom.Qty -= row1.Qty
                            row1.Qty = 0
                            Exit For
                        End If
                        If row1.Qty = 0 Then
                            Exit For
                        End If
                    End If
                Next
                If row1.Qty = 0 Then
                    TBLSC_P_OutAdjust.Rows.Remove(row1)
                End If
            Next
            If TBLSC_P_OutAdjust.Rows.Count <> 0 Then
                MsgBox("Negative stock is not balance2", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            'kalo sampe sini ok
            MsgBox("Data checked: Status OK", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwStep4_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgwStep4.DoWork
        Dim DocNo As String = "RESET2017-3"
        Dim TipeStockInOut As clsTransaksiStockInOut.EnumTipeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust
        Dim TipeSC As ClsTransaksiStockCard.EnumJenisTransaksiAsal = ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockIn_Adjust
        Dim InOut As ClsTransaksiStockCard.EnumJenisStockCard = ClsTransaksiStockCard.EnumJenisStockCard.Stock_In
        Dim PrimaryLama As Double = -1
        Dim Tgl As Date = Me.dtpResetStock.Value.Date
        Dim DTAHeader As New DataSet2TableAdapters.tbltransaksi_stockinout1TableAdapter
        Dim DTASC As New DataSet2TableAdapters.tbltransaksi_stock_card1TableAdapter
        Dim DTACustom As New DataSet2TableAdapters.tbltrans_custom_stockTableAdapter
        Dim rowNegative, rowPositive1, rowPositive2 As DataSet2.tbltrans_custom_stockRow
        Dim JmlPindah As Decimal = 0

        Try
            Me.ObjCD.UpdateThreading()

            Me.prgStep4.Value = 0
            Me.LblPrgStep4.Text = "Initializing..."
            Me.PrgBar = Me.prgStep4
            Me.TipePrg = ProgressBarStyle.Marquee
            Me.bgwStep4.ReportProgress(0)

            'load data dari stock custom, hanya yang positive saja
            Me.LblPrgStep4.Text = "Loading custom stock data..."
            Me.TBLNegative.Rows.Clear()
            Me.TBLPositive1.Rows.Clear()
            Me.TBLPositive2.Rows.Clear()
            DTACustom.FillBy_NegativeOnly(Me.TBLNegative)
            DTACustom.FillBy_PositiveOnly(Me.TBLPositive1)

            'pindahkan data dari tabel positive 1 -> 2, sejumlah baris negative
            Me.prgStep4.Value = 0
            Me.LblPrgStep4.Text = "Initializing..."
            Me.PrgBar = Me.prgStep4
            Me.TipePrg = ProgressBarStyle.Marquee
            Me.bgwStep4.ReportProgress(0)
            Me.prgStep4.Value = 0
            Me.prgStep4.Maximum = TBLNegative.Rows.Count
            For Each rowNegative In TBLNegative.Rows
                Me.prgStep4.Value += 1
                Me.LblPrgStep4.Text = "Moving stock out adjust - data " + Me.prgStep4.Value.ToString + " of " + Me.prgStep4.Maximum.ToString

                JmlPindah = Math.Abs(rowNegative.Qty)
                Me.TBLPositive1.DefaultView.RowFilter = "[Primary Product] = " + rowNegative.Primary_Product.ToString
                If Me.TBLPositive1.DefaultView.Count = 0 Then
                    MsgBox("Logic error: please make sure there is negative stock on " + _
                    Me.ObjProduct.GetProductCode(rowNegative.Primary_Product), MsgBoxStyle.Information, "Information")
                    Exit Sub
                Else
                    For ij As Integer = TBLPositive1.DefaultView.Count - 1 To 0 Step -1
                        If JmlPindah > 0 Then
                            rowPositive1 = TBLPositive1.DefaultView(ij).Row
                            If JmlPindah > rowPositive1.Qty Then
                                'row positive 2 = row positive 1
                                rowPositive2 = TBLPositive2.Newtbltrans_custom_stockRow
                                For jk As Integer = 0 To TBLPositive2.Columns.Count - 1
                                    rowPositive2.Item(jk) = rowPositive1.Item(jk)
                                Next
                                TBLPositive2.Rows.Add(rowPositive2)
                                JmlPindah -= rowPositive1.Qty
                                TBLPositive1.Rows.Remove(rowPositive1)
                            Else
                                'qty rowpositive 2 = qty jml pindah
                                rowPositive2 = TBLPositive2.Newtbltrans_custom_stockRow
                                For jk As Integer = 0 To TBLPositive2.Columns.Count - 1
                                    rowPositive2.Item(jk) = rowPositive1.Item(jk)
                                Next
                                rowPositive2.Qty = JmlPindah
                                TBLPositive2.Rows.Add(rowPositive2)
                                rowPositive1.Qty -= JmlPindah
                                JmlPindah = 0
                                If rowPositive1.Qty = 0 Then
                                    TBLPositive1.Rows.Remove(rowPositive1)
                                End If
                            End If
                        Else
                            Exit For
                        End If
                    Next
                    If JmlPindah > 0 Then
                        MsgBox("Logic error: please make sure there is negative stock on " + _
                        Me.ObjProduct.GetProductCode(rowNegative.Primary_Product), MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                End If
            Next

            'buat stock in adjust dari yang positive, sejumlah yang negative, pada dtp -> tbl positive 2
            'cari primary yang lama, hapus
            Me.PrgBar = Me.prgStep4
            Me.TipePrg = ProgressBarStyle.Marquee
            Me.bgwStep4.ReportProgress(0)
            Me.LblPrgStep4.Text = "Deleting old data..."
            If DTAHeader.ScalarQuery_GetPrimary(DocNo).HasValue = True Then
                PrimaryLama = DTAHeader.ScalarQuery_GetPrimary(DocNo).Value
                'hapus header + item
                DTAHeader.Delete(PrimaryLama)
                'hapus stock card
                DTASC.Delete(PrimaryLama, TipeSC)
            End If

            'masukkan data header
            Me.LblPrgStep4.Text = "Inserting header data..."
            DTAHeader.Insert(DocNo, 0, CInt(TipeStockInOut), Tgl, -1, -1, -1, -1, 0, -1, "", "USER", Now, "USER", Now, Now, "", -1, 0, -1, 0, 0, -1)
            PrimaryLama = DTAHeader.ScalarQuery_GetPrimary(DocNo).Value

            'masukkan data item -> qty harus abs
            Me.LblPrgStep4.Text = "Inserting item data..."
            Me.TipePrg = ProgressBarStyle.Continuous
            Me.bgwStep4.ReportProgress(0)
            Me.prgStep4.Value = 0
            Me.prgStep4.Maximum = Me.TBLPositive2.Rows.Count
            Me.Simpan_ItemInOut_UsingScript(Me.TBLPositive2, PrimaryLama, Me.prgStep4, Me.LblPrgStep4)
            'masukkan data stock card
            Me.prgStep4.Value = 0
            Me.prgStep4.Maximum = Me.TBLPositive2.Rows.Count
            Me.Simpan_ItemSC_UsingScript(Me.TBLPositive2, PrimaryLama, TipeSC, DocNo, InOut, Tgl, Me.prgStep4, Me.LblPrgStep4)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwStep4_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles bgwStep4.ProgressChanged
        Try
            If Me.TipeProgress = EnumTipeProgress.ChangePrgStyle2 Then
                Me.ChangeProgressBar2()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwStep4_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgwStep4.RunWorkerCompleted
        Try
            Me.prgStep4.Value = 0
            Me.LblPrgStep4.Text = "Finished"
            Me.bgwStep5.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwStep5_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgwStep5.DoWork
        Dim DocNo As String = "RESET2017-4"
        Dim TipeStockInOut As clsTransaksiStockInOut.EnumTipeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust
        Dim TipeSC As ClsTransaksiStockCard.EnumJenisTransaksiAsal = ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockOut_Adjust
        Dim InOut As ClsTransaksiStockCard.EnumJenisStockCard = ClsTransaksiStockCard.EnumJenisStockCard.Stock_Out
        Dim PrimaryLama As Double = -1
        Dim Tgl As Date = Me.dtpResetStock.Value.Date
        Dim DTAHeader As New DataSet2TableAdapters.tbltransaksi_stockinout1TableAdapter
        Dim DTASC As New DataSet2TableAdapters.tbltransaksi_stock_card1TableAdapter
        Dim DTACustom As New DataSet2TableAdapters.tbltrans_custom_stockTableAdapter
   
        Try
            Me.ObjCD.UpdateThreading()

            Me.prgStep5.Value = 0
            Me.LblPrgStep5.Text = "Initializing..."
            Me.PrgBar = Me.prgStep5
            Me.TipePrg = ProgressBarStyle.Marquee
            Me.bgwStep5.ReportProgress(0)

            'buat stock out adjust dari yg negative
            'cari primary yang lama, hapus
            Me.PrgBar = Me.prgStep5
            Me.TipePrg = ProgressBarStyle.Marquee
            Me.bgwStep5.ReportProgress(0)
            Me.LblPrgStep5.Text = "Deleting old data..."
            If DTAHeader.ScalarQuery_GetPrimary(DocNo).HasValue = True Then
                PrimaryLama = DTAHeader.ScalarQuery_GetPrimary(DocNo).Value
                'hapus header + item
                DTAHeader.Delete(PrimaryLama)
                'hapus stock card
                DTASC.Delete(PrimaryLama, TipeSC)
            End If

            'masukkan data header
            Me.LblPrgStep5.Text = "Inserting header data..."
            DTAHeader.Insert(DocNo, 0, CInt(TipeStockInOut), Tgl, -1, -1, -1, -1, 0, -1, "", "USER", Now, "USER", Now, Now, "", -1, 0, -1, 0, 0, -1)
            PrimaryLama = DTAHeader.ScalarQuery_GetPrimary(DocNo).Value

            'masukkan data item -> qty harus abs
            Me.LblPrgStep5.Text = "Inserting item data..."
            Me.TipePrg = ProgressBarStyle.Continuous
            Me.bgwStep5.ReportProgress(0)
            Me.prgStep5.Value = 0
            Me.prgStep5.Maximum = Me.TBLNegative.Rows.Count
            Me.Simpan_ItemInOut_UsingScript(Me.TBLNegative, PrimaryLama, Me.prgStep5, Me.LblPrgStep5)
            'masukkan data stock card
            Me.prgStep5.Value = 0
            Me.prgStep5.Maximum = Me.TBLNegative.Rows.Count
            Me.Simpan_ItemSC_UsingScript(Me.TBLNegative, PrimaryLama, TipeSC, DocNo, InOut, Tgl, Me.prgStep5, Me.LblPrgStep5)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwStep5_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles bgwStep5.ProgressChanged
        Try
            If Me.TipeProgress = EnumTipeProgress.ChangePrgStyle2 Then
                Me.ChangeProgressBar2()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwStep5_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgwStep5.RunWorkerCompleted
        Try
            Me.prgStep5.Value = 0
            Me.LblPrgStep5.Text = "Finished"
            Me.bgwStep6.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwStep6_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgwStep6.DoWork
        Dim DocNo As String = "RESET2017-5"
        Dim TipeStockInOut As clsTransaksiStockInOut.EnumTipeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust
        Dim TipeSC As ClsTransaksiStockCard.EnumJenisTransaksiAsal = ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockIn_Adjust
        Dim InOut As ClsTransaksiStockCard.EnumJenisStockCard = ClsTransaksiStockCard.EnumJenisStockCard.Stock_In
        Dim PrimaryLama As Double = -1
        Dim Tgl As Date = Me.dtpResetStock.Value.Date + TimeSpan.FromDays(1)
        Dim DTAHeader As New DataSet2TableAdapters.tbltransaksi_stockinout1TableAdapter
        Dim DTASC As New DataSet2TableAdapters.tbltransaksi_stock_card1TableAdapter
        Dim DTACustom As New DataSet2TableAdapters.tbltrans_custom_stockTableAdapter

        Try
            Me.ObjCD.UpdateThreading()

            Me.prgStep6.Value = 0
            Me.LblPrgStep6.Text = "Initializing..."
            Me.PrgBar = Me.prgStep6
            Me.TipePrg = ProgressBarStyle.Marquee
            Me.bgwStep6.ReportProgress(0)

            'buat stock in adjust dari positive 1
            'cari primary yang lama, hapus
            Me.PrgBar = Me.prgStep6
            Me.TipePrg = ProgressBarStyle.Marquee
            Me.bgwStep6.ReportProgress(0)
            Me.LblPrgStep6.Text = "Deleting old data..."
            If DTAHeader.ScalarQuery_GetPrimary(DocNo).HasValue = True Then
                PrimaryLama = DTAHeader.ScalarQuery_GetPrimary(DocNo).Value
                'hapus header + item
                DTAHeader.Delete(PrimaryLama)
                'hapus stock card
                DTASC.Delete(PrimaryLama, TipeSC)
            End If

            'masukkan data header
            Me.LblPrgStep6.Text = "Inserting header data..."
            DTAHeader.Insert(DocNo, 0, CInt(TipeStockInOut), Tgl, -1, -1, -1, -1, 0, -1, "", "USER", Now, "USER", Now, Now, "", -1, 0, -1, 0, 0, -1)
            PrimaryLama = DTAHeader.ScalarQuery_GetPrimary(DocNo).Value

            'masukkan data item -> qty harus abs
            Me.LblPrgStep6.Text = "Inserting item data..."
            Me.TipePrg = ProgressBarStyle.Continuous
            Me.bgwStep6.ReportProgress(0)
            Me.prgStep6.Value = 0
            Me.prgStep6.Maximum = Me.TBLNegative.Rows.Count
            Me.Simpan_ItemInOut_UsingScript(Me.TBLPositive1, PrimaryLama, Me.prgStep6, Me.LblPrgStep6)
            'masukkan data stock card
            Me.prgStep6.Value = 0
            Me.prgStep6.Maximum = Me.TBLPositive1.Rows.Count
            Me.Simpan_ItemSC_UsingScript(Me.TBLPositive1, PrimaryLama, TipeSC, DocNo, InOut, Tgl, Me.prgStep6, Me.LblPrgStep6)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwStep6_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles bgwStep6.ProgressChanged
        Try
            If Me.TipeProgress = EnumTipeProgress.ChangePrgStyle2 Then
                Me.ChangeProgressBar2()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwStep6_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgwStep6.RunWorkerCompleted
        Try
            Me.prgStep6.Value = 0
            Me.LblPrgStep6.Text = "Finished"
            MsgBox("Data updated", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDelAdj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelAdj.Click
        Dim DTAHeader As New DataSet2TableAdapters.tbltransaksi_stockinout1TableAdapter
        Dim DocNo As String = ""
        Dim PrimaryLama As Double = -1
        Dim TipeSC As ClsTransaksiStockCard.EnumJenisTransaksiAsal = ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockIn_Adjust
        Dim DTASC As New DataSet2TableAdapters.tbltransaksi_stock_card1TableAdapter

        Try
            DocNo = "RESET2017-1"
            TipeSC = ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockOut_Adjust
            If DTAHeader.ScalarQuery_GetPrimary(DocNo).HasValue = True Then
                PrimaryLama = DTAHeader.ScalarQuery_GetPrimary(DocNo).Value
                'hapus header + item
                DTAHeader.Delete(PrimaryLama)
                'hapus stock card
                DTASC.Delete(PrimaryLama, TipeSC)
            End If

            DocNo = "RESET2017-2"
            TipeSC = ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockIn_Adjust
            If DTAHeader.ScalarQuery_GetPrimary(DocNo).HasValue = True Then
                PrimaryLama = DTAHeader.ScalarQuery_GetPrimary(DocNo).Value
                'hapus header + item
                DTAHeader.Delete(PrimaryLama)
                'hapus stock card
                DTASC.Delete(PrimaryLama, TipeSC)
            End If

            DocNo = "RESET2017-3"
            TipeSC = ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockIn_Adjust
            If DTAHeader.ScalarQuery_GetPrimary(DocNo).HasValue = True Then
                PrimaryLama = DTAHeader.ScalarQuery_GetPrimary(DocNo).Value
                'hapus header + item
                DTAHeader.Delete(PrimaryLama)
                'hapus stock card
                DTASC.Delete(PrimaryLama, TipeSC)
            End If

            DocNo = "RESET2017-4"
            TipeSC = ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockOut_Adjust
            If DTAHeader.ScalarQuery_GetPrimary(DocNo).HasValue = True Then
                PrimaryLama = DTAHeader.ScalarQuery_GetPrimary(DocNo).Value
                'hapus header + item
                DTAHeader.Delete(PrimaryLama)
                'hapus stock card
                DTASC.Delete(PrimaryLama, TipeSC)
            End If

            DocNo = "RESET2017-5"
            TipeSC = ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockIn_Adjust
            If DTAHeader.ScalarQuery_GetPrimary(DocNo).HasValue = True Then
                PrimaryLama = DTAHeader.ScalarQuery_GetPrimary(DocNo).Value
                'hapus header + item
                DTAHeader.Delete(PrimaryLama)
                'hapus stock card
                DTASC.Delete(PrimaryLama, TipeSC)
            End If

            MsgBox("Data Deleted", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        MsgBox("Liat tab, reverse adjustment", MsgBoxStyle.Information, "Information")
    End Sub

    Private Sub btnStep8_1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStep8_1.Click
        Dim DTA As New DataSet2TableAdapters.tbltransaksi_stockinout2TableAdapter
        Dim TBL As New DataSet2.tbltransaksi_stockinout2DataTable
        Dim TBL1 As New DataSet2.tbltransaksi_stockinout2DataTable
        Dim row, row1, row2 As DataSet2.tbltransaksi_stockinout2Row
        Dim StartDate, EndDate As Date
        Dim RowDGV As DataGridViewRow
        Dim TipeStockInOut As clsTransaksiStockInOut.EnumTipeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust
        Dim DTAAdj As New DataSet2TableAdapters.tbltrans_custom_adjustmentTableAdapter
        Dim TBLAdj As New DataSet2.tbltrans_custom_adjustmentDataTable
        Dim rowAdj As DataSet2.tbltrans_custom_adjustmentRow

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.dgvAdjustment.ClearSelection()

            'load data berdasarkan remarks header 
            StartDate = #1/1/2017#
            EndDate = #12/31/2017#
            DTA.FillBy_RemarksHeader(TBL, EndDate, StartDate, "%closing%")
            'load data berdasarkan remarks item 
            TBL1.Rows.Clear()
            DTA.FillBy_RemarksItem(TBL1, EndDate, StartDate, "%closing%")
            For ij As Double = TBL1.Rows.Count - 1 To 0 Step -1
                row2 = TBL1.Rows(ij)
                For Each row1 In TBL.Rows
                    If row1.Primary = row2.Primary Then
                        TBL1.Rows.Remove(row2)
                        Exit For
                    End If
                Next
            Next
            If TBL1.Rows.Count > 0 Then
                For Each row2 In TBL1.Rows
                    row1 = TBL.Newtbltransaksi_stockinout2Row
                    For ij As Integer = 0 To TBL.Columns.Count - 1
                        row1.Item(ij) = row2.Item(ij)
                    Next
                    TBL.Rows.Add(row1)
                Next
            End If

            'load data berdasarkan batch item
            TBL1.Rows.Clear()
            DTA.FillBy_BatchItem(TBL1, EndDate, StartDate, "%closing%")
            For ij As Double = TBL1.Rows.Count - 1 To 0 Step -1
                row2 = TBL1.Rows(ij)
                For Each row1 In TBL.Rows
                    If row1.Primary = row2.Primary Then
                        TBL1.Rows.Remove(row2)
                        Exit For
                    End If
                Next
            Next
            If TBL1.Rows.Count > 0 Then
                For Each row2 In TBL1.Rows
                    row1 = TBL.Newtbltransaksi_stockinout2Row
                    For ij As Integer = 0 To TBL.Columns.Count - 1
                        row1.Item(ij) = row2.Item(ij)
                    Next
                    TBL.Rows.Add(row1)
                Next
            End If

            'tampilkan ke dalam dgv
            TBL.DefaultView.RowFilter = ""
            TBL.DefaultView.Sort = "[Tanggal Transaksi], [No Transaksi]"
            Me.dgvAdjustment.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvAdjustment.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.dgvAdjustment.Rows.Clear()
            For Each DV As Data.DataRowView In TBL.DefaultView
                row = DV.Row
                Me.dgvAdjustment.Rows.Add()
                RowDGV = Me.dgvAdjustment.Rows(Me.dgvAdjustment.Rows.Count - 1)
                RowDGV.Cells(Me.ColDGVAdjust_CreatedBy.Index).Value = row.Created_By
                RowDGV.Cells(Me.ColDGVAdjust_CreatedDate.Index).Value = Format(row.Created_Date, "ddd, dd MMM yyyy HH:mm")
                RowDGV.Cells(Me.ColDGVAdjust_Date.Index).Value = Format(row.Tanggal_Transaksi, "dd MMM yyyy")
                RowDGV.Cells(Me.ColDGVAdjust_DocNo.Index).Value = row.No_Transaksi
                RowDGV.Cells(Me.ColDGVAdjust_Include.Index).Value = Boolean.FalseString
                RowDGV.Cells(Me.ColDGVAdjust_ModifiedDate.Index).Value = Format(row.Modified_Date, "ddd, dd MMM yyyy HH:mm")
                RowDGV.Cells(Me.ColDGVAdjust_No.Index).Value = RowDGV.Index + 1
                RowDGV.Cells(Me.ColDGVAdjust_Remarks.Index).Value = row.Remarks
                If row.Tipe_Transaksi = CInt(TipeStockInOut) Then
                    RowDGV.Cells(Me.ColDGVAdjust_Type.Index).Value = 0
                Else
                    RowDGV.Cells(Me.ColDGVAdjust_Type.Index).Value = 1
                End If
                RowDGV.Cells(Me.ColDGVAdjust_Primary.Index).Value = row.Primary
            Next
            Me.dgvAdjustment.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvAdjustment.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            'tampilkan check list
            DTAAdj.Fill(TBLAdj)
            For Each rowAdj In TBLAdj.Rows
                For Each RowDGV In Me.dgvAdjustment.Rows
                    If rowAdj.primary_trans = RowDGV.Cells(Me.ColDGVAdjust_Primary.Index).Value Then
                        RowDGV.Cells(Me.ColDGVAdjust_Include.Index).Value = Boolean.TrueString
                        Exit For
                    End If
                Next
            Next

            'pilih baris pertama
            If Me.dgvAdjustment.Rows.Count > 0 Then
                Me.dgvAdjustment.Rows(0).Selected = True
            End If

            Me.dgvAdjustmentItem2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvAdjustmentItem2.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.dgvAdjustmentItem2.Rows.Clear()
            Me.dgvAdjustmentItem2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvAdjustmentItem2.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvAdjustment_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvAdjustment.SelectionChanged
        Dim DTAItem As New DataSet2TableAdapters.tbltransaksi_stockinout_item2TableAdapter
        Dim TBLItem As New DataSet2.tbltransaksi_stockinout_item2DataTable
        Dim rowItem As DataSet2.tbltransaksi_stockinout_item2Row
        Dim rowDGV, rowDGV1 As DataGridViewRow

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.dgvAdjustmentItem1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvAdjustmentItem1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            If Me.dgvAdjustment.SelectedCells.Count = 0 Then
                Me.dgvAdjustmentItem1.Rows.Clear()
            Else
                Me.dgvAdjustmentItem1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.dgvAdjustmentItem1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                rowDGV = Me.dgvAdjustment.Rows(Me.dgvAdjustment.SelectedCells(0).RowIndex)
                DTAItem.Fill(TBLItem, rowDGV.Cells(Me.ColDGVAdjust_Primary.Index).Value)
                For Each rowItem In TBLItem.Rows
                    Me.dgvAdjustmentItem1.Rows.Add()
                    rowDGV1 = Me.dgvAdjustmentItem1.Rows(Me.dgvAdjustmentItem1.Rows.Count - 1)
                    rowDGV1.Cells(Me.ColAdjItem1_Batch.Index).Value = rowItem.Batch_No
                    rowDGV1.Cells(Me.ColAdjItem1_CodeNo.Index).Value = rowItem.Product_Code
                    rowDGV1.Cells(Me.ColAdjItem1_Desc.Index).Value = rowItem.Description
                    rowDGV1.Cells(Me.ColAdjItem1_Expiry.Index).Value = Format(rowItem.Expiry_Date, "dd MMM yyyy")
                    rowDGV1.Cells(Me.ColAdjItem1_Location.Index).Value = rowItem.Storage_Area_Code
                    rowDGV1.Cells(Me.ColAdjItem1_No.Index).Value = rowDGV1.Index + 1
                    rowDGV1.Cells(Me.ColAdjItem1_Qty.Index).Value = Me.ObjCD.BuangNolDibelakangDecimal(rowItem.Qty)
                    rowDGV1.Cells(Me.ColAdjItem1_Unit.Index).Value = rowItem.Unit_Code
                    rowDGV1.Cells(Me.ColAdjItem1_Remarks.Index).Value = rowItem.Remarks
                Next
                Me.dgvAdjustmentItem1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                Me.dgvAdjustmentItem1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            End If
            Me.dgvAdjustmentItem1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvAdjustmentItem1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Dim RowDGV As DataGridViewRow

        Try
            For Each RowDGV In Me.dgvAdjustment.Rows
                RowDGV.Cells(Me.ColDGVAdjust_Include.Index).Value = Boolean.TrueString
            Next
            Me.dgvAdjustmentItem2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvAdjustmentItem2.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.dgvAdjustmentItem2.Rows.Clear()
            Me.dgvAdjustmentItem2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvAdjustmentItem2.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnStep8_2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStep8_2.Click
        Dim DTA As New DataSet2TableAdapters.tbltrans_custom_adjustmentTableAdapter
        Dim rowDGV As DataGridViewRow

        Try
            If MsgBox("Save to db?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                'hapus db
                DTA.Delete()

                'simpan data
                For Each rowDGV In Me.dgvAdjustment.Rows
                    If rowDGV.Cells(Me.ColDGVAdjust_Include.Index).Value = Boolean.TrueString Then
                        DTA.Insert(rowDGV.Cells(Me.ColDGVAdjust_Primary.Index).Value, rowDGV.Cells(Me.ColDGVAdjust_Type.Index).Value)
                    End If
                Next

                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnStep8_3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStep8_3.Click
        Dim DTA As New DataSet2TableAdapters.tbltransaksi_stockinout_item2TableAdapter
        Dim TBL As New DataSet2.tbltransaksi_stockinout_item2DataTable
        Dim TBLDummy As New DataSet2.tbltransaksi_stockinout_item2DataTable
        Dim row, rowDummy As DataSet2.tbltransaksi_stockinout_item2Row
        Dim rowDGV As DataGridViewRow
        Dim Ketemu As Boolean = False

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            For Each rowDGV In Me.dgvAdjustment.Rows
                If rowDGV.Cells(Me.ColDGVAdjust_Include.Index).Value = Boolean.TrueString Then
                    'fill tbl dummy
                    TBLDummy.Rows.Clear()
                    DTA.Fill(TBLDummy, rowDGV.Cells(Me.ColDGVAdjust_Primary.Index).Value)

                    'tentukan apakah reverse
                    If rowDGV.Cells(Me.ColDGVAdjust_Type.Index).Value = 1 Then
                        For Each rowDummy In TBLDummy.Rows
                            rowDummy.Qty = rowDummy.Qty * -1
                        Next
                    End If

                    'masukkan ke tabel 
                    For Each rowDummy In TBLDummy.Rows
                        Ketemu = False

                        For Each row In TBL.Rows
                            If row.Primary_Product = rowDummy.Primary_Product And row.Batch_No = rowDummy.Batch_No And _
                            row.Primary_Storage_Area = rowDummy.Primary_Storage_Area Then
                                Ketemu = True
                                row.Qty += rowDummy.Qty
                            End If
                        Next

                        If Ketemu = False Then
                            'tambahkan baris baru
                            row = TBL.Newtbltransaksi_stockinout_item2Row
                            For ij As Integer = 0 To TBL.Columns.Count - 1
                                row.Item(ij) = rowDummy.Item(ij)
                            Next
                            TBL.Rows.Add(row)
                        Else
                            If row.Qty = 0 Then
                                TBL.Rows.Remove(row)
                            End If
                        End If
                    Next
                End If
            Next

            'tampilkan data table ke dgv
            Me.dgvAdjustmentItem2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvAdjustmentItem2.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.dgvAdjustmentItem2.Rows.Clear()
            For Each row In TBL.Rows
                If row.Qty <> 0 Then
                    Me.dgvAdjustmentItem2.Rows.Add()
                    rowDGV = Me.dgvAdjustmentItem2.Rows(Me.dgvAdjustmentItem2.Rows.Count - 1)
                    rowDGV.Cells(Me.ColItemAdj2_BatchNo.Index).Value = row.Batch_No
                    rowDGV.Cells(Me.ColItemAdj2_CodeNo.Index).Value = row.Product_Code
                    rowDGV.Cells(Me.ColItemAdj2_Desc.Index).Value = row.Description
                    rowDGV.Cells(Me.ColItemAdj2_Expiry.Index).Value = Format(row.Expiry_Date, "dd MMM yyyy")
                    rowDGV.Cells(Me.ColItemAdj2_Location.Index).Value = row.Storage_Area_Code
                    rowDGV.Cells(Me.ColItemAdj2_No.Index).Value = rowDGV.Index + 1
                    rowDGV.Cells(Me.ColItemAdj2_Qty.Index).Value = Me.ObjCD.BuangNolDibelakangDecimal(row.Qty)
                    rowDGV.Cells(Me.ColItemAdj2_Unit.Index).Value = row.Unit_Code
                End If

            Next
            Me.dgvAdjustmentItem2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvAdjustmentItem2.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnStep8_4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStep8_4.Click
        Dim rowDGV As DataGridViewRow
        Dim TglIn, TglOut, TglTujuan As Date
        Dim TipeSC As ClsTransaksiStockCard.EnumJenisTransaksiAsal = ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockIn_Adjust
        Dim DTASC As New DataSet2TableAdapters.tbltransaksi_stock_cardTableAdapter
        Dim DTA1 As New DataSet2TableAdapters.tbltransaksi_stockinoutTableAdapter

        Try
            TglIn = Me.ObjCD.GetFirstDateofTheMonth(Me.dtp_Step8_4.Value)
            TglOut = Me.ObjCD.GetLasDateOfTheMonth(Me.dtp_Step8_4.Value) - TimeSpan.FromDays(1)

            For Each rowDGV In Me.dgvAdjustment.Rows
                If rowDGV.Cells(Me.ColDGVAdjust_Include.Index).Value = Boolean.TrueString Then
                    If rowDGV.Cells(Me.ColDGVAdjust_Type.Index).Value = 1 Then
                        TipeSC = ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockOut_Adjust
                        TglTujuan = TglOut
                    Else
                        TipeSC = ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockIn_Adjust
                        TglTujuan = TglIn
                    End If

                    DTASC.UpdateQuery(TglTujuan, rowDGV.Cells(Me.ColDGVAdjust_Primary.Index).Value, TipeSC)
                    '1.3. pindahkan tanggal
                    DTA1.UpdateQuery_SetTgl(TglTujuan, rowDGV.Cells(Me.ColDGVAdjust_Primary.Index).Value)
                End If
            Next

            MsgBox("Data updated", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnStep9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStep9.Click
        Dim DTA As New DataSet2TableAdapters.DataTable4TableAdapter
        Dim TBL As New DataSet2.DataTable4DataTable
        Dim TBL1 As New DataSet2.DataTable4DataTable
        Dim row, row1 As DataSet2.DataTable4Row
        Dim rowDGV As DataGridViewRow
        Dim ListPrimaryProduct As New List(Of Double)
        Dim Frm1 As New Form2_AccCode
        Dim DTA2 As New DataSet2TableAdapters.DataTable6TableAdapter
        Dim TBL2 As New DataSet2.DataTable6DataTable
        Dim row2 As DataSet2.DataTable6Row

        Try
            If MsgBox("Load stock data?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                Exit Sub
            End If

            Frm1.ShowDialog()
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'load data dari tabel stock custom -> sudah di rekap per code barang
            'klo ada stock minus, pindahkan ke gudang yang tidak minus
            DTA.Fill(TBL1)
            For Each row1 In TBL1.Rows
                If row1.Qty < 0 Then
                    If ListPrimaryProduct.Contains(row1.Primary_Product) = False Then
                        ListPrimaryProduct.Add(row1.Primary_Product)
                    End If
                End If
            Next
            For Each PrimaryProduct As Double In ListPrimaryProduct
                TBL1.DefaultView.RowFilter = "[Primary Product] = " + PrimaryProduct.ToString
                TBL1.DefaultView.Sort = "Qty"

                While TBL1.DefaultView.Count > 1
                    row1 = TBL1.DefaultView(0).Row
                    If row1.Qty > 0 Then
                        Exit While
                    Else
                        row = TBL1.DefaultView(TBL1.DefaultView.Count - 1).Row
                        If row.Qty < 0 Then
                            Exit While
                        Else
                            If Math.Abs(row.Qty) >= Math.Abs(row1.Qty) Then
                                row.Qty += row1.Qty
                                TBL1.Rows.Remove(row1)
                                If row.Qty <= 0 Then
                                    TBL1.Rows.Remove(row)
                                End If
                            Else
                                row1.Qty += row.Qty
                                TBL1.Rows.Remove(row)
                                If row1.Qty = 0 Then
                                    TBL1.Rows.Remove(row1)
                                End If
                            End If
                        End If
                    End If
                End While
            Next

            'pindah tabel
            For Each row1 In TBL1.Rows
                row = TBL.NewDataTable4Row
                For ij As Integer = 0 To TBL.Columns.Count - 1
                    row.Item(ij) = row1.Item(ij)
                Next
                TBL.Rows.Add(row)
            Next

            'kumpulkan m3
            DTA2.Fill(TBL2)

            'tampilkan ke dalam dgv
            Me.dgvFinal.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvFinal.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.dgvFinal.Rows.Clear()
            For Each row In TBL.Rows
                Me.dgvFinal.Rows.Add()
                rowDGV = Me.dgvFinal.Rows(Me.dgvFinal.Rows.Count - 1)
                rowDGV.Cells(Me.ColFinal_BrandCode.Index).Value = row.Brand_Code
                rowDGV.Cells(Me.ColFinal_BrandDesc.Index).Value = row.Brand_Desc
                rowDGV.Cells(Me.ColFinal_CategoryCode.Index).Value = row.Category_Code
                rowDGV.Cells(Me.ColFinal_CategoryDesc.Index).Value = row.Category_Desc
                rowDGV.Cells(Me.ColFinal_CodeNo.Index).Value = row.Product_Code
                rowDGV.Cells(Me.ColFinal_Desc.Index).Value = row.Description
                rowDGV.Cells(Me.ColFinal_GroupCode.Index).Value = row.Group_Code
                rowDGV.Cells(Me.ColFinal_GroupDesc.Index).Value = row.Group_Desc
                rowDGV.Cells(Me.ColFinal_No.Index).Value = rowDGV.Index + 1
                rowDGV.Cells(Me.ColFinal_Qty.Index).Value = Me.ObjCD.BuangNolDibelakangDecimal(row.Qty)
                rowDGV.Cells(Me.ColFinal_Amouunt.Index).Value = 0
                rowDGV.Cells(Me.ColFinal_SubBrandCode.Index).Value = row.Sub_Brand
                rowDGV.Cells(Me.ColFinal_SubBrandDesc.Index).Value = row.Sub_Brand_Desc
                rowDGV.Cells(Me.ColFinal_Unit.Index).Value = row.Unit_Code
                rowDGV.Cells(Me.ColFinal_UnitPrice.Index).Value = 0
                rowDGV.Cells(Me.ColFinal_Warehouse.Index).Value = row.Warehouse_Code
                rowDGV.Cells(Me.ColFinal_WarehouseDesc.Index).Value = row.Warehouse_Desc
                rowDGV.Cells(Me.ColFinal_AccCode.Index).Value = row.Account_Code
                rowDGV.Cells(Me.ColFinal_AccDesc.Index).Value = row.Account_Desc

                TBL2.DefaultView.RowFilter = "[Primary Product] = " + row.Primary_Product.ToString
                If TBL2.DefaultView.Count = 0 Then
                    rowDGV.Cells(Me.ColFinal_Qtym3.Index).Value = 0
                Else
                    row2 = TBL2.DefaultView(0).Row
                    rowDGV.Cells(Me.ColFinal_Qtym3.Index).Value = Me.ObjCD.DigitGrouping((row.Qty * row2.Qty_Rasio) / row2.Qty_Default, row2.Default_Precision)
                End If
            Next
            Me.dgvFinal.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvFinal.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            MsgBox("Data loaded", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnStep12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStep12.Click
        Dim ij As New Form2_Bundling

        Try
            ij.ShowDialog()
            ij.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnStep13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStep13.Click
        Dim PrimaryClosing As Double = -1
        Dim ObjClosing As New ClsTransaksiMonthlyClosingItem
        Dim RefDate As Date = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpStep13.Value.Date)

        Try
            PrimaryClosing = ObjClosing.GetPrimaryClosing(RefDate)
            If PrimaryClosing <> -1 Then
                If MsgBox("Check closing amount for: " + Format(RefDate, "MMM yyyy"), MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Dim ij As New Form2_EndAmount2
                    ij.Period = Me.dtpStep13.Value.Date
                    ij.ShowDialog()
                End If
            Else
                MsgBox("Please do monthly closing first: " + Format(RefDate, "MMM yyyy"), MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

        'Dim Dummy As String = ""
        'Dim DTA As New DataSet2_EndAmountTableAdapters.tbltrans_custom_sch_cost1TableAdapter
        'Dim TBL As New DataSet2_EndAmount.tbltrans_custom_sch_cost1DataTable
        'Dim row As DataSet2_EndAmount.tbltrans_custom_sch_cost1Row
        'Dim JmlClosing As Decimal = 0
        'Dim DTAClosing As New DataSet2_EndAmountTableAdapters.tbltrans_purchase_closingTableAdapter
        'Dim TBLClosing As New DataSet2_EndAmount.tbltrans_purchase_closingDataTable
        'Dim rowClosing As DataSet2_EndAmount.tbltrans_purchase_closingRow

        'Try
        '    DTA.Connection.ConnectionString = Core_Module.GetConstring
        '    DTAClosing.Connection.ConnectionString = DTA.Connection.ConnectionString

        '    PrimaryClosing = ObjClosing.GetPrimaryClosing(RefDate)
        '    If PrimaryClosing <> -1 Then
        '        If MsgBox("Check closing amount for: " + Format(RefDate, "MMM yyyy"), MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
        '            'load data reference amount
        '            DTA.FillBy_Period(TBL, RefDate)
        '            If TBL.Rows.Count = 0 Then
        '                MsgBox("No reference account data.", MsgBoxStyle.Information, "Information")
        '            End If
        '            DTAClosing.Fill(TBLClosing, PrimaryClosing)

        '            'load data closing per account
        '            JmlClosing = 0
        '            For Each row In TBL.Rows
        '                For Each rowClosing In TBLClosing.Rows
        '                    If rowClosing.Primary_Inventory_Account = row.primary_akun Then
        '                        If rowClosing.Ending_Amount <> row.amount Then
        '                            If Dummy = "" Then
        '                                Dummy = row.Account_Code
        '                            Else
        '                                Dummy = Dummy + ", " + row.Account_Code
        '                            End If
        '                        End If
        '                        Exit For
        '                    End If
        '                Next
        '            Next

        '            If Dummy <> "" Then
        '                MsgBox("Please check the following account code:" + Dummy, MsgBoxStyle.Information, "Information")
        '            Else
        '                MsgBox("Data checked", MsgBoxStyle.Information, "Information")
        '            End If
        '        End If
        '    Else
        '        MsgBox("Please do monthly closing first: " + Format(RefDate, "MMM yyyy"), MsgBoxStyle.Information, "Information")
        '    End If
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'End Try
    End Sub

    Private Sub btnStep8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStep11.Click
        Dim ij As New Form2_EndAmount

        Try
            ij.ShowDialog()
            ij.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnStep14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStep14.Click
        Dim DTAClosing As New DataSet2_CekSelisihQtyTableAdapters.tbltrans_purchase_closingTableAdapter
        Dim DTAClosingItem As New DataSet2_CekSelisihQtyTableAdapters.DataTable1TableAdapter
        Dim TBLClosingItem As New DataSet2_CekSelisihQty.DataTable1DataTable
        Dim RowClosingItem As DataSet2_CekSelisihQty.DataTable1Row
        Dim DTASC As New DataSet2_CekSelisihQtyTableAdapters.DataTable2TableAdapter
        Dim TBLSC As New DataSet2_CekSelisihQty.DataTable2DataTable
        Dim rowSC As DataSet2_CekSelisihQty.DataTable2Row
        Dim PrimaryClosing As Nullable(Of Double)
        Dim Lanjut As Boolean = False

        Try
            'samakan tanggal period
            Me.dtpStep14.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpStep14.Value.Date)
            'pastikan ada closing pada periode yang dimaksud
            PrimaryClosing = DTAClosing.ScalarQuery_GetPrimary(Me.dtpStep14.Value.Date)
            If PrimaryClosing.HasValue = True Then
                If PrimaryClosing.Value = -1 Then
                    Lanjut = False
                Else
                    Lanjut = True
                End If
            Else
                Lanjut = False
            End If

            If Lanjut = True Then
                'ingatkan user untuk menjalankan software repair stock
                If MsgBox("Please do stock repair. Continue checking?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    'load data item closing
                    DTAClosingItem.Fill(TBLClosingItem, PrimaryClosing.Value)

                    'load data item stock card
                    DTASC.Fill(TBLSC, Me.ObjCD.GetLasDateOfTheMonth(Me.dtpStep14.Value.Date))

                    For Each RowClosingItem In TBLClosingItem.Rows
                        TBLSC.DefaultView.RowFilter = "[Primary Product] = " + RowClosingItem.Primary_Product.ToString
                        If TBLSC.DefaultView.Count = 0 Then
                            RowClosingItem.Qty_SC = 0
                        Else
                            rowSC = TBLSC.DefaultView(0).Row
                            RowClosingItem.Qty_SC = rowSC.Qty_SC
                            TBLSC.Rows.Remove(rowSC)
                        End If
                    Next
                    While TBLSC.Rows.Count > 0
                        rowSC = TBLSC.Rows(0)
                        RowClosingItem = TBLClosingItem.NewDataTable1Row
                        RowClosingItem.Description = rowSC.Description
                        RowClosingItem.Ending_Qty = 0
                        RowClosingItem.Primary_Product = rowSC.Primary_Product
                        RowClosingItem.Product_Code = rowSC.Product_Code
                        RowClosingItem.Qty_SC = rowSC.Qty_SC
                        RowClosingItem.Unit_Code = rowSC.Unit_Code
                        RowClosingItem._Decimal = rowSC._Decimal
                        TBLClosingItem.Rows.Add(RowClosingItem)
                        TBLSC.Rows.Remove(rowSC)
                    End While

                    'jika ada qty selisih, tampilkan ke layar
                    For ij As Double = TBLClosingItem.Rows.Count - 1 To 0 Step -1
                        RowClosingItem = TBLClosingItem.Rows(ij)
                        If RowClosingItem.Qty_SC = RowClosingItem.Ending_Qty Then
                            TBLClosingItem.Rows.Remove(RowClosingItem)
                        End If
                    Next
                    If TBLClosingItem.Rows.Count > 0 Then
                        If MsgBox("Discrepancy found, show detail?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                            Dim frm1 As New Form2_SelisihQty
                            frm1.TBLRef = TBLClosingItem
                            frm1.ShowDialog()
                        End If
                    Else : MsgBox("Data checked. Everything is ok", MsgBoxStyle.Information, "Information")
                    End If
                End If
            Else
                MsgBox("No closing detected on specified period", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnStep15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStep15.Click
        Try
            If MsgBox("Delete item aop w/o grn?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Control.CheckForIllegalCrossThreadCalls = False
                Me.bgwStep15.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwStep15_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgwStep15.DoWork
        Dim DTA As New DataSet2_CekSelisihQtyTableAdapters.DataTable3TableAdapter
        Dim TBL As New DataSet2_CekSelisihQty.DataTable3DataTable
        Dim row As DataSet2_CekSelisihQty.DataTable3Row

        Try
            Me.ObjCD.UpdateThreading()

            'fill tabel
            Me.LblPrgStep15.Text = "Preparing data"
            DTA.Fill(TBL)

            'for each row, hapus
            Me.prgStep15.Value = 0
            Me.prgStep15.Maximum = TBL.Rows.Count
            For Each row In TBL.Rows
                Me.prgStep15.Value += 1
                Me.LblPrgStep15.Text = "Deleting data " + Me.prgStep15.Value.ToString + " of " + Me.prgStep15.Maximum.ToString + " - " + row.No_Transaksi
                DTA.DeleteQuery_ItemAOP(row.Primary)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwStep15_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgwStep15.RunWorkerCompleted
        Try
            Me.prgStep15.Value = 0
            Me.LblPrgStep15.Text = "Ready"
            MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnStep16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStep16.Click
        Try
            If MsgBox("Create temporary closing?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Control.CheckForIllegalCrossThreadCalls = False
                Me.bgwStep16.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwStep16_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgwStep16.DoWork
        Dim DTA As New DataSet2_EndAmountTableAdapters.tbltrans_purchase_closing_itemTableAdapter
        Dim RefPeriod As Date = #11/1/2015#
        Dim Jml As Integer = 0
        Dim TargetDate As Date = #12/1/2017#

        Try
            ObjCD.UpdateThreading()

            'get periode max
            If DTA.ScalarQuery_GetMaxClosingPeriod.HasValue = True Then
                RefPeriod = DTA.ScalarQuery_GetMaxClosingPeriod.Value
            End If
            If RefPeriod < TargetDate Then
                While RefPeriod + TimeSpan.FromDays(30 * Jml) < TargetDate
                    Jml += 1
                End While

                Me.prgStep16.Value = 0
                Me.prgStep16.Maximum = Jml
                'mulai dari periode max+1 s.d dec 2017 -> buat closing boongan
                While RefPeriod < TargetDate
                    Me.prgStep16.Value += 1
                    RefPeriod = Me.ObjCD.GetFirstDateofTheMonth(RefPeriod + TimeSpan.FromDays(32))
                    Me.LblPrgStep16.Text = "Creating " + Format(RefPeriod, "MMM yyyy")

                    DTA.InsertQuery(Format(RefPeriod, "MM/yyyy"), Today, RefPeriod, "", "USER", Now, "USER", Now, -1)
                End While
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwStep16_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgwStep16.RunWorkerCompleted
        Try
            Me.prgStep16.Value = 0
            Me.LblPrgStep16.Text = "Ready"
            MsgBox("Data created", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnStep17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStep17.Click
        Try
            If MsgBox("Delete temporary closing?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Control.CheckForIllegalCrossThreadCalls = False
                Me.bgwStep17.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwStep17_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgwStep17.DoWork
        Dim DTA As New DataSet2_EndAmountTableAdapters.DataTable3TableAdapter
        Dim TBL As New DataSet2_EndAmount.DataTable3DataTable
        Dim row As DataSet2_EndAmount.DataTable3Row

        Try
            ObjCD.UpdateThreading()

            'get periode max
            DTA.Fill(TBL)
            Me.prgStep17.Value = 0
            Me.prgStep17.Maximum = TBL.Rows.Count
            For Each row In TBL.Rows
                Me.prgStep17.Value += 1 
                Me.LblPrgStep17.Text = "Deleting " + Me.prgStep17.Value.ToString + " of " + Me.prgStep17.Maximum.ToString
                DTA.DeleteQuery(row.Primary)
            Next 
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwStep17_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgwStep17.RunWorkerCompleted
        Try
            Me.prgStep17.Value = 0
            Me.LblPrgStep17.Text = "Ready"
            MsgBox("Data created", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnStep18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStep18.Click
        Try
            Dim ij As New Form2_EditBundling
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnStep19_Set_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStep19_Set.Click
        Dim DTA As New DataSet2_EndAmountTableAdapters.tbl_smessTableAdapter

        Try
            'insert 1 baris untuk tabel license smess
            DTA.Insert(1)
            MsgBox("Data updated", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnStep1_RemoveMaintenance_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStep1_RemoveMaintenance.Click
        Dim DTA As New DataSet2_EndAmountTableAdapters.tbl_smessTableAdapter

        Try
            'remove baris pada tabel smess
            DTA.Delete()
            MsgBox("Data updated", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnStep20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStep20.Click
        Dim ij As New Form2_CekBundling

        Try
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub btnCek_Bundling_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCek_Bundling.Click
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class