Option Explicit On
Option Strict Off

Partial Public Class FrmTools_MonthlyClosing
    Private Sub T3_btnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles T3_btnFilter.Click
        Dim DTA As New DseToolsTableAdapters.DataTable4TableAdapter
        Dim TBL As New DseTools.DataTable4DataTable
        Dim row As DseTools.DataTable4Row
        Dim DTABundling As New DseToolsTableAdapters.tbltransaksi_stockbundlingTableAdapter
        Dim TBLBundling As New DseTools.tbltransaksi_stockbundlingDataTable
        Dim rowDGV As DataGridViewRow
        Dim PrimaryClosing As Double = -1
        Dim PrimaryTransaksiAsal As Double = -1
        Dim TipeIn As Purchase_Module.ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal = Purchase_Module.ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling
        Dim JmlIn, JmlOut, AllocatedAmount As Decimal
        Dim BundlingNo As String = ""
        Dim DefPrec As Integer = Me.ObjCurr.GetDecimal_DefaultCurrency
        Dim TotalSelisih As Decimal = 0
        Dim DTA_AllocatedAmount As New DseToolsTableAdapters.tbltrans_purchase_closing_item_cost_group_code_noTableAdapter
        Dim TBL_AllocatedAmount As New DseTools.tbltrans_purchase_closing_item_cost_group_code_noDataTable
        Dim row_AllocatedAmount As DseTools.tbltrans_purchase_closing_item_cost_group_code_noRow
        Dim TotalAllocatedAmount As Decimal = 0
        Dim TipeTransaksiAsal1 As Purchase_Module.ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal = Purchase_Module.ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling
        Dim TipeTransaksiAsal2 As Purchase_Module.ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal = Purchase_Module.ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling
        Dim TipeItem1 As Purchase_Module.ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail = Purchase_Module.ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.InItem
        Dim TipeItem2 As Purchase_Module.ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail = Purchase_Module.ClsTransaksiMonthlyClosingItem.EnumTipeItemDetail.OutItem

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.ColT3_In.DefaultCellStyle.Format = "N" + DefPrec.ToString
            Me.ColT3_Out.DefaultCellStyle.Format = "N" + DefPrec.ToString
            Me.ColT3_Selisih.DefaultCellStyle.Format = "N" + DefPrec.ToString
            Me.ColT3_AllocatedAmount.DefaultCellStyle.Format = "N" + DefPrec.ToString

            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTABundling.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA_AllocatedAmount.Connection.ConnectionString = DTA.Connection.ConnectionString
            JmlIn = 0
            JmlOut = 0

            'load data
            PrimaryClosing = Me.ObjClosing.GetPrimary(Me.T3_dtpPeriod.Value.Date)
            DTA.FillBy(TBL, CInt(TipeItem1), CInt(TipeItem2), CInt(TipeTransaksiAsal1), CInt(TipeTransaksiAsal2), PrimaryClosing)

            If PrimaryClosing <> -1 Then
                DTABundling.Fill(TBLBundling, Me.ObjCD.GetLasDateOfTheMonth(Me.T3_dtpPeriod.Value), _
                Me.ObjCD.GetFirstDateofTheMonth(Me.T3_dtpPeriod.Value))
                DTA_AllocatedAmount.Fill(TBL_AllocatedAmount, PrimaryClosing)

                Me.T3_dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.T3_dgv.Rows.Clear()
                'for each row, jika total item in <> total in out, maka tampilkan ke dalam dgv
                For Each row In TBL.Rows
                    If row.primary_transaksi_asal <> PrimaryTransaksiAsal Then
                        'tambahkan ke dalam dgv jika selisih in dan out
                        If PrimaryTransaksiAsal <> -1 Then
                            If JmlIn <> JmlOut Then
                                TBL_AllocatedAmount.DefaultView.RowFilter = "[primary bundling] = " + PrimaryTransaksiAsal.ToString
                                AllocatedAmount = 0
                                For Each DV As DataRowView In TBL_AllocatedAmount.DefaultView
                                    row_AllocatedAmount = DV.Row
                                    AllocatedAmount += row_AllocatedAmount.allocated_amount
                                Next

                                If AllocatedAmount <> JmlIn - JmlOut Then
                                    BundlingNo = Me.GetNoBundling(PrimaryTransaksiAsal, TBLBundling)
                                    Me.T3_dgv.Rows.Add()
                                    rowDGV = Me.T3_dgv.Rows(Me.T3_dgv.Rows.Count - 1)
                                    rowDGV.Cells(Me.ColT3_BundlingNo.Index).Value = BundlingNo
                                    rowDGV.Cells(Me.ColT3_In.Index).Value = JmlIn
                                    rowDGV.Cells(Me.ColT3_No.Index).Value = Me.T3_dgv.Rows.Count
                                    rowDGV.Cells(Me.ColT3_Out.Index).Value = JmlOut
                                    rowDGV.Cells(Me.ColT3_Selisih.Index).Value = JmlIn - JmlOut
                                    TotalSelisih += (JmlIn - JmlOut)

                                    'allocated amount

                                    rowDGV.Cells(Me.ColT3_AllocatedAmount.Index).Value = AllocatedAmount
                                    TotalAllocatedAmount += AllocatedAmount
                                End If
                            End If
                        End If
                        'reset perhitungan
                        JmlIn = 0
                        JmlOut = 0
                        PrimaryTransaksiAsal = row.primary_transaksi_asal
                    End If
                    If row.tipe_item_detail = TipeItem1 Or row.tipe_item_detail = TipeItem2 Then
                        If row.tipe_transaksi_asal = CInt(TipeIn) Then
                            JmlIn += row.amount
                        Else
                            JmlOut += row.amount
                        End If
                    End If
                Next

                'cek baris terakhir
                If JmlIn <> JmlOut Then
                    TBL_AllocatedAmount.DefaultView.RowFilter = "[primary bundling] = " + PrimaryTransaksiAsal.ToString
                    AllocatedAmount = 0
                    For Each DV As DataRowView In TBL_AllocatedAmount.DefaultView
                        row_AllocatedAmount = DV.Row
                        AllocatedAmount += row_AllocatedAmount.allocated_amount
                    Next
                    If AllocatedAmount <> JmlIn - JmlOut Then
                        BundlingNo = Me.GetNoBundling(PrimaryTransaksiAsal, TBLBundling)
                        Me.T3_dgv.Rows.Add()
                        rowDGV = Me.T3_dgv.Rows(Me.T3_dgv.Rows.Count - 1)
                        rowDGV.Cells(Me.ColT3_BundlingNo.Index).Value = BundlingNo
                        rowDGV.Cells(Me.ColT3_In.Index).Value = JmlIn
                        rowDGV.Cells(Me.ColT3_No.Index).Value = Me.T3_dgv.Rows.Count
                        rowDGV.Cells(Me.ColT3_Out.Index).Value = JmlOut
                        rowDGV.Cells(Me.ColT3_Selisih.Index).Value = JmlIn - JmlOut
                        TotalSelisih += (JmlIn - JmlOut)

                        rowDGV.Cells(Me.ColT3_AllocatedAmount.Index).Value = AllocatedAmount
                        TotalAllocatedAmount += AllocatedAmount
                    End If
                End If

                Me.T3_dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                If Me.T3_dgv.Rows.Count = 0 Then
                    Me.T3_LblSelisih.Text = ""
                Else
                    Me.T3_LblSelisih.Text = "Total Selisih: " + Me.ObjCD.DigitGrouping(TotalSelisih, DefPrec) + " - " + Me.ObjCD.DigitGrouping(TotalAllocatedAmount, DefPrec)
                End If

                MsgBox("Checking finish", MsgBoxStyle.Information, "Information")
            Else
                Me.T3_dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.T3_dgv.Rows.Clear()
                Me.T3_dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                MsgBox("No closing on the specified period.", MsgBoxStyle.Information, "Informatio")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetNoBundling(ByVal vPrimaryBundling As Double, ByVal TBLBundling As DseTools.tbltransaksi_stockbundlingDataTable) As String
        Dim rowBundling As DseTools.tbltransaksi_stockbundlingRow
        Try
            TBLBundling.DefaultView.RowFilter = "Primary = " + vPrimaryBundling.ToString
            If TBLBundling.DefaultView.Count = 0 Then
                Return ""
            Else
                rowBundling = TBLBundling.DefaultView(0).Row
                Return rowBundling.No_Transaksi
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function
End Class
