Option Explicit On
Option Strict Off

Public Class Form2
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    'yang harus di ubah tiap kali mau generate laporan baru:
    '1. Header form: PrimaryTransaksiAcuan,PeriodClosing
    '2. Parameter pada fungsi form load: StartDate, EndDate, Primary Transaksi Acuan, Period Closing
    Dim vListPrimaryCategory As New List(Of Double)
    Dim PrimaryTransaksiAcuan As Double = 77
    Dim PeriodClosing As Date = #1/1/2017#
    Dim SudahLoadDGV As Boolean = False
    Dim StartTime, EndTime As Date

    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim ObjProd As New Core_Module.clsMasterProduct
    Dim ObjCat As New Core_Module.ClsMasterCategory
    Dim ObjAdjust As New Inventory_Module.clsTransaksiStockInOut
    Dim DefPrec As Integer = ObjCur.GetDecimal_DefaultCurrency
    Dim DTAStockInOutAdjust As New Dse3TableAdapters.tbltransaksi_stockinoutTableAdapter
    Dim TBLStockInOutAdjust As New Dse3.tbltransaksi_stockinoutDataTable
    Dim rowStockInOut As Dse3.tbltransaksi_stockinoutRow
    Dim DTABundling As New Dse3TableAdapters.tbltransaksi_stockbundlingTableAdapter
    Dim TblBundling As New Dse3.tbltransaksi_stockbundlingDataTable
    Dim rowBundling As Dse3.tbltransaksi_stockbundlingRow
    Dim DTAAOP As New Dse3TableAdapters.tbltrans_aopTableAdapter
    Dim TBLAOP As New Dse3.tbltrans_aopDataTable
    Dim rowAOP As Dse3.tbltrans_aopRow
    Dim DTAInv As New Dse3TableAdapters.tbltrans_salesinvoiceTableAdapter
    Dim TBLInv As New Dse3.tbltrans_salesinvoiceDataTable
    Dim rowInv As Dse3.tbltrans_salesinvoiceRow
    Dim StartDate, EndDate As Date

    Dim TBLDetail As New Inventory_Module.DseLaporanStockDetail2.DataTable1DataTable
    Dim TipePrg As Windows.Forms.ProgressBarStyle = ProgressBarStyle.Marquee
    Dim PrgBar As Windows.Forms.ProgressBar

    Dim TotalRow, CurrentRow As Double
    'Dim PrimaryProductAcuan As Double = 35855
    Dim SaveToFile As String = ""

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

    Private Sub IsikanTipeDoc2Dkk()
        Dim TBLAcuan As New Dse2.DataTable1DataTable
        Dim rowAcuan As Dse2.DataTable1Row
        Dim TipeItem1 As String = "Begining Balance"
        Dim TipeItem2 As String = "In"
        Dim TipeItem3 As String = "Out"
        Dim rowDGV As DataGridViewRow
        Dim Urutan As Double = 0
        Dim QtyAllocated As Decimal = 0
        Dim ErrorMsg As String = ""


        Try
            'isikan tabel acuan dari item2 yang tipe-nya begining balance dan stock in, secara urut
            TBLAcuan.Rows.Clear()
            For Each rowDGV In Me.DataGridView2.Rows
                If rowDGV.Cells(Me.ColItem_CodeNo.Index).Value.ToString.Contains(" '''") = True Then
                    rowDGV.Cells(Me.ColItem_CodeNo.Index).Value = rowDGV.Cells(Me.ColItem_CodeNo.Index).Value.ToString.Replace(" '''", "")
                End If
                If rowDGV.Cells(Me.ColItem_CodeNo.Index).Value.ToString.Contains("'") = True Then
                    rowDGV.Cells(Me.ColItem_CodeNo.Index).Value = rowDGV.Cells(Me.ColItem_CodeNo.Index).Value.ToString.Replace("'", "")
                End If

                If rowDGV.Cells(Me.ColItem_TipeItem.Index).Value = TipeItem1 Or _
                rowDGV.Cells(Me.ColItem_TipeItem.Index).Value = TipeItem2 Then
                    Urutan += 1

                    rowAcuan = TBLAcuan.NewDataTable1Row
                    With rowAcuan
                        .Primary = Urutan
                        .No_Dokumen_Asal = rowDGV.Cells(Me.ColItem_NoDoc.Index).Value
                        .Code_No = rowDGV.Cells(Me.ColItem_CodeNo.Index).Value
                        If rowDGV.Cells(Me.ColItem_TipeItem.Index).Value = TipeItem1 Then
                            .Qty = rowDGV.Cells(Me.ColItem_QtyBB.Index).Value
                        Else
                            .Qty = rowDGV.Cells(Me.ColItem_QtyIn.Index).Value
                        End If
                        .Tipe_Item = rowDGV.Cells(Me.ColItem_TipeItem.Index).Value
                        .Tipe_Transaksi_Asal = rowDGV.Cells(Me.ColItem_TipeDoc.Index).Value
                    End With
                    TBLAcuan.Rows.Add(rowAcuan)
                End If
            Next

            'untuk setiap item yang tipe-nya out, maka ambil tipe transaksi asal, tipe item, no doc asal
            For Each rowDGV In Me.DataGridView2.Rows
                If rowDGV.Cells(Me.ColItem_CodeNo.Index).Value.ToString.Contains(" '''") = True Then
                    rowDGV.Cells(Me.ColItem_CodeNo.Index).Value = rowDGV.Cells(Me.ColItem_CodeNo.Index).Value.ToString.Replace(" '''", "")
                End If
                If rowDGV.Cells(Me.ColItem_CodeNo.Index).Value.ToString.Contains("'") = True Then
                    rowDGV.Cells(Me.ColItem_CodeNo.Index).Value = rowDGV.Cells(Me.ColItem_CodeNo.Index).Value.ToString.Replace("'", "")
                End If
                'If rowDGV.Index = 76843 Then
                '    MsgBox("")
                'End If
                If rowDGV.Cells(Me.ColItem_TipeItem.Index).Value = TipeItem3 Then
                    rowDGV.Cells(Me.ColItem_TipeDoc2.Index).Value = ""
                    rowDGV.Cells(Me.ColItem_DocNo2.Index).Value = ""
                    rowDGV.Cells(Me.ColItem_TipeItem2.Index).Value = ""

                    TBLAcuan.DefaultView.RowFilter = "[Code No] = '" + rowDGV.Cells(Me.ColItem_CodeNo.Index).Value.ToString + "'"
                    TBLAcuan.DefaultView.Sort = "Primary"
                    If TBLAcuan.DefaultView.Count = 0 Then
                        MsgBox("Logic error", MsgBoxStyle.Information, "Information")
                        If ErrorMsg.Contains(rowDGV.Cells(Me.ColItem_CodeNo.Index).Value.ToString) = False Then
                            ErrorMsg = ErrorMsg + "; " + rowDGV.Cells(Me.ColItem_CodeNo.Index).Value.ToString
                        End If
                    Else
                        rowAcuan = TBLAcuan.DefaultView(0).Row
                        If rowAcuan.Qty < CDec(rowDGV.Cells(Me.ColItem_QtyOut.Index).Value) Then
                            QtyAllocated = 0

                            While rowDGV.Cells(Me.ColItem_QtyOut.Index).Value > QtyAllocated
                                If TBLAcuan.DefaultView.Count = 0 Then
                                    MsgBox("Logic error")
                                    If ErrorMsg.Contains(rowDGV.Cells(Me.ColItem_CodeNo.Index).Value.ToString) = False Then
                                        ErrorMsg = ErrorMsg + "; " + rowDGV.Cells(Me.ColItem_CodeNo.Index).Value.ToString
                                    End If
                                    Exit While
                                Else
                                    rowAcuan = TBLAcuan.DefaultView(0).Row

                                    If rowDGV.Cells(Me.ColItem_TipeDoc2.Index).Value = "" Then
                                        rowDGV.Cells(Me.ColItem_TipeDoc2.Index).Value = rowAcuan.Tipe_Transaksi_Asal
                                        rowDGV.Cells(Me.ColItem_DocNo2.Index).Value = rowAcuan.No_Dokumen_Asal
                                        rowDGV.Cells(Me.ColItem_TipeItem2.Index).Value = rowAcuan.Tipe_Item
                                    Else
                                        If rowDGV.Cells(Me.ColItem_TipeDoc2.Index).Value.ToString.Contains(rowAcuan.Tipe_Transaksi_Asal) = False Then
                                            rowDGV.Cells(Me.ColItem_TipeDoc2.Index).Value = rowDGV.Cells(Me.ColItem_TipeDoc2.Index).Value.ToString + ", " + rowAcuan.Tipe_Transaksi_Asal
                                        End If
                                        If rowDGV.Cells(Me.ColItem_DocNo2.Index).Value.ToString.Contains(rowAcuan.No_Dokumen_Asal) = False Then
                                            rowDGV.Cells(Me.ColItem_DocNo2.Index).Value = rowDGV.Cells(Me.ColItem_DocNo2.Index).Value + ", " + rowAcuan.No_Dokumen_Asal
                                        End If
                                        If rowDGV.Cells(Me.ColItem_TipeItem2.Index).Value.ToString.Contains(rowAcuan.Tipe_Item) = False Then
                                            rowDGV.Cells(Me.ColItem_TipeItem2.Index).Value = rowDGV.Cells(Me.ColItem_TipeItem2.Index).Value + ", " + rowAcuan.Tipe_Item
                                        End If
                                    End If

                                    If rowAcuan.Qty <= CDec(rowDGV.Cells(Me.ColItem_QtyOut.Index).Value) - QtyAllocated Then
                                        QtyAllocated += rowAcuan.Qty
                                        rowDGV.Cells(Me.ColItem_DocNo2.Index).Value = rowDGV.Cells(Me.ColItem_DocNo2.Index).Value.ToString + " (" + Me.ObjCD.DigitGrouping(rowAcuan.Qty, 2) + ")"
                                        rowAcuan.Qty = 0
                                        If rowAcuan.Qty = 0 Then
                                            TBLAcuan.Rows.Remove(rowAcuan)
                                        End If
                                    Else
                                        rowAcuan.Qty -= (CDec(rowDGV.Cells(Me.ColItem_QtyOut.Index).Value) - QtyAllocated)
                                        rowDGV.Cells(Me.ColItem_DocNo2.Index).Value = rowDGV.Cells(Me.ColItem_DocNo2.Index).Value.ToString + " (" + Me.ObjCD.DigitGrouping((CDec(rowDGV.Cells(Me.ColItem_QtyOut.Index).Value) - QtyAllocated), 2) + ")"
                                        QtyAllocated += CDec(rowDGV.Cells(Me.ColItem_QtyOut.Index).Value) - QtyAllocated
                                    End If
                                End If
                            End While
                        Else
                            rowDGV.Cells(Me.ColItem_TipeDoc2.Index).Value = rowAcuan.Tipe_Transaksi_Asal
                            rowDGV.Cells(Me.ColItem_DocNo2.Index).Value = rowAcuan.No_Dokumen_Asal
                            rowDGV.Cells(Me.ColItem_TipeItem2.Index).Value = rowAcuan.Tipe_Item

                            rowAcuan.Qty -= CDec(rowDGV.Cells(Me.ColItem_QtyOut.Index).Value)

                            If rowAcuan.Qty = 0 Then
                                TBLAcuan.Rows.Remove(rowAcuan)
                            End If
                        End If
                    End If
                End If
            Next

            'kurangkan qty-nya pada tabel ref

            'jika ada yang qty out > qty ref, maka logic error
            If ErrorMsg <> "" Then
                MsgBox("Invalid data: " + ErrorMsg, MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Recalculate(ByVal vPrimaryClosing As Double)
        Dim DTAProd As New Dse3TableAdapters.tblmaster_productTableAdapter
        Dim TBLProdDummy As New Dse3.tblmaster_productDataTable
        Dim rowProd, rowProdDummy As Dse3.tblmaster_productRow
        Dim TBLProd As New Dse3.tblmaster_productDataTable
        Dim DTAItem As New Dse3TableAdapters.tbltrans_purchase_closing_itemTableAdapter
        Dim TBLItem As New Dse3.tbltrans_purchase_closing_itemDataTable
        Dim rowItem As Dse3.tbltrans_purchase_closing_itemRow
        Dim DTAItemd As New Dse3TableAdapters.tbltrans_purchase_closing_item_detailTableAdapter
        Dim TBLItemd As New Dse3.tbltrans_purchase_closing_item_detailDataTable
        Dim rowItemd As Dse3.tbltrans_purchase_closing_item_detailRow
        Dim rowDGV As DataGridViewRow
        Dim TipeItemd As EnumTipeItemDetail = EnumTipeItemDetail.BeginingBalanceItem
        Dim TipeTransaksiAsald As EnumTipeTranskasiAsal = EnumTipeTranskasiAsal.InAdjust
        Dim DTAC As New Dse3TableAdapters.tblmaster_categoryTableAdapter
        Dim TBLC As New Dse3.tblmaster_categoryDataTable
        Dim rowC As Dse3.tblmaster_categoryRow
        Dim hg As Double = 0

        Try
            '4 parameter dibawah harus di ganti kalo mau update data
            PrimaryTransaksiAcuan = vPrimaryClosing
            Me.StartDate = Me.dtpPeriod.Value
            Me.EndDate = Me.ObjCD.GetLasDateOfTheMonth(Me.StartDate)
            PeriodClosing = Me.StartDate


            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.vListPrimaryCategory.Clear()
            DTAC.Fill(TBLC)
            For Each rowC In TBLC.Rows
                Me.vListPrimaryCategory.Add(rowC.Primary)
            Next

            TBLProd.Rows.Clear()
            For Each PrimaryCategory As Double In Me.vListPrimaryCategory
                TBLProdDummy.Rows.Clear()
                DTAProd.Fill(TBLProdDummy, PrimaryCategory)
                For Each rowProdDummy In TBLProdDummy.Rows
                    'If rowProdDummy.Primary = Me.PrimaryProductAcuan Then
                    rowProd = TBLProd.Newtblmaster_productRow
                    For ij As Integer = 0 To TBLProd.Columns.Count - 1
                        rowProd.Item(ij) = rowProdDummy.Item(ij)
                    Next
                    TBLProd.Rows.Add(rowProd)
                    'End If
                Next
            Next

            'load data tabel acuan
            DTAItem.Fill(TBLItem, PrimaryTransaksiAcuan)
            DTAItemd.Fill(TBLItemd, PrimaryTransaksiAcuan)

         
            Me.IsikanDocNo(TBLItemd)

            Me.DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DataGridView1.Rows.Clear()
            Me.DataGridView2.Rows.Clear()

            Me.TotalRow = TBLProd.Rows.Count
            Me.ProgressBar1.Value = 0
            Me.ProgressBar1.Maximum = TBLProd.Rows.Count
            For Each rowProd In TBLProd.Rows
                Me.ProgressBar1.Value += 1
                hg += 1
                Me.CurrentRow += 1
                If Me.ProgressBar1.Value > 25 Then
                    If Math.IEEERemainder(Me.ProgressBar1.Value, 100) = 0 Then
                        Me.LblProgress.Text = "Phase 2 of 5 - Calculating detail data : Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString + " : " + rowProd.Product_Code
                    ElseIf Me.ProgressBar1.Value = Me.ProgressBar1.Maximum Then
                        Me.LblProgress.Text = "Phase 2 of 5 - Calculating detail data : Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString + " : " + rowProd.Product_Code
                    End If
                Else
                    Me.LblProgress.Text = "Phase 2 of 5 - Calculating detail data : Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString + " : " + rowProd.Product_Code
                End If

                TBLItem.DefaultView.RowFilter = "[Primary Product] = " + rowProd.Primary.ToString
                If TBLItem.DefaultView.Count = 0 Then
                Else
                    'tambahkan baris baru
                    rowItem = TBLItem.DefaultView(0).Row
                    'If rowItem.Begin_Qty > 0 Or rowItem.In_Qty > 0 Or rowItem.Out_Qty > 0 Or rowItem.Ending_Qty > 0 Then
                    Me.DataGridView1.Rows.Add()
                    rowDGV = Me.DataGridView1.Rows(Me.DataGridView1.Rows.Count - 1)
                    With rowItem
                        rowDGV.Cells(Me.ColAmountBB.Index).Value = Me.ObjCD.DigitGrouping(.Begin_Amount, DefPrec)
                        rowDGV.Cells(Me.ColAmountEnd.Index).Value = Me.ObjCD.DigitGrouping(.Ending_Amount, DefPrec)
                        rowDGV.Cells(Me.ColAmountIn.Index).Value = Me.ObjCD.DigitGrouping(.In_Amount, DefPrec)
                        rowDGV.Cells(Me.ColAmountOut.Index).Value = Me.ObjCD.DigitGrouping(.Out_Amount, DefPrec)
                        rowDGV.Cells(Me.ColCategory.Index).Value = rowProd.Category_Code
                        rowDGV.Cells(Me.ColCodeNo.Index).Value = rowProd.Product_Code
                        rowDGV.Cells(Me.ColOutQty.Index).Value = Me.ObjCD.DigitGrouping(.Out_Qty, .Decimal_Product)
                        rowDGV.Cells(Me.ColQtyBB.Index).Value = Me.ObjCD.DigitGrouping(.Begin_Qty, .Decimal_Product)
                        rowDGV.Cells(Me.ColQtyEnd.Index).Value = Me.ObjCD.DigitGrouping(.Ending_Qty, .Decimal_Product)
                        rowDGV.Cells(Me.ColQtyIn.Index).Value = Me.ObjCD.DigitGrouping(.In_Qty, .Decimal_Product)
                        rowDGV.Cells(Me.ColDesc.Index).Value = rowProd.Description
                        rowDGV.Cells(Me.ColUnit.Index).Value = rowProd.Unit_Code
                        rowDGV.Cells(Me.ColDefPrec.Index).Value = rowProd._Decimal
                        rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value = rowProd.Primary
                        'rowDGV.Cells(Me.ColDecimalProd.Index).Value = .Decimal_Product
                    End With

                    'tambahkan baris detail
                    TBLItemd.DefaultView.RowFilter = "[primary code no] = " + rowProd.Primary.ToString
                    TBLItemd.DefaultView.Sort = "primary"
                    For Each DV As DataRowView In TBLItemd.DefaultView
                        rowItemd = DV.Row
                        Me.DataGridView2.Rows.Add()
                        rowDGV = Me.DataGridView2.Rows(Me.DataGridView2.Rows.Count - 1)
                        With rowItemd
                            rowDGV.Cells(Me.ColItem_No.Index).Value = rowDGV.Index + 1
                            rowDGV.Cells(Me.ColItem_AmountBB.Index).Value = Me.ObjCD.DigitGrouping(0, DefPrec)
                            rowDGV.Cells(Me.ColItem_AmountEnd.Index).Value = Me.ObjCD.DigitGrouping(0, DefPrec)
                            rowDGV.Cells(Me.ColItem_AmountIn.Index).Value = Me.ObjCD.DigitGrouping(0, DefPrec)
                            rowDGV.Cells(Me.ColItem_AmountOut.Index).Value = Me.ObjCD.DigitGrouping(0, DefPrec)
                            rowDGV.Cells(Me.ColItem_QtyBB.Index).Value = Me.ObjCD.DigitGrouping(0, rowProd._Decimal)
                            rowDGV.Cells(Me.ColItem_QtyEnd.Index).Value = Me.ObjCD.DigitGrouping(0, rowProd._Decimal)
                            rowDGV.Cells(Me.ColItem_QtyIn.Index).Value = Me.ObjCD.DigitGrouping(0, rowProd._Decimal)
                            rowDGV.Cells(Me.ColItem_QtyOut.Index).Value = Me.ObjCD.DigitGrouping(0, rowProd._Decimal)
                            rowDGV.Cells(Me.ColItem_TipeDoc.Index).Value = Me.GetTipeDocument_TransaksiAsal(rowItemd.tipe_transaksi_asal)
                            If .tipe_item_detail = CInt(EnumTipeItemDetail.BeginingBalanceItem) Then
                                rowDGV.Cells(Me.ColItem_AmountBB.Index).Value = Me.ObjCD.DigitGrouping(.amount, DefPrec)
                                rowDGV.Cells(Me.ColItem_QtyBB.Index).Value = Me.ObjCD.DigitGrouping(.qty, rowProd._Decimal)
                            ElseIf .tipe_item_detail = CInt(EnumTipeItemDetail.InItem) Then
                                rowDGV.Cells(Me.ColItem_AmountIn.Index).Value = Me.ObjCD.DigitGrouping(.amount, DefPrec)
                                rowDGV.Cells(Me.ColItem_QtyIn.Index).Value = Me.ObjCD.DigitGrouping(.qty, rowProd._Decimal)
                            ElseIf .tipe_item_detail = CInt(EnumTipeItemDetail.OutItem) Then
                                rowDGV.Cells(Me.ColItem_AmountOut.Index).Value = Me.ObjCD.DigitGrouping(.amount, DefPrec)
                                rowDGV.Cells(Me.ColItem_QtyOut.Index).Value = Me.ObjCD.DigitGrouping(.qty, rowProd._Decimal)
                            ElseIf .tipe_item_detail = CInt(EnumTipeItemDetail.EndingBalanceItem) Then
                                rowDGV.Cells(Me.ColItem_AmountEnd.Index).Value = Me.ObjCD.DigitGrouping(.amount, DefPrec)
                                rowDGV.Cells(Me.ColItem_QtyEnd.Index).Value = Me.ObjCD.DigitGrouping(.qty, rowProd._Decimal)
                            End If

                            rowDGV.Cells(Me.ColItem_Category.Index).Value = rowProd.Category_Code
                            rowDGV.Cells(Me.ColItem_CodeNo.Index).Value = rowProd.Product_Code
                            rowDGV.Cells(Me.ColItem_Desc.Index).Value = rowProd.Description
                            rowDGV.Cells(Me.ColItem_NoDoc.Index).Value = rowItemd.doc_no
                            rowDGV.Cells(Me.ColItem_Tgl.Index).Value = .tgl_transaksi_asal
                            rowDGV.Cells(Me.ColItem_Unit.Index).Value = rowProd.Unit_Code
                            rowDGV.Cells(Me.ColItem_DefPrec.Index).Value = rowProd._Decimal
                            rowDGV.Cells(Me.ColItem_PrimaryTransaksiAsal.Index).Value = .primary_transaksi_asal
                            If .tipe_item_detail = EnumTipeItemDetail.BeginingBalanceItem Then
                                rowDGV.Cells(Me.ColItem_TipeItem.Index).Value = "Begining Balance"
                            ElseIf .tipe_item_detail = EnumTipeItemDetail.InItem Then
                                rowDGV.Cells(Me.ColItem_TipeItem.Index).Value = "In"
                            ElseIf .tipe_item_detail = EnumTipeItemDetail.OutItem Then
                                rowDGV.Cells(Me.ColItem_TipeItem.Index).Value = "Out"
                            ElseIf .tipe_item_detail = EnumTipeItemDetail.EndingBalanceItem Then
                                rowDGV.Cells(Me.ColItem_TipeItem.Index).Value = "Ending"
                            End If
                            rowDGV.Cells(Me.ColItem_TipeItem2.Index).Value = ""
                            rowDGV.Cells(Me.ColItem_TipeDoc2.Index).Value = ""
                            rowDGV.Cells(Me.ColItem_DocNo2.Index).Value = ""
                        End With
                    Next
                    While TBLItemd.DefaultView.Count > 0
                        TBLItemd.Rows.Remove(TBLItemd.DefaultView(0).Row)
                    End While
                    While TBLItem.DefaultView.Count > 0
                        TBLItem.Rows.Remove(TBLItem.DefaultView(0).Row)
                    End While
                End If
            Next


            'Me.DTAAOP.Fill(Me.TBLAOP, Me.EndDate, #1/1/2015#)
            'Me.DTABundling.Fill(TblBundling, Me.EndDate, #1/1/2015#)
            'Me.DTAInv.Fill(Me.TBLInv, Me.EndDate, #1/1/2015#)
            'Me.DTAStockInOutAdjust.Fill(Me.TBLStockInOutAdjust, Me.EndDate, #1/1/2015#)
            'Me.IsikanTipeDoc2Dkk()
            Me.IsikanDGVDetailWarehouse()
            Me.IsikanDGVDetailUsing()
            'xit Sub

            'Me.DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.ColItem_Tgl.DefaultCellStyle.Format = "dd MMM yyyy"
            'Me.DataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            TBLProdDummy.Rows.Clear()
            TBLProd.Rows.Clear()
            TBLItem.Rows.Clear()
            TBLItemd.Rows.Clear()
            TBLC.Rows.Clear()

            EndTime = Now
            MsgBox("Calculation finished in " + Me.ObjCD.DigitGrouping((EndTime - StartTime).TotalMinutes, 2) + " minutes")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

            DTAProd.Dispose()
            TBLProdDummy.Dispose()
            TBLProd.Dispose()
            DTAItem.Dispose()
            TBLItem.Dispose()
            DTAItemd.Dispose()
            TBLItemd.Dispose()
            DTAC.Dispose()
            TBLC.Dispose()
        End Try
    End Sub

    Private Sub Recalculate(ByVal vPrimaryClosing As Double, ByVal vListPrimaryCodeNo As List(Of Double))
        Dim DTAProd As New Dse3TableAdapters.tblmaster_productTableAdapter
        Dim TBLProdDummy As New Dse3.tblmaster_productDataTable
        Dim rowProd, rowProdDummy As Dse3.tblmaster_productRow
        Dim TBLProd As New Dse3.tblmaster_productDataTable
        Dim DTAItem As New Dse3TableAdapters.tbltrans_purchase_closing_itemTableAdapter
        Dim TBLItem As New Dse3.tbltrans_purchase_closing_itemDataTable
        Dim rowItem As Dse3.tbltrans_purchase_closing_itemRow
        Dim DTAItemd As New Dse3TableAdapters.tbltrans_purchase_closing_item_detailTableAdapter
        Dim TBLItemd As New Dse3.tbltrans_purchase_closing_item_detailDataTable
        Dim rowItemd As Dse3.tbltrans_purchase_closing_item_detailRow
        Dim rowDGV As DataGridViewRow
        Dim TipeItemd As EnumTipeItemDetail = EnumTipeItemDetail.BeginingBalanceItem
        Dim TipeTransaksiAsald As EnumTipeTranskasiAsal = EnumTipeTranskasiAsal.InAdjust
        Dim DTAC As New Dse3TableAdapters.tblmaster_categoryTableAdapter
        Dim TBLC As New Dse3.tblmaster_categoryDataTable
        Dim rowC As Dse3.tblmaster_categoryRow
        Dim hg As Double = 0

        Try
            '4 parameter dibawah harus di ganti kalo mau update data
            PrimaryTransaksiAcuan = vPrimaryClosing
            Me.StartDate = Me.dtpPeriod.Value
            Me.EndDate = Me.ObjCD.GetLasDateOfTheMonth(Me.StartDate)
            PeriodClosing = Me.StartDate


            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.vListPrimaryCategory.Clear()
            DTAC.Fill(TBLC)
            For Each rowC In TBLC.Rows
                Me.vListPrimaryCategory.Add(rowC.Primary)
            Next

            TBLProd.Rows.Clear()
            For Each PrimaryCategory As Double In Me.vListPrimaryCategory
                TBLProdDummy.Rows.Clear()
                DTAProd.Fill(TBLProdDummy, PrimaryCategory)
                For Each rowProdDummy In TBLProdDummy.Rows
                    rowProd = TBLProd.Newtblmaster_productRow
                    For ij As Integer = 0 To TBLProd.Columns.Count - 1
                        rowProd.Item(ij) = rowProdDummy.Item(ij)
                    Next
                    TBLProd.Rows.Add(rowProd)
                Next
            Next

            'load data tabel acuan
            DTAItem.Fill(TBLItem, PrimaryTransaksiAcuan)
            DTAItemd.Fill(TBLItemd, PrimaryTransaksiAcuan)
            Me.DTAAOP.Fill(Me.TBLAOP, Me.EndDate, #1/1/2015#)
            Me.DTABundling.Fill(TblBundling, Me.EndDate, #1/1/2015#)
            Me.DTAInv.Fill(Me.TBLInv, Me.EndDate, #1/1/2015#)
            Me.DTAStockInOutAdjust.Fill(Me.TBLStockInOutAdjust, Me.EndDate, #1/1/2015#)

            Me.DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            'Me.DataGridView1.Rows.Clear()
            'Me.DataGridView2.Rows.Clear()

            Me.TotalRow = TBLProd.Rows.Count

            For Each rowProd In TBLProd.Rows
                hg += 1
                Me.CurrentRow += 1

                If vListPrimaryCodeNo.Contains(rowProd.Primary) = True Then
                    'load dgv item 
                    TBLItem.DefaultView.RowFilter = "[Primary Product] = " + rowProd.Primary.ToString
                    If TBLItem.DefaultView.Count = 0 Then
                        'tambahkan baris baru kosong
                        'Me.DataGridView1.Rows.Add()
                        'rowDGV = Me.DataGridView1.Rows(Me.DataGridView1.Rows.Count - 1)
                        'With rowProd
                        '    rowDGV.Cells(Me.ColAmountBB.Index).Value = Me.ObjCD.DigitGrouping(0, DefPrec)
                        '    rowDGV.Cells(Me.ColAmountEnd.Index).Value = Me.ObjCD.DigitGrouping(0, DefPrec)
                        '    rowDGV.Cells(Me.ColAmountIn.Index).Value = Me.ObjCD.DigitGrouping(0, DefPrec)
                        '    rowDGV.Cells(Me.ColAmountOut.Index).Value = Me.ObjCD.DigitGrouping(0, DefPrec)
                        '    rowDGV.Cells(Me.ColCategory.Index).Value = rowProd.Category_Code
                        '    rowDGV.Cells(Me.ColCodeNo.Index).Value = rowProd.Product_Code
                        '    rowDGV.Cells(Me.ColOutQty.Index).Value = Me.ObjCD.DigitGrouping(0, ._Decimal)
                        '    rowDGV.Cells(Me.ColQtyBB.Index).Value = Me.ObjCD.DigitGrouping(0, ._Decimal)
                        '    rowDGV.Cells(Me.ColQtyEnd.Index).Value = Me.ObjCD.DigitGrouping(0, ._Decimal)
                        '    rowDGV.Cells(Me.ColQtyIn.Index).Value = Me.ObjCD.DigitGrouping(0, ._Decimal)
                        '    rowDGV.Cells(Me.ColDesc.Index).Value = rowProd.Description
                        '    rowDGV.Cells(Me.ColUnit.Index).Value = rowProd.Unit_Code
                        'End With
                    Else
                        'tambahkan baris baru
                        rowItem = TBLItem.DefaultView(0).Row
                        'If rowItem.Begin_Qty > 0 Or rowItem.In_Qty > 0 Or rowItem.Out_Qty > 0 Or rowItem.Ending_Qty > 0 Then
                        Me.DataGridView1.Rows.Add()
                        rowDGV = Me.DataGridView1.Rows(Me.DataGridView1.Rows.Count - 1)
                        With rowItem
                            rowDGV.Cells(Me.ColAmountBB.Index).Value = Me.ObjCD.DigitGrouping(.Begin_Amount, DefPrec)
                            rowDGV.Cells(Me.ColAmountEnd.Index).Value = Me.ObjCD.DigitGrouping(.Ending_Amount, DefPrec)
                            rowDGV.Cells(Me.ColAmountIn.Index).Value = Me.ObjCD.DigitGrouping(.In_Amount, DefPrec)
                            rowDGV.Cells(Me.ColAmountOut.Index).Value = Me.ObjCD.DigitGrouping(.Out_Amount, DefPrec)
                            rowDGV.Cells(Me.ColCategory.Index).Value = rowProd.Category_Code
                            rowDGV.Cells(Me.ColCodeNo.Index).Value = rowProd.Product_Code
                            rowDGV.Cells(Me.ColOutQty.Index).Value = Me.ObjCD.DigitGrouping(.Out_Qty, .Decimal_Product)
                            rowDGV.Cells(Me.ColQtyBB.Index).Value = Me.ObjCD.DigitGrouping(.Begin_Qty, .Decimal_Product)
                            rowDGV.Cells(Me.ColQtyEnd.Index).Value = Me.ObjCD.DigitGrouping(.Ending_Qty, .Decimal_Product)
                            rowDGV.Cells(Me.ColQtyIn.Index).Value = Me.ObjCD.DigitGrouping(.In_Qty, .Decimal_Product)
                            rowDGV.Cells(Me.ColDesc.Index).Value = rowProd.Description
                            rowDGV.Cells(Me.ColUnit.Index).Value = rowProd.Unit_Code
                            rowDGV.Cells(Me.ColDefPrec.Index).Value = rowProd._Decimal
                            rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value = rowProd.Primary
                            'rowDGV.Cells(Me.ColDecimalProd.Index).Value = .Decimal_Product
                        End With
                        'tambahkan baris detail
                        TBLItemd.DefaultView.RowFilter = "[primary code no] = " + rowProd.Primary.ToString
                        TBLItemd.DefaultView.Sort = "primary"
                        For Each DV As DataRowView In TBLItemd.DefaultView
                            rowItemd = DV.Row
                            Me.DataGridView2.Rows.Add()
                            rowDGV = Me.DataGridView2.Rows(Me.DataGridView2.Rows.Count - 1)
                            With rowItemd
                                rowDGV.Cells(Me.ColItem_No.Index).Value = rowDGV.Index + 1
                                rowDGV.Cells(Me.ColItem_AmountBB.Index).Value = Me.ObjCD.DigitGrouping(0, DefPrec)
                                rowDGV.Cells(Me.ColItem_AmountEnd.Index).Value = Me.ObjCD.DigitGrouping(0, DefPrec)
                                rowDGV.Cells(Me.ColItem_AmountIn.Index).Value = Me.ObjCD.DigitGrouping(0, DefPrec)
                                rowDGV.Cells(Me.ColItem_AmountOut.Index).Value = Me.ObjCD.DigitGrouping(0, DefPrec)
                                rowDGV.Cells(Me.ColItem_QtyBB.Index).Value = Me.ObjCD.DigitGrouping(0, rowProd._Decimal)
                                rowDGV.Cells(Me.ColItem_QtyEnd.Index).Value = Me.ObjCD.DigitGrouping(0, rowProd._Decimal)
                                rowDGV.Cells(Me.ColItem_QtyIn.Index).Value = Me.ObjCD.DigitGrouping(0, rowProd._Decimal)
                                rowDGV.Cells(Me.ColItem_QtyOut.Index).Value = Me.ObjCD.DigitGrouping(0, rowProd._Decimal)
                                rowDGV.Cells(Me.ColItem_TipeDoc.Index).Value = Me.GetTipeDocument_TransaksiAsal(rowItemd.tipe_transaksi_asal)
                                If .tipe_item_detail = CInt(EnumTipeItemDetail.BeginingBalanceItem) Then
                                    rowDGV.Cells(Me.ColItem_AmountBB.Index).Value = Me.ObjCD.DigitGrouping(.amount, DefPrec)
                                    rowDGV.Cells(Me.ColItem_QtyBB.Index).Value = Me.ObjCD.DigitGrouping(.qty, rowProd._Decimal)
                                ElseIf .tipe_item_detail = CInt(EnumTipeItemDetail.InItem) Then
                                    rowDGV.Cells(Me.ColItem_AmountIn.Index).Value = Me.ObjCD.DigitGrouping(.amount, DefPrec)
                                    rowDGV.Cells(Me.ColItem_QtyIn.Index).Value = Me.ObjCD.DigitGrouping(.qty, rowProd._Decimal)
                                ElseIf .tipe_item_detail = CInt(EnumTipeItemDetail.OutItem) Then
                                    rowDGV.Cells(Me.ColItem_AmountOut.Index).Value = Me.ObjCD.DigitGrouping(.amount, DefPrec)
                                    rowDGV.Cells(Me.ColItem_QtyOut.Index).Value = Me.ObjCD.DigitGrouping(.qty, rowProd._Decimal)
                                ElseIf .tipe_item_detail = CInt(EnumTipeItemDetail.EndingBalanceItem) Then
                                    rowDGV.Cells(Me.ColItem_AmountEnd.Index).Value = Me.ObjCD.DigitGrouping(.amount, DefPrec)
                                    rowDGV.Cells(Me.ColItem_QtyEnd.Index).Value = Me.ObjCD.DigitGrouping(.qty, rowProd._Decimal)
                                End If

                                rowDGV.Cells(Me.ColItem_Category.Index).Value = rowProd.Category_Code
                                rowDGV.Cells(Me.ColItem_CodeNo.Index).Value = rowProd.Product_Code
                                rowDGV.Cells(Me.ColItem_Desc.Index).Value = rowProd.Description
                                rowDGV.Cells(Me.ColItem_NoDoc.Index).Value = Me.GetDocumentNo(.primary_transaksi_asal, CInt(.tipe_transaksi_asal))
                                rowDGV.Cells(Me.ColItem_Tgl.Index).Value = .tgl_transaksi_asal
                                rowDGV.Cells(Me.ColItem_Unit.Index).Value = rowProd.Unit_Code
                                rowDGV.Cells(Me.ColItem_DefPrec.Index).Value = rowProd._Decimal
                                rowDGV.Cells(Me.ColItem_PrimaryTransaksiAsal.Index).Value = .primary_transaksi_asal
                                If .tipe_item_detail = EnumTipeItemDetail.BeginingBalanceItem Then
                                    rowDGV.Cells(Me.ColItem_TipeItem.Index).Value = "Begining Balance"
                                ElseIf .tipe_item_detail = EnumTipeItemDetail.InItem Then
                                    rowDGV.Cells(Me.ColItem_TipeItem.Index).Value = "In"
                                ElseIf .tipe_item_detail = EnumTipeItemDetail.OutItem Then
                                    rowDGV.Cells(Me.ColItem_TipeItem.Index).Value = "Out"
                                ElseIf .tipe_item_detail = EnumTipeItemDetail.EndingBalanceItem Then
                                    rowDGV.Cells(Me.ColItem_TipeItem.Index).Value = "Ending"
                                End If
                            End With
                        Next
                        'End If
                        'If rowProd.Primary = vPrimaryCodeNo Then
                        '    Exit For
                        'End If
                    End If
                End If
            Next

            'Me.IsikanTipeDoc2Dkk()
            Me.DataGridView3.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For Each PrimaryCodeNo As Double In vListPrimaryCodeNo
                Me.IsikanDGVDetailWarehouse(PrimaryCodeNo)
            Next
            'Me.DataGridView3.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            'Me.IsikanDGVDetailUsing()
            'Exit Sub

            ' Me.DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.ColItem_Tgl.DefaultCellStyle.Format = "dd MMM yyyy"
            'Me.DataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            TBLProdDummy.Rows.Clear()
            TBLProd.Rows.Clear()
            TBLItem.Rows.Clear()
            TBLItemd.Rows.Clear()
            TBLC.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

            DTAProd.Dispose()
            TBLProdDummy.Dispose()
            TBLProd.Dispose()
            DTAItem.Dispose()
            TBLItem.Dispose()
            DTAItemd.Dispose()
            TBLItemd.Dispose()
            DTAC.Dispose()
            TBLC.Dispose()
        End Try
    End Sub

    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.dtpPeriod.Value = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value)
            Me.LblProgress.Text = ""

            If Me.UserCode <> "USER" Then
                Me.btnCek.Visible = False
                Me.btnCek.TabStop = False
                Me.btnShowSum.Visible = False
                Me.btnShowSum.TabStop = False
                Me.Button1.Visible = False
                Me.Button1.TabStop = False
                Me.Button2.Visible = False
                Me.Button2.TabStop = False
                Me.btnCekItemVsDetail.Visible = False
                Me.btnCekItemVsDetail.TabStop = False
                Me.btnRepairCategory.Visible = False
                Me.btnRepairCategory.TabStop = False
            Else
                Me.DataGridView1.Height = Me.DataGridView1.Height - Me.btnCek.Height
                Me.DataGridView1.Top = Me.DataGridView1.Top + Me.btnCek.Height + 5
                'Me.DgvItem.Height = Me.DataGridView1.Height
                'Me.DgvItem.Top = Me.DataGridView1.Top
            End If

            Me.LblProgress.Text = ""
            Me.LblProgress2.Text = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetDocumentNo(ByVal vPrimaryTransksiAsal As Double, ByVal vTipeTransaksiAsal As EnumTipeTranskasiAsal) As String
        Try
            If vTipeTransaksiAsal = EnumTipeTranskasiAsal.InAdjust Or vTipeTransaksiAsal = EnumTipeTranskasiAsal.OutAdjust Then
                Me.TBLStockInOutAdjust.DefaultView.RowFilter = "Primary = " + vPrimaryTransksiAsal.ToString
                If Me.TBLStockInOutAdjust.DefaultView.Count > 0 Then
                    Me.rowStockInOut = Me.TBLStockInOutAdjust.DefaultView(0).Row
                    Return Me.rowStockInOut.No_Transaksi
                Else
                    Return ""
                End If
            ElseIf vTipeTransaksiAsal = EnumTipeTranskasiAsal.InBundling Or vTipeTransaksiAsal = EnumTipeTranskasiAsal.OutBundling Then
                Me.TblBundling.DefaultView.RowFilter = "Primary = " + vPrimaryTransksiAsal.ToString
                If Me.TblBundling.DefaultView.Count > 0 Then
                    Me.rowBundling = Me.TblBundling.DefaultView(0).Row
                    Return Me.rowBundling.No_Transaksi
                Else
                    Return ""
                End If
            ElseIf vTipeTransaksiAsal = EnumTipeTranskasiAsal.InPurchase Then
                Me.TBLAOP.DefaultView.RowFilter = "Primary = " + vPrimaryTransksiAsal.ToString
                If Me.TBLAOP.DefaultView.Count > 0 Then
                    Me.rowAOP = Me.TBLAOP.DefaultView(0).Row
                    Return Me.rowAOP.No_Transaksi
                Else
                    Return ""
                End If
            ElseIf vTipeTransaksiAsal = EnumTipeTranskasiAsal.InSR Then
                Return ""
            ElseIf vTipeTransaksiAsal = EnumTipeTranskasiAsal.InStockTakeAdjust Then
                Return ""
            ElseIf vTipeTransaksiAsal = EnumTipeTranskasiAsal.OutPR Then
                Return ""
            ElseIf vTipeTransaksiAsal = EnumTipeTranskasiAsal.OutSales Then
                Me.TBLInv.DefaultView.RowFilter = "Primary = " + vPrimaryTransksiAsal.ToString
                If Me.TBLInv.DefaultView.Count > 0 Then
                    Me.rowInv = Me.TBLInv.DefaultView(0).Row
                    Return Me.rowInv.No_Transaksi
                Else
                    Return ""
                End If
            ElseIf vTipeTransaksiAsal = EnumTipeTranskasiAsal.OutStockTakeAdjust Then
                Return ""
            End If
            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function GetTipeDocument_TransaksiAsal(ByVal vTipeTransaksiAsal As EnumTipeTranskasiAsal) As String
        Try
            If vTipeTransaksiAsal = EnumTipeTranskasiAsal.InAdjust Then
                Return "In Adjustment"
            ElseIf vTipeTransaksiAsal = EnumTipeTranskasiAsal.InBundling Then
                Return "In Bundling"
            ElseIf vTipeTransaksiAsal = EnumTipeTranskasiAsal.InPurchase Then
                Return "In Purchase"
            ElseIf vTipeTransaksiAsal = EnumTipeTranskasiAsal.InSR Then
                Return "In Sales Return"
            ElseIf vTipeTransaksiAsal = EnumTipeTranskasiAsal.InStockTakeAdjust Then
                Return "In Stock Take"
            ElseIf vTipeTransaksiAsal = EnumTipeTranskasiAsal.OutAdjust Then
                Return "Out Adjustment"
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

            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Sub IsikanDGVDetailUsing()
        Dim rowDGV, rowDGV1 As DataGridViewRow
        Dim KeteranganUsing As String = ""
        Dim PrimaryBundling As Double = -1
        Dim DTA_After As New Dse2TableAdapters.tbltransaksi_stockbundling_item_afterTableAdapter
        Dim TBL_After As New Dse2.tbltransaksi_stockbundling_item_afterDataTable
        Dim row_After As Dse2.tbltransaksi_stockbundling_item_afterRow
        Dim TBL_D As New Dse2.DataTable5DataTable
        Dim row_D As Dse2.DataTable5Row
        Dim defPrec As Integer = Me.ObjCur.GetDecimal_DefaultCurrency
        Dim TipeAdj As String = ""
        Dim ad As Double = 0

        Try
            Me.DataGridView4.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DataGridView4.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.DataGridView4.Rows.Clear()

           
            Me.ProgressBar1.Value = 0
            Me.ProgressBar1.Maximum = Me.DataGridView2.Rows.Count

            For Each rowDGV In Me.DataGridView2.Rows
                Me.ProgressBar1.Value += 1

                If Me.ProgressBar1.Value > 25 Then
                    If Math.IEEERemainder(Me.ProgressBar1.Value, 100) = 0 Then
                        Me.LblProgress.Text = "Phase 4 of 5 - Calculating detail using data : Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString
                    ElseIf Me.ProgressBar1.Value = Me.ProgressBar1.Maximum Then
                        Me.LblProgress.Text = "Phase 4 of 5 - Calculating detail using data : Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString
                    End If
                Else
                    Me.LblProgress.Text = "Phase 4 of 5 - Calculating detail using data : Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString
                End If

                ad += 1

                KeteranganUsing = ""

                If rowDGV.Cells(Me.ColItem_TipeDoc.Index).Value = "Out Bundling" Then
                    'keterangan using, ambil item pertama dari bundling tersebut, lalu tuliskan kategory-nya
                    PrimaryBundling = rowDGV.Cells(Me.ColItem_PrimaryTransaksiAsal.Index).Value
                    DTA_After.Fill(TBL_After, PrimaryBundling)
                    row_After = TBL_After.Rows(0)
                    KeteranganUsing = Me.ObjCat.GetCategoryCode(Me.ObjProd.GetPrimaryCategory(row_After.Primary_Product))
                ElseIf rowDGV.Cells(Me.ColItem_TipeDoc.Index).Value = "Out Purchase Return" Then
                    'keterangan using: purchase return
                    KeteranganUsing = "purchase return"
                ElseIf rowDGV.Cells(Me.ColItem_TipeDoc.Index).Value = "Out Sales" Then
                    'keterangan using: out sales
                    KeteranganUsing = "sales"
                ElseIf rowDGV.Cells(Me.ColItem_TipeDoc.Index).Value = "Out Stock Take" Then
                    'keterangan using: out stock take
                    KeteranganUsing = "stock take"
                ElseIf rowDGV.Cells(Me.ColItem_TipeDoc.Index).Value = "Overhead Cost" Then
                    KeteranganUsing = "overhead cost"
                ElseIf rowDGV.Cells(Me.ColItem_TipeDoc.Index).Value = "Out Adjustment" Then
                    'keterangan using, ambil tipe adjustmentnya
                    TipeAdj = Me.ObjAdjust.GetCodeSettingByPrimary(Me.ObjAdjust.GetPrimaryTipeAdjustment(rowDGV.Cells(Me.ColItem_PrimaryTransaksiAsal.Index).Value.ToString).ToString)
                    If TipeAdj = "" Then
                        TipeAdj = "No Type"
                    End If
                    KeteranganUsing = "adjustment type: " + TipeAdj
                    KeteranganUsing = KeteranganUsing + " - " + Me.ObjAdjust.GetCreatedBy(rowDGV.Cells(Me.ColItem_PrimaryTransaksiAsal.Index).Value)
                Else
                    KeteranganUsing = ""
                End If

                If KeteranganUsing <> "" Then
                    TBL_D.DefaultView.RowFilter = "[Code No] = '" + rowDGV.Cells(Me.ColItem_CodeNo.Index).Value + "' AND [Keterangan Using] = '" + KeteranganUsing + "'"
                    If TBL_D.DefaultView.Count = 0 Then
                        'tambahkan baris baru
                        row_D = TBL_D.NewDataTable5Row
                        With row_D
                            .Category_Code = rowDGV.Cells(Me.ColItem_Category.Index).Value
                            .Code_No = rowDGV.Cells(Me.ColItem_CodeNo.Index).Value
                            .DefPrec = rowDGV.Cells(Me.ColItem_DefPrec.Index).Value
                            .Desc = rowDGV.Cells(Me.ColItem_Desc.Index).Value
                            .Keterangan_Using = KeteranganUsing
                            .Out_Amount = rowDGV.Cells(Me.ColItem_AmountOut.Index).Value
                            .Qty_Out = rowDGV.Cells(Me.ColItem_QtyOut.Index).Value
                            .Unit = rowDGV.Cells(Me.ColItem_Unit.Index).Value
                        End With
                        TBL_D.Rows.Add(row_D)
                    Else
                        'tambahkan amount dan qty using
                        row_D = TBL_D.DefaultView(0).Row
                        With row_D
                            .Qty_Out += rowDGV.Cells(Me.ColItem_QtyOut.Index).Value
                            .Out_Amount += rowDGV.Cells(Me.ColItem_AmountOut.Index).Value
                        End With
                    End If
                End If
            Next

            ad = 0
            Me.ProgressBar1.Value = 0
            Me.ProgressBar1.Maximum = TBL_D.Rows.Count

            For Each row_D In TBL_D.Rows
                Me.ProgressBar1.Value += 1

                If Me.ProgressBar1.Value > 25 Then
                    If Math.IEEERemainder(Me.ProgressBar1.Value, 100) = 0 Then
                        Me.LblProgress.Text = "Phase 5 of 5 - Populating using datagridview : Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString
                    ElseIf Me.ProgressBar1.Value = Me.ProgressBar1.Maximum Then
                        Me.LblProgress.Text = "Phase 5 of 5 - Populating using datagridview : Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString
                    End If
                Else
                    Me.LblProgress.Text = "Phase 5 of 5 - Populating using datagridview : Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString
                End If

                ad += 1
                Me.DataGridView4.Rows.Add()
                rowDGV1 = Me.DataGridView4.Rows(Me.DataGridView4.Rows.Count - 1)
                With rowDGV1
                    .Cells(Me.ColU_AmountEnd.Index).Value = Me.ObjCD.DigitGrouping(row_D.Out_Amount, defPrec)
                    .Cells(Me.ColU_Category.Index).Value = row_D.Category_Code
                    .Cells(Me.ColU_CodeNo.Index).Value = row_D.Code_No
                    .Cells(Me.ColU_Description.Index).Value = row_D.Desc
                    .Cells(Me.ColU_QtyUsing.Index).Value = Me.ObjCD.DigitGrouping(row_D.Qty_Out, row_D.DefPrec)
                    .Cells(Me.ColU_Unit.Index).Value = row_D.Unit
                    .Cells(Me.ColU_UsingDetail.Index).Value = row_D.Keterangan_Using
                End With
            Next
            'Me.DataGridView4.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_After.Dispose()
            TBL_After.Dispose()
            TBL_D.Dispose()
        End Try
    End Sub

    Private Sub AddNewRow(ByVal vrowAcuan As Inventory_Module.DseLaporanStockDetail2.DataTable1Row, _
    ByVal vTBLDetail As Inventory_Module.DseLaporanStockDetail2.DataTable1DataTable)
        Dim rowDummy As Inventory_Module.DseLaporanStockDetail2.DataTable1Row

        Try
            rowDummy = vTBLDetail.NewDataTable1Row
            For ij As Integer = 0 To vTBLDetail.Columns.Count - 1
                rowDummy.Item(ij) = vrowAcuan.Item(ij)
            Next
            vTBLDetail.Rows.Add(rowDummy)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CompactData(ByVal vTBLDetail As Inventory_Module.DseLaporanStockDetail2.DataTable1DataTable)
        Dim rowDetail, rowDetail_B4, rowDummy As Inventory_Module.DseLaporanStockDetail2.DataTable1Row
        Dim TBLDummy As New Inventory_Module.DseLaporanStockDetail2.DataTable1DataTable
        Dim CurIdx As Double = 0

        Try
            'urutkan ulang tabel detail
            vTBLDetail.DefaultView.RowFilter = ""
            vTBLDetail.DefaultView.Sort = "[Primary Product], [Warehouse Code]"
            'salin ke tabel dummy, hanya saja pastikan baris yang terduplikasi dijadikan 1 baris
            For Each DV As DataRowView In vTBLDetail.DefaultView
                rowDetail = DV.Row
                If CurIdx > 0 Then
                    rowDetail_B4 = vTBLDetail.DefaultView(CurIdx - 1).Row
                    If rowDetail.Primary_Product = rowDetail_B4.Primary_Product And _
                    rowDetail.Warehouse_Code = rowDetail_B4.Warehouse_Code Then
                        rowDummy = TBLDummy.Rows(TBLDummy.Rows.Count - 1)
                        rowDummy.Qty += rowDetail.Qty
                    Else
                        Me.AddNewRow(rowDetail, TBLDummy)
                    End If
                Else
                    'langsung tambahkan baris baru
                    Me.AddNewRow(rowDetail, TBLDummy)
                End If

                CurIdx += 1
            Next

            'salin balik
            vTBLDetail.Rows.Clear()
            For Each rowDummy In TBLDummy.Rows
                rowDetail = vTBLDetail.NewDataTable1Row
                For ij As Integer = 0 To vTBLDetail.Columns.Count - 1
                    rowDetail.Item(ij) = rowDummy.Item(ij)
                Next
                vTBLDetail.Rows.Add(rowDetail)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanDGVDetailWarehouse(ByVal vPrimaryCodeNo As Double)
        Dim TBLDetail As New Inventory_Module.DseLaporanStockDetail2.DataTable1DataTable
        Dim rowDetail As Inventory_Module.DseLaporanStockDetail2.DataTable1Row
        Dim ListPrimaryCodeNo As New List(Of Double)
        Dim Qty, Amount, QtyAllocated, AmountAllocated, Amount1, Qty1 As Decimal
        Dim Jml As Integer = 0
        Dim Idx As Integer = 0
        Dim DefPrec As Integer = Me.ObjCur.GetDecimal_DefaultCurrency
        Dim RowDGV As DataGridViewRow
        Dim CategoryCode As String = ""
        Dim CodeNo, Desc, UnitCode As String
        Dim TambahBaris As Boolean = True

        Try
            'Me.DataGridView3.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            'Me.DataGridView3.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            'Me.DataGridView3.Rows.Clear()

            Me.LoadStockDetail(TBLDetail, vPrimaryCodeNo)
            Me.CompactData(TBLDetail)
            ListPrimaryCodeNo.Add(vPrimaryCodeNo)

            For Each PrimaryCodeNo1 As Double In ListPrimaryCodeNo
                TBLDetail.DefaultView.RowFilter = "[Primary Product] = " + PrimaryCodeNo1.ToString
                Me.GetQtyAmountRekap(PrimaryCodeNo1, Qty, Amount, CategoryCode, CodeNo, Desc, UnitCode)
                Jml = TBLDetail.DefaultView.Count
                Idx = 0
                AmountAllocated = 0
                QtyAllocated = 0

                For Each DV As DataRowView In TBLDetail.DefaultView
                    rowDetail = DV.Row
                    Qty1 = rowDetail.Qty
                    QtyAllocated += Qty1

                    If Idx = Jml - 1 Then
                        Amount1 = Amount - AmountAllocated
                        AmountAllocated = Amount
                    Else
                        Amount1 = Math.Round((Qty1 * Amount) / Qty, DefPrec, MidpointRounding.AwayFromZero)
                        AmountAllocated += Amount1
                    End If

                    'tambahkan baris
                    Me.DataGridView3.Rows.Add()
                    RowDGV = Me.DataGridView3.Rows(Me.DataGridView3.Rows.Count - 1)
                    RowDGV.Cells(Me.ColW_AmountEnd.Index).Value = Me.ObjCD.DigitGrouping(Amount1, DefPrec)
                    RowDGV.Cells(Me.ColW_CategoryCode.Index).Value = rowDetail.Categorey_Code
                    RowDGV.Cells(Me.ColW_CodeNo.Index).Value = rowDetail.Product_Code
                    RowDGV.Cells(Me.ColW_Desc.Index).Value = rowDetail.Description
                    RowDGV.Cells(Me.ColW_QtyEnd.Index).Value = Me.ObjCD.DigitGrouping(Qty1, rowDetail._Decimal)
                    RowDGV.Cells(Me.ColW_UnitCode.Index).Value = rowDetail.Unit_Code
                    RowDGV.Cells(Me.ColW_Warehouse.Index).Value = rowDetail.Warehouse_Code + " - " + rowDetail.Warehouse_Desc

                    Idx += 1
                Next

                If QtyAllocated <> Qty Or AmountAllocated <> Amount Then
                    If IsNothing(rowDetail) = True Then
                        TambahBaris = True
                    Else
                        If rowDetail.Product_Code = CodeNo Then
                            TambahBaris = False
                        Else
                            TambahBaris = True
                        End If
                    End If
                    If TambahBaris = True Then
                        'tambahkan baris
                        Me.DataGridView3.Rows.Add()
                        RowDGV = Me.DataGridView3.Rows(Me.DataGridView3.Rows.Count - 1)
                        RowDGV.Cells(Me.ColW_AmountEnd.Index).Value = Me.ObjCD.DigitGrouping(Amount - AmountAllocated, DefPrec)
                        RowDGV.Cells(Me.ColW_CategoryCode.Index).Value = CategoryCode
                        RowDGV.Cells(Me.ColW_CodeNo.Index).Value = CodeNo
                        RowDGV.Cells(Me.ColW_Desc.Index).Value = Desc
                        RowDGV.Cells(Me.ColW_QtyEnd.Index).Value = Me.ObjCD.DigitGrouping(Qty - QtyAllocated, Me.ObjProd.GetAngkaDecimal(PrimaryCodeNo1))
                        RowDGV.Cells(Me.ColW_UnitCode.Index).Value = UnitCode
                        RowDGV.Cells(Me.ColW_Warehouse.Index).Value = ""
                    Else
                        RowDGV = Me.DataGridView3.Rows(Me.DataGridView3.Rows.Count - 1)
                        RowDGV.Cells(Me.ColW_QtyEnd.Index).Value = Me.ObjCD.DigitGrouping(CDec(RowDGV.Cells(Me.ColW_QtyEnd.Index).Value) + _
                        Qty - QtyAllocated, Me.ObjProd.GetAngkaDecimal(PrimaryCodeNo1))
                        RowDGV.Cells(Me.ColW_AmountEnd.Index).Value = Me.ObjCD.DigitGrouping(CDec(RowDGV.Cells(Me.ColW_AmountEnd.Index).Value) + _
                        Amount - AmountAllocated, DefPrec)
                    End If
                End If
            Next
            'Me.DataGridView3.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanDGVDetailWarehouse()
        Dim rowDetail, rowDetail1, rowDetail2 As Inventory_Module.DseLaporanStockDetail2.DataTable1Row
        Dim ListPrimaryCodeNo As New List(Of Double)
        Dim Qty, Amount, QtyAllocated, AmountAllocated, Amount1, Qty1 As Decimal
        Dim Jml As Integer = 0
        Dim Idx As Integer = 0
        Dim DefPrec As Integer = Me.ObjCur.GetDecimal_DefaultCurrency
        Dim RowDGV As DataGridViewRow
        Dim CategoryCode As String = ""
        Dim CodeNo, Desc, UnitCode As String
        Dim TambahBaris As Boolean = False
        Dim PerluRepair As Boolean = False
        Dim ij As Double = 0

        Try

            Me.DataGridView3.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DataGridView3.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.DataGridView3.Rows.Clear()
            'Me.LoadStockDetail(TBLDetail)
            'Me.SudahLoadDGV = False
            'Me.CetakanStockDetail1()
            'While Me.SudahLoadDGV = False
            '    'tunggu 1 menit
            '    Threading.Thread.Sleep(60000)
            'End While

            Me.CompactData(TBLDetail)
            For Each RowDGV In Me.DataGridView1.Rows
                ListPrimaryCodeNo.Add(RowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value)
            Next

            Me.ProgressBar1.Value = 0
            Me.ProgressBar1.Maximum = ListPrimaryCodeNo.Count

            TBLDetail.DefaultView.Sort = "Qty DESC"
            For Each PrimaryCodeNo1 As Double In ListPrimaryCodeNo
                Me.ProgressBar1.Value += 1

                'If Me.ProgressBar1.Value = 5038 Then
                '    MsgBox("")
                'End If
                If Me.ProgressBar1.Value > 25 Then
                    If Math.IEEERemainder(Me.ProgressBar1.Value, 100) = 0 Then
                        Me.LblProgress.Text = "Phase 3 of 5 - Calculating detail warehouse data : Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString
                    ElseIf Me.ProgressBar1.Value = Me.ProgressBar1.Maximum Then
                        Me.LblProgress.Text = "Phase 3 of 5 - Calculating detail data : Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString
                    End If
                Else
                    Me.LblProgress.Text = "Phase 3 of 5 - Calculating detail warehouse data : Data " + Me.ProgressBar1.Value.ToString + " of " + Me.ProgressBar1.Maximum.ToString
                End If

                TBLDetail.DefaultView.RowFilter = "[Primary Product] = " + PrimaryCodeNo1.ToString
                PerluRepair = False
                If TBLDetail.DefaultView.Count > 0 Then
                    If TBLDetail.DefaultView.Count > 1 Then
                        rowDetail1 = TBLDetail.DefaultView(0).Row
                        rowDetail2 = TBLDetail.DefaultView(TBLDetail.DefaultView.Count - 1).Row
                        If rowDetail2.Qty < 0 And rowDetail1.Qty > 0 Then
                            PerluRepair = True
                        End If
                    End If
                    If PerluRepair = False Then
                        For Each DV As DataRowView In TBLDetail.DefaultView
                            rowDetail = DV.Row
                            If rowDetail.Qty = 0 Then
                                PerluRepair = True
                            End If
                        Next
                    End If

                    If PerluRepair = True Then
                        Me.RepairMinus(TBLDetail)
                    End If
                End If

                Me.GetQtyAmountRekap(PrimaryCodeNo1, Qty, Amount, CategoryCode, CodeNo, Desc, UnitCode)
                Jml = TBLDetail.DefaultView.Count
                Idx = 0
                AmountAllocated = 0
                QtyAllocated = 0

                For Each DV As DataRowView In TBLDetail.DefaultView
                    rowDetail = DV.Row
                    Qty1 = rowDetail.Qty
                    QtyAllocated += Qty1

                    If Idx = Jml - 1 Then
                        Amount1 = Amount - AmountAllocated
                        AmountAllocated = Amount
                    Else
                        If Qty = 0 Then
                            Amount1 = 0
                        Else
                            Amount1 = Math.Round((Qty1 * Amount) / Qty, DefPrec, MidpointRounding.AwayFromZero)
                        End If
                        AmountAllocated += Amount1
                    End If

                    'tambahkan baris
                    Me.DataGridView3.Rows.Add()
                    RowDGV = Me.DataGridView3.Rows(Me.DataGridView3.Rows.Count - 1)
                    RowDGV.Cells(Me.ColW_AmountEnd.Index).Value = Me.ObjCD.DigitGrouping(Amount1, DefPrec)
                    RowDGV.Cells(Me.ColW_CategoryCode.Index).Value = rowDetail.Categorey_Code
                    RowDGV.Cells(Me.ColW_CodeNo.Index).Value = rowDetail.Product_Code
                    RowDGV.Cells(Me.ColW_Desc.Index).Value = rowDetail.Description
                    RowDGV.Cells(Me.ColW_QtyEnd.Index).Value = Me.ObjCD.DigitGrouping(Qty1, rowDetail._Decimal)
                    RowDGV.Cells(Me.ColW_UnitCode.Index).Value = rowDetail.Unit_Code
                    RowDGV.Cells(Me.ColW_Warehouse.Index).Value = rowDetail.Warehouse_Code + " - " + rowDetail.Warehouse_Desc

                    Idx += 1
                Next

                If QtyAllocated <> Qty Or AmountAllocated <> Amount Then
                    If IsNothing(rowDetail) = True Then
                        TambahBaris = True
                    Else
                        If rowDetail.RowState = DataRowState.Detached Then
                            TambahBaris = True
                        Else
                            If rowDetail.Product_Code = CodeNo Then
                                TambahBaris = False
                            Else
                                TambahBaris = True
                            End If

                        End If
                    End If

                    If TambahBaris = True Then
                        'tambahkan baris
                        Me.DataGridView3.Rows.Add()
                        RowDGV = Me.DataGridView3.Rows(Me.DataGridView3.Rows.Count - 1)
                        RowDGV.Cells(Me.ColW_AmountEnd.Index).Value = Me.ObjCD.DigitGrouping(Amount - AmountAllocated, DefPrec)
                        RowDGV.Cells(Me.ColW_CategoryCode.Index).Value = CategoryCode
                        RowDGV.Cells(Me.ColW_CodeNo.Index).Value = CodeNo
                        RowDGV.Cells(Me.ColW_Desc.Index).Value = Desc
                        RowDGV.Cells(Me.ColW_QtyEnd.Index).Value = Me.ObjCD.DigitGrouping(Qty - QtyAllocated, Me.ObjProd.GetAngkaDecimal(PrimaryCodeNo1))
                        RowDGV.Cells(Me.ColW_UnitCode.Index).Value = UnitCode
                        RowDGV.Cells(Me.ColW_Warehouse.Index).Value = ""
                    Else
                        RowDGV = Me.DataGridView3.Rows(Me.DataGridView3.Rows.Count - 1)
                        RowDGV.Cells(Me.ColW_QtyEnd.Index).Value = Me.ObjCD.DigitGrouping(CDec(RowDGV.Cells(Me.ColW_QtyEnd.Index).Value) + _
                        Qty - QtyAllocated, Me.ObjProd.GetAngkaDecimal(PrimaryCodeNo1))
                        RowDGV.Cells(Me.ColW_AmountEnd.Index).Value = Me.ObjCD.DigitGrouping(CDec(RowDGV.Cells(Me.ColW_AmountEnd.Index).Value) + _
                        Amount - AmountAllocated, DefPrec)
                    End If
                End If
            Next
            'Me.DataGridView3.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub RepairMinus(ByVal TBLDetail As Inventory_Module.DseLaporanStockDetail2.DataTable1DataTable)
        Dim ij As Decimal = 0
        Dim rowDetail, rowDetail1, rowDetail2 As Inventory_Module.DseLaporanStockDetail2.DataTable1Row

        Try
            ij = 0
            'buang yang nol
            While ij <= TBLDetail.DefaultView.Count - 1
                rowDetail = TBLDetail.DefaultView(ij).Row
                If rowDetail.Qty = 0 Then
                    TBLDetail.Rows.Remove(rowDetail)
                Else
                    ij += 1
                End If
            End While

            'kompensasi yang negatif ke positif
            While TBLDetail.DefaultView.Count > 1
                rowDetail1 = TBLDetail.DefaultView(0).Row
                rowDetail2 = TBLDetail.DefaultView(TBLDetail.DefaultView.Count - 1).Row
                If rowDetail2.Qty < 0 And rowDetail1.Qty > 0 Then
                    If Math.Abs(rowDetail2.Qty) = rowDetail1.Qty Then
                        TBLDetail.Rows.Remove(rowDetail1)
                        TBLDetail.Rows.Remove(rowDetail2)
                    ElseIf Math.Abs(rowDetail2.Qty) > rowDetail1.Qty Then
                        rowDetail2.Qty += rowDetail1.Qty
                        TBLDetail.Rows.Remove(rowDetail1)
                    Else
                        rowDetail1.Qty += rowDetail2.Qty
                        TBLDetail.Rows.Remove(rowDetail2)
                    End If
                Else
                    Exit While
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub GetQtyAmountRekap(ByVal vPrimaryCodeNo As Double, ByRef vQty As Decimal, ByRef vAmount As Decimal, _
    ByRef Category As String, ByRef CodeNo As String, ByRef Desc As String, ByRef UnitCode As String)
        Dim rowDGV As DataGridViewRow
        vQty = 0
        vAmount = 0

        Try
            For Each rowDGV In Me.DataGridView1.Rows
                If rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryCodeNo Then
                    vQty = rowDGV.Cells(Me.ColQtyEnd.Index).Value
                    vAmount = rowDGV.Cells(Me.ColAmountEnd.Index).Value
                    Category = rowDGV.Cells(Me.ColCategory.Index).Value
                    CodeNo = rowDGV.Cells(Me.ColCodeNo.Index).Value
                    Desc = rowDGV.Cells(Me.ColDesc.Index).Value
                    UnitCode = rowDGV.Cells(Me.ColUnit.Index).Value

                    Exit Sub
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function LoadStockDetail(ByVal TBL As Inventory_Module.DseLaporanStockDetail2.DataTable1DataTable) As Boolean
        'Dim DS As New DseLaporanStockDetail2
        Dim DTA As New Inventory_Module.DseLaporanStockDetail2TableAdapters.DataTable1TableAdapter
        Dim Row As Inventory_Module.DseLaporanStockDetail2.DataTable1Row

        'TBLCetak.Clear()
        Dim ROWCetak, rowDummy As Inventory_Module.DseLaporanStockDetail2.TblCetakRow
        Dim tblDummy As New Inventory_Module.DseLaporanStockDetail2.TblCetakDataTable

        Dim ParamPerGrouping As String = ""
        Dim MinDate As Date
        Dim MaxDate As Date
        Dim PrimaryMax As Double = -1
        Dim PrimaryMin As Double = -1
        Dim Filter01 As String = ""
        Dim Filter01_II As String = ""
        Dim Filter02 As String = ""
        Dim Filter03 As String = ""
        Dim Filter04 As String = ""
        Dim ParamGrandTotal As String = ""
        Dim DV1 As DataRowView
        Dim DTAKonv As New Inventory_Module.DseLaporanStockDetail2TableAdapters.TblKonversiTableAdapter
        Dim TBLKonv As New Inventory_Module.DseLaporanStockDetail2.TblKonversiDataTable
        Dim rowKonv As Inventory_Module.DseLaporanStockDetail2.TblKonversiRow
        Dim ListPrimaryMinus As New List(Of Double)
        Dim IdxAcuan As Integer = 0
        Dim PrimaryCodeNo1, LotNo, PrimaryLocation, Mode As String

        Try

            'TBLCetak.Rows.Clear()
            'MaxDate = Me.DtpMaxExpiryDate.Value
            'MinDate = Me.DtpMinExpiryDate.Value
            TBL.Rows.Clear()
            DTA.Fill(TBL)
            TBL.DefaultView.Sort = "Primary"
            Filter01 = "[Primary] = [Primary]"
            Filter01_II = "[Primary] = [Primary]"

            ''================================================================================================================================================
            'If Me.ShowStockAsOff = EnumPilih.Yes Then
            Dim TBL_ As New Inventory_Module.DseLaporanStockDetail2.DataTable1DataTable
            Dim Row_, Row_1 As Inventory_Module.DseLaporanStockDetail2.DataTable1Row
            Dim dtaSC As New Inventory_Module.DseLaporanStockDetail2TableAdapters.DataTable5TableAdapter
            Dim tblSC As New Inventory_Module.DseLaporanStockDetail2.DataTable5DataTable
            Dim rowSC As Inventory_Module.DseLaporanStockDetail2.DataTable5Row
            Dim dtaLTF As New Inventory_Module.DseLaporanStockDetail2TableAdapters.DataTable6TableAdapter
            Dim tblLTF As New Inventory_Module.DseLaporanStockDetail2.DataTable6DataTable
            Dim rowLTF As Inventory_Module.DseLaporanStockDetail2.DataTable6Row
            Dim DTA_ST As New Inventory_Module.DseLaporanStockDetail2TableAdapters.DataTable7TableAdapter
            Dim TBL_ST As New Inventory_Module.DseLaporanStockDetail2.DataTable7DataTable
            Dim row_ST As Inventory_Module.DseLaporanStockDetail2.DataTable7Row

            For Each DV1 In TBL.DefaultView
                Row = DV1.Row
                Row_ = TBL_.NewDataTable1Row
                For i As Integer = 0 To TBL.Columns.Count - 1
                    Row_.Item(i) = Row.Item(i)
                Next
                TBL_.Rows.Add(Row_)
            Next

            'dtaSC.Connection.ConnectionString = Core_Module.GetConstring
            'dtaLTF.Connection.ConnectionString = Core_Module.GetConstring
            'DTA_ST.Connection.ConnectionString = Core_Module.GetConstring

            dtaSC.Fill(tblSC, Me.ObjCD.GetLasDateOfTheMonth(Me.PeriodClosing).Date.AddDays(1), CInt(Inventory_Module.ClsTransaksiStockCard.EnumJenisTransaksiAsal.Stock_Taking))
            dtaLTF.Fill(tblLTF, Me.ObjCD.GetLasDateOfTheMonth(Me.PeriodClosing).Date.AddDays(1))
            DTA_ST.Fill(TBL_ST, Me.ObjCD.GetLasDateOfTheMonth(Me.PeriodClosing).Date.AddDays(1))

            tblSC.DefaultView.RowFilter = Filter01
            tblLTF.DefaultView.RowFilter = Filter01_II
            TBL_ST.DefaultView.RowFilter = Filter04

            'Prg1.Maximum = tblSC.DefaultView.Count
            'Prg1.Value = 0
            'Prg1.Refresh()
            For Each DV1 In tblSC.DefaultView
                rowSC = DV1.Row
                'Prg1.Value += 1
                'Label4.Text = "Stock Card calculation... " & Prg1.Value.ToString & " of " & Prg1.Maximum.ToString & "."
                PrimaryCodeNo1 = rowSC.Primary_Product.ToString
                LotNo = rowSC.Lot_No.ToString
                PrimaryLocation = rowSC.Primary_Storage_Area.ToString
                Mode = "Stock Card"

                TBL_.DefaultView.RowFilter = "[Primary Product] = " & rowSC.Primary_Product.ToString & " AND [Lot No] = '" & Me.ObjCD.FormatSpecialChar_DataFilter(rowSC.Lot_No) & "' AND [Primary Storage Area] = " & rowSC.Primary_Storage_Area.ToString
                If TBL_.DefaultView.Count > 0 Then
                    Row_ = TBL_.DefaultView(0).Row
                    Row_.Qty = Row_.Qty + rowSC._Total__Out_ - rowSC.Total_In
                    If Row_.Qty = 0 Then
                        TBL_.Rows.Remove(Row_)
                    End If
                Else
                    Row_ = TBL_.NewDataTable1Row
                    Row_.Primary_Product = rowSC.Primary_Product
                    Row_.Product_Code = rowSC.Product_Code
                    Row_.Description = rowSC.Description
                    Row_._Decimal = rowSC._Decimal
                    Row_.Batch_No_Controlled = rowSC.Batch_No_Controlled
                    Row_.Location_Controlled = rowSC.Location_Controlled
                    Row_.Expiry_Date_Controlled = rowSC.Expiry_Date_Controlled
                    Row_.Type = rowSC.Type
                    Row_.Qty = rowSC._Total__Out_ - rowSC.Total_In
                    Row_.Primary_Unit = rowSC.Primary_Unit
                    Row_.Unit_Code = rowSC.Unit_Code
                    Row_.Category_Primary = rowSC.Category_Primary
                    Row_.Categorey_Code = rowSC.Categorey_Code
                    Row_.Category_Desc = rowSC.Categrory_Desc
                    Row_.Sub_Brand_Primary = rowSC.Sub_Brand_Primary
                    Row_.Sub_Brand_Code = rowSC.Sub_Brand_Code
                    Row_.Sub_Brand_Desc = rowSC.Sub_Brand_Desc
                    Row_.Brand_Primary = rowSC.Brand_Primary
                    Row_.Brand_Code = rowSC.Brand_Code
                    Row_.Brand_Desc = rowSC.Brand_Desc
                    Row_.Group_Primary = rowSC.Group_Primary
                    Row_.Group_Code = rowSC.Group_Code
                    Row_.Group_Desc = rowSC.Group_Desc
                    Row_.Expiry_Date = rowSC.Expiry_Date
                    Row_.Primary_Storage_Area = rowSC.Primary_Storage_Area
                    Row_.Storage_Area_Code = rowSC.Storage_Area_Code
                    Row_.Storage_Area_Desc = rowSC.Storage_Area_Desc
                    Row_.Lot_No = rowSC.Lot_No
                    Row_.Qty_Secondary = -1
                    Row_.Warehouse_Code = rowSC.Warehouse_Code
                    Row_.Warehouse_Desc = rowSC.Warehouse_Desc

                    If Row_.Qty <> 0 Then
                        TBL_.Rows.Add(Row_)
                    End If
                End If
            Next
            'Label4.Text = "Finishing...."

            For Each DV1 In tblLTF.DefaultView
                rowLTF = DV1.Row
                PrimaryCodeNo1 = rowLTF.Primary_Product.ToString
                LotNo = rowLTF.Lot_No.ToString
                PrimaryLocation = rowLTF.Primary_Storage_Area_II.ToString
                Mode = "Stock Transfer Tujuan"

                TBL_.DefaultView.RowFilter = "[Primary Product] = " & rowLTF.Primary_Product.ToString & " AND [Lot No] = '" & Me.ObjCD.FormatSpecialChar_DataFilter(rowLTF.Lot_No) & "' AND [Primary Storage Area] = " & rowLTF.Primary_Storage_Area_II.ToString & ""
                If TBL_.DefaultView.Count > 0 Then
                    Row_ = TBL_.DefaultView(0).Row
                    Row_.Qty = Row_.Qty - rowLTF.Qty
                    If Row_.Qty = 0 Then
                        TBL_.Rows.Remove(Row_)
                    End If
                Else
                    If rowLTF.Qty > 0 Then
                        Row_ = TBL_.NewDataTable1Row
                        Row_.Primary_Product = rowLTF.Primary_Product
                        Row_.Product_Code = rowLTF.Product_Code
                        Row_.Description = rowLTF.Description
                        Row_._Decimal = rowLTF._Decimal
                        Row_.Batch_No_Controlled = rowLTF.Batch_No_Controlled
                        Row_.Location_Controlled = rowLTF.Location_Controlled
                        Row_.Expiry_Date_Controlled = rowLTF.Expiry_Date_Controlled
                        Row_.Type = rowLTF.Type
                        Row_.Qty = rowLTF.Qty * -1
                        Row_.Primary_Unit = rowLTF.Primary_Unit
                        Row_.Unit_Code = rowLTF.Unit_Code
                        Row_.Category_Primary = rowLTF.Category_Primary
                        Row_.Categorey_Code = rowLTF.Categorey_Code
                        Row_.Category_Desc = rowLTF.Category_Desc
                        Row_.Sub_Brand_Primary = rowLTF.Sub_Brand_Primary
                        Row_.Sub_Brand_Code = rowLTF.Sub_Brand_Code
                        Row_.Sub_Brand_Desc = rowLTF.Sub_Brand_Desc
                        Row_.Brand_Primary = rowLTF.Brand_Primary
                        Row_.Brand_Code = rowLTF.Brand_Code
                        Row_.Brand_Desc = rowLTF.Brand_Desc
                        Row_.Group_Primary = rowLTF.Group_Primary
                        Row_.Group_Code = rowLTF.Group_Code
                        Row_.Group_Desc = rowLTF.Group_Desc
                        Row_.Expiry_Date = rowLTF.Expiry_Date
                        Row_.Primary_Storage_Area = rowLTF.Primary_Storage_Area_II
                        Row_.Storage_Area_Code = rowLTF.Storage_Area_Code_II
                        Row_.Storage_Area_Desc = rowLTF.Storage_Area_Desc_II
                        Row_.Lot_No = rowLTF.Lot_No
                        Row_.Qty_Secondary = -1
                        Row_.Warehouse_Code = rowLTF.Warehouse_Code_II
                        Row_.Warehouse_Desc = rowLTF.Warehouse_Desc_II
                        If Row_.Qty <> 0 Then
                            TBL_.Rows.Add(Row_)
                        End If
                    End If
                End If

                PrimaryCodeNo1 = rowLTF.Primary_Product.ToString
                LotNo = rowLTF.Lot_No.ToString
                PrimaryLocation = rowLTF.Primary_Storage_Area.ToString
                Mode = "Stock Transfer Asal"
                TBL_.DefaultView.RowFilter = "[Primary Product] = " & rowLTF.Primary_Product.ToString & " AND [Lot No] = '" & Me.ObjCD.FormatSpecialChar_DataFilter(rowLTF.Lot_No) & "' AND [Primary Storage Area] = " & rowLTF.Primary_Storage_Area.ToString & ""
                If TBL_.DefaultView.Count > 0 Then
                    Row_ = TBL_.DefaultView(0).Row
                    Row_.Qty = Row_.Qty + rowLTF.Qty
                    If Row_.Qty = 0 Then
                        TBL_.Rows.Remove(Row_)
                    End If
                Else
                    If rowLTF.Qty > 0 Then
                        Row_ = TBL_.NewDataTable1Row
                        Row_.Primary_Product = rowLTF.Primary_Product
                        Row_.Product_Code = rowLTF.Product_Code
                        Row_.Description = rowLTF.Description
                        Row_._Decimal = rowLTF._Decimal
                        Row_.Batch_No_Controlled = rowLTF.Batch_No_Controlled
                        Row_.Location_Controlled = rowLTF.Location_Controlled
                        Row_.Expiry_Date_Controlled = rowLTF.Expiry_Date_Controlled
                        Row_.Type = rowLTF.Type
                        Row_.Qty = rowLTF.Qty
                        Row_.Primary_Unit = rowLTF.Primary_Unit
                        Row_.Unit_Code = rowLTF.Unit_Code
                        Row_.Category_Primary = rowLTF.Category_Primary
                        Row_.Categorey_Code = rowLTF.Categorey_Code
                        Row_.Category_Desc = rowLTF.Categorey_Code
                        Row_.Sub_Brand_Primary = rowLTF.Sub_Brand_Primary
                        Row_.Sub_Brand_Code = rowLTF.Sub_Brand_Code
                        Row_.Sub_Brand_Desc = rowLTF.Sub_Brand_Desc
                        Row_.Brand_Primary = rowLTF.Brand_Primary
                        Row_.Brand_Code = rowLTF.Brand_Code
                        Row_.Brand_Desc = rowLTF.Brand_Desc
                        Row_.Group_Primary = rowLTF.Group_Primary
                        Row_.Group_Code = rowLTF.Group_Code
                        Row_.Group_Desc = rowLTF.Group_Desc
                        Row_.Expiry_Date = rowLTF.Expiry_Date
                        Row_.Primary_Storage_Area = rowLTF.Primary_Storage_Area
                        Row_.Storage_Area_Code = rowLTF.Storage_Area_Code
                        Row_.Storage_Area_Desc = rowLTF.Storage_Area_Desc
                        Row_.Lot_No = rowLTF.Lot_No
                        Row_.Qty_Secondary = -1
                        Row_.Warehouse_Code = rowLTF.Warehouse_Code
                        Row_.Warehouse_Desc = rowLTF.Warehouse_Desc
                        If Row_.Qty <> 0 Then
                            TBL_.Rows.Add(Row_)
                        End If
                    End If
                End If
            Next

            'Prg1.Value = 0
            'Prg1.Maximum = TBL_ST.DefaultView.Count
            'Prg1.Refresh()
            For Each DV1 In TBL_ST.DefaultView
                row_ST = DV1.Row
                'Prg1.Value += 1
                'Label4.Text = "Stock Taking calculation... " & Prg1.Value.ToString & " of " & Prg1.Maximum.ToString & "."
                PrimaryCodeNo1 = row_ST.Primary_Product.ToString
                LotNo = row_ST.Batch_No.ToString
                PrimaryLocation = row_ST.Primary_Location.ToString
                Mode = "Stock Taking"

                TBL_.DefaultView.RowFilter = "[Primary Product] = " & row_ST.Primary_Product.ToString & " AND [Lot No] = '" & Me.ObjCD.FormatSpecialChar_DataFilter(row_ST.Batch_No) & "' AND [Primary Storage Area] = " & row_ST.Primary_Location.ToString
                If TBL_.DefaultView.Count > 0 Then
                    Row_ = TBL_.DefaultView(0).Row
                    Row_.Qty = Row_.Qty - row_ST.Qty
                    If Row_.Qty = 0 Then
                        TBL_.Rows.Remove(Row_)
                    End If
                Else
                    Row_ = TBL_.NewDataTable1Row
                    Row_.Primary_Product = row_ST.Primary_Product
                    Row_.Product_Code = row_ST.Product_Code
                    Row_.Description = row_ST.Description
                    Row_._Decimal = row_ST._Decimal
                    Row_.Batch_No_Controlled = row_ST.Batch_No_Controlled
                    Row_.Location_Controlled = row_ST.Location_Controlled
                    Row_.Expiry_Date_Controlled = row_ST.Expiry_Date_Controlled
                    Row_.Type = row_ST.Type
                    Row_.Qty = row_ST.Qty * (-1)
                    Row_.Primary_Unit = row_ST.Primary_Unit
                    Row_.Unit_Code = row_ST.Unit_Code
                    Row_.Category_Primary = row_ST.Category_Primary
                    Row_.Categorey_Code = row_ST.Category_Code
                    Row_.Category_Desc = row_ST.Category_Desc
                    Row_.Sub_Brand_Primary = row_ST.Sub_Brand_Primary
                    Row_.Sub_Brand_Code = row_ST.Sub_Brand_Code
                    Row_.Sub_Brand_Desc = row_ST.Sub_Brand_Desc
                    Row_.Brand_Primary = row_ST.Brand_Primary
                    Row_.Brand_Code = row_ST.Brand_Code
                    Row_.Brand_Desc = row_ST.Brand_Desc
                    Row_.Group_Primary = row_ST.Group_Primary
                    Row_.Group_Code = row_ST.Product_Group_Code
                    Row_.Group_Desc = row_ST.Product_Group_Desc
                    Row_.Expiry_Date = row_ST.Expiry_Date
                    Row_.Primary_Storage_Area = row_ST.Primary_Location
                    Row_.Storage_Area_Code = row_ST.Storage_Area_Code
                    Row_.Storage_Area_Desc = row_ST.Storage_Area_Desc
                    Row_.Lot_No = row_ST.Batch_No
                    Row_.Qty_Secondary = -1
                    Row_.Warehouse_Code = row_ST.Warehouse_Code
                    Row_.Warehouse_Desc = row_ST.Warehouse_Desc
                    If Row_.Qty <> 0 Then
                        TBL_.Rows.Add(Row_)
                    End If
                End If
            Next
            'Label4.Text = "Finishing...."

            'Dim listOfPrimaryLocation As New List(Of Double)
            'For Each dgvRow As DataGridViewRow In Me.DgvLocation.Rows
            '    If IsNothing(dgvRow.Cells(Me.ColLocation_Primary.Index).Value) = False Then
            '        If listOfPrimaryLocation.Contains(dgvRow.Cells(Me.ColLocation_Primary.Index).Value) = False Then
            '            listOfPrimaryLocation.Add(dgvRow.Cells(Me.ColLocation_Primary.Index).Value)
            '        End If
            '    End If
            'Next

            'cek minus
            For Each Row_ In TBL_
                If Row_.Qty < 0 Then
                    If ListPrimaryMinus.Contains(Row_.Primary_Product) = False Then
                        ListPrimaryMinus.Add(Row_.Primary_Product)
                    End If
                End If
            Next
            For Each PrimaryProduct As Double In ListPrimaryMinus
                TBL_.DefaultView.RowFilter = "[Primary Product] = " + PrimaryProduct.ToString
                TBL_.DefaultView.Sort = "Qty ASC"

                IdxAcuan = 0
                Row_ = TBL_.DefaultView(IdxAcuan).Row
                While TBL_.DefaultView.Count > 1 And IdxAcuan <= TBL_.DefaultView.Count - 1
                    Row_ = TBL_.DefaultView(IdxAcuan).Row
                    Row_1 = TBL_.DefaultView(TBL_.DefaultView.Count - 1).Row

                    If Row_1.Qty < 0 Then
                        Exit While
                    ElseIf Row_.Qty > 0 Then
                        Exit While
                    Else
                        If Row_.Qty + Row_1.Qty = 0 Then
                            TBL_.Rows.Remove(Row_)
                            TBL_.Rows.Remove(Row_1)
                        ElseIf Row_.Qty + Row_1.Qty > 0 Then
                            'lebih besar row_1.qty
                            'buang row_
                            Row_1.Qty += Row_.Qty
                            TBL_.Rows.Remove(Row_)
                        ElseIf Row_.Qty + Row_1.Qty < 0 Then
                            'lebih besar row_1.qty
                            'buang row_
                            Row_.Qty += Row_1.Qty
                            TBL_.Rows.Remove(Row_1)
                        End If
                    End If
                    'IdxAcuan += 1
                End While
            Next

            TBL.Rows.Clear()
            For Each Row_ In TBL_
                Row = TBL.NewDataTable1Row
                For i As Integer = 0 To TBL.Columns.Count - 1
                    Row.Item(i) = Row_.Item(i)
                Next
                TBL.Rows.Add(Row)
            Next
            TBL.DefaultView.RowFilter = ""

            'End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            'DS.Dispose()
            DTA.Dispose()
            TBL.Dispose()
            DTAKonv.Dispose()
            TBLKonv.Dispose()
            'DisableControls(False)
        End Try
    End Function

    Public Function LoadStockDetail(ByVal TBL As Inventory_Module.DseLaporanStockDetail2.DataTable1DataTable, ByVal vPrimaryCodeNo As Double) As Boolean
        'Dim DS As New DseLaporanStockDetail2
        Dim DTA As New Inventory_Module.DseLaporanStockDetail2TableAdapters.DataTable1TableAdapter
        Dim Row As Inventory_Module.DseLaporanStockDetail2.DataTable1Row

        'TBLCetak.Clear()
        'Dim ROWCetak, rowDummy As Inventory_Module.DseLaporanStockDetail2.TblCetakRow
        Dim tblDummy As New Inventory_Module.DseLaporanStockDetail2.TblCetakDataTable

        Dim ParamPerGrouping As String = ""
        'Dim MinDate As Date
        'Dim MaxDate As Date
        Dim PrimaryMax As Double = -1
        Dim PrimaryMin As Double = -1
        Dim Filter01 As String = ""
        Dim Filter01_II As String = ""
        Dim Filter02 As String = ""
        Dim Filter03 As String = ""
        Dim Filter04 As String = ""
        Dim ParamGrandTotal As String = ""
        Dim DV1 As DataRowView
        Dim DTAKonv As New Inventory_Module.DseLaporanStockDetail2TableAdapters.TblKonversiTableAdapter
        Dim TBLKonv As New Inventory_Module.DseLaporanStockDetail2.TblKonversiDataTable
        Dim rowKonv As Inventory_Module.DseLaporanStockDetail2.TblKonversiRow
        Dim ListPrimaryMinus As New List(Of Double)
        Dim IdxAcuan As Integer = 0

        Try
            'Label4.Text = "Initializing..."
            'DisableControls(True)
            'Prg1.Maximum = 100
            'Prg1.Value = 10

            'DTA.Connection.ConnectionString = Core_Module.GetConstring
            'DTAKonv.Connection.ConnectionString = Core_Module.GetConstring

            'TBLCetak.Rows.Clear()
            'MaxDate = Me.DtpMaxExpiryDate.Value
            'MinDate = Me.DtpMinExpiryDate.Value
            DTA.FillBy_PrimaryProduct1(TBL, vPrimaryCodeNo)
            TBL.DefaultView.Sort = "Primary"
            Filter01 = "[Primary] = [Primary]"
            Filter01_II = "[Primary] = [Primary]"


            ''================================================================================================================================================
            'If Me.ShowStockAsOff = EnumPilih.Yes Then
            Dim TBL_ As New Inventory_Module.DseLaporanStockDetail2.DataTable1DataTable
            Dim Row_, Row_1 As Inventory_Module.DseLaporanStockDetail2.DataTable1Row
            Dim dtaSC As New Inventory_Module.DseLaporanStockDetail2TableAdapters.DataTable5TableAdapter
            Dim tblSC As New Inventory_Module.DseLaporanStockDetail2.DataTable5DataTable
            Dim rowSC As Inventory_Module.DseLaporanStockDetail2.DataTable5Row
            Dim dtaLTF As New Inventory_Module.DseLaporanStockDetail2TableAdapters.DataTable6TableAdapter
            Dim tblLTF As New Inventory_Module.DseLaporanStockDetail2.DataTable6DataTable
            Dim rowLTF As Inventory_Module.DseLaporanStockDetail2.DataTable6Row
            Dim DTA_ST As New Inventory_Module.DseLaporanStockDetail2TableAdapters.DataTable7TableAdapter
            Dim TBL_ST As New Inventory_Module.DseLaporanStockDetail2.DataTable7DataTable
            Dim row_ST As Inventory_Module.DseLaporanStockDetail2.DataTable7Row

            For Each DV1 In TBL.DefaultView
                Row = DV1.Row
                Row_ = TBL_.NewDataTable1Row
                For i As Integer = 0 To TBL.Columns.Count - 1
                    Row_.Item(i) = Row.Item(i)
                Next
                TBL_.Rows.Add(Row_)
            Next

            'dtaSC.Connection.ConnectionString = Core_Module.GetConstring
            'dtaLTF.Connection.ConnectionString = Core_Module.GetConstring
            'DTA_ST.Connection.ConnectionString = Core_Module.GetConstring

            dtaSC.FillBy_PrimaryProduct(tblSC, vPrimaryCodeNo, Me.ObjCD.GetLasDateOfTheMonth(Me.PeriodClosing).Date.AddDays(1), _
            CInt(Inventory_Module.ClsTransaksiStockCard.EnumJenisTransaksiAsal.Stock_Taking))
            dtaLTF.FillBy_Primary_Product(tblLTF, vPrimaryCodeNo, Me.ObjCD.GetLasDateOfTheMonth(Me.PeriodClosing).Date.AddDays(1))
            DTA_ST.FillBy_PrimaryProduct(TBL_ST, Me.ObjCD.GetLasDateOfTheMonth(Me.PeriodClosing).Date.AddDays(1), vPrimaryCodeNo)

            tblSC.DefaultView.RowFilter = Filter01
            tblLTF.DefaultView.RowFilter = Filter01_II
            TBL_ST.DefaultView.RowFilter = Filter04

            'Prg1.Maximum = tblSC.DefaultView.Count
            'Prg1.Value = 0
            'Prg1.Refresh()
            For Each DV1 In tblSC.DefaultView
                rowSC = DV1.Row
                'Prg1.Value += 1
                'Label4.Text = "Stock Card calculation... " & Prg1.Value.ToString & " of " & Prg1.Maximum.ToString & "."
                TBL_.DefaultView.RowFilter = "[Primary Product] = " & rowSC.Primary_Product.ToString & " AND [Lot No] = '" & rowSC.Lot_No & "' AND [Primary Storage Area] = " & rowSC.Primary_Storage_Area.ToString
                If TBL_.DefaultView.Count > 0 Then
                    Row_ = TBL_.DefaultView(0).Row
                    Row_.Qty = Row_.Qty + rowSC._Total__Out_ - rowSC.Total_In
                    If Row_.Qty = 0 Then
                        TBL_.Rows.Remove(Row_)
                    End If
                Else
                    Row_ = TBL_.NewDataTable1Row
                    Row_.Primary_Product = rowSC.Primary_Product
                    Row_.Product_Code = rowSC.Product_Code
                    Row_.Description = rowSC.Description
                    Row_._Decimal = rowSC._Decimal
                    Row_.Batch_No_Controlled = rowSC.Batch_No_Controlled
                    Row_.Location_Controlled = rowSC.Location_Controlled
                    Row_.Expiry_Date_Controlled = rowSC.Expiry_Date_Controlled
                    Row_.Type = rowSC.Type
                    Row_.Qty = rowSC._Total__Out_ - rowSC.Total_In
                    Row_.Primary_Unit = rowSC.Primary_Unit
                    Row_.Unit_Code = rowSC.Unit_Code
                    Row_.Category_Primary = rowSC.Category_Primary
                    Row_.Categorey_Code = rowSC.Categorey_Code
                    Row_.Category_Desc = rowSC.Categrory_Desc
                    Row_.Sub_Brand_Primary = rowSC.Sub_Brand_Primary
                    Row_.Sub_Brand_Code = rowSC.Sub_Brand_Code
                    Row_.Sub_Brand_Desc = rowSC.Sub_Brand_Desc
                    Row_.Brand_Primary = rowSC.Brand_Primary
                    Row_.Brand_Code = rowSC.Brand_Code
                    Row_.Brand_Desc = rowSC.Brand_Desc
                    Row_.Group_Primary = rowSC.Group_Primary
                    Row_.Group_Code = rowSC.Group_Code
                    Row_.Group_Desc = rowSC.Group_Desc
                    Row_.Expiry_Date = rowSC.Expiry_Date
                    Row_.Primary_Storage_Area = rowSC.Primary_Storage_Area
                    Row_.Storage_Area_Code = rowSC.Storage_Area_Code
                    Row_.Storage_Area_Desc = rowSC.Storage_Area_Desc
                    Row_.Lot_No = rowSC.Lot_No
                    Row_.Qty_Secondary = -1
                    Row_.Warehouse_Code = rowSC.Warehouse_Code
                    Row_.Warehouse_Desc = rowSC.Warehouse_Desc

                    If Row_.Qty <> 0 Then
                        TBL_.Rows.Add(Row_)
                    End If
                End If
            Next
            'Label4.Text = "Finishing...."

            For Each DV1 In tblLTF.DefaultView
                rowLTF = DV1.Row
                TBL_.DefaultView.RowFilter = "[Primary Product] = " & rowLTF.Primary_Product.ToString & " AND [Lot No] = '" & rowLTF.Lot_No & "' AND [Primary Storage Area] = " & rowLTF.Primary_Storage_Area_II.ToString & ""
                If TBL_.DefaultView.Count > 0 Then
                    Row_ = TBL_.DefaultView(0).Row
                    Row_.Qty = Row_.Qty - rowLTF.Qty
                    If Row_.Qty = 0 Then
                        TBL_.Rows.Remove(Row_)
                    End If
                Else
                    If rowLTF.Qty > 0 Then
                        Row_ = TBL_.NewDataTable1Row
                        Row_.Primary_Product = rowLTF.Primary_Product
                        Row_.Product_Code = rowLTF.Product_Code
                        Row_.Description = rowLTF.Description
                        Row_._Decimal = rowLTF._Decimal
                        Row_.Batch_No_Controlled = rowLTF.Batch_No_Controlled
                        Row_.Location_Controlled = rowLTF.Location_Controlled
                        Row_.Expiry_Date_Controlled = rowLTF.Expiry_Date_Controlled
                        Row_.Type = rowLTF.Type
                        Row_.Qty = rowLTF.Qty * -1
                        Row_.Primary_Unit = rowLTF.Primary_Unit
                        Row_.Unit_Code = rowLTF.Unit_Code
                        Row_.Category_Primary = rowLTF.Category_Primary
                        Row_.Categorey_Code = rowLTF.Categorey_Code
                        Row_.Category_Desc = rowLTF.Category_Desc
                        Row_.Sub_Brand_Primary = rowLTF.Sub_Brand_Primary
                        Row_.Sub_Brand_Code = rowLTF.Sub_Brand_Code
                        Row_.Sub_Brand_Desc = rowLTF.Sub_Brand_Desc
                        Row_.Brand_Primary = rowLTF.Brand_Primary
                        Row_.Brand_Code = rowLTF.Brand_Code
                        Row_.Brand_Desc = rowLTF.Brand_Desc
                        Row_.Group_Primary = rowLTF.Group_Primary
                        Row_.Group_Code = rowLTF.Group_Code
                        Row_.Group_Desc = rowLTF.Group_Desc
                        Row_.Expiry_Date = rowLTF.Expiry_Date
                        Row_.Primary_Storage_Area = rowLTF.Primary_Storage_Area_II
                        Row_.Storage_Area_Code = rowLTF.Storage_Area_Code_II
                        Row_.Storage_Area_Desc = rowLTF.Storage_Area_Desc_II
                        Row_.Lot_No = rowLTF.Lot_No
                        Row_.Qty_Secondary = -1
                        Row_.Warehouse_Code = rowLTF.Warehouse_Code_II
                        Row_.Warehouse_Desc = rowLTF.Warehouse_Desc_II
                        If Row_.Qty <> 0 Then
                            TBL_.Rows.Add(Row_)
                        End If
                    End If
                End If

                TBL_.DefaultView.RowFilter = "[Primary Product] = " & rowLTF.Primary_Product.ToString & " AND [Lot No] = '" & rowLTF.Lot_No & "' AND [Primary Storage Area] = " & rowLTF.Primary_Storage_Area.ToString & ""
                If TBL_.DefaultView.Count > 0 Then
                    Row_ = TBL_.DefaultView(0).Row
                    Row_.Qty = Row_.Qty + rowLTF.Qty
                    If Row_.Qty = 0 Then
                        TBL_.Rows.Remove(Row_)
                    End If
                Else
                    If rowLTF.Qty > 0 Then
                        Row_ = TBL_.NewDataTable1Row
                        Row_.Primary_Product = rowLTF.Primary_Product
                        Row_.Product_Code = rowLTF.Product_Code
                        Row_.Description = rowLTF.Description
                        Row_._Decimal = rowLTF._Decimal
                        Row_.Batch_No_Controlled = rowLTF.Batch_No_Controlled
                        Row_.Location_Controlled = rowLTF.Location_Controlled
                        Row_.Expiry_Date_Controlled = rowLTF.Expiry_Date_Controlled
                        Row_.Type = rowLTF.Type
                        Row_.Qty = rowLTF.Qty
                        Row_.Primary_Unit = rowLTF.Primary_Unit
                        Row_.Unit_Code = rowLTF.Unit_Code
                        Row_.Category_Primary = rowLTF.Category_Primary
                        Row_.Categorey_Code = rowLTF.Categorey_Code
                        Row_.Category_Desc = rowLTF.Categorey_Code
                        Row_.Sub_Brand_Primary = rowLTF.Sub_Brand_Primary
                        Row_.Sub_Brand_Code = rowLTF.Sub_Brand_Code
                        Row_.Sub_Brand_Desc = rowLTF.Sub_Brand_Desc
                        Row_.Brand_Primary = rowLTF.Brand_Primary
                        Row_.Brand_Code = rowLTF.Brand_Code
                        Row_.Brand_Desc = rowLTF.Brand_Desc
                        Row_.Group_Primary = rowLTF.Group_Primary
                        Row_.Group_Code = rowLTF.Group_Code
                        Row_.Group_Desc = rowLTF.Group_Desc
                        Row_.Expiry_Date = rowLTF.Expiry_Date
                        Row_.Primary_Storage_Area = rowLTF.Primary_Storage_Area
                        Row_.Storage_Area_Code = rowLTF.Storage_Area_Code
                        Row_.Storage_Area_Desc = rowLTF.Storage_Area_Desc
                        Row_.Lot_No = rowLTF.Lot_No
                        Row_.Qty_Secondary = -1
                        Row_.Warehouse_Code = rowLTF.Warehouse_Code
                        Row_.Warehouse_Desc = rowLTF.Warehouse_Desc
                        If Row_.Qty <> 0 Then
                            TBL_.Rows.Add(Row_)
                        End If
                    End If
                End If
            Next

            'Prg1.Value = 0
            'Prg1.Maximum = TBL_ST.DefaultView.Count
            'Prg1.Refresh()
            For Each DV1 In TBL_ST.DefaultView
                row_ST = DV1.Row
                'Prg1.Value += 1
                'Label4.Text = "Stock Taking calculation... " & Prg1.Value.ToString & " of " & Prg1.Maximum.ToString & "."
                TBL_.DefaultView.RowFilter = "[Primary Product] = " & row_ST.Primary_Product.ToString & " AND [Lot No] = '" & row_ST.Batch_No & "' AND [Primary Storage Area] = " & row_ST.Primary_Location.ToString
                If TBL_.DefaultView.Count > 0 Then
                    Row_ = TBL_.DefaultView(0).Row
                    Row_.Qty = Row_.Qty - row_ST.Qty
                    If Row_.Qty = 0 Then
                        TBL_.Rows.Remove(Row_)
                    End If
                Else
                    Row_ = TBL_.NewDataTable1Row
                    Row_.Primary_Product = row_ST.Primary_Product
                    Row_.Product_Code = row_ST.Product_Code
                    Row_.Description = row_ST.Description
                    Row_._Decimal = row_ST._Decimal
                    Row_.Batch_No_Controlled = row_ST.Batch_No_Controlled
                    Row_.Location_Controlled = row_ST.Location_Controlled
                    Row_.Expiry_Date_Controlled = row_ST.Expiry_Date_Controlled
                    Row_.Type = row_ST.Type
                    Row_.Qty = row_ST.Qty * (-1)
                    Row_.Primary_Unit = row_ST.Primary_Unit
                    Row_.Unit_Code = row_ST.Unit_Code
                    Row_.Category_Primary = row_ST.Category_Primary
                    Row_.Categorey_Code = row_ST.Category_Code
                    Row_.Category_Desc = row_ST.Category_Desc
                    Row_.Sub_Brand_Primary = row_ST.Sub_Brand_Primary
                    Row_.Sub_Brand_Code = row_ST.Sub_Brand_Code
                    Row_.Sub_Brand_Desc = row_ST.Sub_Brand_Desc
                    Row_.Brand_Primary = row_ST.Brand_Primary
                    Row_.Brand_Code = row_ST.Brand_Code
                    Row_.Brand_Desc = row_ST.Brand_Desc
                    Row_.Group_Primary = row_ST.Group_Primary
                    Row_.Group_Code = row_ST.Product_Group_Code
                    Row_.Group_Desc = row_ST.Product_Group_Desc
                    Row_.Expiry_Date = row_ST.Expiry_Date
                    Row_.Primary_Storage_Area = row_ST.Primary_Location
                    Row_.Storage_Area_Code = row_ST.Storage_Area_Code
                    Row_.Storage_Area_Desc = row_ST.Storage_Area_Desc
                    Row_.Lot_No = row_ST.Batch_No
                    Row_.Qty_Secondary = -1
                    Row_.Warehouse_Code = row_ST.Warehouse_Code
                    Row_.Warehouse_Desc = row_ST.Warehouse_Desc
                    If Row_.Qty <> 0 Then
                        TBL_.Rows.Add(Row_)
                    End If
                End If
            Next
            'Label4.Text = "Finishing...."

            'Dim listOfPrimaryLocation As New List(Of Double)
            'For Each dgvRow As DataGridViewRow In Me.DgvLocation.Rows
            '    If IsNothing(dgvRow.Cells(Me.ColLocation_Primary.Index).Value) = False Then
            '        If listOfPrimaryLocation.Contains(dgvRow.Cells(Me.ColLocation_Primary.Index).Value) = False Then
            '            listOfPrimaryLocation.Add(dgvRow.Cells(Me.ColLocation_Primary.Index).Value)
            '        End If
            '    End If
            'Next

            'cek minus
            For Each Row_ In TBL_
                If Row_.Qty < 0 Then
                    If ListPrimaryMinus.Contains(Row_.Primary_Product) = False Then
                        ListPrimaryMinus.Add(Row_.Primary_Product)
                    End If
                End If
            Next
            For Each PrimaryProduct As Double In ListPrimaryMinus
                TBL_.DefaultView.RowFilter = "[Primary Product] = " + PrimaryProduct.ToString
                TBL_.DefaultView.Sort = "Qty ASC"

                IdxAcuan = 0
                Row_ = TBL_.DefaultView(IdxAcuan).Row
                While TBL_.DefaultView.Count > 1 And IdxAcuan <= TBL_.DefaultView.Count - 1
                    Row_ = TBL_.DefaultView(IdxAcuan).Row
                    Row_1 = TBL_.DefaultView(TBL_.DefaultView.Count - 1).Row

                    If Row_1.Qty < 0 Then
                        Exit While
                    ElseIf Row_.Qty > 0 Then
                        Exit While
                    Else
                        If Row_.Qty + Row_1.Qty = 0 Then
                            TBL_.Rows.Remove(Row_)
                            TBL_.Rows.Remove(Row_1)
                        ElseIf Row_.Qty + Row_1.Qty > 0 Then
                            'lebih besar row_1.qty
                            'buang row_
                            Row_1.Qty += Row_.Qty
                            TBL_.Rows.Remove(Row_)
                        ElseIf Row_.Qty + Row_1.Qty < 0 Then
                            'lebih besar row_1.qty
                            'buang row_
                            Row_.Qty += Row_1.Qty
                            TBL_.Rows.Remove(Row_1)
                        End If
                    End If
                    'IdxAcuan += 1
                End While
            Next

            TBL.Rows.Clear()
            For Each Row_ In TBL_
                Row = TBL.NewDataTable1Row
                For i As Integer = 0 To TBL.Columns.Count - 1
                    Row.Item(i) = Row_.Item(i)
                Next
                TBL.Rows.Add(Row)
            Next
            TBL.DefaultView.RowFilter = ""

            'End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            'DS.Dispose()
            DTA.Dispose()
            TBL.Dispose()
            DTAKonv.Dispose()
            TBLKonv.Dispose()
            'DisableControls(False)
        End Try
    End Function

    Private Sub IsikanDGVDetailWarehouse_Obselete()
        Dim TBL_Awal As New Dse2.DataTable2DataTable
        Dim row_Awal As Dse2.DataTable2Row
        Dim DTAProduct As New Dse2TableAdapters.tblmaster_productTableAdapter
        Dim TBLProduct As New Dse2.tblmaster_productDataTable
        Dim rowProduct As Dse2.tblmaster_productRow
        Dim rowDGV As DataGridViewRow
        Dim DTAQty As New Dse2TableAdapters.tblmaster_productqtyTableAdapter
        Dim TBLQty1 As New Dse2.tblmaster_productqtyDataTable
        Dim TblQty2 As New Dse2.tblmaster_productqtyDataTable
        Dim rowQty1, rowQty2 As Dse2.tblmaster_productqtyRow
        Dim TBLAkhir As New Dse2.DataTable3DataTable
        Dim rowAkhir As Dse2.DataTable3Row
        Dim DV As Data.DataRowView
        Dim DTASC As New Dse2TableAdapters.tbltransaksi_stock_cardTableAdapter
        Dim TBLSC As New Dse2.tbltransaksi_stock_cardDataTable
        Dim rowSC As Dse2.tbltransaksi_stock_cardRow
        Dim DTA_Transfer As New Dse2TableAdapters.tbltransaksi_stocktransfer_itemTableAdapter
        Dim TBL_Transfer As New Dse2.tbltransaksi_stocktransfer_itemDataTable
        Dim row_Transfer As Dse2.tbltransaksi_stocktransfer_itemRow
        Dim TglAcuan1 As Date = (Me.ObjCD.GetLasDateOfTheMonth(Me.PeriodClosing) + TimeSpan.FromDays(1)).Date
        Dim DTAWarehouse As New Dse2TableAdapters.DataTable4TableAdapter
        Dim TBLWarehouse As New Dse2.DataTable4DataTable
        Dim rowWarehouse As Dse2.DataTable4Row
        Dim TblQty3 As New Dse2.tblmaster_productqty1DataTable
        Dim rowQty3 As Dse2.tblmaster_productqty1Row
        Dim AdaMinus As Boolean = False
        Dim QtyMinus As Decimal = 0
        Dim AmountAllocated As Decimal = 0
        Dim QtyAllocated As Decimal = 0
        Dim DefPrec As Integer = Me.ObjCur.GetDecimal_DefaultCurrency
        Dim ProductCodeRef As String = ""

        Try
            TBL_Awal.Rows.Clear()
            TBLQty1.Rows.Clear()
            TBLSC.Rows.Clear()
            TBL_Transfer.Rows.Clear()
            TBLWarehouse.Rows.Clear()

            TBLProduct.Rows.Clear()

            DTAProduct.Fill(TBLProduct)
            DTAQty.Fill(TBLQty1)
            DTASC.Fill(TBLSC, TglAcuan1)
            DTA_Transfer.Fill(TBL_Transfer, TglAcuan1)
            DTAWarehouse.Fill(TBLWarehouse)

            For ij As Integer = 0 To Me.DataGridView1.Rows.Count - 1
                rowDGV = Me.DataGridView1.Rows(ij)
                If rowDGV.Cells(Me.ColCodeNo.Index).Value.ToString.Contains(" '''") = True Then
                    rowDGV.Cells(Me.ColCodeNo.Index).Value = rowDGV.Cells(Me.ColCodeNo.Index).Value.ToString.Replace(" '''", "")
                End If
                If rowDGV.Cells(Me.ColCodeNo.Index).Value.ToString.Contains("'") = True Then
                    rowDGV.Cells(Me.ColCodeNo.Index).Value = rowDGV.Cells(Me.ColCodeNo.Index).Value.ToString.Replace("'", "")
                End If

                If CDec(rowDGV.Cells(Me.ColQtyEnd.Index).Value) > 0 Then
                    TBLProduct.DefaultView.RowFilter = "[Product Code] = '" + rowDGV.Cells(Me.ColCodeNo.Index).Value.ToString + "'"
                    If TBLProduct.DefaultView.Count > 0 Then
                        rowProduct = TBLProduct.DefaultView(0).Row
                        row_Awal = TBL_Awal.NewDataTable2Row
                        With row_Awal
                            .Category_Code = rowProduct.Category_Code
                            .Code_No = rowProduct.Product_Code
                            .Desc = rowProduct.Description
                            .End_Amount = CDec(rowDGV.Cells(Me.ColAmountEnd.Index).Value)
                            .Primary_Code_No = rowProduct.Primary
                            .Qty_End = CDec(rowDGV.Cells(Me.ColQtyEnd.Index).Value)
                            .Unit = rowDGV.Cells(Me.ColUnit.Index).Value
                        End With
                        TBL_Awal.Rows.Add(row_Awal)
                    End If
                End If
            Next

            For Each row_Awal In TBL_Awal.Rows
                ProductCodeRef = row_Awal.Code_No

                'masukkan ke data tabel qty
                TblQty2.Rows.Clear()
                TblQty3.Rows.Clear()

                TBLQty1.DefaultView.RowFilter = "[Primary Product] = " + row_Awal.Primary_Code_No.ToString
                For Each DV In TBLQty1.DefaultView
                    rowQty1 = DV.Row
                    rowQty2 = TblQty2.Newtblmaster_productqtyRow
                    For ij As Integer = 0 To TBLQty1.Columns.Count - 1
                        rowQty2.Item(ij) = rowQty1.Item(ij)
                    Next
                    TblQty2.Rows.Add(rowQty2)
                Next

                'balikkan berdasarkan stock card
                TBLSC.DefaultView.RowFilter = "[Primary Product] = " + row_Awal.Primary_Code_No.ToString
                For Each DV In TBLSC.DefaultView
                    rowSC = DV.Row
                    If rowSC.Qty_In <> rowSC.Qty_Out Then
                        TblQty2.DefaultView.RowFilter = "[Primary Product] = " + rowSC.Primary_Product.ToString + " AND [Primary Storage Area] = " + rowSC.primary_location.ToString
                        If TblQty2.DefaultView.Count = 0 Then
                            'tambahkan baris baru
                            rowQty2 = TblQty2.Newtblmaster_productqtyRow
                            With rowQty2
                                .Primary_Product = rowSC.Primary_Product
                                .Primary_Storage_Area = rowSC.primary_location
                                .Qty = rowSC.Qty_Out - rowSC.Qty_In
                            End With
                            TblQty2.Rows.Add(rowQty2)
                        Else
                            'set qty-nya
                            rowQty2 = TblQty2.DefaultView(0).Row
                            rowQty2.Qty += (rowSC.Qty_Out - rowSC.Qty_In)
                        End If
                    End If
                Next

                'balikkan berdasarkan location transfer
                TBL_Transfer.DefaultView.RowFilter = "[Primary Product] = " + row_Awal.Primary_Code_No.ToString
                For Each DV In TBL_Transfer.DefaultView
                    row_Transfer = DV.Row
                    'kurangkan menggunakan primary location target
                    TblQty2.DefaultView.RowFilter = "[Primary Product] = " + row_Transfer.Primary_Product.ToString + " AND [Primary Storage Area] = " + row_Transfer.primary_destination_area.ToString
                    If TblQty2.DefaultView.Count = 0 Then
                        'tambahkan baris baru
                        rowQty2 = TblQty2.Newtblmaster_productqtyRow
                        With rowQty2
                            .Primary_Product = row_Transfer.Primary_Product
                            .Primary_Storage_Area = row_Transfer.primary_destination_area
                            .Qty = (row_Transfer.Qty * -1)
                        End With
                        TblQty2.Rows.Add(rowQty2)
                    Else
                        'set qty-nya
                        rowQty2 = TblQty2.DefaultView(0).Row
                        rowQty2.Qty -= row_Transfer.Qty
                    End If

                    'tambahkan menggunakan primary location sumber
                    TblQty2.DefaultView.RowFilter = "[Primary Product] = " + row_Transfer.Primary_Product.ToString + " AND [Primary Storage Area] = " + row_Transfer.Primary_Storage_Area.ToString
                    If TblQty2.DefaultView.Count = 0 Then
                        'tambahkan baris baru
                        rowQty2 = TblQty2.Newtblmaster_productqtyRow
                        With rowQty2
                            .Primary_Product = row_Transfer.Primary_Product
                            .Primary_Storage_Area = row_Transfer.Primary_Storage_Area
                            .Qty = row_Transfer.Qty
                        End With
                        TblQty2.Rows.Add(rowQty2)
                    Else
                        'set qty-nya
                        rowQty2 = TblQty2.DefaultView(0).Row
                        rowQty2.Qty += row_Transfer.Qty
                    End If
                Next

                'bagi ke beberapa warehouse
                For Each rowQty2 In TblQty2.Rows
                    TBLWarehouse.DefaultView.RowFilter = "Primary = " + rowQty2.Primary_Storage_Area.ToString
                    If TBLWarehouse.DefaultView.Count = 0 Then
                        TblQty3.DefaultView.RowFilter = "[Primary Product] = " + rowQty2.Primary_Product.ToString + " AND [Primary Warehouse] = -1"
                        If TblQty3.DefaultView.Count = 0 Then
                            'tambahkan baris baru
                            rowQty3 = TblQty3.Newtblmaster_productqty1Row
                            With rowQty3
                                .Primary_Product = rowQty2.Primary_Product
                                .Primary_Warehouse = -1
                                .Qty = rowQty2.Qty
                                .Warehouse_Code = "N/A"
                                .Warehouse_Desc = "N/A"
                            End With
                            TblQty3.Rows.Add(rowQty3)
                        Else
                            'tambahkan qty saja
                            rowQty3 = TblQty3.DefaultView(0).Row
                            rowQty3.Qty += rowQty2.Qty
                        End If
                    Else
                        rowWarehouse = TBLWarehouse.DefaultView(0).Row
                        TblQty3.DefaultView.RowFilter = "[Primary Product] = " + rowQty2.Primary_Product.ToString + " AND [Primary Warehouse] = " + rowWarehouse.Primary_Warehouse.ToString
                        If TblQty3.DefaultView.Count = 0 Then
                            'tambahkan baris baru
                            rowQty3 = TblQty3.Newtblmaster_productqty1Row
                            With rowQty3
                                .Primary_Product = rowQty2.Primary_Product
                                .Primary_Warehouse = rowWarehouse.Primary_Warehouse
                                .Qty = rowQty2.Qty
                                .Warehouse_Code = rowWarehouse.Warehouse_Code
                                .Warehouse_Desc = rowWarehouse.Warehouse_Desc
                            End With
                            TblQty3.Rows.Add(rowQty3)
                        Else
                            'tambahkan qty saja
                            rowQty3 = TblQty3.DefaultView(0).Row
                            rowQty3.Qty += rowQty2.Qty
                        End If
                    End If
                Next

                'yang minus dengan tidak minus, (tidak coding klo hanya ketinggalan yang minus saja)
                AdaMinus = False
                QtyMinus = 0
                For Each rowQty3 In TblQty3.Rows
                    If rowQty3.Qty < 0 Then
                        AdaMinus = True
                    End If
                Next
                If AdaMinus = True Then
                    'buang baris minus dan nol
                    For ij As Integer = TblQty3.Rows.Count - 1 To 0 Step -1
                        rowQty3 = TblQty3.Rows(ij)
                        If rowQty3.Qty <= 0 Then
                            QtyMinus += rowQty3.Qty
                            TblQty3.Rows.Remove(rowQty3)
                        End If
                    Next
                    'sebarkan minusnya
                    While QtyMinus < 0
                        If TblQty3.Rows.Count = 0 Then
                            'MsgBox("Logic error", MsgBoxStyle.Information, "Information")
                            Exit While
                        Else
                            rowQty3 = TblQty3.Rows(0)
                            If QtyMinus + rowQty3.Qty > 0 Then
                                rowQty3.Qty += QtyMinus
                                QtyMinus = 0
                            Else
                                QtyMinus += rowQty3.Qty
                                rowQty3.Qty = 0
                                TblQty3.Rows.RemoveAt(0)
                            End If
                        End If
                    End While
                End If

                'masukkan tbl akhir
                AmountAllocated = 0
                QtyAllocated = 0
                For Each rowQty3 In TblQty3.Rows
                    'tambahkan baris baru
                    rowAkhir = TBLAkhir.NewDataTable3Row
                    With rowAkhir
                        .Category_Code = row_Awal.Category_Code
                        .Code_No = row_Awal.Code_No
                        .Desc = row_Awal.Desc
                        .End_Amount = 0
                        .Primary_Code_No = row_Awal.Primary_Code_No
                        .Primary_Warehouse = rowQty3.Primary_Warehouse
                        .Qty_End = rowQty3.Qty
                        .Unit = row_Awal.Unit
                        .Warehouse_Code = rowQty3.Warehouse_Code
                        .Warehouse_Desc = rowQty3.Warehouse_Desc
                    End With
                    TBLAkhir.Rows.Add(rowAkhir)

                    'hitung amount per barisnya
                    If row_Awal.Qty_End = rowAkhir.Qty_End + QtyAllocated Then
                        'ambil sisanya langsung
                        rowAkhir.End_Amount = row_Awal.End_Amount - AmountAllocated
                    Else
                        'di bagis berdasarkan qty allocated per warehouse
                        rowAkhir.End_Amount = Math.Round(row_Awal.End_Amount / row_Awal.Qty_End, DefPrec, MidpointRounding.AwayFromZero) * rowAkhir.Qty_End
                        AmountAllocated += rowAkhir.End_Amount
                        QtyAllocated += rowAkhir.Qty_End
                    End If
                Next
            Next

            'masukkan ke dgv
            Me.DataGridView3.Rows.Clear()
            Me.DataGridView3.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            TBLAkhir.DefaultView.RowFilter = ""
            For Each DV In TBLAkhir.DefaultView
                rowAkhir = DV.Row
                If rowAkhir.Qty_End > 0 Then
                    Me.DataGridView3.Rows.Add()
                    rowDGV = Me.DataGridView3.Rows(Me.DataGridView3.Rows.Count - 1)
                    rowDGV.Cells(Me.ColW_AmountEnd.Index).Value = Me.ObjCD.DigitGrouping(rowAkhir.End_Amount, DefPrec)
                    rowDGV.Cells(Me.ColW_CategoryCode.Index).Value = rowAkhir.Category_Code
                    rowDGV.Cells(Me.ColW_CodeNo.Index).Value = rowAkhir.Code_No
                    rowDGV.Cells(Me.ColW_Desc.Index).Value = rowAkhir.Desc
                    rowDGV.Cells(Me.ColW_QtyEnd.Index).Value = Me.ObjCD.DigitGrouping(rowAkhir.Qty_End, 2)
                    rowDGV.Cells(Me.ColW_UnitCode.Index).Value = rowAkhir.Unit
                    rowDGV.Cells(Me.ColW_Warehouse.Index).Value = rowAkhir.Warehouse_Code + " - " + rowAkhir.Warehouse_Desc
                End If
            Next
            'Me.DataGridView3.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message + ": " + ProductCodeRef, MsgBoxStyle.Critical, "Error")
        Finally
            TBL_Awal.Dispose()
            DTAProduct.Dispose()
            TBLProduct.Dispose()
            DTAQty.Dispose()
            TBLQty1.Dispose()
            TblQty2.Dispose()
            DTASC.Dispose()
            TBLSC.Dispose()
            DTA_Transfer.Dispose()
            TBL_Transfer.Dispose()
            DTAWarehouse.Dispose()
            TBLWarehouse.Dispose()
        End Try
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim rowDGV As DataGridViewRow
        Dim DTA As New Dse3TableAdapters.tbltrans_custom_itemdetailTableAdapter

        Try
            If MsgBox("Save detail item?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                DTA.Delete()

                Me.Simpan_UsingScript()

                'For Each rowDGV In Me.DataGridView2.Rows
                '    With rowDGV
                '        If IsNothing(.Cells(Me.ColItem_TipeItem2.Index).Value) = True Then
                '            .Cells(Me.ColItem_TipeItem2.Index).Value = ""
                '        End If
                '        If IsNothing(.Cells(Me.ColItem_TipeDoc2.Index).Value) = True Then
                '            .Cells(Me.ColItem_TipeDoc2.Index).Value = ""
                '        End If
                '        If IsNothing(.Cells(Me.ColItem_DocNo2.Index).Value) = True Then
                '            .Cells(Me.ColItem_DocNo2.Index).Value = ""
                '        End If

                '        DTA.Insert(.Cells(Me.ColItem_No.Index).Value, .Cells(Me.ColItem_Category.Index).Value, .Cells(Me.ColCodeNo.Index).Value, .Cells(Me.ColItem_Desc.Index).Value, _
                '        .Cells(Me.ColItem_Unit.Index).Value, .Cells(Me.ColItem_Tgl.Index).Value, .Cells(Me.ColItem_TipeDoc.Index).Value, .Cells(Me.ColItem_TipeItem.Index).Value, _
                '        .Cells(Me.ColItem_NoDoc.Index).Value, .Cells(Me.ColItem_QtyBB.Index).Value, .Cells(Me.ColItem_AmountBB.Index).Value, .Cells(Me.ColItem_QtyIn.Index).Value, _
                '        .Cells(Me.ColItem_AmountIn.Index).Value, .Cells(Me.ColItem_QtyOut.Index).Value, .Cells(Me.ColItem_AmountOut.Index).Value, .Cells(Me.ColItem_QtyEnd.Index).Value, _
                '        .Cells(Me.ColItem_AmountEnd.Index).Value, .Cells(Me.ColItem_TipeItem2.Index).Value, .Cells(Me.ColItem_TipeDoc2.Index).Value, .Cells(Me.ColItem_DocNo2.Index).Value, _
                '        .Cells(Me.ColItem_DefPrec.Index).Value, .Cells(Me.ColItem_PrimaryTransaksiAsal.Index).Value)
                '    End With
                'Next

                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function Simpan_UsingScript() As Boolean
        Dim conStr As String = "server=server;user id=root;Password=#SMeSS_p4ss*897%#;persist security info=True;database=core_db;port=3307;port=3307"
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 50
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < Me.DataGridView2.Rows.Count
                sql = Me.GetInsertString(StartRowIndex, BatchQty, Me.DataGridView2)
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

    Private Function GetInsertString(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, ByVal vdgvItem As DataGridView) As String
        Dim retval As String = ""
        Dim row As DataGridViewRow
        Dim Delimiter As String = ","

        Try
            retval = "INSERT INTO `tbltrans_custom_itemdetail` (`No`, `Category`, `Code No`, `Desc`, `Unit`, `Tgl`, `Tipe Doc`, " + _
            "`Tipe Item`, `Doc No`, `Qty BB`, `Amount BB`, `Qty In`, `Amount In`, `Qty Out`, `Amount Out`, `Qty End`, `Amount End`, " + _
            "`Tipe Item 2`, `Tipe Doc 2`, `Doc No 2`, `DefPrec`, `Primary Transaksi Asal`) VALUES"

            For i As Double = 0 To BatchQty - 1 Step 1
                If vdgvItem.Rows.Count > vStartRowIndex + i Then
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vdgvItem.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    row = vdgvItem.Rows(vStartRowIndex + i)

                    With row
                        If IsNothing(.Cells(Me.ColItem_TipeItem2.Index).Value) = True Then
                            .Cells(Me.ColItem_TipeItem2.Index).Value = ""
                        End If
                        If IsNothing(.Cells(Me.ColItem_TipeDoc2.Index).Value) = True Then
                            .Cells(Me.ColItem_TipeDoc2.Index).Value = ""
                        End If
                        If IsNothing(.Cells(Me.ColItem_DocNo2.Index).Value) = True Then
                            .Cells(Me.ColItem_DocNo2.Index).Value = ""
                        End If
                    End With

                    '         retval = "INSERT INTO `tbltrans_custom_itemdetail` (`No`, `Category`, `Code No`, `Desc`, `Unit`, `Tgl`, `Tipe Doc`, " + _
                    '"`Tipe Item`, `Doc No`, `Qty BB`, `Amount BB`, `Qty In`, `Amount In`, `Qty Out`, `Amount Out`, `Qty End`, `Amount End`, " + _
                    '"`Tipe Item 2`, `Tipe Doc 2`, `Doc No 2`, `DefPrec`, `Primary Transaksi Asal`) VALUES"

                    retval = retval + vbCrLf + "(" + row.Cells(Me.ColItem_No.Index).Value.ToString + "," _
                    + "'" + row.Cells(Me.ColItem_Category.Index).Value.ToString + "'," _
                    + "'" + row.Cells(Me.ColItem_CodeNo.Index).Value.ToString + "'," _
                    + "'" + Me.ObjCD.FormatSpecialChar_MySQL(row.Cells(Me.ColItem_Desc.Index).Value.ToString) + "'," _
                    + "'" + row.Cells(Me.ColItem_Unit.Index).Value.ToString + "'," _
                    + "'" + Format(row.Cells(Me.ColItem_Tgl.Index).Value, "yyyy-MM-dd") + "'," _
                    + "'" + row.Cells(Me.ColItem_TipeDoc.Index).Value.ToString + "'," _
                    + "'" + row.Cells(Me.ColItem_TipeItem.Index).Value.ToString + "'," _
                    + "'" + row.Cells(Me.ColItem_NoDoc.Index).Value.ToString + "'," _
                    + CDec(row.Cells(Me.ColItem_QtyBB.Index).Value).ToString + "," _
                    + CDec(row.Cells(Me.ColItem_AmountBB.Index).Value).ToString + "," _
                    + CDec(row.Cells(Me.ColItem_QtyIn.Index).Value).ToString + "," _
                    + CDec(row.Cells(Me.ColItem_AmountIn.Index).Value).ToString + "," _
                    + CDec(row.Cells(Me.ColItem_QtyOut.Index).Value).ToString + "," _
                    + CDec(row.Cells(Me.ColItem_AmountOut.Index).Value).ToString + "," _
                    + CDec(row.Cells(Me.ColItem_QtyEnd.Index).Value).ToString + "," _
                    + CDec(row.Cells(Me.ColItem_AmountEnd.Index).Value).ToString + "," _
                    + "'" + row.Cells(Me.ColItem_TipeItem2.Index).Value.ToString + "'," _
                    + "'" + row.Cells(Me.ColItem_TipeDoc2.Index).Value.ToString + "'," _
                    + "'" + row.Cells(Me.ColItem_DocNo2.Index).Value.ToString + "'," _
                    + CDbl(row.Cells(Me.ColDefPrec.Index).Value).ToString + "," _
                    + CDbl(row.Cells(Me.ColItem_PrimaryTransaksiAsal.Index).Value).ToString + "" _
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

    Private Sub btnLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoad.Click
        Dim rowDGV As DataGridViewRow
        Dim DTA As New Dse3TableAdapters.tbltrans_custom_itemdetailTableAdapter
        Dim TBL As New Dse3.tbltrans_custom_itemdetailDataTable
        Dim row As Dse3.tbltrans_custom_itemdetailRow
        Dim IntCur As Integer = Me.ObjCur.GetDecimal_DefaultCurrency

        Try
            If MsgBox("Load detail item?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                DTA.Fill(TBL)
                Me.DataGridView2.Rows.Clear()

                For Each row In TBL.Rows
                    Me.DataGridView2.Rows.Add()
                    rowDGV = Me.DataGridView2.Rows(Me.DataGridView2.Rows.Count - 1)
                    With rowDGV
                        .Cells(Me.ColItem_No.Index).Value = row.No
                        .Cells(Me.ColItem_Category.Index).Value = row.Category
                        .Cells(Me.ColCodeNo.Index).Value = row.Code_No
                        .Cells(Me.ColItem_Desc.Index).Value = row.Desc
                        .Cells(Me.ColItem_Unit.Index).Value = row.Unit
                        .Cells(Me.ColItem_Tgl.Index).Value = row.Tgl
                        .Cells(Me.ColItem_TipeDoc.Index).Value = row.Tipe_Doc
                        .Cells(Me.ColItem_TipeItem.Index).Value = row.Tipe_Item
                        .Cells(Me.ColItem_NoDoc.Index).Value = row.Doc_No
                        .Cells(Me.ColItem_QtyBB.Index).Value = Me.ObjCD.DigitGrouping(row.Qty_BB, row.DefPrec)
                        .Cells(Me.ColItem_AmountBB.Index).Value = Me.ObjCD.DigitGrouping(row.Amount_BB, IntCur)
                        .Cells(Me.ColItem_QtyIn.Index).Value = Me.ObjCD.DigitGrouping(row.Qty_In, row.DefPrec)
                        .Cells(Me.ColItem_AmountIn.Index).Value = Me.ObjCD.DigitGrouping(row.Amount_In, IntCur)
                        .Cells(Me.ColItem_QtyOut.Index).Value = Me.ObjCD.DigitGrouping(row.Qty_Out, row.DefPrec)
                        .Cells(Me.ColItem_AmountOut.Index).Value = Me.ObjCD.DigitGrouping(row.Amount_Out, IntCur)
                        .Cells(Me.ColItem_QtyEnd.Index).Value = Me.ObjCD.DigitGrouping(row.Qty_End, row.DefPrec)
                        .Cells(Me.ColItem_AmountEnd.Index).Value = Me.ObjCD.DigitGrouping(row.Amount_End, IntCur)

                        .Cells(Me.ColItem_TipeItem2.Index).Value = row.Tipe_Item_2
                        .Cells(Me.ColItem_TipeDoc2.Index).Value = row.Tipe_Doc_2
                        .Cells(Me.ColItem_DocNo2.Index).Value = row.Doc_No_2
                        .Cells(Me.ColItem_DefPrec.Index).Value = row.DefPrec
                        .Cells(Me.ColItem_PrimaryTransaksiAsal.Index).Value = row.Primary_Transaksi_Asal
                    End With
                Next

                MsgBox("Data loaded", MsgBoxStyle.Information, "Information")
            End If
            Me.ColItem_Tgl.DefaultCellStyle.Format = "dd MMM yyyy"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Private Sub btnCreateCSV_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreateCSV.Click
        Try
            'item detail
            If Me.DataGridView2.Rows.Count > 0 Then
                FrmDia.Title = "Please specify where to save csv file - item detail"
                FrmDia.Filter = "*.csv|csv file"
                FrmDia.FilterIndex = 0
                FrmDia.DefaultExt = ".csv"

                If FrmDia.ShowDialog <> DialogResult.OK Then
                    Exit Sub
                End If
                If FrmDia.FileName = "" Then
                    Exit Sub
                End If

                Me.CreateCsv_Item(FrmDia.FileName.Substring(0, FrmDia.FileName.Length - 4) + " - Item.csv")
                Me.CreateCsv_ItemDetail(FrmDia.FileName.Substring(0, FrmDia.FileName.Length - 4) + " - Item detail.csv")
                Me.CreateCsv_DetailWarehouse(FrmDia.FileName.Substring(0, FrmDia.FileName.Length - 4) + " - Detail warehouse.csv")
                Me.CreateCsv_DetailUsing(FrmDia.FileName.Substring(0, FrmDia.FileName.Length - 4) + " - Detail using.csv")

                MsgBox("Csv file created.", MsgBoxStyle.Information, "Information")
            Else
                MsgBox("Please specify at least 1 stock in purchase item.", _
                MsgBoxStyle.Information, "Information")
            End If

            'Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CreateCsv(ByVal vFileName As String)
        Try
            Me.CreateCsv_Item(vFileName.Substring(0, vFileName.Length - 4) + " - Item.csv")
            Me.CreateCsv_ItemDetail(vFileName.Substring(0, vFileName.Length - 4) + " - Item detail.csv")
            Me.CreateCsv_DetailWarehouse(vFileName.Substring(0, vFileName.Length - 4) + " - Detail warehouse.csv")
            Me.CreateCsv_DetailUsing(vFileName.Substring(0, vFileName.Length - 4) + " - Detail using.csv")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CreateCsv_Item(ByVal vFilePath As String)
        Dim rowDGV As DataGridViewRow
        Dim fileExists As Boolean
        Dim DummyString As String = ""
        Dim DummyIdx As Integer = 0

        Try
            If Me.DataGridView1.RowCount > 0 Then
                'hapus dummy file in program files folder
                fileExists = My.Computer.FileSystem.FileExists(vFilePath)
                If fileExists = True Then
                    My.Computer.FileSystem.DeleteFile(vFilePath, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
                End If
                'create empty file
                My.Computer.FileSystem.WriteAllText(vFilePath, String.Empty, False, System.Text.Encoding.ASCII)
                DummyString = "Category,Code No,Description,Unit,Qty Begining,Amount Begining,Qty In,Amount In,Qty Out,Amount Out,Qty End,Amount End"
                My.Computer.FileSystem.WriteAllText(vFilePath, DummyString + vbCrLf, True, System.Text.Encoding.ASCII)

                For Each rowDGV In Me.DataGridView1.Rows
                    DummyIdx += 1
                    DummyString = ""

                    DummyString = DummyString + rowDGV.Cells(Me.ColCategory.Index).Value.ToString.Replace(",", " ") + ","
                    DummyString = DummyString + rowDGV.Cells(Me.ColCodeNo.Index).Value.ToString.Replace(",", " ") + ","
                    DummyString = DummyString + rowDGV.Cells(Me.ColDesc.Index).Value.ToString.Replace(",", " ") + ","
                    DummyString = DummyString + rowDGV.Cells(Me.ColUnit.Index).Value.ToString + ","
                    DummyString = DummyString + Me.ObjCD.BuangNolDibelakangDecimal(CDec(rowDGV.Cells(Me.ColQtyBB.Index).Value).ToString) + ","
                    DummyString = DummyString + Me.ObjCD.BuangNolDibelakangDecimal(CDec(rowDGV.Cells(Me.ColAmountBB.Index).Value).ToString) + ","
                    DummyString = DummyString + Me.ObjCD.BuangNolDibelakangDecimal(CDec(rowDGV.Cells(Me.ColQtyIn.Index).Value).ToString) + ","
                    DummyString = DummyString + Me.ObjCD.BuangNolDibelakangDecimal(CDec(rowDGV.Cells(Me.ColAmountIn.Index).Value).ToString) + ","
                    DummyString = DummyString + Me.ObjCD.BuangNolDibelakangDecimal(CDec(rowDGV.Cells(Me.ColOutQty.Index).Value).ToString) + ","
                    DummyString = DummyString + Me.ObjCD.BuangNolDibelakangDecimal(CDec(rowDGV.Cells(Me.ColAmountOut.Index).Value).ToString) + ","
                    DummyString = DummyString + Me.ObjCD.BuangNolDibelakangDecimal(CDec(rowDGV.Cells(Me.ColQtyEnd.Index).Value).ToString) + ","
                    DummyString = DummyString + Me.ObjCD.BuangNolDibelakangDecimal(CDec(rowDGV.Cells(Me.ColAmountEnd.Index).Value).ToString)

                    If DummyIdx <> 1 Then
                        DummyString = vbCrLf + DummyString
                    End If

                    My.Computer.FileSystem.WriteAllText(vFilePath, DummyString, True, System.Text.Encoding.ASCII)
                Next
            End If


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CreateCsv_DetailUsing(ByVal vFilePath As String)
        Dim rowDGV As DataGridViewRow
        Dim fileExists As Boolean
        Dim DummyString As String = ""
        Dim DummyIdx As Integer = 0

        Try
            If Me.DataGridView4.RowCount > 0 Then
                'hapus dummy file in program files folder
                fileExists = My.Computer.FileSystem.FileExists(vFilePath)
                If fileExists = True Then
                    My.Computer.FileSystem.DeleteFile(vFilePath, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
                End If
                'create empty file
                My.Computer.FileSystem.WriteAllText(vFilePath, String.Empty, False, System.Text.Encoding.ASCII)
                DummyString = "Category,Code No,Description,Using Detail,Qty Using,Unit,Amount Using"
                My.Computer.FileSystem.WriteAllText(vFilePath, DummyString + vbCrLf, True, System.Text.Encoding.ASCII)

                For Each rowDGV In Me.DataGridView4.Rows
                    DummyIdx += 1
                    DummyString = ""

                    DummyString = DummyString + rowDGV.Cells(Me.ColU_Category.Index).Value.ToString.Replace(",", " ") + ","
                    DummyString = DummyString + rowDGV.Cells(Me.ColU_CodeNo.Index).Value.ToString.Replace(",", " ") + ","
                    DummyString = DummyString + rowDGV.Cells(Me.ColU_Description.Index).Value.ToString.Replace(",", " ") + ","
                    DummyString = DummyString + rowDGV.Cells(Me.ColU_UsingDetail.Index).Value.ToString + ","
                    DummyString = DummyString + Me.ObjCD.BuangNolDibelakangDecimal(CDec(rowDGV.Cells(Me.ColU_QtyUsing.Index).Value).ToString) + ","
                    DummyString = DummyString + rowDGV.Cells(Me.ColU_Unit.Index).Value.ToString + ","
                    DummyString = DummyString + Me.ObjCD.BuangNolDibelakangDecimal(CDec(rowDGV.Cells(Me.ColU_AmountEnd.Index).Value).ToString)

                    If DummyIdx <> 1 Then
                        DummyString = vbCrLf + DummyString
                    End If

                    My.Computer.FileSystem.WriteAllText(vFilePath, DummyString, True, System.Text.Encoding.ASCII)
                Next
            End If


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CreateCsv_DetailWarehouse(ByVal vFilePath As String)
        Dim rowDGV As DataGridViewRow
        Dim fileExists As Boolean
        Dim DummyString As String = ""
        Dim DummyIdx As Integer = 0

        Try
            If Me.DataGridView3.RowCount > 0 Then
                'hapus dummy file in program files folder
                fileExists = My.Computer.FileSystem.FileExists(vFilePath)
                If fileExists = True Then
                    My.Computer.FileSystem.DeleteFile(vFilePath, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
                End If
                'create empty file
                My.Computer.FileSystem.WriteAllText(vFilePath, String.Empty, False, System.Text.Encoding.ASCII)
                DummyString = "Category,Code No,Description,Warehouse,Qty End,Unit,Amount End"
                My.Computer.FileSystem.WriteAllText(vFilePath, DummyString + vbCrLf, True, System.Text.Encoding.ASCII)

                For Each rowDGV In Me.DataGridView3.Rows
                    DummyIdx += 1
                    DummyString = ""

                    DummyString = DummyString + rowDGV.Cells(Me.ColW_CategoryCode.Index).Value.ToString.Replace(",", " ") + ","
                    DummyString = DummyString + rowDGV.Cells(Me.ColW_CodeNo.Index).Value.ToString.Replace(",", " ") + ","
                    DummyString = DummyString + rowDGV.Cells(Me.ColW_Desc.Index).Value.ToString.Replace(",", " ") + ","
                    DummyString = DummyString + rowDGV.Cells(Me.ColW_Warehouse.Index).Value.ToString + ","
                    DummyString = DummyString + Me.ObjCD.BuangNolDibelakangDecimal(CDec(rowDGV.Cells(Me.ColW_QtyEnd.Index).Value).ToString) + ","
                    DummyString = DummyString + rowDGV.Cells(Me.ColW_UnitCode.Index).Value.ToString + ","
                    DummyString = DummyString + Me.ObjCD.BuangNolDibelakangDecimal(CDec(rowDGV.Cells(Me.ColW_AmountEnd.Index).Value).ToString)

                    If DummyIdx <> 1 Then
                        DummyString = vbCrLf + DummyString
                    End If

                    My.Computer.FileSystem.WriteAllText(vFilePath, DummyString, True, System.Text.Encoding.ASCII)
                Next
            End If


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CreateCsv_ItemDetail(ByVal vFilePath As String)
        Dim rowDGV As DataGridViewRow
        Dim fileExists As Boolean
        Dim DummyString As String = ""
        Dim DummyIdx As Integer = 0

        Try
            If Me.DataGridView2.RowCount > 0 Then
                'hapus dummy file in program files folder
                fileExists = My.Computer.FileSystem.FileExists(vFilePath)
                If fileExists = True Then
                    My.Computer.FileSystem.DeleteFile(vFilePath, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
                End If
                'create empty file
                My.Computer.FileSystem.WriteAllText(vFilePath, String.Empty, False, System.Text.Encoding.ASCII)
                DummyString = "No,Category,Code No,Description,Unit,Date,Tipe Doc,Tipe Item,Doc No,Qty Begining,Amount Begining,Qty In,Amount In,Qty Out,Amount Out,Qty End,Amount End,Tipe Item 2,Tipe Doc 2,Doc No 2" ',Def Prec,Primary Transaksi Asal"
                My.Computer.FileSystem.WriteAllText(vFilePath, DummyString + vbCrLf, True, System.Text.Encoding.ASCII)

                For Each rowDGV In Me.DataGridView2.Rows
                    DummyIdx += 1
                    DummyString = ""

                    DummyString = DummyString + rowDGV.Cells(Me.ColItem_No.Index).Value.ToString + ","
                    DummyString = DummyString + rowDGV.Cells(Me.ColItem_Category.Index).Value.ToString.Replace(",", " ") + ","
                    DummyString = DummyString + rowDGV.Cells(Me.ColItem_CodeNo.Index).Value.ToString.Replace(",", " ") + ","
                    DummyString = DummyString + rowDGV.Cells(Me.ColItem_Desc.Index).Value.ToString.Replace(",", " ") + ","
                    DummyString = DummyString + rowDGV.Cells(Me.ColItem_Unit.Index).Value.ToString.Replace(vbCrLf, ", ").Replace(",", " ") + ","
                    DummyString = DummyString + rowDGV.Cells(Me.ColItem_Tgl.Index).Value.ToString + ","
                    DummyString = DummyString + rowDGV.Cells(Me.ColItem_TipeDoc.Index).Value.ToString + ","
                    DummyString = DummyString + rowDGV.Cells(Me.ColItem_TipeItem.Index).Value.ToString.Replace(vbCrLf, ", ").Replace(",", " ") + ","
                    DummyString = DummyString + rowDGV.Cells(Me.ColItem_NoDoc.Index).Value.ToString.Replace(vbCrLf, ", ").Replace(",", " ") + ","
                    DummyString = DummyString + Me.ObjCD.BuangNolDibelakangDecimal(CDec(rowDGV.Cells(Me.ColItem_QtyBB.Index).Value).ToString) + ","
                    DummyString = DummyString + Me.ObjCD.BuangNolDibelakangDecimal(CDec(rowDGV.Cells(Me.ColItem_AmountBB.Index).Value).ToString) + ","
                    DummyString = DummyString + Me.ObjCD.BuangNolDibelakangDecimal(CDec(rowDGV.Cells(Me.ColItem_QtyIn.Index).Value).ToString) + ","
                    DummyString = DummyString + Me.ObjCD.BuangNolDibelakangDecimal(CDec(rowDGV.Cells(Me.ColItem_AmountIn.Index).Value).ToString) + ","
                    DummyString = DummyString + Me.ObjCD.BuangNolDibelakangDecimal(CDec(rowDGV.Cells(Me.ColItem_QtyOut.Index).Value).ToString) + ","
                    DummyString = DummyString + Me.ObjCD.BuangNolDibelakangDecimal(CDec(rowDGV.Cells(Me.ColItem_AmountOut.Index).Value).ToString) + ","
                    DummyString = DummyString + Me.ObjCD.BuangNolDibelakangDecimal(CDec(rowDGV.Cells(Me.ColItem_QtyEnd.Index).Value).ToString) + ","
                    DummyString = DummyString + Me.ObjCD.BuangNolDibelakangDecimal(CDec(rowDGV.Cells(Me.ColItem_AmountEnd.Index).Value).ToString) + ","
                    DummyString = DummyString + rowDGV.Cells(Me.ColItem_TipeItem2.Index).Value.ToString.Replace(",", " ") + ","
                    DummyString = DummyString + rowDGV.Cells(Me.ColItem_TipeDoc2.Index).Value.ToString.Replace(",", " ") + ","
                    DummyString = DummyString + rowDGV.Cells(Me.ColItem_DocNo2.Index).Value.ToString.Replace(",", " ") '+ ","
                    'DummyString = DummyString + rowDGV.Cells(Me.ColItem_DefPrec.Index).Value.ToString.Replace(",", " ") + ","
                    'DummyString = DummyString + rowDGV.Cells(Me.ColItem_PrimaryTransaksiAsal.Index).Value.ToString.Replace(",", " ")

                    If DummyIdx <> 1 Then
                        DummyString = vbCrLf + DummyString
                    End If

                    My.Computer.FileSystem.WriteAllText(vFilePath, DummyString, True, System.Text.Encoding.ASCII)
                Next
            End If


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    '|----------------------------------------|
    '|Ahdar '- Load & Save (6/14/2017)        |
    '|----------------------------------------|
    Private Sub LoadToDgv_Item(ByVal vPeriodClosing As Date, Optional ByVal vCodeNo As String = "")
        Dim DTA As New Dse2TableAdapters.tbltrans_custom_closing_itemTableAdapter
        Dim TBL As New Dse2.tbltrans_custom_closing_itemDataTable
        Dim row As Dse2.tbltrans_custom_closing_itemRow
        Dim rowDGV As New DataGridViewRow
        Dim Lanjut As Boolean = False

        Try
            Me.DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            DTA.Fill(TBL, vPeriodClosing)
            Me.DataGridView1.Rows.Clear()

            For Each row In TBL.Rows
                If vCodeNo = "" Then
                    Lanjut = True
                Else
                    If row.Code_No = vCodeNo Then
                        Lanjut = True
                    Else
                        Lanjut = False
                    End If
                End If
                If Lanjut = True Then
                    Me.DataGridView1.Rows.Add()
                    rowDGV = Me.DataGridView1.Rows(Me.DataGridView1.Rows.Count - 1)
                    With rowDGV
                        .Cells(Me.ColCategory.Index).Value = row.Category
                        .Cells(Me.ColCodeNo.Index).Value = row.Code_No
                        .Cells(Me.ColDesc.Index).Value = row.Desc
                        .Cells(Me.ColUnit.Index).Value = row.Unit
                        .Cells(Me.ColQtyBB.Index).Value = row.Qty_BB
                        .Cells(Me.ColAmountBB.Index).Value = row.Amount_BB
                        .Cells(Me.ColQtyIn.Index).Value = row.Qty_In
                        .Cells(Me.ColAmountIn.Index).Value = row.Amount_In
                        .Cells(Me.ColOutQty.Index).Value = row.Out_Qty
                        .Cells(Me.ColAmountOut.Index).Value = row.Amount_Out
                        .Cells(Me.ColQtyEnd.Index).Value = row.Qty_End
                        .Cells(Me.ColAmountEnd.Index).Value = row.Amount_End
                        .Cells(Me.ColDefPrec.Index).Value = row.Default_Precision
                    End With
                End If
            Next
            'Me.DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            ' Me.DataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
        End Try
    End Sub

    Private Sub LoadToDgv_DetailData(ByVal vPeriodClosing As Date, Optional ByVal vCodeNo As String = "")
        Dim DTA As New Dse2TableAdapters.tbltrans_custom_closing_detail_dataTableAdapter
        Dim TBL As New Dse2.tbltrans_custom_closing_detail_dataDataTable
        Dim row As Dse2.tbltrans_custom_closing_detail_dataRow
        Dim rowDGV As New DataGridViewRow
        Dim Lanjut As Boolean = True

        Try
            Me.DataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DataGridView2.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            DTA.Fill(TBL, vPeriodClosing)
            Me.DataGridView2.Rows.Clear()

            Me.IsikanDocNo(TBL)

            For Each row In TBL.Rows
                If vCodeNo = "" Then
                    Lanjut = True
                Else
                    If row.Code_No = vCodeNo Then
                        Lanjut = True
                    Else
                        Lanjut = False
                    End If
                End If
                If Lanjut = True Then
                    Me.DataGridView2.Rows.Add()
                    rowDGV = Me.DataGridView2.Rows(Me.DataGridView2.Rows.Count - 1)
                    With rowDGV
                        .Cells(Me.ColItem_No.Index).Value = row.No
                        .Cells(Me.ColItem_Category.Index).Value = row.Category
                        .Cells(Me.ColItem_CodeNo.Index).Value = row.Code_No
                        .Cells(Me.ColItem_Desc.Index).Value = row.Desc
                        .Cells(Me.ColItem_Unit.Index).Value = row.Unit
                        .Cells(Me.ColItem_Tgl.Index).Value = row.Tanggal
                        .Cells(Me.ColItem_TipeDoc.Index).Value = row.Tipe_Doc
                        .Cells(Me.ColItem_TipeItem.Index).Value = row.Tipe_Item
                        .Cells(Me.ColItem_NoDoc.Index).Value = row.Doc_No
                        .Cells(Me.ColItem_QtyBB.Index).Value = row.Qty_BB
                        .Cells(Me.ColItem_AmountBB.Index).Value = row.Amount_BB
                        .Cells(Me.ColItem_QtyIn.Index).Value = row.Qty_In
                        .Cells(Me.ColItem_AmountIn.Index).Value = row.Amount_In
                        .Cells(Me.ColItem_QtyOut.Index).Value = row.Out_Qty
                        .Cells(Me.ColItem_AmountOut.Index).Value = row.Amount_Out
                        .Cells(Me.ColItem_QtyEnd.Index).Value = row.Qty_End
                        .Cells(Me.ColItem_AmountEnd.Index).Value = row.Amount_End
                        .Cells(Me.ColItem_TipeItem2.Index).Value = row.Tipe_Item2
                        .Cells(Me.ColItem_TipeDoc2.Index).Value = row.Tipe_Doc2
                        .Cells(Me.ColItem_DocNo2.Index).Value = row.Doc_No2
                        .Cells(Me.ColItem_DefPrec.Index).Value = row.Default_Precision
                        .Cells(Me.ColItem_PrimaryTransaksiAsal.Index).Value = row.Primary_Transaksi_Asal
                    End With
                End If
            Next
            'Me.DataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            'Me.DataGridView2.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
        End Try
    End Sub

    Private Sub IsikanDocNo(ByVal TBL As Dse2.tbltrans_custom_closing_detail_dataDataTable)
        Dim DTAAOP As New Dse2TableAdapters.tbltrans_aopTableAdapter
        Dim TBLAOP As New Dse2.tbltrans_aopDataTable
        Dim rowAOP As Dse2.tbltrans_aopRow
        Dim PrimaryMax, PrimaryMin As Double
        Dim row As Dse2.tbltrans_custom_closing_detail_dataRow
        Dim StartIdx As Double = 0

        Try
            DTAAOP.Connection.ConnectionString = Core_Module.GetConstring
            'stock in purchase
            TBL.DefaultView.RowFilter = "[tipe transaksi asal]" + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InPurchase).ToString
            TBL.DefaultView.Sort = "[primary transaksi asal]"

            PrimaryMax = Double.MinValue
            PrimaryMin = Double.MaxValue
            If TBL.DefaultView.Count > 0 Then
                row = TBL.DefaultView(0).Row
                PrimaryMin = row.Primary_Transaksi_Asal
                row = TBL.DefaultView(TBL.DefaultView.Count - 1).Row
                PrimaryMax = row.Primary_Transaksi_Asal

                DTAAOP.Fill(TBLAOP, PrimaryMax, PrimaryMin)
                While StartIdx <= TBLAOP.Rows.Count - 1
                    rowAOP = TBLAOP.Rows(StartIdx)
                    If rowAOP.Primary < row.Primary_Transaksi_Asal Then
                        StartIdx += 1
                    ElseIf rowAOP.Primary = row.Primary_Transaksi_Asal Then
                        row.Doc_No = rowAOP.No_Transaksi
                        Exit While
                    Else
                        If StartIdx > 1 Then
                            StartIdx -= 1
                        End If
                        Exit While
                    End If
                End While
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAAOP.Dispose()
            TBLAOP.Dispose()
        End Try
    End Sub

    Private Sub IsikanDocNo(ByVal TBL As Dse3.tbltrans_purchase_closing_item_detailDataTable)
        Dim row As Dse3.tbltrans_purchase_closing_item_detailRow
        Try
            For Each row In TBL.Rows
                row.doc_no = ""
            Next

            Me.IsikanDocNo_AOP(TBL)
            Me.IsikanDocNo_StockBundling(TBL)
            Me.IsikanDocNo_StockInOutAdjust(TBL)
            Me.IsikanDocNo_Invoice(TBL)
            Me.IsikanDocNo_PurchaseReturn(TBL)
            Me.IsikanDocNo_SalesReturn(TBL)
            Me.IsikanDocNo_StockTaking(TBL)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanDocNo_StockTaking(ByVal TBL As Dse3.tbltrans_purchase_closing_item_detailDataTable)
        Dim DTASR As New Dse2TableAdapters.tbltrans_stock_takingTableAdapter
        Dim TBLSR As New Dse2.tbltrans_stock_takingDataTable
        Dim rowSR As Dse2.tbltrans_stock_takingRow
        Dim PrimaryMax, PrimaryMin As Double
        Dim row As Dse3.tbltrans_purchase_closing_item_detailRow
        Dim StartIdx As Double = 0

        Try
            DTASR.Connection.ConnectionString = Core_Module.GetConstring
            'stock in purchase
            TBL.DefaultView.RowFilter = "[tipe transaksi asal] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InStockTakeAdjust).ToString + " OR " + _
            "[tipe transaksi asal] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutStockTakeAdjust).ToString
            TBL.DefaultView.Sort = "[primary transaksi asal]"

            PrimaryMax = Double.MinValue
            PrimaryMin = Double.MaxValue
            If TBL.DefaultView.Count > 0 Then
                row = TBL.DefaultView(0).Row
                PrimaryMin = row.primary_transaksi_asal
                row = TBL.DefaultView(TBL.DefaultView.Count - 1).Row
                PrimaryMax = row.primary_transaksi_asal

                DTASR.Fill(TBLSR, PrimaryMax, PrimaryMin)
                For Each DV As DataRowView In TBL.DefaultView
                    row = DV.Row
                    row.doc_no = ""

                    While StartIdx <= TBLSR.Rows.Count - 1
                        rowSR = TBLSR.Rows(StartIdx)
                        If rowSR.Primary < row.primary_transaksi_asal Then
                            StartIdx += 1
                        ElseIf rowSR.Primary = row.primary_transaksi_asal Then
                            row.doc_no = rowSR.No_Transaksi
                            Exit While
                        Else
                            If StartIdx > 1 Then
                                StartIdx -= 1
                            End If
                            Exit While
                        End If
                    End While
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTASR.Dispose()
            TBLSR.Dispose()
        End Try
    End Sub

    Private Sub IsikanDocNo_SalesReturn(ByVal TBL As Dse3.tbltrans_purchase_closing_item_detailDataTable)
        Dim DTASR As New Dse2TableAdapters.DataTable6TableAdapter
        Dim TBLSR As New Dse2.DataTable6DataTable
        Dim rowSR As Dse2.DataTable6Row
        Dim PrimaryMax, PrimaryMin As Double
        Dim row As Dse3.tbltrans_purchase_closing_item_detailRow
        Dim StartIdx As Double = 0

        Try
            DTASR.Connection.ConnectionString = Core_Module.GetConstring
            'stock in purchase
            TBL.DefaultView.RowFilter = "[tipe transaksi asal] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InSR).ToString
            TBL.DefaultView.Sort = "[primary transaksi asal]"

            PrimaryMax = Double.MinValue
            PrimaryMin = Double.MaxValue
            If TBL.DefaultView.Count > 0 Then
                row = TBL.DefaultView(0).Row
                PrimaryMin = row.primary_transaksi_asal
                row = TBL.DefaultView(TBL.DefaultView.Count - 1).Row
                PrimaryMax = row.primary_transaksi_asal

                DTASR.Fill(TBLSR, PrimaryMax, PrimaryMin, CInt(Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn))
                For Each DV As DataRowView In TBL.DefaultView
                    row = DV.Row
                    row.doc_no = ""

                    While StartIdx <= TBLSR.Rows.Count - 1
                        rowSR = TBLSR.Rows(StartIdx)
                        If rowSR.Primary < row.primary_transaksi_asal Then
                            StartIdx += 1
                        ElseIf rowSR.Primary = row.primary_transaksi_asal Then
                            row.doc_no = rowSR.No_Transaksi
                            Exit While
                        Else
                            If StartIdx > 1 Then
                                StartIdx -= 1
                            End If
                            Exit While
                        End If
                    End While
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTASR.Dispose()
            TBLSR.Dispose()
        End Try
    End Sub

    Private Sub IsikanDocNo_PurchaseReturn(ByVal TBL As Dse3.tbltrans_purchase_closing_item_detailDataTable)
        Dim DTAPR As New Dse2TableAdapters.tbltrans_purchasereturnTableAdapter
        Dim TBLPR As New Dse2.tbltrans_purchasereturnDataTable
        Dim rowPR As Dse2.tbltrans_purchasereturnRow
        Dim PrimaryMax, PrimaryMin As Double
        Dim row As Dse3.tbltrans_purchase_closing_item_detailRow
        Dim StartIdx As Double = 0

        Try
            DTAPR.Connection.ConnectionString = Core_Module.GetConstring
            'stock in purchase
            TBL.DefaultView.RowFilter = "[tipe transaksi asal] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutPR).ToString
            TBL.DefaultView.Sort = "[primary transaksi asal]"

            PrimaryMax = Double.MinValue
            PrimaryMin = Double.MaxValue
            If TBL.DefaultView.Count > 0 Then
                row = TBL.DefaultView(0).Row
                PrimaryMin = row.primary_transaksi_asal
                row = TBL.DefaultView(TBL.DefaultView.Count - 1).Row
                PrimaryMax = row.primary_transaksi_asal

                DTAPR.Fill(TBLPR, PrimaryMax, PrimaryMin)
                For Each DV As DataRowView In TBL.DefaultView
                    row = DV.Row
                    row.doc_no = ""

                    While StartIdx <= TBLPR.Rows.Count - 1
                        rowPR = TBLPR.Rows(StartIdx)
                        If rowPR.Primary < row.primary_transaksi_asal Then
                            StartIdx += 1
                        ElseIf rowPR.Primary = row.primary_transaksi_asal Then
                            row.doc_no = rowPR.No_Transaksi
                            Exit While
                        Else
                            If StartIdx > 1 Then
                                StartIdx -= 1
                            End If
                            Exit While
                        End If
                    End While
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAPR.Dispose()
            TBLPR.Dispose()
        End Try
    End Sub

    Private Sub IsikanDocNo_Invoice(ByVal TBL As Dse3.tbltrans_purchase_closing_item_detailDataTable)
        Dim DTAInvoice As New Dse2TableAdapters.tbltrans_salesinvoiceTableAdapter
        Dim TBLInvoice As New Dse2.tbltrans_salesinvoiceDataTable
        Dim rowInvoice As Dse2.tbltrans_salesinvoiceRow
        Dim PrimaryMax, PrimaryMin As Double
        Dim row As Dse3.tbltrans_purchase_closing_item_detailRow
        Dim StartIdx As Double = 0

        Try
            DTAInvoice.Connection.ConnectionString = Core_Module.GetConstring
            'stock in purchase
            TBL.DefaultView.RowFilter = "[tipe transaksi asal] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutSales).ToString  
            TBL.DefaultView.Sort = "[primary transaksi asal]"

            PrimaryMax = Double.MinValue
            PrimaryMin = Double.MaxValue
            If TBL.DefaultView.Count > 0 Then
                row = TBL.DefaultView(0).Row
                PrimaryMin = row.primary_transaksi_asal
                row = TBL.DefaultView(TBL.DefaultView.Count - 1).Row
                PrimaryMax = row.primary_transaksi_asal

                DTAInvoice.Fill(TBLInvoice, PrimaryMax, PrimaryMin)
                For Each DV As DataRowView In TBL.DefaultView
                    row = DV.Row
                    row.doc_no = ""

                    While StartIdx <= TBLInvoice.Rows.Count - 1
                        rowInvoice = TBLInvoice.Rows(StartIdx)
                        If rowInvoice.Primary < row.primary_transaksi_asal Then
                            StartIdx += 1
                        ElseIf rowInvoice.Primary = row.primary_transaksi_asal Then
                            row.doc_no = rowInvoice.No_Transaksi
                            Exit While
                        Else
                            If StartIdx > 1 Then
                                StartIdx -= 1
                            End If
                            Exit While
                        End If
                    End While
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAInvoice.Dispose()
            TBLInvoice.Dispose()
        End Try
    End Sub

    Private Sub IsikanDocNo_StockBundling(ByVal TBL As Dse3.tbltrans_purchase_closing_item_detailDataTable)
        Dim DTAbundling As New Dse2TableAdapters.tbltransaksi_stockbundlingTableAdapter
        Dim TBLbundling As New Dse2.tbltransaksi_stockbundlingDataTable
        Dim rowbundling As Dse2.tbltransaksi_stockbundlingRow
        Dim PrimaryMax, PrimaryMin As Double
        Dim row As Dse3.tbltrans_purchase_closing_item_detailRow
        Dim StartIdx As Double = 0

        Try
            DTAbundling.Connection.ConnectionString = Core_Module.GetConstring
            'stock in purchase
            TBL.DefaultView.RowFilter = "[tipe transaksi asal] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InBundling).ToString + " OR " + _
            "[tipe transaksi asal] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutBundling).ToString
            TBL.DefaultView.Sort = "[primary transaksi asal]"

            PrimaryMax = Double.MinValue
            PrimaryMin = Double.MaxValue
            If TBL.DefaultView.Count > 0 Then
                row = TBL.DefaultView(0).Row
                PrimaryMin = row.primary_transaksi_asal
                row = TBL.DefaultView(TBL.DefaultView.Count - 1).Row
                PrimaryMax = row.primary_transaksi_asal

                DTAbundling.Fill(TBLbundling, PrimaryMax, PrimaryMin)
                For Each DV As DataRowView In TBL.DefaultView
                    row = DV.Row
                    row.doc_no = ""

                    While StartIdx <= TBLbundling.Rows.Count - 1
                        rowbundling = TBLbundling.Rows(StartIdx)
                        If rowbundling.Primary < row.primary_transaksi_asal Then
                            StartIdx += 1
                        ElseIf rowbundling.Primary = row.primary_transaksi_asal Then
                            row.doc_no = rowbundling.No_Transaksi
                            Exit While
                        Else
                            If StartIdx > 1 Then
                                StartIdx -= 1
                            End If
                            Exit While
                        End If
                    End While
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAbundling.Dispose()
            TBLbundling.Dispose()
        End Try
    End Sub

    Private Sub IsikanDocNo_StockInOutAdjust(ByVal TBL As Dse3.tbltrans_purchase_closing_item_detailDataTable)
        Dim DTAInOut As New Dse2TableAdapters.tbltransaksi_stockinoutTableAdapter
        Dim TBLInOut As New Dse2.tbltransaksi_stockinoutDataTable
        Dim rowInOut As Dse2.tbltransaksi_stockinoutRow
        Dim PrimaryMax, PrimaryMin As Double
        Dim row As Dse3.tbltrans_purchase_closing_item_detailRow
        Dim StartIdx As Double = 0

        Try
            DTAInOut.Connection.ConnectionString = Core_Module.GetConstring
            'stock in purchase
            TBL.DefaultView.RowFilter = "[tipe transaksi asal] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InAdjust).ToString + " OR " + _
            "[tipe transaksi asal] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.OutAdjust).ToString
            TBL.DefaultView.Sort = "[primary transaksi asal]"

            PrimaryMax = Double.MinValue
            PrimaryMin = Double.MaxValue
            If TBL.DefaultView.Count > 0 Then
                row = TBL.DefaultView(0).Row
                PrimaryMin = row.primary_transaksi_asal
                row = TBL.DefaultView(TBL.DefaultView.Count - 1).Row
                PrimaryMax = row.primary_transaksi_asal

                DTAInOut.Fill(TBLInOut, PrimaryMax, PrimaryMin)
                For Each DV As DataRowView In TBL.DefaultView
                    row = DV.Row
                    row.doc_no = ""

                    While StartIdx <= TBLInOut.Rows.Count - 1
                        rowInOut = TBLInOut.Rows(StartIdx)
                        If rowInOut.Primary < row.primary_transaksi_asal Then
                            StartIdx += 1
                        ElseIf rowInOut.Primary = row.primary_transaksi_asal Then
                            row.doc_no = rowInOut.No_Transaksi
                            Exit While
                        Else
                            If StartIdx > 1 Then
                                StartIdx -= 1
                            End If
                            Exit While
                        End If
                    End While
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAInOut.Dispose()
            TBLInOut.Dispose()
        End Try
    End Sub

    Private Sub IsikanDocNo_AOP(ByVal TBL As Dse3.tbltrans_purchase_closing_item_detailDataTable)
        Dim DTAAOP As New Dse2TableAdapters.tbltrans_aopTableAdapter
        Dim TBLAOP As New Dse2.tbltrans_aopDataTable
        Dim rowAOP As Dse2.tbltrans_aopRow
        Dim PrimaryMax, PrimaryMin As Double
        Dim row As Dse3.tbltrans_purchase_closing_item_detailRow
        Dim StartIdx As Double = 0

        Try
            DTAAOP.Connection.ConnectionString = Core_Module.GetConstring
            'stock in purchase
            TBL.DefaultView.RowFilter = "[tipe transaksi asal] = " + CInt(ClsTransaksiMonthlyClosingItem.EnumTipeTranskasiAsal.InPurchase).ToString
            TBL.DefaultView.Sort = "[primary transaksi asal]"

            PrimaryMax = Double.MinValue
            PrimaryMin = Double.MaxValue
            If TBL.DefaultView.Count > 0 Then
                row = TBL.DefaultView(0).Row
                PrimaryMin = row.primary_transaksi_asal
                row = TBL.DefaultView(TBL.DefaultView.Count - 1).Row
                PrimaryMax = row.primary_transaksi_asal

                DTAAOP.Fill(TBLAOP, PrimaryMax, PrimaryMin)
                For Each DV As DataRowView In TBL.DefaultView
                    row = DV.Row
                    row.doc_no = ""

                    While StartIdx <= TBLAOP.Rows.Count - 1
                        rowAOP = TBLAOP.Rows(StartIdx)
                        If rowAOP.Primary < row.primary_transaksi_asal Then
                            StartIdx += 1
                        ElseIf rowAOP.Primary = row.primary_transaksi_asal Then
                            row.doc_no = rowAOP.No_Transaksi
                            Exit While
                        Else
                            If StartIdx > 1 Then
                                StartIdx -= 1
                            End If
                            Exit While
                        End If
                    End While
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAAOP.Dispose()
            TBLAOP.Dispose()
        End Try
    End Sub

    Private Sub LoadToDgv_DetailWareHouse(ByVal vPeriodClosing As Date, Optional ByVal vCodeNo As String = "")
        Dim DTA As New Dse2TableAdapters.tbltrans_custom_closing_detail_warehouseTableAdapter
        Dim TBL As New Dse2.tbltrans_custom_closing_detail_warehouseDataTable
        Dim row As Dse2.tbltrans_custom_closing_detail_warehouseRow
        Dim rowDGV As New DataGridViewRow
        Dim Lanjut As Boolean = True
        Try
            Me.DataGridView3.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DataGridView3.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            DTA.Fill(TBL, vPeriodClosing)
            Me.DataGridView3.Rows.Clear()

            For Each row In TBL.Rows
                If vCodeNo = "" Then
                    Lanjut = True
                Else
                    If row.Code_No = vCodeNo Then
                        Lanjut = True
                    Else
                        Lanjut = False
                    End If
                End If
                If Lanjut = True Then
                    Me.DataGridView3.Rows.Add()
                    rowDGV = Me.DataGridView3.Rows(Me.DataGridView3.Rows.Count - 1)
                    With rowDGV
                        .Cells(Me.ColW_CategoryCode.Index).Value = row.Category
                        .Cells(Me.ColW_CodeNo.Index).Value = row.Code_No
                        .Cells(Me.ColW_Desc.Index).Value = row.Desc
                        .Cells(Me.ColW_Warehouse.Index).Value = row.Warehouse
                        .Cells(Me.ColW_QtyEnd.Index).Value = row.Qty_End
                        .Cells(Me.ColW_UnitCode.Index).Value = row.Unit
                        .Cells(Me.ColW_AmountEnd.Index).Value = row.Amount_End
                    End With
                End If
            Next
            'Me.DataGridView3.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            'Me.DataGridView3.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
        End Try
    End Sub

    Private Sub LoadToDgv_DetailUsing(ByVal vPeriodClosing As Date, Optional ByVal vCodeNo As String = "")
        Dim DTA As New Dse2TableAdapters.tbltrans_custom_closing_detail_usingTableAdapter
        Dim TBL As New Dse2.tbltrans_custom_closing_detail_usingDataTable
        Dim row As Dse2.tbltrans_custom_closing_detail_usingRow
        Dim rowDGV As New DataGridViewRow
        Dim Lanjut As Boolean = False
        Try
            Me.DataGridView4.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DataGridView4.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            DTA.Fill(TBL, vPeriodClosing)
            Me.DataGridView4.Rows.Clear()

            For Each row In TBL.Rows
                If vCodeNo = "" Then
                    Lanjut = True
                Else
                    If row.Code_No = vCodeNo Then
                        Lanjut = True
                    Else
                        Lanjut = False
                    End If
                End If
                If Lanjut = True Then
                    Me.DataGridView4.Rows.Add()
                    rowDGV = Me.DataGridView4.Rows(Me.DataGridView4.Rows.Count - 1)
                    With rowDGV
                        .Cells(Me.ColU_Category.Index).Value = row.Category
                        .Cells(Me.ColU_CodeNo.Index).Value = row.Code_No
                        .Cells(Me.ColU_Description.Index).Value = row.Desc
                        .Cells(Me.ColU_UsingDetail.Index).Value = row.Using_Detail
                        .Cells(Me.ColU_QtyUsing.Index).Value = row.Qty_Using
                        .Cells(Me.ColU_Unit.Index).Value = row.Unit
                        .Cells(Me.ColU_AmountEnd.Index).Value = row.Amount_End
                    End With
                End If
            Next
            'Me.DataGridView4.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            'Me.DataGridView4.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
        End Try
    End Sub

    Private Function Simpan_UsingScript_Item(ByVal vClosingPeriod As Date) As Boolean
        Dim conStr As String = Core_Module.GetConstring
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < Me.DataGridView1.Rows.Count
                sql = Me.GetInsertString_Item(StartRowIndex, BatchQty, Me.DataGridView1, vClosingPeriod)
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

    Private Function GetInsertString_Item(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, ByVal vdgvItem As DataGridView, ByVal vClosingPeriod As Date) As String
        Dim retval As String = ""
        Dim row As DataGridViewRow
        Dim Delimiter As String = ","
        Dim StartRowIndex As Double = 0
        Try
            retval = "INSERT INTO `tbltrans_custom_closing_item` (`Category`, `Code No`, `Desc`, `Unit`, `Qty BB`, `Amount BB`, " + _
            "`Qty In`, `Amount In`, `Out Qty`, `Amount Out`, `Qty End`, `Amount End`, `Default Precision`, `Closing Period`) VALUES"

            For i As Double = 0 To BatchQty - 1 Step 1
                If vdgvItem.Rows.Count > vStartRowIndex + i Then
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vdgvItem.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    row = vdgvItem.Rows(vStartRowIndex + i)

                    retval = retval + vbCrLf + "('" + row.Cells(Me.ColItem_Category.Index).Value.ToString + "'," _
                    + "'" + row.Cells(Me.ColCodeNo.Index).Value.ToString + "'," _
                    + "'" + Me.ObjCD.FormatSpecialChar_MySQL(row.Cells(Me.ColDesc.Index).Value.ToString) + "'," _
                    + "'" + row.Cells(Me.ColUnit.Index).Value.ToString + "'," _
                    + "'" + row.Cells(Me.ColQtyBB.Index).Value.ToString + "'," _
                    + "'" + row.Cells(Me.ColAmountBB.Index).Value.ToString + "'," _
                    + "'" + row.Cells(Me.ColQtyIn.Index).Value.ToString + "'," _
                    + "'" + row.Cells(Me.ColAmountIn.Index).Value.ToString + "'," _
                    + "'" + row.Cells(Me.ColOutQty.Index).Value.ToString + "'," _
                    + "'" + row.Cells(Me.ColAmountOut.Index).Value.ToString + "'," _
                    + "'" + row.Cells(Me.ColQtyEnd.Index).Value.ToString + "'," _
                    + "'" + row.Cells(Me.ColAmountEnd.Index).Value.ToString + "'," _
                    + "'" + row.Cells(Me.ColDefPrec.Index).Value.ToString + "'," _
                    + "'" + Format(vClosingPeriod, "yyyy-MM-dd") + "'" _
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

    Private Function Simpan_UsingScript_DetailData(ByVal vClosingPeriod As Date) As Boolean
        Dim conStr As String = Core_Module.GetConstring
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 100
        Dim StartRowIndex As Double = 0

        Try
            'Me.DataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.DataGridView2.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            'open connection
            Conn.Open()

            While StartRowIndex < Me.DataGridView2.Rows.Count
                sql = Me.GetInsertString_DetailData(StartRowIndex, BatchQty, Me.DataGridView2, vClosingPeriod)
                script.Query = sql
                'execute script
                script.Execute()
                StartRowIndex += BatchQty
            End While

            Conn.Close()

            'Me.DataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Private Function GetInsertString_DetailData(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, ByVal vdgvItem As DataGridView, ByVal vClosingPeriod As Date) As String
        Dim retval As String = ""
        Dim row As DataGridViewRow
        Dim Delimiter As String = ","
        Dim StartRowIndex As Double = 0
        Try
            Me.DataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            retval = "INSERT INTO `tbltrans_custom_closing_detail_data` (`No`, `Category`, `Code No`, `Desc`, `Unit`, `Tanggal`, `Tipe Doc`, " + _
         "`Tipe Item`, `Doc No`, `Qty BB`, `Amount BB`, `Qty In`, `Amount In`, `Out Qty`, `Amount Out`, `Qty End`, `Amount End`, " + _
         "`Tipe Item2`, `Tipe Doc2`, `Doc No2`, `Default Precision`, `Primary Transaksi Asal`, `Closing Period`) VALUES"

            For i As Double = 0 To BatchQty - 1 Step 1
                If vdgvItem.Rows.Count > vStartRowIndex + i Then
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vdgvItem.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    row = vdgvItem.Rows(vStartRowIndex + i)
                    'Me.DataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

                    With row
                        If IsNothing(.Cells(Me.ColItem_TipeItem2.Index).Value) = True Then
                            .Cells(Me.ColItem_TipeItem2.Index).Value = ""
                        End If
                        If IsNothing(.Cells(Me.ColItem_TipeDoc2.Index).Value) = True Then
                            .Cells(Me.ColItem_TipeDoc2.Index).Value = ""
                        End If
                        If IsNothing(.Cells(Me.ColItem_DocNo2.Index).Value) = True Then
                            .Cells(Me.ColItem_DocNo2.Index).Value = ""
                        End If
                    End With

                    If row.Cells(Me.ColItem_TipeDoc2.Index).Value.ToString.Length > 250 Then
                        row.Cells(Me.ColItem_TipeDoc2.Index).Value = row.Cells(Me.ColItem_TipeDoc2.Index).Value.ToString.Substring(0, 250)
                    End If
                    If row.Cells(Me.ColItem_DocNo2.Index).Value.ToString.Length > 250 Then
                        row.Cells(Me.ColItem_DocNo2.Index).Value = row.Cells(Me.ColItem_DocNo2.Index).Value.ToString.Substring(0, 250)
                    End If
                    retval = retval + vbCrLf + "(" + row.Cells(Me.ColItem_No.Index).Value.ToString + "," _
                    + "'" + row.Cells(Me.ColItem_Category.Index).Value.ToString + "'," _
                    + "'" + row.Cells(Me.ColItem_CodeNo.Index).Value.ToString + "'," _
                    + "'" + ObjCD.FormatSpecialChar_MySQL(row.Cells(Me.ColItem_Desc.Index).Value.ToString) + "'," _
                    + "'" + row.Cells(Me.ColItem_Unit.Index).Value.ToString + "'," _
                    + "'" + row.Cells(Me.ColItem_Tgl.Index).Value + "'," _
                    + "'" + row.Cells(Me.ColItem_TipeDoc.Index).Value.ToString + "'," _
                    + "'" + row.Cells(Me.ColItem_TipeItem.Index).Value.ToString + "'," _
                    + "'" + row.Cells(Me.ColItem_NoDoc.Index).Value.ToString + "'," _
                    + "'" + row.Cells(Me.ColItem_QtyBB.Index).Value.ToString + "'," _
                    + "'" + row.Cells(Me.ColItem_AmountBB.Index).Value.ToString + "'," _
                    + "'" + row.Cells(Me.ColItem_QtyIn.Index).Value.ToString + "'," _
                    + "'" + row.Cells(Me.ColItem_AmountIn.Index).Value.ToString + "'," _
                    + "'" + row.Cells(Me.ColItem_QtyOut.Index).Value.ToString + "'," _
                    + "'" + row.Cells(Me.ColItem_AmountOut.Index).Value.ToString + "'," _
                    + "'" + row.Cells(Me.ColItem_QtyEnd.Index).Value.ToString + "'," _
                    + "'" + row.Cells(Me.ColItem_AmountEnd.Index).Value.ToString + "'," _
                    + "'" + row.Cells(Me.ColItem_TipeItem2.Index).Value.ToString + "'," _
                    + "'" + row.Cells(Me.ColItem_TipeDoc2.Index).Value.ToString + "'," _
                    + "'" + row.Cells(Me.ColItem_DocNo2.Index).Value.ToString + "'," _
                    + "'" + row.Cells(Me.ColItem_DefPrec.Index).Value.ToString + "'," _
                    + "'" + row.Cells(Me.ColItem_PrimaryTransaksiAsal.Index).Value.ToString + "'," _
                    + "'" + Format(vClosingPeriod, "yyyy-MM-dd") + "'" _
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

    Private Function Simpan_UsingScript_DetailWarehouse(ByVal vClosingPeriod As Date) As Boolean
        Dim conStr As String = Core_Module.GetConstring
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < Me.DataGridView3.Rows.Count
                sql = Me.GetInsertString_DetailWarehouse(StartRowIndex, BatchQty, Me.DataGridView3, vClosingPeriod)
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

    Private Function GetInsertString_DetailWarehouse(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, ByVal vdgvItem As DataGridView, ByVal vClosingPeriod As Date) As String
        Dim retval As String = ""
        Dim row As DataGridViewRow
        Dim Delimiter As String = ","
        Dim StartRowIndex As Double = 0
        Try
            retval = "INSERT INTO `tbltrans_custom_closing_detail_warehouse` (`Category`, `Code No`, `Desc`, `Warehouse`, `Qty End`, `Unit`, " + _
         "`Amount End`, `Closing Period`) VALUES"

            For i As Double = 0 To BatchQty - 1 Step 1
                If vdgvItem.Rows.Count > vStartRowIndex + i Then
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vdgvItem.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    row = vdgvItem.Rows(vStartRowIndex + i)

                    retval = retval + vbCrLf + "('" + row.Cells(Me.ColW_CategoryCode.Index).Value.ToString + "'," _
                    + "'" + row.Cells(Me.ColW_CodeNo.Index).Value.ToString + "'," _
                    + "'" + Me.ObjCD.FormatSpecialChar_MySQL(row.Cells(Me.ColW_Desc.Index).Value.ToString) + "'," _
                    + "'" + row.Cells(Me.ColW_Warehouse.Index).Value.ToString + "'," _
                    + "'" + row.Cells(Me.ColW_QtyEnd.Index).Value.ToString + "'," _
                    + "'" + row.Cells(Me.ColW_UnitCode.Index).Value.ToString + "'," _
                    + "'" + row.Cells(Me.ColW_AmountEnd.Index).Value.ToString + "'," _
                    + "'" + Format(vClosingPeriod, "yyyy-MM-dd") + "'" _
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

    Private Function Simpan_UsingScript_DetailUsing(ByVal vClosingPeriod As Date) As Boolean
        Dim conStr As String = Core_Module.GetConstring
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < Me.DataGridView4.Rows.Count
                sql = Me.GetInsertString_DetailUsing(StartRowIndex, BatchQty, Me.DataGridView4, vClosingPeriod)
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

    Private Function GetInsertString_DetailUsing(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, ByVal vdgvItem As DataGridView, ByVal vClosingPeriod As Date) As String
        Dim retval As String = ""
        Dim row As DataGridViewRow
        Dim Delimiter As String = ","
        Dim StartRowIndex As Double = 0
        Try
            retval = "INSERT INTO `tbltrans_custom_closing_detail_using` (`Category`, `Code No`, `Desc`, `Using Detail`, `Qty Using`, `Unit`, " + _
         "`Amount End`, `Closing Period`) VALUES"

            For i As Double = 0 To BatchQty - 1 Step 1
                If vdgvItem.Rows.Count > vStartRowIndex + i Then
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vdgvItem.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    row = vdgvItem.Rows(vStartRowIndex + i)

                    retval = retval + vbCrLf + "('" + row.Cells(Me.ColU_Category.Index).Value.ToString + "'," _
                    + "'" + row.Cells(Me.ColU_CodeNo.Index).Value.ToString + "'," _
                     + "'" + Me.ObjCD.FormatSpecialChar_MySQL(row.Cells(Me.ColU_Description.Index).Value.ToString) + "'," _
                    + "'" + row.Cells(Me.ColU_UsingDetail.Index).Value.ToString + "'," _
                    + "'" + row.Cells(Me.ColU_QtyUsing.Index).Value.ToString + "'," _
                    + "'" + row.Cells(Me.ColU_Unit.Index).Value.ToString + "'," _
                    + "'" + row.Cells(Me.ColU_AmountEnd.Index).Value.ToString + "'," _
                    + "'" + Format(vClosingPeriod, "yyyy-MM-dd") + "'" _
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

    Private Sub DeleteAllItem(ByVal vClosingPeriod As Date)
        Dim DTAItem As New Dse2TableAdapters.tbltrans_custom_closing_itemTableAdapter
        Dim DTADetailData As New Dse2TableAdapters.tbltrans_custom_closing_detail_dataTableAdapter
        Dim DTADetailWarehouse As New Dse2TableAdapters.tbltrans_custom_closing_detail_warehouseTableAdapter
        Dim DTADetailUsing As New Dse2TableAdapters.tbltrans_custom_closing_detail_usingTableAdapter
        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            DTADetailData.Connection.ConnectionString = DTAItem.Connection.ConnectionString
            DTADetailWarehouse.Connection.ConnectionString = DTAItem.Connection.ConnectionString
            DTADetailUsing.Connection.ConnectionString = DTAItem.Connection.ConnectionString

            DTAItem.DeleteQuery(vClosingPeriod)
            DTADetailData.DeleteQuery(vClosingPeriod)
            DTADetailWarehouse.DeleteQuery(vClosingPeriod)
            DTADetailUsing.DeleteQuery(vClosingPeriod)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            DTADetailData.Dispose()
            DTADetailData.Dispose()
            DTADetailUsing.Dispose()
        End Try
    End Sub

    Private Sub btnSave1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave1.Click
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            If MsgBox("Save detail item?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                DeleteAllItem(Me.PeriodClosing)

                Me.Simpan_UsingScript_Item(Me.PeriodClosing)
                Me.Simpan_UsingScript_DetailData(Me.PeriodClosing)
                Me.Simpan_UsingScript_DetailWarehouse(Me.PeriodClosing)
                Me.Simpan_UsingScript_DetailUsing(Me.PeriodClosing)
                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub BtnLoad1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnLoad1.Click
        Dim PrimaryClosing As Double = -1

        Try
            PrimaryClosing = Me.GetPrimaryClosing
            If PrimaryClosing <> -1 Then
                If MsgBox("Load data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                    PrimaryTransaksiAcuan = PrimaryClosing
                    Me.StartDate = Me.dtpPeriod.Value
                    Me.EndDate = Me.ObjCD.GetLasDateOfTheMonth(Me.StartDate)
                    PeriodClosing = Me.StartDate

                    Me.LoadToDgv_Item(Me.PeriodClosing)
                    Me.LoadToDgv_DetailData(Me.PeriodClosing)
                    Me.LoadToDgv_DetailWareHouse(Me.PeriodClosing)
                    Me.LoadToDgv_DetailUsing(Me.PeriodClosing)
                    MsgBox("Data loaded", MsgBoxStyle.Information, "Information")
                End If
            Else
                MsgBox("No closing data for the selected period.", MsgBoxStyle.Information, "Information")
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub btnCek_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCek.Click
        Dim ij As New Form2_Cek

        Try
            ij.FrmParent = Me
            ij.ClosingPeriod = Me.PeriodClosing
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetPrimaryClosing() As Double
        Dim DTA As New DseCekTableAdapters.DataTable2TableAdapter
        Dim RetVal As Nullable(Of Double)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            RetVal = DTA.ScalarQuery_GetPrimary(Me.dtpPeriod.Value)
            If RetVal.HasValue = True Then
                Return RetVal.Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub btnRecalculate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnRecalculate.Click
        Dim PrimaryClosing As Double = -1

        Try
            PrimaryClosing = Me.GetPrimaryClosing
            If PrimaryClosing <> -1 Then
                If MsgBox("Recalculate data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    FrmDia.Title = "Please specify where to save csv file - item detail"
                    FrmDia.Filter = "*.csv|csv file"
                    FrmDia.FilterIndex = 0
                    FrmDia.DefaultExt = ".csv"

                    If FrmDia.ShowDialog <> DialogResult.OK Then
                        Exit Sub
                    End If
                    If FrmDia.FileName = "" Then
                        Exit Sub
                    End If

                    SaveToFile = FrmDia.FileName
                    StartTime = Now
                    Control.CheckForIllegalCrossThreadCalls = False
                    Me.bgw1.RunWorkerAsync()
                End If
            Else
                MsgBox("No closing is available", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShowSum.Click
        Dim QtyBB, AmountBB, QtyIn, AmountIn, QtyOut, AmountOut, QtyEnd, AmountEnd As Decimal
        Dim rowDGV As DataGridViewRow
        Dim Dummy As String = ""

        Try
            QtyBB = 0
            QtyIn = 0
            QtyOut = 0
            QtyEnd = 0
            AmountBB = 0
            AmountIn = 0
            AmountOut = 0
            AmountEnd = 0
            For Each rowDGV In Me.DataGridView1.Rows
                QtyBB += CDec(rowDGV.Cells(Me.ColQtyBB.Index).Value)
                QtyIn += CDec(rowDGV.Cells(Me.ColQtyIn.Index).Value)
                QtyOut += CDec(rowDGV.Cells(Me.ColOutQty.Index).Value)
                QtyEnd += CDec(rowDGV.Cells(Me.ColQtyEnd.Index).Value)
                AmountBB += CDec(rowDGV.Cells(Me.ColAmountBB.Index).Value)
                AmountIn += CDec(rowDGV.Cells(Me.ColAmountIn.Index).Value)
                AmountOut += CDec(rowDGV.Cells(Me.ColAmountOut.Index).Value)
                AmountEnd += CDec(rowDGV.Cells(Me.ColAmountEnd.Index).Value)
            Next
            Dummy = "Rekap: " + vbCrLf + _
            "Qty BB: " + Me.ObjCD.BuangNolDibelakangDecimal(Me.ObjCD.DigitGrouping(QtyBB, 5, _
            Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)) + " - Amount BB: " + Me.ObjCD.BuangNolDibelakangDecimal(Me.ObjCD.DigitGrouping(AmountBB, 5, _
            Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)) + vbCrLf + _
            "Qty In: " + Me.ObjCD.BuangNolDibelakangDecimal(Me.ObjCD.DigitGrouping(QtyIn, 5, _
            Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)) + " - Amount In: " + Me.ObjCD.BuangNolDibelakangDecimal(Me.ObjCD.DigitGrouping(AmountIn, 5, _
            Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)) + vbCrLf + _
            "Qty Out: " + Me.ObjCD.BuangNolDibelakangDecimal(Me.ObjCD.DigitGrouping(QtyOut, 5, _
            Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)) + " - Amount Out: " + Me.ObjCD.BuangNolDibelakangDecimal(Me.ObjCD.DigitGrouping(AmountOut, 5, _
            Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)) + vbCrLf + _
            "Qty End: " + Me.ObjCD.BuangNolDibelakangDecimal(Me.ObjCD.DigitGrouping(QtyEnd, 5, _
            Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)) + " - Amount End: " + Me.ObjCD.BuangNolDibelakangDecimal(Me.ObjCD.DigitGrouping(AmountEnd, 5, _
            Core_Module.ClsCekData.EnumTipeRounding.Middlepoint))

            Dummy = Dummy + vbCrLf + vbCrLf
            QtyBB = 0
            QtyIn = 0
            QtyOut = 0
            QtyEnd = 0
            AmountBB = 0
            AmountIn = 0
            AmountOut = 0
            AmountEnd = 0
            For Each rowDGV In Me.DataGridView2.Rows
                QtyBB += CDec(rowDGV.Cells(Me.ColItem_QtyBB.Index).Value)
                QtyIn += CDec(rowDGV.Cells(Me.ColItem_QtyIn.Index).Value)
                QtyOut += CDec(rowDGV.Cells(Me.ColItem_QtyOut.Index).Value)
                QtyEnd += CDec(rowDGV.Cells(Me.ColItem_QtyEnd.Index).Value)
                AmountBB += CDec(rowDGV.Cells(Me.ColItem_AmountBB.Index).Value)
                AmountIn += CDec(rowDGV.Cells(Me.ColItem_AmountIn.Index).Value)
                AmountOut += CDec(rowDGV.Cells(Me.ColItem_AmountOut.Index).Value)
                AmountEnd += CDec(rowDGV.Cells(Me.ColItem_AmountEnd.Index).Value)
            Next
            Dummy = Dummy + "Item Detail: " + vbCrLf + _
            "Qty BB: " + Me.ObjCD.BuangNolDibelakangDecimal(Me.ObjCD.DigitGrouping(QtyBB, 5, _
            Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)) + " - Amount BB: " + Me.ObjCD.BuangNolDibelakangDecimal(Me.ObjCD.DigitGrouping(AmountBB, 5, _
            Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)) + vbCrLf + _
            "Qty In: " + Me.ObjCD.BuangNolDibelakangDecimal(Me.ObjCD.DigitGrouping(QtyIn, 5, _
            Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)) + " - Amount In: " + Me.ObjCD.BuangNolDibelakangDecimal(Me.ObjCD.DigitGrouping(AmountIn, 5, _
            Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)) + vbCrLf + _
            "Qty Out: " + Me.ObjCD.BuangNolDibelakangDecimal(Me.ObjCD.DigitGrouping(QtyOut, 5, _
            Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)) + " - Amount Out: " + Me.ObjCD.BuangNolDibelakangDecimal(Me.ObjCD.DigitGrouping(AmountOut, 5, _
            Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)) + vbCrLf + _
            "Qty End: " + Me.ObjCD.BuangNolDibelakangDecimal(Me.ObjCD.DigitGrouping(QtyEnd, 5, _
            Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)) + " - Amount End: " + Me.ObjCD.BuangNolDibelakangDecimal(Me.ObjCD.DigitGrouping(AmountEnd, 5, _
            Core_Module.ClsCekData.EnumTipeRounding.Middlepoint))

            Dummy = Dummy + vbCrLf + vbCrLf
            QtyEnd = 0
            AmountEnd = 0
            For Each rowDGV In Me.DataGridView3.Rows
                QtyEnd += CDec(rowDGV.Cells(Me.ColW_QtyEnd.Index).Value)
                AmountEnd += CDec(rowDGV.Cells(Me.ColW_AmountEnd.Index).Value)
            Next
            Dummy = Dummy + "Item Detail Warehouse: " + vbCrLf + _
            "Qty End: " + Me.ObjCD.BuangNolDibelakangDecimal(Me.ObjCD.DigitGrouping(QtyEnd, 5, _
            Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)) + " - Amount End: " + Me.ObjCD.BuangNolDibelakangDecimal(Me.ObjCD.DigitGrouping(AmountEnd, 5, _
            Core_Module.ClsCekData.EnumTipeRounding.Middlepoint))

            Dummy = Dummy + vbCrLf + vbCrLf
            QtyEnd = 0
            AmountEnd = 0
            For Each rowDGV In Me.DataGridView4.Rows
                QtyEnd += CDec(rowDGV.Cells(Me.ColU_QtyUsing.Index).Value)
                AmountEnd += CDec(rowDGV.Cells(Me.ColW_AmountEnd.Index).Value)
            Next
            Dummy = Dummy + "Item Detail Using: " + vbCrLf + _
            "Qty Using: " + Me.ObjCD.BuangNolDibelakangDecimal(Me.ObjCD.DigitGrouping(QtyEnd, 5, _
            Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)) + " - Amount Using: " + Me.ObjCD.BuangNolDibelakangDecimal(Me.ObjCD.DigitGrouping(AmountEnd, 5, _
            Core_Module.ClsCekData.EnumTipeRounding.Middlepoint))

            MsgBox(Dummy, MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim ij As New Form2_Cek2

        Try
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCekItemVsDetail_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCekItemVsDetail.Click
        Try
            Dim ij As New Form2_Cek3
            ij.PeriodClosing = Me.dtpPeriod.Value.Date
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRepairCategory.Click
        Try
            If MsgBox("Update data category pada item closing custom?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                If MsgBox("Sudah insert index pada tblmaster_product & tbltrans_custom_closing_item?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Dim ij As New Form2_UpdateCategory
                    ij.vPeriod = Me.ObjCD.GetFirstDateofTheMonth(Me.dtpPeriod.Value.Date)
                    ij.ShowDialog()
                    MsgBox("Data updated")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Public Function CetakanStockDetail() As Boolean
    '    Try
    '        Control.CheckForIllegalCrossThreadCalls = False
    '        Me.bgw1.RunWorkerAsync()
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Function

    Public Function CetakanStockDetail1() As Boolean
        Dim DTA1 As New Inventory_Module.DseLaporanStockDetail2TableAdapters.tblmaster_productTableAdapter
        Dim TBL1 As New Inventory_Module.DseLaporanStockDetail2.tblmaster_productDataTable
        Dim row1, row2 As Inventory_Module.DseLaporanStockDetail2.tblmaster_productRow
        Dim DTA2 As New Inventory_Module.DataSet2TableAdapters.tbltrans_purchase_closing_item_cost_groupTableAdapter
        Dim TglAcuan As Date = #11/30/2017#
        Dim CurIdx, TargetIdx, BatchQty, FillIdx1, FillIdx2 As Integer
        Dim dtaSC As New Inventory_Module.DseLaporanStockDetail2TableAdapters.DataTable5TableAdapter
        Dim tblSC As New Inventory_Module.DseLaporanStockDetail2.DataTable5DataTable
        Dim tblSC1 As New Inventory_Module.DseLaporanStockDetail2.DataTable5DataTable
        Dim rowSC, rowSC1 As Inventory_Module.DseLaporanStockDetail2.DataTable5Row
        Dim dtaLTF As New Inventory_Module.DseLaporanStockDetail2TableAdapters.DataTable6TableAdapter
        Dim tblLTF As New Inventory_Module.DseLaporanStockDetail2.DataTable6DataTable
        Dim tblLTF1 As New Inventory_Module.DseLaporanStockDetail2.DataTable6DataTable
        Dim rowLTF, rowLTF1 As Inventory_Module.DseLaporanStockDetail2.DataTable6Row
        Dim DTA_ST As New Inventory_Module.DseLaporanStockDetail2TableAdapters.DataTable7TableAdapter
        Dim TBL_ST As New Inventory_Module.DseLaporanStockDetail2.DataTable7DataTable
        Dim TBL_ST1 As New Inventory_Module.DseLaporanStockDetail2.DataTable7DataTable
        Dim row_ST, row_ST1 As Inventory_Module.DseLaporanStockDetail2.DataTable7Row
        'Dim PrimaryMax, PrimaryMin As Double
        Dim DTA As New Inventory_Module.DseLaporanStockDetail2TableAdapters.DataTable1TableAdapter
        Dim TBL_Prod As New Inventory_Module.DseLaporanStockDetail2.DataTable1DataTable
        Dim TBL_Prod1 As New Inventory_Module.DseLaporanStockDetail2.DataTable1DataTable
        Dim row_Prod, row_Prod1 As Inventory_Module.DseLaporanStockDetail2.DataTable1Row

        Try
            Me.TBLDetail.Rows.Clear()

            FillIdx1 = -1
            FillIdx2 = -1

            Me.ObjCD.UpdateThreading()
            Me.TBLDetail.Rows.Clear()
            CurIdx = 0
            BatchQty = 99
            TglAcuan = Me.ObjCD.GetLasDateOfTheMonth(Me.dtpPeriod.Value.Date)

            'DTA.Connection.ConnectionString = Core_Module.GetConstring
            'dtaSC.Connection.ConnectionString = Core_Module.GetConstring
            'dtaLTF.Connection.ConnectionString = Core_Module.GetConstring
            'DTA_ST.Connection.ConnectionString = Core_Module.GetConstring
            'DTA1.Connection.ConnectionString = Core_Module.GetConstring

            Me.SudahLoadDGV = False
            'DTA1.FillBy_PrimaryProduct(TBL1, PrimaryProductAcuan)
            DTA1.Fill(TBL1)
            Me.ProgressBar1.Value = 0
            Me.ProgressBar1.Maximum = TBL1.Rows.Count
            Me.LblProgress.Text = "Loading detail qty"
            Me.TBLDetail.Rows.Clear()

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

                    Me.LblProgress2.Text = "Load detail product qty"
                    TBL_Prod.Rows.Clear()
                    DTA.FillBy_AllGoods2(TBL_Prod, row2.Primary, row1.Primary)
                    Me.LblProgress2.Text = "Load detail stock card"
                    tblSC.Rows.Clear()
                    dtaSC.FillBy_AllGoods2(tblSC, row2.Primary, row1.Primary, TglAcuan.Date.AddDays(1), CInt(Inventory_Module.ClsTransaksiStockCard.EnumJenisTransaksiAsal.Stock_Taking))
                    Me.LblProgress2.Text = "Load detail stock transfer"
                    tblLTF.Rows.Clear()
                    dtaLTF.FillBy_AllGoods2(tblLTF, row2.Primary, row1.Primary, TglAcuan.Date.AddDays(1))
                    Me.LblProgress2.Text = "Load detail stock taking"
                    TBL_ST.Rows.Clear()
                    DTA_ST.FillBy_AllGoods2(TBL_ST, TglAcuan.Date.AddDays(1), row2.Primary, row1.Primary)
                End If

                Me.ProgressBar1.Value += (TargetIdx - CurIdx)
                row1 = TBL1.Rows(CurIdx)
                row2 = TBL1.Rows(TargetIdx)
                TBL_Prod1.Rows.Clear()
                TBL_ST1.Rows.Clear()
                tblLTF1.Rows.Clear()
                tblSC1.Rows.Clear()


                Me.LblProgress.Text = "Phase 1 of 5 - Loading detail qty: " + "Data " + (TargetIdx + 1).ToString + " of " + Me.ProgressBar1.Maximum.ToString + " - " + row1.Product_Code

                Me.LblProgress.Text = "Phase 1 of 5 - Loading detail qty: " + "transfer detail product qty"
                TBL_Prod.DefaultView.RowFilter = "[Primary Product] <= " + row2.Primary.ToString + " AND [Primary Product] >= " + row1.Primary.ToString
                Me.ProgressBar2.Value = 0
                Me.ProgressBar2.Maximum = TBL_Prod.DefaultView.Count
                While TBL_Prod.DefaultView.Count > 0
                    Me.ProgressBar2.Value += 1
                    If Me.ProgressBar2.Value = Me.ProgressBar2.Maximum Or Math.IEEERemainder(Me.ProgressBar2.Value, 100) = 0 Then
                        Me.LblProgress2.Text = "Data. " & Me.ProgressBar2.Value.ToString & " of " & Me.ProgressBar2.Maximum.ToString & "."
                    End If
                    row_Prod = TBL_Prod.DefaultView(0).Row
                    row_Prod1 = TBL_Prod1.NewDataTable1Row
                    For gh As Integer = 0 To TBL_Prod.Columns.Count - 1
                        row_Prod1.Item(gh) = row_Prod.Item(gh)
                    Next
                    TBL_Prod1.Rows.Add(row_Prod1)
                    TBL_Prod.Rows.Remove(row_Prod)
                End While
                Me.LblProgress.Text = "Phase 1 of 5 - Loading detail qty: " + "transfer detail stock taking"
                TBL_ST.DefaultView.RowFilter = "[Primary Product] <= " + row2.Primary.ToString + " AND [Primary Product] >= " + row1.Primary.ToString
                Me.ProgressBar2.Value = 0
                Me.ProgressBar2.Maximum = TBL_ST.DefaultView.Count
                While TBL_ST.DefaultView.Count > 0
                    Me.ProgressBar2.Value += 1
                    If Me.ProgressBar2.Value = Me.ProgressBar2.Maximum Or Math.IEEERemainder(Me.ProgressBar2.Value, 100) = 0 Then
                        LblProgress2.Text = "Data. " & Me.ProgressBar2.Value.ToString & " of " & Me.ProgressBar2.Maximum.ToString & "."
                    End If
                    row_ST = TBL_ST.DefaultView(0).Row
                    row_ST1 = TBL_ST1.NewDataTable7Row
                    For gh As Integer = 0 To TBL_ST.Columns.Count - 1
                        row_ST1.Item(gh) = row_ST.Item(gh)
                    Next
                    TBL_ST1.Rows.Add(row_ST1)
                    TBL_ST.Rows.Remove(row_ST)
                End While
                Me.LblProgress.Text = "Phase 1 of 5 - Loading detail qty: " + "transfer detail stock transfer"
                tblLTF.DefaultView.RowFilter = "[Primary Product] <= " + row2.Primary.ToString + " AND [Primary Product] >= " + row1.Primary.ToString
                Me.ProgressBar2.Value = 0
                Me.ProgressBar2.Maximum = tblLTF.DefaultView.Count
                While tblLTF.DefaultView.Count > 0
                    Me.ProgressBar2.Value += 1
                    If Me.ProgressBar2.Value = Me.ProgressBar2.Maximum Or Math.IEEERemainder(Me.ProgressBar2.Value, 100) = 0 Then
                        LblProgress2.Text = "Data. " & Me.ProgressBar2.Value.ToString & " of " & Me.ProgressBar2.Maximum.ToString & "."
                    End If
                    rowLTF = tblLTF.DefaultView(0).Row
                    rowLTF1 = tblLTF1.NewDataTable6Row
                    For gh As Integer = 0 To tblLTF.Columns.Count - 1
                        rowLTF1.Item(gh) = rowLTF.Item(gh)
                    Next
                    tblLTF1.Rows.Add(rowLTF1)
                    tblLTF.Rows.Remove(rowLTF)
                End While

                Me.LblProgress.Text = "Phase 1 of 5 - Loading detail qty: " + "transfer detail stock card"
                tblSC.DefaultView.RowFilter = "[Primary Product] <= " + row2.Primary.ToString + " AND [Primary Product] >= " + row1.Primary.ToString
                Me.ProgressBar2.Value = 0
                Me.ProgressBar2.Maximum = tblSC.DefaultView.Count
                While tblSC.DefaultView.Count > 0
                    Me.ProgressBar2.Value += 1
                    If Me.ProgressBar2.Value = Me.ProgressBar2.Maximum Or Math.IEEERemainder(Me.ProgressBar2.Value, 100) = 0 Then
                        LblProgress2.Text = "Data. " & Me.ProgressBar2.Value.ToString & " of " & Me.ProgressBar2.Maximum.ToString & "."
                    End If
                    rowSC = tblSC.DefaultView(0).Row
                    rowSC1 = tblSC1.NewDataTable5Row
                    For gh As Integer = 0 To tblSC.Columns.Count - 1
                        rowSC1.Item(gh) = rowSC.Item(gh)
                    Next
                    tblSC1.Rows.Add(rowSC1)
                    tblSC.Rows.Remove(rowSC)
                End While

                Me.LblProgress.Text = "Phase 1 of 5 - Loading detail qty: " + "Data " + (TargetIdx + 1).ToString + " of " + Me.ProgressBar1.Maximum.ToString + " - " + row1.Product_Code
                Me.CetakanStockDetail2(TBL_Prod1, TglAcuan, tblSC1, tblLTF1, TBL_ST1)
                Me.ProgressBar2.Value = 0
                Me.ProgressBar2.Maximum = TBL_Prod1.Rows.Count
                For Each row_Prod In TBL_Prod1.Rows
                    Me.ProgressBar2.Value += 1
                    If Me.ProgressBar2.Value = Me.ProgressBar2.Maximum Or Math.IEEERemainder(Me.ProgressBar2.Value, 100) = 0 Then
                        LblProgress2.Text = "Moving result. " & Me.ProgressBar2.Value.ToString & " of " & Me.ProgressBar2.Maximum.ToString & "."
                    End If
                    row_Prod1 = Me.TBLDetail.NewDataTable1Row
                    For jk As Integer = 0 To Me.TBLDetail.Columns.Count - 1
                        row_Prod1.Item(jk) = row_Prod.Item(jk)
                    Next
                    Me.TBLDetail.Rows.Add(row_Prod1)
                Next
                CurIdx = TargetIdx + 1
            End While

            ''load ke dgv
            'Me.dgvStock.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            'Me.dgvStock.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            'Me.dgvStockTidakCocok.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            'Me.dgvStockTidakCocok.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            'Me.dgvStockMinus.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            'Me.dgvStockMinus.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            'Me.dgvStock.Rows.Clear()
            'Me.dgvStockTidakCocok.Rows.Clear()
            'Me.dgvStockMinus.Rows.Clear()

            ''Me.TipeProgress = EnumTipeProgress.ChangePrgStyle
            ''Me.bgwStep1_2.ReportProgress(0)

            'Me.TipeProgress = EnumTipeProgress.AddDGVStock
            'Me.bgwStep1_2.ReportProgress(0)
            'LblProgress2.Text = "Load data to DGV..."
            'While Me.SudahLoadDGV = False
            '    'LblProgress2.Text = "Load data to DGV - data " + Me.dgvStock.Rows.Count.ToString + " of " + TBL.Rows.Count.ToString
            '    Threading.Thread.Sleep(60000)
            'End While

            ''Me.TipeProgress = EnumTipeProgress.ChangePrgStyle
            ''Me.bgwStep1_2.ReportProgress(0)
            ''pengecekan stock tidak cocok
            ''If MsgBox("Check unmatched data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
            'Me.CekTidakCocok(TBL, TglAcuan)
            ''End If

            ''simpan ke db
            ''If MsgBox("Simpan ke db?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
            ''hapus db
            'Me.ProgressBar2.Value = 0
            'Me.LblProgress2.Text = "Deleting database on tbl item stock"
            'DTA2.DeleteQuery()
            ''simpan item
            'Me.Simpan_Item_UsingScript(TBL)
            ''End If

            Me.LblProgress2.Text = "Finished"
            Me.SudahLoadDGV = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function CetakanStockDetail2(ByVal TBL_Prod As Inventory_Module.DseLaporanStockDetail2.DataTable1DataTable, ByVal TglAcuan As Date, _
 ByVal tblSC As Inventory_Module.DseLaporanStockDetail2.DataTable5DataTable, ByVal tblLTF As Inventory_Module.DseLaporanStockDetail2.DataTable6DataTable, _
 ByVal TBL_ST As Inventory_Module.DseLaporanStockDetail2.DataTable7DataTable) As Boolean
        Dim Row As Inventory_Module.DseLaporanStockDetail2.DataTable1Row
        Dim rowSC As Inventory_Module.DseLaporanStockDetail2.DataTable5Row
        Dim rowLTF As Inventory_Module.DseLaporanStockDetail2.DataTable6Row
        Dim row_ST As Inventory_Module.DseLaporanStockDetail2.DataTable7Row

        Try
            Me.ProgressBar2.Value = 0
            Me.ProgressBar2.Maximum = tblSC.Rows.Count
            While tblSC.Rows.Count > 0
                rowSC = tblSC.Rows(0)
                Me.ProgressBar2.Value += 1
                If Me.ProgressBar2.Value = Me.ProgressBar2.Maximum Or Math.IEEERemainder(Me.ProgressBar2.Value, 100) = 0 Then
                    LblProgress2.Text = "Stock Card calculation... " & Me.ProgressBar2.Value.ToString & " of " & Me.ProgressBar2.Maximum.ToString & "."
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

            LblProgress2.Text = "Preparing for stock transfer reversal...."
            Me.ProgressBar2.Value = 0
            Me.ProgressBar2.Maximum = tblLTF.Rows.Count
            While tblLTF.Rows.Count > 0
                Me.ProgressBar2.Value += 1
                If Me.ProgressBar2.Value = Me.ProgressBar2.Maximum Or Math.IEEERemainder(Me.ProgressBar2.Value, 100) = 0 Then
                    LblProgress2.Text = "Stock Transfer Calculation... " & Me.ProgressBar2.Value.ToString & " of " & Me.ProgressBar2.Maximum.ToString & "."
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

            Me.ProgressBar2.Value = 0
            Me.ProgressBar2.Maximum = TBL_ST.Rows.Count
            While TBL_ST.Rows.Count > 0
                row_ST = TBL_ST.Rows(0)
                Me.ProgressBar2.Value += 1
                If Me.ProgressBar2.Value = Me.ProgressBar2.Maximum Or Math.IEEERemainder(Me.ProgressBar2.Value, 100) = 0 Then
                    LblProgress2.Text = "Stock Taking calculation... " & Me.ProgressBar2.Value.ToString & " of " & Me.ProgressBar2.Maximum.ToString & "."
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

            LblProgress2.Text = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            tblSC.Rows.Clear()
            tblLTF.Rows.Clear()
            TBL_ST.Rows.Clear()
        End Try
    End Function

    Private Sub bgw1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgw1.DoWork
        Try
            btnRecalculate.Enabled = False
            btnSave1.Enabled = False
            btnCreateCSV.Enabled = False
            BtnLoad1.Enabled = False
            Me.dtpPeriod.Enabled = False

            Me.DataGridView1.Enabled = False
            Me.TabControl1.Enabled = False

            Me.CetakanStockDetail1()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgw1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgw1.RunWorkerCompleted
        Try
            Me.bgw2.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim StartDate As Date = #7/1/2017#
        Dim DTAClosing As New DseClosingCutOff01TableAdapters.tbltrans_purchase_closing1TableAdapter
        Dim TBLClosing As New DseClosingCutOff01.tbltrans_purchase_closing1DataTable
        Dim rowClosing As DseClosingCutOff01.tbltrans_purchase_closing1Row
        Dim DTAJurnal As New DseClosingCutOff01TableAdapters.tbltrans_jurnalTableAdapter
        Dim TBLJurnal As New DseClosingCutOff01.tbltrans_jurnalDataTable
        Dim rowJurnal As DseClosingCutOff01.tbltrans_jurnalRow
        Dim TipeJurnal As AccModule.ClsTransaksiJournal.EnumTipeJournal = AccModule.ClsTransaksiJournal.EnumTipeJournal.GeneralJournal
        Dim TipeJurnal2 As AccModule.ClsTransaksiJournal.EnumTipeJurnal2 = AccModule.ClsTransaksiJournal.EnumTipeJurnal2.Jurnal_NoGrouping
        Dim ObjClosing As New Purchase_Module.ClsMonthlyClosing01
        Dim Frm As FrmTransaksiMonthlyClosing2

        Try
            DTAClosing.Connection.ConnectionString = Core_Module.GetConstring
            DTAJurnal.Connection.ConnectionString = Core_Module.GetConstring

            While StartDate < #1/1/2018#
                'load data closing
                DTAClosing.Fill(TBLClosing, StartDate)
                rowClosing = TBLClosing.Rows(0)

                'buat jurnal closing baru
                Frm = New FrmTransaksiMonthlyClosing2
                Frm.Show()
                Frm.CreateJurnal(rowClosing.Primary)


                'start date naik 1 bulan
                StartDate = Me.ObjCD.GetFirstDateofTheMonth(StartDate + TimeSpan.FromDays(32))

                Frm.Dispose()
            End While

            MsgBox("Done", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgw2_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgw2.DoWork
        Dim PrimaryClosing As Double = -1

        Try
            Me.LblProgress.Text = ""
            Me.LblProgress2.Text = ""
            Me.ProgressBar1.Value = 0
            Me.ProgressBar2.Value = 0

            PrimaryClosing = Me.GetPrimaryClosing
            Me.Recalculate(PrimaryClosing)

            Me.LblProgress.Text = "Refreshing..."
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgw2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgw2.RunWorkerCompleted
        Dim rowDGV1, rowDGV2 As DataGridViewRow

        Try
            'Me.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            'Me.DgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            'Me.DgvItem.Rows.Clear()

            'For Each rowDGV1 In Me.DataGridView1.Rows
            '    Me.DgvItem.Rows.Add()
            '    rowDGV2 = Me.DgvItem.Rows(Me.DgvItem.Rows.Count - 1)

            '    For ij As Integer = 0 To Me.DgvItem.ColumnCount - 1
            '        rowDGV2.Cells(ij).Value = rowDGV1.Cells(ij).Value
            '    Next
            'Next

            'Me.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            btnRecalculate.Enabled = True
            btnSave1.Enabled = True
            btnCreateCSV.Enabled = True
            BtnLoad1.Enabled = True
            Me.dtpPeriod.Enabled = True

            Me.DataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DataGridView2.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.DataGridView3.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DataGridView3.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.DataGridView4.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DataGridView4.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            'Me.DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            'Me.DataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            'Me.DataGridView3.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            'Me.DataGridView4.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            Me.DataGridView1.Enabled = False
            Me.TabControl1.Enabled = False
            Me.LblProgress.Text = ""

            Me.CreateCsv(Me.SaveToFile)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class