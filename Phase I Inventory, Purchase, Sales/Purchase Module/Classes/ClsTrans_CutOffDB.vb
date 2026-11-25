Option Explicit On
Option Strict Off

Public Class ClsTrans_CutOffDB
    Public Enum EnumClosingMethod As Integer
        Average
        FIFO
        LIFO
    End Enum

    Public Function CekUserLogin() As Boolean
        Dim Dta As New DseTrans_CutOffDBTableAdapters.tbltrans_monitorTableAdapter
        Dim Jml As New Nullable(Of Long)

        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            Jml = Dta.ScalarQuery_GetJmlUserLogin()
            If Jml.HasValue = True Then
                If Jml.Value > 1 Then
                    Return False
                Else
                    Return True
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Sub ControlUserLogin(ByRef vUseActiveLicense As Boolean, ByRef vActiveLicense As Integer, _
    Optional ByVal vFromClose As Boolean = False)
        Dim Dta As New DseTrans_CutOffDBTableAdapters.tbl_smessTableAdapter
        Dim Tbl As New DseTrans_CutOffDB.tbl_smessDataTable
        Dim row As DseTrans_CutOffDB.tbl_smessRow

        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            Dta.Fill(Tbl)
            If vFromClose = False Then
                If Tbl.Count > 0 Then
                    row = Tbl.Rows(0)

                    vUseActiveLicense = True
                    vActiveLicense = row.active_license

                    'jika menggunakan active license maka update active license = 1
                    'agar tidak ada user lain yang daat login selain user yang sedang
                    'melakukan cut off database
                    Dta.Update(1, row.primary)
                Else
                    'jika tidak menggunakan active license maka buat active license = 1 
                    'agar tidak ada user lain yang daat login selain user yang sedang
                    'melakukan cut off database
                    Dta.Insert(1)
                End If
            Else
                'jika user yang melakukan proses cut off database menutup form cut of db
                'maka kembalikan settingan kembali seperti semula
                If vUseActiveLicense = True Then
                    If Tbl.Count > 0 Then
                        row = Tbl.Rows(0)
                        Dta.Update(vActiveLicense, row.primary)
                    Else
                        Dta.Insert(vActiveLicense)
                    End If
                Else
                    Dta.Delete()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Dta.Dispose()
            Tbl.Dispose()
        End Try
    End Sub

    Public Function TransferBeginingBalance_GetAmount(ByRef tblAmount As DseTrans_CutOffDB.tbltrans_purchase_closing_item_detailDataTable, _
    ByVal vPrimaryProduct As Double, ByVal vQty As Decimal, ByVal vClosingMethod As EnumClosingMethod) As Decimal
        Dim rowAmount As DseTrans_CutOffDB.tbltrans_purchase_closing_item_detailRow
        Dim QtyHasil As Decimal = 0
        Dim TotalAmount As Decimal = 0

        Try
            tblAmount.DefaultView.RowFilter = "[primary code no] = " & vPrimaryProduct.ToString
            If tblAmount.DefaultView.Count > 0 Then
                If vClosingMethod = EnumClosingMethod.Average Then
                    rowAmount = tblAmount.DefaultView(0).Row
                    TotalAmount = vQty * rowAmount.Unit_Price
                ElseIf vClosingMethod = EnumClosingMethod.FIFO Then
                    For Each dv As DataRowView In tblAmount.DefaultView
                        rowAmount = dv.Row
                        If rowAmount.qty > 0 Then
                            rowAmount.qty = rowAmount.qty - vQty
                            If rowAmount.qty < 0 Then
                                QtyHasil = vQty - (rowAmount.qty * -1)
                                rowAmount.qty = 0
                            Else
                                QtyHasil = vQty
                            End If
                            vQty = vQty - QtyHasil
                        End If
                        TotalAmount += QtyHasil * rowAmount.Unit_Price

                        If vQty <= 0 Then
                            Exit For
                        End If
                    Next
                ElseIf vClosingMethod = EnumClosingMethod.LIFO Then
                    'SEMENTARA UNTUK LIFO MEMANG TIDAK DI BUAT FUNGSI NYA DULU
                End If

                Return TotalAmount
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        End Try
    End Function

    Public Function TransferBeginingBalance_GetTotalAmount(ByVal tblAmount As DseTrans_CutOffDB.tbltrans_purchase_closing_item_detailDataTable, _
    ByVal vPrimaryProduct As Double, ByVal vClosingMethod As EnumClosingMethod) As Decimal
        Dim rowAmount As DseTrans_CutOffDB.tbltrans_purchase_closing_item_detailRow
        Dim TotalAmount As Decimal = 0

        Try
            tblAmount.DefaultView.RowFilter = "[primary code no] = " & vPrimaryProduct.ToString
            If tblAmount.DefaultView.Count > 0 Then
                If vClosingMethod = EnumClosingMethod.Average Then
                    rowAmount = tblAmount.DefaultView(0).Row
                    TotalAmount = rowAmount.amount
                ElseIf vClosingMethod = EnumClosingMethod.FIFO Then
                    For Each dv As DataRowView In tblAmount.DefaultView
                        rowAmount = dv.Row
                        TotalAmount += rowAmount.amount
                    Next
                ElseIf vClosingMethod = EnumClosingMethod.LIFO Then
                    'SEMENTARA UNTUK LIFO MEMANG TIDAK DI BUAT FUNGSI NYA DULU
                End If

                Return TotalAmount
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        End Try
    End Function

    Public Sub UpdateSaldoHeader(ByVal vTBL As DseTrans_CutOffDB.tbl_beginning_balanceDataTable)
        Dim row As DseTrans_CutOffDB.tbl_beginning_balanceRow
        Dim i As Integer = 0
        Try
            For Each row In vTBL.Rows
                If row.Header = 1 Then
                    If row.Saldo_Normal = 0 Then
                        row.Debet = Me.GetSaldoHeader(i, row.Saldo_Normal, False, row.Level_Perkiraan, vTBL)
                        row.Debet_Original = Me.GetSaldoHeader(i, row.Saldo_Normal, True, row.Level_Perkiraan, vTBL)
                    Else
                        row.Kredit = Me.GetSaldoHeader(i, row.Saldo_Normal, False, row.Level_Perkiraan, vTBL)
                        row.Credit_Original = Me.GetSaldoHeader(i, row.Saldo_Normal, True, row.Level_Perkiraan, vTBL)
                    End If
                End If

                i += 1
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetSaldoHeader(ByVal vRowIndex As Double, ByVal vDebetKredit As Integer, ByVal vOriginalCurr As Boolean, _
    ByVal vLevel As Integer, ByVal vTBL As DseTrans_CutOffDB.tbl_beginning_balanceDataTable) As Decimal
        Dim TargetIndex, CurrentIndex As Integer
        Dim row As DseTrans_CutOffDB.tbl_beginning_balanceRow
        Dim RetVal As Double = 0

        Try
            CurrentIndex = vRowIndex + 1
            While CurrentIndex <= vTBL.Rows.Count - 1
                row = vTBL.Rows(CurrentIndex)
                If row.Level_Perkiraan <= vLevel Then
                    TargetIndex = CurrentIndex - 1
                    Exit While
                ElseIf CurrentIndex = vTBL.Rows.Count - 1 Then
                    TargetIndex = CurrentIndex
                    Exit While
                End If
                CurrentIndex += 1
            End While

            For i As Integer = vRowIndex + 1 To TargetIndex
                row = vTBL.Rows(i)
                If row.Header = 1 And row.Level_Perkiraan = vLevel + 1 Then
                    RetVal += Me.GetSaldoHeader(i, vDebetKredit, vOriginalCurr, row.Level_Perkiraan, vTBL)
                ElseIf row.Level_Perkiraan = vLevel + 1 Then
                    If vDebetKredit = 0 Then
                        If vOriginalCurr = False Then
                            RetVal += CDec(row.Debet)
                        Else
                            RetVal += CDec(row.Debet_Original)
                        End If
                    Else
                        If vOriginalCurr = False Then
                            RetVal += CDec(row.Kredit)
                        Else
                            RetVal += CDec(row.Credit_Original)
                        End If
                    End If
                End If
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Sub GetTblPOB_Filter(ByVal TblPL_Item_OutSumQty As DseTrans_CutOffDB.tbltrans_packinglist_itemDataTable, _
    ByRef TblPL_Item_OutSumQty_Filter As DseTrans_CutOffDB.tbltrans_packinglist_itemDataTable, _
    ByVal TblPOC_Item_OutSumQty As DseTrans_CutOffDB.tbltrans_poc_itemDataTable, _
    ByRef TblPOC_Item_OutSumQty_Filter As DseTrans_CutOffDB.tbltrans_poc_itemDataTable, _
    ByVal TblPRet_Item_OutSumQty As DseTrans_CutOffDB.tbltrans_purchasereturn_itemDataTable, _
    ByRef TblPRet_Item_OutSumQty_Filter As DseTrans_CutOffDB.tbltrans_purchasereturn_itemDataTable, _
    ByVal TblSIP_Item_OutSumQty As DseTrans_CutOffDB.tbltransaksi_stockinout_itemDataTable, _
    ByRef TblSIP_Item_OutSumQty_Filter As DseTrans_CutOffDB.tbltransaksi_stockinout_itemDataTable, _
    ByVal PrimaryPO_Min As Double, ByVal PrimaryPO_Max As Double)

        Try
            'packing list item qty sum
            TblPL_Item_OutSumQty_Filter.Clear()
            TblPL_Item_OutSumQty.DefaultView.RowFilter = "[Primary PO] >= " & PrimaryPO_Min & " AND [Primary PO] <= " & PrimaryPO_Max
            TblPL_Item_OutSumQty_Filter.Merge(TblPL_Item_OutSumQty.DefaultView.ToTable)
            'stock in purchase item qty sum
            TblSIP_Item_OutSumQty_Filter.Clear()
            TblSIP_Item_OutSumQty.DefaultView.RowFilter = "[Primary PO Purchase] >= " & PrimaryPO_Min & " AND [Primary PO Purchase] <= " & PrimaryPO_Max
            TblSIP_Item_OutSumQty_Filter = TblSIP_Item_OutSumQty.DefaultView.Table.Copy
            'poc item qty sum
            TblPOC_Item_OutSumQty_Filter.Clear()
            TblPOC_Item_OutSumQty.DefaultView.RowFilter = "[Primary PO] >= " & PrimaryPO_Min & " AND [Primary PO] <= " & PrimaryPO_Max
            TblPOC_Item_OutSumQty_Filter = TblPOC_Item_OutSumQty.DefaultView.Table.Copy
            'pr item qty sum
            TblPRet_Item_OutSumQty_Filter.Clear()
            TblPRet_Item_OutSumQty.DefaultView.RowFilter = "[Primary PO] >= " & PrimaryPO_Min & " AND [Primary PO] <= " & PrimaryPO_Max
            TblPRet_Item_OutSumQty_Filter = TblPRet_Item_OutSumQty.DefaultView.Table.Copy
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferSOB")
        End Try
    End Sub

    Public Sub GetTblPO_Filter(ByVal TblPO_Item As DseTrans_CutOffDB.tbltrans_po_itemDataTable, _
    ByRef TblPO_Item_Filter As DseTrans_CutOffDB.tbltrans_po_itemDataTable, _
    ByVal TblPO_ItemDel As DseTrans_CutOffDB.tbltrans_po_itemdeliveryDataTable, _
    ByRef TblPO_ItemDel_Filter As DseTrans_CutOffDB.tbltrans_po_itemdeliveryDataTable, _
    ByVal PrimaryPO_Min As Double, ByVal PrimaryPO_Max As Double)

        Try
            'po item
            TblPO_Item_Filter.Clear()
            TblPO_Item.DefaultView.RowFilter = "[Primary Transaksi] >= " & PrimaryPO_Min & " AND [Primary Transaksi] <= " & PrimaryPO_Max
            TblPO_Item_Filter = TblPO_Item.DefaultView.Table.Copy
            'po item delivery
            TblPO_ItemDel_Filter.Clear()
            TblPO_ItemDel.DefaultView.RowFilter = "[Primary Transaksi] >= " & PrimaryPO_Min & " AND [Primary Transaksi] <= " & PrimaryPO_Max
            TblPO_ItemDel_Filter = TblPO_ItemDel.DefaultView.Table.Copy
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferSOB")
        End Try
    End Sub

    Public Function GetBalancePO(ByVal vTblPL_Item As DseTrans_CutOffDB.tbltrans_packinglist_itemDataTable, _
    ByVal vTblPOC_Item As DseTrans_CutOffDB.tbltrans_poc_itemDataTable, _
    ByVal vTblSIP_Item As DseTrans_CutOffDB.tbltransaksi_stockinout_itemDataTable, _
    ByVal vTblPRet_Item As DseTrans_CutOffDB.tbltrans_purchasereturn_itemDataTable, _
    ByRef vRowPO_Item As DseTrans_CutOffDB.tbltrans_po_itemRow, _
    ByVal vTblDec_PO As DseTrans_CutOffDB.tblmaster_productdecimalDataTable) As Boolean
        Dim rowPL_Item As DseTrans_CutOffDB.tbltrans_packinglist_itemRow
        Dim rowPOC_Item As DseTrans_CutOffDB.tbltrans_poc_itemRow
        Dim rowSIP_Item As DseTrans_CutOffDB.tbltransaksi_stockinout_itemRow
        Dim rowPRet_Item As DseTrans_CutOffDB.tbltrans_purchasereturn_itemRow
        Dim rowDec_PO As DseTrans_CutOffDB.tblmaster_productdecimalRow

        Dim QtyPL_Multi As Decimal = 0
        Dim QtySIP_Multi As Decimal = 0
        Dim QtyPR_Multi As Decimal = 0
        Dim QtyPOC_Multi As Decimal = 0
        Dim DecimalProduct As Integer = 0
        Dim TransPOItem As Boolean = False

        Try
            'kurangkan qty item po dengan yang ada di packing list
            vTblPL_Item.DefaultView.RowFilter = "[Primary PO] = " & vRowPO_Item.Primary_Transaksi.ToString & " AND [Primary Product] = " & vRowPO_Item.Primary_Product
            If vTblPL_Item.DefaultView.Count > 0 Then
                rowPL_Item = vTblPL_Item.DefaultView(0).Row
                QtyPL_Multi = rowPL_Item.Total_Qty_Multi
                TransPOItem = True
            End If

            'kurangkan qty item po dengan yang ada di stock in purchase
            vTblSIP_Item.DefaultView.RowFilter = "[Primary PO Purchase] = " & vRowPO_Item.Primary_Transaksi.ToString & " AND [Primary Product] = " & vRowPO_Item.Primary_Product.ToString & " AND [Tipe Transaksi] = " & Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase
            If vTblSIP_Item.DefaultView.Count > 0 Then
                rowSIP_Item = vTblSIP_Item.DefaultView(0).Row
                QtySIP_Multi = rowSIP_Item.Qty_Multi_Unit
                QtySIP_Multi -= QtyPL_Multi
                TransPOItem = True
            End If

            'tambahkan qty item po dengan yang ada di purchase return
            vTblPRet_Item.DefaultView.RowFilter = "[Primary PO] = " & vRowPO_Item.Primary_Transaksi.ToString & " AND [Primary Product] = " & vRowPO_Item.Primary_Product
            If vTblPRet_Item.DefaultView.Count > 0 Then
                rowPRet_Item = vTblPRet_Item.DefaultView(0).Row
                QtyPR_Multi = rowPRet_Item.Total_Qty_Multi
                TransPOItem = True
            End If

            'kurangkan qty item po dengan yang ada di poc
            vTblPOC_Item.DefaultView.RowFilter = "[Primary PO] = " & vRowPO_Item.Primary_Transaksi.ToString & " AND [Primary Product] = " & vRowPO_Item.Primary_Product
            If vTblPOC_Item.DefaultView.Count > 0 Then
                rowPOC_Item = vTblPOC_Item.DefaultView(0).Row
                QtyPOC_Multi = rowPOC_Item.Total_Qty_Multi
                TransPOItem = True
            End If

            vTblDec_PO.DefaultView.RowFilter = "[Primary Product] = " & vRowPO_Item.Primary_Product
            If vTblDec_PO.DefaultView.Count > 0 Then
                rowDec_PO = vTblDec_PO.DefaultView(0).Row
                DecimalProduct = rowDec_PO.Decimal_Default
            End If

            vRowPO_Item.Qty_Multi = vRowPO_Item.POB_Multi + QtyPL_Multi + QtySIP_Multi - QtyPR_Multi + QtyPOC_Multi
            vRowPO_Item.Qty = Math.Round(vRowPO_Item.Qty_Multi * vRowPO_Item.Ratio_Default / vRowPO_Item.Ratio_Reference, DecimalProduct, MidpointRounding.AwayFromZero)
            vRowPO_Item.POB_Multi = vRowPO_Item.Qty_Multi
            vRowPO_Item.POB = vRowPO_Item.Qty

            If TransPOItem = True And vRowPO_Item.Qty_Multi = 0 Then
                Return True
                'vRowPO_Item.Qty_Multi = QtyPL_Multi + QtySIP_Multi
                'If vRowPO_Item.Qty_Multi = 0 Then
                '    vRowPO_Item.Qty_Multi = QtyPR_Multi
                'End If
                'vRowPO_Item.Qty = Math.Round(vRowPO_Item.Qty_Multi * vRowPO_Item.Ratio_Default / vRowPO_Item.Ratio_Reference, DecimalProduct, MidpointRounding.AwayFromZero)
            End If
            If vRowPO_Item.Qty_Multi = 0 Then
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: GetBalancePO")
            Return False
        End Try
    End Function

    Public Sub GetTblSO_Filter(ByVal TblSO_Item As DseTrans_CutOffDB.tbltrans_so_itemDataTable, _
    ByRef TblSO_Item_Filter As DseTrans_CutOffDB.tbltrans_so_itemDataTable, _
    ByVal TblSO_ItemDel As DseTrans_CutOffDB.tbltrans_so_itemdeliveryDataTable, _
    ByRef TblSO_ItemDel_Filter As DseTrans_CutOffDB.tbltrans_so_itemdeliveryDataTable, _
    ByVal PrimarySO_Min As Double, ByVal PrimarySO_Max As Double)

        Try
            'so item
            TblSO_Item_Filter.Clear()
            TblSO_Item.DefaultView.RowFilter = "[Primary Transaksi] >= " & PrimarySO_Min & " AND [Primary Transaksi] <= " & PrimarySO_Max
            TblSO_Item_Filter = TblSO_Item.DefaultView.Table.Copy
            'so item delivery
            TblSO_ItemDel_Filter.Clear()
            TblSO_ItemDel.DefaultView.RowFilter = "[Primary Transaksi] >= " & PrimarySO_Min & " AND [Primary Transaksi] <= " & PrimarySO_Max
            TblSO_ItemDel_Filter = TblSO_ItemDel.DefaultView.Table.Copy
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferSOB")
        End Try
    End Sub

    Public Sub GetTblSOB_Filter(ByVal TblDoItem_OutSumQty As DseTrans_CutOffDB.tbltrans_do_itemDataTable, _
    ByRef TblDoItem_OutSumQty_Filter As DseTrans_CutOffDB.tbltrans_do_itemDataTable, _
    ByVal TblSOCItem_OutSumQty As DseTrans_CutOffDB.tbltrans_soc_itemDataTable, _
    ByRef TblSOCItem_OutSumQty_Filter As DseTrans_CutOffDB.tbltrans_soc_itemDataTable, _
    ByVal TblSRItem_OutSumQty As DseTrans_CutOffDB.tbltrans_salesreturn_itemDataTable, _
    ByRef TblSRItem_OutSumQty_Filter As DseTrans_CutOffDB.tbltrans_salesreturn_itemDataTable, _
    ByVal PrimarySO_Min As Double, ByVal PrimarySO_Max As Double)
        Try
            'do item qty sum
            TblDoItem_OutSumQty_Filter.Clear()
            TblDoItem_OutSumQty.DefaultView.RowFilter = "[Primary SO] >= " & PrimarySO_Min & " AND [Primary SO] <= " & PrimarySO_Max
            TblDoItem_OutSumQty_Filter.Merge(TblDoItem_OutSumQty.DefaultView.ToTable)
            'soc item qty sum
            TblSOCItem_OutSumQty_Filter.Clear()
            TblSOCItem_OutSumQty.DefaultView.RowFilter = "[Primary SO] >= " & PrimarySO_Min & " AND [Primary SO] <= " & PrimarySO_Max
            TblSOCItem_OutSumQty_Filter = TblSOCItem_OutSumQty.DefaultView.Table.Copy
            'sr item qty sum
            TblSRItem_OutSumQty_Filter.Clear()
            TblSRItem_OutSumQty.DefaultView.RowFilter = "[Primary SO] >= " & PrimarySO_Min & " AND [Primary SO] <= " & PrimarySO_Max
            TblSRItem_OutSumQty_Filter = TblSRItem_OutSumQty.DefaultView.Table.Copy
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferSOB")
        End Try
    End Sub

    Public Function GetBalanceSO(ByVal vTblDO_Item As DseTrans_CutOffDB.tbltrans_do_itemDataTable, _
          ByVal vTblSOC_Item As DseTrans_CutOffDB.tbltrans_soc_itemDataTable, _
          ByVal vTblSR_Item As DseTrans_CutOffDB.tbltrans_salesreturn_itemDataTable, _
          ByRef vRowSOItem As DseTrans_CutOffDB.tbltrans_so_itemRow, _
          ByVal vTblDec_SO As DseTrans_CutOffDB.tblmaster_productdecimalDataTable, _
          ByVal vTblSISRItem As DseTrans_CutOffDB.tbltransaksi_stockinout_itemDataTable) As Boolean
        Dim rowDoItem As DseTrans_CutOffDB.tbltrans_do_itemRow
        Dim rowSOCItem As DseTrans_CutOffDB.tbltrans_soc_itemRow
        Dim rowSRItem As DseTrans_CutOffDB.tbltrans_salesreturn_itemRow
        Dim rowDec_SO As DseTrans_CutOffDB.tblmaster_productdecimalRow
        Dim rowSISRItem As DseTrans_CutOffDB.tbltransaksi_stockinout_itemRow
        Dim QtyDO_Multi As Decimal = 0
        Dim QtyDO As Decimal = 0
        Dim QtySOS_Multi As Decimal = 0
        Dim QtySOS As Decimal = 0
        Dim QtySR_Multi As Decimal = 0
        Dim QtySR As Decimal = 0
        Dim QtySOC_Multi As Decimal = 0
        Dim QtySOC As Decimal = 0
        Dim DecimalProduct As Integer = 0
        Dim TransSOItem As Boolean = False

        Try
            'delivery order
            vTblDO_Item.DefaultView.RowFilter = "[Primary SO] = " & vRowSOItem.Primary_Transaksi & " AND [Primary Product] = " & vRowSOItem.Primary_Product
            If vTblDO_Item.DefaultView.Count > 0 Then
                rowDoItem = vTblDO_Item.DefaultView(0).Row
                QtyDO_Multi += rowDoItem.Total_Qty_Multi
                QtyDO += rowDoItem.Total_Qty
                TransSOItem = True
            End If

            'sales cancelation
            vTblSOC_Item.DefaultView.RowFilter = "[Primary SO] = " & vRowSOItem.Primary_Transaksi & " AND [Primary Product] = " & vRowSOItem.Primary_Product
            If vTblSOC_Item.DefaultView.Count > 0 Then
                rowSOCItem = vTblSOC_Item.DefaultView(0).Row
                QtySOC_Multi += rowSOCItem.Qty_Multi
                QtySOC += rowSOCItem.Qty
                TransSOItem = True
            End If

            'sales return
            vTblSR_Item.DefaultView.RowFilter = "[Primary SO] = " & vRowSOItem.Primary_Transaksi & " AND [Primary Product] = " & vRowSOItem.Primary_Product
            If vTblSR_Item.DefaultView.Count > 0 Then
                rowSRItem = vTblSR_Item.DefaultView(0).Row
                QtySR_Multi += rowSRItem.Total_Qty_Multi
                QtySR += rowSRItem.Total_Qty
                TransSOItem = True
            End If

            vTblDec_SO.DefaultView.RowFilter = "[Primary Product] = " & vRowSOItem.Primary_Product
            If vTblDec_SO.DefaultView.Count > 0 Then
                rowDec_SO = vTblDec_SO.DefaultView(0).Row
                DecimalProduct = rowDec_SO.Decimal_Default
            End If

            vRowSOItem.Qty_Multi = vRowSOItem.SOB_Multi + QtyDO_Multi + QtySOC_Multi - QtySR_Multi
            vRowSOItem.Qty = Math.Round(vRowSOItem.Qty_Multi * vRowSOItem.Ratio_Default / vRowSOItem.Ratio_Reference, DecimalProduct, MidpointRounding.AwayFromZero)
            vRowSOItem.SOB_Multi = vRowSOItem.Qty_Multi
            vRowSOItem.SOB = vRowSOItem.Qty

            If TransSOItem = True And vRowSOItem.Qty_Multi = 0 Then
                Return True
                'vRowSOItem.Qty_Multi = QtyDO_Multi
                'If vRowSOItem.Qty_Multi = 0 Then
                ''stock in sales return
                'vTblSISRItem.DefaultView.RowFilter = "[Primary PO Purchase] = " & vRowSOItem.Primary_Transaksi & " AND [Primary Product] = " & vRowSOItem.Primary_Product
                'If vTblSISRItem.DefaultView.Count > 0 Then
                '    rowSISRItem = vTblSISRItem.DefaultView(0).Row
                '    vRowSOItem.Qty_Multi = rowSISRItem.Qty_Multi_Unit
                'End If
                'End If
                'vRowSOItem.Qty = Math.Round(vRowSOItem.Qty_Multi * vRowSOItem.Ratio_Default / vRowSOItem.Ratio_Reference, DecimalProduct, MidpointRounding.AwayFromZero)
            End If
            If vRowSOItem.Qty_Multi = 0 Then
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: GetBalanceSO")
            Return False
        End Try
    End Function

    Public Sub GetTblDObySOS(ByVal TblDo3 As DseTrans_CutOffDB.tbltrans_doDataTable, _
    ByRef TblDo3_filter As DseTrans_CutOffDB.tbltrans_doDataTable, _
    ByVal TblDoItem3 As DseTrans_CutOffDB.tbltrans_do_itemDataTable, _
    ByRef TblDoItem3_filter As DseTrans_CutOffDB.tbltrans_do_itemDataTable, _
    ByVal PrimaryDO_Min As Double, ByVal PrimaryDO_Max As Double)
        Try
            TblDo3_filter.Clear()
            TblDo3.DefaultView.RowFilter = "[Primary] >= " & PrimaryDO_Min & " AND [Primary] <= " & PrimaryDO_Max
            TblDo3_filter.Merge(TblDo3.DefaultView.ToTable)

            TblDoItem3_filter.Clear()
            TblDoItem3.DefaultView.RowFilter = "[Primary DO] >= " & PrimaryDO_Min & " AND [Primary DO] <= " & PrimaryDO_Max
            TblDoItem3_filter.Merge(TblDoItem3.DefaultView.ToTable)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error: TransferSOB")
        End Try
    End Sub

    Public Function GetDefPrec(ByVal vTbl As Core_Module.dseMasterCurrency.tblmaster_currency1DataTable, ByVal vPrimaryCurr As Double) As Integer
        Dim row As Core_Module.dseMasterCurrency.tblmaster_currency1Row

        Try
            'vTbl.DefaultView.RowFilter = Nothing
            vTbl.DefaultView.RowFilter = "[Primary] = " & vPrimaryCurr
            If vTbl.DefaultView.Count > 0 Then
                row = vTbl.DefaultView(0).Row
                Return row.Default_Precision
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        End Try
    End Function

    Public Function GetNoPReturn(ByVal vTblPR As DseTrans_CutOffDB.tbltrans_purchasereturn1DataTable, ByVal vPrimaryPR_Obs As Double) As String
        Dim rowRP As DseTrans_CutOffDB.tbltrans_purchasereturn1Row

        Try
            vTblPR.DefaultView.RowFilter = "[Primary] = " & vPrimaryPR_Obs
            If vTblPR.DefaultView.Count > 0 Then
                rowRP = vTblPR.DefaultView(0).Row
                Return rowRP.No_Transaksi
            End If

            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Sub GetNoPO_SIPByAOPQuick(ByVal vTblPO As Core_Module.DseTrans_CutOffDB_Core.tbl_po_noDataTable, _
    ByVal vTblSIP As Core_Module.DseTrans_CutOffDB_Core.tbl_stockinout_noDataTable, _
    ByVal vPrimaryAOP_Quick As Double, ByRef vNoPO As String, ByRef vNoSIP As String)
        Dim rowPO As Core_Module.DseTrans_CutOffDB_Core.tbl_po_noRow
        Dim rowSIP As Core_Module.DseTrans_CutOffDB_Core.tbl_stockinout_noRow

        Try
            vTblPO.DefaultView.RowFilter = "[Primary Quick Acuan] = " & vPrimaryAOP_Quick
            If vTblPO.DefaultView.Count > 0 Then
                rowPO = vTblPO.DefaultView(0).Row
                vNoPO = rowPO.No_Transaksi
            Else
                vNoPO = "-"
            End If
            vTblSIP.DefaultView.RowFilter = "[Primary Quick Acuan] = " & vPrimaryAOP_Quick
            If vTblSIP.DefaultView.Count > 0 Then
                rowSIP = vTblSIP.DefaultView(0).Row
                vNoSIP = rowSIP.No_Transaksi
            Else
                vNoSIP = "-"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetNoAOP(ByVal vTblAOP As DseTrans_CutOffDB.tbltrans_aop1DataTable, ByVal vPrimaryAOP_Obs As Double) As String
        Dim rowAOP As DseTrans_CutOffDB.tbltrans_aop1Row

        Try
            vTblAOP.DefaultView.RowFilter = "[Primary] = " & vPrimaryAOP_Obs
            If vTblAOP.DefaultView.Count > 0 Then
                rowAOP = vTblAOP.DefaultView(0).Row
                Return rowAOP.No_Transaksi
            End If

            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function GetNoHistoricalPurch(ByVal vTbl As DseTrans_CutOffDB.tbltrans_historicalpurchasebalancemanualDataTable, ByVal vPrimary_Obs As Double, _
    Optional ByRef vTglTrans As Date = #1/1/1950#, Optional ByRef vPayTerm As Integer = 0) As String
        Dim row As DseTrans_CutOffDB.tbltrans_historicalpurchasebalancemanualRow

        Try
            vTbl.DefaultView.RowFilter = "[Primary] = " & vPrimary_Obs
            If vTbl.DefaultView.Count > 0 Then
                row = vTbl.DefaultView(0).Row
                vTglTrans = row.AOP_Date
                vPayTerm = DateDiff(DateInterval.Day, row.Due_Date, row.AOP_Date)
                Return row.AOP_Number
            End If

            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function GetNoHistoricalPurchReturn(ByVal vTbl As DseTrans_CutOffDB.tbltrans_historicalpurchasereturnbalancemanualDataTable, ByVal vPrimary_Obs As Double, _
    Optional ByRef vTglTrans As Date = #1/1/1950#, Optional ByRef vPayTerm As Integer = 0) As String
        Dim row As DseTrans_CutOffDB.tbltrans_historicalpurchasereturnbalancemanualRow

        Try
            vTbl.DefaultView.RowFilter = "[Primary] = " & vPrimary_Obs
            If vTbl.DefaultView.Count > 0 Then
                row = vTbl.DefaultView(0).Row
                vTglTrans = row.AOPR_Date
                vPayTerm = DateDiff(DateInterval.Day, row.Due_Date, row.AOPR_Date)
                Return row.AOPR_Number
            End If

            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function GetNoHistoricalSales(ByVal vTbl As DseTrans_CutOffDB.tbltrans_historicalsalesDataTable, ByVal vPrimary_Obs As Double, _
    Optional ByRef vTglTrans As Date = #1/1/1950#, Optional ByRef vPayTerm As Integer = 0) As String
        Dim row As DseTrans_CutOffDB.tbltrans_historicalsalesRow

        Try
            vTbl.DefaultView.RowFilter = "[Primary] = " & vPrimary_Obs
            If vTbl.DefaultView.Count > 0 Then
                row = vTbl.DefaultView(0).Row
                vTglTrans = row.Invoice_Date
                vPayTerm = DateDiff(DateInterval.Day, row.Due_Date, row.Invoice_Date)
                Return row.Invoice_Number
            End If

            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function GetNoHistoricalSalesReturn(ByVal vTbl As DseTrans_CutOffDB.tbltrans_historicalsalesreturn_manualDataTable, ByVal vPrimary_Obs As Double, _
    Optional ByRef vTglTrans As Date = #1/1/1950#, Optional ByRef vPayTerm As Integer = 0) As String
        Dim row As DseTrans_CutOffDB.tbltrans_historicalsalesreturn_manualRow

        Try
            vTbl.DefaultView.RowFilter = "[Primary] = " & vPrimary_Obs
            If vTbl.DefaultView.Count > 0 Then
                row = vTbl.DefaultView(0).Row
                vTglTrans = row.AOSR_Date
                vPayTerm = DateDiff(DateInterval.Day, row.Due_Date, row.AOSR_Date)
                Return row.AOSR_Number
            End If

            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function GetPrimaryUpdate_BalanceManual(ByVal vTblBM_New As DseTrans_CutOffDB.tbltrans_balance_purchaseDataTable, _
    ByVal vNoTransaksi As String, ByVal vTipe As AccModule.ClsTransaksiSalesBalance.EnumTipeSalesBalance) As Double
        Dim rowBM_New As DseTrans_CutOffDB.tbltrans_balance_purchaseRow

        Try
            'vTblAOPDP_New.DefaultView.RowFilter = Nothing
            vTblBM_New.DefaultView.RowFilter = "[AOP Number Manual] = '" & vNoTransaksi.ToString & "' AND [Tipe] = " & vTipe
            If vTblBM_New.DefaultView.Count > 0 Then
                rowBM_New = vTblBM_New.DefaultView(0).Row
                Return rowBM_New.Primary_AOP
            Else
                'sementara kasih -3 agar mudah di ketahui bila ada kesalahan
                Return -3
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetPrimaryUpdate_BalancePurchase(ByVal vTblBM_New As DseTrans_CutOffDB.tbltrans_balance_purchaseDataTable, _
    ByVal vPrimarySupp As Double, ByVal vPrimaryAOP As Double, ByVal vTipe As AccModule.ClsTransaksiSalesBalance.EnumTipeSalesBalance) As Double
        Dim rowBM_New As DseTrans_CutOffDB.tbltrans_balance_purchaseRow

        Try
            'vTblAOPDP_New.DefaultView.RowFilter = Nothing
            vTblBM_New.DefaultView.RowFilter = "[Primary Supplier] = " & vPrimarySupp & " AND [Primary AOP] = " & vPrimaryAOP & " AND [Tipe] = " & vTipe
            If vTblBM_New.DefaultView.Count > 0 Then
                rowBM_New = vTblBM_New.DefaultView(0).Row
                Return rowBM_New.Primary
            Else
                'sementara kasih -3 agar mudah di ketahui bila ada kesalahan
                Return -3
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetPrimaryMaxJurnal(ByVal ConString As String) As Double
        Dim dta As New DseTrans_CutOffDBTableAdapters.tbltrans_jurnalTableAdapter
        Dim Retval As Nullable(Of Double)

        Try
            dta.Connection.ConnectionString = ConString
            Retval = dta.ScalarQuery_GetPrimaryMaxJurnal()
            If Retval.HasValue = True Then
                Return Retval.Value
            End If

            Return 1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 1
        End Try
    End Function

    Public Sub DeleteDOFromQuickInv(ByRef vTblDo_Out As DseTrans_CutOffDB.tbltrans_doDataTable, ByVal vEndDateCutOff As Date, ByVal ConString As String)
        Dim DtaSalesInv As New DseTrans_CutOffDBTableAdapters.tbltrans_salesinvoiceTableAdapter
        Dim TblSalesInv As New DseTrans_CutOffDB.tbltrans_salesinvoiceDataTable
        Dim rowQuickSalesInv As DseTrans_CutOffDB.tbltrans_salesinvoiceRow
        Dim rowDo_Out As DseTrans_CutOffDB.tbltrans_doRow

        Try
            DtaSalesInv.Connection.ConnectionString = ConString
            DtaSalesInv.FillBy_OutCutOff(TblSalesInv, Sales_Module.ClsTransaksiSalesInvoice.EnumTipeInvoice.SalesInvoice, vEndDateCutOff, 1)
            For Each rowQuickSalesInv In TblSalesInv.Rows
                vTblDo_Out.DefaultView.RowFilter = "[Primary Invoice] = " & rowQuickSalesInv.Primary
                If vTblDo_Out.DefaultView.Count > 0 Then
                    rowDo_Out = vTblDo_Out.DefaultView(0).Row
                    vTblDo_Out.Rows.Remove(rowDo_Out)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DtaSalesInv.Dispose()
            TblSalesInv.Dispose()
        End Try
    End Sub

    Public Function GetPrimaryBalanceOrAcuanPurch_New(ByVal vPrimarySupp As Double, ByVal vPrimaryAOP As Double, _
    ByVal vTipe As Integer, ByVal vTblBalPurch As DseTrans_CutOffDB.tbltrans_balance_purchaseDataTable, _
    Optional ByVal vNoTransAcuan As String = "", Optional ByRef PrimaryBalance As Double = -1) As Double
        Dim rowBalPurch As DseTrans_CutOffDB.tbltrans_balance_purchaseRow

        Try
            If vNoTransAcuan <> "" Then
                vTblBalPurch.DefaultView.RowFilter = "[Primary Supplier] = " & vPrimarySupp.ToString & " AND [AOP Number Manual] = '" & vNoTransAcuan.ToString & "' AND [Tipe] = " & vTipe
                If vTblBalPurch.DefaultView.Count > 0 Then
                    rowBalPurch = vTblBalPurch.DefaultView(0).Row
                    PrimaryBalance = rowBalPurch.Primary
                    Return rowBalPurch.Primary_AOP
                Else
                    Return -1
                End If
            Else
                vTblBalPurch.DefaultView.RowFilter = "[Primary Supplier] = " & vPrimarySupp & " AND [Primary AOP] = " & vPrimaryAOP & " AND [Tipe] = " & vTipe
                If vTblBalPurch.DefaultView.Count > 0 Then
                    rowBalPurch = vTblBalPurch.DefaultView(0).Row
                    Return rowBalPurch.Primary
                Else
                    Return -1
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetPrimaryBalanceOrAcuanSales_New(ByVal vPrimaryCust As Double, ByVal vPrimaryInv As Double, _
    ByVal vTipe As Integer, ByVal vTblBalSales As DseTrans_CutOffDB.tbltrans_balance_sales1DataTable, _
   Optional ByVal vNoTransAcuan As String = "", Optional ByRef PrimaryBalance As Double = -1) As Double
        Dim rowBalSales As DseTrans_CutOffDB.tbltrans_balance_sales1Row

        Try
            If vNoTransAcuan <> "" Then
                vTblBalSales.DefaultView.RowFilter = "[Primary Customer] = " & vPrimaryCust.ToString & " AND [Invoice Number Manual] = '" & vNoTransAcuan.ToString & "' AND [Tipe] = " & vTipe
                If vTblBalSales.DefaultView.Count > 0 Then
                    rowBalSales = vTblBalSales.DefaultView(0).Row
                    PrimaryBalance = rowBalSales.Primary
                    Return rowBalSales.Primary_Invoice
                Else
                    Return -1
                End If
            Else
                vTblBalSales.DefaultView.RowFilter = "[Primary Customer] = " & vPrimaryCust & " AND [Primary Invoice] = " & vPrimaryInv & " AND [Tipe] = " & vTipe
                If vTblBalSales.DefaultView.Count > 0 Then
                    rowBalSales = vTblBalSales.DefaultView(0).Row
                    Return rowBalSales.Primary
                Else
                    Return -1
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Sub PurchasePayment_FilterData(ByVal TblJurnal_Item As DseTrans_CutOffDB.tbltrans_jurnal_Item1DataTable, _
    ByRef TblJurnal_Item_Filter As DseTrans_CutOffDB.tbltrans_jurnal_Item1DataTable, _
    ByVal TblJurnal_ItemOther As DseTrans_CutOffDB.tbltrans_jurnal_ItemOther1DataTable, _
    ByRef TblJurnal_ItemOther_Filter As DseTrans_CutOffDB.tbltrans_jurnal_ItemOther1DataTable, _
    ByVal TblJurnal_ItemPay As DseTrans_CutOffDB.tbltrans_jurnal_ItemPayment1DataTable, _
    ByRef TblJurnal_ItemPay_Filter As DseTrans_CutOffDB.tbltrans_jurnal_ItemPayment1DataTable, _
    ByVal TblBalancePurch As DseTrans_CutOffDB.tbltrans_balance_purchaseDataTable, _
    ByRef TblBalancePurch_Filter As DseTrans_CutOffDB.tbltrans_balance_purchaseDataTable, _
    ByVal TblBalancePurch_New As DseTrans_CutOffDB.tbltrans_balance_purchaseDataTable, _
    ByRef TblBalancePurch_New_Filter As DseTrans_CutOffDB.tbltrans_balance_purchaseDataTable, _
    ByVal TblAOPNo As Core_Module.DseTrans_CutOffDB_Core.tbltrans_aopDataTable, _
    ByRef TblAOPNo_Filter As Core_Module.DseTrans_CutOffDB_Core.tbltrans_aopDataTable, _
    ByVal vPrimarySupplier As Double)

        Try
            TblJurnal_Item_Filter.Clear()
            TblJurnal_Item.DefaultView.RowFilter = "[Primary Pihak Acuan] = " & vPrimarySupplier
            If TblJurnal_Item.DefaultView.Count > 0 Then
                TblJurnal_Item_Filter.Merge(TblJurnal_Item.DefaultView.ToTable)
            End If
            TblJurnal_ItemOther_Filter.Clear()
            TblJurnal_ItemOther.DefaultView.RowFilter = "[Primary Pihak Acuan] = " & vPrimarySupplier
            If TblJurnal_ItemOther.DefaultView.Count > 0 Then
                TblJurnal_ItemOther_Filter.Merge(TblJurnal_ItemOther.DefaultView.ToTable)
            End If
            TblJurnal_ItemPay_Filter.Clear()
            TblJurnal_ItemPay.DefaultView.RowFilter = "[Primary Pihak Acuan] = " & vPrimarySupplier
            If TblJurnal_ItemPay.DefaultView.Count > 0 Then
                TblJurnal_ItemPay_Filter.Merge(TblJurnal_ItemPay.DefaultView.ToTable)
            End If
            TblBalancePurch_Filter.Clear()
            TblBalancePurch.DefaultView.RowFilter = "[Primary Supplier] = " & vPrimarySupplier
            If TblBalancePurch.DefaultView.Count > 0 Then
                TblBalancePurch_Filter.Merge(TblBalancePurch.DefaultView.ToTable)
            End If
            TblBalancePurch_New_Filter.Clear()
            TblBalancePurch_New.DefaultView.RowFilter = "[Primary Supplier] = " & vPrimarySupplier
            If TblBalancePurch_New.DefaultView.Count > 0 Then
                TblBalancePurch_New_Filter.Merge(TblBalancePurch_New.DefaultView.ToTable)
            End If
            TblAOPNo_Filter.Clear()
            TblAOPNo.DefaultView.RowFilter = "[Primary Supplier] = " & vPrimarySupplier
            If TblAOPNo.DefaultView.Count > 0 Then
                TblAOPNo_Filter.Merge(TblAOPNo.DefaultView.ToTable)
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SalesReceivement_FilterData(ByVal TblJurnal_Item As DseTrans_CutOffDB.tbltrans_jurnal_Item1DataTable, _
    ByRef TblJurnal_Item_Filter As DseTrans_CutOffDB.tbltrans_jurnal_Item1DataTable, _
    ByVal TblJurnal_ItemOther As DseTrans_CutOffDB.tbltrans_jurnal_ItemOther1DataTable, _
    ByRef TblJurnal_ItemOther_Filter As DseTrans_CutOffDB.tbltrans_jurnal_ItemOther1DataTable, _
    ByVal TblJurnal_ItemPay As DseTrans_CutOffDB.tbltrans_jurnal_ItemPayment1DataTable, _
    ByRef TblJurnal_ItemPay_Filter As DseTrans_CutOffDB.tbltrans_jurnal_ItemPayment1DataTable, _
    ByVal TblBalanceSales As DseTrans_CutOffDB.tbltrans_balance_sales1DataTable, _
    ByRef TblBalanceSales_Filter As DseTrans_CutOffDB.tbltrans_balance_sales1DataTable, _
    ByVal TblBalanceSales_New As DseTrans_CutOffDB.tbltrans_balance_sales1DataTable, _
    ByRef TblBalanceSales_New_Filter As DseTrans_CutOffDB.tbltrans_balance_sales1DataTable, _
    ByVal TblInvNo As Core_Module.DseTrans_CutOffDB_Core.tbl_InvNoDataTable, _
    ByRef TblInvNo_Filter As Core_Module.DseTrans_CutOffDB_Core.tbl_InvNoDataTable, _
    ByVal vPrimaryCustomer As Double)

        Try
            TblJurnal_Item_Filter.Clear()
            TblJurnal_Item.DefaultView.RowFilter = "[Primary Pihak Acuan] = " & vPrimaryCustomer
            If TblJurnal_Item.DefaultView.Count > 0 Then
                TblJurnal_Item_Filter.Merge(TblJurnal_Item.DefaultView.ToTable)
            End If
            TblJurnal_ItemOther_Filter.Clear()
            TblJurnal_ItemOther.DefaultView.RowFilter = "[Primary Pihak Acuan] = " & vPrimaryCustomer
            If TblJurnal_ItemOther.DefaultView.Count > 0 Then
                TblJurnal_ItemOther_Filter.Merge(TblJurnal_ItemOther.DefaultView.ToTable)
            End If
            TblJurnal_ItemPay_Filter.Clear()
            TblJurnal_ItemPay.DefaultView.RowFilter = "[Primary Pihak Acuan] = " & vPrimaryCustomer
            If TblJurnal_ItemPay.DefaultView.Count > 0 Then
                TblJurnal_ItemPay_Filter.Merge(TblJurnal_ItemPay.DefaultView.ToTable)
            End If
            TblBalanceSales_Filter.Clear()
            TblBalanceSales.DefaultView.RowFilter = "[Primary Customer] = " & vPrimaryCustomer
            If TblBalanceSales.DefaultView.Count > 0 Then
                TblBalanceSales_Filter.Merge(TblBalanceSales.DefaultView.ToTable)
            End If
            TblBalanceSales_New_Filter.Clear()
            TblBalanceSales_New.DefaultView.RowFilter = "[Primary Customer] = " & vPrimaryCustomer
            If TblBalanceSales_New.DefaultView.Count > 0 Then
                TblBalanceSales_New_Filter.Merge(TblBalanceSales_New.DefaultView.ToTable)
            End If
            TblInvNo_Filter.Clear()
            TblInvNo.DefaultView.RowFilter = "[Primary Customer] = " & vPrimaryCustomer
            If TblInvNo.DefaultView.Count > 0 Then
                TblInvNo_Filter.Merge(TblInvNo.DefaultView.ToTable)
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
