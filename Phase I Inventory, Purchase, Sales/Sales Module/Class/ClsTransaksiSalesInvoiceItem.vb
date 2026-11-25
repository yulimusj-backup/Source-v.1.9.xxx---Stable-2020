Option Explicit On
Option Strict Off

Public Class ClsTransaksiSalesInvoiceItem
    Public Function GetAliasCode(ByVal vPrimaryProduct As Double, ByVal vPrimaryCust As Double) As String
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.TblAliasTableAdapter
        Dim TBL As New DseTransaksiSalesInvoice.TblAliasDataTable
        Dim row As DseTransaksiSalesInvoice.TblAliasRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, vPrimaryProduct, vPrimaryCust)
            If TBL.Rows.Count = 0 Then
                Return ""
            Else
                row = TBL.Rows(TBL.Rows.Count - 1)
                Return row.Alias_Code
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function GetAliasDesc(ByVal vPrimaryProduct As Double, ByVal vPrimaryCust As Double) As String
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.TblAliasTableAdapter
        Dim TBL As New DseTransaksiSalesInvoice.TblAliasDataTable
        Dim row As DseTransaksiSalesInvoice.TblAliasRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, vPrimaryProduct, vPrimaryCust)
            If TBL.Rows.Count = 0 Then
                Return ""
            Else
                row = TBL.Rows(TBL.Rows.Count - 1)
                Return row.Alias_Desc
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function IsiTBLDO(ByVal vPrimaryInvoice As Double, ByVal TBL As DseTransaksiSalesInvoice.DataTable7DataTable) As Boolean
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.DataTable7TableAdapter
        Try
            DTA.Fill(TBL, vPrimaryInvoice)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml_Default(ByVal vPrimaryInv As Double, ByVal vPrimaryProduct As Double, ByVal vBatchNumber As String, ByVal vPrimaryLocation As Double) As Decimal
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_itemTableAdapter

        Try
            If DTA.ScalarQuery_GetJml(vPrimaryProduct, vPrimaryInv, vBatchNumber, vPrimaryLocation).HasValue = True Then
                Return DTA.ScalarQuery_GetJml(vPrimaryProduct, vPrimaryInv, vBatchNumber, vPrimaryLocation).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal vPrimarySo As Double, ByVal vPrimaryProduct As Double) As Integer
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_itemTableAdapter
        Try
            Return DTA.ScalarQuery_GetJmlByPrimarySO_n_Product(vPrimaryProduct, vPrimarySo).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlBySR(ByVal vPrimarySR As Double) As Double
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_item_salesreturnTableAdapter

        Try
            Return DTA.ScalarQuery_GetJmlByPrimarySR(vPrimarySR).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlByDO(ByVal PrimaryDO As Double) As Integer
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_itemTableAdapter
        Try
            Return DTA.ScalarQuery_GetJmlByPrimaryDO(PrimaryDO).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetSalesInvoiceNoBySR(ByVal vPrimarySR As Double) As String
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_item_salesreturnTableAdapter

        Try
            Return DTA.ScalarQuery_GetInvoiceNo(vPrimarySR)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetSalesInvoiceNo(ByVal vPrimaryDo As Double) As String
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_itemTableAdapter

        Try
            Return DTA.ScalarQuery_GetNoSalesInvoice(vPrimaryDo)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDGV(ByVal vPrimaryTrans As Double, _
            ByVal vTblOldSalesInvoiceItem As CustomSMESS.DseTransferBeginingBalance.tbltrans_salesinvoice_item_copyDataTable, _
            ByVal vTblOldSalesInvoiceItemExchangeRate As CustomSMESS.DseTransferBeginingBalance.tbltrans_salesinvoice_item_exchangerate_copyDataTable, _
            ByVal vTblOldSalesInvoiceItemSalesReturn As CustomSMESS.DseTransferBeginingBalance.tbltrans_salesinvoice_item_salesreturn_copyDataTable, _
            ByVal vTblOldSalesInvoiceItemInvoiceDp As CustomSMESS.DseTransferBeginingBalance.tbltrans_salesinvoice_itemdp_copyDataTable, _
            ByVal vDGVItem As DataGridView, ByVal vDGVExchangeRate As DataGridView, _
            ByVal vDGVDP As DataGridView, ByVal vCalculatedBy As ClsTransaksiSalesInvoice.EnumCalculatesBy, _
            ByVal vDGVSR As DataGridView, ByVal vFrm As FrmTransaksiSalesInvoiceQuick) As Boolean
        Dim objProd As New Core_Module.clsMasterProduct

        Dim rowOldSalesInvoiceItem As CustomSMESS.DseTransferBeginingBalance.tbltrans_salesinvoice_item_copyRow
        Dim rowOldSalesInvoiceItemExchangeRate As CustomSMESS.DseTransferBeginingBalance.tbltrans_salesinvoice_item_exchangerate_copyRow
        Dim rowOldSalesInvoiceItemSalesReturn As CustomSMESS.DseTransferBeginingBalance.tbltrans_salesinvoice_item_salesreturn_copyRow
        Dim rowOldSalesInvoiceItemInvoiceDp As CustomSMESS.DseTransferBeginingBalance.tbltrans_salesinvoice_itemdp_copyRow

        Dim ObjCD As New Core_Module.ClsCekData
        Dim rowDGV As DataGridViewRow

        Dim PrimarySO, PrimaryDO, PrimaryStockOut As Double
        Dim NoDO, NoSO, CustPONumber, NoStockOutSales As String
        Dim ObjSO As New ClsTransaksiSalesOrder
        Dim ObjDO As New ClsTransaksiDO
        Dim ObjInOut As New Inventory_Module.clsTransaksiStockInOut
        Dim DTASO As New DseTransaksiSOTableAdapters.tbltrans_soTableAdapter
        Dim DTADO As New DseTransaksiDOTableAdapters.tbltrans_do3TableAdapter

        Try
            'kosongkan dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            vDGVItem.Rows.Clear()
            vDGVExchangeRate.Rows.Clear()
            vDGVDP.Rows.Clear()
            vDGVSR.Rows.Clear()

            'isikan tabel item
            PrimarySO = -1
            PrimaryDO = -1
            PrimaryStockOut = -1
            NoDO = ""
            NoSO = ""
            CustPONumber = ""
            NoStockOutSales = ""

            vTblOldSalesInvoiceItem.DefaultView.RowFilter = "[Primary Transaksi] = " & vPrimaryTrans.ToString
            For i As Integer = 0 To vTblOldSalesInvoiceItem.DefaultView.Count - 1
                rowOldSalesInvoiceItem = vTblOldSalesInvoiceItem.DefaultView(i).Row
                With rowOldSalesInvoiceItem
                    If .Primary_DO <> PrimaryDO Then
                        PrimaryDO = .Primary_DO
                        NoDO = ObjDO.GetDONO(PrimaryDO, DTADO)
                        NoStockOutSales = ObjInOut.GetNoTransaksiByPrimaryDOLossOrPrimaryDOSales(PrimaryDO, Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Sales)
                    End If
                    If .Primary_SO <> PrimarySO Then
                        PrimarySO = .Primary_SO
                        NoSO = ObjSO.GetNoSO(PrimarySO, DTASO)
                        CustPONumber = ObjSO.GetCustPONo(PrimarySO, DTASO)
                    End If
                    '.SO_Number = NoSO
                    '.Cust_PO_No = CustPONumber
                    '.Do_Number = NoDO
                    '.Stock_Out_No = NoStockOutSales

                    vDGVItem.Rows.Add()
                    rowDGV = vDGVItem.Rows(vDGVItem.RowCount - 1)
                    rowDGV.Cells(vFrm.ColNo.Index).Value = "0"
                    rowDGV.Cells(vFrm.ColCodeNo.Index).Value = objProd.GetProductCode(.Primary_Product)
                    'rowDGV.Cells(vFrm.ColDescription.Index).Value = .Description
                    'rowDGV.Cells(vFrm.ColCustCodeNo.Index).Value = ._Customer_s_Code_No
                    'rowDGV.Cells(vFrm.ColCustDesc.Index).Value = ._Customer_s_Desc
                    'rowDGV.Cells(vFrm.ColSONO.Index).Value = .SO_Number
                    'rowDGV.Cells(vFrm.ColCustPONo.Index).Value = .Cust_PO_No
                    'rowDGV.Cells(vFrm.ColDONO.Index).Value = .Do_Number
                    'rowDGV.Cells(vFrm.ColStockOutNo.Index).Value = .Stock_Out_No
                    rowDGV.Cells(vFrm.ColQtyMaster.Index).Value = .Qty_Master
                    rowDGV.Cells(vFrm.ColQtyPerMaster.Index).Value = .Qty_Per_Master
                    'rowDGV.Cells(vFrm.ColMasterUnits.Index).Value = .Master_Unit_Code
                    rowDGV.Cells(vFrm.ColQty.Index).Value = ObjCD.DigitGrouping(.Qty) ', .AngkaDecimalProduct)
                    rowDGV.Cells(vFrm.ColTotalQty.Index).Value = ObjCD.DigitGrouping(.Total_Qty) ', .AngkaDecimalProduct)
                    'rowDGV.Cells(vFrm.ColUnit.Index).Value = .Unit_Code
                    rowDGV.Cells(vFrm.ColSecQty.Index).Value = ObjCD.DigitGrouping(.Secondary_Qty)
                    'rowDGV.Cells(vFrm.ColSecUnit.Index).Value = .Secondary_Unit_Code
                    'rowDGV.Cells(vFrm.ColCurrency.Index).Value = .Currency_Code
                    rowDGV.Cells(vFrm.ColUnitPrice.Index).Value = ObjCD.DigitGrouping(.Unit_Price) ', .Default_Prec_Currency)
                    rowDGV.Cells(vFrm.ColDiscount.Index).Value = ObjCD.DigitGrouping(.Discount, 2)
                    rowDGV.Cells(vFrm.ColNettPrice.Index).Value = ObjCD.DigitGrouping(.Nett_Price) ', .Default_Prec_Currency)
                    rowDGV.Cells(vFrm.ColAmount.Index).Value = ObjCD.DigitGrouping(.Amount) ', .Default_Prec_Currency)
                    rowDGV.Cells(vFrm.ColRemarksItem.Index).Value = .Remarks
                    rowDGV.Cells(vFrm.ColPrimaryCodeNo.Index).Value = .Primary_Product
                    rowDGV.Cells(vFrm.ColPrimarySO.Index).Value = .Primary_SO
                    rowDGV.Cells(vFrm.ColPrimaryDO.Index).Value = .Primary_DO
                    rowDGV.Cells(vFrm.ColPrimaryMasterUnits.Index).Value = .Primary_Master_Unit
                    rowDGV.Cells(vFrm.ColPrimarySecUnit.Index).Value = .Primary_Secondary_Units
                    rowDGV.Cells(vFrm.ColPrimaryUnit.Index).Value = .Primary_Units
                    rowDGV.Cells(vFrm.ColPrimaryCurrency.Index).Value = .Primary_Currency
                    rowDGV.Cells(vFrm.ColLotNumber.Index).Value = .Batch_Number
                    'If .Location = "" Then
                    '    .Location = "-"
                    'End If
                    rowDGV.Cells(vFrm.ColPrimaryLocation.Index).Value = .Primary_Location
                    'rowDGV.Cells(vFrm.ColLocationCode.Index).Value = .Location
                    'If .Expiry = #1/1/1950# Then
                    rowDGV.Cells(vFrm.ColExpiry.Index).Value = "-"
                    'Else
                    '    rowDGV.Cells(vFrm.ColExpiry.Index).Value = Format(.Expiry, "dd MMM yyyy")
                    'End If
                    rowDGV.Cells(vFrm.ColAliasDesc.Index).Value = .Alias_Desc
                    rowDGV.Cells(vFrm.ColAliasCode.Index).Value = .Alias_Code
                End With
            Next

            'isikan tabel item dp
            vTblOldSalesInvoiceItemInvoiceDp.DefaultView.RowFilter = "[Primary Sales Invoice] = " & vPrimaryTrans.ToString
            For i As Integer = 0 To vTblOldSalesInvoiceItemInvoiceDp.DefaultView.Count - 1
                rowOldSalesInvoiceItemInvoiceDp = vTblOldSalesInvoiceItemInvoiceDp.DefaultView(i).Row
                With rowOldSalesInvoiceItemInvoiceDp
                    vDGVDP.Rows.Add("0", "", "", ObjCD.DigitGrouping(CStr(0), 0), .Remarks, .Primary_Invoice_DP)
                End With
            Next

            'isikan tabel item exchange rate
            vTblOldSalesInvoiceItemExchangeRate.DefaultView.RowFilter = "[Primary Sales Invoice] = " & vPrimaryTrans.ToString
            For i As Integer = 0 To vTblOldSalesInvoiceItemExchangeRate.DefaultView.Count - 1
                rowOldSalesInvoiceItemExchangeRate = vTblOldSalesInvoiceItemExchangeRate.DefaultView(i).Row
                With rowOldSalesInvoiceItemExchangeRate
                    vDGVExchangeRate.Rows.Add("0", "", "", .Exhange_Rate, .Remarks, .Primary_Transaksi_Rujukan, .Primary_Currency)
                End With
            Next

            'isikan dgv SR
            vTblOldSalesInvoiceItemSalesReturn.DefaultView.RowFilter = "[Primary Sales Invoice] = " & vPrimaryTrans.ToString
            For i As Integer = 0 To vTblOldSalesInvoiceItemSalesReturn.DefaultView.Count - 1
                rowOldSalesInvoiceItemSalesReturn = vTblOldSalesInvoiceItemSalesReturn.DefaultView(i).Row
                With rowOldSalesInvoiceItemSalesReturn
                    vDGVSR.Rows.Add("", "", "", .Remarks, .Primary_SR)
                End With
            Next

            'set nomer
            ObjCD.SetNumber(vDGVItem)
            ObjCD.SetNumber(vDGVDP)
            ObjCD.SetNumber(vDGVExchangeRate)
            ObjCD.SetNumber(vDGVSR)

            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function IsiDGV(ByVal vDGVItem As DataGridView, ByVal vDGVExchangeRate As DataGridView, _
  ByVal vDGVDP As DataGridView, ByVal vPrimaryTransaksi As Double, ByVal vCalculatedBy As _
  ClsTransaksiSalesInvoice.EnumCalculatesBy, ByVal vDGVSR As DataGridView, ByVal vFrm As FrmTransaksiSalesInvoiceQuick) As Boolean
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_item_view1TableAdapter
        Dim DTA1 As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_itemdp_viewTableAdapter
        Dim DTA2 As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoiceitem_salesreturn_view0TableAdapter
        Dim DTA3 As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_add_dataTableAdapter
        Dim DTA4 As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_item_costTableAdapter

        Dim TBL As New DseTransaksiSalesInvoice.tbltrans_salesinvoice_item_view1DataTable
        Dim TBL1 As New DseTransaksiSalesInvoice.tbltrans_salesinvoice_itemdp_viewDataTable
        Dim TBL2 As New DseTransaksiSalesInvoice.tbltrans_salesinvoiceitem_salesreturn_view0DataTable
        Dim TBL3 As New DseTransaksiSalesInvoice.tbltrans_salesinvoice_add_dataDataTable
        Dim TBL4 As New DseTransaksiSalesInvoice.tbltrans_salesinvoice_item_costDataTable

        Dim row As DseTransaksiSalesInvoice.tbltrans_salesinvoice_item_view1Row
        Dim row1 As DseTransaksiSalesInvoice.tbltrans_salesinvoice_itemdp_viewRow
        Dim row2 As DseTransaksiSalesInvoice.tbltrans_salesinvoiceitem_salesreturn_view0Row
        Dim row3 As DseTransaksiSalesInvoice.tbltrans_salesinvoice_add_dataRow
        Dim row4 As DseTransaksiSalesInvoice.tbltrans_salesinvoice_item_costRow

        Dim ObjCD As New Core_Module.ClsCekData
        Dim rowDGV As DataGridViewRow

        Dim PrimarySO, PrimaryDO, PrimaryStockOut As Double
        Dim NoDO, NoSO, CustPONumber, NoStockOutSales As String
        Dim ObjSO As New ClsTransaksiSalesOrder
        Dim ObjDO As New ClsTransaksiDO
        Dim ObjInOut As New Inventory_Module.clsTransaksiStockInOut
        Dim DTASO As New DseTransaksiSOTableAdapters.tbltrans_soTableAdapter
        Dim DTADO As New DseTransaksiDOTableAdapters.tbltrans_do3TableAdapter
        Dim DefPrecMulti As Integer = 0
        Dim DefPrecDefault As Integer = 0
        Dim ObjP As New Core_Module.clsMasterProduct

        Try
            'kosongkan dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            vDGVItem.Rows.Clear()
            vDGVExchangeRate.Rows.Clear()
            vDGVDP.Rows.Clear()
            vDGVSR.Rows.Clear()

            'isikan tabel item
            PrimarySO = -1
            PrimaryDO = -1
            PrimaryStockOut = -1
            NoDO = ""
            NoSO = ""
            CustPONumber = ""
            NoStockOutSales = ""

            If DTA.Fill(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    With row
                        If .Primary_DO <> PrimaryDO Then
                            PrimaryDO = .Primary_DO
                            NoDO = ObjDO.GetDONO(PrimaryDO, DTADO)
                            NoStockOutSales = ObjInOut.GetNoTransaksiByPrimaryDOLossOrPrimaryDOSales(PrimaryDO, Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Sales)
                        End If
                        If .Primary_SO <> PrimarySO Then
                            PrimarySO = .Primary_SO
                            NoSO = ObjSO.GetNoSO(PrimarySO, DTASO)
                            CustPONumber = ObjSO.GetCustPONo(PrimarySO, DTASO)
                        End If
                        .SO_Number = NoSO
                        .Cust_PO_No = CustPONumber
                        .Do_Number = NoDO
                        .Stock_Out_No = NoStockOutSales

                        vDGVItem.Rows.Add()
                        rowDGV = vDGVItem.Rows(vDGVItem.RowCount - 1)
                        rowDGV.Cells(vFrm.ColNo.Index).Value = "0"
                        rowDGV.Cells(vFrm.ColCodeNo.Index).Value = .Product_Code
                        rowDGV.Cells(vFrm.ColDescription.Index).Value = .Description
                        rowDGV.Cells(vFrm.ColCustCodeNo.Index).Value = .Customer_s_Code_No
                        rowDGV.Cells(vFrm.ColCustDesc.Index).Value = .Customer_s_Desc
                        rowDGV.Cells(vFrm.ColSONO.Index).Value = .SO_Number
                        rowDGV.Cells(vFrm.ColCustPONo.Index).Value = .Cust_PO_No
                        rowDGV.Cells(vFrm.ColDONO.Index).Value = .Do_Number
                        rowDGV.Cells(vFrm.ColStockOutNo.Index).Value = .Stock_Out_No
                        rowDGV.Cells(vFrm.ColQtyMaster.Index).Value = .Qty_Master
                        rowDGV.Cells(vFrm.ColQtyPerMaster.Index).Value = .Qty_Per_Master
                        rowDGV.Cells(vFrm.ColMasterUnits.Index).Value = .Master_Unit_Code
                        rowDGV.Cells(vFrm.ColQty.Index).Value = ObjCD.DigitGrouping(.Qty, .AngkaDecimalProduct)
                        rowDGV.Cells(vFrm.ColTotalQty.Index).Value = ObjCD.DigitGrouping(.Total_Qty, .AngkaDecimalProduct)
                        rowDGV.Cells(vFrm.ColUnit.Index).Value = .Unit_Code
                        rowDGV.Cells(vFrm.ColSecQty.Index).Value = ObjCD.DigitGrouping(.Secondary_Qty)
                        rowDGV.Cells(vFrm.ColSecUnit.Index).Value = .Secondary_Unit_Code
                        rowDGV.Cells(vFrm.ColCurrency.Index).Value = .Currency_Code
                        rowDGV.Cells(vFrm.ColUnitPrice.Index).Value = ObjCD.DigitGrouping(.Unit_Price, .Default_Prec_Currency)
                        rowDGV.Cells(vFrm.ColDiscount.Index).Value = ObjCD.DigitGrouping(.Discount, 2)
                        rowDGV.Cells(vFrm.ColDiscAmount.Index).Value = ObjCD.DigitGrouping(.Disc_Amount, .Default_Prec_Currency)
                        rowDGV.Cells(vFrm.ColDiscFromPrect.Index).Value = .DiscFromPerc
                        rowDGV.Cells(vFrm.ColNettPrice.Index).Value = ObjCD.DigitGrouping(.Nett_Price, .Default_Prec_Currency)
                        rowDGV.Cells(vFrm.ColAmount.Index).Value = ObjCD.DigitGrouping(.Amount, .Default_Prec_Currency)
                        rowDGV.Cells(vFrm.ColRemarksItem.Index).Value = .Remarks
                        rowDGV.Cells(vFrm.ColPrimaryCodeNo.Index).Value = .Primary_Product
                        rowDGV.Cells(vFrm.ColPrimarySO.Index).Value = .Primary_SO
                        rowDGV.Cells(vFrm.ColPrimaryDO.Index).Value = .Primary_DO
                        rowDGV.Cells(vFrm.ColPrimaryMasterUnits.Index).Value = .Primary_Master_Unit
                        rowDGV.Cells(vFrm.ColPrimarySecUnit.Index).Value = .Primary_Secondary_Units
                        rowDGV.Cells(vFrm.ColPrimaryUnit.Index).Value = .Primary_Units
                        rowDGV.Cells(vFrm.ColPrimaryCurrency.Index).Value = .Primary_Currency
                        rowDGV.Cells(vFrm.ColLotNumber.Index).Value = .Batch_Number
                        rowDGV.Cells(vFrm.ColVAT.Index).Value = ObjCD.digitgrouping(.VAT, .Default_Prec_Currency)
                        rowDGV.Cells(vFrm.ColVAT2.Index).Value = ObjCD.digitgrouping(.VAT2, .Default_Prec_Currency)
                        rowDGV.Cells(vFrm.ColAmount2.Index).Value = ObjCD.digitgrouping(.Amount2, .Default_Prec_Currency)

                        If .Location = "" Then
                            .Location = "-"
                        End If
                        rowDGV.Cells(vFrm.ColPrimaryLocation.Index).Value = .Primary_Location
                        rowDGV.Cells(vFrm.ColLocationCode.Index).Value = .Location
                        If .Expiry = #1/1/1950# Then
                            rowDGV.Cells(vFrm.ColExpiry.Index).Value = "-"
                        Else
                            rowDGV.Cells(vFrm.ColExpiry.Index).Value = Format(.Expiry, "dd MMM yyyy")
                        End If
                        rowDGV.Cells(vFrm.ColAliasDesc.Index).Value = .Alias_Desc
                        rowDGV.Cells(vFrm.ColAliasCode.Index).Value = .Alias_Code

                        DefPrecMulti = ObjP.GetAngkaDecimal(.Primary_Product, .Primary_Units_Multi)
                        DefPrecDefault = ObjP.GetAngkaDecimal(.Primary_Product)
                        rowDGV.Cells(vFrm.ColTotalQtyMulti.Index).Value = ObjCD.DigitGrouping(.Total_Qty_Multi, DefPrecMulti)
                        rowDGV.Cells(vFrm.ColUnitsMulti.Index).Value = .Units_Code_Multi
                        rowDGV.Cells(vFrm.ColPrimaryUnitsMulti.Index).Value = .Primary_Units_Multi
                        rowDGV.Cells(vFrm.ColRatioReference.Index).Value = ObjCD.DigitGrouping(.Ratio_Reference, DefPrecMulti)
                        rowDGV.Cells(vFrm.ColRatioDefault.Index).Value = ObjCD.DigitGrouping(.Ratio_Default, DefPrecDefault)
                        If .Primary_Units <> .Primary_Units_Multi Then
                            rowDGV.Cells(vFrm.ColRatioDesc.Index).Value = ObjCD.DigitGrouping(.Ratio_Reference, DefPrecMulti) & " " & .Units_Code_Multi & "=" & ObjCD.DigitGrouping(.Ratio_Default, DefPrecDefault) & " " & .Unit_Code
                        Else
                            rowDGV.Cells(vFrm.ColRatioDesc.Index).Value = "-"
                        End If
                        'rowDGV.Cells(vFrm.ColPrimarySalesPerson.Index).Value = .primary_sales_person
                        'rowDGV.Cells(vFrm.ColSalesPerson.Index).Value = .sales_person_code & " - " & .sales_person_desc

                        rowDGV.Cells(vFrm.ColPriceByMultiUnit.Index).Value = .price_by_multi_unit
                        If .price_by_multi_unit = 1 Then
                            rowDGV.Cells(vFrm.ColPriceByMultiUnit_Str.Index).Value = "/" + rowDGV.Cells(vFrm.ColUnitsMulti.Index).Value
                        Else
                            rowDGV.Cells(vFrm.ColPriceByMultiUnit_Str.Index).Value = "/" + rowDGV.Cells(vFrm.ColUnit.Index).Value
                        End If
                    End With
                Next
            End If

            'isikan tabel item dp
            If DTA1.FillBy_Primary(TBL1, vPrimaryTransaksi) > 0 Then
                For Each row1 In TBL1.Rows
                    With row1
                        vDGVDP.Rows.Add("0", .No_Invoice_DP, .Currency_Code, ObjCD.DigitGrouping(.Sub_Total, .Default_Precision), .Remarks, _
                        .Primary_Invoice_DP)
                    End With
                Next
            End If

            'isikan tabel item exchange rate
            If vCalculatedBy = ClsTransaksiSalesInvoice.EnumCalculatesBy.Invoice Then
                Dim DTAEX As New DseTransaksiSalesInvoiceTableAdapters.DataTable2TableAdapter
                Dim TBLEX As New DseTransaksiSalesInvoice.DataTable2DataTable
                Dim RowEX As DseTransaksiSalesInvoice.DataTable2Row
                If DTAEX.Fill(TBLEX, vPrimaryTransaksi) > 0 Then
                    For Each RowEX In TBLEX.Rows
                        With RowEX
                            vDGVExchangeRate.Rows.Add("0", .Sales_Invoice_No, .Currency_Code, _
                            .Exhange_Rate, .Remarks, .Primary_Transaksi_Rujukan, .Primary_Currency)
                        End With
                    Next
                End If
                DTAEX.Dispose()
                TBLEX.Dispose()
            ElseIf vCalculatedBy = ClsTransaksiSalesInvoice.EnumCalculatesBy.Delivery_Order Then
                Dim DTAEX As New DseTransaksiSalesInvoiceTableAdapters.DataTable3TableAdapter
                Dim TBLEX As New DseTransaksiSalesInvoice.DataTable3DataTable
                Dim RowEX As DseTransaksiSalesInvoice.DataTable3Row
                If DTAEX.Fill(TBLEX, vPrimaryTransaksi) > 0 Then
                    For Each RowEX In TBLEX.Rows
                        With RowEX
                            vDGVExchangeRate.Rows.Add("0", .Delivery_Order_No, .Currency_Code, _
                            .Exhange_Rate, .Remarks, .Primary_Transaksi_Rujukan, .Primary_Currency)
                        End With
                    Next
                End If
                DTAEX.Dispose()
                TBLEX.Dispose()
            ElseIf vCalculatedBy = ClsTransaksiSalesInvoice.EnumCalculatesBy.Sales_Order Then
                Dim DTAEX As New DseTransaksiSalesInvoiceTableAdapters.DataTable4TableAdapter
                Dim TBLEX As New DseTransaksiSalesInvoice.DataTable4DataTable
                Dim RowEX As DseTransaksiSalesInvoice.DataTable4Row
                If DTAEX.Fill(TBLEX, vPrimaryTransaksi) > 0 Then
                    For Each RowEX In TBLEX.Rows
                        With RowEX
                            vDGVExchangeRate.Rows.Add("0", .Sales_Order_No, .Currency_Code, _
                            .Exhange_Rate, .Remarks, .Primary_Transaksi_Rujukan, .Primary_Currency)
                        End With
                    Next
                End If
                DTAEX.Dispose()
                TBLEX.Dispose()
            End If

            'isikan dgv SR
            DTA2.FillBy_PrimarySalesInvoice(TBL2, vPrimaryTransaksi)
            For Each row2 In TBL2.Rows
                vDGVSR.Rows.Add("", row2.Sales_Return_No, row2.Delivery_Order_No, row2.Remarks, row2.Primary_SR)
            Next

            vFrm.TBLHeaderCost.Rows.Clear()
            If DTA3.FillBy(TBL3, vPrimaryTransaksi) > 0 Then
                For Each row3 In TBL3.Rows
                    With row3
                        vFrm.rowHeaderCost = vFrm.TBLHeaderCost.NewRow
                        vFrm.rowHeaderCost.primary = .primary
                        vFrm.rowHeaderCost.primary_cust_ship_to = .primary_cust_ship_to
                        vFrm.rowHeaderCost.primary_inv = .primary_inv
                        vFrm.rowHeaderCost.primary_volume_unit = .primary_volume_unit
                        vFrm.rowHeaderCost.primary_weight_unit = .primary_weight_unit
                        vFrm.rowHeaderCost.ETA_Date = .ETA_Date
                        vFrm.rowHeaderCost.ETA_Port = .ETA_Port
                        vFrm.rowHeaderCost.ETD_Date = .ETD_Date
                        vFrm.rowHeaderCost.ETD_Port = .ETD_Port
                        vFrm.rowHeaderCost.identifier_cust_ship_to = .identifier_cust_ship_to
                        vFrm.rowHeaderCost.invoice_type = .invoice_type
                        vFrm.rowHeaderCost.shipping_mark = .shipping_mark
                        vFrm.rowHeaderCost.Total_Gross_Volume = .Total_Gross_Volume
                        vFrm.rowHeaderCost.Total_Gross_Weight = .Total_Gross_Weight
                        vFrm.rowHeaderCost.Total_Nett_Volume = .Total_Nett_Volume
                        vFrm.rowHeaderCost.Total_Nett_Weight = .Total_Nett_Weight
                        vFrm.rowHeaderCost.transportation_mode = .transportation_mode
                        vFrm.rowHeaderCost.var1_name = .var1_name
                        vFrm.rowHeaderCost.var1_value = .var1_value
                        vFrm.rowHeaderCost.var2_name = .var2_name
                        vFrm.rowHeaderCost.var2_value = .var2_value
                        vFrm.rowHeaderCost.var3_name = .var3_name
                        vFrm.rowHeaderCost.var3_value = .var3_value
                        vFrm.rowHeaderCost.var4_name = .var4_name
                        vFrm.rowHeaderCost.var4_value = .var4_value
                        vFrm.rowHeaderCost.var5_name = .var5_name
                        vFrm.rowHeaderCost.var5_value = .var5_value
                        vFrm.TBLHeaderCost.Rows.Add(vFrm.rowHeaderCost)
                    End With
                Next
            End If

            vFrm.TBLitemCost.Rows.Clear()
            If DTA4.Fill(TBL4, vPrimaryTransaksi) > 0 Then
                For Each row4 In TBL4.Rows
                    With row4
                        vFrm.rowitemCost = vFrm.TBLitemCost.NewRow
                        vFrm.rowitemCost.Primary = .Primary
                        vFrm.rowitemCost.Primary_Akun = .Primary_Akun
                        vFrm.rowitemCost.Primary_Inv = .Primary_Inv
                        vFrm.rowitemCost.Remarks = .Remarks
                        vFrm.rowitemCost.Cost = .Cost
                        vFrm.rowitemCost.Cost_Description = .Cost_Description
                        vFrm.TBLitemCost.Rows.Add(vFrm.rowitemCost)
                    End With
                Next
            End If


            'set nomer
            ObjCD.SetNumber(vDGVItem)
            ObjCD.SetNumber(vDGVDP)
            ObjCD.SetNumber(vDGVExchangeRate)
            ObjCD.SetNumber(vDGVSR)

            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
            ObjCD = Nothing
            ObjSO = Nothing
            ObjDO = Nothing
            ObjInOut = Nothing
            DTASO.Dispose()
            DTADO.Dispose()
        End Try
    End Function

    Public Function IsiDGV(ByVal vDGVItem As DataGridView, ByVal vDGVExchangeRate As DataGridView, _
    ByVal vDGVDP As DataGridView, ByVal vPrimaryTransaksi As Double, ByVal vCalculatedBy As _
    ClsTransaksiSalesInvoice.EnumCalculatesBy, ByVal vDGVSR As DataGridView, ByVal vFrm As FrmTransaksiSalesInvoice) As Boolean
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_item_view1TableAdapter
        Dim DTA1 As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_itemdp_viewTableAdapter
        Dim DTA2 As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoiceitem_salesreturn_view0TableAdapter
        Dim DTA3 As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_add_dataTableAdapter
        Dim DTA4 As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_item_costTableAdapter

        Dim TBL As New DseTransaksiSalesInvoice.tbltrans_salesinvoice_item_view1DataTable
        Dim TBL1 As New DseTransaksiSalesInvoice.tbltrans_salesinvoice_itemdp_viewDataTable
        Dim TBL2 As New DseTransaksiSalesInvoice.tbltrans_salesinvoiceitem_salesreturn_view0DataTable
        Dim TBL3 As New DseTransaksiSalesInvoice.tbltrans_salesinvoice_add_dataDataTable
        Dim TBL4 As New DseTransaksiSalesInvoice.tbltrans_salesinvoice_item_costDataTable

        Dim row As DseTransaksiSalesInvoice.tbltrans_salesinvoice_item_view1Row
        Dim row1 As DseTransaksiSalesInvoice.tbltrans_salesinvoice_itemdp_viewRow
        Dim row2 As DseTransaksiSalesInvoice.tbltrans_salesinvoiceitem_salesreturn_view0Row
        Dim row3 As DseTransaksiSalesInvoice.tbltrans_salesinvoice_add_dataRow
        Dim row4 As DseTransaksiSalesInvoice.tbltrans_salesinvoice_item_costRow

        Dim ObjCD As New Core_Module.ClsCekData
        Dim rowDGV As DataGridViewRow

        Dim PrimarySO, PrimaryDO, PrimaryStockOut As Double
        Dim NoDO, NoSO, CustPONumber, NoStockOutSales As String
        Dim ObjSO As New ClsTransaksiSalesOrder
        Dim ObjDO As New ClsTransaksiDO
        Dim ObjInOut As New Inventory_Module.clsTransaksiStockInOut
        Dim DTASO As New DseTransaksiSOTableAdapters.tbltrans_soTableAdapter
        Dim DTADO As New DseTransaksiDOTableAdapters.tbltrans_do3TableAdapter
        Dim ObjP As New Core_Module.clsMasterProduct
        Dim DefPrecMulti As Integer = 0
        Dim DefPrecDefault As Integer = 0

        Try
            'kosongkan dgv
            vDGVItem.Rows.Clear()
            vDGVExchangeRate.Rows.Clear()
            vDGVDP.Rows.Clear()
            vDGVSR.Rows.Clear()

            'isikan tabel item
            PrimarySO = -1
            PrimaryDO = -1
            PrimaryStockOut = -1
            NoDO = ""
            NoSO = ""
            CustPONumber = ""
            NoStockOutSales = ""
            If DTA.Fill(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    With row
                        If .Primary_DO <> PrimaryDO Then
                            PrimaryDO = .Primary_DO
                            NoDO = ObjDO.GetDONO(PrimaryDO, DTADO)
                            NoStockOutSales = ObjInOut.GetNoTransaksiByPrimaryDOLossOrPrimaryDOSales(PrimaryDO, Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Sales)
                        End If
                        If .Primary_SO <> PrimarySO Then
                            PrimarySO = .Primary_SO
                            NoSO = ObjSO.GetNoSO(PrimarySO, DTASO)
                            CustPONumber = ObjSO.GetCustPONo(PrimarySO, DTASO)
                        End If
                        .SO_Number = NoSO
                        .Cust_PO_No = CustPONumber
                        .Do_Number = NoDO
                        .Stock_Out_No = NoStockOutSales

                        vDGVItem.Rows.Add()
                        rowDGV = vDGVItem.Rows(vDGVItem.RowCount - 1)
                        rowDGV.Cells(vFrm.ColNo0.Index).Value = "0"
                        rowDGV.Cells(vFrm.ColCodeNo.Index).Value = .Product_Code
                        rowDGV.Cells(vFrm.ColDescription.Index).Value = .Description
                        rowDGV.Cells(vFrm.ColCustCodeNo.Index).Value = .Customer_s_Code_No
                        rowDGV.Cells(vFrm.ColCustDesc.Index).Value = .Customer_s_Desc
                        rowDGV.Cells(vFrm.ColSONO.Index).Value = .SO_Number
                        rowDGV.Cells(vFrm.ColCustPONo.Index).Value = .Cust_PO_No
                        rowDGV.Cells(vFrm.ColDONO.Index).Value = .Do_Number
                        rowDGV.Cells(vFrm.ColStockOutNo.Index).Value = .Stock_Out_No
                        rowDGV.Cells(vFrm.ColQtyMaster.Index).Value = .Qty_Master
                        rowDGV.Cells(vFrm.ColQtyPerMaster.Index).Value = .Qty_Per_Master
                        rowDGV.Cells(vFrm.ColMasterUnits.Index).Value = .Master_Unit_Code
                        rowDGV.Cells(vFrm.ColQty.Index).Value = ObjCD.DigitGrouping(.Qty, .AngkaDecimalProduct)
                        rowDGV.Cells(vFrm.ColTotalQty.Index).Value = ObjCD.DigitGrouping(.Total_Qty, .AngkaDecimalProduct)
                        rowDGV.Cells(vFrm.ColUnit.Index).Value = .Unit_Code
                        rowDGV.Cells(vFrm.ColSecQty.Index).Value = ObjCD.DigitGrouping(.Secondary_Qty)
                        rowDGV.Cells(vFrm.ColSecUnit.Index).Value = .Secondary_Unit_Code
                        rowDGV.Cells(vFrm.ColCurrency.Index).Value = .Currency_Code
                        rowDGV.Cells(vFrm.ColUnitPrice.Index).Value = ObjCD.DigitGrouping(.Unit_Price, .Default_Prec_Currency)
                        rowDGV.Cells(vFrm.ColDiscount.Index).Value = ObjCD.DigitGrouping(.Discount, 2)
                        rowDGV.Cells(vFrm.ColNettPrice.Index).Value = ObjCD.DigitGrouping(.Nett_Price, .Default_Prec_Currency)
                        rowDGV.Cells(vFrm.ColAmount.Index).Value = ObjCD.DigitGrouping(.Amount, .Default_Prec_Currency)
                        rowDGV.Cells(vFrm.ColRemarksItem.Index).Value = ""
                        rowDGV.Cells(vFrm.ColPrimaryCodeNo.Index).Value = .Primary_Product
                        rowDGV.Cells(vFrm.ColPrimarySO.Index).Value = .Primary_SO
                        rowDGV.Cells(vFrm.ColPrimaryDO.Index).Value = .Primary_DO
                        rowDGV.Cells(vFrm.ColPrimaryMasterUnits.Index).Value = .Primary_Master_Unit
                        rowDGV.Cells(vFrm.ColPrimarySecUnit.Index).Value = .Primary_Secondary_Units
                        rowDGV.Cells(vFrm.ColPrimaryUnit.Index).Value = .Primary_Units
                        rowDGV.Cells(vFrm.ColPrimaryCurrency.Index).Value = .Primary_Currency
                        rowDGV.Cells(vFrm.ColLotNumber.Index).Value = .Batch_Number
                        rowDGV.Cells(vFrm.ColDiscFromPerct.Index).Value = .DiscFromPerc
                        rowDGV.Cells(vFrm.ColDiscountAmount.Index).Value = ObjCD.DigitGrouping(.Disc_Amount, .Default_Prec_Currency)
                        If .Expiry = #1/1/1950# Then
                            rowDGV.Cells(vFrm.ColExpiry.Index).Value = "-"
                        Else
                            rowDGV.Cells(vFrm.ColExpiry.Index).Value = Format(.Expiry, "dd MMM yyyy")
                        End If
                        DefPrecMulti = ObjP.GetAngkaDecimal(.Primary_Product, .Primary_Units_Multi)
                        DefPrecDefault = ObjP.GetAngkaDecimal(.Primary_Product)
                        rowDGV.Cells(vFrm.ColTotalQtyMulti.Index).Value = ObjCD.DigitGrouping(.Total_Qty_Multi, DefPrecMulti)
                        rowDGV.Cells(vFrm.ColUnitsCodeMulti.Index).Value = .Units_Code_Multi
                        rowDGV.Cells(vFrm.ColPrimaryUnitsMulti.Index).Value = .Primary_Units_Multi
                        rowDGV.Cells(vFrm.ColRatioReference.Index).Value = ObjCD.DigitGrouping(.Ratio_Reference, DefPrecMulti)
                        rowDGV.Cells(vFrm.ColRatioDefault.Index).Value = ObjCD.DigitGrouping(.Ratio_Default, DefPrecDefault)
                        If .Primary_Units <> .Primary_Units_Multi Then
                            rowDGV.Cells(vFrm.ColRatioDesc.Index).Value = ObjCD.DigitGrouping(.Ratio_Reference, DefPrecMulti) & " " & .Units_Code_Multi & "=" & ObjCD.DigitGrouping(.Ratio_Default, DefPrecDefault) & " " & .Unit_Code
                        Else
                            rowDGV.Cells(vFrm.ColRatioDesc.Index).Value = "-"
                        End If
                        rowDGV.Cells(vFrm.ColPrimarySalesPerson.Index).Value = .primary_sales_person
                        rowDGV.Cells(vFrm.ColSalesPerson.Index).Value = .sales_person_code & " - " & .sales_person_desc

                        rowDGV.Cells(vFrm.ColPriceByMultiUnit.Index).Value = .price_by_multi_unit
                        If .price_by_multi_unit = 1 Then
                            rowDGV.Cells(vFrm.ColPriceByMultiUnit_Str.Index).Value = "/" + rowDGV.Cells(vFrm.ColUnitsCodeMulti.Index).Value
                        Else
                            rowDGV.Cells(vFrm.ColPriceByMultiUnit_Str.Index).Value = "/" + rowDGV.Cells(vFrm.ColUnit.Index).Value
                        End If

                        '2025 Core Tax Tahap 2.1
                        If .Tanggal_Transaksi >= #1/1/2025# Then
                            rowDGV.Cells(vFrm.ColPrimarySalesPerson.Index).Value = .primary_sales_person
                            rowDGV.Cells(vFrm.ColSalesPerson.Index).Value = .sales_person_code & " - " & .sales_person_desc

                            rowDGV.Cells(vFrm.ColPriceByMultiUnit.Index).Value = .price_by_multi_unit
                            If .price_by_multi_unit = 1 Then
                                rowDGV.Cells(vFrm.ColPriceByMultiUnit_Str.Index).Value = "/" + rowDGV.Cells(vFrm.ColUnitsCodeMulti.Index).Value
                            Else
                                rowDGV.Cells(vFrm.ColPriceByMultiUnit_Str.Index).Value = "/" + rowDGV.Cells(vFrm.ColUnit.Index).Value
                            End If
                            '\\TR==GlobalDiscount
                            rowDGV.Cells(vFrm.ColUnitPriceInclude.Index).Value = ObjCD.DigitGrouping(.Unit_Price_Include, .Default_Prec_Currency)
                            rowDGV.Cells(vFrm.ColDiscount1PercInclude.Index).Value = .Disc_From_Perct_1
                            rowDGV.Cells(vFrm.ColDiscount1Include.Index).Value = ObjCD.DigitGrouping(.Discount_1, 2)
                            rowDGV.Cells(vFrm.ColDiscount1AmountInclude.Index).Value = ObjCD.DigitGrouping(.Disc_Amount_1, .Default_Prec_Currency)
                            rowDGV.Cells(vFrm.ColDiscountGlobalPercInclude.Index).Value = ObjCD.DigitGrouping(.Global_Discount_Percent, 2)
                            rowDGV.Cells(vFrm.ColDiscountGlobalAmountInclude.Index).Value = ObjCD.DigitGrouping(.Global_Discount_Amount, .Default_Prec_Currency)
                            rowDGV.Cells(vFrm.ColDiscountGlobalStringInclude.Index).Value = .Global_Discount_String
                            rowDGV.Cells(vFrm.ColDiscountInclude.Index).Value = ObjCD.DigitGrouping(.Discount_Amount_Include_Perc, 2)
                            rowDGV.Cells(vFrm.ColDiscAmountInclude.Index).Value = ObjCD.DigitGrouping(.Discount_Amount_Include, .Default_Prec_Currency)
                            rowDGV.Cells(vFrm.ColNettPriceInclude.Index).Value = ObjCD.DigitGrouping(.Nett_Price_Include, .Default_Prec_Currency)
                            rowDGV.Cells(vFrm.ColAmountInclude.Index).Value = ObjCD.DigitGrouping(.Amount_Include, .Default_Prec_Currency)

                            rowDGV.Cells(vFrm.ColVAT.Index).Value = ObjCD.DigitGrouping(.VAT, .Default_Prec_Currency)
                            rowDGV.Cells(vFrm.ColVAT2.Index).Value = ObjCD.DigitGrouping(.VAT2, .Default_Prec_Currency)
                            rowDGV.Cells(vFrm.ColAmount2.Index).Value = ObjCD.DigitGrouping(.Amount2, .Default_Prec_Currency)

                            If .Used_VAT = True Then
                                rowDGV.Cells(vFrm.ColIsVAT.Index).Value = Boolean.TrueString
                            Else
                                rowDGV.Cells(vFrm.ColIsVAT.Index).Value = Boolean.FalseString
                            End If

                            If .Price_Exclude = .Price_Exclude_SO Then
                                If .Price_Exclude = False Then
                                    rowDGV.Cells(vFrm.ColPriceCond.Index).Value = "Include"
                                Else
                                    rowDGV.Cells(vFrm.ColPriceCond.Index).Value = "Exclude"
                                End If
                            Else
                                If .Price_Exclude_SO = False Then
                                    rowDGV.Cells(vFrm.ColPriceCond.Index).Value = "Include"
                                Else
                                    rowDGV.Cells(vFrm.ColPriceCond.Index).Value = "Exclude"
                                End If
                            End If

                            rowDGV.Cells(vFrm.ColPrimarySORef.Index).Value = -1
                            'aul
                            rowDGV.Cells(vFrm.colPrimaryCostAcc.Index).Value = -1
                            rowDGV.Cells(vFrm.colCostAccCode.Index).Value = ""
                            rowDGV.Cells(vFrm.colCostAccDesc.Index).Value = ""
                        End If
                    End With
                Next
            End If

            'isikan tabel item dp
            If DTA1.FillBy_Primary(TBL1, vPrimaryTransaksi) > 0 Then
                For Each row1 In TBL1.Rows
                    With row1
                        If .Tipe_DP = 1 Then
                            If Me.GetManualDP_Data(.Primary_Invoice_DP, row1) = False Then
                                'jika false artinya balance manual tidak ditemukan
                                GoTo SkipInsertToDgv
                            End If
                        End If
                        vDGVDP.Rows.Add("0", .No_Invoice_DP, .Currency_Code, ObjCD.DigitGrouping(.Sub_Total, .Default_Precision), .Remarks, _
                        .Primary_Invoice_DP, .Tipe_DP)
                    End With
SkipInsertToDgv:
                Next
            End If

            'isikan tabel item exchange rate
            If vCalculatedBy = ClsTransaksiSalesInvoice.EnumCalculatesBy.Invoice Then
                Dim DTAEX As New DseTransaksiSalesInvoiceTableAdapters.DataTable2TableAdapter
                Dim TBLEX As New DseTransaksiSalesInvoice.DataTable2DataTable
                Dim RowEX As DseTransaksiSalesInvoice.DataTable2Row
                If DTAEX.Fill(TBLEX, vPrimaryTransaksi) > 0 Then
                    For Each RowEX In TBLEX.Rows
                        With RowEX
                            vDGVExchangeRate.Rows.Add("0", .Sales_Invoice_No, .Currency_Code, _
                            .Exhange_Rate, .Remarks, .Primary_Transaksi_Rujukan, .Primary_Currency)
                        End With
                    Next
                End If
                DTAEX.Dispose()
                TBLEX.Dispose()
            ElseIf vCalculatedBy = ClsTransaksiSalesInvoice.EnumCalculatesBy.Delivery_Order Then
                Dim DTAEX As New DseTransaksiSalesInvoiceTableAdapters.DataTable3TableAdapter
                Dim TBLEX As New DseTransaksiSalesInvoice.DataTable3DataTable
                Dim RowEX As DseTransaksiSalesInvoice.DataTable3Row
                If DTAEX.Fill(TBLEX, vPrimaryTransaksi) > 0 Then
                    For Each RowEX In TBLEX.Rows
                        With RowEX
                            vDGVExchangeRate.Rows.Add("0", .Delivery_Order_No, .Currency_Code, _
                            .Exhange_Rate, .Remarks, .Primary_Transaksi_Rujukan, .Primary_Currency)
                        End With
                    Next
                End If
                DTAEX.Dispose()
                TBLEX.Dispose()
            ElseIf vCalculatedBy = ClsTransaksiSalesInvoice.EnumCalculatesBy.Sales_Order Then
                Dim DTAEX As New DseTransaksiSalesInvoiceTableAdapters.DataTable4TableAdapter
                Dim TBLEX As New DseTransaksiSalesInvoice.DataTable4DataTable
                Dim RowEX As DseTransaksiSalesInvoice.DataTable4Row
                If DTAEX.Fill(TBLEX, vPrimaryTransaksi) > 0 Then
                    For Each RowEX In TBLEX.Rows
                        With RowEX
                            vDGVExchangeRate.Rows.Add("0", .Sales_Order_No, .Currency_Code, _
                            .Exhange_Rate, .Remarks, .Primary_Transaksi_Rujukan, .Primary_Currency)
                        End With
                    Next
                End If
                DTAEX.Dispose()
                TBLEX.Dispose()
            End If

            'isikan dgv SR
            DTA2.FillBy_PrimarySalesInvoice(TBL2, vPrimaryTransaksi)
            For Each row2 In TBL2.Rows
                vDGVSR.Rows.Add("", row2.Sales_Return_No, row2.Delivery_Order_No, row2.Remarks, row2.Primary_SR)
            Next


            vFrm.TBLHeaderCost.Rows.Clear()
            If DTA3.FillBy(TBL3, vPrimaryTransaksi) > 0 Then
                For Each row3 In TBL3.Rows
                    With row3
                        vFrm.rowHeaderCost = vFrm.TBLHeaderCost.NewRow
                        vFrm.rowHeaderCost.primary = .primary
                        vFrm.rowHeaderCost.primary_cust_ship_to = .primary_cust_ship_to
                        vFrm.rowHeaderCost.primary_inv = .primary_inv
                        vFrm.rowHeaderCost.primary_volume_unit = .primary_volume_unit
                        vFrm.rowHeaderCost.primary_weight_unit = .primary_weight_unit
                        vFrm.rowHeaderCost.ETA_Date = .ETA_Date
                        vFrm.rowHeaderCost.ETA_Port = .ETA_Port
                        vFrm.rowHeaderCost.ETD_Date = .ETD_Date
                        vFrm.rowHeaderCost.ETD_Port = .ETD_Port
                        vFrm.rowHeaderCost.identifier_cust_ship_to = .identifier_cust_ship_to
                        vFrm.rowHeaderCost.invoice_type = .invoice_type
                        vFrm.rowHeaderCost.shipping_mark = .shipping_mark
                        vFrm.rowHeaderCost.Total_Gross_Volume = .Total_Gross_Volume
                        vFrm.rowHeaderCost.Total_Gross_Weight = .Total_Gross_Weight
                        vFrm.rowHeaderCost.Total_Nett_Volume = .Total_Nett_Volume
                        vFrm.rowHeaderCost.Total_Nett_Weight = .Total_Nett_Weight
                        vFrm.rowHeaderCost.transportation_mode = .transportation_mode
                        vFrm.rowHeaderCost.var1_name = .var1_name
                        vFrm.rowHeaderCost.var1_value = .var1_value
                        vFrm.rowHeaderCost.var2_name = .var2_name
                        vFrm.rowHeaderCost.var2_value = .var2_value
                        vFrm.rowHeaderCost.var3_name = .var3_name
                        vFrm.rowHeaderCost.var3_value = .var3_value
                        vFrm.rowHeaderCost.var4_name = .var4_name
                        vFrm.rowHeaderCost.var4_value = .var4_value
                        vFrm.rowHeaderCost.var5_name = .var5_name
                        vFrm.rowHeaderCost.var5_value = .var5_value
                        vFrm.TBLHeaderCost.Rows.Add(vFrm.rowHeaderCost)
                    End With
                Next
            End If

            vFrm.TBLitemCost.Rows.Clear()
            If DTA4.Fill(TBL4, vPrimaryTransaksi) > 0 Then
                For Each row4 In TBL4.Rows
                    With row4
                        vFrm.rowitemCost = vFrm.TBLitemCost.NewRow
                        vFrm.rowitemCost.Primary = .Primary
                        vFrm.rowitemCost.Primary_Akun = .Primary_Akun
                        vFrm.rowitemCost.Primary_Inv = .Primary_Inv
                        vFrm.rowitemCost.Remarks = .Remarks
                        vFrm.rowitemCost.Cost = .Cost
                        vFrm.rowitemCost.Cost_Description = .Cost_Description
                        vFrm.TBLitemCost.Rows.Add(vFrm.rowitemCost)
                    End With
                Next
            End If


            'set nomer
            ObjCD.SetNumber(vDGVItem)
            ObjCD.SetNumber(vDGVDP)
            ObjCD.SetNumber(vDGVExchangeRate)
            ObjCD.SetNumber(vDGVSR)

            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
            ObjCD = Nothing
            ObjSO = Nothing
            ObjDO = Nothing
            ObjInOut = Nothing
            DTASO.Dispose()
            DTADO.Dispose()
        End Try
    End Function

    Public Function Simpan(ByVal vdgvItem As DataGridView, ByVal vDgvExchangeRate As Windows.Forms.DataGridView, _
   ByVal vDGVDP As DataGridView, ByVal vPrimaryTransaksi As Double, ByVal vDGVSR As DataGridView, _
   ByVal vFrm As FrmTransaksiSalesInvoiceQuick, ByVal vTblCustom As CustomSMESS.DseTransaksiSalesInvoiceQuick_AddByCustomProduct.tbltrans_custom_salesinvoice_item1DataTable) As Boolean
        Dim row As DataGridViewRow
        'Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_itemTableAdapter
        Dim DTA1 As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_item_exchangerateTableAdapter
        'Dim DTA2 As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_itemdpTableAdapter
        'Dim DTA3 As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_item_salesreturnTableAdapter
        Dim SecQty As Double = -1
        Dim DTA_Inv_DO As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_invoice_doTableAdapter
        Dim TBL_Inv_DO As New DseTransaksiSalesInvoice.tbltrans_invoice_doDataTable
        Dim rowInvDO As DseTransaksiSalesInvoice.tbltrans_invoice_doRow
        Dim ED As Date = #1/1/1950#
        Dim ObjCustom As New CustomSMESS.ClsTransaksiSalesInvoiceQuick_AddByCustomProduct

        Try
            'simpan data item
            Me.Simpan_UsingScript_QuickInvoice(vdgvItem, vPrimaryTransaksi, vFrm)
            'For Each row In vdgvItem.Rows
            '    'simpan ke basis data item
            '    With vFrm
            '        If IsNumeric(row.Cells(.ColSecQty.Index).Value) = True Then
            '            SecQty = row.Cells(.ColSecQty.Index).Value
            '        Else
            '            SecQty = -1
            '        End If
            '        If IsDate(row.Cells(.ColExpiry.Index).Value) = True Then
            '            ED = row.Cells(.ColExpiry.Index).Value
            '        Else
            '            ED = #1/1/1950#
            '        End If

            '        If IsNothing(row.Cells(.ColAliasCode.Index).Value) = True Then
            '            row.Cells(.ColAliasCode.Index).Value = ""
            '        End If
            '        If IsNothing(row.Cells(.ColAliasDesc.Index).Value) = True Then
            '            row.Cells(.ColAliasDesc.Index).Value = ""
            '        End If
            '        If String.IsNullOrEmpty(.TxtPrimarySalesPerson.Text) = True Then
            '            .TxtPrimarySalesPerson.Text = -1
            '        End If
            '        DTA.InsertQuery_QuickSales(vPrimaryTransaksi, row.Cells(.ColPrimaryCodeNo.Index).Value, row.Cells(.ColPrimaryDO.Index).Value, row.Cells(.ColPrimarySO.Index).Value, _
            '            row.Cells(.ColQtyMaster.Index).Value, row.Cells(.ColPrimaryMasterUnits.Index).Value, row.Cells(.ColQty.Index).Value, row.Cells(.ColTotalQty.Index).Value, _
            '            row.Cells(.ColPrimaryUnit.Index).Value, SecQty, row.Cells(.ColPrimarySecUnit.Index).Value, row.Cells(.ColUnitPrice.Index).Value, _
            '            row.Cells(.ColDiscount.Index).Value, row.Cells(.ColNettPrice.Index).Value, row.Cells(.ColAmount.Index).Value, row.Cells(.ColRemarksItem.Index).Value, _
            '            row.Cells(.ColQtyPerMaster.Index).Value, row.Cells(.ColPrimaryCurrency.Index).Value, row.Cells(.ColLotNumber.Index).Value, ED, _
            '            row.Cells(.ColPrimaryLocation.Index).Value, row.Cells(.ColAliasCode.Index).Value, row.Cells(.ColAliasDesc.Index).Value, _
            '            row.Cells(.ColDiscAmount.Index).Value, row.Cells(.ColDiscFromPrect.Index).Value, _
            '            row.Cells(.ColTotalQtyMulti.Index).Value, row.Cells(.ColPrimaryUnitsMulti.Index).Value, _
            '            row.Cells(.ColRatioReference.Index).Value, row.Cells(.ColRatioDefault.Index).Value, row.Cells(.ColVAT.Index).Value, _
            '            row.Cells(.ColAmount2.Index).Value, row.Cells(.ColVAT2.Index).Value, .TxtPrimarySalesPerson.Text)
            '    End With
            'Next

            'simpan data item exchange rate
            For Each row In vDgvExchangeRate.Rows
                DTA1.Insert(vPrimaryTransaksi, row.Cells(5).Value, row.Cells(6).Value, _
                row.Cells(3).Value, row.Cells(4).Value)
            Next

            'simpan item DP -> quick invoice tidak memiliki DP
            'For Each row In vDGVDP.Rows
            '    DTA2.Insert(vPrimaryTransaksi, row.Cells(5).Value, row.Cells(4).Value, 0)
            'Next

            'simpan item sales return -> quick invoice tidak memiliki sales return
            'For Each row In vDGVSR.Rows
            '    DTA3.Insert(vPrimaryTransaksi, row.Cells(4).Value, row.Cells(3).Value)
            'Next

            'tandai do
            'update data do jadi belum buat inv
            DTA_Inv_DO.FillBy_PrimaryInvoice(TBL_Inv_DO, vPrimaryTransaksi)
            For Each rowInvDO In TBL_Inv_DO.Rows
                DTA_Inv_DO.UpdateQuery_SetSudahBuatInvoice(1, rowInvDO.Primary_Transaksi, rowInvDO.Primary_DO)
            Next

            'ObjCustom.SimpanFromTblDummy(vTblCustom)
            ObjCustom.SimpanFromTblDummy(vTblCustom, vPrimaryTransaksi, CDbl(vFrm.TxtPrimarySalesPerson.Text))

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            'DTA.Dispose()
            DTA1.Dispose()
            'DTA2.Dispose()
            DTA_Inv_DO.Dispose()
            TBL_Inv_DO.Dispose()
        End Try
    End Function

    Public Sub IsikanTBLLocation(ByVal TBL As DseTransaksiSalesInvoice.tbltrans_salesinvoice_item3DataTable, _
  ByVal vPrimaryProduct As Double, ByVal vPrimaryTransaksi As Double, ByVal vBatchNo As String)
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_item3TableAdapter

        Try
            DTA.Fill(TBL, vPrimaryTransaksi, vPrimaryProduct, vBatchNo)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsikanTBLBatchNo(ByVal TBL As DseTransaksiSalesInvoice.tbltrans_salesinvoice_item2DataTable, _
   ByVal vPrimaryProduct As Double, ByVal vPrimaryTransaksi As Double)
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_item2TableAdapter

        Try
            DTA.Fill(TBL, vPrimaryTransaksi, vPrimaryProduct)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function Simpan(ByVal vdgvItem As DataGridView, ByVal vDgvExchangeRate As Windows.Forms.DataGridView, _
    ByVal vDGVDP As DataGridView, ByVal vPrimaryTransaksi As Double, ByVal vDGVSR As DataGridView, _
    ByVal vFrm As FrmTransaksiSalesInvoice) As Boolean
        Dim row As DataGridViewRow
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_itemTableAdapter
        Dim DTA1 As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_item_exchangerateTableAdapter
        Dim DTA2 As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_itemdpTableAdapter
        Dim DTA3 As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_item_salesreturnTableAdapter
        Dim SecQty As Double = -1
        Dim DTA_Inv_DO As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_invoice_doTableAdapter
        Dim TBL_Inv_DO As New DseTransaksiSalesInvoice.tbltrans_invoice_doDataTable
        Dim rowInvDO As DseTransaksiSalesInvoice.tbltrans_invoice_doRow
        Dim ED As Date = #1/1/1950#
        Dim ListPrimaryDO As New List(Of Double)

        Try
            DTA_Inv_DO.Connection.ConnectionString = Core_Module.GetConstring

            'yulimus-optimasi-cut off
            'simpan data item
            Me.Simpan_UsingScript(vdgvItem, vPrimaryTransaksi, vFrm)
            'simpan data item
            For Each row In vdgvItem.Rows
                'simpan ke basis data item
                With vFrm
                    'If IsNumeric(row.Cells(.ColSecQty.Index).Value) = True Then
                    '    SecQty = row.Cells(.ColSecQty.Index).Value
                    'Else
                    '    SecQty = -1
                    'End If
                    'If IsDate(row.Cells(.ColExpiry.Index).Value) = True Then
                    '    ED = row.Cells(.ColExpiry.Index).Value
                    'Else
                    '    ED = #1/1/1950#
                    'End If

                    'DTA.InsertQuery(vPrimaryTransaksi, row.Cells(.ColPrimaryCodeNo.Index).Value, row.Cells(.ColPrimaryDO.Index).Value, row.Cells(.ColPrimarySO.Index).Value, _
                    '    row.Cells(.ColQtyMaster.Index).Value, row.Cells(.ColPrimaryMasterUnits.Index).Value, row.Cells(.ColQty.Index).Value, row.Cells(.ColTotalQty.Index).Value, _
                    '    row.Cells(.ColPrimaryUnit.Index).Value, SecQty, row.Cells(.ColPrimarySecUnit.Index).Value, row.Cells(.ColUnitPrice.Index).Value, _
                    '    row.Cells(.ColDiscount.Index).Value, row.Cells(.ColNettPrice.Index).Value, row.Cells(.ColAmount.Index).Value, row.Cells(.ColRemarksItem.Index).Value, _
                    '    row.Cells(.ColQtyPerMaster.Index).Value, row.Cells(.ColPrimaryCurrency.Index).Value, row.Cells(.ColLotNumber.Index).Value, ED, _
                    '    row.Cells(.ColDiscountAmount.Index).Value, row.Cells(.ColDiscFromPerct.Index).Value, row.Cells(.ColTotalQtyMulti.Index).Value, row.Cells(.ColPrimaryUnitsMulti.Index).Value, _
                    '    row.Cells(.ColRatioReference.Index).Value, row.Cells(.ColRatioDefault.Index).Value)

                    If ListPrimaryDO.Contains(row.Cells(.ColPrimaryDO.Index).Value) = False Then
                        ListPrimaryDO.Add(row.Cells(.ColPrimaryDO.Index).Value)
                    End If
                End With
            Next

            'simpan data item exchange rate
            For Each row In vDgvExchangeRate.Rows
                DTA1.Insert(vPrimaryTransaksi, row.Cells(5).Value, row.Cells(6).Value, _
                row.Cells(3).Value, row.Cells(4).Value)
            Next

            'simpan item DP
            For Each row In vDGVDP.Rows
                DTA2.Insert(vPrimaryTransaksi, row.Cells(5).Value, row.Cells(4).Value, row.Cells(6).Value)
            Next

            'simpan item sales return
            For Each row In vDGVSR.Rows
                DTA3.Insert(vPrimaryTransaksi, row.Cells(4).Value, row.Cells(3).Value)
            Next

            'tandai do
            'update data do jadi belum buat inv
            'yulimus-optimasi-cut off
            For Each PrimaryDO As Double In ListPrimaryDO
                DTA_Inv_DO.UpdateQuery_SetSudahBuatInvoice(1, vPrimaryTransaksi, PrimaryDO)
            Next
            'DTA_Inv_DO.FillBy_PrimaryInvoice(TBL_Inv_DO, vPrimaryTransaksi)
            'For Each rowInvDO In TBL_Inv_DO.Rows
            '    DTA_Inv_DO.UpdateQuery_SetSudahBuatInvoice(1, rowInvDO.Primary_Transaksi, rowInvDO.Primary_DO)
            'Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
            DTA2.Dispose()
            DTA_Inv_DO.Dispose()
            TBL_Inv_DO.Dispose()
        End Try
    End Function

    'yulimus-optimasi-cut off
    Private Function Simpan_UsingScript(ByVal vdgvItem As DataGridView, ByVal vPrimaryTransaksi As Double, _
    ByVal vFrm As FrmTransaksiSalesInvoice) As Boolean
        Dim conStr As String = My.Settings.core_dbConnectionString
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < vdgvItem.Rows.Count
                sql = Me.GetInsertString(StartRowIndex, BatchQty, vdgvItem, vPrimaryTransaksi, vFrm)
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

    Private Function GetInsertString(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, ByVal vdgvItem As DataGridView, _
   ByVal vPrimaryTransaksi As Double, ByVal ijk As FrmTransaksiSalesInvoice) As String
        Dim retval As String = ""
        Dim row As DataGridViewRow
        Dim Delimiter As String = ","
        Dim ED As Date = #1/1/1950#

        Try
            '2025 Core Tax Tahap 2.1
            If ijk.dtpDate.Value >= #1/1/2025# Then
                retval = Me.GetInsertString_Update(vStartRowIndex, BatchQty, vdgvItem, vPrimaryTransaksi, ijk)
                Return retval
            End If

            retval = "INSERT INTO `tbltrans_salesinvoice_item` (`Primary Transaksi`, `Primary Product`, `Primary DO`, `Primary SO`, " + _
            "`Qty Master`, `Qty Per Master`, `Primary Master Unit`, `Qty`, `Total Qty`, `Primary Units`, `Secondary Qty`," + _
            " `Primary Secondary Units`, `Primary Currency`, `Unit Price`, `Discount`, `Nett Price`, `Amount`, `Remarks`," + _
            " `Batch Number`, `Expiry`, `Cost`, `Sales Amount`, `Primary Location`, `Refer To SO`, `Alias Code`, `Alias Desc`," + _
            " `Disc Amount`, `DiscFromPerc`, `Total Qty Multi`, `Primary Units Multi`, `Ratio Reference`, `Ratio Default`, `primary sales person`, `price by multi unit`) VALUES "
            'retval = "INSERT INTO `tbltrans_purchase_closing_item` (`Primary Transaksi`, `Primary Product`, `Primary Unit`, `Begin Qty`, `Begin Amount`, `In Qty`, `In Amount`, `Out Qty`, `Out Amount`, `Ending Qty`, `Ending Amount`, `Remarks`) VALUES "
            For i As Double = 0 To BatchQty - 1 Step 1
                If vdgvItem.Rows.Count > vStartRowIndex + i Then
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vdgvItem.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    row = vdgvItem.Rows(vStartRowIndex + i)

                    If IsDate(row.Cells(ijk.ColExpiry.Index).Value) = True Then
                        ED = row.Cells(ijk.ColExpiry.Index).Value
                    Else
                        ED = #1/1/1950#
                    End If

                    retval = retval + vbCrLf + "(" + vPrimaryTransaksi.ToString + "," + CDbl(row.Cells(ijk.ColPrimaryCodeNo.Index).Value).ToString + "," _
                    + CDbl(row.Cells(ijk.ColPrimaryDO.Index).Value).ToString + "," _
                    + CDbl(row.Cells(ijk.ColPrimarySO.Index).Value).ToString + ",0,0,-1," _
                    + CDec(row.Cells(ijk.ColQty.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.ColTotalQty.Index).Value).ToString + "," _
                    + CDbl(row.Cells(ijk.ColPrimaryUnit.Index).Value).ToString + ",-1,-1," _
                    + CDbl(row.Cells(ijk.ColPrimaryCurrency.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.ColUnitPrice.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.ColDiscount.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.ColNettPrice.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.ColAmount.Index).Value).ToString + "," _
                    + "'" + row.Cells(ijk.ColRemarksItem.Index).Value.ToString + "'," _
                    + "'" + row.Cells(ijk.ColLotNumber.Index).Value.ToString + "'," _
                    + "'" + Format(ED, "yyyy-MM-dd") + "',0,0,-1,0,'',''," _
                    + CDec(row.Cells(ijk.ColDiscountAmount.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.ColDiscFromPerct.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.ColTotalQtyMulti.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.ColPrimaryUnitsMulti.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.ColRatioReference.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.ColRatioDefault.Index).Value).ToString + "," _
                    + CDbl(row.Cells(ijk.ColPrimarySalesPerson.Index).Value).ToString + "," _
                    + CInt(row.Cells(ijk.ColPriceByMultiUnit.Index).Value).ToString + "" _
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

    '2025 Core Tax Tahap 2.1, tambahan untuk simpan ke database invoice
    Private Function GetInsertString_Update(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, ByVal vdgvItem As DataGridView, _
  ByVal vPrimaryTransaksi As Double, ByVal ijk As FrmTransaksiSalesInvoice) As String
        Dim retval As String = ""
        Dim row As DataGridViewRow
        Dim Delimiter As String = ","
        Dim ED As Date = #1/1/1950#
        Dim priceExclude As Integer = 1
        Dim UseVAT As Integer = 1
        Try
            retval = "INSERT INTO `tbltrans_salesinvoice_item` (`Primary Transaksi`, `Primary Product`, `Primary DO`, `Primary SO`, " + _
            "`Qty Master`, `Qty Per Master`, `Primary Master Unit`, `Qty`, `Total Qty`, `Primary Units`, `Secondary Qty`," + _
            " `Primary Secondary Units`, `Primary Currency`, `Unit Price`, `Discount`, `Nett Price`, `Amount`, `Remarks`," + _
            " `Batch Number`, `Expiry`, `Cost`, `Sales Amount`, `Primary Location`, `Refer To SO`, `Alias Code`, `Alias Desc`," + _
            " `Disc Amount`, `DiscFromPerc`, `Total Qty Multi`, `Primary Units Multi`, `Ratio Reference`, `Ratio Default`,`VAT`,`VAT2`,`Amount2`," + _
            " `primary sales person`, `price by multi unit`,`Disc From Perct 1`,`Discount 1`,`Disc Amount 1`,`Global Discount Percent`,`Global Discount Amount`," + _
            " `Global Discount String`, `Unit Price Include`, `Discount Amount Include Perc`, `Discount Amount Include`," + _
            " `Nett Price Include`, `Amount Include`, `Primary Cost Acc`, `Price Exclude`,`Used VAT`) VALUES "
            'retval = "INSERT INTO `tbltrans_purchase_closing_item` (`Primary Transaksi`, `Primary Product`, `Primary Unit`, `Begin Qty`, `Begin Amount`, `In Qty`, `In Amount`, `Out Qty`, `Out Amount`, `Ending Qty`, `Ending Amount`, `Remarks`) VALUES "
            For i As Double = 0 To BatchQty - 1 Step 1
                If vdgvItem.Rows.Count > vStartRowIndex + i Then
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vdgvItem.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    row = vdgvItem.Rows(vStartRowIndex + i)

                    If IsDate(row.Cells(ijk.ColExpiry.Index).Value) = True Then
                        ED = row.Cells(ijk.ColExpiry.Index).Value
                    Else
                        ED = #1/1/1950#
                    End If
                    If row.Cells(ijk.ColPriceCond.Index).Value = "Include" Then
                        priceExclude = 0
                    Else
                        priceExclude = 1
                    End If
                    If row.Cells(ijk.ColIsVAT.Index).Value = Boolean.FalseString Then
                        UseVAT = 0
                    Else
                        UseVAT = 1
                    End If

                    retval = retval + vbCrLf + "(" + vPrimaryTransaksi.ToString + "," + CDbl(row.Cells(ijk.ColPrimaryCodeNo.Index).Value).ToString + "," _
                    + CDbl(row.Cells(ijk.ColPrimaryDO.Index).Value).ToString + "," _
                    + CDbl(row.Cells(ijk.ColPrimarySO.Index).Value).ToString + ",0,0,-1," _
                    + CDec(row.Cells(ijk.ColQty.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.ColTotalQty.Index).Value).ToString + "," _
                    + CDbl(row.Cells(ijk.ColPrimaryUnit.Index).Value).ToString + ",-1,-1," _
                    + CDbl(row.Cells(ijk.ColPrimaryCurrency.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.ColUnitPrice.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.ColDiscount.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.ColNettPrice.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.ColAmount.Index).Value).ToString + "," _
                    + "'" + row.Cells(ijk.ColRemarksItem.Index).Value.ToString + "'," _
                    + "'" + row.Cells(ijk.ColLotNumber.Index).Value.ToString + "'," _
                    + "'" + Format(ED, "yyyy-MM-dd") + "',0,0,-1,0,'',''," _
                    + CDec(row.Cells(ijk.ColDiscountAmount.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.ColDiscFromPerct.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.ColTotalQtyMulti.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.ColPrimaryUnitsMulti.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.ColRatioReference.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.ColRatioDefault.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.ColVAT.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.ColVAT2.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.ColAmount2.Index).Value).ToString + "," _
                    + CDbl(row.Cells(ijk.ColPrimarySalesPerson.Index).Value).ToString + "," _
                    + CInt(row.Cells(ijk.ColPriceByMultiUnit.Index).Value).ToString + "," _
                    + CInt(row.Cells(ijk.ColDiscount1PercInclude.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.ColDiscount1Include.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.ColDiscount1AmountInclude.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.ColDiscountGlobalPercInclude.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.ColDiscountGlobalAmountInclude.Index).Value).ToString + "," _
                    + "'" + row.Cells(ijk.ColDiscountGlobalStringInclude.Index).Value.ToString + "'," _
                    + CDec(row.Cells(ijk.ColUnitPriceInclude.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.ColDiscountInclude.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.ColDiscAmountInclude.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.ColNettPriceInclude.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.ColAmountInclude.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.colPrimaryCostAcc.Index).Value).ToString + "," _
                    + priceExclude.ToString + "," _
                    + UseVAT.ToString + "" _
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

    Private Function Simpan_UsingScript_QuickInvoice(ByVal vdgvItem As DataGridView, ByVal vPrimaryTransaksi As Double, _
  ByVal vFrm As FrmTransaksiSalesInvoiceQuick) As Boolean
        Dim conStr As String = My.Settings.core_dbConnectionString
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < vdgvItem.Rows.Count
                sql = Me.GetInsertString_QuickInvoice(StartRowIndex, BatchQty, vdgvItem, vPrimaryTransaksi, vFrm)
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

    'yulimus-optimasi-cut off
    Private Function GetInsertString_QuickInvoice(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, ByVal vdgvItem As DataGridView, _
    ByVal vPrimaryTransaksi As Double, ByVal ijk As FrmTransaksiSalesInvoiceQuick) As String
        Dim retval As String = ""
        Dim row As DataGridViewRow
        Dim Delimiter As String = ","
        Dim ED As Date = #1/1/1950#

        Try
            'DTA.InsertQuery_QuickSales(
            '            , _
            '            ,)           

            retval = "INSERT INTO `tbltrans_salesinvoice_item` (`Primary Transaksi`, `Primary Product`, `Primary DO`, `Primary SO`, " + _
            "`Qty Master`, `Qty Per Master`, `Primary Master Unit`, `Qty`, `Total Qty`, `Primary Units`, `Secondary Qty`," + _
            " `Primary Secondary Units`, `Primary Currency`, `Unit Price`, `Discount`, `Nett Price`, `Amount`, `Remarks`," + _
            " `Batch Number`, `Expiry`, `Cost`, `Sales Amount`, `Primary Location`, `Refer To SO`, `Alias Code`, `Alias Desc`," + _
            " `Disc Amount`, `DiscFromPerc`, `Total Qty Multi`, `Primary Units Multi`, `Ratio Reference`, `Ratio Default`, " + _
            "`primary sales person`, `price by multi unit`, `VAT`, `Amount2`, `VAT2`) VALUES "
            'retval = "INSERT INTO `tbltrans_purchase_closing_item` (`Primary Transaksi`, `Primary Product`, `Primary Unit`, `Begin Qty`, `Begin Amount`, `In Qty`, `In Amount`, `Out Qty`, `Out Amount`, `Ending Qty`, `Ending Amount`, `Remarks`) VALUES "
            For i As Double = 0 To BatchQty - 1 Step 1
                If vdgvItem.Rows.Count > vStartRowIndex + i Then
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vdgvItem.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    row = vdgvItem.Rows(vStartRowIndex + i)

                    If IsDate(row.Cells(ijk.ColExpiry.Index).Value) = True Then
                        ED = row.Cells(ijk.ColExpiry.Index).Value
                    Else
                        ED = #1/1/1950#
                    End If

                    retval = retval + vbCrLf + "(" + vPrimaryTransaksi.ToString + "," + CDbl(row.Cells(ijk.ColPrimaryCodeNo.Index).Value).ToString + "," _
                    + CDbl(row.Cells(ijk.ColPrimaryDO.Index).Value).ToString + "," _
                    + CDbl(row.Cells(ijk.ColPrimarySO.Index).Value).ToString + ",0,0,-1," _
                    + CDec(row.Cells(ijk.ColQty.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.ColTotalQty.Index).Value).ToString + "," _
                    + CDbl(row.Cells(ijk.ColPrimaryUnit.Index).Value).ToString + ",-1,-1," _
                    + CDbl(row.Cells(ijk.ColPrimaryCurrency.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.ColUnitPrice.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.ColDiscount.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.ColNettPrice.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.ColAmount.Index).Value).ToString + "," _
                    + "'" + row.Cells(ijk.ColRemarksItem.Index).Value.ToString + "'," _
                    + "'" + row.Cells(ijk.ColLotNumber.Index).Value.ToString + "'," _
                    + "'" + Format(ED, "yyyy-MM-dd") + "',0,0," _
                    + CDbl(row.Cells(ijk.ColPrimaryLocation.Index).Value).ToString + ",0,'',''," _
                    + CDec(row.Cells(ijk.ColDiscAmount.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.ColDiscFromPrect.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.ColTotalQtyMulti.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.ColPrimaryUnitsMulti.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.ColRatioReference.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.ColRatioDefault.Index).Value).ToString + "," _
                    + CDbl(ijk.TxtPrimarySalesPerson.Text).ToString + "," _
                    + CInt(row.Cells(ijk.ColPriceByMultiUnit.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.ColVAT.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.ColAmount2.Index).Value).ToString + "," _
                    + CDec(row.Cells(ijk.ColVAT2.Index).Value).ToString + "" _
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

    Public Function Hapus(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_itemTableAdapter
        Dim DTA1 As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_item_exchangerateTableAdapter
        Dim DTA2 As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_item_salesreturnTableAdapter
        Dim DTA3 As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_itemdpTableAdapter
        Dim DTA_Inv_DO As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_invoice_doTableAdapter
        Dim TBL_Inv_DO As New DseTransaksiSalesInvoice.tbltrans_invoice_doDataTable
        Dim rowInvDO As DseTransaksiSalesInvoice.tbltrans_invoice_doRow

        Try
            'update data do jadi belum buat inv
            DTA_Inv_DO.FillBy_PrimaryInvoice(TBL_Inv_DO, vPrimaryTransaksi)
            For Each rowInvDO In TBL_Inv_DO.Rows
                DTA_Inv_DO.UpdateQuery_SetSudahBuatInvoice(0, -1, rowInvDO.Primary_DO)
            Next
            'hapus data item
            DTA.Delete(vPrimaryTransaksi)
            DTA1.Delete(vPrimaryTransaksi)
            DTA2.Delete(vPrimaryTransaksi)
            DTA3.Delete(vPrimaryTransaksi)
            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
            DTA2.Dispose()
            DTA3.Dispose()
            DTA_Inv_DO.Dispose()
            TBL_Inv_DO.Dispose()
        End Try
    End Function

    Public Function GetManualDP_Data(ByVal vPrimaryInv_DP As Double, ByRef row1 As DseTransaksiSalesInvoice.tbltrans_salesinvoice_itemdp_viewRow) As Boolean
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tblbalance_sales_dp1TableAdapter
        Dim Tbl As New DseTransaksiSalesInvoice.tblbalance_sales_dp1DataTable
        Dim row As DseTransaksiSalesInvoice.tblbalance_sales_dp1Row

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Fill(Tbl, 1, vPrimaryInv_DP) > 0 Then
                row = Tbl.Rows(0)
                row1.No_Invoice_DP = row.No_Inv_DP
                row1.Currency_Code = row.Currency_Code
                row1.Default_Precision = row.Default_Precision
                row1.Sub_Total = row.Sub_Total

                Return True
            End If

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
