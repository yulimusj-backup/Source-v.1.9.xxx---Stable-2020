Option Explicit On
Option Strict Off

Public Class ClsTransaksiSalesReturnItem
    Public Function GetJml(ByVal vPrimaryTransaksi As Double, ByVal vPrimaryProduct As Double, ByVal vSecQty As Double, ByVal vPrimaryLocation As Double, _
    ByVal vBatchNo As String, ByVal vPrimaryPdo As Double, ByVal vExpiry As Date) As Decimal
        Dim DTA As New dseTransaksiSalesReturnTableAdapters.tbltrans_salesreturn_itemTableAdapter
        Dim RetVal As Nullable(Of Decimal)

        Try
            RetVal = DTA.ScalarQuery_GetJml(vPrimaryTransaksi, vPrimaryProduct, vBatchNo, vExpiry, vSecQty, vPrimaryLocation)

            If RetVal.HasValue = True Then
                Return RetVal.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetSalesReturnOrderBalance(ByVal vPrimarySO As Double, _
    ByVal vPrimaryCodeNo As Double) As Double
        Dim DTA As New dseTransaksiSalesReturnTableAdapters.salesreturn_orderbalance_view2TableAdapter

        Try
            Return DTA.ScalarQuery_GetJmlOrderBalance(vPrimarySO, vPrimaryCodeNo).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub IsiTBLSO(ByVal vPrimaryCustomer As Double, ByVal TBL As dseTransaksiSalesReturn.salesreturn_orderbalance_view2DataTable)
        Dim DTA As New dseTransaksiSalesReturnTableAdapters.salesreturn_orderbalance_view2TableAdapter

        Try
            DTA.FillBy_PrimaryCustomer(TBL, vPrimaryCustomer)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsikanTBLProduct(ByVal vPrimarySalesOrder As Double, _
    ByVal TBL As dseTransaksiSalesReturn.salesreturn_orderbalance_view2DataTable)
        Dim DTA As New dseTransaksiSalesReturnTableAdapters.salesreturn_orderbalance_view2TableAdapter

        Try
            DTA.FillBy_PrimarySO(TBL, vPrimarySalesOrder)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsiDGVByPrimarySR(ByVal listSR As List(Of Double), ByVal vDGVItem As DataGridView, ByVal vObjCD As Core_Module.ClsCekData)
        Dim DTA As New dseTransaksiSalesReturnTableAdapters.tbltrans_salesreturn_item_view0TableAdapter
        Dim TBL As New dseTransaksiSalesReturn.tbltrans_salesreturn_item_view0DataTable
        Dim row As dseTransaksiSalesReturn.tbltrans_salesreturn_item_view0Row
        Dim vPrimaryTransaksi As Double = 0
        Dim i As New FrmTransaksiSalesInvoice
        Dim RowDgv As DataGridViewRow
        Dim ObjP As New Core_Module.clsMasterProduct

        Try
            'kosongkan dgv
            vDGVItem.Rows.Clear()

            For Each vPrimaryTransaksi In listSR
                'isikan tabel
                If DTA.FillBy_PrimaryTransaksi(TBL, vPrimaryTransaksi) > 0 Then
                    For Each row In TBL.Rows
                        With row
                            vDGVItem.Rows.Add()
                            RowDgv = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                            RowDgv.Cells(i.SR_ColNo.Index).Value = ""
                            RowDgv.Cells(i.SR_ColCodeNo.Index).Value = row.Product_Code
                            RowDgv.Cells(i.SR_ColDesc.Index).Value = row.Description
                            RowDgv.Cells(i.ColSONumber_SR.Index).Value = row.SO_No
                            RowDgv.Cells(i.ColDONumber_SR.Index).Value = row.DO_Number
                            RowDgv.Cells(i.ColSRNumber_SR.Index).Value = row.SR_Number
                            RowDgv.Cells(i.ColBatchNo_SR.Index).Value = row.Batch_No
                            RowDgv.Cells(i.ColExpiryDate_SR.Index).Value = Format(row.Expiry_Date, "dd MMM yyyy")
                            RowDgv.Cells(i.ColQtyMaster_SR.Index).Value = row.Qty_Master
                            RowDgv.Cells(i.ColQtyPerMaster_SR.Index).Value = row.Qty_Per_Master
                            RowDgv.Cells(i.ColMasterUnit_SR.Index).Value = row.Master_Unit
                            RowDgv.Cells(i.ColQtySR.Index).Value = row.Qty
                            RowDgv.Cells(i.COlTotalQtySR.Index).Value = vObjCD.DigitGrouping(row.Total_Qty, ._Decimal)
                            RowDgv.Cells(i.SR_ColUnits.Index).Value = row.Unit_Code
                            RowDgv.Cells(i.ColSecQty_SR.Index).Value = row.Secondary_Qty
                            RowDgv.Cells(i.ColSecQty_SR.Index).Value = row.Sec_Unit
                            RowDgv.Cells(i.ColRemarks.Index).Value = row.Remarks
                            RowDgv.Cells(i.COlPrimaryCodeNoSR.Index).Value = row.Primary_Product
                            RowDgv.Cells(i.ColPrimarySR_SR.Index).Value = row.Primary_Transaksi
                            RowDgv.Cells(i.SR_ColPrimaryMasterUnit.Index).Value = row.Primary_Master_Units
                            RowDgv.Cells(i.SR_ColPrimaryUnits.Index).Value = row.Primary_Units
                            RowDgv.Cells(i.SR_ColPrimarySecUnits.Index).Value = row.Primary_Secondary_Units
                            RowDgv.Cells(i.ColPrimaryDO_SR.Index).Value = row.Primary_DO
                            RowDgv.Cells(i.ColPrimarySO_SR.Index).Value = row.Primary_SO
                            RowDgv.Cells(i.ColTotalQtymulti_SR.Index).Value = vObjCD.DigitGrouping(row.Total_Qty_Multi, ObjP.GetAngkaDecimal(row.Primary_Product, row.Primary_Units_Multi))
                            RowDgv.Cells(i.ColUnitsCodeMulti_SR.Index).Value = row.Units_Code_Multi
                            RowDgv.Cells(i.ColPrimaryUnitsMulti_SR.Index).Value = row.Primary_Units_Multi
                            RowDgv.Cells(i.ColRatioReference_SR.Index).Value = row.Ratio_Reference
                            RowDgv.Cells(i.ColRatioDefault_SR.Index).Value = row.Ratio_Default
                            If .Primary_Units_Multi <> .Primary_Units Then
                                RowDgv.Cells(i.ColRatioDesc_SR.Index).Value = vObjCD.DigitGrouping(.Ratio_Reference, ObjP.GetAngkaDecimal(row.Primary_Product, row.Primary_Units_Multi)) & " " & .Units_Code_Multi & "=" & vObjCD.DigitGrouping(.Ratio_Default, ._Decimal) & " " & .Unit_Code
                            Else
                                RowDgv.Cells(i.ColRatioDesc_SR.Index).Value = "-"
                            End If

                            'vDGVItem.Rows.Add("0", row.Product_Code, row.Description, _
                            'row.SO_No, row.DO_Number, row.SR_Number, row.Batch_No, row.Expiry_Date, row.Qty_Master, .Qty_Per_Master, .Master_Unit, row.Qty, _
                            'vObjCD.DigitGrouping(.Total_Qty, ._Decimal), row.Unit_Code, row.Secondary_Qty, row.Sec_Unit, row.Remarks, row.Primary_Product, _
                            'row.Primary_Transaksi, row.Primary_Master_Units, row.Primary_Units, row.Primary_Secondary_Units, row.Primary_DO, _
                            'row.Primary_SO)
                        End With
                    Next
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
        End Try
    End Sub

    '@rosadinnn-balikin lagi
    Public Function IsiDGV(ByVal vTblItem As CustomSMESS.DseTransferBeginingBalance.tbltrans_salesreturn_item_copyDataTable, _
            ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double, _
            ByVal ij As FrmTransaksiSalesReturn_Quick) As Boolean
        Dim row As CustomSMESS.DseTransferBeginingBalance.tbltrans_salesreturn_item_copyRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim rowDGV As DataGridViewRow
        Dim rounding As Integer = 0
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim PrimaryCur As Double = -1

        Try
            'kosongkan dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            vDGVItem.Rows.Clear()

            'isikan tabel
            vTblItem.DefaultView.RowFilter = "[Primary Transaksi] = " & vPrimaryTransaksi.ToString
            If vTblItem.DefaultView.Count > 0 Then
                For i As Integer = 0 To vTblItem.DefaultView.Count - 1
                    row = vTblItem.DefaultView(i).Row
                    With row
                        If PrimaryCur <> .Primary_Currency Then
                            rounding = ObjCur.GetDefaultPrecission(.Primary_Currency)
                            PrimaryCur = .Primary_Currency
                        End If
                        vDGVItem.Rows.Add()
                        rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                        rowDGV.Cells(ij.ColNo.Index).Value = "0"
                        'rowDGV.Cells(ij.ColCodeNo.Index).Value = row.Product_Code
                        'rowDGV.Cells(ij.ColDescription.Index).Value = row.Description
                        'rowDGV.Cells(ij.ColSONO.Index).Value = row.SO_No
                        rowDGV.Cells(ij.ColBatchNo.Index).Value = row.Batch_No
                        rowDGV.Cells(ij.ColExpiryDate.Index).Value = row.Expiry_Date
                        'rowDGV.Cells(ij.ColQtyMaster.Index).Value = row.Qty_Master
                        'rowDGV.Cells(ij.ColQtyPerMaster.Index).Value = row.Qty_Per_Master
                        'rowDGV.Cells(ij.ColMasterUnits.Index).Value = row.Master_Unit
                        rowDGV.Cells(ij.ColQty.Index).Value = ObjCD.DigitGrouping(row.Total_Qty) ', ._Decimal)
                        'rowDGV.Cells(ij.ColTotalQt.Index).Value = ObjCD.DigitGrouping(row.Total_Qty, ._Decimal)
                        'rowDGV.Cells(ij.ColSONO.Index).Value = row.SO_No
                        'rowDGV.Cells(ij.ColUnit.Index).Value = row.Unit_Code
                        rowDGV.Cells(ij.ColSecondaryQty.Index).Value = row.Secondary_Qty
                        'rowDGV.Cells(ij.ColSecondaryUnit.Index).Value = row.Sec_Unit
                        rowDGV.Cells(ij.ColRemark.Index).Value = row.Remarks
                        rowDGV.Cells(ij.ColPrimaryUnits.Index).Value = row.Primary_Units
                        rowDGV.Cells(ij.ColPrimaryCodeNo.Index).Value = row.Primary_Product
                        rowDGV.Cells(ij.ColPrimarySO.Index).Value = row.Primary_SO
                        'rowDGV.Cells(ij.ColPrimaryMasterUnits.Index).Value = row.Primary_Master_Units
                        'rowDGV.Cells(ij.ColPrimaryUnit.Index).Value = row.Primary_Units
                        rowDGV.Cells(ij.ColPrimarySecondaryUnits.Index).Value = row.Primary_Secondary_Units
                        'rowDGV.Cells(ij.ColCurrency.Index).Value = row.Currency
                        rowDGV.Cells(ij.ColUnitPrice.Index).Value = ObjCD.DigitGrouping(row.Unit_Price, rounding)
                        rowDGV.Cells(ij.ColDisc.Index).Value = ObjCD.DigitGrouping(row.Discount, 2)
                        rowDGV.Cells(ij.ColNettPrice.Index).Value = ObjCD.DigitGrouping(row.Nett_Price, rounding)
                        rowDGV.Cells(ij.ColAmount.Index).Value = ObjCD.DigitGrouping(row.Amount, rounding)
                        rowDGV.Cells(ij.ColExchangeRate.Index).Value = ObjCD.DigitGrouping(row.Exchange_Rate)
                        rowDGV.Cells(ij.ColPrimaryCurrency.Index).Value = row.Primary_Currency
                        rowDGV.Cells(ij.ColPrimaryLocation.Index).Value = row.Primary_Location
                        'rowDGV.Cells(ij.ColWareHouse.Index).Value = row.Warehouse_Code
                        'rowDGV.Cells(ij.ColLocation.Index).Value = row.Storage_Area_Code
                        rowDGV.Cells(ij.ColPrimaryPdO.Index).Value = -1
                        rowDGV.Cells(ij.ColSuppDONo.Index).Value = ""
                    End With
                Next
            End If

            'set nomer
            'ObjCD.SetNumber(vDGVItem)

            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCD = Nothing
            ObjCur = Nothing
        End Try
    End Function

    Public Function IsiDGV(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double, _
    ByVal ij As FrmTransaksiSalesReturn) As Boolean
        Dim DTA As New dseTransaksiSalesReturnTableAdapters.tbltrans_salesreturn_item_view0TableAdapter
        Dim TBL As New dseTransaksiSalesReturn.tbltrans_salesreturn_item_view0DataTable
        Dim row As dseTransaksiSalesReturn.tbltrans_salesreturn_item_view0Row
        Dim ObjCD As New Core_Module.ClsCekData
        Dim rowDGV As DataGridViewRow
        Dim rounding As Integer = 0
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim PrimaryCur As Double = -1
        Dim DefPrecMulti As Integer = 0
        Dim ObjProduct As New Core_Module.clsMasterProduct

        Try
            'kosongkan dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            vDGVItem.Rows.Clear()

            'isikan tabel
            If DTA.FillBy_PrimaryTransaksi(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    With row
                        If PrimaryCur <> .Primary_Currency Then
                            rounding = ObjCur.GetDefaultPrecission(.Primary_Currency)
                            PrimaryCur = .Primary_Currency
                        End If
                        'vDGVItem.Rows.Add("0", row.Product_Code, row.Description, _
                        'row.SO_No, row.Batch_No, row.Expiry_Date, row.Qty_Master, .Qty_Per_Master, .Master_Unit, row.Qty, _
                        '.Total_Qty, row.Unit_Code, row.Secondary_Qty, row.Sec_Unit, row.Remarks, row.Primary_Product, _
                        'row.Primary_SO, row.Primary_Master_Units, row.Primary_Units, row.Primary_Secondary_Units, row.Primary_DO)
                        DefPrecMulti = ObjProduct.GetAngkaDecimal(.Primary_Product, .Primary_Units_Multi)
                        vDGVItem.Rows.Add()
                        rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                        rowDGV.Cells(ij.ColNo.Index).Value = "0"
                        rowDGV.Cells(ij.ColCodeNo.Index).Value = row.Product_Code
                        rowDGV.Cells(ij.ColDescription.Index).Value = row.Description
                        rowDGV.Cells(ij.ColSONO.Index).Value = row.SO_No
                        rowDGV.Cells(ij.ColBatchNo.Index).Value = row.Batch_No
                        rowDGV.Cells(ij.ColExpiryDate.Index).Value = row.Expiry_Date
                        rowDGV.Cells(ij.ColQtyMaster.Index).Value = row.Qty_Master
                        rowDGV.Cells(ij.ColQtyPerMaster.Index).Value = row.Qty_Per_Master
                        rowDGV.Cells(ij.ColMasterUnits.Index).Value = row.Master_Unit
                        rowDGV.Cells(ij.ColQty.Index).Value = row.Qty
                        rowDGV.Cells(ij.ColTotalQty.Index).Value = ObjCD.DigitGrouping(row.Total_Qty, ._Decimal)
                        rowDGV.Cells(ij.ColSONO.Index).Value = row.SO_No
                        rowDGV.Cells(ij.ColUnit.Index).Value = row.Unit_Code
                        rowDGV.Cells(ij.ColSecQty.Index).Value = row.Secondary_Qty
                        rowDGV.Cells(ij.ColSecUnit.Index).Value = row.Sec_Unit
                        rowDGV.Cells(ij.ColRemarks.Index).Value = row.Remarks
                        rowDGV.Cells(ij.ColPrimaryUnit.Index).Value = row.Primary_Units
                        rowDGV.Cells(ij.ColPrimaryCodeNo.Index).Value = row.Primary_Product
                        rowDGV.Cells(ij.ColPrimarySO.Index).Value = row.Primary_SO
                        rowDGV.Cells(ij.ColPrimaryMasterUnits.Index).Value = row.Primary_Master_Units
                        rowDGV.Cells(ij.ColPrimaryUnit.Index).Value = row.Primary_Units
                        rowDGV.Cells(ij.ColPrimarySecUnit.Index).Value = row.Primary_Secondary_Units
                        rowDGV.Cells(ij.ColCurrency.Index).Value = row.Currency
                        rowDGV.Cells(ij.ColUnitPrice.Index).Value = ObjCD.DigitGrouping(row.Unit_Price, rounding)
                        rowDGV.Cells(ij.ColDisc.Index).Value = ObjCD.DigitGrouping(row.Discount, 2)
                        rowDGV.Cells(ij.ColDiscAmount.Index).Value = ObjCD.DigitGrouping(row.Disc_Amount, rounding)
                        rowDGV.Cells(ij.ColDiscFromPerct.Index).Value = row.Disc_From_Perct
                        rowDGV.Cells(ij.ColNettPrice.Index).Value = ObjCD.DigitGrouping(row.Nett_Price, rounding)
                        rowDGV.Cells(ij.ColAmount.Index).Value = ObjCD.DigitGrouping(row.Amount, rounding)
                        rowDGV.Cells(ij.ColExchangeRate.Index).Value = ObjCD.DigitGrouping(row.Exchange_Rate, 4)
                        rowDGV.Cells(ij.ColPrimaryCurrency.Index).Value = row.Primary_Currency

                        rowDGV.Cells(ij.ColTotalQtyMulti.Index).Value = ObjCD.DigitGrouping(.Total_Qty_Multi, DefPrecMulti)
                        rowDGV.Cells(ij.ColUnitsMulti.Index).Value = .Units_Code_Multi
                        rowDGV.Cells(ij.ColPrimaryUnitsMulti.Index).Value = .Primary_Units_Multi
                        rowDGV.Cells(ij.ColRatioReference.Index).Value = .Ratio_Reference
                        rowDGV.Cells(ij.ColRatioDesc.Index).Value = .Ratio_Default
                        If .Primary_Units <> .Primary_Units_Multi Then
                            rowDGV.Cells(ij.ColRatioDesc.Index).Value = ObjCD.DigitGrouping(.Ratio_Reference, DefPrecMulti) & " " & .Units_Code_Multi & "=" & ObjCD.DigitGrouping(.Ratio_Default, ._Decimal) & " " & .Unit_Code
                        Else
                            rowDGV.Cells(ij.ColRatioDesc.Index).Value = "-"
                        End If
                    End With
                Next
            End If

            'set nomer
            ObjCD.SetNumber(vDGVItem)

            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
            ObjCD = Nothing
            ObjCur = Nothing
        End Try
    End Function

    Public Function IsiDGV(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double, _
    ByVal ij As FrmTransaksiSalesReturn_Quick) As Boolean
        Dim DTA As New dseTransaksiSalesReturnTableAdapters.tbltrans_salesreturn_item_view0TableAdapter
        Dim TBL As New dseTransaksiSalesReturn.tbltrans_salesreturn_item_view0DataTable
        Dim row As dseTransaksiSalesReturn.tbltrans_salesreturn_item_view0Row
        Dim ObjCD As New Core_Module.ClsCekData
        Dim rowDGV As DataGridViewRow
        Dim rounding As Integer = 0
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim PrimaryCur As Double = -1
        Dim DefPrecMulti As Integer = 0
        Dim DefPrecDefault As Integer = 0
        Dim ObjP As New Core_Module.clsMasterProduct

        Try
            'kosongkan dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            vDGVItem.Rows.Clear()

            'isikan tabel
            If DTA.FillBy_PrimaryTransaksi(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    With row
                        If PrimaryCur <> .Primary_Currency Then
                            rounding = ObjCur.GetDefaultPrecission(.Primary_Currency)
                            PrimaryCur = .Primary_Currency
                        End If
                        'vDGVItem.Rows.Add("0", row.Product_Code, row.Description, _
                        'row.SO_No, row.Batch_No, row.Expiry_Date, row.Qty_Master, .Qty_Per_Master, .Master_Unit, row.Qty, _
                        '.Total_Qty, row.Unit_Code, row.Secondary_Qty, row.Sec_Unit, row.Remarks, row.Primary_Product, _
                        'row.Primary_SO, row.Primary_Master_Units, row.Primary_Units, row.Primary_Secondary_Units, row.Primary_DO)
                        vDGVItem.Rows.Add()
                        rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                        rowDGV.Cells(ij.ColNo.Index).Value = "0"
                        rowDGV.Cells(ij.ColCodeNo.Index).Value = row.Product_Code
                        rowDGV.Cells(ij.ColDescription.Index).Value = row.Description
                        rowDGV.Cells(ij.ColSONO.Index).Value = row.SO_No
                        rowDGV.Cells(ij.ColBatchNo.Index).Value = row.Batch_No
                        rowDGV.Cells(ij.ColExpiryDate.Index).Value = row.Expiry_Date
                        'rowDGV.Cells(ij.ColQtyMaster.Index).Value = row.Qty_Master
                        'rowDGV.Cells(ij.ColQtyPerMaster.Index).Value = row.Qty_Per_Master
                        'rowDGV.Cells(ij.ColMasterUnits.Index).Value = row.Master_Unit
                        rowDGV.Cells(ij.ColQty.Index).Value = ObjCD.DigitGrouping(row.Total_Qty, ._Decimal)
                        'rowDGV.Cells(ij.ColTotalQt.Index).Value = ObjCD.DigitGrouping(row.Total_Qty, ._Decimal)
                        rowDGV.Cells(ij.ColSONO.Index).Value = row.SO_No
                        rowDGV.Cells(ij.ColUnit.Index).Value = row.Unit_Code
                        rowDGV.Cells(ij.ColSecondaryQty.Index).Value = row.Secondary_Qty
                        rowDGV.Cells(ij.ColSecondaryUnit.Index).Value = row.Sec_Unit
                        rowDGV.Cells(ij.ColRemark.Index).Value = row.Remarks
                        rowDGV.Cells(ij.ColPrimaryUnits.Index).Value = row.Primary_Units
                        rowDGV.Cells(ij.ColPrimaryCodeNo.Index).Value = row.Primary_Product
                        rowDGV.Cells(ij.ColPrimarySO.Index).Value = row.Primary_SO
                        'rowDGV.Cells(ij.ColPrimaryMasterUnits.Index).Value = row.Primary_Master_Units
                        'rowDGV.Cells(ij.ColPrimaryUnit.Index).Value = row.Primary_Units
                        rowDGV.Cells(ij.ColPrimarySecondaryUnits.Index).Value = row.Primary_Secondary_Units
                        rowDGV.Cells(ij.ColCurrency.Index).Value = row.Currency
                        rowDGV.Cells(ij.ColUnitPrice.Index).Value = ObjCD.DigitGrouping(row.Unit_Price, rounding)
                        rowDGV.Cells(ij.ColDisc.Index).Value = ObjCD.DigitGrouping(row.Discount, 2)
                        rowDGV.Cells(ij.ColDiscAmount.Index).Value = ObjCD.DigitGrouping(row.Disc_Amount, rounding)
                        rowDGV.Cells(ij.ColDiscFromPerct.Index).Value = row.Disc_From_Perct
                        rowDGV.Cells(ij.ColNettPrice.Index).Value = ObjCD.DigitGrouping(row.Nett_Price, rounding)
                        rowDGV.Cells(ij.ColAmount.Index).Value = ObjCD.DigitGrouping(row.Amount, rounding)
                        rowDGV.Cells(ij.ColExchangeRate.Index).Value = ObjCD.DigitGrouping(row.Exchange_Rate)
                        rowDGV.Cells(ij.ColPrimaryCurrency.Index).Value = row.Primary_Currency
                        rowDGV.Cells(ij.ColPrimaryLocation.Index).Value = row.Primary_Location
                        rowDGV.Cells(ij.ColWareHouse.Index).Value = row.Warehouse_Code
                        rowDGV.Cells(ij.ColLocation.Index).Value = row.Storage_Area_Code
                        rowDGV.Cells(ij.ColPrimaryPdO.Index).Value = -1
                        rowDGV.Cells(ij.ColSuppDONo.Index).Value = ""

                        DefPrecMulti = ObjP.GetAngkaDecimal(.Primary_Product, .Primary_Units_Multi)
                        rowDGV.Cells(ij.ColQtyMulti.Index).Value = ObjCD.DigitGrouping(.Total_Qty_Multi, DefPrecMulti)
                        rowDGV.Cells(ij.ColUnitsMulti.Index).Value = .Units_Code_Multi
                        rowDGV.Cells(ij.ColPrimaryUnitsMulti.Index).Value = .Primary_Units_Multi
                        rowDGV.Cells(ij.ColRatioReference.Index).Value = ObjCD.DigitGrouping(.Ratio_Reference, DefPrecMulti)
                        rowDGV.Cells(ij.ColRatioDefault.Index).Value = ObjCD.DigitGrouping(.Ratio_Default, ._Decimal)
                        If .Primary_Units <> .Primary_Units_Multi Then
                            rowDGV.Cells(ij.ColRatioDesc.Index).Value = ObjCD.DigitGrouping(row.Ratio_Reference, DefPrecMulti) _
                                    & " " & .Units_Code_Multi & " = " & ObjCD.DigitGrouping(row.Ratio_Default, row._Decimal) & " " & row.Unit_Code
                        Else
                            rowDGV.Cells(ij.ColRatioDesc.Index).Value = "-"
                        End If
                    End With
                Next
            End If

            'set nomer
            ObjCD.SetNumber(vDGVItem)

            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
            ObjCD = Nothing
            ObjCur = Nothing
        End Try
    End Function

    Public Function Simpan(ByVal vdgvItem As DataGridView, ByVal vPrimaryTransaksi As Double, ByVal vFrm As FrmTransaksiSalesReturn) As Boolean
        Dim row As DataGridViewRow
        Dim DTA As New dseTransaksiSalesReturnTableAdapters.tbltrans_salesreturn_itemTableAdapter
        Dim ObjProd As New Core_Module.clsMasterProduct
        Dim Sec_Qty As Double = 0
        Dim BatchNo As String = ""
        Dim ExpiryDate As Date = #1/1/1950#
        Dim ListPrimarySO As New List(Of Double)
        Dim ObjSOItem As New ClsTransaksiSalesOrderItem
        Dim i As New FrmTransaksiSalesReturn

        Try
            For Each row In vdgvItem.Rows
                BatchNo = row.Cells(i.ColBatchNo.Index).Value
                If IsNumeric(row.Cells(i.ColSecQty.Index).Value) = True Then
                    Sec_Qty = row.Cells(i.ColSecQty.Index).Value
                Else
                    Sec_Qty = "-1"
                End If
                If IsDate(row.Cells(i.ColExpiryDate.Index).Value) = False Then
                    ExpiryDate = #1/1/1950#
                Else
                    ExpiryDate = CDate(row.Cells(i.ColExpiryDate.Index).Value)
                End If
                'simpan ke basis data item
                With row
                    DTA.Insert(vPrimaryTransaksi, _
                    row.Cells(i.ColPrimaryCodeNo.Index).Value, _
                    row.Cells(i.ColPrimarySO.Index).Value, _
                    BatchNo, _
                    ExpiryDate, _
                    row.Cells(i.ColQtyMaster.Index).Value, _
                    row.Cells(i.ColQtyPerMaster.Index).Value, _
                    row.Cells(i.ColPrimaryMasterUnits.Index).Value, _
                    row.Cells(i.ColQty.Index).Value, _
                    row.Cells(i.ColTotalQty.Index).Value, _
                    row.Cells(i.ColPrimaryUnit.Index).Value, _
                    Sec_Qty, _
                    row.Cells(i.ColPrimarySecUnit.Index).Value, _
                    row.Cells(i.ColRemarks.Index).Value, _
                    row.Cells(i.ColUnitPrice.Index).Value, _
                    row.Cells(i.ColDisc.Index).Value, _
                    row.Cells(i.ColNettPrice.Index).Value, _
                    row.Cells(i.ColAmount.Index).Value, _
                    row.Cells(i.ColExchangeRate.Index).Value, _
                    row.Cells(i.ColPrimaryCurrency.Index).Value, _
                    -1, _
                    row.Cells(i.ColDiscAmount.Index).Value, _
                    row.Cells(i.ColDiscFromPerct.Index).Value, _
                    row.Cells(i.ColTotalQtyMulti.Index).Value, _
                    row.Cells(i.ColPrimaryUnitsMulti.Index).Value, _
                    row.Cells(i.ColRatioReference.Index).Value, _
                    row.Cells(i.ColRatioDefault.Index).Value)
                End With

                'DTA.Insert(vPrimaryTransaksi, row.Cells(15).Value, row.Cells(16).Value, BatchNo, _
                'ExpiryDate, row.Cells(6).Value, row.Cells(7).Value, row.Cells(17).Value, _
                'row.Cells(9).Value, row.Cells(10).Value, row.Cells(18).Value, Sec_Qty, row.Cells(19).Value, _
                'row.Cells(14).Value, row.Cells(vFrm.ColUnitPrice.Index).Value, row.Cells(vFrm.ColDisc.Index).Value, _
                'row.Cells(vFrm.ColNettPrice.Index).Value, row.Cells(vFrm.ColAmount.Index).Value, row.Cells(vFrm.ColExchangeRate.Index).Value, _
                'row.Cells(vFrm.ColPrimaryCurrency.Index).Value, -1, row.Cells(vFrm.ColDiscAmount.Index).Value, row.Cells(vFrm.ColDiscFromPerct.Index).Value)

                If ListPrimarySO.Contains(row.Cells(i.ColPrimarySO.Index).Value) = False Then
                    ListPrimarySO.Add(row.Cells(i.ColPrimarySO.Index).Value)
                End If
            Next
            'hitung ulang nilai order balance
            ObjSOItem.Set_Order_Balance(ListPrimarySO)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjProd = Nothing
            ObjSOItem = Nothing
            ListPrimarySO = Nothing
        End Try
    End Function

    Public Function Simpan(ByVal vdgvItem As DataGridView, ByVal vPrimaryTransaksi As Double, ByVal vFrm As FrmTransaksiSalesReturn_Quick) As Boolean
        Dim row As DataGridViewRow
        Dim DTA As New dseTransaksiSalesReturnTableAdapters.tbltrans_salesreturn_itemTableAdapter
        Dim ObjProd As New Core_Module.clsMasterProduct
        Dim Sec_Qty As Double = 0
        Dim BatchNo As String = ""
        Dim ExpiryDate As Date = #1/1/1950#
        Dim ListPrimarySO As New List(Of Double)
        Dim ObjSOItem As New ClsTransaksiSalesOrderItem
        Dim i As New FrmTransaksiSalesReturn_Quick

        Try
            For Each row In vdgvItem.Rows
                BatchNo = row.Cells(vFrm.ColBatchNo.Index).Value
                If IsNumeric(row.Cells(vFrm.ColSecondaryQty.Index).Value) = True Then
                    Sec_Qty = row.Cells(vFrm.ColSecondaryQty.Index).Value
                Else
                    Sec_Qty = "-1"
                End If
                If IsDate(row.Cells(vFrm.ColExpiryDate.Index).Value) = False Then
                    ExpiryDate = #1/1/1950#
                Else
                    ExpiryDate = CDate(row.Cells(vFrm.ColExpiryDate.Index).Value)
                End If

                With row
                    DTA.Insert(vPrimaryTransaksi, _
                    row.Cells(i.ColPrimaryCodeNo.Index).Value, _
                    row.Cells(i.ColPrimarySO.Index).Value, _
                    BatchNo, _
                    ExpiryDate, _
                    0, _
                    0, _
                    -1, _
                    0, _
                    row.Cells(i.ColQty.Index).Value, _
                    row.Cells(i.ColPrimaryUnits.Index).Value, _
                    Sec_Qty, _
                    row.Cells(i.ColPrimarySecondaryUnits.Index).Value, _
                    row.Cells(i.ColRemark.Index).Value, _
                    row.Cells(i.ColUnitPrice.Index).Value, _
                    row.Cells(i.ColDisc.Index).Value, _
                    row.Cells(i.ColNettPrice.Index).Value, _
                    row.Cells(i.ColAmount.Index).Value, _
                    row.Cells(i.ColExchangeRate.Index).Value, _
                    row.Cells(i.ColPrimaryCurrency.Index).Value, _
                    row.Cells(i.ColPrimaryLocation.Index).Value, _
                    row.Cells(i.ColDiscAmount.Index).Value, _
                    row.Cells(i.ColDiscFromPerct.Index).Value, _
                    row.Cells(i.ColQtyMulti.Index).Value, _
                    row.Cells(i.ColPrimaryUnitsMulti.Index).Value, _
                    row.Cells(i.ColRatioReference.Index).Value, _
                    row.Cells(i.ColRatioDefault.Index).Value)
                End With

                'simpan ke basis data item 
                'DTA.Insert(vPrimaryTransaksi, row.Cells(vFrm.ColPrimaryCodeNo.Index).Value, -1, BatchNo, _
                'ExpiryDate, 0, 0, -1, 0, row.Cells(vFrm.ColQty.Index).Value, _
                'row.Cells(vFrm.ColPrimaryUnits.Index).Value, Sec_Qty, row.Cells(vFrm.ColPrimarySecondaryUnits.Index).Value, _
                'row.Cells(vFrm.ColRemark.Index).Value, row.Cells(vFrm.ColUnitPrice.Index).Value, row.Cells(vFrm.ColDisc.Index).Value, _
                'row.Cells(vFrm.ColNettPrice.Index).Value, row.Cells(vFrm.ColAmount.Index).Value, row.Cells(vFrm.ColExchangeRate.Index).Value, _
                'row.Cells(vFrm.ColPrimaryCurrency.Index).Value,  row.Cells(vFrm.ColPrimaryLocation.Index).Value)
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjProd = Nothing
            ObjSOItem = Nothing
            ListPrimarySO = Nothing
        End Try
    End Function

    Public Function Hapus(ByVal vPrimaryTransaksi As Double, ByVal ObjSR As ClsTransaksiSalesReturn) As Boolean
        Dim DTA As New dseTransaksiSalesReturnTableAdapters.tbltrans_salesreturn_itemTableAdapter
        Dim ObjSOItem As New ClsTransaksiSalesOrderItem
        Dim ListPrimarySO As New List(Of Double)

        Try
            ObjSR.GetListPrimarySO(vPrimaryTransaksi, ListPrimarySO)
            'hapus data item
            DTA.Delete(vPrimaryTransaksi)
            'set order balance
            ObjSOItem.Set_Order_Balance(ListPrimarySO)
            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjSOItem = Nothing
            ListPrimarySO = Nothing
        End Try
    End Function
End Class
