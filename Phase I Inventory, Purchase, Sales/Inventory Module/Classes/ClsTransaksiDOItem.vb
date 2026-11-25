Option Explicit On
Option Strict Off

Public Class ClsTransaksiDOItem
    Public Sub IsiListNoSO(ByVal vPrimaryDO As Double, ByVal vPrimaryProduct As Double, ByVal TBL As DseTransaksiDO.tbltrans_soDataTable)
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_soTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, CDec(vPrimaryDO), CDec(vPrimaryProduct))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function Simpan_DariFormStockInOut(ByVal vdgvItem As DataGridView, ByVal vPrimaryTransaksi As Double, _
    ByVal TBL_DO_Item As dseTransaksiStockInOut_CoverMinus.tbltrans_do_itemDataTable, Optional ByVal SetOrderBalance As Boolean = True) As Boolean
        Dim row As DataGridViewRow
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do_itemTableAdapter
        Dim ObjProd As New Core_Module.clsMasterProduct
        Dim ObjLS As New Core_Module.clsMasterListedSoftware
        Dim DTA1 As New Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter
        Dim IsInventoryInstalled As Boolean = ObjLS.IsInventoryInstalled
        Dim Sec_Qty, AuthorizedLevelNeeded As Double
        Dim BatchNo As String = ""
        Dim ExpiryDate As Date = #1/1/1950#
        Dim ListPrimarySO As New List(Of Double)
        Dim ObjSOItem As New ClsTransaksiSalesOrder
        Dim FrmStockOutSales As New FrmTransaksiStockOutSales
        Dim row_DO_Item As dseTransaksiStockInOut_CoverMinus.tbltrans_do_itemRow
        Dim DeliveryDate As Date = #1/1/1950#

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Core_Module.SetConstring_DTAProductQty(DTA1)
            For Each row In vdgvItem.Rows
                BatchNo = row.Cells(FrmStockOutSales.ColBatch.Index).Value
                Sec_Qty = "-1"
                If IsDate(row.Cells(FrmStockOutSales.ColExpiredDate.Index).Value) = False Then
                    ExpiryDate = #1/1/1950#
                Else
                    ExpiryDate = CDate(row.Cells(FrmStockOutSales.ColExpiredDate.Index).Value)
                End If
                If IsInventoryInstalled = True Then
                    AuthorizedLevelNeeded = 1
                End If
                DeliveryDate = #1/1/1950#
                For Each row_DO_Item In TBL_DO_Item.Rows
                    If row_DO_Item.Primary_SO = row.Cells(FrmStockOutSales.ColPrimaryPOnoDoNo.Index).Value And _
                    row_DO_Item.Primary_Product = row.Cells(FrmStockOutSales.ColPrimaryCodeNo.Index).Value Then
                        DeliveryDate = row_DO_Item.Delivery_Date
                    End If
                Next
                'simpan ke basis data item
                DTA.Insert(vPrimaryTransaksi, row.Cells(FrmStockOutSales.ColPrimaryCodeNo.Index).Value, row.Cells(FrmStockOutSales.ColPrimaryPOnoDoNo.Index).Value, _
                BatchNo, ExpiryDate, 0, 0, -1, _
                0, row.Cells(FrmStockOutSales.ColQty.Index).Value, row.Cells(FrmStockOutSales.ColPrimaryUnits.Index).Value, Sec_Qty, -1, _
                row.Cells(FrmStockOutSales.ColRemark.Index).Value, row.Cells(FrmStockOutSales.ColQtyMulti.Index).Value, row.Cells(FrmStockOutSales.ColPrimaryUnitsMulti.Index).Value, _
                row.Cells(FrmStockOutSales.ColRatioReference.Index).Value, row.Cells(FrmStockOutSales.ColRatioDefault.Index).Value, _
                DeliveryDate, row.Cells(FrmStockOutSales.ColPrimaryLocation.Index).Value)
                If ListPrimarySO.Contains(row.Cells(FrmStockOutSales.ColPrimaryPOnoDoNo.Index).Value) = False Then
                    ListPrimarySO.Add(row.Cells(FrmStockOutSales.ColPrimaryPOnoDoNo.Index).Value)
                End If
            Next
            If SetOrderBalance = True Then
                'hitung ulang nilai order balance
                ObjSOItem.Set_Order_Balance(ListPrimarySO)
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
            ObjProd = Nothing
            ObjLS = Nothing
            ObjSOItem = Nothing
            ListPrimarySO = Nothing
            FrmStockOutSales = Nothing
        End Try
    End Function

    Public Function SimpanItemDO(ByVal vdgvItem As DataGridView, ByVal vPrimaryTransaksi As Double, Optional ByVal SetOrderBalance As Boolean = True) As Boolean
        Dim vFrm As New FrmTransaksiDO3
        Dim row As DataGridViewRow
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do_itemTableAdapter
        Dim ObjProd As New Core_Module.clsMasterProduct
        Dim ObjLS As New Core_Module.clsMasterListedSoftware
        Dim DTA1 As New Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter
        Dim IsInventoryInstalled As Boolean = ObjLS.IsInventoryInstalled
        Dim Sec_Qty, AuthorizedLevelNeeded As Double
        Dim BatchNo As String = ""
        Dim ExpiryDate As Date = #1/1/1950#
        Dim ListPrimarySO As New List(Of Double)
        Dim ObjSOItem As New ClsTransaksiSalesOrder

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Core_Module.SetConstring_DTAProductQty(DTA1)
            For Each row In vdgvItem.Rows
                BatchNo = row.Cells(vFrm.ColBatchNo.Index).Value
                If IsNumeric(row.Cells(vFrm.ColSecQty.Index).Value) = True Then
                    Sec_Qty = row.Cells(vFrm.ColSecQty.Index).Value
                Else
                    Sec_Qty = "-1"
                End If
                If IsDate(row.Cells(vFrm.ColExpiryDate.Index).Value) = False Then
                    ExpiryDate = #1/1/1950#
                Else
                    ExpiryDate = CDate(row.Cells(vFrm.ColExpiryDate.Index).Value)
                End If
                If IsInventoryInstalled = True Then
                    AuthorizedLevelNeeded = 1
                End If

                'simpan ke basis data item
                DTA.Insert(vPrimaryTransaksi, row.Cells(vFrm.ColPrimaryCodeNo.Index).Value, row.Cells(vFrm.ColPrimarySO.Index).Value, _
                row.Cells(vFrm.ColBatchNo.Index).Value, ExpiryDate, row.Cells(vFrm.ColQtyMaster.Index).Value, _
                row.Cells(vFrm.ColQtyPerMaster.Index).Value, row.Cells(vFrm.ColPrimaryMasterUnits.Index).Value, row.Cells(vFrm.ColQty.Index).Value, _
                row.Cells(vFrm.ColTotalQty.Index).Value, row.Cells(vFrm.ColPrimaryUnit.Index).Value, Sec_Qty, _
                row.Cells(vFrm.ColPrimarySecUnit.Index).Value, row.Cells(vFrm.ColRemarks.Index).Value, row.Cells(vFrm.ColTotalQtyMulti.Index).Value, _
                row.Cells(vFrm.ColPrimaryUnitsMulti.Index).Value, row.Cells(vFrm.ColRatioReference.Index).Value, _
                row.Cells(vFrm.ColRatioDefault.Index).Value, #1/1/1950#, -1)
                If ListPrimarySO.Contains(row.Cells(vFrm.ColPrimarySO.Index).Value) = False Then
                    ListPrimarySO.Add(row.Cells(vFrm.ColPrimarySO.Index).Value)
                End If
            Next
            If SetOrderBalance = True Then
                'hitung ulang nilai order balance
                ObjSOItem.Set_Order_Balance(ListPrimarySO)
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
            ObjProd = Nothing
            ObjLS = Nothing
            ObjSOItem = Nothing
            ListPrimarySO = Nothing
        End Try
    End Function

    Public Function Hapus(ByVal vPrimaryTransaksi As Double, ByVal vObjDO As ClsTransaksiDO) As Boolean
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do_itemTableAdapter
        Dim vListPrimary As New List(Of Double)
        Dim ObjSOItem As New ClsTransaksiSalesOrder

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            vObjDO.GetPrimarySO(vPrimaryTransaksi, vListPrimary)
            'hapus data item
            DTA.Delete(vPrimaryTransaksi)
            'set nilai order balance
            ObjSOItem.Set_Order_Balance(vListPrimary)
            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            vListPrimary = Nothing
            ObjSOItem = Nothing
        End Try
    End Function

    Public Function Hapus(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do_itemTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'hapus data item
            DTA.Delete(CDec(vPrimaryTransaksi))
            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDGV_Default(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do_item_view1TableAdapter
        Dim TBL As New DseTransaksiDO.tbltrans_do_item_view1DataTable
        Dim row As DseTransaksiDO.tbltrans_do_item_view1Row
        Dim Sec_Units As String = ""
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'kosongkan dgv
            vDGVItem.Rows.Clear()
            'isikan tabel
            If DTA.Fill(TBL, CDec(vPrimaryTransaksi)) > 0 Then
                For Each row In TBL.Rows
                    If row.Primary_Secondary_Units <> -1 Then
                        Sec_Units = row.Secondary_Unit
                    Else
                        Sec_Units = ""
                    End If

                    With row
                        vDGVItem.Rows.Add("0", row.Product_Code, row.Product_Desc, "", "", _
                        row.SO_No, row.Batch_No, row.Expiry_Date, row.Qty_Master, .Qty_Per_Master, .Master_Unit, ObjCD.DigitGrouping(row.Qty), _
                        ObjCD.DigitGrouping(.Total_Qty.ToString, ._decimal), row.Unit_Code, row.Secondary_Qty, Sec_Units, row.Remarks, row.Primary_Product, _
                        row.Primary_SO, row.Primary_Master_Units, row.Primary_Units, row.Primary_Secondary_Units)
                    End With
                Next
            End If
            'set nomer
            ObjCD.SetNumber(vDGVItem)
            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
            ObjCD = Nothing
        End Try
    End Function

    Public Function IsiDGV_Multi(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do_item_view1TableAdapter
        Dim TBL As New DseTransaksiDO.tbltrans_do_item_view1DataTable
        Dim row As DseTransaksiDO.tbltrans_do_item_view1Row
        Dim Sec_Units As String = ""
        Dim ObjCD As New Core_Module.ClsCekData

        Dim RowDgv As DataGridViewRow
        Dim VFrm As New FrmTransaksiStockOutSales
        Dim ObjProduct As New Core_Module.clsMasterProduct
        Dim DefPrecMulti As Integer = 0
        Dim TBLMultiUnit As New DseTransaksiDO.DataTable3DataTable
        Dim rowMultiUnit As DseTransaksiDO.DataTable3Row

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'kosongkan dgv
            vDGVItem.Rows.Clear()
            'isikan tabel
            If DTA.Fill(TBL, CDec(vPrimaryTransaksi)) > 0 Then
                For Each row In TBL.Rows
                    If row.Primary_Secondary_Units <> -1 Then
                        Sec_Units = row.Secondary_Unit
                    Else
                        Sec_Units = ""
                    End If

                    With row
                        If .Primary_Units = .Primary_Units_Multi Then
                            DefPrecMulti = ._decimal
                        Else
                            TBLMultiUnit.DefaultView.RowFilter = "[Primary Product] = " + .Primary_Product.ToString + " AND [Primary Units Multi] = " + .Primary_Units_Multi.ToString
                            If TBLMultiUnit.DefaultView.Count = 0 Then
                                DefPrecMulti = ObjProduct.GetAngkaDecimal(.Primary_Product, .Primary_Units_Multi)
                                rowMultiUnit = TBLMultiUnit.NewDataTable3Row
                                With rowMultiUnit
                                    .Primary_Product = row.Primary_Product
                                    .Primary_Units_Multi = row.Primary_Units_Multi
                                    .Decimal_Point = DefPrecMulti
                                End With
                                TBLMultiUnit.Rows.Add(rowMultiUnit)
                            Else
                                rowMultiUnit = TBLMultiUnit.DefaultView(0).Row
                                DefPrecMulti = rowMultiUnit.Decimal_Point
                            End If
                        End If

                        vDGVItem.Rows.Add()
                        RowDgv = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                        RowDgv.Cells(VFrm.ColDO_No.Index).Value = ""
                        RowDgv.Cells(VFrm.ColDO_CodeNo.Index).Value = .Product_Code
                        RowDgv.Cells(VFrm.ColDO_Description.Index).Value = .Product_Desc
                        RowDgv.Cells(VFrm.ColDO_CustCodeNo.Index).Value = ""
                        RowDgv.Cells(VFrm.ColDO_CustDesc.Index).Value = ""
                        RowDgv.Cells(VFrm.ColDO_SONO.Index).Value = .SO_No
                        RowDgv.Cells(VFrm.ColDO_BatchNo.Index).Value = .Batch_No
                        RowDgv.Cells(VFrm.ColDO_ExpiryDate.Index).Value = Format(.Expiry_Date, "dd MMM yyyy")
                        RowDgv.Cells(VFrm.ColDO_QtyMaster.Index).Value = .Qty_Master
                        RowDgv.Cells(VFrm.ColDO_QtyPerMaster.Index).Value = .Qty_Per_Master
                        RowDgv.Cells(VFrm.ColDO_MasterUnits.Index).Value = .Master_Unit
                        RowDgv.Cells(VFrm.ColDO_Qty.Index).Value = .Qty
                        RowDgv.Cells(VFrm.ColDO_TotalQty.Index).Value = ObjCD.DigitGrouping(.Total_Qty, ._decimal)
                        RowDgv.Cells(VFrm.ColDO_Unit.Index).Value = .Unit_Code
                        RowDgv.Cells(VFrm.ColDO_SecQty.Index).Value = .Secondary_Qty
                        RowDgv.Cells(VFrm.ColDO_SecUnit.Index).Value = .Secondary_Unit
                        RowDgv.Cells(VFrm.ColDO_Remarks.Index).Value = .Remarks
                        RowDgv.Cells(VFrm.ColDO_PrimaryCodeNo.Index).Value = .Primary_Product
                        RowDgv.Cells(VFrm.ColDO_PrimarySO.Index).Value = .Primary_SO
                        RowDgv.Cells(VFrm.ColDO_PrimaryMasterUnits.Index).Value = .Primary_Master_Units
                        RowDgv.Cells(VFrm.ColDO_PrimaryUnit.Index).Value = .Primary_Units
                        RowDgv.Cells(VFrm.ColDO_PrimarySecUnit.Index).Value = .Primary_Secondary_Units
                        RowDgv.Cells(VFrm.ColDO_TotalQtyMulti.Index).Value = ObjCD.DigitGrouping(.Total_Qty_Multi, DefPrecMulti)
                        RowDgv.Cells(VFrm.ColDO_UnitsCodeMulti.Index).Value = .Units_Code_Multi
                        RowDgv.Cells(VFrm.ColDO_RatioDefault.Index).Value = ObjCD.DigitGrouping(.Ratio_Default, ._decimal)
                        RowDgv.Cells(VFrm.ColDO_RatioReference.Index).Value = ObjCD.DigitGrouping(.Ratio_Reference, DefPrecMulti)
                        RowDgv.Cells(VFrm.ColDO_PrimaryUnitsMulti.Index).Value = .Primary_Units_Multi
                        RowDgv.Cells(VFrm.ColDO_PrimaryLocation.Index).Value = .primary_location
                        RowDgv.Cells(VFrm.ColDO_Location.Index).Value = .Storage_Area_Code
                        RowDgv.Cells(VFrm.ColDO_Warehouse.Index).Value = .Warehouse_Code

                        If .Primary_Units_Multi <> .Primary_Units Then
                            RowDgv.Cells(VFrm.ColDO_RatioDesc.Index).Value = ObjCD.DigitGrouping(.Ratio_Reference, DefPrecMulti) _
                            & " " & .Units_Code_Multi & " = " & ObjCD.DigitGrouping(.Ratio_Default, ._decimal) & " " & .Unit_Code
                        Else
                            RowDgv.Cells(VFrm.ColDO_RatioDesc.Index).Value = "-"
                        End If
                    End With
                Next
            End If
            'set nomer
            ObjCD.SetNumber(vDGVItem)
            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
            ObjCD = Nothing
        End Try
    End Function

    Public Sub IsikanTBLProductBySO(ByVal vPrimaryDO As Double, ByVal TBL As DseTransaksiDO.DataTable2DataTable, _
    ByVal vPrimarySO As Double)
        Dim DTA As New DseTransaksiDOTableAdapters.DataTable2TableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_PrimarySO(TBL, vPrimaryDO, vPrimarySO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsikanTBLSO(ByVal vPrimaryDO As Double, ByVal TBL As DseTransaksiDO.DataTable2DataTable)
        Dim DTA As New DseTransaksiDOTableAdapters.DataTable2TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_PrimaryDO(TBL, vPrimaryDO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function GetJml4(ByVal vPrimaryDO As Double, ByVal vPrimarySO As Double, _
    ByVal vPrimaryProduct As Double, ByVal vBatchNo As String) As Decimal
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do_itemTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_GetJml4(vPrimaryDO, vPrimaryProduct, vPrimarySO, vBatchNo)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal vPrimaryDO As Double, ByVal vPrimarySO As Double, ByVal vPrimaryProduct As Double) As Decimal
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do_itemTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_GetJml(vPrimaryDO, vPrimaryProduct, vPrimarySO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub IsiSecondaryQty(ByVal vPrimaryDO As Double, ByVal vPrimaryProduct As Double, _
   ByVal vBatchNo As String, ByVal TBL As DseTransaksiDO.tbltrans_do_item2DataTable)
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do_item2TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, vPrimaryDO, vPrimaryProduct, vBatchNo)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsiTBLBatchNo(ByVal vPrimaryDO As Double, ByVal vPrimaryProduct As Double, _
    ByVal TBL As DseTransaksiDO.tbltrans_do_item1DataTable)
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do_item1TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, vPrimaryProduct, vPrimaryDO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    ''' <summary>
    ''' mengembalikan jumlah pending do
    ''' </summary>
    ''' <param name="vPrimaryCodeNo"></param>
    ''' <param name="vBatchNo"></param>
    ''' <param name="vSecQty"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetJmlPendingDO3(ByVal vPrimaryCodeNo As Double, ByVal vBatchNo As String, _
    ByVal vSecQty As Double) As Decimal
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do_itemTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJmlPending5(vPrimaryCodeNo, vBatchNo, vSecQty).HasValue = True Then
                Return DTA.ScalarQuery_GetJmlPending5(vPrimaryCodeNo, vBatchNo, vSecQty).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlPendingDO3_Multi(ByVal vPrimaryCodeNo As Double, ByVal vBatchNo As String, _
    ByVal vSecQty As Double) As Decimal
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do_itemTableAdapter

        Try
            If DTA.ScalarQuery_GetJmlPending5_Multi(vPrimaryCodeNo, vBatchNo, vSecQty).HasValue = True Then
                Return DTA.ScalarQuery_GetJmlPending5_Multi(vPrimaryCodeNo, vBatchNo, vSecQty).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlPendingDO(ByVal vPrimaryCodeNo As Double) As Decimal
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do_itemTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPendingDO(vPrimaryCodeNo).HasValue = True Then
                Return DTA.ScalarQuery_GetPendingDO(vPrimaryCodeNo).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlPendingDO_Multi(ByVal vPrimaryCodeNo As Double) As Decimal
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do_itemTableAdapter

        Try
            If DTA.ScalarQuery_GetPendingDO(vPrimaryCodeNo).HasValue = True Then
                Return DTA.ScalarQuery_GetPendingDO_Multi(vPrimaryCodeNo).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlPendingDO_Multi(ByVal vPrimaryCodeNo As Double, ByVal vPrimarySO As Double) As Decimal
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do_itemTableAdapter

        Try
            If DTA.ScalarQuery_GetJmlPendingDO2_Multi(vPrimaryCodeNo, vPrimarySO).HasValue = True Then
                Return DTA.ScalarQuery_GetJmlPendingDO2_Multi(vPrimaryCodeNo, vPrimarySO)
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml2(ByVal vPrimaryCodeNo As Double, ByVal vPrimaryDO As Double, ByVal vSecQty As Double, _
  ByVal vBatch As String) As Decimal
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do_itemTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_GetJml3(vPrimaryCodeNo, vPrimaryDO, _
            vBatch, vSecQty).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml3(ByVal vPrimaryDO As Double, ByVal vPrimaryProduct As Double, _
   ByVal vSecondaryQty As Double, ByVal vBatchNo As String) As Decimal
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do_itemTableAdapter
        Try
            Return DTA.ScalarQuery_GetJml2(vPrimaryProduct, vPrimaryDO, _
            vBatchNo, vSecondaryQty).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' mengembalikan jumlah pending do selain pending do yang diberikan primarynya 
    ''' </summary>
    ''' <param name="vPrimaryCodeNo"></param>
    ''' <param name="vBatchNo"></param>
    ''' <param name="vSecQty"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetJmlPendingDO_Other(ByVal vPrimaryCodeNo As Double, ByVal vBatchNo As String, _
    ByVal vSecQty As Double, ByVal vPrimaryDO As Double) As Decimal
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do_itemTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJmlPending4(vPrimaryCodeNo, vBatchNo, vSecQty, vPrimaryDO).HasValue = True Then
                Return DTA.ScalarQuery_GetJmlPending4(vPrimaryCodeNo, vBatchNo, vSecQty, vPrimaryDO).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlPendingDO_Other(ByVal vPrimaryCodeNo As Double, ByVal vPrimarySO As Double, ByVal vPrimaryDOOther As Double) As Decimal
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do_itemTableAdapter
        Dim Dummy As Nullable(Of Decimal)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetJmlDOPending3_Default(vPrimaryCodeNo, vPrimarySO, vPrimaryDOOther)
            If Dummy.HasValue = True Then
                Return Dummy.Value
            Else
                Return 0
            End If
            If DTA.ScalarQuery_GetJmlPendingDO2_Default(vPrimaryCodeNo, vPrimarySO).HasValue = True Then
                Return DTA.ScalarQuery_GetJmlPendingDO2_Default(vPrimaryCodeNo, vPrimarySO)
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml_Multi(ByVal vPrimaryDO As Double, ByVal vPrimarySO As Double, ByVal vPrimaryProduct As Double) As Decimal
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do_itemTableAdapter
        Try
            Return DTA.ScalarQuery_GetJml_Multi(vPrimaryDO, vPrimaryProduct, vPrimarySO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml4_Multi(ByVal vPrimaryDO As Double, ByVal vPrimarySO As Double, _
    ByVal vPrimaryProduct As Double, ByVal vBatchNo As String) As Decimal
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do_itemTableAdapter
        Try
            Return DTA.ScalarQuery_GetJml4_Multi(vPrimaryDO, vPrimaryProduct, vPrimarySO, vBatchNo)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
