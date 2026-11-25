Option Explicit On
Option Strict Off

Public Class ClsTransaksiGRNItem

    Public Function GetJmlByPO(ByVal vPrimaryPO As Double, ByVal vPrimaryProduct As Double) As Integer
        Dim DTA As New DseTransaksiGRNTableAdapters.tbltrans_grn_itemTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_GetJmlBYPO_n_Product(vPrimaryPO, vPrimaryProduct).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlByPO(ByVal vPrimaryPO As Double) As Integer
        Dim DTA As New DseTransaksiGRNTableAdapters.tbltrans_grn_itemTableAdapter
        Try
            Return DTA.ScalarQuery_GetJmlByPrimaryPO(vPrimaryPO).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Simpan(ByVal vdgvItem As DataGridView, ByVal vPrimaryTransaksi As Double) As Boolean
        Dim row As DataGridViewRow
        Dim DTA As New DseTransaksiGRNTableAdapters.tbltrans_grn_itemTableAdapter
        Dim ObjProd As New Core_Module.clsMasterProduct
        Dim ObjLS As New Core_Module.clsMasterListedSoftware
        Dim DTA1 As New Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter
        Dim IsInventoryInstalled As Boolean = ObjLS.IsInventoryInstalled
        Dim Sec_Qty As Double
        Dim BatchNo As String = ""
        Dim ExpiryDate As Date = #1/1/1950#
        Dim objPO As New ClsTransaksiPO
        Try
            For Each row In vdgvItem.Rows
                If IsNothing(row.Cells("ColBatchNo").Value) = False Then
                    BatchNo = row.Cells("ColBatchNo").Value
                Else
                    BatchNo = ""
                End If

                If IsNumeric(row.Cells("ColSecQty").Value) = True Then
                    Sec_Qty = row.Cells("ColSecQty").Value
                Else
                    Sec_Qty = "-1"
                End If

                If IsDate(row.Cells(row.Cells.Item("ColExpiryDate").Value).Value) = False Then
                    ExpiryDate = #1/1/1950#
                Else
                    ExpiryDate = CDate(row.Cells(row.Cells.Item("ColExpiryDate").Value).Value)
                End If

                'simpan ke basis data item
                DTA.Insert(vPrimaryTransaksi, row.Cells("ColPrimaryCodeNo").Value, row.Cells("colPrimaryPO").Value, BatchNo, _
                ExpiryDate, row.Cells("ColQtyMaster").Value, row.Cells("ColPrimaryMasterUnits").Value, _
                row.Cells("ColQty").Value, _
                row.Cells("ColTotalQty").Value, row.Cells("ColPrimaryUnit").Value, Sec_Qty, _
                row.Cells("ColPrimarySecUnit").Value, row.Cells("DataGridViewTextBoxColumn13").Value, row.Cells("ColQtyPerMaster").Value)
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
            ObjProd = Nothing
            ObjLS = Nothing
            objPO = Nothing
        End Try
    End Function

    Public Function Hapus(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTransaksiGRNTableAdapters.tbltrans_grn_itemTableAdapter
        Try
            'hapus data item
            DTA.Delete(vPrimaryTransaksi)
            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDGV_Default(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTransaksiGRNTableAdapters.TblDGV_ItemTableAdapter
        Dim TBL As New DseTransaksiGRN.TblDGV_ItemDataTable

        Dim row As DseTransaksiGRN.TblDGV_ItemRow
        Dim Sec_Units As String = ""
        Dim ObjCD As New Core_Module.ClsCekData
        Dim objpo As New ClsTransaksiPO

        Try
            'kosongkan dgv
            vDGVItem.Rows.Clear()
            'isikan tabel
            If DTA.FillBy_PrimaryGRN(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    If row.Primary_Secondary_Units <> -1 Then
                        Sec_Units = row.SecUnit
                    Else
                        Sec_Units = ""
                    End If

                    With row
                        vDGVItem.Rows.Add("0", row.Product_Code, row.ProductDesc, row.Supplier_Code, "", row.NoPO, _
                        row.Batch_No, row.Expiry_Date, row.Qty_Master, row.Qty_Per_Master, .MasterUnit, row.Qty, _
                        ObjCD.DigitGrouping(.Total_Qty, ._Decimal), row.unit, row.Secondary_Qty, Sec_Units, row.Remarks, row.Primary_Product, _
                        row.Primary_Master_Units, row.Primary_Units, row.Primary_Secondary_Units, row.Primary_PO)
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
        Dim DTA As New DseTransaksiGRNTableAdapters.TblDGV_Item_MultiTableAdapter
        Dim TBL As New DseTransaksiGRN.TblDGV_Item_MultiDataTable
        Dim row As DseTransaksiGRN.TblDGV_Item_MultiRow
        Dim Sec_Units As String = ""
        Dim ObjCD As New Core_Module.ClsCekData
        Dim objpo As New ClsTransaksiPO
        Dim RowDgv As DataGridViewRow
        Dim i As New FrmTransaksiPurchaseReturn
        Dim objProduct As New Core_Module.clsMasterProduct
        Try
            vDGVItem.Rows.Clear()
            If DTA.Fill(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    If row.Primary_Secondary_Units <> -1 Then
                        Sec_Units = row.SecUnit
                    Else
                        Sec_Units = ""
                    End If

                    'With row
                    '    vDGVItem.Rows.Add("0", row.Product_Code, row.ProductDesc, row.Supplier_Code, "", row.NoPO, _
                    '    row.Batch_No, row.Expiry_Date, row.Qty_Master, row.Qty_Per_Master, .MasterUnit, row.Qty, _
                    '    ObjCD.DigitGrouping(.Total_Qty, ._Decimal), row.unit, row.Secondary_Qty, Sec_Units, row.Remarks, row.Primary_Product, _
                    '    row.Primary_Master_Units, row.Primary_Units, row.Primary_Secondary_Units, row.Primary_PO)
                    'End With

                    vDGVItem.Rows.Add()
                    RowDgv = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                    With RowDgv
                        .Cells(i.ColNo_GRN.Index).Value = ""
                        .Cells(i.colCodeNoDGVGRN.Index).Value = row.Product_Code
                        .Cells(i.colDescDGVGRN.Index).Value = row.ProductDesc
                        .Cells(i.ColSuppCodeNoDGVGRN.Index).Value = row.Supplier_Code
                        .Cells(i.ColSuppItemDescDGVGRN.Index).Value = ""
                        .Cells(i.colPONoDGVGRN.Index).Value = row.NoPO
                        .Cells(i.colBatchNoDGVGRN.Index).Value = row.Batch_No
                        .Cells(i.colExpiryDateDGVGRN.Index).Value = row.Expiry_Date
                        .Cells(i.colQtyMasterDGVGRN.Index).Value = row.Qty_Master
                        .Cells(i.colQtyPerMasterDGVGRN.Index).Value = row.Qty_Per_Master
                        .Cells(i.colMasterUnitDGVGRN.Index).Value = row.MasterUnit
                        .Cells(i.colQTYDGVGRN.Index).Value = row.Qty
                        .Cells(i.colTotalQtyDGVGRN.Index).Value = ObjCD.DigitGrouping(row.Total_Qty, row._Decimal)
                        .Cells(i.colUnitDGVGRN.Index).Value = row.unit
                        .Cells(i.colSecQtyDGVGRN.Index).Value = row.Secondary_Qty
                        .Cells(i.colSecUnitDGVGRN.Index).Value = row.SecUnit
                        .Cells(i.colRemarksDGVGRN.Index).Value = row.Remarks
                        .Cells(i.colPrimaryCodeNoDGVGRN.Index).Value = row.Primary_Product
                        .Cells(i.colPRimaryMasterUnitDGVGRN.Index).Value = row.Primary_Master_Units
                        .Cells(i.colPrimaryUnitDGVGRN.Index).Value = row.Primary_Units
                        .Cells(i.colPrimarySecUnitDGVGRN.Index).Value = row.Primary_Secondary_Units
                        .Cells(i.colPrimaryPODGVGRN.Index).Value = row.Primary_PO
                        .Cells(i.ColQtyMulti_GRN.Index).Value = ObjCD.DigitGrouping(row.Total_Qty_Multi, ObjProduct.GetAngkaDecimal(row.Primary_Product, row.Primary_Unit_Multi))
                        .Cells(i.ColUnitsCodeMulti_GRN.Index).Value = row.Unit_Code_Multi
                        If row.Primary_Units <> row.Primary_Unit_Multi Then
                            .Cells(i.ColRatioDesc_GRN.Index).Value = ObjCD.DigitGrouping(row.Ratio_Reference, _
                                    objProduct.GetAngkaDecimal(row.Primary_Product, row.Primary_Unit_Multi)) _
                                    & " " & row.Unit_Code_Multi & " = " & ObjCD.DigitGrouping(row.Ratio_Default, _
                                    objProduct.GetAngkaDecimal(row.Primary_Product)) & " " & row.unit
                        Else
                            .Cells(i.ColRatioDesc_GRN.Index).Value = "-"
                        End If
                        .Cells(i.ColRatioReference_GRN.Index).Value = row.Ratio_Reference
                        .Cells(i.ColRatioDefault_GRN.Index).Value = row.Ratio_Default
                        .Cells(i.ColPrimaryUnitMulti_GRN.Index).Value = row.Primary_Unit_Multi
                    End With

                Next
            End If
            ObjCD.SetNumber(vDGVItem)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
            ObjCD = Nothing
        End Try
    End Function
End Class
