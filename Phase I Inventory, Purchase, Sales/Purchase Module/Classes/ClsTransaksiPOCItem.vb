Option Explicit On
Option Strict Off

Public Class ClsTransaksiPOCItem
    Public Function IsiDGV_Default(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTransaksiPOCTableAdapters.tbltrans_poc_view1TableAdapter
        Dim TBL As New DseTransaksiPOC.tbltrans_poc_view1DataTable
        Dim row As DseTransaksiPOC.tbltrans_poc_view1Row

        Dim ObjCD As New Core_Module.ClsCekData

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'kosongkan dgv
            vDGVItem.Rows.Clear()
            'isikan tabel
            If DTA.FillBy_PrimaryTransaksi(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    With row
                        vDGVItem.Rows.Add("0", row.Product_Code, row.Description, .Qty_Master, .Qty_Per_Master, _
                         row.Master_Unit, row.Qty, ObjCD.DigitGrouping(.Total_Qty, ._Decimal), .Units, row.Remarks, _
                        .Primary_Product, row.Primary_Master_Units, .Primary_Units)
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

    ''@rosadinnn
    ''@MUaT
    Public Function IsiDGV_Multi(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTransaksiPOCTableAdapters.tbltrans_poc_view1_MultiTableAdapter
        Dim TBL As New DseTransaksiPOC.tbltrans_poc_view1_MultiDataTable
        Dim row As DseTransaksiPOC.tbltrans_poc_view1_MultiRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim RowDgv As DataGridViewRow
        Dim i As New FrmTransaksiPOC
        Dim ObjProd As New Core_Module.clsMasterProduct

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            vDGVItem.Rows.Clear()
            If DTA.FillBy_PrimaryPOC(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    With row
                        vDGVItem.Rows.Add()
                        RowDgv = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                        RowDgv.Cells(i.ColNo.Index).Value = ""
                        RowDgv.Cells(i.ColCodeNo.Index).Value = .Product_Code
                        RowDgv.Cells(i.ColDescription.Index).Value = .Description
                        RowDgv.Cells(i.ColMasterUnit.Index).Value = .Master_Unit
                        RowDgv.Cells(i.ColQtyMulti.Index).Value = ObjCD.DigitGrouping(.Total_Qty_Multi, ObjProd.GetAngkaDecimal(.Primary_Product, .Primary_Unit_Multi))
                        RowDgv.Cells(i.ColUnitCodeMulti.Index).Value = .Unit_Code_Multi
                        If .Primary_Units <> .Primary_Unit_Multi Then
                            RowDgv.Cells(i.ColRatioDesc.Index).Value = ObjCD.DigitGrouping(.Ratio_Reference, _
                                    ObjProd.GetAngkaDecimal(.Primary_Product, .Primary_Unit_Multi)) _
                                    & " " & .Unit_Code_Multi & " = " & ObjCD.DigitGrouping(.Ratio_Default, _
                                    ObjProd.GetAngkaDecimal(.Primary_Product)) & " " & .Units
                        Else
                            RowDgv.Cells(i.ColRatioDesc.Index).Value = "-"
                        End If
                        RowDgv.Cells(i.ColTotalQty.Index).Value = ObjCD.DigitGrouping(.Total_Qty, ObjProd.GetAngkaDecimal(.Primary_Product))
                        RowDgv.Cells(i.ColUnit.Index).Value = .Units
                        RowDgv.Cells(i.ColRemark.Index).Value = .Remarks
                        RowDgv.Cells(i.ColRatioReference.Index).Value = .Ratio_Reference
                        RowDgv.Cells(i.ColRatioDefault.Index).Value = .Ratio_Default
                        RowDgv.Cells(i.ColPrimaryCodeNo.Index).Value = .Primary_Product
                        RowDgv.Cells(i.ColPrimaryMasterUnit.Index).Value = .Primary_Master_Units
                        RowDgv.Cells(i.ColPrimaryUnits.Index).Value = .Primary_Units
                        RowDgv.Cells(i.ColPrimaryUnitMulti.Index).Value = .Primary_Unit_Multi
                        RowDgv.Cells(i.ColQtyMaster.Index).Value = .Qty_Master
                        RowDgv.Cells(i.ColQtyPerMaster.Index).Value = .Qty_Per_Master
                        RowDgv.Cells(i.ColQty.Index).Value = .Qty
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

    Public Function Simpan_Default(ByVal vdgvItem As DataGridView, ByVal vPrimaryTransaksi As Double) As Boolean
        Dim row As DataGridViewRow
        Dim DTA As New DseTransaksiPOCTableAdapters.tbltrans_poc_itemTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            For Each row In vdgvItem.Rows
                'simpan ke basis data item
                DTA.Insert(vPrimaryTransaksi, CDbl(row.Cells(10).Value), CDbl(row.Cells(3).Value), CDbl(row.Cells(11).Value), _
                CDbl(row.Cells(6).Value), CDbl(row.Cells(7).Value), CDbl(row.Cells(12).Value), _
                row.Cells(9).Value, CDbl(row.Cells(4).Value))
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    ''@rosadinnn
    ''@MUaT
    Public Function Simpan_Multi(ByVal vdgvItem As DataGridView, ByVal vPrimaryTransaksi As Double) As Boolean
        Dim row As DataGridViewRow
        Dim DTA As New DseTransaksiPOCTableAdapters.tbltrans_poc_itemTableAdapter
        Dim i As New FrmTransaksiPOC

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            For Each row In vdgvItem.Rows
                With row
                    DTA.Insert_Multi(vPrimaryTransaksi, .Cells(i.ColPrimaryCodeNo.Index).Value, .Cells(i.ColQtyMaster.Index).Value, .Cells(i.ColPrimaryMasterUnit.Index).Value, _
                    .Cells(i.ColQty.Index).Value, .Cells(i.ColTotalQty.Index).Value, .Cells(i.ColPrimaryUnits.Index).Value, .Cells(i.ColRemark.Index).Value, _
                    .Cells(i.ColQtyPerMaster.Index).Value, .Cells(i.ColQtyMulti.Index).Value, .Cells(i.ColPrimaryUnitMulti.Index).Value, .Cells(i.ColRatioReference.Index).Value, _
                    .Cells(i.ColRatioDefault.Index).Value)
                End With
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Hapus(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTransaksiPOCTableAdapters.tbltrans_poc_itemTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

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

    Public Function GetJml_Default(ByVal vPrimaryPOC As Double, _
            ByVal vPrimaryCodeNo As Double, ByVal vPrimaryPO As Double) As Double
        Dim DTA As New DseTransaksiPOCTableAdapters.tbltrans_poc_itemTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            Return DTA.ScalarQuery_GetJml(vPrimaryPOC, vPrimaryPO, vPrimaryCodeNo).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    ''@rosadinnn
    ''@MUaT
    Public Function GetJml_Multi(ByVal vPrimaryPOC As Double, _
            ByVal vPrimaryCodeNo As Double, ByVal vPrimaryPO As Double) As Double
        Dim DTA As New DseTransaksiPOCTableAdapters.tbltrans_poc_itemTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            Return DTA.ScalarQuery_GetJml_Multi(vPrimaryPOC, vPrimaryPO, vPrimaryCodeNo).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub IsiTBLPO(ByVal vPrimaryPOC As Double, ByVal vTBL As _
    DseTransaksiPOC.tbltrans_poc_itemview2DataTable)
        Dim DTA As New DseTransaksiPOCTableAdapters.tbltrans_poc_itemview2TableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            DTA.FillBy_PrimaryPOC(vTBL, vPrimaryPOC)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsiTBLProductByPO_Default(ByVal vPrimaryPOC As Double, ByVal vTBL As _
    DseTransaksiPOC.tbltrans_poc_itemview2DataTable, ByVal vPrimaryPO As Double)
        Dim DTA As New DseTransaksiPOCTableAdapters.tbltrans_poc_itemview2TableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            DTA.FillBy_PrimaryPOCnPO(vTBL, vPrimaryPOC, vPrimaryPO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    ''@rosadinnn
    ''@MUaT
    Public Sub IsiTBLProductByPO_Multi(ByVal vPrimaryPOC As Double, ByVal vTBL As _
            DseTransaksiPOC.tbltrans_poc_itemview2_MultiDataTable, ByVal vPrimaryPO As Double)
        Dim DTA As New DseTransaksiPOCTableAdapters.tbltrans_poc_itemview2_MultiTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            DTA.FillBy_PrimaryPOCnPO(vTBL, vPrimaryPOC, vPrimaryPO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function GetJmlPOCAktif_Default(ByVal vPrimaryPO As Double, _
    ByVal vPrimaryProduct As Double, ByVal vPrimaryPOC As Double) As Double
        Dim DTA As New DseTransaksiPOCTableAdapters.tbltrans_poc_itemview2TableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            ' If vPrimaryPOC = -1 Then
            'Return DTA.ScalarQuery_GetJmlPrimaryPOnProduct(vPrimaryPO, vPrimaryProduct)
            'Else
            Return DTA.ScalarQuery_GetJmlByPrimaryPO_POC_Product(vPrimaryPOC, vPrimaryPO, vPrimaryProduct).Value
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    ''@rosadinnn
    ''@MUaT
    Public Function GetJmlPOCAktif_Multi(ByVal vPrimaryPO As Double, _
    ByVal vPrimaryProduct As Double, ByVal vPrimaryPOC As Double) As Double
        Dim DTA As New DseTransaksiPOCTableAdapters.tbltrans_poc_itemview2_MultiTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            ' If vPrimaryPOC = -1 Then
            'Return DTA.ScalarQuery_GetJmlPrimaryPOnProduct(vPrimaryPO, vPrimaryProduct)
            'Else
            Return DTA.ScalarQuery_GetJmlByPrimaryPO_POC_Product(vPrimaryPO, vPrimaryPOC, vPrimaryProduct).Value
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
