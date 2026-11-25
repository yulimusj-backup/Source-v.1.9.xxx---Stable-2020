Option Explicit On
Option Strict Off

Public Class ClsTransaksiPackingListItem
    Public Function IsiDgvCheckStock(ByVal vDGV As DataGridView, _
    ByVal TBL As DseTransaksiPackingList.DataTable22DataTable, _
    Optional ByVal vPrimaryProduct As Double = -1) As Boolean
        Dim DTA As New DseTransaksiPackingListTableAdapters.DataTable22TableAdapter
        Try
            If vPrimaryProduct = -1 Then
                DTA.Fill(TBL)
            Else
                DTA.FillBy_PrimaryProduct(TBL, vPrimaryProduct)
            End If
            vDGV.DataSource = TBL.DefaultView
            vDGV.Columns(1).DefaultCellStyle.Format = "dd MMM yyyy"
            vDGV.Columns(4).DefaultCellStyle.Format = "dd MMM yyyy"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' fungsi ini digunakan untuk cek stock. Mengembalikan total barang yang primarynya diberikan
    ''' Untuk GIT yang belum dimasukkan ke dalam stock in
    ''' </summary>
    ''' <param name="vPrimaryCodeNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetJmlAktif(ByVal vPrimaryCodeNo As Double) As Double
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglist_itemTableAdapter
        Try
            If DTA.ScalarQuery_GetJmlAktif(vPrimaryCodeNo).HasValue = True Then
                Return DTA.ScalarQuery_GetJmlAktif(vPrimaryCodeNo).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub TransferPackingList_Default(ByVal i As FrmTransaksiStockInPurchase2, ByVal vPrimaryPackingList As Double)
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglist_item_view4TableAdapter
        Dim TBL As New DseTransaksiPackingList.tbltrans_packinglist_item_view4DataTable
        Dim row As DseTransaksiPackingList.tbltrans_packinglist_item_view4Row
        Dim DGV As DataGridView = i.dgvItem

        Try
            'kosongkan baris dgv
            DGV.Rows.Clear()

            'isikan data dari basis data
            DTA.FillBy_PrimaryPackingList(TBL, vPrimaryPackingList)

            'salin ke dalam dgv
            For Each row In TBL.Rows
                'yang dipindahkan hanya item - item yang tidak: batch controlled, expiry date controlled, 
                'use secondary unit, dan location controlled
                If row.Batch_No_Controlled = 0 And row.Expiry_Date_Controlled = _
                False And row.Location_Controlled = 0 And row.Use_Secondary_Unit = 0 Then
                    i.dgvItem.Rows.Add("", row.PO_Number, "", row.Product_Code, row.Description, i.ObjCD.DigitGrouping(row.Total_Qty, row._Decimal), _
                    row.Unit_Code, -1, "", "", "-", "", "", "", row.Primary_PO, row.Primary_Product, -1, row.Primary_Units, _
                    -1, -1)
                End If
            Next

            'set number & format dgv
            i.FormatDGV()
            i.ObjCD.SetNumber(i.dgvItem)

            'txt packing list dan btn find packing list tidak bisa digunakan
            i.btnFindPackingList.Enabled = False
            i.TxtPackingListNo.ReadOnly = True
            i.TxtPackingListNo.BackColor = i.txtReadOnly.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    ''@rosadinnn
    ''@MUaT
    Public Sub TransferPackingList_Multi(ByVal i As FrmTransaksiStockInPurchase2, ByVal vPrimaryPackingList As Double)
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglist_item_view4_MultiTableAdapter
        Dim TBL As New DseTransaksiPackingList.tbltrans_packinglist_item_view4_MultiDataTable
        Dim row As DseTransaksiPackingList.tbltrans_packinglist_item_view4_MultiRow
        Dim DGV As DataGridView = i.dgvItem

        Dim RowDgv As DataGridViewRow
        Dim ObjProduct As New Core_Module.clsMasterProduct

        Try
            'kosongkan baris dgv
            DGV.Rows.Clear()

            'isikan data dari basis data
            DTA.FillBy_PrimaryPackingList(TBL, vPrimaryPackingList)

            'salin ke dalam dgv
            For Each row In TBL.Rows
                'yang dipindahkan hanya item - item yang tidak: batch controlled, expiry date controlled, 
                'use secondary unit, dan location controlled
                If row.Batch_No_Controlled = 0 And row.Expiry_Date_Controlled = _
                        False And row.Location_Controlled = 0 And row.Use_Secondary_Unit = 0 Then
                    'i.dgvItem.Rows.Add("", row.PO_Number, "", row.Product_Code, row.Description, i.ObjCD.DigitGrouping(row.Total_Qty, row._Decimal), _
                    'row.Unit_Code, -1, "", "", "-", "", "", "", row.Primary_PO, row.Primary_Product, -1, row.Primary_Units, _
                    '-1, -1)

                    i.dgvItem.Rows.Add()
                    RowDgv = i.dgvItem.Rows(i.dgvItem.Rows.Count - 1)
                    With RowDgv
                        .Cells(i.ColNo.Index).Value = ""
                        .Cells(i.ColPONo.Index).Value = row.PO_Number
                        .Cells(i.ColSuppDONo.Index).Value = ""
                        .Cells(i.ColCodeNo.Index).Value = row.Product_Code
                        .Cells(i.ColDescription.Index).Value = row.Description
                        .Cells(i.ColQtyMulti.Index).Value = i.ObjCD.DigitGrouping(row.Total_Qty_Multi, ObjProduct.GetAngkaDecimal(row.Primary_Product, row.Primary_Unit_Multi))
                        .Cells(i.ColUnitCodeMulti.Index).Value = row.Unit_Code_Multi
                        If row.Primary_Units.ToString <> row.Primary_Unit_Multi.ToString Then
                            .Cells(i.ColRatioDesc.Index).Value = i.ObjCD.DigitGrouping(row.Ratio_Reference, ObjProduct.GetAngkaDecimal(row.Primary_Product, row.Primary_Unit_Multi)) _
                            & " " & row.Unit_Code_Multi & " = " & i.ObjCD.DigitGrouping(row.Ratio_Default, _
                            ObjProduct.GetAngkaDecimal(row.Primary_Product)) & " " & row.Unit_Code
                        Else
                            .Cells(i.ColRatioDesc.Index).Value = "-"
                        End If
                        .Cells(i.ColQty.Index).Value = i.ObjCD.DigitGrouping(row.Total_Qty, ObjProduct.GetAngkaDecimal(row.Primary_Product))
                        .Cells(i.ColUnit.Index).Value = row.Unit_Code
                        .Cells(i.ColSecondaryQty.Index).Value = -1
                        .Cells(i.ColSecondaryUnit.Index).Value = ""
                        .Cells(i.ColBatch.Index).Value = ""
                        .Cells(i.ColExpiredDate.Index).Value = "-"
                        .Cells(i.ColLocation.Index).Value = ""
                        .Cells(i.ColWareHouse.Index).Value = ""
                        .Cells(i.ColRemark.Index).Value = ""
                        .Cells(i.ColPrimaryPOno.Index).Value = row.Primary_PO
                        .Cells(i.ColPrimaryCodeNo.Index).Value = row.Primary_Product
                        .Cells(i.ColPrimaryLocation.Index).Value = -1
                        .Cells(i.ColPrimaryUnits.Index).Value = row.Primary_Units
                        .Cells(i.ColPrimarySecondaryUnits.Index).Value = -1
                        .Cells(i.ColPrimaryPdO.Index).Value = -1
                        .Cells(i.ColPrimaryUnitMulti.Index).Value = row.Primary_Unit_Multi
                        .Cells(i.ColRatioDefault.Index).Value = row.Ratio_Default
                        .Cells(i.ColRatioReference.Index).Value = row.Ratio_Reference
                    End With
                End If
            Next

            'set number & format dgv
            i.FormatDGV()
            i.ObjCD.SetNumber(i.dgvItem)

            'txt packing list dan btn find packing list tidak bisa digunakan
            i.btnFindPackingList.Enabled = False
            i.TxtPackingListNo.ReadOnly = True
            i.TxtPackingListNo.BackColor = i.txtReadOnly.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub


    Public Function GetJml_Default(ByVal vPrimaryPackingList As Double, ByVal vPrimaryPO As Double, _
    ByVal vPrimaryProduct As Double) As Double
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglist_itemTableAdapter

        Try
            Return DTA.ScalarQuery_GetJml(vPrimaryPackingList, vPrimaryProduct, vPrimaryPO).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    ''@rosadinnn
    ''@MUaT
    Public Function GetJml_Multi(ByVal vPrimaryPackingList As Double, ByVal vPrimaryPO As Double, _
            ByVal vPrimaryProduct As Double) As Double
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglist_itemTableAdapter

        Try
            Return DTA.ScalarQuery_GetJml_Multi(vPrimaryPackingList, vPrimaryProduct, vPrimaryPO).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDGV2_Default(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglist_item_view1TableAdapter
        Dim TBL As New DseTransaksiPackingList.tbltrans_packinglist_item_view1DataTable
        Dim row As DseTransaksiPackingList.tbltrans_packinglist_item_view1Row

        Dim ObjCD As New Core_Module.ClsCekData

        Try
            'kosongkan dgv
            vDGVItem.Rows.Clear()
            'isikan tabel
            If DTA.Fill(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    With row
                        vDGVItem.Rows.Add("0", row.PO_Number, row.Product_Code, row.Description, ._Supplier_s_Code_No, ._Supplier_s_Desc, _
                        row.Qty_Master, row.Qty_Per_Master, row.Master_Unit, row.Qty, ObjCD.DigitGrouping(.Total_Qty, ._Decimal), .Unit_Code, row.Remarks, _
                        .Primary_Product, row.Primary_Units, row.Primary_Master_Units, .Primary_PO)
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
    Public Function IsiDGV2_Multi(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglist_item_view1_MultiTableAdapter
        Dim TBL As New DseTransaksiPackingList.tbltrans_packinglist_item_view1_MultiDataTable
        Dim row As DseTransaksiPackingList.tbltrans_packinglist_item_view1_MultiRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim RowDgv As DataGridViewRow
        Dim FrmRef As New FrmTransaksiStockInPurchase2
        Dim ObjProduct As New Core_Module.clsMasterProduct

        Try
            'kosongkan dgv
            vDGVItem.Rows.Clear()
            'isikan tabel
            If DTA.Fill(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    vDGVItem.Rows.Add()
                    RowDgv = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                    With RowDgv
                        .Cells(FrmRef.colNoDGVPackingList.Index).Value = ""
                        .Cells(FrmRef.colPONoDGVPackingList.Index).Value = row.PO_Number
                        .Cells(FrmRef.colCodeNoDGVPackingList.Index).Value = row.Product_Code
                        .Cells(FrmRef.colDescDGVPackingLIst.Index).Value = row.Description
                        .Cells(FrmRef.ColSuppCodeNoDGVPackingLIst.Index).Value = row._Supplier_s_Code_No
                        .Cells(FrmRef.ColSuppDescDGVPackingList.Index).Value = row._Supplier_s_Desc
                        .Cells(FrmRef.ColQtyMulti_PL.Index).Value = ObjCD.DigitGrouping(row.Total_Qty_Multi, ObjProduct.GetAngkaDecimal(row.Primary_Product, row.Primary_Unit_Multi))
                        .Cells(FrmRef.ColUnitMulti_PL.Index).Value = row.Unit_Code_Multi
                        .Cells(FrmRef.ColRatio_PL.Index).Value = ObjCD.DigitGrouping(row.Ratio_Reference, ObjProduct.GetAngkaDecimal(row.Primary_Product, row.Primary_Unit_Multi)) & " " & row.Unit_Code_Multi & " = " & ObjCD.DigitGrouping(row.Ratio_Default, ObjProduct.GetAngkaDecimal(row.Primary_Product)) & " " & row.Unit_Code
                        .Cells(FrmRef.ColTotalQtyDGVPackingList.Index).Value = ObjCD.DigitGrouping(row.Total_Qty, ObjProduct.GetAngkaDecimal(row.Primary_Product))
                        .Cells(FrmRef.colUnitsDGVPackingList.Index).Value = row.Unit_Code
                        .Cells(FrmRef.colRemarksDGVPackingList.Index).Value = row.Remarks
                        .Cells(FrmRef.colPrimaryCodeNoDGVPAckingList.Index).Value = row.Primary_Product
                        .Cells(FrmRef.colPrimaryUnitsDGVPackingLIst.Index).Value = row.Primary_Units
                        .Cells(FrmRef.ColPrimaryMasterUnitsDGVPAckingLIst.Index).Value = row.Primary_Master_Units
                        .Cells(FrmRef.ColPrimaryPODGVPackingLIst.Index).Value = row.Primary_PO
                        .Cells(FrmRef.colQtyDGVPackingList.Index).Value = row.Qty
                        .Cells(FrmRef.ColQtyMasterDGVPackingList.Index).Value = row.Qty_Per_Master
                        .Cells(FrmRef.ColMasterUnitDGVPackingList.Index).Value = row.Master_Unit
                        .Cells(FrmRef.ColPrimaryUnitsMulti_PL.Index).Value = row.Primary_Unit_Multi
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

    Public Sub IsiTBLPO(ByVal vPrimaryPackingList As Double, ByVal vTBL As _
    DseTransaksiPackingList.tbltrans_packinglist_item_view3DataTable)
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglist_item_view3TableAdapter
        Try
            DTA.FillBy_PrimaryPackingList(vTBL, vPrimaryPackingList)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsiTBLProductByPO(ByVal vPrimaryPackingList As Double, ByVal vTBL As _
    DseTransaksiPackingList.tbltrans_packinglist_item_view3DataTable, ByVal vPrimaryPO As Double)
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglist_item_view3TableAdapter
        Try
            DTA.FillBy_PrimaryPO(vTBL, vPrimaryPackingList, vPrimaryPO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub
End Class