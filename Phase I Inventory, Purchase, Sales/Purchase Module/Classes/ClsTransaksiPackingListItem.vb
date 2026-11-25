Option Explicit On
Option Strict Off

Public Class ClsTransaksiPackingListItem
    Public Function GetJml_Default(ByVal vPrimaryPackingList As Double, _
    ByVal vPrimaryCodeNo As Double, ByVal vPrimaryPO As Double) As Double
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglist_itemTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_GetJml(vPrimaryPackingList, vPrimaryCodeNo, vPrimaryPO).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    ''@rosadinnn
    ''@MUaT
    Public Function GetJml_Multi(ByVal vPrimaryPackingList As Double, _
    ByVal vPrimaryCodeNo As Double, ByVal vPrimaryPO As Double) As Double
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglist_itemTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            Return DTA.ScalarQuery_GetJmlMulti(vPrimaryPackingList, vPrimaryCodeNo, vPrimaryPO).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub IsiTBLPO_Default(ByVal vPrimaryPackingList As Double, ByVal vTBL As _
    DseTransaksiPackingList.tbltrans_packinglist_item_view3DataTable)
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglist_item_view3TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            DTA.FillBy_PrimaryPackingList(vTBL, vPrimaryPackingList)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    ''@rosadinnn
    ''@MUaT
    Public Sub IsiTBLPO_Multi(ByVal vPrimaryPackingList As Double, ByVal vTBL As _
            DseTransaksiPackingList.tbltrans_packinglist_item_view3MultiDataTable)
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglist_item_view3MultiTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            DTA.FillBy_PrimaryPackingList(vTBL, vPrimaryPackingList)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsiTBLProductByPO_Default(ByVal vPrimaryPackingList As Double, ByVal vTBL As _
    DseTransaksiPackingList.tbltrans_packinglist_item_view3DataTable, ByVal vPrimaryPO As Double)
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglist_item_view3TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            DTA.FillBy_PrimaryPO(vTBL, vPrimaryPackingList, vPrimaryPO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsiTBLProductByPO_Multi(ByVal vPrimaryPackingList As Double, ByVal vTBL As _
    DseTransaksiPackingList.tbltrans_packinglist_item_view3MultiDataTable, ByVal vPrimaryPO As Double)
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglist_item_view3MultiTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            DTA.FillBy_PrimaryPO(vTBL, vPrimaryPackingList, vPrimaryPO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function GetJmlPackingListAktif_Default(ByVal vPrimaryPO As Double, _
    ByVal vPrimaryProduct As Double, ByVal vPrimaryPackingList As Double) As Double
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglist_item_view2TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_GetJmlGITAktif(vPrimaryPO, vPrimaryProduct, vPrimaryPackingList).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    ''@rosadinnn
    ''@MUaT
    Public Function GetJmlPackingListAktif_Multi(ByVal vPrimaryPO As Double, _
    ByVal vPrimaryProduct As Double, ByVal vPrimaryPackingList As Double) As Double
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglist_item_view1_MultiTableAdapter
        Dim NoD As New Nullable(Of Decimal)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            NoD = DTA.ScalarQuery_GetJmlGITAktif_Multi(vPrimaryPO, vPrimaryProduct, vPrimaryPackingList)
            If NoD.HasValue = True Then
                Return NoD.Value
            End If
            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDGV_Default(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglist_item_view1TableAdapter
        Dim TBL As New DseTransaksiPackingList.tbltrans_packinglist_item_view1DataTable
        Dim row As DseTransaksiPackingList.tbltrans_packinglist_item_view1Row

        Dim ObjCD As New Core_Module.ClsCekData

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'kosongkan dgv
            vDGVItem.Rows.Clear()
            'isikan tabel
            If DTA.Fill(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    With row
                        vDGVItem.Rows.Add("0", row.PO_Number, row.Product_Code, row.Description, ._Supplier_s_Code_No, ._Supplier_s_Desc, _
                        row.Qty_Master, row.Qty_Per_Master, row.Master_Unit, ObjCD.SetDecimal(row.Qty, ._Decimal), _
                        ObjCD.SetDecimal(.Total_Qty, ._Decimal), .Unit_Code, row.Remarks, .Primary_Product, _
                        row.Primary_Units, row.Primary_Master_Units, .Primary_PO)
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
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglist_item_view1_MultiTableAdapter
        Dim TBL As New DseTransaksiPackingList.tbltrans_packinglist_item_view1_MultiDataTable
        Dim row As DseTransaksiPackingList.tbltrans_packinglist_item_view1_MultiRow

        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjProduct As New Core_Module.clsMasterProduct
        Dim RowDgv As DataGridViewRow
        Dim FrmRef As New FrmTransaksiPackingList

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'kosongkan dgv
            vDGVItem.Rows.Clear()
            'isikan tabel
            If DTA.Fill(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    With row

                        vDGVItem.Rows.Add()
                        RowDgv = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                        RowDgv.Cells(FrmRef.ColNo.Index).Value = ""
                        RowDgv.Cells(FrmRef.ColPONo.Index).Value = .PO_Number
                        RowDgv.Cells(FrmRef.ColCodeNo.Index).Value = .Product_Code
                        RowDgv.Cells(FrmRef.ColDescription.Index).Value = .Description
                        RowDgv.Cells(FrmRef.ColSuppCodeNo.Index).Value = ""
                        RowDgv.Cells(FrmRef.ColSuppDesc.Index).Value = ""
                        RowDgv.Cells(FrmRef.ColQtyMaster.Index).Value = "0"
                        RowDgv.Cells(FrmRef.ColQtyPerMaster.Index).Value = "0"
                        RowDgv.Cells(FrmRef.ColMasterUnit.Index).Value = ""
                        RowDgv.Cells(FrmRef.ColQty.Index).Value = "0"
                        RowDgv.Cells(FrmRef.ColTotalQty.Index).Value = ObjCD.DigitGrouping(.Total_Qty, ObjProduct.GetAngkaDecimal(.Primary_Product))
                        RowDgv.Cells(FrmRef.ColUnit.Index).Value = .Unit_Code
                        RowDgv.Cells(FrmRef.ColRemark.Index).Value = .Remarks
                        RowDgv.Cells(FrmRef.ColPrimaryCodeNo.Index).Value = .Primary_Product
                        RowDgv.Cells(FrmRef.ColPrimaryUnits.Index).Value = .Primary_Units
                        RowDgv.Cells(FrmRef.ColPrimaryMasterUnits.Index).Value = "-1"
                        RowDgv.Cells(FrmRef.ColPrimaryPO.Index).Value = .Primary_PO
                        RowDgv.Cells(FrmRef.ColQtyMulti.Index).Value = ObjCD.DigitGrouping(.Total_Qty_Multi, ObjProduct.GetAngkaDecimal(.Primary_Product, .Primary_Unit_Multi))
                        RowDgv.Cells(FrmRef.ColUnitCodeMulti.Index).Value = .Unit_Code_Multi
                        If .Primary_Units <> .Primary_Unit_Multi Then
                            RowDgv.Cells(FrmRef.ColRatioDesc.Index).Value = ObjCD.DigitGrouping(.Ratio_Reference, _
                                    ObjProduct.GetAngkaDecimal(.Primary_Product, .Primary_Unit_Multi)) _
                                    & " " & .Unit_Code_Multi & " = " & ObjCD.DigitGrouping(.Ratio_Default, _
                                    ObjProduct.GetAngkaDecimal(.Primary_Product)) & " " & .Unit_Code
                        Else
                            RowDgv.Cells(FrmRef.ColRatioDesc.Index).Value = "-"
                        End If
                        RowDgv.Cells(FrmRef.ColRatioReference.Index).Value = ObjCD.DigitGrouping(.Ratio_Reference, ObjProduct.GetAngkaDecimal(.Primary_Product, .Primary_Unit_Multi))
                        RowDgv.Cells(FrmRef.ColRatioDefault.Index).Value = ObjCD.DigitGrouping(.Ratio_Default, ObjProduct.GetAngkaDecimal(.Primary_Product))
                        RowDgv.Cells(FrmRef.ColPrimaryUnitMulti.Index).Value = .Primary_Unit_Multi

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

    Public Function Simpan_Default(ByVal vdgvItem As DataGridView, ByVal vPrimaryTransaksi As Double) As Boolean
        Dim row As DataGridViewRow
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglist_itemTableAdapter
        Dim ObjPOItem As New ClsTransaksiPOItem
        Dim vListPrimary As New List(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            For Each row In vdgvItem.Rows
                'simpan ke basis data item
                DTA.Insert(vPrimaryTransaksi, row.Cells(13).Value, row.Cells(16).Value, row.Cells(6).Value, _
                row.Cells(7).Value, row.Cells(15).Value, row.Cells(9).Value, _
                row.Cells(10).Value, row.Cells(14).Value, row.Cells(12).Value)
                If vListPrimary.Contains(row.Cells(16).Value) = False Then
                    vListPrimary.Add(row.Cells(16).Value)
                End If
            Next
            'set order balance
            ObjPOItem.SetOrderBalance(vListPrimary)
            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjPOItem = Nothing
            vListPrimary = Nothing
        End Try
    End Function

    ''@rosadinnn
    ''@MUaT
    Public Function Simpan_Multi(ByVal vdgvItem As DataGridView, ByVal vPrimaryTransaksi As Double, _
    Optional ByVal vSETPOB As Boolean = True) As Boolean
        Dim row As DataGridViewRow
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglist_itemTableAdapter
        Dim ObjPOItem As New ClsTransaksiPOItem
        Dim vListPrimary As New List(Of Double)

        Dim FrmRef As New FrmTransaksiPackingList
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            For Each row In vdgvItem.Rows
                'simpan ke basis data item
                With row
                    DTA.InsertQueryMulti(vPrimaryTransaksi, .Cells(FrmRef.ColPrimaryCodeNo.Index).Value, .Cells(FrmRef.ColPrimaryPO.Index).Value, .Cells(FrmRef.ColQtyMaster.Index).Value, _
                            .Cells(FrmRef.ColQtyPerMaster.Index).Value, .Cells(FrmRef.ColPrimaryMasterUnits.Index).Value, .Cells(FrmRef.ColQty.Index).Value, _
                            .Cells(FrmRef.ColTotalQty.Index).Value, .Cells(FrmRef.ColPrimaryUnits.Index).Value, .Cells(FrmRef.ColRemark.Index).Value, _
                            .Cells(FrmRef.ColQtyMulti.Index).Value, .Cells(FrmRef.ColPrimaryUnitMulti.Index).Value, .Cells(FrmRef.ColRatioDefault.Index).Value, .Cells(FrmRef.ColRatioReference.Index).Value)
                    If vListPrimary.Contains(row.Cells(FrmRef.ColPrimaryPO.Index).Value) = False Then
                        vListPrimary.Add(row.Cells(FrmRef.ColPrimaryPO.Index).Value)
                    End If
                End With
            Next

            If vSETPOB = True Then
                'set order balance
                ObjPOItem.SetOrderBalance(vListPrimary)
            End If

            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjPOItem = Nothing
            vListPrimary = Nothing
        End Try
    End Function

    Public Function Hapus(ByVal vPrimaryTransaksi As Double, ByVal ObjPL As ClsTransaksiPackingList, Optional ByVal vSetOB As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiPackingListTableAdapters.tbltrans_packinglist_itemTableAdapter
        Dim ObjPOItem As New ClsTransaksiPOItem
        Dim vListPrimaryPO As New List(Of Double)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            ObjPL.GetPrimaryPO(vPrimaryTransaksi, vListPrimaryPO)
            'hapus data item
            DTA.Delete(vPrimaryTransaksi)
            'set order balance
            If vSetOB = True Then
                ObjPOItem.SetOrderBalance(vListPrimaryPO)
            End If

            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjPOItem = Nothing
            vListPrimaryPO = Nothing
        End Try
    End Function

    ''' <summary>
    ''' fungsi ini digunakan untuk cek stock. Mengembalikan total barang yang primarynya diberikan
    ''' Untuk GIT yang belum dimasukkan ke dalam stock in
    ''' </summary>
    ''' <param name="vPrimaryCodeNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetJml(ByVal vPrimaryCodeNo As Double) As Double
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
End Class
