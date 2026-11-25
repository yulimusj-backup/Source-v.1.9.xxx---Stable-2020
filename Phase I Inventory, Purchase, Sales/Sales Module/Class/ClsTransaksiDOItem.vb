Option Explicit On
Option Strict Off

Public Class ClsTransaksiDOItem
    Public Function GetJmlWithOutLocation(ByVal vPrimaryDO As Double, ByVal vPrimaryProduct As Double) As Double
        Dim DTA As New DseTransaksiDO_AddByLocationTableAdapters.tbltrans_do_itemTableAdapter
        Dim JmlWithoutLocation As Nullable(Of Long)

        Try
            JmlWithoutLocation = DTA.ScalarQuery_GetJmlWithoutLocation(vPrimaryDO, vPrimaryProduct)
            If JmlWithoutLocation.HasValue = False Then
                Return 0
            Else
                Return JmlWithoutLocation.Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetJmlPendingDO_Other_Default(ByVal vPrimaryCodeNo As Double, ByVal vPrimarySO As Double, ByVal vPrimaryDOOther As Double) As Decimal
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do_itemTableAdapter
        Dim Dummy As Nullable(Of Decimal)

        Try
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

    Public Function GetJmlPendingDO_Other_Multi(ByVal vPrimaryCodeNo As Double, ByVal vPrimarySO As Double, ByVal vPrimaryDOOther As Double) As Decimal
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do_itemTableAdapter
        Dim Dummy As Nullable(Of Decimal)

        Try
            Dummy = DTA.ScalarQuery_GetJmlDOPending3_Multi(vPrimaryCodeNo, vPrimarySO, vPrimaryDOOther)
            If Dummy.HasValue = True Then
                Return Dummy.Value
            Else
                Return 0
            End If
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

    Public Function GetJmlPendingDO_Default(ByVal vPrimaryCodeNo As Double, ByVal vPrimarySO As Double) As Decimal
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do_itemTableAdapter

        Try
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

    Public Function GetJmlPendingDO(ByVal vPrimaryCodeNo As Double, ByVal vPrimarySO As Double, ByVal vBatchNo As String) As Decimal
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do_itemTableAdapter

        Try
            If DTA.ScalarQuery_GetJmlPending6(vPrimaryCodeNo, vPrimarySO, vBatchNo).HasValue = True Then
                Return DTA.ScalarQuery_GetJmlPending6(vPrimaryCodeNo, vPrimarySO, vBatchNo).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlPendingDO(ByVal vPrimaryCodeNo As Double, ByVal vPrimarySO As Double, ByVal vBatchNo As String, _
    ByVal vSecQty As Double) As Decimal
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do_itemTableAdapter

        Try
            If DTA.ScalarQuery_GetJmlPending3(vPrimaryCodeNo, vPrimarySO, vBatchNo, vSecQty).HasValue = True Then
                Return DTA.ScalarQuery_GetJmlPending3(vPrimaryCodeNo, vPrimarySO, vBatchNo, vSecQty).Value
            Else
                Return 0
            End If
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

    Public Function GetJmlPendingDO4(ByVal vPrimaryCodeNo As Double, ByVal vBatchNo As String, _
        ByVal vSecQty As Double, ByVal vPrimaryDO As Double) As Decimal
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do_itemTableAdapter

        Try
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

    ''' <summary>
    ''' mengembalikan jumlah pending do
    ''' </summary>
    ''' <param name="vPrimaryCodeNo"></param>
    ''' <param name="vBatchNo"></param>
    ''' <param name="vSecQty"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetJmlPendingDO3_Default(ByVal vPrimaryCodeNo As Double, ByVal vBatchNo As String, _
    ByVal vSecQty As Double) As Decimal
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do_itemTableAdapter

        Try
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

    Public Function GetJmlPendingDO_Default(ByVal vPrimaryCodeNo As Double) As Decimal
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do_itemTableAdapter

        Try
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

    Public Function GetJml_Default(ByVal vPrimaryCodeNo As Double, ByVal vPrimaryDO As Double) As Decimal
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do_itemTableAdapter

        Try
            Return DTA.ScalarQuery_GetJmlByProduct(vPrimaryCodeNo, vPrimaryDO).Value
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
            Return DTA.ScalarQuery_GetJml2(vPrimaryCodeNo, vPrimaryDO, _
            vBatch, vSecQty).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' mengembalikan jumlah item DO yang merujuk pada SO yang diberikan primary-nya
    ''' </summary>
    ''' <param name="vPrimarySO"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetJmlBySO(ByVal vPrimarySO As Double, ByVal vPrimaryProduct As Double) As Decimal
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do_itemTableAdapter
        Try
            Return DTA.ScalarQuery_GetJmlBySO_n_Product(vPrimarySO, vPrimaryProduct).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' mengembalikan jumlah item DO yang sesuai dengan parameter yang diberikan. Jumlahnya disini adalah jumlah do, jika
    ''' do-nya belum di stock out maka yang dikembalikan juga merupakan jumlah pending 
    ''' </summary>
    ''' <param name="vPrimaryDO"></param>
    ''' <param name="vPrimaryProduct"></param>
    ''' <param name="vSecondaryQty"></param>
    ''' <param name="vBatchNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
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

    Public Function GetJml_Default(ByVal vPrimaryDO As Double, ByVal vPrimarySO As Double, ByVal vPrimaryProduct As Double) As Decimal
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do_itemTableAdapter
        Try
            Return DTA.ScalarQuery_GetJml(vPrimaryDO, vPrimaryProduct, vPrimarySO)
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

    Public Function GetJml4_Default(ByVal vPrimaryDO As Double, ByVal vPrimarySO As Double, _
    ByVal vPrimaryProduct As Double, ByVal vBatchNo As String) As Decimal
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do_itemTableAdapter
        Try
            Return DTA.ScalarQuery_GetJml4(vPrimaryDO, vPrimaryProduct, vPrimarySO, vBatchNo)
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

    Public Function GetJmlBySO(ByVal vPrimarySO As Double) As Decimal
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do_itemTableAdapter
        Try
            Return DTA.ScalarQuery_GetJmlByPrimarySO(vPrimarySO).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDGV(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do_item_view1TableAdapter
        Dim TBL As New DseTransaksiDO.tbltrans_do_item_view1DataTable
        Dim row As DseTransaksiDO.tbltrans_do_item_view1Row
        Dim Sec_Units As String = ""
        Dim ObjCD As New Core_Module.ClsCekData
        Dim RowDgv As DataGridViewRow
        Dim VFrm As New FrmTransaksiDO3
        Dim ObjProduct As New Core_Module.clsMasterProduct
        Dim DefPrecMulti As Integer = 0

        Try
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            'kosongkan dgv
            vDGVItem.Rows.Clear()
            'isikan tabel
            If DTA.Fill(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    If row.Primary_Secondary_Units <> -1 Then
                        Sec_Units = row.Secondary_Unit
                    Else
                        Sec_Units = ""
                    End If

                    With row
                        'vDGVItem.Rows.Add("0", row.Product_Code, row.Product_Desc, ._Customer_s_Code_No, ._Customer_s_Desc, _
                        'row.SO_No, row.Batch_No, row.Expiry_Date, row.Qty_Master, .Qty_Per_Master, .Master_Unit, row.Qty, _
                        'ObjCD.DigitGrouping(.Total_Qty, ._decimal), row.Unit_Code, row.Secondary_Qty, Sec_Units, row.Remarks, row.Primary_Product, _
                        'row.Primary_SO, row.Primary_Master_Units, row.Primary_Units, row.Primary_Secondary_Units)
                        DefPrecMulti = ObjProduct.GetAngkaDecimal(.Primary_Product, .Primary_Units_Multi)
                        vDGVItem.Rows.Add()
                        RowDgv = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                        RowDgv.Cells(VFrm.ColNo.Index).Value = ""
                        RowDgv.Cells(VFrm.ColCodeNo.Index).Value = .Product_Code
                        RowDgv.Cells(VFrm.ColDescription.Index).Value = .Product_Desc
                        RowDgv.Cells(VFrm.ColCustCodeNo.Index).Value = ""
                        RowDgv.Cells(VFrm.ColCustDesc.Index).Value = ""
                        RowDgv.Cells(VFrm.ColSONO.Index).Value = .SO_No
                        RowDgv.Cells(VFrm.ColBatchNo.Index).Value = .Batch_No
                        RowDgv.Cells(VFrm.ColExpiryDate.Index).Value = Format(.Expiry_Date, "dd MMM yyyy")
                        RowDgv.Cells(VFrm.ColQtyMaster.Index).Value = .Qty_Master
                        RowDgv.Cells(VFrm.ColQtyPerMaster.Index).Value = .Qty_Per_Master
                        RowDgv.Cells(VFrm.ColMasterUnits.Index).Value = .Master_Unit
                        RowDgv.Cells(VFrm.ColQty.Index).Value = .Qty
                        RowDgv.Cells(VFrm.ColTotalQty.Index).Value = ObjCD.DigitGrouping(.Total_Qty, ._decimal)
                        RowDgv.Cells(VFrm.ColUnit.Index).Value = .Unit_Code
                        RowDgv.Cells(VFrm.ColSecQty.Index).Value = .Secondary_Qty
                        RowDgv.Cells(VFrm.ColSecUnit.Index).Value = .Secondary_Unit
                        RowDgv.Cells(VFrm.ColRemarks.Index).Value = .Remarks
                        RowDgv.Cells(VFrm.ColPrimaryCodeNo.Index).Value = .Primary_Product
                        RowDgv.Cells(VFrm.ColPrimarySO.Index).Value = .Primary_SO
                        RowDgv.Cells(VFrm.ColPrimaryMasterUnits.Index).Value = .Primary_Master_Units
                        RowDgv.Cells(VFrm.ColPrimaryUnit.Index).Value = .Primary_Units
                        RowDgv.Cells(VFrm.ColPrimarySecUnit.Index).Value = .Primary_Secondary_Units
                        RowDgv.Cells(VFrm.ColTotalQtyMulti.Index).Value = ObjCD.DigitGrouping(.Total_Qty_Multi, DefPrecMulti)
                        RowDgv.Cells(VFrm.ColUnitsCodeMulti.Index).Value = .Units_Code_Multi
                        RowDgv.Cells(VFrm.ColRatioDefault.Index).Value = ObjCD.DigitGrouping(.Ratio_Default, ._decimal)
                        RowDgv.Cells(VFrm.ColRatioReference.Index).Value = ObjCD.DigitGrouping(.Ratio_Reference, DefPrecMulti)
                        RowDgv.Cells(VFrm.ColPrimaryUnitsMulti.Index).Value = .Primary_Units_Multi
                        RowDgv.Cells(VFrm.ColPrimaryLocation.Index).Value = .primary_location
                        RowDgv.Cells(VFrm.ColLocation.Index).Value = .Storage_Area_Code
                        RowDgv.Cells(VFrm.ColWarehouse.Index).Value = .Warehouse_Code

                        If .Primary_Units_Multi <> .Primary_Units Then
                            RowDgv.Cells(VFrm.ColRatioDesc.Index).Value = ObjCD.DigitGrouping(.Ratio_Reference, DefPrecMulti) _
                            & " " & .Units_Code_Multi & " = " & ObjCD.DigitGrouping(.Ratio_Default, ._decimal) & " " & .Unit_Code
                        Else
                            RowDgv.Cells(VFrm.ColRatioDesc.Index).Value = "-"
                        End If
                        'TR/Tambahan Form option untuk pilihan delivery date
                        If .Delivery_Date = "#1/1/1950#" Then
                            RowDgv.Cells(VFrm.ColDeliveryDate.Index).Value = "-"
                        Else
                            RowDgv.Cells(VFrm.ColDeliveryDate.Index).Value = Format(.Delivery_Date, "dd MMM yyyy")
                        End If

                    End With
                Next
            End If
            'set nomer
            ObjCD.SetNumber(vDGVItem)
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
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

    Public Function Simpan(ByVal vdgvItem As DataGridView, ByVal vPrimaryTransaksi As Double, Optional ByVal SetOrderBalance As Boolean = True) As Boolean
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
        Dim ObjSOItem As New ClsTransaksiSalesOrderItem
        Dim DeliveryDate As Date = #1/1/1950#


        Try
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

                'TR/Tambahan Form option untuk pilihan delivery date
                If IsDate(row.Cells(vFrm.ColDeliveryDate.Index).Value) = False Then
                    DeliveryDate = #1/1/1950#
                Else
                    DeliveryDate = CDate(row.Cells(vFrm.ColDeliveryDate.Index).Value)
                End If

                'simpan ke basis data item

                'DTA.Insert(vPrimaryTransaksi, row.Cells(17).Value, row.Cells(18).Value, BatchNo, _
                'ExpiryDate, row.Cells(8).Value, row.Cells(9).Value, row.Cells(19).Value, _
                'row.Cells(11).Value, row.Cells(12).Value, row.Cells(20).Value, Sec_Qty, row.Cells(21).Value, _
                'row.Cells(16).Value)
                DTA.Insert(vPrimaryTransaksi, row.Cells(vFrm.ColPrimaryCodeNo.Index).Value, row.Cells(vFrm.ColPrimarySO.Index).Value, _
                row.Cells(vFrm.ColBatchNo.Index).Value, ExpiryDate, row.Cells(vFrm.ColQtyMaster.Index).Value, _
                row.Cells(vFrm.ColQtyPerMaster.Index).Value, row.Cells(vFrm.ColPrimaryMasterUnits.Index).Value, row.Cells(vFrm.ColQty.Index).Value, _
                row.Cells(vFrm.ColTotalQty.Index).Value, row.Cells(vFrm.ColPrimaryUnit.Index).Value, Sec_Qty, _
                row.Cells(vFrm.ColPrimarySecUnit.Index).Value, row.Cells(vFrm.ColRemarks.Index).Value, row.Cells(vFrm.ColTotalQtyMulti.Index).Value, _
                row.Cells(vFrm.ColPrimaryUnitsMulti.Index).Value, row.Cells(vFrm.ColRatioReference.Index).Value, _
                row.Cells(vFrm.ColRatioDefault.Index).Value, DeliveryDate, row.Cells(vFrm.ColPrimaryLocation.Index).Value)
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
        Dim ObjSOItem As New ClsTransaksiSalesOrderItem

        Try
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

    Public Sub IsikanTBLSO(ByVal vPrimaryDO As Double, ByVal TBL As DseTransaksiDO.DataTable2DataTable)
        Dim DTA As New DseTransaksiDOTableAdapters.DataTable2TableAdapter

        Try
            DTA.FillBy_PrimaryDO(TBL, vPrimaryDO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsikanTBLProductByDO(ByVal vPrimaryDO As Double, ByVal TBL As DseTransaksiDO.DataTable2DataTable)
        Dim DTA As New DseTransaksiDOTableAdapters.DataTable2TableAdapter
        Try
            DTA.FillBy_PrimaryDO(TBL, vPrimaryDO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsikanTBLProductBySO(ByVal vPrimaryDO As Double, ByVal TBL As DseTransaksiDO.DataTable2DataTable, _
    ByVal vPrimarySO As Double)
        Dim DTA As New DseTransaksiDOTableAdapters.DataTable2TableAdapter
        Try
            DTA.FillBy_PrimarySO(TBL, vPrimaryDO, vPrimarySO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsiTBL(ByVal vPrimaryDO As Double, ByVal TBL As DseTransaksiDO.tbltrans_do_itemDataTable)
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do_itemTableAdapter

        Try
            DTA.FillBy_PrimaryDO(TBL, vPrimaryDO)
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
            DTA.Fill(TBL, vPrimaryProduct, vPrimaryDO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsiSecondaryQty(ByVal vPrimaryDO As Double, ByVal vPrimaryProduct As Double, _
    ByVal vBatchNo As String, ByVal TBL As DseTransaksiDO.tbltrans_do_item2DataTable)
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do_item2TableAdapter

        Try
            DTA.Fill(TBL, vPrimaryDO, vPrimaryProduct, vBatchNo)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub
End Class
