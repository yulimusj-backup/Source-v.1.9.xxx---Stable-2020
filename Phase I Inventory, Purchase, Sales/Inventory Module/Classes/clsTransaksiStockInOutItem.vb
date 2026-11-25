Option Explicit On
Option Strict Off

Public Class clsTransaksiStockInOutItem

    Private Function ApakahTransferOtomatis(ByVal vTanggal As Date, ByVal vConString As String) As Boolean
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltrans_accounting_closingTableAdapter

        Try

            ''@rosadinnn
            If String.IsNullOrEmpty(vConString) = False Then
                DTA.Connection.ConnectionString = vConString
            End If

            'Transfer otomatis jika:
            '1. Tanggal invoice >= convertion date
            If DTA.ScalarQuery_GetConvertionDate.HasValue = True Then
                If DTA.ScalarQuery_GetConvertionDate.Value.Date <= vTanggal.Date Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Function ApakahTransferOtomatis(ByVal vTanggal As Date) As Boolean
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltrans_accounting_closingTableAdapter

        Try
            'Transfer otomatis jika:
            '1. Tanggal invoice >= convertion date
            If DTA.ScalarQuery_GetConvertionDate.HasValue = True Then
                If DTA.ScalarQuery_GetConvertionDate.Value.Date <= vTanggal.Date Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlReturn(ByVal vPrimarySO As Double, ByVal vPrimaryProduct As Double, _
    ByVal vPrimaryDO As Double, ByVal vPrimaryTransaksi As Double) As Decimal
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinout_itemTableAdapter

        Try
            Return DTA.ScalarQuery_GetJmlBy_PrimaryStockIn_PrimaryProduct_PrimaryDO_PrimarySO(vPrimaryProduct, _
            vPrimarySO, vPrimaryDO, vPrimaryTransaksi).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    'Public Function GetJmlReturn(ByVal vPrimarySO As Double, ByVal vPrimaryProduct As Double, _
    'ByVal vPrimaryDO As Double, ByVal vPrimaryTransaksi As Double, ByVal vBatchNo As String, _
    'ByVal vSecQty As Double) As Double
    '    Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinout_itemTableAdapter

    '    Try
    '        Return DTA.ScalarQuery_GetJml4(vPrimaryTransaksi, vPrimaryProduct, _
    '        vBatchNo, vPrimarySO, vSecQty, vPrimaryDO).Value
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTA.Dispose()
    '    End Try
    'End Function

    'Public Function GetJmlReturn(ByVal vPrimarySO As Double, ByVal vPrimaryProduct As Double, ByVal vPrimaryDO As Double, _
    'ByVal vBatchNo As String, ByVal vSecQty As Double) As Double
    '    Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinout_itemTableAdapter
    '    Dim TipeTransaksi As clsTransaksiStockInOut.EnumTipeStockInOut = _
    '    clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn

    '    Try
    '        Return DTA.ScalarQuery_GetJml3(vPrimaryProduct, _
    '        vBatchNo, vPrimarySO, vSecQty, TipeTransaksi, vPrimaryDO).Value
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTA.Dispose()
    '    End Try
    'End Function

    Public Function GetJmlReturn(ByVal vPrimarySO As Double, ByVal vPrimaryProduct As Double, ByVal vPrimaryDO As Double) As Decimal
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinout_itemTableAdapter
        Dim TipeTransaksi As clsTransaksiStockInOut.EnumTipeStockInOut = _
        clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn

        Try
            Return DTA.ScalarQuery_GetJmlBy_PrimaryProduct_PrimarySO_TipeTransaksi_PrimaryDO(vPrimaryProduct, _
            vPrimarySO, TipeTransaksi, vPrimaryDO).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    'yang dicek hanya item yang tipenya goods
    Public Function ApakahHapusMengakibatkanMinus(ByVal vPrimaryTranskasi As Double, ByVal _
    vTypeTransaksi As clsTransaksiStockInOut.EnumTipeStockInOut, ByVal vTampilkanPesan As Boolean, ByVal vPrimary_Antrian As Double, _
    ByVal vDTAInOUT As Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter, ByRef vSudahHapusAntrian_ByRef As Boolean) As Boolean
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinout_itemTableAdapter
        Dim TBL As New dseTransaksiStockInOut.tbltransaksi_stockinout_itemDataTable
        Dim row As dseTransaksiStockInOut.tbltransaksi_stockinout_itemRow
        Dim JmlLama As Double = 0
        Dim JmlBaru As Double = 0
        Dim ObjProduct As New Core_Module.clsMasterProduct
        Dim DTA1 As New Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter
        Dim Dummy As String = ""
        Dim TipeItemGoods As Core_Module.clsMasterProduct.EnumTipeProduct = Core_Module.clsMasterProduct.EnumTipeProduct.Goods
        Dim ListPrimaryProduct As New List(Of Double)

        Try
            If vTypeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust _
            Or vTypeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_BackProcess _
            Or vTypeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Process _
            Or vTypeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase _
            Or vTypeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn Then
                DTA.FillBy_Primary_Simulasi(TBL, vPrimaryTranskasi, TipeItemGoods)

                For Each row In TBL.Rows
                    JmlLama = ObjProduct.GetJml2(row.Primary_Product, DTA1, row.Secondary_Qty, _
                    row.Primary_Storage_Area, row.Batch_No)
                    If JmlLama <> 0 Then
                        JmlBaru = JmlLama - row.Qty
                        If JmlBaru < 0 Then
                            If ListPrimaryProduct.Contains(row.Primary_Product) = False Then
                                'jumlah setelah hapus transaksi akan menjadi minus
                                If Dummy = "" Then
                                    Dummy = ObjProduct.GetProductCode(row.Primary_Product)
                                Else
                                    Dummy = Dummy + ", " + ObjProduct.GetProductCode(row.Primary_Product)
                                End If
                                ListPrimaryProduct.Add(row.Primary_Product)
                            End If
                        End If
                    Else
                        'tidak ada data digudang, data akan menjadi minus jika item dihapuskan
                        If ListPrimaryProduct.Contains(row.Primary_Product) = False Then
                            If Dummy = "" Then
                                Dummy = ObjProduct.GetProductCode(row.Primary_Product)
                            Else
                                Dummy = Dummy + ", " + ObjProduct.GetProductCode(row.Primary_Product)
                            End If
                            ListPrimaryProduct.Add(row.Primary_Product)
                        End If
                    End If
                Next
            Else
                'transaksi stock out ketika dihapuskan akan menambahkan stock gudang sehingga
                'pasti tidak akan mengakibatkan stock menjadi minus
                Return False
            End If

            'kalo sampe sini return false, tidak ada data yang minus jika transaksi ini dihapuskan
            If Dummy <> "" Then
                If vTampilkanPesan = True Then
                    If vSudahHapusAntrian_ByRef = False Then
                        If vDTAInOUT.Delete(vPrimary_Antrian) = 0 Then
                            MsgBox("Logic error(1), contact your provider when you see this message", _
                            MsgBoxStyle.Information, "Information")
                        End If
                        vSudahHapusAntrian_ByRef = True
                    End If

                    MsgBox("You cannot delete this data because will cause negatif qty one or several product: " + _
                    Dummy, MsgBoxStyle.Information, "Information")
                End If
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
            ObjProduct = Nothing
            TBL.Dispose()
            ListPrimaryProduct = Nothing
        End Try
    End Function

    Public Function GetJml(ByVal vPrimaryStockIn As Double, ByVal vPrimaryCodeNo As Double, ByVal Secondary_Qty As _
    Double, ByVal vPrimaryLocation As Double, ByVal vBatchNo As String, ByVal vPrimaryPdO As Double, _
    ByVal ExpiryDate As Date) As Decimal
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinout_itemTableAdapter
        Try

            If DTA.ScalarQuery_GetJml(Secondary_Qty, vBatchNo, ExpiryDate, vPrimaryLocation, vPrimaryPdO, _
            vPrimaryStockIn, vPrimaryCodeNo).HasValue = True Then
                Return DTA.ScalarQuery_GetJml(Secondary_Qty, vBatchNo, ExpiryDate, vPrimaryLocation, vPrimaryPdO, _
                vPrimaryStockIn, vPrimaryCodeNo).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml_Default(ByVal vPrimaryStockIn As Double, ByVal vPrimaryCodeNo As Double, ByVal vPrimaryPO As Double) As Decimal
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinout_itemTableAdapter
        Try
            If DTA.ScalarQuery_getJml5(vPrimaryStockIn, vPrimaryCodeNo, vPrimaryPO).HasValue = True Then
                Return DTA.ScalarQuery_getJml5(vPrimaryStockIn, vPrimaryCodeNo, vPrimaryPO).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    ''@rosadinnn
    ''@MUaT
    Public Function GetJml_Multi(ByVal vPrimaryStockIn As Double, ByVal vPrimaryCodeNo As Double, ByVal vPrimaryPO As Double) As Decimal
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinout_itemTableAdapter
        Try
            If DTA.ScalarQuery_GetJml5Multi(vPrimaryStockIn, vPrimaryCodeNo, vPrimaryPO).HasValue = True Then
                Return DTA.ScalarQuery_GetJml5Multi(vPrimaryStockIn, vPrimaryCodeNo, vPrimaryPO).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub IsiTBLPO_Default(ByVal vPrimaryStock As Double, ByVal vTBL As _
      dseTransaksiStockInOut.tbltransaksistockinpurchase_viewDataTable)
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksistockinpurchase_viewTableAdapter
        Try
            DTA.FillBy_PrimaryStockInOut(vTBL, vPrimaryStock)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    ''@rosadinnn
    ''@MUaT
    Public Sub IsiTBLPO_Multi(ByVal vPrimaryStock As Double, ByVal vTBL As _
      dseTransaksiStockInOut.tbltransaksistockinpurchase_viewMultiDataTable)
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksistockinpurchase_viewMultiTableAdapter
        Try
            DTA.FillBy_PrimaryStockInOut(vTBL, vPrimaryStock)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    'Public Sub IsiTBLProductByPO(ByVal vPrimaryStock As Double, ByVal vTBL As _
    ' dseTransaksiStockInOut.tbltransaksistockinpurchase_viewDataTable, ByVal vPrimaryPO As Double)
    '    Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksistockinpurchase_viewTableAdapter
    '    Try
    '        DTA.FillBy_PrimaryStockNPrimaryPO(vTBL, vPrimaryStock, vPrimaryPO)
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTA.Dispose()
    '    End Try
    'End Sub

    ''@rosadinnn
    ''@MUaT
    Public Sub IsiTBLProductByPO(ByVal vPrimaryStock As Double, ByVal vTBL As _
            dseTransaksiStockInOut.tbltransaksistockinpurchase_viewMultiDataTable, ByVal vPrimaryPO As Double)
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksistockinpurchase_viewMultiTableAdapter
        Try
            DTA.FillBy_PrimaryStockNPrimaryPO(vTBL, vPrimaryStock, vPrimaryPO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    ''' <summary>
    ''' Menggunakan No SO
    ''' </summary>
    ''' <param name="vDGVItem"></param>
    ''' <param name="vPrimaryTransaksi"></param>
    ''' <param name="vTipeTransaksi"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsiDGV2_(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double, _
      ByVal vTipeTransaksi As clsTransaksiStockInOut.EnumTipeStockInOut) As Boolean
        Dim DTA As New dseTransaksiStockInOutTableAdapters.DataTable2_oldTableAdapter
        Dim TBL As New dseTransaksiStockInOut.DataTable2_oldDataTable
        Dim row As dseTransaksiStockInOut.DataTable2_oldRow
        Dim NoSO, NoPDO, Sec_Units, Location, Warehouse As String
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            'kosongkan dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            vDGVItem.Rows.Clear()

            'isikan data
            If vTipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Sales Or _
            vTipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn Then
                Me.IsiDGVStockOutSales(vDGVItem, vPrimaryTransaksi)
                If vDGVItem.RowCount <> 0 Then
                    Exit Function
                End If
            End If

            'isikan tabel
            If DTA.Fill(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    If row.Primary_PdO = -1 Then
                        NoPDO = ""
                    Else
                        'msgbox("belum")
                        NoPDO = "No-PDO-perbaiki lagi"
                    End If
                    If row.Primary_PO_Purchase = -1 Then
                        NoSO = ""
                    Else
                        NoSO = row.So_No
                    End If
                    If row.Primary_Storage_Area = -1 Then
                        Location = ""
                        Warehouse = ""
                    Else
                        Location = row.Storage_Area_Code
                        Warehouse = row.Warehouse_Code
                    End If
                    If row.Primary_Secondary_Units <> -1 Then
                        Sec_Units = row.Secondary_Units
                    Else
                        Sec_Units = ""
                    End If

                    vDGVItem.Rows.Add("0", NoSO, row.Supp__DO_No, row.Product_Code, row.Product_Description, _
                    ObjCD.DigitGrouping(row.Qty, row._Decimal), row.Unit_Code, row.Secondary_Qty, Sec_Units, row.Batch_No, row.Expiry_Date, _
                    Location, Warehouse, row.Remarks, row.Primary_PO_Purchase, row.Primary_Product, _
                    row.Primary_Storage_Area, row.Primary_Units, row.Primary_Secondary_Units, _
                    row.Primary_PdO)
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

    ''@rosadinnn
    ''@MUaT
    Public Function IsiDGV(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double, _
            ByVal vTipeTransaksi As clsTransaksiStockInOut.EnumTipeStockInOut, _
            ByVal vTbl As CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable) As Boolean

        Dim objProduct As New Core_Module.clsMasterProduct
        Dim vFrmRefPR As New FrmTransaksiStockOutPurchaseReturn2
        Dim vFrmRefPurchase As New FrmTransaksiStockInPurchase2

        Dim DTA As New dseTransaksiStockInOutTableAdapters.transaksi_stockinout_item_viewTableAdapter
        Dim TBL As New dseTransaksiStockInOut.transaksi_stockinout_item_viewDataTable
        Dim row As dseTransaksiStockInOut.transaksi_stockinout_item_viewRow
        Dim NoPO, NoPDO, Sec_Units, Location, Warehouse As String
        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim DefPrec As Integer = 0

        Dim rowDgv As DataGridViewRow
        Dim ObjCustom As New CustomSMESS.ClsTransaksiStockOutSales_AddByCustomProduct
        Try
            'kosongkan dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            vDGVItem.Rows.Clear()
            DefPrec = ObjCur.GetDecimal_DefaultCurrency

            ''isikan data
            'If vTipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Sales Or _
            'vTipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn Or _
            'vTipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust Then
            '    Me.IsiDGVStockOutSales(vDGVItem, vPrimaryTransaksi)
            '    If vDGVItem.RowCount <> 0 Then
            '        Exit Function
            '    End If
            'End If

            'isikan data
            If vTipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust Then
                Me.IsiDGVSales(vDGVItem, vPrimaryTransaksi)
                If vDGVItem.RowCount <> 0 Then
                    Exit Function
                End If
            End If
            If vTipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Sales Then
                Me.IsiDGVStockOutSales(vDGVItem, vPrimaryTransaksi)
                If vDGVItem.RowCount <> 0 Then
                    Exit Function
                End If
            End If
            If vTipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn Then
                Me.IsiDGVStockInSalesReturn(vDGVItem, vPrimaryTransaksi)
                If vDGVItem.RowCount <> 0 Then
                    Exit Function
                End If
            End If

            'isikan tabel
            If DTA.Fill(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    If row.Primary_PdO = -1 Then
                        NoPDO = ""
                    Else
                        'msgbox("belum")
                        NoPDO = "No-PDO-perbaiki lagi"
                    End If
                    If row.Primary_PO_Purchase = -1 Then
                        NoPO = ""
                    Else
                        NoPO = row.No_PO
                    End If
                    If row.Primary_Storage_Area = -1 Then
                        Location = ""
                        Warehouse = ""
                    Else
                        Location = row.Storage_Area_Code
                        Warehouse = row.Warehouse_Code
                    End If
                    If row.Primary_Secondary_Units <> -1 Then
                        Sec_Units = row.Secondary_Units
                    Else
                        Sec_Units = ""
                    End If

                    If vTipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust Then
                        vDGVItem.Rows.Add("0", NoPO, row.Supp__DO_No, row.Product_Code, row.Product_Description, _
                        ObjCD.DigitGrouping(row.Qty, row._Decimal), row.Unit_Code, row.Secondary_Qty, Sec_Units, row.Batch_No, row.Expiry_Date, _
                        Location, Warehouse, row.Remarks, row.Primary_PO_Purchase, row.Primary_Product, _
                        row.Primary_Storage_Area, row.Primary_Units, row.Primary_Secondary_Units, _
                        row.Primary_PdO, ObjCD.DigitGrouping(row.Unit_Price, DefPrec), ObjCD.DigitGrouping(row.Amount, DefPrec))
                    ElseIf vTipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase Then
                        vDGVItem.Rows.Add()
                        rowDgv = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                        With rowDgv
                            .Cells(vFrmRefPurchase.ColNo.Index).Value = ""
                            .Cells(vFrmRefPurchase.ColPONo.Index).Value = NoPO
                            .Cells(vFrmRefPurchase.ColSuppDONo.Index).Value = row.Supp__DO_No
                            .Cells(vFrmRefPurchase.ColCodeNo.Index).Value = row.Product_Code
                            .Cells(vFrmRefPurchase.ColDescription.Index).Value = row.Product_Description
                            .Cells(vFrmRefPurchase.ColQtyMulti.Index).Value = ObjCD.DigitGrouping(row.Qty_Multi, objProduct.GetAngkaDecimal(row.Primary_Product, row.Primary_Unit_Multi))
                            .Cells(vFrmRefPurchase.ColUnitCodeMulti.Index).Value = row.Unit_Code_Multi
                            If row.Primary_Units <> row.Primary_Unit_Multi Then
                                .Cells(vFrmRefPurchase.ColRatioDesc.Index).Value = ObjCD.DigitGrouping(row.Ratio_Reference, _
                                        objProduct.GetAngkaDecimal(row.Primary_Product, row.Primary_Unit_Multi)) _
                                        & " " & row.Unit_Code_Multi & " = " & ObjCD.DigitGrouping(row.Ratio_Default, _
                                        row._Decimal) & " " & row.Unit_Code
                            Else
                                .Cells(vFrmRefPurchase.ColRatioDesc.Index).Value = "-"
                            End If
                            .Cells(vFrmRefPurchase.ColQty.Index).Value = ObjCD.DigitGrouping(row.Qty, row._Decimal)
                            .Cells(vFrmRefPurchase.ColUnit.Index).Value = (row.Unit_Code)
                            .Cells(vFrmRefPurchase.ColSecondaryQty.Index).Value = row.Secondary_Qty
                            .Cells(vFrmRefPurchase.ColSecondaryUnit.Index).Value = Sec_Units
                            .Cells(vFrmRefPurchase.ColBatch.Index).Value = row.Batch_No
                            .Cells(vFrmRefPurchase.ColExpiredDate.Index).Value = Format(row.Expiry_Date, "dd MMM yyyy")
                            .Cells(vFrmRefPurchase.ColLocation.Index).Value = Location
                            .Cells(vFrmRefPurchase.ColWareHouse.Index).Value = Warehouse
                            .Cells(vFrmRefPurchase.ColRemark.Index).Value = row.Remarks
                            .Cells(vFrmRefPurchase.ColPrimaryPOno.Index).Value = row.Primary_PO_Purchase
                            .Cells(vFrmRefPurchase.ColPrimaryCodeNo.Index).Value = row.Primary_Product
                            .Cells(vFrmRefPurchase.ColPrimaryLocation.Index).Value = row.Primary_Storage_Area
                            .Cells(vFrmRefPurchase.ColPrimaryUnits.Index).Value = row.Primary_Units
                            .Cells(vFrmRefPurchase.ColPrimarySecondaryUnits.Index).Value = row.Primary_Secondary_Units
                            .Cells(vFrmRefPurchase.ColPrimaryPdO.Index).Value = row.Primary_PdO
                            .Cells(vFrmRefPurchase.ColPrimaryUnitMulti.Index).Value = row.Primary_Unit_Multi
                            .Cells(vFrmRefPurchase.ColRatioReference.Index).Value = ObjCD.DigitGrouping(row.Ratio_Reference, objProduct.GetAngkaDecimal(row.Primary_Product, row.Primary_Unit_Multi))
                            .Cells(vFrmRefPurchase.ColRatioDefault.Index).Value = ObjCD.DigitGrouping(row.Ratio_Default, objProduct.GetAngkaDecimal(row.Primary_Product))
                        End With
                    ElseIf vTipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn Then
                        vDGVItem.Rows.Add()
                        rowDgv = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                        With rowDgv
                            .Cells(vFrmRefPR.ColNo.Index).Value = ""
                            .Cells(vFrmRefPR.ColPONoDONo.Index).Value = NoPO
                            .Cells(vFrmRefPR.ColSuppDONo.Index).Value = row.Supp__DO_No
                            .Cells(vFrmRefPR.ColCodeNo.Index).Value = row.Product_Code
                            .Cells(vFrmRefPR.ColDescription.Index).Value = row.Product_Description
                            .Cells(vFrmRefPR.ColQty.Index).Value = ObjCD.DigitGrouping(row.Qty, row._Decimal)
                            .Cells(vFrmRefPR.ColUnit.Index).Value = row.Unit_Code
                            .Cells(vFrmRefPR.ColSecondaryQty.Index).Value = row.Secondary_Qty
                            .Cells(vFrmRefPR.ColSecondaryUnit.Index).Value = Sec_Units
                            .Cells(vFrmRefPR.ColBatch.Index).Value = row.Batch_No
                            .Cells(vFrmRefPR.ColExpiredDate.Index).Value = Format(row.Expiry_Date, "dd MMM yyyy")
                            .Cells(vFrmRefPR.ColLocation.Index).Value = Location
                            .Cells(vFrmRefPR.ColWareHouse.Index).Value = Warehouse
                            .Cells(vFrmRefPR.ColRemark.Index).Value = row.Remarks
                            .Cells(vFrmRefPR.ColPrimaryPOnoDoNo.Index).Value = row.Primary_PO_Purchase
                            .Cells(vFrmRefPR.ColPrimaryCodeNo.Index).Value = row.Primary_Product
                            .Cells(vFrmRefPR.ColPrimaryLocation.Index).Value = row.Primary_Storage_Area
                            .Cells(vFrmRefPR.ColPrimaryUnits.Index).Value = row.Primary_Units
                            .Cells(vFrmRefPR.ColPrimarySecondaryUnits.Index).Value = row.Primary_Secondary_Units
                            .Cells(vFrmRefPR.ColPrimaryPdO.Index).Value = row.Primary_PdO

                            .Cells(vFrmRefPR.ColQtyMulti.Index).Value = ObjCD.DigitGrouping(row.Qty_Multi, objProduct.GetAngkaDecimal(row.Primary_Product, row.Primary_Unit_Multi))
                            .Cells(vFrmRefPR.ColUnitCodeMulti.Index).Value = row.Unit_Code_Multi
                            If row.Primary_Units <> row.Primary_Unit_Multi Then
                                .Cells(vFrmRefPR.ColRatioDesc.Index).Value = ObjCD.DigitGrouping(row.Ratio_Reference, _
                                        objProduct.GetAngkaDecimal(row.Primary_Product, row.Primary_Unit_Multi)) _
                                        & " " & row.Unit_Code_Multi & " = " & ObjCD.DigitGrouping(row.Ratio_Default, _
                                        row._Decimal) & " " & row.Unit_Code
                            Else
                                .Cells(vFrmRefPR.ColRatioDesc.Index).Value = "-"
                            End If
                            .Cells(vFrmRefPR.ColRatioReference.Index).Value = ObjCD.DigitGrouping(row.Ratio_Reference, objProduct.GetAngkaDecimal(row.Primary_Product, row.Primary_Unit_Multi))
                            .Cells(vFrmRefPR.ColRatioDefault.Index).Value = ObjCD.DigitGrouping(row.Ratio_Default, objProduct.GetAngkaDecimal(row.Primary_Product))
                            .Cells(vFrmRefPR.ColPrimaryUnitMulti.Index).Value = row.Primary_Unit_Multi
                        End With
                    Else
                        vDGVItem.Rows.Add("0", NoPO, row.Supp__DO_No, row.Product_Code, row.Product_Description, _
                        ObjCD.DigitGrouping(row.Qty, row._Decimal), row.Unit_Code, row.Secondary_Qty, Sec_Units, row.Batch_No, row.Expiry_Date, _
                        Location, Warehouse, row.Remarks, row.Primary_PO_Purchase, row.Primary_Product, _
                        row.Primary_Storage_Area, row.Primary_Units, row.Primary_Secondary_Units, _
                        row.Primary_PdO)
                    End If
                Next
            End If
            'set nomer
            ObjCD.SetNumber(vDGVItem)
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            'nilai balikan

            'ObjCustom.IsiTblDummyFromDB(vPrimaryTransaksi, vTbl)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
            ObjCD = Nothing
            vFrmRefPR.Dispose()
            vFrmRefPurchase.Dispose()
        End Try
    End Function

    Private Function IsiDGVSales(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New dseTransaksiStockInOutTableAdapters.DataTable2_oldTableAdapter
        Dim TBL As New dseTransaksiStockInOut.DataTable2_oldDataTable
        Dim row As dseTransaksiStockInOut.DataTable2_oldRow
        Dim NoPDO, Sec_Units, Location, Warehouse As String
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            'isikan tabel
            If DTA.Fill(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    If row.Primary_PdO = -1 Then
                        NoPDO = ""
                    Else
                        'msgbox("belum")
                        NoPDO = "No-PDO-perbaiki lagi"
                    End If
                    If row.Primary_Storage_Area = -1 Then
                        Location = ""
                        Warehouse = ""
                    Else
                        Location = row.Storage_Area_Code
                        Warehouse = row.Warehouse_Code
                    End If
                    If row.Primary_Secondary_Units <> -1 Then
                        Sec_Units = row.Secondary_Units
                    Else
                        Sec_Units = ""
                    End If

                    vDGVItem.Rows.Add("0", row.So_No, row.Supp__DO_No, row.Product_Code, row.Product_Description, _
                    ObjCD.DigitGrouping(row.Qty, row._Decimal), row.Unit_Code, row.Secondary_Qty, Sec_Units, row.Batch_No, row.Expiry_Date, _
                    Location, Warehouse, row.Remarks, row.Primary_PO_Purchase, row.Primary_Product, _
                    row.Primary_Storage_Area, row.Primary_Units, row.Primary_Secondary_Units, _
                    row.Primary_PdO)
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

    Private Function IsiDGVStockOutSales(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New dseTransaksiStockInOutTableAdapters.DataTable2_NewTableAdapter
        Dim TBL As New dseTransaksiStockInOut.DataTable2_NewDataTable
        Dim row As dseTransaksiStockInOut.DataTable2_NewRow
        Dim NoPDO, Sec_Units, Location, Warehouse As String
        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjProduct As New Core_Module.clsMasterProduct
        Dim i As New FrmTransaksiStockOutSales
        Dim rowDgv As DataGridViewRow

        Try
            'isikan tabel
            If DTA.Fill(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    If row.Primary_PdO = -1 Then
                        NoPDO = ""
                    Else
                        'msgbox("belum")
                        NoPDO = "No-PDO-perbaiki lagi"
                    End If
                    If row.Primary_Storage_Area = -1 Then
                        Location = ""
                        Warehouse = ""
                    Else
                        Location = row.Storage_Area_Code
                        Warehouse = row.Warehouse_Code
                    End If
                    If row.Primary_Secondary_Units <> -1 Then
                        Sec_Units = row.Secondary_Units
                    Else
                        Sec_Units = ""
                    End If

                    Dim DefPrecDefault As Integer = ObjProduct.GetAngkaDecimal(row.Primary_Product)
                    Dim DefPrecMulti As Integer = ObjProduct.GetAngkaDecimal(row.Primary_Product, row.Primary_Multi_Units)
                    vDGVItem.Rows.Add()
                    rowDgv = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                    With rowDgv
                        .Cells(i.ColNo.Index).Value = ""
                        .Cells(i.ColPONoDONo.Index).Value = row.SO_NO
                        .Cells(i.ColSuppDONo.Index).Value = row.Supp__DO_No
                        .Cells(i.ColCodeNo.Index).Value = row.Product_Code
                        .Cells(i.ColDescription.Index).Value = row.Product_Description
                        .Cells(i.ColQty.Index).Value = ObjCD.DigitGrouping(row.Qty, DefPrecDefault)
                        .Cells(i.ColUnit.Index).Value = row.Unit_Code
                        .Cells(i.ColSecondaryQty.Index).Value = row.Secondary_Qty
                        .Cells(i.ColSecondaryUnit.Index).Value = Sec_Units
                        .Cells(i.ColBatch.Index).Value = row.Batch_No
                        .Cells(i.ColExpiredDate.Index).Value = Format(row.Expiry_Date, "dd MMM yyyy")
                        .Cells(i.ColLocation.Index).Value = Location
                        .Cells(i.ColWareHouse.Index).Value = Warehouse
                        .Cells(i.ColRemark.Index).Value = row.Remarks
                        .Cells(i.ColPrimaryPOnoDoNo.Index).Value = row.Primary_PO_Purchase
                        .Cells(i.ColPrimaryCodeNo.Index).Value = row.Primary_Product
                        .Cells(i.ColPrimaryLocation.Index).Value = row.Primary_Storage_Area
                        .Cells(i.ColPrimaryUnits.Index).Value = row.Primary_Units
                        .Cells(i.ColPrimarySecondaryUnits.Index).Value = row.Primary_Secondary_Units
                        .Cells(i.ColPrimaryPdO.Index).Value = row.Primary_PdO
                        .Cells(i.ColQtyMulti.Index).Value = ObjCD.DigitGrouping(row.Qty_Multi_Unit, DefPrecMulti)
                        .Cells(i.ColUnitsCodeMulti.Index).Value = row.Units_Code_Multi
                        .Cells(i.ColPrimaryUnitsMulti.Index).Value = row.Primary_Multi_Units
                        .Cells(i.ColRatioDefault.Index).Value = ObjCD.DigitGrouping(row.Ratio_Default, DefPrecDefault)
                        .Cells(i.ColRatioReference.Index).Value = ObjCD.DigitGrouping(row.Ratio_Reference, DefPrecMulti)
                        If row.Primary_Multi_Units <> row.Primary_Units Then
                            .Cells(i.ColRatioDesc.Index).Value = ObjCD.DigitGrouping(row.Ratio_Reference, DefPrecMulti) _
                            & " " & row.Units_Code_Multi & " = " & ObjCD.DigitGrouping(row.Ratio_Default, DefPrecDefault) & " " & row.Unit_Code
                        Else
                            .Cells(i.ColRatioDesc.Index).Value = "-"
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
        'Dim DTA As New dseTransaksiStockInOutTableAdapters.DataTable2_NewTableAdapter
        'Dim TBL As New dseTransaksiStockInOut.DataTable2_NewDataTable
        'Dim row As dseTransaksiStockInOut.DataTable2_NewRow
        'Dim NoPDO, Sec_Units, Location, Warehouse As String
        'Dim ObjCD As New Core_Module.ClsCekData
        'Dim decimalMulti As Integer = 0
        'Dim ObjProd As New Core_Module.clsMasterProduct
        'Try
        '    'isikan tabel
        '    If DTA.Fill(TBL, vPrimaryTransaksi) > 0 Then
        '        For Each row In TBL.Rows
        '            If row.Primary_PdO = -1 Then
        '                NoPDO = ""
        '            Else
        '                'msgbox("belum")
        '                NoPDO = "No-PDO-perbaiki lagi"
        '            End If
        '            If row.Primary_Storage_Area = -1 Then
        '                Location = ""
        '                Warehouse = ""
        '            Else
        '                Location = row.Storage_Area_Code
        '                Warehouse = row.Warehouse_Code
        '            End If
        '            If row.Primary_Secondary_Units <> -1 Then
        '                Sec_Units = row.Secondary_Units
        '            Else
        '                Sec_Units = ""
        '            End If
        '            decimalMulti = ObjProd.GetAngkaDecimal(row.Primary_Product, row.Primary_Multi_Units)
        '            vDGVItem.Rows.Add("0", row.SO_NO, row.Supp__DO_No, row.Product_Code, row.Product_Description, _
        '            ObjCD.DigitGrouping(row.Qty, row._Decimal), row.Unit_Code, row.Ratio_Default, ObjCD.DigitGrouping(row.Qty_Multi_Unit, decimalMulti), _
        '            row.Qty_Multi_Unit, "", "", row.Batch_No, row.Expiry_Date, Location, Warehouse, row.Remarks, row.Primary_PO_Purchase, _
        '            row.Primary_Product, row.Primary_Storage_Area, row.Primary_Units, row.Primary_Secondary_Units, _
        '            row.Primary_PdO, row.Ratio_Default, row.Ratio_Reference, row.Primary_Multi_Units)
        '        Next
        '    End If
        '    'set nomer
        '    ObjCD.SetNumber(vDGVItem)
        '    'nilai balikan
        '    Return True
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'Finally
        '    DTA.Dispose()
        '    TBL = Nothing
        '    ObjCD = Nothing
        'End Try
    End Function

    Private Function IsiDGVStockInSalesReturn(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New dseTransaksiStockInOutTableAdapters.DataTable2_NewTableAdapter
        Dim TBL As New dseTransaksiStockInOut.DataTable2_NewDataTable
        Dim row As dseTransaksiStockInOut.DataTable2_NewRow
        Dim NoPDO, Sec_Units, Location, Warehouse As String
        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjProduct As New Core_Module.clsMasterProduct
        Dim i As New FrmTransaksiStockInSalesReturn2
        Dim rowDgv As DataGridViewRow

        Try
            'isikan tabel
            If DTA.Fill(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    If row.Primary_PdO = -1 Then
                        NoPDO = ""
                    Else
                        'msgbox("belum")
                        NoPDO = "No-PDO-perbaiki lagi"
                    End If
                    If row.Primary_Storage_Area = -1 Then
                        Location = ""
                        Warehouse = ""
                    Else
                        Location = row.Storage_Area_Code
                        Warehouse = row.Warehouse_Code
                    End If
                    If row.Primary_Secondary_Units <> -1 Then
                        Sec_Units = row.Secondary_Units
                    Else
                        Sec_Units = ""
                    End If

                    Dim DefPrecDefault As Integer = ObjProduct.GetAngkaDecimal(row.Primary_Product)
                    Dim DefPrecMulti As Integer = ObjProduct.GetAngkaDecimal(row.Primary_Product, row.Primary_Multi_Units)
                    vDGVItem.Rows.Add()
                    rowDgv = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                    With rowDgv
                        .Cells(i.ColNo.Index).Value = ""
                        .Cells(i.ColPONoDONo.Index).Value = row.SO_NO
                        .Cells(i.ColSuppDONo.Index).Value = row.Supp__DO_No
                        .Cells(i.ColCodeNo.Index).Value = row.Product_Code
                        .Cells(i.ColDescription.Index).Value = row.Product_Description
                        .Cells(i.ColQty.Index).Value = ObjCD.DigitGrouping(row.Qty, DefPrecDefault)
                        .Cells(i.ColUnit.Index).Value = row.Unit_Code
                        .Cells(i.ColSecondaryQty.Index).Value = row.Secondary_Qty
                        .Cells(i.ColSecondaryUnit.Index).Value = Sec_Units
                        .Cells(i.ColBatch.Index).Value = row.Batch_No
                        .Cells(i.ColExpiredDate.Index).Value = Format(row.Expiry_Date, "dd MMM yyyy")
                        .Cells(i.ColLocation.Index).Value = Location
                        .Cells(i.ColWareHouse.Index).Value = Warehouse
                        .Cells(i.ColRemark.Index).Value = row.Remarks
                        .Cells(i.ColPrimarySO.Index).Value = row.Primary_PO_Purchase
                        .Cells(i.ColPrimaryCodeNo.Index).Value = row.Primary_Product
                        .Cells(i.ColPrimaryLocation.Index).Value = row.Primary_Storage_Area
                        .Cells(i.ColPrimaryUnits.Index).Value = row.Primary_Units
                        .Cells(i.ColPrimarySecondaryUnits.Index).Value = row.Primary_Secondary_Units
                        .Cells(i.ColPrimaryPdO.Index).Value = row.Primary_PdO
                        .Cells(i.ColQtyMulti.Index).Value = ObjCD.DigitGrouping(row.Qty_Multi_Unit, DefPrecMulti)
                        .Cells(i.ColUnitsCodeMulti.Index).Value = row.Units_Code_Multi
                        .Cells(i.ColPrimaryUnitsMulti.Index).Value = row.Primary_Multi_Units
                        .Cells(i.ColRatioDefault.Index).Value = ObjCD.DigitGrouping(row.Ratio_Default, DefPrecDefault)
                        .Cells(i.ColRatioReference.Index).Value = ObjCD.DigitGrouping(row.Ratio_Reference, DefPrecMulti)
                        If row.Primary_Multi_Units <> row.Primary_Units Then
                            .Cells(i.ColRatioDesc.Index).Value = ObjCD.DigitGrouping(row.Ratio_Reference, DefPrecMulti) _
                            & " " & row.Units_Code_Multi & " = " & ObjCD.DigitGrouping(row.Ratio_Default, DefPrecDefault) & " " & row.Unit_Code
                        Else
                            .Cells(i.ColRatioDesc.Index).Value = "-"
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

    Public Function Simpan_StockOutSales(ByVal vdgvItem As DataGridView, ByVal vTanggalTransaksi As Date, _
    ByVal vPrimaryTransaksi As Double, ByVal vNoTransaksi As String, Optional ByVal vSupplierName As String = "", _
    Optional ByVal vCustomerName As String = "", Optional ByVal vSetPOB As Boolean = True, _
    Optional ByVal vBeginingBalance As Boolean = False, Optional ByVal vKoreksiProdQty As Boolean = True) As Boolean
        Dim row As DataGridViewRow
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinout_itemTableAdapter
        Dim ObjProd As New Core_Module.clsMasterProduct
        Dim ObjSC As New ClsTransaksiStockCard
        Dim ObjLS As New Core_Module.clsMasterListedSoftware
        Dim DTA1 As New Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter
        Dim InOut As ClsTransaksiStockCard.EnumJenisStockCard = ClsTransaksiStockCard.EnumJenisStockCard.Stock_In
        Dim JenisSC As ClsTransaksiStockCard.EnumJenisTransaksiAsal = ClsTransaksiStockCard.EnumJenisStockCard.Stock_In
        Dim Qty_In, Qty_Out As Double
        Dim DTASC As New dseTransaksiStockCardTableAdapters.tbltransaksi_stock_cardTableAdapter
        Dim IsSalesInstalled As Boolean = ObjLS.IsSalesInstalled
        Dim IsPurchaseInstalled As Boolean = ObjLS.IsPurchaseInstalled
        Dim DescSC As String = ""
        Dim Sec_Qty, PrimaryLocation, PrimaryPdO As Double
        Dim BatchNo As String = ""
        Dim ExpiryDate As Date = #1/1/1950#
        Dim ijk As New FrmTransaksiStockInAdj
        Dim TransferOtomatis As Boolean = Me.ApakahTransferOtomatis(vTanggalTransaksi)
        Dim ObjPOItem As New ClsTransaksiPOItem
        Dim vListPrimaryPO As New List(Of Double)
        Dim i As New FrmTransaksiStockOutSales
        Try
            JenisSC = ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockOut_Sales
            InOut = ClsTransaksiStockCard.EnumJenisStockCard.Stock_Out
            DescSC = "Sales to " + vCustomerName
            For Each row In vdgvItem.Rows
                BatchNo = row.Cells(i.ColBatch.Index).Value
                PrimaryLocation = row.Cells(i.ColPrimaryLocation.Index).Value
                If IsNumeric(row.Cells(i.ColSecondaryQty.Index).Value) = True Then
                    Sec_Qty = row.Cells(i.ColSecondaryQty.Index).Value
                Else
                    Sec_Qty = "-1"
                End If
                PrimaryPdO = row.Cells(i.ColPrimaryPdO.Index).Value
                If IsDate(row.Cells(i.ColExpiredDate.Index).Value) = False Then
                    ExpiryDate = #1/1/1950#
                Else
                    ExpiryDate = CDate(row.Cells(i.ColExpiredDate.Index).Value)
                End If
                DTA.Insert_Multi(vPrimaryTransaksi, _
                    row.Cells(i.ColPrimaryCodeNo.Index).Value, _
                    row.Cells(i.ColQty.Index).Value, _
                    row.Cells(i.ColPrimaryUnits.Index).Value, _
                    Sec_Qty, _
                    row.Cells(i.ColPrimarySecondaryUnits.Index).Value, _
                    BatchNo, _
                    ExpiryDate, _
                    PrimaryLocation, _
                    row.Cells(i.ColPrimaryPOnoDoNo.Index).Value, _
                    PrimaryPdO, _
                    row.Cells(i.ColSuppDONo.Index).Value, _
                    row.Cells(i.ColRemark.Index).Value, _
                    row.Cells(i.ColQtyMulti.Index).Value, _
                    row.Cells(i.ColPrimaryUnitsMulti.Index).Value, _
                    row.Cells(i.ColRatioReference.Index).Value, _
                    row.Cells(i.ColRatioDefault.Index).Value)
                If TransferOtomatis = True Or vBeginingBalance = True Then
                    If vKoreksiProdQty = True Then
                        Me.Koreksi4SimpanOut(vTanggalTransaksi, ObjProd, row.Cells(i.ColPrimaryCodeNo.Index).Value, DTA1, row.Cells(i.ColQty.Index).Value, Sec_Qty, PrimaryLocation, BatchNo, ExpiryDate)
                    End If
                    Qty_In = 0
                    Qty_Out = row.Cells(i.ColQty.Index).Value
                    ObjSC.Simpan(vPrimaryTransaksi, JenisSC, vNoTransaksi, row.Cells(i.ColPrimaryCodeNo.Index).Value, _
                    Qty_In, Qty_Out, DescSC, InOut, vTanggalTransaksi, DTASC, BatchNo, ExpiryDate, PrimaryLocation, row.Cells(i.ColRemark.Index).Value)
                End If
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
            ObjProd = Nothing
            ObjSC = Nothing
            ObjLS = Nothing
            ObjPOItem = Nothing
            vListPrimaryPO = Nothing
            ijk = Nothing
        End Try
    End Function

    Public Function Simpan_StockInSalesReturn(ByVal vdgvItem As DataGridView, ByVal vTanggalTransaksi As Date, _
            ByVal vPrimaryTransaksi As Double, ByVal vNoTransaksi As String, Optional ByVal vSupplierName As String = "", _
            Optional ByVal vCustomerName As String = "", Optional ByVal vSetPOB As Boolean = True, _
            Optional ByVal vBeginingBalance As Boolean = False, Optional ByVal vKoreksiProdQty As Boolean = True) As Boolean
        Dim row As DataGridViewRow
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinout_itemTableAdapter
        Dim ObjProd As New Core_Module.clsMasterProduct
        Dim ObjSC As New ClsTransaksiStockCard
        Dim ObjLS As New Core_Module.clsMasterListedSoftware
        Dim DTA1 As New Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter
        Dim InOut As ClsTransaksiStockCard.EnumJenisStockCard = ClsTransaksiStockCard.EnumJenisStockCard.Stock_In
        Dim JenisSC As ClsTransaksiStockCard.EnumJenisTransaksiAsal = ClsTransaksiStockCard.EnumJenisStockCard.Stock_In
        Dim Qty_In, Qty_Out As Double
        Dim DTASC As New dseTransaksiStockCardTableAdapters.tbltransaksi_stock_cardTableAdapter
        Dim IsSalesInstalled As Boolean = ObjLS.IsSalesInstalled
        Dim IsPurchaseInstalled As Boolean = ObjLS.IsPurchaseInstalled
        Dim DescSC As String = ""
        Dim Sec_Qty, PrimaryLocation, PrimaryPdO As Double
        Dim BatchNo As String = ""
        Dim ExpiryDate As Date = #1/1/1950#
        Dim ijk As New FrmTransaksiStockInAdj
        Dim TransferOtomatis As Boolean = Me.ApakahTransferOtomatis(vTanggalTransaksi)
        Dim ObjPOItem As New ClsTransaksiPOItem
        Dim vListPrimaryPO As New List(Of Double)

        Dim i As New FrmTransaksiStockInSalesReturn2
        Try
            JenisSC = Me.GetTipeStockCard(clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn)
            InOut = ClsTransaksiStockCard.EnumJenisStockCard.Stock_In
            DescSC = "Sales return from " + vCustomerName

            For Each row In vdgvItem.Rows
                BatchNo = row.Cells(i.ColBatch.Index).Value
                PrimaryLocation = row.Cells(i.ColPrimaryLocation.Index).Value
                If IsNumeric(row.Cells(i.ColSecondaryQty.Index).Value) = True Then
                    Sec_Qty = row.Cells(i.ColSecondaryQty.Index).Value
                Else
                    Sec_Qty = "-1"
                End If
                PrimaryPdO = row.Cells(i.ColPrimaryPdO.Index).Value
                If IsDate(row.Cells(i.ColExpiredDate.Index).Value) = False Then
                    ExpiryDate = #1/1/1950#
                Else
                    ExpiryDate = CDate(row.Cells(i.ColExpiredDate.Index).Value)
                End If

                DTA.Insert_Multi(vPrimaryTransaksi, _
                    row.Cells(i.ColPrimaryCodeNo.Index).Value, _
                    row.Cells(i.ColQty.Index).Value, _
                    row.Cells(i.ColPrimaryUnits.Index).Value, _
                    Sec_Qty, _
                    row.Cells(i.ColPrimarySecondaryUnits.Index).Value, _
                    BatchNo, _
                    ExpiryDate, _
                    PrimaryLocation, _
                    row.Cells(i.ColPrimarySO.Index).Value, _
                    PrimaryPdO, _
                    row.Cells(i.ColSuppDONo.Index).Value, _
                    row.Cells(i.ColRemark.Index).Value, _
                    row.Cells(i.ColQtyMulti.Index).Value, _
                    row.Cells(i.ColPrimaryUnitsMulti.Index).Value, _
                    row.Cells(i.ColRatioReference.Index).Value, _
                    row.Cells(i.ColRatioDefault.Index).Value)

                If TransferOtomatis = True Or vBeginingBalance = True Then
                    If vKoreksiProdQty = True Then
                        Me.Koreksi4SimpanIn(vTanggalTransaksi, ObjProd, row.Cells(i.ColPrimaryCodeNo.Index).Value, DTA1, row.Cells(i.ColQty.Index).Value, _
                        Sec_Qty, PrimaryLocation, BatchNo, PrimaryPdO, ExpiryDate)
                    End If
                    Qty_In = row.Cells(i.ColQty.Index).Value
                    Qty_Out = 0
                    ObjSC.Simpan(vPrimaryTransaksi, JenisSC, vNoTransaksi, row.Cells(i.ColPrimaryCodeNo.Index).Value, _
                    Qty_In, Qty_Out, DescSC, InOut, vTanggalTransaksi, DTASC, BatchNo, ExpiryDate, PrimaryLocation, row.Cells(i.ColRemark.Index).Value)
                End If
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
            ObjProd = Nothing
            ObjSC = Nothing
            ObjLS = Nothing
            ObjPOItem = Nothing
            vListPrimaryPO = Nothing
            ijk = Nothing
        End Try
    End Function

    Public Function Simpan(ByVal vdgvItem As DataGridView, ByVal vTypeTransaksi As _
    clsTransaksiStockInOut.EnumTipeStockInOut, ByVal vTanggalTransaksi As Date, _
    ByVal vPrimaryTransaksi As Double, ByVal vNoTransaksi As String, Optional ByVal vSupplierName As String _
    = "", Optional ByVal vCustomerName As String = "", Optional ByVal vSetPOB As Boolean = True, _
    Optional ByVal vBeginingBalance As Boolean = False, Optional ByVal vKoreksiProdQty As Boolean = True) As Boolean
        Dim row As DataGridViewRow
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinout_itemTableAdapter
        Dim ObjProd As New Core_Module.clsMasterProduct
        Dim ObjSC As New ClsTransaksiStockCard
        Dim ObjLS As New Core_Module.clsMasterListedSoftware
        Dim DTA1 As New Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter

        Dim InOut As ClsTransaksiStockCard.EnumJenisStockCard = ClsTransaksiStockCard.EnumJenisStockCard.Stock_In
        Dim JenisSC As ClsTransaksiStockCard.EnumJenisTransaksiAsal = ClsTransaksiStockCard.EnumJenisStockCard.Stock_In
        Dim Qty_In, Qty_Out As Double
        Dim DTASC As New dseTransaksiStockCardTableAdapters.tbltransaksi_stock_cardTableAdapter

        Dim IsSalesInstalled As Boolean = ObjLS.IsSalesInstalled
        Dim IsPurchaseInstalled As Boolean = ObjLS.IsPurchaseInstalled
        Dim DescSC As String = ""
        Dim Sec_Qty, PrimaryLocation, PrimaryPdO As Double
        Dim BatchNo As String = ""
        Dim ExpiryDate As Date = #1/1/1950#
        Dim ijk As New FrmTransaksiStockInAdj
        Dim TransferOtomatis As Boolean = Me.ApakahTransferOtomatis(vTanggalTransaksi)

        'elemen yang berkaitan deng perhitungan purchase order balance
        Dim ObjPOItem As New ClsTransaksiPOItem
        Dim vListPrimaryPO As New List(Of Double)

        Try
            'cek jenis transaksi stock card
            JenisSC = Me.GetTipeStockCard(vTypeTransaksi)
            'cek in/out
            Select Case vTypeTransaksi
                Case clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust
                    InOut = ClsTransaksiStockCard.EnumJenisStockCard.Stock_In
                    DescSC = "Stock in adjust"
                    Exit Select
                Case clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_BackProcess
                    InOut = ClsTransaksiStockCard.EnumJenisStockCard.Stock_In
                    DescSC = "Stock in back process"
                    Exit Select
                Case clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Process
                    InOut = ClsTransaksiStockCard.EnumJenisStockCard.Stock_In
                    DescSC = "Stock in process"
                    Exit Select
                Case clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase
                    InOut = ClsTransaksiStockCard.EnumJenisStockCard.Stock_In
                    DescSC = "Purchase from " + vSupplierName
                    Exit Select
                Case clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn
                    InOut = ClsTransaksiStockCard.EnumJenisStockCard.Stock_In
                    DescSC = "Sales return from " + vCustomerName
                    Exit Select
                Case clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust
                    InOut = ClsTransaksiStockCard.EnumJenisStockCard.Stock_Out
                    DescSC = "Stock out adjust"
                    Exit Select
                Case clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Process
                    InOut = ClsTransaksiStockCard.EnumJenisStockCard.Stock_Out
                    DescSC = "Stock out process"
                    Exit Select
                Case clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn
                    InOut = ClsTransaksiStockCard.EnumJenisStockCard.Stock_Out
                    DescSC = "Purchase return to " + vSupplierName
                    Exit Select
                Case clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Sales
                    InOut = ClsTransaksiStockCard.EnumJenisStockCard.Stock_Out
                    DescSC = "Sales to " + vCustomerName
                    Exit Select
            End Select

            For Each row In vdgvItem.Rows
                BatchNo = row.Cells(9).Value
                PrimaryLocation = row.Cells(16).Value
                If IsNumeric(row.Cells(7).Value) = True Then
                    Sec_Qty = row.Cells(7).Value
                Else
                    Sec_Qty = "-1"
                End If
                PrimaryPdO = row.Cells(19).Value
                If IsDate(row.Cells(10).Value) = False Then
                    ExpiryDate = #1/1/1950#
                Else
                    ExpiryDate = CDate(row.Cells(10).Value)
                End If

                'simpan ke basis data item
                If vTypeTransaksi <> clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust Then
                    'DTA.Insert(vPrimaryTransaksi, row.Cells(15).Value, row.Cells(5).Value, row.Cells(17).Value, _
                    'Sec_Qty, row.Cells(18).Value, BatchNo, ExpiryDate, _
                    'PrimaryLocation, row.Cells(14).Value, PrimaryPdO, row.Cells(2).Value, _
                    'row.Cells(13).Value)
                    DTA.Insert_StockInPurchase(vPrimaryTransaksi, row.Cells(15).Value, row.Cells(5).Value, row.Cells(17).Value, _
                    Sec_Qty, row.Cells(18).Value, BatchNo, ExpiryDate, _
                    PrimaryLocation, row.Cells(14).Value, PrimaryPdO, row.Cells(2).Value, _
                    row.Cells(13).Value, row.Cells(5).Value, row.Cells(17).Value, 1, 1)
                Else
                    'DTA.InsertQuery(vPrimaryTransaksi, row.Cells(15).Value, row.Cells(5).Value, row.Cells(17).Value, _
                    'Sec_Qty, row.Cells(18).Value, BatchNo, ExpiryDate, _
                    'PrimaryLocation, row.Cells(14).Value, PrimaryPdO, row.Cells(2).Value, _
                    'row.Cells(13).Value, row.Cells(ijk.ColAmount.Index).Value)
                    DTA.InsertQuery1(vPrimaryTransaksi, row.Cells(15).Value, row.Cells(5).Value, row.Cells(17).Value, _
                    Sec_Qty, row.Cells(18).Value, BatchNo, ExpiryDate, _
                    PrimaryLocation, row.Cells(14).Value, PrimaryPdO, row.Cells(2).Value, _
                    row.Cells(13).Value, row.Cells(ijk.ColAmount.Index).Value, row.Cells(17).Value, row.Cells(5).Value, _
                    1, 1, row.Cells(ijk.ColUnitCost.Index).Value)
                End If

                'lakukan koreksi di qty detail
                If TransferOtomatis = True Or vBeginingBalance = True Then
                    If vTypeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust Or _
                                  vTypeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_BackProcess Or _
                                  vTypeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Process Or _
                                  vTypeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase Or _
                                  vTypeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn Then
                        If vKoreksiProdQty = True Then
                            Me.Koreksi4SimpanIn(vTanggalTransaksi, ObjProd, row.Cells(15).Value, DTA1, row.Cells(5).Value, _
                            Sec_Qty, PrimaryLocation, BatchNo, PrimaryPdO, ExpiryDate)
                        End If
                    Else
                        If vKoreksiProdQty = True Then
                            Me.Koreksi4SimpanOut(vTanggalTransaksi, ObjProd, row.Cells(15).Value, DTA1, row.Cells(5).Value, _
                            Sec_Qty, PrimaryLocation, BatchNo, ExpiryDate)
                        End If
                    End If

                    'simpan data stock card
                    If InOut = ClsTransaksiStockCard.EnumJenisStockCard.Stock_In Then
                        Qty_In = row.Cells(5).Value
                        Qty_Out = 0
                    Else
                        Qty_In = 0
                        Qty_Out = row.Cells(5).Value
                    End If
                    ObjSC.Simpan(vPrimaryTransaksi, JenisSC, vNoTransaksi, row.Cells(15).Value, _
                    Qty_In, Qty_Out, DescSC, InOut, vTanggalTransaksi, DTASC, BatchNo, ExpiryDate, PrimaryLocation, row.Cells(13).Value)
                End If

                'untuk transaksi stock in purchase return set nilai purhcase order balance.
                'untuk stock out purchase return, tidak lagi mengeset nilai purchase order balance karena
                'sudah di set di transaksi purchase return
                If vTypeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase Then
                    If vListPrimaryPO.Contains(row.Cells(14).Value) = False Then
                        vListPrimaryPO.Add(row.Cells(14).Value)
                    End If
                End If
            Next
            If vTypeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase Or _
            vTypeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn Then
                'set order balance
                If vSetPOB = True Then
                    ObjPOItem.SetOrderBalance(vListPrimaryPO)
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
            ObjProd = Nothing
            ObjSC = Nothing
            ObjLS = Nothing
            ObjPOItem = Nothing
            vListPrimaryPO = Nothing
            ijk = Nothing
        End Try
    End Function

    ''@rosadinnn
    ''@MUaT
    Public Function Simpan_StockInPurchase(ByVal vdgvItem As DataGridView, ByVal vTanggalTransaksi As Date, _
            ByVal vPrimaryTransaksi As Double, ByVal vNoTransaksi As String, Optional ByVal vSupplierName As String = "", _
            Optional ByVal vCustomerName As String = "", Optional ByVal vSetPOB As Boolean = True, _
            Optional ByVal vBeginingBalance As Boolean = False, Optional ByVal vKoreksiProdQty As Boolean = True) As Boolean

        Dim i As New FrmTransaksiStockInPurchase2

        Dim row As DataGridViewRow
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinout_itemTableAdapter
        Dim ObjProd As New Core_Module.clsMasterProduct
        Dim ObjSC As New ClsTransaksiStockCard
        Dim ObjLS As New Core_Module.clsMasterListedSoftware
        Dim DTA1 As New Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter

        Dim InOut As ClsTransaksiStockCard.EnumJenisStockCard = ClsTransaksiStockCard.EnumJenisStockCard.Stock_In
        Dim JenisSC As ClsTransaksiStockCard.EnumJenisTransaksiAsal = ClsTransaksiStockCard.EnumJenisStockCard.Stock_In
        Dim Qty_In, Qty_Out As Double
        Dim DTASC As New dseTransaksiStockCardTableAdapters.tbltransaksi_stock_cardTableAdapter

        Dim IsSalesInstalled As Boolean = ObjLS.IsSalesInstalled
        Dim IsPurchaseInstalled As Boolean = ObjLS.IsPurchaseInstalled
        Dim DescSC As String = "Purchase from " + vSupplierName
        Dim Sec_Qty, PrimaryLocation, PrimaryPdO As Double
        Dim BatchNo As String = ""
        Dim ExpiryDate As Date = #1/1/1950#
        Dim ijk As New FrmTransaksiStockInAdj
        Dim TransferOtomatis As Boolean = Me.ApakahTransferOtomatis(vTanggalTransaksi)

        'elemen yang berkaitan deng perhitungan purchase order balance
        Dim ObjPOItem As New ClsTransaksiPOItem
        Dim vListPrimaryPO As New List(Of Double)

        Try
            JenisSC = Me.GetTipeStockCard(clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase)

            For Each row In vdgvItem.Rows

                BatchNo = row.Cells(i.ColBatch.Index).Value
                PrimaryLocation = row.Cells(i.ColPrimaryLocation.Index).Value
                If IsNumeric(row.Cells(i.ColSecondaryQty.Index).Value) = True Then
                    Sec_Qty = row.Cells(i.ColSecondaryQty.Index).Value
                Else
                    Sec_Qty = "-1"
                End If
                PrimaryPdO = row.Cells(i.ColPrimaryPdO.Index).Value
                If IsDate(row.Cells(i.ColExpiredDate.Index).Value) = False Then
                    ExpiryDate = #1/1/1950#
                Else
                    ExpiryDate = CDate(row.Cells(i.ColExpiredDate.Index).Value)
                End If

                DTA.Insert_StockInPurchase(vPrimaryTransaksi, row.Cells(i.ColPrimaryCodeNo.Index).Value, row.Cells(i.ColQty.Index).Value, _
                                row.Cells(i.ColPrimaryUnits.Index).Value, Sec_Qty, row.Cells(i.ColPrimarySecondaryUnits.Index).Value, _
                                BatchNo, ExpiryDate, PrimaryLocation, row.Cells(i.ColPrimaryPOno.Index).Value, PrimaryPdO, _
                                row.Cells(i.ColSuppDONo.Index).Value, row.Cells(i.ColRemark.Index).Value, _
                                row.Cells(i.ColQtyMulti.Index).Value, row.Cells(i.ColPrimaryUnitMulti.Index).Value, _
                                row.Cells(i.ColRatioReference.Index).Value, row.Cells(i.ColRatioDefault.Index).Value)

                'lakukan koreksi di qty detail
                If TransferOtomatis = True Or vBeginingBalance = True Then
                    If vKoreksiProdQty = True Then
                        Me.Koreksi4SimpanIn(vTanggalTransaksi, ObjProd, row.Cells(i.ColPrimaryCodeNo.Index).Value, DTA1, row.Cells(i.ColQty.Index).Value, _
                        Sec_Qty, PrimaryLocation, BatchNo, PrimaryPdO, ExpiryDate)
                    End If

                    'simpan data stock card
                    Qty_In = row.Cells(i.ColQty.Index).Value
                    Qty_Out = 0
                    ObjSC.Simpan(vPrimaryTransaksi, JenisSC, vNoTransaksi, row.Cells(i.ColPrimaryCodeNo.Index).Value, _
                    Qty_In, Qty_Out, DescSC, InOut, vTanggalTransaksi, DTASC, BatchNo, ExpiryDate, PrimaryLocation, row.Cells(i.ColRemark.Index).Value)
                End If

                If vListPrimaryPO.Contains(row.Cells(i.ColPrimaryPOno.Index).Value) = False Then
                    vListPrimaryPO.Add(row.Cells(i.ColPrimaryPOno.Index).Value)
                End If

            Next

            If vSetPOB = True Then
                ObjPOItem.SetOrderBalance(vListPrimaryPO)
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
            ObjProd = Nothing
            ObjSC = Nothing
            ObjLS = Nothing
            ObjPOItem = Nothing
            vListPrimaryPO = Nothing
            ijk = Nothing
        End Try
    End Function

    Public Function Simpan_StockOutPurchaseReturn(ByVal vdgvItem As DataGridView, ByVal vTanggalTransaksi As Date, _
            ByVal vPrimaryTransaksi As Double, ByVal vNoTransaksi As String, Optional ByVal vSupplierName As String = "", _
            Optional ByVal vCustomerName As String = "", Optional ByVal vSetPOB As Boolean = True, _
            Optional ByVal vBeginingBalance As Boolean = False, Optional ByVal vKoreksiProdQty As Boolean = True) As Boolean

        Dim i As New FrmTransaksiStockOutPurchaseReturn2

        Dim row As DataGridViewRow
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinout_itemTableAdapter
        Dim ObjProd As New Core_Module.clsMasterProduct
        Dim ObjSC As New ClsTransaksiStockCard
        Dim ObjLS As New Core_Module.clsMasterListedSoftware
        Dim DTA1 As New Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter

        Dim InOut As ClsTransaksiStockCard.EnumJenisStockCard = ClsTransaksiStockCard.EnumJenisStockCard.Stock_Out
        Dim JenisSC As ClsTransaksiStockCard.EnumJenisTransaksiAsal = ClsTransaksiStockCard.EnumJenisStockCard.Stock_In
        Dim Qty_In, Qty_Out As Double
        Dim DTASC As New dseTransaksiStockCardTableAdapters.tbltransaksi_stock_cardTableAdapter

        Dim IsSalesInstalled As Boolean = ObjLS.IsSalesInstalled
        Dim IsPurchaseInstalled As Boolean = ObjLS.IsPurchaseInstalled
        Dim DescSC As String = "Purchase return to " + vSupplierName
        Dim Sec_Qty, PrimaryLocation, PrimaryPdO As Double
        Dim BatchNo As String = ""
        Dim ExpiryDate As Date = #1/1/1950#
        Dim ijk As New FrmTransaksiStockInAdj
        Dim TransferOtomatis As Boolean = Me.ApakahTransferOtomatis(vTanggalTransaksi)

        'elemen yang berkaitan deng perhitungan purchase order balance
        Dim ObjPOItem As New ClsTransaksiPOItem
        Dim vListPrimaryPO As New List(Of Double)

        Try
            JenisSC = Me.GetTipeStockCard(clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn)

            For Each row In vdgvItem.Rows
                BatchNo = row.Cells(i.ColBatch.Index).Value
                PrimaryLocation = row.Cells(i.ColPrimaryLocation.Index).Value
                If IsNumeric(row.Cells(i.ColSecondaryQty.Index).Value) = True Then
                    Sec_Qty = row.Cells(i.ColSecondaryQty.Index).Value
                Else
                    Sec_Qty = "-1"
                End If
                PrimaryPdO = row.Cells(i.ColPrimaryPdO.Index).Value
                If IsDate(row.Cells(i.ColExpiredDate.Index).Value) = False Then
                    ExpiryDate = #1/1/1950#
                Else
                    ExpiryDate = CDate(row.Cells(i.ColExpiredDate.Index).Value)
                End If

                DTA.Insert_StockInPurchase(vPrimaryTransaksi, row.Cells(i.ColPrimaryCodeNo.Index).Value, row.Cells(i.ColQty.Index).Value, _
                                row.Cells(i.ColPrimaryUnits.Index).Value, Sec_Qty, row.Cells(i.ColPrimarySecondaryUnits.Index).Value, _
                                BatchNo, ExpiryDate, PrimaryLocation, row.Cells(i.ColPrimaryPOnoDoNo.Index).Value, PrimaryPdO, _
                                row.Cells(i.ColSuppDONo.Index).Value, row.Cells(i.ColRemark.Index).Value, _
                                row.Cells(i.ColQtyMulti.Index).Value, row.Cells(i.ColPrimaryUnitMulti.Index).Value, _
                                row.Cells(i.ColRatioReference.Index).Value, row.Cells(i.ColRatioDefault.Index).Value)

                'lakukan koreksi di qty detail
                If TransferOtomatis = True Or vBeginingBalance = True Then
                    If vKoreksiProdQty = True Then
                        Me.Koreksi4SimpanOut(vTanggalTransaksi, ObjProd, row.Cells(i.ColPrimaryCodeNo.Index).Value, DTA1, row.Cells(i.ColQty.Index).Value, _
                        Sec_Qty, PrimaryLocation, BatchNo, ExpiryDate)
                    End If

                    'simpan data stock card
                    Qty_In = 0
                    Qty_Out = row.Cells(i.ColQty.Index).Value
                    ObjSC.Simpan(vPrimaryTransaksi, JenisSC, vNoTransaksi, row.Cells(i.ColPrimaryCodeNo.Index).Value, _
                    Qty_In, Qty_Out, DescSC, InOut, vTanggalTransaksi, DTASC, BatchNo, ExpiryDate, PrimaryLocation, row.Cells(i.ColRemark.Index).Value)
                End If

            Next
            If vSetPOB = True Then
                ObjPOItem.SetOrderBalance(vListPrimaryPO)
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
            ObjProd = Nothing
            ObjSC = Nothing
            ObjLS = Nothing
            ObjPOItem = Nothing
            vListPrimaryPO = Nothing
            ijk = Nothing
        End Try
    End Function

    Private Function GetTipeStockCard(ByVal vTipeStockInOut As clsTransaksiStockInOut.EnumTipeStockInOut) As ClsTransaksiStockCard.EnumJenisTransaksiAsal
        Try
            Select Case vTipeStockInOut
                Case clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust
                    Return ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockIn_Adjust
                Case clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_BackProcess
                    Return ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockIn_BackProcess
                Case clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Process
                    Return ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockIn_Process
                Case clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase
                    Return ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockIn_Purchase
                Case clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn
                    Return ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockIn_SalesReturn
                Case clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust
                    Return ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockOut_Adjust
                Case clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Process
                    Return ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockOut_Process
                Case clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn
                    Return ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockOut_PurchaseReturn
                Case clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Sales
                    Return ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockOut_Sales
            End Select
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function Hapus(ByVal vPrimaryTransaksi As Double, _
            ByVal vTipeTransaksi As clsTransaksiStockInOut.EnumTipeStockInOut, ByVal vConString As String, _
            Optional ByVal vSetPOB As Boolean = True, Optional ByVal vBeginingBalance As Boolean = False) As Boolean
        Dim DTA As New dseTransaksiStockInOutTableAdapters.transaksi_stockinout_item_viewTableAdapter
        Dim TBL As New dseTransaksiStockInOut.transaksi_stockinout_item_viewDataTable
        Dim row As dseTransaksiStockInOut.transaksi_stockinout_item_viewRow
        Dim ObjSC As New ClsTransaksiStockCard
        Dim JenisSC As ClsTransaksiStockCard.EnumJenisTransaksiAsal
        Dim MovingIn As Boolean
        Dim DTA1 As New Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter
        Dim DTA2 As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinout_itemTableAdapter
        Dim ObjProd As New Core_Module.clsMasterProduct
        Dim Sec_Qty, PrimaryLocation, PrimaryPdO As Double
        Dim BatchNo As String = ""
        Dim ExpiryDate As Date = #1/1/1950#
        Dim ObjStockInOut As New clsTransaksiStockInOut
        Dim vTanggalTrans As Date = Today

        Try

            ''@rosadinnn
            If String.IsNullOrEmpty(vConString) = False Then
                DTA.Connection.ConnectionString = vConString
                Core_Module.SetConstring_DTAProductQty(DTA1)
                DTA2.Connection.ConnectionString = vConString
            End If

            vTanggalTrans = ObjStockInOut.GetTanggalTransaksi(vPrimaryTransaksi, vConString)

            If vTipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust Or _
            vTipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_BackProcess Or _
            vTipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Process Or _
            vTipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase Or _
            vTipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn Then
                MovingIn = True
            Else
                MovingIn = False
            End If

            'load data item lama
            If Me.ApakahTransferOtomatis(vTanggalTrans, vConString) = True Or vBeginingBalance = True Then
                DTA.Fill(TBL, vPrimaryTransaksi)
                'lakukan koreksi stock di qty detail
                For Each row In TBL.Rows
                    'lakukan koreksi di qty detail
                    If CBool(row.Batch_No_Controlled) = True Then
                        'batch no controlled
                        BatchNo = row.Batch_No
                    Else
                        BatchNo = ""
                    End If
                    If CBool(row.Location_Controlled) = True Then
                        'location controlled
                        PrimaryLocation = row.Primary_Storage_Area
                    Else
                        PrimaryLocation = -1
                    End If
                    If CBool(row.Use_Secondary_Unit) = True Then
                        'use secondary
                        Sec_Qty = row.Secondary_Qty
                    Else
                        Sec_Qty = -1
                    End If
                    If CBool(row.PdOControlled) = True Then
                        'PdO controlled
                        PrimaryPdO = row.Primary_PdO
                    Else
                        PrimaryPdO = -1
                    End If
                    If CBool(row.Expiry_Date_Controlled) = True Then
                        'expiry date controlled
                        ExpiryDate = row.Expiry_Date
                    Else
                        ExpiryDate = "1/1/1950"
                    End If

                    If vTipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust Or _
                    vTipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_BackProcess Or _
                    vTipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Process Or _
                    vTipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase Or _
                    vTipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn Then
                        Me.Koreksi4HapusIn(vTanggalTrans, ObjProd, row.Primary_Product, DTA1, row.Qty, _
                        Sec_Qty, PrimaryLocation, BatchNo, PrimaryPdO, ExpiryDate)
                    Else
                        Me.Koreksi4HapusOut(vTanggalTrans, ObjProd, row.Primary_Product, DTA1, row.Qty, _
                        Sec_Qty, PrimaryLocation, BatchNo, ExpiryDate)
                    End If
                Next
            End If


            'hapus data stock card
            JenisSC = Me.GetTipeStockCard(vTipeTransaksi)
            ObjSC.Hapus(vPrimaryTransaksi, JenisSC)

            'hapus data item
            DTA2.DeleteQuery(vPrimaryTransaksi)

            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
            TBL.Dispose()
            ObjSC = Nothing
            ObjProd = Nothing
            'ObjPOItem = Nothing
            'ObjStock = Nothing
            'vListPrimaryPO = Nothing
            ObjStockInOut = Nothing
        End Try
    End Function

    Public Function Hapus(ByVal vPrimaryTransaksi As Double, _
            ByVal vTipeTransaksi As clsTransaksiStockInOut.EnumTipeStockInOut, Optional ByVal vSetPOB As Boolean = True, Optional ByVal vBeginingBalance As Boolean = False) As Boolean
        Dim DTA As New dseTransaksiStockInOutTableAdapters.transaksi_stockinout_item_viewTableAdapter
        Dim TBL As New dseTransaksiStockInOut.transaksi_stockinout_item_viewDataTable
        Dim row As dseTransaksiStockInOut.transaksi_stockinout_item_viewRow
        Dim ObjSC As New ClsTransaksiStockCard
        Dim JenisSC As ClsTransaksiStockCard.EnumJenisTransaksiAsal
        Dim MovingIn As Boolean
        Dim DTA1 As New Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter
        Dim DTA2 As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinout_itemTableAdapter
        Dim ObjProd As New Core_Module.clsMasterProduct
        Dim Sec_Qty, PrimaryLocation, PrimaryPdO As Double
        Dim BatchNo As String = ""
        Dim ExpiryDate As Date = #1/1/1950#
        'Dim ObjPOItem As New ClsTransaksiPOItem
        'Dim ObjStock As New clsTransaksiStockInOut
        'Dim vListPrimaryPO As New List(Of Double)
        Dim ObjStockInOut As New clsTransaksiStockInOut
        Dim vTanggalTrans As Date = Today

        Try

            vTanggalTrans = ObjStockInOut.GetTanggalTransaksi(vPrimaryTransaksi)

            If vTipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust Or _
            vTipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_BackProcess Or _
            vTipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Process Or _
            vTipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase Or _
            vTipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn Then
                MovingIn = True
            Else
                MovingIn = False
            End If
            'If vTipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase Or _
            'vTipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn Then
            'ObjStock.GetPrimaryPO(vPrimaryTransaksi, vListPrimaryPO)
            'End If
            'load data item lama
            If Me.ApakahTransferOtomatis(vTanggalTrans) = True Or vBeginingBalance = True Then
                DTA.Fill(TBL, vPrimaryTransaksi)
                'lakukan koreksi stock di qty detail
                For Each row In TBL.Rows
                    'lakukan koreksi di qty detail
                    If CBool(row.Batch_No_Controlled) = True Then
                        'batch no controlled
                        BatchNo = row.Batch_No
                    Else
                        BatchNo = ""
                    End If
                    If CBool(row.Location_Controlled) = True Then
                        'location controlled
                        PrimaryLocation = row.Primary_Storage_Area
                    Else
                        PrimaryLocation = -1
                    End If
                    If CBool(row.Use_Secondary_Unit) = True Then
                        'use secondary
                        Sec_Qty = row.Secondary_Qty
                    Else
                        Sec_Qty = -1
                    End If
                    If CBool(row.PdOControlled) = True Then
                        'PdO controlled
                        PrimaryPdO = row.Primary_PdO
                    Else
                        PrimaryPdO = -1
                    End If
                    If CBool(row.Expiry_Date_Controlled) = True Then
                        'expiry date controlled
                        ExpiryDate = row.Expiry_Date
                    Else
                        ExpiryDate = "1/1/1950"
                    End If

                    If vTipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust Or _
                    vTipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_BackProcess Or _
                    vTipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Process Or _
                    vTipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase Or _
                    vTipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn Then
                        Me.Koreksi4HapusIn(vTanggalTrans, ObjProd, row.Primary_Product, DTA1, row.Qty, _
                        Sec_Qty, PrimaryLocation, BatchNo, PrimaryPdO, ExpiryDate)
                    Else
                        Me.Koreksi4HapusOut(vTanggalTrans, ObjProd, row.Primary_Product, DTA1, row.Qty, _
                        Sec_Qty, PrimaryLocation, BatchNo, ExpiryDate)
                    End If
                Next
            End If


            'hapus data stock card
            JenisSC = Me.GetTipeStockCard(vTipeTransaksi)
            ObjSC.Hapus(vPrimaryTransaksi, JenisSC)
            'hapus data item
            DTA2.DeleteQuery(vPrimaryTransaksi)
            'set order balance untuk stock in purchase, untuk stock out purchase return tidak 
            'boleh karena sudah di set di transaksi purchase return
            'If vTipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase Then
            'If vSetPOB = True Then
            'ObjPOItem.SetOrderBalance(vListPrimaryPO)
            'End If
            'End If
            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
            TBL.Dispose()
            ObjSC = Nothing
            ObjProd = Nothing
            'ObjPOItem = Nothing
            'ObjStock = Nothing
            'vListPrimaryPO = Nothing
            ObjStockInOut = Nothing
        End Try
    End Function

    ''' <summary>
    ''' mengisikan TBL yang diberikan untuk mendata jumlah item yang distock out untuk primary transaksi dan 
    ''' primary product tertentu
    ''' </summary>
    ''' <param name="TBL"></param>
    ''' <param name="vPrimaryProduct"></param>
    ''' <param name="vPrimaryTransaksi"></param>
    ''' <remarks></remarks>
    Public Sub IsikanTBLBatchNo(ByVal TBL As dseTransaksiStockInOut.tbltransaksi_stockinout_item1DataTable, _
    ByVal vPrimaryProduct As Double, ByVal vPrimaryTransaksi As Double)
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinout_item1TableAdapter

        Try
            DTA.Fill(TBL, vPrimaryTransaksi, vPrimaryProduct)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsikanTBLLocation(ByVal TBL As dseTransaksiStockInOut.tbltransaksi_stockinout_item2DataTable, _
    ByVal vPrimaryProduct As Double, ByVal vPrimaryTransaksi As Double, ByVal vBatchNo As String)
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinout_item2TableAdapter

        Try
            DTA.Fill(TBL, vPrimaryTransaksi, vPrimaryProduct, vBatchNo)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsikanTBLSecondaryUnit(ByVal TBL As dseTransaksiStockInOut.tbltransaksi_stockinout_item3DataTable, _
    ByVal vPrimaryProduct As Double, ByVal vPrimaryTransaksi As Double, ByVal vBatchNo As String, ByVal vPrimaryLocation As Double)
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinout_item3TableAdapter

        Try
            DTA.Fill(TBL, vPrimaryTransaksi, vPrimaryProduct, vPrimaryLocation, vBatchNo)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function Koreksi4SimpanOut(ByVal vTanggalTrans As Date, ByVal ObjProduct As Core_Module.clsMasterProduct, ByVal vPrimaryProduct As _
    Double, ByVal DTA As Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter, ByVal Qty As Decimal, _
    Optional ByVal vSecondaryQty As Double = -1, Optional ByVal vPrimaryLocation As Double = -1, Optional ByVal _
    vBatchNo As String = "", Optional ByVal vExpiryDate As Date = _
    #1/1/1950#) As Boolean
        Dim JmlLama As Decimal = 0
        Dim JmlBaru As Decimal = 0

        Try
            'If Me.ApakahTransferOtomatis(vTanggalTrans) = True Then
            JmlLama = ObjProduct.GetJml2(vPrimaryProduct, DTA, vSecondaryQty, _
                        vPrimaryLocation, vBatchNo)
            If JmlLama <> 0 Then
                JmlBaru = JmlLama - Qty
                If JmlBaru = 0 Then
                    'hapus data lama
                    ObjProduct.Hapus(vPrimaryProduct, DTA, vSecondaryQty, vPrimaryLocation, _
                    vBatchNo)
                Else
                    'update data jumlah
                    ObjProduct.SetJml(JmlBaru, vPrimaryProduct, DTA, _
                    vSecondaryQty, vPrimaryLocation, vBatchNo)
                End If
            Else
                'masukkan data baru: tambah
                ObjProduct.SimpanQtyBaru(Qty * -1, vPrimaryProduct, DTA, vSecondaryQty, _
                vPrimaryLocation, vBatchNo, vExpiryDate)
            End If
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function Koreksi4HapusOut(ByVal vTanggalTrans As Date, ByVal ObjProduct As Core_Module.clsMasterProduct, ByVal vPrimaryProduct As _
    Double, ByVal DTA As Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter, ByVal Qty As Decimal, _
    Optional ByVal vSecondaryQty As Double = -1, Optional ByVal vPrimaryLocation As Double = -1, Optional ByVal _
    vBatchNo As String = "", Optional ByVal vExpiryDate As Date = _
    #1/1/1950#) As Boolean
        Dim JmlLama As Decimal = 0
        Dim JmlBaru As Decimal = 0
        Try
            'If Me.ApakahTransferOtomatis(vTanggalTrans) = True Then
            JmlLama = ObjProduct.GetJml2(vPrimaryProduct, DTA, vSecondaryQty, _
                       vPrimaryLocation, vBatchNo)
            If JmlLama <> 0 Then
                JmlBaru = JmlLama + Qty
                If JmlBaru = 0 Then
                    'hapus data lama
                    ObjProduct.Hapus(vPrimaryProduct, DTA, vSecondaryQty, vPrimaryLocation, _
                    vBatchNo)
                Else
                    'update data jumlah
                    ObjProduct.SetJml(JmlBaru, vPrimaryProduct, DTA, _
                    vSecondaryQty, vPrimaryLocation, vBatchNo)
                End If
            Else
                'masukkan data baru: tambah
                ObjProduct.SimpanQtyBaru(Qty, vPrimaryProduct, DTA, vSecondaryQty, _
                vPrimaryLocation, vBatchNo, vExpiryDate)
            End If
            ' End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function Koreksi4SimpanIn(ByVal vTanggalTrans As Date, ByVal ObjProduct As Core_Module.clsMasterProduct, ByVal vPrimaryProduct As _
    Double, ByVal DTA As Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter, ByVal Qty As Decimal, _
    Optional ByVal vSecondaryQty As Double = -1, Optional ByVal vPrimaryLocation As Double = -1, Optional ByVal _
    vBatchNo As String = "", Optional ByVal vPrimaryPdO As Double = -1, Optional ByVal vExpiryDate As Date = _
    #1/1/1950#) As Boolean
        Dim JmlLama As Decimal = 0
        Dim JmlBaru As Decimal = 0
        Try
            'If Me.ApakahTransferOtomatis(vTanggalTrans) = True Then
            JmlLama = ObjProduct.GetJml2(vPrimaryProduct, DTA, vSecondaryQty, _
                       vPrimaryLocation, vBatchNo)
            If JmlLama <> 0 Then
                JmlBaru = JmlLama + Qty
                If JmlBaru = 0 Then
                    'hapus data lama
                    ObjProduct.Hapus(vPrimaryProduct, DTA, vSecondaryQty, vPrimaryLocation, _
                    vBatchNo)
                Else
                    'update data jumlah
                    ObjProduct.SetJml(JmlBaru, vPrimaryProduct, DTA, _
                    vSecondaryQty, vPrimaryLocation, vBatchNo)
                End If
            Else
                'masukkan data baru: tambah
                ObjProduct.SimpanQtyBaru(Qty, vPrimaryProduct, DTA, vSecondaryQty, _
                vPrimaryLocation, vBatchNo, vExpiryDate)
            End If
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function Koreksi4HapusIn(ByVal vTanggalTrans As Date, ByVal ObjProduct As Core_Module.clsMasterProduct, ByVal vPrimaryProduct As _
                Double, ByVal DTA As Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter, ByVal Qty As Decimal, _
                Optional ByVal vSecondaryQty As Double = -1, Optional ByVal vPrimaryLocation As Double = -1, Optional ByVal _
                vBatchNo As String = "", Optional ByVal vPrimaryPdO As Double = -1, Optional ByVal vExpiryDate As Date = _
                #1/1/1950#) As Boolean
        Dim JmlLama As Decimal = 0
        Dim JmlBaru As Decimal = 0
        Try
            'If Me.ApakahTransferOtomatis(vTanggalTrans) = True Then
            JmlLama = ObjProduct.GetJml2(vPrimaryProduct, DTA, vSecondaryQty, _
                       vPrimaryLocation, vBatchNo)
            If JmlLama <> 0 Then
                JmlBaru = JmlLama - Qty
                If JmlBaru = 0 Then
                    'hapus data lama
                    ObjProduct.Hapus(vPrimaryProduct, DTA, vSecondaryQty, vPrimaryLocation, _
                    vBatchNo)
                Else
                    'update data jumlah
                    ObjProduct.SetJml(JmlBaru, vPrimaryProduct, DTA, _
                    vSecondaryQty, vPrimaryLocation, vBatchNo)
                End If
            Else
                'masukkan data baru: tambah
                ObjProduct.SimpanQtyBaru(Qty * -1, vPrimaryProduct, DTA, vSecondaryQty, _
                vPrimaryLocation, vBatchNo, vExpiryDate)
            End If
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetJmlReturn_Default(ByVal vPrimarySO As Double, ByVal vPrimaryProduct As Double, _
   ByVal vPrimaryDO As Double, ByVal vPrimaryTransaksi As Double, ByVal vBatchNo As String, _
   ByVal vSecQty As Double) As Decimal
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinout_itemTableAdapter

        Try
            Return DTA.ScalarQuery_GetJml4_Default(vPrimaryTransaksi, vPrimaryProduct, _
            vBatchNo, vPrimarySO, vSecQty, vPrimaryDO).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlReturn_Multi(ByVal vPrimarySO As Double, ByVal vPrimaryProduct As Double, _
    ByVal vPrimaryDO As Double, ByVal vPrimaryTransaksi As Double, ByVal vBatchNo As String, _
    ByVal vSecQty As Double) As Decimal
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinout_itemTableAdapter

        Try
            Return DTA.ScalarQuery_GetJml4_Multi(vPrimaryTransaksi, vPrimaryProduct, _
            vBatchNo, vPrimarySO, vSecQty, vPrimaryDO).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlReturn_Default(ByVal vPrimarySO As Double, ByVal vPrimaryProduct As Double, ByVal vPrimaryDO As Double, _
    ByVal vBatchNo As String, ByVal vSecQty As Double) As Decimal
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinout_itemTableAdapter
        Dim TipeTransaksi As clsTransaksiStockInOut.EnumTipeStockInOut = _
        clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn

        Try
            Return DTA.ScalarQuery_GetJml3_Default(vPrimaryProduct, _
            vBatchNo, vPrimarySO, vSecQty, TipeTransaksi, vPrimaryDO).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlReturn_Multi(ByVal vPrimarySO As Double, ByVal vPrimaryProduct As Double, ByVal vPrimaryDO As Double, _
    ByVal vBatchNo As String, ByVal vSecQty As Double) As Decimal
        Dim DTA As New dseTransaksiStockInOutTableAdapters.tbltransaksi_stockinout_itemTableAdapter
        Dim TipeTransaksi As clsTransaksiStockInOut.EnumTipeStockInOut = _
        clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn

        Try
            Return DTA.ScalarQuery_GetJml3_Multi(vPrimaryProduct, _
            vBatchNo, vPrimarySO, vSecQty, TipeTransaksi, vPrimaryDO).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
