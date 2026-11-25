Option Explicit On
Option Strict Off

Partial Public Class clsTransaksiStockInOutItem
    Public Function IsiDGVMaterialUsage(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double, _
  ByVal vTipeTransaksi As clsTransaksiStockInOut.EnumTipeStockInOut) As Boolean
        Dim DTA As New dseTransaksiStockInOutTableAdapters.transaksi_stockinout_item_viewTableAdapter
        Dim TBL As New dseTransaksiStockInOut.transaksi_stockinout_item_viewDataTable
        Dim row As dseTransaksiStockInOut.transaksi_stockinout_item_viewRow
        Dim NoPO, NoPDO, Sec_Units, Location, Warehouse As String
        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjUnit As New Core_Module.ClsMasterUnit
        Dim ObjProd As New Core_Module.clsMasterProduct
        Dim RowDgv As DataGridViewRow
        Dim frm As New FrmTransaksiAssetAcquisition_MaterialUsage

        Try
            'kosongkan dgv
            vDGVItem.Rows.Clear()

            'isikan data
            If vTipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Sales Or _
            vTipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn Or _
            vTipeTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust Then
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

                    vDGVItem.Rows.Add()
                    RowDgv = vDGVItem.Rows(vDGVItem.RowCount - 1)
                    RowDgv.Cells(frm.ColNo.Index).Value = ""
                    RowDgv.Cells(frm.ColPONoDONo.Index).Value = ""
                    RowDgv.Cells(frm.ColSuppDONo.Index).Value = ""
                    RowDgv.Cells(frm.ColCodeNo.Index).Value = row.Product_Code
                    RowDgv.Cells(frm.ColDescription.Index).Value = row.Product_Description
                    RowDgv.Cells(frm.ColQty.Index).Value = ObjCD.DigitGrouping(row.Qty, row._Decimal)
                    RowDgv.Cells(frm.ColUnit.Index).Value = row.Unit_Code
                    RowDgv.Cells(frm.ColSecondaryQty.Index).Value = row.Secondary_Qty
                    RowDgv.Cells(frm.ColSecondaryUnit.Index).Value = Sec_Units
                    RowDgv.Cells(frm.ColBatch.Index).Value = row.Batch_No
                    RowDgv.Cells(frm.ColExpiredDate.Index).Value = row.Expiry_Date
                    RowDgv.Cells(frm.ColLocation.Index).Value = Location
                    RowDgv.Cells(frm.ColWareHouse.Index).Value = Warehouse
                    RowDgv.Cells(frm.ColRemark.Index).Value = row.Remarks
                    RowDgv.Cells(frm.ColPrimaryPOnoDoNo.Index).Value = row.Primary_PO_Purchase
                    RowDgv.Cells(frm.ColPrimaryCodeNo.Index).Value = row.Primary_Product
                    RowDgv.Cells(frm.ColPrimaryLocation.Index).Value = row.Primary_Storage_Area
                    RowDgv.Cells(frm.ColPrimaryUnits.Index).Value = row.Primary_Units
                    RowDgv.Cells(frm.ColPrimarySecondaryUnits.Index).Value = row.Primary_Secondary_Units
                    RowDgv.Cells(frm.ColPrimaryPdO.Index).Value = row.Primary_PdO
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
            frm.Dispose()
        End Try
    End Function

    Public Function Simpan_Material_Usage(ByVal vdgvItem As DataGridView, ByVal vTypeTransaksi As _
       clsTransaksiStockInOut.EnumTipeStockInOut, ByVal vTanggalTransaksi As Date, _
       ByVal vPrimaryTransaksi As Double, ByVal vNoTransaksi As String, _
       Optional ByVal vSupplierName As String = "", Optional ByVal vCustomerName As String = "", _
       Optional ByVal vSetPOB As Boolean = True, Optional ByVal vUserID As String = "") As Boolean
        Dim row As DataGridViewRow
        Dim DTA As New dseTransaksiStockInOut_AssetTableAdapters.tbltransaksi_stockinout_itemTableAdapter
        Dim ObjProd As New Core_Module.clsMasterProduct
        Dim ObjSC As New ClsTransaksiStockCard
        Dim DTA1 As New Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter

        Dim InOut As ClsTransaksiStockCard.EnumJenisStockCard = ClsTransaksiStockCard.EnumJenisStockCard.Stock_In
        Dim JenisSC As ClsTransaksiStockCard.EnumJenisTransaksiAsal = ClsTransaksiStockCard.EnumJenisStockCard.Stock_In
        Dim Qty_In, Qty_Out As Double
        Dim DTASC As New dseTransaksiStockCardTableAdapters.tbltransaksi_stock_cardTableAdapter

        Dim DescSC As String = ""
        Dim Sec_Qty, PrimaryLocation, PrimaryPdO As Double
        Dim BatchNo As String = ""
        Dim ExpiryDate As Date = #1/1/1950#

        Dim frmMaterialUsage As New FrmTransaksiAssetAcquisition_MaterialUsage

        Try
            'cek jenis transaksi stock card
            JenisSC = Me.GetTipeStockCard(vTypeTransaksi)
            'cek in/out
            Select Case vTypeTransaksi
                Case clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust
                    InOut = ClsTransaksiStockCard.EnumJenisStockCard.Stock_Out
                    DescSC = "Stock out adjust"
                    Exit Select
            End Select

            For Each row In vdgvItem.Rows
                If String.IsNullOrEmpty(row.Cells(frmMaterialUsage.ColBatch.Index).Value.ToString) = False Then
                    BatchNo = row.Cells(frmMaterialUsage.ColBatch.Index).Value.ToString
                Else
                    BatchNo = ""
                End If
                PrimaryLocation = row.Cells(frmMaterialUsage.ColPrimaryLocation.Index).Value
                If IsNumeric(row.Cells(frmMaterialUsage.ColSecondaryQty.Index).Value) = True Then
                    Sec_Qty = row.Cells(frmMaterialUsage.ColSecondaryQty.Index).Value
                Else
                    Sec_Qty = "-1"
                End If
                PrimaryPdO = row.Cells(frmMaterialUsage.ColPrimaryPdO.Index).Value
                If IsDate(row.Cells("ColExpiredDate").Value) = False Then
                    ExpiryDate = #1/1/1950#
                Else
                    ExpiryDate = CDate(row.Cells(frmMaterialUsage.ColExpiredDate.Index).Value)
                End If

                'simpan ke basis data item
                DTA.InsertQuery_Material_Usage(vPrimaryTransaksi, row.Cells(frmMaterialUsage.ColPrimaryCodeNo.Index).Value, row.Cells(frmMaterialUsage.ColQty.Index).Value, _
                row.Cells(frmMaterialUsage.ColPrimaryUnits.Index).Value, BatchNo, ExpiryDate, PrimaryLocation, PrimaryPdO, row.Cells(frmMaterialUsage.ColRemark.Index).Value, _
                row.Cells(frmMaterialUsage.ColSuppDONo.Index).Value, Sec_Qty)

                'lakukan koreksi di qty detail
                Me.Koreksi4SimpanOut(vTanggalTransaksi, ObjProd, row.Cells(frmMaterialUsage.ColPrimaryCodeNo.Index).Value, DTA1, row.Cells(frmMaterialUsage.ColQty.Index).Value, _
                Sec_Qty, PrimaryLocation, BatchNo, ExpiryDate)


                'simpan data stock card
                If InOut = ClsTransaksiStockCard.EnumJenisStockCard.Stock_In Then
                    Qty_In = row.Cells(frmMaterialUsage.ColQty.Index).Value
                    Qty_Out = 0
                Else
                    Qty_In = 0
                    Qty_Out = row.Cells(frmMaterialUsage.ColQty.Index).Value
                End If
                ObjSC.Simpan(vPrimaryTransaksi, JenisSC, vNoTransaksi, row.Cells(frmMaterialUsage.ColPrimaryCodeNo.Index).Value, _
                Qty_In, Qty_Out, DescSC, InOut, vTanggalTransaksi, DTASC, BatchNo, ExpiryDate, PrimaryLocation, row.Cells(frmMaterialUsage.ColRemark.Index).Value)
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
            ObjProd = Nothing
            ObjSC = Nothing
        End Try
    End Function
End Class
