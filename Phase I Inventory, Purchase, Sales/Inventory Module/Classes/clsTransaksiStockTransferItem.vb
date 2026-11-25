Option Explicit On
Option Strict Off

Public Class clsTransaksiStockTransferItem
    Public Function ApakahHapusMengakibatkanMinus(ByVal vPrimaryTranskasi As Double) As Boolean
        Dim DTA As New dseTransaksiStockTransferTableAdapters.tbltransaksi_stocktransfer_itemTableAdapter
        Dim TBL As New dseTransaksiStockTransfer.tbltransaksi_stocktransfer_itemDataTable
        Dim row As dseTransaksiStockTransfer.tbltransaksi_stocktransfer_itemRow
        Dim JmlLama As Double = 0
        Dim JmlBaru As Double = 0
        Dim ObjProduct As New Core_Module.clsMasterProduct
        Dim DTA1 As New Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter

        Try
            DTA.FillBy_PrimaryTransaksi(TBL, vPrimaryTranskasi)
            For Each row In TBL.Rows
                JmlLama = ObjProduct.GetJml2(row.Primary_Product, DTA1, row.Secondary_Qty, _
                row.primary_destination_area, row.Batch_No)
                If JmlLama <> 0 Then
                    JmlBaru = JmlLama - Me.GetJmlDestination(TBL, row.Primary_Product, row.primary_destination_area, _
                    row.Batch_No, row.Secondary_Qty)
                    If JmlBaru < 0 Then
                        'jumlah setelah hapus transaksi akan menjadi minus
                        Return True
                    End If
                Else
                    'tidak ada data digudang, data akan menjadi minus jika item dihapuskan
                    Return True
                End If
            Next

            'kalo sampe sini return false, tidak ada data yang minus jika transaksi ini dihapuskan
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
            ObjProduct = Nothing
            TBL.Dispose()
        End Try
    End Function

    Private Function GetJmlDestination(ByVal TBL As dseTransaksiStockTransfer.tbltransaksi_stocktransfer_itemDataTable, _
    ByVal vPrimaryProduct As Double, ByVal vPrimaryDestination As Double, ByVal vBatchNo As String, _
    ByVal vSecQty As Double) As Double
        Dim row As dseTransaksiStockTransfer.tbltransaksi_stocktransfer_itemRow
        Try
            Dim Qty As Double = 0

            For Each row In TBL.Rows
                If row.primary_destination_area = vPrimaryDestination And row.Primary_Product = vPrimaryProduct And _
                row.Secondary_Qty = vSecQty And vBatchNo = row.Batch_No Then
                    Qty += row.Qty
                End If
            Next

            Return Qty
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetJmlDestination(ByVal vPrimaryTransaksi As Double, ByVal vPrimaryCodeNo As Double, ByVal Secondary_Qty As _
    Double, ByVal vPrimaryDestination As Double, ByVal vBatchNo As String, ByVal ExpiryDate As Date) As Decimal
        Dim DTA As New dseTransaksiStockTransferTableAdapters.tbltransaksi_stocktransfer_itemTableAdapter
        Try

            If DTA.ScalarQuery_GetJmlByDestination(vPrimaryTransaksi, vPrimaryCodeNo, Secondary_Qty, vBatchNo, ExpiryDate, _
            vPrimaryDestination).HasValue = True Then
                Return DTA.ScalarQuery_GetJmlByDestination(vPrimaryTransaksi, vPrimaryCodeNo, Secondary_Qty, vBatchNo, ExpiryDate, _
                vPrimaryDestination).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal vPrimaryTransaksi As Double, ByVal vPrimaryCodeNo As Double, ByVal Secondary_Qty As _
    Double, ByVal vPrimaryLocationAsal As Double, ByVal vBatchNo As String, ByVal ExpiryDate As Date) As Decimal
        Dim DTA As New dseTransaksiStockTransferTableAdapters.tbltransaksi_stocktransfer_itemTableAdapter
        Try

            If DTA.ScalarQuery_GetJml(vPrimaryTransaksi, vPrimaryCodeNo, Secondary_Qty, vBatchNo, ExpiryDate, _
            vPrimaryLocationAsal).HasValue = True Then
                Return DTA.ScalarQuery_GetJml(vPrimaryTransaksi, vPrimaryCodeNo, Secondary_Qty, vBatchNo, ExpiryDate, _
                vPrimaryLocationAsal).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDGV(ByVal vDGVItem As DataGridView, ByVal vPrimaryTransaksi As Double, _
    ByVal vApakahDariTransaksiReceivement As Boolean, ByVal vTransitStatus As clsTransaksiStockTransfer.EnumTransitStatus) As Boolean
        Dim DTA As New dseTransaksiStockTransferTableAdapters.transaksi_stocktransfer_viewTableAdapter
        Dim TBL As New dseTransaksiStockTransfer.transaksi_stocktransfer_viewDataTable
        Dim row As dseTransaksiStockTransfer.transaksi_stocktransfer_viewRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim rd As DataGridViewRow
        Dim frm As New FrmTransaksiStockTransfer

        Try
            'kosongkan dgv
            vDGVItem.Rows.Clear()
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            'isikan tabel
            If DTA.Fill(TBL, vPrimaryTransaksi) > 0 Then
                For Each row In TBL.Rows
                    vDGVItem.Rows.Add()
                    rd = vDGVItem.Rows(vDGVItem.RowCount - 1)
                    rd.Cells(frm.ColNo.Index).Value = vDGVItem.RowCount
                    rd.Cells(frm.ColCodeNo.Index).Value = row.Product_Code
                    rd.Cells(frm.ColDescription.Index).Value = row.Description
                    rd.Cells(frm.ColQty.Index).Value = ObjCD.DigitGrouping(row.Qty, row._Decimal)
                    rd.Cells(frm.ColUnit.Index).Value = row.Unit_Code
                    rd.Cells(frm.ColSecondaryQty.Index).Value = row.Secondary_Qty
                    rd.Cells(frm.ColSecondaryUnit.Index).Value = row.Secondary_Units
                    rd.Cells(frm.ColBatch.Index).Value = row.Batch_No
                    rd.Cells(frm.ColExpiredDate.Index).Value = Format(row.Expiry_Date, "dd MMM yyyy")
                    rd.Cells(frm.ColRemark.Index).Value = row.Remarks
                    rd.Cells(frm.ColPrimaryCodeNo.Index).Value = row.Primary_Product
                    rd.Cells(frm.ColPrimaryUnits.Index).Value = row.Primary_Units
                    rd.Cells(frm.ColPrimarySecondaryUnits.Index).Value = row.Primary_Secondary_Units
                    If vApakahDariTransaksiReceivement = False Then
                        'lokasi asal
                        rd.Cells(frm.ColPrimaryLocation.Index).Value = row.Primary_Storage_Area
                        rd.Cells(frm.ColLocation.Index).Value = row.Storage_Area_Code
                        rd.Cells(frm.ColWareHouse.Index).Value = row.Warehouse_Code
                        'lokasi tujuan
                        rd.Cells(frm.ColPrimaryDestination.Index).Value = row.primary_destination_area
                        rd.Cells(frm.ColDestinationLoc.Index).Value = row.Destination_Area
                        rd.Cells(frm.ColWarehouseDestination.Index).Value = row.Destination_Warehouse
                        'lokasi transit
                        rd.Cells(frm.ColPrimaryTransitLocation.Index).Value = row.primary_transit_area
                        rd.Cells(frm.ColTransitLocation.Index).Value = row.Transit_Area
                        rd.Cells(frm.ColTransitWarehouse.Index).Value = row.Transit_Warehouse
                        If vTransitStatus = clsTransaksiStockTransfer.EnumTransitStatus.Finished_Transit Then
                            rd.Cells(frm.ColQtyReceived.Index).Value = ObjCD.DigitGrouping(row.Qty, row._Decimal)
                        Else
                            rd.Cells(frm.ColQtyReceived.Index).Value = 0
                        End If
                        rd.Cells(frm.ColUnitReceived.Index).Value = row.Unit_Code
                    Else
                        'lokasi asal
                        rd.Cells(frm.ColPrimaryLocation.Index).Value = row.primary_destination_area
                        rd.Cells(frm.ColLocation.Index).Value = row.Destination_Area
                        rd.Cells(frm.ColWareHouse.Index).Value = row.Destination_Warehouse
                        'lokasi tujuan
                        rd.Cells(frm.ColPrimaryDestination.Index).Value = row.primary_transit_area 'row.primary_destination_area
                        rd.Cells(frm.ColDestinationLoc.Index).Value = row.Transit_Area 'row.Destination_Area
                        rd.Cells(frm.ColWarehouseDestination.Index).Value = row.Transit_Warehouse  'row.Destination_Warehouse

                        'lokasi transit
                        rd.Cells(frm.ColPrimaryTransitLocation.Index).Value = -1 'row.primary_transit_area
                        rd.Cells(frm.ColTransitLocation.Index).Value = "" 'row.Transit_Area
                        rd.Cells(frm.ColTransitWarehouse.Index).Value = "" 'row.Transit_Warehouse
                        rd.Cells(frm.ColQtyReceived.Index).Value = 0
                        rd.Cells(frm.ColUnitReceived.Index).Value = "" 'row.Unit_Code
                    End If
                Next
            End If
            'set nomer
            'ObjCD.SetNumber(vDGVItem)
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
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

    Public Function Simpan(ByVal vdgvItem As DataGridView, ByVal vTanggalTransaksi As Date, _
    ByVal vPrimaryTransaksi As Double, ByVal vNoTransaksi As String, Optional ByVal vSupplierName As String _
    = "", Optional ByVal vCustomerName As String = "", Optional ByVal vKoreksiStock As Boolean = True) As Boolean
        Dim row As DataGridViewRow
        Dim DTA As New dseTransaksiStockTransferTableAdapters.tbltransaksi_stocktransfer_itemTableAdapter
        Dim ObjProd As New Core_Module.clsMasterProduct
        Dim DTA1 As New Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter

        Dim Sec_Qty, PrimaryLocation, PrimaryDestination, PrimaryTransitArea As Double
        Dim BatchNo As String = ""
        Dim ExpiryDate As Date = #1/1/1950#
        Dim frm As New FrmTransaksiStockTransfer

        Try
            For Each row In vdgvItem.Rows
                BatchNo = row.Cells(frm.ColBatch.Index).Value
                PrimaryLocation = row.Cells(frm.ColPrimaryLocation.Index).Value
                PrimaryDestination = row.Cells(frm.ColPrimaryDestination.Index).Value
                If IsNumeric(row.Cells(frm.ColSecondaryQty.Index).Value) = True Then
                    Sec_Qty = row.Cells(frm.ColSecondaryQty.Index).Value
                Else
                    Sec_Qty = "-1"
                End If
                If IsDate(row.Cells(frm.ColExpiredDate.Index).Value) = False Then
                    ExpiryDate = #1/1/1950#
                Else
                    ExpiryDate = CDate(row.Cells(frm.ColExpiredDate.Index).Value)
                End If
                'primary transit area
                If String.IsNullOrEmpty(row.Cells(frm.ColPrimaryTransitLocation.Index).Value) = True Then
                    row.Cells(frm.ColPrimaryTransitLocation.Index).Value = -1
                End If
                PrimaryTransitArea = row.Cells(frm.ColPrimaryTransitLocation.Index).Value
                'simpan ke basis data item
                DTA.Insert(vPrimaryTransaksi, row.Cells(frm.ColPrimaryCodeNo.Index).Value, row.Cells(frm.ColQty.Index).Value, _
                row.Cells(frm.ColPrimaryUnits.Index).Value, Sec_Qty, row.Cells(frm.ColPrimarySecondaryUnits.Index).Value, BatchNo, _
                ExpiryDate, PrimaryLocation, PrimaryDestination, row.Cells(frm.ColRemark.Index).Value, PrimaryTransitArea)

                If vKoreksiStock = True Then
                    'lakukan koreksi di qty detail
                    Me.Koreksi4SimpanIn(ObjProd, row.Cells(frm.ColPrimaryCodeNo.Index).Value, DTA1, row.Cells(frm.ColQty.Index).Value, _
                    Sec_Qty, PrimaryDestination, BatchNo, -1, ExpiryDate)
                    Me.Koreksi4SimpanOut(ObjProd, row.Cells(frm.ColPrimaryCodeNo.Index).Value, DTA1, row.Cells(frm.ColQty.Index).Value, _
                    Sec_Qty, PrimaryLocation, BatchNo, ExpiryDate)
                End If
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
            ObjProd = Nothing
            frm.Dispose()
        End Try
    End Function

    Public Function Hapus(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA1 As New Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter
        Dim DTA2 As New dseTransaksiStockTransferTableAdapters.transaksi_stocktransfer_viewTableAdapter
        Dim DTA3 As New dseTransaksiStockTransferTableAdapters.tbltransaksi_stocktransfer_itemTableAdapter

        Dim TBL As New dseTransaksiStockTransfer.transaksi_stocktransfer_viewDataTable
        Dim row As dseTransaksiStockTransfer.transaksi_stocktransfer_viewRow
        Dim ObjProd As New Core_Module.clsMasterProduct
        Dim Sec_Qty, PrimaryLocation, PrimaryDestination As Double
        Dim BatchNo As String = ""
        Dim ExpiryDate As Date = #1/1/1950#

        Try
            'load data item lama
            DTA2.Fill(TBL, vPrimaryTransaksi)

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
                    PrimaryDestination = row.primary_destination_area
                Else
                    PrimaryLocation = -1
                    PrimaryDestination = -1
                End If
                If CBool(row.Use_Secondary_Unit) = True Then
                    'use secondary
                    Sec_Qty = row.Secondary_Qty
                Else
                    Sec_Qty = -1
                End If
                If CBool(row.Expiry_Date_Controlled) = True Then
                    'expiry date controlled
                    ExpiryDate = row.Expiry_Date
                Else
                    ExpiryDate = "1/1/1950"
                End If

                Me.Koreksi4HapusIn(ObjProd, row.Primary_Product, DTA1, row.Qty, _
                Sec_Qty, PrimaryDestination, BatchNo, -1, ExpiryDate)
                Me.Koreksi4HapusOut(ObjProd, row.Primary_Product, DTA1, row.Qty, _
                Sec_Qty, PrimaryLocation, BatchNo, ExpiryDate)
            Next

            'hapus data item
            DTA3.Delete(vPrimaryTransaksi)

            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA1.Dispose()
            DTA2.Dispose()
            DTA3.Dispose()
            TBL.Dispose()
            ObjProd = Nothing
        End Try
    End Function

    Public Function Koreksi4SimpanOut(ByVal ObjProduct As Core_Module.clsMasterProduct, ByVal vPrimaryProduct As _
    Double, ByVal DTA As Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter, ByVal Qty As Decimal, _
    Optional ByVal vSecondaryQty As Double = -1, Optional ByVal vPrimaryLocation As Double = -1, Optional ByVal _
    vBatchNo As String = "", Optional ByVal vExpiryDate As Date = _
    #1/1/1950#) As Boolean
        Dim JmlLama As Decimal = 0
        Dim JmlBaru As Decimal = 0
        Try
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
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
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
    Public Sub IsikanTBLBatchNo(ByVal TBL As dseTransaksiStockTransfer.tbltransaksi_stocktransfer_item1DataTable, _
    ByVal vPrimaryProduct As Double, ByVal vPrimaryTransaksi As Double)
        Dim DTA As New dseTransaksiStockTransferTableAdapters.tbltransaksi_stocktransfer_item1TableAdapter

        Try
            DTA.Fill(TBL, vPrimaryTransaksi, vPrimaryProduct)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsikanTBLLocation(ByVal TBL As dseTransaksiStockTransfer.DataTable2DataTable, _
    ByVal vPrimaryProduct As Double, ByVal vPrimaryTransaksi As Double, ByVal vBatchNo As String)
        Dim DTA As New dseTransaksiStockTransferTableAdapters.DataTable2TableAdapter

        Try
            DTA.Fill(TBL, vPrimaryTransaksi, vPrimaryProduct, vBatchNo)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsikanTBLSecondaryUnit(ByVal TBL As dseTransaksiStockTransfer.tbltransaksi_stocktransfer_item2DataTable, _
    ByVal vPrimaryProduct As Double, ByVal vPrimaryTransaksi As Double, ByVal vBatchNo As String, ByVal vPrimaryLocation As Double)
        Dim DTA As New dseTransaksiStockTransferTableAdapters.tbltransaksi_stocktransfer_item2TableAdapter

        Try
            DTA.Fill(TBL, vPrimaryTransaksi, vPrimaryProduct, vPrimaryLocation, vBatchNo)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function Koreksi4HapusOut(ByVal ObjProduct As Core_Module.clsMasterProduct, ByVal vPrimaryProduct As _
    Double, ByVal DTA As Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter, ByVal Qty As Decimal, _
    Optional ByVal vSecondaryQty As Double = -1, Optional ByVal vPrimaryLocation As Double = -1, Optional ByVal _
    vBatchNo As String = "", Optional ByVal vExpiryDate As Date = _
    #1/1/1950#) As Boolean
        Dim JmlLama As Decimal = 0
        Dim JmlBaru As Decimal = 0
        Try
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
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function Koreksi4SimpanIn(ByVal ObjProduct As Core_Module.clsMasterProduct, ByVal vPrimaryProduct As _
    Double, ByVal DTA As Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter, ByVal Qty As Decimal, _
    Optional ByVal vSecondaryQty As Double = -1, Optional ByVal vPrimaryLocation As Double = -1, Optional ByVal _
    vBatchNo As String = "", Optional ByVal vPrimaryPdO As Double = -1, Optional ByVal vExpiryDate As Date = _
    #1/1/1950#) As Boolean
        Dim JmlLama As Decimal = 0
        Dim JmlBaru As Decimal = 0
        Try
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
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function Koreksi4HapusIn(ByVal ObjProduct As Core_Module.clsMasterProduct, ByVal vPrimaryProduct As _
    Double, ByVal DTA As Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter, ByVal Qty As Decimal, _
    Optional ByVal vSecondaryQty As Double = -1, Optional ByVal vPrimaryLocation As Double = -1, Optional ByVal _
    vBatchNo As String = "", Optional ByVal vPrimaryPdO As Double = -1, Optional ByVal vExpiryDate As Date = _
    #1/1/1950#) As Boolean
        Dim JmlLama As Decimal = 0
        Dim JmlBaru As Decimal = 0
        Try
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
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
End Class
