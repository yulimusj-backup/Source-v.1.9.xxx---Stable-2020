Option Explicit On
Option Strict Off

Public Class clsTransaksiStockBundlingItem
    Public Function ApakahHapusMengakibatkanMinus(ByVal vPrimaryTranskasi As Double) As Boolean
        Dim DTA As New dseTransaksiStockBundlingTableAdapters.tbltransaksi_stockbundling_item_afterTableAdapter
        Dim TBL As New dseTransaksiStockBundling.tbltransaksi_stockbundling_item_afterDataTable
        Dim row As dseTransaksiStockBundling.tbltransaksi_stockbundling_item_afterRow
        Dim JmlLama As Double = 0
        Dim JmlBaru As Double = 0
        Dim ObjProduct As New Core_Module.clsMasterProduct
        Dim DTA1 As New Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter
        Dim DTAB4 As New dseTransaksiStockBundlingTableAdapters.tbltransaksi_stockbundling_item_beforeTableAdapter
        Dim TBLB4 As New dseTransaksiStockBundling.tbltransaksi_stockbundling_item_beforeDataTable
        Dim RowB4 As dseTransaksiStockBundling.tbltransaksi_stockbundling_item_beforeRow
        Dim JmlB4 As Decimal = 0

        Try
            DTAB4.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Connection.ConnectionString = DTAB4.Connection.ConnectionString
            Core_Module.SetConstring_DTAProductQty(DTA1)

            'yang mungkin menjadi minus ketika hapus data transaksi stock bundling adalah item after karena
            'item ini berkurang ketika hapus transaksi, sedangkan item before bertambah ketika hapus 
            'transaksi. item before berpengaruh ketika hapus transaksi karena ada kemungkinan nilai item
            'before sama dengan nilai item after (batch, location, expiry)
            DTA.Fill(TBL, vPrimaryTranskasi)
            DTAB4.Fill(TBLB4, vPrimaryTranskasi)

            For Each row In TBL.Rows
                JmlLama = ObjProduct.GetJml2(row.Primary_Product, DTA1, row.Secondary_Qty, _
                row.Primary_Storage_Area, row.Batch_No)
                TBLB4.DefaultView.RowFilter = "[Primary Product] = " + row.Primary_Product.ToString + _
                " AND [Batch No] = '" + row.Batch_No + "' AND [Primary Storage Area] = " + row.Primary_Storage_Area.ToString
                JmlB4 = 0
                If TBLB4.DefaultView.Count > 0 Then
                    For Each DV As Data.DataRowView In TBLB4.DefaultView
                        RowB4 = DV.Row
                        JmlB4 += RowB4.Qty
                    Next
                End If

                'If JmlLama <> 0 Then
                JmlBaru = JmlLama - row.Qty + JmlB4
                If JmlBaru < 0 Then
                    'jumlah setelah hapus transaksi akan menjadi minus
                    MsgBox("Deleting this data will cause negatif qty for: " + _
                    ObjProduct.GetProductCode(row.Primary_Product) + ".", MsgBoxStyle.Information, "Information")
                    Return True
                End If
                'Else
                ''tidak ada data digudang, data akan menjadi minus jika item dihapuskan
                'MsgBox("Deleting this data will cause negatif qty for: " + _
                'ObjProduct.GetProductCode(row.Primary_Product) + ".", MsgBoxStyle.Information, "Information")
                'Return True
                'End If
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
            DTAB4.Dispose()
            TBLB4.Dispose()
        End Try
    End Function

    Public Function GetJmlB4(ByVal vPrimaryTransaksi As Double, ByVal vPrimaryCodeNo As Double, ByVal Secondary_Qty As _
    Double, ByVal vPrimaryLocation As Double, ByVal vBatchNo As String, ByVal ExpiryDate As Date) As Decimal
        Dim DTA As New dseTransaksiStockBundlingTableAdapters.tbltransaksi_stockbundling_item_beforeTableAdapter
        Try

            If DTA.ScalarQuery_GetJml(vPrimaryTransaksi, vPrimaryCodeNo, Secondary_Qty, vBatchNo, _
            vPrimaryLocation).HasValue = True Then
                Return DTA.ScalarQuery_GetJml(vPrimaryTransaksi, vPrimaryCodeNo, Secondary_Qty, vBatchNo, _
                vPrimaryLocation).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlAfter(ByVal vPrimaryTransaksi As Double, ByVal vPrimaryCodeNo As Double, ByVal Secondary_Qty As _
    Double, ByVal vPrimaryLocation As Double, ByVal vBatchNo As String, ByVal ExpiryDate As Date) As Decimal
        Dim DTA As New dseTransaksiStockBundlingTableAdapters.tbltransaksi_stockbundling_item_afterTableAdapter
        Try

            If DTA.ScalarQuery_GetJml(vPrimaryTransaksi, vPrimaryCodeNo, Secondary_Qty, vBatchNo, _
            vPrimaryLocation).HasValue = True Then
                Return DTA.ScalarQuery_GetJml(vPrimaryTransaksi, vPrimaryCodeNo, Secondary_Qty, vBatchNo, _
                vPrimaryLocation).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetExpiryDateAfter(ByVal vPrimaryProduct As Double, ByVal vBatchNo As String) As Date
        Dim DTA As New dseTransaksiStockBundlingTableAdapters.tbltransaksi_stockbundling_item_afterTableAdapter

        Try
            If DTA.ScalarQuery_GetExpiryDate(vPrimaryProduct, vBatchNo).HasValue = True Then
                Return DTA.ScalarQuery_GetExpiryDate(vPrimaryProduct, vBatchNo).Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDGV(ByVal vDGVItemB4 As DataGridView, ByVal vDGVItemAfter As DataGridView, _
    ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTAB4 As New dseTransaksiStockBundlingTableAdapters.transaksi_stockbundling_itemb4_viewTableAdapter
        Dim TBLB4 As New dseTransaksiStockBundling.transaksi_stockbundling_itemb4_viewDataTable
        Dim rowB4 As dseTransaksiStockBundling.transaksi_stockbundling_itemb4_viewRow
        Dim DTAAfter As New dseTransaksiStockBundlingTableAdapters.transaksi_stockbundling_itemafter_viewTableAdapter
        Dim TBLAfter As New dseTransaksiStockBundling.transaksi_stockbundling_itemafter_viewDataTable
        Dim rowAfter As dseTransaksiStockBundling.transaksi_stockbundling_itemafter_viewRow
        Dim Sec_Units, Location, Warehouse As String
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            'kosongkan dgv
            vDGVItemB4.Rows.Clear()
            vDGVItemAfter.Rows.Clear()
            'isikan tabel
            If DTAB4.Fill(TBLB4, vPrimaryTransaksi) > 0 Then
                For Each rowB4 In TBLB4.Rows
                    If rowB4.Primary_Storage_Area = -1 Then
                        Location = ""
                        Warehouse = ""
                    Else
                        Location = rowB4.StorageAreaCode
                        Warehouse = rowB4.Warehouse_Code
                    End If
                    If rowB4.Primary_Secondary_Units <> -1 Then
                        Sec_Units = rowB4.Secondary_Units
                    Else
                        Sec_Units = ""
                    End If

                    vDGVItemB4.Rows.Add(vDGVItemB4.Rows.Count + 1, rowB4.Product_Code, rowB4.Description, _
                    ObjCD.DigitGrouping(rowB4.Qty, rowB4._Decimal), rowB4.Unit_Code, rowB4.Secondary_Qty, Sec_Units, rowB4.Batch_No, rowB4.Expiry_Date, _
                    Location, Warehouse, rowB4.Remarks, rowB4.Primary_Product, _
                    rowB4.Primary_Storage_Area, rowB4.Primary_Units, rowB4.Primary_Secondary_Units)
                Next
            End If
            If DTAAfter.Fill(TBLAfter, vPrimaryTransaksi) > 0 Then
                For Each rowAfter In TBLAfter.Rows
                    If rowAfter.Primary_Storage_Area = -1 Then
                        Location = ""
                        Warehouse = ""
                    Else
                        Location = rowAfter.StorageAreaCode
                        Warehouse = rowAfter.Warehouse_Code
                    End If
                    If rowAfter.Primary_Secondary_Units <> -1 Then
                        Sec_Units = rowAfter.SecondaryUnits
                    Else
                        Sec_Units = ""
                    End If

                    vDGVItemAfter.Rows.Add(vDGVItemAfter.Rows.Count + 1, rowAfter.Product_Code, rowAfter.Description, _
                    ObjCD.DigitGrouping(rowAfter.Qty, rowAfter._decimal), rowAfter.Unit_Code, rowAfter.Secondary_Qty, Sec_Units, rowAfter.Batch_No, rowAfter.Expiry_Date, _
                    Location, Warehouse, rowAfter.Remarks, rowAfter.Primary_Product, _
                    rowAfter.Primary_Storage_Area, rowAfter.Primary_Units, rowAfter.Primary_Secondary_Units, _
                    ObjCD.DigitGrouping(rowAfter.Faktor_Pemberat, 2))
                Next
            End If
            'set nomer
            'ObjCD.SetNumber(vDGVItemB4)
            'ObjCD.SetNumber(vDGVItemAfter)
            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAB4.Dispose()
            TBLB4 = Nothing
            ObjCD = Nothing
        End Try
    End Function

    Public Function Simpan(ByVal vdgvItemB4 As DataGridView, ByVal vdgvItemAfter As DataGridView, _
    ByVal vTanggalTransaksi As Date, ByVal vPrimaryTransaksi As Double, ByVal vNoTransaksi As String, _
    Optional ByVal vKoreksiStock As Boolean = True) _
    As Boolean
        Dim row As DataGridViewRow
        Dim DTAB4 As New dseTransaksiStockBundlingTableAdapters.tbltransaksi_stockbundling_item_beforeTableAdapter
        Dim DTAAfter As New dseTransaksiStockBundlingTableAdapters.tbltransaksi_stockbundling_item_afterTableAdapter
        Dim ObjProd As New Core_Module.clsMasterProduct
        Dim ObjSC As New ClsTransaksiStockCard
        Dim DTA1 As New Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter

        Dim InOut As ClsTransaksiStockCard.EnumJenisStockCard = ClsTransaksiStockCard.EnumJenisStockCard.Stock_In
        Dim JenisSC As ClsTransaksiStockCard.EnumJenisTransaksiAsal = ClsTransaksiStockCard.EnumJenisTransaksiAsal.Stock_Bundling
        Dim Qty_In, Qty_Out As Double
        Dim DescSC As String = "Stock bundling"
        Dim Sec_Qty, PrimaryLocation As Double
        Dim BatchNo As String = ""
        Dim ExpiryDate As Date = #1/1/1950#

        Try
            'simpan data item b4
            InOut = ClsTransaksiStockCard.EnumJenisStockCard.Stock_Out
            For Each row In vdgvItemB4.Rows
                BatchNo = row.Cells(7).Value
                PrimaryLocation = row.Cells(13).Value
                If IsNumeric(row.Cells(5).Value) = True Then
                    Sec_Qty = row.Cells(5).Value
                Else
                    Sec_Qty = "-1"
                End If
                If IsDate(row.Cells(8).Value) = False Then
                    ExpiryDate = #1/1/1950#
                Else
                    ExpiryDate = CDate(row.Cells(8).Value)
                End If
                'simpan ke basis data item
                DTAB4.Insert(vPrimaryTransaksi, row.Cells(12).Value, row.Cells(3).Value, row.Cells(14).Value, _
                Sec_Qty, row.Cells(15).Value, BatchNo, ExpiryDate, PrimaryLocation, row.Cells(11).Value)

                If vKoreksiStock = True Then
                    'lakukan koreksi di qty detail
                    Me.Koreksi4SimpanOut(ObjProd, row.Cells(12).Value, DTA1, row.Cells(3).Value, _
                    Sec_Qty, PrimaryLocation, BatchNo, ExpiryDate)
                End If

                'simpan data stock card
                If InOut = ClsTransaksiStockCard.EnumJenisStockCard.Stock_In Then
                    Qty_In = row.Cells(3).Value
                    Qty_Out = 0
                Else
                    Qty_In = 0
                    Qty_Out = row.Cells(3).Value
                End If
                ObjSC.Simpan(vPrimaryTransaksi, JenisSC, vNoTransaksi, row.Cells(12).Value, _
                Qty_In, Qty_Out, DescSC, InOut, vTanggalTransaksi, BatchNo, ExpiryDate, PrimaryLocation, row.Cells(11).Value)
            Next

            'simpan data item after
            InOut = ClsTransaksiStockCard.EnumJenisStockCard.Stock_In
            For Each row In vdgvItemAfter.Rows
                BatchNo = row.Cells(7).Value
                PrimaryLocation = row.Cells(13).Value
                If IsNumeric(row.Cells(5).Value) = True Then
                    Sec_Qty = row.Cells(5).Value
                Else
                    Sec_Qty = "-1"
                End If
                If IsDate(row.Cells(8).Value) = False Then
                    ExpiryDate = #1/1/1950#
                Else
                    ExpiryDate = CDate(row.Cells(8).Value)
                End If
                'simpan ke basis data item
                DTAAfter.Insert(vPrimaryTransaksi, row.Cells(12).Value, row.Cells(3).Value, row.Cells(14).Value, _
                Sec_Qty, row.Cells(15).Value, BatchNo, ExpiryDate, PrimaryLocation, row.Cells(11).Value, _
                row.Cells(16).Value)

                If vKoreksiStock = True Then
                    'lakukan koreksi di qty detail
                    Me.Koreksi4SimpanIn(ObjProd, row.Cells(12).Value, DTA1, row.Cells(3).Value, _
                    Sec_Qty, PrimaryLocation, BatchNo, -1, ExpiryDate)
                End If

                'simpan data stock card
                If InOut = ClsTransaksiStockCard.EnumJenisStockCard.Stock_In Then
                    Qty_In = row.Cells(3).Value
                    Qty_Out = 0
                Else
                    Qty_In = 0
                    Qty_Out = row.Cells(3).Value
                End If
                ObjSC.Simpan(vPrimaryTransaksi, JenisSC, vNoTransaksi, row.Cells(12).Value, _
                Qty_In, Qty_Out, DescSC, InOut, vTanggalTransaksi, BatchNo, ExpiryDate, PrimaryLocation, row.Cells(11).Value)
            Next

            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAB4.Dispose()
            DTAAfter.Dispose()
            DTA1.Dispose()
            ObjProd = Nothing
            ObjSC = Nothing
        End Try
    End Function

    Public Function Hapus(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTAB4 As New dseTransaksiStockBundlingTableAdapters.transaksi_stockbundling_itemb4_viewTableAdapter
        Dim TBLB4 As New dseTransaksiStockBundling.transaksi_stockbundling_itemb4_viewDataTable
        Dim rowB4 As dseTransaksiStockBundling.transaksi_stockbundling_itemb4_viewRow
        Dim DTAAfter As New dseTransaksiStockBundlingTableAdapters.transaksi_stockbundling_itemafter_viewTableAdapter
        Dim TBLAfter As New dseTransaksiStockBundling.transaksi_stockbundling_itemafter_viewDataTable
        Dim rowAfter As dseTransaksiStockBundling.transaksi_stockbundling_itemafter_viewRow

        Dim ObjSC As New ClsTransaksiStockCard
        Dim DTA1 As New Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter
        Dim DTA2 As New dseTransaksiStockBundlingTableAdapters.tbltransaksi_stockbundling_item_beforeTableAdapter
        Dim DTA3 As New dseTransaksiStockBundlingTableAdapters.tbltransaksi_stockbundling_item_afterTableAdapter
        Dim ObjProd As New Core_Module.clsMasterProduct
        Dim Sec_Qty, PrimaryLocation As Double
        Dim BatchNo As String = ""
        Dim ExpiryDate As Date = #1/1/1950#

        Try
            'load data item lama
            DTAB4.Fill(TBLB4, vPrimaryTransaksi)
            DTAAfter.Fill(TBLAfter, vPrimaryTransaksi)

            'lakukan koreksi stock di qty detail, untuk item B4
            For Each rowB4 In TBLB4.Rows
                'lakukan koreksi di qty detail
                If CBool(rowB4.Batch_No_Controlled) = True Then
                    'batch no controlled
                    BatchNo = rowB4.Batch_No
                Else
                    BatchNo = ""
                End If
                If CBool(rowB4.Location_Controlled) = True Then
                    'location controlled
                    PrimaryLocation = rowB4.Primary_Storage_Area
                Else
                    PrimaryLocation = -1
                End If
                If CBool(rowB4.Use_Secondary_Unit) = True Then
                    'use secondary
                    Sec_Qty = rowB4.Secondary_Qty
                Else
                    Sec_Qty = -1
                End If
                If CBool(rowB4.Expiry_Date_Controlled) = True Then
                    'expiry date controlled
                    ExpiryDate = rowB4.Expiry_Date
                Else
                    ExpiryDate = "1/1/1950"
                End If

                Me.Koreksi4HapusOut(ObjProd, rowB4.Primary_Product, DTA1, rowB4.Qty, _
                Sec_Qty, PrimaryLocation, BatchNo, ExpiryDate)
            Next

            'lakukan koreksi stock di qty detail, untuk item After
            For Each rowAfter In TBLAfter.Rows
                'lakukan koreksi di qty detail
                If CBool(rowAfter.Batch_No_Controlled) = True Then
                    'batch no controlled
                    BatchNo = rowAfter.Batch_No
                Else
                    BatchNo = ""
                End If
                If CBool(rowAfter.Location_Controlled) = True Then
                    'location controlled
                    PrimaryLocation = rowAfter.Primary_Storage_Area
                Else
                    PrimaryLocation = -1
                End If
                If CBool(rowAfter.Use_Secondary_Unit) = True Then
                    'use secondary
                    Sec_Qty = rowAfter.Secondary_Qty
                Else
                    Sec_Qty = -1
                End If
                If CBool(rowAfter.Expiry_Date_Controlled) = True Then
                    'expiry date controlled
                    ExpiryDate = rowAfter.Expiry_Date
                Else
                    ExpiryDate = "1/1/1950"
                End If

                Me.Koreksi4HapusIn(ObjProd, rowAfter.Primary_Product, DTA1, rowAfter.Qty, _
                Sec_Qty, PrimaryLocation, BatchNo, -1, ExpiryDate)
            Next

            'hapus data stock card
            ObjSC.Hapus(vPrimaryTransaksi, ClsTransaksiStockCard.EnumJenisTransaksiAsal.Stock_Bundling)
            'hapus data item
            DTA2.Delete(vPrimaryTransaksi)
            DTA3.Delete(vPrimaryTransaksi)

            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAB4.Dispose()
            DTAAfter.Dispose()
            DTA1.Dispose()
            DTA2.Dispose()
            DTA3.Dispose()
            TBLB4.Dispose()
            TBLAfter.Dispose()
            ObjSC = Nothing
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
    Public Sub IsikanTBLBatchNo(ByVal TBL As dseTransaksiStockBundling.tbltransaksi_stockbundling_item_before1DataTable, _
    ByVal vPrimaryProduct As Double, ByVal vPrimaryTransaksi As Double)
        Dim DTA As New dseTransaksiStockBundlingTableAdapters.tbltransaksi_stockbundling_item_before1TableAdapter

        Try
            DTA.Fill(TBL, vPrimaryTransaksi, vPrimaryProduct)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsikanTBLLocation(ByVal TBL As dseTransaksiStockBundling.tbltransaksi_stockbundling_item_before2DataTable, _
    ByVal vPrimaryProduct As Double, ByVal vPrimaryTransaksi As Double, ByVal vBatchNo As String)
        Dim DTA As New dseTransaksiStockBundlingTableAdapters.tbltransaksi_stockbundling_item_before2TableAdapter

        Try
            DTA.Fill(TBL, vPrimaryTransaksi, vPrimaryProduct, vBatchNo)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsikanTBLSecondaryUnit(ByVal TBL As dseTransaksiStockBundling.tbltransaksi_stockbundling_item_before3DataTable, _
    ByVal vPrimaryProduct As Double, ByVal vPrimaryTransaksi As Double, ByVal vBatchNo As String, _
    ByVal vPrimaryLocation As Double)
        Dim DTA As New dseTransaksiStockBundlingTableAdapters.tbltransaksi_stockbundling_item_before3TableAdapter

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
