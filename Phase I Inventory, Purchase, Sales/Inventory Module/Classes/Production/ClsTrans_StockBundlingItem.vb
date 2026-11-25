Option Explicit On
Option Strict Off

Public Class clsTrans_StockBundlingItem
    Public Function ApakahHapusMengakibatkanMinus(ByVal vPrimaryTranskasi As Double) As Boolean
        Dim DTA As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling_item_afterTableAdapter
        Dim TBL As New DseTrans_StockBundling.tbltransaksi_stockbundling_item_afterDataTable
        Dim row As DseTrans_StockBundling.tbltransaksi_stockbundling_item_afterRow
        Dim JmlLama As Double = 0
        Dim JmlBaru As Double = 0
        Dim ObjProduct As New Core_Module.clsMasterProduct
        Dim DTA1 As New Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter

        Try
            'yang mungkin menjadi minus ketika hapus data transaksi stock bundling adalah item after karena
            'item ini berkurang ketika hapus transaksi, sedangkan item before bertambah ketika hapus 
            'transaksi
            DTA.Fill(TBL, vPrimaryTranskasi)
            For Each row In TBL.Rows
                JmlLama = ObjProduct.GetJml2(row.Primary_Product, DTA1, row.Secondary_Qty, _
                row.Primary_Storage_Area, row.Batch_No)
                If JmlLama <> 0 Then
                    JmlBaru = JmlLama - row.Qty
                    If JmlBaru < 0 Then
                        'jumlah setelah hapus transaksi akan menjadi minus
                        MsgBox("Deleting this data will cause negatif qty for: " + _
                        ObjProduct.GetProductCode(row.Primary_Product) + ".", MsgBoxStyle.Information, "Information")
                        Return True
                    End If
                Else
                    'tidak ada data digudang, data akan menjadi minus jika item dihapuskan
                    MsgBox("Deleting this data will cause negatif qty for: " + _
                    ObjProduct.GetProductCode(row.Primary_Product) + ".", MsgBoxStyle.Information, "Information")
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

    Public Function GetJmlB4(ByVal vPrimaryTransaksi As Double, ByVal vPrimaryCodeNo As Double, ByVal Secondary_Qty As _
    Double, ByVal vPrimaryLocation As Double, ByVal vBatchNo As String, ByVal ExpiryDate As Date) As Double
        Dim DTA As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling_item_beforeTableAdapter
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
    Double, ByVal vPrimaryLocation As Double, ByVal vBatchNo As String, ByVal ExpiryDate As Date) As Double
        Dim DTA As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling_item_afterTableAdapter
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
        Dim DTA As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling_item_afterTableAdapter

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
    ByVal vPrimaryTransaksi As Double, ByVal vfrm As FrmTrans_StockBundling, ByVal vPrimaryWU As Double) As Boolean
        Dim DTAB4 As New DseTrans_StockBundlingTableAdapters.transaksi_stockbundling_itemb4_view_newTableAdapter
        Dim TBLB4 As New DseTrans_StockBundling.transaksi_stockbundling_itemb4_view_newDataTable
        Dim rowB4 As DseTrans_StockBundling.transaksi_stockbundling_itemb4_view_newRow
        Dim DTAAfter As New DseTrans_StockBundlingTableAdapters.transaksi_stockbundling_itemafter_view_newTableAdapter
        Dim TBLAfter As New DseTrans_StockBundling.transaksi_stockbundling_itemafter_view_newDataTable
        Dim rowAfter As DseTrans_StockBundling.transaksi_stockbundling_itemafter_view_newRow
        Dim Sec_Units, Location, Warehouse As String
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            vDGVItemAfter.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItemAfter.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            vDGVItemB4.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            vDGVItemB4.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

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

                    vDGVItemB4.Rows.Add("0", rowB4.Product_Code, rowB4.Description, _
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

                    vDGVItemAfter.Rows.Add("0", rowAfter.Product_Code, rowAfter.Description, _
                    ObjCD.DigitGrouping(rowAfter.Qty, rowAfter._decimal), rowAfter.Unit_Code, rowAfter.Secondary_Qty, Sec_Units, rowAfter.Batch_No, rowAfter.Expiry_Date, _
                    Location, Warehouse, rowAfter.Remarks, rowAfter.Primary_Product, _
                    rowAfter.Primary_Storage_Area, rowAfter.Primary_Units, rowAfter.Primary_Secondary_Units, _
                    ObjCD.DigitGrouping(rowAfter.Faktor_Pemberat, 2))
                Next
            End If
            'set nomer
            ObjCD.SetNumber(vDGVItemB4)
            ObjCD.SetNumber(vDGVItemAfter)

            '->isi dgvitem process
            Me.isidgv_downtime(vfrm.dgvinput_downtime, vPrimaryTransaksi, vfrm)
            Me.isidgv_reject(vfrm.dgvinput_processreject, vPrimaryTransaksi, ClsTrans_StockBundling.enumtipe_reject.reject_processoutput, vfrm)
            Me.isidgv_reject(vfrm.dgvinput_rmreject, vPrimaryTransaksi, ClsTrans_StockBundling.enumtipe_reject.reject_rm, vfrm)
            Me.isidgv_machine(vfrm.dgvinput_machine, vPrimaryTransaksi, vfrm, vPrimaryWU)
            Me.isidgv_karyawan(vfrm.dgvinput_employee, vPrimaryTransaksi, vfrm, vPrimaryWU)
            Me.isidgv_setuptime(vfrm.dgvInput_SetupTime, vPrimaryTransaksi, vfrm)

            vDGVItemAfter.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVItemB4.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

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
    ByVal vTanggalTransaksi As Date, ByVal vPrimaryTransaksi As Double, ByVal vNoTransaksi As String, ByVal vfrm As FrmTrans_StockBundling) _
    As Boolean
        Dim row As DataGridViewRow
        Dim DTAB4 As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling_item_beforeTableAdapter
        Dim DTAAfter As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling_item_afterTableAdapter
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
        Dim objstockbundling As New ClsTrans_StockBundling
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

                'lakukan koreksi di qty detail
                Me.Koreksi4SimpanOut(ObjProd, row.Cells(12).Value, DTA1, row.Cells(3).Value, _
                Sec_Qty, PrimaryLocation, BatchNo, ExpiryDate)

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

                ''update rm
                'Me.update_rm_by(vfrm.txtidentifier_wo.Text, row.Cells(vfrm.ColPrimaryCodeNo.Index).Value, vfrm.TxtIdentifierJO.Text)
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

                'lakukan koreksi di qty detail
                Me.Koreksi4SimpanIn(ObjProd, row.Cells(12).Value, DTA1, row.Cells(3).Value, _
                Sec_Qty, PrimaryLocation, BatchNo, -1, ExpiryDate)

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

            '->simpan data item process
            Me.simpanitem_downtime(vfrm.dgvinput_downtime, vPrimaryTransaksi, vfrm)
            Me.simpanitem_reject(vfrm.dgvinput_processreject, vPrimaryTransaksi, ClsTrans_StockBundling.enumtipe_reject.reject_processoutput, vfrm)
            Me.simpanitem_reject(vfrm.dgvinput_rmreject, vPrimaryTransaksi, ClsTrans_StockBundling.enumtipe_reject.reject_rm, vfrm)
            Me.simpanitem_karyawan(vfrm.dgvinput_employee, vPrimaryTransaksi, vfrm)
            Me.simpanitem_machine(vfrm.dgvinput_machine, vPrimaryTransaksi, vfrm)
            Me.simpanitem_setuptime(vfrm.dgvInput_SetupTime, vPrimaryTransaksi, vfrm)

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

    Public Function Hapus(ByVal vPrimaryTransaksi As Double, ByVal vfrm As FrmTrans_StockBundling) As Boolean
        Dim DTAB4 As New DseTrans_StockBundlingTableAdapters.transaksi_stockbundling_itemb4_view_newTableAdapter
        Dim TBLB4 As New DseTrans_StockBundling.transaksi_stockbundling_itemb4_view_newDataTable
        Dim rowB4 As DseTrans_StockBundling.transaksi_stockbundling_itemb4_view_newRow
        Dim DTAAfter As New DseTrans_StockBundlingTableAdapters.transaksi_stockbundling_itemafter_view_newTableAdapter
        Dim TBLAfter As New DseTrans_StockBundling.transaksi_stockbundling_itemafter_view_newDataTable
        Dim rowAfter As DseTrans_StockBundling.transaksi_stockbundling_itemafter_view_newRow

        Dim ObjSC As New ClsTransaksiStockCard
        Dim DTA1 As New Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter
        Dim DTA2 As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling_item_beforeTableAdapter
        Dim DTA3 As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling_item_afterTableAdapter
        Dim ObjProd As New Core_Module.clsMasterProduct
        Dim Sec_Qty, PrimaryLocation As Double
        Dim BatchNo As String = ""
        Dim ExpiryDate As Date = #1/1/1950#
        Dim objstockbundling As New ClsTrans_StockBundling
        Dim list_primaryproduct As New List(Of Double)
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

                If list_primaryproduct.Contains(rowB4.Primary_Product) = False Then
                    list_primaryproduct.Add(rowB4.Primary_Product)
                End If
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

            'For Each i As Double In list_primaryproduct
            '    'update rm
            '    Me.update_rm_by(vfrm.txtidentifier_wo.Text, i, vfrm.TxtIdentifierJO.Text)
            'Next

            '-> hapus item dgvitem process
            Me.hapusitem_dgvprocess(vPrimaryTransaksi)

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
    Public Sub IsikanTBLBatchNo(ByVal TBL As DseTrans_StockBundling.tbltransaksi_stockbundling_item_before1DataTable, _
    ByVal vPrimaryProduct As Double, ByVal vPrimaryTransaksi As Double)
        Dim DTA As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling_item_before1TableAdapter

        Try
            DTA.Fill(TBL, vPrimaryTransaksi, vPrimaryProduct)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsikanTBLLocation(ByVal TBL As DseTrans_StockBundling.tbltransaksi_stockbundling_item_before2DataTable, _
    ByVal vPrimaryProduct As Double, ByVal vPrimaryTransaksi As Double, ByVal vBatchNo As String)
        Dim DTA As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling_item_before2TableAdapter

        Try
            DTA.Fill(TBL, vPrimaryTransaksi, vPrimaryProduct, vBatchNo)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsikanTBLSecondaryUnit(ByVal TBL As DseTrans_StockBundling.tbltransaksi_stockbundling_item_before3DataTable, _
    ByVal vPrimaryProduct As Double, ByVal vPrimaryTransaksi As Double, ByVal vBatchNo As String, _
    ByVal vPrimaryLocation As Double)
        Dim DTA As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling_item_before3TableAdapter

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

    '========================simpan item
    Private Function simpanitem_reject(ByVal dgv_reject As DataGridView, ByVal vprimary_bundling As Double, ByVal vtipe_reject As ClsTrans_StockBundling.enumtipe_reject, _
    ByVal frm As FrmTrans_StockBundling) As Boolean
        Dim dta As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling_item_rejectTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring

            If vtipe_reject = ClsTrans_StockBundling.enumtipe_reject.reject_processoutput Then
                For Each vr As DataGridViewRow In dgv_reject.Rows
                    With vr
                        If String.IsNullOrEmpty(.Cells(frm.colinput_processreject_codeno.Index).Value) = False Then
                            dta.Insert(vprimary_bundling, vtipe_reject, .Cells(frm.colinput_processreject_primary.Index).Value, .Cells(frm.colinput_processreject_primaryproduct.Index).Value, _
                             .Cells(frm.colinput_processreject_qty.Index).Value, .Cells(frm.colinput_processreject_remarks.Index).Value)
                        End If
                    End With
                Next
            ElseIf vtipe_reject = ClsTrans_StockBundling.enumtipe_reject.reject_rm Then
                For Each vr As DataGridViewRow In dgv_reject.Rows
                    With vr
                        If String.IsNullOrEmpty(.Cells(frm.colinput_rmreject_codeno.Index).Value) = False Then
                            dta.Insert(vprimary_bundling, vtipe_reject, .Cells(frm.colinput_rmreject_primaryreject.Index).Value, .Cells(frm.colinput_rmreject_primaryproduct.Index).Value, _
                            .Cells(frm.colinput_rmreject_qty.Index).Value, .Cells(frm.colinput_rmreject_remarks.Index).Value)
                        End If
                    End With
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Private Function simpanitem_downtime(ByVal dgv_downtime As DataGridView, ByVal vprimary_bundling As Double, ByVal vfrm As FrmTrans_StockBundling) As Boolean
        Dim dta As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling_item_downtimeTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring

            For Each vr As DataGridViewRow In dgv_downtime.Rows
                With vr
                    If String.IsNullOrEmpty(vr.Cells(vfrm.colinput_employee_code.Index).Value) = False Then
                        dta.Insert(vprimary_bundling, .Cells(vfrm.colinput_downtime_primary.Index).Value, Me.objcd.Convert_ToMinutes(.Cells(vfrm.colinput_downtime_duration.Index).Value), _
                         .Cells(vfrm.colinput_downtime_remarks.Index).Value)
                    End If
                End With
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Private Function simpanitem_setuptime(ByVal dgv_setuptime As DataGridView, ByVal vprimary_bundling As Double, ByVal vfrm As FrmTrans_StockBundling) As Boolean
        Dim dta As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling_item_setuptimeTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring

            For Each vr As DataGridViewRow In dgv_setuptime.Rows
                With vr
                    If String.IsNullOrEmpty(vr.Cells(vfrm.ColST_Code.Index).Value) = False Then
                        dta.Insert(vprimary_bundling, .Cells(vfrm.ColST_PrimarySetupTime.Index).Value, CDec(Me.objcd.Convert_ToMinutes(.Cells(vfrm.ColST_Duration.Index).Value)), _
                         .Cells(vfrm.ColST_Remarks.Index).Value)
                    End If
                End With
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Private Function simpanitem_machine(ByVal dgv_machine As DataGridView, ByVal vprimary_bundling As Double, ByVal vfrm As FrmTrans_StockBundling) As Boolean
        Dim dta As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling_item_machineTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring

            For Each vr As DataGridViewRow In dgv_machine.Rows
                With vr
                    If String.IsNullOrEmpty(vr.Cells(vfrm.colinput_machine_code.Index).Value) = False Then
                        dta.Insert(vprimary_bundling, vr.Cells(vfrm.colinput_machine_primary.Index).Value, CDbl(vr.Cells(vfrm.colinput_machine_prodindex.Index).Value), _
                          .Cells(vfrm.colinput_machine_remarks.Index).Value)
                    End If
                End With
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Private Function simpanitem_karyawan(ByVal dgv_karyawan As DataGridView, ByVal vprimary_bundling As Double, ByVal vfrm As FrmTrans_StockBundling) As Boolean
        Dim dta As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling_item_employeeTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring

            For Each vr As DataGridViewRow In dgv_karyawan.Rows
                With vr
                    If String.IsNullOrEmpty(.Cells(vfrm.colinput_employee_code.Index).Value) = False Then
                        dta.Insert(vprimary_bundling, .Cells(vfrm.colinput_employee_primary.Index).Value, .Cells(vfrm.colinput_employee_prodindex.Index).Value, _
                        .Cells(vfrm.colinput_employee_remarks.Index).Value)
                    End If
                End With
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    '====================isidgv_item
    Dim objcd As New Core_Module.ClsCekData

    Private Function isidgv_reject(ByVal dgv_reject As DataGridView, ByVal vprimary_bundling As Double, ByVal vtipe_reject As ClsTrans_StockBundling.enumtipe_reject, _
    ByVal vfrm As FrmTrans_StockBundling) As Boolean
        Dim dta As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling_item_reject1TableAdapter
        Dim tbl As New DseTrans_StockBundling.tbltransaksi_stockbundling_item_reject1DataTable
        Dim row As DseTrans_StockBundling.tbltransaksi_stockbundling_item_reject1Row
        Dim vr As DataGridViewRow
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dgv_reject.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            dgv_reject.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            If vtipe_reject = ClsTrans_StockBundling.enumtipe_reject.reject_processoutput Then
                dta.Fill(tbl, vprimary_bundling, vtipe_reject)
                vfrm.dgvinput_processreject.AllowUserToAddRows = False
                vfrm.dgvinput_processreject.Rows.Clear()
                For Each row In tbl.Rows
                    With row
                        dgv_reject.Rows.Add()
                        vr = dgv_reject.Rows(vfrm.dgvinput_processreject.Rows.Count - 1)
                        vr.Cells(vfrm.colinput_processreject_no.Index).Value = ""
                        vr.Cells(vfrm.colinput_processreject_primary.Index).Value = .primary_reject
                        vr.Cells(vfrm.colinput_processreject_primaryproduct.Index).Value = .primary_product
                        vr.Cells(vfrm.colinput_processreject_codeno.Index).Value = .code_no
                        vr.Cells(vfrm.colinput_processreject_qty.Index).Value = Me.objcd.DigitGrouping(.qty, ._Decimal)
                        vr.Cells(vfrm.colinput_processreject_rejectcode.Index).Value = .reject_code
                        vr.Cells(vfrm.colinput_processreject_remarks.Index).Value = .remarks
                    End With
                Next
                vfrm.dgvinput_processreject.AllowUserToAddRows = True
                Me.objcd.SetNumber(dgv_reject)

            ElseIf vtipe_reject = ClsTrans_StockBundling.enumtipe_reject.reject_rm Then
                dta.Fill(tbl, vprimary_bundling, vtipe_reject)

                vfrm.dgvinput_rmreject.Rows.Clear()
                vfrm.dgvinput_rmreject.AllowUserToAddRows = False
                For Each row In tbl.Rows
                    With row
                        dgv_reject.Rows.Add()
                        vr = dgv_reject.Rows(vfrm.dgvinput_rmreject.Rows.Count - 1)
                        vr.Cells(vfrm.colinput_rmreject_no.Index).Value = ""
                        vr.Cells(vfrm.colinput_rmreject_primaryproduct.Index).Value = .primary_product
                        vr.Cells(vfrm.colinput_rmreject_primaryreject.Index).Value = .primary_reject
                        vr.Cells(vfrm.colinput_rmreject_codeno.Index).Value = .code_no
                        vr.Cells(vfrm.colinput_rmreject_rejectcode.Index).Value = .reject_code
                        vr.Cells(vfrm.colinput_rmreject_qty.Index).Value = Me.objcd.DigitGrouping(.qty, ._Decimal)
                        vr.Cells(vfrm.colinput_rmreject_remarks.Index).Value = .remarks
                    End With
                Next
                vfrm.dgvinput_rmreject.AllowUserToAddRows = True
                Me.objcd.SetNumber(dgv_reject)
            End If
            dgv_reject.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function isidgv_machine(ByVal dgv_machine As DataGridView, ByVal vprimary_bundling As Double, _
    ByVal vfrm As FrmTrans_StockBundling, ByVal vPrimaryWU As Double) As Boolean
        Dim dta As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling_item_machine1TableAdapter
        Dim tbl As New DseTrans_StockBundling.tbltransaksi_stockbundling_item_machine1DataTable
        Dim row As DseTrans_StockBundling.tbltransaksi_stockbundling_item_machine1Row
        Dim vr As DataGridViewRow
        Dim obj_mesin As New Core_Module.ClsMasterMachineData
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dgv_machine.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            dgv_machine.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            dta.Fill(tbl, vprimary_bundling, vPrimaryWU)
            vfrm.dgvinput_machine.Rows.Clear()
            dgv_machine.AllowUserToAddRows = False
            For Each row In tbl.Rows
                With row
                    dgv_machine.Rows.Add()
                    vr = dgv_machine.Rows(dgv_machine.Rows.Count - 1)
                    vr.Cells(vfrm.colinput_machine_no.Index).Value = ""
                    vr.Cells(vfrm.colinput_machine_primary.Index).Value = .primary_machine
                    vr.Cells(vfrm.colinput_machine_code.Index).Value = .machine_data_code
                    vr.Cells(vfrm.colinput_machine_desc.Index).Value = obj_mesin.GetDesc(.primary_machine)
                    vr.Cells(vfrm.colinput_machine_prodindex.Index).Value = Me.objcd.DigitGrouping(.production_index, 0)
                    vr.Cells(vfrm.colinput_machine_remarks.Index).Value = .remarks
                End With
            Next
            dgv_machine.AllowUserToAddRows = True
            Me.objcd.SetNumber(dgv_machine)
            dgv_machine.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Private Function isidgv_downtime(ByVal dgv_downtime As DataGridView, ByVal vprimary_bundling As Double, _
    ByVal vfrm As FrmTrans_StockBundling) As Boolean
        Dim dta As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling_item_downtime1TableAdapter
        Dim tbl As New DseTrans_StockBundling.tbltransaksi_stockbundling_item_downtime1DataTable
        Dim row As DseTrans_StockBundling.tbltransaksi_stockbundling_item_downtime1Row
        Dim vr As DataGridViewRow
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dgv_downtime.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            dgv_downtime.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            dta.Fill(tbl, vprimary_bundling)
            vfrm.dgvinput_downtime.Rows.Clear()
            dgv_downtime.AllowUserToAddRows = False
            For Each row In tbl.Rows
                With row
                    dgv_downtime.Rows.Add()
                    vr = dgv_downtime.Rows(dgv_downtime.Rows.Count - 1)
                    vr.Cells(vfrm.colinput_downtime_no.Index).Value = ""
                    vr.Cells(vfrm.colinput_downtime_primary.Index).Value = .primary_downtime
                    vr.Cells(vfrm.colinput_downtime_code.Index).Value = .downtime_code
                    vr.Cells(vfrm.colinput_downtime_duration.Index).Value = Me.objcd.Convert_ToString(.duration)
                    vr.Cells(vfrm.colinput_downtime_remarks.Index).Value = .remarks
                End With
            Next
            dgv_downtime.AllowUserToAddRows = True
            Me.objcd.SetNumber(dgv_downtime)
            dgv_downtime.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Private Function isidgv_setuptime(ByVal dgv_setuptime As DataGridView, ByVal vprimary_bundling As Double, _
    ByVal vfrm As FrmTrans_StockBundling) As Boolean
        Dim dta As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling_item_setuptime1TableAdapter
        Dim tbl As New DseTrans_StockBundling.tbltransaksi_stockbundling_item_setuptime1DataTable
        Dim row As DseTrans_StockBundling.tbltransaksi_stockbundling_item_setuptime1Row
        Dim vr As DataGridViewRow
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dgv_setuptime.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            dgv_setuptime.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            dta.Fill(tbl, vprimary_bundling)
            vfrm.dgvInput_SetupTime.Rows.Clear()
            dgv_setuptime.AllowUserToAddRows = False
            For Each row In tbl.Rows
                With row
                    dgv_setuptime.Rows.Add()
                    vr = dgv_setuptime.Rows(dgv_setuptime.Rows.Count - 1)
                    vr.Cells(vfrm.ColST_No.Index).Value = ""
                    vr.Cells(vfrm.ColST_PrimarySetupTime.Index).Value = .primary_setuptime
                    vr.Cells(vfrm.ColST_Code.Index).Value = .Code
                    vr.Cells(vfrm.ColST_Duration.Index).Value = Me.objcd.Convert_ToString(.duration)
                    vr.Cells(vfrm.ColST_Remarks.Index).Value = .remarks
                End With
            Next
            dgv_setuptime.AllowUserToAddRows = True
            Me.objcd.SetNumber(dgv_setuptime)
            dgv_setuptime.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function isidgv_machineby_primaryworkunit(ByVal dgv_machine As DataGridView, ByVal vprimarywu As Double, ByVal vfrm As FrmTrans_StockBundling) As Boolean
        Dim dta As New DseTrans_StockBundlingTableAdapters.tblmachinebyTableAdapter
        Dim tbl As New DseTrans_StockBundling.tblmachinebyDataTable
        Dim row As DseTrans_StockBundling.tblmachinebyRow
        Dim vr As DataGridViewRow
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dgv_machine.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            dgv_machine.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            dta.Fill(tbl, vprimarywu)
            vfrm.dgvinput_machine.Rows.Clear()
            dgv_machine.AllowUserToAddRows = False
            For Each row In tbl.Rows
                With row
                    dgv_machine.Rows.Add()
                    vr = dgv_machine.Rows(dgv_machine.Rows.Count - 1)
                    vr.Cells(vfrm.colinput_machine_no.Index).Value = ""
                    vr.Cells(vfrm.colinput_machine_primary.Index).Value = .primary_machine
                    vr.Cells(vfrm.colinput_machine_code.Index).Value = .Machine_Data_Code
                    vr.Cells(vfrm.colinput_machine_desc.Index).Value = .Description
                    vr.Cells(vfrm.colinput_machine_prodindex.Index).Value = "0"
                    vr.Cells(vfrm.colinput_machine_remarks.Index).Value = .Remarks
                End With
            Next
            dgv_machine.AllowUserToAddRows = True
            Me.objcd.SetNumber(dgv_machine)
            dgv_machine.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function isidgv_karyawanby_primaryworkunit(ByVal dgv_employee As DataGridView, ByVal vprimarywu As Double, _
        ByVal vfrm As FrmTrans_StockBundling) As Boolean

        Dim dta As New DseTrans_StockBundlingTableAdapters.tblkaryawanbyTableAdapter
        Dim tbl As New DseTrans_StockBundling.tblkaryawanbyDataTable
        Dim row As DseTrans_StockBundling.tblkaryawanbyRow
        Dim vr As DataGridViewRow
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dgv_employee.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            dgv_employee.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            dta.Fill(tbl, vprimarywu)
            vfrm.dgvinput_employee.Rows.Clear()
            dgv_employee.AllowUserToAddRows = False
            For Each row In tbl.Rows
                With row
                    dgv_employee.Rows.Add()
                    vr = dgv_employee.Rows(dgv_employee.Rows.Count - 1)
                    vr.Cells(vfrm.colinput_employee_no.Index).Value = ""
                    vr.Cells(vfrm.colinput_employee_primary.Index).Value = .primary_employee
                    vr.Cells(vfrm.colinput_employee_code.Index).Value = .Code
                    vr.Cells(vfrm.colinput_employee_name.Index).Value = .Full_Name
                    vr.Cells(vfrm.colinput_employee_prodindex.Index).Value = "0"
                    vr.Cells(vfrm.colinput_employee_remarks.Index).Value = .remarks
                End With
            Next
            dgv_employee.AllowUserToAddRows = True
            Me.objcd.SetNumber(dgv_employee)
            dgv_employee.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function isidgv_karyawan(ByVal dgv_employee As DataGridView, ByVal vprimary_bundling As Double, _
    ByVal vfrm As FrmTrans_StockBundling, ByVal vPrimaryWU As Double) As Boolean

        Dim dta As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling_item_employee1TableAdapter
        Dim tbl As New DseTrans_StockBundling.tbltransaksi_stockbundling_item_employee1DataTable
        Dim row As DseTrans_StockBundling.tbltransaksi_stockbundling_item_employee1Row
        Dim vr As DataGridViewRow
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dgv_employee.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            dgv_employee.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            dta.Fill(tbl, vprimary_bundling, vPrimaryWU)
            vfrm.dgvinput_employee.Rows.Clear()
            dgv_employee.AllowUserToAddRows = False
            For Each row In tbl.Rows
                With row
                    dgv_employee.Rows.Add()
                    vr = dgv_employee.Rows(dgv_employee.Rows.Count - 1)
                    vr.Cells(vfrm.colinput_employee_no.Index).Value = ""
                    vr.Cells(vfrm.colinput_employee_primary.Index).Value = .primary_employee
                    vr.Cells(vfrm.colinput_employee_code.Index).Value = .code_karyawan
                    vr.Cells(vfrm.colinput_employee_name.Index).Value = .full_name
                    vr.Cells(vfrm.colinput_employee_prodindex.Index).Value = Me.objcd.DigitGrouping(.production_index, 0)
                    vr.Cells(vfrm.colinput_employee_remarks.Index).Value = .remarks
                End With
            Next
            dgv_employee.AllowUserToAddRows = True
            Me.objcd.SetNumber(dgv_employee)
            dgv_employee.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    '============hapus dgv_item
    Private Function hapusitem_dgvprocess(ByVal vprimary_bundling As Double) As Boolean
        Dim dta_reject As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling_item_rejectTableAdapter
        Dim dta_downtime As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling_item_downtimeTableAdapter
        Dim dta_machine As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling_item_machineTableAdapter
        Dim dta_employee As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling_item_employeeTableAdapter
        Dim dta_setuptime As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling_item_setuptimeTableAdapter
        Try
            dta_reject.Connection.ConnectionString = Core_Module.GetConstring
            dta_downtime.Connection.ConnectionString = dta_reject.Connection.ConnectionString
            dta_machine.Connection.ConnectionString = dta_reject.Connection.ConnectionString
            dta_employee.Connection.ConnectionString = dta_reject.Connection.ConnectionString

            dta_reject.Delete(vprimary_bundling, ClsTrans_StockBundling.enumtipe_reject.reject_processoutput)
            dta_reject.Delete(vprimary_bundling, ClsTrans_StockBundling.enumtipe_reject.reject_rm)
            dta_downtime.Delete(vprimary_bundling)
            dta_machine.Delete(vprimary_bundling)
            dta_employee.Delete(vprimary_bundling)
            dta_setuptime.Delete(vprimary_bundling)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            dta_reject.Dispose()
            dta_downtime.Dispose()
            dta_employee.Dispose()
            dta_machine.Dispose()
        End Try
    End Function

    Public Function apakahproductdirm_sudahada(ByVal videntifier_wo As String, ByVal vprimaryproduct_rm As Double, ByVal vIdentifierJO As String) As Boolean
        Dim dta As New DseTrans_StockBundlingTableAdapters.tbltrans_mo_item_rmTableAdapter
        Dim retval As New Nullable(Of Long)

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            retval = dta.ScalarQuery_GetJmlRM(videntifier_wo, vprimaryproduct_rm, vIdentifierJO)

            If retval.HasValue = True Then
                If retval.Value > 0 Then
                    Return True
                Else
                    Return False
                End If
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function get_qtyallocatedrm(ByVal videntifier_wo As String, ByVal vprimary_product As Double, ByVal vIdentifierJO As String) As Decimal
        Dim dta As New DseTrans_StockBundlingTableAdapters.tbltrans_mo_item_rmTableAdapter
        Dim retval As New Nullable(Of Decimal)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            retval = dta.ScalarQuery_GetAllocatedQty(videntifier_wo, vprimary_product, vIdentifierJO)
            If retval.HasValue = True Then
                Return retval.Value
            Else
                Return 0
            End If
            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function get_qtyadditionalrm(ByVal videntifier_wo As String, ByVal vprimary_product As Double, ByVal vIdentifierJO As String) As Decimal
        Dim dta As New DseTrans_StockBundlingTableAdapters.tbltrans_mo_item_rmTableAdapter
        Dim retval As New Nullable(Of Decimal)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            retval = dta.ScalarQuery_GetAdditionalQty(videntifier_wo, vprimary_product, vIdentifierJO)
            If retval.HasValue = True Then
                Return retval.Value
            Else
                Return 0
            End If
            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function get_qtytargetresult(ByVal videntifier_wo As String) As Decimal
        Dim dta As New DseTrans_StockBundlingTableAdapters.tbltrans_mo_item_woTableAdapter
        Dim retval As Nullable(Of Decimal)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            retval = dta.ScalarQuery_GetQtyTargetResult(videntifier_wo)
            If retval.HasValue = True Then
                Return retval.HasValue
            Else
                Return 0
            End If
            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function get_sumqtyrm_alltransactionperwo(ByVal videntifier_wo As String, ByVal vprimary_product As Double, ByVal vIdentifierJO As String) As Decimal
        Dim dta As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling_item_beforeTableAdapter
        Dim retval As New Nullable(Of Decimal)

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            retval = dta.ScalarQuery_sumqtyrm(vprimary_product, videntifier_wo, vIdentifierJO)
            If retval.HasValue = True Then
                If retval.Value > 0 Then
                    Return retval
                Else
                    Return 0
                End If
            End If
            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function get_qtyrmbalance(ByVal videntifier_wo As String, ByVal vprimary_product As Double, ByVal vIdentifierJO As String) As Decimal
        Dim dta As New DseTrans_StockBundlingTableAdapters.tbltrans_mo_item_rmTableAdapter
        Dim retval As New Nullable(Of Decimal)

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            retval = dta.ScalarQuery_GetQtyBalanceRM(videntifier_wo, vprimary_product, vIdentifierJO)
            If retval.HasValue = True Then
                If retval.Value > 0 Then
                    Return retval.Value
                Else
                    Return 0
                End If
            End If
            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            dta.Dispose()
        End Try
    End Function

    'Public Function update_rm_by(ByVal videntifier_wo As String, ByVal vlistprimary_product As List(Of Double), ByVal vIdentifierJO As String) As Boolean
    '    Try
    '        For Each i As Double In vlistprimary_product
    '            Me.update_rm_by(videntifier_wo, i, vIdentifierJO)
    '        Next
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Function

    ''#update ke databasenya untuk rm
    'Public Function update_rm_by(ByVal videntifier_wo As String, ByVal vprimaryproduct_rm As Double, ByVal vIdentifierJO As String) As Boolean
    '    Dim qty, blc, additional, allocated, count As Decimal
    '    Dim dta As New DseTrans_StockBundlingTableAdapters.tblgets_rmTableAdapter
    'Dim objMO As New ClsTrans_ManufacturingOrder

    '    Try
    '        'MsgBox("Consider jo dulu")

    '        dta.Connection.ConnectionString = Core_Module.GetConstring
    '        If Me.apakahproductdirm_sudahada(videntifier_wo, vprimaryproduct_rm, vIdentifierJO) = True Then
    '            qty = Me.get_sumqtyrm_alltransactionperwo(videntifier_wo, vprimaryproduct_rm, vIdentifierJO)
    '            additional = Me.get_qtyadditionalrm(videntifier_wo, vprimaryproduct_rm, vIdentifierJO)
    '            allocated = Me.get_qtyallocatedrm(videntifier_wo, vprimaryproduct_rm, vIdentifierJO)
    '            blc = Me.get_qtyrmbalance(videntifier_wo, vprimaryproduct_rm, vIdentifierJO)
    '            count = allocated + additional - qty
    '            dta.UpdateQuery_RM(qty, count, videntifier_wo, vprimaryproduct_rm, vIdentifierJO)

    'objMO.get_GPActualRM(videntifier_wo, vprimaryproduct_rm, vIdentifierJO)
    '            Return True
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '        Return False
    '    Finally
    '        dta.Dispose()
    '        objMO = Nothing
    '    End Try
    'End Function
End Class
