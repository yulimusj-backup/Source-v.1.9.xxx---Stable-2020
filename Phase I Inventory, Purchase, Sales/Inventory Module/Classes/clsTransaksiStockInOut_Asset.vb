Option Explicit On
Option Strict Off

Partial Public Class clsTransaksiStockInOut
    Private mMaterialUsageAmount As Decimal = 0
    Private ObjAsset As New Core_Module.ClsMasterAsset

    Public Property MaterialUsageAmount() As Decimal
        Get
            Return mMaterialUsageAmount
        End Get
        Set(ByVal value As Decimal)
            mMaterialUsageAmount = value
        End Set
    End Property

    Public Overridable Function SimpanBaru_Material_Usage(ByVal vdgvItem As DataGridView, ByVal vUserCode As String, _
        Optional ByVal vAskEmailNotification As Boolean = True) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim ObjAntri As New Core_Module.ClsAntrian
        Dim SW As New Stopwatch

        Try
            'ambil no antrian
            ObjAntri.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Core_Module.ClsAntrian.EnumTipeAntrian.AntrianStockInOut_DO_Invoice, _
            Me.vPrimary_Antrian, Me.mPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiStockInOut_SimpanBaru, vUserCode)


            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= ObjAntri.GetWaitTime Then
                        Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanBaru1_Material_Usage(vdgvItem, vUserCode, vAskEmailNotification)
                            'hapus data antri untuk transaksi ini
                            If Me.DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
                                MsgBox("Logic error(1), contact your provider when you see this message", _
                                MsgBoxStyle.Information, "Information")
                            End If
                            'nilai balikan
                            Return retval
                        Else
                            'kontrol interval pengecekan
                            TryCounter += 1
                            CheckInterval = ((TryCounter \ 5) + 1) * 500
                            Me.Clock_Antri.Reset()
                            Me.Clock_Antri.Start()
                            Threading.Thread.Sleep(1000)
                        End If
                    End If
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function SimpanBaru1_Material_Usage(ByVal vdgvItem As DataGridView, ByVal vUserCode As String, _
    Optional ByVal vAskEmailNotification As Boolean = True, Optional ByVal vSetPOB As Boolean = True) As Boolean
        Dim DTA As New dseTransaksiStockInOut_AssetTableAdapters.tbltransaksi_stockinoutTableAdapter
        Dim ObjBooking As New ClsTransaksiBookingInv

        Dim ObjItem As New clsTransaksiStockInOutItem
        Dim CustName, SuppName, CustCode As String

        Try
            If Me.ApakahDataBenar_MaterialUsage(True, vdgvItem) = True Then
                'insert query
                'material usage isi 0 dulu
                If DTA.InsertQuery_asset_MaterialUsage(Me.mNoTransaksi, CInt(Me.mNoTransaksi.Substring(0, 5)), Me.mTipeTransaksi, _
                   Me.mTanggalTransaksi, Me.mPrimaryPOPurchase, Me.mPrimaryDOSales, Me.mPrimaryCust, _
                   Me.mPrimarySupplier, Me.mImporDariPackingList, Me.mPrimaryGIT, Me.mRemarks, vUserCode, _
                   Now, vUserCode, Now, Now, Me.mSupplierDoNo, Me.mPrimaryAsset, Me.mInOutAdjust, Me.mMaterialUsageAmount) > 0 Then
                    'update data yang berubah setelah insert query
                    Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mNoTransaksi).Value
                    Me.mNoUrut = CInt(Me.mNoTransaksi.Substring(0, 5))
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode
                    'simpan data item baru
                    If Me.mPrimaryCust <> -1 Then
                        Dim ObjCust As New Core_Module.clsMasterCustomer
                        CustName = ObjCust.GetCustName(Me.mPrimaryCust)
                        CustCode = ObjCust.GetCustomerCode(Me.mPrimaryCust)
                        ObjCust = Nothing
                    Else
                        CustName = ""
                        CustCode = ""
                    End If
                    If Me.mPrimarySupplier <> -1 Then
                        Dim ObjSupp As New Core_Module.clsMasterSupplier
                        SuppName = ObjSupp.GetSupplierName(Me.mPrimarySupplier)
                        ObjSupp = Nothing
                    Else
                        SuppName = ""
                    End If


                    'simpan data item
                    ObjItem.Simpan_Material_Usage(vdgvItem, Me.mTipeTransaksi, Me.mTanggalTransaksi, Me.mPrimary, _
                     Me.mNoTransaksi, SuppName, CustName, vSetPOB, vUserCode)

                    ''berkaitan penomoran yang sama untuk bulan yang berbeda
                    'Me.CekNomor(DTA, Me.mTipeTransaksi)

                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mNoTransaksi + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                Me.mNoTransaksi = Me.GetNoBaru(Me.mTipeTransaksi, Me.mTanggalTransaksi)
                Return Me.SimpanBaru1_Material_Usage(vdgvItem, vUserCode)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
            ObjBooking = Nothing
        End Try
    End Function

    Public Overridable Function SimpanPerubahan_MaterialUsage(ByVal vdgvItem As DataGridView, ByVal vUserCode As String) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim ObjAntri As New Core_Module.ClsAntrian
        Dim SW As New Stopwatch

        Try
            SW.Start()

            'ambil no antrian
            ObjAntri.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Core_Module.ClsAntrian.EnumTipeAntrian.AntrianStockInOut_DO_Invoice, _
            Me.vPrimary_Antrian, Me.mPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiInvoice_SimpanPerubahan, vUserCode)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > ObjAntri.GetWaitTime Then
                    If SW.Elapsed.TotalSeconds >= ObjAntri.GetWaitTime Then
                        Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanPerubahan1_MaterialUsage(vdgvItem, vUserCode)
                            'hapus data antri untuk transaksi ini
                            If Me.DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
                                MsgBox("Logic error(1), contact your provider when you see this message", _
                                MsgBoxStyle.Information, "Information")
                            End If
                            'nilai balikan
                            Return retval
                        Else
                            'kontrol interval pengecekan
                            TryCounter += 1
                            CheckInterval = ((TryCounter \ 5) + 1) * 500
                            Me.Clock_Antri.Reset()
                            Me.Clock_Antri.Start()
                            Threading.Thread.Sleep(1000)
                        End If
                    End If
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjAntri = Nothing
            SW = Nothing
        End Try
    End Function

    Private Function SimpanPerubahan1_MaterialUsage(ByVal dgvItem As DataGridView, ByVal vUserCode As String) As Boolean
        MsgBox("obselete function")
        'Dim DTA As New dseTransaksiStockInOut_AssetTableAdapters.tbltransaksi_stockinoutTableAdapter
        'Dim ObjItem As New clsTransaksiStockInOutItem
        'Dim CustName, SuppName, CustCode As String
        'CustName = ""
        'CustCode = ""
        'Dim ObjBooking As New ClsTransaksiBookingInv
        'Dim ObjPOItem As New ClsTransaksiPOItem
        'Dim ListPrimaryPO As New List(Of Double)
        'Dim ListPrimaryPO2 As New List(Of Double)
        'Dim PrimaryPO As Double = -1

        'Try
        '    If Me.ApakahDataBenar_MaterialUsage(False, dgvItem) = True Then
        '        If DTA.UpdateQuery_AssetMaterialUsage(Me.mNoTransaksi, CInt(Me.mNoTransaksi.Substring(0, 5)), Me.mTipeTransaksi, Me.mTanggalTransaksi, _
        '        Me.mPrimaryPOPurchase, Me.mPrimaryDOSales, Me.mPrimaryCust, Me.mPrimarySupplier, _
        '        Me.mImporDariPackingList, Me.mPrimaryGIT, Me.mRemarks, vUserCode, Now, Now, Me.SupplierDoNo, Me.mPrimaryAsset, Me.mInOutAdjust, _
        '         Me.mMaterialUsageAmount, Me.mPrimary, Me.mTimestamp) > 0 Then
        '            'update data yang berubah setelah simpan perubahan
        '            Me.mNoUrut = Me.mNoTransaksi.Substring(0, 5)
        '            Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
        '            Me.mModifiedDate = Me.mTimestamp
        '            Me.mModifiedBy = vUserCode

        '            'email pemberitahuan
        '            If Me.mTipeTransaksi = EnumTipeStockInOut.StockOut_Sales Then
        '                Dim ObjDO As New ClsTransaksiDO
        '                Dim ObjitemDO As New ClsTransaksiDOItem
        '                If MsgBox("Send notification email?", MsgBoxStyle.YesNo, "Information") = MsgBoxResult.Yes Then
        '                    Me.SendStockOutSalesReadyNotificationEmail(Me.mNoTransaksi, ObjDO.GetDoNo(Me.mPrimaryDOSales), _
        '                    Me.mModifiedBy, CustCode, CustName)
        '                End If
        '                'update data primary stock out sales baru
        '                ObjDO.SetPrimaryStockOut(Me.mPrimaryDOSales, Me.mPrimary)
        '                'update data item delivery order sesuai dengan yang dituliskan pada bon out
        '                ObjitemDO.Hapus(Me.mPrimaryDOSales)
        '                ObjitemDO.Simpan_DariFormStockInOut(dgvItem, Me.mPrimaryDOSales)
        '                ObjDO = Nothing
        '                ObjitemDO = Nothing
        '            ElseIf Me.mTipeTransaksi = EnumTipeStockInOut.StockIn_SalesReturn Then
        '                Dim ObjDO As New ClsTransaksiDO
        '                Dim ObjLS As New Core_Module.clsMasterListedSoftware
        '                Dim ObjCust As New Core_Module.clsMasterCustomer
        '                If ObjLS.IsSalesInstalled = True Then
        '                    'notification email
        '                    If MsgBox("Send notification email?", MsgBoxStyle.YesNo, "Information") = MsgBoxResult.Yes Then
        '                        Me.SendStockInSalesReturnNotificationEmail(Me.mNoTransaksi, ObjDO.GetDoNo(Me.mPrimaryDOSales), _
        '                        Me.mModifiedBy, ObjDO.GetUserModifiedBy(Me.mPrimaryDOSales), ObjCust.GetCustName(Me.mPrimaryCust))
        '                    End If
        '                End If
        '                ObjLS = Nothing
        '                ObjDO = Nothing
        '                ObjCust = Nothing
        '            ElseIf Me.mTipeTransaksi = EnumTipeStockInOut.StockIn_Purchase Then
        '                Dim objGRN As New ClsTransaksiGRN
        '                Dim ObjPR As New ClsTransaksiPurchaseReturn
        '                Dim PrimaryGRN As Double = -1
        '                Dim PrimaryGRNBaru As Double = -1

        '                Dim NoGRN As String = ""

        '                PrimaryGRN = objGRN.GetPrimary_ByPrimaryStockInPurchase(Me.mPrimary)
        '                'hapus dulu GRN
        '                objGRN.HapusDariStockInPurchase(Me.mPrimary)
        '                'simpan GRN
        '                Me.SimpanGRN(dgvItem, vUserCode, NoGRN)

        '                'kirim email
        '                If MsgBox("Send notification email?", MsgBoxStyle.YesNo, "Information") = MsgBoxResult.Yes Then
        '                    Me.SendStockInPurchaseNotificationEmail(Me.mNoTransaksi, NoGRN, vUserCode)
        '                End If

        '                'update primary grn ke dalam semua transaksi purchase return yang merujuk
        '                PrimaryGRNBaru = objGRN.GetPrimary_ByPrimaryStockInPurchase(Me.mPrimary)
        '                ObjPR.SetPrimaryGRN(PrimaryGRN, PrimaryGRNBaru)

        '                objGRN = Nothing
        '                ObjPR = Nothing
        '            End If

        '            'ambil data primary po, untuk digunakan untuk perhitungan order balance
        '            Me.GetPrimaryPO(Me.mPrimary, ListPrimaryPO)
        '            'hapus data item lama
        '            ObjItem.Hapus(Me.mPrimary, Me.mTipeTransaksi, vUserCode)
        '            'simpan data item baru
        '            If Me.mPrimaryCust <> -1 Then
        '                Dim ObjCust As New Core_Module.clsMasterCustomer
        '                CustName = ObjCust.GetCustName(Me.mPrimaryCust)
        '                CustCode = ObjCust.GetCustomerCode(Me.mPrimaryCust)
        '                ObjCust = Nothing
        '            Else
        '                CustName = ""
        '                CustCode = ""
        '            End If
        '            If Me.mPrimarySupplier <> -1 Then
        '                Dim ObjSupp As New Core_Module.clsMasterSupplier
        '                SuppName = ObjSupp.GetSupplierName(Me.mPrimarySupplier)
        '                ObjSupp = Nothing
        '            Else
        '                SuppName = ""
        '            End If

        '            'untuk simpan item disini, tidak perlu set pob, karena perhitungannya akan dilakukan disini
        '            ObjItem.Simpan_Material_Usage(dgvItem, Me.mTipeTransaksi, Me.mTanggalTransaksi, Me.mPrimary, _
        '            Me.mNoTransaksi, SuppName, CustName, False, vUserCode)

        '            'hitung order balance
        '            Me.GetPrimaryPO(Me.mPrimary, ListPrimaryPO2)
        '            For Each PrimaryPO In ListPrimaryPO2
        '                If ListPrimaryPO.Contains(PrimaryPO) = False Then
        '                    ListPrimaryPO.Add(PrimaryPO)
        '                End If
        '            Next
        '            ObjPOItem.SetOrderBalance(ListPrimaryPO)

        '            'berkaitan dengan proses booking inventory
        '            If Me.mTipeTransaksi = EnumTipeStockInOut.StockIn_Purchase Then
        '                If Me.mPrimaryGIT = -1 Then
        '                    'jika langsung dari stock in purchase maka simpan data booking inventory jika ada
        '                    ObjBooking.CreateBookingInventory_TanpaGIT(Me.mPrimary, vUserCode)
        '                Else
        '                    'jika dari packing list, cek apakah ada booking git, jika ada maka 
        '                    'ubah data booking git menjadi booking inventory
        '                    ObjBooking.CreateBookingInventory_DenganGIT(Me.mPrimary, Me.mPrimaryGIT, vUserCode)
        '                End If
        '            End If

        '            'simpan data jurnal
        '            If Me.mTipeTransaksi = EnumTipeStockInOut.StockOut_PurchaseReturn Then
        '                Me.TransferJurnal(Me.mPrimarySupplier, Me.mPrimary, -1)
        '            End If

        '            'nilai balikan
        '            Return True
        '        Else
        '            'cek apakah timestampnya beda, jika ya maka minta user untuk reload data
        '            If DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value <> Me.mTimestamp Then
        '                MsgBox("Other user have modified this transaction, please reload from database.", _
        '                MsgBoxStyle.Information, "Information")
        '            End If
        '            Return False
        '        End If
        '    End If
        'Catch ex As Exception
        '    If ex.Message = "Duplicate entry '" + Me.mNoTransaksi + "' for key 2" Then
        '        'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
        '        Me.mNoTransaksi = Me.GetNoBaru(Me.mTipeTransaksi, Me.mTanggalTransaksi)
        '        Return Me.SimpanPerubahan1(dgvItem, vUserCode)
        '    Else
        '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        '    End If
        'Finally
        '    DTA.Dispose()
        '    ObjItem = Nothing
        '    ObjBooking = Nothing
        '    ObjPOItem = Nothing
        '    ListPrimaryPO = Nothing
        '    ListPrimaryPO2 = Nothing
        'End Try
    End Function

    Public Overridable Function ApakahDataBenar_MaterialUsage(ByVal vBaru As Boolean, ByVal vDGVItem As DataGridView) As Boolean
        Dim DTA As New dseTransaksiStockInOut_AssetTableAdapters.transaksi_stockinout_item_viewTableAdapter
        Dim TBL As New dseTransaksiStockInOut_Asset.transaksi_stockinout_item_viewDataTable
        Dim rowInOutLama As dseTransaksiStockInOut_Asset.transaksi_stockinout_item_viewRow

        Dim DTAMasterProductQty As New Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter
        DTAMasterProductQty.ClearBeforeFill = False
        Dim TBLMasterProductQty As New Core_Module.dseMasterProduct.tblmaster_productqtyDataTable

        Dim rowDGV As DataGridViewRow
        Dim BatchNo As String
        Dim PrimaryProduct, PrimaryLocation, Sec_Qty, Qty As Decimal
        Dim ExpiryDate As Date

        Dim ListPrimaryProduct As New Queue(Of Double)
        Dim rowProduct As Core_Module.dseMasterProduct.tblmaster_productqtyRow
        Dim Dummy As String = ""


        Dim ListPrimaryPO As New List(Of Double)
        Dim ObjProd As New Core_Module.clsMasterProduct
        Dim LotNoKosong As String = ""
        Dim ListDummy As New List(Of String)
        Dim Warning As String = ""
        Dim Msg As String = ""
        Dim i As Integer = 1
        Dim filter As String = "'"
        Dim TBLStock As New dseTransaksiStockInOut_Asset.tbltransaksistockinpurchase_viewDataTable
        Dim ObjPL As New ClsTransaksiPackingList
        Dim NoPL As String = ""

        Dim frm As New FrmTransaksiStockInPurchase2
        Dim frmMaterialUsage As New FrmTransaksiAssetAcquisition_MaterialUsage

        Try
            If vBaru = True Then
                'jika sudah ada, maka ganti no transaksi
                If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                    Me.mNoTransaksi = Me.GetNoBaru(Me.mTipeTransaksi, Me.mTanggalTransaksi, Me.mInOutAdjust)
                End If
            Else
                If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                    If Me.mPrimary <> Me.GetPrimary(Me.mNoTransaksi) Then
                        Me.mNoTransaksi = Me.GetNoBaru(Me.mTipeTransaksi, Me.mTanggalTransaksi, Me.mInOutAdjust)
                    End If
                End If
            End If

            'load data item lama
            If vBaru = False Then
                'load database master product qty berdasarkan data product lama dan data product baru
                TBL.Rows.Clear()
                DTA.Fill(TBL, Me.mPrimary)
                For Each rowInOutLama In TBL.Rows
                    With rowInOutLama
                        DTAMasterProductQty.FillBy_Parameter(TBLMasterProductQty, _
                        .Secondary_Qty, .Primary_Product, .Batch_No, .Expiry_Date)
                    End With
                Next
            End If

            'load data item baru
            For Each rowDGV In vDGVItem.Rows
                BatchNo = rowDGV.Cells(frmMaterialUsage.ColBatch.Index).Value
                PrimaryLocation = rowDGV.Cells(frmMaterialUsage.ColPrimaryLocation.Index).Value
                If IsNumeric(rowDGV.Cells(frmMaterialUsage.ColSecondaryQty.Index).Value) = True Then
                    Sec_Qty = rowDGV.Cells(frmMaterialUsage.ColSecondaryQty.Index).Value
                Else
                    Sec_Qty = "-1"
                End If
                If IsDate(rowDGV.Cells(frmMaterialUsage.ColExpiredDate.Index).Value) = False Then
                    ExpiryDate = #1/1/1950#
                Else
                    ExpiryDate = CDate(rowDGV.Cells(frmMaterialUsage.ColExpiredDate.Index).Value)
                End If
                PrimaryProduct = rowDGV.Cells(frmMaterialUsage.ColPrimaryCodeNo.Index).Value
                DTAMasterProductQty.FillBy_Parameter(TBLMasterProductQty, _
                Sec_Qty, PrimaryProduct, BatchNo, ExpiryDate)
            Next

            If Me.mTipeTransaksi = EnumTipeStockInOut.StockOut_Adjust Or _
            Me.mTipeTransaksi = EnumTipeStockInOut.StockOut_PurchaseReturn Or _
            Me.mTipeTransaksi = EnumTipeStockInOut.StockOut_Process Or _
            Me.mTipeTransaksi = EnumTipeStockInOut.StockOut_Sales Then
                'simulasikan untuk stock out
                'simulasi hapus data item lama
                If TBL.Rows.Count > 0 Then
                    For Each rowInOutLama In TBL.Rows
                        With rowInOutLama
                            Me.Koreksi4HapusOut(TBLMasterProductQty, _
                            .Primary_Product, .Qty, .Secondary_Qty, _
                            .Primary_Storage_Area, .Batch_No, .Expiry_Date)
                        End With
                    Next
                End If
                'simulasi simpan item baru
                For Each rowDGV In vDGVItem.Rows
                    BatchNo = rowDGV.Cells(frmMaterialUsage.ColBatch.Index).Value
                    PrimaryLocation = rowDGV.Cells(frmMaterialUsage.ColPrimaryLocation.Index).Value
                    If IsNumeric(rowDGV.Cells(frmMaterialUsage.ColSecondaryQty.Index).Value) = True Then
                        Sec_Qty = rowDGV.Cells(frmMaterialUsage.ColSecondaryQty.Index).Value
                    Else
                        Sec_Qty = "-1"
                    End If
                    If IsDate(rowDGV.Cells(frmMaterialUsage.ColExpiredDate.Index).Value) = False Then
                        ExpiryDate = #1/1/1950#
                    Else
                        ExpiryDate = CDate(rowDGV.Cells(frmMaterialUsage.ColExpiredDate.Index).Value)
                    End If
                    PrimaryProduct = rowDGV.Cells(frmMaterialUsage.ColPrimaryCodeNo.Index).Value
                    Qty = rowDGV.Cells(frmMaterialUsage.ColQty.Index).Value

                    Me.Koreksi4SimpanOut(TBLMasterProductQty, _
                    PrimaryProduct, Qty, Sec_Qty, PrimaryLocation, _
                    BatchNo, ExpiryDate)
                Next

            End If

            'cek apakah ada qty yang minus setelah dilakukan simulasi, jika ada
            'maka return false
            ListPrimaryProduct.Clear()
            For Each rowProduct In TBLMasterProductQty.Rows
                If rowProduct.Qty < 0 Then
                    ListPrimaryProduct.Enqueue(rowProduct.Primary_Product)
                End If
            Next
            If ListPrimaryProduct.Count > 0 Then
                Dim ObjProduct As New Core_Module.clsMasterProduct
                While ListPrimaryProduct.Count > 0
                    If Dummy = "" Then
                        Dummy = ObjProduct.GetProductCode(ListPrimaryProduct.Dequeue)
                    Else
                        Dummy = Dummy + ", " + ObjProduct.GetProductCode(ListPrimaryProduct.Dequeue)
                    End If
                End While
                ObjProduct = Nothing

                MsgBox("The following product(s) will be minus: " + Dummy + ", process cancelled.", _
                MsgBoxStyle.Information, "Information")

                Return False
            End If

            If Me.InOutAdjust = EnumInOutAdjust.OutAdjustMaterialUsageAssetAcquisition Then
                If Me.ObjAsset.GetStatus(Me.PrimaryAsset) <> Core_Module.ClsMasterAsset.EnumStatusAsset.On_Progress Then
                    MsgBox("Asset status isn't on progress, you can't save this data.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If
            If Me.InOutAdjust = EnumInOutAdjust.OutAdjustMaterialUsageAssetRevaluation Then
                If Me.ObjAsset.GetStatus(Me.PrimaryAsset) <> Core_Module.ClsMasterAsset.EnumStatusAsset.Active Then
                    MsgBox("Asset status isn't active, you can't save this data.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            DTAMasterProductQty.Dispose()
            TBLMasterProductQty.Dispose()
            ListPrimaryProduct = Nothing
            ListPrimaryPO = Nothing
            ListDummy = Nothing
            TBLStock.Dispose()
        End Try
    End Function

    Public Function GetPrimaryAsset(ByVal vPrimaryStockInOut As Double) As Double
        Dim DTA As New dseTransaksiStockInOut_AssetTableAdapters.tbltransaksi_stockinoutTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimaryAsset(vPrimaryStockInOut).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryAsset(vPrimaryStockInOut).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetProductGood(ByVal vPrimaryProduct As Double) As Integer
        Dim DTA As New dseTransaksiStockInOut_AssetTableAdapters.tblmaster_productTableAdapter
        Try
            If DTA.ScalarQuery_GetGoodService(vPrimaryProduct).HasValue = True Then
                Return DTA.ScalarQuery_GetGoodService(vPrimaryProduct).Value
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetJml_Asset(ByVal Criteria As String, _
   ByVal SearchCriteria As Integer, Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
   Optional ByVal vTanggalSampai As Date = (#1/1/1950#), _
   Optional ByVal vInOutAdjust As clsTransaksiStockInOut.EnumInOutAdjust = EnumInOutAdjust.NoTipe) As Long
        Dim DTA As New dseTransaksiStockInOut_AssetTableAdapters.tbl_transaksi_material_usage_viewTableAdapter

        Try
            If SearchCriteria = EnumSreachCriteria_Material_Usage.NoTransaksi Then
                If DTA.ScalarQuery_GetJmlLikeNoTransaksi(Criteria + "%", vInOutAdjust, clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust).HasValue = True Then
                    Return DTA.ScalarQuery_GetJmlLikeNoTransaksi(Criteria + "%", vInOutAdjust, clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust)
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSreachCriteria_Material_Usage._Date Then
                If DTA.ScalarQuery_GetJmlByTanggalTransaksi(vInOutAdjust, clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust, vTanggalDari, vTanggalSampai).HasValue = True Then
                    Return DTA.ScalarQuery_GetJmlByTanggalTransaksi(vInOutAdjust, clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust, vTanggalDari, vTanggalSampai).Value
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSreachCriteria_Material_Usage.AssetCode Then
                If DTA.ScalarQuery_GetJmlLikeAssetCode(vInOutAdjust, clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust, Criteria + "%").HasValue = True Then
                    Return DTA.ScalarQuery_GetJmlLikeAssetCode(vInOutAdjust, clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust, Criteria + "%")
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSreachCriteria_Material_Usage.AssetDescrition Then
                If DTA.ScalarQuery_GetJmlLikeAssetDeskription(vInOutAdjust, clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust, Criteria + "%").HasValue = True Then
                    Return DTA.ScalarQuery_GetJmlLikeAssetDeskription(vInOutAdjust, clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust, Criteria + "%")
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSreachCriteria_Material_Usage.Remarks Then
                If DTA.ScalarQuery_GetJmlLikeRemarks(vInOutAdjust, clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust, Criteria + "%").HasValue = True Then
                    Return DTA.ScalarQuery_GetJmlLikeRemarks(vInOutAdjust, clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust, Criteria + "%")
                Else
                    Return 0
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function CariDetail_MaterialUsage(ByVal vPrimaryTransaksi As Double, _
ByVal vDgvItem As DataGridView) As Boolean
        Dim DTA As New dseTransaksiStockInOut_AssetTableAdapters.tbltransaksi_stockinout_Material_UsageTableAdapter
        Dim TBL As New dseTransaksiStockInOut_Asset.tbltransaksi_stockinout_Material_UsageDataTable
        Dim row As dseTransaksiStockInOut_Asset.tbltransaksi_stockinout_Material_UsageRow
        Dim ObjItem As New clsTransaksiStockInOutItem

        Try
            'fil basis data
            If DTA.FillBy_Primary(TBL, vPrimaryTransaksi) > 0 Then
                'salin data header
                row = TBL.Rows(0)

                With row
                    Me.mImporDariPackingList = .Impor_Dari_Packing_List
                    Me.mNoTransaksi = .No_Transaksi
                    Me.mNoUrut = .Urutan
                    Me.mPrimary = .Primary
                    Me.mPrimaryCust = .Primary_Customer
                    Me.mPrimaryDOSales = .Primary_DO_Sales
                    Me.mPrimaryGIT = .Primary_GIT
                    Me.mPrimaryPOPurchase = .Primary_PO_Purchase
                    Me.mPrimarySupplier = .Primary_Supplier
                    Me.mRemarks = .Remarks
                    Me.mTanggalTransaksi = .Tanggal_Transaksi
                    Me.mTipeTransaksi = .Tipe_Transaksi
                    Me.mPrimaryAsset = .Primary_Asset
                    Me.mInOutAdjust = .In_Out_Adjust
                    Me.mMaterialUsageAmount = .Material_Usage_Amount


                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
                    Me.mSupplierDoNo = ._Supplier_s_DO_Number
                End With
                'salin data item 
                ObjItem.IsiDGVMaterialUsage(vDgvItem, Me.mPrimary, Me.mTipeTransaksi)
                'nilai balikan
                Return True
            Else
                Me.ResetObject(vDgvItem)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjItem = Nothing
        End Try
    End Function
End Class
