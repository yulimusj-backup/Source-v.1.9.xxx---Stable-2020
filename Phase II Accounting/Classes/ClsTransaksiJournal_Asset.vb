Option Explicit On
Option Strict Off

Partial Public Class ClsTransaksiJournal
    Private mPrimaryAssetAccount As Double = -1
    Private mAmount As Decimal = 0
    Private mPrimaryAsset As Double = -1
    Private objAsset As Core_Module.ClsMasterAsset

    Public Property PrimaryAsset() As Double
        Get
            Return Me.mPrimaryAsset
        End Get
        Set(ByVal value As Double)
            mPrimaryAsset = value
        End Set
    End Property

    Public Property Amount() As Decimal
        Get
            Return Me.mAmount
        End Get
        Set(ByVal value As Decimal)
            mAmount = value
        End Set
    End Property

    Public Property PrimaryAssetAccount() As Double
        Get
            Return Me.mPrimaryAssetAccount
        End Get
        Set(ByVal value As Double)
            mPrimaryAssetAccount = value
        End Set
    End Property

    Public Function GetQty(ByVal vPrimaryAop As Double, ByVal vPrimaryProduct As Double, _
                     ByVal vPrimaryPO As Double) As Decimal
        Dim DTA As New DseTransaksiJournal_AssetTableAdapters.tbltrans_aop_itemTableAdapter

        Try
            If DTA.ScalarQuery_GetQty(vPrimaryAop, vPrimaryProduct, vPrimaryPO).HasValue = True Then
                Return DTA.ScalarQuery_GetQty(vPrimaryAop, vPrimaryProduct, vPrimaryPO).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetAtPrice(ByVal vPrimaryAop As Double, ByVal vPrimaryProduct As Double, _
                 ByVal vPrimaryPO As Double) As Decimal
        Dim DTA As New DseTransaksiJournal_AssetTableAdapters.tbltrans_aop_itemTableAdapter

        Try
            If DTA.ScalarQuery_GetAtPrice(vPrimaryAop, vPrimaryProduct, vPrimaryPO).HasValue = True Then
                Return DTA.ScalarQuery_GetAtPrice(vPrimaryAop, vPrimaryProduct, vPrimaryPO).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub CariDataAop(ByVal vPrimarySupp As Double, ByVal CboAop As ComboBox)
        Dim DTA As New DseTransaksiJournal_AssetTableAdapters.tbltrans_aopTableAdapter
        Dim TBL As New DseTransaksiJournal_Asset.tbltrans_aopDataTable
        Try
            DTA.Fill(TBL, vPrimarySupp)
            TBL.DefaultView.Sort = "[No Transaksi]"
            CboAop.DataSource = TBL.DefaultView
            CboAop.DisplayMember = "No Transaksi"
            CboAop.Refresh()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Public Sub CariDataProduct(ByVal vPrimaryAop As Double, ByVal CboProduct As ComboBox)
        Dim DTA As New DseTransaksiJournal_AssetTableAdapters.tbltrans_aop_itemTableAdapter
        Dim TBL As New DseTransaksiJournal_Asset.tbltrans_aop_itemDataTable
        Dim ObjProduct As New Core_Module.clsMasterProduct
        Dim i As Integer = Core_Module.clsMasterProduct.EnumTipeProduct.Goods

        Try
            DTA.FillBy_PrimaryAOP(TBL, vPrimaryAop)
            TBL.DefaultView.Sort = "[Product Code]"
            CboProduct.DataSource = TBL.DefaultView
            CboProduct.DisplayMember = "Product Code"
            CboProduct.Refresh()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Public Function GetNoBaru(ByVal vTanggalTransaksi As Date, _
   Optional ByVal vTipeJurnal2 As EnumTipeJurnal2 = EnumTipeJurnal2.AssetAcquistion_Other) As String
        Dim DTA As New DseTransaksiJournalTableAdapters.tbltrans_jurnalTableAdapter
        Dim TanggalStart, TanggalEnd As Date
        Dim NoUrut, InOut, Bulan As String
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            InOut = ""
            'cari tanggal sampai dan dari
            TanggalEnd = ObjCD.GetLastDateOfTheYear(vTanggalTransaksi)
            TanggalStart = ObjCD.GetFisrtDateOfTheYear(vTanggalTransaksi)

            If vTipeJurnal2 = EnumTipeJurnal2.AssetAcquistion_Other Then
                InOut = "ASSET-ACQUISITION"
            ElseIf vTipeJurnal2 = EnumTipeJurnal2.AssetRevaluation_Other Then
                InOut = "ASSET-REVALUATION"
            End If

            'cari no urut

            If DTA.ScalarQuery_GetJmlByTipeJurnal2(vTipeJurnal2, TanggalStart, TanggalEnd).HasValue = True Then
                NoUrut = CStr(DTA.ScalarQuery_GetJmlByTipeJurnal2(vTipeJurnal2, TanggalStart, TanggalEnd).Value) + 1
                If NoUrut <= 0 Then
                    NoUrut = 1
                End If
            Else
                NoUrut = 1
            End If

            'pastikan no urut terdiri atas 6 huruf
            While NoUrut.Length < 6
                NoUrut = "0" + NoUrut
            End While

            'bulan 
            Bulan = CStr(vTanggalTransaksi.Month)
            While Bulan.Length < 2
                Bulan = "0" + Bulan
            End While
            Return NoUrut + "/" + InOut + "/" + Bulan + "/" + CStr(vTanggalTransaksi.Year)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Function GetPrimaryPO(ByVal vPrimaryAop As Double, ByVal vPrimaryProduct As Double) As Decimal
        Dim DTA As New DseTransaksiJournal_AssetTableAdapters.tbltrans_aop_itemTableAdapter

        Try
            If DTA.ScalarQuery_GetPrimaryPO(vPrimaryAop, vPrimaryProduct).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryPO(vPrimaryAop, vPrimaryProduct).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetBookValue_DefineValue(ByVal vPrimaryAsset As Double, ByVal vPlusMinus As Integer, _
    ByVal vPrimaryJournal As Double, ByVal vRevSendiri As Decimal) As Decimal
        Dim DTARev As New DseTransaksiJournal_AssetTableAdapters.DataTable7TableAdapter
        Dim DTADep As New DseTransaksiJournal_AssetTableAdapters.DataTable8TableAdapter
        Dim DTADisposal As New DseTransaksiJournal_AssetTableAdapters.tbltrans_jurnal_item_disposalTableAdapter
        Dim DefineValue As Decimal = 0
        Dim AcquiredValue As Decimal = Me.objAsset.GetAcquiredValue(vPrimaryAsset)
        Dim Revaluation As Decimal = 0
        Dim DepValue As Decimal = 0
        Dim Disposal As Decimal = 0
        Try
            'khusus revaluasi
            If Me.TipeJurnal2 = EnumTipeJurnal2.AssetRevaluation_Other Then
                If Me.objAsset.GetTypeAcquistion(vPrimaryAsset) = Core_Module.ClsMasterAsset.EnumAcquisitionType.Define_Value Then
                    'cari Acquired value
                    AcquiredValue = Me.objAsset.GetAcquiredValue(vPrimaryAsset)
                    'cari total revaluasi keseluruhan
                    Revaluation = DTARev.ScalarQuery_GetTotalRevaluation(vPrimaryJournal, ClsTransaksiJournal.EnumTipeJournal.GeneralJournal, ClsTransaksiJournal.EnumTipeJurnal2.AssetRevaluation_Other, _
                     vPrimaryAsset)
                    'menentukan plus/minus revaluasi
                    If vPlusMinus = ClsTransaksiJournal.EnumTipeInOut.out Then
                        Revaluation = Revaluation * -1
                    ElseIf vPlusMinus = EnumTipeInOut._In Then
                        vRevSendiri = vRevSendiri * -1
                    End If
                    'cari depresiasi value
                    DepValue = DTADep.ScalarQuery_GetDepreciationValue(vPrimaryJournal, ClsTransaksiJournal.EnumTipeJournal.GeneralJournal, ClsTransaksiJournal.EnumTipeJurnal2.AssetRevaluation_Other, _
                     vPrimaryAsset)
                    'cari disposal
                    Disposal = DTADisposal.ScalarQuery_GetDisposalValue(vPrimaryJournal)

                    DefineValue = AcquiredValue + Revaluation - DepValue - Disposal - vRevSendiri
                    Return DefineValue
                Else
                    Return 0
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetBookValue_BuyOrBuild(ByVal vPrimaryAsset As Double, ByVal vPlusMinus As Integer, _
ByVal vPrimaryJournal As Double, ByVal vRevSendiri As Decimal) As Decimal
        Dim DTARev As New DseTransaksiJournal_AssetTableAdapters.DataTable7TableAdapter
        Dim DTADep As New DseTransaksiJournal_AssetTableAdapters.DataTable8TableAdapter
        Dim DTADisposal As New DseTransaksiJournal_AssetTableAdapters.tbltrans_jurnal_item_disposalTableAdapter
        Dim BuyOrBulid As Decimal = 0
        Dim OtherAcquisition As Decimal = 0
        Dim Revaluation As Decimal = 0
        Dim DepValue As Decimal = 0
        Dim Disposal As Decimal = 0
        Dim MaterialUsage As Decimal = 0
        Try
            'khusus revaluasi
            If Me.TipeJurnal2 = EnumTipeJurnal2.AssetRevaluation_Other Then
                If Me.objAsset.GetTypeAcquistion(vPrimaryAsset) = Core_Module.ClsMasterAsset.EnumAcquisitionType.Define_Value Then
                    'cari Other revaluation
                    OtherAcquisition = DTARev.ScalarQuery_GetTotalRevaluation(vPrimaryJournal, ClsTransaksiJournal.EnumTipeJournal.GeneralJournal, ClsTransaksiJournal.EnumTipeJurnal2.AssetAcquistion_Other, _
                     vPrimaryAsset)
                    'cari total revaluasi keseluruhan
                    Revaluation = DTARev.ScalarQuery_GetTotalRevaluation(vPrimaryJournal, ClsTransaksiJournal.EnumTipeJournal.GeneralJournal, ClsTransaksiJournal.EnumTipeJurnal2.AssetAcquistion_Other, _
                     vPrimaryAsset)
                    'menentukan plus/minus revaluasi
                    If vPlusMinus = ClsTransaksiJournal.EnumTipeInOut.out Then
                        Revaluation = Revaluation * -1
                    ElseIf vPlusMinus = EnumTipeInOut._In Then
                        vRevSendiri = vRevSendiri * -1
                    End If
                    'cari depresiasi value
                    DepValue = DTADep.ScalarQuery_GetDepreciationValue(vPrimaryJournal, ClsTransaksiJournal.EnumTipeJournal.GeneralJournal, ClsTransaksiJournal.EnumTipeJurnal2.AssetRevaluation_Other, _
                     vPrimaryAsset)
                    'cari disposal
                    Disposal = DTADisposal.ScalarQuery_GetDisposalValue(vPrimaryJournal)

                    BuyOrBulid = OtherAcquisition + MaterialUsage + Revaluation - DepValue - Disposal - vRevSendiri
                    Return BuyOrBulid
                Else
                    Return 0
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Overridable Function SimpanPerubahan_Asset(ByVal vDGVItem As DataGridView, ByVal vDGVItem_pay As DataGridView, _
  ByVal vUserCode As String, ByVal frm As FrmTransaksiJournal_Acquisition_Revaluation_Other) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim ObjAntri As New Core_Module.ClsAntrian
        Dim SW As New Stopwatch
        Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter

        Try
            SW.Start()

            'ambil no antrian
            ObjAntri.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, DTAInOUT, Core_Module.ClsAntrian.EnumTipeAntrian.AntrianStockInOut_DO_Invoice, _
            Me.vPrimary_Antrian, Me.mPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiJurnal_Asset_Acq_SimpanPerubahan, vUserCode)


            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= ObjAntri.GetWaitTime Then
                        DTAInOUT.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanPerubahan1_Asset(vDGVItem, vDGVItem_pay, vUserCode, frm)
                            'hapus data antri untuk transaksi ini
                            If DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
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

    Public Overridable Function SimpanBaru_Asset(ByVal vDGVItem As DataGridView, ByVal vDGVItem_Payment As DataGridView, _
    ByVal vUserCode As String, ByVal Frm As FrmTransaksiJournal_Acquisition_Revaluation_Other) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch
        Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter

        Try
            SW.Start()
            'ambil no antrian
            Me.ObjAntrin.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
            Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiJurnal_Asset_Acq_SimpanBaru, vUserCode)
            'Me.GetNoAntrian()

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                If SW.Elapsed.TotalSeconds >= ObjAntrin.GetWaitTime Then
                    DTAInOUT.Delete(Me.vPrimary_Antrian)
                    MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                    Exit Function
                Else
                    'hanya mencoba ke database setelah interval waktu tertentu saja
                    If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                        If DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanBaru1_Asset(vDGVItem, vDGVItem_Payment, vUserCode, Frm)
                            'hapus data antri untuk transaksi ini
                            If DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
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
            SW = Nothing
        End Try
    End Function

    Public Overridable Function SimpanBaru1_Asset(ByVal vDGVItem As DataGridView, ByVal vDGVPayment As DataGridView, _
       ByVal vUserCode As String, ByVal Frm As FrmTransaksiJournal_Acquisition_Revaluation_Other) As Boolean
        Dim DTA As New DseTransaksiJournal_AssetTableAdapters.tbltrans_jurnal_AssetTableAdapter
        Dim ObjItem As New ClsTransaksiJournal_Item
        Try
            Me.mUrutan = -1
            Me.mPrimaryDokumenAcuan = -1
            Me.mPrimaryPihakAcuan = -1

            If Me.ApakahDataBenar1(True, vDGVItem, Frm.ColPrimaryCurrency.Index, Me.Tanggal) = True Then
                If DTA.Insert(Me.RefNumber, Me.Tanggal, Me.Urutan, Me.TipeJournal, Me.PrimaryDokumenAcuan, _
                Me.PrimaryPihakAcuan, Me.Description, vUserCode, Now, _
                 vUserCode, Now, Now, Me.PrimaryAkunPayTo, Me.PrimaryAkunDifference, Me.Paid, Me.ExchangeRate, Me.Paid_IDR, _
                 Me.ARPaid_IDR, Me.Difference, Me.ExchageRateGainLoss, Me.TipeJurnal2, Me.PrimaryCurrency, Me.TipeInOut, _
                 Me.PrimaryAssetAccount, Me.PrimaryAsset) > 0 Then
                    'update data yang berubah setelah insert query
                    Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mRefNumber).Value
                    Me.mTimestamp = DTA.ScalarQuery_GetTimeStamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    'simpan data item
                    ObjItem.SimpanItem_Asset(vDGVItem, Me.mPrimary, Frm, vDGVPayment)

                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mRefNumber + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                ''Me.mNoTransaksi = Me.GetNoBaru(Me.mTanggalTransaksi)
                Return Me.SimpanBaru1_Asset(vDGVItem, vDGVPayment, vUserCode, Frm)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Overridable Function SimpanPerubahan1_Asset(ByVal vDGVItem As DataGridView, ByVal vDGVItem_Pay As DataGridView, _
   ByVal vUserCode As String, ByVal frm As FrmTransaksiJournal_Acquisition_Revaluation_Other) As Boolean
        Dim DTA As New DseTransaksiJournal_AssetTableAdapters.tbltrans_jurnal_AssetTableAdapter
        Dim ObjItem As New ClsTransaksiJournal_Item
        Dim ObjCek As New ClsTransaksiJournal_Item_Cheque

        Try
            Me.mUrutan = -1
            Me.mPrimaryDokumenAcuan = -1
            Me.mPrimaryPihakAcuan = -1
            If Me.ApakahDataBenar1(False, vDGVItem, frm.ColPrimaryCurrency.Index, Me.Tanggal) = True Then
                If DTA.Update(Me.RefNumber, Me.Tanggal, Me.Urutan, Me.TipeJournal, Me.PrimaryDokumenAcuan, _
                Me.PrimaryPihakAcuan, Me.Description, vUserCode, Now, Now, _
                Me.PrimaryAkunPayTo, Me.PrimaryAkunDifference, Me.Paid, Me.ExchangeRate, Me.Paid_IDR, Me.ARPaid_IDR, Me.Difference, _
                Me.ExchageRateGainLoss, Me.TipeJurnal2, Me.PrimaryCurrency, Me.TipeInOut, Me.PrimaryAssetAccount, Me.PrimaryAsset, _
                Me.Primary, Me.Timestamp) > 0 Then
                    'update data yang berubah setelah simpan perubahan
                    Me.mTimestamp = DTA.ScalarQuery_GetTimeStamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode

                    'hapus data item lama
                    ObjItem.HapusItem(Me.mPrimary)
                    ObjCek.HapusItem(Me.mPrimary)

                    'simpan data item baru
                    ObjItem.SimpanItem_Asset(vDGVItem, Me.mPrimary, frm, vDGVItem_Pay)

                    'nilai balikan
                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mRefNumber + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                ''Me.mNoTransaksi = Me.GetNoBaru(Me.mTanggalTransaksi)
                Return Me.SimpanPerubahan1_Asset(vDGVItem, vDGVItem_Pay, vUserCode, frm)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Overridable Function CariDetail_Asset(ByVal vPrimaryMaster As Double, _
 ByVal vDGVItem As DataGridView, ByVal frm As FrmTransaksiJournal_Acquisition_Revaluation_Other, _
 ByVal tipe_journal As Integer, ByVal vDGVItem_payment As DataGridView) As Boolean
        Dim DTA As New DseTransaksiJournal_AssetTableAdapters.tbltrans_jurnal_AssetTableAdapter
        Dim TBL As New DseTransaksiJournal_Asset.tbltrans_jurnal_AssetDataTable
        Dim row As DseTransaksiJournal_Asset.tbltrans_jurnal_AssetRow
        Dim ObjItem As New ClsTransaksiJournal_Item

        Try
            'fil basis data
            If DTA.FillBy_Primary(TBL, vPrimaryMaster, tipe_journal) > 0 Then
                'salin data header
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = vPrimaryMaster
                    Me.mRefNumber = .Ref_Number
                    Me.mDate = ._Date
                    Me.mUrutan = .Urutan
                    Me.mTipeJournal = .Tipe_Journal
                    Me.mPrimaryDokumenAcuan = .Primary_Document_Acuan
                    Me.mPrimaryPihakAcuan = .Primary_Pihak_Acuan
                    Me.mDesc = .Desc
                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
                    Me.mTipeJurnal2 = .Tipe_Jurnal_2

                    Me.mPrimaryAkunDifference = .primary_akun_difference
                    Me.mPrimaryAkunPayTo = .primary_akun_pay_to
                    Me.mPaid = .paid
                    Me.mPaid_IDR = .paid_IDR
                    Me.mARPaid_IDR = .AR_paid_IDR
                    Me.mDifference = .difference
                    Me.mExchangeRate = .exchange_rate
                    Me.mExchageRateGainLoss = .exchange_rate_gain_loss
                    Me.mPrimaryCurr = .Primary_Currency
                    Me.mTipeInOut = .Type_InOut
                    Me.mPrimaryAssetAccount = .Primary_Akun_Asset
                    Me.mPrimaryAsset = .Primary_Asset
                    Me.mReferenceParty = ""
                End With
                'salin data item ada 2 item 
                'CodeNo
                ObjItem.IsiDGVItem_Other(vDGVItem, vPrimaryMaster, frm, vDGVItem_payment)
                'nilai balikan
                Return True
            Else
                Me.ResetObject(vDGVItem)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Sub CariDataPO(ByVal vPrimaryAop As Double, ByVal vPrimaryProduct As Double, _
                          ByVal CboPO As ComboBox)
        Dim DTA As New DseTransaksiJournal_AssetTableAdapters.tbltrans_aop_itemTableAdapter
        Dim TBL As New DseTransaksiJournal_Asset.tbltrans_aop_itemDataTable
        Try
            DTA.FillBy_PrimaryAOP_PrimaryProduct(TBL, vPrimaryAop, vPrimaryProduct)
            TBL.DefaultView.Sort = "[No Transaksi]"
            CboPO.DataSource = TBL.DefaultView
            CboPO.DisplayMember = "No Transaksi"
            CboPO.Refresh()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Public Function GetPrimaryExpense(ByVal vPrimaryProduct As Double) As Double
        Dim DTA As New DseTransaksiJournal_AssetTableAdapters.tblmaster_productTableAdapter

        Try
            If DTA.ScalarQuery_GetPrimaryExpense(vPrimaryProduct).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryExpense(vPrimaryProduct).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
