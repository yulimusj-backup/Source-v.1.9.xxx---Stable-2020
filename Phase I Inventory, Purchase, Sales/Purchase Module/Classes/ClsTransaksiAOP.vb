Option Explicit On
Option Strict Off

Public Class ClsTransaksiAOP
    Private mprimary As Double = -1
    Private mnotransaksi As String = ""
    Private mNoUrut As Integer = 0
    Private mTipeAOP As EnumTipeAop = EnumTipeAop.AOP
    Private mTanggalTransaksi As Date = Now
    Private mPrimarySupplier As Double = -1
    Private mSupplierInvoiceNo As String = ""
    Private mLokal As EnumLokal = EnumLokal.Lokal
    Private mPrimaryCurrency As Double = -1
    Private mSubtotalBeforeDP As Decimal = 0
    Private mCalculatesBy As EnumCalculatesBy = EnumCalculatesBy.AOP
    Private mDownPayment As Decimal = 0
    Private mSubTotal As Decimal = 0
    Private mDutyStamp As Decimal = 0
    Private mTaxExclude As EnumTaxExclude = EnumTaxExclude.Exclude
    Private mtax As Decimal = 0
    Private mTaxExchangeRate As Decimal = 1
    Private mTaxAmount As Decimal = 0
    Private mTotal As Decimal = 0
    Private mExhangeRateToIDR As Decimal = 1
    Private mExchangeRateToUSD As Decimal = 1
    Private mRemarks As String = ""
    Private mSuppAddressIdentifier As String = ""
    Private mPrimaryPO4Dp As Double = -1
    Private mPaymentTerms As Double = 0
    Private mNoFakturPajak As String = ""
    Private mNoNPWPSupplier As String = ""
    Private mVATDiPungut As Boolean = True
    Private mEFak_KodeJenisTransaksiPajak As Integer = -1
    Private mSupplierInvoiceDate As Date = Now
 
    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now

    'elemen yang digunakan untuk antrian simpan data
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianStockInOut_DO_Invoice
    Dim ObjRnd As New Random
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter
    Dim ObjAntrin As New Core_Module.ClsAntrian
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim SudahHapusAntrian As Boolean = True

    '====================================================================================================================================================================
    '====================================================================================================================================================================
    'Erna's Code mulai()
    '====================================================================================================================================================================
    Private mPrimaryAccount As Double = -1
    Private mPrimaryAccountDP As Double = -1
    Private mPrimaryCurrTax As Double = -1

    Dim ObjNumbering As New Core_Module.ClsMasterTransactionNumbering


    '2025 Core Tax Tahap 2.2
    Private mDPPNilaiLain As Decimal = 0

    '2025 Core Tax Tahap 2.2
    Public Property DPPNilaiLain() As Decimal
        Get
            Return Me.mDPPNilaiLain
        End Get
        Set(ByVal value As Decimal)
            Me.mDPPNilaiLain = value
        End Set
    End Property
    '2025 Core Tax Tahap 2.2 End

    'aul add
    Public Enum EnumTipeDP As Integer
        AOPDP
        ManualDP
    End Enum

    Public Property PrimaryAccount() As Double
        Get
            Return Me.mPrimaryAccount
        End Get
        Set(ByVal value As Double)
            Me.mPrimaryAccount = value
        End Set
    End Property

    Public Property PrimaryCurrTax() As Double
        Get
            Return Me.mPrimaryCurrTax
        End Get
        Set(ByVal value As Double)
            Me.mPrimaryCurrTax = value
        End Set
    End Property

    Public Property PrimaryAccountDP() As Double
        Get
            Return Me.mPrimaryAccountDP
        End Get
        Set(ByVal value As Double)
            Me.mPrimaryAccountDP = value
        End Set
    End Property

    Public Property SupplierInvoiceDate() As Date
        Get
            Return Me.mSupplierInvoiceDate
        End Get
        Set(ByVal value As Date)
            Me.mSupplierInvoiceDate = value
        End Set
    End Property

    Public Function GetPrimaryAccountAP(ByVal vPrimaryAOP As Decimal) As Double
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aop3TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetPrimaryAccount(vPrimaryAOP).HasValue AndAlso DTA.ScalarQuery_GetPrimaryAccount(vPrimaryAOP).Value > 0 Then
                Return DTA.ScalarQuery_GetPrimaryAccount(vPrimaryAOP).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            DTA.Dispose()
        End Try
    End Function


    Public Function GetPrimaryAccountDP(ByVal vPrimaryAOP As Decimal) As Double
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aop3TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetPrimaryAccountDP(vPrimaryAOP).HasValue AndAlso DTA.ScalarQuery_GetPrimaryAccountDP(vPrimaryAOP).Value > 0 Then
                Return DTA.ScalarQuery_GetPrimaryAccountDP(vPrimaryAOP).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryCurrTax(ByVal vPrimaryAOP As Decimal) As Double
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aop3TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetPrimaryCurrTax(vPrimaryAOP).HasValue AndAlso DTA.ScalarQuery_GetPrimaryCurrTax(vPrimaryAOP).Value > 0 Then
                Return DTA.ScalarQuery_GetPrimaryCurrTax(vPrimaryAOP).Value
            Else
                Return -1
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function CariDetail(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aop2TableAdapter
        Dim TBL As New DseTransaksiAOP.tbltrans_aop2DataTable
        Dim row As DseTransaksiAOP.tbltrans_aop2Row

        Try
            'fil basis data
            If DTA.FillBy_PrimaryTransaksi(TBL, vPrimaryTransaksi) > 0 Then
                'salin data header
                row = TBL.Rows(0)

                With row

                    Me.mnotransaksi = .No_Transaksi
                    Me.mNoUrut = .Urutan
                    Me.mprimary = .Primary
                    mCalculatesBy = .Calculates_By
                    mPrimaryCurrency = .Primary_Currency
                    mSubtotalBeforeDP = .Sub_Total_Before_DP
                    mDownPayment = .Down_Payment
                    Me.mSubTotal = .Sub_Total

                    If .Tax_Exclude = True Then
                        Me.mTaxExclude = EnumTaxExclude.Exclude
                    Else
                        Me.mTaxExclude = EnumTaxExclude.Include
                    End If
                    mTaxExchangeRate = .Tax_Exchange_Rate
                    Me.mTaxAmount = .Tax_Amount
                    mDutyStamp = .Duty_Stamp
                    Me.mTotal = .Total
                    mExhangeRateToIDR = .Exchange_Rate_To_IDR
                    mExchangeRateToUSD = .Exchange_Rate_To_USD
                    Me.mRemarks = .Remarks
                    Me.mTanggalTransaksi = .Tanggal_Transaksi
                    mTipeAOP = .Tipe_AOP
                    mPrimarySupplier = .Primary_Supplier
                    mSupplierInvoiceNo = .Supplier_Invoice_No
                    mSupplierInvoiceDate = .Supplier_Invoice_Date
                    mLokal = .Lokal
                    mDownPayment = .Down_Payment
                    mSubTotal = .Sub_Total
                    mtax = .Tax
                    mPrimaryPO4Dp = .Primary_PO_4_DP
                    Me.mPaymentTerms = .Payment_Term
                    Me.NoFakturPajak = ""
                    Me.mNoNPWPSupplier = .No_NPWP
                    Me.mNoFakturPajak = .Tax_Invoice_Number
                    If .VAT_Dipungut = 1 Then
                        Me.mVATDiPungut = True
                    Else
                        Me.mVATDiPungut = False
                    End If
                    Me.mEFak_KodeJenisTransaksiPajak = .kode_jenis_transaksi_pajak

                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
                    Me.mPrimaryAccount = .primary_akun_ap
                    Me.mPrimaryCurrTax = .Primary_Curr_Tax
                    Me.mPrimaryAccountDP = .Primary_Akun_DP

                    '2025 Core Tax Tahap 2.2
                    Me.mDPPNilaiLain = .DPPNilaiLain
                End With
                'nilai balikan
                Return True
            Else
                Me.ResetObject()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Overridable Function CariDetail(ByVal vPrimaryTransaksi As Double, _
   ByVal vDgvItem As DataGridView, ByVal vdGVItemDP As DataGridView, ByVal vDGVItemPR As DataGridView, _
   ByVal vDGVItemExchangeRate As DataGridView, ByVal vDGVItemCostData As DataGridView, ByVal vFrm As FrmTransaksiAOP2) As Boolean
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aop2TableAdapter
        Dim TBL As New DseTransaksiAOP.tbltrans_aop2DataTable

        Dim row As DseTransaksiAOP.tbltrans_aop2Row
        Dim ObjItem As New ClsTransaksiAOPItem

        Try
            'fil basis data
            If DTA.FillBy_PrimaryTransaksi(TBL, vPrimaryTransaksi) > 0 Then
                'salin data header
                row = TBL.Rows(0)

                With row
                    mprimary = .Primary
                    mnotransaksi = .No_Transaksi
                    mNoUrut = .Urutan
                    mTipeAOP = .Tipe_AOP
                    mTanggalTransaksi = .Tanggal_Transaksi
                    mPrimarySupplier = .Primary_Supplier
                    mSupplierInvoiceNo = .Supplier_Invoice_No
                    mSupplierInvoiceDate = .Supplier_Invoice_Date
                    mLokal = .Lokal
                    mPrimaryCurrency = .Primary_Currency
                    mSubtotalBeforeDP = .Sub_Total_Before_DP
                    mCalculatesBy = .Calculates_By
                    mDownPayment = .Down_Payment
                    mSubTotal = .Sub_Total
                    mDutyStamp = .Duty_Stamp
                    If .Tax_Exclude = True Then
                        Me.mTaxExclude = EnumTaxExclude.Exclude
                    Else
                        Me.mTaxExclude = EnumTaxExclude.Include
                    End If
                    mtax = .Tax
                    mTaxExchangeRate = .Tax_Exchange_Rate
                    mTaxAmount = .Tax_Amount
                    mTotal = .Total
                    mExhangeRateToIDR = .Exchange_Rate_To_IDR
                    mExchangeRateToUSD = .Exchange_Rate_To_USD
                    mRemarks = .Remarks
                    mSuppAddressIdentifier = .Supplier_Address_Identifier
                    mPrimaryPO4Dp = .Primary_PO_4_DP
                    Me.mPaymentTerms = .Payment_Term
                    Me.mNoNPWPSupplier = .No_NPWP
                    Me.mNoFakturPajak = .Tax_Invoice_Number
                    If .VAT_Dipungut = 1 Then
                        Me.mVATDiPungut = True
                    Else
                        Me.mVATDiPungut = False
                    End If
                    Me.mEFak_KodeJenisTransaksiPajak = .kode_jenis_transaksi_pajak

                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
                    Me.mPrimaryCurrTax = .Primary_Curr_Tax
                    Me.mPrimaryAccount = .primary_akun_ap
                    Me.mPrimaryAccountDP = .Primary_Akun_DP

                    '2025 Core Tax Tahap 2.2
                    Me.mDPPNilaiLain = .DPPNilaiLain
                End With
                'salin data item 
                ObjItem.IsiDGV_Multi(vDgvItem, vDGVItemExchangeRate, vdGVItemDP, Me.mprimary, _
                Me.mCalculatesBy, vDGVItemPR, vDGVItemCostData, vFrm)
                'nilai balikan
                Return True
            Else
                Me.ResetObject(vDgvItem, vdGVItemDP, vDGVItemExchangeRate, vDGVItemPR, vDGVItemCostData)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Overridable Function CariDetail(ByVal vPrimaryTransaksi As Double, _
            ByVal vDgvItem As DataGridView, ByVal vdGVItemDP As DataGridView, ByVal vDGVItemPR As DataGridView, _
            ByVal vDGVItemExchangeRate As DataGridView, ByVal vDGVItemCostData As DataGridView, ByVal vFrm As FrmTransaksiAOP_Quick2, _
            ByVal vTblCustom As CustomSMESS.DseTransaksiAOP_Quick2_AddByCustomProduct.tbltrans_custom_aop_item1DataTable, _
            ByRef vTblCustomInOUt As CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable) As Boolean
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aop2TableAdapter
        Dim TBL As New DseTransaksiAOP.tbltrans_aop2DataTable

        Dim row As DseTransaksiAOP.tbltrans_aop2Row
        Dim ObjItem As New ClsTransaksiAOPItem

        Try
            'fil basis data
            If DTA.FillBy_PrimaryTransaksi(TBL, vPrimaryTransaksi) > 0 Then
                'salin data header
                row = TBL.Rows(0)

                With row
                    mprimary = .Primary
                    mnotransaksi = .No_Transaksi
                    mNoUrut = .Urutan
                    mTipeAOP = .Tipe_AOP
                    mTanggalTransaksi = .Tanggal_Transaksi
                    mPrimarySupplier = .Primary_Supplier
                    mSupplierInvoiceNo = .Supplier_Invoice_No
                    mSupplierInvoiceDate = .Supplier_Invoice_Date
                    mLokal = .Lokal
                    mPrimaryCurrency = .Primary_Currency
                    mSubtotalBeforeDP = .Sub_Total_Before_DP
                    mCalculatesBy = .Calculates_By
                    mDownPayment = .Down_Payment
                    mSubTotal = .Sub_Total
                    mDutyStamp = .Duty_Stamp
                    If .Tax_Exclude = True Then
                        Me.mTaxExclude = EnumTaxExclude.Exclude
                    Else
                        Me.mTaxExclude = EnumTaxExclude.Include
                    End If
                    mtax = .Tax
                    mTaxExchangeRate = .Tax_Exchange_Rate
                    mTaxAmount = .Tax_Amount
                    mTotal = .Total
                    mExhangeRateToIDR = .Exchange_Rate_To_IDR
                    mExchangeRateToUSD = .Exchange_Rate_To_USD
                    mRemarks = .Remarks
                    mSuppAddressIdentifier = .Supplier_Address_Identifier
                    mPrimaryPO4Dp = .Primary_PO_4_DP
                    Me.mPaymentTerms = .Payment_Term
                    Me.mNoNPWPSupplier = .No_NPWP
                    Me.mNoFakturPajak = .Tax_Invoice_Number
                    If .VAT_Dipungut = 1 Then
                        Me.mVATDiPungut = True
                    Else
                        Me.mVATDiPungut = False
                    End If
                    Me.mEFak_KodeJenisTransaksiPajak = .kode_jenis_transaksi_pajak

                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
                    Me.mPrimaryCurrTax = .Primary_Curr_Tax
                    Me.mPrimaryAccount = .primary_akun_ap
                    Me.mPrimaryAccountDP = .Primary_Akun_DP
                End With
                'salin data item 
                ObjItem.IsiDGV_Multi(vDgvItem, vDGVItemExchangeRate, vdGVItemDP, Me.mprimary, _
                Me.mCalculatesBy, vDGVItemPR, vDGVItemCostData, vFrm)

                Dim ObjCustom As New CustomSMESS.ClsTransaksiAOP_Quick2_AddByCustomProduct
                ObjCustom.IsiTblDummyFromDB(Me.mprimary, vTblCustom, vTblCustomInOUt)
                'nilai balikan
                Return True
            Else
                Me.ResetObject(vDgvItem, vdGVItemDP, vDGVItemExchangeRate, vDGVItemPR, vDGVItemCostData)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Overridable Function SimpanBaru(ByVal vUserCode As String, Optional ByVal vTransJurnal As Boolean = True, _
    Optional ByVal vCekDataBenar As Boolean = True, Optional ByVal vFromCutOff As Boolean = False) As Boolean
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aopTableAdapter
        Dim DTA2 As New DseTransaksiAOPTableAdapters.tbltrans_aop2TableAdapter
        Dim ObjItem As New ClsTransaksiAOPItem
        Dim UrutanFakturPajak As String = ""
        Dim TaxExclude As Integer
        Dim ObjCurrency As New Core_Module.clsMasterCurrency
        Dim vVATDipungut As Integer = 1
        Dim ObjPOItem As New ClsTransaksiPOItem
        Dim ObjPO As New ClsTransaksiPO
        Dim ObjCD As New Core_Module.ClsCekData

        Dim CreatedBy_ As String = vUserCode
        Dim CreatedDate_ As Date = Now
        Dim ModifiedBy_ As String = vUserCode
        Dim ModifiedDate_ As Date = Now
        Dim TimeStamp_ As Date = Now

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA2.Connection.ConnectionString = DTA.Connection.ConnectionString

            If vCekDataBenar = True Then
                'khusus untuk dp ganti no tergantung dengan suppllier-nya
                Me.mnotransaksi = Me.GetNoBaru(Me.mTanggalTransaksi, Me.mPrimarySupplier)
            End If

            If Me.ApakahDataBenar(True, "", -1, vCekDataBenar) = True Then
                'pastikan po sudah di otorisasi dengan baik
                If vCekDataBenar = True Then
                    If ObjPO.ApakahSudahOtorisasi(Me.mPrimaryPO4Dp) = False Then
                        MsgBox("Please authorize po first, you can't create purchase downpayment for unauthorized po.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If

                If Me.mTaxExclude = EnumTaxExclude.Exclude Then
                    TaxExclude = 1
                Else
                    TaxExclude = 0
                End If
                If Me.VATDipungut = True Then
                    vVATDipungut = 1
                Else
                    vVATDipungut = 0
                End If

                If vFromCutOff = True Then
                    CreatedBy_ = Me.mCreatedBy
                    CreatedDate_ = Me.mCreatedDate
                    ModifiedBy_ = Me.mModifiedBy
                    ModifiedDate_ = Me.mModifiedDate
                    TimeStamp_ = Me.mTimestamp
                End If

                'insert query
                If DTA2.InsertQuery(Me.mnotransaksi, CInt(Me.mnotransaksi.Substring(0, 6)), _
                 Me.mTipeAOP, Me.mTanggalTransaksi, Me.mPrimarySupplier, Me.mSupplierInvoiceNo, Me.mLokal, _
                 Me.mPrimaryCurrency, Me.mCalculatesBy, Me.SubTotal, Me.mDownPayment, Me.mSubTotal, Me.mDutyStamp, TaxExclude, Me.Tax, _
                 Me.mTaxExchangeRate, Me.mTaxAmount, Me.mTotal, Me.mExchangeRateToUSD, Me.mExhangeRateToIDR, _
                 Me.mRemarks, CreatedBy_, CreatedDate_, ModifiedDate_, ModifiedBy_, TimeStamp_, Me.mSuppAddressIdentifier, Me.mPrimaryPO4Dp, Me.mPaymentTerms, _
                 Me.mNoFakturPajak, vVATDipungut, Me.mNoNPWPSupplier, Me.mEFak_KodeJenisTransaksiPajak, _
                 Me.mPrimaryCurrTax, Me.mPrimaryAccount, Me.mPrimaryAccountDP, Me.SupplierInvoiceDate, Me.mDPPNilaiLain) > 0 Then
                    'update data yang berubah setelah insert query
                    Me.mprimary = DTA.ScalarQuery_GetPrimaryByNoTransaksi(Me.mnotransaksi).Value
                    Me.mNoUrut = CInt(Me.mnotransaksi.Substring(0, 6))
                    Me.mTimestamp = DTA.ScalarQuery_GetTimeStamp(Me.mprimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    If vCekDataBenar = True Then
                        'cek penomoran agar tidak ada yang sama tetapi beda bulan
                        Me.CekNomor(DTA)
                    End If

                    '**************************************************************************************************************************************
                    '**************************************************************************************************************************************
                    'PERUBAHAN SISTEM PERHITUNGAN JURNAL ERROR NO 50 TANGGAL 11 JULI 2016 -- PERBAIKAN 28 JULI 2016
                    'SEBELUMNYA NILAI SALES RUPIAH = NILAI SALES USD * EXCHANGE RATE
                    'SEKARANG NILAINYA DICEK DULU APAKAH ADA PENJUMLAHAN TERHADAP NILAI SALES (CONTOH TAX USD, SHIPMENT DLL),
                    'HASILNYA DITOTAL LALU DIKURANGI SEMUA PENJUMLAHAN, SISANYA DIANGGAP SEBAGAI SALES
                    'HAL INI DILAKUKAN UTK MENCEGAH SELISIH ANTARA SALES DI INVOICE JURNAL DENGAN SALES DI RECEIVEMENT JOURNAL

                    Dim PurchaseRupiah As Decimal = 0
                    Dim TotalPurchase As Decimal = 0
                    Dim AmountTanpaPurchase As Decimal = 0
                    Dim DefprecBase As Integer = 0
                    DefprecBase = ObjCurrency.GetDecimal_DefaultCurrency
                    TotalPurchase = ObjCD.DigitGrouping(Me.mTotal * Me.mTaxExchangeRate, DefprecBase)
                    If Me.mPrimaryCurrTax = Me.mPrimaryCurrency Then
                        AmountTanpaPurchase += ObjCD.DigitGrouping(Me.mTaxAmount * Me.mTaxExchangeRate, DefprecBase)
                        AmountTanpaPurchase += ObjCD.DigitGrouping(Me.mDutyStamp * Me.mTaxExchangeRate, DefprecBase)
                    End If
                    'AmountTanpaSalesRp = Me.ObjCD.DigitGrouping((AmountTanpaSales * Me.mExchangeRate), DefPrecBase)
                    PurchaseRupiah = ObjCD.DigitGrouping(TotalPurchase - AmountTanpaPurchase, DefprecBase)

                    '**************************************************************************************************************************************
                    'simpan jurnal otomatis
                    Me.TransferJurnal_DP1(Me.mprimary, Me.mPrimarySupplier, _
                    PurchaseRupiah, _
                    ObjCD.DigitGrouping(Me.mTaxAmount * Me.mTaxExchangeRate, DefprecBase), _
                    ObjCD.DigitGrouping(Me.mDutyStamp * Me.mTaxExchangeRate, DefprecBase), _
                    Me.mNoUrut, Me.mTanggalTransaksi, Me.mnotransaksi, Me.mModifiedBy, CDbl(Me.mPaymentTerms), _
                    Me.mSupplierInvoiceNo, Me.mSubtotalBeforeDP, Me.mTaxExchangeRate, Me.mPrimaryCurrency, _
                    Me.mPrimaryAccount, Me.mPrimaryCurrTax, Me.mTaxAmount, Me.mDutyStamp, Me.mPrimaryAccountDP)
                    'update po balance, untuk memastikan po tidak bisa dicancel otorisasi
                    ObjPOItem.SetOrderBalance(Me.mPrimaryPO4Dp)
                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mnotransaksi + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                Me.mnotransaksi = Me.GetNoBaru(Me.mTanggalTransaksi, Me.mPrimarySupplier)
                Return Me.SimpanBaru(vUserCode, vTransJurnal, vCekDataBenar, vFromCutOff)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
            ObjCurrency = Nothing
            ObjPOItem = Nothing
            ObjPO = Nothing
        End Try
    End Function

    Public Overridable Function SimpanBaru1(ByVal vdgvItem As DataGridView, ByVal vdgvItemDP As DataGridView, _
    ByVal vDGVItemSR As DataGridView, ByVal vDGVItemExchangeRate As DataGridView, ByVal vUserCode As String, _
    ByVal vDGVItemCostData As DataGridView, ByVal vFrm As FrmTransaksiAOP2, _
    Optional ByVal vCekDataBenar As Boolean = True, Optional ByVal vFromCutOff As Boolean = False) As Boolean
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aopTableAdapter
        Dim DTA2 As New DseTransaksiAOPTableAdapters.tbltrans_aop2TableAdapter
        Dim ObjItem As New ClsTransaksiAOPItem
        Dim UrutanFakturPajak As String = ""
        Dim TaxExclude As Integer
        Dim TotalCOst As Decimal = 0
        Dim rowdgv As DataGridViewRow
        Dim ObjCurr As New Core_Module.clsMasterCurrency
        Dim DPOriginal, DPBaseCurrency As Decimal
        Dim DTA1 As New DseTransaksiAOPTableAdapters.DataTable15TableAdapter
        Dim TBL1 As New DseTransaksiAOP.DataTable15DataTable
        Dim row1 As DseTransaksiAOP.DataTable15Row
        Dim DefPrec As Integer = 0
        Dim vVATDipungut As Integer = 0

        Dim objCD As New Core_Module.ClsCekData
        'Dim Primary4DP As Double = -1

        Dim CreatedBy_ As String = vUserCode
        Dim CreatedDate_ As Date = Now
        Dim ModifiedBy_ As String = vUserCode
        Dim ModifiedDate_ As Date = Now
        Dim TimeStamp_ As Date = Now

        Try
            If Me.ApakahDataBenar1(True, vdgvItem, vFrm, vCekDataBenar) = True Then
                DefPrec = ObjCurr.GetDecimal_DefaultCurrency
                'insert query
                If Me.mTaxExclude = EnumTaxExclude.Exclude Then
                    TaxExclude = 1
                Else
                    TaxExclude = 0
                End If
                If Me.mVATDiPungut = True Then
                    vVATDipungut = 1
                Else
                    vVATDipungut = 0
                End If

                If vFromCutOff = True Then
                    CreatedBy_ = Me.mCreatedBy
                    CreatedDate_ = Me.mCreatedDate
                    ModifiedBy_ = Me.mModifiedBy
                    ModifiedDate_ = Me.mModifiedDate
                    TimeStamp_ = Me.mTimestamp
                End If

                Me.mNoUrut = CInt(Me.mnotransaksi.Substring(0, 6))
                If DTA2.InsertQuery(Me.mnotransaksi, CInt(Me.mnotransaksi.Substring(0, 6)), _
                 Me.mTipeAOP, Me.mTanggalTransaksi, Me.mPrimarySupplier, Me.mSupplierInvoiceNo, Me.mLokal, _
                 Me.mPrimaryCurrency, Me.mCalculatesBy, Me.SubtotalBeforeDP, Me.mDownPayment, Me.mSubTotal, Me.mDutyStamp, TaxExclude, Me.Tax, _
                 Me.mTaxExchangeRate, Me.mTaxAmount, Me.mTotal, Me.mExchangeRateToUSD, Me.mExhangeRateToIDR, _
                 Me.mRemarks, CreatedBy_, CreatedDate_, ModifiedDate_, ModifiedBy_, TimeStamp_, Me.mSuppAddressIdentifier, Me.mPrimaryPO4Dp, Me.mPaymentTerms, _
                 Me.mNoFakturPajak, vVATDipungut, Me.mNoNPWPSupplier, Me.mEFak_KodeJenisTransaksiPajak, _
                 Me.mPrimaryCurrTax, Me.mPrimaryAccount, Me.mPrimaryAccountDP, Me.mSupplierInvoiceDate, Me.mDPPNilaiLain) > 0 Then
                    'update data yang berubah setelah insert query
                    Me.mprimary = Me.GetPrimary(Me.NoTransaksi)
                    Me.mTimestamp = DTA.ScalarQuery_GetTimeStamp(Me.mprimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    'update data yang berkaitan dengan item exchange rate
                    If Me.mCalculatesBy = EnumCalculatesBy.AOP Then
                        vDGVItemExchangeRate.Rows(0).Cells(vFrm.ColPrimaryDokumen.Index).Value = Me.mprimary
                        vDGVItemExchangeRate.Rows(0).Cells(vFrm.ColDocumentNo.Index).Value = Me.mnotransaksi
                    End If
                    'simpan data item baru
                    ObjItem.Simpan_Multi(vdgvItem, vDGVItemExchangeRate, vdgvItemDP, Me.mprimary, vDGVItemSR, vDGVItemCostData, vFrm)
                    If vFromCutOff = False Then
                        Me.CekNomor(DTA)
                    End If

                    'hitung DP jika ada
                    DTA1.Fill(TBL1, Me.mprimary)
                    DPOriginal = 0
                    DPBaseCurrency = 0
                    Dim PrimaryAkunDP As Double = -1
                    For Each row1 In TBL1.Rows
                        DPOriginal += objCD.DigitGrouping(row1.Sub_Total_Before_DP, ObjCurr.GetDefaultPrecission(Me.mPrimaryCurrency))
                        DPBaseCurrency += objCD.DigitGrouping(row1.Sub_Total_Before_DP * row1.Tax_Exchange_Rate, DefPrec)
                        PrimaryAkunDP = row1.Primary_Akun_DP
                    Next

                    'simpan jurnal otomatis
                    For Each rowdgv In vDGVItemCostData.Rows
                        Dim PrimaryCurrCost As Double = -1
                        PrimaryCurrCost = CDbl(rowdgv.Cells(vFrm.ColPrimaryCurrencyCost.Index).Value)
                        TotalCOst += objCD.DigitGrouping(CDec(rowdgv.Cells(vFrm.ColCostAmount.Index).Value), DefPrec)
                    Next

                    '**************************************************************************************************************************************
                    '**************************************************************************************************************************************
                    'PERUBAHAN SISTEM PERHITUNGAN JURNAL ERROR NO 50 TANGGAL 11 JULI 2016 -- PERBAIKAN 28 JULI 2016
                    'SEBELUMNYA NILAI SALES RUPIAH = NILAI SALES USD * EXCHANGE RATE
                    'SEKARANG NILAINYA DICEK DULU APAKAH ADA PENJUMLAHAN TERHADAP NILAI SALES (CONTOH TAX USD, SHIPMENT DLL),
                    'HASILNYA DITOTAL LALU DIKURANGI SEMUA PENJUMLAHAN, SISANYA DIANGGAP SEBAGAI SALES
                    'HAL INI DILAKUKAN UTK MENCEGAH SELISIH ANTARA SALES DI INVOICE JURNAL DENGAN SALES DI RECEIVEMENT JOURNAL

                    Dim DPRupiah As Decimal = 0
                    Dim PurchaseRupiah As Decimal = 0
                    Dim TotalPurchase As Decimal = 0
                    Dim AmountTanpaPurchase As Decimal = 0
                    Dim DefprecBase As Integer = 0
                    Dim DefprecOri As Integer = 0
                    Dim PurchasePlusDP As Decimal = 0


                    DefprecBase = ObjCurr.GetDecimal_DefaultCurrency
                    DefprecOri = ObjCurr.GetDefaultPrecission(Me.mPrimaryCurrency)
                    If IsNumeric(Me.mTotal) Then
                        PurchasePlusDP += objCD.DigitGrouping(CDec(Me.mTotal), DefprecOri)
                    End If
                    If IsNumeric(DPOriginal) Then
                        PurchasePlusDP += objCD.DigitGrouping(CDec(DPOriginal), DefprecOri)
                    End If

                    PurchasePlusDP = objCD.DigitGrouping(PurchasePlusDP * Me.mTaxExchangeRate, DefprecBase)

                    If Me.mPrimaryCurrTax = Me.mPrimaryCurrency Then
                        AmountTanpaPurchase += objCD.DigitGrouping(Me.mTaxAmount * Me.mTaxExchangeRate, DefprecBase)
                        AmountTanpaPurchase += objCD.DigitGrouping(Me.mDutyStamp * Me.mTaxExchangeRate, DefprecBase)
                    End If
                    'AmountTanpaSalesRp = Me.ObjCD.DigitGrouping((AmountTanpaSales * Me.mExchangeRate), DefPrecBase)
                    PurchaseRupiah = objCD.DigitGrouping(PurchasePlusDP - AmountTanpaPurchase + TotalCOst, DefprecBase)
                    '**************************************************************************************************************************************

                    Me.TransferJurnal1(Me.mprimary, Me.mPrimarySupplier, PurchaseRupiah, _
                    objCD.DigitGrouping(Me.mTaxAmount * Me.mTaxExchangeRate, DefprecBase), _
                    objCD.DigitGrouping(Me.mDutyStamp * Me.mTaxExchangeRate, DefprecBase), _
                    Me.mNoUrut, Me.mTanggalTransaksi, Me.mnotransaksi, Me.mModifiedBy, CDbl(Me.mPaymentTerms), TotalCOst, vDGVItemCostData, _
                    vFrm, Me.mSupplierInvoiceNo, Me.mSubtotalBeforeDP, Me.mTaxExchangeRate, _
                    Me.mPrimaryCurrency, DPOriginal, DPBaseCurrency, Me.mPrimaryAccount, Me.mPrimaryCurrTax, Me.mTaxAmount, Me.mDutyStamp, PrimaryAkunDP)

                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mnotransaksi + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                Me.mnotransaksi = Me.GetNoBaru(Me.mTanggalTransaksi, Me.mPrimarySupplier)
                Return Me.SimpanBaru1(vdgvItem, vdgvItemDP, vDGVItemSR, vDGVItemExchangeRate, vUserCode, vDGVItemCostData, vFrm, vCekDataBenar, vFromCutOff)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
            ObjCurr = Nothing
            DTA1.Dispose()
            TBL1.Dispose()
        End Try
    End Function

    Public Overridable Function SimpanBaru1(ByVal vdgvItem As DataGridView, ByVal vdgvItemDP As DataGridView, _
    ByVal vDGVItemSR As DataGridView, ByVal vDGVItemExchangeRate As DataGridView, ByVal vUserCode As String, _
    ByVal vDGVItemCostData As DataGridView, ByVal vFrm As FrmTransaksiAOP_Quick2, _
    ByVal vTblCustom As CustomSMESS.DseTransaksiAOP_Quick2_AddByCustomProduct.tbltrans_custom_aop_item1DataTable, _
    ByVal vTblItemCustomInOut As CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable, _
       Optional ByVal vCekDataBenar As Boolean = True, _
    Optional ByVal NoPO As String = "", Optional ByVal NoStockIn As String = "", Optional ByVal NoGRN As String = "", _
    Optional ByVal vKoreksiStock As Boolean = True, Optional ByVal vSetPOB As Boolean = True, _
    Optional ByVal vFromCutOff As Boolean = False) As Boolean
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aopTableAdapter
        Dim DTA2 As New DseTransaksiAOPTableAdapters.tbltrans_aop2TableAdapter
        Dim ObjItem As New ClsTransaksiAOPItem
        Dim UrutanFakturPajak As String = ""
        Dim TaxExclude As Integer
        Dim TotalCOst As Decimal = 0
        Dim rowdgv As DataGridViewRow
        Dim ObjPO As New ClsTransaksiPO
        Dim ObjStockInOut As New Inventory_Module.clsTransaksiStockInOut
        Dim ObjSupplier As New Core_Module.clsMasterSupplier
        Dim AddIdentifier As String = ""
        Dim PrimaryPO As Double = -1
        Dim PrimaryGRN As Double = -1
        Dim PrimaryStockIn As Double = -1
        Dim ObjUser As New Core_Module.ClsUser
        Dim PrimaryUser As Double = -1
        Dim ObjGRN As New ClsTransaksiGRN
        Dim Lanjut As Boolean = True
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim ObjPOItem As New ClsTransaksiPOItem
        Dim ListPrimaryPO As New List(Of Double)
        Dim vVATDipungut As Integer = 0
        Dim DefprecBase As Integer = 0
        DefprecBase = ObjCur.GetDecimal_DefaultCurrency

        Dim ObjCD As New Core_Module.ClsCekData

        Dim CreatedBy_ As String = vUserCode
        Dim CreatedDate_ As Date = Now
        Dim ModifiedBy_ As String = vUserCode
        Dim ModifiedDate_ As Date = Now
        Dim TimeStamp_ As Date = Now
        Dim ObjSC As New Inventory_Module.ClsTransaksiStockCard

        'by custom admiral
        Dim tblProdQty As New CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_productqtyDataTable
        Dim ObjCustAOP As New CustomSMESS.ClsTransaksiAOP
        Dim ListProductCustom As New List(Of Double)
        Dim rowCustomInOut As CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemRow
        'Dim Primary4DP As Double = -1


        Try
            If vCekDataBenar = True Then
                If Me.ApakahDataBenar(True, NoPO, -1) = False Then
                    Lanjut = False
                End If
            End If
            If Me.mVATDiPungut = True Then
                vVATDipungut = 1
            Else
                vVATDipungut = 0
            End If

            If Lanjut = True Then
                'simpan data po, stock in , grn
                AddIdentifier = ObjSupplier.GetIdentifierAdd(Me.mPrimarySupplier)
                PrimaryUser = ObjUser.GetPrimary(vUserCode)

                'simpan data po
                If Me.ApakahPerluSimpanPO(vFrm, vdgvItem) = True Then
                    PrimaryPO = Me.SimpanPO_Multi(vdgvItem, vFrm, vUserCode, AddIdentifier, ObjPO, PrimaryUser, NoPO, vCekDataBenar, vFromCutOff)
                    If PrimaryPO = -1 Then
                        Return False
                    End If
                End If

                For Each rowCustomInOut In vTblItemCustomInOut.Rows
                    rowCustomInOut.Primary_PO_Purchase = PrimaryPO
                Next

                'simpan data stock in otomatis grn disimpan
                PrimaryStockIn = Me.SimpanStockIn_Multi(vUserCode, vdgvItem, vFrm, PrimaryPO, AddIdentifier, ObjStockInOut, ObjPO.NoTransaksi, vTblItemCustomInOut, NoStockIn, vCekDataBenar, vKoreksiStock, vFromCutOff)
                If PrimaryStockIn = -1 Then
                    ObjPO.CancelAuthorize(PrimaryPO, PrimaryUser, vUserCode)
                    ObjPO.Hapus1(PrimaryPO, vUserCode, False)
                    Return False
                End If
                PrimaryGRN = ObjGRN.GetPrimaryGRNByPrimaryStockIn(PrimaryStockIn)
                If PrimaryGRN = -1 Then
                    ObjPO.CancelAuthorize(PrimaryPO, PrimaryUser, vUserCode)
                    ObjPO.Hapus(PrimaryPO, vUserCode, False)
                    Return False
                End If

                'simpan data orderbalance untuk po yang bukan dari quick aop
                For Each rowdgv In vFrm.dgvItem.Rows
                    If rowdgv.Cells(vFrm.ColReferedToPO.Index).Value = 1 Then
                        If ListPrimaryPO.Contains(rowdgv.Cells(vFrm.ColPrimaryPO.Index).Value) = False Then
                            ListPrimaryPO.Add(rowdgv.Cells(vFrm.ColPrimaryPO.Index).Value)
                        End If
                    End If
                Next
                ObjPOItem.SetOrderBalance(ListPrimaryPO)

                'insert query
                If Me.mTaxExclude = EnumTaxExclude.Exclude Then
                    TaxExclude = 1
                Else
                    TaxExclude = 0
                End If
                If Me.ApakahSudahAda(Me.mnotransaksi) = True Then
                    Me.mnotransaksi = Me.GetNoBaru(Me.mTanggalTransaksi, Me.mPrimarySupplier)
                End If
                Me.mNoUrut = CInt(Me.mnotransaksi.Substring(0, 6))
                If DTA2.InsertQuery(Me.mnotransaksi, CInt(Me.mnotransaksi.Substring(0, 6)), _
                 Me.mTipeAOP, Me.mTanggalTransaksi, Me.mPrimarySupplier, Me.mSupplierInvoiceNo, Me.mLokal, _
                 Me.mPrimaryCurrency, Me.mCalculatesBy, Me.SubTotal, Me.mDownPayment, Me.mSubTotal, Me.mDutyStamp, TaxExclude, Me.Tax, _
                 Me.mTaxExchangeRate, Me.mTaxAmount, Me.mTotal, Me.mExchangeRateToUSD, Me.mExhangeRateToIDR, _
                 Me.mRemarks, CreatedBy_, CreatedDate_, ModifiedDate_, ModifiedBy_, TimeStamp_, Me.mSuppAddressIdentifier, Me.mPrimaryPO4Dp, Me.mPaymentTerms, _
                 Me.mNoFakturPajak, vVATDipungut, Me.mNoNPWPSupplier, Me.mEFak_KodeJenisTransaksiPajak, _
                 Me.mPrimaryCurrTax, Me.mPrimaryAccount, Me.mPrimaryAccountDP, Me.mSupplierInvoiceDate, Me.mDPPNilaiLain) > 0 Then
                    'update data yang berubah setelah insert query
                    Me.mprimary = Me.GetPrimary(Me.NoTransaksi)
                    Me.mTimestamp = DTA.ScalarQuery_GetTimeStamp(Me.mprimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode
                    'update data quick purchase
                    DTA.UpdateQuery_SetQuickPurchase(Me.mprimary)
                    'update data yang berkaitan dengan item exchange rate
                    If Me.mCalculatesBy = EnumCalculatesBy.AOP Then
                        vDGVItemExchangeRate.Rows(0).Cells(vFrm.ColPrimaryDokumen.Index).Value = Me.mprimary
                        vDGVItemExchangeRate.Rows(0).Cells(vFrm.ColDocumentNo.Index).Value = Me.mnotransaksi
                    End If
                    'simpan data item baru
                    For Each rowdgv In vdgvItem.Rows
                        If rowdgv.Cells(vFrm.ColReferedToPO.Index).Value = 0 Then
                            rowdgv.Cells(vFrm.ColPrimaryPO.Index).Value = PrimaryPO
                        End If
                        rowdgv.Cells(vFrm.ColPrimaryGRN.Index).Value = PrimaryGRN
                    Next

                    For Each row As CustomSMESS.DseTransaksiAOP_Quick2_AddByCustomProduct.tbltrans_custom_aop_item1Row In vTblCustom.Rows
                        If row.Refered_To_PO = 0 Then
                            row.Primary_PO = PrimaryPO
                        End If
                        row.Primary_GRN = PrimaryGRN
                    Next
                    ObjItem.Simpan_Multi(vdgvItem, vDGVItemExchangeRate, vdgvItemDP, Me.mprimary, vDGVItemSR, vDGVItemCostData, vFrm, vTblCustom)

                    If vFromCutOff = False Then
                        Me.CekNomor(DTA)
                    End If
                    'simpan jurnal otomatis
                    For Each rowdgv In vDGVItemCostData.Rows
                        Dim PrimaryCurrCost As Double = -1
                        PrimaryCurrCost = CDbl(rowdgv.Cells(vFrm.ColPrimaryCurrencyCost.Index).Value)
                        TotalCOst += ObjCD.DigitGrouping(CDec(rowdgv.Cells(vFrm.ColCostAmount.Index).Value), DefprecBase)
                    Next

                    '**************************************************************************************************************************************
                    '**************************************************************************************************************************************
                    'PERUBAHAN SISTEM PERHITUNGAN JURNAL ERROR NO 50 TANGGAL 11 JULI 2016 -- PERBAIKAN 28 JULI 2016
                    'SEBELUMNYA NILAI SALES RUPIAH = NILAI SALES USD * EXCHANGE RATE
                    'SEKARANG NILAINYA DICEK DULU APAKAH ADA PENJUMLAHAN TERHADAP NILAI SALES (CONTOH TAX USD, SHIPMENT DLL),
                    'HASILNYA DITOTAL LALU DIKURANGI SEMUA PENJUMLAHAN, SISANYA DIANGGAP SEBAGAI SALES
                    'HAL INI DILAKUKAN UTK MENCEGAH SELISIH ANTARA SALES DI INVOICE JURNAL DENGAN SALES DI RECEIVEMENT JOURNAL

                    Dim DPRupiah As Decimal = 0
                    Dim PurchaseRupiah As Decimal = 0
                    Dim TotalPurchase As Decimal = 0
                    Dim AmountTanpaPurchase As Decimal = 0
                    Dim DefprecOri As Integer = 0
                    Dim PurchasePlusDP As Decimal = 0


                    DefprecBase = ObjCur.GetDecimal_DefaultCurrency
                    DefprecOri = ObjCur.GetDefaultPrecission(Me.mPrimaryCurrency)
                    If IsNumeric(Me.mTotal) Then
                        PurchasePlusDP += ObjCD.DigitGrouping(CDec(Me.mTotal), DefprecOri)
                    End If
                    'If IsNumeric(DPOriginal) Then
                    '    PurchasePlusDP += ObjCD.DigitGrouping(CDec(DPOriginal), DefprecOri)
                    'End If

                    PurchasePlusDP = ObjCD.DigitGrouping(PurchasePlusDP * Me.mTaxExchangeRate, DefprecBase)

                    If Me.mPrimaryCurrTax = Me.mPrimaryCurrency Then
                        AmountTanpaPurchase += ObjCD.DigitGrouping(Me.mTaxAmount * Me.mTaxExchangeRate, DefprecBase)
                        AmountTanpaPurchase += ObjCD.DigitGrouping(Me.mDutyStamp * Me.mTaxExchangeRate, DefprecBase)
                    End If
                    'AmountTanpaSalesRp = Me.ObjCD.DigitGrouping((AmountTanpaSales * Me.mExchangeRate), DefPrecBase)
                    PurchaseRupiah = ObjCD.DigitGrouping(PurchasePlusDP - AmountTanpaPurchase + TotalCOst, DefprecBase)
                    '**************************************************************************************************************************************

                    Me.TransferJurnal1(Me.mprimary, Me.mPrimarySupplier, _
                    PurchaseRupiah, ObjCD.DigitGrouping(Me.mTaxAmount * Me.mTaxExchangeRate, DefprecBase), _
                    ObjCD.DigitGrouping(Me.mDutyStamp * Me.mTaxExchangeRate, DefprecBase), Me.mNoUrut, Me.mTanggalTransaksi, _
                    Me.mnotransaksi, Me.mModifiedBy, CDbl(Me.mPaymentTerms), TotalCOst, vDGVItemCostData, vFrm, Me.mSupplierInvoiceNo, Me.mSubtotalBeforeDP, _
                    Me.mTaxExchangeRate, Me.mPrimaryCurrency, Me.mPrimaryAccount, Me.mPrimaryCurrTax, Me.mTaxAmount, Me.mDutyStamp)

                    'tandai po, grn dan stock in purchase yang dirujuk sebagai transaksi quick
                    ObjPO.SetPrimaryQuickAOP(Me.mprimary, PrimaryPO)
                    ObjStockInOut.SetPrimaryTransaksiQuick(Me.mprimary, PrimaryStockIn)
                    ObjSC.UpdateDocNo(Inventory_Module.ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockIn_Purchase, PrimaryStockIn, Me.mnotransaksi)

                    ''################################CUSTOM AOP#############################################
                    ObjCustAOP.Simpan_Custom(Me.mprimary, vTblCustom, tblProdQty, ListProductCustom)
                    ''################################CUSTOM AOP#############################################
                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mnotransaksi + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                Me.mnotransaksi = Me.GetNoBaru(Me.mTanggalTransaksi, Me.mPrimarySupplier)
                Return Me.SimpanBaru1(vdgvItem, vdgvItemDP, vDGVItemSR, vDGVItemExchangeRate, vUserCode, vDGVItemCostData, vFrm, vTblCustom, vTblItemCustomInOut, True, NoPO, NoStockIn, NoGRN, vKoreksiStock, vSetPOB, vFromCutOff)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
            ObjPO = Nothing
            ObjStockInOut = Nothing
            ObjSupplier = Nothing
            ObjUser = Nothing
            ObjGRN = Nothing
            ObjCur = Nothing
            ObjPOItem = Nothing
            ListPrimaryPO = Nothing
            ObjSC = Nothing
        End Try
    End Function


    Public Overridable Function SimpanPerubahan(ByVal vUserCode As String) As Boolean
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aopTableAdapter
        Dim DTA3 As New DseTransaksiAOPTableAdapters.tbltrans_aop2TableAdapter
        Dim ObjItem As New ClsTransaksiAOPItem
        Dim TaxExclude As Integer
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim vVATDipungut As Integer = 1
        Dim ObjPOItem As New ClsTransaksiPOItem
        Dim vPrimaryPODP As Double = -1
        Dim ObjPO As New ClsTransaksiPO
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            If Me.ApakahDataBenar(False, "", -1) = True Then
                'pastikan po sudah di otorisasi dengan baik
                If ObjPO.ApakahSudahOtorisasi(Me.mPrimaryPO4Dp) = False Then
                    MsgBox("Please authorize po first, you can't create purchase downpayment for unauthorized po.", MsgBoxStyle.Information, "Information")
                    Return False
                End If

                If Me.mTaxExclude = EnumTaxExclude.Exclude Then
                    TaxExclude = 1
                Else
                    TaxExclude = 0
                End If
                If Me.mVATDiPungut = True Then
                    vVATDipungut = 1
                Else
                    vVATDipungut = 0
                End If
                'hitung level otorisasi yang dibutuhkan
                vPrimaryPODP = Me.GetPrimaryPO4DP(Me.mprimary)
                If DTA3.UpdateQuery(Me.mnotransaksi, CInt(Me.mnotransaksi.Substring(0, 6)), _
                 Me.mTipeAOP, Me.mTanggalTransaksi, Me.mPrimarySupplier, Me.mSupplierInvoiceNo, Me.mLokal, _
                 Me.mPrimaryCurrency, Me.mCalculatesBy, Me.SubTotal, Me.mDownPayment, Me.mSubTotal, Me.mDutyStamp, TaxExclude, Me.Tax, _
                 Me.mTaxExchangeRate, Me.mTaxAmount, Me.mTotal, Me.mExchangeRateToUSD, Me.mExhangeRateToIDR, _
                 Me.mRemarks, Now, vUserCode, Now, Me.mSuppAddressIdentifier, Me.mPrimaryPO4Dp, Me.mPaymentTerms, _
                 vVATDipungut, Me.mNoFakturPajak, Me.mNoNPWPSupplier, Me.mEFak_KodeJenisTransaksiPajak, _
                 Me.mPrimaryCurrTax, Me.mPrimaryAccount, Me.mPrimaryAccountDP, Me.mSupplierInvoiceDate, Me.mDPPNilaiLain, Me.mprimary, Me.mTimestamp) > 0 Then
                    'update data yang berubah setelah simpan perubahan
                    Me.mNoUrut = Me.mnotransaksi.Substring(0, 6)
                    Me.mTimestamp = DTA.ScalarQuery_GetTimeStamp(Me.mprimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode


                    '**************************************************************************************************************************************
                    '**************************************************************************************************************************************
                    'PERUBAHAN SISTEM PERHITUNGAN JURNAL ERROR NO 50 TANGGAL 11 JULI 2016 -- PERBAIKAN 28 JULI 2016
                    'SEBELUMNYA NILAI SALES RUPIAH = NILAI SALES USD * EXCHANGE RATE
                    'SEKARANG NILAINYA DICEK DULU APAKAH ADA PENJUMLAHAN TERHADAP NILAI SALES (CONTOH TAX USD, SHIPMENT DLL),
                    'HASILNYA DITOTAL LALU DIKURANGI SEMUA PENJUMLAHAN, SISANYA DIANGGAP SEBAGAI SALES
                    'HAL INI DILAKUKAN UTK MENCEGAH SELISIH ANTARA SALES DI INVOICE JURNAL DENGAN SALES DI RECEIVEMENT JOURNAL

                    Dim PurchaseRupiah As Decimal = 0
                    Dim TotalPurchase As Decimal = 0
                    Dim AmountTanpaPurchase As Decimal = 0
                    Dim DefprecBase As Integer = 0
                    DefprecBase = ObjCur.GetDecimal_DefaultCurrency
                    TotalPurchase = ObjCD.DigitGrouping(Me.mTotal * Me.mTaxExchangeRate, DefprecBase)
                    If Me.mPrimaryCurrTax = Me.mPrimaryCurrency Then
                        AmountTanpaPurchase += ObjCD.DigitGrouping(Me.mTaxAmount * Me.mTaxExchangeRate, DefprecBase)
                        AmountTanpaPurchase += ObjCD.DigitGrouping(Me.mDutyStamp * Me.mTaxExchangeRate, DefprecBase)
                    End If
                    'AmountTanpaSalesRp = Me.ObjCD.DigitGrouping((AmountTanpaSales * Me.mExchangeRate), DefPrecBase)
                    PurchaseRupiah = ObjCD.DigitGrouping(TotalPurchase - AmountTanpaPurchase, DefprecBase)

                    '**************************************************************************************************************************************
                    'simpan jurnal otomatis
                    Me.TransferJurnal_DP1(Me.mprimary, Me.mPrimarySupplier, _
                    PurchaseRupiah, _
                    ObjCD.DigitGrouping(Me.mTaxAmount * Me.mTaxExchangeRate, DefprecBase), _
                    ObjCD.DigitGrouping(Me.mDutyStamp * Me.mTaxExchangeRate, DefprecBase), _
                    Me.mNoUrut, Me.mTanggalTransaksi, Me.mnotransaksi, Me.mModifiedBy, CDbl(Me.mPaymentTerms), _
                    Me.mSupplierInvoiceNo, Me.mSubtotalBeforeDP, Me.mTaxExchangeRate, Me.mPrimaryCurrency, _
                    Me.mPrimaryAccount, Me.mPrimaryCurrTax, Me.mTaxAmount, Me.mDutyStamp, Me.mPrimaryAccountDP)
                    'update data order balance, pastikan tidak ada po yang sudah di dp tapi kemudian bisa dicancel otorisasi
                    If Me.mPrimaryPO4Dp <> vPrimaryPODP Then
                        ObjPOItem.SetOrderBalance(vPrimaryPODP)
                    End If
                    ObjPOItem.SetOrderBalance(Me.mPrimaryPO4Dp)

                    'nilai balikan
                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mnotransaksi + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                Me.mnotransaksi = Me.GetNoBaru(Me.mTanggalTransaksi, Me.mPrimarySupplier)
                Return Me.SimpanPerubahan(vUserCode)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            DTA3.Dispose()
            ObjItem = Nothing
            ObjCur = Nothing
            ObjPO = Nothing
            ObjPOItem = Nothing
        End Try
    End Function


    Public Overridable Function SimpanPerubahan1(ByVal dgvItem As DataGridView, ByVal vdgvItemDP As DataGridView, _
        ByVal vDGVItemSR As DataGridView, ByVal vDGVItemExchangeRate As DataGridView, ByVal vUserCode As String, _
        ByVal vDGVItemCostData As DataGridView, ByVal vFrm As FrmTransaksiAOP2) As Boolean
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aopTableAdapter
        Dim DTA2 As New DseTransaksiAOPTableAdapters.tbltrans_aop2TableAdapter
        Dim ObjItem As New ClsTransaksiAOPItem
        Dim TaxExclude As Integer
        'Dim Primary4DP As Double = -1
        Dim rowdgv As DataGridViewRow
        Dim TotalCost As Decimal = 0
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim DPOriginal, DPBaseCurrency As Decimal
        Dim DTA1 As New DseTransaksiAOPTableAdapters.DataTable15TableAdapter
        Dim TBL1 As New DseTransaksiAOP.DataTable15DataTable
        Dim row1 As DseTransaksiAOP.DataTable15Row
        Dim DefPrec As Integer = 0
        Dim vVATDipungut As Integer = 1
        Dim objCD As New Core_Module.ClsCekData

        Try
            If Me.ApakahDataBenar1(False, dgvItem, vFrm) = True Then
                DefPrec = ObjCur.GetDecimal_DefaultCurrency
                If Me.mTaxExclude = EnumTaxExclude.Exclude Then
                    TaxExclude = 1
                Else
                    TaxExclude = 0
                End If
                If Me.mVATDiPungut = True Then
                    vVATDipungut = 1
                Else
                    vVATDipungut = 0
                End If
                'hitung level otorisasi yang dibutuhkan
                If DTA2.UpdateQuery(Me.mnotransaksi, CInt(Me.mnotransaksi.Substring(0, 6)), _
                  Me.mTipeAOP, Me.mTanggalTransaksi, Me.mPrimarySupplier, Me.mSupplierInvoiceNo, Me.mLokal, _
                 Me.mPrimaryCurrency, Me.mCalculatesBy, Me.SubtotalBeforeDP, Me.mDownPayment, Me.mSubTotal, Me.mDutyStamp, TaxExclude, Me.Tax, _
                 Me.mTaxExchangeRate, Me.mTaxAmount, Me.mTotal, Me.mExchangeRateToUSD, Me.mExhangeRateToIDR, _
                 Me.mRemarks, Now, vUserCode, Now, Me.mSuppAddressIdentifier, Me.mPrimaryPO4Dp, Me.mPaymentTerms, _
                 vVATDipungut, Me.mNoFakturPajak, Me.mNoNPWPSupplier, Me.mEFak_KodeJenisTransaksiPajak, _
                 Me.mPrimaryCurrTax, Me.mPrimaryAccount, Me.mPrimaryAccountDP, Me.mSupplierInvoiceDate, Me.mDPPNilaiLain, Me.mprimary, Me.mTimestamp) > 0 Then
                    'update data yang berubah setelah simpan perubahan
                    Me.mNoUrut = Me.mnotransaksi.Substring(0, 6)
                    Me.mTimestamp = DTA.ScalarQuery_GetTimeStamp(Me.mprimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode
                    'hapus data item lama
                    ObjItem.Hapus(Me.mprimary)
                    'simpan data item baru
                    ObjItem.Simpan_Multi(dgvItem, vDGVItemExchangeRate, vdgvItemDP, Me.mprimary, vDGVItemSR, vDGVItemCostData, vFrm)
                    'hitung DP jika ada
                    DTA1.Fill(TBL1, Me.mprimary)
                    DPOriginal = 0
                    DPBaseCurrency = 0
                    Dim PrimaryAkunDP As Double = -1
                    For Each row1 In TBL1.Rows
                        DPOriginal += objCD.DigitGrouping(row1.Sub_Total_Before_DP, ObjCur.GetDefaultPrecission(Me.mPrimaryCurrency))
                        DPBaseCurrency += objCD.DigitGrouping(row1.Sub_Total_Before_DP * row1.Tax_Exchange_Rate, DefPrec)
                        PrimaryAkunDP = row1.Primary_Akun_DP
                    Next

                    'simpan jurnal otomatis
                    For Each rowdgv In vDGVItemCostData.Rows
                        Dim PrimaryCurrCost As Double = -1
                        PrimaryCurrCost = CDbl(rowdgv.Cells(vFrm.ColPrimaryCurrencyCost.Index).Value)
                        TotalCost += objCD.DigitGrouping(CDec(rowdgv.Cells(vFrm.ColCostAmount.Index).Value), DefPrec)
                    Next



                    '**************************************************************************************************************************************
                    '**************************************************************************************************************************************
                    'PERUBAHAN SISTEM PERHITUNGAN JURNAL ERROR NO 50 TANGGAL 11 JULI 2016 -- PERBAIKAN 28 JULI 2016
                    'SEBELUMNYA NILAI SALES RUPIAH = NILAI SALES USD * EXCHANGE RATE
                    'SEKARANG NILAINYA DICEK DULU APAKAH ADA PENJUMLAHAN TERHADAP NILAI SALES (CONTOH TAX USD, SHIPMENT DLL),
                    'HASILNYA DITOTAL LALU DIKURANGI SEMUA PENJUMLAHAN, SISANYA DIANGGAP SEBAGAI SALES
                    'HAL INI DILAKUKAN UTK MENCEGAH SELISIH ANTARA SALES DI INVOICE JURNAL DENGAN SALES DI RECEIVEMENT JOURNAL

                    Dim DPRupiah As Decimal = 0
                    Dim PurchaseRupiah As Decimal = 0
                    Dim TotalPurchase As Decimal = 0
                    Dim AmountTanpaPurchase As Decimal = 0
                    Dim DefprecBase As Integer = 0
                    Dim DefprecOri As Integer = 0
                    Dim PurchasePlusDP As Decimal = 0


                    DefprecBase = ObjCur.GetDecimal_DefaultCurrency
                    DefprecOri = ObjCur.GetDefaultPrecission(Me.mPrimaryCurrency)
                    If IsNumeric(Me.mTotal) Then
                        PurchasePlusDP += objCD.DigitGrouping(CDec(Me.mTotal), DefprecOri)
                    End If
                    If IsNumeric(DPOriginal) Then
                        PurchasePlusDP += objCD.DigitGrouping(CDec(DPOriginal), DefprecOri)
                    End If

                    PurchasePlusDP = objCD.DigitGrouping(PurchasePlusDP * Me.mTaxExchangeRate, DefprecBase)

                    If Me.mPrimaryCurrTax = Me.mPrimaryCurrency Then
                        AmountTanpaPurchase += objCD.DigitGrouping(Me.mTaxAmount * Me.mTaxExchangeRate, DefprecBase)
                        AmountTanpaPurchase += objCD.DigitGrouping(Me.mDutyStamp * Me.mTaxExchangeRate, DefprecBase)
                    End If
                    'AmountTanpaSalesRp = Me.ObjCD.DigitGrouping((AmountTanpaSales * Me.mExchangeRate), DefPrecBase)
                    PurchaseRupiah = objCD.DigitGrouping(PurchasePlusDP - AmountTanpaPurchase + TotalCost, DefprecBase)
                    '**************************************************************************************************************************************


                    Me.TransferJurnal1(Me.mprimary, Me.mPrimarySupplier, PurchaseRupiah, _
                   objCD.DigitGrouping(Me.mTaxAmount * Me.mTaxExchangeRate, DefprecBase), _
                   objCD.DigitGrouping(Me.mDutyStamp * Me.mTaxExchangeRate, DefprecBase), _
                   Me.mNoUrut, Me.mTanggalTransaksi, Me.mnotransaksi, Me.mModifiedBy, CDbl(Me.mPaymentTerms), TotalCost, vDGVItemCostData, _
                   vFrm, Me.mSupplierInvoiceNo, Me.mSubtotalBeforeDP, Me.mTaxExchangeRate, _
                   Me.mPrimaryCurrency, DPOriginal, DPBaseCurrency, Me.mPrimaryAccount, Me.mPrimaryCurrTax, Me.mTaxAmount, Me.mDutyStamp, PrimaryAkunDP)


                    'nilai balikan
                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mnotransaksi + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                Me.mnotransaksi = Me.GetNoBaru(Me.mTanggalTransaksi, Me.mPrimarySupplier)
                Return Me.SimpanPerubahan1(dgvItem, vdgvItemDP, vDGVItemSR, vDGVItemExchangeRate, vUserCode, vDGVItemCostData, vFrm)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
            ObjCur = Nothing
            DTA1.Dispose()
            TBL1.Dispose()
            ObjCur = Nothing
        End Try
    End Function

    Private Sub TransferJurnal_DP1(ByVal vPrimaryAOP As Double, ByVal vPrimarySupplier As Double, ByVal vPurchase As Decimal, ByVal vVAT As Decimal, _
    ByVal vDutystamp As Decimal, ByVal vUrutanAOP As Double, ByVal vTanggalAOP As Date, ByVal vAOPNumber As String, _
    ByVal vUserCode As String, ByVal vPaymentTerm As Double, _
    ByVal vSupplierInvoiceNumber As String, ByVal vPurchaseOriginalAmount As Decimal, _
    ByVal vExchangeRate As Decimal, ByVal vPrimaryCurrency As Double, ByVal vPrimaryAccount As Double, ByVal vPrimaryCurrTax As Double, _
    ByVal vVatOriginal As Decimal, ByVal vDutyStampOriginal As Decimal, ByVal vPrimaryDPAccount As Double)
        Dim DTA_Balance As New DseTransaksiAOP_TransfJurnalTableAdapters.tbltrans_balance_purchaseTableAdapter
        Dim TipeBalance As EnumTipePurchaseBalance = EnumTipePurchaseBalance.Automatic
        Dim DTA_Jurnal As New DseTransaksiAOP_TransfJurnalTableAdapters.tbltrans_jurnalTableAdapter
        Dim DTA_JurnalItem As New DseTransaksiAOP_TransfJurnalTableAdapters.tbltrans_jurnal_itemTableAdapter
        Dim TipeJurnal As EnumTipeJournal = EnumTipeJournal.PurchaseDownPaymentJournal
        Dim RefNumber As String = CStr(vUrutanAOP)
        Dim PrimaryJurnal As Double
        Dim Debet, Credit As Decimal
        Dim ObjSupplier As New Core_Module.clsMasterSupplier
        Dim ObjKartu As New ClsTransaksi_Kartu_APAR
        Dim SuppCode, SuppName As String
        Dim Local As Boolean = False
        Dim LocalStr As String = ""
        Dim MonthStr As String = ""
        Dim PrimaryAkunAP As Double = -1
        Dim PrimaryAccountDP As Double = -1
        Dim PrimaryAkunAP_BaseCurrency As Double = -1
        Dim PrimaryAkunAP_BaseIDR As Double = -1
        Dim PrimaryAkun As Double = -1

        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim DefPrec, DefPrecOriginal As Integer
        Dim ObjMulti As New AccModule.ClsMultiAccount
        Try
            DefPrec = ObjCur.GetDecimal_DefaultCurrency
            DefPrecOriginal = ObjCur.GetDefaultPrecission(vPrimaryCurrency)

            LocalStr = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_AOP_DP)
            If LocalStr = "" Then
                MsgBox("No AOP DP Account found. Please contact your vendor if you see this message", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            'hapuskan data jurnal lama
            DTA_Jurnal.Delete(TipeJurnal, vPrimaryAOP)
            'hapuskan data balance sales lama
            DTA_Balance.Delete(vPrimaryAOP, TipeBalance)
            Local = ObjSupplier.ApakahLokal(vPrimarySupplier)
            If Local = True Then
                LocalStr = "L"
            Else
                LocalStr = "I"
            End If
            'hapus data kartu AP dan AR
            ObjKartu.Hapus(vPrimaryAOP, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis)

            If Me.ApakahTransferOtomatis(vTanggalAOP) = True Then
                SuppCode = ObjSupplier.GetSupplierCode(vPrimarySupplier)
                SuppName = ObjSupplier.GetSupplierName(vPrimarySupplier)
                'buat jurnal baru
                If RefNumber.Length < 6 Then
                    While RefNumber.Length < 6
                        RefNumber = "0" + RefNumber
                    End While
                End If
                MonthStr = vTanggalAOP.Month.ToString
                While MonthStr.Length < 2
                    MonthStr = "0" + MonthStr
                End While
                'data header
                DTA_Jurnal.Insert("PJDP-" + RefNumber + "/" + LocalStr + "/" + MonthStr + "/" + vTanggalAOP.Year.ToString, vTanggalAOP, _
                vUrutanAOP, TipeJurnal, vPrimaryAOP, vPrimarySupplier, "AOP: " + vAOPNumber _
                + " - " + SuppCode + " - " + SuppName, vUserCode, _
                Now, vUserCode, Now, Now)

                'data item
                PrimaryJurnal = DTA_Jurnal.ScalarQuery_GetPrimaryJurnal(TipeJurnal, vPrimaryAOP).Value
                'PrimaryAccount = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_AOP_DP)
                ' PrimaryAccountDP = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_AOP_DP)
                PrimaryAccountDP = vPrimaryDPAccount
                'PrimaryAccountDP = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_AOP_DP)

                PrimaryAkunAP = vPrimaryAccount
                PrimaryAkunAP_BaseCurrency = ObjMulti.GetPrimaryAkunAPBase(vPrimaryCurrency)
                PrimaryAkunAP_BaseIDR = ObjCur.GetPrimaryAkunAP(Core_Module.PrimaryBaseCurrency)

                '--------------------------------------------------------------------------------------------
                'biaya dibayar dimuka - dari awal selalu pilih yang rupiah
                '--------------------------------------------------------------------------------------------
                If vPurchase <> 0 Then
                    PrimaryAkun = PrimaryAccountDP
                    If PrimaryAkun <> -1 Then
                        Debet = vPurchase
                        Credit = 0
                        DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkun, Debet, Credit, "", Debet, Credit, 1)
                    End If
                End If
                '--------------------------------------------------------------------------------------------
                'simpan data vat in-debet
                '--------------------------------------------------------------------------------------------
                If vVAT <> 0 Then
                    PrimaryAkun = Me.GetPrimaryAkunVATIn
                    If PrimaryAkun <> -1 Then
                        If vPrimaryCurrTax = vPrimaryCurrency Then
                            DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkun, vVAT, 0, "", vVAT, 0, 1)
                        Else
                            DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkun, vVatOriginal, 0, "", vVatOriginal, 0, 1)
                        End If
                    End If
                End If
                '--------------------------------------------------------------------------------------------
                'simpan data duty stamp-debet
                '--------------------------------------------------------------------------------------------
                If vDutystamp <> 0 Then
                    PrimaryAkun = Me.GetPrimaryDutyStamp
                    If PrimaryAkun <> -1 Then
                        If vPrimaryCurrTax = vPrimaryCurrency Then
                            DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkun, vDutystamp, 0, "", vDutystamp, 0, 1)
                        Else
                            DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkun, vDutyStampOriginal, 0, "", vDutyStampOriginal, 0, 1)
                        End If
                    End If
                End If
                '--------------------------------------------------------------------------------------------
                'simpan data akun AP
                Debet = 0
                Credit = 0
                If vPrimaryCurrency = Core_Module.PrimaryBaseCurrency Then
                    '========================================================================================================================================
                    'JIKA TRANSAKSI IDR
                    '========================================================================================================================================
                    Credit = vPurchase + vVAT + vDutystamp
                    DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkunAP, Debet, Credit, "", Debet, vPurchaseOriginalAmount + vVatOriginal + vDutyStampOriginal, vExchangeRate)
                Else
                    '========================================================================================================================================
                    'JIKA TRANSAKSI NON IDR
                    '========================================================================================================================================
                    If vPrimaryCurrTax = vPrimaryCurrency Then
                        '----------------------------------------------------------------------------------------------------------------------------------
                        'JIKA Tax NON IDR ---------------------------------------------------------------------------------------------------------------
                        '----------------------------------------------------------------------------------------------------------------------------------
                        Credit = vPurchase + vVAT + vDutystamp
                        DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkunAP, Debet, Credit, "", Debet, vPurchaseOriginalAmount + vVatOriginal + vDutyStampOriginal, vExchangeRate)
                    Else
                        '----------------------------------------------------------------------------------------------------------------------------------
                        'JIKA Tax IDR ---------------------------------------------------------------------------------------------------------------
                        '----------------------------------------------------------------------------------------------------------------------------------
                        Credit = vPurchase
                        DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkunAP, Debet, Credit, "", Debet, vPurchaseOriginalAmount, vExchangeRate)
                        'masukkan data vat + dutystamp ke ap base
                        If vVatOriginal + vDutyStampOriginal <> 0 Then
                            Credit = vVatOriginal + vDutyStampOriginal
                            DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkunAP_BaseIDR, Debet, Credit, "", Debet, Credit, 1)
                        End If
                    End If
                End If

                'If vPrimaryCurrency = Core_Module.PrimaryBaseCurrency Then
                '    '========================================================================================================================================
                '    'JIKA TRANSAKSI IDR
                '    '========================================================================================================================================
                '    If PrimaryAkunAP = PrimaryAkunAP_BaseCurrency Then
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '        'JIKA APACCOUNT yg dipilih = base -------------------------------------------------------------------------------------------------
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '        Credit = vPurchase + vVAT + vDutystamp
                '        DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkunAP, Debet, Credit, "", Debet, vPurchaseOriginalAmount + vVatOriginal + vDutyStampOriginal, vExchangeRate)
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '    Else
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '        'JIKA APACCOUNT yg dipilih <> base -------------------------------------------------------------------------------------------------
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '        Credit = vPurchase
                '        DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkunAP, Debet, Credit, "", Debet, vPurchaseOriginalAmount, vExchangeRate)

                '        'masukkan data vat + dutystamp ke ap base
                '        Credit = vVAT + vDutystamp
                '        If vVAT + vDutystamp <> 0 Then
                '            DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkunAP_BaseCurrency, Debet, Credit, "", Debet, vVatOriginal + vDutyStampOriginal, vExchangeRate)
                '        End If
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '    End If
                'Else
                '    '========================================================================================================================================
                '    'JIKA TRANSAKSI NON IDR
                '    '========================================================================================================================================
                '    If PrimaryAkunAP = PrimaryAkunAP_BaseCurrency Then
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '        'JIKA APACCOUNT yg dipilih = base -------------------------------------------------------------------------------------------------
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '        If vPrimaryCurrTax = vPrimaryCurrency Then
                '            '----------------------------------------------------------------------------------------------------------------------------------
                '            'JIKA Tax NON IDR ---------------------------------------------------------------------------------------------------------------
                '            '----------------------------------------------------------------------------------------------------------------------------------
                '            Credit = vPurchase + vVAT + vDutystamp
                '            DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkunAP, Debet, Credit, "", Debet, vPurchaseOriginalAmount + vVatOriginal + vDutyStampOriginal, vExchangeRate)
                '        Else
                '            '----------------------------------------------------------------------------------------------------------------------------------
                '            'JIKA Tax IDR ------------------------------------------------------------------------------------------------------------------
                '            '----------------------------------------------------------------------------------------------------------------------------------
                '            Credit = vPurchase
                '            DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkunAP, Debet, Credit, "", Debet, vPurchaseOriginalAmount, vExchangeRate)
                '            'masukkan data vat + dutystamp ke ap base
                '            If vVatOriginal + vDutyStampOriginal <> 0 Then
                '                Credit = vVatOriginal + vDutyStampOriginal
                '                DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkunAP_BaseIDR, Debet, Credit, "", Debet, Credit, 1)
                '            End If
                '        End If
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '    Else
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '        'JIKA APACCOUNT yg dipilih <> base -------------------------------------------------------------------------------------------------
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '        If vPrimaryCurrTax = vPrimaryCurrency Then
                '            '----------------------------------------------------------------------------------------------------------------------------------
                '            'JIKA Tax NON IDR ---------------------------------------------------------------------------------------------------------------
                '            '----------------------------------------------------------------------------------------------------------------------------------
                '            '----------------------------------------------------------------------------------------------------------------------------------
                '            Credit = vPurchase
                '            DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkunAP, Debet, Credit, "", Debet, vPurchaseOriginalAmount, vExchangeRate)

                '            If vVAT + vDutystamp <> 0 Then
                '                Credit = vVAT + vDutystamp
                '                DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkunAP_BaseCurrency, Debet, Credit, "", Debet, vVatOriginal + vDutyStampOriginal, vExchangeRate)
                '            End If
                '        Else
                '            '----------------------------------------------------------------------------------------------------------------------------------
                '            'JIKA Tax IDR ------------------------------------------------------------------------------------------------------------------
                '            '----------------------------------------------------------------------------------------------------------------------------------
                '            Credit = vPurchase
                '            DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkunAP, Debet, Credit, "", Debet, vPurchaseOriginalAmount, vExchangeRate)

                '            If vVatOriginal + vDutyStampOriginal <> 0 Then
                '                Credit = vVatOriginal + vDutyStampOriginal
                '                DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkunAP_BaseIDR, Debet, Credit, "", Debet, Credit, 1)
                '            End If
                '        End If
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '    End If
                'End If


                If vPrimaryCurrency = Core_Module.PrimaryBaseCurrency Then
                    '========================================================================================================================================
                    'JIKA TRANSAKSI IDR
                    '========================================================================================================================================
                    'aop idr
                    DTA_Balance.InsertQuery(vPrimarySupplier, vPrimaryAOP, TipeBalance, "", #1/1/1950#, "", vTanggalAOP + TimeSpan.FromDays(vPaymentTerm), _
                    vPurchase, vVAT, vDutystamp, vPurchase + vVAT + vDutystamp, 0, vUserCode, Now, vUserCode, Now, Now, Core_Module.PrimaryBaseCurrency, _
                    vPurchase + vVAT + vDutystamp, 1, PrimaryAkunAP)

                    'buat kartu ap baru
                    ObjKartu.Simpan(vPrimaryAOP, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, vAOPNumber, _
                    vPrimarySupplier, 0, vPurchase + vVAT + vDutystamp, "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
                    vTanggalAOP, vPrimaryAOP, vTanggalAOP, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, _
                    Core_Module.PrimaryBaseCurrency, 0, vPurchase + vVAT + vDutystamp, PrimaryAkunAP_BaseCurrency, 1)
                Else
                    '========================================================================================================================================
                    'JIKA TRANSAKSI NON IDR
                    '========================================================================================================================================
                    If vPrimaryCurrTax = vPrimaryCurrency Then
                        '----------------------------------------------------------------------------------------------------------------------------------
                        'JIKA Tax NON IDR ---------------------------------------------------------------------------------------------------------------
                        '----------------------------------------------------------------------------------------------------------------------------------
                        'balance purchase amount
                        DTA_Balance.InsertQuery(vPrimarySupplier, vPrimaryAOP, TipeBalance, "", #1/1/1950#, "", vTanggalAOP + TimeSpan.FromDays(vPaymentTerm), _
                        vPurchase, vVAT, vDutystamp, _
                        vPurchase + vVAT + vDutystamp, _
                        0, vUserCode, Now, vUserCode, Now, Now, vPrimaryCurrency, _
                         vPurchaseOriginalAmount + vVatOriginal + vDutyStampOriginal, vExchangeRate, _
                        PrimaryAkunAP)

                        'buat kartu ar purchase amount
                        ObjKartu.Simpan(vPrimaryAOP, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, vAOPNumber, _
                        vPrimarySupplier, 0, _
                        vPurchase + vVAT + vDutystamp, _
                        "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
                        vTanggalAOP, vPrimaryAOP, vTanggalAOP, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, _
                        vPrimaryCurrency, 0, _
                        vPurchaseOriginalAmount + vVatOriginal + vDutyStampOriginal, _
                        PrimaryAkunAP, vExchangeRate)
                    Else
                        '----------------------------------------------------------------------------------------------------------------------------------
                        'JIKA Tax IDR ---------------------------------------------------------------------------------------------------------------
                        '----------------------------------------------------------------------------------------------------------------------------------
                        'balance purchase amount
                        DTA_Balance.InsertQuery(vPrimarySupplier, vPrimaryAOP, TipeBalance, "", #1/1/1950#, "", vTanggalAOP + TimeSpan.FromDays(vPaymentTerm), _
                        vPurchase, 0, 0, _
                        vPurchase, _
                        0, vUserCode, Now, vUserCode, Now, Now, vPrimaryCurrency, _
                         vPurchaseOriginalAmount, vExchangeRate, _
                        PrimaryAkunAP)

                        'buat kartu ar purchase amount
                        ObjKartu.Simpan(vPrimaryAOP, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, vAOPNumber, _
                        vPrimarySupplier, 0, _
                        vPurchase, _
                        "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
                        vTanggalAOP, vPrimaryAOP, vTanggalAOP, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, _
                        vPrimaryCurrency, 0, _
                        vPurchaseOriginalAmount, _
                        PrimaryAkunAP, vExchangeRate)

                        If vVAT + vDutystamp <> 0 Then
                            'balance purchase tax
                            DTA_Balance.InsertQuery(vPrimarySupplier, vPrimaryAOP, TipeBalance, "", _
                            #1/1/1950#, "", vTanggalAOP + TimeSpan.FromDays(vPaymentTerm), _
                            0, vVAT, vDutystamp, vVAT + vDutystamp, 0, vUserCode, Now, vUserCode, Now, Now, vPrimaryCurrTax, _
                             vVatOriginal + vDutyStampOriginal, 1, PrimaryAkunAP_BaseIDR)

                            'buat kartu ar purchase tax
                            ObjKartu.Simpan(vPrimaryAOP, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, vAOPNumber, _
                            vPrimarySupplier, 0, vVAT + vDutystamp, _
                            "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
                            vTanggalAOP, vPrimaryAOP, vTanggalAOP, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, _
                            vPrimaryCurrTax, 0, _
                           vVatOriginal + vDutyStampOriginal, PrimaryAkunAP_BaseIDR, 1)
                        End If
                    End If
                End If



                'If vPrimaryCurrency = Core_Module.PrimaryBaseCurrency Then
                '    '========================================================================================================================================
                '    'JIKA TRANSAKSI IDR
                '    '========================================================================================================================================
                '    If PrimaryAkunAP = PrimaryAkunAP_BaseCurrency Then
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '        'JIKA APACCOUNT yg dipilih = base -------------------------------------------------------------------------------------------------
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '        'aop idr
                '        DTA_Balance.InsertQuery(vPrimarySupplier, vPrimaryAOP, TipeBalance, "", #1/1/1950#, "", vTanggalAOP + TimeSpan.FromDays(vPaymentTerm), _
                '        vPurchase, vVAT, vDutystamp, vPurchase + vVAT + vDutystamp, 0, vUserCode, Now, vUserCode, Now, Now, Core_Module.PrimaryBaseCurrency, _
                '        vPurchase + vVAT + vDutystamp, 1, PrimaryAkunAP)

                '        'buat kartu ap baru
                '        ObjKartu.Simpan(vPrimaryAOP, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, vAOPNumber, _
                '        vPrimarySupplier, 0, vPurchase + vVAT + vDutystamp, "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
                '        vTanggalAOP, vPrimaryAOP, vTanggalAOP, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, _
                '        Core_Module.PrimaryBaseCurrency, 0, vPurchase + vVAT + vDutystamp, PrimaryAkunAP_BaseCurrency, 1)
                '    Else
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '        'JIKA APACCOUNT yg dipilih <> base -------------------------------------------------------------------------------------------------
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '        'aop idr
                '        DTA_Balance.InsertQuery(vPrimarySupplier, vPrimaryAOP, TipeBalance, "", #1/1/1950#, "", vTanggalAOP + TimeSpan.FromDays(vPaymentTerm), _
                '        vPurchase, 0, 0, vPurchase, 0, vUserCode, Now, vUserCode, Now, Now, Core_Module.PrimaryBaseCurrency, _
                '        vPurchase, vExchangeRate, PrimaryAkunAP)

                '        'buat kartu ap baru
                '        ObjKartu.Simpan(vPrimaryAOP, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, vAOPNumber, _
                '        vPrimarySupplier, 0, vPurchase, "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
                '        vTanggalAOP, vPrimaryAOP, vTanggalAOP, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, _
                '        Core_Module.PrimaryBaseCurrency, 0, vPurchase, PrimaryAkunAP_BaseCurrency, 1)

                '        If vVAT + vDutystamp <> 0 Then
                '            'TAX
                '            DTA_Balance.InsertQuery(vPrimarySupplier, vPrimaryAOP, TipeBalance, "", _
                '            #1/1/1950#, "", vTanggalAOP + TimeSpan.FromDays(vPaymentTerm), _
                '            0, vVAT, vDutystamp, vVAT + vDutystamp, 0, vUserCode, Now, vUserCode, Now, Now, Core_Module.PrimaryBaseCurrency, _
                '            vVatOriginal + vDutyStampOriginal, vExchangeRate, PrimaryAkunAP_BaseCurrency)

                '            'buat kartu ap baru
                '            ObjKartu.Simpan(vPrimaryAOP, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, vAOPNumber, _
                '            vPrimarySupplier, 0, vVAT + vDutystamp, "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
                '            vTanggalAOP, vPrimaryAOP, vTanggalAOP, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, _
                '            Core_Module.PrimaryBaseCurrency, 0, vVatOriginal + vDutyStampOriginal, PrimaryAkunAP_BaseCurrency, 1)
                '        End If
                '    End If

                'Else
                '    '========================================================================================================================================
                '    'JIKA TRANSAKSI NON IDR
                '    '========================================================================================================================================
                '    If PrimaryAkunAP = PrimaryAkunAP_BaseCurrency Then
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '        'JIKA APACCOUNT yg dipilih = base -------------------------------------------------------------------------------------------------
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '        If vPrimaryCurrTax = vPrimaryCurrency Then
                '            '----------------------------------------------------------------------------------------------------------------------------------
                '            'JIKA Tax NON IDR ---------------------------------------------------------------------------------------------------------------
                '            '----------------------------------------------------------------------------------------------------------------------------------
                '            'balance purchase amount
                '            DTA_Balance.InsertQuery(vPrimarySupplier, vPrimaryAOP, TipeBalance, "", #1/1/1950#, "", vTanggalAOP + TimeSpan.FromDays(vPaymentTerm), _
                '            vPurchase, vVAT, vDutystamp, _
                '            vPurchase + vVAT + vDutystamp, _
                '            0, vUserCode, Now, vUserCode, Now, Now, vPrimaryCurrency, _
                '             vPurchaseOriginalAmount + vVatOriginal + vDutyStampOriginal, vExchangeRate, _
                '            PrimaryAkunAP)

                '            'buat kartu ar purchase amount
                '            ObjKartu.Simpan(vPrimaryAOP, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, vAOPNumber, _
                '            vPrimarySupplier, 0, _
                '            vPurchase + vVAT + vDutystamp, _
                '            "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
                '            vTanggalAOP, vPrimaryAOP, vTanggalAOP, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, _
                '            vPrimaryCurrency, 0, _
                '            vPurchaseOriginalAmount + vVatOriginal + vDutyStampOriginal, _
                '            PrimaryAkunAP, vExchangeRate)
                '        Else
                '            '----------------------------------------------------------------------------------------------------------------------------------
                '            'JIKA Tax IDR -------------------------------------------------------------------------------------------------------------------
                '            '----------------------------------------------------------------------------------------------------------------------------------
                '            'balance purchase amount
                '            DTA_Balance.InsertQuery(vPrimarySupplier, vPrimaryAOP, TipeBalance, "", #1/1/1950#, "", vTanggalAOP + TimeSpan.FromDays(vPaymentTerm), _
                '            vPurchase, 0, 0, _
                '            vPurchase, _
                '            0, vUserCode, Now, vUserCode, Now, Now, vPrimaryCurrency, _
                '             vPurchaseOriginalAmount, vExchangeRate, _
                '            PrimaryAkunAP)

                '            'buat kartu ar purchase amount
                '            ObjKartu.Simpan(vPrimaryAOP, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, vAOPNumber, _
                '            vPrimarySupplier, 0, _
                '            vPurchase, _
                '            "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
                '            vTanggalAOP, vPrimaryAOP, vTanggalAOP, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, _
                '            vPrimaryCurrency, 0, _
                '            vPurchaseOriginalAmount, _
                '            PrimaryAkunAP, vExchangeRate)

                '            If vVAT + vDutystamp <> 0 Then
                '                'balance purchase tax
                '                DTA_Balance.InsertQuery(vPrimarySupplier, vPrimaryAOP, TipeBalance, "", _
                '                #1/1/1950#, "", vTanggalAOP + TimeSpan.FromDays(vPaymentTerm), _
                '                0, vVAT, vDutystamp, vVAT + vDutystamp, 0, vUserCode, Now, vUserCode, Now, Now, vPrimaryCurrTax, _
                '                 vVatOriginal + vDutyStampOriginal, 1, PrimaryAkunAP_BaseIDR)

                '                'buat kartu ar purchase tax
                '                ObjKartu.Simpan(vPrimaryAOP, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, vAOPNumber, _
                '                vPrimarySupplier, 0, vVAT + vDutystamp, _
                '                "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
                '                vTanggalAOP, vPrimaryAOP, vTanggalAOP, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, _
                '                vPrimaryCurrTax, 0, _
                '               vVatOriginal + vDutyStampOriginal, PrimaryAkunAP_BaseIDR, 1)
                '            End If
                '        End If
                '    Else
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '        'JIKA APACCOUNT yg dipilih <> base -------------------------------------------------------------------------------------------------
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '        If vPrimaryCurrTax = vPrimaryCurrency Then
                '            '----------------------------------------------------------------------------------------------------------------------------------
                '            'JIKA Tax Non IDR -------------------------------------------------------------------------------------------------------------------
                '            '----------------------------------------------------------------------------------------------------------------------------------
                '            'balance purchase amount
                '            DTA_Balance.InsertQuery(vPrimarySupplier, vPrimaryAOP, TipeBalance, "", #1/1/1950#, "", vTanggalAOP + TimeSpan.FromDays(vPaymentTerm), _
                '            vPurchase, 0, 0, _
                '            vPurchase, _
                '            0, vUserCode, Now, vUserCode, Now, Now, vPrimaryCurrency, _
                '             vPurchaseOriginalAmount, vExchangeRate, _
                '            PrimaryAkunAP)

                '            'buat kartu ar purchase amount
                '            ObjKartu.Simpan(vPrimaryAOP, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, vAOPNumber, _
                '            vPrimarySupplier, 0, _
                '            vPurchase, _
                '            "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
                '            vTanggalAOP, vPrimaryAOP, vTanggalAOP, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, _
                '            vPrimaryCurrency, 0, _
                '            vPurchaseOriginalAmount, _
                '            PrimaryAkunAP, vExchangeRate)

                '            If vVAT + vDutystamp <> 0 Then
                '                'balance purchase tax
                '                DTA_Balance.InsertQuery(vPrimarySupplier, vPrimaryAOP, TipeBalance, "", #1/1/1950#, "", vTanggalAOP + TimeSpan.FromDays(vPaymentTerm), _
                '                0, vVAT, vDutystamp, vVAT + vDutystamp, 0, vUserCode, Now, vUserCode, Now, Now, vPrimaryCurrTax, _
                '                 vVatOriginal + vDutyStampOriginal, vExchangeRate, PrimaryAkunAP_BaseCurrency)

                '                'buat kartu ar purchase tax
                '                ObjKartu.Simpan(vPrimaryAOP, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, vAOPNumber, _
                '                vPrimarySupplier, 0, vVAT + vDutystamp, _
                '                "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
                '                vTanggalAOP, vPrimaryAOP, vTanggalAOP, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, _
                '                vPrimaryCurrTax, 0, _
                '               vVatOriginal + vDutyStampOriginal, PrimaryAkunAP_BaseCurrency, vExchangeRate)
                '            End If

                '        Else
                '            '----------------------------------------------------------------------------------------------------------------------------------
                '            'JIKA Tax IDR -------------------------------------------------------------------------------------------------------------------
                '            '----------------------------------------------------------------------------------------------------------------------------------
                '            'balance purchase amount
                '            DTA_Balance.InsertQuery(vPrimarySupplier, vPrimaryAOP, TipeBalance, "", #1/1/1950#, "", vTanggalAOP + TimeSpan.FromDays(vPaymentTerm), _
                '            vPurchase, 0, 0, _
                '            vPurchase, _
                '            0, vUserCode, Now, vUserCode, Now, Now, vPrimaryCurrency, _
                '             vPurchaseOriginalAmount, vExchangeRate, _
                '            PrimaryAkunAP)

                '            'buat kartu ar purchase amount
                '            ObjKartu.Simpan(vPrimaryAOP, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, vAOPNumber, _
                '            vPrimarySupplier, 0, _
                '            vPurchase, _
                '            "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
                '            vTanggalAOP, vPrimaryAOP, vTanggalAOP, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, _
                '            vPrimaryCurrency, 0, _
                '            vPurchaseOriginalAmount, _
                '            PrimaryAkunAP, vExchangeRate)

                '            If vVatOriginal + vDutyStampOriginal <> 0 Then
                '                'balance purchase tax
                '                DTA_Balance.InsertQuery(vPrimarySupplier, vPrimaryAOP, TipeBalance, _
                '                "", #1/1/1950#, "", vTanggalAOP + TimeSpan.FromDays(vPaymentTerm), _
                '                0, vVatOriginal, vDutyStampOriginal, vVatOriginal + vDutyStampOriginal, _
                '                0, vUserCode, Now, vUserCode, Now, Now, vPrimaryCurrTax, _
                '                 vVatOriginal + vDutyStampOriginal, 1, PrimaryAkunAP_BaseCurrency)

                '                'buat kartu ar purchase tax
                '                ObjKartu.Simpan(vPrimaryAOP, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, vAOPNumber, _
                '                vPrimarySupplier, 0, vVatOriginal + vDutyStampOriginal, _
                '                "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
                '                vTanggalAOP, vPrimaryAOP, vTanggalAOP, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, _
                '                vPrimaryCurrTax, 0, _
                '               vVatOriginal + vDutyStampOriginal, PrimaryAkunAP_BaseCurrency, 1)
                '            End If

                '        End If
                '    End If
                'End If
            End If


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_Balance.Dispose()
            DTA_Jurnal.Dispose()
            DTA_JurnalItem.Dispose()
            ObjSupplier = Nothing
            ObjKartu = Nothing
            ObjSetting = Nothing
            ObjCur = Nothing
        End Try
    End Sub

    Private Sub TransferJurnal1(ByVal vPrimaryAOP As Double, ByVal vPrimarySupplier As Double, ByVal vPurchase As Decimal, ByVal vVAT As Decimal, _
ByVal vDutystamp As Decimal, ByVal vUrutanAOP As Double, ByVal vTanggalAOP As Date, ByVal vAOPNumber As String, _
ByVal vUserCode As String, ByVal vPaymentTerm As Double, ByVal vTotalCostIDR As Decimal, _
ByVal vDGVCost As DataGridView, ByVal vFrm As FrmTransaksiAOP2, _
ByVal vSupplierInvoiceNumber As String, ByVal vPurchaseOriginalAmount As Decimal, _
ByVal vExchangeRate As Decimal, ByVal vPrimaryCurrency As Double, _
ByVal vDP_Original As Decimal, ByVal vDPBaseCurrency As Decimal, ByVal vPrimaryAccountAP As Double, ByVal vPrimaryCurrTax As Double, _
ByVal vVATOriginal As Decimal, ByVal vDutyStampOriginal As Decimal, ByVal vPrimaryAccountDP As Double)

        Dim DTA_Balance As New DseTransaksiAOP_TransfJurnalTableAdapters.tbltrans_balance_purchaseTableAdapter
        Dim TipeBalance As EnumTipePurchaseBalance = EnumTipePurchaseBalance.Automatic
        Dim DTA_Jurnal As New DseTransaksiAOP_TransfJurnalTableAdapters.tbltrans_jurnalTableAdapter
        Dim DTA_JurnalItem As New DseTransaksiAOP_TransfJurnalTableAdapters.tbltrans_jurnal_itemTableAdapter
        Dim TipeJurnal As EnumTipeJournal = EnumTipeJournal.PurchaseJournal
        Dim RefNumber As String = CStr(vUrutanAOP)
        Dim PrimaryJurnal As Double
        Dim Debet, Credit, DebetOriginal, CreditOriginal As Decimal
        Dim rowDGV As DataGridViewRow
        Dim DTAAkun As New DseTransaksiAOP_TransfJurnalTableAdapters.tblmaster_akunTableAdapter
        Dim ObjSupplier As New Core_Module.clsMasterSupplier
        Dim Local As Boolean = False
        Dim LocalStr As String = ""
        Dim ObjKartu As New ClsTransaksi_Kartu_APAR
        Dim SuppCode, SuppName As String
        Dim MonthStr As String = ""
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim AP_PadaCost As Decimal = 0
        Dim AP_PadaCost_Original As Decimal = 0
        Dim AP_PadaCost_Base As Decimal = 0
        Dim PrimaryAkunAP As Double = -1
        Dim PrimaryAkunAP_BaseCurrency As Double = -1
        Dim PrimaryAkunAP_BaseIDR As Double = -1
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim PrimaryAkun As Double = -1
        Dim DTAInv As New DseTransaksiAOP_TransfJurnalTableAdapters.TBLInventoryTableAdapter
        Dim TBLInv As New DseTransaksiAOP_TransfJurnal.TBLInventoryDataTable
        Dim rowInv As DseTransaksiAOP_TransfJurnal.TBLInventoryRow
        Dim ListAkunInv As New List(Of Double)
        Dim ObjCD As New Core_Module.ClsCekData
        Dim DefPrecBase As Integer = 0
        Dim SelisihDP As Decimal = 0
        Dim Dummy1 As Decimal = 0
        Dim ListPrimaryAkunCost As New List(Of Double)
        Dim ObjMulti As New AccModule.ClsMultiAccount
        Dim objAkun As New AccModule.ClsMasterAkun

        Try
            DefPrecBase = ObjCur.GetDecimal_DefaultCurrency
            DTAInv.Connection.ConnectionString = Core_Module.GetConstring
            DTA_Jurnal.Connection.ConnectionString = DTAInv.Connection.ConnectionString
            DTA_Balance.Connection.ConnectionString = DTAInv.Connection.ConnectionString
            DTAAkun.Connection.ConnectionString = DTAInv.Connection.ConnectionString

            'hapuskan data jurnal lama
            DTA_Jurnal.Delete(TipeJurnal, vPrimaryAOP)
            'hapuskan data balance sales lama
            DTA_Balance.Delete(vPrimaryAOP, TipeBalance)
            Local = ObjSupplier.ApakahLokal(vPrimarySupplier)
            If Local = True Then
                LocalStr = "L"
            Else
                LocalStr = "I"
            End If
            'hapus data kartu AP dan AR
            ObjKartu.Hapus(vPrimaryAOP, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis)

            If Me.ApakahTransferOtomatis(vTanggalAOP) = True Then
                SuppCode = ObjSupplier.GetSupplierCode(vPrimarySupplier)
                SuppName = ObjSupplier.GetSupplierName(vPrimarySupplier)
                'buat jurnal baru
                If RefNumber.Length < 6 Then
                    While RefNumber.Length < 6
                        RefNumber = "0" + RefNumber
                    End While
                End If
                MonthStr = vTanggalAOP.Month.ToString
                While MonthStr.Length < 2
                    MonthStr = "0" + MonthStr
                End While
                'data header
                DTA_Jurnal.Insert("PJ-" + RefNumber + "/" + LocalStr + "/" + MonthStr + "/" + vTanggalAOP.Year.ToString, vTanggalAOP, _
                vUrutanAOP, TipeJurnal, vPrimaryAOP, vPrimarySupplier, "AOP: " + vAOPNumber _
                + " - " + SuppCode + " - " + SuppName, vUserCode, _
                Now, vUserCode, Now, Now)

                'data item
                PrimaryJurnal = DTA_Jurnal.ScalarQuery_GetPrimaryJurnal(TipeJurnal, vPrimaryAOP).Value

                'simpan data purchase-inventory-debet
                DTAInv.Fill(TBLInv, vPrimaryAOP)

                For Each rowInv In TBLInv.Rows
                    If ListAkunInv.Contains(rowInv.Primary_Inventory_Account) = False Then
                        ListAkunInv.Add(rowInv.Primary_Inventory_Account)
                        Debet = Me.GetInventoryPurchased(TBLInv, rowInv.Primary_Inventory_Account, vPurchase)
                        Credit = 0
                        DTA_JurnalItem.InsertQuery(PrimaryJurnal, rowInv.Primary_Inventory_Account, Debet, Credit, "", Debet, Credit, 1)
                    End If
                Next

                'simpan data vat in-debet
                If vVAT <> 0 Then
                    PrimaryAkun = Me.GetPrimaryAkunVATIn
                    If PrimaryAkun <> -1 Then
                        If vPrimaryCurrTax > 1 Then
                            DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkun, vVAT, 0, "", vVAT, 0, 1)
                        Else
                            DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkun, vVATOriginal, 0, "", vVATOriginal, 0, 1)
                        End If
                    End If
                End If

                'simpan data duty stamp-debet
                If vDutystamp <> 0 Then
                    PrimaryAkun = Me.GetPrimaryDutyStamp
                    If PrimaryAkun <> -1 Then
                        If vPrimaryCurrTax > 1 Then
                            DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkun, vDutystamp, 0, "", vDutystamp, 0, 1)
                        Else
                            DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkun, vDutyStampOriginal, 0, "", vDutyStampOriginal, 0, 1)
                        End If
                    End If
                End If

                '========================================================================================
                'Code Erna
                '========================================================================================
                'simpan data ap, kredit
                PrimaryAkunAP = vPrimaryAccountAP

                PrimaryAkunAP_BaseCurrency = ObjMulti.GetPrimaryAkunAPBase(vPrimaryCurrency)
                PrimaryAkunAP_BaseIDR = ObjCur.GetPrimaryAkunAP(Core_Module.PrimaryBaseCurrency)
                Debet = 0
                Credit = 0
                Dim costOriginal As Decimal = 0
                Dim cost As Decimal = 0

                If vPrimaryCurrency = Core_Module.PrimaryBaseCurrency Then
                    '========================================================================================================================================
                    'JIKA TRANSAKSI IDR
                    '========================================================================================================================================
                    Credit = (vPurchase - vTotalCostIDR) + vVAT + vDutystamp
                    'masukkan data cost yang ap + cost
                    cost = 0
                    costOriginal = 0
                    For Each rowDGV In vDGVCost.Rows
                        If CDbl(rowDGV.Cells(vFrm.ColPrimaryAccount.Index).Value) = PrimaryAkunAP Then
                            Dim PrimaryCurCost As Double = -1
                            PrimaryCurCost = rowDGV.Cells(vFrm.ColPrimaryCurrencyCost.Index).Value
                            costOriginal += ObjCD.DigitGrouping(rowDGV.Cells(vFrm.ColCostNominal.Index).Value, ObjCur.GetDefaultPrecission(PrimaryCurCost))
                            cost += ObjCD.DigitGrouping(rowDGV.Cells(vFrm.ColCostAmount.Index).Value, ObjCur.GetDefaultPrecission(PrimaryCurCost))
                        End If
                    Next
                    Credit += cost
                    DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkunAP, Debet, Credit, "", Debet, vPurchaseOriginalAmount + vVATOriginal + vDutyStampOriginal + costOriginal, vExchangeRate)

                Else
                    '========================================================================================================================================
                    'JIKA TRANSAKSI NON IDR
                    '========================================================================================================================================
                    If vPrimaryCurrTax = vPrimaryCurrency Then
                        '----------------------------------------------------------------------------------------------------------------------------------
                        'JIKA Tax NON IDR ---------------------------------------------------------------------------------------------------------------
                        '----------------------------------------------------------------------------------------------------------------------------------
                        Credit = (vPurchase - vTotalCostIDR) + vVAT + vDutystamp
                        'masukkan data cost yang ap + cost
                        cost = 0
                        costOriginal = 0
                        For Each rowDGV In vDGVCost.Rows
                            If CDbl(rowDGV.Cells(vFrm.ColPrimaryAccount.Index).Value) = PrimaryAkunAP Then
                                Dim PrimaryCurCost As Double = -1
                                PrimaryCurCost = rowDGV.Cells(vFrm.ColPrimaryCurrencyCost.Index).Value
                                costOriginal += ObjCD.DigitGrouping(rowDGV.Cells(vFrm.ColCostNominal.Index).Value, ObjCur.GetDefaultPrecission(PrimaryCurCost))
                                cost += ObjCD.DigitGrouping(rowDGV.Cells(vFrm.ColCostAmount.Index).Value, ObjCur.GetDefaultPrecission(PrimaryCurCost))
                            End If
                        Next
                        Credit += cost
                        DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkunAP, Debet, Credit, "", Debet, vPurchaseOriginalAmount + vVATOriginal + vDutyStampOriginal + costOriginal, vExchangeRate)
                    Else
                        '----------------------------------------------------------------------------------------------------------------------------------
                        'JIKA Tax IDR ---------------------------------------------------------------------------------------------------------------
                        '----------------------------------------------------------------------------------------------------------------------------------
                        Credit = (vPurchase - vTotalCostIDR)
                        'masukkan data cost yang ap + cost ke ap base
                        cost = 0
                        costOriginal = 0
                        For Each rowDGV In vDGVCost.Rows
                            If CDbl(rowDGV.Cells(vFrm.ColPrimaryAccount.Index).Value) = PrimaryAkunAP Then
                                Dim PrimaryCurCost As Double = -1
                                PrimaryCurCost = rowDGV.Cells(vFrm.ColPrimaryCurrencyCost.Index).Value
                                costOriginal += ObjCD.DigitGrouping(rowDGV.Cells(vFrm.ColCostNominal.Index).Value, ObjCur.GetDefaultPrecission(PrimaryCurCost))
                                cost += ObjCD.DigitGrouping(rowDGV.Cells(vFrm.ColCostAmount.Index).Value, ObjCur.GetDefaultPrecission(PrimaryCurCost))
                            End If
                        Next
                        Credit += cost
                        DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkunAP, Debet, Credit, "", Debet, vPurchaseOriginalAmount + costOriginal, vExchangeRate)


                        'masukkan data cost yang ap + cost ke ap base
                        Credit = vVATOriginal + vDutyStampOriginal
                        cost = 0
                        costOriginal = 0
                        For Each rowDGV In vDGVCost.Rows
                            If CDbl(rowDGV.Cells(vFrm.ColPrimaryAccount.Index).Value) = PrimaryAkunAP_BaseIDR Then
                                Dim PrimaryCurCost As Double = -1
                                PrimaryCurCost = rowDGV.Cells(vFrm.ColPrimaryCurrencyCost.Index).Value
                                costOriginal += ObjCD.DigitGrouping(rowDGV.Cells(vFrm.ColCostNominal.Index).Value, ObjCur.GetDefaultPrecission(PrimaryCurCost))
                                cost += ObjCD.DigitGrouping(rowDGV.Cells(vFrm.ColCostAmount.Index).Value, ObjCur.GetDefaultPrecission(PrimaryCurCost))
                            End If
                        Next

                        Credit += costOriginal
                        If vVATOriginal + vDutyStampOriginal + costOriginal <> 0 Then
                            DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkunAP_BaseIDR, Debet, Credit, "", Debet, Credit, 1)
                        End If
                    End If
                End If

                Dim CostIDR As Decimal = 0
                Dim CostBase As Decimal = 0
                Dim CostBaseOriginal As Decimal = 0
                For Each rowDGV In vDGVCost.Rows
                    If rowDGV.Cells(vFrm.ColPrimaryAccount.Index).Value <> -1 Then
                        Dim tipe1 As Integer = 0
                        tipe1 = objAkun.GetTipeAkun(rowDGV.Cells(vFrm.ColPrimaryAccount.Index).Value)
                        If tipe1 = AccModule.ClsMasterAkun.EnumTipeAkun.Liabilities Then
                            Dim PrimaryCurCost As Double = -1
                            PrimaryCurCost = rowDGV.Cells(vFrm.ColPrimaryCurrencyCost.Index).Value

                            If CDbl(rowDGV.Cells(vFrm.ColPrimaryAccount.Index).Value) = PrimaryAkunAP OrElse _
                               CDbl(rowDGV.Cells(vFrm.ColPrimaryAccount.Index).Value) = PrimaryAkunAP_BaseIDR Then

                                If CDbl(rowDGV.Cells(vFrm.ColPrimaryAccount.Index).Value) = PrimaryAkunAP Then
                                    AP_PadaCost += ObjCD.DigitGrouping(rowDGV.Cells(vFrm.ColCostAmount.Index).Value, ObjCur.GetDecimal_DefaultCurrency)
                                    AP_PadaCost_Original += ObjCD.DigitGrouping(rowDGV.Cells(vFrm.ColCostNominal.Index).Value, ObjCur.GetDefaultPrecission(PrimaryCurCost))
                                End If

                                If PrimaryAkunAP <> PrimaryAkunAP_BaseIDR AndAlso _
                                   CDbl(rowDGV.Cells(vFrm.ColPrimaryAccount.Index).Value) = PrimaryAkunAP_BaseIDR Then
                                    AP_PadaCost_Base += ObjCD.DigitGrouping(rowDGV.Cells(vFrm.ColCostAmount.Index).Value, ObjCur.GetDefaultPrecission(PrimaryCurCost))
                                End If
                            Else
                                If ListPrimaryAkunCost.Contains(rowDGV.Cells(vFrm.ColPrimaryAccount.Index).Value) = False Then
                                    ListPrimaryAkunCost.Add(rowDGV.Cells(vFrm.ColPrimaryAccount.Index).Value)
                                    Debet = 0
                                    Credit = Me.GetCostIDR(rowDGV.Cells(vFrm.ColPrimaryAccount.Index).Value, vDGVCost, vFrm)
                                    DebetOriginal = 0
                                    CreditOriginal = Credit
                                    DTA_JurnalItem.InsertQuery(PrimaryJurnal, rowDGV.Cells(vFrm.ColPrimaryAccount.Index).Value, Debet, Credit, _
                                    "", DebetOriginal, CreditOriginal, 1)
                                End If
                            End If

                        Else
                            If ListPrimaryAkunCost.Contains(rowDGV.Cells(vFrm.ColPrimaryAccount.Index).Value) = False Then
                                ListPrimaryAkunCost.Add(rowDGV.Cells(vFrm.ColPrimaryAccount.Index).Value)
                                Debet = 0
                                Credit = Me.GetCostIDR(rowDGV.Cells(vFrm.ColPrimaryAccount.Index).Value, vDGVCost, vFrm)
                                DebetOriginal = 0
                                CreditOriginal = Credit
                                DTA_JurnalItem.InsertQuery(PrimaryJurnal, rowDGV.Cells(vFrm.ColPrimaryAccount.Index).Value, Debet, Credit, _
                                "", DebetOriginal, CreditOriginal, 1)
                            End If
                        End If
                    End If
                Next

                If vPrimaryCurrency = Core_Module.PrimaryBaseCurrency Then
                    'kurangkan data hutang dengan DP
                    If vDP_Original <> 0 Then
                        Debet = 0
                        Credit = 0
                        Debet = vDP_Original
                        DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkunAP, Debet, Credit, "", Debet, Credit, 1)

                        Debet = 0
                        Credit = 0
                        Credit = vDP_Original
                        DTA_JurnalItem.InsertQuery(PrimaryJurnal, vPrimaryAccountDP, Debet, Credit, "", Debet, Credit, 1)
                    End If
                Else
                    'kurangkan data hutang dengan DP
                    If vDP_Original <> 0 Then
                        Debet = 0
                        Credit = 0
                        Debet = ObjCD.DigitGrouping(vDP_Original * vExchangeRate, DefPrecBase)
                        DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkunAP, Debet, Credit, _
                         "", vDP_Original, Credit, vExchangeRate)

                        Debet = 0
                        Credit = 0
                        Credit = vDPBaseCurrency
                        DTA_JurnalItem.InsertQuery(PrimaryJurnal, vPrimaryAccountDP, Debet, Credit, "", Debet, Credit, 1)

                        SelisihDP = ObjCD.DigitGrouping((ObjCD.DigitGrouping(vDP_Original * vExchangeRate, DefPrecBase) - vDPBaseCurrency), DefPrecBase)

                        Debet = 0
                        Credit = 0
                        If SelisihDP > 0 Then
                            Debet = 0
                            Credit = SelisihDP
                        ElseIf SelisihDP < 0 Then
                            Debet = Math.Abs(SelisihDP)
                            Credit = 0
                        End If

                        If SelisihDP <> 0 Then
                            DTA_JurnalItem.InsertQuery(PrimaryJurnal, ObjCur.GetPrimary_Akun_Realize_Exch_Rate(vPrimaryCurrency), _
                            Debet, Credit, "", Debet, Credit, 1)
                        End If
                    End If
                End If



                If vPrimaryCurrency = Core_Module.PrimaryBaseCurrency Then
                    '========================================================================================================================================
                    'JIKA TRANSAKSI IDR
                    '========================================================================================================================================
                    'aop idr
                    DTA_Balance.InsertQuery(vPrimarySupplier, vPrimaryAOP, TipeBalance, "", #1/1/1950#, "", vTanggalAOP + TimeSpan.FromDays(vPaymentTerm), _
                    (vPurchase - vTotalCostIDR) + AP_PadaCost - vDP_Original, vVAT, vDutystamp, (vPurchase - vTotalCostIDR) + vVAT + vDutystamp + AP_PadaCost - vDP_Original, 0, vUserCode, Now, vUserCode, Now, Now, Core_Module.PrimaryBaseCurrency, _
                    (vPurchase - vTotalCostIDR) + vVAT + vDutystamp + AP_PadaCost_Original - vDP_Original, 1, PrimaryAkunAP)


                    'buat kartu ap baru
                    ObjKartu.Simpan(vPrimaryAOP, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, vAOPNumber, _
                    vPrimarySupplier, 0, (vPurchase - vTotalCostIDR) + vVAT + vDutystamp + AP_PadaCost - vDP_Original, "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
                    vTanggalAOP, vPrimaryAOP, vTanggalAOP, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, _
                    Core_Module.PrimaryBaseCurrency, 0, (vPurchase - vTotalCostIDR) + vVAT + vDutystamp + AP_PadaCost_Original - vDP_Original, PrimaryAkunAP_BaseCurrency, 1)
                Else

                    '========================================================================================================================================
                    'JIKA TRANSAKSI NON IDR
                    '========================================================================================================================================
                    If vPrimaryCurrTax = vPrimaryCurrency Then
                        '----------------------------------------------------------------------------------------------------------------------------------
                        'JIKA Tax NON IDR ---------------------------------------------------------------------------------------------------------------
                        '----------------------------------------------------------------------------------------------------------------------------------
                        'balance purchase amount
                        DTA_Balance.InsertQuery(vPrimarySupplier, vPrimaryAOP, TipeBalance, "", #1/1/1950#, "", vTanggalAOP + TimeSpan.FromDays(vPaymentTerm), _
                        (vPurchase - vTotalCostIDR) + AP_PadaCost - ObjCD.DigitGrouping(vDP_Original * vExchangeRate, DefPrecBase), vVAT, vDutystamp, _
                        (vPurchase - vTotalCostIDR) + AP_PadaCost + vVAT + vDutystamp - ObjCD.DigitGrouping(vDP_Original * vExchangeRate, DefPrecBase), _
                        0, vUserCode, Now, vUserCode, Now, Now, vPrimaryCurrency, _
                         vPurchaseOriginalAmount + AP_PadaCost_Original + vVATOriginal + vDutyStampOriginal - vDP_Original, vExchangeRate, _
                        PrimaryAkunAP)

                        'buat kartu ar purchase amount
                        ObjKartu.Simpan(vPrimaryAOP, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, vAOPNumber, _
                        vPrimarySupplier, 0, _
                        (vPurchase - vTotalCostIDR) + AP_PadaCost + vVAT + vDutystamp - ObjCD.DigitGrouping(vDP_Original * vExchangeRate, DefPrecBase), _
                        "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
                        vTanggalAOP, vPrimaryAOP, vTanggalAOP, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, _
                        vPrimaryCurrency, 0, _
                        vPurchaseOriginalAmount + AP_PadaCost_Original + vVATOriginal + vDutyStampOriginal - vDP_Original, _
                        PrimaryAkunAP, vExchangeRate)
                    Else
                        '----------------------------------------------------------------------------------------------------------------------------------
                        'JIKA Tax IDR ---------------------------------------------------------------------------------------------------------------
                        '----------------------------------------------------------------------------------------------------------------------------------
                        'balance purchase amount
                        DTA_Balance.InsertQuery(vPrimarySupplier, vPrimaryAOP, TipeBalance, "", #1/1/1950#, "", vTanggalAOP + TimeSpan.FromDays(vPaymentTerm), _
                        (vPurchase - vTotalCostIDR) + AP_PadaCost - ObjCD.DigitGrouping(vDP_Original * vExchangeRate, DefPrecBase), 0, 0, _
                        (vPurchase - vTotalCostIDR) + AP_PadaCost - ObjCD.DigitGrouping(vDP_Original * vExchangeRate, DefPrecBase), _
                        0, vUserCode, Now, vUserCode, Now, Now, vPrimaryCurrency, _
                         vPurchaseOriginalAmount + AP_PadaCost_Original - vDP_Original, vExchangeRate, _
                        PrimaryAkunAP)

                        'buat kartu ar purchase amount
                        ObjKartu.Simpan(vPrimaryAOP, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, vAOPNumber, _
                        vPrimarySupplier, 0, _
                        (vPurchase - vTotalCostIDR) + AP_PadaCost - ObjCD.DigitGrouping(vDP_Original * vExchangeRate, DefPrecBase), _
                        "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
                        vTanggalAOP, vPrimaryAOP, vTanggalAOP, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, _
                        vPrimaryCurrency, 0, _
                        vPurchaseOriginalAmount + AP_PadaCost_Original - vDP_Original, _
                        PrimaryAkunAP, vExchangeRate)

                        If vVATOriginal + vDutyStampOriginal + AP_PadaCost_Base <> 0 Then
                            'balance purchase tax
                            DTA_Balance.InsertQuery(vPrimarySupplier, vPrimaryAOP, TipeBalance, "", #1/1/1950#, "", vTanggalAOP + TimeSpan.FromDays(vPaymentTerm), _
                            0, vVATOriginal + AP_PadaCost_Base, vDutyStampOriginal, vVATOriginal + AP_PadaCost_Base + vDutyStampOriginal, 0, vUserCode, Now, vUserCode, Now, Now, vPrimaryCurrTax, _
                             vVATOriginal + vDutyStampOriginal + AP_PadaCost_Base, 1, PrimaryAkunAP_BaseCurrency)

                            'buat kartu ar purchase tax
                            ObjKartu.Simpan(vPrimaryAOP, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, vAOPNumber, _
                            vPrimarySupplier, 0, vVATOriginal + vDutyStampOriginal + AP_PadaCost_Base, _
                            "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
                            vTanggalAOP, vPrimaryAOP, vTanggalAOP, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, _
                            vPrimaryCurrTax, 0, _
                           vVATOriginal + vDutyStampOriginal + AP_PadaCost_Base, PrimaryAkunAP_BaseCurrency, 1)
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_Balance.Dispose()
            DTA_Jurnal.Dispose()
            DTAInv.Dispose()
            TBLInv.Dispose()
            DTA_JurnalItem.Dispose()
            DTAAkun.Dispose()
            ObjSupplier = Nothing
            ObjKartu = Nothing
            ObjSetting = Nothing
            ListAkunInv = Nothing
            ObjCD = Nothing
            ListPrimaryAkunCost = Nothing
        End Try
    End Sub


    Private Sub TransferJurnal1(ByVal vPrimaryAOP As Double, ByVal vPrimarySupplier As Double, ByVal vPurchase As Decimal, ByVal vVAT As Decimal, _
 ByVal vDutystamp As Decimal, ByVal vUrutanAOP As Double, ByVal vTanggalAOP As Date, ByVal vAOPNumber As String, _
 ByVal vUserCode As String, ByVal vPaymentTerm As Double, ByVal vTotalCostIDR As Decimal, _
 ByVal vDGVCost As DataGridView, ByVal vFrm As FrmTransaksiAOP_Quick2, ByVal vSupplierInvoiceNumber As String, _
 ByVal vPurchaseOriginalAmount As Decimal, ByVal vExchangeRate As Decimal, ByVal vPrimaryCurrency As Double, _
 ByVal vPrimaryAccountAP As Double, ByVal vPrimaryCurrTax As Double, _
  ByVal vVATOriginal As Decimal, ByVal vDutyStampOriginal As Decimal)
        Dim DTA_Balance As New DseTransaksiAOP_TransfJurnalTableAdapters.tbltrans_balance_purchaseTableAdapter
        Dim TipeBalance As EnumTipePurchaseBalance = EnumTipePurchaseBalance.Automatic
        Dim DTA_Jurnal As New DseTransaksiAOP_TransfJurnalTableAdapters.tbltrans_jurnalTableAdapter
        Dim DTA_JurnalItem As New DseTransaksiAOP_TransfJurnalTableAdapters.tbltrans_jurnal_itemTableAdapter
        Dim TipeJurnal As EnumTipeJournal = EnumTipeJournal.PurchaseJournal
        Dim RefNumber As String = CStr(vUrutanAOP)
        Dim PrimaryJurnal As Double
        Dim Debet, Credit, DebetOriginal, CreditOriginal As Decimal
        Dim rowDGV As DataGridViewRow
        Dim DTAAkun As New DseTransaksiAOP_TransfJurnalTableAdapters.tblmaster_akunTableAdapter
        Dim ObjSupplier As New Core_Module.clsMasterSupplier
        Dim Local As Boolean = False
        Dim LocalStr As String = ""
        Dim ObjKartu As New ClsTransaksi_Kartu_APAR
        Dim SuppCode, SuppName As String
        Dim MonthStr As String = ""
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim AP_PadaCost As Decimal = 0
        Dim AP_PadaCost_Original As Decimal = 0
        Dim AP_PadaCost_Base As Decimal = 0
        Dim PrimaryAkunAP As Double = -1
        Dim PrimaryAkunAP_BaseCurrency As Double = -1
        Dim PrimaryAkunAP_BaseIDR As Double = -1
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim PrimaryAkun As Double = -1
        Dim DTAInv As New DseTransaksiAOP_TransfJurnalTableAdapters.TBLInventoryTableAdapter
        Dim TBLInv As New DseTransaksiAOP_TransfJurnal.TBLInventoryDataTable
        Dim rowInv As DseTransaksiAOP_TransfJurnal.TBLInventoryRow
        Dim ListAkunInv As New List(Of Double)

        Dim ListPrimaryAkunCost As New List(Of Double)
        Dim ObjMulti As New AccModule.ClsMultiAccount
        Dim objAkun As New AccModule.ClsMasterAkun

        Try
            DTAInv.Connection.ConnectionString = Core_Module.GetConstring
            DTA_Jurnal.Connection.ConnectionString = DTAInv.Connection.ConnectionString
            DTA_Balance.Connection.ConnectionString = DTAInv.Connection.ConnectionString
            DTAAkun.Connection.ConnectionString = DTAInv.Connection.ConnectionString

            'hapuskan data jurnal lama
            DTA_Jurnal.Delete(TipeJurnal, vPrimaryAOP)
            'hapuskan data balance sales lama
            DTA_Balance.Delete(vPrimaryAOP, TipeBalance)
            Local = ObjSupplier.ApakahLokal(vPrimarySupplier)
            If Local = True Then
                LocalStr = "L"
            Else
                LocalStr = "I"
            End If

            'hapus data kartu AP dan AR
            ObjKartu.Hapus(vPrimaryAOP, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis)

            If Me.ApakahTransferOtomatis(vTanggalAOP) = True Then
                SuppCode = ObjSupplier.GetSupplierCode(vPrimarySupplier)
                SuppName = ObjSupplier.GetSupplierName(vPrimarySupplier)
                'buat jurnal baru
                If RefNumber.Length < 6 Then
                    While RefNumber.Length < 6
                        RefNumber = "0" + RefNumber
                    End While
                End If
                MonthStr = vTanggalAOP.Month.ToString
                While MonthStr.Length < 2
                    MonthStr = "0" + MonthStr
                End While
                'data header
                DTA_Jurnal.Insert("PJ-" + RefNumber + "/" + LocalStr + "/" + MonthStr + "/" + vTanggalAOP.Year.ToString, vTanggalAOP, _
                vUrutanAOP, TipeJurnal, vPrimaryAOP, vPrimarySupplier, "AOP: " + vAOPNumber _
                + " - " + SuppCode + " - " + SuppName, vUserCode, _
                Now, vUserCode, Now, Now)

                'data item
                PrimaryJurnal = DTA_Jurnal.ScalarQuery_GetPrimaryJurnal(TipeJurnal, vPrimaryAOP).Value

                'simpan data purchase-inventory-debet
                DTAInv.Fill(TBLInv, vPrimaryAOP)

                'Dim decPurchase As Decimal = 0
                'decPurchase = vPurchase
                'For intA As Integer = 0 To TBLInv.Rows.Count - 1
                '    rowInv = TBLInv.Rows(intA)
                '    If intA <> TBLInv.Rows.Count - 1 Then
                '        decPurchase -= rowInv.At_Amount
                '    Else
                '        rowInv.At_Amount = decPurchase
                '    End If
                'Next


                For Each rowInv In TBLInv.Rows
                    If ListAkunInv.Contains(rowInv.Primary_Inventory_Account) = False Then
                        ListAkunInv.Add(rowInv.Primary_Inventory_Account)
                        Debet = Me.GetInventoryPurchased(TBLInv, rowInv.Primary_Inventory_Account, vPurchase)
                        Credit = 0
                        DTA_JurnalItem.InsertQuery(PrimaryJurnal, rowInv.Primary_Inventory_Account, Debet, Credit, "", Debet, Credit, 1)
                    End If
                Next



                'simpan data vat in-debet
                If vVAT <> 0 Then
                    PrimaryAkun = Me.GetPrimaryAkunVATIn
                    If PrimaryAkun <> -1 Then
                        If vPrimaryCurrTax = vPrimaryCurrency Then
                            DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkun, vVAT, 0, "", vVAT, 0, 1)
                        Else
                            DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkun, vVATOriginal, 0, "", vVATOriginal, 0, 1)
                        End If
                    End If
                End If

                'simpan data duty stamp-debet
                If vDutystamp <> 0 Then
                    PrimaryAkun = Me.GetPrimaryDutyStamp
                    If PrimaryAkun <> -1 Then
                        If vPrimaryCurrTax = vPrimaryCurrency Then
                            DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkun, vDutystamp, 0, "", vDutystamp, 0, 1)
                        Else
                            DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkun, vDutyStampOriginal, 0, "", vDutyStampOriginal, 0, 1)
                        End If
                    End If
                End If

                ''simpan data vat in-debet
                'If vVAT <> 0 Then
                '    PrimaryAkun = Me.GetPrimaryAkunVATIn
                '    If PrimaryAkun <> -1 Then
                '        DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkun, vVAT, 0, "", vVAT, 0, 1)
                '    End If
                'End If

                ''simpan data duty stamp-debet
                'If vDutystamp <> 0 Then
                '    PrimaryAkun = Me.GetPrimaryDutyStamp
                '    If PrimaryAkun <> -1 Then
                '        DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkun, vDutystamp, 0, "", vDutystamp, 0, 1)
                '    End If
                'End If
                '========================================================================================
                'Code Erna
                '========================================================================================
                'simpan data ap, kredit
                PrimaryAkunAP = vPrimaryAccountAP
                'PrimaryAkunAP = Me.GetPrimaryAkunAP(vPrimaryAOP)
                'PrimaryAkunAP_BaseCurrency = ObjCur.GetPrimaryAkunAP(Core_Module.PrimaryBaseCurrency)
                '========================================================================================

                PrimaryAkunAP_BaseCurrency = ObjMulti.GetPrimaryAkunAPBase(vPrimaryCurrency)
                PrimaryAkunAP_BaseIDR = ObjCur.GetPrimaryAkunAP(Core_Module.PrimaryBaseCurrency)
                Debet = 0
                Credit = 0
                Dim costOriginal As Decimal = 0
                Dim cost As Decimal = 0

                If vPrimaryCurrency = Core_Module.PrimaryBaseCurrency Then
                    '========================================================================================================================================
                    'JIKA TRANSAKSI IDR
                    '========================================================================================================================================
                    Credit = (vPurchase - vTotalCostIDR) + vVAT + vDutystamp
                    'masukkan data cost yang ap + cost
                    cost = 0
                    costOriginal = 0
                    For Each rowDGV In vDGVCost.Rows
                        If CDbl(rowDGV.Cells(vFrm.ColPrimaryAccount.Index).Value) = PrimaryAkunAP Then
                            costOriginal += rowDGV.Cells(vFrm.ColCostNominal.Index).Value
                            cost += rowDGV.Cells(vFrm.ColCostAmount.Index).Value
                        End If
                    Next
                    Credit += cost
                    DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkunAP, Debet, Credit, "", Debet, vPurchaseOriginalAmount + vVATOriginal + vDutyStampOriginal + costOriginal, vExchangeRate)
                Else
                    '========================================================================================================================================
                    'JIKA TRANSAKSI NON IDR
                    '========================================================================================================================================
                    If vPrimaryCurrTax = vPrimaryCurrency Then
                        '----------------------------------------------------------------------------------------------------------------------------------
                        'JIKA Tax NON IDR ---------------------------------------------------------------------------------------------------------------
                        '----------------------------------------------------------------------------------------------------------------------------------
                        Credit = (vPurchase - vTotalCostIDR) + vVAT + vDutystamp
                        'masukkan data cost yang ap + cost ke ap base
                        cost = 0
                        costOriginal = 0
                        For Each rowDGV In vDGVCost.Rows
                            If CDbl(rowDGV.Cells(vFrm.ColPrimaryAccount.Index).Value) = PrimaryAkunAP Then
                                costOriginal += rowDGV.Cells(vFrm.ColCostNominal.Index).Value
                                cost += rowDGV.Cells(vFrm.ColCostAmount.Index).Value
                            End If
                        Next
                        Credit += cost
                        DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkunAP, Debet, Credit, "", Debet, vPurchaseOriginalAmount + vVATOriginal + vDutyStampOriginal + costOriginal, vExchangeRate)
                    Else
                        '----------------------------------------------------------------------------------------------------------------------------------
                        'JIKA Tax IDR ------------------------------------------------------------------------------------------------------------------
                        '----------------------------------------------------------------------------------------------------------------------------------
                        Credit = (vPurchase - vTotalCostIDR)
                        'masukkan data cost yang ap + cost ke ap base
                        cost = 0
                        costOriginal = 0
                        For Each rowDGV In vDGVCost.Rows
                            If CDbl(rowDGV.Cells(vFrm.ColPrimaryAccount.Index).Value) = PrimaryAkunAP Then
                                costOriginal += rowDGV.Cells(vFrm.ColCostNominal.Index).Value
                                cost += rowDGV.Cells(vFrm.ColCostAmount.Index).Value
                            End If
                        Next
                        Credit += cost
                        DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkunAP, Debet, Credit, "", Debet, vPurchaseOriginalAmount + costOriginal, vExchangeRate)

                        'masukkan data cost yang ap + cost ke ap base
                        Credit = vVATOriginal + vDutyStampOriginal
                        cost = 0
                        costOriginal = 0
                        For Each rowDGV In vDGVCost.Rows
                            If CDbl(rowDGV.Cells(vFrm.ColPrimaryAccount.Index).Value) = PrimaryAkunAP_BaseIDR Then
                                costOriginal += rowDGV.Cells(vFrm.ColCostNominal.Index).Value
                                cost += rowDGV.Cells(vFrm.ColCostAmount.Index).Value
                            End If
                        Next
                        Credit += costOriginal
                        If vVATOriginal + vDutyStampOriginal + costOriginal <> 0 Then
                            DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkunAP_BaseIDR, Debet, Credit, "", Debet, Credit, 1)
                        End If
                    End If
                End If



                'If vPrimaryCurrency = Core_Module.PrimaryBaseCurrency Then
                '    '========================================================================================================================================
                '    'JIKA TRANSAKSI IDR
                '    '========================================================================================================================================
                '    If PrimaryAkunAP = PrimaryAkunAP_BaseCurrency Then
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '        'JIKA APACCOUNT yg dipilih = base -------------------------------------------------------------------------------------------------
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '        Credit = vPurchase + vVAT + vDutystamp
                '        'masukkan data cost yang ap + cost
                '        cost = 0
                '        costOriginal = 0
                '        For Each rowDGV In vDGVCost.Rows
                '            If CDbl(rowDGV.Cells(vFrm.ColPrimaryAccount.Index).Value) = PrimaryAkunAP Then
                '                costOriginal += rowDGV.Cells(vFrm.ColCostNominal.Index).Value
                '                cost += rowDGV.Cells(vFrm.ColCostAmount.Index).Value
                '            End If
                '        Next
                '        Credit += cost
                '        DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkunAP, Debet, Credit, "", Debet, vPurchaseOriginalAmount + vVATOriginal + vDutyStampOriginal + costOriginal, vExchangeRate)
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '    Else
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '        'JIKA APACCOUNT yg dipilih <> base -------------------------------------------------------------------------------------------------
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '        Credit = vPurchase
                '        'masukkan data cost yang ap + cost
                '        cost = 0
                '        costOriginal = 0
                '        For Each rowDGV In vDGVCost.Rows
                '            If CDbl(rowDGV.Cells(vFrm.ColPrimaryAccount.Index).Value) = PrimaryAkunAP Then
                '                costOriginal += rowDGV.Cells(vFrm.ColCostNominal.Index).Value
                '                cost += rowDGV.Cells(vFrm.ColCostAmount.Index).Value
                '            End If
                '        Next
                '        Credit += cost
                '        DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkunAP, Debet, Credit, "", Debet, vPurchaseOriginalAmount + costOriginal, vExchangeRate)

                '        'masukkan data vat + dutystamp ke ap base
                '        Credit = vVAT + vDutystamp
                '        'masukkan data cost yang ap + cost ke ap base
                '        cost = 0
                '        costOriginal = 0
                '        For Each rowDGV In vDGVCost.Rows
                '            If CDbl(rowDGV.Cells(vFrm.ColPrimaryAccount.Index).Value) = PrimaryAkunAP_BaseCurrency Then
                '                costOriginal += rowDGV.Cells(vFrm.ColCostNominal.Index).Value
                '                cost += rowDGV.Cells(vFrm.ColCostAmount.Index).Value
                '            End If
                '        Next
                '        Credit += cost

                '        If vVAT + vDutystamp + cost <> 0 Then
                '            DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkunAP_BaseCurrency, Debet, Credit, "", Debet, vVATOriginal + vDutyStampOriginal + costOriginal, vExchangeRate)
                '        End If
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '    End If
                'Else
                '    '========================================================================================================================================
                '    'JIKA TRANSAKSI NON IDR
                '    '========================================================================================================================================
                '    If PrimaryAkunAP = PrimaryAkunAP_BaseCurrency Then
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '        'JIKA APACCOUNT yg dipilih = base -------------------------------------------------------------------------------------------------
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '        If vPrimaryCurrTax = vPrimaryCurrency Then
                '            '----------------------------------------------------------------------------------------------------------------------------------
                '            'JIKA Tax NON IDR ---------------------------------------------------------------------------------------------------------------
                '            '----------------------------------------------------------------------------------------------------------------------------------
                '            Credit = vPurchase + vVAT + vDutystamp
                '            'masukkan data cost yang ap + cost ke ap base
                '            cost = 0
                '            costOriginal = 0
                '            For Each rowDGV In vDGVCost.Rows
                '                If CDbl(rowDGV.Cells(vFrm.ColPrimaryAccount.Index).Value) = PrimaryAkunAP Then
                '                    costOriginal += rowDGV.Cells(vFrm.ColCostNominal.Index).Value
                '                    cost += rowDGV.Cells(vFrm.ColCostAmount.Index).Value
                '                End If
                '            Next
                '            Credit += cost
                '            DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkunAP, Debet, Credit, "", Debet, vPurchaseOriginalAmount + vVATOriginal + vDutyStampOriginal + costOriginal, vExchangeRate)
                '        Else
                '            '----------------------------------------------------------------------------------------------------------------------------------
                '            'JIKA Tax IDR ------------------------------------------------------------------------------------------------------------------
                '            '----------------------------------------------------------------------------------------------------------------------------------
                '            Credit = vPurchase
                '            'masukkan data cost yang ap + cost ke ap base
                '            cost = 0
                '            costOriginal = 0
                '            For Each rowDGV In vDGVCost.Rows
                '                If CDbl(rowDGV.Cells(vFrm.ColPrimaryAccount.Index).Value) = PrimaryAkunAP Then
                '                    costOriginal += rowDGV.Cells(vFrm.ColCostNominal.Index).Value
                '                    cost += rowDGV.Cells(vFrm.ColCostAmount.Index).Value
                '                End If
                '            Next
                '            Credit += cost
                '            DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkunAP, Debet, Credit, "", Debet, vPurchaseOriginalAmount + costOriginal, vExchangeRate)


                '            'masukkan data cost yang ap + cost ke ap base
                '            Credit = vVATOriginal + vDutyStampOriginal
                '            cost = 0
                '            costOriginal = 0
                '            For Each rowDGV In vDGVCost.Rows
                '                If CDbl(rowDGV.Cells(vFrm.ColPrimaryAccount.Index).Value) = PrimaryAkunAP_BaseIDR Then
                '                    costOriginal += rowDGV.Cells(vFrm.ColCostNominal.Index).Value
                '                    cost += rowDGV.Cells(vFrm.ColCostAmount.Index).Value
                '                End If
                '            Next
                '            Credit += costOriginal
                '            If vVATOriginal + vDutyStampOriginal + costOriginal <> 0 Then
                '                DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkunAP_BaseIDR, Debet, Credit, "", Debet, Credit, 1)
                '            End If
                '        End If
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '    Else
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '        'JIKA APACCOUNT yg dipilih <> base -------------------------------------------------------------------------------------------------
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '        If vPrimaryCurrTax = vPrimaryCurrency Then
                '            '----------------------------------------------------------------------------------------------------------------------------------
                '            'JIKA Tax NON IDR ---------------------------------------------------------------------------------------------------------------
                '            '----------------------------------------------------------------------------------------------------------------------------------
                '            '----------------------------------------------------------------------------------------------------------------------------------
                '            Credit = vPurchase
                '            'masukkan data cost yang ap + cost
                '            cost = 0
                '            costOriginal = 0
                '            For Each rowDGV In vDGVCost.Rows
                '                If CDbl(rowDGV.Cells(vFrm.ColPrimaryAccount.Index).Value) = PrimaryAkunAP Then
                '                    costOriginal += rowDGV.Cells(vFrm.ColCostNominal.Index).Value
                '                    cost += rowDGV.Cells(vFrm.ColCostAmount.Index).Value
                '                End If
                '            Next
                '            Credit += cost
                '            DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkunAP, Debet, Credit, "", Debet, vPurchaseOriginalAmount + costOriginal, vExchangeRate)

                '            Credit = vVAT + vDutystamp
                '            'masukkan data cost yang ap + cost
                '            cost = 0
                '            costOriginal = 0
                '            For Each rowDGV In vDGVCost.Rows
                '                If CDbl(rowDGV.Cells(vFrm.ColPrimaryAccount.Index).Value) = PrimaryAkunAP_BaseCurrency Then
                '                    costOriginal += rowDGV.Cells(vFrm.ColCostNominal.Index).Value
                '                    cost += rowDGV.Cells(vFrm.ColCostAmount.Index).Value
                '                End If
                '            Next
                '            Credit += cost
                '            If vVAT + vDutystamp + cost <> 0 Then
                '                DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkunAP_BaseCurrency, Debet, Credit, "", Debet, vVATOriginal + vDutyStampOriginal + costOriginal, vExchangeRate)
                '            End If
                '        Else
                '            '----------------------------------------------------------------------------------------------------------------------------------
                '            'JIKA Tax IDR ------------------------------------------------------------------------------------------------------------------
                '            '----------------------------------------------------------------------------------------------------------------------------------
                '            Credit = vPurchase
                '            'masukkan data cost yang ap + cost
                '            cost = 0
                '            costOriginal = 0
                '            For Each rowDGV In vDGVCost.Rows
                '                If CDbl(rowDGV.Cells(vFrm.ColPrimaryAccount.Index).Value) = PrimaryAkunAP Then
                '                    costOriginal += rowDGV.Cells(vFrm.ColCostNominal.Index).Value
                '                    cost += rowDGV.Cells(vFrm.ColCostAmount.Index).Value
                '                End If
                '            Next
                '            Credit += cost
                '            DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkunAP, Debet, Credit, "", Debet, vPurchaseOriginalAmount + costOriginal, vExchangeRate)

                '            Credit = vVATOriginal + vDutyStampOriginal
                '            'masukkan data cost yang ap + cost
                '            cost = 0
                '            costOriginal = 0
                '            For Each rowDGV In vDGVCost.Rows
                '                If CDbl(rowDGV.Cells(vFrm.ColPrimaryAccount.Index).Value) = PrimaryAkunAP_BaseIDR Then
                '                    costOriginal += rowDGV.Cells(vFrm.ColCostNominal.Index).Value
                '                    cost += rowDGV.Cells(vFrm.ColCostAmount.Index).Value
                '                End If
                '            Next
                '            Credit += costOriginal
                '            If vVATOriginal + vDutyStampOriginal + costOriginal <> 0 Then
                '                DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkunAP_BaseIDR, Debet, Credit, "", Debet, Credit, 1)
                '            End If
                '        End If
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '    End If
                'End If

                Dim CostIDR As Decimal = 0
                Dim CostBase As Decimal = 0
                Dim CostBaseOriginal As Decimal = 0
                For Each rowDGV In vDGVCost.Rows
                    If rowDGV.Cells(vFrm.ColPrimaryAccount.Index).Value <> -1 Then
                        Dim tipe1 As Integer = 0
                        tipe1 = objAkun.GetTipeAkun(rowDGV.Cells(vFrm.ColPrimaryAccount.Index).Value)
                        If tipe1 = AccModule.ClsMasterAkun.EnumTipeAkun.Liabilities Then
                            If CDbl(rowDGV.Cells(vFrm.ColPrimaryAccount.Index).Value) = PrimaryAkunAP OrElse _
                            CDbl(rowDGV.Cells(vFrm.ColPrimaryAccount.Index).Value) = PrimaryAkunAP_BaseIDR Then

                                If CDbl(rowDGV.Cells(vFrm.ColPrimaryAccount.Index).Value) = PrimaryAkunAP Then
                                    AP_PadaCost += CDec(rowDGV.Cells(vFrm.ColCostAmount.Index).Value)
                                    AP_PadaCost_Original += CDec(rowDGV.Cells(vFrm.ColCostNominal.Index).Value)
                                End If

                                If PrimaryAkunAP <> PrimaryAkunAP_BaseIDR AndAlso _
                              CDbl(rowDGV.Cells(vFrm.ColPrimaryAccount.Index).Value) = PrimaryAkunAP_BaseIDR Then
                                    AP_PadaCost_Base += CDec(rowDGV.Cells(vFrm.ColCostAmount.Index).Value)
                                End If

                            Else
                                If ListPrimaryAkunCost.Contains(rowDGV.Cells(vFrm.ColPrimaryAccount.Index).Value) = False Then
                                    ListPrimaryAkunCost.Add(rowDGV.Cells(vFrm.ColPrimaryAccount.Index).Value)
                                    Debet = 0
                                    Credit = Me.GetCostIDRQuick(rowDGV.Cells(vFrm.ColPrimaryAccount.Index).Value, vDGVCost, vFrm)
                                    DebetOriginal = 0
                                    CreditOriginal = Credit
                                    DTA_JurnalItem.InsertQuery(PrimaryJurnal, rowDGV.Cells(vFrm.ColPrimaryAccount.Index).Value, Debet, Credit, _
                                    "", DebetOriginal, CreditOriginal, 1)
                                End If
                            End If


                        Else
                            If ListPrimaryAkunCost.Contains(rowDGV.Cells(vFrm.ColPrimaryAccount.Index).Value) = False Then
                                ListPrimaryAkunCost.Add(rowDGV.Cells(vFrm.ColPrimaryAccount.Index).Value)
                                Debet = 0
                                Credit = Me.GetCostIDRQuick(rowDGV.Cells(vFrm.ColPrimaryAccount.Index).Value, vDGVCost, vFrm)
                                DebetOriginal = 0
                                CreditOriginal = Credit
                                DTA_JurnalItem.InsertQuery(PrimaryJurnal, rowDGV.Cells(vFrm.ColPrimaryAccount.Index).Value, Debet, Credit, _
                                "", DebetOriginal, CreditOriginal, 1)
                            End If
                        End If
                    End If
                Next



                If vPrimaryCurrency = Core_Module.PrimaryBaseCurrency Then
                    '========================================================================================================================================
                    'JIKA TRANSAKSI IDR
                    '========================================================================================================================================
                    'aop idr
                    DTA_Balance.InsertQuery(vPrimarySupplier, vPrimaryAOP, TipeBalance, "", #1/1/1950#, "", vTanggalAOP + TimeSpan.FromDays(vPaymentTerm), _
                    (vPurchase - vTotalCostIDR) + AP_PadaCost, vVAT, vDutystamp, (vPurchase - vTotalCostIDR) + vVAT + vDutystamp + AP_PadaCost, 0, vUserCode, Now, vUserCode, Now, Now, Core_Module.PrimaryBaseCurrency, _
                    (vPurchase - vTotalCostIDR) + vVAT + vDutystamp + AP_PadaCost_Original, 1, PrimaryAkunAP)
                    'DTA_Balance.InsertQuery(vPrimarySupplier, vPrimaryAOP, TipeBalance, "", #1/1/1950#, "", vTanggalAOP + TimeSpan.FromDays(vPaymentTerm), _
                    'vPurchase + AP_PadaCost, vVAT, vDutystamp, vPurchase + vVAT + vDutystamp + AP_PadaCost, 0, vUserCode, Now, vUserCode, Now, Now, Core_Module.PrimaryBaseCurrency, _
                    'vPurchase + vVAT + vDutystamp + AP_PadaCost_Original, 1, PrimaryAkunAP)

                    'buat kartu ap baru
                    ObjKartu.Simpan(vPrimaryAOP, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, vAOPNumber, _
                    vPrimarySupplier, 0, (vPurchase - vTotalCostIDR) + vVAT + vDutystamp + AP_PadaCost, "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
                    vTanggalAOP, vPrimaryAOP, vTanggalAOP, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, _
                    Core_Module.PrimaryBaseCurrency, 0, (vPurchase - vTotalCostIDR) + vVAT + vDutystamp + AP_PadaCost_Original, PrimaryAkunAP_BaseCurrency, 1)
                Else
                    '========================================================================================================================================
                    'JIKA TRANSAKSI NON IDR
                    '========================================================================================================================================
                    If vPrimaryCurrTax = vPrimaryCurrency Then
                        '----------------------------------------------------------------------------------------------------------------------------------
                        'JIKA Tax NON IDR ---------------------------------------------------------------------------------------------------------------
                        '----------------------------------------------------------------------------------------------------------------------------------
                        'balance purchase amount
                        DTA_Balance.InsertQuery(vPrimarySupplier, vPrimaryAOP, TipeBalance, "", #1/1/1950#, "", vTanggalAOP + TimeSpan.FromDays(vPaymentTerm), _
                        (vPurchase - vTotalCostIDR) + AP_PadaCost, vVAT, vDutystamp, _
                        (vPurchase - vTotalCostIDR) + AP_PadaCost + vVAT + vDutystamp, _
                        0, vUserCode, Now, vUserCode, Now, Now, vPrimaryCurrency, _
                        vPurchaseOriginalAmount + AP_PadaCost_Original + vVATOriginal + vDutyStampOriginal, vExchangeRate, _
                        PrimaryAkunAP)

                        'buat kartu ar purchase amount
                        ObjKartu.Simpan(vPrimaryAOP, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, vAOPNumber, _
                        vPrimarySupplier, 0, _
                        (vPurchase - vTotalCostIDR) + AP_PadaCost + vVAT + vDutystamp, _
                        "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
                        vTanggalAOP, vPrimaryAOP, vTanggalAOP, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, _
                        vPrimaryCurrency, 0, _
                        vPurchaseOriginalAmount + AP_PadaCost_Original + vVATOriginal + vDutyStampOriginal, _
                        PrimaryAkunAP, vExchangeRate)
                    Else
                        '----------------------------------------------------------------------------------------------------------------------------------
                        'JIKA Tax IDR ---------------------------------------------------------------------------------------------------------------
                        '----------------------------------------------------------------------------------------------------------------------------------
                        'balance purchase amount
                        'DTA_Balance.InsertQuery(vPrimarySupplier, vPrimaryAOP, TipeBalance, "", #1/1/1950#, "", vTanggalAOP + TimeSpan.FromDays(vPaymentTerm), _
                        'vPurchase + AP_PadaCost, 0, 0, _
                        'vPurchase + AP_PadaCost, _
                        '0, vUserCode, Now, vUserCode, Now, Now, vPrimaryCurrency, _
                        ' vPurchaseOriginalAmount + AP_PadaCost_Original, vExchangeRate, _
                        'PrimaryAkunAP)
                        DTA_Balance.InsertQuery(vPrimarySupplier, vPrimaryAOP, TipeBalance, "", #1/1/1950#, "", vTanggalAOP + TimeSpan.FromDays(vPaymentTerm), _
                                               (vPurchase - vTotalCostIDR) + AP_PadaCost, 0, 0, _
                                               (vPurchase - vTotalCostIDR) + AP_PadaCost, _
                                               0, vUserCode, Now, vUserCode, Now, Now, vPrimaryCurrency, _
                                               vPurchaseOriginalAmount + AP_PadaCost_Original, vExchangeRate, _
                                               PrimaryAkunAP)

                        'buat kartu ar purchase amount
                        ObjKartu.Simpan(vPrimaryAOP, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, vAOPNumber, _
                        vPrimarySupplier, 0, _
                        (vPurchase - vTotalCostIDR) + AP_PadaCost, _
                        "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
                        vTanggalAOP, vPrimaryAOP, vTanggalAOP, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, _
                        vPrimaryCurrency, 0, _
                        vPurchaseOriginalAmount + AP_PadaCost_Original, _
                        PrimaryAkunAP, vExchangeRate)

                        If vVAT + vDutystamp + AP_PadaCost_Base <> 0 Then
                            'balance purchase tax
                            DTA_Balance.InsertQuery(vPrimarySupplier, vPrimaryAOP, TipeBalance, "", #1/1/1950#, "", vTanggalAOP + TimeSpan.FromDays(vPaymentTerm), _
                            0, vVATOriginal + AP_PadaCost_Base, vDutyStampOriginal, vVATOriginal + vDutyStampOriginal + AP_PadaCost_Base, 0, vUserCode, Now, vUserCode, Now, Now, vPrimaryCurrTax, _
                             vVATOriginal + vDutyStampOriginal + AP_PadaCost_Base, 1, PrimaryAkunAP_BaseIDR)

                            'buat kartu ar purchase tax
                            ObjKartu.Simpan(vPrimaryAOP, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, vAOPNumber, _
                            vPrimarySupplier, 0, vVATOriginal + vDutyStampOriginal + AP_PadaCost_Base, _
                            "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
                            vTanggalAOP, vPrimaryAOP, vTanggalAOP, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, _
                            vPrimaryCurrTax, 0, _
                           vVATOriginal + vDutyStampOriginal + AP_PadaCost_Base, PrimaryAkunAP_BaseIDR, 1)
                        End If
                    End If
                End If


                'If vPrimaryCurrency = Core_Module.PrimaryBaseCurrency Then
                '    '========================================================================================================================================
                '    'JIKA TRANSAKSI IDR
                '    '========================================================================================================================================
                '    If PrimaryAkunAP = PrimaryAkunAP_BaseCurrency Then
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '        'JIKA APACCOUNT yg dipilih = base -------------------------------------------------------------------------------------------------
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '        'aop idr
                '        DTA_Balance.InsertQuery(vPrimarySupplier, vPrimaryAOP, TipeBalance, "", #1/1/1950#, "", vTanggalAOP + TimeSpan.FromDays(vPaymentTerm), _
                '        vPurchase + AP_PadaCost, vVAT, vDutystamp, vPurchase + vVAT + vDutystamp + AP_PadaCost, 0, vUserCode, Now, vUserCode, Now, Now, Core_Module.PrimaryBaseCurrency, _
                '        vPurchase + vVAT + vDutystamp + AP_PadaCost_Original, 1, PrimaryAkunAP)

                '        'buat kartu ap baru
                '        ObjKartu.Simpan(vPrimaryAOP, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, vAOPNumber, _
                '        vPrimarySupplier, 0, vPurchase + vVAT + vDutystamp + AP_PadaCost, "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
                '        vTanggalAOP, vPrimaryAOP, vTanggalAOP, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, _
                '        Core_Module.PrimaryBaseCurrency, 0, vPurchase + vVAT + vDutystamp + AP_PadaCost_Original, PrimaryAkunAP_BaseCurrency, 1)
                '    Else
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '        'JIKA APACCOUNT yg dipilih <> base -------------------------------------------------------------------------------------------------
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '        'aop idr
                '        DTA_Balance.InsertQuery(vPrimarySupplier, vPrimaryAOP, TipeBalance, "", #1/1/1950#, "", vTanggalAOP + TimeSpan.FromDays(vPaymentTerm), _
                '        vPurchase + AP_PadaCost, 0, 0, vPurchase + AP_PadaCost, 0, vUserCode, Now, vUserCode, Now, Now, Core_Module.PrimaryBaseCurrency, _
                '        vPurchase + AP_PadaCost_Original, vExchangeRate, PrimaryAkunAP)

                '        'buat kartu ap baru
                '        ObjKartu.Simpan(vPrimaryAOP, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, vAOPNumber, _
                '        vPrimarySupplier, 0, vPurchase + AP_PadaCost, "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
                '        vTanggalAOP, vPrimaryAOP, vTanggalAOP, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, _
                '        Core_Module.PrimaryBaseCurrency, 0, vPurchase + AP_PadaCost_Original, PrimaryAkunAP_BaseCurrency, 1)

                '        If vVAT + vDutystamp + AP_PadaCost_Base <> 0 Then
                '            'TAX
                '            DTA_Balance.InsertQuery(vPrimarySupplier, vPrimaryAOP, TipeBalance, "", #1/1/1950#, "", vTanggalAOP + TimeSpan.FromDays(vPaymentTerm), _
                '            0, vVAT + AP_PadaCost_Base, vDutystamp, vVAT + vDutystamp + AP_PadaCost_Base, 0, vUserCode, Now, vUserCode, Now, Now, Core_Module.PrimaryBaseCurrency, _
                '            vVATOriginal + vDutyStampOriginal + AP_PadaCost_Base, vExchangeRate, PrimaryAkunAP_BaseCurrency)

                '            'buat kartu ap baru
                '            ObjKartu.Simpan(vPrimaryAOP, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, vAOPNumber, _
                '            vPrimarySupplier, 0, vVAT + vDutystamp + AP_PadaCost_Base, "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
                '            vTanggalAOP, vPrimaryAOP, vTanggalAOP, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, _
                '            Core_Module.PrimaryBaseCurrency, 0, vVATOriginal + vDutyStampOriginal + AP_PadaCost_Base, PrimaryAkunAP_BaseCurrency, 1)
                '        End If
                '    End If
                'Else
                '    '========================================================================================================================================
                '    'JIKA TRANSAKSI NON IDR
                '    '========================================================================================================================================
                '    'If vDP_Original <> 0 Then
                '    '    Debet = vDP_Original
                '    '    Credit = 0
                '    '    DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkunAP, ObjCD.DigitGrouping(vDP_Original * vExchangeRate, DefPrecBase), _
                '    '    Credit, "", vDP_Original, Credit, vExchangeRate)
                '    '    Debet = 0
                '    '    Credit = vDPBaseCurrency
                '    '    DTA_JurnalItem.InsertQuery(PrimaryJurnal, _
                '    '    ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
                '    '    Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_AOP_DP), Debet, Credit, "", Debet, Credit, 1)
                '    '    SelisihDP = ObjCD.DigitGrouping(vDP_Original * vExchangeRate, DefPrecBase) - vDPBaseCurrency
                '    '    If SelisihDP > 0 Then
                '    '        Debet = 0
                '    '        Credit = SelisihDP
                '    '    ElseIf SelisihDP < 0 Then
                '    '        Debet = Math.Abs(SelisihDP)
                '    '        Credit = 0
                '    '    End If
                '    '    If SelisihDP <> 0 Then
                '    '        DTA_JurnalItem.InsertQuery(PrimaryJurnal, ObjCur.GetPrimary_Akun_Realize_Exch_Rate(vPrimaryCurrency), _
                '    '        Debet, Credit, "", Debet, Credit, 1)
                '    '    End If
                '    'End If


                '    If PrimaryAkunAP = PrimaryAkunAP_BaseCurrency Then
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '        'JIKA APACCOUNT yg dipilih = base -------------------------------------------------------------------------------------------------
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '        If vPrimaryCurrTax = vPrimaryCurrency Then
                '            '----------------------------------------------------------------------------------------------------------------------------------
                '            'JIKA Tax NON IDR ---------------------------------------------------------------------------------------------------------------
                '            '----------------------------------------------------------------------------------------------------------------------------------
                '            'balance purchase amount
                '            DTA_Balance.InsertQuery(vPrimarySupplier, vPrimaryAOP, TipeBalance, "", #1/1/1950#, "", vTanggalAOP + TimeSpan.FromDays(vPaymentTerm), _
                '            vPurchase + AP_PadaCost, vVAT, vDutystamp, _
                '            vPurchase + AP_PadaCost + vVAT + vDutystamp, _
                '            0, vUserCode, Now, vUserCode, Now, Now, vPrimaryCurrency, _
                '             vPurchaseOriginalAmount + AP_PadaCost_Original + vVATOriginal + vDutyStampOriginal, vExchangeRate, _
                '            PrimaryAkunAP)

                '            'buat kartu ar purchase amount
                '            ObjKartu.Simpan(vPrimaryAOP, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, vAOPNumber, _
                '            vPrimarySupplier, 0, _
                '            vPurchase + AP_PadaCost + vVAT + vDutystamp, _
                '            "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
                '            vTanggalAOP, vPrimaryAOP, vTanggalAOP, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, _
                '            vPrimaryCurrency, 0, _
                '            vPurchaseOriginalAmount + AP_PadaCost_Original + vVATOriginal + vDutyStampOriginal, _
                '            PrimaryAkunAP, vExchangeRate)
                '        Else
                '            '----------------------------------------------------------------------------------------------------------------------------------
                '            'JIKA Tax IDR -------------------------------------------------------------------------------------------------------------------
                '            '----------------------------------------------------------------------------------------------------------------------------------
                '            'balance purchase amount
                '            DTA_Balance.InsertQuery(vPrimarySupplier, vPrimaryAOP, TipeBalance, "", #1/1/1950#, "", vTanggalAOP + TimeSpan.FromDays(vPaymentTerm), _
                '            vPurchase + AP_PadaCost, 0, 0, _
                '            vPurchase + AP_PadaCost, _
                '            0, vUserCode, Now, vUserCode, Now, Now, vPrimaryCurrency, _
                '             vPurchaseOriginalAmount + AP_PadaCost_Original, vExchangeRate, _
                '            PrimaryAkunAP)

                '            'buat kartu ar purchase amount
                '            ObjKartu.Simpan(vPrimaryAOP, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, vAOPNumber, _
                '            vPrimarySupplier, 0, _
                '            vPurchase + AP_PadaCost, _
                '            "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
                '            vTanggalAOP, vPrimaryAOP, vTanggalAOP, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, _
                '            vPrimaryCurrency, 0, _
                '            vPurchaseOriginalAmount + AP_PadaCost_Original, _
                '            PrimaryAkunAP, vExchangeRate)

                '            If vVAT + vDutystamp + AP_PadaCost_Base <> 0 Then
                '                'balance purchase tax
                '                DTA_Balance.InsertQuery(vPrimarySupplier, vPrimaryAOP, TipeBalance, "", #1/1/1950#, "", vTanggalAOP + TimeSpan.FromDays(vPaymentTerm), _
                '                0, vVAT + AP_PadaCost_Base, vDutystamp, vVAT + vDutystamp + AP_PadaCost_Base, 0, vUserCode, Now, vUserCode, Now, Now, vPrimaryCurrTax, _
                '                 vVATOriginal + vDutyStampOriginal + AP_PadaCost_Base, 1, PrimaryAkunAP_BaseIDR)

                '                'buat kartu ar purchase tax
                '                ObjKartu.Simpan(vPrimaryAOP, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, vAOPNumber, _
                '                vPrimarySupplier, 0, vVAT + vDutystamp + AP_PadaCost_Base, _
                '                "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
                '                vTanggalAOP, vPrimaryAOP, vTanggalAOP, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, _
                '                vPrimaryCurrTax, 0, _
                '               vVATOriginal + vDutyStampOriginal + AP_PadaCost_Original, PrimaryAkunAP_BaseIDR, 1)
                '            End If
                '        End If
                '    Else
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '        'JIKA APACCOUNT yg dipilih <> base -------------------------------------------------------------------------------------------------
                '        '----------------------------------------------------------------------------------------------------------------------------------
                '        If vPrimaryCurrTax = vPrimaryCurrency Then
                '            '----------------------------------------------------------------------------------------------------------------------------------
                '            'JIKA Tax Non IDR -------------------------------------------------------------------------------------------------------------------
                '            '----------------------------------------------------------------------------------------------------------------------------------
                '            'balance purchase amount
                '            DTA_Balance.InsertQuery(vPrimarySupplier, vPrimaryAOP, TipeBalance, "", #1/1/1950#, "", vTanggalAOP + TimeSpan.FromDays(vPaymentTerm), _
                '            vPurchase + AP_PadaCost, 0, 0, _
                '            vPurchase + AP_PadaCost, _
                '            0, vUserCode, Now, vUserCode, Now, Now, vPrimaryCurrency, _
                '             vPurchaseOriginalAmount + AP_PadaCost_Original, vExchangeRate, _
                '            PrimaryAkunAP)

                '            'buat kartu ar purchase amount
                '            ObjKartu.Simpan(vPrimaryAOP, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, vAOPNumber, _
                '            vPrimarySupplier, 0, _
                '            vPurchase + AP_PadaCost, _
                '            "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
                '            vTanggalAOP, vPrimaryAOP, vTanggalAOP, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, _
                '            vPrimaryCurrency, 0, _
                '            vPurchaseOriginalAmount + AP_PadaCost_Original, _
                '            PrimaryAkunAP, vExchangeRate)

                '            If vVAT + vDutystamp <> 0 Then
                '                'balance purchase tax
                '                DTA_Balance.InsertQuery(vPrimarySupplier, vPrimaryAOP, TipeBalance, "", #1/1/1950#, "", vTanggalAOP + TimeSpan.FromDays(vPaymentTerm), _
                '                0, vVAT, vDutystamp, vVAT + vDutystamp, 0, vUserCode, Now, vUserCode, Now, Now, vPrimaryCurrTax, _
                '                 vVATOriginal + vDutyStampOriginal, vExchangeRate, PrimaryAkunAP_BaseCurrency)

                '                'buat kartu ar purchase tax
                '                ObjKartu.Simpan(vPrimaryAOP, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, vAOPNumber, _
                '                vPrimarySupplier, 0, vVAT + vDutystamp, _
                '                "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
                '                vTanggalAOP, vPrimaryAOP, vTanggalAOP, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, _
                '                vPrimaryCurrTax, 0, _
                '               vVATOriginal + vDutyStampOriginal, PrimaryAkunAP_BaseCurrency, vExchangeRate)
                '            End If
                '        Else
                '            '----------------------------------------------------------------------------------------------------------------------------------
                '            'JIKA Tax IDR -------------------------------------------------------------------------------------------------------------------
                '            '----------------------------------------------------------------------------------------------------------------------------------
                '            'balance purchase amount
                '            DTA_Balance.InsertQuery(vPrimarySupplier, vPrimaryAOP, TipeBalance, "", #1/1/1950#, "", vTanggalAOP + TimeSpan.FromDays(vPaymentTerm), _
                '            vPurchase + AP_PadaCost, 0, 0, _
                '            vPurchase + AP_PadaCost, _
                '            0, vUserCode, Now, vUserCode, Now, Now, vPrimaryCurrency, _
                '             vPurchaseOriginalAmount + AP_PadaCost_Original, vExchangeRate, _
                '            PrimaryAkunAP)

                '            'buat kartu ar purchase amount
                '            ObjKartu.Simpan(vPrimaryAOP, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, vAOPNumber, _
                '            vPrimarySupplier, 0, _
                '            vPurchase + AP_PadaCost, _
                '            "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
                '            vTanggalAOP, vPrimaryAOP, vTanggalAOP, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, _
                '            vPrimaryCurrency, 0, _
                '            vPurchaseOriginalAmount + AP_PadaCost_Original, _
                '            PrimaryAkunAP, vExchangeRate)

                '            If vVATOriginal + AP_PadaCost_Base + vDutyStampOriginal <> 0 Then
                '                'balance purchase tax
                '                DTA_Balance.InsertQuery(vPrimarySupplier, vPrimaryAOP, TipeBalance, "", #1/1/1950#, "", vTanggalAOP + TimeSpan.FromDays(vPaymentTerm), _
                '                0, vVATOriginal + AP_PadaCost_Base, vDutyStampOriginal, vVATOriginal + AP_PadaCost_Base + vDutyStampOriginal, 0, vUserCode, Now, vUserCode, Now, Now, vPrimaryCurrTax, _
                '                 vVATOriginal + vDutyStampOriginal + AP_PadaCost_Base, 1, PrimaryAkunAP_BaseCurrency)

                '                'buat kartu ar purchase tax
                '                ObjKartu.Simpan(vPrimaryAOP, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, vAOPNumber, _
                '                vPrimarySupplier, 0, vVATOriginal + vDutyStampOriginal + AP_PadaCost_Base, _
                '                "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
                '                vTanggalAOP, vPrimaryAOP, vTanggalAOP, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, _
                '                vPrimaryCurrTax, 0, _
                '               vVATOriginal + vDutyStampOriginal + AP_PadaCost_Base, PrimaryAkunAP_BaseCurrency, 1)
                '            End If
                '        End If
                '    End If
                'End If

            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_Balance.Dispose()
            DTA_Jurnal.Dispose()
            DTAInv.Dispose()
            TBLInv.Dispose()
            DTA_JurnalItem.Dispose()
            DTAAkun.Dispose()
            ObjSupplier = Nothing
            ObjKartu = Nothing
            ObjSetting = Nothing
            ListAkunInv = Nothing
        End Try
    End Sub


    Public Function GetJmlAOPSama(ByVal vPrimaryPO As Double, ByVal VPrimaryAccount As Double, ByVal vPrimaryCurrTax As Double, ByVal vTipeAOP As ClsTransaksiAOP.EnumTipeAop) As Decimal
        Dim DTA As New DseTransaksiAOPTableAdapters.BatasanAOPDPTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJmlByPrimaryPONTipeNCurrNAccount(vPrimaryPO, vTipeAOP, vPrimaryCurrTax, VPrimaryAccount).HasValue AndAlso _
               DTA.ScalarQuery_GetJmlByPrimaryPONTipeNCurrNAccount(vPrimaryPO, vTipeAOP, vPrimaryCurrTax, VPrimaryAccount).Value > 0 Then
                Return DTA.ScalarQuery_GetJmlByPrimaryPONTipeNCurrNAccount(vPrimaryPO, vTipeAOP, vPrimaryCurrTax, VPrimaryAccount).Value
            Else
                Return 0
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlAOPwithPOSama(ByVal vPrimaryPO As Double, ByVal vTipeAOP As ClsTransaksiAOP.EnumTipeAop) As Decimal
        Dim DTA As New DseTransaksiAOPTableAdapters.BatasanAOPDPTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJMLByPrimaryPONTipe(vPrimaryPO, vTipeAOP).HasValue AndAlso _
            DTA.ScalarQuery_GetJMLByPrimaryPONTipe(vPrimaryPO, vTipeAOP).Value > 0 Then
                Return DTA.ScalarQuery_GetJMLByPrimaryPONTipe(vPrimaryPO, vTipeAOP).Value
            Else
                Return 0
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryAOPwithPOSama(ByVal vPrimaryPO As Double, ByVal vTipeAOP As ClsTransaksiAOP.EnumTipeAop) As Decimal
        Dim DTA As New DseTransaksiAOPTableAdapters.BatasanAOPDPTableAdapter
        Dim tbl As New DseTransaksiAOP.BatasanAOPDPDataTable
        Dim row As DseTransaksiAOP.BatasanAOPDPRow
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            DTA.FillBy_PrimaryPONTipeAOP(tbl, vPrimaryPO, vTipeAOP)
            If tbl.Rows.Count > 0 Then
                row = tbl.Rows(0)
                Return row.Primary
            End If

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
    '====================================================================================================================================================================
    'Erna's Code selesai()
    '====================================================================================================================================================================
    '====================================================================================================================================================================

    'elemen yang digunakan pada automatic invoice transfer
    Public Enum EnumTipePurchaseBalance As Integer
        Automatic
        Manual
        Manual_Transfer_Invoice ' tidak terpakai lagi
        '2 berikut ini baru ditambahkan di v.1.0.5, berkaitan penggabungan jurnal sales dan jurnal sales return.
        Automatic_SR
        Manual_SR
    End Enum

    Public Enum EnumTipeJournal As Integer
        GeneralJournal
        AdjustmenJournal
        ClosingJournal
        SalesJournal
        SalesReturnJournal
        SalesDownPayment
        PurchaseJournal
        PurchaseReturnJournal
        PurchaseDownPaymentJournal
        SalesReceivementJournal
        SalesReturnPaymentJournal
        PurchasePaymentJournal
        PurchaseReturnReceivement
    End Enum
    Public Enum Enum_Tipe_Setting
        None
        Debet
        Kredit
    End Enum

    Public Enum EnumTaxExclude As Integer
        Exclude
        Include
    End Enum

    Public Enum EnumCalculatesBy As Integer
        AOP ''Invoice
        GRN ''Delivery_Order
        PO ''Sales_Order
    End Enum

    Public Enum EnumTipeAop As Integer
        AOP
        AOP_DP
    End Enum

    Public Enum EnumLokal As Integer
        Lokal
        Import
    End Enum

    Public Property Primary() As Double
        Get
            Return Me.mprimary
        End Get
        Set(ByVal value As Double)
            Me.mprimary = value
        End Set
    End Property

    Public Property EFak_KodeJenisTransaksiPajak() As Integer
        Get
            Return Me.mEFak_KodeJenisTransaksiPajak
        End Get
        Set(ByVal value As Integer)
            mEFak_KodeJenisTransaksiPajak = value
        End Set
    End Property

    Public Property VATDipungut() As Boolean
        Get
            Return Me.mVATDiPungut
        End Get
        Set(ByVal value As Boolean)
            mVATDiPungut = value
        End Set
    End Property

    Public Property NoNPWPSupplier() As String
        Get
            Return Me.mNoNPWPSupplier
        End Get
        Set(ByVal value As String)
            mNoNPWPSupplier = value
        End Set
    End Property

    Public Property NoFakturPajak() As String
        Get
            Return Me.mNoFakturPajak
        End Get
        Set(ByVal value As String)
            mNoFakturPajak = value
        End Set
    End Property

    Public Property PaymentTerms() As Double
        Get
            Return Me.mPaymentTerms
        End Get
        Set(ByVal value As Double)
            mPaymentTerms = value
        End Set
    End Property

    Public Property NoTransaksi() As String
        Get
            Return Me.mnotransaksi
        End Get
        Set(ByVal value As String)
            Me.mnotransaksi = value
        End Set
    End Property

    Public Property SuppIdentifier() As String
        Get
            Return Me.mSuppAddressIdentifier
        End Get
        Set(ByVal value As String)
            Me.mSuppAddressIdentifier = value
        End Set
    End Property

    Public Property urutan() As Integer
        Get
            Return Me.mNoUrut
        End Get
        Set(ByVal value As Integer)
            Me.mNoUrut = value
        End Set
    End Property
    Public Property TanggalTransaksi() As Date
        Get
            Return Me.mTanggalTransaksi
        End Get
        Set(ByVal value As Date)
            Me.mTanggalTransaksi = value
        End Set
    End Property

    Public Property PrimarySupplier() As Double
        Get
            Return Me.mPrimarySupplier
        End Get
        Set(ByVal value As Double)
            Me.mPrimarySupplier = value
        End Set
    End Property

    Public Property SupplierInvoiceNo() As String
        Get
            Return Me.mSupplierInvoiceNo
        End Get
        Set(ByVal value As String)
            Me.mSupplierInvoiceNo = value
        End Set
    End Property

    Public Property PrimaryCurrency() As Double
        Get
            Return Me.mPrimaryCurrency
        End Get
        Set(ByVal value As Double)
            Me.mPrimaryCurrency = value
        End Set
    End Property


    Public Property Lokal() As EnumLokal
        Get
            Return Me.mLokal
        End Get
        Set(ByVal value As EnumLokal)
            Me.mLokal = value
        End Set
    End Property

    Public Property CalculatesBy() As EnumCalculatesBy
        Get
            Return Me.mCalculatesBy
        End Get
        Set(ByVal value As EnumCalculatesBy)
            mCalculatesBy = value
        End Set
    End Property

    Public Property SubtotalBeforeDP() As Decimal
        Get
            Return Me.mSubtotalBeforeDP
        End Get
        Set(ByVal value As Decimal)
            Me.mSubtotalBeforeDP = value
        End Set
    End Property

    Public Property DownPayment() As Decimal
        Get
            Return Me.mDownPayment
        End Get
        Set(ByVal value As Decimal)
            mDownPayment = value
        End Set
    End Property

    Public Property SubTotal() As Decimal
        Get
            Return Me.mSubTotal
        End Get
        Set(ByVal value As Decimal)
            mSubTotal = value
        End Set
    End Property

    Public Property CreatedBy() As String
        Get
            Return mCreatedBy
        End Get
        Set(ByVal value As String)
            mCreatedBy = value
        End Set
    End Property

    Public Property ModifiedBy() As String
        Get
            Return mModifiedBy
        End Get
        Set(ByVal value As String)
            mModifiedBy = value
        End Set
    End Property

    Public Property CreatedDate() As Date
        Get
            Return mCreatedDate
        End Get
        Set(ByVal value As Date)
            mCreatedDate = value
        End Set
    End Property

    Public Property ModifiedDate() As Date
        Get
            Return mModifiedDate
        End Get
        Set(ByVal value As Date)
            mModifiedDate = value
        End Set
    End Property

    Public Property TimeStamp() As Date
        Get
            Return mTimestamp
        End Get
        Set(ByVal value As Date)
            mTimestamp = value
        End Set
    End Property

    Public Property DutyStamp() As Decimal
        Get
            Return Me.mDutyStamp
        End Get
        Set(ByVal value As Decimal)
            mDutyStamp = value
        End Set
    End Property

    Public Property TaxExclude() As EnumTaxExclude
        Get
            Return Me.mTaxExclude
        End Get
        Set(ByVal value As EnumTaxExclude)
            Me.mTaxExclude = value
        End Set
    End Property

    Public Property Tax() As Decimal
        Get
            Return Me.mtax
        End Get
        Set(ByVal value As Decimal)
            mtax = value
        End Set
    End Property

    Public Property TaxExchangeRate() As Decimal
        Get
            Return Me.mTaxExchangeRate
        End Get
        Set(ByVal value As Decimal)
            mTaxExchangeRate = value
        End Set
    End Property

    Public Property TaxAmount() As Decimal
        Get
            Return Me.mTaxAmount
        End Get
        Set(ByVal value As Decimal)
            mTaxAmount = value
        End Set
    End Property

    Public Property Total() As Decimal
        Get
            Return Me.mTotal
        End Get
        Set(ByVal value As Decimal)
            mTotal = value
        End Set
    End Property

    Public Property tipeAOP() As EnumTipeAop
        Get
            Return Me.mTipeAOP
        End Get
        Set(ByVal value As EnumTipeAop)
            mTipeAOP = value
        End Set
    End Property

    Public Property ExchangRateToIDR() As Decimal
        Get
            Return Me.mExhangeRateToIDR
        End Get
        Set(ByVal value As Decimal)
            mExhangeRateToIDR = value
        End Set
    End Property

    Public Property ExhangeRateToUSD() As Decimal
        Get
            Return Me.mExchangeRateToUSD
        End Get
        Set(ByVal value As Decimal)
            Me.mExchangeRateToUSD = value
        End Set
    End Property

    Public Property Remarks() As String
        Get
            Return Me.mRemarks
        End Get
        Set(ByVal value As String)
            Me.mRemarks = value
        End Set
    End Property

    Public Property PrimaryPO4DP() As Double
        Get
            Return mPrimaryPO4Dp
        End Get
        Set(ByVal value As Double)
            mPrimaryPO4Dp = value
        End Set
    End Property

    Public Function CetakAop(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal vPrimaryAOP As Double) As Boolean
        Dim ObjSettin As New Core_Module.ClsTransaksiSetting
        If ObjSettin.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.AOP, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic) = "1" Then
            'Dim objCustom As New CustomSMESS.ClsTransaksiAOP
            'objCustom.CetakCustomAOP(CRV, vPrimaryAOP)
            'objCustom = Nothing
            Exit Function
        End If
        ObjSettin = Nothing

        Dim Report As New CrpCetakTransaksiAOPLokal
        Dim DS As New DseTransaksiAOP
        Dim FormatDay As Integer = 1
        Dim FormatMonth As Integer = 2
        Dim FormatYears As Integer = 1
        Dim FormatDecimal As Integer = 1
        Dim StockInList As New List(Of String)
        Dim rowCetak As DseTransaksiAOP.TblCetakItemLokalRow

        Try
            Dim DTA1 As New DseTransaksiAOPTableAdapters.TblHeaderLokalCetakTableAdapter
            DTA1.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Fill(DS.TblHeaderLokalCetak, vPrimaryAOP)
            DTA1.Dispose()
            Dim DTA2 As New DseTransaksiAOPTableAdapters.TblCetakItemLokalTableAdapter
            DTA2.Connection.ConnectionString = Core_Module.GetConstring
            DTA2.Fill(DS.TblCetakItemLokal, vPrimaryAOP)
            'aul
            For Each rowCetak In DS.TblCetakItemLokal
                With rowCetak
                    If .Primary_Units <> .Primary_Units_Multi Then
                        .Decimal_Multi_ = ObjProduct.GetAngkaDecimal(.Primary_Product, .Primary_Units_Multi)
                    Else
                        .Decimal_Multi_ = ._Decimal
                    End If
                End With
            Next
            '---
            Me.IsikanInformasiPO(DS.TblCetakItemLokal)
            DTA2.Dispose()

            Report.SetDataSource(DS)
            Report.SetParameterValue(Report.Parameter_FormatDay.ParameterFieldName, FormatDay)
            Report.SetParameterValue(Report.Parameter_FormatMonth.ParameterFieldName, FormatMonth)
            Report.SetParameterValue(Report.Parameter_FormatYears.ParameterFieldName, FormatYears)
            Report.SetParameterValue(Report.Parameter_FormatDecimal.ParameterFieldName, FormatDecimal)
            CRV.ReportSource = Report
            CRV.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            StockInList = Nothing
        End Try
    End Function

    Private Sub IsikanInformasiPO(ByVal vTBL As DseTransaksiAOP.TblCetakItemLokalDataTable)
        Dim rowItem As DseTransaksiAOP.TblCetakItemLokalRow
        Dim DTA As New DseTransaksiAOPTableAdapters.TblCetakItem1TableAdapter
        Dim TBL As New DseTransaksiAOP.TblCetakItem1DataTable
        Dim rowPO As DseTransaksiAOP.TblCetakItem1Row
        Dim ListPrimaryPO As New List(Of Double)
        Dim filter As String = ""
        Dim DV As DataRowView

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            DTA.ClearBeforeFill = False
            For Each rowItem In vTBL.Rows
                If ListPrimaryPO.Contains(rowItem.Primary_PO) = False Then
                    ListPrimaryPO.Add(rowItem.Primary_PO)
                    DTA.Fill(TBL, rowItem.Primary_PO)
                End If
            Next

            For Each rowPO In TBL.Rows
                filter = "[Primary PO] = " + rowPO.Primary.ToString
                vTBL.DefaultView.RowFilter = filter
                For Each DV In vTBL.DefaultView
                    rowItem = DV.Row
                    With rowItem
                        .No_PO = rowPO.No_Transaksi
                        .Currency = rowPO.Currency_Code
                        .Def_Prec_Curr = rowPO.Default_Precision
                    End With
                Next
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ListPrimaryPO = Nothing
        End Try
    End Sub

    Public Overridable Function ApakahSudahAda(ByVal vNoTransaksi As String) As Boolean
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aopTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetJmlByNoTransaksi(vNoTransaksi).Value > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetNoAOPBy_PrimaryPurchaseReturn(ByVal vPrimaryPurchaseReturn As Double) As String
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aopitem_purchasereturn_view1TableAdapter
        Dim NoAOP As String = ""

        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetNoAOP(vPrimaryPurchaseReturn)) = False Then
                Return DTA.ScalarQuery_GetNoAOP(vPrimaryPurchaseReturn)
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetSupplier_NPWPNumber(ByVal vNoAOP As String) As String
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aopTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If String.IsNullOrEmpty(DTA.ScalarQuery_GetNoNPWPSupplier(vNoAOP)) = True Then
                Return ""
            Else
                Return DTA.ScalarQuery_GetNoNPWPSupplier(vNoAOP)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetSupplier_InvoiceNumber(ByVal vNoAOP As String) As String
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aopTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If String.IsNullOrEmpty(DTA.ScalarQuery_GetSupplierInvNo(vNoAOP)) = True Then
                Return ""
            Else
                Return DTA.ScalarQuery_GetSupplierInvNo(vNoAOP)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function


    Public Function GetSubTotal(ByVal vPrimaryPO As Double, ByVal vPrimaryAOP As Double) As Decimal
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aopTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetSubTotalDP(vPrimaryPO, vPrimaryAOP).HasValue = True Then
                Return DTA.ScalarQuery_GetSubTotalDP(vPrimaryPO, vPrimaryAOP).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetNoBaru(ByVal vTanggalTransaksi As Date, ByVal vPrimarySupplier As Double) As String
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aopTableAdapter
        Dim TanggalStart, TanggalEnd As Date
        Dim NoUrut, Bulan As String
        Dim ObjSupplier As New Core_Module.clsMasterSupplier
        Dim ApakahLokal As Boolean = False
        Dim Foreign As Integer = 0
        Dim LokalStr As String = ""
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'cari tanggal sampai dan dari
            TanggalEnd = ObjCD.GetLastDateOfTheYear(vTanggalTransaksi)
            TanggalStart = ObjCD.GetFisrtDateOfTheYear(vTanggalTransaksi)

            ApakahLokal = ObjSupplier.ApakahLokal(vPrimarySupplier)
            If ApakahLokal = True Then
                Foreign = 0
                LokalStr = "L"
            Else
                Foreign = 1
                LokalStr = "I"
            End If

            'cari no urut
            If DTA.ScalarQuery_GetNoUrutanMAX(TanggalStart, TanggalEnd, Foreign).HasValue = True Then
                NoUrut = CStr(DTA.ScalarQuery_GetNoUrutanMAX(TanggalStart, TanggalEnd, Foreign).Value + 1)
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

            While Me.ApakahSudahAda(NoUrut + "/" + "AOP" + "/" + LokalStr + "/" + Bulan + "/" + CStr(vTanggalTransaksi.Year)) = True
                NoUrut = CStr(CDbl(NoUrut) + 1)
                'pastikan no urut terdiri atas 6 huruf
                While NoUrut.Length < 6
                    NoUrut = "0" + NoUrut
                End While
            End While

            Return NoUrut + "/" + "AOP" + "/" + LokalStr + "/" + Bulan + "/" + CStr(vTanggalTransaksi.Year)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            ObjSupplier = Nothing
            ObjCD = Nothing
        End Try
    End Function

    Public Overridable Function ApakahDataBenar(ByVal vBaru As Boolean, ByVal vPONumber As String, _
    ByVal vPrimaryAOPQuick As Double, Optional ByVal vCekDataBenar As Boolean = True) As Boolean
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim PrimaryAkunDummy As String = ""
        Dim ObjPO As New ClsTransaksiPO

        Try
            If vCekDataBenar = False Then
                Return True
            End If

            'msgbox("belum")
            'data salah bila:
            '1. Sudah ada no transaksi yang sama sebelumnya (untuk transaksi baru dan edit BEDA)
            '       untuk transaksi baru - dan autonumber-langsung ambil nomor baru
            '       untuk transaksi edit - dan autonumber-dianggap sama karena memang no transaksi tidak bisa diedit.
            If vBaru = True Then
                If Me.ApakahSudahAda(Me.mnotransaksi) = True Then
                    Me.mnotransaksi = Me.GetNoBaru(Me.mTanggalTransaksi, Me.mPrimarySupplier)
                End If
            End If

            '2. Jika semua transaksi sales sedang dikunci

            '3. Jika sudah melewati batas penguncian transaksi

            '4. jika tidak ada item yang akan disimpan

            'berkaitan dengan jurnal
            'jika ada faktur pajak, cek apakah ada akun vat in
            If Me.mTaxAmount > 0 Then
                PrimaryAkunDummy = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_VAT_In)
                If PrimaryAkunDummy = "" Or PrimaryAkunDummy = "-1" Then
                    MsgBox("Please specify vat in account in global setting", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If
            If Me.mDutyStamp > 0 Then
                PrimaryAkunDummy = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_Dutystamp_Expense)
                If PrimaryAkunDummy = "" Or PrimaryAkunDummy = "-1" Then
                    MsgBox("Please specify stamp duty expense account in global setting", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'jika AOP quick dan PO Number yang dirujuk sudah digunakan oleh transaksi lain
            If vPONumber <> "" Then
                If ObjPO.ApakahSudahAda(vPONumber, vPrimaryAOPQuick) = True Then
                    MsgBox("PO Number duplication, please specify alternative PO number", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If



            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjSetting = Nothing
            ObjPO = Nothing
        End Try
    End Function


    Public Function ApakahDPSudahDipakai(ByVal vPrimaryAOP As Double, ByVal vTampilkanPesan As Boolean, ByVal vfrm As FrmTransaksiAOP2) As Boolean
        Dim DTA As New DseTransaksiAOPTableAdapters.BATASANDPTableAdapter
        Dim tbl As New DseTransaksiAOP.BATASANDPDataTable
        Dim row As DseTransaksiAOP.BATASANDPRow
        Dim rowDGV As New DataGridViewRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(tbl, vPrimaryAOP)
            Dim strMSG As String = ""
            For Each rowDGV In vfrm.DGVdp.Rows
                With rowDGV
                    tbl.DefaultView.RowFilter = Nothing
                    tbl.DefaultView.RowFilter = "[Primary AOP DP] = " & .Cells(vfrm.ColDP_PrimaryAOPDP.Index).Value
                    If tbl.DefaultView.Count > 0 Then
                        For intB As Integer = 0 To tbl.DefaultView.Count - 1
                            row = tbl.DefaultView(intB).Row
                            strMSG &= vbNewLine & " - AOP DP No " & .Cells(vfrm.ColDP_AOP_DP_No.Index).Value & " already used in " & Me.GetNoTransaksi(row.Primary_AOP)
                        Next
                    End If
                End With
            Next


            If strMSG <> "" Then
                If vTampilkanPesan = True Then
                    MsgBox("You cannot save data, because there's AOP DP have been redeem in another AOP no, please re-redeem your down payment: " & strMSG, MsgBoxStyle.Information, "Information")
                End If
                Return True
            Else
                Return False
            End If



        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            tbl.Dispose()
        End Try


    End Function

    Public Overridable Function ApakahDataBenar1(ByVal vBaru As Boolean, ByVal vDGVItem As DataGridView, _
 ByVal vFrm As FrmTransaksiAOP2, Optional ByVal vCekDataBenar As Boolean = True) As Boolean
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim PrimaryAkunDummy As String = ""
        Dim ObjPO As New ClsTransaksiPO
        Dim DTAGRN As New DseTransaksiGRNTableAdapters.tbltrans_grnTableAdapter
        Dim ListPrimaryGRN As New List(Of Double)
        Dim PrimaryGRN As Double = -1
        Dim rowDGV As DataGridViewRow
        Dim JmlGRN As Nullable(Of Long)
        Dim ListGRNSalah As New List(Of Double)
        Dim Dummy As String = ""
        Dim ObjGRN As New ClsTransaksiGRN
        Dim vPrimaryAOP As Nullable(Of Double)
        Dim DTA_DP As New DseTransaksiAOPTableAdapters.tbltrans_aop_itemdp1TableAdapter
        Dim TBL_DP As New DseTransaksiAOP.tbltrans_aop_itemdp1DataTable
        Dim row_DP As DseTransaksiAOP.tbltrans_aop_itemdp1Row
        Dim JmlDP As Nullable(Of Long)

        Try
            If vCekDataBenar = False Then
                Return True
            End If

            DTAGRN.Connection.ConnectionString = Core_Module.GetConstring

            If Me.ApakahDataBenar(vBaru, "", -1) = False Then
                Return False
            End If

            For Each rowDGV In vDGVItem.Rows
                If ListPrimaryGRN.Contains(rowDGV.Cells(vFrm.ColPrimaryGRN.Index).Value) = False Then
                    ListPrimaryGRN.Add(rowDGV.Cells(vFrm.ColPrimaryGRN.Index).Value)
                End If
            Next

            Dummy = ""
            For Each PrimaryGRN In ListPrimaryGRN
                JmlGRN = DTAGRN.ScalarQuery_GetJml(PrimaryGRN)
                If JmlGRN.HasValue = False Then
                    ListGRNSalah.Add(PrimaryGRN)
                Else
                    If JmlGRN.Value = 0 Then
                        ListGRNSalah.Add(PrimaryGRN)
                    End If
                End If
            Next
            If ListGRNSalah.Count > 0 Then
                If Dummy = "" Then
                    Dummy = Me.GetNoGRN(PrimaryGRN, vDGVItem, vFrm)
                Else
                    Dummy = Dummy + ", " + Me.GetNoGRN(PrimaryGRN, vDGVItem, vFrm)
                End If

                MsgBox("Invalid GRN: " + Dummy + ". Please redo item input.", _
               MsgBoxStyle.Information, "Information")
                Return False
            End If

            'pastikan semua grn yang ada didalam transaksi ini tidak disimpan oleh item lain
            Dummy = ""
            ListGRNSalah.Clear()
            For Each PrimaryGRN In ListPrimaryGRN
                vPrimaryAOP = DTAGRN.ScalarQuery_GetPrimaryAOP(PrimaryGRN)
                If vPrimaryAOP.HasValue = True Then
                    If vPrimaryAOP.Value <> Me.mprimary And vPrimaryAOP.Value <> -1 Then
                        ListGRNSalah.Add(PrimaryGRN)
                        If Dummy <> "" Then
                            Dummy = Dummy + ", " + ObjGRN.GetNoStockIn(PrimaryGRN) + _
                            "(" + Me.GetNoTransaksi(vPrimaryAOP.Value) + ")"
                        Else
                            Dummy = ObjGRN.GetNoStockIn(ObjGRN.GetNoGRN(PrimaryGRN)) + _
                            "(" + Me.GetNoTransaksi(vPrimaryAOP.Value) + ")"
                        End If
                    End If
                End If
            Next
            If ListGRNSalah.Count > 0 Then
                MsgBox("The following stock is refered by other aop: " + Dummy + ". Process cancelled.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            'cek masalah DP
            For Each rowDGV In vFrm.DGVdp.Rows
                If rowDGV.Cells(vFrm.ColDP_PrimaryAOPDP.Index).Value <> -1 Then
                    JmlDP = DTA_DP.ScalarQuery_GetJmlReference(Me.mprimary, rowDGV.Cells(vFrm.ColDP_PrimaryAOPDP.Index).Value)
                    If JmlDP.HasValue = True Then
                        If JmlDP.Value > 0 Then
                            DTA_DP.FillBy_PrimaryAOPDP(TBL_DP, rowDGV.Cells(vFrm.ColDP_PrimaryAOPDP.Index).Value)
                            row_DP = TBL_DP.Rows(0)

                            MsgBox("The down payment AOP: " + rowDGV.Cells(vFrm.ColDP_AOP_DP_No.Index).Value.ToString + " is refered by: " + _
                            Me.GetNoTransaksi(row_DP.Primary_AOP), _
                            MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    End If
                End If
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjSetting = Nothing
            ObjPO = Nothing
            DTAGRN.Dispose()
            ListPrimaryGRN = Nothing
            ListGRNSalah = Nothing
            ObjGRN = Nothing
        End Try
    End Function

    Private Function GetNoGRN(ByVal vPrimaryGRN As Double, ByVal vDGVItem As DataGridView, ByVal vFrm As FrmTransaksiAOP2) As String
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In vDGVItem.Rows
                If rowDGV.Cells(vFrm.ColPrimaryGRN.Index).Value = vPrimaryGRN Then
                    Return rowDGV.Cells(vFrm.ColGRNNO.Index).Value
                End If
            Next
            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Sub Get_PrimaryCurrency_TaxPercent_TaxExclude_TaxExchangeRate(ByVal vPrimaryAOP As Double, ByRef vTaxPercent As Decimal, _
    ByRef vTaxExclude As Boolean, ByRef vPrimaryCurrency As Double, ByRef vTaxExchangeRate As Decimal)
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aopTableAdapter
        Dim TBL As New DseTransaksiAOP.tbltrans_aopDataTable
        Dim row As DseTransaksiAOP.tbltrans_aopRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            DTA.FillBy_Primary(TBL, vPrimaryAOP)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    vTaxPercent = .Tax
                    vTaxExclude = .Tax_Exclude
                    vPrimaryCurrency = .Primary_Currency
                    vTaxExchangeRate = .Tax_Exchange_Rate
                End With
            Else
                vTaxPercent = 0
                vTaxExclude = False
                vPrimaryCurrency = -1
                vTaxExchangeRate = 1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function GetPrimaryAOP_ByPrimaryGRN(ByVal vPrimaryGRN As Double) As Double
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aop_itemTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetPrimaryAOP(vPrimaryGRN).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryAOP(vPrimaryGRN).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTipePerhitungan(ByVal vPrimaryAOP As Double) As EnumCalculatesBy
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aopTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If Not String.IsNullOrEmpty(DTA.ScalarQuery_GetCalculatedBy(vPrimaryAOP)) Then
                Return DTA.ScalarQuery_GetCalculatedBy(vPrimaryAOP)
            Else
                Return EnumCalculatesBy.AOP
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTanggalAOP(ByVal vPrimary As Double) As Date
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aopTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetTanggalTransaksi(vPrimary).HasValue = True Then
                Return DTA.ScalarQuery_GetTanggalTransaksi(vPrimary).Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahVATDipungut(ByVal vPrimaryAOP As Double) As Boolean
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aopTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetVATDipungut(vPrimaryAOP).HasValue = True Then
                If DTA.ScalarQuery_GetVATDipungut(vPrimaryAOP).Value = 1 Then
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

    Public Function ApakahAOPDP(ByVal vPrimaryAOP As Double) As Boolean
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aopTableAdapter

        Try
            If DTA.ScalarQuery_GetTipeAOP(vPrimaryAOP).HasValue = True Then
                If DTA.ScalarQuery_GetTipeAOP(vPrimaryAOP).Value = 1 Then
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

    Public Function ApakahQuickPurchase(ByVal vPrimaryAOP As Double) As Boolean
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aopTableAdapter

        Try
            If DTA.ScalarQuery_GetApakahQuickPurchase(vPrimaryAOP).HasValue = True Then
                If DTA.ScalarQuery_GetApakahQuickPurchase(vPrimaryAOP).Value = 1 Then
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

    Public Function ApakahLokal(ByVal vPrimaryAOP As Double) As Boolean
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aopTableAdapter

        Try
            If DTA.ScalarQuery_GetLokal(vPrimaryAOP).HasValue = True Then
                Return DTA.ScalarQuery_GetLokal(vPrimaryAOP).Value = True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function ApakahBolehEdit(ByVal vPrimaryTransaksi As Double, _
      ByVal vSalesStaff As String, ByVal StrictEditting As Boolean, _
      Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aopTableAdapter
        Dim DTA1 As New DseTransaksiAOPTableAdapters.tbltrans_aop_itemdpTableAdapter
        Dim TanggalDO As Date = #1/1/1950#
        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjCL As New ClsTransaksiMonthlyClosing
        Dim ObjJurnal As New AccModule.ClsTransaksiJournal
        Dim Jurnal As String = ""
        'Dim ObjCustom As New CustomSMESS.ClsTransaksiAOP
        Dim ObjPR As New AccModule.ClsTransaksiPurchaseReturn
        Dim NoPR As String = ""
        Dim ObjClosingAcc As New AccModule.ClsTransaksiClosing
        Dim TglClosingAcc As Date = Today
        Dim Dummy As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString

            'jika sudah melewati batas penguncian transaksi maka tidak bisa edit data
            'changed on 2016-07-21: request dari lmb dan setelah diconsider oleh gw dan cien, tidak ada efek ke inventory, sehingga tidak masalah
            TanggalDO = Me.GetTanggalAOP(vPrimaryTransaksi)
            If Me.ApakahAOPDP(vPrimaryTransaksi) = True Then
                'di batasi oleh tanggal closing accounting
                TglClosingAcc = ObjClosingAcc.GetClosingDate
                If TanggalDO.Date <= TglClosingAcc.Date Then
                    MsgBox("You cannot edit this data because transaction date <= accounting closing date(" + Format(TglClosingAcc, "dd MMM yyyy") + ")" _
                        , MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                If ObjCD.GetLasDateOfTheMonth(ObjCL.GetLastClosingPeriod()) >= TanggalDO Then
                    MsgBox("The period " + Format(TanggalDO, "MMM yyyy") + " has been closed, you cannot edit this data.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'msgbox("belum")
            'batasan edit invoice:
            '1. Jika strict editting maka hanya sales staff yang menciptakan yang boleh edit
            If StrictEditting = True Then
                Dim CreatedBy As String = DTA.ScalarQuery_GetCreatedBy(vPrimaryTransaksi)
                If CreatedBy = vSalesStaff Then
                Else
                    If TampilkanPesan = True Then
                        MsgBox("Only user with ID:" + CreatedBy + " can edit this data.", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If

            '2. AOP DP, maka jika sudah ada AOP lain yang merujuk ke AOPDP tersebut maka data tidak bisa diedit
            Dummy = DTA1.ScalarQuery_GetNoAOP1(vPrimaryTransaksi)
            If String.IsNullOrEmpty(Dummy) = False Then
                If TampilkanPesan = True Then
                    MsgBox("This invoice is refered as down payment by: " + _
                    Dummy + ", you cannot edit this data.", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            '3. Jika sudah ada purchase payment jurnal yang merujuk ke transaksi tersebut, maka data tidak
            '   bisa diedit
            Jurnal = ObjJurnal.Get_AOP_Jurnal(vPrimaryTransaksi)
            If Jurnal <> "" Then
                If TampilkanPesan = True Then
                    MsgBox("This data is refered by purchase payment jounal: " + Jurnal + _
                    ". You cannot edit this data.", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            '4. Jika sudah ada transfer balance AOP secara manual maka tidak boleh edit data AOP


            '5. Jika sudah pernah dicetak maka, ketika akan melakukan edit harus dengan persetujuan 
            'orang yang memiliki hak khusus

            '6. Jika sudah ada purchase return yang revisi aop, maka tidak bisa diedit,
            'karena data exchange rate, vat dipungut, data harga, semuanya dirujuk.
            NoPR = ObjPR.GetNoTransaksi_NotaRetur_ByAOP(Me.mprimary)
            If NoPR <> "" Then
                MsgBox("This aop refered by purchase return: " + NoPR + ". You cannot edit this data.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            '7. custom smess
            'If ObjCustom.ApakahBolehEdit_Custom(vPrimaryTransaksi) = False Then
            '    Return False
            'End If

            'msgbox("belum")

            'sampe sini balikan true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
            ObjCD = Nothing
            ObjCL = Nothing
            ObjJurnal = Nothing
            'ObjCustom = Nothing
            ObjPR = Nothing
            ObjClosingAcc = Nothing
        End Try
    End Function

    Public Overridable Function ApakahBolehHapus(ByVal vPrimaryTransaksi As Double, _
        ByVal StrictSalesStaffDelete As Boolean, ByVal vSalesStaff As String, _
        Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aopTableAdapter
        Dim DTA1 As New DseTransaksiAOPTableAdapters.tbltrans_aop_itemdpTableAdapter
        Dim AOPDummy As String = ""
        Dim TanggalDO As Date = #1/1/1950#
        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjCL As New ClsTransaksiMonthlyClosing
        Dim DTA_PR As New DseTransaksiAOPTableAdapters.tbltrans_aop_view1TableAdapter
        Dim TBL_PR As New DseTransaksiAOP.tbltrans_aop_view1DataTable
        Dim row_PR As DseTransaksiAOP.tbltrans_aop_view1Row
        Dim TipePR As ClsTransaksiPurchaseReturn.EnumTipe_PurchaseReturn = ClsTransaksiPurchaseReturn.EnumTipe_PurchaseReturn.Nota_Retur
        Dim PR_Dummy As String = ""
        Dim Jurnal As String = ""
        Dim ObjJurnal As New AccModule.ClsTransaksiJournal
        'Dim ObjCustom As New CustomSMESS.ClsTransaksiAOP
        Dim ObjClosingAcc As New AccModule.ClsTransaksiClosing
        Dim TglClosingAcc As Date = Today
        Dim Dummy As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString

            'jika sudah melewati batas penguncian transaksi maka tidak bisa hapus data
            TanggalDO = Me.GetTanggalAOP(vPrimaryTransaksi)
            If Me.ApakahAOPDP(vPrimaryTransaksi) = False Then
                If ObjCD.GetLasDateOfTheMonth(ObjCL.GetLastClosingPeriod()) >= TanggalDO Then
                    MsgBox("The period " + Format(TanggalDO, "MMM yyyy") + " has been closed, you cannot delete this data.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                TglClosingAcc = ObjClosingAcc.GetClosingDate
                If TanggalDO.Date <= TglClosingAcc.Date Then
                    MsgBox("You cannot delete this data because transaction date <= accounting closing date(" + Format(TglClosingAcc, "dd MMM yyyy") + ")" _
                        , MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'data SO tidak boleh dihapus jika
            '1. Jika strict sales Staff delete maka yang boleh hapus hanya sales staff yang bersangkutan
            Dim CreatedBy As String = DTA.ScalarQuery_GetCreatedBy(vPrimaryTransaksi)
            If CreatedBy = vSalesStaff Then
            Else
                If TampilkanPesan = True Then
                    MsgBox("Only user with ID:" + CreatedBy + " can delete this data.", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            '2. Jika sudah pernah dicetak, maka minta minta otorisasi sebelum menghapus invoice
            'msgbox("belum")

            '3. Tidak bisa dihapus jika sudah dirujuk oleh AOP lain sebagai item DP
            Dummy = DTA1.ScalarQuery_GetNoAOP1(vPrimaryTransaksi)
            If String.IsNullOrEmpty(Dummy) = False Then
                If TampilkanPesan = True Then
                    MsgBox("This invoice is refered as down payment by: " + _
                    Dummy + ", you cannot delete this data.", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            'If DTA1.ScalarQuery_GetJmlAOP_UsingDP(vPrimaryTransaksi).Value > 0 Then
            '    If TampilkanPesan = True Then
            '        If String.IsNullOrEmpty(DTA1.ScalarQuery_GetNoAOP_ByPrimaryDP(vPrimaryTransaksi)) = True Then
            '            AOPDummy = ""
            '        Else
            '            AOPDummy = DTA1.ScalarQuery_GetNoAOP_ByPrimaryDP(vPrimaryTransaksi)
            '        End If

            '        MsgBox("Cannot delete this data, this transaction refered by " _
            '        + AOPDummy, MsgBoxStyle.Information, "Information")
            '    End If
            '    Return False
            'End If

            '4. jika sudah dirujuk oleh transaksi purchase return--> tipe nota return
            DTA_PR.Fill(TBL_PR, vPrimaryTransaksi, TipePR)
            If TBL_PR.Rows.Count > 0 Then
                For Each row_PR In TBL_PR.Rows
                    If PR_Dummy = "" Then
                        PR_Dummy = row_PR.No_Transaksi
                    Else
                        PR_Dummy = PR_Dummy + ", " + row_PR.No_Transaksi
                    End If
                Next

                MsgBox("Cannot delete this data, this transaction refered by purchase return: " _
                + PR_Dummy, MsgBoxStyle.Information, "Information")

                Return False
            End If

            '5. jika sudah ada purchase payment jurnal maka data ini tidak bisa diedit
            Jurnal = ObjJurnal.Get_AOP_Jurnal(vPrimaryTransaksi)
            If Jurnal <> "" Then
                If TampilkanPesan = True Then
                    MsgBox("This data is refered by purchase payment jounal: " + Jurnal + _
                    ". You cannot delete this data.", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            'custom
            'If ObjCustom.ApakahBolehHapus_Custom(vPrimaryTransaksi) = False Then
            '    Return False
            'End If

            'nilai balikan true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            'DTA1.Dispose()
            ObjCD = Nothing
            ObjCL = Nothing
            DTA_PR.Dispose()
            TBL_PR.Dispose()
            DTA.Dispose()
            DTA1.Dispose()
            ObjJurnal = Nothing
            'ObjCustom = Nothing
            ObjClosingAcc = Nothing
        End Try
    End Function

    Public Function GetNoTransaksi(ByVal vPrimary As Double) As String
        Dim dta As New DseTransaksiAOPTableAdapters.tbltrans_aopTableAdapter
        Try
            If String.IsNullOrEmpty(dta.ScalarQuery_GetNoTransaksiAOP(vPrimary)) = False Then
                Return dta.ScalarQuery_GetNoTransaksiAOP(vPrimary)
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Sub ResetObject()
        Try
            mprimary = -1
            mnotransaksi = ""
            mNoUrut = 0
            mTipeAOP = EnumTipeAop.AOP
            mTanggalTransaksi = Now
            mPrimarySupplier = -1
            mSupplierInvoiceNo = ""
            mLokal = EnumLokal.Lokal
            mPrimaryCurrency = -1
            mSubtotalBeforeDP = 0
            mCalculatesBy = EnumCalculatesBy.AOP
            mDownPayment = 0
            mSubTotal = 0
            mDutyStamp = 0
            mTaxExclude = EnumTaxExclude.Exclude
            mtax = 0
            mTaxExchangeRate = 1
            mTaxAmount = 0
            mTotal = 0
            mExhangeRateToIDR = 1
            mExchangeRateToUSD = 1
            mRemarks = ""
            mPrimaryPO4Dp = -1
            mPrimaryAccount = -1
            mPrimaryCurrTax = -1

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub ResetObject(ByVal vDgvItem As DataGridView, ByVal vdgvItemDP As DataGridView, _
    ByVal vDGVExchangeRate As DataGridView, ByVal vDgvItemSR As DataGridView, ByVal vDGVItemCostData As DataGridView)
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            mprimary = -1
            mnotransaksi = ""
            mNoUrut = 0
            mTipeAOP = EnumTipeAop.AOP
            mTanggalTransaksi = Now
            mPrimarySupplier = -1
            mSupplierInvoiceNo = ""
            mLokal = EnumLokal.Lokal
            mPrimaryCurrency = -1
            mSubtotalBeforeDP = 0
            mCalculatesBy = EnumCalculatesBy.AOP
            mDownPayment = 0
            mSubTotal = 0
            mDutyStamp = 0
            mTaxExclude = EnumTaxExclude.Exclude
            mtax = 0
            mTaxExchangeRate = 1
            mTaxAmount = 0
            mTotal = 0
            mExhangeRateToIDR = 1
            mExchangeRateToUSD = 1
            mRemarks = ""
            mPrimaryPO4Dp = -1
            mPrimaryAccount = -1
            mPrimaryCurrTax = -1


            vDgvItem.Rows.Clear()
            vdgvItemDP.Rows.Clear()
            vDGVExchangeRate.Rows.Clear()
            vDgvItemSR.Rows.Clear()
            vDGVItemCostData.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCD = Nothing
        End Try
    End Sub

    Public Function GetTimestamp(ByVal vPrimary As Double) As Date
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aopTableAdapter
        Try
            If DTA.ScalarQuery_GetTimeStamp(vPrimary).HasValue = True Then
                Return DTA.ScalarQuery_GetTimeStamp(vPrimary).Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function


    Public Function GetDPNo(ByVal vPrimaryAOP As Double) As String
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aop_itemdpview3TableAdapter
        Dim TBL As New DseTransaksiAOP.tbltrans_aop_itemdpview3DataTable
        Dim rowDP As DseTransaksiAOP.tbltrans_aop_itemdpview3Row
        Dim retval As String = ""

        Try
            DTA.Fill(TBL, vPrimaryAOP)
            For Each rowDP In TBL.Rows
                If retval = "" Then
                    retval = rowDP.No_AOP
                Else
                    retval = retval + ", " + rowDP.No_AOP
                End If
            Next

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' mengembalikan primary invoice dengan no transaksi yang mirip dengan revisi sebelumnya
    ''' </summary>
    ''' <param name="vNoTransaksiLIKE"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetPrimaryByNoTransaksiLIKE(ByVal vNoTransaksiLIKE As String) As Double
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aopTableAdapter

        Try
            If DTA.ScalarQuery_GetPrimaryByNoTransaksi(vNoTransaksiLIKE + "%").HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryByNoTransaksi(vNoTransaksiLIKE + "%").Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetAOPNoByPO_n_Product(ByVal vPrimaryPO As Double, _
          ByVal vPrimaryProduct As Double) As String
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aopTableAdapter
        Dim TBL As New DseTransaksiAOP.tbltrans_aopDataTable
        Dim row As DseTransaksiAOP.tbltrans_aopRow
        Dim retval As String = ""

        Try
            If DTA.FillBy_PrimaryPOnProduct(TBL, vPrimaryPO, vPrimaryProduct) > 0 Then
                For Each row In TBL.Rows
                    If retval = "" Then
                        retval = row.No_Transaksi
                    Else
                        retval = retval + ", " + row.No_Transaksi
                    End If
                Next

                Return retval
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Sub GetAOPNoByPO_n_Product(ByVal vPrimaryPO As Double, _
    ByVal vPrimaryProduct As Double, ByVal vListNoTrans As List(Of String))
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aopTableAdapter
        Dim TBL As New DseTransaksiAOP.tbltrans_aopDataTable
        Dim row As DseTransaksiAOP.tbltrans_aopRow
        Dim retval As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            vListNoTrans.Clear()

            If DTA.FillBy_PrimaryPOnProduct(TBL, vPrimaryPO, vPrimaryProduct) > 0 Then
                For Each row In TBL.Rows
                    vListNoTrans.Add(row.No_Transaksi)
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub


    ''' <summary>
    ''' mengisikan tabel dengan invoice DP yang belum pernah diinvoice, 
    ''' yang merujuk pada SO yang primary-nya diberikan
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub IsiDgvDP(ByVal ListPrimaryPO As List(Of Double), ByVal Dgv As DataGridView, _
    ByVal vCurrencyCode As String, ByVal ObjCD As Core_Module.ClsCekData, ByVal vPrimaryAOP As Double)
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aop_itemdp_view4TableAdapter
        Dim TBL As New DseTransaksiAOP.tbltrans_aop_itemdp_view4DataTable
        Dim PrimaryPO As Double
        DTA.ClearBeforeFill = True
        Dim row As DseTransaksiAOP.tbltrans_aop_itemdp_view4Row
        Dim AdaBedaCurrency As Boolean = False
        Dim BedaStr As String = ""
        Dim count As Integer = 0
        Dim MasukkanDP As Boolean = False

        'aul add
        Dim DTA1 As New DseTransaksiAOPTableAdapters.tblbalance_purchase_dpTableAdapter
        Dim TBL1 As New DseTransaksiAOP.tblbalance_purchase_dpDataTable
        Dim row1 As DseTransaksiAOP.tblbalance_purchase_dpRow

        Try
            'kosongkan dgv
            Dgv.Rows.Clear()
            'isikan SO
            For Each PrimaryPO In ListPrimaryPO
                If DTA.Fill(TBL, PrimaryPO) > 0 Then
                    For Each row In TBL.Rows
                        MasukkanDP = False

                        With row
                            If .Jml = 1 Then
                                If .Primary_AOP = -1 Then
                                    'masukkan dp
                                    MasukkanDP = True
                                ElseIf .Primary_AOP = vPrimaryAOP Then
                                    'masukkan dp
                                    MasukkanDP = True
                                End If
                            ElseIf .Jml = 0 Then
                                'masukkan dp
                                MasukkanDP = True
                            End If
                            If MasukkanDP = True Then
                                If .Currency_Code = vCurrencyCode Then
                                    Dgv.Rows.Add("0", row.No_AOP_DP, row.Currency_Code, ObjCD.DigitGrouping(.Sub_Total, .Default_Precision), "", .PrimaryAOP, EnumTipeDP.AOPDP)
                                Else
                                    AdaBedaCurrency = True
                                    count += 1
                                    BedaStr = BedaStr + count.ToString + ". " + row.No_AOP_DP + ": " + _
                                    row.Currency_Code + " " + ObjCD.DigitGrouping(row.Sub_Total.ToString, .Default_Precision) + vbCrLf
                                End If
                            End If
                        End With
                    Next

                    If AdaBedaCurrency = True Then
                        MsgBox("Cannot insert downpayment with diffrent currency. Downpayment list:" _
                        + vbCrLf + BedaStr, MsgBoxStyle.Information, "Information")
                    End If
                End If

                'aul add
                DTA1.Connection.ConnectionString = Core_Module.GetConstring
                DTA1.Fill(TBL1, AccModule.ClsTransaksiBalanceManual_DP.EnumTipeDP.DP_AOP, PrimaryPO)
                If TBL1.Count > 0 Then
                    For Each row1 In TBL1.Rows
                        MasukkanDP = False

                        With row1
                            If .Count = 1 Then
                                If .Primary_AOP = -1 Then
                                    'masukkan dp
                                    MasukkanDP = True
                                ElseIf .Primary_AOP = vPrimaryAOP Then
                                    'masukkan dp
                                    MasukkanDP = True
                                End If
                            ElseIf .Count = 0 Then
                                'masukkan dp
                                MasukkanDP = True
                            End If
                            If MasukkanDP = True Then
                                If .Currency_Code = vCurrencyCode Then
                                    Dgv.Rows.Add("0", .No_AOP_DP, .Currency_Code, ObjCD.DigitGrouping(.Sub_Total, .Default_Precision), "", .Primary, EnumTipeDP.ManualDP)
                                Else
                                    AdaBedaCurrency = True
                                    count += 1
                                    BedaStr = BedaStr + count.ToString + ". " + .No_AOP_DP + ": " + _
                                    .Currency_Code + " " + ObjCD.DigitGrouping(.Sub_Total.ToString, .Default_Precision) + vbCrLf
                                End If
                            End If
                        End With
                    Next

                    If AdaBedaCurrency = True Then
                        MsgBox("Cannot insert downpayment with diffrent currency. Downpayment list:" _
                        + vbCrLf + BedaStr, MsgBoxStyle.Information, "Information")
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub


    Public Overridable Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseTransaksiAOP_TransfJurnal.tbltrans_aop_view0DataTable, ByVal TipeAOP As EnumTipeAop, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, ByVal vForeignSupplier As Integer, ByVal vQuickPurchase As Boolean, _
    Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
    Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Boolean
        Dim DTA1 As New DseTransaksiAOP_TransfJurnalTableAdapters.tbltrans_aop_view0TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Dim IdxColPO As Integer = 0

        Try
            DTA1.Connection.ConnectionString = Core_Module.GetConstring

            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If TipeAOP = EnumTipeAop.AOP_DP Then
                If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                    QueryResult = DTA1.FillBy_NoTransaksiLIKE2(TBL, TipeAOP, Criteria + "%")
                ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                    QueryResult = DTA1.FillBy_TglSampaiDari2(TBL, TipeAOP, vTanggalSampai, vTanggalDari)
                ElseIf SearchCriteria = EnumSearchCriteria.ByKodeSupplier Then
                    QueryResult = DTA1.FillBy_KodeSupplierLIKE2(TBL, TipeAOP, Criteria + "%")
                ElseIf SearchCriteria = EnumSearchCriteria.BySupplierName Then
                    QueryResult = DTA1.FillBy_SupplierNameLIKE2(TBL, TipeAOP, Criteria + "%")
                ElseIf SearchCriteria = EnumSearchCriteria.ByPONo Then
                    QueryResult = DTA1.FillBy_PONumberLIKE2(TBL, TipeAOP, Criteria + "%")
                End If
            Else
                If vQuickPurchase = False Then
                    If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                        QueryResult = DTA1.FillBy_NoTransaksiLIKE3(TBL, TipeAOP, Criteria + "%", vForeignSupplier)
                    ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                        QueryResult = DTA1.FillBy_TanggalSampaiDari3(TBL, TipeAOP, vForeignSupplier, vTanggalSampai, vTanggalDari)
                    ElseIf SearchCriteria = EnumSearchCriteria.ByKodeSupplier Then
                        QueryResult = DTA1.FillBy_KodeSupplierLIKE3(TBL, TipeAOP, vForeignSupplier, Criteria + "%")
                    ElseIf SearchCriteria = EnumSearchCriteria.BySupplierName Then
                        QueryResult = DTA1.FillBy_SupplierNameLIKE3(TBL, TipeAOP, vForeignSupplier, Criteria + "%")
                    ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                        QueryResult = DTA1.FillBy_ProductCodeLIKE3(TBL, TipeAOP, vForeignSupplier, Criteria + "%")
                    ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                        QueryResult = DTA1.FillBy_ProductDescLIKE3(TBL, TipeAOP, vForeignSupplier, Criteria + "%")
                    ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                        QueryResult = DTA1.FillBy_ProductSpecLIKE3(TBL, TipeAOP, vForeignSupplier, Criteria + "%")
                    ElseIf SearchCriteria = EnumSearchCriteria.ByPONo Then
                        QueryResult = DTA1.FillBy_PONumberLIKE3(TBL, TipeAOP, vForeignSupplier, Criteria + "%")
                    ElseIf SearchCriteria = EnumSearchCriteria.ByStockInDate Then
                        QueryResult = DTA1.FillBy_TanggalSampaiDari3(TBL, TipeAOP, vForeignSupplier, vTanggalSampai, vTanggalDari)
                    ElseIf SearchCriteria = EnumSearchCriteria.ByNoGRN Then
                        QueryResult = DTA1.FillBy_StockInNoLIKE3(TBL, TipeAOP, vForeignSupplier, Criteria + "%")
                    End If
                Else
                    If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                        QueryResult = DTA1.FillBy_NoTransaksiLIKE(TBL, TipeAOP, Criteria + "%")
                    ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                        QueryResult = DTA1.FillBy_TglSampaiDari(TBL, TipeAOP, vTanggalSampai, vTanggalDari)
                    ElseIf SearchCriteria = EnumSearchCriteria.ByKodeSupplier Then
                        QueryResult = DTA1.FillBy_KodeSupplierLIKE(TBL, TipeAOP, Criteria + "%")
                    ElseIf SearchCriteria = EnumSearchCriteria.BySupplierName Then
                        QueryResult = DTA1.FillBy_SupplierNameLIKE(TBL, TipeAOP, Criteria + "%")
                    ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                        QueryResult = DTA1.FillBy_ProductCodeLIKE(TBL, TipeAOP, Criteria + "%")
                    ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                        QueryResult = DTA1.FillBy_ProductDescLIKE(TBL, TipeAOP, Criteria + "%")
                    ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                        QueryResult = DTA1.FillBy_ProductSpecLIKE(TBL, TipeAOP, Criteria + "%")
                    ElseIf SearchCriteria = EnumSearchCriteria.ByPONo Then
                        QueryResult = DTA1.FillBy_PONumberLIKE(TBL, TipeAOP, Criteria + "%")
                    End If
                End If
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(2).DefaultCellStyle.Format = "dd MMM yyyy"
                IdxColPO = DGV.ColumnCount - 1
                If TipeAOP = EnumTipeAop.AOP Then
                    DGV.Columns(IdxColPO).Visible = False
                Else
                    DGV.Columns(IdxColPO).Visible = True
                End If
            End If

            Dim TS As TimeSpan = Now - StartT
            If QueryResult > 0 Then
                If TS.Ticks > 10000000 Then
                    tsl.Text = CStr(QueryResult) + " result(s) in " + CStr(TS.Ticks / TimeSpan.TicksPerSecond) + " seconds"
                Else
                    tsl.Text = CStr(QueryResult) + " result(s) in " + CStr(Math.Round(TS.Ticks / TimeSpan.TicksPerMillisecond, 0, MidpointRounding.AwayFromZero)) + " miliseconds"
                End If
            Else
                tsl.Text = "No result found"
            End If
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA1.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Function GetJml_AOPDP_ByPrimaryPO(ByVal vPrimaryPO As Double) As Double
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aop2TableAdapter
        Dim RetVal As Nullable(Of Long)
        Dim TipeAOP As EnumTipeAop = EnumTipeAop.AOP_DP

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetJmlDP_ByPrimaryPO(CInt(TipeAOP), vPrimaryPO)
            If RetVal.HasValue = False Then
                Return 0
            Else
                Return RetVal.Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal Criteria As String, _
       ByVal SearchCriteria As EnumSearchCriteria, ByVal TipeAOP As EnumTipeAop, ByVal ForeignSupplier As Integer, Optional _
       ByVal vTanggalDari As Date = (#1/1/1950#), Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Long
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aopTableAdapter
        Try
            If TipeAOP = EnumTipeAop.AOP_DP Then
                If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                    Return DTA.ScalarQuery_GetjmlByNoTransaksiLIKE(Criteria + "%", TipeAOP).Value
                ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                    Return DTA.ScalarQuery_GetJmlbyTgltransaksi(TipeAOP, vTanggalDari, vTanggalSampai).Value
                ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                    Return DTA.ScalarQuery_GetJmlByProductCodeLIKE(TipeAOP, Criteria + "%").Value
                ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                    Return DTA.ScalarQuery_GetJmlByProductDescLIKE(TipeAOP, Criteria + "%").Value
                ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                    Return DTA.ScalarQuery_GetJmlProductSpecLIKE(TipeAOP, Criteria + "%").Value
                ElseIf SearchCriteria = EnumSearchCriteria.ByKodeSupplier Then
                    Return DTA.ScalarQuery_GetJmlBySuppCodeLIKE(Criteria + "%", TipeAOP).Value
                ElseIf SearchCriteria = EnumSearchCriteria.BySupplierName Then
                    Return DTA.ScalarQuery_GetJmlBySuppNameLIKE(TipeAOP, Criteria + "%").Value
                ElseIf SearchCriteria = EnumSearchCriteria.ByPONo Then
                    Return DTA.ScalarQuery_GetJmlByPONO_LIKE(Criteria + "%").Value
                ElseIf SearchCriteria = EnumSearchCriteria.ByStockInDate Then
                    Return DTA.ScalarQuery_GetJmlByInPurchaseDate2(vTanggalDari, vTanggalSampai, TipeAOP).Value
                End If
            Else
                If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                    Return DTA.ScalarQuery_GetJmlByNoTransaksiLIKE2(Criteria + "%", TipeAOP, ForeignSupplier).Value
                ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                    Return DTA.ScalarQuery_GetJmlByTanggalTransaksi2(TipeAOP, vTanggalDari, vTanggalSampai, ForeignSupplier).Value
                ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                    Return DTA.ScalarQuery_GetJmlByProductCodeLIKE2(TipeAOP, Criteria + "%", ForeignSupplier).Value
                ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                    Return DTA.ScalarQuery_GetJmlByProductDescLIKE2(TipeAOP, Criteria + "%", ForeignSupplier).Value
                ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                    Return DTA.ScalarQuery_GetJmlByProductSpecLIKE2(TipeAOP, Criteria + "%", ForeignSupplier).Value
                ElseIf SearchCriteria = EnumSearchCriteria.ByKodeSupplier Then
                    Return DTA.ScalarQuery_GetJmlBySuppCodeLIKE2(Criteria + "%", TipeAOP, ForeignSupplier).Value
                ElseIf SearchCriteria = EnumSearchCriteria.BySupplierName Then
                    Return DTA.ScalarQuery_GetJmlBySuppNameLIKE2(TipeAOP, Criteria + "%", ForeignSupplier).Value
                ElseIf SearchCriteria = EnumSearchCriteria.BySupplierName Then
                    Return DTA.ScalarQuery_GetJmlByGRNNoLike2(TipeAOP, ForeignSupplier, Criteria + "%").Value
                ElseIf SearchCriteria = EnumSearchCriteria.ByPONo Then
                    Return DTA.ScalarQuery_GetJmlByPONO_LIKE2(TipeAOP, ForeignSupplier, Criteria + "%").Value
                ElseIf SearchCriteria = EnumSearchCriteria.ByStockInDate Then
                    Return DTA.ScalarQuery_GetJmlByInPuchaseDate(vTanggalDari, vTanggalSampai, ForeignSupplier, TipeAOP).Value
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function


    Public Function GetPrimaryPO4DP(ByVal vPrimaryAOP As Double) As Double
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aopTableAdapter
        Dim retVal As Nullable(Of Double)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            retVal = DTA.ScalarQuery_GetPrimaryPO4DP(vPrimaryAOP)
            If retVal.HasValue = True Then
                Return retVal.Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function Hapus(ByVal vPrimaryTransaksi As Double, _
        ByVal UserCode As String, ByVal StrictSalesStaffDelete As Boolean) As Boolean
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aopTableAdapter
        Dim TipeAOP As EnumTipeAop
        Dim DTA5 As New DseTransaksiAOPTableAdapters.tbltrans_aop_item1TableAdapter
        Dim ObjPOItem As New ClsTransaksiPOItem
        Dim PrimaryPO As Double = -1

        Try
            If Me.ApakahBolehHapus(vPrimaryTransaksi, StrictSalesStaffDelete, _
            UserCode) = True Then
                'hapus data jurnal
                TipeAOP = DTA.ScalarQuery_GetTipeAOP(vPrimaryTransaksi).Value
                If TipeAOP = EnumTipeAop.AOP_DP Then
                    Me.Hapus_Jurnal_DP(vPrimaryTransaksi)
                Else
                    Me.Hapus_Jurnal(vPrimaryTransaksi)
                End If

                'update primary aop pada tabel GRN
                DTA5.UpdateQuery_SetPrimaryMin(vPrimaryTransaksi)

                PrimaryPO = Me.GetPrimaryPO4DP(vPrimaryTransaksi)

                'hapus data header, data item otomatis dihapuskan oleh server
                If DTA.Delete(vPrimaryTransaksi) > 0 Then
                    'update jumlah acuan po, memastikan po bisa cancel authorize
                    If PrimaryPO <> -1 Then
                        ObjPOItem.SetOrderBalance(PrimaryPO)
                    End If

                    Return True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA5.Dispose()
            ObjPOItem = Nothing
        End Try
    End Function

    Public Function Hapus(ByVal vPrimaryTransaksi As Double, _
   ByVal UserCode As String, ByVal StrictSalesStaffDelete As Boolean, ByVal vFrm As FrmTransaksiAOP_Quick2) As Boolean
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
            Me.vPrimary_Antrian, vPrimaryTransaksi, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiAOP_Hapus, UserCode)

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
                            retval = Me.Hapus1(vPrimaryTransaksi, UserCode, StrictSalesStaffDelete, vFrm)
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

    Public Overridable Function Hapus1(ByVal vPrimaryTransaksi As Double, _
  ByVal UserCode As String, ByVal StrictSalesStaffDelete As Boolean, ByVal vFrm As FrmTransaksiAOP_Quick2, _
  Optional ByVal vCekMinusQty As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aopTableAdapter
        Dim TipeAOP As EnumTipeAop
        Dim ObjItem As New Inventory_Module.clsTransaksiStockInOutItem
        Dim PrimaryStockIn As Double = -1
        Dim PrimaryPO As Double
        Dim DTAItem As New DseTransaksiAOPTableAdapters.tbltrans_aop_itemTableAdapter
        Dim TBLItem As New DseTransaksiAOP.tbltrans_aop_itemDataTable
        Dim rowItem As DseTransaksiAOP.tbltrans_aop_itemRow
        Dim ObjGRN As New ClsTransaksiGRN
        Dim ObjPO As New ClsTransaksiPO
        Dim ObjStockIn As New Inventory_Module.clsTransaksiStockInOut
        Dim Lanjut As Boolean = True
        Dim DTA5 As New DseTransaksiAOPTableAdapters.tbltrans_aop_item1TableAdapter
        Dim ListPrimaryPO As New List(Of Double)
        Dim vTblItemCustomInOut As New CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable
        Try
            If Me.ApakahBolehHapus(vPrimaryTransaksi, StrictSalesStaffDelete, _
            UserCode) = True Then
                DTAItem.FillBy_PrimaryTransaksi(TBLItem, vPrimaryTransaksi)
                If TBLItem.Rows.Count > 0 Then
                    rowItem = TBLItem.Rows(0)
                    PrimaryPO = Me.GetPrimaryPOQuick(TBLItem)
                    PrimaryStockIn = ObjGRN.GetPrimaryStockIn(rowItem.Primary_GRN)
                    Me.GetListPrimaryPO_Refered(TBLItem, ListPrimaryPO)
                End If

                If vCekMinusQty = True Then
                    If ObjItem.ApakahHapusMengakibatkanMinus(PrimaryStockIn, Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase, True, _
                    Me.vPrimary_Antrian, Me.DTAInOUT, Me.SudahHapusAntrian) = True Then
                        Lanjut = False
                    Else
                        Lanjut = True
                    End If
                End If
                If Lanjut = True Then
                    'hapus data jurnal
                    TipeAOP = DTA.ScalarQuery_GetTipeAOP(vPrimaryTransaksi).Value
                    If TipeAOP = EnumTipeAop.AOP_DP Then
                        Me.Hapus_Jurnal_DP(vPrimaryTransaksi)
                    Else
                        Me.Hapus_Jurnal(vPrimaryTransaksi)
                    End If

                    'update primary aop pada tabel GRN
                    DTA5.UpdateQuery_SetPrimaryMin(vPrimaryTransaksi)

                    'hapus data header, data item otomatis dihapuskan oleh server
                    If DTA.Delete(vPrimaryTransaksi) > 0 Then
                        'hapuskan data stock in
                        Dim ij As New FrmTransaksiStockInPurchase2 'new
                        ObjStockIn.CariDetail(PrimaryStockIn, ij.dgvItem, vTblItemCustomInOut) 'new
                        ObjStockIn.CreatedBy = UserCode
                        ObjStockIn.Primary = PrimaryStockIn
                        ObjStockIn.Hapus1(PrimaryStockIn, UserCode, vTblItemCustomInOut, False)

                        'hapuskan data po
                        ObjPO.Hapus1(PrimaryPO, UserCode, False)

                        'hitung ulang order balance yang dirujuk oleh transaksi aop
                        If ListPrimaryPO.Count > 0 Then
                            Dim ObjPOItem As New ClsTransaksiPOItem
                            ObjPOItem.SetOrderBalance(ListPrimaryPO)
                            ObjPOItem = Nothing
                        End If

                        Return True
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTAItem.Dispose()
            TBLItem.Dispose()
            ObjItem = Nothing
            ObjGRN = Nothing
            ObjPO = Nothing
            ObjStockIn = Nothing
            DTA5.Dispose()
            ListPrimaryPO = Nothing
        End Try
    End Function

    Private Sub GetListPrimaryPO_Refered(ByVal TBLItem As DseTransaksiAOP.tbltrans_aop_itemDataTable, _
    ByVal vListPrimaryPO As List(Of Double))
        Dim rowItem As DseTransaksiAOP.tbltrans_aop_itemRow

        Try
            vListPrimaryPO.Clear()
            For Each rowItem In TBLItem.Rows
                If rowItem.Refered_To_PO = 1 Then
                    If vListPrimaryPO.Contains(rowItem.Primary_PO) = False Then
                        vListPrimaryPO.Add(rowItem.Primary_PO)
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetPrimaryPOQuick(ByVal TBLItem As DseTransaksiAOP.tbltrans_aop_itemDataTable) As Double
        Dim rowItem As DseTransaksiAOP.tbltrans_aop_itemRow

        Try
            For Each rowItem In TBLItem.Rows
                If rowItem.Refered_To_PO = 0 Then
                    Return rowItem.Primary_PO
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Overridable Function SimpanBaru(ByVal vdgvItem As DataGridView, ByVal vdgvItemDP As DataGridView, _
       ByVal vDGVItemSR As DataGridView, ByVal vDGVItemExchangeRate As DataGridView, ByVal vUserCode As String, _
       ByVal vDGVItemCostData As DataGridView, ByVal vFrm As FrmTransaksiAOP2) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            SW.Start()
            'ambil no antrian
            Me.ObjAntrin.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
            Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiAOP_SimpanBaru, vUserCode)
            'Me.GetNoAntrian()

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                If SW.Elapsed.TotalSeconds >= ObjAntrin.GetWaitTime Then
                    Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                    MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                    Exit Function
                Else
                    'hanya mencoba ke database setelah interval waktu tertentu saja
                    If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanBaru1(vdgvItem, vdgvItemDP, vDGVItemSR, vDGVItemExchangeRate, _
                            vUserCode, vDGVItemCostData, vFrm)
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
            SW = Nothing
        End Try
    End Function

    Private Function ApakahPerluSimpanPO(ByVal vFrm As FrmTransaksiAOP_Quick2, ByVal vDGV As DataGridView) As Boolean
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In vDGV.Rows
                If rowDGV.Cells(vFrm.ColReferedToPO.Index).Value = 0 Then
                    Return True
                End If
            Next

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Overridable Function SimpanBaru(ByVal vdgvItem As DataGridView, ByVal vdgvItemDP As DataGridView, _
     ByVal vDGVItemSR As DataGridView, ByVal vDGVItemExchangeRate As DataGridView, ByVal vUserCode As String, _
     ByVal vDGVItemCostData As DataGridView, ByVal vFrm As FrmTransaksiAOP_Quick2, ByVal vPONUmber As String, _
     ByVal vTblCustom As CustomSMESS.DseTransaksiAOP_Quick2_AddByCustomProduct.tbltrans_custom_aop_item1DataTable, _
     ByVal vTblItemCustomInOut As CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            SW.Start()
            'ambil no antrian
            Me.ObjAntrin.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
            Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiAOP_SimpanBaru, vUserCode)
            'Me.GetNoAntrian()

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                If SW.Elapsed.TotalSeconds >= ObjAntrin.GetWaitTime Then
                    MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                    Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                    Exit Function
                Else
                    'hanya mencoba ke database setelah interval waktu tertentu saja
                    If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanBaru1(vdgvItem, vdgvItemDP, vDGVItemSR, vDGVItemExchangeRate, _
                            vUserCode, vDGVItemCostData, vFrm, vTblCustom, vTblItemCustomInOut, True, vPONUmber)
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
            SW = Nothing
        End Try
    End Function

    Public Function GetPrimary(ByVal VnoTrans As String) As Double
        Dim dta As New DseTransaksiAOPTableAdapters.tbltrans_aopTableAdapter
        Try
            If dta.ScalarQuery_GetPrimaryByNoTransaksi(VnoTrans).HasValue Then
                Return dta.ScalarQuery_GetPrimaryByNoTransaksi(VnoTrans).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Overridable Function SimpanPerubahan(ByVal dgvItem As DataGridView, ByVal vdgvItemDP As DataGridView, _
  ByVal vDGVItemSR As DataGridView, ByVal vDGVItemExchangeRate As DataGridView, ByVal vUserCode As String, _
  ByVal vDGVItemCostData As DataGridView, ByVal vFrm As FrmTransaksiAOP_Quick2, ByVal vPONumber As String, _
  ByVal vTblCustom As CustomSMESS.DseTransaksiAOP_Quick2_AddByCustomProduct.tbltrans_custom_aop_item1DataTable, _
  ByVal vTblItemCustomInOut As CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable) As Boolean
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
            Me.vPrimary_Antrian, Me.mprimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiAOP_SimpanPerubahan, vUserCode)


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
                            retval = Me.SimpanPerubahan1(dgvItem, vdgvItemDP, vDGVItemSR, vDGVItemExchangeRate, vUserCode, _
                            vDGVItemCostData, vFrm, vPONumber, vTblCustom, vTblItemCustomInOut)
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

    Public Overridable Function SimpanPerubahan1(ByVal dgvItem As DataGridView, ByVal vdgvItemDP As DataGridView, _
    ByVal vDGVItemSR As DataGridView, ByVal vDGVItemExchangeRate As DataGridView, ByVal vUserCode As String, _
    ByVal vDGVItemCostData As DataGridView, ByVal vFrm As FrmTransaksiAOP_Quick2, ByVal vPONumber As String, _
    ByVal vTblCustom As CustomSMESS.DseTransaksiAOP_Quick2_AddByCustomProduct.tbltrans_custom_aop_item1DataTable, _
    ByVal vTblItemCustomInOut As CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable) As Boolean
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aopTableAdapter
        Dim ObjItem As New ClsTransaksiAOPItem
        'Dim TaxExclude As Integer
        'Dim Primary4DP As Double = -1
        'Dim rowdgv As DataGridViewRow
        Dim TotalCost As Decimal = 0
        Dim rowItem As DseTransaksiAOP.tbltrans_aop_itemRow
        Dim TBLItem As New DseTransaksiAOP.tbltrans_aop_itemDataTable
        Dim DTAItem As New DseTransaksiAOPTableAdapters.tbltrans_aop_itemTableAdapter
        Dim PrimaryStockIn As Double
        Dim ObjGRN As New ClsTransaksiGRN
        '\\TR = No Transaksi Admiral
        Dim NoPO As String = ""
        Dim NoStockIn As String = ""
        Try
            If Me.ApakahDataBenar(False, vPONumber, Me.mprimary) = True Then
                DTAItem.FillBy_PrimaryTransaksi(TBLItem, Me.mprimary)
                If TBLItem.Rows.Count > 0 Then
                    rowItem = TBLItem.Rows(0)
                    PrimaryStockIn = ObjGRN.GetPrimaryStockIn(rowItem.Primary_GRN)
                Else
                    PrimaryStockIn = -1
                End If
                '20190711 /fungsi keep numbering hanya di khususkan untuk, PO, Stock In  yang auto numbering saja
                Me.GetNumbering_PO_StockIn(Me.mprimary, Me.TanggalTransaksi, NoPO, NoStockIn, Me.PrimarySupplier)
                If Me.ApakahTidakMinus(PrimaryStockIn, dgvItem, vFrm) = True Then
                    If Me.Hapus1(Me.mprimary, vUserCode, False, vFrm, False) = True Then
                        Return Me.SimpanBaru1(dgvItem, vdgvItemDP, vDGVItemSR, vDGVItemExchangeRate, vUserCode, _
                        vDGVItemCostData, vFrm, vTblCustom, vTblItemCustomInOut, False, NoPO, NoStockIn)
                    End If
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mnotransaksi + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                Me.mnotransaksi = Me.GetNoBaru(Me.mTanggalTransaksi, Me.mPrimarySupplier)
                Return Me.SimpanPerubahan1(dgvItem, vdgvItemDP, vDGVItemSR, vDGVItemExchangeRate, vUserCode, vDGVItemCostData, vFrm, vPONumber, vTblCustom, vTblItemCustomInOut)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
                MsgBox("Error on aop")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
            DTAItem.Dispose()
            TBLItem.Dispose()
            ObjGRN = Nothing
        End Try
    End Function

    Public Overridable Function SimpanPerubahan(ByVal dgvItem As DataGridView, ByVal vdgvItemDP As DataGridView, _
          ByVal vDGVItemSR As DataGridView, ByVal vDGVItemExchangeRate As DataGridView, ByVal vUserCode As String, _
          ByVal vDGVItemCostData As DataGridView, ByVal vFrm As FrmTransaksiAOP2) As Boolean
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
            Me.vPrimary_Antrian, Me.mprimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiAOP_SimpanPerubahan, vUserCode)


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
                            retval = Me.SimpanPerubahan1(dgvItem, vdgvItemDP, vDGVItemSR, vDGVItemExchangeRate, vUserCode, vDGVItemCostData, vFrm)
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

    Private Sub CekNomor(ByVal DTA As DseTransaksiAOPTableAdapters.tbltrans_aopTableAdapter)
        Dim TanggalDari, TanggalSampai As Date
        Dim Urutan As Integer = Me.urutan
        Dim LoopGo As Boolean = True
        Dim JmlTrans As Integer
        Dim ObjSupp As New Core_Module.clsMasterSupplier
        Dim Foreign, LoopCount As Integer
        Dim ObjCD As New Core_Module.ClsCekData
        Dim PrimaryMax As Double = -1

        Try
            If ObjSupp.ApakahLokal(Me.mPrimarySupplier) = True Then
                Foreign = 0
            Else
                Foreign = 1
            End If
            TanggalDari = ObjCD.GetFisrtDateOfTheYear(Me.mTanggalTransaksi)
            If TanggalDari.Date < #5/1/2010# Then
                TanggalDari = #5/1/2010#
            End If
            TanggalSampai = ObjCD.GetLastDateOfTheYear(Me.mTanggalTransaksi)

            LoopCount = 0
            While LoopGo = True
                If DTA.ScalarQuery_GetJmlByUrutan(TanggalSampai, TanggalDari, Me.urutan, Foreign).HasValue = True Then
                    JmlTrans = DTA.ScalarQuery_GetJmlByUrutan(TanggalSampai, TanggalDari, Me.urutan, Foreign).Value
                Else
                    JmlTrans = 1
                End If
                If JmlTrans <= 1 Then
                    LoopGo = False
                Else
                    If DTA.ScalarQuery_GetPrimaryMaxByUrutan(Me.mNoUrut, TanggalSampai, TanggalDari, Foreign).HasValue = True Then
                        PrimaryMax = DTA.ScalarQuery_GetPrimaryMaxByUrutan(Me.mNoUrut, TanggalSampai, TanggalDari, Foreign)
                    Else
                        PrimaryMax = -1
                    End If
                    If PrimaryMax = Me.mprimary Then
                        Me.mnotransaksi = Me.GetNoBaru(Me.mTanggalTransaksi, Me.mPrimarySupplier)
                        Me.mNoUrut = CInt(Me.mnotransaksi.Substring(0, 6))
                    End If
                End If
                LoopCount += 1
            End While
            If LoopCount > 1 Then
                DTA.UpdateQuery_NoTrans_NoUrut(Me.mnotransaksi, Me.mNoUrut, Me.mprimary)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjSupp = Nothing
            ObjCD = Nothing
        End Try
    End Sub


    'Private Sub TransferJurnal_DP1(ByVal vPrimaryAOP As Double, ByVal vPrimarySupplier As Double, ByVal vPurchase As Decimal, ByVal vVAT As Decimal, _
    'ByVal vDutystamp As Decimal, ByVal vUrutanAOP As Double, ByVal vTanggalAOP As Date, ByVal vAOPNumber As String, _
    'ByVal vUserCode As String, ByVal vPaymentTerm As Double, _
    'ByVal vSupplierInvoiceNumber As String, ByVal vPurchaseOriginalAmount As Decimal, _
    'ByVal vExchangeRate As Decimal, ByVal vPrimaryCurrency As Double, ByVal vPrimaryAccount As Double, ByVal vPrimaryCurrTax As Double, _
    'ByVal vVatOriginal As Decimal, ByVal vDutyStampOriginal As Decimal)
    '    Dim DTA_Balance As New DseTransaksiAOP_TransfJurnalTableAdapters.tbltrans_balance_purchaseTableAdapter
    '    Dim TipeBalance As EnumTipePurchaseBalance = EnumTipePurchaseBalance.Automatic
    '    Dim DTA_Jurnal As New DseTransaksiAOP_TransfJurnalTableAdapters.tbltrans_jurnalTableAdapter
    '    Dim DTA_JurnalItem As New DseTransaksiAOP_TransfJurnalTableAdapters.tbltrans_jurnal_itemTableAdapter
    '    Dim TipeJurnal As EnumTipeJournal = EnumTipeJournal.PurchaseDownPaymentJournal
    '    Dim RefNumber As String = CStr(vUrutanAOP)
    '    Dim PrimaryJurnal As Double
    '    Dim Debet, Credit As Decimal
    '    Dim ObjSupplier As New Core_Module.clsMasterSupplier
    '    Dim ObjKartu As New ClsTransaksi_Kartu_APAR
    '    Dim SuppCode, SuppName As String
    '    Dim Local As Boolean = False
    '    Dim LocalStr As String = ""
    '    Dim MonthStr As String = ""
    '    Dim PrimaryAccount As Double = -1
    '    Dim ObjSetting As New Core_Module.ClsTransaksiSetting
    '    Dim ObjCur As New Core_Module.clsMasterCurrency
    '    Dim DefPrec, DefPrecOriginal As Integer
    '    Try
    '        DefPrec = ObjCur.GetDecimal_DefaultCurrency
    '        DefPrecOriginal = ObjCur.GetDefaultPrecission(vPrimaryCurrency)

    '        LocalStr = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_AOP_DP)
    '        If LocalStr = "" Then
    '            MsgBox("No AOP DP Account found. Please contact your vendor if you see this message", MsgBoxStyle.Information, "Information")
    '            Exit Sub
    '        End If

    '        'hapuskan data jurnal lama
    '        DTA_Jurnal.Delete(TipeJurnal, vPrimaryAOP)
    '        'hapuskan data balance sales lama
    '        DTA_Balance.Delete(vPrimaryAOP, TipeBalance)
    '        Local = ObjSupplier.ApakahLokal(vPrimarySupplier)
    '        If Local = True Then
    '            LocalStr = "L"
    '        Else
    '            LocalStr = "I"
    '        End If
    '        'hapus data kartu AP dan AR
    '        ObjKartu.Hapus(vPrimaryAOP, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis)

    '        If Me.ApakahTransferOtomatis(vTanggalAOP) = True Then
    '            SuppCode = ObjSupplier.GetSupplierCode(vPrimarySupplier)
    '            SuppName = ObjSupplier.GetSupplierName(vPrimarySupplier)
    '            'buat jurnal baru
    '            If RefNumber.Length < 6 Then
    '                While RefNumber.Length < 6
    '                    RefNumber = "0" + RefNumber
    '                End While
    '            End If
    '            MonthStr = vTanggalAOP.Month.ToString
    '            While MonthStr.Length < 2
    '                MonthStr = "0" + MonthStr
    '            End While
    '            'data header
    '            DTA_Jurnal.Insert("PJDP-" + RefNumber + "/" + LocalStr + "/" + MonthStr + "/" + vTanggalAOP.Year.ToString, vTanggalAOP, _
    '            vUrutanAOP, TipeJurnal, vPrimaryAOP, vPrimarySupplier, "AOP: " + vAOPNumber _
    '            + " - " + SuppCode + " - " + SuppName, vUserCode, _
    '            Now, vUserCode, Now, Now)

    '            'data item
    '            PrimaryJurnal = DTA_Jurnal.ScalarQuery_GetPrimaryJurnal(TipeJurnal, vPrimaryAOP).Value
    '            'biaya dibayar dimuka
    '            PrimaryAccount = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_AOP_DP)
    '            Debet = vPurchase
    '            Credit = 0
    '            DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAccount, Debet, Credit, "", Debet, Credit, 1)

    '            'PPN Masukan
    '            If vVAT > 0 Then
    '                PrimaryAccount = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_VAT_In)
    '                Debet = vVAT
    '                Credit = 0
    '                DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAccount, Debet, Credit, "", Debet, Credit, 1)
    '            End If

    '            'Stamp Duty Expense
    '            If vDutystamp > 0 Then
    '                PrimaryAccount = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_Dutystamp_Expense)
    '                Debet = vDutystamp
    '                Credit = 0
    '                DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAccount, Debet, Credit, "", Debet, Credit, 1)
    '            End If

    '            If vPrimaryCurrency <> Core_Module.PrimaryBaseCurrency Then
    '                'Hutang Non-IDR
    '                PrimaryAccount = ObjCur.GetPrimaryAkunAP(vPrimaryCurrency)
    '                Debet = 0
    '                Credit = 0
    '                DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAccount, Debet, vPurchase, "", Debet, vPurchaseOriginalAmount, vExchangeRate)
    '                'purchase amount
    '                DTA_Balance.InsertQuery(vPrimarySupplier, vPrimaryAOP, TipeBalance, "", #1/1/1950#, "", vTanggalAOP + TimeSpan.FromDays(vPaymentTerm), _
    '                vPurchase, 0, 0, vPurchase, 0, vUserCode, Now, vUserCode, Now, Now, vPrimaryCurrency, vPurchaseOriginalAmount, vExchangeRate, _
    '                PrimaryAccount)
    '                'buat kartu ar purchase amount
    '                ObjKartu.Simpan(vPrimaryAOP, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, vAOPNumber, _
    '                vPrimarySupplier, 0, vPurchase, "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
    '                vTanggalAOP, vPrimaryAOP, vTanggalAOP, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, _
    '                vPrimaryCurrency, 0, vPurchaseOriginalAmount, PrimaryAccount, vExchangeRate)

    '                'Hutang IDR
    '                If vDutystamp + vVAT <> 0 Then
    '                    'jurnal item
    '                    PrimaryAccount = ObjCur.GetPrimaryAkunAP(Core_Module.PrimaryBaseCurrency)
    '                    Debet = 0
    '                    Credit = vDutystamp + vVAT
    '                    DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAccount, Debet, Credit, "", Debet, Credit, 1)
    '                    'buat kartu ar tax amount + duty stamp
    '                    ObjKartu.Simpan(vPrimaryAOP, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, vAOPNumber, _
    '                    vPrimarySupplier, 0, vVAT + vDutystamp, "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
    '                    vTanggalAOP, vPrimaryAOP, vTanggalAOP, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, _
    '                    Core_Module.PrimaryBaseCurrency, 0, vVAT + vDutystamp, PrimaryAccount, 1)
    '                    'tax amount + duty stamp
    '                    DTA_Balance.InsertQuery(vPrimarySupplier, vPrimaryAOP, TipeBalance, "", #1/1/1950#, "", vTanggalAOP + TimeSpan.FromDays(vPaymentTerm), _
    '                    0, vVAT, vDutystamp, vVAT + vDutystamp, 0, vUserCode, Now, vUserCode, Now, Now, Core_Module.PrimaryBaseCurrency, vVAT + vDutystamp, 1, _
    '                    PrimaryAccount)
    '                End If
    '            Else
    '                'hutang IDR
    '                PrimaryAccount = ObjCur.GetPrimaryAkunAP(vPrimaryCurrency)
    '                Debet = 0
    '                Credit = vDutystamp + vVAT + vPurchase
    '                DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAccount, Debet, Credit, "", Debet, Credit, 1)

    '                'balance purchase
    '                DTA_Balance.InsertQuery(vPrimarySupplier, vPrimaryAOP, TipeBalance, "", #1/1/1950#, "", vTanggalAOP + TimeSpan.FromDays(vPaymentTerm), _
    '                vPurchase, vVAT, vDutystamp, vPurchase + vVAT + vDutystamp, 0, vUserCode, Now, vUserCode, Now, Now, Core_Module.PrimaryBaseCurrency, _
    '                vPurchase + vVAT + vDutystamp, 1, PrimaryAccount)

    '                'buat kartu ap baru
    '                ObjKartu.Simpan(vPrimaryAOP, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, vAOPNumber, _
    '                vPrimarySupplier, 0, vPurchase + vVAT + vDutystamp, "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
    '                vTanggalAOP, vPrimaryAOP, vTanggalAOP, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, _
    '                Core_Module.PrimaryBaseCurrency, 0, vPurchase + vVAT + vDutystamp, PrimaryAccount, 1)
    '            End If
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTA_Balance.Dispose()
    '        DTA_Jurnal.Dispose()
    '        DTA_JurnalItem.Dispose()
    '        ObjSupplier = Nothing
    '        ObjKartu = Nothing
    '        ObjSetting = Nothing
    '        ObjCur = Nothing
    '    End Try
    'End Sub


    Private Function GetCostIDR(ByVal vPrimaryAkunCost As Double, ByVal vDGV As DataGridView, ByVal vFrm As FrmTransaksiAOP2) As Decimal
        Dim row As DataGridViewRow
        Dim Retval As Decimal = 0

        Try
            For Each row In vDGV.Rows
                If CDbl(row.Cells(vFrm.ColPrimaryAccount.Index).Value) = vPrimaryAkunCost Then
                    Retval += row.Cells(vFrm.ColCostAmount.Index).Value
                End If
            Next

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetCostOriginal(ByVal vPrimaryAkunCost As Double, ByVal vDGV As DataGridView, ByVal vFrm As FrmTransaksiAOP2) As Decimal
        Dim row As DataGridViewRow
        Dim Retval As Decimal = 0

        Try
            For Each row In vDGV.Rows
                If CDbl(row.Cells(vFrm.ColPrimaryAccount.Index).Value) = vPrimaryAkunCost Then
                    Retval += row.Cells(vFrm.ColCostNominal.Index).Value
                End If
            Next

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetCostIDRQuick(ByVal vPrimaryAkunCost As Double, ByVal vDGV As DataGridView, ByVal vFrm As FrmTransaksiAOP_Quick2) As Decimal
        Dim row As DataGridViewRow
        Dim Retval As Decimal = 0

        Try
            For Each row In vDGV.Rows
                If CDbl(row.Cells(vFrm.ColPrimaryAccount.Index).Value) = vPrimaryAkunCost Then
                    Retval += row.Cells(vFrm.ColCostAmount.Index).Value
                End If
            Next

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetCostOriginalQuick(ByVal vPrimaryAkunCost As Double, ByVal vDGV As DataGridView, ByVal vFrm As FrmTransaksiAOP_Quick2) As Decimal
        Dim row As DataGridViewRow
        Dim Retval As Decimal = 0

        Try
            For Each row In vDGV.Rows
                If CDbl(row.Cells(vFrm.ColPrimaryAccount.Index).Value) = vPrimaryAkunCost Then
                    Retval += row.Cells(vFrm.ColCostNominal.Index).Value
                End If
            Next

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetInventoryPurchased(ByVal vTBLInv As DseTransaksiAOP_TransfJurnal.TBLInventoryDataTable, ByVal vPrimaryAkunInv As Double, ByVal vPurchaseAmount As Decimal) As Decimal
        Dim row As DseTransaksiAOP_TransfJurnal.TBLInventoryRow
        Dim retVal As Decimal = 0
        Dim objcd As New Core_Module.ClsCekData

        Dim objCurr As New Core_Module.clsMasterCurrency
        Try

            Dim decPurchaseAmount As Decimal = vPurchaseAmount
            For intP As Integer = 0 To vTBLInv.Rows.Count - 1
                row = vTBLInv.Rows(intP)
                If intP <> vTBLInv.Rows.Count - 1 Then
                    decPurchaseAmount -= objcd.DigitGrouping(row.At_Amount, objCurr.GetDecimal_DefaultCurrency)
                Else
                    row.At_Amount = objcd.DigitGrouping(decPurchaseAmount, objCurr.GetDecimal_DefaultCurrency)
                End If
            Next


            For Each row In vTBLInv.Rows
                If row.Primary_Inventory_Account = vPrimaryAkunInv Then
                    retVal += objcd.DigitGrouping(row.At_Amount, objCurr.GetDecimal_DefaultCurrency)
                End If
            Next

            Return retVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetPrimaryDutyStamp() As Double
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim Dummy As String = ""

        Try
            Dummy = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_Dutystamp_Expense)

            If Dummy = "" Then
                Return -1
            Else
                Return Dummy
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            ObjSetting = Nothing
        End Try
    End Function

    Private Function GetPrimaryAkunVATIn() As Double
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim Dummy As String = ""

        Try
            Dummy = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_VAT_In)
            If Dummy = "" Then
                Return -1
            Else
                Return Dummy
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            ObjSetting = Nothing
        End Try
    End Function

    Public Function GetPrimaryCurrency(ByVal vPrimaryAOP As Double) As Double
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aopTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetPrimaryCurrency(vPrimaryAOP).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryCurrency(vPrimaryAOP).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Function GetPrimaryAkunAP(ByVal vPrimaryAOP As Double) As Double
        Dim ObjCur As New Core_Module.clsMasterCurrency

        Try
            Return ObjCur.GetPrimaryAkunAP(Me.GetPrimaryCurrency(vPrimaryAOP))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCur = Nothing
        End Try
    End Function

    Private Function ApakahTransferOtomatis(ByVal vTanggalInvoice As Date) As Boolean
        Dim DTA As New DseTransaksiAOP_TransfJurnalTableAdapters.tbltrans_accounting_closingTableAdapter

        Try
            'Transfer otomatis jika:
            '1. Tanggal invoice >= convertion date
            If DTA.ScalarQuery_GetConvertionDate.HasValue = True Then
                If DTA.ScalarQuery_GetConvertionDate.Value.Date <= vTanggalInvoice.Date Then
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

    Private Sub Hapus_Jurnal_DP(ByVal vPrimaryAOP As Double)
        Dim DTA_Balance As New DseTransaksiAOP_TransfJurnalTableAdapters.tbltrans_balance_purchaseTableAdapter
        Dim DTA_Jurnal As New DseTransaksiAOP_TransfJurnalTableAdapters.tbltrans_jurnalTableAdapter
        Dim TipeBalance As EnumTipePurchaseBalance = EnumTipePurchaseBalance.Automatic
        Dim TipeJurnal As EnumTipeJournal = EnumTipeJournal.PurchaseDownPaymentJournal
        Dim Obj As New ClsTransaksi_Kartu_APAR

        Try
            'hapuskan data jurnal lama
            DTA_Jurnal.Delete(TipeJurnal, vPrimaryAOP)
            'hapuskan data balance sales lama
            DTA_Balance.Delete(vPrimaryAOP, TipeBalance)
            'hapuskan data kartu AP
            Obj.Hapus(vPrimaryAOP, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_Balance.Dispose()
            DTA_Jurnal.Dispose()
            Obj = Nothing
        End Try
    End Sub

    Private Sub Hapus_Jurnal(ByVal vPrimaryAOP As Double)
        Dim DTA_Balance As New DseTransaksiAOP_TransfJurnalTableAdapters.tbltrans_balance_purchaseTableAdapter
        Dim DTA_Jurnal As New DseTransaksiAOP_TransfJurnalTableAdapters.tbltrans_jurnalTableAdapter
        Dim TipeBalance As EnumTipePurchaseBalance = EnumTipePurchaseBalance.Automatic
        Dim TipeJurnal As EnumTipeJournal = EnumTipeJournal.PurchaseJournal
        Dim Obj As New ClsTransaksi_Kartu_APAR

        Try
            'hapuskan data jurnal lama
            DTA_Jurnal.Delete(TipeJurnal, vPrimaryAOP)
            'hapuskan data balance sales lama
            DTA_Balance.Delete(vPrimaryAOP, TipeBalance)
            'hapuskan data kartu AP
            Obj.Hapus(vPrimaryAOP, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_Balance.Dispose()
            DTA_Jurnal.Dispose()
            Obj = Nothing
        End Try
    End Sub

    Public Function GetConvertionDate() As Date
        Dim DTA As New DseTransaksiAOP_TransfJurnalTableAdapters.tbltrans_accounting_closingTableAdapter

        Try
            If DTA.ScalarQuery_GetConvertionDate.HasValue = True Then
                Return DTA.ScalarQuery_GetConvertionDate.Value
            Else
                Return Date.MaxValue
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return Date.MaxValue
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Function SimpanPO_Default(ByVal vDGVItem As DataGridView, ByVal vFrm As FrmTransaksiAOP_Quick2, _
   ByVal vUserCode As String, ByVal vSupplierAddIdentifier As String, ByVal ObjPO As ClsTransaksiPO, ByVal vPrimaryUser As Double, _
    ByVal vPONumber As String) As Double
        Dim ij As New FrmTransaksiPO
        Dim rowPO, rowAOP As DataGridViewRow
        Dim Lanjut As Boolean = False
        Dim AutoNumberPO As Boolean = False
        Dim ObjNumbering As New Core_Module.ClsMasterTransactionNumbering
        Dim ObjSupplier As New Core_Module.clsMasterSupplier

        Try
            If ObjSupplier.ApakahLokal(Me.mPrimarySupplier) = True Then
                AutoNumberPO = ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Impor)
            Else
                AutoNumberPO = ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Lokal)
            End If

            rowAOP = vDGVItem.Rows(0)
            'salin data header
            With ObjPO
                .primary = -1
                .AcknowledgeBy = ""
                .AcknowledgeBy_Jabatan = ""
                .ApprovedBy = ""
                .ApprovedBy_Jabatan = ""
                .Authorizationlevel = 3
                .AuthorizationNeeded = 3
                .IdentifierSuppAddress = vSupplierAddIdentifier
                If AutoNumberPO = True Then
                    .NoTransaksi = ObjPO.GetNoBaru1(Me.mTanggalTransaksi, Me.mPrimarySupplier)
                Else
                    .NoTransaksi = vPONumber
                End If
                .PaymentDue = 0
                .PaymentTerm = ""
                .PreparedBy = ""
                .PreparedBy_Jabatan = ""
                .PrimaryCurrency = rowAOP.Cells(vFrm.ColPrimaryCurrency.Index).Value
                .PrimaryDirectorStaffAuthorization = -1
                .PrimaryManagerStaffAuthorization = -1
                .PrimaryPurchaseStaffAuthorization = -1
                .primarySupplier = Me.mPrimarySupplier
                .PurchaseDirectorAuthorizationDate = Now
                .PurchaseManagerAuthorizationDate = Now
                .PurchaseStaffAuthorizationDate = Now
                .ReferencePaymentDue = 0
                .Remarks = ""
                .ShipVia = ""
                .SubTotal = Me.mSubTotal
                .SuppContactPerson = ""
                .SuppPhoneNo = ""
                .TanggalTransaksi = Me.mTanggalTransaksi
                .Tax = Me.mtax
                .TaxAmount = Me.mTaxAmount
                If Me.mTaxExclude = EnumTaxExclude.Exclude Then
                    .TaxExclude = True
                Else
                    .TaxExclude = False
                End If
                .Total = Me.mTotal
                .urutan = 0
            End With

            'salin data item
            For Each rowAOP In vDGVItem.Rows
                If rowAOP.Cells(vFrm.ColReferedToPO.Index).Value = 0 Then
                    ij.dgvItem.Rows.Add()
                    rowPO = ij.dgvItem.Rows(ij.dgvItem.Rows.Count - 1)

                    With rowPO
                        .Cells(ij.ColNo.Index).Value = ""
                        .Cells(ij.ColCodeNo.Index).Value = ""
                        .Cells(ij.ColDescription.Index).Value = ""
                        .Cells(ij.ColIPONo.Index).Value = ""
                        .Cells(ij.ColQty.Index).Value = rowAOP.Cells(vFrm.ColTotalQty.Index).Value
                        .Cells(ij.ColUnit.Index).Value = ""
                        .Cells(ij.ColUnitPrice.Index).Value = rowAOP.Cells(vFrm.ColUnitPrice.Index).Value
                        .Cells(ij.ColDiscount.Index).Value = rowAOP.Cells(vFrm.ColDiscount.Index).Value
                        .Cells(ij.ColNettPrice.Index).Value = rowAOP.Cells(vFrm.ColNettPrice.Index).Value
                        .Cells(ij.ColAmount.Index).Value = rowAOP.Cells(vFrm.ColAmount.Index).Value
                        .Cells(ij.ColDeliveryCont.Index).Value = False
                        .Cells(ij.ColDeliveryDate.Index).Value = #1/1/1950#
                        .Cells(ij.ColRemark.Index).Value = ""
                        .Cells(ij.ColPrimaryCodeNo.Index).Value = rowAOP.Cells(vFrm.ColPrimaryCodeNo.Index).Value
                        .Cells(ij.ColPrimaryUnits.Index).Value = rowAOP.Cells(vFrm.ColPrimaryUnit.Index).Value
                        .Cells(ij.ColPrimaryIPO.Index).Value = -1
                        .Cells(ij.colUserSecondaryUnits.Index).Value = False
                        .Cells(ij.ColIdentifier.Index).Value = "-1" + (ij.dgvItem.RowCount - 1).ToString
                    End With
                End If
            Next

            'hitung ulang amount, karena tidak semua item po masuk ke dalam po quick puchase
            Me.Hitung_Total_PO(ObjPO, ij.dgvItem, ij.objCurrency.GetDefaultPrecission(ObjPO.PrimaryCurrency), ij.ObjCD, ij)

            'simpan data po
            ObjPO.SimpanBaru(ij.dgvItem, vUserCode, ij, ij.vTBLItemDelivery, False)

            'authorize po
            ObjPO.Authorize(ObjPO.primary, vPrimaryUser, vUserCode)

            Return ObjPO.primary
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij = Nothing
            ObjNumbering = Nothing
            ObjSupplier = Nothing
        End Try
    End Function

    ''@rosadinnn
    ''@MUaT
    Private Function SimpanPO_Multi(ByVal vDGVItem As DataGridView, ByVal vFrm As FrmTransaksiAOP_Quick2, _
            ByVal vUserCode As String, ByVal vSupplierAddIdentifier As String, ByVal ObjPO As ClsTransaksiPO, ByVal vPrimaryUser As Double, _
            ByVal vPONumber As String, Optional ByVal vCekDataBenar As Boolean = True, _
            Optional ByVal vFromCutOff As Boolean = False) As Double
        Dim ij As New FrmTransaksiPO
        Dim rowPO, rowAOP As DataGridViewRow
        Dim Lanjut As Boolean = False
        Dim AutoNumberPO As Boolean = False
        Dim ObjNumbering As New Core_Module.ClsMasterTransactionNumbering
        Dim ObjSupplier As New Core_Module.clsMasterSupplier

        Try
            If vCekDataBenar = True Then
                If ObjSupplier.ApakahLokal(Me.mPrimarySupplier) = True Then
                    AutoNumberPO = ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Impor)
                Else
                    AutoNumberPO = ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Lokal)
                End If
            End If

            rowAOP = vDGVItem.Rows(0)
            'salin data header
            With ObjPO
                .primary = -1
                .AcknowledgeBy = ""
                .AcknowledgeBy_Jabatan = ""
                .ApprovedBy = ""
                .ApprovedBy_Jabatan = ""
                .Authorizationlevel = 3
                .AuthorizationNeeded = 3
                .IdentifierSuppAddress = vSupplierAddIdentifier
                If AutoNumberPO = True Then
                    .NoTransaksi = ObjPO.GetNoBaru1(Me.mTanggalTransaksi, Me.mPrimarySupplier)
                Else
                    .NoTransaksi = vPONumber
                End If
                .PaymentDue = 0
                .PaymentTerm = ""
                .PreparedBy = ""
                .PreparedBy_Jabatan = ""
                .PrimaryCurrency = rowAOP.Cells(vFrm.ColPrimaryCurrency.Index).Value
                .PrimaryDirectorStaffAuthorization = -1
                .PrimaryManagerStaffAuthorization = -1
                .PrimaryPurchaseStaffAuthorization = -1
                .primarySupplier = Me.mPrimarySupplier
                .PurchaseDirectorAuthorizationDate = Now
                .PurchaseManagerAuthorizationDate = Now
                .PurchaseStaffAuthorizationDate = Now
                .ReferencePaymentDue = 0
                .Remarks = ""
                .ShipVia = ""
                .SubTotal = Me.mSubTotal
                .SuppContactPerson = ""
                .SuppPhoneNo = ""
                .TanggalTransaksi = Me.mTanggalTransaksi
                .Tax = Me.mtax
                .TaxAmount = Me.mTaxAmount
                If Me.mTaxExclude = EnumTaxExclude.Exclude Then
                    .TaxExclude = True
                Else
                    .TaxExclude = False
                End If
                .Total = Me.mTotal
                .urutan = 0
                If vFromCutOff = True Then
                    .CreatedBy = Me.mCreatedBy
                    .CreatedDate = Me.mCreatedDate
                    .ModifiedBy = Me.mModifiedBy
                    .ModifiedDate = Me.mModifiedDate
                    .TimeStamp = Me.mTimestamp
                End If
            End With

            'salin data item
            For Each rowAOP In vDGVItem.Rows
                If rowAOP.Cells(vFrm.ColReferedToPO.Index).Value = 0 Then
                    ij.dgvItem.Rows.Add()
                    rowPO = ij.dgvItem.Rows(ij.dgvItem.Rows.Count - 1)

                    With rowPO
                        .Cells(ij.ColNo.Index).Value = ""
                        .Cells(ij.ColCodeNo.Index).Value = ""
                        .Cells(ij.ColDescription.Index).Value = ""
                        .Cells(ij.ColIPONo.Index).Value = ""
                        .Cells(ij.ColQty.Index).Value = rowAOP.Cells(vFrm.ColTotalQty.Index).Value
                        .Cells(ij.ColUnit.Index).Value = ""
                        .Cells(ij.ColUnitPrice.Index).Value = rowAOP.Cells(vFrm.ColUnitPrice.Index).Value
                        .Cells(ij.ColDiscount.Index).Value = rowAOP.Cells(vFrm.ColDiscount.Index).Value
                        .Cells(ij.ColNettPrice.Index).Value = rowAOP.Cells(vFrm.ColNettPrice.Index).Value
                        .Cells(ij.ColAmount.Index).Value = rowAOP.Cells(vFrm.ColAmount.Index).Value
                        .Cells(ij.ColDeliveryCont.Index).Value = False
                        .Cells(ij.ColDeliveryDate.Index).Value = #1/1/1950#
                        .Cells(ij.ColRemark.Index).Value = ""
                        .Cells(ij.ColPrimaryCodeNo.Index).Value = rowAOP.Cells(vFrm.ColPrimaryCodeNo.Index).Value
                        .Cells(ij.ColPrimaryUnits.Index).Value = rowAOP.Cells(vFrm.ColPrimaryUnit.Index).Value
                        .Cells(ij.ColPrimaryIPO.Index).Value = -1
                        .Cells(ij.colUserSecondaryUnits.Index).Value = False
                        .Cells(ij.ColIdentifier.Index).Value = "-1" + (ij.dgvItem.RowCount - 1).ToString

                        .Cells(ij.ColQtyMulti.Index).Value = rowAOP.Cells(vFrm.ColTotalQtyMulti.Index).Value
                        .Cells(ij.ColPrimaryUnitMulti.Index).Value = rowAOP.Cells(vFrm.ColPrimaryUnitsMulti.Index).Value
                        .Cells(ij.ColRatioReference.Index).Value = rowAOP.Cells(vFrm.ColRatioReference.Index).Value
                        .Cells(ij.ColRatioDefault.Index).Value = rowAOP.Cells(vFrm.ColRatioDefault.Index).Value
                        .Cells(ij.ColPriceByMultiUnit.Index).Value = rowAOP.Cells(vFrm.ColPriceByMultiUnit.Index).Value
                    End With
                End If
            Next

            'hitung ulang amount, karena tidak semua item po masuk ke dalam po quick puchase
            Me.Hitung_Total_PO(ObjPO, ij.dgvItem, ij.objCurrency.GetDefaultPrecission(ObjPO.PrimaryCurrency), ij.ObjCD, ij)

            'simpan data po
            ObjPO.SimpanBaru1(ij.dgvItem, vUserCode, ij, ij.vTBLItemDelivery, False, vCekDataBenar, vFromCutOff)

            'authorize po
            ObjPO.Authorize(ObjPO.primary, vPrimaryUser, vUserCode)

            Return ObjPO.primary
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            MsgBox("Error on po")
        Finally
            ij = Nothing
            ObjNumbering = Nothing
            ObjSupplier = Nothing
        End Try
    End Function

    Private Sub Hitung_Total_PO(ByVal vObjPO As ClsTransaksiPO, ByVal vDGV_PO As DataGridView, _
    ByVal DefaultPrecision As Integer, ByVal ObjCD As Core_Module.ClsCekData, ByVal vFrm As FrmTransaksiPO)
        Dim row As DataGridViewRow
        Dim SubTotal As Decimal = 0

        Try
            'hitung sub total
            For Each row In vDGV_PO.Rows
                SubTotal += CDec(row.Cells(vFrm.ColAmount.Index).Value)
            Next
            vObjPO.SubTotal = ObjCD.DigitGrouping(SubTotal, DefaultPrecision)
            'hitung pajak
            If vObjPO.TaxExclude = True Then
                vObjPO.TaxAmount = ObjCD.DigitGrouping(vObjPO.Tax * SubTotal / 100, _
                DefaultPrecision)
            Else
                vObjPO.Tax = 0
                vObjPO.TaxAmount = 0
            End If
            'hitung total = sub total + tax
            vObjPO.Total = vObjPO.SubTotal + vObjPO.TaxAmount
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function SimpanStockIn_Default(ByVal vUserCode As String, ByVal vdgvItem As DataGridView, ByVal vFrm As FrmTransaksiAOP_Quick2, _
   ByVal PrimaryPO As Double, ByVal AddIdentifier As String, ByVal ObjStock As Inventory_Module.clsTransaksiStockInOut, ByVal NoPO As String, Optional ByVal vNoTransaksi As String = "") As Double
        Dim ij As New FrmTransaksiStockInPurchase2
        Dim rowAOP, rowStock As DataGridViewRow
        Dim TblItemCustom As New CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable 'kemungkinan dikirim dari form quick nya

        Try
            'salin data header
            With ObjStock
                .ImporDariPackingList = False
                If vNoTransaksi = "" Then
                    .NoTransaksi = ObjStock.GetNoBaru(Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase, Me.TanggalTransaksi)
                Else
                    .NoTransaksi = vNoTransaksi
                End If
                .NoUrut = 0
                .Primary = -1
                .PrimaryCustomer = -1
                .PrimaryDOSales = -1
                .PrimaryGIT = -1
                .PrimaryPOPurchase = -1
                .PrimarySupplier = Me.mPrimarySupplier
                .Remarks = ""
                .SupplierDoNo = ""
                .TanggalTransaksi = Me.TanggalTransaksi
                .TipeTransaksi = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase

                'salin data item
                ij.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                ij.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                For Each rowAOP In vdgvItem.Rows
                    ij.dgvItem.Rows.Add()
                    rowStock = ij.dgvItem.Rows(ij.dgvItem.Rows.Count - 1)
                    With rowStock
                        .Cells(ij.ColNo.Index).Value = ""
                        .Cells(ij.ColPONo.Index).Value = ""
                        .Cells(ij.ColSuppDONo.Index).Value = ""
                        .Cells(ij.ColCodeNo.Index).Value = ""
                        .Cells(ij.ColDescription.Index).Value = ""
                        .Cells(ij.ColQty.Index).Value = rowAOP.Cells(vFrm.ColTotalQty.Index).Value
                        .Cells(ij.ColUnit.Index).Value = ""
                        .Cells(ij.ColSecondaryQty.Index).Value = -1
                        .Cells(ij.ColSecondaryUnit.Index).Value = ""
                        .Cells(ij.ColBatch.Index).Value = rowAOP.Cells(vFrm.ColLotNumber.Index).Value
                        .Cells(ij.ColExpiredDate.Index).Value = rowAOP.Cells(vFrm.ColExpiry.Index).Value
                        .Cells(ij.ColLocation.Index).Value = ""
                        .Cells(ij.ColWareHouse.Index).Value = ""
                        .Cells(ij.ColRemark.Index).Value = ""
                        If rowAOP.Cells(vFrm.ColReferedToPO.Index).Value = 1 Then
                            .Cells(ij.ColPrimaryPOno.Index).Value = rowAOP.Cells(vFrm.ColPrimaryPO.Index).Value
                        Else
                            .Cells(ij.ColPrimaryPOno.Index).Value = PrimaryPO
                        End If
                        .Cells(ij.ColPrimaryCodeNo.Index).Value = rowAOP.Cells(vFrm.ColPrimaryCodeNo.Index).Value
                        .Cells(ij.ColPrimaryLocation.Index).Value = rowAOP.Cells(vFrm.ColPrimaryLocation.Index).Value
                        .Cells(ij.ColPrimaryUnits.Index).Value = rowAOP.Cells(vFrm.ColPrimaryUnit.Index).Value
                        .Cells(ij.ColPrimarySecondaryUnits.Index).Value = -1
                        .Cells(ij.ColPrimaryPdO.Index).Value = -1
                    End With
                Next
            End With

            'simpan
            ObjStock.SimpanBaru1(ij.dgvItem, vUserCode, TblItemCustom, False, False)

            'retval
            Return ObjStock.Primary
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            ij = Nothing
            ObjStock = Nothing
        End Try
    End Function

    Private Function SimpanStockIn_Multi(ByVal vUserCode As String, ByVal vdgvItem As DataGridView, ByVal vFrm As FrmTransaksiAOP_Quick2, _
            ByVal PrimaryPO As Double, ByVal AddIdentifier As String, ByVal ObjStock As Inventory_Module.clsTransaksiStockInOut, _
            ByVal NoPO As String, ByVal TblItemCustom As CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable, _
            Optional ByVal vNoTransaksi As String = "", Optional ByVal vCekDataBenar As Boolean = True, _
            Optional ByVal vKoreksiStock As Boolean = True, Optional ByVal vFromCutOff As Boolean = False) As Double
        Dim ij As New FrmTransaksiStockInPurchase2
        Dim rowAOP, rowStock As DataGridViewRow
        'kemungkinan dikirim dari form quick nya

        Try
            'salin data header
            With ObjStock
                .ImporDariPackingList = False
                If vNoTransaksi = "" Then
                    .NoTransaksi = ObjStock.GetNoBaru(Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase, Me.TanggalTransaksi)
                Else
                    .NoTransaksi = vNoTransaksi
                End If
                .NoUrut = 0
                .Primary = -1
                .PrimaryCustomer = -1
                .PrimaryDOSales = -1
                .PrimaryGIT = -1
                .PrimaryPOPurchase = -1
                .PrimarySupplier = Me.mPrimarySupplier
                .Remarks = ""
                .SupplierDoNo = ""
                .TanggalTransaksi = Me.TanggalTransaksi
                .TipeTransaksi = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase
                If vFromCutOff = True Then
                    .CreatedBy = Me.mCreatedBy
                    .CreatedDate = Me.mCreatedDate
                    .ModifiedBy = Me.mModifiedBy
                    .ModifiedDate = Me.mModifiedDate
                    .TimeStamp = Me.mTimestamp
                End If

                'salin data item
                ij.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                ij.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                For Each rowAOP In vdgvItem.Rows
                    ij.dgvItem.Rows.Add()
                    rowStock = ij.dgvItem.Rows(ij.dgvItem.Rows.Count - 1)
                    With rowStock
                        .Cells(ij.ColNo.Index).Value = ""
                        .Cells(ij.ColPONo.Index).Value = ""
                        .Cells(ij.ColSuppDONo.Index).Value = ""
                        .Cells(ij.ColCodeNo.Index).Value = ""
                        .Cells(ij.ColDescription.Index).Value = ""
                        .Cells(ij.ColQty.Index).Value = rowAOP.Cells(vFrm.ColTotalQty.Index).Value
                        .Cells(ij.ColUnit.Index).Value = ""
                        .Cells(ij.ColSecondaryQty.Index).Value = -1
                        .Cells(ij.ColSecondaryUnit.Index).Value = ""
                        .Cells(ij.ColBatch.Index).Value = rowAOP.Cells(vFrm.ColLotNumber.Index).Value
                        .Cells(ij.ColExpiredDate.Index).Value = rowAOP.Cells(vFrm.ColExpiry.Index).Value
                        .Cells(ij.ColLocation.Index).Value = ""
                        .Cells(ij.ColWareHouse.Index).Value = ""
                        .Cells(ij.ColRemark.Index).Value = ""
                        If rowAOP.Cells(vFrm.ColReferedToPO.Index).Value = 1 Then
                            .Cells(ij.ColPrimaryPOno.Index).Value = rowAOP.Cells(vFrm.ColPrimaryPO.Index).Value
                        Else
                            .Cells(ij.ColPrimaryPOno.Index).Value = PrimaryPO
                        End If
                        .Cells(ij.ColPrimaryCodeNo.Index).Value = rowAOP.Cells(vFrm.ColPrimaryCodeNo.Index).Value
                        .Cells(ij.ColPrimaryLocation.Index).Value = rowAOP.Cells(vFrm.ColPrimaryLocation.Index).Value
                        .Cells(ij.ColPrimaryUnits.Index).Value = rowAOP.Cells(vFrm.ColPrimaryUnit.Index).Value
                        .Cells(ij.ColPrimarySecondaryUnits.Index).Value = -1
                        .Cells(ij.ColPrimaryPdO.Index).Value = -1

                        ''@rosadinnn
                        ''@MUaT
                        .Cells(ij.ColQtyMulti.Index).Value = rowAOP.Cells(vFrm.ColTotalQtyMulti.Index).Value
                        .Cells(ij.ColPrimaryUnitMulti.Index).Value = rowAOP.Cells(vFrm.ColPrimaryUnitsMulti.Index).Value
                        .Cells(ij.ColRatioDefault.Index).Value = rowAOP.Cells(vFrm.ColRatioDefault.Index).Value
                        .Cells(ij.ColRatioReference.Index).Value = rowAOP.Cells(vFrm.ColRatioReference.Index).Value

                    End With
                Next
            End With

            'simpan
            ObjStock.SimpanBaru1(ij.dgvItem, vUserCode, TblItemCustom, False, False, -1, vCekDataBenar, vKoreksiStock, vFromCutOff)

            'retval
            Return ObjStock.Primary
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            MsgBox("Error on stock in")
            Return -1
        Finally
            ij = Nothing
            ObjStock = Nothing
        End Try
    End Function

    Private Function ApakahTidakMinus(ByVal vPrimaryStockInPurchase As Double, ByVal vDGVItem As DataGridView, _
  ByVal vFrm As FrmTransaksiAOP_Quick2) As Boolean
        Dim DTA As New Inventory_Module.dseTransaksiStockInOutTableAdapters.transaksi_stockinout_item_viewTableAdapter
        Dim TBL As New Inventory_Module.dseTransaksiStockInOut.transaksi_stockinout_item_viewDataTable
        Dim rowInOutLama As Inventory_Module.dseTransaksiStockInOut.transaksi_stockinout_item_viewRow

        Dim DTAMasterProductQty As New Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter
        DTAMasterProductQty.ClearBeforeFill = False
        Dim TBLMasterProductQty As New Core_Module.dseMasterProduct.tblmaster_productqtyDataTable

        Dim rowDGV As DataGridViewRow
        Dim BatchNo As String = ""
        Dim PrimaryProduct, PrimaryLocation As Double
        Dim Sec_Qty, Qty As Decimal
        Dim ExpiryDate As Date

        Dim ListPrimaryProduct As New Queue(Of Double)
        Dim rowProduct As Core_Module.dseMasterProduct.tblmaster_productqtyRow
        Dim Dummy As String = ""

        Try
            'load database master product qty berdasarkan data product lama dan data product baru
            TBL.Rows.Clear()
            DTA.Fill(TBL, vPrimaryStockInPurchase)
            For Each rowInOutLama In TBL.Rows
                With rowInOutLama
                    DTAMasterProductQty.FillBy_Parameter(TBLMasterProductQty, _
                    .Secondary_Qty, .Primary_Product, .Batch_No, .Expiry_Date)
                End With
            Next

            'load data item baru
            For Each rowDGV In vDGVItem.Rows
                BatchNo = rowDGV.Cells(vFrm.ColLotNumber.Index).Value
                PrimaryLocation = rowDGV.Cells(vFrm.ColPrimaryLocation.Index).Value
                Sec_Qty = -1

                If IsDate(rowDGV.Cells(vFrm.ColExpiry.Index).Value) = False Then
                    ExpiryDate = #1/1/1950#
                Else
                    ExpiryDate = CDate(rowDGV.Cells(vFrm.ColExpiry.Index).Value)
                End If
                PrimaryProduct = rowDGV.Cells(vFrm.ColPrimaryCodeNo.Index).Value
                DTAMasterProductQty.FillBy_Parameter(TBLMasterProductQty, _
                Sec_Qty, PrimaryProduct, BatchNo, ExpiryDate)
            Next

            'simulasikan untuk stock in 
            'simulasi hapus data item lama
            If TBL.Rows.Count > 0 Then
                For Each rowInOutLama In TBL.Rows
                    With rowInOutLama
                        Me.Koreksi4HapusIn(TBLMasterProductQty, _
                        .Primary_Product, .Qty, .Secondary_Qty, _
                        .Primary_Storage_Area, .Batch_No, -1, .Expiry_Date)
                    End With
                Next
            End If
            'simulasi simpan item baru
            For Each rowDGV In vDGVItem.Rows
                BatchNo = rowDGV.Cells(vFrm.ColLotNumber.Index).Value
                PrimaryLocation = rowDGV.Cells(vFrm.ColPrimaryLocation.Index).Value
                Sec_Qty = -1
                If IsDate(rowDGV.Cells(vFrm.ColExpiry.Index).Value) = False Then
                    ExpiryDate = #1/1/1950#
                Else
                    ExpiryDate = CDate(rowDGV.Cells(vFrm.ColExpiry.Index).Value)
                End If
                PrimaryProduct = rowDGV.Cells(vFrm.ColPrimaryCodeNo.Index).Value
                Qty = rowDGV.Cells(vFrm.ColTotalQty.Index).Value

                Me.Koreksi4SimpanIn(TBLMasterProductQty, _
                PrimaryProduct, Qty, Sec_Qty, PrimaryLocation, _
                BatchNo, -1, ExpiryDate)
            Next

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

            'cek batasan minus stock inout custom
            '################################CUSTOM INOUT#############################################
            Dim ObjCustInOut As New CustomSMESS.ClsTransaksiStockInOut
            Dim ListProductCustom As New List(Of Double)
            Dim vTblItemCustom As New CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable
            Dim tblProdQty As New CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_productqtyDataTable
            If ObjCustInOut.ApakahDataBenar_CustomStockInOut(vPrimaryStockInPurchase, False, CustomSMESS.ClsTransaksiStockInOut.EnumJenisStockCard.Stock_In, _
            tblProdQty, vFrm.TblItemCustomInOut, ListProductCustom, CustomSMESS.ClsTransaksiStockInOut.EnumJenisTransaksiAsal.StockIn_Purchase) = False Then
                Return False
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
        End Try
    End Function

    Public Function Koreksi4HapusIn(ByVal TBLMasterProduct As Core_Module.dseMasterProduct.tblmaster_productqtyDataTable, _
ByVal vPrimaryProduct As Double, ByVal Qty As Decimal, Optional ByVal vSecondaryQty As Double = -1, _
Optional ByVal vPrimaryLocation As Double = -1, Optional ByVal _
vBatchNo As String = "", Optional ByVal vPrimaryPdO As Double = -1, _
Optional ByVal vExpiryDate As Date = #1/1/1950#) As Boolean
        Dim JmlLama As Decimal = 0
        Dim JmlBaru As Decimal = 0

        Try
            JmlLama = Me.GetJml_Simulasi(TBLMasterProduct, vPrimaryProduct, vSecondaryQty, _
            vPrimaryLocation, vBatchNo)

            If JmlLama <> 0 Then
                JmlBaru = JmlLama - Qty
                If JmlBaru = 0 Then
                    'hapus data lama
                    Me.Hapus_Simulasi(TBLMasterProduct, vPrimaryProduct, vSecondaryQty, vPrimaryLocation, _
                    vBatchNo)
                Else
                    'update data jumlah
                    Me.SetJml_Simulasi(TBLMasterProduct, vPrimaryProduct, _
                    vSecondaryQty, vPrimaryLocation, vBatchNo, JmlBaru)
                End If
            Else
                'masukkan data baru: tambah
                Me.Insert_Simulasi(TBLMasterProduct, vPrimaryProduct, vSecondaryQty, _
                vPrimaryLocation, vBatchNo, Qty * -1)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function Koreksi4SimpanIn(ByVal TBLMasterProduct As Core_Module.dseMasterProduct.tblmaster_productqtyDataTable, _
ByVal vPrimaryProduct As Double, _
ByVal Qty As Decimal, Optional ByVal vSecondaryQty As Decimal = -1, Optional ByVal _
vPrimaryLocation As Double = -1, Optional ByVal vBatchNo As String = "", _
Optional ByVal vPrimaryPdO As Double = -1, Optional ByVal vExpiryDate As Date = _
#1/1/1950#) As Boolean
        Dim JmlLama As Decimal = 0
        Dim JmlBaru As Decimal = 0
        Try
            JmlLama = Me.GetJml_Simulasi(TBLMasterProduct, vPrimaryProduct, vSecondaryQty, _
            vPrimaryLocation, vBatchNo)
            If JmlLama <> 0 Then
                JmlBaru = JmlLama + Qty
                If JmlBaru = 0 Then
                    'hapus data lama
                    Me.Hapus_Simulasi(TBLMasterProduct, vPrimaryProduct, vSecondaryQty, vPrimaryLocation, _
                    vBatchNo)
                Else
                    'update data jumlah
                    Me.SetJml_Simulasi(TBLMasterProduct, vPrimaryProduct, vSecondaryQty, _
                    vPrimaryLocation, vBatchNo, JmlBaru)
                End If
            Else
                'masukkan data baru: tambah
                Me.Insert_Simulasi(TBLMasterProduct, vPrimaryProduct, vSecondaryQty, _
                vPrimaryLocation, vBatchNo, Qty)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub Hapus_Simulasi(ByVal TBL As Core_Module.dseMasterProduct.tblmaster_productqtyDataTable, _
            ByVal vPrimaryProduct As Double, ByVal vSecQty As Decimal, _
            ByVal vPrimaryLocation As Double, ByVal vBatchNo As String)
        Dim row As Core_Module.dseMasterProduct.tblmaster_productqtyRow
        Dim filter As String = ""

        Try
            filter = "[Primary Product] = " + vPrimaryProduct.ToString + " and [Qty Secondary] = " + _
            vSecQty.ToString + " and [Lot No] = '" + vBatchNo + "' and [Primary Storage Area] = " + _
            vPrimaryLocation.ToString + ""
            TBL.DefaultView.RowFilter = filter
            If TBL.DefaultView.Count > 0 Then
                row = TBL.DefaultView(0).Row
                TBL.Rows.Remove(row)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetJml_Simulasi(ByVal TBL As Core_Module.dseMasterProduct.tblmaster_productqtyDataTable, _
    ByVal vPrimaryProduct As Double, ByVal vSecQty As Decimal, _
    ByVal vPrimaryLocation As Double, ByVal vBatchNo As String) As Decimal
        Dim row As Core_Module.dseMasterProduct.tblmaster_productqtyRow
        Dim filter As String = ""

        Try
            filter = "[Primary Product] = " + vPrimaryProduct.ToString + " and [Qty Secondary] = " + _
            vSecQty.ToString + " and [Lot No] = '" + vBatchNo + "' and [Primary Storage Area] = " _
            + vPrimaryLocation.ToString + ""
            TBL.DefaultView.RowFilter = filter
            If TBL.DefaultView.Count > 0 Then
                row = TBL.DefaultView(0).Row
                Return row.Qty
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub SetJml_Simulasi(ByVal TBL As Core_Module.dseMasterProduct.tblmaster_productqtyDataTable, _
    ByVal vPrimaryProduct As Double, ByVal vSecQty As Decimal, _
    ByVal vPrimaryLocation As Double, ByVal vBatchNo As String, _
    ByVal JmlBaru As Decimal)
        Dim row As Core_Module.dseMasterProduct.tblmaster_productqtyRow
        Dim filter As String = ""

        Try
            filter = "[Primary Product] = " + vPrimaryProduct.ToString + " and [Qty Secondary] = " + _
            vSecQty.ToString + " and [Lot No] = '" + vBatchNo + "' and [Primary Storage Area] = " _
            + vPrimaryLocation.ToString + ""
            TBL.DefaultView.RowFilter = filter
            If TBL.DefaultView.Count > 0 Then
                row = TBL.DefaultView(0).Row
                row.Qty = JmlBaru
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Insert_Simulasi(ByVal TBL As Core_Module.dseMasterProduct.tblmaster_productqtyDataTable, _
    ByVal vPrimaryProduct As Double, ByVal vSecQty As Decimal, _
    ByVal vPrimaryLocation As Double, ByVal vBatchNo As String, _
    ByVal JmlBaru As Decimal)
        Dim row As Core_Module.dseMasterProduct.tblmaster_productqtyRow

        Try
            row = TBL.Newtblmaster_productqtyRow

            row.Primary = Me.GetPrimaryMax_TBLQtySimulasi(TBL) + 1
            row.Primary_Product = vPrimaryProduct
            row.Primary_Storage_Area = vPrimaryLocation
            row.Qty_Secondary = vSecQty
            row.Lot_No = vBatchNo
            'expiry date diabaikan dalam simulasi karena tidak berpengaruh ke databse
            'yang sebenarnya
            row.Expiry_Date = #1/1/1950#
            row.Qty = JmlBaru

            TBL.Rows.Add(row)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetPrimaryMax_TBLQtySimulasi(ByVal TBL As Core_Module.dseMasterProduct.tblmaster_productqtyDataTable) As Double
        Dim Sort As String = "[Primary] DESC"
        Dim row As Core_Module.dseMasterProduct.tblmaster_productqtyRow

        Try
            TBL.DefaultView.RowFilter = ""
            TBL.DefaultView.Sort = Sort
            If TBL.DefaultView.Count > 0 Then
                row = TBL.DefaultView(0).Row
                Return row.Primary
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Sub New()
        Core_Module.SetConstring_DTAAntrian(Me.DTAInOUT)
    End Sub

    Public Function GetDPAmount_ExcludePrimaryAOP(ByVal vPrimaryPO As Double, ByVal vPrimaryAopExclude As Double) As Decimal
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aopTableAdapter
        Dim RetVal As Nullable(Of Decimal)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetDPAmountByPrimaryPO(vPrimaryAopExclude, vPrimaryPO)
            If RetVal.HasValue = True Then
                Return RetVal.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetDPAmount(ByVal vPrimaryPO As Double) As Decimal
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aopTableAdapter
        Dim RetVal As Nullable(Of Decimal)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetDPAmountByPrimaryPO2(vPrimaryPO)
            If RetVal.HasValue = True Then
                Return RetVal.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            DTA.Dispose()
        End Try
    End Function

    '\\TR = No Transaksi Admiral
    'Hanya berlaku jika format numbering untuk SO, DO, Out Sales
    Public Sub GetNumbering_PO_StockIn(ByVal vPrimaryAOP As Double, ByVal vAOPDate As Date, _
    ByRef vPOno As String, ByRef vOutStockInNo As String, ByVal vPrimarySupp As Double)
        Dim dta As New DseTransaksiAOP_GenerateNoTransaksiTableAdapters.DataTable1TableAdapter
        Dim Tbl As New DseTransaksiAOP_GenerateNoTransaksi.DataTable1DataTable
        Dim row As DseTransaksiAOP_GenerateNoTransaksi.DataTable1Row
        Dim objSupplier As New Core_Module.clsMasterSupplier
        Dim ObjPO As New ClsTransaksiPO
        Dim ObjInOut As New Inventory_Module.clsTransaksiStockInOut
        Dim Lokal As Boolean = True

        Try
            'vSOno = ""
            'vDOno = ""
            'vOutSalesNo = ""




            If objSupplier.ApakahLokal(vPrimarySupp) = False Then
                Lokal = False
            End If

            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.Fill_PrimaryAOP(Tbl, vPrimaryAOP) > 0 Then
                row = Tbl.Rows(0)
                With row
                    'PO NUMBERING
                    If Lokal = True Then
                        If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Lokal) = True Then
                            If Me.ObjNumbering.GetResetType(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Lokal) = Core_Module.ClsMasterTransactionNumbering.EnumReset.Monthly Then
                                'monthly
                                If Format(.PO_Date, "MMM yyyy") = Format(vAOPDate, "MMM yyyy") Then
                                    vPOno = .PO_No
                                End If
                            Else
                                'yearly
                                If Format(.PO_Date, "yyyy") = Format(vAOPDate, "yyyy") Then
                                    If Format(.PO_Date, "MMM") = Format(vAOPDate, "MMM") Then
                                        'jika sama tahun dan bulannya maka keep no lama
                                        vPOno = .PO_No
                                    Else
                                        'jika sama tahun tapi bulannya beda maka ganti bulannya saja dan nomor nya tetap di keep
                                        vPOno = Me.ObjNumbering.GetNomorTransaksiBaru1(vAOPDate, .PO_Urutan, Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Lokal)
                                    End If
                                Else
                                    'jika tahun dan bulannya tidak sama maka get no baru
                                    vPOno = ObjPO.GetNoBaru1(vAOPDate)
                                End If
                            End If
                        End If
                    Else
                        If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Impor) = True Then
                            If Me.ObjNumbering.GetResetType(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Impor) = Core_Module.ClsMasterTransactionNumbering.EnumReset.Monthly Then
                                'monthly
                                If Format(.PO_Date, "MMM yyyy") = Format(vAOPDate, "MMM yyyy") Then
                                    vPOno = .PO_No
                                End If
                            Else
                                'yearly
                                If Format(.PO_Date, "yyyy") = Format(vAOPDate, "yyyy") Then
                                    If Format(.PO_Date, "MMM") = Format(vAOPDate, "MMM") Then
                                        'jika sama tahun dan bulannya maka keep no lama
                                        vPOno = .PO_No
                                    Else
                                        'jika sama tahun tapi bulannya beda maka ganti bulannya saja dan nomor nya tetap di keep
                                        vPOno = Me.ObjNumbering.GetNomorTransaksiBaru1(vAOPDate, .PO_Urutan, Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Impor)
                                    End If
                                Else
                                    'jika tahun dan bulannya tidak sama maka get no baru
                                    vPOno = ObjPO.GetNoBaru1(vAOPDate)
                                End If
                            End If
                        End If
                    End If

                    '11 Juli 2019 /kodisi saat ini untuk inout numbering masih di hardcode selalu auto numbering dan reset number pertahun di coding, 
                    'jadi coding agak beda dengan SO, DO
                    If Format(.Stock_In_Date, "yyyy") = Format(vAOPDate, "yyyy") Then
                        If Format(.Stock_In_Date, "MMM") = Format(vAOPDate, "MMM") Then
                            'jika sama tahun dan bulannya maka keep no lama
                            vOutStockInNo = .Stock_In_No
                        Else
                            'jika sama tahun tapi bulannya beda maka ganti bulannya saja dan nomor nya tetap di keep
                            vOutStockInNo = ObjInOut.GetNoBaru_20190711(Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase, _
                            vAOPDate, Inventory_Module.clsTransaksiStockInOut.EnumInOutAdjust.NoTipe, .Stock_In_Urutan)
                        End If
                    Else
                        'jika tahun dan bulannya tidak sama maka get no baru
                        If Me.tipeAOP = EnumTipeAop.AOP Or Me.tipeAOP = EnumTipeAop.AOP_DP Then
                            vOutStockInNo = ObjInOut.GetNoBaru_20190711(Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase, vAOPDate)
                        Else
                            vOutStockInNo = ObjInOut.GetNoBaru_20190711(Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Sales, vAOPDate)
                        End If
                    End If
                End With
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
