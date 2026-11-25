Option Explicit On
Option Strict Off

Imports CrystalDecisions.CrystalReports.Engine

Public Class ClsTransaksiSalesInvoice
    Dim ObjCD As New Core_Module.ClsCekData

    'Mata uang yang digunakan untuk DP harus sama dengan mata uang yang digunakan untuk invoice
    Private mPrimary As Double = -1
    ''' <summary>
    ''' digunakan untuk menggenerate no transaksi secara otomatis, default value = 0
    ''' </summary>
    ''' <remarks></remarks>
    Private mNoUrut As Double = 0
    Private mNoTransaksi As String = ""
    Private mTanggalTransaksi As Date = Now
    Private mPrimaryCust As Double = -1
    Private mIdentifier_CustAddress As String = ""
    Private mRemarks As String = ""
    Private mCalculatesBy As EnumCalculatesBy = EnumCalculatesBy.Invoice
    Private mJenisFakturPajak As EnumJenisFakturPajak = EnumJenisFakturPajak.FakturPajakSederhana
    Private mTipeInvoice As EnumTipeInvoice = EnumTipeInvoice.InvoiceDP
    Private mPaymentTerm As String = ""
    Private mRevisi As Integer = 0

    Private mPrimaryCurrency As Double = -1
    Private mSubTotalB4DP As Double = 0
    Private mDownPayment As Double = 0
    Private mSubTotal As Double = 0
    Private mUseTax As Integer = 0
    Private mTaxPercentage As Double = 0
    Private mTaxAmount As Double = 0
    Private mTaxAmount1 As Double = 0 'new tax amount1
    Private mTaxExchangeRate As Double = 1
    Private mTotal As Double = 0
    Private mDutyStamp As Double = 0
    Private mExhangeRateToIDR As Double = 1
    Private mExchangeRateToUSD As Double = 1
    Private mPrimarySO4DP As Double = -1

    Private mNoSeriFakturPajak As String = ""
    Private mNoFakturPajak As String = ""
    Private mUrutanFakturPajak As Double = 0
    Private mNoNPWP As String = ""
    Private mNoFakturPajakSederhana As String = ""
    Private mNamaSigned As String = ""
    Private mJabatanSigned As String = ""
    Private mAlamatPajak1 As String = ""
    Private mAlamatPajak2 As String = ""
    Private mExchangeRate As Decimal = 1
    Private mOngkosKirim As Decimal = 0
    Private mShipmentWithTax As Boolean = True
    Private mVATDipungut As Boolean = True
    Private mDiscString As String = "0"
    Private mDiscAmount As Decimal = 0
    Private mPrimarySalesPerson As Double = -1

    Private mEFak_NamaJalan As String = ""
    Private mEFak_Blok As String = ""
    Private mEFak_No As String = ""
    Private mEFak_RT As String = ""
    Private mEFak_RW As String = ""
    Private mEFak_Kecamatan As String = ""
    Private mEFak_Kelurahan As String = ""
    Private mEFak_Kabupaten As String = ""
    Private mEFak_Propinsi As String = ""
    Private mEFak_KodePos As String = ""
    Private mEFak_Telepon As String = ""
    Private mEFak_KodeJenisTransaksiPajak As Integer = -1

    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now
    Dim ObjNumbering As New Core_Module.ClsMasterTransactionNumbering

    'elemen stock yang digunakan untuk antrian simpan data
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianStockInOut_DO_Invoice
    Dim ObjRnd As New Random
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter
    Dim ObjAntrian As New Core_Module.ClsAntrian

    '//TR : Tambahan baru Nik/Pas
    Private mOptionNIKPASS As Integer = 0
    Private mNoNikPass As String = ""

    '====================================================================================================================================================================
    '====================================================================================================================================================================
    'Erna's Code mulai()
    '====================================================================================================================================================================
    Private mPrimaryCurrTax As Double = -1
    Private mPrimaryAccount As Double = -1
    Private mPrimaryAccountDP As Double = -1
    Private mOtherCharge As Decimal = 0

    '2025 Core Tax Tahap 2.2
    Private mPrimaryAddInfo As Double = -1
    Private mPrimaryFacilityStamp As Double = -1
    Private mTaxType As EnumTaxType = EnumTaxType.Tax04
    Private mTaxRevised As Boolean = False
    Private mSPPBDate As Date = Now
    Private mDPPNilaiLain As Decimal = 0

    '2025 Core Tax Tahap 2.2
    Public Enum EnumTaxType As Integer
        Tax01
        Tax02
        Tax03
        Tax04
        Tax05
        Tax06
        Tax07
        Tax08
        Tax09
        Tax10
    End Enum

    Public Property SPPBDate() As Date
        Get
            Return Me.mSPPBDate
        End Get
        Set(ByVal value As Date)
            mSPPBDate = value
        End Set
    End Property

    Public Property PrimaryAddInfo() As Double
        Get
            Return Me.mPrimaryAddInfo
        End Get
        Set(ByVal value As Double)
            mPrimaryAddInfo = value
        End Set
    End Property

    Public Property PrimaryFacilityStamp() As Double
        Get
            Return Me.mPrimaryFacilityStamp
        End Get
        Set(ByVal value As Double)
            mPrimaryFacilityStamp = value
        End Set
    End Property


    Public Property TaxRevised() As Boolean
        Get
            Return Me.mTaxRevised
        End Get
        Set(ByVal value As Boolean)
            mTaxRevised = value
        End Set
    End Property

    Public Property TaxType() As EnumTaxType
        Get
            Return Me.mTaxType
        End Get
        Set(ByVal value As EnumTaxType)
            mTaxType = value
        End Set
    End Property

    Public Property DPPNilaiLain() As Decimal
        Get
            Return Me.mDPPNilaiLain
        End Get
        Set(ByVal value As Decimal)
            Me.mDPPNilaiLain = value
        End Set
    End Property

    '2025 Core Tax Tahap 2.2 End

    Public Enum EnumTipeDP As Integer
        InvDP
        ManualDP
    End Enum

    Public Property TaxAmount1() As Double 'new tax amount1
        Get
            Return Me.mTaxAmount1
        End Get
        Set(ByVal value As Double)
            mTaxAmount1 = value
        End Set
    End Property

    '//TR : Tambahan baru Nik/Pas
    Public Property OptionNIKPASS() As Integer
        Get
            Return Me.mOptionNIKPASS
        End Get
        Set(ByVal value As Integer)
            mOptionNIKPASS = value
        End Set
    End Property
    Public Property NoNikPass() As String
        Get
            Return Me.mNoNikPass
        End Get
        Set(ByVal value As String)
            mNoNikPass = value
        End Set
    End Property


    Public Property PrimaryCurrTax() As Double
        Get
            Return Me.mPrimaryCurrTax
        End Get
        Set(ByVal value As Double)
            mPrimaryCurrTax = value
        End Set
    End Property

    Public Property PrimaryAccount() As Double
        Get
            Return Me.mPrimaryAccount
        End Get
        Set(ByVal value As Double)
            mPrimaryAccount = value
        End Set
    End Property

    Public Property PrimaryAccountDP() As Double
        Get
            Return Me.mPrimaryAccountDP
        End Get
        Set(ByVal value As Double)
            mPrimaryAccountDP = value
        End Set
    End Property

    Public Property OtherCharge() As Double
        Get
            Return Me.mOtherCharge
        End Get
        Set(ByVal value As Double)
            mOtherCharge = value
        End Set
    End Property

    Public Function GetPrimaryMax(ByVal vPrimaryCustomer As Double) As Double
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_NewTableAdapter

        Try
            If DTA.ScalarQuery_GetPrimaryMax(vPrimaryCustomer).HasValue AndAlso DTA.ScalarQuery_GetPrimaryMax(vPrimaryCustomer).Value > 0 Then
                Return DTA.ScalarQuery_GetPrimaryMax(vPrimaryCustomer).Value
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

    Public Overridable Function CariDetail(ByVal vTbl As DseTransaksiSalesInvoice.tbltrans_salesinvoice_NewDataTable, _
            ByVal vTblOldSalesInvoiceItem As CustomSMESS.DseTransferBeginingBalance.tbltrans_salesinvoice_item_copyDataTable, _
            ByVal vTblOldSalesInvoiceItemExchangeRate As CustomSMESS.DseTransferBeginingBalance.tbltrans_salesinvoice_item_exchangerate_copyDataTable, _
            ByVal vTblOldSalesInvoiceItemSalesReturn As CustomSMESS.DseTransferBeginingBalance.tbltrans_salesinvoice_item_salesreturn_copyDataTable, _
            ByVal vTblOldSalesInvoiceItemInvoiceDp As CustomSMESS.DseTransferBeginingBalance.tbltrans_salesinvoice_itemdp_copyDataTable, _
            ByVal vDgvItem As DataGridView, ByVal vdGVItemDP As DataGridView, ByVal vDGVItemSR As DataGridView, _
            ByVal vDGVItemExchangeRate As DataGridView, ByVal vFrm As FrmTransaksiSalesInvoiceQuick, _
            ByRef vTblCustom As CustomSMESS.DseTransaksiSalesInvoiceQuick_AddByCustomProduct.tbltrans_custom_salesinvoice_item1DataTable, _
            ByRef vTblCustomInOutDummy As CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable) As Boolean

        Dim ObjItem As New ClsTransaksiSalesInvoiceItem
        Dim row As DseTransaksiSalesInvoice.tbltrans_salesinvoice_NewRow
        Dim ObjCustom As New CustomSMESS.ClsTransaksiSalesInvoiceQuick_AddByCustomProduct

        Try
            row = vTbl.Rows(0)
            With row
                Me.mPrimarySO4DP = .Primary_SO_4_DP
                mTipeInvoice = .Tipe_Invoice
                mNoSeriFakturPajak = .No_Seri_Faktur_Pajak
                mNoFakturPajak = .No_Faktur_Pajak
                mUrutanFakturPajak = .Urutan_Faktur_Pajak
                mNoNPWP = .No_NPWP
                Me.mIdentifier_CustAddress = .Address_Identifier
                Me.mNoTransaksi = .No_Transaksi
                Me.mNoUrut = .Urutan
                Me.mPrimary = .Primary
                Me.mPaymentTerm = .Payment_Term
                Me.mPrimaryCust = .Primary_Customer
                Me.mRevisi = .Revisi
                mCalculatesBy = .Calculates_By
                mPrimaryCurrency = .Primary_Currency
                mSubTotalB4DP = .Sub_Total_Before_DP
                mDownPayment = .Down_Payment
                Me.mSubTotal = .Sub_Total
                Me.mUseTax = .Tax_Exclude
                mTaxExchangeRate = .Tax_Exchange_Rate
                Me.mTaxAmount = .Tax_Amount
                Me.mTaxAmount1 = .Tax_Amount1 'new tax amount1
                Me.mTaxPercentage = .Tax
                mDutyStamp = .Duty_Stamp
                Me.mTotal = .Total
                mExhangeRateToIDR = .Exchange_Rate_To_IDR
                mExchangeRateToUSD = .Exchange_Rate_To_USD
                Me.mExchangeRate = .exchange_rate
                Me.mRemarks = .Remarks
                Me.mTanggalTransaksi = .Tanggal_Transaksi
                If .Jenis_Faktur_Pajak = True Then
                    Me.mJenisFakturPajak = EnumJenisFakturPajak.FakturPajakStandar
                Else
                    Me.mJenisFakturPajak = EnumJenisFakturPajak.FakturPajakSederhana
                End If
                Me.mNoFakturPajakSederhana = .No_Faktur_Pajak_Sederhana
                Me.mAlamatPajak1 = .Alamat_Pajak_1
                Me.mAlamatPajak2 = .Alamat_Pajak_2

                Me.mCreatedBy = .Created_By
                Me.mCreatedDate = .Created_Date
                Me.mModifiedBy = .Modified_By
                Me.mModifiedDate = .Modified_Date
                Me.mTimestamp = .Timestamp

                Me.mNamaSigned = .Nama_Signed
                Me.mJabatanSigned = .Jabatan_Signed
                Me.mOngkosKirim = .ongkos_kirim
                If .Shipment_With_VAT = 1 Then
                    Me.mShipmentWithTax = True
                Else
                    Me.mShipmentWithTax = False
                End If
                If .VAT_Dipungut = 1 Then
                    Me.mVATDipungut = True
                Else
                    Me.mVATDipungut = False
                End If

                Me.mDiscAmount = .disc_amount
                Me.mDiscString = .disc_string
                Me.mPrimarySalesPerson = .primary_sales_person

                Me.mEFak_Blok = .Blok
                Me.mEFak_Kabupaten = .Kabupaten
                Me.mEFak_Kecamatan = .Kecamatan
                Me.mEFak_Kelurahan = .Kelurahan
                Me.mEFak_KodePos = .Kode_Pos
                Me.mEFak_NamaJalan = .Nama_Jalan
                Me.mEFak_No = .Nomor
                Me.mEFak_Propinsi = .Propinsi
                Me.mEFak_RT = .RT
                Me.mEFak_RW = .RW
                Me.mEFak_Telepon = .Telepon
                Me.mEFak_KodeJenisTransaksiPajak = .kode_jenis_transaksi_pajak
                Me.mPrimaryAccount = .primary_akun_ar
                Me.mPrimaryCurrTax = .Primary_Curr_Tax
                Me.mPrimaryAccountDP = .Primary_Akun_DP
                Me.mOtherCharge = .Other_Charges

                '2025 Core Tax Tahap 2.2
                Me.mTaxType = .Tax_Type
                Me.mTaxRevised = .Tax_Revised
                Me.mPrimaryAddInfo = .Primary_Additional_Info
                Me.mPrimaryFacilityStamp = .Primary_Cap_Fasilitas
                Me.mSPPBDate = .SPPB_Date
                Me.mDPPNilaiLain = .DPPNilaiLain
            End With
            ObjItem.IsiDGV(Me.mPrimary, vTblOldSalesInvoiceItem, vTblOldSalesInvoiceItemExchangeRate, _
                    vTblOldSalesInvoiceItemSalesReturn, vTblOldSalesInvoiceItemInvoiceDp, _
                    vDgvItem, vDGVItemExchangeRate, vdGVItemDP, Me.mCalculatesBy, vDGVItemSR, vFrm)

            ObjCustom.IsiTblDummyFromDB(Me.mPrimary, vTblCustom)
            ObjCustom.IsiTblOutSales(Me.mPrimary, vTblCustom, vTblCustomInOutDummy) 'new
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function CariDetail(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_NewTableAdapter
        Dim TBL As New DseTransaksiSalesInvoice.tbltrans_salesinvoice_NewDataTable
        Dim row As DseTransaksiSalesInvoice.tbltrans_salesinvoice_NewRow

        Try
            'fil basis data
            If DTA.FillBy(TBL, vPrimaryTransaksi) > 0 Then
                'salin data header
                row = TBL.Rows(0)

                With row
                    Me.mPrimarySO4DP = .Primary_SO_4_DP
                    mTipeInvoice = .Tipe_Invoice
                    mNoSeriFakturPajak = .No_Seri_Faktur_Pajak
                    mNoFakturPajak = .No_Faktur_Pajak
                    mUrutanFakturPajak = .Urutan_Faktur_Pajak
                    mNoNPWP = .No_NPWP
                    Me.mIdentifier_CustAddress = .Address_Identifier
                    Me.mNoTransaksi = .No_Transaksi
                    Me.mNoUrut = .Urutan
                    Me.mPrimary = .Primary
                    Me.mPaymentTerm = .Payment_Term
                    Me.mPrimaryCust = .Primary_Customer
                    Me.mRevisi = .Revisi

                    mCalculatesBy = .Calculates_By
                    mPrimaryCurrency = .Primary_Currency
                    mSubTotalB4DP = .Sub_Total_Before_Dp
                    mDownPayment = .Down_Payment
                    Me.mSubTotal = .Sub_Total
                    Me.mUseTax = .Tax_Exclude
                    mTaxExchangeRate = .Tax_Exchange_Rate
                    Me.mTaxAmount = .Tax_Amount
                    Me.mTaxAmount1 = .Tax_Amount1 'new tax amount1
                    Me.mTaxPercentage = .Tax
                    mDutyStamp = .Duty_Stamp
                    Me.mTotal = .Total
                    mExhangeRateToIDR = .Exchange_Rate_To_IDR
                    mExchangeRateToUSD = .Exchange_Rate_To_USD
                    Me.mRemarks = .Remarks
                    Me.mTanggalTransaksi = .Tanggal_Transaksi
                    If .Jenis_Faktur_Pajak = True Then
                        Me.mJenisFakturPajak = EnumJenisFakturPajak.FakturPajakStandar
                    Else
                        Me.mJenisFakturPajak = EnumJenisFakturPajak.FakturPajakSederhana
                    End If
                    Me.mNoFakturPajakSederhana = .No_Faktur_Pajak_Sederhana
                    Me.mAlamatPajak1 = .Alamat_Pajak_1
                    Me.mAlamatPajak2 = .Alamat_Pajak_2

                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
                    Me.mExchangeRate = .exchange_rate

                    Me.mNamaSigned = .Nama_Signed
                    Me.mJabatanSigned = .Jabatan_Signed
                    Me.mOngkosKirim = .ongkos_kirim
                    If .VAT_Dipungut = 1 Then
                        Me.mVATDipungut = True
                    Else
                        Me.mVATDipungut = False
                    End If
                    If .Shipment_With_VAT = 1 Then
                        Me.mShipmentWithTax = True
                    Else
                        Me.mShipmentWithTax = False
                    End If

                    Me.mDiscString = .disc_string
                    Me.mDiscAmount = .disc_amount
                    Me.mPrimarySalesPerson = .primary_sales_person

                    Me.mEFak_Blok = .Blok
                    Me.mEFak_Kabupaten = .Kabupaten
                    Me.mEFak_Kecamatan = .Kecamatan
                    Me.mEFak_Kelurahan = .Kelurahan
                    Me.mEFak_KodePos = .Kode_Pos
                    Me.mEFak_NamaJalan = .Nama_Jalan
                    Me.mEFak_No = .Nomor
                    Me.mEFak_Propinsi = .Propinsi
                    Me.mEFak_RT = .RT
                    Me.mEFak_RW = .RW
                    Me.mEFak_Telepon = .Telepon
                    Me.mEFak_KodeJenisTransaksiPajak = .kode_jenis_transaksi_pajak
                    Me.mPrimaryAccount = .primary_akun_ar
                    Me.mPrimaryCurrTax = .Primary_Curr_Tax
                    Me.mPrimaryAccountDP = .Primary_Akun_DP
                    Me.mOtherCharge = .Other_Charges

                    '2025 Core Tax Tahap 2.2
                    Me.mTaxType = .Tax_Type
                    Me.mTaxRevised = .Tax_Revised
                    Me.mPrimaryAddInfo = .Primary_Additional_Info
                    Me.mPrimaryFacilityStamp = .Primary_Cap_Fasilitas
                    Me.mSPPBDate = .SPPB_Date
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
    ByVal vDgvItem As DataGridView, ByVal vdGVItemDP As DataGridView, ByVal vDGVItemSR As DataGridView, _
    ByVal vDGVItemExchangeRate As DataGridView, ByVal vFrm As FrmTransaksiSalesInvoice) As Boolean
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_NewTableAdapter
        Dim TBL As New DseTransaksiSalesInvoice.tbltrans_salesinvoice_NewDataTable

        Dim row As DseTransaksiSalesInvoice.tbltrans_salesinvoice_NewRow
        Dim ObjItem As New ClsTransaksiSalesInvoiceItem

        Try
            'fil basis data
            If DTA.FillBy(TBL, vPrimaryTransaksi) > 0 Then
                'salin data header
                row = TBL.Rows(0)

                With row
                    Me.mPrimarySO4DP = .Primary_SO_4_DP
                    mTipeInvoice = .Tipe_Invoice
                    mNoSeriFakturPajak = .No_Seri_Faktur_Pajak
                    mNoFakturPajak = .No_Faktur_Pajak
                    mUrutanFakturPajak = .Urutan_Faktur_Pajak
                    mNoNPWP = .No_NPWP
                    Me.mIdentifier_CustAddress = .Address_Identifier
                    Me.mNoTransaksi = .No_Transaksi
                    Me.mNoUrut = .Urutan
                    Me.mPrimary = .Primary
                    Me.mPaymentTerm = .Payment_Term
                    Me.mPrimaryCust = .Primary_Customer
                    Me.mRevisi = .Revisi
                    mCalculatesBy = .Calculates_By
                    mPrimaryCurrency = .Primary_Currency
                    mSubTotalB4DP = .Sub_Total_Before_DP
                    mDownPayment = .Down_Payment
                    Me.mSubTotal = .Sub_Total
                    Me.mUseTax = .Tax_Exclude
                    mTaxExchangeRate = .Tax_Exchange_Rate
                    Me.mTaxAmount = .Tax_Amount
                    Me.mTaxAmount1 = .Tax_Amount1 'new tax amount1
                    Me.mTaxPercentage = .Tax
                    mDutyStamp = .Duty_Stamp
                    Me.mTotal = .Total
                    mExhangeRateToIDR = .Exchange_Rate_To_IDR
                    mExchangeRateToUSD = .Exchange_Rate_To_USD
                    Me.mExchangeRate = .exchange_rate
                    Me.mRemarks = .Remarks
                    Me.mTanggalTransaksi = .Tanggal_Transaksi
                    If .Jenis_Faktur_Pajak = True Then
                        Me.mJenisFakturPajak = EnumJenisFakturPajak.FakturPajakStandar
                    Else
                        Me.mJenisFakturPajak = EnumJenisFakturPajak.FakturPajakSederhana
                    End If
                    Me.mNoFakturPajakSederhana = .No_Faktur_Pajak_Sederhana
                    Me.mAlamatPajak1 = .Alamat_Pajak_1
                    Me.mAlamatPajak2 = .Alamat_Pajak_2

                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp

                    Me.mNamaSigned = .Nama_Signed
                    Me.mJabatanSigned = .Jabatan_Signed
                    Me.mOngkosKirim = .ongkos_kirim
                    If .VAT_Dipungut = 1 Then
                        Me.mVATDipungut = True
                    Else
                        Me.mVATDipungut = False
                    End If
                    If .Shipment_With_VAT = 1 Then
                        Me.mShipmentWithTax = True
                    Else
                        Me.mShipmentWithTax = False
                    End If

                    Me.mPrimarySalesPerson = .primary_sales_person
                    Me.mDiscString = .disc_string
                    Me.mDiscAmount = .disc_amount

                    Me.mEFak_Blok = .Blok
                    Me.mEFak_Kabupaten = .Kabupaten
                    Me.mEFak_Kecamatan = .Kecamatan
                    Me.mEFak_Kelurahan = .Kelurahan
                    Me.mEFak_KodePos = .Kode_Pos
                    Me.mEFak_NamaJalan = .Nama_Jalan
                    Me.mEFak_No = .Nomor
                    Me.mEFak_Propinsi = .Propinsi
                    Me.mEFak_RT = .RT
                    Me.mEFak_RW = .RW
                    Me.mEFak_Telepon = .Telepon
                    Me.mEFak_KodeJenisTransaksiPajak = .kode_jenis_transaksi_pajak
                    Me.mPrimaryAccount = .primary_akun_ar
                    Me.mPrimaryCurrTax = .Primary_Curr_Tax
                    Me.mPrimaryAccountDP = .Primary_Akun_DP
                    Me.mOtherCharge = .Other_Charges

                    '//TR : Tambahan baru Nik/Pas
                    Me.mOptionNIKPASS = .Pilih_NIKPASS
                    Me.mNoNikPass = .No_NIKPASS

                    '2025 Core Tax Tahap 2.2
                    Me.mTaxType = .Tax_Type
                    Me.mTaxRevised = .Tax_Revised
                    Me.mPrimaryAddInfo = .Primary_Additional_Info
                    Me.mPrimaryFacilityStamp = .Primary_Cap_Fasilitas
                    Me.mSPPBDate = .SPPB_Date
                    Me.mDPPNilaiLain = .DPPNilaiLain

                End With

                'salin data item 
                ObjItem.IsiDGV(vDgvItem, vDGVItemExchangeRate, vdGVItemDP, Me.mPrimary, Me.mCalculatesBy, vDGVItemSR, vFrm)
                'nilai balikan
                Return True
            Else
                Me.ResetObject(vDgvItem, vdGVItemDP, vDGVItemExchangeRate, vDGVItemSR)
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
   ByVal vDgvItem As DataGridView, ByVal vdGVItemDP As DataGridView, ByVal vDGVItemSR As DataGridView, _
   ByVal vDGVItemExchangeRate As DataGridView, ByVal vFrm As FrmTransaksiSalesInvoiceQuick, _
   ByRef vTBL As CustomSMESS.DseTransaksiSalesInvoiceQuick_AddByCustomProduct.tbltrans_custom_salesinvoice_item1DataTable, _
    ByRef vTblCustomInOutDummy As CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable) As Boolean
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_NewTableAdapter
        Dim TBL As New DseTransaksiSalesInvoice.tbltrans_salesinvoice_NewDataTable

        Dim row As DseTransaksiSalesInvoice.tbltrans_salesinvoice_NewRow
        Dim ObjItem As New ClsTransaksiSalesInvoiceItem
        Dim ObjCustom As New CustomSMESS.ClsTransaksiSalesInvoiceQuick_AddByCustomProduct
        Try
            'fil basis data
            If DTA.FillBy(TBL, vPrimaryTransaksi) > 0 Then
                'salin data header
                row = TBL.Rows(0)

                With row
                    Me.mPrimarySO4DP = .Primary_SO_4_DP
                    mTipeInvoice = .Tipe_Invoice
                    mNoSeriFakturPajak = .No_Seri_Faktur_Pajak
                    mNoFakturPajak = .No_Faktur_Pajak
                    mUrutanFakturPajak = .Urutan_Faktur_Pajak
                    mNoNPWP = .No_NPWP
                    Me.mIdentifier_CustAddress = .Address_Identifier
                    Me.mNoTransaksi = .No_Transaksi
                    Me.mNoUrut = .Urutan
                    Me.mPrimary = .Primary
                    Me.mPaymentTerm = .Payment_Term
                    Me.mPrimaryCust = .Primary_Customer
                    Me.mRevisi = .Revisi
                    mCalculatesBy = .Calculates_By
                    mPrimaryCurrency = .Primary_Currency
                    mSubTotalB4DP = .Sub_Total_Before_DP
                    mDownPayment = .Down_Payment
                    Me.mSubTotal = .Sub_Total
                    Me.mUseTax = .Tax_Exclude
                    mTaxExchangeRate = .Tax_Exchange_Rate
                    Me.mTaxAmount = .Tax_Amount
                    Me.mTaxAmount1 = .Tax_Amount1 'new tax amount1
                    Me.mTaxPercentage = .Tax
                    mDutyStamp = .Duty_Stamp
                    Me.mTotal = .Total
                    mExhangeRateToIDR = .Exchange_Rate_To_IDR
                    mExchangeRateToUSD = .Exchange_Rate_To_USD
                    Me.mExchangeRate = .exchange_rate
                    Me.mRemarks = .Remarks
                    Me.mTanggalTransaksi = .Tanggal_Transaksi
                    If .Jenis_Faktur_Pajak = True Then
                        Me.mJenisFakturPajak = EnumJenisFakturPajak.FakturPajakStandar
                    Else
                        Me.mJenisFakturPajak = EnumJenisFakturPajak.FakturPajakSederhana
                    End If
                    Me.mNoFakturPajakSederhana = .No_Faktur_Pajak_Sederhana
                    Me.mAlamatPajak1 = .Alamat_Pajak_1
                    Me.mAlamatPajak2 = .Alamat_Pajak_2

                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp

                    Me.mNamaSigned = .Nama_Signed
                    Me.mJabatanSigned = .Jabatan_Signed
                    Me.mOngkosKirim = .ongkos_kirim
                    If .Shipment_With_VAT = 1 Then
                        Me.mShipmentWithTax = True
                    Else
                        Me.mShipmentWithTax = False
                    End If
                    If .VAT_Dipungut = 1 Then
                        Me.mVATDipungut = True
                    Else
                        Me.mVATDipungut = False
                    End If

                    Me.mDiscAmount = .disc_amount
                    Me.mDiscString = .disc_string
                    Me.mPrimarySalesPerson = .primary_sales_person

                    Me.mEFak_Blok = .Blok
                    Me.mEFak_Kabupaten = .Kabupaten
                    Me.mEFak_Kecamatan = .Kecamatan
                    Me.mEFak_Kelurahan = .Kelurahan
                    Me.mEFak_KodePos = .Kode_Pos
                    Me.mEFak_NamaJalan = .Nama_Jalan
                    Me.mEFak_No = .Nomor
                    Me.mEFak_Propinsi = .Propinsi
                    Me.mEFak_RT = .RT
                    Me.mEFak_RW = .RW
                    Me.mEFak_Telepon = .Telepon
                    Me.mEFak_KodeJenisTransaksiPajak = .kode_jenis_transaksi_pajak
                    Me.mPrimaryAccount = .primary_akun_ar
                    Me.mPrimaryCurrTax = .Primary_Curr_Tax
                    Me.mPrimaryAccountDP = .Primary_Akun_DP
                    Me.mOtherCharge = .Other_Charges
                    '//TR : Tambahan baru Nik/Pas
                    Me.mOptionNIKPASS = .Pilih_NIKPASS
                    Me.mNoNikPass = .No_NIKPASS

                    '2025 Core Tax Tahap 2.2
                    Me.mTaxType = .Tax_Type
                    Me.mTaxRevised = .Tax_Revised
                    Me.mPrimaryAddInfo = .Primary_Additional_Info
                    Me.mPrimaryFacilityStamp = .Primary_Cap_Fasilitas
                    Me.mSPPBDate = .SPPB_Date
                    Me.mDPPNilaiLain = .DPPNilaiLain
                End With

                'salin data item 
                ObjItem.IsiDGV(vDgvItem, vDGVItemExchangeRate, vdGVItemDP, Me.mPrimary, Me.mCalculatesBy, vDGVItemSR, vFrm)

                ObjCustom.IsiTblDummyFromDB(Me.mPrimary, vTBL)
                ObjCustom.IsiTblOutSales(Me.mPrimary, vTBL, vTblCustomInOutDummy) 'new
                'nilai balikan
                Return True
            Else
                Me.ResetObject(vDgvItem, vdGVItemDP, vDGVItemExchangeRate, vDGVItemSR)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjItem = Nothing
        End Try
    End Function


    Public Overridable Function SimpanBaru1(ByVal vdgvItem As DataGridView, ByVal vdgvItemDP As DataGridView, _
    ByVal vDGVItemSR As DataGridView, ByVal vDGVItemExchangeRate As DataGridView, ByVal vUserCode As String, ByVal vFrm As FrmTransaksiSalesInvoice, _
    Optional ByVal vCekDataBenar As Boolean = True, Optional ByVal vFromCutOff As Boolean = False) As Boolean
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice1TableAdapter
        Dim DTA2 As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_NewTableAdapter
        Dim DTAHeaderCost As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_add_dataTableAdapter
        Dim DTAItemCost As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_item_costTableAdapter
        Dim ObjItem As New ClsTransaksiSalesInvoiceItem
        Dim UrutanFakturPajak As String = ""
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim NoUrut As Double = 0
        Dim Obj As New ClsTransaksiSalesInvoice_SettingTaxInvoice
        Dim PrimaryFakturPajak As Double = -1
        Dim NoFakturPajakMax As String = ""
        Dim DefPrecBase As Integer = 0

        Dim DPOriginal, DPBaseCurrency As Decimal
        Dim DTA1 As New DseTransaksiSalesInvoiceTableAdapters.DataTable13TableAdapter
        Dim TBL1 As New DseTransaksiSalesInvoice.DataTable13DataTable
        Dim row1 As DseTransaksiSalesInvoice.DataTable13Row
        Dim TipeDP1 As EnumTipeDP = EnumTipeDP.InvDP
        Dim TBL2 As New DseTransaksiSalesInvoice.DataTable13DataTable
        Dim row2 As DseTransaksiSalesInvoice.DataTable13Row
        Dim TipeDP2 As EnumTipeDP = EnumTipeDP.ManualDP
        Dim PrimaryAkunDP As Double = -1
        Dim vShipmentTax, vVatDipungut As Integer

        Dim CreatedBy_ As String = vUserCode
        Dim CreatedDate_ As Date = Now
        Dim ModifiedBy_ As String = vUserCode
        Dim ModifiedDate_ As Date = Now
        Dim Timestamp_ As Date = Now

        Try
            If Me.ApakahDataBenar(True, vdgvItem, vFrm, vCekDataBenar, vFromCutOff) = True Then
                DefPrecBase = ObjCur.GetDecimal_DefaultCurrency

                '2025 Core Tax Tahap 2.2
                ''yulimus-optimasi-cut off
                'If vFromCutOff = False Then
                '    If Me.mTanggalTransaksi >= Obj.BatasStartDate Then
                '        Me.mUrutanFakturPajak = Obj.GetUrutan(Me.mNoFakturPajak, Me.mTanggalTransaksi)
                '    Else
                '        If Me.mJenisFakturPajak = EnumJenisFakturPajak.FakturPajakStandar Then
                '            Me.mUrutanFakturPajak = Me.NoFakturPajak.Substring(4)
                '        Else
                '            Me.mUrutanFakturPajak = 0
                '            Me.mNoSeriFakturPajak = ""
                '            Me.mNoFakturPajak = ""
                '        End If
                '    End If
                'End If

                NoUrut = Me.ObjNumbering.GetUrutan(Me.mNoTransaksi, Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Invoice, _
                Me.mTanggalTransaksi)

                If Me.mShipmentWithTax = True Then
                    vShipmentTax = 1
                Else
                    vShipmentTax = 0
                End If

                If Me.mVATDipungut = True Then
                    vVatDipungut = 1
                Else
                    vVatDipungut = 0
                End If

                If vFromCutOff = True Then
                    CreatedBy_ = Me.mCreatedBy
                    CreatedDate_ = Me.mCreatedDate
                    ModifiedBy_ = Me.mModifiedBy
                    ModifiedDate_ = Me.mModifiedDate
                    Timestamp_ = Me.mTimestamp
                End If

                'insert query
                If DTA2.InsertQuery2(Me.mNoTransaksi, NoUrut, Me.mUrutanFakturPajak, _
                Me.mTipeInvoice, Me.mTanggalTransaksi, Me.mPrimaryCust, Me.mPrimaryCurrency, Me.mCalculatesBy, _
                Me.mSubTotalB4DP, Me.mDownPayment, Me.mSubTotal, Me.mDutyStamp, Me.mUseTax, Me.mTaxPercentage, _
                Me.mTaxExchangeRate, Me.mTaxAmount, Me.mTotal, Me.mExchangeRateToUSD, Me.mExhangeRateToIDR, _
                Me.mNoSeriFakturPajak, Me.mNoFakturPajak, Me.mNoNPWP, Me.mIdentifier_CustAddress, _
                Me.mRemarks, CreatedBy_, CreatedDate_, ModifiedBy_, ModifiedDate_, Timestamp_, Me.mJenisFakturPajak, Me.mPaymentTerm, _
                Me.mPrimarySO4DP, Me.mNoFakturPajakSederhana, Me.mNamaSigned, Me.mJabatanSigned, 0, Me.mRevisi, _
                Me.mAlamatPajak1, Me.mAlamatPajak2, Me.mExchangeRate, Me.mOngkosKirim, vVatDipungut, vShipmentTax, _
                Me.mPrimarySalesPerson, Me.mDiscString, Me.mDiscAmount, Me.mEFak_NamaJalan, Me.mEFak_Blok, _
                Me.mEFak_No, Me.mEFak_RT, Me.mEFak_RW, Me.mEFak_Kecamatan, Me.mEFak_Kelurahan, Me.mEFak_Kabupaten, _
                Me.mEFak_Propinsi, Me.mEFak_KodePos, Me.mEFak_Telepon, Me.mEFak_KodeJenisTransaksiPajak, _
                Me.mPrimaryAccount, Me.mPrimaryCurrTax, Me.mPrimaryAccountDP, Me.mOtherCharge, Me.mOptionNIKPASS, Me.mNoNikPass, Me.mTaxAmount1, _
                Me.PrimaryAddInfo, Me.PrimaryFacilityStamp, Me.TaxRevised, Me.TaxType, Me.SPPBDate, Me.mDPPNilaiLain) > 0 Then  '//TR : Tambahan baru Nik/Pas
                    'update data yang berubah setelah insert query
                    Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mNoTransaksi).Value
                    Me.mNoUrut = NoUrut
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    '2025 Core Tax Tahap 2.2
                    ''update data berkaitan dengan no faktur pajak
                    ''yulimus-optimasi-cut off
                    'If vFromCutOff = False Then
                    '    If Me.mJenisFakturPajak = EnumJenisFakturPajak.FakturPajakStandar And Me.UseTax = True Then
                    '        While DTA.ScalarQuery_GetJmlByNoFakturPajak(Me.mNoFakturPajak).Value > 1
                    '            'MsgBox(DTA.ScalarQuery_GetJmlByNoFakturPajak(Me.mNoFakturPajak).Value.ToString)
                    '            If DTA.ScalarQuery_GetPrimaryMaxByNoFakturPajak(Me.mNoFakturPajak).Value = _
                    '            Me.mPrimary Then
                    '                'get no baru
                    '                Me.mNoFakturPajak = Me.GetNoFakturPajakStandarBaru(Me.mTanggalTransaksi)
                    '                'Exit While

                    '                'update no-nya
                    '                If Me.mTanggalTransaksi >= Obj.BatasStartDate Then
                    '                    Me.mUrutanFakturPajak = Obj.GetUrutan(Me.mNoFakturPajak, Me.mTanggalTransaksi)
                    '                Else
                    '                    If Me.mJenisFakturPajak = EnumJenisFakturPajak.FakturPajakStandar Then
                    '                        Me.mUrutanFakturPajak = Me.NoFakturPajak.Substring(4)
                    '                    Else
                    '                        Me.mUrutanFakturPajak = 0
                    '                        Me.mNoSeriFakturPajak = ""
                    '                        Me.mNoFakturPajak = ""
                    '                    End If
                    '                End If

                    '                DTA.UpdateQuery_UpdateNoFakturPajak(Me.mUrutanFakturPajak, Me.mNoFakturPajak, _
                    '                Me.mPrimary)
                    '                'Exit While
                    '            End If
                    '        End While
                    '    End If
                    'End If

                    'update data yang berkaitan dengan item exchange rate
                    If Me.mCalculatesBy = EnumCalculatesBy.Invoice Then
                        vDGVItemExchangeRate.Rows(0).Cells(5).Value = Me.mPrimary
                        vDGVItemExchangeRate.Rows(0).Cells(1).Value = Me.mNoTransaksi
                    End If

                    'simpan berkaitan other charge
                    If vFrm.TBLHeaderCost.Rows.Count > 0 Then
                        For Each vFrm.rowHeaderCost In vFrm.TBLHeaderCost.Rows
                            With vFrm.rowHeaderCost
                                DTAHeaderCost.InsertQuery(Me.mPrimary, .ETD_Date, .ETD_Port, .ETA_Date, .ETA_Port, .Total_Nett_Weight, _
                                .Total_Gross_Weight, .primary_weight_unit, .Total_Nett_Volume, .Total_Gross_Volume, .primary_volume_unit, _
                                .primary_cust_ship_to, .identifier_cust_ship_to, .shipping_mark, .transportation_mode, .invoice_type, _
                                .var1_name, .var1_value, .var2_name, .var2_value, .var3_name, .var3_value, .var4_name, .var4_value, _
                                .var5_name, .var5_value)
                            End With
                        Next
                    End If

                    If vFrm.TBLitemCost.Rows.Count > 0 Then
                        For Each vFrm.rowitemCost In vFrm.TBLitemCost.Rows
                            With vFrm.rowitemCost
                                DTAItemCost.InsertQuery(Me.mPrimary, .Cost_Description, .Cost, .Remarks, .Primary_Akun)
                            End With
                        Next
                    End If

                    'simpan data item baru
                    ObjItem.Simpan(vdgvItem, vDGVItemExchangeRate, vdgvItemDP, Me.mPrimary, vDGVItemSR, vFrm)

                    'cek berkaitan nomor sama untuk beda bulan -> obselete as off 2017-01-31 by iwan
                    'If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Invoice) = True Then
                    '    Me.CekNomor2(DTA)
                    'End If

                    'simpan data sales amount
                    Me.SimpanDataSalesAmount(Me.mCalculatesBy, Me.mPrimary, _
                    Math.Round(Me.mSubTotalB4DP * Me.mTaxExchangeRate, DefPrecBase, MidpointRounding.AwayFromZero))

                    ''hitung DP
                    'DTA1.Fill(TBL1, Me.mPrimary)
                    'DPOriginal = 0
                    'DPBaseCurrency = 0
                    'Dim PrimaryAkunDP As Double = -1
                    'For Each row1 In TBL1.Rows
                    '    DPOriginal += ObjCD.DigitGrouping(row1.Sub_Total_Before_DP, ObjCur.GetDefaultPrecission(Me.mPrimaryCurrTax))
                    '    DPBaseCurrency += ObjCD.DigitGrouping(row1.Sub_Total_Before_DP * row1.Tax_Exchange_Rate, DefPrecBase)
                    '    PrimaryAkunDP = Me.GetPrimaryAccountDP(row1.Primary_Invoice_DP)
                    'Next
                    'hitung DP
                    DPOriginal = 0
                    DPBaseCurrency = 0
                    If Me.mDownPayment > 0 Then
                        DTA1.Fill(TBL1, Me.mPrimary, CInt(TipeDP1))
                        For Each row1 In TBL1.Rows
                            DPOriginal += ObjCD.DigitGrouping(row1.Sub_Total_Before_DP, ObjCur.GetDefaultPrecission(Me.mPrimaryCurrTax))
                            DPBaseCurrency += ObjCD.DigitGrouping(row1.Sub_Total_Before_DP * row1.Tax_Exchange_Rate, DefPrecBase)
                            If PrimaryAkunDP = -1 Then
                                PrimaryAkunDP = Me.GetPrimaryAccountDP(row1.Primary_Invoice_DP)
                            End If
                        Next

                        DTA1.FillBy_DPBalance(TBL2, Me.mPrimary, CInt(TipeDP2))
                        For Each row2 In TBL2.Rows
                            DPOriginal += ObjCD.DigitGrouping(row2.Sub_Total_Before_DP, ObjCur.GetDefaultPrecission(Me.mPrimaryCurrTax))
                            DPBaseCurrency += ObjCD.DigitGrouping(row2.Sub_Total_Before_DP * row2.Tax_Exchange_Rate, DefPrecBase)
                            If PrimaryAkunDP = -1 Then
                                PrimaryAkunDP = Me.GetPrimaryAccountDP_BalanceDP(row2.Primary_Invoice_DP)
                            End If
                        Next
                    End If

                    '**************************************************************************************************************************************
                    '**************************************************************************************************************************************
                    'PERUBAHAN SISTEM PERHITUNGAN JURNAL ERROR NO 50 TANGGAL 11 JULI 2016 -- PERBAIKAN 28 JULI 2016
                    'SEBELUMNYA NILAI SALES RUPIAH = NILAI SALES USD * EXCHANGE RATE
                    'SEKARANG NILAINYA DICEK DULU APAKAH ADA PENJUMLAHAN TERHADAP NILAI SALES (CONTOH TAX USD, SHIPMENT DLL),
                    'HASILNYA DITOTAL LALU DIKURANGI SEMUA PENJUMLAHAN, SISANYA DIANGGAP SEBAGAI SALES
                    'HAL INI DILAKUKAN UTK MENCEGAH SELISIH ANTARA SALES DI INVOICE JURNAL DENGAN SALES DI RECEIVEMENT JOURNAL


                    Dim SalesRupiah As Decimal = 0
                    Dim TotalSales As Decimal = 0
                    Dim AmountTanpaSales As Decimal = 0
                    Dim SalesPlusDP As Decimal = 0


                    SalesPlusDP = Me.ObjCD.DigitGrouping(Me.mTotal + DPOriginal, ObjCur.GetDefaultPrecission(Me.mPrimaryCurrency))
                    TotalSales = Me.ObjCD.DigitGrouping(SalesPlusDP * Me.mTaxExchangeRate, DefPrecBase)
                    AmountTanpaSales += Me.ObjCD.DigitGrouping(Me.mOngkosKirim * Me.mTaxExchangeRate, DefPrecBase)
                    AmountTanpaSales += Me.ObjCD.DigitGrouping(Me.mDiscAmount * Me.mTaxExchangeRate, DefPrecBase)
                    AmountTanpaSales += Me.ObjCD.DigitGrouping(Me.mOtherCharge * Me.mTaxExchangeRate, DefPrecBase)
                    If Me.mPrimaryCurrTax = Me.mPrimaryCurrency Then
                        AmountTanpaSales += Me.ObjCD.DigitGrouping(Me.mTaxAmount * Me.mTaxExchangeRate, DefPrecBase)
                        AmountTanpaSales += Me.ObjCD.DigitGrouping(Me.mDutyStamp * Me.mTaxExchangeRate, DefPrecBase)
                    End If
                    'AmountTanpaSalesRp = Me.ObjCD.DigitGrouping((AmountTanpaSales * Me.mExchangeRate), DefPrecBase)
                    SalesRupiah = Me.ObjCD.DigitGrouping(TotalSales - AmountTanpaSales, DefPrecBase)

                    '**************************************************************************************************************************************


                    'simpan jurnal otomatis
                    DefPrecBase = ObjCur.GetDecimal_DefaultCurrency
                    Me.TransferJurnal_Invoice2(Me.mPrimary, Me.mPrimaryCust, _
                    SalesRupiah, _
                    ObjCD.DigitGrouping(Me.mTaxAmount * Me.mTaxExchangeRate, DefPrecBase), _
                    ObjCD.DigitGrouping(Me.mDutyStamp * Me.mTaxExchangeRate, DefPrecBase), _
                    Me.mNoUrut, Me.mTanggalTransaksi, Me.mNoTransaksi, Me.mModifiedBy, CDbl(Me.mPaymentTerm), Me.mTaxExchangeRate, _
                    Me.mSubTotalB4DP, Me.mPrimaryCurrency, ObjCD.DigitGrouping(Me.mOngkosKirim * Me.mTaxExchangeRate, DefPrecBase), _
                    DPOriginal, DPBaseCurrency, Me.mOngkosKirim, Me.mDiscAmount, _
                    ObjCD.DigitGrouping(Me.mDiscAmount * Me.mTaxExchangeRate, DefPrecBase), _
                    Me.mPrimaryCurrTax, Me.mPrimaryAccount, Me.mDutyStamp, Me.mTaxAmount, PrimaryAkunDP, _
                    Me.ObjCD.DigitGrouping(Me.mOtherCharge * Me.mTaxExchangeRate, DefPrecBase), _
                    Me.mOtherCharge, vFrm.TBLitemCost, vFrm.rowitemCost)

                    '2025 Core Tax Tahap 2.2
                    ''simpan no urut max untuk numbering
                    'If Obj.BatasStartDate <= Me.mTanggalTransaksi Then
                    '    If Me.mJenisFakturPajak = EnumJenisFakturPajak.FakturPajakStandar Then
                    '        DTA.UpdateQuery_PrimaryFakturPajak(Obj.GetPrimary(Me.mTanggalTransaksi, Me.mNoFakturPajak), Me.mPrimary)
                    '        If vFromCutOff = False Then
                    '            PrimaryFakturPajak = DTA.ScalarQuery_GetPrimaryFakturPajak(Me.mPrimary)
                    '            NoFakturPajakMax = Me.GetNoFakturPajakMax(PrimaryFakturPajak, Me.mPrimary)

                    '            If Obj.GetUrutan(NoFakturPajakMax, Me.mTanggalTransaksi) < Obj.GetUrutan(Me.mNoFakturPajak, Me.mTanggalTransaksi) Then
                    '                Obj.SetNoUrutMaksimal(Me.mTanggalTransaksi, Me.mNoFakturPajak)
                    '            Else
                    '                Obj.SetNoUrutMaksimal(PrimaryFakturPajak, NoFakturPajakMax)
                    '            End If
                    '        End If
                    '    End If
                    'End If


                    ''simpan jurnal otomatis
                    'DefPrecBase = ObjCur.GetDecimal_DefaultCurrency
                    'Me.TransferJurnal_Invoice2(Me.mPrimary, Me.mPrimaryCust, _
                    'Math.Round(Me.mSubTotalB4DP * Me.mTaxExchangeRate, DefPrecBase, MidpointRounding.AwayFromZero), _
                    'Math.Round(Me.mTaxAmount * Me.mTaxExchangeRate, DefPrecBase, MidpointRounding.AwayFromZero), _
                    'Math.Round(Me.mDutyStamp * Me.mTaxExchangeRate, DefPrecBase, MidpointRounding.AwayFromZero), _
                    'Me.mNoUrut, Me.mTanggalTransaksi, Me.mNoTransaksi, Me.mModifiedBy, CDbl(Me.mPaymentTerm), Me.mTaxExchangeRate, _
                    'Me.mSubTotalB4DP, Me.mPrimaryCurrency, Math.Round(Me.mOngkosKirim * Me.mTaxExchangeRate, DefPrecBase, MidpointRounding.AwayFromZero), _
                    'DPOriginal, DPBaseCurrency, Me.mOngkosKirim, Me.mDiscAmount, _
                    'Math.Round(Me.mDiscAmount * Me.mTaxExchangeRate, DefPrecBase, MidpointRounding.AwayFromZero), _
                    'Me.mPrimaryCurrTax, Me.mPrimaryAccount, Me.mDutyStamp, Me.mTaxAmount, PrimaryAkunDP, _
                    'Me.mOtherCharge * Me.mTaxExchangeRate, Me.mOtherCharge, vFrm.TBLitemCost, vFrm.rowitemCost)
                    ''simpan no urut max untuk numbering
                    'If Obj.BatasStartDate <= Me.mTanggalTransaksi Then
                    '    If Me.mJenisFakturPajak = EnumJenisFakturPajak.FakturPajakStandar Then
                    '        DTA.UpdateQuery_PrimaryFakturPajak(Obj.GetPrimary(Me.mTanggalTransaksi, Me.mNoFakturPajak), Me.mPrimary)
                    '        PrimaryFakturPajak = DTA.ScalarQuery_GetPrimaryFakturPajak(Me.mPrimary)
                    '        NoFakturPajakMax = Me.GetNoFakturPajakMax(PrimaryFakturPajak, Me.mPrimary)

                    '        If Obj.GetUrutan(NoFakturPajakMax, Me.mTanggalTransaksi) < Obj.GetUrutan(Me.mNoFakturPajak, Me.mTanggalTransaksi) Then
                    '            Obj.SetNoUrutMaksimal(Me.mTanggalTransaksi, Me.mNoFakturPajak)
                    '        Else
                    '            Obj.SetNoUrutMaksimal(PrimaryFakturPajak, NoFakturPajakMax)
                    '        End If
                    '    End If
                    'End If

                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mNoTransaksi + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Invoice) = True Then
                    Me.mNoTransaksi = Me.GetNoBaru1(Me.mTanggalTransaksi)
                    Return Me.SimpanBaru1(vdgvItem, vdgvItemDP, vDGVItemSR, vDGVItemExchangeRate, vUserCode, vFrm, vCekDataBenar, vFromCutOff)
                Else
                    MsgBox("Invoice number duplication, please change invoice number.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            DTA2.Dispose()
            ObjItem = Nothing
            ObjCur = Nothing
            DTA1.Dispose()
            TBL1.Dispose()
        End Try
    End Function

    Public Overridable Function SimpanBaru(ByVal vdgvItem As DataGridView, ByVal vdgvItemDP As DataGridView, _
    ByVal vDGVItemSR As DataGridView, ByVal vDGVItemExchangeRate As DataGridView, ByVal vUserCode As String, ByVal vFrm As FrmTransaksiSalesInvoice) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            SW.Start()

            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
            Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiInvoice_SimpanBaru, vUserCode)
            'Me.GetNoAntrian()

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()

            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > Me.ObjAntrian.GetWaitTime Then
                    If SW.Elapsed.TotalSeconds >= ObjAntrian.GetWaitTime Then
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                        Exit Function
                    Else
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanBaru1(vdgvItem, vdgvItemDP, _
                            vDGVItemSR, vDGVItemExchangeRate, vUserCode, vFrm)
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

    Public Overridable Function SimpanBaru(ByVal vdgvItem As DataGridView, ByVal vdgvItemDP As DataGridView, _
  ByVal vDGVItemSR As DataGridView, ByVal vDGVItemExchangeRate As DataGridView, ByVal vUserCode As String, _
  ByVal vFrm As FrmTransaksiSalesInvoiceQuick, ByVal vNoDO As String, _
  ByVal vTblCustom As CustomSMESS.DseTransaksiSalesInvoiceQuick_AddByCustomProduct.tbltrans_custom_salesinvoice_item1DataTable, _
    ByVal TblCustomInOutDummy As CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable, Optional ByVal vCekDataBenar As Boolean = True) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            SW.Start()

            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
           Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiInvoice_Quick_SimpanBaru, vUserCode)

            'Me.GetNoAntrian()

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()

            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > ObjAntrian.GetWaitTime Then
                    If SW.Elapsed.TotalSeconds >= ObjAntrian.GetWaitTime Then
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                        Exit Function
                    Else
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanBaru1(vdgvItem, vdgvItemDP, _
                            vDGVItemSR, vDGVItemExchangeRate, vUserCode, vFrm, vCekDataBenar, vNoDO, vTblCustom, TblCustomInOutDummy)
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

    Public Overridable Function SimpanBaru1(ByVal vdgvItem As DataGridView, ByVal vdgvItemDP As DataGridView, _
    ByVal vDGVItemSR As DataGridView, ByVal vDGVItemExchangeRate As DataGridView, ByVal vUserCode As String, _
    ByVal vFrm As FrmTransaksiSalesInvoiceQuick, ByVal vCekApakahDataBenar As Boolean, ByVal vNoDO As String, _
    ByVal vTblCustom As CustomSMESS.DseTransaksiSalesInvoiceQuick_AddByCustomProduct.tbltrans_custom_salesinvoice_item1DataTable, _
    ByVal TblCustomInOutDummy As CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable, _
    Optional ByVal vSetSOB As Boolean = True, Optional ByVal vKoreksiStock As Boolean = True, _
    Optional ByVal vFromCutOff As Boolean = False, Optional ByVal vNoSO As String = "", Optional ByVal vNoOutSales As String = "") As Boolean
        Dim DTAHeaderCost As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_add_dataTableAdapter
        Dim DTAItemCost As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_item_costTableAdapter
        Dim DTA2 As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_NewTableAdapter
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice1TableAdapter
        Dim ObjItem As New ClsTransaksiSalesInvoiceItem
        Dim UrutanFakturPajak As String = ""
        Dim PrimarySO As Double = -1
        Dim PrimaryDO As Double = -1
        Dim PrimaryStockOut As Double = -1
        Dim CustIndentifier As String = ""

        Dim rowDGV As DataGridViewRow
        Dim ObjCust As New Core_Module.clsMasterCustomer
        Dim ObjSO As New ClsTransaksiSalesOrder
        Dim ObjDO As New ClsTransaksiDO
        Dim Lanjut As Boolean = True
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim NoUrut As Double = 0
        Dim Obj As New ClsTransaksiSalesInvoice_SettingTaxInvoice
        Dim PrimaryFakturPajak As Double = -1
        Dim NoFakturPajakMax As String = ""
        Dim DefPrecBase As Integer = 0

        Dim DPOriginal, DPBaseCurrency As Decimal
        'revisi: 2018-07-14: quick invoice ga pake dp; iwan
        'Dim DTA1 As New DseTransaksiSalesInvoiceTableAdapters.DataTable13TableAdapter
        'Dim TBL1 As New DseTransaksiSalesInvoice.DataTable13DataTable
        'Dim row1 As DseTransaksiSalesInvoice.DataTable13Row
        Dim PrimaryAkunDP As Double = -1
        Dim DefPrec As Integer = 0
        Dim vShipmentTax, vVatDipungut As Integer

        Dim CreatedBy_ As String = vUserCode
        Dim CreatedDate_ As Date = Now
        Dim ModifiedBy_ As String = vUserCode
        Dim ModifiedDate_ As Date = Now
        Dim TimeStamp_ As Date = Now
        Dim ObjSC As New Inventory_Module.ClsTransaksiStockCard
        Dim rowCustomInOutDummy As CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemRow

        Try
            If vCekApakahDataBenar = True Then
                If Me.ApakahDataBenar(True, vdgvItem, vFrm, vNoDO, -2) = False Then
                    Lanjut = False
                End If
            End If

            If Lanjut = True Then
                DefPrecBase = ObjCur.GetDecimal_DefaultCurrency
                'simpan data so, do, stock out
                'CustIndentifier = ObjCust.GetIdentifierAdd(Me.PrimaryCustomer)
                CustIndentifier = vFrm.txtCustomerAddIdentifier.Text
                'simpan data so
                PrimarySO = Me.SimpanBaruSO(vUserCode, vdgvItem, vFrm, CustIndentifier, ObjSO, Me.mPrimarySalesPerson, Me.mTaxExchangeRate, _
                vCekApakahDataBenar, vSetSOB, vFromCutOff, vNoSO)

                If PrimarySO = -1 Then
                    Return False
                End If
                'simpan data do
                PrimaryDO = Me.SimpanBaruDO(vUserCode, vdgvItem, vFrm, PrimarySO, CustIndentifier, ObjDO, ObjSO.NoTransaksi, vNoDO, _
                vCekApakahDataBenar, vSetSOB, vFromCutOff)
                If PrimaryDO = -1 Then
                    ObjSO.Hapus(PrimarySO, vUserCode, False)
                    Return False
                End If

                For Each rowCustomInOutDummy In TblCustomInOutDummy.Rows
                    rowCustomInOutDummy.Primary_PO_Purchase = PrimarySO
                Next

                'simpan data stock out sales
                PrimaryStockOut = Me.SimpanStockOut(vUserCode, vdgvItem, vFrm, PrimarySO, _
                CustIndentifier, PrimaryDO, TblCustomInOutDummy, vCekApakahDataBenar, vKoreksiStock, vSetSOB, vFromCutOff, vNoOutSales)
                If PrimaryStockOut = -1 Then
                    ObjDO.Hapus1(PrimaryDO, vUserCode, False)
                    ObjSO.Hapus(PrimarySO, vUserCode, False)
                    Return False
                End If

                '2025 Core Tax Tahap 2.2
                'If vFromCutOff = False Then
                '    If Me.mTanggalTransaksi >= Obj.BatasStartDate Then
                '        Me.mUrutanFakturPajak = Obj.GetUrutan(Me.mNoFakturPajak, Me.mTanggalTransaksi)
                '    Else
                '        If Me.mJenisFakturPajak = EnumJenisFakturPajak.FakturPajakStandar Then
                '            Me.mUrutanFakturPajak = Me.NoFakturPajak.Substring(4)
                '        Else
                '            Me.mUrutanFakturPajak = 0
                '            Me.mNoSeriFakturPajak = ""
                '            Me.mNoFakturPajak = ""
                '        End If
                '    End If
                'End If

                If Me.mShipmentWithTax = True Then
                    vShipmentTax = 1
                Else
                    vShipmentTax = 0
                End If
                If Me.mVATDipungut = True Then
                    vVatDipungut = 1
                Else
                    vVatDipungut = 0
                End If

                If vFromCutOff = True Then
                    CreatedBy_ = Me.mCreatedBy
                    CreatedDate_ = Me.mCreatedDate
                    ModifiedBy_ = Me.mModifiedBy
                    ModifiedDate_ = Me.mModifiedDate
                    TimeStamp_ = Me.mTimestamp
                End If

                'insert query
                NoUrut = Me.ObjNumbering.GetUrutan(Me.mNoTransaksi, Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Invoice, _
                Me.mTanggalTransaksi)
                If DTA2.InsertQuery2(Me.mNoTransaksi, NoUrut, Me.mUrutanFakturPajak, _
                Me.mTipeInvoice, Me.mTanggalTransaksi, Me.mPrimaryCust, Me.mPrimaryCurrency, Me.mCalculatesBy, _
                Me.mSubTotalB4DP, Me.mDownPayment, Me.mSubTotal, Me.mDutyStamp, Me.mUseTax, Me.mTaxPercentage, _
                Me.mTaxExchangeRate, Me.mTaxAmount, Me.mTotal, Me.mExchangeRateToUSD, Me.mExhangeRateToIDR, _
                Me.mNoSeriFakturPajak, Me.mNoFakturPajak, Me.mNoNPWP, Me.mIdentifier_CustAddress, _
                Me.mRemarks, CreatedBy_, CreatedDate_, ModifiedBy_, ModifiedDate_, TimeStamp_, Me.mJenisFakturPajak, Me.mPaymentTerm, _
                Me.mPrimarySO4DP, Me.mNoFakturPajakSederhana, Me.mNamaSigned, Me.mJabatanSigned, 0, Me.mRevisi, _
                Me.mAlamatPajak1, Me.mAlamatPajak2, Me.mExchangeRate, Me.mOngkosKirim, vVatDipungut, vShipmentTax, _
                Me.mPrimarySalesPerson, Me.mDiscString, Me.mDiscAmount, Me.mEFak_NamaJalan, Me.mEFak_Blok, _
                Me.mEFak_No, Me.mEFak_RT, Me.mEFak_RW, Me.mEFak_Kecamatan, Me.mEFak_Kelurahan, Me.mEFak_Kabupaten, _
                Me.mEFak_Propinsi, Me.mEFak_KodePos, Me.mEFak_Telepon, Me.mEFak_KodeJenisTransaksiPajak, _
                Me.mPrimaryAccount, Me.mPrimaryCurrTax, Me.mPrimaryAccountDP, Me.mOtherCharge, Me.mOptionNIKPASS, Me.mNoNikPass, Me.mTaxAmount1, _
                Me.PrimaryAddInfo, Me.PrimaryFacilityStamp, Me.TaxRevised, Me.TaxType, Me.SPPBDate, Me.mDPPNilaiLain) > 0 Then '//TR : Tambahan baru Nik/Pas
                    'update data yang berubah setelah insert query
                    Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mNoTransaksi).Value
                    Me.mNoUrut = NoUrut
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode
                    'update status quick invoice
                    DTA.UpdateQuery_SetQuickInvoice(1, Me.mPrimary)
                    'set so yang dibuat sebagai quick so
                    If PrimarySO <> -1 Then
                        ObjSO.SetAsQuickSO(Me.mPrimary, PrimarySO)
                    End If
                    ObjSC.UpdateDocNo(Inventory_Module.ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockOut_Sales, PrimaryStockOut, Me.mNoTransaksi)

                    '2025 Core Tax Tahap 2.2
                    'If vFromCutOff = False Then
                    '    'update data berkaitan dengan no faktur pajak
                    '    If Me.mJenisFakturPajak = EnumJenisFakturPajak.FakturPajakStandar And Me.UseTax = True Then
                    '        While DTA.ScalarQuery_GetJmlByNoFakturPajak(Me.mNoFakturPajak).Value > 1
                    '            If DTA.ScalarQuery_GetPrimaryMaxByNoFakturPajak(Me.mNoFakturPajak).Value = _
                    '            Me.mPrimary Then
                    '                'get no baru
                    '                Me.mNoFakturPajak = Me.GetNoFakturPajakStandarBaru(Me.mTanggalTransaksi)
                    '                'update no-nya
                    '                If Me.mTanggalTransaksi >= Obj.BatasStartDate Then
                    '                    Me.mUrutanFakturPajak = Obj.GetUrutan(Me.mNoFakturPajak, Me.mTanggalTransaksi)
                    '                Else
                    '                    If Me.mJenisFakturPajak = EnumJenisFakturPajak.FakturPajakStandar Then
                    '                        Me.mUrutanFakturPajak = Me.NoFakturPajak.Substring(4)
                    '                    Else
                    '                        Me.mUrutanFakturPajak = 0
                    '                        Me.mNoSeriFakturPajak = ""
                    '                        Me.mNoFakturPajak = ""
                    '                    End If
                    '                End If

                    '                DTA.UpdateQuery_UpdateNoFakturPajak(Me.mUrutanFakturPajak, Me.mNoFakturPajak, _
                    '                Me.mPrimary)
                    '            End If
                    '        End While
                    '    End If
                    'End If

                    'update data yang berkaitan dengan item exchange rate
                    If Me.mCalculatesBy = EnumCalculatesBy.Invoice Then
                        vDGVItemExchangeRate.Rows(0).Cells(5).Value = Me.mPrimary
                        vDGVItemExchangeRate.Rows(0).Cells(1).Value = Me.mNoTransaksi
                    End If

                    'simpan data item baru
                    For Each rowDGV In vdgvItem.Rows
                        rowDGV.Cells(vFrm.ColPrimarySO.Index).Value = PrimarySO
                        rowDGV.Cells(vFrm.ColPrimaryDO.Index).Value = PrimaryDO
                    Next

                    'udpate tblcustom
                    For Each row As CustomSMESS.DseTransaksiSalesInvoiceQuick_AddByCustomProduct.tbltrans_custom_salesinvoice_item1Row In vTblCustom.Rows
                        row.Primary_Transaksi = Me.mPrimary
                        row.Primary_SO = PrimarySO
                        row.Primary_DO = PrimaryDO
                    Next
                    ObjItem.Simpan(vdgvItem, vDGVItemExchangeRate, vdgvItemDP, Me.mPrimary, vDGVItemSR, vFrm, vTblCustom)

                    'simpan berkaitan other charge
                    If vFrm.TBLHeaderCost.Rows.Count > 0 Then
                        For Each vFrm.rowHeaderCost In vFrm.TBLHeaderCost.Rows
                            With vFrm.rowHeaderCost
                                DTAHeaderCost.InsertQuery(Me.mPrimary, .ETD_Date, .ETD_Port, .ETA_Date, .ETA_Port, .Total_Nett_Weight, _
                                .Total_Gross_Weight, .primary_weight_unit, .Total_Nett_Volume, .Total_Gross_Volume, .primary_volume_unit, _
                                .primary_cust_ship_to, .identifier_cust_ship_to, .shipping_mark, .transportation_mode, .invoice_type, _
                                .var1_name, .var1_value, .var2_name, .var2_value, .var3_name, .var3_value, .var4_name, .var4_value, _
                                .var5_name, .var5_value)
                            End With
                        Next
                    End If

                    If vFrm.TBLitemCost.Rows.Count > 0 Then
                        For Each vFrm.rowitemCost In vFrm.TBLitemCost.Rows
                            With vFrm.rowitemCost
                                DTAItemCost.InsertQuery(Me.mPrimary, .Cost_Description, .Cost, .Remarks, .Primary_Akun)
                            End With
                        Next
                    End If


                    'cek berkaitan nomor sama untuk beda bulan -> obselete as off 2017-01-31 by iwan
                    'If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Invoice) = True Then
                    '    Me.CekNomor1(DTA)
                    'End If

                    'simpan data sales amount
                    Me.SimpanDataSalesAmount(Me.mCalculatesBy, Me.mPrimary, _
                    Math.Round(Me.mSubTotalB4DP * Me.mTaxExchangeRate, ObjCur.GetDecimal_DefaultCurrency, MidpointRounding.AwayFromZero))

                    Dim objCustomSalesInv As New CustomSMESS.ClsTransaksiSalesInvoice
                    Dim objIdentifier As New CustomSMESS.ClsIdentifier
                    If objIdentifier.GetIdentifier() = "smess_binaya" Then
                        'Me.SimpanDataSalesAmount_Custom(Me.mCalculatesBy, Me.mPrimary, _
                        '                   Math.Round(Me.mSubTotalB4DP * Me.mTaxExchangeRate, ObjCur.GetDecimal_DefaultCurrency, MidpointRounding.AwayFromZero))
                        objCustomSalesInv.SimpanDataSalesAmount_Custom(Me.mCalculatesBy, Me.mPrimary, _
                                         Math.Round(Me.mSubTotalB4DP * Me.mTaxExchangeRate, ObjCur.GetDecimal_DefaultCurrency, MidpointRounding.AwayFromZero))

                    End If


                    'hitung DP
                    'revisi: 2018-07-14: quick invoice ga pake dp; iwan
                    DPOriginal = 0
                    DPBaseCurrency = 0
                    'DTA1.Fill(TBL1, Me.mPrimary)
                    'For Each row1 In TBL1.Rows
                    '    DPOriginal += ObjCD.DigitGrouping(row1.Sub_Total_Before_DP, ObjCur.GetDefaultPrecission(Me.mPrimaryCurrency))
                    '    DPBaseCurrency += ObjCD.DigitGrouping(row1.Sub_Total_Before_DP * row1.Tax_Exchange_Rate, DefPrecBase)
                    '    PrimaryAkunDP = Me.GetPrimaryAccountDP(row1.Primary_Invoice_DP)
                    'Next

                    '**************************************************************************************************************************************
                    '**************************************************************************************************************************************
                    'PERUBAHAN SISTEM PERHITUNGAN JURNAL ERROR NO 50 TANGGAL 11 JULI 2016 -- PERBAIKAN 28 JULI 2016
                    'SEBELUMNYA NILAI SALES RUPIAH = NILAI SALES USD * EXCHANGE RATE
                    'SEKARANG NILAINYA DICEK DULU APAKAH ADA PENJUMLAHAN TERHADAP NILAI SALES (CONTOH TAX USD, SHIPMENT DLL),
                    'HASILNYA DITOTAL LALU DIKURANGI SEMUA PENJUMLAHAN, SISANYA DIANGGAP SEBAGAI SALES
                    'HAL INI DILAKUKAN UTK MENCEGAH SELISIH ANTARA SALES DI INVOICE JURNAL DENGAN SALES DI RECEIVEMENT JOURNAL


                    Dim SalesRupiah As Decimal = 0
                    Dim TotalSales As Decimal = 0
                    Dim AmountTanpaSales As Decimal = 0
                    Dim SalesPlusDP As Decimal = 0


                    SalesPlusDP = Me.ObjCD.DigitGrouping(Me.mTotal + DPOriginal, ObjCur.GetDefaultPrecission(Me.mPrimaryCurrency))
                    TotalSales = Me.ObjCD.DigitGrouping(SalesPlusDP * Me.mTaxExchangeRate, DefPrecBase)

                    AmountTanpaSales += Me.ObjCD.DigitGrouping(Me.mOngkosKirim * Me.mTaxExchangeRate, DefPrecBase)
                    AmountTanpaSales += Me.ObjCD.DigitGrouping(Me.mDiscAmount * Me.mTaxExchangeRate, DefPrecBase)
                    AmountTanpaSales += Me.ObjCD.DigitGrouping(Me.mOtherCharge * Me.mTaxExchangeRate, DefPrecBase)
                    AmountTanpaSales -= Me.ObjCD.DigitGrouping(DPOriginal * Me.mTaxExchangeRate, DefPrecBase)
                    If Me.mPrimaryCurrTax = Me.mPrimaryCurrency Then
                        AmountTanpaSales += Me.ObjCD.DigitGrouping(Me.mTaxAmount * Me.mTaxExchangeRate, DefPrecBase)
                        AmountTanpaSales += Me.ObjCD.DigitGrouping(Me.mDutyStamp * Me.mTaxExchangeRate, DefPrecBase)
                    End If
                    'AmountTanpaSalesRp = Me.ObjCD.DigitGrouping((AmountTanpaSales * Me.mExchangeRate), DefPrecBase)
                    SalesRupiah = Me.ObjCD.DigitGrouping(TotalSales - AmountTanpaSales, DefPrecBase)


                    '**************************************************************************************************************************************

                    'simpan jurnal otomatis
                    DefPrecBase = ObjCur.GetDecimal_DefaultCurrency
                    Me.TransferJurnal_Invoice2(Me.mPrimary, Me.mPrimaryCust, _
                    SalesRupiah, _
                    Me.ObjCD.DigitGrouping(Me.mTaxAmount * Me.mTaxExchangeRate, DefPrecBase), _
                    Me.ObjCD.DigitGrouping(Me.mDutyStamp * Me.mTaxExchangeRate, DefPrecBase), _
                    Me.mNoUrut, Me.mTanggalTransaksi, Me.mNoTransaksi, Me.mModifiedBy, CDbl(Me.mPaymentTerm), Me.mTaxExchangeRate, _
                    Me.mSubTotalB4DP, Me.mPrimaryCurrency, Me.ObjCD.DigitGrouping(Me.mOngkosKirim * Me.mTaxExchangeRate, DefPrecBase), _
                    DPOriginal, DPBaseCurrency, Me.mOngkosKirim, Me.mDiscAmount, _
                    Me.ObjCD.DigitGrouping(Me.mDiscAmount * Me.mExchangeRate, DefPrecBase), _
                    Me.mPrimaryCurrTax, Me.mPrimaryAccount, Me.mDutyStamp, Me.mTaxAmount, PrimaryAkunDP, _
                    Me.ObjCD.DigitGrouping(Me.mOtherCharge * Me.mTaxExchangeRate, DefPrecBase), Me.mOtherCharge, vFrm.TBLitemCost, vFrm.rowitemCost)


                    'set sob menjadi no
                    ObjSO.SetOrderBalance_QuickInvoice_Nol(PrimarySO)

                    '2025 Core Tax Tahap 2.2
                    ''simpan no urut max untuk numbering
                    'If Obj.BatasStartDate <= Me.mTanggalTransaksi Then
                    '    If Me.mJenisFakturPajak = EnumJenisFakturPajak.FakturPajakStandar Then
                    '        DTA.UpdateQuery_PrimaryFakturPajak(Obj.GetPrimary(Me.mTanggalTransaksi, Me.mNoFakturPajak), Me.mPrimary)
                    '        If vFromCutOff = False Then
                    '            PrimaryFakturPajak = DTA.ScalarQuery_GetPrimaryFakturPajak(Me.mPrimary)
                    '            NoFakturPajakMax = Me.GetNoFakturPajakMax(PrimaryFakturPajak, Me.mPrimary)

                    '            If Obj.GetUrutan(NoFakturPajakMax, Me.mTanggalTransaksi) < Obj.GetUrutan(Me.mNoFakturPajak, Me.mTanggalTransaksi) Then
                    '                Obj.SetNoUrutMaksimal(Me.mTanggalTransaksi, Me.mNoFakturPajak)
                    '            Else
                    '                Obj.SetNoUrutMaksimal(PrimaryFakturPajak, NoFakturPajakMax)
                    '            End If
                    '        End If
                    '    End If
                    'End If

                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mNoTransaksi + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Invoice) = True Then
                    Me.mNoTransaksi = Me.GetNoBaru1(Me.mTanggalTransaksi)
                    Return Me.SimpanBaru1(vdgvItem, vdgvItemDP, vDGVItemSR, vDGVItemExchangeRate, vUserCode, vFrm, vCekApakahDataBenar, vNoDO, vTblCustom, TblCustomInOutDummy, vSetSOB, vKoreksiStock, vFromCutOff) 'new
                    'Return Me.SimpanBaru1(vdgvItem, vdgvItemDP, vDGVItemSR, vDGVItemExchangeRate, vUserCode, vFrm, vCekApakahDataBenar, vNoDO, vTblCustom, vSetSOB, vKoreksiStock, vFromCutOff)
                Else
                    MsgBox("Invoice number duplication, please change invoice number.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            DTA2.Dispose()
            ObjItem = Nothing
            ObjCust = Nothing
            ObjSO = Nothing
            ObjDO = Nothing
            ObjCur = Nothing
            ObjSC = Nothing
            'DTA1.Dispose()
            'TBL1.Dispose()
        End Try
    End Function


    Public Overridable Function SimpanBaru(ByVal vUserCode As String, _
    Optional ByVal vCekDataBenar As Boolean = True, Optional ByVal vApakahSimpanJurnal As Boolean = True, _
    Optional ByVal vFromCutOff As Boolean = False) As Boolean
        Dim DTA3 As New DseTransaksiSalesInvoice_TaxRateTableAdapters.tbltrans_salesinvoice_NEWTableAdapter
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice1TableAdapter
        Dim ObjItem As New ClsTransaksiSalesInvoiceItem
        Dim JmlFakturPajak As Double = 0
        Dim NoUrut As Double = 0
        Dim Obj As New ClsTransaksiSalesInvoice_SettingTaxInvoice
        Dim PrimaryFakturPajak As Double = -1
        Dim NoFakturPajakMax As String = ""
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim vShipmentTax, vVatDipungut As Integer

        Dim CreatedBy_ As String = vUserCode
        Dim CreatedDate_ As Date = Now
        Dim ModifiedBy_ As String = vUserCode
        Dim ModifiedDate_ As Date = Now
        Dim TimeStamp_ As Date = Now

        Try
            If Me.ApakahDataBenar_DP(True, Me.mTanggalTransaksi, vCekDataBenar, vFromCutOff) = True Then
                '2025 Core Tax Tahap 2.2
                'If vFromCutOff = False Then
                '    'MsgBox("Tgl Transaksi: " + Format(Me.mTanggalTransaksi, "dd MMM yyyy") + " - Batas start date: " + Format(Obj.BatasStartDate, "dd MMM yyyy"), MsgBoxStyle.Information, "information")
                '    '"Batas start date: "+format(Obj.BatasStartDate, "dd MMM yyyy"))
                '    If Me.mTanggalTransaksi >= Obj.BatasStartDate Then
                '        Me.mUrutanFakturPajak = Obj.GetUrutan(Me.mNoFakturPajak, Me.mTanggalTransaksi)
                '    Else
                '        If Me.mJenisFakturPajak = EnumJenisFakturPajak.FakturPajakStandar Then
                '            Me.mUrutanFakturPajak = Me.NoFakturPajak.Substring(4)
                '        Else
                '            Me.mUrutanFakturPajak = 0
                '            Me.mNoSeriFakturPajak = ""
                '            Me.mNoFakturPajak = ""
                '        End If
                '    End If
                'End If

                If Me.mShipmentWithTax = True Then
                    vShipmentTax = 1
                Else
                    vShipmentTax = 0
                End If
                If Me.mVATDipungut = True Then
                    vVatDipungut = 1
                Else
                    vVatDipungut = 0
                End If

                NoUrut = Me.ObjNumbering.GetUrutan(Me.mNoTransaksi, Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Invoice, _
                               Me.mTanggalTransaksi)

                If vFromCutOff = True Then
                    CreatedBy_ = Me.mCreatedBy
                    CreatedDate_ = Me.mCreatedDate
                    ModifiedBy_ = Me.mModifiedBy
                    ModifiedDate_ = Me.mModifiedDate
                    TimeStamp_ = Me.mTimestamp
                End If

                'insert query
                If DTA3.InsertQuery(Me.mNoTransaksi, NoUrut, Me.mUrutanFakturPajak, _
                Me.mTipeInvoice, Me.mTanggalTransaksi, Me.mPrimaryCust, Me.mPrimaryCurrency, Me.mCalculatesBy, _
                Me.mSubTotalB4DP, Me.mDownPayment, Me.mSubTotal, Me.mDutyStamp, Me.mUseTax, Me.mTaxPercentage, _
                Me.mTaxExchangeRate, Me.mTaxAmount, Me.mTotal, Me.mExchangeRateToUSD, Me.mExhangeRateToIDR, _
                Me.mNoSeriFakturPajak, Me.mNoFakturPajak, Me.mNoNPWP, Me.mIdentifier_CustAddress, _
                Me.mRemarks, CreatedBy_, CreatedDate_, ModifiedBy_, ModifiedDate_, TimeStamp_, Me.mJenisFakturPajak, Me.mPaymentTerm, Me.mPrimarySO4DP, _
                Me.mNoFakturPajakSederhana, Me.mNamaSigned, Me.mJabatanSigned, Me.mAlamatPajak1, Me.mAlamatPajak2, _
                Me.mExchangeRate, Me.mOngkosKirim, vVatDipungut, vShipmentTax, Me.mPrimarySalesPerson, Me.mDiscString, _
                Me.mDiscAmount, Me.mEFak_NamaJalan, Me.mEFak_Blok, Me.mEFak_No, Me.mEFak_RT, Me.mEFak_RW, Me.mEFak_Kecamatan, _
                Me.mEFak_Kelurahan, Me.mEFak_Kabupaten, Me.mEFak_Propinsi, _
                Me.mEFak_KodePos, Me.mEFak_Telepon, Me.mEFak_KodeJenisTransaksiPajak, _
                Me.mPrimaryAccount, Me.mPrimaryCurrTax, Me.mPrimaryAccountDP, Me.mOtherCharge, Me.mOptionNIKPASS, Me.mNoNikPass, Me.mTaxAmount1, _
                Me.PrimaryAddInfo, Me.PrimaryFacilityStamp, Me.TaxRevised, Me.TaxType, Me.SPPBDate, Me.mDPPNilaiLain) > 0 Then '//TR : Tambahan baru Nik/Pas
                    'update data yang berubah setelah insert query
                    Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mNoTransaksi).Value
                    Me.mNoUrut = NoUrut
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    '2025 Core Tax Tahap 2.2
                    'If vFromCutOff = False Then
                    '    'update data berkaitan dengan no faktur pajak
                    '    If Me.mJenisFakturPajak = EnumJenisFakturPajak.FakturPajakStandar And Me.UseTax = True Then
                    '        While DTA.ScalarQuery_GetJmlByNoFakturPajak(Me.mNoFakturPajak).Value > 1
                    '            If DTA.ScalarQuery_GetPrimaryMaxByNoFakturPajak(Me.mNoFakturPajak).Value = _
                    '            Me.mPrimary Then
                    '                'get no baru
                    '                Me.mNoFakturPajak = Me.GetNoFakturPajakStandarBaru(Me.mTanggalTransaksi)
                    '                'update no-nya
                    '                If Me.mTanggalTransaksi >= Obj.BatasStartDate Then
                    '                    Me.mUrutanFakturPajak = Obj.GetUrutan(Me.mNoFakturPajak, Me.mTanggalTransaksi)
                    '                Else
                    '                    If Me.mJenisFakturPajak = EnumJenisFakturPajak.FakturPajakStandar Then
                    '                        Me.mUrutanFakturPajak = Me.NoFakturPajak.Substring(4)
                    '                    Else
                    '                        Me.mUrutanFakturPajak = 0
                    '                        Me.mNoSeriFakturPajak = ""
                    '                        Me.mNoFakturPajak = ""
                    '                    End If
                    '                End If

                    '                DTA.UpdateQuery_UpdateNoFakturPajak(Me.mUrutanFakturPajak, Me.mNoFakturPajak, _
                    '                Me.mPrimary)
                    '            End If
                    '        End While
                    '    End If

                    'update data berkaitan dengan cust po no
                    '    If Me.mJenisFakturPajak = EnumJenisFakturPajak.FakturPajakStandar And Me.UseTax = True Then
                    '        While DTA.ScalarQuery_GetJmlByNoFakturPajak(Me.mNoFakturPajak).Value > 1
                    '            If DTA.ScalarQuery_GetPrimaryMaxByNoFakturPajak(Me.mNoFakturPajak).Value = _
                    '            Me.mPrimary Then
                    '                'get no baru
                    '                Me.mNoFakturPajak = Me.GetNoFakturPajakStandarBaru(Me.mTanggalTransaksi)
                    '                'update no-nya
                    '                Me.mUrutanFakturPajak = Me.NoFakturPajak.Substring(4)
                    '                DTA.UpdateQuery_UpdateNoFakturPajak(Me.mUrutanFakturPajak, Me.mNoFakturPajak, _
                    '                Me.mPrimary)
                    '            End If
                    '        End While
                    '    End If
                    'End If

                    'cek berkaitan nomor sama untuk beda bulan -> obselete as off 2017-01-31 by iwan
                    'If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Invoice) = True Then
                    '    Me.CekNomor1(DTA)
                    'End If


                    '**************************************************************************************************************************************
                    '**************************************************************************************************************************************
                    'PERUBAHAN SISTEM PERHITUNGAN JURNAL ERROR NO 50 TANGGAL 11 JULI 2016 -- PERBAIKAN 28 JULI 2016
                    'SEBELUMNYA NILAI SALES RUPIAH = NILAI SALES USD * EXCHANGE RATE
                    'SEKARANG NILAINYA DICEK DULU APAKAH ADA PENJUMLAHAN TERHADAP NILAI SALES (CONTOH TAX USD, SHIPMENT DLL),
                    'HASILNYA DITOTAL LALU DIKURANGI SEMUA PENJUMLAHAN, SISANYA DIANGGAP SEBAGAI SALES
                    'HAL INI DILAKUKAN UTK MENCEGAH SELISIH ANTARA SALES DI INVOICE JURNAL DENGAN SALES DI RECEIVEMENT JOURNAL


                    Dim DefPrecBase As Integer = 0
                    Dim SalesRupiah As Decimal = 0
                    Dim TotalSales As Decimal = 0
                    Dim AmountTanpaSales As Decimal = 0

                    DefPrecBase = ObjCur.GetDecimal_DefaultCurrency

                    TotalSales = Me.ObjCD.DigitGrouping(Me.mTotal * Me.mTaxExchangeRate, DefPrecBase)

                    AmountTanpaSales += Me.ObjCD.DigitGrouping(Me.mOngkosKirim * Me.mTaxExchangeRate, DefPrecBase)
                    AmountTanpaSales += Me.ObjCD.DigitGrouping(Me.mDiscAmount * Me.mTaxExchangeRate, DefPrecBase)
                    AmountTanpaSales += Me.ObjCD.DigitGrouping(Me.mOtherCharge * Me.mTaxExchangeRate, DefPrecBase)
                    If Me.mPrimaryCurrTax = Me.mPrimaryCurrency Then
                        AmountTanpaSales += Me.ObjCD.DigitGrouping(Me.mTaxAmount * Me.mTaxExchangeRate, DefPrecBase)
                        AmountTanpaSales += Me.ObjCD.DigitGrouping(Me.mDutyStamp * Me.mTaxExchangeRate, DefPrecBase)
                    End If
                    'AmountTanpaSalesRp = Me.ObjCD.DigitGrouping((AmountTanpaSales * Me.mExchangeRate), DefPrecBase)
                    SalesRupiah = Me.ObjCD.DigitGrouping(TotalSales - AmountTanpaSales, DefPrecBase)


                    '**************************************************************************************************************************************


                    'simpan jurnal otomatis
                    Me.TransferJurnal_Invoice_DP2(Me.mPrimary, Me.mPrimaryCust, _
                    SalesRupiah, _
                    Me.ObjCD.DigitGrouping(Me.mTaxAmount * Me.mTaxExchangeRate, DefPrecBase), _
                    Me.ObjCD.DigitGrouping(Me.mDutyStamp * Me.mTaxExchangeRate, DefPrecBase), _
                    Me.mNoUrut, Me.mTanggalTransaksi, Me.mNoTransaksi, Me.mModifiedBy, _
                    CDbl(Me.mPaymentTerm), Me.mPrimaryCurrency, Me.mSubTotalB4DP, Me.mTaxExchangeRate, _
                    Me.mPrimaryCurrTax, Me.mPrimaryAccount, Me.mPrimaryAccountDP, Me.mTaxAmount, Me.mDutyStamp)

                    '2025 Core Tax Tahap 2.2
                    ''simpan no urut max untuk numbering
                    'If Obj.BatasStartDate <= Me.mTanggalTransaksi Then
                    '    If Me.mJenisFakturPajak = EnumJenisFakturPajak.FakturPajakStandar Then
                    '        DTA.UpdateQuery_PrimaryFakturPajak(Obj.GetPrimary(Me.mTanggalTransaksi, Me.mNoFakturPajak), Me.mPrimary)
                    '        If vFromCutOff = False Then
                    '            PrimaryFakturPajak = DTA.ScalarQuery_GetPrimaryFakturPajak(Me.mPrimary)
                    '            NoFakturPajakMax = Me.GetNoFakturPajakMax(PrimaryFakturPajak, Me.mPrimary)

                    '            If Obj.GetUrutan(NoFakturPajakMax, Me.mTanggalTransaksi) < Obj.GetUrutan(Me.mNoFakturPajak, Me.mTanggalTransaksi) Then
                    '                Obj.SetNoUrutMaksimal(Me.mTanggalTransaksi, Me.mNoFakturPajak)
                    '            Else
                    '                Obj.SetNoUrutMaksimal(PrimaryFakturPajak, NoFakturPajakMax)
                    '            End If
                    '        End If
                    '    End If
                    'End If

                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mNoTransaksi + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Invoice) = True Then
                    Me.mNoTransaksi = Me.GetNoBaru1(Me.mTanggalTransaksi)
                    Return Me.SimpanBaru(vUserCode, vCekDataBenar, vApakahSimpanJurnal, vFromCutOff)
                Else
                    MsgBox("Invoice number duplication, please change invoice number.", MsgBoxStyle.Information, "Information")
                    Return True
                End If
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
            ObjCur = Nothing
        End Try
    End Function

    'Hanya berlaku jika format numbering untuk SO, DO, Out Sales
    Public Sub GetNumbering_SO_DO_OutSales(ByVal vPrimaryInvoice As Double, ByVal vInvoiceDate As Date, _
    ByRef vSOno As String, ByRef vDOno As String, ByRef vOutSalesNo As String)
        Dim dta As New DseTransaksiSalesInvoiceTableAdapters.DataTable5TableAdapter
        Dim Tbl As New DseTransaksiSalesInvoice.DataTable5DataTable
        Dim row As DseTransaksiSalesInvoice.DataTable5Row
        Dim ObjSO As New ClsTransaksiSalesOrder
        Dim ObjDO As New ClsTransaksiDO
        Dim ObjInOut As New Inventory_Module.clsTransaksiStockInOut

        Try
            'vSOno = ""
            'vDOno = ""
            'vOutSalesNo = ""

            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.Fill(Tbl, vPrimaryInvoice) > 0 Then
                row = Tbl.Rows(0)
                With row
                    'SO NUMBERING
                    If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.SO) = True Then
                        If Me.ObjNumbering.GetResetType(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.SO) = Core_Module.ClsMasterTransactionNumbering.EnumReset.Monthly Then
                            'monthly
                            If Format(.so_date, "MMM yyyy") = Format(vInvoiceDate, "MMM yyyy") Then
                                vSOno = .so_no
                            End If
                        Else
                            'yearly
                            If Format(.so_date, "yyyy") = Format(vInvoiceDate, "yyyy") Then
                                If Format(.so_date, "MMM") = Format(vInvoiceDate, "MMM") Then
                                    'jika sama tahun dan bulannya maka keep no lama
                                    vSOno = .so_no
                                Else
                                    'jika sama tahun tapi bulannya beda maka ganti bulannya saja dan nomor nya tetap di keep
                                    vSOno = Me.ObjNumbering.GetNomorTransaksiBaru1(vInvoiceDate, .so_urutan, Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.SO)
                                End If
                            Else
                                'jika tahun dan bulannya tidak sama maka get no baru
                                vSOno = ObjSO.GetNoBaru1(vInvoiceDate)
                            End If
                        End If
                    End If

                    'DO NUMBERING
                    If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._DO) = True Then
                        If Me.ObjNumbering.GetResetType(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._DO) = Core_Module.ClsMasterTransactionNumbering.EnumReset.Monthly Then
                            'monthly
                            If Format(.do_date, "MMM yyyy") = Format(vInvoiceDate, "MMM yyyy") Then
                                vDOno = .do_no
                            End If
                        Else
                            'yearly
                            If Format(.do_date, "yyyy") = Format(vInvoiceDate, "yyyy") Then
                                If Format(.do_date, "MMM") = Format(vInvoiceDate, "MMM") Then
                                    'jika sama tahun dan bulannya maka keep no lama
                                    vDOno = .do_no
                                Else
                                    'jika sama tahun tapi bulannya beda maka ganti bulannya saja dan nomor nya tetap di keep
                                    vDOno = Me.ObjNumbering.GetNomorTransaksiBaru1(vInvoiceDate, .do_urutan, Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._DO)
                                End If
                            Else
                                'jika tahun dan bulannya tidak sama maka get no baru
                                vDOno = ObjDO.GetNoBaru1(vInvoiceDate)
                            End If
                        End If
                    End If

                    '11 Juli 2019 /kodisi saat ini untuk inout numbering masih di hardcode selalu auto numbering dan reset number pertahun di coding, 
                    'jadi coding agak beda dengan SO, DO
                    If Format(.out_sales_date, "yyyy") = Format(vInvoiceDate, "yyyy") Then
                        If Format(.out_sales_date, "MMM") = Format(vInvoiceDate, "MMM") Then
                            'jika sama tahun dan bulannya maka keep no lama
                            vOutSalesNo = .out_sales_no
                        Else
                            'jika sama tahun tapi bulannya beda maka ganti bulannya saja dan nomor nya tetap di keep
                            vOutSalesNo = ObjInOut.GetNoBaru_20190711(Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Sales, vInvoiceDate, _
                            Inventory_Module.clsTransaksiStockInOut.EnumInOutAdjust.NoTipe, .out_sales_urutan)
                        End If
                    Else
                        'jika tahun dan bulannya tidak sama maka get no baru
                        vOutSalesNo = ObjInOut.GetNoBaru_20190711(Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Sales, vInvoiceDate)
                    End If
                End With
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Overridable Function SimpanPerubahan1(ByVal dgvItem As DataGridView, ByVal vdgvItemDP As DataGridView, _
   ByVal vDGVItemSR As DataGridView, ByVal vDGVItemExchangeRate As DataGridView, ByVal vUserCode As String, _
   ByVal vFrm As FrmTransaksiSalesInvoiceQuick, ByVal vNoDO As String, _
   ByVal vTblCustom As CustomSMESS.DseTransaksiSalesInvoiceQuick_AddByCustomProduct.tbltrans_custom_salesinvoice_item1DataTable, _
   ByVal TblCustomInOutDummy As CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable) As Boolean
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice1TableAdapter
        Dim ObjItem As New ClsTransaksiSalesInvoiceItem
        Dim UrutanFakturPajak As String = ""
        Dim ij As Double = 0
        Dim PrimaryMax As Double = -1
        Dim PrimaryStockOut As Double = -1
        Dim ObjDO As New ClsTransaksiDO

        Dim NoSO As String = ""
        Dim NoOutSales As String = ""

        Try
            If Me.ApakahBolehEdit(Me.mPrimary, vUserCode, True, True) = True And Me.ApakahDataBenar(False, dgvItem, _
            vFrm, vNoDO, Me.mPrimary) = True Then
                PrimaryStockOut = ObjDO.GetPrimaryStockOut(Me.GetPrimaryDO_QuickSales(Me.mPrimary))
                If Me.ApakahStockTidakMinus(dgvItem, vFrm, PrimaryStockOut) = True Then
                    '20190711 /fungsi keep numbering hanya di khususkan untuk SO, DO, Out Sales yang auto numbering saja
                    Me.GetNumbering_SO_DO_OutSales(Me.mPrimary, Me.TanggalTransaksi, NoSO, vNoDO, NoOutSales)
                    If Me.Hapus1(Me.mPrimary, vUserCode, True, True, True) = True Then
                        Return Me.SimpanBaru1(dgvItem, vdgvItemDP, vDGVItemSR, vDGVItemExchangeRate, vUserCode, vFrm, False, vNoDO, vTblCustom, TblCustomInOutDummy, True, True, False, NoSO, NoOutSales)
                    End If
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mNoTransaksi + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Invoice) = True Then
                    Me.mNoTransaksi = Me.GetNoBaru1(Me.mTanggalTransaksi)
                    Return Me.SimpanPerubahan1(dgvItem, vdgvItemDP, vDGVItemSR, vDGVItemExchangeRate, vUserCode, vFrm, vNoDO, vTblCustom, TblCustomInOutDummy)
                Else
                    MsgBox("Invoice number duplication, please change invoice number.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
            ObjDO = Nothing
        End Try
    End Function

    Public Overridable Function SimpanPerubahan(ByVal vUserCode As String) As Boolean
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice1TableAdapter
        Dim DTA2 As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_NewTableAdapter
        Dim ObjItem As New ClsTransaksiSalesInvoiceItem
        Dim ij As Double = 0
        Dim PrimaryMax As Double = -1
        Dim NoUrut As Double = 0
        Dim Obj As New ClsTransaksiSalesInvoice_SettingTaxInvoice
        Dim PrimaryFakturPajak As Double = -1
        Dim NoFakturPajakMax As String = ""
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim DefPrecBase As Integer = 0
        Dim vShipmentTax, vVatDipungut As Integer

        Try
            If Me.ApakahDataBenar_DP(False, Me.mTanggalTransaksi) = True Then
                DefPrecBase = ObjCur.GetDecimal_DefaultCurrency

                '2025 Core Tax Tahap 2.2
                'If Me.mTanggalTransaksi >= Obj.BatasStartDate Then
                '    Me.mUrutanFakturPajak = Obj.GetUrutan(Me.mNoFakturPajak, Me.mTanggalTransaksi)
                'Else
                '    If Me.mJenisFakturPajak = EnumJenisFakturPajak.FakturPajakStandar Then
                '        Me.mUrutanFakturPajak = Me.NoFakturPajak.Substring(4)
                '    Else
                '        Me.mUrutanFakturPajak = 0
                '        Me.mNoSeriFakturPajak = ""
                '        Me.mNoFakturPajak = ""
                '    End If
                'End If
                If Me.mShipmentWithTax = True Then
                    vShipmentTax = 1
                Else
                    vShipmentTax = 0
                End If
                If Me.mVATDipungut = True Then
                    vVatDipungut = 1
                Else
                    vVatDipungut = 0
                End If

                'hitung level otorisasi yang dibutuhkan
                NoUrut = Me.ObjNumbering.GetUrutan(Me.mNoTransaksi, Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Invoice, _
               Me.mTanggalTransaksi)
                If DTA2.UpdateQuery(Me.mNoTransaksi, NoUrut, Me.mUrutanFakturPajak, Me.mTipeInvoice, _
                Me.mTanggalTransaksi, Me.mPrimaryCust, Me.mPrimaryCurrency, Me.mCalculatesBy, Me.mSubTotalB4DP, _
                Me.mDownPayment, Me.mSubTotal, Me.mDutyStamp, Me.mUseTax, Me.mTaxPercentage, Me.mTaxExchangeRate, _
                Me.mTaxAmount, Me.mTotal, Me.mExchangeRateToUSD, Me.mExhangeRateToIDR, Me.mNoSeriFakturPajak, _
                Me.mNoFakturPajak, Me.mNoNPWP, Me.mIdentifier_CustAddress, Me.mRemarks, vUserCode, Now, Now, _
                Me.mJenisFakturPajak, Me.mPaymentTerm, Me.mPrimarySO4DP, Me.mNoFakturPajakSederhana, _
                Me.mNamaSigned, Me.mJabatanSigned, Me.mAlamatPajak1, Me.mAlamatPajak2, Me.mExchangeRate, Me.mOngkosKirim, _
                vVatDipungut, vShipmentTax, Me.mPrimarySalesPerson, Me.mDiscString, Me.mDiscAmount, Me.mEFak_NamaJalan, _
                Me.mEFak_Blok, Me.mEFak_No, Me.mEFak_RT, Me.mEFak_RW, Me.mEFak_Kecamatan, Me.mEFak_Kelurahan, Me.mEFak_Kabupaten, _
                Me.mEFak_Propinsi, Me.mEFak_KodePos, Me.mEFak_Telepon, Me.mEFak_KodeJenisTransaksiPajak, _
                Me.mPrimaryAccount, Me.mPrimaryCurrTax, Me.mPrimaryAccountDP, Me.mOtherCharge, Me.mOptionNIKPASS, Me.mNoNikPass, Me.mTaxAmount1, _
                Me.PrimaryAddInfo, Me.PrimaryFacilityStamp, Me.TaxRevised, Me.TaxType, Me.SPPBDate, Me.mDPPNilaiLain, Me.mPrimary, Me.mTimestamp) > 0 Then '//TR : Tambahan baru Nik/Pas
                    'update data yang berubah setelah simpan perubahan
                    Me.mNoUrut = NoUrut
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode

                    '2025 Core Tax Tahap 2.2
                    ''update data berkaitan dengan no faktur pajak
                    'If Me.mJenisFakturPajak = EnumJenisFakturPajak.FakturPajakStandar And Me.UseTax = True Then
                    '    While DTA.ScalarQuery_GetJmlByNoFakturPajak(Me.mNoFakturPajak).Value > 1
                    '        PrimaryMax = DTA.ScalarQuery_GetPrimaryMaxByNoFakturPajak(Me.mNoFakturPajak).Value
                    '        If PrimaryMax = _
                    '        Me.mPrimary Then
                    '            'get no baru
                    '            Me.mNoFakturPajak = Me.GetNoFakturPajakStandarBaru(Me.mTanggalTransaksi)
                    '            'update no-nya
                    '            If Me.mTanggalTransaksi >= Obj.BatasStartDate Then
                    '                Me.mUrutanFakturPajak = Obj.GetUrutan(Me.mNoFakturPajak, Me.mTanggalTransaksi)
                    '            Else
                    '                If Me.mJenisFakturPajak = EnumJenisFakturPajak.FakturPajakStandar Then
                    '                    Me.mUrutanFakturPajak = Me.NoFakturPajak.Substring(4)
                    '                Else
                    '                    Me.mUrutanFakturPajak = 0
                    '                    Me.mNoSeriFakturPajak = ""
                    '                    Me.mNoFakturPajak = ""
                    '                End If
                    '            End If
                    '            DTA.UpdateQuery_UpdateNoFakturPajak(Me.mUrutanFakturPajak, Me.mNoFakturPajak, _
                    '            Me.mPrimary)
                    '        Else
                    '            ij += 1
                    '            If ij > 10 Then
                    '                If DTA.ScalarQuery_GetPrimaryMaxByNoFakturPajakStandar2(Me.mNoFakturPajak, _
                    '                PrimaryMax).Value = Me.mPrimary Then
                    '                    'get no baru
                    '                    Me.mNoFakturPajak = Me.GetNoFakturPajakStandarBaru(Me.mTanggalTransaksi)
                    '                    'update no-nya
                    '                    If Me.mTanggalTransaksi >= Obj.BatasStartDate Then
                    '                        Me.mUrutanFakturPajak = Obj.GetUrutan(Me.mNoFakturPajak, Me.mTanggalTransaksi)
                    '                    Else
                    '                        If Me.mJenisFakturPajak = EnumJenisFakturPajak.FakturPajakStandar Then
                    '                            Me.mUrutanFakturPajak = Me.NoFakturPajak.Substring(4)
                    '                        Else
                    '                            Me.mUrutanFakturPajak = 0
                    '                            Me.mNoSeriFakturPajak = ""
                    '                            Me.mNoFakturPajak = ""
                    '                        End If
                    '                    End If
                    '                    DTA.UpdateQuery_UpdateNoFakturPajak(Me.mUrutanFakturPajak, Me.mNoFakturPajak, _
                    '                    Me.mPrimary)
                    '                End If
                    '                ij = 0
                    '            End If
                    '        End If
                    '    End While
                    'End If

                    ''simpan berkaitan other charge
                    'If vFrm.TBLHeaderCost.Rows.Count > 0 Then
                    '    For Each vFrm.rowHeaderCost In vFrm.TBLHeaderCost.Rows
                    '        With vFrm.rowHeaderCost
                    '            DTAHeaderCost.InsertQuery(Me.mPrimary, .ETD_Date, .ETD_Port, .ETA_Date, .ETA_Port, .Total_Nett_Weight, _
                    '            .Total_Gross_Weight, .primary_weight_unit, .Total_Nett_Volume, .Total_Gross_Volume, .primary_volume_unit, _
                    '            .primary_cust_ship_to, .identifier_cust_ship_to, .shipping_mark, .transportation_mode, .invoice_type, _
                    '            .var1_name, .var1_value, .var2_name, .var2_value, .var3_name, .var3_value, .var4_name, .var4_value, _
                    '            .var5_name, .var5_value)
                    '        End With
                    '    Next
                    'End If

                    'If vFrm.TBLitemCost.Rows.Count > 0 Then
                    '    For Each vFrm.rowitemCost In vFrm.TBLitemCost.Rows
                    '        With vFrm.rowitemCost
                    '            DTAItemCost.InsertQuery(Me.mPrimary, .Cost_Description, .Cost, .Remarks, .Primary_Akun)
                    '        End With
                    '    Next
                    'End If

                    '**************************************************************************************************************************************
                    '**************************************************************************************************************************************
                    'PERUBAHAN SISTEM PERHITUNGAN JURNAL ERROR NO 50 TANGGAL 11 JULI 2016 -- PERBAIKAN 28 JULI 2016
                    'SEBELUMNYA NILAI SALES RUPIAH = NILAI SALES USD * EXCHANGE RATE
                    'SEKARANG NILAINYA DICEK DULU APAKAH ADA PENJUMLAHAN TERHADAP NILAI SALES (CONTOH TAX USD, SHIPMENT DLL),
                    'HASILNYA DITOTAL LALU DIKURANGI SEMUA PENJUMLAHAN, SISANYA DIANGGAP SEBAGAI SALES
                    'HAL INI DILAKUKAN UTK MENCEGAH SELISIH ANTARA SALES DI INVOICE JURNAL DENGAN SALES DI RECEIVEMENT JOURNAL


                    Dim SalesRupiah As Decimal = 0
                    Dim TotalSales As Decimal = 0
                    Dim AmountTanpaSales As Decimal = 0

                    TotalSales = Me.ObjCD.DigitGrouping(Me.mTotal * Me.mTaxExchangeRate, DefPrecBase)

                    If Me.mPrimaryCurrTax = Me.mPrimaryCurrency Then
                        AmountTanpaSales += Me.ObjCD.DigitGrouping(Me.mTaxAmount * Me.mTaxExchangeRate, DefPrecBase)
                        AmountTanpaSales += Me.ObjCD.DigitGrouping(Me.mDutyStamp * Me.mTaxExchangeRate, DefPrecBase)
                    End If
                    'AmountTanpaSalesRp = Me.ObjCD.DigitGrouping((AmountTanpaSales * Me.mExchangeRate), DefPrecBase)
                    SalesRupiah = Me.ObjCD.DigitGrouping(TotalSales - AmountTanpaSales, DefPrecBase)


                    '**************************************************************************************************************************************
                    'simpan jurnal otomatis
                    Me.TransferJurnal_Invoice_DP2(Me.mPrimary, Me.mPrimaryCust, _
                    SalesRupiah, _
                    Me.ObjCD.DigitGrouping(Me.mTaxAmount * Me.mTaxExchangeRate, DefPrecBase), _
                    Me.ObjCD.DigitGrouping(Me.mDutyStamp * Me.mTaxExchangeRate, DefPrecBase), _
                    Me.mNoUrut, Me.mTanggalTransaksi, Me.mNoTransaksi, Me.mModifiedBy, _
                    CDbl(Me.mPaymentTerm), Me.mPrimaryCurrency, Me.mSubTotalB4DP, Me.mTaxExchangeRate, _
                    Me.mPrimaryCurrTax, Me.mPrimaryAccount, Me.mPrimaryAccountDP, Me.mTaxAmount, Me.mDutyStamp)

                    '2025 Core Tax Tahap 2.2
                    ''simpan no urut max untuk numbering
                    'If Obj.BatasStartDate <= Me.mTanggalTransaksi Then
                    '    If Me.mJenisFakturPajak = EnumJenisFakturPajak.FakturPajakStandar Then
                    '        DTA.UpdateQuery_PrimaryFakturPajak(Obj.GetPrimary(Me.mTanggalTransaksi, Me.mNoFakturPajak), Me.mPrimary)
                    '        PrimaryFakturPajak = DTA.ScalarQuery_GetPrimaryFakturPajak(Me.mPrimary)
                    '        NoFakturPajakMax = Me.GetNoFakturPajakMax(PrimaryFakturPajak, Me.mPrimary)

                    '        If Obj.GetUrutan(NoFakturPajakMax, Me.mTanggalTransaksi) < Obj.GetUrutan(Me.mNoFakturPajak, Me.mTanggalTransaksi) Then
                    '            Obj.SetNoUrutMaksimal(Me.mTanggalTransaksi, Me.mNoFakturPajak)
                    '        Else
                    '            Obj.SetNoUrutMaksimal(PrimaryFakturPajak, NoFakturPajakMax)
                    '        End If
                    '    End If
                    'End If

                    'nilai balikan
                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mNoTransaksi + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Invoice) = True Then
                    Me.mNoTransaksi = Me.GetNoBaru1(Me.mTanggalTransaksi)
                    Return Me.SimpanPerubahan(vUserCode)
                Else
                    MsgBox("Invoice number duplication, please change invoice number.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            DTA2.Dispose()
            ObjItem = Nothing
            ObjCur = Nothing
        End Try
    End Function


    Public Overridable Function SimpanPerubahan(ByVal dgvItem As DataGridView, ByVal vdgvItemDP As DataGridView, _
    ByVal vDGVItemSR As DataGridView, ByVal vDGVItemExchangeRate As DataGridView, ByVal vUserCode As String, ByVal vFrm As FrmTransaksiSalesInvoice) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            SW.Start()

            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
             Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiInvoice_SimpanPerubahan, vUserCode)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= ObjAntrian.GetWaitTime Then
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                        Exit Function
                    Else
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanPerubahan1(dgvItem, vdgvItemDP, _
                            vDGVItemSR, vDGVItemExchangeRate, vUserCode, vFrm)

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

    Public Overridable Function SimpanPerubahan(ByVal dgvItem As DataGridView, ByVal vdgvItemDP As DataGridView, _
  ByVal vDGVItemSR As DataGridView, ByVal vDGVItemExchangeRate As DataGridView, ByVal vUserCode As String, _
  ByVal vFrm As FrmTransaksiSalesInvoiceQuick, ByVal vNoDO As String, _
  ByVal vTblCustom As CustomSMESS.DseTransaksiSalesInvoiceQuick_AddByCustomProduct.tbltrans_custom_salesinvoice_item1DataTable, _
  ByVal TblCustomInOutDummy As CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            SW.Start()

            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
            Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiInvoice_Quick_SimpanPerubahan, vUserCode)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= Me.ObjAntrian.GetWaitTime Then
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                        Exit Function
                    Else
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanPerubahan1(dgvItem, vdgvItemDP, _
                            vDGVItemSR, vDGVItemExchangeRate, vUserCode, vFrm, vNoDO, vTblCustom, TblCustomInOutDummy)

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

    Public Overridable Function SimpanPerubahan1(ByVal dgvItem As DataGridView, ByVal vdgvItemDP As DataGridView, _
    ByVal vDGVItemSR As DataGridView, ByVal vDGVItemExchangeRate As DataGridView, ByVal vUserCode As String, _
    ByVal vFrm As FrmTransaksiSalesInvoice) As Boolean
        Dim DTA2 As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_NewTableAdapter
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice1TableAdapter
        Dim DTAHeaderCost As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_add_dataTableAdapter
        Dim DTAItemCost As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_item_costTableAdapter
        Dim ObjItem As New ClsTransaksiSalesInvoiceItem
        Dim UrutanFakturPajak As String = ""
        Dim ij As Double = 0
        Dim PrimaryMax As Double = -1
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim NoUrut As Double = 0
        Dim Obj As New ClsTransaksiSalesInvoice_SettingTaxInvoice
        Dim PrimaryFakturPajak As Double = -1
        Dim NoFakturPajakMax As String = ""
        Dim DefPrecBase As Integer = 0
        Dim vShipmentTax, vVatDipungut As Integer

        Dim DPOriginal, DPBaseCurrency As Decimal
        Dim DTA1 As New DseTransaksiSalesInvoiceTableAdapters.DataTable13TableAdapter
        Dim TBL1 As New DseTransaksiSalesInvoice.DataTable13DataTable
        Dim row1 As DseTransaksiSalesInvoice.DataTable13Row
        Dim TipeDP1 As EnumTipeDP = EnumTipeDP.InvDP
        Dim TBL2 As New DseTransaksiSalesInvoice.DataTable13DataTable
        Dim row2 As DseTransaksiSalesInvoice.DataTable13Row
        Dim TipeDP2 As EnumTipeDP = EnumTipeDP.ManualDP
        Dim PrimaryAkunDP As Double = -1

        Try
            If Me.ApakahDataBenar(False, dgvItem, vFrm, True, False) = True Then
                DefPrecBase = ObjCur.GetDecimal_DefaultCurrency

                '2025 Core Tax Tahap 2.2
                'If Me.mTanggalTransaksi >= Obj.BatasStartDate Then
                '    Me.mUrutanFakturPajak = Obj.GetUrutan(Me.mNoFakturPajak, Me.mTanggalTransaksi)
                'Else
                '    If Me.mJenisFakturPajak = EnumJenisFakturPajak.FakturPajakStandar Then
                '        Me.mUrutanFakturPajak = Me.NoFakturPajak.Substring(4)
                '    Else
                '        Me.mUrutanFakturPajak = 0
                '        Me.mNoSeriFakturPajak = ""
                '        Me.mNoFakturPajak = ""
                '    End If
                'End If

                'hitung level otorisasi yang dibutuhkan
                If Me.mShipmentWithTax = True Then
                    vShipmentTax = 1
                Else
                    vShipmentTax = 0
                End If
                If Me.mVATDipungut = True Then
                    vVatDipungut = 1
                Else
                    vVatDipungut = 0
                End If
                NoUrut = Me.ObjNumbering.GetUrutan(Me.mNoTransaksi, Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Invoice, _
               Me.mTanggalTransaksi)
                If DTA2.UpdateQuery(Me.mNoTransaksi, NoUrut, Me.mUrutanFakturPajak, Me.mTipeInvoice, _
                Me.mTanggalTransaksi, Me.mPrimaryCust, Me.mPrimaryCurrency, Me.mCalculatesBy, Me.mSubTotalB4DP, _
                Me.mDownPayment, Me.mSubTotal, Me.mDutyStamp, Me.mUseTax, Me.mTaxPercentage, Me.mTaxExchangeRate, _
                Me.mTaxAmount, Me.mTotal, Me.mExchangeRateToUSD, Me.mExhangeRateToIDR, Me.mNoSeriFakturPajak, _
                Me.mNoFakturPajak, Me.mNoNPWP, Me.mIdentifier_CustAddress, Me.mRemarks, vUserCode, Now, Now, _
                Me.mJenisFakturPajak, Me.mPaymentTerm, Me.mPrimarySO4DP, Me.mNoFakturPajakSederhana, _
                Me.mNamaSigned, Me.mJabatanSigned, Me.mAlamatPajak1, Me.mAlamatPajak2, Me.mExchangeRate, Me.mOngkosKirim, _
                vVatDipungut, vShipmentTax, Me.mPrimarySalesPerson, Me.mDiscString, Me.mDiscAmount, Me.mEFak_NamaJalan, _
                Me.mEFak_Blok, Me.mEFak_No, Me.mEFak_RT, Me.mEFak_RW, Me.mEFak_Kecamatan, Me.mEFak_Kelurahan, Me.mEFak_Kabupaten, _
                Me.mEFak_Propinsi, Me.mEFak_KodePos, Me.mEFak_Telepon, Me.mEFak_KodeJenisTransaksiPajak, _
                Me.mPrimaryAccount, Me.mPrimaryCurrTax, Me.mPrimaryAccountDP, Me.mOtherCharge, Me.mOptionNIKPASS, Me.mNoNikPass, Me.mTaxAmount1, _
                Me.PrimaryAddInfo, Me.PrimaryFacilityStamp, Me.TaxRevised, Me.TaxType, Me.SPPBDate, Me.mDPPNilaiLain, Me.mPrimary, Me.mTimestamp) > 0 Then '//TR : Tambahan baru Nik/Pas
                    'update data yang berubah setelah simpan perubahan
                    Me.mNoUrut = NoUrut
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode

                    '2025 Core Tax Tahap 2.2
                    ''update data berkaitan dengan no faktur pajak
                    'If Me.mJenisFakturPajak = EnumJenisFakturPajak.FakturPajakStandar And Me.UseTax = True Then
                    '    While DTA.ScalarQuery_GetJmlByNoFakturPajak(Me.mNoFakturPajak).Value > 1
                    '        PrimaryMax = DTA.ScalarQuery_GetPrimaryMaxByNoFakturPajak(Me.mNoFakturPajak).Value
                    '        If PrimaryMax = _
                    '        Me.mPrimary Then
                    '            'get no baru
                    '            Me.mNoFakturPajak = Me.GetNoFakturPajakStandarBaru(Me.mTanggalTransaksi)
                    '            'update no-nya
                    '            If Me.mTanggalTransaksi >= Obj.BatasStartDate Then
                    '                Me.mUrutanFakturPajak = Obj.GetUrutan(Me.mNoFakturPajak, Me.mTanggalTransaksi)
                    '            Else
                    '                If Me.mJenisFakturPajak = EnumJenisFakturPajak.FakturPajakStandar Then
                    '                    Me.mUrutanFakturPajak = Me.NoFakturPajak.Substring(4)
                    '                Else
                    '                    Me.mUrutanFakturPajak = 0
                    '                    Me.mNoSeriFakturPajak = ""
                    '                    Me.mNoFakturPajak = ""
                    '                End If
                    '            End If
                    '            DTA.UpdateQuery_UpdateNoFakturPajak(Me.mUrutanFakturPajak, Me.mNoFakturPajak, _
                    '            Me.mPrimary)
                    '        Else
                    '            ij += 1
                    '            If ij > 10 Then
                    '                If DTA.ScalarQuery_GetPrimaryMaxByNoFakturPajakStandar2(Me.mNoFakturPajak, _
                    '                PrimaryMax).Value = Me.mPrimary Then
                    '                    'get no baru
                    '                    Me.mNoFakturPajak = Me.GetNoFakturPajakStandarBaru(Me.mTanggalTransaksi)
                    '                    'update no-nya
                    '                    If Me.mTanggalTransaksi >= Obj.BatasStartDate Then
                    '                        Me.mUrutanFakturPajak = Obj.GetUrutan(Me.mNoFakturPajak, Me.mTanggalTransaksi)
                    '                    Else
                    '                        If Me.mJenisFakturPajak = EnumJenisFakturPajak.FakturPajakStandar Then
                    '                            Me.mUrutanFakturPajak = Me.NoFakturPajak.Substring(4)
                    '                        Else
                    '                            Me.mUrutanFakturPajak = 0
                    '                            Me.mNoSeriFakturPajak = ""
                    '                            Me.mNoFakturPajak = ""
                    '                        End If
                    '                    End If
                    '                    DTA.UpdateQuery_UpdateNoFakturPajak(Me.mUrutanFakturPajak, Me.mNoFakturPajak, _
                    '                    Me.mPrimary)
                    '                End If
                    '                ij = 0
                    '            End If
                    '        End If
                    '    End While
                    'End If

                    'hapus data item lama
                    ObjItem.Hapus(Me.mPrimary)
                    'simpan data item baru
                    ObjItem.Simpan(dgvItem, vDGVItemExchangeRate, vdgvItemDP, Me.mPrimary, vDGVItemSR, vFrm)
                    'simpan data sales amount
                    Me.SimpanDataSalesAmount(Me.mCalculatesBy, Me.mPrimary, _
                    Math.Round(Me.mSubTotalB4DP * Me.mTaxExchangeRate, DefPrecBase, MidpointRounding.AwayFromZero))

                    'hitung DP
                    DPOriginal = 0
                    DPBaseCurrency = 0
                    If Me.mDownPayment > 0 Then
                        DTA1.Fill(TBL1, Me.mPrimary, CInt(TipeDP1))
                        For Each row1 In TBL1.Rows
                            DPOriginal += ObjCD.DigitGrouping(row1.Sub_Total_Before_DP, ObjCur.GetDefaultPrecission(Me.mPrimaryCurrTax))
                            DPBaseCurrency += ObjCD.DigitGrouping(row1.Sub_Total_Before_DP * row1.Tax_Exchange_Rate, DefPrecBase)
                            If PrimaryAkunDP = -1 Then
                                PrimaryAkunDP = Me.GetPrimaryAccountDP(row1.Primary_Invoice_DP)
                            End If
                        Next

                        DTA1.FillBy_DPBalance(TBL2, Me.mPrimary, CInt(TipeDP2))
                        For Each row2 In TBL2.Rows
                            DPOriginal += ObjCD.DigitGrouping(row2.Sub_Total_Before_DP, ObjCur.GetDefaultPrecission(Me.mPrimaryCurrTax))
                            DPBaseCurrency += ObjCD.DigitGrouping(row2.Sub_Total_Before_DP * row2.Tax_Exchange_Rate, DefPrecBase)
                            If PrimaryAkunDP = -1 Then
                                PrimaryAkunDP = Me.GetPrimaryAccountDP_BalanceDP(row2.Primary_Invoice_DP)
                            End If
                        Next
                    End If

                    'simpan berkaitan other charge
                    If vFrm.TBLHeaderCost.Rows.Count > 0 Then
                        DTAHeaderCost.DeleteQuery(Me.mPrimary)
                        For Each vFrm.rowHeaderCost In vFrm.TBLHeaderCost.Rows
                            With vFrm.rowHeaderCost
                                DTAHeaderCost.InsertQuery(Me.mPrimary, .ETD_Date, .ETD_Port, .ETA_Date, .ETA_Port, .Total_Nett_Weight, _
                                .Total_Gross_Weight, .primary_weight_unit, .Total_Nett_Volume, .Total_Gross_Volume, .primary_volume_unit, _
                                .primary_cust_ship_to, .identifier_cust_ship_to, .shipping_mark, .transportation_mode, .invoice_type, _
                                .var1_name, .var1_value, .var2_name, .var2_value, .var3_name, .var3_value, .var4_name, .var4_value, _
                                .var5_name, .var5_value)
                            End With
                        Next
                    End If

                    DTAItemCost.DeleteQuery(Me.mPrimary)
                    If vFrm.TBLitemCost.Rows.Count > 0 Then
                        For Each vFrm.rowitemCost In vFrm.TBLitemCost.Rows
                            With vFrm.rowitemCost
                                DTAItemCost.InsertQuery(Me.mPrimary, .Cost_Description, .Cost, .Remarks, .Primary_Akun)
                            End With
                        Next
                    End If


                    'simpan jurnal otomatis
                    DefPrecBase = ObjCur.GetDecimal_DefaultCurrency

                    '**************************************************************************************************************************************
                    '**************************************************************************************************************************************
                    'PERUBAHAN SISTEM PERHITUNGAN JURNAL ERROR NO 50 TANGGAL 11 JULI 2016 -- PERBAIKAN 28 JULI 2016
                    'SEBELUMNYA NILAI SALES RUPIAH = NILAI SALES USD * EXCHANGE RATE
                    'SEKARANG NILAINYA DICEK DULU APAKAH ADA PENJUMLAHAN TERHADAP NILAI SALES (CONTOH TAX USD, SHIPMENT DLL),
                    'HASILNYA DITOTAL LALU DIKURANGI SEMUA PENJUMLAHAN, SISANYA DIANGGAP SEBAGAI SALES
                    'HAL INI DILAKUKAN UTK MENCEGAH SELISIH ANTARA SALES DI INVOICE JURNAL DENGAN SALES DI RECEIVEMENT JOURNAL


                    Dim SalesRupiah As Decimal = 0
                    Dim TotalSales As Decimal = 0
                    Dim AmountTanpaSales As Decimal = 0

                    Dim SalesDanDP As Decimal = 0

                    SalesDanDP = Me.ObjCD.DigitGrouping((Me.mTotal + DPOriginal), ObjCur.GetDefaultPrecission(Me.mPrimaryCurrency))
                    TotalSales = Me.ObjCD.DigitGrouping(SalesDanDP * Me.mTaxExchangeRate, DefPrecBase)

                    AmountTanpaSales += Me.ObjCD.DigitGrouping(Me.mOngkosKirim * Me.mTaxExchangeRate, DefPrecBase)
                    AmountTanpaSales += Me.ObjCD.DigitGrouping(Me.mDiscAmount * Me.mTaxExchangeRate, DefPrecBase)
                    AmountTanpaSales += Me.ObjCD.DigitGrouping(Me.mOtherCharge * Me.mTaxExchangeRate, DefPrecBase)
                    If Me.mPrimaryCurrTax = Me.mPrimaryCurrency Then
                        AmountTanpaSales += Me.ObjCD.DigitGrouping(Me.mTaxAmount * Me.mTaxExchangeRate, DefPrecBase)
                        AmountTanpaSales += Me.ObjCD.DigitGrouping(Me.mDutyStamp * Me.mTaxExchangeRate, DefPrecBase)
                    End If
                    'AmountTanpaSalesRp = Me.ObjCD.DigitGrouping((AmountTanpaSales * Me.mExchangeRate), DefPrecBase)
                    SalesRupiah = Me.ObjCD.DigitGrouping(TotalSales - AmountTanpaSales, DefPrecBase)


                    '**************************************************************************************************************************************





                    Me.TransferJurnal_Invoice2(Me.mPrimary, Me.mPrimaryCust, _
                     SalesRupiah, _
                    ObjCD.DigitGrouping(Me.mTaxAmount * Me.mTaxExchangeRate, DefPrecBase), _
                    ObjCD.DigitGrouping(Me.mDutyStamp * Me.mTaxExchangeRate, DefPrecBase), _
                    Me.mNoUrut, Me.mTanggalTransaksi, Me.mNoTransaksi, Me.mModifiedBy, CDbl(Me.mPaymentTerm), Me.mTaxExchangeRate, _
                    Me.mSubTotalB4DP, Me.mPrimaryCurrency, Me.ObjCD.DigitGrouping(Me.mOngkosKirim * Me.mTaxExchangeRate, DefPrecBase), _
                    DPOriginal, DPBaseCurrency, Me.mOngkosKirim, Me.mDiscAmount, _
                    ObjCD.DigitGrouping(Me.mDiscAmount * Me.mTaxExchangeRate, DefPrecBase), _
                    Me.mPrimaryCurrTax, Me.mPrimaryAccount, Me.mDutyStamp, Me.mTaxAmount, PrimaryAkunDP, _
                    Me.mOtherCharge * Me.mTaxExchangeRate, Me.mOtherCharge, vFrm.TBLitemCost, vFrm.rowitemCost)


                    '2025 Core Tax Tahap 2.2
                    ''simpan no urut max untuk numbering
                    'If Obj.BatasStartDate <= Me.mTanggalTransaksi Then
                    '    If Me.mJenisFakturPajak = EnumJenisFakturPajak.FakturPajakStandar Then
                    '        DTA.UpdateQuery_PrimaryFakturPajak(Obj.GetPrimary(Me.mTanggalTransaksi, Me.mNoFakturPajak), Me.mPrimary)
                    '        PrimaryFakturPajak = DTA.ScalarQuery_GetPrimaryFakturPajak(Me.mPrimary)
                    '        NoFakturPajakMax = Me.GetNoFakturPajakMax(PrimaryFakturPajak, Me.mPrimary)

                    '        If Obj.GetUrutan(NoFakturPajakMax, Me.mTanggalTransaksi) < Obj.GetUrutan(Me.mNoFakturPajak, Me.mTanggalTransaksi) Then
                    '            Obj.SetNoUrutMaksimal(Me.mTanggalTransaksi, Me.mNoFakturPajak)
                    '        Else
                    '            Obj.SetNoUrutMaksimal(PrimaryFakturPajak, NoFakturPajakMax)
                    '        End If
                    '    End If
                    'End If

                    'nilai balikan
                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mNoTransaksi + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Invoice) = True Then
                    Me.mNoTransaksi = Me.GetNoBaru1(Me.mTanggalTransaksi)
                    Return Me.SimpanPerubahan1(dgvItem, vdgvItemDP, vDGVItemSR, vDGVItemExchangeRate, vUserCode, vFrm)
                Else
                    MsgBox("Invoice number duplication, please change invoice number", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
            ObjCur = Nothing
            DTA1.Dispose()
            TBL1.Dispose()
        End Try
    End Function

    Public Function GetPrimaryAccountDP_BalanceDP(ByVal vPrimary As Double) As Double
        Dim DTA As New Core_Module.DseBalanceDPTableAdapters.tbltrans_historicalbalancemanual_dpTableAdapter

        Try
            If DTA.ScalarQuery_GetPrimaryDPAccount(vPrimary).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryDPAccount(vPrimary).Value
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

    Public Function GetPrimaryAccountDP(ByVal vPrimary As Double) As Double
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_NewTableAdapter

        Try
            If DTA.ScalarQuery_GetPrimaryAccountDP(vPrimary).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryAccountDP(vPrimary).Value
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

    Public Function GetPrimaryAccountAR(ByVal vPrimary As Double) As Double
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_NewTableAdapter

        Try
            If DTA.ScalarQuery_GetPrimaryAccount(vPrimary).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryAccount(vPrimary).Value
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

    Public Function GetPrimaryCurrTax(ByVal vPrimary As Double) As Double
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_NewTableAdapter

        Try
            If DTA.ScalarQuery_GetPrimaryCurrTax(vPrimary).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryCurrTax(vPrimary).Value
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

    Public Function GetJmlAOPSama(ByVal vPrimaryPO As Double, ByVal VPrimaryAccount As Double, ByVal vPrimaryCurrTax As Double, ByVal vTipeAOP As ClsTransaksiSalesInvoice.EnumTipeInvoice) As Decimal
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.BatasanInvoiceDPTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJmlInvoiceDataSama(vTipeAOP, vPrimaryPO, VPrimaryAccount, vPrimaryCurrTax).HasValue AndAlso _
               DTA.ScalarQuery_GetJmlInvoiceDataSama(vTipeAOP, vPrimaryPO, VPrimaryAccount, vPrimaryCurrTax).Value > 0 Then
                Return DTA.ScalarQuery_GetJmlInvoiceDataSama(vTipeAOP, vPrimaryPO, VPrimaryAccount, vPrimaryCurrTax).Value
            Else
                Return 0
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlAOPwithPOSama(ByVal vPrimaryPO As Double, ByVal vTipeAOP As ClsTransaksiSalesInvoice.EnumTipeInvoice) As Decimal
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.BatasanInvoiceDPTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJmlBySONTipe(vTipeAOP, vPrimaryPO).HasValue AndAlso _
            DTA.ScalarQuery_GetJmlBySONTipe(vTipeAOP, vPrimaryPO).Value > 0 Then
                Return DTA.ScalarQuery_GetJmlBySONTipe(vTipeAOP, vPrimaryPO).Value
            Else
                Return 0
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryAOPwithPOSama(ByVal vPrimaryPO As Double, ByVal vTipeAOP As ClsTransaksiSalesInvoice.EnumTipeInvoice) As Decimal
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.BatasanInvoiceDPTableAdapter
        Dim tbl As New DseTransaksiSalesInvoice.BatasanInvoiceDPDataTable
        Dim row As DseTransaksiSalesInvoice.BatasanInvoiceDPRow
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            DTA.FillBy_PrimarySONTipeInvoice(tbl, vPrimaryPO, vTipeAOP)
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
    '====================================================================================================================================================================
    'Erna's Code selesai()
    '====================================================================================================================================================================

    'elemen yang digunakan pada automatic invoice transfer
    Public Enum EnumTipeSalesBalance As Integer
        Automatic
        Manual
        Manual_Transfer '--> tidak digunakan lagi, hanya pada cmsi saja
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

    Public Enum EnumTipeInvoice As Integer
        InvoiceDP
        SalesInvoice
    End Enum

    Public Property DiscString() As String
        Get
            Return Me.mDiscString
        End Get
        Set(ByVal value As String)
            mDiscString = value
        End Set
    End Property

    Public Property DiscAmount() As Decimal
        Get
            Return Me.mDiscAmount
        End Get
        Set(ByVal value As Decimal)
            mDiscAmount = value
        End Set
    End Property

    Public Property PrimarySalesPerson() As Double
        Get
            Return Me.mPrimarySalesPerson
        End Get
        Set(ByVal value As Double)
            mPrimarySalesPerson = value
        End Set
    End Property

    Public Property VATDipungut() As Boolean
        Get
            Return Me.mVATDipungut
        End Get
        Set(ByVal value As Boolean)
            mVATDipungut = value
        End Set
    End Property

    Public Property ShipmentWithTax() As Boolean
        Get
            Return Me.mShipmentWithTax
        End Get
        Set(ByVal value As Boolean)
            mShipmentWithTax = value
        End Set
    End Property

    Public Property OngkosKirim() As Decimal
        Get
            Return Me.mOngkosKirim
        End Get
        Set(ByVal value As Decimal)
            mOngkosKirim = value
        End Set
    End Property

    Public Property ExchangeRate() As Decimal
        Get
            Return Me.mExchangeRate
        End Get
        Set(ByVal value As Decimal)
            mExchangeRate = value
        End Set
    End Property

    Public Property AlamatPajak1() As String
        Get
            Return Me.mAlamatPajak1
        End Get
        Set(ByVal value As String)
            mAlamatPajak1 = value
        End Set
    End Property

    Public Property AlamatPajak2() As String
        Get
            Return Me.mAlamatPajak2
        End Get
        Set(ByVal value As String)
            mAlamatPajak2 = value
        End Set
    End Property

    Public Property NamaSigned() As String
        Get
            Return Me.mNamaSigned
        End Get
        Set(ByVal value As String)
            mNamaSigned = value
        End Set
    End Property

    Public Property JabatanSigned() As String
        Get
            Return Me.mJabatanSigned
        End Get
        Set(ByVal value As String)
            mJabatanSigned = value
        End Set
    End Property

    Public Property NoFakturPajakSederhana() As String
        Get
            Return Me.mNoFakturPajakSederhana
        End Get
        Set(ByVal value As String)
            mNoFakturPajakSederhana = value
        End Set
    End Property

    Public Enum EnumCalculatesBy As Integer
        Invoice
        Delivery_Order
        Sales_Order
    End Enum

    Public Enum EnumJenisFakturPajak As Integer
        FakturPajakSederhana
        FakturPajakStandar
    End Enum

    Public Property PrimarySO4DP() As Double
        Get
            Return Me.mPrimarySO4DP
        End Get
        Set(ByVal value As Double)
            mPrimarySO4DP = value
        End Set
    End Property

    Public Property TipeInvoice() As EnumTipeInvoice
        Get
            Return Me.mTipeInvoice
        End Get
        Set(ByVal value As EnumTipeInvoice)
            mTipeInvoice = value
        End Set
    End Property

    Public Property Revisi() As Integer
        Get
            Return Me.mRevisi
        End Get
        Set(ByVal value As Integer)
            mRevisi = value
        End Set
    End Property

    Public Property PaymentTerm() As String
        Get
            Return Me.mPaymentTerm
        End Get
        Set(ByVal value As String)
            mPaymentTerm = value
        End Set
    End Property

    Public Property NoNPWP() As String
        Get
            Return Me.mNoNPWP
        End Get
        Set(ByVal value As String)
            mNoNPWP = value
        End Set
    End Property

    Public Property TaxExchangeRate() As Double
        Get
            Return Me.mTaxExchangeRate
        End Get
        Set(ByVal value As Double)
            mTaxExchangeRate = value
        End Set
    End Property

    Public Property NoSeriFakturPajak() As String
        Get
            Return Me.mNoSeriFakturPajak
        End Get
        Set(ByVal value As String)
            mNoSeriFakturPajak = value
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

    Public Property UrutanFakturPajak() As Double
        Get
            Return Me.mUrutanFakturPajak
        End Get
        Set(ByVal value As Double)
            mUrutanFakturPajak = value
        End Set
    End Property

    Public Property DutyStamp() As Double
        Get
            Return Me.mDutyStamp
        End Get
        Set(ByVal value As Double)
            mDutyStamp = value
        End Set
    End Property

    Public Property SubTotalB4DP() As Double
        Get
            Return Me.mSubTotalB4DP
        End Get
        Set(ByVal value As Double)
            mSubTotalB4DP = value
        End Set
    End Property

    Public Property DownPayment() As Double
        Get
            Return Me.mDownPayment
        End Get
        Set(ByVal value As Double)
            mDownPayment = value
        End Set
    End Property

    Public Property ExchangRateToIDR() As Double
        Get
            Return Me.mExhangeRateToIDR
        End Get
        Set(ByVal value As Double)
            mExhangeRateToIDR = value
        End Set
    End Property

    Public Property ExhangeRateToUSD() As Double
        Get
            Return Me.mExchangeRateToUSD
        End Get
        Set(ByVal value As Double)
            Me.mExchangeRateToUSD = value
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

    Public Property JenisFakturPajak() As EnumJenisFakturPajak
        Get
            Return Me.mJenisFakturPajak
        End Get
        Set(ByVal value As EnumJenisFakturPajak)
            mJenisFakturPajak = value
        End Set
    End Property

    Public Property Total() As Double
        Get
            Return Me.mTotal
        End Get
        Set(ByVal value As Double)
            mTotal = value
        End Set
    End Property

    Public Property TaxAmount() As Double
        Get
            Return Me.mTaxAmount
        End Get
        Set(ByVal value As Double)
            mTaxAmount = value
        End Set
    End Property

    Public Property TaxPercentage() As Double
        Get
            Return Me.mTaxPercentage
        End Get
        Set(ByVal value As Double)
            mTaxPercentage = value
        End Set
    End Property

    Public Property UseTax() As Boolean
        Get
            If Me.mUseTax = 0 Then
                Return False
            Else
                Return True
            End If
        End Get
        Set(ByVal value As Boolean)
            If value = False Then
                Me.mUseTax = 0
            Else
                Me.mUseTax = 1
            End If
        End Set
    End Property

    Public Property SubTotal() As Double
        Get
            Return Me.mSubTotal
        End Get
        Set(ByVal value As Double)
            mSubTotal = value
        End Set
    End Property

    Public Property PrimaryCurrency() As Double
        Get
            Return Me.mPrimaryCurrency
        End Get
        Set(ByVal value As Double)
            mPrimaryCurrency = value
        End Set
    End Property

    Public Property CustIdentifierAdd() As String
        Get
            Return Me.mIdentifier_CustAddress
        End Get
        Set(ByVal value As String)
            mIdentifier_CustAddress = value
        End Set
    End Property

    Public Property PrimaryCustomer() As Double
        Get
            Return Me.mPrimaryCust
        End Get
        Set(ByVal value As Double)
            mPrimaryCust = value
        End Set
    End Property

    Public Property Remarks() As String
        Get
            Return Me.mRemarks
        End Get
        Set(ByVal value As String)
            mRemarks = value
        End Set
    End Property

    Public Property TanggalTransaksi() As Date
        Get
            Return Me.mTanggalTransaksi
        End Get
        Set(ByVal value As Date)
            mTanggalTransaksi = value
        End Set
    End Property

    Public Property NoTransaksi() As String
        Get
            Return Me.mNoTransaksi
        End Get
        Set(ByVal value As String)
            mNoTransaksi = value
        End Set
    End Property

    Public Property NoUrut() As Double
        Get
            Return Me.mNoUrut
        End Get
        Set(ByVal value As Double)
            mNoUrut = value
        End Set
    End Property

    Public Property Primary() As Double
        Get
            Return mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
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

    Public Property EFak_KodeJenisTransaksiPajak() As Integer
        Get
            Return Me.mEFak_KodeJenisTransaksiPajak
        End Get
        Set(ByVal value As Integer)
            mEFak_KodeJenisTransaksiPajak = value
        End Set
    End Property

    Public Property EFak_Telepon() As String
        Get
            Return Me.mEFak_Telepon
        End Get
        Set(ByVal value As String)
            mEFak_Telepon = value
        End Set
    End Property

    Public Property EFak_KodePos() As String
        Get
            Return Me.mEFak_KodePos
        End Get
        Set(ByVal value As String)
            mEFak_KodePos = value
        End Set
    End Property

    Public Property EFak_Propinsi() As String
        Get
            Return Me.mEFak_Propinsi
        End Get
        Set(ByVal value As String)
            mEFak_Propinsi = value
        End Set
    End Property

    Public Property EFak_Kabupaten() As String
        Get
            Return Me.mEFak_Kabupaten
        End Get
        Set(ByVal value As String)
            mEFak_Kabupaten = value
        End Set
    End Property

    Public Property EFak_Kelurahan() As String
        Get
            Return Me.mEFak_Kelurahan
        End Get
        Set(ByVal value As String)
            mEFak_Kelurahan = value
        End Set
    End Property

    Public Property EFak_Kecamatan() As String
        Get
            Return Me.mEFak_Kecamatan
        End Get
        Set(ByVal value As String)
            mEFak_Kecamatan = value
        End Set
    End Property

    Public Property EFak_RW() As String
        Get
            Return Me.mEFak_RW
        End Get
        Set(ByVal value As String)
            mEFak_RW = value
        End Set
    End Property

    Public Property EFak_RT() As String
        Get
            Return Me.mEFak_RT
        End Get
        Set(ByVal value As String)
            mEFak_RT = value
        End Set
    End Property

    Public Property EFak_No() As String
        Get
            Return Me.mEFak_No
        End Get
        Set(ByVal value As String)
            mEFak_No = value
        End Set
    End Property

    Public Property EFak_Blok() As String
        Get
            Return Me.mEFak_Blok
        End Get
        Set(ByVal value As String)
            mEFak_Blok = value
        End Set
    End Property

    Public Property EFak_NamaJalan() As String
        Get
            Return Me.mEFak_NamaJalan
        End Get
        Set(ByVal value As String)
            mEFak_NamaJalan = value
        End Set
    End Property

    Public Function GetInvoiceNoBySO_n_Product(ByVal vPrimarySO As Double, _
       ByVal vPrimaryProduct As Double) As String
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice3TableAdapter
        Dim TBL As New DseTransaksiSalesInvoice.tbltrans_salesinvoice3DataTable
        Dim row As DseTransaksiSalesInvoice.tbltrans_salesinvoice3Row
        Dim retval As String = ""

        Try
            If DTA.Fill(TBL, vPrimaryProduct, vPrimarySO) > 0 Then
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

    Public Function GetMaximumUrutanFP(ByVal vPrimaryFP As Double) As Double
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice1TableAdapter
        Dim Obj As New ClsTransaksiSalesInvoice_SettingTaxInvoice

        Try
            If DTA.ScalarQuery_GetMaxUrutanFP(vPrimaryFP, Obj.BatasStartDate).HasValue = True Then
                Return DTA.ScalarQuery_GetMaxUrutanFP(vPrimaryFP, Obj.BatasStartDate).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Obj = Nothing
        End Try
    End Function

    Public Function GetMinimumUrutanFP(ByVal vPrimaryFP As Double) As Double
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice1TableAdapter
        Dim Obj As New ClsTransaksiSalesInvoice_SettingTaxInvoice

        Try
            If DTA.ScalarQuery_GetMinUrutanFP(vPrimaryFP, Obj.BatasStartDate).HasValue = True Then
                Return DTA.ScalarQuery_GetMinUrutanFP(vPrimaryFP, Obj.BatasStartDate).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Obj = Nothing
        End Try
    End Function


    Public Function Get_Primary_Currency(ByVal vPrimary As Double) As Double
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoiceTableAdapter

        Try
            If DTA.ScalarQuery_GetPrimaryCurrency(vPrimary).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryCurrency(vPrimary).Value
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

    Public Function Get_Primary_Cust(ByVal vPrimaryInv As Double) As Double
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice1TableAdapter

        Try
            If DTA.ScalarQuery_GetPrimaryCust(vPrimaryInv).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryCust(vPrimaryInv).Value
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

    Public Sub GetAlamatPajak(ByVal vPrimaryInvoice As Double, ByRef vAlamat1 As String, ByRef vAlamat2 As String)
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice1TableAdapter

        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetAlamatPajak1(vPrimaryInvoice)) = True Then
                vAlamat1 = ""
            Else
                vAlamat1 = DTA.ScalarQuery_GetAlamatPajak1(vPrimaryInvoice)
            End If
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetAlamatPajak2(vPrimaryInvoice)) = True Then
                vAlamat2 = ""
            Else
                vAlamat2 = DTA.ScalarQuery_GetAlamatPajak2(vPrimaryInvoice)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function ApakahMenggunakanSettingFP(ByVal vPrimarySettingFP As Double) As Boolean
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice1TableAdapter
        Dim ObjTaxSetting As New ClsTransaksiSalesInvoice_SettingTaxInvoice

        Try
            If DTA.ScalarQuery_GetJmlByFP(vPrimarySettingFP, ObjTaxSetting.BatasStartDate).HasValue = True Then
                If DTA.ScalarQuery_GetJmlByFP(vPrimarySettingFP, ObjTaxSetting.BatasStartDate).Value > 0 Then
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

    Public Function GetDPNo(ByVal vPrimaryInvoice As Double) As String
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_itemdp_viewTableAdapter
        Dim TBL As New DseTransaksiSalesInvoice.tbltrans_salesinvoice_itemdp_viewDataTable
        Dim rowDP As DseTransaksiSalesInvoice.tbltrans_salesinvoice_itemdp_viewRow
        Dim retval As String = ""

        Try
            DTA.FillBy_Primary(TBL, vPrimaryInvoice)
            For Each rowDP In TBL.Rows
                If retval = "" Then
                    retval = rowDP.No_Invoice_DP
                Else
                    retval = retval + ", " + rowDP.No_Invoice_DP
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

    Private Sub KurangkanItemInvoice(ByVal rowReturn As dseTransaksiSalesReturn.tbltrans_salesreturn_item_view0Row, _
    ByVal vTBLIsi As DseTransaksiSalesInvoice.TblCetakIsiDataTable, ByVal DefaultPrec As Integer)
        Dim rowInv As DseTransaksiSalesInvoice.TblCetakIsiRow
        Dim counter As Integer = 0
        Dim HargaLama, HargaBaru As Decimal

        Try
            While rowReturn.Total_Qty_Multi > 0
                While counter < vTBLIsi.Rows.Count
                    rowInv = vTBLIsi.Rows(counter)
                    If rowInv.Product_Code = rowReturn.Product_Code And rowInv.No_SO = rowReturn.SO_No And _
                    rowInv.No_DO = rowReturn.DO_Number Then
                        If rowInv.Total_Qty = rowReturn.Total_Qty_Multi Then
                            vTBLIsi.Rows.Remove(rowInv)
                            Exit Sub
                        ElseIf rowInv.Total_Qty < rowReturn.Total_Qty_Multi Then
                            rowReturn.Total_Qty_Multi -= rowInv.Total_Qty
                            vTBLIsi.Rows.Remove(rowInv)
                        ElseIf rowInv.Total_Qty > rowReturn.Total_Qty_Multi Then
                            rowInv.Total_Qty -= rowReturn.Total_Qty_Multi
                            'hitung ulang amount
                            HargaLama = rowInv.Amount
                            HargaBaru = Math.Round(rowInv.Amount * rowInv.Total_Qty / _
                            (rowInv.Total_Qty + rowReturn.Total_Qty_Multi), DefaultPrec, MidpointRounding.AwayFromZero)
                            rowInv.Amount = HargaBaru
                            Exit Sub
                        End If
                    Else
                        counter += 1
                    End If
                End While
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function CetakInvoice2(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, _
ByVal vprimarysalesInvoice As Double, ByVal vTanggalInvoice As Date, ByVal DecimalHarga As Integer) As Boolean
        Dim report As New CrpCetakTransaksiInvoice2
        Dim DS As New DseTransaksiSalesInvoice
        Dim FormatDay As Integer = 1
        Dim FormatMonth As Integer = 2
        Dim FormatYears As Integer = 1
        Dim FormatDecimal As Integer = 2
        Dim ListStr As New List(Of String)
        Dim rowIsi As DseTransaksiSalesInvoice.TblCetakIsiRow
        Dim rowHeader As DseTransaksiSalesInvoice.TblCetakHeaderRow
        Dim counter As Integer = 0
        Dim Dummy As String = ""
        Dim DTAReturn As New DseTransaksiSalesInvoiceTableAdapters.DataTable9TableAdapter
        Dim TBLReturn As New DseTransaksiSalesInvoice.DataTable9DataTable
        Dim rowReturn As DseTransaksiSalesInvoice.DataTable9Row
        Dim i As New FrmTransaksiSalesInvoice_EditDONumber
        Dim ObjCD As New Core_Module.ClsCekData
        Dim OngkosKirim As Decimal = 0

        Try
            Dim DTAHeader As New DseTransaksiSalesInvoiceTableAdapters.TblCetakHeaderTableAdapter
            DTAHeader.FillBy_PrimaryInvoice(DS.TblCetakHeader, vprimarysalesInvoice)
            DTAHeader.Dispose()
            Dim DTAIsi As New DseTransaksiSalesInvoiceTableAdapters.TblCetakIsiTableAdapter
            DTAIsi.Fill(DS.TblCetakIsi, vprimarysalesInvoice)
            DTAIsi.Dispose()

            'kurangkan isi yang terjadi dengan sales return
            If DTAReturn.FillBy_PrimaryInvoice(TBLReturn, vprimarysalesInvoice) > 0 Then
                Dim DTAReturn_Detail As New dseTransaksiSalesReturnTableAdapters.tbltrans_salesreturn_item_view0TableAdapter
                Dim TBLReturn_Detail As New dseTransaksiSalesReturn.tbltrans_salesreturn_item_view0DataTable
                Dim rowReturnDetail As dseTransaksiSalesReturn.tbltrans_salesreturn_item_view0Row

                For Each rowReturn In TBLReturn.Rows
                    DTAReturn_Detail.FillBy_PrimaryTransaksi(TBLReturn_Detail, rowReturn.Primary_SR)
                    For Each rowReturnDetail In TBLReturn_Detail.Rows
                        Me.KurangkanItemInvoice(rowReturnDetail, DS.TblCetakIsi, DecimalHarga)
                    Next
                Next

                DTAReturn_Detail.Dispose()
                TBLReturn_Detail.Dispose()
            End If

            'update data customer order
            For Each rowIsi In DS.TblCetakIsi.Rows
                If ListStr.Contains(rowIsi.No_DO) = False Then
                    ListStr.Add(rowIsi.No_DO)
                End If
            Next
            rowHeader = DS.TblCetakHeader.Rows(0)
            If rowHeader.Shipment_With_VAT = 1 Then
                rowHeader.ongkos_kirim = 0
                rowHeader.Sub_Total = rowHeader.Sub_Total + OngkosKirim
            End If
            If rowHeader.Currency_Code = "IDR" Then
                rowHeader.Currency_Code = ""
            Else
                If rowHeader.VAT_Dipungut = 1 Then
                    rowHeader.Tax_Amount = rowHeader.Sub_Total * rowHeader.Tax / 100
                End If
                If rowHeader.Shipment_With_VAT = 1 Then
                    rowHeader.Total = rowHeader.Sub_Total + rowHeader.Tax_Amount - rowHeader.Down_Payment
                Else
                    rowHeader.Total = rowHeader.Sub_Total + rowHeader.Tax_Amount + OngkosKirim - rowHeader.Down_Payment
                End If
            End If

            'modifikasi item sehingga untuk product yang sama tetapi batch, unit price, berbeda tidak keluar 2 kali.
            Me.PindahkanItem(DS)

            'tambahkan ongkos kirim
            OngkosKirim = Me.GetOngkosKirim(vprimarysalesInvoice)
            If OngkosKirim <> 0 And Me.IsShipmentWithTax(vprimarysalesInvoice) = True Then
                rowIsi = DS.TblCetakIsi.NewTblCetakIsiRow
                With rowIsi
                    rowIsi.Amount = OngkosKirim
                    .Category_Code = ""
                    .Category_Desc = ""
                    .Cust_PO_No = ""
                    .Description = "Ongkos Kirim"
                    .Discount = 0
                    .Nett_Price = OngkosKirim
                    .No_DO = ""
                    .No_SO = ""
                    .PaymentTermSO = ""
                    .Primary_Category = -1
                    .Primary_DO = -1
                    .Primary_Product = -1
                    .Primary_SO = -1
                    .Primary_Transaksi = -1
                    .Product_Code = "OK"
                    .Qty_36 = 0
                    .Qty_37 = 0
                    .Qty_38 = 0
                    .Qty_39 = 0
                    .Qty_40 = 0
                    .Qty_41 = 0
                    .Qty_42 = 0
                    .Qty_43 = 0
                    .Qty_44 = 0
                    .Qty_45 = 0
                    .Qty_46 = 0
                    .Remarks = ""
                    .Sec_Unit = ""
                    .Secondary_Qty = -1
                    .Size = 0
                    .TglSO = Today
                    .Total_Qty = 1
                    .Unit_Code = "UNIT"
                    .Unit_Price = OngkosKirim
                    .Use_Sizing = 0
                    ._Decimal = 0
                End With
                DS.TblCetakIsi.Rows.Add(rowIsi)
            End If

            rowHeader.No_DO_Buatan_Sendiri = ""
            For Each Dummy In ListStr
                counter += 1
                If rowHeader.IsNo_DO_Buatan_SendiriNull = True Then
                    If counter = ListStr.Count Then
                        rowHeader.No_DO_Buatan_Sendiri = Dummy
                    Else
                        rowHeader.No_DO_Buatan_Sendiri = Dummy.Substring(0, 6)
                    End If
                Else
                    If counter = ListStr.Count Then
                        rowHeader.No_DO_Buatan_Sendiri = rowHeader.No_DO_Buatan_Sendiri + ", " + Dummy
                    Else
                        rowHeader.No_DO_Buatan_Sendiri = rowHeader.No_DO_Buatan_Sendiri + ", " + Dummy.Substring(0, 6)
                    End If
                End If
            Next
            i.TextBox1.Text = rowHeader.No_DO_Buatan_Sendiri
            If i.ShowDialog = DialogResult.OK Then
                rowHeader.No_DO_Buatan_Sendiri = i.TextBox1.Text
                rowHeader.Sub_Total += OngkosKirim
            End If

            'set sising untuk format baru
            Me.SetSizing(DS.TblCetakIsi)

            report.SetDataSource(DS)
            report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, FormatDay)
            report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, FormatMonth)
            report.SetParameterValue(report.Parameter_FormatYears.ParameterFieldName, FormatYears)
            report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, DecimalHarga)
            report.SetParameterValue(report.Parameter_Tanggal_Invoice.ParameterFieldName, vTanggalInvoice)
            report.SetParameterValue(report.Parameter_Tanggal_STR.ParameterFieldName, ObjCD.GetDateString_Indo(vTanggalInvoice))
            report.SetParameterValue(report.Parameter_Jabatan.ParameterFieldName, rowHeader.Jabatan_Signed)
            report.SetParameterValue(report.Parameter_Remarks.ParameterFieldName, rowHeader.Remarks)
            report.SetParameterValue(report.Parameter_Address.ParameterFieldName, Me.GetAddress(DS.TblCetakHeader))

            CRV.ReportSource = report
            CRV.Show()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            i.Dispose()
            DTAReturn.Dispose()
            TBLReturn.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Function GetOngkosKirim(ByVal vPrimaryInv As Double) As Decimal
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice1TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetOngkosKirim(vPrimaryInv).HasValue = True Then
                Return DTA.ScalarQuery_GetOngkosKirim(vPrimaryInv).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsShipmentWithTax(ByVal vPrimaryInvoice As Double) As Boolean
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice1TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_IsShipmentWithVAT(vPrimaryInvoice).HasValue = True Then
                If DTA.ScalarQuery_IsShipmentWithVAT(vPrimaryInvoice).Value = 1 Then
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

    Public Function CetakInvoice(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, _
    ByVal vprimarysalesInvoice As Double, ByVal vTanggalInvoice As Date, ByVal DecimalHarga As Integer, ByVal vUserCode As String) As Boolean
        Dim DS As New DseTransaksiSalesInvoice
        Dim FormatDay As Integer = 1
        Dim FormatMonth As Integer = 2
        Dim FormatYears As Integer = 1
        Dim FormatDecimal As Integer = 2
        Dim ListStr As New List(Of String)
        Dim rowIsi As DseTransaksiSalesInvoice.TblCetakIsiRow
        Dim rowHeader As DseTransaksiSalesInvoice.TblCetakHeaderRow
        Dim counter As Integer = 0
        Dim Dummy As String = ""
        Dim DTAReturn As New DseTransaksiSalesInvoiceTableAdapters.DataTable9TableAdapter
        Dim TBLReturn As New DseTransaksiSalesInvoice.DataTable9DataTable
        Dim rowReturn As DseTransaksiSalesInvoice.DataTable9Row
        Dim i As New FrmTransaksiSalesInvoice_EditDONumber
        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim report As New CrpCetakTransaksiInvoice
        Dim ObjCustom As New CustomSMESS.ClsTransaksiSalesInvoice
        Dim OngkosKirim As Decimal = 0
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim ObjP As New Core_Module.clsMasterProduct

        Try
            'If Me.ApakahUseSizing(vprimarysalesInvoice) = True Then
            '    Return Me.CetakInvoice2(CRV, vprimarysalesInvoice, vTanggalInvoice, DecimalHarga)
            '    Exit Function
            'End If

            If ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.PrintUseCustomInvoice_Bool) = "1" Then
                Return ObjCustom.CetakInvoiceCustom(CRV, vprimarysalesInvoice, vTanggalInvoice, DecimalHarga, report.Parameter_FormatDay.ParameterFieldName, _
                report.Parameter_FormatMonth.ParameterFieldName, report.Parameter_FormatYears.ParameterFieldName, report.Parameter_FormatDecimal.ParameterFieldName, _
                report.Parameter_Tanggal_Invoice.ParameterFieldName, report.Parameter_Tanggal_STR.ParameterFieldName, _
                report.Parameter_Jabatan.ParameterFieldName, report.Parameter_Remarks.ParameterFieldName, vUserCode)
                Exit Function
            End If

            Dim DTAHeader As New DseTransaksiSalesInvoiceTableAdapters.TblCetakHeaderTableAdapter
            DTAHeader.FillBy_PrimaryInvoice(DS.TblCetakHeader, vprimarysalesInvoice)
            DTAHeader.Dispose()
            Dim DTAIsi As New DseTransaksiSalesInvoiceTableAdapters.TblCetakIsiTableAdapter
            DTAIsi.Fill(DS.TblCetakIsi, vprimarysalesInvoice)
            For Each rowIsi In DS.TblCetakIsi.Rows
                rowIsi._Decimal = ObjP.GetAngkaDecimal(rowIsi.Primary_Product, rowIsi.Primary_Units_Multi)
            Next
            DTAIsi.Dispose()


            'kurangkan isi yang terjadi dengan sales return
            If DTAReturn.FillBy_PrimaryInvoice(TBLReturn, vprimarysalesInvoice) > 0 Then
                Dim DTAReturn_Detail As New dseTransaksiSalesReturnTableAdapters.tbltrans_salesreturn_item_view0TableAdapter
                Dim TBLReturn_Detail As New dseTransaksiSalesReturn.tbltrans_salesreturn_item_view0DataTable
                Dim rowReturnDetail As dseTransaksiSalesReturn.tbltrans_salesreturn_item_view0Row

                For Each rowReturn In TBLReturn.Rows
                    DTAReturn_Detail.FillBy_PrimaryTransaksi(TBLReturn_Detail, rowReturn.Primary_SR)
                    For Each rowReturnDetail In TBLReturn_Detail.Rows
                        Me.KurangkanItemInvoice(rowReturnDetail, DS.TblCetakIsi, DecimalHarga)
                    Next
                Next

                DTAReturn_Detail.Dispose()
                TBLReturn_Detail.Dispose()
            End If

            'tambahkan ongkos kirim
            OngkosKirim = Me.GetOngkosKirim(vprimarysalesInvoice)
            If OngkosKirim <> 0 And Me.IsShipmentWithTax(vprimarysalesInvoice) = True Then
                rowIsi = DS.TblCetakIsi.NewTblCetakIsiRow
                With rowIsi
                    rowIsi.Amount = OngkosKirim
                    .Category_Code = ""
                    .Category_Desc = ""
                    .Cust_PO_No = ""
                    .Description = "Shipment"
                    .Discount = 0
                    .Nett_Price = OngkosKirim
                    .No_DO = ""
                    .No_SO = ""
                    .PaymentTermSO = ""
                    .Primary_Category = -1
                    .Primary_DO = -1
                    .Primary_Product = -1
                    .Primary_SO = -1
                    .Primary_Transaksi = -1
                    .Product_Code = "Shipment"
                    .Qty_36 = 0
                    .Qty_37 = 0
                    .Qty_38 = 0
                    .Qty_39 = 0
                    .Qty_40 = 0
                    .Qty_41 = 0
                    .Qty_42 = 0
                    .Qty_43 = 0
                    .Qty_44 = 0
                    .Qty_45 = 0
                    .Qty_46 = 0
                    .Remarks = ""
                    .Sec_Unit = ""
                    .Secondary_Qty = -1
                    .Size = 0
                    .TglSO = Today
                    .Total_Qty = 1
                    .Unit_Code = "UNIT"
                    .Unit_Price = OngkosKirim
                    .Use_Sizing = 0
                    ._Decimal = 0
                    .Alias_Code = ""
                    .Alias_Desc = ""
                    .Primary_Units_Multi = -1
                End With
                DS.TblCetakIsi.Rows.Add(rowIsi)
            End If

            'masukkan kode alias
            For Each rowIsi In DS.TblCetakIsi.Rows
                If rowIsi.IsAlias_CodeNull = False Then
                    If rowIsi.Alias_Code.Trim <> "" Then
                        rowIsi.Product_Code = rowIsi.Alias_Code.Trim
                    End If
                End If
                If rowIsi.IsAlias_DescNull = False Then
                    If rowIsi.Alias_Desc.Trim <> "" Then
                        rowIsi.Description = rowIsi.Alias_Desc.Trim
                    End If
                End If
            Next

            'update data customer order
            For Each rowIsi In DS.TblCetakIsi.Rows
                If ListStr.Contains(rowIsi.No_DO) = False And rowIsi.No_DO <> "" Then
                    ListStr.Add(rowIsi.No_DO)
                End If
            Next


            rowHeader = DS.TblCetakHeader.Rows(0)
            If rowHeader.Shipment_With_VAT = 1 Then
                rowHeader.ongkos_kirim = 0
                rowHeader.Sub_Total = rowHeader.Sub_Total + OngkosKirim
            End If
            If rowHeader.Currency_Code = "IDR" Then
                rowHeader.Currency_Code = ""
            Else

                If rowHeader.VAT_Dipungut = 1 Then
                    Dim decTaxAmount As Decimal = 0
                    decTaxAmount = ((rowHeader.Sub_Total - rowHeader.Down_Payment) * rowHeader.Tax / 100)
                    If rowHeader.Primary_Curr_Tax = 1 Then
                        rowHeader.Tax_Amount = ObjCD.DigitGrouping(ObjCD.DigitGrouping(decTaxAmount * rowHeader.Tax_Exchange_Rate, 0, Core_Module.ClsCekData.EnumTipeRounding.Down), ObjCur.GetDefaultPrecission(rowHeader.Primary_Curr_Tax), Core_Module.ClsCekData.EnumTipeRounding.Down)
                    Else
                        rowHeader.Tax_Amount = ObjCD.DigitGrouping(decTaxAmount, ObjCur.GetDefaultPrecission(rowHeader.Primary_Curr_Tax), Core_Module.ClsCekData.EnumTipeRounding.Down)
                    End If
                End If



                If rowHeader.Shipment_With_VAT = 1 Then
                    If rowHeader.Primary_Curr_Tax = rowHeader.Primary_Currency Then
                        rowHeader.Total = (rowHeader.Sub_Total + rowHeader.Other_Charges + rowHeader.Tax_Amount + rowHeader.Duty_Stamp) - rowHeader.Down_Payment
                    Else
                        rowHeader.Total = rowHeader.Sub_Total + rowHeader.Other_Charges - rowHeader.Down_Payment
                    End If
                Else
                    If rowHeader.Primary_Curr_Tax = rowHeader.Primary_Currency Then
                        rowHeader.Total = (rowHeader.Sub_Total + rowHeader.Other_Charges + rowHeader.Tax_Amount + rowHeader.Duty_Stamp + OngkosKirim) - rowHeader.Down_Payment
                    Else
                        rowHeader.Total = (rowHeader.Sub_Total + rowHeader.Other_Charges + OngkosKirim) - rowHeader.Down_Payment
                    End If
                End If
            End If

            'modifikasi item sehingga untuk product yang sama tetapi batch, unit price, berbeda tidak keluar 2 kali.
            Me.PindahkanItem(DS)

            For Each Dummy In ListStr
                counter += 1
                If rowHeader.IsNo_DO_Buatan_SendiriNull = True Then
                    If counter = ListStr.Count Then
                        rowHeader.No_DO_Buatan_Sendiri = Dummy
                    Else
                        rowHeader.No_DO_Buatan_Sendiri = Dummy.Substring(0, 6)
                    End If
                Else
                    If counter = ListStr.Count Then
                        rowHeader.No_DO_Buatan_Sendiri = rowHeader.No_DO_Buatan_Sendiri + ", " + Dummy
                    Else
                        rowHeader.No_DO_Buatan_Sendiri = rowHeader.No_DO_Buatan_Sendiri + ", " + Dummy.Substring(0, 6)
                    End If
                End If
            Next

            'cust po number
            ListStr.Clear()
            For Each rowIsi In DS.TblCetakIsi.Rows
                If ListStr.Contains(rowIsi.Cust_PO_No) = False And rowIsi.Cust_PO_No <> "" Then
                    ListStr.Add(rowIsi.Cust_PO_No)
                End If
            Next
            counter = 0
            For Each Dummy In ListStr
                counter += 1
                If rowHeader.IsNo_PO_Buatan_SendiriNull = True Then
                    If counter = ListStr.Count Then
                        rowHeader.No_PO_Buatan_Sendiri = Dummy
                    Else
                        rowHeader.No_PO_Buatan_Sendiri = Dummy
                    End If
                Else
                    If counter = ListStr.Count Then
                        rowHeader.No_PO_Buatan_Sendiri = rowHeader.No_PO_Buatan_Sendiri + ", " + Dummy
                    Else
                        rowHeader.No_PO_Buatan_Sendiri = rowHeader.No_PO_Buatan_Sendiri + ", " + Dummy
                    End If
                End If
            Next

            If rowHeader.IsNo_DO_Buatan_SendiriNull = True Then
                rowHeader.No_DO_Buatan_Sendiri = ""
            End If
            i.TextBox1.Text = rowHeader.No_DO_Buatan_Sendiri
            If i.ShowDialog = DialogResult.OK Then
                rowHeader.No_DO_Buatan_Sendiri = i.TextBox1.Text
            Else
                Exit Function
            End If
            If rowHeader.IsNo_PO_Buatan_SendiriNull = True Then
                rowHeader.No_PO_Buatan_Sendiri = ""
            End If

            report.SetDataSource(DS)
            report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, FormatDay)
            report.SetParameterValue(report.Parameter_FormatDecimalIDR.ParameterFieldName, ObjCur.GetDecimal_DefaultCurrency)
            report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, FormatMonth)
            report.SetParameterValue(report.Parameter_FormatYears.ParameterFieldName, FormatYears)
            report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, ObjCur.GetDefaultPrecission(rowHeader.Primary_Currency))
            report.SetParameterValue(report.Parameter_Tanggal_Invoice.ParameterFieldName, vTanggalInvoice)
            report.SetParameterValue(report.Parameter_Tanggal_STR.ParameterFieldName, ObjCD.GetDateString_Indo(vTanggalInvoice))
            report.SetParameterValue(report.Parameter_Jabatan.ParameterFieldName, rowHeader.Jabatan_Signed)
            report.SetParameterValue(report.Parameter_Remarks.ParameterFieldName, rowHeader.Remarks)

            CRV.ReportSource = report
            CRV.Show()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            i.Dispose()
            DTAReturn.Dispose()
            TBLReturn.Dispose()
            ObjCD = Nothing
            ObjSetting = Nothing
            ObjCustom = Nothing
            ObjCur = Nothing
        End Try
    End Function

    Private Function ApakahUseSizing(ByVal vPrimaryInvoice As Double) As Boolean
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_itemTableAdapter

        Try
            If DTA.ScalarQuery_GetJmlUseSizing(vPrimaryInvoice).HasValue = True Then
                If DTA.ScalarQuery_GetJmlUseSizing(vPrimaryInvoice).Value > 0 Then
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


    Private Function GetAddress(ByVal vTBL As DseTransaksiSalesInvoice.TblCetakHeaderDataTable) As String
        Dim row As DseTransaksiSalesInvoice.TblCetakHeaderRow
        Dim dummy As String = ""

        Try
            If vTBL.Rows.Count = 0 Then
                Return ""
            Else
                row = vTBL.Rows(0)
                If row.Line_1 <> "" Then
                    If dummy = "" Then
                        dummy = row.Line_1
                    Else
                        dummy = dummy + vbCrLf + row.Line_1
                    End If
                End If
                If row.Line_2 <> "" Then
                    If dummy = "" Then
                        dummy = row.Line_2
                    Else
                        dummy = dummy + vbCrLf + row.Line_2
                    End If
                End If
                If row.Line_3 <> "" Then
                    If dummy = "" Then
                        dummy = row.Line_3
                    Else
                        dummy = dummy + vbCrLf + row.Line_3
                    End If
                End If
                If row.Line_4 <> "" Then
                    If dummy = "" Then
                        dummy = row.Line_4
                    Else
                        dummy = dummy + vbCrLf + row.Line_4
                    End If
                End If
                If row.Line_5 <> "" Then
                    If dummy = "" Then
                        dummy = row.Line_5
                    Else
                        dummy = dummy + vbCrLf + row.Line_5
                    End If
                End If
            End If
            Return dummy
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Sub SetSizing(ByVal vTBL As DseTransaksiSalesInvoice.TblCetakIsiDataTable)
        Dim TBLDummy As New DseTransaksiSalesInvoice.TblCetakIsiDataTable
        Dim row, row1 As DseTransaksiSalesInvoice.TblCetakIsiRow
        Dim Filter As String = ""

        Try
            While vTBL.Rows.Count > 0
                row = vTBL.Rows(0)

                If row.Use_Sizing = 0 Then
                    'langsung pindahkan apa adanya ke tabel dummy
                    row1 = TBLDummy.NewTblCetakIsiRow

                    With row1
                        .Amount = row.Amount
                        .Category_Code = row.Category_Code
                        .Category_Desc = row.Category_Desc
                        .Cust_PO_No = row.Cust_PO_No
                        .Description = row.Description
                        .Discount = row.Discount
                        .Nett_Price = row.Nett_Price
                        .No_DO = row.No_DO
                        .No_SO = row.No_SO
                        .PaymentTermSO = row.PaymentTermSO
                        .Primary_Product = row.Primary_Product
                        .Primary_SO = row.Primary_SO
                        .Primary_DO = row.Primary_DO
                        .Primary_Category = row.Primary_Category
                        .Primary_Transaksi = row.Primary_Transaksi
                        .Product_Code = row.Product_Code
                        .Qty_36 = 0
                        .Qty_37 = 0
                        .Qty_38 = 0
                        .Qty_39 = 0
                        .Qty_40 = 0
                        .Remarks = row.Remarks
                        .Sec_Unit = row.Sec_Unit
                        .Secondary_Qty = row.Secondary_Qty
                        .Size = row.Size
                        .TglSO = row.TglSO
                        .Total_Qty = row.Total_Qty
                        .Unit_Code = row.Unit_Code
                        .Unit_Price = row.Unit_Price
                        .Use_Sizing = row.Use_Sizing
                        ._Decimal = row._Decimal
                    End With

                    TBLDummy.Rows.Add(row1)
                Else
                    'jika use sizing maka pindahkan berdasarkan category-nya
                    Filter = "[Primary SO] =" + row.Primary_SO.ToString + " and [Primary DO] =" + _
                    row.Primary_DO.ToString + " and [Primary Category] =" + row.Primary_Category.ToString + _
                    " AND [Nett Price] =" + row.Nett_Price.ToString

                    TBLDummy.DefaultView.RowFilter = Filter
                    If TBLDummy.DefaultView.Count > 0 Then
                        row1 = TBLDummy.DefaultView(0).Row
                        'masukkan jumlah qty 
                        row1.Total_Qty += row.Total_Qty

                        'hitung ulang amount
                        row1.Amount = row1.Total_Qty * row1.Nett_Price

                        'masukkan qty sesuai size
                        If row.Size = 36 Then
                            row1.Qty_36 += row.Total_Qty
                        ElseIf row.Size = 37 Then
                            row1.Qty_37 += row.Total_Qty
                        ElseIf row.Size = 38 Then
                            row1.Qty_38 += row.Total_Qty
                        ElseIf row.Size = 39 Then
                            row1.Qty_39 += row.Total_Qty
                        ElseIf row.Size = 40 Then
                            row1.Qty_40 += row.Total_Qty
                        End If
                    Else
                        'masukkan baris baru
                        'langsung pindahkan apa adanya ke tabel dummy
                        row1 = TBLDummy.NewTblCetakIsiRow

                        With row1
                            .Amount = row.Amount
                            .Category_Code = row.Category_Code
                            .Category_Desc = row.Category_Desc
                            .Cust_PO_No = row.Cust_PO_No
                            .Description = row.Category_Desc
                            .Discount = row.Discount
                            .Nett_Price = row.Nett_Price
                            .No_DO = row.No_DO
                            .No_SO = row.No_SO
                            .PaymentTermSO = row.PaymentTermSO
                            .Primary_Product = row.Primary_Product
                            .Primary_SO = row.Primary_SO
                            .Primary_DO = row.Primary_DO
                            .Primary_Category = row.Primary_Category
                            .Primary_Transaksi = row.Primary_Transaksi
                            .Product_Code = row.Category_Code
                            .Qty_36 = 0
                            .Qty_37 = 0
                            .Qty_38 = 0
                            .Qty_39 = 0
                            .Qty_40 = 0
                            .Remarks = row.Remarks
                            .Sec_Unit = row.Sec_Unit
                            .Secondary_Qty = row.Secondary_Qty
                            .Size = row.Size
                            .TglSO = row.TglSO
                            .Total_Qty = row.Total_Qty
                            .Unit_Code = row.Unit_Code
                            .Unit_Price = row.Unit_Price
                            .Use_Sizing = row.Use_Sizing
                            ._Decimal = row._Decimal
                            If row.Size = 36 Then
                                row1.Qty_36 += row.Total_Qty
                            ElseIf row.Size = 37 Then
                                row1.Qty_37 += row.Total_Qty
                            ElseIf row.Size = 38 Then
                                row1.Qty_38 += row.Total_Qty
                            ElseIf row.Size = 39 Then
                                row1.Qty_39 += row.Total_Qty
                            ElseIf row.Size = 40 Then
                                row1.Qty_40 += row.Total_Qty
                            End If
                        End With

                        TBLDummy.Rows.Add(row1)
                    End If
                End If

                vTBL.Rows.Remove(row)
            End While

            'pindahkan hasil tbl dummy kembali ke tbl asal
            While TBLDummy.Rows.Count > 0
                row = TBLDummy.Rows(0)

                'langsung pindahkan apa adanya ke tabel dummy
                row1 = vTBL.NewTblCetakIsiRow

                With row1
                    .Amount = row.Amount
                    .Category_Code = row.Category_Code
                    .Category_Desc = row.Category_Desc
                    .Cust_PO_No = row.Cust_PO_No
                    .Description = row.Description
                    .Discount = row.Discount
                    .Nett_Price = row.Nett_Price
                    .No_DO = row.No_DO
                    .No_SO = row.No_SO
                    .PaymentTermSO = row.PaymentTermSO
                    .Primary_Product = row.Primary_Product
                    .Primary_SO = row.Primary_SO
                    .Primary_DO = row.Primary_DO
                    .Primary_Category = row.Primary_Category
                    .Primary_Transaksi = row.Primary_Transaksi
                    .Product_Code = row.Product_Code
                    .Qty_36 = row.Qty_36
                    .Qty_37 = row.Qty_37
                    .Qty_38 = row.Qty_38
                    .Qty_39 = row.Qty_39
                    .Qty_40 = row.Qty_40
                    .Remarks = row.Remarks
                    .Sec_Unit = row.Sec_Unit
                    .Secondary_Qty = row.Secondary_Qty
                    .Size = row.Size
                    .TglSO = row.TglSO
                    .Total_Qty = row.Total_Qty
                    .Unit_Code = row.Unit_Code
                    .Unit_Price = row.Unit_Price
                    .Use_Sizing = row.Use_Sizing
                    ._Decimal = row._Decimal
                End With

                vTBL.Rows.Add(row1)

                TBLDummy.Rows.Remove(row)
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TBLDummy.Dispose()
        End Try
    End Sub

    Private Sub PindahkanItem(ByVal DS As DseTransaksiSalesInvoice)
        Dim TBLIsi_Dummy As New DseTransaksiSalesInvoice.TblCetakIsiDataTable
        Dim rowIsi_Dummy, rowIsi As DseTransaksiSalesInvoice.TblCetakIsiRow
        Dim Ketemu As Boolean = False

        Try
            For Each rowIsi In DS.TblCetakIsi.Rows
                rowIsi_Dummy = TBLIsi_Dummy.NewTblCetakIsiRow
                With rowIsi_Dummy
                    .Amount = rowIsi.Amount
                    .Cust_PO_No = rowIsi.Cust_PO_No
                    .Description = rowIsi.Description
                    .Discount = rowIsi.Discount
                    .Nett_Price = rowIsi.Nett_Price
                    .No_DO = rowIsi.No_DO
                    .No_SO = rowIsi.No_SO
                    .PaymentTermSO = rowIsi.PaymentTermSO
                    .Primary_Product = rowIsi.Primary_Product
                    .Primary_SO = rowIsi.Primary_SO
                    .Primary_Transaksi = rowIsi.Primary_Transaksi
                    .Product_Code = rowIsi.Product_Code
                    .Remarks = rowIsi.Remarks
                    .Sec_Unit = rowIsi.Sec_Unit
                    .Secondary_Qty = rowIsi.Secondary_Qty
                    .TglSO = rowIsi.TglSO
                    .Total_Qty = rowIsi.Total_Qty
                    .Unit_Code = rowIsi.Unit_Code
                    .Unit_Price = rowIsi.Unit_Price
                    ._Decimal = rowIsi._Decimal
                    .Use_Sizing = rowIsi.Use_Sizing
                    .Size = rowIsi.Size
                    .Category_Code = rowIsi.Category_Code
                    .Category_Desc = rowIsi.Category_Desc
                    .Primary_DO = rowIsi.Primary_DO
                    .Primary_Category = rowIsi.Primary_Category
                    .Primary_Units_Multi = rowIsi.Primary_Units_Multi
                End With
                TBLIsi_Dummy.Rows.Add(rowIsi_Dummy)
            Next
            DS.TblCetakIsi.Rows.Clear()
            For Each rowIsi In TBLIsi_Dummy.Rows
                Ketemu = False
                For Each rowIsi_Dummy In DS.TblCetakIsi.Rows
                    If rowIsi_Dummy.No_SO = rowIsi.No_SO And rowIsi_Dummy.No_DO = rowIsi.No_DO _
                    And rowIsi_Dummy.Primary_Product = rowIsi.Primary_Product _
                    And rowIsi_Dummy.Unit_Price = rowIsi.Unit_Price _
                    And rowIsi_Dummy.Discount = rowIsi.Discount _
                    And rowIsi_Dummy.Nett_Price = rowIsi.Nett_Price Then
                        'tandai sudah ketemu
                        Ketemu = True

                        'modifikasi qty dan harga
                        With rowIsi_Dummy
                            .Total_Qty += rowIsi.Total_Qty
                            .Amount = .Amount + (rowIsi.Amount)
                        End With

                        Exit For
                    End If
                Next

                If Ketemu = False Then
                    rowIsi_Dummy = DS.TblCetakIsi.NewTblCetakIsiRow
                    With rowIsi_Dummy
                        .Amount = rowIsi.Amount
                        .Cust_PO_No = rowIsi.Cust_PO_No
                        .Description = rowIsi.Description
                        .Discount = rowIsi.Discount
                        .Nett_Price = rowIsi.Nett_Price
                        .No_DO = rowIsi.No_DO
                        .No_SO = rowIsi.No_SO
                        .PaymentTermSO = rowIsi.PaymentTermSO
                        .Primary_Product = rowIsi.Primary_Product
                        .Primary_SO = rowIsi.Primary_SO
                        .Primary_Transaksi = rowIsi.Primary_Transaksi
                        .Product_Code = rowIsi.Product_Code
                        .Remarks = rowIsi.Remarks
                        .Sec_Unit = rowIsi.Sec_Unit
                        .Secondary_Qty = rowIsi.Secondary_Qty
                        .TglSO = rowIsi.TglSO
                        .Total_Qty = rowIsi.Total_Qty
                        .Unit_Code = rowIsi.Unit_Code
                        .Unit_Price = rowIsi.Unit_Price
                        ._Decimal = rowIsi._Decimal
                        .Use_Sizing = rowIsi.Use_Sizing
                        .Size = rowIsi.Size
                        .Category_Code = rowIsi.Category_Code
                        .Category_Desc = rowIsi.Category_Desc
                        .Primary_DO = rowIsi.Primary_DO
                        .Primary_Category = rowIsi.Primary_Category
                        .Primary_Units_Multi = rowIsi.Primary_Units_Multi
                    End With
                    DS.TblCetakIsi.Rows.Add(rowIsi_Dummy)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TBLIsi_Dummy.Dispose()
        End Try
    End Sub

    Public Function CetakFakturPajakSederhana(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, _
    ByVal vprimarysalesInvoice As Double, ByVal vNoFakturPajak As String, ByVal DecimalHarga As Integer, _
    ByVal TotalHargaJual As Double, ByVal TotalPPN As Double, ByVal TotalDitagih As Double) As Boolean
        Dim report As New CrpCetakFakturPajakSederhana
        Dim DS As New DseTransaksiSalesInvoice
        Dim FormatDay As Integer = 1
        Dim FormatMonth As Integer = 2
        Dim FormatYears As Integer = 1
        Dim FormatDecimal As Integer = 2
        Dim ListStr As New List(Of String)
        Dim rowIsi As DseTransaksiSalesInvoice.TblCetakIsiRow
        Dim rowHeader As DseTransaksiSalesInvoice.TblCetakHeaderRow
        Dim counter As Integer = 0
        Dim Dummy As String = ""
        Dim DTAReturn As New DseTransaksiSalesInvoiceTableAdapters.DataTable9TableAdapter
        Dim TBLReturn As New DseTransaksiSalesInvoice.DataTable9DataTable
        Dim rowReturn As DseTransaksiSalesInvoice.DataTable9Row
        Dim ObjProd As New Core_Module.clsMasterProduct

        Try
            Dim DTAHeader As New DseTransaksiSalesInvoiceTableAdapters.TblCetakHeaderTableAdapter
            DTAHeader.FillBy_PrimaryInvoice(DS.TblCetakHeader, vprimarysalesInvoice)
            DTAHeader.Dispose()
            Dim DTAIsi As New DseTransaksiSalesInvoiceTableAdapters.TblCetakIsiTableAdapter
            DTAIsi.Fill(DS.TblCetakIsi, vprimarysalesInvoice)
            DTAIsi.Dispose()
            'kurangkan isi yang terjadi dengan sales return
            If DTAReturn.FillBy_PrimaryInvoice(TBLReturn, vprimarysalesInvoice) > 0 Then
                Dim DTAReturn_Detail As New dseTransaksiSalesReturnTableAdapters.tbltrans_salesreturn_item_view0TableAdapter
                Dim TBLReturn_Detail As New dseTransaksiSalesReturn.tbltrans_salesreturn_item_view0DataTable
                Dim rowReturnDetail As dseTransaksiSalesReturn.tbltrans_salesreturn_item_view0Row

                For Each rowReturn In TBLReturn.Rows
                    DTAReturn_Detail.FillBy_PrimaryTransaksi(TBLReturn_Detail, rowReturn.Primary_SR)
                    For Each rowReturnDetail In TBLReturn_Detail.Rows
                        Me.KurangkanItemInvoice(rowReturnDetail, DS.TblCetakIsi, DecimalHarga)
                    Next
                Next

                DTAReturn_Detail.Dispose()
                TBLReturn_Detail.Dispose()
            End If

            'update data customer order
            For Each rowIsi In DS.TblCetakIsi.Rows
                rowIsi._Decimal = ObjProd.GetAngkaDecimal(rowIsi.Primary_Product, rowIsi.Primary_Units_Multi)
                If ListStr.Contains(rowIsi.No_DO) = False Then
                    ListStr.Add(rowIsi.No_DO)
                End If
            Next
            rowHeader = DS.TblCetakHeader.Rows(0)
            For Each Dummy In ListStr
                counter += 1
                If rowHeader.IsNo_DO_Buatan_SendiriNull = True Then
                    If counter = ListStr.Count Then
                        rowHeader.No_DO_Buatan_Sendiri = Dummy
                    Else
                        rowHeader.No_DO_Buatan_Sendiri = Dummy.Substring(0, 6)
                    End If
                Else
                    If counter = ListStr.Count Then
                        rowHeader.No_DO_Buatan_Sendiri = rowHeader.No_DO_Buatan_Sendiri + ", " + Dummy
                    Else
                        rowHeader.No_DO_Buatan_Sendiri = rowHeader.No_DO_Buatan_Sendiri + ", " + Dummy.Substring(0, 6)
                    End If
                End If
            Next

            report.SetDataSource(DS)

            report.SetParameterValue(report.Parameter_DecimalHarga.ParameterFieldName, DecimalHarga)
            report.SetParameterValue(report.Parameter_PPN_Amount.ParameterFieldName, TotalPPN)
            report.SetParameterValue(report.Parameter_Total_Ditagih.ParameterFieldName, TotalDitagih)
            report.SetParameterValue(report.Parameter_TotalHargaJual.ParameterFieldName, TotalHargaJual)
            report.SetParameterValue(report.Parameter_NoFakturPajak.ParameterFieldName, vNoFakturPajak)

            CRV.ReportSource = report
            CRV.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTAReturn.Dispose()
            TBLReturn.Dispose()
            ObjProd = Nothing
        End Try
    End Function

    Public Function CetakInvoiceDP(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, _
    ByVal vprimarysalesInvoice As Double, ByVal vTanggalInvoice As Date, ByVal DecimalHarga As Integer, ByVal vUserCode As String) As Boolean
        Dim report As New CrpCetakTransaksiInvoice
        Dim DS As New DseTransaksiSalesInvoice
        Dim FormatDay As Integer = 1
        Dim FormatMonth As Integer = 2
        Dim FormatYears As Integer = 1
        Dim FormatDecimal As Integer = 2
        Dim ListStr As New List(Of String)
        Dim rowIsi As DseTransaksiSalesInvoice.TblCetakIsiRow
        Dim rowHeader As DseTransaksiSalesInvoice.TblCetakHeaderRow
        Dim ObjCD As New Core_Module.ClsCekData

        Dim Dummy As String = ""
        Dim CustPONumber As String = ""
        Dim ObjSO As New ClsTransaksiSalesOrder
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim ObjCustom As New CustomSMESS.ClsTransaksiSalesInvoice

        Try
            If ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.PrintUseCustomInvoice_Bool) = "1" Then
                Return ObjCustom.CetakInvoiceDPCustom(CRV, vprimarysalesInvoice, vTanggalInvoice, DecimalHarga, vUserCode)
                Exit Function
            End If

            Dim DTAHeader As New DseTransaksiSalesInvoiceTableAdapters.TblCetakHeaderTableAdapter
            DTAHeader.FillBy_PrimaryInvoice(DS.TblCetakHeader, vprimarysalesInvoice)
            DTAHeader.Dispose()

            'isikan data DP
            rowHeader = DS.TblCetakHeader.Rows(0)
            rowIsi = DS.TblCetakIsi.NewTblCetakIsiRow
            rowIsi.Amount = rowHeader.Sub_Total
            rowIsi.Cust_PO_No = rowHeader.Cust_PO_No_DP
            CustPONumber = ObjSO.GetCustPONumber(rowHeader.SO_DP)
            If CustPONumber = "" Then
                rowIsi.Description = "Down payment for SO: " + rowHeader.SO_DP
            Else
                rowIsi.Description = "Down payment for PO: " + CustPONumber
            End If
            'rowIsi.Description = "DP atas sales order: " + rowHeader.SO_DP
            'If rowHeader.Cust_PO_No_DP <> "" Then
            'rowIsi.Description = rowIsi.Description + "(Customer's PO Number: " + _
            'rowHeader.Cust_PO_No_DP(+")")
            'End If
            rowIsi._Decimal = 0
            rowIsi.Discount = 0
            rowIsi.Nett_Price = 0
            rowIsi.No_DO = ""
            rowIsi.No_SO = ""
            rowIsi.PaymentTermSO = ""
            rowIsi.Primary_Transaksi = rowHeader.Primary
            rowIsi.Product_Code = ""
            rowIsi.Remarks = ""
            rowIsi.Sec_Unit = ""
            rowIsi.Secondary_Qty = -1
            rowIsi.TglSO = vTanggalInvoice
            rowIsi.Total_Qty = 0
            rowIsi.Unit_Code = ""
            rowIsi.Unit_Price = 0
            rowIsi.Primary_Product = -1
            rowIsi.Primary_Units_Multi = -1

            rowIsi.Use_Sizing = 0
            rowIsi.Size = 36
            rowIsi.Category_Code = ""
            rowIsi.Category_Desc = ""
            rowIsi.Primary_DO = -1
            rowIsi.Primary_Category = -1

            DS.TblCetakIsi.Rows.Add(rowIsi)

            'update data customer order
            For Each rowIsi In DS.TblCetakIsi.Rows
                If ListStr.Contains(rowIsi.No_DO) = False Then
                    ListStr.Add(rowIsi.No_DO)
                End If
            Next

            rowHeader = DS.TblCetakHeader.Rows(0)
            rowHeader.ongkos_kirim = 0
            rowHeader.Sub_Total = rowHeader.Sub_Total '+ OngkosKirim
            rowHeader.Total = rowHeader.Sub_Total


            If rowHeader.Currency_Code = "IDR" Then
                rowHeader.Currency_Code = ""
            Else
                'If rowHeader.VAT_Dipungut = 1 Then
                '    rowHeader.Tax_Amount = rowHeader.Sub_Total * rowHeader.Tax / 100
                'End If
            End If


            If rowHeader.Primary_Curr_Tax = rowHeader.Primary_Currency Then
                rowHeader.Total += rowHeader.Tax_Amount
                rowHeader.Total += rowHeader.Duty_Stamp
            End If

            For Each Dummy In ListStr
                If rowHeader.IsNo_DO_Buatan_SendiriNull = True Then
                    rowHeader.No_DO_Buatan_Sendiri = Dummy
                Else
                    rowHeader.No_DO_Buatan_Sendiri = rowHeader.No_DO_Buatan_Sendiri + ", " + Dummy
                End If
            Next

            ''update data pajak untuk row header, jika invoice yang terjadi adalah invoice non-IDR
            'If rowHeader.Primary_Currency <> 1 Then
            '    With rowHeader
            '        .Tax_Amount = ObjCD.DigitGrouping(.Sub_Total * .Tax / 100, DecimalHarga)
            '        .Total = .Sub_Total + .Tax_Amount
            '    End With
            'End If

            report.SetDataSource(DS)
            report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, FormatDay)
            report.SetParameterValue(report.Parameter_FormatDecimalIDR.ParameterFieldName, ObjCur.GetDecimal_DefaultCurrency)
            report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, FormatMonth)
            report.SetParameterValue(report.Parameter_FormatYears.ParameterFieldName, FormatYears)
            report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, ObjCur.GetDefaultPrecission(rowHeader.Primary_Currency))
            report.SetParameterValue(report.Parameter_Tanggal_Invoice.ParameterFieldName, vTanggalInvoice)
            report.SetParameterValue(report.Parameter_Tanggal_STR.ParameterFieldName, ObjCD.GetDateString_Indo(vTanggalInvoice))
            report.SetParameterValue(report.Parameter_Jabatan.ParameterFieldName, rowHeader.Jabatan_Signed)
            report.SetParameterValue(report.Parameter_Remarks.ParameterFieldName, rowHeader.Remarks)

            CRV.ReportSource = report
            CRV.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            ObjSO = Nothing
        End Try
    End Function

    Public Function GetNoBaru1(ByVal vTanggalTransaksi As Date) As String
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoiceTableAdapter
        Dim TanggalStart, TanggalEnd As Date
        Dim NoUrut As Double

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'cari tanggal sampai dan dari
            Me.ObjNumbering.GetTanggalDariSampai(TanggalStart, TanggalEnd, Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Invoice, _
            vTanggalTransaksi)

            'cari no urut
            If DTA.ScalarQuery_GetNoUrutanMax(TanggalEnd, TanggalStart).HasValue = True Then
                NoUrut = CStr(DTA.ScalarQuery_GetNoUrutanMax(TanggalEnd, TanggalStart).Value + 1)
                If NoUrut <= 0 Then
                    NoUrut = 1
                End If
            Else
                NoUrut = 1
            End If

            Return Me.ObjNumbering.GetNomorTransaksiBaru1(vTanggalTransaksi, NoUrut, Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Invoice)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    'update ory
    Public Function apakahNoUrutSudahAda(ByVal exPrimary As Double, ByVal noUrut As Integer, ByVal vTanggalTransaksi As Date) As Boolean
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoiceTableAdapter
        Dim TanggalStart, TanggalEnd As Date
        Dim Retval As Nullable(Of Long)
        Try
            Me.ObjNumbering.GetTanggalDariSampai(TanggalStart, TanggalEnd, _
                Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Invoice, vTanggalTransaksi)

            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetJmlBy_UrutanTglDariSampai(exPrimary, noUrut, TanggalStart, TanggalEnd).Value > 0 Then
                Return True
            End If

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return True
        Finally
            DTA.Dispose()
            TanggalStart = Nothing
            TanggalEnd = Nothing
            Retval = Nothing
        End Try
    End Function

    Public Function GetNoBaru_Obs(ByVal vTanggalTransaksi As Date) As String
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoiceTableAdapter
        Dim TanggalStart, TanggalEnd As Date
        Dim NoUrut, Bulan As String
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            'cari tanggal sampai dan dari
            TanggalEnd = ObjCD.GetLastDateOfTheYear(vTanggalTransaksi)
            TanggalStart = ObjCD.GetFisrtDateOfTheYear(vTanggalTransaksi)

            'cari no urut
            If DTA.ScalarQuery_GetNoUrutanMax(TanggalEnd, TanggalStart).HasValue = True Then
                NoUrut = CStr(DTA.ScalarQuery_GetNoUrutanMax(TanggalEnd, TanggalStart).Value + 1)
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
            Return NoUrut + "/" + "INV" + "/" + Bulan + "/" + CStr(vTanggalTransaksi.Year)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Function GetNoFakturPajakStandarBaru(ByVal vTanggalTransaksi As Date) As String
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice1TableAdapter
        Dim TanggalStart, TanggalEnd As Date
        Dim NoUrut As String
        Dim Obj As New ClsTransaksiSalesInvoice_SettingTaxInvoice

        Try
            If vTanggalTransaksi >= #4/1/2013# Then
                Return Obj.GetNoSeriFakturPajakBaru(vTanggalTransaksi)
            Else
                'cari tanggal sampai dan dari
                TanggalEnd = CDate("12/31/" + CStr(vTanggalTransaksi.Year))
                If TanggalEnd > #4/1/2013# Then
                    TanggalEnd = #3/31/2013#
                End If
                TanggalStart = CDate("1/1/" + CStr(vTanggalTransaksi.Year))

                'cari no urut
                If DTA.ScalarQuery_GetUrutanFakturPajakMax(TanggalEnd, TanggalStart).HasValue = True Then
                    NoUrut = CStr(DTA.ScalarQuery_GetUrutanFakturPajakMax(TanggalEnd, TanggalStart).Value + 1)
                    If NoUrut <= 0 Then
                        NoUrut = 1
                    End If
                Else
                    NoUrut = 1
                End If

                'pastikan no urut terdiri atas 8 huruf
                While NoUrut.Length < 8
                    NoUrut = "0" + NoUrut
                End While

                'bulan 
                Return CStr(vTanggalTransaksi.Year).Substring(2) + "." + NoUrut
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            Obj = Nothing
        End Try
    End Function

    Public Overridable Function ApakahSudahAda(ByVal vNoTransaksi As String) As Boolean
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice1TableAdapter

        Try
            If DTA.ScalarQuery_GetJmlByNoInvoice(vNoTransaksi).Value > 0 Then
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

    Public Overridable Function ApakahBolehEdit(ByVal vPrimaryTransaksi As Double, _
    ByVal vSalesStaff As String, ByVal StrictEditting As Boolean, _
    Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice1TableAdapter
        Dim DTA1 As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_itemdpTableAdapter
        Dim ObjCD As New Core_Module.ClsCekData
        Dim OBJCL As New ClsTransaksiPurchaseClosing
        Dim TanggalInvoice As Date = #1/1/1950#
        Dim ObjJurnal As New ClsTransaksiJurnal
        Dim Jurnal As String = ""
        Dim ObjInv_C As New CustomSMESS.ClsTransaksiSalesInvoice
        Dim TBL As New DseTransaksiSalesInvoice.tbltrans_salesinvoice1DataTable
        Dim ObjClosing As New Core_Module.ClsTransaksiJournal1
        'Dim OBJCL As New ClsTransaksiPurchaseClosing

        Try
            '4. Jika sudah lewat tanggal closing
            TanggalInvoice = Me.GetTanggalTransaksi(vPrimaryTransaksi)
            If Me.TipeInvoice = EnumTipeInvoice.SalesInvoice Then
                If ObjCD.GetLasDateOfTheMonth(OBJCL.GetLastClosingPeriod()) >= TanggalInvoice Then
                    MsgBox("The period " + Format(TanggalInvoice, "MMM yyyy") + " has been closed, you cannot edit this data.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                If TanggalInvoice.Date <= ObjClosing.GetClosingDate.Date Then
                    MsgBox("You cannot edit this data because transaction date <= closing date" _
                        , MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'jika sudah melewati batas penguncian transaksi maka tidak bisa edit data
            'msgbox("belum")
            'batasan edit invoice:
            '0. Jika invoice tersebut masih ada. Misal quick invoice dibuka(posisi find) pd form invoice, lalu dihapuskan dari form quick invoice
            '   lalu yg sudah ditampilkan pada form invoice di edit, maka invoce saat itu sudah tidak ada.
            DTA.FillBy_Primary(TBL, vPrimaryTransaksi)
            If TBL.Rows.Count <= 0 Then
                MsgBox("Other user have deleted this data, you cannot edit this data.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            '1. Jika strict editting maka hanya sales staff yang menciptakan yang boleh edit
            Dim ObjUser As New Core_Module.ClsUser
            If ObjUser.ApakahMemilikiHakAkses(vSalesStaff, "FrmTransaksiSalesInvoice_HakAksesEditHapus") = False Then
                If StrictEditting = True Then
                    Dim CreatedBy As String = DTA.ScalarQuery_GetCreatedBy(vPrimaryTransaksi)
                    If CreatedBy <> vSalesStaff Then
                        If TampilkanPesan = True Then
                            MsgBox("Only user with ID:" + CreatedBy + " can edit this data.", _
                            MsgBoxStyle.Information, "Information")
                        End If
                        Return False
                    End If
                End If
            End If

            '2. Jika invoice tersebut adalah invoice DP dan sudah dirujuk oleh invoice lainnya
            If DTA1.ScalarQuery_GetJmlByPrimaryInvoiceDO(vPrimaryTransaksi).Value > 0 Then
                If TampilkanPesan = True Then
                    MsgBox("This invoice is refered as down payment by: " + _
                    CStr(DTA1.ScalarQuery_GetNoInvoice(vPrimaryTransaksi)) + ", you cannot edit this data.", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            '3. Jika sudah pernah dicetak maka, ketika akan melakukan edit harus dengan persetujuan 
            'orang yang memiliki hak khusus

            ''4. Jika sudah melewati tanggal closing maka tidak boleh diedit
            'TanggalInvoice = Me.GetTanggalTransaksi(vPrimaryTransaksi)
            'If ObjCD.GetLasDateOfTheMonth(OBJCL.GetLastClosingPeriod()) >= TanggalInvoice Then
            '    MsgBox("The period " + Format(TanggalInvoice, "MMM yyyy") + " has been closed, you cannot edit this data.", _
            '    MsgBoxStyle.Information, "Information")
            '    Return False
            'End If

            '5. Masih boleh diedit jika belum ada jurnal sales receivement yang merujuk ke jurnal sales ybs
            Jurnal = ObjJurnal.Get_SalesInvoiceJurnal(vPrimaryTransaksi)
            If Jurnal <> "" Then
                If TampilkanPesan = True Then
                    MsgBox("This invoice referred by sales receivement journal: " + Jurnal + _
                    ". You cannot edit this data.", MsgBoxStyle.Information, "Information")
                End If

                Return False
            End If
            'msgbox("belum")

            '6. kalo transaksi custom bilang ok
            If ObjInv_C.ApakahBolehEdit_Custom(vPrimaryTransaksi) = False Then
                Return False
            End If

            'sampe sini balikan true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
            ObjCD = Nothing
            OBJCL = Nothing
            ObjJurnal = Nothing
            ObjInv_C = Nothing
            ObjClosing = Nothing
        End Try
    End Function

    Public Function GetTanggalTransaksi_ByPrimaryDO(ByVal vPrimaryDO As Double) As Date
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_itemTableAdapter

        Try
            If DTA.ScalarQuery_GetPrimaryInvoice_ByPrimaryDO(vPrimaryDO).HasValue = True Then
                Return Me.GetTanggalTransaksi(DTA.ScalarQuery_GetPrimaryInvoice_ByPrimaryDO(vPrimaryDO).Value)
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return #1/1/1950#
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTanggalTransaksi(ByVal vPrimaryInvoice As Double) As Date
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice1TableAdapter

        Try
            If DTA.ScalarQuery_GetTanggal(vPrimaryInvoice).HasValue = True Then
                Return DTA.ScalarQuery_GetTanggal(vPrimaryInvoice).Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function ApakahBolehRevisi(ByVal vPrimaryTransaksi As Double, _
    ByVal vSalesStaff As String, ByVal StrictEditting As Boolean, _
    Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice1TableAdapter
        Dim DTA1 As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_itemdpTableAdapter
        Dim ObjCD As New Core_Module.ClsCekData
        Dim OBJCL As New ClsTransaksiPurchaseClosing
        Dim TanggalInvoice As Date = #1/1/1950#
        Dim ObjJurnal As New ClsTransaksiJurnal
        Dim Jurnal As String = ""
        Dim ObjInv_C As New CustomSMESS.ClsTransaksiSalesInvoice

        Try
            'jika sudah melewati batas penguncian transaksi maka tidak bisa edit data
            'msgbox("belum")
            'batasan edit invoice:
            '1. Jika strict editting maka hanya sales staff yang menciptakan yang boleh edit
            Dim ObjUser As New Core_Module.ClsUser
            If ObjUser.ApakahMemilikiHakAkses(vSalesStaff, "FrmTransaksiSalesInvoice_HakAksesEditHapus") = False Then
                Dim CreatedBy As String = DTA.ScalarQuery_GetCreatedBy(vPrimaryTransaksi)
                If CreatedBy <> vSalesStaff Then
                    If TampilkanPesan = True Then
                        MsgBox("Only user with ID:" + CreatedBy + " can delete this data.", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If

            '2. Jika invoice tersebut adalah invoice DP dan sudah dirujuk oleh invoice lainnya
            If DTA1.ScalarQuery_GetJmlByPrimaryInvoiceDO(vPrimaryTransaksi).Value > 0 Then
                If TampilkanPesan = True Then
                    MsgBox("This invoice is refered as down payment by: " + _
                    CStr(DTA1.ScalarQuery_GetNoInvoice(vPrimaryTransaksi)) + ", you cannot revise this data.", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            '3. Jika sudah pernah dicetak maka, ketika akan melakukan edit harus dengan persetujuan 
            'orang yang memiliki hak khusus

            '4. Jika sudah lewat tanggal closing
            TanggalInvoice = Me.GetTanggalTransaksi(vPrimaryTransaksi)
            If ObjCD.GetLasDateOfTheMonth(OBJCL.GetLastClosingPeriod()) >= TanggalInvoice Then
                MsgBox("The period " + Format(TanggalInvoice, "MMM yyyy") + " has been closed, you cannot revise this data.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            '5. Masih boleh diedit jika belum ada jurnal sales receivement yang merujuk ke jurnal sales ybs
            Jurnal = ObjJurnal.Get_SalesInvoiceJurnal(vPrimaryTransaksi)
            If Jurnal <> "" Then
                If TampilkanPesan = True Then
                    MsgBox("This invoice referred by sales receivement journal: " + Jurnal + _
                    ". You cannot revise this data.", MsgBoxStyle.Information, "Information")
                End If

                Return False
            End If

            'msgbox("belum")
            If ObjInv_C.ApakahBolehRevisiCustom(vPrimaryTransaksi) = False Then
                Return False
            End If

            'sampe sini balikan true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
            ObjCD = Nothing
            OBJCL = Nothing
            ObjJurnal = Nothing
            ObjInv_C = Nothing
        End Try
    End Function

    Public Overridable Function ApakahBolehHapus(ByVal vPrimaryTransaksi As Double, _
    ByVal StrictSalesStaffDelete As Boolean, ByVal vSalesStaff As String, _
    Optional ByVal TampilkanPesan As Boolean = True, Optional ByVal QuickUpdate As Boolean = False) As Boolean
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice1TableAdapter
        Dim DTA1 As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_itemdpTableAdapter
        Dim ObjCD As New Core_Module.ClsCekData
        Dim OBJCL As New ClsTransaksiPurchaseClosing
        Dim TanggalInvoice As Date = #1/1/1950#
        Dim ObjJurnal As New ClsTransaksiJurnal
        Dim Jurnal As String = ""
        Dim ObjInv_C As New CustomSMESS.ClsTransaksiSalesInvoice
        Dim DTA_SR As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_doTableAdapter
        Dim TBL_SR As New DseTransaksiSalesInvoice.tbltrans_doDataTable
        Dim row_SR As DseTransaksiSalesInvoice.tbltrans_doRow
        Dim Dummy As String = ""
        Dim No As Integer = 1
        Dim ObjSR As New ClsTransaksiSalesReturn
        Dim ObjClosing As New Core_Module.ClsTransaksiJournal1

        Try
            DTA_SR.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Connection.ConnectionString = DTA_SR.Connection.ConnectionString
            DTA1.Connection.ConnectionString = DTA_SR.Connection.ConnectionString

            'jika sudah melewati batas penguncian transaksi maka tidak bisa hapus data
            'msgbox("belum")

            'data SO tidak boleh dihapus jika
            '1. Jika strict sales Staff delete maka yang boleh hapus hanya sales staff yang bersangkutan
            Dim CreatedBy As String = DTA.ScalarQuery_GetCreatedBy(vPrimaryTransaksi)
            If CreatedBy <> vSalesStaff Then
                If TampilkanPesan = True Then
                    MsgBox("Only user with ID:" + CreatedBy + " can delete this data.", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            '2. Jika sudah pernah dicetak, maka minta minta otorisasi sebelum menghapus invoice
            'msgbox("belum")

            '3. Tidak bisa dihapus jika sudah dirujuk oleh invoice lain sebagai item DP
            If DTA1.ScalarQuery_GetJmlByPrimaryInvoiceDO(vPrimaryTransaksi).Value > 0 Then
                If TampilkanPesan = True Then
                    MsgBox("This invoice is refered as down payment by: " + _
                    CStr(DTA1.ScalarQuery_GetNoInvoice(vPrimaryTransaksi)) + ", you cannot delete this data.", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            '4. Jika sudah lewat tanggal closing
            TanggalInvoice = Me.GetTanggalTransaksi(vPrimaryTransaksi)
            If Me.TipeInvoice = EnumTipeInvoice.InvoiceDP Then
                If TanggalInvoice.Date <= ObjClosing.GetClosingDate.Date Then
                    MsgBox("You cannot delete this data because transaction date <= closing date" _
                        , MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                If ObjCD.GetLasDateOfTheMonth(OBJCL.GetLastClosingPeriod()) >= TanggalInvoice Then
                    MsgBox("The period " + Format(TanggalInvoice, "MMM yyyy") + " has been closed, you cannot delete this data.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            '5. Masih boleh diedit jika belum ada jurnal sales receivement yang merujuk ke jurnal sales ybs
            Jurnal = ObjJurnal.Get_SalesInvoiceJurnal(vPrimaryTransaksi)
            If Jurnal <> "" Then
                If TampilkanPesan = True Then
                    MsgBox("This invoice referred by sales receivement journal: " + Jurnal + _
                    ". You cannot delete this data.", MsgBoxStyle.Information, "Information")
                End If

                Return False
            End If

            '6. jika sudah ada stock in sales return yang merujuk pada transaksi ini
            DTA_SR.Fill(TBL_SR, vPrimaryTransaksi, CInt(Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn))
            If TBL_SR.Rows.Count > 0 Then
                No = 1
                For Each row_SR In TBL_SR.Rows
                    If Dummy = "" Then
                        Dummy = No.ToString + ". " + row_SR.DO_Number + "(" + row_SR.No_Stock_In_SR + ")"
                    Else
                        Dummy = Dummy + vbCrLf + No.ToString + ". " + row_SR.DO_Number + "(" + row_SR.No_Stock_In_SR + ")"
                    End If
                    No += 1
                Next

                'MsgBox("The following delivery order has sales return, you cannot delete this data: " + vbCrLf + Dummy, _
                'MsgBoxStyle.Information, "Information")
                'Return False

            End If

            '7. kalo transaksi custom bilang ok
            If ObjInv_C.ApakahBolehHapus_Custom(vPrimaryTransaksi, QuickUpdate, vSalesStaff) = False Then
                Return False
            End If

            'nilai balikan true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
            ObjCD = Nothing
            OBJCL = Nothing
            ObjJurnal = Nothing
            ObjInv_C = Nothing
            DTA_SR.Dispose()
            TBL_SR.Dispose()
            ObjSR = Nothing
            ObjClosing = Nothing
        End Try
    End Function

    Public Function ApakahBolehEdit_QuickInvoice(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA_SR As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_doTableAdapter
        Dim TBL_SR As New DseTransaksiSalesInvoice.tbltrans_doDataTable
        Dim row_SR As DseTransaksiSalesInvoice.tbltrans_doRow
        Dim Dummy As String = ""
        Dim No As Integer = 1

        Try
            '1. jika sudah ada stock in sales return yang merujuk pada transaksi ini
            DTA_SR.Fill(TBL_SR, vPrimaryTransaksi, CInt(Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn))
            If TBL_SR.Rows.Count > 0 Then
                No = 1
                For Each row_SR In TBL_SR.Rows
                    If Dummy = "" Then
                        Dummy = No.ToString + ". " + row_SR.DO_Number + "(" + row_SR.No_Stock_In_SR + ")"
                    Else
                        Dummy = Dummy + vbCrLf + No.ToString + ". " + row_SR.DO_Number + "(" + row_SR.No_Stock_In_SR + ")"
                    End If
                    No += 1
                Next

                MsgBox("The following delivery order has sales return, you cannot edit this data: " + vbCrLf + Dummy, _
                MsgBoxStyle.Information, "Information")

                Return False
            End If

            'nilai balikan true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_SR.Dispose()
            TBL_SR.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' mengisikan tabel dengan invoice DP yang belum pernah diinvoice, 
    ''' yang merujuk pada SO yang primary-nya diberikan
    ''' </summary>
    ''' <param name="ListPrimarySO"></param>
    ''' <remarks></remarks>
    Public Sub IsiDgvDP(ByVal ListPrimarySO As List(Of Double), ByVal Dgv As DataGridView, ByVal vprimaryCurrencyInvoice As Double, ByVal vPrimaryInvoice As Double)
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice2TableAdapter
        Dim TBL As New DseTransaksiSalesInvoice.tbltrans_salesinvoice2DataTable
        Dim PrimarySO As Double
        DTA.ClearBeforeFill = True
        Dim row As DseTransaksiSalesInvoice.tbltrans_salesinvoice2Row
        Dim Inv_DP_BedaCurrency As String = ""
        Dim ObjCD As New Core_Module.ClsCekData
        Dim DTAB As New DseTransaksiSalesInvoiceTableAdapters.BATASAN_ADD_DPTableAdapter
        Dim tblB As New DseTransaksiSalesInvoice.BATASAN_ADD_DPDataTable
        Dim rowB As DseTransaksiSalesInvoice.BATASAN_ADD_DPRow

        'add aul
        Dim DTA1 As New DseTransaksiSalesInvoiceTableAdapters.tblbalance_sales_dpTableAdapter
        Dim tbl1 As New DseTransaksiSalesInvoice.tblbalance_sales_dpDataTable
        Dim row1 As DseTransaksiSalesInvoice.tblbalance_sales_dpRow
        Dim MasukkanDP As Boolean = False

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = Core_Module.GetConstring
            DTAB.Connection.ConnectionString = Core_Module.GetConstring
            DTAB.Fill(tblB, vPrimaryInvoice)
            'kosongkan dgv
            Dgv.Rows.Clear()
            'isikan SO
            For Each PrimarySO In ListPrimarySO
                DTA.FillBy_PrimarySO(TBL, PrimarySO)
                For Each row In TBL.Rows
                    If row.Primary_Currency <> vprimaryCurrencyInvoice Then
                        If Inv_DP_BedaCurrency <> "" Then
                            Inv_DP_BedaCurrency = row.DP_Invoice_No
                        Else
                            Inv_DP_BedaCurrency = Inv_DP_BedaCurrency + ", " + row.DP_Invoice_No
                        End If
                    Else
                        '************************************************************************************************************************************
                        'PERUBAHAN INSTRUKSI KO IWAN TGL 31 JULI 2016 
                        'JIKA 1 SO sama dengan 2 do yg berbeda, lalu buat 1 DP, saat add DP harusnya cm 1 invoice yg bisa add
                        '************************************************************************************************************************************
                        With row
                            Dim PrimaryInvoiceDP As Double = -1
                            PrimaryInvoiceDP = Me.GetPrimaryInvoice(.DP_Invoice_No)
                            tblB.DefaultView.RowFilter = Nothing
                            tblB.DefaultView.RowFilter = "[Primary Invoice DP] = " & PrimaryInvoiceDP
                            If tblB.DefaultView.Count = 0 Then
                                Dgv.Rows.Add("0", row.DP_Invoice_No, row.Currency_Code, ObjCD.DigitGrouping(.Sub_Total, .Default_Precision), "", .Primary_Invoice, EnumTipeDP.InvDP)
                            End If
                        End With
                        '************************************************************************************************************************************
                    End If
                Next
            Next

            'add aul
            For Each PrimarySO In ListPrimarySO
                DTA1.Fill(tbl1, PrimarySO)
                For Each row1 In tbl1.Rows
                    MasukkanDP = False

                    With row1
                        If .Jml_DP = 1 Then
                            If .Primary_Invoice = -1 Then
                                'masukkan dp
                                MasukkanDP = True
                            ElseIf .Primary_Invoice = vPrimaryInvoice Then
                                'masukkan dp
                                MasukkanDP = True
                            End If
                        ElseIf .Jml_DP = 0 Then
                            'masukkan dp
                            MasukkanDP = True
                        End If
                        If MasukkanDP = True Then
                            If .primary_currency = vprimaryCurrencyInvoice Then
                                Dgv.Rows.Add("0", .DP_Invoice_No, .Currency_Code, ObjCD.DigitGrouping(.Sub_Total, .Default_Precision), "", .Primary, EnumTipeDP.ManualDP)
                            Else
                                If Inv_DP_BedaCurrency <> "" Then
                                    Inv_DP_BedaCurrency = .DP_Invoice_No
                                Else
                                    Inv_DP_BedaCurrency = Inv_DP_BedaCurrency + ", " + .DP_Invoice_No
                                End If
                            End If
                        End If
                    End With
                Next
            Next

            If Inv_DP_BedaCurrency <> "" Then
                MsgBox("The following downpayment invoice have different currency: " _
                + Inv_DP_BedaCurrency + ". We cannot enter the data.", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjCD = Nothing
        End Try
    End Sub

    Public Function ApakahAdaDPSudahDipakai(ByVal vPrimaryInvoice As Double, ByVal TampilkanPesan As Boolean, ByVal vFrm As FrmTransaksiSalesInvoice) As Boolean
        Dim DTAB As New DseTransaksiSalesInvoiceTableAdapters.BATASAN_ADD_DPTableAdapter
        Dim tblB As New DseTransaksiSalesInvoice.BATASAN_ADD_DPDataTable
        Dim rowB As DseTransaksiSalesInvoice.BATASAN_ADD_DPRow
        Dim rowDGV As New DataGridViewRow

        Try
            DTAB.Connection.ConnectionString = Core_Module.GetConstring
            DTAB.Fill(tblB, vPrimaryInvoice)

            Dim strMsg As String = ""
            For Each rowDGV In vFrm.dgvDP.Rows
                With rowDGV
                    tblB.DefaultView.RowFilter = Nothing
                    tblB.DefaultView.RowFilter = "[Primary Invoice DP] = " & .Cells(vFrm.ColPrimaryInvDP.Index).Value
                    If tblB.DefaultView.Count > 0 Then
                        For IntB As Integer = 0 To tblB.DefaultView.Count - 1
                            rowB = tblB.DefaultView(IntB).Row
                            strMsg &= vbNewLine & " - Invoice DP No " & .Cells(vFrm.ColDPNo.Index).Value & " already used in " & Me.GetNoInvoice(rowB.Primary_Sales_Invoice)
                        Next
                    End If
                End With
            Next


            If strMsg <> "" Then
                If TampilkanPesan = True Then
                    MsgBox("You cannot save data, because there's invoice DP have been redeem in another invoice no, please re-redeem your down payment: " & strMsg, MsgBoxStyle.Information, "Information")
                End If
                Return True
            Else
                Return False
            End If



        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function ApakahDataBenar(ByVal vBaru As Boolean, ByVal vDGV As DataGridView, _
   ByVal vFrm As FrmTransaksiSalesInvoiceQuick, ByVal vNoDO As String, ByVal vPrimaryQuickInv As Double) As Boolean
        Dim ListPrimaryDO As New List(Of Double)
        Dim rowDGV As DataGridViewRow
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.DataTable10TableAdapter
        Dim TBL As New DseTransaksiSalesInvoice.DataTable10DataTable
        Dim row As DseTransaksiSalesInvoice.DataTable10Row
        Dim PrimaryDO As Double = -1
        Dim Inv_dummy As String = ""
        Dim Jurnal As String = ""
        Dim ObjJurnal As New ClsTransaksiJurnal
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim PrimaryAkunDummy As String = ""
        Dim OBJCL As New ClsTransaksiPurchaseClosing
        Dim DTA2 As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice1TableAdapter
        Dim Lanjut As Boolean = False
        Dim PrimaryDummy As Nullable(Of Double)
        Dim DummyNoFakturPajakBaru As String = ""
        Dim Obj As New ClsTransaksiSalesInvoice_SettingTaxInvoice
        Dim JmlAcuan As Integer = 0

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring 
            DTA2.Connection.ConnectionString = DTA.Connection.ConnectionString

            If ObjCD.GetLasDateOfTheMonth(OBJCL.GetLastClosingPeriod()) >= vFrm.dtpDate.Value.Date Then
                MsgBox("The period " + Format(vFrm.dtpDate.Value.Date, "MMM yyyy") + " has been closed, you cannot delete this data.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            'msgbox("belum")
            'data salah bila:
            '1. Sudah ada no transaksi yang sama sebelumnya (untuk transaksi baru dan edit BEDA)
            '       untuk transaksi baru - dan autonumber-langsung ambil nomor baru
            '       untuk transaksi edit - dan autonumber-dianggap sama karena memang no transaksi tidak bisa diedit.
            If vBaru = True Then
                If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Invoice) = True Then
                    If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                        Me.mNoTransaksi = Me.GetNoBaru1(Me.mTanggalTransaksi)
                    End If
                Else
                    If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                        MsgBox("Invoice number duplication, please change invoice number.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            Else
                If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Invoice) = True Then
                    If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                        If Me.GetPrimaryInvoice(Me.mNoTransaksi) <> Me.mPrimary Then
                            Me.mNoTransaksi = Me.GetNoBaru1(Me.mTanggalTransaksi)
                        End If
                    End If
                Else
                    If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                        If Me.GetPrimaryInvoice(Me.mNoTransaksi) <> Me.mPrimary Then
                            MsgBox("Invoice number duplication, please change invoice number.", MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    End If
                End If
            End If

            '2. Jika semua transaksi sales sedang dikunci

            '3. Jika DO yang dinvoicekan sudah diinvoicekan pada transaksi lain
            For Each rowDGV In vDGV.Rows
                PrimaryDO = rowDGV.Cells(vFrm.ColPrimaryDO.Index).Value
                If ListPrimaryDO.Contains(PrimaryDO) = False Then
                    If DTA.Fill(TBL, PrimaryDO) > 0 Then
                        For Each row In TBL.Rows
                            If vBaru = True Then
                                If Inv_dummy = "" Then
                                    Inv_dummy = row.No_Transaksi
                                Else
                                    Inv_dummy = Inv_dummy + "," + row.No_Transaksi
                                End If
                            Else
                                'untuk transaksi edit, tidak boleh memperhitungkan dirinya sendiri
                                If Me.mPrimary <> row.Primary_Invoice Then
                                    If Inv_dummy = "" Then
                                        Inv_dummy = row.No_Transaksi
                                    Else
                                        Inv_dummy = Inv_dummy + "," + row.No_Transaksi
                                    End If
                                End If
                            End If
                        Next
                    End If

                    ListPrimaryDO.Add(PrimaryDO)
                End If

                If Inv_dummy <> "" Then
                    MsgBox("The delivery order: " + rowDGV.Cells(vFrm.ColDONO.Index).Value.ToString + _
                    " have been invoiced in: " + Inv_dummy, MsgBoxStyle.Information, "Information")

                    Return False
                End If
            Next

            '5. Jika sudah ada pembayaran untuk AR ybs
            Jurnal = ObjJurnal.Get_SalesInvoiceJurnal(Me.mPrimary)
            If Jurnal <> "" Then
                MsgBox("This invoice referred by sales receivement journal: " + Jurnal + _
                ". You cannot change this data.", MsgBoxStyle.Information, "Information")

                Return False
            End If

            If Me.TaxAmount > 0 Then
                PrimaryAkunDummy = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_VAT_Out)
                If PrimaryAkunDummy = "" Or PrimaryAkunDummy = "-1" Then
                    MsgBox("Please specify vat out account on global setting.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'jika dutystamp >, maka harus ada akun dutystamp expense
            If Me.mDutyStamp > 0 Then
                PrimaryAkunDummy = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_Dutystamp_Expense)
                If PrimaryAkunDummy = "" Or PrimaryAkunDummy = "-1" Then
                    MsgBox("Please specify stamp duty expense account on global setting.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'jika ada global disc, maka harus ada akun biaya global disc
            If Me.mDiscAmount > 0 Then
                PrimaryAkunDummy = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.GlobalDiscAcc)
                If PrimaryAkunDummy = "" Or PrimaryAkunDummy = "-1" Then
                    MsgBox("Please specify invoice global discount account on global setting.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            '2025 Core Tax Tahap 2.2 
            ''test nomor faktur pajak: double cek: disini dan difungsi simpan baru dan simpan perubahan
            ''perlu coding disini untuk handle masalah ketika sudah tidak ada lagi nomor faktur pajak kemudian user maksa pake yang sudah ada
            'If Me.mJenisFakturPajak = EnumJenisFakturPajak.FakturPajakStandar And Me.UseTax = True Then
            '    If vBaru = True Then
            '        JmlAcuan = 0
            '    Else
            '        JmlAcuan = 1
            '    End If
            '    While DTA2.ScalarQuery_GetJmlByNoFakturPajak(Me.mNoFakturPajak).Value > JmlAcuan
            '        PrimaryDummy = DTA2.ScalarQuery_GetPrimaryMaxByNoFakturPajak(Me.mNoFakturPajak)
            '        If PrimaryDummy.HasValue = True Then
            '            If vBaru = True Then
            '                If PrimaryDummy.Value <> -1 Then
            '                    Lanjut = True
            '                Else
            '                    Lanjut = False
            '                End If
            '            Else
            '                If PrimaryDummy.Value = Me.mPrimary Then
            '                    Lanjut = False
            '                Else
            '                    Lanjut = True
            '                End If
            '            End If
            '        Else
            '            'jika tidak ada dalam db, maka ga perlu cek lebih jauh
            '            Lanjut = False
            '        End If

            '        If Lanjut = True Then
            '            'get no baru
            '            DummyNoFakturPajakBaru = Me.GetNoFakturPajakStandarBaru(Me.mTanggalTransaksi)
            '            If DummyNoFakturPajakBaru <> "" Then
            '                Me.mNoFakturPajak = DummyNoFakturPajakBaru
            '                'Exit While

            '                'update no-nya
            '                If Me.mTanggalTransaksi >= Obj.BatasStartDate Then
            '                    Me.mUrutanFakturPajak = Obj.GetUrutan(Me.mNoFakturPajak, Me.mTanggalTransaksi)
            '                Else
            '                    If Me.mJenisFakturPajak = EnumJenisFakturPajak.FakturPajakStandar Then
            '                        Me.mUrutanFakturPajak = Me.NoFakturPajak.Substring(4)
            '                    Else
            '                        Me.mUrutanFakturPajak = 0
            '                        Me.mNoSeriFakturPajak = ""
            '                        Me.mNoFakturPajak = ""
            '                    End If
            '                End If
            '            Else
            '                MsgBox("There are no new tax invoice number available. Please specify  new tax invoice number in your global setting.", MsgBoxStyle.Information, "Information")
            '                Return False
            '            End If
            '        Else
            '            Exit While
            '        End If
            '    End While
            'End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimaryDO = Nothing
            DTA.Dispose()
            TBL.Dispose()
            ObjSetting = Nothing
            OBJCL = Nothing
            DTA2.Dispose()
        End Try
    End Function

    Private Function ApakahDataBenar(ByVal vBaru As Boolean, ByVal vDGV As DataGridView, ByVal vFrm As FrmTransaksiSalesInvoice, _
    Optional ByVal vCekDataBenar As Boolean = True, Optional ByVal vFromCutOff As Boolean = True) As Boolean
        If vCekDataBenar = False Then
            Return True
        End If

        Dim ListPrimaryDO As New List(Of Double)
        Dim rowDGV As DataGridViewRow
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.DataTable10TableAdapter
        Dim TBL As New DseTransaksiSalesInvoice.DataTable10DataTable
        Dim row As DseTransaksiSalesInvoice.DataTable10Row
        Dim PrimaryDO As Double = -1
        Dim Inv_dummy As String = ""
        Dim Urutan As Double = 0
        Dim PrimaryInvoiceRujukan As Double = -1
        Dim DTA1 As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoiceTableAdapter
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim PrimaryAkunDummy As String = ""
        Dim OBJCL As New ClsTransaksiPurchaseClosing
        Dim DTA2 As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice1TableAdapter
        Dim Lanjut As Boolean = False
        Dim PrimaryDummy As Nullable(Of Double)
        Dim DummyNoFakturPajakBaru As String = ""
        Dim Obj As New ClsTransaksiSalesInvoice_SettingTaxInvoice
        Dim JmlAcuan As Integer = 0

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA2.Connection.ConnectionString = DTA.Connection.ConnectionString

            '4. Jika sudah lewat tanggal closing
            If ObjCD.GetLasDateOfTheMonth(OBJCL.GetLastClosingPeriod()) >= vFrm.dtpDate.Value.Date Then
                MsgBox("The period " + Format(vFrm.dtpDate.Value.Date, "MMM yyyy") + " has been closed, you cannot delete this data.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            'msgbox("belum")
            'data salah bila:
            '1. Sudah ada no transaksi yang sama sebelumnya (untuk transaksi baru dan edit BEDA)
            '       untuk transaksi baru - dan autonumber-langsung ambil nomor baru
            '       untuk transaksi edit - dan autonumber-dianggap sama karena memang no transaksi tidak bisa diedit.
            If vBaru = True Then
                If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Invoice) = True Then
                    If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                        Me.mNoTransaksi = Me.GetNoBaru1(Me.mTanggalTransaksi)
                    End If
                Else
                    If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                        MsgBox("Invoice number duplication, please change invoice number.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            Else
                If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Invoice) = True Then
                    If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                        If Me.GetPrimaryInvoice(Me.mNoTransaksi) <> Me.mPrimary Then
                            Me.mNoTransaksi = Me.GetNoBaru1(Me.mTanggalTransaksi)
                        End If
                    End If
                Else
                    If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                        If Me.GetPrimaryInvoice(Me.mNoTransaksi) <> Me.mPrimary Then
                            MsgBox("Invoice number duplication, please change invoice number.", MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    End If
                End If
            End If

            '2. Jika DO yang dinvoicekan sudah diinvoicekan pada transaksi lain
            For Each rowDGV In vDGV.Rows
                PrimaryDO = rowDGV.Cells(vFrm.ColPrimaryDO.Index).Value
                If ListPrimaryDO.Contains(PrimaryDO) = False Then
                    If DTA.Fill(TBL, PrimaryDO) > 0 Then
                        For Each row In TBL.Rows
                            If vBaru = True Then
                                If Inv_dummy = "" Then
                                    Inv_dummy = row.No_Transaksi
                                Else
                                    Inv_dummy = Inv_dummy + "," + row.No_Transaksi
                                End If
                            Else
                                'untuk transaksi edit, tidak boleh memperhitungkan dirinya sendiri
                                If Me.mPrimary <> row.Primary_Invoice Then
                                    If Inv_dummy = "" Then
                                        Inv_dummy = row.No_Transaksi
                                    Else
                                        Inv_dummy = Inv_dummy + "," + row.No_Transaksi
                                    End If
                                End If
                            End If
                        Next
                    End If

                    ListPrimaryDO.Add(PrimaryDO)
                End If

                If Inv_dummy <> "" Then
                    MsgBox("The delivery order: " + rowDGV.Cells(vFrm.ColDONO.Index).Value.ToString + _
                    " have been invoiced in: " + Inv_dummy, MsgBoxStyle.Information, "Information")

                    Return False
                End If
            Next

            'jika tidak ada akun vat out maka tidak bisa simpan
            If Me.TaxAmount > 0 Then
                PrimaryAkunDummy = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_VAT_Out)
                If PrimaryAkunDummy = "" Or PrimaryAkunDummy = "-1" Then
                    MsgBox("Please specify vat out account on global setting.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'jika dutystamp >, maka harus ada akun dutystamp expense
            If Me.mDutyStamp > 0 Then
                PrimaryAkunDummy = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_Dutystamp_Expense)
                If PrimaryAkunDummy = "" Or PrimaryAkunDummy = "-1" Then
                    MsgBox("Please specify stamp duty expense account on global setting.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'jika shipment <> 0, maka harus ada akun dutystamp expense
            If Me.mOngkosKirim > 0 Then
                PrimaryAkunDummy = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_ShippingIncome)
                If PrimaryAkunDummy = "" Or PrimaryAkunDummy = "-1" Then
                    MsgBox("Please specify income from shipment account on global setting.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'jika ada global disc, maka harus ada akun biaya global disc
            If Me.mDiscAmount > 0 Then
                PrimaryAkunDummy = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.GlobalDiscAcc)
                If PrimaryAkunDummy = "" Or PrimaryAkunDummy = "-1" Then
                    MsgBox("Please specify invoice global discount account on global setting.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            If Me.ApakahAdaDPSudahDipakai(Me.mPrimary, True, vFrm) = True Then
                Return False
            End If

            '2025 Core Tax Tahap 2.2 
            ''test nomor faktur pajak: double cek: disini dan difungsi simpan baru dan simpan perubahan
            ''perlu coding disini untuk handle masalah ketika sudah tidak ada lagi nomor faktur pajak kemudian user maksa pake yang sudah ada
            'If vFromCutOff = False Then
            '    If Me.mJenisFakturPajak = EnumJenisFakturPajak.FakturPajakStandar And Me.UseTax = True Then
            '        If vBaru = True Then
            '            JmlAcuan = 0
            '        Else
            '            JmlAcuan = 1
            '        End If
            '        While DTA2.ScalarQuery_GetJmlByNoFakturPajak(Me.mNoFakturPajak).Value > JmlAcuan
            '            PrimaryDummy = DTA2.ScalarQuery_GetPrimaryMaxByNoFakturPajak(Me.mNoFakturPajak)
            '            If PrimaryDummy.HasValue = True Then
            '                If vBaru = True Then
            '                    If PrimaryDummy.Value <> -1 Then
            '                        Lanjut = True
            '                    Else
            '                        Lanjut = False
            '                    End If
            '                Else
            '                    If PrimaryDummy.Value = Me.mPrimary Then
            '                        Lanjut = False
            '                    Else
            '                        Lanjut = True
            '                    End If
            '                End If
            '            Else
            '                'jika tidak ada dalam db, maka ga perlu cek lebih jauh
            '                Lanjut = False
            '            End If

            '            If Lanjut = True Then
            '                'get no baru
            '                DummyNoFakturPajakBaru = Me.GetNoFakturPajakStandarBaru(Me.mTanggalTransaksi)
            '                If DummyNoFakturPajakBaru <> "" Then
            '                    Me.mNoFakturPajak = DummyNoFakturPajakBaru
            '                    'Exit While

            '                    'update no-nya
            '                    If Me.mTanggalTransaksi >= Obj.BatasStartDate Then
            '                        Me.mUrutanFakturPajak = Obj.GetUrutan(Me.mNoFakturPajak, Me.mTanggalTransaksi)
            '                    Else
            '                        If Me.mJenisFakturPajak = EnumJenisFakturPajak.FakturPajakStandar Then
            '                            Me.mUrutanFakturPajak = Me.NoFakturPajak.Substring(4)
            '                        Else
            '                            Me.mUrutanFakturPajak = 0
            '                            Me.mNoSeriFakturPajak = ""
            '                            Me.mNoFakturPajak = ""
            '                        End If
            '                    End If
            '                Else
            '                    MsgBox("There are no new tax invoice number available. Please specify  new tax invoice number in your global setting.", MsgBoxStyle.Information, "Information")
            '                    Return False
            '                End If
            '            Else
            '                Exit While
            '            End If
            '        End While
            '    End If
            'End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimaryDO = Nothing
            DTA.Dispose()
            TBL.Dispose()
            ObjSetting = Nothing
            OBJCL = Nothing
            DTA1.Dispose()
            DTA2.Dispose()
        End Try
    End Function

    Private Function ApakahDataBenar_DP(ByVal vBaru As Boolean, ByVal vTglTransaksi As Date, _
    Optional ByVal vCekDataBenar As Boolean = True, _
    Optional ByVal vFromCutOff As Boolean = False) As Boolean
        Dim ListPrimaryDO As New List(Of Double)
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.DataTable10TableAdapter
        Dim TBL As New DseTransaksiSalesInvoice.DataTable10DataTable
        Dim PrimaryDO As Double = -1
        Dim Inv_dummy As String = ""
        Dim Jurnal As String = ""
        Dim ObjJurnal As New ClsTransaksiJurnal
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim PrimaryAkunDummy As String = ""
        Dim OBJCL As New ClsTransaksiPurchaseClosing
        Dim DTA2 As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice1TableAdapter
        Dim Lanjut As Boolean = False
        Dim PrimaryDummy As Nullable(Of Double)
        Dim DummyNoFakturPajakBaru As String = ""
        Dim Obj As New ClsTransaksiSalesInvoice_SettingTaxInvoice
        Dim JmlAcuan As Integer = 0
        Dim ObjClosing As New Core_Module.ClsTransaksiJournal1

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA2.Connection.ConnectionString = DTA.Connection.ConnectionString

            If vCekDataBenar = False Then
                Return True
            End If

            '2018-07-09: iwan, diganti ke journal locking, karena dp tidak mempengaruhi inventory
            'If ObjCD.GetLasDateOfTheMonth(OBJCL.GetLastClosingPeriod()) >= vTglTransaksi Then
            '    MsgBox("The period " + Format(vTglTransaksi, "MMM yyyy") + " has been closed, you cannot save this data.", _
            '    MsgBoxStyle.Information, "Information")
            '    Return False
            'End If
            If vTglTransaksi.Date <= ObjClosing.GetClosingDate.Date Then
                MsgBox("You cannot save this data because transaction date <= closing date" _
                    , MsgBoxStyle.Information, "Information")
                Return False
            End If

            'msgbox("belum")
            'data salah bila:
            '1. Sudah ada no transaksi yang sama sebelumnya (untuk transaksi baru dan edit BEDA)
            '       untuk transaksi baru - dan autonumber-langsung ambil nomor baru
            '       untuk transaksi edit - dan autonumber-dianggap sama karena memang no transaksi tidak bisa diedit.
            If vBaru = True Then
                If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Invoice) = True Then
                    If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                        Me.mNoTransaksi = Me.GetNoBaru1(Me.mTanggalTransaksi)
                    End If
                Else
                    If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                        MsgBox("Invoice number duplication, please change invoice number.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            Else
                If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Invoice) = True Then
                    If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                        If Me.GetPrimaryInvoice(Me.mNoTransaksi) <> Me.mPrimary Then
                            Me.mNoTransaksi = Me.GetNoBaru1(Me.mTanggalTransaksi)
                        End If
                    End If
                Else
                    If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                        If Me.GetPrimaryInvoice(Me.mNoTransaksi) <> Me.mPrimary Then
                            MsgBox("Invoice number duplication, please change invoice number.", MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    End If
                End If
            End If

            '2. Jika semua transaksi sales sedang dikunci

            '3. Jika DO yang dinvoicekan sudah diinvoicekan pada transaksi lain

            '4. jika tidak ada item yang akan disimpan

            '5. Jika sudah ada pembayaran untuk AR ybs
            Jurnal = ObjJurnal.Get_SalesInvoiceJurnal(Me.mPrimary)
            If Jurnal <> "" Then
                MsgBox("This invoice referred by sales receivement journal: " + Jurnal + _
                ". You cannot change this data.", MsgBoxStyle.Information, "Information")

                Return False
            End If

            If Me.TaxAmount > 0 Then
                PrimaryAkunDummy = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_VAT_Out)
                If PrimaryAkunDummy = "" Or PrimaryAkunDummy = "-1" Then
                    MsgBox("Please specify vat out account on global setting.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'jika dutystamp >, maka harus ada akun dutystamp expense
            If Me.mDutyStamp > 0 Then
                PrimaryAkunDummy = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_Dutystamp_Expense)
                If PrimaryAkunDummy = "" Or PrimaryAkunDummy = "-1" Then
                    MsgBox("Please specify stamp duty expense account on global setting.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            '2025 Core Tax Tahap 2.2 
            ''test nomor faktur pajak: double cek: disini dan difungsi simpan baru dan simpan perubahan
            ''perlu coding disini untuk handle masalah ketika sudah tidak ada lagi nomor faktur pajak kemudian user maksa pake yang sudah ada
            'If Me.mJenisFakturPajak = EnumJenisFakturPajak.FakturPajakStandar And Me.UseTax = True Then
            '    If vBaru = True Then
            '        JmlAcuan = 0
            '    Else
            '        JmlAcuan = 1
            '    End If
            '    While DTA2.ScalarQuery_GetJmlByNoFakturPajak(Me.mNoFakturPajak).Value > JmlAcuan
            '        PrimaryDummy = DTA2.ScalarQuery_GetPrimaryMaxByNoFakturPajak(Me.mNoFakturPajak)
            '        If PrimaryDummy.HasValue = True Then
            '            If vBaru = True Then
            '                If PrimaryDummy.Value <> -1 Then
            '                    Lanjut = True
            '                Else
            '                    Lanjut = False
            '                End If
            '            Else
            '                If PrimaryDummy.Value = Me.mPrimary Then
            '                    Lanjut = False
            '                Else
            '                    Lanjut = True
            '                End If
            '            End If
            '        Else
            '            'jika tidak ada dalam db, maka ga perlu cek lebih jauh
            '            Lanjut = False
            '        End If

            '        If Lanjut = True Then
            '            'get no baru
            '            DummyNoFakturPajakBaru = Me.GetNoFakturPajakStandarBaru(Me.mTanggalTransaksi)
            '            If DummyNoFakturPajakBaru <> "" Then
            '                Me.mNoFakturPajak = DummyNoFakturPajakBaru
            '                'Exit While

            '                'update no-nya
            '                If Me.mTanggalTransaksi >= Obj.BatasStartDate Then
            '                    Me.mUrutanFakturPajak = Obj.GetUrutan(Me.mNoFakturPajak, Me.mTanggalTransaksi)
            '                Else
            '                    If Me.mJenisFakturPajak = EnumJenisFakturPajak.FakturPajakStandar Then
            '                        Me.mUrutanFakturPajak = Me.NoFakturPajak.Substring(4)
            '                    Else
            '                        Me.mUrutanFakturPajak = 0
            '                        Me.mNoSeriFakturPajak = ""
            '                        Me.mNoFakturPajak = ""
            '                    End If
            '                End If
            '            Else
            '                MsgBox("There are no new tax invoice number available. Please specify  new tax invoice number in your global setting.", MsgBoxStyle.Information, "Information")
            '                Return False
            '            End If
            '        Else
            '            Exit While
            '        End If
            '    End While
            'End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimaryDO = Nothing
            DTA.Dispose()
            TBL.Dispose()
            ObjJurnal = Nothing
            ObjSetting = Nothing
            OBJCL = Nothing
            ObjClosing = Nothing
        End Try
    End Function


    Private Function SimpanBaruSO(ByVal vUserCode As String, ByVal vDGVItem As DataGridView, _
    ByVal vFrm As FrmTransaksiSalesInvoiceQuick, ByVal vIdentifierCust As String, ByVal ObjSO As ClsTransaksiSalesOrder, _
    ByVal vPrimarySalesPerson As Double, ByVal vExchangeRate As Decimal, Optional ByVal vCekDataBenar As Boolean = True, _
    Optional ByVal vSetSOB As Boolean = True, Optional ByVal vFromCutOff As Boolean = False, Optional ByVal vNoSO As String = "") As Double
        Dim ij As New FrmTransaksiSO
        Dim rowDGV As DataGridViewRow
        Dim rowSO As DataGridViewRow
        Dim ObjNumbering As New Core_Module.ClsMasterTransactionNumbering

        Try
            rowSO = vDGVItem.Rows(0)
            With ObjSO
                .PrimaryCurrency = rowSO.Cells(vFrm.ColPrimaryCurrency.Index).Value
                'salin data header
                .CustContactPerson = ""
                .PrimaryCustomer = Me.PrimaryCustomer
                .CustIdentifierAdd = vIdentifierCust
                .CustPhone = ""
                .CustPONo = ""
                .DiskonSama = False
                .JenisFakturPajak = ClsTransaksiSalesOrder.EnumJenisFakturPajak.FakturPajakSederhana
                If ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.SO) = True Then
                    If vNoSO = "" Then '20190711
                        .NoTransaksi = .GetNoBaru1(Me.TanggalTransaksi)
                    Else
                        .NoTransaksi = vNoSO
                    End If
                Else
                    .NoTransaksi = Me.NoTransaksi
                End If
                .PrimarySalesPerson = vPrimarySalesPerson
                .ExchangeRate = vExchangeRate
                .NoUrut = -1
                .OrderBy = ClsTransaksiSalesOrder.EnumOrderBy.PO
                .PaymentDue = 0
                .PaymentTerm = ""
                .PriceCondition = ClsTransaksiSalesOrder.EnumPriceCondition.Exclude
                .PriceConditionSama = False
                .Primary = -1
                .ReferenceDiscount = 0
                .RefPaymentDate = ClsTransaksiSalesOrder.EnumRefPaymentDate.PODate
                .Remarks = ""
                .ShipVia = ""
                .SubTotal = Me.SubTotal - Me.OngkosKirim
                .TanggalTransaksi = Me.TanggalTransaksi
                .TaxAmount = Me.TaxAmount
                .TaxPercentage = Me.TaxPercentage
                .Total = Me.Total - Me.OngkosKirim - Me.DutyStamp - Me.OtherCharge
                .UseTax = Me.UseTax
                If vFromCutOff = True Then
                    .CreatedBy = Me.mCreatedBy
                    .CreatedDate = Me.mCreatedDate
                    .ModifiedBy = Me.mModifiedBy
                    .ModifiedDate = Me.mModifiedDate
                    .TimeStamp = Me.mTimestamp
                End If
                .IdentifierAddressShipTo = vFrm.TxtIdentifierAddressShipTo.Text
                .PrimaryCustomerShipTo = vFrm.TxtPrimaryCustomerShipTo.Text
                .IdentifierAddressCustomerShipTo = vFrm.TxtIdentifierAddressCustomerShipTo.Text

                'salin data item
                ij.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                ij.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                ij.dgvItem.Rows.Clear()
                For Each rowDGV In vDGVItem.Rows
                    ij.dgvItem.Rows.Add()
                    rowSO = ij.dgvItem.Rows(ij.dgvItem.Rows.Count - 1)
                    With rowSO
                        .Cells(ij.ColNo.Index).Value = ""
                        .Cells(ij.ColCodeNo.Index).Value = ""
                        .Cells(ij.ColDescription.Index).Value = ""
                        .Cells(ij.ColSupplierCodeNo.Index).Value = ""
                        .Cells(ij.ColSuppDesc.Index).Value = ""
                        .Cells(ij.ColQty.Index).Value = rowDGV.Cells(vFrm.ColTotalQty.Index).Value
                        .Cells(ij.ColUnit.Index).Value = ""
                        .Cells(ij.ColUnitPrice.Index).Value = rowDGV.Cells(vFrm.ColUnitPrice.Index).Value
                        .Cells(ij.ColDiscount.Index).Value = rowDGV.Cells(vFrm.ColDiscount.Index).Value
                        .Cells(ij.ColNettPrice.Index).Value = rowDGV.Cells(vFrm.ColNettPrice.Index).Value
                        .Cells(ij.ColAmount.Index).Value = rowDGV.Cells(vFrm.ColAmount.Index).Value
                        .Cells(ij.ColPriceCond.Index).Value = "Exclude"
                        .Cells(ij.ColDeliveryCont.Index).Value = False
                        .Cells(ij.ColDeliveryDate.Index).Value = "-"
                        .Cells(ij.ColRemark.Index).Value = ""
                        .Cells(ij.ColPrimaryCodeNo.Index).Value = rowDGV.Cells(vFrm.ColPrimaryCodeNo.Index).Value
                        .Cells(ij.ColPrimaryUnits.Index).Value = rowDGV.Cells(vFrm.ColPrimaryUnit.Index).Value
                        .Cells(ij.ColUseSecondary.Index).Value = False
                        .Cells(ij.ColIdentifier.Index).Value = "-1" + ij.dgvItem.RowCount - 1
                        .Cells(ij.ColDiscAmount.Index).Value = rowDGV.Cells(vFrm.ColDiscAmount.Index).Value
                        .Cells(ij.ColDiscFromPerct.Index).Value = rowDGV.Cells(vFrm.ColDiscFromPrect.Index).Value

                        .Cells(ij.ColQtyMulti.Index).Value = rowDGV.Cells(vFrm.ColTotalQtyMulti.Index).Value
                        .Cells(ij.ColPrimaryUnitsMulti.Index).Value = rowDGV.Cells(vFrm.ColPrimaryUnitsMulti.Index).Value
                        .Cells(ij.ColRatioDefault.Index).Value = rowDGV.Cells(vFrm.ColRatioDefault.Index).Value
                        .Cells(ij.ColRatioReference.Index).Value = rowDGV.Cells(vFrm.ColRatioReference.Index).Value
                        .Cells(ij.ColToleranceDown.Index).Value = "0"
                        .Cells(ij.ColToleranceDownFromPerc.Index).Value = "1"
                        .Cells(ij.ColToleranceDownPerc.Index).Value = "0"
                        .Cells(ij.ColToleranceUp.Index).Value = "0"
                        .Cells(ij.ColToleranceUpFromPerc.Index).Value = "1"
                        .Cells(ij.ColToleranceUpPerc.Index).Value = "0"
                        .Cells(ij.ColPriceByMultiUnit_Str.Index).Value = rowDGV.Cells(vFrm.ColPriceByMultiUnit_Str.Index).Value
                        .Cells(ij.ColPriceByMultiUnit.Index).Value = rowDGV.Cells(vFrm.ColPriceByMultiUnit.Index).Value

                        '.Cells(ij.ColIdentifier.Index).Value = ij.GetIdentifierBaru
                    End With
                Next
                'simpan data baru
                .SimpanBaru(ij.dgvItem, vUserCode, ij, ij.vTBLItemDelivery, vCekDataBenar, vFromCutOff)
            End With

            Return ObjSO.Primary
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            ij = Nothing
            ObjNumbering = Nothing
        End Try
    End Function

    Private Function SimpanBaruDO(ByVal vUserCode As String, ByVal vDGVItem As DataGridView, _
    ByVal vFrm As FrmTransaksiSalesInvoiceQuick, ByVal vPrimarySO As Double, ByVal CustIdentifier As String, ByVal ObjDO As ClsTransaksiDO, _
    ByVal vSONumber As String, ByVal vDONumber As String, Optional ByVal vCekDataBenar As Boolean = False, _
    Optional ByVal vSetSOB As Boolean = True, Optional ByVal vFromCutOff As Boolean = False) As Double
        Dim ij As New FrmTransaksiDO3
        Dim rowDO, rowInv As DataGridViewRow
        Dim ObjNumbering As New Core_Module.ClsMasterTransactionNumbering

        Try
            With ObjDO
                'salin data header
                .Primary = -1
                .AuthorizationLevel = 1
                .AuthorizationLevelNeeded = 1
                .CetakFaktur = False
                .CetakKwitansi = True
                .CustAddIdentifier = CustIdentifier
                If ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._DO) = True Then
                    If vDONumber = "" Then
                        .NoTransaksi = ObjDO.GetNoBaru1(Me.TanggalTransaksi)
                    Else
                        .NoTransaksi = vDONumber '20190711
                    End If
                Else
                    .NoTransaksi = vDONumber
                End If
                .NoUrut = 0
                .Primary = -1
                .PrimaryCustomer = Me.PrimaryCustomer
                .PrimaryStockOutSales = -1
                .Remarks = ""
                .TanggalTransaksi = Me.TanggalTransaksi
                If vFromCutOff = True Then
                    .CreatedBy = Me.mCreatedBy
                    .CreatedDate = Me.mCreatedDate
                    .ModifiedBy = Me.mModifiedBy
                    .ModifiedDate = Me.mModifiedDate
                    .TimeStamp = Me.mTimestamp
                End If
                .IdentifierAddressShipTo = vFrm.TxtIdentifierAddressShipTo.Text
                .PrimaryCustomerShipTo = vFrm.TxtPrimaryCustomerShipTo.Text
                .IdentifierAddressCustomerShipTo = vFrm.TxtIdentifierAddressCustomerShipTo.Text

                'salin data item
                ij.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                ij.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

                For Each rowInv In vDGVItem.Rows
                    ij.dgvItem.Rows.Add()
                    rowDO = ij.dgvItem.Rows(ij.dgvItem.Rows.Count - 1)

                    With rowDO
                        .Cells(ij.ColNo.Index).Value = ""
                        .Cells(ij.ColCodeNo.Index).Value = rowInv.Cells(vFrm.ColCodeNo.Index).Value
                        .Cells(ij.ColDescription.Index).Value = ""
                        .Cells(ij.ColCustCodeNo.Index).Value = ""
                        .Cells(ij.ColCustDesc.Index).Value = ""
                        .Cells(ij.ColSONO.Index).Value = vSONumber
                        .Cells(ij.ColBatchNo.Index).Value = rowInv.Cells(vFrm.ColLotNumber.Index).Value
                        .Cells(ij.ColExpiryDate.Index).Value = rowInv.Cells(vFrm.ColExpiry.Index).Value
                        .Cells(ij.ColQtyMaster.Index).Value = 0
                        .Cells(ij.ColQtyPerMaster.Index).Value = 0
                        .Cells(ij.ColQty.Index).Value = 0
                        .Cells(ij.ColMasterUnits.Index).Value = ""
                        .Cells(ij.ColTotalQty.Index).Value = rowInv.Cells(vFrm.ColTotalQty.Index).Value
                        .Cells(ij.ColUnit.Index).Value = ""
                        .Cells(ij.ColSecQty.Index).Value = -1
                        .Cells(ij.ColSecUnit.Index).Value = ""
                        .Cells(ij.ColRemarks.Index).Value = ""
                        .Cells(ij.ColPrimaryCodeNo.Index).Value = rowInv.Cells(vFrm.ColPrimaryCodeNo.Index).Value
                        .Cells(ij.ColPrimarySO.Index).Value = vPrimarySO
                        .Cells(ij.ColPrimaryMasterUnits.Index).Value = -1
                        .Cells(ij.ColPrimaryUnit.Index).Value = rowInv.Cells(vFrm.ColPrimaryUnit.Index).Value
                        .Cells(ij.ColPrimarySecUnit.Index).Value = -1

                        .Cells(ij.ColTotalQtyMulti.Index).Value = rowInv.Cells(vFrm.ColTotalQtyMulti.Index).Value
                        .Cells(ij.ColPrimaryUnitsMulti.Index).Value = rowInv.Cells(vFrm.ColPrimaryUnitsMulti.Index).Value
                        .Cells(ij.ColRatioDefault.Index).Value = rowInv.Cells(vFrm.ColRatioDefault.Index).Value
                        .Cells(ij.ColRatioReference.Index).Value = rowInv.Cells(vFrm.ColRatioReference.Index).Value

                    End With
                Next
                'simpan baru
                ObjDO.SimpanBaru1(ij.dgvItem, vUserCode, ij, False, False, vCekDataBenar, vFromCutOff)
            End With

            Return ObjDO.Primary
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            ij = Nothing
        End Try
    End Function

    Private Function SimpanStockOut(ByVal vUserCode As String, ByVal vDGVItem As DataGridView, _
    ByVal vFrm As FrmTransaksiSalesInvoiceQuick, ByVal vPrimarySO As Double, ByVal CustIdentifier As String, _
    ByVal vPrimaryDO As Double, ByVal TblCustomInOutDummy As CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable, _
    Optional ByVal vCekDataBenar As Boolean = True, _
    Optional ByVal vSetSOB As Boolean = True, Optional ByVal vKoreksiStock As Boolean = True, _
    Optional ByVal vFromCutOff As Boolean = False, Optional ByVal vOutSalesNo As String = "") As Double
        Dim ObjStock As New Inventory_Module.clsTransaksiStockInOut
        ObjStock.TipeTransaksi = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Sales
        Dim ij As New FrmTransaksiStockOutSales
        Dim rowInv, rowStock As DataGridViewRow

        Try
            'salin data header
            With ObjStock
                .ImporDariPackingList = False
                If vOutSalesNo = "" Then
                    .NoTransaksi = ObjStock.GetNoBaru(Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Sales, Me.TanggalTransaksi)
                Else
                    .NoTransaksi = vOutSalesNo '20190711
                End If
                .NoUrut = 0
                .Primary = -1
                .PrimaryCustomer = Me.PrimaryCustomer
                .PrimaryDOSales = vPrimaryDO
                .PrimaryGIT = -1
                .PrimaryPOPurchase = -1
                .PrimarySupplier = -1
                .Remarks = ""
                .SupplierDoNo = ""
                .TanggalTransaksi = Me.TanggalTransaksi
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
                For Each rowInv In vDGVItem.Rows
                    ij.dgvItem.Rows.Add()
                    rowStock = ij.dgvItem.Rows(ij.dgvItem.Rows.Count - 1)
                    With rowStock
                        .Cells(ij.ColNo.Index).Value = ""
                        .Cells(ij.ColPONoDONo.Index).Value = ""
                        .Cells(ij.ColSuppDONo.Index).Value = ""
                        .Cells(ij.ColCodeNo.Index).Value = ""
                        .Cells(ij.ColDescription.Index).Value = ""
                        .Cells(ij.ColQty.Index).Value = rowInv.Cells(vFrm.ColTotalQty.Index).Value
                        .Cells(ij.ColUnit.Index).Value = ""
                        .Cells(ij.ColSecondaryQty.Index).Value = -1
                        .Cells(ij.ColSecondaryUnit.Index).Value = ""
                        .Cells(ij.ColBatch.Index).Value = rowInv.Cells(vFrm.ColLotNumber.Index).Value
                        .Cells(ij.ColExpiredDate.Index).Value = rowInv.Cells(vFrm.ColExpiry.Index).Value
                        .Cells(ij.ColLocation.Index).Value = ""
                        .Cells(ij.ColWareHouse.Index).Value = ""
                        .Cells(ij.ColRemark.Index).Value = ""
                        .Cells(ij.ColPrimaryPOnoDoNo.Index).Value = vPrimarySO
                        .Cells(ij.ColPrimaryCodeNo.Index).Value = rowInv.Cells(vFrm.ColPrimaryCodeNo.Index).Value
                        .Cells(ij.ColPrimaryLocation.Index).Value = rowInv.Cells(vFrm.ColPrimaryLocation.Index).Value
                        .Cells(ij.ColPrimaryUnits.Index).Value = rowInv.Cells(vFrm.ColPrimaryUnit.Index).Value
                        .Cells(ij.ColPrimarySecondaryUnits.Index).Value = -1
                        .Cells(ij.ColPrimaryPdO.Index).Value = -1

                        .Cells(ij.ColQtyMulti.Index).Value = rowInv.Cells(vFrm.ColTotalQtyMulti.Index).Value
                        .Cells(ij.ColPrimaryUnitsMulti.Index).Value = rowInv.Cells(vFrm.ColPrimaryUnitsMulti.Index).Value
                        .Cells(ij.ColRatioDefault.Index).Value = rowInv.Cells(vFrm.ColRatioDefault.Index).Value
                        .Cells(ij.ColRatioReference.Index).Value = rowInv.Cells(vFrm.ColRatioReference.Index).Value

                    End With
                Next
            End With

            'simpan
            ObjStock.SimpanBaru1(ij.dgvItem, vUserCode, TblCustomInOutDummy)

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

    Public Function ApakahInvoiceDP(ByVal vPrimaryInvoice As Double) As Boolean
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice1TableAdapter

        Try
            If DTA.ScalarQuery_GetTipeInvoice(vPrimaryInvoice).HasValue = True Then
                If DTA.ScalarQuery_GetTipeInvoice(vPrimaryInvoice).Value = 0 Then
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

    Public Function ApakahQuickInvoice(ByVal vPrimaryInvoice As Double) As Boolean
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice1TableAdapter

        Try
            If DTA.ScalarQuery_GetApakahQuickInvoice(vPrimaryInvoice).HasValue = True Then
                If DTA.ScalarQuery_GetApakahQuickInvoice(vPrimaryInvoice).Value = 0 Then
                    Return False
                Else
                    Return True
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

    'menyimpan data sales amount yang digunakan untuk laporan sales analysis dan sales statistic
    Public Sub SimpanDataSalesAmount(ByVal vCalculatesBy As ClsTransaksiSalesInvoice.EnumCalculatesBy, _
    ByVal vPrimarySalesInvoice As Double, ByVal TotalSales As Decimal)
        Dim Rate As Decimal = 1
        Dim vPrimaryTransaksiRujukan As Double = -1
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_item1TableAdapter
        Dim TBL As New DseTransaksiSalesInvoice.tbltrans_salesinvoice_item1DataTable
        Dim row1 As DseTransaksiSalesInvoice.tbltrans_salesinvoice_item1Row
        Dim SalesAmount As Decimal = 0
        Dim DTAReturn As New DseTransaksiSalesInvoiceTableAdapters.DataTable11TableAdapter
        Dim TBLReturn As New DseTransaksiSalesInvoice.DataTable11DataTable
        Dim i As Integer = 0
        Dim DefPrecBase As Integer = 0
        Dim ObjCur As New Core_Module.clsMasterCurrency

        Try
            DefPrecBase = ObjCur.GetDecimal_DefaultCurrency
            'isikan data item
            DTA.Fill(TBL, vPrimarySalesInvoice)
            If TBL.Rows.Count = 0 Then
                'kalo ga ada itemnya maka ga usah update data cost
                Exit Sub
            End If
            DTAReturn.Fill(TBLReturn, vPrimarySalesInvoice)

            'kurangkan dengan item return
            Me.KurangkanItemReturn1(TBLReturn, TBL)
            If TBL.Rows.Count = 0 Then
                'kalo ga ada itemnya maka ga usah update data cost
                Exit Sub
            End If

            'hitung exchange rate
            If vCalculatesBy = EnumCalculatesBy.Invoice Then
                row1 = TBL.Rows(0)
                If vCalculatesBy = ClsTransaksiSalesInvoice.EnumCalculatesBy.Delivery_Order Then
                    vPrimaryTransaksiRujukan = row1.Primary_DO
                ElseIf vCalculatesBy = ClsTransaksiSalesInvoice.EnumCalculatesBy.Invoice Then
                    vPrimaryTransaksiRujukan = row1.Primary_Invoice
                ElseIf vCalculatesBy = ClsTransaksiSalesInvoice.EnumCalculatesBy.Sales_Order Then
                    vPrimaryTransaksiRujukan = row1.Primary_SO
                End If
                Rate = Me.GetExhangeRate(vPrimarySalesInvoice, vPrimaryTransaksiRujukan)
            End If

            For i = 0 To TBL.Rows.Count - 1
                row1 = TBL.Rows(i)
                'hitung exchange rate
                If vCalculatesBy <> EnumCalculatesBy.Invoice Then
                    If vCalculatesBy = ClsTransaksiSalesInvoice.EnumCalculatesBy.Delivery_Order Then
                        vPrimaryTransaksiRujukan = row1.Primary_DO
                    ElseIf vCalculatesBy = ClsTransaksiSalesInvoice.EnumCalculatesBy.Invoice Then
                        vPrimaryTransaksiRujukan = row1.Primary_Invoice
                    ElseIf vCalculatesBy = ClsTransaksiSalesInvoice.EnumCalculatesBy.Sales_Order Then
                        vPrimaryTransaksiRujukan = row1.Primary_SO
                    End If
                    Rate = Me.GetExhangeRate(vPrimarySalesInvoice, vPrimaryTransaksiRujukan)
                End If

                If i <> TBL.Rows.Count - 1 Then
                    'untuk setiap item, hitung amount
                    SalesAmount = Math.Round(Math.Round(row1.Amount * Rate, _
                    row1.Default_Precision, MidpointRounding.AwayFromZero) * row1.Tax_Exchange_Rate, _
                    DefPrecBase, MidpointRounding.AwayFromZero)
                    'update data sales amount
                    DTA.UpdateQuery_SetSalesAmount(SalesAmount, row1.Primary)
                    TotalSales -= SalesAmount
                Else
                    SalesAmount = TotalSales
                    'update data sales amount
                    DTA.UpdateQuery_SetSalesAmount(SalesAmount, row1.Primary)
                    If SalesAmount <= 0 Then
                        MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjCur = Nothing
        End Try
    End Sub

    Public Function GetExhangeRate(ByVal vPrimarySalesInvoice As Double, _
ByVal vPrimaryTransaksiRujukan As Double) As Decimal
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_item_exchangerateTableAdapter

        Try
            If DTA.ScalarQuery_GetExhangeRate(vPrimarySalesInvoice, vPrimaryTransaksiRujukan).HasValue = True Then
                Return DTA.ScalarQuery_GetExhangeRate(vPrimarySalesInvoice, vPrimaryTransaksiRujukan).Value
            Else
                Return 1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Sub KurangkanItemReturn1(ByVal vTBLReturn As DseTransaksiSalesInvoice.DataTable11DataTable, _
ByVal vTBLInvoice As DseTransaksiSalesInvoice.tbltrans_salesinvoice_item1DataTable)
        Dim Filter As String = ""
        Dim row1 As DseTransaksiSalesInvoice.tbltrans_salesinvoice_item1Row
        Dim rowReturn As DseTransaksiSalesInvoice.DataTable11Row
        Dim i As Integer = 0

        Try
            For Each rowReturn In vTBLReturn.Rows
                Filter = "[Primary Invoice] =" + rowReturn.Primary_Sales_Invoice.ToString + " AND [Primary Product] =" + _
                rowReturn.Primary_Product.ToString + " AND [Primary SO] =" + rowReturn.Primary_SO.ToString + _
                " AND [Primary DO] =" + rowReturn.Primary_DO.ToString
                vTBLInvoice.DefaultView.RowFilter = Filter
                i = 0
                While vTBLInvoice.DefaultView.Count > i And rowReturn.Total_Qty_Multi > 0
                    row1 = vTBLInvoice.DefaultView(i).Row
                    If row1.Total_Qty_Multi >= rowReturn.Total_Qty_Multi Then
                        row1.Amount = Math.Round(((row1.Total_Qty_Multi - rowReturn.Total_Qty_Multi) * row1.Amount) / row1.Total_Qty_Multi, row1.Default_Precision_Item, MidpointRounding.AwayFromZero)
                        row1.Total_Qty_Multi -= rowReturn.Total_Qty_Multi
                        rowReturn.Total_Qty_Multi = 0
                    Else
                        row1.Total_Qty_Multi = 0
                        rowReturn.Total_Qty_Multi -= row1.Total_Qty_Multi
                    End If

                    If row1.Total_Qty_Multi = 0 Then
                        vTBLInvoice.Rows.Remove(row1)
                    End If
                End While
            Next

            vTBLInvoice.DefaultView.RowFilter = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahTransferOtomatis(ByVal vTanggalInvoice As Date) As Boolean
        Dim DTA As New DseTransaksiSalesInvoice_TransfJurnalTableAdapters.tbltrans_accounting_closingTableAdapter

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

    ''' <summary>
    ''' memindahkan data dari tabel sales invoice aktif ke tabel sales invoice obselete
    ''' </summary>
    ''' <param name="vPrimaryInvoiceLama"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SimpanDataObseleteInvoice(ByVal vPrimaryInvoiceLama As Double) As Boolean
        Dim DTAInvoiceLama As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice1TableAdapter
        Dim DTADPLama As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_itemdpTableAdapter
        Dim DTAItemLama As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_itemTableAdapter
        Dim DTASRLama As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_item_salesreturnTableAdapter
        Dim DTAExchangeRateLama As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_item_exchangerateTableAdapter

        Dim TBLExhangeRate As New DseTransaksiSalesInvoice.tbltrans_salesinvoice_item_exchangerateDataTable
        Dim TBLSR As New DseTransaksiSalesInvoice.tbltrans_salesinvoice_item_salesreturnDataTable
        Dim TBLItem As New DseTransaksiSalesInvoice.tbltrans_salesinvoice_itemDataTable
        Dim TBLDP As New DseTransaksiSalesInvoice.tbltrans_salesinvoice_itemdpDataTable
        Dim TBLInvoice As New DseTransaksiSalesInvoice.tbltrans_salesinvoice1DataTable

        Dim rowInv As DseTransaksiSalesInvoice.tbltrans_salesinvoice1Row
        Dim rowDP As DseTransaksiSalesInvoice.tbltrans_salesinvoice_itemdpRow
        Dim rowSR As DseTransaksiSalesInvoice.tbltrans_salesinvoice_item_salesreturnRow
        Dim rowItem As DseTransaksiSalesInvoice.tbltrans_salesinvoice_itemRow
        Dim rowExchange As DseTransaksiSalesInvoice.tbltrans_salesinvoice_item_exchangerateRow

        Dim DTAInvoiceObselete As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_obseleteTableAdapter
        Dim DTAItemObselete As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoiceobselete_itemTableAdapter
        Dim DTADPObselete As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoiceobselete_itemdpTableAdapter
        Dim DTASRObselete As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoiceobselete_item_salesreturnTableAdapter
        Dim DTAExchangeRateobselete As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoiceobselete_item_exchangerateTableAdapter

        Try
            'isikan tabel lama
            DTAInvoiceLama.FillBy_Primary(TBLInvoice, vPrimaryInvoiceLama)
            DTADPLama.FillBy_PrimaryInvoice(TBLDP, vPrimaryInvoiceLama)
            DTAItemLama.FillBy_PrimaryInvoice(TBLItem, vPrimaryInvoiceLama)
            DTASRLama.FillBy_PrimarySalesInvoice(TBLSR, vPrimaryInvoiceLama)
            DTAExchangeRateLama.FillBy_PrimarySalesInvoice(TBLExhangeRate, vPrimaryInvoiceLama)

            'simpan ke dalam tabel baru
            If TBLInvoice.Rows.Count > 0 Then
                'simpan data invoice
                rowInv = TBLInvoice.Rows(0)
                Dim JenisFaktur, TaxExclude As Integer
                If rowInv.Jenis_Faktur_Pajak = False Then
                    JenisFaktur = 0
                Else
                    JenisFaktur = 1
                End If
                If rowInv.Tax_Exclude = False Then
                    TaxExclude = 0
                Else
                    TaxExclude = 1
                End If
                With rowInv
                    DTAInvoiceObselete.Insert(.No_Transaksi, .Urutan, .Urutan_Faktur_Pajak, .Tipe_Invoice, _
                    .Tanggal_Transaksi, .Primary_Customer, .Primary_Currency, .Calculates_By, .Sub_Total_Before_DP, _
                    .Down_Payment, .Sub_Total, .Duty_Stamp, TaxExclude, .Tax, .Tax_Exchange_Rate, .Tax_Amount, _
                    .Total, .Exchange_Rate_To_USD, .Exchange_Rate_To_IDR, .No_Seri_Faktur_Pajak, .No_Faktur_Pajak, _
                    .No_NPWP, .Address_Identifier, .Remarks, .Created_By, .Created_Date, .Modified_By, .Modified_Date, _
                    .Timestamp, JenisFaktur, .Payment_Term, 0, .Revisi, .Primary_SO_4_DP, .Nama_Signed, .Jabatan_Signed, _
                    .No_Faktur_Pajak_Sederhana, .Alamat_Pajak_1, .Alamat_Pajak_2, .exchange_rate, .ongkos_kirim, .VAT_Dipungut, _
                    .Shipment_With_VAT, .Nama_Jalan, .Blok, .Nomor, .RT, .RW, .Kecamatan, .Kelurahan, .Kabupaten, .Propinsi, _
                    .Kode_Pos, .Telepon, .kode_jenis_transaksi_pajak)
                End With
                Dim PrimaryInvObselete As Double = DTAInvoiceObselete.ScalarQuery_GetPrimary(rowInv.No_Transaksi)

                'simpan data item
                For Each rowItem In TBLItem.Rows
                    With rowItem
                        DTAItemObselete.Insert(PrimaryInvObselete, .Primary_Product, .Primary_DO, .Primary_SO, _
                        .Qty_Master, .Qty_Per_Master, .Primary_Master_Unit, .Qty, .Total_Qty, .Primary_Units, _
                        .Secondary_Qty, .Primary_Secondary_Units, .Primary_Currency, .Unit_Price, .Discount, _
                        .Nett_Price, .Amount, .Remarks, .Batch_Number, .Expiry, .Alias_Code, .Alias_Desc, .Disc_Amount, .DiscFromPerc, _
                        .Total_Qty_Multi, .Primary_Units_Multi, .Ratio_Reference, .Ratio_Default)
                    End With
                Next

                'simpan data item DP
                For Each rowDP In TBLDP.Rows
                    With rowDP
                        DTADPObselete.Insert(PrimaryInvObselete, .Primary_Invoice_DP, .Remarks)
                    End With
                Next

                'simpan data item SR
                For Each rowSR In TBLSR.Rows
                    With rowSR
                        DTASRObselete.Insert(PrimaryInvObselete, .Primary_SR, .Remarks)
                    End With
                Next

                'simpan data exhange rate
                Dim CalculatesBy As EnumCalculatesBy = rowInv.Calculates_By
                For Each rowExchange In TBLExhangeRate
                    If CalculatesBy = EnumCalculatesBy.Invoice Then
                        With rowExchange
                            DTAExchangeRateobselete.Insert(PrimaryInvObselete, PrimaryInvObselete, _
                            .Primary_Currency, .Exhange_Rate, .Remarks)
                        End With
                    Else
                        With rowExchange
                            DTAExchangeRateobselete.Insert(PrimaryInvObselete, .Primary_Transaksi_Rujukan, _
                            .Primary_Currency, .Exhange_Rate, .Remarks)
                        End With
                    End If
                Next
            End If

            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAInvoiceLama.Dispose()
            DTADPLama.Dispose()
            DTAItemLama.Dispose()
            DTASRLama.Dispose()
            DTAExchangeRateLama.Dispose()

            TBLExhangeRate.Dispose()
            TBLSR.Dispose()
            TBLItem.Dispose()
            TBLDP.Dispose()
            TBLInvoice.Dispose()

            DTAInvoiceObselete.Dispose()
            DTAItemObselete.Dispose()
            DTADPObselete.Dispose()
            DTASRObselete.Dispose()
            DTAExchangeRateobselete.Dispose()
        End Try
    End Function

    'Public Function SimpanDataInvoiceRevisiLama(ByVal vPrimaryInvoiceLama As Double) As Boolean
    '    Dim DTAInvoiceLama As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice1TableAdapter
    '    Dim DTADPLama As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_itemdpTableAdapter
    '    Dim DTAItemLama As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_itemTableAdapter
    '    Dim DTASRLama As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_item_salesreturnTableAdapter
    '    Dim DTAExchangeRateLama As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_item_exchangerateTableAdapter

    '    Dim TBLExhangeRate As New DseTransaksiSalesInvoice.tbltrans_salesinvoiceobselete_item_exchangerateDataTable
    '    Dim TBLSR As New DseTransaksiSalesInvoice.tbltrans_salesinvoiceobselete_item_salesreturnDataTable
    '    Dim TBLItem As New DseTransaksiSalesInvoice.tbltrans_salesinvoiceobselete_itemDataTable
    '    Dim TBLDP As New DseTransaksiSalesInvoice.tbltrans_salesinvoiceobselete_itemdpDataTable
    '    Dim TBLInvoice As New DseTransaksiSalesInvoice.tbltrans_salesinvoice_obseleteDataTable

    '    Dim rowInv As DseTransaksiSalesInvoice.tbltrans_salesinvoice_obseleteRow
    '    Dim rowDP As DseTransaksiSalesInvoice.tbltrans_salesinvoiceobselete_itemdpRow
    '    Dim rowSR As DseTransaksiSalesInvoice.tbltrans_salesinvoiceobselete_item_salesreturnRow
    '    Dim rowItem As DseTransaksiSalesInvoice.tbltrans_salesinvoiceobselete_itemRow
    '    Dim rowExchange As DseTransaksiSalesInvoice.tbltrans_salesinvoiceobselete_item_exchangerateRow

    '    Dim DTAInvoiceObselete As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_obseleteTableAdapter
    '    Dim DTAItemObselete As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoiceobselete_itemTableAdapter
    '    Dim DTADPObselete As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoiceobselete_itemdpTableAdapter
    '    Dim DTASRObselete As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoiceobselete_item_salesreturnTableAdapter
    '    Dim DTAExchangeRateobselete As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoiceobselete_item_exchangerateTableAdapter

    '    Dim DTA_Inv_DO As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_invoice_doTableAdapter
    '    Dim TBL_Inv_DO As New DseTransaksiSalesInvoice.tbltrans_invoice_doDataTable
    '    Dim rowInvDO As DseTransaksiSalesInvoice.tbltrans_invoice_doRow
    '    Dim Obj As New ClsTransaksiSalesInvoice_SettingTaxInvoice
    '    Dim PrimaryFakturPajak As Double = -1
    '    Dim NoFakturPajakMax As String = ""
    '    Dim DefPrecBase As Integer = 0
    '    Dim ObjCur As New Core_Module.clsMasterCurrency

    '    Dim DPOriginal, DPBaseCurrency As Decimal
    '    Dim DTA1 As New DseTransaksiSalesInvoiceTableAdapters.DataTable13TableAdapter
    '    Dim TBL1 As New DseTransaksiSalesInvoice.DataTable13DataTable
    '    Dim row1 As DseTransaksiSalesInvoice.DataTable13Row
    '    Dim DefPrec As Integer = 0

    '    Dim TBLitemCost As New DseTransaksiSalesInvoice.tbltrans_salesinvoice_item_costDataTable
    '    Dim RowitemCost As DseTransaksiSalesInvoice.tbltrans_salesinvoice_item_costRow


    '    Try
    '        'isikan tabel lama
    '        DTAInvoiceObselete.FillBy_PrimaryInvoice(TBLInvoice, vPrimaryInvoiceLama)
    '        DTADPObselete.FillBy_PrimaryInvoice(TBLDP, vPrimaryInvoiceLama)
    '        DTAItemObselete.FillBy_PrimaryInvoice(TBLItem, vPrimaryInvoiceLama)
    '        DTASRObselete.FillBy_PrimaryInvoice(TBLSR, vPrimaryInvoiceLama)
    '        DTAExchangeRateobselete.FillBy_PrimaryInvoice(TBLExhangeRate, vPrimaryInvoiceLama)

    '        'simpan ke dalam tabel baru
    '        If TBLInvoice.Rows.Count > 0 Then
    '            'simpan data invoice
    '            rowInv = TBLInvoice.Rows(0)
    '            Dim JenisFaktur, TaxExclude As Integer
    '            If rowInv.Jenis_Faktur_Pajak = False Then
    '                JenisFaktur = 0
    '            Else
    '                JenisFaktur = 1
    '            End If
    '            If rowInv.Tax_Exclude = False Then
    '                TaxExclude = 0
    '            Else
    '                TaxExclude = 1
    '            End If

    '            With rowInv
    '                DTAInvoiceLama.InsertQuery(.No_Transaksi, .Urutan, .Urutan_Faktur_Pajak, .Tipe_Invoice, _
    '                .Tanggal_Transaksi, .Primary_Customer, .Primary_Currency, .Calculates_By, .Sub_Total_Before_DP, _
    '                .Down_Payment, .Sub_Total, .Duty_Stamp, TaxExclude, .Tax, .Tax_Exchange_Rate, .Tax_Amount, _
    '                .Total, .Exchange_Rate_To_USD, .Exchange_Rate_To_IDR, .No_Seri_Faktur_Pajak, .No_Faktur_Pajak, _
    '                .No_NPWP, .Address_Identifier, .Remarks, .Created_By, .Created_Date, .Modified_By, .Modified_Date, _
    '                .Timestamp, JenisFaktur, .Payment_Term, .Primary_SO_4_DP, _
    '                .No_Faktur_Pajak_Sederhana, .Nama_Signed, .Jabatan_Signed, .JmlCetak, .Revisi, .Alamat_Pajak_1, _
    '                .Alamat_Pajak_2, .exchange_rate, .ongkos_kirim, .VAT_Dipungut, .Shipment_WIth_VAT, .primary_sales_person, _
    '                .disc_string, .disc_amount, .Nama_Jalan, .Blok, .Nomor, .RT, .RW, .Kecamatan, .Kelurahan, .Kabupaten, _
    '                .Propinsi, .Kode_Pos, .Telepon, .kode_jenis_transaksi_pajak)
    '            End With
    '            Dim PrimaryInvObselete As Double = DTAInvoiceLama.ScalarQuery_GetPrimary(rowInv.No_Transaksi)

    '            'simpan data item
    '            For Each rowItem In TBLItem.Rows
    '                With rowItem
    '                    DTAItemLama.Insert(PrimaryInvObselete, .Primary_Product, .Primary_DO, .Primary_SO, _
    '                      .Qty_Master, .Primary_Master_Unit, .Qty, .Total_Qty, .Primary_Units, _
    '                    .Secondary_Qty, .Primary_Secondary_Units, .Unit_Price, .Discount, _
    '                    .Nett_Price, .Amount, .Remarks, .Qty_Per_Master, .Primary_Currency, .Batch_Number, .Expiry, _
    '                    .Alias_Code, .Alias_Desc, .Disc_Amount, .Disc_From_Perct, .Total_Qty_Multi, .Primary_Units_Multi, .Ratio_Reference, .Ratio_Default)
    '                End With
    '            Next

    '            'simpan data item DP
    '            For Each rowDP In TBLDP.Rows
    '                With rowDP
    '                    DTADPLama.Insert(PrimaryInvObselete, .Primary_Invoice_DP, .Remarks , 0)
    '                End With
    '            Next

    '            'simpan data item SR
    '            For Each rowSR In TBLSR.Rows
    '                With rowSR
    '                    DTASRLama.Insert(PrimaryInvObselete, .Primary_SR, .Remarks)
    '                End With
    '            Next

    '            'simpan data exhange rate
    '            Dim CalculatesBy As EnumCalculatesBy = rowInv.Calculates_By
    '            For Each rowExchange In TBLExhangeRate
    '                If CalculatesBy = EnumCalculatesBy.Invoice Then
    '                    With rowExchange
    '                        DTAExchangeRateLama.Insert(PrimaryInvObselete, PrimaryInvObselete, _
    '                        .Primary_Currency, .Exhange_Rate, .Remarks)
    '                    End With
    '                Else
    '                    With rowExchange
    '                        DTAExchangeRateLama.Insert(PrimaryInvObselete, .Primary_Transaksi_Rujukan, _
    '                        .Primary_Currency, .Exhange_Rate, .Remarks)
    '                    End With
    '                End If
    '            Next

    '            'tandai do
    '            DTA_Inv_DO.FillBy_PrimaryInvoice(TBL_Inv_DO, PrimaryInvObselete)
    '            For Each rowInvDO In TBL_Inv_DO.Rows
    '                DTA_Inv_DO.UpdateQuery_SetSudahBuatInvoice(1, rowInvDO.Primary_Transaksi, rowInvDO.Primary_DO)
    '            Next

    '            'simpan data sales amount 
    '            With rowInv
    '                Me.SimpanDataSalesAmount(rowInv.Calculates_By, _
    '                PrimaryInvObselete, Math.Round(.Sub_Total_Before_DP * .Tax_Exchange_Rate, ObjCur.GetDecimal_DefaultCurrency, MidpointRounding.AwayFromZero))
    '            End With

    '            'hitung DP
    '            DTA1.Fill(TBL1, Me.mPrimary)
    '            DPOriginal = 0
    '            DPBaseCurrency = 0
    '            Dim PrimaryAkunDP As Double = -1
    '            For Each row1 In TBL1.Rows
    '                DPOriginal += row1.Sub_Total_Before_DP
    '                DPBaseCurrency += Math.Round(row1.Sub_Total_Before_DP * row1.Tax_Exchange_Rate, DefPrecBase, MidpointRounding.AwayFromZero)
    '                PrimaryAkunDP = row1.Primary_Invoice_DP
    '            Next

    '            'simpan data jurnal
    '            With rowInv
    '                DefPrecBase = ObjCur.GetDecimal_DefaultCurrency
    '                Me.TransferJurnal_Invoice2(Me.mPrimary, Me.mPrimaryCust, _
    '                  Math.Round(Me.mSubTotalB4DP * Me.mTaxExchangeRate, DefPrecBase, MidpointRounding.AwayFromZero), _
    '                  Math.Round(Me.mTaxAmount * Me.mTaxExchangeRate, DefPrecBase, MidpointRounding.AwayFromZero), _
    '                  Math.Round(Me.mDutyStamp * Me.mTaxExchangeRate, DefPrecBase, MidpointRounding.AwayFromZero), _
    '                  Me.mNoUrut, Me.mTanggalTransaksi, Me.mNoTransaksi, Me.mModifiedBy, CDbl(Me.mPaymentTerm), Me.mTaxExchangeRate, _
    '                  Me.mSubTotalB4DP, Me.mPrimaryCurrency, Math.Round(Me.mOngkosKirim * Me.mTaxExchangeRate, DefPrecBase, MidpointRounding.AwayFromZero), _
    '                  DPOriginal, DPBaseCurrency, Me.mOngkosKirim, Me.mDiscAmount, _
    '                  Math.Round(Me.mDiscAmount * Me.mTaxExchangeRate, DefPrecBase, MidpointRounding.AwayFromZero), _
    '                  Me.mPrimaryCurrTax, Me.mPrimaryAccount, Me.mDutyStamp, Me.mTaxAmount, PrimaryAkunDP, _
    '                  Me.mOtherCharge * Me.mTaxExchangeRate, Me.mOtherCharge, TBLitemCost, RowitemCost)
    '            End With

    '            'simpan no urut max untuk numbering
    '            If Obj.BatasStartDate <= Me.mTanggalTransaksi Then
    '                If Me.mJenisFakturPajak = EnumJenisFakturPajak.FakturPajakStandar Then
    '                    DTAInvoiceLama.UpdateQuery_PrimaryFakturPajak(Obj.GetPrimary(rowInv.Tanggal_Transaksi, rowInv.No_Faktur_Pajak), PrimaryInvObselete)
    '                    PrimaryFakturPajak = DTAInvoiceLama.ScalarQuery_GetPrimaryFakturPajak(PrimaryInvObselete)
    '                    NoFakturPajakMax = Me.GetNoFakturPajakMax(PrimaryFakturPajak, PrimaryInvObselete)

    '                    If Obj.GetUrutan(NoFakturPajakMax, rowInv.Tanggal_Transaksi) < Obj.GetUrutan(rowInv.No_Faktur_Pajak, rowInv.Tanggal_Transaksi) Then
    '                        Obj.SetNoUrutMaksimal(rowInv.Tanggal_Transaksi, rowInv.No_Faktur_Pajak)
    '                    Else
    '                        Obj.SetNoUrutMaksimal(PrimaryFakturPajak, NoFakturPajakMax)
    '                    End If
    '                End If
    '            End If
    '        End If

    '        'nilai balikan
    '        Return True
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTAInvoiceLama.Dispose()
    '        DTADPLama.Dispose()
    '        DTAItemLama.Dispose()
    '        DTASRLama.Dispose()
    '        DTAExchangeRateLama.Dispose()

    '        TBLExhangeRate.Dispose()
    '        TBLSR.Dispose()
    '        TBLItem.Dispose()
    '        TBLDP.Dispose()
    '        TBLInvoice.Dispose()

    '        DTAInvoiceObselete.Dispose()
    '        DTAItemObselete.Dispose()
    '        DTADPObselete.Dispose()
    '        DTASRObselete.Dispose()
    '        DTAExchangeRateobselete.Dispose()

    '        DTA_Inv_DO.Dispose()
    '        TBL_Inv_DO.Dispose()
    '        ObjCur = Nothing
    '        Obj = Nothing

    '        DTA1.Dispose()
    '        TBL1.Dispose()
    '    End Try
    'End Function

    Public Function SimpanRevisi(ByVal vdgvItem As DataGridView, ByVal vdgvItemDP As DataGridView, _
  ByVal vDGVItemSR As DataGridView, ByVal vDGVItemExchangeRate As DataGridView, ByVal vUserCode As String, _
  ByVal vPrimaryRevisiLama As Double, ByVal vFrm As FrmTransaksiSalesInvoiceQuick, ByVal vNoDO As String, _
  ByVal vTblCustom As CustomSMESS.DseTransaksiSalesInvoiceQuick_AddByCustomProduct.tbltrans_custom_salesinvoice_item1DataTable, _
  ByVal TblCustomInOutDummy As CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable) As Boolean
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice1TableAdapter
        Dim DTA_Inv_DO As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_invoice_doTableAdapter
        Dim TBL_Inv_DO As New DseTransaksiSalesInvoice.tbltrans_invoice_doDataTable
        Dim rowInvDO As DseTransaksiSalesInvoice.tbltrans_invoice_doRow

        Try
            'simpan data lama invoice pada tabel obselete invoice
            Me.SimpanDataObseleteInvoice(vPrimaryRevisiLama)
            'hapus data jurnal yang lama
            Me.Hapus_Jurnal(vPrimaryRevisiLama)
            'hapus data lama invoice-semua data item juga akan hilang
            DTA_Inv_DO.FillBy_PrimaryInvoice(TBL_Inv_DO, vPrimaryRevisiLama)
            'tandai do belum buat inv
            For Each rowInvDO In TBL_Inv_DO.Rows
                DTA_Inv_DO.UpdateQuery_SetSudahBuatInvoice(0, -1, rowInvDO.Primary_DO)
            Next
            DTA.Delete(vPrimaryRevisiLama)
            'simpan data revisi baru
            Me.SimpanBaru(vdgvItem, vdgvItemDP, vDGVItemSR, vDGVItemExchangeRate, vUserCode, vFrm, vNoDO, vTblCustom, TblCustomInOutDummy)
            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA_Inv_DO.Dispose()
            TBL_Inv_DO.Dispose()
        End Try
    End Function

    Public Function SimpanRevisi(ByVal vdgvItem As DataGridView, ByVal vdgvItemDP As DataGridView, _
    ByVal vDGVItemSR As DataGridView, ByVal vDGVItemExchangeRate As DataGridView, ByVal vUserCode As String, _
    ByVal vPrimaryRevisiLama As Double, ByVal vFrm As FrmTransaksiSalesInvoice) As Boolean
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice1TableAdapter
        Dim DTA_Inv_DO As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_invoice_doTableAdapter
        Dim TBL_Inv_DO As New DseTransaksiSalesInvoice.tbltrans_invoice_doDataTable
        Dim rowInvDO As DseTransaksiSalesInvoice.tbltrans_invoice_doRow

        Try
            'simpan data lama invoice pada tabel obselete invoice
            Me.SimpanDataObseleteInvoice(vPrimaryRevisiLama)
            'hapus data jurnal yang lama
            Me.Hapus_Jurnal(vPrimaryRevisiLama)
            'hapus data lama invoice-semua data item juga akan hilang
            DTA_Inv_DO.FillBy_PrimaryInvoice(TBL_Inv_DO, vPrimaryRevisiLama)
            'tandai do belum buat inv
            For Each rowInvDO In TBL_Inv_DO.Rows
                DTA_Inv_DO.UpdateQuery_SetSudahBuatInvoice(0, -1, rowInvDO.Primary_DO)
            Next
            DTA.Delete(vPrimaryRevisiLama)
            'simpan data revisi baru
            Me.SimpanBaru(vdgvItem, vdgvItemDP, vDGVItemSR, vDGVItemExchangeRate, vUserCode, vFrm)
            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA_Inv_DO.Dispose()
            TBL_Inv_DO.Dispose()
        End Try
    End Function

    'Private Sub CekNomor2(ByVal DTA As DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice1TableAdapter)
    '    Dim TanggalDari, TanggalSampai As Date
    '    Dim Urutan As Double = Me.mNoUrut
    '    Dim LoopGo As Boolean = True
    '    Dim JmlTrans As Integer
    '    Dim LoopCount As Integer
    '    Dim ObjCD As New Core_Module.ClsCekData
    '    Dim PrimaryMax As Double = -1

    '    Try
    '        TanggalDari = ObjCD.GetFisrtDateOfTheYear(Me.mTanggalTransaksi)
    '        If TanggalDari.Date < #5/1/2010# Then
    '            'tanggan ini dipilih 1 mei 2010 karena yakin pada tanggal tersebut di CMSI
    '            'tidak ada penomoran bulanan lagi
    '            TanggalDari = #5/1/2010#
    '        End If
    '        TanggalSampai = ObjCD.GetLastDateOfTheYear(Me.mTanggalTransaksi)

    '        LoopCount = 0
    '        While LoopGo = True
    '            If DTA.ScalarQuery_GetJmlByUrutan(Me.mNoUrut, TanggalSampai, TanggalDari).HasValue = True Then
    '                JmlTrans = DTA.ScalarQuery_GetJmlByUrutan(Me.mNoUrut, TanggalSampai, TanggalDari).Value
    '            Else
    '                JmlTrans = 1
    '            End If
    '            If JmlTrans <= 1 Then
    '                LoopGo = False
    '            Else
    '                If DTA.ScalarQuery_GetPrimaryMaxByUrutan(Me.mNoUrut, TanggalSampai, TanggalDari).HasValue = True Then
    '                    PrimaryMax = DTA.ScalarQuery_GetPrimaryMaxByUrutan(Me.mNoUrut, TanggalSampai, TanggalDari)
    '                Else
    '                    PrimaryMax = -1
    '                End If
    '                If PrimaryMax = Me.mPrimary Then
    '                    Me.mNoTransaksi = Me.GetNoBaru1(Me.mTanggalTransaksi)
    '                    Me.mNoUrut = CInt(Me.mNoTransaksi.Substring(0, 6))
    '                End If
    '            End If
    '            LoopCount += 1
    '        End While
    '        If LoopCount > 1 Then
    '            DTA.UpdateQuery_NoTrans_NoUrut(Me.mNoTransaksi, Me.mNoUrut, Me.mPrimary)
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        ObjCD = Nothing
    '    End Try
    'End Sub


    Private Function ApakahStockTidakMinus(ByVal vDGVItem As DataGridView, ByVal vFrm As FrmTransaksiSalesInvoiceQuick, _
    ByVal vPrimaryStockOutSales As Double) As Boolean
        Dim DTA As New Inventory_Module.dseTransaksiStockInOutTableAdapters.transaksi_stockinout_item_viewTableAdapter
        Dim TBL As New Inventory_Module.dseTransaksiStockInOut.transaksi_stockinout_item_viewDataTable
        Dim rowInOutLama As Inventory_Module.dseTransaksiStockInOut.transaksi_stockinout_item_viewRow
        Dim TipeProduct As Core_Module.clsMasterProduct.EnumTipeProduct = Core_Module.clsMasterProduct.EnumTipeProduct.Goods
        Dim DTAMasterProductQty As New Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter
        DTAMasterProductQty.ClearBeforeFill = False
        Dim TBLMasterProductQty As New Core_Module.dseMasterProduct.tblmaster_productqtyDataTable

        Dim rowDGV As DataGridViewRow
        Dim BatchNo As String
        Dim PrimaryProduct, PrimaryLocation, Sec_Qty, Qty As Double
        Dim ExpiryDate As Date

        Dim ListPrimaryProduct As New Queue(Of Double)
        Dim rowProduct As Core_Module.dseMasterProduct.tblmaster_productqtyRow
        Dim Dummy As String = ""
        Dim ObjProd As New Core_Module.clsMasterProduct

        Try
            'Simulasikan terlebih dahulu perubahan data yang akan terjadi, pastikan tidak ada yang negatif
            'load data item lama

            'load database master product qty berdasarkan data product lama dan data product baru
            TBL.Rows.Clear()
            DTA.FillBy_Primary_Simulasi(TBL, vPrimaryStockOutSales, CInt(TipeProduct))
            For Each rowInOutLama In TBL.Rows
                With rowInOutLama
                    DTAMasterProductQty.FillBy_Parameter2(TBLMasterProductQty, _
                    .Secondary_Qty, .Primary_Product, .Batch_No, .Expiry_Date, CInt(TipeProduct))
                End With
            Next

            'load data item baru
            For Each rowDGV In vDGVItem.Rows
                BatchNo = rowDGV.Cells(vFrm.ColLotNumber.Index).Value
                PrimaryLocation = rowDGV.Cells(vFrm.ColPrimaryLocation.Index).Value
                If IsNumeric(rowDGV.Cells(vFrm.ColSecQty.Index).Value) = True Then
                    Sec_Qty = rowDGV.Cells(vFrm.ColSecQty.Index).Value
                Else
                    Sec_Qty = "-1"
                End If
                If IsDate(rowDGV.Cells(vFrm.ColExpiry.Index).Value) = False Then
                    ExpiryDate = #1/1/1950#
                Else
                    ExpiryDate = CDate(rowDGV.Cells(vFrm.ColExpiry.Index).Value)
                End If
                PrimaryProduct = rowDGV.Cells(vFrm.ColPrimaryCodeNo.Index).Value
                DTAMasterProductQty.FillBy_Parameter2(TBLMasterProductQty, _
                Sec_Qty, PrimaryProduct, BatchNo, ExpiryDate, CInt(TipeProduct))
            Next

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
                BatchNo = rowDGV.Cells(vFrm.ColLotNumber.Index).Value
                PrimaryLocation = rowDGV.Cells(vFrm.ColPrimaryLocation.Index).Value
                If IsNumeric(rowDGV.Cells(vFrm.ColSecQty.Index).Value) = True Then
                    Sec_Qty = rowDGV.Cells(vFrm.ColSecQty.Index).Value
                Else
                    Sec_Qty = "-1"
                End If
                If IsDate(rowDGV.Cells(vFrm.ColExpiry.Index).Value) = False Then
                    ExpiryDate = #1/1/1950#
                Else
                    ExpiryDate = CDate(rowDGV.Cells(vFrm.ColExpiry.Index).Value)
                End If
                PrimaryProduct = rowDGV.Cells(vFrm.ColPrimaryCodeNo.Index).Value
                Qty = rowDGV.Cells(vFrm.ColTotalQty.Index).Value

                Me.Koreksi4SimpanOut(TBLMasterProductQty, _
                PrimaryProduct, Qty, Sec_Qty, PrimaryLocation, _
                BatchNo, ExpiryDate)
            Next

            'cek apakah ada qty yang minus setelah dilakukan simulasi, jika ada
            'maka return false
            ListPrimaryProduct.Clear()
            For Each rowProduct In TBLMasterProductQty.Rows
                If rowProduct.Qty < 0 Then
                    If ObjProd.ApakahGoods(rowProduct.Primary_Product) = True Then
                        ListPrimaryProduct.Enqueue(rowProduct.Primary_Product)
                    End If
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

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            DTAMasterProductQty.Dispose()
            TBLMasterProductQty.Dispose()
            ListPrimaryProduct = Nothing
            ObjProd = Nothing
        End Try
    End Function

    Public Function Koreksi4SimpanOut(ByVal TBLMasterProduct As Core_Module.dseMasterProduct.tblmaster_productqtyDataTable, ByVal vPrimaryProduct As _
Double, ByVal Qty As Decimal, Optional ByVal vSecondaryQty As Decimal = -1, _
Optional ByVal vPrimaryLocation As Double = -1, Optional ByVal _
vBatchNo As String = "", Optional ByVal vExpiryDate As Date = #1/1/1950#) As Boolean
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

    Public Function Koreksi4HapusOut(ByVal TBLMasterProduct As Core_Module.dseMasterProduct.tblmaster_productqtyDataTable, _
ByVal vPrimaryProduct As Double, ByVal Qty As Decimal, Optional ByVal vSecondaryQty _
As Double = -1, Optional ByVal vPrimaryLocation As Double = -1, Optional ByVal _
vBatchNo As String = "", Optional ByVal vExpiryDate As Date = _
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
                    Me.SetJml_Simulasi(TBLMasterProduct, vPrimaryProduct, _
                    vSecondaryQty, vPrimaryLocation, vBatchNo, JmlBaru)
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
    ByVal vPrimaryProduct As Double, ByVal vSecQty As Double, _
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
    ByVal vPrimaryProduct As Double, ByVal vSecQty As Double, _
    ByVal vPrimaryLocation As Double, ByVal vBatchNo As String) As Double
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
    ByVal vPrimaryProduct As Double, ByVal vSecQty As Double, _
    ByVal vPrimaryLocation As Double, ByVal vBatchNo As String, _
    ByVal JmlBaru As Double)
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

    Private Function GetPrimaryMax_TBLQtySimulasi(ByVal TBL As _
Core_Module.dseMasterProduct.tblmaster_productqtyDataTable) As Double
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

    Public Function GetPrimaryDO_QuickSales(ByVal vPrimaryQuickSales As Double) As Double
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_itemTableAdapter
        Dim TBL As New DseTransaksiSalesInvoice.tbltrans_salesinvoice_itemDataTable
        Dim row As DseTransaksiSalesInvoice.tbltrans_salesinvoice_itemRow

        Try
            DTA.FillBy_PrimaryInvoice(TBL, vPrimaryQuickSales)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                Return row.Primary_DO
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function


    Public Function GetNoFakturPajakMax(ByVal vPrimaryFakturPajak As Double, ByVal vPrimaryInvoiceException As Double) As String
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice1TableAdapter
        Dim Primary, UrutanMax As Double

        Try
            If DTA.ScalarQuery_GetUrutanFakturPajakMax2(vPrimaryInvoiceException, vPrimaryFakturPajak).HasValue Then
                UrutanMax = DTA.ScalarQuery_GetUrutanFakturPajakMax2(vPrimaryInvoiceException, vPrimaryFakturPajak).Value
                If DTA.ScalarQuery_GetPrimary2(UrutanMax, vPrimaryFakturPajak).HasValue = True Then
                    Primary = DTA.ScalarQuery_GetPrimary2(UrutanMax, vPrimaryFakturPajak).Value
                    If String.IsNullOrEmpty(DTA.ScalarQuery_GetNoFakturPajak(Primary)) = True Then
                        Return ""
                    Else
                        Return DTA.ScalarQuery_GetNoFakturPajak(Primary)
                    End If
                Else
                    Return ""
                End If
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

    Public Function Hapus(ByVal vPrimaryTransaksi As Double, _
    ByVal UserCode As String, ByVal StrictSalesStaffDelete As Boolean, ByVal QuickSales As Boolean) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            SW.Start()

            'ambil no antrian
            If QuickSales = True Then
                Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, vPrimaryTransaksi, _
               Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiInvoice_Quick_Hapus, UserCode)
            Else
                Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, vPrimaryTransaksi, _
               Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiInvoice_Hapus, UserCode)
            End If

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > ObjAntrian.GetWaitTime Then
                    If SW.Elapsed.TotalSeconds >= ObjAntrian.GetWaitTime Then
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                        Exit Function
                    Else
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.Hapus1(vPrimaryTransaksi, _
                            UserCode, StrictSalesStaffDelete, QuickSales)
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

    Public Overridable Function Hapus1(ByVal vPrimaryTransaksi As Double, _
    ByVal UserCode As String, ByVal StrictSalesStaffDelete As Boolean, ByVal QuickInvoice As Boolean, Optional ByVal IsQuickUpdate As Boolean = False) As Boolean
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice1TableAdapter
        Dim DTAObs As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_obseleteTableAdapter
        Dim DTA_Inv_DO As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_invoice_doTableAdapter
        Dim TBL_Inv_DO As New DseTransaksiSalesInvoice.tbltrans_invoice_doDataTable
        Dim rowInvDO As DseTransaksiSalesInvoice.tbltrans_invoice_doRow
        Dim TipeInvoice As EnumTipeInvoice

        Dim Revisi As Double = 0
        Dim Inv_Conflict As String = ""

        Dim PrimarySO, PrimaryDO, PrimaryStockOut As Double

        Dim TBLItem As New DseTransaksiSalesInvoice.tbltrans_salesinvoice_itemDataTable
        Dim DTAItem As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_itemTableAdapter
        Dim rowItem As DseTransaksiSalesInvoice.tbltrans_salesinvoice_itemRow
        Dim ObjDO As New ClsTransaksiDO
        Dim ij As New FrmTransaksiStockOutSales

        Dim PrimaryFakturPajak As Double = -1
        Dim NoFakturPajakMax As String = ""
        Dim Obj As New ClsTransaksiSalesInvoice_SettingTaxInvoice
        Dim TblItemCustom As New CustomSMESS.dseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable

        Dim ObjInv_C As New CustomSMESS.ClsTransaksiSalesInvoice

        Try
            If Me.ApakahBolehHapus(vPrimaryTransaksi, StrictSalesStaffDelete, _
            UserCode, , IsQuickUpdate) = True Then
                '2025 Core Tax Tahap 2.2, dikomen
                'PrimaryFakturPajak = DTA.ScalarQuery_GetPrimaryFakturPajak(vPrimaryTransaksi)
                'NoFakturPajakMax = Me.GetNoFakturPajakMax(PrimaryFakturPajak, vPrimaryTransaksi)

                TipeInvoice = DTA.ScalarQuery_GetTipeInvoice(vPrimaryTransaksi)
                Revisi = DTA.ScalarQuery_GetRevisi(vPrimaryTransaksi).Value

                'hapuskan data so, do, stock out untuk quick sales
                If QuickInvoice = True Then
                    DTAItem.FillBy_PrimaryInvoice(TBLItem, vPrimaryTransaksi)
                    If TBLItem.Rows.Count > 0 Then
                        rowItem = TBLItem.Rows(0)
                        PrimarySO = rowItem.Primary_SO
                        PrimaryDO = rowItem.Primary_DO
                        PrimaryStockOut = ObjDO.GetPrimaryStockOut(PrimaryDO)
                    End If
                End If

                If Revisi = 0 Then
                    DTA_Inv_DO.FillBy_PrimaryInvoice(TBL_Inv_DO, vPrimaryTransaksi)
                    'hapus data header, data item otomatis dihapuskan oleh server
                    If DTA.Delete(vPrimaryTransaksi) > 0 Then
                        '2025 Core Tax Tahap 2.2, dikomen
                        'set urutan faktur pajak max
                        'Obj.SetNoUrutMaksimal(PrimaryFakturPajak, NoFakturPajakMax)

                        'hapus data jurnal
                        If TipeInvoice = EnumTipeInvoice.SalesInvoice Then
                            Me.Hapus_Jurnal(vPrimaryTransaksi)
                        Else
                            Me.Hapus_Jurnal_DP(vPrimaryTransaksi)
                        End If
                        'tandai do belum buat inv
                        For Each rowInvDO In TBL_Inv_DO.Rows
                            DTA_Inv_DO.UpdateQuery_SetSudahBuatInvoice(0, -1, rowInvDO.Primary_DO)
                        Next

                        If QuickInvoice = True Then
                            Dim ObjStock As New Inventory_Module.clsTransaksiStockInOut
                            ObjStock.CariDetail(PrimaryStockOut, ij.dgvItem, TblItemCustom)
                            ObjStock.Hapus1(PrimaryStockOut, UserCode, TblItemCustom)
                            ObjStock = Nothing

                            ObjDO.Hapus1(PrimaryDO, UserCode, False)

                            Dim ObjSO As New ClsTransaksiSalesOrder
                            ObjSO.Hapus(PrimarySO, UserCode, False)
                            ObjSO = Nothing
                        End If

                        'Update SMESS 2026
                        If IsQuickUpdate = False Then
                            If TipeInvoice = EnumTipeInvoice.SalesInvoice Then
                                ObjInv_C.AfterHapusInvoice(CustomSMESS.ClsTransaksiSalesInvoice.EnumDeleteType.Invoice, vPrimaryTransaksi, Date.Now, True)
                            Else
                                ObjInv_C.AfterHapusInvoice(CustomSMESS.ClsTransaksiSalesInvoice.EnumDeleteType.Invoice_DP, vPrimaryTransaksi, Date.Now, True)
                            End If
                        End If

                        Return True
                    End If
                Else
                    'jika ada conflict untuk revisi sebelum dengan invoice yang sudah ada, maka hapus semua revisi yang lama
                    'tanpa kecuali
                    Inv_Conflict = Me.Apakah_Revisi_Sebelumnya_Conflict(vPrimaryTransaksi)
                    If Inv_Conflict = "" Then
                        'revisi sebelumnya tidak ada conflict
                        If MsgBox("Delete all revision history? By clicking yes, all previous invoice revision will all be delete.", _
                        MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                            'hapus data jurnal
                            If TipeInvoice = EnumTipeInvoice.SalesInvoice Then
                                Me.Hapus_Jurnal(vPrimaryTransaksi)
                            Else
                                Me.Hapus_Jurnal_DP(vPrimaryTransaksi)
                            End If
                            ''hapus semua revisi
                            'Return Me.HapusRevisi(vPrimaryTransaksi)
                        Else
                            Dim NoTransaksiAcuan As String = DTA.ScalarQuery_GetNoTransaksi(vPrimaryTransaksi)
                            NoTransaksiAcuan = NoTransaksiAcuan.Substring(0, NoTransaksiAcuan.Length - 1)
                            'hapus data invoice aktif
                            DTA_Inv_DO.FillBy_PrimaryInvoice(TBL_Inv_DO, vPrimaryTransaksi)
                            DTA.Delete(vPrimaryTransaksi)
                            'hapus data jurnal
                            If TipeInvoice = EnumTipeInvoice.SalesInvoice Then
                                Me.Hapus_Jurnal(vPrimaryTransaksi)
                            Else
                                Me.Hapus_Jurnal_DP(vPrimaryTransaksi)
                            End If
                            For Each rowInvDO In TBL_Inv_DO.Rows
                                DTA_Inv_DO.UpdateQuery_SetSudahBuatInvoice(0, -1, rowInvDO.Primary_DO)
                            Next
                            '2025 Core Tax Tahap 2.2, dikomen
                            'set urutan faktur pajak max
                            'Obj.SetNoUrutMaksimal(PrimaryFakturPajak, NoFakturPajakMax)
                            'hapus semua revisi
                            DTAObs.DeleteQueryAllRevisi(NoTransaksiAcuan + "%")

                            'Update SMESS 2026
                            If IsQuickUpdate = False Then
                                If TipeInvoice = EnumTipeInvoice.SalesInvoice Then
                                    ObjInv_C.AfterHapusInvoice(CustomSMESS.ClsTransaksiSalesInvoice.EnumDeleteType.Invoice, vPrimaryTransaksi, Date.Now, True)
                                Else
                                    ObjInv_C.AfterHapusInvoice(CustomSMESS.ClsTransaksiSalesInvoice.EnumDeleteType.Invoice_DP, vPrimaryTransaksi, Date.Now, True)
                                End If
                            End If

                            Return True
                        End If
                    Else
                        'ada conflict do dengan revisi sebelumnya, maka hapus semua data tanpa pandang bulu ^^
                        If MsgBox("The previous revision of this invoice conflict the following invoice:" + Inv_Conflict + ", regarding delivery order." + _
                        "Deleting this data will delete all previous revision of the invoice, continue?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                            Dim NoTransaksiAcuan As String = DTA.ScalarQuery_GetNoTransaksi(vPrimaryTransaksi)
                            NoTransaksiAcuan = NoTransaksiAcuan.Substring(0, NoTransaksiAcuan.Length - 1)
                            'hapus data invoice aktif
                            DTA_Inv_DO.FillBy_PrimaryInvoice(TBL_Inv_DO, vPrimaryTransaksi)
                            DTA.Delete(vPrimaryTransaksi)
                            For Each rowInvDO In TBL_Inv_DO.Rows
                                DTA_Inv_DO.UpdateQuery_SetSudahBuatInvoice(0, -1, rowInvDO.Primary_DO)
                            Next
                            'hapus data jurnal
                            If TipeInvoice = EnumTipeInvoice.SalesInvoice Then
                                Me.Hapus_Jurnal(vPrimaryTransaksi)
                            Else
                                Me.Hapus_Jurnal_DP(vPrimaryTransaksi)
                            End If
                            '2025 Core Tax Tahap 2.2, dikomen
                            'set urutan faktur pajak max
                            'Obj.SetNoUrutMaksimal(PrimaryFakturPajak, NoFakturPajakMax)
                            'hapus semua revisi
                            DTAObs.DeleteQueryAllRevisi(NoTransaksiAcuan + "%")

                            'Update SMESS 2026
                            If IsQuickUpdate = False Then
                                If TipeInvoice = EnumTipeInvoice.SalesInvoice Then
                                    ObjInv_C.AfterHapusInvoice(CustomSMESS.ClsTransaksiSalesInvoice.EnumDeleteType.Invoice, vPrimaryTransaksi, Date.Now, True)
                                Else
                                    ObjInv_C.AfterHapusInvoice(CustomSMESS.ClsTransaksiSalesInvoice.EnumDeleteType.Invoice_DP, vPrimaryTransaksi, Date.Now, True)
                                End If
                            End If

                            Return True
                        Else
                            'Update SMESS 2026
                            If IsQuickUpdate = False Then
                                If TipeInvoice = EnumTipeInvoice.SalesInvoice Then
                                    ObjInv_C.AfterHapusInvoice(CustomSMESS.ClsTransaksiSalesInvoice.EnumDeleteType.Invoice, vPrimaryTransaksi, Date.Now, False)
                                Else
                                    ObjInv_C.AfterHapusInvoice(CustomSMESS.ClsTransaksiSalesInvoice.EnumDeleteType.Invoice_DP, vPrimaryTransaksi, Date.Now, False)
                                End If
                            End If

                            Return False
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA_Inv_DO.Dispose()
            TBL_Inv_DO.Dispose()
            ObjDO = Nothing
            Obj = Nothing
        End Try
    End Function

    Private Function Apakah_Revisi_Sebelumnya_Conflict(ByVal vPrimaryTransaksi As Double) As String
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice1TableAdapter
        Dim DTAObs As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_obseleteTableAdapter
        Dim NoTransaksiAcuan As String = ""
        Dim PrimaryObselete As Double = 0
        Dim DTA_DO_Obs As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoiceobselete_item1TableAdapter
        Dim TBL_DO_Obs As New DseTransaksiSalesInvoice.tbltrans_salesinvoiceobselete_item1DataTable
        Dim row_DO_Obs As DseTransaksiSalesInvoice.tbltrans_salesinvoiceobselete_item1Row
        Dim DTA1 As New DseTransaksiSalesInvoiceTableAdapters.DataTable10TableAdapter
        Dim TBL As New DseTransaksiSalesInvoice.DataTable10DataTable
        Dim row As DseTransaksiSalesInvoice.DataTable10Row
        Dim Retval As String = ""

        Try
            NoTransaksiAcuan = DTA.ScalarQuery_GetNoTransaksi(vPrimaryTransaksi)
            NoTransaksiAcuan = NoTransaksiAcuan.Substring(0, NoTransaksiAcuan.Length - 1)
            PrimaryObselete = DTAObs.ScalarQuery_GetMaxPrimary(NoTransaksiAcuan + "%").Value

            'tarik semua data DO untuk invoice revisi lama
            DTA_DO_Obs.Fill(TBL_DO_Obs, PrimaryObselete)

            'cek untuk semua item tersebut apakah ada yang sudah berinvoice, dengan mengecualikan 
            'invoice revisi baru yang akan dihapus menjadi revisi lama
            For Each row_DO_Obs In TBL_DO_Obs.Rows
                If DTA1.Fill(TBL, row_DO_Obs.Primary_DO) > 0 Then
                    For Each row In TBL.Rows
                        'untuk transaksi edit, tidak boleh memperhitungkan dirinya sendiri
                        If vPrimaryTransaksi <> row.Primary_Invoice Then
                            If Retval = "" Then
                                Retval = row.No_Transaksi
                            Else
                                Retval = Retval + ", " + row.No_Transaksi
                            End If
                        End If
                    Next
                End If
            Next

            'kalo sampe sini return true
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            DTAObs.Dispose()
            DTA_DO_Obs.Dispose()
            TBL_DO_Obs.Dispose()
            DTA1.Dispose()
            TBL.Dispose()
        End Try
    End Function

    'Public Function HapusRevisi(ByVal vPrimaryTransaksi As Double) As Boolean
    '    Dim NoTransaksiAcuan As String = ""
    '    Dim PrimaryObselete As Double = 0

    '    Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice1TableAdapter
    '    Dim DTAObs As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_obseleteTableAdapter
    '    Dim DTA_Inv_DO As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_invoice_doTableAdapter
    '    Dim TBL_Inv_DO As New DseTransaksiSalesInvoice.tbltrans_invoice_doDataTable
    '    Dim rowInvDO As DseTransaksiSalesInvoice.tbltrans_invoice_doRow
    '    Dim PrimaryFakturPajak As Double = -1
    '    Dim NoFakturPajakMax As String = ""
    '    Dim Obj As New ClsTransaksiSalesInvoice_SettingTaxInvoice

    '    Try
    '        'cari primary obselete
    '        NoTransaksiAcuan = DTA.ScalarQuery_GetNoTransaksi(vPrimaryTransaksi)
    '        NoTransaksiAcuan = NoTransaksiAcuan.Substring(0, NoTransaksiAcuan.Length - 1)
    '        PrimaryObselete = DTAObs.ScalarQuery_GetMaxPrimary(NoTransaksiAcuan + "%").Value

    '        'set urutan maksimal fp
    '        PrimaryFakturPajak = DTA.ScalarQuery_GetPrimaryFakturPajak(vPrimaryTransaksi)
    '        NoFakturPajakMax = Me.GetNoFakturPajakMax(PrimaryFakturPajak, vPrimaryTransaksi)
    '        Obj.SetNoUrutMaksimal(PrimaryFakturPajak, NoFakturPajak)

    '        'hapuskan data aktif
    '        DTA_Inv_DO.FillBy_PrimaryInvoice(TBL_Inv_DO, vPrimaryTransaksi)
    '        'tandai do belum buat inv
    '        For Each rowInvDO In TBL_Inv_DO.Rows
    '            DTA_Inv_DO.UpdateQuery_SetSudahBuatInvoice(0, -1, rowInvDO.Primary_DO)
    '        Next
    '        DTA.Delete(vPrimaryTransaksi)

    '        'simpan data obselete ke data aktif
    '        Me.SimpanDataInvoiceRevisiLama(PrimaryObselete)

    '        'hapuskan data obselete yang dipindahkan
    '        DTAObs.Delete(PrimaryObselete)

    '        'nilai balikan
    '        Return True
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTA_Inv_DO.Dispose()
    '        TBL_Inv_DO.Dispose()
    '        DTA.Dispose()
    '        DTAObs.Dispose()
    '        Obj = Nothing
    '    End Try
    'End Function

    Public Overridable Function IsiDGV(ByVal Criteria As String, _
   ByVal DGV As DataGridView, ByVal TBL As DseTransaksiSalesInvoice.tbltrans_salesinvoice_viewDataTable, _
   ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, ByVal TipeInvoice As EnumTipeInvoice, ByVal vIsQuick As Boolean, _
   Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
   Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Boolean
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_viewTableAdapter
        Dim Row As DseTransaksiSalesInvoice.tbltrans_salesinvoice_viewRow
        Dim QueryResult As New Double
        Dim StartT As Date
        Try
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            DGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                If TipeInvoice = EnumTipeInvoice.InvoiceDP Then
                    QueryResult = DTA.FillBy_NoTransaksiLIKE2(TBL, TipeInvoice, Criteria + "%")
                Else
                    QueryResult = DTA.FillBy_NoTransaksiLIKE(TBL, TipeInvoice, Criteria + "%")
                End If

            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                If TipeInvoice = EnumTipeInvoice.InvoiceDP Then
                    QueryResult = DTA.FillBy_TanggalTransaksi2(TBL, TipeInvoice, vTanggalSampai, vTanggalDari)
                Else
                    QueryResult = DTA.FillBy_TanggalTransaksi(TBL, TipeInvoice, vTanggalSampai, vTanggalDari)
                End If

            ElseIf SearchCriteria = EnumSearchCriteria.TaxInvoiceNo Then
                If TipeInvoice = EnumTipeInvoice.InvoiceDP Then
                    QueryResult = DTA.FillBy_NoTaxInvoiceLIKE2(TBL, TipeInvoice, Criteria + "%")
                Else
                    QueryResult = DTA.FillBy_NoTaxInvoiceLIKE(TBL, TipeInvoice, Criteria + "%")
                End If

            ElseIf SearchCriteria = EnumSearchCriteria.ByKodeCustomer Then
                If TipeInvoice = EnumTipeInvoice.InvoiceDP Then
                    QueryResult = DTA.FillBy_CustCodeLIKE2(TBL, TipeInvoice, Criteria + "%")
                Else
                    QueryResult = DTA.FillBy_CustCodeLIKE(TBL, TipeInvoice, Criteria + "%")
                End If

            ElseIf SearchCriteria = EnumSearchCriteria.ByNamaCustomer Then
                If TipeInvoice = EnumTipeInvoice.InvoiceDP Then
                    QueryResult = DTA.FillBy_CustName2(TBL, TipeInvoice, Criteria + "%")
                Else
                    QueryResult = DTA.FillBy_CustName(TBL, TipeInvoice, Criteria + "%")
                End If

            ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                QueryResult = DTA.FillBy_ProductCodeLIKE(TBL, TipeInvoice, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                QueryResult = DTA.FillBy_ProductDescLIKE(TBL, TipeInvoice, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                QueryResult = DTA.FillBy_ProductSpec(TBL, TipeInvoice, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByDeliveryOrderNo Then
                QueryResult = DTA.FillBy_DONumberLIKE(TBL, TipeInvoice, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.BySalesPersonCode Then
                If vIsQuick = True Then
                    QueryResult = DTA.FillBy_SalesPersonCodeLIKE1(TBL, Criteria + "%")
                Else
                    If TipeInvoice = EnumTipeInvoice.InvoiceDP Then
                        QueryResult = DTA.FillBy_SalesPersonCodeLIKE2(TBL, Criteria + "%")
                    Else
                        QueryResult = DTA.FillBy_SalesPersonCodeLIKE3(TBL, Criteria + "%")
                    End If
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.BySalesPersonName Then
                If vIsQuick = True Then
                    QueryResult = DTA.FillBy_SalesPersonDescLIKE1(TBL, Criteria + "%")
                Else
                    If TipeInvoice = EnumTipeInvoice.InvoiceDP Then
                        QueryResult = DTA.FillBy_SalesPersonDescLIKE2(TBL, Criteria + "%")
                    Else
                        QueryResult = DTA.FillBy_SalesPersonDescLIKE3(TBL, Criteria + "%")
                    End If
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.BySalesOrderNo Then
                If TipeInvoice = EnumTipeInvoice.InvoiceDP Then
                    QueryResult = DTA.FillBy_SONumberLIKE2(TBL, TipeInvoice, Criteria + "%")
                Else
                    QueryResult = DTA.FillBy_SONumberLIKE(TBL, TipeInvoice, Criteria + "%")
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByCustomerShiptoCode Then
                If TipeInvoice = EnumTipeInvoice.InvoiceDP Then
                    QueryResult = DTA.FillBy_ShiptoCustomerCode2(TBL, TipeInvoice, Criteria + "%")
                Else
                    QueryResult = DTA.FillBy_ShiptoCustomerCode(TBL, TipeInvoice, Criteria + "%")
                End If

            ElseIf SearchCriteria = EnumSearchCriteria.ByCustomerShiptoName Then
                If TipeInvoice = EnumTipeInvoice.InvoiceDP Then
                    QueryResult = DTA.FillBy_ShiptoCustomerName2(TBL, TipeInvoice, Criteria + "%")
                Else
                    QueryResult = DTA.FillBy_ShiptoCustomerName(TBL, TipeInvoice, Criteria + "%")
                End If

            End If

            'kosongkan data shipto
            For Each Row In TBL.Rows
                If Row.IsShip_to_Customer_CodeNull = False Then
                    If Row.Customer_Code = Row.Ship_to_Customer_Code Then
                        Row.Ship_to_Customer_Code = ""
                        Row.Ship_to_Customer_Name = ""
                    End If
                End If
            Next

            TBL.DefaultView.RowFilter = Nothing
            If Criteria <> "" Then
                Criteria.Substring(Criteria.Length - 1, 1).Replace("'", "").Replace("%", "")
                If SearchCriteria = EnumSearchCriteria.ByCustomerShiptoCode Then
                    TBL.DefaultView.RowFilter = "[Ship to customer Code] LIKE '" & Criteria & "%'"
                ElseIf SearchCriteria = EnumSearchCriteria.ByCustomerShiptoName Then
                    TBL.DefaultView.RowFilter = "[Ship to customer Name] LIKE '" & Criteria & "%'"
                End If
                QueryResult = TBL.DefaultView.Count
            End If


            '
            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(2).DefaultCellStyle.Format = "dd MMM yyyy"
                DGV.Columns(5).DisplayIndex = 7
                DGV.Columns(6).DisplayIndex = 5
                DGV.Columns(7).DisplayIndex = 6
            End If

            Dim TS As TimeSpan = Now - StartT
            If QueryResult > 0 Then
                If TS.TotalSeconds > 1 Then
                    tsl.Text = CStr(QueryResult) + " result(s) in " + ObjCD.DigitGrouping(TS.TotalSeconds, 2) + " seconds"
                Else
                    tsl.Text = CStr(QueryResult) + " result(s) in " + ObjCD.DigitGrouping(TS.TotalMilliseconds, 2, MidpointRounding.AwayFromZero) + " miliseconds"
                End If
            Else
                tsl.Text = "No result found"
            End If
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            DGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Sub ResetObject()
        Try
            Me.mPrimarySO4DP = -1
            mPrimary = -1
            mNoUrut = 0
            mUrutanFakturPajak = 0
            mNoTransaksi = ""
            Me.mCalculatesBy = EnumCalculatesBy.Invoice
            mTanggalTransaksi = Now
            mPrimaryCust = -1
            mIdentifier_CustAddress = ""
            mPaymentTerm = ""
            mRemarks = ""
            mPrimaryCurrency = 1
            mSubTotalB4DP = 0
            mDownPayment = 0
            mSubTotal = 0
            mUseTax = 0
            mTaxPercentage = 0
            mTaxExchangeRate = 1
            mDutyStamp = 0
            mTaxAmount = 0
            mTotal = 0
            mJenisFakturPajak = EnumJenisFakturPajak.FakturPajakSederhana
            mExchangeRateToUSD = 1
            mExhangeRateToIDR = 1
            mNoNPWP = ""
            mNoFakturPajak = ""
            mNoSeriFakturPajak = ""
            Me.mAlamatPajak1 = ""
            Me.mAlamatPajak2 = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub ResetObject(ByVal vDgvItem As DataGridView, ByVal vdgvItemDP As DataGridView, _
    ByVal vDGVExchangeRate As DataGridView, ByVal vDgvItemSR As DataGridView)
        Try
            Me.mPrimarySO4DP = -1
            mPrimary = -1
            mNoUrut = 0
            mUrutanFakturPajak = 0
            mNoTransaksi = ""
            Me.mCalculatesBy = EnumCalculatesBy.Invoice
            mTanggalTransaksi = Now
            mPrimaryCust = -1
            mIdentifier_CustAddress = ""
            mPaymentTerm = ""
            mRemarks = ""
            mPrimaryCurrency = 1
            mSubTotalB4DP = 0
            mDownPayment = 0
            mSubTotal = 0
            mUseTax = 0
            mTaxPercentage = 0
            mTaxExchangeRate = 1
            mDutyStamp = 0
            mTaxAmount = 0
            mTotal = 0
            mJenisFakturPajak = EnumJenisFakturPajak.FakturPajakSederhana
            mExchangeRateToUSD = 1
            mExhangeRateToIDR = 1
            mNoNPWP = ""
            mNoFakturPajak = ""
            mNoSeriFakturPajak = ""
            Me.mAlamatPajak1 = ""
            Me.mAlamatPajak2 = ""

            vDgvItem.Rows.Clear()
            vdgvItemDP.Rows.Clear()
            vDGVExchangeRate.Rows.Clear()
            vDgvItemSR.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub Get_PrimaryCurrency_TaxPercent_TaxExclude_TaxExchangeRate(ByVal vPrimaryInvoice As Double, _
    ByRef vTaxExclude As Boolean, ByRef vTaxPercent As Double, ByRef vTaxExchangeRate As Double, _
    ByRef vPrimaryCurrency As Double)
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice1TableAdapter
        Dim TBL As New DseTransaksiSalesInvoice.tbltrans_salesinvoice1DataTable
        Dim row As DseTransaksiSalesInvoice.tbltrans_salesinvoice1Row

        Try
            'fil basis data
            If DTA.FillBy_Primary(TBL, vPrimaryInvoice) > 0 Then
                'salin data header
                row = TBL.Rows(0)

                With row
                    vPrimaryCurrency = .Primary_Currency
                    vTaxExclude = .Tax_Exclude
                    vTaxExchangeRate = .Tax_Exchange_Rate
                    vTaxPercent = .Tax
                End With
            Else
                vPrimaryCurrency = -1
                vTaxExclude = False
                vTaxExchangeRate = 1
                vTaxPercent = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Public Function GetPrimaryCurrency(ByVal vPrimaryInvoice As Double) As Double
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoiceTableAdapter
        Dim Retval As Nullable(Of Double)

        Try
            Retval = DTA.ScalarQuery_GetPrimaryCurrency(vPrimaryInvoice)
            If Retval.HasValue = True Then
                Return Retval
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

    Public Function GetTaxInvoiceNumber(ByVal vPrimaryInvoice As Double) As String
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoiceTableAdapter
        Dim Retval As String = ""

        Try
            Retval = DTA.ScalarQuery_GetTaxInvoiceNumber(vPrimaryInvoice)
            If String.IsNullOrEmpty(Retval) = True Then
                Return ""
            Else
                Return Retval
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTipePerhitungan(ByVal vPrimaryInvoice As Double) As EnumCalculatesBy
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice1TableAdapter

        Try
            If DTA.ScalarQuery_GetCalculateBy(vPrimaryInvoice).HasValue = True Then
                Return DTA.ScalarQuery_GetCalculateBy(vPrimaryInvoice).Value
            Else
                Return EnumCalculatesBy.Invoice
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return EnumCalculatesBy.Invoice
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryInvoice_ByPrimaryDO(ByVal vPrimaryDO As String) As Double
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_itemTableAdapter

        Try
            If DTA.ScalarQuery_GetPrimaryInvoice_ByPrimaryDO(vPrimaryDO).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryInvoice_ByPrimaryDO(vPrimaryDO).Value
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

    Public Function GetPrimaryInvoice(ByVal vNoTransaksi As String) As Double
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice1TableAdapter

        Try
            If DTA.ScalarQuery_GetPrimaryByNoTransaksi(vNoTransaksi).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryByNoTransaksi(vNoTransaksi).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
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
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice1TableAdapter

        Try
            If DTA.ScalarQuery_GetPrimaryByNoTransaksiLIKE(vNoTransaksiLIKE + "%").HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryByNoTransaksiLIKE(vNoTransaksiLIKE + "%").Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub GetNoInvoice(ByVal vPrimarySO As Double, ByVal vPrimaryProduct As Double, ByVal vListRetVal As List(Of String))
        Dim DTA As New DseTransaksiSalesInvoice_TransfJurnalTableAdapters.tbltrans_salesinvoiceTableAdapter
        Dim TBL As New DseTransaksiSalesInvoice_TransfJurnal.tbltrans_salesinvoiceDataTable
        Dim row As DseTransaksiSalesInvoice_TransfJurnal.tbltrans_salesinvoiceRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            TBL.Rows.Clear()
            vListRetVal.Clear()

            DTA.Fill(TBL, vPrimaryProduct, vPrimarySO)

            For Each row In TBL.Rows
                vListRetVal.Add(row.No_Transaksi)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Public Function GetNoInvoice(ByVal vPrimaryInvoice As Double) As String
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoiceTableAdapter

        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetNoInvoice(vPrimaryInvoice)) = True Then
                Return ""
            Else
                Return DTA.ScalarQuery_GetNoInvoice(vPrimaryInvoice)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryInvoiceByNoFakturPajak(ByVal vNoFakturPajak As String) As Double
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoiceTableAdapter

        Try
            If DTA.ScalarQuery_GetPrimaryInvoice(vNoFakturPajak).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryInvoice(vNoFakturPajak).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahSudahAdaInvoice(ByVal vPrimaryDO As Double) As Boolean
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_itemTableAdapter

        Try
            If DTA.ScalarQuery_GetJmlByPrimaryDO(vPrimaryDO).HasValue = True Then
                If DTA.ScalarQuery_GetJmlByPrimaryDO(vPrimaryDO).Value = 0 Then
                    Return False
                Else
                    Return True
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

    Public Function GetJml(ByVal Criteria As String, _
   ByVal SearchCriteria As EnumSearchCriteria, ByVal TipeInvoice As EnumTipeInvoice, ByVal vIsQuick As Boolean, Optional _
   ByVal vTanggalDari As Date = (#1/1/1950#), Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Long
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_viewTableAdapter
        ' ahdar tambahan find by ship to customer
        Try
            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                Return DTA.ScalarQuery_GetJmlByNoTransaksiLIKE(TipeInvoice, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                Return DTA.ScalarQuery_GetJmlByTanggalTransaksi(TipeInvoice, vTanggalSampai, vTanggalDari).Value
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                Return DTA.ScalarQuery_GetJmlByProductCodeLIKE(TipeInvoice, Criteria + "%").Value
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                Return DTA.ScalarQuery_GetJmlByProductDescLIKE(TipeInvoice, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                Return DTA.ScalarQuery_GetJmlByProductSpecLIKE(TipeInvoice, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByKodeCustomer Then
                Return DTA.ScalarQuery_GetJmlByCustCodeLIKE(TipeInvoice, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByNamaCustomer Then
                Return DTA.ScalarQuery_GetJmlByCustNameLIKE(TipeInvoice, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByNamaCustomer Then
                Return DTA.ScalarQuery_GetJmlByCustNameLIKE(TipeInvoice, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.TaxInvoiceNo Then
                Return DTA.ScalarQuery_GetJmlByTaxInvoiceLIKE(TipeInvoice, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByDeliveryOrderNo Then
                Return DTA.ScalarQuery_GetJmlBy_DONumberLIKE(TipeInvoice, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.BySalesPersonCode Then

                If vIsQuick = True Then
                    'quikc invoice, cari langsung dari header
                    Return DTA.ScalarQuery_GetJmlBySalesPersonCode1(Criteria + "%")
                Else
                    If TipeInvoice = EnumTipeInvoice.InvoiceDP Then
                        Return DTA.ScalarQuery_GetJmlBySalesPersonCode2(Criteria + "%")
                    Else
                        Return DTA.ScalarQuery_GetJmlBySalesPersonCode3(Criteria + "%")
                    End If
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.BySalesPersonName Then
                If vIsQuick = True Then
                    'quikc invoice, cari langsung dari header
                    Return DTA.ScalarQuery_GetJmlBySalesPersonDesc1(Criteria + "%")
                Else
                    If TipeInvoice = EnumTipeInvoice.InvoiceDP Then
                        Return DTA.ScalarQuery_GetJmlBySalesPersonDesc2(Criteria + "%")
                    Else
                        Return DTA.ScalarQuery_GetJmlBySalesPersonDesc3(Criteria + "%")
                    End If
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.BySalesOrderNo Then
                If TipeInvoice = EnumTipeInvoice.InvoiceDP Then
                    Return DTA.ScalarQuery_GetJmlBySONoLIKE2(TipeInvoice, Criteria + "%")
                Else
                    Return DTA.ScalarQuery_GetJmlBySONoLIKE(TipeInvoice, Criteria + "%")
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByCustomerShiptoCode Then
                Return DTA.ScalarQuery_GetJmlByShiptoCustomerCode(TipeInvoice, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByCustomerShiptoName Then
                Return DTA.ScalarQuery_GetJmlByShiptoCustomerName(TipeInvoice, Criteria + "%")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetSONumber_ForInvoiceDP(ByVal vPrimaryInvoiceDP As Double) As String
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice1TableAdapter
        Dim RetVal As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetSODPNumber(vPrimaryInvoiceDP)

            If String.IsNullOrEmpty(RetVal) = True Then
                Return ""
            Else
                Return RetVal
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetCustomerPONumber_ForInvoiceDP(ByVal vPrimaryInvoiceDP As Double) As String
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice1TableAdapter
        Dim RetVal As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetCustomerPONumber_SO_DP(vPrimaryInvoiceDP)

            If String.IsNullOrEmpty(RetVal) = True Then
                Return ""
            Else
                Return RetVal
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTimestamp(ByVal vPrimary As Double) As Date
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice1TableAdapter

        Try
            If DTA.ScalarQuery_GetTimestamp(vPrimary).HasValue = True Then
                Return DTA.ScalarQuery_GetTimestamp(vPrimary).Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function


    Private Sub Hapus_Jurnal(ByVal vPrimaryInvoice As Double)
        Dim DTA_Balance As New DseTransaksiSalesInvoice_TransfJurnalTableAdapters.tbltrans_balance_salesTableAdapter
        Dim DTA_Jurnal As New DseTransaksiSalesInvoice_TransfJurnalTableAdapters.tbltrans_jurnalTableAdapter
        Dim TipeBalance As EnumTipeSalesBalance = EnumTipeSalesBalance.Automatic
        Dim TipeJurnal As EnumTipeJournal = EnumTipeJournal.SalesJournal
        Dim Obj_KartuAR As New ClsTransaksi_Kartu_APAR

        Try
            'hapuskan data jurnal lama
            DTA_Jurnal.Delete(TipeJurnal, vPrimaryInvoice)
            'hapuskan data balance sales lama
            DTA_Balance.DeleteQuery(vPrimaryInvoice, TipeBalance)
            'hapuskan kartu AR lama
            Obj_KartuAR.Hapus(vPrimaryInvoice, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Otomatis)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_Balance.Dispose()
            DTA_Jurnal.Dispose()
            Obj_KartuAR = Nothing
        End Try
    End Sub

    Private Sub Hapus_Jurnal_DP(ByVal vPrimaryInvoice As Double)
        Dim DTA_Balance As New DseTransaksiSalesInvoice_TransfJurnalTableAdapters.tbltrans_balance_salesTableAdapter
        Dim DTA_Jurnal As New DseTransaksiSalesInvoice_TransfJurnalTableAdapters.tbltrans_jurnalTableAdapter
        Dim TipeBalance As EnumTipeSalesBalance = EnumTipeSalesBalance.Automatic
        Dim TipeJurnal As EnumTipeJournal = EnumTipeJournal.SalesDownPayment
        Dim Obj_KartuAR As New ClsTransaksi_Kartu_APAR

        Try
            'hapuskan data jurnal lama
            DTA_Jurnal.Delete(TipeJurnal, vPrimaryInvoice)
            'hapuskan data balance sales lama
            DTA_Balance.DeleteQuery(vPrimaryInvoice, TipeBalance)
            'hapuskan kartu AR lama
            Obj_KartuAR.Hapus(vPrimaryInvoice, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Otomatis)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_Balance.Dispose()
            DTA_Jurnal.Dispose()
            Obj_KartuAR = Nothing
        End Try
    End Sub

    Private Sub TransferJurnal_Invoice2(ByVal vPrimaryInvoice As Double, ByVal vPrimaryCustomer As Double, ByVal vSales As Decimal, ByVal vVAT As Decimal, _
       ByVal vDutystamp As Decimal, ByVal vUrutanInvoice As Double, ByVal vTanggalInvoice As Date, ByVal vInvoiceNumber As String, _
       ByVal vUserCode As String, ByVal vPaymentTerm As Double, ByVal vExchangeRate As Decimal, ByVal vSalesOriginalAmount As Decimal, _
       ByVal vPrimaryCurrency As Double, ByVal vOngkosKirimIDR As Decimal, ByVal vDP_Original As Decimal, ByVal vDPBaseCurrency As Decimal, _
       ByVal vOngkosKirimOriginal As Decimal, ByVal vGlobalDiscOriginal As Decimal, ByVal vGlobalDiscIDR As Decimal, _
       ByVal vPrimaryCurrTax As Double, ByVal vPrimaryAccount As Double, ByVal vDutyStampOriginal As Decimal, _
       ByVal vVATOriginal As Decimal, ByVal vPrimaryAccountDP As Double, ByVal vOtherCost As Decimal, ByVal vOtherCostOriginal As Decimal, _
       ByVal vtblItem As DseTransaksiSalesInvoice.tbltrans_salesinvoice_item_costDataTable, _
       ByVal vrowItem As DseTransaksiSalesInvoice.tbltrans_salesinvoice_item_costRow)
        Dim DTA_Balance As New DseTransaksiSalesInvoice_TransfJurnalTableAdapters.tbltrans_balance_salesTableAdapter
        Dim TipeBalance As EnumTipeSalesBalance = EnumTipeSalesBalance.Automatic
        Dim DTA_Jurnal As New DseTransaksiSalesInvoice_TransfJurnalTableAdapters.tbltrans_jurnalTableAdapter
        Dim DTA_JurnalItem As New DseTransaksiSalesInvoice_TransfJurnalTableAdapters.tbltrans_jurnal_itemTableAdapter
        Dim TipeJurnal As EnumTipeJournal = EnumTipeJournal.SalesJournal
        Dim RefNumber As String = CStr(vUrutanInvoice)
        Dim PrimaryJurnal As Double
        Dim Debet, Credit As Decimal
        Dim ObjCust As New Core_Module.clsMasterCustomer
        Dim Obj_KartuAR As New ClsTransaksi_Kartu_APAR
        Dim MonthStr As String = ""
        Dim PrimaryAkunAR, PrimaryAkunAR_BaseCurrency, PrimaryAkunAR_BaseIDR As Double
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim TBLSales As New DseTransaksiSalesInvoice_TransfJurnal.Sales_TransfDataTable
        Dim rowSales As DseTransaksiSalesInvoice_TransfJurnal.Sales_TransfRow
        Dim PrimaryAkun As Double = -1
        Dim GabungDgnSales As Boolean = False
        Dim SelisihDP As Decimal = 0
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim DefPrecBase As Integer = 0
        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjMulti As New ClsMultiAccount
        Dim Dummy As String = ""
        Dim ObjJurnal As New ClsTransaksiJurnal
        Dim DPIDR2 As Decimal = 0

        Try
            DTA_Jurnal.Connection.ConnectionString = Core_Module.GetConstring

            DefPrecBase = ObjCur.GetDecimal_DefaultCurrency
            'hapuskan data jurnal lama
            DTA_Jurnal.Delete(TipeJurnal, vPrimaryInvoice)
            'hapuskan data balance sales lama
            DTA_Balance.DeleteQuery(vPrimaryInvoice, TipeBalance)
            'hapus kartu lama
            Obj_KartuAR.Hapus(vPrimaryInvoice, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Otomatis)

            If Me.ApakahTransferOtomatis(vTanggalInvoice) = True Then
                'buat jurnal baru
                MonthStr = vTanggalInvoice.Month.ToString
                While MonthStr.Length < 2
                    MonthStr = "0" + MonthStr
                End While
                If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Invoice) = True Then
                    If RefNumber.Length < 6 Then
                        While RefNumber.Length < 6
                            RefNumber = "0" + RefNumber
                        End While
                    End If
                    RefNumber = RefNumber + "/" + MonthStr + "/" + vTanggalInvoice.Year.ToString
                Else
                    RefNumber = vInvoiceNumber
                End If

                'data header
                While ObjJurnal.ApakahSudahAda("SJ-" + RefNumber + Dummy, DTA_Jurnal) = True
                    If Dummy = "" Then
                        Dummy = "-1"
                    Else
                        Dummy = CInt(Dummy) - 1
                    End If
                End While
                DTA_Jurnal.Insert("SJ-" + RefNumber + Dummy, vTanggalInvoice, _
                vUrutanInvoice, TipeJurnal, vPrimaryInvoice, PrimaryCustomer, "Invoice: " + _
                vInvoiceNumber + " - " + ObjCust.GetCustomerCode(vPrimaryCustomer) + " - " + ObjCust.GetCustName(vPrimaryCustomer), vUserCode, _
                Now, vUserCode, Now, Now)

                'data item
                PrimaryJurnal = DTA_Jurnal.ScalarQuery_GetPrimaryJurnal(TipeJurnal, vPrimaryInvoice).Value


                'simpan data ar, debet
                'PrimaryAkunAR = Me.GetPrimaryAkunAR(vPrimaryInvoice)
                'PrimaryAkunAR_BaseCurrency = ObjCur.GetPrimaryAkunAR(Core_Module.PrimaryBaseCurrency)
                PrimaryAkunAR = vPrimaryAccount
                PrimaryAkunAR_BaseCurrency = ObjMulti.GetPrimaryAkunARBase(vPrimaryCurrency)
                PrimaryAkunAR_BaseIDR = ObjCur.GetPrimaryAkunAR(Core_Module.PrimaryBaseCurrency)

                Debet = 0
                Credit = 0

                If vPrimaryCurrency = Core_Module.PrimaryBaseCurrency Then
                    '========================================================================================================================================
                    'JIKA TRANSAKSI IDR
                    '========================================================================================================================================
                    Debet = vSales + vVAT + vDutystamp + vOngkosKirimIDR - vGlobalDiscIDR + vOtherCost
                    'masukkan data cost yang ap 
                    DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkunAR, Debet, Credit, "", _
                    vSalesOriginalAmount + vVATOriginal + vDutyStampOriginal + vOngkosKirimOriginal - vGlobalDiscOriginal + vOtherCostOriginal, Credit, vExchangeRate)
                Else
                    '========================================================================================================================================
                    'JIKA TRANSAKSI NON IDR
                    '========================================================================================================================================
                    '----------------------------------------------------------------------------------------------------------------------------------
                    'JIKA APACCOUNT yg dipilih = base -------------------------------------------------------------------------------------------------
                    '----------------------------------------------------------------------------------------------------------------------------------
                    If vPrimaryCurrTax = vPrimaryCurrency Then
                        '----------------------------------------------------------------------------------------------------------------------------------
                        'JIKA Tax NON IDR ---------------------------------------------------------------------------------------------------------------
                        '----------------------------------------------------------------------------------------------------------------------------------
                        Debet = vSales + vVAT + vDutystamp + vOngkosKirimIDR - vGlobalDiscIDR + vOtherCost
                        DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkunAR, Debet, Credit, "", _
                       vSalesOriginalAmount + vVATOriginal + vDutyStampOriginal + vOngkosKirimOriginal - vGlobalDiscOriginal + vOtherCostOriginal, Credit, vExchangeRate)
                    Else
                        '----------------------------------------------------------------------------------------------------------------------------------
                        'JIKA Tax IDR ------------------------------------------------------------------------------------------------------------------
                        '----------------------------------------------------------------------------------------------------------------------------------
                        Debet = vSales + vOngkosKirimIDR - vGlobalDiscIDR + vOtherCost
                        DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkunAR, Debet, Credit, "", _
                       vSalesOriginalAmount + vOngkosKirimOriginal - vGlobalDiscOriginal + vOtherCostOriginal, Credit, vExchangeRate)


                        'masukkan data cost yang ap  ke ap base
                        If vVATOriginal + vDutyStampOriginal <> 0 Then
                            Debet = vVATOriginal + vDutyStampOriginal
                            DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkunAR_BaseIDR, Debet, Credit, "", Debet, Credit, 1)
                        End If

                    End If
                End If


                'Debet = 0
                'Credit = 0
                'If PrimaryAkunAR = PrimaryAkunAR_BaseCurrency Then
                '    'idr
                '    Debet = vSales + vVAT + vDutystamp + vOngkosKirimOriginal - vGlobalDiscOriginal
                '    DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkunAR, Debet, Credit, "", Debet, Credit, 1)
                'Else
                '    'non idr
                '    Debet = vSales + vOngkosKirimIDR - vGlobalDiscIDR
                '    DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkunAR, Debet, Credit, "", _
                '    vSalesOriginalAmount + vOngkosKirimOriginal - vGlobalDiscOriginal, 0, vExchangeRate)
                '    Debet = vVAT + vDutystamp
                '    If Debet <> 0 Then
                '        DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkunAR_BaseCurrency, Debet, Credit, "", Debet, Credit, 1)
                '    End If
                'End If

                'simpan data biaya diskon
                If vGlobalDiscOriginal <> 0 Then
                    PrimaryAkun = Me.GetPrimaryAkunGlobalDiscInv
                    If PrimaryAkun <> -1 Then
                        DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkun, vGlobalDiscIDR, 0, "", vGlobalDiscIDR, 0, 1)
                    End If
                End If

                'simpan data sales
                Me.GetSalesTransfer(TBLSales, vPrimaryInvoice)

                Dim decSales As Decimal = 0
                decSales = vSales
                For intA As Integer = 0 To TBLSales.Rows.Count - 1
                    rowSales = TBLSales.Rows(intA)
                    If intA <> TBLSales.Rows.Count - 1 Then
                        decSales -= rowSales.Amount
                    Else
                        rowSales.Amount = decSales
                    End If
                Next

                'simpan data pendapatan ongkos kirim
                If vOngkosKirimIDR <> 0 Then
                    PrimaryAkun = Me.GetPrimaryAkunPendapatanOngkosKirim
                    For Each rowSales In TBLSales.Rows
                        If rowSales.Primary_Akun = PrimaryAkun Then
                            rowSales.Amount += vOngkosKirimIDR
                            rowSales.Amount_Original += vOngkosKirimIDR
                            GabungDgnSales = True
                            Exit For
                        End If
                    Next
                End If


                For Each rowSales In TBLSales.Rows
                    decSales -= rowSales.Amount
                    DTA_JurnalItem.InsertQuery(PrimaryJurnal, rowSales.Primary_Akun, 0, rowSales.Amount, "", 0, rowSales.Amount, 1)
                Next

                'simpan data pendapatan ongkos kirim
                If GabungDgnSales = False Then
                    If vOngkosKirimIDR <> 0 Then
                        PrimaryAkun = Me.GetPrimaryAkunPendapatanOngkosKirim
                        If PrimaryAkun <> -1 Then
                            DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkun, 0, vOngkosKirimIDR, "", 0, vOngkosKirimIDR, 1)
                        End If
                    End If
                End If


                'simpan data vat in
                If vVAT <> 0 Then
                    PrimaryAkun = Me.GetPrimaryAkunVATOut
                    If PrimaryAkun <> -1 Then
                        If vPrimaryCurrTax = vPrimaryCurrency Then
                            DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkun, 0, vVAT, "", 0, vVAT, 1)
                        Else
                            DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkun, 0, vVATOriginal, "", 0, vVATOriginal, 1)
                        End If
                    End If
                End If

                'simpan data duty stamp
                If vDutystamp <> 0 Then
                    PrimaryAkun = Me.GetPrimaryDutyStamp
                    If PrimaryAkun <> -1 Then
                        If vPrimaryCurrTax = vPrimaryCurrency Then
                            DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkun, 0, vDutystamp, "", 0, vDutystamp, 1)
                        Else
                            DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkun, 0, vDutyStampOriginal, "", 0, vDutyStampOriginal, 1)
                        End If
                    End If
                End If


                'simpan data Other cost
                If vOtherCost <> 0 Then
                    If vtblItem.Rows.Count > 0 Then
                        Dim ListPrimaryAccount As New List(Of Double)
                        For Each vrowItem In vtblItem.Rows
                            With vrowItem
                                If ListPrimaryAccount.Contains(.Primary_Akun) = False Then
                                    ListPrimaryAccount.Add(.Primary_Akun)
                                End If
                            End With
                        Next

                        For i As Integer = 0 To ListPrimaryAccount.Count - 1
                            Dim Cost As Decimal = 0
                            Dim Cost_Original As Decimal = 0
                            vtblItem.DefaultView.RowFilter = Nothing
                            vtblItem.DefaultView.RowFilter = " [Primary Akun] = " & ListPrimaryAccount.Item(i)
                            If vtblItem.DefaultView.Count > 0 Then
                                Dim dv As Data.DataRowView
                                For Each dv In vtblItem.DefaultView
                                    vrowItem = dv.Row
                                    With vrowItem
                                        Cost_Original += CDec(.Cost)
                                        Cost += CDec(.Cost) * vExchangeRate
                                    End With
                                Next
                                DTA_JurnalItem.InsertQuery(PrimaryJurnal, ListPrimaryAccount.Item(i), 0, Cost, "", 0, Cost, 1)
                            End If
                        Next
                    End If
                End If



                ''simpan data vat out
                'If vVAT <> 0 Then
                '    PrimaryAkun = Me.GetPrimaryAkunVATOut
                '    If PrimaryAkun <> -1 Then
                '        DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkun, 0, vVAT, "", 0, vVAT, 1)
                '    End If
                'End If

                ''simpan data duty stamp
                'If vDutystamp <> 0 Then
                '    PrimaryAkun = Me.GetPrimaryDutyStamp
                '    If PrimaryAkun <> -1 Then
                '        DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkun, 0, vDutystamp, "", 0, vDutystamp, 1)
                '    End If
                'End If



                If vPrimaryCurrency = Core_Module.PrimaryBaseCurrency Then
                    'kurangkan data hutang dengan DP
                    If vDP_Original <> 0 Then
                        Debet = 0
                        Credit = 0
                        Debet = vDP_Original
                        DTA_JurnalItem.InsertQuery(PrimaryJurnal, vPrimaryAccountDP, Debet, Credit, "", Debet, Credit, 1)

                        Debet = 0
                        Credit = 0
                        Credit = vDP_Original
                        DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkunAR, Debet, Credit, "", Debet, Credit, 1)
                    End If
                Else
                    'kurangkan data hutang dengan DP
                    If vDP_Original <> 0 Then
                        Debet = 0
                        Credit = 0
                        Debet = vDPBaseCurrency
                        DTA_JurnalItem.InsertQuery(PrimaryJurnal, vPrimaryAccountDP, Debet, Credit, "", Debet, Credit, 1)

                        Debet = 0
                        Credit = 0
                        Credit = ObjCD.DigitGrouping(vDP_Original * vExchangeRate, DefPrecBase)
                        DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkunAR, Debet, Credit, _
                         "", Debet, vDP_Original, vExchangeRate)

                        SelisihDP = vDPBaseCurrency - ObjCD.DigitGrouping(vDP_Original * vExchangeRate, DefPrecBase)

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




                Dim AR_PadaCost As Decimal = 0
                Dim AR_PadaCost_Original As Decimal = 0
                Dim AR_PadaCost_Base As Decimal = 0

                Credit = 0
                Debet = 0
                If vPrimaryCurrency = Core_Module.PrimaryBaseCurrency Then
                    '========================================================================================================================================
                    'JIKA TRANSAKSI IDR
                    '========================================================================================================================================
                    DTA_Balance.Insert(vPrimaryCustomer, vPrimaryInvoice, TipeBalance, "", #1/1/1950#, "", _
                    vTanggalInvoice + TimeSpan.FromDays(vPaymentTerm), _
                    vSales + vOngkosKirimIDR + vOtherCost - vDPBaseCurrency - vGlobalDiscIDR, vVAT, vDutystamp, _
                    vSales + vOngkosKirimIDR + vVAT + vDutystamp + vOtherCost - vDPBaseCurrency - vGlobalDiscIDR, 0, _
                    vUserCode, Now, vUserCode, Now, Now, vPrimaryCurrency, _
                    vSalesOriginalAmount + vVATOriginal + vDutyStampOriginal + vOngkosKirimOriginal _
                    + vOtherCostOriginal - vDP_Original - vGlobalDiscOriginal, 1, PrimaryAkunAR)

                    'buat kartu ar baru
                    Obj_KartuAR.Simpan1(vPrimaryInvoice, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Otomatis, vInvoiceNumber, _
                    vPrimaryCustomer, vSales + vVAT + vDutystamp + vOngkosKirimIDR + vOtherCost - vDPBaseCurrency - vGlobalDiscIDR, 0, "-", ClsTransaksi_Kartu_APAR.EnumJenisStockCard.AR, _
                    vTanggalInvoice, vPrimaryInvoice, vTanggalInvoice, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Otomatis, _
                    vPrimaryCurrency, vSalesOriginalAmount + vVATOriginal + vDutyStampOriginal + _
                    vOngkosKirimOriginal + vOtherCostOriginal - vDP_Original - vGlobalDiscOriginal, 0, 1, PrimaryAkunAR)
                Else
                    '========================================================================================================================================
                    'JIKA TRANSAKSI NON IDR
                    '========================================================================================================================================
                    '20190113: iwan
                    'DP IDR yang dihitung disini, dihitung menggunakan exc. rate ketika invoice
                    'ketika jurnal, menggunakan exc. rate ketika dp, karena dikompensasi oleh exchange gain/loss
                    'nilai IDR yang dimasukkan ke balance sales/kartu ar  = SelisihDP + (vSales - vDPBaseCurrency) atau = vSales - ObjCD.DigitGrouping(vDP_Original * vExchangeRate, DefPrecBase)
                    DPIDR2 = ObjCD.DigitGrouping(vDP_Original * vExchangeRate, DefPrecBase)

                    If vPrimaryCurrTax = vPrimaryCurrency Then
                        '----------------------------------------------------------------------------------------------------------------------------------
                        'JIKA Tax NON IDR ---------------------------------------------------------------------------------------------------------------
                        '----------------------------------------------------------------------------------------------------------------------------------
                        'balance purchase amount
                        DTA_Balance.Insert(vPrimaryCustomer, vPrimaryInvoice, TipeBalance, "", #1/1/1950#, "", _
                                          vTanggalInvoice + TimeSpan.FromDays(vPaymentTerm), _
                                          vSales + vOngkosKirimIDR + vOtherCost - DPIDR2 - vGlobalDiscIDR, vVAT, vDutystamp, _
                                          vSales + vVAT + vDutystamp + vOngkosKirimIDR + vOtherCost - DPIDR2 - vGlobalDiscIDR, 0, _
                                          vUserCode, Now, vUserCode, Now, Now, vPrimaryCurrency, _
                                          vSalesOriginalAmount + vVATOriginal + vDutyStampOriginal + vOngkosKirimOriginal + vOtherCostOriginal - _
                                          vDP_Original - vGlobalDiscOriginal, vExchangeRate, PrimaryAkunAR)

                        'buat kartu ar baru
                        Obj_KartuAR.Simpan1(vPrimaryInvoice, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Otomatis, vInvoiceNumber, _
                        vPrimaryCustomer, vSales + vVAT + vDutystamp + vOngkosKirimIDR + vOtherCost - DPIDR2 - vGlobalDiscIDR, 0, _
                        "-", ClsTransaksi_Kartu_APAR.EnumJenisStockCard.AR, vTanggalInvoice, vPrimaryInvoice, vTanggalInvoice, vPaymentTerm, _
                        ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Otomatis, vPrimaryCurrency, _
                        vSalesOriginalAmount + vVATOriginal + vDutyStampOriginal + vOngkosKirimOriginal + vOtherCostOriginal - vDP_Original - _
                        vGlobalDiscOriginal, 0, vExchangeRate, PrimaryAkunAR)
                    Else
                        '----------------------------------------------------------------------------------------------------------------------------------
                        'JIKA Tax IDR -------------------------------------------------------------------------------------------------------------------
                        '----------------------------------------------------------------------------------------------------------------------------------
                        'balance purchase amount
                        'DTA_Balance.Insert(vPrimaryCustomer, vPrimaryInvoice, TipeBalance, "", #1/1/1950#, "", _
                        '                  vTanggalInvoice + TimeSpan.FromDays(vPaymentTerm), _
                        '                  vSales + vOngkosKirimIDR + vOtherCost - vDPBaseCurrency - vGlobalDiscIDR, 0, 0, _
                        '                  vSales + vOngkosKirimIDR + vOtherCost - vDPBaseCurrency - vGlobalDiscIDR, 0, _
                        '                  vUserCode, Now, vUserCode, Now, Now, vPrimaryCurrency, _
                        '                  vSalesOriginalAmount + vOngkosKirimOriginal + vOtherCostOriginal - vDP_Original _
                        '                  - vGlobalDiscOriginal, vExchangeRate, PrimaryAkunAR)


                        ''buat kartu ar purchase amount
                        'Obj_KartuAR.Simpan(vPrimaryInvoice, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Otomatis, vInvoiceNumber, _
                        '                  vPrimaryCustomer, vSales + vOngkosKirimIDR + vOtherCost - vDPBaseCurrency - vGlobalDiscIDR, 0, "-", _
                        '                  ClsTransaksi_Kartu_APAR.EnumJenisStockCard.AR, vTanggalInvoice, vPrimaryInvoice, vTanggalInvoice, _
                        '                  vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Otomatis, vPrimaryCurrency, _
                        '                  vSalesOriginalAmount + vOngkosKirimOriginal + vOtherCostOriginal - vDP_Original - vGlobalDiscOriginal, 0, vExchangeRate, PrimaryAkunAR)
                        DTA_Balance.Insert(vPrimaryCustomer, vPrimaryInvoice, TipeBalance, "", #1/1/1950#, "", _
                                          vTanggalInvoice + TimeSpan.FromDays(vPaymentTerm), _
                                          vSales + vOngkosKirimIDR + vOtherCost - DPIDR2 - vGlobalDiscIDR, 0, 0, _
                                          vSales + vOngkosKirimIDR + vOtherCost - DPIDR2 - vGlobalDiscIDR, 0, _
                                          vUserCode, Now, vUserCode, Now, Now, vPrimaryCurrency, _
                                          vSalesOriginalAmount + vOngkosKirimOriginal + vOtherCostOriginal - vDP_Original _
                                          - vGlobalDiscOriginal, vExchangeRate, PrimaryAkunAR)


                        'buat kartu ar purchase amount
                        Obj_KartuAR.Simpan1(vPrimaryInvoice, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Otomatis, vInvoiceNumber, _
                                          vPrimaryCustomer, vSales + vOngkosKirimIDR + vOtherCost - DPIDR2 - vGlobalDiscIDR, 0, "-", _
                                          ClsTransaksi_Kartu_APAR.EnumJenisStockCard.AR, vTanggalInvoice, vPrimaryInvoice, vTanggalInvoice, _
                                          vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Otomatis, vPrimaryCurrency, _
                                          vSalesOriginalAmount + vOngkosKirimOriginal + vOtherCostOriginal - vDP_Original - vGlobalDiscOriginal, 0, vExchangeRate, PrimaryAkunAR)

                        If vVATOriginal + vDutyStampOriginal > 0 Then
                            'balance purchase tax
                            DTA_Balance.Insert(vPrimaryCustomer, vPrimaryInvoice, TipeBalance, "", #1/1/1950#, "", _
                                               vTanggalInvoice + TimeSpan.FromDays(vPaymentTerm), _
                                               0, vVATOriginal, vDutyStampOriginal, _
                                                vVATOriginal + vDutyStampOriginal, 0, _
                                               vUserCode, Now, vUserCode, Now, Now, Core_Module.PrimaryBaseCurrency, _
                                               vVATOriginal + vDutyStampOriginal, 1, PrimaryAkunAR_BaseIDR)



                            'buat kartu ar purchase tax
                            Obj_KartuAR.Simpan1(vPrimaryInvoice, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Otomatis, vInvoiceNumber, _
                            vPrimaryCustomer, vVATOriginal + vDutyStampOriginal, 0, "-", ClsTransaksi_Kartu_APAR.EnumJenisStockCard.AR, _
                            vTanggalInvoice, vPrimaryInvoice, vTanggalInvoice, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Otomatis, _
                            Core_Module.PrimaryBaseCurrency, vVATOriginal + vDutyStampOriginal, 0, 1, PrimaryAkunAR_BaseIDR)
                        End If
                    End If
                End If
            End If


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_Balance.Dispose()
            DTA_Jurnal.Dispose()
            DTA_JurnalItem.Dispose()
            ObjCust = Nothing
            Obj_KartuAR = Nothing
            ObjCur = Nothing
            TBLSales = Nothing
            ObjSetting = Nothing
            ObjCD = Nothing
            ObjJurnal = Nothing
        End Try
    End Sub

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

    Public Function GetPrimaryAkunPendapatanOngkosKirim() As Double
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim Dummy As String = ""

        Try
            Dummy = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_ShippingIncome)
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

    Private Function GetPrimaryAkunVATOut() As Double
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim Dummy As String = ""

        Try
            Dummy = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_VAT_Out)
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

    Private Function GetPrimaryAkunGlobalDiscInv() As Double
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim Dummy As String = ""

        Try
            Dummy = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.GlobalDiscAcc)
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

    Private Function GetPrimaryAkunAR(ByVal vPrimaryInv As Double) As Double
        Dim ObjCur As New Core_Module.clsMasterCurrency

        Try
            Return ObjCur.GetPrimaryAkunAR(Me.Get_Primary_Currency(vPrimaryInv))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCur = Nothing
        End Try
    End Function

    Private Sub GetSalesTransfer(ByVal vTBLAkun As DseTransaksiSalesInvoice_TransfJurnal.Sales_TransfDataTable, ByVal vPrimaryInv As Double)
        Dim DTA As New DseTransaksiSalesInvoice_TransfJurnalTableAdapters.DataTable1TableAdapter
        Dim TBL As New DseTransaksiSalesInvoice_TransfJurnal.DataTable1DataTable
        Dim row As DseTransaksiSalesInvoice_TransfJurnal.DataTable1Row
        Dim rowAkun As DseTransaksiSalesInvoice_TransfJurnal.Sales_TransfRow

        Dim Filter As String = ""

        Try
            'isikan item
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, vPrimaryInv)

            'kosongkan tabel
            vTBLAkun.Rows.Clear()

            'masukkan ke dalam tabel akun
            For Each row In TBL.Rows
                Filter = "[Primary Akun] = " + row.Primary_Sales_Account.ToString
                vTBLAkun.DefaultView.RowFilter = Filter

                If vTBLAkun.DefaultView.Count > 0 Then
                    'tambahkan amount saja
                    rowAkun = vTBLAkun.DefaultView(0).Row
                    rowAkun.Amount += row.Sales_Amount
                    rowAkun.Amount_Original += row.Amount
                Else
                    'buat baris baru
                    rowAkun = vTBLAkun.NewSales_TransfRow
                    rowAkun.Amount = row.Sales_Amount
                    rowAkun.Amount_Original = row.Amount
                    rowAkun.Primary_Akun = row.Primary_Sales_Account
                    vTBLAkun.Rows.Add(rowAkun)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Private Sub TransferJurnal_Invoice_DP2(ByVal vPrimaryInvoice As Double, ByVal vPrimaryCustomer As Double, ByVal vSales As Decimal, ByVal vVAT As Decimal, _
     ByVal vDutystamp As Decimal, ByVal vUrutanInvoice As Double, ByVal vTanggalInvoice As Date, ByVal vInvoiceNumber As String, _
     ByVal vUserCode As String, ByVal vPaymentTerm As Double, ByVal vPrimaryCurrency As Double, ByVal vSalesOriginalAmount As Decimal, _
     ByVal vExchangeRate As Decimal, ByVal vPrimaryCurrTax As Double, ByVal vPrimaryAccount As Double, ByVal vPrimaryAccountDP As Double, _
     ByVal vVatOriginal As Decimal, ByVal vDutyStampOriginal As Decimal)
        Dim DTA_Balance As New DseTransaksiSalesInvoice_TransfJurnalTableAdapters.tbltrans_balance_salesTableAdapter
        Dim TipeBalance As EnumTipeSalesBalance = EnumTipeSalesBalance.Automatic
        Dim DTA_Jurnal As New DseTransaksiSalesInvoice_TransfJurnalTableAdapters.tbltrans_jurnalTableAdapter
        Dim DTA_JurnalItem As New DseTransaksiSalesInvoice_TransfJurnalTableAdapters.tbltrans_jurnal_itemTableAdapter
        Dim TipeJurnal As EnumTipeJournal = EnumTipeJournal.SalesDownPayment
        Dim RefNumber As String = CStr(vUrutanInvoice)
        Dim PrimaryJurnal As Double
        Dim Debet, Credit As Double
        Dim ObjCust As New Core_Module.clsMasterCustomer
        Dim Obj_KartuAR As New ClsTransaksi_Kartu_APAR
        Dim MonthStr As String = ""
        Dim PrimaryAccount As Double = -1
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim PrimaryAkunAR, PrimaryAkunAR_BaseCurrency, PrimaryAkunAR_BaseIDR As Double
        Dim ObjMulti As New ClsMultiAccount
        Dim Dummy As String = ""
        Dim ObjJurnal As New ClsTransaksiJurnal

        Try
            DTA_Jurnal.Connection.ConnectionString = Core_Module.GetConstring

            'hapuskan data jurnal lama
            DTA_Jurnal.Delete(TipeJurnal, vPrimaryInvoice)
            'hapuskan data balance sales lama
            DTA_Balance.DeleteQuery(vPrimaryInvoice, TipeBalance)
            'hapus kartu lama
            Obj_KartuAR.Hapus(vPrimaryInvoice, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Otomatis)

            If Me.ApakahTransferOtomatis(vTanggalInvoice) = True Then
                'buat jurnal baru
                MonthStr = vTanggalInvoice.Month.ToString
                While MonthStr.Length < 2
                    MonthStr = "0" + MonthStr
                End While
                If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Invoice) = True Then
                    If RefNumber.Length < 6 Then
                        While RefNumber.Length < 6
                            RefNumber = "0" + RefNumber
                        End While
                    End If
                    RefNumber = RefNumber + "/" + MonthStr + "/" + vTanggalInvoice.Year.ToString
                Else
                    RefNumber = vInvoiceNumber
                End If

                'data header
                While ObjJurnal.ApakahSudahAda("SJDP-" + RefNumber + Dummy, DTA_Jurnal) = True
                    If Dummy = "" Then
                        Dummy = "-1"
                    Else
                        Dummy = CInt(Dummy) - 1
                    End If
                End While

                DTA_Jurnal.Insert("SJDP-" + RefNumber + Dummy, vTanggalInvoice, _
                vUrutanInvoice, TipeJurnal, vPrimaryInvoice, PrimaryCustomer, "Invoice: " + vInvoiceNumber _
                + " - " + ObjCust.GetCustomerCode(vPrimaryCustomer) + " - " + ObjCust.GetCustName(vPrimaryCustomer), vUserCode, _
                Now, vUserCode, Now, Now)

                PrimaryJurnal = DTA_Jurnal.ScalarQuery_GetPrimaryJurnal(TipeJurnal, vPrimaryInvoice).Value

                PrimaryAkunAR = vPrimaryAccount
                PrimaryAkunAR_BaseCurrency = ObjMulti.GetPrimaryAkunARBase(vPrimaryCurrency)
                PrimaryAkunAR_BaseIDR = ObjCur.GetPrimaryAkunAR(Core_Module.PrimaryBaseCurrency)



                If vPrimaryCurrency = Core_Module.PrimaryBaseCurrency Then
                    '========================================================================================================================================
                    'JIKA TRANSAKSI IDR
                    '========================================================================================================================================
                    '----------------------------------------------------------------------------------------------------------------------------------
                    'JIKA APACCOUNT yg dipilih = base -------------------------------------------------------------------------------------------------
                    '----------------------------------------------------------------------------------------------------------------------------------
                    Debet = vSales + vVAT + vDutystamp
                    'masukkan data cost yang ap 
                    DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkunAR, Debet, Credit, "", _
                    vSalesOriginalAmount + vVatOriginal + vDutyStampOriginal, Credit, vExchangeRate)
                Else
                    '========================================================================================================================================
                    'JIKA TRANSAKSI NON IDR
                    '========================================================================================================================================
                    If vPrimaryCurrTax = vPrimaryCurrency Then
                        '----------------------------------------------------------------------------------------------------------------------------------
                        'JIKA Tax NON IDR ---------------------------------------------------------------------------------------------------------------
                        '----------------------------------------------------------------------------------------------------------------------------------
                        Debet = vSales + vVAT + vDutystamp
                        DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkunAR, Debet, Credit, "", _
                       vSalesOriginalAmount + vVatOriginal + vDutyStampOriginal, Credit, vExchangeRate)
                    Else
                        '----------------------------------------------------------------------------------------------------------------------------------
                        'JIKA Tax IDR ------------------------------------------------------------------------------------------------------------------
                        '----------------------------------------------------------------------------------------------------------------------------------
                        Debet = vSales
                        DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkunAR, Debet, Credit, "", _
                       vSalesOriginalAmount, Credit, vExchangeRate)


                        'masukkan data cost yang ap  ke ap base
                        If vVatOriginal + vDutyStampOriginal <> 0 Then
                            Debet = vVatOriginal + vDutyStampOriginal
                            DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAkunAR_BaseIDR, Debet, Credit, "", Debet, Credit, 1)
                        End If
                    End If
                End If



                'biaya dibayar dimuka
                If vPrimaryAccountDP > -1 Then
                    PrimaryAccount = vPrimaryAccountDP
                Else
                    PrimaryAccount = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_Invoice_DP)
                End If

                Debet = 0
                Credit = vSales
                DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAccount, Debet, Credit, "", Debet, Credit, 1)


                'simpan data vat in-debet
                If vVAT <> 0 Then
                    PrimaryAccount = Me.GetPrimaryAkunVATOut
                    If PrimaryAccount <> -1 Then
                        If vPrimaryCurrency = vPrimaryCurrTax Then
                            DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAccount, 0, vVAT, "", 0, vVAT, 1)
                        Else
                            DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAccount, 0, vVatOriginal, "", 0, vVatOriginal, 1)
                        End If
                    End If
                End If

                'simpan data duty stamp-debet
                If vDutystamp <> 0 Then
                    PrimaryAccount = Me.GetPrimaryDutyStamp
                    If PrimaryAccount <> -1 Then
                        If vPrimaryCurrTax = vPrimaryCurrency Then
                            DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAccount, 0, vDutystamp, "", 0, vDutystamp, 1)
                        Else
                            DTA_JurnalItem.InsertQuery(PrimaryJurnal, PrimaryAccount, 0, vDutyStampOriginal, "", 0, vDutyStampOriginal, 1)
                        End If
                    End If
                End If


                Dim AR_PadaCost As Decimal = 0
                Dim AR_PadaCost_Original As Decimal = 0
                Dim AR_PadaCost_Base As Decimal = 0

                If vPrimaryCurrency = Core_Module.PrimaryBaseCurrency Then
                    '========================================================================================================================================
                    'JIKA TRANSAKSI IDR
                    '========================================================================================================================================
                    DTA_Balance.Insert(vPrimaryCustomer, vPrimaryInvoice, TipeBalance, "", #1/1/1950#, "", _
                    vTanggalInvoice + TimeSpan.FromDays(vPaymentTerm), _
                    vSales, vVAT, vDutystamp, _
                    vSales + vVAT + vDutystamp, 0, _
                    vUserCode, Now, vUserCode, Now, Now, vPrimaryCurrency, _
                    vSalesOriginalAmount + vVatOriginal + vDutyStampOriginal, 1, PrimaryAkunAR)

                    'buat kartu ar baru
                    Obj_KartuAR.Simpan1(vPrimaryInvoice, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Otomatis, vInvoiceNumber, _
                    vPrimaryCustomer, vSales + vVAT + vDutystamp, 0, "-", ClsTransaksi_Kartu_APAR.EnumJenisStockCard.AR, _
                    vTanggalInvoice, vPrimaryInvoice, vTanggalInvoice, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Otomatis, _
                    vPrimaryCurrency, vSalesOriginalAmount + vVatOriginal + vDutyStampOriginal, 0, 1, PrimaryAkunAR)
                Else
                    '========================================================================================================================================
                    'JIKA TRANSAKSI NON IDR
                    '========================================================================================================================================
                    If vPrimaryCurrTax = vPrimaryCurrency Then
                        '----------------------------------------------------------------------------------------------------------------------------------
                        'JIKA Tax NON IDR ---------------------------------------------------------------------------------------------------------------
                        '----------------------------------------------------------------------------------------------------------------------------------
                        'balance purchase amount
                        DTA_Balance.Insert(vPrimaryCustomer, vPrimaryInvoice, TipeBalance, "", #1/1/1950#, "", _
                        vTanggalInvoice + TimeSpan.FromDays(vPaymentTerm), _
                        vSales, vVAT, vDutystamp, _
                        vSales + vVAT + vDutystamp, 0, _
                        vUserCode, Now, vUserCode, Now, Now, vPrimaryCurrency, _
                        vSalesOriginalAmount + vVatOriginal + vDutyStampOriginal, vExchangeRate, PrimaryAkunAR)

                        'buat kartu ar baru
                        Obj_KartuAR.Simpan1(vPrimaryInvoice, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Otomatis, vInvoiceNumber, _
                        vPrimaryCustomer, vSales + vVAT + vDutystamp, 0, "-", ClsTransaksi_Kartu_APAR.EnumJenisStockCard.AR, _
                        vTanggalInvoice, vPrimaryInvoice, vTanggalInvoice, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Otomatis, _
                        vPrimaryCurrency, vSalesOriginalAmount + vVAT + vDutystamp, 0, vExchangeRate, PrimaryAkunAR)
                    Else
                        '----------------------------------------------------------------------------------------------------------------------------------
                        'JIKA Tax IDR -------------------------------------------------------------------------------------------------------------------
                        '----------------------------------------------------------------------------------------------------------------------------------
                        'balance purchase amount
                        DTA_Balance.Insert(vPrimaryCustomer, vPrimaryInvoice, TipeBalance, "", #1/1/1950#, "", _
                                          vTanggalInvoice + TimeSpan.FromDays(vPaymentTerm), _
                                          vSales, 0, 0, _
                                          vSales, 0, _
                                          vUserCode, Now, vUserCode, Now, Now, vPrimaryCurrency, _
                                          vSalesOriginalAmount + vVatOriginal + vDutyStampOriginal, vExchangeRate, PrimaryAkunAR)


                        'buat kartu ar purchase amount
                        Obj_KartuAR.Simpan1(vPrimaryInvoice, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Otomatis, vInvoiceNumber, _
                                          vPrimaryCustomer, vSales, 0, "-", ClsTransaksi_Kartu_APAR.EnumJenisStockCard.AR, _
                                          vTanggalInvoice, vPrimaryInvoice, vTanggalInvoice, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Otomatis, _
                                          vPrimaryCurrency, vSalesOriginalAmount + vVatOriginal + vDutyStampOriginal, 0, vExchangeRate, PrimaryAkunAR)

                        If vVatOriginal + vDutyStampOriginal <> 0 Then
                            'balance purchase tax
                            DTA_Balance.Insert(vPrimaryCustomer, vPrimaryInvoice, TipeBalance, "", #1/1/1950#, "", _
                                             vTanggalInvoice + TimeSpan.FromDays(vPaymentTerm), _
                                             0, vVatOriginal, vDutyStampOriginal, _
                                              vVatOriginal + vDutyStampOriginal, 0, _
                                             vUserCode, Now, vUserCode, Now, Now, Core_Module.PrimaryBaseCurrency, _
                                             vVatOriginal + vDutyStampOriginal, 1, PrimaryAkunAR_BaseCurrency)

                            'buat kartu ar purchase tax
                            Obj_KartuAR.Simpan1(vPrimaryInvoice, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Otomatis, vInvoiceNumber, _
                            vPrimaryCustomer, vVatOriginal + vDutyStampOriginal, 0, "-", ClsTransaksi_Kartu_APAR.EnumJenisStockCard.AR, _
                            vTanggalInvoice, vPrimaryInvoice, vTanggalInvoice, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Invoice_Otomatis, _
                            Core_Module.PrimaryBaseCurrency, vVatOriginal + vDutyStampOriginal, 0, 1, PrimaryAkunAR_BaseCurrency)
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_Balance.Dispose()
            DTA_Jurnal.Dispose()
            DTA_JurnalItem.Dispose()
            ObjCust = Nothing
            Obj_KartuAR = Nothing
            ObjSetting = Nothing
            ObjCur = Nothing
            ObjJurnal = Nothing
        End Try
    End Sub

    Public Function GetPrintStatus(ByVal vPrimary As Double) As String
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoiceTableAdapter
        Dim RetVal_Date As Nullable(Of Date)
        Dim RetVal_PrintedBy As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal_Date = DTA.ScalarQuery_GetPrintedDate(vPrimary)
            If RetVal_Date.HasValue = True Then
                If RetVal_Date <> #1/1/1950# Then
                    RetVal_PrintedBy = DTA.ScalarQuery_GetPrintedBy(vPrimary)
                    Return RetVal_PrintedBy + " - " + Format(RetVal_Date.Value, "dd-MMM-yyyy") + "-" + RetVal_Date.Value.ToShortTimeString
                Else
                    Return ""
                End If
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function
    Public Function GetNoNikPassbyInvoice(ByVal vPrimary As Double) As String
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_NewTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetNoNikPasByPrimary(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetNoNikPasByPrimary(vPrimary)
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

    Public Function GetPrimaryCurrency_ManualBalance(ByVal vPrimaryTransaksi As Double) As Double
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tblbalance_sales_dp1TableAdapter
        Dim Retval As System.Nullable(Of Double)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetPrimaryCurr(vPrimaryTransaksi)
            If Retval.HasValue Then
                Return Retval.Value
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

    Public Function GetPrimaryAPAR_Acc_ManualBalance(ByVal vPrimaryTransaksi As Double) As Double
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tblbalance_sales_dp1TableAdapter
        Dim Retval As System.Nullable(Of Double)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetPrimaryAPARAcc(vPrimaryTransaksi)
            If Retval.HasValue Then
                Return Retval.Value
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

    Public Function GetPrimary_ManualBalance(ByVal vDPNo As String, ByVal vTipe As EnumTipeDP) As Double
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tblbalance_sales_dp1TableAdapter
        Dim Retval As System.Nullable(Of Double)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetPrimaryManualDP(vDPNo, vTipe)
            If Retval.HasValue Then
                Return Retval.Value
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

    Public Function GetDPAmount(ByVal vPrimarySO As Double) As Decimal
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice1TableAdapter
        Dim RetVal As Nullable(Of Decimal)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetDPAmountByPrimarySO(vPrimarySO)
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

    Public Function GetDPAmount_ExcludePrimaryInv(ByVal vPrimarySO As Double, ByVal vPrimaryInvExclude As Double) As Decimal
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice1TableAdapter
        Dim RetVal As Nullable(Of Decimal)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetDPAmountByPrimarySO2(vPrimarySO, vPrimaryInvExclude)
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

    Public Function GePrimarySalesPerson(ByVal vPrimarySO As Double, ByVal vPrimaryDO As Double) As Double
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_itemTableAdapter
        Dim retval As Nullable(Of Double)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            retval = DTA.ScalarQuery_GetPrimarySalesPersonByPrimarySO_DO(vPrimaryDO, vPrimarySO)
            If retval.HasValue = True Then
                Return retval.Value
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

    Public Function IsiTblOutSales(ByVal vPrimaryTransaksi As Double, ByVal vTBLInvoice As DataTable, _
    ByRef vTBL As DataTable) As Boolean

    End Function

    Public Function SetDontCollect(ByVal vNoSeriFakturPajak As String, ByVal vUseTax As Boolean, ByVal vCboPPNDipungut As Integer) As Boolean
        Try
            If vUseTax = True And vCboPPNDipungut = 0 Then 'Cbo PPN Dipungun masih di set collect
                If vNoSeriFakturPajak.StartsWith("070") = True Or vNoSeriFakturPajak.StartsWith("071") = True Then
                    If MsgBox("Don't collect tax?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Return True
                    End If
                End If
            End If

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    'keperluan custom binaya
    '  Public Sub SimpanDataSalesAmount_Custom(ByVal vCalculatesBy As ClsTransaksiSalesInvoice.EnumCalculatesBy, _
    'ByVal vPrimarySalesInvoice As Double, ByVal TotalSales As Decimal)
    '      Dim Rate As Decimal = 1
    '      Dim vPrimaryTransaksiRujukan As Double = -1
    '      Dim DTA As New CustomSMESS.DseTransaksiSalesInvoiceQuick_AddByCustomProductTableAdapters.DataTable2TableAdapter
    '      Dim TBL As New CustomSMESS.DseTransaksiSalesInvoiceQuick_AddByCustomProduct.DataTable2DataTable
    '      Dim row1 As CustomSMESS.DseTransaksiSalesInvoiceQuick_AddByCustomProduct.DataTable2Row
    '      Dim SalesAmount As Decimal = 0
    '      Dim DTAReturn As New CustomSMESS.DseTransaksiSalesInvoiceQuick_AddByCustomProductTableAdapters.DataTable3TableAdapter
    '      Dim TBLReturn As New CustomSMESS.DseTransaksiSalesInvoiceQuick_AddByCustomProduct.DataTable3DataTable
    '      Dim i As Integer = 0
    '      Dim DefPrecBase As Integer = 0
    '      Dim ObjCur As New Core_Module.clsMasterCurrency

    '      Try
    '          DefPrecBase = ObjCur.GetDecimal_DefaultCurrency
    '          'isikan data item
    '          DTA.Fill(TBL, vPrimarySalesInvoice)
    '          If TBL.Rows.Count = 0 Then
    '              'kalo ga ada itemnya maka ga usah update data cost
    '              Exit Sub
    '          End If
    '          DTAReturn.Fill(TBLReturn, vPrimarySalesInvoice)

    '          'kurangkan dengan item return
    '          Me.KurangkanItemReturn2(TBLReturn, TBL)
    '          If TBL.Rows.Count = 0 Then
    '              'kalo ga ada itemnya maka ga usah update data cost
    '              Exit Sub
    '          End If

    '          'hitung exchange rate
    '          If vCalculatesBy = EnumCalculatesBy.Invoice Then
    '              row1 = TBL.Rows(0)
    '              If vCalculatesBy = ClsTransaksiSalesInvoice.EnumCalculatesBy.Delivery_Order Then
    '                  vPrimaryTransaksiRujukan = row1.Primary_DO
    '              ElseIf vCalculatesBy = ClsTransaksiSalesInvoice.EnumCalculatesBy.Invoice Then
    '                  vPrimaryTransaksiRujukan = row1.Primary_Transaksi
    '              ElseIf vCalculatesBy = ClsTransaksiSalesInvoice.EnumCalculatesBy.Sales_Order Then
    '                  vPrimaryTransaksiRujukan = row1.Primary_SO
    '              End If
    '              Rate = Me.GetExhangeRate(vPrimarySalesInvoice, vPrimaryTransaksiRujukan)
    '          End If

    '          For i = 0 To TBL.Rows.Count - 1
    '              row1 = TBL.Rows(i)
    '              'hitung exchange rate
    '              If vCalculatesBy <> EnumCalculatesBy.Invoice Then
    '                  If vCalculatesBy = ClsTransaksiSalesInvoice.EnumCalculatesBy.Delivery_Order Then
    '                      vPrimaryTransaksiRujukan = row1.Primary_DO
    '                  ElseIf vCalculatesBy = ClsTransaksiSalesInvoice.EnumCalculatesBy.Invoice Then
    '                      vPrimaryTransaksiRujukan = row1.Primary_Transaksi
    '                  ElseIf vCalculatesBy = ClsTransaksiSalesInvoice.EnumCalculatesBy.Sales_Order Then
    '                      vPrimaryTransaksiRujukan = row1.Primary_SO
    '                  End If
    '                  Rate = Me.GetExhangeRate(vPrimarySalesInvoice, vPrimaryTransaksiRujukan)
    '              End If

    '              If i <> TBL.Rows.Count - 1 Then
    '                  'untuk setiap item, hitung amount
    '                  SalesAmount = Math.Round(Math.Round(row1.Amount * Rate, _
    '                  row1.Default_Precision, MidpointRounding.AwayFromZero) * row1.Tax_Exchange_Rate, _
    '                  DefPrecBase, MidpointRounding.AwayFromZero)
    '                  'update data sales amount
    '                  DTA.UpdateQuery_SetSalesAmount(SalesAmount, row1.Primary)
    '                  TotalSales -= SalesAmount
    '              Else
    '                  SalesAmount = TotalSales
    '                  'update data sales amount
    '                  DTA.UpdateQuery_SetSalesAmount(SalesAmount, row1.Primary)
    '                  If SalesAmount <= 0 Then
    '                      MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
    '                  End If
    '              End If
    '          Next
    '      Catch ex As Exception
    '          MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '      Finally
    '          DTA.Dispose()
    '          TBL.Dispose()
    '          ObjCur = Nothing
    '      End Try
    '  End Sub

    ' Private Sub KurangkanItemReturn2(ByVal vTBLReturn As CustomSMESS.DseTransaksiSalesInvoiceQuick_AddByCustomProduct.DataTable3DataTable, _
    'ByVal vTBLInvoice As CustomSMESS.DseTransaksiSalesInvoiceQuick_AddByCustomProduct.DataTable2DataTable)
    '     Dim Filter As String = ""
    '     Dim row1 As CustomSMESS.DseTransaksiSalesInvoiceQuick_AddByCustomProduct.DataTable2Row
    '     Dim rowReturn As CustomSMESS.DseTransaksiSalesInvoiceQuick_AddByCustomProduct.DataTable3Row
    '     Dim i As Integer = 0

    '     Try
    '         For Each rowReturn In vTBLReturn.Rows
    '             Filter = "[Primary Invoice] =" + rowReturn.Primary_Sales_Invoice.ToString + " AND [Primary Product] =" + _
    '             rowReturn.Primary_Product.ToString + " AND [Primary SO] =" + rowReturn.Primary_SO.ToString + _
    '             " AND [Primary DO] =" + rowReturn.Primary_DO.ToString
    '             vTBLInvoice.DefaultView.RowFilter = Filter
    '             i = 0
    '             While vTBLInvoice.DefaultView.Count > i And rowReturn.Total_Qty_Multi > 0
    '                 row1 = vTBLInvoice.DefaultView(i).Row
    '                 If row1.Total_Qty_Multi >= rowReturn.Total_Qty_Multi Then
    '                     row1.Amount = Math.Round(((row1.Total_Qty_Multi - rowReturn.Total_Qty_Multi) * row1.Amount) / row1.Total_Qty_Multi, row1.Default_Precision_Item, MidpointRounding.AwayFromZero)
    '                     row1.Total_Qty_Multi -= rowReturn.Total_Qty_Multi
    '                     rowReturn.Total_Qty_Multi = 0
    '                 Else
    '                     row1.Total_Qty_Multi = 0
    '                     rowReturn.Total_Qty_Multi -= row1.Total_Qty_Multi
    '                 End If

    '                 If row1.Total_Qty_Multi = 0 Then
    '                     vTBLInvoice.Rows.Remove(row1)
    '                 End If
    '             End While
    '         Next

    '         vTBLInvoice.DefaultView.RowFilter = ""
    '     Catch ex As Exception
    '         MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '     End Try
    ' End Sub

    ''' <summary>
    ''' Cek setiap SO, jika ada item So dengan SOB > 0, return false
    ''' jika semua SOB <= 0, return true
    ''' </summary>
    ''' <param name="ListPrimarySO"></param>
    ''' <remarks></remarks>
    Public Function CekSOBNol(ByVal ListPrimarySO As List(Of Double)) As Boolean
        Dim DTA As New DseTransaksiSO_OrderBalanceTableAdapters.tbltrans_so_itemTableAdapter
        Dim TBL As New DseTransaksiSO_OrderBalance.tbltrans_so_itemDataTable
        Dim PrimarySO As Double
        Dim row As DseTransaksiSO_OrderBalance.tbltrans_so_itemRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim RetVal As Boolean = True
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'isikan SO
            For Each PrimarySO In ListPrimarySO
                DTA.FillBy_PrimarySO(TBL, PrimarySO)
                For Each row In TBL.Rows
                    If row.SOB > 0 Then
                        RetVal = False
                    End If
                Next
            Next
            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return True
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    '2025 Core Tax Tahap 2.2 
    Public Function SetDontCollect1(ByVal vCBOTypeTax As Integer, ByVal vUseTax As Boolean, ByVal vCboPPNDipungut As Integer) As Boolean
        Try
            If vUseTax = True And vCboPPNDipungut = 0 Then 'Cbo PPN Dipungun masih di set collect
                If vCBOTypeTax = ClsTransaksiSalesInvoice.EnumTaxType.Tax08 Or vCBOTypeTax = ClsTransaksiSalesInvoice.EnumTaxType.Tax07 Then
                    If MsgBox("Don't collect tax?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Return True
                    End If
                End If
            End If

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    '2025 Core Tax Tahap 2.2
    Public Function getLastTaxType() As Integer
        Dim DTA As New DseTransaksiSalesInvoice_TaxRateTableAdapters.tbltrans_salesinvoiceTableAdapter
        Try
            If DTA.ScalarQuery_GetLastTaxType().HasValue = True Then
                Return DTA.ScalarQuery_GetLastTaxType().Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function


End Class
