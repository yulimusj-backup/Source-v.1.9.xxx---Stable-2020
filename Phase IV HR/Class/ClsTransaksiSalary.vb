Option Explicit On
Option Strict Off

Imports Ciloci.Flee
Imports Ciloci.Flee.CalcEngine
Imports NUnit.Framework
Imports System.Windows.Forms, Core_Module
Public Class ClsTransaksiSalary
    Private mPrimary As Double = -1
    Private mNoTransaksi As String = ""
    Private mTanggalTransaksi As Date = #1/1/1950#
    Private mUrutan As Integer = 0
    Private mStartDate As Date = #1/1/1950#
    Private mEndDate As Date = #1/1/1950#
    Private mPrimaryEmployee As Double = -1
    Private mPrimaryDept As Double = -1
    Private mPrimarySalaryFormula As Double = -1
    Private mPrimaryPosition As Double = -1
    Private mTax As EnumTax = EnumTax.Gross_Up
    Private mPrimaryTaxExpense As Double = -1
    Private mIncome As Decimal = 0
    Private mDeduction As Decimal = 0
    Private mPrimaryLoanPayment As Double = -1
    Private mLoanInstallment As Decimal = 0
    Private mIncomeTax As Decimal = 0
    Private mSalaryReceived As Decimal = 0
    Private mPaymentStatus As EnumPaymentStatus = EnumPaymentStatus.unpaid
    Private mPaymentDate As Date = #1/1/1950#
    Private mRemarks As String = ""
    '-------- Tidak dipakai, didalam versi update terakhir --------
    Private mAdj_IncomeAdjust As Decimal = 0
    Private mAdj_PositionExpense As Decimal = 0
    Private mAdj_PositionExpenseAdjust As Decimal = 0
    Private mAdj_PTKPMonthly As Decimal = 0
    Private mAdj_PKPValue As Decimal = 0
    Private mAdj_PPhValue As Decimal = 0
    '---------------------------------------------------------------
    Private mPeriod_Paid As EnumPeriodPaid = EnumPeriodPaid.One_Month
    Private m01Gaji As Decimal = 0
    Private m01GajiRutin As Decimal = 0
    Private m02TunjanganPPh As Decimal = 0
    Private m02TunjanganPPHNonRutin As Decimal = 0
    Private m02TunjanganPPHRutin As Decimal = 0
    Private m02TunjanganPPHDisetahunkan As Decimal = 0
    Private m02TunjanganPPHDisetahunkanRutin As Decimal = 0
    Private m03Tunjanganlainnya As Decimal = 0
    Private m03TunjanganLainnyaRutin As Decimal = 0
    Private m04Honorarium As Decimal = 0
    Private m04HonorariumRutin As Decimal = 0
    Private m05PremiAsuransi As Decimal = 0
    Private m05PremiAsuransiRutin As Decimal = 0
    Private m06Natura As Decimal = 0
    Private m06NaturaRutin As Decimal = 0
    Private m07Tantiem As Decimal = 0
    Private m07TantiemRutin As Decimal = 0
    Private m08PenghasilanBruto As Decimal = 0
    Private m08PenghasilanBrutoRutin As Decimal = 0
    Private m09BiayaJabatan As Decimal = 0
    Private m09BiayaJabatanRutin As Decimal = 0
    Private m10IuranPensiun As Decimal = 0
    Private m10IuranPensiunRutin As Decimal = 0
    Private m11JumlahPengurangan As Decimal = 0
    Private m11JumlahPenguranganRutin As Decimal = 0
    Private m12JumlahPenghasilanBruto As Decimal = 0
    Private m13JumlahPenghasilanMsPjkSblm As Decimal = 0
    Private m13JumlahPenghasilanMsPjkSblm_1 As Decimal = 0
    Private m14JumlahPenghasilanNettoUtkPPh As Decimal = 0
    Private m14JumlahPenghasilanNettoUtkPPhRutin As Decimal = 0
    Private m15PTKP As Decimal = 0
    Private m16PKPSetahun As Decimal = 0
    Private m16PKPSetahunRutin As Decimal = 0
    Private m17PPh21 As Decimal = 0
    Private m17PPh21Rutin As Decimal = 0
    Private m18PPh21YgDiPotongMsPjkSblm As Decimal = 0
    Private m18PPh21YgDiPotongMsPjkSblm_1 As Decimal = 0
    Private m19PPhTerhutang As Decimal = 0
    Private m19PPhTerhutangSetahun As Decimal = 0
    Private m19PPhTerhutangSetahunRutin As Decimal = 0
    Private m19PPhTerhutangSebulanNonRutin As Decimal = 0
    Private m19PPhTerhutangSebulanRutin As Decimal = 0
    Private m20PPhPs21danPPh26Lunas As Decimal = 0
    Private m21Pesangon As Decimal = 0
    Private m21TunjPPhFinalPesangon As Decimal = 0
    Private m21PPhFinalPesangon As Decimal = 0
    Private m21TarifPPhFinalPesangon As Decimal = 0

    Private mA_NoForm1721A1 As String = ""
    Private mA_NPWP_Pemotongan As String = ""
    Private mA_Nama_Pemotongan As String = ""
    Private mA_NPWP As String = ""
    Private mA_NIK As String = ""
    Private mA_Nama As String = ""
    Private mA_Alamat1 As String = ""
    Private mA_Alamat2 As String = ""
    Private mA_Jenis_Kelamin As ClsTransaksiSalary.EnumJenisKelamin = ClsTransaksiSalary.EnumJenisKelamin.Wanita
    Private mA_Status_NPWP As Core_Module.ClsMasterKaryawan.GolonganPajak = Core_Module.ClsMasterKaryawan.GolonganPajak.TK0
    Private mA_Nama_Jabatan As String = 0
    Private mA_Karyawan_Asing As String = ""
    Private mA_Kode_Negara_Domisili As String = ""
    Private mC_NPWP As String = ""
    Private mC_Nama As String = ""
    Private mA_UrutanForm1721A1 As Integer = 0
    Private mPrimaryPayment As Double = -1
    Private mGrossSalary As Decimal = 0
    Private mOtherDeduction As Decimal = 0
    Private mKurangBayarTax As Decimal = 0
    Private mCompanyExpense As Decimal = 0
    Private mEmployeeExpense As Decimal = 0
    Private mPeriodStart As Integer = 1
    Private mPeriodEnd As Integer = 12
    Private mPeriodCount As Integer = 12
    Private mIsResign As Integer = 0
    Private mLosesTaxObligation As Integer = 0

    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now

    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianTransaksiHRPayroll
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter
    Dim ObjRnd As New Random
    Dim ObjAntrian As New Core_Module.ClsAntrian
    Dim ObjKary As New Core_Module.ClsMasterKaryawan

    Public Property LosesTaxObligation() As Integer
        Get
            Return Me.mLosesTaxObligation
        End Get
        Set(ByVal value As Integer)
            Me.mLosesTaxObligation = value
        End Set
    End Property

    Public Property PeriodStart() As Integer
        Get
            Return Me.mPeriodStart
        End Get
        Set(ByVal value As Integer)
            mPeriodStart = value
        End Set
    End Property

    Public Property PeriodEnd() As Integer
        Get
            Return Me.mPeriodEnd
        End Get
        Set(ByVal value As Integer)
            mPeriodEnd = value
        End Set
    End Property

    Public Property PeriodCount() As Integer
        Get
            Return Me.mPeriodCount
        End Get
        Set(ByVal value As Integer)
            mPeriodCount = value
        End Set
    End Property

    Public Property IsResign() As Integer
        Get
            Return Me.mIsResign
        End Get
        Set(ByVal value As Integer)
            mIsResign = value
        End Set
    End Property

    Public Property CompanyExpense() As Decimal
        Get
            Return Me.mCompanyExpense
        End Get
        Set(ByVal value As Decimal)
            mCompanyExpense = value
        End Set
    End Property

    Public Property EmployeeExpense() As Decimal
        Get
            Return Me.mEmployeeExpense
        End Get
        Set(ByVal value As Decimal)
            mEmployeeExpense = value
        End Set
    End Property

    Public Property PrimaryPayment() As Double
        Get
            Return mPrimaryPayment
        End Get
        Set(ByVal value As Double)
            mPrimaryPayment = value
        End Set
    End Property

    Public Property GrossSalary() As Decimal
        Get
            Return mGrossSalary
        End Get
        Set(ByVal value As Decimal)
            mGrossSalary = value
        End Set
    End Property

    Public Property OtherDeduction() As Decimal
        Get
            Return mOtherDeduction
        End Get
        Set(ByVal value As Decimal)
            mOtherDeduction = value
        End Set
    End Property

    Public Property KurangBayarTax() As Decimal
        Get
            Return mKurangBayarTax
        End Get
        Set(ByVal value As Decimal)
            mKurangBayarTax = value
        End Set
    End Property

    Public Property A_NoForm1721A1() As String
        Get
            Return mA_NoForm1721A1
        End Get
        Set(ByVal value As String)
            mA_NoForm1721A1 = value
        End Set
    End Property

    Public Property A_NPWP_Pemotongan() As String
        Get
            Return mA_NPWP_Pemotongan
        End Get
        Set(ByVal value As String)
            mA_NPWP_Pemotongan = value
        End Set
    End Property

    Public Property A_Nama_Pemotongan() As String
        Get
            Return Me.mA_Nama_Pemotongan
        End Get
        Set(ByVal value As String)
            mA_Nama_Pemotongan = value
        End Set
    End Property

    Public Property A_NPWP() As String
        Get
            Return Me.mA_NPWP
        End Get
        Set(ByVal value As String)
            mA_NPWP = value
        End Set
    End Property

    Public Property A_NIK() As String
        Get
            Return Me.mA_NIK
        End Get
        Set(ByVal value As String)
            Me.mA_NIK = value
        End Set
    End Property

    Public Property A_Nama() As String
        Get
            Return Me.mA_Nama
        End Get
        Set(ByVal value As String)
            mA_Nama = value
        End Set
    End Property

    Public Property A_Alamat1() As String
        Get
            Return mA_Alamat1
        End Get
        Set(ByVal value As String)
            mA_Alamat1 = value
        End Set
    End Property

    Public Property A_Alamat2() As String
        Get
            Return Me.mA_Alamat2
        End Get
        Set(ByVal value As String)
            mA_Alamat2 = value
        End Set
    End Property

    Public Property A_Jenis_Kelamin() As ClsTransaksiSalary.EnumJenisKelamin
        Get
            Return mA_Jenis_Kelamin
        End Get
        Set(ByVal value As ClsTransaksiSalary.EnumJenisKelamin)
            mA_Jenis_Kelamin = value
        End Set
    End Property

    Public Property A_Status_NPWP() As Core_Module.ClsMasterKaryawan.GolonganPajak
        Get
            Return Me.mA_Status_NPWP
        End Get
        Set(ByVal value As Core_Module.ClsMasterKaryawan.GolonganPajak)
            mA_Status_NPWP = value
        End Set
    End Property

    Public Property A_Nama_Jabatan() As String
        Get
            Return Me.mA_Nama_Jabatan
        End Get
        Set(ByVal value As String)
            mA_Nama_Jabatan = value
        End Set
    End Property

    Public Property A_Karyawan_Asing() As String
        Get
            Return Me.mA_Karyawan_Asing
        End Get
        Set(ByVal value As String)
            mA_Karyawan_Asing = value
        End Set
    End Property

    Public Property A_Kode_Negara_Domisili() As String
        Get
            Return Me.mA_Kode_Negara_Domisili
        End Get
        Set(ByVal value As String)
            mA_Kode_Negara_Domisili = value
        End Set
    End Property

    Public Property C_NPWP() As String
        Get
            Return mC_NPWP
        End Get
        Set(ByVal value As String)
            mC_NPWP = value
        End Set
    End Property

    Public Property C_Nama() As String
        Get
            Return mC_Nama
        End Get
        Set(ByVal value As String)
            mC_Nama = value
        End Set
    End Property

    Public Property A_UrutanForm1721A1() As Integer
        Get
            Return mA_UrutanForm1721A1
        End Get
        Set(ByVal value As Integer)
            mA_UrutanForm1721A1 = value
        End Set
    End Property

    Public Enum EnumJenisKelamin As Integer
        Pria
        Wanita
    End Enum

    Public Property _01Gaji() As Decimal
        Get
            Return Me.m01Gaji
        End Get
        Set(ByVal value As Decimal)
            m01Gaji = value
        End Set
    End Property

    Public Property _01GajiRutin() As Decimal
        Get
            Return Me.m01GajiRutin
        End Get
        Set(ByVal value As Decimal)
            m01GajiRutin = value
        End Set
    End Property

    Public Property _02TunjanganPPh() As Decimal
        Get
            Return m02TunjanganPPh
        End Get
        Set(ByVal value As Decimal)
            m02TunjanganPPh = value
        End Set
    End Property

    Public Property _02TunjanganPPhDisetahunkan() As Decimal
        Get
            Return m02TunjanganPPHDisetahunkan
        End Get
        Set(ByVal value As Decimal)
            m02TunjanganPPHDisetahunkan = value
        End Set
    End Property

    Public Property _02TunjanganPPhDisetahunkanRutin() As Decimal
        Get
            Return m02TunjanganPPHDisetahunkanRutin
        End Get
        Set(ByVal value As Decimal)
            m02TunjanganPPHDisetahunkanRutin = value
        End Set
    End Property

    Public Property _02TunjanganPPhNonRutin() As Decimal
        Get
            Return m02TunjanganPPHNonRutin
        End Get
        Set(ByVal value As Decimal)
            m02TunjanganPPHNonRutin = value
        End Set
    End Property

    Public Property _02TunjanganPPhRutin() As Decimal
        Get
            Return m02TunjanganPPHRutin
        End Get
        Set(ByVal value As Decimal)
            m02TunjanganPPHRutin = value
        End Set
    End Property

    Public Property _03Tunjanganlainnya() As Decimal
        Get
            Return m03Tunjanganlainnya
        End Get
        Set(ByVal value As Decimal)
            m03Tunjanganlainnya = value
        End Set
    End Property

    Public Property _03TunjanganlainnyaRutin() As Decimal
        Get
            Return m03TunjanganLainnyaRutin
        End Get
        Set(ByVal value As Decimal)
            m03TunjanganLainnyaRutin = value
        End Set
    End Property

    Public Property _04Honorarium() As Decimal
        Get
            Return m04Honorarium
        End Get
        Set(ByVal value As Decimal)
            m04Honorarium = value
        End Set
    End Property

    Public Property _04HonorariumRutin() As Decimal
        Get
            Return m04HonorariumRutin
        End Get
        Set(ByVal value As Decimal)
            m04HonorariumRutin = value
        End Set
    End Property

    Public Property _05PremiAsuransi() As Decimal
        Get
            Return m05PremiAsuransi
        End Get
        Set(ByVal value As Decimal)
            m05PremiAsuransi = value
        End Set
    End Property

    Public Property _05PremiAsuransiRutin() As Decimal
        Get
            Return m05PremiAsuransiRutin
        End Get
        Set(ByVal value As Decimal)
            m05PremiAsuransiRutin = value
        End Set
    End Property

    Public Property _06Natura() As Decimal
        Get
            Return Me.m06Natura
        End Get
        Set(ByVal value As Decimal)
            m06Natura = value
        End Set
    End Property

    Public Property _06NaturaRutin() As Decimal
        Get
            Return Me.m06NaturaRutin
        End Get
        Set(ByVal value As Decimal)
            m06NaturaRutin = value
        End Set
    End Property

    Public Property _07Tantiem() As Decimal
        Get
            Return m07Tantiem
        End Get
        Set(ByVal value As Decimal)
            m07Tantiem = value
        End Set
    End Property

    Public Property _07TantiemRutin() As Decimal
        Get
            Return m07TantiemRutin
        End Get
        Set(ByVal value As Decimal)
            m07TantiemRutin = value
        End Set
    End Property

    Public Property _08PenghasilanBruto() As Decimal
        Get
            Return m08PenghasilanBruto
        End Get
        Set(ByVal value As Decimal)
            m08PenghasilanBruto = value
        End Set
    End Property

    Public Property _08PenghasilanBrutoRutin() As Decimal
        Get
            Return m08PenghasilanBrutoRutin
        End Get
        Set(ByVal value As Decimal)
            m08PenghasilanBrutoRutin = value
        End Set
    End Property

    Public Property _09BiayaJabatan() As Decimal
        Get
            Return m09BiayaJabatan
        End Get
        Set(ByVal value As Decimal)
            m09BiayaJabatan = value
        End Set
    End Property

    Public Property _09BiayaJabatanRutin() As Decimal
        Get
            Return m09BiayaJabatanRutin
        End Get
        Set(ByVal value As Decimal)
            m09BiayaJabatanRutin = value
        End Set
    End Property

    Public Property _10IuranPensiun() As Decimal
        Get
            Return Me.m10IuranPensiun
        End Get
        Set(ByVal value As Decimal)
            m10IuranPensiun = value
        End Set
    End Property

    Public Property _10IuranPensiunRutin() As Decimal
        Get
            Return Me.m10IuranPensiunRutin
        End Get
        Set(ByVal value As Decimal)
            m10IuranPensiunRutin = value
        End Set
    End Property

    Public Property _11JumlahPengurangan() As Decimal
        Get
            Return m11JumlahPengurangan
        End Get
        Set(ByVal value As Decimal)
            m11JumlahPengurangan = value
        End Set
    End Property

    Public Property _11JumlahPenguranganRutin() As Decimal
        Get
            Return m11JumlahPenguranganRutin
        End Get
        Set(ByVal value As Decimal)
            m11JumlahPenguranganRutin = value
        End Set
    End Property

    Public Property _12JumlahPenghasilanBruto() As Decimal
        Get
            Return m12JumlahPenghasilanBruto
        End Get
        Set(ByVal value As Decimal)
            m12JumlahPenghasilanBruto = value
        End Set
    End Property

    Public Property _13JumlahPenghasilanMsPjkSblm() As Decimal
        Get
            Return m13JumlahPenghasilanMsPjkSblm
        End Get
        Set(ByVal value As Decimal)
            m13JumlahPenghasilanMsPjkSblm = value
        End Set
    End Property

    Public Property _13JumlahPenghasilanMsPjkSblm_1() As Decimal
        Get
            Return m13JumlahPenghasilanMsPjkSblm_1
        End Get
        Set(ByVal value As Decimal)
            m13JumlahPenghasilanMsPjkSblm_1 = value
        End Set
    End Property

    Public Property _14JumlahPenghasilanNettoUtkPPh() As Decimal
        Get
            Return Me.m14JumlahPenghasilanNettoUtkPPh
        End Get
        Set(ByVal value As Decimal)
            m14JumlahPenghasilanNettoUtkPPh = value
        End Set
    End Property

    Public Property _14JumlahPenghasilanNettoUtkPPhRutin() As Decimal
        Get
            Return Me.m14JumlahPenghasilanNettoUtkPPhRutin
        End Get
        Set(ByVal value As Decimal)
            m14JumlahPenghasilanNettoUtkPPhRutin = value
        End Set
    End Property

    Public Property _15PTKP() As Decimal
        Get
            Return Me.m15PTKP
        End Get
        Set(ByVal value As Decimal)
            m15PTKP = value
        End Set
    End Property

    Public Property _16PKPSetahun() As Decimal
        Get
            Return Me.m16PKPSetahun
        End Get
        Set(ByVal value As Decimal)
            m16PKPSetahun = value
        End Set
    End Property

    Public Property _16PKPSetahunRutin() As Decimal
        Get
            Return Me.m16PKPSetahunRutin
        End Get
        Set(ByVal value As Decimal)
            m16PKPSetahunRutin = value
        End Set
    End Property

    Public Property _17PPh21() As Decimal
        Get
            Return m17PPh21
        End Get
        Set(ByVal value As Decimal)
            m17PPh21 = value
        End Set
    End Property

    Public Property _17PPh21Rutin() As Decimal
        Get
            Return m17PPh21Rutin
        End Get
        Set(ByVal value As Decimal)
            m17PPh21Rutin = value
        End Set
    End Property

    Public Property _18PPh21YgDiPotongMsPjkSblm() As Decimal
        Get
            Return Me.m18PPh21YgDiPotongMsPjkSblm
        End Get
        Set(ByVal value As Decimal)
            m18PPh21YgDiPotongMsPjkSblm = value
        End Set
    End Property

    Public Property _18PPh21YgDiPotongMsPjkSblm_1() As Decimal
        Get
            Return Me.m18PPh21YgDiPotongMsPjkSblm_1
        End Get
        Set(ByVal value As Decimal)
            m18PPh21YgDiPotongMsPjkSblm_1 = value
        End Set
    End Property

    Public Property _19PPhTerhutang() As Decimal
        Get
            Return m19PPhTerhutang
        End Get
        Set(ByVal value As Decimal)
            m19PPhTerhutang = value
        End Set
    End Property

    Public Property _19PPhTerhutangSebulanNonRutin() As Decimal
        Get
            Return m19PPhTerhutangSebulanNonRutin
        End Get
        Set(ByVal value As Decimal)
            m19PPhTerhutangSebulanNonRutin = value
        End Set
    End Property

    Public Property _19PPhTerhutangSebulanRutin() As Decimal
        Get
            Return m19PPhTerhutangSebulanRutin
        End Get
        Set(ByVal value As Decimal)
            m19PPhTerhutangSebulanRutin = value
        End Set
    End Property

    Public Property _19PPhTerhutangSetahun() As Decimal
        Get
            Return m19PPhTerhutangSetahun
        End Get
        Set(ByVal value As Decimal)
            m19PPhTerhutangSetahun = value
        End Set
    End Property

    Public Property _19PPhTerhutangSetahunRutin() As Decimal
        Get
            Return m19PPhTerhutangSetahunRutin
        End Get
        Set(ByVal value As Decimal)
            m19PPhTerhutangSetahunRutin = value
        End Set
    End Property

    Public Property _20PPhPs21danPPh26Lunas() As Decimal
        Get
            Return m20PPhPs21danPPh26Lunas
        End Get
        Set(ByVal value As Decimal)
            m20PPhPs21danPPh26Lunas = value
        End Set
    End Property

    Public Property _21Pesangon() As Decimal
        Get
            Return Me.m21Pesangon
        End Get
        Set(ByVal value As Decimal)
            Me.m21Pesangon = value
        End Set
    End Property

    Public Property _21TunjPPhFinalPesangon() As Decimal
        Get
            Return Me.m21TunjPPhFinalPesangon
        End Get
        Set(ByVal value As Decimal)
            Me.m21TunjPPhFinalPesangon = value
        End Set
    End Property

    Public Property _21PPhFinalPesangon() As Decimal
        Get
            Return Me.m21PPhFinalPesangon
        End Get
        Set(ByVal value As Decimal)
            Me.m21PPhFinalPesangon = value
        End Set
    End Property

    Public Property _21TarifPPhFinalPesangon() As Decimal
        Get
            Return Me.m21TarifPPhFinalPesangon
        End Get
        Set(ByVal value As Decimal)
            Me.m21TarifPPhFinalPesangon = value
        End Set
    End Property

    Public Enum EnumPeriodPaid As Integer
        One_Month
        One_Week
        Two_Weeks
    End Enum

    Public Property Period_Paid() As EnumPeriodPaid
        Get
            Return mPeriod_Paid
        End Get
        Set(ByVal value As EnumPeriodPaid)
            mPeriod_Paid = value
        End Set
    End Property

    Public Property Adj_PPhValue() As Decimal
        Get
            Return mAdj_PPhValue
        End Get
        Set(ByVal value As Decimal)
            mAdj_PPhValue = value
        End Set
    End Property

    Public Property Adj_IncomeAdjust() As Decimal
        Get
            Return mAdj_IncomeAdjust
        End Get
        Set(ByVal value As Decimal)
            mAdj_IncomeAdjust = value
        End Set
    End Property

    Public Property Adj_PositionExpense() As Decimal
        Get
            Return mAdj_PositionExpense
        End Get
        Set(ByVal value As Decimal)
            mAdj_PositionExpense = value
        End Set
    End Property

    Public Property Adj_PositionExpenseAdjust() As Decimal
        Get
            Return mAdj_PositionExpenseAdjust
        End Get
        Set(ByVal value As Decimal)
            mAdj_PositionExpenseAdjust = value
        End Set
    End Property

    Public Property Adj_PTKPMonthly() As Decimal
        Get
            Return mAdj_PTKPMonthly
        End Get
        Set(ByVal value As Decimal)
            mAdj_PTKPMonthly = value
        End Set
    End Property

    Public Property Adj_PKPValue() As Decimal
        Get
            Return mAdj_PKPValue
        End Get
        Set(ByVal value As Decimal)
            mAdj_PKPValue = value
        End Set
    End Property

    Public Enum EnumPaymentStatus As Integer
        unpaid
        paid
    End Enum

    Public Enum EnumTax As Integer
        Gross_Up
        Netto
        Gross
        Manual_Tax_Calculation
    End Enum

#Region "Property"
    Public Property Remarks() As String
        Get
            Return mRemarks
        End Get
        Set(ByVal value As String)
            mRemarks = value
        End Set
    End Property

    Public Property PaymentDate() As Date
        Get
            Return mPaymentDate
        End Get
        Set(ByVal value As Date)
            mPaymentDate = value
        End Set
    End Property

    Public Property PaymentStatus() As EnumPaymentStatus
        Get
            Return mPaymentStatus
        End Get
        Set(ByVal value As EnumPaymentStatus)
            mPaymentStatus = value
        End Set
    End Property

    Public Property SalaryReceived() As Decimal
        Get
            Return mSalaryReceived
        End Get
        Set(ByVal value As Decimal)
            mSalaryReceived = value
        End Set
    End Property

    Public Property IncomeTax() As Decimal
        Get
            Return mIncomeTax
        End Get
        Set(ByVal value As Decimal)
            mIncomeTax = value
        End Set
    End Property

    Public Property LoanInstallment() As Decimal
        Get
            Return mLoanInstallment
        End Get
        Set(ByVal value As Decimal)
            mLoanInstallment = value
        End Set
    End Property

    Public Property PrimaryLoanPayment() As Double
        Get
            Return mPrimaryLoanPayment
        End Get
        Set(ByVal value As Double)
            mPrimaryLoanPayment = value
        End Set
    End Property

    Public Property Deduction() As Decimal
        Get
            Return mDeduction
        End Get
        Set(ByVal value As Decimal)
            mDeduction = value
        End Set
    End Property

    Public Property Income() As Decimal
        Get
            Return mIncome
        End Get
        Set(ByVal value As Decimal)
            mIncome = value
        End Set
    End Property

    Public Property PrimaryTaxExpense() As Double
        Get
            Return mPrimaryTaxExpense
        End Get
        Set(ByVal value As Double)
            mPrimaryTaxExpense = value
        End Set
    End Property

    Public Property Tax() As EnumTax
        Get
            Return mTax
        End Get
        Set(ByVal value As EnumTax)
            mTax = value
        End Set
    End Property

    Public Property PrimaryPosition() As Double
        Get
            Return mPrimaryPosition
        End Get
        Set(ByVal value As Double)
            mPrimaryPosition = value
        End Set
    End Property

    Public Property PrimarySalaryFormula() As Double
        Get
            Return mPrimarySalaryFormula
        End Get
        Set(ByVal value As Double)
            mPrimarySalaryFormula = value
        End Set
    End Property

    Public Property PrimaryEmployee() As Double
        Get
            Return mPrimaryEmployee
        End Get
        Set(ByVal value As Double)
            mPrimaryEmployee = value
        End Set
    End Property

    Public Property PrimaryDept() As Double
        Get
            Return mPrimaryDept
        End Get
        Set(ByVal value As Double)
            mPrimaryDept = value
        End Set
    End Property

    Public Property EndDate() As Date
        Get
            Return mEndDate
        End Get
        Set(ByVal value As Date)
            mEndDate = value
        End Set
    End Property

    Public Property StartDate() As Date
        Get
            Return mStartDate
        End Get
        Set(ByVal value As Date)
            mStartDate = value
        End Set
    End Property

    Public Property Urutan() As Integer
        Get
            Return mUrutan
        End Get
        Set(ByVal value As Integer)
            mUrutan = value
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

    Public Property NoTransaksi() As String
        Get
            Return mNoTransaksi
        End Get
        Set(ByVal value As String)
            mNoTransaksi = value
        End Set
    End Property

    Public Property TanggalTransaksi() As Date
        Get
            Return mTanggalTransaksi
        End Get
        Set(ByVal value As Date)
            mTanggalTransaksi = value
        End Set
    End Property

    Public Property CreatedBy() As String
        Get
            Return Me.mCreatedBy
        End Get
        Set(ByVal value As String)
            mCreatedBy = value
        End Set
    End Property

    Public Property CreatedDate() As Date
        Get
            Return Me.mCreatedDate
        End Get
        Set(ByVal value As Date)
            mCreatedDate = value
        End Set
    End Property

    Public Property ModifiedBy() As String
        Get
            Return Me.mModifiedBy
        End Get
        Set(ByVal value As String)
            mModifiedBy = value
        End Set
    End Property

    Public Property ModifiedDate() As Date
        Get
            Return Me.mModifiedDate
        End Get
        Set(ByVal value As Date)
            mModifiedDate = value
        End Set
    End Property

    Public Property TimeStamp() As Date
        Get
            Return Me.mTimestamp
        End Get
        Set(ByVal value As Date)
            Me.mTimestamp = value
        End Set
    End Property
#End Region

    Public Function GetJmlByNoTransaksi(ByVal vNoTrans As String) As Double
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salaryTableAdapter
        Dim Dummy As System.Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetJmlByNoTransaksi(vNoTrans)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryEmployee(ByVal vNoTrans As String) As Double
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salaryTableAdapter
        Dim Dummy As System.Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetPrimaryEmployee(Me.GetPrimary(vNoTrans))
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryEmployeeByPrimary(ByVal vPrimarySalary As Double) As Double
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salary2TableAdapter
        Dim Dummy As System.Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetPrimaryEmployee(vPrimarySalary)
            If Dummy.HasValue Then
                Return Dummy.Value
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
    ''' mengembalikan primary transaksi yang tanggal periode penggajiannya mencakup tanggal yang diberikan pd parameter
    ''' </summary>
    ''' <param name="vDate"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetPrimary_ByPeriod(ByVal vDate As Date, ByVal vPrimaryKary As Double) As Double
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salary2TableAdapter
        Dim Dummy As System.Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_PrimaryByTgl(vDate, vPrimaryKary)
            If Dummy.HasValue Then
                Return Dummy.Value
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

    Public Function GetStartDate(ByVal vPrimary As Double) As Date
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salary2TableAdapter
        Dim Dummy As System.Nullable(Of Date)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetStartDate(vPrimary)
            If Dummy.HasValue Then
                Return Dummy.Value
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

    Public Function GetPrimary(ByVal vNoTrans As String) As Double
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salaryTableAdapter
        Dim Dummy As System.Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetPrimary(vNoTrans)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetCreatedBy(ByVal vPrimary As Double) As String
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salaryTableAdapter
        Dim Dummy As String
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetCreatedBy(vPrimary)
            If String.IsNullOrEmpty(Dummy) = False Then
                Return Dummy
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function GetMAXUrutan(ByVal vEndDate As Date, ByVal vStartDate As Date) As Double
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salaryTableAdapter
        Dim Dummy As System.Nullable(Of Integer)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetMAXUrutan(vEndDate, vStartDate)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetNoTransaksi(ByVal vPrimary As Double) As String
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salaryTableAdapter
        Dim Dummy As String
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetNoTransaksi(vPrimary)
            If String.IsNullOrEmpty(Dummy) = False Then
                Return Dummy
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

    Public Function GetNoTransaksi_ByAttendance(ByVal vPrimaryEmployee As Double, ByVal vDate As Date) As String
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salaryTableAdapter
        Dim Dummy As String
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetNoTransaksi1(vPrimaryEmployee, vDate)
            If String.IsNullOrEmpty(Dummy) = False Then
                Return Dummy
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

    Public Function GetTimestamp(ByVal vPrimary As Double) As Date
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salaryTableAdapter
        Dim Dummy As System.Nullable(Of Date)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetTimestamp(vPrimary)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return Date.MinValue
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTanggalTransaksi(ByVal vPrimary As Double) As Date
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salaryTableAdapter
        Dim Dummy As System.Nullable(Of Date)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetTanggalTransaksi(vPrimary)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return Date.MinValue
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahBolehFindDataSalary(ByVal vPrimarySalary As Double, ByVal vUserCode As String) As Boolean
        Dim DTAUser As New DseTransPaymentTableAdapters.TblPayrollAccessTableAdapter
        Dim TBLUser As New DseTransPayment.TblPayrollAccessDataTable
        Dim DTASalary As New DseTransPaymentTableAdapters.tbltrans_salaryTableAdapter
        Dim TBLSalary As New DseTransPayment.tbltrans_salaryDataTable
        Dim RowSalary As DseTransPayment.tbltrans_salaryRow
        Dim PrimarySalaryFormula As Double = -1
        Dim SalaryNo As String = ""
        Dim Filter As String = ""
        Dim Retval As String = ""
        Try
            DTAUser.Connection.ConnectionString = Core_Module.GetConstring
            DTASalary.Connection.ConnectionString = Core_Module.GetConstring
            DTAUser.FillBy_KodeUser(TBLUser, vUserCode)
            DTASalary.FillBy_PrimarySalary(TBLSalary, vPrimarySalary)
            If TBLSalary.Rows.Count > 0 Then
                RowSalary = TBLSalary.Rows(0)
                PrimarySalaryFormula = RowSalary.Primary_Salary_Formula
                Filter = "[Primary Salary Formula] = " + PrimarySalaryFormula.ToString + ""
                TBLUser.DefaultView.RowFilter = Filter
                If TBLUser.DefaultView.Count <= 0 Then
                    MsgBox(vUserCode + " doesn't have payroll access for salary no : " + RowSalary.No_Transaksi.ToString + ".", MsgBoxStyle.Information, "Information")
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
            DTAUser.Dispose()
            TBLUser.Dispose()
            DTASalary.Dispose()
            TBLSalary.Dispose()
        End Try
    End Function

    Public Function GetPenghasilanBruto(ByVal vPrimary As Double) As Decimal
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salary2TableAdapter
        Dim Retval As Nullable(Of Decimal)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetPenghasilanBruto(vPrimary)
            If Retval.HasValue = True Then
                Return Retval
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

#Region "Standar Coding"
    Public Function CariDetail(ByVal vPrimary As Double, ByVal vObjCD As Core_Module.ClsCekData, _
    ByVal vDGVItem As DataGridView, ByVal vDGVItemOther As DataGridView, _
    ByVal vDGVPayment As DataGridView, ByVal vDef_Prec As Double, ByVal i As FrmTrans_Gaji, _
    ByVal ij As FrmTrans_Gaji_Payment, ByVal vTBLItemIterasi As DseTransaksiSalary.tbltrans_salary_item_iterationDataTable, _
    ByVal vTBLItemTax As DseTransaksiSalary.tbltrans_salary_itemtaxDataTable, ByVal vTBLOtherPayment As DseTransaksiSalary.tbltrans_salary_item_other1DataTable, _
    ByRef vPenghasilanBruto As Decimal, ByVal vTBLPrevIncome As DseTransaksiSalary.tbltrans_salary_item_prev_income1DataTable) As Boolean
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salary2TableAdapter
        Dim TBL As New DseTransaksiSalary.tbltrans_salary2DataTable
        Dim row As DseTransaksiSalary.tbltrans_salary2Row
        Dim RowDGV As DataGridViewRow
        Dim ObjItem As New ClsTransaksiSalary_Item
        Dim ObjCD As New Core_Module.ClsCekData
        Dim FrmOther As New FrmTrans_Gaji_OtherDeduction

        Try
            vPenghasilanBruto = Me.GetPenghasilanBruto(vPrimary)
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DTA.FillBy_Primary(TBL, vPrimary)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = .Primary
                    Me.mNoTransaksi = .No_Transaksi
                    Me.mTanggalTransaksi = .Tanggal_Transaksi
                    Me.mUrutan = .Urutan
                    Me.mPrimaryEmployee = .Primary_Employee
                    Me.mStartDate = .Start_Date
                    Me.mEndDate = .End_Date
                    Me.mPrimarySalaryFormula = .Primary_Salary_Formula
                    Me.mPrimaryPosition = .Primary_Position
                    Me.mTax = .Tax
                    Me.mPrimaryTaxExpense = .Primary_Tax_Expense_Account
                    Me.mIncome = .Income
                    Me.mDeduction = .Deduction
                    Me.mPrimaryLoanPayment = .Primary_Loan_Payment
                    Me.mLoanInstallment = .Loan_Installment
                    Me.mIncomeTax = .Income_Tax
                    Me.mSalaryReceived = .Salary_Received
                    Me.mPaymentStatus = .Payment_Status
                    Me.mPaymentDate = .Payment_Date
                    Me.mRemarks = .Remarks
                    Me.mEmployeeExpense = .Employee_Expense
                    Me.mCompanyExpense = .Company_Expense
                    Me.mPeriodCount = .Period_Count
                    Me.mPeriodEnd = .Period_End
                    Me.mPeriodStart = .Period_Start
                    Me.mPrimaryDept = .Primary_Dept
                    Me.mIsResign = .Is_Resign

                    '---------------------------------
                    Me.mAdj_IncomeAdjust = 0
                    Me.mAdj_PKPValue = 0
                    Me.mAdj_PositionExpense = 0
                    Me.mAdj_PositionExpenseAdjust = 0
                    Me.mAdj_PTKPMonthly = 0
                    Me.mAdj_PPhValue = 0
                    '----------------------------------
                    Me.mPeriod_Paid = .Period_Paid
                    Me.m01Gaji = ._01Gaji
                    Me.m01GajiRutin = ._01_Gaji_Rutin
                    Me.m02TunjanganPPh = ._02TunjanganPPh
                    Me.m02TunjanganPPHDisetahunkan = ._02_Tunjangan_PPH_Disetahunkan
                    Me.m02TunjanganPPHDisetahunkanRutin = ._02_Tunjangan_PPH_Disetahunkan_Rutin
                    Me.m02TunjanganPPHNonRutin = ._02_Tunjangan_PPH_Non_Rutin
                    Me.m02TunjanganPPHRutin = ._02_Tunjangan_PPH_Rutin
                    Me.m03Tunjanganlainnya = ._03TunjanganLainnya
                    Me.m03TunjanganLainnyaRutin = ._03_Tunjangan_Lainnya_Rutin
                    Me.m04Honorarium = ._04Honorarium
                    Me.m04HonorariumRutin = ._04_Honorarium_Rutin
                    Me.m05PremiAsuransi = ._05PremiAsuransi
                    Me.m05PremiAsuransiRutin = ._05_Premi_Asurasi_Rutin
                    Me.m06Natura = ._06Natura
                    Me.m06NaturaRutin = ._06_Natura_Rutin
                    Me.m07Tantiem = ._07Tantiem
                    Me.m07TantiemRutin = ._07_Tantiem_Rutin
                    Me.m08PenghasilanBruto = ._08PenghasilanBruto
                    Me.m08PenghasilanBrutoRutin = ._08_Penghasilan_Bruto_Rutin
                    Me.m09BiayaJabatan = ._09BiayaJabatan
                    Me.m09BiayaJabatanRutin = ._09_Biaya_Jabatan_Rutin
                    Me.m10IuranPensiun = ._10IuranPensiun
                    Me.m10IuranPensiunRutin = ._10_Iuran_Pensiun_Rutin
                    Me.m11JumlahPengurangan = row._11JumlahPengurangan
                    Me.m11JumlahPenguranganRutin = row._11_Jumlah_Pengurang_Rutin
                    Me.m12JumlahPenghasilanBruto = row._12JumlahPenghasilanNetto
                    Me.m13JumlahPenghasilanMsPjkSblm = row._13JumlahPenghasilanMasaPjkSblmnya
                    Me.m13JumlahPenghasilanMsPjkSblm_1 = row._13JumlahPenghasilanMasaPjkSblmnya1
                    Me.m14JumlahPenghasilanNettoUtkPPh = row._14JumlahPenghasilanNettoUtkPPh
                    Me.m14JumlahPenghasilanNettoUtkPPhRutin = row._14_Jumlah_Penghasilan_Netto_Unt_PPH_Rutin
                    Me.m15PTKP = ._15PTKP
                    Me.m16PKPSetahunRutin = row._16_PKP_Setahun_Rutin
                    Me.m16PKPSetahun = ._16PKPSetahun
                    Me.m17PPh21 = ._17PPh21
                    Me.m17PPh21Rutin = ._17_PPH_21_Rutin
                    Me.m18PPh21YgDiPotongMsPjkSblm = ._18PPh21YgTelahDipotongMasaPjkSblmnya
                    Me.m18PPh21YgDiPotongMsPjkSblm_1 = ._18PPh21YgTelahDipotongMasaPjkSblmnya1
                    Me.m19PPhTerhutang = ._19PPhTerhutang
                    Me.m19PPhTerhutangSebulanNonRutin = ._19_PPH_Terhutang_Sebulan_Non_Rutin
                    Me.m19PPhTerhutangSebulanRutin = ._19_PPH_Terutang_Sebulan_Rutin
                    Me.m19PPhTerhutangSetahun = ._19_PPH_Terhutang_Setahun
                    Me.m19PPhTerhutangSetahunRutin = ._19_PPH_Terhutang_Setahun_Rutin
                    Me.m20PPhPs21danPPh26Lunas = ._20PPhPs21DanPPhPs26YgTelahDipotongNLunas
                    Me.m21Pesangon = ._21_Pesangon
                    Me.m21PPhFinalPesangon = ._21_PPH_Final_Pesangon
                    Me.m21TunjPPhFinalPesangon = ._21_Tunj_PPH_FInal_Pesangon
                    Me.m21TarifPPhFinalPesangon = .Tarif_Pesangon

                    Me.mA_NoForm1721A1 = .NoForm1721A1
                    Me.mA_NPWP_Pemotongan = .NPWP_Pemotong
                    Me.mA_Nama_Pemotongan = .Nama_Pemotong
                    Me.mA_NPWP = .A_NPWP
                    Me.mA_NIK = .A_NIK
                    Me.mA_Nama = .A_Nama
                    Me.mA_Alamat1 = .A_Alamat1
                    Me.mA_Alamat2 = .A_Alamat2
                    Me.mA_Jenis_Kelamin = .A_Jenis_Kelamin
                    Me.mA_Status_NPWP = .A_Status_NPWP
                    Me.mA_Nama_Jabatan = .A_Nama_Jabatan
                    Me.mA_Karyawan_Asing = .A_Karyawan_Asing
                    Me.mA_Kode_Negara_Domisili = .A_Kode_Negara_Domisili
                    Me.mC_NPWP = .C_NPWP
                    Me.mC_Nama = .C_Nama
                    Me.mA_UrutanForm1721A1 = .UrutanForm1721A1
                    Me.mPrimaryPayment = .Primary_Payment
                    Me.mGrossSalary = .Gross_Salary
                    Me.mOtherDeduction = .Other_Deduction
                    Me.mKurangBayarTax = .Kurang_Bayar_Tax

                    Me.mLosesTaxObligation = .loses_subjective_tax_obligations

                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
                End With
                'Pengecekan Status Paid dan UnPaid
                If Me.GetPaymentStatus(Me.mPrimary) = EnumPaymentStatus.paid Then
                    i.DtpPaymentStatus.Value = Me.GetPaymentDate(Me.mPrimary)
                    i.CboPaymentStatus.SelectedIndex = EnumPaymentStatus.paid
                    i.LblPaid.Text = "Paid"
                Else
                    i.DtpPaymentStatus.Value = #1/1/1950#
                    i.CboPaymentStatus.SelectedIndex = EnumPaymentStatus.unpaid
                    i.LblPaid.Text = "Unpaid"
                End If
                'cari data item
                ObjItem.IsiDGV(Primary, vDGVItem, vDef_Prec, i, vTBLItemIterasi, vTBLItemTax, vTBLOtherPayment, vTBLPrevIncome)
                ObjItem.IsiDGVOtherDeduction(Primary, vDGVItemOther, vDef_Prec, FrmOther, i)
                i.TxtOtherDedution.Text = "0"
                If vDGVItemOther.Rows.Count > 0 Then
                    For Each RowDGV In vDGVItemOther.Rows
                        i.TxtOtherDedution.Text = vObjCD.DigitGrouping(CDec(i.TxtOtherDedution.Text) + RowDGV.Cells(i.ColAmount.Index).Value, vDef_Prec)
                    Next
                End If
                'cari data payment item
                ObjItem.IsiDGVItemPayment(Primary, vDGVPayment, vDef_Prec, ij)
                i.TxtPayment.Text = "0"
                Dim HitunganPayment As Decimal = 0
                If vDGVPayment.Rows.Count > 0 Then
                    For Each RowDGV In vDGVPayment.Rows
                        HitunganPayment += Convert.ToDecimal(RowDGV.Cells(ij.ColDetailPayment_Amount.Index).Value)
                    Next
                End If
                i.TxtPayment.Text = ObjCD.DigitGrouping(HitunganPayment, vDef_Prec)
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjItem = Nothing
            ObjCD = Nothing
            FrmOther.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As Core_Module.EnumSearchCriteria, _
    Optional ByVal vEndDate As Date = #1/1/1950#, Optional ByVal vStartDate As Date = #1/1/1950#) As Double
        Dim DTA As New DseTransaksiSalaryTableAdapters.DataTable1TableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vSearchBy = Core_Module.EnumSearchCriteria.ByNoTransaksi Then
                Return DTA.ScalarQuery_GetJmlByNoTransaksiLIKE(vKriteria + "%").Value
            ElseIf vSearchBy = Core_Module.EnumSearchCriteria.ByCode Then
                Return DTA.ScalarQuery_GetJmlByEmployeeCodeLIKE(vKriteria + "%").Value
            ElseIf vSearchBy = Core_Module.EnumSearchCriteria.ByDate Then
                Return DTA.ScalarQuery_GetJmlByTanggalDariSampai(vEndDate, vStartDate).Value
            ElseIf vSearchBy = Core_Module.EnumSearchCriteria.ByRemarks Then
                Return DTA.ScalarQuery_GetJmlByRemarksLIKE(vKriteria + "%").Value
            ElseIf vSearchBy = Core_Module.EnumSearchCriteria.ByEmployeeName Then
                Return DTA.ScalarQuery_GetJmlByEmployeeName(vKriteria + "%").Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseTransaksiSalary.DataTable8DataTable, _
    ByVal SearchCriteria As Core_Module.EnumSearchCriteria, ByVal tsl As ToolStripLabel, _
    Optional ByVal vEndDate As Date = #1/1/1950#, Optional ByVal vStartDate As Date = #1/1/1950#) As Boolean
        Dim DTA As New DseTransaksiSalaryTableAdapters.DataTable8TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            DGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            If SearchCriteria = Core_Module.EnumSearchCriteria.ByNoTransaksi Then
                QueryResult = DTA.FillBy_NoTransaksiLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = Core_Module.EnumSearchCriteria.ByDate Then
                QueryResult = DTA.FillBy_TanggalDariSampai(TBL, vEndDate, vStartDate)
            ElseIf SearchCriteria = Core_Module.EnumSearchCriteria.ByCode Then
                QueryResult = DTA.FillBy_EmployeeCodeLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = Core_Module.EnumSearchCriteria.ByRemarks Then
                QueryResult = DTA.FillBy_RemarksLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = Core_Module.EnumSearchCriteria.ByEmployeeName Then
                QueryResult = DTA.FillBy_EmployeeNameLIKE(TBL, Criteria + "%")
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(0).Frozen = True
                DGV.Columns(1).Frozen = True
                DGV.Columns(2).DefaultCellStyle.Format = "dd MMM yyyy"
                'DGV.Columns(1).HeaderText = "Code"
                'DGV.Columns(2).HeaderText = "Description"
                'DGV.Columns(3).HeaderText = "City"
                'DGV.Columns(4).HeaderText = "Remarks"
            End If


            Dim TS As TimeSpan = Now - StartT
            If QueryResult > 0 Then
                If TS.Ticks > TimeSpan.TicksPerSecond Then
                    tsl.Text = CStr(QueryResult) + " result(s) in " + CStr(TS.Seconds) + " seconds"
                Else
                    tsl.Text = CStr(QueryResult) + " result(s) in " + CStr(Math.Round(TS.Ticks / TimeSpan.TicksPerMillisecond, _
                    0, MidpointRounding.AwayFromZero)) + " miliseconds"
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

    Public Overridable Function SimpanBaru(ByVal vUserCode As String, ByVal vObjCD As Core_Module.ClsCekData, _
    ByVal vPrimary As Double, ByVal vdgvItem As DataGridView, ByVal vdgvAttd As DataGridView, _
    ByVal vDGVOtherDeduction As DataGridView, ByVal vDGVPayment As DataGridView, ByVal vLoanPayment_No As String, ByVal ij As FrmTrans_Gaji, _
    ByVal ijk As FrmTransaksiLoan_Payment, ByVal ikl As FrmTrans_Gaji_OtherDeduction, ByVal ijn As FrmTrans_Gaji_Payment, _
    ByVal vTBLItemIterasi As DseTransaksiSalary.tbltrans_salary_item_iterationDataTable, _
    ByVal vTBLItemTax As DseTransaksiSalary.tbltrans_salary_itemtaxDataTable, _
    ByVal vTBLOtherPayment As DseTransaksiSalary.tbltrans_salary_item_other1DataTable, ByVal vPenghasilanBruto As Decimal, _
    ByVal vTBLPrevIncome As DseTransaksiSalary.tbltrans_salary_item_prev_income1DataTable) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch
        Try
            SW.Start()

            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
            Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_Salary_SimpanBaru, vUserCode)
            'Me.GetNoAntrian()

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= Me.ObjAntrian.GetWaitTime Then
                        Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanBaru1(vUserCode, vObjCD, vPrimary, vdgvItem, vdgvAttd, vDGVOtherDeduction, vDGVPayment, _
                            vLoanPayment_No, ij, ijk, ikl, ijn, vTBLItemIterasi, vTBLItemTax, vTBLOtherPayment, _
                            vPenghasilanBruto, vTBLPrevIncome)
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

    Public Overridable Function SimpanPerubahan(ByVal vUserCode As String, ByVal vObjCD As Core_Module.ClsCekData, _
    ByVal vPrimary As Double, ByVal vdgvItem As DataGridView, ByVal vDGVAttd As DataGridView, ByVal vDGVOtherDeduction As DataGridView, _
    ByVal vDGVPayment As DataGridView, ByVal ij As FrmTrans_Gaji, ByVal ijk As FrmTransaksiLoan_Payment, ByVal ikl As FrmTrans_Gaji_OtherDeduction, _
    ByVal ijn As FrmTrans_Gaji_Payment, ByVal vTBLItemIterasi As DseTransaksiSalary.tbltrans_salary_item_iterationDataTable, _
    ByVal vTBLItemTax As DseTransaksiSalary.tbltrans_salary_itemtaxDataTable, ByVal vTBLOtherPayment As DseTransaksiSalary.tbltrans_salary_item_other1DataTable, _
    ByVal vPenghasilanBruto As Decimal, ByVal vTBLPrevIncome As DseTransaksiSalary.tbltrans_salary_item_prev_income1DataTable) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch
        Try
            SW.Start()
            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
            Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_Salary_SimpanPerubahan, vUserCode)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= ObjAntrian.GetWaitTime Then
                        Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanPerubahan1(vUserCode, vPrimary, vObjCD, vdgvItem, vDGVAttd, _
                            vDGVOtherDeduction, vDGVPayment, ij, ijk, ikl, ijn, vTBLItemIterasi, vTBLItemTax, vTBLOtherPayment, vPenghasilanBruto, _
                            vTBLPrevIncome)
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

    Public Function Hapus(ByVal vUserCode As String, ByVal vObjCD As Core_Module.ClsCekData, _
    ByVal vPrimary As Double, ByVal vdgvItem As DataGridView, ByVal ij As FrmTrans_Gaji, ByVal ijk As FrmTransaksiLoan_Payment, _
    Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim ObjAntri As New Core_Module.ClsAntrian
        Dim SW As New Stopwatch
        Try
            SW.Start()

            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, _
            Me.vPrimary_Antrian, vPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_Salary_Hapus, vUserCode)

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
                        If CInt(Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value) = CInt(Me.vPrimary_Antrian) Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.Hapus1(vPrimary, vUserCode, True, TampilkanPesan, vdgvItem, ij, ijk)
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

    Private Sub update_PenghasilanBruto(ByVal vPrimary As Double, ByVal vPenghasilanBruto As Decimal, _
    ByVal DTA_Baru As DseTransaksiSalaryTableAdapters.tbltrans_salary2TableAdapter)
        Try
            'tidak setting constring lagi karena sudah dilakukan pada fungsi sebelumnya
            DTA_Baru.UpdateQuery_SetPenghasilanBruto(vPenghasilanBruto, vPrimary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function SimpanBaru1(ByVal vUserCode As String, _
    ByVal vObjCD As Core_Module.ClsCekData, ByVal vPrimary As Double, _
    ByVal vDGVItem As DataGridView, ByVal vDGVAttd As DataGridView, _
    ByVal vDGVOtherDeduction As DataGridView, ByVal vDGVPayment As DataGridView, _
    ByVal vLoanPayment_No As String, ByVal i As FrmTrans_Gaji, _
    ByVal ijk As FrmTransaksiLoan_Payment, ByVal ikl As FrmTrans_Gaji_OtherDeduction, _
    ByVal ijn As FrmTrans_Gaji_Payment, ByVal vTBLItemIterasi As DseTransaksiSalary.tbltrans_salary_item_iterationDataTable, _
    ByVal vTBLItemTax As DseTransaksiSalary.tbltrans_salary_itemtaxDataTable, ByVal vTBLOtherPayment As DseTransaksiSalary.tbltrans_salary_item_other1DataTable, _
    ByVal vPenghasilanBruto As Decimal, ByVal vTBLPrevIncome As DseTransaksiSalary.tbltrans_salary_item_prev_income1DataTable) As Boolean
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salaryTableAdapter
        Dim DTA_Tax As New DseTransaksiSalaryTableAdapters.tbltrans_salary_hitungpph21TableAdapter
        Dim DTA_Baru As New DseTransaksiSalaryTableAdapters.tbltrans_salary2TableAdapter
        Dim ObjItem As New ClsTransaksiSalary_Item
        Dim ObjLoanPay As New ClsTransaksiLoanPayment
        Dim ObjCust As New CustomSMESS.ClsTransaksiSalary

        Try
            vTBLItemIterasi.Rows.Clear()
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA_Baru.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahDataBenar(True, vPrimary, True, vObjCD, vDGVItem, vDGVAttd, i) = True Then
                Me.mUrutan = CDbl(Me.mNoTransaksi.Substring(0, 6))
                If Me.mA_NoForm1721A1 <> "" Then
                    Me.mA_UrutanForm1721A1 = CDbl(Me.mA_NoForm1721A1.Substring(10))
                Else
                    Me.mA_UrutanForm1721A1 = 0
                End If
                'Get Primary Loan Payment, karena -1
                If ijk.FungsiBtnOK_Salary() = False Then
                    Return False
                End If
                Me.mPrimaryLoanPayment = ObjLoanPay.GetPrimary(vLoanPayment_No)
                'Insert Data
                If DTA_Baru.InsertQuery(Me.mNoTransaksi, Me.mTanggalTransaksi, Me.mUrutan, Me.mStartDate, _
                Me.mEndDate, Me.mPrimaryEmployee, Me.mPrimarySalaryFormula, Me.mPrimaryPosition, _
                Me.mTax, Me.PrimaryTaxExpense, Me.mIncome, Me.mDeduction, Me.mPrimaryLoanPayment, _
                Me.mLoanInstallment, Me.mIncomeTax, Me.mSalaryReceived, Me.mPaymentStatus, _
                Me.mPaymentDate, Me.mRemarks, vUserCode, Now, vUserCode, Now, Now, _
                Me.mPeriod_Paid, Me.m01Gaji, Me.m02TunjanganPPh, Me.m03Tunjanganlainnya, Me.m04Honorarium, _
                Me.m05PremiAsuransi, Me.m06Natura, Me.m07Tantiem, Me.m08PenghasilanBruto, Me.m09BiayaJabatan, _
                Me.m10IuranPensiun, Me.m11JumlahPengurangan, Me.m12JumlahPenghasilanBruto, Me.m13JumlahPenghasilanMsPjkSblm, _
                Me.m14JumlahPenghasilanNettoUtkPPh, Me.m15PTKP, Me.m16PKPSetahun, Me.m17PPh21, Me.m18PPh21YgDiPotongMsPjkSblm, _
                Me.m19PPhTerhutang, Me.m20PPhPs21danPPh26Lunas, Me.mA_NoForm1721A1, Me.mA_NPWP_Pemotongan, _
                Me.mA_Nama_Pemotongan, Me.mA_NPWP, Me.mA_NIK, Me.mA_Nama, Me.mA_Alamat1, Me.mA_Alamat2, _
                Me.mA_Jenis_Kelamin, Me.mA_Status_NPWP, Me.mA_Nama_Jabatan, Me.mA_Karyawan_Asing, _
                Me.mA_Kode_Negara_Domisili, Me.mC_NPWP, Me.mC_Nama, Me.mA_UrutanForm1721A1, Me.mPrimaryPayment, _
                Me.mGrossSalary, Me.mOtherDeduction, Me.mKurangBayarTax, Me.mCompanyExpense, Me.mEmployeeExpense, _
                Me.m01GajiRutin, Me.m02TunjanganPPHNonRutin, Me.m02TunjanganPPHRutin, Me.m02TunjanganPPHDisetahunkan, Me.m02TunjanganPPHDisetahunkanRutin, _
                Me.m03TunjanganLainnyaRutin, Me.m04HonorariumRutin, Me.m05PremiAsuransiRutin, Me.m06NaturaRutin, Me.m07TantiemRutin, _
                Me.m08PenghasilanBrutoRutin, Me.m09BiayaJabatanRutin, Me.m10IuranPensiunRutin, Me.m11JumlahPenguranganRutin, _
                Me.m14JumlahPenghasilanNettoUtkPPhRutin, Me.m16PKPSetahunRutin, Me.m17PPh21Rutin, _
                Me.m19PPhTerhutangSetahun, Me.m19PPhTerhutangSetahunRutin, Me.m19PPhTerhutangSebulanNonRutin, _
                Me.m19PPhTerhutangSebulanRutin, Me.m21Pesangon, Me.m21TunjPPhFinalPesangon, Me.m21PPhFinalPesangon, Me.m21TarifPPhFinalPesangon, _
                Me.mPeriodStart, Me.mPeriodEnd, Me.mPeriodCount, Me.mPrimaryDept, Me.mIsResign, Me.mLosesTaxObligation, Me._13JumlahPenghasilanMsPjkSblm_1, Me._18PPh21YgDiPotongMsPjkSblm_1) > 0 Then
                    Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mNoTransaksi)
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    'Update data primary salary di Transaksi Loan Payment
                    ObjLoanPay.UpdatePrimarySalary(Me.mPrimary, Me.mPrimaryLoanPayment)

                    'simpan data item
                    ObjItem.SimpanItem(Me.mPrimary, vDGVItem, i, vTBLItemIterasi, vTBLItemTax, vTBLOtherPayment, vTBLPrevIncome)
                    ObjItem.SimpanItemOtherDeduction(Me.mPrimary, vDGVOtherDeduction, ikl)
                    ObjItem.SimpanItemPayment(Me.mPrimary, vDGVPayment, ijn)

                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    'Simpan data Perhitungan PPh 21
                    If Me.mTax = EnumTax.Gross_Up Then
                        Dim vApakahAkhirTahun As ClsTransaksiSalary.EnumHitungPPh21AkhirAtauBukan
                        If Me.mEndDate.Month = 12 Then
                            vApakahAkhirTahun = EnumHitungPPh21AkhirAtauBukan.AkhirTahun
                        Else
                            vApakahAkhirTahun = EnumHitungPPh21AkhirAtauBukan.BukanAkhirTahun
                        End If
                        If Me.PPh21_GetJumlah(CDec(Me.m08PenghasilanBruto) - CDec(Me.m02TunjanganPPh), _
                        Me.m11JumlahPengurangan, Me.m15PTKP, vApakahAkhirTahun) <= 0 Then
                            DTA_Tax.Insert(CDec(Me.m08PenghasilanBruto) - CDec(Me.m02TunjanganPPh), _
                            Me.m11JumlahPengurangan, Me.m15PTKP, Me.m19PPhTerhutang, vApakahAkhirTahun, Me.m17PPh21)
                        End If
                    End If

                    Me.update_PenghasilanBruto(Me.mPrimary, vPenghasilanBruto, DTA_Baru)

                    'simpan item detail pada custom
                    ObjCust.SimpanDetailOvertime_Custom(Me.mPrimary)

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
            DTA_Baru.Dispose()
            DTA_Tax.Dispose()
            ObjItem = Nothing
            ObjLoanPay = Nothing
            ObjCust = Nothing
        End Try
    End Function

    Public Function GetPrimarySalaryFormula(ByVal vPrimarySalary As Double) As Double
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salaryTableAdapter
        Dim Dummy As New System.Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetPrimaryLoanPayment(vPrimarySalary)
            If Dummy.HasValue Then
                Return Dummy
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function SimpanPerubahan1(ByVal vUserCode As String, _
    ByVal vPrimary As Double, ByVal vObjCD As Core_Module.ClsCekData, _
    ByVal vDGVItem As DataGridView, ByVal vDGVAttd As DataGridView, _
    ByVal vDGVOtherDeduction As DataGridView, ByVal vDGVPayment As DataGridView, _
    ByVal i As FrmTrans_Gaji, ByVal ijk As FrmTransaksiLoan_Payment, _
    ByVal ikl As FrmTrans_Gaji_OtherDeduction, ByVal ijn As FrmTrans_Gaji_Payment, ByVal vTBLItemIterasi As DseTransaksiSalary.tbltrans_salary_item_iterationDataTable, _
    ByVal vTBLItemTax As DseTransaksiSalary.tbltrans_salary_itemtaxDataTable, ByVal vTBLOtherPayment As DseTransaksiSalary.tbltrans_salary_item_other1DataTable, _
    ByVal vPenghasilanBruto As Decimal, ByVal vTBLPrevIncome As DseTransaksiSalary.tbltrans_salary_item_prev_income1DataTable) As Boolean
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salaryTableAdapter
        Dim DTA_Baru As New DseTransaksiSalaryTableAdapters.tbltrans_salary2TableAdapter
        Dim DTA_Tax As New DseTransaksiSalaryTableAdapters.tbltrans_salary_hitungpph21TableAdapter
        Dim ObjItem As New ClsTransaksiSalary_Item
        Dim ObjCust As New CustomSMESS.ClsTransaksiSalary

        Try
            vTBLItemIterasi.Rows.Clear()
            If Me.ApakahDataBenar(False, vPrimary, True, vObjCD, vDGVItem, vDGVAttd, i) = True Then
                DTA.Connection.ConnectionString = Core_Module.GetConstring
                DTA_Baru.Connection.ConnectionString = Core_Module.GetConstring
                Me.mUrutan = CDbl(Me.mNoTransaksi.Substring(0, 6))
                If Me.mA_NoForm1721A1 <> "" Then
                    Me.mA_UrutanForm1721A1 = CDbl(Me.mA_NoForm1721A1.Substring(10))
                Else
                    Me.mA_UrutanForm1721A1 = 0
                End If
                'Update Loan Payment
                If ijk.FungsiBtnOK_Salary() = False Then
                    Return False
                End If
                If Me.GetPrimarySalaryFormula(Me.mPrimary) = -1 Then
                    Me.mPrimaryLoanPayment = -1
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                End If

                If DTA_Baru.UpdateQuery(Me.mNoTransaksi, Me.mTanggalTransaksi, Me.mUrutan, Me.mStartDate, _
                Me.mEndDate, Me.mPrimaryEmployee, Me.mPrimarySalaryFormula, Me.mPrimaryPosition, _
                Me.mTax, Me.PrimaryTaxExpense, Me.mIncome, Me.mDeduction, Me.mPrimaryLoanPayment, _
                Me.mLoanInstallment, Me.mIncomeTax, Me.mSalaryReceived, Me.mPaymentStatus, _
                Me.mPaymentDate, Me.mRemarks, vUserCode, Now, Now, Me.mPeriod_Paid, _
                Me.m01Gaji, Me.m02TunjanganPPh, Me.m03Tunjanganlainnya, Me.m04Honorarium, _
                Me.m05PremiAsuransi, Me.m06Natura, Me.m07Tantiem, Me.m08PenghasilanBruto, Me.m09BiayaJabatan, _
                Me.m10IuranPensiun, Me.m11JumlahPengurangan, Me.m12JumlahPenghasilanBruto, Me.m13JumlahPenghasilanMsPjkSblm, _
                Me.m14JumlahPenghasilanNettoUtkPPh, Me.m15PTKP, Me.m16PKPSetahun, Me.m17PPh21, Me.m18PPh21YgDiPotongMsPjkSblm, _
                Me.m19PPhTerhutang, Me.m20PPhPs21danPPh26Lunas, Me.mA_NoForm1721A1, Me.mA_NPWP_Pemotongan, _
                Me.mA_Nama_Pemotongan, Me.mA_NPWP, Me.mA_NIK, Me.mA_Nama, Me.mA_Alamat1, Me.mA_Alamat2, _
                Me.mA_Jenis_Kelamin, Me.mA_Status_NPWP, Me.mA_Nama_Jabatan, Me.mA_Karyawan_Asing, _
                Me.mA_Kode_Negara_Domisili, Me.mC_NPWP, Me.mC_Nama, Me.mA_UrutanForm1721A1, _
                Me.mPrimaryPayment, Me.mGrossSalary, Me.mOtherDeduction, Me.mKurangBayarTax, _
                Me.mCompanyExpense, Me.mEmployeeExpense, Me.m01GajiRutin, Me.m02TunjanganPPHNonRutin, _
                Me.m02TunjanganPPHRutin, Me.m02TunjanganPPHDisetahunkan, Me.m02TunjanganPPHDisetahunkanRutin, _
                Me.m03TunjanganLainnyaRutin, Me.m04HonorariumRutin, Me.m05PremiAsuransiRutin, Me.m06NaturaRutin, _
                Me.m07TantiemRutin, Me.m08PenghasilanBrutoRutin, Me.m09BiayaJabatanRutin, _
                Me.m10IuranPensiunRutin, Me.m11JumlahPenguranganRutin, Me.m14JumlahPenghasilanNettoUtkPPhRutin, _
                Me.m16PKPSetahunRutin, Me.m17PPh21Rutin, Me.m19PPhTerhutangSetahun, _
                Me.m19PPhTerhutangSetahunRutin, Me.m19PPhTerhutangSebulanNonRutin, _
                Me.m19PPhTerhutangSebulanRutin, Me.m21Pesangon, Me.m21TunjPPhFinalPesangon, Me.m21PPhFinalPesangon, Me.m21TarifPPhFinalPesangon, _
                Me.mPeriodStart, Me.mPeriodEnd, Me.mPeriodCount, Me.mPrimaryDept, Me.mIsResign, Me.mLosesTaxObligation, _
                Me._13JumlahPenghasilanMsPjkSblm_1, Me._18PPh21YgDiPotongMsPjkSblm_1, Me.mPrimary, Me.mTimestamp) > 0 Then
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode
                    'hapus data item lama
                    ObjItem.HapusItem(Me.mPrimary)
                    ObjItem.HapusItemOtherDeduction(Me.mPrimary)
                    ObjItem.HapusItemPayment(Me.mPrimary)

                    'simpan data item
                    ObjItem.SimpanItem(Me.mPrimary, vDGVItem, i, vTBLItemIterasi, vTBLItemTax, vTBLOtherPayment, vTBLPrevIncome)
                    ObjItem.SimpanItemOtherDeduction(Me.mPrimary, vDGVOtherDeduction, ikl)
                    ObjItem.SimpanItemPayment(Me.mPrimary, vDGVPayment, ijn)

                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode

                    'Simpan data Perhitungan PPh 21
                    If Me.mTax = EnumTax.Gross_Up Then
                        Dim vApakahAkhirTahun As ClsTransaksiSalary.EnumHitungPPh21AkhirAtauBukan
                        If Me.mEndDate.Month = 12 Then
                            vApakahAkhirTahun = EnumHitungPPh21AkhirAtauBukan.AkhirTahun
                        Else
                            vApakahAkhirTahun = EnumHitungPPh21AkhirAtauBukan.BukanAkhirTahun
                        End If
                        If Me.PPh21_GetJumlah(CDec(Me.m08PenghasilanBruto) - CDec(Me.m02TunjanganPPh), _
                        Me.m11JumlahPengurangan, Me.m15PTKP, vApakahAkhirTahun) <= 0 Then
                            DTA_Tax.Insert(CDec(Me.m08PenghasilanBruto) - CDec(Me.m02TunjanganPPh), _
                            Me.m11JumlahPengurangan, Me.m15PTKP, Me.m19PPhTerhutang, vApakahAkhirTahun, Me.m17PPh21)
                        End If
                    End If

                    Me.update_PenghasilanBruto(Me.mPrimary, vPenghasilanBruto, DTA_Baru)

                    'hapus item detail pada custom
                    ObjCust.HapusDetailOvertime_Custom(Me.mPrimary)
                    'simpan item detail pada custom
                    ObjCust.SimpanDetailOvertime_Custom(Me.mPrimary)

                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA_Baru.Dispose()
            DTA_Tax.Dispose()
            ObjItem = Nothing
            ObjCust = Nothing
        End Try
    End Function

    Public Function Hapus1(ByVal vPrimary As Double, ByVal vUserCode As String, _
    ByVal vStrictDelete As Boolean, ByVal vShowMsg As Boolean, _
    ByVal vDGVItem As DataGridView, ByVal i As FrmTrans_Gaji, _
    ByVal ijk As FrmTransaksiLoan_Payment) As Boolean
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salaryTableAdapter
        Dim ObjItem As New ClsTransaksiSalary_Item
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahBolehHapus(vPrimary, vUserCode, vStrictDelete, vShowMsg, vDGVItem, i) = True Then
                'Delete Loan Payment 
                If ijk.FungsiBtnHapus_Salary() = False Then
                    Return False
                End If
                'Hapus item
                ObjItem.HapusItem(vPrimary)
                ObjItem.HapusItemOtherDeduction(vPrimary)
                ObjItem.HapusItemPayment(vPrimary)
                'Hapus
                DTA.Delete(vPrimary)
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Sub ResetObject(ByVal vDgvItem As DataGridView, ByVal vDGVOtherDeduction As DataGridView)
        Try
            mPrimary = -1
            mNoTransaksi = ""
            mTanggalTransaksi = #1/1/1950#
            mUrutan = 0
            mStartDate = #1/1/1950#
            mEndDate = #1/1/1950#
            mPrimaryEmployee = -1
            mPrimarySalaryFormula = -1
            mPrimaryPosition = -1
            mTax = EnumTax.Gross_Up
            mPrimaryTaxExpense = -1
            mIncome = 0
            mDeduction = 0
            mPrimaryLoanPayment = -1
            mLoanInstallment = 0
            mIncomeTax = 0
            mSalaryReceived = 0
            mPaymentStatus = EnumPaymentStatus.unpaid
            mPaymentDate = #1/1/1950#
            mRemarks = ""
            '-------- Tidak dipakai, didalam versi update terakhir --------
            mAdj_IncomeAdjust = 0
            mAdj_PositionExpense = 0
            mAdj_PositionExpenseAdjust = 0
            mAdj_PTKPMonthly = 0
            mAdj_PKPValue = 0
            mAdj_PPhValue = 0
            '---------------------------------------------------------------
            mPeriod_Paid = EnumPeriodPaid.One_Month
            m01Gaji = 0
            m02TunjanganPPh = 0
            m03Tunjanganlainnya = 0
            m04Honorarium = 0
            m05PremiAsuransi = 0
            m06Natura = 0
            m07Tantiem = 0
            m08PenghasilanBruto = 0
            m09BiayaJabatan = 0
            m10IuranPensiun = 0
            m11JumlahPengurangan = 0
            m12JumlahPenghasilanBruto = 0
            m13JumlahPenghasilanMsPjkSblm = 0
            m14JumlahPenghasilanNettoUtkPPh = 0
            m15PTKP = 0
            m16PKPSetahun = 0
            m17PPh21 = 0
            m18PPh21YgDiPotongMsPjkSblm = 0
            m19PPhTerhutang = 0
            m20PPhPs21danPPh26Lunas = 0
            m21Pesangon = 0
            m21PPhFinalPesangon = 0
            m21TarifPPhFinalPesangon = 0
            m21TunjPPhFinalPesangon = 0
            mA_NoForm1721A1 = ""
            mA_NPWP_Pemotongan = ""
            mA_Nama_Pemotongan = ""
            mA_NPWP = ""
            mA_NIK = ""
            mA_Nama = ""
            mA_Alamat1 = ""
            mA_Alamat2 = ""
            mA_Jenis_Kelamin = ClsTransaksiSalary.EnumJenisKelamin.Wanita
            mA_Status_NPWP = Core_Module.ClsMasterKaryawan.GolonganPajak.TK0
            mA_Nama_Jabatan = 0
            mA_Karyawan_Asing = ""
            mA_Kode_Negara_Domisili = ""
            mC_NPWP = ""
            mC_Nama = ""
            mA_UrutanForm1721A1 = 0
            mPrimaryPayment = -1
            mGrossSalary = 0
            mOtherDeduction = 0
            mKurangBayarTax = 0

            mCreatedBy = ""
            mCreatedDate = Today
            mModifiedBy = ""
            mModifiedDate = Today
            mTimestamp = Today

            'Item
            vDgvItem.Rows.Clear()
            vDGVOtherDeduction.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function ApakahBolehEdit(ByVal vPrimary As Double, ByVal vUserCode As String, _
      ByVal vStrictEdit As Boolean, ByVal vShowMsg As Boolean) As Boolean
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salaryTableAdapter
        Dim pCreatedBy As String = ""
        Dim PrimaryKaryawan1 As Double = -1
        Dim ObjKaryawan As New Core_Module.ClsMasterKaryawan
        Dim TglResign As Date = #1/1/1950#
        Dim PrimarySalaryRef As Double = -1
        Dim ObjTR As New ClsTrans_TaxRecalculate
        Dim NoTR As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            '1. tidak boleh hapus jika melanggar aturan strict edit
            If vStrictEdit = True Then
                If IsNothing(DTA.ScalarQuery_GetCreatedBy(vPrimary)) = False Then
                    pCreatedBy = DTA.ScalarQuery_GetCreatedBy(vPrimary)
                Else
                    pCreatedBy = ""
                End If

                If pCreatedBy <> vUserCode Then
                    If vShowMsg = True Then
                        MsgBox("Only " + pCreatedBy + " can edit this data.", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If
            '2. Berkaitan dengan transaksi salary Payment
            'MsgBox("Berkaitan dengan batasan edit salary payment masih pending", MsgBoxStyle.Information, "Informasi")
            If Me.mPaymentStatus = EnumPaymentStatus.paid Then
                MsgBox("Can't edit this transaction because it is referred by salary payment.", MsgBoxStyle.Information, "Infomation")
                Return False
            End If

            '3. Batasan bulan 1 sampai 11, jika sudah ada transaksi di bulan 12, maka tidak boleh edit, dan hapus salary
            If Me.mTanggalTransaksi.Month <> 12 Then
                If Me.ApakahSelamaBulan1sampai11TransaksiBerjalanUdahAdaTransaksiBulan12a(Me.mTanggalTransaksi, Me.mPrimaryEmployee, vUserCode) = False Then
                    Return False
                End If
            End If

            'jika ada sudah resign dan perhitungan gaji untuk bulan resign sudah dilakukan
            PrimaryKaryawan1 = Me.GetPrimaryEmployeeByPrimary(vPrimary)
            If ObjKaryawan.ApakahAktif(PrimaryKaryawan1) = False Then
                TglResign = ObjKaryawan.GetResignDate(PrimaryKaryawan1)
                PrimarySalaryRef = Me.GetPrimary_ByPeriod(TglResign, PrimaryKaryawan1)
                If PrimarySalaryRef <> -1 And vPrimary <> PrimarySalaryRef Then
                    MsgBox("This employee is not active, the last salary calculation is: " + Me.GetNoTransaksi(PrimarySalaryRef) + _
                    ". You can't edit this data.", MsgBoxStyle.Information, "Informtion")
                    Return False
                End If
            End If

            'jika sudah di recalculate, maka tidak bisa edit
            NoTR = ObjTR.GetNoTransaksi_ByPrimarySalary(vPrimary)
            If NoTR <> "" Then
                MsgBox("This transaction is refered by tax recalculation: " + NoTR + ", you can't edit this data.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjKaryawan = Nothing
            ObjTR = Nothing
        End Try
    End Function

    'Public Function ApakahBolehHapus1(ByVal vPrimary As Double, ByVal vUserCode As String, _
    'ByVal vStrictDelete As Boolean, ByVal vShowMsg As Boolean, _
    'ByVal vDGVItem As DataGridView, ByVal i As FrmTrans_Gaji) As Boolean
    '    Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salaryTableAdapter
    '    Dim pCreatedBy As String = ""
    '    Dim ObjItem As New ClsTransaksiLoanPayment_Item
    '    Dim InfoData As String = ""
    '    Dim TBL As New DseTransaksiSalary.tbltrans_salaryDataTable
    '    Dim row As DseTransaksiSalary.tbltrans_salaryRow
    '    Try
    '        DTA.Connection.ConnectionString = Core_Module.GetConstring
    '        '1. tidak boleh hapus jika melanggar aturan strict edit
    '        If vStrictDelete = True Then
    '            If IsNothing(DTA.ScalarQuery_GetCreatedBy(vPrimary)) = False Then
    '                pCreatedBy = DTA.ScalarQuery_GetCreatedBy(vPrimary)
    '            Else
    '                pCreatedBy = ""
    '            End If
    '            If pCreatedBy <> vUserCode Then
    '                If vShowMsg = True Then
    '                    MsgBox("User that can delete this data is: " + pCreatedBy, _
    '                    MsgBoxStyle.Information, "Information")
    '                End If
    '                Return False
    '            End If
    '        End If

    '        '2. Berkaitan dengan transaksi salary Payment
    '        'MsgBox("Berkaitan dengan batasan edit salary payment masih pending", MsgBoxStyle.Information, "Informasi")
    '        DTA.FillBy_Primary(TBL, vPrimary)
    '        If TBL.Rows.Count > 0 Then
    '            row = TBL.Rows(0)
    '            If row.Payment_Status = EnumPaymentStatus.paid Then
    '                MsgBox("This transaction can't deleted because it is referring to transaction salary payment.", MsgBoxStyle.Information, "Infomation")
    '                Return False
    '            End If

    '            '3. Batasan bulan 1 sampai 11, jika sudah ada transaksi di bulan 12, maka tidak boleh edit, dan hapus salary
    '            If row.Tanggal_Transaksi.Month <> 12 Then
    '                If Me.ApakahSelamaBulan1sampai11TransaksiBerjalanUdahAdaTransaksiBulan12(row.Tanggal_Transaksi, row.Primary_Employee) = False Then
    '                    Return False
    '                End If
    '            End If
    '        End If

    '        Return True
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTA.Dispose()
    '        ObjItem = Nothing
    '        TBL.Dispose()
    '    End Try
    'End Function

    Public Function ApakahBolehHapus(ByVal vPrimary As Double, ByVal vUserCode As String, _
    ByVal vStrictDelete As Boolean, ByVal vShowMsg As Boolean, _
    ByVal vDGVItem As DataGridView, ByVal i As FrmTrans_Gaji) As Boolean
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salaryTableAdapter
        Dim pCreatedBy As String = ""
        Dim ObjItem As New ClsTransaksiLoanPayment_Item
        Dim InfoData As String = ""
        Dim PrimaryKaryawan1 As Double = -1
        Dim ObjKaryawan As New Core_Module.ClsMasterKaryawan
        Dim TglResign As Date = #1/1/1950#
        Dim PrimarySalaryRef As Double = -1
        Dim ObjTR As New ClsTrans_TaxRecalculate
        Dim NoTR As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            '1. tidak boleh hapus jika melanggar aturan strict edit
            If vStrictDelete = True Then
                If IsNothing(DTA.ScalarQuery_GetCreatedBy(vPrimary)) = False Then
                    pCreatedBy = DTA.ScalarQuery_GetCreatedBy(vPrimary)
                Else
                    pCreatedBy = ""
                End If
                If pCreatedBy <> vUserCode Then
                    If vShowMsg = True Then
                        MsgBox("User that can delete this data is: " + pCreatedBy, _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If

            '2. Berkaitan dengan transaksi salary Payment
            'MsgBox("Berkaitan dengan batasan edit salary payment masih pending", MsgBoxStyle.Information, "Informasi")
            If Me.mPaymentStatus = EnumPaymentStatus.paid Then
                MsgBox("Can't delete this transaction because it is referenced by salary payment.", MsgBoxStyle.Information, "Infomation")
                Return False
            End If

            '3. Batasan bulan 1 sampai 11, jika sudah ada transaksi di bulan 12, maka tidak boleh edit, dan hapus salary
            If Me.mTanggalTransaksi.Month <> 12 Then
                If Me.ApakahSelamaBulan1sampai11TransaksiBerjalanUdahAdaTransaksiBulan12a(Me.mTanggalTransaksi, Me.mPrimaryEmployee, "") = False Then
                    Return False
                End If
            End If

            PrimaryKaryawan1 = Me.GetPrimaryEmployeeByPrimary(vPrimary)
            If ObjKaryawan.ApakahAktif(PrimaryKaryawan1) = False Then
                TglResign = ObjKaryawan.GetResignDate(PrimaryKaryawan1)
                PrimarySalaryRef = Me.GetPrimary_ByPeriod(TglResign, PrimaryKaryawan1)
                If PrimarySalaryRef <> -1 And vPrimary <> PrimarySalaryRef Then
                    MsgBox("This employee is not active, the last salary calculation is: " + Me.GetNoTransaksi(PrimarySalaryRef) + _
                    ". You can't delete this data.", MsgBoxStyle.Information, "Informtion")
                    Return False
                End If
            End If

            'jika sudah di recalculate, maka tidak bisa edit
            NoTR = ObjTR.GetNoTransaksi_ByPrimarySalary(vPrimary)
            If NoTR <> "" Then
                MsgBox("This transaction is referenced by tax recalculation: " + NoTR + ", you can't delete this data.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjItem = Nothing
            ObjKaryawan = Nothing
            ObjTR = Nothing
        End Try
    End Function

    Public Function GetJmlByTanggalDariSampai_BuatAntaraTaxDenganSalary(ByVal vEndDate As Date, ByVal vStartDate As Date, ByVal vPrimaryEmployee As Double) As Double
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salary1TableAdapter
        Dim Dummy As System.Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetTanggalDariSampai_123(vEndDate, vStartDate, vPrimaryEmployee)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahDataBenar(ByVal vBaru As Boolean, ByVal vPrimary As Double, _
        ByVal vShowMsg As Boolean, ByVal vObjCD As Core_Module.ClsCekData, _
        ByVal vDGVItem As DataGridView, ByVal vDGVAttd As DataGridView, ByVal i As FrmTrans_Gaji) As Boolean
        Dim ObjItem As New ClsTransaksiSalary_Item
        Dim ij As Integer = 0
        Dim ListString As New List(Of String)
        Dim TanggalStart As Date = i.dtpPeriod.Value
        Dim TanggalEnd As Date = i.DtpTo.Value
        Dim FrmAtt As New FrmTrans_Gaji_ShowAttendance
        Dim Str As String = ""
        Dim ObjClosing As New AccModule.ClsTransaksiClosing
        Dim ConventionDate As Date = ObjClosing.GetPayrollConvertionDate()
        Dim vFirstTaxPeriod As Date = #1/1/1950#
        Dim vJmlAcuan As Integer = 0

        Try
            vFirstTaxPeriod = ObjKary.GetFirstTaxPeriod(Me.mPrimaryEmployee)

            'data salah bila:
            '1. Sudah ada no transaksi yang sama sebelumnya (untuk transaksi baru dan edit BEDA)
            '       untuk transaksi baru - dan autonumber-langsung ambil nomor baru
            '       untuk transaksi edit - dan autonumber-dianggap sama karena memang no transaksi tidak bisa diedit.
            If vBaru = True Then
                If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                    Me.mNoTransaksi = Me.GetNoBaru(Me.mTanggalTransaksi, vObjCD)
                End If
            Else
                Dim NoTransaksiLama As String = Me.GetNoTransaksi(Me.mPrimary)
                If NoTransaksiLama <> Me.mNoTransaksi Then
                    If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                        Me.mNoTransaksi = Me.GetNoBaru(Me.mTanggalTransaksi, vObjCD)
                    End If
                End If
            End If

            'yang harus diisikan:
            '1. kode employee code
            If Me.mPrimaryEmployee = "-1" Then
                If vShowMsg = True Then
                    MsgBox("Please input employee code.", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            '2. kode cash account
            'If Me.mPrimaryTaxExpense = "-1" Then
            '    If vShowMsg = True Then
            '        MsgBox("Please input tax expense account.", MsgBoxStyle.Information, "Information")
            '    End If
            '    Return False
            'End If
            '3. Remarks
            If vObjCD.CekDataKosong(Me.mRemarks, False) = False Then
                If vShowMsg = True Then
                    MsgBox("Data shouldn't include the character: ' , % , or * .", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            '4. Salary Received
            If vObjCD.CekNonNegatif(Me.mSalaryReceived) = False Then
                If vShowMsg = True Then
                    MsgBox("Salary received is not negative value.", MsgBoxStyle.Information, "Infomation")
                End If
                Return False
            End If
            '5. DGV tidak boleh kosong
            If i.dgvItem.Rows.Count <= 0 Then
                If vShowMsg = True Then
                    MsgBox("Please insert data item.", MsgBoxStyle.Information, "Infomation")
                End If
                Return False
            End If
            '6. batasan dengan Tax beginning balance 
            If Me.mEndDate.Month = 12 And Me.mEndDate.Year = ConventionDate.Year Then
                If Me.mTax <> EnumTax.Manual_Tax_Calculation Then
                    Dim MonthConvertion As Integer = ConventionDate.Month - 1
                    Dim DateConvention As Date = ConventionDate
                    Dim Hasil_1 As MsgBoxResult
                    If Me.GetJmlByPeriodTaxBeginning(ConventionDate, vObjCD, mPrimaryEmployee) <> MonthConvertion Then
                        Hasil_1 = MsgBox("This employee doesn't have tax beginning balance calculation for the following month : Jan - " _
                         + Format(DateConvention, "MMM").ToString + " Continue ?", MsgBoxStyle.YesNo, "Confirm")
                        If Hasil_1 = MsgBoxResult.No Then
                            Return False
                        End If
                    End If
                End If
            End If
            '7. Batasan dengan Tax Beginning Balance
            If Me.mEndDate.Month = 12 Then
                If Me.mTax <> EnumTax.Manual_Tax_Calculation Then
                    Dim YearConvertion As Integer = ConventionDate.Year
                    If YearConvertion <> Me.mTanggalTransaksi.Year Then
                        Dim Hasil_2 As MsgBoxResult
                        Dim mFirstDateYear As Date = vObjCD.GetFisrtDateOfTheYear(Me.mTanggalTransaksi)
                        Dim mLastDateYear As Date = "#11/30/" + Me.mTanggalTransaksi.Year.ToString + "#"
                        If vFirstTaxPeriod.Year < Me.mEndDate.Year Then
                            If Me.GetJmlByTanggalDariSampai_BuatAntaraTaxDenganSalary(mLastDateYear, mFirstDateYear, mPrimaryEmployee) <> 11 Then
                                Hasil_2 = MsgBox("This employee doesn't have salary calculation for the following month : Jan - " _
                                + Format(CDate("#11/1/1950#"), "MMM").ToString + " Continue ?", MsgBoxStyle.YesNo, "Confirm")
                                If Hasil_2 = MsgBoxResult.No Then
                                    Return False
                                End If
                            End If
                        Else
                            If vFirstTaxPeriod.Year = Me.mEndDate.Year Then
                                If vFirstTaxPeriod.Month <> 12 Then
                                    vJmlAcuan = 12 - vFirstTaxPeriod.Month
                                    If Me.GetJmlByTanggalDariSampai_BuatAntaraTaxDenganSalary(mLastDateYear, vFirstTaxPeriod, mPrimaryEmployee) <> vJmlAcuan Then
                                        Hasil_2 = MsgBox("This employee doesn't have salary calculation for the following month : " + _
                                        Format(CDate("#" + vFirstTaxPeriod.Month.ToString + "/1/1950#"), "MMM").ToString + " - " _
                                        + Format(CDate("#11/1/1950#"), "MMM").ToString + " Continue ?", MsgBoxStyle.YesNo, "Confirm")
                                        If Hasil_2 = MsgBoxResult.No Then
                                            Return False
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        Dim MonthConvertion As Integer = ConventionDate.Month
                        Dim Hasil_2 As MsgBoxResult
                        Dim mFirstDateYear As Date = vObjCD.GetFisrtDateOfTheYear(Me.mTanggalTransaksi)
                        Dim mLastDateYear As Date = "#11/30/" + Me.mTanggalTransaksi.Year.ToString + "#"
                        If 11 - Me.GetJmlByPeriodTaxBeginning(ConventionDate, vObjCD, mPrimaryEmployee) <> Me.GetJmlByTanggalDariSampai_BuatAntaraTaxDenganSalary(mLastDateYear, mFirstDateYear, mPrimaryEmployee) Then
                            Hasil_2 = MsgBox("This employee doesn't have salary calculation for the following month : " + Format(CDate("#" + MonthConvertion.ToString + "/1/1950#"), "MMM") + " - " _
                            + Format(CDate("#11/1/1950#"), "MMM").ToString + " Continue ?", MsgBoxStyle.YesNo, "Confirm")
                            If Hasil_2 = MsgBoxResult.No Then
                                Return False
                            End If
                        End If
                    End If
                End If
            End If

            '8. batasan bulan 12, cek bulan 1-11 metode tax calculation harus sama
            'If Me.ApakahSelama1TahunTaxCalculationSama(Me.mTanggalTransaksi, Me.mTax, Me.mPrimaryEmployee, Me.mPrimary) = False Then
            '    If i.UserCode = "USER" Then
            '        'kalo disini, pasti di override pada formnya, jadi biarin lewat aj
            '    Else
            '        Return False
            '    End If
            'End If

            '9. Batasan salary dengan tax beginning balance, jika tax calulation = netto/gross maka tidak boleh menginput transaksi di bulan 12
            If Me.TanggalTransaksi.Month = 12 Then
                If Me.mTax = EnumTax.Gross Or Me.mTax = EnumTax.Netto Then
                    If Me.ApakahTunjanganPPhLebihBesarDariNol(Me.mTanggalTransaksi, Me.mPrimaryEmployee) = False Then
                        Return False
                    End If
                End If
            End If

            'kalo mpe sini ok
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjItem = Nothing
            ListString = Nothing
            ObjClosing = Nothing
            FrmAtt.Dispose()
        End Try
    End Function

    Public Function GetJmlByTanggalTransaksi(ByVal vEndDate As Date, ByVal vStartDate As Date) As Double
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salary1TableAdapter
        Dim Dummy As System.Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetJmlByTanggalDariSampai(vEndDate, vStartDate)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlByPeriodTaxBeginning(ByVal vConventionDate As Date, ByVal vObjCD As Core_Module.ClsCekData, _
    ByVal vPrimaryEmployee As Double) As Double
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_tax_beginning_balanceTableAdapter
        Dim Dummy As System.Nullable(Of Long)
        Dim EndDate As Date = vConventionDate - TimeSpan.FromDays(1)
        Dim StartDate As Date = vObjCD.GetFisrtDateOfTheYear(vConventionDate)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetJmlByPeriodTax(EndDate, StartDate, vPrimaryEmployee)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetNoForm1721A1(ByVal vPrimary As Double) As String
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salary2TableAdapter
        Dim Dummy As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetNoForm1721A1(vPrimary)
            If String.IsNullOrEmpty(Dummy) = False Then
                Return Dummy
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

    Public Function GetNoBaru_BuatNPWP(ByVal vTanggalTransaksi As Date, ByVal vObjCD As Core_Module.ClsCekData) As String
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salary2TableAdapter
        Dim TanggalStart, TanggalEnd As Date
        Dim NoUrut, Bulan As String
        Dim InOut As String = "1.1-"
        NoUrut = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'cari tanggal sampai dan dari
            TanggalEnd = vObjCD.GetLastDateOfTheYear(vTanggalTransaksi)
            TanggalStart = vObjCD.GetFisrtDateOfTheYear(vTanggalTransaksi)

            'cari no urut
            If DTA.ScalarQuery_GetMAXUrutanPPh21(TanggalEnd, TanggalStart).HasValue = True Then
                NoUrut = CStr(DTA.ScalarQuery_GetMAXUrutanPPh21(TanggalEnd, TanggalStart).Value + 1)
                If NoUrut <= 0 Then
                    NoUrut = 1
                End If
            Else
                NoUrut = 1
            End If


            'pastikan no urut terdiri atas 6 huruf
            While NoUrut.Length < 7
                NoUrut = "0" + NoUrut
            End While

            'bulan 
            Bulan = CStr(vTanggalTransaksi.Month)
            While Bulan.Length < 2
                Bulan = "0" + Bulan
            End While

            Return InOut + Bulan + "." + Format(vTanggalTransaksi, "yy") + "-" + NoUrut
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetNoBaru(ByVal vTanggalTransaksi As Date, ByVal vObjCD As Core_Module.ClsCekData) As String
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salaryTableAdapter
        Dim TanggalStart, TanggalEnd As Date
        Dim NoUrut, Bulan As String
        Dim InOut As String = "SLR"
        NoUrut = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'cari tanggal sampai dan dari
            TanggalEnd = vObjCD.GetLastDateOfTheYear(vTanggalTransaksi)
            TanggalStart = vObjCD.GetFisrtDateOfTheYear(vTanggalTransaksi)

            'cari no urut
            If DTA.ScalarQuery_GetMAXUrutan(TanggalEnd, TanggalStart).HasValue = True Then
                NoUrut = CStr(DTA.ScalarQuery_GetMAXUrutan(TanggalEnd, TanggalStart).Value + 1)
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
        End Try
    End Function

    Public Function GetTanggalBaru() As Date
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salaryTableAdapter
        Dim Dummy As System.Nullable(Of Date)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetMAX_EndDate()
            If Dummy.HasValue Then
                Return Dummy.Value.Date.AddDays(1)
            Else
                Return Today
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return Today
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahSudahAda(ByVal vNoTransaksi As String) As Boolean
        Try
            If Me.GetJmlByNoTransaksi(vNoTransaksi) > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
#End Region

    Public Function GetSalaryRecieved(ByVal vPrimary As Double) As Decimal
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salaryTableAdapter
        Dim Dummy As New System.Nullable(Of Decimal)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetSalaryRecieved(vPrimary)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetIncomeTax(ByVal vPrimary As Double) As Decimal
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salaryTableAdapter
        Dim Dummy As New System.Nullable(Of Decimal)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetIncomeTax(vPrimary)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetLoanInstallment(ByVal vPrimary As Double) As Decimal
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salaryTableAdapter
        Dim Dummy As New System.Nullable(Of Decimal)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetLoanInstallment(vPrimary)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTaxTypeCalculation(ByVal vPrimary As Double) As Integer
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salaryTableAdapter
        Dim Dummy As New System.Nullable(Of Integer)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetTaxCalculation(vPrimary)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub Cetak(ByVal vCRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal vPrimary As Double, _
    ByVal vDef_Prec As Double) ', Optional ByVal vCreatedFile As Boolean = False)
        Dim FormatDecimal As Integer = vDef_Prec
        Dim Report As New CrpTransaksiSalary
        Dim DS As New DseTransaksiSalary
        Dim DTA As New DseTransaksiSalaryTableAdapters.TblCetakTableAdapter
        Dim TBL As DseTransaksiSalary.TblCetakDataTable = DS.TblCetak
        Dim Row As DseTransaksiSalary.TblCetakRow
        Dim TanggalTransaksi As Date = #1/1/1950#
        Dim FullName As String = "-"
        Dim Counter As Integer = 0
        Dim LoanInstallment As Decimal = 0
        Dim IncomeTax As Decimal = 0
        Dim SalaryReceived As Decimal = 0
        Dim TaxCalculation As Integer = -1

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_Primary(TBL, vPrimary)
            While Counter <= TBL.Rows.Count - 1
                Row = TBL.Rows(Counter)
                If Row.Amount = 0 Then
                    TBL.Rows.Remove(Row)
                Else
                    If Row.Type = ClsTransaksiSalary_Item.EnumType.Deduction Then
                        Row.Amount = Row.Amount * -1
                    End If
                    TanggalTransaksi = Row.Tanggal_Transaksi
                    FullName = Row.Full_Name
                    Counter += 1
                End If
            End While
            'Loan Installment 
            LoanInstallment = Me.GetLoanInstallment(vPrimary)
            If LoanInstallment > 0 Then
                Row = TBL.NewRow
                Row.Primary = -1
                Row.No_Transaksi = "-"
                Row.Code = "-"
                Row.Full_Name = FullName
                Row.current_address = "-"
                Row.Description = "Loan Installment"
                Row.Amount = LoanInstallment * -1
                Row.Type = ClsTransaksiSalary_Item.EnumType.Deduction
                Row.Tanggal_Transaksi = TanggalTransaksi
                TBL.Rows.Add(Row)
            End If
            IncomeTax = Me.GetIncomeTax(vPrimary)
            TaxCalculation = Me.GetTaxTypeCalculation(vPrimary)
            If IncomeTax > 0 Then
                Row = TBL.NewRow
                Row.Primary = -1
                Row.No_Transaksi = "-"
                Row.Code = "-"
                Row.Full_Name = FullName
                Row.current_address = "-"
                Row.Description = "Income Tax"
                If TaxCalculation = ClsTransaksiSalary.EnumTax.Gross_Up Then
                    Row.Description += " (Gross Up)"
                Else
                    Row.Description += " (Netto)"
                End If
                Row.Amount = IncomeTax * -1
                Row.Type = ClsTransaksiSalary_Item.EnumType.Deduction
                Row.Tanggal_Transaksi = TanggalTransaksi
                TBL.Rows.Add(Row)
            End If
            SalaryReceived = Me.GetSalaryRecieved(vPrimary)

            Report.SetDataSource(DS)
            Report.SetParameterValue(Report.Parameter_Format_Decimal.ParameterFieldName, FormatDecimal)
            Report.SetParameterValue(Report.Parameter_Salary_Received.ParameterFieldName, SalaryReceived)

            vCRV.ReportSource = Report
            vCRV.Show()

            'If vCreatedFile = True Then
            '    Report.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, _
            '    Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\Salary Slip - " + _
            '    Me.ObjKary.GetFullNameByPrimary(Me.GetPrimaryEmployeeByPrimary(vPrimary)) + ".pdf")
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTA.Dispose()
        End Try
    End Sub

    Public Sub FillSalaryComponent_Item(ByVal TBL As DseTransaksiSalary.tblmaster_salary_formula_salarycomponentDataTable, _
    ByVal PrimaryMasterSalaryComponent As Double)

        Dim DTA As New DseTransaksiSalaryTableAdapters.tblmaster_salary_formula_salarycomponentTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            TBL.Rows.Clear()
            DTA.FillBy_PrimarySalaryFormula(TBL, PrimaryMasterSalaryComponent)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub FillSalaryComponent_Item_OtherDeduction(ByVal TBL As DseTransaksiSalary.tblmaster_salary_formula_otherdeductionDataTable, _
    ByVal PrimaryEmployee As Double)
        Dim DTA As New DseTransaksiSalaryTableAdapters.tblmaster_salary_formula_otherdeductionTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            TBL.Rows.Clear()
            DTA.Fill(TBL, PrimaryEmployee)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub FillSalaryComponent_DetailEmployee(ByVal TBL As DseTransaksiSalary.tblmaster_salary_formula_component_detailemployeeDataTable, _
    ByVal PrimaryKary As Double)
        Dim DTA As New DseTransaksiSalaryTableAdapters.tblmaster_salary_formula_component_detailemployeeTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            TBL.Rows.Clear()
            DTA.FillBy_PrimaryEmployee(TBL, PrimaryKary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function GetPPHTerutang_Rutin_NonRutin(ByVal vPrimaryKaryawan As Double, ByVal vStartDate As Date, ByVal vEndDate As Date) As Decimal
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salary2TableAdapter
        Dim Dummy As System.Nullable(Of Decimal)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetPPHTerutanRutin_NonRutin(vPrimaryKaryawan, vEndDate, vStartDate)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPPHSudahDipotongLunas(ByVal vPrimaryKaryawan As Double, ByVal vStartDate As Date, ByVal vEndDate As Date) As Decimal
        'Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salary2TableAdapter
        'Dim Dummy As System.Nullable(Of Decimal)

        'Try
        '    DTA.Connection.ConnectionString = Core_Module.GetConstring
        '    Dummy = DTA.ScalarQuery_GetPPHTerutanRutin_NonRutin(vPrimaryKaryawan, vEndDate, vStartDate)
        '    If Dummy.HasValue Then
        '        Return Dummy.Value
        '    Else
        '        Return 0
        '    End If
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'Finally
        '    DTA.Dispose()
        'End Try
    End Function

    Public Function GetPrimaryMasterSalaryFormula(ByVal vPrimaryEmployee As Double) As Double
        Dim DTA As New DseTransaksiSalaryTableAdapters.tblmaster_salary_formula_component_detailemployeeTableAdapter
        Dim Dummy As System.Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetPrimarySalaryFormulaByEmployee(vPrimaryEmployee)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsConstraint_OtherDeduction(ByVal vPrimaryOther As Double, _
    ByVal vPrimaryEmployee As Double, ByVal vTglDari As Date, ByVal vTglSampai As Date, _
    ByVal TBLParam1 As DseMasterParameter.DataTable1DataTable, ByVal TBLParam2 As DseMasterParameter.tbltrans_attendanceDataTable, _
    ByVal TBLWeekDay As DseMasterParameter.tbltrans_attendance1DataTable, ByVal TBLHoliday As DseMasterParameter.tbltrans_attendance1DataTable, _
    ByVal TblKaryawan As DseMasterParameter.tblmaster_karyawanDataTable, ByVal vObjParam As ClsMasterParameterGaji, _
    ByVal TBLConst As DseTransaksiSalary.tblmaster_salary_formula_otherdeduction_constraintDataTable, ByVal vObjKompetensi As Core_Module.ClsMasterKompetensi, _
    ByVal vObjKaryawan As Core_Module.ClsMasterKaryawan) As Boolean
        Dim RowConst As DseTransaksiSalary.tblmaster_salary_formula_otherdeduction_constraintRow
        Dim RV As Data.DataRowView
        Dim Filter As String = ""
        Dim Var1 As String = ""
        Dim Var2 As String = ""
        'Parameter Leave Code Constraint
        Dim objItemOtherDeduction As New Core_Module.ClsTrans_OtherDeduction_item
        Dim mPrimaryLeave As Double = -1
        Try
            Filter = "[Primary Transaksi] = " + vPrimaryOther.ToString + ""
            TBLConst.DefaultView.RowFilter = Filter
            If TBLConst.DefaultView.Count > 0 Then
                For Each RV In TBLConst.DefaultView
                    RowConst = RV.Row
                    'Pass Parameter Primary Leave ke Fungsi GetValue
                    mPrimaryLeave = -1
                    If RowConst.Primary_Parameter = 133 Then
                        'Get Leave Code
                        mPrimaryLeave = objItemOtherDeduction.GetRefValueConstraint_Others(RowConst.Primary_Transaksi, RowConst.Primary_Parameter)
                    End If
                    Var1 = vObjParam.GetValue(TBLParam1, TBLParam2, TBLWeekDay, TBLHoliday, TblKaryawan, vPrimaryEmployee, _
                    RowConst.Primary_Parameter, vTglDari, vTglSampai, RowConst.Type2, vObjKaryawan, mPrimaryLeave)
                    MsgBox(Var1)
                    Var2 = RowConst.Ref_Value
                    If vObjParam.IsConstraintBenar(Var1, Var2, RowConst.Compare, RowConst.Primary_Parameter, RowConst.Type2, vObjKompetensi) = False Then
                        Return False
                    End If
                Next
            End If
            Return True
            TBLConst.DefaultView.RowFilter = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objItemOtherDeduction = Nothing
        End Try
    End Function

    Public Function IsConstraint(ByVal vIdentifierFormula As String, _
    ByVal vPrimaryEmployee As Double, ByVal vTglDari As Date, ByVal vTglSampai As Date, _
    ByVal TBLParam1 As DseMasterParameter.DataTable1DataTable, ByVal TBLParam2 As DseMasterParameter.tbltrans_attendanceDataTable, _
    ByVal TBLWeekDay As DseMasterParameter.tbltrans_attendance1DataTable, ByVal TBLHoliday As DseMasterParameter.tbltrans_attendance1DataTable, _
    ByVal TblKaryawan As DseMasterParameter.tblmaster_karyawanDataTable, ByVal vObjParam As ClsMasterParameterGaji, _
    ByVal TBLConst As DseTransaksiSalary.tblmaster_salary_formula_salarycomponent_constraintDataTable, ByVal vObjKompetensi As Core_Module.ClsMasterKompetensi, _
    ByVal vObjKary As Core_Module.ClsMasterKaryawan) As Boolean
        Dim RowConst As DseTransaksiSalary.tblmaster_salary_formula_salarycomponent_constraintRow
        Dim RV As Data.DataRowView
        Dim Filter As String = ""
        Dim Var1 As String = ""
        Dim Var2 As String = ""
        'Parameter Leave Code Constraint
        Dim objItemSalaryFormulaItem As New Core_Module.ClsMasterSalaryFormula_item
        Dim mPrimaryLeave As Double = -1
        Try
            Filter = "[Identifier salary Component] = '" + vIdentifierFormula.ToString + "'"
            TBLConst.DefaultView.RowFilter = Filter
            If TBLConst.DefaultView.Count > 0 Then
                For Each RV In TBLConst.DefaultView
                    RowConst = RV.Row
                    'Pass Parameter Primary Leave ke Fungsi GetValue
                    mPrimaryLeave = -1
                    If RowConst.Primary_Parameter = 133 Then
                        'Get Primary Leave 
                        mPrimaryLeave = objItemSalaryFormulaItem.GetRefValueConstraint(RowConst.Identifier_Salary_Component, RowConst.Primary_Parameter)
                    End If
                    Var1 = vObjParam.GetValue(TBLParam1, TBLParam2, TBLWeekDay, TBLHoliday, TblKaryawan, vPrimaryEmployee, _
                    RowConst.Primary_Parameter, vTglDari, vTglSampai, RowConst.Type2, vObjKary, mPrimaryLeave)
                    Var2 = RowConst.Ref_Value
                    If vObjParam.IsConstraintBenar(Var1, Var2, RowConst.Compare, RowConst.Primary_Parameter, RowConst.Type2, vObjKompetensi) = False Then
                        Return False
                    End If
                Next
            End If
            Return True
            TBLConst.DefaultView.RowFilter = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objItemSalaryFormulaItem = Nothing
        End Try
    End Function

    Public Function TranslateFormulaMonthlyYearly_OtherDeduction(ByVal Rumus As String, ByVal vPrimaryOther As Double, _
    ByVal vPrimaryEmployee As Double, ByVal vTglDari As Date, ByVal vTglSampai As Date, _
    ByVal TBLParam1 As DseMasterParameter.DataTable1DataTable, ByVal TBLParam2 As DseMasterParameter.tbltrans_attendanceDataTable, _
    ByVal TBLWeekDay As DseMasterParameter.tbltrans_attendance1DataTable, ByVal TBLHoliday As DseMasterParameter.tbltrans_attendance1DataTable, _
    ByVal TblKaryawan As DseMasterParameter.tblmaster_karyawanDataTable, ByVal vObjParam As ClsMasterParameterGaji, ByVal vObjKary As Core_Module.ClsMasterKaryawan) As Decimal
        Dim Formula As New ExpressionContext
        Dim var As VariableCollection = Formula.Variables
        Dim DTA As New DseTransaksiSalaryTableAdapters.tblmaster_salary_formula_otherdeduction_variableTableAdapter
        Dim TBL As New DseTransaksiSalary.tblmaster_salary_formula_otherdeduction_variableDataTable
        Dim Row As DseTransaksiSalary.tblmaster_salary_formula_otherdeduction_variableRow
        Dim Hasil As Decimal = 0
        Dim Filter As String = ""
        Dim DV As Data.DataRowView
        Dim ObjMasterSF_Item As New Core_Module.ClsTrans_OtherDeduction_item
        Dim ObjCD As New Core_Module.ClsCekData
        Dim vDef_Prec As Decimal = 0
        'Parameter Total Leave Balance Yearly
        Dim objItemSalaryFormulaItem_OtherDeduction As New Core_Module.ClsTrans_OtherDeduction_item
        Dim mPrimaryLeave As Double = -1
        'Parameter Rounding untuk Translate Formula Other
        Dim mSIC_RoundingTipe As ClsTrans_OtherDeduction_item.EnumRoundingType = ClsTrans_OtherDeduction_item.EnumRoundingType.No_Rounding
        Dim mSIC_RoundingUnit As Decimal = 0

        Try
            Formula.ParserOptions.DecimalSeparator = "."
            Formula.ParserOptions.FunctionArgumentSeparator = ";"
            Formula.Options.ParseCulture.NumberFormat.NumberDecimalSeparator = "."
            Formula.Options.ParseCulture.NumberFormat.NumberGroupSeparator = ","
            Formula.ParserOptions.RecreateParser()
            Rumus = Rumus.Replace(",", "")

            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_PrimaryOther(TBL, vPrimaryOther)
            TBL.DefaultView.RowFilter = "[Primary Transaksi] = " + vPrimaryOther.ToString + ""
            If TBL.DefaultView.Count > 0 Then
                For Each DV In TBL.DefaultView
                    Row = DV.Row
                    'Pass Parameter Primary Leave ke Fungsi GetValue
                    mPrimaryLeave = -1
                    If Row.Primary_Parameter = 132 Then
                        'Total Leave Balance Yearly, Primary Param Leave Code = 133
                        mPrimaryLeave = objItemSalaryFormulaItem_OtherDeduction.GetRefValueConstraint_Others(Row.Primary_Transaksi, 133)
                    End If
                    If Row.Rounding_Type = ClsTrans_OtherDeduction_item.EnumRoundingType.No_Rounding Then
                        var.Add(Row.Variable.ToString, CDbl(vObjParam.GetValue(TBLParam1, TBLParam2, TBLWeekDay, _
                        TBLHoliday, TblKaryawan, vPrimaryEmployee, Row.Primary_Parameter, vTglDari, vTglSampai, Row.Type2, vObjKary, mPrimaryLeave)))
                    ElseIf Row.Rounding_Type = ClsTrans_OtherDeduction_item.EnumRoundingType.Round_Down Then
                        var.Add(Row.Variable.ToString, CDbl(ObjCD.GetRoundingValue(CDec(vObjParam.GetValue(TBLParam1, TBLParam2, TBLWeekDay, _
                        TBLHoliday, TblKaryawan, vPrimaryEmployee, Row.Primary_Parameter, vTglDari, vTglSampai, Row.Type2, vObjKary, mPrimaryLeave)), 0, ClsCekData.EnumTipeRounding.Down, Row.Rounding_Unit, ObjCD)))
                    ElseIf Row.Rounding_Type = ClsTrans_OtherDeduction_item.EnumRoundingType.Round_Middlepoint Then
                        var.Add(Row.Variable.ToString, CDbl(ObjCD.GetRoundingValue(CDec(vObjParam.GetValue(TBLParam1, TBLParam2, TBLWeekDay, _
                        TBLHoliday, TblKaryawan, vPrimaryEmployee, Row.Primary_Parameter, vTglDari, vTglSampai, Row.Type2, vObjKary, mPrimaryLeave)), 0, ClsCekData.EnumTipeRounding.Middlepoint, Row.Rounding_Unit, ObjCD)))
                    ElseIf Row.Rounding_Type = ClsTrans_OtherDeduction_item.EnumRoundingType.Round_Up Then
                        var.Add(Row.Variable.ToString, CDbl(ObjCD.GetRoundingValue(CDec(vObjParam.GetValue(TBLParam1, TBLParam2, TBLWeekDay, _
                        TBLHoliday, TblKaryawan, vPrimaryEmployee, Row.Primary_Parameter, vTglDari, vTglSampai, Row.Type2, vObjKary, mPrimaryLeave)), 0, ClsCekData.EnumTipeRounding.Up, Row.Rounding_Unit, ObjCD)))
                    End If
                Next

                'Result
                mSIC_RoundingTipe = objItemSalaryFormulaItem_OtherDeduction.Get_SCI_RoundingType(vPrimaryOther)
                mSIC_RoundingUnit = objItemSalaryFormulaItem_OtherDeduction.Get_SCIRoundingUnit(vPrimaryOther)
                If mSIC_RoundingTipe = ClsMasterSalaryFormula_item.EnumRoundingType.No_Rounding Then
                    Hasil = CDec(Formula.CompileDynamic(Rumus).Evaluate)
                ElseIf mSIC_RoundingTipe = ClsTrans_OtherDeduction_item.EnumRoundingType.Round_Down Then
                    Hasil = ObjCD.GetRoundingValue(CDec(Formula.CompileDynamic(Rumus).Evaluate), vDef_Prec, ClsCekData.EnumTipeRounding.Down, mSIC_RoundingUnit, ObjCD)
                ElseIf mSIC_RoundingTipe = ClsTrans_OtherDeduction_item.EnumRoundingType.Round_Middlepoint Then
                    Hasil = ObjCD.GetRoundingValue(CDec(Formula.CompileDynamic(Rumus).Evaluate), vDef_Prec, ClsCekData.EnumTipeRounding.Middlepoint, mSIC_RoundingUnit, ObjCD)
                ElseIf mSIC_RoundingTipe = ClsTrans_OtherDeduction_item.EnumRoundingType.Round_Up Then
                    Hasil = ObjCD.GetRoundingValue(CDec(Formula.CompileDynamic(Rumus).Evaluate), vDef_Prec, ClsCekData.EnumTipeRounding.Up, mSIC_RoundingUnit, ObjCD)
                End If
            End If
            Return Hasil
        Catch ex As Exception
            MsgBox(Me.GetDescriptionOtherDeduction(vPrimaryOther).ToString + " formula is incorrect, you cannot get result formula value.", MsgBoxStyle.Information, "Information")
            'MsgBox("Wrong formula, you cannot get result formula value.", MsgBoxStyle.Information, "Information")
            'MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjMasterSF_Item = Nothing
            objItemSalaryFormulaItem_OtherDeduction = Nothing
            ObjCD = Nothing
        End Try
    End Function

    Public Function TranslateFormula(ByVal Rumus As String, ByVal vPrimaryFormula As Double, _
    ByVal vPrimaryEmployee As Double, ByVal vTglDari As Date, ByVal vTglSampai As Date, _
    ByVal TBLParam1 As DseMasterParameter.DataTable1DataTable, ByVal TBLParam2 As DseMasterParameter.tbltrans_attendanceDataTable, _
    ByVal TBLWeekDay As DseMasterParameter.tbltrans_attendance1DataTable, ByVal TBLHoliday As DseMasterParameter.tbltrans_attendance1DataTable, _
    ByVal TblKaryawan As DseMasterParameter.tblmaster_karyawanDataTable, ByVal vObjParam As ClsMasterParameterGaji, _
    ByVal vIdentifierFormula As String, ByVal vObjKary As Core_Module.ClsMasterKaryawan) As Decimal
        Dim Formula As New ExpressionContext
        Dim var As VariableCollection = Formula.Variables
        Dim DTA As New DseTransaksiSalaryTableAdapters.tblmaster_salary_formula_salarycomponent_variableTableAdapter
        Dim TBL As New DseTransaksiSalary.tblmaster_salary_formula_salarycomponent_variableDataTable
        Dim Row As DseTransaksiSalary.tblmaster_salary_formula_salarycomponent_variableRow
        Dim Hasil As Decimal = 0
        Dim Filter As String = ""
        Dim DV As Data.DataRowView
        Dim ObjMasterSF_Item As New Core_Module.ClsMasterSalaryFormula_item
        Dim ObjCD As New Core_Module.ClsCekData
        Dim vDef_Prec As Decimal = 0
        'Parameter Total Leave Balance Yearly
        Dim objItemSalaryFormulaItem As New Core_Module.ClsMasterSalaryFormula_item
        Dim mPrimaryLeave As Double = -1
        'Parameter Rounding untuk Translate Formula
        Dim mSIC_RoundingTipe As ClsMasterSalaryFormula_item.EnumRoundingType = ClsMasterSalaryFormula_item.EnumRoundingType.No_Rounding
        Dim mSIC_RoundingUnit As Decimal = 0

        Try
            Formula.ParserOptions.DecimalSeparator = "."
            Formula.ParserOptions.FunctionArgumentSeparator = ";"
            Formula.Options.ParseCulture.NumberFormat.NumberDecimalSeparator = "."
            Formula.Options.ParseCulture.NumberFormat.NumberGroupSeparator = ","
            Formula.ParserOptions.RecreateParser()
            Rumus = Rumus.Replace(",", "")

            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_Primary(TBL, vPrimaryFormula)
            TBL.DefaultView.RowFilter = "[Identifier Salary Component] = '" + vIdentifierFormula.ToString + "'"
            If TBL.DefaultView.Count > 0 Then
                For Each DV In TBL.DefaultView
                    Row = DV.Row
                    'Pass Parameter Primary Leave ke Fungsi GetValue
                    mPrimaryLeave = -1
                    If Row.Primary_Parameter = 132 Then
                        'Total Leave Balance Yearly, Primary Param Leave Code = 133
                        mPrimaryLeave = objItemSalaryFormulaItem.GetRefValueConstraint(Row.Identifier_Salary_Component, 133)
                    End If
                    If Row.Rounding_Type = ClsMasterSalaryFormula_item.EnumRoundingType.No_Rounding Then
                        var.Add(Row.Variable.ToString, CDbl(vObjParam.GetValue(TBLParam1, TBLParam2, TBLWeekDay, TBLHoliday, TblKaryawan, vPrimaryEmployee, Row.Primary_Parameter, vTglDari, vTglSampai, Row.Type2, vObjKary, mPrimaryLeave)))
                    ElseIf Row.Rounding_Type = ClsMasterSalaryFormula_item.EnumRoundingType.Round_Down Then
                        var.Add(Row.Variable.ToString, CDbl(ObjCD.GetRoundingValue(CDec(vObjParam.GetValue(TBLParam1, TBLParam2, TBLWeekDay, TBLHoliday, TblKaryawan, vPrimaryEmployee, Row.Primary_Parameter, vTglDari, vTglSampai, Row.Type2, vObjKary, mPrimaryLeave)), vDef_Prec, ClsCekData.EnumTipeRounding.Down, Row.Rounding_Unit, ObjCD)))
                    ElseIf Row.Rounding_Type = ClsMasterSalaryFormula_item.EnumRoundingType.Round_Middlepoint Then
                        var.Add(Row.Variable.ToString, CDbl(ObjCD.GetRoundingValue(CDec(vObjParam.GetValue(TBLParam1, TBLParam2, TBLWeekDay, _
                        TBLHoliday, TblKaryawan, vPrimaryEmployee, Row.Primary_Parameter, vTglDari, vTglSampai, Row.Type2, vObjKary, mPrimaryLeave)), vDef_Prec, ClsCekData.EnumTipeRounding.Middlepoint, Row.Rounding_Unit, ObjCD)))
                    ElseIf Row.Rounding_Type = ClsMasterSalaryFormula_item.EnumRoundingType.Round_Up Then
                        var.Add(Row.Variable.ToString, CDbl(ObjCD.GetRoundingValue(CDec(vObjParam.GetValue(TBLParam1, TBLParam2, TBLWeekDay, TBLHoliday, TblKaryawan, vPrimaryEmployee, Row.Primary_Parameter, vTglDari, vTglSampai, Row.Type2, vObjKary, mPrimaryLeave)), vDef_Prec, ClsCekData.EnumTipeRounding.Up, Row.Rounding_Unit, ObjCD)))
                    End If
                Next

                'Result 
                mSIC_RoundingTipe = objItemSalaryFormulaItem.Get_SCI_RoundingType(vIdentifierFormula)
                mSIC_RoundingUnit = objItemSalaryFormulaItem.Get_SCIRoundingUnit(vIdentifierFormula)
                If mSIC_RoundingTipe = ClsMasterSalaryFormula_item.EnumRoundingType.No_Rounding Then
                    Hasil = CDec(Formula.CompileDynamic(Rumus).Evaluate)
                ElseIf mSIC_RoundingTipe = ClsMasterSalaryFormula_item.EnumRoundingType.Round_Down Then
                    Hasil = ObjCD.GetRoundingValue(CDec(Formula.CompileDynamic(Rumus).Evaluate), vDef_Prec, ClsCekData.EnumTipeRounding.Down, mSIC_RoundingUnit, ObjCD)
                ElseIf mSIC_RoundingTipe = ClsMasterSalaryFormula_item.EnumRoundingType.Round_Middlepoint Then
                    Hasil = ObjCD.GetRoundingValue(CDec(Formula.CompileDynamic(Rumus).Evaluate), vDef_Prec, ClsCekData.EnumTipeRounding.Middlepoint, mSIC_RoundingUnit, ObjCD)
                ElseIf mSIC_RoundingTipe = ClsMasterSalaryFormula_item.EnumRoundingType.Round_Up Then
                    Hasil = ObjCD.GetRoundingValue(CDec(Formula.CompileDynamic(Rumus).Evaluate), vDef_Prec, ClsCekData.EnumTipeRounding.Up, mSIC_RoundingUnit, ObjCD)
                End If
            End If
            Return Hasil
        Catch ex As Exception
            MsgBox(ObjMasterSF_Item.GetDescComponent(vIdentifierFormula).ToString + " formula is incorrect, you cannot get result formula value.", MsgBoxStyle.Information, "Information")
            'MsgBox("Wrong formula, you cannot get result formula value.", MsgBoxStyle.Information, "Information")
            'MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjMasterSF_Item = Nothing
            objItemSalaryFormulaItem = Nothing
            ObjCD = Nothing
        End Try
    End Function

    Public Function IsConstraint_Daily_OtherDeduction(ByVal vPrimaryOther As Double, _
    ByVal vPrimaryEmployee As Double, ByVal vTglDari As Date, ByVal vTglSampai As Date, _
    ByVal TBLParam1 As DseMasterParameter.DataTable1DataTable, ByVal TBLParam2 As DseMasterParameter.tbltrans_attendanceDataTable, _
    ByVal TBLWeekDay As DseMasterParameter.tbltrans_attendance1DataTable, ByVal TBLHoliday As DseMasterParameter.tbltrans_attendance1DataTable, _
    ByVal TblKaryawan As DseMasterParameter.tblmaster_karyawanDataTable, ByVal vObjParam As ClsMasterParameterGaji, _
    ByVal TBLConst As DseTransaksiSalary.tblmaster_salary_formula_otherdeduction_constraintDataTable, ByVal vObjKrywn As Core_Module.ClsMasterKaryawan, _
    ByVal vObjKompetensi As Core_Module.ClsMasterKompetensi) As Boolean
        Dim RowConst As DseTransaksiSalary.tblmaster_salary_formula_otherdeduction_constraintRow
        Dim RV As Data.DataRowView
        Dim Filter As String = ""
        Dim Var1 As String = ""
        Dim Var2 As String = ""
        'Parameter Leave Code Constraint
        Dim objItemOtherDeduction As New Core_Module.ClsTrans_OtherDeduction_item
        Dim mPrimaryLeave As Double = -1
        Try
            Filter = "[Primary Transaksi] = " + vPrimaryOther.ToString + ""
            TBLConst.DefaultView.RowFilter = Filter
            If TBLConst.DefaultView.Count > 0 Then
                For Each RV In TBLConst.DefaultView
                    RowConst = RV.Row
                    'Pass Parameter Primary Leave ke Fungsi GetValue
                    mPrimaryLeave = -1
                    If RowConst.Primary_Parameter = 133 Then
                        'Get Leave Code
                        mPrimaryLeave = objItemOtherDeduction.GetRefValueConstraint_Others(RowConst.Primary_Transaksi, RowConst.Primary_Parameter)
                    End If
                    Var1 = vObjParam.GetValue_Daily(TBLParam1, TBLParam2, TBLWeekDay, TBLHoliday, TblKaryawan, vPrimaryEmployee, _
                    RowConst.Primary_Parameter, vTglDari, RowConst.Type2, vObjKrywn, mPrimaryLeave)
                    Var2 = RowConst.Ref_Value
                    If vObjParam.IsConstraintBenar(Var1, Var2, RowConst.Compare, RowConst.Primary_Parameter, RowConst.Type2, vObjKompetensi) = False Then
                        Return False
                    End If
                Next
            End If
            Return True
            TBLConst.DefaultView.RowFilter = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objItemOtherDeduction = Nothing
        End Try
    End Function

    Public Function IsConstraint_Daily(ByVal vIdentifierFormula As String, _
   ByVal vPrimaryEmployee As Double, ByVal vTglDari As Date, ByVal vTglSampai As Date, _
   ByVal TBLParam1 As DseMasterParameter.DataTable1DataTable, ByVal TBLParam2 As DseMasterParameter.tbltrans_attendanceDataTable, _
   ByVal TBLWeekDay As DseMasterParameter.tbltrans_attendance1DataTable, ByVal TBLHoliday As DseMasterParameter.tbltrans_attendance1DataTable, _
   ByVal TblKaryawan As DseMasterParameter.tblmaster_karyawanDataTable, ByVal vObjParam As ClsMasterParameterGaji, _
   ByVal TBLConst As DseTransaksiSalary.tblmaster_salary_formula_salarycomponent_constraintDataTable, ByVal vObjKary As Core_Module.ClsMasterKaryawan, _
   ByVal vObjKompetensi As Core_Module.ClsMasterKompetensi) As Boolean
        Dim RowConst As DseTransaksiSalary.tblmaster_salary_formula_salarycomponent_constraintRow
        Dim RV As Data.DataRowView
        Dim Filter As String = ""
        Dim Var1 As String = ""
        Dim Var2 As String = ""
        'Parameter Leave Code Constraint
        Dim objItemSalaryFormulaItem As New Core_Module.ClsMasterSalaryFormula_item
        Dim mPrimaryLeave As Double = -1
        Try
            Filter = "[Identifier salary Component] = '" + vIdentifierFormula.ToString + "'"
            TBLConst.DefaultView.RowFilter = Filter
            If TBLConst.DefaultView.Count > 0 Then
                For Each RV In TBLConst.DefaultView
                    RowConst = RV.Row
                    'Pass Parameter Primary Leave ke Fungsi GetValue
                    mPrimaryLeave = -1
                    If RowConst.Primary_Parameter = 133 Then
                        'Get Primary Leave 
                        mPrimaryLeave = objItemSalaryFormulaItem.GetRefValueConstraint(RowConst.Identifier_Salary_Component, RowConst.Primary_Parameter)
                    End If
                    Var1 = vObjParam.GetValue_Daily(TBLParam1, TBLParam2, TBLWeekDay, TBLHoliday, TblKaryawan, vPrimaryEmployee, _
                    RowConst.Primary_Parameter, vTglDari, RowConst.Type2, vObjKary, mPrimaryLeave)

                    Var2 = RowConst.Ref_Value

                    If vObjParam.IsConstraintBenar(Var1, Var2, RowConst.Compare, RowConst.Primary_Parameter, RowConst.Type2, vObjKompetensi) = False Then
                        Return False
                    End If
                Next
            End If
            Return True
            TBLConst.DefaultView.RowFilter = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objItemSalaryFormulaItem = Nothing
        End Try
    End Function

    Public Function GetDescriptionOtherDeduction(ByVal vPrimaryOther As Double) As String
        Dim DTA As New DseTransaksiSalaryTableAdapters.tblmaster_salary_formula_otherdeductionTableAdapter
        Dim Dummy As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetDescriptionByPrimaryOther(vPrimaryOther)
            If String.IsNullOrEmpty(Dummy) = False Then
                Return Dummy
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

    Public Function TranslateFormula_Daily_OtherDeduction(ByVal Rumus As String, ByVal vPrimaryOther As Double, _
    ByVal vPrimaryEmployee As Double, ByVal vTglDari As Date, ByVal vTglSampai As Date, _
    ByVal TBLParam1 As DseMasterParameter.DataTable1DataTable, ByVal TBLParam2 As DseMasterParameter.tbltrans_attendanceDataTable, _
    ByVal TBLWeekDay As DseMasterParameter.tbltrans_attendance1DataTable, ByVal TBLHoliday As DseMasterParameter.tbltrans_attendance1DataTable, _
    ByVal TblKaryawan As DseMasterParameter.tblmaster_karyawanDataTable, ByVal vObjParam As ClsMasterParameterGaji, ByVal vObjKary As Core_Module.ClsMasterKaryawan) As Decimal
        Dim Formula As New ExpressionContext
        Dim var As VariableCollection = Formula.Variables
        Dim DTA As New DseTransaksiSalaryTableAdapters.tblmaster_salary_formula_otherdeduction_variableTableAdapter
        Dim TBL As New DseTransaksiSalary.tblmaster_salary_formula_otherdeduction_variableDataTable
        Dim Row As DseTransaksiSalary.tblmaster_salary_formula_otherdeduction_variableRow
        Dim Hasil As Decimal = 0
        Dim Filter As String = ""
        Dim DV As Data.DataRowView
        Dim ObjMasterSF_Item As New Core_Module.ClsTrans_OtherDeduction_item
        Dim ObjCD As New Core_Module.ClsCekData
        Dim vDef_Prec As Decimal = 0
        'Parameter Rounding untuk Translate Formula
        Dim objItemSalaryFormulaItem_Other As New Core_Module.ClsMasterSalaryFormula_item
        Dim mSIC_RoundingTipe As ClsTrans_OtherDeduction_item.EnumRoundingType = ClsTrans_OtherDeduction_item.EnumRoundingType.No_Rounding
        Dim mSIC_RoundingUnit As Decimal = 0

        Try
            Formula.ParserOptions.DecimalSeparator = "."
            Formula.ParserOptions.FunctionArgumentSeparator = ";"
            Formula.Options.ParseCulture.NumberFormat.NumberDecimalSeparator = "."
            Formula.Options.ParseCulture.NumberFormat.NumberGroupSeparator = ","
            Formula.ParserOptions.RecreateParser()
            Rumus = Rumus.Replace(",", "")

            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_PrimaryOther(TBL, vPrimaryOther)
            TBL.DefaultView.RowFilter = "[Primary Transaksi] = " + vPrimaryOther.ToString + ""
            If TBL.DefaultView.Count > 0 Then
                For Each DV In TBL.DefaultView
                    Row = DV.Row
                    If Row.Rounding_Type = ClsTrans_OtherDeduction_item.EnumRoundingType.No_Rounding Then
                        var.Add(Row.Variable.ToString, CDbl(vObjParam.GetValue_Daily(TBLParam1, TBLParam2, TBLWeekDay, _
                        TBLHoliday, TblKaryawan, vPrimaryEmployee, Row.Primary_Parameter, vTglDari, Row.Type2, vObjKary, -1)))
                    ElseIf Row.Rounding_Type = ClsTrans_OtherDeduction_item.EnumRoundingType.Round_Down Then
                        var.Add(Row.Variable.ToString, CDbl(ObjCD.GetRoundingValue(CDec(vObjParam.GetValue_Daily(TBLParam1, TBLParam2, TBLWeekDay, _
                        TBLHoliday, TblKaryawan, vPrimaryEmployee, Row.Primary_Parameter, vTglDari, Row.Type2, vObjKary, -1)), 0, ClsCekData.EnumTipeRounding.Down, Row.Rounding_Unit, ObjCD)))
                    ElseIf Row.Rounding_Type = ClsTrans_OtherDeduction_item.EnumRoundingType.Round_Middlepoint Then
                        var.Add(Row.Variable.ToString, CDbl(ObjCD.GetRoundingValue(CDec(vObjParam.GetValue_Daily(TBLParam1, TBLParam2, TBLWeekDay, _
                        TBLHoliday, TblKaryawan, vPrimaryEmployee, Row.Primary_Parameter, vTglDari, Row.Type2, vObjKary, -1)), 0, ClsCekData.EnumTipeRounding.Middlepoint, Row.Rounding_Unit, ObjCD)))
                    ElseIf Row.Rounding_Type = ClsTrans_OtherDeduction_item.EnumRoundingType.Round_Up Then
                        var.Add(Row.Variable.ToString, CDbl(ObjCD.GetRoundingValue(CDec(vObjParam.GetValue_Daily(TBLParam1, TBLParam2, TBLWeekDay, _
                        TBLHoliday, TblKaryawan, vPrimaryEmployee, Row.Primary_Parameter, vTglDari, Row.Type2, vObjKary, -1)), 0, ClsCekData.EnumTipeRounding.Up, Row.Rounding_Unit, ObjCD)))
                    End If
                Next

                'Result
                mSIC_RoundingTipe = objItemSalaryFormulaItem_Other.Get_SCI_RoundingType(vPrimaryOther)
                mSIC_RoundingUnit = objItemSalaryFormulaItem_Other.Get_SCIRoundingUnit(vPrimaryOther)
                If mSIC_RoundingTipe = ClsMasterSalaryFormula_item.EnumRoundingType.No_Rounding Then
                    Hasil = CDec(Formula.CompileDynamic(Rumus).Evaluate)
                ElseIf mSIC_RoundingTipe = ClsTrans_OtherDeduction_item.EnumRoundingType.Round_Down Then
                    Hasil = ObjCD.GetRoundingValue(CDec(Formula.CompileDynamic(Rumus).Evaluate), vDef_Prec, ClsCekData.EnumTipeRounding.Down, mSIC_RoundingUnit, ObjCD)
                ElseIf mSIC_RoundingTipe = ClsTrans_OtherDeduction_item.EnumRoundingType.Round_Middlepoint Then
                    Hasil = ObjCD.GetRoundingValue(CDec(Formula.CompileDynamic(Rumus).Evaluate), vDef_Prec, ClsCekData.EnumTipeRounding.Middlepoint, mSIC_RoundingUnit, ObjCD)
                ElseIf mSIC_RoundingTipe = ClsTrans_OtherDeduction_item.EnumRoundingType.Round_Up Then
                    Hasil = ObjCD.GetRoundingValue(CDec(Formula.CompileDynamic(Rumus).Evaluate), vDef_Prec, ClsCekData.EnumTipeRounding.Up, mSIC_RoundingUnit, ObjCD)
                End If
            End If
            Return Hasil
        Catch ex As Exception
            MsgBox(Me.GetDescriptionOtherDeduction(vPrimaryOther).ToString + " formula is incorrect, you cannot get result formula value.", MsgBoxStyle.Information, "Information")
            'MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjMasterSF_Item = Nothing
            ObjCD = Nothing
            objItemSalaryFormulaItem_Other = Nothing
        End Try
    End Function


    Public Function TranslateFormula_Daily(ByVal Rumus As String, ByVal vPrimaryFormula As Double, _
    ByVal vPrimaryEmployee As Double, ByVal vTglDari As Date, ByVal vTglSampai As Date, _
    ByVal TBLParam1 As DseMasterParameter.DataTable1DataTable, ByVal TBLParam2 As DseMasterParameter.tbltrans_attendanceDataTable, _
    ByVal TBLWeekDay As DseMasterParameter.tbltrans_attendance1DataTable, ByVal TBLHoliday As DseMasterParameter.tbltrans_attendance1DataTable, _
    ByVal TblKaryawan As DseMasterParameter.tblmaster_karyawanDataTable, ByVal vObjParam As ClsMasterParameterGaji, _
    ByVal vIdentifierFormula As String, ByVal vObjKary As Core_Module.ClsMasterKaryawan) As Decimal
        Dim Formula As New ExpressionContext
        Dim var As VariableCollection = Formula.Variables
        Dim DTA As New DseTransaksiSalaryTableAdapters.tblmaster_salary_formula_salarycomponent_variableTableAdapter
        Dim TBL As New DseTransaksiSalary.tblmaster_salary_formula_salarycomponent_variableDataTable
        Dim Row As DseTransaksiSalary.tblmaster_salary_formula_salarycomponent_variableRow
        Dim Hasil As Decimal = 0
        Dim Filter As String = ""
        Dim DV As Data.DataRowView
        Dim ObjMasterSF_Item As New Core_Module.ClsMasterSalaryFormula_item
        Dim ObjCD As New Core_Module.ClsCekData
        Dim vDef_Prec As Decimal = 0
        'Parameter Rounding untuk Translate Formula
        Dim objItemSalaryFormulaItem As New Core_Module.ClsMasterSalaryFormula_item
        Dim mSIC_RoundingTipe As ClsMasterSalaryFormula_item.EnumRoundingType = ClsMasterSalaryFormula_item.EnumRoundingType.No_Rounding
        Dim mSIC_RoundingUnit As Decimal = 0

        Try
            Formula.ParserOptions.DecimalSeparator = "."
            Formula.ParserOptions.FunctionArgumentSeparator = ";"
            Formula.Options.ParseCulture.NumberFormat.NumberDecimalSeparator = "."
            Formula.Options.ParseCulture.NumberFormat.NumberGroupSeparator = ","
            Formula.ParserOptions.RecreateParser()
            Rumus = Rumus.Replace(",", "")
            'Formula.CompileDynamic(Rumus)

            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_Primary(TBL, vPrimaryFormula)
            TBL.DefaultView.RowFilter = "[Identifier Salary Component] = '" + vIdentifierFormula.ToString + "'"
            If TBL.DefaultView.Count > 0 Then
                For Each DV In TBL.DefaultView
                    Row = DV.Row
                    If Row.Rounding_Type = ClsMasterSalaryFormula_item.EnumRoundingType.No_Rounding Then
                        var.Add(Row.Variable.ToString, CDbl(vObjParam.GetValue_Daily(TBLParam1, TBLParam2, TBLWeekDay, TBLHoliday, TblKaryawan, vPrimaryEmployee, Row.Primary_Parameter, vTglDari, Row.Type2, vObjKary, -1)))
                    ElseIf Row.Rounding_Type = ClsMasterSalaryFormula_item.EnumRoundingType.Round_Down Then
                        var.Add(Row.Variable.ToString, CDbl(ObjCD.GetRoundingValue(CDec(vObjParam.GetValue_Daily(TBLParam1, TBLParam2, TBLWeekDay, _
                        TBLHoliday, TblKaryawan, vPrimaryEmployee, Row.Primary_Parameter, vTglDari, Row.Type2, vObjKary, -1)), vDef_Prec, ClsCekData.EnumTipeRounding.Down, Row.Rounding_Unit, ObjCD)))
                    ElseIf Row.Rounding_Type = ClsMasterSalaryFormula_item.EnumRoundingType.Round_Middlepoint Then
                        var.Add(Row.Variable.ToString, CDbl(ObjCD.GetRoundingValue(CDec(vObjParam.GetValue_Daily(TBLParam1, TBLParam2, TBLWeekDay, _
                        TBLHoliday, TblKaryawan, vPrimaryEmployee, Row.Primary_Parameter, vTglDari, Row.Type2, vObjKary, -1)), vDef_Prec, ClsCekData.EnumTipeRounding.Middlepoint, Row.Rounding_Unit, ObjCD)))
                    ElseIf Row.Rounding_Type = ClsMasterSalaryFormula_item.EnumRoundingType.Round_Up Then
                        var.Add(Row.Variable.ToString, CDbl(ObjCD.GetRoundingValue(CDec(vObjParam.GetValue_Daily(TBLParam1, TBLParam2, TBLWeekDay, _
                        TBLHoliday, TblKaryawan, vPrimaryEmployee, Row.Primary_Parameter, vTglDari, Row.Type2, vObjKary, -1)), vDef_Prec, ClsCekData.EnumTipeRounding.Up, Row.Rounding_Unit, ObjCD)))
                    End If
                Next

                'Result
                mSIC_RoundingTipe = objItemSalaryFormulaItem.Get_SCI_RoundingType(vIdentifierFormula)
                mSIC_RoundingUnit = objItemSalaryFormulaItem.Get_SCIRoundingUnit(vIdentifierFormula)
                If mSIC_RoundingTipe = ClsMasterSalaryFormula_item.EnumRoundingType.No_Rounding Then
                    Hasil = CDec(Formula.CompileDynamic(Rumus).Evaluate)
                ElseIf mSIC_RoundingTipe = ClsMasterSalaryFormula_item.EnumRoundingType.Round_Down Then
                    Hasil = ObjCD.GetRoundingValue(CDec(Formula.CompileDynamic(Rumus).Evaluate), vDef_Prec, ClsCekData.EnumTipeRounding.Down, mSIC_RoundingUnit, ObjCD)
                ElseIf mSIC_RoundingTipe = ClsMasterSalaryFormula_item.EnumRoundingType.Round_Middlepoint Then
                    Hasil = ObjCD.GetRoundingValue(CDec(Formula.CompileDynamic(Rumus).Evaluate), vDef_Prec, ClsCekData.EnumTipeRounding.Middlepoint, mSIC_RoundingUnit, ObjCD)
                ElseIf mSIC_RoundingTipe = ClsMasterSalaryFormula_item.EnumRoundingType.Round_Up Then
                    Hasil = ObjCD.GetRoundingValue(CDec(Formula.CompileDynamic(Rumus).Evaluate), vDef_Prec, ClsCekData.EnumTipeRounding.Up, mSIC_RoundingUnit, ObjCD)
                End If
            End If
            Return Hasil
        Catch ex As Exception
            MsgBox(ObjMasterSF_Item.GetDescComponent(vIdentifierFormula).ToString + " formula is incorrect, you cannot get result formula value.", MsgBoxStyle.Information, "Information")
            'MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjMasterSF_Item = Nothing
            ObjCD = Nothing
            objItemSalaryFormulaItem = Nothing
        End Try
    End Function

    Public Function GetMAXPrimary(ByVal vPrimaryEmployee As Double) As Double
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salaryTableAdapter
        Dim Dummy As System.Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetMAXPrimaryByEmployee(vPrimaryEmployee)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetEndDate(ByVal vPrimary As Double) As Date
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salaryTableAdapter
        Dim RetVal As Nullable(Of Date)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetEndDate(vPrimary)

            If RetVal.HasValue = True Then
                Return RetVal.Value
            Else
                Return "#1/1/" + Today.Year.ToString + "#"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "#1/1/" + Today.Year.ToString + "#"
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetEndDate_Minimum(ByVal vPrimaryKaryawan As Double) As Date
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salaryTableAdapter
        Dim RetVal As Nullable(Of Date)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetEndDateMinimum(vPrimaryKaryawan)

            If RetVal.HasValue = True Then
                Return RetVal.Value
            Else
                Return "#1/1/" + Today.Year.ToString + "#"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "#1/1/" + Today.Year.ToString + "#"
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPeriodPaid(ByVal vPrimary As Double) As Double
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salaryTableAdapter
        Dim Dummy As System.Nullable(Of Integer)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetPeriodPaid(vPrimary)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Enum EnumNilaiAwal As Integer
        Tax_Calculation
        Periode_Hitung
    End Enum

    Public Function MenentukanNilaiAwal(ByVal vEnumNilaiAwal As EnumNilaiAwal, ByVal vPrimaryEmployee As Double) As Double
        Dim PrimaryMAX As Double = Me.GetMAXPrimary(vPrimaryEmployee)
        Try
            If vEnumNilaiAwal = EnumNilaiAwal.Tax_Calculation Then
                Return Me.GetTaxTypeCalculation(PrimaryMAX)
            Else
                Return Me.GetPeriodPaid(PrimaryMAX)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function EndResultValue_Final_OtherDeduction(ByVal vApakahDaily As Boolean, ByRef vParamYearlyMax As String, ByRef vParamYearlyMin As String, ByRef vParamMonthlyMax As String, ByRef vParamMonthlyMin As String, _
    ByVal vPrimaryEmployee As Double, ByVal DtpFrom As Date, ByVal CalculateOn As ClsMasterSalaryFormula_item.EnumCalculateOn, ByVal vTotalRoundingPeriod As Decimal, ByVal vPrimaryOther As Double, _
    ByVal vTglDari As Date, ByVal vTglSampai As Date, _
    ByVal TBLParam1 As DseMasterParameter.DataTable1DataTable, ByVal TBLParam2 As DseMasterParameter.tbltrans_attendanceDataTable, _
    ByVal TBLWeekDay As DseMasterParameter.tbltrans_attendance1DataTable, ByVal TBLHoliday As DseMasterParameter.tbltrans_attendance1DataTable, _
    ByVal TblKaryawan As DseMasterParameter.tblmaster_karyawanDataTable, ByVal vObjParam As ClsMasterParameterGaji, ByVal vObjKary As Core_Module.ClsMasterKaryawan) As Decimal
        Dim Retval As Decimal = 0
        Try
            If vApakahDaily = True Then
                If IsNumeric(vParamYearlyMax) = False Then
                    vParamMonthlyMax = Me.TranslateFormula_Daily_OtherDeduction(vParamYearlyMax, vPrimaryOther, vPrimaryEmployee, vTglDari, vTglSampai, _
                     TBLParam1, TBLParam2, TBLWeekDay, TBLHoliday, TblKaryawan, vObjParam, vObjKary)
                End If
                If IsNumeric(vParamYearlyMin) = False Then
                    vParamYearlyMin = Me.TranslateFormula_Daily_OtherDeduction(vParamYearlyMin, vPrimaryOther, vPrimaryEmployee, vTglDari, vTglSampai, _
                     TBLParam1, TBLParam2, TBLWeekDay, TBLHoliday, TblKaryawan, vObjParam, vObjKary)
                End If
                If IsNumeric(vParamMonthlyMax) = False Then
                    vParamMonthlyMax = Me.TranslateFormula_Daily_OtherDeduction(vParamMonthlyMax, vPrimaryOther, vPrimaryEmployee, vTglDari, vTglSampai, _
                    TBLParam1, TBLParam2, TBLWeekDay, TBLHoliday, TblKaryawan, vObjParam, vObjKary)
                End If
                If IsNumeric(vParamMonthlyMin) = False Then
                    vParamMonthlyMin = Me.TranslateFormula_Daily_OtherDeduction(vParamMonthlyMin, vPrimaryOther, vPrimaryEmployee, vTglDari, vTglSampai, _
                    TBLParam1, TBLParam2, TBLWeekDay, TBLHoliday, TblKaryawan, vObjParam, vObjKary)
                End If
            ElseIf vApakahDaily = False Then
                If IsNumeric(vParamYearlyMax) = False Then
                    vParamMonthlyMax = Me.TranslateFormulaMonthlyYearly_OtherDeduction(vParamYearlyMax, vPrimaryOther, vPrimaryEmployee, vTglDari, vTglSampai, _
                     TBLParam1, TBLParam2, TBLWeekDay, TBLHoliday, TblKaryawan, vObjParam, vObjKary)
                End If
                If IsNumeric(vParamYearlyMin) = False Then
                    vParamYearlyMin = Me.TranslateFormulaMonthlyYearly_OtherDeduction(vParamYearlyMin, vPrimaryOther, vPrimaryEmployee, vTglDari, vTglSampai, _
                     TBLParam1, TBLParam2, TBLWeekDay, TBLHoliday, TblKaryawan, vObjParam, vObjKary)
                End If
                If IsNumeric(vParamMonthlyMax) = False Then
                    vParamMonthlyMax = Me.TranslateFormulaMonthlyYearly_OtherDeduction(vParamMonthlyMax, vPrimaryOther, vPrimaryEmployee, vTglDari, vTglSampai, _
                    TBLParam1, TBLParam2, TBLWeekDay, TBLHoliday, TblKaryawan, vObjParam, vObjKary)
                End If
                If IsNumeric(vParamMonthlyMin) = False Then
                    vParamMonthlyMin = Me.TranslateFormulaMonthlyYearly_OtherDeduction(vParamMonthlyMin, vPrimaryOther, vPrimaryEmployee, vTglDari, vTglSampai, _
                    TBLParam1, TBLParam2, TBLWeekDay, TBLHoliday, TblKaryawan, vObjParam, vObjKary)
                End If
            End If
            Retval = Me.EndResultValue_OtherDeduction(vPrimaryEmployee, DtpFrom, CalculateOn, vTotalRoundingPeriod, vParamYearlyMax, _
            vParamYearlyMin, vParamMonthlyMax, vParamYearlyMin, vPrimaryOther)

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function EndResultValue_OtherDeduction(ByVal vPrimaryEmployee As Double, ByVal DtpFrom As Date, ByVal CalculateOn As ClsMasterSalaryFormula_item.EnumCalculateOn, _
    ByVal vTotalRoundingPeriod As Decimal, ByVal YearlyMax As Decimal, ByVal YearlyMin As Decimal, ByVal MonthlyMax As Decimal, ByVal MonthlyMin As Decimal, _
    ByVal vPrimaryOther As Double) As Decimal
        Dim ObjCD As New ClsCekData
        Dim ObjFormulaItem As New ClsMasterSalaryFormula_item
        Dim Bulan As Integer = 0
        Dim TglDari As Date = #1/1/1950#
        Dim TglSampai As Date = #1/1/1950#
        Dim Retval As Decimal = 0
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salaryTableAdapter
        Dim Dummy As System.Nullable(Of Decimal)
        Dim AmountValue As Decimal = 0
        'Dim AwalPerhitunganDalamTahun As Boolean = False
        Try
            TglDari = CDate("#1/1/" + DtpFrom.Year.ToString + "#")
            TglSampai = CDate(ObjCD.GetLasDateOfTheMonth(DtpFrom.Date)) ' - TimeSpan.FromDays(1))
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'Bener in hanya amount dengan identifier tertentu / dengan kategori potongan tertentu tidak semua ditarik.
            Dummy = DTA.ScalarQuery_GetAmountOtherDeduction(vPrimaryEmployee, TglDari, TglSampai, vPrimaryOther)
            If Dummy.HasValue Then
                AmountValue = Dummy
            Else
                AmountValue = 0
            End If
            'Rounding Period-nya di cek masuk batasan Monthly MIN, dan MAX
            Dim CekMonthlyMin As Boolean = True
            If MonthlyMax > 0 Then
                If vTotalRoundingPeriod > MonthlyMax Then
                    vTotalRoundingPeriod = MonthlyMax
                    CekMonthlyMin = False
                End If
            End If
            If CekMonthlyMin = True Then
                If MonthlyMin > 0 Then
                    If vTotalRoundingPeriod < MonthlyMin Then
                        vTotalRoundingPeriod = MonthlyMin
                    End If
                End If
            End If
            ''
            AmountValue += vTotalRoundingPeriod
            Retval = vTotalRoundingPeriod
            'Yearly
            If YearlyMax > 0 Then
                If AmountValue > YearlyMax Then
                    AmountValue -= vTotalRoundingPeriod
                    If AmountValue > 0 Then
                        Retval = YearlyMax - AmountValue
                        If ObjCD.CekBilanganPositif(Retval) = False Then
                            Retval = 0
                        End If
                        Return Retval
                    End If
                    ''
                End If
            End If
            Bulan = ObjFormulaItem.ConversionMonth(CalculateOn)
            If TglSampai.Month = CDate("#" + Bulan.ToString + "/1/1950#").AddMonths(-1).Month Then
                'Jika Akhir Tahun perhitungan, maka cek dengan batasan Yearly Min
                If YearlyMin > 0 Then
                    If AmountValue < YearlyMin Then
                        AmountValue -= vTotalRoundingPeriod
                        Retval = YearlyMin - AmountValue
                    Else
                        Retval = vTotalRoundingPeriod
                    End If
                End If
            End If
            'Monthly
            If MonthlyMax > 0 Then
                If Retval > MonthlyMax Then
                    Retval = MonthlyMax
                    Return Retval
                End If
            End If
            If MonthlyMin > 0 Then
                If Retval < MonthlyMin Then
                    Retval = MonthlyMin
                    Return Retval
                End If
            End If
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjCD = Nothing
            ObjFormulaItem = Nothing
        End Try
    End Function

    Public Function EndResultValue_Final(ByVal vApakahDaily As Boolean, ByRef vParamYearlyMax As String, ByRef vParamYearlyMin As String, ByRef vParamMonthlyMax As String, ByRef vParamMonthlyMin As String, _
    ByVal vPrimaryEmployee As Double, ByVal DtpFrom As Date, ByVal CalculateOn As ClsMasterSalaryFormula_item.EnumCalculateOn, ByVal vTotalRoundingPeriod As Decimal, ByVal vIdentifierSalaryComponent As String, _
    ByVal vPrimaryFormula As Double, ByVal vTglDari As Date, ByVal vTglSampai As Date, _
    ByVal TBLParam1 As DseMasterParameter.DataTable1DataTable, ByVal TBLParam2 As DseMasterParameter.tbltrans_attendanceDataTable, _
    ByVal TBLWeekDay As DseMasterParameter.tbltrans_attendance1DataTable, ByVal TBLHoliday As DseMasterParameter.tbltrans_attendance1DataTable, _
    ByVal TblKaryawan As DseMasterParameter.tblmaster_karyawanDataTable, ByVal vObjParam As ClsMasterParameterGaji, ByVal vObjKary As Core_Module.ClsMasterKaryawan) As Decimal
        Dim Retval As Decimal = 0
        Try
            If vApakahDaily = True Then
                If IsNumeric(vParamYearlyMax) = False Then
                    vParamMonthlyMax = Me.TranslateFormula_Daily(vParamYearlyMax, vPrimaryFormula, vPrimaryEmployee, vTglDari, vTglSampai, _
                     TBLParam1, TBLParam2, TBLWeekDay, TBLHoliday, TblKaryawan, vObjParam, vIdentifierSalaryComponent, vObjKary)
                End If
                If IsNumeric(vParamYearlyMin) = False Then
                    vParamYearlyMin = Me.TranslateFormula_Daily(vParamYearlyMin, vPrimaryFormula, vPrimaryEmployee, vTglDari, vTglSampai, _
                     TBLParam1, TBLParam2, TBLWeekDay, TBLHoliday, TblKaryawan, vObjParam, vIdentifierSalaryComponent, vObjKary)
                End If
                If IsNumeric(vParamMonthlyMax) = False Then
                    vParamMonthlyMax = Me.TranslateFormula_Daily(vParamMonthlyMax, vPrimaryFormula, vPrimaryEmployee, vTglDari, vTglSampai, _
                    TBLParam1, TBLParam2, TBLWeekDay, TBLHoliday, TblKaryawan, vObjParam, vIdentifierSalaryComponent, vObjKary)
                End If
                If IsNumeric(vParamMonthlyMin) = False Then
                    vParamMonthlyMin = Me.TranslateFormula_Daily(vParamMonthlyMin, vPrimaryFormula, vPrimaryEmployee, vTglDari, vTglSampai, _
                    TBLParam1, TBLParam2, TBLWeekDay, TBLHoliday, TblKaryawan, vObjParam, vIdentifierSalaryComponent, vObjKary)
                End If
            ElseIf vApakahDaily = False Then
                If IsNumeric(vParamYearlyMax) = False Then
                    vParamMonthlyMax = Me.TranslateFormula(vParamYearlyMax, vPrimaryFormula, vPrimaryEmployee, vTglDari, vTglSampai, _
                     TBLParam1, TBLParam2, TBLWeekDay, TBLHoliday, TblKaryawan, vObjParam, vIdentifierSalaryComponent, vObjKary)
                End If
                If IsNumeric(vParamYearlyMin) = False Then
                    vParamYearlyMin = Me.TranslateFormula(vParamYearlyMin, vPrimaryFormula, vPrimaryEmployee, vTglDari, vTglSampai, _
                     TBLParam1, TBLParam2, TBLWeekDay, TBLHoliday, TblKaryawan, vObjParam, vIdentifierSalaryComponent, vObjKary)
                End If
                If IsNumeric(vParamMonthlyMax) = False Then
                    vParamMonthlyMax = Me.TranslateFormula(vParamMonthlyMax, vPrimaryFormula, vPrimaryEmployee, vTglDari, vTglSampai, _
                    TBLParam1, TBLParam2, TBLWeekDay, TBLHoliday, TblKaryawan, vObjParam, vIdentifierSalaryComponent, vObjKary)
                End If
                If IsNumeric(vParamMonthlyMin) = False Then
                    vParamMonthlyMin = Me.TranslateFormula(vParamMonthlyMin, vPrimaryFormula, vPrimaryEmployee, vTglDari, vTglSampai, _
                    TBLParam1, TBLParam2, TBLWeekDay, TBLHoliday, TblKaryawan, vObjParam, vIdentifierSalaryComponent, vObjKary)
                End If
            End If
            Retval = Me.EndResultValue(vPrimaryEmployee, DtpFrom, CalculateOn, vTotalRoundingPeriod, vParamYearlyMax, _
            vParamYearlyMin, vParamMonthlyMax, vParamMonthlyMin, vIdentifierSalaryComponent)

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function EndResultValue(ByVal vPrimaryEmployee As Double, ByVal DtpFrom As Date, ByVal CalculateOn As ClsMasterSalaryFormula_item.EnumCalculateOn, _
    ByVal vTotalRoundingPeriod As Decimal, ByVal YearlyMax As Decimal, ByVal YearlyMin As Decimal, ByVal MonthlyMax As Decimal, ByVal MonthlyMin As Decimal, _
    ByVal vIdentifierSalaryComponent As String) As Decimal
        Dim ObjCD As New ClsCekData
        Dim ObjFormulaItem As New ClsMasterSalaryFormula_item
        Dim Bulan As Integer = 0
        Dim TglDari As Date = #1/1/1950#
        Dim TglSampai As Date = #1/1/1950#
        Dim Retval As Decimal = 0
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salaryTableAdapter
        Dim Dummy As System.Nullable(Of Decimal)
        Dim AmountValue As Decimal = 0
        'Dim AwalPerhitunganDalamTahun As Boolean = False
        Try
            TglDari = CDate("#1/1/" + DtpFrom.Year.ToString + "#")
            TglSampai = CDate(ObjCD.GetLasDateOfTheMonth(DtpFrom.Date)) ' - TimeSpan.FromDays(1))
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'Bener in hanya amount dengan identifier tertentu / dengan kategori potongan tertentu tidak semua ditarik.
            Dummy = DTA.ScalarQuery_GetAmountValue(vPrimaryEmployee, TglDari, TglSampai, vIdentifierSalaryComponent)
            If Dummy.HasValue Then
                AmountValue = Dummy
            Else
                AmountValue = 0
            End If
            'Rounding Period-nya di cek masuk batasan Monthly MIN, dan MAX
            Dim CekMonthlyMin As Boolean = True
            If MonthlyMax > 0 Then
                If vTotalRoundingPeriod > MonthlyMax Then
                    vTotalRoundingPeriod = MonthlyMax
                    CekMonthlyMin = False
                End If
            End If
            If CekMonthlyMin = True Then
                If MonthlyMin > 0 Then
                    If vTotalRoundingPeriod < MonthlyMin Then
                        vTotalRoundingPeriod = MonthlyMin
                    End If
                End If
            End If
            ''
            AmountValue += vTotalRoundingPeriod
            Retval = vTotalRoundingPeriod
            'Yearly
            If YearlyMax > 0 Then
                If AmountValue > YearlyMax Then
                    AmountValue -= vTotalRoundingPeriod
                    If AmountValue > 0 Then
                        Retval = YearlyMax - AmountValue
                        If ObjCD.CekBilanganPositif(Retval) = False Then
                            Retval = 0
                        End If
                        Return Retval
                    End If
                    ''
                End If
            End If
            Bulan = ObjFormulaItem.ConversionMonth(CalculateOn)
            If TglSampai.Month = CDate("#" + Bulan.ToString + "/1/1950#").AddMonths(-1).Month Then
                'Jika Akhir Tahun perhitungan, maka cek dengan batasan Yearly Min
                If YearlyMin > 0 Then
                    If AmountValue < YearlyMin Then
                        AmountValue -= vTotalRoundingPeriod
                        Retval = YearlyMin - AmountValue
                    Else
                        Retval = vTotalRoundingPeriod
                    End If
                End If
            End If
            'Monthly
            If MonthlyMax > 0 Then
                If Retval > MonthlyMax Then
                    Retval = MonthlyMax
                    Return Retval
                End If
            End If
            If MonthlyMin > 0 Then
                If Retval < MonthlyMin Then
                    Retval = MonthlyMin
                    Return Retval
                End If
            End If
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjCD = Nothing
            ObjFormulaItem = Nothing
        End Try
    End Function

    Public Function FillTBLStartEnd(ByVal vPrimaryEmployee As Double, ByVal EndDate As Date, ByVal StartDate As Date, _
    Optional ByVal ApakahBaru As Boolean = True, Optional ByVal vPrimaryTransSalary As Double = -1, _
    Optional ByVal vShowMsg As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salary1TableAdapter
        Dim TBL As New DseTransaksiSalary.tbltrans_salary1DataTable
        Dim Row As DseTransaksiSalary.tbltrans_salary1Row
        Dim ij As Integer = 1
        Dim Retval As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If ApakahBaru = True Then
                DTA.FillBy_PrimaryEmployee(TBL, EndDate, StartDate, vPrimaryEmployee)
            Else
                DTA.FillBy_PrimaryEmployeePosisiEdit(TBL, EndDate, StartDate, vPrimaryEmployee, vPrimaryTransSalary)
            End If
            If TBL.Rows.Count > 0 Then
                For Each Row In TBL.Rows
                    If ij = 1 Then
                        Retval = "You can't create salary from " + Format(StartDate, "dd MMM yyyy") + " to " + Format(EndDate, "dd MMM yyyy") + " because it's used for another salary."
                        Retval += vbCrLf + "The Following is a list of salary transactions have been used : "
                        Retval += vbCrLf + ij.ToString + ". " + Row.No_Transaksi.ToString + ", From : " + Format(Row.Start_Date, "dd MMM yyyy") + " to " + Format(Row.End_Date, "dd MMM yyyy")
                    Else
                        Retval += vbCrLf + ij.ToString + ". " + Row.No_Transaksi.ToString + ", From : " + Format(Row.Start_Date, "dd MMM yyyy") + " to " + Format(Row.End_Date, "dd MMM yyyy")
                    End If
                    ij += 1
                Next
                If vShowMsg = True Then
                    MsgBox(Retval.ToString, MsgBoxStyle.Information, "information")
                End If
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function CekValidPeriodEmp_HubunganAttendance(ByVal vPrimaryEmp As Double, _
    ByVal vStartDate As Date, ByVal vEndDate As Date, Optional ByVal ShowMsg As Boolean = True) As Boolean
        Dim dta As New DseTrans_AttendanceTableAdapters.tbltrans_attendanceTableAdapter
        Dim tbl As New DseTrans_Attendance.tbltrans_attendanceDataTable
        Dim row As DseTrans_Attendance.tbltrans_attendanceRow
        Dim Dta1 As New DseTransaksiSalaryTableAdapters.DataTable5TableAdapter
        Dim tbl1 As New DseTransaksiSalary.DataTable5DataTable
        Dim row1 As DseTransaksiSalary.DataTable5Row
        Dim DTA2 As New DseTransaksiSalaryTableAdapters.tbltrans_attendanceTableAdapter
        Dim TBL2 As New DseTransaksiSalary.tbltrans_attendanceDataTable
        Dim DTA3 As New DseTransaksiSalaryTableAdapters.tbltrans_attendance_itemTableAdapter
        Dim JumlahDate As Double = 0
        Dim TS As TimeSpan
        Dim ParamHari As Double = 0
        Dim ListPeriod As String = ""
        Dim ListPrimaryAttendance As New List(Of Double)
        Dim PrimaryAttendance As Double = -1
        Dim i As Integer = 1
        Dim Filter As String = ""
        Dim MunculkanMsgBox As Boolean = False
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            Dta1.Connection.ConnectionString = Core_Module.GetConstring
            DTA2.Connection.ConnectionString = Core_Module.GetConstring
            DTA3.Connection.ConnectionString = Core_Module.GetConstring

            'Dapatkan List Primary Attendance 
            Dta1.Fill(tbl1, vEndDate, vStartDate, vPrimaryEmp)
            For Each row1 In tbl1.Rows
                If ListPrimaryAttendance.Contains(row1.primary_transaksi) = False Then
                    ListPrimaryAttendance.Add(row1.primary_transaksi)
                End If
            Next

            'Masukin Message Box 
            dta.FillBy_PeriodEmp(tbl, vEndDate, vStartDate, vPrimaryEmp)
            If tbl.Rows.Count > 0 Then
                For Each row In tbl.Rows
                    ListPeriod = ListPeriod + i.ToString + ". " + row.no_transaksi.ToString + ", Period : " + _
                    Format(row.start_date, "dd MMM yyyy").ToString + " - " + Format(row.end_date, "dd MMM yyyy").ToString + vbNewLine
                    i = i + 1
                Next
            End If

            'DTA2.Fill(TBL2, vStartDate, vEndDate, vPrimaryEmp)
            'For Each PrimaryAttendance In ListPrimaryAttendance
            '    Filter = "[primary] = " + PrimaryAttendance.ToString + ""
            '    TBL2.DefaultView.RowFilter = Filter
            '    If TBL2.DefaultView.Count > 0 Then
            '    Else
            '        MunculkanMsgBox = True
            '        Exit For
            '    End If
            'Next

            If DTA3.ScalarQuery_GetJmlByDateBolong(vStartDate, vEndDate, vPrimaryEmp).HasValue Then
                JumlahDate = DTA3.ScalarQuery_GetJmlByDateBolong(vStartDate, vEndDate, vPrimaryEmp).Value
            Else
                JumlahDate = 0
            End If
            TS = vEndDate - vStartDate
            ParamHari = TS.Days + 1
            If JumlahDate <> ParamHari Then
                MunculkanMsgBox = True
            End If
            If MunculkanMsgBox = True Then
                If tbl1.Rows.Count <= 0 Then
                    If ShowMsg = True Then
                        MsgBox("This period salary have no attendance data : " + vbCrLf + Format(vStartDate, "dd MMM yyyy") + _
                        " - " + Format(vEndDate, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                Else
                    If ShowMsg = True Then
                        MsgBox("You can't use this period because this period isn't suitable with period attendance : " + vbNewLine + ListPeriod, MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            dta.Dispose()
            tbl.Dispose()
            Dta1.Dispose()
            tbl1.Dispose()
            DTA2.Dispose()
            TBL2.Dispose()
            DTA3.Dispose()
        End Try
    End Function

    Public Enum EnumPeriodePKP As Integer
        One_Month
        One_Week
        Two_Weeks
    End Enum

    Public Function GetPKPValue(ByVal vTanggalTransaksi As Date, ByVal vDateAwalThn As Date) As Decimal
        Dim ObjCd As New ClsCekData
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salaryTableAdapter
        Dim Dummy As System.Nullable(Of Decimal)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetZPKPValue(vDateAwalThn, vTanggalTransaksi)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCd = Nothing
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPPhValue(ByVal vTanggalTransaksi As Date, ByVal vDateAwalThn As Date) As Decimal
        Dim ObjCd As New ClsCekData
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salaryTableAdapter
        Dim Dummy As System.Nullable(Of Decimal)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetZPPhValue(vDateAwalThn, vTanggalTransaksi)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCd = Nothing
            DTA.Dispose()
        End Try
    End Function

    Public Sub IsiKeteranganKolomTabPage3(ByVal ij As FrmTrans_Gaji)
        Dim ObjCompany As New Core_Module.ClsMasterCompanyData
        Dim ObjKaryawan As New Core_Module.ClsMasterKaryawan
        Try
            ij.Txt_P_NPWPPemotong.Text = ObjCompany.GetCompany_NPWP()
            ij.Txt_P_NamaPemotong.Text = ObjCompany.GetCompanyName()
            ij.Txt_P_A_NPWP.Text = ObjKaryawan.GetNoNPWP(ij.TxtPrimaryEmployee.Text)
            ij.Txt_P_NoNIK.Text = ObjKaryawan.GetNIK(ij.TxtPrimaryEmployee.Text)
            ij.Txt_P_ANama.Text = ObjKaryawan.GetFullNameByPrimary(ij.TxtPrimaryEmployee.Text)
            ij.Txt_P_Alamat1.Text = ObjKaryawan.GetNPWPAddress(ij.TxtPrimaryEmployee.Text)
            ij.Txt_P_Alamat2.Text = ""
            If ObjKaryawan.GetSex(ij.TxtPrimaryEmployee.Text) = ClsMasterKaryawan.EnumSex.Male Then
                ij.Txt_P_JenisKelamin.Text = "LAKI-LAKI"
            Else
                ij.Txt_P_JenisKelamin.Text = "PEREMPUAN"
            End If
            If ObjKaryawan.GetNPWPCategory(ij.TxtPrimaryEmployee.Text) = ClsMasterKaryawan.GolonganPajak.TK0 Then
                ij.Txt_P_StatusNPWP.Text = "TK0"
            ElseIf ObjKaryawan.GetNPWPCategory(ij.TxtPrimaryEmployee.Text) = ClsMasterKaryawan.GolonganPajak.TK1 Then
                ij.Txt_P_StatusNPWP.Text = "TK1"
            ElseIf ObjKaryawan.GetNPWPCategory(ij.TxtPrimaryEmployee.Text) = ClsMasterKaryawan.GolonganPajak.TK2 Then
                ij.Txt_P_StatusNPWP.Text = "TK2"
            ElseIf ObjKaryawan.GetNPWPCategory(ij.TxtPrimaryEmployee.Text) = ClsMasterKaryawan.GolonganPajak.TK3 Then
                ij.Txt_P_StatusNPWP.Text = "TK3"
            ElseIf ObjKaryawan.GetNPWPCategory(ij.TxtPrimaryEmployee.Text) = ClsMasterKaryawan.GolonganPajak.K0 Then
                ij.Txt_P_StatusNPWP.Text = "K0"
            ElseIf ObjKaryawan.GetNPWPCategory(ij.TxtPrimaryEmployee.Text) = ClsMasterKaryawan.GolonganPajak.K1 Then
                ij.Txt_P_StatusNPWP.Text = "K1"
            ElseIf ObjKaryawan.GetNPWPCategory(ij.TxtPrimaryEmployee.Text) = ClsMasterKaryawan.GolonganPajak.K2 Then
                ij.Txt_P_StatusNPWP.Text = "K2"
            ElseIf ObjKaryawan.GetNPWPCategory(ij.TxtPrimaryEmployee.Text) = ClsMasterKaryawan.GolonganPajak.K3 Then
                ij.Txt_P_StatusNPWP.Text = "K3"
            End If
            ij.Txt_P_NamaJabatan.Text = ObjKaryawan.GetPositionDescription(ij.TxtPrimaryEmployee.Text)
            If ObjKaryawan.GetForeignEmployee(ij.TxtPrimaryEmployee.Text) = ClsMasterKaryawan.EnumForeign.Yes Then
                ij.Txt_P_KaryawanAsing.Text = "YA"
                ij.Txt_P_KodeNegaraDomisili.Text = ObjKaryawan.GetKodeNegara(ObjKaryawan.GetPrimaryCountry(ij.TxtPrimaryEmployee.Text))
            Else
                ij.Txt_P_KaryawanAsing.Text = "TIDAK"
                ij.Txt_P_KodeNegaraDomisili.Text = "INA"
            End If
            'sementara
            ij.Txt_P_NPWP.Text = ij.Txt_P_NPWPPemotong.Text
            ij.Txt_P_Nama.Text = ij.Txt_P_NamaPemotong.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCompany = Nothing
            ObjKaryawan = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' mengembalikan primary transaksi gaji pertama yang terjadi untuk karyawan ybs
    ''' dipake untuk memastikan faktor kali untuk perhitungan pajak disetahunkan
    ''' </summary>
    ''' <param name="vPrimaryKaryawan"></param>
    ''' <param name="vAdmissionDate"></param>
    ''' <param name="vEndDate"></param>
    ''' <param name="vPrimaryTransaksiAcuan"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetPrimary_TransaksiGajiPertama(ByVal vPrimaryKaryawan As Double, ByVal vAdmissionDate As Date, ByVal vEndDate As Date, _
    ByVal vPrimaryTransaksiAcuan As Double) As Double
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salaryTableAdapter
        Dim RetVal As Nullable(Of Double)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetPrimaryTransaksi1(vPrimaryKaryawan, vEndDate, vAdmissionDate, vPrimaryTransaksiAcuan)
            If RetVal.HasValue = True Then
                Return RetVal.Value
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

    Public Sub ClearObjectTxtPPh21(ByVal ij As FrmTrans_Gaji)
        Try
            ij.Txt01Pjk_Gaji.Text = "0"
            ij.Txt02PjkTunjanganPPh.Text = "0"
            ij.Txt03Pjk_TunjLainnya.Text = "0"
            ij.Txt04PjkHonorarium.Text = "0"
            ij.Txt05PjkPremiAsuransi.Text = "0"
            ij.Txt06PjkNatura.Text = "0"
            ij.Txt07PjkTantiem.Text = "0"
            ij.Txt08PjkPenghasilanBruto.Text = "0"
            ij.Txt09PjkBiayaJabatan.Text = "0"
            ij.Txt10PjkIuranPensiun.Text = "0"
            ij.Txt11PjkJmlPengurangan.Text = "0"
            ij.Txt12PjkJmlPenghasilanNetto.Text = "0"
            ij.Txt13PjkJmlHasilMsPjkSblm.Text = "0"
            ij.Txt14PjkJmlHasilNettoUPPh.Text = "0"
            ij.Txt15Pjk_PTKP.Text = "0"
            ij.Txt16Pjk_PKPSetahun.Text = "0"
            ij.Txt17Pjk_PPH21.Text = "0"
            ij.Txt18.Text = "0"
            ij.Txt19Pjk_PPH21Terutang.Text = "0"
            ij.Txt20.Text = "0"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Enum EnumHitungPPh21AkhirAtauBukan As Integer
        BukanAkhirTahun
        AkhirTahun
    End Enum

    'Pengecekan TBLPPh21
    Public Function PPh21_GetJumlah(ByVal vPenghasilanBrutoDiluarTunjangan As Decimal, ByVal vNo11JumlahPengurangan As Decimal, _
    ByVal vNo15PTKP As Decimal, ByVal vApakahAkhirTahun As EnumHitungPPh21AkhirAtauBukan) As Double
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salary_hitungpph21TableAdapter
        Dim Dummy As System.Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetJumlah(vPenghasilanBrutoDiluarTunjangan, vNo11JumlahPengurangan, vNo15PTKP, vApakahAkhirTahun)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    'Perhitungan TBLPPh21
    Public Sub PPh21_GetNilaiTblHitungPPh21(ByRef vNo17NilaiPPh21Value As Decimal, ByRef vNo19PPh21TerhutangValue As Decimal, _
    ByRef vNo2TunjanganPPh21Value As Decimal, ByVal vPenghasilanBrutoDiluarTunjangan As Decimal, ByVal vNo19PPh21Terhutang As Decimal, _
    ByVal vNo11JumlahPengurang As Decimal, ByVal vNo15PTKP As Decimal, ByVal vApakahAkhirTahun As EnumHitungPPh21AkhirAtauBukan)
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salary_hitungpph21_NilaiTableAdapter
        Dim TBL As New DseTransaksiSalary.tbltrans_salary_hitungpph21_NilaiDataTable
        Dim Row As DseTransaksiSalary.tbltrans_salary_hitungpph21_NilaiRow
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, vPenghasilanBrutoDiluarTunjangan, vNo19PPh21Terhutang, vNo11JumlahPengurang, vNo15PTKP, vApakahAkhirTahun)
            If TBL.Rows.Count > 0 Then
                Row = TBL.Rows(0)
                vNo17NilaiPPh21Value = Row._17PPh21
                vNo19PPh21TerhutangValue = Row._19PPh21Terhutang
                If vApakahAkhirTahun = EnumHitungPPh21AkhirAtauBukan.AkhirTahun Then
                    vNo2TunjanganPPh21Value = Row._02TunjanganPPh21
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    'Perhitungan Income Tax Yang Baru (Gross Up)
    'Public Function HitungIncomeTax_GrossUp_Bulan12(ByVal ij As FrmTrans_Gaji) As Decimal
    '    Dim ObjKrywn As New Core_Module.ClsMasterKaryawan
    '    'Parameter Hasil
    '    Dim vNo17NilaiPPh21Value As Decimal = 0
    '    Dim vNo19PPh21TerhutangValue As Decimal = 0
    '    Dim vNo2TunjanganPPh21Value As Decimal = 0
    '    'Parameter Fungsi
    '    Dim vPenghasilanBrutoDiluarTunjangan As Decimal
    '    Dim vNo19PPh21Terhutang As Decimal
    '    Dim vNo11JumlahPengurang As Decimal
    '    Dim vNo15PTKP As Decimal
    '    Dim vApakahAkhirTahun As EnumHitungPPh21AkhirAtauBukan
    '    'Stop watch
    '    Dim BatasanWaktu As Double = 300
    '    Dim SW As New Stopwatch
    '    '----------------
    '    Dim ObjCD As New Core_Module.ClsCekData
    '    Dim Bulan As Integer = 0
    '    Dim CekJmlPPh21 As Double = 0
    '    Dim Retval As Decimal = 0
    '    Try
    '        'Pertama --> Tunjangan PPh21 = 0, pakai fungsi yang bukan untuk Gross Up (ij, False, 0)
    '        Me.HitungIncomeTax_GrossNNetto_Bulan12(ij, False, 0)
    '        '
    '        vPenghasilanBrutoDiluarTunjangan = CDec(ij.Txt08PjkPenghasilanBruto.Text) - CDec(ij.Txt02PjkTunjanganPPh.Text)
    '        vNo19PPh21Terhutang = CDec(ij.Txt19Pjk_PPH21Terutang.Text)
    '        vNo11JumlahPengurang = CDec(ij.Txt11PjkJmlPengurangan.Text)
    '        vNo15PTKP = CDec(ij.Txt15Pjk_PTKP.Text)
    '        vApakahAkhirTahun = EnumHitungPPh21AkhirAtauBukan.AkhirTahun
    '        '
    '        CekJmlPPh21 = Me.PPh21_GetJumlah(vPenghasilanBrutoDiluarTunjangan, vNo11JumlahPengurang, vNo15PTKP, vApakahAkhirTahun)
    '        If CekJmlPPh21 > 0 Then
    '            Me.PPh21_GetNilaiTblHitungPPh21(vNo17NilaiPPh21Value, vNo19PPh21TerhutangValue, vNo2TunjanganPPh21Value, _
    '            vPenghasilanBrutoDiluarTunjangan, vNo19PPh21Terhutang, vNo11JumlahPengurang, vNo15PTKP, vApakahAkhirTahun)
    '            'Set ulang
    '            Retval = Me.SetUlang_PerhitungIncomeTax_GrossUp_Bulan12(ij, vNo17NilaiPPh21Value, vNo19PPh21TerhutangValue, vNo2TunjanganPPh21Value)
    '        Else
    '            If ObjKrywn.GetIsNPWP(ij.TxtPrimaryEmployee.Text) = ClsMasterKaryawan.EnumIsNPWP.No Then
    '                vNo19PPh21TerhutangValue = ObjCD.DigitGrouping(CDec(ij.Txt19Pjk_PPH21Terutang.Text) * 1.2, 0)
    '            Else
    '                vNo19PPh21TerhutangValue = CDec(ij.Txt19Pjk_PPH21Terutang.Text)
    '            End If
    '            'Get Tunjangan PPh 21, baru dengan konsep While
    '            SW.Start()
    '            While vNo2TunjanganPPh21Value <> vNo19PPh21TerhutangValue
    '                'If SW.Elapsed.TotalSeconds > BatasanWaktu Then
    '                '    MsgBox("Logic error, please contact your vendor.", MsgBoxStyle.Information, "Information")
    '                '    Exit While
    '                'Else
    '                Me.HitungIncomeTax_GrossNNetto_Bulan12(ij, True, vNo2TunjanganPPh21Value)
    '                vNo19PPh21TerhutangValue = CDec(ij.Txt19Pjk_PPH21Terutang.Text)
    '                If vNo19PPh21TerhutangValue > vNo2TunjanganPPh21Value Then
    '                    'If ((vNo19PPh21TerhutangValue - vNo2TunjanganPPh21Value) / 2) < 1 Then
    '                    '    Exit While
    '                    'Else
    '                    vNo2TunjanganPPh21Value = (((vNo19PPh21TerhutangValue - vNo2TunjanganPPh21Value) / 2) + vNo19PPh21TerhutangValue)
    '                    'End If
    '                ElseIf vNo19PPh21TerhutangValue < vNo2TunjanganPPh21Value Then
    '                    'If ((vNo2TunjanganPPh21Value - vNo19PPh21TerhutangValue) / 2) < 1 Then
    '                    '    Exit While
    '                    'Else
    '                    vNo2TunjanganPPh21Value = ((vNo2TunjanganPPh21Value - vNo19PPh21TerhutangValue) / 2)
    '                    vNo2TunjanganPPh21Value = vNo19PPh21TerhutangValue - vNo2TunjanganPPh21Value
    '                    'End If
    '                End If
    '                'End If
    '            End While
    '            ij.Txt02PjkTunjanganPPh.Text = ObjCD.DigitGrouping(vNo2TunjanganPPh21Value, 0)
    '            SW.Stop()
    '            Retval = CDec(ij.Txt19Pjk_PPH21Terutang.Text)
    '        End If

    '        Return Retval
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        ObjCD = Nothing
    '        ObjKrywn = Nothing
    '    End Try
    'End Function

    'Public Function SetUlang_PerhitungIncomeTax_GrossUp_Bulan12(ByVal ij As FrmTrans_Gaji, _
    'ByVal vNo17NilaiPPh21Value As Decimal, ByVal vNo19PPh21TerhutangValue As Decimal, _
    'ByVal vNo2TunjanganPPh21Value As Decimal) As Decimal
    '    'SUM PPh21 TblTrans_Gaji (nomor 1 ampe 20)
    '    Dim DTA As New DseTransaksiSalaryTableAdapters.DataTable6TableAdapter
    '    Dim TBL As New DseTransaksiSalary.DataTable6DataTable
    '    Dim Row As DseTransaksiSalary.DataTable6Row
    '    'SUM PPh21 TblTrans_Tax_Beginning (nomor 1 ampe 20)
    '    Dim DTA_Tax As New DseTransaksiSalaryTableAdapters.DataTable7TableAdapter
    '    Dim TBL_Tax As New DseTransaksiSalary.DataTable7DataTable
    '    Dim Row_Tax As DseTransaksiSalary.DataTable7Row
    '    Dim StartDate As Date = #1/1/1950#
    '    Dim EndDate As Date = #1/1/1950#
    '    Dim ObjCD As New ClsCekData
    '    Dim ObjEmployee As New ClsMasterKaryawan
    '    Dim ObjBeginningTax As New ClsTrans_Tax_Beginning_Balance
    '    Dim ObjIncomeTax As New ClsMasterSettingPajak
    '    Dim TipePajak As ClsMasterSettingPajak.EnumCriteria = ClsMasterSettingPajak.EnumCriteria.TK0
    '    Dim No09 As Decimal = 0
    '    Dim No15 As Decimal = 0
    '    Dim Retval As Decimal = 0
    '    Try
    '        DTA.Connection.ConnectionString = Core_Module.GetConstring
    '        DTA_Tax.Connection.ConnectionString = Core_Module.GetConstring

    '        Me.HitungIncomeTax_GrossNNetto_DariAwalSampe11(ij, True, vNo2TunjanganPPh21Value)
    '        StartDate = ObjCD.GetFisrtDateOfTheYear(ij.DtpDate.Value)
    '        EndDate = ObjCD.GetLasDateOfTheMonth(ObjCD.GetLastDateOfTheYear(ij.DtpDate.Value) - TimeSpan.FromDays(32))
    '        DTA.FillBy_StartEndDate(TBL, StartDate, EndDate, ij.TxtPrimaryEmployee.Text)
    '        DTA_Tax.Fill(TBL_Tax, EndDate, StartDate, ij.TxtPrimaryEmployee.Text)
    '        'Gabungkan Nilai dari salary + Tax_Beginning 
    '        If TBL_Tax.Rows.Count > 0 Then
    '            Row_Tax = TBL_Tax.Rows(0)
    '            If TBL.Rows.Count > 0 Then
    '                Row = TBL.Rows(0)
    '                For idx As Integer = 0 To TBL_Tax.Columns.Count - 1
    '                    Row.Item(idx) += Row_Tax.Item(idx)
    '                Next
    '            End If
    '        End If
    '        'Hitung Pennjumlahan
    '        If TBL.Rows.Count > 0 Then
    '            Row = TBL.Rows(0)
    '            ij.Txt01Pjk_Gaji.Text = ObjCD.DigitGrouping(CDec(ij.Txt01Pjk_Gaji.Text) + Row._01, 0)
    '            ij.Txt02PjkTunjanganPPh.Text = ObjCD.DigitGrouping(vNo2TunjanganPPh21Value, 0)
    '            ij.Txt03Pjk_TunjLainnya.Text = ObjCD.DigitGrouping(CDec(ij.Txt03Pjk_TunjLainnya.Text) + Row._03, 0)
    '            ij.Txt04PjkHonorarium.Text = ObjCD.DigitGrouping(CDec(ij.Txt04PjkHonorarium.Text) + Row._04, 0)
    '            ij.Txt05PjkPremiAsuransi.Text = ObjCD.DigitGrouping(CDec(ij.Txt05PjkPremiAsuransi.Text) + Row._05, 0)
    '            ij.Txt06PjkNatura.Text = ObjCD.DigitGrouping(CDec(ij.Txt06PjkNatura.Text) + Row._06, 0)
    '            ij.Txt07PjkTantiem.Text = ObjCD.DigitGrouping(CDec(ij.Txt07PjkTantiem.Text) + Row._07, 0)
    '            ij.Txt08PjkPenghasilanBruto.Text = ObjCD.DigitGrouping(CDec(ij.Txt08PjkPenghasilanBruto.Text) + Row._08, 0)
    '            No09 = 0.05 * CDec(ij.Txt08PjkPenghasilanBruto.Text)
    '            If No09 > 6000000 Then
    '                No09 = 6000000
    '            End If
    '            ij.Txt09PjkBiayaJabatan.Text = ObjCD.DigitGrouping(No09, 0)
    '            ij.Txt10PjkIuranPensiun.Text = ObjCD.DigitGrouping(CDec(ij.Txt10PjkIuranPensiun.Text) + Row._10, 0)
    '            ij.Txt11PjkJmlPengurangan.Text = ObjCD.DigitGrouping(Me.No11_HitungJumlahPengurang(CDec(ij.Txt09PjkBiayaJabatan.Text), CDec(ij.Txt10PjkIuranPensiun.Text)), 0)
    '            ij.Txt12PjkJmlPenghasilanNetto.Text = ObjCD.DigitGrouping(Me.No12_HItungJumlahPenghasilanNetto(CDec(ij.Txt08PjkPenghasilanBruto.Text), CDec(ij.Txt11PjkJmlPengurangan.Text)), 0)
    '            ij.Txt14PjkJmlHasilNettoUPPh.Text = ObjCD.DigitGrouping(ij.Txt12PjkJmlPenghasilanNetto.Text, 0)
    '            TipePajak = ObjEmployee.GetGolonganPajak(ij.TxtPrimaryEmployee.Text)
    '            No15 = ObjCD.DigitGrouping(ObjIncomeTax.GetPTKPPerTahun(TipePajak), 0)
    '            ij.Txt15Pjk_PTKP.Text = ObjCD.DigitGrouping(No15, 0)
    '            ij.Txt16Pjk_PKPSetahun.Text = ObjCD.DigitGrouping(Me.No16_HitungJumlahPenghasilanKenaPajakDiSetahunkan(CDec(ij.Txt14PjkJmlHasilNettoUPPh.Text), CDec(ij.Txt15Pjk_PTKP.Text)), 0)
    '            ij.Txt17Pjk_PPH21.Text = ObjCD.DigitGrouping(vNo17NilaiPPh21Value, 0)
    '            ij.Txt19Pjk_PPH21Terutang.Text = ObjCD.DigitGrouping(vNo19PPh21TerhutangValue, 0)
    '            Retval = CDec(ij.Txt19Pjk_PPH21Terutang.Text)
    '        End If

    '        Return Retval
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTA.Dispose()
    '        TBL.Dispose()
    '        DTA_Tax.Dispose()
    '        TBL_Tax.Dispose()
    '        ObjCD = Nothing
    '        ObjEmployee = Nothing
    '        ObjBeginningTax = Nothing
    '        ObjIncomeTax = Nothing
    '    End Try
    'End Function

    'Public Function HitungIncomeTax_GrossUp_DariAwalSampe11(ByVal ij As FrmTrans_Gaji) As Decimal
    '    Dim ObjKrywn As New Core_Module.ClsMasterKaryawan
    '    'Parameter Hasil
    '    Dim vNo17NilaiPPh21Value As Decimal = 0
    '    Dim vNo19PPh21TerhutangValue As Decimal = 0
    '    Dim vNo2TunjanganPPh21Value As Decimal = 0
    '    'Parameter Fungsi
    '    Dim vPenghasilanBrutoDiluarTunjangan As Decimal
    '    Dim vNo19PPh21Terhutang As Decimal
    '    Dim vNo11JumlahPengurang As Decimal
    '    Dim vNo15PTKP As Decimal
    '    Dim vApakahAkhirTahun As EnumHitungPPh21AkhirAtauBukan
    '    'Stop watch
    '    Dim BatasanWaktu As Double = 300
    '    Dim SW As New Stopwatch
    '    '----------------
    '    Dim ObjCD As New Core_Module.ClsCekData
    '    Dim Bulan As Integer = 0
    '    Dim CekJmlPPh21 As Double = 0
    '    Dim Retval As Decimal = 0
    '    Dim HarusHitungSekaliLagi As Boolean = False

    '    Try
    '        'Pertama --> Tunjangan PPh21 = 0, pakai fungsi yang bukan untuk Gross Up (ij, False, 0)
    '        Me.HitungIncomeTax_GrossNNetto_DariAwalSampe11(ij, False, 0)
    '        '
    '        vPenghasilanBrutoDiluarTunjangan = CDec(ij.Txt08PjkPenghasilanBruto.Text) - CDec(ij.Txt02PjkTunjanganPPh.Text)
    '        vNo19PPh21Terhutang = CDec(ij.Txt19Pjk_PPH21Terutang.Text)
    '        vNo11JumlahPengurang = CDec(ij.Txt11PjkJmlPengurangan.Text)
    '        vNo15PTKP = CDec(ij.Txt15Pjk_PTKP.Text)
    '        vApakahAkhirTahun = EnumHitungPPh21AkhirAtauBukan.BukanAkhirTahun
    '        '
    '        CekJmlPPh21 = Me.PPh21_GetJumlah(vPenghasilanBrutoDiluarTunjangan, vNo11JumlahPengurang, vNo15PTKP, vApakahAkhirTahun)
    '        If CekJmlPPh21 > 0 Then
    '            Me.PPh21_GetNilaiTblHitungPPh21(vNo17NilaiPPh21Value, vNo19PPh21TerhutangValue, vNo2TunjanganPPh21Value, _
    '            vPenghasilanBrutoDiluarTunjangan, vNo19PPh21Terhutang, vNo11JumlahPengurang, vNo15PTKP, vApakahAkhirTahun)
    '            'Set ulang
    '            Retval = Me.SetUlang_PerhitungIncomeTax_GrossUp_DariAwalSampe11(ij, vNo17NilaiPPh21Value, vNo19PPh21TerhutangValue, vNo2TunjanganPPh21Value)
    '        Else
    '            If ObjKrywn.GetIsNPWP(ij.TxtPrimaryEmployee.Text) = ClsMasterKaryawan.EnumIsNPWP.No Then
    '                vNo19PPh21TerhutangValue = ObjCD.DigitGrouping(CDec(ij.Txt19Pjk_PPH21Terutang.Text) * 1.2, 0)
    '            Else
    '                vNo19PPh21TerhutangValue = CDec(ij.Txt19Pjk_PPH21Terutang.Text)
    '            End If
    '            'Get Tunjangan PPh 21, baru dengan konsep While
    '            SW.Start()
    '            While vNo2TunjanganPPh21Value <> vNo19PPh21TerhutangValue
    '                If SW.Elapsed.TotalSeconds > BatasanWaktu Then
    '                    MsgBox("Logic error, please contact your vendor.", MsgBoxStyle.Information, "Information")
    '                    Exit While
    '                Else
    '                    Me.HitungIncomeTax_GrossNNetto_DariAwalSampe11(ij, True, vNo2TunjanganPPh21Value)
    '                    vNo19PPh21TerhutangValue = CDec(ij.Txt19Pjk_PPH21Terutang.Text)
    '                    If vNo19PPh21TerhutangValue > vNo2TunjanganPPh21Value Then
    '                        'If ((vNo19PPh21TerhutangValue - vNo2TunjanganPPh21Value) / 2) < 1 Then
    '                        '    Exit While
    '                        'Else
    '                        vNo2TunjanganPPh21Value = Math.Round((((vNo19PPh21TerhutangValue - vNo2TunjanganPPh21Value) / 2) + vNo19PPh21TerhutangValue), 0, MidpointRounding.AwayFromZero)
    '                        'End If
    '                    ElseIf vNo19PPh21TerhutangValue < vNo2TunjanganPPh21Value Then
    '                        'If ((vNo2TunjanganPPh21Value - vNo19PPh21TerhutangValue) / 2) < 1 Then
    '                        '    Exit While
    '                        'Else
    '                        vNo2TunjanganPPh21Value = ((vNo2TunjanganPPh21Value - vNo19PPh21TerhutangValue) / 2)
    '                        vNo2TunjanganPPh21Value = Math.Round(vNo19PPh21TerhutangValue - vNo2TunjanganPPh21Value, 0, MidpointRounding.AwayFromZero)
    '                        'End If
    '                    End If
    '                End If
    '            End While
    '            'Update Coding By Kleren - 20150406
    '            If ij.Txt02PjkTunjanganPPh.Text <> vNo2TunjanganPPh21Value Then
    '                HarusHitungSekaliLagi = True
    '            End If
    '            ij.Txt02PjkTunjanganPPh.Text = vNo2TunjanganPPh21Value
    '            If HarusHitungSekaliLagi = True Then
    '                Me.HitungIncomeTax_GrossNNetto_DariAwalSampe11(ij, True, vNo2TunjanganPPh21Value)
    '            End If

    '            SW.Stop()
    '            Retval = CDec(ij.Txt19Pjk_PPH21Terutang.Text)
    '        End If

    '        Return Retval
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        ObjCD = Nothing
    '        ObjKrywn = Nothing
    '    End Try
    'End Function

    'Private Function SetUlang_PerhitungIncomeTax_GrossUp_DariAwalSampe11(ByVal ij As FrmTrans_Gaji, _
    'ByVal vNo17NilaiPPh21Value As Decimal, ByVal vNo19PPh21TerhutangValue As Decimal, _
    'ByVal vNo2TunjanganPPh21Value As Decimal) As Decimal
    '    Dim RowDGV As DataGridViewRow
    '    Dim CboDgv As DataGridViewComboBoxCell
    '    Dim CboDgvType As DataGridViewComboBoxCell
    '    Dim ObjIncomeTax As New ClsMasterSettingPajak
    '    Dim ObjEmployee As New ClsMasterKaryawan
    '    Dim ObjBeginningTax As New ClsTrans_Tax_Beginning_Balance
    '    Dim ObjCD As New Core_Module.ClsCekData
    '    Dim TipePajak As ClsMasterSettingPajak.EnumCriteria = ClsMasterSettingPajak.EnumCriteria.TK0
    '    Dim Retval As Decimal = 0
    '    Dim No1 As Decimal = 0
    '    Dim No3 As Decimal = 0
    '    Dim No4 As Decimal = 0
    '    Dim No5 As Decimal = 0
    '    Dim No6 As Decimal = 0
    '    Dim No7 As Decimal = 0
    '    Dim No8 As Decimal = 0
    '    Dim No9 As Decimal = 0
    '    Dim No10 As Decimal = 0
    '    Dim No11 As Decimal = 0
    '    Dim No12 As Decimal = 0
    '    Dim No14 As Decimal = 0
    '    Dim No15 As Decimal = 0
    '    Dim No16 As Decimal = 0
    '    Dim No17 As Decimal = 0
    '    Dim No19 As Decimal = 0
    '    Try
    '        If ij.dgvItem.Rows.Count > 0 Then
    '            For Each RowDGV In ij.dgvItem.Rows
    '                CboDgv = RowDGV.Cells(ij.ColTaxElement.Index)
    '                CboDgvType = RowDGV.Cells(ij.ColTipeKomponen.Index)
    '                If CboDgv.Items.IndexOf(RowDGV.Cells(ij.ColTaxElement.Index).Value) = ClsTransaksiSalary_Item.EnumTaxElement.No1 Then
    '                    If CboDgvType.Items.IndexOf(RowDGV.Cells(ij.ColTipeKomponen.Index).Value) = ClsTransaksiSalary_Item.EnumType.Income Then
    '                        No1 += CDec(RowDGV.Cells(ij.ColAmount.Index).Value)
    '                    ElseIf CboDgvType.Items.IndexOf(RowDGV.Cells(ij.ColTipeKomponen.Index).Value) = ClsTransaksiSalary_Item.EnumType.Deduction Then
    '                        No1 -= CDec(RowDGV.Cells(ij.ColAmount.Index).Value)
    '                    End If
    '                ElseIf CboDgv.Items.IndexOf(RowDGV.Cells(ij.ColTaxElement.Index).Value) = ClsTransaksiSalary_Item.EnumTaxElement.No3 Then
    '                    If CboDgvType.Items.IndexOf(RowDGV.Cells(ij.ColTipeKomponen.Index).Value) = ClsTransaksiSalary_Item.EnumType.Income Then
    '                        No3 += CDec(RowDGV.Cells(ij.ColAmount.Index).Value)
    '                    ElseIf CboDgvType.Items.IndexOf(RowDGV.Cells(ij.ColTipeKomponen.Index).Value) = ClsTransaksiSalary_Item.EnumType.Deduction Then
    '                        No3 -= CDec(RowDGV.Cells(ij.ColAmount.Index).Value)
    '                    End If
    '                ElseIf CboDgv.Items.IndexOf(RowDGV.Cells(ij.ColTaxElement.Index).Value) = ClsTransaksiSalary_Item.EnumTaxElement.No4 Then
    '                    If CboDgvType.Items.IndexOf(RowDGV.Cells(ij.ColTipeKomponen.Index).Value) = ClsTransaksiSalary_Item.EnumType.Income Then
    '                        No4 += CDec(RowDGV.Cells(ij.ColAmount.Index).Value)
    '                    ElseIf CboDgvType.Items.IndexOf(RowDGV.Cells(ij.ColTipeKomponen.Index).Value) = ClsTransaksiSalary_Item.EnumType.Deduction Then
    '                        No4 -= CDec(RowDGV.Cells(ij.ColAmount.Index).Value)
    '                    End If
    '                ElseIf CboDgv.Items.IndexOf(RowDGV.Cells(ij.ColTaxElement.Index).Value) = ClsTransaksiSalary_Item.EnumTaxElement.No5 Then
    '                    If CboDgvType.Items.IndexOf(RowDGV.Cells(ij.ColTipeKomponen.Index).Value) = ClsTransaksiSalary_Item.EnumType.Income Then
    '                        No5 += CDec(RowDGV.Cells(ij.ColAmount.Index).Value)
    '                    ElseIf CboDgvType.Items.IndexOf(RowDGV.Cells(ij.ColTipeKomponen.Index).Value) = ClsTransaksiSalary_Item.EnumType.Deduction Then
    '                        No5 -= CDec(RowDGV.Cells(ij.ColAmount.Index).Value)
    '                    End If
    '                ElseIf CboDgv.Items.IndexOf(RowDGV.Cells(ij.ColTaxElement.Index).Value) = ClsTransaksiSalary_Item.EnumTaxElement.No6 Then
    '                    If CboDgvType.Items.IndexOf(RowDGV.Cells(ij.ColTipeKomponen.Index).Value) = ClsTransaksiSalary_Item.EnumType.Income Then
    '                        No6 += CDec(RowDGV.Cells(ij.ColAmount.Index).Value)
    '                    ElseIf CboDgvType.Items.IndexOf(RowDGV.Cells(ij.ColTipeKomponen.Index).Value) = ClsTransaksiSalary_Item.EnumType.Deduction Then
    '                        No6 -= CDec(RowDGV.Cells(ij.ColAmount.Index).Value)
    '                    End If
    '                ElseIf CboDgv.Items.IndexOf(RowDGV.Cells(ij.ColTaxElement.Index).Value) = ClsTransaksiSalary_Item.EnumTaxElement.No7 Then
    '                    If CboDgvType.Items.IndexOf(RowDGV.Cells(ij.ColTipeKomponen.Index).Value) = ClsTransaksiSalary_Item.EnumType.Income Then
    '                        No7 += CDec(RowDGV.Cells(ij.ColAmount.Index).Value)
    '                    ElseIf CboDgvType.Items.IndexOf(RowDGV.Cells(ij.ColTipeKomponen.Index).Value) = ClsTransaksiSalary_Item.EnumType.Deduction Then
    '                        No7 -= CDec(RowDGV.Cells(ij.ColAmount.Index).Value)
    '                    End If
    '                ElseIf CboDgv.Items.IndexOf(RowDGV.Cells(ij.ColTaxElement.Index).Value) = ClsTransaksiSalary_Item.EnumTaxElement.No10 Then
    '                    No10 -= CDec(RowDGV.Cells(ij.ColAmount.Index).Value)
    '                End If
    '            Next
    '        End If

    '        No8 = Me.No8_HitungJumlahPenghasilanBruto(No1, vNo2TunjanganPPh21Value, No3, No4, No5, No6, No7)
    '        No9 = 0.05 * No8
    '        If No9 > 500000 Then
    '            No9 = 500000
    '        End If
    '        No11 = Me.No11_HitungJumlahPengurang(No9, No10)
    '        No12 = Me.No12_HItungJumlahPenghasilanNetto(No8, No11)
    '        No14 = No12 * 12
    '        TipePajak = ObjEmployee.GetGolonganPajak(ij.TxtPrimaryEmployee.Text)
    '        No15 = ObjIncomeTax.GetPTKPPerTahun(TipePajak)
    '        No16 = Me.No16_HitungJumlahPenghasilanKenaPajakDiSetahunkan(No14, No15)
    '        No17 = vNo17NilaiPPh21Value
    '        No19 = vNo19PPh21TerhutangValue
    '        'Input ke text box
    '        ij.Txt01Pjk_Gaji.Text = ObjCD.DigitGrouping(No1, 0)
    '        ij.Txt02PjkTunjanganPPh.Text = ObjCD.DigitGrouping(vNo2TunjanganPPh21Value, 0)
    '        ij.Txt03Pjk_TunjLainnya.Text = ObjCD.DigitGrouping(No3, 0)
    '        ij.Txt04PjkHonorarium.Text = ObjCD.DigitGrouping(No4, 0)
    '        ij.Txt05PjkPremiAsuransi.Text = ObjCD.DigitGrouping(No5, 0)
    '        ij.Txt06PjkNatura.Text = ObjCD.DigitGrouping(No6, 0)
    '        ij.Txt07PjkTantiem.Text = ObjCD.DigitGrouping(No7, 0)
    '        ij.Txt08PjkPenghasilanBruto.Text = ObjCD.DigitGrouping(No8, 0)
    '        ij.Txt09PjkBiayaJabatan.Text = ObjCD.DigitGrouping(No9, 0)
    '        ij.Txt10PjkIuranPensiun.Text = ObjCD.DigitGrouping(No10, 0)
    '        ij.Txt11PjkJmlPengurangan.Text = ObjCD.DigitGrouping(No11, 0)
    '        ij.Txt12PjkJmlPenghasilanNetto.Text = ObjCD.DigitGrouping(No12, 0)
    '        ij.Txt13PjkJmlHasilMsPjkSblm.Text = 0
    '        ij.Txt14PjkJmlHasilNettoUPPh.Text = ObjCD.DigitGrouping(No14, 0)
    '        ij.Txt15Pjk_PTKP.Text = ObjCD.DigitGrouping(No15, 0)
    '        ij.Txt16Pjk_PKPSetahun.Text = ObjCD.DigitGrouping(No16, 0)
    '        ij.Txt17Pjk_PPH21.Text = ObjCD.DigitGrouping(No17, 0)
    '        ij.Txt18.Text = "0"
    '        ij.Txt19Pjk_PPH21Terutang.Text = ObjCD.DigitGrouping(No19, 0)
    '        ij.Txt20.Text = "0"
    '        Retval = CDec(ij.Txt19Pjk_PPH21Terutang.Text)

    '        Return Retval
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        ObjIncomeTax = Nothing
    '        ObjEmployee = Nothing
    '        ObjBeginningTax = Nothing
    '        ObjCD = Nothing
    '    End Try
    'End Function

    'Perhitungan Income Tax Yang Baru (Gross dan Netto)
    'Public Function HitungIncomeTax_GrossNNetto_Bulan12(ByVal ij As FrmTrans_Gaji, _
    'Optional ByVal vApakahUntukGrossUp As Boolean = False, Optional ByVal vNo2TunjanganPPh As Decimal = 0) As Decimal
    '    'SUM PPh21 TblTrans_Gaji (nomor 1 ampe 20)
    '    Dim DTA As New DseTransaksiSalaryTableAdapters.DataTable6TableAdapter
    '    Dim TBL As New DseTransaksiSalary.DataTable6DataTable
    '    Dim Row As DseTransaksiSalary.DataTable6Row
    '    '---Hitung jumlah bayar dari bulan 1 ampe 11---
    '    Dim IncomeTaxSampaiBulan11 As Decimal = 0
    '    'SUM PPh21 TblTrans_Tax_Beginning (nomor 1 ampe 20)
    '    Dim DTA_Tax As New DseTransaksiSalaryTableAdapters.DataTable7TableAdapter
    '    Dim TBL_Tax As New DseTransaksiSalary.DataTable7DataTable
    '    Dim Row_Tax As DseTransaksiSalary.DataTable7Row
    '    Dim StartDate As Date = #1/1/1950#
    '    Dim EndDate As Date = #1/1/1950#
    '    Dim ObjCD As New ClsCekData
    '    Dim ObjEmployee As New ClsMasterKaryawan
    '    Dim ObjBeginningTax As New ClsTrans_Tax_Beginning_Balance
    '    Dim ObjIncomeTax As New ClsMasterSettingPajak
    '    Dim ObjKrywn As New Core_Module.ClsMasterKaryawan
    '    Dim TipePajak As ClsMasterSettingPajak.EnumCriteria = ClsMasterSettingPajak.EnumCriteria.TK0
    '    Dim No09 As Decimal = 0
    '    Dim No15 As Decimal = 0
    '    Dim Retval As Decimal = 0
    '    Try
    '        DTA.Connection.ConnectionString = Core_Module.GetConstring
    '        DTA_Tax.Connection.ConnectionString = Core_Module.GetConstring

    '        If vApakahUntukGrossUp = False Then
    '            Me.HitungIncomeTax_GrossNNetto_DariAwalSampe11(ij)
    '        Else
    '            Me.HitungIncomeTax_GrossNNetto_DariAwalSampe11(ij, True, vNo2TunjanganPPh)
    '        End If
    '        StartDate = ObjCD.GetFisrtDateOfTheYear(ij.DtpDate.Value)
    '        EndDate = ObjCD.GetLasDateOfTheMonth(ObjCD.GetLastDateOfTheYear(ij.DtpDate.Value) - TimeSpan.FromDays(32))
    '        DTA.FillBy_StartEndDate(TBL, StartDate, EndDate, ij.TxtPrimaryEmployee.Text)
    '        'Hitung Income Tax Yang Pernah di bayar dari Bulan 1 ampe 11
    '        If TBL.Rows.Count > 0 Then
    '            Row = TBL.Rows(0)
    '            IncomeTaxSampaiBulan11 = Row._19
    '        End If
    '        DTA_Tax.Fill(TBL_Tax, EndDate, StartDate, ij.TxtPrimaryEmployee.Text)
    '        'Gabungkan Nilai dari salary + Tax_Beginning 
    '        If TBL_Tax.Rows.Count > 0 Then
    '            Row_Tax = TBL_Tax.Rows(0)
    '            If TBL.Rows.Count > 0 Then
    '                Row = TBL.Rows(0)
    '                For idx As Integer = 0 To TBL_Tax.Columns.Count - 1
    '                    Row.Item(idx) += Row_Tax.Item(idx)
    '                Next
    '            End If
    '        End If
    '        'Lakukan Penjumlahan
    '        If TBL.Rows.Count > 0 Then
    '            Row = TBL.Rows(0)
    '            ij.Txt01Pjk_Gaji.Text = ObjCD.DigitGrouping(CDec(ij.Txt01Pjk_Gaji.Text) + Row._01, 0)
    '            ij.Txt02PjkTunjanganPPh.Text = ObjCD.DigitGrouping(Row._02, 0)
    '            ij.Txt03Pjk_TunjLainnya.Text = ObjCD.DigitGrouping(CDec(ij.Txt03Pjk_TunjLainnya.Text) + Row._03, 0)
    '            ij.Txt04PjkHonorarium.Text = ObjCD.DigitGrouping(CDec(ij.Txt04PjkHonorarium.Text) + Row._04, 0)
    '            ij.Txt05PjkPremiAsuransi.Text = ObjCD.DigitGrouping(CDec(ij.Txt05PjkPremiAsuransi.Text) + Row._05, 0)
    '            ij.Txt06PjkNatura.Text = ObjCD.DigitGrouping(CDec(ij.Txt06PjkNatura.Text) + Row._06, 0)
    '            ij.Txt07PjkTantiem.Text = ObjCD.DigitGrouping(CDec(ij.Txt07PjkTantiem.Text) + Row._07, 0)
    '            ij.Txt08PjkPenghasilanBruto.Text = ObjCD.DigitGrouping(CDec(ij.Txt08PjkPenghasilanBruto.Text) + Row._01 + Row._03 + Row._04 + Row._05 + Row._06 + Row._07, 0)

    '            No09 = 0.05 * CDec(ij.Txt08PjkPenghasilanBruto.Text)
    '            If No09 > 6000000 Then
    '                No09 = 6000000
    '            End If
    '            ij.Txt09PjkBiayaJabatan.Text = ObjCD.DigitGrouping(No09, 0)
    '            ij.Txt10PjkIuranPensiun.Text = ObjCD.DigitGrouping(CDec(ij.Txt10PjkIuranPensiun.Text) + Row._10, 0)
    '            ij.Txt11PjkJmlPengurangan.Text = ObjCD.DigitGrouping(Me.No11_HitungJumlahPengurang(CDec(ij.Txt09PjkBiayaJabatan.Text), CDec(ij.Txt10PjkIuranPensiun.Text)), 0)
    '            ij.Txt12PjkJmlPenghasilanNetto.Text = ObjCD.DigitGrouping(Me.No12_HItungJumlahPenghasilanNetto(CDec(ij.Txt08PjkPenghasilanBruto.Text), CDec(ij.Txt11PjkJmlPengurangan.Text)), 0)
    '            ij.Txt14PjkJmlHasilNettoUPPh.Text = ObjCD.DigitGrouping(ij.Txt12PjkJmlPenghasilanNetto.Text, 0)
    '            TipePajak = ObjEmployee.GetGolonganPajak(ij.TxtPrimaryEmployee.Text)
    '            No15 = ObjCD.DigitGrouping(ObjIncomeTax.GetPTKPPerTahun(TipePajak), 0)
    '            ij.Txt15Pjk_PTKP.Text = ObjCD.DigitGrouping(No15, 0)
    '            ij.Txt16Pjk_PKPSetahun.Text = ObjCD.DigitGrouping(Me.No16_HitungJumlahPenghasilanKenaPajakDiSetahunkan(CDec(ij.Txt14PjkJmlHasilNettoUPPh.Text), CDec(ij.Txt15Pjk_PTKP.Text)), 0)
    '            If vApakahUntukGrossUp = False Then
    '                If ObjKrywn.GetIsNPWP(ij.TxtPrimaryEmployee.Text) = ClsMasterKaryawan.EnumIsNPWP.No Then
    '                    ij.Txt17Pjk_PPH21.Text = ObjCD.DigitGrouping(ObjBeginningTax.HitungPKP(CDec(ij.Txt16Pjk_PKPSetahun.Text)) * 1.2, 0)
    '                    ij.Txt19Pjk_PPH21Terutang.Text = ij.Txt17Pjk_PPH21.Text
    '                Else
    '                    ij.Txt17Pjk_PPH21.Text = ObjCD.DigitGrouping(ObjBeginningTax.HitungPKP(CDec(ij.Txt16Pjk_PKPSetahun.Text)), 0)
    '                    ij.Txt19Pjk_PPH21Terutang.Text = ij.Txt17Pjk_PPH21.Text
    '                End If
    '            Else
    '                ij.Txt17Pjk_PPH21.Text = ObjCD.DigitGrouping(ObjBeginningTax.HitungPKP(CDec(ij.Txt16Pjk_PKPSetahun.Text)), 0)
    '                ij.Txt19Pjk_PPH21Terutang.Text = ij.Txt17Pjk_PPH21.Text
    '            End If
    '            If CDec(ij.Txt19Pjk_PPH21Terutang.Text) > IncomeTaxSampaiBulan11 Then
    '                ij.TxtKurangBayarTax.Text = ObjCD.DigitGrouping(CDec(ij.Txt19Pjk_PPH21Terutang.Text) - IncomeTaxSampaiBulan11, 0)
    '            Else
    '                ij.TxtKurangBayarTax.Text = "0"
    '            End If
    '            Retval = CDec(ij.Txt19Pjk_PPH21Terutang.Text)
    '        End If

    '        Return Retval
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTA.Dispose()
    '        TBL.Dispose()
    '        DTA_Tax.Dispose()
    '        TBL_Tax.Dispose()
    '        ObjCD = Nothing
    '        ObjEmployee = Nothing
    '        ObjBeginningTax = Nothing
    '        ObjIncomeTax = Nothing
    '        ObjKrywn = Nothing
    '    End Try
    'End Function

    'Public Function HitungIncomeTax_GrossNNetto_DariAwalSampe11(ByVal ij As FrmTrans_Gaji, _
    'Optional ByVal vApakahUntukGrossUp As Boolean = False, Optional ByVal vNo2TunjanganPPh As Decimal = 0) As Decimal
    '    Dim RowDGV As DataGridViewRow
    '    Dim CboDgv As DataGridViewComboBoxCell
    '    Dim CboDgvType As DataGridViewComboBoxCell
    '    Dim ObjIncomeTax As New ClsMasterSettingPajak
    '    Dim ObjEmployee As New ClsMasterKaryawan
    '    Dim ObjBeginningTax As New ClsTrans_Tax_Beginning_Balance
    '    Dim ObjCD As New Core_Module.ClsCekData
    '    Dim ObjKrywn As New Core_Module.ClsMasterKaryawan
    '    Dim TipePajak As ClsMasterSettingPajak.EnumCriteria = ClsMasterSettingPajak.EnumCriteria.TK0
    '    Dim Retval As Decimal = 0
    '    Dim No1 As Decimal = 0
    '    Dim No3 As Decimal = 0
    '    Dim No4 As Decimal = 0
    '    Dim No5 As Decimal = 0
    '    Dim No6 As Decimal = 0
    '    Dim No7 As Decimal = 0
    '    Dim No8 As Decimal = 0
    '    Dim No9 As Decimal = 0
    '    Dim No10 As Decimal = 0
    '    Dim No11 As Decimal = 0
    '    Dim No12 As Decimal = 0
    '    Dim No14 As Decimal = 0
    '    Dim No15 As Decimal = 0
    '    Dim No16 As Decimal = 0
    '    Dim No17 As Decimal = 0
    '    Dim No19 As Decimal = 0
    '    Try
    '        If ij.dgvItem.Rows.Count > 0 Then
    '            For Each RowDGV In ij.dgvItem.Rows
    '                CboDgv = RowDGV.Cells(ij.ColTaxElement.Index)
    '                CboDgvType = RowDGV.Cells(ij.ColTipeKomponen.Index)
    '                If CboDgv.Items.IndexOf(RowDGV.Cells(ij.ColTaxElement.Index).Value) = ClsTransaksiSalary_Item.EnumTaxElement.No1 Then
    '                    If CboDgvType.Items.IndexOf(RowDGV.Cells(ij.ColTipeKomponen.Index).Value) = ClsTransaksiSalary_Item.EnumType.Income Then
    '                        No1 += CDec(RowDGV.Cells(ij.ColAmount.Index).Value)
    '                    ElseIf CboDgvType.Items.IndexOf(RowDGV.Cells(ij.ColTipeKomponen.Index).Value) = ClsTransaksiSalary_Item.EnumType.Deduction Then
    '                        No1 -= CDec(RowDGV.Cells(ij.ColAmount.Index).Value)
    '                    End If
    '                ElseIf CboDgv.Items.IndexOf(RowDGV.Cells(ij.ColTaxElement.Index).Value) = ClsTransaksiSalary_Item.EnumTaxElement.No3 Then
    '                    If CboDgvType.Items.IndexOf(RowDGV.Cells(ij.ColTipeKomponen.Index).Value) = ClsTransaksiSalary_Item.EnumType.Income Then
    '                        No3 += CDec(RowDGV.Cells(ij.ColAmount.Index).Value)
    '                    ElseIf CboDgvType.Items.IndexOf(RowDGV.Cells(ij.ColTipeKomponen.Index).Value) = ClsTransaksiSalary_Item.EnumType.Deduction Then
    '                        No3 -= CDec(RowDGV.Cells(ij.ColAmount.Index).Value)
    '                    End If
    '                ElseIf CboDgv.Items.IndexOf(RowDGV.Cells(ij.ColTaxElement.Index).Value) = ClsTransaksiSalary_Item.EnumTaxElement.No4 Then
    '                    If CboDgvType.Items.IndexOf(RowDGV.Cells(ij.ColTipeKomponen.Index).Value) = ClsTransaksiSalary_Item.EnumType.Income Then
    '                        No4 += CDec(RowDGV.Cells(ij.ColAmount.Index).Value)
    '                    ElseIf CboDgvType.Items.IndexOf(RowDGV.Cells(ij.ColTipeKomponen.Index).Value) = ClsTransaksiSalary_Item.EnumType.Deduction Then
    '                        No4 -= CDec(RowDGV.Cells(ij.ColAmount.Index).Value)
    '                    End If
    '                ElseIf CboDgv.Items.IndexOf(RowDGV.Cells(ij.ColTaxElement.Index).Value) = ClsTransaksiSalary_Item.EnumTaxElement.No5 Then
    '                    If CboDgvType.Items.IndexOf(RowDGV.Cells(ij.ColTipeKomponen.Index).Value) = ClsTransaksiSalary_Item.EnumType.Income Then
    '                        No5 += CDec(RowDGV.Cells(ij.ColAmount.Index).Value)
    '                    ElseIf CboDgvType.Items.IndexOf(RowDGV.Cells(ij.ColTipeKomponen.Index).Value) = ClsTransaksiSalary_Item.EnumType.Deduction Then
    '                        No5 -= CDec(RowDGV.Cells(ij.ColAmount.Index).Value)
    '                    End If
    '                ElseIf CboDgv.Items.IndexOf(RowDGV.Cells(ij.ColTaxElement.Index).Value) = ClsTransaksiSalary_Item.EnumTaxElement.No6 Then
    '                    If CboDgvType.Items.IndexOf(RowDGV.Cells(ij.ColTipeKomponen.Index).Value) = ClsTransaksiSalary_Item.EnumType.Income Then
    '                        No6 += CDec(RowDGV.Cells(ij.ColAmount.Index).Value)
    '                    ElseIf CboDgvType.Items.IndexOf(RowDGV.Cells(ij.ColTipeKomponen.Index).Value) = ClsTransaksiSalary_Item.EnumType.Deduction Then
    '                        No6 -= CDec(RowDGV.Cells(ij.ColAmount.Index).Value)
    '                    End If
    '                ElseIf CboDgv.Items.IndexOf(RowDGV.Cells(ij.ColTaxElement.Index).Value) = ClsTransaksiSalary_Item.EnumTaxElement.No7 Then
    '                    If CboDgvType.Items.IndexOf(RowDGV.Cells(ij.ColTipeKomponen.Index).Value) = ClsTransaksiSalary_Item.EnumType.Income Then
    '                        No7 += CDec(RowDGV.Cells(ij.ColAmount.Index).Value)
    '                    ElseIf CboDgvType.Items.IndexOf(RowDGV.Cells(ij.ColTipeKomponen.Index).Value) = ClsTransaksiSalary_Item.EnumType.Deduction Then
    '                        No7 -= CDec(RowDGV.Cells(ij.ColAmount.Index).Value)
    '                    End If
    '                ElseIf CboDgv.Items.IndexOf(RowDGV.Cells(ij.ColTaxElement.Index).Value) = ClsTransaksiSalary_Item.EnumTaxElement.No10 Then
    '                    No10 -= CDec(RowDGV.Cells(ij.ColAmount.Index).Value)
    '                End If
    '            Next
    '        End If
    '        If vApakahUntukGrossUp = False Then
    '            No8 = Me.No8_HitungJumlahPenghasilanBruto(No1, 0, No3, No4, No5, No6, No7)
    '        Else
    '            No8 = Me.No8_HitungJumlahPenghasilanBruto(No1, vNo2TunjanganPPh, No3, No4, No5, No6, No7)
    '        End If
    '        No9 = 0.05 * No8
    '        If No9 > 500000 Then
    '            No9 = 500000
    '        End If
    '        No11 = Me.No11_HitungJumlahPengurang(No9, No10)
    '        No12 = Me.No12_HItungJumlahPenghasilanNetto(No8, No11)
    '        If vApakahUntukGrossUp = False Then
    '            No14 = No12 * 12
    '        Else
    '            No14 = ObjCD.DigitGrouping(No12, 0) * 12
    '        End If
    '        TipePajak = ObjEmployee.GetGolonganPajak(ij.TxtPrimaryEmployee.Text)
    '        No15 = ObjIncomeTax.GetPTKPPerTahun(TipePajak)
    '        No16 = Me.No16_HitungJumlahPenghasilanKenaPajakDiSetahunkan(No14, No15)
    '        No17 = ObjBeginningTax.HitungPKP(No16)
    '        No19 = No17 / 12
    '        'Input ke text box
    '        ij.Txt01Pjk_Gaji.Text = ObjCD.DigitGrouping(No1, 0)
    '        If vApakahUntukGrossUp = False Then
    '            ij.Txt02PjkTunjanganPPh.Text = "0"
    '        Else
    '            ij.Txt02PjkTunjanganPPh.Text = ObjCD.DigitGrouping(vNo2TunjanganPPh, 0)
    '        End If
    '        ij.Txt03Pjk_TunjLainnya.Text = ObjCD.DigitGrouping(No3, 0)
    '        ij.Txt04PjkHonorarium.Text = ObjCD.DigitGrouping(No4, 0)
    '        ij.Txt05PjkPremiAsuransi.Text = ObjCD.DigitGrouping(No5, 0)
    '        ij.Txt06PjkNatura.Text = ObjCD.DigitGrouping(No6, 0)
    '        ij.Txt07PjkTantiem.Text = ObjCD.DigitGrouping(No7, 0)
    '        ij.Txt08PjkPenghasilanBruto.Text = ObjCD.DigitGrouping(No8, 0)
    '        ij.Txt09PjkBiayaJabatan.Text = ObjCD.DigitGrouping(No9, 0)
    '        ij.Txt10PjkIuranPensiun.Text = ObjCD.DigitGrouping(No10, 0)
    '        ij.Txt11PjkJmlPengurangan.Text = ObjCD.DigitGrouping(No11, 0)
    '        ij.Txt12PjkJmlPenghasilanNetto.Text = ObjCD.DigitGrouping(No12, 0)
    '        ij.Txt13PjkJmlHasilMsPjkSblm.Text = 0
    '        ij.Txt14PjkJmlHasilNettoUPPh.Text = ObjCD.DigitGrouping(No14, 0)
    '        ij.Txt15Pjk_PTKP.Text = ObjCD.DigitGrouping(No15, 0)
    '        ij.Txt16Pjk_PKPSetahun.Text = ObjCD.DigitGrouping(No16, 0)
    '        ij.Txt17Pjk_PPH21.Text = ObjCD.DigitGrouping(No17, 0)
    '        ij.Txt18.Text = "0"
    '        If vApakahUntukGrossUp = False Then
    '            If ObjKrywn.GetIsNPWP(ij.TxtPrimaryEmployee.Text) = ClsMasterKaryawan.EnumIsNPWP.No Then
    '                ij.Txt19Pjk_PPH21Terutang.Text = ObjCD.DigitGrouping(No19 * 1.2, 0)
    '            Else
    '                ij.Txt19Pjk_PPH21Terutang.Text = ObjCD.DigitGrouping(No19, 0)
    '            End If
    '        Else
    '            ij.Txt19Pjk_PPH21Terutang.Text = ObjCD.DigitGrouping(No19, 0)
    '        End If
    '        ij.Txt20.Text = "0"
    '        Retval = CDec(ij.Txt19Pjk_PPH21Terutang.Text)

    '        Return Retval
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        ObjIncomeTax = Nothing
    '        ObjEmployee = Nothing
    '        ObjBeginningTax = Nothing
    '        ObjCD = Nothing
    '        ObjKrywn = Nothing
    '    End Try
    'End Function

    Public Function No8_HitungJumlahPenghasilanBruto(ByVal No1 As Decimal, ByVal No2 As Decimal, ByVal No3 As Decimal, ByVal No4 As Decimal, _
    ByVal No5 As Decimal, ByVal No6 As Decimal, ByVal No7 As Decimal) As Decimal
        Dim Retval As Decimal = 0
        Try
            Retval = No1 + No2 + No3 + No4 + No5 + No6 + No7
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function No11_HitungJumlahPengurang(ByVal no9 As Decimal, ByVal no10 As Decimal) As Decimal
        Dim Retval As Decimal = 0
        Try
            Retval = no9 + no10
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function No12_HItungJumlahPenghasilanNetto(ByVal no8 As Decimal, ByVal no11 As Decimal) As Decimal
        Dim Retval As Decimal = 0
        Try
            Retval = no8 - no11
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function No16_HitungJumlahPenghasilanKenaPajakDiSetahunkan(ByVal no14 As Decimal, ByVal no15 As Decimal) As Decimal
        Dim Retval As Decimal = 0
        Dim ObjCD As New ClsCekData
        Try
            Retval = ObjCD.GetRoundingValue(no14 - no15, 0, ClsCekData.EnumTipeRounding.Down, 1000, ObjCD)
            If Retval <= 0 Then
                Retval = 0
            End If
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCD = Nothing
        End Try
    End Function

    'Cara HitungYgLama(Tidakdipergunakan)
    'Public Function HitungIncomeTax(ByVal vBaru As Boolean, ByVal vTanggal As Date, ByVal vPrimary As Double, _
    'ByVal vIncome As Decimal, ByVal vDeduction As Decimal, ByVal vPrimaryEmployee As Double, _
    'ByVal vBasicSalary As Decimal, ByVal vDefPrec As Double, ByVal vPeriodePKP As EnumPeriodePKP, _
    'ByVal ijk As FrmTrans_Gaji) As Decimal
    '    Dim DTA As New DseTransaksiSalaryTableAdapters.tblmaster_tarif_pphTableAdapter
    '    Dim TBL As New DseTransaksiSalary.tblmaster_tarif_pphDataTable
    '    Dim Row As DseTransaksiSalary.tblmaster_tarif_pphRow
    '    Dim Status As ClsMasterKaryawan.GolonganPajak
    '    Dim ObjKrywn As New ClsMasterKaryawan
    '    Dim ObjSetting As New ClsMasterSettingPajak
    '    Dim Objcd As New ClsCekData
    '    Dim FaktorKali As Decimal = 1
    '    Dim PKPTotalBefore As Decimal = 0
    '    Dim PPhTotalBefore As Decimal = 0
    '    Dim vTanggalDari As Date = Objcd.GetFisrtDateOfTheYear(vTanggal)
    '    Dim Retval As Decimal = 0
    '    Dim Total As Decimal = 0
    '    Try
    '        DTA.Connection.ConnectionString = Core_Module.GetConstring
    '        If vPeriodePKP = EnumPeriodePKP.One_Month Then
    '            FaktorKali = 1
    '        ElseIf vPeriodePKP = EnumPeriodePKP.One_Week Then
    '            FaktorKali = 4
    '        ElseIf vPeriodePKP = EnumPeriodePKP.Two_Weeks Then
    '            FaktorKali = 2
    '        End If
    '        DTA.Fill(TBL)
    '        Status = ObjKrywn.GetGolonganPajak(vPrimaryEmployee)
    '        ijk.Txt_A_IncomeAdjust.Text = (vIncome - vDeduction) * FaktorKali
    '        ijk.Txt_A_PositionExpense.Text = (0.05 * vBasicSalary)
    '        ijk.Txt_A_PositionExpenseAdj.Text = ((vIncome - vDeduction) * FaktorKali) - (0.05 * vBasicSalary)
    '        ijk.Txt_A_PTKPMonthly.Text = ObjSetting.GetPTKPPerTahun(Status) / 12
    '        ijk.Txt_A_PKPValue.Text = Objcd.GetRoundingValue(CDec(ijk.Txt_A_PositionExpenseAdj.Text) - CDec(ijk.Txt_A_PTKPMonthly.Text), vDefPrec, ClsCekData.EnumTipeRounding.Down, 1000, Objcd)
    '        If CDec(ijk.Txt_A_PKPValue.Text) < 0 Then
    '            Return 0
    '        End If
    '        PKPTotalBefore = Me.GetPKPValue(vTanggal, vTanggalDari) + CDec(ijk.Txt_A_PKPValue.Text)
    '        Dim Counter As Decimal = 0
    '        For ik As Integer = 0 To TBL.Rows.Count - 1
    '            Row = TBL.Rows(ik)
    '            If ik = 0 Then
    '                Counter = Row.maximum_income - PKPTotalBefore
    '                If Counter > 0 Then
    '                    Retval += PKPTotalBefore * (Row.tarif / 100)
    '                    Exit For
    '                Else
    '                    Counter = Math.Abs(Counter)
    '                    Retval += Counter * (Row.tarif / 100)
    '                End If
    '            Else
    '                If Row.maximum_income - Counter > 0 Then
    '                    Retval += Counter * (Row.tarif / 100)
    '                    Exit For
    '                Else
    '                    Counter = Row.maximum_income - Counter
    '                    Counter = Math.Abs(Counter)
    '                    Retval += Counter * (Row.tarif / 100)
    '                End If
    '            End If
    '        Next
    '        PPhTotalBefore = Me.GetPPhValue(vTanggal, vTanggalDari)
    '        Retval = Retval - PPhTotalBefore
    '        ijk.Txt_A_PPhValue.Text = Retval
    '        Return Retval
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTA.Dispose()
    '        TBL.Dispose()
    '        ObjKrywn = Nothing
    '        Objcd = Nothing
    '        ObjSetting = Nothing
    '    End Try
    'End Function

    Public Function GetTaxCalculationDefault(ByVal vPrimaryEmployee As Double, ByVal vPrimaryMasterSalaryFormula As Double) As Integer
        Dim DTA As New DseTransaksiSalaryTableAdapters.tblmaster_salary_formula_component_detailemployeeTableAdapter
        Dim Dummy As System.Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetTaxCalculation(vPrimaryMasterSalaryFormula, vPrimaryEmployee)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlByTanggalDariSampai(ByVal vEndDate As Date, ByVal vStartDate As Date, ByVal vPrimaryEmployee As Double, _
    Optional ByVal PrimarySalary As Double = -1) As Double
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salaryTableAdapter
        Dim Dummy As System.Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If PrimarySalary = -1 Then
                Dummy = DTA.ScalarQuery_GetJmlByTanggalDariSampai(vEndDate, vStartDate, vPrimaryEmployee)
            Else
                Dummy = DTA.ScalarQuery_GetJmlByTanggalDariSampai_PosisiEdit(vEndDate, vStartDate, PrimarySalary, vPrimaryEmployee)
            End If
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub UpdatePrimaryLoanPayment(ByVal vPrimaryLoanPayment As Double, ByVal vPrimarySalary As Double)
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salary1TableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.UpdateQuery_BuatUpdatePrimaryLoanPayment(vPrimaryLoanPayment, vPrimarySalary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    'Batasan pengecekan bulan 12, selama 1 tahun 1-11 perhitungan harus sama
    Public Function ApakahSelama1TahunTaxCalculationSama(ByVal vTanggalTransSalary As Date, _
    ByVal vEnumTaxCalculation As ClsTransaksiSalary.EnumTax, ByVal vPrimaryEmployee As Double, _
    ByVal vPrimarySalary As Double) As Boolean
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salary2TableAdapter
        Dim TBL As New DseTransaksiSalary.tbltrans_salary2DataTable
        Dim Row As DseTransaksiSalary.tbltrans_salary2Row
        Dim StartDate As Date = #1/1/1950#
        Dim EndDate As Date = #1/1/1950#
        Dim ObjCD As New Core_Module.ClsCekData
        Dim Retval As String = ""
        Dim TipeTax As String = ""
        Dim ij As Integer = 1
        Try
            Return True
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            StartDate = ObjCD.GetFisrtDateOfTheYear(vTanggalTransSalary)
            EndDate = ObjCD.GetLasDateOfTheMonth(ObjCD.GetLastDateOfTheYear(vTanggalTransSalary) - TimeSpan.FromDays(32))
            DTA.FillBy_TglDariSampai_TanggalTransaksi1(TBL, EndDate, StartDate, vPrimaryEmployee, vPrimarySalary)
            If TBL.Rows.Count > 0 Then
                For Each Row In TBL.Rows
                    If Row.Tax <> vEnumTaxCalculation Then
                        If Row.Tax = ClsTransaksiSalary.EnumTax.Gross_Up Then
                            TipeTax = "gross up"
                        ElseIf Row.Tax = ClsTransaksiSalary.EnumTax.Netto Then
                            TipeTax = "netto"
                        ElseIf Row.Tax = ClsTransaksiSalary.EnumTax.Gross Then
                            TipeTax = "Gross"
                        End If
                        If String.IsNullOrEmpty(Retval) = False Then
                            Retval += vbCrLf + ij.ToString + ". " + Row.No_Transaksi.ToString + ", " + Format(Row.Tanggal_Transaksi, "dd MMM yyyy") + ", tax calulation : " + TipeTax.ToString + "."
                        Else
                            Retval = ij.ToString + ". " + Row.No_Transaksi.ToString + ", " + Format(Row.Tanggal_Transaksi, "dd MMM yyyy") + ", tax calulation : " + TipeTax.ToString + "."
                        End If
                        ij += 1
                    End If
                Next
                If String.IsNullOrEmpty(Retval) = False Then
                    'MsgBox("Please cek & update tax calculation suitable with period december " + vTanggalTransSalary.Year.ToString + "." + vbCrLf + _
                    '"The following salary transaction have to suitable :" + vbCrLf + Retval, MsgBoxStyle.Information, "Information")
                    'Return False
                    'MsgBox("Please cek tax calculation because you must choose the same tax calculation for around one year.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCD = Nothing
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function ApakahSelamaBulan1sampai11TransaksiBerjalanUdahAdaTransaksiBulan12a(ByVal vTanggalTransSalary As Date, _
    ByVal vPrimaryEmployee As Double, ByVal vUserCode As String) As Boolean
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salary2TableAdapter
        Dim TBL As New DseTransaksiSalary.tbltrans_salary2DataTable
        Dim Row As DseTransaksiSalary.tbltrans_salary2Row
        Dim StartDate As Date = #1/1/1950#
        Dim EndDate As Date = #1/1/1950#
        Dim ObjCD As New Core_Module.ClsCekData
        Dim Retval As String = ""
        Dim TipeTax As String = ""
        Dim ij As Integer = 1
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            StartDate = ObjCD.GetFirstDateofTheMonth(ObjCD.GetLastDateOfTheYear(vTanggalTransSalary))
            EndDate = ObjCD.GetLastDateOfTheYear(vTanggalTransSalary)
            DTA.FillBy_TglDariSampai_TanggalTransaksi(TBL, EndDate, StartDate, vPrimaryEmployee)
            If TBL.Rows.Count > 0 Then
                Row = TBL.Rows(0)
                If vUserCode.ToUpper = "USER" Then
                    If MsgBox("This data cannot edit or delete because this data has already used at period december " + vTanggalTransSalary.Year.ToString + ", " + vbCrLf + _
                    "document no :" + Row.No_Transaksi + ", date : " + Format(Row.Tanggal_Transaksi, "dd MMM yyyy") + ". Continue anyway?", MsgBoxStyle.YesNo, "Information") <> MsgBoxResult.Yes Then
                        Return False
                    End If
                Else
                    MsgBox("This data cannot edit or delete because this data has already used at period december " + vTanggalTransSalary.Year.ToString + ", " + vbCrLf + _
                    "document no :" + Row.No_Transaksi + ", date : " + Format(Row.Tanggal_Transaksi, "dd MMM yyyy") + ".", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    'Batasan salary dengan tax beginning balance, jika tax calulation = netto/gross maka tidak boleh menginput transaksi di bulan 12
    Public Function ApakahTunjanganPPhLebihBesarDariNol(ByVal vTanggalTransSalary As Date, ByVal vPrimaryEmployee As Double) As Boolean
        Dim DTA As New DseTransaksiSalaryTableAdapters.DataTable7TableAdapter
        Dim TBL As New DseTransaksiSalary.DataTable7DataTable
        Dim Row As DseTransaksiSalary.DataTable7Row
        Dim StartDate As Date = #1/1/1950#
        Dim EndDate As Date = #1/1/1950#
        Dim ObjCD As New Core_Module.ClsCekData
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            StartDate = ObjCD.GetFisrtDateOfTheYear(vTanggalTransSalary)
            EndDate = ObjCD.GetLasDateOfTheMonth(ObjCD.GetLastDateOfTheYear(vTanggalTransSalary) - TimeSpan.FromDays(32))
            DTA.Fill(TBL, EndDate, StartDate, vPrimaryEmployee)
            If TBL.Rows.Count > 0 Then
                Row = TBL.Rows(0)
                If Row._02 <> 0 Then
                    MsgBox("Tunjangan PPh in tax beginning balance have to zero value because tax calculation method in salary transaction is netto or gross.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    'Status Payment = UnPaid ato Paid, Payment date = .....
    Public Sub GetTBLPengecekanStatusPayment(ByVal TBL As DseTransaksiSalary.tbltrans_salary2DataTable, ByVal vPrimarySalaryMin As Double)
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salary2TableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_PrimarySalaryMin(TBL, vPrimarySalaryMin)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function GetPrimaryPayment(ByVal vPrimarySalary As Double) As Double
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salary2TableAdapter
        Dim Dummy As System.Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetPrimaryPaymentByPrimarySalary(vPrimarySalary)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPaymentDate(ByVal TBL As DseTransaksiSalary.tbltrans_salary2DataTable, _
    ByVal vPrimary As Double) As Date
        Dim Row As DseTransaksiSalary.tbltrans_salary2Row
        Dim TanggalTransaksi As Date = #1/1/1950#
        Dim Filter As String = ""
        Try
            Filter = "[Primary] = " + vPrimary.ToString + ""
            TBL.DefaultView.RowFilter = Filter
            If TBL.DefaultView.Count > 0 Then
                Row = TBL.DefaultView(0).Row
                TanggalTransaksi = Row.Tanggal_Transaksi
            Else
                TanggalTransaksi = #1/1/1950#
            End If
            TBL.DefaultView.RowFilter = ""
            'Nilai balikan
            Return TanggalTransaksi
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetPaymentStatus(ByVal TBL As DseTransaksiSalary.tbltrans_salary2DataTable, _
    ByVal vPrimary As Double) As EnumPaymentStatus
        Dim Row As DseTransaksiSalary.tbltrans_salary2Row
        Dim StatusPayment As Integer = -1
        Dim Filter As String = ""
        Try
            Filter = "[Primary] = " + vPrimary.ToString + ""
            TBL.DefaultView.RowFilter = Filter
            If TBL.DefaultView.Count > 0 Then
                Row = TBL.DefaultView(0).Row
                StatusPayment = Row.Payment_Status
            Else
                StatusPayment = -1
            End If
            TBL.DefaultView.RowFilter = ""
            'Nilai balikan
            Return StatusPayment
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetPaymentDate(ByVal vPrimary As Double) As Date
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salaryTableAdapter
        Dim Dummy As System.Nullable(Of Date)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_PaymentDate1(vPrimary)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return Date.MinValue
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPaymentStatus(ByVal vPrimary As Double) As EnumPaymentStatus
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salaryTableAdapter
        Dim Dummy As System.Nullable(Of Integer)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_PaymentStatus1(vPrimary)
            If String.IsNullOrEmpty(Dummy) = False Then
                Return Dummy
            Else
                Return EnumPaymentStatus.unpaid
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub UpdatePayment(ByVal vTanggalPay As Date, ByVal vPaymentStatus As ClsTransaksiSalary.EnumPaymentStatus, _
    ByVal vPrimaryPayment As Double, ByVal vPrimarySalary As Double)
        Dim DTA As New DseTransaksiSalaryTableAdapters.tbltrans_salary2TableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.UpdateQuery_PaymentStatusDate(vPaymentStatus, vTanggalPay, vPrimaryPayment, vPrimarySalary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub BuatPaymentData_Posisi_Edit(ByVal vDGVPayment As DataGridView, ByVal vPrimarySalary As Double, _
    ByVal FrmPaymentData As FrmTrans_Gaji_Payment, ByVal ObjAkun As AccModule.ClsMasterAkun, ByVal Def_Prec As Double, _
    ByRef AmountPayment As Decimal)
        Dim RowDGV As DataGridViewRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim DTA As New DseTransaksiSalaryTableAdapters.DataTable9TableAdapter
        Dim TBL As New DseTransaksiSalary.DataTable9DataTable
        Dim Row As DseTransaksiSalary.DataTable9Row
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_PrimarySalary(TBL, vPrimarySalary)
            vDGVPayment.Rows.Clear()
            If TBL.Rows.Count > 0 Then
                For Each Row In TBL.Rows
                    vDGVPayment.Rows.Add()
                    RowDGV = vDGVPayment.Rows(vDGVPayment.Rows.Count - 1)
                    RowDGV.Cells(FrmPaymentData.ColDetailPayment_PrimaryAkun.Index).Value = Row.primary_akun_payment
                    RowDGV.Cells(FrmPaymentData.ColDetailPayment_AccountCode.Index).Value = ObjAkun.GetAccountCode(Row.primary_akun_payment)
                    RowDGV.Cells(FrmPaymentData.ColDetailPayment_AccountDesc.Index).Value = ObjAkun.GetDescription_ByPrimaryAcc(Row.primary_akun_payment)
                    RowDGV.Cells(FrmPaymentData.ColDetailPayment_Remarks.Index).Value = Row.remarks
                    RowDGV.Cells(FrmPaymentData.ColDetailPayment_Identifier.Index).Value = Row.identifier
                    RowDGV.Cells(FrmPaymentData.ColDetailPayment_IndexRowSalary.Index).Value = "0"
                    RowDGV.Cells(FrmPaymentData.ColDetailPayment_Amount.Index).Value = ObjCD.DigitGrouping(Row.amount, Def_Prec)
                    AmountPayment += Row.amount
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjCD = Nothing
        End Try
    End Sub

    Public Function GetNoUrutBuktiPotongMax(ByVal vPeriod As Date) As Double
        Dim DTA As New DseTrans_Export_eSPTTableAdapters.tbltrans_salary_monthlyTableAdapter
        Dim ObjCD As New Core_Module.ClsCekData
        Dim RetVal As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetUrutanBuktiPotongMax(ObjCD.GetLastDateOfTheYear(vPeriod), ObjCD.GetFisrtDateOfTheYear(vPeriod))
            If String.IsNullOrEmpty(RetVal) = False Then
                If IsNumeric(RetVal) = True Then
                    Return RetVal
                Else
                    Return 0
                End If
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            DTA.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Sub HitungPajak1_Aktual(ByVal UseNPWP As Boolean, ByVal vTBLTarif As DseTransaksiSalary.tblmaster_tarif_pphDataTable, ByVal FaktorKaliDisetahunkan As Integer, _
    ByVal TBLItemTax As DseTransaksiSalary.tbltrans_salary_itemtaxDataTable, ByVal Perct_BiayaJabatan As Decimal, ByVal BiayaJabatanMaxValue As Decimal, _
    ByVal ObjCD As Core_Module.ClsCekData, ByVal NPWPCategory As Core_Module.ClsMasterKaryawan.GolonganPajak, ByVal DescNPWPCategory As String, _
    ByVal vPrimaryTransaksi As Double, ByVal vPrimaryEmployee As Double, ByVal vTglTransaksi As Date, ByVal vDocumentNo As String, _
    ByRef vTunjanganPPH_Setahun_Rutin As Decimal, ByRef vTunjanganPPH_Setahun As Decimal, ByRef Txt02Pjk_TunjPPHDisetahunkanRutin As Decimal, _
    ByRef Txt02Pjk_TunjPPHDisetahunkan As Decimal, ByRef Txt02Pjk_TunjPPHRutin As Decimal, ByRef Txt02Pjk_TunjPPHNonRutin As Decimal, ByRef Txt02PjkTunjanganPPh As Decimal, _
    ByRef Txt01Pjk_GajiRutin As Decimal, ByRef Txt01Pjk_Gaji As Decimal, ByRef Txt03Pjk_TunjLainnya As Decimal, ByRef Txt03Pjk_TunjLainnyaRutin As Decimal, _
    ByRef Txt04PjkHonorarium As Decimal, ByRef Txt04PjkHonorariumRutin As Decimal, ByRef Txt05PjkPremiAsuransi As Decimal, ByRef Txt05PjkPremiAsuransiRutin As Decimal, _
    ByRef Txt06PjkNatura As Decimal, ByRef Txt06PjkNaturaRutin As Decimal, ByRef Txt07PjkTantiem As Decimal, ByRef Txt07PjkTantiemRutin As Decimal, _
    ByRef Txt08PjkPenghasilanBruto As Decimal, ByRef Txt08PjkPenghasilanBrutoRutin As Decimal, ByRef Txt09PjkBiayaJabatan As Decimal, _
    ByRef Txt09PjkBiayaJabatanRutin As Decimal, ByRef Txt10PjkIuranPensiun As Decimal, ByRef Txt10PjkIuranPensiunRutin As Decimal, _
    ByRef Txt11PjkJmlPengurangan As Decimal, ByRef Txt11PjkJmlPenguranganRutin As Decimal, ByRef Txt12PjkJmlPenghasilanNetto As Decimal, _
    ByRef Txt13PjkJmlHasilMsPjkSblm As Decimal, ByRef Txt14PjkJmlHasilNettoUPPh As Decimal, ByRef Txt14PjkJmlHasilNettoUPPhRutin As Decimal, _
    ByRef Txt15Pjk_PTKP As Decimal, ByRef Txt16Pjk_PKPSetahun As Decimal, ByRef Txt16Pjk_PKPSetahunRutin As Decimal, ByRef Txt17Pjk_PPH21 As Decimal, _
    ByRef Txt17Pjk_PPH21Rutin As Decimal, ByRef Txt18 As Decimal, ByRef Txt19Pjk_PPH21TerutangSetahun As Decimal, ByRef Txt19Pjk_PPH21TerutangSetahunRutin As Decimal, _
    ByRef Txt19Pjk_PPH21TerutangSebulanNonRutin As Decimal, ByRef Txt19Pjk_PPH21TerutangSebulanRutin As Decimal, ByRef Txt19Pjk_PPH21Terutang As Decimal, _
    ByRef TxtIncomeTax As Decimal, ByRef TxtKurangBayarTax As Decimal, ByRef Txt20PPhTerutangYgSudahDilunasi As Decimal)
        Dim PengahasilanKenaPajak, PPHTerhutangRutin, PPHTerhutangGabungan As Decimal
        Dim FaktorKali As Decimal = 1
        Dim RemarksPerhitunganPPHGabungan As String = ""
        Dim RemarksPerhitunganPPHRutin As String = ""
        Dim RemarksTunjanganPPHNonRutin As String = ""
        Dim RemarksTunjanganPPHRutin As String = ""

        Try
            'buang item tax hasil perhtungan
            Me.BuangItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No2_TunjanganPPH, TBLItemTax)
            Me.BuangItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No8_Jumlah_Penghasilan_Bruto, TBLItemTax)
            Me.BuangItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No9_BiayaJabatan_BiayaPensiun, TBLItemTax)
            Me.BuangItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No11_JumlahPengurang, TBLItemTax)
            Me.BuangItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No12_JumlahPenghasilanUntukPerhitunganPPh, TBLItemTax)
            'Me.BuangItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No13_PenghasilanNettoMasaSebelum, TBLItemTax)
            Me.BuangItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No14_PenghasilanNettoKenaPajak, TBLItemTax)
            Me.BuangItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No15_PTKP, TBLItemTax)
            Me.BuangItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No16_PKP, TBLItemTax)
            Me.BuangItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No17_PPHSetahun, TBLItemTax)
            Me.BuangItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No18_PPH21DipotongMasaPajakSebelum, TBLItemTax)
            Me.BuangItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No19_PPHTerutang, TBLItemTax)
            Me.BuangItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No20_PPHLunas_Terpotong, TBLItemTax)

            If vTunjanganPPH_Setahun <> 0 Then
                Txt02Pjk_TunjPPHDisetahunkanRutin = vTunjanganPPH_Setahun_Rutin
                Txt02Pjk_TunjPPHDisetahunkan = vTunjanganPPH_Setahun
                Txt02Pjk_TunjPPHRutin = ObjCD.RoundDown1(vTunjanganPPH_Setahun_Rutin / FaktorKaliDisetahunkan, 0)
                Txt02Pjk_TunjPPHNonRutin = vTunjanganPPH_Setahun - vTunjanganPPH_Setahun_Rutin
                Txt02PjkTunjanganPPh = vTunjanganPPH_Setahun

                ''gabungan
                Me.MasukkanItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No2_TunjanganPPH, Txt02Pjk_TunjPPHNonRutin, "", _
                0, Txt02Pjk_TunjPPHDisetahunkan, ObjCD.DigitGrouping(Txt02Pjk_TunjPPHDisetahunkan, 0) + " - " + _
                ObjCD.DigitGrouping(Txt02Pjk_TunjPPHDisetahunkanRutin, 0), 0, "", TBLItemTax, vPrimaryTransaksi, vPrimaryEmployee, vTglTransaksi, vDocumentNo)
                'rutin
                Me.MasukkanItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No2_TunjanganPPH, Txt02Pjk_TunjPPHRutin, "", _
                1, 0, "", Txt02Pjk_TunjPPHDisetahunkanRutin, _
                ObjCD.DigitGrouping(Txt02Pjk_TunjPPHDisetahunkanRutin, 0) + " / " + FaktorKaliDisetahunkan.ToString, TBLItemTax, vPrimaryTransaksi, vPrimaryEmployee, vTglTransaksi, vDocumentNo)
            Else
                'update txt tunjangan PPH
                Txt02PjkTunjanganPPh = 0
                Txt02Pjk_TunjPPHRutin = 0
                Txt02Pjk_TunjPPHNonRutin = 0
                Txt02Pjk_TunjPPHDisetahunkanRutin = 0
                Txt02Pjk_TunjPPHDisetahunkan = 0
                'tetap masukkan data item
                RemarksTunjanganPPHNonRutin = ""
                RemarksTunjanganPPHRutin = ""
                Me.MasukkanItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No2_TunjanganPPH, Txt02PjkTunjanganPPh, "", _
                1, Txt02Pjk_TunjPPHDisetahunkan, RemarksTunjanganPPHNonRutin, Txt02Pjk_TunjPPHDisetahunkanRutin, _
                RemarksTunjanganPPHRutin, TBLItemTax, vPrimaryTransaksi, vPrimaryEmployee, vTglTransaksi, vDocumentNo)
            End If

            'hitung total penghasilan bruto
            Txt01Pjk_GajiRutin = Me.GetAmountRutin(ClsTransaksiSalary_Item.EnumTaxElement.No1_Gaji_Pensiunan_THT_JHT, TBLItemTax)
            Txt01Pjk_Gaji = Me.GetAmountGabungan(ClsTransaksiSalary_Item.EnumTaxElement.No1_Gaji_Pensiunan_THT_JHT, TBLItemTax)
            Txt03Pjk_TunjLainnya = Me.GetAmountGabungan(ClsTransaksiSalary_Item.EnumTaxElement.No3_TunjLain_lembur_dsb, TBLItemTax)
            Txt03Pjk_TunjLainnyaRutin = Me.GetAmountRutin(ClsTransaksiSalary_Item.EnumTaxElement.No3_TunjLain_lembur_dsb, TBLItemTax)
            Txt04PjkHonorarium = Me.GetAmountGabungan(ClsTransaksiSalary_Item.EnumTaxElement.No4_Honorarium_ImbalanLain_Sejenisnya, TBLItemTax)
            Txt04PjkHonorariumRutin = Me.GetAmountRutin(ClsTransaksiSalary_Item.EnumTaxElement.No4_Honorarium_ImbalanLain_Sejenisnya, TBLItemTax)
            Txt05PjkPremiAsuransi = Me.GetAmountGabungan(ClsTransaksiSalary_Item.EnumTaxElement.No5_PremiAsuransi_DibayarkanPemberiKerja, TBLItemTax)
            Txt05PjkPremiAsuransiRutin = Me.GetAmountRutin(ClsTransaksiSalary_Item.EnumTaxElement.No5_PremiAsuransi_DibayarkanPemberiKerja, TBLItemTax)
            Txt06PjkNatura = Me.GetAmountGabungan(ClsTransaksiSalary_Item.EnumTaxElement.No6_Natura_KenikmatanLain_KenaPPH, TBLItemTax)
            Txt06PjkNaturaRutin = Me.GetAmountRutin(ClsTransaksiSalary_Item.EnumTaxElement.No6_Natura_KenikmatanLain_KenaPPH, TBLItemTax)
            Txt07PjkTantiem = Me.GetAmountGabungan(ClsTransaksiSalary_Item.EnumTaxElement.No7_Tantiem_Bonus_Gratifikasi_JasaProduksi, TBLItemTax)
            Txt07PjkTantiemRutin = Me.GetAmountRutin(ClsTransaksiSalary_Item.EnumTaxElement.No7_Tantiem_Bonus_Gratifikasi_JasaProduksi, TBLItemTax)
            Txt08PjkPenghasilanBruto = Txt01Pjk_Gaji + Txt02Pjk_TunjPPHDisetahunkan + _
            Txt03Pjk_TunjLainnya + Txt04PjkHonorarium + Txt05PjkPremiAsuransi + Txt06PjkNatura + _
            Txt07PjkTantiem
            Txt08PjkPenghasilanBrutoRutin = Txt01Pjk_GajiRutin + Txt02Pjk_TunjPPHDisetahunkanRutin + _
            Txt03Pjk_TunjLainnyaRutin + Txt04PjkHonorariumRutin + Txt05PjkPremiAsuransiRutin + Txt06PjkNaturaRutin + _
            Txt07PjkTantiemRutin
            'masukkan item tax
            Me.MasukkanItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No8_Jumlah_Penghasilan_Bruto, 0, "", _
            1, Txt08PjkPenghasilanBruto, "sum of point 1 to 7", Txt08PjkPenghasilanBrutoRutin, "", TBLItemTax, _
            vPrimaryTransaksi, vPrimaryEmployee, vTglTransaksi, vDocumentNo)

            'hitung biaya jabatan
            Txt09PjkBiayaJabatan = ObjCD.RoundDown1(Perct_BiayaJabatan * Txt08PjkPenghasilanBruto, 0)
            If Txt09PjkBiayaJabatan > ((BiayaJabatanMaxValue * FaktorKaliDisetahunkan) / 12) Then
                Txt09PjkBiayaJabatan = ((BiayaJabatanMaxValue * FaktorKaliDisetahunkan) / 12)
            End If
            Txt09PjkBiayaJabatanRutin = ObjCD.RoundDown1(Perct_BiayaJabatan * Txt08PjkPenghasilanBrutoRutin, 0)
            If Txt09PjkBiayaJabatanRutin > ((BiayaJabatanMaxValue * FaktorKaliDisetahunkan) / 12) Then
                Txt09PjkBiayaJabatanRutin = ((BiayaJabatanMaxValue * FaktorKaliDisetahunkan) / 12)
            End If
            'masukkan item tax
            Me.MasukkanItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No9_BiayaJabatan_BiayaPensiun, 0, "", 1, _
            Txt09PjkBiayaJabatan, "5 % of point 8, max value " + ObjCD.DigitGrouping(((BiayaJabatanMaxValue * FaktorKaliDisetahunkan) / 12), 0), _
            Txt09PjkBiayaJabatanRutin, "", TBLItemTax, vPrimaryTransaksi, vPrimaryEmployee, vTglTransaksi, vDocumentNo)

            'hitung iuran THT/JHT
            Txt10PjkIuranPensiun = Me.GetAmountGabungan(ClsTransaksiSalary_Item.EnumTaxElement.No10_IuranPensiun_THT_JHT, TBLItemTax)
            Txt10PjkIuranPensiunRutin = Me.GetAmountRutin(ClsTransaksiSalary_Item.EnumTaxElement.No10_IuranPensiun_THT_JHT, TBLItemTax)

            'hitung total pengurang
            Txt11PjkJmlPengurangan = Txt09PjkBiayaJabatan + Txt10PjkIuranPensiun
            Txt11PjkJmlPenguranganRutin = Txt09PjkBiayaJabatanRutin + Txt10PjkIuranPensiunRutin

            'masukkan item tax
            Me.MasukkanItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No11_JumlahPengurang, 0, "", 1, _
            Txt11PjkJmlPengurangan, "sum of point 9 & 10", _
            Txt11PjkJmlPenguranganRutin, "", TBLItemTax, vPrimaryTransaksi, vPrimaryEmployee, vTglTransaksi, vDocumentNo)

            'hitung jumlah penghasil netto untuk perhitungan pph 21
            Txt12PjkJmlPenghasilanNetto = "0" '-> ini yang belum disetahunkan, anggap saja nol
            Txt13PjkJmlHasilMsPjkSblm = 0 '-> ini sementara tidak diperhitungkan oleh smess
            Txt14PjkJmlHasilNettoUPPh = Txt08PjkPenghasilanBruto - Txt11PjkJmlPengurangan
            If Txt14PjkJmlHasilNettoUPPh < 0 Then
                Txt14PjkJmlHasilNettoUPPh = 0
            End If
            Txt14PjkJmlHasilNettoUPPhRutin = Txt08PjkPenghasilanBrutoRutin - Txt11PjkJmlPenguranganRutin
            If Txt14PjkJmlHasilNettoUPPhRutin < 0 Then
                Txt14PjkJmlHasilNettoUPPhRutin = 0
            End If
            Me.MasukkanItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No14_PenghasilanNettoKenaPajak, 0, "", 1, _
            Txt14PjkJmlHasilNettoUPPh, "difference of point 8 & 11", _
            Txt14PjkJmlHasilNettoUPPhRutin, "", TBLItemTax, vPrimaryTransaksi, vPrimaryEmployee, vTglTransaksi, vDocumentNo)

            'ambil data PTKP
            NPWPCategory = NPWPCategory
            'masukkan item tax
            Me.MasukkanItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No15_PTKP, 0, "", 1, _
            Txt15Pjk_PTKP, "category : " + DescNPWPCategory, _
            Txt15Pjk_PTKP, "", TBLItemTax, vPrimaryTransaksi, vPrimaryEmployee, vTglTransaksi, vDocumentNo)

            'hitung penghasilan kena pajak
            Txt16Pjk_PKPSetahun = ObjCD.RoundDown2(Txt14PjkJmlHasilNettoUPPh - Txt15Pjk_PTKP, 1000)
            If Txt16Pjk_PKPSetahun < 0 Then
                Txt16Pjk_PKPSetahun = 0
            End If
            Txt16Pjk_PKPSetahunRutin = ObjCD.RoundDown2(Txt14PjkJmlHasilNettoUPPhRutin - Txt15Pjk_PTKP, 1000)
            If Txt16Pjk_PKPSetahunRutin < 0 Then
                Txt16Pjk_PKPSetahunRutin = 0
            End If
            'masukkan item
            Me.MasukkanItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No16_PKP, 0, "", 1, _
            Txt16Pjk_PKPSetahun, "difference of point 12 and 13, rounddown 1,000", _
            Txt16Pjk_PKPSetahunRutin, "", TBLItemTax, vPrimaryTransaksi, vPrimaryEmployee, vTglTransaksi, vDocumentNo)

            'hitung pph terhutang
            Txt17Pjk_PPH21 = 0
            Txt17Pjk_PPH21Rutin = 0
            Txt18 = 0
            If UseNPWP = True Then
                FaktorKali = 1
            Else
                FaktorKali = 1.2
            End If
            'gabungan setahun
            PengahasilanKenaPajak = Txt16Pjk_PKPSetahun

            PPHTerhutangGabungan = 0
            RemarksPerhitunganPPHGabungan = ""
            Me.HitungPPHTerhutang(vTBLTarif, PPHTerhutangGabungan, PengahasilanKenaPajak, FaktorKali, RemarksPerhitunganPPHGabungan, ObjCD)
            Txt19Pjk_PPH21TerutangSetahun = PPHTerhutangGabungan
            'rutin setahun
            PengahasilanKenaPajak = Txt16Pjk_PKPSetahunRutin
            PPHTerhutangRutin = 0
            RemarksPerhitunganPPHRutin = ""
            Me.HitungPPHTerhutang(vTBLTarif, PPHTerhutangRutin, PengahasilanKenaPajak, FaktorKali, RemarksPerhitunganPPHRutin, ObjCD)
            Txt19Pjk_PPH21TerutangSetahunRutin = PPHTerhutangRutin
            'masukan item tax
            Me.MasukkanItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No19_PPHTerutang, 0, "", 1, _
            PPHTerhutangGabungan, RemarksPerhitunganPPHGabungan, _
            0, "", TBLItemTax, vPrimaryTransaksi, vPrimaryEmployee, vTglTransaksi, vDocumentNo)
            If PPHTerhutangRutin <> 0 Then
                Me.MasukkanItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No19_PPHTerutang, 0, "", 1, _
                0, "", _
                PPHTerhutangRutin, RemarksPerhitunganPPHRutin, TBLItemTax, vPrimaryTransaksi, vPrimaryEmployee, vTglTransaksi, vDocumentNo)
            End If
            'sebulan non rutin
            Txt19Pjk_PPH21TerutangSebulanNonRutin = PPHTerhutangGabungan - PPHTerhutangRutin
            If Txt19Pjk_PPH21TerutangSebulanNonRutin <> 0 Then
                '#sinkron2018_pphcls
                Me.MasukkanItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No19_PPHTerutang, Txt19Pjk_PPH21TerutangSebulanNonRutin, "", 1, _
                0, ObjCD.DigitGrouping(PPHTerhutangGabungan, 0) + " - " + ObjCD.DigitGrouping(PPHTerhutangRutin, 0), _
                0, "", TBLItemTax, vPrimaryTransaksi, vPrimaryEmployee, vTglTransaksi, vDocumentNo)
            Else
                Me.MasukkanItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No19_PPHTerutang, 0, "", 1, _
                0, ObjCD.DigitGrouping(PPHTerhutangGabungan, 0) + " - " + ObjCD.DigitGrouping(PPHTerhutangRutin, 0), 0, "", TBLItemTax, vPrimaryTransaksi, vPrimaryEmployee, vTglTransaksi, vDocumentNo)
            End If
            'sebulan rutin
            Txt19Pjk_PPH21TerutangSebulanRutin = ObjCD.RoundDown1(PPHTerhutangRutin / FaktorKaliDisetahunkan, 0)
            If Txt19Pjk_PPH21TerutangSebulanRutin <> 0 Then
                Me.MasukkanItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No19_PPHTerutang, Txt19Pjk_PPH21TerutangSebulanRutin, "", 1, _
                0, "", 0, ObjCD.DigitGrouping(PPHTerhutangRutin, 0) + " / " + FaktorKaliDisetahunkan.ToString, _
                TBLItemTax, vPrimaryTransaksi, vPrimaryEmployee, vTglTransaksi, vDocumentNo)
            Else
                Me.MasukkanItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No19_PPHTerutang, 0, "", 1, _
                0, "", 0, "", TBLItemTax, vPrimaryTransaksi, vPrimaryEmployee, vTglTransaksi, vDocumentNo)
            End If
            'total pph terhutang
            Txt19Pjk_PPH21Terutang = ObjCD.DigitGrouping(Txt19Pjk_PPH21TerutangSebulanRutin + _
            Txt19Pjk_PPH21TerutangSebulanNonRutin, 0)
            If Txt19Pjk_PPH21Terutang <> 0 Then
                Me.MasukkanItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No19_PPHTerutang, Txt19Pjk_PPH21Terutang, "", 1, _
                0, ObjCD.DigitGrouping(Txt19Pjk_PPH21TerutangSebulanRutin, 0) + " + " + _
                ObjCD.DigitGrouping(Txt19Pjk_PPH21TerutangSebulanNonRutin, 0), _
                0, "", TBLItemTax, vPrimaryTransaksi, vPrimaryEmployee, vTglTransaksi, vDocumentNo)
            Else
                Me.MasukkanItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No19_PPHTerutang, 0, "", 1, _
                0, "", 0, "", TBLItemTax, vPrimaryTransaksi, vPrimaryEmployee, vTglTransaksi, vDocumentNo)
            End If

            TxtIncomeTax = ObjCD.DigitGrouping(Txt19Pjk_PPH21Terutang, 0)
            TxtKurangBayarTax = ObjCD.DigitGrouping(Txt19Pjk_PPH21Terutang, 0)
            Txt20PPhTerutangYgSudahDilunasi = TxtIncomeTax
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub HitungPajak1_Aktual_EndOfYear(ByVal vTunjanganPPH_Sebulan_Rutin As Decimal, ByVal vTunjanganPPH_Sebulan_NonRutin As Decimal, _
    ByVal UseNPWP As Boolean, ByVal vTBLTarif As DseTransaksiSalary.tblmaster_tarif_pphDataTable, _
    ByVal vTunjanganPPH_Setahun_Rutin As Decimal, ByVal vTunjanganPPH_Setahun As Decimal, ByVal vStartWorkingPeriod As Date, ByVal vResignDate As Date, _
    ByVal FaktorKaliDisetahunkan As Integer, ByVal TBLItemTax As DseTransaksiSalary.tbltrans_salary_itemtaxDataTable, ByVal Perct_BiayaJabatan As Decimal, ByVal BiayaJabatanMaxValue As Decimal, _
    ByVal ObjCD As Core_Module.ClsCekData, ByVal NPWPCategory As Core_Module.ClsMasterKaryawan.GolonganPajak, ByVal DescNPWPCategory As String, _
    ByVal vPrimaryTransaksi As Double, ByVal vPrimaryEmployee As Double, ByVal vTglTransaksi As Date, ByVal vDocumentNo As String, _
    ByVal vSalary_EndPeriod As Date, ByRef vPPhTerutangPeriodeSebelumnya As Decimal, _
    ByRef Txt02Pjk_TunjPPHDisetahunkanRutin As Decimal, _
    ByRef Txt02Pjk_TunjPPHDisetahunkan As Decimal, ByRef Txt02Pjk_TunjPPHRutin As Decimal, ByRef Txt02Pjk_TunjPPHNonRutin As Decimal, ByRef Txt02PjkTunjanganPPh As Decimal, _
    ByRef Txt01Pjk_GajiRutin As Decimal, ByRef Txt01Pjk_Gaji As Decimal, ByRef Txt03Pjk_TunjLainnya As Decimal, ByRef Txt03Pjk_TunjLainnyaRutin As Decimal, _
    ByRef Txt04PjkHonorarium As Decimal, ByRef Txt04PjkHonorariumRutin As Decimal, ByRef Txt05PjkPremiAsuransi As Decimal, ByRef Txt05PjkPremiAsuransiRutin As Decimal, _
    ByRef Txt06PjkNatura As Decimal, ByRef Txt06PjkNaturaRutin As Decimal, ByRef Txt07PjkTantiem As Decimal, ByRef Txt07PjkTantiemRutin As Decimal, _
    ByRef Txt08PjkPenghasilanBruto As Decimal, ByRef Txt08PjkPenghasilanBrutoRutin As Decimal, ByRef Txt09PjkBiayaJabatan As Decimal, _
    ByRef Txt09PjkBiayaJabatanRutin As Decimal, ByRef Txt10PjkIuranPensiun As Decimal, ByRef Txt10PjkIuranPensiunRutin As Decimal, _
    ByRef Txt11PjkJmlPengurangan As Decimal, ByRef Txt11PjkJmlPenguranganRutin As Decimal, ByRef Txt12PjkJmlPenghasilanNetto As Decimal, _
    ByRef Txt13PenghasilanMsPjkSblm As Decimal, ByRef Txt14PjkJmlHasilNettoUPPh As Decimal, ByRef Txt14PjkJmlHasilNettoUPPhRutin As Decimal, _
    ByRef Txt15Pjk_PTKP As Decimal, ByRef Txt16Pjk_PKPSetahun As Decimal, ByRef Txt16Pjk_PKPSetahunRutin As Decimal, ByRef Txt17Pjk_PPH21 As Decimal, _
    ByRef Txt17Pjk_PPH21Rutin As Decimal, ByRef Txt18 As Decimal, ByRef Txt19Pjk_PPH21TerutangSetahun As Decimal, ByRef Txt19Pjk_PPH21TerutangSetahunRutin As Decimal, _
    ByRef Txt19Pjk_PPH21TerutangSebulanNonRutin As Decimal, ByRef Txt19Pjk_PPH21TerutangSebulanRutin As Decimal, ByRef Txt19Pjk_PPH21Terutang As Decimal, _
    ByRef TxtIncomeTax As Decimal, ByRef TxtKurangBayarTax As Decimal, ByRef Txt20PPhTerutangYgSudahDilunasi As Decimal, ByVal vLosesTaxObligation As Integer, _
    ByVal vEndPeriod As Integer, ByVal vPeriodCount As Integer, ByVal vIsResign As Boolean, ByVal vTaxPeriod As Date, ByVal vStartPeriod As Integer, _
    ByVal Txt13PenghasilanMsPjkSblm1 As Decimal, ByVal Txt18_PPh21_yang_sdh_DisetorMasaSblm1 As Decimal, _
    ByVal vTBLPrevIncome As DseTransaksiSalary.tbltrans_salary_item_prev_income1DataTable)
        Dim rowPrevIncome As DseTransaksiSalary.tbltrans_salary_item_prev_income1Row
        Dim PengahasilanKenaPajak, PPHTerhutangRutin, PPHTerhutangGabungan As Decimal
        Dim FaktorKali As Decimal = 1
        Dim RemarksPerhitunganPPHGabungan As String = ""
        Dim RemarksPerhitunganPPHRutin As String = ""
        Dim RemarksTunjanganPPHNonRutin As String = ""
        Dim RemarksTunjanganPPHRutin As String = ""
        Dim Dummy1 As String = ""
        Dim DescPPH As FrmTrans_Gaji_IncomeTax2.EnumDescPPHTerhutang = FrmTrans_Gaji_IncomeTax2.EnumDescPPHTerhutang.Pajak_Terhutang_Bulanan

        Try
            'buang item tax hasil perhtungan
            Me.BuangItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No2_TunjanganPPH, TBLItemTax)
            Me.BuangItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No8_Jumlah_Penghasilan_Bruto, TBLItemTax)
            Me.BuangItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No9_BiayaJabatan_BiayaPensiun, TBLItemTax)
            Me.BuangItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No11_JumlahPengurang, TBLItemTax)
            Me.BuangItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No12_JumlahPenghasilanUntukPerhitunganPPh, TBLItemTax)
            Me.BuangItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No13_PenghasilanNettoMasaSebelum, TBLItemTax)
            Me.BuangItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No14_PenghasilanNettoKenaPajak, TBLItemTax)
            Me.BuangItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No15_PTKP, TBLItemTax)
            Me.BuangItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No16_PKP, TBLItemTax)
            Me.BuangItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No17_PPHSetahun, TBLItemTax)
            Me.BuangItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No18_PPH21DipotongMasaPajakSebelum, TBLItemTax)
            Me.BuangItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No19_PPHTerutang, TBLItemTax)
            Me.BuangItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No20_PPHLunas_Terpotong, TBLItemTax)

            If vTunjanganPPH_Setahun <> 0 Then
                'update txt tunjangan PPH
                Txt02PjkTunjanganPPh = vTunjanganPPH_Sebulan_NonRutin + vTunjanganPPH_Sebulan_Rutin
                Txt02Pjk_TunjPPHRutin = vTunjanganPPH_Sebulan_Rutin
                Txt02Pjk_TunjPPHNonRutin = vTunjanganPPH_Sebulan_NonRutin
                Txt02Pjk_TunjPPHDisetahunkanRutin = vTunjanganPPH_Setahun_Rutin
                Txt02Pjk_TunjPPHDisetahunkan = vTunjanganPPH_Setahun

                'update data item
                'rutin
                Me.MasukkanItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No2_TunjanganPPH, CDec(Txt02Pjk_TunjPPHRutin), "", _
                1, 0, ObjCD.DigitGrouping(CDec(Txt02Pjk_TunjPPHDisetahunkan), 0) + " - " + _
                ObjCD.DigitGrouping(CDec(Txt02Pjk_TunjPPHDisetahunkanRutin), 0), CDec(Txt02Pjk_TunjPPHDisetahunkanRutin), _
                "", _
                TBLItemTax, vPrimaryTransaksi, vPrimaryEmployee, vTglTransaksi, vDocumentNo)
                'gabungan
                Me.MasukkanItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No2_TunjanganPPH, CDec(Txt02Pjk_TunjPPHNonRutin), "", _
                0, CDec(Txt02Pjk_TunjPPHDisetahunkan), "", 0, "", _
                TBLItemTax, vPrimaryTransaksi, vPrimaryEmployee, vTglTransaksi, vDocumentNo)
            Else
                'update txt tunjangan PPH
                Txt02PjkTunjanganPPh = 0
                Txt02Pjk_TunjPPHRutin = 0
                Txt02Pjk_TunjPPHNonRutin = 0
                Txt02Pjk_TunjPPHDisetahunkanRutin = 0
                Txt02Pjk_TunjPPHDisetahunkan = 0
                'tetap masukkan data item
                RemarksTunjanganPPHRutin = ""
                RemarksTunjanganPPHRutin = ""

                Me.MasukkanItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No2_TunjanganPPH, CDec(Txt02PjkTunjanganPPh), "", _
                1, CDec(Txt02Pjk_TunjPPHDisetahunkan), RemarksTunjanganPPHNonRutin, CDec(Txt02Pjk_TunjPPHDisetahunkanRutin), _
                RemarksTunjanganPPHRutin, _
                TBLItemTax, vPrimaryTransaksi, vPrimaryEmployee, vTglTransaksi, vDocumentNo)
            End If

            'hitung total penghasilan bruto
            Txt01Pjk_GajiRutin = Me.GetAmountRutin(ClsTransaksiSalary_Item.EnumTaxElement.No1_Gaji_Pensiunan_THT_JHT, TBLItemTax)
            Txt01Pjk_Gaji = Me.GetAmountGabungan(ClsTransaksiSalary_Item.EnumTaxElement.No1_Gaji_Pensiunan_THT_JHT, TBLItemTax)
            Txt03Pjk_TunjLainnya = Me.GetAmountGabungan(ClsTransaksiSalary_Item.EnumTaxElement.No3_TunjLain_lembur_dsb, TBLItemTax)
            Txt03Pjk_TunjLainnyaRutin = Me.GetAmountRutin(ClsTransaksiSalary_Item.EnumTaxElement.No3_TunjLain_lembur_dsb, TBLItemTax)
            Txt04PjkHonorarium = Me.GetAmountGabungan(ClsTransaksiSalary_Item.EnumTaxElement.No4_Honorarium_ImbalanLain_Sejenisnya, TBLItemTax)
            Txt04PjkHonorariumRutin = Me.GetAmountRutin(ClsTransaksiSalary_Item.EnumTaxElement.No4_Honorarium_ImbalanLain_Sejenisnya, TBLItemTax)
            Txt05PjkPremiAsuransi = Me.GetAmountGabungan(ClsTransaksiSalary_Item.EnumTaxElement.No5_PremiAsuransi_DibayarkanPemberiKerja, TBLItemTax)
            Txt05PjkPremiAsuransiRutin = Me.GetAmountRutin(ClsTransaksiSalary_Item.EnumTaxElement.No5_PremiAsuransi_DibayarkanPemberiKerja, TBLItemTax)
            Txt06PjkNatura = Me.GetAmountGabungan(ClsTransaksiSalary_Item.EnumTaxElement.No6_Natura_KenikmatanLain_KenaPPH, TBLItemTax)
            Txt06PjkNaturaRutin = Me.GetAmountRutin(ClsTransaksiSalary_Item.EnumTaxElement.No6_Natura_KenikmatanLain_KenaPPH, TBLItemTax)
            Txt07PjkTantiem = Me.GetAmountGabungan(ClsTransaksiSalary_Item.EnumTaxElement.No7_Tantiem_Bonus_Gratifikasi_JasaProduksi, TBLItemTax)
            Txt07PjkTantiemRutin = Me.GetAmountRutin(ClsTransaksiSalary_Item.EnumTaxElement.No7_Tantiem_Bonus_Gratifikasi_JasaProduksi, TBLItemTax)
            Debug.Print("F: Tunj. PPH Gabungan (Disetahunkan): " + ObjCD.DigitGrouping(Txt02Pjk_TunjPPHDisetahunkan, 0))
            Txt08PjkPenghasilanBruto = CDec(Txt01Pjk_Gaji) + CDec(Txt02Pjk_TunjPPHDisetahunkan) + _
            CDec(Txt03Pjk_TunjLainnya) + CDec(Txt04PjkHonorarium) + CDec(Txt05PjkPremiAsuransi) + CDec(Txt06PjkNatura) + _
            CDec(Txt07PjkTantiem)
            Debug.Print("F: Tunj. PPH Rutin (Disetahunkan): " + ObjCD.DigitGrouping(Txt02Pjk_TunjPPHDisetahunkanRutin, 0))
            Txt08PjkPenghasilanBrutoRutin = CDec(Txt01Pjk_GajiRutin) + CDec(Txt02Pjk_TunjPPHDisetahunkanRutin) + _
            CDec(Txt03Pjk_TunjLainnyaRutin) + CDec(Txt04PjkHonorariumRutin) + CDec(Txt05PjkPremiAsuransiRutin) + CDec(Txt06PjkNaturaRutin) + _
            CDec(Txt07PjkTantiemRutin)
            'masukkan item tax
            Me.MasukkanItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No8_Jumlah_Penghasilan_Bruto, 0, "", _
            1, CDec(Txt08PjkPenghasilanBruto), "sum of point 1 to 7", CDec(Txt08PjkPenghasilanBrutoRutin), "", _
                TBLItemTax, vPrimaryTransaksi, vPrimaryEmployee, vTglTransaksi, vDocumentNo)

            'hitung biaya jabatan
            Txt09PjkBiayaJabatan = ObjCD.RoundDown1(Perct_BiayaJabatan * CDec(Txt08PjkPenghasilanBruto), 0)
            If CDec(Txt09PjkBiayaJabatan) > ((BiayaJabatanMaxValue * FaktorKaliDisetahunkan) / 12) Then
                Txt09PjkBiayaJabatan = ((BiayaJabatanMaxValue * FaktorKaliDisetahunkan) / 12)
            End If
            Txt09PjkBiayaJabatanRutin = ObjCD.RoundDown1(Perct_BiayaJabatan * CDec(Txt08PjkPenghasilanBrutoRutin), 0)
            If CDec(Txt09PjkBiayaJabatanRutin) > ((BiayaJabatanMaxValue * FaktorKaliDisetahunkan) / 12) Then
                Txt09PjkBiayaJabatanRutin = ((BiayaJabatanMaxValue * FaktorKaliDisetahunkan) / 12)
            End If
            'masukkan item tax
            Me.MasukkanItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No9_BiayaJabatan_BiayaPensiun, 0, "", 1, _
            CDec(Txt09PjkBiayaJabatan), "5 % of point 8, max value " + ObjCD.DigitGrouping(((BiayaJabatanMaxValue * FaktorKaliDisetahunkan) / 12), 0), _
            CDec(Txt09PjkBiayaJabatanRutin), "", _
                TBLItemTax, vPrimaryTransaksi, vPrimaryEmployee, vTglTransaksi, vDocumentNo)

            'hitung iuran THT/JHT
            Txt10PjkIuranPensiun = Me.GetAmountGabungan(ClsTransaksiSalary_Item.EnumTaxElement.No10_IuranPensiun_THT_JHT, TBLItemTax)
            Txt10PjkIuranPensiunRutin = Me.GetAmountRutin(ClsTransaksiSalary_Item.EnumTaxElement.No10_IuranPensiun_THT_JHT, TBLItemTax)

            'hitung total pengurang
            Txt11PjkJmlPengurangan = CDec(Txt09PjkBiayaJabatan) + CDec(Txt10PjkIuranPensiun)
            Txt11PjkJmlPenguranganRutin = CDec(Txt09PjkBiayaJabatanRutin) + CDec(Txt10PjkIuranPensiunRutin)
            'masukkan item tax
            Me.MasukkanItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No11_JumlahPengurang, 0, "", 1, _
            CDec(Txt11PjkJmlPengurangan), "sum of point 9 & 10", _
            CDec(Txt11PjkJmlPenguranganRutin), "", _
                TBLItemTax, vPrimaryTransaksi, vPrimaryEmployee, vTglTransaksi, vDocumentNo)

            'hitung jumlah penghasil netto untuk perhitungan pph 21
            Txt12PjkJmlPenghasilanNetto = "0" '-> ini yang belum disetahunkan, anggap saja nol
            Txt13PenghasilanMsPjkSblm = Txt13PenghasilanMsPjkSblm1
            If Txt13PenghasilanMsPjkSblm <> 0 Then
                rowPrevIncome = vTBLPrevIncome.Rows(0)

                ' Me.MasukkanItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No13_PenghasilanNettoMasaSebelum, 0, "", 1, _
                'CDec(Txt13PenghasilanMsPjkSblm), "", _
                'CDec(Txt13PenghasilanMsPjkSblm), "", _
                'TBLItemTax, vPrimaryTransaksi, vPrimaryEmployee, vTglTransaksi, vDocumentNo)

                Me.MasukkanItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No13_PenghasilanNettoMasaSebelum, 0, "", 1, _
                CDec(Txt13PenghasilanMsPjkSblm), "A1 Form No:" + rowPrevIncome.a1_form_no + _
                " (" + Format(rowPrevIncome.start_period, "MMM") + " - " + Format(rowPrevIncome.start_period, "MMM yyyy") + " - " + rowPrevIncome.employer + ")", _
                CDec(Txt13PenghasilanMsPjkSblm), "", _
                TBLItemTax, vPrimaryTransaksi, vPrimaryEmployee, vTglTransaksi, vDocumentNo)
            End If
            If vEndPeriod <> 12 And vIsResign = True And vLosesTaxObligation = 1 Then
                Txt14PjkJmlHasilNettoUPPh = Math.Round(((CDec(Txt08PjkPenghasilanBruto) - CDec(Txt11PjkJmlPengurangan) + Txt13PenghasilanMsPjkSblm) * 12) / vPeriodCount, 0, MidpointRounding.AwayFromZero)
                If CDec(Txt14PjkJmlHasilNettoUPPh) < 0 Then
                    Txt14PjkJmlHasilNettoUPPh = 0
                End If
                Txt14PjkJmlHasilNettoUPPhRutin = Math.Round(((CDec(Txt08PjkPenghasilanBrutoRutin) - CDec(Txt11PjkJmlPenguranganRutin) + Txt13PenghasilanMsPjkSblm) * 12) / vPeriodCount, 0, MidpointRounding.AwayFromZero)
                If CDec(Txt14PjkJmlHasilNettoUPPhRutin) < 0 Then
                    Txt14PjkJmlHasilNettoUPPhRutin = 0
                End If
                '#sinkron2018_pphcls
                If CDec(Txt13PenghasilanMsPjkSblm) = 0 Then
                    Me.MasukkanItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No14_PenghasilanNettoKenaPajak, 0, "", 1, _
                    CDec(Txt14PjkJmlHasilNettoUPPh), "((difference of point 8 & 11) * 12) /" + ObjCD.DigitGrouping(vPeriodCount, 0), _
                    CDec(Txt14PjkJmlHasilNettoUPPhRutin), "", _
                        TBLItemTax, vPrimaryTransaksi, vPrimaryEmployee, vTglTransaksi, vDocumentNo)
                Else
                    Me.MasukkanItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No14_PenghasilanNettoKenaPajak, 0, "", 1, _
                    CDec(Txt14PjkJmlHasilNettoUPPh), "((point 8 - 11 + 12) * 12) /" + ObjCD.DigitGrouping(vPeriodCount, 0), _
                    CDec(Txt14PjkJmlHasilNettoUPPhRutin), "", _
                    TBLItemTax, vPrimaryTransaksi, vPrimaryEmployee, vTglTransaksi, vDocumentNo)
                End If
            Else
                Txt14PjkJmlHasilNettoUPPh = CDec(Txt08PjkPenghasilanBruto) - CDec(Txt11PjkJmlPengurangan) + Txt13PenghasilanMsPjkSblm
                If CDec(Txt14PjkJmlHasilNettoUPPh) < 0 Then
                    Txt14PjkJmlHasilNettoUPPh = 0
                End If
                Txt14PjkJmlHasilNettoUPPhRutin = CDec(Txt08PjkPenghasilanBrutoRutin) - CDec(Txt11PjkJmlPenguranganRutin) + Txt13PenghasilanMsPjkSblm
                If CDec(Txt14PjkJmlHasilNettoUPPhRutin) < 0 Then
                    Txt14PjkJmlHasilNettoUPPhRutin = 0
                End If
                '#sinkron2018_pphcls
                If CDec(Txt13PenghasilanMsPjkSblm) = 0 Then
                    Me.MasukkanItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No14_PenghasilanNettoKenaPajak, 0, "", 1, _
                    CDec(Txt14PjkJmlHasilNettoUPPh), "difference of point 8 & 11", _
                    CDec(Txt14PjkJmlHasilNettoUPPhRutin), "", _
                        TBLItemTax, vPrimaryTransaksi, vPrimaryEmployee, vTglTransaksi, vDocumentNo)
                Else
                    Me.MasukkanItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No14_PenghasilanNettoKenaPajak, 0, "", 1, _
                    CDec(Txt14PjkJmlHasilNettoUPPh), "point 8 - 11 + 12", _
                    CDec(Txt14PjkJmlHasilNettoUPPhRutin), "", _
                    TBLItemTax, vPrimaryTransaksi, vPrimaryEmployee, vTglTransaksi, vDocumentNo)
                End If
            End If

            'ambil data PTKP
            'NPWPCategory = Me.ObjKrywn.GetNPWPCategory(Me.TxtPrimaryEmployee.Text)
            'Txt15Pjk_PTKP = Me.ObjPTKP.GetPTKPPerTahun(NPWPCategory)
            'masukkan item tax
            Me.MasukkanItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No15_PTKP, 0, "", 1, _
            CDec(Txt15Pjk_PTKP), "category : " + DescNPWPCategory, _
            CDec(Txt15Pjk_PTKP), "", _
                TBLItemTax, vPrimaryTransaksi, vPrimaryEmployee, vTglTransaksi, vDocumentNo)

            'hitung penghasilan kena pajak
            Txt16Pjk_PKPSetahun = ObjCD.RoundDown2(CDec(Txt14PjkJmlHasilNettoUPPh) - CDec(Txt15Pjk_PTKP), 1000)
            If CDec(Txt16Pjk_PKPSetahun) < 0 Then
                Txt16Pjk_PKPSetahun = 0
            End If
            Txt16Pjk_PKPSetahunRutin = ObjCD.RoundDown2(CDec(Txt14PjkJmlHasilNettoUPPhRutin) - CDec(Txt15Pjk_PTKP), 1000)
            If CDec(Txt16Pjk_PKPSetahunRutin) < 0 Then
                Txt16Pjk_PKPSetahunRutin = 0
            End If
            'masukkan item
            If CDec(Txt13PenghasilanMsPjkSblm) = 0 Then
                Me.MasukkanItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No16_PKP, 0, "", 1, _
                            CDec(Txt16Pjk_PKPSetahun), "difference of point 12 and 13, rounddown 1,000", _
                            CDec(Txt16Pjk_PKPSetahunRutin), "", _
                                TBLItemTax, vPrimaryTransaksi, vPrimaryEmployee, vTglTransaksi, vDocumentNo)
            Else
                Me.MasukkanItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No16_PKP, 0, "", 1, _
            CDec(Txt16Pjk_PKPSetahun), "difference of point 13 and 14, rounddown 1,000", _
            CDec(Txt16Pjk_PKPSetahunRutin), "", _
                TBLItemTax, vPrimaryTransaksi, vPrimaryEmployee, vTglTransaksi, vDocumentNo)
            End If

            'hitung pph terhutang
            Txt17Pjk_PPH21 = 0
            Txt17Pjk_PPH21Rutin = 0
            Txt18 = 0
            If UseNPWP = True Then
                FaktorKali = 1
            Else
                FaktorKali = 1.2
            End If
            'gabungan setahun
            PengahasilanKenaPajak = CDec(Txt16Pjk_PKPSetahun)
            PPHTerhutangGabungan = 0
            RemarksPerhitunganPPHGabungan = ""
            Me.HitungPPHTerhutang(vTBLTarif, PPHTerhutangGabungan, PengahasilanKenaPajak, FaktorKali, RemarksPerhitunganPPHGabungan, ObjCD)
            Txt19Pjk_PPH21TerutangSetahun = PPHTerhutangGabungan
            'rutin setahun
            PengahasilanKenaPajak = CDec(Txt16Pjk_PKPSetahunRutin)
            PPHTerhutangRutin = 0
            RemarksPerhitunganPPHRutin = ""
            Me.HitungPPHTerhutang(vTBLTarif, PPHTerhutangRutin, PengahasilanKenaPajak, FaktorKali, RemarksPerhitunganPPHRutin, ObjCD)
            Txt19Pjk_PPH21TerutangSetahunRutin = PPHTerhutangRutin
            Txt19Pjk_PPH21TerutangSebulanRutin = Txt19Pjk_PPH21TerutangSetahunRutin
            'masukan item tax
            DescPPH = FrmTrans_Gaji_IncomeTax2.EnumDescPPHTerhutang.Pajak_Terhutang_Disetahunkan2
            Me.MasukkanItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No19_PPHTerutang, PPHTerhutangGabungan, "", 1, _
            0, RemarksPerhitunganPPHGabungan, _
            0, CInt(DescPPH).ToString, _
                TBLItemTax, vPrimaryTransaksi, vPrimaryEmployee, vTglTransaksi, vDocumentNo)

            DescPPH = FrmTrans_Gaji_IncomeTax2.EnumDescPPHTerhutang.Pajak_Terhutang_Penghasilan_Rutin1
            If PPHTerhutangRutin <> 0 Then
                'Me.MasukkanItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No19_PPHTerutang, 0, "", 1, _
                '0, "", PPHTerhutangRutin, RemarksPerhitunganPPHRutin, _
                'TBLItemTax, vPrimaryTransaksi, vPrimaryEmployee, vTglTransaksi, vDocumentNo)
                '#sinkron2018_pphcls
                Me.MasukkanItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No19_PPHTerutang, PPHTerhutangRutin, "", 1, _
                0, RemarksPerhitunganPPHRutin, 0, CInt(DescPPH).ToString, _
                TBLItemTax, vPrimaryTransaksi, vPrimaryEmployee, vTglTransaksi, vDocumentNo)
            Else
                Me.MasukkanItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No19_PPHTerutang, 0, "", 1, _
                0, "", 0, CInt(DescPPH).ToString, _
                TBLItemTax, vPrimaryTransaksi, vPrimaryEmployee, vTglTransaksi, vDocumentNo)
            End If
            'sebulan non rutin
            DescPPH = FrmTrans_Gaji_IncomeTax2.EnumDescPPHTerhutang.Pajak_Terhutang_Penghasilan_TidakRutin1
            Txt19Pjk_PPH21TerutangSebulanNonRutin = PPHTerhutangGabungan - PPHTerhutangRutin
            If CDec(Txt19Pjk_PPH21TerutangSebulanNonRutin) <> 0 Then
                'Me.MasukkanItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No19_PPHTerutang, 0, "", 1, _
                'CDec(Txt19Pjk_PPH21TerutangSebulanNonRutin), ObjCD.DigitGrouping(PPHTerhutangGabungan, 0) + " - " + ObjCD.DigitGrouping(PPHTerhutangRutin, 0), _
                '0, "", _
                'TBLItemTax, vPrimaryTransaksi, vPrimaryEmployee, vTglTransaksi, vDocumentNo)
                '#sinkron2018_pphcls
                Me.MasukkanItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No19_PPHTerutang, CDec(Txt19Pjk_PPH21TerutangSebulanNonRutin), "", 1, _
                0, ObjCD.DigitGrouping(PPHTerhutangGabungan, 0) + " - " + ObjCD.DigitGrouping(PPHTerhutangRutin, 0), _
                0, CInt(DescPPH).ToString, TBLItemTax, vPrimaryTransaksi, vPrimaryEmployee, vTglTransaksi, vDocumentNo)
            Else
                Me.MasukkanItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No19_PPHTerutang, 0, "", 1, _
                CDec(Txt19Pjk_PPH21TerutangSebulanNonRutin), CInt(DescPPH).ToString, _
                0, "", _
                TBLItemTax, vPrimaryTransaksi, vPrimaryEmployee, vTglTransaksi, vDocumentNo)
            End If

            'hitung pajak akhir tahun
            Txt17Pjk_PPH21 = "0"
            Txt17Pjk_PPH21Rutin = "0"

            Txt19Pjk_PPH21Terutang = Txt19Pjk_PPH21TerutangSetahun

            'kurangkan dengan PPH yang sudah dibayarkan pada masa sebelumnya
            DescPPH = FrmTrans_Gaji_IncomeTax2.EnumDescPPHTerhutang.Pajak_Penghasilan_Yang_Telah_Dibayarkan_masa_Sblm
            If Txt18_PPh21_yang_sdh_DisetorMasaSblm1 <> 0 Then
                Me.MasukkanItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No19_PPHTerutang, CDec(Txt18_PPh21_yang_sdh_DisetorMasaSblm1), "", 1, _
                0, "", 0, CInt(DescPPH).ToString, _
                TBLItemTax, vPrimaryTransaksi, vPrimaryEmployee, vTglTransaksi, vDocumentNo)
            Else
                Txt18_PPh21_yang_sdh_DisetorMasaSblm1 = 0
                Me.MasukkanItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No19_PPHTerutang, 0, "", 1, _
               0, "", 0, CInt(DescPPH).ToString, _
               TBLItemTax, vPrimaryTransaksi, vPrimaryEmployee, vTglTransaksi, vDocumentNo)
            End If

            'kurangkan dengan PPH yang sudah dibayarkan
            Txt18 = vPPhTerutangPeriodeSebelumnya
            DescPPH = FrmTrans_Gaji_IncomeTax2.EnumDescPPHTerhutang.Pajak_Penghasilan_Telah_Dilunasi
            If Txt18 <> 0 Then
                Me.MasukkanItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No19_PPHTerutang, CDec(Txt18), "", 1, _
                0, "", 0, CInt(DescPPH).ToString, _
                TBLItemTax, vPrimaryTransaksi, vPrimaryEmployee, vTglTransaksi, vDocumentNo)
            Else
                Txt18 = 0
                Me.MasukkanItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No19_PPHTerutang, 0, "", 1, _
               0, "", 0, CInt(DescPPH).ToString, _
               TBLItemTax, vPrimaryTransaksi, vPrimaryEmployee, vTglTransaksi, vDocumentNo)
            End If

            'lebih/kurang bayar
            DescPPH = FrmTrans_Gaji_IncomeTax2.EnumDescPPHTerhutang.Pajak_Penghasilan_Lebih_Dipotong
            If vEndPeriod <> 12 And vIsResign = True And vLosesTaxObligation = 1 Then
                TxtIncomeTax = ObjCD.DigitGrouping(Math.Round((CDec(Txt19Pjk_PPH21Terutang) / 12) * vPeriodCount, _
                0, MidpointRounding.AwayFromZero) - CDec(Txt18), 0)
                Debug.Print("F: PPH Terutang: " + ObjCD.DigitGrouping(Txt19Pjk_PPH21Terutang, 0))
                Debug.Print("F: PPH Terutang modified: " + ObjCD.DigitGrouping(Math.Round((CDec(Txt19Pjk_PPH21Terutang) / 12) * vPeriodCount, _
                0, MidpointRounding.AwayFromZero), 0))
                Debug.Print("F: PPH Yg sudah dibayar masa pajak sblm: " + ObjCD.DigitGrouping(Txt18, 0))
                Debug.Print("F: Kurang Bayar: " + ObjCD.DigitGrouping(TxtIncomeTax, 0))

                Dummy1 = "((Pajak Terhutang periode: " + Format(CDate(vStartPeriod.ToString + "/1/" + vTaxPeriod.Year.ToString), "MMM") + " s.d " _
                + Format(CDate(vEndPeriod.ToString + "/1/" + vTaxPeriod.Year.ToString), "MMM yyyy") + " * 12)/" + vPeriodCount.ToString + ")"

                If vEndPeriod > 1 And vStartPeriod <> vEndPeriod Then
                    Dummy1 = Dummy1 + " -  Pajak yang sudah dibayarkan periode: " _
                    + Format(CDate(vStartPeriod.ToString + "/1/" + vTaxPeriod.Year.ToString), "MMM") + " s.d " _
                    + Format(CDate((vEndPeriod - 1).ToString + "/1/" + vTaxPeriod.Year.ToString), "MMM yyyy")
                End If
                'If CDec(Txt18) > 0 Then
                '    Dummy1 = Dummy1 + "-  Pajak dibayarkan masa sebelumnya"
                'End If

                Me.MasukkanItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No19_PPHTerutang, Math.Round((CDec(Txt19Pjk_PPH21Terutang) / 12) * vPeriodCount, _
                0, MidpointRounding.AwayFromZero) - CDec(Txt18), "", 1, _
                0, Dummy1, 0, CInt(DescPPH).ToString, _
                TBLItemTax, vPrimaryTransaksi, vPrimaryEmployee, vTglTransaksi, vDocumentNo)
            Else
                TxtIncomeTax = ObjCD.DigitGrouping(CDec(Txt19Pjk_PPH21Terutang) - CDec(Txt18) - CDec(Txt18_PPh21_yang_sdh_DisetorMasaSblm1), 0)
                Me.MasukkanItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No19_PPHTerutang, CDec(Txt19Pjk_PPH21Terutang) - CDec(Txt18) - CDec(Txt18_PPh21_yang_sdh_DisetorMasaSblm1), "", 1, _
                0, "", 0, CInt(DescPPH).ToString, _
                    TBLItemTax, vPrimaryTransaksi, vPrimaryEmployee, vTglTransaksi, vDocumentNo)
            End If

            TxtKurangBayarTax = ObjCD.DigitGrouping(TxtIncomeTax)
            Txt20PPhTerutangYgSudahDilunasi = TxtIncomeTax
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub HitungPPhFinalPesangon_Aktual(ByVal vTunjanganPPHFinal As Decimal, _
    ByVal UseNPWP As Boolean, ByVal vTBLTarif As DseTransaksiSalary.tblmaster_tarif_pph_final_pesangonDataTable, ByVal JmlPesangon As Decimal, _
    ByVal vPrimaryTransaksi As Double, ByVal vPrimaryEmployee As Double, ByVal vTglTransaksi As Date, ByVal vDocumentNo As String, _
    ByVal TBLItemTax As DseTransaksiSalary.tbltrans_salary_itemtaxDataTable, ByVal ObjCD As Core_Module.ClsCekData, ByVal vSalary_EndPeriod As Date, _
    ByRef Txt21Pjk_Pesangon As Decimal, ByRef Txt21Pjk_TunjPPhPesangon As Decimal, _
    ByRef Txt21Pjk_PPhFinalPesangon As Decimal, ByRef vTarifPPhFinal As Decimal, ByRef Txt19Pjk_PPH21Terutang As Decimal, ByRef Txt18 As Decimal, _
    ByRef TxtIncomeTax As Decimal, ByRef TxtKurangBayarTax As Decimal, ByRef Txt20PPhTerutangYgSudahDilunasi As Decimal, ByVal vEndPeriod As Integer, _
    ByVal vIsResign As Boolean, ByVal vLosesTaxObligation As Integer, ByVal vPeriodCount As Integer, ByVal Txt18_PPh21_yang_sdh_DisetorMasaSblm1 As Decimal)
        Dim NPWPCategory As Integer = 0
        Dim vPesangon, PPHTerhutang As Decimal
        Dim FaktorKali As Decimal = 1
        Dim RemarksPerhitungan As String = ""


        Try
            'buang item tax hasil perhtungan
            Me.BuangItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No21_Pesangon, TBLItemTax)
            Me.BuangItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No21_PPHFinalPesangon, TBLItemTax)
            Me.BuangItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No21_TunjPPHPesangon, TBLItemTax)

            'masukkan data pesangon
            Txt21Pjk_Pesangon = JmlPesangon
            If JmlPesangon <> 0 Then
                Me.MasukkanItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No21_Pesangon, CDec(Txt21Pjk_Pesangon), "", _
                0, 0, "", 0, "", _
                TBLItemTax, vPrimaryTransaksi, vPrimaryEmployee, vTglTransaksi, vDocumentNo)
            End If

            'tunjangan pesangon
            Txt21Pjk_TunjPPhPesangon = vTunjanganPPHFinal
            'update data item
            If JmlPesangon <> 0 Then
                Me.MasukkanItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No21_TunjPPHPesangon, Txt21Pjk_TunjPPhPesangon, "", _
                0, 0, "", 0, "", _
                TBLItemTax, vPrimaryTransaksi, vPrimaryEmployee, vTglTransaksi, vDocumentNo)
            End If

            'hitung pph final terhutang
            Txt21Pjk_PPhFinalPesangon = 0
            If UseNPWP = True Then
                FaktorKali = 1
            Else
                FaktorKali = 1.2
            End If
            vPesangon = CDec(Txt21Pjk_Pesangon)
            PPHTerhutang = 0
            RemarksPerhitungan = ""
            vTarifPPhFinal = 0
            Me.HitungPPHFinalTerhutang_Pesangon(vTBLTarif, PPHTerhutang, vPesangon + vTunjanganPPHFinal, FaktorKali, RemarksPerhitungan, ObjCD, vTarifPPhFinal)
            Txt21Pjk_PPhFinalPesangon = PPHTerhutang
            'masukan item tax pp final
            Txt21Pjk_PPhFinalPesangon = ObjCD.RoundDown1(PPHTerhutang, 0)
            If JmlPesangon <> 0 Then
                Me.MasukkanItemTax(ClsTransaksiSalary_Item.EnumTaxElement.No21_PPHFinalPesangon, CDec(Txt21Pjk_PPhFinalPesangon), "", 1, _
                0, RemarksPerhitungan, _
                0, "", _
                TBLItemTax, vPrimaryTransaksi, vPrimaryEmployee, vTglTransaksi, vDocumentNo)
            End If

            If vSalary_EndPeriod.Month = 12 Then
                TxtIncomeTax = ObjCD.DigitGrouping(CDec(Txt19Pjk_PPH21Terutang) - CDec(Txt18) - CDec(Txt18_PPh21_yang_sdh_DisetorMasaSblm1) + PPHTerhutang, 0)
                TxtKurangBayarTax = ObjCD.DigitGrouping(TxtIncomeTax)
            Else
                If vEndPeriod <> 12 And vIsResign = True And vLosesTaxObligation = 1 Then
                    TxtIncomeTax = ObjCD.DigitGrouping(Math.Round((CDec(Txt19Pjk_PPH21Terutang) / 12) * vPeriodCount, _
                    0, MidpointRounding.AwayFromZero) - CDec(Txt18) - CDec(Txt18_PPh21_yang_sdh_DisetorMasaSblm1) + PPHTerhutang, 0)
                    TxtKurangBayarTax = ObjCD.DigitGrouping(TxtIncomeTax)
                Else
                    TxtIncomeTax = ObjCD.DigitGrouping(CDec(Txt19Pjk_PPH21Terutang) - CDec(Txt18) - CDec(Txt18_PPh21_yang_sdh_DisetorMasaSblm1) + PPHTerhutang, 0)
                    TxtKurangBayarTax = ObjCD.DigitGrouping(TxtIncomeTax)
                End If
            End If

            Txt20PPhTerutangYgSudahDilunasi = TxtIncomeTax
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HitungPPHFinalTerhutang_Pesangon(ByVal vTBLTarif As DseTransaksiSalary.tblmaster_tarif_pph_final_pesangonDataTable, ByRef vPPHTerhutang As Decimal, _
   ByRef vPesangon As Decimal, ByVal vFaktorKali As Decimal, ByRef vRemarksPerhitungan As String, ByVal ObjCD As Core_Module.ClsCekData, ByRef vTarifPPhFinal As Decimal)
        Dim vPenghasilanSudahHitungPajak As Decimal = 0
        Dim Dummy As String = ""

        Try
            vRemarksPerhitungan = ""
            For Each rowTarif As DseTransaksiSalary.tblmaster_tarif_pph_final_pesangonRow In vTBLTarif.Rows
                If vPenghasilanSudahHitungPajak < vPesangon Then
                    If rowTarif.maximum_income < vPesangon Then
                        vPPHTerhutang += (rowTarif.maximum_income - vPenghasilanSudahHitungPajak) * (rowTarif.tarif / 100) * vFaktorKali
                        If vPenghasilanSudahHitungPajak = 0 Then
                            Dummy = ObjCD.DigitGrouping(rowTarif.maximum_income, 0)
                        Else
                            Dummy = "(" + ObjCD.DigitGrouping(rowTarif.maximum_income, 0) + " - " + _
                            ObjCD.DigitGrouping(vPenghasilanSudahHitungPajak, 0) + ")"
                        End If
                        If vRemarksPerhitungan = "" Then
                            vRemarksPerhitungan = Dummy + " * " + ObjCD.DigitGrouping(rowTarif.tarif, 0) + "%"
                        Else
                            vRemarksPerhitungan += " + " + vbCrLf + Dummy + " * " + ObjCD.DigitGrouping(rowTarif.tarif, 0) + "%"
                        End If
                        If vFaktorKali <> 1 Then
                            vRemarksPerhitungan += " * " + ObjCD.DigitGrouping(vFaktorKali, 1)
                        End If
                    Else
                        vPPHTerhutang += (vPesangon - vPenghasilanSudahHitungPajak) * (rowTarif.tarif / 100) * vFaktorKali
                        If vPenghasilanSudahHitungPajak = 0 Then
                            Dummy = ObjCD.DigitGrouping(vPesangon, 0)
                        Else
                            Dummy = "(" + ObjCD.DigitGrouping(vPesangon, 0) + " - " + _
                            ObjCD.DigitGrouping(vPenghasilanSudahHitungPajak, 0) + ")"
                        End If
                        If vRemarksPerhitungan = "" Then
                            vRemarksPerhitungan = Dummy + " * " + ObjCD.DigitGrouping(rowTarif.tarif, 0) + "%"
                        Else
                            vRemarksPerhitungan += " + " + vbCrLf + Dummy + " * " + ObjCD.DigitGrouping(rowTarif.tarif, 0) + "%"
                        End If
                        If vFaktorKali <> 1 Then
                            vRemarksPerhitungan += " * " + ObjCD.DigitGrouping(vFaktorKali, 1)
                        End If
                    End If
                    vPenghasilanSudahHitungPajak = rowTarif.maximum_income
                    vTarifPPhFinal = rowTarif.tarif
                Else
                    Exit For
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HitungPPHTerhutang(ByVal vTBLTarif As DseTransaksiSalary.tblmaster_tarif_pphDataTable, ByRef vPPHTerhutangGabungan As Decimal, _
    ByRef vPengahasilanKenaPajak As Decimal, ByVal vFaktorKali As Decimal, ByRef vRemarksPerhitunganPPH As String, ByVal ObjCD As Core_Module.ClsCekData)
        Dim vPenghasilanSudahHitungPajak As Decimal = 0
        Dim Dummy As String = ""

        Try
            vRemarksPerhitunganPPH = ""
            For Each rowTarif As DseTransaksiSalary.tblmaster_tarif_pphRow In vTBLTarif.Rows
                If vPenghasilanSudahHitungPajak < vPengahasilanKenaPajak Then
                    If rowTarif.maximum_income < vPengahasilanKenaPajak Then
                        vPPHTerhutangGabungan += (rowTarif.maximum_income - vPenghasilanSudahHitungPajak) * (rowTarif.tarif / 100) * vFaktorKali
                        If vPenghasilanSudahHitungPajak = 0 Then
                            Dummy = ObjCD.DigitGrouping(rowTarif.maximum_income, 0)
                        Else
                            Dummy = "(" + ObjCD.DigitGrouping(rowTarif.maximum_income, 0) + " - " + _
                            ObjCD.DigitGrouping(vPenghasilanSudahHitungPajak, 0) + ")"
                        End If
                        If vRemarksPerhitunganPPH = "" Then
                            vRemarksPerhitunganPPH = Dummy + " * " + ObjCD.DigitGrouping(rowTarif.tarif, 0) + "%"
                        Else
                            vRemarksPerhitunganPPH += " + " + vbCrLf + Dummy + " * " + ObjCD.DigitGrouping(rowTarif.tarif, 0) + "%"
                        End If
                        If vFaktorKali <> 1 Then
                            vRemarksPerhitunganPPH += " * " + ObjCD.DigitGrouping(vFaktorKali, 1)
                        End If
                    Else
                        vPPHTerhutangGabungan += (vPengahasilanKenaPajak - vPenghasilanSudahHitungPajak) * (rowTarif.tarif / 100) * vFaktorKali
                        If vPenghasilanSudahHitungPajak = 0 Then
                            Dummy = ObjCD.DigitGrouping(vPengahasilanKenaPajak, 0)
                        Else
                            Dummy = "(" + ObjCD.DigitGrouping(vPengahasilanKenaPajak, 0) + " - " + _
                            ObjCD.DigitGrouping(vPenghasilanSudahHitungPajak, 0) + ")"
                        End If
                        If vRemarksPerhitunganPPH = "" Then
                            vRemarksPerhitunganPPH = Dummy + " * " + ObjCD.DigitGrouping(rowTarif.tarif, 0) + "%"
                        Else
                            vRemarksPerhitunganPPH += " + " + vbCrLf + Dummy + " * " + ObjCD.DigitGrouping(rowTarif.tarif, 0) + "%"
                        End If
                        If vFaktorKali <> 1 Then
                            vRemarksPerhitunganPPH += " * " + ObjCD.DigitGrouping(vFaktorKali, 1)
                        End If
                    End If
                    vPenghasilanSudahHitungPajak = rowTarif.maximum_income
                Else
                    Exit For
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetAmountRutin(ByVal vTipeElemenPajak As ClsTransaksiSalary_Item.EnumTaxElement, ByVal TBLItemTax As DseTransaksiSalary.tbltrans_salary_itemtaxDataTable) As Decimal
        Dim DV As DataRowView
        Dim row As DseTransaksiSalary.tbltrans_salary_itemtaxRow
        Dim RetVal As Decimal
        Try
            TBLItemTax.DefaultView.RowFilter = "[tipe tax] = " + CInt(vTipeElemenPajak).ToString
            For Each DV In TBLItemTax.DefaultView
                row = DV.Row
                RetVal += row.amount_rutin_disetahunkan
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetAmountGabungan(ByVal vTipeElemenPajak As ClsTransaksiSalary_Item.EnumTaxElement, ByVal TBLItemTax As DseTransaksiSalary.tbltrans_salary_itemtaxDataTable) As Decimal
        Dim DV As DataRowView
        Dim row As DseTransaksiSalary.tbltrans_salary_itemtaxRow
        Dim RetVal As Decimal

        Try
            TBLItemTax.DefaultView.RowFilter = "[tipe tax] = " + CInt(vTipeElemenPajak).ToString
            For Each DV In TBLItemTax.DefaultView
                row = DV.Row
                RetVal += row.amount_gabungan_disetahunkan
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub MasukkanItemTax(ByVal vTipeElemenPajak As ClsTransaksiSalary_Item.EnumTaxElement, ByVal vAmountMonth As Decimal, _
  ByVal vIdentifier As String, ByVal vIsRutin As Integer, ByVal vAmountGabunganDisetahunkan As Decimal, ByVal vRemarksAmountGabungan As String, _
  ByVal vAmountRutinDisetahunkan As Decimal, ByVal vRemarksRutinDisetahunkan As String, ByVal TBLItemTax As DseTransaksiSalary.tbltrans_salary_itemtaxDataTable, _
  ByVal vPrimaryTransaksi As Double, ByVal vPrimaryEmployee As Double, ByVal vTanggalTransaksi As Date, ByVal vNoDocument As String)
        Dim rowItemTax As DseTransaksiSalary.tbltrans_salary_itemtaxRow
        Try
            rowItemTax = TBLItemTax.Newtbltrans_salary_itemtaxRow
            With rowItemTax
                .tipe_tax = CInt(vTipeElemenPajak)
                .amount_month = vAmountMonth

                .identifier_salary_component = vIdentifier
                .primary = Me.GetPrimaryMax(TBLItemTax) + 1
                .primary_transaksi = vPrimaryTransaksi
                .is_rutin = vIsRutin
                .amount_gabungan_disetahunkan = vAmountGabunganDisetahunkan
                .remarks_gabungan = vRemarksAmountGabungan

                .amount_rutin_disetahunkan = vAmountRutinDisetahunkan
                .remarks_rutin = vRemarksRutinDisetahunkan

                .No_Transaksi = vNoDocument
                .Tanggal_Transaksi = vTanggalTransaksi
                .Primary_Employee = vPrimaryEmployee
            End With
            TBLItemTax.Rows.Add(rowItemTax)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetPrimaryMax(ByVal vTBLItemTax As DseTransaksiSalary.tbltrans_salary_itemtaxDataTable) As Double
        Dim rowItemTax As DseTransaksiSalary.tbltrans_salary_itemtaxRow
        Dim RetVal As Double = Double.MinValue

        Try
            For Each rowItemTax In vTBLItemTax.Rows
                If RetVal < rowItemTax.primary Then
                    RetVal = rowItemTax.primary
                End If
            Next

            If RetVal = Double.MinValue Then
                Return 0
            Else
                Return RetVal
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub BuangItemTax(ByVal vTipeElemenPajak As ClsTransaksiSalary_Item.EnumTaxElement, ByVal TBLItemTax As DseTransaksiSalary.tbltrans_salary_itemtaxDataTable)
        Dim DV As DataRowView
        Dim row As DseTransaksiSalary.tbltrans_salary_itemtaxRow

        Try
            TBLItemTax.DefaultView.RowFilter = "[tipe tax] = " + CInt(vTipeElemenPajak).ToString
            For Each DV In TBLItemTax.DefaultView
                row = DV.Row
                TBLItemTax.Rows.Remove(row)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
