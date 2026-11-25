Option Explicit On
Option Strict Off

Public Class ClsMasterKaryawan
    Dim ObjCD As New Core_Module.ClsCekData

    Dim mPrimary As Double = -1
    Dim mCode As String = ""
    Dim mFullName As String = ""
    Dim mPrefix As String = ""
    Dim mFristName As String = ""
    Dim mMiddleName As String = ""
    Dim mLastName As String = ""
    Dim mSuffixName As String = ""
    Dim mNIK As String = ""
    Dim mRemark As String = ""
    Dim mAktif As EnumAktif = EnumAktif.Active
    Dim mResignDate As Date = #1/1/1950#     'aul tambahan
    Dim mCreatedBy As String = ""
    Dim mCreatedDate As DateTime = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As DateTime = Now
    Dim mTimestamp As Date = Now
    Dim mPrimaryDep As Double = -1
    Dim mPicturePacth As String = ""
    Dim mAdmisionDate As Date = #1/1/1950#
    Dim mAppointmentDate As Date = #1/1/1950#
    Dim mPrimaryPosition As Double = -1
    Dim mKPA_Astek_No As String = ""
    Dim mKPJ_no As String = ""
    Dim mId_Card_no As String = ""
    Dim mId_Card_Exp As Date
    Dim mId_Card_Add As String = ""
    Dim mNPWP_No As String = ""
    Dim mNPWP_Category As GolonganPajak = GolonganPajak.TK0
    Dim mNPWP_Add As String = ""
    Dim mMartial_Status As EnumMartialStatus = EnumMartialStatus._Single
    Dim mMariageDAte As Date
    Dim mEducation As String = ""
    Dim mSex As EnumSex = EnumSex.Male
    Dim mCurentAdd As String = ""
    Dim mReligion As EnumReligion = EnumReligion.Islam
    Dim mPhoneNo As String = ""
    Dim mEmployStatus As EnumEmployStatus = EnumEmployStatus.Probation
    Dim mBrithPalce As String = ""
    Dim mDateBrith As Date
    Dim mBloodType As EnumBloodType = EnumBloodType.A
    Dim mNationality As String = ""
    Dim mEmail As String = ""
    Dim mMotherName As String = ""
    Dim mMotherAddres As String = ""
    Dim mBankCode As String = ""    'aul tambahan
    Dim mBankName As String = ""
    Dim mBankBranch As String = ""
    Dim mBankAccountNo As String = ""
    Dim mBankAccountHolder As String = ""
    Dim mBalaiPengobatanJPK As String = ""
    Dim mSalaryFormula As String = ""
    Dim mIsNPWP As Integer = ClsMasterKaryawan.EnumIsNPWP.No
    Dim mForeignEmployee As EnumForeign = EnumForeign.No
    Dim mPrimaryCountry As Double = -1
    Dim mFingerprintID As String = ""
    Dim mOverrideClockIn As Integer = 0
    Dim mDoesNotClockIn As Integer = 0
    Dim mOverrideClockOut As Integer
    Dim mDoesNotClockOut As Integer
    Private mOverrideEarlyOT As Integer = 0
    Private mIsCalculateEarlyOT As Integer = 1
    Private mOverrideLate As Integer = 0
    Private mIsCalculateLate As Integer = 1
    Private mOverrideEarlyOut As Integer = 0
    Private mIsCalculateEarlyOut As Integer = 1
    Private mOverrideOT As Integer = 0
    Private mUserOvertimeOrder As Integer = 0
    Private mIsCalculateOT As Integer = 1
    Dim mFirstTaxPeriod As Date = #1/1/1950#

    Dim mCalculateAfterEarlyOT As Double = 0
    Dim mCalculateAfterLate As Double = 0
    Dim mCalculateAfterEarlyOut As Double = 0
    Dim mCalculateAfterOT As Double = 0
    Dim mRoundingTypeEarlyOT As Integer = 0
    Dim mRoundingTypeLate As Integer = 0
    Dim mRoundingTypeEarlyOut As Integer = 0
    Dim mRoundingTypeOT As Integer = 0
    Dim mRoundingValueEarlyOT As Decimal = 0
    Dim mRoundingValueLate As Decimal = 0
    Dim mRoundingValueEarlyOut As Decimal = 0
    Dim mRoundingValueOT As Decimal = 0
    Dim mMaxValueEarlyOT As Decimal = 0
    Dim mMaxValueLate As Decimal = 0
    Dim mMaxValueEarlyOut As Decimal = 0
    Dim mMaxValueOT As Decimal = 0

    Dim mLosesTaxObligations As Integer = 0

    'digunakan untuk antrian simpan data
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianMaster
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAAntrian As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter
    Dim ObjRnd As New Random
    Dim ObjAntrian As New Core_Module.ClsAntrian

    Dim objitem As New ClsMasterKaryawan_Item

    Public Enum EnumDefaultFindBy As Integer
        ByCode
        ByName
    End Enum

    Public Enum EnumDefaultFindBy_SettingUser As Integer
        ByGlobalSetting
        ByCode
        ByName
    End Enum

    Public Property LosesTaxObligation() As Integer
        Get
            Return Me.mLosesTaxObligations
        End Get
        Set(ByVal value As Integer)
            Me.mLosesTaxObligations = value
        End Set
    End Property

    Public Property FirstTaxPeriod() As Date
        Get
            Return Me.mFirstTaxPeriod
        End Get
        Set(ByVal value As Date)
            mFirstTaxPeriod = value
        End Set
    End Property

  
    Public Property OverrideEarlyOT() As Integer
        Get
            Return mOverrideEarlyOT
        End Get
        Set(ByVal value As Integer)
            mOverrideEarlyOT = value
        End Set
    End Property

    Public Property IsCalculateEarlyOT() As Integer
        Get
            Return mIsCalculateEarlyOT
        End Get
        Set(ByVal value As Integer)
            mIsCalculateEarlyOT = value
        End Set
    End Property

    Public Property OverrideLate() As Integer
        Get
            Return mOverrideLate
        End Get
        Set(ByVal value As Integer)
            mOverrideLate = value
        End Set
    End Property

    Public Property IsCalculateLate() As Integer
        Get
            Return mIsCalculateLate
        End Get
        Set(ByVal value As Integer)
            mIsCalculateLate = value
        End Set
    End Property

    Public Property OverrideEarlyOut() As Integer
        Get
            Return mOverrideEarlyOut
        End Get
        Set(ByVal value As Integer)
            mOverrideEarlyOut = value
        End Set
    End Property

    Public Property IsCalculateEarlyOut() As Integer
        Get
            Return mIsCalculateEarlyOut
        End Get
        Set(ByVal value As Integer)
            mIsCalculateEarlyOut = value
        End Set
    End Property

    Public Property OverrideOT() As Integer
        Get
            Return mOverrideOT
        End Get
        Set(ByVal value As Integer)
            mOverrideOT = value
        End Set
    End Property

    Public Property IsCalculateOT() As Integer
        Get
            Return mIsCalculateOT
        End Get
        Set(ByVal value As Integer)
            mIsCalculateOT = value
        End Set
    End Property

    Public Property CalculateAfterEarlyOT() As Double
        Get
            Return Me.mCalculateAfterEarlyOT
        End Get
        Set(ByVal value As Double)
            Me.mCalculateAfterEarlyOT = value
        End Set
    End Property

    Public Property CalculateAfterLate() As Double
        Get
            Return Me.mCalculateAfterLate
        End Get
        Set(ByVal value As Double)
            Me.mCalculateAfterLate = value
        End Set
    End Property

    Public Property CalculateAfterEarlyOut() As Double
        Get
            Return Me.mCalculateAfterEarlyOut
        End Get
        Set(ByVal value As Double)
            Me.mCalculateAfterEarlyOut = value
        End Set
    End Property

    Public Property CalculateAfterOT() As Double
        Get
            Return Me.mCalculateAfterOT
        End Get
        Set(ByVal value As Double)
            Me.mCalculateAfterOT = value
        End Set
    End Property


    Public Property RoundingTypeEarlyOT() As Integer
        Get
            Return Me.mRoundingTypeEarlyOT
        End Get
        Set(ByVal value As Integer)
            Me.mRoundingTypeEarlyOT = value
        End Set
    End Property

    Public Property RoundingTypeLate() As Integer
        Get
            Return Me.mRoundingTypeLate
        End Get
        Set(ByVal value As Integer)
            Me.mRoundingTypeLate = value
        End Set
    End Property

    Public Property RoundingTypeEarlyOut() As Integer
        Get
            Return Me.mRoundingTypeEarlyOut
        End Get
        Set(ByVal value As Integer)
            Me.mRoundingTypeEarlyOut = value
        End Set
    End Property

    Public Property RoundingTypeOT() As Integer
        Get
            Return Me.mRoundingTypeOT
        End Get
        Set(ByVal value As Integer)
            Me.mRoundingTypeOT = value
        End Set
    End Property


    Public Property RoundingValueEarlyOT() As Decimal
        Get
            Return Me.mRoundingValueEarlyOT
        End Get
        Set(ByVal value As Decimal)
            Me.mRoundingValueEarlyOT = value
        End Set
    End Property

    Public Property RoundingValueLate() As Decimal
        Get
            Return Me.mRoundingValueLate
        End Get
        Set(ByVal value As Decimal)
            Me.mRoundingValueLate = value
        End Set
    End Property

    Public Property RoundingValueEarlyOut() As Decimal
        Get
            Return Me.mRoundingValueEarlyOut
        End Get
        Set(ByVal value As Decimal)
            Me.mRoundingValueEarlyOut = value
        End Set
    End Property

    Public Property RoundingValueOT() As Decimal
        Get
            Return Me.mRoundingValueOT
        End Get
        Set(ByVal value As Decimal)
            Me.mRoundingValueOT = value
        End Set
    End Property


    Public Property MaxValueEarlyOT() As Decimal
        Get
            Return Me.mMaxValueEarlyOT
        End Get
        Set(ByVal value As Decimal)
            Me.mMaxValueEarlyOT = value
        End Set
    End Property

    Public Property MaxValueLate() As Decimal
        Get
            Return Me.mMaxValueLate
        End Get
        Set(ByVal value As Decimal)
            Me.mMaxValueLate = value
        End Set
    End Property

    Public Property MaxValueEarlyOut() As Decimal
        Get
            Return Me.mMaxValueEarlyOut
        End Get
        Set(ByVal value As Decimal)
            Me.mMaxValueEarlyOut = value
        End Set
    End Property

    Public Property MaxValueOT() As Decimal
        Get
            Return Me.mMaxValueOT
        End Get
        Set(ByVal value As Decimal)
            Me.mMaxValueOT = value
        End Set
    End Property

    Private mUseOtDuration As Integer = 0
    Public Property UseOtDuration() As Integer
        Get
            Return Me.mUseOtDuration
        End Get
        Set(ByVal value As Integer)
            Me.mUseOtDuration = value
        End Set
    End Property

    Private mNotificationEmailOt As Integer = 0
    Public Property NotificationEmailOt() As Integer
        Get
            Return Me.mNotificationEmailOt
        End Get
        Set(ByVal value As Integer)
            Me.mNotificationEmailOt = value
        End Set
    End Property

    Private mOvverideOTOrder As Integer = 0
    Public Property OvverideOTOrder() As Integer
        Get
            Return Me.mOvverideOTOrder
        End Get
        Set(ByVal value As Integer)
            Me.mOvverideOTOrder = value
        End Set
    End Property

#Region "Property"
    Public Property ResignDate() As Date
        Get
            Return Me.mResignDate
        End Get
        Set(ByVal value As Date)
            mResignDate = value
        End Set
    End Property
    Public Property AppointmentDate() As Date
        Get
            Return Me.mAppointmentDate
        End Get
        Set(ByVal value As Date)
            mAppointmentDate = value
        End Set
    End Property
    Public Property PrimaryCountry() As Double
        Get
            Return mPrimaryCountry
        End Get
        Set(ByVal value As Double)
            mPrimaryCountry = value
        End Set
    End Property

    Public Property ForeignEmployee() As EnumForeign
        Get
            Return mForeignEmployee
        End Get
        Set(ByVal value As EnumForeign)
            mForeignEmployee = value
        End Set
    End Property

    Public Property IsNPWP() As Integer
        Get
            Return mIsNPWP
        End Get
        Set(ByVal value As Integer)
            Me.mIsNPWP = value
        End Set
    End Property

    Public Property Primary() As Double
        Get
            Return Me.mPrimary
        End Get
        Set(ByVal value As Double)
            Me.mPrimary = value
        End Set
    End Property

    Public Property Code() As String
        Get
            Return Me.mCode
        End Get
        Set(ByVal value As String)
            Me.mCode = value
        End Set
    End Property

    Public Property FullName() As String
        Get
            Return Me.mFullName
        End Get
        Set(ByVal value As String)
            Me.mFullName = value
        End Set
    End Property

    Public Property Prepix() As String
        Get
            Return Me.mPrefix
        End Get
        Set(ByVal value As String)
            Me.mPrefix = value
        End Set
    End Property

    Public Property FristName() As String
        Get
            Return Me.mFristName
        End Get
        Set(ByVal value As String)
            Me.mFristName = value
        End Set
    End Property

    Public Property MiddleName() As String
        Get
            Return Me.mMiddleName
        End Get
        Set(ByVal value As String)
            Me.mMiddleName = value
        End Set
    End Property

    Public Property LastName() As String
        Get
            Return Me.mLastName
        End Get
        Set(ByVal value As String)
            Me.mLastName = value
        End Set
    End Property

    Public Property Suffix() As String
        Get
            Return Me.mSuffixName
        End Get
        Set(ByVal value As String)
            Me.mSuffixName = value
        End Set
    End Property

    Public Property NIK() As String
        Get
            Return Me.mNIK
        End Get
        Set(ByVal value As String)
            Me.mNIK = value
        End Set
    End Property

    Public Property remark() As String
        Get
            Return Me.mRemark
        End Get
        Set(ByVal value As String)
            Me.mRemark = value
        End Set
    End Property

    Public Property Aktif() As EnumAktif
        Get
            Return Me.mAktif
        End Get
        Set(ByVal value As EnumAktif)
            Me.mAktif = value
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

    Public Property PrimaryDep() As Double
        Get
            Return Me.mPrimaryDep
        End Get
        Set(ByVal value As Double)
            Me.mPrimaryDep = value
        End Set
    End Property

    Public Property PicturePath() As String
        Get
            Return Me.mPicturePacth
        End Get
        Set(ByVal value As String)
            Me.mPicturePacth = value
        End Set
    End Property

    Public Property AdmisionDate() As String
        Get
            Return Me.mAdmisionDate
        End Get
        Set(ByVal value As String)
            Me.mAdmisionDate = value
        End Set
    End Property

    Public Property PrimaryPosition() As Double
        Get
            Return Me.mPrimaryPosition
        End Get
        Set(ByVal value As Double)
            Me.mPrimaryPosition = value
        End Set
    End Property

    Public Property KPA_Astek_No() As String
        Get
            Return Me.mKPA_Astek_No
        End Get
        Set(ByVal value As String)
            Me.mKPA_Astek_No = value
        End Set
    End Property

    Public Property KPJ_no() As String
        Get
            Return Me.mKPJ_no
        End Get
        Set(ByVal value As String)
            Me.mKPJ_no = value
        End Set
    End Property

    Public Property Id_Card_no() As String
        Get
            Return Me.mId_Card_no
        End Get
        Set(ByVal value As String)
            Me.mId_Card_no = value
        End Set
    End Property

    Public Property Id_Card_Exp() As Date
        Get
            Return Me.mId_Card_Exp
        End Get
        Set(ByVal value As Date)
            Me.mId_Card_Exp = value
        End Set
    End Property

    Public Property Id_Card_Add() As String
        Get
            Return Me.mId_Card_Add
        End Get
        Set(ByVal value As String)
            Me.mId_Card_Add = value
        End Set
    End Property

    Public Property NPWP_No() As String
        Get
            Return Me.mNPWP_No
        End Get
        Set(ByVal value As String)
            Me.mNPWP_No = value
        End Set
    End Property

    Public Property NPWP_Category() As GolonganPajak
        Get
            Return Me.mNPWP_Category
        End Get
        Set(ByVal value As GolonganPajak)
            Me.mNPWP_Category = value
        End Set
    End Property

    Public Property NPWP_Add() As String
        Get
            Return Me.mNPWP_Add
        End Get
        Set(ByVal value As String)
            Me.mNPWP_Add = value
        End Set
    End Property

    Public Property MartialStatus() As EnumMartialStatus
        Get
            Return Me.mMartial_Status
        End Get
        Set(ByVal value As EnumMartialStatus)
            Me.mMartial_Status = value
        End Set
    End Property

    Public Property MariageDAte() As Date
        Get
            Return Me.mMariageDAte
        End Get
        Set(ByVal value As Date)
            Me.mMariageDAte = value
        End Set
    End Property

    Public Property Education() As String
        Get
            Return Me.mEducation
        End Get
        Set(ByVal value As String)
            Me.mEducation = value
        End Set
    End Property

    Public Property Sex() As EnumSex
        Get
            Return Me.mSex
        End Get
        Set(ByVal value As EnumSex)
            Me.mSex = value
        End Set
    End Property

    Public Property CurentAdd() As String
        Get
            Return Me.mCurentAdd
        End Get
        Set(ByVal value As String)
            Me.mCurentAdd = value
        End Set
    End Property

    Public Property Religion() As EnumReligion
        Get
            Return Me.mReligion
        End Get
        Set(ByVal value As EnumReligion)
            Me.mReligion = value
        End Set
    End Property

    Public Property PhoneNo() As String
        Get
            Return Me.mPhoneNo
        End Get
        Set(ByVal value As String)
            Me.mPhoneNo = value
        End Set
    End Property

    Public Property EmployStatus() As EnumEmployStatus
        Get
            Return Me.mEmployStatus
        End Get
        Set(ByVal value As EnumEmployStatus)
            Me.mEmployStatus = value
        End Set
    End Property

    Public Property BrithPalce() As String
        Get
            Return Me.mBrithPalce
        End Get
        Set(ByVal value As String)
            Me.mBrithPalce = value
        End Set
    End Property

    Public Property DateBrith() As Date
        Get
            Return Me.mDateBrith
        End Get
        Set(ByVal value As Date)
            Me.mDateBrith = value
        End Set
    End Property

    Public Property BloodType() As EnumBloodType
        Get
            Return Me.mBloodType
        End Get
        Set(ByVal value As EnumBloodType)
            Me.mBloodType = value
        End Set
    End Property

    Public Property Nationality() As String
        Get
            Return Me.mNationality
        End Get
        Set(ByVal value As String)
            Me.mNationality = value
        End Set
    End Property

    Public Property Email() As String
        Get
            Return Me.mEmail
        End Get
        Set(ByVal value As String)
            Me.mEmail = value
        End Set
    End Property

    Public Property MotherName() As String
        Get
            Return Me.mMotherName
        End Get
        Set(ByVal value As String)
            Me.mMotherName = value
        End Set
    End Property

    Public Property MotherAddres() As String
        Get
            Return Me.mMotherAddres
        End Get
        Set(ByVal value As String)
            Me.mMotherAddres = value
        End Set
    End Property

    'aul tambahan
    Public Property BankCode() As String
        Get
            Return Me.mBankCode
        End Get
        Set(ByVal value As String)
            Me.mBankCode = value
        End Set
    End Property

    Public Property BankName() As String
        Get
            Return Me.mBankName
        End Get
        Set(ByVal value As String)
            Me.mBankName = value
        End Set
    End Property

    Public Property BankBranch() As String
        Get
            Return Me.mBankBranch
        End Get
        Set(ByVal value As String)
            Me.mBankBranch = value
        End Set
    End Property

    Public Property BankAccountNo() As String
        Get
            Return Me.mBankAccountNo
        End Get
        Set(ByVal value As String)
            Me.mBankAccountNo = value
        End Set
    End Property

    Public Property BankAccountHolder() As String
        Get
            Return Me.mBankAccountHolder
        End Get
        Set(ByVal value As String)
            Me.mBankAccountHolder = value
        End Set
    End Property

    Public Property BalaiPengobatanJPK() As String
        Get
            Return Me.mBalaiPengobatanJPK
        End Get
        Set(ByVal value As String)
            Me.mBalaiPengobatanJPK = value
        End Set
    End Property

    Public Property SalaryFormula() As String
        Get
            Return Me.mSalaryFormula
        End Get
        Set(ByVal value As String)
            Me.mSalaryFormula = value
        End Set
    End Property

    Public Property FingerprintID() As String
        Get
            Return mFingerprintID
        End Get
        Set(ByVal value As String)
            mFingerprintID = value
        End Set
    End Property

    Public Property OverrideClockIn() As Integer
        Get
            Return mOverrideClockIn
        End Get
        Set(ByVal value As Integer)
            mOverrideClockIn = value
        End Set
    End Property

    Public Property DoesNotClockIn() As Integer
        Get
            Return mDoesNotClockIn
        End Get
        Set(ByVal value As Integer)
            mDoesNotClockIn = value
        End Set
    End Property

    Public Property OverrideClockOut() As Integer
        Get
            Return mOverrideClockOut
        End Get
        Set(ByVal value As Integer)
            mOverrideClockOut = value
        End Set
    End Property

    Public Property DoesNotClockOut() As Integer
        Get
            Return mDoesNotClockOut
        End Get
        Set(ByVal value As Integer)
            mDoesNotClockOut = value
        End Set
    End Property


    Private mOverrideCalculateEarlyOTTolerance As Integer
    Public Property OverrideCalculateEarlyOTTolerance() As Integer
        Get
            Return mOverrideCalculateEarlyOTTolerance
        End Get
        Set(ByVal value As Integer)
            mOverrideCalculateEarlyOTTolerance = value
        End Set
    End Property

    Private mOverrideCalculateLateTolerance As Integer
    Public Property OverrideCalculateLateTolerance() As Integer
        Get
            Return mOverrideCalculateLateTolerance
        End Get
        Set(ByVal value As Integer)
            mOverrideCalculateLateTolerance = value
        End Set
    End Property

    Private mOverrideCalculateEarlyOutTolerance As Integer
    Public Property OverrideCalculateEarlyOutTolerance() As Integer
        Get
            Return mOverrideCalculateEarlyOutTolerance
        End Get
        Set(ByVal value As Integer)
            mOverrideCalculateEarlyOutTolerance = value
        End Set
    End Property

    Private mOverrideCalculateOTTolerance As Integer
    Public Property OverrideCalculateOTTolerance() As Integer
        Get
            Return mOverrideCalculateOTTolerance
        End Get
        Set(ByVal value As Integer)
            mOverrideCalculateOTTolerance = value
        End Set
    End Property

    Private mCalculateEarlyOTOnWeekend As Integer = 0
    Public Property CalculateEarlyOTOnWeekend() As Integer
        Get
            Return mCalculateEarlyOTOnWeekend
        End Get
        Set(ByVal value As Integer)
            mCalculateEarlyOTOnWeekend = value
        End Set
    End Property

    Private mCalculateLateOnWeekend As Integer = 0
    Public Property CalculateLateOnWeekend() As Integer
        Get
            Return mCalculateLateOnWeekend
        End Get
        Set(ByVal value As Integer)
            mCalculateLateOnWeekend = value
        End Set
    End Property

    Private mCalculateEarlyOutOnWeekend As Integer = 0
    Public Property CalculateEarlyOutOnWeekend() As Integer
        Get
            Return mCalculateEarlyOutOnWeekend
        End Get
        Set(ByVal value As Integer)
            mCalculateEarlyOutOnWeekend = value
        End Set
    End Property

    Private mCalculateOTOnWeekend As Integer = 1
    Public Property CalculateOTOnWeekend() As Integer
        Get
            Return mCalculateOTOnWeekend
        End Get
        Set(ByVal value As Integer)
            mCalculateOTOnWeekend = value
        End Set
    End Property

    Private mCalculateEarlyOTOnHoliday As Integer = 0
    Public Property CalculateEarlyOTOnHoliday() As Integer
        Get
            Return mCalculateEarlyOTOnHoliday
        End Get
        Set(ByVal value As Integer)
            mCalculateEarlyOTOnHoliday = value
        End Set
    End Property

    Private mCalculateLateOnHoliday As Integer = 0
    Public Property CalculateLateOnHoliday() As Integer
        Get
            Return mCalculateLateOnHoliday
        End Get
        Set(ByVal value As Integer)
            mCalculateLateOnHoliday = value
        End Set
    End Property

    Private mCalculateEarlyOutOnHoliday As Integer = 0
    Public Property CalculateEarlyOutOnHoliday() As Integer
        Get
            Return mCalculateEarlyOutOnHoliday
        End Get
        Set(ByVal value As Integer)
            mCalculateEarlyOutOnHoliday = value
        End Set
    End Property

    Private mCalculateOTOnHoliday As Integer = 1
    Public Property CalculateOTOnHoliday() As Integer
        Get
            Return mCalculateOTOnHoliday
        End Get
        Set(ByVal value As Integer)
            mCalculateOTOnHoliday = value
        End Set
    End Property

#End Region

#Region "Enum"

    Public Enum enumRoundingType As Integer
        No_Rounding
        Round_Up
        Round_down
        Middle
    End Enum

    Public Enum EnumForeign As Integer
        No
        Yes
    End Enum

    Public Enum EnumIsNPWP As Integer
        No
        Yes
    End Enum

    Public Enum EnumBloodType As Integer
        A
        B
        AB
        O
    End Enum

    Public Enum EnumEmployStatus As Integer
        Probation
        Contract
        Jobholder
    End Enum

    Public Enum EnumReligion As Integer
        Islam
        Katolik
        Kristen
        Hindu
        Buddha
        Konghucu
        Others
    End Enum

    Public Enum EnumSex As Integer
        Male
        Female
    End Enum

    Public Enum EnumMartialStatus As Integer
        _Single
        Married
    End Enum

    Public Enum EnumAktif As Integer
        Active
        NonActive
    End Enum

    Public Enum GolonganPajak As Integer
        TK0
        TK1
        TK2
        TK3
        K0
        K1
        K2
        K3
    End Enum

    Public Enum EnumSearchCriteriaEmployee As Integer
        ByCode
        ByName
        ByDepartement
        ByPosition
        ByNote
        ByResignDate
        ByFirstAttendanceDate
    End Enum
#End Region

    Public Overridable Function ApakahBolehEdit(ByVal vPrimary As Double, ByVal vUser As String) As Boolean
        'Dim dta As New DseMasterKaryawanTableAdapters.tbltrans_cuti_pengajuan_settingTableAdapter
        'Dim dta1 As New DseMasterKaryawanTableAdapters.tblmaster_salary_formula_component_detailemployeeTableAdapter
        'Dim Dummy As Nullable(Of Long)
        'Dim DummyUser As String
        Try
            ''tambahan dari leave app setting
            'dta.Connection.ConnectionString = Core_Module.GetConstring
            'dta1.Connection.ConnectionString = dta.Connection.ConnectionString 

            'Dummy = dta.ScalarQuery_GetJmlByPrimaryKaryawan(vPrimary)
            'If Dummy.HasValue = True Then
            '    If Dummy.Value > 0 Then
            '        MsgBox("Employee data cannot edit because this data is refering to transaction leave application setting.", _
            '        MsgBoxStyle.Information, "Information")
            '        Return False
            '    End If
            'End If

            ''salary formula
            'Dummy = dta1.ScalarQuery_GetJmlByPrimaryKaryawan(vPrimary)
            'If Dummy.HasValue = True Then
            '    If Dummy.Value > 0 Then
            '        MsgBox("Employee data cannot edit because this data is refering to salary formula.", _
            '        MsgBoxStyle.Information, "Information")
            '        Return False
            '    End If
            'End If

            'dta.Connection.ConnectionString = Core_Module.GetConstring
            'strict edit
            'DummyUser = dta.ScalarQuery_GetCreated(vPrimary)
            'If DummyUser <> vUser Then
            '    MsgBox("Only " + DummyUser + " user is allowed to edit this transaction.", _
            '    MsgBoxStyle.Information, "Information")
            '    Return False
            'End If

            'benar
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            'dta.Dispose()
            'dta1.Dispose()
        End Try
    End Function

    Public Function CekMasterLeave(ByVal vPrimary As Double) As Double
        Dim DTA As New DseMasterKaryawanTableAdapters.tblmaster_cuti_itemTableAdapter
        Dim Dummy As System.Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetJmlByEmployee(vPrimary)
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

    Public Function CekMasterWorkshiftAllocate(ByVal vPrimary As Double) As Double
        Dim DTA As New DseMasterKaryawanTableAdapters.tblmaster_workshift_itemallocateTableAdapter
        Dim Dummy As System.Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetJmlByEmployee(vPrimary)
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


    Public Function ApakahBolehHapus(ByVal vPrimary As Double, ByVal vUser As String) As Boolean
        Dim dta As New DseMasterKaryawanTableAdapters.tbltrans_cuti_pengajuan_settingTableAdapter
        Dim dta1 As New DseMasterKaryawanTableAdapters.tblmaster_salary_formula_component_detailemployeeTableAdapter
        Dim dta2 As New DseMasterKaryawanTableAdapters.tblmaster_karyawanTableAdapter

        Dim Dummy As Nullable(Of Long)

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dta1.Connection.ConnectionString = Core_Module.GetConstring
            dta2.Connection.ConnectionString = Core_Module.GetConstring

            '----------------Berlaku Umum-------------------
            'tambahan dari master leave 
            If Me.CekMasterLeave(vPrimary) > 0 Then
                MsgBox("Employee data can not be deleted because the data referred by leave data master.", MsgBoxStyle.Information, "Information")
                Return False
            End If
            'tambahan dari master workshift allocation
            If Me.CekMasterWorkshiftAllocate(vPrimary) > 0 Then
                MsgBox("Employee data can not be deleted because the data referred by workshift data master.", MsgBoxStyle.Information, "Information")
                Return False
            End If
            'tambahan dari leave app setting
            Dummy = dta.ScalarQuery_GetJmlByPrimaryKaryawan(vPrimary)
            If Dummy.HasValue = True Then
                If Dummy.Value > 0 Then
                    MsgBox("Employee data can not be deleted because the data referred by leave application data setting.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If
            'salary formula
            Dummy = dta1.ScalarQuery_GetJmlByPrimaryKaryawan(vPrimary)
            If Dummy.HasValue = True Then
                If Dummy.Value > 0 Then
                    MsgBox("Employee data can not be deleted because the data referred by salary formula data.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'tidak boleh hapus, jika dirujuk oleh work unit
            Dummy = dta2.ScalarQuery_GetJmlWorkUnit(vPrimary)
            If Dummy.HasValue = True Then
                If Dummy.Value > 0 Then
                    MsgBox("Employee data can not be deleted because the data referred by work unit data master.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'strict edit
            'DummyUser = dta.ScalarQuery_GetCreated(vPrimary)
            'If DummyUser <> vUser Then
            '    MsgBox("Only " + DummyUser + " user is allowed to delete this transaction.", _
            '    MsgBoxStyle.Information, "Information")
            '    Return False
            'End If

            If Me.ApakahSudahDigunakanTransaksiPreviousIncome(Me.mPrimary) = True Then
                Return False
            End If

            'benar
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Erorr")
        Finally
            dta.Dispose()
            dta1.Dispose()
            dta2.Dispose()
        End Try
    End Function

    Public Function ApakahSudahDigunakanTransaksiPreviousIncome(ByVal vPrimaryEmployee As Double) As Boolean
        Dim dta As New DseMasterKaryawanTableAdapters.tbltrans_salary_prev_incomeTableAdapter
        Dim tbl As New DseMasterKaryawan.tbltrans_salary_prev_incomeDataTable
        Dim row As DseMasterKaryawan.tbltrans_salary_prev_incomeRow
        Dim A1_Form_No As String = ""

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dta.Fill(tbl, vPrimaryEmployee)
            For Each row In tbl
                If A1_Form_No <> "" Then
                    A1_Form_No &= ", "
                End If
                A1_Form_No &= row.a1_form_no
            Next

            If A1_Form_No <> "" Then
                MsgBox("Employee data can not be deleted because the data referred by transaction previous income (" & A1_Form_No & ") .", MsgBoxStyle.Information, "Information")
                Return True
            End If

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function ApakahSudahAda(ByVal vCode As String) As Boolean
        Dim dta As New DseMasterKaryawanTableAdapters.tblmaster_karyawanTableAdapter
        Dim dummy As Nullable(Of Long)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dummy = dta.ScalarQuery_GetJmlPrimaryByCode(vCode)
            If dummy.HasValue Then
                If dummy.Value > 0 Then
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
            dta.Dispose()
        End Try
    End Function

    Public Function ApakahAktif(ByVal vPrimary As Double) As Boolean
        Dim dta As New DseMasterKaryawanTableAdapters.tblmaster_karyawanTableAdapter
        Dim dummy As Nullable(Of Integer)

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dummy = dta.ScalarQuery_GetAktif(vPrimary)
            If dummy.HasValue Then
                If dummy.Value > 0 Then
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
            dta.Dispose()
        End Try
    End Function

    Public Function ApakahSudahAda_ShowSalary(ByVal vCode As String, ByVal UserCode As String) As Boolean
        Dim dta As New DseMasterKaryawanTableAdapters.tblmaster_karyawan1TableAdapter
        Dim dummy As Nullable(Of Long)
        Dim PrimaryUser As Double = -1
        Dim ObjUser As New ClsUser
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            PrimaryUser = ObjUser.GetPrimary(UserCode)
            dummy = dta.ScalarQuery_GetJmlByEmployeeCodeSalary(vCode, PrimaryUser)
            If dummy.HasValue Then
                If dummy.Value > 0 Then
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
            dta.Dispose()
            ObjUser = Nothing
        End Try
    End Function

    Public Function GetKompetensi(ByVal vPrimaryKaryawan As Double, ByVal vPrimaryKompetensi As Double) As String
        Dim dta As New DseMasterKaryawanTableAdapters.tblmaster_karyawan_item_kompetensiTableAdapter
        Dim Dummy As String
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = dta.ScalarQuery_GetValue(vPrimaryKaryawan, vPrimaryKompetensi)
            If String.IsNullOrEmpty(Dummy) = False Then
                Return Dummy
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

    Public Function GetPositionCode(ByVal vPrimary As Double) As String
        Dim DTA As New DseMasterKaryawanTableAdapters.DataTable5TableAdapter
        Dim Dummy As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetPositionCode(vPrimary)
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

    Public Function GetCodeByPrimary(ByVal vPrimary As Double) As String
        Dim dta As New DseMasterKaryawanTableAdapters.tblmaster_karyawanTableAdapter
        Dim Dummy As String
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = dta.ScalarQuery_GetCodeByPrimary(vPrimary)
            If String.IsNullOrEmpty(Dummy) = False Then
                Return Dummy
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

    Public Function GetJobByPrimary(ByVal vPrimary As Double) As Double
        Dim dta As New DseMasterKaryawanTableAdapters.tblmaster_karyawanTableAdapter
        Dim Dummy As Nullable(Of Double)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = dta.ScalarQuery_GetPrimaryJobByPrimary(vPrimary)
            If Dummy.HasValue Then
                Return Dummy
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetFullNameByPrimary(ByVal vPrimary As Double) As String
        Dim dta As New DseMasterKaryawanTableAdapters.tblmaster_karyawanTableAdapter
        Dim Dummy As String
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = dta.ScalarQuery_GetFullNameByPrimary(vPrimary)
            If String.IsNullOrEmpty(Dummy) = False Then
                Return Dummy
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

    'mengembalika jumlah karyawan yang menggunakan NIK selain kary dengan primary yang diberikan
    Public Function GetJmlKaryByNIK(ByVal vPrimaryKary As Double, ByVal NIKAcuan As String) As String
        Dim DTA As New DseMasterKaryawanTableAdapters.tblmaster_karyawanTableAdapter
        Dim Dummy As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetJml(vPrimaryKary, NIKAcuan)
            If Dummy.HasValue = False Then
                Return 0
            Else
                Return Dummy.Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetNIK(ByVal vPrimary As Double) As String
        Dim DTA As New DseMasterKaryawanTableAdapters.tblmaster_karyawanTableAdapter
        Dim Dummy As String
        Try
            Dummy = DTA.ScalarQuery_GetNIK(vPrimary)
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

    Public Function GetPrimaryByCode(ByVal vCode As String) As Double
        Dim dta As New DseMasterKaryawanTableAdapters.tblmaster_karyawanTableAdapter
        Dim Dummy As Nullable(Of Double)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = dta.ScalarQuery_GetPrimaryByCode(vCode)
            If Dummy.HasValue = True Then
                Return Dummy.Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Overridable Function SimpanBaru(ByVal vUserCode As String, ByVal vDGVFamily As DataGridView, _
  ByVal vDGVKompetensi As DataGridView, ByVal frm As FrmMasterKaryawan) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            DTAAntrian.Connection.ConnectionString = Core_Module.GetConstring
            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, ClsAntrian.EnumTipeAntrian.AntrianMaster, _
            Me.vPrimary_Antrian, Me.mPrimary, ClsAntrian.EnumTipeTransaksiAsalAntrian.MasterKaryawan_SimpanBaru, vUserCode)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= Me.ObjAntrian.GetWaitTime Then
                        Me.DTAAntrian.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If Me.DTAAntrian.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanBaru1(vUserCode, vDGVFamily, vDGVKompetensi, frm)
                            'hapus data antri untuk transaksi ini
                            If Me.DTAAntrian.Delete(Me.vPrimary_Antrian) = 0 Then
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

    Public Function SimpanBaru1(ByVal vUserCode As String, ByVal vDGVfamily As DataGridView, ByVal vDGVKompetensi As DataGridView, _
    ByVal frm As FrmMasterKaryawan) As Boolean
        Dim DTA As New DseMasterKaryawan1TableAdapters.tblmaster_karyawanTableAdapter
        Dim DTA2 As New DseMasterKaryawanTableAdapters.tblmaster_karyawanTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA2.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.Insert(Me.mCode, Me.mFullName, Me.mPrefix, Me.FristName, Me.mMiddleName, Me.mLastName, Me.mSuffixName, Me.mNIK, _
            Me.mRemark, Me.mAktif, vUserCode, Now, vUserCode, Now, Now, Me.mPrimaryDep, Me.mPicturePacth, Me.mAdmisionDate, Me.mPrimaryPosition, _
            Me.mKPA_Astek_No, Me.mKPJ_no, Me.mId_Card_no, Me.mId_Card_Exp, Me.mId_Card_Add, Me.mNPWP_No, Me.mNPWP_Category, Me.mNPWP_Add, _
            Me.mMartial_Status, Me.mMariageDAte, Me.mEducation, Me.mSex, Me.mCurentAdd, Me.mReligion, Me.mPhoneNo, Me.mEmployStatus, _
            Me.mBrithPalce, Me.mDateBrith, Me.mBloodType, Me.mNationality, Me.mEmail, Me.mMotherName, Me.mMotherAddres, Me.mBankName, _
            Me.mBankBranch, Me.mBankAccountNo, Me.mBankAccountHolder, Me.mBalaiPengobatanJPK, Me.mFullName, Me.mIsNPWP, Me.mForeignEmployee, _
            Me.mPrimaryCountry, Me.mAppointmentDate, Me.FingerprintID, Me.OverrideClockIn, Me.DoesNotClockIn, Me.OverrideClockOut, Me.DoesNotClockOut, _
            Me.OverrideEarlyOT, Me.IsCalculateEarlyOT, Me.OverrideLate, Me.IsCalculateLate, Me.OverrideEarlyOut, Me.IsCalculateEarlyOut, Me.OverrideOT, _
            Me.IsCalculateOT, Me.mCalculateAfterEarlyOT, Me.mCalculateAfterLate, Me.mCalculateAfterEarlyOut, Me.mCalculateAfterOT, _
            Me.mRoundingTypeEarlyOT, Me.mRoundingTypeLate, Me.mRoundingTypeEarlyOut, Me.mRoundingTypeOT, Me.mRoundingValueEarlyOT, Me.mRoundingValueLate, _
            Me.mRoundingValueEarlyOut, Me.mRoundingValueOT, Me.mMaxValueEarlyOT, Me.mMaxValueLate, Me.mMaxValueEarlyOut, Me.mMaxValueOT, Me.mBankCode, Me.mResignDate, _
            Me.mOverrideCalculateEarlyOTTolerance, Me.mOverrideCalculateLateTolerance, Me.mOverrideCalculateEarlyOutTolerance, Me.mOverrideCalculateOTTolerance, _
            Me.mCalculateEarlyOTOnWeekend, Me.mCalculateLateOnWeekend, Me.mCalculateEarlyOutOnWeekend, Me.mCalculateOTOnWeekend, _
            Me.mCalculateEarlyOTOnHoliday, Me.mCalculateLateOnHoliday, Me.mCalculateEarlyOutOnHoliday, Me.mCalculateOTOnHoliday, Me.mFirstTaxPeriod, Me.mUseOtDuration, Me.mNotificationEmailOt, Me.mOvverideOTOrder, Me.mLosesTaxObligations) > 0 Then
                Me.mPrimary = DTA2.ScalarQuery_GetPrimaryByCode(Me.mCode)
                Me.mTimestamp = DTA2.ScalarQuery_GetTimestampByPrimary(Me.mPrimary)
                Me.mCreatedDate = Me.mTimestamp
                Me.mModifiedDate = Me.mTimestamp
                Me.mCreatedBy = vUserCode
                Me.mModifiedBy = vUserCode
                'DTA.UpdateQuery_UpdateName(Me.mFullName, Me.mPrimary)

                'simpan item
                Me.objitem.SimpanFamily(vDGVfamily, Me.mPrimary, frm)
                Me.objitem.SimpanKompetensi(vDGVKompetensi, Me.mPrimary, frm)
                Me.objitem.SimpanKontrak(frm.DGVKontrak, Me.mPrimary, frm)
                Me.objitem.SimpanAttchKontrak(frm.DGVAttanchmentKontrak, Me.mPrimary, frm)
                Me.objitem.SimpanAttachment(frm.DGVAttanchment, Me.mPrimary, frm)

                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA2.Dispose()
        End Try
    End Function

    Public Overridable Function SimpanPerubahan(ByVal vUserCode As String, ByVal vDGVfamily As DataGridView, ByVal vDGVKompetensi As DataGridView, _
    ByVal frm As FrmMasterKaryawan) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            DTAAntrian.Connection.ConnectionString = Core_Module.GetConstring
            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, ClsAntrian.EnumTipeAntrian.AntrianMaster, _
            Me.vPrimary_Antrian, Me.mPrimary, ClsAntrian.EnumTipeTransaksiAsalAntrian.MasterKaryawan_SimpanPerubahan, vUserCode)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= Me.ObjAntrian.GetWaitTime Then
                        Me.DTAAntrian.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If Me.DTAAntrian.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanPerubahan1(vUserCode, vDGVfamily, vDGVKompetensi, frm)
                            'hapus data antri untuk transaksi ini
                            If Me.DTAAntrian.Delete(Me.vPrimary_Antrian) = 0 Then
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

    Public Function SimpanPerubahan1(ByVal vUser As String, ByVal vDGVfamily As DataGridView, ByVal vDGVKompetensi As DataGridView, _
    ByVal frm As FrmMasterKaryawan) As Boolean
        Dim DTA As New DseMasterKaryawan1TableAdapters.tblmaster_karyawanTableAdapter
        Dim DTA2 As New DseMasterKaryawanTableAdapters.tblmaster_karyawanTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA2.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahBolehEdit(Me.mPrimary, vUser) = True Then
                If DTA.UpdateQuery(Me.mCode, Me.mFullName, Me.mPrefix, Me.FristName, Me.mMiddleName, Me.mLastName, Me.mSuffixName, Me.mNIK, _
                Me.mRemark, Me.mAktif, vUser, Now, Now, Me.mPrimaryDep, Me.mPicturePacth, Me.mAdmisionDate, Me.mPrimaryPosition, _
                Me.mKPA_Astek_No, Me.mKPJ_no, Me.mId_Card_no, Me.mId_Card_Exp, Me.mId_Card_Add, Me.mNPWP_No, Me.mNPWP_Category, Me.mNPWP_Add, _
                Me.mMartial_Status, Me.mMariageDAte, Me.mEducation, Me.mSex, Me.mCurentAdd, Me.mReligion, Me.mPhoneNo, Me.mEmployStatus, _
                Me.mBrithPalce, Me.mDateBrith, Me.mBloodType, Me.mNationality, Me.mEmail, Me.mMotherName, Me.mMotherAddres, Me.mBankName, _
                Me.mBankBranch, Me.mBankAccountNo, Me.mBankAccountHolder, Me.mBalaiPengobatanJPK, Me.mIsNPWP, Me.mForeignEmployee, Me.mPrimaryCountry, _
                Me.mAppointmentDate, Me.mFullName, Me.mFingerprintID, Me.mOverrideClockIn, Me.mDoesNotClockIn, Me.mOverrideClockOut, Me.mDoesNotClockOut, _
                Me.OverrideEarlyOT, Me.IsCalculateEarlyOT, Me.OverrideLate, Me.IsCalculateLate, Me.OverrideEarlyOut, Me.IsCalculateEarlyOut, _
                Me.OverrideOT, Me.IsCalculateOT, Me.mBankCode, Me.mCalculateAfterEarlyOT, Me.mCalculateAfterLate, Me.mCalculateAfterEarlyOut, _
                Me.mCalculateAfterOT, Me.mRoundingTypeEarlyOT, Me.mRoundingTypeLate, Me.mRoundingTypeEarlyOut, Me.mRoundingTypeOT, Me.mRoundingValueEarlyOT, _
                Me.mRoundingValueLate, Me.mRoundingValueEarlyOut, Me.mRoundingValueOT, Me.mMaxValueEarlyOT, Me.mMaxValueLate, Me.mMaxValueEarlyOut, _
                Me.mMaxValueOT, Me.mResignDate, _
                Me.mOverrideCalculateEarlyOTTolerance, Me.mOverrideCalculateLateTolerance, Me.mOverrideCalculateEarlyOutTolerance, Me.mOverrideCalculateOTTolerance, _
                Me.mCalculateEarlyOTOnWeekend, Me.mCalculateLateOnWeekend, Me.mCalculateEarlyOutOnWeekend, Me.mCalculateOTOnWeekend, _
                Me.mCalculateEarlyOTOnHoliday, Me.mCalculateLateOnHoliday, Me.mCalculateEarlyOutOnHoliday, Me.mCalculateOTOnHoliday, _
                Me.mFirstTaxPeriod, Me.mUseOtDuration, Me.mNotificationEmailOt, Me.mOvverideOTOrder, Me.mLosesTaxObligations, Me.mPrimary, Me.mTimestamp) > 0 Then
                    Me.mTimestamp = DTA2.ScalarQuery_GetTimestampByPrimary(Me.Primary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUser

                    'hapus item
                    Me.objitem.HapusItem(Me.mPrimary)

                    'simpan item
                    Me.objitem.SimpanFamily(vDGVfamily, Me.mPrimary, frm)
                    Me.objitem.SimpanKompetensi(vDGVKompetensi, Me.mPrimary, frm)
                    Me.objitem.SimpanKontrak(frm.DGVKontrak, Me.mPrimary, frm)
                    Me.objitem.SimpanAttchKontrak(frm.DGVAttanchmentKontrak, Me.mPrimary, frm)
                    Me.objitem.SimpanAttachment(frm.DGVAttanchment, Me.mPrimary, frm)
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
            DTA2.Dispose()
        End Try
    End Function

    Public Function Hapus(ByVal vPrimaryTransaksi As Double, ByVal vUser As String) As Boolean

        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            DTAAntrian.Connection.ConnectionString = Core_Module.GetConstring
            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, ClsAntrian.EnumTipeAntrian.AntrianMaster, _
            Me.vPrimary_Antrian, Me.mPrimary, ClsAntrian.EnumTipeTransaksiAsalAntrian.MasterKaryawan_Hapus, vUser)
            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= Me.ObjAntrian.GetWaitTime Then
                        Me.DTAAntrian.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If Me.DTAAntrian.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.Hapus1(vPrimaryTransaksi, vUser)
                            'hapus data antri untuk transaksi ini
                            If Me.DTAAntrian.Delete(Me.vPrimary_Antrian) = 0 Then
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

    Public Function Hapus1(ByVal vPrimary As Double, ByVal vUser As String) As Boolean
        Dim DTA As New DseMasterKaryawanTableAdapters.tblmaster_karyawanTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahBolehHapus(vPrimary, vUser) = True Then
                DTA.Delete1(vPrimary)
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

    Public Function GetFirstTaxPeriod(ByVal vPrimaryKaryawan As Double) As Date
        Dim DTA As New DseMasterKaryawanTableAdapters.tblmaster_karyawanTableAdapter
        Dim RetVal As Nullable(Of Date)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetFirstTaxPeriod(vPrimaryKaryawan)

            If RetVal.HasValue = True Then
                Return RetVal.Value
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

    Public Function GetFirstSalaryEndPeriod(ByVal vPrimaryKaryawan As Double) As Date
        Dim DTA As New DseMasterKaryawanTableAdapters.tblmaster_karyawanTableAdapter
        Dim RetVal As Nullable(Of Date)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetEndDateMinimum(vPrimaryKaryawan)

            If RetVal.HasValue = True Then
                Return RetVal.Value
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

    Public Function GetLasSalaryEndPeriod(ByVal vPrimaryKaryawan As Double) As Date
        Dim DTA As New DseMasterKaryawanTableAdapters.tblmaster_karyawanTableAdapter
        Dim RetVal As Nullable(Of Date)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetEndDateMaximum(vPrimaryKaryawan)

            If RetVal.HasValue = True Then
                Return RetVal.Value
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

    Public Function CariDetail(ByVal vPrimary As Double, ByVal vFrm As FrmMasterKaryawan, _
    ByVal vDgvFamily As DataGridView, ByVal vDgvKompetensi As DataGridView) As Boolean
        Dim DTA As New DseMasterKaryawanTableAdapters.tblmaster_karyawanTableAdapter
        Dim TBL As New DseMasterKaryawan.tblmaster_karyawanDataTable
        Dim row As DseMasterKaryawan.tblmaster_karyawanRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_Primary(TBL, vPrimary)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = .Primary
                    Me.mCode = .Code
                    Me.mFullName = .Full_Name
                    Me.mPrefix = .Prefix
                    Me.mFristName = .First_Name
                    Me.mMiddleName = .Middle_Name
                    Me.mLastName = .Last_Name
                    Me.mSuffixName = .Suffix
                    Me.mNIK = .NIK
                    Me.mRemark = .Remarks
                    Me.mAktif = .Aktif
                    Me.mResignDate = .Resign_Date
                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
                    Me.mPrimaryDep = .primary_departemen
                    Me.mPicturePacth = .picture_path
                    Me.mAdmisionDate = .admision_date
                    Me.mPrimaryPosition = .primary_position
                    Me.mKPA_Astek_No = .kpa_astek_no
                    Me.mKPJ_no = .kpj_no
                    Me.mId_Card_no = .id_card_no
                    Me.mId_Card_Exp = .id_card_expiry
                    Me.mId_Card_Add = .id_card_address
                    Me.mIsNPWP = .Is_NPWP
                    Me.mNPWP_No = .npwp_no
                    Me.mNPWP_Category = .npwp_category
                    Me.mNPWP_Add = .npwp_address
                    Me.mMartial_Status = .marital_status
                    Me.mMariageDAte = .mariage_date
                    Me.mEducation = .education
                    Me.mSex = .sex
                    Me.mCurentAdd = .current_address
                    Me.mReligion = .religion
                    Me.mPhoneNo = .phone_no
                    Me.mEmployStatus = .Employee_Status
                    Me.mBrithPalce = .Birth_Place
                    Me.mDateBrith = .Date_Of_Birth
                    Me.mBloodType = .Blood_Type
                    Me.mNationality = .Nationality
                    Me.mEmail = .email
                    Me.mMotherName = .mother_name
                    Me.mMotherAddres = .mother_address
                    Me.mBankCode = .Bank_Code 'aul tambahan
                    Me.mBankName = .Bank_Name
                    Me.mBankBranch = .Bank_Branch
                    Me.mBankAccountNo = .Bank_Account_No
                    Me.mBankAccountHolder = .Bank_Account_Holder
                    Me.mBalaiPengobatanJPK = .Balai_Pengobatan_JPK
                    Me.mForeignEmployee = .Foreign_Employee
                    Me.mPrimaryCountry = .Primary_Country
                    Me.mLosesTaxObligations = .loses_subjective_tax_obligations

                    Me.objitem.IsiDGVItemFamily(vDgvFamily, .Primary, vFrm)
                    Me.objitem.IsiDGVItemKompetensi(vDgvKompetensi, .Primary, vFrm, .primary_position)
                    Me.objitem.IsiDGVItemKontrak(vFrm.DGVKontrak, .Primary, vFrm)
                    Me.objitem.IsiDGVItemAttchKontrak(vFrm.DGVAttanchmentKontrak, .Primary, vFrm)
                    Me.objitem.IsiDGVItemAttachment(vFrm.DGVAttanchment, .Primary, vFrm)

                    Me.mSalaryFormula = Me.GetSalaryDesc()
                    Me.mAppointmentDate = .appointment_date

                    Me.mFingerprintID = .fingerprint_id
                    Me.mOverrideClockIn = .override_clock_in
                    Me.mDoesNotClockIn = .does_not_clock_in
                    Me.mOverrideClockOut = .override_clock_out
                    Me.mDoesNotClockOut = .does_not_clock_out

                    Me.OverrideEarlyOT = .override_early_ot
                    Me.IsCalculateEarlyOT = .is_early_ot
                    Me.OverrideLate = .override_late
                    Me.IsCalculateLate = .is_late
                    Me.OverrideEarlyOut = .override_early_out
                    Me.IsCalculateEarlyOut = .is_early_out
                    Me.OverrideOT = .override_ot
                    Me.IsCalculateOT = .is_ot

                    Me.mCalculateAfterEarlyOT = .calculate_after_early_OT
                    Me.mCalculateAfterLate = .calculate_after_late
                    Me.mCalculateAfterEarlyOut = .calculate_after_early_out
                    Me.mCalculateAfterOT = .calculate_after_ot
                    Me.mRoundingTypeEarlyOT = .rounding_type_early_ot
                    Me.mRoundingTypeLate = .rounding_type_late
                    Me.mRoundingTypeEarlyOut = .rounding_type_early_out
                    Me.mRoundingTypeOT = .rounding_type_ot
                    Me.mRoundingValueEarlyOT = .rounding_value_early_ot
                    Me.mRoundingValueLate = .rounding_value_late
                    Me.mRoundingValueEarlyOut = .rounding_value_early_out
                    Me.mRoundingValueOT = .rounding_value_ot
                    Me.mMaxValueEarlyOT = .max_value_early_ot
                    Me.mMaxValueLate = .max_value_late
                    Me.mMaxValueEarlyOut = .max_value_early_out
                    Me.mMaxValueOT = .max_value_ot

                    Me.mOverrideCalculateEarlyOTTolerance = .calculate_early_ot_tolerance
                    Me.mOverrideCalculateLateTolerance = .calculate_late_tolerance
                    Me.mOverrideCalculateEarlyOutTolerance = .calculate_early_out_tolerance
                    Me.mOverrideCalculateOTTolerance = .calculate_ot_tolerance

                    Me.mCalculateEarlyOTOnWeekend = .calculate_early_ot_on_weekend
                    Me.mCalculateLateOnWeekend = .calculate_late_on_weekend
                    Me.mCalculateEarlyOutOnWeekend = .calculate_early_out_on_weekend
                    Me.mCalculateOTOnWeekend = .calculate_ot_on_weekend
                    Me.mCalculateEarlyOTOnHoliday = .calculate_early_ot_on_holiday
                    Me.mCalculateLateOnHoliday = .calculate_late_on_holiday
                    Me.mCalculateEarlyOutOnHoliday = .calculate_early_out_on_holiday
                    Me.mCalculateOTOnHoliday = .calculate_ot_on_holiday

                    Me.mFirstTaxPeriod = .first_tax_period
                    Me.mUseOtDuration = .use_ot_order
                    Me.mNotificationEmailOt = .notification_email_ot
                    Me.mOvverideOTOrder = .override_ot_order
                End With
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function GetTimeStampByPrimary(ByVal vPrimary As Double) As Date
        Dim dta As New DseMasterKaryawanTableAdapters.tblmaster_karyawanTableAdapter
        Dim Dummy As String
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = dta.ScalarQuery_GetTimeStampByPrimary(vPrimary)
            If String.IsNullOrEmpty(Dummy) = False Then
                Return Dummy
            Else
                Return "1/1/1950"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetJml_Salary(ByVal vKriteria As String, ByVal vSearchBy As _
   EnumSearchCriteriaEmployee, ByVal vUserCode As String) As Double
        Dim DTA As New DseMasterKaryawanTableAdapters.tblmaster_karyawan1TableAdapter
        Dim ObjUser As New Core_Module.ClsUser
        Dim PrimaryUser As Double = -1
        Dim Dummy As Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            PrimaryUser = ObjUser.GetPrimary(vUserCode)
            If vSearchBy = EnumSearchCriteriaEmployee.ByCode Then
                Dummy = DTA.ScalarQuery_GetJmlPrimaryByCodeLIKE(vKriteria + "%", PrimaryUser)
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteriaEmployee.ByName Then
                Dummy = DTA.ScalarQuery_GetJmlPrimaryByFullNameLIKE(vKriteria + "%", PrimaryUser)
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteriaEmployee.ByDepartement Then
                Dummy = DTA.ScalarQuery_GetJmlPrimaryByDeptCodeLIKE(vKriteria + "%", PrimaryUser)
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteriaEmployee.ByPosition Then
                Dummy = DTA.ScalarQuery_GetJmlPrimaryByPositionCodeLIKE(vKriteria + "%", PrimaryUser)
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteriaEmployee.ByNote Then
                Dummy = DTA.ScalarQueryGetJmlPrimaryByNoteLIKE(vKriteria + "%", PrimaryUser)
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjUser = Nothing
        End Try
    End Function

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As _
    EnumSearchCriteriaEmployee, ByVal vStartDate As Date, ByVal vEndDate As Date) As Double
        Dim DTA As New DseMasterKaryawanTableAdapters.DataTable1TableAdapter
        Dim Dummy As Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If vSearchBy = EnumSearchCriteriaEmployee.ByCode Then
                Dummy = DTA.ScalarQuery_GetJmlByPrimaryByCodeLike(vKriteria + "%")
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteriaEmployee.ByName Then
                Dummy = DTA.ScalarQuery_GetJmlPrimaryByFullNameLIKE(vKriteria + "%")
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteriaEmployee.ByDepartement Then
                Dummy = DTA.ScalarQuery_GetJmlPrimaryByDeptCode(vKriteria + "%")
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteriaEmployee.ByPosition Then
                Dummy = DTA.ScalarQuery_GetJmlPrimaryByPosition(vKriteria + "%")
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteriaEmployee.ByNote Then
                Dummy = DTA.ScalarQuery_GetJmlPrimaryByNoteLike(vKriteria + "%")
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteriaEmployee.ByResignDate Then
                Dummy = DTA.ScalarQuery_GetJmlByResignDate(vEndDate, vStartDate)
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteriaEmployee.ByFirstAttendanceDate Then
                Dummy = DTA.ScalarQuery_GetJmlByMinAttandanceDate(vEndDate, vStartDate)
                If Dummy.HasValue = True Then
                    Return Dummy.Value
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

    Public Function IsiDGV_Salary(ByVal Criteria As String, ByVal UserCode As String, _
  ByVal DGV As DataGridView, ByVal TBL As DseMasterKaryawan.tblmaster_karyawan1DataTable, _
  ByVal SearchCriteria As EnumSearchCriteriaEmployee, ByVal tsl As ToolStripLabel) As Boolean
        Dim DTA As New DseMasterKaryawanTableAdapters.tblmaster_karyawan1TableAdapter
        Dim ObjUser As New Core_Module.ClsUser
        Dim PrimaryUser As Double = -1
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            PrimaryUser = ObjUser.GetPrimary(UserCode)
            If SearchCriteria = EnumSearchCriteriaEmployee.ByCode Then
                QueryResult = DTA.FillBy_CodeLIKE(TBL, Criteria + "%", PrimaryUser)
            ElseIf SearchCriteria = EnumSearchCriteriaEmployee.ByName Then
                QueryResult = DTA.FillBy_NamaLIKE(TBL, Criteria + "%", PrimaryUser)
            ElseIf SearchCriteria = EnumSearchCriteriaEmployee.ByDepartement Then
                QueryResult = DTA.FillBy_DepartementLIKE(TBL, Criteria + "%", PrimaryUser)
            ElseIf SearchCriteria = EnumSearchCriteriaEmployee.ByPosition Then
                QueryResult = DTA.FillBy_PositionLIKE(TBL, Criteria + "%", PrimaryUser)
            ElseIf SearchCriteria = EnumSearchCriteriaEmployee.ByNote Then
                QueryResult = DTA.FillBy_RemarksLIKE(TBL, Criteria + "%", PrimaryUser)
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Columns(0).Visible = False
            DGV.Refresh()

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

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Function IsiDGV(ByVal Criteria As String, _
   ByVal DGV As DataGridView, ByVal TBL As DseMasterKaryawan.DataTable1DataTable, _
   ByVal SearchCriteria As EnumSearchCriteriaEmployee, ByVal tsl As ToolStripLabel, ByVal vStartDate As Date, ByVal vEndDate As Date) As Boolean
        Dim DTA As New DseMasterKaryawanTableAdapters.DataTable1TableAdapter
        'Dim row As DseMasterKaryawan.DataTable1Row
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If SearchCriteria = EnumSearchCriteriaEmployee.ByCode Then
                QueryResult = DTA.FillBy_LikeCode(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteriaEmployee.ByName Then
                QueryResult = DTA.FillBy_LikeName(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteriaEmployee.ByDepartement Then
                QueryResult = DTA.FillBy_LikeDepartement(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteriaEmployee.ByPosition Then
                QueryResult = DTA.FillBy_LikePosition(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteriaEmployee.ByNote Then
                QueryResult = DTA.FillBy_LikeNote(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteriaEmployee.ByResignDate Then
                QueryResult = DTA.FillBy_ResignDate(TBL, vEndDate, vStartDate)
            ElseIf SearchCriteria = EnumSearchCriteriaEmployee.ByFirstAttendanceDate Then
                QueryResult = DTA.FillBy_FirstAttendanceDate(TBL, vEndDate, vStartDate)
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Columns(0).Visible = False
            DGV.Refresh()

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

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Function ApakahSesuaiDenganJobList(ByVal vPrimaryJob As Double, ByVal i As FrmMasterKaryawan) As Boolean
        Dim Dta As New DseMasterKaryawanTableAdapters.tblmaster_jobposition_item_kompetensiTableAdapter
        Dim Tbl As New DseMasterKaryawan.tblmaster_jobposition_item_kompetensiDataTable
        Dim row As DseMasterKaryawan.tblmaster_jobposition_item_kompetensiRow
        Dim RowDGV As DataGridViewRow
        Dim Ketemu As Boolean = False
        Dim ListString As New List(Of String)
        Dim Index As String
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            Dta.Fill(Tbl, vPrimaryJob)
            For Each row In Tbl.Rows
                Ketemu = False
                For Each RowDGV In i.DGVKompetensi.Rows
                    If RowDGV.Cells(i.ColCompetence_Primary.Index).Value = row.primary_kompetensi Then
                        Ketemu = True
                        Exit For
                    End If
                Next
                If Ketemu = False Then
                    ListString.Add("False")
                Else
                    ListString.Add("True")
                End If
            Next
            For Each Index In ListString
                If Index = "False" Then
                    Return False
                End If
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Dta.Dispose()
            Tbl.Dispose()
        End Try
    End Function

    Public Function GetListKompeten(ByVal vPrimaryJob As Double) As String
        Dim Dta As New DseMasterKaryawanTableAdapters.tblmaster_jobposition_item_kompetensiTableAdapter
        Dim Tbl As New DseMasterKaryawan.tblmaster_jobposition_item_kompetensiDataTable
        Dim row As DseMasterKaryawan.tblmaster_jobposition_item_kompetensiRow
        Dim RetVal As String = ""
        Dim i As Integer = 1
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            Dta.Fill(Tbl, vPrimaryJob)
            For Each row In Tbl.Rows
                RetVal = RetVal + i.ToString + ". " + row.Data_Code.ToString + vbNewLine
                i = i + 1
            Next
            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            Dta.Dispose()
            Tbl.Dispose()
        End Try
    End Function

    Private Function GetSalaryDesc() As String
        Dim Dta As New DseMasterKaryawanTableAdapters.DataTable4TableAdapter
        Dim tbl As New DseMasterKaryawan.DataTable4DataTable
        Dim row As DseMasterKaryawan.DataTable4Row
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            Dta.Fill(tbl, Me.Primary)
            If tbl.Rows.Count > 1 Then
                MsgBox("ada data salah")
                Return ""
            ElseIf tbl.Rows.Count = 1 Then
                For Each row In tbl.Rows
                    Return row.desc
                Next
            Else
                Return ""
            End If

            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            Dta.Dispose()
            tbl.Dispose()
        End Try
    End Function

    Public Sub CetakanDataKaryawan(ByVal vCRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal vPrimaryKaryawan As Double)
        Dim Report As New CrpMasterKaryawanCetak
        Dim ds As New DseMasterKaryawan
        Dim DtaHeader As New DseMasterKaryawanTableAdapters.tblmaster_karyawan_printTableAdapter
        Dim TblHeader As DseMasterKaryawan.tblmaster_karyawan_printDataTable = ds.tblmaster_karyawan_print
        Dim RowHeader As DseMasterKaryawan.tblmaster_karyawan_printRow
        Dim DtaItem As New DseMasterKaryawanTableAdapters.tblmaster_karyawan_itemTableAdapter
        Dim TblItem As New DseMasterKaryawan.tblmaster_karyawan_itemDataTable
        Dim RowItem As DseMasterKaryawan.tblmaster_karyawan_itemRow
        Dim TblCetakItem As DseMasterKaryawan.Tbl_cetak_item_karyawanDataTable = ds.Tbl_cetak_item_karyawan
        Dim RowCetakItem As DseMasterKaryawan.Tbl_cetak_item_karyawanRow

        Dim EmployStat As String = ""
        Dim gender As String = ""
        Dim StatPernikahan As String = ""
        Dim GolDarah As String = ""
        Dim Agama As String = ""
        Dim TglMenikah As String = ""

        Try
            DtaHeader.Connection.ConnectionString = Core_Module.GetConstring
            DtaItem.Connection.ConnectionString = DtaHeader.Connection.ConnectionString


            DtaHeader.FillBy(TblHeader, vPrimaryKaryawan)
            DtaItem.FillBy(TblItem, vPrimaryKaryawan)

            For Each RowHeader In TblHeader.Rows
                'status
                If RowHeader.Employee_Status = ClsMasterKaryawan.EnumEmployStatus.Contract Then
                    EmployStat = "Kontrak"
                ElseIf RowHeader.Employee_Status = ClsMasterKaryawan.EnumEmployStatus.Jobholder Then
                    EmployStat = "Tetap"
                ElseIf RowHeader.Employee_Status = ClsMasterKaryawan.EnumEmployStatus.Probation Then
                    EmployStat = "Percobaan"
                End If
                'gender
                If RowHeader.sex = ClsMasterKaryawan.EnumSex.Female Then
                    gender = "Perempuan"
                ElseIf RowHeader.sex = ClsMasterKaryawan.EnumSex.Male Then
                    gender = "Laki-laki"
                End If
                'Status pernikahan
                If RowHeader.marital_status = ClsMasterKaryawan.EnumMartialStatus.Married Then
                    StatPernikahan = "Menikah"
                    TglMenikah = Format(RowHeader.mariage_date, "dd MMMM yyyy").ToString
                ElseIf RowHeader.marital_status = ClsMasterKaryawan.EnumMartialStatus._Single Then
                    StatPernikahan = "Belum menikah"
                    TglMenikah = "-"
                End If
                'Gol darah
                If RowHeader.Blood_Type = ClsMasterKaryawan.EnumBloodType.A Then
                    GolDarah = "A"
                ElseIf RowHeader.Blood_Type = ClsMasterKaryawan.EnumBloodType.AB Then
                    GolDarah = "AB"
                ElseIf RowHeader.Blood_Type = ClsMasterKaryawan.EnumBloodType.B Then
                    GolDarah = "B"
                ElseIf RowHeader.Blood_Type = ClsMasterKaryawan.EnumBloodType.O Then
                    GolDarah = "O"
                End If
                'Agama
                If RowHeader.religion = ClsMasterKaryawan.EnumReligion.Islam Then
                    Agama = "Islam"
                ElseIf RowHeader.religion = ClsMasterKaryawan.EnumReligion.Buddha Then
                    Agama = "Budha"
                ElseIf RowHeader.religion = ClsMasterKaryawan.EnumReligion.Hindu Then
                    Agama = "Hindu"
                ElseIf RowHeader.religion = ClsMasterKaryawan.EnumReligion.Katolik Then
                    Agama = "Katolik"
                ElseIf RowHeader.religion = ClsMasterKaryawan.EnumReligion.Konghucu Then
                    Agama = "Konghucu"
                ElseIf RowHeader.religion = ClsMasterKaryawan.EnumReligion.Kristen Then
                    Agama = "Kristen"
                ElseIf RowHeader.religion = ClsMasterKaryawan.EnumReligion.Others Then
                    Agama = "Other"
                End If
                RowHeader.Str_Cat_NPWP = Me.GetNPWPCategoryStr(RowHeader.npwp_category)
            Next

            For Each RowItem In TblItem.Rows
                RowCetakItem = TblCetakItem.NewRow
                RowCetakItem.Nama = RowItem.nama
                If RowItem.relationship = ClsMasterKaryawan_Item.EnumHubKeluarga.Aunt Then
                    RowCetakItem.Hub_Keluarga = "Aunt"
                ElseIf RowItem.relationship = ClsMasterKaryawan_Item.EnumHubKeluarga.Brother Then
                    RowCetakItem.Hub_Keluarga = "Brother"
                ElseIf RowItem.relationship = ClsMasterKaryawan_Item.EnumHubKeluarga.Daughter Then
                    RowCetakItem.Hub_Keluarga = "Daughter"
                ElseIf RowItem.relationship = ClsMasterKaryawan_Item.EnumHubKeluarga.Father Then
                    RowCetakItem.Hub_Keluarga = "Father"
                ElseIf RowItem.relationship = ClsMasterKaryawan_Item.EnumHubKeluarga.Husband Then
                    RowCetakItem.Hub_Keluarga = "Husband"
                ElseIf RowItem.relationship = ClsMasterKaryawan_Item.EnumHubKeluarga.Mother Then
                    RowCetakItem.Hub_Keluarga = "Mother"
                ElseIf RowItem.relationship = ClsMasterKaryawan_Item.EnumHubKeluarga.Other Then
                    RowCetakItem.Hub_Keluarga = "Other"
                ElseIf RowItem.relationship = ClsMasterKaryawan_Item.EnumHubKeluarga.Sister Then
                    RowCetakItem.Hub_Keluarga = "Sister"
                ElseIf RowItem.relationship = ClsMasterKaryawan_Item.EnumHubKeluarga.Son Then
                    RowCetakItem.Hub_Keluarga = "Son"
                ElseIf RowItem.relationship = ClsMasterKaryawan_Item.EnumHubKeluarga.Uncle Then
                    RowCetakItem.Hub_Keluarga = "Uncle"
                ElseIf RowItem.relationship = ClsMasterKaryawan_Item.EnumHubKeluarga.Wife Then
                    RowCetakItem.Hub_Keluarga = "Wife"
                End If
                RowCetakItem.Pendidikan = RowItem.education
                If RowItem.sex = ClsMasterKaryawan_Item.EnumSEx.Female Then
                    RowCetakItem._L_P = "P"
                ElseIf RowItem.sex = ClsMasterKaryawan_Item.EnumSEx.Male Then
                    RowCetakItem._L_P = "L"
                End If
                RowCetakItem.Tempat_Lahir = RowItem.birth_place
                RowCetakItem.Tanggal_Lahir = RowItem.birthday
                RowCetakItem.Keterangan = RowItem.remarks
                TblCetakItem.AddTbl_cetak_item_karyawanRow(RowCetakItem)
            Next


            Report.SetDataSource(ds)
            Report.SetParameterValue(Report.Parameter_Status_Karyawan.ParameterFieldName, EmployStat)
            Report.SetParameterValue(Report.Parameter_Jenis_Kelamin.ParameterFieldName, gender)
            Report.SetParameterValue(Report.Parameter_Status_Pernikahan.ParameterFieldName, StatPernikahan)
            Report.SetParameterValue(Report.Parameter_Golongan_Darah.ParameterFieldName, GolDarah)
            Report.SetParameterValue(Report.Parameter_Agama.ParameterFieldName, Agama)
            Report.SetParameterValue(Report.Parameter_Tanggal_Menikah.ParameterFieldName, TglMenikah)
            vCRV.ReportSource = Report
            vCRV.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DtaHeader.Dispose()
            DtaItem.Dispose()
        End Try
    End Sub

    Public Function GetGolonganPajak(ByVal vPrimary As Double) As Integer
        Dim DTA As New DseMasterKaryawanTableAdapters.tblmaster_karyawan2TableAdapter
        Dim Dummy As System.Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetGolonganPajak(vPrimary)
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

    Public Function GetIsNPWP(ByVal vPrimary As Double) As ClsMasterKaryawan.EnumIsNPWP
        Dim DTA As New DseMasterKaryawanTableAdapters.tblmaster_karyawanTableAdapter
        Dim Dummy As System.Nullable(Of Integer)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetIsNPWP(vPrimary)
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

    Public Function GetJmlByPrimaryDept(ByVal vPrimaryDept As Double) As Double
        Dim DTA As New DseMasterKaryawanTableAdapters.tblmaster_karyawanTableAdapter
        Dim Dummy As System.Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetJmlByPrimaryDept(vPrimaryDept)
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

    Public Function GetNoNPWP(ByVal vPrimary As Double) As String
        Dim DTA As New DseMasterKaryawanTableAdapters.tblmaster_karyawanTableAdapter
        Dim Dummy As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetNoNPWP(vPrimary)
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

    Public Function GetNPWPAddress(ByVal vPrimary As Double) As String
        Dim DTA As New DseMasterKaryawanTableAdapters.tblmaster_karyawanTableAdapter
        Dim Dummy As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetNPWPAddress(vPrimary)
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

    Public Function GetNPWPCategory(ByVal vPrimary As Double) As Integer
        Dim DTA As New DseMasterKaryawanTableAdapters.DataTable5TableAdapter
        Dim Dummy As System.Nullable(Of Integer)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetNPWPCategory(vPrimary)
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

    Public Function GetDeptDescription(ByVal vPrimary As Double) As String
        Dim DTA As New DseMasterKaryawanTableAdapters.DataTable5TableAdapter
        Dim Dummy As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetDeptDesc(vPrimary)
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

    Public Function GetPositionDescription(ByVal vPrimary As Double) As String
        Dim DTA As New DseMasterKaryawanTableAdapters.DataTable5TableAdapter
        Dim Dummy As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetPositionDesc(vPrimary)
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

    Public Sub IsiTBLKodeNegara(ByVal CBO As ComboBox, ByVal TBL As DseMasterKaryawan.tblmaster_kode_negaraDataTable, _
    ByVal DTA As DseMasterKaryawanTableAdapters.tblmaster_kode_negaraTableAdapter, ByVal i As FrmMasterKaryawan)
        Dim row As DseMasterKaryawan.tblmaster_kode_negaraRow
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            CBO.DataSource = Nothing
            CBO.Items.Clear()

            DTA.Fill(TBL)

            CBO.DataSource = TBL.DefaultView
            row = TBL.DefaultView(0).Row
            TBL.DefaultView.Sort = "[Primary] ASC"
            CBO.DisplayMember = "Nama Negara"
            i.TxtPrimaryCountryCode.Text = row.Primary
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetSex(ByVal vPrimary As Double) As Integer
        Dim DTA As New DseMasterKaryawanTableAdapters.tblmaster_karyawanTableAdapter
        Dim Dummy As System.Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetSex(vPrimary)
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

    Public Function GetForeignEmployee(ByVal vPrimary As Double) As Integer
        Dim DTA As New DseMasterKaryawanTableAdapters.DataTable5TableAdapter
        Dim Dummy As System.Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetForeignEmployee(vPrimary)
            If Dummy.HasValue Then
                Return Dummy
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetKodeNegara(ByVal vPrimaryKodeNegara As Double) As String
        Dim DTA As New DseMasterKaryawanTableAdapters.tblmaster_kode_negaraTableAdapter
        Dim Dummy As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetKodeNegara(vPrimaryKodeNegara)
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

    Public Function GetPrimaryCountry(ByVal vPrimary As Double) As Double
        Dim DTA As New DseMasterKaryawanTableAdapters.DataTable5TableAdapter
        Dim Dummy As System.Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetPrimaryCountry(vPrimary)
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

    Public Function GetPrimaryDept(ByVal vPrimary As Double) As Double
        Dim DTA As New DseMasterKaryawanTableAdapters.DataTable5TableAdapter
        Dim Dummy As System.Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetPrimaryDept(vPrimary)
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

    Public Function GetPrimaryJobPosition(ByVal vPrimary As Double) As Double
        Dim DTA As New DseMasterKaryawanTableAdapters.DataTable5TableAdapter
        Dim Dummy As System.Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetPrimaryJobPosition(vPrimary)
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

    Public Function GetDeptByPrimary(ByVal vPrimary As Double) As Double
        Dim dta As New DseMasterKaryawanTableAdapters.tblmaster_karyawanTableAdapter
        Dim Dummy As Nullable(Of Double)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = dta.ScalarQuery_GetPrimaryDeptByPrimary(vPrimary)
            If Dummy.HasValue Then
                Return Dummy
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetMartialStatus(ByVal vPrimary As Double) As Double
        Dim DTA As New DseMasterKaryawanTableAdapters.tblmaster_karyawanTableAdapter
        Dim Dummy As System.Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetMartialStatus(vPrimary)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return ClsMasterKaryawan.EnumMartialStatus._Single
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetIsOverrideClockIn(ByVal vPrimary As Double) As Boolean
        Dim Dta As New DseMasterKaryawanTableAdapters.tblmaster_karyawanTableAdapter
        Dim IsOverrideClockIn_ As Nullable(Of Integer)
        Try
            IsOverrideClockIn_ = Dta.ScalarQuery_GetIsOverrideClockIn(vPrimary)
            If IsOverrideClockIn_.HasValue = True Then
                If IsOverrideClockIn_.Value = 0 Then
                    Return False
                Else
                    Return True
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Function GetIsOverrideClockOut(ByVal vPrimary As Double) As Boolean
        Dim Dta As New DseMasterKaryawanTableAdapters.tblmaster_karyawanTableAdapter
        Dim IsOverrideClockOut_ As Nullable(Of Integer)
        Try
            IsOverrideClockOut_ = Dta.ScalarQuery_GetIsOverrideClockOut(vPrimary)
            If IsOverrideClockOut_.HasValue = True Then
                If IsOverrideClockOut_.Value = 0 Then
                    Return False
                Else
                    Return True
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Function GetDoesNotClockIn(ByVal vPrimary As Double) As Boolean
        Dim Dta As New DseMasterKaryawanTableAdapters.tblmaster_karyawanTableAdapter
        Dim DoesNotClockIn_ As Nullable(Of Integer)
        Try
            DoesNotClockIn_ = Dta.ScalarQuery_GetDoesNotClockIn(vPrimary)
            If DoesNotClockIn_.HasValue = True Then
                If DoesNotClockIn_.Value = 0 Then
                    Return False
                Else
                    Return True
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Function GetDoesNotClockOut(ByVal vPrimary As Double) As Boolean
        Dim Dta As New DseMasterKaryawanTableAdapters.tblmaster_karyawanTableAdapter
        Dim DoesNotClockOut_ As Nullable(Of Integer)
        Try
            DoesNotClockOut_ = Dta.ScalarQuery_GetDoesNotClockOut(vPrimary)
            If DoesNotClockOut_.HasValue = True Then
                If DoesNotClockOut_.Value = 0 Then
                    Return False
                Else
                    Return True
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Function GetFingerprintID(ByVal vPrimary As Double) As String
        Dim Dta As New Core_Module.DseMasterKaryawanTableAdapters.tblmaster_karyawanTableAdapter
        Dim FingerPrintID_ As String = ""
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            FingerPrintID_ = Dta.ScalarQuery_GetFingerprintID(vPrimary)
            If String.IsNullOrEmpty(FingerPrintID_) = False Then
                Return FingerPrintID_
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Function IsFingerprintIDAlreadyExist(ByVal vFingerprintID As String, ByVal vPrimary As Double) As Boolean
        Dim Dta As New Core_Module.DseMasterKaryawanTableAdapters.tblmaster_karyawanTableAdapter
        Dim count_ As Nullable(Of Long)
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            count_ = Dta.ScalarQuery_GetCountByFingerprintIDAndPrimary(vFingerprintID, vPrimary)
            If count_.HasValue = True Then
                If count_.Value > 0 Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Function GetLengthOfService_Days_AfterAppointment(ByVal vPrimaryKary As Double, ByVal vRefDate As Date) As Double
        Dim vAppointmentDate As Nullable(Of Date)
        Dim DTA As New DseMasterKaryawanTableAdapters.tblmaster_karyawanTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            vAppointmentDate = DTA.ScalarQuery_GetAppointmentDate(vPrimaryKary)
            If vAppointmentDate.HasValue = True Then
                If vAppointmentDate.Value.Date > vRefDate.Date Then
                    Return 0
                Else
                    Return (vRefDate.Date - vAppointmentDate.Value.Date).Days
                End If
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetLengthOfService_Months_AfterAppointment(ByVal vPrimaryKary As Double, ByVal vRefDate As Date) As Decimal
        Dim vAppointmentDate As Nullable(Of Date)
        Dim DTA As New DseMasterKaryawanTableAdapters.tblmaster_karyawanTableAdapter
        Dim DummyRetval As Decimal
        Dim DummyRefDate As Date

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            vAppointmentDate = DTA.ScalarQuery_GetAppointmentDate(vPrimaryKary)
            If vAppointmentDate.HasValue = True Then
                DummyRetval = 0
                DummyRefDate = vAppointmentDate.Value
                If DummyRefDate.Date <= vRefDate.Date Then
                    While DummyRefDate.Date.AddMonths(1).Date - TimeSpan.FromDays(1) <= vRefDate.Date
                        DummyRefDate = DummyRefDate.AddMonths(1)
                        DummyRetval += 1
                    End While
                    'hitung sisanya: dibagi 31 hari, kalo di bagi 30, masalah kalo nanti selisih hari-nya 30 hari
                    If DummyRefDate.Date < vRefDate.Date Then
                        DummyRetval += Math.Round((vRefDate.Date - DummyRefDate.Date).Days / 31, 2, MidpointRounding.AwayFromZero)
                    End If

                    Return DummyRetval
                Else
                    Return 0
                End If
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetLengthOfService_Years_AfterAppointment(ByVal vPrimaryKary As Double, ByVal vRefDate As Date) As Decimal
        Dim vAppointmentDate As Nullable(Of Date)
        Dim DTA As New DseMasterKaryawanTableAdapters.tblmaster_karyawanTableAdapter
        Dim DummyRetval As Decimal
        Dim DummyRefDate As Date

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            vAppointmentDate = DTA.ScalarQuery_GetAppointmentDate(vPrimaryKary)
            If vAppointmentDate.HasValue = True Then
                DummyRetval = 0
                DummyRefDate = vAppointmentDate.Value
                If DummyRefDate.Date <= vRefDate.Date Then
                    DummyRetval = 0
                    While DummyRefDate.Date.AddYears(1).Date - TimeSpan.FromDays(1) <= vRefDate.Date
                        DummyRefDate = DummyRefDate.AddYears(1)
                        DummyRetval += 1
                    End While
                    'hitung sisanya: dibagi 366 hari, kalo dibagi 365, nanti masalah kalo tahunnya kabisat
                    'rounding value 3, karena kalo 2, durasi 1 hari ga akan keliatan dalam decimal
                    If DummyRefDate.Date < vRefDate.Date Then
                        DummyRetval += Math.Round((vRefDate.Date - DummyRefDate.Date).Days / 366, 3, MidpointRounding.AwayFromZero)
                    End If

                    Return DummyRetval
                Else
                    Return 0
                End If
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetLengthOfService_Days_Total(ByVal vPrimaryKary As Double, ByVal vRefDate As Date) As Double
        Dim vAdmisionDate As Nullable(Of Date)
        Dim DTA As New DseMasterKaryawanTableAdapters.tblmaster_karyawanTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            vAdmisionDate = DTA.ScalarQuery_GetAdmissionDate(vPrimaryKary)
            If vAdmisionDate.HasValue = True Then
                If vAdmisionDate.Value.Date > vRefDate.Date Then
                    Return 0
                Else
                    Return (vRefDate.Date - vAdmisionDate.Value.Date).Days
                End If
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetLengthOfService_Months_Total(ByVal vPrimaryKary As Double, ByVal vRefDate As Date) As Decimal
        Dim vAdmisionDate As Nullable(Of Date)
        Dim DTA As New DseMasterKaryawanTableAdapters.tblmaster_karyawanTableAdapter
        Dim DummyRetval As Decimal
        Dim DummyRefDate As Date

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            vAdmisionDate = DTA.ScalarQuery_GetAdmissionDate(vPrimaryKary)
            If vAdmisionDate.HasValue = True Then
                DummyRetval = 0
                DummyRefDate = vAdmisionDate.Value
                If DummyRefDate.Date <= vRefDate.Date Then
                    While DummyRefDate.Date.AddMonths(1).Date - TimeSpan.FromDays(1) <= vRefDate.Date
                        DummyRefDate = DummyRefDate.AddMonths(1)
                        DummyRetval += 1
                    End While
                    'hitung sisanya: dibagi 31 hari, kalo di bagi 30, masalah kalo nanti selisih hari-nya 30 hari
                    If DummyRefDate.Date < vRefDate.Date Then
                        DummyRetval += Math.Round((vRefDate.Date - DummyRefDate.Date).Days / 31, 2, MidpointRounding.AwayFromZero)
                    End If

                    Return DummyRetval
                Else
                    Return 0
                End If
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetResignDate(ByVal vPrimaryKary As Double) As Date
        Dim vResignDate As Nullable(Of Date)
        Dim DTA As New DseMasterKaryawanTableAdapters.tblmaster_karyawanTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            vResignDate = DTA.ScalarQuery_GetResignDate(vPrimaryKary)
            If vResignDate.HasValue = True Then
                Return vResignDate
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return Today
        Finally
            DTA.Dispose()
        End Try
    End Function 

    Public Function GetAdmissionDate(ByVal vPrimaryKary As Double) As Date
        Dim vAdmisionDate As Nullable(Of Date)
        Dim DTA As New DseMasterKaryawanTableAdapters.tblmaster_karyawanTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            vAdmisionDate = DTA.ScalarQuery_GetAdmissionDate(vPrimaryKary)
            If vAdmisionDate.HasValue = True Then
                Return vAdmisionDate
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

    Public Function GetLengthOfService_Years_Total(ByVal vPrimaryKary As Double, ByVal vRefDate As Date) As Decimal
        Dim vAdmisionDate As Nullable(Of Date)
        Dim DTA As New DseMasterKaryawanTableAdapters.tblmaster_karyawanTableAdapter
        Dim DummyRetval As Decimal
        Dim DummyRefDate As Date

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            vAdmisionDate = DTA.ScalarQuery_GetAdmissionDate(vPrimaryKary)
            If vAdmisionDate.HasValue = True Then
                DummyRetval = 0
                DummyRefDate = vAdmisionDate.Value
                If DummyRefDate.Date <= vRefDate.Date Then
                    DummyRetval = 0
                    While DummyRefDate.Date.AddYears(1).Date - TimeSpan.FromDays(1) <= vRefDate.Date
                        DummyRefDate = DummyRefDate.AddYears(1)
                        DummyRetval += 1
                    End While
                    'hitung sisanya: dibagi 366 hari, kalo dibagi 365, nanti masalah kalo tahunnya kabisat
                    'rounding value 3, karena kalo 2, durasi 1 hari ga akan keliatan dalam decimal
                    If DummyRefDate.Date < vRefDate.Date Then
                        DummyRetval += Math.Round((vRefDate.Date - DummyRefDate.Date).Days / 366, 3, MidpointRounding.AwayFromZero)
                    End If

                    Return DummyRetval
                Else
                    Return 0
                End If
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function


    Public Sub GetNewOverrideSettings(ByVal vPrimary As Double, _
                ByRef IsOverrideEarlyOTRule As Boolean, ByRef IsCalculateEarlyOT As Boolean, _
                ByRef IsOverrideLateRule As Boolean, ByRef IsCalculateLate As Boolean, _
                ByRef IsOverrideEarlyOutRule As Boolean, ByRef IsCalculateEarlyOut As Boolean, _
                ByRef IsOverrideOTRule As Boolean, ByRef IsCalculateOT As Boolean, _
                ByRef vCalculateAfterEarlyOT As Double, ByRef vCalculateAfterLate As Double, ByRef vCalculateAfterEarlyOut As Double, ByRef vCalculateAfterOT As Double, _
                ByRef vRoundingTypeEarlyOT As Integer, ByRef vRoundingTypeLate As Integer, ByRef vRoundingTypeEarlyOut As Integer, ByRef vRoundingTypeOT As Integer, _
                ByRef vRoundingValueEarlyOT As Decimal, ByRef vRoundingValueLate As Decimal, ByRef vRoundingValueEarlyOut As Decimal, ByRef vRoundingValueOT As Decimal, _
                ByRef vMaxValueEarlyOT As Double, ByRef vMaxValueLate As Double, ByRef vMaxValueEarlyOut As Double, ByRef vMaxValueOT As Double, _
                ByRef vCalculateEarlyOTTolerance As Integer, ByRef vCalculateLateTolerance As Integer, ByRef vCalculateEarlyOutTolerance As Integer, ByRef vCalculateOTTolerance As Integer, _
                ByRef vCalculateEarlyOTOnWeekend As Integer, ByRef vCalculateLateOnWeekend As Integer, ByRef vCalculateEarlyOutOnWeekend As Integer, ByRef vCalculateOTOnWeekend As Integer, _
                ByRef vCalculateEarlyOTOnHoliday As Integer, ByRef vCalculateLateOnHoliday As Integer, ByRef vCalculateEarlyOutOnHoliday As Integer, ByRef vCalculateOTOnHoliday As Integer)
        Dim Dta As New DseMasterKaryawanTableAdapters.employee_override_settingsTableAdapter
        Dim Tbl As New DseMasterKaryawan.employee_override_settingsDataTable
        Dim Row As DseMasterKaryawan.employee_override_settingsRow
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            If Dta.FillBy_Primary(Tbl, vPrimary) > 0 Then
                Row = Tbl.Rows(0)
                With Row
                    If .override_early_ot = 0 Then
                        IsOverrideEarlyOTRule = False
                    Else
                        IsOverrideEarlyOTRule = True
                    End If


                    If .is_early_ot = 0 Then
                        IsCalculateEarlyOT = False
                        vCalculateAfterEarlyOT = 0
                        vRoundingTypeEarlyOT = 0
                        vRoundingValueEarlyOT = 0
                        vMaxValueEarlyOT = 0
                    Else
                        IsCalculateEarlyOT = True
                        vCalculateAfterEarlyOT = .calculate_after_early_OT
                        vRoundingTypeEarlyOT = .rounding_type_early_ot
                        vRoundingValueEarlyOT = .rounding_value_early_ot
                        vMaxValueEarlyOT = .max_value_early_ot
                    End If

                    If .override_late = 0 Then
                        IsOverrideLateRule = False
                    Else
                        IsOverrideLateRule = True
                    End If


                    If .is_late = 0 Then
                        IsCalculateLate = False
                        vCalculateAfterLate = 0
                        vRoundingTypeLate = 0
                        vRoundingValueLate = 0
                        vMaxValueLate = 0
                    Else
                        IsCalculateLate = True
                        vCalculateAfterLate = .calculate_after_late
                        vRoundingTypeLate = .rounding_type_late
                        vRoundingValueLate = .rounding_value_late
                        vMaxValueLate = .max_value_late
                    End If


                    If .override_early_out = 0 Then
                        IsOverrideEarlyOutRule = False
                    Else
                        IsOverrideEarlyOutRule = True
                    End If


                    If .is_early_out = 0 Then
                        IsCalculateEarlyOut = False
                        vCalculateAfterEarlyOut = 0
                        vRoundingTypeEarlyOut = 0
                        vRoundingValueEarlyOut = 0
                        vMaxValueEarlyOut = 0
                    Else
                        IsCalculateEarlyOut = True
                        vCalculateAfterEarlyOut = .calculate_after_early_out
                        vRoundingTypeEarlyOut = .rounding_type_early_out
                        vRoundingValueEarlyOut = .rounding_value_early_out
                        vMaxValueEarlyOut = .max_value_early_out
                    End If


                    If .override_ot = 0 Then
                        IsOverrideOTRule = False
                    Else
                        IsOverrideOTRule = True
                    End If

                    If .is_ot = 0 Then
                        IsCalculateOT = False
                        vCalculateAfterOT = 0
                        vRoundingTypeOT = 0
                        vRoundingValueOT = 0
                        vMaxValueOT = 0
                    Else
                        IsCalculateOT = True
                        vCalculateAfterOT = .calculate_after_ot
                        vRoundingTypeOT = .rounding_type_ot
                        vRoundingValueOT = .rounding_value_ot
                        vMaxValueOT = .max_value_ot
                    End If

                    vCalculateEarlyOTTolerance = .calculate_early_ot_tolerance
                    vCalculateLateTolerance = .calculate_late_tolerance
                    vCalculateEarlyOutTolerance = .calculate_early_out_tolerance
                    vCalculateOTTolerance = .calculate_ot_tolerance

                    vCalculateEarlyOTOnWeekend = .calculate_early_ot_on_weekend
                    vCalculateEarlyOTOnHoliday = .calculate_early_ot_on_holiday
                    vCalculateLateOnWeekend = .calculate_late_on_weekend
                    vCalculateLateOnHoliday = .calculate_late_on_holiday
                    vCalculateEarlyOutOnWeekend = .calculate_early_out_on_weekend
                    vCalculateEarlyOutOnHoliday = .calculate_early_out_on_holiday
                    vCalculateOTOnWeekend = .calculate_ot_on_weekend
                    vCalculateOTOnHoliday = .calculate_ot_on_holiday
                End With

            Else
                IsOverrideEarlyOTRule = False
                IsCalculateEarlyOT = False
                IsOverrideLateRule = False
                IsCalculateLate = False
                IsOverrideEarlyOutRule = False
                IsCalculateEarlyOut = False
                IsOverrideOTRule = False
                IsCalculateOT = False
                vCalculateAfterOT = 0
                vRoundingTypeOT = 0
                vRoundingValueOT = 0
                vMaxValueOT = 0
                vCalculateAfterEarlyOut = 0
                vRoundingTypeEarlyOut = 0
                vRoundingValueEarlyOut = 0
                vMaxValueEarlyOut = 0
                vCalculateAfterLate = 0
                vRoundingTypeLate = 0
                vRoundingValueLate = 0
                vMaxValueLate = 0
                vCalculateAfterEarlyOT = 0
                vRoundingTypeEarlyOT = 0
                vRoundingValueEarlyOT = 0
                vMaxValueEarlyOT = 0
                vCalculateEarlyOTTolerance = 1
                vCalculateLateTolerance = 1
                vCalculateEarlyOutTolerance = 1
                vCalculateOTTolerance = 1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Dta.Dispose()
        End Try
    End Sub

    'aul 09 jun 16
    Public Function GetValue(ByVal vPrimaryKaryawan As Double, ByVal vPrimaryKompetensi As Double) As String
        Dim DTA As New DseMasterKaryawanTableAdapters.tblmaster_karyawan_item_kompetensiTableAdapter
        Dim Dummy As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetValue(vPrimaryKaryawan, vPrimaryKompetensi)
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

    Public Function GetJmlJobPositionKompetensi(ByVal vPrimaryItemKompetensi As Double, ByVal vPrimaryJobPosition As _
    Double) As Double
        Dim DTA As New DseMasterKaryawanTableAdapters.DataTable1TableAdapter
        Dim Dummy As Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            Dummy = DTA.ScalarQuery_GetJmlJobPositionItemKompetensi(vPrimaryJobPosition, vPrimaryItemKompetensi)
            If Dummy.HasValue = True Then
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
	
    Public Function GetNPWPCategoryStr(ByVal vInt_Category As Integer) As String
        Dim vStr_Category As String = ""
        Try
            If vInt_Category = 0 Then
                vStr_Category = "TK0"
            ElseIf vInt_Category = 1 Then
                vStr_Category = "TK1"
            ElseIf vInt_Category = 2 Then
                vStr_Category = "TK2"
            ElseIf vInt_Category = 3 Then
                vStr_Category = "TK3"
            ElseIf vInt_Category = 4 Then
                vStr_Category = "K0"
            ElseIf vInt_Category = 5 Then
                vStr_Category = "K1"
            ElseIf vInt_Category = 6 Then
                vStr_Category = "K2"
            ElseIf vInt_Category = 7 Then
                vStr_Category = "K3"
            End If
            Return vStr_Category
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function ApakahOverrideOTOrder_Emp(ByVal vPrimary As Double) As Boolean
        Dim DTA As New DseMasterKaryawanTableAdapters.tblmaster_karyawanTableAdapter
        Dim Dummy As System.Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetOverrideOTOrder(vPrimary)
            If Dummy.HasValue Then
                If Dummy.Value > 0 Then
                    Return True
                Else
                    Return False
                End If
            End If

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahUseOTOrder_Emp(ByVal vPrimary As Double) As Boolean
        Dim DTA As New DseMasterKaryawanTableAdapters.tblmaster_karyawanTableAdapter
        Dim Dummy As System.Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetUserOTOrder(vPrimary)
            If Dummy.HasValue Then
                If Dummy.Value > 0 Then
                    Return True
                Else
                    Return False
                End If
            End If

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahUseNotifEmailOTOrder_Emp(ByVal vPrimary As Double) As Boolean
        Dim DTA As New DseMasterKaryawanTableAdapters.tblmaster_karyawanTableAdapter
        Dim Dummy As System.Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetNotifEmailOTOrder(vPrimary)
            If Dummy.HasValue Then
                If Dummy.Value > 0 Then
                    Return True
                Else
                    Return False
                End If
            End If

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetSalaryCalcNo(ByVal vPrimaryEmployee As Double, ByVal vRefPeriod As Date) As String
        Dim DTA As New DseMasterKaryawanTableAdapters.tblmaster_karyawanTableAdapter
        Dim Retval As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetSalaryCalcNo(vRefPeriod, vPrimaryEmployee)
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

    Public Function GetSalaryStartPeriod(ByVal vPrimaryEmployee As Double, ByVal vRefPeriod As Date) As String
        Dim DTA As New DseMasterKaryawanTableAdapters.tblmaster_karyawanTableAdapter
        Dim Retval As Nullable(Of Date)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetSalaryStartPeriod(vRefPeriod, vPrimaryEmployee)
            If Retval.HasValue = True Then
                Return Retval
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

    Public Function GetSalaryEndPeriod(ByVal vPrimaryEmployee As Double, ByVal vRefPeriod As Date) As String
        Dim DTA As New DseMasterKaryawanTableAdapters.tblmaster_karyawanTableAdapter
        Dim Retval As Nullable(Of Date)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetSalaryEndPeriod(vRefPeriod, vPrimaryEmployee)
            If Retval.HasValue = True Then
                Return Retval
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

    Public Function GetPrimarySalaryFormula(ByVal vPrimaryKaryawan As Double) As Double
        Dim DTA As New DseMasterSalaryFormulaTableAdapters.tblmaster_salary_formula_component_detailemployeeTableAdapter
        Dim RetVal As Nullable(Of Double)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetPrimaryTransaksi(vPrimaryKaryawan)
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

    Public Function GetLosesTaxObligation(ByVal vPrimaryKary As Double) As Integer
        Dim DTA As New DseMasterKaryawanTableAdapters.tblmaster_karyawanTableAdapter
        Dim RetVal As Nullable(Of Integer)

        Try
            RetVal = DTA.ScalarQuery_GetLosesTaxObligations(vPrimaryKary)
            If RetVal.HasValue = True Then
                Return RetVal.Value
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
