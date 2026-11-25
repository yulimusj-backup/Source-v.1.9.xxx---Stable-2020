Option Explicit On
Option Strict Off

Public Class ClsMasterAsset
    Dim ObjCD As New Core_Module.ClsCekData

    Dim mPrimary As Double = -1
    Dim mAssetCode As String = ""
    Dim mStatusAsset As Integer = EnumStatusAsset.On_Progress
    Dim mDesc As String = ""
    Dim mPrimaryGroupAssetCode As Double = -1
    Dim mQty As Decimal = 0
    Dim mPrimaryUnit As Double = -1
    Dim mremarks As String = "'"
    Dim mAcquisitionType As Integer = EnumAcquisitionType.Define_Value
    Dim mAcquisitionFinish As Integer = EnumAcquisitionFinish._true
    Dim mDepreciateThisAsset As Integer = EnumDepreciateThisAsset._true
    Dim mDepreciationMethod As Integer = EnumDepreciationMethod.Straight_Line
    Dim mDepreciationLife As Decimal = 0
    Dim mAcquiredValue As Decimal = 0
    Dim mSalvageValue As Decimal = 0
    Dim mBookValue As Decimal = 0
    Dim mStartDepreciation As Date = "01-01-1950"
    Dim mFrecuency As Decimal = EnumFrequency.Monthly
    Dim mDepreciateOn As Date = "01-01-1950"
    Dim mPrimaryAkunOnProcess As Double = -1
    Dim mPrimaryAkunAsset As Double = -1
    Dim mPrimaryAkunDepreciationAsset As Double = -1
    Dim mPrimaryAkunAccumDepreciation As Double = -1
    Dim mPrimaryAkunDisposalIncome As Double = -1
    Dim mPrimaryAkunGainLossDisposal As Double = -1
    Dim mPrimaryJournalSetFinish As Double = -1
    Dim mTanggalSetFinish As Date = "01-01-1950"

    Dim mCreatedBy As String = ""
    Dim mCreatedDate As DateTime = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As DateTime = Now
    Dim mTimestamp As Date = Now

    'elemen stock yang digunakan untuk antrian simpan data
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianAssetModule
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAAsset As New DseMasterAssetTableAdapters.tbltrans_antrianTableAdapter
    Dim ObjRnd As New Random

    Public Enum EnumAkunAsset As Integer
        OnProcess
        Asset
        DeprciationExpense
        AccumDepreciation
        DisposalIncome
        GainLossDisposal
    End Enum

    Public Enum EnumStatusAsset As Integer
        On_Progress
        Active
        Dispose
    End Enum

    Public Enum EnumAcquisitionType As Integer
        Define_Value
        BuyOrBuild
    End Enum

    Public Enum EnumAcquisitionFinish As Integer
        _true
        _false
    End Enum

    Public Enum EnumDepreciateThisAsset As Integer
        _true
        _false
    End Enum

    Public Enum EnumDepreciationMethod As Integer
        Straight_Line
        Sum_year_Digits
        Double_Declining_Balance
        No_Depreciation_Method
    End Enum

    Public Enum EnumFrequency As Integer
        Monthly
        Yearly
    End Enum

    Public Property Primary() As Double
        Get
            Return Me.mPrimary
        End Get
        Set(ByVal value As Double)
            Me.mPrimary = value
        End Set
    End Property

    Public Property AssetCode() As String
        Get
            Return Me.mAssetCode
        End Get
        Set(ByVal value As String)
            Me.mAssetCode = value
        End Set
    End Property

    Public Property StatusAsset() As EnumStatusAsset
        Get
            Return Me.mStatusAsset
        End Get
        Set(ByVal value As EnumStatusAsset)
            Me.mStatusAsset = value
        End Set
    End Property

    Public Property Desc() As String
        Get
            Return Me.mDesc
        End Get
        Set(ByVal value As String)
            Me.mDesc = value
        End Set
    End Property

    Public Property PrimaryAssetGroupCode() As Double
        Get
            Return Me.mPrimaryGroupAssetCode
        End Get
        Set(ByVal value As Double)
            Me.mPrimaryGroupAssetCode = value
        End Set
    End Property

    Public Property Qty() As Decimal
        Get
            Return Me.mQty
        End Get
        Set(ByVal value As Decimal)
            Me.mQty = value
        End Set
    End Property

    Public Property PrimaryUnit() As Double
        Get
            Return Me.mPrimaryUnit
        End Get
        Set(ByVal value As Double)
            Me.mPrimaryUnit = value
        End Set
    End Property

    Public Property Remarks() As String
        Get
            Return Me.mremarks
        End Get
        Set(ByVal value As String)
            Me.mremarks = value
        End Set
    End Property

    Public Property AcquicitionType() As EnumAcquisitionType
        Get
            Return Me.mAcquisitionType
        End Get
        Set(ByVal value As EnumAcquisitionType)
            Me.mAcquisitionType = value
        End Set
    End Property

    Public Property AcquicitionFinish() As Integer
        Get
            Return Me.mAcquisitionFinish
        End Get
        Set(ByVal value As Integer)
            Me.mAcquisitionFinish = value
        End Set
    End Property

    Public Property DepreciationThisAsset() As EnumDepreciateThisAsset
        Get
            Return Me.mDepreciateThisAsset
        End Get
        Set(ByVal value As EnumDepreciateThisAsset)
            Me.mDepreciateThisAsset = value
        End Set
    End Property

    Public Property DepreciationMethod() As EnumDepreciationMethod
        Get
            Return Me.mDepreciationMethod
        End Get
        Set(ByVal value As EnumDepreciationMethod)
            Me.mDepreciationMethod = value
        End Set
    End Property

    Public Property DepreciationLife() As Decimal
        Get
            Return Me.mDepreciationLife
        End Get
        Set(ByVal value As Decimal)
            Me.mDepreciationLife = value
        End Set
    End Property

    Public Property AcquiredValue() As Decimal
        Get
            Return Me.mAcquiredValue
        End Get
        Set(ByVal value As Decimal)
            Me.mAcquiredValue = value
        End Set
    End Property

    Public Property SalvageValue() As Decimal
        Get
            Return Me.mSalvageValue
        End Get
        Set(ByVal value As Decimal)
            Me.mSalvageValue = value
        End Set
    End Property

    Public Property BookValue() As Decimal
        Get
            Return Me.mBookValue
        End Get
        Set(ByVal value As Decimal)
            Me.mBookValue = value
        End Set
    End Property

    Public Property StartDepreciateOn() As Date
        Get
            Return Me.mStartDepreciation
        End Get
        Set(ByVal value As Date)
            Me.mStartDepreciation = value
        End Set
    End Property

    Public Property Frequency() As EnumFrequency
        Get
            Return Me.mFrecuency
        End Get
        Set(ByVal value As EnumFrequency)
            Me.mFrecuency = value
        End Set
    End Property

    Public Property DepreciateOn() As Date
        Get
            Return Me.mDepreciateOn
        End Get
        Set(ByVal value As Date)
            Me.mDepreciateOn = value
        End Set
    End Property

    Public Property PrimaryAkunOnProcess() As Double
        Get
            Return Me.mPrimaryAkunOnProcess
        End Get
        Set(ByVal value As Double)
            Me.mPrimaryAkunOnProcess = value
        End Set
    End Property

    Public Property PrimaryAkunAsset() As Double
        Get
            Return Me.mPrimaryAkunAsset
        End Get
        Set(ByVal value As Double)
            Me.mPrimaryAkunAsset = value
        End Set
    End Property

    Public Property PrimaryAkunDepreciationAsset() As Double
        Get
            Return Me.mPrimaryAkunDepreciationAsset
        End Get
        Set(ByVal value As Double)
            Me.mPrimaryAkunDepreciationAsset = value
        End Set
    End Property

    Public Property PrimaryAkunAccumDepreciation() As Double
        Get
            Return Me.mPrimaryAkunAccumDepreciation
        End Get
        Set(ByVal value As Double)
            Me.mPrimaryAkunAccumDepreciation = value
        End Set
    End Property

    Public Property PrimaryAkunDisposalIncome() As Double
        Get
            Return Me.mPrimaryAkunDisposalIncome
        End Get
        Set(ByVal value As Double)
            Me.mPrimaryAkunDisposalIncome = value
        End Set
    End Property

    Public Property PrimaryAkunGainLossDisposal() As Double
        Get
            Return Me.mPrimaryAkunGainLossDisposal
        End Get
        Set(ByVal value As Double)
            Me.mPrimaryAkunGainLossDisposal = value
        End Set
    End Property

    Public Property PrimaryJournalSetfinish() As Double
        Get
            Return Me.mPrimaryJournalSetFinish
        End Get
        Set(ByVal value As Double)
            Me.mPrimaryJournalSetFinish = value
        End Set
    End Property

    Public Property TanggalSetfinish() As Date
        Get
            Return Me.mTanggalSetFinish
        End Get
        Set(ByVal value As Date)
            Me.mTanggalSetFinish = value
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

    Public Function GetMAXDateDepreciation() As Date
        Dim DTA As New DseMasterAssetTableAdapters.tbltrans_jurnalTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetMaxDateDepresiasi().HasValue Then
                Return DTA.ScalarQuery_GetMaxDateDepresiasi().Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function ApakahDataBenar(ByVal vBaru As Boolean, ByVal vShowMsg As Boolean) As Boolean
        Try
            'tidak boleh ada kode yang sama
            If vBaru = True Then
                If Me.ApakahSudahAda(Me.mAssetCode) = True Then
                    If vShowMsg = True Then
                        MsgBox("Duplication in asset code, please change your code", MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            Else
                Dim CodeLama As String = Me.GetCode(Me.mPrimary)
                If CodeLama <> Me.mAssetCode Then
                    If Me.ApakahSudahAda(Me.mAssetCode) = True Then
                        If vShowMsg = True Then
                            MsgBox("Duplication in asset code, please change your code", MsgBoxStyle.Information, "Information")
                        End If
                        Return False
                    End If
                End If
            End If

            'yang harus diisikan:
            '1. kode
            If Me.mAssetCode = "" Then
                If vShowMsg = True Then
                    MsgBox("Please input asset code.", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If


            'kalo mpe sini ok
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function ApakahSudahAda(ByVal vCode As String) As Boolean
        Dim DTA As New DseMasterAssetTableAdapters.tblmaster_assetTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimary(vCode).HasValue Then
                If DTA.ScalarQuery_GetPrimary(vCode).Value > 0 Then
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

    Public Overridable Function ApakahBolehEdit(ByVal vPrimary As Double, ByVal StrictEditting As Boolean, Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseMasterAssetTableAdapters.tblmaster_assetTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If StrictEditting = True Then
                If DTA.ScalarQuery_GetStatus(vPrimary).Value = ClsMasterAsset.EnumStatusAsset.Dispose Then
                    If TampilkanPesan = True Then
                        MsgBox("This asset can't edit because status asset dispose.", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                Else
                    Return True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahBolehHapus(ByVal vPrimary As Double, Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseMasterAssetTableAdapters.tblmaster_assetTableAdapter
        Dim ObjJournal As New ClsTransaksiJournal1
        Dim ObjStockInOut As New clsTransaksiStockInOut1
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'If DTA.ScalarQuery_GetStatus(vPrimary).Value = Me.EnumStatusAsset.Dispose Then
            '    If TampilkanPesan = True Then
            '        MsgBox("This asset can't delete because status asset dispose.", _
            '        MsgBoxStyle.Information, "Information")
            '    End If
            '    Return False
            'End If

            If Me.CekAdaTransaksi(vPrimary) = True Then
                Return False
            End If


            Return True


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Erorr")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function CekAdaTransaksi(ByVal vPrimary As Double) As Boolean
        Dim ObjStockInOut As New clsTransaksiStockInOut1
        Dim ObjJurnal As New ClsTransaksiJournal1
        Try

            If vPrimary = -1 Then
                Return False
            End If
            'cek di material usage
            If ObjStockInOut.GetJmlByAssetNInOut(vPrimary, clsTransaksiStockInOut1.EnumInOutAdjust.OutAdjustMaterialUsageAssetAcquisition) > 0 Then
                'ada
                MsgBox("This asset can't delete because this asset data is rerferring to transaction acquisition material usage.", MsgBoxStyle.Information, "Information")
                Return True
            End If
            If ObjStockInOut.GetJmlByAssetNInOut(vPrimary, clsTransaksiStockInOut1.EnumInOutAdjust.OutAdjustMaterialUsageAssetRevaluation) > 0 Then
                'ada
                MsgBox("This asset can't delete because this asset data is rerferring to transaction acquisition revaluation material usage.", MsgBoxStyle.Information, "Information")
                Return True
            End If
            'cek di acquistion other
            If ObjJurnal.GetJmlByTipeNPrimaryAsset(vPrimary, ClsTransaksiJournal1.EnumTipeJurnal2.AssetAcquistion_Other) > 0 Then
                'ada
                MsgBox("This asset can't delete because this asset data is rerferring to transaction acquistion other.", MsgBoxStyle.Information, "Information")
                Return True
            End If
            'cek di revaluation other
            If ObjJurnal.GetJmlByTipeNPrimaryAsset(vPrimary, ClsTransaksiJournal1.EnumTipeJurnal2.AssetRevaluation_Other) > 0 Then
                'ada
                MsgBox("This asset can't delete because this asset data is rerferring to transaction revaluation other.", MsgBoxStyle.Information, "Information")
                Return True
            End If
            'cek di depresiasi
            If ObjJurnal.GetJmlByTipeNPrimaryAsset(vPrimary, ClsTransaksiJournal1.EnumTipeJurnal2.AssetDepreciation) > 0 Then
                'ada
                MsgBox("This asset can't delete because this asset data is rerferring to transaction depreciation.", MsgBoxStyle.Information, "Information")
                Return True
            End If
            'cek di disposal
            If ObjJurnal.GetJmlByTipeNPrimaryAsset(vPrimary, ClsTransaksiJournal1.EnumTipeJurnal2.AssetDisposal) > 0 Then
                'ada
                MsgBox("This asset can't delete because this asset data is rerferring to transaction disposal.", MsgBoxStyle.Information, "Information")
                Return True
            End If

            'jika tidak ada
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjStockInOut = Nothing
            ObjJurnal = Nothing
        End Try
    End Function

    Public Function CariDetail(ByVal vPrimary As Double) As Boolean
        Dim DTA As New DseMasterAssetTableAdapters.tblmaster_assetTableAdapter
        Dim TBL As New DseMasterAsset.tblmaster_assetDataTable
        Dim row As DseMasterAsset.tblmaster_assetRow
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_Primary(TBL, vPrimary)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = .Primary
                    Me.mAssetCode = .Asset_Code
                    Me.mStatusAsset = .Status_Asset
                    Me.mDesc = .Description
                    Me.mPrimaryGroupAssetCode = .Primary_Group_Asset_Code
                    Me.mQty = .Qty
                    Me.PrimaryUnit = .Primary_Unit
                    Me.mremarks = .Remarks
                    Me.mAcquisitionType = .Acquisition_Type
                    Me.mAcquisitionFinish = .Acquisition_Finish
                    Me.mDepreciateThisAsset = .Depreciate_This_Asset
                    Me.mDepreciationMethod = .Depreciation_Method
                    Me.mDepreciationLife = .Depreciation_Life
                    Me.mAcquiredValue = .Acquired_Value
                    Me.mSalvageValue = .Salvage_Value
                    Me.mBookValue = .Book_Value
                    Me.mFrecuency = .Frequency
                    Me.mStartDepreciation = .Start_Depreciate_On
                    Me.mDepreciateOn = .Depreciate_On
                    Me.mPrimaryAkunOnProcess = .Primary_Akun_On_Process
                    Me.mPrimaryAkunAsset = .Primary_Akun_Asset
                    Me.mPrimaryAkunDepreciationAsset = .Primary_Akun_Depreciation_Asset
                    Me.mPrimaryAkunAccumDepreciation = .Primary_Akun_Accum_Depreciation
                    Me.mPrimaryAkunDisposalIncome = .Primary_Akun_Dispossal_Income
                    Me.mPrimaryAkunGainLossDisposal = .Primary_Akun_Gain_Loss_Dispossal
                    Me.mPrimaryJournalSetFinish = .Primary_Jounal_Set_Finish
                    Me.mTanggalSetFinish = .Tanggal_Set_Finish
                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
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

    Public Function GetJmlByGroup(ByVal vPrimaryGroup As Double) As Integer
        Dim dta As New DseMasterAssetTableAdapters.tblmaster_assetTableAdapter

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.ScalarQuery_GetJmlByPrimaryGroup(vPrimaryGroup).HasValue = True Then
                Return dta.ScalarQuery_GetJmlByPrimaryGroup(vPrimaryGroup).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetCode(ByVal vPrimary As Double) As String
        Dim DTA As New DseMasterAssetTableAdapters.tblmaster_assetTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetAssetCode(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetAssetCode(vPrimary)
            End If
            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTimestamp(ByVal vPrimary As Double) As Date
        Dim DTA As New DseMasterAssetTableAdapters.tblmaster_assetTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetTimestamp(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetTimestamp(vPrimary)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Sub GetNoAntrian()
        Try
            Me.DTAAsset.Connection.ConnectionString = Core_Module.GetConstring
            Me.vIdentifier = Me.ObjRnd.Next(1000000)
            'jangan sampai sama dengan identifier yang sudah ada
            While Me.DTAAsset.ScalarQuery_GetJmlIdentifier(Me.vTipeAntrian, Me.vIdentifier).Value > 0
                Me.vIdentifier = Me.ObjRnd.Next(1000000)
            End While
            'masukan data baru ke antrian
            Me.DTAAsset.Insert(Me.vTipeAntrian, Me.vIdentifier)

            'cek jangan sampai ada 2 data yang sama
            While Me.DTAAsset.ScalarQuery_GetJmlIdentifier(Me.vTipeAntrian, Me.vIdentifier).Value > 1
                'bila ada yang sama maka harus masukkan data baru
                DTAAsset.DeleteQuery(Me.vTipeAntrian, Me.vIdentifier)
                'minta identifier baru
                Me.vIdentifier = Me.ObjRnd.Next(1000000)
                'jangan sampai sama dengan identifier yang sudah ada
                While Me.DTAAsset.ScalarQuery_GetJmlIdentifier(Me.vTipeAntrian, Me.vIdentifier).Value > 0
                    Me.vIdentifier = Me.ObjRnd.Next(1000000)
                    'masukan data baru ke antrian
                    Me.DTAAsset.Insert(Me.vTipeAntrian, Me.vIdentifier)
                End While
            End While

            'ambil no antrian
            Me.vPrimary_Antrian = DTAAsset.Scalar_Query_GetPrimary(Me.vIdentifier, Me.vTipeAntrian).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Overridable Function SimpanBaru(ByVal vUserCode As String) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True

        Try
            Me.DTAAsset.Connection.ConnectionString = Core_Module.GetConstring
            'ambil no antrian
            Me.GetNoAntrian()

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If Me.DTAAsset.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                        'clock antri matikan
                        Me.Clock_Antri.Stop()
                        'sudah saatnya melakukan proses simpan baru
                        retval = Me.SimpanBaru1(vUserCode)
                        'hapus data antri untuk transaksi ini
                        If Me.DTAAsset.Delete(Me.vPrimary_Antrian) = 0 Then
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
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function SimpanBaru1(ByVal vUserCode As String) As Boolean
        Dim DTA As New DseMasterAssetTableAdapters.tblmaster_assetTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.InsertQuery(Me.mAssetCode, Me.mStatusAsset, Me.mDesc, Me.mPrimaryGroupAssetCode, Me.mQty, _
            Me.mPrimaryUnit, Me.mremarks, Me.mAcquisitionType, Me.mAcquisitionFinish, Me.mDepreciateThisAsset, _
            Me.mDepreciationMethod, Me.mDepreciationLife, Me.mAcquiredValue, Me.mSalvageValue, Me.mBookValue, _
            Me.StartDepreciateOn, Me.mFrecuency, Me.mDepreciateOn, Me.mPrimaryAkunOnProcess, Me.mPrimaryAkunAsset, _
            Me.mPrimaryAkunDepreciationAsset, Me.mPrimaryAkunAccumDepreciation, Me.mPrimaryAkunDisposalIncome, _
            Me.mPrimaryAkunGainLossDisposal, Me.mPrimaryJournalSetFinish, Me.mTanggalSetFinish, vUserCode, Now, _
            vUserCode, Now, Now) > 0 Then
                Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mAssetCode)
                Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                Me.mCreatedDate = Me.mTimestamp
                Me.mModifiedDate = Me.mTimestamp
                Me.mCreatedBy = vUserCode
                Me.mModifiedBy = vUserCode
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

    Public Overridable Function SimpanPerubahan(ByVal vUserCode As String) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True

        Try
            Me.DTAAsset.Connection.ConnectionString = Core_Module.GetConstring
            'ambil no antrian
            Me.GetNoAntrian()

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If Me.DTAAsset.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                        'clock antri matikan
                        Me.Clock_Antri.Stop()
                        'sudah saatnya melakukan proses simpan baru
                        retval = Me.SimpanPerubahan1(vUserCode)
                        'hapus data antri untuk transaksi ini
                        If Me.DTAAsset.Delete(Me.vPrimary_Antrian) = 0 Then
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
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function SimpanPerubahan1(ByVal vUserCode As String) As Boolean
        Dim DTA As New DseMasterAssetTableAdapters.tblmaster_assetTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.UpdateQuery(Me.mAssetCode, Me.mStatusAsset, Me.mDesc, Me.mPrimaryGroupAssetCode, Me.mQty, _
            Me.mPrimaryUnit, Me.mremarks, Me.mAcquisitionType, Me.mAcquisitionFinish, Me.mDepreciateThisAsset, _
            Me.mDepreciationMethod, Me.mDepreciationLife, Me.mAcquiredValue, Me.mSalvageValue, Me.mBookValue, _
            Me.StartDepreciateOn, Me.mFrecuency, Me.mDepreciateOn, Me.mPrimaryAkunOnProcess, Me.mPrimaryAkunAsset, _
            Me.mPrimaryAkunDepreciationAsset, Me.mPrimaryAkunAccumDepreciation, Me.mPrimaryAkunDisposalIncome, _
            Me.mPrimaryAkunGainLossDisposal, Me.mPrimaryJournalSetFinish, Me.mTanggalSetFinish, vUserCode, Now, _
            vUserCode, Now, Now, Me.mPrimary, Me.mTimestamp) > 0 Then
                Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.Primary)
                Me.mModifiedDate = Me.mTimestamp
                Me.mModifiedBy = vUserCode

                'Panggil Fungsi Jurnal Set Finish
                Me.Simpan_SetFinish_Asset(Me.mPrimary, vUserCode)

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

    Public Function Hapus(ByVal vPrimaryTransaksi As Double, ByVal vUser As String) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True

        Try
            Me.DTAAsset.Connection.ConnectionString = Core_Module.GetConstring
            'ambil no antrian
            Me.GetNoAntrian()
            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If Me.DTAAsset.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                        'clock antri matikan
                        Me.Clock_Antri.Stop()
                        'sudah saatnya melakukan proses simpan baru
                        retval = Me.Hapus1(vPrimaryTransaksi)
                        'hapus data antri untuk transaksi ini
                        If Me.DTAAsset.Delete(Me.vPrimary_Antrian) = 0 Then
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
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function Hapus1(ByVal vPrimary As Double) As Boolean
        Dim DTA As New DseMasterAssetTableAdapters.tblmaster_assetTableAdapter
        Dim PrimaryJournal As Double = -1
        Dim DTA_Journal As New DseMasterAssetTableAdapters.tbltrans_jurnal_setfinishTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA_Journal.Connection.ConnectionString = DTA.Connection.ConnectionString
            If Me.ApakahBolehHapus(vPrimary) = True Then
                PrimaryJournal = Me.GetPrimaryJournal(vPrimary)
                DTA_Journal.DeleteQuery_Jurnal_Primary(PrimaryJournal)
                'Hapus data
                DTA.Delete(vPrimary)
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA_Journal.Dispose()
        End Try
    End Function

    Private Function GetPrimaryJournal(ByVal vPrimaryAsset As Double) As Double
        Dim DTA As New DseMasterAssetTableAdapters.tbltrans_jurnal_setfinishTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimaryByPrimaryAsset(vPrimaryAsset).HasValue Then
                Return DTA.ScalarQuery_GetPrimaryByPrimaryAsset(vPrimaryAsset).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseMasterAsset.DataTable1DataTable, _
    ByVal SearchCriteria As EnumSearchCriteriaAsset, ByVal tsl As ToolStripLabel) As Boolean
        Dim DTA As New DseMasterAssetTableAdapters.DataTable1TableAdapter
        Dim row As DseMasterAsset.DataTable1Row
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If SearchCriteria = EnumSearchCriteriaAsset.ByAssetCode Then
                QueryResult = DTA.FillBy_LikeAssetCode(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteriaAsset.ByDescription Then
                QueryResult = DTA.FillBy_LikeDesc(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteriaAsset.ByRemark Then
                QueryResult = DTA.FillBy_LikeRemarks(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteriaAsset.ByGroupCode Then
                QueryResult = DTA.FillBy_GroupCode(TBL, Criteria + "%")
            End If

            'update dgv
            DGV.DataBindings.Clear()
            DGV.DataSource = Nothing
            DGV.Rows.Clear()
            For Each row In TBL.Rows
                DGV.Rows.Add(row.Primary, row.Asset_Code, row.Description, row.Remarks)
            Next
            'If DGV.ColumnCount > 0 Then
            '    DGV.Columns(0).Visible = False
            '    DGV.Columns(1).HeaderText = "Group code"
            '    DGV.Columns(0).Frozen = True
            '    DGV.Columns(1).Frozen = True
            'End If


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


    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As _
      EnumSearchCriteriaAsset) As Double
        Dim DTA As New DseMasterAssetTableAdapters.DataTable1TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vSearchBy = EnumSearchCriteriaAsset.ByAssetCode Then
                If DTA.ScalarQuery_GetJmlLikeAssetCode(vKriteria + "%").HasValue Then
                    Return DTA.ScalarQuery_GetJmlLikeAssetCode(vKriteria + "%").Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteriaAsset.ByDescription Then
                If DTA.ScalarQuery_GetJmlLikeDesc(vKriteria + "%").HasValue Then
                    Return DTA.ScalarQuery_GetJmlLikeDesc(vKriteria + "%").Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteriaAsset.ByRemark Then
                If DTA.ScalarQuery_GetJmlLikeRemarks(vKriteria + "%").HasValue Then
                    Return DTA.ScalarQuery_GetJmlLikeRemarks(vKriteria + "%").Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteriaAsset.ByGroupCode Then
                If DTA.ScalarQuery_GetJmlLikeGroupCode(vKriteria + "%").HasValue Then
                    Return DTA.ScalarQuery_GetJmlLikeGroupCode(vKriteria + "%").Value
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

    'Fungsi Update BookValue (digunakan ketika simpan di transaksi revaluation, depreciation, dan disposal)
    Public Function GetAcquiredValue(ByVal vPrimary As Double) As Decimal
        Dim DTA As New DseMasterAssetTableAdapters.tblmaster_assetTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetAcquiredValue(vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetAcquiredValue(vPrimary).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetDepreciatedValue(ByVal vPrimaryAsset As Double) As Decimal
        Dim DTA As New DseMasterAssetTableAdapters.tbltrans_jurnal_item_depresiasi_asetTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetDepreciatedValue(vPrimaryAsset).HasValue Then
                Return DTA.ScalarQuery_GetDepreciatedValue(vPrimaryAsset).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function UpdateBookValue(ByVal vBookValue As Decimal, ByVal vPrimaryAsset As Double) As Boolean
        Dim DTA As New DseMasterAssetTableAdapters.tblmaster_assetTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.UpdateQuery_BookValue(vBookValue, vPrimaryAsset)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function HitungBookValue(ByVal vPrimaryAsset As Double) As Decimal
        Dim AcqValue As Decimal = 0
        Dim RevValue As Decimal
        Dim DeptValue As Decimal
        Dim DisposalValue As Decimal
        Dim Retval As Decimal = 0
        Dim Filter As String = ""
        Dim RowView As DataRowView
        Dim DTAJournal As New DseMasterAssetTableAdapters.tbltrans_jurnalTableAdapter
        Dim TBLJournal As New DseMasterAsset.tbltrans_jurnalDataTable
        Dim RowJournal As DseMasterAsset.tbltrans_jurnalRow
        Dim PrimaryMinJournal As Double = -1
        Dim InValue As Double = 0
        Dim OutValue As Double = 1
        Dim PositifRevValue As Decimal = 0
        Dim NegativeRevValue As Decimal = 0
        Dim DTAPayDetail As New DseMasterAssetTableAdapters.tbltrans_jurnal_item_payment_detailTableAdapter
        Dim TBLPayDetail As New DseMasterAsset.tbltrans_jurnal_item_payment_detailDataTable
        Dim RowPayDetail As DseMasterAsset.tbltrans_jurnal_item_payment_detailRow
        Dim DTADisposal As New DseMasterAssetTableAdapters.tbltrans_jurnal_item_disposalTableAdapter
        Dim TBLDisposal As New DseMasterAsset.tbltrans_jurnal_item_disposalDataTable
        Dim RowDisposal As DseMasterAsset.tbltrans_jurnal_item_disposalRow
        Dim MaterialUsage As Decimal = 0
        Dim DTA_Asset As New DseMasterAssetTableAdapters.tbltransaksi_stockinoutTableAdapter
        Dim TipeTransaksi_MaterialUsageRev As clsTransaksiStockInOut1.EnumInOutAdjust = clsTransaksiStockInOut1.EnumInOutAdjust.OutAdjustMaterialUsageAssetRevaluation

        Try
            DTA_Asset.Connection.ConnectionString = Core_Module.GetConstring
            DTAJournal.Connection.ConnectionString = DTA_Asset.Connection.ConnectionString
            DTAPayDetail.Connection.ConnectionString = DTA_Asset.Connection.ConnectionString
            DTADisposal.Connection.ConnectionString = DTA_Asset.Connection.ConnectionString


            'acq value
            AcqValue = Me.GetAcquiredValue(vPrimaryAsset)
            'Get Revaluation Value (EnumTipeJurnal2 = 4)
            'positive revaluation
            Filter = ""
            PositifRevValue = 0
            TBLPayDetail.DefaultView.RowFilter = ""
            DTAJournal.FillBy_InOut(TBLJournal, vPrimaryAsset, InValue, 4)
            If TBLJournal.Rows.Count > 0 Then
                RowJournal = TBLJournal.Rows(0)
                PrimaryMinJournal = RowJournal.Primary
                DTAPayDetail.FillBy_PrimaryJournal(TBLPayDetail, PrimaryMinJournal)
                For Each RowJournal In TBLJournal.Rows
                    Filter = "[Primary Jurnal] = " + RowJournal.Primary.ToString + ""
                    TBLPayDetail.DefaultView.RowFilter = Filter
                    If TBLPayDetail.DefaultView.Count > 0 Then
                        For Each RowView In TBLPayDetail.DefaultView
                            RowPayDetail = RowView.Row
                            PositifRevValue += RowPayDetail.Amount
                        Next
                    End If
                Next
            End If
            'negative revaluation
            Filter = ""
            NegativeRevValue = 0
            TBLPayDetail.DefaultView.RowFilter = ""
            DTAJournal.FillBy_InOut(TBLJournal, vPrimaryAsset, OutValue, 4)
            If TBLJournal.Rows.Count > 0 Then
                RowJournal = TBLJournal.Rows(0)
                PrimaryMinJournal = RowJournal.Primary
                DTAPayDetail.FillBy_PrimaryJournal(TBLPayDetail, PrimaryMinJournal)
                For Each RowJournal In TBLJournal.Rows
                    Filter = "[Primary Jurnal] = " + RowJournal.Primary.ToString + ""
                    TBLPayDetail.DefaultView.RowFilter = Filter
                    If TBLPayDetail.DefaultView.Count > 0 Then
                        For Each RowView In TBLPayDetail.DefaultView
                            RowPayDetail = RowView.Row
                            NegativeRevValue += RowPayDetail.Amount
                        Next
                    End If
                Next
            End If
            RevValue = PositifRevValue - NegativeRevValue

            'Get Depreciation Value (EnumTipeJurnal2 = 5)
            DeptValue = Me.GetDepreciatedValue(vPrimaryAsset)

            'Get Dispossal Value (EnumTipeJurnal2 = 6)
            Filter = ""
            DisposalValue = 0
            DTAJournal.FillBy_PrimaryAsset(TBLJournal, vPrimaryAsset, 6)
            If TBLJournal.Rows.Count > 0 Then
                RowJournal = TBLJournal.Rows(0)
                PrimaryMinJournal = RowJournal.Primary
                DTADisposal.FillBy_PrimaryJurnal(TBLDisposal, PrimaryMinJournal)
                For Each RowJournal In TBLJournal.Rows
                    Filter = "[primary jurnal] = " + RowJournal.Primary.ToString + ""
                    TBLDisposal.DefaultView.RowFilter = Filter
                    If TBLDisposal.DefaultView.Count > 0 Then
                        For Each RowView In TBLDisposal.DefaultView
                            RowDisposal = RowView.Row
                            DisposalValue += RowDisposal.disposed_asset_amount
                        Next
                    End If
                Next
            End If

            'material usage revaluasi
            If DTA_Asset.ScalarQuery_GetMaterialUsage(CInt(TipeTransaksi_MaterialUsageRev), vPrimaryAsset).HasValue = True Then
                MaterialUsage = DTA_Asset.ScalarQuery_GetMaterialUsage(CInt(TipeTransaksi_MaterialUsageRev), vPrimaryAsset).Value
            Else
                MaterialUsage = 0
            End If

            'Rumus Book Value
            Retval = AcqValue + MaterialUsage + RevValue - DeptValue - DisposalValue

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTADisposal.Dispose()
            DTAJournal.Dispose()
            DTAPayDetail.Dispose()
            TBLDisposal.Dispose()
            TBLJournal.Dispose()
            TBLPayDetail.Dispose()
            DTA_Asset.Dispose()
        End Try
    End Function

    Public Function ApakahSudahAda_JournalRefNumber_SetFinish(ByVal vRefNumber As String) As Boolean
        Dim DTA As New DseMasterAssetTableAdapters.tbltrans_jurnal_setfinishTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJmlByRefNumber(vRefNumber).Value > 0 Then
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

    Public Function Simpan_SetFinish_Asset(ByVal vPrimary As Double, ByVal vUserCode As String) As Boolean
        Dim DTA As New DseMasterAssetTableAdapters.tblmaster_asset_SetFinishTableAdapter
        Dim DTA_InOut As New DseMasterAssetTableAdapters.tbltransaksi_stockinout_setfinishTableAdapter
        Dim DTA_Journal As New DseMasterAssetTableAdapters.tbltrans_jurnal_setfinishTableAdapter
        Dim DTA_Journal_Item As New DseMasterAssetTableAdapters.tbltrans_jurnal_item_setfinishTableAdapter
        Dim Status_Finish_Lama As EnumAcquisitionFinish
        Dim Status_Finish_Baru As EnumAcquisitionFinish
        '----
        Dim RefNumber As String = ""
        Dim Format_Kode_Desc As String = ""
        Dim AcqFinishDate As Date = #1/1/1950#
        Dim i As Integer = 0
        Dim Primary_GJ As Double = -1
        Dim Primary_GJ_Lama As Double = -1
        Dim RefNumber_Tambahan As Double = 1
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA_InOut.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA_Journal.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA_Journal_Item.Connection.ConnectionString = DTA.Connection.ConnectionString

            If DTA_InOut.ScalarQuery_GetJmlByPrimaryAsset(vPrimary).Value <= 0 Then
                Return False
            End If
            If DTA_Journal.ScalarQuery_GetJmlByPrimaryAsset(vPrimary).Value <= 0 Then
                Return False
            End If
            'Cek Status Akun Set Finish 
            If DTA.ScalarQuery_AcquisitionFinish(vPrimary).Value <> -1 Then
                Status_Finish_Lama = DTA.ScalarQuery_AcquisitionFinish(vPrimary).Value
            Else
                Status_Finish_Lama = EnumAcquisitionFinish._false
            End If
            Status_Finish_Baru = Me.mAcquisitionFinish
            If Status_Finish_Lama <> Status_Finish_Baru Then
                If Status_Finish_Baru = EnumAcquisitionFinish._true Then
                    'Hapus Akun Set Finish
                    Primary_GJ_Lama = DTA_Journal.ScalarQuery_GetPrimaryByPrimaryAsset(vPrimary)
                    DTA_Journal.DeleteQuery_Jurnal_Primary(Primary_GJ_Lama)
                    'Simpan Akun Set Finish 
                    If DTA.ScalarQuery_GetAcquisitionType(vPrimary).Value = EnumAcquisitionType.BuyOrBuild Then
                        If DTA.ScalarQuery_GetStatusAsset(vPrimary).Value = EnumStatusAsset.On_Progress Then
                            RefNumber = DTA.ScalarQuery_GetKodeAsset(vPrimary) + "-ACQUISITION FINISH"
                            If Me.ApakahSudahAda_JournalRefNumber_SetFinish(RefNumber) = True Then
                                Dim Idx_P As Integer = RefNumber.Length
                                Dim Idx_P_Value As String = RefNumber.Substring(Idx_P - 3)
                                If IsNumeric(Idx_P_Value) = False Then
                                    RefNumber += RefNumber_Tambahan.ToString
                                Else
                                    RefNumber_Tambahan = CInt(Idx_P_Value) + 1
                                    RefNumber += RefNumber_Tambahan.ToString
                                End If
                            End If
                            If DTA.ScalarQuery_GetAcqFinishDate(vPrimary).HasValue Then
                                AcqFinishDate = DTA.ScalarQuery_GetAcqFinishDate(vPrimary).Value
                            Else
                                AcqFinishDate = #1/1/1950#
                            End If
                            Format_Kode_Desc = DTA.ScalarQuery_GetKode_Description(vPrimary)
                            DTA_Journal.InsertQuery_JurnalSetFinish(RefNumber, AcqFinishDate, 0, 0, -1, _
                             -1, Format_Kode_Desc, vUserCode, Now, vUserCode, Now, Now, -1, -1, 0, _
                             1, 0, 0, 0, 0, 7, -1, 0, vPrimary, -1)
                            Primary_GJ = DTA_Journal.ScalarQuery_GetPrimary(RefNumber)
                            '----Data item journal sepasang data baris 1 debet, data baris 2 kredit
                            For i = 0 To 1
                                If i = 0 Then
                                    DTA_Journal_Item.InsertQuery_JurnalItem_SetFinish(Primary_GJ, Me.mPrimaryAkunAsset, Me.mAcquiredValue, 0, "", _
                                    Me.mAcquiredValue, 0, 1, -1, -1)
                                Else
                                    DTA_Journal_Item.InsertQuery_JurnalItem_SetFinish(Primary_GJ, Me.mPrimaryAkunAsset, 0, Me.mAcquiredValue, "", _
                                     0, Me.mAcquiredValue, 1, -1, -1)
                                End If
                                i += 1
                            Next
                        End If
                    End If
                Else
                    If Status_Finish_Baru = EnumAcquisitionFinish._false Then
                        'Hapus Akun Set Finish
                        Primary_GJ_Lama = DTA_Journal.ScalarQuery_GetPrimaryByPrimaryAsset(vPrimary)
                        DTA_Journal.DeleteQuery_Jurnal_Primary(Primary_GJ_Lama)
                    End If
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA_Journal.Dispose()
            DTA_Journal_Item.Dispose()
            DTA_InOut.Dispose()
        End Try
    End Function

    'Fungsi yang digunakan untuk transaksi depreciation
    Public Function GetAcquiredValue_Dept(ByVal vPrimary As Double) As Decimal
        Dim DTA As New Core_Module.DseMasterAssetTableAdapters.tblmaster_asset_SetFinishTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetAcquiredValue(vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetAcquiredValue(vPrimary).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetSalvageValue_Dept(ByVal vPrimary As Double) As Decimal
        Dim DTA As New Core_Module.DseMasterAssetTableAdapters.tblmaster_asset_SetFinishTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetSalvageValue(vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetSalvageValue(vPrimary).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetSalvageValue(ByVal vPrimary As Double) As Decimal
        Dim DTA As New DseMasterAssetTableAdapters.tblmaster_assetTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetSalvageValue(vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetSalvageValue(vPrimary).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetDepreciationLife_Dept(ByVal vPrimary As Double) As Decimal
        Dim DTA As New DseMasterAssetTableAdapters.tblmaster_asset_SetFinishTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetDepreciationLife(vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetDepreciationLife(vPrimary).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetDepreciationThisAsset(ByVal vPrimary As Double) As Integer
        Dim DTA As New DseMasterAssetTableAdapters.tblmaster_asset_SetFinishTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetDepreciationLife(vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetDepreciationLife(vPrimary).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetBookValue_Dept(ByVal vPrimary As Double) As Decimal
        Dim DTA As New DseMasterAssetTableAdapters.tblmaster_asset_SetFinishTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetBookValue(vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetBookValue(vPrimary).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetDesc(ByVal vPrimaryAsset As Double) As String
        Dim DTA As New DseMasterAssetTableAdapters.tblmaster_assetTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQueryGetDescription(vPrimaryAsset)) = False Then
                Return DTA.ScalarQueryGetDescription(vPrimaryAsset)
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vCode As String) As Double
        Dim DTA As New DseMasterAssetTableAdapters.tblmaster_assetTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimary(vCode).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimary(vCode).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetStatus(ByVal vPrimaryAsset As Double) As Integer
        Dim DTA As New DseMasterAssetTableAdapters.tblmaster_assetTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetStatus(vPrimaryAsset).HasValue = True Then
                Return DTA.ScalarQuery_GetStatus(vPrimaryAsset).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryAkunOnProgress(ByVal vPrimaryAsset As Double) As Integer
        Dim DTA As New DseMasterAssetTableAdapters.tblmaster_assetTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimaryAkunOnProgress(vPrimaryAsset).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryAkunOnProgress(vPrimaryAsset).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryAkunAccumDept(ByVal vPrimaryAsset As Double) As Integer
        Dim DTA As New DseMasterAssetTableAdapters.tblmaster_assetTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimaryAccumDept(vPrimaryAsset).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryAccumDept(vPrimaryAsset).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryAkunAsset(ByVal vPrimaryAsset As Double) As Integer
        Dim DTA As New DseMasterAssetTableAdapters.DataTable1TableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimaryAkunAsset(vPrimaryAsset).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryAkunAsset(vPrimaryAsset).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetDeptDate(ByVal vPrimaryAsset As Double, ByVal vDept As Integer) As Date
        Dim DTA As New DseMasterAssetTableAdapters.tblmaster_assetTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetDepreciateDate(vPrimaryAsset, vDept)) = False Then
                Return DTA.ScalarQuery_GetDepreciateDate(vPrimaryAsset, vDept)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTypeAcquistion(ByVal vPrimary As Double) As Integer
        Dim DTA As New DseMasterAssetTableAdapters.tblmaster_assetTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetTipeAcquisition(vPrimary).HasValue = True Then
                Return DTA.ScalarQuery_GetTipeAcquisition(vPrimary).Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        End Try
    End Function

    Public Sub UpdateBookValue(ByVal vPrimaryAsset As Double)
        Dim BookValue As Decimal = 0
        Dim DTA As New DseMasterAssetTableAdapters.tblmaster_assetTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            BookValue = Me.HitungBookValue(vPrimaryAsset)
            DTA.UpdateQuery_BookValue(BookValue, vPrimaryAsset)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub UpdateAcquiredValue(ByVal vPrimaryAsset As Double)
        Dim DTA As New DseMasterAssetTableAdapters.tbltransaksi_stockinoutTableAdapter
        Dim TBL As New DseMasterAsset.tbltransaksi_stockinoutDataTable
        Dim row As DseMasterAsset.tbltransaksi_stockinoutRow
        Dim DTA1 As New DseMasterAssetTableAdapters.tbltrans_jurnal1TableAdapter
        Dim TBL1 As New DseMasterAsset.tbltrans_jurnal1DataTable
        Dim row1 As DseMasterAsset.tbltrans_jurnal1Row
        Dim TipeStockInOut As clsTransaksiStockInOut1.EnumInOutAdjust = clsTransaksiStockInOut1.EnumInOutAdjust.OutAdjustMaterialUsageAssetAcquisition
        Dim TipeJurnal2 As ClsTransaksiJournal1.EnumTipeJurnal2 = ClsTransaksiJournal1.EnumTipeJurnal2.AssetAcquistion_Other
        Dim TipeJurnal As ClsTransaksiJournal1.EnumTipeJournal = ClsTransaksiJournal1.EnumTipeJournal.GeneralJournal
        Dim MaterialUsage, Other As Decimal
        Dim DTA_Asset As New DseMasterAssetTableAdapters.tblmaster_assetTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA_Asset.Connection.ConnectionString = DTA.Connection.ConnectionString

            MaterialUsage = 0
            Other = 0

            'acquired value = material usage + other acquisition
            '1. material usage
            DTA.FillBy_PrimaryAsset(TBL, TipeStockInOut, vPrimaryAsset)
            For Each row In TBL.Rows
                MaterialUsage += row.Material_Usage_Amount
            Next
            '2. other
            DTA1.Fill(TBL1, vPrimaryAsset, TipeJurnal, TipeJurnal2)
            For Each row1 In TBL1.Rows
                Other += row1.Amount_original
            Next

            'update k database
            DTA_Asset.UpdateQuery_AcquiredValue(Other + MaterialUsage, vPrimaryAsset)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            DTA1.Dispose()
            TBL1.Dispose()
            DTA_Asset.Dispose()
        End Try
    End Sub
End Class
