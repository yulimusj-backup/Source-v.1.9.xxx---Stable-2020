Option Explicit On
Option Strict Off
Imports Core_Module

Public Class ClsMast_Diskon
    Private mPrimary As Double = -1
    Private mCode As String = ""
    Private mDesc As String = ""
    Private mUrutanDisk As Double = 0
    Private mType As EnumType = EnumType.Nominal_Percentage
    Private mMemberShip As EnumMembership = EnumMembership.Apply_For_All
    Private mStatus As EnumStatus = EnumStatus.Active
    Private mMultidiskon As Integer = 0
    Private mShowDiskonOnReceipt As Integer = 0
    Private mRemarks As String = ""
    Private mDiskonRate As String = ""
    Private mAllowToSellMoreThanMaxQty As Integer = 0
    Private mBaseQty As Decimal = 0
    Private mFreeQty As Decimal = 0
    Private mMaxQty As Decimal = 0
    Private mMinQty As Decimal = 0
    Private mMultiply As Integer = 0
    Private mProductDataValidity As EnumProductDataValidity = EnumProductDataValidity.ApplyForAllItem
    Private mBranchDataValidity As EnumBranchDataValidity = EnumBranchDataValidity.ApplyForAllBranch
    Private mPeriod As EnumPeriod = EnumPeriod.AlwaysValid
    Private mFromDate As Date = Now
    Private mToDate As Date = Now
    Private mFromTime As DateTime = "1950/01/01 00:00:00"
    Private mToTime As DateTime = "1950/01/01 00:00:00"
    Private mAdvanceFiltering As EnumAdvanceFilter = EnumAdvanceFilter.ApplyAlways
    Private mIsMonday As Integer = 0
    Private mIsTuesday As Integer = 0
    Private mIsWednesday As Integer = 0
    Private mIsThursday As Integer = 0
    Private mIsFriday As Integer = 0
    Private mIsSaturday As Integer = 0
    Private mIsSunday As Integer = 0
    Private mIs1 As Integer = 0
    Private mIs2 As Integer = 0
    Private mIs3 As Integer = 0
    Private mIs4 As Integer = 0
    Private mIs5 As Integer = 0
    Private mIs6 As Integer = 0
    Private mIs7 As Integer = 0
    Private mIs8 As Integer = 0
    Private mIs9 As Integer = 0
    Private mIs10 As Integer = 0
    Private mIs11 As Integer = 0
    Private mIs12 As Integer = 0
    Private mIs13 As Integer = 0
    Private mIs14 As Integer = 0
    Private mIs15 As Integer = 0
    Private mIs16 As Integer = 0
    Private mIs17 As Integer = 0
    Private mIs18 As Integer = 0
    Private mIs19 As Integer = 0
    Private mIs20 As Integer = 0
    Private mIs21 As Integer = 0
    Private mIs22 As Integer = 0
    Private mIs23 As Integer = 0
    Private mIs24 As Integer = 0
    Private mIs25 As Integer = 0
    Private mIs26 As Integer = 0
    Private mIs27 As Integer = 0
    Private mIs28 As Integer = 0
    Private mIs29 As Integer = 0
    Private mIs30 As Integer = 0
    Private mIs31 As Integer = 0
    Private mRounding As EnumRounding = EnumRounding.Up


    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now

    'elemen stock yang digunakan untuk antrian simpan data
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianMaster
    Dim ObjRnd As New Random
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAAntrian As New DseTransaksiAntrianTableAdapters.tbltrans_antrianTableAdapter

    Public Enum EnumRounding As Integer
        Up
        Down
        Middle
    End Enum

    Public Property Rounding() As EnumRounding
        Get
            Return mRounding
        End Get
        Set(ByVal value As EnumRounding)
            mRounding = value
        End Set
    End Property

    Public Property Is1() As Integer
        Get
            Return mIs1
        End Get
        Set(ByVal value As Integer)
            mIs1 = value
        End Set
    End Property

    Public Property Is2() As Integer
        Get
            Return mIs2
        End Get
        Set(ByVal value As Integer)
            mIs2 = value
        End Set
    End Property

    Public Property Is3() As Integer
        Get
            Return mIs3
        End Get
        Set(ByVal value As Integer)
            mIs3 = value
        End Set
    End Property

    Public Property Is4() As Integer
        Get
            Return mIs4
        End Get
        Set(ByVal value As Integer)
            mIs4 = value
        End Set
    End Property

    Public Property Is5() As Integer
        Get
            Return mIs5
        End Get
        Set(ByVal value As Integer)
            mIs5 = value
        End Set
    End Property

    Public Property Is6() As Integer
        Get
            Return mIs6
        End Get
        Set(ByVal value As Integer)
            mIs6 = value
        End Set
    End Property

    Public Property Is7() As Integer
        Get
            Return mIs7
        End Get
        Set(ByVal value As Integer)
            mIs7 = value
        End Set
    End Property

    Public Property Is8() As Integer
        Get
            Return mIs8
        End Get
        Set(ByVal value As Integer)
            mIs8 = value
        End Set
    End Property

    Public Property Is9() As Integer
        Get
            Return mIs9
        End Get
        Set(ByVal value As Integer)
            mIs9 = value
        End Set
    End Property

    Public Property Is10() As Integer
        Get
            Return mIs10
        End Get
        Set(ByVal value As Integer)
            mIs10 = value
        End Set
    End Property

    Public Property Is11() As Integer
        Get
            Return mIs11
        End Get
        Set(ByVal value As Integer)
            mIs11 = value
        End Set
    End Property

    Public Property Is12() As Integer
        Get
            Return mIs12
        End Get
        Set(ByVal value As Integer)
            mIs12 = value
        End Set
    End Property

    Public Property Is13() As Integer
        Get
            Return mIs13
        End Get
        Set(ByVal value As Integer)
            mIs13 = value
        End Set
    End Property

    Public Property Is14() As Integer
        Get
            Return mIs14
        End Get
        Set(ByVal value As Integer)
            mIs14 = value
        End Set
    End Property

    Public Property Is15() As Integer
        Get
            Return mIs15
        End Get
        Set(ByVal value As Integer)
            mIs15 = value
        End Set
    End Property

    Public Property Is16() As Integer
        Get
            Return mIs16
        End Get
        Set(ByVal value As Integer)
            mIs16 = value
        End Set
    End Property

    Public Property Is17() As Integer
        Get
            Return mIs17
        End Get
        Set(ByVal value As Integer)
            mIs17 = value
        End Set
    End Property

    Public Property Is18() As Integer
        Get
            Return mIs18
        End Get
        Set(ByVal value As Integer)
            mIs18 = value
        End Set
    End Property

    Public Property Is19() As Integer
        Get
            Return mIs19
        End Get
        Set(ByVal value As Integer)
            mIs19 = value
        End Set
    End Property

    Public Property Is20() As Integer
        Get
            Return mIs20
        End Get
        Set(ByVal value As Integer)
            mIs20 = value
        End Set
    End Property

    Public Property Is21() As Integer
        Get
            Return mIs21
        End Get
        Set(ByVal value As Integer)
            mIs21 = value
        End Set
    End Property

    Public Property Is22() As Integer
        Get
            Return mIs22
        End Get
        Set(ByVal value As Integer)
            mIs22 = value
        End Set
    End Property

    Public Property Is23() As Integer
        Get
            Return mIs23
        End Get
        Set(ByVal value As Integer)
            mIs23 = value
        End Set
    End Property

    Public Property Is24() As Integer
        Get
            Return mIs24
        End Get
        Set(ByVal value As Integer)
            mIs24 = value
        End Set
    End Property

    Public Property Is25() As Integer
        Get
            Return mIs25
        End Get
        Set(ByVal value As Integer)
            mIs25 = value
        End Set
    End Property

    Public Property Is26() As Integer
        Get
            Return mIs26
        End Get
        Set(ByVal value As Integer)
            mIs26 = value
        End Set
    End Property

    Public Property Is27() As Integer
        Get
            Return mIs27
        End Get
        Set(ByVal value As Integer)
            mIs27 = value
        End Set
    End Property

    Public Property Is28() As Integer
        Get
            Return mIs28
        End Get
        Set(ByVal value As Integer)
            mIs28 = value
        End Set
    End Property

    Public Property Is29() As Integer
        Get
            Return mIs29
        End Get
        Set(ByVal value As Integer)
            mIs29 = value
        End Set
    End Property

    Public Property Is30() As Integer
        Get
            Return mIs30
        End Get
        Set(ByVal value As Integer)
            mIs30 = value
        End Set
    End Property

    Public Property Is31() As Integer
        Get
            Return mIs31
        End Get
        Set(ByVal value As Integer)
            mIs31 = value
        End Set
    End Property

    Public Property IsMonday() As Integer
        Get
            Return mIsMonday
        End Get
        Set(ByVal value As Integer)
            mIsMonday = value
        End Set
    End Property

    Public Property IsTuesday() As Integer
        Get
            Return mIsTuesday
        End Get
        Set(ByVal value As Integer)
            mIsTuesday = value
        End Set
    End Property

    Public Property IsWednesday() As Integer
        Get
            Return mIsWednesday
        End Get
        Set(ByVal value As Integer)
            mIsWednesday = value
        End Set
    End Property

    Public Property IsThursday() As Integer
        Get
            Return mIsThursday
        End Get
        Set(ByVal value As Integer)
            mIsThursday = value
        End Set
    End Property

    Public Property IsFriday() As Integer
        Get
            Return mIsFriday
        End Get
        Set(ByVal value As Integer)
            mIsFriday = value
        End Set
    End Property

    Public Property IsSaturday() As Integer
        Get
            Return mIsSaturday
        End Get
        Set(ByVal value As Integer)
            mIsSaturday = value
        End Set
    End Property

    Public Property IsSunday() As Integer
        Get
            Return mIsSunday
        End Get
        Set(ByVal value As Integer)
            mIsSunday = value
        End Set
    End Property

    Public Enum EnumType As Integer
        Nominal_Percentage
        Gimmick
    End Enum

    Public Enum EnumMembership As Integer
        Apply_For_All
        Apply_For_Member_Only
    End Enum

    Public Enum EnumStatus As Integer
        Active
        Not_Active
    End Enum

    Public Enum EnumProductDataValidity As Integer
        ApplyForAllItem
        ApplyForSpesificItemOnly
    End Enum

    Public Enum EnumBranchDataValidity As Integer
        ApplyForAllBranch
        ApplyForSpecificBranch
    End Enum

    Public Enum EnumPeriod As Integer
        AlwaysValid
        ValidForCertainPeriod
    End Enum

    Public Enum EnumAdvanceFilter As Integer
        ApplyAlways
        ApplyOnSpecificDayOfTheWeek
        ApplyOnSpecificDateOfTheMonth
    End Enum

    Public Property CreatedBy() As String
        Get
            Return mCreatedBy
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

    Public Property Primary() As Double
        Get
            Return mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property

    Public Property Code() As String
        Get
            Return mCode
        End Get
        Set(ByVal value As String)
            mCode = value
        End Set
    End Property

    Public Property Desc() As String
        Get
            Return mDesc
        End Get
        Set(ByVal value As String)
            mDesc = value
        End Set
    End Property

    Public Property UrutanDisk() As Double
        Get
            Return mUrutanDisk
        End Get
        Set(ByVal value As Double)
            mUrutanDisk = value
        End Set
    End Property

    Public Property Type() As EnumType
        Get
            Return mType
        End Get
        Set(ByVal value As EnumType)
            mType = value
        End Set
    End Property

    Public Property Membership() As EnumMembership
        Get
            Return mMemberShip
        End Get
        Set(ByVal value As EnumMembership)
            mMemberShip = value
        End Set
    End Property

    Public Property Status() As EnumStatus
        Get
            Return mStatus
        End Get
        Set(ByVal value As EnumStatus)
            mStatus = value
        End Set
    End Property

    Public Property Multidiskon() As Integer
        Get
            Return mMultidiskon
        End Get
        Set(ByVal value As Integer)
            mMultidiskon = value
        End Set
    End Property

    Public Property ShowDiskonOnReceipt() As Integer
        Get
            Return mShowDiskonOnReceipt
        End Get
        Set(ByVal value As Integer)
            mShowDiskonOnReceipt = value
        End Set
    End Property

    Public Property Remarks() As String
        Get
            Return mRemarks
        End Get
        Set(ByVal value As String)
            mRemarks = value
        End Set
    End Property

    Public Property DiskonRate() As String
        Get
            Return mDiskonRate
        End Get
        Set(ByVal value As String)
            mDiskonRate = value
        End Set
    End Property

    Public Property AllowToSellMoreThanMaxQty() As Integer
        Get
            Return mAllowToSellMoreThanMaxQty
        End Get
        Set(ByVal value As Integer)
            mAllowToSellMoreThanMaxQty = value
        End Set
    End Property

    Public Property BaseQty() As Integer
        Get
            Return mBaseQty
        End Get
        Set(ByVal value As Integer)
            mBaseQty = value
        End Set
    End Property

    Public Property FreeQty() As Double
        Get
            Return mFreeQty
        End Get
        Set(ByVal value As Double)
            mFreeQty = value
        End Set
    End Property

    Public Property MaxQty() As Double
        Get
            Return mMaxQty
        End Get
        Set(ByVal value As Double)
            mMaxQty = value
        End Set
    End Property

    Public Property MinQty() As Double
        Get
            Return mMinQty
        End Get
        Set(ByVal value As Double)
            mMinQty = value
        End Set
    End Property

    Public Property Multiply() As Integer
        Get
            Return mMultiply
        End Get
        Set(ByVal value As Integer)
            mMultiply = value
        End Set
    End Property

    Public Property ProductDataValidity() As EnumProductDataValidity
        Get
            Return mProductDataValidity
        End Get
        Set(ByVal value As EnumProductDataValidity)
            mProductDataValidity = value
        End Set
    End Property

    Public Property BranchDataValidity() As EnumBranchDataValidity
        Get
            Return mBranchDataValidity
        End Get
        Set(ByVal value As EnumBranchDataValidity)
            mBranchDataValidity = value
        End Set
    End Property

    Public Property Period() As EnumPeriod
        Get
            Return mPeriod
        End Get
        Set(ByVal value As EnumPeriod)
            mPeriod = value
        End Set
    End Property

    Public Property FromDate() As Date
        Get
            Return mFromDate
        End Get
        Set(ByVal value As Date)
            mFromDate = value
        End Set
    End Property

    Public Property ToDate() As Date
        Get
            Return mToDate
        End Get
        Set(ByVal value As Date)
            mToDate = value
        End Set
    End Property

    Public Property FromTime() As DateTime
        Get
            Return mFromTime
        End Get
        Set(ByVal value As DateTime)
            mFromTime = value
        End Set
    End Property

    Public Property ToTime() As DateTime
        Get
            Return mToTime
        End Get
        Set(ByVal value As DateTime)
            mToTime = value
        End Set
    End Property

    Public Property AdvanceFiltering() As EnumAdvanceFilter
        Get
            Return mAdvanceFiltering
        End Get
        Set(ByVal value As EnumAdvanceFilter)
            mAdvanceFiltering = value
        End Set
    End Property

    Public Function GetConstring() As String
        Try
            Return "server=" + My.Settings.server + ";user id=" + My.Settings.user_id + ";Password=" + My.Settings.password + _
            ";persist security info=True;database=" + My.Settings.database_name + ";port=" + My.Settings.port_number + ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "server=localhost;user id=root;Password=123;persist security info=True;database=core_db;port=3307"
        End Try
    End Function

    ''Coding standarnya 
    Private Sub GetNoAntrian()
        Try
            DTAAntrian.Connection.ConnectionString = Me.GetConstring
            Me.vIdentifier = Me.ObjRnd.Next(1000000)
            'jangan sampai sama dengan identifier yang sudah ada
            While Me.DTAAntrian.ScalarQuery_GetJmlIdentifier(Me.vTipeAntrian, Me.vIdentifier).Value > 0
                Me.vIdentifier = Me.ObjRnd.Next(1000000)
            End While
            'masukan data baru ke antrian
            Me.DTAAntrian.Insert(Me.vTipeAntrian, Me.vIdentifier)

            'cek jangan sampai ada 2 data yang sama
            While Me.DTAAntrian.ScalarQuery_GetJmlIdentifier(Me.vTipeAntrian, Me.vIdentifier).Value > 1
                'bila ada yang sama maka harus masukkan data baru
                DTAAntrian.DeleteQuery(Me.vTipeAntrian, Me.vIdentifier)
                'minta identifier baru
                Me.vIdentifier = Me.ObjRnd.Next(1000000)
                'jangan sampai sama dengan identifier yang sudah ada
                While Me.DTAAntrian.ScalarQuery_GetJmlIdentifier(Me.vTipeAntrian, Me.vIdentifier).Value > 0
                    Me.vIdentifier = Me.ObjRnd.Next(1000000)
                    'masukan data baru ke antrian
                    Me.DTAAntrian.Insert(Me.vTipeAntrian, Me.vIdentifier)
                End While
            End While

            'ambil no antrian
            Me.vPrimary_Antrian = DTAAntrian.ScalarQuery_GetPrimary(Me.vIdentifier, Me.vTipeAntrian).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetCreatedBy(ByVal vPrimary As Double) As String
        Dim DTA As New DseMast_DiskonTableAdapters.tblmaster_discountTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetCreatedBy(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetCreatedBy(vPrimary)
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

    Public Function GetCode(ByVal vPrimary As Double) As String
        Dim DTA As New DseMast_DiskonTableAdapters.tblmaster_discountTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetCode(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetCode(vPrimary)
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
        Dim DTA As New DseMast_DiskonTableAdapters.tblmaster_discountTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetTimestamp(vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetTimestamp(vPrimary).Value
            Else
                Return Date.MinValue
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vCode As String) As Double
        Dim DTA As New DseMast_DiskonTableAdapters.tblmaster_discountTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetPrimary(vCode).HasValue Then
                Return DTA.ScalarQuery_GetPrimary(vCode).Value
            Else
                Return DTA.ScalarQuery_GetPrimary(vCode).Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetMultidiskon(ByVal vPrimary As Double) As Double
        Dim DTA As New DseMast_DiskonTableAdapters.tblmaster_discountTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetMultidiskon(vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetMultidiskon(vPrimary).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlByCode(ByVal vCode As String) As Double
        Dim DTA As New DseMast_DiskonTableAdapters.tblmaster_discountTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetJmlByCode(vCode).HasValue Then
                Return DTA.ScalarQuery_GetJmlByCode(vCode).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetStatus(ByVal vPrimary As Double) As Double
        Dim DTA As New DseMast_DiskonTableAdapters.tblmaster_discountTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetStatus(vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetStatus(vPrimary).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetDiscountRate(ByVal vPrimary As Double) As String
        Dim DTA As New DseMast_DiskonTableAdapters.tblmaster_discountTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetDiscountRate(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetDiscountRate(vPrimary)
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

    Public Sub UpdateUrutanDiskonSatuBandingSatu(ByVal idxLama As Integer, ByVal idxBaru As Integer, _
    ByVal DGV As DataGridView, ByVal i As FrmMast_Discount_Main)
        Dim DTA As New DseMast_DiskonTableAdapters.tblmaster_discountTableAdapter
        Dim hitung As Double = 0
        Dim Urutan As Double = -1
        Dim Primary As Double = -1
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            While hitung <= 1
                If hitung = 0 Then
                    Urutan = DGV.Item(i.ColNo.Index, idxLama).Value
                    Primary = DGV.Item(i.ColPrimary.Index, idxLama).Value
                Else
                    Urutan = DGV.Item(i.ColNo.Index, idxBaru).Value
                    Primary = DGV.Item(i.ColPrimary.Index, idxBaru).Value
                End If
                DTA.UpdateQuery_UrutanDiskon(Urutan, Primary)
                hitung += 1
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub UpdateUrutanDiskonDeleting(ByVal UrutanDiskYgDihapus As Double, ByVal Primary As Double, _
    ByVal DGV As DataGridView, ByVal i As FrmMast_Discount_Main)
        Dim DTA As New DseMast_DiskonTableAdapters.tblmaster_discountTableAdapter
        Dim TBL As New DseMast_Diskon.tblmaster_discountDataTable
        Dim Row As DseMast_Diskon.tblmaster_discountRow
        Try
            DTA.Connection.ConnectionString = Me.GetConstring

            'Filter data mana aja yang harus di update
            DTA.FillBy_PrimaryLebihBesar(TBL, Primary)

            For Each Row In TBL.Rows
                DTA.UpdateQuery_UrutanDiskon(UrutanDiskYgDihapus, Row.Primary)
                UrutanDiskYgDihapus += 1
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Public Sub LoadDataMain(ByVal vDGV As DataGridView, ByVal i As FrmMast_Discount_Main)
        Dim DTA As New DseMast_DiskonTableAdapters.tblmaster_discount1TableAdapter
        Dim TBL As New DseMast_Diskon.tblmaster_discount1DataTable
        Dim Row As DseMast_Diskon.tblmaster_discount1Row
        Dim RowDGV As DataGridViewRow
        Dim DV As Data.DataRowView
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            DTA.Fill(TBL)

            vDGV.Rows.Clear()
            TBL.DefaultView.Sort = "[Urutan diskon] ASC"
            For Each DV In TBL.DefaultView
                Row = DV.Row
                vDGV.Rows.Add()
                RowDGV = vDGV.Rows(vDGV.Rows.Count - 1)
                With RowDGV
                    .Cells(i.ColPrimary.Index).Value = Row.Primary
                    .Cells(i.ColNo.Index).Value = Row.Urutan_diskon
                    .Cells(i.ColCode.Index).Value = Row.Code
                    .Cells(i.ColDesc.Index).Value = Row.Description
                    .Cells(i.ColMultiDiscount.Index).Value = Row.Multidiskon

                End With
            Next

            'fokuskan ke baris terakhir
            'While vDGV.Rows(vDGV.Rows.Count - 1).Displayed = False _
            'Or vDGV.DisplayedRowCount(True) <> vDGV.DisplayedRowCount(False)
            'vDGV.FirstDisplayedScrollingRowIndex = vDGV.FirstDisplayedScrollingRowIndex + 1
            'End While

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Public Function ApakahSudahAda(ByVal vCode As String) As Boolean
        Dim DTA As New DseMast_DiskonTableAdapters.tblmaster_discountTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetJmlByCode(vCode).Value > 0 Then
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

    ''' <summary>
    ''' data ini digunakan untuk mengecek kebenaran data yang diberikan
    ''' </summary>
    ''' <param name="vBaru"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ApakahDataBenar(ByVal vBaru As Boolean, ByVal vPrimary As Double, _
    ByVal vShowMsg As Boolean, ByVal vObjCD As ClsCekData, _
    ByVal vEnumBhs As EnumBahasa, ByVal vDGVProduct As DataGridView, _
    ByVal vDGVBranch As DataGridView, ByVal i As FrmMast_Discount) As Boolean
        'Dim ObjClosing As New ClsTransaksiClosing
        Try
            'data salah bila:
            '1. Sudah ada no transaksi yang sama sebelumnya (untuk transaksi baru dan edit BEDA)
            '       untuk transaksi baru - dan autonumber-langsung ambil nomor baru
            '       untuk transaksi edit - dan autonumber-dianggap sama karena memang no transaksi tidak bisa diedit.
            If vBaru = True Then
                If Me.ApakahSudahAda(Me.mCode) = True Then
                    If vShowMsg = True Then
                        If vEnumBhs = EnumBahasa.Inggris Then
                            MsgBox("Duplication in code, please change your code", _
                            MsgBoxStyle.Information, "Information")
                        ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                            MsgBox("Duplikasi code, silahkan ubah code", _
                            MsgBoxStyle.Information, "Informasi")
                        End If
                    End If
                    Return False
                End If
            Else
                Dim NoTransaksiLama As String = Me.GetCode(Me.mPrimary)
                If NoTransaksiLama <> Me.mCode Then
                    If Me.ApakahSudahAda(Me.mCode) = True Then
                        If vShowMsg = True Then
                            If vEnumBhs = EnumBahasa.Inggris Then
                                MsgBox("Duplication in code, please change your code", _
                                MsgBoxStyle.Information, "Information")
                            ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                                MsgBox("Duplikasi code, silahkan ubah code", _
                                MsgBoxStyle.Information, "Informasi")
                            End If
                        End If
                        Return False
                    End If
                End If
            End If

            'yang harus diisikan:
            '1. kode customer
            If String.IsNullOrEmpty(Me.mCode) = True Then
                If vShowMsg = True Then
                    If vEnumBhs = EnumBahasa.Inggris Then
                        MsgBox("Please input code.", MsgBoxStyle.Information, "Information")
                    ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                        MsgBox("Masukan code.", MsgBoxStyle.Information, "Informasi")
                    End If
                End If
                Return False
            End If

            '2. Description
            If String.IsNullOrEmpty(Me.mDesc) = True Then
                If vShowMsg = True Then
                    If vEnumBhs = EnumBahasa.Inggris Then
                        MsgBox("Please input description.", MsgBoxStyle.Information, "Information")
                    ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                        MsgBox("Masukan description.", MsgBoxStyle.Information, "Informasi")
                    End If
                End If
                Return False
            End If

            '3. Branch Details
            If Me.mType = EnumType.Nominal_Percentage Then
                If Me.mDiskonRate = "" Then
                    If vShowMsg = True Then
                        If vEnumBhs = EnumBahasa.Inggris Then
                            MsgBox("Please input discount rate.", MsgBoxStyle.Information, "Information")
                        ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                            MsgBox("Masukan discount rate.", MsgBoxStyle.Information, "Informasi")
                        End If
                    End If
                    Return False
                End If
                If Me.mMinQty = "0" Then
                    If vShowMsg = True Then
                        If vEnumBhs = EnumBahasa.Inggris Then
                            MsgBox("Please input min qty.", MsgBoxStyle.Information, "Information")
                        ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                            MsgBox("Masukan min qty.", MsgBoxStyle.Information, "Informasi")
                        End If
                    End If
                    Return False
                End If
                If Me.mMaxQty = "0" Then
                    If vShowMsg = True Then
                        If vEnumBhs = EnumBahasa.Inggris Then
                            MsgBox("Please input max qty.", MsgBoxStyle.Information, "Information")
                        ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                            MsgBox("Masukan max qty.", MsgBoxStyle.Information, "Informasi")
                        End If
                    End If
                    Return False
                End If
            ElseIf Me.mType = EnumType.Gimmick Then
                If Me.mMinQty = "0" Then
                    If vShowMsg = True Then
                        If vEnumBhs = EnumBahasa.Inggris Then
                            MsgBox("Please input min qty.", MsgBoxStyle.Information, "Information")
                        ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                            MsgBox("Masukan min qty.", MsgBoxStyle.Information, "Informasi")
                        End If
                    End If
                    Return False
                End If
                If Me.mMaxQty = "0" Then
                    If vShowMsg = True Then
                        If vEnumBhs = EnumBahasa.Inggris Then
                            MsgBox("Please input max qty.", MsgBoxStyle.Information, "Information")
                        ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                            MsgBox("Masukan max qty.", MsgBoxStyle.Information, "Informasi")
                        End If
                    End If
                    Return False
                End If
                If Me.mBaseQty = "0" Then
                    If vShowMsg = True Then
                        If vEnumBhs = EnumBahasa.Inggris Then
                            MsgBox("Please input base qty.", MsgBoxStyle.Information, "Information")
                        ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                            MsgBox("Masukan base qty.", MsgBoxStyle.Information, "Informasi")
                        End If
                    End If
                    Return False
                End If
                If Me.mFreeQty = "0" Then
                    If vShowMsg = True Then
                        If vEnumBhs = EnumBahasa.Inggris Then
                            MsgBox("Please input free qty.", MsgBoxStyle.Information, "Information")
                        ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                            MsgBox("Masukan free qty.", MsgBoxStyle.Information, "Informasi")
                        End If
                    End If
                    Return False
                End If
            End If
            If Me.mProductDataValidity = EnumProductDataValidity.ApplyForSpesificItemOnly Then
                If vDGVProduct.Rows.Count <= 0 Then
                    If vShowMsg = True Then
                        If vEnumBhs = EnumBahasa.Inggris Then
                            MsgBox("Please input product data.", MsgBoxStyle.Information, "Information")
                        ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                            MsgBox("Masukan product data.", MsgBoxStyle.Information, "Informasi")
                        End If
                    End If
                    Return False
                End If
            End If
            If Me.mBranchDataValidity = EnumBranchDataValidity.ApplyForSpecificBranch Then
                If vDGVBranch.Rows.Count <= 0 Then
                    If vShowMsg = True Then
                        If vEnumBhs = EnumBahasa.Inggris Then
                            MsgBox("Please input branch data.", MsgBoxStyle.Information, "Information")
                        ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                            MsgBox("Masukan server branch data.", MsgBoxStyle.Information, "Informasi")
                        End If
                    End If
                    Return False
                End If
            End If

            'Tanggal Transaksi berada dalam Periode yang sudah closing 
            'If ObjClosing.GetClosingDate.Date >= Me.TanggalTransaksi Then
            'MsgBox("Periode <= " + Format(ObjClosing.GetClosingDate.Date, "dd MMM yyyy") _
            '+ " sudah di kunci. Hubungi admin Anda.", MsgBoxStyle.Information, "Informasi")
            'Return False
            'End If

            'kalo mpe sini ok
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            'ObjClosing = Nothing
        End Try
    End Function

    Public Function ApakahBolehEdit(ByVal vPrimary As Double, ByVal vUserCode As String, _
    ByVal vStrictEdit As Boolean, ByVal vShowMsg As Boolean, ByVal vEnumBhs As EnumBahasa) As Boolean
        Dim DTA As New DseMast_DiskonTableAdapters.tblmaster_discountTableAdapter
        Dim pCreatedBy As String = ""
        'Dim ObjClosing As New ClsTransaksiClosing
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            Return True
            Exit Try

            '1. tidak boleh hapus jika melanggar aturan strict edit
            If vStrictEdit = True Then
                If IsNothing(DTA.ScalarQuery_GetCreatedBy(vPrimary)) = False Then
                    pCreatedBy = DTA.ScalarQuery_GetCreatedBy(vPrimary)
                Else
                    pCreatedBy = ""
                End If

                If pCreatedBy <> vUserCode Then
                    If vShowMsg = True Then
                        If vEnumBhs = EnumBahasa.Inggris Then
                            MsgBox("User can edit this data: " + pCreatedBy, _
                            MsgBoxStyle.Information, "Information")
                        ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                            MsgBox("Data ini hanya dapat diedit oleh: " + pCreatedBy, _
                            MsgBoxStyle.Information, "Informasi")
                        End If
                    End If
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            'ObjClosing = Nothing
        End Try
    End Function

    Public Function ApakahBolehHapus(ByVal vPrimary As Double, ByVal vUserCode As String, _
    ByVal vStrictDelete As Boolean, ByVal vShowMsg As Boolean, ByVal vEnumBhs As EnumBahasa, _
    ByVal i As FrmMast_Discount) As Boolean
        Dim DTA As New DseMast_DiskonTableAdapters.tblmaster_discountTableAdapter
        Dim pCreatedBy As String = ""
        Dim ObjPOSItem As New ClsTrans_POS_Item
        Dim ObjPOS As New ClsTrans_POS
        'Dim ObjClosing As New ClsTransaksiClosing
        Try
            DTA.Connection.ConnectionString = Me.GetConstring

            '1. tidak boleh hapus jika melanggar aturan strict edit
            If vStrictDelete = True Then
                If IsNothing(DTA.ScalarQuery_GetCreatedBy(vPrimary)) = False Then
                    pCreatedBy = DTA.ScalarQuery_GetCreatedBy(vPrimary)
                Else
                    pCreatedBy = ""
                End If
                If pCreatedBy <> vUserCode Then
                    If vShowMsg = True Then
                        If vEnumBhs = EnumBahasa.Inggris Then
                            MsgBox("User can delete this data: " + pCreatedBy, _
                            MsgBoxStyle.Information, "Information")
                        ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                            MsgBox("Data ini hanya dapat didelete oleh: " + pCreatedBy, _
                            MsgBoxStyle.Information, "Informasi")
                        End If
                    End If
                    Return False
                End If
            End If

            '2. Trans POS
            If ObjPOS.GetJmlByPrimaryDiskon(vPrimary) > 0 Then
                If vEnumBhs = EnumBahasa.Inggris Then
                    MsgBox("Data cannot be removed because there are other transactions that refer", MsgBoxStyle.Information, "Information")
                ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                    MsgBox("Data tidak dapat dihapus karena ada transaksi lain yang merujuk", MsgBoxStyle.Information, "Informasi")
                End If
                Return False
            End If
            If ObjPOSItem.GetJmlByPrimaryDisc1(vPrimary) > 0 Then
                If vEnumBhs = EnumBahasa.Inggris Then
                    MsgBox("Data cannot be deleted because there are other transaction that refering this data", MsgBoxStyle.Information, "Information")
                ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                    MsgBox("Data tidak dapat dihapus karena ada transaksi lain yang merujuk", MsgBoxStyle.Information, "Informasi")
                End If
                Return False
            End If
            If ObjPOSItem.GetJmlByPrimaryDisc2(vPrimary) > 0 Then
                If vEnumBhs = EnumBahasa.Inggris Then
                    MsgBox("Data cannot be deleted because there are other transaction that refering this data", MsgBoxStyle.Information, "Information")
                ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                    MsgBox("Data tidak dapat dihapus karena ada transaksi lain yang merujuk", MsgBoxStyle.Information, "Informasi")
                End If
                Return False
            End If
            If ObjPOSItem.GetJmlByPrimaryDisc3(vPrimary) > 0 Then
                If vEnumBhs = EnumBahasa.Inggris Then
                    MsgBox("Data cannot be deleted because there are other transaction that refering this data", MsgBoxStyle.Information, "Information")
                ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                    MsgBox("Data tidak dapat dihapus karena ada transaksi lain yang merujuk", MsgBoxStyle.Information, "Informasi")
                End If
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjPOSItem = Nothing
            ObjPOS = Nothing
            'ObjClosing = Nothing
        End Try
    End Function

    Public Function CariDetail(ByVal vPrimary As Double, _
    ByVal vObjCD As ClsCekData, ByVal vDGVProd As DataGridView, ByVal vDGVBranch As DataGridView, _
    ByVal i As FrmMast_Discount) As Boolean
        Dim DTA As New DseMast_DiskonTableAdapters.tblmaster_discount1TableAdapter
        Dim TBL As New DseMast_Diskon.tblmaster_discount1DataTable
        Dim row As DseMast_Diskon.tblmaster_discount1Row
        Dim ObjItem As New ClsMast_Diskon_Item
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            DTA.FillBy_Primary(TBL, vPrimary)

            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = .Primary
                    Me.mCode = .Code
                    Me.mDesc = .Description
                    Me.mUrutanDisk = .Urutan_diskon
                    Me.mType = .Type
                    Me.mMemberShip = .Membership
                    Me.mStatus = .Status
                    Me.mMultidiskon = .Multidiskon
                    Me.mShowDiskonOnReceipt = .Show_Discount_On_Receipt
                    Me.mDiskonRate = .Discount_Rate
                    Me.mRemarks = .Remarks
                    Me.mAllowToSellMoreThanMaxQty = .Allow_To_Sell_More_Than_Max_Qty
                    Me.mBaseQty = .Base_Qty
                    Me.mFreeQty = .Free_Qty
                    Me.mMinQty = .Min_Qty
                    Me.MaxQty = .Max_Qty
                    Me.mMultiply = .Multiply
                    Me.mProductDataValidity = .Product_Data_Validity
                    Me.mBranchDataValidity = .Branch_Data_Validity
                    Me.mPeriod = .Period
                    Me.mFromDate = .From_Date
                    Me.mToDate = .From_To
                    Me.FromTime = .From_Time
                    Me.mToTime = .To_Time
                    Me.AdvanceFiltering = .Advance_Filtering
                    Me.mIsMonday = .Is_Monday
                    Me.mIsTuesday = .Is_Tuesday
                    Me.mIsWednesday = .Is_Wednesday
                    Me.mIsThursday = .Is_Thursday
                    Me.mIsFriday = .Is_Friday
                    Me.mIsSaturday = .Is_Saturday
                    Me.mIsSunday = .Is_Sunday
                    Me.mIs1 = .Is_1
                    Me.mIs2 = .Is_2
                    Me.mIs3 = .Is_3
                    Me.mIs4 = .Is_4
                    Me.mIs5 = .Is_5
                    Me.mIs6 = .Is_6
                    Me.mIs7 = .Is_7
                    Me.mIs8 = .Is_8
                    Me.mIs9 = .Is_9
                    Me.mIs10 = .Is_10
                    Me.mIs11 = .Is_11
                    Me.mIs12 = .Is_12
                    Me.mIs13 = .Is_13
                    Me.mIs14 = .Is_14
                    Me.mIs15 = .Is_15
                    Me.mIs16 = .Is_16
                    Me.mIs17 = .Is_17
                    Me.mIs18 = .Is_18
                    Me.mIs19 = .Is_19
                    Me.mIs20 = .Is_20
                    Me.mIs21 = .Is_21
                    Me.mIs22 = .Is_22
                    Me.mIs23 = .Is_23
                    Me.mIs24 = .Is_24
                    Me.mIs25 = .Is_25
                    Me.mIs26 = .Is_26
                    Me.mIs27 = .Is_27
                    Me.mIs28 = .Is_28
                    Me.mIs29 = .Is_29
                    Me.mIs30 = .Is_30
                    Me.mIs31 = .Is_31
                    Me.mRounding = .Rounding

                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
                End With
            End If

            ObjItem.IsiDGV(vPrimary, vDGVProd, vDGVBranch, i, vObjCD)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Sub New()
        Try
            Me.DTAAntrian.Connection.ConnectionString = Me.GetConstring
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Overridable Function SimpanBaru(ByVal vUserCode As String, _
    ByVal vObjCD As ClsCekData, _
    ByVal vPrimary As Double, ByVal vEnumBhs As EnumBahasa, _
    ByVal vDGVProd As DataGridView, ByVal vDGVBranch As DataGridView, _
    ByVal i As FrmMast_Discount) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Try
            'ambil no antrian
            Me.GetNoAntrian()

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If Me.DTAAntrian.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                        'clock antri matikan
                        Me.Clock_Antri.Stop()
                        'sudah saatnya melakukan proses simpan baru
                        retval = Me.SimpanBaru1(vUserCode, vPrimary, vObjCD, _
                        vEnumBhs, vDGVProd, vDGVBranch, i)

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
                    End If
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function SimpanBaru1(ByVal vUserCode As String, _
    ByVal vPrimary As Double, ByVal vObjCD As ClsCekData, _
    ByVal vEnumBhs As EnumBahasa, ByVal vDGVProd As DataGridView, _
    ByVal vDGVBranch As DataGridView, ByVal i As FrmMast_Discount) As Boolean
        Dim DTA1 As New DseMast_DiskonTableAdapters.tblmaster_discount1TableAdapter
        Dim DTA As New DseMast_DiskonTableAdapters.tblmaster_discountTableAdapter
        Dim ObjItem As New ClsMast_Diskon_Item
        Try
            DTA1.Connection.ConnectionString = Me.GetConstring
            DTA.Connection.ConnectionString = Me.GetConstring

            If Me.ApakahDataBenar(True, vPrimary, True, vObjCD, vEnumBhs, vDGVProd, vDGVBranch, i) = True Then
                If DTA1.Insert(Me.mCode, Me.mDesc, Me.mUrutanDisk, Me.mType, Me.mMemberShip, _
                Me.mStatus, Me.mMultidiskon, Me.mShowDiskonOnReceipt, Me.mRemarks, Me.mDiskonRate, _
                Me.mAllowToSellMoreThanMaxQty, Me.mBaseQty, Me.mFreeQty, Me.mMaxQty, Me.mMinQty, _
                Me.mMultiply, Me.mProductDataValidity, Me.mBranchDataValidity, Me.mPeriod, _
                Me.mFromDate, Me.mToDate, Me.mFromTime, Me.mToTime, Me.mAdvanceFiltering, _
                vUserCode, Now, vUserCode, Now, Now, _
                Me.mIsMonday, Me.mIsTuesday, Me.mIsWednesday, Me.mIsThursday, Me.mIsFriday, _
                Me.mIsSaturday, Me.mIsSunday, Me.mIs1, Me.mIs2, Me.mIs3, Me.mIs4, Me.mIs5, Me.mIs6, _
                Me.mIs7, Me.mIs8, Me.mIs9, Me.mIs10, Me.mIs11, Me.mIs12, Me.mIs13, Me.mIs14, Me.mIs15, _
                Me.mIs16, Me.mIs17, Me.mIs18, Me.mIs19, Me.mIs20, Me.mIs21, Me.mIs22, Me.mIs23, _
                Me.mIs24, Me.mIs25, Me.mIs26, Me.mIs27, Me.mIs28, Me.mIs29, Me.mIs30, Me.mIs31, Me.mRounding) > 0 Then
                    Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mCode)
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    'simpan data item
                    ObjItem.SimpanItem(Me.mPrimary, vDGVProd, vDGVBranch, i)

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
            DTA1.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Overridable Function SimpanPerubahan(ByVal vUserCode As String, _
    ByVal vObjCD As ClsCekData, ByVal vPrimary As Double, _
    ByVal vEnumBhs As EnumBahasa, ByVal vDGVProd As DataGridView, _
    ByVal vDGVBranch As DataGridView, ByVal i As FrmMast_Discount) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Try
            'ambil no antrian
            Me.GetNoAntrian()

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If Me.DTAAntrian.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                        'clock antri matikan
                        Me.Clock_Antri.Stop()
                        'sudah saatnya melakukan proses simpan baru
                        retval = Me.SimpanPerubahan1(vUserCode, vPrimary, vObjCD, vEnumBhs, vDGVProd, vDGVBranch, i)

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
                    End If
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function SimpanPerubahan1(ByVal vUserCode As String, _
    ByVal vPrimary As Double, ByVal vObjCD As ClsCekData, _
    ByVal vEnumBhs As EnumBahasa, ByVal vDGVProd As DataGridView, _
    ByVal vDGVBranch As DataGridView, ByVal i As FrmMast_Discount) As Boolean
        Dim DTA1 As New DseMast_DiskonTableAdapters.tblmaster_discount1TableAdapter
        Dim DTA As New DseMast_DiskonTableAdapters.tblmaster_discountTableAdapter
        Dim ObjItem As New ClsMast_Diskon_Item
        Try
            DTA1.Connection.ConnectionString = Me.GetConstring
            DTA.Connection.ConnectionString = Me.GetConstring

            If Me.ApakahDataBenar(False, vPrimary, True, vObjCD, vEnumBhs, vDGVProd, vDGVBranch, i) = True Then
                If DTA1.Update(Me.mCode, Me.mDesc, Me.mUrutanDisk, Me.mType, Me.mMemberShip, _
                    Me.mStatus, Me.mMultidiskon, Me.mShowDiskonOnReceipt, Me.mRemarks, Me.mDiskonRate, _
                    Me.mAllowToSellMoreThanMaxQty, Me.mBaseQty, Me.mFreeQty, Me.mMaxQty, Me.mMinQty, _
                    Me.mMultiply, Me.mProductDataValidity, Me.mBranchDataValidity, Me.mPeriod, _
                    Me.mFromDate, Me.mToDate, Me.mFromTime, Me.mToTime, Me.mAdvanceFiltering, _
                    vUserCode, Now, Now, _
                    Me.mIsMonday, Me.mIsTuesday, Me.mIsWednesday, Me.mIsThursday, Me.mIsFriday, _
                    Me.mIsSaturday, Me.mIsSunday, Me.mIs1, Me.mIs2, Me.mIs3, Me.mIs4, Me.mIs5, Me.mIs6, _
                    Me.mIs7, Me.mIs8, Me.mIs9, Me.mIs10, Me.mIs11, Me.mIs12, Me.mIs13, Me.mIs14, Me.mIs15, _
                    Me.mIs16, Me.mIs17, Me.mIs18, Me.mIs19, Me.mIs20, Me.mIs21, Me.mIs22, Me.mIs23, _
                    Me.mIs24, Me.mIs25, Me.mIs26, Me.mIs27, Me.mIs28, Me.mIs29, Me.mIs30, Me.mIs31, Me.mRounding, _
                    Me.mPrimary, Me.mTimestamp) > 0 Then

                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode

                    'hapus data item lama
                    ObjItem.HapusItem(Me.mPrimary)

                    'simpan data item
                    ObjItem.SimpanItem(Me.mPrimary, vDGVProd, vDGVBranch, i)

                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Overridable Function Hapus(ByVal vPrimary As Double, ByVal vUserCode As String, _
    ByVal vStrictDelete As Boolean, ByVal vShowMsg As Boolean, _
    ByVal vEnumBhs As EnumBahasa, ByVal i As FrmMast_Discount) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Try
            'ambil no antrian
            Me.GetNoAntrian()

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If Me.DTAAntrian.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                        'clock antri matikan
                        Me.Clock_Antri.Stop()
                        'sudah saatnya melakukan proses simpan baru
                        retval = Me.Hapus1(vPrimary, vUserCode, _
                        vStrictDelete, vShowMsg, vEnumBhs, i)
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
                    End If
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function Hapus1(ByVal vPrimary As Double, ByVal vUserCode As String, _
    ByVal vStrictDelete As Boolean, ByVal vShowMsg As Boolean, _
    ByVal vEnumBhs As EnumBahasa, ByVal i As FrmMast_Discount) As Boolean
        Dim DTA As New DseMast_DiskonTableAdapters.tblmaster_discount1TableAdapter
        Dim ObjItem As New ClsMast_Diskon_Item
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If Me.ApakahBolehHapus(vPrimary, vUserCode, _
                vStrictDelete, vShowMsg, vEnumBhs, i) = True Then
                'Hapus item
                ObjItem.HapusItem(vPrimary)
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

    Public Sub ResetObject(ByVal vDGVProd As DataGridView, ByVal vDGVBranch As DataGridView)
        Try
            Me.mPrimary = -1
            Me.mCode = ""
            Me.mDesc = ""
            Me.mUrutanDisk = 0
            Me.mType = EnumType.Nominal_Percentage
            Me.mMemberShip = EnumMembership.Apply_For_All
            Me.mStatus = EnumStatus.Active
            Me.mMultidiskon = 0
            Me.mShowDiskonOnReceipt = 0
            Me.mDiskonRate = ""
            Me.mRemarks = ""
            Me.mAllowToSellMoreThanMaxQty = 0
            Me.mBaseQty = 0
            Me.mFreeQty = 0
            Me.mMinQty = 0
            Me.MaxQty = 0
            Me.mMultiply = 0
            Me.mProductDataValidity = EnumProductDataValidity.ApplyForAllItem
            Me.mBranchDataValidity = EnumBranchDataValidity.ApplyForAllBranch
            Me.mPeriod = EnumPeriod.AlwaysValid
            Me.mFromDate = "1950/01/01"
            Me.mToDate = "1950/01/01"
            Me.mFromTime = "1950/01/01 00:00:00"
            Me.mToTime = "1950/01/01 00:00:00"
            Me.AdvanceFiltering = EnumAdvanceFilter.ApplyAlways
            Me.mIsMonday = 0
            Me.mIsTuesday = 0
            Me.mIsWednesday = 0
            Me.mIsThursday = 0
            Me.mIsFriday = 0
            Me.mIsSaturday = 0
            Me.mIsSunday = 0
            Me.mIs1 = 0
            Me.mIs2 = 0
            Me.mIs3 = 0
            Me.mIs4 = 0
            Me.mIs5 = 0
            Me.mIs6 = 0
            Me.mIs7 = 0
            Me.mIs8 = 0
            Me.mIs9 = 0
            Me.mIs10 = 0
            Me.mIs11 = 0
            Me.mIs12 = 0
            Me.mIs13 = 0
            Me.mIs14 = 0
            Me.mIs15 = 0
            Me.mIs16 = 0
            Me.mIs17 = 0
            Me.mIs18 = 0
            Me.mIs19 = 0
            Me.mIs20 = 0
            Me.mIs21 = 0
            Me.mIs22 = 0
            Me.mIs23 = 0
            Me.mIs24 = 0
            Me.mIs25 = 0
            Me.mIs26 = 0
            Me.mIs27 = 0
            Me.mIs28 = 0
            Me.mIs29 = 0
            Me.mIs30 = 0
            Me.mIs31 = 0

            Me.mCreatedBy = ""
            Me.mCreatedDate = Today
            Me.mModifiedBy = ""
            Me.mModifiedDate = Today
            Me.mTimestamp = Today

            vDGVProd.Rows.Clear()
            vDGVBranch.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub FungsiInputProductCode(ByVal vDGVItemProd As DataGridView, ByVal DTA As DseMast_DiskonTableAdapters.DataTable4TableAdapter, _
    ByVal vPrimaryGroup As Double, ByVal vPrimaryCategory As Double, ByVal vIndexCBO As Integer, _
    ByVal vObjCD As ClsCekData, ByVal i As FrmMast_Discount)
        Dim TBL As New DseMast_Diskon.DataTable4DataTable
        Dim Row As DseMast_Diskon.DataTable4Row
        Dim RowDGV As DataGridViewRow
        Dim Filter As String = ""
        Dim DV As Data.DataRowView
        Dim ListPrimary As New List(Of Double)
        'Level Group Product
        '1. Sub Brand, 2. Brand, 3. Product Group
        Dim ProductGroupLevel As Double = -1
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            vDGVItemProd.AllowUserToAddRows = False
            If vDGVItemProd.RowCount > 0 Then
                For Each RowDGV In vDGVItemProd.Rows
                    If ListPrimary.Contains(RowDGV.Cells(i.ColPrimaryProduct.Index).Value) = False Then
                        ListPrimary.Add(RowDGV.Cells(i.ColPrimaryProduct.Index).Value)
                    End If
                Next
            End If

            DTA.Fill(TBL)
            If vIndexCBO = 0 Then
                'Product Group
                'ProductGroupLevel = 3
                'Filter = "[Primary Group] = " + vPrimaryGroup.ToString + " AND [Group Level] = " + ProductGroupLevel.ToString + ""
                Filter = "[Primary Product Group] = " + vPrimaryGroup.ToString + ""
                TBL.DefaultView.RowFilter = Filter
                If TBL.DefaultView.Count > 0 Then
                    For Each DV In TBL.DefaultView
                        Row = DV.Row
                        If ListPrimary.Contains(Row.Primary) = False Then
                            vDGVItemProd.Rows.Add()
                            RowDGV = vDGVItemProd.Rows(vDGVItemProd.Rows.Count - 1)
                            With RowDGV
                                .Cells(i.ColCodeNo.Index).Value = Row.Product_Code
                                .Cells(i.ColDesc.Index).Value = Row.Description
                                .Cells(i.ColRemarks.Index).Value = ""
                                .Cells(i.ColPrimaryProduct.Index).Value = Row.Primary
                            End With
                            ListPrimary.Add(Row.Primary)
                        End If
                    Next
                End If
            ElseIf vIndexCBO = 1 Then
                'Brand
                'ProductGroupLevel = 2
                'Filter = "[Primary Group] = " + vPrimaryGroup.ToString + " AND [Group Level] = " + ProductGroupLevel.ToString + ""
                Filter = "[Primary Brand] = " + vPrimaryGroup.ToString + ""
                TBL.DefaultView.RowFilter = Filter
                If TBL.DefaultView.Count > 0 Then
                    For Each DV In TBL.DefaultView
                        Row = DV.Row
                        If ListPrimary.Contains(Row.Primary) = False Then
                            vDGVItemProd.Rows.Add()
                            RowDGV = vDGVItemProd.Rows(vDGVItemProd.Rows.Count - 1)
                            With RowDGV
                                .Cells(i.ColCodeNo.Index).Value = Row.Product_Code
                                .Cells(i.ColDesc.Index).Value = Row.Description
                                .Cells(i.ColRemarks.Index).Value = ""
                                .Cells(i.ColPrimaryProduct.Index).Value = Row.Primary
                            End With
                            ListPrimary.Add(Row.Primary)
                        End If
                    Next
                End If
            ElseIf vIndexCBO = 2 Then
                'Sub brand
                'ProductGroupLevel = 1
                'Filter = "[Primary Group] = " + vPrimaryGroup.ToString + " AND [Group Level] = " + ProductGroupLevel.ToString + ""
                Filter = "[Primary Sub Brand] = " + vPrimaryGroup.ToString + ""
                TBL.DefaultView.RowFilter = Filter
                If TBL.DefaultView.Count > 0 Then
                    For Each DV In TBL.DefaultView
                        Row = DV.Row
                        If ListPrimary.Contains(Row.Primary) = False Then
                            vDGVItemProd.Rows.Add()
                            RowDGV = vDGVItemProd.Rows(vDGVItemProd.Rows.Count - 1)
                            With RowDGV
                                .Cells(i.ColCodeNo.Index).Value = Row.Product_Code
                                .Cells(i.ColDesc.Index).Value = Row.Description
                                .Cells(i.ColRemarks.Index).Value = ""
                                .Cells(i.ColPrimaryProduct.Index).Value = Row.Primary
                            End With
                            ListPrimary.Add(Row.Primary)
                        End If
                    Next
                End If
            ElseIf vIndexCBO = 3 Then
                'category
                'ProductGroupLevel = -1
                Filter = "[Primary Category] = " + vPrimaryCategory.ToString + ""
                TBL.DefaultView.RowFilter = Filter
                If TBL.DefaultView.Count > 0 Then
                    For Each DV In TBL.DefaultView
                        Row = DV.Row
                        If ListPrimary.Contains(Row.Primary) = False Then
                            vDGVItemProd.Rows.Add()
                            RowDGV = vDGVItemProd.Rows(vDGVItemProd.Rows.Count - 1)
                            With RowDGV
                                .Cells(i.ColCodeNo.Index).Value = Row.Product_Code
                                .Cells(i.ColDesc.Index).Value = Row.Description
                                .Cells(i.ColRemarks.Index).Value = ""
                                .Cells(i.ColPrimaryProduct.Index).Value = Row.Primary
                            End With
                            ListPrimary.Add(Row.Primary)
                        End If
                    Next
                End If
            End If

            vObjCD.SetNumber(vDGVItemProd)
            vDGVItemProd.AllowUserToAddRows = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TBL.Dispose()
            ListPrimary = Nothing
        End Try
    End Sub

    Public Function GetMAXUrutan() As Double
        Dim DTA As New DseMast_DiskonTableAdapters.tblmaster_discountTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetMaxUrutan().HasValue Then
                Return DTA.ScalarQuery_GetMaxUrutan.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub UpdateMultidiskon(ByVal vPrimary As Double, ByVal vValueMultidiskon As Integer)
        Dim DTA As New DseMast_DiskonTableAdapters.tblmaster_discountTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            DTA.UpdateQuery_Multidiskon(vValueMultidiskon, vPrimary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    ''' <summary>
    ''' menghitung item product disc, untuk primary discount yang diberikan
    ''' </summary>
    ''' <param name="vPrimaryDisc"></param>
    ''' <param name="vPrimaryProduct"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetJml_ByProduct(ByVal vPrimaryDisc As Double, ByVal vPrimaryProduct As Double) As Double
        Dim DTA As New DseMast_DiskonTableAdapters.tblmaster_discount_itemproductTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetJmlByPrimaryDiskProd(vPrimaryDisc, vPrimaryProduct).HasValue Then
                Return DTA.ScalarQuery_GetJmlByPrimaryDiskProd(vPrimaryDisc, vPrimaryProduct).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Function GetJml_ByBranch(ByVal vPrimaryDisc As Double, ByVal vPrimaryBranch As Double) As Double
        Dim DTA As New DseMast_DiskonTableAdapters.tblmaster_discount_itembranchTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetJmlByPrimaryDiskBranch(vPrimaryDisc, vPrimaryBranch).HasValue Then
                Return DTA.ScalarQuery_GetJmlByPrimaryDiskBranch(vPrimaryDisc, vPrimaryBranch).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="vPrimaryBranch"></param>
    ''' <param name="vPrimaryProduct"></param>
    ''' <param name="vPrimaryMember"></param>
    ''' <param name="vDateTime"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Get_Disc_String(ByVal vPrimaryBranch As Double, ByVal vPrimaryProduct As Double, _
    ByVal vPrimaryMember As Double, ByVal vDateTime As Date, ByVal vQty As Decimal, _
    ByVal TBL As DseMast_Diskon.tblmaster_discount2DataTable, ByRef vMaxQty As Decimal) As String
        Dim ind As Integer = 0
        Dim DTA As New DseMast_DiskonTableAdapters.tblmaster_discount2TableAdapter
        Dim row As DseMast_Diskon.tblmaster_discount2Row
        Dim AllMember As EnumMembership = EnumMembership.Apply_For_All
        Dim AllProduct As EnumProductDataValidity = EnumProductDataValidity.ApplyForAllItem
        Dim AllPeriod As EnumPeriod = EnumPeriod.AlwaysValid
        Dim AdvanceTime As EnumAdvanceFilter = EnumAdvanceFilter.ApplyAlways
        Dim AllBranch As EnumBranchDataValidity = EnumBranchDataValidity.ApplyForAllBranch
        Dim Status As EnumStatus = EnumStatus.Active
        Dim Cek As Boolean = True
        Dim JmlMultiDiskon As Integer = 0
        Dim Retval As String = ""
        Try
            '1. isikan tabel, berdasarkan urutan diskon
            DTA.Connection.ConnectionString = Me.GetConstring
            TBL.Rows.Clear()
            DTA.Fill(TBL)

            '2. buang baris yang tidak valid
            'tidak valid jika:
            While ind <= TBL.Rows.Count - 1
                row = TBL.Rows(ind)
                Cek = True
                '   a. member tidak cocok
                AllMember = row.All_Member
                If AllMember <> EnumMembership.Apply_For_All Then
                    If vPrimaryMember = -1 Then
                        TBL.Rows.RemoveAt(ind)
                        Cek = False
                    End If
                End If
                '   b. product tidak cocok
                If Cek = True Then
                    AllProduct = row.All_Product
                    If AllProduct <> EnumProductDataValidity.ApplyForAllItem Then
                        If Me.GetJml_ByProduct(row.Primary, vPrimaryProduct) = 0 Then
                            TBL.Rows.RemoveAt(ind)
                            Cek = False
                        End If
                    End If
                End If
                '   c. cabang tidak cocok
                If Cek = True Then
                    AllBranch = row.All_Branch
                    If AllBranch <> EnumBranchDataValidity.ApplyForAllBranch Then
                        If Me.GetJml_ByBranch(row.Primary, vPrimaryBranch) = 0 Then
                            TBL.Rows.RemoveAt(ind)
                            Cek = False
                        End If
                    End If
                End If
                '   d. jika diskon tidak aktif
                If Cek = True Then
                    Status = row.Status
                    If Status <> EnumStatus.Active Then
                        TBL.Rows.RemoveAt(ind)
                        Cek = False
                    End If
                End If
                '   e. qty tidak cocok, kurang dari min qty, lebih dari max
                If Cek = True Then
                    If vQty > row.Max_Qty Then
                        TBL.Rows.RemoveAt(ind)
                        Cek = False
                    End If
                End If
                If Cek = True Then
                    If vQty < row.Min_Qty Then
                        TBL.Rows.RemoveAt(ind)
                        Cek = False
                    End If
                End If
                '   f. time tidak cocok
                If Cek = True Then
                    AllPeriod = row.All_Period
                    AdvanceTime = row.Advance_Filtering
                    If AllPeriod <> EnumPeriod.AlwaysValid Then
                        If AdvanceTime = EnumAdvanceFilter.ApplyAlways Then
                            If vDateTime >= row.From_Date Then
                                If vDateTime <= row.To_Time Then
                                    If Format(vDateTime, "HH:mm") >= Format(row.From_Time, "HH:mm") Then
                                        If Format(vDateTime, "HH:mm") <= Format(row.To_Time, "HH:mm") Then
                                        Else
                                            TBL.Rows.RemoveAt(ind)
                                            Cek = False
                                        End If
                                    Else
                                        TBL.Rows.RemoveAt(ind)
                                        Cek = False
                                    End If
                                Else
                                    TBL.Rows.RemoveAt(ind)
                                    Cek = False
                                End If
                            Else
                                TBL.Rows.RemoveAt(ind)
                                Cek = False
                            End If
                        ElseIf AdvanceTime = EnumAdvanceFilter.ApplyOnSpecificDayOfTheWeek Then
                            If vDateTime.Day = 1 Then
                                If row.Is_Monday = 0 Then
                                    TBL.Rows.RemoveAt(ind)
                                    Cek = False
                                End If
                            ElseIf vDateTime.Day = 2 Then
                                If row.Is_Tuesday = 0 Then
                                    TBL.Rows.RemoveAt(ind)
                                    Cek = False
                                End If
                            ElseIf vDateTime.Day = 3 Then
                                If row.Is_Wednesday = 0 Then
                                    TBL.Rows.RemoveAt(ind)
                                    Cek = False
                                End If
                            ElseIf vDateTime.Day = 4 Then
                                If row.Is_Thursday = 0 Then
                                    TBL.Rows.RemoveAt(ind)
                                    Cek = False
                                End If
                            ElseIf vDateTime.Day = 5 Then
                                If row.Is_Friday = 0 Then
                                    TBL.Rows.RemoveAt(ind)
                                    Cek = False
                                End If
                            ElseIf vDateTime.Day = 6 Then
                                If row.Is_Saturday = 0 Then
                                    TBL.Rows.RemoveAt(ind)
                                    Cek = False
                                End If
                            ElseIf vDateTime.Day = 7 Then
                                If row.Is_Sunday = 0 Then
                                    TBL.Rows.RemoveAt(ind)
                                    Cek = False
                                End If
                            End If
                        ElseIf AdvanceTime = EnumAdvanceFilter.ApplyOnSpecificDateOfTheMonth Then
                            If Format(vDateTime, "dd") = 1 Then
                                If row.Is_1 = 0 Then
                                    TBL.Rows.RemoveAt(ind)
                                    Cek = False
                                End If
                            End If
                            If Format(vDateTime, "dd") = 2 Then
                                If row.Is_2 = 0 Then
                                    TBL.Rows.RemoveAt(ind)
                                    Cek = False
                                End If
                            End If
                            If Format(vDateTime, "dd") = 3 Then
                                If row.Is_3 = 0 Then
                                    TBL.Rows.RemoveAt(ind)
                                    Cek = False
                                End If
                            End If
                            If Format(vDateTime, "dd") = 4 Then
                                If row.Is_4 = 0 Then
                                    TBL.Rows.RemoveAt(ind)
                                    Cek = False
                                End If
                            End If
                            If Format(vDateTime, "dd") = 5 Then
                                If row.Is_5 = 0 Then
                                    TBL.Rows.RemoveAt(ind)
                                    Cek = False
                                End If
                            End If
                            If Format(vDateTime, "dd") = 6 Then
                                If row.Is_6 = 0 Then
                                    TBL.Rows.RemoveAt(ind)
                                    Cek = False
                                End If
                            End If
                            If Format(vDateTime, "dd") = 7 Then
                                If row.Is_7 = 0 Then
                                    TBL.Rows.RemoveAt(ind)
                                    Cek = False
                                End If
                            End If
                            If Format(vDateTime, "dd") = 8 Then
                                If row.Is_8 = 0 Then
                                    TBL.Rows.RemoveAt(ind)
                                    Cek = False
                                End If
                            End If
                            If Format(vDateTime, "dd") = 9 Then
                                If row.Is_9 = 0 Then
                                    TBL.Rows.RemoveAt(ind)
                                    Cek = False
                                End If
                            End If
                            If Format(vDateTime, "dd") = 10 Then
                                If row.Is_10 = 0 Then
                                    TBL.Rows.RemoveAt(ind)
                                    Cek = False
                                End If
                            End If
                            If Format(vDateTime, "dd") = 11 Then
                                If row.Is_11 = 0 Then
                                    TBL.Rows.RemoveAt(ind)
                                    Cek = False
                                End If
                            End If
                            If Format(vDateTime, "dd") = 12 Then
                                If row.Is_12 = 0 Then
                                    TBL.Rows.RemoveAt(ind)
                                    Cek = False
                                End If
                            End If
                            If Format(vDateTime, "dd") = 13 Then
                                If row.Is_13 = 0 Then
                                    TBL.Rows.RemoveAt(ind)
                                    Cek = False
                                End If
                            End If
                            If Format(vDateTime, "dd") = 14 Then
                                If row.Is_14 = 0 Then
                                    TBL.Rows.RemoveAt(ind)
                                    Cek = False
                                End If
                            End If
                            If Format(vDateTime, "dd") = 15 Then
                                If row.Is_15 = 0 Then
                                    TBL.Rows.RemoveAt(ind)
                                    Cek = False
                                End If
                            End If
                            If Format(vDateTime, "dd") = 16 Then
                                If row.Is_16 = 0 Then
                                    TBL.Rows.RemoveAt(ind)
                                    Cek = False
                                End If
                            End If
                            If Format(vDateTime, "dd") = 17 Then
                                If row.Is_17 = 0 Then
                                    TBL.Rows.RemoveAt(ind)
                                    Cek = False
                                End If
                            End If
                            If Format(vDateTime, "dd") = 18 Then
                                If row.Is_18 = 0 Then
                                    TBL.Rows.RemoveAt(ind)
                                    Cek = False
                                End If
                            End If
                            If Format(vDateTime, "dd") = 19 Then
                                If row.Is_19 = 0 Then
                                    TBL.Rows.RemoveAt(ind)
                                    Cek = False
                                End If
                            End If
                            If Format(vDateTime, "dd") = 20 Then
                                If row.Is_20 = 0 Then
                                    TBL.Rows.RemoveAt(ind)
                                    Cek = False
                                End If
                            End If
                            If Format(vDateTime, "dd") = 21 Then
                                If row.Is_21 = 0 Then
                                    TBL.Rows.RemoveAt(ind)
                                    Cek = False
                                End If
                            End If
                            If Format(vDateTime, "dd") = 22 Then
                                If row.Is_22 = 0 Then
                                    TBL.Rows.RemoveAt(ind)
                                    Cek = False
                                End If
                            End If
                            If Format(vDateTime, "dd") = 23 Then
                                If row.Is_23 = 0 Then
                                    TBL.Rows.RemoveAt(ind)
                                    Cek = False
                                End If
                            End If
                            If Format(vDateTime, "dd") = 24 Then
                                If row.Is_24 = 0 Then
                                    TBL.Rows.RemoveAt(ind)
                                    Cek = False
                                End If
                            End If
                            If Format(vDateTime, "dd") = 25 Then
                                If row.Is_25 = 0 Then
                                    TBL.Rows.RemoveAt(ind)
                                    Cek = False
                                End If
                            End If
                            If Format(vDateTime, "dd") = 26 Then
                                If row.Is_26 = 0 Then
                                    TBL.Rows.RemoveAt(ind)
                                    Cek = False
                                End If
                            End If
                            If Format(vDateTime, "dd") = 27 Then
                                If row.Is_27 = 0 Then
                                    TBL.Rows.RemoveAt(ind)
                                    Cek = False
                                End If
                            End If
                            If Format(vDateTime, "dd") = 28 Then
                                If row.Is_28 = 0 Then
                                    TBL.Rows.RemoveAt(ind)
                                    Cek = False
                                End If
                            End If
                            If Format(vDateTime, "dd") = 29 Then
                                If row.Is_29 = 0 Then
                                    TBL.Rows.RemoveAt(ind)
                                    Cek = False
                                End If
                            End If
                            If Format(vDateTime, "dd") = 30 Then
                                If row.Is_30 = 0 Then
                                    TBL.Rows.RemoveAt(ind)
                                    Cek = False
                                End If
                            End If
                            If Format(vDateTime, "dd") = 31 Then
                                If row.Is_31 = 0 Then
                                    TBL.Rows.RemoveAt(ind)
                                    Cek = False
                                End If
                            End If
                        End If
                    End If
                End If
                If Cek = True Then
                    If row.Type <> EnumType.Nominal_Percentage Then
                        TBL.Rows.RemoveAt(ind)
                        Cek = False
                    End If
                End If
                If Cek = True Then
                    ind += 1
                End If
            End While

            '3. Jika tidak ada diskon yang memenuhi kriteria yang diberikan, maka kembalikan nol
            If TBL.Rows.Count = 0 Then
                Return "0"
            Else
                If TBL.Rows.Count = 1 Then
                    row = TBL.Rows(0)
                    vMaxQty = row.Max_Qty
                    Return row.Discount_Rate
                Else
                    'cek multidiskon, jika ada yang multidiskon, maka semua harus multi diskon
                    For Each row In TBL.Rows
                        If row.Multidiskon = 1 Then
                            JmlMultiDiskon += 1
                        End If
                    Next
                    If JmlMultiDiskon = TBL.Rows.Count Then
                        'kembalikan kombinasi diskon
                        For Each row In TBL.Rows
                            If Retval = "" Then
                                Retval = row.Discount_Rate
                                vMaxQty = row.Max_Qty
                            Else
                                Retval = Retval + "+" + row.Discount_Rate
                                If vMaxQty > row.Max_Qty Then
                                    vMaxQty = row.Max_Qty
                                End If
                            End If
                        Next
                        Return Retval
                    Else
                        'kembalikan diskon pada baris pertama
                        row = TBL.Rows(0)
                        vMaxQty = row.Max_Qty
                        Return row.Discount_Rate
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "0"
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
