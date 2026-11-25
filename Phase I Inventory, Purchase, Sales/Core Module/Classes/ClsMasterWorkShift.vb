Option Explicit On
Option Strict Off

Public Class ClsMasterWorkShift
    Dim ObjCD As New Core_Module.ClsCekData

    Dim mPrimary As Double = -1
    Dim mTipeWorkshift As EnumTipeWorkshift = EnumTipeWorkshift.HR
    Dim mCode As String = ""
    Dim mDesc As String = ""
    Dim mRemark As String = ""
    Dim mCreatedBy As String = ""
    Dim mCreatedDate As DateTime = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As DateTime = Now
    Dim mTimestamp As Date = Now

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
    Dim mUseOtOrder As Integer = 0
    Dim mNotificationEmailOt As Integer = 0
    'digunakan untuk antrian simpan data
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianMaster
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAAntrian As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter
    Dim ObjRnd As New Random
    Dim ObjAntrian As New Core_Module.ClsAntrian

    Dim ObjItem As New ClsMasterWorkShift_Item

    Public Enum EnumTipeWorkshift As Integer
        HR
        Production
    End Enum

    Public Enum TipeWorkdays As Integer
        Weekday
        Holiday
    End Enum

    Public Enum TipeOvertime As Integer
        Weekday
        Holiday
        No_Overtime
    End Enum

    Public Enum enumDuration
        fullday
        halfday
    End Enum

    Public Enum enumRoundingType As Integer
        No_Rounding
        Round_Up
        Round_down
        Middle
    End Enum
    Enum EnumSearchWorkshift
        Code
        Desc
        Remarks
    End Enum

    Enum EnumPilihan
        weekly
        holidaylist
        AttendanceSetting
        All
    End Enum

    Public Enum EnumScheduleType As Integer
        SCHEDULE_IN
        BREAK_OUT
        BREAK_IN
        SCHEDULE_OUT
    End Enum

    Public Enum EnumDefaultIfNoClockInOut As Integer
        SCHEDULE
        CUSTOM
    End Enum

    Public Enum EnumDefaultIfOutstation As Integer
        SCHEDULE
        CUSTOM
    End Enum

    Public Enum EnumCheckRange As Integer
        AllRange
        SchIn_SchBkOut
        SchBOut_SchBIn
        SchBIn_SchOut
    End Enum

    Public Property Primary() As Double
        Get
            Return Me.mPrimary
        End Get
        Set(ByVal value As Double)
            Me.mPrimary = value
        End Set
    End Property

    Public Property TipeWorkshift() As EnumTipeWorkshift
        Get
            Return Me.mTipeWorkshift
        End Get
        Set(ByVal value As EnumTipeWorkshift)
            Me.mTipeWorkshift = value
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

    Public Property Desc() As String
        Get
            Return Me.mDesc
        End Get
        Set(ByVal value As String)
            Me.mDesc = value
        End Set
    End Property

    Public Property Remarks() As String
        Get
            Return Me.mRemark
        End Get
        Set(ByVal value As String)
            Me.mRemark = value
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

    Private mIsCalculateEarlyOT As Integer
    Public Property IsCalculateEarlyOT() As Integer
        Get
            Return mIsCalculateEarlyOT
        End Get
        Set(ByVal value As Integer)
            mIsCalculateEarlyOT = value
        End Set
    End Property

    Private mIsCalculateLate As Integer
    Public Property IsCalculateLate() As Integer
        Get
            Return mIsCalculateLate
        End Get
        Set(ByVal value As Integer)
            mIsCalculateLate = value
        End Set
    End Property

    Private mIsCalculateEarlyOut As Integer
    Public Property IsCalculateEarlyOut() As Integer
        Get
            Return mIsCalculateEarlyOut
        End Get
        Set(ByVal value As Integer)
            mIsCalculateEarlyOut = value
        End Set
    End Property

    Private mIsCalculateOT As Integer
    Public Property IsCalculateOT() As Integer
        Get
            Return mIsCalculateOT
        End Get
        Set(ByVal value As Integer)
            mIsCalculateOT = value
        End Set
    End Property


    Private mCalculateEarlyOTTolerance As Integer
    Public Property CalculateEarlyOTTolerance() As Integer
        Get
            Return mCalculateEarlyOTTolerance
        End Get
        Set(ByVal value As Integer)
            mCalculateEarlyOTTolerance = value
        End Set
    End Property

    Private mCalculateLateTolerance As Integer
    Public Property CalculateLateTolerance() As Integer
        Get
            Return mCalculateLateTolerance
        End Get
        Set(ByVal value As Integer)
            mCalculateLateTolerance = value
        End Set
    End Property

    Private mCalculateEarlyOutTolerance As Integer
    Public Property CalculateEarlyOutTolerance() As Integer
        Get
            Return mCalculateEarlyOutTolerance
        End Get
        Set(ByVal value As Integer)
            mCalculateEarlyOutTolerance = value
        End Set
    End Property

    Private mCalculateOTTolerance As Integer
    Public Property CalculateOTTolerance() As Integer
        Get
            Return mCalculateOTTolerance
        End Get
        Set(ByVal value As Integer)
            mCalculateOTTolerance = value
        End Set
    End Property

    Private mCalculateEarlyOTOnWeekend As Integer
    Public Property CalculateEarlyOTOnWeekend() As Integer
        Get
            Return mCalculateEarlyOTOnWeekend
        End Get
        Set(ByVal value As Integer)
            mCalculateEarlyOTOnWeekend = value
        End Set
    End Property

    Private mCalculateLateOnWeekend As Integer
    Public Property CalculateLateOnWeekend() As Integer
        Get
            Return mCalculateLateOnWeekend
        End Get
        Set(ByVal value As Integer)
            mCalculateLateOnWeekend = value
        End Set
    End Property

    Private mCalculateEarlyOutOnWeekend As Integer
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

    Private mCalculateEarlyOTOnHoliday As Integer
    Public Property CalculateEarlyOTOnHoliday() As Integer
        Get
            Return mCalculateEarlyOTOnHoliday
        End Get
        Set(ByVal value As Integer)
            mCalculateEarlyOTOnHoliday = value
        End Set
    End Property

    Private mCalculateLateOnHoliday As Integer
    Public Property CalculateLateOnHoliday() As Integer
        Get
            Return mCalculateLateOnHoliday
        End Get
        Set(ByVal value As Integer)
            mCalculateLateOnHoliday = value
        End Set
    End Property

    Private mCalculateEarlyOutOnHoliday As Integer
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

    Public Property UseOtOrder() As Integer
        Get
            Return Me.mUseOtOrder
        End Get
        Set(ByVal value As Integer)
            Me.mUseOtOrder = value
        End Set
    End Property

    Public Property NotificationEmailOt() As Integer
        Get
            Return Me.mNotificationEmailOt
        End Get
        Set(ByVal value As Integer)
            Me.mNotificationEmailOt = value
        End Set
    End Property

    Public Function ApakahBolehHapus(ByVal vPrimary As Double) As Boolean
        Dim dtaAtt As New DseMasterWorkShiftTableAdapters.tbl_trans_att_viewTableAdapter
        Dim dtaAllocate As New DseMasterWorkShiftTableAdapters.tblmaster_workshift_itemallocateTableAdapter
        Dim dta1 As New DseMasterWorkShiftTableAdapters.tblmaster_workshiftTableAdapter

        Dim Dummy As Nullable(Of Long)
        Try
            dtaAtt.Connection.ConnectionString = Core_Module.GetConstring
            dtaAllocate.Connection.ConnectionString = Core_Module.GetConstring
            dta1.Connection.ConnectionString = Core_Module.GetConstring

            'cek workshift di attendance
            Dummy = dtaAtt.ScalarQuery_PrimaryWS(vPrimary)
            If Dummy.HasValue = True Then
                If Dummy.Value > 0 Then
                    MsgBox("This data referred by attendance data, you can't delete this data.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If
            'cek workshift di allocate
            Dummy = dtaAllocate.ScalarQuery_GetJmlByPrimaryWS(vPrimary)
            If Dummy.HasValue = True Then
                If Dummy.Value > 0 Then
                    MsgBox("This data referred by workshift allocation data, you can't delete this data.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If
            Dummy = dta1.ScalarQuery_GetJmlInAllocate(vPrimary)
            If Dummy.HasValue = True Then
                If Dummy.Value > 0 Then
                    MsgBox("This data referred by workshift allocation data, you can't delete this data.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If
            'cek workshift di work unit 
            Dummy = dta1.ScalarQuery_GetJmlWorkUnit(vPrimary)
            If Dummy.HasValue = True Then
                If Dummy.Value > 0 Then
                    MsgBox("This data referred by work unit data, you can't delete this data.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'benar
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Erorr")
        Finally
            dtaAtt.Dispose()
            dtaAllocate.Dispose()
            dta1.Dispose()
        End Try
    End Function

    Public Function ApakahBolehEdit(ByVal vPrimary As Double) As Boolean
        Dim dtaAtt As New DseMasterWorkShiftTableAdapters.tbl_trans_att_viewTableAdapter
        Dim dtaAllocate As New DseMasterWorkShiftTableAdapters.tblmaster_workshift_itemallocateTableAdapter
        Dim Dummy As Nullable(Of Long)
        Try
            dtaAtt.Connection.ConnectionString = Core_Module.GetConstring
            dtaAllocate.Connection.ConnectionString = Core_Module.GetConstring
            'cek workshift di attendance
            Dummy = dtaAtt.ScalarQuery_PrimaryWS(vPrimary)
            If Dummy.HasValue = True Then
                If Dummy.Value > 0 Then
                    Return False
                End If
            End If
            'cek workshift di allocate
            Dummy = dtaAllocate.ScalarQuery_GetJmlByPrimaryWS(vPrimary)
            If Dummy.HasValue = True Then
                If Dummy.Value > 0 Then
                    Return False
                End If
            End If

            'benar
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Erorr")
        Finally
            dtaAtt.Dispose()
            dtaAllocate.Dispose()
        End Try
    End Function

    Public Overridable Function SimpanBaru(ByVal vUserCode As String, ByVal vDGVweekly As DataGridView, ByVal vDGVholiday As DataGridView, _
   ByVal frm As FrmMasterWorkShift) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            DTAAntrian.Connection.ConnectionString = Core_Module.GetConstring
            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, ClsAntrian.EnumTipeAntrian.AntrianMaster, _
            Me.vPrimary_Antrian, Me.mPrimary, ClsAntrian.EnumTipeTransaksiAsalAntrian.MasterWorkshift_SimpanBaru, vUserCode)

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
                            retval = Me.SimpanBaru1(vUserCode, vDGVweekly, frm, vDGVholiday)
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

    Public Function SimpanBaru1(ByVal vUserCode As String, ByVal vDGVweekly As DataGridView, ByVal frm As FrmMasterWorkShift, _
    ByVal vDGVHoliday As DataGridView) As Boolean
        Dim DTA2 As New DseMasterWorkShiftTableAdapters.tblmaster_workshift_newTableAdapter
        Try
            DTA2.Connection.ConnectionString = Core_Module.GetConstring
            If DTA2.InsertQuery(Me.mCode, Me.mDesc, Me.mRemark, vUserCode, Me.mCreatedDate, vUserCode, Me.mModifiedDate, Me.mTimestamp, _
                    Me.IsCalculateEarlyOT, Me.IsCalculateLate, Me.IsCalculateEarlyOut, Me.IsCalculateEarlyOT, _
                    Me.mCalculateAfterEarlyOT, Me.mCalculateAfterLate, Me.mCalculateAfterEarlyOut, Me.mCalculateAfterOT, _
                    Me.mRoundingTypeEarlyOT, Me.mRoundingTypeLate, Me.mRoundingTypeEarlyOut, Me.mRoundingTypeOT, _
                    Me.mRoundingValueEarlyOT, Me.mRoundingValueLate, Me.mRoundingValueEarlyOut, Me.mRoundingValueOT, _
                    Me.mMaxValueEarlyOT, Me.mMaxValueLate, Me.mMaxValueEarlyOut, Me.mMaxValueOT, _
                    Me.mCalculateEarlyOTTolerance, Me.mCalculateLateTolerance, Me.mCalculateEarlyOutTolerance, Me.mCalculateOTTolerance, _
                    Me.mCalculateEarlyOTOnWeekend, Me.mCalculateLateOnWeekend, Me.mCalculateEarlyOutOnWeekend, Me.mCalculateOTOnWeekend, _
                    Me.mCalculateEarlyOTOnHoliday, Me.mCalculateLateOnHoliday, Me.mCalculateEarlyOutOnHoliday, Me.mCalculateOTOnHoliday, _
                    Me.mTipeWorkshift, Me.mUseOtOrder, Me.mNotificationEmailOt) > 0 Then

                Me.mPrimary = Me.GetPrimary(Me.mCode, Me.mTipeWorkshift)
                Me.mTimestamp = Me.GetTimeStampByPrimary(Me.mPrimary)
                Me.mCreatedDate = Me.mTimestamp
                Me.mModifiedDate = Me.mTimestamp
                Me.mCreatedBy = vUserCode
                Me.mModifiedBy = vUserCode

                'simpan item
                Me.ObjItem.SimpanItem(vDGVweekly, vDGVHoliday, Me.mPrimary, frm)
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA2.Dispose()
        End Try
    End Function

    Public Overridable Function SimpanPerubahan(ByVal vUserCode As String, ByVal vDGVweekly As DataGridView, ByVal vDGVholiday As DataGridView, _
   ByVal frm As FrmMasterWorkShift) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            DTAAntrian.Connection.ConnectionString = Core_Module.GetConstring
            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, ClsAntrian.EnumTipeAntrian.AntrianMaster, _
            Me.vPrimary_Antrian, Me.mPrimary, ClsAntrian.EnumTipeTransaksiAsalAntrian.MasterWorkshift_SimpanPerubahan, vUserCode)

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
                            retval = Me.SimpanPerubahan1(vUserCode, vDGVweekly, frm, vDGVholiday)
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

    Public Function SimpanPerubahan1(ByVal vUserCode As String, ByVal vDGVweekly As DataGridView, ByVal frm As FrmMasterWorkShift, _
    ByVal vDGVHoliday As DataGridView) As Boolean
        Dim DTA As New DseMasterWorkShiftTableAdapters.tblmaster_workshift_newTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'If Me.ApakahBolehEdit(Me.mPrimary) = True Then
            If DTA.UpdateQuery(Me.mCode, Me.mDesc, Me.mRemark, vUserCode, Me.mModifiedDate, Me.mTimestamp, _
                    Me.IsCalculateEarlyOT, Me.IsCalculateLate, Me.IsCalculateEarlyOut, Me.IsCalculateOT, _
                    Me.mCalculateAfterEarlyOT, Me.mCalculateAfterLate, Me.mCalculateAfterEarlyOut, Me.mCalculateAfterOT, _
                    Me.mRoundingTypeEarlyOT, Me.mRoundingTypeLate, Me.mRoundingTypeEarlyOut, Me.mRoundingTypeOT, _
                    Me.mRoundingValueEarlyOT, Me.mRoundingValueLate, Me.mRoundingValueEarlyOut, Me.mRoundingValueOT, _
                    Me.mMaxValueEarlyOT, Me.mMaxValueLate, Me.mMaxValueEarlyOut, Me.mMaxValueOT, _
                    Me.mCalculateEarlyOTTolerance, Me.mCalculateLateTolerance, Me.mCalculateEarlyOutTolerance, Me.mCalculateOTTolerance, _
                    Me.mCalculateEarlyOTOnWeekend, Me.mCalculateLateOnWeekend, Me.mCalculateEarlyOutOnWeekend, Me.mCalculateOTOnWeekend, _
                    Me.mCalculateEarlyOTOnHoliday, Me.mCalculateLateOnHoliday, Me.mCalculateEarlyOutOnHoliday, Me.mCalculateOTOnHoliday, _
                    Me.mTipeWorkshift, Me.mUseOtOrder, Me.mNotificationEmailOt, Me.mPrimary) > 0 Then
                Me.mTimestamp = Me.GetTimeStampByPrimary(Me.mPrimary)
                Me.mModifiedDate = Me.mTimestamp
                Me.mModifiedBy = vUserCode

                'hapus item
                Me.ObjItem.HapusItem(Me.mPrimary)

                'simpan item
                Me.ObjItem.SimpanItem(vDGVweekly, vDGVHoliday, Me.mPrimary, frm)
                Return True
            Else
                Return False
            End If
            'Else
            'MsgBox("Workshift data can't edit because this data has refered allocation item or attendance item", _
            'MsgBoxStyle.Information, "Information")
            'Return False
            'End If
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
        Dim SW As New Stopwatch

        Try
            DTAAntrian.Connection.ConnectionString = Core_Module.GetConstring
            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, ClsAntrian.EnumTipeAntrian.AntrianMaster, _
            Me.vPrimary_Antrian, Me.mPrimary, ClsAntrian.EnumTipeTransaksiAsalAntrian.MasterWorkshift_Hapus, vUser)

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
                            retval = Me.Hapus1(vPrimaryTransaksi)
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

    Public Function Hapus1(ByVal vPrimary As Double) As Boolean
        Dim DTA As New DseMasterWorkShiftTableAdapters.tblmaster_workshiftTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahBolehHapus(vPrimary) = True Then
                DTA.Delete(vPrimary)
                Me.ObjItem.HapusItem(vPrimary)
                Return True
            Else
                MsgBox("Workshift data can't delete because this data has refered allocation item or attendance item", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahSudahAda(ByVal vCode As String, ByVal vTipeWS As EnumTipeWorkshift) As Boolean
        Dim dta As New DseMasterWorkShiftTableAdapters.tblmaster_workshiftTableAdapter
        Dim dummy As Nullable(Of Long)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dummy = dta.ScalarQuery_GetJmlPrimaryByCode(vCode, vTipeWS)
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

    Public Function GetCodeByPrimary(ByVal vPrimary As Double) As String
        Dim dta As New DseMasterWorkShiftTableAdapters.tblmaster_workshiftTableAdapter
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

    Public Function GetJmlAtt(ByVal vDate As Date, ByVal vPrimaryWS As Double) As Double
        Dim dta As New DseMasterWorkShiftTableAdapters.tbl_trans_att_viewTableAdapter
        Dim Dummy As Nullable(Of Long)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = dta.ScalarQuery_GetJmlByHolDateNPrimaryWS(vPrimaryWS, vDate)
            If Dummy.HasValue = True Then
                Return Dummy.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetTimeStampByPrimary(ByVal vPrimary As Double) As Date
        Dim dta As New DseMasterWorkShiftTableAdapters.tblmaster_workshiftTableAdapter
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

    Public Function CariDetail(ByVal vPrimary As Double, ByVal vFrm As FrmMasterWorkShift, _
    ByVal vDgvWeekly As DataGridView, ByVal vDgvHoliday As DataGridView, ByVal vObjCd As Core_Module.ClsCekData) As Boolean
        Dim DTA As New DseMasterWorkShiftTableAdapters.tblmaster_workshift_newTableAdapter
        Dim TBL As New DseMasterWorkShift.tblmaster_workshift_newDataTable
        Dim row As DseMasterWorkShift.tblmaster_workshift_newRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_Primary(TBL, vPrimary)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = .primary
                    Me.mCode = .code
                    Me.mDesc = .description
                    Me.Remarks = .remarks
                    Me.IsCalculateEarlyOT = .is_calculate_early_ot
                    Me.IsCalculateLate = .is_calculate_late
                    Me.IsCalculateEarlyOut = .is_calculate_early_out
                    Me.IsCalculateOT = .is_calculate_ot
                    Me.mCreatedBy = .created_by
                    Me.mCreatedDate = .created_date
                    Me.mModifiedBy = .modified_by
                    Me.mModifiedDate = .modified_date
                    Me.mTimestamp = .timestamp
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
                    Me.mCalculateEarlyOTTolerance = .calculate_early_ot_tolerance
                    Me.mCalculateLateTolerance = .calculate_late_tolerance
                    Me.mCalculateEarlyOutTolerance = .calculate_early_out_tolerance
                    Me.mCalculateOTTolerance = .calculate_ot_tolerance
                    Me.mCalculateEarlyOTOnWeekend = .calculate_early_ot_on_weekend
                    Me.mCalculateLateOnWeekend = .calculate_late_on_weekend
                    Me.mCalculateEarlyOutOnWeekend = .calculate_early_out_on_weekend
                    Me.mCalculateOTOnWeekend = .calculate_ot_on_weekend
                    Me.mCalculateEarlyOTOnHoliday = .calculate_early_ot_on_holiday
                    Me.mCalculateLateOnHoliday = .calculate_late_on_holiday
                    Me.mCalculateEarlyOutOnHoliday = .calculate_early_out_on_holiday
                    Me.mCalculateOTOnHoliday = .calculate_ot_on_holiday
                    Me.mTipeWorkshift = .tipe_workshif
                    Me.mUseOtOrder = .use_ot_order
                    Me.NotificationEmailOt = .notification_email_ot
                    Me.ObjItem.IsiDGVItem(vDgvWeekly, vDgvHoliday, vPrimary, vFrm, vObjCd, ClsMasterWorkShift.EnumPilihan.All, .tipe_workshif)
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

    Public Function IsiDGV(ByVal Criteria As String, ByVal vDGV As DataGridView, ByVal TBL As DseMasterWorkShift.tblmaster_workshiftDataTable, _
    ByVal SearchCriteria As EnumSearchWorkshift, ByVal tsl As ToolStripLabel, ByVal vTipeWS As EnumTipeWorkshift) As Boolean
        Dim DTA As New DseMasterWorkShiftTableAdapters.tblmaster_workshiftTableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If SearchCriteria = EnumSearchWorkshift.Code Then
                QueryResult = DTA.FillBy_LikeCode(TBL, Criteria + "%", vTipeWS)
            ElseIf SearchCriteria = EnumSearchWorkshift.Desc Then
                QueryResult = DTA.FillBy_LikeDesc(TBL, Criteria + "%", vTipeWS)
            ElseIf SearchCriteria = EnumSearchWorkshift.Remarks Then
                QueryResult = DTA.FillBy_LikeRemarks(TBL, Criteria + "%", vTipeWS)
            End If

            'update dgv
            vDGV.DataSource = TBL.DefaultView
            vDGV.Refresh()

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
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As _
    EnumSearchWorkshift, ByVal vTipeWS As EnumTipeWorkshift) As Double
        Dim DTA As New DseMasterWorkShiftTableAdapters.tblmaster_workshiftTableAdapter
        Dim Dummy As Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If vSearchBy = EnumSearchWorkshift.Code Then
                Dummy = DTA.ScalarQuery_GetJmlByLikeCode(vKriteria + "%", vTipeWS)
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchWorkshift.Desc Then
                Dummy = DTA.ScalarQuery_GetJmlByLikeDesc(vKriteria + "%", vTipeWS)
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchWorkshift.Remarks Then
                Dummy = DTA.ScalarQuery_GetJmlByLikeRemarks(vKriteria + "%", vTipeWS)
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

    Public Function HitungBreakDuration(ByVal vSchin As Double, ByVal vSchout As Double, ByVal vSchBreakOut As Double, ByVal vSchBreakIn As Double) As Double
        Dim duration As Double
        Try
            If vSchin > vSchout Then
                If vSchBreakOut > vSchBreakIn Then
                    duration = ((24 * 60) - vSchBreakOut) + vSchBreakIn
                Else
                    duration = vSchBreakIn - vSchBreakOut
                End If
            Else
                duration = vSchBreakIn - vSchBreakOut
            End If

            Return duration
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetDateMAXAttendance(ByVal vPrimaryWorkshift As Double) As Date
        Dim DTA As New DseMasterWorkShiftTableAdapters.DataTable1TableAdapter
        Dim Dummy As System.Nullable(Of Date)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetNilaiMaxAttendanceByWorkshift(vPrimaryWorkshift)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub New()
        Me.DTAAntrian.Connection.ConnectionString = Core_Module.GetConstring
    End Sub

    Public Function GetLimitBefore(ByVal vPrimary As Double, ByVal vSchType As ClsMasterWorkShift.EnumScheduleType) As Integer
        Dim Dta As New DseMasterWorkShiftTableAdapters.tblmaster_workshift_itemparametersTableAdapter
        Dim LimitBefore_ As Nullable(Of Decimal)
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            LimitBefore_ = Dta.ScalarQuery_GetLimitBefore(vSchType, vPrimary)
            If LimitBefore_.HasValue = True Then
                Return CInt(LimitBefore_.Value)
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Function GetLimitAfter(ByVal vPrimary As Double, ByVal vSchType As ClsMasterWorkShift.EnumScheduleType) As Integer
        Dim Dta As New DseMasterWorkShiftTableAdapters.tblmaster_workshift_itemparametersTableAdapter
        Dim LimitAfter_ As Nullable(Of Decimal)
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            LimitAfter_ = Dta.ScalarQuery_GetLimitAfter(vSchType, vPrimary)
            If LimitAfter_.HasValue = True Then
                Return CInt(LimitAfter_.Value)
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Function GetIsGetValue(ByVal vPrimary As Double, ByVal vSchType As ClsMasterWorkShift.EnumScheduleType) As Boolean
        Dim Dta As New DseMasterWorkShiftTableAdapters.tblmaster_workshift_itemparametersTableAdapter
        Dim IsGetValue_ As Nullable(Of Integer)
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            IsGetValue_ = Dta.ScalarQuery_GetIsGetValue(vSchType, vPrimary)
            If IsGetValue_.HasValue = True Then
                If IsGetValue_.Value = 0 Then
                    Return False
                Else
                    Return True
                End If
            Else
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return True
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Function GetFirstOrLast(ByVal vPrimary As Double, ByVal vSchType As ClsMasterWorkShift.EnumScheduleType) As Integer
        Dim Dta As New DseMasterWorkShiftTableAdapters.tblmaster_workshift_itemparametersTableAdapter
        Dim FirstOrLast_ As Nullable(Of Long)
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            FirstOrLast_ = Dta.ScalarQuery_GetFirstOrLast(vSchType, vPrimary)
            If FirstOrLast_.HasValue = True Then
                Return FirstOrLast_
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Function GetDefaultIfNoClockInOut(ByVal vPrimary As Double, ByVal vSchType As ClsMasterWorkShift.EnumScheduleType) As ClsMasterWorkShift.EnumDefaultIfNoClockInOut
        Dim Dta As New DseMasterWorkShiftTableAdapters.tblmaster_workshift_itemparametersTableAdapter
        Dim DefaultIfNoClockInOut_ As Nullable(Of Long)
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            DefaultIfNoClockInOut_ = Dta.ScalarQuery_GetDefault(vSchType, vPrimary)
            If DefaultIfNoClockInOut_.HasValue Then
                Return DefaultIfNoClockInOut_
            Else
                Return EnumDefaultIfNoClockInOut.SCHEDULE
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return EnumDefaultIfNoClockInOut.SCHEDULE
        Finally
            Dta.Dispose()
        End Try
    End Function

    'aul -penambahan outsatation-
    Public Function GetDefaultIfOutstationInOut(ByVal vPrimary As Double, ByVal vSchType As ClsMasterWorkShift.EnumScheduleType) As ClsMasterWorkShift.EnumDefaultIfOutstation
        Dim Dta As New DseMasterWorkShiftTableAdapters.tblmaster_workshift_itemparametersTableAdapter
        Dim DefaultIfOutstationInOut_ As Nullable(Of Long)
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            DefaultIfOutstationInOut_ = Dta.ScalarQuery_GetDefaultOutstation(vSchType, vPrimary)
            If DefaultIfOutstationInOut_.HasValue Then
                Return DefaultIfOutstationInOut_
            Else
                Return EnumDefaultIfOutstation.SCHEDULE
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return EnumDefaultIfOutstation.SCHEDULE
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Function GetDefaultCustom(ByVal vPrimary As Double, ByVal vSchType As ClsMasterWorkShift.EnumScheduleType) As String
        Dim Dta As New DseMasterWorkShiftTableAdapters.tblmaster_workshift_itemparametersTableAdapter
        Dim DefaultCustom_ As String
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            DefaultCustom_ = Dta.ScalarQuery_GetDefaultCustom(vSchType, vPrimary)
            If String.IsNullOrEmpty(DefaultCustom_) = False Then
                Return DefaultCustom_
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

    Public Function GetAttendaceSettings(ByVal vPrimary As Double, ByRef IsCalculateEarlyOTWS As Boolean, ByRef IsCalculateLateWS As Boolean, _
            ByRef IsCalculateEarlyOutWS As Boolean, ByRef IsCalculateOTWS As Boolean, _
            ByRef CalculateAfterEarlyOTWS As Double, ByRef CalculateAfterLateWS As Double, ByRef CalculateAfterEarlyOutWS As Double, ByRef CalculateAfterOTWS As Double, _
            ByRef RoundingTypeEarlyOTWS As Integer, ByRef RoundingTypeLateWS As Integer, ByRef RoundingTypeEarlyOutWS As Integer, ByRef RoundingTypeOTWS As Integer, _
            ByRef RoundingValueEarlyOTWS As Integer, ByRef RoundingValueLateWS As Integer, ByRef RoundingValueEarlyOutWS As Integer, ByRef RoundingValueOTWS As Integer, _
            ByRef MaxValueEarlyOTWS As Integer, ByRef MaxValueLateWS As Integer, ByRef MaxValueEarlyOutWS As Integer, ByRef MaxValueOTWS As Integer, _
            ByRef vCalculateEarlyOTTolerance As Integer, ByRef vCalculateLateTolerance As Integer, ByRef vCalculateEarlyOutTolerance As Integer, ByRef vCalculateOTTolerance As Integer, _
            ByRef vCalculateEarlyOTOnWeekend As Integer, ByRef vCalculateLateOnWeekend As Integer, ByRef vCalculateEarlyOutOnWeekend As Integer, ByRef vCalculateOTOnWeekend As Integer, _
            ByRef vCalculateEarlyOTOnHoliday As Integer, ByRef vCalculateLateOnHoliday As Integer, ByRef vCalculateEarlyOutOnHoliday As Integer, ByRef vCalculateOTOnHoliday As Integer) As Boolean
        Dim Dta As New DseMasterWorkShiftTableAdapters.tblmaster_workshift_newTableAdapter
        Dim Tbl As New DseMasterWorkShift.tblmaster_workshift_newDataTable
        Dim Row As DseMasterWorkShift.tblmaster_workshift_newRow
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            If Dta.FillBy_Primary(Tbl, vPrimary) > 0 Then
                Row = Tbl.Rows(0)
                With Row
                    If .is_calculate_early_ot = 0 Then
                        IsCalculateEarlyOTWS = False
                        CalculateAfterEarlyOTWS = 0
                        RoundingTypeEarlyOTWS = 0
                        RoundingValueEarlyOTWS = 0
                        MaxValueEarlyOTWS = 0
                    Else
                        IsCalculateEarlyOTWS = True
                        CalculateAfterEarlyOTWS = .calculate_after_early_OT
                        RoundingTypeEarlyOTWS = .rounding_type_early_ot
                        RoundingValueEarlyOTWS = .rounding_value_early_ot
                        MaxValueEarlyOTWS = .max_value_early_ot
                    End If

                    If .is_calculate_late = 0 Then
                        IsCalculateLateWS = False
                        CalculateAfterLateWS = 0
                        RoundingTypeLateWS = 0
                        RoundingValueLateWS = 0
                        MaxValueLateWS = 0
                    Else
                        IsCalculateLateWS = True
                        CalculateAfterLateWS = .calculate_after_late
                        RoundingTypeLateWS = .rounding_type_late
                        RoundingValueLateWS = .rounding_value_late
                        MaxValueLateWS = .max_value_late
                    End If

                    If .is_calculate_early_out = 0 Then
                        IsCalculateEarlyOutWS = False
                        CalculateAfterEarlyOutWS = 0
                        RoundingTypeEarlyOutWS = 0
                        RoundingValueEarlyOutWS = 0
                        MaxValueEarlyOutWS = 0
                    Else
                        IsCalculateEarlyOutWS = True
                        CalculateAfterEarlyOutWS = .calculate_after_early_out
                        RoundingTypeEarlyOutWS = .rounding_type_early_out
                        RoundingValueEarlyOutWS = .rounding_value_early_out
                        MaxValueEarlyOutWS = .max_value_early_out
                    End If

                    If .is_calculate_ot = 0 Then
                        IsCalculateOTWS = False
                        CalculateAfterOTWS = 0
                        RoundingTypeOTWS = 0
                        RoundingValueOTWS = 0
                        MaxValueOTWS = 0
                    Else
                        IsCalculateOTWS = True
                        CalculateAfterOTWS = .calculate_after_ot
                        RoundingTypeOTWS = .rounding_type_ot
                        RoundingValueOTWS = .rounding_value_ot
                        MaxValueOTWS = .max_value_ot
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
                CalculateAfterEarlyOTWS = 0
                RoundingTypeEarlyOTWS = 0
                RoundingValueEarlyOTWS = 0
                MaxValueEarlyOTWS = 0
                CalculateAfterLateWS = 0
                RoundingTypeLateWS = 0
                RoundingValueLateWS = 0
                MaxValueLateWS = 0
                CalculateAfterEarlyOutWS = 0
                RoundingTypeEarlyOutWS = 0
                RoundingValueEarlyOutWS = 0
                MaxValueEarlyOutWS = 0
                CalculateAfterOTWS = 0
                RoundingTypeEarlyOutWS = 0
                RoundingValueEarlyOutWS = 0
                MaxValueEarlyOutWS = 0

                vCalculateEarlyOTTolerance = 1
                vCalculateLateTolerance = 1
                vCalculateEarlyOutTolerance = 1
                vCalculateOTTolerance = 1

                vCalculateEarlyOTOnWeekend = 0
                vCalculateEarlyOTOnHoliday = 0
                vCalculateLateOnWeekend = 0
                vCalculateLateOnHoliday = 0
                vCalculateEarlyOutOnWeekend = 0
                vCalculateEarlyOutOnHoliday = 0
                vCalculateOTOnWeekend = 1
                vCalculateOTOnHoliday = 1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function ApakahUseOTOrder_WS(ByVal vPrimary As Double) As Boolean
        Dim DTA As New DseMasterWorkShiftTableAdapters.tblmaster_workshift_newTableAdapter
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

    Public Function ApakahUseNotifEmailOTOrder_WS(ByVal vPrimary As Double) As Boolean
        Dim DTA As New DseMasterWorkShiftTableAdapters.tblmaster_workshift_newTableAdapter
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
End Class
