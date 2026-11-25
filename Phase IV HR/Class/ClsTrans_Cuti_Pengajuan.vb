Option Explicit On
Option Strict Off
Imports System.Windows.Forms

Public Class ClsTrans_Cuti_Pengajuan
    Dim ObjCD As New Core_Module.ClsCekData

    Dim mPrimary As Double = -1
    Dim mNoTransaksi As String = ""
    Dim mUrutuan As Double = -1
    Dim mTanggal As Date = #1/1/1950#
    Dim mPrimaryKaryawan As Double = -1
    Dim mPrimaryJob As Double = -1
    Dim mPrimaryCuti As Double = -1
    Dim mStatus As EnumStatus = EnumStatus.Waiting
    Dim mStatusNote As String = ""
    Dim mAppBy As String = ""
    Dim mAppDate As Date = #1/1/1950#
    Dim mStartDate As Date = #1/1/1950#
    Dim mHalfDayStart As EnumHalfDay = EnumHalfDay.FullDay
    Dim mEndDAte As Date = #1/1/1950#
    Dim mHalfDayEnd As EnumHalfDay = EnumHalfDay.FullDay
    Dim mDuration As Decimal = 0
    Dim mReason As String = ""
    Dim mCreatedBy As String = ""
    Dim mCreatedDate As DateTime = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As DateTime = Now
    Dim mTimestamp As Date = Now
    'digunakan untuk antrian simpan data
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianTransaksiHRPayroll
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAAntrian As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter
    Dim ObjRnd As New Random
    Dim ObjAntrian As New Core_Module.ClsAntrian

    Public Enum EnumPrivilagePengajuanCuti As Integer
        All_Employee
        Concerned_Employee
    End Enum

    Enum EnumStatus
        Waiting
        App
        Rejected
    End Enum

    Enum EnumHalfDay
        FullDay
        HalfDay
    End Enum

    Enum EnumSreach
        DocumentNo
        _Date
        EmployeeCode
        LeaveCode
    End Enum

    Public Enum EnumStatusExplain As Integer '-
        Leave
        Sick
        Permit
        Absent
    End Enum

#Region "Property"
    Public Property Primary() As Double
        Get
            Return Me.mPrimary
        End Get
        Set(ByVal value As Double)
            Me.mPrimary = value
        End Set
    End Property

    Public Property NoTransaksi() As String
        Get
            Return Me.mNoTransaksi
        End Get
        Set(ByVal value As String)
            Me.mNoTransaksi = value
        End Set
    End Property

    Public Property Urutan() As Double
        Get
            Return Me.mUrutuan
        End Get
        Set(ByVal value As Double)
            Me.mUrutuan = value
        End Set
    End Property

    Public Property Tanggal() As Date
        Get
            Return Me.mTanggal
        End Get
        Set(ByVal value As Date)
            Me.mTanggal = value
        End Set
    End Property

    Public Property PrimaryKaryawan() As Double
        Get
            Return Me.mPrimaryKaryawan
        End Get
        Set(ByVal value As Double)
            Me.mPrimaryKaryawan = value
        End Set
    End Property

    Public Property Primaryjjob() As Double
        Get
            Return Me.mPrimaryJob
        End Get
        Set(ByVal value As Double)
            Me.mPrimaryJob = value
        End Set
    End Property

    Public Property PrimaryCuti() As Double
        Get
            Return Me.mPrimaryCuti
        End Get
        Set(ByVal value As Double)
            Me.mPrimaryCuti = value
        End Set
    End Property

    Public Property Status() As EnumStatus
        Get
            Return Me.mStatus
        End Get
        Set(ByVal value As EnumStatus)
            Me.mStatus = value
        End Set
    End Property

    Public Property StatusNote() As String
        Get
            Return Me.mStatusNote
        End Get
        Set(ByVal value As String)
            Me.mStatusNote = value
        End Set
    End Property

    Public Property AppBy() As String
        Get
            Return Me.mAppBy
        End Get
        Set(ByVal value As String)
            Me.mAppBy = value
        End Set
    End Property

    Public Property AppDate() As Date
        Get
            Return Me.mAppDate
        End Get
        Set(ByVal value As Date)
            Me.mAppDate = value
        End Set
    End Property

    Public Property StartDate() As Date
        Get
            Return Me.mStartDate
        End Get
        Set(ByVal value As Date)
            Me.mStartDate = value
        End Set
    End Property

    Public Property HalfDayStart() As EnumHalfDay
        Get
            Return Me.mHalfDayStart
        End Get
        Set(ByVal value As EnumHalfDay)
            Me.mHalfDayStart = value
        End Set
    End Property

    Public Property EndDate() As Date
        Get
            Return Me.mEndDAte
        End Get
        Set(ByVal value As Date)
            Me.mEndDAte = value
        End Set
    End Property

    Public Property HalfDayEnd() As EnumHalfDay
        Get
            Return Me.mHalfDayEnd
        End Get
        Set(ByVal value As EnumHalfDay)
            Me.mHalfDayEnd = value
        End Set
    End Property

    Public Property Duration() As Decimal
        Get
            Return Me.mDuration
        End Get
        Set(ByVal value As Decimal)
            Me.mDuration = value
        End Set
    End Property

    Public Property Reason() As String
        Get
            Return Me.mReason
        End Get
        Set(ByVal value As String)
            Me.mReason = value
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

    Private Function ApakahDataBenar(ByVal vUser As String, Optional ByVal vBaru As Boolean = True) As Boolean
        Dim Dummy As String = ""
        Dim DummyStartDate As Date = Today
        Dim JmlAttendance As Integer = 0
        Dim ObjAttendanceItem As New ClsTrans_Attendance_Item

        Try
            If vBaru = True Then
                If Me.mStatus <> EnumStatus.Waiting Then
                    MsgBox("Aplication status isn't waiting of approval, you can't save this data.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            If vBaru = False Then
                If Me.ApakahBolehEdit(vUser) = False Then
                    Return False
                End If
            End If

            Dummy = ""
            DummyStartDate = Me.mStartDate
            While DummyStartDate.Date <= Me.mEndDAte
                JmlAttendance = ObjAttendanceItem.GetJml(Me.mPrimaryKaryawan, DummyStartDate)
                If JmlAttendance > 0 Then
                    If Dummy = "" Then
                        Dummy = Format(DummyStartDate, "dd MMM yyyy")
                    Else
                        Dummy = Dummy + ", " + Format(DummyStartDate, "dd MMM yyyy")
                    End If
                End If
                DummyStartDate = DummyStartDate + TimeSpan.FromDays(1)
            End While
            'MsgBox("Tolong direpair")
            If Dummy <> "" Then
                MsgBox("There are attendance items for the following date: " + Dummy + _
                ". Process cancelled.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjAttendanceItem = Nothing
        End Try
    End Function

    Public Function ApakahBolehEdit(ByVal vUser As String) As Boolean
        Dim dta As New DseTrans_Cuti_PengajuanTableAdapters.tbltrans_cuti_pengajuanTableAdapter
        Dim dummy As String
        Try
            'strict edit
            dummy = dta.ScalarQueryGetCreatedByPrimary(Me.mPrimary)
            If String.IsNullOrEmpty(dummy) = False Then
                If dummy <> vUser Then
                    MsgBox("Only user " + dummy + " can edit this transaction", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If
            Me.mStatus = Me.GetStatus(Me.mPrimary)
            'jika reject
            If Me.mStatus = EnumStatus.Rejected Then
                MsgBox("Application status is rejected, you can't edit this data", MsgBoxStyle.Information, "Information")
                Return False
            End If
            'jika approved
            If Me.mStatus = EnumStatus.App Then
                MsgBox("Application status is approved, you can't edit this data", MsgBoxStyle.Information, "Information")
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function ApakahBolehHapus(ByVal vUser As String) As Boolean
        Dim dta As New DseTrans_Cuti_PengajuanTableAdapters.tbltrans_cuti_pengajuanTableAdapter
        Dim dummy As String
        Try
            'strict delete
            dummy = dta.ScalarQueryGetCreatedByPrimary(Me.mPrimary)
            If String.IsNullOrEmpty(dummy) = False Then
                If dummy <> vUser Then
                    MsgBox("Only user " + dummy + " can dalete this transaction", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If
            Me.mStatus = Me.GetStatus(Me.mPrimary)
            'jika reject
            If Me.mStatus = EnumStatus.Rejected Then
                MsgBox("Application status is rejected, you can't delete this data", MsgBoxStyle.Information, "Information")
                Return False
            End If
            'jika approved
            If Me.mStatus = EnumStatus.App Then
                MsgBox("Application status is approved, you can't delete this data", MsgBoxStyle.Information, "Information")
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Overridable Function SimpanBaruSetting(ByVal vFrm As FrmTransaksiCuti_Pengajuan_Setting, ByVal vUserCode As String) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            SW.Start()

            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, Core_Module.ClsAntrian.EnumTipeAntrian.AntrianTransaksiHRPayroll, _
            Me.vPrimary_Antrian, Me.mPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiLeaveSetting_SimpanBaru, vUserCode)

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
                            Me.HapusSetting()
                            Me.SimpanSetting(vFrm)

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
            Windows.Forms.Cursor.Current = Cursors.Default
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            SW = Nothing
        End Try
    End Function

    Public Function HapusSetting() As Boolean
        Dim Dta As New DseTrans_Cuti_PengajuanTableAdapters.tbltrans_cuti_pengajuan_settingTableAdapter
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            If Dta.Delete() > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Function SimpanSetting(ByVal vFrm As FrmTransaksiCuti_Pengajuan_Setting) As Boolean
        Dim Dta As New DseTrans_Cuti_PengajuanTableAdapters.tbltrans_cuti_pengajuan_settingTableAdapter
        Dim row As DataGridViewRow
        Dim CboDgv As DataGridViewComboBoxCell
        Dim remark As String = ""
        Dim privilage As ClsTrans_Cuti_Pengajuan.EnumPrivilagePengajuanCuti = EnumPrivilagePengajuanCuti.All_Employee
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            For Each row In vFrm.DgvItem.Rows
                If row.Cells(vFrm.ColUserCode.Index).Value <> "" Then
                    CboDgv = row.Cells(vFrm.ColPrivilage.Index)
                    If String.IsNullOrEmpty(row.Cells(vFrm.ColRemarks.Index).Value) = False Then
                        remark = row.Cells(vFrm.ColRemarks.Index).Value
                    Else
                        remark = ""
                    End If
                    Dta.Insert(row.Cells(vFrm.ColPrimaryUser.Index).Value, row.Cells(vFrm.ColPrimaryEmployee.Index).Value, _
                    CboDgv.Items.IndexOf(row.Cells(vFrm.ColPrivilage.Index).Value), remark)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Overridable Function SimpanBaru(ByVal vUserCode As String, ByVal FrmExplain As FrmTransaksiCuti_Pengajuan_ExplainDuration, _
    ByVal vDGV As DataGridView, ByVal vTblList_StatusExplain As DseTrans_Cuti_Pengajuan.List_StatusExplainDataTable, ByVal vFrm As FrmTransaksiCuti_Pengajuan) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            SW.Start()

            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, Core_Module.ClsAntrian.EnumTipeAntrian.AntrianTransaksiHRPayroll, _
            Me.vPrimary_Antrian, Me.mPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiLeaveApplication_SimpanBaru, vUserCode)

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
                            If Me.ApakahDataBenar(vUserCode) = True Then
                                retval = Me.SimpanBaru1(vUserCode, FrmExplain, vDGV, vTblList_StatusExplain, vFrm)
                            Else
                                retval = False
                            End If
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

    Public Function SimpanBaru1(ByVal vUserCode As String, ByVal vFrmExplain As FrmTransaksiCuti_Pengajuan_ExplainDuration, _
    ByVal vDGV As DataGridView, ByVal vTblList_StatusExplain As DseTrans_Cuti_Pengajuan.List_StatusExplainDataTable, ByVal vFrm As FrmTransaksiCuti_Pengajuan) As Boolean
        Dim DTA As New DseTrans_Cuti_PengajuanTableAdapters.tbltrans_cuti_pengajuanTableAdapter
        Dim ObjItem As New ClsTrans_Cuti_Pengajuan_Item
        Dim ObjCuti As New Core_Module.ClsMasterCuti
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Me.mUrutuan = Me.mNoTransaksi.Substring(0, 6)
            If DTA.Insert(Me.NoTransaksi, Me.mUrutuan, Me.mTanggal, Me.mPrimaryKaryawan, Me.mPrimaryJob, Me.mPrimaryCuti, _
             Me.mStatus, Me.mStatusNote, Me.mAppBy, Me.mAppDate, Me.mStartDate, Me.mHalfDayStart, Me.mEndDAte, Me.mHalfDayEnd, _
             Me.mDuration, Me.mReason, vUserCode, Me.mCreatedDate, vUserCode, Me.mModifiedDate, Me.mTimestamp) > 0 Then
                Me.mPrimary = DTA.ScalarQueryGetPrimaryByNoTransaksi(Me.mNoTransaksi)
                Me.mTimestamp = DTA.ScalarQueryGetTimeStampByPrimary(Me.mPrimary)
                Me.mCreatedDate = Me.mTimestamp
                Me.mModifiedDate = Me.mTimestamp
                Me.mCreatedBy = vUserCode
                Me.mModifiedBy = vUserCode

                'simpan item
                ObjItem.SimpanItem(vDGV, Me.mPrimary, vFrmExplain, vTblList_StatusExplain, vFrm)

                'Panggil fungsi set balance by primary karyawan dan primary cuti
                ObjCuti.SetBalance(Me.mPrimaryKaryawan, Me.mPrimaryCuti, Me.mStartDate)

                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjItem = Nothing
            ObjCuti = Nothing
        End Try
    End Function

    Public Overridable Function SimpanPerubahan(ByVal vUserCode As String, ByVal vFrmExplain As FrmTransaksiCuti_Pengajuan_ExplainDuration, _
    ByVal vDGV As DataGridView, ByVal vTblList_StatusExplain As DseTrans_Cuti_Pengajuan.List_StatusExplainDataTable, ByVal vFrm As FrmTransaksiCuti_Pengajuan) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            SW.Start()

            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, Core_Module.ClsAntrian.EnumTipeAntrian.AntrianTransaksiHRPayroll, _
            Me.vPrimary_Antrian, Me.mPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiLeaveApplication_SimpanPerubahan, vUserCode)

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
                            If Me.ApakahDataBenar(vUserCode, False) Then
                                retval = Me.SimpanPerubahan1(vUserCode, vFrmExplain, vDGV, vTblList_StatusExplain, vFrm)
                            Else
                                retval = False
                            End If
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

    Public Function SimpanPerubahan1(ByVal vUserCode As String, ByVal vFrmExplain As FrmTransaksiCuti_Pengajuan_ExplainDuration, _
    ByVal vDGV As DataGridView, ByVal vTblList_StatusExplain As DseTrans_Cuti_Pengajuan.List_StatusExplainDataTable, ByVal vFrm As FrmTransaksiCuti_Pengajuan) As Boolean
        Dim DTA As New DseTrans_Cuti_PengajuanTableAdapters.tbltrans_cuti_pengajuanTableAdapter
        Dim ObjItem As New ClsTrans_Cuti_Pengajuan_Item
        Dim ObjCuti As New Core_Module.ClsMasterCuti
        Dim PrimaryKaryawanLama As Double = -1
        Dim PrimaryCutiLama As Double = -1
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            ' Ambil dulu get primary karyawan dengan primary cuti yang lama 
            PrimaryKaryawanLama = DTA.ScalarQuery_GetPrimaryKaryawanByPrimary(Me.mPrimary)
            PrimaryCutiLama = DTA.ScalarQuery_GetPrimaryCutiByPrimary(Me.mPrimary)
            Me.mUrutuan = Me.mNoTransaksi.Substring(0, 6)

            If DTA.Update(Me.NoTransaksi, Me.mUrutuan, Me.mTanggal, Me.mPrimaryKaryawan, Me.mPrimaryJob, Me.mPrimaryCuti, _
             Me.mStatus, Me.mStatusNote, Me.mAppBy, Me.mAppDate, Me.mStartDate, Me.mHalfDayStart, Me.mEndDAte, Me.mHalfDayEnd, _
             Me.mDuration, Me.mReason, Me.mModifiedBy, Now, Me.mTimestamp, Me.mPrimary, Me.mTimestamp) > 0 Then
                Me.mTimestamp = DTA.ScalarQueryGetTimeStampByPrimary(Me.mPrimary)
                Me.mModifiedDate = Me.mTimestamp
                Me.mModifiedBy = vUserCode

                'hapus item
                ObjItem.HapusItem(Me.mPrimary, vFrmExplain, vTblList_StatusExplain, vFrm)

                'Panggil Fungsi set Balance By Primary Karyawan lama dan Primary cuti lama
                ObjCuti.SetBalance(PrimaryKaryawanLama, PrimaryKaryawanLama, Me.mStartDate)

                'simpan item
                ObjItem.SimpanItem(vDGV, Me.mPrimary, vFrmExplain, vTblList_StatusExplain, vFrm)

                'Panggil Fungsi set Balance By Primary Karyawan Baru dan Primary cuti Baru
                ObjCuti.SetBalance(Me.mPrimaryKaryawan, Me.mPrimaryCuti, Me.mStartDate)

                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjItem = Nothing
            ObjCuti = Nothing
        End Try
    End Function

    Public Function Hapus(ByVal vPrimaryTransaksi As Double, ByVal vUser As String) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            SW.Start()

            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, Core_Module.ClsAntrian.EnumTipeAntrian.AntrianTransaksiHRPayroll, _
            Me.vPrimary_Antrian, Me.mPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiLeaveApplication_Hapus, vUser)

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
                            If Me.ApakahBolehHapus(vUser) = True Then
                                retval = Me.Hapus1(vPrimaryTransaksi, vUser)
                            Else
                                retval = False
                            End If
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
        Dim DTA As New DseTrans_Cuti_PengajuanTableAdapters.tbltrans_cuti_pengajuanTableAdapter
        Dim vPrimaryKary, vPrimaryCuti As Double
        Dim RetVal As Nullable(Of Double)
        Dim ObjCuti As New Core_Module.ClsMasterCuti
        Dim Period As Date = Today
        Dim RetVal_date As Nullable(Of Date)
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'Get Primary Karyawan Lama dan Primary Cuti Lama 
            RetVal = DTA.ScalarQuery_GetPrimaryKaryawanByPrimary(vPrimary)
            If RetVal.HasValue = True Then
                vPrimaryKary = RetVal.Value
            End If
            RetVal = DTA.ScalarQuery_GetPrimaryCutiByPrimary(vPrimary)
            If RetVal.HasValue = True Then
                vPrimaryCuti = RetVal.Value
            End If
            RetVal_date = DTA.ScalarQuery_GetStartDate(vPrimary)
            If RetVal_date.HasValue = True Then
                Period = RetVal_date.Value
            End If

            'hapus            
            DTA.Delete(vPrimary)

            'Panggil set Balance By Primary Karyawan lama dan Primary Cuti Lama
            ObjCuti.SetBalance(vPrimaryKary, vPrimaryCuti, Period)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            DTA.Dispose()
            ObjCuti = Nothing
            ObjCD = Nothing
        End Try
    End Function

    Public Function GetNoBaru(ByVal vTanggalTransaksi As Date) As String
        Dim DTA As New DseTrans_Cuti_PengajuanTableAdapters.tbltrans_cuti_pengajuanTableAdapter
        Dim TanggalStart, TanggalEnd As Date
        Dim NoUrut As String = ""
        Dim ObjCD As New Core_Module.ClsCekData
        Dim dummy As Nullable(Of Double)

        Try
            'cari tanggal sampai dan dari
            TanggalEnd = ObjCD.GetLastDateOfTheYear(vTanggalTransaksi)
            TanggalStart = ObjCD.GetFisrtDateOfTheYear(vTanggalTransaksi)

            'cari no urut
            dummy = DTA.ScalarQuery_GetMaxUrutan(TanggalStart, TanggalEnd)
            If dummy.HasValue = True Then
                NoUrut = CStr(dummy.Value) + 1
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

            Return NoUrut + "/HR-LA/" + ObjCD.GetMonthRomawi(vTanggalTransaksi).ToUpper + "/" + vTanggalTransaksi.Year.ToString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Function CariDetail(ByVal vPrimary As Double, ByRef vFrmExplain As FrmTransaksiCuti_Pengajuan_ExplainDuration, _
   ByRef vDgv As DataGridView, ByRef vTblList_StatusExplain As DseTrans_Cuti_Pengajuan.List_StatusExplainDataTable, ByVal vFrm As FrmTransaksiCuti_Pengajuan) As Boolean
        Dim DTA As New DseTrans_Cuti_PengajuanTableAdapters.tbltrans_cuti_pengajuanTableAdapter
        Dim TBL As New DseTrans_Cuti_Pengajuan.tbltrans_cuti_pengajuanDataTable
        Dim row As DseTrans_Cuti_Pengajuan.tbltrans_cuti_pengajuanRow
        Dim ObjItem As New ClsTrans_Cuti_Pengajuan_Item
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillByPrimary(TBL, vPrimary)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = .primary
                    Me.mNoTransaksi = .no_transaksi
                    Me.mUrutuan = .urutan
                    Me.mTanggal = .tanggal
                    Me.mPrimaryKaryawan = .primary_karyawan
                    Me.mPrimaryJob = .primary_job_position
                    Me.mPrimaryCuti = .primary_cuti
                    Me.mStatus = .status
                    Me.mStatusNote = .status_notes
                    Me.mAppBy = .approved_by
                    Me.mAppDate = .approved_date
                    Me.mStartDate = .start_date
                    Me.mHalfDayStart = .half_day_start
                    Me.mEndDAte = .end_date
                    Me.mHalfDayEnd = .half_day_end
                    Me.mDuration = .duration
                    Me.mReason = .reason
                    Me.mCreatedBy = .created_by
                    Me.mCreatedDate = .created_date
                    Me.mModifiedBy = .modified_by
                    Me.mModifiedDate = .modified_date
                    Me.mTimestamp = .timestamp

                    ObjItem.IsiDGVItem(vDgv, .primary, vFrmExplain, vTblList_StatusExplain, vFrm)
                End With
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Function IsiDGV(ByVal Criteria As String, ByVal vDGV As DataGridView, ByVal TBL As DseTrans_Cuti_Pengajuan.tbltrancutipengajuan_viewDataTable, _
    ByVal SearchCriteria As EnumSreach, ByVal tsl As ToolStripLabel, ByVal vTanggalDari As Date, ByVal vTanggalSampai As Date, _
    ByVal vAllData As Boolean, ByVal vPrimaryKaryawan As Double) As Boolean
        Dim DTA As New DseTrans_Cuti_PengajuanTableAdapters.tbltrancutipengajuan_viewTableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If vAllData = True Then
                If SearchCriteria = EnumSreach.DocumentNo Then
                    QueryResult = DTA.FillByLikeDocumentNo(TBL, Criteria + "%")
                ElseIf SearchCriteria = EnumSreach._Date Then
                    QueryResult = DTA.FillByDate(TBL, vTanggalDari.Date, vTanggalSampai.Date)
                ElseIf SearchCriteria = EnumSreach.EmployeeCode Then
                    QueryResult = DTA.FillByLikeEmployee(TBL, Criteria + "%")
                ElseIf SearchCriteria = EnumSreach.LeaveCode Then
                    QueryResult = DTA.FillByLikeLeave(TBL, Criteria + "%")
                End If
            Else
                If SearchCriteria = EnumSreach.DocumentNo Then
                    QueryResult = DTA.FillByLikeDocumentNoNPrimaryKaryawan(TBL, Criteria + "%", vPrimaryKaryawan)
                ElseIf SearchCriteria = EnumSreach._Date Then
                    QueryResult = DTA.FillByDateNPrimaryKaryawan(TBL, vTanggalDari.Date, vTanggalSampai.Date, vPrimaryKaryawan)
                ElseIf SearchCriteria = EnumSreach.EmployeeCode Then
                    QueryResult = DTA.FillByLikeEmployeeNPrimaryKaryawan(TBL, Criteria + "%", vPrimaryKaryawan)
                ElseIf SearchCriteria = EnumSreach.LeaveCode Then
                    QueryResult = DTA.FillByLikeLeaveNPrimaryKaryawan(TBL, Criteria + "%", vPrimaryKaryawan)
                End If
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

    Public Function GetDuration(ByVal vPrimaryTransaksi As Double, ByVal vPrimaryKaryawan As Double, ByVal vPrimaryCuti As Double) As Decimal
        Dim DTA As New DseTrans_Cuti_PengajuanTableAdapters.tbltrans_cuti_pengajuan_itemTableAdapter
        Dim Retval As Nullable(Of Decimal)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetJml(vPrimaryKaryawan, vPrimaryCuti, vPrimaryTransaksi)
            If Retval.HasValue = True Then
                Return Retval.Value
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

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As _
    EnumSreach, ByVal vTanggalDari As Date, ByVal vTanggalSampai As Date, _
    ByVal vAlldata As Boolean, ByVal vPrimaryKaryawan As Double) As Double
        Dim DTA As New DseTrans_Cuti_PengajuanTableAdapters.tbltrancutipengajuan_viewTableAdapter
        Dim Dummy As Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If vAlldata = True Then
                If vSearchBy = EnumSreach.DocumentNo Then
                    Dummy = DTA.ScalarQueryGetJmlByLikeDocumentNo(vKriteria + "%")
                    If Dummy.HasValue = True Then
                        Return Dummy.Value
                    Else
                        Return 0
                    End If
                ElseIf vSearchBy = EnumSreach._Date Then
                    Dummy = DTA.ScalarQueryGetJmlByDate(vTanggalDari, vTanggalSampai)
                    If Dummy.HasValue = True Then
                        Return Dummy.Value
                    Else
                        Return 0
                    End If
                ElseIf vSearchBy = EnumSreach.EmployeeCode Then
                    Dummy = DTA.ScalarQueryGetJmlByLikeEmployee(vKriteria + "%")
                    If Dummy.HasValue = True Then
                        Return Dummy.Value
                    Else
                        Return 0
                    End If
                ElseIf vSearchBy = EnumSreach.LeaveCode Then
                    Dummy = DTA.ScalarQueryGetJmlByLikeLeave(vKriteria + "%")
                    If Dummy.HasValue = True Then
                        Return Dummy.Value
                    Else
                        Return 0
                    End If
                End If
            ElseIf vAlldata = False Then
                If vSearchBy = EnumSreach.DocumentNo Then
                    Dummy = DTA.ScalarQueryGetJmlByDocNoNPrimaryKaryawan(vKriteria + "%", vPrimaryKaryawan)
                    If Dummy.HasValue = True Then
                        Return Dummy.Value
                    Else
                        Return 0
                    End If
                ElseIf vSearchBy = EnumSreach._Date Then
                    Dummy = DTA.ScalarQueryGetJmlByDateNPrimaryKaryawan(vTanggalDari, vTanggalSampai, vPrimaryKaryawan)
                    If Dummy.HasValue = True Then
                        Return Dummy.Value
                    Else
                        Return 0
                    End If
                ElseIf vSearchBy = EnumSreach.EmployeeCode Then
                    Dummy = DTA.ScalarQueryGetJmlByPrimaryKaryawan(vKriteria + "%", vPrimaryKaryawan)
                    If Dummy.HasValue = True Then
                        Return Dummy.Value
                    Else
                        Return 0
                    End If
                ElseIf vSearchBy = EnumSreach.LeaveCode Then
                    Dummy = DTA.ScalarQueryGetJmlByLikeLeaveNPrimaryKaryawan(vKriteria + "%", vPrimaryKaryawan)
                    If Dummy.HasValue = True Then
                        Return Dummy.Value
                    Else
                        Return 0
                    End If
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryJob(ByVal vCode As String) As Double
        Dim DTA As New DseTrans_Cuti_PengajuanTableAdapters.tblmaster_cutiTableAdapter
        Dim Dummy As Nullable(Of Double)
        Try
            Dummy = DTA.ScalarQuery_GetPrimaryByCodeDesc(vCode)
            If Dummy.HasValue = True Then
                Return Dummy
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Public Function GetStatusNotes(ByVal vPrimary As Double) As String
        Dim DTA As New DseTrans_Cuti_PengajuanTableAdapters.tbltrans_cuti_pengajuanTableAdapter
        Dim Dummy As String = ""

        Try
            Dummy = DTA.ScalarQuery_GetStatusNotes(vPrimary)
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

    Public Function GetStatus(ByVal vPrimary As String) As Integer
        Dim DTA As New DseTrans_Cuti_PengajuanTableAdapters.tbltrans_cuti_pengajuanTableAdapter
        Dim Dummy As Nullable(Of Integer)
        Try
            Dummy = DTA.ScalarQuery_GetStatusByPrimary(vPrimary)
            If Dummy.HasValue = True Then
                Return Dummy
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Public Sub IsiCobLeave(ByVal vFrm As FrmTransaksiCuti_Pengajuan)
        Try
            vFrm.Tblmaster_cutiTableAdapter.Fill(vFrm.DseTrans_Cuti_Pengajuan.tblmaster_cuti, _
            vFrm.TxtPrimaryEmployee.Text, vFrm.ObjCekData.GetFisrtDateOfTheYear(vFrm.DtpDate.Value), vFrm.ObjCekData.GetLastDateOfTheYear(vFrm.DtpDate.Value))
            'vFrm.cboLeave.Refresh()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub HitungDurasi(ByVal vPrimaryEmployee As Double, ByVal vDateStart As Date, ByVal vDateEnd As Date, ByVal vTxtDuration As TextBox, _
    ByVal vChkStart As CheckBox, ByVal vChkEnd As CheckBox, ByVal vFrmExplain As FrmTransaksiCuti_Pengajuan_ExplainDuration, _
    ByRef vTblList As DseTrans_Cuti_Pengajuan.List_StatusExplainDataTable, ByVal vFrm As FrmTransaksiCuti_Pengajuan)
        Dim TblAntara As New DseTrans_Cuti_Pengajuan.TBLAntaraDataTable
        Dim rowAntara As DseTrans_Cuti_Pengajuan.TBLAntaraRow
        Dim DtaAllocate As New DseTrans_Cuti_PengajuanTableAdapters.tblmaster_workshift_itemallocateTableAdapter
        Dim TblAllocate As New DseTrans_Cuti_Pengajuan.tblmaster_workshift_itemallocateDataTable
        Dim rowAllocate As DseTrans_Cuti_Pengajuan.tblmaster_workshift_itemallocateRow
        Dim Filter As String = ""
        Dim DtaWeekly As New DseTrans_Cuti_PengajuanTableAdapters.TBLWeeklyTableAdapter
        Dim TBLWeekly As New DseTrans_Cuti_Pengajuan.TBLWeeklyDataTable
        Dim rowWeekly As DseTrans_Cuti_Pengajuan.TBLWeeklyRow
        Dim DtaHoliday As New DseTrans_Cuti_PengajuanTableAdapters.TBLHolidayListTableAdapter
        Dim TBLHoliday As New DseTrans_Cuti_Pengajuan.TBLHolidayListDataTable
        Dim duration As Decimal = 0
        Dim ObjCd As New Core_Module.ClsCekData
        Dim Objjob As New Core_Module.ClsMasterJobPosition
        Dim ObjWS As New Core_Module.ClsMasterWorkShift
        Dim BufferPrimary As Double = -1
        Dim BufferDate As Date = #1/1/1950#
        Dim _date As Date = #1/1/1950#
        Dim rowDgv As DataGridViewRow
        Dim vIndex As Integer = 0 '-
        Dim ObjItem As New ClsTrans_Cuti_Pengajuan_Item '-
        Try
            vFrmExplain.dgvExplain.Rows.Clear()
            'bikin baris sesuai durasi
            _date = vDateStart
            While (_date.Date <= vDateEnd.Date)
                'masukin tanggal, set primary workshift -1 dan durasi 0
                TblAntara.AddTBLAntaraRow(-1, _date, 0, True)
                _date = _date.AddDays(1)
            End While

            'kasih tanda jika bukan leave
            For i As Integer = TblAntara.Rows.Count - 1 To 0 Step -1
                rowAntara = TblAntara.Rows(i)
                vTblList.DefaultView.RowFilter = "[Date]='" & rowAntara.tanggal_startend & "' And [Status] <> 'Leave'"
                If vTblList.DefaultView.Count > 0 Then
                    rowAntara.is_leave = False
                End If
            Next

            'filter table workshift allocate
            DtaAllocate.FillBy_PrimaryEmployee(TblAllocate, vPrimaryEmployee)
            For Each rowAntara In TblAntara.Rows
                If rowAntara.is_leave = True Then
                    Filter = "[start date] <=#" + rowAntara.tanggal_startend.Date.ToString + _
                                    "# AND [end date] >=#" + rowAntara.tanggal_startend.Date.ToString + "#"
                    TblAllocate.DefaultView.RowFilter = Filter
                    If TblAllocate.DefaultView.Count > 0 Then
                        rowAllocate = TblAllocate.DefaultView(0).Row
                        'masukin nilai primary workshift
                        rowAntara.primary_workshift = rowAllocate.primary_workshift
                        If BufferPrimary = rowAllocate.primary_workshift And _
                        BufferDate = rowAntara.tanggal_startend Then
                            'jika ada primary workshift dan tanggalnya sama
                            'ada workshift yang double
                            MsgBox("Logic error, please contact your vendor if you see this message", MsgBoxStyle.Critical, "Error")
                        Else
                            'jika half day start di ceklist, durasi 0,5
                            'If rowAntara.tanggal_startend = vDateStart Then
                            '    If vChkStart.Checked = True Then
                            '        rowAntara.duration = 0.5
                            '        duration = duration + rowAntara.duration
                            '    End If
                            'End If
                            ''jika half end start di ceklist, durasi 0,5
                            'If rowAntara.tanggal_startend = vDateEnd Then
                            '    If vChkEnd.Checked = True Then
                            '        rowAntara.duration = 0.5
                            '        duration = duration + rowAntara.duration
                            '    End If
                            'End If
                            'jika half start dan end tidak di ceklist, arti nya colom durasi masih 0
                            'maka hitung durasi
                            If rowAntara.duration = 0 Then
                                DtaWeekly.FillBy_PrimaryWorkshiftNDay(TBLWeekly, rowAntara.primary_workshift, _
                                                        rowAntara.tanggal_startend.DayOfWeek)
                                rowWeekly = TBLWeekly.Rows(0)
                                If rowWeekly.tipe_workdays = Core_Module.ClsMasterWorkShift.TipeWorkdays.Holiday Then
                                    rowAntara.duration = 0
                                    duration = duration + rowAntara.duration
                                Else
                                    DtaHoliday.FillBy_PrimaryWorkshiftNDate(TBLHoliday, rowAntara.primary_workshift, _
                                    rowAntara.tanggal_startend)
                                    If TBLHoliday.Rows.Count > 0 Then
                                        rowAntara.duration = 0
                                        duration = duration + rowAntara.duration
                                    Else
                                        If rowWeekly.duratioin = Core_Module.ClsMasterWorkShift.enumDuration.fullday Then
                                            If rowAntara.tanggal_startend = vDateStart Then
                                                If vChkStart.Checked = True Then
                                                    rowAntara.duration = 0.5
                                                    duration = duration + rowAntara.duration
                                                Else
                                                    rowAntara.duration = 1
                                                    duration = duration + rowAntara.duration
                                                End If
                                            ElseIf rowAntara.tanggal_startend = vDateEnd Then
                                                If vChkEnd.Checked = True Then
                                                    rowAntara.duration = 0.5
                                                    duration = duration + rowAntara.duration
                                                Else
                                                    rowAntara.duration = 1
                                                    duration = duration + rowAntara.duration
                                                End If
                                            Else
                                                rowAntara.duration = 1
                                                duration = duration + rowAntara.duration
                                            End If
                                        Else
                                            rowAntara.duration = 0.5
                                            duration = duration + rowAntara.duration
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        BufferPrimary = rowAllocate.primary_workshift
                        BufferDate = rowAntara.tanggal_startend
                        'masukan ke dgv explain

                        vFrmExplain.dgvExplain.Rows.Add()
                        rowDgv = vFrmExplain.dgvExplain.Rows(vFrmExplain.dgvExplain.RowCount - 1)
                        rowDgv.Cells(vFrmExplain.ColDate.Index).Value = Format(rowAntara.tanggal_startend, "dd MMM yyyy")
                        rowDgv.Cells(vFrmExplain.ColDuration.Index).Value = ObjCd.DigitGrouping(rowAntara.duration, 1)
                        rowDgv.Cells(vFrmExplain.ColPrimaryWorkShift.Index).Value = rowAntara.primary_workshift
                        rowDgv.Cells(vFrmExplain.ColStatus.Index).Value = ObjItem.GetStatusExplain(vTblList, vFrmExplain, vIndex, vPrimaryEmployee, rowDgv.Cells(vFrmExplain.ColDate.Index).Value) '-
                        rowDgv.Cells(vFrmExplain.ColWorkShift.Index).Value = ObjWS.GetCodeByPrimary(rowAntara.primary_workshift)
                        vIndex += 1 '-
                    End If
                Else
                    Filter = "[start date] <=#" + rowAntara.tanggal_startend.Date.ToString + _
                                 "# AND [end date] >=#" + rowAntara.tanggal_startend.Date.ToString + "#"
                    TblAllocate.DefaultView.RowFilter = Filter
                    If TblAllocate.DefaultView.Count > 0 Then
                        rowAllocate = TblAllocate.DefaultView(0).Row
                        'masukin nilai primary workshift
                        rowAntara.primary_workshift = rowAllocate.primary_workshift
                    End If
                    vFrmExplain.dgvExplain.Rows.Add()
                    rowDgv = vFrmExplain.dgvExplain.Rows(vFrmExplain.dgvExplain.RowCount - 1)
                    rowDgv.Cells(vFrmExplain.ColDate.Index).Value = Format(rowAntara.tanggal_startend, "dd MMM yyyy")
                    rowDgv.Cells(vFrmExplain.ColDuration.Index).Value = ObjCd.DigitGrouping(rowAntara.duration, 1)
                    rowDgv.Cells(vFrmExplain.ColPrimaryWorkShift.Index).Value = rowAntara.primary_workshift
                    rowDgv.Cells(vFrmExplain.ColStatus.Index).Value = ObjItem.GetStatusExplain(vTblList, vFrmExplain, vIndex, vPrimaryEmployee, rowDgv.Cells(vFrmExplain.ColDate.Index).Value) '-
                    rowDgv.Cells(vFrmExplain.ColWorkShift.Index).Value = ObjWS.GetCodeByPrimary(rowAntara.primary_workshift)
                    vIndex += 1 '-
                    End If
            Next
            'vTblList.Rows.Clear()
            'ObjItem.Add_To_List(vTblList, vFrmExplain, vFrm)
            vTxtDuration.Text = ObjCd.DigitGrouping(duration, 1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        Finally
            DtaAllocate.Dispose()
            TblAllocate.Dispose()
            TblAntara.Dispose()
            DtaHoliday.Dispose()
            DtaWeekly.Dispose()
            TBLHoliday.Dispose()
            TBLWeekly.Dispose()
            ObjCd = Nothing
        End Try
    End Sub

    Public Function GetJmlUser(ByVal vPrimaryUser) As Double
        Dim DTA As New DseTrans_Cuti_PengajuanTableAdapters.tbltrans_cuti_pengajuan_settingTableAdapter
        Dim Dummy As Nullable(Of Long)
        Try
            Dummy = DTA.ScalarQuery_GetJmlUser(vPrimaryUser)
            If Dummy.HasValue = True Then
                Return Dummy
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Public Function GetPrivilegeUser(ByVal vPrimaryUser) As Integer
        Dim DTA As New DseTrans_Cuti_PengajuanTableAdapters.tbltrans_cuti_pengajuan_settingTableAdapter
        Dim Dummy As Nullable(Of Integer)
        Try
            Dummy = DTA.ScalarQuery_GetPrivilegeByUser(vPrimaryUser)
            If Dummy.HasValue = True Then
                Return Dummy
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Public Function GetPrimaryKaryawabByUser(ByVal vPrimaryUser) As Double
        Dim DTA As New DseTrans_Cuti_PengajuanTableAdapters.tbltrans_cuti_pengajuan_settingTableAdapter
        Dim Dummy As Nullable(Of Double)
        Try
            Dummy = DTA.ScalarQuery_GetPrimaryEmployeeByUser(vPrimaryUser)
            If Dummy.HasValue = True Then
                Return Dummy
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Public Sub IsiDGVsetting(ByVal vDgv As DataGridView, ByVal vFrm As FrmTransaksiCuti_Pengajuan_Setting)
        Dim DTA As New DseTrans_Cuti_PengajuanTableAdapters.tbltrans_cuti_pengajuan_settingTableAdapter
        Dim TBL As New DseTrans_Cuti_Pengajuan.tbltrans_cuti_pengajuan_settingDataTable
        Dim row As DseTrans_Cuti_Pengajuan.tbltrans_cuti_pengajuan_settingRow
        Dim RowDgv As DataGridViewRow
        Dim ObjKaryawan As New Core_Module.ClsMasterKaryawan
        Dim ObjUser As New Core_Module.ClsUser
        Dim CboDgv As DataGridViewComboBoxCell
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DTA.Fill(TBL)
            If TBL.Rows.Count > 0 Then
                For Each row In TBL.Rows
                    vDgv.Rows.Add()
                    RowDgv = vDgv.Rows(vDgv.RowCount - 2)
                    RowDgv.Cells(vFrm.ColEmployeeCode.Index).Value = ObjKaryawan.GetCodeByPrimary(row.primary_employee)
                    RowDgv.Cells(vFrm.ColEmployeeName.Index).Value = ObjKaryawan.GetFullNameByPrimary(row.primary_employee)
                    RowDgv.Cells(vFrm.ColPrimaryEmployee.Index).Value = row.primary_employee
                    RowDgv.Cells(vFrm.ColPrimaryUser.Index).Value = row.primary_user
                    CboDgv = RowDgv.Cells(vFrm.ColPrivilage.Index)
                    RowDgv.Cells(vFrm.ColPrivilage.Index).Value = CboDgv.Items.Item(row.privilage)
                    RowDgv.Cells(vFrm.ColRemarks.Index).Value = row.remarks
                    RowDgv.Cells(vFrm.ColUserCode.Index).Value = ObjUser.GetKodeUser(row.primary_user)
                    RowDgv.Cells(vFrm.ColUserName.Index).Value = ObjUser.GetNamaByPrimary(row.primary_user)
                Next
            End If
            Windows.Forms.Cursor.Current = Cursors.Default
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjKaryawan = Nothing
            ObjUser = Nothing
        End Try
    End Sub

    Public Function CekDurasi(ByVal vPrimaryKaryawan As Double, ByVal vPrimaryLeave As Double, _
    ByVal vStrDate As Date, ByVal vEndDate As Date, ByVal vBaru As Boolean, ByVal vPrimary As Double) As Boolean
        Dim dta As New DseTrans_Cuti_PengajuanTableAdapters.tbltrancutipengajuan_viewTableAdapter
        Dim tbl As New DseTrans_Cuti_Pengajuan.tbltrancutipengajuan_viewDataTable
        Dim row As DseTrans_Cuti_Pengajuan.tbltrancutipengajuan_viewRow
        Dim ListStr As String = ""
        Dim i As Integer = 0
        Try
            If vBaru = True Then
                dta.FillBy_PrimaryEmpNPrimaryLeaveNDate(tbl, vPrimaryLeave, vPrimaryKaryawan, vStrDate, vEndDate)
            Else
                dta.FillBy_CekDurasiEdit(tbl, vPrimaryLeave, vPrimaryKaryawan, vStrDate, vEndDate, vPrimary)
            End If

            If tbl.Rows.Count > 0 Then
                i = 1
                For Each row In tbl.Rows
                    ListStr = ListStr + i.ToString + ". " + row.no_transaksi + " : " + Format(row.start_date, "dd MMM yyyy") + " - " + Format(row.end_date, "dd MMM yyyy") + vbNewLine
                    i = i + 1
                Next
                MsgBox("You can't create leave application from " + Format(vStrDate, "dd MMM yyyy") + " to " + Format(vEndDate, "dd MMM yyyy") + " because it's used for another leave application." + _
                vbNewLine + "The Following is a list of leave application that have been used : " + vbNewLine + ListStr, MsgBoxStyle.Information, "Error")
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

End Class
