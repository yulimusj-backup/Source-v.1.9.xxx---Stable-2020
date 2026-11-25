Option Explicit On
Option Strict Off

Public Class ClsTrans_Attendance
    Dim mPrimary As Double = -1
    Dim mNoTrans As String = ""
    Dim mTanggal As Date = #1/1/1950#
    Dim mPrimaryEmp As Double = -1
    Dim mPrimaryPosition As Double = -1
    Dim mStrDate As Date = #1/1/1950#
    Dim mEndDate As Date = #1/1/1950#
    Dim mRemarks As String = ""
    Dim mCreatedBy As String = ""
    Dim mCreatedDate As DateTime = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As DateTime = Now
    Dim mTimestamp As Date = Now

    Dim ObjCd As New Core_Module.ClsCekData
    Dim ObjWs As New Core_Module.ClsMasterWorkShift
    Dim ObjItem As New ClsTrans_Attendance_Item
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting
    Dim ObjCuti As New Core_Module.ClsMasterCuti
    Dim ObjCuti_App As New ClsTrans_Cuti_Pengajuan

    'digunakan untuk antrian simpan data
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianTransaksiHRPayroll
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAAntrian As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter
    Dim ObjRnd As New Random
    Dim ObjAntrian As New Core_Module.ClsAntrian

    Public Enum EnumRemarks As Integer
        NO_ACTUAL_IN
        NO_ACTUAL_OUT
        LEAVE_BUT_PRESENT
    End Enum

#Region "Property"
    Public Property Primary() As Double
        Get
            Return mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property

    Public Property NoTrans() As String
        Get
            Return mNoTrans
        End Get
        Set(ByVal value As String)
            mNoTrans = value
        End Set
    End Property

    Public Property Tanggal() As Date
        Get
            Return mTanggal
        End Get
        Set(ByVal value As Date)
            mTanggal = value
        End Set
    End Property

    Public Property PrimaryEmp() As Double
        Get
            Return mPrimaryEmp
        End Get
        Set(ByVal value As Double)
            mPrimaryEmp = value
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

    Public Property StrDate() As Date
        Get
            Return mStrDate
        End Get
        Set(ByVal value As Date)
            mStrDate = value
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

    Public Property Remarks() As String
        Get
            Return mRemarks
        End Get
        Set(ByVal value As String)
            mRemarks = value
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

    Enum EnumWorkday
        Workday
        Holiday
    End Enum

    Enum EnumDurasi
        Full_Day
        Half_Day
    End Enum

    Enum EnumStatus
        Present
        Leave
        Half_day_Leave
        Sick
        Half_Day_Sick '4
        Permit
        Half_Day_Permit
        Absent
        Half_Absent
        Holiday
        Half_Holiday
    End Enum

    Enum EnumTipeOt
        No_Overtime
        Weekday_Overtime
        Holiday_Overtime
    End Enum

    Public Enum EnumSearch As Integer
        DocNo
        Tgl
        EmpCode
        EmpName
        Remarks
    End Enum

    Public Function ApakahSudahAda(ByVal vNoTrans As String) As Boolean
        Dim dta As New DseTrans_AttendanceTableAdapters.tbltrans_attendanceTableAdapter
        Dim dummy As Nullable(Of Long)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dummy = dta.ScalarQuery_GetJmlByNoTrans(vNoTrans)
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

    Public Overridable Function ApakahBolehEdit(ByVal vPrimary As Double, ByVal vUser As String, _
    Optional ByVal HanyaDipanggilKetikaQuickAttd As Boolean = False, Optional ByVal ShowMsgBox As Boolean = True) As Boolean
        Dim dta As New DseTrans_AttendanceTableAdapters.tbltrans_attendanceTableAdapter
        Dim DtaSalary As New DseTrans_AttendanceTableAdapters.DataTable3TableAdapter
        Dim TblSalary As New DseTrans_Attendance.DataTable3DataTable
        Dim DummyUser As String
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            DtaSalary.Connection.ConnectionString = Core_Module.GetConstring

            If HanyaDipanggilKetikaQuickAttd = False Then
                'strict edit
                DummyUser = dta.ScalarQuery_GetCreatedBy(vPrimary)
                If DummyUser <> vUser Then
                    If ShowMsgBox = True Then
                        MsgBox("Only user: " + DummyUser + " is allowed to edit this transaction.", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If

            'cek disalary
            DtaSalary.FillBy_PrimayEmpNDate(TblSalary, Me.PrimaryEmp, Me.EndDate, Me.StrDate)
            If TblSalary.Rows.Count > 0 Then
                If ShowMsgBox = True Then
                    MsgBox("Attendance data is refered by salary calculation, you can't edit this data.", _
                    MsgBoxStyle.Information, "information")
                End If
                Return False
            End If

            'Cek jika ada perubahan employee code, jika sudah ada salary maka tidak boleh edit
            Dim PrimaryEmpDatabase As Double = Me.GetPrimaryEmployeeAttendance(Me.mPrimary)
            DtaSalary.FillBy_PrimayEmpNDate(TblSalary, PrimaryEmpDatabase, Me.EndDate, Me.StrDate)
            If TblSalary.Rows.Count > 0 Then
                If ShowMsgBox = True Then
                    MsgBox("Attendance data is refered by salary calculation, you can't change employee code.", _
                    MsgBoxStyle.Information, "information")
                End If
                Return False
            End If

            'benar
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
            DtaSalary.Dispose()
            TblSalary.Dispose()
        End Try
    End Function

    Public Function GetPrimaryEmployeeAttendance(ByVal vPrimaryAttendance As Double) As Double
        Dim DTA As New DseTrans_AttendanceTableAdapters.tbltrans_attendanceTableAdapter
        Dim Dummy As System.Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetPrimaryKaryawanByPrimary(vPrimaryAttendance)
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

    Public Overridable Function ApakahBolehHapus(ByVal vPrimary As Double, ByVal vUser As String) As Boolean
        Dim dta As New DseTrans_AttendanceTableAdapters.tbltrans_attendanceTableAdapter
        Dim DtaSalary As New DseTrans_AttendanceTableAdapters.DataTable3TableAdapter
        Dim TblSalary As New DseTrans_Attendance.DataTable3DataTable
        Dim DummyUser As String = ""

        Try
            'tambahan dari leave app setting
            dta.Connection.ConnectionString = Core_Module.GetConstring
            DtaSalary.Connection.ConnectionString = Core_Module.GetConstring

            'strict edit
            DummyUser = dta.ScalarQuery_GetCreatedBy(vPrimary)
            If DummyUser <> vUser Then
                MsgBox("Only user: " + DummyUser + " is allowed to delete this transaction.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            'cek disalary
            DtaSalary.FillBy_PrimayEmpNDate(TblSalary, Me.PrimaryEmp, Me.EndDate, Me.StrDate)
            If TblSalary.Rows.Count > 0 Then
                MsgBox("Attendance data is refered by salary calculation, you can't delete this data..", _
                MsgBoxStyle.Information, "information")
                Return False
            End If
            'benar
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
            DtaSalary.Dispose()
            TblSalary.Dispose()
        End Try
    End Function

    Public Function GetTanggalTransaksi(ByVal vPrimary As Double) As Date
        Dim DTA As New DseTrans_AttendanceTableAdapters.tbltrans_attendanceTableAdapter
        Dim Dummy As System.Nullable(Of Date)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetTanggalTransaksiAttendance(vPrimary)
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

    Public Function GetNoBaru(ByVal vTanggalTransaksi As Date) As String
        Dim DTA As New DseTrans_AttendanceTableAdapters.tbltrans_attendanceTableAdapter
        Dim TanggalStart, TanggalEnd As Date
        Dim NoUrut As String
        Dim JmlDb As Nullable(Of Long)
        Try
            'cari tanggal sampai dan dari
            TanggalEnd = Me.ObjCd.GetLastDateOfTheYear(vTanggalTransaksi)
            TanggalStart = Me.ObjCd.GetFisrtDateOfTheYear(vTanggalTransaksi)


            'cari no urut
            JmlDb = DTA.ScalarQuery_GetMaxUrutan(TanggalStart, TanggalEnd)
            If JmlDb.HasValue = True Then
                NoUrut = CStr(CInt(JmlDb.Value)) + 1
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
            Dim bln As String = ""
            bln = Me.ObjCd.GetMonthRomawi(vTanggalTransaksi)

            Return NoUrut + "/" + "ATT" + "/" + bln + "/" + CStr(vTanggalTransaksi.Year)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetNoTransByPrimary(ByVal vPrimary As Double) As String
        Dim dta As New DseTrans_AttendanceTableAdapters.tbltrans_attendanceTableAdapter
        Dim Dummy As String
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = dta.ScalarQuery_GetNoTransByPrimary(vPrimary)
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

    Public Function GetTimeStampByPrimary(ByVal vPrimary As Double) As Date
        Dim dta As New DseTrans_AttendanceTableAdapters.tbltrans_attendanceTableAdapter
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

    Public Function CekValidPeriodEmp(ByVal vPrimary As Double, ByVal vPrimaryEmp As Double, _
    ByVal vStrDate As Date, ByVal vEndDate As Date, ByVal vBaru As Boolean, _
    Optional ByVal vShowDialog As Boolean = True) As Boolean
        Dim dta As New DseTrans_AttendanceTableAdapters.tbltrans_attendanceTableAdapter
        Dim tbl As New DseTrans_Attendance.tbltrans_attendanceDataTable
        Dim row As DseTrans_Attendance.tbltrans_attendanceRow
        Dim ListPeriod As String = ""
        Dim i As Integer = 1
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If vBaru = True Then
                dta.FillBy_PeriodEmp(tbl, vEndDate, vStrDate, vPrimaryEmp)
            Else
                dta.FillBy_PeriodEmpEdit(tbl, vEndDate, vStrDate, vPrimaryEmp, vPrimary)
            End If

            If tbl.Rows.Count > 0 Then
                For Each row In tbl.Rows
                    ListPeriod = ListPeriod + i.ToString + ". " + row.no_transaksi.ToString + ", Period : " + _
                    Format(row.start_date, "dd MMM yyyy").ToString + " - " + Format(row.end_date, "dd MMM yyyy").ToString + vbNewLine
                    i = i + 1
                Next
                If vShowDialog = True Then
                    MsgBox("This period already used : " + vbNewLine + ListPeriod, MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            dta.Dispose()
            tbl.Dispose()
        End Try
    End Function

    Public Function CariDetail(ByVal vPrimary As Double, ByVal vFrm As FrmTrans_Attendance, _
    ByRef TblOT_Order As DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_ot_1DataTable, _
    ByVal vJoinWithAttPrev As Boolean) As Boolean
        Dim dta As New DseTrans_AttendanceTableAdapters.tbltrans_attendanceTableAdapter
        Dim TBL As New DseTrans_Attendance.tbltrans_attendanceDataTable
        Dim row As DseTrans_Attendance.tbltrans_attendanceRow
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dta.FillBy_Primary(TBL, vPrimary)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = .primary
                    Me.mNoTrans = .no_transaksi
                    Me.mTanggal = .tanggal
                    Me.mPrimaryEmp = .primary_karyawan
                    Me.mStrDate = .start_date
                    Me.mEndDate = .end_date
                    Me.Remarks = .remarks
                    Me.mCreatedBy = .created_by
                    Me.mCreatedDate = .created_date
                    Me.mModifiedBy = .modified_by
                    Me.mModifiedDate = .modified_date
                    Me.mTimestamp = .timestamp

                    Me.IsiTblSpl(vJoinWithAttPrev, .primary_karyawan, vFrm.DtpStartAttPrev, .start_date, .end_date, TblOT_Order)
                    Me.ObjItem.IsiDgv(vPrimary, vFrm, TblOT_Order)
                End With
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "error")
        Finally
            dta.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function CekHalfStatus(ByVal vStatus As ClsTrans_Attendance.EnumStatus, Optional ByVal vPresent As Boolean = False) As Boolean
        Try
            If vStatus = ClsTrans_Attendance.EnumStatus.Half_Absent Then
                Return True
            ElseIf vStatus = ClsTrans_Attendance.EnumStatus.Half_day_Leave Then
                Return True
            ElseIf vStatus = ClsTrans_Attendance.EnumStatus.Half_Day_Permit Then
                Return True
            ElseIf vStatus = ClsTrans_Attendance.EnumStatus.Half_Day_Sick Then
                Return True
            ElseIf vStatus = ClsTrans_Attendance.EnumStatus.Half_Holiday Then
                Return True
            ElseIf vPresent = True Then
                If vStatus = ClsTrans_Attendance.EnumStatus.Present Then
                    Return True
                End If
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function CekValidStatus(ByVal workday As EnumWorkday, ByVal durasi As EnumDurasi, ByVal status As EnumStatus) As Boolean
        Try
            If workday = EnumWorkday.Workday And durasi = EnumDurasi.Half_Day Then
                If Me.CekHalfStatus(status, True) = False Then
                    MsgBox("Status allowed is : " + vbNewLine + _
                           "1. Present" + vbNewLine + _
                           "2. Halfday Leave" + vbNewLine + _
                           "3. Halfday Sick" + vbNewLine + _
                           "4. Halfday Permit" + vbNewLine + _
                           "5. Halfday Absent" + vbNewLine + _
                           "6. Halfday Holiday" + vbNewLine, MsgBoxStyle.Information, "Information")
                    Return False
                End If
            ElseIf workday = EnumWorkday.Holiday And durasi = EnumDurasi.Full_Day Then
                If status = EnumStatus.Present Or status = EnumStatus.Holiday Then
                    Return True
                End If
                MsgBox("Status allowed is Present or Holiday status", MsgBoxStyle.Information, "Information")
                Return False
            ElseIf workday = EnumWorkday.Holiday And durasi = EnumDurasi.Half_Day Then
                If status = EnumStatus.Present Or status = EnumStatus.Half_Holiday Then
                    Return True
                End If
                MsgBox("Status allowed is Present or Halfday holiday status", MsgBoxStyle.Information, "Information")
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Sub ReloadData(ByVal vFrm As FrmTrans_Attendance, ByVal vJoinWithAttPrev As Boolean, _
    ByRef TblOT_Order As DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_ot_1DataTable, _
    ByRef DtpStartAttNew As Date)
        Dim Dta As New DseTrans_AttendanceTableAdapters.DataTable2TableAdapter
        Dim Tbl As New DseTrans_Attendance.DataTable2DataTable
        Dim Row As DseTrans_Attendance.DataTable2Row
        Dim DtaPengajuan As New DseTrans_AttendanceTableAdapters.tbltrans_cuti_pengajuanTableAdapter
        Dim CountDate As Integer = 1

        Dim TglAwal As Date = #1/1/1950#
        Dim TglAkhir As Date = #1/1/1950#

        Dim RowDgv As Windows.Forms.DataGridViewRow
        Dim ObjApp As New ClsTrans_Cuti_Pengajuan

        'aul tambahan
        Dim DtpStarDum As Date = vFrm.dtpStart.Value.Date
        Dim DtpToDum As Date = vFrm.DtpEnd.Value.Date
        Dim IndexDgvNew As Integer = 0

        Try
            'aul tambahan
            If vJoinWithAttPrev = True Then
                IndexDgvNew = vFrm.dgvItem.RowCount
                DtpStarDum = CDate(vFrm.dgvItem.Item(vFrm.ColDate.Index, IndexDgvNew - 1).Value).Date + TimeSpan.FromDays(1)
                If DtpStarDum.Date > DtpToDum.Date Then
                    For CountDate = 1 To (DtpStarDum.Date - (DtpToDum.Date + TimeSpan.FromDays(1)).Date).TotalDays
                        vFrm.dgvItem.Rows.RemoveAt(vFrm.dgvItem.RowCount - 1)
                    Next

                    Exit Sub
                End If
            End If
            DtpStartAttNew = DtpStarDum

            Dta.Fill(Tbl, vFrm.TxtPrimaryEmployee.Text, DtpStarDum.Date, DtpToDum.Date)
            If Tbl.Rows.Count > 0 Then
                For Each Row In Tbl.Rows
                    'cari waktu yang akan dimasukin ke dgv nya
                    If Row.start_date < DtpStarDum.Date And Row.end_date > DtpToDum.Date Then
                        TglAwal = DtpStarDum
                        TglAkhir = DtpToDum
                    ElseIf Row.start_date < DtpStarDum.Date And Row.end_date <= DtpToDum.Date Then
                        TglAwal = DtpStarDum
                        TglAkhir = Row.end_date
                    ElseIf Row.start_date >= DtpStarDum.Date And Row.end_date > DtpToDum.Date Then
                        TglAwal = Row.start_date
                        TglAkhir = DtpToDum
                    ElseIf Row.start_date >= DtpStarDum.Date And Row.end_date <= DtpToDum.Date Then
                        TglAwal = Row.start_date
                        TglAkhir = Row.end_date
                    End If
                    Me.IsiDgvItem(vFrm, TglAwal, TglAkhir, Row)
                Next
            Else
                MsgBox("Please allocate workshift at this period before reload data", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            'cari leave app yang blum di approved
            Dim DummyPrimary As New List(Of Double)
            Dim Status, i As Integer
            Dim ListWaiting As String = ""

            i = 1
            For Each RowDgv In vFrm.dgvItem.Rows
                'jika ada primary leave application
                If RowDgv.Cells(vFrm.ColPrimaryLA.Index).Value > 0 Then
                    'primary leave applicationnya ada d dalam list ga ??
                    If DummyPrimary.Contains(RowDgv.Cells(vFrm.ColPrimaryLA.Index).Value) = False Then
                        'status leave approved ??
                        'jika waiting masukin kedalam list yang belum d approved
                        Status = ObjApp.GetStatus(RowDgv.Cells(vFrm.ColPrimaryLA.Index).Value)
                        If Status = ClsTrans_Cuti_Pengajuan.EnumStatus.Waiting Then
                            ListWaiting += i.ToString + ". " + DtaPengajuan.ScalarQuery_GetListByPrimary(RowDgv.Cells(vFrm.ColPrimaryLA.Index).Value).ToString _
                            + vbNewLine
                            i += 1
                            DummyPrimary.Add(RowDgv.Cells(vFrm.ColPrimaryLA.Index).Value)
                        End If
                    End If
                End If
            Next
            'For IdxDgv As Integer = IndexDgvNew To vFrm.dgvItem.Rows.Count - 1
            '    RowDgv = vFrm.dgvItem.Rows(IdxDgv)
            '    'jika ada primary leave application
            '    If RowDgv.Cells(vFrm.ColPrimaryLA.Index).Value > 0 Then
            '        'primary leave applicationnya ada d dalam list ga ??
            '        If DummyPrimary.Contains(RowDgv.Cells(vFrm.ColPrimaryLA.Index).Value) = False Then
            '            'status leave approved ??
            '            'jika waiting masukin kedalam list yang belum d approved
            '            Status = ObjApp.GetStatus(RowDgv.Cells(vFrm.ColPrimaryLA.Index).Value)
            '            If Status = ClsTrans_Cuti_Pengajuan.EnumStatus.Waiting Then
            '                ListWaiting += i.ToString + ". " + DtaPengajuan.ScalarQuery_GetListByPrimary(RowDgv.Cells(vFrm.ColPrimaryLA.Index).Value).ToString _
            '                + vbNewLine
            '                i += 1
            '                DummyPrimary.Add(RowDgv.Cells(vFrm.ColPrimaryLA.Index).Value)
            '            End If
            '        End If
            '    End If
            'Next
            'kasih peringatan jika ada yg waiting
            If ListWaiting <> "" Then
                MsgBox("The following leave application must be approved before save :" + vbNewLine + _
                       ListWaiting, MsgBoxStyle.Information, "Information")
                vFrm.dgvItem.Rows.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Dta.Dispose()
            Tbl.Dispose()
            DtaPengajuan.Dispose()
            ObjApp = Nothing
        End Try
    End Sub

    Public Sub IsiTblSpl(ByVal vJoinWithAttPrev As Boolean, ByVal vPrimaryEmp As Double, ByVal vDateFromPrev As Date, ByVal vDateFrom As Date, _
    ByVal vDateTo As Date, ByRef TblOT_Order As DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_ot_1DataTable)
        Dim DtaOT_Order As New DseTrans_SPLTableAdapters.tbltrans_spl_overtime_order_item_approved_ot_1TableAdapter

        Try
            DtaOT_Order.Connection.ConnectionString = Core_Module.GetConstring
            If vJoinWithAttPrev = True Then
                DtaOT_Order.Fill(TblOT_Order, vPrimaryEmp, vDateFromPrev, vDateTo)
            Else
                DtaOT_Order.Fill(TblOT_Order, vPrimaryEmp, vDateFrom, vDateTo)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'aul tambah overtime order
    Public Sub UpdateAttBySpl(ByVal vPrimaryEmp As Double, ByVal vDateFrom As Date, ByVal vDateTo As Date, _
    ByRef TblOT_Order As DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_ot_1DataTable, ByVal vFrm As FrmTrans_Attendance)
        Dim DtaWS As New DseTrans_SPLTableAdapters.tblmaster_workshiftTableAdapter
        Dim TblWS As New DseTrans_SPL.tblmaster_workshiftDataTable
        Dim rowWS As DseTrans_SPL.tblmaster_workshiftRow
        Dim RowDgv As Windows.Forms.DataGridViewRow
        Dim OT_Dgv As Double = 0
        Dim EarlyOT_Dgv As Double = 0
        Dim ObjEmp As New Core_Module.ClsMasterKaryawan
        Dim ApaOverrideOTOrder_Emp As Boolean = ObjEmp.ApakahOverrideOTOrder_Emp(vPrimaryEmp)
        Dim ApaUseOTOrder_Emp As Boolean = ObjEmp.ApakahUseOTOrder_Emp(vPrimaryEmp)
        ObjEmp = Nothing
        Dim UpdateColDuration As Boolean = False
        Dim rowOT_Order As DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_ot_1Row

        Try
            DtaWS.Connection.ConnectionString = Core_Module.GetConstring
            DtaWS.Fill_all(TblWS, Core_Module.ClsMasterWorkShift.EnumTipeWorkshift.HR)

            'jika pada master karyawan di override dan tidak menggunakan use ot maka karyawan tsb dianggap tidak menggunakan spl
            If ApaOverrideOTOrder_Emp = True And ApaUseOTOrder_Emp = False Then
                Exit Sub
            End If

            'HANYA FILTER YANG STATUSNYA WAIT FOR APPROVAL & APPROVED SAJA.
            For Each RowDgv In vFrm.dgvItem.Rows
                UpdateColDuration = False

                If RowDgv.Cells(vFrm.ColPrimaryOvertimeOrder.Index).Value <> -1 Then
                    TblOT_Order.DefaultView.RowFilter = "[Date] = #" & RowDgv.Cells(vFrm.ColDate.Index).Value & "#"
                    If TblOT_Order.DefaultView.Count > 0 Then
                        rowOT_Order = TblOT_Order.DefaultView(0).Row
                        'isi status remarks/reason overtime order
                        RowDgv.Cells(vFrm.ColStatusRemark.Index).Value = rowOT_Order.Reason
                        'isi overtime order no
                        RowDgv.Cells(vFrm.ColOvertimeOrderNo.Index).Value = rowOT_Order.No_Transaksi
                    End If
                End If

                'hanya input berdasarkan periode from to
                If RowDgv.Cells(vFrm.ColDate.Index).Value >= vDateFrom And RowDgv.Cells(vFrm.ColDate.Index).Value <= vDateTo Then
                    'jika pada master karyawan di override ot order maka langsung mengacu ke master karyawan
                    If ApaOverrideOTOrder_Emp = True Then
                        UpdateColDuration = True
                    Else
                        TblWS.DefaultView.RowFilter = "[Primary] = " & RowDgv.Cells(vFrm.ColPrimaryWorkShift.Index).Value
                        If TblWS.DefaultView.Count > 0 Then
                            rowWS = TblWS.DefaultView(0).Row
                            'jika pada workshift yang tidak use ot order maka tidak usah di update by ot order, walaupun sudah di buat ot ordernya.
                            If rowWS.use_ot_order > 0 Then
                                UpdateColDuration = True
                            End If
                        End If
                    End If

                    If UpdateColDuration = True Then
                        'jika pada periode dimana menggunaka ot order namun belum di approve maka tidak bisa buat attendancenya.
                        If Me.UpdateAttBySpl_Detail(vFrm.TxtPrimaryEmployee.Text, RowDgv, TblOT_Order, vFrm, True) = False Then
                            vFrm.dgvItem.Rows.Clear()
                            Exit Sub
                        End If
                    Else
                        RowDgv.Cells(vFrm.ColPrimaryOvertimeOrder.Index).Value = -1
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'aul tambah overtime order detail
    Public Function UpdateAttBySpl_Detail(ByVal vPrimaryEmp As Double, ByVal RowDgv As Windows.Forms.DataGridViewRow, _
    ByRef TblOT_Order As DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_ot_1DataTable, _
    ByVal vFrm As FrmTrans_Attendance, ByVal vShowMsg As Boolean) As Boolean
        Dim rowOT_Order As DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_ot_1Row
        Dim OT_Dgv As Double = 0
        Dim EarlyOT_Dgv As Double = 0
        Dim EndTime_OT_Dgv As Double = 0
        Dim StartTime_EarlyOT_Dgv As Double = 0

        Dim CboDgv As Windows.Forms.DataGridViewComboBoxCell
        Dim workday As ClsTrans_Attendance.EnumWorkday
        Dim durasi As ClsTrans_Attendance.EnumDurasi
        Dim status As ClsTrans_Attendance.EnumStatus
        Dim CalAfterEOT As Double = 0
        Dim CalAfterOT As Double = 0
        Dim RTypeEOT As Integer = 0
        Dim RTypeOT As Integer = 0
        Dim RValueEOT As Decimal = 0
        Dim RValueOT As Decimal = 0
        Dim MValueEOT As Double = 0
        Dim MValueOT As Double = 0

        'YANG MASUK KE FUNGSI INI HANYA WORKSHIFT YANG MENGGUNAKAN USE OT ORDER SAJA.

        Try
            'jika ada primary leave application, maka overtime order dianggap ditak ada
            If RowDgv.Cells(vFrm.ColPrimaryLA.Index).Value = -1 Then
                'workday
                CboDgv = RowDgv.Cells(vFrm.ColWorkdays.Index)
                workday = CboDgv.Items.IndexOf(RowDgv.Cells(vFrm.ColWorkdays.Index).Value)
                'fullday / half day
                CboDgv = RowDgv.Cells(vFrm.ColDuration.Index)
                durasi = CboDgv.Items.IndexOf(RowDgv.Cells(vFrm.ColDuration.Index).Value)
                'status
                CboDgv = RowDgv.Cells(vFrm.ColAttendanceStatus.Index)
                status = CboDgv.Items.IndexOf(RowDgv.Cells(vFrm.ColAttendanceStatus.Index).Value)

                CalAfterEOT = Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(vFrm.ColCalculateAfterEarlyOT.Index).Value)
                CalAfterOT = Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(vFrm.ColCalculateAfterOT.Index).Value)

                CboDgv = RowDgv.Cells(vFrm.ColRoundingTypeEarlyOT.Index)
                RTypeEOT = CboDgv.Items.IndexOf(RowDgv.Cells(vFrm.ColRoundingTypeEarlyOT.Index).Value)
                RTypeOT = CboDgv.Items.IndexOf(RowDgv.Cells(vFrm.ColRoundingTypeOT.Index).Value)

                RValueEOT = Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(vFrm.ColRoundingValueEarlyOT.Index).Value)
                RValueOT = Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(vFrm.ColRoundingValueOT.Index).Value)

                MValueEOT = Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(vFrm.ColMaxValueEarlyOT.Index).Value)
                MValueOT = Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(vFrm.ColMaxValueOT.Index).Value)

                TblOT_Order.DefaultView.RowFilter = "[Date] = #" & CDate(RowDgv.Cells(vFrm.ColDate.Index).Value) & "#"
                If TblOT_Order.DefaultView.Count > 0 Then
                    rowOT_Order = TblOT_Order.DefaultView(0).Row
                    If rowOT_Order.Status = ClsTrans_SPL.EnumStatus.Approved Then
                        RowDgv.Cells(vFrm.ColPrimaryOvertimeOrder.Index).Value = rowOT_Order.Primary_Transaksi
                        RowDgv.Cells(vFrm.ColStatusRemark.Index).Value = rowOT_Order.Reason
                        RowDgv.Cells(vFrm.ColOvertimeOrderNo.Index).Value = rowOT_Order.No_Transaksi
                        RowDgv.Cells(vFrm.ColEarlyOTOrderDuration.Index).Value = Me.ObjCd.Convert_ToString(rowOT_Order.OT_Early_Approved)
                        RowDgv.Cells(vFrm.ColOTOrderDuration.Index).Value = Me.ObjCd.Convert_ToString(rowOT_Order.OT_Aprroved)

                        If Me.ObjCd.Is_Time_String(RowDgv.Cells(vFrm.ColActualIn.Index).Value.ToString) = True Or _
                        Me.ObjCd.Is_Time_String(RowDgv.Cells(vFrm.ColActualOut.Index).Value.ToString) = True Then
                            'update kolom early ot ################################################################################### 
                            EarlyOT_Dgv = Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(vFrm.ColEarlyOTDuration.Index).Value)
                            If EarlyOT_Dgv > 0 Then
                                StartTime_EarlyOT_Dgv = Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(vFrm.ColActualIn.Index).Value)
                                If StartTime_EarlyOT_Dgv >= rowOT_Order.End_Time_Early Then
                                    EarlyOT_Dgv = 0
                                ElseIf StartTime_EarlyOT_Dgv < rowOT_Order.End_Time_Early Then
                                    If rowOT_Order.Break_Out_Time_Early = 0 Or rowOT_Order.Break_In_Time_Early = 0 Then
                                        If StartTime_EarlyOT_Dgv < rowOT_Order.Start_Time_Early Then
                                            EarlyOT_Dgv = rowOT_Order.OT_Early_Approved
                                        End If
                                    Else
                                        If StartTime_EarlyOT_Dgv > rowOT_Order.Break_In_Time_Early Then
                                            EarlyOT_Dgv = rowOT_Order.End_Time_Early - StartTime_EarlyOT_Dgv
                                        ElseIf StartTime_EarlyOT_Dgv > rowOT_Order.Break_Out_Time_Early And StartTime_EarlyOT_Dgv <= rowOT_Order.Break_In_Time_Early Then
                                            EarlyOT_Dgv = rowOT_Order.End_Time_Early - rowOT_Order.Break_In_Time_Early
                                        ElseIf StartTime_EarlyOT_Dgv < rowOT_Order.Break_Out_Time_Early Then
                                            If StartTime_EarlyOT_Dgv > rowOT_Order.Start_Time_Early Then
                                                EarlyOT_Dgv = (rowOT_Order.Break_Out_Time_Early - StartTime_EarlyOT_Dgv) + (rowOT_Order.End_Time_Early - rowOT_Order.Break_In_Time_Early)
                                            ElseIf StartTime_EarlyOT_Dgv <= rowOT_Order.Start_Time_Early Then
                                                EarlyOT_Dgv = rowOT_Order.OT_Early_Approved
                                            End If
                                        End If
                                    End If
                                End If

                                'kurangi dengan tolerance
                                If RowDgv.Cells(vFrm.ColCalculateEarlyOTTolerance.Index).Value = 0 Then
                                    If EarlyOT_Dgv - Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(vFrm.ColCalculateAfterEarlyOT.Index).Value) > 0 Then
                                        EarlyOT_Dgv = EarlyOT_Dgv - Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(vFrm.ColCalculateAfterEarlyOT.Index).Value)
                                    End If
                                End If

                                'rounding early in
                                If RTypeEOT = Core_Module.ClsMasterWorkShift.enumRoundingType.No_Rounding Then
                                ElseIf RTypeEOT = Core_Module.ClsMasterWorkShift.enumRoundingType.Round_down Then
                                    EarlyOT_Dgv = Me.ObjCd.RoundDown2(CDec(EarlyOT_Dgv), RValueEOT)
                                ElseIf RTypeEOT = Core_Module.ClsMasterWorkShift.enumRoundingType.Round_Up Then
                                    EarlyOT_Dgv = Me.ObjCd.RoundUp2(CDec(EarlyOT_Dgv), RValueEOT)
                                ElseIf RTypeEOT = Core_Module.ClsMasterWorkShift.enumRoundingType.Middle Then
                                    EarlyOT_Dgv = Me.ObjCd.MiddleRound(RValueEOT, CDec(EarlyOT_Dgv))
                                End If

                                If MValueEOT > 0 Then
                                    If MValueEOT >= EarlyOT_Dgv Then
                                        RowDgv.Cells(vFrm.ColEarlyOTDuration.Index).Value = Me.ObjCd.Convert_ToString(EarlyOT_Dgv)
                                    ElseIf MValueEOT < EarlyOT_Dgv Then
                                        RowDgv.Cells(vFrm.ColEarlyOTDuration.Index).Value = Me.ObjCd.Convert_ToString(MValueEOT)
                                    End If
                                Else
                                    RowDgv.Cells(vFrm.ColEarlyOTDuration.Index).Value = Me.ObjCd.Convert_ToString(EarlyOT_Dgv)
                                End If

                                'cek type ot
                                CboDgv = RowDgv.Cells(vFrm.ColOT.Index)
                                If RowDgv.Cells(vFrm.ColOT.Index).Value <> CboDgv.Items.Item(ClsTrans_Attendance.EnumTipeOt.No_Overtime) Then
                                    If EarlyOT_Dgv > 0 Then
                                        RowDgv.Cells(vFrm.ColIsEarlyOvertime.Index).Value = 1
                                    Else
                                        RowDgv.Cells(vFrm.ColIsEarlyOvertime.Index).Value = 0
                                    End If
                                Else
                                    RowDgv.Cells(vFrm.ColIsEarlyOvertime.Index).Value = 0
                                End If

                                RowDgv.Cells(vFrm.ColEarlyOTDuration.Index).Value = Me.ObjCd.Convert_ToString(EarlyOT_Dgv)
                            End If
                            'tutup update kolom early ot ################################################################################### 

                            'update kolom ot duration ###################################################################################
                            OT_Dgv = Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(vFrm.ColOTDuration.Index).Value)
                            Dim KenaLintasHari As Boolean = False
                            If OT_Dgv > 0 Then
                                'jam pulang lembur = schedule out + ot duration
                                EndTime_OT_Dgv = Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(vFrm.ColActualOut.Index).Value)
                                If workday = EnumWorkday.Workday Then
                                    If Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(vFrm.ColActualOut.Index).Value) < Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(vFrm.ColSchOut.Index).Value) Then
                                        EndTime_OT_Dgv += 1440
                                        KenaLintasHari = True
                                    End If
                                End If
                                If EndTime_OT_Dgv <= rowOT_Order.Start_Time Then
                                    OT_Dgv = 0
                                ElseIf EndTime_OT_Dgv > rowOT_Order.Start_Time Then
                                    If rowOT_Order.Break_Out_Time = 0 Or rowOT_Order.Break_In_Time = 0 Then
                                        If KenaLintasHari = True And rowOT_Order.End_Time < Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(vFrm.ColSchOut.Index).Value) Then
                                            EndTime_OT_Dgv -= 1440
                                        End If
                                        If EndTime_OT_Dgv > rowOT_Order.End_Time Then
                                            OT_Dgv = rowOT_Order.OT_Aprroved
                                        End If
                                    Else
                                        Dim BuffStart As Integer = rowOT_Order.Start_Time
                                        Dim BuffBreakOut As Integer = rowOT_Order.Break_Out_Time
                                        Dim BuffBreakIn As Integer = rowOT_Order.Break_In_Time
                                        Dim BuffEnd As Integer = rowOT_Order.End_Time
                                        If rowOT_Order.Start_Time < Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(vFrm.ColSchOut.Index).Value) Then
                                            BuffStart += 1440
                                        End If
                                        If rowOT_Order.Break_Out_Time < Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(vFrm.ColSchOut.Index).Value) Then
                                            BuffBreakOut += 1440
                                        End If
                                        If rowOT_Order.Break_In_Time < Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(vFrm.ColSchOut.Index).Value) Then
                                            BuffBreakIn += 1440
                                        End If
                                        If rowOT_Order.End_Time < Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(vFrm.ColSchOut.Index).Value) Then
                                            BuffEnd += 1440
                                        End If

                                        If EndTime_OT_Dgv <= BuffBreakOut Then
                                            OT_Dgv = EndTime_OT_Dgv - BuffStart
                                        ElseIf EndTime_OT_Dgv > BuffBreakOut And EndTime_OT_Dgv <= BuffBreakIn Then
                                            OT_Dgv = BuffBreakOut - BuffStart
                                        ElseIf EndTime_OT_Dgv > BuffBreakIn Then
                                            If EndTime_OT_Dgv < BuffEnd Then
                                                OT_Dgv = (EndTime_OT_Dgv - BuffBreakIn) + (BuffBreakOut - BuffStart)
                                            ElseIf EndTime_OT_Dgv >= BuffEnd Then
                                                OT_Dgv = rowOT_Order.OT_Aprroved
                                            End If
                                        End If
                                    End If
                                End If
BackOT:
                                'kurangi dengan tolerance
                                If RowDgv.Cells(vFrm.ColCalculateOTTolerance.Index).Value = 0 Then
                                    If OT_Dgv - Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(vFrm.ColCalculateAfterOT.Index).Value) > 0 Then
                                        OT_Dgv = OT_Dgv - Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(vFrm.ColCalculateAfterOT.Index).Value)
                                    End If
                                End If

                                'rounding ot
                                If RTypeOT = Core_Module.ClsMasterWorkShift.enumRoundingType.No_Rounding Then

                                ElseIf RTypeOT = Core_Module.ClsMasterWorkShift.enumRoundingType.Round_down Then
                                    OT_Dgv = Me.ObjCd.RoundDown2(CDec(OT_Dgv), RValueOT)
                                ElseIf RTypeOT = Core_Module.ClsMasterWorkShift.enumRoundingType.Round_Up Then
                                    OT_Dgv = Me.ObjCd.RoundUp2(CDec(OT_Dgv), RValueOT)
                                ElseIf RTypeOT = Core_Module.ClsMasterWorkShift.enumRoundingType.Middle Then
                                    OT_Dgv = Me.ObjCd.MiddleRound(RValueOT, OT_Dgv)
                                End If

                                If MValueOT > 0 Then
                                    If MValueOT >= OT_Dgv Then
                                        RowDgv.Cells(vFrm.ColOTDuration.Index).Value = Me.ObjCd.Convert_ToString(OT_Dgv)
                                    ElseIf MValueOT < OT_Dgv Then
                                        RowDgv.Cells(vFrm.ColOTDuration.Index).Value = Me.ObjCd.Convert_ToString(MValueOT)
                                    End If
                                Else
                                    RowDgv.Cells(vFrm.ColOTDuration.Index).Value = Me.ObjCd.Convert_ToString(OT_Dgv)
                                End If

                                'cek tipe ot
                                CboDgv = RowDgv.Cells(vFrm.ColOT.Index)
                                If RowDgv.Cells(vFrm.ColOT.Index).Value <> CboDgv.Items.Item(ClsTrans_Attendance.EnumTipeOt.No_Overtime) Then
                                    If OT_Dgv > 0 Then
                                        RowDgv.Cells(vFrm.ColIsOvertime.Index).Value = 1
                                    Else
                                        RowDgv.Cells(vFrm.ColIsOvertime.Index).Value = 0
                                    End If
                                Else
                                    RowDgv.Cells(vFrm.ColIsOvertime.Index).Value = 0
                                End If

                                If OT_Dgv > rowOT_Order.OT_Aprroved Then
                                    OT_Dgv = rowOT_Order.OT_Aprroved
                                    GoTo BackOT
                                End If

                                RowDgv.Cells(vFrm.ColOTDuration.Index).Value = Me.ObjCd.Convert_ToString(OT_Dgv)
                            End If
                            'tutup update kolom ot duration ###################################################################################
                        Else
                            RowDgv.Cells(vFrm.ColIsEarlyOvertime.Index).Value = 0
                            RowDgv.Cells(vFrm.ColEarlyOTDuration.Index).Value = Me.ObjCd.Convert_ToString(EarlyOT_Dgv)
                            RowDgv.Cells(vFrm.ColIsOvertime.Index).Value = 0
                            RowDgv.Cells(vFrm.ColOTDuration.Index).Value = Me.ObjCd.Convert_ToString(OT_Dgv)
                        End If
                    ElseIf rowOT_Order.Status = ClsTrans_SPL.EnumStatus.WaitingForApproval Then
                        If vShowMsg = True Then
                            MsgBox("The following overtime order must be approved before save :" + vbNewLine + _
                            rowOT_Order.No_Transaksi, MsgBoxStyle.Information, "Information")
                        End If

                        Return False
                    Else
                        'jika workshift use ot order namun tidak di buat ot ordernya maka di set 0 ot duratioannya.
                        RowDgv.Cells(vFrm.ColPrimaryOvertimeOrder.Index).Value = -1
                        RowDgv.Cells(vFrm.ColStatusRemark.Index).Value = ""
                        RowDgv.Cells(vFrm.ColOvertimeOrderNo.Index).Value = ""
                        RowDgv.Cells(vFrm.ColEarlyOTDuration.Index).Value = "00.00"
                        RowDgv.Cells(vFrm.ColIsEarlyOvertime.Index).Value = 0
                        RowDgv.Cells(vFrm.ColOTDuration.Index).Value = "00.00"
                        RowDgv.Cells(vFrm.ColIsOvertime.Index).Value = 0
                    End If
                Else
                    'jika workshift use ot order namun tidak di buat ot ordernya maka di set 0 ot duratioannya.
                    RowDgv.Cells(vFrm.ColPrimaryOvertimeOrder.Index).Value = -1
                    RowDgv.Cells(vFrm.ColStatusRemark.Index).Value = ""
                    RowDgv.Cells(vFrm.ColOvertimeOrderNo.Index).Value = ""
                    RowDgv.Cells(vFrm.ColEarlyOTDuration.Index).Value = "00.00"
                    RowDgv.Cells(vFrm.ColIsEarlyOvertime.Index).Value = 0
                    RowDgv.Cells(vFrm.ColOTDuration.Index).Value = "00.00"
                    RowDgv.Cells(vFrm.ColIsOvertime.Index).Value = 0
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    'aul tambah overtime order detail
    '    Public Function UpdateAttBySpl_Detail(ByVal vPrimaryEmp As Double, ByVal RowDgv As Windows.Forms.DataGridViewRow, _
    '    ByRef TblOT_Order As DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_ot_1DataTable, _
    '    ByVal vFrm As FrmTrans_Attendance, ByVal vShowMsg As Boolean) As Boolean
    '        Dim rowOT_Order As DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_ot_1Row
    '        Dim OT_Dgv As Double = 0
    '        Dim EarlyOT_Dgv As Double = 0
    '        Dim EndTime_OT_Dgv As Double = 0
    '        Dim StartTime_EarlyOT_Dgv As Double = 0

    '        Dim CboDgv As Windows.Forms.DataGridViewComboBoxCell
    '        Dim workday As ClsTrans_Attendance.EnumWorkday
    '        Dim durasi As ClsTrans_Attendance.EnumDurasi
    '        Dim status As ClsTrans_Attendance.EnumStatus
    '        Dim CalAfterEOT As Double = 0
    '        Dim CalAfterOT As Double = 0
    '        Dim RTypeEOT As Integer = 0
    '        Dim RTypeOT As Integer = 0
    '        Dim RValueEOT As Decimal = 0
    '        Dim RValueOT As Decimal = 0
    '        Dim MValueEOT As Double = 0
    '        Dim MValueOT As Double = 0

    '        'YANG MASUK KE FUNGSI INI HANYA WORKSHIFT YANG MENGGUNAKAN USE OT ORDER SAJA.

    '        Try
    '            'jika ada primary leave application, maka overtime order dianggap ditak ada
    '            If RowDgv.Cells(vFrm.ColPrimaryLA.Index).Value = -1 Then
    '                If Me.ObjCd.Is_Time_String(RowDgv.Cells(vFrm.ColActualIn.Index).Value.ToString) = False Or _
    '                Me.ObjCd.Is_Time_String(RowDgv.Cells(vFrm.ColActualOut.Index).Value.ToString) = False Then
    '                    Return True
    '                End If

    '                'workday
    '                CboDgv = RowDgv.Cells(vFrm.ColWorkdays.Index)
    '                workday = CboDgv.Items.IndexOf(RowDgv.Cells(vFrm.ColWorkdays.Index).Value)
    '                'fullday / half day
    '                CboDgv = RowDgv.Cells(vFrm.ColDuration.Index)
    '                durasi = CboDgv.Items.IndexOf(RowDgv.Cells(vFrm.ColDuration.Index).Value)
    '                'status
    '                CboDgv = RowDgv.Cells(vFrm.ColAttendanceStatus.Index)
    '                status = CboDgv.Items.IndexOf(RowDgv.Cells(vFrm.ColAttendanceStatus.Index).Value)

    '                CalAfterEOT = Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(vFrm.ColCalculateAfterEarlyOT.Index).Value)
    '                CalAfterOT = Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(vFrm.ColCalculateAfterOT.Index).Value)

    '                CboDgv = RowDgv.Cells(vFrm.ColRoundingTypeEarlyOT.Index)
    '                RTypeEOT = CboDgv.Items.IndexOf(RowDgv.Cells(vFrm.ColRoundingTypeEarlyOT.Index).Value)
    '                RTypeOT = CboDgv.Items.IndexOf(RowDgv.Cells(vFrm.ColRoundingTypeOT.Index).Value)

    '                RValueEOT = Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(vFrm.ColRoundingValueEarlyOT.Index).Value)
    '                RValueOT = Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(vFrm.ColRoundingValueOT.Index).Value)

    '                MValueEOT = Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(vFrm.ColMaxValueEarlyOT.Index).Value)
    '                MValueOT = Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(vFrm.ColMaxValueOT.Index).Value)

    '                TblOT_Order.DefaultView.RowFilter = "[Date] = #" & CDate(RowDgv.Cells(vFrm.ColDate.Index).Value) & "#"
    '                If TblOT_Order.DefaultView.Count > 0 Then
    '                    rowOT_Order = TblOT_Order.DefaultView(0).Row
    '                    If rowOT_Order.Status = ClsTrans_SPL.EnumStatus.Approved Then
    '                        RowDgv.Cells(vFrm.ColPrimaryOvertimeOrder.Index).Value = rowOT_Order.Primary_Transaksi
    '                        RowDgv.Cells(vFrm.ColStatusRemark.Index).Value = rowOT_Order.Reason
    '                        RowDgv.Cells(vFrm.ColOvertimeOrderNo.Index).Value = rowOT_Order.No_Transaksi
    '                        RowDgv.Cells(vFrm.ColEarlyOTOrderDuration.Index).Value = Me.ObjCd.Convert_ToString(rowOT_Order.OT_Early_Approved)
    '                        RowDgv.Cells(vFrm.ColOTOrderDuration.Index).Value = Me.ObjCd.Convert_ToString(rowOT_Order.OT_Aprroved)

    '                        'update kolom early ot ################################################################################### 
    '                        EarlyOT_Dgv = Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(vFrm.ColEarlyOTDuration.Index).Value)
    '                        If EarlyOT_Dgv > 0 Then
    '                            StartTime_EarlyOT_Dgv = Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(vFrm.ColActualIn.Index).Value)
    '                            If StartTime_EarlyOT_Dgv >= rowOT_Order.End_Time_Early Then
    '                                EarlyOT_Dgv = 0
    '                            ElseIf StartTime_EarlyOT_Dgv < rowOT_Order.End_Time_Early Then
    '                                If rowOT_Order.Break_Out_Time_Early = 0 Or rowOT_Order.Break_In_Time_Early = 0 Then
    '                                    If StartTime_EarlyOT_Dgv < rowOT_Order.Start_Time_Early Then
    '                                        EarlyOT_Dgv = rowOT_Order.OT_Early_Approved
    '                                    End If
    '                                Else
    '                                    If StartTime_EarlyOT_Dgv > rowOT_Order.Break_In_Time_Early Then
    '                                        EarlyOT_Dgv = rowOT_Order.End_Time_Early - StartTime_EarlyOT_Dgv
    '                                    ElseIf StartTime_EarlyOT_Dgv > rowOT_Order.Break_Out_Time_Early And StartTime_EarlyOT_Dgv <= rowOT_Order.Break_In_Time_Early Then
    '                                        EarlyOT_Dgv = rowOT_Order.End_Time_Early - rowOT_Order.Break_In_Time_Early
    '                                    ElseIf StartTime_EarlyOT_Dgv < rowOT_Order.Break_Out_Time_Early Then
    '                                        If StartTime_EarlyOT_Dgv > rowOT_Order.Start_Time_Early Then
    '                                            EarlyOT_Dgv = (rowOT_Order.Break_Out_Time_Early - StartTime_EarlyOT_Dgv) + (rowOT_Order.End_Time_Early - rowOT_Order.Break_In_Time_Early)
    '                                        ElseIf StartTime_EarlyOT_Dgv <= rowOT_Order.Start_Time_Early Then
    '                                            EarlyOT_Dgv = rowOT_Order.OT_Early_Approved
    '                                        End If
    '                                    End If
    '                                End If
    '                            End If

    '                            'kurangi dengan tolerance
    '                            If RowDgv.Cells(vFrm.ColCalculateEarlyOTTolerance.Index).Value = 0 Then
    '                                If EarlyOT_Dgv - Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(vFrm.ColCalculateAfterEarlyOT.Index).Value) > 0 Then
    '                                    EarlyOT_Dgv = EarlyOT_Dgv - Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(vFrm.ColCalculateAfterEarlyOT.Index).Value)
    '                                End If
    '                            End If

    '                            'rounding early in
    '                            If RTypeEOT = Core_Module.ClsMasterWorkShift.enumRoundingType.No_Rounding Then
    '                            ElseIf RTypeEOT = Core_Module.ClsMasterWorkShift.enumRoundingType.Round_down Then
    '                                EarlyOT_Dgv = Me.ObjCd.RoundDown2(CDec(EarlyOT_Dgv), RValueEOT)
    '                            ElseIf RTypeEOT = Core_Module.ClsMasterWorkShift.enumRoundingType.Round_Up Then
    '                                EarlyOT_Dgv = Me.ObjCd.RoundUp2(CDec(EarlyOT_Dgv), RValueEOT)
    '                            ElseIf RTypeEOT = Core_Module.ClsMasterWorkShift.enumRoundingType.Middle Then
    '                                EarlyOT_Dgv = Me.ObjCd.MiddleRound(RValueEOT, CDec(EarlyOT_Dgv))
    '                            End If

    '                            If MValueEOT > 0 Then
    '                                If MValueEOT >= EarlyOT_Dgv Then
    '                                    RowDgv.Cells(vFrm.ColEarlyOTDuration.Index).Value = Me.ObjCd.Convert_ToString(EarlyOT_Dgv)
    '                                ElseIf MValueEOT < EarlyOT_Dgv Then
    '                                    RowDgv.Cells(vFrm.ColEarlyOTDuration.Index).Value = Me.ObjCd.Convert_ToString(MValueEOT)
    '                                End If
    '                            Else
    '                                RowDgv.Cells(vFrm.ColEarlyOTDuration.Index).Value = Me.ObjCd.Convert_ToString(EarlyOT_Dgv)
    '                            End If

    '                            'cek type ot
    '                            CboDgv = RowDgv.Cells(vFrm.ColOT.Index)
    '                            If RowDgv.Cells(vFrm.ColOT.Index).Value <> CboDgv.Items.Item(ClsTrans_Attendance.EnumTipeOt.No_Overtime) Then
    '                                If EarlyOT_Dgv > 0 Then
    '                                    RowDgv.Cells(vFrm.ColIsEarlyOvertime.Index).Value = 1
    '                                Else
    '                                    RowDgv.Cells(vFrm.ColIsEarlyOvertime.Index).Value = 0
    '                                End If
    '                            Else
    '                                RowDgv.Cells(vFrm.ColIsEarlyOvertime.Index).Value = 0
    '                            End If

    '                            RowDgv.Cells(vFrm.ColEarlyOTDuration.Index).Value = Me.ObjCd.Convert_ToString(EarlyOT_Dgv)
    '                        End If
    '                        'tutup update kolom early ot ################################################################################### 

    '                        'update kolom ot duration ###################################################################################
    '                        OT_Dgv = Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(vFrm.ColOTDuration.Index).Value)
    '                        If OT_Dgv > 0 Then
    '                            'jam pulang lembur = schedule out + ot duration
    '                            EndTime_OT_Dgv = Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(vFrm.ColActualOut.Index).Value)
    '                            If workday = EnumWorkday.Workday Then
    '                                If Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(vFrm.ColActualOut.Index).Value) < Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(vFrm.ColSchOut.Index).Value) Then
    '                                    EndTime_OT_Dgv += 1440
    '                                End If
    '                            End If
    '                            If EndTime_OT_Dgv <= rowOT_Order.Start_Time Then
    '                                OT_Dgv = 0
    '                            ElseIf EndTime_OT_Dgv > rowOT_Order.Start_Time Then
    '                                If rowOT_Order.Break_Out_Time = 0 Or rowOT_Order.Break_In_Time = 0 Then
    '                                    If EndTime_OT_Dgv > rowOT_Order.End_Time Then
    '                                        OT_Dgv = rowOT_Order.OT_Aprroved
    '                                    End If
    '                                Else
    '                                    'If workday = EnumWorkday.Holiday Then
    '                                    '    If EndTime_OT_Dgv <= rowOT_Order.Break_Out_Time Then
    '                                    '        OT_Dgv = EndTime_OT_Dgv - Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(vFrm.ColActualIn.Index).Value)
    '                                    '    ElseIf EndTime_OT_Dgv > rowOT_Order.Break_Out_Time And EndTime_OT_Dgv <= rowOT_Order.Break_In_Time Then
    '                                    '        OT_Dgv = rowOT_Order.Break_Out_Time - Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(vFrm.ColActualIn.Index).Value)
    '                                    '    ElseIf EndTime_OT_Dgv > rowOT_Order.Break_In_Time Then
    '                                    '        If EndTime_OT_Dgv < rowOT_Order.End_Time Then
    '                                    '            OT_Dgv = (EndTime_OT_Dgv - rowOT_Order.Break_In_Time) + (rowOT_Order.Break_Out_Time - rowOT_Order.Start_Time)
    '                                    '        ElseIf EndTime_OT_Dgv >= rowOT_Order.End_Time Then
    '                                    '            OT_Dgv = rowOT_Order.OT_Aprroved
    '                                    '        End If
    '                                    '    End If
    '                                    'Else
    '                                    If EndTime_OT_Dgv <= rowOT_Order.Break_Out_Time Then
    '                                        OT_Dgv = EndTime_OT_Dgv - rowOT_Order.Start_Time
    '                                    ElseIf EndTime_OT_Dgv > rowOT_Order.Break_Out_Time And EndTime_OT_Dgv <= rowOT_Order.Break_In_Time Then
    '                                        OT_Dgv = rowOT_Order.Break_Out_Time - rowOT_Order.Start_Time
    '                                    ElseIf EndTime_OT_Dgv > rowOT_Order.Break_In_Time Then
    '                                        If EndTime_OT_Dgv < rowOT_Order.End_Time Then
    '                                            OT_Dgv = (EndTime_OT_Dgv - rowOT_Order.Break_In_Time) + (rowOT_Order.Break_Out_Time - rowOT_Order.Start_Time)
    '                                        ElseIf EndTime_OT_Dgv >= rowOT_Order.End_Time Then
    '                                            OT_Dgv = rowOT_Order.OT_Aprroved
    '                                        End If
    '                                    End If
    '                                    'End If
    '                                End If
    '                            End If

    'BackOT:

    '                            'kurangi dengan tolerance
    '                            If RowDgv.Cells(vFrm.ColCalculateOTTolerance.Index).Value = 0 Then
    '                                If OT_Dgv - Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(vFrm.ColCalculateAfterOT.Index).Value) > 0 Then
    '                                    OT_Dgv = OT_Dgv - Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(vFrm.ColCalculateAfterOT.Index).Value)
    '                                End If
    '                            End If

    '                            'rounding ot
    '                            If RTypeOT = Core_Module.ClsMasterWorkShift.enumRoundingType.No_Rounding Then

    '                            ElseIf RTypeOT = Core_Module.ClsMasterWorkShift.enumRoundingType.Round_down Then
    '                                OT_Dgv = Me.ObjCd.RoundDown2(CDec(OT_Dgv), RValueOT)
    '                            ElseIf RTypeOT = Core_Module.ClsMasterWorkShift.enumRoundingType.Round_Up Then
    '                                OT_Dgv = Me.ObjCd.RoundUp2(CDec(OT_Dgv), RValueOT)
    '                            ElseIf RTypeOT = Core_Module.ClsMasterWorkShift.enumRoundingType.Middle Then
    '                                OT_Dgv = Me.ObjCd.MiddleRound(RValueOT, OT_Dgv)
    '                            End If

    '                            If MValueOT > 0 Then
    '                                If MValueOT >= OT_Dgv Then
    '                                    RowDgv.Cells(vFrm.ColOTDuration.Index).Value = Me.ObjCd.Convert_ToString(OT_Dgv)
    '                                ElseIf MValueOT < OT_Dgv Then
    '                                    RowDgv.Cells(vFrm.ColOTDuration.Index).Value = Me.ObjCd.Convert_ToString(MValueOT)
    '                                End If
    '                            Else
    '                                RowDgv.Cells(vFrm.ColOTDuration.Index).Value = Me.ObjCd.Convert_ToString(OT_Dgv)
    '                            End If

    '                            'cek tipe ot
    '                            CboDgv = RowDgv.Cells(vFrm.ColOT.Index)
    '                            If RowDgv.Cells(vFrm.ColOT.Index).Value <> CboDgv.Items.Item(ClsTrans_Attendance.EnumTipeOt.No_Overtime) Then
    '                                If OT_Dgv > 0 Then
    '                                    RowDgv.Cells(vFrm.ColIsOvertime.Index).Value = 1
    '                                Else
    '                                    RowDgv.Cells(vFrm.ColIsOvertime.Index).Value = 0
    '                                End If
    '                            Else
    '                                RowDgv.Cells(vFrm.ColIsOvertime.Index).Value = 0
    '                            End If

    '                            If OT_Dgv > rowOT_Order.OT_Aprroved Then
    '                                OT_Dgv = rowOT_Order.OT_Aprroved
    '                                GoTo BackOT
    '                            End If

    '                            RowDgv.Cells(vFrm.ColOTDuration.Index).Value = Me.ObjCd.Convert_ToString(OT_Dgv)
    '                        End If
    '                            'tutup update kolom ot duration ###################################################################################
    '                        ElseIf rowOT_Order.Status = ClsTrans_SPL.EnumStatus.WaitingForApproval Then
    '                            If vShowMsg = True Then
    '                                MsgBox("The following overtime order must be approved before save :" + vbNewLine + _
    '                                rowOT_Order.No_Transaksi, MsgBoxStyle.Information, "Information")
    '                            End If

    '                            Return False
    '                        Else
    '                            'jika workshift use ot order namun tidak di buat ot ordernya maka di set 0 ot duratioannya.
    '                            RowDgv.Cells(vFrm.ColPrimaryOvertimeOrder.Index).Value = -1
    '                            RowDgv.Cells(vFrm.ColStatusRemark.Index).Value = ""
    '                            RowDgv.Cells(vFrm.ColOvertimeOrderNo.Index).Value = ""
    '                            RowDgv.Cells(vFrm.ColEarlyOTDuration.Index).Value = "00.00"
    '                            RowDgv.Cells(vFrm.ColIsEarlyOvertime.Index).Value = 0
    '                            RowDgv.Cells(vFrm.ColOTDuration.Index).Value = "00.00"
    '                            RowDgv.Cells(vFrm.ColIsOvertime.Index).Value = 0
    '                        End If
    '                    Else
    '                        'jika workshift use ot order namun tidak di buat ot ordernya maka di set 0 ot duratioannya.
    '                        RowDgv.Cells(vFrm.ColPrimaryOvertimeOrder.Index).Value = -1
    '                        RowDgv.Cells(vFrm.ColStatusRemark.Index).Value = ""
    '                        RowDgv.Cells(vFrm.ColOvertimeOrderNo.Index).Value = ""
    '                        RowDgv.Cells(vFrm.ColEarlyOTDuration.Index).Value = "00.00"
    '                        RowDgv.Cells(vFrm.ColIsEarlyOvertime.Index).Value = 0
    '                        RowDgv.Cells(vFrm.ColOTDuration.Index).Value = "00.00"
    '                        RowDgv.Cells(vFrm.ColIsOvertime.Index).Value = 0
    '                    End If
    '                End If

    '                Return True
    '        Catch ex As Exception
    '            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '        End Try
    '    End Function

    Private Function FilterLeaveApproval(ByRef tblExp As DataTable, ByVal tblLeave As DataTable) As Boolean
        Dim RowExp As DataRow
        Try
            If tblLeave.Rows.Count > 0 Then
                For ij As Integer = tblExp.Rows.Count - 1 To 0 Step -1
                    RowExp = tblExp.Rows(ij)
                    tblLeave.DefaultView.RowFilter = "[primary]=" & RowExp.Item("primary") & "AND [status] =" & ClsTrans_Cuti_Pengajuan.EnumStatus.App
                    If tblLeave.DefaultView.Count = 0 Then
                        tblExp.Rows.Remove(RowExp)
                    End If
                Next
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            tblLeave.Rows.Clear()
        End Try
    End Function

    'change by ahdar
    Private Sub IsiDgvItem(ByVal vFrm As FrmTrans_Attendance, ByVal vAwal As Date, ByVal vAkhir As Date, _
    ByVal vRow As DseTrans_Attendance.DataTable2Row)
        Dim DtaHol As New DseTrans_AttendanceTableAdapters.tblmaster_workshift_itemholidayTableAdapter
        Dim TblHol As New DseTrans_Attendance.tblmaster_workshift_itemholidayDataTable
        Dim RowHol As DseTrans_Attendance.tblmaster_workshift_itemholidayRow
        Dim DtaWork As New DseTrans_AttendanceTableAdapters.tblmaster_workshift_itemTableAdapter
        Dim TblWork As New DseTrans_Attendance.tblmaster_workshift_itemDataTable
        Dim RowWork As DseTrans_Attendance.tblmaster_workshift_itemRow
        Dim DtaLeave As New DseTrans_AttendanceTableAdapters.tbltrans_cuti_pengajuanTableAdapter
        Dim TblLeave As New DseTrans_Attendance.tbltrans_cuti_pengajuanDataTable
        Dim CountDate As Integer = 0
        Dim RowDgv As Windows.Forms.DataGridViewRow
        Dim TglDgv As Date = #1/1/1950#
        Dim objEmp As New Core_Module.ClsMasterKaryawan

        Dim CboDgv As Windows.Forms.DataGridViewComboBoxCell
        Dim workday As ClsTrans_Attendance.EnumWorkday
        Dim dtaHoliday As New DseTrans_AttendanceTableAdapters.tblmaster_workshift_itemholiday1TableAdapter
        Dim tblHoliday As New DseTrans_Attendance.tblmaster_workshift_itemholiday1DataTable

        Dim DtaExp As New DseTrans_Cuti_PengajuanTableAdapters.tblexplainTableAdapter
        Dim TblExp As New DseTrans_Cuti_Pengajuan.tblexplainDataTable
        Dim RowExp As DseTrans_Cuti_Pengajuan.tblexplainRow
        Dim StatusExplain As Integer = -1
        Dim DurationExplain As Decimal = 0
        Try
            vFrm.dgvItem.AutoSizeColumnsMode = Windows.Forms.DataGridViewAutoSizeColumnsMode.None
            vFrm.dgvItem.AutoSizeRowsMode = Windows.Forms.DataGridViewAutoSizeRowsMode.None

            dtaHoliday.Connection.ConnectionString = Core_Module.GetConstring
            dtaHoliday.Fill(tblHoliday, vFrm.dtpStart.Value.Date, vFrm.DtpEnd.Value.Date)
            DtaExp.Connection.ConnectionString = dtaHoliday.Connection.ConnectionString
            DtaExp.Fill(TblExp)
            DtaLeave.Fill(TblLeave)
            Me.FilterLeaveApproval(TblExp, TblLeave)
            TglDgv = vAwal

            For CountDate = 0 To (vAkhir.Date - vAwal.Date).TotalDays
                vFrm.dgvItem.Rows.Add()
                RowDgv = vFrm.dgvItem.Rows(vFrm.dgvItem.Rows.Count - 1)
                'no
                RowDgv.Cells(vFrm.ColNo.Index).Value = vFrm.dgvItem.Rows.Count
                'date
                RowDgv.Cells(vFrm.ColDate.Index).Value = Format(TglDgv, "dd MMM yyyy")
                'primary ws
                RowDgv.Cells(vFrm.ColWorkshift.Index).Value = Me.ObjWs.GetCodeByPrimary(vRow.Primary_Workshift)
                'holiday / workday
                DtaHol.FillBy_PrimaryWSnDate(TblHol, vRow.Primary_Workshift, TglDgv)
                TblExp.DefaultView.RowFilter = "[primary karyawan]=" & vFrm.TxtPrimaryEmployee.Text & "And [code workshift]='" & RowDgv.Cells(vFrm.ColWorkshift.Index).Value & "'" & _
                " And [date]='" & RowDgv.Cells(vFrm.ColDate.Index).Value & "'"
                StatusExplain = -1
                DurationExplain = 0
                If TblExp.DefaultView.Count > 0 Then
                    RowExp = TblExp.DefaultView(0).Row
                    StatusExplain = RowExp.status
                    DurationExplain = RowExp.duration
                End If

                If TblHol.Rows.Count > 0 Then
                    'kondisi holiday
                    For Each RowHol In TblHol.Rows
                        Me.InitHoliday(vFrm, RowDgv, RowHol, TglDgv, DtaLeave, TblLeave)
                        Me.KosongkanNonvisibleDGv(vFrm, RowDgv)
                    Next
                Else
                    'kondisi workday
                    DtaWork.FillBy_PrimaryWSnDay(TblWork, vRow.Primary_Workshift, TglDgv.DayOfWeek)
                    For Each RowWork In TblWork.Rows
                        Me.InitWorkDay(vFrm, RowDgv, RowWork, TglDgv, DtaLeave, TblLeave, StatusExplain, DurationExplain)
                        Me.KosongkanNonvisibleDGv(vFrm, RowDgv)
                    Next
                End If

                CboDgv = vFrm.dgvItem.Item(vFrm.ColWorkdays.Index, RowDgv.Index)
                workday = CboDgv.Items.IndexOf(vFrm.dgvItem.Item(vFrm.ColWorkdays.Index, RowDgv.Index).Value)
                vFrm.ApplyAttendanceSettings(RowDgv, vRow.Primary_Workshift, vFrm.GetCalculationType(workday, vFrm.dgvItem.Item(vFrm.ColDate.Index, RowDgv.Index).Value, vFrm.dgvItem.Item(vFrm.ColPrimaryWorkShift.Index, RowDgv.Index).Value, tblHoliday))

                TglDgv = TglDgv.AddDays(1)
            Next

            vFrm.dgvItem.AutoSizeColumnsMode = Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DtaHol.Dispose()
            DtaWork.Dispose()
            TblHol.Dispose()
            TblWork.Dispose()
            DtaLeave.Dispose()
            TblLeave.Dispose()
        End Try
        'Dim DtaHol As New DseTrans_AttendanceTableAdapters.tblmaster_workshift_itemholidayTableAdapter
        'Dim TblHol As New DseTrans_Attendance.tblmaster_workshift_itemholidayDataTable
        'Dim RowHol As DseTrans_Attendance.tblmaster_workshift_itemholidayRow
        'Dim DtaWork As New DseTrans_AttendanceTableAdapters.tblmaster_workshift_itemTableAdapter
        'Dim TblWork As New DseTrans_Attendance.tblmaster_workshift_itemDataTable
        'Dim RowWork As DseTrans_Attendance.tblmaster_workshift_itemRow
        'Dim DtaLeave As New DseTrans_AttendanceTableAdapters.tbltrans_cuti_pengajuanTableAdapter
        'Dim TblLeave As New DseTrans_Attendance.tbltrans_cuti_pengajuanDataTable
        'Dim CountDate As Integer = 0
        'Dim RowDgv As Windows.Forms.DataGridViewRow
        'Dim TglDgv As Date = #1/1/1950#
        'Dim objEmp As New Core_Module.ClsMasterKaryawan

        'Dim CboDgv As Windows.Forms.DataGridViewComboBoxCell
        'Dim workday As ClsTrans_Attendance.EnumWorkday
        'Dim dtaHoliday As New DseTrans_AttendanceTableAdapters.tblmaster_workshift_itemholiday1TableAdapter
        'Dim tblHoliday As New DseTrans_Attendance.tblmaster_workshift_itemholiday1DataTable

        'Try
        '    vFrm.dgvItem.AutoSizeColumnsMode = Windows.Forms.DataGridViewAutoSizeColumnsMode.None
        '    vFrm.dgvItem.AutoSizeRowsMode = Windows.Forms.DataGridViewAutoSizeRowsMode.None

        '    dtaHoliday.Connection.ConnectionString = Core_Module.GetConstring
        '    dtaHoliday.Fill(tblHoliday, vFrm.dtpStart.Value.Date, vFrm.DtpEnd.Value.Date)

        '    TglDgv = vAwal

        '    For CountDate = 0 To (vAkhir.Date - vAwal.Date).TotalDays
        '        vFrm.dgvItem.Rows.Add()
        '        RowDgv = vFrm.dgvItem.Rows(vFrm.dgvItem.Rows.Count - 1)
        '        'no
        '        RowDgv.Cells(vFrm.ColNo.Index).Value = vFrm.dgvItem.Rows.Count
        '        'date
        '        RowDgv.Cells(vFrm.ColDate.Index).Value = Format(TglDgv, "dd MMM yyyy")
        '        'primary ws
        '        RowDgv.Cells(vFrm.ColWorkshift.Index).Value = Me.ObjWs.GetCodeByPrimary(vRow.Primary_Workshift)
        '        'holiday / workday
        '        DtaHol.FillBy_PrimaryWSnDate(TblHol, vRow.Primary_Workshift, TglDgv)
        '        If TblHol.Rows.Count > 0 Then
        '            'kondisi holiday
        '            For Each RowHol In TblHol.Rows
        '                Me.InitHoliday(vFrm, RowDgv, RowHol, TglDgv, DtaLeave, TblLeave)
        '                Me.KosongkanNonvisibleDGv(vFrm, RowDgv)
        '            Next
        '        Else
        '            'kondisi workday
        '            DtaWork.FillBy_PrimaryWSnDay(TblWork, vRow.Primary_Workshift, TglDgv.DayOfWeek)
        '            For Each RowWork In TblWork.Rows
        '                Me.InitWorkDay(vFrm, RowDgv, RowWork, TglDgv, DtaLeave, TblLeave)
        '                Me.KosongkanNonvisibleDGv(vFrm, RowDgv)
        '            Next
        '        End If

        '        CboDgv = vFrm.dgvItem.Item(vFrm.ColWorkdays.Index, RowDgv.Index)
        '        workday = CboDgv.Items.IndexOf(vFrm.dgvItem.Item(vFrm.ColWorkdays.Index, RowDgv.Index).Value)
        '        vFrm.ApplyAttendanceSettings(RowDgv, vRow.Primary_Workshift, vFrm.GetCalculationType(workday, vFrm.dgvItem.Item(vFrm.ColDate.Index, RowDgv.Index).Value, vFrm.dgvItem.Item(vFrm.ColPrimaryWorkShift.Index, RowDgv.Index).Value, tblHoliday))

        '        TglDgv = TglDgv.AddDays(1)
        '    Next

        '    vFrm.dgvItem.AutoSizeColumnsMode = Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'Finally
        '    DtaHol.Dispose()
        '    DtaWork.Dispose()
        '    TblHol.Dispose()
        '    TblWork.Dispose()
        '    DtaLeave.Dispose()
        '    TblLeave.Dispose()
        'End Try
    End Sub

    Private Sub InitHoliday(ByVal vFrm As FrmTrans_Attendance, ByVal RowDgv As Windows.Forms.DataGridViewRow, _
    ByVal RowHol As DseTrans_Attendance.tblmaster_workshift_itemholidayRow, ByVal vDate As Date, _
    ByVal DtaLeave As DseTrans_AttendanceTableAdapters.tbltrans_cuti_pengajuanTableAdapter, _
    ByVal TblLeave As DseTrans_Attendance.tbltrans_cuti_pengajuanDataTable)
        Dim CboDgv As Windows.Forms.DataGridViewComboBoxCell
        Dim Workday As EnumWorkday
        Dim Durasi As EnumDurasi
        Dim TypeOt As EnumTipeOt
        Dim status As EnumStatus
        Dim rowLeave As DseTrans_Attendance.tbltrans_cuti_pengajuanRow
        Try
            'primary ws
            RowDgv.Cells(vFrm.ColPrimaryWorkShift.Index).Value = RowHol.primary_transaksi
            'workshift
            Workday = EnumWorkday.Holiday
            CboDgv = RowDgv.Cells(vFrm.ColWorkdays.Index)
            RowDgv.Cells(vFrm.ColWorkdays.Index).Value = CboDgv.Items.Item(ClsTrans_Attendance.EnumWorkday.Holiday)
            'durasi
            Durasi = RowHol.duration
            CboDgv = RowDgv.Cells(vFrm.ColDuration.Index)
            RowDgv.Cells(vFrm.ColDuration.Index).Value = CboDgv.Items.Item(RowHol.duration)
            'status
            CboDgv = RowDgv.Cells(vFrm.ColAttendanceStatus.Index)
            If Durasi = EnumDurasi.Full_Day Then
                RowDgv.Cells(vFrm.ColAttendanceStatus.Index).Value = CboDgv.Items.Item(ClsTrans_Attendance.EnumStatus.Holiday)
            ElseIf Durasi = EnumDurasi.Half_Day Then
                RowDgv.Cells(vFrm.ColAttendanceStatus.Index).Value = CboDgv.Items.Item(ClsTrans_Attendance.EnumStatus.Half_Holiday)
            End If
            'sch in
            RowDgv.Cells(vFrm.ColSchIn.Index).Value = Me.ObjCd.Convert_ToString(RowHol.sch_in)
            'sch break out
            RowDgv.Cells(vFrm.ColSchBreakOut.Index).Value = Me.ObjCd.Convert_ToString(RowHol.schedule_break_out)
            'sch break in
            RowDgv.Cells(vFrm.ColSchBreakIn.Index).Value = Me.ObjCd.Convert_ToString(RowHol.schedule_break_in)
            'sch out
            RowDgv.Cells(vFrm.ColSchOut.Index).Value = Me.ObjCd.Convert_ToString(RowHol.sch_out)
            'act in
            RowDgv.Cells(vFrm.ColActualIn.Index).Value = Me.ObjCd.Convert_ToString(0)
            'act break out
            RowDgv.Cells(vFrm.ColActualBreakOut.Index).Value = Me.ObjCd.Convert_ToString(0)
            'act break in
            RowDgv.Cells(vFrm.ColActualBreakIn.Index).Value = Me.ObjCd.Convert_ToString(0)
            'act out
            RowDgv.Cells(vFrm.ColActualOut.Index).Value = Me.ObjCd.Convert_ToString(0)
            'status remarks, leave code, leave desc, leave doc
            'jika ada data master cuti di ws holiday maka ambil dari ws holiday
            'jika tidak ada ambil dari leave app
            If RowHol.primary_leave > 0 Then

                Dim DurasiLeave As EnumDurasi
                'cek di leave app
                DtaLeave.FillBy_PrimaryEmpNDate(TblLeave, vDate, vFrm.TxtPrimaryEmployee.Text)
                If TblLeave.Rows.Count > 0 Then
                    For Each rowLeave In TblLeave.Rows
                        RowDgv.Cells(vFrm.ColStatusRemark.Index).Value = rowLeave.reason
                        RowDgv.Cells(vFrm.ColLeaveCode.Index).Value = rowLeave.CutiCode
                        RowDgv.Cells(vFrm.ColLeaveDesc.Index).Value = rowLeave.CutiDesc
                        RowDgv.Cells(vFrm.ColLeaveDoc.Index).Value = rowLeave.no_transaksi
                        RowDgv.Cells(vFrm.ColPrimaryLeave.Index).Value = rowLeave.primary_cuti
                        RowDgv.Cells(vFrm.ColPrimaryLA.Index).Value = rowLeave.primary
                        'ubah status leave
                        'durasi ngikutin durasi di leave app
                        If vDate = rowLeave.start_date Then
                            If rowLeave.half_day_start = ClsTrans_Cuti_Pengajuan.EnumHalfDay.FullDay Then
                                DurasiLeave = EnumDurasi.Full_Day
                            ElseIf rowLeave.half_day_start = ClsTrans_Cuti_Pengajuan.EnumHalfDay.HalfDay Then
                                DurasiLeave = EnumDurasi.Half_Day
                            End If
                        ElseIf vDate = rowLeave.end_date Then
                            If rowLeave.half_day_start = ClsTrans_Cuti_Pengajuan.EnumHalfDay.FullDay Then
                                DurasiLeave = EnumDurasi.Full_Day
                            ElseIf rowLeave.half_day_start = ClsTrans_Cuti_Pengajuan.EnumHalfDay.HalfDay Then
                                DurasiLeave = EnumDurasi.Half_Day
                            End If
                        Else
                            DurasiLeave = Durasi
                        End If
                        CboDgv = RowDgv.Cells(vFrm.ColAttendanceStatus.Index)
                        If DurasiLeave = EnumDurasi.Full_Day Then
                            RowDgv.Cells(vFrm.ColAttendanceStatus.Index).Value = CboDgv.Items.Item(ClsTrans_Attendance.EnumStatus.Leave)
                        ElseIf DurasiLeave = EnumDurasi.Half_Day Then
                            RowDgv.Cells(vFrm.ColAttendanceStatus.Index).Value = CboDgv.Items.Item(ClsTrans_Attendance.EnumStatus.Half_day_Leave)
                        End If
                    Next
                End If
            Else
                'cek holiday list
                RowDgv.Cells(vFrm.ColStatusRemark.Index).Value = ""
                RowDgv.Cells(vFrm.ColLeaveCode.Index).Value = RowHol.CutiCode
                RowDgv.Cells(vFrm.ColLeaveDesc.Index).Value = RowHol.CutiDesc
                RowDgv.Cells(vFrm.ColLeaveDoc.Index).Value = ""
                RowDgv.Cells(vFrm.ColPrimaryLeave.Index).Value = RowHol.primary_leave
                'ubah status leave
                Durasi = RowHol.duration
                CboDgv = RowDgv.Cells(vFrm.ColAttendanceStatus.Index)
                If Durasi = EnumDurasi.Full_Day Then
                    RowDgv.Cells(vFrm.ColAttendanceStatus.Index).Value = CboDgv.Items.Item(ClsTrans_Attendance.EnumStatus.Holiday)
                ElseIf Durasi = EnumDurasi.Half_Day Then
                    RowDgv.Cells(vFrm.ColAttendanceStatus.Index).Value = CboDgv.Items.Item(ClsTrans_Attendance.EnumStatus.Half_Holiday)
                    'act in
                    RowDgv.Cells(vFrm.ColActualIn.Index).Value = Me.ObjCd.Convert_ToString(RowHol.sch_in)
                    'act break out
                    RowDgv.Cells(vFrm.ColActualBreakOut.Index).Value = Me.ObjCd.Convert_ToString(RowHol.schedule_break_out)
                    'act break in
                    RowDgv.Cells(vFrm.ColActualBreakIn.Index).Value = Me.ObjCd.Convert_ToString(RowHol.schedule_break_in)
                    'act out
                    RowDgv.Cells(vFrm.ColActualOut.Index).Value = Me.ObjCd.Convert_ToString(RowHol.sch_out)
                End If
            End If
            'is on holiday list =1
            RowDgv.Cells(vFrm.ColIsOnHolidayList.Index).Value = 1
            'type ot
            TypeOt = EnumTipeOt.Holiday_Overtime
            CboDgv = RowDgv.Cells(vFrm.ColOT.Index)
            RowDgv.Cells(vFrm.ColOT.Index).Value = CboDgv.Items.Item(TypeOt)
            'cek status
            Me.InitIsStatus(vFrm, RowDgv, Durasi)
            'masukin yang berkaitan dengan hitungan
            Me.InitHitungan(vFrm, RowDgv)
            'kosongkan yang fullday tidak masuk
            CboDgv = RowDgv.Cells(vFrm.ColAttendanceStatus.Index)
            status = CboDgv.Items.IndexOf(RowDgv.Cells(vFrm.ColAttendanceStatus.Index).Value)
            Me.KosongkanFullTidakMasuk(vFrm, RowDgv, status)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'change by ahdar
    Private Sub InitWorkDay(ByVal vFrm As FrmTrans_Attendance, ByVal RowDgv As Windows.Forms.DataGridViewRow, _
    ByVal RowWork As DseTrans_Attendance.tblmaster_workshift_itemRow, ByVal vDate As Date, _
    ByVal DtaLeave As DseTrans_AttendanceTableAdapters.tbltrans_cuti_pengajuanTableAdapter, _
    ByVal TblLeave As DseTrans_Attendance.tbltrans_cuti_pengajuanDataTable, _
    ByVal vStatusExplain As ClsTrans_Cuti_Pengajuan.EnumStatusExplain, ByVal vDuration As Decimal)
        Dim CboDgv As Windows.Forms.DataGridViewComboBoxCell
        Dim Workday As EnumWorkday
        Dim Durasi As EnumDurasi
        Dim TypeOt As EnumTipeOt
        Dim Status As EnumStatus
        Dim rowLeave As DseTrans_Attendance.tbltrans_cuti_pengajuanRow
        Try
            RowDgv.Cells(vFrm.ColShowLog.Index).Value = "Show"
            'primary ws
            RowDgv.Cells(vFrm.ColPrimaryWorkShift.Index).Value = RowWork.primary_transaksi
            'workshift
            CboDgv = RowDgv.Cells(vFrm.ColWorkdays.Index)
            If RowWork.tipe_workdays = EnumWorkday.Workday Then
                Workday = EnumWorkday.Workday
                RowDgv.Cells(vFrm.ColWorkdays.Index).Value = CboDgv.Items.Item(ClsTrans_Attendance.EnumWorkday.Workday)
                TypeOt = EnumTipeOt.Weekday_Overtime
            ElseIf RowWork.tipe_workdays = EnumWorkday.Holiday Then
                Workday = EnumWorkday.Holiday
                RowDgv.Cells(vFrm.ColWorkdays.Index).Value = CboDgv.Items.Item(ClsTrans_Attendance.EnumWorkday.Holiday)
                TypeOt = EnumTipeOt.Holiday_Overtime
            End If
            'durasi
            Durasi = RowWork.duratioin
            CboDgv = RowDgv.Cells(vFrm.ColDuration.Index)
            RowDgv.Cells(vFrm.ColDuration.Index).Value = CboDgv.Items.Item(RowWork.duratioin)
            'status
            CboDgv = RowDgv.Cells(vFrm.ColAttendanceStatus.Index)
            'sebelumnya cek dulu status di pengajuan cuti explain 
            If vStatusExplain <> -1 And Workday <> EnumWorkday.Holiday Then
                Dim DurasiLeave As EnumDurasi
                RowDgv.Cells(vFrm.ColAttendanceStatus.Index).Value = CboDgv.Items.Item(Me.GetStatusByExplain(vStatusExplain, vDuration, RowDgv.Cells(vFrm.ColDuration.Index).Value, DurasiLeave))
            Else
                If Durasi = EnumDurasi.Full_Day And Workday = EnumWorkday.Holiday Then
                    RowDgv.Cells(vFrm.ColAttendanceStatus.Index).Value = CboDgv.Items.Item(ClsTrans_Attendance.EnumStatus.Holiday)
                ElseIf Durasi = EnumDurasi.Full_Day And Workday = EnumWorkday.Holiday Then
                    RowDgv.Cells(vFrm.ColAttendanceStatus.Index).Value = CboDgv.Items.Item(ClsTrans_Attendance.EnumStatus.Half_Holiday)
                Else
                    RowDgv.Cells(vFrm.ColAttendanceStatus.Index).Value = CboDgv.Items.Item(ClsTrans_Attendance.EnumStatus.Present)
                End If
            End If

            'sch in
            RowDgv.Cells(vFrm.ColSchIn.Index).Value = Me.ObjCd.Convert_ToString(RowWork.schedule_in)
            'sch break out
            RowDgv.Cells(vFrm.ColSchBreakOut.Index).Value = Me.ObjCd.Convert_ToString(RowWork.schedule_break_out)
            'sch break in
            RowDgv.Cells(vFrm.ColSchBreakIn.Index).Value = Me.ObjCd.Convert_ToString(RowWork.schedule_break_in)
            'sch out
            RowDgv.Cells(vFrm.ColSchOut.Index).Value = Me.ObjCd.Convert_ToString(RowWork.schedule_out)
            'act in
            RowDgv.Cells(vFrm.ColActualIn.Index).Value = Me.ObjCd.Convert_ToString(RowWork.schedule_in)
            'act break out
            RowDgv.Cells(vFrm.ColActualBreakOut.Index).Value = Me.ObjCd.Convert_ToString(RowWork.schedule_break_out)
            'act break in
            RowDgv.Cells(vFrm.ColActualBreakIn.Index).Value = Me.ObjCd.Convert_ToString(RowWork.schedule_break_in)
            'act out
            RowDgv.Cells(vFrm.ColActualOut.Index).Value = Me.ObjCd.Convert_ToString(RowWork.schedule_out)
            'cek leave app
            'jika holiday ga usah dicek
            If Workday = EnumWorkday.Workday Then
                DtaLeave.FillBy_PrimaryEmpNDate(TblLeave, vDate, vFrm.TxtPrimaryEmployee.Text)
                If TblLeave.Rows.Count > 0 Then
                    Dim DurasiLeave As EnumDurasi
                    For Each rowLeave In TblLeave.Rows
                        RowDgv.Cells(vFrm.ColStatusRemark.Index).Value = rowLeave.reason
                        RowDgv.Cells(vFrm.ColLeaveCode.Index).Value = rowLeave.CutiCode
                        RowDgv.Cells(vFrm.ColLeaveDesc.Index).Value = rowLeave.CutiDesc
                        RowDgv.Cells(vFrm.ColLeaveDoc.Index).Value = rowLeave.no_transaksi
                        RowDgv.Cells(vFrm.ColPrimaryLeave.Index).Value = rowLeave.primary_cuti
                        RowDgv.Cells(vFrm.ColPrimaryLA.Index).Value = rowLeave.primary
                        'ubah status leave
                        CboDgv = RowDgv.Cells(vFrm.ColAttendanceStatus.Index)
                        'durasi ngikutin durasi di leave app
                        'sebelumnya cek dulu status di pengajuan cuti explain 
                        If vStatusExplain <> -1 And Workday <> EnumWorkday.Holiday Then
                            RowDgv.Cells(vFrm.ColAttendanceStatus.Index).Value = CboDgv.Items.Item(Me.GetStatusByExplain(vStatusExplain, vDuration, RowDgv.Cells(vFrm.ColDuration.Index).Value, DurasiLeave))
                            Dim setInOut As Boolean = False
                            If DurasiLeave = EnumDurasi.Full_Day Then
                                setInOut = True
                            Else
                                If RowWork.duratioin = EnumDurasi.Half_Day Then
                                    setInOut = True
                                End If
                            End If
                            If setInOut = True Then
                                'act in
                                RowDgv.Cells(vFrm.ColActualIn.Index).Value = Me.ObjCd.Convert_ToString(0)
                                'act break out
                                RowDgv.Cells(vFrm.ColActualBreakOut.Index).Value = Me.ObjCd.Convert_ToString(0)
                                'act break in
                                RowDgv.Cells(vFrm.ColActualBreakIn.Index).Value = Me.ObjCd.Convert_ToString(0)
                                'act out
                                RowDgv.Cells(vFrm.ColActualOut.Index).Value = Me.ObjCd.Convert_ToString(0)
                            End If
                        Else
                            If vDate = rowLeave.start_date Then
                                If rowLeave.half_day_start = ClsTrans_Cuti_Pengajuan.EnumHalfDay.FullDay Then
                                    DurasiLeave = EnumDurasi.Full_Day
                                ElseIf rowLeave.half_day_start = ClsTrans_Cuti_Pengajuan.EnumHalfDay.HalfDay Then
                                    DurasiLeave = EnumDurasi.Half_Day
                                End If
                            ElseIf vDate = rowLeave.end_date Then
                                If rowLeave.half_day_end = ClsTrans_Cuti_Pengajuan.EnumHalfDay.FullDay Then
                                    DurasiLeave = EnumDurasi.Full_Day
                                ElseIf rowLeave.half_day_end = ClsTrans_Cuti_Pengajuan.EnumHalfDay.HalfDay Then
                                    DurasiLeave = EnumDurasi.Half_Day
                                End If
                            Else
                                DurasiLeave = Durasi
                            End If

                            CboDgv = RowDgv.Cells(vFrm.ColAttendanceStatus.Index)
                            If DurasiLeave = EnumDurasi.Full_Day Then
                                RowDgv.Cells(vFrm.ColAttendanceStatus.Index).Value = CboDgv.Items.Item(ClsTrans_Attendance.EnumStatus.Leave)
                                'act in
                                RowDgv.Cells(vFrm.ColActualIn.Index).Value = Me.ObjCd.Convert_ToString(0)
                                'act break out
                                RowDgv.Cells(vFrm.ColActualBreakOut.Index).Value = Me.ObjCd.Convert_ToString(0)
                                'act break in
                                RowDgv.Cells(vFrm.ColActualBreakIn.Index).Value = Me.ObjCd.Convert_ToString(0)
                                'act out
                                RowDgv.Cells(vFrm.ColActualOut.Index).Value = Me.ObjCd.Convert_ToString(0)
                            ElseIf DurasiLeave = EnumDurasi.Half_Day Then
                                RowDgv.Cells(vFrm.ColAttendanceStatus.Index).Value = CboDgv.Items.Item(ClsTrans_Attendance.EnumStatus.Half_day_Leave)
                            End If
                        End If
                    Next
                End If
            End If
            'is on holiday list = 0
            RowDgv.Cells(vFrm.ColIsOnHolidayList.Index).Value = 0
            'type ot
            CboDgv = RowDgv.Cells(vFrm.ColOT.Index)
            RowDgv.Cells(vFrm.ColOT.Index).Value = CboDgv.Items.Item(TypeOt)
            'cek status
            Me.InitIsStatus(vFrm, RowDgv, Durasi)
            'masukin yang berkaitan dengan hitungan
            Me.InitHitungan(vFrm, RowDgv)
            'kosongkan yang fullday tidak masuk
            CboDgv = RowDgv.Cells(vFrm.ColAttendanceStatus.Index)
            Status = CboDgv.Items.IndexOf(RowDgv.Cells(vFrm.ColAttendanceStatus.Index).Value)
            Me.KosongkanFullTidakMasuk(vFrm, RowDgv, Status)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
        'Dim CboDgv As Windows.Forms.DataGridViewComboBoxCell
        'Dim Workday As EnumWorkday
        'Dim Durasi As EnumDurasi
        'Dim TypeOt As EnumTipeOt
        'Dim Status As EnumStatus
        'Dim rowLeave As DseTrans_Attendance.tbltrans_cuti_pengajuanRow
        'Try
        '    RowDgv.Cells(vFrm.ColShowLog.Index).Value = "Show"
        '    'primary ws
        '    RowDgv.Cells(vFrm.ColPrimaryWorkShift.Index).Value = RowWork.primary_transaksi
        '    'workshift
        '    CboDgv = RowDgv.Cells(vFrm.ColWorkdays.Index)
        '    If RowWork.tipe_workdays = EnumWorkday.Workday Then
        '        Workday = EnumWorkday.Workday
        '        RowDgv.Cells(vFrm.ColWorkdays.Index).Value = CboDgv.Items.Item(ClsTrans_Attendance.EnumWorkday.Workday)
        '        TypeOt = EnumTipeOt.Weekday_Overtime
        '    ElseIf RowWork.tipe_workdays = EnumWorkday.Holiday Then
        '        Workday = EnumWorkday.Holiday
        '        RowDgv.Cells(vFrm.ColWorkdays.Index).Value = CboDgv.Items.Item(ClsTrans_Attendance.EnumWorkday.Holiday)
        '        TypeOt = EnumTipeOt.Holiday_Overtime
        '    End If
        '    'durasi
        '    Durasi = RowWork.duratioin
        '    CboDgv = RowDgv.Cells(vFrm.ColDuration.Index)
        '    RowDgv.Cells(vFrm.ColDuration.Index).Value = CboDgv.Items.Item(RowWork.duratioin)
        '    'status
        '    CboDgv = RowDgv.Cells(vFrm.ColAttendanceStatus.Index)
        '    If Durasi = EnumDurasi.Full_Day And Workday = EnumWorkday.Holiday Then
        '        RowDgv.Cells(vFrm.ColAttendanceStatus.Index).Value = CboDgv.Items.Item(ClsTrans_Attendance.EnumStatus.Holiday)
        '    ElseIf Durasi = EnumDurasi.Full_Day And Workday = EnumWorkday.Holiday Then
        '        RowDgv.Cells(vFrm.ColAttendanceStatus.Index).Value = CboDgv.Items.Item(ClsTrans_Attendance.EnumStatus.Half_Holiday)
        '    Else
        '        RowDgv.Cells(vFrm.ColAttendanceStatus.Index).Value = CboDgv.Items.Item(ClsTrans_Attendance.EnumStatus.Present)
        '    End If
        '    'sch in
        '    RowDgv.Cells(vFrm.ColSchIn.Index).Value = Me.ObjCd.Convert_ToString(RowWork.schedule_in)
        '    'sch break out
        '    RowDgv.Cells(vFrm.ColSchBreakOut.Index).Value = Me.ObjCd.Convert_ToString(RowWork.schedule_break_out)
        '    'sch break in
        '    RowDgv.Cells(vFrm.ColSchBreakIn.Index).Value = Me.ObjCd.Convert_ToString(RowWork.schedule_break_in)
        '    'sch out
        '    RowDgv.Cells(vFrm.ColSchOut.Index).Value = Me.ObjCd.Convert_ToString(RowWork.schedule_out)
        '    'act in
        '    RowDgv.Cells(vFrm.ColActualIn.Index).Value = Me.ObjCd.Convert_ToString(RowWork.schedule_in)
        '    'act break out
        '    RowDgv.Cells(vFrm.ColActualBreakOut.Index).Value = Me.ObjCd.Convert_ToString(RowWork.schedule_break_out)
        '    'act break in
        '    RowDgv.Cells(vFrm.ColActualBreakIn.Index).Value = Me.ObjCd.Convert_ToString(RowWork.schedule_break_in)
        '    'act out
        '    RowDgv.Cells(vFrm.ColActualOut.Index).Value = Me.ObjCd.Convert_ToString(RowWork.schedule_out)
        '    'cek leave app
        '    'jika holiday ga usah dicek
        '    If Workday = EnumWorkday.Workday Then
        '        DtaLeave.FillBy_PrimaryEmpNDate(TblLeave, vDate, vFrm.TxtPrimaryEmployee.Text)
        '        If TblLeave.Rows.Count > 0 Then
        '            Dim DurasiLeave As EnumDurasi
        '            For Each rowLeave In TblLeave.Rows
        '                RowDgv.Cells(vFrm.ColStatusRemark.Index).Value = rowLeave.reason
        '                RowDgv.Cells(vFrm.ColLeaveCode.Index).Value = rowLeave.CutiCode
        '                RowDgv.Cells(vFrm.ColLeaveDesc.Index).Value = rowLeave.CutiDesc
        '                RowDgv.Cells(vFrm.ColLeaveDoc.Index).Value = rowLeave.no_transaksi
        '                RowDgv.Cells(vFrm.ColPrimaryLeave.Index).Value = rowLeave.primary_cuti
        '                RowDgv.Cells(vFrm.ColPrimaryLA.Index).Value = rowLeave.primary
        '                'ubah status leave
        '                CboDgv = RowDgv.Cells(vFrm.ColAttendanceStatus.Index)
        '                'durasi ngikutin durasi di leave app
        '                If vDate = rowLeave.start_date Then
        '                    If rowLeave.half_day_start = ClsTrans_Cuti_Pengajuan.EnumHalfDay.FullDay Then
        '                        DurasiLeave = EnumDurasi.Full_Day
        '                    ElseIf rowLeave.half_day_start = ClsTrans_Cuti_Pengajuan.EnumHalfDay.HalfDay Then
        '                        DurasiLeave = EnumDurasi.Half_Day
        '                    End If
        '                ElseIf vDate = rowLeave.end_date Then
        '                    If rowLeave.half_day_end = ClsTrans_Cuti_Pengajuan.EnumHalfDay.FullDay Then
        '                        DurasiLeave = EnumDurasi.Full_Day
        '                    ElseIf rowLeave.half_day_end = ClsTrans_Cuti_Pengajuan.EnumHalfDay.HalfDay Then
        '                        DurasiLeave = EnumDurasi.Half_Day
        '                    End If
        '                Else
        '                    DurasiLeave = Durasi
        '                End If
        '                CboDgv = RowDgv.Cells(vFrm.ColAttendanceStatus.Index)
        '                If DurasiLeave = EnumDurasi.Full_Day Then
        '                    RowDgv.Cells(vFrm.ColAttendanceStatus.Index).Value = CboDgv.Items.Item(ClsTrans_Attendance.EnumStatus.Leave)
        '                    'act in
        '                    RowDgv.Cells(vFrm.ColActualIn.Index).Value = Me.ObjCd.Convert_ToString(0)
        '                    'act break out
        '                    RowDgv.Cells(vFrm.ColActualBreakOut.Index).Value = Me.ObjCd.Convert_ToString(0)
        '                    'act break in
        '                    RowDgv.Cells(vFrm.ColActualBreakIn.Index).Value = Me.ObjCd.Convert_ToString(0)
        '                    'act out
        '                    RowDgv.Cells(vFrm.ColActualOut.Index).Value = Me.ObjCd.Convert_ToString(0)
        '                ElseIf DurasiLeave = EnumDurasi.Half_Day Then
        '                    RowDgv.Cells(vFrm.ColAttendanceStatus.Index).Value = CboDgv.Items.Item(ClsTrans_Attendance.EnumStatus.Half_day_Leave)
        '                End If
        '            Next
        '        End If
        '    End If
        '    'is on holiday list = 0
        '    RowDgv.Cells(vFrm.ColIsOnHolidayList.Index).Value = 0
        '    'type ot
        '    CboDgv = RowDgv.Cells(vFrm.ColOT.Index)
        '    RowDgv.Cells(vFrm.ColOT.Index).Value = CboDgv.Items.Item(TypeOt)
        '    'cek status
        '    Me.InitIsStatus(vFrm, RowDgv, Durasi)
        '    'masukin yang berkaitan dengan hitungan
        '    Me.InitHitungan(vFrm, RowDgv)
        '    'kosongkan yang fullday tidak masuk
        '    CboDgv = RowDgv.Cells(vFrm.ColAttendanceStatus.Index)
        '    Status = CboDgv.Items.IndexOf(RowDgv.Cells(vFrm.ColAttendanceStatus.Index).Value)
        '    Me.KosongkanFullTidakMasuk(vFrm, RowDgv, Status)
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'End Try
    End Sub

    ''' <summary>
    ''' ini hanya digunakan oleh programmer aja, kalo misalnya ada cuti yang ketinggalan.
    ''' Jangan kasih tau user ada fitur ini, nanti bikin ribet dan user manjadi tidak disiplin
    ''' </summary>
    ''' <param name="vFrm"></param>
    ''' <param name="RowDgv"></param>
    ''' <param name="vDate"></param>
    ''' <remarks></remarks>
    Public Sub InitLeave(ByVal vFrm As FrmTrans_Attendance, ByVal RowDgv As Windows.Forms.DataGridViewRow, _
    ByVal vDate As Date)
        Dim DtaLeave As New DseTrans_AttendanceTableAdapters.tbltrans_cuti_pengajuanTableAdapter
        Dim TblLeave As New DseTrans_Attendance.tbltrans_cuti_pengajuanDataTable
        Dim CboDgv As Windows.Forms.DataGridViewComboBoxCell
        Dim Durasi As EnumDurasi
        Dim rowLeave As DseTrans_Attendance.tbltrans_cuti_pengajuanRow
        Try
            DtaLeave.FillBy_PrimaryEmpNDate(TblLeave, vDate, vFrm.TxtPrimaryEmployee.Text)
            If TblLeave.Rows.Count > 0 Then
                Dim DurasiLeave As EnumDurasi
                For Each rowLeave In TblLeave.Rows
                    RowDgv.Cells(vFrm.ColStatusRemark.Index).Value = rowLeave.reason
                    RowDgv.Cells(vFrm.ColLeaveCode.Index).Value = rowLeave.CutiCode
                    RowDgv.Cells(vFrm.ColLeaveDesc.Index).Value = rowLeave.CutiDesc
                    RowDgv.Cells(vFrm.ColLeaveDoc.Index).Value = rowLeave.no_transaksi
                    RowDgv.Cells(vFrm.ColPrimaryLeave.Index).Value = rowLeave.primary_cuti
                    RowDgv.Cells(vFrm.ColPrimaryLA.Index).Value = rowLeave.primary
                    'ubah status leave
                    CboDgv = RowDgv.Cells(vFrm.ColAttendanceStatus.Index)
                    'durasi ngikutin durasi di leave app
                    If vDate = rowLeave.start_date Then
                        If rowLeave.half_day_start = ClsTrans_Cuti_Pengajuan.EnumHalfDay.FullDay Then
                            DurasiLeave = EnumDurasi.Full_Day
                        ElseIf rowLeave.half_day_start = ClsTrans_Cuti_Pengajuan.EnumHalfDay.HalfDay Then
                            DurasiLeave = EnumDurasi.Half_Day
                        End If
                    ElseIf vDate = rowLeave.end_date Then
                        If rowLeave.half_day_end = ClsTrans_Cuti_Pengajuan.EnumHalfDay.FullDay Then
                            DurasiLeave = EnumDurasi.Full_Day
                        ElseIf rowLeave.half_day_end = ClsTrans_Cuti_Pengajuan.EnumHalfDay.HalfDay Then
                            DurasiLeave = EnumDurasi.Half_Day
                        End If
                    Else
                        DurasiLeave = Durasi
                    End If
                    CboDgv = RowDgv.Cells(vFrm.ColAttendanceStatus.Index)
                    If DurasiLeave = EnumDurasi.Full_Day Then
                        RowDgv.Cells(vFrm.ColAttendanceStatus.Index).Value = CboDgv.Items.Item(ClsTrans_Attendance.EnumStatus.Leave)
                        'act in
                        RowDgv.Cells(vFrm.ColActualIn.Index).Value = Me.ObjCd.Convert_ToString(0)
                        'act break out
                        RowDgv.Cells(vFrm.ColActualBreakOut.Index).Value = Me.ObjCd.Convert_ToString(0)
                        'act break in
                        RowDgv.Cells(vFrm.ColActualBreakIn.Index).Value = Me.ObjCd.Convert_ToString(0)
                        'act out
                        RowDgv.Cells(vFrm.ColActualOut.Index).Value = Me.ObjCd.Convert_ToString(0)
                    ElseIf DurasiLeave = EnumDurasi.Half_Day Then
                        RowDgv.Cells(vFrm.ColAttendanceStatus.Index).Value = CboDgv.Items.Item(ClsTrans_Attendance.EnumStatus.Half_day_Leave)
                    End If
                Next
            End If
            'cek status
            Me.InitIsStatus(vFrm, RowDgv, Durasi)
            'masukin yang berkaitan dengan hitungan
            Me.InitHitungan(vFrm, RowDgv)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DtaLeave.Dispose()
            TblLeave.Dispose()
        End Try
    End Sub

    Private Sub InitIsStatus(ByVal vFrm As FrmTrans_Attendance, ByVal RowDgv As Windows.Forms.DataGridViewRow, _
    ByVal durasi As EnumDurasi)
        Dim CboDgv As Windows.Forms.DataGridViewComboBoxCell
        Try
            'ahdar
            'cek status
            CboDgv = RowDgv.Cells(vFrm.ColAttendanceStatus.Index)
            'is present
            If RowDgv.Cells(vFrm.ColAttendanceStatus.Index).Value = CboDgv.Items.Item(ClsTrans_Attendance.EnumStatus.Present) Then
                If durasi = EnumDurasi.Full_Day Then
                    RowDgv.Cells(vFrm.ColIsPresent.Index).Value = 1
                Else
                    RowDgv.Cells(vFrm.ColIsPresent.Index).Value = 0.5
                End If
            ElseIf RowDgv.Cells(vFrm.ColAttendanceStatus.Index).Value = CboDgv.Items.Item(ClsTrans_Attendance.EnumStatus.Half_Absent) Or _
              RowDgv.Cells(vFrm.ColAttendanceStatus.Index).Value = CboDgv.Items.Item(ClsTrans_Attendance.EnumStatus.Half_day_Leave) Or _
              RowDgv.Cells(vFrm.ColAttendanceStatus.Index).Value = CboDgv.Items.Item(ClsTrans_Attendance.EnumStatus.Half_Day_Permit) Or _
              RowDgv.Cells(vFrm.ColAttendanceStatus.Index).Value = CboDgv.Items.Item(ClsTrans_Attendance.EnumStatus.Half_Day_Sick) Or _
              RowDgv.Cells(vFrm.ColAttendanceStatus.Index).Value = CboDgv.Items.Item(ClsTrans_Attendance.EnumStatus.Half_Holiday) Then
                If durasi = EnumDurasi.Full_Day Then
                    RowDgv.Cells(vFrm.ColIsPresent.Index).Value = 0.5
                ElseIf durasi = EnumDurasi.Half_Day Then
                    RowDgv.Cells(vFrm.ColIsPresent.Index).Value = 0
                End If
            Else
                RowDgv.Cells(vFrm.ColIsPresent.Index).Value = 0
            End If
            'is sick
            If RowDgv.Cells(vFrm.ColAttendanceStatus.Index).Value = CboDgv.Items.Item(ClsTrans_Attendance.EnumStatus.Sick) Then
                RowDgv.Cells(vFrm.ColIsSick.Index).Value = 1
            ElseIf RowDgv.Cells(vFrm.ColAttendanceStatus.Index).Value = CboDgv.Items.Item(ClsTrans_Attendance.EnumStatus.Half_Day_Sick) Then
                RowDgv.Cells(vFrm.ColIsSick.Index).Value = 0.5
            Else
                RowDgv.Cells(vFrm.ColIsSick.Index).Value = 0
            End If
            'is leave
            If RowDgv.Cells(vFrm.ColAttendanceStatus.Index).Value = CboDgv.Items.Item(ClsTrans_Attendance.EnumStatus.Leave) Then
                RowDgv.Cells(vFrm.ColIsLeave.Index).Value = 1
            ElseIf RowDgv.Cells(vFrm.ColAttendanceStatus.Index).Value = CboDgv.Items.Item(ClsTrans_Attendance.EnumStatus.Half_day_Leave) Then
                RowDgv.Cells(vFrm.ColIsLeave.Index).Value = 0.5
            Else
                RowDgv.Cells(vFrm.ColIsLeave.Index).Value = 0
            End If
            'is alfa
            If RowDgv.Cells(vFrm.ColAttendanceStatus.Index).Value = CboDgv.Items.Item(ClsTrans_Attendance.EnumStatus.Absent) Then
                RowDgv.Cells(vFrm.ColIsAlfa.Index).Value = 1
            ElseIf RowDgv.Cells(vFrm.ColAttendanceStatus.Index).Value = CboDgv.Items.Item(ClsTrans_Attendance.EnumStatus.Half_Absent) Then
                RowDgv.Cells(vFrm.ColIsAlfa.Index).Value = 0.5
            Else
                RowDgv.Cells(vFrm.ColIsAlfa.Index).Value = 0
            End If
            'is permit
            If RowDgv.Cells(vFrm.ColAttendanceStatus.Index).Value = CboDgv.Items.Item(ClsTrans_Attendance.EnumStatus.Permit) Then
                RowDgv.Cells(vFrm.ColIsPermit.Index).Value = 1
            ElseIf RowDgv.Cells(vFrm.ColAttendanceStatus.Index).Value = CboDgv.Items.Item(ClsTrans_Attendance.EnumStatus.Half_Day_Permit) Then
                RowDgv.Cells(vFrm.ColIsPermit.Index).Value = 0.5
            Else
                RowDgv.Cells(vFrm.ColIsPermit.Index).Value = 0
            End If
            'is holiday
            If RowDgv.Cells(vFrm.ColAttendanceStatus.Index).Value = CboDgv.Items.Item(ClsTrans_Attendance.EnumStatus.Holiday) Then
                RowDgv.Cells(vFrm.ColIsHoliday.Index).Value = 1
            ElseIf RowDgv.Cells(vFrm.ColAttendanceStatus.Index).Value = CboDgv.Items.Item(ClsTrans_Attendance.EnumStatus.Half_Holiday) Then
                RowDgv.Cells(vFrm.ColIsHoliday.Index).Value = 0.5
            Else
                RowDgv.Cells(vFrm.ColIsHoliday.Index).Value = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub InitHitungan(ByVal vFrm As FrmTrans_Attendance, ByVal RowDgv As Windows.Forms.DataGridViewRow)
        Dim sett As String = ""
        Dim ActIn As Double = 0
        Dim ActOut As Double = 0
        Dim SchIn As Double = 0
        Dim SchOut As Double = 0
        Dim SchBreakOut As Double = 0
        Dim SchBreakIn As Double = 0
        Dim ActBreakOut As Double = 0
        Dim ActBreakIn As Double = 0
        Dim ActWork As Double = 0
        Dim ActBreak As Double = 0
        Dim SchWork As Double = 0
        Dim SchBreak As Double = 0
        Dim vDummy As String = ""

        Try
            ActIn = Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(vFrm.ColActualIn.Index).Value)
            ActOut = Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(vFrm.ColActualOut.Index).Value)
            SchIn = Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(vFrm.ColSchIn.Index).Value)
            SchOut = Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(vFrm.ColSchOut.Index).Value)
            SchBreakOut = Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(vFrm.ColSchBreakOut.Index).Value)
            SchBreakIn = Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(vFrm.ColSchBreakIn.Index).Value)
            ActBreakOut = Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(vFrm.ColActualBreakOut.Index).Value)
            ActBreakIn = Me.ObjCd.Convert_ToMinutes(RowDgv.Cells(vFrm.ColActualBreakIn.Index).Value)
            sett = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, Core_Module.ClsTransaksiSetting.EnumNamaParameter.IgnoreBreakData)

            'Sch Break Duration & act break duration
            If sett = "1" Then
                vDummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Break_duration)
                If vDummy = "" Then
                    RowDgv.Cells(vFrm.ColSchBreakDuration.Index).Value = Me.ObjCd.Convert_ToString(0)
                Else
                    RowDgv.Cells(vFrm.ColSchBreakDuration.Index).Value = Me.ObjCd.Convert_ToString(vDummy)
                End If
                vDummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.HR_Attendance, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Break_duration)
                If vDummy = "" Then
                    RowDgv.Cells(vFrm.ColActualBreakDuration.Index).Value = Me.ObjCd.Convert_ToString(0)
                Else
                    RowDgv.Cells(vFrm.ColActualBreakDuration.Index).Value = Me.ObjCd.Convert_ToString(vDummy)
                End If
            Else
                If SchBreakOut > SchBreakIn Then
                    SchBreak = ((24 * 60) - SchBreakOut) + SchBreakIn
                Else
                    SchBreak = SchBreakIn - SchBreakOut
                End If
                RowDgv.Cells(vFrm.ColSchBreakDuration.Index).Value = Me.ObjCd.Convert_ToString(SchBreak)
                If ActBreakOut > ActBreakIn Then
                    ActBreak = ((24 * 60) - ActBreakOut) + ActBreakIn
                Else
                    ActBreak = ActBreakIn - ActBreakOut
                End If
                RowDgv.Cells(vFrm.ColActualBreakDuration.Index).Value = Me.ObjCd.Convert_ToString(ActBreak)
            End If

            'sch work duration
            If SchOut > SchIn Then

                RowDgv.Cells(vFrm.ColSchWorkDuration.Index).Value = Me.ObjCd.Convert_ToString((SchOut - SchIn) - SchBreak)
            Else
                RowDgv.Cells(vFrm.ColSchWorkDuration.Index).Value = Me.ObjCd.Convert_ToString((((24 * 60) - SchIn) + SchOut) - SchBreak)
            End If

            'act work duration
            If ActOut > ActIn Then
                RowDgv.Cells(vFrm.ColActualWorkDuration.Index).Value = Me.ObjCd.Convert_ToString((ActOut - ActIn) - ActBreak)
            Else
                RowDgv.Cells(vFrm.ColActualWorkDuration.Index).Value = Me.ObjCd.Convert_ToString((((24 * 60) - ActIn) + ActOut) - ActBreak)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function IsiDGV(ByVal Criteria As String, ByVal DGV As Windows.Forms.DataGridView, _
    ByVal TBL As DseTrans_Attendance.DataTable4DataTable, ByVal SearchCriteria As ClsTrans_Attendance.EnumSearch, _
    ByVal tsl As Windows.Forms.ToolStripLabel, ByVal dta As DseTrans_AttendanceTableAdapters.DataTable4TableAdapter, _
    ByVal vStrDate As Date, ByVal vEndDate As Date) As Boolean
        Dim QueryResult As Double
        Dim StartT As Date

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            StartT = Now
            Windows.Forms.Cursor.Current = Windows.Forms.Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = Windows.Forms.DataGridViewAutoSizeColumnsMode.None

            If SearchCriteria = EnumSearch.DocNo Then
                QueryResult = dta.FillBy_LikeDocNo(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearch.Tgl Then
                QueryResult = dta.FillBy_Date(TBL, vStrDate, vEndDate)
            ElseIf SearchCriteria = EnumSearch.EmpCode Then
                QueryResult = dta.FillBy_LikeEmpCode(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearch.EmpName Then
                QueryResult = dta.FillBy_LIKEName(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearch.Remarks Then
                QueryResult = dta.FillBy_LikeRemarks(TBL, Criteria + "%")
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()

            Dim TS As TimeSpan = Now - StartT
            If QueryResult > 0 Then
                If TS.TotalSeconds > 1 Then
                    tsl.Text = CStr(QueryResult) + " result(s) in " + ObjCd.DigitGrouping(TS.TotalSeconds, 2) + " seconds"
                Else
                    tsl.Text = CStr(QueryResult) + " result(s) in " + ObjCd.DigitGrouping(TS.TotalMilliseconds, 2, MidpointRounding.AwayFromZero) + " miliseconds"
                End If
            Else
                tsl.Text = "No result found"
            End If

            DGV.Columns(0).Visible = False
            DGV.Columns(1).Visible = False
            DGV.Columns(2).HeaderText = "Document No"
            DGV.Columns(3).HeaderText = "Date"
            DGV.Columns(3).DefaultCellStyle.Format = "dd MMM yyyy"
            DGV.Columns(4).HeaderText = "Employee Code"
            DGV.Columns(5).HeaderText = "Name"
            DGV.Columns(6).HeaderText = "Start Date"
            DGV.Columns(7).HeaderText = "End Date"
            DGV.Columns(6).DefaultCellStyle.Format = "dd MMM yyyy"
            DGV.Columns(7).DefaultCellStyle.Format = "dd MMM yyyy"
            DGV.Columns(8).HeaderText = "Remarks"

            DGV.AutoSizeColumnsMode = Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
            Windows.Forms.Cursor.Current = Windows.Forms.Cursors.Arrow
        End Try
    End Function

    Private Sub KosongkanNonvisibleDGv(ByVal vFrm As FrmTrans_Attendance, _
    ByVal RowDgv As Windows.Forms.DataGridViewRow)
        Try
            'break data
            If vFrm.CheckBox1.Checked = True Then
                RowDgv.Cells(vFrm.ColActualBreakOut.Index).Value = Me.ObjCd.Convert_ToString(0)
                RowDgv.Cells(vFrm.ColActualBreakIn.Index).Value = Me.ObjCd.Convert_ToString(0)
                RowDgv.Cells(vFrm.ColEarlyBreakOutDuration.Index).Value = Me.ObjCd.Convert_ToString(0)
                RowDgv.Cells(vFrm.ColLateBreakInDuration.Index).Value = Me.ObjCd.Convert_ToString(0)
            End If
            'early in
            If vFrm.CheckBox2.Checked = True Then
                RowDgv.Cells(vFrm.ColEarlyInDuration.Index).Value = Me.ObjCd.Convert_ToString(0)
                RowDgv.Cells(vFrm.ColEarlyOTDuration.Index).Value = Me.ObjCd.Convert_ToString(0)
            End If
            'early out
            If vFrm.CheckBox3.Checked = True Then
                RowDgv.Cells(vFrm.ColEarlyOutDuration.Index).Value = Me.ObjCd.Convert_ToString(0)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanFullTidakMasuk(ByVal vFrm As FrmTrans_Attendance, _
    ByVal RowDgv As Windows.Forms.DataGridViewRow, ByVal Status As EnumStatus)
        Try
            If Me.CekHalfStatus(Status, True) = False Then
                RowDgv.Cells(vFrm.ColActualIn.Index).Value = Me.ObjCd.Convert_ToString(0)
                RowDgv.Cells(vFrm.ColActualOut.Index).Value = Me.ObjCd.Convert_ToString(0)
                RowDgv.Cells(vFrm.ColActualBreakOut.Index).Value = Me.ObjCd.Convert_ToString(0)
                RowDgv.Cells(vFrm.ColActualBreakIn.Index).Value = Me.ObjCd.Convert_ToString(0)
                RowDgv.Cells(vFrm.ColActualWorkDuration.Index).Value = Me.ObjCd.Convert_ToString(0)
                RowDgv.Cells(vFrm.ColActualBreakDuration.Index).Value = Me.ObjCd.Convert_ToString(0)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As ClsTrans_Attendance.EnumSearch, _
    ByVal vStrDate As Date, ByVal vEndDate As Date) As Double
        Dim Dummy As Nullable(Of Long)
        Dim dta As New DseTrans_AttendanceTableAdapters.tbl_trans_att_cari_viewTableAdapter

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If vSearchBy = EnumSearch.DocNo Then
                Dummy = dta.ScalarQuery_GetJmlLikeDocNo(vKriteria + "%")
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearch.Tgl Then
                Dummy = dta.ScalarQuery_GetJmlByDate(vStrDate, vEndDate)
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearch.EmpCode Then
                Dummy = dta.ScalarQuery_GetJmlLikeEmpCode(vKriteria + "%")
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearch.EmpName Then
                Dummy = dta.ScalarQuery_GetJmlByName(vKriteria + "%")
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearch.Remarks Then
                Dummy = dta.ScalarQuery_GetJmlLikeRemarks(vKriteria + "%")
                If Dummy.HasValue = True Then
                    Return Dummy.Value
                Else
                    Return 0
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Overridable Function SimpanBaru(ByVal vUserCode As String, ByVal vDGV As Windows.Forms.DataGridView, _
    ByVal frm As FrmTrans_Attendance, ByVal TblOT_Order As DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_ot_1DataTable) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            Core_Module.SetConstring_DTAAntrian(Me.DTAAntrian)
            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, Me.vTipeAntrian, _
            Me.vPrimary_Antrian, Me.mPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_Attd_SimpanBaru, vUserCode)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > Me.ObjAntrian.GetWaitTime Then
                    If SW.Elapsed.TotalSeconds >= Me.ObjAntrian.GetWaitTime Then
                        Me.DTAAntrian.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If Me.DTAAntrian.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanBaru1(vUserCode, vDGV, frm, TblOT_Order)
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

    Public Function SimpanBaru1(ByVal vUserCode As String, ByVal vDGV As Windows.Forms.DataGridView, _
    ByVal frm As FrmTrans_Attendance, ByVal TblOT_Order As DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_ot_1DataTable) As Boolean
        Dim DTA As New DseTrans_AttendanceTableAdapters.tbltrans_attendanceTableAdapter
        Dim Urutan As Integer = 0
        Dim ListCuti As New List(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Urutan = Me.mNoTrans.Substring(0, 6)
            If DTA.Insert(Me.mNoTrans, Urutan, Me.mTanggal, Me.mPrimaryEmp, Me.mStrDate, Me.mEndDate, Me.mRemarks, _
             vUserCode, Now, vUserCode, Now, Now) > 0 Then
                Me.mPrimary = DTA.ScalarQuery_GetPrimaryByNoTrans(Me.mNoTrans)
                Me.mTimestamp = DTA.ScalarQuery_GetTimeStampByPrimary(Me.mPrimary)
                Me.mCreatedDate = Me.mTimestamp
                Me.mModifiedDate = Me.mTimestamp
                Me.mCreatedBy = vUserCode
                Me.mModifiedBy = vUserCode

                'simpan item
                ListCuti.Clear()
                Me.ObjItem.SimpanItem(Me.Primary, vDGV, frm, ListCuti, TblOT_Order, False, Me.mPrimaryEmp)

                'set balance
                Me.ObjCuti.SetBalance(Me.PrimaryEmp, ListCuti, Me.ObjCd.GetFisrtDateOfTheYear(Me.mStrDate))

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

    Public Overridable Function SimpanPerubahan(ByVal vUserCode As String, ByVal vDGV As Windows.Forms.DataGridView, _
    ByVal frm As FrmTrans_Attendance, ByVal TblOT_Order As DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_ot_1DataTable) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            Core_Module.SetConstring_DTAAntrian(DTAAntrian)
            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, Me.vTipeAntrian, _
            Me.vPrimary_Antrian, Me.mPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_Attd_SimpanPerubahan, vUserCode)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > Me.ObjAntrian.GetWaitTime Then
                    If SW.Elapsed.TotalSeconds >= ObjAntrian.GetWaitTime Then
                        Me.DTAAntrian.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If Me.DTAAntrian.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanPerubahan1(vUserCode, vDGV, frm, TblOT_Order)
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

    Public Function SimpanPerubahan1(ByVal vUserCode As String, ByVal vDGV As Windows.Forms.DataGridView, ByVal frm As FrmTrans_Attendance, _
    ByVal TblOT_Order As DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_ot_1DataTable) As Boolean
        Dim DTA As New DseTrans_AttendanceTableAdapters.tbltrans_attendanceTableAdapter
        Dim DtaItem As New DseTrans_AttendanceTableAdapters.DataTable5TableAdapter
        Dim TblItem As New DseTrans_Attendance.DataTable5DataTable
        Dim row As DseTrans_Attendance.DataTable5Row
        Dim Urutan As Integer = 0
        Dim ListCuti As New List(Of Double)
        Dim PrimaryKaryawanLama As Double = -1
        Try
            If Me.ApakahBolehEdit(Me.Primary, vUserCode) = True Then
                Urutan = Me.mNoTrans.Substring(0, 6)
                DTA.Connection.ConnectionString = Core_Module.GetConstring
                'cari primary karyawan lama
                PrimaryKaryawanLama = DTA.ScalarQuery_GetPrimaryKaryawanByPrimary(Me.mPrimary)
                If DTA.Update(Me.mNoTrans, Urutan, Me.mTanggal, Me.mPrimaryEmp, Me.mStrDate, Me.mEndDate, Me.mRemarks, _
                Me.mModifiedBy, Now, Now, Me.mPrimary, Me.mTimestamp) > 0 Then
                    Me.mTimestamp = DTA.ScalarQuery_GetTimeStampByPrimary(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode
                    'cari list sebelumnya
                    ListCuti.Clear()
                    DtaItem.FillBy_PrimaryTransaksi(TblItem, Me.mPrimary)
                    For Each row In TblItem.Rows
                        ListCuti.Add(row.primary_cuti)
                    Next
                    'hapus item
                    Me.ObjItem.HapusItem(Me.mPrimary, frm, TblOT_Order, TblItem, Me.PrimaryEmp, True)
                    'set balance yang lama
                    Me.ObjCuti.SetBalance(PrimaryKaryawanLama, ListCuti, Me.ObjCd.GetFisrtDateOfTheYear(Me.mStrDate))

                    ListCuti.Clear()
                    'simpan item
                    Me.ObjItem.SimpanItem(Me.Primary, vDGV, frm, ListCuti, TblOT_Order, False, Me.mPrimaryEmp)
                    'set balance yang baru 
                    Me.ObjCuti.SetBalance(Me.PrimaryEmp, ListCuti, Me.ObjCd.GetFisrtDateOfTheYear(Me.mStrDate))

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
            DtaItem.Dispose()
            TblItem.Dispose()
        End Try
    End Function

    Public Function Hapus(ByVal vPrimaryTransaksi As Double, ByVal vUserCode As String, ByVal Frm As FrmTrans_Attendance, _
    ByVal TblOT_Order As DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_ot_1DataTable, ByVal UpdByTbl As Boolean) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            Core_Module.SetConstring_DTAAntrian(DTAAntrian)
            'ambil no antrian
            Me.ObjAntrian.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAAntrian, Me.vTipeAntrian, _
            Me.vPrimary_Antrian, Me.mPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.Trans_Attd_Hapus, vUserCode)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > Me.ObjAntrian.GetWaitTime Then
                    If SW.Elapsed.TotalSeconds >= Me.ObjAntrian.GetWaitTime Then
                        Me.DTAAntrian.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If Me.DTAAntrian.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.Hapus1(vPrimaryTransaksi, vUserCode, Frm, TblOT_Order, UpdByTbl)
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

    Public Function Hapus1(ByVal vPrimary As Double, ByVal vUserCode As String, ByVal Frm As FrmTrans_Attendance, _
    ByVal TblOT_Order As DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_ot_1DataTable, _
    ByVal UpdByTbl As Boolean) As Boolean
        Dim DTA As New DseTrans_AttendanceTableAdapters.tbltrans_attendanceTableAdapter
        Dim DtaItem As New DseTrans_AttendanceTableAdapters.DataTable5TableAdapter
        Dim TblItem As New DseTrans_Attendance.DataTable5DataTable
        Dim row As DseTrans_Attendance.DataTable5Row
        Dim ListCuti As New List(Of Double)
        Dim PrimaryKaryawanLama As Double = -1

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahBolehHapus(vPrimary, vUserCode) = True Then
                'cari primary karyawan lama
                PrimaryKaryawanLama = DTA.ScalarQuery_GetPrimaryKaryawanByPrimary(Me.mPrimary)
                'list cuti
                ListCuti.Clear()
                DtaItem.FillBy_PrimaryTransaksi(TblItem, Me.mPrimary)
                For Each row In TblItem.Rows
                    ListCuti.Add(row.primary_cuti)
                Next

                'hapus header, item otomatis dihapuskan
                DTA.Delete(vPrimary)

                'set balance yang lama
                Me.ObjCuti.SetBalance(Me.PrimaryEmp, ListCuti, Me.ObjCd.GetFisrtDateOfTheYear(Me.mStrDate))

                '################## update juga ke overtime order item approval ot
                Me.ObjItem.UpdateOvertimeOrder(vPrimary, Frm, TblOT_Order, TblItem, Me.PrimaryEmp, UpdByTbl, True)
                '################## update juga ke overtime order item approval ot

                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DtaItem.Dispose()
            TblItem.Dispose()
        End Try
    End Function

    Public Function GetJmlTransAtt(ByVal vPrimaryAtt As Double) As Double
        Dim DTA As New DseTrans_AttendanceTableAdapters.tbltrans_attendanceTableAdapter
        Dim Dummy As System.Nullable(Of Long)
        Try
            Dummy = DTA.ScalarQuery_GetJmlByPrimaryAttendance(vPrimaryAtt)
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

    Public Function GetStartPeriod(ByVal vPrimary As Double) As Date
        Dim DTA As New DseTrans_AttendanceTableAdapters.tbltrans_attendanceTableAdapter
        Dim Dummy As System.Nullable(Of Date)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetStartPeriod(vPrimary)
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

    Public Function GetEndPeriod(ByVal vPrimary As Double) As Date
        'Dim DTA As New DseTrans_AttendanceTableAdapters.tbltrans_attendanceTableAdapter
        'tidak di buat langsung dsetrans_attendance karena saat edit dse lalu simpan selalu error
        Dim DTA As New DseTrans_SPLTableAdapters.tbltrans_spl_overtime_order_scalarTableAdapter
        Dim Dummy As System.Nullable(Of Date)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetEndPeriodAttendance(vPrimary)
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

    'aul tambahan
    Public Function GetPrimaryPrevAttendance(ByVal vPrimaryEmployee As Double, ByVal vStartDate As Date) As Double
        Dim Dta As New DseTrans_AttendanceTableAdapters.tbltrans_attendanceTableAdapter
        Dim EndDatePrev As Date = vStartDate.Date - TimeSpan.FromDays(1)
        Dim PrimaryAttendance As Nullable(Of Double)

        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            PrimaryAttendance = Dta.ScalarQuery_GetPrimaryByPrimaryEmpAndEndDate(vPrimaryEmployee, EndDatePrev)
            If PrimaryAttendance.HasValue = True Then
                Return PrimaryAttendance.Value
            End If

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Sub KosongkanObject()

        Try
            Me.mPrimary = -1
            Me.mNoTrans = ""
            Me.mTanggal = #1/1/1950#
            Me.mPrimaryEmp = -1
            Me.mStrDate = #1/1/1950#
            Me.mEndDate = #1/1/1950#
            Me.Remarks = ""
            Me.mCreatedBy = ""
            Me.mCreatedDate = #1/1/1950#
            Me.mModifiedBy = ""
            Me.mModifiedDate = #1/1/1950#
            Me.mTimestamp = #1/1/1950#
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'create by ahdar
    Private Function GetStatusByExplain(ByVal vStatusExplain As Integer, ByVal vDuration As Decimal, ByVal vDuration2 As String, ByRef vEnumDuration As EnumDurasi) As ClsTrans_Attendance.EnumStatus
        Dim vStatus As ClsTrans_Attendance.EnumStatus = EnumStatus.Present
        Try
            If vStatusExplain = ClsTrans_Cuti_Pengajuan.EnumStatusExplain.Leave Then
                If vDuration = 0.5 Then
                    vStatus = EnumStatus.Half_day_Leave
                    vEnumDuration = EnumDurasi.Half_Day
                Else
                    vStatus = EnumStatus.Leave
                    vEnumDuration = EnumDurasi.Full_Day
                End If
            ElseIf vStatusExplain = ClsTrans_Cuti_Pengajuan.EnumStatusExplain.Sick Then
                If vDuration2 <> "Full Day" Then
                    vStatus = EnumStatus.Half_Day_Sick
                    vEnumDuration = EnumDurasi.Half_Day
                Else
                    vStatus = EnumStatus.Sick
                    vEnumDuration = EnumDurasi.Full_Day
                End If
            ElseIf vStatusExplain = ClsTrans_Cuti_Pengajuan.EnumStatusExplain.Permit Then
                If vDuration2 <> "Full Day" Then
                    vStatus = EnumStatus.Half_Day_Permit
                    vEnumDuration = EnumDurasi.Half_Day
                Else
                    vStatus = EnumStatus.Permit
                    vEnumDuration = EnumDurasi.Full_Day
                End If
            ElseIf vStatusExplain = ClsTrans_Cuti_Pengajuan.EnumStatusExplain.Absent Then
                If vDuration2 <> "Full Day" Then
                    vStatus = EnumStatus.Half_Absent
                    vEnumDuration = EnumDurasi.Half_Day
                Else
                    vStatus = EnumStatus.Absent
                    vEnumDuration = EnumDurasi.Full_Day
                End If
            End If
            Return vStatus
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
End Class
