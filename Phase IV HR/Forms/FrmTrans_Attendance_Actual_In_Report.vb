Option Explicit On
Option Strict Off

Imports Core_Module
Imports System.Windows.Forms
Public Class FrmTrans_Attendance_Actual_In_Report

    Public UserCode As String = ""
    Public BackColorAll As Drawing.Color = Core_Module.GetSMESSBackColorAll
    Dim ObjCd As New Core_Module.ClsCekData
    Dim ObjSrv As New ClsMasterAttendanceServer
    Dim ObjSettings As New Core_Module.ClsTransaksiSetting
    Dim ObjSettingsUser As New Core_Module.ClsTransaksiSettingUser
    Dim Report As New CrpTrans_Attendance_Actual_In_Report

    Dim DS As New DseTrans_Attendance_Actual_In_Report
    Dim Tbl As DseTrans_Attendance_Actual_In_Report.tblLaporanClockInOutDataTable = DS.tblLaporanClockInOut
    Dim row As DseTrans_Attendance_Actual_In_Report.tblLaporanClockInOutRow

    Private Sub FrmTrans_Attendance_Actual_In_Report_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.DgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.LoadDefaultTimeSettings()
            Me.LoadLocation()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub FrmTrans_Attendance_Actual_In_Report_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            Me.ObjSettings.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Default_TimeFrom_ActualClockInOutReport, "1950-01-01 " & Format(Me.DtpTimeFrom.Value, "HH:mm") & ":00")
            Me.ObjSettings.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Default_TimeTo_ActualClockInOutReport, "1950-01-01 " & Format(Me.DtpTimeTo.Value, "HH:mm") & ":00")
            If Me.CboIfMultiple.SelectedIndex = 0 Then
                Me.ObjSettings.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Default_IfMultipleClockInOut_ActualClockInOutReport, "0")
            Else
                Me.ObjSettings.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Default_IfMultipleClockInOut_ActualClockInOutReport, "1")
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub LoadDefaultTimeSettings()
        Dim DefaultTimeFrom As String = ""
        Dim DefaultTimeTo As String = ""
        Dim DefaultIfMultipleClock As String = ""
        Try
            DefaultTimeFrom = Me.ObjSettings.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Default_TimeFrom_ActualClockInOutReport)
            DefaultTimeTo = Me.ObjSettings.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Default_TimeTo_ActualClockInOutReport)
            DefaultIfMultipleClock = Me.ObjSettings.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Default_IfMultipleClockInOut_ActualClockInOutReport)

            If String.IsNullOrEmpty(DefaultTimeFrom) = True Then
                Me.ObjSettings.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Default_TimeFrom_ActualClockInOutReport, "1950-01-01 01:00:00")
                DefaultTimeFrom = "1950-01-01 01:00:00"
            End If
            If String.IsNullOrEmpty(DefaultTimeTo) = True Then
                Me.ObjSettings.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Default_TimeTo_ActualClockInOutReport, "1950-01-01 23:00:00")
                DefaultTimeTo = "1950-01-01 23:00:00"
            End If
            If String.IsNullOrEmpty(Me.ObjSettings.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Default_IfMultipleClockInOut_ActualClockInOutReport)) = True Then
                Me.ObjSettings.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Default_IfMultipleClockInOut_ActualClockInOutReport, "0")
                DefaultIfMultipleClock = "0"
            End If

            Me.DtpTimeFrom.Value = CDate(DefaultTimeFrom)
            Me.DtpTimeTo.Value = CDate(DefaultTimeTo)
            If DefaultIfMultipleClock = "0" Then
                Me.CboIfMultiple.SelectedIndex = 0
            Else
                Me.CboIfMultiple.SelectedIndex = 1
            End If

            'tambahan
            Me.CboFormat.SelectedIndex = 0
            Me.CboGroupBy.SelectedIndex = 0
            Me.CrvActulaIn.DisplayStatusBar = True
            Me.CrvActulaIn.DisplayToolbar = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub LoadLocation()
        Dim ListLocation As New List(Of String)
        Try
            Me.Cursor = Cursors.WaitCursor
            ListLocation = Me.ObjSrv.GetListLocationTerminal(True)
            Me.ClbLocation.Items.Clear()
            Me.ClbLocation.Items.Add("From File", System.Windows.Forms.CheckState.Checked)
            If ListLocation.Count > 0 Then
                For Each location As String In ListLocation
                    If String.IsNullOrEmpty(location) = False Then
                        Me.ClbLocation.Items.Add(location, System.Windows.Forms.CheckState.Checked)
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub BtnShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnShow.Click
        Dim SelectedCount_ As Integer = 0
        Try
			If CDate(Format(CDate("1950-01-01 " & Format(Me.DtpTimeFrom.Value, "HH:mm") & ":00"))) > CDate(Format(CDate("1950-01-01 " & Format(Me.DtpTimeTo.Value, "HH:mm") & ":00"))) Then
                MsgBox("Invalid time range.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            For Each items_ As Object In Me.ClbLocation.CheckedItems
                SelectedCount_ += 1
            Next
            If SelectedCount_ < 1 Then
                MsgBox("Please select at least 1 location.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            If Me.CboFormat.SelectedIndex = 0 Then
                If Me.DtpDate.Value.Date > Me.DtpDate.Value.Date Then
                    MsgBox("Invalid time range.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
                Me.Cursor = Cursors.WaitCursor
                ShowLog2()
            Else
                ShowLog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub BtnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub IsiDataHoliday(ByVal vDateFrom As Date, ByVal vDateTo As Date, _
    ByVal TblWorkshift As DseTrans_Attendance_Actual_In_Report.tblmaster_workshift_itemallocateDataTable)
        Dim rowworkshift As DseTrans_Attendance_Actual_In_Report.tblmaster_workshift_itemallocateRow
        Dim DtaWSItemAlloc As New DseTrans_Attendance_Actual_In_ReportTableAdapters.tblmaster_workshift_itemallocateTableAdapter
        Dim TblWSItemAlloc As New DseTrans_Attendance_Actual_In_Report.tblmaster_workshift_itemallocateDataTable
        Dim rowWSItemAlloc As DseTrans_Attendance_Actual_In_Report.tblmaster_workshift_itemallocateRow
        Dim DtaLeave As New DseTrans_Attendance_Actual_In_ReportTableAdapters.tbl_leaveTableAdapter
        Dim TblLeave As New DseTrans_Attendance_Actual_In_Report.tbl_leaveDataTable
        Dim rowLeave As DseTrans_Attendance_Actual_In_Report.tbl_leaveRow

        Try
            DtaWSItemAlloc.Connection.ConnectionString = core_module.getconstring
            DtaWSItemAlloc.FillBy_StartAndEnd(TblWSItemAlloc, vDateTo, vDateFrom)
            DtaLeave.Connection.ConnectionString = core_module.getconstring
            DtaLeave.Fill(TblLeave, vDateFrom, vDateTo, 0)

            TblWSItemAlloc.DefaultView.Sort = "[primary transaksi]"
            TblLeave.DefaultView.Sort = "[]"

            TblWorkshift.DefaultView.RowFilter = ""
            TblWorkshift.DefaultView.Sort = "[primary entity]"
            For i As Integer = 0 To TblWorkshift.DefaultView.Count - 1
                rowworkshift = TblWorkshift.DefaultView(i).Row
                With rowworkshift

                End With
            Next

            TblWorkshift.DefaultView.Sort = "[primary workshift]"
            For i As Integer = 0 To TblWorkshift.DefaultView.Count - 1
                rowworkshift = TblWorkshift.DefaultView(i).Row
                With rowworkshift

                End With
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ShowLog2()
        Dim DTA As New DseTrans_Attendance_Actual_In_ReportTableAdapters.tblmaster_fingerprint_server_machinesTableAdapter
        Dim DTAKaryawan As New DseTrans_Attendance_Actual_In_ReportTableAdapters.tblKaryawanTableAdapter
        Dim DTAEmployee As New DseTrans_Attendance_Actual_In_ReportTableAdapters.tblmaster_karyawanTableAdapter
        Dim DTAWorkshift As New DseTrans_Attendance_Actual_In_ReportTableAdapters.tblmaster_workshift_itemallocateTableAdapter
        Dim DTAWorkshiftItem As New DseTrans_Attendance_Actual_In_ReportTableAdapters.tblmaster_workshift_itemparametersTableAdapter

        Dim TblEmployee As New DseTrans_Attendance_Actual_In_Report.tblmaster_karyawanDataTable
        Dim TblAttendance, TblKaryDummy As New DseTrans_Attendance_Actual_In_Report.tblKaryawanDataTable
        Dim TblWorkshift As New DseTrans_Attendance_Actual_In_Report.tblmaster_workshift_itemallocateDataTable
        Dim TblWorksfhitItem As New DseTrans_Attendance_Actual_In_Report.tblmaster_workshift_itemparametersDataTable

        Dim rowAttendance, rowKaryDummy As DseTrans_Attendance_Actual_In_Report.tblKaryawanRow
        Dim rowEmployee As DseTrans_Attendance_Actual_In_Report.tblmaster_karyawanRow
        Dim rowworkshift As DseTrans_Attendance_Actual_In_Report.tblmaster_workshift_itemallocateRow
        Dim rowworkshiftItem As DseTrans_Attendance_Actual_In_Report.tblmaster_workshift_itemparametersRow

        Dim listKary As New List(Of String)
        Dim StartDate As Date
        Dim EndDate As Date
        Dim tanggalAcuan As Date
        Dim JamMasuk As Double = 0
        Dim JamKeluar As Double = 0
        Dim JamX As Double = 0
        Dim ActualInDone As Boolean = False
        Dim ActualOutDone As Boolean = False
        Dim ActualBreakinDone As Boolean = False
        Dim ActualBreakOutDone As Boolean = False
        Dim counter As Integer = 0
        Dim report1 As New CrpLaporanClockInOutGroupByEmp
        Dim report2 As New CrpLaporanClockInOutGroupByShift

        Dim ActOutLintasHari As Date = Now

        StartDate = Me.DtpDate.Value.Date + " " + "00:00:00" 'Format(Me.DtpDate.Value.Date, "MM/dd/yyyy") + " " + "00:00:00"
        EndDate = Me.DtpDateTo.Value.Date + " " + "23:59:59" 'Format(Me.DtpDateTo.Value.Date, "MM/dd/yyyy") + " " + "23:59:59"

        Dim DtaHoliday As New DseTrans_Attendance_Actual_In_ReportTableAdapters.tblmaster_workshift_itemholidayTableAdapter
        Dim TblHoliday As New DseTrans_Attendance_Actual_In_Report.tblmaster_workshift_itemholidayDataTable
        Dim rowHoliday As DseTrans_Attendance_Actual_In_Report.tblmaster_workshift_itemholidayRow
        Dim DtaLeave As New DseTrans_Attendance_Actual_In_ReportTableAdapters.tbl_leaveTableAdapter
        Dim TblLeave As New DseTrans_Attendance_Actual_In_Report.tbl_leaveDataTable
        Dim rowLeave As DseTrans_Attendance_Actual_In_Report.tbl_leaveRow
        Dim StatReason As String = "Workday"

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTAKaryawan.Connection.ConnectionString = Core_Module.GetConstring
            DTAWorkshift.Connection.ConnectionString = Core_Module.GetConstring
            DTAWorkshiftItem.Connection.ConnectionString = Core_Module.GetConstring
            DTAEmployee.Connection.ConnectionString = Core_Module.GetConstring

            Tbl.Rows.Clear()

            If DTAEmployee.Fill(TblEmployee) > 0 Then
                DTAWorkshift.FillBy_StartAndEnd(TblWorkshift, EndDate, StartDate)
                DTAKaryawan.FillBy_Periode(TblAttendance, StartDate, EndDate)
                DTAWorkshiftItem.Fill(TblWorksfhitItem)

                DtaHoliday.Connection.ConnectionString = Core_Module.GetConstring
                DtaHoliday.Fill(TblHoliday, StartDate, EndDate)
                DtaLeave.Connection.ConnectionString = core_module.getconstring
                DtaLeave.Fill(TblLeave, StartDate, EndDate, 1)

                counter = 0

                For x As Integer = TblEmployee.Rows.Count - 1 To 0 Step -1
                    tanggalAcuan = StartDate
                    rowEmployee = TblEmployee.Rows(x)

                    TblWorkshift.DefaultView.RowFilter = "[primary entity] = " & rowEmployee.Primary
                    If TblWorkshift.DefaultView.Count = 0 Then
                        TblEmployee.Rows.Remove(rowEmployee)
                    Else
                        While tanggalAcuan <= EndDate
                            ActualInDone = False
                            ActualOutDone = False
                            ActualBreakinDone = False
                            ActualBreakOutDone = False

                            For i As Integer = 0 To TblWorkshift.DefaultView.Count - 1
                                rowworkshift = TblWorkshift.DefaultView(i).Row

                                'jika employee sudah tidak aktif maka cek tanggal resign nya
                                If rowEmployee.Aktif <> 0 Then
                                    'jika tanggal acuan lebih besar dari tanggal resign maka tidak ditampilkan
                                    If rowEmployee.Resign_Date < tanggalAcuan Then
                                        GoTo Skip2
                                    End If
                                End If

                                If tanggalAcuan >= rowworkshift.start_date And tanggalAcuan <= rowworkshift.end_date Then
                                    If rowworkshift.day_of_week = tanggalAcuan.DayOfWeek Then
                                        StatReason = "Workday"

                                        row = Tbl.NewtblLaporanClockInOutRow
                                        row.Primary_Karyawan = rowworkshift.primary_entity
                                        row.Tanggal = Format(tanggalAcuan, "dd MMM yyyy")
                                        row.Primary_Departement = rowworkshift.Primary_Dept
                                        row.Kode_Karyawan = rowworkshift.Code
                                        row.Nama_Karyawan = rowworkshift.Full_Name
                                        row.Dept_Code = rowworkshift.Dept_Code
                                        row.Dept_Desc = rowworkshift.Description
                                        row.Schedule_In = ObjCd.Convert_ToString(rowworkshift.schedule_in)
                                        row.Schedule_Out = ObjCd.Convert_ToString(rowworkshift.schedule_out)
                                        row.Schedule_break_out = ObjCd.Convert_ToString(rowworkshift.schedule_break_out)
                                        row.Schedule_break_in = ObjCd.Convert_ToString(rowworkshift.schedule_break_in)
                                        row.WorkShift_Code = rowworkshift.Workshift_Code
                                        row.WorkShift_Desc = rowworkshift.Workshift_Desc
                                        row.Primary_Workshift = rowworkshift.primary_workshift
                                        row.Actual_In = ObjCd.Convert_ToString(rowworkshift.schedule_in)
                                        row.Actual_Out = ObjCd.Convert_ToString(rowworkshift.schedule_out)
                                        row.Actual_Break_In = ObjCd.Convert_ToString(rowworkshift.schedule_break_in)
                                        row.Actual_Break_Out = ObjCd.Convert_ToString(rowworkshift.schedule_break_out)

                                        If rowworkshift.tipe_workdays = 0 Then
                                            TblHoliday.DefaultView.RowFilter = "[primary transaksi] = " & rowworkshift.primary_workshift & " AND [date] = #" & tanggalAcuan & "#"
                                            If TblHoliday.DefaultView.Count > 0 Then
                                                rowHoliday = TblHoliday.DefaultView(0).Row
                                                StatReason = "Holiday"
                                            Else
                                                TblLeave.DefaultView.RowFilter = "[primary karyawan] = " & rowEmployee.Primary & " AND [date] = #" & tanggalAcuan & "#"
                                                If TblLeave.DefaultView.Count > 0 Then
                                                    rowLeave = TblLeave.DefaultView(0).Row
                                                    StatReason = "Leave"
                                                End If
                                            End If
                                        Else
                                            StatReason = "Holiday"
                                        End If
                                        row.reason_status = StatReason

                                            'Actual in
                                            TblWorksfhitItem.DefaultView.RowFilter = "[primary transaksi] = " & rowworkshift.primary_workshift & " AND [schedule type] = 0"
                                            If TblWorksfhitItem.DefaultView.Count > 0 Then
                                                rowworkshiftItem = TblWorksfhitItem.DefaultView(0).Row
                                                If rowworkshiftItem.is_get_value = 1 And rowworkshiftItem._default = 1 Then
                                                    row.Actual_In = "-"
                                                    TblAttendance.DefaultView.RowFilter = "[Primary Karyawan] = " & row.Primary_Karyawan
                                                    If TblAttendance.DefaultView.Count > 0 Then
                                                        For ij As Integer = 0 To TblAttendance.DefaultView.Count - 1
                                                            rowAttendance = TblAttendance.DefaultView(ij).Row
                                                            If rowAttendance.transaction_time.Date = row.Tanggal.Date Then
                                                                If ActualInDone = False Then
                                                                    JamX = ObjCd.Convert_ToMinutes(rowAttendance.transaction_time.Hour & "." & rowAttendance.transaction_time.Minute)
                                                                    JamMasuk = rowworkshift.schedule_in
                                                                    If JamX >= JamMasuk - rowworkshiftItem.limit_before And JamX <= JamMasuk + rowworkshiftItem.limit_after Then
                                                                        If rowworkshiftItem.first_or_last = 0 Then
                                                                            ActualInDone = True
                                                                        Else
                                                                            ActualInDone = False
                                                                        End If
                                                                        row.Actual_In = ObjCd.Convert_ToString(JamX)
                                                                    End If
                                                                End If
                                                            End If
                                                        Next
                                                    End If
                                                End If
                                            End If

                                            'Actual Break Out
                                            TblWorksfhitItem.DefaultView.RowFilter = "[primary transaksi] = " & rowworkshift.primary_workshift & " AND [schedule type] = 1"
                                            If TblWorksfhitItem.DefaultView.Count > 0 Then
                                                rowworkshiftItem = TblWorksfhitItem.DefaultView(0).Row
                                                If rowworkshiftItem.is_get_value = 1 And rowworkshiftItem._default = 1 Then
                                                    row.Actual_Break_Out = "-"
                                                    TblAttendance.DefaultView.RowFilter = "[Primary Karyawan] = " & row.Primary_Karyawan
                                                    If TblAttendance.DefaultView.Count > 0 Then
                                                        For ij As Integer = 0 To TblAttendance.DefaultView.Count - 1
                                                            rowAttendance = TblAttendance.DefaultView(ij).Row
                                                            If rowAttendance.transaction_time.Date = row.Tanggal.Date Then
                                                                If ActualBreakOutDone = False Then
                                                                    JamX = ObjCd.Convert_ToMinutes(rowAttendance.transaction_time.Hour & "." & rowAttendance.transaction_time.Minute)
                                                                    JamKeluar = rowworkshift.schedule_break_out
                                                                    If JamX >= JamKeluar - rowworkshiftItem.limit_before And JamX <= JamKeluar + rowworkshiftItem.limit_after Then
                                                                        If rowworkshiftItem.first_or_last = 0 Then
                                                                            ActualBreakOutDone = True
                                                                        Else
                                                                            ActualBreakOutDone = False
                                                                        End If
                                                                        row.Actual_Break_Out = ObjCd.Convert_ToString(JamX)
                                                                    End If
                                                                End If
                                                            End If
                                                        Next
                                                    End If
                                                End If
                                            End If

                                            'Actual Break In
                                            TblWorksfhitItem.DefaultView.RowFilter = "[primary transaksi] = " & rowworkshift.primary_workshift & " AND [schedule type] = 2"
                                            If TblWorksfhitItem.DefaultView.Count > 0 Then
                                                rowworkshiftItem = TblWorksfhitItem.DefaultView(0).Row
                                                If rowworkshiftItem.is_get_value = 1 And rowworkshiftItem._default = 1 Then
                                                    row.Actual_Break_In = "-"
                                                    TblAttendance.DefaultView.RowFilter = "[Primary Karyawan] = " & row.Primary_Karyawan
                                                    If TblAttendance.DefaultView.Count > 0 Then
                                                        For ij As Integer = 0 To TblAttendance.DefaultView.Count - 1
                                                            rowAttendance = TblAttendance.DefaultView(ij).Row
                                                            If rowAttendance.transaction_time.Date = row.Tanggal.Date Then
                                                                If ActualBreakinDone = False Then
                                                                    JamX = ObjCd.Convert_ToMinutes(rowAttendance.transaction_time.Hour & "." & rowAttendance.transaction_time.Minute)
                                                                    JamMasuk = rowworkshift.schedule_break_in
                                                                    If JamX >= JamMasuk - rowworkshiftItem.limit_before And JamX <= JamMasuk + rowworkshiftItem.limit_after Then
                                                                        If rowworkshiftItem.first_or_last = 0 Then
                                                                            ActualBreakinDone = True
                                                                        Else
                                                                            ActualBreakinDone = False
                                                                        End If
                                                                        row.Actual_Break_In = ObjCd.Convert_ToString(JamX)
                                                                    End If
                                                                End If
                                                            End If
                                                        Next
                                                    End If
                                                End If
                                            End If

                                            'Actual out
                                            TblWorksfhitItem.DefaultView.RowFilter = "[primary transaksi] = " & rowworkshift.primary_workshift & " AND [schedule type] = 3"
                                            If TblWorksfhitItem.DefaultView.Count > 0 Then
                                                rowworkshiftItem = TblWorksfhitItem.DefaultView(0).Row
                                                If rowworkshiftItem.is_get_value = 1 And rowworkshiftItem._default = 1 Then
                                                    row.Actual_Out = "-"
                                                    TblAttendance.DefaultView.RowFilter = "[Primary Karyawan] = " & row.Primary_Karyawan
                                                    If TblAttendance.DefaultView.Count > 0 Then
                                                        For ij As Integer = 0 To TblAttendance.DefaultView.Count - 1
                                                            rowAttendance = TblAttendance.DefaultView(ij).Row
                                                            'jika sch in lebih besar dari sch out maka jadwal nya lintas hari
                                                            ActOutLintasHari = row.Tanggal.Date
                                                            If row.Schedule_In > row.Schedule_Out Then
                                                                ActOutLintasHari = ActOutLintasHari.AddDays(1)
                                                            End If

                                                            If rowAttendance.transaction_time.Date = ActOutLintasHari Then 'row.Tanggal.Date Then 
                                                                If ActualOutDone = False Then
                                                                    JamX = ObjCd.Convert_ToMinutes(rowAttendance.transaction_time.Hour & "." & rowAttendance.transaction_time.Minute)
                                                                    JamKeluar = rowworkshift.schedule_out
                                                                    If JamX >= JamKeluar - rowworkshiftItem.limit_before And JamX <= JamKeluar + rowworkshiftItem.limit_after Then
                                                                        If rowworkshiftItem.first_or_last = 0 Then
                                                                            ActualOutDone = True
                                                                        Else
                                                                            ActualOutDone = False
                                                                        End If
                                                                        row.Actual_Out = ObjCd.Convert_ToString(JamX)
                                                                    End If
                                                                End If
                                                            End If
                                                        Next
                                                    End If
                                                End If
                                            End If

                                            Tbl.Rows.Add(row)
                                            GoTo Skip2
                                    End If
                                End If
                            Next
Skip2:
                            tanggalAcuan = tanggalAcuan.AddDays(1)

                        End While
                    End If
                    counter += 1
                Next
                If ChkNoClockInOut.Checked = True Then
                    counter = 0
                    While counter <= Tbl.Rows.Count - 1
                        row = Tbl.Item(counter)
                        If row.Actual_In = "-" And row.Actual_Out = "-" And row.reason_status = "Workday" Then
                            counter += 1
                        Else
                            Tbl.Rows.Remove(row)
                        End If
                    End While
                End If

                Dim dtaAtt As New DseTrans_Attendance_Actual_In_ReportTableAdapters.tbl_attendanceTableAdapter
                Dim tblAtt As New DseTrans_Attendance_Actual_In_Report.tbl_attendanceDataTable
                Dim rowAtt As DseTrans_Attendance_Actual_In_Report.tbl_attendanceRow

                StartDate = Me.DtpDate.Value.Date + " " + "00:00:00" 'Format(Me.DtpDate.Value.Date, "MM/dd/yyyy") + " " + "00:00:00"
                EndDate = Me.DtpDateTo.Value.Date + " " + "23:59:59" 'Format(Me.DtpDateTo.Value.Date, "MM/dd/yyyy") + " " + "23:59:59"

                dtaAtt.Connection.ConnectionString = core_module.getconstring
                dtaAtt.Fill(tblAtt, StartDate, EndDate) 'sudah di urutkan by primary karyawan, date
                Tbl.DefaultView.RowFilter = ""
                Tbl.DefaultView.Sort = "[Primary Karyawan], [Tanggal]"
                For i As Integer = 0 To Tbl.DefaultView.Count - 1
                    row = Tbl.DefaultView(i).Row
                    With row
                        While tblAtt.Rows.Count > 0
                            rowAtt = tblAtt.Rows(0)
                            If .Primary_Karyawan < rowAtt.primary_karyawan Then
                                Exit While
                            ElseIf .Primary_Karyawan = rowAtt.primary_karyawan Then
                                If .Tanggal < rowAtt._date Then
                                    Exit While
                                ElseIf .Tanggal = rowAtt._date Then
                                    Dim IsReason As Boolean = True
                                    If rowAtt.Is_alfa > 0 Then
                                        If rowAtt.Is_alfa = 0.5 Then
                                            .reason_status &= " - Half Alfa"
                                        Else
                                            .reason_status &= " - Alfa"
                                        End If
                                    ElseIf rowAtt.is_permit > 0 Or rowAtt.Permit_duration > 0 Then
                                        If rowAtt.is_permit = 0.5 Then
                                            .reason_status &= " - Half Permit"
                                        Else
                                            .reason_status &= " - Permit"
                                        End If
                                    ElseIf rowAtt.is_sick > 0 Then
                                        If rowAtt.is_sick = 0.5 Then
                                            .reason_status &= " - Half Sick"
                                        Else
                                            .reason_status &= " - Sick"
                                        End If
                                    Else
                                        IsReason = False
                                    End If

                                    If rowAtt.is_late > 0 Then
                                        If IsReason = True Then
                                            .reason_status &= "/Late"
                                        Else
                                            .reason_status &= " - Late"
                                        End If
                                        IsReason = True
                                    End If
                                    If rowAtt.is_early_out > 0 Then
                                        If IsReason = True Then
                                            .reason_status &= "/Early Out"
                                        Else
                                            .reason_status &= " - Early Out"
                                        End If
                                        IsReason = True
                                    End If
                                    tblAtt.Rows.Remove(rowAtt)
                                ElseIf .Tanggal > rowAtt._date Then
                                    tblAtt.Rows.Remove(rowAtt)
                                End If
                            ElseIf .Primary_Karyawan > rowAtt.primary_karyawan Then
                                tblAtt.Rows.Remove(rowAtt)
                            End If
                        End While
                    End With
                Next

                TampilkandiDGV()

                If Me.CboGroupBy.SelectedIndex = 0 Then
                    report1.SetDataSource(DS)
                    report1.SetParameterValue(report1.Parameter_paramStart.ParameterFieldName, Format(StartDate, "dd MMM yyyy"))
                    report1.SetParameterValue(report1.Parameter_ParamTo.ParameterFieldName, Format(EndDate, "dd MMM yyyy"))
                    Me.CrvActulaIn.ReportSource = report1
                Else
                    report2.SetDataSource(DS)
                    report2.SetParameterValue(report2.Parameter_paramStart.ParameterFieldName, Format(StartDate, "dd MMM yyyy"))
                    report2.SetParameterValue(report2.Parameter_ParamTo.ParameterFieldName, Format(EndDate, "dd MMM yyyy"))
                    Me.CrvActulaIn.ReportSource = report2
                End If
                Me.CrvActulaIn.Show()
            End If







        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub
    Private Sub TampilkandiDGV()
        Dim Dv As New DataGridViewRow
        Try
            Tbl.DefaultView.RowFilter = Nothing
            If TabControl1.SelectedIndex = 0 Then
                Tbl.DefaultView.Sort = "[Dept Code] ASC, [Kode Karyawan] ASC, [Tanggal] ASC"
                If Me.CboGroupBy.SelectedIndex = 0 Then
                    Tbl.DefaultView.Sort = "[Dept Code] ASC, [Kode Karyawan] ASC, [Tanggal] ASC"
                Else
                    Tbl.DefaultView.Sort = "[WorkShift Code] ASC, [Dept Code] ASC,[Kode Karyawan] ASC,[Tanggal] ASC"
                End If
            Else
                If Me.ColDeptCode.Visible = True Then
                    Tbl.DefaultView.Sort = "[Dept Code] ASC, [Kode Karyawan] ASC, [Tanggal] ASC"
                Else
                    Tbl.DefaultView.Sort = "[Kode Karyawan] ASC, [Tanggal] ASC"
                End If
            End If

            Me.DgvItem.Rows.Clear()
            Me.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            For ij As Integer = 0 To Tbl.DefaultView.Count - 1
                row = Tbl.DefaultView(ij).Row
                Me.DgvItem.Rows.Add()
                Dv = DgvItem.Rows(Me.DgvItem.Rows.Count - 1)
                Dv.Cells(Me.ColDate.Index).Value = Format(row.Tanggal, "dd MMM yyyy")
                Dv.Cells(Me.ColDeptCode.Index).Value = row.Dept_Code
                Dv.Cells(Me.ColDeptDesc.Index).Value = row.Dept_Desc
                Dv.Cells(Me.ColEmpCode.Index).Value = row.Kode_Karyawan
                Dv.Cells(Me.ColEmpName.Index).Value = row.Nama_Karyawan
                Dv.Cells(Me.ColShift.Index).Value = row.WorkShift_Code
                Dv.Cells(Me.ColShiftDesc.Index).Value = row.WorkShift_Desc
                Dv.Cells(Me.ColStatus.Index).Value = row.reason_status
                Dv.Cells(Me.ColSchIn.Index).Value = row.Schedule_In
                Dv.Cells(Me.ColSchOut.Index).Value = row.Schedule_Out
                Dv.Cells(Me.ColSchBreakIn.Index).Value = row.Schedule_break_in
                Dv.Cells(Me.ColSchBreakOut.Index).Value = row.Schedule_break_out
                Dv.Cells(Me.ColActIn.Index).Value = row.Actual_In
                Dv.Cells(Me.ColActOut.Index).Value = row.Actual_Out
                Dv.Cells(Me.ColActBreakIn.Index).Value = row.Actual_Break_In
                Dv.Cells(Me.ColActBreakOut.Index).Value = row.Actual_Break_Out
            Next
            Me.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.DgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub



    Private Sub ShowLog()
        Dim DS As New DseTrans_Attendance_Actual_In_Report
        Dim Dta As New DseTrans_Attendance_Actual_In_ReportTableAdapters.DataTable1TableAdapter
        Dim TblDummy As New DseTrans_Attendance_Actual_In_Report.DataTable1DataTable
        Dim Tbl As DseTrans_Attendance_Actual_In_Report.DataTable1DataTable = DS.DataTable1
        Dim RowDummy, Row As DseTrans_Attendance_Actual_In_Report.DataTable1Row

        Dim DtaLocation As New DseTrans_Attendance_Actual_In_ReportTableAdapters.tblmaster_fingerprint_server_machinesTableAdapter
        Dim TblLocation As New DseTrans_Attendance_Actual_In_Report.tblmaster_fingerprint_server_machinesDataTable
        Dim RowLocation As DseTrans_Attendance_Actual_In_Report.tblmaster_fingerprint_server_machinesRow

        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            If Dta.FillBy_Date2(TblDummy, CDate(Format(Me.DtpDate.Value, "dd MMM yyyy " & Format(Me.DtpTimeFrom.Value, "HH:mm") & ":00")), CDate(Format(Me.DtpDate.Value, "dd MMM yyyy " & Format(Me.DtpTimeTo.Value, "HH:mm") & ":00"))) > 0 Then
                ''set location column
                DtaLocation.Connection.ConnectionString = Core_Module.GetConstring
                DtaLocation.Fill(TblLocation)
                For Each RowDummy In TblDummy
                    TblLocation.DefaultView.RowFilter = "[primary server] = " & RowDummy.primary_attendance_server.ToString & " AND [terminal id] = '" & RowDummy.terminal_id & "'"
                    If TblLocation.DefaultView.Count > 0 Then
                        RowLocation = TblLocation.DefaultView(0).Row
                        RowDummy.location = RowLocation.remarks
                    End If
                Next

                ''filter by location and first or last
                For Each items_ As Object In Me.ClbLocation.CheckedItems
                    TblDummy.DefaultView.RowFilter = "[location] = '" & items_.ToString & "'"
                    If Me.CboIfMultiple.SelectedIndex = 0 Then
                        TblDummy.DefaultView.Sort = "[transaction time] ASC"
                    Else
                        TblDummy.DefaultView.Sort = "[transaction time] DESC"
                    End If
                    For Each DV As DataRowView In TblDummy.DefaultView
                        RowDummy = DV.Row
                        Tbl.DefaultView.RowFilter = "[location] = '" & RowDummy.location & "' AND [primary employee] = " & RowDummy.primary_employee.ToString
                        If Tbl.DefaultView.Count = 0 Then
                            Row = Tbl.NewDataTable1Row
                            For i As Integer = 0 To TblDummy.Columns.Count - 1
                                Row.Item(i) = RowDummy.Item(i)
                            Next
                            Tbl.Rows.Add(Row)
                        End If
                    Next
                Next
                If Tbl.Count > 0 Then
                    Me.Report.SetDataSource(DS)
                    Me.Report.SetParameterValue(Me.Report.Parameter_P_Date.ParameterFieldName, Me.DtpDate.Value)
                    Me.CrvActulaIn.ReportSource = Me.Report
                Else
                    Me.CrvActulaIn.ReportSource = Nothing
                    MsgBox("Nobody clock in/out at selected location.", MsgBoxStyle.Information, "Information")
                End If
            Else
                Me.CrvActulaIn.ReportSource = Nothing
                If Me.DtpDate.Value.Date = Now.Date Then
                    MsgBox("Nobody clock in/out today.", MsgBoxStyle.Information, "Information")
                Else
                    MsgBox("Nobody clock in/out on " & Format(Me.DtpDate.Value.Date, "dd MMM yyyy") & ".", MsgBoxStyle.Information, "Information")
                End If
            End If
            Me.CrvActulaIn.Zoom(90)
            Me.CrvActulaIn.Refresh()
            Me.CrvActulaIn.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub CboFormat_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboFormat.SelectedIndexChanged
        Try
            If Me.CboFormat.SelectedIndex = 0 Then
                Me.LblDate.Text = "Date From"
                Me.LblDateTo.Text = "Date To"
                Me.DtpDateTo.Visible = True
                Me.DtpTimeFrom.Visible = False
                Me.Label4.Visible = False 'strip
                Me.DtpTimeTo.Visible = False
                Me.LblGroupOrMultipleClock.Text = "Group By"
                Me.CboIfMultiple.Enabled = False
                Me.CboIfMultiple.Visible = False
                Me.CboGroupBy.Visible = True
                Me.CboGroupBy.Enabled = True
                Me.ChkNoClockInOut.Enabled = True
                Me.LlbColSetting.Enabled = True
                Me.ClbLocation.Enabled = False
            Else
                Me.LblDate.Text = "Date"
                Me.LblDateTo.Text = "Time"
                Me.DtpDateTo.Visible = False
                Me.DtpTimeFrom.Visible = True
                Me.Label4.Visible = True 'strip
                Me.DtpTimeTo.Visible = True
                Me.LblGroupOrMultipleClock.Text = "If multiple clock in/out"
                Me.CboIfMultiple.Enabled = True
                Me.CboIfMultiple.Visible = True
                Me.CboGroupBy.Visible = False
                Me.CboGroupBy.Enabled = False
                Me.ChkNoClockInOut.Enabled = False
                Me.LlbColSetting.Enabled = False
                Me.TabControl1.SelectedIndex = 0
                Me.ClbLocation.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SettingColumn()
        Try
            'OT
            If Me.ObjSettingsUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_ClockInOut_ShowSchedule, Me.UserCode) = "1" Then
                Me.ColSchIn.Visible = True
                Me.ColSchOut.Visible = True

                If Me.ObjSettingsUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_ClockInOut_ShowScheduleBreak, Me.UserCode) = "1" Then
                    Me.ColSchBreakOut.Visible = True
                    Me.ColSchBreakIn.Visible = True
                Else
                    Me.ColSchBreakOut.Visible = False
                    Me.ColSchBreakIn.Visible = False
                End If
            Else
                Me.ColSchIn.Visible = False
                Me.ColSchOut.Visible = False
                Me.ColSchBreakOut.Visible = False
                Me.ColSchBreakIn.Visible = False
            End If
            If Me.ObjSettingsUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_ClockInOut_ShowActual, Me.UserCode) = "1" Then
                Me.ColActIn.Visible = True
                Me.ColActOut.Visible = True

                If Me.ObjSettingsUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_ClockInOut_ShowActualBreak, Me.UserCode) = "1" Then
                    Me.ColActBreakIn.Visible = True
                    Me.ColActBreakOut.Visible = True
                Else
                    Me.ColActBreakIn.Visible = False
                    Me.ColActBreakOut.Visible = False
                End If
            Else
                Me.ColActIn.Visible = False
                Me.ColActBreakOut.Visible = False
                Me.ColActBreakIn.Visible = False
                Me.ColActOut.Visible = False
            End If

            If Me.ObjSettingsUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_ClockInOut_ShowShift, Me.UserCode) = "1" Then
                Me.ColShift.Visible = True
                Me.ColShiftDesc.Visible = True
            Else
                Me.ColShift.Visible = False
                Me.ColShiftDesc.Visible = False
            End If

            If Me.ObjSettingsUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.HR_Attendance, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.HR_ClockInOut_ShowDept, Me.UserCode) = "1" Then
                Me.ColDeptCode.Visible = True
                Me.ColDeptDesc.Visible = True
            Else
                Me.ColDeptCode.Visible = False
                Me.ColDeptDesc.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LlbColSetting_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbColSetting.LinkClicked
        Dim ij As New FrmTrans_Attendance_Actual_In_Report_ColSetting

        Try
            ij.UserCode = Me.UserCode
            ij.ShowDialog()
            Me.SettingColumn()
            Me.TampilkandiDGV()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        Try
            If Me.TabControl1.SelectedIndex = 0 Then
                Me.CrvActulaIn.DisplayStatusBar = True
                Me.CrvActulaIn.DisplayToolbar = True
            Else
                If Me.CboFormat.SelectedIndex = 1 Then
                    MsgBox("Format 2 only for format period", MsgBoxStyle.Information, "Information")
                    Me.TabControl1.SelectedIndex = 0
                End If
                Me.SettingColumn()
                Me.TampilkandiDGV()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class