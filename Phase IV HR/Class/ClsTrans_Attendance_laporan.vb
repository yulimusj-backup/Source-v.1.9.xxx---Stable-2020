Option Explicit On
Option Strict Off
Imports System.Windows.Forms, Core_Module

Public Class ClsTrans_Attendance_laporan

    Dim obj As New ClsCekData
    Dim Objcd As New Core_Module.ClsCekData
    Dim report As New CrpLaporanAtt

    Private Enum enumOT As Integer
        noOt
        weekdayOt
        holidayOt
    End Enum

    Public Enum EnumOvertime As Integer
        No_Overtime
        Weekdays_Overtime
        Weekends_Overtime
    End Enum

    Public Enum EnumStatus As Integer
        Present
        Leave
        Half_day_Leave
        Sick
        Half_day_Sick
        Permit
        Half_day_Permit
        Absent
        Half_Absent
        Holiday
        Half_Holiday
    End Enum

    Public Sub cetak_detail(ByVal vStrDate As Date, ByVal vEndDate As Date, _
    ByVal vCrv As CrystalDecisions.Windows.Forms.CrystalReportViewer)
        Dim ReportFormatA4 As New CrpLaporanAtt_Detail
        Dim DS As New DseTrans_Attendance_Cetak_IGI
        Dim DTA As New DseTrans_Attendance_Cetak_IGITableAdapters.DataTable1TableAdapter
        Dim TBL As DseTrans_Attendance_Cetak_IGI.DataTable1DataTable = DS.DataTable1
        Dim Row, RowTotal As DseTrans_Attendance_Cetak_IGI.DataTable1Row
        'Parameter Total 
        Dim Total_Weekday_OT_1 As Integer = 0
        Dim Total_Weekday_OT_2 As Integer = 0
        Dim Total_Weekday_OT_3 As Integer = 0
        Dim Total_Weekend_OT_1 As Integer = 0
        Dim Total_Weekend_OT_2 As Integer = 0
        Dim Total_Weekend_OT_3 As Integer = 0
        Dim Total_Sick As Decimal = 0
        Dim Total_Leave As Decimal = 0
        Dim Total_Permit As Decimal = 0
        Dim Total_Late As Decimal = 0
        Dim Total_Outstation As Decimal = 0
        Dim Sub_Total_Actual As Decimal = 0
        Dim Total_Permit_duration As Integer = 0

        Dim RowSetting As DseTrans_Attendance_Cetak_IGI.TblSettingRow
        Dim PrimaryAttMax As Double = -1
        Dim PrimaryAttMin As Double = -1
        Dim PrimaryEmp_Buffer As Double = -1
        Dim ListPrimaryAtt_Delete As New List(Of Double)
        Dim LastPrimary As Double = -1
        Dim ApaHitungTotal As Boolean = False

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'Fill data
            DTA.FillBy_DateFromTo(TBL, vStrDate, vEndDate)

            'Koding Laporan
            If TBL.Rows.Count > 0 Then
                For i As Integer = 0 To TBL.Rows.Count - 1
                    Row = TBL.Rows(i)
                    ApaHitungTotal = True
                    With Row
                        .start_date = vStrDate
                        .end_date = vEndDate
                        .sch_in_format = Me.Objcd.Convert_ToString(.sch_in)
                        .sch_out_format = Me.Objcd.Convert_ToString(.sch_out)
                        If .presence_status <> EnumStatus.Sick And .presence_status <> EnumStatus.Leave And _
                        .presence_status <> EnumStatus.Permit And .presence_status <> EnumStatus.Holiday And _
                        .presence_status <> EnumStatus.Absent Then
                            .act_in_format = Me.Objcd.Convert_ToString(.actual_in)
                            .act_out_format = Me.Objcd.Convert_ToString(.actual_out)
                        End If
                        If .actual_work_duration > 0 Then
                            .total_actual_work_duration_format = Me.Objcd.Convert_ToString(.actual_work_duration)
                        End If

                        ._date = ._date

                        'mengambil data overtime 
                        If .OT_Duration > 0 Then
                            If .tipe_overtime = EnumOvertime.Weekdays_Overtime Then
                                If .OT_Duration > 540 Then
                                    .OT_Weekday_1 = 60
                                    .OT_Weekday_2 = 480
                                    .OT_Weekday_3 = .OT_Duration - 540
                                ElseIf .OT_Duration > 60 And .OT_Duration <= 540 Then
                                    .OT_Weekday_1 = 60
                                    .OT_Weekday_2 = .OT_Duration - 60
                                ElseIf .OT_Duration <= 60 Then
                                    .OT_Weekday_1 = .OT_Duration
                                End If

                                If .OT_Weekday_1 > 0 Then
                                    .OT_Weekday_1_Format = Me.Objcd.Convert_ToString(.OT_Weekday_1)
                                End If
                                If .OT_Weekday_2 > 0 Then
                                    .OT_Weekday_2_Format = Me.Objcd.Convert_ToString(.OT_Weekday_2)
                                End If
                                If .OT_Weekday_3 > 0 Then
                                    .OT_Weekday_3_Format = Me.Objcd.Convert_ToString(.OT_Weekday_3)
                                End If
                            ElseIf .tipe_overtime = EnumOvertime.Weekends_Overtime Then
                                If .OT_Duration > 540 Then
                                    .OT_Weekend_1 = 480
                                    .OT_Weekend_2 = 60
                                    .OT_Weekend_3 = .OT_Duration - 540
                                ElseIf .OT_Duration > 480 And .OT_Duration <= 540 Then
                                    .OT_Weekend_1 = 480
                                    .OT_Weekend_2 = .OT_Duration - 480
                                ElseIf .OT_Duration <= 480 Then
                                    .OT_Weekend_1 = .OT_Duration
                                End If

                                If .OT_Weekend_1 > 0 Then
                                    .OT_Weekday_1_Format = Me.Objcd.Convert_ToString(.OT_Weekend_1)
                                End If
                                If .OT_Weekend_2 > 0 Then
                                    .OT_Weekday_2_Format = Me.Objcd.Convert_ToString(.OT_Weekend_2)
                                End If
                                If Row.OT_Weekend_3 > 0 Then
                                    .OT_Weekday_3_Format = Me.Objcd.Convert_ToString(.OT_Weekend_3)
                                End If
                            End If
                        End If

                        'Leave, Late, Sick
                        .late_duration_format = Me.Objcd.Convert_ToString(.late_duration)
                        If .presence_status = EnumStatus.Present Then
                            If .late_duration > 0 Then
                                .Late_Leave_Sick = "Late: " + .late_duration_format.ToString
                                If .Permit_duration > 0 Then
                                    If String.IsNullOrEmpty(.Late_Leave_Sick) = False Then
                                        .Late_Leave_Sick += ", permit: " + Me.Objcd.Convert_ToString(.Permit_duration).ToString
                                    End If
                                End If
                            Else
                                If .Permit_duration > 0 Then
                                    .Late_Leave_Sick = "Permit: " + Me.Objcd.Convert_ToString(.Permit_duration).ToString
                                End If
                            End If
                        End If
                        If .presence_status = EnumStatus.Sick Then
                            .Late_Leave_Sick = "Sick"
                        ElseIf .presence_status = EnumStatus.Half_day_Sick Then
                            .Late_Leave_Sick = "Half day sick"
                        End If
                        If .presence_status = EnumStatus.Leave Then
                            If String.IsNullOrEmpty(.cuti_description) = False Then
                                .Late_Leave_Sick = "Leave: " + .cuti_description.ToString
                            End If
                        ElseIf .presence_status = EnumStatus.Half_day_Leave Then
                            If String.IsNullOrEmpty(.cuti_description) = False Then
                                .Late_Leave_Sick = "Half day leave: " + .cuti_description.ToString
                            End If
                        End If
                        If .presence_status = EnumStatus.Permit Then
                            .Late_Leave_Sick = "Permit"
                        ElseIf .presence_status = EnumStatus.Half_day_Permit Then
                            .Late_Leave_Sick = "Half day permit"
                        ElseIf .presence_status = EnumStatus.Half_Absent Then
                            .Late_Leave_Sick = "Half day absent"
                        ElseIf .presence_status = EnumStatus.Absent Then
                            .Late_Leave_Sick = "Absent"
                        End If
                        If String.IsNullOrEmpty(.Late_Leave_Sick) = False Then
                            If String.IsNullOrEmpty(.remarks) = False Then
                                .Late_Leave_Sick += "-" + .remarks.ToString
                            End If
                        Else
                            If String.IsNullOrEmpty(.remarks) = False Then
                                .Late_Leave_Sick = .remarks.ToString
                            End If
                        End If

                        If PrimaryEmp_Buffer <> .primary_karyawan And PrimaryEmp_Buffer <> -1 Then
HitungTotal:
                            TBL.DefaultView.RowFilter = "[primary karyawan] = " & PrimaryEmp_Buffer
                            If TBL.DefaultView.Count > 0 Then
                                RowTotal = TBL.DefaultView(0).Row
                                RowTotal.Total_Weekdays_OT_1 = Me.Objcd.Convert_ToString(Total_Weekday_OT_1).ToString
                                RowTotal.Total_Weekdays_OT_2 = Me.Objcd.Convert_ToString(Total_Weekday_OT_2).ToString
                                RowTotal.Total_Weekdays_OT_3 = Me.Objcd.Convert_ToString(Total_Weekday_OT_3).ToString
                                RowTotal.Total_Weekend_OT_1 = Me.Objcd.Convert_ToString(Total_Weekend_OT_1).ToString
                                RowTotal.Total_Weekend_OT_2 = Me.Objcd.Convert_ToString(Total_Weekend_OT_2).ToString
                                RowTotal.Total_Weekend_OT_3 = Me.Objcd.Convert_ToString(Total_Weekend_OT_3).ToString
                                RowTotal.Total_Permit = Total_Permit.ToString
                                RowTotal.Total_Sick = Total_Sick.ToString
                                RowTotal.Total_Leave = Total_Leave.ToString
                                RowTotal.Total_Late = Total_Late.ToString
                                RowTotal.Total_Outstation = Total_Outstation.ToString
                                RowTotal.Total_Duration = Me.Objcd.Convert_ToString(Total_Permit_duration).ToString
                            End If

                            Total_Weekday_OT_1 = 0
                            Total_Weekday_OT_2 = 0
                            Total_Weekday_OT_3 = 0
                            Total_Weekend_OT_1 = 0
                            Total_Weekend_OT_2 = 0
                            Total_Weekend_OT_3 = 0
                            Total_Permit = 0
                            Total_Sick = 0
                            Total_Leave = 0
                            Total_Late = 0
                            Total_Outstation = 0
                            Sub_Total_Actual = 0
                            Total_Permit_duration = 0

                            If i = TBL.Count - 1 Then
                                Exit For
                            End If
                            ApaHitungTotal = False
                        End If

                        'hitung total
                        Total_Weekday_OT_1 += .OT_Weekday_1
                        Total_Weekday_OT_2 += .OT_Weekday_2
                        Total_Weekday_OT_3 += .OT_Weekday_3
                        Total_Weekend_OT_1 += .OT_Weekend_1
                        Total_Weekend_OT_2 += .OT_Weekend_2
                        Total_Weekend_OT_3 += .OT_Weekend_3
                        Total_Permit += .is_permit
                        Total_Sick += .is_sick
                        Total_Leave += .is_cuti
                        Total_Late += .is_late
                        Total_Outstation += .is_outstation
                        Total_Permit_duration += .Permit_duration
                        Sub_Total_Actual += .actual_work_duration
                        .Grand_Total = Me.Objcd.Convert_ToString(Sub_Total_Actual).ToString
                        If .Grand_Total = "00.00" Then
                            .Grand_Total = "-"
                        End If
                        .Total_OT_1 = Me.Objcd.Convert_ToString(Total_Weekday_OT_1 + Total_Weekend_OT_1).ToString
                        If .Total_OT_1 = "00.00" Then
                            .Total_OT_1 = "-"
                        End If
                        .Total_OT_2 = Me.Objcd.Convert_ToString(Total_Weekday_OT_2 + Total_Weekend_OT_2).ToString
                        If .Total_OT_2 = "00.00" Then
                            .Total_OT_2 = "-"
                        End If
                        .Total_OT_3 = Me.Objcd.Convert_ToString(Total_Weekday_OT_3 + Total_Weekend_OT_3).ToString
                        If .Total_OT_3 = "00.00" Then
                            .Total_OT_3 = "-"
                        End If

                        PrimaryEmp_Buffer = .primary_karyawan
                        If ApaHitungTotal = True And i = TBL.Count - 1 Then
                            GoTo HitungTotal
                        End If
                    End With
                Next

                If TBL.Rows.Count > 0 Then
                    Row = TBL.Rows(TBL.Count - 1)
                    LastPrimary = Row.primary_transaksi
                End If

                'Setting Laporan
                ReportFormatA4.SetDataSource(DS)
                ReportFormatA4.SetParameterValue(ReportFormatA4.Parameter_LastPrimary.ParameterFieldName, LastPrimary)
                vCrv.ReportSource = ReportFormatA4
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub cetak(ByVal vStrDate As Date, ByVal vEndDate As Date, _
            ByVal vCrv As CrystalDecisions.Windows.Forms.CrystalReportViewer)
        Dim ds As New DseLaporanAttendance
        Dim tblCetak As DseLaporanAttendance.tblCetakDataTable = ds.tblCetak
        Dim rowCetak As DseLaporanAttendance.tblCetakRow
        'Dim rowCetakBefore As DseLaporanAttendance.TblCetakRow
        Dim dtaDb As New DseLaporanAttendanceTableAdapters.tblAttendanceTableAdapter
        Dim tblDb As New DseLaporanAttendance.tblAttendanceDataTable
        Dim rowDb As DseLaporanAttendance.tblAttendanceRow
        Dim Filter As String = ""
        Dim isSick As Double = 0
        Dim isPermit As Double = 0
        Dim isAbsen As Double = 0
        Dim isPresent As Double = 0
        Dim isLate As Double = 0
        Dim isLateQty As Integer = 0
        Dim isLeaveBiasa As Double = 0
        Dim isLeaveSpesial As Double = 0
        Dim otWeekday1 As Double = 0
        Dim otWeekday2 As Double = 0
        Dim otWeekday9 As Double = 0
        Dim otHoliday1 As Double = 0
        Dim otHoliday8 As Double = 0
        Dim otHoliday9 As Double = 0
        Dim sisaOt As Double = 0

        Dim PrimaryDeptSebelumnya As Double = -1
        Dim TotalPermit As Double = 0
        Dim TotalLate As Double = 0
        Dim TotalLateQty As Integer = 0
        Dim TotalOTW1 As Double = 0
        Dim TotalOTW2 As Double = 0
        Dim TotalOTW9 As Double = 0
        Dim TotalOTH1 As Double = 0
        Dim TotalOTH8 As Double = 0
        Dim TotalOTH9 As Double = 0
        Try
            'cari berdasarkan tanggal dari sampai
            dtaDb.Fill(tblDb, vStrDate, vEndDate)
            For Each rowDb In tblDb.Rows
                'reset semua nilai
                Filter = ""
                isSick = 0
                isPermit = 0
                isAbsen = 0
                isPresent = 0
                isLate = 0
                isLeaveBiasa = 0
                isLeaveSpesial = 0
                otWeekday1 = 0
                otWeekday2 = 0
                otWeekday9 = 0
                otHoliday1 = 0
                otHoliday8 = 0
                otHoliday9 = 0

                'Cari nilai IsSick, IsPermit, IsAbsent, IsPresent, IsLeave
                '=IsSick
                '===Ambil apa adanya dari IsSick database, jika setengah hari set nilai IsPresent 0,5
                isSick += rowDb.is_sick
                If isSick = 0.5 Then
                    isPresent += 0.5
                End If

                '=IsPermit
                '===Cek dulu jika IsPermit db > 0,  langsung kalikan dengan sch workday
                '===jika IsPermit db = 0,5 jika setengah hari set nilai var IsPresent 0,5
                '===jika IsPresent db > 0 cek apakah ada permit durasi , kalo ada masukan ke var IsPermit
                If rowDb.is_permit > 0 Then
                    If rowDb.is_permit = 0.5 Then
                        isPresent += 0.5
                    End If
                    If rowDb.sch_work_duration <= 0 Then
                        rowDb.sch_work_duration = 480
                    End If
                    isPermit += rowDb.is_permit * rowDb.sch_work_duration
                End If
                If rowDb.is_present > 0 Then
                    If rowDb.Permit_duration > 0 Then
                        isPermit += rowDb.Permit_duration
                    End If
                End If

                '=IsAbsen
                '===Ambil apa adanya dari IsAbsen database, jika setengah hari set nilai IsPresent 0,5
                isAbsen += rowDb.Is_alfa
                If rowDb.Is_alfa = 0.5 Then
                    isPresent += 0.5
                End If

                '=IsLate
                '===Ambil dari late durasi
                isLate += rowDb.late_duration

                '=IsPresent
                '===Ambil dari db
                isPresent += rowDb.is_present

                ''Tambahan Aul 120915
                If rowDb.late_duration > Me.obj.Convert_ToMinutes("00.00") Then
                    isLateQty = 1
                Else
                    isLateQty = 0
                End If
                ''

                '=IsLeave
                '===Ambil dari setting (biasa/spesial)
                '===Jika Leave DB = 0.5 tambahkan isPresent dengan 0.5
                If rowDb.is_leave > 0 Then
                    If rowDb.Status = 0 Then
                        isLeaveBiasa += rowDb.is_leave
                    Else
                        isLeaveSpesial += rowDb.is_leave
                    End If
                End If
                If rowDb.is_leave = 0.5 Then
                    isPresent += 0.5
                End If


                '=IsOvertime ( ambil durasinya )
                '=== ambil tipe ot [no/weekday/holiday]
                '=== jika no, jangan dihitung
                '=== jika weekday, dapatkan jml_ot
                '=== periksa >= 1, update kolom `1` | sisanya simpan di `> 1`
                '=== jika holiday, dapatkan jml_ot
                '=== periksa <= 8, jika ya simpan total_ot , lainnya '0' kan
                '=== jika >8 and <=9, isi `1-8` dengan '8', kolom `9` isi dengan [sisa_ot]
                '=== jika >9, simpan ....., isi kolom `>9` dengan [sisa_ot]
                If rowDb.tipe_overtime = ClsTrans_Attendance_laporan.enumOT.weekdayOt Then
                    If rowDb.jml <= Me.obj.Convert_ToMinutes("01.00") Then
                        otWeekday1 = rowDb.jml
                        ''Tambahan Aul 120915
                    ElseIf rowDb.jml > Me.obj.Convert_ToMinutes("01.00") And rowDb.jml <= Me.obj.Convert_ToMinutes("09.00") Then
                        otWeekday1 = Me.obj.Convert_ToMinutes("01.00")
                        otWeekday2 = rowDb.jml - Me.obj.Convert_ToMinutes("01.00")
                    ElseIf rowDb.jml > Me.obj.Convert_ToMinutes("09.00") Then
                        otWeekday1 = Me.obj.Convert_ToMinutes("01.00")
                        otWeekday2 = Me.obj.Convert_ToMinutes("09.00") - otWeekday1
                        otWeekday9 = rowDb.jml - Me.obj.Convert_ToMinutes("09.00")
                        ''
                    End If
                ElseIf rowDb.tipe_overtime = ClsTrans_Attendance_laporan.enumOT.holidayOt Then
                    If rowDb.jml <= Me.obj.Convert_ToMinutes("08.00") Then
                        otHoliday1 = rowDb.jml
                    ElseIf rowDb.jml > Me.obj.Convert_ToMinutes("08.00") And rowDb.jml <= Me.obj.Convert_ToMinutes("09.00") Then
                        otHoliday1 = Me.obj.Convert_ToMinutes("08.00")
                        otHoliday8 = rowDb.jml - Me.obj.Convert_ToMinutes("08.00")
                    ElseIf rowDb.jml > Me.obj.Convert_ToMinutes("09.00") Then
                        otHoliday1 = Me.obj.Convert_ToMinutes("08.00")
                        otHoliday8 = Me.obj.Convert_ToMinutes("01.00")
                        otHoliday9 = rowDb.jml - Me.obj.Convert_ToMinutes("09.00")
                    End If
                End If

                '=== Jika IsOvertime > 0
                '=== Weekday = 1 jam pertama dan selanjutnya
                '=== holiday = 8 jam pertama, jam ke-9, lebih dari 9 jam
                '=== tinggal cek type ot'y trus ambil durasinya

                'filter tbl cetak manual berdasarkan primary karyawan
                'jika tidak ada tambah baris baru
                'jika ada jumlah kan masing" IsSick, IsPermit, IsAbsent, IsPresent, IsLeave
                Filter = "[Primary Emp] = " & rowDb.Primary_Emp.ToString
                tblCetak.DefaultView.RowFilter = Filter
                If tblCetak.DefaultView.Count = 0 Then
                    rowCetak = tblCetak.NewtblCetakRow
                    With rowCetak
                        .Dept_Code = rowDb.Dept_Code
                        .Name = rowDb.Full_Name
                        .Job_Code = rowDb.Job_Code
                        .Is_Sick = isSick
                        .Is_Permit = Me.obj.Convert_ToString(isPermit)
                        If .Is_Permit = "00.00" Then
                            .Is_Permit = "-"
                        End If
                        .Is_Absen = isAbsen
                        .Is_Present = isPresent
                        .Is_Late = Me.obj.Convert_ToString(isLate)
                        If .Is_Late = "00.00" Then
                            .Is_Late = "-"
                        End If
                        .Is_Late_Qty += isLateQty
                        .Is_Leave_Biasa = isLeaveBiasa
                        .Is_Leave_Spesial = isLeaveSpesial
                        .Primary_Emp = rowDb.Primary_Emp
                        .Ot_Weekday_1 = Me.obj.Convert_ToString(otWeekday1)
                        If .Ot_Weekday_1 = "00.00" Then
                            .Ot_Weekday_1 = "-"
                        End If
                        .Ot_Weekday_2 = Me.obj.Convert_ToString(otWeekday2)
                        If .Ot_Weekday_2 = "00.00" Then
                            .Ot_Weekday_2 = "-"
                        End If
                        ''Tambahan Aul 120915
                        .Ot_Weekday_9 = Me.obj.Convert_ToString(otWeekday9)
                        If .Ot_Weekday_9 = "00.00" Then
                            .Ot_Weekday_9 = "-"
                        End If
                        ''
                        .Ot_Holiday_1 = Me.obj.Convert_ToString(otHoliday1)
                        If .Ot_Holiday_1 = "00.00" Then
                            .Ot_Holiday_1 = "-"
                        End If
                        .Ot_Holiday_8 = Me.obj.Convert_ToString(otHoliday8)
                        If .Ot_Holiday_8 = "00.00" Then
                            .Ot_Holiday_8 = "-"
                        End If
                        .Ot_Holiday_9 = Me.obj.Convert_ToString(otHoliday9)
                        If .Ot_Holiday_9 = "00.00" Then
                            .Ot_Holiday_9 = "-"
                        End If
                        If PrimaryDeptSebelumnya = rowDb.Primary_Dept Then
                            TotalPermit += isPermit
                            TotalLate += isLate
                            TotalLateQty += isLateQty
                            TotalOTW1 += otWeekday1
                            TotalOTW2 += otWeekday2
                            TotalOTW9 += otWeekday9
                            TotalOTH1 += otHoliday1
                            TotalOTH8 += otHoliday8
                            TotalOTH9 += otHoliday9
                        Else
                            TotalPermit = isPermit
                            TotalLate = isLate
                            TotalLateQty = isLateQty
                            TotalOTW1 = otWeekday1
                            TotalOTW2 = otWeekday2
                            TotalOTW9 = otWeekday9
                            TotalOTH1 = otHoliday1
                            TotalOTH8 = otHoliday8
                            TotalOTH9 = otHoliday9
                        End If
                        .Total_Permit = Me.obj.Convert_ToString(TotalPermit)
                        If .Total_Permit = "00.00" Then
                            .Total_Permit = "-"
                        End If
                        .Total_Telat = Me.obj.Convert_ToString(TotalLate)
                        If .Total_Telat = "00.00" Then
                            .Total_Telat = "-"
                        End If
                        .Total_Telat_Qty = TotalLateQty
                        .Total_OT_Weekday_1 = Me.obj.Convert_ToString(TotalOTW1)
                        If .Total_OT_Weekday_1 = "00.00" Then
                            .Ot_Weekday_1 = "-"
                        End If
                        .Total_OT_Weekday_1 = Me.obj.Convert_ToString(TotalOTW1)
                        If .Total_OT_Weekday_1 = "00.00" Then
                            .Ot_Weekday_1 = "-"
                        End If
                        .Total_OT_Weekday_2 = Me.obj.Convert_ToString(TotalOTW2)
                        If .Ot_Weekday_2 = "00.00" Then
                            .Ot_Weekday_2 = "-"
                        End If
                        ''Tambahan Aul 120915
                        .Total_OT_Weekday_9 = Me.obj.Convert_ToString(TotalOTW9)
                        If .Total_OT_Weekday_9 = "00.00" Then
                            .Total_OT_Weekday_9 = "-"
                        End If
                        ''
                        .Total_OT_Holiday_1 = Me.obj.Convert_ToString(TotalOTH1)
                        If .Total_OT_Holiday_1 = "00.00" Then
                            .Ot_Holiday_1 = "-"
                        End If
                        .Total_OT_Holiday_8 = Me.obj.Convert_ToString(TotalOTH8)
                        If .Ot_Holiday_8 = "00.00" Then
                            .Ot_Holiday_8 = "-"
                        End If
                        .Total_OT_Holiday_9 = Me.obj.Convert_ToString(TotalOTH9)
                        If .Ot_Holiday_9 = "00.00" Then
                            .Total_OT_Holiday_9 = "-"
                        End If
                    End With
                    tblCetak.AddtblCetakRow(rowCetak)
                Else
                    rowCetak = tblCetak.DefaultView(0).Row
                    With rowCetak
                        .Is_Sick += isSick
                        If .Is_Permit = "-" Then
                            .Is_Permit = "00.00"
                        End If
                        .Is_Permit = Me.obj.Convert_ToString(Me.obj.Convert_ToMinutes(.Is_Permit) + isPermit)
                        If .Is_Permit = "00.00" Then
                            .Is_Permit = "-"
                        End If
                        .Is_Absen += isAbsen
                        .Is_Present += isPresent
                        If .Is_Late = "-" Then
                            .Is_Late = "00.00"
                        End If
                        .Is_Late = Me.obj.Convert_ToString(Me.obj.Convert_ToMinutes(.Is_Late) + isLate)
                        If .Is_Late = "00.00" Then
                            .Is_Late = "-"
                        End If
                        .Is_Late_Qty += isLateQty
                        .Is_Leave_Biasa += isLeaveBiasa
                        .Is_Leave_Spesial += isLeaveSpesial
                        If .Ot_Weekday_1 = "-" Then
                            .Ot_Weekday_1 = "00.00"
                        End If
                        .Ot_Weekday_1 = Me.obj.Convert_ToString(Me.obj.Convert_ToMinutes(.Ot_Weekday_1) + otWeekday1)
                        If .Ot_Weekday_1 = "00.00" Then
                            .Ot_Weekday_1 = "-"
                        End If
                        If .Ot_Weekday_2 = "-" Then
                            .Ot_Weekday_2 = "00.00"
                        End If
                        .Ot_Weekday_2 = Me.obj.Convert_ToString(Me.obj.Convert_ToMinutes(.Ot_Weekday_2) + otWeekday2)
                        If .Ot_Weekday_2 = "00.00" Then
                            .Ot_Weekday_2 = "-"
                        End If
                        '' Tambahan aul 120915
                        If .Ot_Weekday_9 = "-" Then
                            .Ot_Weekday_9 = "00.00"
                        End If
                        .Ot_Weekday_9 = Me.obj.Convert_ToString(Me.obj.Convert_ToMinutes(.Ot_Weekday_9) + otWeekday9)
                        If .Ot_Weekday_9 = "00.00" Then
                            .Ot_Weekday_9 = "-"
                        End If
                        ''
                        If .Ot_Holiday_1 = "-" Then
                            .Ot_Holiday_1 = "00.00"
                        End If
                        .Ot_Holiday_1 = Me.obj.Convert_ToString(Me.obj.Convert_ToMinutes(.Ot_Holiday_1) + otHoliday1)
                        If .Ot_Holiday_1 = "00.00" Then
                            .Ot_Holiday_1 = "-"
                        End If
                        If .Ot_Holiday_8 = "-" Then
                            .Ot_Holiday_8 = "00.00"
                        End If
                        .Ot_Holiday_8 = Me.obj.Convert_ToString(Me.obj.Convert_ToMinutes(.Ot_Holiday_8) + otHoliday8)
                        If .Ot_Holiday_8 = "00.00" Then
                            .Ot_Holiday_8 = "-"
                        End If
                        If .Ot_Holiday_9 = "-" Then
                            .Ot_Holiday_9 = "00.00"
                        End If
                        .Ot_Holiday_9 = Me.obj.Convert_ToString(Me.obj.Convert_ToMinutes(.Ot_Holiday_9) + otHoliday9)
                        If .Ot_Holiday_9 = "00.00" Then
                            .Ot_Holiday_9 = "-"
                        End If
                        If PrimaryDeptSebelumnya = rowDb.Primary_Dept Then
                            TotalPermit += isPermit
                            TotalLate += isLate
                            TotalLateQty += isLateQty
                            TotalOTW1 += otWeekday1
                            TotalOTW2 += otWeekday2
                            TotalOTW9 += otWeekday9
                            TotalOTH1 += otHoliday1
                            TotalOTH8 += otHoliday8
                            TotalOTH9 += otHoliday9
                        Else
                            TotalPermit = isPermit
                            TotalLate = isLate
                            TotalLateQty = isLateQty
                            TotalOTW1 = otWeekday1
                            TotalOTW2 = otWeekday2
                            TotalOTW9 = otWeekday9
                            TotalOTH1 = otHoliday1
                            TotalOTH8 = otHoliday8
                            TotalOTH9 = otHoliday9
                        End If
                        .Total_Permit = Me.obj.Convert_ToString(TotalPermit)
                        If .Total_Permit = "00.00" Then
                            .Total_Permit = "-"
                        End If
                        .Total_Telat = Me.obj.Convert_ToString(TotalLate)
                        If .Total_Telat = "00.00" Then
                            .Total_Telat = "-"
                        End If
                        .Total_Telat_Qty = TotalLateQty
                        .Total_OT_Weekday_1 = Me.obj.Convert_ToString(TotalOTW1)
                        If .Total_OT_Weekday_1 = "00.00" Then
                            .Total_OT_Weekday_1 = "-"
                        End If
                        .Total_OT_Weekday_2 = Me.obj.Convert_ToString(TotalOTW2)
                        If .Total_OT_Weekday_2 = "00.00" Then
                            .Total_OT_Weekday_2 = "-"
                        End If
                        '' Tambahan aul 120915
                        .Total_OT_Weekday_9 = Me.obj.Convert_ToString(TotalOTW9)
                        If .Total_OT_Weekday_9 = "00.00" Then
                            .Total_OT_Weekday_9 = "-"
                        End If
                        ''
                        .Total_OT_Holiday_1 = Me.obj.Convert_ToString(TotalOTH1)
                        If .Total_OT_Holiday_1 = "00.00" Then
                            .Total_OT_Holiday_1 = "-"
                        End If
                        .Total_OT_Holiday_8 = Me.obj.Convert_ToString(TotalOTH8)
                        If .Total_OT_Holiday_8 = "00.00" Then
                            .Total_OT_Holiday_8 = "-"
                        End If
                        .Total_OT_Holiday_9 = Me.obj.Convert_ToString(TotalOTH9)
                        If .Total_OT_Holiday_9 = "00.00" Then
                            .Total_OT_Holiday_9 = "-"
                        End If
                    End With
                End If

                'hitung total(permit,telat,otw1,otw2,oth1,oth8,oth9) per department
                'periksa baris sebelumnya tblCetak berdasarkan dept code
                'jika tidak ditemukan, simpan langsung nilai tersebut
                'jika ditemukan, jumlahkan dengan nilai sebelumnya lalu simpan

                PrimaryDeptSebelumnya = rowDb.Primary_Dept
            Next

            ''kirim ke rpt
            If tblCetak.Rows.Count > 0 Then
                report.SetDataSource(ds)
                report.SetParameterValue("tglDari", vStrDate)
                report.SetParameterValue("tglSampai", vEndDate)
                vCrv.ReportSource = report
                vCrv.ShowRefreshButton = False
                vCrv.Show()
                vCrv.Refresh()
            Else
                MsgBox("No data for this period.", MsgBoxStyle.Information, "Information")
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class