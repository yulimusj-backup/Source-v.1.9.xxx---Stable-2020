Option Explicit On
Option Strict Off

Imports System.Windows.Forms, CrystalDecisions.Windows.Forms.CrystalReportViewer

Public Class FrmTrans_Attendance_Cetak
    Public mTblSetting As New DseTrans_Attendance_Cetak_IGI.TblSettingDataTable
    Public mIndexActual As Integer = 0
    Public mIndexMin As Integer = 0
    Public mIndexMax As Integer = 0
    'Penanda dipanggil dari form attendance atau quick attendance
    Public mShowQuickAttendance As Boolean = False
    'Settingan Single
    Public mPrimaryAttendance As Double = -1
    'Settingan Quick
    Public mPrimaryMinimumAttendance As Double = -1
    Public mListPrimaryAttendance As New List(Of Double)
    Public FromAttendanceEmail As Boolean = False
    Public FormatCetak As Integer = 0

    Dim ReportFormatA5_Batch As New CrpCetakTrans_Attendace_IGI_A5_Batch
    Dim ReportFormatA4_Batch As New CrpCetakTrans_Attendace_IGI_A4_Batch
    Dim ReportFormatA5 As New CrpCetakTrans_Attendace_IGI_A5
    Dim ReportFormatA4 As New CrpCetakTrans_Attendace_IGI_A4
    Dim Rowsetting As DseTrans_Attendance_Cetak_IGI.TblSettingRow
    Dim Objcd As New Core_Module.ClsCekData

    Public Enum EnumPrint As Integer
        Print
        Print_All
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
    Public Enum EnumWorkshift As Integer
        Workdays
        Holiday
    End Enum
    Public Enum EnumTipeFormatCetak As Integer
        FormatA5
        FormatA4
    End Enum
    Public Enum EnumOvertime As Integer
        No_Overtime
        Weekdays_Overtime
        Weekends_Overtime
    End Enum
    Public Sub Aktifkandetail()
        Try
            If Me.mShowQuickAttendance = False Then
                Me.TxtNoAttendance.Enabled = False
                Me.BtnPrevious.Enabled = False
                Me.BtnNext.Enabled = False
                Me.BtnPrintAll.Enabled = False
                Me.DtpEnd.Enabled = False
                Me.dtpStart.Enabled = False
                Me.BtnPrevious.Visible = False
                Me.BtnNext.Visible = False
                Me.BtnPrintAll.Visible = False
            Else
                Me.TxtNoAttendance.Enabled = False
                Me.BtnPrevious.Enabled = True
                Me.BtnNext.Enabled = True
                Me.BtnPrintAll.Enabled = True
                Me.DtpEnd.Enabled = False
                Me.dtpStart.Enabled = False
                Me.BtnPrevious.Visible = True
                Me.BtnNext.Visible = True
                Me.BtnPrintAll.Visible = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub btnShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub FrmTrans_Gaji_Cetak_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.CboFormat.SelectedIndex = 0
            If Me.FromAttendanceEmail = False Then
                Me.RboA5.Checked = True
                Me.Aktifkandetail()
                If Me.mTblSetting.Rows.Count > 0 Then
                    Rowsetting = Me.mTblSetting.Rows(0)
                    Me.TxtPrimaryAttendance.Text = Rowsetting.primary_transaksi
                    Me.DtpEnd.Value = Rowsetting.end_date
                    Me.dtpStart.Value = Rowsetting.start_date
                    Me.TxtIndex.Text = Rowsetting.index
                    Me.TxtNoAttendance.Text = Rowsetting.Employee_Code
                    If Me.mIndexActual = Me.mIndexMin Then
                        Me.BtnPrevious.Enabled = False
                    Else
                        Me.BtnPrevious.Enabled = True
                    End If
                End If
                If Me.FormatCetak = 0 Then 'Format cetak A5
                    Me.CetakAttendance_Setting(Me.crv1, Me.TxtPrimaryAttendance.Text, EnumTipeFormatCetak.FormatA5)
                Else 'Format cetak A4
                    Me.CetakAttendance_Setting(Me.crv1, Me.TxtPrimaryAttendance.Text, EnumTipeFormatCetak.FormatA4)
                End If
                'If Me.RboA5.Checked = True Then
                '    Me.CetakAttendance_Single_Setting(Me.crv1, Me.TxtPrimaryAttendance.Text, EnumTipeFormatCetak.FormatA5)
                'Else
                '    Me.CetakAttendance_Single_Setting(Me.crv1, Me.TxtPrimaryAttendance.Text, EnumTipeFormatCetak.FormatA4)
                'End If
            Else
                Me.Hide()
            End If
            Me.ShowIcon = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub CreatePDF(ByVal vPrimary As Double, ByVal vNoAttendance As String)
        Try
            Me.CetakAttendance_Single_Setting(Me.crv1, vPrimary, EnumTipeFormatCetak.FormatA4)
            Me.ReportFormatA4.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, _
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\Attendance-" + vNoAttendance + ".pdf")

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub CetakAttendance_Setting(ByVal vCRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, _
    ByVal vPrimaryAttendance As Double, ByVal vEnumTipeFormatCetak As EnumTipeFormatCetak, _
    Optional ByVal vEnumPrint As EnumPrint = EnumPrint.Print, Optional ByVal vNamaPrinter As String = "")
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
        Dim PrimaryAtt_Buffer As Double = -1
        Dim ListPrimaryAtt_Delete As New List(Of Double)
        Dim LastPrimary As Double = -1
        Dim ApaHitungTotal As Boolean = False

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'Fill data
            If Me.mTblSetting.Count > 0 Then
                Me.mTblSetting.DefaultView.RowFilter = Nothing
                Me.mTblSetting.DefaultView.Sort = "[primary transaksi] ASC"
                RowSetting = Me.mTblSetting.DefaultView(0).Row
                PrimaryAttMin = RowSetting.primary_transaksi
                RowSetting = Me.mTblSetting.DefaultView(Me.mTblSetting.DefaultView.Count - 1).Row
                PrimaryAttMax = RowSetting.primary_transaksi
                DTA.FillBy_PrimaryMinMax(TBL, PrimaryAttMin, PrimaryAttMax)
            End If

            'Koding Laporan
            If TBL.Rows.Count > 0 Then
                For i As Integer = 0 To TBL.Rows.Count - 1
                    Row = TBL.Rows(i)
                    ApaHitungTotal = True
                    With Row
                        Me.mTblSetting.DefaultView.RowFilter = "[primary transaksi] = " & .primary_transaksi
                        If Me.mTblSetting.DefaultView.Count > 0 Then
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

                                'tambah untuk early out
                                If String.IsNullOrEmpty(.Late_Leave_Sick) = False Then
                                    If .early_out_duration > 0 Then
                                        .Late_Leave_Sick += ", EO: " + Me.Objcd.Convert_ToString(.early_out_duration).ToString
                                    End If
                                Else
                                    If .early_out_duration > 0 Then
                                        .Late_Leave_Sick += "EO: " + Me.Objcd.Convert_ToString(.early_out_duration).ToString
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

                            If PrimaryAtt_Buffer <> .primary_transaksi And PrimaryAtt_Buffer <> -1 Then
HitungTotal:
                                TBL.DefaultView.RowFilter = "[primary transaksi] = " & PrimaryAtt_Buffer
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

                            PrimaryAtt_Buffer = .primary_transaksi
                            If ApaHitungTotal = True And i = TBL.Count - 1 Then
                                GoTo HitungTotal
                            End If
                        Else
                            If ListPrimaryAtt_Delete.Contains(.primary_transaksi) = False Then
                                ListPrimaryAtt_Delete.Add(.primary_transaksi)
                            End If
                        End If
                    End With
                Next

                'hapus attendace yang tidak ada di list cetak
                For i As Integer = TBL.Rows.Count - 1 To 0 Step -1
                    Row = TBL.Rows(i)
                    If ListPrimaryAtt_Delete.Contains(Row.primary_transaksi) = True Then
                        TBL.Rows.Remove(Row)
                    End If
                Next

                If TBL.Rows.Count > 0 Then
                    Row = TBL.Rows(TBL.Count - 1)
                    LastPrimary = Row.primary_transaksi
                End If

                'Setting Laporan
                If vEnumPrint = EnumPrint.Print Then
                    If vEnumTipeFormatCetak = EnumTipeFormatCetak.FormatA5 Then
                        Me.ReportFormatA5_Batch.SetDataSource(DS)
                        Me.ReportFormatA5_Batch.PrintOptions.PaperSize = CType(System.Drawing.Printing.PaperKind.A5, CrystalDecisions.Shared.PaperSize)
                        Me.ReportFormatA5_Batch.PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.Landscape
                        Me.ReportFormatA5_Batch.PrintOptions.ApplyPageMargins(New CrystalDecisions.Shared.PageMargins(0.63, 0.63, 0.63, 0.63))
                        Me.ReportFormatA5_Batch.SetParameterValue(Me.ReportFormatA5_Batch.Parameter_LastPrimary.ParameterFieldName, LastPrimary)
                        'Me.ReportFormatA5_Batch.PrintToPrinter(1, False, 0, 0)
                        vCRV.ReportSource = Me.ReportFormatA5_Batch
                        vCRV.Show()
                    ElseIf vEnumTipeFormatCetak = EnumTipeFormatCetak.FormatA4 Then
                        Me.ReportFormatA4_Batch.SetDataSource(DS)
                        Me.ReportFormatA4_Batch.SetParameterValue(Me.ReportFormatA4_Batch.Parameter_LastPrimary.ParameterFieldName, LastPrimary)
                        vCRV.ReportSource = Me.ReportFormatA4_Batch
                        vCRV.Show()
                    End If
                Else
                    If vEnumTipeFormatCetak = EnumTipeFormatCetak.FormatA5 Then

                    ElseIf vEnumTipeFormatCetak = EnumTipeFormatCetak.FormatA4 Then
                        Me.ReportFormatA4_Batch.SetDataSource(DS)
                        Me.ReportFormatA4_Batch.SetParameterValue(Me.ReportFormatA4_Batch.Parameter_LastPrimary.ParameterFieldName, LastPrimary)
                        vCRV.ReportSource = Me.ReportFormatA4_Batch
                        Me.ReportFormatA4_Batch.PrintToPrinter(1, True, 1, 100)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Private Sub SetTotalHitung(ByVal vPrimaryAtt As Double, ByVal vTBL As DseTrans_Attendance_Cetak_IGI.DataTable1DataTable)
        Dim Row As DseTrans_Attendance_Cetak_IGI.DataTable1Row

        Try
            'insert hasil hitung total
            vTBL.DefaultView.RowFilter = "[primary transaksi] = " & vPrimaryAtt
            For Each dv2 As DataRowView In vTBL.DefaultView
                Row = dv2.Row
                With Row
                    .Total_Weekdays_OT_1 = Me.Objcd.Convert_ToString(.OT_Weekday_1).ToString
                    .Total_Weekdays_OT_2 = Me.Objcd.Convert_ToString(.OT_Weekday_2).ToString
                    .Total_Weekdays_OT_3 = Me.Objcd.Convert_ToString(.OT_Weekday_3).ToString
                    .Total_Weekend_OT_1 = Me.Objcd.Convert_ToString(.OT_Weekend_1).ToString
                    .Total_Weekend_OT_2 = Me.Objcd.Convert_ToString(.OT_Weekend_2).ToString
                    .Total_Weekend_OT_3 = Me.Objcd.Convert_ToString(.OT_Weekend_3).ToString
                    .Total_Permit = Me.Objcd.Convert_ToString(.is_permit).ToString
                    .Total_Sick = Me.Objcd.Convert_ToString(.is_sick).ToString
                    .Total_Leave = Me.Objcd.Convert_ToString(.is_cuti).ToString
                    .Total_Late = Me.Objcd.Convert_ToString(.is_late).ToString
                    .Total_Outstation = Me.Objcd.Convert_ToString(.is_outstation).ToString
                    .Grand_Total = Me.Objcd.Convert_ToString(.actual_work_duration).ToString
                    .Total_Duration = Me.Objcd.Convert_ToString(.Permit_duration).ToString
                End With

            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub CetakAttendance_Single_Setting(ByVal vCRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, _
    ByVal vPrimaryAttendance As Double, ByVal vEnumTipeFormatCetak As EnumTipeFormatCetak, _
    Optional ByVal vEnumPrint As EnumPrint = EnumPrint.Print, Optional ByVal vNamaPrinter As String = "")
        Dim DS As New DseTrans_Attendance_Cetak_IGI
        Dim DTA As New DseTrans_Attendance_Cetak_IGITableAdapters.DataTable1TableAdapter
        Dim TBL As DseTrans_Attendance_Cetak_IGI.DataTable1DataTable = DS.DataTable1
        Dim Row As DseTrans_Attendance_Cetak_IGI.DataTable1Row
        'Parameter Total untuk parameter Field Total
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
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'Fill data
            DTA.FillBy_PrimaryTransaksi(TBL, vPrimaryAttendance)
            'Koding Laporan
            If TBL.Rows.Count > 0 Then
                For Each Row In TBL.Rows
                    Row.sch_in_format = Me.Objcd.Convert_ToString(Row.sch_in)
                    Row.sch_out_format = Me.Objcd.Convert_ToString(Row.sch_out)
                    If Row.actual_work_duration = 0 Then
                        Row.total_actual_work_duration_format = "-"
                    Else
                        Row.total_actual_work_duration_format = Me.Objcd.Convert_ToString(Row.actual_work_duration)
                    End If

                    If Row.presence_status = EnumStatus.Present Or _
                    Row.presence_status = EnumStatus.Half_day_Sick Or _
                    Row.presence_status = EnumStatus.Half_day_Leave Or _
                    Row.presence_status = EnumStatus.Half_day_Permit Then
                        Row.act_in_format = Me.Objcd.Convert_ToString(Row.actual_in)
                        Row.act_out_format = Me.Objcd.Convert_ToString(Row.actual_out)
                    ElseIf Row.presence_status = EnumStatus.Sick Or _
                    Row.presence_status = EnumStatus.Leave Or _
                    Row.presence_status = EnumStatus.Permit Or _
                    Row.presence_status = EnumStatus.Holiday Then
                        Row.act_in_format = "-"
                        Row.act_out_format = "-"
                    Else
                        Row.act_in_format = Me.Objcd.Convert_ToString(Row.actual_in)
                        Row.act_out_format = Me.Objcd.Convert_ToString(Row.actual_out)
                    End If

                    'mengambil data overtime 
                    If Row.tipe_overtime = EnumOvertime.Weekdays_Overtime Then
                        If Row.OT_Duration > 540 Then
                            Row.OT_Weekday_1 = 60
                            Row.OT_Weekday_2 = 480
                            Row.OT_Weekday_3 = Row.OT_Duration - 540
                        ElseIf Row.OT_Duration > 60 And Row.OT_Duration <= 540 Then
                            Row.OT_Weekday_1 = 60
                            Row.OT_Weekday_2 = Row.OT_Duration - 60
                            'If Row.OT_Weekday_2 > 90 Then
                            '    Row.OT_Weekday_2 = Row.OT_Weekday_2 - 30
                            'Else
                            '    Row.OT_Weekday_2 = Row.OT_Weekday_2
                            'End If
                        ElseIf Row.OT_Duration <= 60 Then
                            Row.OT_Weekday_1 = Row.OT_Duration
                        End If
                    End If
                    If Row.tipe_overtime = EnumOvertime.Weekends_Overtime Then
                        If Row.actual_work_duration > 540 Then
                            Row.OT_Weekend_1 = 480
                            Row.OT_Weekend_2 = 60
                            Row.OT_Weekend_3 = Row.actual_work_duration - 540
                        ElseIf Row.actual_work_duration > 480 And Row.actual_work_duration <= 540 Then
                            Row.OT_Weekend_1 = 480
                            Row.OT_Weekend_2 = Row.actual_work_duration - 480
                        ElseIf Row.actual_work_duration <= 480 Then
                            Row.OT_Weekend_1 = Row.actual_work_duration
                        End If
                    End If

                    If Row.tipe_overtime = EnumOvertime.Weekdays_Overtime Then
                        If Row.OT_Weekday_1 = 0 Then
                            Row.OT_Weekday_1_Format = "-"
                        Else
                            Row.OT_Weekday_1_Format = Me.Objcd.Convert_ToString(Row.OT_Weekday_1)
                        End If
                        If Row.OT_Weekday_2 = 0 Then
                            Row.OT_Weekday_2_Format = "-"
                        Else
                            Row.OT_Weekday_2_Format = Me.Objcd.Convert_ToString(Row.OT_Weekday_2)
                        End If
                        If Row.OT_Weekday_3 = 0 Then
                            Row.OT_Weekday_3_Format = "-"
                        Else
                            Row.OT_Weekday_3_Format = Me.Objcd.Convert_ToString(Row.OT_Weekday_3)
                        End If

                    ElseIf Row.tipe_overtime = EnumOvertime.Weekends_Overtime Then
                        If Row.OT_Weekend_1 = 0 Then
                            Row.OT_Weekday_1_Format = "-"
                        Else
                            Row.OT_Weekday_1_Format = Me.Objcd.Convert_ToString(Row.OT_Weekend_1)
                        End If
                        If Row.OT_Weekend_2 = 0 Then
                            Row.OT_Weekday_2_Format = "-"
                        Else
                            Row.OT_Weekday_2_Format = Me.Objcd.Convert_ToString(Row.OT_Weekend_2)
                        End If
                        If Row.OT_Weekend_3 = 0 Then
                            Row.OT_Weekday_3_Format = "-"
                        Else
                            Row.OT_Weekday_3_Format = Me.Objcd.Convert_ToString(Row.OT_Weekend_3)
                        End If
                    ElseIf Row.tipe_overtime = EnumOvertime.No_Overtime Then
                        Row.OT_Weekday_1_Format = "-"
                        Row.OT_Weekday_2_Format = "-"
                        Row.OT_Weekday_3_Format = "-"
                    End If


                    'Total
                    Total_Weekday_OT_1 += Row.OT_Weekday_1
                    Total_Weekday_OT_2 += Row.OT_Weekday_2
                    Total_Weekday_OT_3 += Row.OT_Weekday_3
                    Total_Weekend_OT_1 += Row.OT_Weekend_1
                    Total_Weekend_OT_2 += Row.OT_Weekend_2
                    Total_Weekend_OT_3 += Row.OT_Weekend_3
                    Row.late_duration_format = Me.Objcd.Convert_ToString(Row.late_duration)
                    Total_Permit += Row.is_permit
                    Total_Sick += Row.is_sick
                    Total_Leave += Row.is_cuti
                    Total_Late += Row.is_late
                    Total_Outstation += Row.is_outstation
                    Sub_Total_Actual += Row.actual_work_duration
                    If Row.Permit_duration > 0 Then
                        Total_Permit_duration += Row.Permit_duration
                    End If
                    'Leave, Late, Sick
                    If Row.presence_status = EnumStatus.Present Then
                        If Row.late_duration > 0 Then
                            Row.Late_Leave_Sick = "Late: " + Row.late_duration_format.ToString
                            If Row.Permit_duration > 0 Then
                                If String.IsNullOrEmpty(Row.Late_Leave_Sick) = False Then
                                    Row.Late_Leave_Sick += ", permit: " + Me.Objcd.Convert_ToString(Row.Permit_duration).ToString
                                End If
                            End If
                        Else
                            If Row.Permit_duration > 0 Then
                                Row.Late_Leave_Sick = "Permit: " + Me.Objcd.Convert_ToString(Row.Permit_duration).ToString
                            End If
                        End If
                    End If
                    If Row.presence_status = EnumStatus.Sick Then
                        Row.Late_Leave_Sick = "Sick"
                    End If
                    If Row.presence_status = EnumStatus.Half_day_Sick Then
                        Row.Late_Leave_Sick = "Half day sick"
                    End If
                    If Row.presence_status = EnumStatus.Leave Then
                        If String.IsNullOrEmpty(Row.cuti_description) = False Then
                            Row.Late_Leave_Sick = "Leave: " + Row.cuti_description.ToString
                        End If
                    End If
                    If Row.presence_status = EnumStatus.Half_day_Leave Then
                        If String.IsNullOrEmpty(Row.cuti_description) = False Then
                            Row.Late_Leave_Sick = "Half day leave: " + Row.cuti_description.ToString
                        End If
                    End If
                    If Row.presence_status = EnumStatus.Permit Then
                        Row.Late_Leave_Sick = "Permit"
                    End If
                    If Row.presence_status = EnumStatus.Half_day_Permit Then
                        Row.Late_Leave_Sick = "Half day permit"
                    End If
                    If String.IsNullOrEmpty(Row.Late_Leave_Sick) = False Then
                        If String.IsNullOrEmpty(Row.remarks) = False Then
                            Row.Late_Leave_Sick += "-" + Row.remarks.ToString
                        End If
                    Else
                        If String.IsNullOrEmpty(Row.remarks) = False Then
                            Row.Late_Leave_Sick = Row.remarks.ToString
                        End If
                    End If
                Next

                'Dim X As New Formx
                'X.DataGridView1.DataSource = Nothing
                'X.DataGridView1.DataSource = TBL
                'X.ShowDialog()
                'Setting Laporan
                If vEnumPrint = EnumPrint.Print Then
                    If vEnumTipeFormatCetak = EnumTipeFormatCetak.FormatA5 Then
                        Me.ReportFormatA5.SetDataSource(DS)
                        Me.ReportFormatA5.SetParameterValue(Me.ReportFormatA5.Parameter_Total_Weekdays_OT_1.ParameterFieldName, Me.Objcd.Convert_ToString(Total_Weekday_OT_1))
                        Me.ReportFormatA5.SetParameterValue(Me.ReportFormatA5.Parameter_Total_Weekdays_OT_2.ParameterFieldName, Me.Objcd.Convert_ToString(Total_Weekday_OT_2))
                        Me.ReportFormatA5.SetParameterValue(Me.ReportFormatA5.Parameter_Total_Weekdays_OT_3.ParameterFieldName, Me.Objcd.Convert_ToString(Total_Weekday_OT_3))
                        Me.ReportFormatA5.SetParameterValue(Me.ReportFormatA5.Parameter_Total_Weekend_OT_1.ParameterFieldName, Me.Objcd.Convert_ToString(Total_Weekend_OT_1))
                        Me.ReportFormatA5.SetParameterValue(Me.ReportFormatA5.Parameter_Total_Weekend_OT_2.ParameterFieldName, Me.Objcd.Convert_ToString(Total_Weekend_OT_2))
                        Me.ReportFormatA5.SetParameterValue(Me.ReportFormatA5.Parameter_Total_Weekend_OT_3.ParameterFieldName, Me.Objcd.Convert_ToString(Total_Weekend_OT_3))
                        Me.ReportFormatA5.SetParameterValue(Me.ReportFormatA5.Parameter_Total_Leave.ParameterFieldName, Total_Leave)
                        Me.ReportFormatA5.SetParameterValue(Me.ReportFormatA5.Parameter_Total_Sick.ParameterFieldName, Total_Sick)
                        Me.ReportFormatA5.SetParameterValue(Me.ReportFormatA5.Parameter_Total_Permit.ParameterFieldName, Total_Permit)
                        Me.ReportFormatA5.SetParameterValue(Me.ReportFormatA5.Parameter_Total_Late.ParameterFieldName, Total_Late)
                        Me.ReportFormatA5.SetParameterValue(Me.ReportFormatA5.Parameter_Total_Outstation.ParameterFieldName, Total_Outstation)
                        Me.ReportFormatA5.SetParameterValue(Me.ReportFormatA5.Parameter_Grand_Total_Actual.ParameterFieldName, Me.Objcd.Convert_ToString(Sub_Total_Actual))
                        Me.ReportFormatA5.SetParameterValue(Me.ReportFormatA5.Parameter_Total_Duration.ParameterFieldName, Me.Objcd.Convert_ToString(Total_Permit_duration))
                        Me.ReportFormatA5.PrintOptions.PaperSize = CType(System.Drawing.Printing.PaperKind.A5, CrystalDecisions.Shared.PaperSize)
                        Me.ReportFormatA5.PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.Landscape
                        Me.ReportFormatA5.PrintOptions.ApplyPageMargins(New CrystalDecisions.Shared.PageMargins(0.63, 0.63, 0.63, 0.63))
                        'Me.ReportFormatA5.PrintToPrinter(1, False, 0, 0)
                        vCRV.ReportSource = Me.ReportFormatA5
                        vCRV.Show()
                    ElseIf vEnumTipeFormatCetak = EnumTipeFormatCetak.FormatA4 Then
                        Me.ReportFormatA4.SetDataSource(DS)
                        Me.ReportFormatA4.SetParameterValue(Me.ReportFormatA4.Parameter_Total_Weekdays_OT_1.ParameterFieldName, Me.Objcd.Convert_ToString(Total_Weekday_OT_1))
                        Me.ReportFormatA4.SetParameterValue(Me.ReportFormatA4.Parameter_Total_Weekdays_OT_2.ParameterFieldName, Me.Objcd.Convert_ToString(Total_Weekday_OT_2))
                        Me.ReportFormatA4.SetParameterValue(Me.ReportFormatA4.Parameter_Total_Weekdays_OT_3.ParameterFieldName, Me.Objcd.Convert_ToString(Total_Weekday_OT_3))
                        Me.ReportFormatA4.SetParameterValue(Me.ReportFormatA4.Parameter_Total_Weekend_OT_1.ParameterFieldName, Me.Objcd.Convert_ToString(Total_Weekend_OT_1))
                        Me.ReportFormatA4.SetParameterValue(Me.ReportFormatA4.Parameter_Total_Weekend_OT_2.ParameterFieldName, Me.Objcd.Convert_ToString(Total_Weekend_OT_2))
                        Me.ReportFormatA4.SetParameterValue(Me.ReportFormatA4.Parameter_Total_Weekend_OT_3.ParameterFieldName, Me.Objcd.Convert_ToString(Total_Weekend_OT_3))
                        Me.ReportFormatA4.SetParameterValue(Me.ReportFormatA4.Parameter_Total_Leave.ParameterFieldName, Total_Leave)
                        Me.ReportFormatA4.SetParameterValue(Me.ReportFormatA4.Parameter_Total_Sick.ParameterFieldName, Total_Sick)
                        Me.ReportFormatA4.SetParameterValue(Me.ReportFormatA4.Parameter_Total_Permit.ParameterFieldName, Total_Permit)
                        Me.ReportFormatA4.SetParameterValue(Me.ReportFormatA4.Parameter_Total_Late.ParameterFieldName, Total_Late)
                        Me.ReportFormatA4.SetParameterValue(Me.ReportFormatA4.Parameter_Total_Outstation.ParameterFieldName, Total_Outstation)
                        Me.ReportFormatA4.SetParameterValue(Me.ReportFormatA4.Parameter_Grand_Total.ParameterFieldName, Me.Objcd.Convert_ToString(Sub_Total_Actual))
                        Me.ReportFormatA4.SetParameterValue(Me.ReportFormatA4.Parameter_Total_duration.ParameterFieldName, Me.Objcd.Convert_ToString(Total_Permit_duration))
                        vCRV.ReportSource = Me.ReportFormatA4
                        vCRV.Show()
                    End If
                Else
                    If vEnumTipeFormatCetak = EnumTipeFormatCetak.FormatA5 Then
                        'Me.ReportFormatA5.SetDataSource(DS)
                        'Me.ReportFormatA5.SetParameterValue(Me.ReportFormatA5.Parameter_Total_Weekdays_OT_1.ParameterFieldName, Me.Objcd.Convert_ToString(Total_Weekday_OT_1))
                        'Me.ReportFormatA5.SetParameterValue(Me.ReportFormatA5.Parameter_Total_Weekdays_OT_2.ParameterFieldName, Me.Objcd.Convert_ToString(Total_Weekday_OT_2))
                        'Me.ReportFormatA5.SetParameterValue(Me.ReportFormatA5.Parameter_Total_Weekdays_OT_3.ParameterFieldName, Me.Objcd.Convert_ToString(Total_Weekday_OT_3))
                        'Me.ReportFormatA5.SetParameterValue(Me.ReportFormatA5.Parameter_Total_Weekend_OT_1.ParameterFieldName, Me.Objcd.Convert_ToString(Total_Weekend_OT_1))
                        'Me.ReportFormatA5.SetParameterValue(Me.ReportFormatA5.Parameter_Total_Weekend_OT_2.ParameterFieldName, Me.Objcd.Convert_ToString(Total_Weekend_OT_2))
                        'Me.ReportFormatA5.SetParameterValue(Me.ReportFormatA5.Parameter_Total_Weekend_OT_3.ParameterFieldName, Me.Objcd.Convert_ToString(Total_Weekend_OT_3))
                        'Me.ReportFormatA5.SetParameterValue(Me.ReportFormatA5.Parameter_Total_Leave.ParameterFieldName, Total_Leave)
                        'Me.ReportFormatA5.SetParameterValue(Me.ReportFormatA5.Parameter_Total_Sick.ParameterFieldName, Total_Sick)
                        'Me.ReportFormatA5.SetParameterValue(Me.ReportFormatA5.Parameter_Total_Permit.ParameterFieldName, Total_Permit)
                        'Me.ReportFormatA5.SetParameterValue(Me.ReportFormatA5.Parameter_Total_Late.ParameterFieldName, Total_Late)
                        'Me.ReportFormatA5.SetParameterValue(Me.ReportFormatA5.Parameter_Total_Outstation.ParameterFieldName, Total_Outstation)
                        'Me.ReportFormatA5.SetParameterValue(Me.ReportFormatA5.Parameter_Grand_Total_Actual.ParameterFieldName, Me.Objcd.Convert_ToString(Sub_Total_Actual))
                        'Me.ReportFormatA5.SetParameterValue(Me.ReportFormatA5.Parameter_Total_Duration.ParameterFieldName, Me.Objcd.Convert_ToString(Total_Permit_duration))
                        'vCRV.ReportSource = Me.ReportFormatA5
                        'Me.ReportFormatA5.PrintToPrinter(1, True, 1, 100)
                    ElseIf vEnumTipeFormatCetak = EnumTipeFormatCetak.FormatA4 Then
                        Me.ReportFormatA4.SetDataSource(DS)
                        Me.ReportFormatA4.SetParameterValue(Me.ReportFormatA4.Parameter_Total_Weekdays_OT_1.ParameterFieldName, Me.Objcd.Convert_ToString(Total_Weekday_OT_1))
                        Me.ReportFormatA4.SetParameterValue(Me.ReportFormatA4.Parameter_Total_Weekdays_OT_2.ParameterFieldName, Me.Objcd.Convert_ToString(Total_Weekday_OT_2))
                        Me.ReportFormatA4.SetParameterValue(Me.ReportFormatA4.Parameter_Total_Weekdays_OT_3.ParameterFieldName, Me.Objcd.Convert_ToString(Total_Weekday_OT_3))
                        Me.ReportFormatA4.SetParameterValue(Me.ReportFormatA4.Parameter_Total_Weekend_OT_1.ParameterFieldName, Me.Objcd.Convert_ToString(Total_Weekend_OT_1))
                        Me.ReportFormatA4.SetParameterValue(Me.ReportFormatA4.Parameter_Total_Weekend_OT_2.ParameterFieldName, Me.Objcd.Convert_ToString(Total_Weekend_OT_2))
                        Me.ReportFormatA4.SetParameterValue(Me.ReportFormatA4.Parameter_Total_Weekend_OT_3.ParameterFieldName, Me.Objcd.Convert_ToString(Total_Weekend_OT_3))
                        Me.ReportFormatA4.SetParameterValue(Me.ReportFormatA4.Parameter_Total_Leave.ParameterFieldName, Total_Leave)
                        Me.ReportFormatA4.SetParameterValue(Me.ReportFormatA4.Parameter_Total_Sick.ParameterFieldName, Total_Sick)
                        Me.ReportFormatA4.SetParameterValue(Me.ReportFormatA4.Parameter_Total_Permit.ParameterFieldName, Total_Permit)
                        Me.ReportFormatA4.SetParameterValue(Me.ReportFormatA4.Parameter_Total_Late.ParameterFieldName, Total_Late)
                        Me.ReportFormatA4.SetParameterValue(Me.ReportFormatA4.Parameter_Total_Outstation.ParameterFieldName, Total_Outstation)
                        Me.ReportFormatA4.SetParameterValue(Me.ReportFormatA4.Parameter_Grand_Total.ParameterFieldName, Me.Objcd.Convert_ToString(Sub_Total_Actual))
                        Me.ReportFormatA4.SetParameterValue(Me.ReportFormatA4.Parameter_Total_duration.ParameterFieldName, Me.Objcd.Convert_ToString(Total_Permit_duration))
                        vCRV.ReportSource = Me.ReportFormatA4
                        Me.ReportFormatA4.PrintToPrinter(1, True, 1, 100)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Private Function DelayFunction(ByVal mDelay As Integer) As Boolean
        Dim SW As New Stopwatch
        Dim Start As Integer = 0
        Try
            SW.Start()
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            While Start = 0
                If SW.Elapsed.TotalSeconds >= mDelay Then
                    Exit While
                End If
            End While
            SW.Stop()
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub RboA5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RboA5.CheckedChanged
        Try
            If String.IsNullOrEmpty(Me.TxtPrimaryAttendance.Text) = True Then
                Me.TxtPrimaryAttendance.Text = "-1"
            End If
            Me.DelayFunction(0.7)
            If Me.RboA5.Checked = True Then
                Me.CetakAttendance_Single_Setting(Me.crv1, Me.TxtPrimaryAttendance.Text, EnumTipeFormatCetak.FormatA5)
            Else
                Me.CetakAttendance_Single_Setting(Me.crv1, Me.TxtPrimaryAttendance.Text, EnumTipeFormatCetak.FormatA4)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Private Sub RboA4_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RboA4.CheckedChanged
    '    Try
    '        If String.IsNullOrEmpty(Me.TxtPrimaryAttendance.Text) = True Then
    '            Me.TxtPrimaryAttendance.Text = "-1"
    '        End If
    '        Me.DelayFunction(0.6)
    '        If Me.RboA5.Checked = True Then
    '            Me.CetakAttendance_Single_Setting(Me.crv1, Me.TxtPrimaryAttendance.Text, EnumTipeFormatCetak.FormatA5)
    '        Else
    '            Me.CetakAttendance_Single_Setting(Me.crv1, Me.TxtPrimaryAttendance.Text, EnumTipeFormatCetak.FormatA4)
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Private Sub BtnPrevious_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnPrevious.Click
        Dim Filter As String = ""
        Try
            If Me.mTblSetting.Rows.Count > 0 Then
                Me.mIndexActual -= 1
                If Me.mIndexActual < Me.mIndexMin Then
                    Me.mIndexActual = Me.mIndexMin
                Else
                    Filter = "[index] =" + Me.mIndexActual.ToString + ""
                    Me.mTblSetting.DefaultView.RowFilter = Filter
                    If Me.mTblSetting.Rows.Count > 0 Then
                        Rowsetting = Me.mTblSetting.DefaultView(0).Row
                        Me.TxtPrimaryAttendance.Text = Rowsetting.primary_transaksi
                        Me.TxtNoAttendance.Text = Rowsetting.Employee_Code
                        Me.DtpEnd.Value = Rowsetting.end_date
                        Me.dtpStart.Value = Rowsetting.start_date
                        Me.TxtIndex.Text = Rowsetting.index
                    End If
                    Me.mTblSetting.DefaultView.RowFilter = ""
                    If Me.mIndexActual = Me.mIndexMin Then
                        Me.BtnPrevious.Enabled = False
                    Else
                        Me.BtnPrevious.Enabled = True
                        Me.BtnNext.Enabled = True
                    End If
                    If Me.RboA4.Checked = True Then
                        Me.CetakAttendance_Single_Setting(Me.crv1, Me.TxtPrimaryAttendance.Text, EnumTipeFormatCetak.FormatA4)
                    Else
                        Me.CetakAttendance_Single_Setting(Me.crv1, Me.TxtPrimaryAttendance.Text, EnumTipeFormatCetak.FormatA5)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnNext.Click
        Dim Filter As String = ""
        Try
            If Me.mTblSetting.Rows.Count > 0 Then
                Me.mIndexActual += 1
                If Me.mIndexActual > Me.mIndexMax Then
                    Me.mIndexActual = Me.mIndexMax
                Else
                    Filter = "[index] =" + Me.mIndexActual.ToString + ""
                    Me.mTblSetting.DefaultView.RowFilter = Filter
                    If Me.mTblSetting.Rows.Count > 0 Then
                        Rowsetting = Me.mTblSetting.DefaultView(0).Row
                        Me.TxtPrimaryAttendance.Text = Rowsetting.primary_transaksi
                        Me.TxtNoAttendance.Text = Rowsetting.Employee_Code
                        Me.DtpEnd.Value = Rowsetting.end_date
                        Me.dtpStart.Value = Rowsetting.start_date
                        Me.TxtIndex.Text = Rowsetting.index
                    End If
                    Me.mTblSetting.DefaultView.RowFilter = ""
                    If Me.mIndexActual = Me.mIndexMax Then
                        Me.BtnNext.Enabled = False
                    Else
                        Me.BtnPrevious.Enabled = True
                        Me.BtnNext.Enabled = True
                    End If
                    If Me.RboA4.Checked = True Then
                        Me.CetakAttendance_Single_Setting(Me.crv1, Me.TxtPrimaryAttendance.Text, EnumTipeFormatCetak.FormatA4)
                    Else
                        Me.CetakAttendance_Single_Setting(Me.crv1, Me.TxtPrimaryAttendance.Text, EnumTipeFormatCetak.FormatA5)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnPrintAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnPrintAll.Click
        Dim PrinterSetting As String = ""
        Dim Hasil As MsgBoxResult
        Try
            Hasil = MsgBox("Do you want to print all this page ?", MsgBoxStyle.YesNo, "Confirm")
            If Hasil = MsgBoxResult.Yes Then
                If Me.mTblSetting.Rows.Count > 0 Then
                    Me.PrintDialog1.ShowDialog()
                    If Me.RboA4.Checked = True Then
                        Me.ReportFormatA4.PrintOptions.PrinterName = Me.PrintDialog1.PrinterSettings.PrinterName
                        'Me.ReportFormatA4.PrintOptions.PrinterName = "\\FITRI-PC\HP Deskjet 1010 series (Copy 1)"
                    Else
                        Me.ReportFormatA5.PrintOptions.PrinterName = Me.PrintDialog1.PrinterSettings.PrinterName
                        'Me.ReportFormatA5.PrintOptions.PrinterName = "\\FITRI-PC\HP Deskjet 1010 series (Copy 1)"
                    End If
                    For Each Rowsetting In Me.mTblSetting.Rows
                        If Me.RboA4.Checked = True Then
                            Me.CetakAttendance_Single_Setting(Me.crv1, Rowsetting.primary_transaksi, EnumTipeFormatCetak.FormatA4, EnumPrint.Print_All)
                        Else
                            Me.CetakAttendance_Single_Setting(Me.crv1, Rowsetting.primary_transaksi, EnumTipeFormatCetak.FormatA5, EnumPrint.Print_All)
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Dim BufferFormat As Integer = 0
    Private Sub CboFormat_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboFormat.SelectedIndexChanged
        Try
            Me.Cursor = Cursors.WaitCursor
            'Me.DelayFunction(0.7)
            If Me.CboFormat.SelectedIndex <> Me.BufferFormat Then
                If Me.CboFormat.SelectedIndex = 0 Then
                    Me.BufferFormat = 0
                    Me.CetakAttendance_Setting(Me.crv1, Me.TxtPrimaryAttendance.Text, EnumTipeFormatCetak.FormatA5)
                Else
                    Me.BufferFormat = 1
                    Me.CetakAttendance_Setting(Me.crv1, Me.TxtPrimaryAttendance.Text, EnumTipeFormatCetak.FormatA4)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub
End Class