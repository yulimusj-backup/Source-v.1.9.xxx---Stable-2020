Option Explicit On
Option Strict Off

Public Class FrmMasterWorkShift
    Dim txtSelect As TextBox

    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public TipeWorkshift As ClsMasterWorkShift.EnumTipeWorkshift = ClsMasterWorkShift.EnumTipeWorkshift.HR
    Dim baru As Boolean = True
    Dim buffer As String
    Dim ObjCd As New ClsCekData
    Dim ObjWorkShift As New ClsMasterWorkShift
    Dim ObjWorkShiftItem As New ClsMasterWorkShift_Item
    Dim ObjCuti As New Core_Module.ClsMasterCuti
    Dim EraseRowIndex As Integer = 0

    Dim dtaAtt As New DseMasterWorkShiftTableAdapters.tbl_trans_att_viewTableAdapter
    Dim DateDgv As Date

    Public Sub SetDgvWeekly()
        Dim rowDgv As DataGridViewRow
        Dim CboDgv As DataGridViewComboBoxCell
        Try
            For Each day As DayOfWeek In [Enum].GetValues(GetType(DayOfWeek))
                dgvDailyActivities.Rows.Add("")
                rowDgv = dgvDailyActivities.Rows(dgvDailyActivities.Rows.Count - 1)
                rowDgv.Cells(Me.ColDay.Index).Value = Me.ObjCd.GetHari_English(day)
                CboDgv = rowDgv.Cells(Me.ColWorkdays.Index)
                rowDgv.Cells(Me.ColWorkdays.Index).Value = CboDgv.Items.Item(ClsMasterWorkShift.TipeWorkdays.Weekday)
                CboDgv = rowDgv.Cells(Me.ColOvertime.Index)
                rowDgv.Cells(Me.ColOvertime.Index).Value = CboDgv.Items.Item(ClsMasterWorkShift.TipeOvertime.Weekday)
                CboDgv = rowDgv.Cells(Me.ColDuration.Index)
                rowDgv.Cells(Me.ColDuration.Index).Value = CboDgv.Items.Item(ClsMasterWorkShift.enumDuration.fullday)
                rowDgv.Cells(Me.ColSchIn.Index).Value = "00.00"
                rowDgv.Cells(Me.ColSchOut.Index).Value = "00.00"
                rowDgv.Cells(Me.ColSchBreakIn_WA.Index).Value = "00.00"
                rowDgv.Cells(Me.ColSchBreakOut_WA.Index).Value = "00.00"
                rowDgv.Cells(Me.COlEnumDayOfWeek.Index).Value = day
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SetDgvAttSettings()
        Try
            Dim DgvRow As DataGridViewRow

            Me.DgvAttParams.Rows.Add("")
            DgvRow = Me.DgvAttParams.Rows(Me.DgvAttParams.Rows.Count - 1)
            DgvRow.Cells(Me.ColAtt_Schedule.Index).Value = "Schedule In"
            DgvRow.Cells(Me.ColAtt_GetValue.Index).Value = Boolean.TrueString
            DgvRow.Cells(Me.ColAtt_LimitBefore.Index).Value = "0"
            DgvRow.Cells(Me.ColAtt_LimitAfter.Index).Value = "0"
            DgvRow.Cells(Me.ColAtt_FirstOrLast.Index).Value = "First"
            DgvRow.Cells(Me.ColAtt_Default.Index).Value = "Schedule"
            DgvRow.Cells(Me.ColAtt_Outstation.Index).Value = "Schedule"
            DgvRow.Cells(Me.ColAtt_DefaultCustom.Index).Value = ""
            DgvRow.Cells(Me.ColAtt_ScheduleType.Index).Value = ClsMasterWorkShift.EnumScheduleType.SCHEDULE_IN

            Me.DgvAttParams.Rows.Add("")
            DgvRow = Me.DgvAttParams.Rows(Me.DgvAttParams.Rows.Count - 1)
            DgvRow.Cells(Me.ColAtt_Schedule.Index).Value = "Break Out"
            DgvRow.Cells(Me.ColAtt_GetValue.Index).Value = Boolean.FalseString
            DgvRow.Cells(Me.ColAtt_LimitBefore.Index).Value = "0"
            DgvRow.Cells(Me.ColAtt_LimitAfter.Index).Value = "0"
            DgvRow.Cells(Me.ColAtt_FirstOrLast.Index).Value = "Last"
            DgvRow.Cells(Me.ColAtt_Default.Index).Value = "Schedule"
            DgvRow.Cells(Me.ColAtt_Outstation.Index).Value = "Schedule"
            DgvRow.Cells(Me.ColAtt_DefaultCustom.Index).Value = ""
            DgvRow.Cells(Me.ColAtt_ScheduleType.Index).Value = ClsMasterWorkShift.EnumScheduleType.BREAK_OUT

            Me.DgvAttParams.Rows.Add("")
            DgvRow = Me.DgvAttParams.Rows(Me.DgvAttParams.Rows.Count - 1)
            DgvRow.Cells(Me.ColAtt_Schedule.Index).Value = "Break In"
            DgvRow.Cells(Me.ColAtt_GetValue.Index).Value = Boolean.FalseString
            DgvRow.Cells(Me.ColAtt_LimitBefore.Index).Value = "0"
            DgvRow.Cells(Me.ColAtt_LimitAfter.Index).Value = "0"
            DgvRow.Cells(Me.ColAtt_FirstOrLast.Index).Value = "First"
            DgvRow.Cells(Me.ColAtt_Default.Index).Value = "Schedule"
            DgvRow.Cells(Me.ColAtt_Outstation.Index).Value = "Schedule"
            DgvRow.Cells(Me.ColAtt_DefaultCustom.Index).Value = ""
            DgvRow.Cells(Me.ColAtt_ScheduleType.Index).Value = ClsMasterWorkShift.EnumScheduleType.BREAK_IN

            Me.DgvAttParams.Rows.Add("")
            DgvRow = Me.DgvAttParams.Rows(Me.DgvAttParams.Rows.Count - 1)
            DgvRow.Cells(Me.ColAtt_Schedule.Index).Value = "Schedule Out"
            DgvRow.Cells(Me.ColAtt_GetValue.Index).Value = Boolean.TrueString
            DgvRow.Cells(Me.ColAtt_LimitBefore.Index).Value = "0"
            DgvRow.Cells(Me.ColAtt_LimitAfter.Index).Value = "0"
            DgvRow.Cells(Me.ColAtt_FirstOrLast.Index).Value = "Last"
            DgvRow.Cells(Me.ColAtt_Default.Index).Value = "Schedule"
            DgvRow.Cells(Me.ColAtt_Outstation.Index).Value = "Schedule"
            DgvRow.Cells(Me.ColAtt_DefaultCustom.Index).Value = ""
            DgvRow.Cells(Me.ColAtt_ScheduleType.Index).Value = ClsMasterWorkShift.EnumScheduleType.SCHEDULE_OUT
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim row As DataGridViewRow

        Dim CboDgv As DataGridViewComboBoxCell
        'Dim MsgHasil As MsgBoxResult
        Try
            'tidak boleh kosong
            KoleksiNamaData.Add("Code")
            KoleksiTextBox.Add(Me.TxtJobCode)
            'KoleksiNamaData.Add("Description")
            'KoleksiTextBox.Add(Me.TxtJobName)
            If Me.ObjCd.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If
            KoleksiNamaData.Clear()
            KoleksiTextBox.Clear()

            'harus memiliki minimal 1 item
            If Me.dgvDailyActivities.RowCount < 7 Then
                MsgBox("Please input all days in weekly activities.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'tidak boleh double
            If Me.baru = True Then
                If Me.ObjWorkShift.ApakahSudahAda(Me.TxtJobCode.Text, Me.TipeWorkshift) = True Then
                    MsgBox("Duplication in code, please change your code", MsgBoxStyle.Information, "Information")
                    Me.TxtJobCode.Focus()
                    Return False
                End If
            Else
                Dim CodeLama As String = Me.ObjWorkShift.GetCodeByPrimary(Me.TxtPrimaryJob.Text)
                If CodeLama <> Me.TxtJobCode.Text Then
                    If Me.ObjWorkShift.ApakahSudahAda(Me.TxtJobCode.Text, Me.TipeWorkshift) = True Then
                        MsgBox("Duplication in code, please change your code", MsgBoxStyle.Information, "Information")
                        Me.TxtJobCode.Focus()
                        Return False
                    End If
                End If

                'data salah jika nilai time stamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses(edit)
                Dim OldTimeStamp As Date = Me.ObjWorkShift.GetTimeStampByPrimary(Me.TxtPrimaryJob.Text)
                If Me.ObjWorkShift.TimeStamp <> OldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.ObjWorkShift.CariDetail(Me.ObjWorkShift.Primary, Me, Me.dgvDailyActivities, Me.DgvHolidayList, Me.ObjCd)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            'DgvWeekly
            Dim BatasDurasi As Double = 0
            Dim Durasi As Double = 0
            Dim SchIn As Double = 0
            Dim SchBreakOut As Double = 0
            Dim SchBreakIn As Double = 0
            Dim SchOut As Double = 0
            Dim BreakDuration As Double = 0
            Dim ListSalah As String = ""

            ListSalah = ""
            For Each row In dgvDailyActivities.Rows
                SchIn = Me.ObjCd.Convert_ToMinutes(row.Cells(Me.ColSchIn.Index).Value)
                SchBreakOut = Me.ObjCd.Convert_ToMinutes(row.Cells(Me.ColSchBreakOut_WA.Index).Value)
                SchBreakIn = Me.ObjCd.Convert_ToMinutes(row.Cells(Me.ColSchBreakIn_WA.Index).Value)
                SchOut = Me.ObjCd.Convert_ToMinutes(row.Cells(Me.ColSchOut.Index).Value)
                BreakDuration = Me.ObjWorkShift.HitungBreakDuration(SchIn, SchOut, SchBreakOut, SchBreakIn)

                'jika semua baris 00.00
                If SchIn = 0 And _
                SchBreakOut = 0 And _
                SchBreakIn = 0 And _
                SchOut = 0 Then
                    Me.TabControl1.SelectedIndex = 0
                    Me.dgvDailyActivities.ClearSelection()
                    Me.dgvDailyActivities.Rows.Item(row.Index).Selected = True
                    MsgBox("Work duration : " + row.Cells(Me.ColDay.Index).Value.ToString + vbNewLine + _
                    "Schedule in , schedule break out , schedule break in and schedule out cannot fill 00.00.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
                'cari durasinya
                If SchIn < SchOut Then
                    'hitung schOut - schIn
                    Durasi = (SchOut - SchIn) - BreakDuration
                ElseIf SchIn > SchOut Then
                    'hitung (24.00-schIn)+(schOut-00.00)
                    Durasi = (((24 * 60) - SchIn) + SchOut) - BreakDuration
                ElseIf SchIn = SchOut Then
                    Me.TabControl1.SelectedIndex = 0
                    Me.dgvDailyActivities.ClearSelection()
                    Me.dgvDailyActivities.Rows.Item(row.Index).Selected = True
                    MsgBox("Work duration : " + row.Cells(Me.ColDay.Index).Value.ToString + _
                    " ,Schedule in not be the same with schedule out.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
                'cek duration, jika fullday batasan 8 jam, jika halfday batasan 4 jam.
                'dalam menit
                CboDgv = row.Cells(Me.ColDuration.Index)
                If row.Cells(Me.ColDuration.Index).Value = CboDgv.Items.Item(ClsMasterWorkShift.enumDuration.fullday) Then
                    BatasDurasi = 480
                    'jika durasi melebihi batas durasi kasih peringatan
                    If Durasi > BatasDurasi Then
                        ListSalah = ListSalah + "Work duration : " + row.Cells(Me.ColDay.Index).Value + " ,is larger than " _
                        + (BatasDurasi / 60).ToString + " hours." + vbNewLine
                    End If
                ElseIf row.Cells(Me.ColDuration.Index).Value = CboDgv.Items.Item(ClsMasterWorkShift.enumDuration.halfday) Then
                    BatasDurasi = 240
                    'jika durasi melebihi batas durasi kasih peringatan
                    If Durasi > BatasDurasi Then
                        ListSalah = ListSalah + "Work duration : " + row.Cells(Me.ColDay.Index).Value + " ,is larger than " _
                        + (BatasDurasi / 60).ToString + " hours." + vbNewLine
                    ElseIf Durasi > 480 Then 'jika durasi lebih dari 8 jam
                        ListSalah = ListSalah + "Work duration : " + row.Cells(Me.ColDay.Index).Value.ToString + _
                        " ,Duration halfday cannot more or same with 8 hours, in weekly activity item" + vbNewLine
                    End If
                End If

                'cek break duration
                If SchBreakOut <> 0 And SchBreakIn <> 0 Then
                    If SchOut > SchIn Then
                        'jika jam normal
                        If (SchIn > SchBreakOut And SchBreakOut < SchOut) Or (SchIn > SchBreakIn And SchBreakIn < SchOut) Then
                            'istirahat tidak boleh ada diluar jam kerja
                            Me.TabControl1.SelectedIndex = 0
                            Me.dgvDailyActivities.ClearSelection()
                            Me.dgvDailyActivities.Rows.Item(row.Index).Selected = True
                            MsgBox("Work duration : " + row.Cells(Me.ColDay.Index).Value.ToString + _
                            " ,Schedule break must in work duration.", MsgBoxStyle.Information, "Information")
                            Return False
                        Else
                            If SchBreakOut > SchBreakIn Then
                                Me.TabControl1.SelectedIndex = 0
                                Me.dgvDailyActivities.ClearSelection()
                                Me.dgvDailyActivities.Rows.Item(row.Index).Selected = True
                                MsgBox("Work duration : " + row.Cells(Me.ColDay.Index).Value.ToString + _
                                " ,Schedule break out cannot more than schedule break in.", MsgBoxStyle.Information, "Information")
                                Return False
                            End If
                        End If
                    Else
                        'istirahat tidak boleh ada diluar jam kerja, jam nyebrang hari
                        If ((SchBreakOut >= SchIn And SchBreakOut < (24 * 60)) Or SchBreakOut = 0) And _
                            (SchBreakIn <= SchOut) Then
                            'kondisi benar
                        ElseIf (SchBreakIn >= SchIn And SchBreakIn < (24 * 60) Or SchBreakIn = 0) And _
                                (SchBreakOut <= SchOut) Then
                            'kondisi terbalik
                            Me.TabControl1.SelectedIndex = 0
                            Me.dgvDailyActivities.ClearSelection()
                            Me.dgvDailyActivities.Rows.Item(row.Index).Selected = True
                            MsgBox("Work duration : " + row.Cells(Me.ColDay.Index).Value.ToString + _
                            " ,Schedule break out cannot more than schedule break in.", MsgBoxStyle.Information, "Information")
                            Return False
                        ElseIf ((SchBreakIn >= SchIn And SchBreakIn < (24 * 60) Or SchBreakIn = 0) And _
                        (SchBreakOut >= SchIn And SchBreakOut < (24 * 60) Or SchBreakOut = 0)) Or _
                        (SchBreakOut <= SchOut And SchBreakIn <= SchOut) Then
                            'semua sebelah kiri atao sebelah kanan
                            If SchBreakOut > SchBreakIn Then
                                Me.TabControl1.SelectedIndex = 0
                                Me.dgvDailyActivities.ClearSelection()
                                Me.dgvDailyActivities.Rows.Item(row.Index).Selected = True
                                MsgBox("Work duration : " + row.Cells(Me.ColDay.Index).Value.ToString + _
                                " ,Schedule break out cannot more than schedule break in.", MsgBoxStyle.Information, "Information")
                                Return False
                            End If
                        ElseIf (SchIn > SchBreakOut And SchBreakOut > SchOut) And _
                        (SchIn > SchBreakIn And SchBreakIn > SchOut) Then
                            'di luar jam  kerja
                            Me.TabControl1.SelectedIndex = 0
                            Me.dgvDailyActivities.ClearSelection()
                            Me.dgvDailyActivities.Rows.Item(row.Index).Selected = True
                            MsgBox("Work duration : " + row.Cells(Me.ColDay.Index).Value.ToString + _
                            " ,Schedule break must in work duration.", MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    End If
                End If
            Next
            'jika ada baris yang salah
            Dim hasil As DialogResult
            If ListSalah <> "" Then
                hasil = MsgBox(ListSalah + vbNewLine + "Continue save data ?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = Windows.Forms.DialogResult.No Then
                    Return False
                End If
            End If



            'holiday
            ListSalah = ""
            Dim i As Integer = 1
            For Each row In Me.DgvHolidayList.Rows
                If i <> Me.DgvHolidayList.RowCount Then
                    If row.Cells(Me.ColItemHoliday_PrimaryLeave.Index).Value > 0 Then
                        If Me.ObjCuti.GetPeriodByPrimary(row.Cells(Me.ColItemHoliday_PrimaryLeave.Index).Value).Date.Year <> _
                            CDate(row.Cells(Me.ColItemHoliday_Date.Index).Value).Date.Year Then
                            Me.TabControl1.SelectedIndex = 1
                            Me.dgvDailyActivities.ClearSelection()
                            Me.dgvDailyActivities.Rows.Item(row.Index).Selected = True
                            MsgBox("Holiday list : " + row.Cells(Me.ColItemHoliday_Date.Index).Value.ToString + _
                            ", leave code period is not appropriate.", MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    End If
                    SchIn = Me.ObjCd.Convert_ToMinutes(row.Cells(Me.ColItemHoliday_SchIn.Index).Value)
                    SchBreakOut = Me.ObjCd.Convert_ToMinutes(row.Cells(Me.ColSchBreakOut_Holiday.Index).Value)
                    SchBreakIn = Me.ObjCd.Convert_ToMinutes(row.Cells(Me.ColSchBreakIn_Holiday.Index).Value)
                    SchOut = Me.ObjCd.Convert_ToMinutes(row.Cells(Me.ColItemHoliday_SchOut.Index).Value)
                    BreakDuration = Me.ObjWorkShift.HitungBreakDuration(SchIn, SchOut, SchBreakOut, SchBreakIn)

                    If String.IsNullOrEmpty(row.Cells(Me.ColItemHoliday_Date.Index).Value) = True Then
                        Me.TabControl1.SelectedIndex = 1
                        Me.dgvDailyActivities.ClearSelection()
                        Me.dgvDailyActivities.Rows.Item(row.Index).Selected = True
                        MsgBox("Holiday list : " + row.Cells(Me.ColItemHoliday_Date.Index).Value.ToString + _
                        " ,please input date", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                    If String.IsNullOrEmpty(row.Cells(Me.ColItemHoliday_Description.Index).Value) = True Then
                        Me.TabControl1.SelectedIndex = 1
                        Me.DgvHolidayList.ClearSelection()
                        Me.DgvHolidayList.Rows.Item(row.Index).Selected = True
                        MsgBox("Holiday list : " + row.Cells(Me.ColItemHoliday_Date.Index).Value.ToString + _
                        " ,please input description in holiday list", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                    'jika semua baris 00.00
                    If SchIn = 0 And _
                    SchBreakOut = 0 And _
                    SchBreakIn = 0 And _
                    SchOut = 0 Then
                        Me.TabControl1.SelectedIndex = 1
                        Me.DgvHolidayList.ClearSelection()
                        Me.DgvHolidayList.Rows.Item(row.Index).Selected = True
                        MsgBox("Holiday list : " + row.Cells(Me.ColItemHoliday_Date.Index).Value.ToString + _
                        " ,Schedule in , schedule break out , schedule break in and schedule out cannot fill 00.00.", _
                        MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                    'cari durasinya
                    If SchIn < SchOut Then
                        'hitung schOut - schIn
                        Durasi = (SchOut - SchIn) - BreakDuration
                    ElseIf SchIn > SchOut Then
                        'hitung (24.00-schIn)+(schOut-00.00)
                        Durasi = (((24 * 60) - SchIn) + SchOut) - BreakDuration
                    ElseIf SchIn = SchOut Then
                        Me.TabControl1.SelectedIndex = 0
                        Me.dgvDailyActivities.ClearSelection()
                        Me.dgvDailyActivities.Rows.Item(row.Index).Selected = True
                        MsgBox("Holiday list : " + row.Cells(Me.ColItemHoliday_Date.Index).Value.ToString + _
                        " ,Schedule in not be the same with schedule out.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If

                    'cek duration, jika fullday batasan 8 jam, jika halfday batasan 4 jam.
                    'dalam menit
                    CboDgv = row.Cells(Me.ColItemHoliday_Duration.Index)
                    If row.Cells(Me.ColItemHoliday_Duration.Index).Value = CboDgv.Items.Item(ClsMasterWorkShift.enumDuration.fullday) Then
                        BatasDurasi = 480
                        If Durasi > BatasDurasi Then
                            ListSalah = ListSalah + "Holiday list : " + row.Cells(Me.ColItemHoliday_Date.Index).Value.ToString + " ,is larger than " _
                            + (BatasDurasi / 60).ToString + " hours." + vbNewLine
                        End If
                    ElseIf row.Cells(Me.ColItemHoliday_Duration.Index).Value = CboDgv.Items.Item(ClsMasterWorkShift.enumDuration.halfday) Then
                        BatasDurasi = 240
                        If Durasi > BatasDurasi Then
                            ListSalah = ListSalah + "Holiday list : " + row.Cells(Me.ColItemHoliday_Date.Index).Value.ToString + " ,is larger than " _
                            + (BatasDurasi / 60).ToString + " hours." + vbNewLine
                        ElseIf Durasi > 480 Then 'jika durasi lebih dari 8 jam
                            ListSalah = ListSalah + "Holiday list : " + row.Cells(Me.ColItemHoliday_Date.Index).Value.ToString _
                            + " ,Duration halfday cannot more or same with 8 hours."
                        End If
                    End If
                End If
                If SchBreakOut <> 0 And SchBreakIn <> 0 Then
                    'cek break duration
                    If SchOut > SchIn Then
                        'jika jam normal
                        If (SchIn > SchBreakOut And SchBreakOut < SchOut) Or (SchIn > SchBreakIn And SchBreakIn < SchOut) Then
                            'istirahat tidak boleh ada diluar jam kerja
                            Me.TabControl1.SelectedIndex = 1
                            Me.DgvHolidayList.ClearSelection()
                            Me.DgvHolidayList.Rows.Item(row.Index).Selected = True
                            MsgBox("Holiday list : " + row.Cells(Me.ColItemHoliday_Date.Index).Value.ToString + _
                            " ,Schedule break must in work duration.", MsgBoxStyle.Information, "Information")
                            Return False
                        Else
                            If SchBreakOut > SchBreakIn Then
                                Me.TabControl1.SelectedIndex = 1
                                Me.DgvHolidayList.ClearSelection()
                                Me.DgvHolidayList.Rows.Item(row.Index).Selected = True
                                MsgBox("Holiday list : " + row.Cells(Me.ColItemHoliday_Date.Index).Value.ToString + _
                                " ,Schedule break out cannot more than schedule break in.", MsgBoxStyle.Information, "Information")
                                Return False
                            End If
                        End If
                    Else
                        'istirahat tidak boleh ada diluar jam kerja, jam nyebrang hari
                        If ((SchBreakOut >= SchIn And SchBreakOut < (24 * 60)) Or SchBreakOut = 0) And _
                            (SchBreakIn <= SchOut) Then
                            'kondisi benar
                        ElseIf (SchBreakIn >= SchIn And SchBreakIn < (24 * 60) Or SchBreakIn = 0) And _
                                (SchBreakOut <= SchOut) Then
                            'kondisi terbalik
                            Me.TabControl1.SelectedIndex = 1
                            Me.DgvHolidayList.ClearSelection()
                            Me.DgvHolidayList.Rows.Item(row.Index).Selected = True
                            MsgBox("Holiday list : " + row.Cells(Me.ColItemHoliday_Date.Index).Value.ToString + _
                            " ,Schedule break out cannot more than schedule break in.", MsgBoxStyle.Information, "Information")
                            Return False
                        ElseIf ((SchBreakIn >= SchIn And SchBreakIn < (24 * 60) Or SchBreakIn = 0) And _
                        (SchBreakOut >= SchIn And SchBreakOut < (24 * 60) Or SchBreakOut = 0)) Or _
                        (SchBreakOut <= SchOut And SchBreakIn <= SchOut) Then
                            'semua sebelah kiri atao sebelah kanan
                            If SchBreakOut > SchBreakIn Then
                                Me.TabControl1.SelectedIndex = 1
                                Me.DgvHolidayList.ClearSelection()
                                Me.DgvHolidayList.Rows.Item(row.Index).Selected = True
                                MsgBox("Holiday list : " + row.Cells(Me.ColItemHoliday_Date.Index).Value.ToString + _
                                " ,Schedule break out cannot more than schedule break in.", MsgBoxStyle.Information, "Information")
                                Return False
                            End If
                        ElseIf (SchIn > SchBreakOut And SchBreakOut > SchOut) And _
                        (SchIn > SchBreakIn And SchBreakIn > SchOut) Then
                            'di luar jam  kerja
                            Me.TabControl1.SelectedIndex = 1
                            Me.DgvHolidayList.ClearSelection()
                            Me.DgvHolidayList.Rows.Item(row.Index).Selected = True
                            MsgBox("Holiday list : " + row.Cells(Me.ColItemHoliday_Date.Index).Value.ToString + _
                            " ,Schedule break must in work duration.", MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    End If
                End If
                i = i + 1
            Next

            If ListSalah <> "" Then
                'tambahan dikit #byahdar
                Dim ArryList As String() = ListSalah.Split(vbNewLine)
                If ArryList.Length > 10 Then
                    ListSalah = ""
                    For ij As Integer = 1 To 10
                        If ListSalah = "" Then
                            ListSalah = ArryList.GetValue(ij)
                        Else
                            ListSalah &= vbNewLine & ArryList.GetValue(ij)
                        End If
                    Next
                    ListSalah &= vbNewLine & "Any more."
                End If
                hasil = MsgBox(ListSalah + vbNewLine + "Continue save data ?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = Windows.Forms.DialogResult.No Then
                    Return False
                End If
            End If

            If AttParametersCheck(ClsMasterWorkShift.EnumCheckRange.AllRange) = False Then
                Return False
            End If

            If Me.cboRoundingTypeEarlyOT.SelectedIndex <> Core_Module.ClsMasterWorkShift.enumRoundingType.No_Rounding AndAlso _
           Me.ObjCd.Convert_ToMinutes(Me.txtRoundingEarlyOT.Text) = 0 Then
                MsgBox("Rounding type for early overtime : " & cboRoundingTypeEarlyOT.Text & _
                       ", rounding value for early overtime cannot fill 00.00.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            If Me.cboRoundingTypeLate.SelectedIndex <> Core_Module.ClsMasterWorkShift.enumRoundingType.No_Rounding AndAlso _
           Me.ObjCd.Convert_ToMinutes(Me.txtRoundingLate.Text) = 0 Then
                MsgBox("Rounding type for late : " & cboRoundingTypeLate.Text & _
                       ", rounding value for late cannot fill 00.00.", MsgBoxStyle.Information, "Information")
                Return False

            End If

            If Me.cboRoundingTypeEarlyOut.SelectedIndex <> Core_Module.ClsMasterWorkShift.enumRoundingType.No_Rounding AndAlso _
           Me.ObjCd.Convert_ToMinutes(Me.txtRoundingEarlyOut.Text) = 0 Then
                MsgBox("Rounding type for early out : " & cboRoundingTypeEarlyOut.Text & _
                      ", rounding value for early out cannot fill 00.00.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            If Me.cboRoundingTypeOT.SelectedIndex <> Core_Module.ClsMasterWorkShift.enumRoundingType.No_Rounding AndAlso _
           Me.ObjCd.Convert_ToMinutes(Me.txtRoundingOT.Text) = 0 Then
                MsgBox("Rounding type for overtime : " & cboRoundingTypeOT.Text & _
                     ", rounding value for overtime cannot fill 00.00.", MsgBoxStyle.Information, "Information")
                Return False
            End If


            If Me.ObjCd.Convert_ToMinutes(Me.txtCalculateAfterEarlyOT.Text) > 0 AndAlso _
               Me.ObjCd.Convert_ToMinutes(Me.txtMaxValueEarlyOT.Text) > 0 AndAlso _
               Me.ObjCd.Convert_ToMinutes(Me.txtCalculateAfterEarlyOT.Text) >= Me.ObjCd.Convert_ToMinutes(Me.txtMaxValueEarlyOT.Text) Then
                MsgBox("cannot fill calculate after early ot >= maximum value early ot", MsgBoxStyle.Information, "Information")
                Return False
            End If

            If Me.ObjCd.Convert_ToMinutes(Me.txtCalculateAfterLate.Text) > 0 AndAlso _
            Me.ObjCd.Convert_ToMinutes(Me.txtMaxValueLate.Text) > 0 AndAlso _
            Me.ObjCd.Convert_ToMinutes(Me.txtCalculateAfterLate.Text) >= Me.ObjCd.Convert_ToMinutes(Me.txtMaxValueLate.Text) Then
                MsgBox("cannot fill calculate after late >= maximum value late", MsgBoxStyle.Information, "Information")
                Return False
            End If

            If Me.ObjCd.Convert_ToMinutes(Me.txtCalculateAfterEarlyOut.Text) > 0 AndAlso _
               Me.ObjCd.Convert_ToMinutes(Me.txtMaxValueEarlyOut.Text) > 0 AndAlso _
            Me.ObjCd.Convert_ToMinutes(Me.txtCalculateAfterEarlyOut.Text) >= Me.ObjCd.Convert_ToMinutes(Me.txtMaxValueEarlyOut.Text) Then
                MsgBox("cannot fill calculate after early out >= maximum value early out", MsgBoxStyle.Information, "Information")
                Return False
            End If

            If Me.ObjCd.Convert_ToMinutes(Me.txtCalculateAfterOT.Text) > 0 AndAlso _
               Me.ObjCd.Convert_ToMinutes(Me.txtMaxValueOT.Text) > 0 AndAlso _
               Me.ObjCd.Convert_ToMinutes(Me.txtCalculateAfterOT.Text) >= Me.ObjCd.Convert_ToMinutes(Me.txtMaxValueOT.Text) Then
                MsgBox("cannot fill calculate after overtime >= maximum value overtime ", MsgBoxStyle.Information, "Information")
                Return False
            End If




            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub KosongkanField()
        Try
            Me.ObjCd.KosongkanControl(Me.gboDetail.Controls, True, False, False)
            Me.ObjCd.KosongkanControl(Me.TabPage3.Controls, False, False, True)
            Me.chkCalculateEarlyOTTolerance.Checked = True
            Me.chkCalculateLateTolerance.Checked = True
            Me.chkCalculateEarlyOutTolerance.Checked = True
            Me.chkCalculateOTTolerance.Checked = True
            Me.dgvDailyActivities.Rows.Clear()
            Me.DgvHolidayList.Rows.Clear()
            Me.DgvAttParams.Rows.Clear()
            Me.DgvAttParams.Rows.Clear()
            Me.txtCalculateAfterEarlyOT.Text = Me.ObjCd.Convert_ToString(0)
            Me.txtCalculateAfterLate.Text = Me.ObjCd.Convert_ToString(0)
            Me.txtCalculateAfterEarlyOut.Text = Me.ObjCd.Convert_ToString(0)
            Me.txtCalculateAfterOT.Text = Me.ObjCd.Convert_ToString(0)
            Me.txtRoundingEarlyOT.Text = Me.ObjCd.Convert_ToString(0)
            Me.txtRoundingLate.Text = Me.ObjCd.Convert_ToString(0)
            Me.txtRoundingEarlyOut.Text = Me.ObjCd.Convert_ToString(0)
            Me.txtRoundingOT.Text = Me.ObjCd.Convert_ToString(0)
            Me.txtMaxValueEarlyOT.Text = Me.ObjCd.Convert_ToString(0)
            Me.txtMaxValueLate.Text = Me.ObjCd.Convert_ToString(0)
            Me.txtMaxValueEarlyOut.Text = Me.ObjCd.Convert_ToString(0)
            Me.txtMaxValueOT.Text = Me.ObjCd.Convert_ToString(0)
            Me.TxtPrimaryJob.Text = -1
            Me.chkUseOvertimeOrder.Checked = False
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AktifkanDetail()
        Try
            Me.ObjCd.KendalikanControl(Me.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCd.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCd.KendalikanControl(Me.TabPage3.Controls, False, False, False, False, False, False, True, True, True, True)
            Me.DgvHolidayList.AllowUserToAddRows = True
            Me.DgvHolidayList.AllowUserToDeleteRows = True
            Me.DgvAttParams.ReadOnly = False
            Me.LinkLabel1.Enabled = True
            Me.CancelButton = Me.btnCancel

            If Me.CboIsCalculateEarlyOT.SelectedIndex = 0 Then
                Me.txtCalculateAfterEarlyOT.Enabled = True
                Me.cboRoundingTypeEarlyOT.Enabled = True
                If Me.cboRoundingTypeEarlyOT.SelectedIndex = Core_Module.ClsMasterWorkShift.enumRoundingType.No_Rounding Then
                    Me.txtRoundingEarlyOT.Enabled = False
                Else
                    Me.txtRoundingEarlyOT.Enabled = True
                End If
                Me.txtMaxValueEarlyOT.Enabled = True
                Me.chkCalculateEarlyOTTolerance.Enabled = True
            Else
                Me.txtCalculateAfterEarlyOT.Enabled = False
                Me.cboRoundingTypeEarlyOT.Enabled = False
                Me.txtRoundingEarlyOT.Enabled = False
                Me.txtMaxValueEarlyOT.Enabled = False
            End If


            If Me.CboIsCalculateLate.SelectedIndex = 0 Then
                Me.txtCalculateAfterLate.Enabled = True
                Me.cboRoundingTypeLate.Enabled = True
                If Me.cboRoundingTypeLate.SelectedIndex = Core_Module.ClsMasterWorkShift.enumRoundingType.No_Rounding Then
                    Me.txtRoundingLate.Enabled = False
                Else
                    Me.txtRoundingLate.Enabled = True
                End If
                Me.txtMaxValueLate.Enabled = True
                Me.chkCalculateLateTolerance.Enabled = True
            Else
                Me.txtCalculateAfterLate.Enabled = False
                Me.cboRoundingTypeLate.Enabled = False
                Me.txtRoundingLate.Enabled = False
                Me.txtMaxValueLate.Enabled = False
            End If

            If Me.CboIsCalculateEarlyOut.SelectedIndex = 0 Then
                Me.txtCalculateAfterEarlyOut.Enabled = True
                Me.cboRoundingTypeEarlyOut.Enabled = True
                If Me.cboRoundingTypeEarlyOut.SelectedIndex = Core_Module.ClsMasterWorkShift.enumRoundingType.No_Rounding Then
                    Me.txtRoundingEarlyOut.Enabled = False
                Else
                    Me.txtRoundingEarlyOut.Enabled = True
                End If
                Me.txtMaxValueEarlyOut.Enabled = True
                Me.chkCalculateEarlyOutTolerance.Enabled = True
            Else
                Me.txtCalculateAfterEarlyOut.Enabled = False
                Me.cboRoundingTypeEarlyOut.Enabled = False
                Me.txtRoundingEarlyOut.Enabled = False
                Me.txtMaxValueEarlyOut.Enabled = False
            End If

            If Me.CboIsCalculateOT.SelectedIndex = 0 Then
                Me.txtCalculateAfterOT.Enabled = True
                Me.cboRoundingTypeOT.Enabled = True
                If Me.cboRoundingTypeOT.SelectedIndex = Core_Module.ClsMasterWorkShift.enumRoundingType.No_Rounding Then
                    Me.txtRoundingOT.Enabled = False
                Else
                    Me.txtRoundingOT.Enabled = True
                End If
                Me.txtMaxValueOT.Enabled = True
                Me.chkCalculateOTTolerance.Enabled = True
            Else
                Me.txtCalculateAfterOT.Enabled = False
                Me.cboRoundingTypeOT.Enabled = False
                Me.txtRoundingOT.Enabled = False
                Me.txtMaxValueOT.Enabled = False
            End If
            If Me.chkUseOvertimeOrder.Checked = False Then
                Me.chkSendNotifEmailOvertimeOrder.Checked = False
                Me.chkSendNotifEmailOvertimeOrder.Enabled = False
            Else
                Me.chkSendNotifEmailOvertimeOrder.Enabled = True
            End If

            If Me.CboIsCalculateEarlyOT.SelectedIndex = 1 And Me.CboIsCalculateOT.SelectedIndex = 1 Then
                Me.chkUseOvertimeOrder.Enabled = False
                Me.chkSendNotifEmailOvertimeOrder.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            Me.ObjCd.KendalikanControl(Me.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCd.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCd.KendalikanControl(Me.TabPage3.Controls, False, False, False, False, False, False, True, False, True, False)
            Me.DgvHolidayList.AllowUserToAddRows = False
            Me.DgvHolidayList.AllowUserToDeleteRows = False
            Me.DgvAttParams.ReadOnly = True
            Me.LinkLabel1.Enabled = False
            Me.CancelButton = Me.btnExit

            Me.txtCalculateAfterEarlyOT.Enabled = False
            Me.cboRoundingTypeEarlyOT.Enabled = False
            Me.txtRoundingEarlyOT.Enabled = False
            Me.txtMaxValueEarlyOT.Enabled = False

            Me.txtCalculateAfterLate.Enabled = False
            Me.cboRoundingTypeLate.Enabled = False
            Me.txtRoundingLate.Enabled = False
            Me.txtMaxValueLate.Enabled = False


            Me.txtCalculateAfterEarlyOut.Enabled = False
            Me.cboRoundingTypeEarlyOut.Enabled = False
            Me.txtRoundingEarlyOut.Enabled = False
            Me.txtMaxValueEarlyOut.Enabled = False


            Me.txtCalculateAfterOT.Enabled = False
            Me.cboRoundingTypeOT.Enabled = False
            Me.txtRoundingOT.Enabled = False
            Me.txtMaxValueOT.Enabled = False

            Me.chkCalculateEarlyOTTolerance.Enabled = False
            Me.chkCalculateLateTolerance.Enabled = False
            Me.chkCalculateEarlyOutTolerance.Enabled = False
            Me.chkCalculateOTTolerance.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.ObjWorkShift
                Me.TipeWorkshift = .TipeWorkshift
                Me.TxtPrimaryJob.Text = .Primary
                Me.TxtJobCode.Text = .Code
                Me.TxtJobName.Text = .Desc
                Me.TxtRemark.Text = .Remarks
                If .IsCalculateEarlyOT = 0 Then
                    Me.CboIsCalculateEarlyOT.SelectedIndex = 1
                    Me.txtCalculateAfterEarlyOT.Text = Me.ObjCd.Convert_ToString(0)
                    Me.cboRoundingTypeEarlyOT.SelectedIndex = 0
                    Me.txtRoundingEarlyOT.Text = Me.ObjCd.Convert_ToString(0)
                    Me.txtMaxValueEarlyOT.Text = Me.ObjCd.Convert_ToString(0)
                    'Me.txtCalculateAfterEarlyOT.Enabled = False
                    'Me.cboRoundingTypeEarlyOT.Enabled = False
                    'Me.txtRoundingEarlyOT.Enabled = False
                    'Me.txtMaxValueEarlyOT.Enabled = False
                Else
                    Me.CboIsCalculateEarlyOT.SelectedIndex = 0
                    Me.txtCalculateAfterEarlyOT.Text = Me.ObjCd.Convert_ToString(.CalculateAfterEarlyOT)
                    Me.cboRoundingTypeEarlyOT.SelectedIndex = .RoundingTypeEarlyOT
                    Me.txtRoundingEarlyOT.Text = Me.ObjCd.Convert_ToString(.RoundingValueEarlyOT)
                    Me.txtMaxValueEarlyOT.Text = Me.ObjCd.Convert_ToString(.MaxValueEarlyOT)

                    'Me.txtCalculateAfterEarlyOT.Enabled = True
                    'Me.cboRoundingTypeEarlyOT.Enabled = True
                    'Me.txtRoundingEarlyOT.Enabled = True
                    'Me.txtMaxValueEarlyOT.Enabled = True
                End If

                If .IsCalculateLate = 0 Then
                    Me.CboIsCalculateLate.SelectedIndex = 1
                    Me.txtCalculateAfterLate.Text = Me.ObjCd.Convert_ToString(0)
                    Me.cboRoundingTypeLate.SelectedIndex = 0
                    Me.txtRoundingLate.Text = Me.ObjCd.Convert_ToString(0)
                    Me.txtMaxValueLate.Text = Me.ObjCd.Convert_ToString(0)
                    'Me.txtCalculateAfterLate.Enabled = False
                    'Me.cboRoundingTypeLate.Enabled = False
                    'Me.txtRoundingLate.Enabled = False
                    'Me.txtMaxValueLate.Enabled = False
                Else
                    Me.CboIsCalculateLate.SelectedIndex = 0
                    Me.txtCalculateAfterLate.Text = Me.ObjCd.Convert_ToString(.CalculateAfterLate)
                    Me.cboRoundingTypeLate.SelectedIndex = .RoundingTypeLate
                    Me.txtRoundingLate.Text = Me.ObjCd.Convert_ToString(.RoundingValueLate)
                    Me.txtMaxValueLate.Text = Me.ObjCd.Convert_ToString(.MaxValueLate)
                    'Me.txtCalculateAfterLate.Enabled = True
                    'Me.cboRoundingTypeLate.Enabled = True
                    'Me.txtRoundingLate.Enabled = True
                    'Me.txtMaxValueLate.Enabled = True
                End If

                If .IsCalculateEarlyOut = 0 Then
                    Me.CboIsCalculateEarlyOut.SelectedIndex = 1
                    Me.txtCalculateAfterEarlyOut.Text = Me.ObjCd.Convert_ToString(0)
                    Me.cboRoundingTypeEarlyOut.SelectedIndex = 0
                    Me.txtRoundingEarlyOut.Text = Me.ObjCd.Convert_ToString(0)
                    Me.txtMaxValueEarlyOut.Text = Me.ObjCd.Convert_ToString(0)
                    'Me.txtCalculateAfterEarlyOut.Enabled = False
                    'Me.cboRoundingTypeEarlyOut.Enabled = False
                    'Me.txtRoundingEarlyOut.Enabled = False
                    'Me.txtMaxValueEarlyOut.Enabled = False
                Else
                    Me.CboIsCalculateEarlyOut.SelectedIndex = 0
                    Me.txtCalculateAfterEarlyOut.Text = Me.ObjCd.Convert_ToString(.CalculateAfterEarlyOut)
                    Me.cboRoundingTypeEarlyOut.SelectedIndex = .RoundingTypeEarlyOut
                    Me.txtRoundingEarlyOut.Text = Me.ObjCd.Convert_ToString(.RoundingValueEarlyOut)
                    Me.txtMaxValueEarlyOut.Text = Me.ObjCd.Convert_ToString(.MaxValueEarlyOut)
                    'Me.txtCalculateAfterEarlyOut.Enabled = True
                    'Me.cboRoundingTypeEarlyOut.Enabled = True
                    'Me.txtRoundingEarlyOut.Enabled = True
                    'Me.txtMaxValueEarlyOut.Enabled = True
                End If

                If .IsCalculateOT = 0 Then
                    Me.CboIsCalculateOT.SelectedIndex = 1
                    Me.txtCalculateAfterOT.Text = Me.ObjCd.Convert_ToString(0)
                    Me.cboRoundingTypeOT.SelectedIndex = 0
                    Me.txtRoundingOT.Text = Me.ObjCd.Convert_ToString(0)
                    Me.txtMaxValueOT.Text = Me.ObjCd.Convert_ToString(0)
                    'Me.txtCalculateAfterOT.Enabled = False
                    'Me.cboRoundingTypeOT.Enabled = False
                    'Me.txtRoundingOT.Enabled = False
                    'Me.txtMaxValueOT.Enabled = False
                Else
                    Me.CboIsCalculateOT.SelectedIndex = 0
                    Me.txtCalculateAfterOT.Text = Me.ObjCd.Convert_ToString(.CalculateAfterOT)
                    Me.cboRoundingTypeOT.SelectedIndex = .RoundingTypeOT
                    Me.txtRoundingOT.Text = Me.ObjCd.Convert_ToString(.RoundingValueOT)
                    Me.txtMaxValueOT.Text = Me.ObjCd.Convert_ToString(.MaxValueOT)


                    'Me.txtCalculateAfterOT.Enabled = True
                    'Me.cboRoundingTypeOT.Enabled = True
                    'Me.txtRoundingOT.Enabled = True
                    'Me.txtMaxValueOT.Enabled = True
                End If

                If .CalculateEarlyOTTolerance = 1 Then
                    Me.chkCalculateEarlyOTTolerance.Checked = True
                Else
                    Me.chkCalculateEarlyOTTolerance.Checked = False
                End If
                If .CalculateLateTolerance = 1 Then
                    Me.chkCalculateLateTolerance.Checked = True
                Else
                    Me.chkCalculateLateTolerance.Checked = False
                End If
                If .CalculateEarlyOutTolerance = 1 Then
                    Me.chkCalculateEarlyOutTolerance.Checked = True
                Else
                    Me.chkCalculateEarlyOutTolerance.Checked = False
                End If
                If .CalculateOTTolerance = 1 Then
                    Me.chkCalculateOTTolerance.Checked = True
                Else
                    Me.chkCalculateOTTolerance.Checked = False
                End If
                If .UseOtOrder = 1 Then
                    Me.chkUseOvertimeOrder.Checked = True
                Else
                    Me.chkUseOvertimeOrder.Checked = False
                End If
                If .NotificationEmailOt = 1 Then
                    Me.chkSendNotifEmailOvertimeOrder.Checked = True
                Else
                    Me.chkSendNotifEmailOvertimeOrder.Checked = False
                End If
                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SalinKeObject()
        Try
            With Me.ObjWorkShift
                .TipeWorkshift = Me.TipeWorkshift
                .Code = Me.TxtJobCode.Text
                .Desc = Me.TxtJobName.Text
                .Remarks = Me.TxtRemark.Text

                If Me.CboIsCalculateEarlyOT.SelectedIndex = 0 Then
                    .IsCalculateEarlyOT = 1
                    .CalculateAfterEarlyOT = Me.ObjCd.Convert_ToMinutes(Me.txtCalculateAfterEarlyOT.Text)
                    .RoundingTypeEarlyOT = Me.cboRoundingTypeEarlyOT.SelectedIndex
                    .RoundingValueEarlyOT = Me.ObjCd.Convert_ToMinutes(Me.txtRoundingEarlyOT.Text)
                    .MaxValueEarlyOT = Me.ObjCd.Convert_ToMinutes(Me.txtMaxValueEarlyOT.Text)
                Else
                    .IsCalculateEarlyOT = 0
                    .CalculateAfterEarlyOT = 0
                    .RoundingTypeEarlyOT = 0
                    .RoundingValueEarlyOT = 0
                    .MaxValueEarlyOT = 0
                End If


                If Me.CboIsCalculateLate.SelectedIndex = 0 Then
                    .IsCalculateLate = 1
                    .CalculateAfterLate = Me.ObjCd.Convert_ToMinutes(Me.txtCalculateAfterLate.Text)
                    .RoundingTypeLate = Me.cboRoundingTypeLate.SelectedIndex
                    .RoundingValueLate = Me.ObjCd.Convert_ToMinutes(Me.txtRoundingLate.Text)
                    .MaxValueLate = Me.ObjCd.Convert_ToMinutes(Me.txtMaxValueLate.Text)
                Else
                    .IsCalculateLate = 0
                    .CalculateAfterLate = 0
                    .RoundingTypeLate = 0
                    .RoundingValueLate = 0
                    .MaxValueLate = 0
                End If


                If Me.CboIsCalculateEarlyOut.SelectedIndex = 0 Then
                    .IsCalculateEarlyOut = 1
                    .CalculateAfterEarlyOut = Me.ObjCd.Convert_ToMinutes(Me.txtCalculateAfterEarlyOut.Text)
                    .RoundingTypeEarlyOut = Me.cboRoundingTypeEarlyOut.SelectedIndex
                    .RoundingValueEarlyOut = Me.ObjCd.Convert_ToMinutes(Me.txtRoundingEarlyOut.Text)
                    .MaxValueEarlyOut = Me.ObjCd.Convert_ToMinutes(Me.txtMaxValueEarlyOut.Text)
                Else
                    .IsCalculateEarlyOut = 0
                    .CalculateAfterEarlyOut = 0
                    .RoundingTypeEarlyOut = 0
                    .RoundingValueEarlyOut = 0
                    .MaxValueEarlyOut = 0
                End If


                If Me.CboIsCalculateOT.SelectedIndex = 0 Then
                    .IsCalculateOT = 1
                    .CalculateAfterOT = Me.ObjCd.Convert_ToMinutes(Me.txtCalculateAfterOT.Text)
                    .RoundingTypeOT = Me.cboRoundingTypeOT.SelectedIndex
                    .RoundingValueOT = Me.ObjCd.Convert_ToMinutes(Me.txtRoundingOT.Text)
                    .MaxValueOT = Me.ObjCd.Convert_ToMinutes(Me.txtMaxValueOT.Text)
                Else
                    .IsCalculateOT = 0
                    .CalculateAfterOT = 0
                    .RoundingTypeOT = 0
                    .RoundingValueOT = 0
                    .MaxValueOT = 0
                End If

                If Me.chkCalculateEarlyOTTolerance.Checked = True Then
                    .CalculateEarlyOTTolerance = 1
                Else
                    .CalculateEarlyOTTolerance = 0
                End If
                If Me.chkCalculateLateTolerance.Checked = True Then
                    .CalculateLateTolerance = 1
                Else
                    .CalculateLateTolerance = 0
                End If
                If Me.chkCalculateEarlyOutTolerance.Checked = True Then
                    .CalculateEarlyOutTolerance = 1
                Else
                    .CalculateEarlyOutTolerance = 0
                End If
                If Me.chkCalculateOTTolerance.Checked = True Then
                    .CalculateOTTolerance = 1
                Else
                    .CalculateOTTolerance = 0
                End If
                If Me.chkUseOvertimeOrder.Checked = True Then
                    .UseOtOrder = 1
                Else
                    .UseOtOrder = 0
                End If
                If Me.chkSendNotifEmailOvertimeOrder.Checked = True Then
                    .NotificationEmailOt = 1
                Else
                    .NotificationEmailOt = 0
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Dim FrmCopyData As New FrmMasterWorkshift_CopyData
        Dim frmcari As New FrmMasterWorkShift_Cari(Me.TipeWorkshift)
        Try
            frmcari.BackColorAll = Me.BackColorAll
            frmcari.ShowDialog()
            If frmcari.PrimaryWorkshift <> -1 Then
                FrmCopyData.tipeWorkshift = Me.TipeWorkshift
                FrmCopyData.ShowDialog()
                If FrmCopyData.Batal = False Then
                    Me.ObjWorkShiftItem.IsiDGVItem(Me.dgvDailyActivities, Me.DgvHolidayList, frmcari.PrimaryWorkshift, Me, Me.ObjCd, FrmCopyData.pilihan, Me.TipeWorkshift)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FrmCopyData.Dispose()
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            Me.baru = True
            Me.KosongkanField()
            Me.AktifkanDetail()
            Me.SetDgvWeekly()
            Me.SetDgvAttSettings()
            Me.TxtJobCode.Focus()
            Me.tslCreatedBy.Text = "Created By: " + UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterWorkShift_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.TxtJobCode.Enabled = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObject()
                        If baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjWorkShift.SimpanBaru(Me.UserCode, Me.dgvDailyActivities, Me.DgvHolidayList, Me) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnFind.Focus()
                            Else
                                e.Cancel = True
                            End If
                        Else
                            'simpan perubahan
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjWorkShift.SimpanPerubahan(Me.UserCode, Me.dgvDailyActivities, Me.DgvHolidayList, Me) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnEdit.Focus()
                            Else
                                e.Cancel = True
                            End If
                        End If
                    Else
                        e.Cancel = True
                    End If
                ElseIf hasil = MsgBoxResult.Cancel Then
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterWorkShift_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.MinimizeBox = True
            If Me.TipeWorkshift = ClsMasterWorkShift.EnumTipeWorkshift.Production Then
                Me.Text = "Production Workshift"
            Else
                Me.Text = "HR Workshift"
            End If
            Me.NonAktifkanDetail()
            Me.dgvDailyActivities.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvHolidayList.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvAttParams.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dtaAtt.Connection.ConnectionString = Core_Module.GetConstring
            If Me.TipeWorkshift = ClsMasterWorkShift.EnumTipeWorkshift.Production Then
                Me.TabControl1.TabPages.Remove(Me.TabPage3)
            End If

            'WAJIB SINKRON 2019-03-02
            Me.ColItemHoliday_Date.Frozen = True
            Me.ColDay.Frozen = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            'cek data kosong
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObject()
                hasil = MsgBox("Save data ?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                    If baru = True Then
                        If Me.ObjWorkShift.SimpanBaru(Me.UserCode, Me.dgvDailyActivities, Me.DgvHolidayList, Me) = True Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                        End If
                    Else
                        If Me.ObjWorkShift.SimpanPerubahan(Me.UserCode, Me.dgvDailyActivities, Me.DgvHolidayList, Me) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnEdit.Focus()
                        End If
                    End If
                    Windows.Forms.Cursor.Current = Cursors.Default
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtJobCode.Text <> "" And Me.TxtJobCode.Enabled = True Then
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If baru = True Then
                        Me.NonAktifkanDetail()
                        Me.KosongkanField()
                    Else
                        Me.ObjWorkShift.CariDetail(Me.TxtPrimaryJob.Text, Me, Me.dgvDailyActivities, Me.DgvHolidayList, Me.ObjCd)
                        Me.NonAktifkanDetail()
                        Me.SalinKeField()
                    End If
                Else
                    Exit Sub
                End If
            Else
                Me.NonAktifkanDetail()
                Me.KosongkanField()
            End If
            Me.btnNew.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            If Me.TxtJobCode.Text = "" Then
                Me.btnFind.PerformClick()
            End If
            If Me.TxtJobCode.Text <> "" Then
                Me.baru = False
                Me.ObjWorkShift.CariDetail(Me.TxtPrimaryJob.Text, Me, Me.dgvDailyActivities, Me.DgvHolidayList, Me.ObjCd)
                Me.AktifkanDetail()
                Me.TxtJobCode.Focus()

                If Me.DgvAttParams.RowCount = 0 Then
                    SetDgvAttSettings()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtJobCode.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtJobCode.Text <> "" Then
                If Me.ObjWorkShift.ApakahBolehHapus(Me.TxtPrimaryJob.Text) = True Then
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.ObjWorkShift.Hapus(Me.TxtPrimaryJob.Text, Me.UserCode) = True Then
                            Me.KosongkanField()
                            Me.btnNew.Focus()
                            MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub btnFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim hj As New FrmMasterWorkShift_Cari(Me.TipeWorkshift)
        Try
            hj.BackColorAll = Me.BackColorAll
            hj.ShowDialog()
            If Not hj.PrimaryWorkshift = -1 Then
                If Me.ObjWorkShift.CariDetail(hj.PrimaryWorkshift, Me, Me.dgvDailyActivities, Me.DgvHolidayList, Me.ObjCd) = True Then
                    Me.SalinKeField()
                    Me.btnEdit.Focus()
                    'Me.NonAktifkanDetail()
                Else
                    Me.btnFind.Focus()
                End If
                Me.btnFind.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub dgvDailyActivities_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvDailyActivities.CellBeginEdit
        'Dim CboDgv As DataGridViewComboBoxCell
        Try
            If Me.TxtJobCode.Enabled = True Then
                If e.ColumnIndex = Me.ColWorkdays.Index Or e.ColumnIndex = Me.ColOvertime.Index Or _
                    e.ColumnIndex = Me.ColSchIn.Index Or e.ColumnIndex = Me.ColSchOut.Index Or _
                    e.ColumnIndex = Me.ColDuration.Index Or e.ColumnIndex = Me.ColRemarks.Index Or _
                    e.ColumnIndex = Me.ColSchBreakIn_WA.Index Or e.ColumnIndex = Me.ColSchBreakOut_WA.Index Then
                    Me.buffer = Me.dgvDailyActivities.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                e.Cancel = True
            End If

            ''jika holiday
            'CboDgv = Me.dgvDailyActivities.Item(Me.ColWorkdays.Index, e.RowIndex)
            'If Me.dgvDailyActivities.Item(Me.ColWorkdays.Index, e.RowIndex).Value = _
            'CboDgv.Items.Item(ClsMasterWorkShift.TipeWorkdays.Holiday) Then
            '    'cek lagi durasinya
            '    CboDgv = Me.dgvDailyActivities.Item(Me.ColDuration.Index, e.RowIndex)
            '    If Me.dgvDailyActivities.Item(Me.ColDuration.Index, e.RowIndex).Value = _
            '    CboDgv.Items.Item(ClsMasterWorkShift.enumDuration.fullday) Then
            '        'fullday
            '        If e.ColumnIndex = Me.ColWorkdays.Index Or e.ColumnIndex = Me.ColOvertime.Index Or _
            '            e.ColumnIndex = Me.ColDuration.Index Or e.ColumnIndex = Me.ColRemarks.Index Then
            '            Me.buffer = Me.dgvDailyActivities.Item(e.ColumnIndex, e.RowIndex).Value
            '        Else
            '            e.Cancel = True
            '        End If
            '    Else
            '        'halfday
            '        If e.ColumnIndex = Me.ColWorkdays.Index Or e.ColumnIndex = Me.ColOvertime.Index Or _
            '            e.ColumnIndex = Me.ColSchIn.Index Or e.ColumnIndex = Me.ColSchOut.Index Or _
            '            e.ColumnIndex = Me.ColDuration.Index Or e.ColumnIndex = Me.ColRemarks.Index Then
            '            Me.buffer = Me.dgvDailyActivities.Item(e.ColumnIndex, e.RowIndex).Value
            '        Else
            '            e.Cancel = True
            '        End If
            '    End If
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvDailyActivities_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvDailyActivities.CellEndEdit
        Dim row As DataGridViewRow = Me.dgvDailyActivities.Rows(e.RowIndex)

        Dim CboDgv As DataGridViewComboBoxCell
        Try
            With row
                If e.ColumnIndex = Me.ColWorkdays.Index Then
                    CboDgv = Me.dgvDailyActivities.Item(e.ColumnIndex, e.RowIndex)
                    If Me.dgvDailyActivities.Item(e.ColumnIndex, e.RowIndex).Value = _
                        CboDgv.Items.Item(ClsMasterWorkShift.TipeWorkdays.Holiday) Then
                        Me.dgvDailyActivities.Item(Me.ColSchIn.Index, e.RowIndex).Value = "00.00"
                        Me.dgvDailyActivities.Item(Me.ColSchOut.Index, e.RowIndex).Value = "00.00"
                        Me.dgvDailyActivities.Item(Me.ColSchBreakIn_WA.Index, e.RowIndex).Value = "00.00"
                        Me.dgvDailyActivities.Item(Me.ColSchBreakOut_WA.Index, e.RowIndex).Value = "00.00"
                    End If
                End If

                If e.ColumnIndex = Me.ColSchIn.Index Then
                    'cek format
                    If String.IsNullOrEmpty(Me.dgvDailyActivities.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        If Me.ObjCd.Is_Time_String(Me.dgvDailyActivities.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            If Me.ObjCd.Convert_ToMinutes(Me.dgvDailyActivities.Item(e.ColumnIndex, e.RowIndex).Value) > 1439 Then
                                MsgBox("Schedule in not more than 23.59 or 1439 minutes", MsgBoxStyle.Information, "Information")
                                Me.dgvDailyActivities.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            End If
                            Me.ComparedWithDgvHoliday(Me.dgvDailyActivities.Item(Me.ColDay.Index, e.RowIndex).Value, _
                            Me.dgvDailyActivities.Item(Me.ColSchIn.Index, e.RowIndex).Value, Me.dgvDailyActivities.Item(Me.ColSchOut.Index, e.RowIndex).Value, _
                            Me.dgvDailyActivities.Item(Me.ColSchBreakIn_WA.Index, e.RowIndex).Value, Me.dgvDailyActivities.Item(Me.ColSchBreakOut_WA.Index, e.RowIndex).Value)
                        Else
                            MsgBox("Please input hh.mm format", MsgBoxStyle.Information, "Information")
                            Me.dgvDailyActivities.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        End If
                    Else
                        Me.dgvDailyActivities.Item(e.ColumnIndex, e.RowIndex).Value = "00.00"
                    End If
                End If
                If e.ColumnIndex = Me.ColSchOut.Index Then
                    'cek format
                    If String.IsNullOrEmpty(Me.dgvDailyActivities.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        If Me.ObjCd.Is_Time_String(Me.dgvDailyActivities.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            If Me.ObjCd.Convert_ToMinutes(Me.dgvDailyActivities.Item(e.ColumnIndex, e.RowIndex).Value) > 1439 Then
                                MsgBox("Schedule Out not more than 23.59 or 1439 minutes", MsgBoxStyle.Information, "Information")
                                Me.dgvDailyActivities.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            End If
                            Me.ComparedWithDgvHoliday(Me.dgvDailyActivities.Item(Me.ColDay.Index, e.RowIndex).Value, _
                           Me.dgvDailyActivities.Item(Me.ColSchIn.Index, e.RowIndex).Value, Me.dgvDailyActivities.Item(Me.ColSchOut.Index, e.RowIndex).Value, _
                           Me.dgvDailyActivities.Item(Me.ColSchBreakIn_WA.Index, e.RowIndex).Value, Me.dgvDailyActivities.Item(Me.ColSchBreakOut_WA.Index, e.RowIndex).Value)
                        Else
                            MsgBox("Please input hh.mm format", MsgBoxStyle.Information, "Information")
                            Me.dgvDailyActivities.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        End If
                    Else
                        Me.dgvDailyActivities.Item(e.ColumnIndex, e.RowIndex).Value = "00.00"
                    End If
                End If
                If e.ColumnIndex = Me.ColSchBreakOut_WA.Index Then
                    'cek format
                    If String.IsNullOrEmpty(Me.dgvDailyActivities.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        If Me.ObjCd.Is_Time_String(Me.dgvDailyActivities.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            If Me.ObjCd.Convert_ToMinutes(Me.dgvDailyActivities.Item(e.ColumnIndex, e.RowIndex).Value) > 1439 Then
                                MsgBox("Schedule Out not more than 23.59 or 1439 minutes", MsgBoxStyle.Information, "Information")
                                Me.dgvDailyActivities.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            End If
                            Me.ComparedWithDgvHoliday(Me.dgvDailyActivities.Item(Me.ColDay.Index, e.RowIndex).Value, _
                           Me.dgvDailyActivities.Item(Me.ColSchIn.Index, e.RowIndex).Value, Me.dgvDailyActivities.Item(Me.ColSchOut.Index, e.RowIndex).Value, _
                           Me.dgvDailyActivities.Item(Me.ColSchBreakIn_WA.Index, e.RowIndex).Value, Me.dgvDailyActivities.Item(Me.ColSchBreakOut_WA.Index, e.RowIndex).Value)
                        Else
                            MsgBox("Please input hh.mm format", MsgBoxStyle.Information, "Information")
                            Me.dgvDailyActivities.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        End If
                    Else
                        Me.dgvDailyActivities.Item(e.ColumnIndex, e.RowIndex).Value = "00.00"
                    End If
                End If
                If e.ColumnIndex = Me.ColSchBreakIn_WA.Index Then
                    'cek format
                    If String.IsNullOrEmpty(Me.dgvDailyActivities.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        If Me.ObjCd.Is_Time_String(Me.dgvDailyActivities.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            If Me.ObjCd.Convert_ToMinutes(Me.dgvDailyActivities.Item(e.ColumnIndex, e.RowIndex).Value) > 1439 Then
                                MsgBox("Schedule Out not more than 23.59 or 1439 minutes", MsgBoxStyle.Information, "Information")
                                Me.dgvDailyActivities.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            End If
                            Me.ComparedWithDgvHoliday(Me.dgvDailyActivities.Item(Me.ColDay.Index, e.RowIndex).Value, _
                           Me.dgvDailyActivities.Item(Me.ColSchIn.Index, e.RowIndex).Value, Me.dgvDailyActivities.Item(Me.ColSchOut.Index, e.RowIndex).Value, _
                           Me.dgvDailyActivities.Item(Me.ColSchBreakIn_WA.Index, e.RowIndex).Value, Me.dgvDailyActivities.Item(Me.ColSchBreakOut_WA.Index, e.RowIndex).Value)
                        Else
                            MsgBox("Please input hh.mm format", MsgBoxStyle.Information, "Information")
                            Me.dgvDailyActivities.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        End If
                    Else
                        Me.dgvDailyActivities.Item(e.ColumnIndex, e.RowIndex).Value = "00.00"
                    End If
                End If
                ''date
                'If e.ColumnIndex = Me.ColItemHoliday_Date.Index Then
                '    If Me.ObjWorkShift.GetJmlAtt(Me.DgvHolidayList.Item(Me.ColItemHoliday_Date.Index, e.RowIndex).Value, Me.TxtPrimaryJob.Text) > 0 Then
                '        MsgBox("Holiday list : " + Me.DgvHolidayList.Item(Me.ColItemHoliday_Date.Index, e.RowIndex).Value + _
                '        " , Cannot delete because refered attendance item.", MsgBoxStyle.Information, "Information")
                '        Me.DgvHolidayList.Item(Me.ColItemHoliday_Date.Index, e.RowIndex).Value = Me.buffer
                '    End If
                'End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvHolidayList_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvHolidayList.CellBeginEdit
        Dim tgl As Date = #1/1/1950#
        Try
            If Me.TxtJobCode.Enabled = False Then
                e.Cancel = True
            Else
                My.Application.ChangeUICulture("id-ID")
                My.Application.ChangeCulture("id-ID")
                tgl = CDate(Me.DgvHolidayList.Item(Me.ColItemHoliday_Date.Index, e.RowIndex).Value)
                If String.IsNullOrEmpty(Me.DgvHolidayList(Me.ColItemHoliday_Date.Index, e.RowIndex).Value) = False And _
                Me.TxtJobCode.Enabled = True Then
                    If e.ColumnIndex = Me.ColItemHoliday_Duration.Index Or e.ColumnIndex = Me.ColItemHoliday_SchIn.Index Or _
                    e.ColumnIndex = Me.ColItemHoliday_SchOut.Index Or e.ColumnIndex = Me.ColItemHoliday_Description.Index Or _
                    e.ColumnIndex = Me.ColItemHoliday_Overtime.Index Or e.ColumnIndex = Me.ColItemHoliday_Leave.Index Or _
                    e.ColumnIndex = Me.ColItemHoliday_Date.Index Or e.ColumnIndex = Me.ColItemHoliday_Remarks.Index Or _
                    e.ColumnIndex = Me.ColSchBreakIn_Holiday.Index Or e.ColumnIndex = Me.ColSchBreakOut_Holiday.Index Then
                        'If Me.ObjWorkShift.GetJmlAtt(tgl, Me.TxtPrimaryJob.Text) > 0 Then
                        '    MsgBox("Holiday list : " + Me.DgvHolidayList.Item(Me.ColItemHoliday_Date.Index, e.RowIndex).Value + _
                        '    " , Cannot edit because refered attendance item.", MsgBoxStyle.Information, "Information")
                        '    e.Cancel = True
                        'Else
                        '    Me.buffer = Me.DgvHolidayList.Item(e.ColumnIndex, e.RowIndex).Value
                        'End If
                        Me.buffer = Me.DgvHolidayList.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        e.Cancel = True
                    End If
                Else
                    'If Me.ObjWorkShift.GetJmlAtt(tgl, Me.TxtPrimaryJob.Text) > 0 Then
                    '    MsgBox("Holiday list : " + Me.DgvHolidayList.Item(Me.ColItemHoliday_Date.Index, e.RowIndex).Value + _
                    '    " , Cannot edit because refered attendance item.", MsgBoxStyle.Information, "Information")
                    '    e.Cancel = True
                    'Else
                    If e.ColumnIndex = Me.ColItemHoliday_Date.Index Then
                        Me.buffer = Me.DgvHolidayList.Item(e.ColumnIndex, e.RowIndex).Value
                    Else
                        e.Cancel = True
                    End If
                    'End If
                End If
                My.Application.ChangeCulture("en-US")
                My.Application.ChangeUICulture("en-US")
            End If
            Me.buffer = Me.DgvHolidayList.Item(e.ColumnIndex, e.RowIndex).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvHolidayList_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvHolidayList.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvHolidayList.Rows(e.RowIndex)
        Dim Primary As Double = -1
        Dim ObjCuti As New ClsMasterCuti
        Dim rowDGV, rowDGV1, rowDGVB4 As DataGridViewRow
        Dim DOW As Integer = 0

        Try
            With row
                'date
                If e.ColumnIndex = Me.ColItemHoliday_Date.Index Then
                    If String.IsNullOrEmpty(Me.DgvHolidayList.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        'set region ke indonesia
                        'My.Application.ChangeUICulture("id-ID")
                        'My.Application.ChangeCulture("id-ID")
                        If IsDate(Me.DgvHolidayList.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            MsgBox("Must input date format.", MsgBoxStyle.Information, "Information")
                            Me.EraseRowIndex = row.Index
                            Me.BackgroundWorker1.RunWorkerAsync()
                            Exit Sub
                        Else
                            Me.DgvHolidayList.Item(e.ColumnIndex, e.RowIndex).Value = _
                            Format(CDate(Me.DgvHolidayList.Item(e.ColumnIndex, e.RowIndex).Value), "dd MMM yyyy")
                            Me.DgvHolidayList.Item(Me.ColItemHoliday_Day.Index, e.RowIndex).Value = _
                            Me.ObjCd.GetHari_English(CDate(Me.DgvHolidayList.Item(e.ColumnIndex, e.RowIndex).Value))
                        End If
                        'kembalikan ke region defaultnya
                        'My.Application.ChangeCulture("en-US")
                        'My.Application.ChangeUICulture("en-US")
                        Dim idx As Integer = 0
                        For Each row In Me.DgvHolidayList.Rows
                            If idx <> e.RowIndex Then 'And idx <> 1
                                If row.Cells(Me.ColItemHoliday_Date.Index).Value = _
                                Me.DgvHolidayList.Item(e.ColumnIndex, e.RowIndex).Value Then
                                    MsgBox("Duplication in date, please change your date", MsgBoxStyle.Information, "Information")
                                    Me.EraseRowIndex = e.RowIndex
                                    Me.BackgroundWorker1.RunWorkerAsync()
                                End If
                            End If
                            idx = idx + 1
                        Next
                        'sch in dan sch out
                        rowDGV = Me.DgvHolidayList.Rows(Me.DgvHolidayList.Rows.Count - 2)
                        For Each rowDGV1 In Me.dgvDailyActivities.Rows
                            If CDate(rowDGV.Cells(Me.ColItemHoliday_Date.Index).Value).DayOfWeek = DOW Then
                                rowDGV.Cells(Me.ColItemHoliday_SchIn.Index).Value = rowDGV1.Cells(Me.ColSchIn.Index).Value
                                rowDGV.Cells(Me.ColItemHoliday_SchOut.Index).Value = rowDGV1.Cells(Me.ColSchOut.Index).Value
                                rowDGV.Cells(Me.ColSchBreakIn_Holiday.Index).Value = rowDGV1.Cells(Me.ColSchBreakIn_WA.Index).Value
                                rowDGV.Cells(Me.ColSchBreakOut_Holiday.Index).Value = rowDGV1.Cells(Me.ColSchBreakOut_WA.Index).Value
                                Exit For
                            End If
                            DOW += 1
                        Next

                        'WAJIB SINKRON 2019-03-01
                        If rowDGV.Index <> 0 Then
                            rowDGVB4 = Me.DgvHolidayList.Rows(rowDGV.Index - 1)
                            rowDGV.Cells(Me.ColItemHoliday_Overtime.Index).Value = rowDGVB4.Cells(Me.ColItemHoliday_Overtime.Index).Value
                        End If
                    Else
                        If e.RowIndex <> Me.DgvHolidayList.RowCount - 1 Then
                            Me.EraseRowIndex = e.RowIndex
                            Me.BackgroundWorker1.RunWorkerAsync()
                        End If

                    End If
                End If


                'sch in
                If e.ColumnIndex = Me.ColItemHoliday_SchIn.Index Then
                    'cek format
                    If String.IsNullOrEmpty(Me.DgvHolidayList.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        If Me.ObjCd.Is_Time_String(Me.DgvHolidayList.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            If Me.ObjCd.Convert_ToMinutes(Me.DgvHolidayList.Item(e.ColumnIndex, e.RowIndex).Value) > 1439 Then
                                MsgBox("Schedule in not more than 23.59 or 1439 minutes", MsgBoxStyle.Information, "Information")
                                Me.DgvHolidayList.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            End If
                        Else
                            MsgBox("Please input hh.mm format", MsgBoxStyle.Information, "Information")
                            Me.DgvHolidayList.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        End If
                    Else
                        Me.DgvHolidayList.Item(e.ColumnIndex, e.RowIndex).Value = "00.00"
                    End If
                End If
                'sch out
                If e.ColumnIndex = Me.ColItemHoliday_SchOut.Index Then
                    'cek format
                    If String.IsNullOrEmpty(Me.DgvHolidayList.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        If Me.ObjCd.Is_Time_String(Me.DgvHolidayList.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            If Me.ObjCd.Convert_ToMinutes(Me.DgvHolidayList.Item(e.ColumnIndex, e.RowIndex).Value) > 1439 Then
                                MsgBox("Schedule Out not more than 23.59 or 1439 minutes", MsgBoxStyle.Information, "Information")
                                Me.DgvHolidayList.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            End If
                        Else
                            MsgBox("Please input hh.mm format", MsgBoxStyle.Information, "Information")
                            Me.DgvHolidayList.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        End If
                    Else
                        Me.DgvHolidayList.Item(e.ColumnIndex, e.RowIndex).Value = "00.00"
                    End If
                End If
                'sch break out
                If e.ColumnIndex = Me.ColSchBreakOut_Holiday.Index Then
                    'cek format
                    If String.IsNullOrEmpty(Me.DgvHolidayList.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        If Me.ObjCd.Is_Time_String(Me.DgvHolidayList.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            If Me.ObjCd.Convert_ToMinutes(Me.DgvHolidayList.Item(e.ColumnIndex, e.RowIndex).Value) > 1439 Then
                                MsgBox("Schedule break out not more than 23.59 or 1439 minutes", MsgBoxStyle.Information, "Information")
                                Me.DgvHolidayList.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            End If
                        Else
                            MsgBox("Please input hh.mm format", MsgBoxStyle.Information, "Information")
                            Me.DgvHolidayList.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        End If
                    Else
                        Me.DgvHolidayList.Item(e.ColumnIndex, e.RowIndex).Value = "00.00"
                    End If
                End If
                'sch break in
                If e.ColumnIndex = Me.ColSchBreakIn_Holiday.Index Then
                    'cek format
                    If String.IsNullOrEmpty(Me.DgvHolidayList.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        If Me.ObjCd.Is_Time_String(Me.DgvHolidayList.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            If Me.ObjCd.Convert_ToMinutes(Me.DgvHolidayList.Item(e.ColumnIndex, e.RowIndex).Value) > 1439 Then
                                MsgBox("Schedule break in not more than 23.59 or 1439 minutes", MsgBoxStyle.Information, "Information")
                                Me.DgvHolidayList.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            End If
                        Else
                            MsgBox("Please input hh.mm format", MsgBoxStyle.Information, "Information")
                            Me.DgvHolidayList.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        End If
                    Else
                        Me.DgvHolidayList.Item(e.ColumnIndex, e.RowIndex).Value = "00.00"
                    End If
                End If
                'Leave(code)
                If e.ColumnIndex = Me.ColItemHoliday_Leave.Index Then
                    Dim ij As New FrmMasterCuti_Cari
                    If String.IsNullOrEmpty(Me.DgvHolidayList.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        If Me.ObjCuti.ApakahSudahAda(Me.DgvHolidayList.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            Primary = Me.ObjCuti.GetPrimaryByCode(Me.DgvHolidayList.Item(e.ColumnIndex, e.RowIndex).Value)
                            Me.DgvHolidayList.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCuti.GetCodeByPrimary(Primary)
                            Me.DgvHolidayList.Item(Me.ColItemHoliday_PrimaryLeave.Index, e.RowIndex).Value = Primary
                            If ObjCuti.ApakahSudahAda(Me.DgvHolidayList.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                                MsgBox("Leave code is not valid.", MsgBoxStyle.Information, "Information")
                                Me.DgvHolidayList.Item(e.ColumnIndex, e.RowIndex).Value = ""
                                Me.DgvHolidayList.Item(Me.ColItemHoliday_PrimaryLeave.Index, e.RowIndex).Value = -1
                            End If
                            'If ObjCuti.GetPeriodByPrimary(Me.ObjCuti.GetPrimaryByCode(Me.DgvHolidayList.Item(e.ColumnIndex, e.RowIndex).Value)).Date.Year <> _
                            'Format(, "yyyy") Then
                            '    MsgBox("Leave code period is not appropriate.", MsgBoxStyle.Information, "Information")
                            '    Me.DgvHolidayList.Item(e.ColumnIndex, e.RowIndex).Value = ""
                            '    Me.DgvHolidayList.Item(Me.ColItemHoliday_PrimaryLeave.Index, e.RowIndex).Value = -1
                            'End If
                        Else
                            ij.txtCari.Text = Me.DgvHolidayList.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryCuti = -1 Then
                                Me.DgvHolidayList.Item(e.ColumnIndex, e.RowIndex).Value = ""
                            Else
                                Me.DgvHolidayList.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCuti.GetCodeByPrimary(ij.PrimaryCuti)
                                Me.DgvHolidayList.Item(Me.ColItemHoliday_PrimaryLeave.Index, e.RowIndex).Value = ij.PrimaryCuti
                                If ObjCuti.ApakahSudahAda(Me.DgvHolidayList.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                                    MsgBox("Leave code is not valid.", MsgBoxStyle.Information, "Information")
                                    Me.DgvHolidayList.Item(e.ColumnIndex, e.RowIndex).Value = ""
                                    Me.DgvHolidayList.Item(Me.ColItemHoliday_PrimaryLeave.Index, e.RowIndex).Value = -1
                                End If
                                If ObjCuti.GetPeriodByPrimary(Me.ObjCuti.GetPrimaryByCode(Me.DgvHolidayList.Item(e.ColumnIndex, e.RowIndex).Value)).Date.Year <> _
                                Format(Now, "yyyy") Then
                                    MsgBox("Leave code period is not appropriate.", MsgBoxStyle.Information, "Information")
                                    Me.DgvHolidayList.Item(e.ColumnIndex, e.RowIndex).Value = ""
                                    Me.DgvHolidayList.Item(Me.ColItemHoliday_PrimaryLeave.Index, e.RowIndex).Value = -1
                                End If
                            End If
                            ij = Nothing
                        End If
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvHolidayList_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DgvHolidayList.RowsAdded
        Dim CboDgv As DataGridViewComboBoxCell
        Try
            Me.ObjCd.SetNumber(Me.DgvHolidayList)
            CboDgv = Me.DgvHolidayList.Item(Me.ColItemHoliday_Duration.Index, e.RowIndex)
            Me.DgvHolidayList.Item(Me.ColItemHoliday_Duration.Index, e.RowIndex).Value = _
            CboDgv.Items.Item(ClsMasterWorkShift.enumDuration.fullday)
            Me.DgvHolidayList.Item(Me.ColItemHoliday_SchIn.Index, e.RowIndex).Value = "00.00"
            Me.DgvHolidayList.Item(Me.ColItemHoliday_SchOut.Index, e.RowIndex).Value = "00.00"
            Me.DgvHolidayList.Item(Me.ColSchBreakIn_Holiday.Index, e.RowIndex).Value = "00.00"
            Me.DgvHolidayList.Item(Me.ColSchBreakOut_Holiday.Index, e.RowIndex).Value = "00.00"
            CboDgv = Me.DgvHolidayList.Item(Me.ColItemHoliday_Overtime.Index, e.RowIndex)
            Me.DgvHolidayList.Item(Me.ColItemHoliday_Overtime.Index, e.RowIndex).Value = _
            CboDgv.Items.Item(ClsMasterWorkShift.TipeOvertime.Holiday)
            Me.DgvHolidayList.Item(Me.ColItemHoliday_PrimaryLeave.Index, e.RowIndex).Value = -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvHolidayList_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles DgvHolidayList.RowsRemoved
        Dim CboDgv As DataGridViewComboBoxCell
        Try
            If Me.TxtJobCode.Enabled = True Then
                If Me.DgvHolidayList.RowCount > 0 Then
                    Me.ObjCd.SetNumber(Me.DgvHolidayList)
                    CboDgv = Me.DgvHolidayList.Item(Me.ColItemHoliday_Duration.Index, e.RowIndex)
                    Me.DgvHolidayList.Item(Me.ColItemHoliday_Duration.Index, e.RowIndex).Value = _
                    CboDgv.Items.Item(ClsMasterWorkShift.enumDuration.fullday)
                    Me.DgvHolidayList.Item(Me.ColItemHoliday_SchIn.Index, e.RowIndex).Value = "00.00"
                    Me.DgvHolidayList.Item(Me.ColItemHoliday_SchOut.Index, e.RowIndex).Value = "00.00"
                    Me.DgvHolidayList.Item(Me.ColSchBreakIn_Holiday.Index, e.RowIndex).Value = "00.00"
                    Me.DgvHolidayList.Item(Me.ColSchBreakOut_Holiday.Index, e.RowIndex).Value = "00.00"
                    CboDgv = Me.DgvHolidayList.Item(Me.ColItemHoliday_Overtime.Index, e.RowIndex)
                    Me.DgvHolidayList.Item(Me.ColItemHoliday_Overtime.Index, e.RowIndex).Value = _
                    CboDgv.Items.Item(ClsMasterWorkShift.TipeOvertime.Holiday)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvHolidayList_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvHolidayList.Sorted
        Me.ObjCd.SetNumber(Me.DgvHolidayList)
    End Sub

    Private Sub TxtJobCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtJobCode.GotFocus
        Me.TxtJobCode.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtJobCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtJobCode.LostFocus
        Me.TxtJobCode.BackColor = Color.White
    End Sub

    Private Sub TxtJobName_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtJobName.GotFocus
        Me.TxtJobName.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtJobName_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtJobName.LostFocus
        Me.TxtJobName.BackColor = Color.White
    End Sub

    Private Sub TxtRemark_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemark.GotFocus
        Me.TxtRemark.BackColor = Me.BackColorAll
    End Sub

    Private Sub TxtRemark_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemark.LostFocus
        Me.TxtRemark.BackColor = Color.White
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Me.DgvHolidayList.Rows.RemoveAt(Me.EraseRowIndex)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvHolidayList_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles DgvHolidayList.UserDeletingRow
        Dim tgl As Date
        Try
            'My.Application.ChangeUICulture("id-ID")
            'My.Application.ChangeCulture("id-ID")
            'tgl = CDate(Me.DgvHolidayList.Item(Me.ColItemHoliday_Date.Index, e.Row.Index).Value)
            'My.Application.ChangeCulture("en-US")
            'My.Application.ChangeUICulture("en-US")
            'If Me.ObjWorkShift.GetJmlAtt(tgl, Me.TxtPrimaryJob.Text) > 0 Then
            '    MsgBox("Holiday list : " + Me.DgvHolidayList.Item(Me.ColItemHoliday_Date.Index, e.Row.Index).Value + _
            '    " , Cannot delete because refered attendance item.", MsgBoxStyle.Information, "Information")
            '    e.Cancel = True
            'End If
            'Batasan cek tanggal attendance maximum, tidak boleh hapus holiday list dibawah tanggal maksimum
            My.Application.ChangeUICulture("id-ID")
            My.Application.ChangeCulture("id-ID")
            tgl = CDate(e.Row.Cells(ColItemHoliday_Date.Index).Value)
            My.Application.ChangeCulture("en-US")
            My.Application.ChangeUICulture("en-US")
            If tgl < Me.ObjWorkShift.GetDateMAXAttendance(Me.TxtPrimaryJob.Text) Then
                MsgBox("Holiday list period " + e.Row.Cells(ColItemHoliday_Date.Index).Value.ToString + " has already made attendance data.", MsgBoxStyle.Information, "Information")
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvAttParams_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvAttParams.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColAtt_Schedule.Index Then
                e.Cancel = True
            Else
                If e.ColumnIndex = Me.ColAtt_DefaultCustom.Index And Me.DgvAttParams.Item(Me.ColAtt_Default.Index, e.RowIndex).Value.ToString.ToUpper = _
                ClsMasterWorkShift.EnumDefaultIfNoClockInOut.SCHEDULE.ToString And Me.DgvAttParams.Item(Me.ColAtt_Outstation.Index, e.RowIndex).Value.ToString.ToUpper = _
                ClsMasterWorkShift.EnumDefaultIfOutstation.SCHEDULE.ToString Then
                    e.Cancel = True
                Else
                    If e.ColumnIndex <> Me.ColAtt_GetValue.Index And Me.DgvAttParams.Item(Me.ColAtt_GetValue.Index, e.RowIndex).Value = Boolean.FalseString Then
                        e.Cancel = True
                    Else
                        Me.buffer = Me.DgvAttParams.Item(e.ColumnIndex, e.RowIndex).Value
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvAttParams_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvAttParams.CellEndEdit
        Dim vCheckRange As ClsMasterWorkShift.EnumCheckRange = ClsMasterWorkShift.EnumCheckRange.AllRange
        Try
            If e.ColumnIndex = Me.ColAtt_GetValue.Index And Me.DgvAttParams.Item(Me.ColAtt_ScheduleType.Index, e.RowIndex).Value.ToString = ClsMasterWorkShift.EnumScheduleType.BREAK_OUT.ToString Then
                For Each rowDgv As DataGridViewRow In Me.DgvAttParams.Rows
                    If rowDgv.Cells(Me.ColAtt_ScheduleType.Index).Value.ToString = ClsMasterWorkShift.EnumScheduleType.BREAK_IN.ToString Then
                        If Me.DgvAttParams.Item(Me.ColAtt_GetValue.Index, e.RowIndex).Value.ToString = Boolean.TrueString Then
                            rowDgv.Cells(Me.ColAtt_GetValue.Index).Value = Boolean.TrueString
                        Else
                            rowDgv.Cells(Me.ColAtt_GetValue.Index).Value = Boolean.FalseString
                        End If
                    End If
                Next
            End If
            If e.ColumnIndex = Me.ColAtt_GetValue.Index And Me.DgvAttParams.Item(Me.ColAtt_ScheduleType.Index, e.RowIndex).Value.ToString = ClsMasterWorkShift.EnumScheduleType.BREAK_IN.ToString Then
                For Each rowDgv As DataGridViewRow In Me.DgvAttParams.Rows
                    If rowDgv.Cells(Me.ColAtt_ScheduleType.Index).Value.ToString = ClsMasterWorkShift.EnumScheduleType.BREAK_OUT.ToString Then
                        If Me.DgvAttParams.Item(Me.ColAtt_GetValue.Index, e.RowIndex).Value.ToString = Boolean.TrueString Then
                            rowDgv.Cells(Me.ColAtt_GetValue.Index).Value = Boolean.TrueString
                        Else
                            rowDgv.Cells(Me.ColAtt_GetValue.Index).Value = Boolean.FalseString
                        End If
                    End If
                Next
            End If
            If e.ColumnIndex = Me.ColAtt_LimitBefore.Index Or e.ColumnIndex = Me.ColAtt_LimitAfter.Index Then
                If IsNothing(Me.DgvAttParams.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If Me.DgvAttParams.Item(e.ColumnIndex, e.RowIndex).Value.ToString <> "" Then
                        If IsNumeric(Me.DgvAttParams.Item(e.ColumnIndex, e.RowIndex).Value.ToString) = False Then
                            MsgBox("You have to input numeric.", MsgBoxStyle.Information, "Information")
                            If Me.buffer <> "" Then
                                Me.DgvAttParams.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Else
                                Me.DgvAttParams.Item(e.ColumnIndex, e.RowIndex).Value = "0"
                            End If
                        End If
                        If Me.DgvAttParams.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains(",") Or _
                            Me.DgvAttParams.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Contains(".") Then
                            MsgBox("You have to input numeric.", MsgBoxStyle.Information, "Information")
                            If Me.buffer <> "" Then
                                Me.DgvAttParams.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Else
                                Me.DgvAttParams.Item(e.ColumnIndex, e.RowIndex).Value = "0"
                            End If
                        End If
                        Me.DgvAttParams.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCd.DigitGrouping(Me.DgvAttParams.Item(e.ColumnIndex, e.RowIndex).Value.ToString, 0)

                        ''hanya cek validitas parameter ke jangkauan yang mungkin aja
                        If (Me.DgvAttParams.Item(Me.ColAtt_ScheduleType.Index, e.RowIndex).Value.ToString = ClsMasterWorkShift.EnumScheduleType.SCHEDULE_IN.ToString And e.ColumnIndex = Me.ColAtt_LimitAfter.Index) Or _
                                (Me.DgvAttParams.Item(Me.ColAtt_ScheduleType.Index, e.RowIndex).Value.ToString = ClsMasterWorkShift.EnumScheduleType.BREAK_OUT.ToString And e.ColumnIndex = Me.ColAtt_LimitBefore.Index) Then
                            vCheckRange = ClsMasterWorkShift.EnumCheckRange.SchIn_SchBkOut
                        ElseIf (Me.DgvAttParams.Item(Me.ColAtt_ScheduleType.Index, e.RowIndex).Value.ToString = ClsMasterWorkShift.EnumScheduleType.BREAK_OUT.ToString And e.ColumnIndex = Me.ColAtt_LimitAfter.Index) Or _
                                (Me.DgvAttParams.Item(Me.ColAtt_ScheduleType.Index, e.RowIndex).Value.ToString = ClsMasterWorkShift.EnumScheduleType.BREAK_IN.ToString And e.ColumnIndex = Me.ColAtt_LimitBefore.Index) Then
                            vCheckRange = ClsMasterWorkShift.EnumCheckRange.SchBOut_SchBIn
                        ElseIf (Me.DgvAttParams.Item(Me.ColAtt_ScheduleType.Index, e.RowIndex).Value.ToString = ClsMasterWorkShift.EnumScheduleType.BREAK_IN.ToString And e.ColumnIndex = Me.ColAtt_LimitAfter.Index) Or _
                                (Me.DgvAttParams.Item(Me.ColAtt_ScheduleType.Index, e.RowIndex).Value.ToString = ClsMasterWorkShift.EnumScheduleType.SCHEDULE_OUT.ToString And e.ColumnIndex = Me.ColAtt_LimitBefore.Index) Then
                            vCheckRange = ClsMasterWorkShift.EnumCheckRange.SchBIn_SchOut
                        End If
                        If AttParametersCheck(vCheckRange) = False Then
                            If Me.buffer <> "" Then
                                Me.DgvAttParams.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Else
                                Me.DgvAttParams.Item(e.ColumnIndex, e.RowIndex).Value = "0"
                            End If
                        End If

                    End If
                Else
                    If Me.buffer <> "" Then
                        Me.DgvAttParams.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Else
                        Me.DgvAttParams.Item(e.ColumnIndex, e.RowIndex).Value = "0"
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function AttParametersCheck(ByVal vCheckRange As ClsMasterWorkShift.EnumCheckRange) As Boolean
        Dim SchInDummy_, SchBOutDummy_, SchBInDummy_, SchOutDummy_, SchIn_, SchBOut_, SchBIn_, SchOut_ As Integer
        Dim LimitAfterSchIn, LimitBeforeSchBOut, LimitAfterSchBOut, LimitBeforeSchBIn, LimitAfterSchBIn, LimitBeforeSchOut As Integer
        Dim CboDgv As DataGridViewComboBoxCell

        Try
            For Each rowDgv As DataGridViewRow In Me.DgvAttParams.Rows
                With rowDgv
                    If .Cells(Me.ColAtt_ScheduleType.Index).Value.ToString = ClsMasterWorkShift.EnumScheduleType.SCHEDULE_IN.ToString Then
                        LimitAfterSchIn = .Cells(Me.ColAtt_LimitAfter.Index).Value.ToString
                    End If
                    If .Cells(Me.ColAtt_ScheduleType.Index).Value.ToString = ClsMasterWorkShift.EnumScheduleType.BREAK_OUT.ToString Then
                        LimitBeforeSchBOut = .Cells(Me.ColAtt_LimitBefore.Index).Value.ToString
                    End If
                    If .Cells(Me.ColAtt_ScheduleType.Index).Value.ToString = ClsMasterWorkShift.EnumScheduleType.BREAK_OUT.ToString Then
                        LimitAfterSchBOut = .Cells(Me.ColAtt_LimitAfter.Index).Value.ToString
                    End If
                    If .Cells(Me.ColAtt_ScheduleType.Index).Value.ToString = ClsMasterWorkShift.EnumScheduleType.BREAK_IN.ToString Then
                        LimitBeforeSchBIn = .Cells(Me.ColAtt_LimitBefore.Index).Value.ToString
                    End If
                    If .Cells(Me.ColAtt_ScheduleType.Index).Value.ToString = ClsMasterWorkShift.EnumScheduleType.BREAK_IN.ToString Then
                        LimitAfterSchBIn = .Cells(Me.ColAtt_LimitAfter.Index).Value.ToString
                    End If
                    If .Cells(Me.ColAtt_ScheduleType.Index).Value.ToString = ClsMasterWorkShift.EnumScheduleType.SCHEDULE_OUT.ToString Then
                        LimitBeforeSchOut = .Cells(Me.ColAtt_LimitBefore.Index).Value.ToString
                    End If
                End With
            Next

            For Each rowDgv As DataGridViewRow In Me.dgvDailyActivities.Rows
                CboDgv = rowDgv.Cells(Me.ColDuration.Index)
                'If rowDgv.Cells(Me.ColDuration.Index).Value = CboDgv.Items.Item(ClsMasterWorkShift.enumDuration.fullday) Then
                With rowDgv
                    SchInDummy_ = Me.ObjCd.Convert_ToMinutes(.Cells(Me.ColSchIn.Index).Value.ToString)
                    SchBOutDummy_ = Me.ObjCd.Convert_ToMinutes(.Cells(Me.ColSchBreakOut_WA.Index).Value.ToString)
                    SchBInDummy_ = Me.ObjCd.Convert_ToMinutes(.Cells(Me.ColSchBreakIn_WA.Index).Value.ToString)
                    SchOutDummy_ = Me.ObjCd.Convert_ToMinutes(.Cells(Me.ColSchOut.Index).Value.ToString)

                    SchIn_ = SchInDummy_
                    ''jika melewati hari::
                    If SchBOutDummy_ < SchInDummy_ Then
                        SchBOut_ = SchBOutDummy_ + Me.ObjCd.Convert_ToMinutes("23.59") + 1
                    Else
                        SchBOut_ = SchBOutDummy_
                    End If
                    If SchBInDummy_ < SchInDummy_ Then
                        SchBIn_ = SchBInDummy_ + Me.ObjCd.Convert_ToMinutes("23.59") + 1
                    Else
                        SchBIn_ = SchBInDummy_
                    End If
                    If SchBOutDummy_ < SchInDummy_ Then
                        SchOut_ = SchOutDummy_ + Me.ObjCd.Convert_ToMinutes("23.59") + 1
                    Else
                        SchOut_ = SchOutDummy_
                    End If

                    If SchBOutDummy_ <> SchBInDummy_ Then
                        If SchOut_ > SchIn_ Then
                            'cek jarak waktu limit yang diperbolehkan antara sch in, break out, break in, dan sch out
                            If vCheckRange = ClsMasterWorkShift.EnumCheckRange.AllRange Or vCheckRange = ClsMasterWorkShift.EnumCheckRange.SchIn_SchBkOut Then
                                If (SchIn_ + LimitAfterSchIn) >= (SchBOut_ - LimitBeforeSchBOut) Then
                                    MsgBox("[Limit After] for schedule in and [Limit Before] for break out " & vbNewLine & "is not valid if combine with '" & .Cells(Me.ColDay.Index).Value.ToString & "' settings.", MsgBoxStyle.Information, "Information")
                                    Return False
                                End If
                            End If
                            If vCheckRange = ClsMasterWorkShift.EnumCheckRange.AllRange Or vCheckRange = ClsMasterWorkShift.EnumCheckRange.SchBOut_SchBIn Then
                                If (SchBOut_ + LimitAfterSchBOut) >= (SchBIn_ - LimitBeforeSchBIn) Then
                                    MsgBox("[Limit After] for break out and [Limit Before] for break in " & vbNewLine & "is not valid if combine with '" & .Cells(Me.ColDay.Index).Value.ToString & "' settings.", MsgBoxStyle.Information, "Information")
                                    Return False
                                End If
                            End If
                            If vCheckRange = ClsMasterWorkShift.EnumCheckRange.AllRange Or vCheckRange = ClsMasterWorkShift.EnumCheckRange.SchBIn_SchOut Then
                                If (SchBIn_ + LimitAfterSchBIn) >= (SchOut_ - LimitBeforeSchOut) Then
                                    MsgBox("[Limit After] for break in and [Limit Before] for schedule out " & vbNewLine & "is not valid if combine with '" & .Cells(Me.ColDay.Index).Value.ToString & "' settings.", MsgBoxStyle.Information, "Information")
                                    Return False
                                End If
                            End If
                        Else
                            'cek jarak waktu limit yang diperbolehkan antara sch in, break out, break in, dan sch out
                            If vCheckRange = ClsMasterWorkShift.EnumCheckRange.AllRange Or vCheckRange = ClsMasterWorkShift.EnumCheckRange.SchIn_SchBkOut Then
                                If SchIn_ < SchBOut_ Then
                                    If (SchIn_ + LimitAfterSchIn) >= (SchBOut_ - LimitBeforeSchBOut) Then
                                        MsgBox("[Limit After] for schedule in and [Limit Before] for break out " & vbNewLine & "is not valid if combine with '" & .Cells(Me.ColDay.Index).Value.ToString & "' settings.", MsgBoxStyle.Information, "Information")
                                        Return False
                                    End If
                                Else
                                    If (SchIn_ + LimitAfterSchIn) >= (SchBOut_ + (24 * 60) - LimitBeforeSchBOut) Then
                                        MsgBox("[Limit After] for schedule in and [Limit Before] for break out " & vbNewLine & "is not valid if combine with '" & .Cells(Me.ColDay.Index).Value.ToString & "' settings.", MsgBoxStyle.Information, "Information")
                                        Return False
                                    End If
                                End If
                               
                            End If
                            If vCheckRange = ClsMasterWorkShift.EnumCheckRange.AllRange Or vCheckRange = ClsMasterWorkShift.EnumCheckRange.SchBOut_SchBIn Then
                                If SchBOut_ < SchBIn_ Then
                                    If (SchBOut_ + LimitAfterSchBOut) >= (SchBIn_ - LimitBeforeSchBIn) Then
                                        MsgBox("[Limit After] for break out and [Limit Before] for break in " & vbNewLine & "is not valid if combine with '" & .Cells(Me.ColDay.Index).Value.ToString & "' settings.", MsgBoxStyle.Information, "Information")
                                        Return False
                                    End If
                                Else
                                    If (SchBOut_ + LimitAfterSchBOut) >= (SchBIn_ + (24 * 60) - LimitBeforeSchBIn) Then
                                        MsgBox("[Limit After] for break out and [Limit Before] for break in " & vbNewLine & "is not valid if combine with '" & .Cells(Me.ColDay.Index).Value.ToString & "' settings.", MsgBoxStyle.Information, "Information")
                                        Return False
                                    End If
                                End If
                            End If
                            If vCheckRange = ClsMasterWorkShift.EnumCheckRange.AllRange Or vCheckRange = ClsMasterWorkShift.EnumCheckRange.SchBIn_SchOut Then
                                If SchBIn_ < SchOut_ Then
                                    If (SchBIn_ + LimitAfterSchBIn) >= (SchOut_ - LimitBeforeSchOut) Then
                                        MsgBox("[Limit After] for break in and [Limit Before] for schedule out " & vbNewLine & "is not valid if combine with '" & .Cells(Me.ColDay.Index).Value.ToString & "' settings.", MsgBoxStyle.Information, "Information")
                                        Return False
                                    End If
                                Else
                                    If (SchBIn_ + LimitAfterSchBIn) >= (SchOut_ + (24 * 60) - LimitBeforeSchOut) Then
                                        MsgBox("[Limit After] for break in and [Limit Before] for schedule out " & vbNewLine & "is not valid if combine with '" & .Cells(Me.ColDay.Index).Value.ToString & "' settings.", MsgBoxStyle.Information, "Information")
                                        Return False
                                    End If
                                End If
                            End If
                        End If
                    Else
                        'cek jarak waktu limit yang diperbolehkan antara sch in dan sch out
                        If vCheckRange = ClsMasterWorkShift.EnumCheckRange.AllRange Then
                            If (SchIn_ + LimitAfterSchIn) >= (SchOut_ - LimitBeforeSchOut) Then
                                MsgBox("[Limit After] for schedule in and [Limit Before] for schedule out " & vbNewLine & "is not valid if combine with '" & .Cells(Me.ColDay.Index).Value.ToString & "' settings.", MsgBoxStyle.Information, "Information")
                                Return False
                            End If
                        End If
                    End If
                End With
                'Else
                'jika half day, maka batasannya agak sedikit berbeda
                '
                'End If
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub CboIsCalculateEarlyOT_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboIsCalculateEarlyOT.SelectedIndexChanged
        Try
            If Me.TxtJobCode.Enabled = True Then
                If Me.CboIsCalculateEarlyOT.SelectedIndex = 0 Then
                    Me.txtCalculateAfterEarlyOT.Enabled = True
                    Me.cboRoundingTypeEarlyOT.Enabled = True
                    If Me.cboRoundingTypeEarlyOT.SelectedIndex <> 0 Then
                        Me.txtRoundingEarlyOT.Enabled = True
                    End If
                    Me.txtMaxValueEarlyOT.Enabled = True
                    Me.chkCalculateEarlyOTTolerance.Enabled = True
                    Me.chkCalculateEarlyOTTolerance.Checked = True
                    Me.ObjWorkShift.CalculateEarlyOTOnWeekend = 1
                    Me.ObjWorkShift.CalculateEarlyOTOnHoliday = 1

                    Me.chkUseOvertimeOrder.Enabled = True
                Else
                    Me.txtCalculateAfterEarlyOT.Enabled = False
                    Me.cboRoundingTypeEarlyOT.Enabled = False
                    Me.txtRoundingEarlyOT.Enabled = False
                    Me.txtMaxValueEarlyOT.Enabled = False
                    Me.txtCalculateAfterEarlyOT.Text = "00.00"
                    Me.cboRoundingTypeEarlyOT.SelectedIndex = 0
                    Me.txtRoundingEarlyOT.Text = "00.00"
                    Me.txtMaxValueEarlyOT.Text = "00.00"
                    Me.chkCalculateEarlyOTTolerance.Enabled = False
                    Me.chkCalculateEarlyOTTolerance.Checked = False
                    Me.ObjWorkShift.CalculateEarlyOTOnWeekend = 0
                    Me.ObjWorkShift.CalculateEarlyOTOnHoliday = 0

                    If Me.CboIsCalculateOT.SelectedIndex = 1 Then
                        Me.chkUseOvertimeOrder.Checked = False
                        Me.chkSendNotifEmailOvertimeOrder.Checked = False
                        Me.chkUseOvertimeOrder.Enabled = False
                        Me.chkSendNotifEmailOvertimeOrder.Enabled = False
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboIsCalculateLate_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboIsCalculateLate.SelectedIndexChanged
        Try
            If Me.TxtJobCode.Enabled = True Then
                If Me.CboIsCalculateLate.SelectedIndex = 0 Then
                    Me.txtCalculateAfterLate.Enabled = True
                    Me.cboRoundingTypeLate.Enabled = True
                    If Me.cboRoundingTypeLate.SelectedIndex <> 0 Then
                        Me.txtRoundingLate.Enabled = True
                    End If
                    Me.txtMaxValueLate.Enabled = True
                    Me.chkCalculateLateTolerance.Enabled = True
                    Me.chkCalculateLateTolerance.Checked = True
                    Me.ObjWorkShift.CalculateLateOnWeekend = 1
                    Me.ObjWorkShift.CalculateLateOnHoliday = 1
                Else
                    Me.txtCalculateAfterLate.Enabled = False
                    Me.cboRoundingTypeLate.Enabled = False
                    Me.txtRoundingLate.Enabled = False
                    Me.txtMaxValueLate.Enabled = False
                    Me.txtCalculateAfterLate.Text = "00.00"
                    Me.cboRoundingTypeLate.SelectedIndex = 0
                    Me.txtRoundingLate.Text = "00.00"
                    Me.txtMaxValueLate.Text = "00.00"
                    Me.chkCalculateLateTolerance.Enabled = False
                    Me.chkCalculateLateTolerance.Checked = False
                    Me.ObjWorkShift.CalculateLateOnWeekend = 0
                    Me.ObjWorkShift.CalculateLateOnHoliday = 0
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboIsCalculateEarlyOut_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboIsCalculateEarlyOut.SelectedIndexChanged
        Try
            If Me.TxtJobCode.Enabled = True Then
                If Me.CboIsCalculateEarlyOut.SelectedIndex = 0 Then
                    Me.txtCalculateAfterEarlyOut.Enabled = True
                    Me.cboRoundingTypeEarlyOut.Enabled = True
                    If Me.cboRoundingTypeEarlyOut.SelectedIndex <> 0 Then
                        Me.txtRoundingEarlyOut.Enabled = True
                    End If
                    Me.txtMaxValueEarlyOut.Enabled = True
                    Me.chkCalculateEarlyOutTolerance.Enabled = True
                    Me.chkCalculateEarlyOutTolerance.Checked = True
                    Me.ObjWorkShift.CalculateEarlyOutOnWeekend = 1
                    Me.ObjWorkShift.CalculateEarlyOutOnHoliday = 1
                Else
                    Me.txtCalculateAfterEarlyOut.Enabled = False
                    Me.cboRoundingTypeEarlyOut.Enabled = False
                    Me.txtRoundingEarlyOut.Enabled = False
                    Me.txtMaxValueEarlyOut.Enabled = False
                    Me.txtCalculateAfterEarlyOut.Text = "00.00"
                    Me.cboRoundingTypeEarlyOut.SelectedIndex = 0
                    Me.txtRoundingEarlyOut.Text = "00.00"
                    Me.txtMaxValueEarlyOut.Text = "00.00"
                    Me.chkCalculateEarlyOutTolerance.Enabled = False
                    Me.chkCalculateEarlyOutTolerance.Checked = False
                    Me.ObjWorkShift.CalculateEarlyOutOnWeekend = 0
                    Me.ObjWorkShift.CalculateEarlyOutOnHoliday = 0
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
   
    Private Sub CboIsCalculateOT_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboIsCalculateOT.SelectedIndexChanged
        Try
            If Me.TxtJobCode.Enabled = True Then
                If Me.CboIsCalculateOT.SelectedIndex = 0 Then
                    Me.txtCalculateAfterOT.Enabled = True
                    Me.cboRoundingTypeOT.Enabled = True
                    If Me.cboRoundingTypeOT.SelectedIndex <> 0 Then
                        Me.txtRoundingOT.Enabled = True
                    End If
                    Me.txtMaxValueOT.Enabled = True
                    Me.chkCalculateOTTolerance.Enabled = True
                    Me.chkCalculateOTTolerance.Checked = True
                    Me.ObjWorkShift.CalculateOTOnWeekend = 1
                    Me.ObjWorkShift.CalculateOTOnHoliday = 1

                    Me.chkUseOvertimeOrder.Enabled = True
                Else
                    Me.txtCalculateAfterOT.Enabled = False
                    Me.cboRoundingTypeOT.Enabled = False
                    Me.txtRoundingOT.Enabled = False
                    Me.txtMaxValueOT.Enabled = False
                    Me.txtCalculateAfterOT.Text = "00.00"
                    Me.cboRoundingTypeOT.SelectedIndex = 0
                    Me.txtRoundingOT.Text = "00.00"
                    Me.txtMaxValueOT.Text = "00.00"
                    Me.chkCalculateOTTolerance.Enabled = False
                    Me.chkCalculateOTTolerance.Checked = False
                    Me.ObjWorkShift.CalculateOTOnWeekend = 0
                    Me.ObjWorkShift.CalculateOTOnHoliday = 0

                    If Me.CboIsCalculateEarlyOT.SelectedIndex = 1 Then
                        Me.chkUseOvertimeOrder.Checked = False
                        Me.chkSendNotifEmailOvertimeOrder.Checked = False
                        Me.chkUseOvertimeOrder.Enabled = False
                        Me.chkSendNotifEmailOvertimeOrder.Enabled = False
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboRoundingTypeEarlyOT_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboRoundingTypeEarlyOT.SelectedIndexChanged
        Try
            If Me.TxtJobCode.Enabled = True Then
                If Me.cboRoundingTypeEarlyOT.SelectedIndex <> 0 Then
                    Me.txtRoundingEarlyOT.Enabled = True
                Else
                    Me.txtRoundingEarlyOT.Enabled = False
                    Me.txtRoundingEarlyOT.Text = "00.00"
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboRoundingTypeLate_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboRoundingTypeLate.SelectedIndexChanged
        Try
            If Me.TxtJobCode.Enabled = True Then
                If Me.cboRoundingTypeLate.SelectedIndex <> 0 Then
                    Me.txtRoundingLate.Enabled = True
                Else
                    Me.txtRoundingLate.Enabled = False
                    Me.txtRoundingLate.Text = "00.00"
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboRoundingTypeEarlyOut_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboRoundingTypeEarlyOut.SelectedIndexChanged
        Try
            If Me.TxtJobCode.Enabled = True Then
                If Me.cboRoundingTypeEarlyOut.SelectedIndex <> 0 Then
                    Me.txtRoundingEarlyOut.Enabled = True
                Else
                    Me.txtRoundingEarlyOut.Enabled = False
                    Me.txtRoundingEarlyOut.Text = "00.00"
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboRoundingTypeOT_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboRoundingTypeOT.SelectedIndexChanged
        Try
            If Me.TxtJobCode.Enabled = True Then
                If Me.cboRoundingTypeOT.SelectedIndex <> 0 Then
                    Me.txtRoundingOT.Enabled = True
                Else
                    Me.txtRoundingOT.Enabled = False
                    Me.txtRoundingOT.Text = "00.00"
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Try
            Me.txtSelect.SelectAll()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCalculateAfterEarlyOT_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCalculateAfterEarlyOT.Enter, txtCalculateAfterLate.Enter, txtCalculateAfterEarlyOut.Enter, txtCalculateAfterOT.Enter, txtRoundingEarlyOT.Enter, txtRoundingLate.Enter, txtRoundingEarlyOut.Enter, txtRoundingOT.Enter, txtMaxValueEarlyOT.Enter, txtMaxValueLate.Enter, txtMaxValueEarlyOut.Enter, txtMaxValueOT.Enter
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnOK.Enabled = True Then
                    If Txt.Enabled = True Then
                        Me.buffer = Txt.Text
                        Me.txtSelect = Txt
                        Me.BackgroundWorker2.RunWorkerAsync()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCalculateAfterEarlyOT_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCalculateAfterEarlyOT.GotFocus, txtCalculateAfterLate.GotFocus, txtCalculateAfterEarlyOut.GotFocus, txtCalculateAfterOT.GotFocus, txtRoundingEarlyOT.GotFocus, txtRoundingLate.GotFocus, txtRoundingEarlyOut.GotFocus, txtRoundingOT.GotFocus, txtMaxValueEarlyOT.GotFocus, txtMaxValueLate.GotFocus, txtMaxValueEarlyOut.GotFocus, txtMaxValueOT.GotFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnOK.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.TxtReadOnly.BackColor
                        Else
                            Txt.BackColor = Me.BackColorAll
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCalculateAfterEarlyOT_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCalculateAfterEarlyOT.Leave, txtCalculateAfterLate.Leave, txtCalculateAfterEarlyOut.Leave, txtCalculateAfterOT.Leave, txtRoundingEarlyOT.Leave, txtRoundingLate.Leave, txtRoundingEarlyOut.Leave, txtRoundingOT.Leave, txtMaxValueEarlyOT.Leave, txtMaxValueLate.Leave, txtMaxValueEarlyOut.Leave, txtMaxValueOT.Leave
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnOK.Enabled = True Then
                    If Txt.Enabled = True Then
                        If String.IsNullOrEmpty(Txt.Text) = False Then
                            If Me.ObjCd.Is_Time_String(Txt.Text) = True Then
                                If Me.ObjCd.Convert_ToMinutes(Txt.Text) > 1439 Then
                                    MsgBox("Schedule in not more than 23.59 or 1439 minutes", MsgBoxStyle.Information, "Information")
                                    Txt.Text = Me.buffer
                                End If
                            Else
                                MsgBox("Please input hh.mm format", MsgBoxStyle.Information, "Information")
                                Txt.Text = Me.buffer
                            End If
                        Else
                            Txt.Text = "00.00"
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCalculateAfterEarlyOT_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCalculateAfterEarlyOT.LostFocus, txtCalculateAfterLate.LostFocus, txtCalculateAfterEarlyOut.LostFocus, txtCalculateAfterOT.LostFocus, txtRoundingEarlyOT.LostFocus, txtRoundingLate.LostFocus, txtRoundingEarlyOut.LostFocus, txtRoundingOT.LostFocus, txtMaxValueEarlyOT.LostFocus, txtMaxValueLate.LostFocus, txtMaxValueEarlyOut.LostFocus, txtMaxValueOT.LostFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnOK.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.TxtReadOnly.BackColor
                        Else
                            Txt.BackColor = Me.TxtPrimaryJob.BackColor
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Dim frmMasterWorkShift_AdvanceSettings As New FrmMasterWorkShift_AdvanceSettings
        Try
            If Me.TxtPrimaryJob.Text = "-1" Or (Me.CboIsCalculateEarlyOT.SelectedIndex = 1 And Me.CboIsCalculateLate.SelectedIndex = 1 And _
                    Me.CboIsCalculateEarlyOut.SelectedIndex = 1 And Me.CboIsCalculateOT.SelectedIndex = 1) Then
                Exit Sub
            End If
            frmMasterWorkShift_AdvanceSettings.BackColorAll = Me.BackColorAll
            frmMasterWorkShift_AdvanceSettings.UserCode = Me.UserCode
            frmMasterWorkShift_AdvanceSettings.ObjWS = Me.ObjWorkShift
            If Me.CboIsCalculateEarlyOT.SelectedIndex = 1 Then
                frmMasterWorkShift_AdvanceSettings.disabledEarlyOT = True
            End If
            If Me.CboIsCalculateLate.SelectedIndex = 1 Then
                frmMasterWorkShift_AdvanceSettings.disabledLate = True
            End If
            If Me.CboIsCalculateEarlyOut.SelectedIndex = 1 Then
                frmMasterWorkShift_AdvanceSettings.disabledEarlyOut = True
            End If
            If Me.CboIsCalculateOT.SelectedIndex = 1 Then
                frmMasterWorkShift_AdvanceSettings.disabledOT = True
            End If
            If Me.btnOK.Enabled = True Then
                frmMasterWorkShift_AdvanceSettings.ReadOnly_ = False
            End If
            If frmMasterWorkShift_AdvanceSettings.ShowDialog = Windows.Forms.DialogResult.OK Then
                With Me.ObjWorkShift
                    If frmMasterWorkShift_AdvanceSettings.chkCalculateEarlyOTOnWeekend.Checked = True Then
                        .CalculateEarlyOTOnWeekend = 1
                    Else
                        .CalculateEarlyOTOnWeekend = 0
                    End If
                    If frmMasterWorkShift_AdvanceSettings.chkCalculateLateOnWeekend.Checked = True Then
                        .CalculateLateOnWeekend = 1
                    Else
                        .CalculateLateOnWeekend = 0
                    End If
                    If frmMasterWorkShift_AdvanceSettings.chkCalculateEarlyOutOnWeekend.Checked = True Then
                        .CalculateEarlyOutOnWeekend = 1
                    Else
                        .CalculateEarlyOutOnWeekend = 0
                    End If
                    If frmMasterWorkShift_AdvanceSettings.chkCalculateOTOnWeekend.Checked = True Then
                        .CalculateOTOnWeekend = 1
                    Else
                        .CalculateOTOnWeekend = 0
                    End If
                    If frmMasterWorkShift_AdvanceSettings.chkCalculateEarlyOTOnHoliday.Checked = True Then
                        .CalculateEarlyOTOnHoliday = 1
                    Else
                        .CalculateEarlyOTOnHoliday = 0
                    End If
                    If frmMasterWorkShift_AdvanceSettings.chkCalculateLateOnHoliday.Checked = True Then
                        .CalculateLateOnHoliday = 1
                    Else
                        .CalculateLateOnHoliday = 0
                    End If
                    If frmMasterWorkShift_AdvanceSettings.chkCalculateEarlyOutOnHoliday.Checked = True Then
                        .CalculateEarlyOutOnHoliday = 1
                    Else
                        .CalculateEarlyOutOnHoliday = 0
                    End If
                    If frmMasterWorkShift_AdvanceSettings.chkCalculateOTOnHoliday.Checked = True Then
                        .CalculateOTOnHoliday = 1
                    Else
                        .CalculateOTOnHoliday = 0
                    End If
                End With
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkUseOvertimeOrder_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkUseOvertimeOrder.CheckedChanged
        Try
            If Me.chkUseOvertimeOrder.Checked = True And Me.chkUseOvertimeOrder.Enabled = True Then
                Me.chkSendNotifEmailOvertimeOrder.Enabled = True
            ElseIf Me.chkUseOvertimeOrder.Checked = True And Me.chkUseOvertimeOrder.Enabled = False Then
                Me.chkSendNotifEmailOvertimeOrder.Enabled = False
            ElseIf Me.chkUseOvertimeOrder.Checked = False And Me.chkSendNotifEmailOvertimeOrder.Enabled = True Then
                Me.chkSendNotifEmailOvertimeOrder.Enabled = False
                Me.chkSendNotifEmailOvertimeOrder.Checked = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ComparedWithDgvHoliday(ByVal vDaysOfWeek As String, ByVal vIn As String, ByVal vOut As String, ByVal vBreakIn As String, ByVal vBreakOut As String) As Boolean
        Try
            For Each row As DataGridViewRow In Me.DgvHolidayList.Rows
                With row
                    If String.IsNullOrEmpty(.Cells(Me.ColItemHoliday_Day.Index).Value) = False Then
                        If .Cells(Me.ColItemHoliday_Day.Index).Value = vDaysOfWeek Then
                            .Cells(Me.ColItemHoliday_SchIn.Index).Value = vIn
                            .Cells(Me.ColItemHoliday_SchOut.Index).Value = vOut
                            .Cells(Me.ColSchBreakIn_Holiday.Index).Value = vBreakIn
                            .Cells(Me.ColSchBreakOut_Holiday.Index).Value = vBreakOut
                        End If
                    End If
                End With
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
End Class