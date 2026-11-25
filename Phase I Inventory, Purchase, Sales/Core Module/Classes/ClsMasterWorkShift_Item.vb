Option Explicit On
Option Strict Off

Public Class ClsMasterWorkShift_Item
    Dim ObjCd As New ClsCekData

    Public Function IsiDGVItem(ByVal vDGVWeekly As DataGridView, ByVal vDGVHolidayList As DataGridView, _
    ByVal vPrimaryWorkShift As Double, ByVal Frm As FrmMasterWorkShift, ByVal vObjCd As Core_Module.ClsCekData, _
    ByVal vPilihan As Integer, ByVal vTipeWS As ClsMasterWorkShift.EnumTipeWorkshift) As Boolean
        Dim DTAweekly As New DseMasterWorkShiftTableAdapters.tblmaster_workshift_itemTableAdapter
        Dim TBLweekly As New DseMasterWorkShift.tblmaster_workshift_itemDataTable
        Dim rowWeekly As DseMasterWorkShift.tblmaster_workshift_itemRow
        Dim rowDGVweekly As DataGridViewRow

        Dim DTAholidaylist As New DseMasterWorkShiftTableAdapters.tblmaster_workshift_itemholidayTableAdapter
        Dim TBLholidaylist As New DseMasterWorkShift.tblmaster_workshift_itemholidayDataTable
        Dim rowholidaylist As DseMasterWorkShift.tblmaster_workshift_itemholidayRow
        Dim rowDGVholidaylist As DataGridViewRow

        Dim DtaAttParams As New DseMasterWorkShiftTableAdapters.tblmaster_workshift_itemparametersTableAdapter
        Dim TblAttParams As New DseMasterWorkShift.tblmaster_workshift_itemparametersDataTable
        Dim RowAttParams As DseMasterWorkShift.tblmaster_workshift_itemparametersRow

        Dim DTASett As New DseMasterWorkShiftTableAdapters.tblmaster_workshift_newTableAdapter
        Dim tblSett As New DseMasterWorkShift.tblmaster_workshift_newDataTable
        Dim rowSett As DseMasterWorkShift.tblmaster_workshift_newRow

        Dim CboDgv As DataGridViewComboBoxCell
        Dim ObjCuti As New ClsMasterCuti
        Try
            DTAweekly.Connection.ConnectionString = Core_Module.GetConstring
            DTAholidaylist.Connection.ConnectionString = Core_Module.GetConstring
            DtaAttParams.Connection.ConnectionString = Core_Module.GetConstring
            DTASett.Connection.ConnectionString = Core_Module.GetConstring

            ''isikan(tabel)
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            'DTAweekly.FillBy_PrimaryWorksift(TBLweekly, vPrimaryWorkShift)
            If vPilihan = ClsMasterWorkShift.EnumPilihan.weekly Or vPilihan = ClsMasterWorkShift.EnumPilihan.All Then
                vDGVWeekly.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                vDGVWeekly.Rows.Clear()
                If DTAweekly.FillBy_PrimaryWorksift(TBLweekly, vPrimaryWorkShift) > 0 Then
                    For Each rowWeekly In TBLweekly.Rows
                        With rowWeekly
                            vDGVWeekly.Rows.Add("")
                            rowDGVweekly = vDGVWeekly.Rows(vDGVWeekly.Rows.Count - 1)
                            rowDGVweekly.Cells(Frm.ColDay.Index).Value = vObjCd.GetHari_English(.day_of_week)
                            CboDgv = rowDGVweekly.Cells(Frm.ColWorkdays.Index)
                            rowDGVweekly.Cells(Frm.ColWorkdays.Index).Value = CboDgv.Items.Item(.tipe_workdays)
                            CboDgv = rowDGVweekly.Cells(Frm.ColOvertime.Index)
                            rowDGVweekly.Cells(Frm.ColOvertime.Index).Value = CboDgv.Items.Item(.tipe_overtime)
                            CboDgv = rowDGVweekly.Cells(Frm.ColDuration.Index)
                            rowDGVweekly.Cells(Frm.ColDuration.Index).Value = CboDgv.Items.Item(.duratioin)
                            rowDGVweekly.Cells(Frm.ColSchIn.Index).Value = vObjCd.Convert_ToString(.schedule_in)
                            rowDGVweekly.Cells(Frm.ColSchOut.Index).Value = vObjCd.Convert_ToString(.schedule_out)
                            rowDGVweekly.Cells(Frm.ColRemarks.Index).Value = .remarks
                            rowDGVweekly.Cells(Frm.COlEnumDayOfWeek.Index).Value = .day_of_week
                            rowDGVweekly.Cells(Frm.ColSchBreakIn_WA.Index).Value = vObjCd.Convert_ToString(.schedule_break_in)
                            rowDGVweekly.Cells(Frm.ColSchBreakOut_WA.Index).Value = vObjCd.Convert_ToString(.schedule_break_out)
                        End With
                    Next
                End If
            End If


            If vPilihan = ClsMasterWorkShift.EnumPilihan.holidaylist Or vPilihan = ClsMasterWorkShift.EnumPilihan.All Then
                vDGVHolidayList.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                vDGVHolidayList.Rows.Clear()
                vDGVHolidayList.AllowUserToAddRows = False
                If DTAholidaylist.FillBy_PrimaryWorksift(TBLholidaylist, vPrimaryWorkShift) > 0 Then
                    For Each rowholidaylist In TBLholidaylist.Rows
                        With rowholidaylist
                            Me.GetSchedule(._date.DayOfWeek, Frm, TBLholidaylist, .sch_in, .sch_out, .schedule_break_in, .schedule_break_out)
                            vDGVHolidayList.Rows.Add()
                            rowDGVholidaylist = vDGVHolidayList.Rows(vDGVHolidayList.Rows.Count - 1)
                            'My.Application.ChangeCulture("id-ID")
                            'My.Application.ChangeUICulture("id-ID")
                            rowDGVholidaylist.Cells(Frm.ColItemHoliday_Date.Index).Value = Format(._date, "dd MMM yyyy")
                            'My.Application.ChangeCulture("en-US")
                            'My.Application.ChangeUICulture("en-US")
                            rowDGVholidaylist.Cells(Frm.ColItemHoliday_Day.Index).Value = vObjCd.GetHari_English(._date.DayOfWeek)
                            CboDgv = rowDGVholidaylist.Cells(Frm.ColItemHoliday_Duration.Index)
                            rowDGVholidaylist.Cells(Frm.ColItemHoliday_Duration.Index).Value = CboDgv.Items.Item(.duration)
                            rowDGVholidaylist.Cells(Frm.ColItemHoliday_SchIn.Index).Value = vObjCd.Convert_ToString(.sch_in)
                            rowDGVholidaylist.Cells(Frm.ColItemHoliday_SchOut.Index).Value = vObjCd.Convert_ToString(.sch_out)
                            rowDGVholidaylist.Cells(Frm.ColItemHoliday_Description.Index).Value = .holiday_description
                            CboDgv = rowDGVholidaylist.Cells(Frm.ColItemHoliday_Overtime.Index)
                            rowDGVholidaylist.Cells(Frm.ColItemHoliday_Overtime.Index).Value = CboDgv.Items.Item(.tipe_overtime)
                            rowDGVholidaylist.Cells(Frm.ColItemHoliday_Leave.Index).Value = ObjCuti.GetCodeByPrimary(.primary_leave)
                            rowDGVholidaylist.Cells(Frm.ColItemHoliday_PrimaryLeave.Index).Value = .primary_leave
                            rowDGVholidaylist.Cells(Frm.ColSchBreakIn_Holiday.Index).Value = vObjCd.Convert_ToString(.schedule_break_in)
                            rowDGVholidaylist.Cells(Frm.ColSchBreakOut_Holiday.Index).Value = vObjCd.Convert_ToString(.schedule_break_out)
                            rowDGVholidaylist.Cells(Frm.ColItemHoliday_Remarks.Index).Value = .remarks
                        End With
                    Next
                End If
                If Frm.btnOK.Enabled = True Then
                    vDGVHolidayList.AllowUserToAddRows = True
                End If
            End If

            'Frm.CboIsCalculateEarlyOT.SelectedIndex = 1
            'Frm.txtCalculateAfterEarlyOT.Text = Me.ObjCd.Convert_ToString(0)
            'Frm.cboRoundingTypeEarlyOT.SelectedIndex = ClsMasterWorkShift.enumRoundingType.No_Rounding
            'Frm.txtRoundingEarlyOT.Text = Me.ObjCd.Convert_ToString(0)
            'Frm.txtMaxValueEarlyOT.Text = Me.ObjCd.Convert_ToString(0)
            'Frm.CboIsCalculateLate.SelectedIndex = 1
            'Frm.txtCalculateAfterLate.Text = Me.ObjCd.Convert_ToString(0)
            'Frm.cboRoundingTypeLate.SelectedIndex = ClsMasterWorkShift.enumRoundingType.No_Rounding
            'Frm.txtRoundingLate.Text = Me.ObjCd.Convert_ToString(0)
            'Frm.txtMaxValueLate.Text = Me.ObjCd.Convert_ToString(0)
            'Frm.CboIsCalculateEarlyOut.SelectedIndex = 1
            'Frm.txtCalculateAfterLate.Text = Me.ObjCd.Convert_ToString(0)
            'Frm.cboRoundingTypeLate.SelectedIndex = ClsMasterWorkShift.enumRoundingType.No_Rounding
            'Frm.txtRoundingLate.Text = Me.ObjCd.Convert_ToString(0)
            'Frm.txtMaxValueLate.Text = Me.ObjCd.Convert_ToString(0)
            'Frm.CboIsCalculateOT.SelectedIndex = 1
            'Frm.txtCalculateAfterOT.Text = Me.ObjCd.Convert_ToString(0)
            'Frm.cboRoundingTypeOT.SelectedIndex = ClsMasterWorkShift.enumRoundingType.No_Rounding
            'Frm.txtRoundingOT.Text = Me.ObjCd.Convert_ToString(0)
            'Frm.txtMaxValueOT.Text = Me.ObjCd.Convert_ToString(0)

            If vTipeWS <> ClsMasterWorkShift.EnumTipeWorkshift.Production And (vPilihan = ClsMasterWorkShift.EnumPilihan.AttendanceSetting Or vPilihan = ClsMasterWorkShift.EnumPilihan.All) Then
                DTASett.FillBy_Primary(tblSett, vPrimaryWorkShift)
                For Each rowSett In tblSett.Rows
                    With rowSett
                        If .is_calculate_early_ot = 0 Then
                            Frm.CboIsCalculateEarlyOT.SelectedIndex = 1
                            Frm.txtCalculateAfterEarlyOT.Text = Me.ObjCd.Convert_ToString(0)
                            Frm.cboRoundingTypeEarlyOT.SelectedIndex = ClsMasterWorkShift.enumRoundingType.No_Rounding
                            Frm.txtRoundingEarlyOT.Text = Me.ObjCd.Convert_ToString(0)
                            Frm.txtMaxValueEarlyOT.Text = Me.ObjCd.Convert_ToString(0)
                            Frm.chkCalculateEarlyOTTolerance.Enabled = False
                        Else
                            Frm.CboIsCalculateEarlyOT.SelectedIndex = 0
                            Frm.txtCalculateAfterEarlyOT.Text = Me.ObjCd.Convert_ToString(.calculate_after_early_OT)
                            Frm.cboRoundingTypeEarlyOT.SelectedIndex = .rounding_type_early_ot
                            Frm.txtRoundingEarlyOT.Text = Me.ObjCd.Convert_ToString(.rounding_value_early_ot)
                            Frm.txtMaxValueEarlyOT.Text = Me.ObjCd.Convert_ToString(.max_value_early_ot)
                            If Frm.TxtRemark.Enabled = True Then Frm.chkCalculateEarlyOTTolerance.Enabled = True
                        End If


                        If .is_calculate_late = 0 Then
                            Frm.CboIsCalculateLate.SelectedIndex = 1
                            Frm.txtCalculateAfterLate.Text = Me.ObjCd.Convert_ToString(0)
                            Frm.cboRoundingTypeLate.SelectedIndex = ClsMasterWorkShift.enumRoundingType.No_Rounding
                            Frm.txtRoundingLate.Text = Me.ObjCd.Convert_ToString(0)
                            Frm.txtMaxValueLate.Text = Me.ObjCd.Convert_ToString(0)
                            Frm.chkCalculateLateTolerance.Enabled = False
                        Else
                            Frm.CboIsCalculateLate.SelectedIndex = 0
                            Frm.txtCalculateAfterLate.Text = Me.ObjCd.Convert_ToString(.calculate_after_late)
                            Frm.cboRoundingTypeLate.SelectedIndex = .rounding_type_late
                            Frm.txtRoundingLate.Text = Me.ObjCd.Convert_ToString(.rounding_value_late)
                            Frm.txtMaxValueLate.Text = Me.ObjCd.Convert_ToString(.max_value_early_ot)
                            If Frm.TxtRemark.Enabled = True Then Frm.chkCalculateLateTolerance.Enabled = True
                        End If


                        If .is_calculate_early_out = 0 Then
                            Frm.CboIsCalculateEarlyOut.SelectedIndex = 1
                            Frm.txtCalculateAfterLate.Text = Me.ObjCd.Convert_ToString(0)
                            Frm.cboRoundingTypeLate.SelectedIndex = ClsMasterWorkShift.enumRoundingType.No_Rounding
                            Frm.txtRoundingLate.Text = Me.ObjCd.Convert_ToString(0)
                            Frm.txtMaxValueLate.Text = Me.ObjCd.Convert_ToString(0)
                            Frm.chkCalculateEarlyOutTolerance.Enabled = False
                        Else
                            Frm.CboIsCalculateEarlyOut.SelectedIndex = 0
                            Frm.txtCalculateAfterEarlyOut.Text = Me.ObjCd.Convert_ToString(.calculate_after_early_out)
                            Frm.cboRoundingTypeEarlyOut.SelectedIndex = .rounding_type_early_out
                            Frm.txtRoundingEarlyOut.Text = Me.ObjCd.Convert_ToString(.rounding_value_early_out)
                            Frm.txtMaxValueEarlyOut.Text = Me.ObjCd.Convert_ToString(.max_value_early_out)
                            If Frm.TxtRemark.Enabled = True Then Frm.chkCalculateEarlyOutTolerance.Enabled = True
                        End If

                        If .is_calculate_ot = 0 Then
                            Frm.CboIsCalculateOT.SelectedIndex = 1
                            Frm.txtCalculateAfterOT.Text = Me.ObjCd.Convert_ToString(0)
                            Frm.cboRoundingTypeOT.SelectedIndex = ClsMasterWorkShift.enumRoundingType.No_Rounding
                            Frm.txtRoundingOT.Text = Me.ObjCd.Convert_ToString(0)
                            Frm.txtMaxValueOT.Text = Me.ObjCd.Convert_ToString(0)
                            Frm.chkCalculateOTTolerance.Enabled = False
                        Else
                            Frm.CboIsCalculateOT.SelectedIndex = 0
                            Frm.txtCalculateAfterOT.Text = Me.ObjCd.Convert_ToString(.calculate_after_ot)
                            Frm.cboRoundingTypeOT.SelectedIndex = .rounding_type_ot
                            Frm.txtRoundingOT.Text = Me.ObjCd.Convert_ToString(.rounding_value_ot)
                            Frm.txtMaxValueOT.Text = Me.ObjCd.Convert_ToString(.max_value_ot)
                            If Frm.TxtRemark.Enabled = True Then Frm.chkCalculateOTTolerance.Enabled = True
                        End If

                        If .calculate_early_ot_tolerance = 1 Then
                            Frm.chkCalculateEarlyOTTolerance.Checked = True
                        Else
                            Frm.chkCalculateEarlyOTTolerance.Checked = False
                        End If
                        If .calculate_late_tolerance = 1 Then
                            Frm.chkCalculateLateTolerance.Checked = True
                        Else
                            Frm.chkCalculateLateTolerance.Checked = False
                        End If
                        If .calculate_early_out_tolerance = 1 Then
                            Frm.chkCalculateEarlyOutTolerance.Checked = True
                        Else
                            Frm.chkCalculateEarlyOutTolerance.Checked = False
                        End If
                        If .calculate_ot_tolerance = 1 Then
                            Frm.chkCalculateOTTolerance.Checked = True
                        Else
                            Frm.chkCalculateOTTolerance.Checked = False
                        End If
                        If .use_ot_order = 1 Then
                            Frm.chkUseOvertimeOrder.Checked = True
                        Else
                            Frm.chkUseOvertimeOrder.Checked = False
                        End If
                        If .notification_email_ot = 1 Then
                            Frm.chkSendNotifEmailOvertimeOrder.Checked = True
                        Else
                            Frm.chkSendNotifEmailOvertimeOrder.Checked = False
                        End If
                    End With
                Next

                If DtaAttParams.FillBy_PrimaryTransaksi(TblAttParams, vPrimaryWorkShift) > 0 Then
                    Frm.DgvAttParams.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                    Frm.DgvAttParams.Rows.Clear()
                    For Each RowAttParams In TblAttParams
                        Frm.DgvAttParams.Rows.Add()
                        rowDGVweekly = Frm.DgvAttParams.Rows(Frm.DgvAttParams.Rows.Count - 1)
                        With rowDGVweekly
                            If RowAttParams.schedule_type = ClsMasterWorkShift.EnumScheduleType.SCHEDULE_IN Then
                                .Cells(Frm.ColAtt_Schedule.Index).Value = "Schedule In"
                            ElseIf RowAttParams.schedule_type = ClsMasterWorkShift.EnumScheduleType.BREAK_OUT Then
                                .Cells(Frm.ColAtt_Schedule.Index).Value = "Break Out"
                            ElseIf RowAttParams.schedule_type = ClsMasterWorkShift.EnumScheduleType.BREAK_IN Then
                                .Cells(Frm.ColAtt_Schedule.Index).Value = "Break In"
                            ElseIf RowAttParams.schedule_type = ClsMasterWorkShift.EnumScheduleType.SCHEDULE_OUT Then
                                .Cells(Frm.ColAtt_Schedule.Index).Value = "Schedule Out"
                            End If
                            If RowAttParams.is_get_value = 0 Then
                                .Cells(Frm.ColAtt_GetValue.Index).Value = Boolean.FalseString
                            Else
                                .Cells(Frm.ColAtt_GetValue.Index).Value = Boolean.TrueString
                            End If
                            .Cells(Frm.ColAtt_LimitBefore.Index).Value = Me.ObjCd.DigitGrouping(RowAttParams.limit_before, 0)
                            .Cells(Frm.ColAtt_LimitAfter.Index).Value = Me.ObjCd.DigitGrouping(RowAttParams.limit_after, 0)
                            If RowAttParams.first_or_last = 0 Then
                                .Cells(Frm.ColAtt_FirstOrLast.Index).Value = "First"
                            Else
                                .Cells(Frm.ColAtt_FirstOrLast.Index).Value = "Last"
                            End If
                            If RowAttParams._default = ClsMasterWorkShift.EnumDefaultIfNoClockInOut.SCHEDULE Then
                                .Cells(Frm.ColAtt_Default.Index).Value = "Schedule"
                            ElseIf RowAttParams._default = ClsMasterWorkShift.EnumDefaultIfNoClockInOut.CUSTOM Then
                                .Cells(Frm.ColAtt_Default.Index).Value = "Custom"
                            End If

                            'aul -penambahan outstation-
                            .Cells(Frm.ColAtt_DefaultCustom.Index).Value = RowAttParams.default_custom
                            If RowAttParams.default_outstation = ClsMasterWorkShift.EnumDefaultIfOutstation.SCHEDULE Then
                                .Cells(Frm.ColAtt_Outstation.Index).Value = "Schedule"
                            ElseIf RowAttParams.default_outstation = ClsMasterWorkShift.EnumDefaultIfOutstation.CUSTOM Then
                                .Cells(Frm.ColAtt_Outstation.Index).Value = "Custom"
                            End If

                            If RowAttParams.schedule_type = ClsMasterWorkShift.EnumScheduleType.SCHEDULE_IN Then
                                .Cells(Frm.ColAtt_ScheduleType.Index).Value = ClsMasterWorkShift.EnumScheduleType.SCHEDULE_IN.ToString
                            ElseIf RowAttParams.schedule_type = ClsMasterWorkShift.EnumScheduleType.BREAK_OUT Then
                                .Cells(Frm.ColAtt_ScheduleType.Index).Value = ClsMasterWorkShift.EnumScheduleType.BREAK_OUT.ToString
                            ElseIf RowAttParams.schedule_type = ClsMasterWorkShift.EnumScheduleType.BREAK_IN Then
                                .Cells(Frm.ColAtt_ScheduleType.Index).Value = ClsMasterWorkShift.EnumScheduleType.BREAK_IN.ToString
                            ElseIf RowAttParams.schedule_type = ClsMasterWorkShift.EnumScheduleType.SCHEDULE_OUT Then
                                .Cells(Frm.ColAtt_ScheduleType.Index).Value = ClsMasterWorkShift.EnumScheduleType.SCHEDULE_OUT.ToString
                            End If
                        End With
                    Next
                End If

            End If

            vObjCd.SetNumber(vDGVHolidayList)

            vDGVWeekly.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVHolidayList.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Frm.DgvAttParams.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Windows.Forms.Cursor.Current = Cursors.Default

            'nilai balikan 
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAweekly.Dispose()
            TBLweekly = Nothing
            DTAholidaylist.Dispose()
            TBLholidaylist = Nothing
        End Try
    End Function

    Public Sub SimpanItem(ByVal vDgvWeekly As DataGridView, ByVal vDgvHoliday As DataGridView, _
    ByVal vPrimaryworkshift As Double, ByVal frm As FrmMasterWorkShift)
        Dim DtaWeekly As New DseMasterWorkShiftTableAdapters.tblmaster_workshift_itemTableAdapter
        Dim Dtaholiday As New DseMasterWorkShiftTableAdapters.tblmaster_workshift_itemholidayTableAdapter
        Dim DtaAttParams As New DseMasterWorkShiftTableAdapters.tblmaster_workshift_itemparametersTableAdapter
        Dim row As DataGridViewRow
        Dim CboDgv As DataGridViewComboBoxCell
        Dim Remarks As String = ""
        Dim i As Integer = 1
        Dim idxWorkday, idxOvertime, idxDuration As Integer
        Dim tgl As Date = #1/1/1950#
        Dim BreakDuration As Double = 0

        Dim ScheduleType_ As ClsMasterWorkShift.EnumScheduleType
        Dim IsGet_, FirstOrLast_ As Integer
        Dim DefaultIfNoClockInOut_ As ClsMasterWorkShift.EnumDefaultIfNoClockInOut
        Dim DefaultIfNoClockInOutOutstation_ As ClsMasterWorkShift.EnumDefaultIfOutstation

        Try
            Windows.Forms.Cursor.Current = Cursors.Default
            DtaWeekly.Connection.ConnectionString = Core_Module.GetConstring
            Dtaholiday.Connection.ConnectionString = Core_Module.GetConstring
            DtaAttParams.Connection.ConnectionString = Core_Module.GetConstring

            For Each row In vDgvWeekly.Rows
                With row
                    CboDgv = .Cells(frm.ColWorkdays.Index)
                    idxWorkday = CboDgv.Items.IndexOf(.Cells(frm.ColWorkdays.Index).Value)
                    CboDgv = .Cells(frm.ColOvertime.Index)
                    idxOvertime = CboDgv.Items.IndexOf(.Cells(frm.ColOvertime.Index).Value)
                    CboDgv = .Cells(frm.ColDuration.Index)
                    idxDuration = CboDgv.Items.IndexOf(.Cells(frm.ColDuration.Index).Value)
                    If String.IsNullOrEmpty(.Cells(frm.ColRemarks.Index).Value) = False Then
                        Remarks = .Cells(frm.ColRemarks.Index).Value
                    Else
                        Remarks = ""
                    End If
                    If Me.ObjCd.Convert_ToMinutes(.Cells(frm.ColSchBreakIn_WA.Index).Value) > Me.ObjCd.Convert_ToMinutes(.Cells(frm.ColSchBreakOut_WA.Index).Value) Then
                        BreakDuration = Me.ObjCd.Convert_ToMinutes(.Cells(frm.ColSchBreakIn_WA.Index).Value) - Me.ObjCd.Convert_ToMinutes(.Cells(frm.ColSchBreakOut_WA.Index).Value)
                    Else
                        ''rosadin-r4
                        BreakDuration = 0
                    End If
                    DtaWeekly.Insert(vPrimaryworkshift, .Cells(frm.COlEnumDayOfWeek.Index).Value, idxWorkday, _
                     idxOvertime, Me.ObjCd.Convert_ToMinutes(.Cells(frm.ColSchIn.Index).Value), _
                     Me.ObjCd.Convert_ToMinutes(.Cells(frm.ColSchOut.Index).Value), Remarks, _
                    idxDuration, Me.ObjCd.Convert_ToMinutes(.Cells(frm.ColSchBreakOut_WA.Index).Value), _
                    Me.ObjCd.Convert_ToMinutes(.Cells(frm.ColSchBreakIn_WA.Index).Value), BreakDuration)
                End With
            Next

            For Each row In vDgvHoliday.Rows
                'ahdar
                If i <> vDgvHoliday.Rows.Count Then
                    With row
                        CboDgv = .Cells(frm.ColItemHoliday_Duration.Index)
                        idxDuration = CboDgv.Items.IndexOf(.Cells(frm.ColItemHoliday_Duration.Index).Value)
                        CboDgv = .Cells(frm.ColItemHoliday_Overtime.Index)
                        idxOvertime = CboDgv.Items.IndexOf(.Cells(frm.ColItemHoliday_Overtime.Index).Value)
                        If String.IsNullOrEmpty(.Cells(frm.ColItemHoliday_Remarks.Index).Value) = False Then
                            Remarks = .Cells(frm.ColItemHoliday_Remarks.Index).Value
                        Else
                            Remarks = ""
                        End If
                        My.Application.ChangeCulture("id-ID")
                        My.Application.ChangeUICulture("id-ID")
                        tgl = .Cells(frm.ColItemHoliday_Date.Index).Value
                        My.Application.ChangeCulture("en-US")
                        My.Application.ChangeUICulture("en-US")
                        If Me.ObjCd.Convert_ToMinutes(.Cells(frm.ColSchBreakIn_Holiday.Index).Value) - Me.ObjCd.Convert_ToMinutes(.Cells(frm.ColSchBreakOut_Holiday.Index).Value) Then
                            BreakDuration = Me.ObjCd.Convert_ToMinutes(.Cells(frm.ColSchBreakIn_Holiday.Index).Value) - Me.ObjCd.Convert_ToMinutes(.Cells(frm.ColSchBreakOut_Holiday.Index).Value)
                        Else
                        End If
                        Dtaholiday.Insert(vPrimaryworkshift, tgl, .Cells(frm.ColItemHoliday_Description.Index).Value, _
                         .Cells(frm.ColItemHoliday_PrimaryLeave.Index).Value, idxOvertime, Remarks, idxDuration, _
                         Me.ObjCd.Convert_ToMinutes(.Cells(frm.ColItemHoliday_SchIn.Index).Value), _
                         Me.ObjCd.Convert_ToMinutes(.Cells(frm.ColItemHoliday_SchOut.Index).Value), _
                         Me.ObjCd.Convert_ToMinutes(.Cells(frm.ColSchBreakOut_Holiday.Index).Value), _
                         Me.ObjCd.Convert_ToMinutes(.Cells(frm.ColSchBreakIn_Holiday.Index).Value), BreakDuration)
                        i = i + 1
                    End With
                End If
            Next

            For Each row In frm.DgvAttParams.Rows
                With row
                    If .Cells(frm.ColAtt_ScheduleType.Index).Value.ToString.ToUpper = ClsMasterWorkShift.EnumScheduleType.SCHEDULE_IN.ToString Then
                        ScheduleType_ = ClsMasterWorkShift.EnumScheduleType.SCHEDULE_IN
                    ElseIf .Cells(frm.ColAtt_ScheduleType.Index).Value.ToString.ToUpper = ClsMasterWorkShift.EnumScheduleType.BREAK_OUT.ToString Then
                        ScheduleType_ = ClsMasterWorkShift.EnumScheduleType.BREAK_OUT
                    ElseIf .Cells(frm.ColAtt_ScheduleType.Index).Value.ToString.ToUpper = ClsMasterWorkShift.EnumScheduleType.BREAK_IN.ToString Then
                        ScheduleType_ = ClsMasterWorkShift.EnumScheduleType.BREAK_IN
                    ElseIf .Cells(frm.ColAtt_ScheduleType.Index).Value.ToString.ToUpper = ClsMasterWorkShift.EnumScheduleType.SCHEDULE_OUT.ToString Then
                        ScheduleType_ = ClsMasterWorkShift.EnumScheduleType.SCHEDULE_OUT
                    End If
                    If .Cells(frm.ColAtt_GetValue.Index).Value.ToString = Boolean.TrueString Then
                        IsGet_ = 1
                    ElseIf .Cells(frm.ColAtt_GetValue.Index).Value.ToString = Boolean.FalseString Then
                        IsGet_ = 0
                    End If
                    If .Cells(frm.ColAtt_FirstOrLast.Index).Value.ToString = "First" Then
                        FirstOrLast_ = 0
                    ElseIf .Cells(frm.ColAtt_FirstOrLast.Index).Value.ToString = "Last" Then
                        FirstOrLast_ = 1
                    End If
                    If .Cells(frm.ColAtt_Default.Index).Value = "Schedule" Then
                        DefaultIfNoClockInOut_ = ClsMasterWorkShift.EnumDefaultIfNoClockInOut.SCHEDULE
                    ElseIf .Cells(frm.ColAtt_Default.Index).Value = "Custom" Then
                        DefaultIfNoClockInOut_ = ClsMasterWorkShift.EnumDefaultIfNoClockInOut.CUSTOM
                    End If
                    If String.IsNullOrEmpty(.Cells(frm.ColAtt_DefaultCustom.Index).Value) = True Then
                        Remarks = ""
                    Else
                        Remarks = .Cells(frm.ColAtt_DefaultCustom.Index).Value.ToString
                    End If

                    'aul -penambahan outstation-
                    If .Cells(frm.ColAtt_Outstation.Index).Value = "Schedule" Then
                        DefaultIfNoClockInOutOutstation_ = ClsMasterWorkShift.EnumDefaultIfOutstation.SCHEDULE
                    ElseIf .Cells(frm.ColAtt_Outstation.Index).Value = "Custom" Then
                        DefaultIfNoClockInOutOutstation_ = ClsMasterWorkShift.EnumDefaultIfOutstation.CUSTOM
                    End If
                    DtaAttParams.Insert(vPrimaryworkshift, ScheduleType_, IsGet_, CDec(.Cells(frm.ColAtt_LimitBefore.Index).Value.ToString), _
                            CDec(.Cells(frm.ColAtt_LimitAfter.Index).Value.ToString), FirstOrLast_, DefaultIfNoClockInOut_, Remarks, DefaultIfNoClockInOutOutstation_)
                End With
            Next
            Windows.Forms.Cursor.Current = Cursors.Default
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DtaWeekly.Dispose()
            Dtaholiday.Dispose()
        End Try
    End Sub

    Public Function HapusItem(ByVal vPrimaryWorkshift As Double) As Boolean
        Dim DtaWeekly As New DseMasterWorkShiftTableAdapters.tblmaster_workshift_itemTableAdapter
        Dim DtaHolidaylist As New DseMasterWorkShiftTableAdapters.tblmaster_workshift_itemholidayTableAdapter
        Dim DtaAttParams As New DseMasterWorkShiftTableAdapters.tblmaster_workshift_itemparametersTableAdapter
        Try
            'hapus data item
            DtaWeekly.Connection.ConnectionString = Core_Module.GetConstring
            DtaHolidaylist.Connection.ConnectionString = Core_Module.GetConstring
            DtaAttParams.Connection.ConnectionString = Core_Module.GetConstring

            DtaWeekly.Delete(vPrimaryWorkshift)
            DtaHolidaylist.Delete(vPrimaryWorkshift)
            DtaAttParams.Delete(vPrimaryWorkshift)
            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DtaWeekly.Dispose()
            DtaHolidaylist.Dispose()
        End Try
    End Function

    Public Function GetSchIn(ByVal vPrimaryWS As Double, ByVal vDayOfWeek As Integer) As Double
        Dim DtaWeekly As New DseMasterWorkShiftTableAdapters.tblmaster_workshift_itemTableAdapter
        Dim Dummy As System.Nullable(Of Double)
        Try
            DtaWeekly.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DtaWeekly.ScalarQuery_GetScheduleIn(vPrimaryWS, vDayOfWeek)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DtaWeekly.Dispose()
        End Try
    End Function

    Private Function GetSchedule(ByVal vDayOfWeek As Integer, ByVal vFrm As FrmMasterWorkShift, ByVal vTbl As DataTable, ByRef vIn As Double, ByRef vOut As Double, ByRef vBreakIn As Double, ByRef vBreakOut As Double) As Double
        Dim vDgv As DataGridView = vFrm.dgvDailyActivities
        Dim Hari As String = Me.ObjCd.GetHari_English(vDayOfWeek)
        Try
            For Each rowDGV As DataGridViewRow In vDgv.Rows
                If Hari = rowDGV.Cells(vFrm.ColDay.Index).Value Then
                    vIn = Me.ObjCd.Convert_ToMinutes(rowDGV.Cells(vFrm.ColSchIn.Index).Value)
                    vOut = Me.ObjCd.Convert_ToMinutes(rowDGV.Cells(vFrm.ColSchOut.Index).Value)
                    vBreakIn = Me.ObjCd.Convert_ToMinutes(rowDGV.Cells(vFrm.ColSchBreakIn_WA.Index).Value)
                    vBreakOut = Me.ObjCd.Convert_ToMinutes(rowDGV.Cells(vFrm.ColSchBreakOut_WA.Index).Value)
                    Return True
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
End Class
