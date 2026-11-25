Option Explicit On
Option Strict Off
Imports System.Windows.Forms

Public Class ClsTrans_Attendance_Item
    Dim ObjCd As New Core_Module.ClsCekData

    '''mengembalikan nilai 1 jika ada item attendance untuk transaksi tsb
    '''mengembalikan nilai 0 jika tidak ada item attendance untuk transaksi tsb
    Public Function GetJml(ByVal vPrimaryKary As Double, ByVal vDate As Date) As Double
        Dim DTA As New DseTrans_AttendanceTableAdapters.tbltrans_attendanceTableAdapter
        Dim Jml As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Jml = DTA.ScalarQuery_GetJml(vDate, vPrimaryKary)

            If Jml.HasValue = True Then
                Return Jml.Value
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

    Public Function GetMaxAttendanceDate(ByVal vPrimaryKaryawan As Double) As Date
        Dim DTA As New DseTrans_AttendanceTableAdapters.DataTable5TableAdapter
        Dim Retval As Nullable(Of Date)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetMaxAttendanceDate(vPrimaryKaryawan)
            If Retval.HasValue = True Then
                Return Retval.Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return #1/1/1950#
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Sub IsiOvertimeOrderData(ByVal vPrimaryOvertimeOrder As Double, ByVal rowD As DataGridViewRow, _
    ByVal TblOT_Order As DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_ot_1DataTable, ByVal i As FrmTrans_Attendance)
        Dim rowOT_Order As DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_ot_1Row

        Try
            If vPrimaryOvertimeOrder <> -1 Then
                TblOT_Order.DefaultView.RowFilter = "[Date] = #" & CDate(rowD.Cells(i.ColDate.Index).Value) & "#"
                If TblOT_Order.DefaultView.Count > 0 Then
                    rowOT_Order = TblOT_Order.DefaultView(0).Row
                    rowD.Cells(i.ColPrimaryOvertimeOrder.Index).Value = vPrimaryOvertimeOrder
                    rowD.Cells(i.ColOvertimeOrderNo.Index).Value = rowOT_Order.No_Transaksi
                    rowD.Cells(i.ColEarlyOTOrderDuration.Index).Value = Me.ObjCd.Convert_ToString(rowOT_Order.OT_Early_Approved)
                    rowD.Cells(i.ColOTOrderDuration.Index).Value = Me.ObjCd.Convert_ToString(rowOT_Order.OT_Aprroved)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub IsiDgv(ByVal vPrimaryTrans As Double, ByVal vFrm As FrmTrans_Attendance, _
    ByVal TblOT_Order As DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_ot_1DataTable)
        Dim Dta As New DseTrans_AttendanceTableAdapters.DataTable5TableAdapter
        Dim tbl As New DseTrans_Attendance.DataTable5DataTable
        Dim row As DseTrans_Attendance.DataTable5Row
        Dim rowOT_Order As DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_ot_1Row
        Dim RowDgv As System.Windows.Forms.DataGridViewRow
        Dim CboDgv As System.Windows.Forms.DataGridViewComboBoxCell
        Dim ShowLog As Boolean = False

        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            Dta.FillBy_PrimaryTransaksi(tbl, vPrimaryTrans)
            vFrm.dgvItem.Rows.Clear()

            vFrm.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vFrm.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            For Each row In tbl.Rows
                vFrm.dgvItem.Rows.Add()
                RowDgv = vFrm.dgvItem.Rows(vFrm.dgvItem.Rows.Count - 1)
                RowDgv.Cells(vFrm.ColNo.Index).Value = vFrm.dgvItem.Rows.Count
                'My.Application.ChangeCulture("id-ID")
                'My.Application.ChangeUICulture("id-ID")
                RowDgv.Cells(vFrm.ColDate.Index).Value = Format(row._date, "dd MMM yyyy")
                'My.Application.ChangeCulture("en-US")
                'My.Application.ChangeUICulture("en-US")
                RowDgv.Cells(vFrm.ColWorkshift.Index).Value = row.Code_WS
                CboDgv = RowDgv.Cells(vFrm.ColWorkdays.Index)
                RowDgv.Cells(vFrm.ColWorkdays.Index).Value = CboDgv.Items.Item(row.Tipe_Workdays)
                CboDgv = RowDgv.Cells(vFrm.ColDuration.Index)
                RowDgv.Cells(vFrm.ColDuration.Index).Value = CboDgv.Items.Item(row.duration)
                CboDgv = RowDgv.Cells(vFrm.ColAttendanceStatus.Index)
                RowDgv.Cells(vFrm.ColAttendanceStatus.Index).Value = CboDgv.Items.Item(row.presence_status)
                RowDgv.Cells(vFrm.ColSchIn.Index).Value = Me.ObjCd.Convert_ToString(row.sch_in)
                RowDgv.Cells(vFrm.ColSchBreakOut.Index).Value = Me.ObjCd.Convert_ToString(row.sch_break_out)
                RowDgv.Cells(vFrm.ColSchBreakIn.Index).Value = Me.ObjCd.Convert_ToString(row.sch_break_in)
                RowDgv.Cells(vFrm.ColSchOut.Index).Value = Me.ObjCd.Convert_ToString(row.sch_out)
                RowDgv.Cells(vFrm.ColActualIn.Index).Value = Me.ObjCd.Convert_ToString(row.actual_in)
                RowDgv.Cells(vFrm.ColActualBreakOut.Index).Value = Me.ObjCd.Convert_ToString(row.actual_break_out)
                RowDgv.Cells(vFrm.ColActualBreakIn.Index).Value = Me.ObjCd.Convert_ToString(row.actual_break_in)
                RowDgv.Cells(vFrm.ColActualOut.Index).Value = Me.ObjCd.Convert_ToString(row.actual_out)
                RowDgv.Cells(vFrm.ColStatusRemark.Index).Value = row.presence_status_remarks
                RowDgv.Cells(vFrm.ColLeaveCode.Index).Value = row.Code_Cuti
                RowDgv.Cells(vFrm.ColLeaveDesc.Index).Value = row.Desc_cuti
                RowDgv.Cells(vFrm.ColLeaveDoc.Index).Value = row.Cuti_app
                RowDgv.Cells(vFrm.ColEarlyInDuration.Index).Value = Me.ObjCd.Convert_ToString(row.early_in_duration)
                RowDgv.Cells(vFrm.ColLateDuration.Index).Value = Me.ObjCd.Convert_ToString(row.late_duration)
                RowDgv.Cells(vFrm.ColEarlyOutDuration.Index).Value = Me.ObjCd.Convert_ToString(row.early_out_duration)
                RowDgv.Cells(vFrm.ColEarlyInDuration.Index).Value = Me.ObjCd.Convert_ToString(row.early_in_duration)
                CboDgv = RowDgv.Cells(vFrm.ColOT.Index)
                RowDgv.Cells(vFrm.ColOT.Index).Value = CboDgv.Items.Item(row.tipe_overtime)
                RowDgv.Cells(vFrm.ColOTDuration.Index).Value = Me.ObjCd.Convert_ToString(row.ot_duration)
                RowDgv.Cells(vFrm.ColEarlyOTDuration.Index).Value = Me.ObjCd.Convert_ToString(row.early_ot_duration)
                RowDgv.Cells(vFrm.ColEarlyBreakOutDuration.Index).Value = Me.ObjCd.Convert_ToString(row.early_break_out_duration)
                RowDgv.Cells(vFrm.ColLateBreakInDuration.Index).Value = Me.ObjCd.Convert_ToString(row.late_break_in_duration)
                RowDgv.Cells(vFrm.ColRemarks.Index).Value = row.remarks
                RowDgv.Cells(vFrm.ColIsLate.Index).Value = row.is_late
                RowDgv.Cells(vFrm.ColIsEarlyIn.Index).Value = row.is_early_in
                RowDgv.Cells(vFrm.ColIsEarlyOut.Index).Value = row.is_early_out
                RowDgv.Cells(vFrm.ColIsEarlyBreakOut.Index).Value = row.is_early_break_out
                RowDgv.Cells(vFrm.ColIsLateBreakIn.Index).Value = row.is_late_break_in
                RowDgv.Cells(vFrm.ColSchWorkDuration.Index).Value = Me.ObjCd.Convert_ToString(row.sch_work_duration)
                RowDgv.Cells(vFrm.ColSchBreakDuration.Index).Value = Me.ObjCd.Convert_ToString(row.sch_break_duration)
                RowDgv.Cells(vFrm.ColActualWorkDuration.Index).Value = Me.ObjCd.Convert_ToString(row.actual_work_duration)
                RowDgv.Cells(vFrm.ColActualBreakDuration.Index).Value = Me.ObjCd.Convert_ToString(row.actual_break_duration)
                RowDgv.Cells(vFrm.ColIsPresent.Index).Value = row.is_present
                RowDgv.Cells(vFrm.ColIsSick.Index).Value = row.is_sick
                RowDgv.Cells(vFrm.ColIsLeave.Index).Value = row.is_leave
                RowDgv.Cells(vFrm.ColIsAlfa.Index).Value = row.Is_alfa
                RowDgv.Cells(vFrm.ColIsPermit.Index).Value = row.is_permit
                RowDgv.Cells(vFrm.ColIsOvertime.Index).Value = row.is_overtime
                RowDgv.Cells(vFrm.ColIsEarlyOvertime.Index).Value = row.is_early_overtime
                RowDgv.Cells(vFrm.ColIsHoliday.Index).Value = row.is_holiday
                RowDgv.Cells(vFrm.ColPrimaryWorkShift.Index).Value = row.primary_workshift
                RowDgv.Cells(vFrm.ColPrimaryLeave.Index).Value = row.primary_cuti
                RowDgv.Cells(vFrm.ColPrimaryLA.Index).Value = row.Primary_Leave_App
                RowDgv.Cells(vFrm.ColPermitOut.Index).Value = Me.ObjCd.Convert_ToString(row.Permit_out)
                RowDgv.Cells(vFrm.ColPermitIn.Index).Value = Me.ObjCd.Convert_ToString(row.Permit_in)
                RowDgv.Cells(vFrm.ColPermitDuration.Index).Value = Me.ObjCd.Convert_ToString(row.Permit_duration)
                If row.Permit_duration > 0 Then
                    RowDgv.Cells(vFrm.ColPermit.Index).Value = True
                Else
                    RowDgv.Cells(vFrm.ColPermit.Index).Value = False
                End If
                RowDgv.Cells(vFrm.ColIsOnHolidayList.Index).Value = row.is_on_holiday_list
                RowDgv.Cells(vFrm.ColLogContents.Index).Value = row.log_content
                RowDgv.Cells(vFrm.ColShowLog.Index).Value = "Show"
                If row.log_content.Trim <> "" Then
                    ShowLog = True
                End If
                If row.is_calculate_early_ot = 0 Then
                    RowDgv.Cells(vFrm.ColIsCalculateEarlyOT.Index).Value = False
                Else
                    RowDgv.Cells(vFrm.ColIsCalculateEarlyOT.Index).Value = True
                End If
                If row.is_calculate_late = 0 Then
                    RowDgv.Cells(vFrm.ColIsCalculateLate.Index).Value = False
                Else
                    RowDgv.Cells(vFrm.ColIsCalculateLate.Index).Value = True
                End If
                If row.is_calculate_early_out = 0 Then
                    RowDgv.Cells(vFrm.ColIsCalculateEarlyOut.Index).Value = False
                Else
                    RowDgv.Cells(vFrm.ColIsCalculateEarlyOut.Index).Value = True
                End If
                If row.is_calculate_ot = 0 Then
                    RowDgv.Cells(vFrm.ColIsCalculateOT.Index).Value = False
                Else
                    RowDgv.Cells(vFrm.ColIsCalculateOT.Index).Value = True
                End If

                RowDgv.Cells(vFrm.ColCalculateAfterEarlyOT.Index).Value = Me.ObjCd.Convert_ToString(row.calculate_after_early_OT)
                RowDgv.Cells(vFrm.ColCalculateAfterEarlyLate.Index).Value = Me.ObjCd.Convert_ToString(row.calculate_after_late)
                RowDgv.Cells(vFrm.ColCalculateAfterEarlyOut.Index).Value = Me.ObjCd.Convert_ToString(row.calculate_after_early_out)
                RowDgv.Cells(vFrm.ColCalculateAfterOT.Index).Value = Me.ObjCd.Convert_ToString(row.calculate_after_ot)

                RowDgv.Cells(vFrm.ColRoundingValueEarlyOT.Index).Value = Me.ObjCd.Convert_ToString(row.rounding_value_early_ot)
                RowDgv.Cells(vFrm.ColRoundingValueLate.Index).Value = Me.ObjCd.Convert_ToString(row.rounding_value_late)
                RowDgv.Cells(vFrm.ColRoundingValueEarlyOut.Index).Value = Me.ObjCd.Convert_ToString(row.rounding_value_early_out)
                RowDgv.Cells(vFrm.ColRoundingValueOT.Index).Value = Me.ObjCd.Convert_ToString(row.rounding_value_ot)

                RowDgv.Cells(vFrm.ColMaxValueEarlyOT.Index).Value = Me.ObjCd.Convert_ToString(row.max_value_early_ot)
                RowDgv.Cells(vFrm.ColMaxValueLate.Index).Value = Me.ObjCd.Convert_ToString(row.max_value_late)
                RowDgv.Cells(vFrm.ColMaxValueEarlyOut.Index).Value = Me.ObjCd.Convert_ToString(row.max_value_early_out)
                RowDgv.Cells(vFrm.ColMaxValueOT.Index).Value = Me.ObjCd.Convert_ToString(row.max_value_ot)

                CboDgv = RowDgv.Cells(vFrm.ColRoundingTypeEarlyOT.Index)
                RowDgv.Cells(vFrm.ColRoundingTypeEarlyOT.Index).Value = CboDgv.Items.Item(row.rounding_type_early_ot)
                RowDgv.Cells(vFrm.ColRoundingTypeLate.Index).Value = CboDgv.Items.Item(row.rounding_type_late)
                RowDgv.Cells(vFrm.ColRoundingTypeEarlyOut.Index).Value = CboDgv.Items.Item(row.rounding_type_early_out)
                RowDgv.Cells(vFrm.ColRoundingTypeOT.Index).Value = CboDgv.Items.Item(row.rounding_type_ot)

                If row.is_outstation = 0 Then
                    RowDgv.Cells(vFrm.ColIsOutstation.Index).Value = False
                Else
                    RowDgv.Cells(vFrm.ColIsOutstation.Index).Value = True
                End If

                With RowDgv
                    .Cells(vFrm.ColCalculateEarlyOTTolerance.Index).Value = row.calculate_early_ot_tolerance
                    .Cells(vFrm.ColCalculateLateTolerance.Index).Value = row.calculate_late_tolerance
                    .Cells(vFrm.ColCalculateEarlyOutTolerance.Index).Value = row.calculate_early_out_tolerance
                    .Cells(vFrm.ColCalculateOTTolerance.Index).Value = row.calculate_ot_tolerance
                End With

                Me.IsiOvertimeOrderData(row.Primary_Overtime_Order, RowDgv, TblOT_Order, vFrm)
            Next
            If ShowLog = True Then
                vFrm.ColShowLog.Visible = True
            Else
                vFrm.ColShowLog.Visible = False
            End If
            vFrm.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vFrm.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Dta.Dispose()
            tbl.Dispose()
        End Try
    End Sub

    Public Overridable Function SimpanItem(ByVal vPrimaryTransaksi As Double, ByVal vDGV As Windows.Forms.DataGridView, _
    ByVal Frm As FrmTrans_Attendance, ByVal vList As List(Of Double), _
    ByVal TblOT_Order As DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_ot_1DataTable, _
     ByVal UpdByTbl As Boolean, ByVal vPrimaryKary As Double) As Boolean
        Dim DTA As New DseTrans_AttendanceTableAdapters.DataTable5TableAdapter
        Dim DTA1 As New DseTrans_Attendance1TableAdapters.tbltrans_attendance_itemTableAdapter
        Dim TblItem As DseTrans_Attendance.DataTable5DataTable
        Dim row As Windows.Forms.DataGridViewRow
        Dim Date_ As Date = #1/1/1950#
        Dim WorkDay As ClsTrans_Attendance.EnumWorkday
        Dim durasi As ClsTrans_Attendance.EnumDurasi
        Dim Status As ClsTrans_Attendance.EnumStatus
        Dim Ot As ClsTrans_Attendance.EnumTipeOt
        Dim CboDgv As System.Windows.Forms.DataGridViewComboBoxCell
        Dim remarks As String = ""
        Dim StatusRemark As String = ""

        Dim calculateEarlyOT, calculateLate, calculateEarlyOut, calculateOT, RTypeEOT, RTypeLate, RTypeEOut, RTypeOT, IsOutstation As Integer

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString

            vList.Clear()
            For Each row In Frm.dgvItem.Rows
                With row
                    My.Application.ChangeCulture("id-ID")
                    My.Application.ChangeUICulture("id-ID")
                    If IsDate(row.Cells(Frm.ColDate.Index).Value) = True Then
                        Date_ = row.Cells(Frm.ColDate.Index).Value
                    End If
                    My.Application.ChangeCulture("en-US")
                    My.Application.ChangeUICulture("en-US")
                    'MsgBox(Date_.ToString)
                    CboDgv = .Cells(Frm.ColWorkdays.Index)
                    WorkDay = CboDgv.Items.IndexOf(.Cells(Frm.ColWorkdays.Index).Value)
                    CboDgv = .Cells(Frm.ColDuration.Index)
                    durasi = CboDgv.Items.IndexOf(.Cells(Frm.ColDuration.Index).Value)
                    CboDgv = .Cells(Frm.ColAttendanceStatus.Index)
                    Status = CboDgv.Items.IndexOf(.Cells(Frm.ColAttendanceStatus.Index).Value)
                    CboDgv = .Cells(Frm.ColOT.Index)
                    Ot = CboDgv.Items.IndexOf(.Cells(Frm.ColOT.Index).Value)

                    If String.IsNullOrEmpty(.Cells(Frm.ColRemarks.Index).Value) = True Then
                        remarks = ""
                    Else
                        remarks = .Cells(Frm.ColRemarks.Index).Value
                    End If
                    If String.IsNullOrEmpty(.Cells(Frm.ColStatusRemark.Index).Value) = True Then
                        StatusRemark = ""
                    Else
                        StatusRemark = .Cells(Frm.ColStatusRemark.Index).Value
                    End If
                    If IsNothing(row.Cells(Frm.ColLogContents.Index).Value) = True Then
                        row.Cells(Frm.ColLogContents.Index).Value = ""
                    End If
                    If row.Cells(Frm.ColLogContents.Index).Value.ToString.Length > 250 Then
                        row.Cells(Frm.ColLogContents.Index).Value = row.Cells(Frm.ColLogContents.Index).Value.ToString.Substring(0, 240) + "...(truncated)"
                    End If

                    If row.Cells(Frm.ColIsCalculateEarlyOT.Index).Value = True Then
                        calculateEarlyOT = 1
                    Else
                        calculateEarlyOT = 0
                    End If
                    If row.Cells(Frm.ColIsCalculateLate.Index).Value = True Then
                        calculateLate = 1
                    Else
                        calculateLate = 0
                    End If
                    If row.Cells(Frm.ColIsCalculateEarlyOut.Index).Value = True Then
                        calculateEarlyOut = 1
                    Else
                        calculateEarlyOut = 0
                    End If
                    If row.Cells(Frm.ColIsCalculateOT.Index).Value = True Then
                        calculateOT = 1
                    Else
                        calculateOT = 0
                    End If
                    If row.Cells(Frm.ColIsOutstation.Index).Value = True Then
                        IsOutstation = 1
                    Else
                        IsOutstation = 0
                    End If

                    CboDgv = .Cells(Frm.ColRoundingTypeEarlyOT.Index)
                    RTypeEOT = CboDgv.Items.IndexOf(.Cells(Frm.ColRoundingTypeEarlyOT.Index).Value)
                    CboDgv = .Cells(Frm.ColRoundingTypeLate.Index)
                    RTypeLate = CboDgv.Items.IndexOf(.Cells(Frm.ColRoundingTypeLate.Index).Value)
                    CboDgv = .Cells(Frm.ColRoundingTypeEarlyOut.Index)
                    RTypeEOut = CboDgv.Items.IndexOf(.Cells(Frm.ColRoundingTypeEarlyOut.Index).Value)
                    CboDgv = .Cells(Frm.ColRoundingTypeOT.Index)
                    RTypeOT = CboDgv.Items.IndexOf(.Cells(Frm.ColRoundingTypeOT.Index).Value)

                    'MsgBox(row.Cells(Frm.ColSchBreakDuration.Index).Value)
                    DTA.InsertQuery1(vPrimaryTransaksi, .Cells(Frm.ColPrimaryWorkShift.Index).Value, Date_, _
                        Me.ObjCd.Convert_ToMinutes(.Cells(Frm.ColSchIn.Index).Value), Me.ObjCd.Convert_ToMinutes(.Cells(Frm.ColSchBreakOut.Index).Value), _
                        Me.ObjCd.Convert_ToMinutes(.Cells(Frm.ColSchBreakIn.Index).Value), Me.ObjCd.Convert_ToMinutes(.Cells(Frm.ColSchOut.Index).Value), _
                        Me.ObjCd.Convert_ToMinutes(.Cells(Frm.ColActualIn.Index).Value), Me.ObjCd.Convert_ToMinutes(.Cells(Frm.ColActualBreakOut.Index).Value), _
                        Me.ObjCd.Convert_ToMinutes(.Cells(Frm.ColActualBreakIn.Index).Value), Me.ObjCd.Convert_ToMinutes(.Cells(Frm.ColActualOut.Index).Value), _
                        remarks, .Cells(Frm.ColIsLate.Index).Value, Me.ObjCd.Convert_ToMinutes(.Cells(Frm.ColLateDuration.Index).Value), _
                        .Cells(Frm.ColIsLeave.Index).Value, .Cells(Frm.ColPrimaryLeave.Index).Value, WorkDay, Date_.DayOfWeek, durasi, Status, StatusRemark, _
                        .Cells(Frm.ColIsEarlyIn.Index).Value, Me.ObjCd.Convert_ToMinutes(.Cells(Frm.ColEarlyInDuration.Index).Value), .Cells(Frm.ColIsEarlyOut.Index).Value, _
                        Me.ObjCd.Convert_ToMinutes(.Cells(Frm.ColEarlyOutDuration.Index).Value), .Cells(Frm.ColIsOvertime.Index).Value, Ot, Me.ObjCd.Convert_ToMinutes(.Cells(Frm.ColOTDuration.Index).Value), _
                        Me.ObjCd.Convert_ToMinutes(.Cells(Frm.ColEarlyOTDuration.Index).Value), .Cells(Frm.ColIsEarlyBreakOut.Index).Value, Me.ObjCd.Convert_ToMinutes(.Cells(Frm.ColEarlyBreakOutDuration.Index).Value), _
                        .Cells(Frm.ColIsLateBreakIn.Index).Value, Me.ObjCd.Convert_ToMinutes(.Cells(Frm.ColLateBreakInDuration.Index).Value), Me.ObjCd.Convert_ToMinutes(.Cells(Frm.ColSchWorkDuration.Index).Value), _
                        Me.ObjCd.Convert_ToMinutes(.Cells(Frm.ColSchBreakDuration.Index).Value), Me.ObjCd.Convert_ToMinutes(.Cells(Frm.ColActualWorkDuration.Index).Value), _
                        Me.ObjCd.Convert_ToMinutes(.Cells(Frm.ColActualBreakDuration.Index).Value), .Cells(Frm.ColIsPresent.Index).Value, .Cells(Frm.ColIsSick.Index).Value, _
                        .Cells(Frm.ColIsAlfa.Index).Value, .Cells(Frm.ColIsPermit.Index).Value, .Cells(Frm.ColIsEarlyOvertime.Index).Value, .Cells(Frm.ColIsHoliday.Index).Value, _
                        row.Cells(Frm.ColPrimaryLA.Index).Value, .Cells(Frm.ColIsLeave.Index).Value, Me.ObjCd.Convert_ToMinutes(.Cells(Frm.ColPermitOut.Index).Value), _
                        Me.ObjCd.Convert_ToMinutes(.Cells(Frm.ColPermitIn.Index).Value), Me.ObjCd.Convert_ToMinutes(.Cells(Frm.ColPermitDuration.Index).Value), _
                        row.Cells(Frm.ColIsOnHolidayList.Index).Value, row.Cells(Frm.ColLogContents.Index).Value, calculateEarlyOT, calculateLate, calculateEarlyOut, calculateOT, _
                        Me.ObjCd.Convert_ToMinutes(.Cells(Frm.ColCalculateAfterEarlyOT.Index).Value), Me.ObjCd.Convert_ToMinutes(.Cells(Frm.ColCalculateAfterEarlyLate.Index).Value), _
                        Me.ObjCd.Convert_ToMinutes(.Cells(Frm.ColCalculateAfterEarlyOut.Index).Value), Me.ObjCd.Convert_ToMinutes(.Cells(Frm.ColCalculateAfterOT.Index).Value), _
                        RTypeEOT, RTypeLate, RTypeEOut, RTypeOT, _
                        Me.ObjCd.Convert_ToMinutes(.Cells(Frm.ColRoundingValueEarlyOT.Index).Value), Me.ObjCd.Convert_ToMinutes(.Cells(Frm.ColRoundingValueLate.Index).Value), _
                        Me.ObjCd.Convert_ToMinutes(.Cells(Frm.ColRoundingValueEarlyOut.Index).Value), Me.ObjCd.Convert_ToMinutes(.Cells(Frm.ColRoundingValueOT.Index).Value), _
                        Me.ObjCd.Convert_ToMinutes(.Cells(Frm.ColMaxValueEarlyOT.Index).Value), Me.ObjCd.Convert_ToMinutes(.Cells(Frm.ColMaxValueLate.Index).Value), _
                        Me.ObjCd.Convert_ToMinutes(.Cells(Frm.ColMaxValueEarlyOut.Index).Value), Me.ObjCd.Convert_ToMinutes(.Cells(Frm.ColMaxValueOT.Index).Value), IsOutstation, _
                        .Cells(Frm.ColCalculateEarlyOTTolerance.Index).Value, .Cells(Frm.ColCalculateLateTolerance.Index).Value, .Cells(Frm.ColCalculateEarlyOutTolerance.Index).Value, _
                        .Cells(Frm.ColCalculateOTTolerance.Index).Value, .Cells(Frm.ColPrimaryOvertimeOrder.Index).Value)
                End With

                If row.Cells(Frm.ColPrimaryLeave.Index).Value > 0 Then
                    vList.Add(row.Cells(Frm.ColPrimaryLeave.Index).Value)
                End If
            Next

            ''################## update juga ke overtime order item approval ot
            'Me.UpdateOvertimeOrder(vPrimaryTransaksi, Frm, TblOT_Order, TblItem, Frm.TxtPrimary.Text, UpdByTbl, False)
            ''################## update juga ke overtime order item approval ot
            DTA1.UpdateQuery_SetPrimaryKary(vPrimaryKary, vPrimaryTransaksi)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
        End Try
    End Function

    Public Function HapusItem(ByVal vPrimaryTransaksi As Double, ByVal Frm As FrmTrans_Attendance, _
    ByVal TblOT_Order As DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_ot_1DataTable, _
    ByVal TblItem As DseTrans_Attendance.DataTable5DataTable, ByVal vPrimaryEmp As Double, ByVal UpdByTbl As Boolean) As Boolean
        Dim dta As New DseTrans_AttendanceTableAdapters.DataTable5TableAdapter
        Dim row As Windows.Forms.DataGridViewRow

        Try
            'hapus data item
            dta.Connection.ConnectionString = Core_Module.GetConstring

            dta.DeleteQuery(vPrimaryTransaksi)

            ''################## update juga ke overtime order item approval ot
            'Me.UpdateOvertimeOrder(vPrimaryTransaksi, Frm, TblOT_Order, TblItem, vPrimaryEmp, UpdByTbl, True)
            ''################## update juga ke overtime order item approval ot

            'nilai balikan
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Sub UpdateOvertimeOrder(ByVal vPrimaryTranssaksi As Double, ByVal Frm As FrmTrans_Attendance, _
    ByVal TblOT_Order As DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_ot_1DataTable, _
    ByVal TblItem As DseTrans_Attendance.DataTable5DataTable, ByVal vPrimaryEmp As Double, ByVal UpdByTbl As Boolean, _
    ByVal SetOvertimeDefault As Boolean)
        Dim rowOT_Order As DseTrans_SPL.tbltrans_spl_overtime_order_item_approved_ot_1Row
        Dim row As Windows.Forms.DataGridViewRow
        Dim rowItem As DseTrans_Attendance.DataTable5Row
        Dim PrimarySPL As Double = -1
        Dim PrimaryAtt As Double = -1
        Dim EarlyOTDur As Double = 0
        Dim OTDur As Double = 0
        Dim ObjSplItem As New ClsTrans_SPL_Item

        Try
            If UpdByTbl = True Then
                'fungsi ini di buat karena di panggil dari hapus dari quick attendance (langsung pas primary attendance tanpa isi dgvitem dahulu.
                For Each rowItem In TblItem
                    With rowItem
                        If SetOvertimeDefault = True Then
                            PrimaryAtt = -1
                            EarlyOTDur = 0
                            OTDur = 0
                        Else
                            PrimarySPL = .Primary_Overtime_Order
                            PrimaryAtt = vPrimaryTranssaksi
                            EarlyOTDur = .early_ot_duration
                            OTDur = .ot_duration
                        End If

                        '################## update juga ke overtime order item approval ot
                        If PrimarySPL = -1 Then
                            TblOT_Order.DefaultView.RowFilter = "[Primary Employee] = " & vPrimaryEmp & " AND [Primary Shift] = " & _
                            .primary_workshift & " AND [Date] = #" & ._date & "# AND [Status] = " & ClsTrans_SPL.EnumStatus.Approved
                            If TblOT_Order.DefaultView.Count > 0 Then
                                'jika ada spl yang approved di periode terfilter dan pada item attendance primary overtime order nya -1 
                                'maka update primary attendanya saja pada tabel item approval ot nya dan duration actual early & ot nya di set 0
                                rowOT_Order = TblOT_Order.DefaultView(0).Row
                                PrimarySPL = rowOT_Order.Primary_Transaksi
                                EarlyOTDur = 0
                                OTDur = 0
                            End If
                        End If

                        If PrimarySPL <> -1 Then
                            ObjSplItem.UpdateDurationOvertime(EarlyOTDur, OTDur, PrimaryAtt, _
                            PrimarySPL, .primary_workshift, vPrimaryEmp, ._date)
                        End If
                        '################## update juga ke overtime order item approval ot
                    End With
                Next
            Else
                For Each row In Frm.dgvItem.Rows
                    With row
                        If SetOvertimeDefault = True Then
                            PrimaryAtt = -1
                            EarlyOTDur = 0
                            OTDur = 0
                        Else
                            PrimarySPL = CDbl(.Cells(Frm.ColPrimaryOvertimeOrder.Index).Value)
                            PrimaryAtt = vPrimaryTranssaksi
                            EarlyOTDur = Me.ObjCd.Convert_ToMinutes(.Cells(Frm.ColEarlyOTDuration.Index).Value)
                            OTDur = Me.ObjCd.Convert_ToMinutes(.Cells(Frm.ColOTDuration.Index).Value)
                        End If

                        '################## update juga ke overtime order item approval ot
                        If PrimarySPL = -1 Then
                            TblOT_Order.DefaultView.RowFilter = "[Primary Employee] = " & CDbl(Frm.TxtPrimaryEmployee.Text) & " AND [Primary Shift] = " & _
                            CDbl(.Cells(Frm.ColPrimaryWorkShift.Index).Value) & " AND [Date] = #" & CDate(.Cells(Frm.ColDate.Index).Value) & _
                            "# AND [Status] = " & ClsTrans_SPL.EnumStatus.Approved
                            If TblOT_Order.DefaultView.Count > 0 Then
                                'jika ada spl yang approved di periode terfilter dan pada item attendance primary overtime order nya -1 
                                'maka update primary attendanya saja pada tabel item approval ot nya dan duration actual early & ot nya di set 0
                                rowOT_Order = TblOT_Order.DefaultView(0).Row
                                PrimarySPL = rowOT_Order.Primary_Transaksi
                                EarlyOTDur = 0
                                OTDur = 0
                            End If
                        End If

                        If PrimarySPL <> -1 Then
                            ObjSplItem.UpdateDurationOvertime(EarlyOTDur, OTDur, PrimaryAtt, _
                            PrimarySPL, .Cells(Frm.ColPrimaryWorkShift.Index).Value, Frm.TxtPrimaryEmployee.Text, _
                            CDate(.Cells(Frm.ColDate.Index).Value))
                        End If
                        '################## update juga ke overtime order item approval ot
                    End With
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
