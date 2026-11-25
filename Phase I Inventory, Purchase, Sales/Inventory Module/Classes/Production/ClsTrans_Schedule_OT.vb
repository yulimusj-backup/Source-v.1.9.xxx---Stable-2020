Option Explicit On
Option Strict Off

Public Class ClsTrans_Schedule_OT

    Dim rowSchBar As DseTrans_Schedule.tbl_sch_barRow

    Public Function RecalculateAfterAddSpace(ByVal vPrimaryWorkUnit As Double, ByVal vFirstOTEndTime As Date, _
            ByVal vTBLSpaceBeforeOT As DseTrans_Schedule.tbl_sch_barDataTable, ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable) As Boolean
        Dim objSch As New ClsTrans_Schedule
        Dim TBLSpaceAfterOT As New DseTrans_Schedule.tbl_sch_barDataTable
        Dim TBLWU As New DseTrans_Schedule.tblmaster_product_item_workunit1DataTable
        Dim FirmPlan As Date = objSch.GetFirmPlan
        Dim rowSchBar0 As DseTrans_Schedule.tbl_sch_barRow
        Try
            objSch.LoadScheduleBar_TanpaScheduledItem(FirmPlan, #12/31/2200#, TBLSpaceAfterOT, vPrimaryWorkUnit)
            If vTBLSchBar.Count > 0 Then
                vTBLSchBar.DefaultView.RowFilter = "[primary work unit] = " & vPrimaryWorkUnit.ToString
                vTBLSchBar.DefaultView.Sort = "[plan start time]"
                Dim AffectedSch As Boolean = False
                Dim planStartTime As Date = #1/1/1950#
                Dim planEndTime As Date = #1/1/1950#
                Dim delayFromPreviousSch As Double = 0
                Dim timeAsReference As Date
                For i As Integer = 0 To vTBLSchBar.DefaultView.Count - 1
                    rowSchBar = vTBLSchBar.DefaultView(i).Row
                    If i = 0 Then
                        timeAsReference = FirmPlan
                    Else
                        rowSchBar0 = vTBLSchBar.DefaultView(i - 1).Row
                        timeAsReference = rowSchBar0.plan_end_time
                    End If
                    If planEndTime = #1/1/1950# Or AffectedSch = False Then planEndTime = timeAsReference
                    If timeAsReference = rowSchBar.plan_start_time Then
                        delayFromPreviousSch = 0
                    Else
                        delayFromPreviousSch = objSch.GetWorkDuration(timeAsReference, rowSchBar.plan_start_time, vPrimaryWorkUnit, vTBLSpaceBeforeOT)
                    End If
                    If rowSchBar.plan_end_time >= vFirstOTEndTime Or AffectedSch Then
                        AffectedSch = True
                        If delayFromPreviousSch = 0 Then
                            planStartTime = planEndTime
                        Else
                            planStartTime = objSch.GetEndTime(vPrimaryWorkUnit, planEndTime, delayFromPreviousSch, TBLSpaceAfterOT)
                        End If
                        planEndTime = objSch.GetEndTime(vPrimaryWorkUnit, planStartTime, rowSchBar.total_work_hour, TBLSpaceAfterOT)
                        objSch.UpdatePlanStartEndSchedule(planStartTime, planEndTime, rowSchBar.identifier_jo, rowSchBar.identifier_wo, vPrimaryWorkUnit)
                        objSch.UpdatePlanStartEndScheduleWO(rowSchBar.identifier_jo, rowSchBar.identifier_wo)
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function RecalculateAfterDeleteOT(ByVal vPrimaryWorkUnit As Double, ByVal vFirstOTEndTime As Date, _
            ByVal vTBLSpaceBeforeDeleteOT As DseTrans_Schedule.tbl_sch_barDataTable, ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable, _
            ByVal vDgvOT As DataGridView, ByRef vTBLUpdate As DseTrans_OT_Item.DataTable2DataTable) As Boolean
        Dim RowUpdate As DseTrans_OT_Item.DataTable2Row
        Dim FrmOT As New FrmTrans_work_unitionSchedule_EditOT
        Dim objSch As New ClsTrans_Schedule
        Dim TBLSpaceAfterDeleteOT As New DseTrans_Schedule.tbl_sch_barDataTable
        Dim TBLWU As New DseTrans_Schedule.tblmaster_product_item_workunit1DataTable
        Dim FirmPlan As Date = objSch.GetFirmPlan
        Dim rowSchBar0 As DseTrans_Schedule.tbl_sch_barRow

        Dim DtaJO As New DseTrans_OT_ItemTableAdapters.DataTable3TableAdapter
        Dim TblJO As New DseTrans_OT_Item.DataTable3DataTable
        Dim RowJO As DseTrans_OT_Item.DataTable3Row

        Dim DtaBundling As New DseTrans_OT_ItemTableAdapters.DataTable4TableAdapter
        Dim TblBundling As New DseTrans_OT_Item.DataTable4DataTable
        Dim RowBundling As DseTrans_OT_Item.DataTable4Row

        Try

            ''simulasi hapus OT
            TBLSpaceAfterDeleteOT = vTBLSpaceBeforeDeleteOT.Copy
            For Each RowDgv As DataGridViewRow In vDgvOT.Rows
                If RowDgv.Cells(FrmOT.ColDelete.Index).Value = Boolean.TrueString Then
                    TBLSpaceAfterDeleteOT.DefaultView.RowFilter = "[tipe] = 3 and [primary work unit] = " & RowDgv.Cells(FrmOT.ColPrimaryWU.Index).Value.ToString & " and [plan start time] = #" & Format(CDate(RowDgv.Cells(FrmOT.ColStartDate.Index).Value & " " & RowDgv.Cells(FrmOT.ColStartTime.Index).Value.ToString.Replace(".", ":") & " :00"), "MM/dd/yyyy HH:mm:ss") & "#"
                    While TBLSpaceAfterDeleteOT.DefaultView.Count > 0
                        rowSchBar0 = TBLSpaceAfterDeleteOT.DefaultView(0).Row
                        TBLSpaceAfterDeleteOT.Rows.Remove(rowSchBar0)
                    End While
                End If
            Next

            ''simulasi hitung ulang
            If vTBLSchBar.Count > 0 Then
                vTBLSchBar.DefaultView.RowFilter = "[primary work unit] = " & vPrimaryWorkUnit.ToString
                vTBLSchBar.DefaultView.Sort = "[plan start time]"
                Dim AffectedSch As Boolean = False
                Dim planStartTime As Date = #1/1/1950#
                Dim planEndTime As Date = #1/1/1950#
                Dim delayFromPreviousSch As Double = 0
                Dim timeAsReference As Date
                For i As Integer = 0 To vTBLSchBar.DefaultView.Count - 1
                    rowSchBar = vTBLSchBar.DefaultView(i).Row
                    If i = 0 Then
                        timeAsReference = FirmPlan
                    Else
                        rowSchBar0 = vTBLSchBar.DefaultView(i - 1).Row
                        timeAsReference = rowSchBar0.plan_end_time
                    End If
                    If planEndTime = #1/1/1950# Or AffectedSch = False Then planEndTime = timeAsReference
                    If timeAsReference = rowSchBar.plan_start_time Then
                        delayFromPreviousSch = 0
                    Else
                        delayFromPreviousSch = objSch.GetWorkDuration(timeAsReference, rowSchBar.plan_start_time, vPrimaryWorkUnit, vTBLSpaceBeforeDeleteOT)
                    End If
                    If rowSchBar.plan_end_time >= vFirstOTEndTime Or AffectedSch Then
                        AffectedSch = True
                        If delayFromPreviousSch = 0 Then
                            planStartTime = planEndTime
                        Else
                            planStartTime = objSch.GetEndTime(vPrimaryWorkUnit, planEndTime, delayFromPreviousSch, TBLSpaceAfterDeleteOT)
                        End If
                        planEndTime = objSch.GetEndTime(vPrimaryWorkUnit, planStartTime, rowSchBar.total_work_hour, TBLSpaceAfterDeleteOT)
                        RowUpdate = vTBLUpdate.NewDataTable2Row
                        RowUpdate.Identifier_JO = rowSchBar.identifier_jo
                        RowUpdate.Identifier_WO = rowSchBar.identifier_wo
                        RowUpdate.Primary_Workunit = vPrimaryWorkUnit
                        RowUpdate.Plan_Start_Time = planStartTime
                        RowUpdate.Plan_End_Time = planEndTime
                        RowUpdate.Workunit = rowSchBar.work_unit_code
                        vTBLUpdate.Rows.Add(RowUpdate)
                    End If
                Next
            End If

            ''update JO state
            DtaJO.Connection.ConnectionString = Core_Module.GetConstring
            DtaBundling.Connection.ConnectionString = Core_Module.GetConstring
            DtaJO.Fill(TblJO)
            DtaBundling.Fill(TblBundling, #1/1/1950#)
            For Each RowUpdate In vTBLUpdate
                TblJO.DefaultView.RowFilter = "[Identifier] = '" & RowUpdate.Identifier_JO & "'"
                If TblJO.DefaultView.Count > 0 Then
                    RowJO = TblJO.DefaultView(0).Row
                    RowUpdate.JO_Closed = True
                    RowUpdate.JO_No = RowJO.JO_Number
                End If
                TblBundling.DefaultView.RowFilter = "[identifier jo] = '" & RowUpdate.Identifier_JO & "' and [identifier wo]='" & RowUpdate.Identifier_WO & "'and [primary work unit] = " & RowUpdate.Primary_Workunit.ToString
                If TblBundling.DefaultView.Count > 0 Then
                    RowBundling = TblBundling.DefaultView(0).Row
                    RowUpdate.Bundling = True
                    RowUpdate.JO_No = RowBundling.JO_Number
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function RecalculateAfterDeleteAllocation(ByVal vPrimaryWorkUnit As Double, ByVal vFirstOTEndTime As Date, _
            ByVal vTBLSpaceBeforeDeleteAllocation As DseTrans_Schedule.tbl_sch_barDataTable, ByVal vTBLSchBar As DseTrans_Schedule.tbl_sch_barDataTable, _
            ByRef vTBLUpdate As DseTrans_OT_Item.DataTable2DataTable, _
            ByVal vStartDate As Date, ByVal vEndDate As Date, ByVal vPrimaryWS As Double) As Boolean
        Dim RowUpdate As DseTrans_OT_Item.DataTable2Row
        Dim FrmOT As New FrmTrans_work_unitionSchedule_EditOT
        Dim objSch As New ClsTrans_Schedule
        Dim TBLSpaceAfterDeleteAllocation As New DseTrans_Schedule.tbl_sch_barDataTable
        Dim TBLWU As New DseTrans_Schedule.tblmaster_product_item_workunit1DataTable
        Dim FirmPlan As Date = objSch.GetFirmPlan
        Dim rowSchBar0 As DseTrans_Schedule.tbl_sch_barRow

        Dim DtaJO As New DseTrans_OT_ItemTableAdapters.DataTable3TableAdapter
        Dim TblJO As New DseTrans_OT_Item.DataTable3DataTable
        Dim RowJO As DseTrans_OT_Item.DataTable3Row

        Dim DtaBundling As New DseTrans_OT_ItemTableAdapters.DataTable4TableAdapter
        Dim TblBundling As New DseTrans_OT_Item.DataTable4DataTable
        Dim RowBundling As DseTrans_OT_Item.DataTable4Row

        Dim dtaAllocation As New DseTrans_ScheduleTableAdapters.tbltrans_mo_workshift_allocate1TableAdapter
        Dim tblAllocation As New DseTrans_Schedule.tbltrans_mo_workshift_allocate1DataTable
        Dim rowAllocation As DseTrans_Schedule.tbltrans_mo_workshift_allocate1Row

        Dim tblAllocation2 As New DseTrans_Schedule.DataTable9DataTable
        Dim rowAllocation2 As DseTrans_Schedule.DataTable9Row

        Dim dtaWorkshift As New DseTrans_ScheduleTableAdapters.tblmaster_workshift_item1TableAdapter
        Dim tblWorkshift As New DseTrans_Schedule.tblmaster_workshift_item1DataTable
        Dim rowWorkshift As DseTrans_Schedule.tblmaster_workshift_item1Row
        Try

            ''simulasi hapus OT
            dtaWorkshift.Connection.ConnectionString = Core_Module.GetConstring
            dtaWorkshift.Fill(tblWorkshift, Core_Module.ClsMasterWorkShift.EnumTipeWorkshift.Production)
            dtaAllocation.Connection.ConnectionString = Core_Module.GetConstring
            dtaAllocation.Fill(tblAllocation, vStartDate, vEndDate, vPrimaryWorkUnit, vPrimaryWS)
            For Each rowAllocation In tblAllocation
                tblWorkshift.DefaultView.RowFilter = "[primary transaksi] = " & rowAllocation.primary_workshift.ToString & " and [day of week] = " & rowAllocation._date.Date.DayOfWeek
                If tblWorkshift.DefaultView.Count > 0 Then
                    rowWorkshift = tblWorkshift.DefaultView(0).Row
                    rowAllocation2 = tblAllocation2.NewDataTable9Row
                    rowAllocation2.primary_workunit = rowAllocation.primary_work_unit
                    rowAllocation2.primary_workshift = rowAllocation.primary_workshift
                    rowAllocation2.day_of_week = rowAllocation._date.DayOfWeek
                    rowAllocation2.before_break = True
                    rowAllocation2.start = rowAllocation._date.Date.AddMinutes(rowWorkshift.schedule_in)
                    rowAllocation2._end = rowAllocation._date.AddMinutes(rowWorkshift.schedule_break_out)
                    tblAllocation2.Rows.Add(rowAllocation2)
                    rowAllocation2 = tblAllocation2.NewDataTable9Row
                    rowAllocation2.primary_workunit = rowAllocation.primary_work_unit
                    rowAllocation2.primary_workshift = rowAllocation.primary_workshift
                    rowAllocation2.day_of_week = rowAllocation._date.DayOfWeek
                    rowAllocation2.before_break = False
                    rowAllocation2.start = rowAllocation._date.Date.AddMinutes(rowWorkshift.schedule_break_in)
                    rowAllocation2._end = rowAllocation._date.AddMinutes(rowWorkshift.schedule_out)
                    tblAllocation2.Rows.Add(rowAllocation2)
                End If
            Next
            TBLSpaceAfterDeleteAllocation = vTBLSpaceBeforeDeleteAllocation.Copy
            For Each rowAllocation2 In tblAllocation2
                TBLSpaceAfterDeleteAllocation.DefaultView.RowFilter = "[tipe] = 0 and [primary work unit] = " & vPrimaryWorkUnit.ToString & " and [plan start time] = #" & Format(rowAllocation2.start, "MM/dd/yyyy HH:mm") & "# and [plan end time]=#" & Format(rowAllocation2._end, "MM/dd/yyyy HH:mm") & "#"
                While TBLSpaceAfterDeleteAllocation.DefaultView.Count > 0
                    rowSchBar0 = TBLSpaceAfterDeleteAllocation.DefaultView(0).Row
                    rowSchBar0.tipe = ClsTrans_Schedule.EnumTipeSchBar.off_time
                End While
            Next

            ''simulasi hitung ulang
            If vTBLSchBar.Count > 0 Then
                vTBLSchBar.DefaultView.RowFilter = "[primary work unit] = " & vPrimaryWorkUnit.ToString
                vTBLSchBar.DefaultView.Sort = "[plan start time]"
                Dim AffectedSch As Boolean = False
                Dim planStartTime As Date = #1/1/1950#
                Dim planEndTime As Date = #1/1/1950#
                Dim delayFromPreviousSch As Double = 0
                Dim timeAsReference As Date
                For i As Integer = 0 To vTBLSchBar.DefaultView.Count - 1
                    rowSchBar = vTBLSchBar.DefaultView(i).Row
                    If i = 0 Then
                        timeAsReference = FirmPlan
                    Else
                        rowSchBar0 = vTBLSchBar.DefaultView(i - 1).Row
                        timeAsReference = rowSchBar0.plan_end_time
                    End If
                    If planEndTime = #1/1/1950# Or AffectedSch = False Then planEndTime = timeAsReference
                    If timeAsReference = rowSchBar.plan_start_time Then
                        delayFromPreviousSch = 0
                    Else
                        delayFromPreviousSch = objSch.GetWorkDuration(timeAsReference, rowSchBar.plan_start_time, vPrimaryWorkUnit, vTBLSpaceBeforeDeleteAllocation)
                    End If
                    If rowSchBar.plan_end_time >= vFirstOTEndTime Or AffectedSch Then
                        AffectedSch = True
                        If delayFromPreviousSch = 0 Then
                            planStartTime = planEndTime
                        Else
                            planStartTime = objSch.GetEndTime(vPrimaryWorkUnit, planEndTime, delayFromPreviousSch, TBLSpaceAfterDeleteAllocation)
                        End If
                        planEndTime = objSch.GetEndTime(vPrimaryWorkUnit, planStartTime, rowSchBar.total_work_hour, TBLSpaceAfterDeleteAllocation)
                        RowUpdate = vTBLUpdate.NewDataTable2Row
                        RowUpdate.Identifier_JO = rowSchBar.identifier_jo
                        RowUpdate.Identifier_WO = rowSchBar.identifier_wo
                        RowUpdate.Primary_Workunit = vPrimaryWorkUnit
                        RowUpdate.Plan_Start_Time = planStartTime
                        RowUpdate.Plan_End_Time = planEndTime
                        RowUpdate.Workunit = rowSchBar.work_unit_code
                        vTBLUpdate.Rows.Add(RowUpdate)
                    End If
                Next
            End If

            ''update JO state
            DtaJO.Connection.ConnectionString = Core_Module.GetConstring
            DtaBundling.Connection.ConnectionString = Core_Module.GetConstring
            DtaJO.Fill(TblJO)
            DtaBundling.Fill(TblBundling, #1/1/1950#)
            For Each RowUpdate In vTBLUpdate
                TblJO.DefaultView.RowFilter = "[Identifier] = '" & RowUpdate.Identifier_JO & "'"
                If TblJO.DefaultView.Count > 0 Then
                    RowJO = TblJO.DefaultView(0).Row
                    RowUpdate.JO_Closed = True
                    RowUpdate.JO_No = RowJO.JO_Number
                End If
                TblBundling.DefaultView.RowFilter = "[identifier jo] = '" & RowUpdate.Identifier_JO & "' and [identifier wo]='" & RowUpdate.Identifier_WO & "'and [primary work unit] = " & RowUpdate.Primary_Workunit.ToString & " and [end time actual] > #" & Format(vStartDate.Date, "MM/dd/yyyy") & "#"
                If TblBundling.DefaultView.Count > 0 Then
                    RowBundling = TblBundling.DefaultView(0).Row
                    RowUpdate.Bundling = True
                    RowUpdate.JO_No = RowBundling.JO_Number
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function EnoughAllocation(ByVal vTBLUpdate As DseTrans_OT_Item.DataTable2DataTable) As Boolean
        Dim ListOfErr As New List(Of String)
        Dim ErrCode As String = ""
        Dim objWU As New Core_Module.ClsMasterWorkUnit
        Dim RowUpdate As DseTrans_OT_Item.DataTable2Row
        Try
            vTBLUpdate.DefaultView.RowFilter = "[plan end time] >= #1/1/2200#"
            If vTBLUpdate.DefaultView.Count > 0 Then
                For Each Row As Data.DataRowView In vTBLUpdate.DefaultView
                    RowUpdate = Row.Row
                    If Not ListOfErr.Contains(RowUpdate.Primary_Workunit) Then
                        ListOfErr.Add(RowUpdate.Primary_Workunit)
                        If ErrCode = "" Then
                            ErrCode = ErrCode & objWU.GetCode(RowUpdate.Primary_Workunit)
                        Else
                            ErrCode = ErrCode & ", " & objWU.GetCode(RowUpdate.Primary_Workunit)
                        End If
                    End If
                Next
                MsgBox("Can't delete OT data, not enough work shift allocation for these work unit:" & vbNewLine & ErrCode & ".", MsgBoxStyle.Information, "Information")
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function AnyClosedJO(ByVal vTBLUpdate As DseTrans_OT_Item.DataTable2DataTable) As Boolean
        Dim ListOfErr As New List(Of String)
        Dim ErrCode As String = ""
        Dim RowUpdate As DseTrans_OT_Item.DataTable2Row
        Try
            vTBLUpdate.DefaultView.RowFilter = "[JO Closed] = True"
            If vTBLUpdate.DefaultView.Count > 0 Then
                For Each Row As Data.DataRowView In vTBLUpdate.DefaultView
                    RowUpdate = Row.Row
                    If Not ListOfErr.Contains(RowUpdate.JO_No) Then
                        ListOfErr.Add(RowUpdate.JO_No)
                        If ErrCode = "" Then
                            ErrCode = ErrCode & RowUpdate.JO_No
                        Else
                            ErrCode = ErrCode & ", " & RowUpdate.JO_No
                        End If
                    End If
                Next
                MsgBox("Can't delete OT data, the following JO are closed:" & vbNewLine & ErrCode & ".", MsgBoxStyle.Information, "Information")
                Return True
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function AnyProducedJO(ByVal vTBLUpdate As DseTrans_OT_Item.DataTable2DataTable, ByRef vMessage As String) As Boolean
        Dim ListOfErr As New List(Of String)
        Dim ErrCode As String = ""
        Dim RowUpdate As DseTrans_OT_Item.DataTable2Row
        Try
            vTBLUpdate.DefaultView.RowFilter = "[Bundling] = True"
            If vTBLUpdate.DefaultView.Count > 0 Then
                For Each Row As Data.DataRowView In vTBLUpdate.DefaultView
                    RowUpdate = Row.Row
                    If Not ListOfErr.Contains(RowUpdate.JO_No) Then
                        ListOfErr.Add(RowUpdate.JO_No)
                        If ErrCode = "" Then
                            ErrCode = ErrCode & RowUpdate.JO_No
                        Else
                            ErrCode = ErrCode & ", " & RowUpdate.JO_No
                        End If
                    End If
                Next
                vMessage = ErrCode
                Return True
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function AnyClosedJO(ByVal vTblSchBar As DseTrans_Schedule.tbl_sch_barDataTable, ByVal vStartDate As Date, ByVal vListPrimaryWorkunit As List(Of Double)) As Boolean
        Dim rowSchBar As DseTrans_Schedule.tbl_sch_barRow
        Dim DtaJO As New DseTrans_OT_ItemTableAdapters.DataTable3TableAdapter
        Dim TblJO As New DseTrans_OT_Item.DataTable3DataTable
        Dim rowJO As DseTrans_OT_Item.DataTable3Row
        Dim ListOfErr As New List(Of String)
        Dim ErrCode As String = ""

        Try
            DtaJO.Connection.ConnectionString = Core_Module.GetConstring
            DtaJO.Fill(TblJO)
            For Each PrimaryWorkunit As Double In vListPrimaryWorkunit
                vTblSchBar.DefaultView.RowFilter = "[primary work unit] = " & PrimaryWorkunit.ToString & " and [plan start time] > #" & Format(vStartDate.AddDays(-1), "MM/dd/yyyy") & "#"
                If vTblSchBar.DefaultView.Count > 0 Then
                    For Each DV As Data.DataRowView In vTblSchBar.DefaultView
                        rowSchBar = DV.Row
                        TblJO.DefaultView.RowFilter = "[Identifier] = '" & rowSchBar.identifier_jo & "'"
                        If TblJO.DefaultView.Count > 0 Then
                            rowJO = TblJO.DefaultView(0).Row
                            If Not ListOfErr.Contains(rowJO.JO_Number) Then
                                ListOfErr.Add(rowJO.JO_Number)
                                If ErrCode = "" Then
                                    ErrCode = ErrCode & rowJO.JO_Number
                                Else
                                    ErrCode = ErrCode & ", " & rowJO.JO_Number
                                End If
                            End If
                        End If
                    Next
                End If
            Next
            If Not String.IsNullOrEmpty(ErrCode) Then
                MsgBox("Can't add OT data, the following JO are closed:" & vbNewLine & ErrCode & ".", MsgBoxStyle.Information, "Information")
                Return True
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function AnyProducedJO(ByVal vTblSchBar As DseTrans_Schedule.tbl_sch_barDataTable, ByVal vStartDate As Date, ByVal vListPrimaryWorkunit As List(Of Double), _
    ByRef vMessage As String) As Boolean
        Dim rowSchBar As DseTrans_Schedule.tbl_sch_barRow
        Dim DtaBundling As New DseTrans_OT_ItemTableAdapters.DataTable4TableAdapter
        Dim TblBundling As New DseTrans_OT_Item.DataTable4DataTable
        Dim RowBundling As DseTrans_OT_Item.DataTable4Row
        Dim ListOfErr As New List(Of String)
        Dim ErrCode As String = ""

        Try
            DtaBundling.Connection.ConnectionString = Core_Module.GetConstring
            DtaBundling.Fill(TblBundling, #1/1/1950#)
            For Each PrimaryWorkunit As Double In vListPrimaryWorkunit
                vTblSchBar.DefaultView.RowFilter = "[primary work unit] = " & PrimaryWorkunit.ToString & " and [plan end time] > #" & Format(vStartDate.AddDays(-1), "MM/dd/yyyy") & "#"
                If vTblSchBar.DefaultView.Count > 0 Then
                    For Each DV As Data.DataRowView In vTblSchBar.DefaultView
                        rowSchBar = DV.Row
                        TblBundling.DefaultView.RowFilter = "[identifier jo] = '" & rowSchBar.identifier_jo & "' and [identifier wo] = '" & rowSchBar.identifier_wo & "' and [end time actual] > #" & Format(vStartDate.Date, "MM/dd/yyyy") & "#"
                        If TblBundling.DefaultView.Count > 0 Then
                            RowBundling = TblBundling.DefaultView(0).Row
                            If Not ListOfErr.Contains(RowBundling.JO_Number) Then
                                ListOfErr.Add(RowBundling.JO_Number)
                                If ErrCode = "" Then
                                    ErrCode = ErrCode & RowBundling.JO_Number
                                Else
                                    ErrCode = ErrCode & ", " & RowBundling.JO_Number
                                End If
                            End If
                        End If
                    Next
                End If
            Next
            If Not String.IsNullOrEmpty(ErrCode) Then
                vMessage = ErrCode
                Return True
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function AnyProducedJO(ByVal vTblSchBar As DseTrans_Schedule.tbl_sch_barDataTable, ByVal vStartDate As Date, ByVal vListPrimaryWorkunit As List(Of Double), _
    ByRef vMessage As String, ByVal vEndDate As Date) As Boolean
        Dim rowSchBar As DseTrans_Schedule.tbl_sch_barRow
        Dim DtaBundling As New DseTrans_OT_ItemTableAdapters.DataTable4TableAdapter
        Dim TblBundling As New DseTrans_OT_Item.DataTable4DataTable
        Dim RowBundling As DseTrans_OT_Item.DataTable4Row
        Dim ListOfErr As New List(Of String)
        Dim ErrCode As String = ""

        Try
            DtaBundling.Connection.ConnectionString = Core_Module.GetConstring
            DtaBundling.FillBy_start_enddate(TblBundling, vStartDate, vEndDate)
            For Each PrimaryWorkunit As Double In vListPrimaryWorkunit
                vTblSchBar.DefaultView.RowFilter = "[primary work unit] = " & PrimaryWorkunit.ToString & " and [plan end time] > #" & Format(vStartDate.AddDays(-1), "MM/dd/yyyy") & "#"
                If vTblSchBar.DefaultView.Count > 0 Then
                    For Each DV As Data.DataRowView In vTblSchBar.DefaultView
                        rowSchBar = DV.Row
                        TblBundling.DefaultView.RowFilter = "[identifier jo] = '" & rowSchBar.identifier_jo & "' and [identifier wo] = '" & rowSchBar.identifier_wo & "' and [end time actual] > #" & Format(vStartDate.Date, "MM/dd/yyyy") & "#"
                        If TblBundling.DefaultView.Count > 0 Then
                            RowBundling = TblBundling.DefaultView(0).Row
                            If Not ListOfErr.Contains(RowBundling.JO_Number) Then
                                ListOfErr.Add(RowBundling.JO_Number)
                                If ErrCode = "" Then
                                    ErrCode = ErrCode & RowBundling.JO_Number
                                Else
                                    ErrCode = ErrCode & ", " & RowBundling.JO_Number
                                End If
                            End If
                        End If
                    Next
                End If
            Next
            If Not String.IsNullOrEmpty(ErrCode) Then
                vMessage = ErrCode
                Return True
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function UpdateScheduleChange(ByVal vTBLUpdate As DseTrans_OT_Item.DataTable2DataTable) As Integer
        Dim DataChangeCount As Integer = 0
        Dim RowUpdate As DseTrans_OT_Item.DataTable2Row
        Dim ObjSch As New ClsTrans_Schedule
        Try
            For Each RowUpdate In vTBLUpdate
                DataChangeCount += ObjSch.UpdatePlanStartEndSchedule(RowUpdate.Plan_Start_Time, RowUpdate.Plan_End_Time, RowUpdate.Identifier_JO, RowUpdate.Identifier_WO, RowUpdate.Primary_Workunit)
                ObjSch.UpdatePlanStartEndScheduleWO(RowUpdate.Identifier_JO, RowUpdate.Identifier_WO)
            Next
            Return DataChangeCount
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return DataChangeCount
        End Try
    End Function
End Class
