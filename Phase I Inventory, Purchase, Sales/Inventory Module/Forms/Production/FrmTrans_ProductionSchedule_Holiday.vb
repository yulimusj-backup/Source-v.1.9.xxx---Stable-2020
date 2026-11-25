Option Explicit On
Option Strict Off
Public Class FrmTrans_ProductionSchedule_Holiday
    Public BackColorAll As Color
    Public IsEdited As Boolean = False
    Public UserCode As String = ""

    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjSch As New ClsTrans_Schedule
    Dim Buffer As String
    Dim TblWS As New DseTrans_ProductionSchedule_Holiday.tblmaster_workshiftDataTable
    Dim RowWS As DseTrans_ProductionSchedule_Holiday.tblmaster_workshiftRow
    Dim RowDgv As DataGridViewRow
    Dim ObjWU As New Core_Module.ClsMasterWorkUnit
    Dim InsufficientAllocation As String = ""
    Dim ThereIsError As Boolean = False

    Dim TblSchBar As New DseTrans_Schedule.tbl_sch_barDataTable
    Dim RowSchBar As DseTrans_Schedule.tbl_sch_barRow

#Region "events"
    Private Sub FrmTrans_ProductionSchedule_Holiday_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.Panel1.Visible = False
            Me.DgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.LoadWS()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwAddHoliday_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwAddHoliday.DoWork
        Try
            Me.Cursor = Cursors.WaitCursor
            Me.ObjCD.UpdateThreading()

            If Me.AddHoliday() Then
                Me.UpdateSchedule()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub BgwAddHoliday_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwAddHoliday.RunWorkerCompleted
        Try
            Me.LoadHolidayList()
            Me.DtpFrom.Value = Now.Date
            Me.DtpTo.Value = Now.Date
            Me.TxtDescription.Clear()
            Me.TxtRemarks.Clear()
            Me.BtnAdd.Focus()
            Me.ControlForm(True)
            MsgBox("Add holiday and reschedule completed.", MsgBoxStyle.Information, "Information")
            Me.IsEdited = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BgwValidationAddHoliday_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwValidationAddHoliday.DoWork
        Dim ReloadSchFromDB As Boolean = True
        Dim HolidayList As New List(Of Date)
        Dim FirmPlanned As Date
        Try
            Me.Cursor = Cursors.WaitCursor
            Me.ObjCD.UpdateThreading()
            Me.InsufficientAllocation = ""
            Me.ThereIsError = False
            FirmPlanned = Me.ObjSch.GetFirmPlan
            For i As Integer = 0 To DateDiff(DateInterval.Day, Me.DtpFrom.Value.Date, Me.DtpTo.Value.Date)
                HolidayList.Add(Me.DtpFrom.Value.Date.AddDays(i))
                Me.AddHolidaySimulation(Me.DtpFrom.Value.Date.AddDays(i), ReloadSchFromDB, HolidayList, FirmPlanned)
                If Me.InsufficientAllocation <> "" Then Exit Sub
                If Me.ThereIsError Then Exit Sub
                If ReloadSchFromDB Then ReloadSchFromDB = False
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub BgwValidationAddHoliday_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwValidationAddHoliday.RunWorkerCompleted
        Try
            Me.ControlForm(True)
            If Me.InsufficientAllocation <> "" Then
                MsgBox(Me.InsufficientAllocation, MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            If Me.ThereIsError Then Exit Sub
            If MsgBox("Are you sure to add holiday?", MsgBoxStyle.YesNo, "Information") = MsgBoxResult.Yes Then
                Control.CheckForIllegalCrossThreadCalls = False
                Me.ControlForm(False)
                Me.BgwAddHoliday.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnAdd.Click
        Try
            Me.TryAddHoliday()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnRemove.Click
        Try
            If Me.IsValidRemoveHoloday Then
                If MsgBox("Are you sure to remove selected holiday?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                    If Me.RemoveHoliday() Then
                        Control.CheckForIllegalCrossThreadCalls = False
                        Me.ControlForm(False)
                        Me.ThereIsError = False
                        Me.BgwRemoveHoliday.RunWorkerAsync()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboWorkshift_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboWorkshift.SelectedIndexChanged
        Try
            If Me.CboWorkshift.SelectedIndex > -1 Then
                Me.LoadHolidayList()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvItem.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColDelete.Index Or e.ColumnIndex = Me.ColDescription.Index Or e.ColumnIndex = Me.ColRemarks.Index Then
                Me.Buffer = Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvItem.CellEndEdit
        Try
            If Not IsNothing(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) Then
                If Not String.IsNullOrEmpty(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString) Then
                    Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.Replace("'", "").Replace("*", "").Replace("%", "")
                    If Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString = Me.Buffer Then
                        Exit Sub
                    End If
                    Me.UpdateDescRemarksToDB(sender, e)
                Else
                    If Not IsNothing(Me.Buffer) Then
                        If Not String.IsNullOrEmpty(Me.Buffer) Then
                            Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                    End If
                End If
            Else
                If Not IsNothing(Me.Buffer) Then
                    If Not String.IsNullOrEmpty(Me.Buffer) Then
                        Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItem_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DgvItem.DataError

    End Sub

    Private Sub LlbCheckAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbCheckAll.LinkClicked
        Try
            Me.CheckDgv(True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LlbUncheckAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbUncheckAll.LinkClicked
        Try
            Me.CheckDgv(False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDescription_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDescription.GotFocus
        Try
            Me.TxtDescription.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDescription_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDescription.LostFocus
        Try
            Me.TxtDescription.BackColor = Color.White
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarks_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.GotFocus
        Try
            Me.TxtRemarks.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarks_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.LostFocus
        Try
            Me.TxtRemarks.BackColor = Color.White
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
#End Region

#Region "function"
    Private Function AddHoliday() As Boolean
        Dim DtaHoliday As New DseTrans_ProductionSchedule_HolidayTableAdapters.tblmaster_workshift_itemholidayTableAdapter
        Try
            Me.Cursor = Cursors.WaitCursor
            For i As Integer = 0 To DateDiff(DateInterval.Day, Me.DtpFrom.Value.Date, Me.DtpTo.Value.Date)
                DtaHoliday.InsertQuery_NewHoliday(Me.CboWorkshift.SelectedValue, Me.DtpFrom.Value.Date.AddDays(i), Me.TxtDescription.Text, Me.TxtRemarks.Text, Core_Module.ClsMasterWorkShift.TipeOvertime.Holiday, 0, 0)
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Function

    Private Sub AddHolidaySimulation(ByVal vHolidayDate As Date, ByVal vReloadSchFromDB As Boolean, ByVal vHolidayList As List(Of Date), ByVal vFirmPlanned As Date)
        Dim DtaAllocate As New DseTrans_ProductionSchedule_HolidayTableAdapters.tbltrans_mo_workshift_allocateTableAdapter
        Dim TblAllocate As New DseTrans_ProductionSchedule_Holiday.tbltrans_mo_workshift_allocateDataTable
        Dim ListOfWorkunit As New List(Of Double)
        Dim FirstLoop As Boolean = True
        Dim TblSchBarTanpaWO_Full As New DseTrans_Schedule.tbl_sch_barDataTable
        Dim TblSchBarTanpaWO_WithHoliday As New DseTrans_Schedule.tbl_sch_barDataTable
        Dim TblSchBarTanpaWO_AfterHoliday As New DseTrans_Schedule.tbl_sch_barDataTable
        Dim MoveDuration As Double = 0
        Dim FirstPlanStartDate, BufferPlanStartTime, BufferPlanFinishTime As Date
        Try
            DtaAllocate.Connection.ConnectionString = Core_Module.GetConstring
            DtaAllocate.Fill(TblAllocate, vHolidayDate, Me.CboWorkshift.SelectedValue)

            ''load semua schedule yang mungkin bergeser karena penambahan holiday
            If vReloadSchFromDB Then
                Me.TblSchBar.Rows.Clear()
                Me.ObjSch.LoadScheduleBar_WOItemOnly_EndDateMoreThan(vHolidayDate, Me.TblSchBar)
            End If
            FirstPlanStartDate = Me.GetFirstPlanStartDate(Me.TblSchBar)
            ListOfWorkunit = Me.GetWorkunitList(Me.TblSchBar, vHolidayDate)
            ''untuk masing-masing workunit yang terlibat, cek:
            For Each PrimaryWorkunit As Double In ListOfWorkunit
                TblAllocate.DefaultView.RowFilter = "[primary work unit] = " & PrimaryWorkunit.ToString
                ''apakah pada tanggal tsb, workunit tsb menggunakan workshift yang dimaksud
                If TblAllocate.DefaultView.Count > 0 Then
                    ''jika ya, geser start time atau end time menggunakan schedule yang asumsi holiday-nya sudah ditambahkan
                    TblSchBar.DefaultView.RowFilter = "[primary work unit] = " & PrimaryWorkunit.ToString
                    TblSchBar.DefaultView.Sort = "[plan start time]"
                    ''load beberapa alokasi untuk keperluan geser start-finish:
                    Me.ObjSch.LoadScheduleBar_TanpaScheduledItem(FirstPlanStartDate, #12/31/2200#, TblSchBarTanpaWO_Full, PrimaryWorkunit)
                    TblSchBarTanpaWO_WithHoliday = TblSchBarTanpaWO_Full.Copy
                    Me.DeleteSchedule(TblSchBarTanpaWO_WithHoliday, vHolidayList)
                    Me.ObjSch.LoadScheduleBar_TanpaScheduledItem(vHolidayDate.AddDays(1).Date, #12/31/2200#, TblSchBarTanpaWO_AfterHoliday, PrimaryWorkunit)
                    FirstLoop = True
                    Me.ProgressBar1.Value = 0
                    Me.ProgressBar1.Maximum = Me.TblSchBar.DefaultView.Count
                    For Each DV As Data.DataRowView In TblSchBar.DefaultView
                        Me.RowSchBar = DV.Row
                        Me.ProgressBar1.Value += 1
                        Me.Label6.Text = "Validating, workunit: " & Me.ObjWU.GetCode(PrimaryWorkunit) & " [data " & ListOfWorkunit.IndexOf(PrimaryWorkunit) + 1 & " of " & ListOfWorkunit.Count & "], schedule: " & Me.ProgressBar1.Value & " of " & Me.ProgressBar1.Maximum & "."
                        BufferPlanStartTime = Me.RowSchBar.plan_start_time
                        BufferPlanFinishTime = Me.RowSchBar.plan_end_time
                        If FirstLoop Then
                            If Me.RowSchBar.plan_start_time > vHolidayDate.Date Then
                                ''harus geser termasuk start nya juga
                                Me.RowSchBar.plan_start_time = Me.ObjSch.GetStartTime_ByWorkUnit(TblSchBarTanpaWO_AfterHoliday, PrimaryWorkunit, vFirmPlanned)
                                MoveDuration = Me.ObjSch.GetWorkDuration(BufferPlanStartTime, Me.RowSchBar.plan_start_time, PrimaryWorkunit, TblSchBarTanpaWO_Full)
                                Me.RowSchBar.plan_end_time = Me.ObjSch.GetEndTime(PrimaryWorkunit, Me.RowSchBar.plan_end_time, MoveDuration, TblSchBarTanpaWO_Full)
                            Else
                                ''hanya geser end nya aja
                                Me.RowSchBar.plan_end_time = Me.ObjSch.GetEndTime(PrimaryWorkunit, Me.RowSchBar.plan_start_time, Me.RowSchBar.total_work_hour, TblSchBarTanpaWO_WithHoliday)
                                MoveDuration = Me.ObjSch.GetWorkDuration(BufferPlanFinishTime, Me.RowSchBar.plan_end_time, PrimaryWorkunit, TblSchBarTanpaWO_Full)
                            End If
                            FirstLoop = False
                        Else
                            ''lanjutkan geser sesuai durasi geser schedule pertama
                            Me.RowSchBar.plan_start_time = Me.ObjSch.GetEndTime(PrimaryWorkunit, Me.RowSchBar.plan_start_time, MoveDuration, TblSchBarTanpaWO_Full)
                            Me.RowSchBar.plan_end_time = Me.ObjSch.GetEndTime(PrimaryWorkunit, Me.RowSchBar.plan_end_time, MoveDuration, TblSchBarTanpaWO_Full)
                        End If
                        If Me.RowSchBar.plan_start_time = #1/1/2200# Or Me.RowSchBar.plan_end_time = #1/1/2200# Then
                            If Me.ObjSch.ApakahFreeTime(BufferPlanStartTime, PrimaryWorkunit, TblSchBarTanpaWO_Full) = False Or Me.ObjSch.ApakahFreeTime(BufferPlanFinishTime, PrimaryWorkunit, TblSchBarTanpaWO_Full) = False Then
                                ''masuk sini jika ada schedule lama yang salah (diluar work time/ot), sementara, balikin lagi ke jadwal tsb.
                                Me.RowSchBar.plan_start_time = BufferPlanStartTime
                                Me.RowSchBar.plan_end_time = BufferPlanFinishTime
                            Else
                                Me.InsufficientAllocation = "Insufficient workshift allocation, please allocate more workshift for " + Me.ObjWU.GetCode(PrimaryWorkunit) + "."
                                Exit Sub
                            End If
                        End If
                    Next
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Me.ThereIsError = True
        End Try
    End Sub

    Private Sub CheckDgv(ByVal vCheck As Boolean)
        Try
            For Each Me.RowDgv In Me.DgvItem.Rows
                If vCheck Then
                    Me.RowDgv.Cells(Me.ColDelete.Index).Value = Boolean.TrueString
                Else
                    Me.RowDgv.Cells(Me.ColDelete.Index).Value = Boolean.FalseString
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ControlForm(ByVal vEnabled As Boolean)
        Try
            If vEnabled Then
                Me.ControlBox = True
                Me.CboWorkshift.Enabled = True
                Me.TxtDescription.Enabled = True
                Me.TxtRemarks.Enabled = True
                Me.DtpFrom.Enabled = True
                Me.DtpTo.Enabled = True
                Me.BtnAdd.Enabled = True
                Me.BtnRemove.Enabled = True
                Me.LlbCheckAll.Enabled = True
                Me.LlbUncheckAll.Enabled = True
                Me.DgvItem.ReadOnly = False
                Me.Panel1.Visible = False
            Else
                Me.ControlBox = False
                Me.CboWorkshift.Enabled = False
                Me.TxtDescription.Enabled = False
                Me.TxtRemarks.Enabled = False
                Me.DtpFrom.Enabled = False
                Me.DtpTo.Enabled = False
                Me.BtnAdd.Enabled = False
                Me.BtnRemove.Enabled = False
                Me.LlbCheckAll.Enabled = False
                Me.LlbUncheckAll.Enabled = False
                Me.DgvItem.ReadOnly = True
                Me.Panel1.Visible = True
                Me.Label6.Text = "Loading..."
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub DeleteSchedule(ByRef vTblSchBar As DseTrans_Schedule.tbl_sch_barDataTable, ByVal vListOfHoliday As List(Of Date))
        Try
            For Each vDate As Date In vListOfHoliday
                'vTblSchBar.DefaultView.RowFilter = "[plan start time] >= #" & Format(vDate, "MM/dd/yyyy") & "# AND [plan start time] < #" & Format(vDate.AddDays(1), "MM/dd/yyyy") & "#" '
                vTblSchBar.DefaultView.RowFilter = "([plan start time] >= #" & Format(vDate, "MM/dd/yyyy") & "# AND [plan start time] < #" & Format(vDate.AddDays(1), "MM/dd/yyyy") & "#) OR (tipe = " + CInt(ClsTrans_Schedule.EnumTipeSchBar.off_time).ToString & " AND [plan end time] >= #" & Format(vDate, "MM/dd/yyyy") & "# AND [plan start time] < #" & Format(vDate.AddDays(1), "MM/dd/yyyy") & "#)"
                While vTblSchBar.DefaultView.Count > 0
                    Me.RowSchBar = vTblSchBar.DefaultView(0).Row
                    vTblSchBar.Rows.Remove(Me.RowSchBar)
                End While
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Function GetFirstPlanStartDate(ByVal vTblSchBar As DseTrans_Schedule.tbl_sch_barDataTable) As Date
        Try
            vTblSchBar.DefaultView.RowFilter = ""
            vTblSchBar.DefaultView.Sort = "[plan start time]"
            If vTblSchBar.DefaultView.Count > 0 Then
                Me.RowSchBar = vTblSchBar.DefaultView(0).Row
                Return Me.RowSchBar.plan_start_time.Date
            End If
            Return #1/1/1950#
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return #1/1/1950#
        End Try
    End Function

    Private Function GetWorkunitList(ByVal vTblSchBar As DseTrans_Schedule.tbl_sch_barDataTable, ByVal vHolidayDate As Date) As List(Of Double)
        GetWorkunitList = New List(Of Double)
        Try
            GetWorkunitList.Clear()
            vTblSchBar.DefaultView.RowFilter = "[plan end time] >= #" & Format(vHolidayDate, "MM/dd/yyyy") & "# and [plan start time] <= #" & Format(vHolidayDate, "MM/dd/yyyy 23:59:59") & "#"
            For Each DV As Data.DataRowView In vTblSchBar.DefaultView
                Me.RowSchBar = DV.Row
                If Not GetWorkunitList.Contains(Me.RowSchBar.primary_work_unit) Then
                    GetWorkunitList.Add(Me.RowSchBar.primary_work_unit)
                End If
            Next
            Return GetWorkunitList
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return GetWorkunitList
        End Try
    End Function

    Private Function GetWorkunitListFromAllocation(ByVal vTblAllocate As DseTrans_ProductionSchedule_Holiday.tbltrans_mo_workshift_allocateDataTable) As List(Of Double)
        Dim RowAllocate As DseTrans_ProductionSchedule_Holiday.tbltrans_mo_workshift_allocateRow
        GetWorkunitListFromAllocation = New List(Of Double)
        Try
            GetWorkunitListFromAllocation.Clear()
            For Each RowAllocate In vTblAllocate
                If Not GetWorkunitListFromAllocation.Contains(RowAllocate.primary_work_unit) Then
                    GetWorkunitListFromAllocation.Add(RowAllocate.primary_work_unit)
                End If
            Next
            Return GetWorkunitListFromAllocation
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return GetWorkunitListFromAllocation
        End Try
    End Function

    Private Function IsConflictHoliday(ByVal vFrom As Date, ByVal vTo As Date) As Boolean
        Dim DtaHoliday As New DseTrans_ProductionSchedule_HolidayTableAdapters.tblmaster_workshift_itemholidayTableAdapter
        Dim TblHoliday As New DseTrans_ProductionSchedule_Holiday.tblmaster_workshift_itemholidayDataTable
        Dim ListOfConflict As String = ""
        Try
            Me.Cursor = Cursors.WaitCursor
            DtaHoliday.Connection.ConnectionString = Core_Module.GetConstring
            If DtaHoliday.Fill(TblHoliday, Core_Module.ClsMasterWorkShift.EnumTipeWorkshift.Production, Me.CboWorkshift.SelectedValue) > 0 Then
                For i As Integer = 0 To DateDiff(DateInterval.Day, vFrom, vTo)
                    TblHoliday.DefaultView.RowFilter = "[date] = #" & Format(vFrom.AddDays(i), "MM") & "/" & Format(vFrom.AddDays(i), "dd") & "/" & Format(vFrom.AddDays(i), "yyyy") & "#"
                    If TblHoliday.DefaultView.Count > 0 Then
                        If ListOfConflict = "" Then
                            ListOfConflict = Format(vFrom.AddDays(i), "dd MMM yyyy")
                        Else
                            ListOfConflict = ListOfConflict & ", " & Format(vFrom.AddDays(i), "dd MMM yyyy")
                        End If
                    End If
                Next
            End If
            If ListOfConflict <> "" Then
                MsgBox("Holiday on the following date is already exist:" & vbNewLine & ListOfConflict & ".", MsgBoxStyle.Information, "Information")
                Return True
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Function

    Function IsValidRemoveHoloday() As Boolean
        Dim SelectedCount As Integer = 0
        Dim FirmPlanned As Date
        Try
            For Each Me.RowDgv In Me.DgvItem.Rows
                If Me.RowDgv.Cells(Me.ColDelete.Index).Value = Boolean.TrueString Then
                    SelectedCount += 1
                End If
            Next
            If SelectedCount = 0 Then
                MsgBox("Select at least one day to remove.", MsgBoxStyle.Information, "Information")
                Me.DgvItem.Focus()
                Return False
            End If
            FirmPlanned = Me.ObjSch.GetFirmPlan
            If Me.DtpFrom.Value.Date <= FirmPlanned Then
                MsgBox("Every schedule <= " + Format(FirmPlanned, "dd MMM yyyy") + " is firm planned, you can't remove holiday.", MsgBoxStyle.Information, "Information")
                Return False
            End If
            If Me.ThereAreOvertime Then
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Sub LoadHolidayList()
        Dim DtaHoliday As New DseTrans_ProductionSchedule_HolidayTableAdapters.tblmaster_workshift_itemholidayTableAdapter
        Dim TblHoliday As New DseTrans_ProductionSchedule_Holiday.tblmaster_workshift_itemholidayDataTable
        Dim RowHoliday As DseTrans_ProductionSchedule_Holiday.tblmaster_workshift_itemholidayRow
        Try
            Me.Cursor = Cursors.WaitCursor
            Me.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DgvItem.Rows.Clear()
            DtaHoliday.Connection.ConnectionString = Core_Module.GetConstring
            If DtaHoliday.Fill(TblHoliday, Core_Module.ClsMasterWorkShift.EnumTipeWorkshift.Production, Me.CboWorkshift.SelectedValue) > 0 Then
                For Each RowHoliday In TblHoliday
                    RowDgv = Me.DgvItem.Rows(Me.DgvItem.Rows.Add)
                    RowDgv.Cells(Me.ColDate.Index).Value = Format(RowHoliday._Date, "dd MMM yyyy")
                    RowDgv.Cells(Me.ColDay.Index).Value = Me.ObjCD.GetHari_English(RowHoliday._Date)
                    RowDgv.Cells(Me.ColDescription.Index).Value = RowHoliday.Holiday_Description
                    RowDgv.Cells(Me.ColNo.Index).Value = Me.DgvItem.Rows.Count.ToString
                    RowDgv.Cells(Me.ColRemarks.Index).Value = RowHoliday.Holiday_Remarks
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub LoadWS()
        Dim DtaWS As New DseTrans_ProductionSchedule_HolidayTableAdapters.tblmaster_workshiftTableAdapter
        Try
            Me.Cursor = Cursors.WaitCursor
            DtaWS.Connection.ConnectionString = Core_Module.GetConstring
            If DtaWS.Fill(TblWS, Core_Module.ClsMasterWorkShift.EnumTipeWorkshift.Production) > 0 Then
                TblWS.DefaultView.RowFilter = ""
                TblWS.DefaultView.Sort = "Code"
                Me.CboWorkshift.ValueMember = "Primary"
                Me.CboWorkshift.DisplayMember = "Code"
                Me.CboWorkshift.DataSource = TblWS.DefaultView
                Me.CboWorkshift.SelectedIndex = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Function RemoveHoliday() As Boolean
        Dim DtaHoliday As New DseTrans_ProductionSchedule_HolidayTableAdapters.tblmaster_workshift_itemholidayTableAdapter
        Try
            DtaHoliday.Connection.ConnectionString = Core_Module.GetConstring
            For Each Me.RowDgv In Me.DgvItem.Rows
                If Me.RowDgv.Cells(Me.ColDelete.Index).Value = Boolean.TrueString Then
                    DtaHoliday.DeleteQuery_Holiday(Me.CboWorkshift.SelectedValue, CDate(Me.RowDgv.Cells(Me.ColDate.Index).Value))
                End If
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Function ThereAreBundling() As Boolean
        Dim ListOfBundling As String = ""
        Dim DtaAllocate As New DseTrans_ProductionSchedule_HolidayTableAdapters.tbltrans_mo_workshift_allocateTableAdapter
        Dim TblAllocate As New DseTrans_ProductionSchedule_Holiday.tbltrans_mo_workshift_allocateDataTable
        Dim RowAllocate As DseTrans_ProductionSchedule_Holiday.tbltrans_mo_workshift_allocateRow
        Dim DtaBundling As New DseTrans_ProductionSchedule_HolidayTableAdapters.DataTable1TableAdapter
        Dim TblBundling As New DseTrans_ProductionSchedule_Holiday.DataTable1DataTable
        Dim RowBundling As DseTrans_ProductionSchedule_Holiday.DataTable1Row
        Dim Counter As Integer = 0
        Try
            Me.Cursor = Cursors.WaitCursor
            DtaAllocate.Connection.ConnectionString = Core_Module.GetConstring
            DtaBundling.Connection.ConnectionString = Core_Module.GetConstring
            For i As Integer = 0 To DateDiff(DateInterval.Day, Me.DtpFrom.Value.Date, Me.DtpTo.Value.Date)
                If DtaAllocate.Fill(TblAllocate, Me.DtpFrom.Value.Date.AddDays(i), Me.CboWorkshift.SelectedValue) > 0 Then
                    DtaBundling.Fill(TblBundling, Me.DtpFrom.Value.Date.AddDays(i), Me.DtpFrom.Value.Date.AddDays(i).AddMinutes(1439))
                    For Each RowAllocate In TblAllocate
                        TblBundling.DefaultView.RowFilter = "[primary work unit] = " & RowAllocate.primary_work_unit.ToString
                        If TblBundling.DefaultView.Count > 0 Then
                            Counter += 1
                            RowBundling = TblBundling.DefaultView(0).Row
                            If ListOfBundling = "" Then
                                ListOfBundling = Format(Me.DtpFrom.Value.Date.AddDays(i), "dd MMM yyyy") & " " & RowAllocate.work_unit_code & ": " & RowBundling.No_Transaksi
                            Else
                                If Counter = 20 Then
                                    ListOfBundling = ListOfBundling & vbNewLine & "..."
                                    GoTo Show
                                End If
                                ListOfBundling = ListOfBundling & vbNewLine & Format(Me.DtpFrom.Value.Date.AddDays(i), "dd MMM yyyy") & " " & RowAllocate.work_unit_code & ": " & RowBundling.No_Transaksi
                            End If
                        End If
                    Next
                End If
            Next
Show:
            If ListOfBundling <> "" Then
                MsgBox("There are already exist bundling transaction for the following date:" & vbNewLine & ListOfBundling & ".", MsgBoxStyle.Information, "Information")
                Return True
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
            Me.Cursor = Cursors.Default
        End Try
    End Function

    Private Function ThereAreOvertime() As Boolean
        Dim ListOfOT As String = ""
        Dim DtaAllocate As New DseTrans_ProductionSchedule_HolidayTableAdapters.tbltrans_mo_workshift_allocateTableAdapter
        Dim TblAllocate As New DseTrans_ProductionSchedule_Holiday.tbltrans_mo_workshift_allocateDataTable
        Dim RowAllocate As DseTrans_ProductionSchedule_Holiday.tbltrans_mo_workshift_allocateRow
        Dim DtaOT As New DseTrans_ProductionSchedule_HolidayTableAdapters.tbltrans_overtime_itemTableAdapter
        Dim TblOT As New DseTrans_ProductionSchedule_Holiday.tbltrans_overtime_itemDataTable
        Dim RowOT As DseTrans_ProductionSchedule_Holiday.tbltrans_overtime_itemRow
        Dim Counter As Integer = 0
        Try
            Me.Cursor = Cursors.WaitCursor
            DtaAllocate.Connection.ConnectionString = Core_Module.GetConstring
            DtaOT.Connection.ConnectionString = Core_Module.GetConstring
            For i As Integer = 0 To DateDiff(DateInterval.Day, Me.DtpFrom.Value.Date, Me.DtpTo.Value.Date)
                If DtaAllocate.Fill(TblAllocate, Me.DtpFrom.Value.Date.AddDays(i), Me.CboWorkshift.SelectedValue) > 0 Then
                    DtaOT.Fill(TblOT, Me.DtpFrom.Value.Date.AddDays(i), Me.DtpFrom.Value.Date.AddDays(i).AddMinutes(1439))
                    For Each RowAllocate In TblAllocate
                        TblOT.DefaultView.RowFilter = "[primary workunit] = " & RowAllocate.primary_work_unit.ToString
                        If TblOT.DefaultView.Count > 0 Then
                            Counter += 1
                            RowOT = TblOT.DefaultView(0).Row
                            If ListOfOT = "" Then
                                ListOfOT = Format(Me.DtpFrom.Value.Date.AddDays(i), "dd MMM yyyy") & ": " & RowAllocate.work_unit_code
                            Else
                                If Counter = 20 Then
                                    ListOfOT = ListOfOT & vbNewLine & "..."
                                    GoTo Show
                                End If
                                ListOfOT = ListOfOT & vbNewLine & Format(Me.DtpFrom.Value.Date.AddDays(i), "dd MMM yyyy") & ": " & RowAllocate.work_unit_code
                            End If
                        End If
                    Next
                End If
            Next
Show:
            If ListOfOT <> "" Then
                MsgBox("There are already exist overtime allocation for the following date:" & vbNewLine & ListOfOT & ".", MsgBoxStyle.Information, "Information")
                Return True
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Function

    Private Sub TryAddHoliday()
        Dim FirmPlanned As Date
        Try
            Me.Cursor = Cursors.WaitCursor
            ''cek validasi standar:
            If Me.TxtDescription.Text.Trim.Length = 0 Then
                MsgBox("Input holiday description.", MsgBoxStyle.Information, "Information")
                Me.TxtDescription.Focus()
                Exit Sub
            End If
            If Me.DtpFrom.Value.Date > Me.DtpTo.Value.Date Then
                MsgBox("Invalid holiday period.", MsgBoxStyle.Information, "Information")
                Me.DtpTo.Focus()
                Exit Sub
            End If
            If Me.IsConflictHoliday(Me.DtpFrom.Value.Date, Me.DtpTo.Value.Date) Then
                Exit Sub
            End If
            FirmPlanned = Me.ObjSch.GetFirmPlan
            If Me.DtpFrom.Value.Date <= FirmPlanned Then
                MsgBox("Every schedule <= " + Format(FirmPlanned, "dd MMM yyyy") + " is firm planned, you can't add holiday.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            If Me.ThereAreBundling Then
                Exit Sub
            End If

            Control.CheckForIllegalCrossThreadCalls = False
            Me.ControlForm(False)
            ''cek validasi pergeseran schedule, jika OK, lanjutkan konfirm dan update DB
            Me.BgwValidationAddHoliday.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Function UpdateDescRemarksToDB(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) As Boolean
        Dim DtaHoliday As New DseTrans_ProductionSchedule_HolidayTableAdapters.tblmaster_workshift_itemholidayTableAdapter
        Try
            DtaHoliday.Connection.ConnectionString = Core_Module.GetConstring
            If DtaHoliday.UpdateQuery_DescRemarks(Me.DgvItem.Item(Me.ColDescription.Index, e.RowIndex).Value, Me.DgvItem.Item(Me.ColRemarks.Index, e.RowIndex).Value, Me.CboWorkshift.SelectedValue, CDate(Me.DgvItem.Item(Me.ColDate.Index, e.RowIndex).Value)) > 0 Then
                Return True
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Sub UpdateSchedule()
        Try
            Me.ProgressBar1.Value = 0
            Me.ProgressBar1.Maximum = Me.TblSchBar.Rows.Count
            For Each RowSchBar In Me.TblSchBar
                Me.ProgressBar1.Value += 1
                Me.Label6.Text = "Updating schedule: data " & Me.ProgressBar1.Value & " of " & Me.ProgressBar1.Maximum & "..."
                Me.ObjSch.UpdatePlanStartEndSchedule(RowSchBar.plan_start_time, RowSchBar.plan_end_time, RowSchBar.identifier_jo, RowSchBar.identifier_wo, RowSchBar.primary_work_unit)
                Me.ObjSch.UpdatePlanStartEndScheduleWO(RowSchBar.identifier_jo, RowSchBar.identifier_wo)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
#End Region

    Private Sub RemoveHolidaySimulation(ByVal vHolidayList As List(Of Date), ByVal vFirmPlanned As Date)
        Dim DtaAllocate As New DseTrans_ProductionSchedule_HolidayTableAdapters.tbltrans_mo_workshift_allocateTableAdapter
        Dim TblAllocate As New DseTrans_ProductionSchedule_Holiday.tbltrans_mo_workshift_allocateDataTable
        Dim ListOfWorkunit As New List(Of Double)
        Dim FirstLoop As Boolean = True
        Dim TblSchBarTanpaWO_Full As New DseTrans_Schedule.tbl_sch_barDataTable
        Dim TblSchBarTanpaWO_FromRemovedHoliday As New DseTrans_Schedule.tbl_sch_barDataTable
        Dim MoveDuration As Double = 0
        Dim FirstPlanStartDate, BufferPlanStartTime, BufferPlanFinishTime As Date
        Try
            ''load workunit yang mungkin terpengaruh
            DtaAllocate.Connection.ConnectionString = Core_Module.GetConstring
            TblAllocate.Clear()
            DtaAllocate.ClearBeforeFill = False
            For Each vRemovedHolidayDate As Date In vHolidayList
                DtaAllocate.Fill(TblAllocate, vRemovedHolidayDate, Me.CboWorkshift.SelectedValue)
            Next

            ''load semua schedule yang mungkin bergeser karena penghapusan holiday
            Me.TblSchBar.Rows.Clear()
            Me.ObjSch.LoadScheduleBar_WOItemOnly_EndDateMoreThan(vHolidayList.Item(0), Me.TblSchBar)
            FirstPlanStartDate = Me.GetFirstPlanStartDate(Me.TblSchBar)
            ListOfWorkunit = Me.GetWorkunitListFromAllocation(TblAllocate)

            ''untuk masing-masing workunit yang terlibat, cek:
            For Each PrimaryWorkunit As Double In ListOfWorkunit
                TblSchBar.DefaultView.RowFilter = "[primary work unit] = " & PrimaryWorkunit.ToString
                TblSchBar.DefaultView.Sort = "[plan start time]"

                ''load beberapa alokasi untuk keperluan geser start-finish:
                Me.ObjSch.LoadScheduleBar_TanpaScheduledItem(FirstPlanStartDate, #12/31/2200#, TblSchBarTanpaWO_Full, PrimaryWorkunit)
                Me.ObjSch.LoadScheduleBar_TanpaScheduledItem(vHolidayList.Item(0), #12/31/2200#, TblSchBarTanpaWO_FromRemovedHoliday, PrimaryWorkunit)
                FirstLoop = True
                Me.ProgressBar1.Value = 0
                Me.ProgressBar1.Maximum = Me.TblSchBar.DefaultView.Count
                For Each DV As Data.DataRowView In TblSchBar.DefaultView
                    Me.RowSchBar = DV.Row
                    Me.ProgressBar1.Value += 1
                    Me.Label6.Text = "Rescheduling, workunit: " & Me.ObjWU.GetCode(PrimaryWorkunit) & " [data " & ListOfWorkunit.IndexOf(PrimaryWorkunit) + 1 & " of " & ListOfWorkunit.Count & "], schedule: " & Me.ProgressBar1.Value & " of " & Me.ProgressBar1.Maximum & "."
                    BufferPlanStartTime = Me.RowSchBar.plan_start_time
                    BufferPlanFinishTime = Me.RowSchBar.plan_end_time
                    If FirstLoop Then
                        If Me.RowSchBar.plan_start_time > vHolidayList.Item(0) Then
                            ''harus geser termasuk start nya juga
                            Me.RowSchBar.plan_start_time = Me.ObjSch.GetStartTime_ByWorkUnit(TblSchBarTanpaWO_FromRemovedHoliday, PrimaryWorkunit, vFirmPlanned)
                            MoveDuration = -1 * Me.ObjSch.GetWorkDuration(Me.RowSchBar.plan_start_time, BufferPlanStartTime, PrimaryWorkunit, TblSchBarTanpaWO_Full)
                            Me.RowSchBar.plan_end_time = Me.ObjSch.GetEndTime(PrimaryWorkunit, Me.RowSchBar.plan_start_time, Me.RowSchBar.total_work_hour, TblSchBarTanpaWO_Full)
                        Else
                            ''hanya geser end nya aja
                            Me.RowSchBar.plan_end_time = Me.ObjSch.GetEndTime(PrimaryWorkunit, Me.RowSchBar.plan_start_time, Me.RowSchBar.total_work_hour, TblSchBarTanpaWO_Full)
                            MoveDuration = -1 * Me.ObjSch.GetWorkDuration(Me.RowSchBar.plan_end_time, BufferPlanFinishTime, PrimaryWorkunit, TblSchBarTanpaWO_Full)
                        End If
                        FirstLoop = False
                    Else
                        ''lanjutkan geser sesuai durasi geser schedule pertama
                        Me.RowSchBar.plan_start_time = Me.ObjSch.GetStartTime(PrimaryWorkunit, Me.RowSchBar.plan_start_time, MoveDuration * -1, TblSchBarTanpaWO_Full)
                        Me.RowSchBar.plan_end_time = Me.ObjSch.GetEndTime(PrimaryWorkunit, Me.RowSchBar.plan_start_time, Me.RowSchBar.total_work_hour, TblSchBarTanpaWO_Full)
                    End If
                    If Me.RowSchBar.plan_start_time = #1/1/2200# Or Me.RowSchBar.plan_end_time = #1/1/2200# Then
                        Me.RowSchBar.plan_start_time = BufferPlanStartTime
                        Me.RowSchBar.plan_end_time = BufferPlanFinishTime
                    End If
                    ''update schedule
                    Me.ObjSch.UpdatePlanStartEndSchedule(RowSchBar.plan_start_time, RowSchBar.plan_end_time, RowSchBar.identifier_jo, RowSchBar.identifier_wo, RowSchBar.primary_work_unit)
                    Me.ObjSch.UpdatePlanStartEndScheduleWO(RowSchBar.identifier_jo, RowSchBar.identifier_wo)
                Next
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub 

    Function GetRemovedHolidayList() As List(Of Date)
        GetRemovedHolidayList = New List(Of Date)
        Try
            GetRemovedHolidayList.Clear()
            For Each RowDgv In Me.DgvItem.Rows
                If RowDgv.Cells(Me.ColDelete.Index).Value = Boolean.TrueString Then
                    GetRemovedHolidayList.Add(CDate(RowDgv.Cells(Me.ColDate.Index).Value))
                End If
            Next
            Return GetRemovedHolidayList
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return GetRemovedHolidayList
        Finally
        End Try
    End Function

    Private Sub BgwRemoveHoliday_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgwRemoveHoliday.DoWork
        Try
            Me.Cursor = Cursors.WaitCursor
            Me.ObjCD.UpdateThreading()
            Me.RemoveHolidaySimulation(Me.GetRemovedHolidayList, Me.ObjSch.GetFirmPlan)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub BgwValidationRemoveHoliday_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgwRemoveHoliday.RunWorkerCompleted
        Try
            Me.LoadHolidayList()
            Me.DtpFrom.Value = Now.Date
            Me.DtpTo.Value = Now.Date
            Me.TxtDescription.Clear()
            Me.TxtRemarks.Clear()
            Me.BtnAdd.Focus()
            Me.ControlForm(True)
            MsgBox("Remove holiday and reschedule completed.", MsgBoxStyle.Information, "Information")
            Me.IsEdited = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class