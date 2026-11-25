Option Explicit On
Option Strict Off

Public Class FrmTrans_ProductionSchedule_SplitMerge
    Public UserCode As String = String.Empty
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public BackColorSource As Color = Color.Pink
    Public TblSchBar As New DseTrans_Schedule.tbl_sch_barDataTable
    Dim TblSchBarTanpaWO As New DseTrans_Schedule.tbl_sch_barDataTable
    Public ObjSch As New ClsTrans_Schedule
    Public PrimaryWU As Double = -1

    Dim objCD As New Core_Module.ClsCekData
    Dim objUserSetting As New Core_Module.ClsTransaksiSettingUser
    Dim objWorkUnit As New Core_Module.ClsMasterWorkUnit
    Dim objLocation As New Core_Module.clsMasterStorageArea
    Dim RowDgv As DataGridViewRow
    Dim Navigate As Boolean = False
    Dim Buffer As String = ""
    Dim listOfRowSelectedRowIndex As New List(Of Integer)
    Dim ObjMO As New ClsTrans_ManufacturingOrder
    Dim objLog As New ClsTrans_StockBundling
    Dim TblSchBar2 As New DseTrans_Schedule.tbl_sch_barDataTable
    Dim frmRate As New FrmTrans_ProductionSchedule_SplitMerge_Rate
    Dim rowSchBarRate As DseTrans_Schedule.tbl_sch_barRow
    Dim rowSchBar As DseTrans_Schedule.tbl_sch_barRow
    Dim RowIndex As Integer = -1
    Dim FirstCellEndEdit As Boolean = True

    Dim BufferTime As Date = Today
    Dim rowSch As DseTrans_Schedule.tbl_sch_barRow
    Dim objWU As New Core_Module.ClsMasterWorkUnit

    Private Sub LinkLabel3_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel3.LinkClicked
        Dim FrmTrans_ProductionSchedule_SplitMerge_ColSetting As New FrmTrans_ProductionSchedule_SplitMerge_ColSetting
        Try
            FrmTrans_ProductionSchedule_SplitMerge_ColSetting.UserCode = Me.UserCode
            If FrmTrans_ProductionSchedule_SplitMerge_ColSetting.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.SetColumnVisibility()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetColumnVisibility()
        Dim MO, JO, WO, Process, Code, Desc, Rate, Qty, Unit, WH, DayStart, DayEnd, Progress As String
        Try
            If Me.objUserSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_ChangeWU, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ChangeWU_ShowMO, Me.UserCode) <> "0" Then
                MO = "1"
            Else
                MO = "0"
            End If
            If Me.objUserSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_ChangeWU, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ChangeWU_ShowJO, Me.UserCode) <> "0" Then
                JO = "1"
            Else
                JO = "0"
            End If
            If Me.objUserSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_ChangeWU, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ChangeWU_ShowWO, Me.UserCode) <> "0" Then
                WO = "1"
            Else
                WO = "0"
            End If
            If Me.objUserSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_ChangeWU, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ChangeWU_ShowProcess, Me.UserCode) <> "0" Then
                Process = "1"
            Else
                Process = "0"
            End If
            If Me.objUserSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_ChangeWU, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ChangeWU_ShowCodeNo, Me.UserCode) <> "0" Then
                Code = "1"
            Else
                Code = "0"
            End If
            If Me.objUserSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_ChangeWU, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ChangeWU_ShowDesc, Me.UserCode) <> "0" Then
                Desc = "1"
            Else
                Desc = "0"
            End If
            If Me.objUserSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_ChangeWU, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ChangeWU_ShowRate, Me.UserCode) <> "0" Then
                Rate = "1"
            Else
                Rate = "0"
            End If
            If Me.objUserSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_ChangeWU, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ChangeWU_ShowQty, Me.UserCode) <> "0" Then
                Qty = "1"
            Else
                Qty = "0"
            End If
            If Me.objUserSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_ChangeWU, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ChangeWU_ShowUnit, Me.UserCode) <> "0" Then
                Unit = "1"
            Else
                Unit = "0"
            End If
            If Me.objUserSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_ChangeWU, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ChangeWU_ShowWH, Me.UserCode) <> "0" Then
                WH = "1"
            Else
                WH = "0"
            End If
            If Me.objUserSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_ChangeWU, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ChangeWU_ShowStartDay, Me.UserCode) <> "0" Then
                DayStart = "1"
            Else
                DayStart = "0"
            End If
            If Me.objUserSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_ChangeWU, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ChangeWU_ShowEndDay, Me.UserCode) <> "0" Then
                DayEnd = "1"
            Else
                DayEnd = "0"
            End If
            If Me.objUserSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_ChangeWU, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ChangeWU_ShowProgress, Me.UserCode) <> "0" Then
                Progress = "1"
            Else
                Progress = "0"
            End If
            If MO = "1" Then
                Me.ColMoNo.Visible = True
                Me.Col2MO_No.Visible = True
            Else
                Me.ColMoNo.Visible = False
                Me.Col2MO_No.Visible = False
            End If
            If JO = "1" Then
                Me.ColJONo.Visible = True
                Me.Col2JO_No.Visible = True
            Else
                Me.ColJONo.Visible = False
                Me.Col2JO_No.Visible = False
            End If
            If WO = "1" Then
                Me.ColWONo.Visible = True
                Me.Col2WO_No.Visible = True
            Else
                Me.ColWONo.Visible = False
                Me.Col2WO_No.Visible = False
            End If
            If Process = "1" Then
                Me.ColProcessCode.Visible = True
                Me.Col2ProcessCode.Visible = True
            Else
                Me.ColProcessCode.Visible = False
                Me.Col2ProcessCode.Visible = False
            End If
            If Code = "1" Then
                Me.ColCodeNo.Visible = True
                Me.Col2CodeNo.Visible = True
            Else
                Me.ColCodeNo.Visible = False
                Me.Col2CodeNo.Visible = False
            End If
            If Desc = "1" Then
                Me.ColDesc.Visible = True
                Me.Col2Desc.Visible = True
            Else
                Me.ColDesc.Visible = False
                Me.Col2Desc.Visible = False
            End If
            If Rate = "1" Then
                Me.ColRate.Visible = True
                Me.Col2Rate.Visible = True
            Else
                Me.ColRate.Visible = False
                Me.Col2Rate.Visible = False
            End If
            If Qty = "1" Then
                Me.ColQty.Visible = True
                Me.Col2Qty.Visible = True
            Else
                Me.ColQty.Visible = False
                Me.Col2Qty.Visible = False
            End If
            If Unit = "1" Then
                Me.ColUnit.Visible = True
                Me.Col2UnitCode.Visible = True
            Else
                Me.ColUnit.Visible = False
                Me.Col2UnitCode.Visible = False
            End If
            If WH = "1" Then
                Me.ColWH.Visible = True
                Me.Col2WorkHour.Visible = True
            Else
                Me.ColWH.Visible = False
                Me.Col2WorkHour.Visible = False
            End If
            If DayStart = "1" Then
                Me.ColStartDay.HeaderText = "Start"
                Me.Col2StartDay.HeaderText = "Start"
                Me.ColStartDate.HeaderText = ""
                Me.Col2StartDate.HeaderText = ""
                Me.ColStartDay.Visible = True
                Me.Col2StartDay.Visible = True
            Else
                Me.ColStartDate.HeaderText = "Start"
                Me.Col2StartDate.HeaderText = "Start"
                Me.ColStartDay.Visible = False
                Me.Col2StartDay.Visible = False
            End If
            If DayEnd = "1" Then
                Me.ColEndDay.HeaderText = "End"
                Me.Col2EndDay.HeaderText = "End"
                Me.ColEndDate.HeaderText = ""
                Me.Col2EndDate.HeaderText = ""
                Me.ColEndDay.Visible = True
                Me.Col2EndDay.Visible = True
            Else
                Me.ColEndDate.HeaderText = "End"
                Me.Col2EndDate.HeaderText = "End"
                Me.ColEndDay.Visible = False
                Me.Col2EndDay.Visible = False
            End If
            If Progress = "1" Then
                Me.ColProgress.Visible = True
                Me.Col2Progress.Visible = True
            Else
                Me.ColProgress.Visible = False
                Me.Col2Progress.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_ProductionSchedule_ChangeWU_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        FrmTrans_ProductionSchedule_SplitMerge.CheckForIllegalCrossThreadCalls = False
        Try
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvItem2.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.EnableControl(False)
            Me.Cursor = Cursors.WaitCursor
            Me.BackgroundWorker2.RunWorkerAsync()
            Me.SetColumnVisibility()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub EnableControl(ByVal vEnabledControl As Boolean)
        Try
            If vEnabledControl Then
                Me.Panel1.Visible = False
                Me.ControlBox = True
                Me.LinkLabel1.Enabled = True
                Me.LinkLabel2.Enabled = True
                Me.dgvItem.Enabled = True
                'Me.Btn1Back.Enabled = True
                Me.Btn1Next.Enabled = True
                Me.LinkLabel3.Enabled = True
                'Me.BtnOK.Enabled = True
                Me.BtnCancel.Enabled = True
                Me.TxtWUCode.Enabled = True
                Me.BtnFindWU.Enabled = True
                Me.ChkSetAsLast.Enabled = True
                Me.DgvItem2.Enabled = True
            Else
                Me.Panel1.Visible = True
                Me.ControlBox = False
                Me.LinkLabel1.Enabled = False
                Me.LinkLabel2.Enabled = False
                Me.dgvItem.Enabled = False
                Me.Btn1Back.Enabled = False
                Me.Btn1Next.Enabled = False
                Me.LinkLabel3.Enabled = False
                Me.BtnOK.Enabled = False
                Me.BtnCancel.Enabled = False
                Me.TxtWUCode.Enabled = False
                Me.BtnFindWU.Enabled = False
                Me.ChkSetAsLast.Enabled = False
                Me.DgvItem2.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnCancel.Click
        Try
            Me.DialogResult = Windows.Forms.DialogResult.Cancel
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        Try
            If Me.TabControl1.SelectedTab.Name = Me.TabPage1.Name Then
                Me.BtnOK.Enabled = False
                Me.Btn1Back.Enabled = False
                Me.Btn1Next.Enabled = True
                Me.BtnOK.Enabled = False
                listOfRowSelectedRowIndex.Clear()
                listOfRowSelectedRowIndex = Me.getListOfSelectedRowIndex
            ElseIf Me.TabControl1.SelectedTab.Name = Me.TabPage2.Name Then
                Me.BtnOK.Enabled = True
                Me.Btn1Back.Enabled = True
                Me.Btn1Next.Enabled = False
                Me.BtnOK.Enabled = True
                If IsDifferentList(listOfRowSelectedRowIndex, Me.getListOfSelectedRowIndex) Then
                    Me.TxtPrimaryWU.Text = -1
                    Me.TxtWUCode.Clear()
                    Me.TxtWUDesc.Clear()
                    Me.TxtWUCode.Focus()
                Else
                    Me.TxtWUCode.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Function GetListOfSelectedRowIndex() As List(Of Integer)
        GetListOfSelectedRowIndex = New List(Of Integer)
        Try
            For Each Me.RowDgv In Me.dgvItem.Rows
                If Me.RowDgv.Cells(Me.ColModify.Index).Value = Boolean.TrueString Then
                    GetListOfSelectedRowIndex.Add(Me.RowDgv.Index)
                End If
            Next
            Return GetListOfSelectedRowIndex
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return getListOfSelectedRowIndex
        End Try
    End Function

    Function IsDifferentList(ByVal vFirstList As List(Of Integer), ByVal vSecondList As List(Of Integer)) As Boolean
        Try
            If vFirstList.Count <> vSecondList.Count Then
                Return True
            End If
            For Each i As Integer In vFirstList
                If vSecondList.Contains(i) = False Then
                    Return True
                End If
            Next
            For Each i As Integer In vSecondList
                If vFirstList.Contains(i) = False Then
                    Return True
                End If
            Next
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub Btn1Back_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn1Back.Click
        Try
            Me.Navigate = True
            Me.TabControl1.SelectedTab = Me.TabPage1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Btn1Next_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn1Next.Click
        Try
            If Me.IsValidToMove Then
                Me.Navigate = True
                Me.TabControl1.SelectedTab = Me.TabPage2
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Function IsValidToMove() As Boolean
        Dim SelectedCount As Integer = 0
        Dim LastSelectedIndex As Integer = -1
        Dim FirmPlanned As Date
        Dim listOfClosedJO As New List(Of String)
        Dim ListOfLoggedSch As New List(Of String)
        Try

            ''01. minimal ada 1 item yang dipindah
            For Each Me.RowDgv In Me.dgvItem.Rows
                If Me.RowDgv.Cells(Me.ColModify.Index).Value = Boolean.TrueString Then
                    SelectedCount += 1
                End If
            Next
            If SelectedCount = 0 Then
                MsgBox("Select at least 1 item to move.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            ''02. item yang dipindah jika lebih dari 1 item, harus berurutan barisnya
            If SelectedCount > 1 Then
                For Each Me.RowDgv In Me.dgvItem.Rows
                    If Me.RowDgv.Cells(Me.ColModify.Index).Value = Boolean.TrueString Then
                        If Me.RowDgv.Index - 1 > LastSelectedIndex And LastSelectedIndex > -1 Then
                            MsgBox("The selected items must be scheduled in sequence.", MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                        LastSelectedIndex = Me.RowDgv.Index
                    End If
                Next
            End If

            ''03. item yang dipilih untuk dipindah minimal 1 hari setelah tanggal firm planned
            FirmPlanned = Me.ObjSch.GetFirmPlan
            For Each Me.RowDgv In Me.dgvItem.Rows
                If Me.RowDgv.Cells(Me.Col2Insert.Index).Value = Boolean.TrueString Then
                    If CDate(Me.RowDgv.Cells(Me.ColPlanStartTime.Index).Value).Date <= FirmPlanned Then
                        MsgBox("The selected schedule is firm planned, you can't move the schedule.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                    If Me.ObjMO.IsJOFinished(Me.RowDgv.Cells(Me.ColIdentifierJO.Index).Value) = True Then
                        If listOfClosedJO.Contains(Me.RowDgv.Cells(Me.ColJONo.Index).Value) = False Then listOfClosedJO.Add(Me.RowDgv.Cells(Me.ColJONo.Index).Value)
                    End If
                    If Me.objLog.GetJml_ByJO_WO_PrimaryWorkUnit(Me.RowDgv.Cells(Me.ColIdentifierJO.Index).Value, Me.RowDgv.Cells(Me.colIdentifierWO.Index).Value, Me.PrimaryWU) > 0 Then
                        If ListOfLoggedSch.Contains("JO: " & Me.RowDgv.Cells(Me.ColJONo.Index).Value & " WO: " & Me.RowDgv.Cells(Me.ColWONo.Index).Value) = False Then ListOfLoggedSch.Add("JO: " & Me.RowDgv.Cells(Me.ColJONo.Index).Value & " WO: " & Me.RowDgv.Cells(Me.ColWONo.Index).Value)
                    End If
                End If
            Next

            ''04. item-item yang dipindah gak boleh atas JO closed
            If listOfClosedJO.Count > 0 Then
                Dim ClosedJO As String = ""
                For Each s As String In listOfClosedJO
                    If ClosedJO = "" Then
                        ClosedJO = s
                    Else
                        ClosedJO = ClosedJO & ", " & s
                    End If
                Next
                MsgBox("This following job order is closed, you can't move the schedule:" & vbNewLine & ClosedJO & ".", MsgBoxStyle.Information, "Information")
                Return False
            End If

            ''05. item-item yang di pindah gak boleh yang sudah ada log
            If ListOfLoggedSch.Count > 0 Then
                Dim LoggedSch As String = ""
                For Each s As String In ListOfLoggedSch
                    If LoggedSch = "" Then
                        LoggedSch = s
                    Else
                        LoggedSch = vbNewLine & s
                    End If
                Next
                MsgBox("The selected schedule is refered by 1 or more production log, you can't move the schedule.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            If e.ColumnIndex <> Me.ColModify.Index Then
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvItem.DataError

    End Sub

    Private Sub dgvItem_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.Sorted
        Try
            Me.objCD.SetNumber(Me.dgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Function GetFirstMoved() As Integer
        Try
            For Each Me.RowDgv In Me.DgvItem2.Rows
                If Me.RowDgv.Cells(Me.Col2ReadOnly.Index).Value = "1" Then Return Me.DgvItem2.Rows.IndexOf(Me.RowDgv)
            Next
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub DgvItem2_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvItem2.CellBeginEdit
        Dim FirmPlanned As Date
        Try
            If Me.FirstCellEndEdit = True And e.RowIndex = 0 And e.ColumnIndex = 0 Then
                Me.FirstCellEndEdit = False
                e.Cancel = True
                Exit Sub
            End If

            If (e.ColumnIndex <> Me.Col2Insert.Index And e.ColumnIndex <> Me.Col2StartDate.Index) Or (e.ColumnIndex = Me.Col2Insert.Index And Me.DgvItem2.Item(Me.Col2ReadOnly.Index, e.RowIndex).Value = "1") Then
                e.Cancel = True
                Exit Sub
            End If

            ''hanya bisa edit start time pertama yang dipindahkan
            If e.ColumnIndex = Me.Col2StartDate.Index Then
                If e.RowIndex <> Me.GetFirstMoved Then
                    MsgBox("You can only modify first moved item.", MsgBoxStyle.Information, "Information")
                    e.Cancel = True
                    Exit Sub
                End If
                Me.Buffer = Me.DgvItem2.Item(e.ColumnIndex, e.RowIndex).Value
                Me.BufferTime = CDate(Me.DgvItem2.Item(Me.Col2StartDate.Index, e.RowIndex).Value).Date + TimeSpan.FromMinutes(Me.objCD.Convert_ToMinutes(Me.DgvItem2.Item(Me.Col2StartTime.Index, e.RowIndex).Value))
                Me.DgvItem2.Item(e.ColumnIndex, e.RowIndex).Value = Format(Me.BufferTime, "dd MMM yyyy HH.mm")
            End If

            If Me.DgvItem2.Item(Me.Col2Insert.Index, e.RowIndex).Value = Boolean.TrueString Then
                e.Cancel = True
                Exit Sub
            End If

            If IsDate(Me.DgvItem2.Item(Me.Col2PlanStartTime.Index, e.RowIndex).Value) = True Then
                FirmPlanned = Me.ObjSch.GetFirmPlan
                If CDate(Me.DgvItem2.Item(Me.Col2PlanStartTime.Index, e.RowIndex).Value).Date <= FirmPlanned Then
                    MsgBox("Selected schedule is firm planned, you can't insert on it schedule.", MsgBoxStyle.Information, "Information")
                    e.Cancel = True
                    Exit Sub
                End If
            End If

            If String.IsNullOrEmpty(Me.DgvItem2.Item(Me.Col2IdentifierJO.Index, e.RowIndex).Value) = False Then
                If Me.objLog.GetJml_ByJO_WO_PrimaryWorkUnit(Me.DgvItem2.Item(Me.Col2IdentifierJO.Index, e.RowIndex).Value, Me.DgvItem2.Item(Me.Col2IdentifierWO.Index, e.RowIndex).Value, Me.TxtPrimaryWU.Text) > 0 Then
                    MsgBox("This schedule is refered by 1 or more production log, you can't insert on it schedule.", MsgBoxStyle.Information, "Information")
                    e.Cancel = True
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItem2_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvItem2.CellEndEdit
        Dim TimeComponent As String()
        Dim StartTimeLama, EndTimeLama, EndTimeBaru As Date
        Dim DummyTime As Date = Today
        Dim dd, MMM, yyyy, HHmm As String
        Dim FormatBetul As Boolean = True
        Dim IdentifierWO As String = ""
        Dim IdentifierJO As String = ""
        Dim PrimaryWU As Double = -1
        Dim DurasiGeser As Double = 0

        Try
            If e.ColumnIndex = Me.Col2Insert.Index Then
                If Me.DgvItem2.Item(e.ColumnIndex, e.RowIndex).Value = Boolean.TrueString Then
                    Me.RowIndex = e.RowIndex
                    For Each Me.RowDgv In Me.DgvItem2.Rows
                        If Me.RowDgv.Index <> e.RowIndex Then
                            Me.RowDgv.Cells(Me.Col2Insert.Index).Value = Boolean.FalseString
                        End If
                    Next
                    Me.BackgroundWorker1.RunWorkerAsync()
                End If
            End If

            If e.ColumnIndex = Me.Col2StartDate.Index Then
                ''cek validitas data:
                If IsNothing(Me.DgvItem2.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                    Me.DgvItem2.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                Else
                    TimeComponent = Me.DgvItem2.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Split(" ")
                    If TimeComponent.Length <> 4 Then
                        FormatBetul = False
                    Else
                        dd = TimeComponent(0)
                        MMM = TimeComponent(1)
                        yyyy = TimeComponent(2)
                        HHmm = TimeComponent(3)
                        If IsDate(dd + " " + MMM + " " + yyyy) = False Then
                            FormatBetul = False
                        End If
                        If Me.objCD.Is_Time_String(HHmm) = False Then
                            FormatBetul = False
                        End If
                        If FormatBetul = True Then
                            DummyTime = CDate(dd + " " + MMM + " " + yyyy).Date + TimeSpan.FromMinutes(Me.objCD.Convert_ToMinutes(HHmm))
                            If CDate(Me.BufferTime) = DummyTime Then
                                Me.DgvItem2.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                Exit Sub
                            End If
                        End If
                    End If
                    If FormatBetul = False Then
                        MsgBox("Please specify format with pattern dd MMM yyy HH.mm.", MsgBoxStyle.Information, "Information")
                        Me.DgvItem2.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    Else
                        StartTimeLama = Me.BufferTime
                        EndTimeLama = CDate(Me.DgvItem2.Item(Me.Col2EndDate.Index, e.RowIndex).Value).Date + _
                        TimeSpan.FromMinutes(Me.objCD.Convert_ToMinutes(Me.DgvItem2.Item(Me.Col2EndTime.Index, e.RowIndex).Value))
                        IdentifierWO = Me.DgvItem2.Item(Me.Col2IdentifierWO.Index, e.RowIndex).Value
                        IdentifierJO = Me.DgvItem2.Item(Me.Col2IdentifierJO.Index, e.RowIndex).Value
                        PrimaryWU = Me.TxtPrimaryWU.Text
                        For Each Me.rowSch In Me.frmRate.TblSchBar.Rows
                            If Me.rowSch.primary_work_unit = PrimaryWU Then
                                If Me.rowSch.identifier_wo = IdentifierWO And Me.rowSch.identifier_jo = IdentifierJO Then
                                    If Me.rowSch.plan_start_time = StartTimeLama And Me.rowSch.plan_end_time = EndTimeLama Then
                                        ''cek validitas schedule:
                                        EndTimeBaru = Me.ObjSch.GetEndTime(PrimaryWU, DummyTime, rowSch.total_work_hour, Me.TblSchBarTanpaWO)
                                        If EndTimeBaru = #1/1/2200# Then
                                            MsgBox("Insufficient workshift allocation, please allocate more workshift for " + Me.objWU.GetCode(rowSch.primary_work_unit) + ".", MsgBoxStyle.Information, "Information")
                                            Me.DgvItem2.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                            Exit Sub
                                        End If
                                        If DummyTime <= Me.BufferTime Then
                                            DurasiGeser = Me.ObjSch.GetWorkDuration(DummyTime, Me.BufferTime, Me.rowSch.primary_work_unit, Me.TblSchBarTanpaWO)
                                            DurasiGeser = DurasiGeser * (-1)
                                        Else
                                            DurasiGeser = Me.ObjSch.GetWorkDuration(Me.BufferTime, DummyTime, Me.rowSch.primary_work_unit, Me.TblSchBarTanpaWO)
                                        End If
                                        If Me.ObjSch.ApakahBolehGeserStartTime_FormUtama(Me.rowSch, Me.TblSchBarTanpaWO, Me.frmRate.TblSchBar, _
                                                DummyTime, EndTimeBaru, StartTimeLama, DurasiGeser) = False Then
                                            Me.DgvItem2.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                            Exit Sub
                                        End If
                                        If Me.ObjSch.ApakahBolehGeserStartTime_FormUtama(Me.rowSch, Me.TblSchBarTanpaWO, Me.TblSchBar2, _
                                                DummyTime, EndTimeBaru, StartTimeLama, DurasiGeser) = False Then
                                            Me.DgvItem2.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                            Exit Sub
                                        End If
                                        ''update tbl & dgv ybs:
                                        Me.rowSch.plan_start_time = DummyTime
                                        Me.rowSch.plan_end_time = Me.ObjSch.GetEndTime(Me.rowSch.primary_work_unit, Me.rowSch.plan_start_time, Me.rowSch.total_work_hour, Me.TblSchBarTanpaWO)
                                        Me.DgvItem2.Item(Me.Col2StartDay.Index, e.RowIndex).Value = Me.objCD.GetHari_English(Me.rowSch.plan_start_time.Date.DayOfWeek)
                                        Me.DgvItem2.Item(Me.Col2StartDate.Index, e.RowIndex).Value = Format(Me.rowSch.plan_start_time, "dd MMM yyyy")
                                        Me.DgvItem2.Item(Me.Col2StartTime.Index, e.RowIndex).Value = Format(Me.rowSch.plan_start_time, "HH.mm")
                                        Me.DgvItem2.Item(Me.Col2EndDay.Index, e.RowIndex).Value = Me.objCD.GetHari_English(Me.rowSch.plan_end_time.Date.DayOfWeek)
                                        Me.DgvItem2.Item(Me.Col2EndDate.Index, e.RowIndex).Value = Format(Me.rowSch.plan_end_time, "dd MMM yyyy")
                                        Me.DgvItem2.Item(Me.Col2EndTime.Index, e.RowIndex).Value = Format(Me.rowSch.plan_end_time, "HH.mm")
                                        ''geserin jadwal setelahnya (tbl & dgv):
                                        Me.MoveSchedule(frmRate.TblSchBar, Me.rowSch.identifier_jo, Me.rowSch.identifier_wo, DurasiGeser, Me.TblSchBarTanpaWO)
                                        If Me.DgvItem2.Rows.Count - 1 > e.RowIndex Then
                                            Me.DgvItem2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                                            For rowIdx As Integer = Me.DgvItem2.Rows.Count - 1 To e.RowIndex + 1 Step -1
                                                Me.DgvItem2.Rows.RemoveAt(rowIdx)
                                            Next
                                            frmRate.TblSchBar.DefaultView.RowFilter = "[Merge] = False"
                                            frmRate.TblSchBar.DefaultView.Sort = "[plan start time]"
                                            Dim Skip As Boolean = True
                                            For Each DV As Data.DataRowView In frmRate.TblSchBar.DefaultView
                                                If Skip Then
                                                    Skip = False
                                                Else
                                                    rowSchBarRate = DV.Row
                                                    RowDgv = Me.DgvItem2.Rows(Me.DgvItem2.Rows.Add())
                                                    With RowDgv
                                                        .Cells(Me.Col2Insert.Index).Value = Boolean.FalseString
                                                        .Cells(Me.Col2IdentifierJO.Index).Value = rowSchBarRate.identifier_jo
                                                        .Cells(Me.Col2IdentifierWO.Index).Value = rowSchBarRate.identifier_wo
                                                        .Cells(Me.Col2MO_No.Index).Value = rowSchBarRate.mo_no
                                                        .Cells(Me.Col2JO_No.Index).Value = rowSchBarRate.jo_no
                                                        .Cells(Me.Col2ProcessCode.Index).Value = rowSchBarRate.process_code
                                                        .Cells(Me.Col2WO_No.Index).Value = rowSchBarRate.wo_no
                                                        .Cells(Me.Col2CodeNo.Index).Value = rowSchBarRate.product_code
                                                        .Cells(Me.Col2Desc.Index).Value = rowSchBarRate.product_desc
                                                        .Cells(Me.Col2Rate.Index).Value = Me.objCD.DigitGrouping(rowSchBarRate.rate, rowSchBarRate.decimal_point)
                                                        .Cells(Me.Col2Qty.Index).Value = Me.objCD.DigitGrouping(rowSchBarRate.qty_produce, rowSchBarRate.decimal_point)
                                                        .Cells(Me.Col2UnitCode.Index).Value = rowSchBarRate.unit_code
                                                        .Cells(Me.Col2WorkHour.Index).Value = Me.objCD.Convert_ToString(rowSchBarRate.total_work_hour)
                                                        .Cells(Me.Col2StartDay.Index).Value = Me.objCD.GetHari_English(rowSchBarRate.plan_start_time.Date.DayOfWeek)
                                                        .Cells(Me.Col2StartDate.Index).Value = Format(rowSchBarRate.plan_start_time, "dd MMM yyyy")
                                                        .Cells(Me.Col2StartTime.Index).Value = Format(rowSchBarRate.plan_start_time, "HH.mm")
                                                        .Cells(Me.Col2EndDay.Index).Value = Me.objCD.GetHari_English(rowSchBarRate.plan_end_time.Date.DayOfWeek)
                                                        .Cells(Me.Col2EndDate.Index).Value = Format(rowSchBarRate.plan_end_time, "dd MMM yyyy")
                                                        .Cells(Me.Col2EndTime.Index).Value = Format(rowSchBarRate.plan_end_time, "HH.mm")
                                                        .Cells(Me.Col2Progress.Index).Value = Me.objCD.DigitGrouping(rowSchBarRate.progress, 2)
                                                        .Cells(Me.Col2PrimaryMO.Index).Value = rowSchBarRate.primary_mo
                                                        .Cells(Me.Col2PrimaryProcess.Index).Value = rowSchBarRate.primary_process
                                                        .Cells(Me.Col2PrimaryCodeNo.Index).Value = rowSchBarRate.primary_product
                                                        .Cells(Me.Col2PrimaryUnit.Index).Value = rowSchBarRate.primary_unit
                                                        .Cells(Me.Col2ReadOnly.Index).Value = 1
                                                        .Cells(Me.Col2PlanStartTime.Index).Value = rowSchBarRate.plan_start_time
                                                        .DefaultCellStyle.BackColor = Me.BackColorSource
                                                    End With
                                                End If
                                            Next
                                            Me.DgvItem2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                                            Me.DgvItem2.ClearSelection()
                                            Me.FirstCellEndEdit = True
                                        End If
                                        Exit For
                                    End If
                                End If
                            End If
                        Next
                    End If
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub MoveSchedule(ByRef vTblSchBar As DseTrans_Schedule.tbl_sch_barDataTable, ByVal vIdentifierJO As String, ByVal vIdentifierWO As String, ByVal vMoveDuration As Integer, ByVal vEmptyTblSchedule As DseTrans_Schedule.tbl_sch_barDataTable)
        Dim StartUpdate As Boolean = False
        Try
            For Each Me.rowSch In vTblSchBar.Rows
                If StartUpdate = True Then
                    If vMoveDuration > 0 Then
                        Me.rowSch.plan_start_time = Me.ObjSch.GetEndTime(Me.TxtPrimaryWU.Text, Me.rowSch.plan_start_time, vMoveDuration, vEmptyTblSchedule)
                    Else
                        Me.rowSch.plan_start_time = Me.ObjSch.GetStartTime(Me.TxtPrimaryWU.Text, Me.rowSch.plan_start_time, vMoveDuration * -1, vEmptyTblSchedule)
                    End If
                    Me.rowSch.plan_end_time = Me.ObjSch.GetEndTime(Me.TxtPrimaryWU.Text, Me.rowSch.plan_start_time, Me.rowSch.total_work_hour, vEmptyTblSchedule)
                Else
                    If Me.rowSch.identifier_wo = vIdentifierWO And Me.rowSch.identifier_jo = vIdentifierJO Then
                        StartUpdate = True
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Function ReOrderSchedule() As Boolean
        Dim rowSchBar As DseTrans_Schedule.tbl_sch_barRow
        Dim rowSchBarRate As DseTrans_Schedule.tbl_sch_barRow
        Dim RowIndex As Integer = -1
        Dim ReBuilStartEnd As Boolean = False

        Dim AlreadyCheckInsertBox As Boolean = False
        Dim ReCheckInsertBox As Boolean = False

        Dim endTimeBefore As Date = #1/1/1950#
        Dim endTimeBefore_AfterReBuild As Date
        Dim endTimeBeforeRate As Date = #1/1/1950#
        Dim endTimeBeforeRate_AfterRebulid As Date
        Dim FirstLoop As Boolean = True
        Dim FirstLoop1 As Boolean = True
        Dim delay0 As Decimal = 0
        Dim delay As Decimal = 0
        Dim LastReadOnlyIndex As Integer = -1

        Try

            For Each Me.RowDgv In Me.DgvItem2.Rows
                If Me.RowDgv.Cells(Me.Col2ReadOnly.Index).Value = "1" Then
                    LastReadOnlyIndex = Me.RowDgv.Index
                End If
            Next
            If Me.RowIndex > LastReadOnlyIndex Then
                Me.RowIndex -= frmRate.TblSchBar.DefaultView.Count
            End If

            Me.DgvItem2.Rows.Clear()
            For Each DV As Data.DataRowView In Me.TblSchBar2.DefaultView
                rowSchBar = DV.Row
                delay0 = Me.ObjSch.GetWorkDuration(endTimeBefore, rowSchBar.plan_start_time, Me.TxtPrimaryWU.Text, Me.TblSchBarTanpaWO)
                endTimeBefore = rowSchBar.plan_end_time
                RowIndex += 1
                If RowIndex = Me.RowIndex Then
                    For Each DV2 As Data.DataRowView In frmRate.TblSchBar.DefaultView
                        rowSchBarRate = DV2.Row
                        delay = Me.ObjSch.GetWorkDuration(endTimeBeforeRate, rowSchBarRate.plan_start_time, Me.TxtPrimaryWU.Text, Me.TblSchBarTanpaWO)
                        endTimeBeforeRate = rowSchBarRate.plan_end_time
                        If FirstLoop Then
                            rowSchBarRate.plan_start_time = rowSchBar.Item("original plan start time") 'rowSchBar.plan_start_time
                            FirstLoop = False
                        Else
                            rowSchBarRate.plan_start_time = Me.ObjSch.GetEndTime(Me.TxtPrimaryWU.Text, endTimeBeforeRate_AfterRebulid, delay, Me.TblSchBarTanpaWO)
                        End If
                        rowSchBarRate.plan_end_time = Me.ObjSch.GetEndTime(Me.TxtPrimaryWU.Text, rowSchBarRate.plan_start_time, rowSchBarRate.total_work_hour, Me.TblSchBarTanpaWO)
                        endTimeBeforeRate_AfterRebulid = rowSchBarRate.plan_end_time
                        RowDgv = Me.DgvItem2.Rows(Me.DgvItem2.Rows.Add())
                        With RowDgv
                            .Cells(Me.Col2Insert.Index).Value = Boolean.FalseString
                            .Cells(Me.Col2IdentifierJO.Index).Value = rowSchBarRate.identifier_jo
                            .Cells(Me.Col2IdentifierWO.Index).Value = rowSchBarRate.identifier_wo
                            .Cells(Me.Col2MO_No.Index).Value = rowSchBarRate.mo_no
                            .Cells(Me.Col2JO_No.Index).Value = rowSchBarRate.jo_no
                            .Cells(Me.Col2ProcessCode.Index).Value = rowSchBarRate.process_code
                            .Cells(Me.Col2WO_No.Index).Value = rowSchBarRate.wo_no
                            .Cells(Me.Col2CodeNo.Index).Value = rowSchBarRate.product_code
                            .Cells(Me.Col2Desc.Index).Value = rowSchBarRate.product_desc
                            .Cells(Me.Col2Rate.Index).Value = Me.objCD.DigitGrouping(rowSchBarRate.rate, rowSchBar.decimal_point)
                            .Cells(Me.Col2Qty.Index).Value = Me.objCD.DigitGrouping(rowSchBarRate.qty_produce, rowSchBar.decimal_point)
                            .Cells(Me.Col2UnitCode.Index).Value = rowSchBarRate.unit_code
                            .Cells(Me.Col2WorkHour.Index).Value = Me.objCD.Convert_ToString(rowSchBarRate.total_work_hour)
                            .Cells(Me.Col2StartDay.Index).Value = Me.objCD.GetHari_English(rowSchBarRate.plan_start_time.Date.DayOfWeek)
                            .Cells(Me.Col2StartDate.Index).Value = Format(rowSchBarRate.plan_start_time, "dd MMM yyyy")
                            .Cells(Me.Col2StartTime.Index).Value = Format(rowSchBarRate.plan_start_time, "HH.mm")
                            .Cells(Me.Col2EndDay.Index).Value = Me.objCD.GetHari_English(rowSchBarRate.plan_end_time.Date.DayOfWeek)
                            .Cells(Me.Col2EndDate.Index).Value = Format(rowSchBarRate.plan_end_time, "dd MMM yyyy")
                            .Cells(Me.Col2EndTime.Index).Value = Format(rowSchBarRate.plan_end_time, "HH.mm")
                            .Cells(Me.Col2Progress.Index).Value = Me.objCD.DigitGrouping(rowSchBarRate.progress, 2)
                            .Cells(Me.Col2PrimaryMO.Index).Value = rowSchBarRate.primary_mo
                            .Cells(Me.Col2PrimaryProcess.Index).Value = rowSchBarRate.primary_process
                            .Cells(Me.Col2PrimaryCodeNo.Index).Value = rowSchBarRate.primary_product
                            .Cells(Me.Col2PrimaryUnit.Index).Value = rowSchBarRate.primary_unit
                            .Cells(Me.Col2ReadOnly.Index).Value = 1
                            .Cells(Me.Col2PlanStartTime.Index).Value = rowSchBarRate.plan_start_time
                            .DefaultCellStyle.BackColor = Me.BackColorSource
                        End With
                    Next
                    ReBuilStartEnd = True
                End If
                If ReBuilStartEnd = True Then
                    If FirstLoop1 Then
                        rowSchBar.plan_start_time = endTimeBeforeRate_AfterRebulid
                        FirstLoop1 = False
                    Else
                        rowSchBar.plan_start_time = Me.ObjSch.GetEndTime(Me.TxtPrimaryWU.Text, endTimeBefore_AfterReBuild, delay0, Me.TblSchBarTanpaWO)
                    End If
                    rowSchBar.plan_end_time = Me.ObjSch.GetEndTime(Me.TxtPrimaryWU.Text, rowSchBar.plan_start_time, rowSchBar.total_work_hour, Me.TblSchBarTanpaWO)
                    endTimeBefore_AfterReBuild = rowSchBar.plan_end_time
                    If AlreadyCheckInsertBox = False Then
                        ReCheckInsertBox = True
                        AlreadyCheckInsertBox = True
                    Else
                        ReCheckInsertBox = False
                    End If
                End If
                RowDgv = Me.DgvItem2.Rows(Me.DgvItem2.Rows.Add())
                With RowDgv
                    .Cells(Me.Col2Insert.Index).Value = ReCheckInsertBox
                    .Cells(Me.Col2IdentifierJO.Index).Value = rowSchBar.identifier_jo
                    .Cells(Me.Col2IdentifierWO.Index).Value = rowSchBar.identifier_wo
                    .Cells(Me.Col2MO_No.Index).Value = rowSchBar.mo_no
                    .Cells(Me.Col2JO_No.Index).Value = rowSchBar.jo_no
                    .Cells(Me.Col2ProcessCode.Index).Value = rowSchBar.process_code
                    .Cells(Me.Col2WO_No.Index).Value = rowSchBar.wo_no
                    .Cells(Me.Col2CodeNo.Index).Value = rowSchBar.product_code
                    .Cells(Me.Col2Desc.Index).Value = rowSchBar.product_desc
                    .Cells(Me.Col2Rate.Index).Value = Me.objCD.DigitGrouping(rowSchBar.rate, rowSchBar.decimal_point)
                    .Cells(Me.Col2Qty.Index).Value = Me.objCD.DigitGrouping(rowSchBar.qty_produce, rowSchBar.decimal_point)
                    .Cells(Me.Col2UnitCode.Index).Value = rowSchBar.unit_code
                    .Cells(Me.Col2WorkHour.Index).Value = Me.objCD.Convert_ToString(rowSchBar.total_work_hour)
                    .Cells(Me.Col2StartDay.Index).Value = Me.objCD.GetHari_English(rowSchBar.plan_start_time.Date.DayOfWeek)
                    .Cells(Me.Col2StartDate.Index).Value = Format(rowSchBar.plan_start_time, "dd MMM yyyy")
                    .Cells(Me.Col2StartTime.Index).Value = Format(rowSchBar.plan_start_time, "HH.mm")
                    .Cells(Me.Col2EndDay.Index).Value = Me.objCD.GetHari_English(rowSchBar.plan_end_time.Date.DayOfWeek)
                    .Cells(Me.Col2EndDate.Index).Value = Format(rowSchBar.plan_end_time, "dd MMM yyyy")
                    .Cells(Me.Col2EndTime.Index).Value = Format(rowSchBar.plan_end_time, "HH.mm")
                    .Cells(Me.Col2Progress.Index).Value = Me.objCD.DigitGrouping(rowSchBar.progress, 2)
                    .Cells(Me.Col2PrimaryMO.Index).Value = rowSchBar.primary_mo
                    .Cells(Me.Col2PrimaryProcess.Index).Value = rowSchBar.primary_process
                    .Cells(Me.Col2PrimaryCodeNo.Index).Value = rowSchBar.primary_product
                    .Cells(Me.Col2PrimaryUnit.Index).Value = rowSchBar.primary_unit
                    .Cells(Me.Col2ReadOnly.Index).Value = 0
                    .Cells(Me.Col2PlanStartTime.Index).Value = rowSchBar.plan_start_time
                End With
            Next
            Me.DgvItem2.ClearSelection()
            Me.FirstCellEndEdit = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Sub DgvItem2_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DgvItem2.DataError

    End Sub

    Private Sub DgvItem2_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvItem2.Sorted
        Try
            Me.objCD.SetNumber(Me.DgvItem2)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Dim ForceStop As Boolean = True
        Dim IdentifierJO_Reference As String = ""
        Try
            For Each Me.RowDgv In Me.dgvItem.Rows
                If Me.RowDgv.Cells(Me.ColModify.Index).Value = Boolean.TrueString Then
                    ForceStop = False
                    IdentifierJO_Reference = Me.RowDgv.Cells(Me.ColIdentifierJO.Index).Value
                    Exit For
                End If
            Next
            If ForceStop = True Then
                MsgBox("Select at least 1 item as reference.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            For Each Me.RowDgv In Me.dgvItem.Rows
                If Me.RowDgv.Cells(Me.ColIdentifierJO.Index).Value = IdentifierJO_Reference Then
                    Me.RowDgv.Cells(Me.ColModify.Index).Value = Boolean.TrueString
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CheckAllDgv(ByVal vDgv As DataGridView, ByVal vColIndex As Integer, ByVal vCheck As Boolean)
        Try
            For Each Me.RowDgv In vDgv.Rows
                If vCheck Then
                    Me.RowDgv.Cells(vColIndex).Value = Boolean.TrueString
                Else
                    Me.RowDgv.Cells(vColIndex).Value = Boolean.FalseString
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Try
            Me.CheckAllDgv(Me.dgvItem, Me.ColModify.Index, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnOK.Click
        Dim RowSchBar2 As DseTrans_Schedule.tbl_sch_barRow
        Dim PreviousRowPlanEndTime As Date = #1/1/1950#
        Dim ObjUnit As New Core_Module.ClsMasterWorkUnit
        Dim FirstLoop As Boolean = True
        Try
            If Me.DgvItem2.Rows.Count = 0 Then
                MsgBox("Select destination work unit.", MsgBoxStyle.Information, "Information")
                Me.TxtWUCode.Focus()
                Exit Sub
            End If
            If MsgBox("Are you sure to save change?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then

                ''update qty produce pada workunit asal dengan sisa setelah dikurangi oleh qty yang di share ke workunit tujuan
                Me.TblSchBar.DefaultView.RowFilter = "[Primary Work Unit] = " & Me.PrimaryWU.ToString
                Me.TblSchBar.DefaultView.Sort = "[plan start time]"
                For Each DV As Data.DataRowView In TblSchBar.DefaultView
                    rowSchBar = DV.Row
                    frmRate.TblSchBar.DefaultView.RowFilter = "[identifier jo] = '" & Me.rowSchBar.identifier_jo & "' and [identifier wo] = '" & Me.rowSchBar.identifier_wo & "' and [primary work unit] = " & Me.TxtPrimaryWU.Text.ToString
                    If frmRate.TblSchBar.DefaultView.Count > 0 Then
                        Me.rowSchBarRate = Me.frmRate.TblSchBar.DefaultView(0).Row
                        Me.rowSchBar.qty_produce = Me.rowSchBarRate.Item("Qty Sisa")
                        Me.rowSchBar.total_work_hour = (Me.rowSchBarRate.Item("Qty Sisa") / Me.rowSchBar.rate) * 60
                    End If
                Next

                ''hitung delay
                Me.TblSchBar.DefaultView.RowFilter = "[primary work unit] = " + Me.PrimaryWU.ToString
                Me.TblSchBar.DefaultView.Sort = "[plan start time]"
                FirstLoop = True
                For Each DV As Data.DataRowView In Me.TblSchBar.DefaultView
                    rowSchBar = DV.Row
                    If FirstLoop Then
                        FirstLoop = False
                        rowSchBar.delay = 0
                    Else
                        If PreviousRowPlanEndTime = rowSchBar.plan_start_time Then
                            rowSchBar.delay = 0
                        Else
                            rowSchBar.delay = Me.ObjSch.GetWorkDuration(PreviousRowPlanEndTime, rowSchBar.plan_start_time, Me.PrimaryWU, Me.TblSchBarTanpaWO)
                        End If
                    End If
                    PreviousRowPlanEndTime = rowSchBar.plan_end_time
                Next

                ''hapus schedule pada workunit asal yang qty-nya = 0
                Me.TblSchBar.DefaultView.RowFilter = "([total work hour]=0 OR [qty produce]=0) and [primary work unit] = " & Me.PrimaryWU.ToString
                While Me.TblSchBar.DefaultView.Count > 0
                    Me.rowSchBar = Me.TblSchBar.DefaultView(0).Row
                    Me.ObjSch.DeleteZeroQty(Me.rowSchBar.identifier_jo, Me.rowSchBar.identifier_wo, Me.PrimaryWU)
                    Me.TblSchBar.Rows.Remove(Me.rowSchBar)
                End While

                ''geser schedule schedule pada workunit asal yang dipindah dan yang setelahnya pada DataTable dan update ke DB
                For Each Me.RowDgv In Me.dgvItem.Rows
                    If Me.RowDgv.Cells(Me.ColModify.Index).Value = Boolean.TrueString Then
                        Me.TblSchBar.DefaultView.RowFilter = "[Primary Work Unit] = " & Me.PrimaryWU.ToString
                        Me.TblSchBar.DefaultView.Sort = "[plan start time]"
                        For i As Integer = 0 To Me.TblSchBar.DefaultView.Count - 1
                            Me.rowSchBar = Me.TblSchBar.DefaultView(i).Row
                            If i > 0 Then
                                If Me.rowSchBar.delay = 0 Then
                                    Me.rowSchBar.plan_start_time = PreviousRowPlanEndTime
                                Else
                                    Me.rowSchBar.plan_start_time = Me.ObjSch.GetEndTime(Me.PrimaryWU, PreviousRowPlanEndTime, Me.rowSchBar.delay, Me.TblSchBarTanpaWO)
                                End If
                            End If
                            Me.rowSchBar.plan_end_time = Me.ObjSch.GetEndTime(Me.PrimaryWU, Me.rowSchBar.plan_start_time, Me.rowSchBar.total_work_hour, Me.TblSchBarTanpaWO)
                            PreviousRowPlanEndTime = Me.rowSchBar.plan_end_time
                            Me.ObjSch.UpdateScheduleAndQty(Me.rowSchBar.plan_start_time, Me.rowSchBar.plan_end_time, Me.rowSchBar.qty_produce, Me.rowSchBar.total_work_hour, Me.rowSchBar.identifier_jo, Me.rowSchBar.identifier_wo, Me.PrimaryWU)
                            Me.ObjSch.UpdatePlanStartEndScheduleWO(Me.rowSchBar.identifier_jo, Me.rowSchBar.identifier_wo) ''rosadin-r22
                        Next
                        Exit For
                    End If
                Next

                ''update schedule pada work unit tujuan berdasarkan perubahan urutan user terakhir pada DataTable dan update DB
                Me.TblSchBar2.DefaultView.RowFilter = "[Primary Work Unit] = " & Me.TxtPrimaryWU.Text
                Me.TblSchBar2.DefaultView.Sort = "[plan start time]"
                For Each DV As Data.DataRowView In Me.TblSchBar2.DefaultView
                    RowSchBar2 = DV.Row
                    Me.TblSchBar.DefaultView.RowFilter = "[Primary MO]=" & RowSchBar2.primary_mo.ToString & " AND [Identifier JO]='" & RowSchBar2.identifier_jo & "' AND [Identifier WO]='" & RowSchBar2.identifier_wo & "' AND [Primary Work Unit]=" & Me.TxtPrimaryWU.Text
                    If Me.TblSchBar.DefaultView.Count > 0 Then
                        Me.rowSchBar = Me.TblSchBar.DefaultView(0).Row
                        Me.rowSchBar.plan_start_time = RowSchBar2.plan_start_time
                        Me.rowSchBar.plan_end_time = RowSchBar2.plan_end_time
                        Me.rowSchBar.qty_produce = RowSchBar2.qty_produce
                        Me.rowSchBar.total_work_hour = RowSchBar2.total_work_hour
                        Me.ObjSch.UpdateScheduleAndQty(Me.rowSchBar.plan_start_time, Me.rowSchBar.plan_end_time, Me.rowSchBar.qty_produce, Me.rowSchBar.total_work_hour, Me.rowSchBar.identifier_jo, Me.rowSchBar.identifier_wo, Me.TxtPrimaryWU.Text)
                        Me.ObjSch.UpdatePlanStartEndScheduleWO(Me.rowSchBar.identifier_jo, Me.rowSchBar.identifier_wo) ''rosadin-r22
                    Else
                        MsgBox("Logic Error, inconsistent TblSchBar2 with TblSchBar.", MsgBoxStyle.Critical, "Err")
                    End If
                Next

                'tambahkan sch-sch baru yang dipindahkan, pada work unit tujuan pada DataTable dan DB
                frmRate.TblSchBar.DefaultView.RowFilter = "[Primary Work Unit] = " & Me.TxtPrimaryWU.Text & " AND ([Qty Sisa] > 0 OR ([Qty Sisa] = 0 AND [Merge] = False))"
                frmRate.TblSchBar.DefaultView.Sort = "[plan start time]"
                For Each DV2 As Data.DataRowView In frmRate.TblSchBar.DefaultView
                    rowSchBarRate = DV2.Row
                    rowSchBar = Me.TblSchBar.Newtbl_sch_barRow
                    For i As Integer = 0 To TblSchBar.Columns.Count - 1
                        rowSchBar.Item(i) = rowSchBarRate.Item(i)
                    Next
                    Me.TblSchBar.Rows.Add(rowSchBar)
                    Me.ObjSch.InsertNewSchedule(rowSchBar)
                    Me.ObjSch.UpdatePlanStartEndScheduleWO(rowSchBar.identifier_jo, rowSchBar.identifier_wo) ''rosadin-r22
                Next

                Me.DialogResult = Windows.Forms.DialogResult.OK
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnFindWU_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindWU.Click
        Dim frm As New Core_Module.FrmMasterWorkUnit_Cari
        Try
            frm.txtCari.Text = Me.TxtWUCode.Text
            frm.ShowDialog()
            If frm.PrimaryWorkUnit > 0 Then
                If frm.PrimaryWorkUnit <> Me.PrimaryWU Then
                    Me.TxtPrimaryWU.Text = frm.PrimaryWorkUnit
                    Me.TxtWUCode.Text = Me.objWorkUnit.GetCode(frm.PrimaryWorkUnit)
                    Me.TxtWUDesc.Text = Me.objWorkUnit.GetDesc(frm.PrimaryWorkUnit)
                    Me.TxtWUDesc.Focus()
                Else
                    Me.TxtPrimaryWU.Text = "-1"
                    Me.TxtWUCode.Text = String.Empty
                    Me.TxtWUDesc.Text = String.Empty
                    MsgBox("Select another work unit.", MsgBoxStyle.Information, "Information")
                End If
            Else
                Me.TxtPrimaryWU.Text = "-1"
                Me.TxtWUCode.Text = String.Empty
                Me.TxtWUDesc.Text = String.Empty
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtWUCode_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtWUCode.Enter
        Try
            Me.Buffer = Me.TxtWUCode.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtWUCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtWUCode.GotFocus
        Try
            Me.TxtWUCode.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtWUCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtWUCode.KeyDown
        Try
            If e.KeyCode = 13 Then
                Me.TxtWUCode.Text = Me.TxtWUCode.Text.Trim.Replace("'", "").Replace("%", "").Replace("*", "")
                If Me.objWorkUnit.ApakahSudahAda(Me.TxtWUCode.Text) Then
                    If Me.objWorkUnit.GetPrimary(Me.TxtWUCode.Text) <> Me.PrimaryWU Then
                        Me.TxtPrimaryWU.Text = Me.objWorkUnit.GetPrimary(Me.TxtWUCode.Text)
                        Me.TxtWUDesc.Text = Me.objWorkUnit.GetDesc(Me.TxtPrimaryWU.Text)
                        Me.TxtWUDesc.Focus()
                    Else
                        MsgBox("Select another work unit.", MsgBoxStyle.Information, "Information")
                        Me.TxtWUCode.Text = Me.Buffer
                        Exit Sub
                    End If
                Else
                    Me.BtnFindWU.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtWUCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtWUCode.Leave
        Try
            Me.TxtWUCode.Text = Me.TxtWUCode.Text.Trim.Replace("'", "").Replace("%", "").Replace("*", "")
            If Me.TxtWUCode.Text <> String.Empty Then
                If Me.objWorkUnit.ApakahSudahAda(Me.TxtWUCode.Text) Then
                    If Me.objWorkUnit.GetPrimary(Me.TxtWUCode.Text) <> Me.PrimaryWU Then
                        Me.TxtPrimaryWU.Text = Me.objWorkUnit.GetPrimary(Me.TxtWUCode.Text)
                        Me.TxtWUDesc.Text = Me.objWorkUnit.GetDesc(Me.TxtPrimaryWU.Text)
                    Else
                        MsgBox("Select another work unit.", MsgBoxStyle.Information, "Information")
                        Me.TxtWUCode.Text = Me.Buffer
                        Exit Sub
                    End If
                Else
                    Me.BtnFindWU.PerformClick()
                End If
            Else
                Me.TxtPrimaryWU.Text = "-1"
                Me.TxtWUCode.Text = String.Empty
                Me.TxtWUDesc.Text = String.Empty
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtWUCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtWUCode.LostFocus
        Try
            Me.TxtWUCode.BackColor = Color.White
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TabControl1_Selecting(ByVal sender As Object, ByVal e As System.Windows.Forms.TabControlCancelEventArgs) Handles TabControl1.Selecting
        Try
            If Not Me.Navigate Then
                e.Cancel = True
                Exit Sub
            Else
                Me.Navigate = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub LoadDgv1()
        Dim rowSchBar As DseTrans_Schedule.tbl_sch_barRow
        Dim rowDGV As DataGridViewRow
        Try
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.IsikanDataCodeNoDescUnitCode(TblSchBar, Me.PrimaryWU)
            TblSchBar.DefaultView.RowFilter = "[primary work unit] = " + Me.PrimaryWU.ToString
            TblSchBar.DefaultView.Sort = "[plan start time]"
            Me.dgvItem.Rows.Clear()
            For Each DV As Data.DataRowView In Me.TblSchBar.DefaultView
                rowSchBar = DV.Row
                Me.dgvItem.Rows.Add()
                rowDGV = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
                With rowDGV
                    .Cells(Me.ColModify.Index).Value = Boolean.FalseString
                    .Cells(Me.ColMoNo.Index).Value = rowSchBar.mo_no
                    .Cells(Me.ColJONo.Index).Value = rowSchBar.jo_no
                    .Cells(Me.ColProcessCode.Index).Value = rowSchBar.process_code
                    .Cells(Me.ColWONo.Index).Value = rowSchBar.wo_no
                    .Cells(Me.ColCodeNo.Index).Value = rowSchBar.product_code
                    .Cells(Me.ColDesc.Index).Value = rowSchBar.product_desc
                    .Cells(Me.ColRate.Index).Value = Me.objCD.DigitGrouping(rowSchBar.rate, rowSchBar.decimal_point)
                    .Cells(Me.ColQty.Index).Value = Me.objCD.DigitGrouping(rowSchBar.qty_produce, rowSchBar.decimal_point)
                    .Cells(Me.ColUnit.Index).Value = rowSchBar.unit_code
                    .Cells(Me.ColWH.Index).Value = Me.objCD.Convert_ToString(rowSchBar.total_work_hour)
                    .Cells(Me.ColStartDay.Index).Value = Me.objCD.GetHari_English(rowSchBar.plan_start_time.Date.DayOfWeek)
                    .Cells(Me.ColStartDate.Index).Value = Format(rowSchBar.plan_start_time, "dd MMM yyyy")
                    .Cells(Me.ColStartHour.Index).Value = Format(rowSchBar.plan_start_time, "HH.mm")
                    .Cells(Me.ColEndDay.Index).Value = Me.objCD.GetHari_English(rowSchBar.plan_end_time.Date.DayOfWeek)
                    .Cells(Me.ColEndDate.Index).Value = Format(rowSchBar.plan_end_time, "dd MMM yyyy")
                    .Cells(Me.ColEndHour.Index).Value = Format(rowSchBar.plan_end_time, "HH.mm")
                    .Cells(Me.ColProgress.Index).Value = Me.objCD.DigitGrouping(rowSchBar.progress, 2)
                    .Cells(Me.ColPrimaryMO.Index).Value = rowSchBar.primary_mo
                    .Cells(Me.ColIdentifierJO.Index).Value = rowSchBar.identifier_jo
                    .Cells(Me.ColPrimaryProcess.Index).Value = rowSchBar.primary_process
                    .Cells(Me.colIdentifierWO.Index).Value = rowSchBar.identifier_wo
                    .Cells(Me.ColPrimaryCodeNo.Index).Value = rowSchBar.primary_product
                    .Cells(Me.ColPrimaryUnit.Index).Value = rowSchBar.primary_unit
                    .Cells(Me.ColPlanStartTime.Index).Value = rowSchBar.plan_start_time
                End With
            Next
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvItem.ClearSelection()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub LoadDgv2()
        Dim rowSchBar As DseTrans_Schedule.tbl_sch_barRow
        Dim rowDGV As DataGridViewRow

        Dim dtaWU_Product As New DseTrans_ProductionSchedule_ChangeWUTableAdapters.tblmaster_product_item_workunitTableAdapter
        Dim tblWU_Product As New DseTrans_ProductionSchedule_ChangeWU.tblmaster_product_item_workunitDataTable
        Dim listOfNotHandledCode As New List(Of String)

        Dim lastEndTime As Date = #1/1/1950#
        Dim endTimeBefore As Date = #1/1/1950#
        Try
            ''inisial awal, kosongkan dgv dan set new sch di ujung sch tujuan
            Me.ChkSetAsLast.Checked = True
            Me.DgvItem2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DgvItem2.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.DgvItem2.Rows.Clear()

            ''cek apakah work unit yang dipilih bisa meng-handle produk2 yang dipilih.
            dtaWU_Product.Connection.ConnectionString = Core_Module.GetConstring
            dtaWU_Product.ClearBeforeFill = False
            tblWU_Product.Clear()
            For Each Me.RowDgv In Me.dgvItem.Rows
                If Me.RowDgv.Cells(Me.ColModify.Index).Value = Boolean.TrueString Then
                    If dtaWU_Product.Fill(tblWU_Product, Me.RowDgv.Cells(Me.ColPrimaryCodeNo.Index).Value, Me.TxtPrimaryWU.Text) <= 0 Then
                        listOfNotHandledCode.Add(Me.RowDgv.Cells(Me.ColCodeNo.Index).Value)
                    End If
                End If
            Next
            If listOfNotHandledCode.Count > 0 Then
                Dim notHandledCode As String = ""
                For Each s As String In listOfNotHandledCode
                    If notHandledCode = "" Then
                        notHandledCode = s
                    Else
                        notHandledCode = notHandledCode & vbNewLine & s
                    End If
                Next
                MsgBox("The following product code is not produced by work unit " & Me.TxtWUCode.Text & "." & vbNewLine & notHandledCode, MsgBoxStyle.Information, "Information")
                Me.TxtPrimaryWU.Text = "-1"
                Me.TxtWUCode.Text = String.Empty
                Me.TxtWUDesc.Text = String.Empty
                Exit Sub
            End If

            ''change rate
            frmRate.TblSchBar.Rows.Clear()
            For Each Me.RowDgv In Me.dgvItem.Rows
                If Me.RowDgv.Cells(Me.ColModify.Index).Value = Boolean.TrueString Then
                    Me.TblSchBar.DefaultView.RowFilter = "[primary mo] = " & Me.RowDgv.Cells(Me.ColPrimaryMO.Index).Value & " and [identifier jo] = '" & Me.RowDgv.Cells(Me.ColIdentifierJO.Index).Value & "' and [identifier wo] = '" & Me.RowDgv.Cells(Me.colIdentifierWO.Index).Value & "' and [primary work unit] = " & Me.PrimaryWU.ToString & " and [tipe] = " & CInt(ClsTrans_Schedule.EnumTipeSchBar.work_time).ToString
                    If Me.TblSchBar.DefaultView.Count > 0 Then
                        For Each DV As Data.DataRowView In Me.TblSchBar.DefaultView
                            rowSchBar = DV.Row
                            rowSchBarRate = frmRate.TblSchBar.Newtbl_sch_barRow
                            For i As Integer = 0 To TblSchBar.Columns.Count - 1
                                rowSchBarRate.Item(i) = rowSchBar.Item(i)
                            Next
                            rowSchBarRate.primary_work_unit = Me.TxtPrimaryWU.Text
                            rowSchBarRate.work_unit_code = Me.objWorkUnit.GetCode(Me.TxtPrimaryWU.Text)
                            rowSchBarRate.primary_input_location = Me.objWorkUnit.GetPrimaryInputLocation(Me.TxtPrimaryWU.Text)
                            rowSchBarRate.input_location = Me.objLocation.GetStorageAreaCode(rowSchBarRate.primary_input_location)
                            rowSchBarRate.primary_output_location = Me.objWorkUnit.GetPrimaryOutputLocation(Me.TxtPrimaryWU.Text)
                            rowSchBarRate.output_location = Me.objLocation.GetStorageAreaCode(rowSchBarRate.primary_output_location)
                            frmRate.TblSchBar.Rows.Add(rowSchBarRate)
                        Next
                    End If
                End If
            Next
            frmRate.PrimaryWU_Destination = Me.TxtPrimaryWU.Text
            frmRate.UserCode = Me.UserCode
            frmRate.BackColorAll = Me.BackColorAll
            If frmRate.ShowDialog() = Windows.Forms.DialogResult.OK Then
                Me.LoadDataToDgv2(lastEndTime, endTimeBefore)
                Me.TxtWUDesc.SelectedText = ""
            Else
                Me.TxtPrimaryWU.Text = -1
                Me.TxtWUCode.Text = String.Empty
                Me.TxtWUDesc.Text = String.Empty
                Me.DgvItem2.Rows.Clear()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub LoadDataToDgv2(ByVal lastEndTime As Date, ByVal endTimeBefore As Date)
        Dim rowSchBarRate As DseTrans_Schedule.tbl_sch_barRow
        Dim rowSchBar2 As DseTrans_Schedule.tbl_sch_barRow
        Dim FirstLoop As Boolean = True
        Dim RowIndex As Integer = -1
        Dim PreviousRowPlanEndTime As Date = #1/1/1950#
        Try
            Me.IsikanDataCodeNoDescUnitCode(TblSchBar, Me.TxtPrimaryWU.Text)
            TblSchBar2 = TblSchBar.Copy

            ''tambahkan kolom2 yang diperlukan
            If TblSchBar2.Columns.Contains("orginal plan start time") = False Then
                TblSchBar2.Columns.Add("original plan start time", System.Type.GetType("System.DateTime"))
                TblSchBar2.Columns("original plan start time").DefaultValue = #1/1/1950#
            End If
            If TblSchBar2.Columns.Contains("Delay2") = False Then
                TblSchBar2.Columns.Add("Delay2", System.Type.GetType("System.Double"))
                TblSchBar2.Columns("Delay2").DefaultValue = 0
            End If
            If TblSchBar2.Columns.Contains("RowIndex") = False Then
                TblSchBar2.Columns.Add("RowIndex", System.Type.GetType("System.Double"))
                TblSchBar2.Columns("RowIndex").DefaultValue = 0
            End If
            If frmRate.TblSchBar.Columns.Contains("Delay2") = False Then
                frmRate.TblSchBar.Columns.Add("Delay2", System.Type.GetType("System.Double"))
                frmRate.TblSchBar.Columns("Delay2").DefaultValue = 0
            End If
            If frmRate.TblSchBar.Columns.Contains("Merge") = False Then
                frmRate.TblSchBar.Columns.Add("Merge", System.Type.GetType("System.Boolean"))
                frmRate.TblSchBar.Columns("Merge").DefaultValue = False
            End If
            For Each rowSchBarRate In frmRate.TblSchBar
                rowSchBarRate.Item("Merge") = False
            Next

            ''hitung delay pada workunit tujuan, tandai index juga, kalo ga, biasanya ngaco saat re-schedule with sort by plan start time
            TblSchBar2.DefaultView.RowFilter = "[primary work unit] = " + Me.TxtPrimaryWU.Text
            TblSchBar2.DefaultView.Sort = "[plan start time]"
            FirstLoop = True
            For Each DV As Data.DataRowView In TblSchBar2.DefaultView
                RowIndex += 1
                rowSchBar2 = DV.Row
                If FirstLoop Then
                    FirstLoop = False
                    rowSchBar2.Item("Delay2") = 0
                Else
                    If PreviousRowPlanEndTime = rowSchBar2.plan_start_time Then
                        rowSchBar2.Item("Delay2") = 0
                    Else
                        rowSchBar2.Item("Delay2") = Me.ObjSch.GetWorkDuration(PreviousRowPlanEndTime, rowSchBar2.plan_start_time, Me.TxtPrimaryWU.Text, Me.TblSchBarTanpaWO)
                    End If
                End If
                rowSchBar2.Item("RowIndex") = RowIndex
                PreviousRowPlanEndTime = rowSchBar2.plan_end_time
            Next

            ''cari apakah ada schedule yang sama (JO-WO) pada workunit tujuan, jika ada, tambahkan qty dan hitung ulang W/H jadwal setelahnya
            frmRate.TblSchBar.DefaultView.RowFilter = ""
            frmRate.TblSchBar.DefaultView.Sort = "[plan start time]"
            For Each DV As Data.DataRowView In frmRate.TblSchBar.DefaultView
                rowSchBarRate = DV.Row
                TblSchBar2.DefaultView.RowFilter = "[primary work unit] = " + Me.TxtPrimaryWU.Text
                TblSchBar2.DefaultView.Sort = "[RowIndex]"
                Dim ViewIndex As Integer = -1
                For Each DV2 As Data.DataRowView In TblSchBar2.DefaultView
                    ViewIndex += 1
                    rowSchBar2 = DV2.Row
                    If rowSchBar2.identifier_jo = rowSchBarRate.identifier_jo And rowSchBar2.identifier_wo = rowSchBarRate.identifier_wo Then
                        rowSchBarRate.Item("Merge") = True
                        FirstLoop = True
                        For ViewIndex = ViewIndex To TblSchBar2.DefaultView.Count - 1
                            rowSchBar2 = TblSchBar2.DefaultView(ViewIndex).Row
                            rowSchBar2.BeginEdit()
                            If FirstLoop Then
                                FirstLoop = False
                                rowSchBar2.qty_produce += rowSchBarRate.qty_produce
                                rowSchBar2.total_work_hour = Me.objCD.DigitGrouping(rowSchBar2.qty_produce / rowSchBar2.rate * 60, 0, Core_Module.ClsCekData.EnumTipeRounding.Up)
                            Else
                                If rowSchBar2.Item("Delay2") = 0 Then
                                    rowSchBar2.plan_start_time = PreviousRowPlanEndTime
                                Else
                                    rowSchBar2.plan_start_time = Me.ObjSch.GetEndTime(Me.TxtPrimaryWU.Text, PreviousRowPlanEndTime, rowSchBar2.Item("Delay2"), Me.TblSchBarTanpaWO)
                                End If
                            End If
                            rowSchBar2.plan_end_time = Me.ObjSch.GetEndTime(Me.TxtPrimaryWU.Text, rowSchBar2.plan_start_time, rowSchBar2.total_work_hour, Me.TblSchBarTanpaWO)
                            PreviousRowPlanEndTime = rowSchBar2.plan_end_time
                            rowSchBar2.EndEdit()
                        Next
                        Exit For
                    End If
                Next
            Next

            ''show schedule workunit tujuan pada datagridview setelah di-merge 
            TblSchBar2.DefaultView.RowFilter = "[primary work unit] = " + Me.TxtPrimaryWU.Text
            TblSchBar2.DefaultView.Sort = "[plan start time]"
            For Each DV As Data.DataRowView In TblSchBar2.DefaultView
                rowSchBar2 = DV.Row
                rowSchBar2.Item("original plan start time") = rowSchBar2.plan_start_time
                lastEndTime = rowSchBar2.plan_end_time
                RowDgv = Me.DgvItem2.Rows(Me.DgvItem2.Rows.Add())
                With RowDgv
                    .Cells(Me.Col2Insert.Index).Value = Boolean.FalseString
                    .Cells(Me.Col2IdentifierJO.Index).Value = rowSchBar2.identifier_jo
                    .Cells(Me.Col2IdentifierWO.Index).Value = rowSchBar2.identifier_wo
                    .Cells(Me.Col2MO_No.Index).Value = rowSchBar2.mo_no
                    .Cells(Me.Col2JO_No.Index).Value = rowSchBar2.jo_no
                    .Cells(Me.Col2ProcessCode.Index).Value = rowSchBar2.process_code
                    .Cells(Me.Col2WO_No.Index).Value = rowSchBar2.wo_no
                    .Cells(Me.Col2CodeNo.Index).Value = rowSchBar2.product_code
                    .Cells(Me.Col2Desc.Index).Value = rowSchBar2.product_desc
                    .Cells(Me.Col2Rate.Index).Value = Me.objCD.DigitGrouping(rowSchBar2.rate, rowSchBar2.decimal_point)
                    .Cells(Me.Col2Qty.Index).Value = Me.objCD.DigitGrouping(rowSchBar2.qty_produce, rowSchBar2.decimal_point)
                    .Cells(Me.Col2UnitCode.Index).Value = rowSchBar2.unit_code
                    .Cells(Me.Col2WorkHour.Index).Value = Me.objCD.Convert_ToString(rowSchBar2.total_work_hour)
                    .Cells(Me.Col2StartDay.Index).Value = Me.objCD.GetHari_English(rowSchBar2.plan_start_time.Date.DayOfWeek)
                    .Cells(Me.Col2StartDate.Index).Value = Format(rowSchBar2.plan_start_time, "dd MMM yyyy")
                    .Cells(Me.Col2StartTime.Index).Value = Format(rowSchBar2.plan_start_time, "HH.mm")
                    .Cells(Me.Col2EndDay.Index).Value = Me.objCD.GetHari_English(rowSchBar2.plan_end_time.Date.DayOfWeek)
                    .Cells(Me.Col2EndDate.Index).Value = Format(rowSchBar2.plan_end_time, "dd MMM yyyy")
                    .Cells(Me.Col2EndTime.Index).Value = Format(rowSchBar2.plan_end_time, "HH.mm")
                    .Cells(Me.Col2Progress.Index).Value = Me.objCD.DigitGrouping(rowSchBar2.progress, 2)
                    .Cells(Me.Col2PrimaryMO.Index).Value = rowSchBar2.primary_mo
                    .Cells(Me.Col2PrimaryProcess.Index).Value = rowSchBar2.primary_process
                    .Cells(Me.Col2PrimaryCodeNo.Index).Value = rowSchBar2.primary_product
                    .Cells(Me.Col2PrimaryUnit.Index).Value = rowSchBar2.primary_unit
                    .Cells(Me.Col2ReadOnly.Index).Value = 0
                    .Cells(Me.Col2PlanStartTime.Index).Value = rowSchBar2.plan_start_time
                End With
            Next

            ''hitung delay pada shcedule yang di pindahkan
            frmRate.TblSchBar.DefaultView.RowFilter = ""
            frmRate.TblSchBar.DefaultView.Sort = "[plan start time]"
            FirstLoop = True
            For Each DV As Data.DataRowView In frmRate.TblSchBar.DefaultView
                rowSchBarRate = DV.Row
                If FirstLoop Then
                    FirstLoop = False
                    rowSchBarRate.Item("Delay2") = 0
                Else
                    If PreviousRowPlanEndTime = rowSchBarRate.plan_start_time Then
                        rowSchBarRate.Item("Delay2") = 0
                    Else
                        rowSchBarRate.Item("Delay2") = Me.ObjSch.GetWorkDuration(PreviousRowPlanEndTime, rowSchBarRate.plan_start_time, Me.TxtPrimaryWU.Text, Me.TblSchBarTanpaWO)
                    End If
                End If
                PreviousRowPlanEndTime = rowSchBarRate.plan_end_time
            Next

            '''hitung ulang jadwal
            If lastEndTime = #1/1/1950# Then
                lastEndTime = Me.ObjSch.GetStartTime_ByWorkUnit(Me.TblSchBarTanpaWO, Me.TxtPrimaryWU.Text, Me.ObjSch.GetFirmPlan)
            End If
            frmRate.TblSchBar.DefaultView.RowFilter = "[Merge] = False"
            frmRate.TblSchBar.DefaultView.Sort = "[plan start time]"
            FirstLoop = True
            For Each DV As Data.DataRowView In frmRate.TblSchBar.DefaultView
                rowSchBarRate = DV.Row
                If FirstLoop Then
                    FirstLoop = False
                    rowSchBarRate.plan_start_time = lastEndTime
                Else
                    If rowSchBarRate.Item("Delay2") = 0 Then
                        rowSchBarRate.plan_start_time = PreviousRowPlanEndTime
                    Else
                        rowSchBarRate.plan_start_time = Me.ObjSch.GetEndTime(Me.TxtPrimaryWU.Text, PreviousRowPlanEndTime, rowSchBarRate.Item("Delay2"), Me.TblSchBarTanpaWO)
                    End If
                End If
                rowSchBarRate.plan_end_time = Me.ObjSch.GetEndTime(Me.TxtPrimaryWU.Text, rowSchBarRate.plan_start_time, rowSchBarRate.total_work_hour, Me.TblSchBarTanpaWO)
                PreviousRowPlanEndTime = rowSchBarRate.plan_end_time
            Next

            ''tampilkan sch sumber di akhir sch tujuan (kecuali yang di-merge)
            frmRate.TblSchBar.DefaultView.RowFilter = "[Merge] = False"
            frmRate.TblSchBar.DefaultView.Sort = "[plan start time]"
            For Each DV As Data.DataRowView In frmRate.TblSchBar.DefaultView
                rowSchBarRate = DV.Row
                RowDgv = Me.DgvItem2.Rows(Me.DgvItem2.Rows.Add())
                With RowDgv
                    .Cells(Me.Col2Insert.Index).Value = Boolean.FalseString
                    .Cells(Me.Col2IdentifierJO.Index).Value = rowSchBarRate.identifier_jo
                    .Cells(Me.Col2IdentifierWO.Index).Value = rowSchBarRate.identifier_wo
                    .Cells(Me.Col2MO_No.Index).Value = rowSchBarRate.mo_no
                    .Cells(Me.Col2JO_No.Index).Value = rowSchBarRate.jo_no
                    .Cells(Me.Col2ProcessCode.Index).Value = rowSchBarRate.process_code
                    .Cells(Me.Col2WO_No.Index).Value = rowSchBarRate.wo_no
                    .Cells(Me.Col2CodeNo.Index).Value = rowSchBarRate.product_code
                    .Cells(Me.Col2Desc.Index).Value = rowSchBarRate.product_desc
                    .Cells(Me.Col2Rate.Index).Value = Me.objCD.DigitGrouping(rowSchBarRate.rate, rowSchBarRate.decimal_point)
                    .Cells(Me.Col2Qty.Index).Value = Me.objCD.DigitGrouping(rowSchBarRate.qty_produce, rowSchBarRate.decimal_point)
                    .Cells(Me.Col2UnitCode.Index).Value = rowSchBarRate.unit_code
                    .Cells(Me.Col2WorkHour.Index).Value = Me.objCD.Convert_ToString(rowSchBarRate.total_work_hour)
                    .Cells(Me.Col2StartDay.Index).Value = Me.objCD.GetHari_English(rowSchBarRate.plan_start_time.Date.DayOfWeek)
                    .Cells(Me.Col2StartDate.Index).Value = Format(rowSchBarRate.plan_start_time, "dd MMM yyyy")
                    .Cells(Me.Col2StartTime.Index).Value = Format(rowSchBarRate.plan_start_time, "HH.mm")
                    .Cells(Me.Col2EndDay.Index).Value = Me.objCD.GetHari_English(rowSchBarRate.plan_end_time.Date.DayOfWeek)
                    .Cells(Me.Col2EndDate.Index).Value = Format(rowSchBarRate.plan_end_time, "dd MMM yyyy")
                    .Cells(Me.Col2EndTime.Index).Value = Format(rowSchBarRate.plan_end_time, "HH.mm")
                    .Cells(Me.Col2Progress.Index).Value = Me.objCD.DigitGrouping(rowSchBarRate.progress, 2)
                    .Cells(Me.Col2PrimaryMO.Index).Value = rowSchBarRate.primary_mo
                    .Cells(Me.Col2PrimaryProcess.Index).Value = rowSchBarRate.primary_process
                    .Cells(Me.Col2PrimaryCodeNo.Index).Value = rowSchBarRate.primary_product
                    .Cells(Me.Col2PrimaryUnit.Index).Value = rowSchBarRate.primary_unit
                    .Cells(Me.Col2ReadOnly.Index).Value = 1
                    .Cells(Me.Col2PlanStartTime.Index).Value = rowSchBarRate.plan_start_time
                    .DefaultCellStyle.BackColor = Me.BackColorSource
                End With
            Next
            Me.DgvItem2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.DgvItem2.ClearSelection()
            Me.FirstCellEndEdit = True

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanDataCodeNoDescUnitCode(ByVal vTblSchBar As DseTrans_Schedule.tbl_sch_barDataTable, ByVal vPrimaryWU As Double)
        Dim PrimaryMOMin As Double = Double.MaxValue
        Dim PrimaryMOMax As Double = Double.MinValue
        Dim rowSchBar As DseTrans_Schedule.tbl_sch_barRow
        Dim DTA As New DseTrans_ScheduleTableAdapters.tbltrans_mo_item_woTableAdapter
        Dim TBL As New DseTrans_Schedule.tbltrans_mo_item_woDataTable
        Dim row As DseTrans_Schedule.tbltrans_mo_item_woRow
        Dim ListPrimaryCodeNo As New List(Of Double)
        Dim ListPrimaryUnit As New List(Of Double)
        Dim DTAUnit As New DseTrans_ScheduleTableAdapters.tblmaster_master_unitTableAdapter
        Dim TBLUnit As New DseTrans_Schedule.tblmaster_master_unitDataTable
        Dim rowUnit As DseTrans_Schedule.tblmaster_master_unitRow
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTAUnit.Connection.ConnectionString = DTA.Connection.ConnectionString

            vTblSchBar.DefaultView.RowFilter = "[primary work unit] = " + vPrimaryWU.ToString + " AND tipe = " + CInt(ClsTrans_Schedule.EnumTipeSchBar.work_time).ToString
            vTblSchBar.DefaultView.Sort = "[plan start time]"

            If vTblSchBar.DefaultView.Count = 0 Then
                Exit Sub
            End If
            For Each DV As Data.DataRowView In vTblSchBar.DefaultView
                rowSchBar = DV.Row
                If rowSchBar.primary_mo > PrimaryMOMax Then
                    PrimaryMOMax = rowSchBar.primary_mo
                End If
                If PrimaryMOMin > rowSchBar.primary_mo Then
                    PrimaryMOMin = rowSchBar.primary_mo
                End If
            Next

            DTA.Fill(TBL, PrimaryMOMin, PrimaryMOMax)
            DTAUnit.Fill(TBLUnit)
            For Each DV As Data.DataRowView In vTblSchBar.DefaultView
                rowSchBar = DV.Row
                For Each row In TBL.Rows
                    If rowSchBar.identifier_jo = row.Identifier_JO And rowSchBar.identifier_wo = row.Identifier_WO Then
                        rowSchBar.product_code = row.Product_Code
                        rowSchBar.product_desc = row.Description
                        rowSchBar.primary_product = row.Primary_Product
                        Exit For
                    End If
                Next
                For Each rowUnit In TBLUnit.Rows
                    If rowSchBar.primary_unit = rowUnit.Primary Then
                        rowSchBar.unit_code = rowUnit.Unit_Code
                        Exit For
                    End If
                Next
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtPrimaryWU_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPrimaryWU.TextChanged
        Try
            If Me.TxtPrimaryWU.Text = "" Or Me.TxtPrimaryWU.Text = "-1" Then
                Me.DgvItem2.Rows.Clear()
            ElseIf Me.TxtPrimaryWU.Text <> Me.Buffer Then
                Me.Buffer = Me.TxtPrimaryWU.Text
                Me.LoadDgv2()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Me.objCD.UpdateThreading()
            Me.Cursor = Cursors.WaitCursor
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Me.ReOrderSchedule()
            Me.ChkSetAsLast.Checked = False
            Me.Cursor = Cursors.Default
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkSetAsLast_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkSetAsLast.Click
        Dim AnyCheck As Boolean = False
        Try
            If Me.ChkSetAsLast.Checked Then
                Me.DgvItem2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.DgvItem2.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                Me.DgvItem2.Rows.Clear()
                Me.LoadDataToDgv2(#1/1/1950#, #1/1/1950#)
            Else
                For Each Me.RowDgv In Me.DgvItem2.Rows
                    If Me.RowDgv.Cells(Me.Col2Insert.Index).Value = Boolean.TrueString Then
                        AnyCheck = True
                    End If
                Next
                If AnyCheck = False Then
                    Me.ChkSetAsLast.Checked = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Try
            Me.Label2.Text = "Loading schedule..."
            Me.LoadTblSch(Me.ObjSch.GetFirmPlan(), #12/31/2200#)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadTblSch(ByVal vStartDate As Date, ByVal vEndDate As Date)
        Dim MOMin, MOMax As Double
        Dim DTA As New DseTrans_ScheduleTableAdapters.DataTable2TableAdapter
        Dim TBL As New DseTrans_Schedule.DataTable2DataTable
        Dim row As DseTrans_Schedule.DataTable2Row
        Dim mn As Double = 0
        Dim jk As Double = 0
        Dim rowSch1 As DseTrans_Schedule.tbl_sch_barRow
        Dim DTAWU As New DseTrans_ScheduleTableAdapters.tblmaster_workunitTableAdapter
        Dim TBLWU As New DseTrans_Schedule.tblmaster_workunitDataTable
        Dim rowWU As DseTrans_Schedule.tblmaster_workunitRow
        Dim DTAWU1 As New DseTrans_ScheduleTableAdapters.tblmaster_product_item_workunit1TableAdapter
        Dim TBLWU1 As New DseTrans_Schedule.tblmaster_product_item_workunit1DataTable
        Dim StartTime As Date = Now
        Dim Ketemu As Boolean = False
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTAWU.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTAWU1.Connection.ConnectionString = DTA.Connection.ConnectionString

            Me.TblSchBar.Rows.Clear()
            Me.ObjSch.LoadScheduleBar_WOItemOnly(vStartDate, vEndDate, Me.TblSchBar)
            Me.ObjSch.LoadScheduleBar_TanpaScheduledItem(vStartDate, vEndDate, Me.TblSchBarTanpaWO)

            Dim ij As Integer = 0
            While ij <= TblSchBar.Rows.Count - 1
                rowSch = TblSchBar.Rows(ij)
                If rowSch.tipe <> ClsTrans_Schedule.EnumTipeSchBar.work_time Then
                    Me.TblSchBar.Rows.Remove(rowSch)
                Else
                    ij += 1
                End If
            End While

            'isikan nilai yang dibutuhkan
            If Me.TblSchBar.Rows.Count > 0 Then
                MOMin = Double.MaxValue
                MOMax = Double.MinValue
                For Each rowSch In Me.TblSchBar.Rows
                    If rowSch.primary_mo <> -1 Then
                        If rowSch.primary_mo < MOMin Then
                            MOMin = rowSch.primary_mo
                        End If
                        If rowSch.primary_mo > MOMax Then
                            MOMax = rowSch.primary_mo
                        End If
                    End If
                Next

                DTA.Fill(TBL, MOMax, MOMin, vStartDate)

                While mn <= Me.TblSchBar.Rows.Count - 1
                    Ketemu = False
                    rowSch = Me.TblSchBar.Rows(mn)
                    If rowSch.primary_mo <> -1 Then
                        If rowSch.wo_no = "" Then
                            For Each row In TBL.Rows
                                If row.Identifier_WO = rowSch.identifier_wo And rowSch.wo_no = "" Then
                                    rowSch.wo_no = row.WO_Number
                                    rowSch.jo_no = row.JO_Number
                                    rowSch.process_code = row.Process_Code
                                    rowSch.mo_no = row.MO_Number
                                    rowSch.unit_code = row.Unit_Code
                                    rowSch.is_closed = row.Is_Closed

                                    jk = mn + 1
                                    While jk <= Me.TblSchBar.Rows.Count - 1
                                        rowSch1 = Me.TblSchBar.Rows(jk)
                                        If rowSch1.identifier_wo = rowSch.identifier_wo Then
                                            rowSch1.wo_no = row.WO_Number
                                            rowSch1.jo_no = row.JO_Number
                                            rowSch1.process_code = row.Process_Code
                                            rowSch1.mo_no = row.MO_Number
                                            rowSch1.unit_code = row.Unit_Code
                                            rowSch1.is_closed = row.Is_Closed
                                        End If
                                        jk += 1
                                    End While

                                    TBL.Rows.Remove(row)
                                    Ketemu = True
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                    mn += 1
                    If Ketemu = False And rowSch.wo_no = "" Then
                        MsgBox("Logic error: additional data not found. Please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                    End If
                End While
                TBL.Rows.Clear()

                mn = 0
                DTAWU.Fill(TBLWU)
                While mn <= Me.TblSchBar.Rows.Count - 1
                    Ketemu = False
                    rowSch = Me.TblSchBar.Rows(mn)
                    If rowSch.identifier_wo <> "" Then
                        If rowSch.work_unit_code = "" Then
                            For Each rowWU In TBLWU.Rows
                                If rowWU.primary = rowSch.primary_work_unit And rowSch.work_unit_code = "" Then
                                    rowSch.work_unit_code = rowWU.work_unit_code

                                    jk = mn + 1
                                    While jk <= Me.TblSchBar.Rows.Count - 1
                                        rowSch1 = Me.TblSchBar.Rows(jk)
                                        If rowSch1.primary_work_unit = rowSch.primary_work_unit Then
                                            rowSch1.work_unit_code = rowWU.work_unit_code
                                        End If
                                        jk += 1
                                    End While

                                    Ketemu = True
                                    TBLWU.Rows.Remove(rowWU)
                                    Exit For
                                End If
                            Next
                        End If
                    End If

                    mn += 1
                    If Ketemu = False And rowSch.work_unit_code = "" Then
                        MsgBox("Logic error: work unit data not found. Please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                    End If
                End While
                TBLWU.Rows.Clear()

            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Try
            Me.LoadDgv1()
            Me.EnableControl(True)
            Me.Cursor = Cursors.Default
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class