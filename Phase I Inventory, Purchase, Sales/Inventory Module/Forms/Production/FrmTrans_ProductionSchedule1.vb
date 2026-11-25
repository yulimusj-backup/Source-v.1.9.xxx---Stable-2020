Option Explicit On
Option Strict Off

Partial Public Class FrmTrans_ProductionSchedule
    Dim Buffer_A As String = ""
    Dim Buffer_B As String = ""
    Dim Reload_TBLTanpaItemWO As Boolean = True

    Dim ShowWO, ShowCodeNo, ShowDesc, ShowDayStart, ShowDayFinish, ShowProcess, ShowProgress_, ShowRate, ShowQty, ShowUnit, ShowWH, ShowInputLocation, ShowOutputLocation, ShowActualGoodQty, ShowActualRejectQty, ShowActualProduceQty As String
    Sub setColumnVisibility_1()
        Dim vGroup As enumGroup = Me.cboGroupBy.SelectedIndex
        Try
            ShowWO = Me.objSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowWO, Me.UserCode)
            If ShowWO = "" Then ShowWO = "1"
            ShowCodeNo = Me.objSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowCodeNo, Me.UserCode)
            If ShowCodeNo = "" Then ShowCodeNo = "1"
            ShowDesc = Me.objSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowDesc, Me.UserCode)
            If ShowDesc = "" Then ShowDesc = "1"
            ShowDayStart = Me.objSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowDayStart, Me.UserCode)
            If ShowDayStart = "" Then ShowDayStart = "1"
            ShowDayFinish = Me.objSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowDayFinish, Me.UserCode)
            If ShowDayFinish = "" Then ShowDayFinish = "1"
            ShowProcess = Me.objSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowProcess, Me.UserCode)
            If ShowProcess = "" Then ShowProcess = "1"
            ShowProgress_ = Me.objSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowProgress, Me.UserCode)
            If ShowProgress_ = "" Then ShowProgress_ = "1"
            ShowRate = Me.objSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowRate, Me.UserCode)
            If ShowRate = "" Then ShowRate = "1"
            ShowQty = Me.objSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowQty, Me.UserCode)
            If ShowQty = "" Then ShowQty = "1"
            ShowUnit = Me.objSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowUnit, Me.UserCode)
            If ShowUnit = "" Then ShowUnit = "1"
            ShowWH = Me.objSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowWH, Me.UserCode)
            If ShowWH = "" Then ShowWH = "1"

            ShowInputLocation = Me.objSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowInputLocation, Me.UserCode)
            If ShowInputLocation = "" Then ShowInputLocation = "1"
            ShowOutputLocation = Me.objSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowOutputLocation, Me.UserCode)
            If ShowOutputLocation = "" Then ShowOutputLocation = "1"

            ShowActualGoodQty = Me.objSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowActualGoodQty, Me.UserCode)
            If ShowActualGoodQty = "" Then ShowActualGoodQty = "1"
            ShowActualRejectQty = Me.objSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowActualRejectQty, Me.UserCode)
            If ShowActualRejectQty = "" Then ShowActualRejectQty = "1"
            ShowActualProduceQty = Me.objSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowActualProduceQty, Me.UserCode)
            If ShowActualProduceQty = "" Then ShowActualProduceQty = "1"

            If ShowWO = "1" Then
                Me.ColLeft_WONo.Visible = True
            Else
                Me.ColLeft_WONo.Visible = False
            End If
            If ShowCodeNo = "1" Then
                Me.ColLeft_CodeNo.Visible = True
            Else
                Me.ColLeft_CodeNo.Visible = False
            End If
            If ShowDesc = "1" Then
                Me.ColLeft_Description.Visible = True
            Else
                Me.ColLeft_Description.Visible = False
            End If
            If ShowDayStart = "1" Then
                Me.ColLeft_DayOfWeekStart.Visible = True
                Me.ColLeft_DayOfWeekStart.HeaderText = "Start"
                Me.ColLeft_StartDate.HeaderText = ""
            Else
                Me.ColLeft_DayOfWeekStart.Visible = False
                Me.ColLeft_DayOfWeekStart.HeaderText = ""
                Me.ColLeft_StartDate.HeaderText = "Start"
            End If
            If ShowDayFinish = "1" Then
                Me.ColLeft_DayOfWeekFinish.Visible = True
                Me.ColLeft_DayOfWeekFinish.HeaderText = "End"
                Me.ColLeft_FinishDate.HeaderText = ""
            Else
                Me.ColLeft_DayOfWeekFinish.Visible = False
                Me.ColLeft_DayOfWeekFinish.HeaderText = ""
                Me.ColLeft_FinishDate.HeaderText = "End"
            End If
            If Me.cboGroupBy.SelectedIndex = enumGroup.BY_JO Then
                Me.ColLeft_WorkUnit.Visible = True
                Me.ColLeft_ProcessCode.Visible = False
            Else
                Me.ColLeft_WorkUnit.Visible = False
                If ShowProcess = "1" Then
                    Me.ColLeft_ProcessCode.Visible = True
                Else
                    Me.ColLeft_ProcessCode.Visible = False
                End If
                If JmlNodeExpand = 0 Then
                    Me.ColLeft_DayOfWeekFinish.Visible = False
                    Me.ColLeft_FinishDate.Visible = False
                    Me.ColLeft_FinishTime.Visible = False
                    Me.ColLeft_DayOfWeekStart.Visible = False
                    Me.ColLeft_StartDate.Visible = False
                    Me.ColLeft_StartTime.Visible = False
                Else
                    Me.ColLeft_FinishDate.Visible = True
                    Me.ColLeft_FinishTime.Visible = True
                    Me.ColLeft_StartDate.Visible = True
                    Me.ColLeft_StartTime.Visible = True
                End If
            End If
            'If ShowProgress_ = "1" Then
            '    Me.ColLeft_Progress.Visible = True
            'Else
            '    Me.ColLeft_Progress.Visible = False
            'End If
            If ShowRate = "1" Then
                Me.ColLeft_ProductionRate.Visible = True
            Else
                Me.ColLeft_ProductionRate.Visible = False
            End If
            If ShowQty = "1" Then
                Me.ColLeft_QtyProduce.Visible = True
            Else
                Me.ColLeft_QtyProduce.Visible = False
            End If
            If ShowUnit = "1" Then
                Me.ColLeft_Unit.Visible = True
            Else
                Me.ColLeft_Unit.Visible = False
            End If
            If ShowWH = "1" Then
                Me.ColLeft_TotalWH.Visible = True
            Else
                Me.ColLeft_TotalWH.Visible = False
            End If

            If ShowInputLocation = "1" Then
                Me.ColLeft_LocationInput.Visible = True
            Else
                Me.ColLeft_LocationInput.Visible = False
            End If
            If ShowOutputLocation = "1" Then
                Me.ColLeft_LocationOutput.Visible = True
            Else
                Me.ColLeft_LocationOutput.Visible = False
            End If

            If ShowActualGoodQty = "1" Then
                Me.ColLeft_ActualGood.Visible = True
            Else
                Me.ColLeft_ActualGood.Visible = False
            End If
            If ShowActualRejectQty = "1" Then
                Me.ColLeft_ActualReject.Visible = True
            Else
                Me.ColLeft_ActualReject.Visible = False
            End If
            If ShowActualProduceQty = "1" Then
                Me.ColLeft_ActualProduce.Visible = True
            Else
                Me.ColLeft_ActualProduce.Visible = False
            End If

            Me.SetSomeColumnVisibilty()
            Me.ResizeMainPanel()

            'Me.ColLeft_IdentifierWO.Visible = True
            'Me.ColLeft_IdentifierJO.Visible = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpFrom_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpFrom.Enter
        Try
            Me.Buffer_A = Me.dtpFrom.Value.ToString

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpFrom_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpFrom.Leave
        Try
            Me.BackgroundWorker3.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker3_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker3.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Dim ForceRefresh As Boolean = False
    Private Sub BackgroundWorker3_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker3.RunWorkerCompleted
        Dim Dummy As String = ""
        Dim StartDate, EndDate As Date

        Try
            If ForceRefresh Then
                ForceRefresh = False
                'Dummy = Me.ObjSetting1.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
                'Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_SchHor_EndDate)
                'If Dummy = "" Then
                '    Dummy = "30"
                'End If
                Dummy = Me.ScheduleHorizon
                EndDate = Me.dtpFrom.Value + TimeSpan.FromDays(Dummy)
                Me.dtpTo.Value = EndDate
                StartDate = Me.dtpFrom.Value
                Me.FullReload()
                'Me.LoadTblSch(StartDate, EndDate) 
                'Me.reloadGanttChart()
            Else
                If Me.Buffer_A <> Me.dtpFrom.Value.ToString Then
                    'hitung end date
                    'Dummy = Me.ObjSetting1.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
                    'Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_SchHor_EndDate)
                    'If Dummy = "" Then
                    '    Dummy = "30"
                    'End If
                    Dummy = Me.ScheduleHorizon
                    EndDate = Me.dtpFrom.Value + TimeSpan.FromDays(Dummy)
                    Me.dtpTo.Value = EndDate

                    'Dummy = Me.ObjSetting1.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
                    'Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_SchHor_StartDate)
                    'If Dummy = "" Then
                    '    Dummy = "-3"
                    'End If
                    StartDate = Me.dtpFrom.Value

                    'load sch bar
                    'Me.LoadTblSch(StartDate, EndDate)
                    Me.FullReload()

                    'ubah nilai dtp
                    Me.refreshBar = True
                    Me.dtpBarFrom.Value = StartDate

                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
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

            Me.tblSchBar.Rows.Clear()
            'DTAWU1.FillBy_AdaSch(TBLWU1, Me.dtpTo.Value.Date + TimeSpan.FromDays(8), Me.dtpFrom.Value.Date - TimeSpan.FromDays(7))
            'Me.ObjSch.LoadScheduleBar(vStartDate, vEndDate, Me.tblSchBar, TBLWU1)
            StartTime = Now
            Me.ObjSch.LoadScheduleBar_WOItemOnly(vStartDate, vEndDate, Me.tblSchBar)
            'MsgBox("Reload Time: Item Sch - " + Me.objCD.DigitGrouping((Now - StartTime).TotalSeconds, 0) + " seconds", MsgBoxStyle.Information, "Information")

            If Me.Reload_TBLTanpaItemWO = True Then
                StartTime = Now
                Me.tblSchBar_TanpaItemWO.Rows.Clear()
                Me.ObjSch.LoadScheduleBar_TanpaScheduledItem(vStartDate, vEndDate, Me.tblSchBar_TanpaItemWO)
                'MsgBox("Reload Time: Sch Tanpa Item WO - " + Me.objCD.DigitGrouping((Now - StartTime).TotalSeconds, 0) + " seconds", MsgBoxStyle.Information, "Information")
            End If

            Dim ij As Integer = 0
            While ij <= tblSchBar.Rows.Count - 1
                rowSch = tblSchBar.Rows(ij)
                If rowSch.tipe <> ClsTrans_Schedule.EnumTipeSchBar.work_time Then
                    Me.tblSchBar.Rows.Remove(rowSch)
                Else
                    ij += 1
                End If
            End While

            'isikan nilai yang dibutuhkan
            If Me.tblSchBar.Rows.Count > 0 Then
                StartTime = Now

                MOMin = Double.MaxValue
                MOMax = Double.MinValue
                For Each rowSch In Me.tblSchBar.Rows
                    If rowSch.primary_mo <> -1 Then
                        If rowSch.primary_mo < MOMin Then
                            MOMin = rowSch.primary_mo
                        End If
                        If rowSch.primary_mo > MOMax Then
                            MOMax = rowSch.primary_mo
                        End If
                    End If
                Next

                DTA.Fill(TBL, MOMax, MOMin, CDate(Format(Me.dtpFrom.Value, "dd MMM yyyy") + " 00:00:00"))

                While mn <= Me.tblSchBar.Rows.Count - 1
                    Ketemu = False

                    rowSch = Me.tblSchBar.Rows(mn)
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
                                    While jk <= Me.tblSchBar.Rows.Count - 1
                                        rowSch1 = Me.tblSchBar.Rows(jk)
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
                While mn <= Me.tblSchBar.Rows.Count - 1
                    Ketemu = False
                    rowSch = Me.tblSchBar.Rows(mn)
                    If rowSch.identifier_wo <> "" Then
                        If rowSch.work_unit_code = "" Then
                            For Each rowWU In TBLWU.Rows
                                If rowWU.primary = rowSch.primary_work_unit And rowSch.work_unit_code = "" Then
                                    rowSch.work_unit_code = rowWU.work_unit_code

                                    jk = mn + 1
                                    While jk <= Me.tblSchBar.Rows.Count - 1
                                        rowSch1 = Me.tblSchBar.Rows(jk)
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

                'MsgBox("Reload Time: Set WU + MO/JO No - " + Me.objCD.DigitGrouping((Now - StartTime).TotalSeconds, 0) + " seconds", MsgBoxStyle.Information, "Information")
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            DTAWU.Dispose()
            TBLWU.Dispose()
            DTAWU1.Dispose()
            TBLWU1.Dispose()
            'MsgBox("Load sch bar: " + Me.objCD.DigitGrouping((Now - StartTime).TotalSeconds, 0) + " seconds")
        End Try
    End Sub

    Private Sub btnAddSchedule_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddSchedule.Click
        Try
            If Me.ObjMOItem.ApakahAdaYg_Ready_BelumTerjadwal = True Then
                Me.AddSch()
            Else
                MsgBox("No new manufacturing order ready for scheduling.", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub tms_ChangeColor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tms_ChangeColor.Click
        Dim IdentifierJO As String = ""
        'Dim rowDGV As DataGridView
        Dim ij As New ColorDialog

        Try
            If Me.ShowGanttChart = False Then
                Exit Sub
            End If
            IdentifierJO = Me.tgvSch.Item(Me.ColLeft_IdentifierJO.Index, Me.tgvSch.SelectedCells(0).RowIndex).Value
            If IdentifierJO <> "" Then
                If ij.ShowDialog = Windows.Forms.DialogResult.OK Then
                    'simpan ke database
                    Me.ObjMO.Set_argb(IdentifierJO, ij.Color.A, ij.Color.R, ij.Color.G, ij.Color.B)

                    ''update ke tbl sch
                    'For Each Me.rowSch In Me.tblSchBar.Rows
                    '    If rowSch.identifier_jo = IdentifierJO Then
                    '        rowSch.a = ij.Color.A
                    '        rowSch.r = ij.Color.R
                    '        rowSch.g = ij.Color.G
                    '        rowSch.b = ij.Color.B
                    '    End If
                    'Next

                    ''ubah tampilan
                    'Me.reloadGanttChart()

                    Me.tblData.DefaultView.RowFilter = "[identifier jo] = '" & IdentifierJO & "'"
                    If Me.tblData.DefaultView.Count > 0 Then
                        For i As Integer = 0 To Me.tblData.DefaultView.Count - 1
                            Me.rowData = Me.tblData.DefaultView(i).Row
                            Me.rowData.a = ij.Color.A
                            Me.rowData.r = ij.Color.R
                            Me.rowData.g = ij.Color.G
                            Me.rowData.b = ij.Color.B
                            Dim drawBar As Boolean = False
                            If Me.rowData.Item("Is Header") = 1 Then
                                If Me.dgvBar.Rows(Me.tblData.Rows.IndexOf(Me.rowData) + 1).Visible = False Then
                                    drawBar = True
                                End If
                            End If
                            Me.reDrawBarHeader(Me.rowData, drawBar)
                        Next
                    End If
                    Me.tblSchBar.DefaultView.RowFilter = "[identifier jo] = '" & IdentifierJO & "'"
                    If tblSchBar.DefaultView.Count > 0 Then
                        For i2 As Integer = 0 To Me.tblSchBar.DefaultView.Count - 1
                            rowSch = tblSchBar.DefaultView(i2).Row
                            rowSch.a = ij.Color.A
                            rowSch.r = ij.Color.R
                            rowSch.g = ij.Color.G
                            rowSch.b = ij.Color.B
                        Next
                    End If

                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnAddOT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddOT.Click
        Dim ij As New FrmTrans_work_unitionSchedule_EditOT

        Try
            ij.UserCode = Me.UserCode
            ij.BackColorAll = Me.BackColorAll
            ij.ShowDialog()

            If ij.IsEdited Then
                Me.refreshBar = False
                'Me.LoadTblSch(Me.dtpFrom.Value, Me.dtpTo.Value)
                'Me.reloadGanttChart()
                Me.FullReload()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpFirmPlanned_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpFirmPlanned.Enter
        Try
            Me.Buffer_B = Me.dtpFirmPlanned.Value.Date.ToString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpFirmPlanned_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpFirmPlanned.Leave
        Try
            Me.BackgroundWorker4.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker4_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker4.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker4_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker4.RunWorkerCompleted
        Try
            If Me.Buffer_B <> Me.dtpFirmPlanned.Value.Date.ToString Then
                Me.ObjSch.SetFirmPlan(Me.dtpFirmPlanned.Value.Date, Today)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Try
            Me.Timer1.Enabled = False
            If Me.objMonitor.getUserControl(Me.userIdentifier, Me.userType) = Core_Module.ClsTransaksiMonitor_Production.enumUserControl.FULL_CONTROL Then
                If Me.ObjMOItem.ApakahAdaYg_Ready_BelumTerjadwal = True And Me.IsFullControl = True Then
                    If MsgBox("There are manufacturing order ready for scheduling, add new schedule?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.AddSch()
                    End If
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnWorkshift_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnWorkshift.Click
        Dim i As New Core_Module.FrmMasterWorkShift
        Try
            i.UserCode = Me.UserCode
            i.BackColorAll = Me.BackColorAll
            i.TipeWorkshift = Core_Module.ClsMasterWorkShift.EnumTipeWorkshift.Production
            i.ShowDialog(Me)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i = Nothing
        End Try
    End Sub

    Private Sub btnAllocateWorkshift_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAllocateWorkshift.Click
        Dim i As New Inventory_Module.FrmTransMO_WorkShift_Allocate
        Try
            i.UserCode = Me.UserCode
            i.BackColorAll = Me.BackColorAll
            i.ShowDialog(Me)
            If i.IsEdited Then
                Me.refreshBar = False
                'Me.LoadTblSch(Me.dtpFrom.Value, Me.dtpTo.Value)
                'Me.reloadGanttChart()
                Me.FullReload()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i = Nothing
        End Try
    End Sub

    Private Sub btnPrintSchedule_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPrintSchedule.Click
        Try
            Dim ij As Inventory_Module.FrmLaporanProd_PlanningReport
            Dim j As Windows.Forms.Form

            For Each j In Me.MdiChildren
                ij = TryCast(j, Inventory_Module.FrmLaporanProd_PlanningReport)
                If IsNothing(ij) = False Then
                    ij.BringToFront()
                    ij.Focus()
                    ij.dtpTo.Value = Me.dtpTo.Value
                    ij.dtpWOStart.Value = Me.DtpFromDate.Value
                    ij.KosongkanReport()
                    Exit Sub
                End If
            Next

            Dim i As New Inventory_Module.FrmLaporanProd_PlanningReport
            ' Increment the caption counter.
            i.UserCode = Me.UserCode
            ' Set the caption to be unique.
            i.BackColorAll = Me.BackColorAll
            i.MdiParent = Me.ParentForm
            'i.ShowIcon = False
            i.dtpTo.Value = Me.dtpTo.Value
            i.dtpWOStart.Value = Me.dtpFrom.Value
            i.KosongkanReport()

            i.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FullReload_ScheduleOnly()
        Try
            Reload_TBLTanpaItemWO = False '-> tabel tanpa item wo ga usah dihitung
            Control.CheckForIllegalCrossThreadCalls = False
            Me.GroupBox3.Visible = True
            Me.NonAktifkanDetail()
            Me.bgwFullReload.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FullReload()
        Try
            Me.Reload_TBLTanpaItemWO = True
            Control.CheckForIllegalCrossThreadCalls = False
            Me.GroupBox3.Visible = True
            Me.NonAktifkanDetail()
            Me.bgwFullReload.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AktifkanDetail()
        Try
            Me.ControlBox = True
            Me.BtnAdd.Enabled = True
            Me.btnAddDelay.Enabled = True
            Me.btnAddDownTime.Enabled = True
            Me.btnAddOT.Enabled = True
            Me.btnAddRemoveHoliday.Enabled = True
            Me.btnAddSchedule.Enabled = True
            Me.btnAddSetupTime.Enabled = True
            Me.btnAllocateWorkshift.Enabled = True
            Me.BtnClear.Enabled = True
            Me.BtnClearGanttchart.Enabled = True
            Me.btnException.Enabled = True
            Me.BtnLoad.Enabled = True
            Me.btnPrintPickList.Enabled = True
            Me.btnPrintSchedule.Enabled = True
            Me.btnRefresh.Enabled = True
            Me.BtnRemove.Enabled = True
            Me.BtnSave.Enabled = True
            Me.BtnSet2.Enabled = True
            Me.btnWorkshift.Enabled = True
            Me.llbCollapsAll.Enabled = True
            Me.dtpFrom.Enabled = True
            Me.dtpFirmPlanned.Enabled = True
            Me.chkShowProgress.Enabled = True
            Me.llbColumnSettingSch.Enabled = True
            Me.cboGroupBy.Enabled = True
            Me.tgvSch.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            Me.ControlBox = False
            Me.BtnAdd.Enabled = False
            Me.btnAddDelay.Enabled = False
            Me.btnAddDownTime.Enabled = False
            Me.btnAddOT.Enabled = False
            Me.btnAddRemoveHoliday.Enabled = False
            Me.btnAddSchedule.Enabled = False
            Me.btnAddSetupTime.Enabled = False
            Me.btnAllocateWorkshift.Enabled = False
            Me.BtnClear.Enabled = False
            Me.BtnClearGanttchart.Enabled = False
            Me.btnException.Enabled = False
            Me.BtnLoad.Enabled = False
            Me.btnPrintPickList.Enabled = False
            Me.btnPrintSchedule.Enabled = False
            Me.btnRefresh.Enabled = False
            Me.BtnRemove.Enabled = False
            Me.BtnSave.Enabled = False
            Me.BtnSet2.Enabled = False
            Me.btnWorkshift.Enabled = False
            Me.llbCollapsAll.Enabled = False
            Me.dtpFrom.Enabled = False
            Me.dtpFirmPlanned.Enabled = False
            Me.chkShowProgress.Enabled = False
            Me.llbColumnSettingSch.Enabled = False
            Me.cboGroupBy.Enabled = False
            Me.tgvSch.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwFullReload_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgwFullReload.DoWork
        Try
            Me.objCD.UpdateThreading()
            Me.LblProgress.Text = "Loading from database"
            Me.LoadTblSch(Me.dtpFrom.Value, Me.dtpTo.Value)
            Me.LblProgress.Text = "Showing data"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwFullReload_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgwFullReload.RunWorkerCompleted
        Try
            Me.reloadGanttChart()
            Me.AktifkanDetail()
            Me.GroupBox3.Visible = False
            Me.refreshBar = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadDataAwal1()
        Dim Dummy As String = ""
        Dim EndDate As Date = Today
        Dim StartTime As Date = Now
        Try
            Me.dtpFirmPlanned.Value = Me.ObjSch.GetFirmPlan
            Dummy = Me.objSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_StartDateDiff, Me.UserCode)
            If Dummy = "" Then
                Dummy = "0"
            End If
            Me.dtpFrom.Value = Me.dtpFirmPlanned.Value.Date + TimeSpan.FromDays(Dummy)
            'While Me.dtpFrom.Value.Date.DayOfWeek <> DayOfWeek.Monday
            '    Me.dtpFrom.Value = (Me.dtpFrom.Value - TimeSpan.FromDays(1)).Date
            'End While
            Me.dtpBarFrom.Value = Me.dtpFrom.Value

            Dummy = Me.ScheduleHorizon ' tadinya 12 * 7
            EndDate = Me.dtpFrom.Value.Date + TimeSpan.FromDays(Dummy)
            Me.dtpTo.Value = EndDate

            Control.CheckForIllegalCrossThreadCalls = False
            Me.GroupBox3.Visible = True
            Me.NonAktifkanDetail()
            Me.bgwFormLoad.RunWorkerAsync()

            'MsgBox("Total duration: " + Me.objCD.DigitGrouping((Now - StartTime).TotalMinutes, 1))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwFormLoad_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgwFormLoad.DoWork
        Try
            Me.objCD.UpdateThreading()
            Me.LblProgress.Text = "Loading from database"
            Me.LoadTblSch(Me.dtpFrom.Value, Me.dtpTo.Value)
            Me.LblProgress.Text = "Showing data"
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgwFormLoad_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgwFormLoad.RunWorkerCompleted
        Dim ObjRdn As New Random

        Try
            Me.reloadGanttChart()
            Me.AktifkanDetail()
            Me.GroupBox3.Visible = False

            Me.Buffer = Me.dtpFrom.Value.ToString
            Me.ExpandAll = False
            Me.Timer1.Enabled = True

            If Me.objUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_ProductionSchedule_SendTakeControl") = True Then
                Me.userType = Core_Module.ClsTransaksiMonitor_Production.enumUserType.READ_OR_FULLCONTROL
            Else
                Me.userType = Core_Module.ClsTransaksiMonitor_Production.enumUserType.READ_ONLY
            End If
            Me.objMonitor = New Core_Module.ClsTransaksiMonitor_Production(Me.UserCode, Me.userIdentifier, Me.userType)
            Me.ToolStripStatusLabel1.IsLink = True
            Me.ToolStripStatusLabel1.ActiveLinkColor = Color.Blue
            Me.ToolStripStatusLabel1.VisitedLinkColor = Color.Blue
            Me.ToolStripStatusLabel1.LinkColor = Color.Blue
            Me.ToolStripStatusLabel1.LinkBehavior = LinkBehavior.NeverUnderline
            Me.userIdentifier = ObjRdn.Next(0, Integer.MaxValue)
            While Me.objMonitor.isExistIdentifier(Me.userIdentifier) = True
                Me.userIdentifier = ObjRdn.Next(0, Integer.MaxValue)
            End While

            Me.objMonitor.Set_UserIdentifier(Me.userIdentifier)
            Me.objMonitor.startProductionMonitor(3 * 60000, 5 * 60000)
            Dim ListOfHandShakeButton As New List(Of Button)
            Dim ListOfHandShakeDateTimePicker As New List(Of DateTimePicker)
            ListOfHandShakeButton.Add(Me.BtnAdd)
            ListOfHandShakeButton.Add(Me.btnAddDelay)
            ListOfHandShakeButton.Add(Me.btnAddDownTime)
            ListOfHandShakeButton.Add(Me.btnAddOT)
            ListOfHandShakeButton.Add(Me.btnAddSchedule)
            ListOfHandShakeButton.Add(Me.btnAddSetupTime)
            ListOfHandShakeButton.Add(Me.btnAllocateWorkshift)
            ListOfHandShakeButton.Add(Me.btnWorkshift)
            ListOfHandShakeDateTimePicker.Add(Me.dtpFirmPlanned)
            Me.objMonitor.startUserControlMonitor(Me.userControl, Me.ToolStripStatusLabel1, 10000, ListOfHandShakeButton, ListOfHandShakeDateTimePicker)
            Me.objMonitor.startNotifMonitor(10000)

            Me.dgvBar.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjRdn = Nothing
        End Try
    End Sub
End Class
