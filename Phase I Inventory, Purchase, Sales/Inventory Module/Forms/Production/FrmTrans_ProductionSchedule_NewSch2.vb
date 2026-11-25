Option Explicit On
Option Strict Off

Public Class FrmTrans_ProductionSchedule_NewSch2
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public IdentifierJO As String = ""
    Public CurrentJO, TotalJO As Double
    Public Stopped As Boolean = False
    Public vRescheduleStartTime As Date = #1/1/1950#

    Dim ObjLocation As New Core_Module.clsMasterStorageArea
    Dim ObjSch As New ClsTrans_Schedule
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjWU As New Core_Module.ClsMasterWorkUnit
    Dim TBLSchBar As New DseTrans_Schedule.tbl_sch_barDataTable
    Dim TBLSchBarKosong As New DseTrans_Schedule.tbl_sch_barDataTable
    Dim TBLSchBarTanpaWO As New DseTrans_Schedule.tbl_sch_barDataTable
    Dim TBLItemWO As New DseTrans_Schedule.tbltrans_mo_item_wo1DataTable
    Dim TBLItemWO_ItemSch As New DseTrans_Schedule.tbltrans_mo_item_wo_item_schDataTable
    Dim Buffer As String = ""
    Dim BufferTime As Date = Today
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim LangsungTutup As Boolean = False

    'data terkait modifikasi jumlah wu yang di-alokasikan
    Dim r1, r2, c1, c2 As Double
    Dim PrevIdx As Integer = -1
    Dim IdentifierAcuan As String = ""
    Dim ListPrimaryWU As New List(Of String)
    Dim TBLWUAcuan As New DseTrans_Schedule.TblWU1DataTable

    Dim callEndOfScroll As Boolean = True
    Dim rowSch As DseTrans_Schedule.tbl_sch_barRow
    Dim TBL_D As New DseTrans_Schedule.tbltrans_mo_item_woDataTable
    Dim tblBarIndex As New DseTrans_Schedule.BarIndexDataTable
    Dim rowBarIndex As DseTrans_Schedule.BarIndexRow
    Dim intervalPerBar_ByHour As TimeSpan = TimeSpan.FromMinutes(5)
    Dim intervalPerBar_ByDay As TimeSpan = TimeSpan.FromHours(1)
    Dim intervalPerBar_ByWeek As TimeSpan = TimeSpan.FromHours(12)
    Dim RefreshBar As Boolean = True
    Dim AlternateBarColor As Color = Color.WhiteSmoke
    Dim ObjSetting As New Core_Module.ClsTransaksiSettingUser
    Dim ObjSetting1 As New Core_Module.ClsTransaksiSetting
    Dim toolTipText As String = ""
    Dim frmTooltip As FrmTrans_ProductionSchedule_NewSch2_Tooltip

    Dim a, r, g, b As Integer
    Dim tblProcess As New DseTrans_ProductionSchedule_NewSch3.DataTable5DataTable
    Dim rowProcess As DseTrans_ProductionSchedule_NewSch3.DataTable5Row
    Dim tblWorkUnit As New DseTrans_ProductionSchedule_NewSch3.DataTable6DataTable
    Dim rowWorkUnit As DseTrans_ProductionSchedule_NewSch3.DataTable6Row
    Dim tblHeader As New DseTrans_ProductionSchedule_NewSch3.DataTable1DataTable
    Dim rowHeader As DseTrans_ProductionSchedule_NewSch3.DataTable1Row
    Dim PrimaryPredefinedConfig As Double = -1
    Dim RefrehGanttChart As Boolean = True

    Private Sub SetPlanedFinishDate()
        Dim MaxEndDate As Date = #1/1/1950#

        Try
            For Each rowWU As DseTrans_Schedule.tbltrans_mo_item_wo_item_schRow In Me.TBLItemWO_ItemSch.Rows
                If rowWU.plan_end_time > MaxEndDate Then
                    MaxEndDate = rowWU.plan_end_time
                End If
            Next

            If MaxEndDate = #1/1/1950# Then
                Me.dtpPlannedFinishDate.Visible = False
            Else
                Me.dtpPlannedFinishDate.Visible = True
                Me.dtpPlannedFinishDate.Value = MaxEndDate
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Function ShowSchedulePlanCalculationMethod(ByRef vPrimaryPredefineConfig As Double) As Boolean
        Dim f As New FrmTrans_ProductionSchedule_NewSch3
        Dim res As DialogResult
        Try
            f.UserCode = Me.UserCode
            f.BackColorAll = Me.BackColorAll
            f.IdentifierJO = Me.IdentifierJO
            res = f.ShowDialog
            If res <> Windows.Forms.DialogResult.OK Then
                Me.Stopped = f.Stopped
                Return False
            End If
            vPrimaryPredefineConfig = f.PrimaryConfig
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            f.Dispose()
        End Try
    End Function

    Private Sub FrmTrans_ProductionSchedule_NewSch2_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.LangsungTutup = False Then
                If MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirmation") <> MsgBoxResult.Yes Then
                    e.Cancel = True
                End If
            Else

            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_ProductionSchedule_NewSch2_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim DefColor As Color = Me.ObjSch.GetDefaultColor_AllocatedWS
        Dim StartTime As Date = Now
        Dim a, r, g, b As String

        Try
            Me.tgvSch.ReadOnly = False
            Me.ColLeft_InputLocation.ReadOnly = False
            Me.ColLeft_OutputLocation.ReadOnly = False

            Me.dgvHeader1.ColumnHeadersHeight = Me.tgvSch.Location.Y
            Me.dgvHeader0.Location = New Point(Me.dgvHeader0.Location.X, Me.tgvSch.Location.Y)
            Me.dgvBar.Location = New Point(Me.dgvBar.Location.X, Me.tgvSch.Location.Y + Me.tgvSch.ColumnHeadersHeight)
            Me.dgvBar.Height = Me.tgvSch.Height - Me.tgvSch.ColumnHeadersHeight

            If Me.ShowSchedulePlanCalculationMethod(PrimaryPredefinedConfig) = False Then
                Me.LangsungTutup = True
                Me.Close()
                Exit Sub
            End If
            StartTime = Now
            If PrimaryPredefinedConfig <> -1 Then
                Me.CoAll_WU_Allocate.Visible = False
                Me.ColAll_Process_TotalWU.ReadOnly = True
            End If

            Me.a = DefColor.A
            Me.r = DefColor.R
            Me.g = DefColor.G
            Me.b = DefColor.B
            pb_Sch.BackColor = DefColor

            Me.cboTimeframe.SelectedIndex = 2
            Me.dgvException.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            'Me.dgvSch.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvWUAllocate_WU.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvWUAllocate_Process.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            'load data detail jo
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.LoadDataDetailJO()

            'set schedule
            Me.SetScheduleAwal()
            Me.cboOptimationPlan.Focus()

            Me.SetPlanedFinishDate()
            Me.ColAll_Process_Rate.HeaderText = "Cummulative Rate (Unit/Hour)"


            'Me.dgvHeader0.ColumnHeadersHeight = Me.dgvSch.ColumnHeadersHeight
            Me.dgvBar.Location = New Point(Me.dgvBar.Location.X, Me.dgvHeader0.Location.Y + Me.dgvHeader0.ColumnHeadersHeight)

            Me.dgvBar.DefaultCellStyle.SelectionBackColor = Me.AlternateBarColor
            Me.setColumnVisibility()
            Me.LlbColumnSettingSch.Visible = False
            'Me.dgvAddListener(Me.dgvBar)

            Me.Text = Me.Text + " - Data: " + Me.CurrentJO.ToString + " of " + Me.TotalJO.ToString + ""
            a = ""
            a = Me.ObjSetting1.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_AllocatedWS_a)
            r = Me.ObjSetting1.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_AllocatedWS_r)
            g = Me.ObjSetting1.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_AllocatedWS_g)
            b = Me.ObjSetting1.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_AllocatedWS_b)
            If a = "" Then
                Me.SetColor(Core_Module.GetSMESSBackColorAll)
            Else
                Me.SetColor(Color.FromArgb(a, r, g, b))
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadDataDetailJO()
        Dim DTA As New DseTrans_ScheduleTableAdapters.tbltrans_mo_item1TableAdapter
        Dim TBL As New DseTrans_Schedule.tbltrans_mo_item1DataTable
        Dim row As DseTrans_Schedule.tbltrans_mo_item1Row

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, Me.IdentifierJO)
            row = TBL.Rows(0)
            With row
                Me.TxtCodeNo.Text = .Product_Code
                Me.TxtDesc.Text = .Description
                Me.TxtJoNumber.Text = .JO_Number
                Me.TxtMONumber.Text = .MO_Number
                Me.TxtPrimaryMO.Text = .Primary_MO
                Me.TxtQtyProduce.Text = Me.ObjCD.DigitGrouping(.Produce_Qty, ._Decimal)
                Me.TxtQtyTarget.Text = Me.ObjCD.DigitGrouping(.Target_Qty, ._Decimal)
                Me.dtpJODate.Text = ._Date
                If .Target_Finish_Date_R2 <> #1/1/1950# Then
                    Me.dtpTargetFinishDate.Value = .Target_Finish_Date_R2
                    Me.LblTargetFinishDate.Text = "Target Finish Date (R2)"
                ElseIf .Target_Finish_Date_R1 <> #1/1/1950# Then
                    Me.dtpTargetFinishDate.Value = .Target_Finish_Date_R1
                    Me.LblTargetFinishDate.Text = "Target Finish Date (R1)"
                Else
                    Me.dtpTargetFinishDate.Value = .Target_Finish_Date
                End If
                Me.pb_Sch.BackColor = Color.FromArgb(.a, .r, .g, .b)
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Private Sub SetScheduleAwal()
        Dim TBLWU As New DseTrans_Schedule.tblmaster_product_item_workunit1DataTable
        Dim DTAWU As New DseTrans_ScheduleTableAdapters.tblmaster_product_item_workunit1TableAdapter
        Dim rowWU As DseTrans_Schedule.tblmaster_product_item_workunit1Row
        Dim DTAWU2 As New DseTrans_ScheduleTableAdapters.tbltrans_mo_master_predefine_workunit_itemworkunit1TableAdapter
        Dim TBLWU2 As New DseTrans_Schedule.tbltrans_mo_master_predefine_workunit_itemworkunit1DataTable
        Dim rowWU2 As DseTrans_Schedule.tbltrans_mo_master_predefine_workunit_itemworkunit1Row

        Try
            'hitung schedule
            DTAWU.Connection.ConnectionString = Core_Module.GetConstring
            DTAWU2.Connection.ConnectionString = DTAWU.Connection.ConnectionString

            If Me.PrimaryPredefinedConfig = -1 Then
                DTAWU.Fill(TBLWU, Me.IdentifierJO)
            Else
                DTAWU2.Fill(TBLWU2, Me.PrimaryPredefinedConfig)
                TBLWU.Rows.Clear()
                For Each rowWU2 In TBLWU2.Rows
                    rowWU = TBLWU.Newtblmaster_product_item_workunit1Row
                    rowWU.primary_work_unit = rowWU2.primary_work_unit
                    TBLWU.Rows.Add(rowWU)
                Next
            End If


            Me.ObjSch.HitungJadwalBaru(Me.TBLSchBar, Me.IdentifierJO, Me.TBLItemWO, Me.TBLItemWO_ItemSch, _
            Me.dtpTargetFinishDate.Value, Me.PrimaryPredefinedConfig, TBLWU, Me.TBLSchBarKosong, Me.TBLSchBarTanpaWO)

            'tampilkan jadwal ke dalam dgv
            Me.LoadDGVProses()

            'isikan exception

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TBLWU.Dispose()
            DTAWU.Dispose()
            DTAWU2.Dispose()
            TBLWU.dispose()
        End Try
    End Sub

    Private Sub LoadDGVProses(Optional ByVal r1 As Integer = -1, Optional ByVal c1 As Integer = -1)
        Dim rowItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1Row
        Dim rowDGV As DataGridViewRow

        Try
            Me.DgvWUAllocate_Process.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DgvWUAllocate_Process.Rows.Clear()

            'isikan data dari tbl item wo
            For Each rowItemWO In Me.TBLItemWO.Rows
                Me.DgvWUAllocate_Process.Rows.Add()
                rowDGV = Me.DgvWUAllocate_Process.Rows(Me.DgvWUAllocate_Process.Rows.Count - 1)
                With rowDGV
                    .Cells(Me.ColAll_Process_Code.Index).Value = rowItemWO.Process_Code
                    .Cells(Me.ColAll_Process_CodeNo.Index).Value = rowItemWO.Product_Code
                    .Cells(Me.ColAll_Process_Desc.Index).Value = rowItemWO.Product_Desc
                    .Cells(Me.ColAll_Process_No.Index).Value = rowDGV.Index + 1
                    .Cells(Me.ColAll_Process_PrimaryCodeNo.Index).Value = rowItemWO.Primary_Product
                    .Cells(Me.ColAll_Process_PrimaryProcess.Index).Value = rowItemWO.Primary_Process
                    .Cells(Me.ColAll_Process_Qty.Index).Value = Me.ObjCD.DigitGrouping(rowItemWO.Qty_Target_Prod + rowItemWO.Qty_Target_Prod_Add, rowItemWO._Decimal)
                    .Cells(Me.ColAll_Process_Rate.Index).Value = Me.ObjCD.DigitGrouping(rowItemWO.Plan_Prod_Rate, rowItemWO._Decimal)
                    .Cells(Me.ColAll_UnitCode.Index).Value = rowItemWO.Unit_Code
                    .Cells(Me.ColAll_Process_IdentifierWO.Index).Value = rowItemWO.Identifier
                    .Cells(Me.ColAll_Process_TotalWorkHour.Index).Value = Me.ObjCD.Convert_ToString(rowItemWO.Plan_Work_Hour)
                    .Cells(Me.ColAll_Process_TotalWU.Index).Value = rowItemWO.WU_Allocated
                    .Cells(Me.ColAll_Process_Decimal.Index).Value = rowItemWO._Decimal
                    .Cells(Me.ColAll_Process_StartDate.Index).Value = Format(rowItemWO.Plan_Start_Time, "dd MMM yyyy")
                    .Cells(Me.ColAll_Process_StartTime.Index).Value = Format(rowItemWO.Plan_Start_Time, "HH.mm")
                    .Cells(Me.ColAll_Process_EndDate.Index).Value = Format(rowItemWO.Plan_End_Time, "dd MMM yyyy")
                    .Cells(Me.ColAll_Process_EndTime.Index).Value = Format(rowItemWO.Plan_End_Time, "HH.mm")
                    .Cells(Me.ColAll_Process_TipeDependency.Index).Value = Me.ObjProduct.GetTipeDependencyStr(rowItemWO.tipe_dependency)
                    .Cells(Me.ColAll_Process_DelayDependency.Index).Value = Me.ObjCD.Convert_ToString(rowItemWO.delay_dependency)
                End With
            Next

            Me.DgvWUAllocate_Process.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.DgvWUAllocate_Process.ClearSelection()
            If Me.DgvWUAllocate_Process.Rows.Count > 0 Then
                If r1 = -1 Then
                    Me.DgvWUAllocate_Process.Rows(0).Selected = True
                Else
                    Me.DgvWUAllocate_Process.Item(c1, r1).Selected = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'mengembalikan end time max, atas item wo sch yang identifier parentnya diberikan
    Private Function GetMaxEndTime(ByVal vIdentifierParent As String) As Date
        Dim RetVal As Date = #1/1/1950#
        Dim RowSchBar1 As DseTrans_Schedule.tbl_sch_barRow
        Dim rowItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1Row

        Try
            Me.TBLItemWO.DefaultView.RowFilter = "[Identifier WO Parent] = '" + vIdentifierParent + "'"
            For Each DV As Data.DataRowView In Me.TBLItemWO.DefaultView
                rowItemWO = DV.Row

                Me.TBLSchBar.DefaultView.RowFilter = "[identifier wo] = '" + rowItemWO.Identifier + "' AND tipe = " + CInt(ClsTrans_Schedule.EnumTipeSchBar.work_time).ToString
                Me.TBLSchBar.DefaultView.Sort = "[plan end time] DESC"
                If Me.TBLSchBar.DefaultView.Count > 0 Then
                    RowSchBar1 = Me.TBLSchBar.DefaultView(0).Row
                    If RetVal < RowSchBar1.plan_end_time Then
                        RetVal = RowSchBar1.plan_end_time
                    End If
                End If
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetStartTime(ByVal vPrimaryWU As Double) As Date
        Dim RetVal As Date = #1/1/1950#
        Dim RowSchBar1 As DseTrans_Schedule.tbl_sch_barRow

        Try
            Me.TBLSchBar.DefaultView.RowFilter = "[primary work unit] = " + vPrimaryWU.ToString + _
            " AND [identifier wo] <> '' AND tipe = " + CInt(ClsTrans_Schedule.EnumTipeSchBar.work_time).ToString
            Me.TBLSchBar.DefaultView.Sort = "[plan end time] DESC"
            If Me.TBLSchBar.DefaultView.Count > 0 Then
                RowSchBar1 = Me.TBLSchBar.DefaultView(0).Row
                RetVal = RowSchBar1.plan_end_time
            End If

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub LoadDGVWU()
        Dim rowDGV, rowDGV1 As DataGridViewRow
        Dim DTAWU As New DseTrans_ScheduleTableAdapters.tblmaster_product_item_workunitTableAdapter
        Dim TBLWU As New DseTrans_Schedule.tblmaster_product_item_workunitDataTable
        Dim rowWU As DseTrans_Schedule.tblmaster_product_item_workunitRow
        Dim rowItemWO_ItemSch As DseTrans_Schedule.tbltrans_mo_item_wo_item_schRow
        Dim Ketemu As Boolean = False
        Dim DummyStart As Date = #1/1/1950#
        Dim FirmPlanDate As Date = Today
        Dim StartDateMin As Date = Today
        Dim PrevIdx As Integer = 0

        Try
            DTAWU.Connection.ConnectionString = Core_Module.GetConstring
            Me.DgvWUAllocate_WU.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DgvWUAllocate_WU.Rows.Clear()

            If Me.DgvWUAllocate_Process.SelectedCells.Count = 0 Then
                Me.DgvWUAllocate_WU.Rows.Clear()
            Else
                rowDGV = Me.DgvWUAllocate_Process.Rows(Me.DgvWUAllocate_Process.SelectedCells(0).RowIndex)
                If IsNothing(rowDGV.Cells(Me.ColAll_Process_IdentifierWO.Index).Value) = True Then
                    Exit Sub
                End If
                'isikan data semua mesin yang mungkin, urutkan berdasarkan code no
                TBLWU.Rows.Clear()
                DTAWU.FillBy_PrimaryProduct(TBLWU, rowDGV.Cells(Me.ColAll_Process_PrimaryCodeNo.Index).Value)
                For Each rowWU In TBLWU.Rows
                    rowWU.effective_rate = Math.Round((rowWU.efficiency * rowWU.rate_per_hour) / 100, rowWU._Decimal, MidpointRounding.AwayFromZero)
                    Me.DgvWUAllocate_WU.Rows.Add()
                    rowDGV1 = Me.DgvWUAllocate_WU.Rows(Me.DgvWUAllocate_WU.Rows.Count - 1)
                    With rowDGV1
                        .Cells(Me.CoAll_WU_Allocate.Index).Value = Boolean.FalseString
                        .Cells(Me.CoAll_WU_Code.Index).Value = rowWU.work_unit_code
                        .Cells(Me.CoAll_WU_Desc.Index).Value = rowWU.description
                        .Cells(Me.CoAll_WU_Efficiency.Index).Value = Me.ObjCD.DigitGrouping(rowWU.efficiency, 2)
                        .Cells(Me.CoAll_WU_EndDate.Index).Value = ""
                        .Cells(Me.CoAll_WU_EndTime.Index).Value = ""
                        .Cells(Me.CoAll_WU_No.Index).Value = rowDGV1.Index + 1
                        .Cells(Me.CoAll_WU_PlanRate.Index).Value = rowWU.effective_rate
                        .Cells(Me.CoAll_WU_PrimaryWU.Index).Value = rowWU.primary_work_unit
                        .Cells(Me.CoAll_WU_ProdRate.Index).Value = rowWU.rate_per_hour
                        .Cells(Me.CoAll_WU_ProduceQty.Index).Value = 0
                        .Cells(Me.CoAll_WU_StartDate.Index).Value = ""
                        .Cells(Me.CoAll_WU_StartTime.Index).Value = ""
                        .Cells(Me.CoAll_WU_WorkHour.Index).Value = ""
                        .Cells(Me.CoAll_WU_IdentifierWO.Index).Value = ""
                    End With
                Next

                'isikan data mesin yang digunakan ke dgv
                Me.TBLItemWO_ItemSch.DefaultView.RowFilter = "[identifier wo] = '" + rowDGV.Cells(Me.ColAll_Process_IdentifierWO.Index).Value.ToString + "'"
                For Each DV As Data.DataRowView In Me.TBLItemWO_ItemSch.DefaultView
                    rowItemWO_ItemSch = DV.Row
                    Ketemu = False
                    For Each rowDGV1 In Me.DgvWUAllocate_WU.Rows
                        If rowDGV1.Cells(Me.CoAll_WU_PrimaryWU.Index).Value = rowItemWO_ItemSch.primary_work_unit Then
                            With rowDGV1
                                .Cells(Me.CoAll_WU_Allocate.Index).Value = Boolean.TrueString
                                .Cells(Me.CoAll_WU_Efficiency.Index).Value = Me.ObjCD.DigitGrouping(rowItemWO_ItemSch.plan_prod_rate_efficiency, 2)
                                .Cells(Me.CoAll_WU_EndDate.Index).Value = Format(rowItemWO_ItemSch.plan_end_time, "dd MMM yyyy")
                                .Cells(Me.CoAll_WU_EndTime.Index).Value = Format(rowItemWO_ItemSch.plan_end_time, "HH.mm")
                                .Cells(Me.CoAll_WU_PlanRate.Index).Value = Me.ObjCD.DigitGrouping(rowItemWO_ItemSch.plan_prod_rate, rowItemWO_ItemSch._Decimal)
                                .Cells(Me.CoAll_WU_ProdRate.Index).Value = Me.ObjCD.DigitGrouping(rowItemWO_ItemSch.plan_prod_rate_normal, rowItemWO_ItemSch._Decimal)
                                .Cells(Me.CoAll_WU_ProduceQty.Index).Value = Me.ObjCD.DigitGrouping(rowItemWO_ItemSch.qty_produce, rowItemWO_ItemSch._Decimal)
                                .Cells(Me.CoAll_WU_StartDate.Index).Value = Format(rowItemWO_ItemSch.plan_start_time, "dd MMM yyyy")
                                .Cells(Me.CoAll_WU_StartTime.Index).Value = Format(rowItemWO_ItemSch.plan_start_time, "HH.mm")
                                .Cells(Me.CoAll_WU_WorkHour.Index).Value = Me.ObjCD.Convert_ToString(rowItemWO_ItemSch.plan_work_duration)
                                .Cells(Me.CoAll_WU_IdentifierWO.Index).Value = rowItemWO_ItemSch.identifier_wo
                                .Cells(Me.CoAll_WU_DelayDependency.Index).Value = Me.ObjCD.Convert_ToString(rowItemWO_ItemSch.delay_dependency)
                                .Cells(Me.CoAll_WU_TipeDependency.Index).Value = Me.ObjProduct.GetTipeDependencyStr(rowItemWO_ItemSch.tipe_dependency)
                                .Cells(Me.CoAll_WU_PrimaryInputLocation.Index).Value = rowItemWO_ItemSch.primary_input_location
                                .Cells(Me.CoAll_WU_PrimaryOutputLocation.Index).Value = rowItemWO_ItemSch.primary_output_location
                                .Cells(Me.CoAll_WU_InputLocation.Index).Value = Me.ObjLocation.GetStorageAreaCode(rowItemWO_ItemSch.primary_input_location)
                                .Cells(Me.CoAll_WU_OutputLocation.Index).Value = Me.ObjLocation.GetStorageAreaCode(rowItemWO_ItemSch.primary_output_location)
                            End With

                            Ketemu = True
                        End If
                    Next

                    If Ketemu = False Then
                        'tambahkan baris baru, ada kemungkinan dimana item work unit sudah tidak sama dengan yang ada di master
                        Me.DgvWUAllocate_WU.Rows.Add()
                        rowDGV1 = Me.DgvWUAllocate_WU.Rows(Me.DgvWUAllocate_WU.Rows.Count - 1)
                        With rowDGV1
                            .Cells(Me.CoAll_WU_Allocate.Index).Value = Boolean.TrueString
                            .Cells(Me.CoAll_WU_Code.Index).Value = Me.ObjWU.GetCode(rowItemWO_ItemSch.primary_work_unit)
                            .Cells(Me.CoAll_WU_Desc.Index).Value = Me.ObjWU.GetDesc(rowItemWO_ItemSch.primary_work_unit)
                            .Cells(Me.CoAll_WU_Efficiency.Index).Value = Me.ObjCD.DigitGrouping(rowItemWO_ItemSch.plan_prod_rate_efficiency, 2)
                            .Cells(Me.CoAll_WU_EndDate.Index).Value = Format(rowItemWO_ItemSch.plan_end_time, "dd MMM yyyy")
                            .Cells(Me.CoAll_WU_EndTime.Index).Value = Format(rowItemWO_ItemSch.plan_end_time, "HH.mm")
                            .Cells(Me.CoAll_WU_No.Index).Value = rowDGV1.Index + 1
                            .Cells(Me.CoAll_WU_PlanRate.Index).Value = Me.ObjCD.DigitGrouping(rowItemWO_ItemSch.plan_prod_rate, rowItemWO_ItemSch._Decimal)
                            .Cells(Me.CoAll_WU_PrimaryWU.Index).Value = rowItemWO_ItemSch.primary_work_unit
                            .Cells(Me.CoAll_WU_ProdRate.Index).Value = Me.ObjCD.DigitGrouping(rowItemWO_ItemSch.plan_prod_rate_normal, rowItemWO_ItemSch._Decimal)
                            .Cells(Me.CoAll_WU_ProduceQty.Index).Value = Me.ObjCD.DigitGrouping(rowItemWO_ItemSch.qty_produce, rowItemWO_ItemSch._Decimal)
                            .Cells(Me.CoAll_WU_StartDate.Index).Value = Format(rowItemWO_ItemSch.plan_start_time, "dd MMM yyyy")
                            .Cells(Me.CoAll_WU_StartTime.Index).Value = Format(rowItemWO_ItemSch.plan_start_time, "HH.mm")
                            .Cells(Me.CoAll_WU_WorkHour.Index).Value = Me.ObjCD.Convert_ToString(rowItemWO_ItemSch.plan_work_duration)
                            .Cells(Me.CoAll_WU_IdentifierWO.Index).Value = rowItemWO_ItemSch.identifier_wo
                            .Cells(Me.CoAll_WU_DelayDependency.Index).Value = Me.ObjCD.Convert_ToString(rowItemWO_ItemSch.delay_dependency)
                            .Cells(Me.CoAll_WU_TipeDependency.Index).Value = Me.ObjProduct.GetTipeDependencyStr(rowItemWO_ItemSch.tipe_dependency)
                        End With
                    End If
                Next

                'isikan start time
                FirmPlanDate = Me.ObjSch.GetFirmPlan
                For Each rowDGV In Me.DgvWUAllocate_WU.Rows
                    If rowDGV.Cells(Me.CoAll_WU_StartDate.Index).Value = "" Then
                        'cari tanggal end terakhir pada sch bar, untuk wu ybs
                        DummyStart = Me.GetStartTime(rowDGV.Cells(Me.CoAll_WU_PrimaryWU.Index).Value)
                        If DummyStart <> #1/1/1950# Then
                            'bandingkan dengan tanggal terakhir, proses sebelumnya
                            StartDateMin = Me.GetMaxEndTime(Me.DgvWUAllocate_Process.Item(Me.ColAll_Process_IdentifierWO.Index, _
                            Me.DgvWUAllocate_Process.SelectedCells(0).RowIndex).Value.ToString)
                            If StartDateMin > DummyStart Then
                                rowDGV.Cells(Me.CoAll_WU_StartDate.Index).Value = Format(StartDateMin, "dd MMM yyyy")
                                rowDGV.Cells(Me.CoAll_WU_StartTime.Index).Value = Format(StartDateMin, "HH.mm")
                            Else
                                rowDGV.Cells(Me.CoAll_WU_StartDate.Index).Value = Format(DummyStart, "dd MMM yyyy")
                                rowDGV.Cells(Me.CoAll_WU_StartTime.Index).Value = Format(DummyStart, "HH.mm")
                            End If
                        Else
                            'jika ada alokasi, maka masukkan jam mulai pertama pada tgl ybs
                            StartDateMin = Me.ObjSch.GetFirstStartTime(FirmPlanDate + TimeSpan.FromDays(1), rowDGV.Cells(Me.CoAll_WU_PrimaryWU.Index).Value)
                            If StartDateMin <> #1/1/1950# Then
                                rowDGV.Cells(Me.CoAll_WU_StartDate.Index).Value = Format(StartDateMin, "dd MMM yyyy")
                                rowDGV.Cells(Me.CoAll_WU_StartTime.Index).Value = Format(StartDateMin, "HH.mm")
                            Else
                                rowDGV.Cells(Me.CoAll_WU_StartDate.Index).Value = Format(FirmPlanDate, "dd MMM yyyy")
                                rowDGV.Cells(Me.CoAll_WU_StartTime.Index).Value = Format(FirmPlanDate, "HH.mm")
                            End If
                        End If
                    End If
                Next
            End If
            Me.DgvWUAllocate_WU.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            'fokus pada row yang mengerjakana
            For Each rowDGV1 In Me.DgvWUAllocate_WU.Rows
                If rowDGV1.Cells(Me.CoAll_WU_ProduceQty.Index).Value > 0 Then
                    Me.DgvWUAllocate_WU.ClearSelection()
                    rowDGV1.Selected = True
                    Me.DgvWUAllocate_WU.FirstDisplayedScrollingRowIndex = rowDGV1.Index

                    'PrevIdx = Me.DgvWUAllocate_WU.FirstDisplayedScrollingRowIndex
                    'While rowDGV1.Displayed = False
                    '    Me.DgvWUAllocate_WU.FirstDisplayedScrollingRowIndex += 1
                    '    If PrevIdx <> Me.DgvWUAllocate_WU.FirstDisplayedScrollingRowIndex Then
                    '        PrevIdx = Me.DgvWUAllocate_WU.FirstDisplayedScrollingRowIndex
                    '    Else
                    '        Exit Sub
                    '    End If
                    'End While
                    'Me.DgvWUAllocate_WU.FirstDisplayedScrollingRowIndex += 10
                    ''If rowDGV1.Index < Me.DgvWUAllocate_WU.Rows.Count - 1 Then
                    ''    Me.DgvWUAllocate_WU.FirstDisplayedScrollingRowIndex += 1
                    ''End If
                    ''If rowDGV1.Index < Me.DgvWUAllocate_WU.Rows.Count - 1 Then
                    ''    Me.DgvWUAllocate_WU.FirstDisplayedScrollingRowIndex += 1
                    ''End If

                    Exit Sub
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAWU.Dispose()
            TBLWU.Dispose()
        End Try
    End Sub

    Private Sub DgvWUAllocate_Process_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvWUAllocate_Process.CellBeginEdit
        Dim rowDGV As DataGridViewRow

        Try
            If e.ColumnIndex = Me.ColAll_Process_TotalWU.Index Or _
            e.ColumnIndex = Me.ColAll_Process_TipeDependency.Index Or e.ColumnIndex = Me.ColAll_Process_DelayDependency.Index Then
                rowDGV = Me.DgvWUAllocate_Process.Rows(e.RowIndex)
                Me.TBLItemWO.DefaultView.RowFilter = "[Identifier WO Parent] = '" + rowDGV.Cells(Me.ColAll_Process_IdentifierWO.Index).Value + "'"
                If Me.TBLItemWO.DefaultView.Count > 0 Then
                    Me.Buffer = Me.DgvWUAllocate_Process.Item(e.ColumnIndex, e.RowIndex).Value.ToString
                Else
                    e.Cancel = True
                End If
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvWUAllocate_Process_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvWUAllocate_Process.CellClick
        Try
            If e.RowIndex <> -1 Then
                If e.ColumnIndex = Me.ColAll_Process_EditQty.Index Then
                    'tandai posisi
                    Me.r1 = e.RowIndex
                    Me.c1 = e.ColumnIndex
                    Me.IdentifierAcuan = Me.DgvWUAllocate_Process.Item(Me.ColAll_Process_IdentifierWO.Index, e.RowIndex).Value

                    Me.bgw3.RunWorkerAsync()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvWUAllocate_Process_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvWUAllocate_Process.CellEndEdit
        Dim Dummy As String = ""
        Dim JmlWUMax As Double = 0
        Dim rowDGV As DataGridViewRow
        Dim RowItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1Row
        Dim Start As Date
        Dim DummyStartTime As Date
        Dim RowItemWO_ItemSch As DseTrans_Schedule.tbltrans_mo_item_wo_item_schRow
        Dim TipeDependencyInt As Integer = 0
        Dim DelayDependency As Double = 0

        Try
            If e.ColumnIndex = Me.ColAll_Process_TotalWU.Index Then
                If Me.DgvWUAllocate_Process.Item(e.ColumnIndex, e.RowIndex).Value.ToString <> Me.Buffer Then
                    'pastikan angka positif
                    If Me.ObjCD.CekBilanganPositif(Me.DgvWUAllocate_Process.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        MsgBox("Please specify a positive number.", MsgBoxStyle.Information, "Information")
                        Me.DgvWUAllocate_Process.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If

                    'digit grouping
                    Me.DgvWUAllocate_Process.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.DgvWUAllocate_Process.Item(e.ColumnIndex, e.RowIndex).Value, 0)

                    'pastikan total allocated work unit, maksimal, jumlah work unit yang mungkin
                    JmlWUMax = Me.ObjProduct.GetJml_WorkUnitAllocated(Me.DgvWUAllocate_Process.Item(Me.ColAll_Process_PrimaryCodeNo.Index, e.RowIndex).Value)
                    If Me.DgvWUAllocate_Process.Item(e.ColumnIndex, e.RowIndex).Value > JmlWUMax Then
                        MsgBox("Maximum allocated work unit is: " + JmlWUMax.ToString + ".", MsgBoxStyle.Information, "Information")
                        Me.DgvWUAllocate_Process.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If

                    If MsgBox("Change total work unit allocated?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then

                        'catat posisi lama
                        r1 = Me.DgvWUAllocate_Process.SelectedCells(0).RowIndex
                        c1 = Me.DgvWUAllocate_Process.SelectedCells(0).ColumnIndex

                        'catat identifier acuan
                        IdentifierAcuan = Me.DgvWUAllocate_Process.Item(CInt(Me.ColAll_Process_IdentifierWO.Index), CInt(r1)).Value

                        Me.bgw2.RunWorkerAsync()
                    Else
                        Me.DgvWUAllocate_Process.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    End If
                End If
            ElseIf e.ColumnIndex = Me.ColAll_Process_TipeDependency.Index Or e.ColumnIndex = Me.ColAll_Process_DelayDependency.Index Then
                If Me.DgvWUAllocate_Process.Item(e.ColumnIndex, e.RowIndex).Value.ToString <> Me.Buffer Then
                    Start = Now

                    'pastikan format tanggalnya benar kalo dia time string -> termasuk tidak boleh minus
                    If e.ColumnIndex = Me.ColAll_Process_DelayDependency.Index Then
                        If Me.ObjCD.Is_Time_String(Me.DgvWUAllocate_Process.Item(e.ColumnIndex, e.RowIndex).Value, True) = False Then
                            Me.DgvWUAllocate_Process.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                        If Me.ObjCD.Convert_ToMinutes(Me.DgvWUAllocate_Process(e.ColumnIndex, e.RowIndex).Value) < 0 Then
                            Me.DgvWUAllocate_Process.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                    End If
                    rowDGV = Me.DgvWUAllocate_Process.Rows(e.RowIndex)

                    Me.TBLSchBar.DefaultView.RowFilter = "[identifier jo] = '" + Me.IdentifierJO + "'" '+ "' and [identifier wo] ='" + Me.DgvWUAllocate_Process.Item(Me.ColAll_Process_IdentifierWO.Index, e.RowIndex).Value.ToString + "'"
                    'Me.TBLSchBar.DefaultView.Sort = "[plan start time]"

                    TipeDependencyInt = Me.ObjProduct.GetTipeDependency(Me.DgvWUAllocate_Process.Item(Me.ColAll_Process_TipeDependency.Index, e.RowIndex).Value)
                    DelayDependency = Me.ObjCD.Convert_ToMinutes(Me.DgvWUAllocate_Process.Item(Me.ColAll_Process_DelayDependency.Index, e.RowIndex).Value)
                    For Each dv As Data.DataRowView In Me.TBLSchBar.DefaultView
                        rowSch = dv.Row
                        'ubah data tipe dependency
                        For Each RowItemWO In Me.TBLItemWO.Rows
                            If RowItemWO.Identifier = Me.DgvWUAllocate_Process.Item(Me.ColAll_Process_IdentifierWO.Index, e.RowIndex).Value And _
                            RowItemWO.Identifier_JO = Me.IdentifierJO Then
                                RowItemWO.tipe_dependency = TipeDependencyInt
                                RowItemWO.delay_dependency = DelayDependency
                                Exit For
                            Else
                                DummyStartTime = RowItemWO.Plan_Start_Time
                            End If
                        Next
                        For Each RowItemWO_ItemSch In Me.TBLItemWO_ItemSch.Rows
                            If RowItemWO_ItemSch.identifier_wo = Me.DgvWUAllocate_Process.Item(Me.ColAll_Process_IdentifierWO.Index, e.RowIndex).Value And _
                            RowItemWO_ItemSch.identifier_jo = Me.IdentifierJO Then
                                RowItemWO_ItemSch.tipe_dependency = TipeDependencyInt
                                RowItemWO_ItemSch.delay_dependency = DelayDependency

                                If DummyStartTime > RowItemWO_ItemSch.plan_start_time Then
                                    DummyStartTime = RowItemWO_ItemSch.plan_start_time
                                End If
                            End If
                        Next
                        'hitung ulang jadwal baru
                        Me.ObjSch.Reschedule_JadwalBaru1(Me.IdentifierJO, Me.TBLSchBar, Me.TBLItemWO, Me.TBLItemWO_ItemSch, rowSch.a, rowSch.r, rowSch.g, _
                        rowSch.b, Me.TBLSchBarKosong, Me.TBLSchBarTanpaWO, DummyStartTime, Me.PrimaryPredefinedConfig)
                        Me.r1 = e.RowIndex
                        Me.c1 = Me.ColAll_Process_EndTime.Index

                        'MsgBox(Me.ObjCD.DigitGrouping((Now - Start).TotalMinutes, 1))
                        Me.bgw4.RunWorkerAsync()

                        Exit Sub
                    Next
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvWUAllocate_Process_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvWUAllocate_Process.SelectionChanged
        Try
            If Me.DgvWUAllocate_Process.SelectedCells.Count > 0 Then
                If Me.DgvWUAllocate_Process.SelectedCells(0).RowIndex <> Me.PrevIdx Then
                    Me.LoadDGVWU()
                    Me.PrevIdx = Me.DgvWUAllocate_Process.SelectedCells(0).RowIndex
                End If
            Else
                Me.PrevIdx = -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadException()
        Dim rowDGV, rowDGV1 As DataGridViewRow
        Dim RowItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1Row
        Dim ListException1 As New List(Of String)
        Dim CodeNo, Process As String
        Dim HasilSplit As String()
        Dim ListExceptionWorkshiftAllocate As New List(Of Double)

        Try
            Me.dgvException.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvException.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.dgvException.Rows.Clear()

            'target finish date < planned finish date (boleh ignore)
            If Me.dtpTargetFinishDate.Value.Date < Me.dtpPlannedFinishDate.Value.Date Then
                Me.dgvException.Rows.Add()
                rowDGV = Me.dgvException.Rows(Me.dgvException.Rows.Count - 1)
                rowDGV.Cells(Me.ColException_Desc.Index).Value = "Target finish date is smaller than planned finish date."
                rowDGV.Cells(Me.ColException_No.Index).Value = rowDGV.Index + 1
                rowDGV.Cells(Me.ColException_Ignore.Index).Value = 0
            End If

            'ada proses yang selesai sebelum proses anaknya selesai (boleh ignore)
            ListException1.Clear()
            For Each rowDGV1 In Me.DgvWUAllocate_Process.Rows
                Me.TBLItemWO.DefaultView.RowFilter = (Me.ObjSch.GetFilterString_IdentifierWOParent(rowDGV1.Cells(Me.ColAll_Process_IdentifierWO.Index).Value, Me.TBLItemWO)).Replace("[identifier wo]", "identifier")
                For Each DV As Data.DataRowView In Me.TBLItemWO.DefaultView
                    RowItemWO = DV.Row

                    If RowItemWO.Plan_Start_Time > CDate(rowDGV1.Cells(Me.ColAll_Process_StartDate.Index).Value + " " + rowDGV1.Cells(Me.ColAll_Process_StartTime.Index).Value.ToString.Replace(".", ":")) Then
                        ListException1.Add(rowDGV1.Cells(Me.ColAll_Process_Code.Index).Value.ToString + ";" + rowDGV1.Cells(Me.ColAll_Process_CodeNo.Index).Value.ToString)
                    ElseIf RowItemWO.Plan_End_Time > CDate(rowDGV1.Cells(Me.ColAll_Process_EndDate.Index).Value + " " + rowDGV1.Cells(Me.ColAll_Process_EndTime.Index).Value.ToString.Replace(".", ":")) Then
                        Me.dgvException.Rows.Add()
                        rowDGV = Me.dgvException.Rows(Me.dgvException.Rows.Count - 1)
                        rowDGV.Cells(Me.ColException_Desc.Index).Value = "Work is done before the previous process is done: " + rowDGV1.Cells(Me.ColAll_Process_Code.Index).Value.ToString + " : " + rowDGV1.Cells(Me.ColAll_Process_CodeNo.Index).Value.ToString
                        rowDGV.Cells(Me.ColException_No.Index).Value = rowDGV.Index + 1
                        rowDGV.Cells(Me.ColException_Ignore.Index).Value = 0
                    End If
                Next
            Next

            For Each Kombinasi As String In ListException1
                HasilSplit = Kombinasi.Split(";")
                CodeNo = HasilSplit(1)
                Process = HasilSplit(0)
                Me.dgvException.Rows.Add()
                rowDGV = Me.dgvException.Rows(Me.dgvException.Rows.Count - 1)
                rowDGV.Cells(Me.ColException_Desc.Index).Value = "Work is done before the previous process is started: " + CodeNo + " : " + Process
                rowDGV.Cells(Me.ColException_No.Index).Value = rowDGV.Index + 1
                rowDGV.Cells(Me.ColException_Ignore.Index).Value = 0
            Next

            'ada item yang selesai 1 jan 2200, artinya tidak selesai, atau workshift allocate tidak cukup
            'tidak boleh ignore
            For Each rowDGV1 In Me.DgvWUAllocate_Process.Rows
                If rowDGV1.Cells(Me.ColAll_Process_EndDate.Index).Value = #1/1/2200# Then
                    If ListExceptionWorkshiftAllocate.Contains(rowDGV1.Cells(Me.ColAll_Process_PrimaryProcess.Index).Value) = False Then
                        Me.dgvException.Rows.Add()

                        rowDGV = Me.dgvException.Rows(Me.dgvException.Rows.Count - 1)
                        rowDGV.Cells(Me.ColException_Desc.Index).Value = "Work shift allocation is need on process: " + _
                        rowDGV1.Cells(Me.ColAll_Process_Code.Index).Value
                        rowDGV.Cells(Me.ColException_No.Index).Value = rowDGV.Index + 1
                        rowDGV.Cells(Me.ColException_Ignore.Index).Value = 1
                        ListExceptionWorkshiftAllocate.Add(rowDGV1.Cells(Me.ColAll_Process_PrimaryProcess.Index).Value)
                    End If
                End If
            Next
            Me.dgvException.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        'Dim a, r, g, b As Integer
        'Dim rowSchItem As DseTrans_Schedule.tbltrans_mo_item_wo_item_schRow

        Try
            'cek exception
            If Me.ApakahDataBenar = False Then
                Exit Sub
            End If
            'rowSchItem = Me.TBLItemWO_ItemSch.Rows(0)
            'a = rowSchItem.a
            'r = rowSchItem.r
            'g = rowSchItem.g
            'b = rowSchItem.b

            If MsgBox("Save data?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                If Me.ObjSch.SimpanJadwalBaru(Me.TBLSchBar, Me.IdentifierJO, Me.TBLItemWO, Me.TBLItemWO_ItemSch, Me.pb_Sch.BackColor.A, Me.pb_Sch.BackColor.R, _
                Me.pb_Sch.BackColor.G, Me.pb_Sch.BackColor.B, Me.PrimaryPredefinedConfig) = True Then
                    MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                    'Me.SimpanPredefined()
                    Me.LangsungTutup = True
                    Me.Close()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataBenar() As Boolean
        'MsgBox("belum: cek exception")

        Try
            Me.LoadException()
            If Me.dgvException.Rows.Count > 0 Then
                Dim ij As New FrmTrans_ProductionSchedule_Exception
                ij.FrmRef = Me
                ij.BackColorAll = Me.BackColorAll

                If ij.ShowDialog <> Windows.Forms.DialogResult.OK Then
                    ij.Dispose()
                    Return False
                End If
                ij.Dispose()
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        'Try
        '    If MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
        '        Me.LangsungTutup = True
        '        Me.Close()
        '    Else
        '        Me.LangsungTutup = False
        '        Exit Sub
        '    End If
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'End Try
    End Sub

    Private Sub DgvWUAllocate_WU_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvWUAllocate_WU.CellBeginEdit
        Dim Jml As Integer = 0
        Dim rowDGV As DataGridViewRow
        Dim DummyTime As Date

        Try
            If e.ColumnIndex = Me.CoAll_WU_Allocate.Index Then
                For Each rowDGV In Me.DgvWUAllocate_WU.Rows
                    If rowDGV.Cells(Me.CoAll_WU_Allocate.Index).Value = Boolean.TrueString Then
                        Jml += 1
                    End If
                Next
                If Jml > 1 Then
                    Me.Buffer = Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value.ToString
                Else
                    If Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value = Boolean.TrueString Then
                        e.Cancel = True
                    Else
                        Me.Buffer = Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value.ToString
                    End If
                End If
            ElseIf e.ColumnIndex = Me.CoAll_WU_EndDate.Index Then 'or e.ColumnIndex = Me.CoAll_WU_EndTime.Index Or
                If CDec(Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_ProduceQty.Index, e.RowIndex).Value) > 0 Then
                    'masukkan nilainya menjadi datetime
                    Me.Buffer = Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value
                    DummyTime = CDate(Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_EndDate.Index, e.RowIndex).Value).Date + TimeSpan.FromMinutes(Me.ObjCD.Convert_ToMinutes(Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_EndTime.Index, e.RowIndex).Value))
                    Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value = Format(DummyTime, "dd MMM yyyy HH.mm")
                    Me.BufferTime = DummyTime
                    Me.r1 = Me.DgvWUAllocate_Process.SelectedCells(0).RowIndex
                    Me.c1 = Me.DgvWUAllocate_Process.SelectedCells(0).ColumnIndex
                    Me.r2 = Me.DgvWUAllocate_WU.SelectedCells(0).RowIndex
                    Me.c2 = Me.DgvWUAllocate_WU.SelectedCells(0).ColumnIndex
                Else
                    e.Cancel = True
                End If
            ElseIf e.ColumnIndex = Me.CoAll_WU_StartDate.Index Then
                If CDec(Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_ProduceQty.Index, e.RowIndex).Value) > 0 Then
                    'masukkan nilainya menjadi datetime
                    Me.Buffer = Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value
                    DummyTime = CDate(Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_StartDate.Index, e.RowIndex).Value).Date + TimeSpan.FromMinutes(Me.ObjCD.Convert_ToMinutes(Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_StartTime.Index, e.RowIndex).Value))
                    Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value = Format(DummyTime, "dd MMM yyyy HH.mm")
                    Me.BufferTime = DummyTime
                    Me.r1 = Me.DgvWUAllocate_Process.SelectedCells(0).RowIndex
                    Me.c1 = Me.DgvWUAllocate_Process.SelectedCells(0).ColumnIndex
                    Me.r2 = Me.DgvWUAllocate_WU.SelectedCells(0).RowIndex
                    Me.c2 = Me.DgvWUAllocate_WU.SelectedCells(0).ColumnIndex
                Else
                    e.Cancel = True
                End If
            ElseIf e.ColumnIndex = Me.CoAll_WU_InputLocation.Index Or e.ColumnIndex = Me.CoAll_WU_OutputLocation.Index Then
                If CDec(Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_ProduceQty.Index, e.RowIndex).Value) > 0 Then
                    Me.Buffer = Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            ElseIf e.ColumnIndex = Me.CoAll_WU_PlanRate.Index Or e.ColumnIndex = Me.CoAll_WU_WorkHour.Index Then
                If CDec(Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_ProduceQty.Index, e.RowIndex).Value) > 0 Then
                    'catat buffer
                    Me.Buffer = Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvWUAllocate_WU_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvWUAllocate_WU.CellEndEdit
        Dim Dummy As String = ""
        Dim rowDGV As DataGridViewRow
        Dim DummyTime As Date = Today
        Dim DummyTime1 As Date = Today
        Dim DummyTime2 As Date = Today
        Dim DummyTime3 As Date = #1/1/1950#
        Dim TimeComponent As String()
        Dim dd, MMM, yyyy, HHmm As String
        Dim FormatBetul As Boolean = True
        Dim rowSch1 As DseTrans_Schedule.tbl_sch_barRow
        Dim RowItemWO_ItemSch As DseTrans_Schedule.tbltrans_mo_item_wo_item_schRow
        Dim Start As Date
        Dim JmlItemSch As Integer = 0
        Dim NewDelayDependency As Double = 0
        Dim RowItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1Row
        Dim PrimaryLocation As Double = -1

        Try
            If e.ColumnIndex = Me.CoAll_WU_Allocate.Index Then
                If Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value.ToString <> Me.Buffer Then
                    If MsgBox("Change work unit allocation?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                        'catat posisi lama
                        r1 = Me.DgvWUAllocate_Process.SelectedCells(0).RowIndex
                        r2 = Me.DgvWUAllocate_WU.SelectedCells(0).RowIndex
                        c1 = Me.DgvWUAllocate_Process.SelectedCells(0).ColumnIndex
                        c2 = Me.DgvWUAllocate_WU.SelectedCells(0).ColumnIndex

                        'catat identifier acuan
                        IdentifierAcuan = Me.DgvWUAllocate_Process.Item(CInt(Me.ColAll_Process_IdentifierWO.Index), CInt(r1)).Value

                        'catat list primary wu yang dialokasikan
                        Me.ListPrimaryWU.Clear()
                        For Each rowDGV In Me.DgvWUAllocate_WU.Rows
                            If rowDGV.Cells(Me.CoAll_WU_Allocate.Index).Value = Boolean.TrueString Then
                                If Dummy = "" Then
                                    Dummy = rowDGV.Cells(Me.CoAll_WU_PrimaryWU.Index).Value.ToString
                                Else
                                    Dummy = Dummy + "-" + rowDGV.Cells(Me.CoAll_WU_PrimaryWU.Index).Value.ToString
                                End If
                            End If
                        Next
                        Me.ListPrimaryWU.Add(Dummy)

                        Me.bgw1.RunWorkerAsync()
                    Else
                        Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    End If
                End If
            ElseIf e.ColumnIndex = Me.CoAll_WU_InputLocation.Index Or e.ColumnIndex = Me.CoAll_WU_OutputLocation.Index Then
                If IsNothing(Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                    Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                Else
                    If Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                        Exit Sub
                    End If
                    If Me.ObjLocation.ApakahSudahAda(Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                        PrimaryLocation = Me.ObjLocation.GetPrimary(Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value)
                        If e.ColumnIndex = Me.CoAll_WU_PrimaryInputLocation.Index Then
                            Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_PrimaryInputLocation.Index, e.RowIndex).Value = PrimaryLocation
                        ElseIf e.ColumnIndex = Me.CoAll_WU_PrimaryOutputLocation.Index Then
                            Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_PrimaryOutputLocation.Index, e.RowIndex).Value = PrimaryLocation
                        End If
                    Else
                        Dim ij As New Core_Module.FrmCariMasterStorageArea
                        ij.txtCari.Text = Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value
                        ij.BackColorAll = Me.BackColorAll
                        ij.ShowDialog()
                        If ij.PrimaryStorageArea <> -1 Then
                            PrimaryLocation = ij.PrimaryStorageArea
                            If e.ColumnIndex = Me.CoAll_WU_InputLocation.Index Then
                                Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_PrimaryInputLocation.Index, e.RowIndex).Value = PrimaryLocation
                                Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_InputLocation.Index, e.RowIndex).Value = Me.ObjLocation.GetStorageAreaCode(PrimaryLocation)
                            Else
                                Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_PrimaryOutputLocation.Index, e.RowIndex).Value = PrimaryLocation
                                Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_OutputLocation.Index, e.RowIndex).Value = Me.ObjLocation.GetStorageAreaCode(PrimaryLocation)
                            End If
                        Else
                            Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                    End If
                    'set upper case
                    Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper

                    'set primary location input ke tbl wo item sch
                    Me.TBLSchBar.DefaultView.RowFilter = "[identifier wo] = '" + Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_IdentifierWO.Index, e.RowIndex).Value.ToString + _
                    "' AND [primary work unit] = " + Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_PrimaryWU.Index, e.RowIndex).Value.ToString
                    If Me.TBLSchBar.DefaultView.Count <> 1 Then
                        MsgBox("Logic error: invalid row count. Please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                        Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    For Each RowItemWO_ItemSch In Me.TBLItemWO_ItemSch.Rows
                        If RowItemWO_ItemSch.identifier_wo = Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_IdentifierWO.Index, e.RowIndex).Value.ToString And _
                        RowItemWO_ItemSch.primary_work_unit = Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_PrimaryWU.Index, e.RowIndex).Value And _
                        RowItemWO_ItemSch.identifier_jo = Me.IdentifierJO Then
                            If e.ColumnIndex = Me.CoAll_WU_InputLocation.Index Then
                                RowItemWO_ItemSch.primary_input_location = PrimaryLocation
                            ElseIf e.ColumnIndex = Me.CoAll_WU_OutputLocation.Index Then
                                RowItemWO_ItemSch.primary_output_location = PrimaryLocation
                            End If
                            Exit For
                        End If
                    Next
                End If
            ElseIf e.ColumnIndex = Me.CoAll_WU_WorkHour.Index Then
                Start = Now
                If IsNothing(Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                    Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                Else
                    'pastikan time string
                    If Me.ObjCD.Is_Time_String(Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        MsgBox("Please time string format HH.mm.", MsgBoxStyle.Information, "Information")
                        Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    'pastikan time string positif
                    If Me.ObjCD.Convert_ToMinutes(Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value) <= 0 Then
                        MsgBox("Please specify a positive value time string.", MsgBoxStyle.Information, "Information")
                        Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If

                    Me.TBLSchBar.DefaultView.RowFilter = "[identifier wo] = '" + Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_IdentifierWO.Index, e.RowIndex).Value.ToString + _
                                         "' AND [primary work unit] = " + Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_PrimaryWU.Index, e.RowIndex).Value.ToString
                    If Me.TBLSchBar.DefaultView.Count <> 1 Then
                        MsgBox("Logic error: invalid row count. Please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                        Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    For Each DV As Data.DataRowView In Me.TBLSchBar.DefaultView
                        rowSch1 = DV.Row
                        'digit grouping
                        Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value = _
                        Me.ObjCD.DigitGrouping(Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value, CInt(rowSch1.decimal_point))

                        'reschedule
                        rowSch1.total_work_hour = Me.ObjCD.Convert_ToMinutes(Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value)
                        rowSch1.plan_end_time = Me.ObjSch.GetEndTime(rowSch1.primary_work_unit, rowSch1.plan_start_time, rowSch1.total_work_hour, Me.TBLSchBarTanpaWO)
                        rowSch1.rate = Math.Round((rowSch1.qty_produce * 60 / rowSch1.total_work_hour), CInt(rowSch1.decimal_point), MidpointRounding.AwayFromZero)

                        'update data item wo sch
                        For Each RowItemWO_ItemSch In Me.TBLItemWO_ItemSch.Rows
                            If RowItemWO_ItemSch.identifier_wo = Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_IdentifierWO.Index, e.RowIndex).Value.ToString And _
                            RowItemWO_ItemSch.primary_work_unit = Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_PrimaryWU.Index, e.RowIndex).Value And _
                            RowItemWO_ItemSch.identifier_jo = Me.IdentifierJO Then
                                RowItemWO_ItemSch.plan_end_time = rowSch1.plan_end_time
                                RowItemWO_ItemSch.plan_end_time_minimum = RowItemWO_ItemSch.plan_end_time
                                RowItemWO_ItemSch.plan_work_duration = rowSch1.total_work_hour
                                RowItemWO_ItemSch.plan_total_duration = RowItemWO_ItemSch.plan_work_duration + RowItemWO_ItemSch.plan_setup_duration + RowItemWO_ItemSch.plan_downtime
                                RowItemWO_ItemSch.plan_prod_rate = rowSch1.rate
                                RowItemWO_ItemSch.plan_prod_rate_normal = Math.Round(RowItemWO_ItemSch.plan_prod_rate * 100 / RowItemWO_ItemSch.plan_prod_rate_efficiency, CInt(RowItemWO_ItemSch._Decimal), MidpointRounding.AwayFromZero)

                                For Each RowItemWO In Me.TBLItemWO.Rows
                                    If RowItemWO.Identifier = RowItemWO_ItemSch.identifier_wo Then
                                        Exit For
                                    Else
                                        DummyTime3 = RowItemWO.Plan_Start_Time
                                    End If
                                Next

                                Exit For
                            End If
                        Next

                        'geser sch bar
                        Me.ObjSch.Reschedule_JadwalBaru1(Me.IdentifierJO, Me.TBLSchBar, Me.TBLItemWO, Me.TBLItemWO_ItemSch, _
                        rowSch1.a, rowSch1.r, rowSch1.g, rowSch1.b, Me.TBLSchBarKosong, Me.TBLSchBarTanpaWO, DummyTime3, Me.PrimaryPredefinedConfig)

                        Me.r1 = Me.DgvWUAllocate_Process.SelectedCells(0).RowIndex
                        Me.c1 = Me.DgvWUAllocate_Process.SelectedCells(0).ColumnIndex
                        Me.r2 = e.RowIndex
                        Me.c2 = Me.CoAll_WU_EndDate.Index
                        'MsgBox(Me.ObjCD.DigitGrouping((Now - Start).TotalMinutes, 1))
                        Me.bgw5.RunWorkerAsync()
                    Next
                End If
            ElseIf e.ColumnIndex = Me.CoAll_WU_PlanRate.Index Then
                Start = Now
                If IsNothing(Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                    Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                Else
                    'pastikan angka positif
                    If Me.ObjCD.CekBilanganPositif(Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        MsgBox("Please specify a positif number for rate data.", MsgBoxStyle.Information, "Information")
                        Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If

                    Me.TBLSchBar.DefaultView.RowFilter = "[identifier wo] = '" + Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_IdentifierWO.Index, e.RowIndex).Value.ToString + _
                      "' AND [primary work unit] = " + Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_PrimaryWU.Index, e.RowIndex).Value.ToString
                    If Me.TBLSchBar.DefaultView.Count <> 1 Then
                        MsgBox("Logic error: invalid row count. Please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                        Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    For Each DV As Data.DataRowView In Me.TBLSchBar.DefaultView
                        rowSch1 = DV.Row
                        'digit grouping
                        Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value = _
                        Me.ObjCD.DigitGrouping(Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value, CInt(rowSch1.decimal_point))

                        'reschedule
                        rowSch1.rate = CDec(Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value)
                        rowSch1.total_work_hour = Me.ObjCD.RoundUp1(rowSch1.qty_produce * 60 / rowSch1.rate, 0)
                        rowSch1.plan_end_time = Me.ObjSch.GetEndTime(rowSch1.primary_work_unit, rowSch1.plan_start_time, rowSch1.total_work_hour, Me.TBLSchBarTanpaWO)

                        'update data item wo sch
                        For Each RowItemWO_ItemSch In Me.TBLItemWO_ItemSch.Rows
                            If RowItemWO_ItemSch.identifier_wo = Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_IdentifierWO.Index, e.RowIndex).Value.ToString And _
                            RowItemWO_ItemSch.primary_work_unit = Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_PrimaryWU.Index, e.RowIndex).Value And _
                            RowItemWO_ItemSch.identifier_jo = Me.IdentifierJO Then
                                RowItemWO_ItemSch.plan_end_time = rowSch1.plan_end_time
                                RowItemWO_ItemSch.plan_end_time_minimum = RowItemWO_ItemSch.plan_end_time
                                RowItemWO_ItemSch.plan_work_duration = rowSch1.total_work_hour
                                RowItemWO_ItemSch.plan_total_duration = RowItemWO_ItemSch.plan_work_duration + RowItemWO_ItemSch.plan_setup_duration + RowItemWO_ItemSch.plan_downtime
                                RowItemWO_ItemSch.plan_prod_rate = rowSch1.rate
                                RowItemWO_ItemSch.plan_prod_rate_normal = Math.Round(RowItemWO_ItemSch.plan_prod_rate * 100 / RowItemWO_ItemSch.plan_prod_rate_efficiency, CInt(RowItemWO_ItemSch._Decimal), MidpointRounding.AwayFromZero)

                                For Each RowItemWO In Me.TBLItemWO.Rows
                                    If RowItemWO.Identifier = RowItemWO_ItemSch.identifier_wo Then
                                        Exit For
                                    Else
                                        DummyTime3 = RowItemWO.Plan_Start_Time
                                    End If
                                Next

                                Exit For
                            End If
                        Next

                        'geser sch bar
                        Me.ObjSch.Reschedule_JadwalBaru1(Me.IdentifierJO, Me.TBLSchBar, Me.TBLItemWO, Me.TBLItemWO_ItemSch, _
                        rowSch1.a, rowSch1.r, rowSch1.g, rowSch1.b, Me.TBLSchBarKosong, Me.TBLSchBarTanpaWO, DummyTime3, Me.PrimaryPredefinedConfig)

                        Me.r1 = Me.DgvWUAllocate_Process.SelectedCells(0).RowIndex
                        Me.c1 = Me.DgvWUAllocate_Process.SelectedCells(0).ColumnIndex
                        Me.r2 = e.RowIndex
                        Me.c2 = Me.CoAll_WU_EndDate.Index
                        'MsgBox(Me.ObjCD.DigitGrouping((Now - Start).TotalMinutes, 1))
                        Me.bgw5.RunWorkerAsync()
                    Next
                End If
            ElseIf e.ColumnIndex = Me.CoAll_WU_StartDate.Index Then
                Start = Now
                If IsNothing(Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                    Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                ElseIf e.ColumnIndex = Me.CoAll_WU_StartDate.Index Then
                    TimeComponent = Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Split(" ")
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
                        If Me.ObjCD.Is_Time_String(HHmm) = False Then
                            FormatBetul = False
                        End If

                        If FormatBetul = True Then
                            DummyTime = CDate(dd + " " + MMM + " " + yyyy).Date + TimeSpan.FromMinutes(Me.ObjCD.Convert_ToMinutes(HHmm))
                            If CDate(Me.BufferTime) = DummyTime Then
                                Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                Exit Sub
                            End If
                        End If
                    End If

                    If FormatBetul = False Then
                        MsgBox("Please specify format with pattern dd MMM yyy HH.mm.", MsgBoxStyle.Information, "Information")
                        Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    Else
                        Me.TBLItemWO_ItemSch.DefaultView.RowFilter = "[identifier wo] = '" + Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_IdentifierWO.Index, e.RowIndex).Value.ToString + "'"
                        JmlItemSch = Me.TBLItemWO_ItemSch.DefaultView.Count

                        DummyTime1 = CDate(Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_EndDate.Index, e.RowIndex).Value).Date + _
                        TimeSpan.FromMinutes(Me.ObjCD.Convert_ToMinutes(Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_EndTime.Index, e.RowIndex).Value))

                        Me.TBLSchBar.DefaultView.RowFilter = "[identifier wo] = '" + Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_IdentifierWO.Index, e.RowIndex).Value.ToString + _
                        "' AND [primary work unit] = " + Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_PrimaryWU.Index, e.RowIndex).Value.ToString
                        If Me.TBLSchBar.DefaultView.Count <> 1 Then
                            MsgBox("Logic error: invalid row count. Please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                            Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                        For Each DV As Data.DataRowView In Me.TBLSchBar.DefaultView
                            rowSch1 = DV.Row

                            If rowSch1.plan_start_time = BufferTime And rowSch1.plan_end_time = DummyTime1 Then
                                DummyTime2 = Me.ObjSch.GetEndTime(rowSch1.primary_work_unit, DummyTime, rowSch1.total_work_hour, Me.TBLSchBarTanpaWO)
                                If DummyTime2 = #1/1/2200# Then
                                    MsgBox("Insufficient workshift allocation, please allocate more workshift for " + Me.ObjWU.GetCode(rowSch1.primary_work_unit) + ".", MsgBoxStyle.Information, "Information")
                                    Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Exit Sub
                                End If
                                If Me.ObjSch.ApakahBolehGeserStartTime(rowSch1, Me.TBLSchBarKosong, Me.TBLSchBar, DummyTime, DummyTime2, rowSch1.plan_start_time, _
                                Me.ObjProduct.GetTipeDependency(Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_TipeDependency.Index, e.RowIndex).Value), Me.TBLItemWO) = False Then
                                    Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Exit Sub
                                Else
                                    'update data sch bar
                                    rowSch1.plan_start_time = DummyTime
                                    rowSch1.plan_end_time = DummyTime2

                                    'update data item wo sch
                                    For Each RowItemWO_ItemSch In Me.TBLItemWO_ItemSch.Rows
                                        If RowItemWO_ItemSch.identifier_wo = Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_IdentifierWO.Index, e.RowIndex).Value.ToString And _
                                        RowItemWO_ItemSch.primary_work_unit = Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_PrimaryWU.Index, e.RowIndex).Value And _
                                        RowItemWO_ItemSch.identifier_jo = Me.IdentifierJO Then
                                            'khusus kalo menggunakan 1 mesin, update data delay ke item wo sch dan item wo
                                            'agar kalo ada perubahan lain, start time tidak bergeser
                                            'untuk yang multi mesin belum tau caranya, nanti aja klo ketemu kasusnya
                                            If JmlItemSch = 1 Then
                                                For Each RowItemWO In Me.TBLItemWO.Rows
                                                    If RowItemWO.Identifier = RowItemWO_ItemSch.identifier_wo Then
                                                        NewDelayDependency = Me.GetDelayDependency(RowItemWO, rowSch1.plan_start_time, rowSch1.primary_work_unit)
                                                        RowItemWO_ItemSch.delay_dependency = NewDelayDependency

                                                        RowItemWO.delay_dependency = RowItemWO_ItemSch.delay_dependency

                                                        Exit For
                                                    Else
                                                        DummyTime3 = RowItemWO.Plan_Start_Time + TimeSpan.FromMinutes(1)
                                                    End If
                                                Next
                                            Else
                                                For Each RowItemWO In Me.TBLItemWO.Rows
                                                    If RowItemWO.Identifier = RowItemWO_ItemSch.identifier_wo Then
                                                        Exit For
                                                    Else
                                                        DummyTime3 = RowItemWO.Plan_Start_Time + TimeSpan.FromMinutes(1)
                                                    End If
                                                Next
                                            End If

                                            RowItemWO_ItemSch.plan_start_time = rowSch1.plan_start_time
                                            RowItemWO_ItemSch.plan_end_time = rowSch1.plan_end_time
                                            RowItemWO_ItemSch.plan_end_time_minimum = RowItemWO_ItemSch.plan_end_time

                                            Exit For
                                        End If
                                    Next

                                    'geser sch bar
                                    Me.ObjSch.Reschedule_JadwalBaru1(Me.IdentifierJO, Me.TBLSchBar, Me.TBLItemWO, Me.TBLItemWO_ItemSch, _
                                    rowSch1.a, rowSch1.r, rowSch1.g, rowSch1.b, Me.TBLSchBarKosong, Me.TBLSchBarTanpaWO, DummyTime3, _
                                    Me.PrimaryPredefinedConfig, rowSch1.primary_work_unit, rowSch1.identifier_wo)

                                    Me.r1 = Me.DgvWUAllocate_Process.SelectedCells(0).RowIndex
                                    Me.c1 = Me.DgvWUAllocate_Process.SelectedCells(0).ColumnIndex
                                    Me.r2 = e.RowIndex
                                    Me.c2 = Me.CoAll_WU_StartDate.Index
                                    'MsgBox(Me.ObjCD.DigitGrouping((Now - Start).TotalMinutes, 1))
                                    Me.bgw5.RunWorkerAsync()
                                End If
                            End If
                        Next
                    End If
                End If
            ElseIf e.ColumnIndex = Me.CoAll_WU_EndDate.Index Then
                Start = Now
                If IsNothing(Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                    Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                ElseIf e.ColumnIndex = Me.CoAll_WU_EndDate.Index Then
                    TimeComponent = Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Split(" ")
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
                        If Me.ObjCD.Is_Time_String(HHmm) = False Then
                            FormatBetul = False
                        End If

                        If FormatBetul = True Then
                            DummyTime = CDate(dd + " " + MMM + " " + yyyy).Date + TimeSpan.FromMinutes(Me.ObjCD.Convert_ToMinutes(HHmm))
                            If CDate(Me.BufferTime) = DummyTime Then
                                Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                Exit Sub
                            End If
                        End If
                    End If

                    If FormatBetul = False Then
                        MsgBox("Please specify format with pattern dd MMM yyy HH.mm.", MsgBoxStyle.Information, "Information")
                        Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    Else
                        DummyTime1 = CDate(Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_StartDate.Index, e.RowIndex).Value).Date + _
                        TimeSpan.FromMinutes(Me.ObjCD.Convert_ToMinutes(Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_StartTime.Index, e.RowIndex).Value))

                        Me.TBLSchBar.DefaultView.RowFilter = "[identifier wo] = '" + Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_IdentifierWO.Index, e.RowIndex).Value.ToString + _
                        "' AND [primary work unit] = " + Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_PrimaryWU.Index, e.RowIndex).Value.ToString
                        If Me.TBLSchBar.DefaultView.Count <> 1 Then
                            MsgBox("Logic error: invalid row count. Please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                            Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        End If
                        For Each DV As Data.DataRowView In Me.TBLSchBar.DefaultView
                            rowSch1 = DV.Row

                            If rowSch1.plan_end_time = BufferTime And rowSch1.plan_start_time = DummyTime1 Then
                                If Me.ObjSch.ApakahBolehGeserEndTime(rowSch1, Me.TBLSchBarTanpaWO, DummyTime) = False Then
                                    Me.DgvWUAllocate_WU.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Exit Sub
                                Else
                                    'update data sch bar
                                    rowSch1.plan_end_time = DummyTime
                                    rowSch1.total_work_hour = Me.ObjSch.GetWorkDuration(rowSch1.plan_start_time, rowSch1.plan_end_time, rowSch1.primary_work_unit, Me.TBLSchBarTanpaWO)
                                    rowSch1.rate = Math.Round((rowSch1.qty_produce * 60 / rowSch1.total_work_hour), CInt(rowSch1.decimal_point), MidpointRounding.AwayFromZero)

                                    'update data item wo sch
                                    For Each RowItemWO_ItemSch In Me.TBLItemWO_ItemSch.Rows
                                        If RowItemWO_ItemSch.identifier_wo = Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_IdentifierWO.Index, e.RowIndex).Value.ToString And _
                                        RowItemWO_ItemSch.primary_work_unit = Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_PrimaryWU.Index, e.RowIndex).Value Then
                                            RowItemWO_ItemSch.plan_end_time = rowSch1.plan_end_time
                                            RowItemWO_ItemSch.plan_end_time_minimum = RowItemWO_ItemSch.plan_end_time
                                            RowItemWO_ItemSch.plan_work_duration = rowSch1.total_work_hour
                                            RowItemWO_ItemSch.plan_total_duration = RowItemWO_ItemSch.plan_work_duration + RowItemWO_ItemSch.plan_setup_duration + RowItemWO_ItemSch.plan_downtime
                                            RowItemWO_ItemSch.plan_prod_rate = rowSch1.rate
                                            RowItemWO_ItemSch.plan_prod_rate_normal = Math.Round(RowItemWO_ItemSch.plan_prod_rate * 100 / RowItemWO_ItemSch.plan_prod_rate_efficiency, CInt(RowItemWO_ItemSch._Decimal), MidpointRounding.AwayFromZero)

                                            For Each RowItemWO In Me.TBLItemWO.Rows
                                                If RowItemWO.Identifier = RowItemWO_ItemSch.identifier_wo Then
                                                    Exit For
                                                Else
                                                    DummyTime3 = RowItemWO.Plan_Start_Time
                                                End If
                                            Next

                                            Exit For
                                        End If
                                    Next

                                    'geser sch bar
                                    Me.ObjSch.Reschedule_JadwalBaru1(Me.IdentifierJO, Me.TBLSchBar, Me.TBLItemWO, Me.TBLItemWO_ItemSch, _
                                    rowSch1.a, rowSch1.r, rowSch1.g, rowSch1.b, Me.TBLSchBarKosong, Me.TBLSchBarTanpaWO, DummyTime3, Me.PrimaryPredefinedConfig)

                                    Me.r1 = Me.DgvWUAllocate_Process.SelectedCells(0).RowIndex
                                    Me.c1 = Me.DgvWUAllocate_Process.SelectedCells(0).ColumnIndex
                                    Me.r2 = e.RowIndex
                                    Me.c2 = Me.CoAll_WU_EndDate.Index
                                    'MsgBox(Me.ObjCD.DigitGrouping((Now - Start).TotalMinutes, 1))
                                    Me.bgw5.RunWorkerAsync()
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

    Private Function GetDelayDependency(ByVal vRowItemWOAcuan As DseTrans_Schedule.tbltrans_mo_item_wo1Row, ByVal vStartDate As Date, ByVal vPrimaryWU As Double) As Double
        Dim RowItemWOTarget As DseTrans_Schedule.tbltrans_mo_item_wo1Row
        Dim RefDate As Date = Date.MinValue
        Dim TipeDependency As Core_Module.clsMasterProduct.EnumTipeDependency = Core_Module.clsMasterProduct.EnumTipeDependency.End_Start1
        Dim RefDate1 As Date = Today

        Try
            'cari item wo parent,
            Me.TBLItemWO.DefaultView.RowFilter = "[Identifier WO Parent] = '" + vRowItemWOAcuan.Identifier + "'"

            'berdasarkan tipe dependency, tentukan delay dependency yang terjadi
            If Me.TBLItemWO.DefaultView.Count = 0 Then
                'tidak ada dependency, kembalikan nilai asal
                Return vRowItemWOAcuan.delay_dependency
            Else
                TipeDependency = vRowItemWOAcuan.tipe_dependency

                For Each DV As Data.DataRowView In Me.TBLItemWO.DefaultView
                    RowItemWOTarget = DV.Row
                    If TipeDependency = Core_Module.clsMasterProduct.EnumTipeDependency.End_Start1 Then
                        RefDate1 = RowItemWOTarget.Plan_End_Time
                    ElseIf TipeDependency = Core_Module.clsMasterProduct.EnumTipeDependency.Start_Start_Strict1 Or TipeDependency = Core_Module.clsMasterProduct.EnumTipeDependency.Start_Start1 Then
                        RefDate1 = RowItemWOTarget.Plan_Start_Time
                    Else
                        RefDate1 = RowItemWOTarget.Plan_Start_Time
                    End If
                    If RefDate < RefDate1 Then
                        RefDate = RefDate1
                    End If
                Next

                Return Me.ObjSch.GetWorkDuration(RefDate, vStartDate, vPrimaryWU, Me.TBLSchBarTanpaWO)
            End If
            Return vRowItemWOAcuan.delay_dependency
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return vRowItemWOAcuan.delay_dependency
        End Try
    End Function

    Private Sub bgw1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgw1.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgw1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgw1.RunWorkerCompleted
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'hitung ulang item modifikasi
            Me.ObjSch.HitungJadwalModifikasi(Me.TBLSchBar, Me.IdentifierJO, Me.TBLItemWO, Me.TBLItemWO_ItemSch, IdentifierAcuan, ListPrimaryWU, Me.PrimaryPredefinedConfig)

            'load dgv proses + dgv wu
            Me.LoadDGVProses(r1, c1)

            'kembalikan fokus ke kolom semula
            Me.DgvWUAllocate_WU.ClearSelection()
            Me.DgvWUAllocate_WU.Item(CInt(c2), CInt(r2)).Selected = True

            Me.SetPlanedFinishDate()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgw2_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgw2.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgw2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgw2.RunWorkerCompleted
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'hitung ulang item modifikasi
            Me.ObjSch.HitungJadwalModifikasi2(Me.TBLSchBar, Me.IdentifierJO, Me.TBLItemWO, Me.TBLItemWO_ItemSch, _
            IdentifierAcuan, Me.DgvWUAllocate_Process(Me.ColAll_Process_TotalWU.Index, CInt(r1)).Value)

            'load dgv proses + dgv wu
            Me.LoadDGVProses(r1, c1)

            Me.SetPlanedFinishDate()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgw3_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgw3.DoWork
        Try
            Threading.Thread.Sleep(200)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgw3_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgw3.RunWorkerCompleted
        Dim Frm1 As New FrmTrans_ProductionSchedule_EditQty
        Dim rowDGV, rowDGV1 As DataGridViewRow
        Dim rowWUAcuan As DseTrans_Schedule.TblWU1Row

        Try
            Frm1.mDefPrec = Me.DgvWUAllocate_Process.Item(Me.ColAll_Process_Decimal.Index, CInt(r1)).Value
            Frm1.mJmlProduce = Me.DgvWUAllocate_Process.Item(Me.ColAll_Process_Qty.Index, CInt(r1)).Value
            Frm1.mPrimaryProduct = Me.DgvWUAllocate_Process.Item(Me.ColAll_Process_PrimaryCodeNo.Index, CInt(r1)).Value
            Frm1.mUnitCode = Me.DgvWUAllocate_Process.Item(Me.ColAll_UnitCode.Index, CInt(r1)).Value
            Frm1.mPrimaryPredefinedWU = Me.PrimaryPredefinedConfig

            For Each rowDGV In Me.DgvWUAllocate_WU.Rows
                Frm1.DgvWUAllocate_WU.Rows.Add()
                rowDGV1 = Frm1.DgvWUAllocate_WU.Rows(Frm1.DgvWUAllocate_WU.Rows.Count - 1)
                With rowDGV1
                    .Cells(Frm1.CoAll_WU_Allocate.Index).Value = rowDGV.Cells(Me.CoAll_WU_Allocate.Index).Value
                    .Cells(Frm1.CoAll_WU_Code.Index).Value = rowDGV.Cells(Me.CoAll_WU_Code.Index).Value
                    .Cells(Frm1.CoAll_WU_Desc.Index).Value = rowDGV.Cells(Me.CoAll_WU_Desc.Index).Value
                    .Cells(Frm1.CoAll_WU_Efficiency.Index).Value = rowDGV.Cells(Me.CoAll_WU_Efficiency.Index).Value
                    .Cells(Frm1.CoAll_WU_No.Index).Value = rowDGV.Cells(Me.CoAll_WU_No.Index).Value
                    .Cells(Frm1.CoAll_WU_PlanRate.Index).Value = rowDGV.Cells(Me.CoAll_WU_PlanRate.Index).Value
                    .Cells(Frm1.CoAll_WU_PrimaryWU.Index).Value = rowDGV.Cells(Me.CoAll_WU_PrimaryWU.Index).Value
                    .Cells(Frm1.CoAll_WU_ProdRate.Index).Value = rowDGV.Cells(Me.CoAll_WU_ProdRate.Index).Value
                    .Cells(Frm1.CoAll_WU_ProduceQty.Index).Value = rowDGV.Cells(Me.CoAll_WU_ProduceQty.Index).Value
                End With
                If rowDGV.Cells(Me.CoAll_WU_Allocate.Index).Value = Boolean.TrueString Then
                    Frm1.ListPrimaryWU.Add(rowDGV.Cells(Me.CoAll_WU_PrimaryWU.Index).Value)
                End If
            Next

            If Frm1.ShowDialog = Windows.Forms.DialogResult.OK Then
                Windows.Forms.Cursor.Current = Cursors.WaitCursor

                'isikan tabel wu acuan
                Me.TBLWUAcuan.Rows.Clear()
                For Each rowDGV In Frm1.DgvWUAllocate_WU.Rows
                    If rowDGV.Cells(Frm1.CoAll_WU_Allocate.Index).Value = Boolean.TrueString Then
                        rowWUAcuan = Me.TBLWUAcuan.NewTblWU1Row
                        With rowWUAcuan
                            .Primary_WU = rowDGV.Cells(Frm1.CoAll_WU_PrimaryWU.Index).Value
                            .Qty_Produce = rowDGV.Cells(Frm1.CoAll_WU_ProduceQty.Index).Value
                        End With
                        Me.TBLWUAcuan.Rows.Add(rowWUAcuan)
                    End If
                Next

                'hitung ulang item modifikasi
                Me.ObjSch.HitungJadwalModifikasi3(Me.TBLSchBar, Me.IdentifierJO, Me.TBLItemWO, Me.TBLItemWO_ItemSch, Me.IdentifierAcuan, Me.TBLWUAcuan, Me.PrimaryPredefinedConfig)

                'load dgv proses + dgv wu
                Me.LoadDGVProses(r1, c1)

                Me.SetPlanedFinishDate()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Frm1 = Nothing
        End Try
    End Sub

    Private Sub SetColor(ByVal vColor As Color)
        Try
            Me.pb_Sch.BackColor = vColor

            Me.TBLSchBar.DefaultView.RowFilter = "[identifier jo] = '" + Me.IdentifierJO + "'"
            For Each dv As Data.DataRowView In Me.TBLSchBar.DefaultView
                rowSch = dv.Row
                rowSch.a = vColor.A
                rowSch.r = vColor.R
                rowSch.g = vColor.G
                rowSch.b = vColor.B
            Next

            Me.tblData.DefaultView.RowFilter = "[identifier jo] = '" & Me.IdentifierJO & "'"
            If Me.tblData.DefaultView.Count > 0 Then
                For i As Integer = 0 To Me.tblData.DefaultView.Count - 1
                    Me.rowData = Me.tblData.DefaultView(i).Row
                    Me.rowData.a = vColor.A
                    Me.rowData.r = vColor.R
                    Me.rowData.g = vColor.G
                    Me.rowData.b = vColor.B
                    Dim drawBar As Boolean = False
                    If Me.rowData.Item("Is Header") = 1 Then
                        If Me.dgvBar.Rows(Me.tblData.Rows.IndexOf(Me.rowData) + 1).Visible = False Then
                            drawBar = True
                        End If
                    End If
                    Me.reDrawBarHeader(Me.rowData, drawBar)
                Next
            End If


            If Me.TabControl1.SelectedTab.Name = Me.TabPage2.Name Then
                Me.LlbColumnSettingSch.Visible = True
            Else
                Me.LlbColumnSettingSch.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnSet_SchDT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSetColor_Sch.Click
        Dim ij As New ColorDialog

        Try
            If ij.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.pb_Sch.BackColor = ij.Color

                Me.TBLSchBar.DefaultView.RowFilter = "[identifier jo] = '" + Me.IdentifierJO + "'"
                For Each dv As Data.DataRowView In Me.TBLSchBar.DefaultView
                    rowSch = dv.Row
                    rowSch.a = ij.Color.A
                    rowSch.r = ij.Color.R
                    rowSch.g = ij.Color.G
                    rowSch.b = ij.Color.B
                Next

                Me.tblData.DefaultView.RowFilter = "[identifier jo] = '" & Me.IdentifierJO & "'"
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


                If Me.TabControl1.SelectedTab.Name = Me.TabPage2.Name Then
                    Me.LlbColumnSettingSch.Visible = True
                Else
                    Me.LlbColumnSettingSch.Visible = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStop.Click
        Try
            If MsgBox("Stop adding new schedule and discard data changes?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                Me.Stopped = True
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvWUAllocate_WU_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvWUAllocate_WU.CellClick
        Dim MaxDate, MinDate As Date

        Try
            If e.ColumnIndex = Me.CoAll_WU_Detail.Index Then
                Dim ij As New FrmTrans_ProductionSchedule_NewSch2_Detail
                ij.UserCode = Me.UserCode
                ij.PrimaryWU = Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_PrimaryWU.Index, e.RowIndex).Value
                ij.TBLSchBar = Me.TBLSchBar
                ij.BackColorAll = Me.BackColorAll

                Me.TBLSchBar.DefaultView.RowFilter = "[identifier jo] = '" + Me.IdentifierJO + "'"
                Me.TBLSchBar.DefaultView.Sort = "[plan start time]"
                MinDate = Date.MaxValue
                MaxDate = Date.MinValue
                For Each DV As Data.DataRowView In Me.TBLSchBar.DefaultView
                    Me.rowSch = DV.Row
                    If rowSch.plan_start_time.Date < MinDate Then
                        MinDate = rowSch.plan_start_time.Date
                    End If
                    If rowSch.plan_end_time > MaxDate Then
                        MaxDate = rowSch.plan_end_time.Date
                    End If
                Next
                ij.dtpTo.Value = MaxDate
                ij.dtpFrom.Value = MinDate
                ij.IdentifierWO = Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_IdentifierWO.Index, e.RowIndex).Value
                ij.StartDate = CDate(Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_StartDate.Index, e.RowIndex).Value) + TimeSpan.FromMinutes(Me.ObjCD.Convert_ToMinutes(Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_StartTime.Index, e.RowIndex).Value))
                If Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_EndDate.Index, e.RowIndex).Value = "" Then
                    ij.EndDate = CDate(Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_StartDate.Index, e.RowIndex).Value) + TimeSpan.FromMinutes(Me.ObjCD.Convert_ToMinutes(Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_StartTime.Index, e.RowIndex).Value))
                Else
                    ij.EndDate = CDate(Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_EndDate.Index, e.RowIndex).Value) + TimeSpan.FromMinutes(Me.ObjCD.Convert_ToMinutes(Me.DgvWUAllocate_WU.Item(Me.CoAll_WU_EndTime.Index, e.RowIndex).Value))
                End If
                ij.ShowDialog()
                ij = Nothing
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SimpanPredefined()
        Dim primaryProduct As Double = -1
        Dim dtaHeader As New DseTrans_ProductionSchedule_NewSch3TableAdapters.DataTable1TableAdapter
        Dim dtaProcess As New DseTrans_ProductionSchedule_NewSch3TableAdapters.DataTable5TableAdapter
        Dim dtaWorkUnit As New DseTrans_ProductionSchedule_NewSch3TableAdapters.DataTable6TableAdapter
        Dim dtaProcessCurrentJO As New DseTrans_ProductionSchedule_NewSch3TableAdapters.DataTable4TableAdapter
        Dim tblProcessCurrentJO As New DseTrans_ProductionSchedule_NewSch3.DataTable4DataTable
        Dim dtaProcessOtherJO As New DseTrans_ProductionSchedule_NewSch3TableAdapters.DataTable2TableAdapter
        Dim tblProcessOtherJO As New DseTrans_ProductionSchedule_NewSch3.DataTable2DataTable
        Dim rowProcessOtherJO As DseTrans_ProductionSchedule_NewSch3.DataTable2Row
        Dim listOfPrimaryMasterExist As New List(Of Double)
        Dim listOfPrimaryMasterCorrect As New List(Of Double)

        Dim dtaConfig As New DseTrans_ProductionSchedule_NewSch3TableAdapters.DataTable1TableAdapter
        Dim valConfigCount As New Nullable(Of Long)
        Try
            primaryProduct = Me.ObjProduct.GetPrimary(Me.TxtCodeNo.Text)
            dtaConfig.Connection.ConnectionString = Core_Module.GetConstring
            valConfigCount = dtaConfig.ScalarQuery_GetJmlByPrimaryProductOut(primaryProduct)
            If valConfigCount.HasValue Then
                If valConfigCount.Value > 0 Then
                    If Me.PrimaryPredefinedConfig = -1 Then
                        Exit Sub
                    End If
                    dtaHeader.Connection.ConnectionString = Core_Module.GetConstring
                    dtaProcess.Connection.ConnectionString = Core_Module.GetConstring
                    dtaWorkUnit.Connection.ConnectionString = Core_Module.GetConstring
                    dtaProcessCurrentJO.Connection.ConnectionString = Core_Module.GetConstring
                    dtaProcessOtherJO.Connection.ConnectionString = Core_Module.GetConstring
                    tblHeader.Rows.Clear()
                    dtaProcess.Fill(tblProcess, primaryProduct)
                    dtaWorkUnit.Fill(tblWorkUnit, primaryProduct)
                    dtaProcessCurrentJO.Fill(tblProcessCurrentJO, Me.IdentifierJO)
                    dtaProcessOtherJO.Fill(tblProcessOtherJO, primaryProduct)

                    For Each rowProcessOtherJO In tblProcessOtherJO
                        If listOfPrimaryMasterExist.Contains(rowProcessOtherJO.primary_master) = False Then listOfPrimaryMasterExist.Add(rowProcessOtherJO.primary_master)
                    Next

                    For Each primaryMasterExist As Double In listOfPrimaryMasterExist
                        tblProcessOtherJO.DefaultView.RowFilter = "[primary master] = " & primaryMasterExist.ToString
                        If tblProcessOtherJO.DefaultView.Count = tblProcessCurrentJO.Count Then
                            Dim Correct As Boolean = True
                            For i As Integer = 0 To tblProcessOtherJO.DefaultView.Count - 1
                                rowProcessOtherJO = tblProcessOtherJO.DefaultView(i).Row
                                tblProcessCurrentJO.DefaultView.RowFilter = "[Primary Process] = " & rowProcessOtherJO.primary_process.ToString & " AND [Primary Product] = " & rowProcessOtherJO.primary_product.ToString
                                If tblProcessCurrentJO.DefaultView.Count = 0 Then Correct = False
                            Next
                            If Correct And listOfPrimaryMasterCorrect.Contains(primaryMasterExist) = False Then listOfPrimaryMasterCorrect.Add(primaryMasterExist)
                        End If
                    Next

                    dtaHeader.ClearBeforeFill = False
                    For Each primaryMasterCorrect As Double In listOfPrimaryMasterCorrect
                        dtaHeader.FillBy_Primary(tblHeader, primaryMasterCorrect)
                    Next
                    dtaHeader.ClearBeforeFill = True

                    If tblHeader.Rows.Count <= 0 Then
                        If MsgBox("Do you want to save work unit configuration for this product?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                            SimpanPredefined_()
                        End If
                    Else
                        If MsgBox("Do you want to update work unit configuration?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                            UdaptePredefined_()
                        End If
                    End If
                Else
                    If MsgBox("Do you want to save work unit configuration for this product?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                        SimpanPredefined_()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub SimpanPredefined_()
        Dim dtaSch As New DseTrans_ProductionSchedule_NewSch3TableAdapters.DataTable9TableAdapter
        Dim tblSch As New DseTrans_ProductionSchedule_NewSch3.DataTable9DataTable
        Dim rowSch As DseTrans_ProductionSchedule_NewSch3.DataTable9Row
        Dim f As New FrmTrans_ProductionSchedule_NewSch4
        Try
            dtaSch.Connection.ConnectionString = Core_Module.GetConstring
            dtaSch.Fill(tblSch, Me.IdentifierJO)
            f.PrimaryConfig = -1
            f.PrimaryProduct = Me.ObjProduct.GetPrimary(Me.TxtCodeNo.Text)
            f.UserCode = Me.UserCode
            f.BackColorAll = Me.BackColorAll
            f.Baru = True
            f.LoadFromSch = True
            f.TblWorkUnitSch = tblSch
            f.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub UdaptePredefined_()
        Dim dtaSch As New DseTrans_ProductionSchedule_NewSch3TableAdapters.DataTable9TableAdapter
        Dim tblSch As New DseTrans_ProductionSchedule_NewSch3.DataTable9DataTable
        Dim rowSch As DseTrans_ProductionSchedule_NewSch3.DataTable9Row
        Dim f As New FrmTrans_ProductionSchedule_NewSch4
        Try
            dtaSch.Connection.ConnectionString = Core_Module.GetConstring
            dtaSch.Fill(tblSch, Me.IdentifierJO)
            f.PrimaryConfig = Me.PrimaryPredefinedConfig
            f.PrimaryProduct = Me.ObjProduct.GetPrimary(Me.TxtCodeNo.Text)
            f.UserCode = Me.UserCode
            f.BackColorAll = Me.BackColorAll
            f.Baru = False
            f.LoadFromSch = True
            f.TblWorkUnitSch = tblSch
            f.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Sub ResizeMainPanel()
    '    Try
    '        Me.SplitContainer3.SplitterDistance = Me.GetVisibleTGVWidth
    '        Me.dgvBar.Width = Me.SplitContainer3.Panel2.Width
    '        Me.dgvHeader1.Width = Me.SplitContainer3.Panel2.Width
    '        Me.dgvHeader0.Width = Me.SplitContainer3.Panel2.Width
    '        Me.SplitContainer3.Refresh()
    '        Me.dgvSch.Left = 0
    '        Me.dgvSch.Width = Me.SplitContainer3.SplitterDistance
    '        Me.dgvSch.Refresh()
    '        'Me.llbCollapsAll.Left = Me.SplitContainer1.Panel1.Width - 5 - Me.llbCollapsAll.Width
    '        'Me.llbExpandAll.Left = Me.llbCollapsAll.Left - 5 - Me.llbExpandAll.Width
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    'Function GetVisibleTGVWidth() As Integer
    '    Dim GetVisibleTGVWidth_ As Integer = 0
    '    Try
    '        For i As Integer = 0 To Me.dgvSch.ColumnCount - 1
    '            If Me.dgvSch.Columns(i).Visible Then
    '                GetVisibleTGVWidth_ += Me.dgvSch.Columns(i).Width
    '            End If
    '        Next

    '        If Me.dgvSch.DisplayedRowCount(False) <> Me.dgvSch.RowCount Then
    '            GetVisibleTGVWidth_ += 20
    '        End If
    '        Return GetVisibleTGVWidth_
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '        Return 0
    '    End Try
    'End Function

    Private Sub FrmTrans_ProductionSchedule_NewSch2_LocationChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.LocationChanged

    End Sub

    'Private Sub FrmTrans_ProductionSchedule_NewSch2_ResizeEnd(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.ResizeEnd
    '    Try
    '        Me.ResizeMainPanel()
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    'Private Sub FrmTrans_ProductionSchedule_NewSch2_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
    '    Try
    '        Me.ResizeMainPanel()
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    'Private Sub FrmTrans_ProductionSchedule_NewSch2_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.SizeChanged
    '    Try
    '        Me.completeRowBar()
    '        Me.ResizeMainPanel()
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Private Sub bgw4_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgw4.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgw4_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgw4.RunWorkerCompleted
        Dim FirstIndex As Integer = 0
        Try
            FirstIndex = Me.DgvWUAllocate_Process.FirstDisplayedCell.RowIndex

            Me.DgvWUAllocate_Process.ClearSelection()
            Me.LoadDGVProses(Me.r1, Me.c1)

            While Me.DgvWUAllocate_Process.FirstDisplayedCell.RowIndex <> FirstIndex
                If FirstIndex < Me.DgvWUAllocate_Process.FirstDisplayedCell.RowIndex Then
                    Me.DgvWUAllocate_Process.FirstDisplayedScrollingRowIndex = Me.DgvWUAllocate_Process.FirstDisplayedScrollingRowIndex - 1
                    If Me.DgvWUAllocate_Process.FirstDisplayedScrollingRowIndex = 0 Then
                        Exit While
                    End If
                ElseIf FirstIndex > Me.DgvWUAllocate_Process.FirstDisplayedCell.RowIndex Then
                    Me.DgvWUAllocate_Process.FirstDisplayedScrollingRowIndex = Me.DgvWUAllocate_Process.FirstDisplayedScrollingRowIndex + 1
                    If Me.DgvWUAllocate_Process.FirstDisplayedScrollingRowIndex = Me.DgvWUAllocate_Process.Rows.Count - 1 Then
                        Exit While
                    End If
                End If
            End While
            Me.SetPlanedFinishDate()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgw5_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgw5.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgw5_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgw5.RunWorkerCompleted
        Dim FirstIndex As Integer = 0

        Try
            FirstIndex = Me.DgvWUAllocate_Process.FirstDisplayedCell.RowIndex

            Me.DgvWUAllocate_Process.ClearSelection()
            Me.LoadDGVProses(Me.r1, Me.c1)
            Me.DgvWUAllocate_WU.ClearSelection()
            Me.DgvWUAllocate_WU.Item(CInt(Me.c2), CInt(Me.r2)).Selected = True

            While Me.DgvWUAllocate_Process.FirstDisplayedCell.RowIndex <> FirstIndex
                Me.DgvWUAllocate_Process.FirstDisplayedScrollingRowIndex = Me.DgvWUAllocate_Process.FirstDisplayedScrollingRowIndex + 1
            End While
            Me.SetPlanedFinishDate()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbSetDependency_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbSetDependency.LinkClicked
        Dim ij As New FrmTrans_ProductionSchedule_NewSch5
        Dim rowDGV As DataGridViewRow
        Dim ListPrimaryProcess As New List(Of Double)
        Dim rowRef As DseTrans_Schedule.TblDependencyRow
        Dim PrimaryProses As Double = -1
        Dim DummyStartTime As Date = Date.MaxValue
        Dim RowItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1Row
        Dim TipeDependencyInt As Integer = 0
        Dim DelayDependency As Double = 0
        Dim RowItemWO_ItemSch As DseTrans_Schedule.tbltrans_mo_item_wo_item_schRow
        Dim StartEvaluate As Boolean = False

        Try
            'isi tblref 
            For Each rowDGV In Me.DgvWUAllocate_Process.Rows
                If ListPrimaryProcess.Contains(rowDGV.Cells(Me.ColAll_Process_PrimaryProcess.Index).Value) = False Then
                    ListPrimaryProcess.Add(rowDGV.Cells(Me.ColAll_Process_PrimaryProcess.Index).Value)
                    rowRef = ij.TblRef.NewTblDependencyRow
                    rowRef.Delay_Dependency = Me.ObjCD.Convert_ToMinutes(rowDGV.Cells(Me.ColAll_Process_DelayDependency.Index).Value)
                    rowRef.Tipe_Dependency = Me.ObjProduct.GetTipeDependency(rowDGV.Cells(Me.ColAll_Process_TipeDependency.Index).Value)
                    rowRef.Primary_Process = rowDGV.Cells(Me.ColAll_Process_PrimaryProcess.Index).Value
                    rowRef.Process_Code = rowDGV.Cells(Me.ColAll_Process_Code.Index).Value

                    ij.TblRef.Rows.Add(rowRef)
                End If
            Next

            If ij.ShowDialog() = Windows.Forms.DialogResult.OK Then
                TipeDependencyInt = ij.cboDependency.SelectedIndex
                DelayDependency = Me.ObjCD.Convert_ToMinutes(ij.TxtDelay.Text)

                For Each rowRef In ij.TblRef
                    If rowRef.Process_Code = ij.cboProcess.Text Then
                        'updte ke dgv
                        Me.TBLSchBar.DefaultView.RowFilter = "[identifier jo] = '" + Me.IdentifierJO + "'" '+ "' and [identifier wo] ='" + Me.DgvWUAllocate_Process.Item(Me.ColAll_Process_IdentifierWO.Index, e.RowIndex).Value.ToString + "'"

                        For Each dv As Data.DataRowView In Me.TBLSchBar.DefaultView
                            rowSch = dv.Row
                            'ubah data tipe dependency
                            For Each RowItemWO In Me.TBLItemWO.Rows
                                If RowItemWO.Primary_Process = rowRef.Primary_Process Then
                                    StartEvaluate = True

                                    RowItemWO.tipe_dependency = TipeDependencyInt
                                    RowItemWO.delay_dependency = DelayDependency

                                    For Each RowItemWO_ItemSch In Me.TBLItemWO_ItemSch.Rows
                                        If RowItemWO_ItemSch.identifier_wo = RowItemWO.Identifier And _
                                        RowItemWO_ItemSch.identifier_jo = RowItemWO.Identifier_JO Then
                                            RowItemWO_ItemSch.tipe_dependency = TipeDependencyInt
                                            RowItemWO_ItemSch.delay_dependency = DelayDependency
                                        End If
                                    Next
                                End If

                                If StartEvaluate = True Then
                                    If DummyStartTime > RowItemWO.Plan_Start_Time Then
                                        DummyStartTime = RowItemWO.Plan_Start_Time
                                    End If
                                End If
                            Next

                            'hitung ulang jadwal baru
                            Me.ObjSch.Reschedule_JadwalBaru1(Me.IdentifierJO, Me.TBLSchBar, Me.TBLItemWO, Me.TBLItemWO_ItemSch, rowSch.a, rowSch.r, rowSch.g, _
                            rowSch.b, Me.TBLSchBarKosong, Me.TBLSchBarTanpaWO, DummyStartTime, Me.PrimaryPredefinedConfig)

                            If Me.DgvWUAllocate_Process.SelectedCells.Count > 0 Then
                                Me.r1 = Me.DgvWUAllocate_Process.SelectedCells(0).RowIndex
                                Me.c1 = Me.DgvWUAllocate_Process.SelectedCells(0).ColumnIndex
                            Else
                                Me.r1 = 0
                                Me.c1 = 0
                            End If

                            'MsgBox(Me.ObjCD.DigitGrouping((Now - Start).TotalMinutes, 1))
                            Me.bgw4.RunWorkerAsync()

                            Exit Sub
                        Next

                        Exit Sub
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
            ListPrimaryProcess = Nothing
        End Try
    End Sub

    Private Sub llbOptimizePlan_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbOptimizePlan.LinkClicked
        Dim rowDGV, rowDGV1 As DataGridViewRow
        Dim PrimaryProses As Double = -1
        Dim StartProses As Boolean = False
        Dim rowItemWO As DseTrans_Schedule.tbltrans_mo_item_wo1Row
        Dim rowItemWOSch As DseTrans_Schedule.tbltrans_mo_item_wo_item_schRow
        Dim rowItemSchBar, rowItemSchBar1 As DseTrans_Schedule.tbl_sch_barRow
        Dim TipeSchBar As ClsTrans_Schedule.EnumTipeSchBar = ClsTrans_Schedule.EnumTipeSchBar.work_time
        Dim DurasiGeser As Double = 0
        Dim TBLSchBarKosongDummy As New DseTrans_Schedule.tbl_sch_barDataTable
        Dim PrimaryWUDummy As Double = -1
        Dim TipeDependency As Core_Module.clsMasterProduct.EnumTipeDependency = Core_Module.clsMasterProduct.EnumTipeDependency.End_Start1
        Dim DelayDependency As Double = 0
        Dim StartTime, EndTime As Date

        Try
            If Me.DgvWUAllocate_Process.SelectedCells.Count > 0 Then
                rowDGV = Me.DgvWUAllocate_Process.Rows(Me.DgvWUAllocate_Process.SelectedCells(0).RowIndex)

                If MsgBox("Optimize schedule for the following process: " + rowDGV.Cells(Me.ColAll_Process_Code.Index).Value.ToString + "?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    'untuk setiap item wo, dari proses yang dirujuk, mulai dari item wo pertama
                    PrimaryProses = rowDGV.Cells(Me.ColAll_Process_PrimaryProcess.Index).Value
                    Me.TBLSchBar.DefaultView.Sort = "[plan start time]"

                    For Each rowDGV In Me.DgvWUAllocate_Process.Rows
                        If PrimaryProses = rowDGV.Cells(Me.ColAll_Process_PrimaryProcess.Index).Value Then
                            If StartProses = True Then
                                Me.TBLItemWO.DefaultView.RowFilter = "[Identifier JO] = '" + Me.IdentifierJO + "' AND Identifier = '" + rowDGV.Cells(Me.ColAll_Process_IdentifierWO.Index).Value.ToString + "'"
                                rowItemWO = Me.TBLItemWO.DefaultView(0).Row
                                StartTime = Date.MaxValue
                                EndTime = Date.MinValue

                                Me.TBLItemWO_ItemSch.DefaultView.RowFilter = "[identifier jo] = '" + Me.IdentifierJO + "' AND [identifier wo] = '" + rowDGV.Cells(Me.ColAll_Process_IdentifierWO.Index).Value.ToString + "'"

                                'geser baris kedua dan selanjutnya, yang harus diubah: plan start, plan end, delay dependency berdasarkan jadwal akhir dari work unit ybs
                                For Each DV As Data.DataRowView In Me.TBLItemWO_ItemSch.DefaultView
                                    rowItemWOSch = DV.Row

                                    If PrimaryWUDummy <> rowItemWOSch.primary_work_unit Then
                                        PrimaryWUDummy = rowItemWOSch.primary_work_unit
                                        TBLSchBarKosongDummy.Rows.Clear()
                                        TBLSchBarKosong.DefaultView.RowFilter = "[primary work unit] = " + rowItemWOSch.primary_work_unit.ToString
                                        For Each DV1 As Data.DataRowView In TBLSchBarKosong.DefaultView
                                            rowItemSchBar = DV1.Row
                                            rowItemSchBar1 = TBLSchBarKosongDummy.Newtbl_sch_barRow
                                            For jk As Integer = 0 To TBLSchBarKosongDummy.Columns.Count - 1
                                                rowItemSchBar1.Item(jk) = rowItemSchBar.Item(jk)
                                            Next
                                            TBLSchBarKosongDummy.Rows.Add(rowItemSchBar1)
                                        Next
                                    End If

                                    'copy item sch bar kosong tanpa wo, khusus untuk 
                                    rowDGV1 = Me.DgvWUAllocate_Process.Rows(rowDGV.Index - 1)

                                    Me.TBLSchBar.DefaultView.RowFilter = "[primary work unit] = " + rowItemWOSch.primary_work_unit.ToString + " AND tipe = " + CInt(TipeSchBar).ToString + _
                                    " AND [identifier jo] = '" + Me.IdentifierJO + "' AND [identifier wo] = '" + rowDGV.Cells(Me.ColAll_Process_IdentifierWO.Index).Value.ToString + "'"
                                    rowItemSchBar = Me.TBLSchBar.DefaultView(0).Row
                                    Me.TBLSchBar.DefaultView.RowFilter = "[primary work unit] = " + rowItemWOSch.primary_work_unit.ToString + " AND tipe = " + CInt(TipeSchBar).ToString + _
                                    " AND [identifier jo] = '" + Me.IdentifierJO + "' AND [identifier wo] = '" + rowDGV1.Cells(Me.ColAll_Process_IdentifierWO.Index).Value.ToString + "'"
                                    rowItemSchBar1 = Me.TBLSchBar.DefaultView(0).Row

                                    If rowItemSchBar1.plan_end_time < rowItemSchBar.plan_start_time Then
                                        'geser kiri
                                        DurasiGeser = Me.ObjSch.GetWorkDuration(rowItemSchBar1.plan_end_time, rowItemSchBar.plan_start_time, rowItemSchBar.primary_work_unit, TBLSchBarKosongDummy)
                                        rowItemSchBar.plan_start_time = Me.ObjSch.GetStartTime(rowItemSchBar.primary_work_unit, rowItemSchBar.plan_start_time, DurasiGeser, TBLSchBarKosongDummy)
                                    Else
                                        'geser kanan
                                        DurasiGeser = Me.ObjSch.GetWorkDuration(rowItemSchBar.plan_start_time, rowItemSchBar1.plan_end_time, rowItemSchBar.primary_work_unit, TBLSchBarKosongDummy)
                                        rowItemSchBar.plan_start_time = Me.ObjSch.GetEndTime(rowItemSchBar.primary_work_unit, rowItemSchBar.plan_start_time, DurasiGeser, TBLSchBarKosongDummy)
                                    End If
                                    'rowItemSchBar.plan_end_time = Me.ObjSch.GetEndTime(rowItemSchBar.primary_work_unit, rowItemSchBar.plan_start_time, rowItemSchBar.total_work_hour, TBLSchBarKosongDummy)

                                    If DurasiGeser > 0 Then
                                        'rowItemSchBar.plan_start_time = Me.ObjSch.GetStartTime(rowItemSchBar.primary_work_unit, rowItemSchBar.plan_start_time, DurasiGeser, TBLSchBarKosongDummy)
                                        rowItemSchBar.plan_end_time = Me.ObjSch.GetEndTime(rowItemSchBar.primary_work_unit, rowItemSchBar.plan_start_time, rowItemSchBar.total_work_hour, TBLSchBarKosongDummy)

                                        TipeDependency = rowItemWOSch.tipe_dependency
                                        rowItemWOSch.plan_start_time = rowItemSchBar.plan_start_time
                                        rowItemWOSch.plan_end_time = rowItemSchBar.plan_end_time
                                        rowItemWOSch.plan_end_time_minimum = rowItemWOSch.plan_end_time
                                        rowItemWOSch.delay_dependency = Me.GetDelayDependency(rowItemWO, rowItemWOSch.plan_start_time, rowItemWOSch.primary_work_unit)
                                        DelayDependency = rowItemWOSch.delay_dependency

                                        If StartTime > rowItemSchBar.plan_start_time Then
                                            StartTime = rowItemSchBar.plan_start_time
                                        End If
                                        If EndTime < rowItemSchBar.plan_end_time Then
                                            EndTime = rowItemSchBar.plan_end_time
                                        End If

                                    End If
                                Next

                                'ubah plan start dan plan end untuk setiap item wo
                                rowItemWO.Plan_Start_Time = StartTime
                                rowItemWO.Plan_End_Time = EndTime
                                'jika item wo ybs, hanya terdiri 1 wu, maka ubah juga tipe dependency dan delay dependency, disamakan dengan item sch wo
                                If Me.TBLItemWO_ItemSch.DefaultView.Count = 1 Then
                                    rowItemWO.tipe_dependency = CInt(TipeDependency)
                                    rowItemWO.delay_dependency = DelayDependency
                                End If
                            Else
                                'baris pertama digunakan sebagai acuan
                                StartProses = True
                            End If
                        Else
                            If StartProses = True Then
                                'semua baris dgv atas proses tsb sudah di update, keluar dari loop
                                Exit For
                            End If
                        End If
                    Next

                    If Me.DgvWUAllocate_Process.SelectedCells.Count > 0 Then
                        Me.r1 = Me.DgvWUAllocate_Process.SelectedCells(0).RowIndex
                        Me.c1 = Me.DgvWUAllocate_Process.SelectedCells(0).ColumnIndex
                    Else
                        Me.r1 = 0
                        Me.c1 = 0
                    End If

                    Me.bgw4.RunWorkerAsync()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TBLSchBarKosongDummy.Dispose()
        End Try
    End Sub


#Region "rosadin-r5b"

    Dim tblData As New DseTrans_Schedule.tbl_sch_barDataTable
    Dim rowData As DseTrans_Schedule.tbl_sch_barRow
    Dim programmaticExpandCollaps As Boolean = False
    Dim ExpandAll As Boolean = False
    Dim objProcess As New Core_Module.ClsMasterProses
    Dim objMO As New ClsTrans_ManufacturingOrder
    Dim objUnit As New Core_Module.ClsMasterUnit

    Sub reloadGanttChart()
        Try
            Me.dgvBar.RowTemplate.Height = Me.tgvSch.RowTemplate.Height
            Me.Cursor = Cursors.WaitCursor
            Me.translateData()
            Me.fillUnitData(Me.tblData)
            Me.filltgvSch()
            Me.generateBarColumns(Me.cboTimeframe.SelectedIndex, Me.dtpBarFrom.Value)
            'Me.SetSomeColumnVisibilty()
            Me.CollapsAll()
            Me.tgvSch.SelectionMode = DataGridViewSelectionMode.CellSelect
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Sub SetSomeColumnVisibilty()
        Try
            Me.ColLeft_ProcessCode.Visible = False
            Me.ColLeft_WONo.Visible = False
            Me.ColLeft_CodeNo.Visible = False
            Me.ColLeft_Description.Visible = False
            Me.ColLeft_QtyProduce.Visible = False
            Me.ColLeft_Unit.Visible = False
            Me.ColLeft_TotalWH.Visible = False
            Me.ColLeft_WorkUnit.Visible = False
            Me.ColLeft_ProductionRate.Visible = False
            Me.ColLeft_InputLocation.Visible = False
            Me.ColLeft_OutputLocation.Visible = False

            For Each row As DataGridViewRow In Me.tgvSch.Rows
                If row.Visible = True Then
                    If Not String.IsNullOrEmpty(row.Cells(Me.ColLeft_ProcessCode.Index).Value) Then
                        If ShowProcess = "1" Then Me.ColLeft_ProcessCode.Visible = True
                    End If
                    If Not String.IsNullOrEmpty(row.Cells(Me.ColLeft_WONo.Index).Value) Then
                        If ShowWO = "1" Then Me.ColLeft_WONo.Visible = True
                    End If
                    If Not String.IsNullOrEmpty(row.Cells(Me.ColLeft_CodeNo.Index).Value) Then
                        If ShowCodeNo = "1" Then Me.ColLeft_CodeNo.Visible = True
                    End If
                    If Not String.IsNullOrEmpty(row.Cells(Me.ColLeft_Description.Index).Value) Then
                        If ShowDesc = "1" Then Me.ColLeft_Description.Visible = True
                    End If
                    If Not String.IsNullOrEmpty(row.Cells(Me.ColLeft_QtyProduce.Index).Value) Then
                        If ShowQty = "1" Then Me.ColLeft_QtyProduce.Visible = True
                    End If
                    If Not String.IsNullOrEmpty(row.Cells(Me.ColLeft_Unit.Index).Value) Then
                        If ShowUnit = "1" Then Me.ColLeft_Unit.Visible = True
                    End If
                    If Not String.IsNullOrEmpty(row.Cells(Me.ColLeft_TotalWH.Index).Value) Then
                        If ShowWH = "1" Then Me.ColLeft_TotalWH.Visible = True
                    End If
                    If Not String.IsNullOrEmpty(row.Cells(Me.ColLeft_WorkUnit.Index).Value) Then
                        Me.ColLeft_WorkUnit.Visible = True
                    End If
                    If Not String.IsNullOrEmpty(row.Cells(Me.ColLeft_ProductionRate.Index).Value) Then
                        If ShowRate = "1" Then Me.ColLeft_ProductionRate.Visible = True
                    End If

                    If Not String.IsNullOrEmpty(row.Cells(Me.ColLeft_InputLocation.Index).Value) Then
                        If ShowInputLocation = "1" Then Me.ColLeft_InputLocation.Visible = True
                    End If
                    If Not String.IsNullOrEmpty(row.Cells(Me.ColLeft_OutputLocation.Index).Value) Then
                        If ShowOutputLocation = "1" Then Me.ColLeft_OutputLocation.Visible = True
                    End If

                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Sub SetSomeColumnVisibilty()
    '    Try
    '        Me.ColLeft_WONo.Visible = False
    '        Me.ColLeft_WorkUnit.Visible = False
    '        Me.ColLeft_CodeNo.Visible = False
    '        Me.ColLeft_Description.Visible = False
    '        Me.ColLeft_QtyProduce.Visible = False
    '        Me.ColLeft_Unit.Visible = False
    '        Me.ColLeft_TotalWH.Visible = False
    '        Me.ColLeft_ProductionRate.Visible = False
    '        For Each row As DataGridViewRow In Me.tgvSch.Rows
    '            If row.Visible = True Then
    '                If Not String.IsNullOrEmpty(row.Cells(Me.ColLeft_WONo.Index).Value) Then
    '                    If Me.ShowWO = "1" Then Me.ColLeft_WONo.Visible = True
    '                End If
    '                If Not String.IsNullOrEmpty(row.Cells(Me.ColLeft_WorkUnit.Index).Value) Then
    '                    Me.ColLeft_WorkUnit.Visible = True
    '                End If
    '                If Not String.IsNullOrEmpty(row.Cells(Me.ColLeft_CodeNo.Index).Value) Then
    '                    If ShowCodeNo = "1" Then Me.ColLeft_CodeNo.Visible = True
    '                End If
    '                If Not String.IsNullOrEmpty(row.Cells(Me.ColLeft_Description.Index).Value) Then
    '                    If ShowDesc = "1" Then Me.ColLeft_Description.Visible = True
    '                End If
    '                If Not String.IsNullOrEmpty(row.Cells(Me.ColLeft_QtyProduce.Index).Value) Then
    '                    Me.ColLeft_QtyProduce.Visible = True
    '                End If
    '                If Not String.IsNullOrEmpty(row.Cells(Me.ColLeft_Unit.Index).Value) Then
    '                    Me.ColLeft_Unit.Visible = True
    '                End If
    '                If Not String.IsNullOrEmpty(row.Cells(Me.ColLeft_TotalWH.Index).Value) Then
    '                    Me.ColLeft_TotalWH.Visible = True
    '                End If
    '                If Not String.IsNullOrEmpty(row.Cells(Me.ColLeft_ProductionRate.Index).Value) Then
    '                    Me.ColLeft_ProductionRate.Visible = True
    '                End If
    '            End If
    '        Next
    '        Me.ResizeMainPanel()
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Sub CollapsAll()
        Try
            Me.ExpandAll = True
            Me.tgvSch.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For Each node As AdvancedDataGridView.TreeGridNode In Me.tgvSch.Nodes
                For Each node1 As AdvancedDataGridView.TreeGridNode In node.Nodes
                    If node1.RowIndex > -1 Then node1.Collapse()
                Next
                If node.RowIndex > -1 Then node.Collapse()
            Next
            Me.SetSomeColumnVisibilty()
            Me.tgvSch.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.ResizeMainPanel()
            Me.ExpandAll = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Function isOnAlternateColumn(ByVal i As Integer) As Boolean
        Try
            If (i >= 1 And i <= 14) Or (i >= 29 And i <= 42) Or (i >= 57 And i <= 70) Or _
                    (i >= 85 And i <= 98) Or (i >= 113 And i <= 126) Or (i >= 141 And i <= 154) Then Return True
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Sub drawBar(ByVal vIntervalScale As ClsTrans_Schedule.enumDurationScale, ByVal vFrom As Date)
        Dim intervalPerBar As TimeSpan
        Dim barValueFrom, barValueTo As Date
        Dim dr As DataGridViewRow
        Try
            Me.dgvBar.Rows.Clear()
            If vIntervalScale = ClsTrans_Schedule.enumDurationScale.By_Hour Then
                intervalPerBar = intervalPerBar_ByHour
            ElseIf vIntervalScale = ClsTrans_Schedule.enumDurationScale.By_Day Then
                intervalPerBar = intervalPerBar_ByDay
            ElseIf vIntervalScale = ClsTrans_Schedule.enumDurationScale.By_Weeks Then
                intervalPerBar = intervalPerBar_ByWeek
            End If
            If Not Me.tblData.Columns.Contains("bar index") Then
                Me.tblData.Columns.Add("bar index", System.Type.GetType("System.Double"))
                Me.tblData.Columns("bar index").DefaultValue = "-1"
            End If
            For Each Me.rowData In Me.tblData
                With Me.rowData
                    Me.rowData("bar index") = Me.tblData.Rows.IndexOf(Me.rowData)
                    Me.dgvBar.Rows.Add()
                    dr = Me.dgvBar.Rows(Me.dgvBar.Rows.Count - 1)
                    barValueFrom = vFrom
                    barValueTo = vFrom + intervalPerBar
                    For i As Integer = 1 To Me.dgvBar.Columns.Count
                        If Me.rowData("Is Header") = 0 And ((barValueFrom <= .plan_start_time And barValueTo > .plan_start_time) Or _
                                (barValueFrom >= .plan_start_time And barValueTo <= .plan_end_time) Or _
                                (barValueFrom < .plan_end_time And barValueTo >= .plan_end_time)) Then
                            dr.Cells("ColBar" & i).Style.BackColor = System.Drawing.Color.FromArgb(.a, .r, .g, .b)
                        Else
                            If Me.isOnAlternateColumn(i) Then
                                dr.Cells("ColBar" & i).Style.BackColor = Me.AlternateBarColor
                            Else
                                dr.Cells("ColBar" & i).Style.BackColor = Color.White
                            End If
                        End If
                        barValueFrom += intervalPerBar
                        barValueTo += intervalPerBar
                    Next
                End With
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub generateBarColumns(ByVal vIntervalScale As ClsTrans_Schedule.enumDurationScale, ByVal vFrom As Date)
        Dim barColumnCount As Integer = 0
        Dim header0ColumnCount As Integer = 0
        Dim header1ColumnCount As Integer = 0
        Dim dgvBarWidth As Integer = 0
        Dim dgvHeader0Width As Integer = 0
        Dim dgvHeader1Width As Integer = 0
        Dim dc As DataGridViewColumn

        Dim durationHeader0 As Date = vFrom
        Dim incrementHeader0 As TimeSpan
        Dim formatDurationHeader0 As String = ""

        Dim durationHeader1 As Date = vFrom
        Dim incrementHeader1 As TimeSpan
        Dim formatDurationHeader1 As String = ""

        Try
            If vIntervalScale = ClsTrans_Schedule.enumDurationScale.By_Hour Then
                barColumnCount = 288
                header0ColumnCount = 24
                header1ColumnCount = 1
                dgvBarWidth = 10
                dgvHeader0Width = dgvBarWidth * 12
                dgvHeader1Width = dgvHeader0Width * 24
                incrementHeader0 = TimeSpan.FromHours(1)
                formatDurationHeader0 = "HH.mm"
                incrementHeader1 = TimeSpan.FromDays(1)
                formatDurationHeader1 = "dd-MMM-yy"
            ElseIf vIntervalScale = ClsTrans_Schedule.enumDurationScale.By_Day Then
                barColumnCount = 336
                header0ColumnCount = 14
                header1ColumnCount = 1
                dgvBarWidth = 10
                dgvHeader0Width = dgvBarWidth * 24
                dgvHeader1Width = dgvHeader0Width * 14
                incrementHeader0 = TimeSpan.FromHours(12)
                formatDurationHeader0 = "dd-MMM-yy"
            ElseIf vIntervalScale = ClsTrans_Schedule.enumDurationScale.By_Weeks Then
                barColumnCount = 168
                header0ColumnCount = barColumnCount / 2
                header1ColumnCount = header0ColumnCount / 7
                dgvBarWidth = 10
                dgvHeader0Width = dgvBarWidth * 2
                dgvHeader1Width = dgvHeader0Width * 7
                incrementHeader0 = TimeSpan.FromDays(1)
                incrementHeader1 = TimeSpan.FromDays(7)
                formatDurationHeader1 = "dd-MMM-yy"
            End If

            Me.dgvBar.Columns.Clear()
            For i As Integer = 1 To barColumnCount
                Dim c As New DataGridViewImageColumn
                c.Name = "ColBar" & i
                c.DefaultCellStyle.NullValue = Nothing
                'dc = Me.dgvBar.Columns(Me.dgvBar.Columns.Add("ColBar" & i, ""))
                dc = Me.dgvBar.Columns(Me.dgvBar.Columns.Add(c))
                dc.Width = dgvBarWidth
            Next

            Me.dgvHeader0.Columns.Clear()
            For i As Integer = 1 To header0ColumnCount
                If vIntervalScale <> ClsTrans_Schedule.enumDurationScale.By_Weeks Then
                    dc = Me.dgvHeader0.Columns(Me.dgvHeader0.Columns.Add("ColHeader0" & i, Format(durationHeader0, formatDurationHeader0)))
                Else
                    dc = Me.dgvHeader0.Columns(Me.dgvHeader0.Columns.Add("ColHeader0" & i, Me.ObjCD.GetHari_English(durationHeader0).Substring(0, 1)))
                End If
                durationHeader0 = durationHeader0.Add(incrementHeader0)
                dc.Width = dgvHeader0Width
                dc.SortMode = DataGridViewColumnSortMode.NotSortable
                dc.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            Next

            Me.dgvHeader1.Columns.Clear()
            For i As Integer = 1 To header1ColumnCount
                If vIntervalScale = ClsTrans_Schedule.enumDurationScale.By_Day Then
                    dc = Me.dgvHeader1.Columns(Me.dgvHeader1.Columns.Add("ColHeader1" & i, ""))
                ElseIf vIntervalScale = ClsTrans_Schedule.enumDurationScale.By_Hour Then
                    dc = Me.dgvHeader1.Columns(Me.dgvHeader1.Columns.Add("ColHeader1" & i, Format(durationHeader1, formatDurationHeader1)))
                ElseIf vIntervalScale = ClsTrans_Schedule.enumDurationScale.By_Weeks Then
                    If durationHeader1.Year <> durationHeader1.Add(incrementHeader1).Year Then
                        dc = Me.dgvHeader1.Columns(Me.dgvHeader1.Columns.Add("ColHeader1" & i, Format(durationHeader1, "dd MMM yy") & " - " & Format(durationHeader1.Add(incrementHeader1), "dd MMM yyyy")))
                    ElseIf durationHeader1.Month <> durationHeader1.Add(incrementHeader1).Month Then
                        dc = Me.dgvHeader1.Columns(Me.dgvHeader1.Columns.Add("ColHeader1" & i, Format(durationHeader1, "dd MMM") & " - " & Format(durationHeader1.Add(incrementHeader1), "dd MMM yyyy")))
                    Else
                        dc = Me.dgvHeader1.Columns(Me.dgvHeader1.Columns.Add("ColHeader1" & i, Format(durationHeader1, "dd") & " - " & Format(durationHeader1.Add(incrementHeader1), "dd MMM yyyy")))
                    End If
                End If
                durationHeader1 = durationHeader1.Add(incrementHeader1)
                dc.Width = dgvHeader1Width
                dc.SortMode = DataGridViewColumnSortMode.NotSortable
                dc.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            Next

            Me.drawBar(vIntervalScale, vFrom)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Function GetVisibleTGVWidth() As Integer
        Dim GetVisibleTGVWidth_ As Integer = 0
        Try
            For i As Integer = 0 To Me.tgvSch.ColumnCount - 1
                If Me.tgvSch.Columns(i).Visible Then
                    GetVisibleTGVWidth_ += Me.tgvSch.Columns(i).Width
                End If
            Next
            If Me.tgvSch.DisplayedRowCount(False) <> Me.tgvSch.Rows.Count Then
                GetVisibleTGVWidth_ += 20
            End If
            Return GetVisibleTGVWidth_
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        End Try
    End Function

    Sub ResizeMainPanel()
        Try
            Me.SplitContainer3.SplitterDistance = Me.GetVisibleTGVWidth
            Me.dgvBar.Width = Me.SplitContainer3.Panel2.Width
            Me.dgvHeader1.Width = Me.SplitContainer3.Panel2.Width
            Me.dgvHeader0.Width = Me.SplitContainer3.Panel2.Width
            Me.tgvSch.Width = Me.SplitContainer3.SplitterDistance
            Me.llbCollapsAll.Left = Me.SplitContainer3.Panel1.Width - 5 - Me.llbCollapsAll.Width
            Me.llbExpandAll.Left = Me.llbCollapsAll.Left - 5 - Me.llbExpandAll.Width
            Me.Refresh()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Dim ShowWO, ShowCodeNo, ShowDesc, ShowDayStart, ShowDayFinish, ShowProcess, ShowRate, ShowQty, ShowUnit, ShowWH, ShowInputLocation, ShowOutputLocation As String
    Sub setColumnVisibility()
        Try
            ShowWO = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_ShowWO, Me.UserCode)
            If ShowWO = "" Then ShowWO = "1"
            ShowCodeNo = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_ShowCodeNo, Me.UserCode)
            If ShowCodeNo = "" Then ShowCodeNo = "1"
            ShowDesc = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_ShowDesc, Me.UserCode)
            If ShowDesc = "" Then ShowDesc = "1"
            ShowDayStart = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_ShowDayStart, Me.UserCode)
            If ShowDayStart = "" Then ShowDayStart = "1"
            ShowDayFinish = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_ShowDayFinish, Me.UserCode)
            If ShowDayFinish = "" Then ShowDayFinish = "1"
            ShowProcess = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_ShowProcess, Me.UserCode)
            If ShowProcess = "" Then ShowProcess = "1"
            ShowRate = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_ShowRate, Me.UserCode)
            If ShowRate = "" Then ShowRate = "1"
            ShowQty = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_ShowQty, Me.UserCode)
            If ShowQty = "" Then ShowQty = "1"
            ShowUnit = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_ShowUnit, Me.UserCode)
            If ShowUnit = "" Then ShowUnit = "1"
            ShowWH = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_ShowWH, Me.UserCode)
            If ShowWH = "" Then ShowWH = "1"

            ShowInputLocation = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_ShowInputLocation, Me.UserCode)
            If ShowInputLocation = "" Then ShowInputLocation = "1"
            ShowOutputLocation = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_ShowOutputLocation, Me.UserCode)
            If ShowOutputLocation = "" Then ShowOutputLocation = "1"

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
            If ShowProcess = "1" Then
                Me.ColLeft_ProcessCode.Visible = True
            Else
                Me.ColLeft_ProcessCode.Visible = False
            End If
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


            Me.ColLeft_FinishTime.HeaderText = ""
            Me.ColLeft_StartTime.HeaderText = ""
            Me.SetSomeColumnVisibilty()
            Me.ResizeMainPanel()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub filltgvSch()
        Dim DTA_D As New DseTrans_ScheduleTableAdapters.tbltrans_mo_item_woTableAdapter
        Dim row_D As DseTrans_Schedule.tbltrans_mo_item_woRow
        Dim BarFrom As Date = #1/1/2200#
        Dim NodeLevel0, NodeLevel1, NodeApply As AdvancedDataGridView.TreeGridNode

        Try
            DTA_D.Connection.ConnectionString = Core_Module.GetConstring
            DTA_D.Fill_All(Me.TBL_D)

            Me.tgvSch.Nodes.Clear()
            Me.tgvSch.Rows.Clear()
            Me.tgvSch.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            If Me.tblData.Count = 0 Then Exit Sub
            For Each Me.rowData In Me.tblData
                If BarFrom > Me.rowData.plan_start_time Then BarFrom = Me.rowData.plan_start_time
                While BarFrom.DayOfWeek <> DayOfWeek.Monday
                    BarFrom = BarFrom.AddDays(-1)
                End While
                Me.TBL_D.DefaultView.RowFilter = "[Identifier WO] = '" + Me.rowData.identifier_wo + "'"
                If Me.rowData("Tree Level") = 1 Then
                    NodeLevel0 = New AdvancedDataGridView.TreeGridNode()
                    NodeLevel0.CreateCells(Me.tgvSch)
                    Me.tgvSch.Nodes.Add(NodeLevel0)
                    NodeApply = NodeLevel0
                ElseIf rowData("Tree Level") = 2 Then
                    NodeLevel1 = New AdvancedDataGridView.TreeGridNode()
                    NodeLevel1.CreateCells(Me.tgvSch)
                    NodeLevel0.Nodes.Add(NodeLevel1)
                    NodeApply = NodeLevel1
                End If
                NodeApply.Visible = True
                With NodeApply
                    If Me.rowData("Tree Level") = 0 Then
                        .Cells(Me.ColLeft_Tree.Index).Value = Me.rowData("Tree")
                    ElseIf Me.rowData("Tree Level") = 1 Then
                        .Cells(Me.ColLeft_Tree.Index).Value = Me.rowData.process_code
                    Else
                        .Cells(Me.ColLeft_Tree.Index).Value = Me.rowData.jo_no
                    End If

                    If TBL_D.DefaultView.Count = 0 Then
                        .Cells(Me.ColLeft_WONo.Index).Value = ""
                        If rowData.Isproduct_codeNull = False And (rowData("Tree Level") <> 1) Then
                            .Cells(Me.ColLeft_CodeNo.Index).Value = rowData.product_code
                        Else
                            .Cells(Me.ColLeft_CodeNo.Index).Value = ""
                        End If
                        If rowData.Isproduct_descNull = False And (rowData("Tree Level") <> 1) Then
                            .Cells(Me.ColLeft_Description.Index).Value = rowData.product_desc
                        Else
                            .Cells(Me.ColLeft_Description.Index).Value = ""
                        End If
                    Else
                        row_D = Me.TBL_D.DefaultView(0).Row
                        .Cells(Me.ColLeft_WONo.Index).Value = row_D.WO_Number
                        .Cells(Me.ColLeft_CodeNo.Index).Value = row_D.Product_Code
                        .Cells(Me.ColLeft_Description.Index).Value = row_D.Description
                    End If
                    .Cells(Me.ColLeft_WONo.Index).Value = Me.rowData.wo_no
                    .Cells(Me.ColLeft_WorkUnit.Index).Value = Me.rowData.work_unit_code
                    .Cells(Me.ColLeft_DayOfWeekStart.Index).Value = Me.ObjCD.GetHari_English(Me.rowData.plan_start_time)
                    .Cells(Me.ColLeft_StartDate.Index).Value = Format(Me.rowData.plan_start_time, "dd MMM yy")
                    .Cells(Me.ColLeft_StartTime.Index).Value = Format(Me.rowData.plan_start_time, "HH.mm")
                    .Cells(Me.ColLeft_DayOfWeekFinish.Index).Value = Me.ObjCD.GetHari_English(Me.rowData.plan_end_time)
                    .Cells(Me.ColLeft_FinishDate.Index).Value = Format(Me.rowData.plan_end_time, "dd MMM yy")
                    .Cells(Me.ColLeft_FinishTime.Index).Value = Format(Me.rowData.plan_end_time, "HH.mm")

                    .Cells(Me.ColLeft_ProcessCode.Index).Value = Me.rowData.process_code
                    .Cells(Me.ColLeft_PrimaryWorkUnit.Index).Value = Me.rowData.primary_work_unit
                    .Cells(Me.ColLeft_PrimaryProcess.Index).Value = Me.rowData.primary_process
                    .Cells(Me.ColLeft_PrimaryMO.Index).Value = Me.rowData.primary_mo
                    .Cells(Me.ColLeft_IdentifierJO.Index).Value = Me.rowData.identifier_jo
                    .Cells(Me.ColLeft_IdentifierWO.Index).Value = Me.rowData.identifier_wo
                    .Cells(Me.ColLeft_a.Index).Value = Me.rowData.a
                    .Cells(Me.ColLeft_r.Index).Value = Me.rowData.r
                    .Cells(Me.ColLeft_g.Index).Value = Me.rowData.g
                    .Cells(Me.ColLeft_b.Index).Value = Me.rowData.b
                    .Cells(Me.ColLeft_NodeLevel.Index).Value = Me.rowData("Tree Level")
                    .Cells(Me.ColLeft_NodeLv0Identifier.Index).Value = Me.rowData("Node Identifier Lv 0")
                    .Cells(Me.ColLeft_NodeLv1Identifier.Index).Value = Me.rowData("Node Identifier Lv 1")
                    '.Cells(Me.ColLeft_Progress.Index).Value = Me.ObjCD.DigitGrouping(Me.rowData.progress, 2)
                    If Me.rowData("Is Header") = 0 Then
                        .Cells(Me.ColLeft_ProductionRate.Index).Value = Me.ObjCD.DigitGrouping(Me.rowData.rate, Me.rowData.decimal_point)
                        .Cells(Me.ColLeft_QtyProduce.Index).Value = Me.ObjCD.DigitGrouping(Me.rowData.qty_produce, Me.rowData.decimal_point)
                        .Cells(Me.ColLeft_TotalWH.Index).Value = Me.ObjCD.Convert_ToString(Me.rowData.total_work_hour)
                        .Cells(Me.ColLeft_Unit.Index).Value = Me.rowData.unit_code
                    ElseIf Me.rowData("Is Header") = 1 And Me.rowData("Tree Level") = 0 Then
                        .Cells(Me.ColLeft_QtyProduce.Index).Value = Me.ObjCD.DigitGrouping(Me.rowData.qty_produce, Me.rowData.decimal_point)
                        .Cells(Me.ColLeft_TotalWH.Index).Value = Me.ObjCD.Convert_ToString(Me.rowData.total_work_hour)
                        .Cells(Me.ColLeft_Unit.Index).Value = Me.rowData.unit_code
                    End If
                    .Cells(Me.ColLeft_PrimaryInputLocation.Index).Value = Me.rowData.primary_input_location
                    .Cells(Me.ColLeft_InputLocation.Index).Value = Me.rowData.input_location
                    .Cells(Me.ColLeft_PrimaryOutputLocation.Index).Value = Me.rowData.primary_output_location
                    .Cells(Me.ColLeft_OutputLocation.Index).Value = Me.rowData.output_location
                    Me.programmaticExpandCollaps = True
                    .Expand()
                End With
            Next
            Me.RefreshBar = False
            Me.dtpBarFrom.Value = CDate(Format(BarFrom, "dd MMMM yyyy") & " 00:00:00")
            Me.setColumnVisibility()
            Me.tgvSch.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_D.Dispose()
            TBL_D.Dispose()
        End Try
    End Sub

    Sub fillUnitData(ByRef vTblData As DseTrans_Schedule.tbl_sch_barDataTable)
        Dim dtaUnit As New DseTrans_ProductionSchedule_NewSch3TableAdapters.DataTable11TableAdapter
        Dim tblUnit As New DseTrans_ProductionSchedule_NewSch3.DataTable11DataTable
        Dim rowUnit As DseTrans_ProductionSchedule_NewSch3.DataTable11Row
        Dim dtaLocation As New Core_Module.dseMasterStorageAreaTableAdapters.tblmaster_storage_areaTableAdapter
        Dim tblLocation As New Core_Module.dseMasterStorageArea.tblmaster_storage_areaDataTable
        Dim rowLocation As Core_Module.dseMasterStorageArea.tblmaster_storage_areaRow

        Try
            dtaUnit.Connection.ConnectionString = Core_Module.GetConstring
            dtaUnit.Fill(tblUnit)
            For Each rowUnit In tblUnit
                vTblData.DefaultView.RowFilter = "[primary unit] = " & rowUnit.Primary.ToString
                If vTblData.DefaultView.Count > 0 Then
                    For i As Integer = 0 To vTblData.DefaultView.Count - 1
                        rowData = vTblData.DefaultView(i).Row
                        rowData.unit_code = rowUnit.Unit_Code
                    Next
                End If
            Next

            dtaLocation.Connection.ConnectionString = Core_Module.GetConstring
            dtaLocation.Fill(tblLocation)
            For Each rowLocation In tblLocation
                vTblData.DefaultView.RowFilter = "[primary input location] = " & rowLocation.Primary.ToString
                If vTblData.DefaultView.Count > 0 Then
                    For i As Integer = 0 To vTblData.DefaultView.Count - 1
                        rowData = vTblData.DefaultView(i).Row
                        rowData.input_location = rowLocation.Storage_Area_Code
                    Next
                End If
                vTblData.DefaultView.RowFilter = "[primary output location] = " & rowLocation.Primary.ToString
                If vTblData.DefaultView.Count > 0 Then
                    For i As Integer = 0 To vTblData.DefaultView.Count - 1
                        rowData = vTblData.DefaultView(i).Row
                        rowData.output_location = rowLocation.Storage_Area_Code
                    Next
                End If
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub translateData()
        Dim listOfWorkUnit As New List(Of String)
        Dim NodeIdentiferLv0 As Integer = -1
        Dim NodeIdentiferLv1 As Integer = -1
        Try
            Me.tblData.Rows.Clear()
            If tblData.Columns.Contains("Tree") = False Then
                Me.tblData.Columns.Add("Tree", System.Type.GetType("System.String"))
                Me.tblData.Columns("Tree").DefaultValue = 0
            End If
            If Me.tblData.Columns.Contains("Tree Level") = False Then
                Me.tblData.Columns.Add("Tree Level", System.Type.GetType("System.Double"))
                Me.tblData.Columns("Tree Level").DefaultValue = "0"
            End If
            If Me.tblData.Columns.Contains("Node Identifier Lv 0") = False Then
                Me.tblData.Columns.Add("Node Identifier Lv 0", System.Type.GetType("System.String"))
                Me.tblData.Columns("Node Identifier Lv 0").DefaultValue = "-1"
            End If
            If Me.tblData.Columns.Contains("Node Identifier Lv 1") = False Then
                Me.tblData.Columns.Add("Node Identifier Lv 1", System.Type.GetType("System.String"))
                Me.tblData.Columns("Node Identifier Lv 1").DefaultValue = "-1"
            End If
            If Me.tblData.Columns.Contains("Is Header") = False Then
                Me.tblData.Columns.Add("Is Header", System.Type.GetType("System.Double"))
                Me.tblData.Columns("Is Header").DefaultValue = 0
            End If
            Me.ColLeft_Tree.HeaderText = "Job Order"

            NodeIdentiferLv1 = -1
            Me.rowData = Me.tblData.Newtbl_sch_barRow
            NodeIdentiferLv0 += 1
            With Me.rowData
                .tipe = ClsTrans_Schedule.EnumTipeSchBar.work_time
                TBLSchBar.DefaultView.RowFilter = "[identifier jo]='" & Me.IdentifierJO & "'"
                TBLSchBar.DefaultView.Sort = "plan start time ASC"
                If TBLSchBar.DefaultView.Count = 0 Then
                    Exit Sub
                End If
                rowSch = Me.TBLSchBar.DefaultView(0).Row
                .plan_start_time = Me.rowSch.plan_start_time
                TBLSchBar.DefaultView.Sort = "plan end time DESC"
                rowSch = Me.TBLSchBar.DefaultView(0).Row
                .plan_end_time = Me.rowSch.plan_end_time
                .primary_mo = Me.rowSch.primary_mo
                .identifier_jo = Me.IdentifierJO
                .a = Me.rowSch.a
                .r = Me.rowSch.r
                .g = Me.rowSch.g
                .b = Me.rowSch.b
                If String.IsNullOrEmpty(Me.rowSch.mo_no) Then
                    .mo_no = Me.objMO.GetNoTransaksi(Me.rowSch.primary_mo)
                Else
                    .mo_no = Me.rowSch.primary_mo
                End If
                If String.IsNullOrEmpty(Me.rowSch.jo_no) Then
                    .jo_no = Me.objMO.GetJONumber(Me.rowSch.identifier_jo)
                Else
                    .jo_no = Me.rowSch.jo_no
                End If
                .product_code = Me.TxtCodeNo.Text
                .product_desc = Me.TxtDesc.Text
                .qty_produce = Me.TxtQtyProduce.Text
                Me.rowData("Tree") = Me.TxtJoNumber.Text
                Me.rowData("Tree Level") = 0
                Me.rowData("Node Identifier Lv 0") = NodeIdentiferLv0
                Me.rowData("Node Identifier Lv 1") = NodeIdentiferLv1
                Me.rowData("Is Header") = 1

                Me.rowData.primary_input_location = -1
                Me.rowData.input_location = String.Empty
                Me.rowData.primary_output_location = -1
                Me.rowData.output_location = String.Empty

            End With
            'Me.tblData.Rows.Add(Me.rowData)
            ''===
            Me.TBLSchBar.DefaultView.RowFilter = "[identifier jo]='" & Me.IdentifierJO & "'"
            Me.TBLSchBar.DefaultView.Sort = "plan start time ASC"
            Dim listOfProcess As New List(Of Double)
            For i As Integer = 0 To Me.TBLSchBar.DefaultView.Count - 1
                Me.rowSch = Me.TBLSchBar.DefaultView(i).Row
                If listOfProcess.Contains(Me.rowSch.primary_process) = False Then listOfProcess.Add(Me.rowSch.primary_process)
            Next
            For Each primaryProcess As Double In listOfProcess
                Me.TBLSchBar.DefaultView.RowFilter = "[identifier jo]='" & Me.IdentifierJO & "' AND [primary process] = " & primaryProcess.ToString
                Me.rowData = Me.tblData.Newtbl_sch_barRow
                NodeIdentiferLv1 += 1
                With Me.rowData
                    .tipe = ClsTrans_Schedule.EnumTipeSchBar.work_time
                    TBLSchBar.DefaultView.Sort = "plan start time ASC"
                    Me.rowSch = Me.TBLSchBar.DefaultView(0).Row
                    .plan_start_time = Me.rowSch.plan_start_time
                    Me.TBLSchBar.DefaultView.Sort = "plan end time DESC"
                    Me.rowSch = Me.TBLSchBar.DefaultView(0).Row
                    .plan_end_time = Me.rowSch.plan_end_time
                    .primary_mo = Me.rowSch.primary_mo
                    .identifier_jo = Me.IdentifierJO
                    .primary_process = primaryProcess
                    .a = Me.rowSch.a
                    .r = Me.rowSch.r
                    .g = Me.rowSch.g
                    .b = Me.rowSch.b
                    If String.IsNullOrEmpty(Me.rowSch.process_code) Then
                        .process_code = Me.objProcess.GetCode(primaryProcess)
                    Else
                        .process_code = Me.rowSch.process_code
                    End If
                    If String.IsNullOrEmpty(Me.rowSch.mo_no) Then
                        .mo_no = Me.objMO.GetNoTransaksi(Me.rowSch.primary_mo)
                    Else
                        .mo_no = Me.rowSch.mo_no
                    End If
                    If String.IsNullOrEmpty(Me.rowSch.jo_no) Then
                        .jo_no = Me.objMO.GetJONumber(Me.rowSch.identifier_jo)
                    Else
                        .jo_no = Me.rowSch.jo_no
                    End If
                    Me.rowData("Tree") = .process_code
                    Me.rowData("Tree Level") = 1
                    Me.rowData("Node Identifier Lv 0") = NodeIdentiferLv0
                    Me.rowData("Node Identifier Lv 1") = NodeIdentiferLv1
                    Me.rowData("Is Header") = 1

                    Me.rowData.primary_input_location = -1
                    Me.rowData.input_location = String.Empty
                    Me.rowData.primary_output_location = -1
                    Me.rowData.output_location = String.Empty

                End With
                Me.tblData.Rows.Add(Me.rowData)
                ''===
                Me.TBLSchBar.DefaultView.Sort = "[wo no]"
                For i As Integer = 0 To Me.TBLSchBar.DefaultView.Count - 1
                    Me.rowSch = Me.TBLSchBar.DefaultView(i).Row
                    Me.rowData = Me.tblData.Newtbl_sch_barRow
                    With Me.rowData
                        .tipe = Me.rowSch.tipe
                        .plan_start_time = Me.rowSch.plan_start_time
                        .plan_end_time = Me.rowSch.plan_end_time
                        .primary_mo = Me.rowSch.primary_mo
                        .identifier_jo = Me.rowSch.identifier_jo
                        .identifier_wo = Me.rowSch.identifier_wo
                        .primary_process = Me.rowSch.primary_process
                        .primary_work_unit = Me.rowSch.primary_work_unit
                        .primary_unit = Me.rowSch.primary_unit
                        .a = Me.rowSch.a
                        .r = Me.rowSch.r
                        .g = Me.rowSch.g
                        .b = Me.rowSch.b
                        .progress = Me.rowSch.progress
                        .duration = Me.rowSch.duration
                        If String.IsNullOrEmpty(Me.rowSch.process_code) Then
                            .process_code = Me.objProcess.GetCode(primaryProcess)
                        Else
                            .process_code = Me.rowSch.process_code
                        End If
                        If String.IsNullOrEmpty(Me.rowSch.mo_no) Then
                            .mo_no = Me.objMO.GetNoTransaksi(Me.rowSch.primary_mo)
                        Else
                            .mo_no = Me.rowSch.mo_no
                        End If
                        If String.IsNullOrEmpty(Me.rowSch.jo_no) Then
                            .jo_no = Me.objMO.GetJONumber(Me.rowSch.identifier_jo)
                        Else
                            .jo_no = Me.rowSch.jo_no
                        End If
                        If String.IsNullOrEmpty(Me.rowSch.wo_no) Then
                            .wo_no = Me.objMO.GetWONumber(Me.rowSch.identifier_wo, Me.rowSch.identifier_jo)
                        Else
                            .wo_no = Me.rowSch.wo_no
                        End If
                        If String.IsNullOrEmpty(Me.rowSch.work_unit_code) Then
                            .work_unit_code = Me.ObjWU.GetCode(Me.rowSch.primary_work_unit)
                        Else
                            .work_unit_code = Me.rowSch.work_unit_code
                        End If
                        .rate = Me.rowSch.rate
                        If Me.rowSch.Isdecimal_pointNull = True Then
                            .decimal_point = 2
                        Else
                            .decimal_point = Me.rowSch.decimal_point
                        End If
                        If rowSch.Istotal_work_hourNull Then
                            .total_work_hour = 0
                        Else
                            .total_work_hour = rowSch.total_work_hour
                        End If
                        .qty_produce = rowSch.qty_produce
                        If String.IsNullOrEmpty(Me.rowSch.unit_code) Then
                            .unit_code = Me.objUnit.GetUnitCode(Me.rowSch.primary_unit)
                        Else
                            .unit_code = rowSch.unit_code
                        End If
                        .primary_wu_config = rowSch.primary_wu_config
                        Me.rowData("Tree") = .wo_no & " : " & .process_code
                        Me.rowData("Tree Level") = 2
                        Me.rowData.primary_input_location = Me.rowSch.primary_input_location
                        Me.rowData.input_location = Me.rowSch.input_location
                        Me.rowData.primary_output_location = Me.rowSch.primary_output_location
                        Me.rowData.output_location = Me.rowSch.output_location

                    End With
                    Me.tblData.Rows.Add(Me.rowData)
                Next
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_ProductionSchedule_NewSch2_ResizeEnd(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.ResizeEnd
        Try
            Me.ResizeMainPanel()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_ProductionSchedule_NewSch2_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        Try
            Me.ResizeMainPanel()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_ProductionSchedule_NewSch2_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.SizeChanged
        Try
            Me.ResizeMainPanel()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker5_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker5.DoWork
        Try
            Threading.Thread.Sleep(10)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker5_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker5.RunWorkerCompleted
        Try
            Me.setColumnVisibility()
            Me.ResizeMainPanel()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub tgvSch_NodeCollapsed(ByVal sender As Object, ByVal e As AdvancedDataGridView.CollapsedEventArgs) Handles tgvSch.NodeCollapsed
        Try
            If Me.programmaticExpandCollaps = False Then
                Me.collapsOrExpandBar()
                If e.Node.RowIndex > -1 Then
                    Me.tblData.DefaultView.RowFilter = "[Tree Level] = " & e.Node.DataGridView.Item(Me.ColLeft_NodeLevel.Index, e.Node.RowIndex).Value.ToString & _
                        " AND [Node Identifier Lv 0] = '" & e.Node.DataGridView.Item(Me.ColLeft_NodeLv0Identifier.Index, e.Node.RowIndex).Value & "'" & _
                        " AND [Node Identifier Lv 1] = '" & e.Node.DataGridView.Item(Me.ColLeft_NodeLv1Identifier.Index, e.Node.RowIndex).Value & "'"
                    Me.rowData = Me.tblData.DefaultView(0).Row
                    Me.reFillTgvHeader(Me.rowData, e.Node.RowIndex, True)
                    Me.reDrawBarHeader(Me.rowData, True)
                End If
                If Me.ExpandAll = False Then
                    Me.BackgroundWorker5.RunWorkerAsync()
                End If
            Else
                Me.programmaticExpandCollaps = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub reDrawBarHeader(ByVal vRowData As Inventory_Module.DseTrans_Schedule.tbl_sch_barRow, ByVal vShowBar As Boolean)
        Dim rowIndex As Integer = Me.tblData.Rows.IndexOf(vRowData)
        Dim dr As DataGridViewRow

        Dim intervalPerBar As TimeSpan
        Dim barValueFrom, barValueTo As Date

        Try
            If Me.cboTimeframe.SelectedIndex = ClsTrans_Schedule.enumDurationScale.By_Hour Then
                intervalPerBar = intervalPerBar_ByHour
            ElseIf Me.cboTimeframe.SelectedIndex = ClsTrans_Schedule.enumDurationScale.By_Day Then
                intervalPerBar = intervalPerBar_ByDay
            ElseIf Me.cboTimeframe.SelectedIndex = ClsTrans_Schedule.enumDurationScale.By_Weeks Then
                intervalPerBar = intervalPerBar_ByWeek
            End If

            dr = Me.dgvBar.Rows(Me.tblData.Rows.IndexOf(vRowData))
            barValueFrom = Me.dtpBarFrom.Value
            barValueTo = barValueFrom + intervalPerBar
            Dim workBarCount As Integer = 0
            Dim firstBarIndex As Integer = -1
            Dim barProgressCount As Integer = 0
            For i As Integer = 1 To Me.dgvBar.Columns.Count
                With vRowData
                    If ((rowData.Item("Is Header") = 0) Or (rowData.Item("Is Header") = 1 And vShowBar = True)) And ((barValueFrom <= .plan_start_time And barValueTo > .plan_start_time) Or _
                        (barValueFrom >= .plan_start_time And barValueTo <= .plan_end_time) Or _
                        (barValueFrom < .plan_end_time And barValueTo >= .plan_end_time)) Then
                        dr.Cells("ColBar" & i).Style.BackColor = System.Drawing.Color.FromArgb(.a, .r, .g, .b)
                        If firstBarIndex = -1 Then
                            firstBarIndex = i
                        End If
                        workBarCount += 1
                    Else
                        dr.Cells("ColBar" & i).Value = Nothing
                        If Me.isOnAlternateColumn(i) Then
                            dr.Cells("ColBar" & i).Style.BackColor = Me.AlternateBarColor
                        Else
                            dr.Cells("ColBar" & i).Style.BackColor = Color.White
                        End If
                    End If
                    barValueFrom += intervalPerBar
                    barValueTo += intervalPerBar
                End With
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub reFillTgvHeader(ByVal vRowData As Inventory_Module.DseTrans_Schedule.tbl_sch_barRow, ByVal vRowIndex As Integer, ByVal vShowTime As Boolean)
        Dim dr As DataGridViewRow = Me.tgvSch.Rows(vRowIndex)
        Try
            With dr
                If Me.rowData("Is Header") = 1 Then
                    .Cells(Me.ColLeft_DayOfWeekStart.Index).Value = Me.ObjCD.GetHari_English(Me.rowData.plan_start_time)
                    .Cells(Me.ColLeft_StartDate.Index).Value = Format(Me.rowData.plan_start_time, "dd MMM yy")
                    .Cells(Me.ColLeft_StartTime.Index).Value = Format(Me.rowData.plan_start_time, "HH.mm")
                    .Cells(Me.ColLeft_DayOfWeekFinish.Index).Value = Me.ObjCD.GetHari_English(Me.rowData.plan_end_time)
                    .Cells(Me.ColLeft_FinishDate.Index).Value = Format(Me.rowData.plan_end_time, "dd MMM yy")
                    .Cells(Me.ColLeft_FinishTime.Index).Value = Format(Me.rowData.plan_end_time, "HH.mm")
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub collapsOrExpandBar()
        Dim rowIndex As Integer = 0
        Try
            For Each node As AdvancedDataGridView.TreeGridNode In Me.tgvSch.Nodes
                If node.Nodes.Count > 0 Then
                    For Each childNode0 As AdvancedDataGridView.TreeGridNode In node.Nodes
                        Me.setBarVisibility(rowIndex, childNode0)
                    Next
                End If
                rowIndex += 1
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Function setBarVisibility(ByRef vRowIndex As Integer, ByVal vNode As AdvancedDataGridView.TreeGridNode) As Boolean
        Try
            vRowIndex += 1
            Me.ColLeft_ProcessCode.Visible = False
            If vNode.IsSited Then
                Me.dgvBar.Rows(vRowIndex).Visible = True
            Else
                Me.dgvBar.Rows(vRowIndex).Visible = False
            End If
            If vNode.Nodes.Count > 0 Then
                For Each childNode1 As AdvancedDataGridView.TreeGridNode In vNode.Nodes
                    setBarVisibility(vRowIndex, childNode1)
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub tgvSch_NodeExpanded(ByVal sender As Object, ByVal e As AdvancedDataGridView.ExpandedEventArgs) Handles tgvSch.NodeExpanded
        Try
            If Me.programmaticExpandCollaps = False Then
                Me.collapsOrExpandBar()
                If e.Node.RowIndex > -1 Then
                    For Each Me.rowData In Me.tblData.Rows
                        If Me.rowData("Tree Level") = e.Node.DataGridView.Item(Me.ColLeft_NodeLevel.Index, e.Node.RowIndex).Value.ToString And _
                         Me.rowData("Node Identifier Lv 0") = e.Node.DataGridView.Item(Me.ColLeft_NodeLv0Identifier.Index, e.Node.RowIndex).Value.ToString And _
                         Me.rowData("Node Identifier Lv 1") = e.Node.DataGridView.Item(Me.ColLeft_NodeLv1Identifier.Index, e.Node.RowIndex).Value.ToString Then
                            Exit For
                        End If
                    Next
                    Me.reDrawBarHeader(Me.rowData, False)
                End If
                If Me.ExpandAll = False Then
                    Me.BackgroundWorker5.RunWorkerAsync()
                End If
            Else
                Me.programmaticExpandCollaps = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub tgvSch_Scroll(ByVal sender As Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles tgvSch.Scroll
        Try
            If Me.tgvSch.Rows.Count <= 0 Then
                Exit Sub
            End If

            If Me.getShowRowCount = Me.tgvSch.Rows.Count Then
                Dim index As Integer = Me.tgvSch.FirstDisplayedScrollingRowIndex + Me.getHidenRowCount(Me.dgvBar.FirstDisplayedScrollingRowIndex)
                While Me.dgvBar.Rows(index).Visible = False
                    If e.NewValue > e.OldValue Then
                        index += 1
                    Else
                        index -= 1
                    End If
                End While
                Me.dgvBar.FirstDisplayedScrollingRowIndex = index
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function getShowRowCount() As Integer
        Dim v As Integer = 0
        Try
            For Each row As DataGridViewRow In Me.dgvBar.Rows
                If row.Visible = True Then
                    v += 1
                End If
            Next
            Return v
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        End Try
    End Function

    Function getHidenRowCount(ByVal vLastRowCheck As Integer) As Integer
        Dim count As Integer = 0
        Try
            For i As Integer = 0 To vLastRowCheck - 1
                If Me.dgvBar.Rows(i).Visible = False Then
                    count += 1
                End If
            Next
            Return count
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub dgvBar_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgvBar.MouseMove
        Dim objMousePosition As Point
        Dim objHitTestInfo As DataGridView.HitTestInfo
        Dim toolTipTextCurrent As String = ""
        Dim toolTipTitleCurrent As String = ""
        Dim job_order_number As String = ""
        Dim wo_no As String = ""
        Dim process_code As String = ""
        Dim work_unit_code As String = ""
        Dim code_no As String = ""
        Dim description As String = ""
        Dim rate_unit_per_hour As Decimal = 0
        Dim qty As Decimal = 0
        Dim unit_code As String = ""
        Dim work_hour As Decimal = 0
        Dim start As Date = #1/1/1950#
        Dim finish As Date = #1/1/1950#
        Dim progress As Decimal = 0
        Try
            objMousePosition = Me.dgvBar.PointToClient(Control.MousePosition)
            objHitTestInfo = Me.dgvBar.HitTest(objMousePosition.X, objMousePosition.Y)
            If objHitTestInfo.ColumnIndex > -1 And objHitTestInfo.RowIndex > -1 Then
                toolTipTextCurrent = Me.getTooltipText(Me.cboTimeframe.SelectedIndex, Me.dtpBarFrom.Value, _
                                        Me.dgvBar.Columns(objHitTestInfo.ColumnIndex).Name, objHitTestInfo.RowIndex, toolTipTitleCurrent, _
                                         job_order_number, wo_no, process_code, work_unit_code, code_no, description, rate_unit_per_hour, qty, unit_code, _
                                         work_hour, start, finish, progress, objHitTestInfo.RowIndex + 1)
                If toolTipTextCurrent <> "" Then
                    If Me.toolTipText <> toolTipTextCurrent Then
                        Me.toolTipText = toolTipTextCurrent
                        Me.dgvBar.Cursor = Cursors.Hand
                        Me.frmTooltip = New FrmTrans_ProductionSchedule_NewSch2_Tooltip
                        Me.frmTooltip.Location = New Point(Me.PointToClient(Control.MousePosition).X + 20, Me.PointToClient(Control.MousePosition).Y + 40)
                        Me.frmTooltip.tooltipColor = Color.PeachPuff
                        Me.frmTooltip.job_order_number = job_order_number
                        Me.frmTooltip.wo_no = wo_no
                        Me.frmTooltip.process_code = process_code
                        Me.frmTooltip.work_unit_code = work_unit_code
                        Me.frmTooltip.code_no = code_no
                        Me.frmTooltip.description = description
                        Me.frmTooltip.rate_unit_per_hour = rate_unit_per_hour
                        Me.frmTooltip.qty = qty
                        Me.frmTooltip.unit_code = unit_code
                        Me.frmTooltip.work_hour = work_hour
                        Me.frmTooltip.start = start
                        Me.frmTooltip.finish = finish
                        Me.frmTooltip.progress = -1
                        Me.frmTooltip.Show()
                    End If
                Else
                    Me.toolTipText = ""
                    Me.dgvBar.Cursor = Cursors.Default
                    Me.dgvBar.DefaultCellStyle.SelectionBackColor = Color.Transparent
                    If IsNothing(Me.frmTooltip) = False Then
                        Me.frmTooltip.Dispose()
                    End If
                End If
            Else
                Me.toolTipText = ""
                Me.dgvBar.Cursor = Cursors.Default
                Me.dgvBar.DefaultCellStyle.SelectionBackColor = Color.Transparent
                If IsNothing(Me.frmTooltip) = False Then
                    Me.frmTooltip.Dispose()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Function getTooltipText(ByVal vIntervalScale As ClsTrans_Schedule.enumDurationScale, ByVal vFrom As Date, ByVal vColName As String, _
            ByVal vBarIndex As Integer, ByRef vTooltipTitle As String, _
            ByRef job_order_number As String, ByRef wo_no As String, ByRef process_code As String, ByRef work_unit_code As String, _
            ByRef code_no As String, ByRef description As String, ByRef rate_unit_per_hour As Decimal, ByRef qty As Decimal, _
            ByRef unit_code As String, ByRef work_hour As Decimal, ByRef start As Date, ByRef finish As Date, ByRef progress As Decimal, _
            ByVal vNextRowIndex As Integer) As String
        Dim retVal As String = ""
        Dim strBarStart As String = ""
        Dim strBarEnd As String = ""
        Try
            strBarStart = Format(Me.getBarTime(True, vIntervalScale, vFrom, vColName), "MM/dd/yyyy HH:mm:00")
            strBarEnd = Format(Me.getBarTime(False, vIntervalScale, vFrom, vColName), "MM/dd/yyyy HH:mm:00")
            Me.tblData.DefaultView.RowFilter = "[bar index] = " & vBarIndex.ToString & _
                                               " AND (([plan start time] >= #" & strBarStart & "# AND [plan start time] <= #" & strBarEnd & _
                                               "#) OR ([plan start time] <= #" & strBarStart & "# AND [plan end time] >= #" & strBarEnd & _
                                               "#) OR ([plan end time] >= #" & strBarStart & "# AND [plan end time] <= #" & strBarEnd & "#))"
            If Me.tblData.DefaultView.Count > 0 Then
                Me.rowData = Me.tblData.DefaultView(0).Row
                If Me.rowData("Is Header") = 1 Then
                    If vNextRowIndex > Me.dgvBar.Rows.Count - 1 Then
                        vTooltipTitle = ""
                        Return ""
                    End If
                    If Me.dgvBar.Rows(vNextRowIndex).Visible Then
                        vTooltipTitle = ""
                        Return ""
                    End If
                End If
                job_order_number = Me.rowData.jo_no
                wo_no = Me.rowData.wo_no
                process_code = Me.rowData.process_code
                work_unit_code = Me.rowData.work_unit_code
                code_no = Me.rowData.product_code
                description = Me.rowData.product_desc
                rate_unit_per_hour = Me.rowData.rate
                qty = Me.rowData.qty_produce
                unit_code = Me.rowData.unit_code
                work_hour = Me.rowData.total_work_hour
                start = Me.rowData.plan_start_time
                finish = Me.rowData.plan_end_time
                'progress = Me.rowData.progress
                'line_code = Me.rowData("line code")
                retVal = job_order_number & wo_no & process_code & work_unit_code & code_no & description & rate_unit_per_hour.ToString & qty.ToString & unit_code & work_hour.ToString & start.ToString & finish.ToString & progress.ToString
            End If
            Return retVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return retVal
        End Try
    End Function

    Function getBarTime(ByVal vBarStart As Boolean, ByVal vIntervalScale As ClsTrans_Schedule.enumDurationScale, ByVal vFrom As Date, ByVal vColName As String) As Date
        Dim intervalPerBar As TimeSpan
        Dim retVal As Date = vFrom

        Try
            If vIntervalScale = ClsTrans_Schedule.enumDurationScale.By_Hour Then
                intervalPerBar = intervalPerBar_ByHour
            ElseIf vIntervalScale = ClsTrans_Schedule.enumDurationScale.By_Day Then
                intervalPerBar = intervalPerBar_ByDay
            ElseIf vIntervalScale = ClsTrans_Schedule.enumDurationScale.By_Weeks Then
                intervalPerBar = intervalPerBar_ByWeek
            End If
            For i As Integer = 1 To CInt(vColName.Replace("ColBar", ""))
                retVal += intervalPerBar
            Next
            If vBarStart Then retVal -= intervalPerBar
            Return retVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub dgvBar_Scroll(ByVal sender As Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles dgvBar.Scroll
        Try
            If Me.tgvSch.RowCount <= 0 Then
                Exit Sub
            End If

            Me.dgvHeader0.HorizontalScrollingOffset = Me.dgvBar.HorizontalScrollingOffset
            Me.dgvHeader1.HorizontalScrollingOffset = Me.dgvBar.HorizontalScrollingOffset
            If Me.getShowRowCount = Me.tgvSch.Rows.Count Then
                Me.tgvSch.FirstDisplayedScrollingRowIndex = Me.dgvBar.FirstDisplayedScrollingRowIndex - Me.getHidenRowCount(Me.dgvBar.FirstDisplayedScrollingRowIndex)
            End If
            If (Me.dgvBar.HorizontalScrollingOffset > 0 And Me.dgvBar.HorizontalScrollingOffset < 2011) Then
                Me.callEndOfScroll = False
            Else
                Me.callEndOfScroll = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbExpandAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbExpandAll.LinkClicked
        Try
            Me.ExpandAll = True
            Me.tgvSch.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For Each node As AdvancedDataGridView.TreeGridNode In Me.tgvSch.Nodes
                node.Expand()
                For Each node1 As AdvancedDataGridView.TreeGridNode In node.Nodes
                    node1.Expand()
                Next
            Next
            Me.tgvSch.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.setColumnVisibility()
            Me.SetSomeColumnVisibilty()
            Me.ExpandAll = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbCollapsAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbCollapsAll.LinkClicked
        Try
            Me.CollapsAll()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LlbColumnSettingSch_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbColumnSettingSch.LinkClicked
        Dim ij As New FrmTrans_ProductionSchedule_NewSch2_ColumnSetting
        Try
            ij.CalledFromMainGantt = False
            ij.UserCode = Me.UserCode

            If ij.ShowDialog() = Windows.Forms.DialogResult.OK Then
                Me.setColumnVisibility()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

#End Region

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        Try
            'Me.RefrehGanttChart = True
            If Me.TabControl1.SelectedTab.Name = Me.TabPage2.Name Then
                'Me.dgvBar.Height = Me.tgvSch.Height
                Me.LlbColumnSettingSch.Visible = True
                If Me.RefrehGanttChart Then
                    Me.reloadGanttChart()
                    Me.RefrehGanttChart = False
                End If
            ElseIf Me.TabControl1.SelectedTab.Name = Me.TabPage3.Name Then
                Me.LoadException()
            Else
                Me.LlbColumnSettingSch.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Private Sub dgvException_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvException.RowsAdded
    '    Try
    '        MsgBox("")
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Private Sub llbSetRate_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbSetRate.LinkClicked
        Dim ij As New FrmTrans_ProductionSchedule_NewSch6
        Dim rowDGV As DataGridViewRow
        Dim ListDummy As New List(Of String)
        Dim PrimaryProses As Double = -1
        Dim DummyStartTime As Date = Date.MaxValue 
        Dim RowItemWO_ItemSch As DseTrans_Schedule.tbltrans_mo_item_wo_item_schRow
        Dim StartEvaluate As Boolean = False
        Dim TBLRef As New DseTrans_Schedule2.TblRefDataTable
        Dim rowRef As DseTrans_Schedule2.TblRefRow

        Try
            'isi tblref 
            For Each RowItemWO_ItemSch In Me.TBLItemWO_ItemSch
                If ListDummy.Contains(RowItemWO_ItemSch.primary_process.ToString + "-" + RowItemWO_ItemSch.primary_work_unit.ToString) = False Then
                    ListDummy.Add(RowItemWO_ItemSch.primary_process.ToString + "-" + RowItemWO_ItemSch.primary_work_unit.ToString)

                    rowRef = ij.TblRef.NewTblRefRow
                    With rowRef
                        .Primary_Process = RowItemWO_ItemSch.primary_process
                        .Process_Code = Me.objProcess.GetCode(.Primary_Process)
                        .Primary_WU = RowItemWO_ItemSch.primary_work_unit
                        .Work_Unit_Code = Me.ObjWU.GetCode(.Primary_WU)
                        .Work_Unit_Desc = Me.ObjWU.GetDesc(.Primary_WU)
                        .Rate = RowItemWO_ItemSch.plan_prod_rate
                    End With
                    ij.TblRef.Rows.Add(rowRef)
                End If
            Next

            If ij.ShowDialog() = Windows.Forms.DialogResult.OK Then
                'siapkan tbl ref rate
                PrimaryProses = Me.objProcess.GetPrimary(ij.cboProcess.Text)

                For Each rowDGV In ij.dgvWU.Rows
                    rowref = TBLRef.NewTblRefRow
                    With rowRef
                        .Primary_Process = PrimaryProses
                        .Process_Code = ij.cboProcess.Text
                        .Primary_WU = rowDGV.Cells(ij.ColPrimaryWorkUnit.Index).Value
                        .Work_Unit_Code = rowDGV.Cells(ij.ColWorkUnitCode.Index).Value
                        .Work_Unit_Desc = rowDGV.Cells(ij.ColWorkUnitDesc.Index).Value
                        .Rate = rowDGV.Cells(ij.ColRate.Index).Value
                    End With
                    TBLRef.Rows.Add(rowref)
                Next

                'panggil fungsi reschedule new rate
                Me.TBLSchBar.DefaultView.RowFilter = "[identifier jo] = '" + Me.IdentifierJO + "'"
                For Each dv As Data.DataRowView In Me.TBLSchBar.DefaultView
                    rowSch = dv.Row

                    Me.ObjSch.Reschedule_NewRate(Me.IdentifierJO, Me.TBLSchBar, Me.TBLItemWO, Me.TBLItemWO_ItemSch, rowSch.a, rowSch.r, rowSch.g, _
                    rowSch.b, Me.TBLSchBarKosong, Me.TBLSchBarTanpaWO, Me.PrimaryPredefinedConfig, TBLRef)

                    'set posisi acuan + reload data 
                    If Me.DgvWUAllocate_Process.SelectedCells.Count > 0 Then
                        Me.r1 = Me.DgvWUAllocate_Process.SelectedCells(0).RowIndex
                        Me.c1 = Me.DgvWUAllocate_Process.SelectedCells(0).ColumnIndex
                    Else
                        Me.r1 = 0
                        Me.c1 = 0
                    End If
                    Me.bgw4.RunWorkerAsync()

                    Exit For
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
            ListDummy = Nothing
            TBLRef.Rows.Clear()
            TBLRef.Dispose()
        End Try
    End Sub 
End Class