Option Explicit On
Option Strict Off
Public Class FrmTrans_ProductionSchedule
    Public UserCode As String = "USER"
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public userIdentifier As Double = 0 '-> ganti ke random Double.Parse(System.Text.RegularExpressions.Regex.Replace(System.Net.NetworkInformation.NetworkInterface.GetAllNetworkInterfaces()(0).GetPhysicalAddress.ToString, "[^\d]", ""))
    Public userType As Core_Module.ClsTransaksiMonitor_Production.enumUserType = Core_Module.ClsTransaksiMonitor_Production.enumUserType.READ_OR_FULLCONTROL

    Dim objMonitor As Core_Module.ClsTransaksiMonitor_Production
    Dim userControl As Core_Module.ClsTransaksiMonitor_Production.enumUserControl = Core_Module.ClsTransaksiMonitor_Production.enumUserControl.READ_ONLY

    Dim ObjLocation As New Core_Module.clsMasterStorageArea
    Dim objConfig As New FrmTrans_ProductionSchedule_NewSch3
    Dim objLine As New FrmTrans_ProductionSchedule_NewSch2_ColumnSetting
    Dim listOfLine As New List(Of Double)
    Dim ObjLicense As New Core_Module.ClsLicense
    Dim objMasterLine As New Core_Module.ClsMasterMachineLine

    Dim ObjProd As New Core_Module.clsMasterProduct
    Dim objCD As New Core_Module.ClsCekData
    Dim objUser As New Core_Module.ClsUser
    Dim ObjSch As New ClsTrans_Schedule
    Dim ObjSetting1 As New Core_Module.ClsTransaksiSetting
    Dim tblSchBar As New DseTrans_Schedule.tbl_sch_barDataTable
    Dim tblSchBar_TanpaItemWO As New DseTrans_Schedule.tbl_sch_barDataTable
    Dim rowSch As DseTrans_Schedule.tbl_sch_barRow
    Dim tblData As New DseTrans_Schedule.tbl_sch_barDataTable
    Dim rowData As DseTrans_Schedule.tbl_sch_barRow
    Dim tblBarIndex As New DseTrans_Schedule.BarIndexDataTable
    Dim rowBarIndex As DseTrans_Schedule.BarIndexRow
    Dim intervalPerBar_ByHour As TimeSpan = TimeSpan.FromMinutes(5)
    Dim intervalPerBar_ByDay As TimeSpan = TimeSpan.FromHours(1)
    Dim intervalPerBar_ByWeek As TimeSpan = TimeSpan.FromHours(12)
    Dim toolTipText As String = ""
    Dim callEndOfScroll As Boolean = True
    Dim objWU As New Core_Module.ClsMasterWorkUnit
    Dim refreshBar As Boolean = True
    Dim frmTooltip As FrmTrans_ProductionSchedule_NewSch2_Tooltip
    Dim TBL_D As New DseTrans_Schedule.tbltrans_mo_item_woDataTable
    Dim objSetting As New Core_Module.ClsTransaksiSettingUser
    Dim alternateBarColor As Color = Color.WhiteSmoke
    Dim Buffer As String = ""
    Dim BufferTime As Date = Today
    Dim line_firmplanned As Bitmap = My.Resources.line_firmplanned
    Dim line_today As Bitmap = My.Resources.line_today
    Dim gantt_progress As Bitmap = My.Resources.gantt_progress
    Dim gantt_progress_firmplanned As Bitmap = My.Resources.gantt_progress_firmplanned
    Dim gantt_progress_today As Bitmap = My.Resources.gantt_progress_today
    Dim LineTodayIndex As Integer = -1
    Dim LineFirmPlannedIndex As Integer = -1
    Dim ObjGlobalSetting As New Core_Module.ClsTransaksiSetting
    Dim ShowGanttChart As Boolean = False
    Dim EnumGroupBy As enumGroup = enumGroup.BY_JO
    Dim OnPorgressLoadingTGV As Boolean = False
    Dim ScheduleHorizon As Integer = 30 '-> tadinya 12 * 7, sementara sampai gantt chart dinyalain lagi
    Dim JmlNodeExpand As Double = 0

    Enum enumGroup As Integer
        BY_JO
        BY_WORKUNIT
    End Enum

    Dim programmaticExpandCollaps As Boolean = False
    Dim mNodeLevel As Integer = -1
    Dim mNodeLevel0 As String = ""
    Dim mNodeLevel1 As String = ""
    Dim ObjMOItem As New ClsTrans_ManufacturingOrder_Item
    Dim ObjMO As New ClsTrans_ManufacturingOrder
    Dim ObjLog As New ClsTrans_StockBundling

    Sub collapsOrExpandBar(ByVal vNode As AdvancedDataGridView.TreeGridNode, ByVal rowIndex As Integer)
        Try
            If Me.ShowGanttChart = True Then
                If vNode.Nodes.Count > 0 Then
                    For Each childNode0 As AdvancedDataGridView.TreeGridNode In vNode.Nodes
                        Me.setBarVisibility(rowIndex, childNode0)
                    Next
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Function dgvAddListener(ByVal dataGridView As DataGridView) As Boolean
        Dim pInfo As Reflection.PropertyInfo
        Dim dgvScrollBar As ScrollBar
        Try
            pInfo = dataGridView.GetType.GetProperty("HorizontalScrollBar", Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic)
            If pInfo Is Nothing Then Return False
            dgvScrollBar = CType(pInfo.GetValue(dataGridView, Nothing), ScrollBar)
            If dgvScrollBar Is Nothing Then Return False
            AddHandler dgvScrollBar.Scroll, New ScrollEventHandler(AddressOf Me.dgvScroll)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Sub dgvScroll(ByVal sender As Object, ByVal e As ScrollEventArgs)
        Dim firstDisplayedScrollingRowIndex As Integer = 0
        Try
            firstDisplayedScrollingRowIndex = Me.dgvBar.FirstDisplayedScrollingRowIndex
            If Me.dgvBar.Rows.Count >= firstDisplayedScrollingRowIndex And Me.tgvSch.Rows.Count >= firstDisplayedScrollingRowIndex And _
                            Me.dgvBar.Rows.Count > 0 And Me.tgvSch.Rows.Count > 0 Then
                Me.dgvBar.FirstDisplayedScrollingRowIndex = firstDisplayedScrollingRowIndex
                Me.tgvSch.FirstDisplayedScrollingRowIndex = firstDisplayedScrollingRowIndex
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

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
                    Dim resetRow As Boolean = False
                    Dim ResetRow1 As Boolean = False
                    For i As Integer = 1 To Me.dgvBar.Columns.Count
                        If barValueFrom.Date = Now.Date And resetRow = False Then
                            Me.LineTodayIndex = i
                            resetRow = True
                            dr.Cells("ColBar" & i).Value = line_today
                            dr.Cells("ColBar" & i).Style.Alignment = DataGridViewContentAlignment.MiddleRight
                        End If
                        If barValueFrom.Date = Me.dtpFirmPlanned.Value.Date And ResetRow1 = False Then
                            Me.LineFirmPlannedIndex = i
                            ResetRow1 = True
                            dr.Cells("ColBar" & i).Value = line_firmplanned
                            dr.Cells("ColBar" & i).Style.Alignment = DataGridViewContentAlignment.MiddleRight
                        End If
                        If Me.rowData("Is Header") = 0 And ((barValueFrom <= .plan_start_time And barValueTo > .plan_start_time) Or _
                                                            (barValueFrom >= .plan_start_time And barValueTo <= .plan_end_time) Or _
                                                            (barValueFrom < .plan_end_time And barValueTo >= .plan_end_time)) Then
                            dr.Cells("ColBar" & i).Style.BackColor = System.Drawing.Color.FromArgb(.a, .r, .g, .b)
                        Else
                            If Me.isOnAlternateColumn(i) Then
                                dr.Cells("ColBar" & i).Style.BackColor = Me.alternateBarColor
                            Else
                                dr.Cells("ColBar" & i).Style.BackColor = Color.White
                            End If
                        End If
                        barValueFrom += intervalPerBar
                        barValueTo += intervalPerBar
                    Next
                End With
            Next
            Me.ShowProgress(vIntervalScale, vFrom)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub ShowProgress(ByVal vIntervalScale As ClsTrans_Schedule.enumDurationScale, ByVal vFrom As Date)
        Dim dr As DataGridViewRow
        Dim intervalPerBar As TimeSpan
        Dim barValueFrom, barValueTo As Date
        Try
            If Me.ShowGanttChart = False Then
                Exit Sub
            End If
            If vIntervalScale = ClsTrans_Schedule.enumDurationScale.By_Hour Then
                intervalPerBar = intervalPerBar_ByHour
            ElseIf vIntervalScale = ClsTrans_Schedule.enumDurationScale.By_Day Then
                intervalPerBar = intervalPerBar_ByDay
            ElseIf vIntervalScale = ClsTrans_Schedule.enumDurationScale.By_Weeks Then
                intervalPerBar = intervalPerBar_ByWeek
            End If
            For Each Me.rowData In Me.tblData
                dr = Me.dgvBar.Rows(Me.tblData.Rows.IndexOf(Me.rowData))
                barValueFrom = vFrom
                barValueTo = vFrom + intervalPerBar
                Dim workBarCount As Integer = 0
                Dim firstBarIndex As Integer = -1
                Dim barProgressCount As Integer = 0
                With Me.rowData
                    For i As Integer = 1 To Me.dgvBar.Columns.Count
                        If ((barValueFrom <= .plan_start_time And barValueTo > .plan_start_time) Or _
                                (barValueFrom >= .plan_start_time And barValueTo <= .plan_end_time) Or _
                                (barValueFrom < .plan_end_time And barValueTo >= .plan_end_time)) Then
                            If firstBarIndex = -1 Then
                                firstBarIndex = i
                            End If
                            workBarCount += 1
                        End If
                        barValueFrom += intervalPerBar
                        barValueTo += intervalPerBar
                    Next
                    barProgressCount = Me.objCD.DigitGrouping(Me.rowData.progress / 100 * workBarCount, 0, Core_Module.ClsCekData.EnumTipeRounding.Down)
                    If barProgressCount > 0 Then
                        For i As Integer = firstBarIndex To firstBarIndex - 1 + barProgressCount
                            If Me.chkShowProgress.Checked Then
                                Dim Draw As Boolean = False
                                If Me.tblData.Rows.IndexOf(Me.rowData) = Me.dgvBar.Rows.Count - 1 Then
                                    Draw = True
                                ElseIf Me.rowData.is_header = 0 Or (Me.rowData.is_header = 1 And Me.dgvBar.Rows(Me.tblData.Rows.IndexOf(Me.rowData) + 1).Visible = False) Then
                                    Draw = True
                                End If
                                If Draw Then
                                    If i = Me.LineTodayIndex Then
                                        dr.Cells("ColBar" & i).Value = gantt_progress_today
                                    ElseIf i = Me.LineFirmPlannedIndex Then
                                        dr.Cells("ColBar" & i).Value = gantt_progress_firmplanned
                                    Else
                                        dr.Cells("ColBar" & i).Value = gantt_progress
                                    End If
                                End If
                            Else
                                If i = Me.LineTodayIndex Then
                                    dr.Cells("ColBar" & i).Value = Me.line_today
                                ElseIf i = Me.LineFirmPlannedIndex Then
                                    dr.Cells("ColBar" & i).Value = Me.line_firmplanned
                                Else
                                    dr.Cells("ColBar" & i).Value = Nothing
                                End If
                            End If
                        Next
                    End If
                End With
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub filltgvSch()
        'Dim DTA_D As New DseTrans_ScheduleTableAdapters.tbltrans_mo_item_woTableAdapter
        'Dim row_D As DseTrans_Schedule.tbltrans_mo_item_woRow
        Dim BarFrom As Date = Me.dtpBarFrom.Value
        Dim NodeLevel0, NodeLevel1, NodeLevel2, NodeApply As AdvancedDataGridView.TreeGridNode
        Dim vGroup As enumGroup = Me.cboGroupBy.SelectedIndex
        Try
            Me.OnPorgressLoadingTGV = True
            'DTA_D.Connection.ConnectionString = Core_Module.GetConstring
            'DTA_D.Fill_All(Me.TBL_D)
            Me.tgvSch.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.tgvSch.Nodes.Clear()
            Me.tgvSch.Rows.Clear()
            If Me.tblData.Count = 0 Then Exit Sub
            For Each Me.rowData In Me.tblData
                If BarFrom > Me.rowData.plan_start_time Then BarFrom = Me.rowData.plan_start_time
                While BarFrom.DayOfWeek <> DayOfWeek.Monday
                    BarFrom = BarFrom.AddDays(-1)
                End While
                Me.TBL_D.DefaultView.RowFilter = "[Identifier WO] = '" + Me.rowData.identifier_wo + "'"
                If Me.rowData("Tree Level") = 0 Then
                    NodeLevel0 = New AdvancedDataGridView.TreeGridNode()
                    NodeLevel0.CreateCells(Me.tgvSch)
                    Me.tgvSch.Nodes.Add(NodeLevel0)
                    NodeApply = NodeLevel0
                ElseIf rowData("Tree Level") = 1 Then
                    NodeLevel1 = New AdvancedDataGridView.TreeGridNode()
                    NodeLevel1.CreateCells(Me.tgvSch)
                    NodeLevel0.Nodes.Add(NodeLevel1)
                    NodeApply = NodeLevel1
                ElseIf rowData("Tree Level") = 2 Then
                    NodeLevel2 = New AdvancedDataGridView.TreeGridNode()
                    NodeLevel2.CreateCells(Me.tgvSch)
                    NodeLevel1.Nodes.Add(NodeLevel2)
                    NodeApply = NodeLevel2
                End If
                NodeApply.Visible = True
                With NodeApply
                    If vGroup = enumGroup.BY_JO Then
                        If Me.rowData("Tree Level") = 0 Then
                            .Cells(Me.ColLeft_Tree.Index).Value = Me.rowData("Tree")

                            If Me.rowData("is closed") = 1 Then
                                .Cells(Me.ColLeft_Tree.Index).Value += " (Closed)"
                            End If

                        ElseIf Me.rowData("Tree Level") = 1 Then
                            .Cells(Me.ColLeft_Tree.Index).Value = Me.rowData.process_code
                        Else
                            .Cells(Me.ColLeft_Tree.Index).Value = Me.rowData.wo_no
                        End If
                    Else
                        If Me.rowData("Tree Level") = 0 Then
                            .Cells(Me.ColLeft_Tree.Index).Value = Me.rowData("Tree")
                        Else
                            .Cells(Me.ColLeft_Tree.Index).Value = Me.rowData.jo_no
                        End If
                    End If

                    If vGroup = enumGroup.BY_JO Then
                        If (rowData("Tree Level") = 2) Then
                            'row_D = Me.TBL_D.DefaultView(0).Row
                            .Cells(Me.ColLeft_WONo.Index).Value = ""
                            .Cells(Me.ColLeft_CodeNo.Index).Value = ""
                            .Cells(Me.ColLeft_Description.Index).Value = ""
                            'Me.rowData.product_code = row_D.Product_Code
                        Else
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
                        End If
                    Else
                        .Cells(Me.ColLeft_WONo.Index).Value = ""
                        .Cells(Me.ColLeft_CodeNo.Index).Value = ""
                        .Cells(Me.ColLeft_Description.Index).Value = ""
                        'Me.rowData.product_code = row_D.Product_Code
                    End If


                    'If TBL_D.DefaultView.Count = 0 Then
                    '    .Cells(Me.ColLeft_WONo.Index).Value = ""
                    '    If rowData.Isproduct_codeNull = False And (rowData("Tree Level") <> 1) Then
                    '        .Cells(Me.ColLeft_CodeNo.Index).Value = rowData.product_code
                    '    Else
                    '        .Cells(Me.ColLeft_CodeNo.Index).Value = ""
                    '    End If
                    '    If rowData.Isproduct_descNull = False And (rowData("Tree Level") <> 1) Then
                    '        .Cells(Me.ColLeft_Description.Index).Value = rowData.product_desc
                    '    Else
                    '        .Cells(Me.ColLeft_Description.Index).Value = ""
                    '    End If
                    'Else
                    '    row_D = Me.TBL_D.DefaultView(0).Row
                    '    .Cells(Me.ColLeft_WONo.Index).Value = row_D.WO_Number
                    '    .Cells(Me.ColLeft_CodeNo.Index).Value = row_D.Product_Code
                    '    .Cells(Me.ColLeft_Description.Index).Value = row_D.Description
                    '    Me.rowData.product_code = row_D.Product_Code
                    'End If
                    .Cells(Me.ColLeft_WONo.Index).Value = Me.rowData.wo_no
                    .Cells(Me.ColLeft_WorkUnit.Index).Value = Me.rowData.work_unit_code
                    If Me.cboGroupBy.SelectedIndex = enumGroup.BY_WORKUNIT And Me.rowData("Tree Level") = 0 Then
                        ''jika by mesin dan header, gak usah dimunculkan waktunya.
                    Else
                        .Cells(Me.ColLeft_DayOfWeekStart.Index).Value = Me.objCD.GetHari_English(Me.rowData.plan_start_time)
                        .Cells(Me.ColLeft_StartDate.Index).Value = Format(Me.rowData.plan_start_time, "dd MMM yy")
                        .Cells(Me.ColLeft_StartTime.Index).Value = Format(Me.rowData.plan_start_time, "HH.mm")
                        .Cells(Me.ColLeft_DayOfWeekFinish.Index).Value = Me.objCD.GetHari_English(Me.rowData.plan_end_time)
                        .Cells(Me.ColLeft_FinishDate.Index).Value = Format(Me.rowData.plan_end_time, "dd MMM yy")
                        .Cells(Me.ColLeft_FinishTime.Index).Value = Format(Me.rowData.plan_end_time, "HH.mm")
                    End If
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
                    If Me.rowData("Is Header") = 0 Then
                        .Cells(Me.ColLeft_ProductionRate.Index).Value = Me.objCD.DigitGrouping(Me.rowData.rate, Me.rowData.decimal_point)
                        .Cells(Me.ColLeft_QtyProduce.Index).Value = Me.objCD.DigitGrouping(Me.rowData.qty_produce, Me.rowData.decimal_point)
                        .Cells(Me.ColLeft_TotalWH.Index).Value = Me.objCD.Convert_ToString(Me.rowData.total_work_hour)
                        .Cells(Me.ColLeft_Unit.Index).Value = Me.rowData.unit_code
                    ElseIf Me.rowData("Is Header") = 1 And Me.rowData("Tree Level") = 0 And Me.cboGroupBy.SelectedIndex = enumGroup.BY_JO Then
                        .Cells(Me.ColLeft_QtyProduce.Index).Value = Me.objCD.DigitGrouping(Me.rowData.qty_produce, Me.rowData.decimal_point)
                        .Cells(Me.ColLeft_TotalWH.Index).Value = Me.objCD.Convert_ToString(Me.rowData.total_work_hour)
                        .Cells(Me.ColLeft_Unit.Index).Value = Me.rowData.unit_code
                    End If
                    .Cells(Me.ColLeft_PrimaryLine.Index).Value = Me.rowData.primary_line
                    .Cells(Me.ColLeft_LineCode.Index).Value = Me.rowData.line_code
                    .Cells(Me.ColLeft_PrimaryLocationInput.Index).Value = Me.rowData.primary_input_location
                    .Cells(Me.ColLeft_LocationInput.Index).Value = Me.rowData.input_location
                    .Cells(Me.ColLeft_PrimaryLocationOutput.Index).Value = Me.rowData.primary_output_location
                    .Cells(Me.ColLeft_LocationOutput.Index).Value = Me.rowData.output_location

                    If Me.EnumGroupBy = enumGroup.BY_JO Then
                        .Cells(Me.ColLeft_Actualproduce.Index).Value = Me.objCD.DigitGrouping(Me.rowData.actual_produce_qty, Me.rowData.decimal_point)
                        .Cells(Me.ColLeft_ActualGood.Index).Value = Me.objCD.DigitGrouping(Me.rowData.actual_good_qty, Me.rowData.decimal_point)
                        .Cells(Me.ColLeft_ActualReject.Index).Value = Me.objCD.DigitGrouping(Me.rowData.actual_reject_qty, Me.rowData.decimal_point)
                        .Cells(Me.ColLeft_Progress.Index).Value = Me.objCD.DigitGrouping(Me.rowData.progress, 2)
                    Else
                        If .Cells(Me.ColLeft_NodeLevel.Index).Value = 0 Then
                            .Cells(Me.ColLeft_Actualproduce.Index).Value = ""
                            .Cells(Me.ColLeft_ActualGood.Index).Value = ""
                            .Cells(Me.ColLeft_ActualReject.Index).Value = ""
                            .Cells(Me.ColLeft_Progress.Index).Value = ""
                        Else
                            .Cells(Me.ColLeft_Actualproduce.Index).Value = Me.objCD.DigitGrouping(Me.rowData.actual_produce_qty, Me.rowData.decimal_point)
                            .Cells(Me.ColLeft_ActualGood.Index).Value = Me.objCD.DigitGrouping(Me.rowData.actual_good_qty, Me.rowData.decimal_point)
                            .Cells(Me.ColLeft_ActualReject.Index).Value = Me.objCD.DigitGrouping(Me.rowData.actual_reject_qty, Me.rowData.decimal_point)
                            .Cells(Me.ColLeft_Progress.Index).Value = Me.objCD.DigitGrouping(Me.rowData.progress, 2)
                        End If
                    End If
                    Me.programmaticExpandCollaps = True
                    '.Expand()
                End With
            Next
            Me.refreshBar = False
            Me.dtpBarFrom.Value = CDate(Format(BarFrom, "dd MMMM yyyy") & " 00:00:00")
            Me.setColumnVisibility_1()
            Me.tgvSch.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            'DTA_D.Dispose()
            TBL_D.Dispose()
            Me.OnPorgressLoadingTGV = False
            Me.JmlNodeExpand = 0
        End Try
    End Sub

    'Sub filltgvSch_obs()
    '    Dim DTA_D As New DseTrans_ScheduleTableAdapters.tbltrans_mo_item_woTableAdapter
    '    Dim row_D As DseTrans_Schedule.tbltrans_mo_item_woRow
    '    Dim BarFrom As Date = Me.dtpBarFrom.Value
    '    Dim NodeLevel0, NodeLevel1, NodeLevel2, NodeApply As AdvancedDataGridView.TreeGridNode
    '    Dim vGroup As enumGroup = Me.cboGroupBy.SelectedIndex
    '    Try
    '        DTA_D.Connection.ConnectionString = Core_Module.GetConstring
    '        DTA_D.Fill_All(Me.TBL_D)
    '        Me.tgvSch.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
    '        Me.tgvSch.Nodes.Clear()
    '        Me.tgvSch.Rows.Clear()
    '        If Me.tblData.Count = 0 Then Exit Sub
    '        For Each Me.rowData In Me.tblData
    '            If BarFrom > Me.rowData.plan_start_time Then BarFrom = Me.rowData.plan_start_time
    '            While BarFrom.DayOfWeek <> DayOfWeek.Monday
    '                BarFrom = BarFrom.AddDays(-1)
    '            End While
    '            Me.TBL_D.DefaultView.RowFilter = "[Identifier WO] = '" + Me.rowData.identifier_wo + "'"
    '            If Me.rowData("Tree Level") = 0 Then
    '                NodeLevel0 = New AdvancedDataGridView.TreeGridNode()
    '                NodeLevel0.CreateCells(Me.tgvSch)
    '                Me.tgvSch.Nodes.Add(NodeLevel0)
    '                NodeApply = NodeLevel0
    '            ElseIf rowData("Tree Level") = 1 Then
    '                NodeLevel1 = New AdvancedDataGridView.TreeGridNode()
    '                NodeLevel1.CreateCells(Me.tgvSch)
    '                NodeLevel0.Nodes.Add(NodeLevel1)
    '                NodeApply = NodeLevel1
    '            ElseIf rowData("Tree Level") = 2 Then
    '                NodeLevel2 = New AdvancedDataGridView.TreeGridNode()
    '                NodeLevel2.CreateCells(Me.tgvSch)
    '                NodeLevel1.Nodes.Add(NodeLevel2)
    '                NodeApply = NodeLevel2
    '            End If
    '            NodeApply.Visible = True
    '            With NodeApply
    '                If vGroup = enumGroup.BY_JO Then
    '                    If Me.rowData("Tree Level") = 0 Then
    '                        .Cells(Me.ColLeft_Tree.Index).Value = Me.rowData("Tree")

    '                        If Me.rowData("is closed") = 1 Then
    '                            .Cells(Me.ColLeft_Tree.Index).Value += " (Closed)"
    '                        End If

    '                    ElseIf Me.rowData("Tree Level") = 1 Then
    '                        .Cells(Me.ColLeft_Tree.Index).Value = Me.rowData.process_code
    '                    Else
    '                        .Cells(Me.ColLeft_Tree.Index).Value = Me.rowData.wo_no
    '                    End If
    '                Else
    '                    If Me.rowData("Tree Level") = 0 Then
    '                        .Cells(Me.ColLeft_Tree.Index).Value = Me.rowData("Tree")
    '                    Else
    '                        .Cells(Me.ColLeft_Tree.Index).Value = Me.rowData.jo_no
    '                    End If
    '                End If
    '                If TBL_D.DefaultView.Count = 0 Then
    '                    .Cells(Me.ColLeft_WONo.Index).Value = ""
    '                    If rowData.Isproduct_codeNull = False And (rowData("Tree Level") <> 1) Then
    '                        .Cells(Me.ColLeft_CodeNo.Index).Value = rowData.product_code
    '                    Else
    '                        .Cells(Me.ColLeft_CodeNo.Index).Value = ""
    '                    End If
    '                    If rowData.Isproduct_descNull = False And (rowData("Tree Level") <> 1) Then
    '                        .Cells(Me.ColLeft_Description.Index).Value = rowData.product_desc
    '                    Else
    '                        .Cells(Me.ColLeft_Description.Index).Value = ""
    '                    End If
    '                Else
    '                    row_D = Me.TBL_D.DefaultView(0).Row
    '                    .Cells(Me.ColLeft_WONo.Index).Value = row_D.WO_Number
    '                    .Cells(Me.ColLeft_CodeNo.Index).Value = row_D.Product_Code
    '                    .Cells(Me.ColLeft_Description.Index).Value = row_D.Description
    '                    Me.rowData.product_code = row_D.Product_Code
    '                End If
    '                .Cells(Me.ColLeft_WONo.Index).Value = Me.rowData.wo_no
    '                .Cells(Me.ColLeft_WorkUnit.Index).Value = Me.rowData.work_unit_code
    '                If Me.cboGroupBy.SelectedIndex = enumGroup.BY_WORKUNIT And Me.rowData("Tree Level") = 0 Then
    '                    ''jika by mesin dan header, gak usah dimunculkan waktunya.
    '                Else
    '                    .Cells(Me.ColLeft_DayOfWeekStart.Index).Value = Me.objCD.GetHari_English(Me.rowData.plan_start_time)
    '                    .Cells(Me.ColLeft_StartDate.Index).Value = Format(Me.rowData.plan_start_time, "dd MMM yy")
    '                    .Cells(Me.ColLeft_StartTime.Index).Value = Format(Me.rowData.plan_start_time, "HH.mm")
    '                    .Cells(Me.ColLeft_DayOfWeekFinish.Index).Value = Me.objCD.GetHari_English(Me.rowData.plan_end_time)
    '                    .Cells(Me.ColLeft_FinishDate.Index).Value = Format(Me.rowData.plan_end_time, "dd MMM yy")
    '                    .Cells(Me.ColLeft_FinishTime.Index).Value = Format(Me.rowData.plan_end_time, "HH.mm")
    '                End If
    '                .Cells(Me.ColLeft_ProcessCode.Index).Value = Me.rowData.process_code
    '                .Cells(Me.ColLeft_PrimaryWorkUnit.Index).Value = Me.rowData.primary_work_unit
    '                .Cells(Me.ColLeft_PrimaryProcess.Index).Value = Me.rowData.primary_process
    '                .Cells(Me.ColLeft_PrimaryMO.Index).Value = Me.rowData.primary_mo
    '                .Cells(Me.ColLeft_IdentifierJO.Index).Value = Me.rowData.identifier_jo
    '                .Cells(Me.ColLeft_IdentifierWO.Index).Value = Me.rowData.identifier_wo
    '                .Cells(Me.ColLeft_a.Index).Value = Me.rowData.a
    '                .Cells(Me.ColLeft_r.Index).Value = Me.rowData.r
    '                .Cells(Me.ColLeft_g.Index).Value = Me.rowData.g
    '                .Cells(Me.ColLeft_b.Index).Value = Me.rowData.b
    '                .Cells(Me.ColLeft_NodeLevel.Index).Value = Me.rowData("Tree Level")
    '                .Cells(Me.ColLeft_NodeLv0Identifier.Index).Value = Me.rowData("Node Identifier Lv 0")
    '                .Cells(Me.ColLeft_NodeLv1Identifier.Index).Value = Me.rowData("Node Identifier Lv 1")
    '                .Cells(Me.ColLeft_Progress.Index).Value = Me.objCD.DigitGrouping(Me.rowData.progress, 2)
    '                If Me.rowData("Is Header") = 0 Then
    '                    .Cells(Me.ColLeft_ProductionRate.Index).Value = Me.objCD.DigitGrouping(Me.rowData.rate, Me.rowData.decimal_point)
    '                    .Cells(Me.ColLeft_QtyProduce.Index).Value = Me.objCD.DigitGrouping(Me.rowData.qty_produce, Me.rowData.decimal_point)
    '                    .Cells(Me.ColLeft_TotalWH.Index).Value = Me.objCD.Convert_ToString(Me.rowData.total_work_hour)
    '                    .Cells(Me.ColLeft_Unit.Index).Value = Me.rowData.unit_code
    '                ElseIf Me.rowData("Is Header") = 1 And Me.rowData("Tree Level") = 0 And Me.cboGroupBy.SelectedIndex = enumGroup.BY_JO Then
    '                    .Cells(Me.ColLeft_QtyProduce.Index).Value = Me.objCD.DigitGrouping(Me.rowData.qty_produce, Me.rowData.decimal_point)
    '                    .Cells(Me.ColLeft_TotalWH.Index).Value = Me.objCD.Convert_ToString(Me.rowData.total_work_hour)
    '                    .Cells(Me.ColLeft_Unit.Index).Value = Me.rowData.unit_code
    '                End If
    '                .Cells(Me.ColLeft_PrimaryLine.Index).Value = Me.rowData.primary_line
    '                .Cells(Me.ColLeft_LineCode.Index).Value = Me.rowData.line_code
    '                .Cells(Me.ColLeft_PrimaryLocationInput.Index).Value = Me.rowData.primary_input_location
    '                .Cells(Me.ColLeft_LocationInput.Index).Value = Me.rowData.input_location
    '                .Cells(Me.ColLeft_PrimaryLocationOutput.Index).Value = Me.rowData.primary_output_location
    '                .Cells(Me.ColLeft_LocationOutput.Index).Value = Me.rowData.output_location
    '                .Cells(Me.ColLeft_Actualproduce.Index).Value = Me.objCD.DigitGrouping(Me.rowData.actual_produce_qty, Me.rowData.decimal_point)
    '                .Cells(Me.ColLeft_ActualGood.Index).Value = Me.objCD.DigitGrouping(Me.rowData.actual_good_qty, Me.rowData.decimal_point)
    '                .Cells(Me.ColLeft_ActualReject.Index).Value = Me.objCD.DigitGrouping(Me.rowData.actual_reject_qty, Me.rowData.decimal_point)
    '                Me.programmaticExpandCollaps = True
    '                .Expand()
    '            End With
    '        Next
    '        Me.refreshBar = False
    '        Me.dtpBarFrom.Value = CDate(Format(BarFrom, "dd MMMM yyyy") & " 00:00:00")
    '        Me.setColumnVisibility_1()
    '        Me.tgvSch.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTA_D.Dispose()
    '        TBL_D.Dispose()
    '    End Try
    'End Sub

    Sub Find(ByVal vPrimaryMO As Double, ByVal vIdentifierJO As String, ByVal vIdentifierWO As String, ByVal vPrimaryProcess As Double, ByVal vPrimaryWorkunit As Double)
        Try
            Me.CollapsAll()

            If Me.cboGroupBy.SelectedIndex = 0 Then
                For Each Node0 As AdvancedDataGridView.TreeGridNode In tgvSch.Nodes
                    If Me.tgvSch.Item(Me.ColLeft_IdentifierJO.Index, Node0.RowIndex).Value = vIdentifierJO Then
                        Node0.Expand()
                        For Each Node1 As AdvancedDataGridView.TreeGridNode In Node0.Nodes
                            If Me.tgvSch.Item(Me.ColLeft_PrimaryProcess.Index, Node1.RowIndex).Value = vPrimaryProcess Then
                                Node1.Expand()
                                For Each Node2 As AdvancedDataGridView.TreeGridNode In Node1.Nodes
                                    If Me.tgvSch.Item(Me.ColLeft_IdentifierWO.Index, Node2.RowIndex).Value = vIdentifierWO And Me.tgvSch.Item(Me.ColLeft_PrimaryWorkUnit.Index, Node2.RowIndex).Value = vPrimaryWorkunit Then
                                        Me.tgvSch.ClearSelection()
                                        Node2.Selected = True
                                        Me.tgvSch.FirstDisplayedScrollingRowIndex = Node0.RowIndex
                                        Exit Sub
                                    End If
                                Next
                            End If
                        Next
                    End If
                Next
            Else
                For Each Node0 As AdvancedDataGridView.TreeGridNode In tgvSch.Nodes
                    If Me.tgvSch.Item(Me.ColLeft_PrimaryWorkUnit.Index, Node0.RowIndex).Value = vPrimaryWorkunit Then
                        Node0.Expand()
                        For Each Node1 As AdvancedDataGridView.TreeGridNode In Node0.Nodes
                            If Me.tgvSch.Item(Me.ColLeft_IdentifierJO.Index, Node1.RowIndex).Value = vIdentifierJO And Me.tgvSch.Item(Me.ColLeft_IdentifierWO.Index, Node1.RowIndex).Value = vIdentifierWO And Me.tgvSch.Item(Me.ColLeft_PrimaryProcess.Index, Node1.RowIndex).Value = vPrimaryProcess Then
                                Me.tgvSch.ClearSelection()
                                Node1.Selected = True
                                Me.tgvSch.FirstDisplayedScrollingRowIndex = Node0.RowIndex
                                Exit Sub
                            End If
                        Next
                    End If
                Next
            End If
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
                dc = Me.dgvBar.Columns(Me.dgvBar.Columns.Add(c))
                dc.Width = dgvBarWidth
            Next
            Me.dgvHeader0.Columns.Clear()
            For i As Integer = 1 To header0ColumnCount
                If vIntervalScale <> ClsTrans_Schedule.enumDurationScale.By_Weeks Then
                    dc = Me.dgvHeader0.Columns(Me.dgvHeader0.Columns.Add("ColHeader0" & i, Format(durationHeader0, formatDurationHeader0)))
                Else
                    dc = Me.dgvHeader0.Columns(Me.dgvHeader0.Columns.Add("ColHeader0" & i, Me.objCD.GetHari_English(durationHeader0).Substring(0, 1)))
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
                        dc = Me.dgvHeader1.Columns(Me.dgvHeader1.Columns.Add("ColHeader1" & i, Format(durationHeader1, "dd MMM yy") & " - " & Format(durationHeader1.Add(incrementHeader1).AddDays(-1), "dd MMM yyyy")))
                    ElseIf durationHeader1.Month <> durationHeader1.Add(incrementHeader1).Month Then
                        dc = Me.dgvHeader1.Columns(Me.dgvHeader1.Columns.Add("ColHeader1" & i, Format(durationHeader1, "dd MMM") & " - " & Format(durationHeader1.Add(incrementHeader1).AddDays(-1), "dd MMM yyyy")))
                    Else
                        dc = Me.dgvHeader1.Columns(Me.dgvHeader1.Columns.Add("ColHeader1" & i, Format(durationHeader1, "dd") & " - " & Format(durationHeader1.Add(incrementHeader1).AddDays(-1), "dd MMM yyyy")))
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

    Function getAllNodesCount(ByVal vOnlyDisplayed As Boolean) As Integer
        Dim vNodesCount As Integer = 0
        Try
            For Each node As AdvancedDataGridView.TreeGridNode In Me.tgvSch.Nodes
                Me.getAllNodesCount(vNodesCount, node, vOnlyDisplayed)
                If vOnlyDisplayed Then
                    If node.Displayed = True Then vNodesCount += 1
                Else
                    vNodesCount += 1
                End If
            Next
            Return vNodesCount
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Sub getAllNodescount(ByRef vNodeCount As Integer, ByRef vNode As AdvancedDataGridView.TreeGridNode, ByVal vOnlyDisplayed As Boolean)
        Try
            If vOnlyDisplayed Then
                If vNode.Displayed Then vNodeCount += 1
            Else
                vNodeCount += 1
            End If
            For Each node As AdvancedDataGridView.TreeGridNode In vNode.Nodes
                getAllNodesCount(vNodeCount, node, vOnlyDisplayed)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Function getBarIdentifierWO(ByVal vBarIndex As Integer) As String
        Try
            Me.tblBarIndex.DefaultView.RowFilter = "[Bar Index] = " & vBarIndex.ToString
            If Me.tblBarIndex.DefaultView.Count > 0 Then
                Me.rowBarIndex = Me.tblBarIndex.DefaultView(0).Row
                Return Me.rowBarIndex.Identifier_WO
            End If
            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Function getBarPrimaryWorkUnit(ByVal vBarIndex As Integer) As Double
        Try
            Me.tblBarIndex.DefaultView.RowFilter = "[Bar Index] = " & vBarIndex.ToString
            If Me.tblBarIndex.DefaultView.Count > 0 Then
                Me.rowBarIndex = Me.tblBarIndex.DefaultView(0).Row
                Return Me.rowBarIndex.Primary_Work_Unit
            End If
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
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

    Function getDgvBarRowsCount(ByVal vDisplayedOnly As Boolean) As Integer
        Dim rowsCount As Integer = 0
        Try
            For Each row As DataGridViewRow In Me.dgvBar.Rows
                If vDisplayedOnly Then
                    If row.Visible Then rowsCount += 1
                Else
                    rowsCount += 1
                End If
            Next
            Return rowsCount
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

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

    Sub fillOtherData(ByRef vTblData As DseTrans_Schedule.tbl_sch_barDataTable)
        'Dim dtaUnit As New DseTrans_ProductionSchedule_NewSch3TableAdapters.DataTable11TableAdapter
        'Dim tblUnit As New DseTrans_ProductionSchedule_NewSch3.DataTable11DataTable
        'Dim rowUnit As DseTrans_ProductionSchedule_NewSch3.DataTable11Row
        'Dim dtaJO As New DseTrans_ScheduleTableAdapters.tbltrans_mo_item2TableAdapter
        'Dim tblJO As New DseTrans_Schedule.tbltrans_mo_item2DataTable
        'Dim rowJO As DseTrans_Schedule.tbltrans_mo_item2Row
        'Dim dtaLocation As New Core_Module.dseMasterStorageAreaTableAdapters.tblmaster_storage_areaTableAdapter
        'Dim tblLocation As New Core_Module.dseMasterStorageArea.tblmaster_storage_areaDataTable
        'Dim rowLocation As Core_Module.dseMasterStorageArea.tblmaster_storage_areaRow
        Dim dtaLocation1 As New DseTrans_ScheduleTableAdapters.tbltrans_mo_item_wo_item_sch6TableAdapter
        Dim tblLocation1 As New DseTrans_Schedule.tbltrans_mo_item_wo_item_sch6DataTable
        Dim rowLocation1 As DseTrans_Schedule.tbltrans_mo_item_wo_item_sch6Row
        Dim dtaLocation2 As New DseTrans_ScheduleTableAdapters.tbltrans_mo_item_wo_item_sch7TableAdapter
        Dim tblLocation2 As New DseTrans_Schedule.tbltrans_mo_item_wo_item_sch7DataTable
        Dim rowLocation2 As DseTrans_Schedule.tbltrans_mo_item_wo_item_sch7Row
        Dim dtaBundling As New DseTrans_ScheduleTableAdapters.tbltransaksi_stockbundlingTableAdapter
        Dim tblBundling As New DseTrans_Schedule.tbltransaksi_stockbundlingDataTable
        Dim rowBundling As DseTrans_Schedule.tbltransaksi_stockbundlingRow
        Dim tblTotal0 As New DseTrans_Schedule.DataTable6DataTable
        Dim rowTotal0 As DseTrans_Schedule.DataTable6Row
        Dim tblTotal1 As New DseTrans_Schedule.DataTable6DataTable
        Dim rowTotal1 As DseTrans_Schedule.DataTable6Row
        Dim ketemu As Boolean = False

        Try
            'dtaUnit.Connection.ConnectionString = Core_Module.GetConstring
            'dtaUnit.Fill(tblUnit)
            'For Each rowUnit In tblUnit
            '    vTblData.DefaultView.RowFilter = "[primary unit] = " & rowUnit.Primary.ToString
            '    If vTblData.DefaultView.Count > 0 Then
            '        For i As Integer = 0 To vTblData.DefaultView.Count - 1
            '            rowData = vTblData.DefaultView(i).Row
            '            rowData.unit_code = rowUnit.Unit_Code
            '        Next
            '    End If
            'Next

            'dtaJO.Connection.ConnectionString = Core_Module.GetConstring
            'dtaJO.Fill(tblJO)
            'If tblData.Columns.Contains("is closed") = False Then
            '    tblData.Columns.Add("is closed", System.Type.GetType("System.Double"))
            '    tblData.Columns("is closed").DefaultValue = 0
            'End If
            'For Each rowJO In tblJO
            '    vTblData.DefaultView.RowFilter = "[identifier jo] = '" & rowJO.Identifier & "'"
            '    If vTblData.DefaultView.Count > 0 Then
            '        For i As Integer = 0 To vTblData.DefaultView.Count - 1
            '            rowData = vTblData.DefaultView(i).Row
            '            rowData("is closed") = rowJO.Is_Closed
            '        Next
            '    End If
            'Next

            dtaLocation1.Connection.ConnectionString = Core_Module.GetConstring
            dtaLocation2.Connection.ConnectionString = dtaLocation1.Connection.ConnectionString
            dtaLocation1.Fill(tblLocation1, CDate(Format(Me.dtpFrom.Value, "dd MMM yyyy") + " 00:00:00"))
            dtaLocation2.Fill(tblLocation2, CDate(Format(Me.dtpFrom.Value, "dd MMM yyyy") + " 00:00:00"))
            While tblLocation2.Rows.Count > 0
                rowLocation2 = tblLocation2.Rows(0)
                ketemu = False
                For Each rowLocation1 In tblLocation1.Rows
                    If rowLocation1.primary_location = rowLocation2.primary_location Then
                        ketemu = True
                        Exit For
                    End If
                Next

                If ketemu = False Then
                    rowLocation1 = tblLocation1.Newtbltrans_mo_item_wo_item_sch6Row
                    rowLocation1.primary_location = rowLocation2.primary_location
                    rowLocation1.Storage_Area_Code = rowLocation2.Storage_Area_Code
                    tblLocation1.Rows.Add(rowLocation1)
                End If

                tblLocation2.Rows.Remove(rowLocation2)
            End While

            While tblLocation1.Rows.Count > 0
                rowLocation1 = tblLocation1.Rows(0)
                vTblData.DefaultView.RowFilter = "[primary input location] = " & rowLocation1.primary_location.ToString
                If vTblData.DefaultView.Count > 0 Then
                    For i As Integer = 0 To vTblData.DefaultView.Count - 1
                        rowData = vTblData.DefaultView(i).Row
                        rowData.input_location = rowLocation1.Storage_Area_Code
                    Next
                End If
                vTblData.DefaultView.RowFilter = "[primary output location] = " & rowLocation1.primary_location.ToString
                If vTblData.DefaultView.Count > 0 Then
                    For i As Integer = 0 To vTblData.DefaultView.Count - 1
                        rowData = vTblData.DefaultView(i).Row
                        rowData.output_location = rowLocation1.Storage_Area_Code
                    Next
                End If
                tblLocation1.Rows.Remove(rowLocation1)
            End While

            'dtaLocation.Connection.ConnectionString = Core_Module.GetConstring
            'dtaLocation.Fill(tblLocation)
            'For Each rowLocation In tblLocation
            'vTblData.DefaultView.RowFilter = "[primary input location] = " & rowLocation.Primary.ToString
            'If vTblData.DefaultView.Count > 0 Then
            '    For i As Integer = 0 To vTblData.DefaultView.Count - 1
            '        rowData = vTblData.DefaultView(i).Row
            '        rowData.input_location = rowLocation.Storage_Area_Code
            '    Next
            'End If
            'vTblData.DefaultView.RowFilter = "[primary output location] = " & rowLocation.Primary.ToString
            'If vTblData.DefaultView.Count > 0 Then
            '    For i As Integer = 0 To vTblData.DefaultView.Count - 1
            '        rowData = vTblData.DefaultView(i).Row
            '        rowData.output_location = rowLocation.Storage_Area_Code
            '    Next
            'End If
            'Next

            dtaBundling.Connection.ConnectionString = Core_Module.GetConstring
            dtaBundling.Fill(tblBundling, ClsTrans_StockBundling.EnumTipeBundling.Produksi1, Me.dtpFrom.Value.Date)
            For Each rowData In tblData
                tblBundling.DefaultView.RowFilter = "[identifier jo] = '" & rowData.identifier_jo & "' and [identifier wo] = '" & rowData.identifier_wo & "' and [primary work unit] = " & rowData.primary_work_unit.ToString
                If tblBundling.DefaultView.Count > 0 Then
                    rowBundling = tblBundling.DefaultView(0).Row
                    rowData.actual_produce_qty = rowBundling.total_produce
                    rowData.actual_good_qty = rowBundling.total_good
                    rowData.actual_reject_qty = rowBundling.total_reject
                    If Me.cboGroupBy.SelectedIndex = 0 Then
                        tblTotal0.DefaultView.RowFilter = "[identifier jo] ='" & rowData.identifier_jo & "'"
                        If tblTotal0.DefaultView.Count > 0 Then
                            rowTotal0 = tblTotal0.DefaultView(0).Row
                            rowTotal0.actual_good += rowBundling.total_good
                            rowTotal0.actual_reject += rowBundling.total_reject
                            rowTotal0.actual_produce += rowBundling.total_produce
                        Else
                            rowTotal0 = tblTotal0.NewDataTable6Row
                            rowTotal0.identifier_jo = rowBundling.identifier_jo
                            rowTotal0.actual_good += rowBundling.total_good
                            rowTotal0.actual_reject += rowBundling.total_reject
                            rowTotal0.actual_produce += rowBundling.total_produce
                            tblTotal0.Rows.Add(rowTotal0)
                        End If
                        tblTotal1.DefaultView.RowFilter = "[identifier jo] = '" & rowData.identifier_jo & "' and [primary process] = " & rowData.primary_process.ToString
                        If tblTotal1.DefaultView.Count > 0 Then
                            rowTotal1 = tblTotal1.DefaultView(0).Row
                            rowTotal1.actual_good += rowBundling.total_good
                            rowTotal1.actual_reject += rowBundling.total_reject
                            rowTotal1.actual_produce += rowBundling.total_produce
                        Else
                            rowTotal1 = tblTotal1.NewDataTable6Row
                            rowTotal1.identifier_jo = rowData.identifier_jo
                            rowTotal1.primary_process = rowData.primary_process
                            rowTotal1.actual_good += rowBundling.total_good
                            rowTotal1.actual_reject += rowBundling.total_reject
                            rowTotal1.actual_produce += rowBundling.total_produce
                            tblTotal1.Rows.Add(rowTotal1)
                        End If
                    ElseIf Me.cboGroupBy.SelectedIndex = 1 Then
                        tblTotal0.DefaultView.RowFilter = "[primary workunit] =" & rowData.primary_work_unit.ToString
                        If tblTotal0.DefaultView.Count > 0 Then
                            rowTotal0 = tblTotal0.DefaultView(0).Row
                            rowTotal0.actual_good += rowBundling.total_good
                            rowTotal0.actual_reject += rowBundling.total_reject
                            rowTotal0.actual_produce += rowBundling.total_produce
                        Else
                            rowTotal0 = tblTotal0.NewDataTable6Row
                            rowTotal0.primary_workunit = rowData.primary_work_unit
                            rowTotal0.actual_good += rowBundling.total_good
                            rowTotal0.actual_reject += rowBundling.total_reject
                            rowTotal0.actual_produce += rowBundling.total_produce
                            tblTotal0.Rows.Add(rowTotal0)
                        End If
                    End If
                End If
            Next

            tblData.DefaultView.RowFilter = "[Is Header] = 1"
            If tblData.DefaultView.Count > 0 Then
                For Each row As Data.DataRowView In tblData.DefaultView
                    rowData = row.Row
                    If Me.cboGroupBy.SelectedIndex = 0 Then
                        If rowData.Item("Tree Level") = 0 Then
                            tblTotal0.DefaultView.RowFilter = "[identifier jo] = '" & rowData.identifier_jo & "'"
                            If tblTotal0.DefaultView.Count > 0 Then
                                rowTotal0 = tblTotal0.DefaultView(0).Row
                                rowData.actual_good_qty = rowTotal0.actual_good
                                rowData.actual_reject_qty = rowTotal0.actual_reject
                                rowData.actual_produce_qty = rowTotal0.actual_produce
                            End If
                        ElseIf rowData.Item("Tree Level") = 1 Then
                            tblTotal1.DefaultView.RowFilter = "[identifier jo] = '" & rowData.identifier_jo & "' and [primary process] = " & rowData.primary_process.ToString
                            If tblTotal1.DefaultView.Count > 0 Then
                                rowTotal1 = tblTotal1.DefaultView(0).Row
                                rowData.actual_good_qty = rowTotal1.actual_good
                                rowData.actual_reject_qty = rowTotal1.actual_reject
                                rowData.actual_produce_qty = rowTotal1.actual_produce
                            End If
                        End If
                    ElseIf Me.cboGroupBy.SelectedIndex = 1 Then
                        tblTotal0.DefaultView.RowFilter = "[primary workunit] = " & rowData.primary_work_unit.ToString
                        If tblTotal0.DefaultView.Count > 0 Then
                            rowTotal0 = tblTotal0.DefaultView(0).Row
                            rowData.actual_good_qty = rowTotal0.actual_good
                            rowData.actual_reject_qty = rowTotal0.actual_reject
                            rowData.actual_produce_qty = rowTotal0.actual_produce
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Function GetFilterTimeHorizon() As String
        Try
            Return "[plan start time] >= #" & Format(Me.dtpFrom.Value.Date, "MM") & "/" & Format(Me.dtpFrom.Value.Date, "dd") & "/" & Format(Me.dtpFrom.Value.Date, "yyyy") & _
                        "# AND [plan start time] <= #" & Format(Me.dtpTo.Value.Date, "MM") & "/" & Format(Me.dtpTo.Value.Date, "dd") & "/" & Format(Me.dtpTo.Value.Date, "yyyy") & " 23:59:00#"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Function getTooltipText(ByVal vIntervalScale As ClsTrans_Schedule.enumDurationScale, ByVal vFrom As Date, ByVal vColName As String, _
            ByVal vBarIndex As Integer, ByRef vTooltipTitle As String, _
            ByRef job_order_number As String, ByRef wo_no As String, ByRef process_code As String, ByRef work_unit_code As String, _
            ByRef code_no As String, ByRef description As String, ByRef rate_unit_per_hour As Decimal, ByRef qty As Decimal, _
            ByRef unit_code As String, ByRef work_hour As Decimal, ByRef start As Date, ByRef finish As Date, ByRef progress As Decimal, _
            ByRef line_code As String, ByVal vNextRowIndex As Integer) As String
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
                progress = Me.rowData.progress
                line_code = Me.rowData("line code")
                retVal = job_order_number & wo_no & process_code & work_unit_code & code_no & description & rate_unit_per_hour.ToString & qty.ToString & unit_code & work_hour.ToString & start.ToString & finish.ToString & progress.ToString & line_code
            End If
            Return retVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return retVal
        End Try
    End Function

    Function isOnAlternateColumn(ByVal i As Integer) As Boolean
        Try
            If (i >= 1 And i <= 14) Or (i >= 29 And i <= 42) Or (i >= 57 And i <= 70) Or _
                    (i >= 85 And i <= 98) Or (i >= 113 And i <= 126) Or (i >= 141 And i <= 154) Then Return True
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

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
            Dim resetRow As Boolean = False
            Dim ResetRow1 As Boolean = False
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
                            dr.Cells("ColBar" & i).Style.BackColor = Me.alternateBarColor
                        Else
                            dr.Cells("ColBar" & i).Style.BackColor = Color.White
                        End If
                    End If
                    If barValueFrom.Date = Now.Date And resetRow = False Then
                        Me.LineTodayIndex = i
                        resetRow = True
                        dr.Cells("ColBar" & i).Value = line_today
                        dr.Cells("ColBar" & i).Style.Alignment = DataGridViewContentAlignment.MiddleRight
                    End If
                    If barValueFrom.Date = Me.dtpFirmPlanned.Value.Date And ResetRow1 = False Then
                        Me.LineFirmPlannedIndex = i
                        ResetRow1 = True
                        dr.Cells("ColBar" & i).Value = line_firmplanned
                        dr.Cells("ColBar" & i).Style.Alignment = DataGridViewContentAlignment.MiddleRight
                    End If
                    barValueFrom += intervalPerBar
                    barValueTo += intervalPerBar
                End With
            Next
            barProgressCount = Me.objCD.DigitGrouping(Me.rowData.progress / 100 * workBarCount, 0, Core_Module.ClsCekData.EnumTipeRounding.Down)
            If barProgressCount > 0 Then
                For i As Integer = firstBarIndex To firstBarIndex - 1 + barProgressCount
                    If Me.chkShowProgress.Checked Then
                        If i = Me.LineTodayIndex Then
                            dr.Cells("ColBar" & i).Value = gantt_progress_today
                        ElseIf i = Me.LineFirmPlannedIndex Then
                            dr.Cells("ColBar" & i).Value = gantt_progress_firmplanned
                        Else
                            dr.Cells("ColBar" & i).Value = gantt_progress
                        End If
                    Else
                        dr.Cells("ColBar" & i).Value = Nothing
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub reFillTgvHeader(ByVal vRowData As Inventory_Module.DseTrans_Schedule.tbl_sch_barRow, ByVal vRowIndex As Integer, ByVal vShowTime As Boolean)
        Dim dr As DataGridViewRow = Me.tgvSch.Rows(vRowIndex)
        Try
            With dr
                If vShowTime = False And Me.cboGroupBy.SelectedIndex = enumGroup.BY_WORKUNIT And vRowData("Is Header") = 1 Then
                    .Cells(Me.ColLeft_DayOfWeekStart.Index).Value = ""
                    .Cells(Me.ColLeft_StartDate.Index).Value = ""
                    .Cells(Me.ColLeft_StartTime.Index).Value = ""
                    .Cells(Me.ColLeft_DayOfWeekFinish.Index).Value = ""
                    .Cells(Me.ColLeft_FinishDate.Index).Value = ""
                    .Cells(Me.ColLeft_FinishTime.Index).Value = ""
                Else
                    .Cells(Me.ColLeft_DayOfWeekStart.Index).Value = Me.objCD.GetHari_English(Me.rowData.plan_start_time)
                    .Cells(Me.ColLeft_StartDate.Index).Value = Format(Me.rowData.plan_start_time, "dd MMM yy")
                    .Cells(Me.ColLeft_StartTime.Index).Value = Format(Me.rowData.plan_start_time, "HH.mm")
                    .Cells(Me.ColLeft_DayOfWeekFinish.Index).Value = Me.objCD.GetHari_English(Me.rowData.plan_end_time)
                    .Cells(Me.ColLeft_FinishDate.Index).Value = Format(Me.rowData.plan_end_time, "dd MMM yy")
                    .Cells(Me.ColLeft_FinishTime.Index).Value = Format(Me.rowData.plan_end_time, "HH.mm")
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub reloadGanttChart()
        Dim StartTime, EndTime As Date

        Try
            Me.Cursor = Cursors.WaitCursor
            Me.listOfLine.Clear()
            Me.listOfLine = objLine.GetListLinePerUser(Me.UserCode)
            Me.fillProductionLine()
            StartTime = Now
            Me.translateData(Me.cboGroupBy.SelectedIndex)
            'MsgBox("Translate data: " + Me.objCD.DigitGrouping((Now - StartTime).TotalMinutes, 2) + "minutes")
            'StartTime = Now
            Me.fillOtherData(Me.tblData)
            'MsgBox("Fill other data: " + Me.objCD.DigitGrouping((Now - StartTime).TotalMinutes, 2) + "minutes")
            'StartTime = Now
            Me.filltgvSch()
            'MsgBox("Fill tgv: " + Me.objCD.DigitGrouping((Now - StartTime).TotalSeconds, 0) + "seconds")
            If Me.ShowGanttChart = True Then
                Me.generateBarColumns(Me.cboTimeframe.SelectedIndex, Me.dtpBarFrom.Value)
            End If
            Me.CollapsAll()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Public Sub fillProductionLine()
        Dim dtaWU As New DseTrans_ScheduleTableAdapters.tblmaster_workunit1TableAdapter
        Dim tblWU As New DseTrans_Schedule.tblmaster_workunit1DataTable
        Dim rowWU As DseTrans_Schedule.tblmaster_workunit1Row
        Try
            dtaWU.Connection.ConnectionString = Core_Module.GetConstring
            dtaWU.Fill(tblWU)
            For Each rowWU In tblWU
                Me.tblSchBar.DefaultView.RowFilter = "[primary work unit] = " & rowWU.primary.ToString
                For Each DV As Data.DataRowView In Me.tblSchBar.DefaultView
                    Me.rowSch = DV.Row
                    Me.rowSch.primary_line = rowWU.primary_production_line
                    Me.rowSch.line_code = rowWU.code
                Next
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

    Sub SetARGBofJO(ByVal vTBLJOCOlor As DseTrans_ManufacturingOrder.JOcolorDataTable, ByVal vIdentifierJO As String, _
            ByRef vA As Double, ByRef vR As Double, ByRef vG As Double, ByRef vB As Double, ByRef vCode As String, ByRef vDesc As String, _
            ByRef vQty As Decimal, ByRef vUnitCode As String, ByRef vDecimal As Integer)
        Dim rowJOColor As DseTrans_ManufacturingOrder.JOcolorRow
        Try
            vTBLJOCOlor.DefaultView.RowFilter = "Identifier='" & vIdentifierJO & "'"
            If vTBLJOCOlor.DefaultView.Count > 0 Then
                rowJOColor = vTBLJOCOlor.DefaultView(0).Row
                vA = rowJOColor.a
                vR = rowJOColor.r
                vG = rowJOColor.g
                vB = rowJOColor.b
                vCode = rowJOColor.Product_Code
                vDesc = rowJOColor.Description
                vQty = rowJOColor.Target_Qty
                vUnitCode = rowJOColor.Unit_Code
                vDecimal = rowJOColor._Decimal
            Else
                vA = Color.Blue.A
                vR = Color.Blue.R
                vG = Color.Blue.G
                vB = Color.Blue.B
                vCode = ""
                vDesc = ""
                vQty = 0
                vUnitCode = ""
                vDecimal = 2
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub translateData(ByVal vGroup As enumGroup)
        Dim dtaJOcolor As New DseTrans_ManufacturingOrderTableAdapters.JOcolorTableAdapter
        Dim tblJOcolor As New DseTrans_ManufacturingOrder.JOcolorDataTable
        Dim a As Double = Color.Blue.A
        Dim r As Double = Color.Blue.R
        Dim g As Double = Color.Blue.G
        Dim b As Double = Color.Blue.B
        Dim code_no As String = ""
        Dim desc As String = ""
        Dim qty As Decimal = 0
        Dim unit As String = ""
        Dim _decimal As Integer = 2
        Dim listOfJO As New List(Of String)
        Dim listOfWorkUnit As New List(Of String)
        Dim NodeIdentiferLv0 As Integer = -1
        Dim NodeIdentiferLv1 As Integer = -1
        Try
            dtaJOcolor.Connection.ConnectionString = Core_Module.GetConstring
            Me.tblData.Rows.Clear()
            If vGroup = enumGroup.BY_JO Then
                dtaJOcolor.Fill(tblJOcolor, Me.dtpFrom.Value.Date - TimeSpan.FromDays(7), #1/1/2200#)
                Me.ColLeft_Tree.HeaderText = "Job Order"
                Me.tblSchBar.DefaultView.RowFilter = Me.GetFilterLine("primary line")
                If Me.tblSchBar.DefaultView.RowFilter <> "" Then
                    Me.tblSchBar.DefaultView.RowFilter = Me.tblSchBar.DefaultView.RowFilter & " AND " & GetFilterTimeHorizon()
                Else
                    Me.tblSchBar.DefaultView.RowFilter = Me.GetFilterTimeHorizon
                End If
                Me.tblSchBar.DefaultView.Sort = "[plan start time]"
                For i As Integer = 0 To Me.tblSchBar.DefaultView.Count - 1
                    Me.rowSch = Me.tblSchBar.DefaultView(i).Row
                    If Not listOfJO.Contains(Me.rowSch.identifier_jo) Then listOfJO.Add(Me.rowSch.identifier_jo)
                Next
                For Each indentiferJO As String In listOfJO
                    NodeIdentiferLv1 = -1
                    Me.rowData = Me.tblData.Newtbl_sch_barRow
                    NodeIdentiferLv0 += 1
                    With Me.rowData
                        Me.SetARGBofJO(tblJOcolor, indentiferJO, a, r, g, b, code_no, desc, qty, unit, _decimal)
                        .tipe = ClsTrans_Schedule.EnumTipeSchBar.work_time
                        If Me.GetFilterLine("primary line") = "" Then
                            tblSchBar.DefaultView.RowFilter = "[identifier jo]='" & indentiferJO & "'" & " AND " & GetFilterTimeHorizon()
                        Else
                            tblSchBar.DefaultView.RowFilter = "[identifier jo]='" & indentiferJO & "'" & " AND " & GetFilterTimeHorizon() & " AND " & Me.GetFilterLine("primary line")
                        End If
                        tblSchBar.DefaultView.Sort = "plan start time ASC"
                        rowSch = Me.tblSchBar.DefaultView(0).Row
                        .plan_start_time = Me.rowSch.plan_start_time
                        tblSchBar.DefaultView.Sort = "plan end time DESC"
                        rowSch = Me.tblSchBar.DefaultView(0).Row
                        .plan_end_time = Me.rowSch.plan_end_time
                        .primary_mo = -1
                        .identifier_jo = indentiferJO
                        .identifier_wo = ""
                        .primary_process = -1
                        .primary_work_unit = -1
                        .a = a
                        .r = r
                        .g = g
                        .b = b
                        .progress = 0
                        .duration = 0
                        .process_code = ""
                        .mo_no = ""
                        .jo_no = Me.rowSch.jo_no
                        .wo_no = ""
                        .work_unit_code = ""
                        .rate = 0
                        .product_code = code_no
                        .product_desc = desc
                        .qty_produce = qty
                        .decimal_point = _decimal
                        .unit_code = unit
                        .primary_wu_config = rowSch.primary_wu_config
                        .tree = Me.rowSch.jo_no
                        .tree_level = 0
                        .node_identifier_lv_0 = NodeIdentiferLv0
                        .node_identifier_lv_1 = NodeIdentiferLv1
                        .is_header = 1
                        .primary_input_location = -1
                        .input_location = String.Empty
                        .primary_output_location = -1
                        .output_location = String.Empty
                    End With
                    Me.tblData.Rows.Add(Me.rowData)
                    If Me.GetFilterLine("primary line") = "" Then
                        Me.tblSchBar.DefaultView.RowFilter = "[identifier jo]='" & indentiferJO & "'" & " AND " & GetFilterTimeHorizon()
                    Else
                        Me.tblSchBar.DefaultView.RowFilter = "[identifier jo]='" & indentiferJO & "'" & " AND " & GetFilterTimeHorizon() & " and " & Me.GetFilterLine("primary line")
                    End If
                    Me.tblSchBar.DefaultView.Sort = "plan start time ASC"
                    Dim listOfProcess As New List(Of Double)
                    For i As Integer = 0 To Me.tblSchBar.DefaultView.Count - 1
                        Me.rowSch = Me.tblSchBar.DefaultView(i).Row
                        If listOfProcess.Contains(Me.rowSch.primary_process) = False Then listOfProcess.Add(Me.rowSch.primary_process)
                    Next
                    For Each primaryProcess As Double In listOfProcess
                        If Me.GetFilterLine("primary line") = "" Then
                            Me.tblSchBar.DefaultView.RowFilter = "[identifier jo]='" & indentiferJO & "' AND [primary process] = " & primaryProcess.ToString & " AND " & GetFilterTimeHorizon()
                        Else
                            Me.tblSchBar.DefaultView.RowFilter = "[identifier jo]='" & indentiferJO & "' AND [primary process] = " & primaryProcess.ToString & " AND " & GetFilterTimeHorizon() & " and " & Me.GetFilterLine("primary line")
                        End If
                        Me.rowData = Me.tblData.Newtbl_sch_barRow
                        NodeIdentiferLv1 += 1
                        With Me.rowData
                            Me.SetARGBofJO(tblJOcolor, indentiferJO, a, r, g, b, code_no, desc, qty, unit, _decimal)
                            .tipe = ClsTrans_Schedule.EnumTipeSchBar.work_time
                            tblSchBar.DefaultView.Sort = "plan start time ASC"
                            Me.rowSch = Me.tblSchBar.DefaultView(0).Row
                            .plan_start_time = Me.rowSch.plan_start_time
                            Me.tblSchBar.DefaultView.Sort = "plan end time DESC"
                            Me.rowSch = Me.tblSchBar.DefaultView(0).Row
                            .plan_end_time = Me.rowSch.plan_end_time
                            .primary_mo = -1
                            .identifier_jo = indentiferJO
                            .identifier_wo = ""
                            .primary_process = primaryProcess
                            .primary_work_unit = -1
                            .a = a
                            .r = r
                            .g = g
                            .b = b
                            .progress = 0
                            .duration = 0
                            .process_code = Me.rowSch.process_code
                            .mo_no = ""
                            .jo_no = Me.rowSch.jo_no
                            .wo_no = ""
                            .work_unit_code = ""
                            .rate = 0
                            .product_code = ""
                            .product_desc = ""
                            .qty_produce = 0
                            .decimal_point = _decimal
                            .unit_code = unit
                            .primary_wu_config = rowSch.primary_wu_config
                            .tree = Me.rowSch.process_code
                            .tree_level = 1
                            .node_identifier_lv_0 = NodeIdentiferLv0
                            .node_identifier_lv_1 = NodeIdentiferLv1
                            .is_header = 1
                            .primary_input_location = -1
                            .input_location = String.Empty
                            .primary_output_location = -1
                            .output_location = String.Empty
                        End With
                        Me.tblData.Rows.Add(Me.rowData)
                        Me.tblSchBar.DefaultView.Sort = "[wo no]"
                        For i As Integer = 0 To Me.tblSchBar.DefaultView.Count - 1
                            Me.rowSch = Me.tblSchBar.DefaultView(i).Row
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
                                .process_code = Me.rowSch.process_code
                                .mo_no = Me.rowSch.mo_no
                                .jo_no = Me.rowSch.jo_no
                                .wo_no = Me.rowSch.wo_no
                                .work_unit_code = Me.rowSch.work_unit_code
                                .rate = Me.rowSch.rate
                                If Me.rowSch.Isdecimal_pointNull = True Then
                                    .decimal_point = 2
                                Else
                                    .decimal_point = Me.rowSch.decimal_point
                                End If
                                .total_work_hour = rowSch.total_work_hour
                                .qty_produce = rowSch.qty_produce
                                .unit_code = rowSch.unit_code
                                .primary_wu_config = rowSch.primary_wu_config
                                .tree = Me.rowSch.jo_no
                                .primary_line = rowSch.primary_line
                                .line_code = rowSch.line_code
                                .tree_level = 2
                                .primary_input_location = Me.rowSch.primary_input_location
                                .input_location = Me.rowSch.input_location
                                .primary_output_location = Me.rowSch.primary_output_location
                                .output_location = Me.rowSch.output_location
                            End With
                            Me.tblData.Rows.Add(Me.rowData)
                        Next
                    Next
                Next
                Me.CountProgress()
            ElseIf vGroup = enumGroup.BY_WORKUNIT Then
                Me.ColLeft_Tree.HeaderText = "Work Unit"
                Me.tblSchBar.DefaultView.RowFilter = Me.GetFilterLine("primary line")
                If Me.tblSchBar.DefaultView.RowFilter <> "" Then
                    Me.tblSchBar.DefaultView.RowFilter = Me.tblSchBar.DefaultView.RowFilter & " AND " & GetFilterTimeHorizon()
                Else
                    Me.tblSchBar.DefaultView.RowFilter = Me.GetFilterTimeHorizon
                End If
                Me.tblSchBar.DefaultView.Sort = "[work unit code]"
                For i As Integer = 0 To Me.tblSchBar.DefaultView.Count - 1
                    Me.rowSch = Me.tblSchBar.DefaultView(i).Row
                    If Not listOfWorkUnit.Contains(Me.rowSch.primary_work_unit) Then listOfWorkUnit.Add(Me.rowSch.primary_work_unit)
                Next
                For Each primaryWorkUnit As Double In listOfWorkUnit
                    NodeIdentiferLv1 = -1
                    Dim filterLine As String = Me.GetFilterLine("primary line")
                    If filterLine = "" Then
                        Me.tblSchBar.DefaultView.RowFilter = "[primary work unit]='" & primaryWorkUnit & "'" & " AND " & GetFilterTimeHorizon()
                    Else
                        Me.tblSchBar.DefaultView.RowFilter = "[primary work unit]='" & primaryWorkUnit & "' AND " & filterLine & " AND " & GetFilterTimeHorizon()
                    End If

                    Me.rowData = Me.tblData.Newtbl_sch_barRow
                    NodeIdentiferLv0 += 1
                    With Me.rowData
                        .tipe = ClsTrans_Schedule.EnumTipeSchBar.work_time
                        tblSchBar.DefaultView.Sort = "plan start time ASC"
                        Me.rowSch = Me.tblSchBar.DefaultView(0).Row
                        .plan_start_time = Me.rowSch.plan_start_time
                        Me.tblSchBar.DefaultView.Sort = "plan end time DESC"
                        Me.rowSch = Me.tblSchBar.DefaultView(0).Row
                        .plan_end_time = Me.rowSch.plan_end_time
                        .primary_mo = -1
                        .identifier_jo = ""
                        .identifier_wo = ""
                        .primary_process = -1
                        .primary_work_unit = primaryWorkUnit
                        .a = Color.Blue.A
                        .r = Color.Blue.R
                        .g = Color.Blue.G
                        .b = Color.Blue.B
                        .progress = 0
                        .duration = 0
                        .process_code = ""
                        .mo_no = ""
                        .jo_no = ""
                        .wo_no = ""
                        .work_unit_code = ""
                        .rate = 0
                        If Me.rowSch.Isdecimal_pointNull = True Then
                            .decimal_point = 2
                        Else
                            .decimal_point = Me.rowSch.decimal_point
                        End If
                        .tree = Me.rowSch.jo_no
                        .primary_line = rowSch.primary_line
                        .tree = Me.rowSch.work_unit_code & " : " & Me.rowSch.process_code
                        .tree_level = 0
                        .node_identifier_lv_0 = NodeIdentiferLv0
                        .node_identifier_lv_1 = NodeIdentiferLv1
                        .is_header = 1
                        .primary_input_location = Me.rowSch.primary_input_location
                        .input_location = Me.rowSch.input_location
                        .primary_output_location = Me.rowSch.primary_output_location
                        .output_location = Me.rowSch.output_location
                    End With
                    Me.tblData.Rows.Add(Me.rowData)
                    Me.tblSchBar.DefaultView.Sort = "[plan start time]"
                    For i As Integer = 0 To Me.tblSchBar.DefaultView.Count - 1
                        Me.rowSch = tblSchBar.DefaultView(i).Row
                        Me.rowData = tblData.Newtbl_sch_barRow
                        NodeIdentiferLv1 += 1
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
                            .a = rowSch.a
                            .r = rowSch.r
                            .g = rowSch.g
                            .b = rowSch.b
                            .progress = Me.rowSch.progress
                            .duration = Me.rowSch.duration
                            .process_code = Me.rowSch.process_code
                            .mo_no = Me.rowSch.mo_no
                            .jo_no = Me.rowSch.jo_no
                            .wo_no = Me.rowSch.wo_no
                            .work_unit_code = Me.rowSch.work_unit_code
                            .rate = Me.rowSch.rate
                            If Me.rowSch.Isdecimal_pointNull = True Then
                                .decimal_point = 2
                            Else
                                .decimal_point = Me.rowSch.decimal_point
                            End If
                            .total_work_hour = rowSch.total_work_hour
                            .qty_produce = rowSch.qty_produce
                            .unit_code = rowSch.unit_code
                            .primary_line = rowSch.primary_line
                            .line_code = rowSch.line_code
                            .tree = Me.rowSch.jo_no & " : " & Me.rowSch.wo_no
                            .tree_level = 1
                            .node_identifier_lv_0 = NodeIdentiferLv0
                            .node_identifier_lv_1 = NodeIdentiferLv1
                            .primary_input_location = Me.rowSch.primary_input_location
                            .input_location = Me.rowSch.input_location
                            .primary_output_location = Me.rowSch.primary_output_location
                            .output_location = Me.rowSch.output_location
                        End With
                        Me.tblData.Rows.Add(Me.rowData)
                    Next
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Dim tblProgress As New DseTrans_Schedule.DataTable4DataTable
    Dim rowProgress As DseTrans_Schedule.DataTable4Row
    Dim tblProgress2 As New DseTrans_Schedule.DataTable4DataTable
    Dim rowProgress2 As DseTrans_Schedule.DataTable4Row
    Sub CountProgress()
        Try
            ''hitung progress pada header process
            tblProgress.Clear()
            For Each rowData In tblData
                If rowData.is_header = 0 Then
                    tblProgress.DefaultView.RowFilter = "[identifier JO] = '" & rowData.identifier_jo & "' AND [primary process] = " & rowData.primary_process.ToString
                    If tblProgress.DefaultView.Count = 0 Then
                        rowProgress = tblProgress.NewDataTable4Row
                        rowProgress.identifier_JO = rowData.identifier_jo
                        rowProgress.primary_process = rowData.primary_process
                        rowProgress.target_qty = rowData.qty_produce
                        rowProgress.progress = rowData.progress
                        rowProgress.actual_qty = rowData.qty_produce * rowData.progress / 100
                        rowProgress.WH = rowData.total_work_hour
                        tblProgress.Rows.Add(rowProgress)
                    Else
                        rowProgress = tblProgress.DefaultView(0).Row
                        rowProgress.target_qty += rowData.qty_produce
                        rowProgress.actual_qty += rowData.qty_produce * rowData.progress / 100
                        rowProgress.WH += rowData.total_work_hour
                        rowProgress.progress = Me.objCD.DigitGrouping(rowProgress.actual_qty / rowProgress.target_qty * 100, 2)
                    End If
                End If
            Next
            For Each rowData In tblData
                If rowData.is_header = 1 Then
                    tblProgress.DefaultView.RowFilter = "[identifier JO] = '" & rowData.identifier_jo & "' AND [primary process] = " & rowData.primary_process.ToString
                    If tblProgress.DefaultView.Count > 0 Then
                        rowProgress = tblProgress.DefaultView(0).Row
                        rowData.progress = rowProgress.progress
                    End If
                End If
            Next

            ''hitung progress pada header JO
            tblProgress2.Clear()
            For Each rowProgress In tblProgress
                tblProgress2.DefaultView.RowFilter = "[identifier JO] = '" & rowProgress.identifier_JO & "'"
                If tblProgress2.DefaultView.Count = 0 Then
                    rowProgress2 = tblProgress2.NewDataTable4Row
                    rowProgress2.identifier_JO = rowProgress.identifier_JO
                    rowProgress2.target_qty = rowProgress.target_qty
                    rowProgress2.actual_qty = rowProgress.actual_qty
                    rowProgress2.progress = rowProgress2.actual_qty / rowProgress2.target_qty * 100
                    rowProgress2.WH = rowProgress.WH
                    tblProgress2.Rows.Add(rowProgress2)
                Else
                    rowProgress2 = tblProgress2.DefaultView(0).Row
                    rowProgress2.target_qty += rowProgress.target_qty
                    rowProgress2.actual_qty += rowProgress.actual_qty
                    rowProgress2.WH += rowProgress.WH
                    rowProgress2.progress = rowProgress2.actual_qty / rowProgress2.target_qty * 100
                End If
            Next
            For Each Me.rowData In Me.tblData
                If Me.rowData.is_header = 1 And Me.rowData.tree_level = 0 Then
                    tblProgress2.DefaultView.RowFilter = "[identifier JO] = '" & rowData.identifier_jo & "'"
                    If tblProgress2.DefaultView.Count > 0 Then
                        rowProgress2 = tblProgress2.DefaultView(0).Row
                        rowData.progress = rowProgress2.progress
                        rowData.total_work_hour = rowProgress2.WH
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
        End Try
    End Sub

    Private Sub cboTimeframe_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboTimeframe.SelectedIndexChanged
        Try
            Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
            Me.generateBarColumns(Me.cboTimeframe.SelectedIndex, Me.dtpBarFrom.Value)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = System.Windows.Forms.Cursors.Default
        End Try
    End Sub

    Private Sub dtpBarFrom_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpBarFrom.ValueChanged
        'Try
        '    If Me.refreshBar Then
        '        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        '        Me.reloadGanttChart()
        '    Else
        '        Me.refreshBar = True
        '    End If
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'Finally
        '    Me.Cursor = System.Windows.Forms.Cursors.Default
        'End Try
    End Sub

    Private Sub dgvBarScroll(ByVal sender As Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles dgvBar.Scroll
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

    Private Sub tgvSch_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles tgvSch.CellBeginEdit
        Dim DummyTime As Date
        Dim IdentifierJO As String = ""
        Dim IdentifierWO As String = ""
        Dim PrimaryWU As Double = -1

        Try
            If Me.ToolStripStatusLabel1.Text.ToLower.Contains("full control") = False Then
                e.Cancel = True
                Exit Sub
            End If
            If e.ColumnIndex = Me.ColLeft_LocationInput.Index Or e.ColumnIndex = Me.ColLeft_LocationOutput.Index Then
                Me.Buffer = Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value
            End If
            If e.ColumnIndex = Me.ColLeft_StartDate.Index Then
                If (Me.tgvSch.Item(Me.ColLeft_NodeLevel.Index, e.RowIndex).Value = 2 And Me.cboGroupBy.SelectedIndex = 0) Or _
                (Me.tgvSch.Item(Me.ColLeft_NodeLevel.Index, e.RowIndex).Value = 1 And Me.cboGroupBy.SelectedIndex = 1) Then
                    If CDate(Me.tgvSch.Item(Me.ColLeft_StartDate.Index, e.RowIndex).Value).Date <= Me.ObjSch.GetFirmPlan.Date Then
                        MsgBox("The specified schedule start before firm plan, you can't change the schedule.", MsgBoxStyle.Information, "Information")
                        e.Cancel = True
                        Exit Sub
                    End If

                    IdentifierJO = Me.tgvSch.Item(Me.ColLeft_IdentifierJO.Index, e.RowIndex).Value
                    IdentifierWO = Me.tgvSch.Item(Me.ColLeft_IdentifierWO.Index, e.RowIndex).Value
                    PrimaryWU = Me.tgvSch.Item(Me.ColLeft_PrimaryWorkUnit.Index, e.RowIndex).Value

                    'bukan dari bagian jo yang sudah selesai
                    If ObjMO.ApakahJOSudahSelesai(IdentifierJO) = True Then
                        MsgBox("The specified job order is closed, you can't change the schedule.", MsgBoxStyle.Information, "Information")
                        e.Cancel = True
                        Exit Sub
                    End If
                    'belum ada log untuk item sch tsb - dicancel dulu
                    'If ObjLog.GetJml_ByIdentifierWO_PrimaryWU(IdentifierWO, PrimaryWU) > 0 Then
                    '    MsgBox("There is 1 or more production log refering to this schedule, you can't change the schedule.", MsgBoxStyle.Information, "Information")
                    '    e.Cancel = True
                    '    Exit Sub
                    'End If

                    Me.Buffer = Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value
                    DummyTime = CDate(Me.tgvSch.Item(Me.ColLeft_StartDate.Index, e.RowIndex).Value).Date + _
                    TimeSpan.FromMinutes(Me.objCD.Convert_ToMinutes(Me.tgvSch.Item(Me.ColLeft_StartTime.Index, e.RowIndex).Value))
                    Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value = Format(DummyTime, "dd MMM yyyy HH.mm")
                    Me.BufferTime = DummyTime
                Else
                    e.Cancel = True
                End If
            ElseIf e.ColumnIndex = Me.ColLeft_FinishDate.Index Then
                'rowTGV = Me.tgvSch.Nodes(e.RowIndex)
                If (Me.tgvSch.Item(Me.ColLeft_NodeLevel.Index, e.RowIndex).Value = 2 And Me.cboGroupBy.SelectedIndex = 0) Or _
                (Me.tgvSch.Item(Me.ColLeft_NodeLevel.Index, e.RowIndex).Value = 1 And Me.cboGroupBy.SelectedIndex = 1) Then
                    If CDate(Me.tgvSch.Item(Me.ColLeft_StartDate.Index, e.RowIndex).Value).Date <= Me.ObjSch.GetFirmPlan.Date Then
                        e.Cancel = True
                        Exit Sub
                    End If

                    IdentifierJO = Me.tgvSch.Item(Me.ColLeft_IdentifierJO.Index, e.RowIndex).Value
                    IdentifierWO = Me.tgvSch.Item(Me.ColLeft_IdentifierWO.Index, e.RowIndex).Value
                    PrimaryWU = Me.tgvSch.Item(Me.ColLeft_PrimaryWorkUnit.Index, e.RowIndex).Value

                    'bukan dari bagian jo yang sudah selesai
                    If ObjMO.ApakahJOSudahSelesai(IdentifierJO) = True Then
                        MsgBox("The specified job order is closed, you can't change the schedule.", MsgBoxStyle.Information, "Information")
                        e.Cancel = True
                        Exit Sub
                    End If
                    'belum ada log untuk item sch tsb
                    If ObjLog.GetJml_ByIdentifierWO_PrimaryWU(IdentifierWO, PrimaryWU) > 0 Then
                        MsgBox("There is 1 or more production log refering to this schedule, you can't change the schedule.", MsgBoxStyle.Information, "Information")
                        e.Cancel = True
                        Exit Sub
                    End If

                    Me.Buffer = Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value
                    DummyTime = CDate(Me.tgvSch.Item(Me.ColLeft_FinishDate.Index, e.RowIndex).Value).Date + _
                    TimeSpan.FromMinutes(Me.objCD.Convert_ToMinutes(Me.tgvSch.Item(Me.ColLeft_FinishTime.Index, e.RowIndex).Value))
                    Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value = Format(DummyTime, "dd MMM yyyy HH.mm")
                    Me.BufferTime = DummyTime
                Else
                    e.Cancel = True
                End If
            ElseIf e.ColumnIndex = Me.ColLeft_ProductionRate.Index Or e.ColumnIndex = Me.ColLeft_TotalWH.Index Then
                If (Me.tgvSch.Item(Me.ColLeft_NodeLevel.Index, e.RowIndex).Value = 2 And Me.cboGroupBy.SelectedIndex = 0) Or _
                (Me.tgvSch.Item(Me.ColLeft_NodeLevel.Index, e.RowIndex).Value = 1 And Me.cboGroupBy.SelectedIndex = 1) Then
                    If CDate(Me.tgvSch.Item(Me.ColLeft_StartDate.Index, e.RowIndex).Value).Date <= Me.ObjSch.GetFirmPlan.Date Then
                        e.Cancel = True
                        Exit Sub
                    End If

                    Me.Buffer = Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            ElseIf e.ColumnIndex = Me.ColLeft_LocationInput.Index Or e.ColumnIndex = Me.ColLeft_LocationOutput.Index Then
                Me.Buffer = Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UpdateBarTime(ByVal vIdentifierWO As String, ByVal vPrimaryWorkUnit As Double, ByVal vPlanStartTime As Date, ByVal vPlanEndTime As Date)
        Try
            Me.tblData.DefaultView.RowFilter = "[identifier wo] = '" & vIdentifierWO & "' AND [primary work unit] = " & vPrimaryWorkUnit.ToString
            If Me.tblData.DefaultView.Count > 0 Then
                Me.rowData = Me.tblData.DefaultView(0).Row
                Me.rowData.plan_start_time = vPlanStartTime
                Me.rowData.plan_end_time = vPlanEndTime
                Me.reDrawBarHeader(Me.rowData, True)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub tgvSch_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles tgvSch.CellEndEdit
        Dim TimeComponent As String()
        Dim dd, MMM, yyyy, HHmm As String
        Dim FormatBetul As Boolean = True
        Dim DummyTime As Date = Today
        Dim StartTimeLama, EndTimeLama, EndTimeBaru As Date
        Dim IdentifierWO As String = ""
        Dim IdentifierJO As String = ""
        Dim PrimaryWU As Double = -1
        Dim DTA As New DseTrans_ScheduleTableAdapters.tbltrans_mo_item_wo_item_sch2TableAdapter
        Dim DurasiGeser As Double = 0
        Dim TotalWHBaru As Double = 0
        Dim ProdEfficiency As Decimal = 0
        Dim PlanProdRateNormal As Decimal = 0
        Dim DecimalPoint As Integer = 0
        Dim frmFindLoc As New Core_Module.FrmCariMasterStorageArea
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If e.ColumnIndex = Me.ColLeft_LocationInput.Index Then
                If IsNothing(Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If String.IsNullOrEmpty(Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value = Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.ToUpper
                        If Me.Buffer = Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value Then
                            Exit Sub
                        End If
                        If Me.ObjLocation.ApakahSudahAda(Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value) Then
                            Me.tgvSch.Item(Me.ColLeft_PrimaryLocationInput.Index, e.RowIndex).Value = Me.ObjLocation.GetPrimary(Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value)
                        Else
                            frmFindLoc.txtCari.Text = Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value
                            frmFindLoc.ShowDialog()
                            If frmFindLoc.PrimaryStorageArea > 0 Then
                                Me.tgvSch.Item(Me.ColLeft_PrimaryLocationInput.Index, e.RowIndex).Value = frmFindLoc.PrimaryStorageArea
                                Me.tgvSch.Item(Me.ColLeft_LocationInput.Index, e.RowIndex).Value = Me.ObjLocation.GetStorageAreaCode(frmFindLoc.PrimaryStorageArea)
                            Else
                                Me.tgvSch.Item(Me.ColLeft_LocationInput.Index, e.RowIndex).Value = Me.Buffer
                            End If
                        End If
                        Me.tblSchBar.DefaultView.RowFilter = "[identifier jo] = '" & Me.tgvSch.Item(Me.ColLeft_IdentifierJO.Index, e.RowIndex).Value & _
                        "' and [identifier wo] = '" & Me.tgvSch.Item(Me.ColLeft_IdentifierWO.Index, e.RowIndex).Value & "' and [primary work unit] = " & Me.tgvSch.Item(Me.ColLeft_PrimaryWorkUnit.Index, e.RowIndex).Value.ToString
                        If Me.tblSchBar.DefaultView.Count > 0 Then
                            Me.rowSch = Me.tblSchBar.DefaultView(0).Row
                            Me.rowSch.primary_input_location = Me.tgvSch.Item(Me.ColLeft_PrimaryLocationInput.Index, e.RowIndex).Value
                            Me.rowSch.input_location = Me.tgvSch.Item(Me.ColLeft_LocationInput.Index, e.RowIndex).Value
                        End If
                        Me.tblData.DefaultView.RowFilter = "[identifier jo] = '" & Me.tgvSch.Item(Me.ColLeft_IdentifierJO.Index, e.RowIndex).Value & _
                        "' and [identifier wo] = '" & Me.tgvSch.Item(Me.ColLeft_IdentifierWO.Index, e.RowIndex).Value & "' and [primary work unit] = " & Me.tgvSch.Item(Me.ColLeft_PrimaryWorkUnit.Index, e.RowIndex).Value.ToString
                        If Me.tblData.DefaultView.Count > 0 Then
                            Me.rowSch = Me.tblData.DefaultView(0).Row
                            Me.rowSch.primary_input_location = Me.tgvSch.Item(Me.ColLeft_PrimaryLocationInput.Index, e.RowIndex).Value
                            Me.rowSch.input_location = Me.tgvSch.Item(Me.ColLeft_LocationInput.Index, e.RowIndex).Value
                        End If
                        Me.ObjSch.UpdateInputOutputLocation(Me.tgvSch.Item(Me.ColLeft_IdentifierJO.Index, e.RowIndex).Value, Me.tgvSch.Item(Me.ColLeft_IdentifierWO.Index, e.RowIndex).Value, _
                        Me.tgvSch.Item(Me.ColLeft_PrimaryWorkUnit.Index, e.RowIndex).Value, Me.tgvSch.Item(Me.ColLeft_PrimaryLocationInput.Index, e.RowIndex).Value, Me.tgvSch.Item(Me.ColLeft_PrimaryLocationOutput.Index, e.RowIndex).Value)
                    Else
                        Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    End If
                Else
                    Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                End If
            End If
            If e.ColumnIndex = Me.ColLeft_LocationOutput.Index Then
                If IsNothing(Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If String.IsNullOrEmpty(Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value = Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.ToUpper
                        If Me.Buffer = Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value Then
                            Exit Sub
                        End If
                        If Me.ObjLocation.ApakahSudahAda(Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value) Then
                            Me.tgvSch.Item(Me.ColLeft_PrimaryLocationOutput.Index, e.RowIndex).Value = Me.ObjLocation.GetPrimary(Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value)
                        Else
                            frmFindLoc.txtCari.Text = Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value
                            frmFindLoc.ShowDialog()
                            If frmFindLoc.PrimaryStorageArea > 0 Then
                                Me.tgvSch.Item(Me.ColLeft_PrimaryLocationOutput.Index, e.RowIndex).Value = frmFindLoc.PrimaryStorageArea
                                Me.tgvSch.Item(Me.ColLeft_LocationOutput.Index, e.RowIndex).Value = Me.ObjLocation.GetStorageAreaCode(frmFindLoc.PrimaryStorageArea)
                            Else
                                Me.tgvSch.Item(Me.ColLeft_LocationOutput.Index, e.RowIndex).Value = Me.Buffer
                            End If
                        End If
                        Me.tblSchBar.DefaultView.RowFilter = "[identifier jo] = '" & Me.tgvSch.Item(Me.ColLeft_IdentifierJO.Index, e.RowIndex).Value & _
                        "' and [identifier wo] = '" & Me.tgvSch.Item(Me.ColLeft_IdentifierWO.Index, e.RowIndex).Value & "' and [primary work unit] = " & Me.tgvSch.Item(Me.ColLeft_PrimaryWorkUnit.Index, e.RowIndex).Value.ToString
                        If Me.tblSchBar.DefaultView.Count > 0 Then
                            Me.rowSch = Me.tblSchBar.DefaultView(0).Row
                            Me.rowSch.primary_output_location = Me.tgvSch.Item(Me.ColLeft_PrimaryLocationOutput.Index, e.RowIndex).Value
                            Me.rowSch.output_location = Me.tgvSch.Item(Me.ColLeft_LocationOutput.Index, e.RowIndex).Value
                        End If
                        Me.tblData.DefaultView.RowFilter = "[identifier jo] = '" & Me.tgvSch.Item(Me.ColLeft_IdentifierJO.Index, e.RowIndex).Value & _
                        "' and [identifier wo] = '" & Me.tgvSch.Item(Me.ColLeft_IdentifierWO.Index, e.RowIndex).Value & "' and [primary work unit] = " & Me.tgvSch.Item(Me.ColLeft_PrimaryWorkUnit.Index, e.RowIndex).Value.ToString
                        If Me.tblData.DefaultView.Count > 0 Then
                            Me.rowSch = Me.tblData.DefaultView(0).Row
                            Me.rowSch.primary_output_location = Me.tgvSch.Item(Me.ColLeft_PrimaryLocationOutput.Index, e.RowIndex).Value
                            Me.rowSch.output_location = Me.tgvSch.Item(Me.ColLeft_LocationOutput.Index, e.RowIndex).Value
                        End If
                        Me.ObjSch.UpdateInputOutputLocation(Me.tgvSch.Item(Me.ColLeft_IdentifierJO.Index, e.RowIndex).Value, Me.tgvSch.Item(Me.ColLeft_IdentifierWO.Index, e.RowIndex).Value, _
                        Me.tgvSch.Item(Me.ColLeft_PrimaryWorkUnit.Index, e.RowIndex).Value, Me.tgvSch.Item(Me.ColLeft_PrimaryLocationInput.Index, e.RowIndex).Value, Me.tgvSch.Item(Me.ColLeft_PrimaryLocationOutput.Index, e.RowIndex).Value)
                    Else
                        Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    End If
                Else
                    Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                End If
            End If

            If e.ColumnIndex = Me.ColLeft_StartDate.Index Then
                If IsNothing(Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                    Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                Else
                    TimeComponent = Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Split(" ")
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
                                Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                Exit Sub
                            End If
                        End If
                    End If

                    If FormatBetul = False Then
                        MsgBox("Please specify format with pattern dd MMM yyy HH.mm.", MsgBoxStyle.Information, "Information")
                        Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    Else
                        StartTimeLama = Me.BufferTime
                        EndTimeLama = CDate(Me.tgvSch.Item(Me.ColLeft_FinishDate.Index, e.RowIndex).Value).Date + _
                        TimeSpan.FromMinutes(Me.objCD.Convert_ToMinutes(Me.tgvSch.Item(Me.ColLeft_FinishTime.Index, e.RowIndex).Value))
                        If EndTimeLama = #1/1/2000# Then
                            EndTimeLama = #1/1/2200#
                        End If
                        IdentifierWO = Me.tgvSch.Item(Me.ColLeft_IdentifierWO.Index, e.RowIndex).Value
                        IdentifierJO = Me.tgvSch.Item(Me.ColLeft_IdentifierJO.Index, e.RowIndex).Value
                        PrimaryWU = Me.tgvSch.Item(Me.ColLeft_PrimaryWorkUnit.Index, e.RowIndex).Value

                        For Each Me.rowSch In Me.tblSchBar.Rows
                            If Me.rowSch.primary_work_unit = PrimaryWU Then
                                If Me.rowSch.identifier_wo = IdentifierWO And Me.rowSch.identifier_jo = IdentifierJO Then
                                    If Me.rowSch.plan_start_time = StartTimeLama And Me.rowSch.plan_end_time = EndTimeLama Then
                                        EndTimeBaru = Me.ObjSch.GetEndTime(PrimaryWU, DummyTime, rowSch.total_work_hour, Me.tblSchBar_TanpaItemWO)
                                        If EndTimeBaru = #1/1/2200# Then
                                            MsgBox("Insufficient workshift allocation, please allocate more workshift for " + Me.objWU.GetCode(rowSch.primary_work_unit) + ".", MsgBoxStyle.Information, "Information")
                                            Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                            Exit Sub
                                        End If
                                        If DummyTime <= Me.BufferTime Then
                                            DurasiGeser = Me.ObjSch.GetWorkDuration(DummyTime, Me.BufferTime, Me.rowSch.primary_work_unit, Me.tblSchBar_TanpaItemWO)
                                            DurasiGeser = DurasiGeser * (-1)
                                        Else
                                            DurasiGeser = Me.ObjSch.GetWorkDuration(Me.BufferTime, DummyTime, Me.rowSch.primary_work_unit, Me.tblSchBar_TanpaItemWO)
                                        End If
                                        If Me.ObjSch.ApakahBolehGeserStartTime_FormUtama(Me.rowSch, Me.tblSchBar_TanpaItemWO, Me.tblSchBar, _
                                        DummyTime, EndTimeBaru, StartTimeLama, DurasiGeser) = False Then
                                            Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                            Exit Sub
                                        End If

                                        'hitung ulang start time dan end time
                                        Me.rowSch.plan_start_time = DummyTime
                                        Me.rowSch.plan_end_time = Me.ObjSch.GetEndTime(Me.rowSch.primary_work_unit, _
                                        Me.rowSch.plan_start_time, Me.rowSch.total_work_hour, Me.tblSchBar_TanpaItemWO)

                                        'geser yang sebelah kanan dari jadwal
                                        If Me.rowSch.plan_start_time < Me.BufferTime Then
                                            Me.ObjSch.GeserJadwal(Me.rowSch.primary_work_unit, DurasiGeser, Me.tblSchBar_TanpaItemWO, Me.rowSch.identifier_wo, _
                                            Me.rowSch.plan_start_time, rowSch.identifier_jo)
                                        Else
                                            Me.ObjSch.GeserJadwal(Me.rowSch.primary_work_unit, DurasiGeser, Me.tblSchBar_TanpaItemWO, Me.rowSch.identifier_wo, _
                                            Me.BufferTime, rowSch.identifier_jo)
                                        End If

                                        'simpan data ke db
                                        DTA.UpdateQuery_SetJadwalBaru3(Me.rowSch.plan_start_time, Me.rowSch.plan_end_time, Me.rowSch.identifier_jo, Me.rowSch.identifier_wo)

                                        'update item wo
                                        Me.ObjSch.UpdateStartEndTime_ItemWO(rowSch.identifier_jo, Me.rowSch.identifier_wo)

                                        'tampilkan ulang dgv
                                        'Windows.Forms.Cursor.Current = Cursors.WaitCursor
                                        'Me.UpdateBarTime(Me.rowSch.identifier_wo, Me.rowSch.primary_work_unit, Me.rowSch.plan_start_time, Me.rowSch.plan_end_time)
                                        Me.tgvSch.Item(Me.ColLeft_StartDate.Index, e.RowIndex).Value = Format(Me.rowSch.plan_start_time, "dd MMM yyyy")
                                        Me.tgvSch.Item(Me.ColLeft_StartTime.Index, e.RowIndex).Value = Format(Me.rowSch.plan_start_time, "HH.mm")

                                        'reload gantt chart
                                        Me.BackgroundWorker6.RunWorkerAsync()

                                        Exit For
                                    End If
                                End If
                            End If
                        Next
                    End If
                End If
            ElseIf e.ColumnIndex = Me.ColLeft_FinishDate.Index Then
                If IsNothing(Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                    Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                Else
                    TimeComponent = Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Split(" ")
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
                                Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                Exit Sub
                            End If
                        End If

                        If FormatBetul = False Then
                            MsgBox("Please specify format with pattern dd MMM yyy HH.mm.", MsgBoxStyle.Information, "Information")
                            Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                            Exit Sub
                        Else
                            StartTimeLama = CDate(Me.tgvSch.Item(Me.ColLeft_StartDate.Index, e.RowIndex).Value).Date + _
                            TimeSpan.FromMinutes(Me.objCD.Convert_ToMinutes(Me.tgvSch.Item(Me.ColLeft_StartTime.Index, e.RowIndex).Value))
                            EndTimeLama = Me.BufferTime
                            IdentifierWO = Me.tgvSch.Item(Me.ColLeft_IdentifierWO.Index, e.RowIndex).Value
                            IdentifierJO = Me.tgvSch.Item(Me.ColLeft_IdentifierJO.Index, e.RowIndex).Value
                            PrimaryWU = Me.tgvSch.Item(Me.ColLeft_PrimaryWorkUnit.Index, e.RowIndex).Value

                            For Each Me.rowSch In Me.tblSchBar.Rows
                                If Me.rowSch.primary_work_unit = PrimaryWU Then
                                    If Me.rowSch.identifier_wo = IdentifierWO And Me.rowSch.identifier_jo = IdentifierJO Then
                                        If Me.rowSch.plan_start_time = StartTimeLama And Me.rowSch.plan_end_time = EndTimeLama Then
                                            If DummyTime <= Me.BufferTime Then
                                                DurasiGeser = Me.ObjSch.GetWorkDuration(DummyTime, Me.BufferTime, Me.rowSch.primary_work_unit, Me.tblSchBar_TanpaItemWO)
                                                DurasiGeser = DurasiGeser * (-1)
                                            Else
                                                DurasiGeser = Me.ObjSch.GetWorkDuration(Me.BufferTime, DummyTime, Me.rowSch.primary_work_unit, Me.tblSchBar_TanpaItemWO)
                                            End If

                                            If Me.ObjSch.ApakahBolehGeserEndTime_FormUtama(Me.rowSch, Me.tblSchBar_TanpaItemWO, Me.tblSchBar, _
                                            DummyTime, EndTimeLama, StartTimeLama, DurasiGeser) = False Then
                                                Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                                Exit Sub
                                            End If

                                            DecimalPoint = Me.ObjSch.GetDecimalPoint(IdentifierJO, IdentifierWO)
                                            ProdEfficiency = Me.ObjSch.GetProdEfficiency(IdentifierJO, IdentifierWO)
                                            Me.rowSch.plan_end_time = DummyTime
                                            Me.rowSch.total_work_hour = Me.ObjSch.GetWorkDuration(Me.rowSch.plan_start_time, Me.rowSch.plan_end_time, Me.rowSch.primary_work_unit, Me.tblSchBar_TanpaItemWO)
                                            Me.rowSch.rate = Math.Round((Me.rowSch.qty_produce * 60) / Me.rowSch.total_work_hour, DecimalPoint, MidpointRounding.AwayFromZero)
                                            PlanProdRateNormal = Math.Round((Me.rowSch.rate * 100) / ProdEfficiency, DecimalPoint, MidpointRounding.AwayFromZero)

                                            Me.ObjSch.GeserJadwal(Me.rowSch.primary_work_unit, DurasiGeser, Me.tblSchBar_TanpaItemWO, Me.rowSch.identifier_wo, _
                                            Me.rowSch.plan_start_time, rowSch.identifier_jo)

                                            'simpan data ke db
                                            DTA.UpdateQuery_SetJadwalBaru2(Me.rowSch.plan_start_time, Me.rowSch.plan_end_time, Me.rowSch.total_work_hour, PlanProdRateNormal, _
                                            Me.rowSch.rate, Me.rowSch.identifier_wo, Me.rowSch.primary_work_unit, Me.rowSch.identifier_jo)

                                            'update item wo
                                            Me.ObjSch.UpdateStartEndTime_ItemWO2(rowSch.identifier_jo, Me.rowSch.identifier_wo)

                                            'reload gantt chart
                                            'tampilkan ulang dgv
                                            'Windows.Forms.Cursor.Current = Cursors.WaitCursor
                                            'Me.UpdateBarTime(Me.rowSch.identifier_wo, Me.rowSch.primary_work_unit, Me.rowSch.plan_start_time, Me.rowSch.plan_end_time)
                                            Me.tgvSch.Item(Me.ColLeft_StartDate.Index, e.RowIndex).Value = Format(Me.rowSch.plan_start_time, "dd MMM yyyy")
                                            Me.tgvSch.Item(Me.ColLeft_StartTime.Index, e.RowIndex).Value = Format(Me.rowSch.plan_start_time, "HH.mm")

                                            'reload gantt chart
                                            Me.BackgroundWorker6.RunWorkerAsync()

                                            Exit For
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    End If
                End If
            ElseIf e.ColumnIndex = Me.ColLeft_ProductionRate.Index Then
                If IsNothing(Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                    Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                Else
                    IdentifierWO = Me.tgvSch.Item(Me.ColLeft_IdentifierWO.Index, e.RowIndex).Value
                    IdentifierJO = Me.tgvSch.Item(Me.ColLeft_IdentifierJO.Index, e.RowIndex).Value
                    PrimaryWU = Me.tgvSch.Item(Me.ColLeft_PrimaryWorkUnit.Index, e.RowIndex).Value

                    'pastikan angka
                    If IsNumeric(Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        MsgBox("Please specify a positive number", MsgBoxStyle.Information, "Information")
                        Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    'digit grouping
                    DecimalPoint = Me.ObjSch.GetDecimalPoint(IdentifierJO, IdentifierWO)
                    Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value = Me.objCD.DigitGrouping(Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value, _
                    DecimalPoint)
                    'pastikan positif
                    If Me.objCD.CekBilanganPositif(Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        MsgBox("Please specify a positive number", MsgBoxStyle.Information, "Information")
                        Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If

                    For Each Me.rowSch In Me.tblSchBar.Rows
                        If Me.rowSch.primary_work_unit = PrimaryWU Then
                            If Me.rowSch.identifier_wo = IdentifierWO And Me.rowSch.identifier_jo = IdentifierJO Then
                                TotalWHBaru = Math.Ceiling(CDec(Me.rowSch.qty_produce * 60 / Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value))
                                DurasiGeser = TotalWHBaru - Me.rowSch.total_work_hour

                                If Me.ObjSch.ApakahBolehUbahRate_FormUtama(Me.rowSch, Me.tblSchBar_TanpaItemWO, DurasiGeser, _
                                Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                                    Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Exit Sub
                                Else
                                    Me.rowSch.plan_end_time = Me.ObjSch.GetEndTime(Me.rowSch.primary_work_unit, _
                                    Me.rowSch.plan_start_time, Me.rowSch.total_work_hour, Me.tblSchBar_TanpaItemWO)

                                    ProdEfficiency = Me.ObjSch.GetProdEfficiency(IdentifierJO, IdentifierWO)
                                    Me.rowSch.total_work_hour = TotalWHBaru
                                    Me.rowSch.plan_end_time = Me.ObjSch.GetEndTime(Me.rowSch.primary_work_unit, Me.rowSch.plan_start_time, TotalWHBaru, Me.tblSchBar_TanpaItemWO)
                                    Me.rowSch.rate = Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value
                                    PlanProdRateNormal = Math.Round((Me.rowSch.rate * 100) / ProdEfficiency, DecimalPoint, MidpointRounding.AwayFromZero)

                                    Me.ObjSch.GeserJadwal(Me.rowSch.primary_work_unit, DurasiGeser, Me.tblSchBar_TanpaItemWO, Me.rowSch.identifier_wo, _
                                    Me.rowSch.plan_start_time, rowSch.identifier_jo)

                                    'simpan data ke db
                                    DTA.UpdateQuery_SetJadwalBaru2(Me.rowSch.plan_start_time, Me.rowSch.plan_end_time, Me.rowSch.total_work_hour, PlanProdRateNormal, _
                                    Me.rowSch.rate, Me.rowSch.identifier_wo, Me.rowSch.primary_work_unit, Me.rowSch.identifier_jo)

                                    'update item wo
                                    Me.ObjSch.UpdateStartEndTime_ItemWO2(rowSch.identifier_jo, Me.rowSch.identifier_wo)

                                    'reload gantt chart
                                    Me.BackgroundWorker6.RunWorkerAsync()
                                End If
                            End If
                        End If
                    Next
                End If
            ElseIf e.ColumnIndex = Me.ColLeft_TotalWH.Index Then
                If IsNothing(Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                    Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                Else
                    IdentifierWO = Me.tgvSch.Item(Me.ColLeft_IdentifierWO.Index, e.RowIndex).Value
                    IdentifierJO = Me.tgvSch.Item(Me.ColLeft_IdentifierJO.Index, e.RowIndex).Value
                    PrimaryWU = Me.tgvSch.Item(Me.ColLeft_PrimaryWorkUnit.Index, e.RowIndex).Value

                    'pastikan time string
                    If Me.objCD.Is_Time_String(Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value, True, False) = False Then
                        MsgBox("Please use time string with HH.mm format", MsgBoxStyle.Information, "Information")
                        Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If

                    For Each Me.rowSch In Me.tblSchBar.Rows
                        If Me.rowSch.primary_work_unit = PrimaryWU Then
                            If Me.rowSch.identifier_wo = IdentifierWO And Me.rowSch.identifier_jo = IdentifierJO Then
                                TotalWHBaru = Me.objCD.Convert_ToMinutes(Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value)
                                DurasiGeser = TotalWHBaru - Me.rowSch.total_work_hour

                                If Me.ObjSch.ApakahBolehUbahWH_FormUtama(Me.rowSch, Me.tblSchBar_TanpaItemWO, DurasiGeser) = False Then
                                    Me.tgvSch.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Exit Sub
                                Else
                                    Me.rowSch.total_work_hour = TotalWHBaru
                                    Me.rowSch.plan_end_time = Me.ObjSch.GetEndTime(Me.rowSch.primary_work_unit, _
                                    Me.rowSch.plan_start_time, Me.rowSch.total_work_hour, Me.tblSchBar_TanpaItemWO)

                                    ProdEfficiency = Me.ObjSch.GetProdEfficiency(IdentifierJO, IdentifierWO)
                                    Me.rowSch.rate = Me.objCD.DigitGrouping((Me.rowSch.qty_produce * 60) / Me.rowSch.total_work_hour, DecimalPoint, Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)
                                    PlanProdRateNormal = Math.Round((Me.rowSch.rate * 100) / ProdEfficiency, DecimalPoint, MidpointRounding.AwayFromZero)

                                    Me.ObjSch.GeserJadwal(Me.rowSch.primary_work_unit, DurasiGeser, Me.tblSchBar_TanpaItemWO, Me.rowSch.identifier_wo, _
                                    Me.rowSch.plan_start_time, rowSch.identifier_jo)

                                    'simpan data ke db
                                    DTA.UpdateQuery_SetJadwalBaru2(Me.rowSch.plan_start_time, Me.rowSch.plan_end_time, Me.rowSch.total_work_hour, PlanProdRateNormal, _
                                    Me.rowSch.rate, Me.rowSch.identifier_wo, Me.rowSch.primary_work_unit, Me.rowSch.identifier_jo)

                                    'update item wo
                                    Me.ObjSch.UpdateStartEndTime_ItemWO2(rowSch.identifier_jo, Me.rowSch.identifier_wo)

                                    'reload gantt chart
                                    Me.BackgroundWorker6.RunWorkerAsync()
                                End If
                            End If
                        End If
                    Next
                End If

            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.tgvSch.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect
        End Try
    End Sub

    Private Sub tgvSch_CellEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles tgvSch.CellEnter
        Try
            If e.ColumnIndex = Me.ColLeft_Tree.Index Then
                Me.mNodeLevel = Me.tgvSch.Item(Me.ColLeft_NodeLevel.Index, e.RowIndex).Value
                Me.mNodeLevel0 = Me.tgvSch.Item(Me.ColLeft_NodeLv0Identifier.Index, e.RowIndex).Value
                Me.mNodeLevel1 = Me.tgvSch.Item(Me.ColLeft_NodeLv1Identifier.Index, e.RowIndex).Value
            Else
                Me.mNodeLevel = "-1"
                Me.mNodeLevel0 = ""
                Me.mNodeLevel1 = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub 

    Private Sub tgvSch_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tgvSch.KeyDown
        Try
            If Me.mNodeLevel = 0 And String.IsNullOrEmpty(Me.mNodeLevel0) Then Exit Sub
            If Me.mNodeLevel = 1 And String.IsNullOrEmpty(Me.mNodeLevel1) Then Exit Sub

            If e.KeyCode = Keys.F And e.Modifiers = Keys.Control Then
                Dim frmFind As New FrmTrans_ProductionSchedule_Find
                frmFind.TBLSchBar = Me.tblData.Copy
                frmFind.GroupBy = Me.cboGroupBy.SelectedIndex
                If frmFind.ShowDialog = Windows.Forms.DialogResult.OK Then
                    Me.Find(frmFind.PrimaryMO, frmFind.IdentifierJO, frmFind.IdentifierWO, frmFind.PrimaryProcess, frmFind.PrimaryWorkunit)
                End If
            End If

            If Me.mNodeLevel = 0 Then
                If e.KeyCode = Keys.Left Then
                    tgvSch.Nodes(Me.mNodeLevel0).Collapse()
                ElseIf e.KeyCode = Keys.Right Then
                    tgvSch.Nodes(Me.mNodeLevel0).Expand()
                End If
            ElseIf Me.mNodeLevel = 1 Then
                If e.KeyCode = Keys.Left Then
                    tgvSch.Nodes(Me.mNodeLevel0).Nodes(Me.mNodeLevel1).Collapse()
                ElseIf e.KeyCode = Keys.Right Then
                    tgvSch.Nodes(Me.mNodeLevel0).Nodes(Me.mNodeLevel1).Expand()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub tgvSch_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles tgvSch.MouseClick
        Dim rowDGV As DataGridViewRow
        Dim rowTGV As AdvancedDataGridView.TreeGridNode
        Try
            If e.Button = Windows.Forms.MouseButtons.Right Then
                Me.tgvSch.ClearSelection()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub tgvSch_NodeCollapsed(ByVal sender As Object, ByVal e As AdvancedDataGridView.CollapsedEventArgs) Handles tgvSch.NodeCollapsed
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            If Me.ShowGanttChart = True Then
                If Me.programmaticExpandCollaps = False Then
                    If e.Node.RowIndex > -1 Then
                        Me.tblData.DefaultView.RowFilter = "[Tree Level] = " & e.Node.DataGridView.Item(Me.ColLeft_NodeLevel.Index, e.Node.RowIndex).Value.ToString & _
                        " AND [Node Identifier Lv 0] = '" & e.Node.DataGridView.Item(Me.ColLeft_NodeLv0Identifier.Index, e.Node.RowIndex).Value & "'" & _
                        " AND [Node Identifier Lv 1] = '" & e.Node.DataGridView.Item(Me.ColLeft_NodeLv1Identifier.Index, e.Node.RowIndex).Value & "'"
                        Me.rowData = Me.tblData.DefaultView(0).Row
                        Me.collapsOrExpandBar(e.Node, Me.tblData.Rows.IndexOf(Me.rowData))
                        Me.reFillTgvHeader(Me.rowData, e.Node.RowIndex, True)
                        Me.reDrawBarHeader(Me.rowData, True)
                    End If
                    If Me.ExpandAll = False Then
                        If Not Me.BackgroundWorker5.IsBusy Then
                            Me.BackgroundWorker5.RunWorkerAsync()
                        End If
                    End If
                Else
                    Me.programmaticExpandCollaps = False
                End If
            End If

            Me.JmlNodeExpand -= 1
            Me.setColumnVisibility_1()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Dim ExpandAll As Boolean = True

    Private Sub tgvSch_NodeExpanded(ByVal sender As Object, ByVal e As AdvancedDataGridView.ExpandedEventArgs) Handles tgvSch.NodeExpanded
        If Me.OnPorgressLoadingTGV = True Then
            Exit Sub
        End If

        Dim DTA_D As New DseTrans_ScheduleTableAdapters.tbltrans_mo_item_woTableAdapter
        Dim row_D As DseTrans_Schedule.tbltrans_mo_item_woRow
        Dim NodeItem, NodeItem1 As AdvancedDataGridView.TreeGridNode

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DTA_D.Connection.ConnectionString = Core_Module.GetConstring

            'ambil ulang product code, wo number, product desc + masukin ke cell-nya
            If Me.EnumGroupBy = enumGroup.BY_JO Then
                If e.Node.DataGridView.Item(Me.ColLeft_NodeLevel.Index, e.Node.RowIndex).Value.ToString = "1" Then
                    'hanya jalan klo: group by jo + yang di expand adalah prosesnya
                    Me.TBL_D.Rows.Clear()
                    DTA_D.FillBy_IdentifierJO_PrimaryProses(Me.TBL_D, _
                    e.Node.DataGridView.Item(Me.ColLeft_IdentifierJO.Index, e.Node.RowIndex).Value.ToString(), _
                    e.Node.DataGridView.Item(Me.ColLeft_PrimaryProcess.Index, e.Node.RowIndex).Value.ToString())
                    For Each NodeItem In e.Node.Nodes
                        For Each row_D In TBL_D.Rows
                            If NodeItem.Cells(Me.ColLeft_IdentifierWO.Index).Value = row_D.Identifier_WO Then
                                With NodeItem
                                    .Cells(Me.ColLeft_WONo.Index).Value = row_D.WO_Number
                                    .Cells(Me.ColLeft_CodeNo.Index).Value = row_D.Product_Code
                                    .Cells(Me.ColLeft_Description.Index).Value = row_D.Description
                                End With
                                Exit For
                            End If
                        Next
                    Next
                End If
            Else
                If e.Node.Nodes.Count > 0 Then
                    NodeItem = e.Node.Nodes(0)
                    If NodeItem.Cells(Me.ColLeft_CodeNo.Index).Value.ToString = "" Then
                        For gh As Integer = 0 To e.Node.Nodes.Count - 1
                            NodeItem = e.Node.Nodes(gh)
                            If NodeItem.Cells(Me.ColLeft_CodeNo.Index).Value.ToString = "" Then
                                Me.TBL_D.Rows.Clear()
                                DTA_D.FillBy_IdentifierJO1(Me.TBL_D, _
                                NodeItem.Cells(Me.ColLeft_IdentifierJO.Index).Value.ToString)
                                For fg As Integer = gh To e.Node.Nodes.Count - 1
                                    NodeItem1 = e.Node.Nodes(fg)
                                    For Each row_D In TBL_D.Rows
                                        If row_D.Identifier_JO = NodeItem1.Cells(Me.ColLeft_IdentifierJO.Index).Value And _
                                        row_D.Identifier_WO = NodeItem1.Cells(Me.ColLeft_IdentifierWO.Index).Value Then
                                            With NodeItem1
                                                .Cells(Me.ColLeft_WONo.Index).Value = row_D.WO_Number
                                                .Cells(Me.ColLeft_CodeNo.Index).Value = row_D.Product_Code
                                                .Cells(Me.ColLeft_Description.Index).Value = row_D.Description
                                            End With
                                            Exit For
                                        End If
                                    Next
                                Next
                            End If
                        Next
                    End If
                End If
                Me.ColLeft_WorkUnit.Visible = False
            End If

            If Me.ShowGanttChart = True Then
                If Me.programmaticExpandCollaps = False Then
                    If e.Node.RowIndex > -1 Then
                        Me.tblData.DefaultView.RowFilter = "[Tree Level] = " & e.Node.DataGridView.Item(Me.ColLeft_NodeLevel.Index, e.Node.RowIndex).Value.ToString & _
                        " AND [Node Identifier Lv 0] = '" & e.Node.DataGridView.Item(Me.ColLeft_NodeLv0Identifier.Index, e.Node.RowIndex).Value & "'" & _
                        " AND [Node Identifier Lv 1] = '" & e.Node.DataGridView.Item(Me.ColLeft_NodeLv1Identifier.Index, e.Node.RowIndex).Value & "'"
                        Me.rowData = Me.tblData.DefaultView(0).Row
                        Me.collapsOrExpandBar(e.Node, Me.tblData.Rows.IndexOf(Me.rowData))
                        Me.reFillTgvHeader(Me.rowData, e.Node.RowIndex, True)
                        Me.reDrawBarHeader(Me.rowData, False)
                    End If
                    If Me.ExpandAll = False Then
                        If Not Me.BackgroundWorker5.IsBusy Then
                            Me.BackgroundWorker5.RunWorkerAsync()
                        End If
                    End If
                Else
                    Me.programmaticExpandCollaps = False
                End If
            End If

            Me.JmlNodeExpand += 1
            Me.setColumnVisibility_1()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub tgvSch_Scroll(ByVal sender As Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles tgvSch.Scroll
        Try
            If Me.ShowGanttChart = True Then
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
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

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
        Dim line_code As String = ""
        Try
            objMousePosition = Me.dgvBar.PointToClient(Control.MousePosition)
            objHitTestInfo = Me.dgvBar.HitTest(objMousePosition.X, objMousePosition.Y)
            If objHitTestInfo.ColumnIndex > -1 And objHitTestInfo.RowIndex > -1 Then
                toolTipTextCurrent = Me.getTooltipText(Me.cboTimeframe.SelectedIndex, Me.dtpBarFrom.Value, _
                                        Me.dgvBar.Columns(objHitTestInfo.ColumnIndex).Name, objHitTestInfo.RowIndex, toolTipTitleCurrent, _
                                         job_order_number, wo_no, process_code, work_unit_code, code_no, description, rate_unit_per_hour, qty, unit_code, _
                                         work_hour, start, finish, progress, line_code, objHitTestInfo.RowIndex + 1)
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
                        Me.frmTooltip.progress = progress
                        Me.frmTooltip.line_code = line_code
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

    Private Sub llbColumnSettingSch_LinkClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbColumnSettingSch.LinkClicked
        Dim ij As New FrmTrans_ProductionSchedule_NewSch2_ColumnSetting
        Try
            ij.CalledFromMainGantt = True
            ij.UserCode = Me.UserCode
            Dim oldListofLine As New List(Of Double)
            For Each d As Double In Me.listOfLine
                oldListofLine.Add(d)
            Next
            If ij.ShowDialog() = Windows.Forms.DialogResult.OK Then
                Me.listOfLine.Clear()
                Me.listOfLine = objLine.GetListLinePerUser(Me.UserCode)

                If IsDifferentList(oldListofLine, Me.listOfLine) Then
                    Me.reloadGanttChart()
                    Me.SetLlbFilterText()
                Else
                    Me.setColumnVisibility_1()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetLlbFilterText()
        Try
            Me.llbColumnSettingSch.Text = "Column Settings & Data Filter"
            If Me.listOfLine.Count > 0 Then
                Me.llbColumnSettingSch.Text = "Column Settings & Data Filter: "
                For Each PrimaryLine As Double In Me.listOfLine
                    If Me.llbColumnSettingSch.Text = "Column Settings & Data Filter: " Then
                        Me.llbColumnSettingSch.Text = Me.llbColumnSettingSch.Text + Me.objMasterLine.GetCode(PrimaryLine)
                    Else
                        Me.llbColumnSettingSch.Text = Me.llbColumnSettingSch.Text + ", " + Me.objMasterLine.GetCode(PrimaryLine)
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Function IsDifferentList(ByVal vOldList As List(Of Double), ByVal vNewList As List(Of Double)) As Boolean
        Try
            If vOldList.Count <> vNewList.Count Then
                Return True
            End If
            For Each d As Double In vOldList
                If vNewList.Contains(d) = False Then
                    Return True
                End If
            Next
            For Each d As Double In vNewList
                If vOldList.Contains(d) = False Then
                    Return True
                End If
            Next
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return True
        End Try
    End Function

    Private Sub cboGroupBy_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboGroupBy.SelectedIndexChanged
        Dim ij As New FrmTrans_ProductionSchedule_Progress5

        Try
            Me.EnumGroupBy = Me.cboGroupBy.SelectedIndex
            If Me.cboGroupBy.Items.Count > 0 And Me.cboGroupBy.SelectedIndex > -1 Then
                If Me.cboGroupBy.SelectedIndex = enumGroup.BY_JO Then
                    Me.ColLeft_ProcessCode.Visible = False
                    Me.ColLeft_WONo.Visible = False
                Else
                    Dim ShowProcess As String = Me.objSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowProcess, Me.UserCode)
                    If ShowProcess = "" Or ShowProcess = "1" Then
                        Me.ColLeft_ProcessCode.Visible = True
                    Else
                        Me.ColLeft_ProcessCode.Visible = False
                    End If
                    Dim ShowWO As String = Me.objSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_ShowWO, Me.UserCode)
                    If ShowWO = "" Or ShowWO = "1" Then
                        Me.ColLeft_WONo.Visible = True
                    Else
                        Me.ColLeft_WONo.Visible = False
                    End If
                End If
                'ij.ShowDialog()
                'Me.GroupBox3.Visible = True
                Me.reloadGanttChart()
                'Me.GroupBox3.Visible = False
            End If
            If Me.cboGroupBy.SelectedIndex <> -1 Then
                Me.EnumGroupBy = Me.cboGroupBy.SelectedIndex
            End If
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
            Me.JmlNodeExpand = 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub SetSomeColumnVisibilty()
        Dim JmlNodeExpand As Integer = 0
        Dim NodeItem As AdvancedDataGridView.TreeGridNode

        Try
            Me.ColLeft_LineCode.Visible = False
            Me.ColLeft_WONo.Visible = False
            Me.ColLeft_WorkUnit.Visible = False
            Me.ColLeft_ProductionRate.Visible = False
            Me.ColLeft_LocationInput.Visible = False
            Me.ColLeft_LocationOutput.Visible = False
            Me.ColLeft_ActualGood.Visible = False
            Me.ColLeft_ActualReject.Visible = False
            Me.ColLeft_Actualproduce.Visible = False
            Me.ColLeft_Progress.Visible = False

            'If ShowProgress_ = "1" Then
            '    Me.ColLeft_Progress.Visible = True
            'Else
            '    Me.ColLeft_Progress.Visible = False
            'End If

            If Me.cboGroupBy.SelectedIndex = enumGroup.BY_JO Then
                If ShowCodeNo = "1" Then Me.ColLeft_CodeNo.Visible = True
                If ShowDesc = "1" Then Me.ColLeft_Description.Visible = True
                If ShowQty = "1" Then Me.ColLeft_QtyProduce.Visible = True
                If ShowUnit = "1" Then Me.ColLeft_Unit.Visible = True
                If ShowWH = "1" Then Me.ColLeft_TotalWH.Visible = True
            Else 
                If Me.JmlNodeExpand = 0 Then
                    Me.ColLeft_CodeNo.Visible = False
                    Me.ColLeft_Description.Visible = False
                Else
                    If ShowCodeNo = "1" Then Me.ColLeft_CodeNo.Visible = True Else Me.ColLeft_CodeNo.Visible = False
                    If ShowDesc = "1" Then Me.ColLeft_Description.Visible = True Else Me.ColLeft_Description.Visible = False
                End If

                Me.ColLeft_ProcessCode.Visible = False
                Me.ColLeft_QtyProduce.Visible = False
                Me.ColLeft_Unit.Visible = False
                Me.ColLeft_TotalWH.Visible = False
                Me.ColLeft_WorkUnit.Visible = False

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
                    End If
                Next
            End If
            For Each row As DataGridViewRow In Me.tgvSch.Rows
                If row.Visible = True Then
                    If Not String.IsNullOrEmpty(row.Cells(Me.ColLeft_LineCode.Index).Value) Then
                        Me.ColLeft_LineCode.Visible = True
                    End If
                    If Me.EnumGroupBy <> enumGroup.BY_WORKUNIT Then
                        If Not String.IsNullOrEmpty(row.Cells(Me.ColLeft_WorkUnit.Index).Value) Then
                            Me.ColLeft_WorkUnit.Visible = True
                        End If
                    Else
                        Me.ColLeft_WorkUnit.Visible = False
                    End If
                    If Not String.IsNullOrEmpty(row.Cells(Me.ColLeft_ProductionRate.Index).Value) Then
                        If ShowRate = "1" Then Me.ColLeft_ProductionRate.Visible = True
                    End If
                    If Not String.IsNullOrEmpty(row.Cells(Me.ColLeft_LocationInput.Index).Value) Then
                        If ShowInputLocation = "1" Then Me.ColLeft_LocationInput.Visible = True
                    End If
                    If Not String.IsNullOrEmpty(row.Cells(Me.ColLeft_LocationOutput.Index).Value) Then
                        If ShowOutputLocation = "1" Then Me.ColLeft_LocationOutput.Visible = True
                    End If
                    If row.Cells(Me.ColLeft_ActualGood.Index).Value <> "" Then
                        If 0 < CDec(row.Cells(Me.ColLeft_ActualGood.Index).Value) Then
                            If ShowActualGoodQty = "1" Then Me.ColLeft_ActualGood.Visible = True
                        End If
                    End If
                    If row.Cells(Me.ColLeft_ActualReject.Index).Value <> "" Then
                        If 0 < CDec(row.Cells(Me.ColLeft_ActualReject.Index).Value) Then
                            If ShowActualRejectQty = "1" Then Me.ColLeft_ActualReject.Visible = True
                        End If
                    End If
                    If row.Cells(Me.ColLeft_Actualproduce.Index).Value <> "" Then
                        If 0 < CDec(row.Cells(Me.ColLeft_Actualproduce.Index).Value) Then
                            If ShowActualProduceQty = "1" Then Me.ColLeft_Actualproduce.Visible = True
                        End If
                    End If
                    If Not String.IsNullOrEmpty(row.Cells(Me.ColLeft_Progress.Index).Value) Then
                        If ShowProgress_ = "1" Then Me.ColLeft_Progress.Visible = True
                    End If 
                End If
            Next
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
            Me.setColumnVisibility_1()
            Me.ExpandAll = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_ProductionSchedule_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            Me.objMonitor.stopProductionMonitor()
            Me.objMonitor.stopUserControlMonitor()
            Me.objMonitor.stopNotifMonitor()
            Me.objMonitor = Nothing

            objSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_StartDateDiff, _
            (Me.dtpFrom.Value.Date - Me.dtpFirmPlanned.Value.Date).TotalDays.ToString, Me.UserCode)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_ProductionSchedule_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try
            Me.btnFindMO.Visible = False
            Me.btnFindMO.TabStop = False
            Me.chkShowProgress.TabStop = False
            Me.chkShowProgress.Visible = False

            Me.llbExpandAll.Visible = False
            If Me.ShowGanttChart = False Then
                Me.tgvSch.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Top
                Me.SplitContainer1.Panel2Collapsed = True
            End If

            Me.ColLeft_LocationInput.ReadOnly = False
            Me.ColLeft_LocationOutput.ReadOnly = False
            Me.dtpTo.Enabled = False
            Me.llbColumnSettingSch.Text = "Column Settings && Data Filter"
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.btnRefresh.Size = New Size(70, Me.dtpTo.Height)
            Me.btnRefresh.Location = New Point(Me.dtpTo.Location.X + Me.dtpTo.Width + 5, Me.dtpTo.Location.Y)
            Me.btnRefresh.Visible = True

            Me.dgvBar.RowTemplate.Height = Me.tgvSch.RowTemplate.Height
            Me.dgvHeader1.ColumnHeadersHeight = Me.tgvSch.Location.Y - Me.dgvHeader1.Location.Y
            Me.dgvBar.Location = New Point(Me.dgvBar.Location.X, Me.tgvSch.Location.Y + Me.tgvSch.ColumnHeadersHeight)
            Me.dgvBar.Height = Me.tgvSch.Height - Me.tgvSch.ColumnHeadersHeight
            Me.dgvAddListener(Me.dgvBar)
            Me.cboTimeframe.SelectedIndex = 2
            Me.cboGroupBy.SelectedIndex = 0
            Me.setColumnVisibility_1()

            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.ToolStripStatusLabel1.Text = ""
            Me.LoadDataAwal1()
            Me.SetLlbFilterText()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AddSch()
        Try
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Today Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If Me.ObjMOItem.ApakahAdaYg_Ready_BelumTerjadwal = True Then
                Dim count As Integer = Me.GetSchDBRowsCount
                Dim ij As New FrmTrans_ProductionSchedule_NewSch
                ij.UserCode = Me.UserCode
                ij.BackColorAll = Me.BackColorAll
                ij.ShowDialog()
                ij = Nothing
                If count <> Me.GetSchDBRowsCount Then
                    Me.FullReload_ScheduleOnly()
                    Me.refreshBar = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Function GetSchDBRowsCount() As Integer
        Dim d As New DseTrans_ScheduleTableAdapters.tbltrans_mo_item_wo_item_schTableAdapter
        Dim v As New Nullable(Of Long)
        Try
            d.Connection.ConnectionString = Core_Module.GetConstring
            v = d.ScalarQuery_GetRowsCount
            If v.HasValue Then
                Return v.Value
            End If
            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        End Try
    End Function

    'Private Sub LoadDataAwal()
    '    Dim Dummy As String = ""
    '    Dim EndDate As Date = Today
    '    Dim StartTime As Date = Now
    '    Try
    '        Me.dtpFirmPlanned.Value = Me.ObjSch.GetFirmPlan
    '        Dummy = Me.ObjSetting1.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
    '           Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_SchHor_StartDate)
    '        If Dummy = "" Then
    '            Dummy = "-3"
    '        End If
    '        Me.dtpFrom.Value = Me.dtpFirmPlanned.Value + TimeSpan.FromDays(Dummy)
    '        While Me.dtpFrom.Value.Date.DayOfWeek <> DayOfWeek.Monday
    '            Me.dtpFrom.Value = (Me.dtpFrom.Value - TimeSpan.FromDays(1)).Date
    '        End While
    '        Me.dtpBarFrom.Value = Me.dtpFrom.Value

    '        'Dummy = Me.ObjSetting1.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, _
    '        '    Core_Module.ClsTransaksiSetting.EnumNamaParameter.ProductionSetting_SchHor_EndDate)
    '        'If Dummy = "" Then
    '        '    Dummy = "30"
    '        'End If
    '        'sementara untuk menyamakan dengan rosadin
    '        Dummy = Me.ScheduleHorizon ' tadinya 12 * 7
    '        EndDate = Me.dtpFrom.Value + TimeSpan.FromDays(Dummy)
    '        Me.dtpTo.Value = EndDate

    '        Me.FullReload()
    '        Me.refreshBar = False

    '        'MsgBox("Total duration: " + Me.objCD.DigitGrouping((Now - StartTime).TotalMinutes, 1))
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Sub ResizeMainPanel()
        Try
            If Me.ShowGanttChart = True Then
                Me.SplitContainer1.SplitterDistance = Me.GetVisibleTGVWidth
                Me.dgvBar.Width = Me.SplitContainer1.Panel2.Width
                Me.dgvHeader1.Width = Me.SplitContainer1.Panel2.Width
                Me.dgvHeader0.Width = Me.SplitContainer1.Panel2.Width
                Me.tgvSch.Width = Me.SplitContainer1.SplitterDistance
                Me.llbCollapsAll.Left = Me.SplitContainer1.Panel1.Width - 5 - Me.llbCollapsAll.Width
                Me.llbExpandAll.Left = Me.llbCollapsAll.Left - 5 - Me.llbExpandAll.Width
            Else
                Me.tgvSch.Width = Me.SplitContainer1.Panel1.Width - Me.tgvSch.Left - 5
            End If

            Me.Refresh()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
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

    Private Sub FrmTrans_ProductionSchedule_ResizeEnd(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.ResizeEnd
        Try
            Me.ResizeMainPanel()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_ProductionSchedule_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        Try
            Me.ResizeMainPanel()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_ProductionSchedule_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.SizeChanged
        Try
            Me.ResizeMainPanel()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnSet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefresh.Click
        Try
            If Not Me.BackgroundWorker3.IsBusy Then
                Me.ForceRefresh = True
                Me.Cursor = Cursors.WaitCursor
                Me.BackgroundWorker3.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkShowProgress_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowProgress.CheckedChanged
        Try
            Me.ShowProgress(Me.cboTimeframe.SelectedIndex, Me.dtpBarFrom.Value)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvBar_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvBar.SelectionChanged
        Try
            If Me.dgvBar.SelectedCells.Count > 0 Then
                Me.dgvBar.DefaultCellStyle.SelectionForeColor = Me.dgvBar.SelectedCells(0).Style.BackColor
                Me.dgvBar.DefaultCellStyle.SelectionBackColor = Me.dgvBar.SelectedCells(0).Style.BackColor
            End If
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
            Me.setColumnVisibility_1()
            'Me.SetSomeColumnVisibilty()
            Me.ResizeMainPanel()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ToolStripStatusLabel1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripStatusLabel1.Click
        Dim ij As New Core_Module.FrmTransaksiMonitor_Production
        Try
            ij.userIdentifier_1 = Me.userIdentifier
            ij.UserType = Me.userType
            ij.userCode = Me.UserCode
            ij.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
        End Try
    End Sub

    Private Function IsFullControl() As Boolean
        Try
            If Me.ToolStripStatusLabel1.Text.ToLower.Contains("full control") = True Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub cms1_Opening(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles cms1.Opening
        Dim IdentifierWO As String = ""
        Dim PrimaryWorkUnit As Double = -1
        Dim objUser As New Core_Module.ClsUser
        Dim Level As Integer = 0
        Dim IsFullControl As Boolean = False

        Dim objMousePosition As Point
        Dim objHitTestInfo As DataGridView.HitTestInfo
        Dim objBundling As New ClsTrans_StockBundling
        Dim IdentifierJO As String

        Try
            If Me.tgvSch.Nodes.Count <= 0 Then
                Exit Sub
            End If
            objMousePosition = Me.tgvSch.PointToClient(Control.MousePosition)
            objHitTestInfo = Me.tgvSch.HitTest(objMousePosition.X, objMousePosition.Y)
            Me.tgvSch.Item(objHitTestInfo.ColumnIndex, objHitTestInfo.RowIndex).Selected = True
            IdentifierJO = Me.tgvSch.Item(Me.ColLeft_IdentifierJO.Index, Me.tgvSch.SelectedCells(0).RowIndex).Value
            If Me.tgvSch.SelectedCells.Count = 0 Then
                e.Cancel = True
                Exit Sub
            End If
            If Me.ToolStripStatusLabel1.Text.ToLower.Contains("full control") = True Then
                IsFullControl = True
            Else
                IsFullControl = False
            End If
            IdentifierWO = Me.tgvSch.Item(Me.ColLeft_IdentifierWO.Index, Me.tgvSch.SelectedCells(0).RowIndex).Value
            PrimaryWorkUnit = Me.tgvSch.Item(Me.ColLeft_PrimaryWorkUnit.Index, Me.tgvSch.SelectedCells(0).RowIndex).Value
            Level = Me.tgvSch.Item(Me.ColLeft_NodeLevel.Index, Me.tgvSch.SelectedCells(0).RowIndex).Value

            ''none
            Me.CreateLogToolStripMenuItem.Enabled = True
            Me.EditLogToolStripMenuItem.Enabled = True
            Me.DeleteLogToolStripMenuItem.Enabled = True
            Me.ShowAllMOScheduleToolStripMenuItem.Enabled = True

            Me.ViewLogToolStripMenuItem.Enabled = True
            Me.CloseJOToolStripMenuItem.Enabled = True
            If objUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_ManufacturingOrder_SetFinishJO") = False Then
                Me.CloseJOToolStripMenuItem.Enabled = False
            End If

            ''hak akses
            If objUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiStockBundling2") = False Then
                Me.ViewLogToolStripMenuItem.Enabled = False
                Me.CreateLogToolStripMenuItem.Enabled = False
                Me.EditLogToolStripMenuItem.Enabled = False
                Me.DeleteLogToolStripMenuItem.Enabled = False
            End If

            ''data validation 
            If IdentifierWO = "" Or PrimaryWorkUnit = -1 Then
                'pas di item product
                Me.ViewLogToolStripMenuItem.Enabled = False
                Me.CreateLogToolStripMenuItem.Enabled = False
                Me.EditLogToolStripMenuItem.Enabled = False
                Me.DeleteLogToolStripMenuItem.Enabled = False
                Me.ShowAllMOScheduleToolStripMenuItem.Enabled = False
            End If

            If ObjMO.IsJOFinished(IdentifierJO) = True Then
                Me.CreateLogToolStripMenuItem.Enabled = False
                Me.EditLogToolStripMenuItem.Enabled = False
                Me.DeleteLogToolStripMenuItem.Enabled = False
            End If

            If IsFullControl = True Then
                Me.tms_ChangeColor.Enabled = True
                If Me.cboGroupBy.SelectedIndex = 0 Then
                    If Level = 0 Then
                        Me.DeleteJobOrderToolStripMenuItem.Enabled = True
                        Me.OptimizeScheduleToolStripMenuItem.Enabled = True
                    Else
                        Me.DeleteJobOrderToolStripMenuItem.Enabled = False
                        Me.OptimizeScheduleToolStripMenuItem.Enabled = False
                    End If
                Else
                    Me.DeleteJobOrderToolStripMenuItem.Enabled = False
                    Me.OptimizeScheduleToolStripMenuItem.Enabled = False
                End If
                If Me.cboGroupBy.SelectedIndex = 1 Then
                    'Me.ReorderScheduleToolStripMenuItem.Enabled = True
                    Me.SplitMergeScheduleToolStripMenuItem.Enabled = True
                Else
                    'Me.ReorderScheduleToolStripMenuItem.Enabled = False
                    Me.SplitMergeScheduleToolStripMenuItem.Enabled = False
                End If

                If Me.cboGroupBy.SelectedIndex = enumGroup.BY_WORKUNIT Or Level > 0 Then
                    Me.CloseJOToolStripMenuItem.Enabled = False
                End If
                If Me.ObjMO.IsJOFinished(IdentifierJO) = True Then
                    Me.CloseJOToolStripMenuItem.Enabled = False
                End If
            Else
                Me.tms_ChangeColor.Enabled = False
                Me.DeleteJobOrderToolStripMenuItem.Enabled = False
                Me.ReorderScheduleToolStripMenuItem.Enabled = False
                Me.OptimizeScheduleToolStripMenuItem.Enabled = False
                Me.SplitMergeScheduleToolStripMenuItem.Enabled = False
                Me.CloseJOToolStripMenuItem.Enabled = False
            End If

            Me.OptimizeScheduleToolStripMenuItem.Visible = False
            If Me.ShowGanttChart = False Then
                Me.tms_ChangeColor.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CreateLogToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CreateLogToolStripMenuItem.Click
        Dim IdentifierWO As String = ""
        Dim IdentifierJO As String = ""
        Dim PrimaryWorkUnit As Double = -1
        Dim i As New Inventory_Module.FrmTrans_StockBundling

        Try
            IdentifierWO = Me.tgvSch.Item(Me.ColLeft_IdentifierWO.Index, Me.tgvSch.SelectedCells(0).RowIndex).Value
            IdentifierJO = Me.tgvSch.Item(Me.ColLeft_IdentifierJO.Index, Me.tgvSch.SelectedCells(0).RowIndex).Value
            PrimaryWorkUnit = Me.tgvSch.Item(Me.ColLeft_PrimaryWorkUnit.Index, Me.tgvSch.SelectedCells(0).RowIndex).Value

            If IdentifierWO <> "" And PrimaryWorkUnit <> -1 Then
                i.UserCode = Me.UserCode
                i.BackColorAll = Me.BackColorAll
                i._IdentifierWO = IdentifierWO
                i._IdentifierJO = IdentifierJO
                i._PrimaryWorkUnit = PrimaryWorkUnit
                i.ObjStockBundling.TipeBundling = ClsTrans_StockBundling.EnumTipeBundling.Produksi1
                i._Command = FrmTrans_StockBundling.EnumCommand.NEW_
                If Me.ObjGlobalSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.ShowDialogOfProductionLog) = "0" Then
                    Dim ij As Inventory_Module.FrmTrans_StockBundling
                    For Each j As Windows.Forms.Form In Me.ParentForm.MdiChildren
                        ij = TryCast(j, Inventory_Module.FrmTrans_StockBundling)
                        If IsNothing(ij) = False Then
                            ij.BringToFront()
                            ij.Focus()
                            Exit Sub
                        End If
                    Next
                    i.MdiParent = Me.ParentForm
                    i.ShowIcon = False
                    i.Show()
                Else
                    i.ShowDialog()
                End If
                Me.UpdateProgress(Me.tgvSch.SelectedCells(0).RowIndex)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
        End Try
    End Sub

    Sub UpdateProgress(ByVal vRowIndex As Integer)
        Dim NewProgress As Decimal = 0
        Dim OldProgress As Decimal = 0

        Dim OldActualQty As Decimal = 0
        Dim rowData_ As DseTrans_Schedule.tbl_sch_barRow
        Try
            ''catat progress dan work hour sebelum perubahan, get juga nilai yang baru dari database
            OldProgress = Me.tgvSch.Item(Me.ColLeft_Progress.Index, vRowIndex).Value
            NewProgress = Me.ObjSch.GetProgress(Me.tgvSch.Item(Me.ColLeft_IdentifierJO.Index, vRowIndex).Value, _
            Me.tgvSch.Item(Me.ColLeft_IdentifierWO.Index, vRowIndex).Value, Me.tgvSch.Item(Me.ColLeft_PrimaryWorkUnit.Index, vRowIndex).Value)

            ''update tgv atas baris sch ybs
            Me.tgvSch.Item(Me.ColLeft_Progress.Index, vRowIndex).Value = Me.objCD.DigitGrouping(NewProgress, 2)

            ''update data table schBar dan tblData
            Me.tblSchBar.DefaultView.RowFilter = "[identifier jo] = '" & Me.tgvSch.Item(Me.ColLeft_IdentifierJO.Index, vRowIndex).Value & "' and [identifier wo] = '" & Me.tgvSch.Item(Me.ColLeft_IdentifierWO.Index, vRowIndex).Value & "' and [primary work unit] = " & Me.tgvSch.Item(Me.ColLeft_PrimaryWorkUnit.Index, vRowIndex).Value
            If Me.tblSchBar.DefaultView.Count > 0 Then
                Me.rowSch = Me.tblSchBar.DefaultView(0).Row
                Me.rowSch.progress = NewProgress
            End If
            Me.tblData.DefaultView.RowFilter = "[identifier jo] = '" & Me.tgvSch.Item(Me.ColLeft_IdentifierJO.Index, vRowIndex).Value & "' and [identifier wo] = '" & Me.tgvSch.Item(Me.ColLeft_IdentifierWO.Index, vRowIndex).Value & "' and [primary work unit] = " & Me.tgvSch.Item(Me.ColLeft_PrimaryWorkUnit.Index, vRowIndex).Value
            If Me.tblData.DefaultView.Count > 0 Then
                Me.rowData = Me.tblData.DefaultView(0).Row
                Me.rowData.progress = NewProgress
            End If

            If Me.cboGroupBy.SelectedIndex = 0 Then
                ''hitung progress dan wh untuk group jo-process
                Me.tblProgress.DefaultView.RowFilter = "[identifier JO] = '" & Me.tgvSch.Item(Me.ColLeft_IdentifierJO.Index, vRowIndex).Value & "' AND [primary process] = " & Me.tgvSch.Item(Me.ColLeft_PrimaryProcess.Index, vRowIndex).Value.ToString
                If Me.tblProgress.DefaultView.Count > 0 Then
                    Me.rowProgress = Me.tblProgress.DefaultView(0).Row
                    OldActualQty = Me.rowProgress.actual_qty
                    Me.rowProgress.actual_qty = (Me.rowProgress.actual_qty - Me.rowData.qty_produce * OldProgress / 100) + (Me.rowData.qty_produce * NewProgress / 100)
                    Me.rowProgress.progress = Me.rowProgress.actual_qty / Me.rowProgress.target_qty * 100
                End If
                ''naik ke atas dari sebelum baris tgv ybs sampai ketemu headernya
                For i As Integer = vRowIndex - 1 To 0 Step -1
                    If Me.tgvSch.Item(Me.ColLeft_NodeLevel.Index, i).Value = "1" Then
                        Me.tgvSch.Item(Me.ColLeft_Progress.Index, i).Value = Me.objCD.DigitGrouping(Me.rowProgress.progress, 2)
                        Me.tblData.DefaultView.RowFilter = "[is header] = 1 AND [tree level] = 1 AND [node identifier lv 0] = '" & Me.tgvSch.Item(Me.ColLeft_NodeLv0Identifier.Index, i).Value & "' AND [node identifier lv 1] = '" & Me.tgvSch.Item(Me.ColLeft_NodeLv1Identifier.Index, i).Value & "'"
                        If Me.tblData.DefaultView.Count > 0 Then
                            rowData_ = Me.tblData.DefaultView(0).Row
                            rowData_.progress = Me.rowProgress.progress
                        End If
                        Exit For
                    End If
                Next
                ''hitung progress dan wh untuk group jo
                Me.tblProgress2.DefaultView.RowFilter = "[identifier JO] = '" & Me.rowProgress.identifier_JO & "'"
                If Me.tblProgress2.DefaultView.Count > 0 Then
                    Me.rowProgress2 = Me.tblProgress2.DefaultView(0).Row
                    Me.rowProgress2.actual_qty = Me.rowProgress2.actual_qty - OldActualQty + Me.rowProgress.actual_qty
                    Me.rowProgress2.progress = Me.rowProgress2.actual_qty / Me.rowProgress2.target_qty * 100
                End If
                ''naik ke atas dari sebelum baris tgv ybs sampai ketemu super headernya
                For i As Integer = vRowIndex - 1 To 0 Step -1
                    If Me.tgvSch.Item(Me.ColLeft_NodeLevel.Index, i).Value = "0" Then
                        Me.tgvSch.Item(Me.ColLeft_Progress.Index, i).Value = Me.objCD.DigitGrouping(Me.rowProgress2.progress, 2)
                        Me.tblData.DefaultView.RowFilter = "[is header] = 1 AND [tree level] = 0 AND [node identifier lv 0] = '" & Me.tgvSch.Item(Me.ColLeft_NodeLv0Identifier.Index, i).Value & "'"
                        If Me.tblData.DefaultView.Count > 0 Then
                            rowData_ = Me.tblData.DefaultView(0).Row
                            rowData_.progress = Me.rowProgress2.progress
                        End If
                        Exit For
                    End If
                Next
            End If

            ''betulkan progress bar pada dgvBar
            Me.UpdateProgressBar()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
        End Try
    End Sub

    Sub UpdateProgressBar()
        Dim dr As DataGridViewRow
        Dim barValueFrom, barValueTo As Date

        Try
            If Me.ShowGanttChart = True Then
                dr = Me.dgvBar.Rows(Me.tblData.Rows.IndexOf(Me.rowData))
                barValueFrom = Me.dtpBarFrom.Value
                barValueTo = Me.dtpBarFrom.Value + intervalPerBar_ByWeek
                Dim workBarCount As Integer = 0
                Dim firstBarIndex As Integer = -1
                Dim barProgressCount As Integer = 0
                With Me.rowData
                    For i As Integer = 1 To Me.dgvBar.Columns.Count
                        If ((barValueFrom <= .plan_start_time And barValueTo > .plan_start_time) Or _
                                (barValueFrom >= .plan_start_time And barValueTo <= .plan_end_time) Or _
                                (barValueFrom < .plan_end_time And barValueTo >= .plan_end_time)) Then
                            If firstBarIndex = -1 Then
                                firstBarIndex = i
                            End If
                            workBarCount += 1
                        End If
                        barValueFrom += intervalPerBar_ByWeek
                        barValueTo += intervalPerBar_ByWeek
                    Next
                    barProgressCount = Me.objCD.DigitGrouping(Me.rowData.progress / 100 * workBarCount, 0, Core_Module.ClsCekData.EnumTipeRounding.Down)
                    For i As Integer = firstBarIndex To Me.dgvBar.Columns.Count
                        If i = Me.LineTodayIndex Then
                            dr.Cells("ColBar" & i).Value = Me.line_today
                        ElseIf i = Me.LineFirmPlannedIndex Then
                            dr.Cells("ColBar" & i).Value = Me.line_firmplanned
                        Else
                            dr.Cells("ColBar" & i).Value = Nothing
                        End If
                        If Me.chkShowProgress.Checked And i <= firstBarIndex - 1 + barProgressCount Then
                            Dim Draw As Boolean = False
                            If Me.tblData.Rows.IndexOf(Me.rowData) = Me.dgvBar.Rows.Count - 1 Then
                                Draw = True
                            ElseIf Me.rowData.is_header = 0 Or (Me.rowData.is_header = 1 And Me.dgvBar.Rows(Me.tblData.Rows.IndexOf(Me.rowData) + 1).Visible = False) Then
                                Draw = True
                            End If
                            If Draw Then
                                If i = Me.LineTodayIndex Then
                                    dr.Cells("ColBar" & i).Value = gantt_progress_today
                                ElseIf i = Me.LineFirmPlannedIndex Then
                                    dr.Cells("ColBar" & i).Value = gantt_progress_firmplanned
                                Else
                                    dr.Cells("ColBar" & i).Value = gantt_progress
                                End If
                            End If
                        End If
                    Next
                End With
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EditLogToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditLogToolStripMenuItem.Click
        Dim IdentifierWO As String = ""
        Dim IdentifierJO As String = ""
        Dim PrimaryWorkUnit As Double = -1
        Dim i As New Inventory_Module.FrmTrans_StockBundling
        Try
            IdentifierWO = Me.tgvSch.Item(Me.ColLeft_IdentifierWO.Index, Me.tgvSch.SelectedCells(0).RowIndex).Value
            IdentifierJO = Me.tgvSch.Item(Me.ColLeft_IdentifierJO.Index, Me.tgvSch.SelectedCells(0).RowIndex).Value
            PrimaryWorkUnit = Me.tgvSch.Item(Me.ColLeft_PrimaryWorkUnit.Index, Me.tgvSch.SelectedCells(0).RowIndex).Value
            If IdentifierWO <> "" And PrimaryWorkUnit <> -1 Then
                Dim d As New DseTrans_ScheduleTableAdapters.DataTable5TableAdapter
                Dim t As New DseTrans_Schedule.DataTable5DataTable
                Dim r As DseTrans_Schedule.DataTable5Row
                d.Connection.ConnectionString = Core_Module.GetConstring
                d.Fill(t, ClsTrans_StockBundling.EnumTipeBundling.Produksi1, IdentifierWO, PrimaryWorkUnit, IdentifierJO)

                If t.Count > 1 Then
                    Dim f As New FrmTrans_StockBundling_ChooseOne(t)
                    f.ShowDialog()
                    i._PrimaryBundling = f.PrimaryBundling
                ElseIf t.Count = 1 Then
                    r = t.Rows(0)
                    i._PrimaryBundling = r.Primary
                ElseIf t.Count = 0 Then
                    MsgBox("This WO number don't have a log.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
                If i._PrimaryBundling <> -1 Then
                    i.UserCode = Me.UserCode
                    i.BackColorAll = Me.BackColorAll
                    i._Command = FrmTrans_StockBundling.EnumCommand.EDIT
                    If Me.ObjGlobalSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.ShowDialogOfProductionLog) = "0" Then
                        Dim ij As Inventory_Module.FrmTrans_StockBundling
                        For Each j As Windows.Forms.Form In Me.ParentForm.MdiChildren
                            ij = TryCast(j, Inventory_Module.FrmTrans_StockBundling)
                            If IsNothing(ij) = False Then
                                ij.BringToFront()
                                ij.Focus()
                                Exit Sub
                            End If
                        Next
                        i.MdiParent = Me.ParentForm
                        i.ShowIcon = False
                        i.Show()
                    Else
                        i.ShowDialog()
                    End If
                    Me.UpdateProgress(Me.tgvSch.SelectedCells(0).RowIndex)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
        End Try
    End Sub

    Private Sub DeleteLogToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteLogToolStripMenuItem.Click
        Dim IdentifierWO As String = ""
        Dim IdentifierJO As String = ""
        Dim PrimaryWorkUnit As Double = -1
        Dim i As New Inventory_Module.FrmTrans_StockBundling
        Try
            IdentifierWO = Me.tgvSch.Item(Me.ColLeft_IdentifierWO.Index, Me.tgvSch.SelectedCells(0).RowIndex).Value
            IdentifierJO = Me.tgvSch.Item(Me.ColLeft_IdentifierJO.Index, Me.tgvSch.SelectedCells(0).RowIndex).Value
            PrimaryWorkUnit = Me.tgvSch.Item(Me.ColLeft_PrimaryWorkUnit.Index, Me.tgvSch.SelectedCells(0).RowIndex).Value
            Dim d As New DseTrans_ScheduleTableAdapters.DataTable5TableAdapter
            Dim t As New DseTrans_Schedule.DataTable5DataTable
            Dim r As DseTrans_Schedule.DataTable5Row
            d.Connection.ConnectionString = Core_Module.GetConstring
            d.Fill(t, ClsTrans_StockBundling.EnumTipeBundling.Produksi1, IdentifierWO, PrimaryWorkUnit, IdentifierJO)
            If t.Count > 1 Then
                Dim f As New FrmTrans_StockBundling_ChooseOne(t)
                f.ShowDialog()
                i._PrimaryBundling = f.PrimaryBundling
            ElseIf t.Count = 1 Then
                r = t.Rows(0)
                i._PrimaryBundling = r.Primary
            ElseIf t.Count = 0 Then
                MsgBox("This WO number don't have a log.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            If i._PrimaryBundling <> -1 Then
                i.UserCode = Me.UserCode
                i.BackColorAll = Me.BackColorAll
                i._Command = FrmTrans_StockBundling.EnumCommand.DELETE
                If Me.ObjGlobalSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ProductionSetting, Core_Module.ClsTransaksiSetting.EnumNamaParameter.ShowDialogOfProductionLog) = "0" Then
                    Dim ij As Inventory_Module.FrmTrans_StockBundling
                    For Each j As Windows.Forms.Form In Me.ParentForm.MdiChildren
                        ij = TryCast(j, Inventory_Module.FrmTrans_StockBundling)
                        If IsNothing(ij) = False Then
                            ij.BringToFront()
                            ij.Focus()
                            Exit Sub
                        End If
                    Next
                    i.MdiParent = Me.ParentForm
                    i.ShowIcon = False
                    i.Show()
                Else
                    i.ShowDialog()
                End If
                Me.UpdateProgress(Me.tgvSch.SelectedCells(0).RowIndex)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DeleteJobOrderToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DeleteJobOrderToolStripMenuItem.Click
        Dim IdentiferJO As String = ""
        Try
            IdentiferJO = Me.tgvSch.Item(Me.ColLeft_IdentifierJO.Index, Me.tgvSch.SelectedCells(0).RowIndex).Value
            If Me.ObjSch.ApakahBolehHapusJadwalJO(IdentiferJO, Me.ObjMO, Me.ObjLog) = True Then
                If MsgBox("Delete schedule for job order: " + _
                            Me.ObjMO.GetJONumber(Me.tgvSch.Item(Me.ColLeft_IdentifierJO.Index, Me.tgvSch.SelectedCells(0).RowIndex).Value), _
                            MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    'cek apakah boleh hapus jadwal jo
                    'hapus jadwal jo
                    Me.ObjSch.HapusJadwal(IdentiferJO)

                    'Dim ij As New FrmTrans_ProductionSchedule_Progress3
                    'Me.ObjSch.HapusJadwal(IdentiferJO, ij)
                    'ij = Nothing

                    'tampilkan ulang dgv
                    Windows.Forms.Cursor.Current = Cursors.WaitCursor
                    Me.FullReload_ScheduleOnly()
                    'Me.LoadTblSch(Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date)
                    'Me.reloadGanttChart()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker6_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker6.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker6_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker6.RunWorkerCompleted
        Try
            'load sch bar
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.FullReload_ScheduleOnly()
            'Me.LoadTblSch(Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date)
            'Me.reloadGanttChart()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ReorderScheduleToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ReorderScheduleToolStripMenuItem.Click
        Dim ij As New FrmTrans_ProductionSchedule_Reorder
        Dim rowSchBar0 As DseTrans_Schedule.tbl_sch_barRow
        'Dim Node As AdvancedDataGridView.TreeGridNode

        Try
            If Me.tgvSch.SelectedCells.Count > 0 Then
                ij.UserCode = Me.UserCode
                ij.TblSchBar = Me.tblSchBar.Copy()
                'ij.TblSchBarTanpaWO = Me.tblSchBar_TanpaItemWO
                ij.ObjSch = Me.ObjSch
                ij.PrimaryWU = Me.tgvSch.Item(Me.ColLeft_PrimaryWorkUnit.Index, Me.tgvSch.SelectedCells(0).RowIndex).Value
                'Node = Me.tgvSch.Nodes(Me.tgvSch.SelectedCells(0).RowIndex)
                ij.IdentifierJO = Me.tgvSch.Item(Me.ColLeft_IdentifierJO.Index, Me.tgvSch.SelectedCells(0).RowIndex).Value
                ij.IdentifierWO = Me.tgvSch.Item(Me.ColLeft_IdentifierWO.Index, Me.tgvSch.SelectedCells(0).RowIndex).Value

                If ij.ShowDialog = Windows.Forms.DialogResult.Yes Then
                    Me.Cursor = Cursors.WaitCursor
                    ij.TblSchBar.DefaultView.RowFilter = "[primary work unit] = " & ij.PrimaryWU.ToString
                    For Each row As Data.DataRowView In ij.TblSchBar.DefaultView
                        rowSchBar0 = row.Row
                        Me.ObjSch.UpdatePlanStartEndSchedule(rowSchBar0.plan_start_time, rowSchBar0.plan_end_time, rowSchBar0.identifier_jo, rowSchBar0.identifier_wo, ij.PrimaryWU)
                        Me.tblSchBar.DefaultView.RowFilter = "[primary work unit] = " & ij.PrimaryWU.ToString & " and [identifier jo] = '" & rowSchBar0.identifier_jo & "' and [identifier wo] = '" & rowSchBar0.identifier_wo & "'"
                        If Me.tblSchBar.DefaultView.Count > 0 Then
                            Me.rowSch = Me.tblSchBar.DefaultView(0).Row
                            Me.rowSch.plan_start_time = rowSchBar0.plan_start_time
                            Me.rowSch.plan_end_time = rowSchBar0.plan_end_time
                        End If
                    Next
                    Me.reloadGanttChart()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Function getBarNodeIdentiferLevel(ByVal vLevel As Integer, ByVal vBarIndex As Integer) As Double
        Try
            Me.tblBarIndex.DefaultView.RowFilter = "[Bar Index] = " & vBarIndex.ToString
            If Me.tblBarIndex.DefaultView.Count > 0 Then
                Me.rowBarIndex = Me.tblBarIndex.DefaultView(0).Row
                If vLevel = 0 Then
                    Return Me.rowBarIndex("Node Identifier Lv 0")
                Else
                    Return Me.rowBarIndex("Node Identifier Lv 1")
                End If
            End If
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Function GetFilterLine(ByVal vColumnName As String) As String
        Dim v As String = ""
        Try
            If Me.listOfLine.Count > 0 Then
                For Each s As String In Me.listOfLine
                    If v = "" Then
                        v = "[" & vColumnName & "] = " & s
                    Else
                        v = v & " OR [" & vColumnName & "] = " & s
                    End If
                Next
            Else
                v = ""
            End If
            If v = "" Then
                Return v
            Else
                Return "(" & v & ")"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Sub CloseJOToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CloseJOToolStripMenuItem.Click
        Dim IdentifierJO As String = ""
        Dim objMO As New ClsTrans_ManufacturingOrder
        Dim objCloseJO As New FrmTrans_ManufacturingOrder_SetFinishJO
        Dim frmDate As New FrmTrans_ManufacturingOrder_SetFinishJO_Date
        Try
            IdentifierJO = Me.tgvSch.Item(Me.ColLeft_IdentifierJO.Index, Me.tgvSch.SelectedCells(0).RowIndex).Value
            If objMO.ApakahBolehCloseJO(IdentifierJO, True) Then
                If frmDate.ShowDialog <> Windows.Forms.DialogResult.OK Then
                    Exit Sub
                End If
                If objCloseJO.UpdateStatus(IdentifierJO, True, frmDate.DateTimePicker1.Value.Date) = False Then
                    MsgBox("Close JO failed.", MsgBoxStyle.Information, "Information")
                Else
                    MsgBox("JO closed.", MsgBoxStyle.Information, "Information")
                    Me.tgvSch.Item(Me.ColLeft_Tree.Index, Me.tgvSch.SelectedCells(0).RowIndex).Value += " (Closed)"
                    Me.ResizeMainPanel()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ViewLogToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ViewLogToolStripMenuItem.Click
        Dim IdentifierWO As String = ""
        Dim IdentifierJO As String = ""
        Dim PrimaryWorkUnit As Double = -1
        Dim i As New Inventory_Module.FrmTrans_StockBundling
        Try
            IdentifierWO = Me.tgvSch.Item(Me.ColLeft_IdentifierWO.Index, Me.tgvSch.SelectedCells(0).RowIndex).Value
            IdentifierJO = Me.tgvSch.Item(Me.ColLeft_IdentifierJO.Index, Me.tgvSch.SelectedCells(0).RowIndex).Value
            PrimaryWorkUnit = Me.tgvSch.Item(Me.ColLeft_PrimaryWorkUnit.Index, Me.tgvSch.SelectedCells(0).RowIndex).Value
            If IdentifierWO <> "" And PrimaryWorkUnit <> -1 Then
                Dim d As New DseTrans_ScheduleTableAdapters.DataTable5TableAdapter
                Dim t As New DseTrans_Schedule.DataTable5DataTable
                Dim r As DseTrans_Schedule.DataTable5Row
                d.Connection.ConnectionString = Core_Module.GetConstring
                d.Fill(t, ClsTrans_StockBundling.EnumTipeBundling.Produksi1, IdentifierWO, PrimaryWorkUnit, IdentifierJO)

                If t.Count > 1 Then
                    Dim f As New FrmTrans_StockBundling_ChooseOne(t)
                    f.ShowDialog()
                    i._PrimaryBundling = f.PrimaryBundling
                ElseIf t.Count = 1 Then
                    r = t.Rows(0)
                    i._PrimaryBundling = r.Primary
                ElseIf t.Count = 0 Then
                    MsgBox("This WO number don't have a log.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
                If i._PrimaryBundling <> -1 Then
                    Dim ij As Inventory_Module.FrmTrans_StockBundling
                    Dim j As Windows.Forms.Form
                    For Each j In Me.ParentForm.MdiChildren
                        ij = TryCast(j, Inventory_Module.FrmTrans_StockBundling)
                        If IsNothing(ij) = False Then
                            ''jika sudah terbuka sebelumnya
                            ij.UserCode = Me.UserCode
                            ij._PrimaryBundling = i._PrimaryBundling
                            ij._Command = FrmTrans_StockBundling.EnumCommand.VIEW
                            ij.BringToFront()
                            ij.Focus()
                            Exit Sub
                        End If
                    Next
                    ''jika tidak sedang terbuka
                    i.UserCode = Me.UserCode
                    i.BackColorAll = Core_Module.GetSMESSBackColorAll
                    i._Command = FrmTrans_StockBundling.EnumCommand.VIEW
                    i.MdiParent = Me.ParentForm
                    i.ShowIcon = False
                    i.Show()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
        End Try
    End Sub

    Private Sub SplitMergeScheduleToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SplitMergeScheduleToolStripMenuItem.Click
        Dim frm As New FrmTrans_ProductionSchedule_SplitMerge
        Try
            frm.UserCode = Me.UserCode
            frm.BackColorAll = Me.BackColorAll
            'frm.TblSchBar = Me.tblSchBar.Copy()
            'frm.TblSchBarTanpaWO = Me.tblSchBar_TanpaItemWO
            frm.ObjSch = Me.ObjSch
            frm.PrimaryWU = Me.tgvSch.Item(Me.ColLeft_PrimaryWorkUnit.Index, Me.tgvSch.SelectedCells(0).RowIndex).Value
            If frm.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.tblSchBar.Rows.Clear()
                Me.tblSchBar = frm.TblSchBar.Copy
                Me.reloadGanttChart()
                MsgBox("Split schedule completed.", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FindScheduleToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FindScheduleToolStripMenuItem.Click
        Try
            Dim frmFind As New FrmTrans_ProductionSchedule_Find
            frmFind.TBLSchBar = Me.tblData.Copy
            frmFind.GroupBy = Me.cboGroupBy.SelectedIndex
            If frmFind.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.Find(frmFind.PrimaryMO, frmFind.IdentifierJO, frmFind.IdentifierWO, frmFind.PrimaryProcess, frmFind.PrimaryWorkunit)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ShowAllMOScheduleToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ShowAllMOScheduleToolStripMenuItem.Click
        Dim IdentiferJO As String = ""
        Dim Frm As New FrmTrans_ProductionSchedule_DetailWO

        Try
            IdentiferJO = Me.tgvSch.Item(Me.ColLeft_IdentifierJO.Index, Me.tgvSch.SelectedCells(0).RowIndex).Value
            Frm.IdentifierJO = IdentiferJO
            Frm.UserCode = Me.UserCode
            Frm.tblSchBar = Me.tblSchBar
            Frm.tblSchBar_TanpaItemWO = Me.tblSchBar_TanpaItemWO
            Frm.ShowDialog()
            If Frm.AnyChange Then
                Me.btnRefresh.PerformClick()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnAddRemoveHoliday_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddRemoveHoliday.Click
        Dim ij As New FrmTrans_ProductionSchedule_Holiday
        Try
            ij.UserCode = Me.UserCode
            ij.BackColorAll = Me.BackColorAll
            ij.ShowDialog()
            If ij.IsEdited Then
                Me.refreshBar = False
                Me.FullReload()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    '  Public Sub LoadDgvItemWO_TreeView(ByVal vIdentifierJO As String, ByVal vDataTable_ As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable, _
    'ByVal vTgvWO As AdvancedDataGridView.TreeGridView, ByVal vFrm As FrmTrans_ManufacturingOrder)
    '      Dim DataTable_ As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable
    '      Dim DataRow_ As DseTrans_ManufacturingOrder.tbltrans_mo_item_woRow
    '      Dim ProsesCode As String = ""
    '      Dim ProsesDesc As String = ""
    '      Dim WorkUnitCode As String = ""
    '      Dim Status_ As String = ""
    '      Dim Progress_ As Decimal = 0
    '      Dim ListPrimary As New List(Of Double)

    '      Try
    '          vTgvWO.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
    '          vTgvWO.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
    '          vTgvWO.AllowUserToResizeColumns = False
    '          vTgvWO.AllowUserToResizeRows = False

    '          If vTgvWO.Nodes.Count > 0 Then
    '              vTgvWO.Nodes.Clear()
    '              'vTgvWO.Rows.Clear()
    '          End If

    '          DataTable_ = vDataTable_.Copy
    '          vDataTable_.DefaultView.RowFilter = "[Identifier JO] = '" & vIdentifierJO & "' AND [Identifier WO Parent] = '-'"
    '          If vDataTable_.DefaultView.Count > 0 Then
    '              DataRow_ = vDataTable_.DefaultView(0).Row
    '              ProsesCode = objProses.GetCode(DataRow_.Primary_Process)
    '              ProsesDesc = objProses.GetDesc(DataRow_.Primary_Process)
    '              WorkUnitCode = Me.GetWorkUnitCode(DataRow_.Identifier) 'objWU.GetCode(DataRow_.Primary_Work_Unit)
    '              If DataRow_.Status = ClsTrans_ManufacturingOrder.EnumPilihStatusWO.Free Then
    '                  Status_ = "Free"
    '              ElseIf DataRow_.Status = ClsTrans_ManufacturingOrder.EnumPilihStatusWO.Planned Then
    '                  Status_ = "Planned"
    '              ElseIf DataRow_.Status = ClsTrans_ManufacturingOrder.EnumPilihStatusWO.Scheduled Then
    '                  Status_ = "Scheduled"
    '              ElseIf DataRow_.Status = ClsTrans_ManufacturingOrder.EnumPilihStatusWO.On_Progress Then
    '                  Status_ = "On Progress"
    '              ElseIf DataRow_.Status = ClsTrans_ManufacturingOrder.EnumPilihStatusWO.Finish Then
    '                  Status_ = "Finish"
    '              End If
    '              Progress_ = Me.ObjCD.DigitGrouping((DataRow_.Qty_Actual_Prod / DataRow_.Qty_Target_Prod) * 100, 2)
    '              If Progress_ = 0 Then
    '                  Progress_ = 0
    '              End If

    '              vTgvWO.Nodes.Add("", Status_.ToString, Progress_.ToString, DataRow_.WO_Number.ToString, DataRow_.Product_Code, DataRow_.Description, ProsesCode.ToString, ProsesDesc.ToString, WorkUnitCode.ToString, DataRow_.Remarks, "", DataRow_.Primary_Process, DataRow_.Identifier, DataRow_.Identifier_JO)
    '              FillNodes(DataTable_, DataRow_, vTgvWO.Nodes(0), True, vTgvWO, vFrm)
    '              'fungsi untuk expand treegridview
    '              vTgvWO.Nodes(0).Expand()
    '              Me.Update_TGV_GetExplodeItem(ListPrimary, vTgvWO.Nodes(0))
    '          End If

    '          vTgvWO.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
    '      Catch ex As Exception
    '          MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '      End Try
    '  End Sub
End Class
