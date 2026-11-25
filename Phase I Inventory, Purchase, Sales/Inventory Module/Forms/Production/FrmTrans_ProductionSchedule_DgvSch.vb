Option Explicit On
Option Strict Off

Partial Public Class FrmTrans_ProductionSchedule
    'Private Sub SetTampilanDGVSch()
    '    Try
    '        If Me.vGroupBy = EnumGrouping.ByWorkUnit Then
    '            Me.ColSch_JONumber1.Visible = False
    '            Me.ColSch_JONumber2.Visible = True
    '        Else
    '            Me.ColSch_JONumber1.Visible = True
    '            Me.ColSch_JONumber2.Visible = False
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    'Private Sub LoadDGVSch()
    '    Try
    '        'kosongkan dgv
    '        Windows.Forms.Cursor.Current = Cursors.WaitCursor
    '        Me.dgvSch.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
    '        Me.dgvSch.Rows.Clear()

    '        'load data table sch dari db
    '        Me.ObjSch.LoadScheduleBar(Me.dtpFrom.Value, Me.dtpTo.Value, Me.vTblSch)
    '        If Me.vTblSch.Rows.Count > 0 Then
    '            'isikan no transaksi dll
    '            Me.IsikanNoTransaksiSch()
    '            'tampilkan ke dalam dgv
    '            If Me.vGroupBy = EnumGrouping.ByWorkUnit Then
    '                Me.LoadDGVSch_GroupByWorkUnit()
    '            Else
    '                Me.LoadDGVSch_GroupByJobOrder()
    '            End If
    '        End If

    '        Me.dgvSch.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        Windows.Forms.Cursor.Current = Cursors.Arrow
    '    End Try
    'End Sub

    'Private Sub IsikanNoTransaksiSch()
    '    Dim vRowSch As DseTrans_Schedule.tbl_sch_barRow
    '    Dim Idx, Idx_Ref, Idx_Ref1, Idx_Ref2 As Integer
    '    'Dim DTA_WS As New DseTrans_ScheduleTableAdapters.tblmaster_workshiftTableAdapter
    '    'Dim TBL_WS As New DseTrans_Schedule.tblmaster_workshiftDataTable
    '    'Dim row_WS As DseTrans_Schedule.tblmaster_workshiftRow
    '    Dim DTA_Process As New DseTrans_ScheduleTableAdapters.tblmaster_prosesTableAdapter
    '    Dim TBL_Process As New DseTrans_Schedule.tblmaster_prosesDataTable
    '    Dim row_Process As DseTrans_Schedule.tblmaster_prosesRow
    '    Dim DTA_WU As New DseTrans_ScheduleTableAdapters.tblmaster_workunitTableAdapter
    '    Dim TBL_WU As New DseTrans_Schedule.tblmaster_workunitDataTable
    '    Dim row_WU As DseTrans_Schedule.tblmaster_workunitRow
    '    Dim DTA_MO As New DseTrans_ScheduleTableAdapters.tbltrans_moTableAdapter
    '    Dim TBL_MO As New DseTrans_Schedule.tbltrans_moDataTable
    '    Dim row_MO As DseTrans_Schedule.tbltrans_moRow
    '    Dim DTA_JO As New DseTrans_ScheduleTableAdapters.tbltrans_mo_itemTableAdapter
    '    Dim TBL_JO As New DseTrans_Schedule.tbltrans_mo_itemDataTable
    '    Dim row_JO As DseTrans_Schedule.tbltrans_mo_itemRow
    '    Dim DTA_WO As New DseTrans_ScheduleTableAdapters.tbltrans_mo_item_woTableAdapter
    '    Dim TBL_WO As New DseTrans_Schedule.tbltrans_mo_item_woDataTable
    '    Dim row_WO As DseTrans_Schedule.tbltrans_mo_item_woRow
    '    Dim PrimaryMin, PrimaryMax As Double

    '    Try
    '        'DTA_WS.Connection.ConnectionString = Core_Module.GetConstring
    '        DTA_Process.Connection.ConnectionString = Core_Module.GetConstring
    '        DTA_WU.Connection.ConnectionString = DTA_Process.Connection.ConnectionString
    '        DTA_MO.Connection.ConnectionString = DTA_Process.Connection.ConnectionString
    '        DTA_JO.Connection.ConnectionString = DTA_Process.Connection.ConnectionString
    '        DTA_WO.Connection.ConnectionString = DTA_Process.Connection.ConnectionString

    '        ''mengisikan workshift code, 
    '        'Me.vTblSch.DefaultView.RowFilter = ""
    '        'Me.vTblSch.DefaultView.Sort = "[primary workshift]"
    '        'DTA_WS.Fill(TBL_WS)
    '        'Idx = 0
    '        'Idx_Ref = 0
    '        'While Idx <= vTblSch.DefaultView.Count - 1 And Idx_Ref <= TBL_WS.Rows.Count - 1
    '        '    vRowSch = vTblSch.DefaultView(Idx).Row
    '        '    row_WS = TBL_WS.Rows(Idx_Ref)
    '        '    If vRowSch.primary_workshift = -1 Then
    '        '        vRowSch.workshift_code = ""
    '        '        Idx += 1
    '        '    Else
    '        '        If row_WS.primary = vRowSch.primary_workshift Then
    '        '            vRowSch.workshift_code = row_WS.code
    '        '            Idx += 1
    '        '        ElseIf vRowSch.primary_workshift > row_WS.primary Then
    '        '            Idx_Ref += 1
    '        '        End If
    '        '    End If
    '        'End While

    '        'mengisikan process code
    '        Me.vTblSch.DefaultView.Sort = "[primary process]"
    '        DTA_Process.Fill(TBL_Process)
    '        Idx = 0
    '        Idx_Ref = 0
    '        While Idx <= vTblSch.DefaultView.Count - 1 And Idx_Ref <= TBL_Process.Rows.Count - 1
    '            vRowSch = vTblSch.DefaultView(Idx).Row
    '            If vRowSch.primary_process = -1 Then
    '                vRowSch.process_code = ""
    '                Idx += 1
    '            Else
    '                row_Process = TBL_Process.Rows(Idx_Ref)
    '                If row_Process.Primary = vRowSch.primary_process Then
    '                    vRowSch.process_code = row_Process.Code
    '                    Idx += 1
    '                ElseIf vRowSch.primary_process > row_Process.Primary Then
    '                    Idx_Ref += 1
    '                End If
    '            End If
    '        End While

    '        'mengisikan work unit code
    '        Me.vTblSch.DefaultView.Sort = "[primary work unit]"
    '        DTA_WU.Fill(TBL_WU)
    '        Idx = 0
    '        Idx_Ref = 0
    '        While Idx <= vTblSch.DefaultView.Count - 1 And Idx_Ref <= TBL_WU.Rows.Count - 1
    '            vRowSch = vTblSch.DefaultView(Idx).Row
    '            If vRowSch.primary_work_unit = -1 Then
    '                vRowSch.work_unit_code = ""
    '                Idx += 1
    '            Else
    '                row_WU = TBL_WU.Rows(Idx_Ref)
    '                If row_WU.primary = vRowSch.primary_work_unit Then
    '                    vRowSch.work_unit_code = row_WU.work_unit_code
    '                    Idx += 1
    '                ElseIf vRowSch.primary_work_unit > row_WU.primary Then
    '                    Idx_Ref += 1
    '                End If
    '            End If
    '        End While

    '        'mengisikan mo no, jo no, wo no
    '        Me.vTblSch.DefaultView.Sort = "[primary mo], [identifier jo], [identifier wo]"
    '        vRowSch = vTblSch.DefaultView(0).Row
    '        PrimaryMin = vRowSch.primary_mo
    '        vRowSch = vTblSch.DefaultView(vTblSch.DefaultView.Count - 1).Row
    '        PrimaryMax = vRowSch.primary_mo
    '        DTA_MO.Fill(TBL_MO, PrimaryMin, PrimaryMax)
    '        DTA_JO.Fill(TBL_JO, PrimaryMin, PrimaryMax)
    '        DTA_WO.Fill(TBL_WO, PrimaryMin, PrimaryMax)
    '        Idx = 0
    '        Idx_Ref = 0
    '        Idx_Ref1 = 0
    '        Idx_Ref2 = 0
    '        While Idx <= vTblSch.DefaultView.Count - 1 And Idx_Ref <= TBL_WU.Rows.Count - 1 And _
    '        Idx_Ref1 <= vTblSch.DefaultView.Count - 1 And Idx_Ref2 <= TBL_WU.Rows.Count - 1
    '            vRowSch = vTblSch.DefaultView(Idx).Row
    '            If vRowSch.primary_mo = -1 Then
    '                vRowSch.mo_no = ""
    '                vRowSch.jo_no = ""
    '                vRowSch.wo_no = ""
    '                Idx += 1
    '            Else
    '                row_MO = TBL_MO.Rows(Idx_Ref)
    '                row_JO = TBL_JO.DefaultView(Idx_Ref1).Row
    '                row_WO = TBL_WO.Rows(Idx_Ref2)

    '                If vRowSch.primary_mo <> row_MO.Primary Then
    '                    Idx_Ref += 1
    '                Else
    '                    If vRowSch.identifier_jo <> row_JO.Identifier_JO Then
    '                        Idx_Ref1 += 1
    '                    Else
    '                        If vRowSch.identifier_wo <> row_WO.Identifier_WO Then
    '                            Idx_Ref2 += 1
    '                        Else
    '                            'primary mo, identifier jo, identifier wo sama semua, copy data no transaksi
    '                            vRowSch.mo_no = row_MO.MO_Number
    '                            vRowSch.jo_no = row_JO.JO_Number
    '                            vRowSch.wo_no = row_WO.WO_Number
    '                        End If
    '                    End If
    '                End If
    '            End If
    '        End While
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        'DTA_WS.Dispose()
    '        'TBL_WS.Dispose()
    '        DTA_Process.Dispose()
    '        TBL_Process.Dispose()
    '        DTA_WU.Dispose()
    '        TBL_WU.Dispose()
    '        DTA_MO.Dispose()
    '        TBL_MO.Dispose()
    '        DTA_JO.Dispose()
    '        TBL_JO.Dispose()
    '        DTA_WO.Dispose()
    '        TBL_WO.Dispose()
    '    End Try
    'End Sub

    'Private Sub LoadDGVSch_GroupByWorkUnit()
    '    Dim vRowSch As DseTrans_Schedule.tbl_sch_barRow
    '    Dim rowDGV As DataGridViewRow
    '    Dim DV As Data.DataRowView
    '    Dim DummyWUCode As String = ""
    '    Dim HideEndDate As Boolean = True
    '    Dim ShowFreeTime As String = ""
    '    Dim ShowOffTime As String = ""
    '    Dim AddRow As Boolean = True

    '    Try
    '        ShowFreeTime = Me.ObjSettingUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.ProductionSetting, _
    '        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.ProductionSetting_ShowFreeTime, Me.UserCode)
    '        If ShowFreeTime = "" Then
    '            If Me.ObjSch.GetDefault_ShowFreeTime = True Then
    '                ShowFreeTime = "1"
    '            Else
    '                ShowFreeTime = "0"
    '            End If
    '            Me.ObjSettingUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.ProductionSetting, _
    '            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.ProductionSetting_ShowFreeTime, ShowFreeTime, Me.UserCode)
    '        End If

    '        ShowOffTime = Me.ObjSettingUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.ProductionSetting, _
    '        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.ProductionSetting_ShowOffTime, Me.UserCode)
    '        If ShowOffTime = "" Then
    '            If Me.ObjSch.GetDefault_ShowOffTime = True Then
    '                ShowOffTime = "1"
    '            Else
    '                ShowOffTime = "0"
    '            End If
    '            Me.ObjSettingUser.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.ProductionSetting, _
    '            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.ProductionSetting_ShowOffTime, ShowOffTime, Me.UserCode)
    '        End If

    '        Me.vTblSch.DefaultView.RowFilter = ""
    '        Me.vTblSch.DefaultView.Sort = "[work unit code], [plan start time]"
    '        For Each DV In Me.vTblSch.DefaultView
    '            vRowSch = DV.Row
    '            AddRow = True

    '            If vRowSch.tipe = CInt(ClsTrans_Schedule.EnumTipeSchBar.free_time) Then
    '                If ShowFreeTime <> "1" Then
    '                    AddRow = False
    '                End If
    '            ElseIf vRowSch.tipe = CInt(ClsTrans_Schedule.EnumTipeSchBar.off_time) Then
    '                If ShowOffTime <> "1" Then
    '                    AddRow = False
    '                End If
    '            End If

    '            If AddRow = True Then
    '                Me.dgvSch.Rows.Add()
    '                rowDGV = Me.dgvSch.Rows(Me.dgvSch.Rows.Count - 1)
    '                With vRowSch
    '                    rowDGV.Cells(Me.ColSch_a.Index).Value = .a
    '                    rowDGV.Cells(Me.ColSch_b.Index).Value = .b
    '                    rowDGV.Cells(Me.ColSch_EndDate.Index).Value = .plan_end_time
    '                    rowDGV.Cells(Me.ColSch_EndTime.Index).Value = .plan_end_time
    '                    rowDGV.Cells(Me.ColSch_g.Index).Value = .g
    '                    rowDGV.Cells(Me.ColSch_IdentifierJO.Index).Value = .identifier_jo
    '                    rowDGV.Cells(Me.ColSch_IdentifierWO.Index).Value = .identifier_wo
    '                    rowDGV.Cells(Me.ColSch_JONumber1.Index).Value = ""
    '                    rowDGV.Cells(Me.ColSch_JONumber2.Index).Value = .jo_no
    '                    rowDGV.Cells(Me.ColSch_MONumber.Index).Value = .mo_no
    '                    rowDGV.Cells(Me.ColSch_PrimaryMO.Index).Value = .primary_mo
    '                    rowDGV.Cells(Me.ColSch_PrimaryProcess.Index).Value = .primary_process
    '                    'rowDGV.Cells(Me.ColSch_PrimaryWorkShift.Index).Value = .primary_workshift
    '                    rowDGV.Cells(Me.ColSch_PrimaryWorkUnit.Index).Value = .primary_work_unit
    '                    rowDGV.Cells(Me.ColSch_Process.Index).Value = .process_code
    '                    rowDGV.Cells(Me.ColSch_Progress.Index).Value = .progress
    '                    rowDGV.Cells(Me.ColSch_r.Index).Value = .r
    '                    rowDGV.Cells(Me.ColSch_StartDate.Index).Value = .plan_start_time
    '                    rowDGV.Cells(Me.ColSch_StartTime.Index).Value = .plan_start_time
    '                    rowDGV.Cells(Me.ColSch_WorkOrderNo.Index).Value = .wo_no
    '                    rowDGV.Cells(Me.ColSch_WorkUnit.Index).Value = .work_unit_code
    '                    rowDGV.Cells(Me.ColSch_Type.Index).Value = Me.ObjSch.GetTipeSchBar_String(.tipe)
    '                    'rowDGV.DefaultCellStyle.BackColor = Color.FromArgb(.a, .r, .g, .b)
    '                    If .plan_start_time.Date <> .plan_end_time.Date Then
    '                        HideEndDate = False
    '                    End If
    '                End With
    '            End If
    '        Next

    '        Me.ColSch_EndDate.DefaultCellStyle.Format = "dd MMM yyyy"
    '        Me.ColSch_EndTime.DefaultCellStyle.Format = "HH.mm"
    '        Me.ColSch_StartDate.DefaultCellStyle.Format = "dd MMM yyyy"
    '        Me.ColSch_StartTime.DefaultCellStyle.Format = "HH.mm"
    '        If HideEndDate = True Then
    '            Me.ColSch_EndDate.Visible = False
    '        Else
    '            Me.ColSch_EndDate.Visible = True
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    'Private Sub LoadDGVSch_GroupByJobOrder()
    '    Try

    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub
End Class
