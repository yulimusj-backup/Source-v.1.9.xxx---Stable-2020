Option Explicit On
Option Strict Off

Public Class FrmTrans_ProductionSchedule_NewSch2_Detail
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public TBLSchBar As New DseTrans_Schedule.tbl_sch_barDataTable
    Public PrimaryWU As Double = -1
    Public IdentifierWO As String = ""
    Public StartDate, EndDate As Date
    Public UserCode As String = ""

    Dim ObjSetting As New Core_Module.ClsTransaksiSettingUser
    Dim rowSchBar, rowSchBar1 As DseTrans_Schedule.tbl_sch_barRow
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjWU As New Core_Module.ClsMasterWorkUnit
    Dim BufferA As String = ""
    Dim BufferB As String = ""
    Dim Buffer As String = ""

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub FrmTrans_ProductionSchedule_NewSch2_Detail_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.Text = Me.Text + ": " + Me.ObjWU.GetCode(Me.PrimaryWU)
            Me.DgvWUAllocate_Process.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvDetailWH.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            Me.LoadDGVWorkQueue()
            Me.LoadDGVDetailWH()
            Me.DgvWUAllocate_Process.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadDGVWorkQueue()
        Dim DTA As New DseTrans_ScheduleTableAdapters.DataTable2TableAdapter
        Dim TBL As New DseTrans_Schedule.DataTable2DataTable
        Dim row As DseTrans_Schedule.DataTable2Row
        Dim PrimaryMOMax, PrimaryMOMin As Double
        Dim mn, jk As Double
        Dim rowDGV As DataGridViewRow
        'Dim BoldFont As New Font(Me.DgvWUAllocate_Process.DefaultCellStyle.Font, FontStyle.Bold)

        Try
            Me.DgvWUAllocate_Process.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DgvWUAllocate_Process.Rows.Clear()

            Me.TBLSchBar.DefaultView.RowFilter = "[primary work unit] = " + Me.PrimaryWU.ToString + " AND tipe = " + CInt(ClsTrans_Schedule.EnumTipeSchBar.work_time).ToString
            Me.TBLSchBar.DefaultView.Sort = "[plan start time]"
            If Me.TBLSchBar.DefaultView.Count > 0 Then
                PrimaryMOMax = Double.MinValue
                PrimaryMOMin = Double.MaxValue
                DTA.Connection.ConnectionString = Core_Module.GetConstring

                For Each DV As Data.DataRowView In Me.TBLSchBar.DefaultView
                    rowSchBar = DV.Row
                    If rowSchBar.primary_mo < PrimaryMOMin Then
                        PrimaryMOMin = rowSchBar.primary_mo
                    End If
                    If rowSchBar.primary_mo > PrimaryMOMax Then
                        PrimaryMOMax = rowSchBar.primary_mo
                    End If
                Next
                DTA.FillBy_PrimaryMaxMin(TBL, PrimaryMOMax, PrimaryMOMin)

                mn = 0
                For Each DV As Data.DataRowView In Me.TBLSchBar.DefaultView
                    rowSchBar = DV.Row
                    If (rowSchBar.plan_end_time.Date <= Me.dtpTo.Value.Date And Me.rowSchBar.plan_end_time.Date >= Me.dtpFrom.Value.Date) Or _
                    (rowSchBar.plan_start_time.Date <= Me.dtpTo.Value.Date And Me.rowSchBar.plan_start_time.Date >= Me.dtpFrom.Value.Date) Then
                        If rowSchBar.wo_no = "" Then
                            TBL.DefaultView.RowFilter = "[Identifier WO] = '" + rowSchBar.identifier_wo + "'"
                            row = TBL.DefaultView(0).Row
                            jk = 0
                            While mn + jk <= Me.TBLSchBar.DefaultView.Count - 1
                                rowSchBar1 = Me.TBLSchBar.DefaultView(mn + jk).Row
                                If rowSchBar.plan_start_time.Date > Me.dtpTo.Value Then
                                    Exit While
                                Else
                                    If rowSchBar1.identifier_wo = rowSchBar.identifier_wo Then
                                        rowSchBar1.mo_no = row.MO_Number
                                        rowSchBar1.wo_no = row.WO_Number
                                        rowSchBar1.jo_no = row.JO_Number
                                        rowSchBar1.process_code = row.Process_Code
                                        rowSchBar1.product_code = row.Product_Code
                                        rowSchBar1.product_desc = row.Product_Desc
                                        rowSchBar1.decimal_point = row._Decimal
                                        rowSchBar1.unit_code = row.Unit_Code
                                    End If
                                End If

                                jk += 1
                            End While
                        End If

                        Me.DgvWUAllocate_Process.Rows.Add()
                        rowDGV = Me.DgvWUAllocate_Process.Rows(Me.DgvWUAllocate_Process.Rows.Count - 1)
                        rowDGV.Cells(Me.ColCodeNo.Index).Value = rowSchBar.product_code
                        rowDGV.Cells(Me.ColDecimal.Index).Value = rowSchBar.decimal_point
                        rowDGV.Cells(Me.ColDesc.Index).Value = rowSchBar.product_desc
                        rowDGV.Cells(Me.ColEndDate.Index).Value = Format(rowSchBar.plan_end_time, "dd MMM yyyy")
                        rowDGV.Cells(Me.ColEndTime.Index).Value = Format(rowSchBar.plan_end_time, "HH.mm")
                        rowDGV.Cells(Me.ColIdentifierWO.Index).Value = rowSchBar.identifier_jo
                        rowDGV.Cells(Me.ColJONo.Index).Value = rowSchBar.jo_no
                        rowDGV.Cells(Me.ColPrimaryProcess.Index).Value = rowSchBar.primary_process
                        rowDGV.Cells(Me.ColProcessCode.Index).Value = rowSchBar.process_code
                        rowDGV.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(rowSchBar.qty_produce, rowSchBar.decimal_point)
                        rowDGV.Cells(Me.ColRate.Index).Value = Me.ObjCD.DigitGrouping(rowSchBar.rate, rowSchBar.decimal_point)
                        rowDGV.Cells(Me.ColStartDate.Index).Value = Format(rowSchBar.plan_start_time, "dd MMM yyyy")
                        rowDGV.Cells(Me.ColStartTime.Index).Value = Format(rowSchBar.plan_start_time, "HH.mm")
                        rowDGV.Cells(Me.ColWONumber.Index).Value = rowSchBar.wo_no
                        rowDGV.Cells(Me.ColTotalWorkHour.Index).Value = Me.ObjCD.Convert_ToString(rowSchBar.total_work_hour)
                        rowDGV.Cells(Me.ColUnitCode.Index).Value = rowSchBar.unit_code

                        If rowSchBar.identifier_wo = Me.IdentifierWO Then
                            'rowDGV.DefaultCellStyle.Font = BoldFont
                            rowDGV.DefaultCellStyle.ForeColor = Color.Red
                        End If
                    ElseIf rowSchBar.plan_start_time.Date > Me.dtpTo.Value Then
                        Exit For
                    End If
                Next
            End If

            Me.DgvWUAllocate_Process.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Private Sub LoadDGVDetailWH()
        Dim TBLAcuan As New DseTrans_Schedule.DataTable3DataTable
        Dim rowAcuan As DseTrans_Schedule.DataTable3Row
        Dim DTAWS_Allocate As New DseTrans_OT_ItemTableAdapters.tbltrans_mo_workshift_allocateTableAdapter
        Dim TBLWS_Allocate As New DseTrans_OT_Item.tbltrans_mo_workshift_allocateDataTable
        Dim rowWS_Allocate, rowWS_Allocate1 As DseTrans_OT_Item.tbltrans_mo_workshift_allocateRow
        Dim DTAWS_Item As New DseTrans_OT_ItemTableAdapters.tblmaster_workshift_itemTableAdapter
        Dim TBLWS_Item As New DseTrans_OT_Item.tblmaster_workshift_itemDataTable
        Dim rowWS_Item As DseTrans_OT_Item.tblmaster_workshift_itemRow
        Dim DTAWS_Holiday As New DseTrans_OT_ItemTableAdapters.tblmaster_workshift_itemholidayTableAdapter
        Dim TBLWS_Holiday As New DseTrans_OT_Item.tblmaster_workshift_itemholidayDataTable
        Dim mn, jk As Double
        Dim DTAOT As New DseTrans_OT_ItemTableAdapters.tbltrans_overtime_itemTableAdapter
        Dim TBLOT As New DseTrans_OT_Item.tbltrans_overtime_itemDataTable
        Dim rowOT As DseTrans_OT_Item.tbltrans_overtime_itemRow
        Dim AdaTglBeda As Boolean = False

        Try
            DTAWS_Allocate.Connection.ConnectionString = Core_Module.GetConstring
            DTAWS_Item.Connection.ConnectionString = DTAWS_Allocate.Connection.ConnectionString
            DTAWS_Holiday.Connection.ConnectionString = DTAWS_Allocate.Connection.ConnectionString
            DTAOT.Connection.ConnectionString = DTAWS_Allocate.Connection.ConnectionString
            TBLAcuan.Rows.Clear()
            Me.dgvDetailWH.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvDetailWH.Rows.Clear()
            If Me.StartDate <> Me.EndDate Then
                'load data allocated workshift, dari tanggal start -1 sampai dengan end +1, consider holiday juga, tambahkan ke tabel acuan
                DTAWS_Allocate.FillBy_Primary_WorkUnit(TBLWS_Allocate, Me.EndDate.Date + TimeSpan.FromDays(1), Me.StartDate.Date - TimeSpan.FromDays(1), Me.PrimaryWU)
                DTAWS_Item.FillBy2(TBLWS_Item)
                DTAWS_Holiday.Fill(TBLWS_Holiday, Me.StartDate.Date - TimeSpan.FromDays(1))
                mn = 0
                For Each rowWS_Allocate In TBLWS_Allocate.Rows
                    rowWS_Allocate.day_of_week = rowWS_Allocate._date.DayOfWeek
                Next
                For Each rowWS_Allocate In TBLWS_Allocate.Rows
                    If rowWS_Allocate.sch_in = -1 Then
                        TBLWS_Item.DefaultView.RowFilter = "[primary transaksi] = " + rowWS_Allocate.primary_workshift.ToString + _
                        " AND [day of week] = " + CInt(rowWS_Allocate.day_of_week).ToString
                        rowWS_Item = TBLWS_Item.DefaultView(0).Row
                        jk = 0
                        While mn + jk <= TBLWS_Allocate.Rows.Count - 1
                            rowWS_Allocate1 = TBLWS_Allocate.Rows(mn + jk)
                            If rowWS_Allocate1.primary_workshift = rowWS_Allocate.primary_workshift And rowWS_Allocate1.day_of_week = rowWS_Allocate.day_of_week Then
                                rowWS_Allocate1.sch_in = rowWS_Item.schedule_in
                                rowWS_Allocate1.sch_out = rowWS_Item.schedule_out
                                rowWS_Allocate1.break_out = rowWS_Item.schedule_break_out
                                rowWS_Allocate1.break_in = rowWS_Item.schedule_break_in
                                rowWS_Allocate1.tipe_workdays = rowWS_Item.tipe_workdays
                                rowWS_Allocate1.shift_code = rowWS_Item.workshift_code
                            End If
                            jk += 1
                        End While
                    End If
                    mn += 1
                Next
                For Each rowWS_Allocate In TBLWS_Allocate.Rows
                    TBLWS_Holiday.DefaultView.RowFilter = "[primary transaksi] = " + rowWS_Allocate.primary_workshift.ToString + _
                    " AND date = #" + Format(rowWS_Allocate._date, "MM/dd/yyyy") + "#"
                    If TBLWS_Holiday.DefaultView.Count > 0 Then
                        rowWS_Allocate.tipe_workdays = CInt(Core_Module.ClsMasterWorkShift.TipeWorkdays.Holiday)
                    End If
                    If rowWS_Allocate.tipe_workdays = CInt(Core_Module.ClsMasterWorkShift.TipeWorkdays.Weekday) Then
                        With rowAcuan
                            If rowWS_Allocate.break_in <> rowWS_Allocate.break_out Then
                                'before break
                                rowAcuan = TBLAcuan.NewDataTable3Row
                                With rowAcuan
                                    .Tipe = rowWS_Allocate.shift_code
                                    If rowWS_Allocate.sch_in < rowWS_Allocate.break_out Then
                                        .Start = rowWS_Allocate._date.Date + TimeSpan.FromMinutes(rowWS_Allocate.sch_in)
                                        ._End = rowWS_Allocate._date.Date + TimeSpan.FromMinutes(rowWS_Allocate.break_out)
                                    Else
                                        .Start = rowWS_Allocate._date.Date + TimeSpan.FromMinutes(rowWS_Allocate.sch_in)
                                        ._End = rowWS_Allocate._date.Date + TimeSpan.FromDays(1) + TimeSpan.FromMinutes(rowWS_Allocate.break_out)
                                    End If
                                End With
                                TBLAcuan.Rows.Add(rowAcuan)
                                'after break
                                rowAcuan = TBLAcuan.NewDataTable3Row
                                With rowAcuan
                                    .Tipe = rowWS_Allocate.shift_code
                                    If rowWS_Allocate.break_in < rowWS_Allocate.sch_out Then
                                        .Start = rowWS_Allocate._date.Date + TimeSpan.FromMinutes(rowWS_Allocate.break_in)
                                        ._End = rowWS_Allocate._date.Date + TimeSpan.FromMinutes(rowWS_Allocate.sch_out)
                                    Else
                                        .Start = rowWS_Allocate._date.Date + TimeSpan.FromMinutes(rowWS_Allocate.break_in)
                                        ._End = rowWS_Allocate._date.Date + TimeSpan.FromDays(1) + TimeSpan.FromMinutes(rowWS_Allocate.sch_out)
                                    End If
                                End With
                                TBLAcuan.Rows.Add(rowAcuan)
                            Else
                                rowAcuan = TBLAcuan.NewDataTable3Row
                                With rowAcuan
                                    .Tipe = rowWS_Allocate.shift_code
                                    If rowWS_Allocate.sch_in < rowWS_Allocate.sch_out Then
                                        .Start = rowWS_Allocate._date.Date + TimeSpan.FromMinutes(rowWS_Allocate.sch_in)
                                        ._End = rowWS_Allocate._date.Date + TimeSpan.FromMinutes(rowWS_Allocate.sch_out)
                                    Else
                                        .Start = rowWS_Allocate._date.Date + TimeSpan.FromMinutes(rowWS_Allocate.sch_in)
                                        ._End = rowWS_Allocate._date.Date + TimeSpan.FromDays(1) + TimeSpan.FromMinutes(rowWS_Allocate.sch_out)
                                    End If
                                End With
                                TBLAcuan.Rows.Add(rowAcuan)
                            End If
                        End With
                    End If
                Next

                'load data overtime, dari tanggal start -1 sampai dengan end +1, tambahkan ke tabel acuan
                DTAOT.FillBy1(TBLOT, Me.PrimaryWU, Me.StartDate.Date - TimeSpan.FromDays(1))
                For Each rowOT In TBLOT.Rows
                    rowAcuan = TBLAcuan.NewDataTable3Row
                    With rowAcuan
                        .Start = rowOT.start
                        ._End = rowOT._end
                        .Tipe = "Overtime"
                    End With
                    TBLAcuan.Rows.Add(rowAcuan)
                Next

                'mulai dari start date - end date, tampilkan secara berurutan
                TBLAcuan.DefaultView.RowFilter = ""
                TBLAcuan.DefaultView.Sort = "Start ASC"
                For Each DV As Data.DataRowView In TBLAcuan.DefaultView
                    rowAcuan = DV.Row
                    If Me.EndDate <= rowAcuan.Start Then
                        Exit For
                    ElseIf rowAcuan.Start >= Me.StartDate And rowAcuan._End <= Me.EndDate Then
                        'ambil secara keseluruhan
                        Me.AddRowDGV(rowAcuan.Tipe, rowAcuan.Start, rowAcuan._End)
                        If rowAcuan.Start.Date <> rowAcuan._End.Date Then
                            AdaTglBeda = True
                        End If
                    ElseIf rowAcuan.Start < Me.StartDate And rowAcuan._End <= Me.EndDate And rowAcuan._End > Me.StartDate Then
                        'ambil sebelah kanan dari row acuan saja
                        Me.AddRowDGV(rowAcuan.Tipe, Me.StartDate, rowAcuan._End)
                        If Me.StartDate.Date <> rowAcuan._End.Date Then
                            AdaTglBeda = True
                        End If
                    ElseIf rowAcuan._End > Me.EndDate And rowAcuan.Start > Me.StartDate And rowAcuan._End >= Me.EndDate Then
                        'ambil sebelah kiri row acuan
                        Me.AddRowDGV(rowAcuan.Tipe, rowAcuan.Start, Me.EndDate)
                        If rowAcuan.Start.Date <> Me.EndDate.Date Then
                            AdaTglBeda = True
                        End If
                    ElseIf Me.StartDate >= rowAcuan.Start And Me.EndDate <= rowAcuan._End Then
                        'ambil start date - end date
                        Me.AddRowDGV(rowAcuan.Tipe, Me.StartDate, Me.EndDate)
                        If Me.StartDate.Date <> Me.EndDate.Date Then
                            AdaTglBeda = True
                        End If
                    End If
                Next
            End If

            If AdaTglBeda = False Then
                Me.ColDetail_EndDate.Visible = False
                Me.ColDetail_StartDate.HeaderText = "Date"
                Me.ColDetail_StartTime.HeaderText = "Start"
                Me.ColDetail_EndTime.HeaderText = "End"
            Else
                Me.ColDetail_EndDate.Visible = True
                Me.ColDetail_StartDate.HeaderText = "Start"
                Me.ColDetail_StartTime.HeaderText = ""
                Me.ColDetail_EndTime.HeaderText = ""
            End If
            Me.dgvDetailWH.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TBLAcuan.Dispose()
            DTAWS_Allocate.Dispose()
            TBLWS_Allocate.Dispose()
            DTAWS_Item.Dispose()
            TBLWS_Item.Dispose()
            DTAWS_Holiday.Dispose()
            TBLWS_Holiday.Dispose()
            DTAOT.Dispose()
            TBLOT.Dispose()
        End Try
    End Sub

    Private Sub AddRowDGV(ByVal vTipe As String, ByVal vStartTime As Date, ByVal vEndTime As Date)
        Dim rowDGV As DataGridViewRow

        Try
            Me.dgvDetailWH.Rows.Add()
            rowDGV = Me.dgvDetailWH.Rows(Me.dgvDetailWH.Rows.Count - 1)
            rowDGV.Cells(Me.ColDetail_EndDate.Index).Value = Format(vEndTime, "dd MMM yyyy")
            rowDGV.Cells(Me.ColDetail_EndTime.Index).Value = Format(vEndTime, "HH.mm")
            rowDGV.Cells(Me.ColDetail_No.Index).Value = rowDGV.Index + 1
            rowDGV.Cells(Me.ColDetail_StartDate.Index).Value = Format(vStartTime, "dd MMM yyyy")
            rowDGV.Cells(Me.ColDetail_StartTime.Index).Value = Format(vStartTime, "HH.mm")
            rowDGV.Cells(Me.ColDetail_TipeWH.Index).Value = vTipe
            rowDGV.Cells(Me.ColDetail_WH.Index).Value = Me.ObjCD.Convert_ToString((vEndTime - vStartTime).TotalMinutes)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpFrom_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpFrom.Enter
        Try
            Me.BufferA = Me.dtpFrom.Value.Date.ToString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpFrom_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpFrom.Leave
        Try
            Me.bgw1.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpTo_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpTo.Enter
        Try
            Me.BufferB = Me.dtpTo.Value.Date.ToString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpTo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpTo.Leave
        Try
            Me.bgw2.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgw1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgw1.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgw1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgw1.RunWorkerCompleted
        Try
            If Me.BufferA <> Me.dtpFrom.Value.Date.ToString Then
                'cek tanggal dari sampai
                If Me.dtpFrom.Value.Date > Me.dtpTo.Value.Date Then
                    Me.dtpFrom.Value = CDate(Me.BufferA)
                    Exit Sub
                End If

                'load ulang data dgv work queue
                Me.LoadDGVWorkQueue()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgw2_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bgw2.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bgw2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgw2.RunWorkerCompleted
        Try
            If Me.BufferB <> Me.dtpTo.Value.Date.ToString Then
                'cek tanggal dari sampai
                If Me.dtpFrom.Value.Date > Me.dtpTo.Value.Date Then
                    Me.dtpTo.Value = CDate(Me.BufferB)
                    Exit Sub
                End If

                'load ulang data dgv work queue
                Me.LoadDGVWorkQueue()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub SetColumnVisibility()
        Dim ShowJO, ShowProcess, ShowCodeNo, ShowDesc As String
        Try
            ShowJO = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2_Detail, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_Detail_ShowJO, Me.UserCode)
            ShowProcess = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2_Detail, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_Detail_ShowProcess, Me.UserCode)
            ShowCodeNo = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2_Detail, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_Detail_ShowCodeNo, Me.UserCode)
            ShowDesc = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2_Detail, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_Detail_ShowDesc, Me.UserCode)
            If ShowJO = "" And ShowProcess = "" And ShowCodeNo = "" And ShowDesc = "" Then
                ShowJO = "1"
                ShowProcess = "1"
                ShowCodeNo = "1"
                ShowDesc = "1"
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2_Detail, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_Detail_ShowJO, ShowJO, Me.UserCode)
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2_Detail, _
                         Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_Detail_ShowProcess, ShowProcess, Me.UserCode)
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2_Detail, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_Detail_ShowCodeNo, ShowCodeNo, Me.UserCode)
                Me.ObjSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_NewSch2_Detail, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_NewSch2_Detail_ShowDesc, ShowDesc, Me.UserCode)
            End If
            If ShowJO = "1" Then
                Me.ColJONo.Visible = True
            Else
                Me.ColJONo.Visible = False
            End If
            If ShowProcess = "1" Then
                Me.ColProcessCode.Visible = True
            Else
                Me.ColProcessCode.Visible = False
            End If
            If ShowCodeNo = "1" Then
                Me.ColCodeNo.Visible = True
            Else
                Me.ColCodeNo.Visible = False
            End If
            If ShowDesc = "1" Then
                Me.ColDesc.Visible = True
            Else
                Me.ColDesc.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbColumnsSetting_LinkClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbColumnsSetting.LinkClicked
        Dim ij As New FrmTrans_ProductionSchedule_NewSch2_Detail_ColSettings
        Try
            ij.UserCode = Me.UserCode
            If ij.ShowDialog() = Windows.Forms.DialogResult.OK Then
                Me.SetColumnVisibility()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        Try
            If Me.TabControl1.SelectedTab.Name = Me.TabPage1.Name Then
                Me.llbColumnsSetting.Visible = True
            ElseIf Me.TabControl1.SelectedTab.Name = Me.TabPage2.Name Then
                Me.llbColumnsSetting.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class