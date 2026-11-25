Option Explicit On
Option Strict Off
Public Class FrmTrans_ProductionSchedule_DetailWO
    Dim TblItemWO As New DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable
    Dim ObjMOItem As New ClsTrans_ManufacturingOrder_Item
    Dim ObjMO As New ClsTrans_ManufacturingOrder
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim ObjProses As New Core_Module.ClsMasterProses
    Dim ObjCD As New Core_Module.ClsCekData
    Dim TBLWorkUnit As New DseTrans_ManufacturingOrder.DataTable9DataTable
    Dim DTAWorkUnit As New DseTrans_ManufacturingOrderTableAdapters.DataTable9TableAdapter
    Dim rowWorkUnit As DseTrans_ManufacturingOrder.DataTable9Row

    Public IdentifierJO As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public UserCode As String = ""
    Public TblSchBar As New DseTrans_Schedule.tbl_sch_barDataTable
    Public TblSchBar_TanpaItemWO As New DseTrans_Schedule.tbl_sch_barDataTable
    Public AnyChange As Boolean = False
    Dim ObjSch As New ClsTrans_Schedule
    Dim RowSch As DseTrans_Schedule.tbl_sch_barRow
    Dim ShowCodeNo, ShowDesc, ShowProcess, ShowProcessDesc As String
    Dim ObjSetting As New Core_Module.ClsTransaksiSettingUser

    Private Sub FrmTrans_ProductionSchedule_DetailWO_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.DTAWorkUnit.Connection.ConnectionString = Core_Module.GetConstring
            Me.dgvItemWO.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            'load data item wo
            Me.TblItemWO.Rows.Clear()
            Me.ObjMOItem.IsiDataItemWO(Me.IdentifierJO, Me.TblItemWO)
            Me.DTAWorkUnit.Fill(Me.TBLWorkUnit, Me.IdentifierJO)

            'tampilkan ke dalam dgv
            Me.LoadDgvItemWO_TableView(Me.IdentifierJO, Me.TblItemWO, Me.dgvItemWO)

            Me.TxtJONumber.Text = Me.ObjMO.GetJONumber(Me.IdentifierJO)
            Me.TxtMONumber.Text = Me.ObjMO.GetNoTransaksi(Me.IdentifierJO.Substring(0, Me.IdentifierJO.IndexOf("-")))
            Me.dgvItemWO.ReadOnly = True
            Me.dgvItemWO.AllowUserToAddRows = False
            Me.dgvItemWO.AllowUserToDeleteRows = False
            Me.SetColumnVisibility()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub LoadDgvItemWO_TableView(ByVal vIdentifierJO As String, ByVal vDataTable_ As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable, _
  ByVal vDGV As DataGridView)
        Dim DataTable_ As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable
        Dim RowTable_ As DseTrans_ManufacturingOrder.tbltrans_mo_item_woRow
        Dim rowDGV As DataGridViewRow
        Dim TBLDummy As New DseTrans_ManufacturingOrder.tbl_codenoDataTable
        Dim rowDummy As DseTrans_ManufacturingOrder.tbl_codenoRow
        Dim TBLDummy_Process As New DseTrans_ManufacturingOrder.tbl_processDataTable
        Dim rowDummy_Process As DseTrans_ManufacturingOrder.tbl_processRow
        'Dim Status_ As String = ""

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            vDGV.Rows.Clear()

            DataTable_ = vDataTable_.Copy
            vDataTable_.DefaultView.RowFilter = "[Identifier JO] = '" & vIdentifierJO & "'"
            vDataTable_.DefaultView.Sort = "Level DESC, Primary"

            If vDataTable_.DefaultView.Count > 0 Then
                For Each DV As Data.DataRowView In vDataTable_.DefaultView
                    RowTable_ = DV.Row
                    If RowTable_.Qty_Target_Prod > 0 Then
                        vDGV.Rows.Add()
                        rowDGV = vDGV.Rows(vDGV.Rows.Count - 1)
                        TBLDummy.DefaultView.RowFilter = "[Primary Product] = " + RowTable_.Primary_Product.ToString
                        If TBLDummy.DefaultView.Count = 0 Then
                            rowDummy = TBLDummy.Newtbl_codenoRow
                            rowDummy.Primary_Product = RowTable_.Primary_Product
                            rowDummy.Code_No = Me.ObjProduct.GetProductCode(RowTable_.Primary_Product)
                            rowDummy.Desc = Me.ObjProduct.GetDescription(RowTable_.Primary_Product)
                            TBLDummy.Rows.Add(rowDummy)
                        Else
                            rowDummy = TBLDummy.DefaultView(0).Row
                        End If
                        rowDGV.Cells(Me.ColWOTbl_CodeNo.Index).Value = rowDummy.Code_No
                        rowDGV.Cells(Me.ColWOTbl_Desc.Index).Value = rowDummy.Desc
                        rowDGV.Cells(Me.ColWOTbl_IdentifierJO.Index).Value = RowTable_.Identifier_JO
                        rowDGV.Cells(Me.ColWOTbl_IdentifierWO.Index).Value = RowTable_.Identifier
                        rowDGV.Cells(Me.ColWOTbl_No.Index).Value = rowDGV.Index + 1
                        rowDGV.Cells(Me.ColWOTbl_PrimaryProcess.Index).Value = RowTable_.Primary_Process
                        TBLDummy_Process.DefaultView.RowFilter = "[Primary Process] = " + RowTable_.Primary_Process.ToString
                        If TBLDummy_Process.DefaultView.Count = 0 Then
                            rowDummy_Process = TBLDummy_Process.Newtbl_processRow
                            rowDummy_Process.Primary_Process = RowTable_.Primary_Process
                            rowDummy_Process.Code = Me.objProses.GetCode(RowTable_.Primary_Process)
                            rowDummy_Process.Desc = Me.objProses.GetDesc(rowDummy_Process.Primary_Process)
                            TBLDummy_Process.Rows.Add(rowDummy_Process)
                        Else
                            rowDummy_Process = TBLDummy_Process.DefaultView(0).Row
                        End If
                        rowDGV.Cells(Me.ColWOTbl_ProcessCode.Index).Value = rowDummy_Process.Code
                        rowDGV.Cells(Me.ColWOTbl_ProcessDesc.Index).Value = rowDummy_Process.Desc
                        rowDGV.Cells(Me.ColWOTbl_Progress.Index).Value = Me.ObjCD.DigitGrouping((RowTable_.Qty_Actual_Prod / RowTable_.Qty_Target_Prod) * 100, 2)
                        rowDGV.Cells(Me.ColWOTbl_Remarks.Index).Value = RowTable_.Remarks 
                        'If RowTable_.Status = ClsTrans_ManufacturingOrder.EnumPilihStatusWO.Free Then
                        '    Status_ = "Free"
                        'ElseIf RowTable_.Status = ClsTrans_ManufacturingOrder.EnumPilihStatusWO.Planned Then
                        '    Status_ = "Planned"
                        'ElseIf RowTable_.Status = ClsTrans_ManufacturingOrder.EnumPilihStatusWO.Scheduled Then
                        '    Status_ = "Scheduled"
                        'ElseIf RowTable_.Status = ClsTrans_ManufacturingOrder.EnumPilihStatusWO.On_Progress Then
                        '    Status_ = "On Progress"
                        'ElseIf RowTable_.Status = ClsTrans_ManufacturingOrder.EnumPilihStatusWO.Finish Then
                        '    Status_ = "Finish"
                        'End If
                        'rowDGV.Cells(Me.ColWOTbl_Status.Index).Value = Status_
                        rowDGV.Cells(Me.ColWOTbl_WONO.Index).Value = RowTable_.WO_Number
                        Me.TBLWorkUnit.DefaultView.RowFilter = "[identifier wo] = '" + RowTable_.Identifier + "'"
                        If Me.TBLWorkUnit.DefaultView.Count = 0 Then
                            MsgBox("Logic error, work unit not found.", MsgBoxStyle.Information, "Information")
                        ElseIf Me.TBLWorkUnit.DefaultView.Count = 1 Then
                            Me.rowWorkUnit = Me.TBLWorkUnit.DefaultView(0).Row
                            rowDGV.Cells(Me.ColWOTbl_WorkUnit.Index).Value = Me.rowWorkUnit.work_unit_code
                            rowDGV.Cells(Me.ColWOTbl_PrimaryWorkunit.Index).Value = Me.rowWorkUnit.primary_work_unit
                        Else
                            rowDGV.Cells(Me.ColWOTbl_WorkUnit.Index).Value = "Multiple"
                        End If

                        rowDGV.Cells(Me.ColWOTbl_StartDate.Index).Value = Format(RowTable_.Plan_Start_Time, "dd MMM yyyy")
                        rowDGV.Cells(Me.ColWOTbl_StartTime.Index).Value = Format(RowTable_.Plan_Start_Time, "HH.mm")
                        rowDGV.Cells(Me.ColWOTbl_EndDate.Index).Value = Format(RowTable_.Plan_End_Time, "dd MMM yyyy")
                        rowDGV.Cells(Me.ColWOTbl_EndTime.Index).Value = Format(RowTable_.Plan_End_Time, "HH.mm")
                    End If
                Next
            End If

            If vDGV.RowCount > 0 Then
                Me.ColWOTbl_StartDate.Visible = True
                Me.ColWOTbl_StartTime.Visible = True
                Me.ColWOTbl_EndDate.Visible = True
                Me.ColWOTbl_EndTime.Visible = True
            End If
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanWorkUnitCode()
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemWO_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItemWO.CellContentClick
        Dim ij As New FrmTrans_ProductionSchedule_Reorder
        Dim identifierJO As String = ""
        Dim primaryWorkunit As Double = -1
        Dim rowSchBar0 As DseTrans_Schedule.tbl_sch_barRow
        Try
            If e.ColumnIndex = Me.ColWOTbl_Modify.Index Then
                If Me.dgvItemWO.Item(Me.ColWOTbl_WorkUnit.Index, e.RowIndex).Value = "Multiple" Then
                    Dim ij2 As New FrmTrans_StockBundling_Quick2
                    ij2.IdentifierJO = Me.dgvItemWO.Item(Me.ColWOTbl_IdentifierJO.Index, e.RowIndex).Value
                    ij2.IdentifierWO = Me.dgvItemWO.Item(Me.ColWOTbl_IdentifierWO.Index, e.RowIndex).Value
                    If ij2.ShowDialog = Windows.Forms.DialogResult.OK Then
                        Dim rowWU As DseTrans_ManufacturingOrder.DataTable10Row
                        rowWU = ij2.TBLWU.DefaultView(ij2.cboWorkUnit.SelectedIndex).Row
                        primaryWorkunit = rowWU.primary_work_unit
                    Else
                        Exit Sub
                    End If
                Else
                    primaryWorkunit = Me.dgvItemWO.Item(Me.ColWOTbl_PrimaryWorkunit.Index, e.RowIndex).Value
                End If
                ij.UserCode = Me.UserCode
                ij.TblSchBar = Me.TblSchBar.Copy()
                'ij.TblSchBarTanpaWO = Me.TblSchBar_TanpaItemWO
                ij.ObjSch = Me.objSch
                ij.PrimaryWU = primaryWorkunit
                ij.IdentifierJO = Me.dgvItemWO.Item(Me.ColWOTbl_IdentifierJO.Index, e.RowIndex).Value
                ij.IdentifierWO = Me.dgvItemWO.Item(Me.ColWOTbl_IdentifierWO.Index, e.RowIndex).Value
                If ij.ShowDialog = Windows.Forms.DialogResult.Yes Then
                    ij.TblSchBar.DefaultView.RowFilter = "[primary work unit] = " & ij.PrimaryWU.ToString
                    For Each row As Data.DataRowView In ij.TblSchBar.DefaultView
                        rowSchBar0 = row.Row
                        Me.objSch.UpdatePlanStartEndSchedule(rowSchBar0.plan_start_time, rowSchBar0.plan_end_time, rowSchBar0.identifier_jo, rowSchBar0.identifier_wo, ij.PrimaryWU)
                        Me.objSch.UpdatePlanStartEndScheduleWO(rowSchBar0.identifier_jo, rowSchBar0.identifier_wo)
                        Me.tblSchBar.DefaultView.RowFilter = "[primary work unit] = " & ij.PrimaryWU.ToString & " and [identifier jo] = '" & rowSchBar0.identifier_jo & "' and [identifier wo] = '" & rowSchBar0.identifier_wo & "'"
                        If Me.tblSchBar.DefaultView.Count > 0 Then
                            Me.rowSch = Me.tblSchBar.DefaultView(0).Row
                            Me.rowSch.plan_start_time = rowSchBar0.plan_start_time
                            Me.rowSch.plan_end_time = rowSchBar0.plan_end_time
                        End If
                    Next
                    Me.TblItemWO.Rows.Clear()
                    Me.ObjMOItem.IsiDataItemWO(Me.IdentifierJO, Me.TblItemWO)
                    Me.LoadDgvItemWO_TableView(Me.IdentifierJO, Me.TblItemWO, Me.dgvItemWO)
                    Me.AnyChange = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbColumnSetting_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbColumnSetting.LinkClicked
        Dim ij As New FrmTrans_ProductionSchedule_DetailWO_ColSetting
        Try
            ij.UserCode = Me.UserCode
            If ij.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.SetColumnVisibility()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub SetColumnVisibility()
        Try
            ShowCodeNo = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_DetailWO, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_DetailWO_ShowCodeNo, Me.UserCode)
            ShowDesc = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_DetailWO, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_DetailWO_ShowDesc, Me.UserCode)
            ShowProcess = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_DetailWO, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_DetailWO_ShowProcess, Me.UserCode)
            ShowProcessDesc = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ProductionSchedule_DetailWO, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ProductionSchedule_DetailWO_ShowProcessDesc, Me.UserCode)
            If ShowCodeNo = "1" Or ShowCodeNo = "" Then
                Me.ColWOTbl_CodeNo.Visible = True
            Else
                Me.ColWOTbl_CodeNo.Visible = False
            End If
            If ShowDesc = "1" Or ShowDesc = "" Then
                Me.ColWOTbl_Desc.Visible = True
            Else
                Me.ColWOTbl_Desc.Visible = False
            End If
            If ShowProcess = "1" Or ShowProcess = "" Then
                Me.ColWOTbl_ProcessCode.Visible = True
            Else
                Me.ColWOTbl_ProcessCode.Visible = False
            End If
            If ShowProcessDesc = "1" Or ShowProcessDesc = "" Then
                Me.ColWOTbl_ProcessDesc.Visible = True
            Else
                Me.ColWOTbl_ProcessDesc.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class