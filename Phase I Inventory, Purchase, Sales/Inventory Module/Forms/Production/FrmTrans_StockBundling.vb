Option Explicit On
Option Strict Off

Imports Core_Module
Public Class FrmTrans_StockBundling
    Public UserCode As String = "USER"
    Public ObjStockBundling As New ClsTrans_StockBundling
    Public PrimaryBundling As Double = -1
    Public _IdentifierWO As String = ""
    Public _IdentifierJO As String = ""
    Public _PrimaryWorkUnit As Double = -1
    Public _PrimaryBundling As Double = -1
    Public _Command As EnumCommand = EnumCommand.NOTHING_
    Public _Stop As Boolean = False

    Dim BolehOverlapSch As Boolean = False
    Dim ObjUser As New Core_Module.ClsUser
    Dim ObjLicense As New Core_Module.ClsLicense

    Public Enum EnumCommand As Integer
        NOTHING_
        NEW_
        EDIT
        DELETE
        VIEW
        NEW2
    End Enum

    Dim PrimaryInputLocation As Double = -1
    Dim PrimaryOutputLocation As Double = -1
    Dim InputLocation As String = String.Empty
    Dim InputLocationDesc As String = String.Empty
    Dim OutputLocation As String = String.Empty
    Dim IsWorkShiftRequired As Boolean = False
    Dim ObjWorkShift As New Core_Module.ClsMasterWorkShift

    Dim SedangAddRM As Boolean = False 
    Dim ObjItemInOut As New clsTransaksiStockInOut
    Dim ObjItemBundling As New clsTrans_StockBundlingItem
    Dim ObjCL As New ClsTransaksiMonthlyClosingItem
    Dim TBLLocationPriority As New Core_Module.DseMasterLocationPriority.tblmaster_location_priority1_DataTable
    Dim ObjLocationPriority As New Core_Module.ClsMasterLocationPriority
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll

    Dim baru As Boolean
    Dim ObjCD As New ClsCekData
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting
    Dim ObjCustom As New CustomSMESS.ClsTransaksiStockBundling
    Dim ObjProduct As New clsMasterProduct
    Dim objStockBundlingItem As New clsTrans_StockBundlingItem
    Dim TBL As New dseMasterProduct.tblmaster_productqty1DataTable
    Dim TBL2 As New dseMasterProduct.stock_location_viewDataTable
    Dim TBL3 As New dseMasterProduct.tblmaster_productqty2DataTable
    Dim TBL4 As New DseTrans_StockBundling.tbltransaksi_stockbundling_item_before1DataTable
    Dim TBL5 As New DseTrans_StockBundling.tbltransaksi_stockbundling_item_before2DataTable
    Dim TBL6 As New DseTrans_StockBundling.tbltransaksi_stockbundling_item_before3DataTable
    Dim buffer As String = ""

    '===================================tambahan berkaitan produksi===================================
    Public isProductCode As String = ""
    Public isBalanceProduce As Decimal = 0
    Public isDecimal As Decimal = 0
    Public isBalance_Qtyrm As Decimal = 0

    Dim TBL_BOM As New DseTrans_StockBundling.tbltrans_mo_item_bomDataTable
    Dim row_BOM As DseTrans_StockBundling.tbltrans_mo_item_bomRow
    Dim tbl_schedule As New DseTrans_StockBundling.tblmaster_workshift_itemDataTable
    Dim row_schedule As DseTrans_StockBundling.tblmaster_workshift_itemRow
    Dim tipe_bundling As ClsTrans_StockBundling.EnumTipeBundling = ClsTrans_StockBundling.EnumTipeBundling.Produksi1
    Dim tbl_workunit As New DseTrans_StockBundling.tblworkunitDataTable
    Dim row As DseTrans_StockBundling.tblworkunitRow
    Dim vscheduled As ClsTrans_ManufacturingOrder.EnumPilihStatusWO = ClsTrans_ManufacturingOrder.EnumPilihStatusWO.Scheduled
    Dim objdowntime As New Core_Module.ClsMasterDownTime
    Dim objmachine As New Core_Module.ClsMasterMachineData
    Dim objemployee As New Core_Module.ClsMasterKaryawan
    Dim objsetupTime As New Core_Module.ClsMasterSetupTime
    Dim objreject As New Core_Module.ClsMasterReject
    Dim eraserowindex_downtime As Integer = 0
    Dim eraserowindex_rm As Integer = 0
    Dim eraserowindex_machine As Integer = 0
    Dim eraserowindex_employee As Integer = 0
    Dim eraserowindex_setuptime As Integer = 0
    Dim tbl_rm As New DseTrans_StockBundling.tblgets_rmDataTable
    Dim row_rm As DseTrans_StockBundling.tblgets_rmRow
    Dim tbl_wo_4Header As New DseTrans_StockBundling.tblwoitemDataTable
    Dim row_wo_4Header As DseTrans_StockBundling.tblwoitemRow
    Dim tbl_wo_4Item As New DseTrans_StockBundling.tbltrans_mo_item_wo_item_sch1DataTable
    Dim row_wo_4Item As DseTrans_StockBundling.tbltrans_mo_item_wo_item_sch1Row
    Dim buffer_pj As String = ""
    Dim buffer_rmreject As String = ""
    Dim buffer_downtime As String = ""
    Dim buffer_employee As String = ""
    Dim buffer_machine As String = ""
    Dim buffer_setuptime As String = ""
    Dim eraserowindex_pj As Integer = 0
    Dim vr As DataGridViewRow
    Dim buffer_st As String = "00.01"
    Dim BufferWorkHour As String = "0"
    Dim PrimaryWorkUnit_ReffOfBundling As Double = -1
    Dim Buffer_ActualStartDate As Date = Now
    Dim Buffer_ActualEndDate As Date = Now
    Dim BufferTime_Start As New Date
    Dim BufferTime_End As New Date
    '========================================================================

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryStockBundling.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer4.Panel1.Controls, Me.TxtPrimaryStockBundling.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer4.Panel2.Controls, Me.TxtPrimaryStockBundling.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gboItemb4.Controls, Me.TxtPrimaryStockBundling.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gboItemAfter.Controls, Me.TxtPrimaryStockBundling.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer5.Panel1.Controls, Me.TxtPrimaryStockBundling.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer5.Panel2.Controls, Me.TxtPrimaryStockBundling.BackColor, False)
            Me.ObjCD.SetBackColor(Me.tabpage_rejectdowntimeoperdata.Controls, Me.TxtPrimaryStockBundling.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.TxtPrimaryStockBundling.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gbo_actual.Controls, Me.TxtPrimaryStockBundling.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gbo_plan.Controls, Me.TxtPrimaryStockBundling.BackColor, False)
            Me.ObjCD.SetBackColor(Me.tabpage_processoutput.Controls, Me.TxtPrimaryStockBundling.BackColor, False)

            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer4.Panel1.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer4.Panel2.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.gboItemb4.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.gbo_plan.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.gbo_actual.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.gboItemAfter.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer5.Panel1.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer5.Panel2.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.tabpage_rejectdowntimeoperdata.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.tabpage_processoutput.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)

            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            If Me.dgvItemb4.Rows.Count = 0 Then
                Me.btnUPb4.Enabled = False
                Me.btnDownb4.Enabled = False
            End If
            If Me.dgvItemAfter.Rows.Count = 0 Then
                Me.btnUpAfter.Enabled = False
                Me.btnDownAfter.Enabled = False
            End If
            '
            Me.SplitContainer1.TabStop = False
            Me.SplitContainer2.TabStop = False
            Me.SplitContainer3.TabStop = False

            Me.dtpplan_start.Enabled = False
            Me.dtpplan_end.Enabled = False
            Me.dgvItemb4.AllowUserToDeleteRows = True
            Me.dgvItemb4.ReadOnly = False
            Me.dgvItemAfter.AllowUserToDeleteRows = True
            Me.dgvItemAfter.ReadOnly = False
            '
            Me.dgvinput_employee.TabStop = True
            Me.dgvinput_machine.TabStop = True
            Me.dgvinput_processreject.TabStop = True
            Me.dgvinput_downtime.TabStop = True
            Me.dgvinput_employee.TabStop = True
            Me.dgvInput_SetupTime.TabStop = True
            '
            Me.dgvinput_employee.AllowUserToDeleteRows = True
            Me.dgvinput_machine.AllowUserToDeleteRows = True
            Me.dgvinput_processreject.AllowUserToDeleteRows = True
            Me.dgvinput_downtime.AllowUserToDeleteRows = True
            Me.dgvinput_rmreject.AllowUserToDeleteRows = True
            Me.dgvInput_SetupTime.AllowUserToDeleteRows = True
            '
            Me.dgvinput_employee.ReadOnly = False
            Me.dgvinput_machine.ReadOnly = False
            Me.dgvinput_processreject.ReadOnly = False
            Me.dgvinput_downtime.ReadOnly = False
            Me.dgvinput_rmreject.ReadOnly = False
            Me.dgvInput_SetupTime.ReadOnly = False
            Me.TxtPrimaryStockBundling.BackColor = Me.TxtPrimaryCodeNoAfter.BackColor
            Me.txttotal_produce.Enabled = True
            Me.txttotal_good.Enabled = True
            Me.txttotal_reject.Enabled = True
            Me.txtgood_perct.Enabled = True
            Me.AturWoNumberdanCboWorkUnit()
            '
            Me.DtpExpiredb4.Enabled = False
            Me.llb_BOM.Enabled = True

            Me.SetGboItemAfter()
            Me.SetGboItemb4()
            Me.FormatDGV(Me.dgvItemAfter)
            Me.FormatDGV(Me.dgvItemb4)

            Me.btnCancel.Enabled = True
            Me.btnCancelTP1.Enabled = True
            Me.btnCancelTP2.Enabled = True
            Me.btnOK.Enabled = True
            Me.ForceNavigate = True
            Me.tc_processoutput.SelectedTab = Me.tabpage_processoutput

            Me.dgvinput_processreject.AllowUserToAddRows = True
            Me.dgvinput_downtime.AllowUserToAddRows = True
            Me.dgvInput_SetupTime.AllowUserToAddRows = True
            If Not IsWorkShiftRequired Then
                Me.TxtWorkShiftCode.Enabled = False
                Me.BtnFindWorkShift.Enabled = False
            End If

            Me.btnNext.Enabled = True
            Me.btnPrev.Enabled = True

            If Me.btnStop.Visible = True Then
                Me.btnStop.Enabled = True
            End If

            Me.BolehOverlapSch = Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiStockBundling2_OverlapSchedule")
            'Me.btnNextTab1.Visible = True
            'Me.btnNextTab2.Visible = True
            'Me.btnPrevTab2.Visible = True
            'Me.btnPrevTab3.Visible = True

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            ModCekData.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer4.Panel1.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer4.Panel2.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.gboItemb4.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.gboItemAfter.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer5.Panel1.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer5.Panel2.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.tabpage_rejectdowntimeoperdata.Controls, True, False, True, False, True, False, True, False)
            Me.btnPrevTab2.Enabled = True
            Me.btnNextTab2.Enabled = True
            ModCekData.KendalikanControl(Me.tabpage_processoutput.Controls, True, False, True, False, True, False, True, False)
            Me.btnNextTab1.Enabled = True
            ModCekData.KendalikanControl(Me.gbo_actual.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.gbo_plan.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)
            Me.CancelButton = Me.btnExit
            Me.dgvItemb4.AllowUserToDeleteRows = False
            Me.dgvItemb4.ReadOnly = True
            Me.dgvItemAfter.AllowUserToDeleteRows = False
            Me.dgvItemAfter.ReadOnly = True
            '
            Me.SplitContainer1.TabStop = True
            Me.SplitContainer2.TabStop = True
            Me.SplitContainer3.TabStop = True
            Me.dgvinput_employee.ReadOnly = True
            Me.dgvinput_machine.ReadOnly = True
            Me.dgvinput_processreject.ReadOnly = True
            Me.dgvinput_downtime.ReadOnly = True
            Me.dgvinput_rmreject.ReadOnly = True
            Me.dgvInput_SetupTime.ReadOnly = True
            '
            Me.dgvinput_employee.TabStop = False
            Me.dgvinput_machine.TabStop = False
            Me.dgvinput_processreject.TabStop = False
            Me.dgvinput_downtime.TabStop = False
            Me.dgvinput_employee.TabStop = False
            Me.dgvInput_SetupTime.TabStop = False
            '
            Me.dgvinput_employee.AllowUserToDeleteRows = False
            Me.dgvinput_machine.AllowUserToDeleteRows = False
            Me.dgvinput_processreject.AllowUserToDeleteRows = False
            Me.dgvinput_downtime.AllowUserToDeleteRows = False
            Me.dgvinput_rmreject.AllowUserToDeleteRows = False
            Me.dgvInput_SetupTime.AllowUserToDeleteRows = False
            Me.txttotal_produce.Enabled = False
            Me.txttotal_good.Enabled = False
            Me.txttotal_reject.Enabled = False
            Me.txtgood_perct.Enabled = False
            '
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer4.Panel1.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer4.Panel2.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gboItemb4.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gboItemAfter.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer5.Panel1.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer5.Panel2.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.tabpage_rejectdowntimeoperdata.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gbo_plan.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gbo_actual.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.tabpage_processoutput.Controls, Me.BackColor, False)
            Me.llb_BOM.Enabled = False
            Me.btnCancel.Enabled = False
            Me.btnCancelTP1.Enabled = False
            Me.btnCancelTP2.Enabled = False
            Me.btnOK.Enabled = False
            Me.btnPrev.Enabled = True
            Me.btnNext.Enabled = True
            Me.dgvinput_processreject.AllowUserToAddRows = False
            Me.dgvinput_downtime.AllowUserToAddRows = False
            Me.dgvInput_SetupTime.AllowUserToAddRows = False
            Me.btnNext.Enabled = False
            Me.btnPrev.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField(Optional ByVal headerwosaja As Boolean = False)
        Try
            Me.KosongkanItemB4()
            Me.KosongkanItemAfter()
            Me.KosongkanItemProcess()

            '->header
            If headerwosaja = False Then
                ModCekData.KosongkanControl(Me.gboDetail.Controls, True, True, False)
                ModCekData.KosongkanControl(Me.SplitContainer4.Panel1.Controls, True, True, False)
                ModCekData.KosongkanControl(Me.SplitContainer4.Panel2.Controls, True, True, False)
                'Me.dtpDate.Value = Now
                Me.TxtPrimaryStockBundling.Text = "-1"
                Me.txtstockbundling.Text = ""
            End If
            Me.TxtPrimaryWorkShift.Text = "-1"

            Me.txtidentifier_wo.Text = ""
            Me.TxtIdentifierJO.Text = ""

            Me.txtprimarywu.Text = "-1"
            Me.txtprimaryprocess.Text = "-1"
            Me.txtprimaryws.Text = "-1"
            Me.CboWorkUnit.DataSource = Nothing
            Me.CboWorkUnit.Items.Clear()
            Me.txtwonumber.Clear()
            Me.txtworkunitDesc.Clear()
            Me.txtprocess_code.Clear()
            Me.txtprocess_desc.Clear()
            Me.txtmo_number.Clear()
            Me.txtjo_number.Clear()
            Me.txttotal_produce.Text = "0"
            Me.txtgood_perct.Text = "0.00"
            Me.txttotal_good.Text = "0"
            Me.txttotal_reject.Text = "0"
            Me.dtpactual_start.Value = Now
            Me.dtpactual_end.Value = Now
            Me.txtactual_workhour.Text = "00.00"
            Me.txtactual_setuptime.Text = "00.00"
            Me.txtactual_downtime.Text = "00.00"
            Me.txtactual_total.Text = "00.00"
            Me.txtactual_prodrate.Text = "0.00"
            Me.dtpplan_start.Value = Now
            Me.dtpplan_end.Value = Now
            Me.txtplan_workhour.Text = "00.00"
            Me.txtplan_setuptime.Text = "00.00"
            Me.txtplan_total.Text = "00.00"
            Me.txtplan_prodrate.Text = "0.00"
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
            Me.dgvItemb4.Rows.Clear()
            Me.dgvItemAfter.Rows.Clear()
            Me.PrimaryWorkUnit_ReffOfBundling = -1
            Me.tc_processoutput.SelectedTab = Me.tabpage_processoutput
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField2()
        Try
            Me.KosongkanItemB4()
            Me.KosongkanItemAfter()
            Me.KosongkanItemProcess()

            Me.txtidentifier_wo.Text = ""
            Me.TxtIdentifierJO.Text = ""
            Me.txtprimarywu.Text = "-1"
            Me.txtprimaryprocess.Text = "-1"
            Me.txtprimaryws.Text = "-1"
            Me.CboWorkUnit.DataSource = Nothing
            Me.CboWorkUnit.Items.Clear()
            Me.txtwonumber.Clear()
            Me.txtworkunitDesc.Clear()
            Me.txtprocess_code.Clear()
            Me.txtprocess_desc.Clear()
            Me.txtmo_number.Clear()
            Me.txtjo_number.Clear()
            Me.txttotal_produce.Text = "0"
            Me.txtgood_perct.Text = "0.00"
            Me.txttotal_good.Text = "0"
            Me.txttotal_reject.Text = "0"
            Me.dtpactual_start.Value = Now
            Me.dtpactual_end.Value = Now
            Me.txtactual_workhour.Text = "00.00"
            Me.txtactual_setuptime.Text = "00.00"
            Me.txtactual_downtime.Text = "00.00"
            Me.txtactual_total.Text = "00.00"
            Me.txtactual_prodrate.Text = "0.00"
            Me.dtpplan_start.Value = Now
            Me.dtpplan_end.Value = Now
            Me.txtplan_workhour.Text = "00.00"
            Me.txtplan_setuptime.Text = "00.00"
            Me.txtplan_total.Text = "00.00"
            Me.txtplan_prodrate.Text = "0.00"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanItemAfter()
        Try
            ModCekData.KosongkanControl(Me.gboItemAfter.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer5.Panel1.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer5.Panel2.Controls, True, True, False)
            Me.TxtPrimaryCodeNoAfter.Text = "-1"
            Me.txtPrimaryUnitsAfter.Text = "-1"
            Me.txtPrimaryPDOAfter.Text = "-1"
            Me.txtPrimarySecondaryUnitsAfter.Text = "-1"
            Me.txtPrimaryLocationAfter.Text = "-1"
            Me.txtPrimaryCategoryAfter.Text = "-1"
            Me.TxtQtyperBatchAfter.Text = "0"
            Me.txtSecondaryQtyAfter.Text = "0"
            Me.txtFaktorPemberat.Text = "0"
            Me.chkBatchNoControlledAfter.Checked = False
            Me.chkExpiryDateControlledAfter.Checked = False
            Me.chkLocationControlledAfter.Checked = False
            Me.chkPdoControlledAfter.Checked = False
            Me.chkUseSecondaryUnitAfter.Checked = False
            Me.txtSecondaryQtyAfter.Visible = False
            Me.lblSecondaryUnitsAfter.Visible = False
            'Me.TxtQtyperBatchAfter.Width = 239
            Me.lblUnitsAfter.Left = Me.lblSecondaryUnitsAfter.Left
            Me.lblSecQtyAfter.Visible = False
            Me.lblExpiryDateAfter.Visible = False
            Me.DtpExpiredAfter.Value = Today
            Me.DtpExpiredAfter.Visible = True
            Me.lblLocationAfter.Visible = False
            Me.lblWarehouseAfter.Visible = False
            Me.TxtLocationAfter.Visible = True
            Me.TxtWarehouseAfter.Visible = True
            Me.btnFindLocationAfter.Visible = True
            Me.lblBatchNoAfter.Visible = False
            Me.txtBatchNoAfter.Visible = True
            Me.txtBatchNoAfter.TabStop = True
            Me.txtBatchNoAfter.ReadOnly = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanItemB4()
        Try
            ModCekData.KosongkanControl(Me.gboItemb4.Controls, True, True, False)
            Me.TxtPrimaryCodeNob4.Text = "-1"
            Me.txtPrimaryUnitsb4.Text = "-1"
            Me.txtPrimaryPDOb4.Text = "-1"
            Me.txtPrimarySecondaryUnitsb4.Text = "-1"
            Me.txtPrimaryLocationb4.Text = "-1"
            Me.txtPrimaryCategoryb4.Text = "-1"
            Me.TxtQtyperBatchb4.Text = "0"
            Me.txtSecondaryQtyb4.Text = "0"
            Me.chkBatchNoControlledb4.Checked = False
            Me.chkExpiryDateControlledb4.Checked = False
            Me.chkLocationControlledb4.Checked = False
            Me.chkPdoControlledb4.Checked = False
            Me.chkUseSecondaryUnitb4.Checked = False
            Me.lblSecondaryUnitsb4.Visible = False
            Me.TxtQtyperBatchb4.Width = 239
            Me.lblUnitsb4.Left = Me.lblSecondaryUnitsb4.Left
            Me.lblSecQtyb4.Visible = False
            Me.lblExpiryDateb4.Visible = False
            Me.DtpExpiredb4.Visible = True
            Me.DtpExpiredb4.Value = Today
            Me.lblLocationb4.Visible = False
            Me.lblWarehouseb4.Visible = False
            Me.cboLocationb4.DataSource = Nothing
            Me.cboLocationb4.Visible = True
            Me.cboLocationb4.Items.Clear()
            Me.TxtWarehouseb4.Visible = True
            Me.lblBatchNob4.Visible = False
            Me.cboBatchNob4.Visible = True
            Me.cboBatchNob4.TabStop = True
            Me.cboBatchNob4.DataSource = Nothing
            Me.cboBatchNob4.Items.Clear()
            Me.cboSecondaryUnitsb4.Visible = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Function ApakahDataItemB4Benar(ByVal vCekBalanceRM As Boolean) As Boolean
        Dim ObjSC As New ClsTransaksiStockCard

        Try
            'data item salah bila:
            '1. product code yang diisikan salah
            If Me.txtDescb4.Text = "" Then
                MsgBox("Please enter code no.", MsgBoxStyle.Information, "Information")
                Me.TxtCodeNob4.Text = ""
                Me.TxtCodeNob4.Focus()
                Return False
            End If

            '----

            'product harus termasuk wo ybs
            If ObjStockBundling.ApakahCodeNoDiRMSudahAda(txtidentifier_wo.Text, TxtPrimaryCodeNob4.Text, Me.TxtIdentifierJO.Text) = False Then
                MsgBox("Raw material product code is not member of output product.", MsgBoxStyle.Information, "Information")
                Return False
            End If
            '---

            '2. nilai qty <= 0
            If Me.ObjCD.CekBilanganPositif(Me.TxtQtyperBatchb4.Text) = False Then
                MsgBox("Please enter a positive number for quantity data.", MsgBoxStyle.Information, _
                "Information")
                Me.TxtQtyperBatchb4.Focus()
                Return False
            End If

            '3. Bila batch no controlled, batch no tidak diisikan
            If Me.chkBatchNoControlledb4.Checked = True Then
                If Me.cboBatchNob4.Text = "" Then
                    MsgBox("Please enter batch no, this product is batch no controlled.", MsgBoxStyle.Information, "Information")
                    Me.cboBatchNob4.Focus()
                    Return False
                End If
            End If
            '4.a Untuk transaksi proses (in/back/out proses) no Pdo harus diisikan
            '4.b Untuk transaksi yang berkaitan dengan purchase (stock out purchase return dan stout in purchase)
            '    harus menyertakan no PO bila terhubungkan dengan software purchase
            '4.c Untuk transaksi yang berkaitan dengan sales (stock out sales dan stock in sales return) harrus
            '    menyertakan no DO bila terhubungkan dengan software sales
            '5. Bila location controlled, kode lokasi tidak diisikan 
            If Me.chkLocationControlledb4.Checked = True Then
                If Me.txtPrimaryLocationb4.Text = "-1" Then
                    MsgBox("Please enter location, this product is location controlled.", MsgBoxStyle.Information, "Information")
                    Me.cboLocationb4.Focus()
                    Return False
                End If
            End If

            ''5.a Jika strict-location tak boleh ambil RM dari tempat selain yang sudah ditentukan
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiStockBundling_NoStrictLocation2") = False Then
                If Me.txtPrimaryLocationb4.Text <> Me.PrimaryInputLocation Then
                    MsgBox("You don't have access to location '" & Me.cboLocationb4.Text & "'.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            '6. Bila expiry date controlled, expiry date back date dibandingkan dengan tanggal transaksi
            '   hanya diberikan peringatan, karena mungkin saja mau jual barang expired dengan harga 
            '   yang sangat murah -  atau adjust untuk dikeluarkan dari data komputer
            If Me.chkExpiryDateControlledb4.Checked = True Then
                If Me.DtpExpiredb4.Value.Date < Me.dtpDate.Value.Date Then
                    If MsgBox("This product expiry date is smaller than transaction date. Do you want to continue?", _
                    MsgBoxStyle.YesNo, "Expired Product") = MsgBoxResult.No Then
                        Me.DtpExpiredb4.Focus()
                        Return False
                    End If
                End If
            End If
            '7. Bila use secondary unit, nilai qty secondary <=0
            If Me.chkUseSecondaryUnitb4.Checked = True Then
                If Me.ObjCD.CekBilanganPositif(Me.txtSecondaryQtyb4.Text) = False Then
                    MsgBox("Please enter a positive number for secondary quantity data.", MsgBoxStyle.Information, "Information")
                    Me.txtSecondaryQtyb4.Focus()
                    Return False
                End If
            End If
            '8a. Bila product yang bersangkutan batch no controlled dan expiry date controlled dan expiry date
            '   yang dimasukkan pada tidak sama dengan expiry date yang ada didalam basis data maka data salah
            If Me.chkBatchNoControlledb4.Checked = True And Me.chkExpiryDateControlledb4.Checked = True Then
                Dim ExpiryDateLama As Date = Me.ObjProduct.GetExpiryDate(Me.TxtPrimaryCodeNob4.Text, Me.cboBatchNob4.Text)
                If ExpiryDateLama <> #1/1/1950# Then
                    If ExpiryDateLama.Date <> Me.DtpExpiredb4.Value.Date Then
                        MsgBox("Expiry date for batch no: " + Me.cboBatchNob4.Text + " is " + _
                        Format(ExpiryDateLama, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                        Me.DtpExpiredb4.Focus()
                        Return False
                    End If
                End If
                '8b. Bila product yang bersangkutan batch no controlled dan expiry date controlled dan expiry date yang dimasukkan
                '    tidak sama dengan expiry date yang ada sudah ada didalam dgv
                Dim row As DataGridViewRow
                For Each row In Me.dgvItemb4.Rows
                    If row.Cells(7).Value = Me.cboBatchNob4.Text And Me.cboBatchNob4.Text <> "" And row.Cells(12).Value = _
                    Me.TxtPrimaryCodeNob4.Text Then
                        If CDate(row.Cells(8).Value) <> Me.DtpExpiredb4.Value.Date Then
                            MsgBox("You already entered expiry date for batch no: " + Me.cboBatchNob4.Text + " on " + _
                            Format(CDate(row.Cells(8).Value), "dd MMM yyyy") + ". Please change your expiry date data.", _
                            MsgBoxStyle.Information, "Information")
                            Me.DtpExpiredb4.Focus()
                            Return False
                        End If
                    End If
                Next
                '8.c Bila product yang bersangkutan batch no controlled dan expiry date controlled dan expiry date yang dimasukkan
                '    tidak sama dengan expiry date yang ada sudah ada didalam transaksi lainnya: stock in adjust, stock in purchase, 
                '    stock in sales return, stock tag, transaksi bundling (item before) 
                '    (untuk stock out tidak dicek karena semua harus masuk dulu sebelum keluar, jadi dimasukka ke dalam pengecekan
                '    adalah semua transaksi yang menjadi pintu masuk inventory)
                'msgbox("belum")
            End If

            '-cek ke balance qty rm
            If vCekBalanceRM = True Then
                If ApakahQtyRMUsageSudahBenar(Me.TxtPrimaryStockBundling.Text, Me.TxtPrimaryCodeNob4.Text, Me.txtidentifier_wo.Text, Me.TxtQtyperBatchb4.Text, False, Me.TxtIdentifierJO.Text) = False Then
                    Return False
                End If
            End If

            '9. Bila setelah diinput jumlah product menjadi minus
            Dim JmlGudang, JmlObjLama, JmlSetelahTambah, JmlDGV As Decimal
            JmlGudang = Me.ObjProduct.GetJml3(Me.TxtPrimaryCodeNob4.Text, Me.txtSecondaryQtyb4.Text, Me.txtPrimaryLocationb4.Text, Me.cboBatchNob4.Text)
            'perhitungan menggunakan nilai jumlah dgv yang baru
            JmlObjLama = Me.objStockBundlingItem.GetJmlB4(Me.TxtPrimaryStockBundling.Text, Me.TxtPrimaryCodeNob4.Text, _
            Me.txtSecondaryQtyb4.Text, Me.txtPrimaryLocationb4.Text, Me.cboBatchNob4.Text, Me.DtpExpiredb4.Value)
            JmlDGV = Me.GetJmlDGV(Me.TxtPrimaryCodeNob4.Text, Me.txtSecondaryQtyb4.Text, Me.txtPrimaryLocationb4.Text, _
            Me.cboBatchNob4.Text, Me.chkBatchNoControlledb4.Checked, _
            Me.chkLocationControlledb4.Checked, Me.chkUseSecondaryUnitb4.Checked, Me.dgvItemb4)
            JmlSetelahTambah = JmlGudang + JmlObjLama - JmlDGV - CDbl(Me.TxtQtyperBatchb4.Text)
            If JmlSetelahTambah < 0 Then
                If Me.chkUseSecondaryUnitb4.Checked = True Then
                    MsgBox("Maximum value you can enter is " + Me.ObjCD.DigitGrouping(CStr(JmlGudang + JmlObjLama - JmlDGV), Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNob4.Text)) + " " + _
                    Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNob4.Text) + " @ " + Me.txtSecondaryQtyb4.Text + " " + _
                    Me.ObjProduct.GetSecondaryUnitCode(Me.TxtPrimaryCodeNob4.Text) + ".", MsgBoxStyle.Information, _
                    "Information")
                Else
                    MsgBox("Maximum value you can enter is " + Me.ObjCD.DigitGrouping(CStr(JmlGudang + JmlObjLama - JmlDGV), Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNob4.Text)) + " " + _
                    CStr(Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNob4.Text)) + ".", MsgBoxStyle.Information, _
                    "Information")
                End If
                Me.TxtQtyperBatchb4.Focus()
                Return False
            End If

            '9.b. Cek stock pada tanggal yang bersangkutan
            JmlGudang = ObjSC.GetStock(Me.TxtPrimaryCodeNob4.Text, Me.dtpDate.Value)
            JmlSetelahTambah = JmlGudang + JmlObjLama - JmlDGV - CDbl(Me.TxtQtyperBatchb4.Text)
            If JmlSetelahTambah < 0 Then
                If JmlGudang + JmlObjLama - JmlDGV > 0 Then
                    If Me.chkUseSecondaryUnitb4.Checked = True Then
                        MsgBox("Maximum value you can enter " + " on " + Format(Me.dtpDate.Value, "dd MMM yyyy") + " is " + Me.ObjCD.DigitGrouping(CStr(JmlGudang + JmlObjLama - JmlDGV), Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNob4.Text)) + " " + _
                        Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNob4.Text) + " @ " + Me.txtSecondaryQtyb4.Text + " " + _
                        Me.ObjProduct.GetSecondaryUnitCode(Me.TxtPrimaryCodeNob4.Text) + ".", MsgBoxStyle.Information, _
                        "Information")
                    Else
                        MsgBox("Maximum value you can enter " + " on " + Format(Me.dtpDate.Value, "dd MMM yyyy") + " is " + Me.ObjCD.DigitGrouping(CStr(JmlGudang + JmlObjLama - JmlDGV), Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNob4.Text)) + " " + _
                        CStr(Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNob4.Text)) + ".", MsgBoxStyle.Information, _
                        "Information")
                    End If
                Else
                    MsgBox("Stock level on " + Format(Me.dtpDate.Value, "dd MMM yyyy") + _
                    " is 0. Entering this data will cause negative value in your stock.", _
                    MsgBoxStyle.Information, "Information")
                End If

                Me.TxtQtyperBatchb4.Focus()
                Return False
            End If

            '10. Bila ada item bundling after yang sama dengan item bundling before, akan mengakibatkan
            'perhitungan in amount pada bundling ngaco
            Dim rowDGV As DataGridViewRow
            For Each rowDGV In Me.dgvItemAfter.Rows
                If rowDGV.Cells(Me.ColPrimaryCodeNoAfter.Index).Value = Me.TxtPrimaryCodeNob4.Text Then
                    MsgBox("There product with code no: " + Me.TxtCodeNob4.Text + _
                    " in item after, we cannot enter this item", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Next

            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjSC = Nothing
        End Try
    End Function

    Private Function ApakahItemProcessSudahBenar(ByVal vtglmulai As Date, ByVal vtglselesai As Date, ByVal vprimary_wu As Double) As Boolean
        Dim ListNoBundlingStr As String = ""
        Try
            '->wo tidak boleh kosong
            If Me.txtwonumber.Text = "" Then
                MsgBox("WO number cannot be an empty string.", MsgBoxStyle.Information, "Information")
                Me.txtwonumber.Focus()
                Return False
            End If

            If Me.BolehOverlapSch = False Then
                Dim ObjCls_Wu As New ClsMasterWorkUnit
                If Me.ObjStockBundling.ApakahWorkTimeBundlingBenar_ByWu(Me.txtprimarywu.Text, Me.TxtPrimaryStockBundling.Text, Me.dtpactual_start.Value, Me.dtpactual_end.Value, ListNoBundlingStr, ClsTrans_StockBundling.EnumTipeBundling.Produksi1) = False Then
                    MsgBox("Actual work time of work unit " & ObjCls_Wu.GetCode(Me.txtprimarywu.Text) & " is conflict with the following production log: " & vbNewLine & ListNoBundlingStr, MsgBoxStyle.Information, "Information")
                    ObjCls_Wu = Nothing
                    Return False
                End If
            End If


            If Me.ObjCD.Is_Time_String(Me.txtactual_setuptime.Text) = False Then
                MsgBox("Please input hh.mm format for actual setup time.", MsgBoxStyle.Information, "Information")
            End If

            If Me.Apakah_Worktime(Me.dtpactual_start.Value) = False Then
                MsgBox(Format(Me.dtpactual_start.Value, "dd MMM yyyy, HH.mm") & " is not valid worktime.", MsgBoxStyle.Information, "Information")
                Me.dtpactual_start.Value = Me.BufferTime_Start
                Me.BufferWorkHour = Me.txtactual_workhour.Text
                Exit Function
            End If

            If Me.Apakah_Worktime(Me.dtpactual_end.Value) = False Then
                MsgBox(Format(Me.dtpactual_end.Value, "dd MMM yyyy, HH.mm") & " is not valid worktime.", MsgBoxStyle.Information, "Information")
                Me.dtpactual_end.Value = Me.BufferTime_End
                Me.BufferWorkHour = Me.txtactual_workhour.Text
                Exit Function
            End If

            If Me.ObjCD.Convert_ToMinutes(Me.txtactual_workhour.Text) <= 0 Then
                MsgBox("Work hour duration must be greather than 0", MsgBoxStyle.Information, "Information")
                Exit Function
            End If

            '->harus ada minimal 1 item di employee atau di machine -> hold dulu sementara
            'Me.dgvinput_employee.AllowUserToAddRows = False
            'Me.dgvinput_machine.AllowUserToAddRows = False
            'If dgvinput_machine.RowCount <= 0 And dgvinput_employee.RowCount <= 0 Then
            '    MsgBox("Please enter 1 item data in employe or machine ", MsgBoxStyle.Information, "Information")
            '    Me.dgvinput_employee.AllowUserToAddRows = True
            '    Me.dgvinput_machine.AllowUserToAddRows = True
            '    Me.tc_processoutput.SelectedIndex = 1
            '    Return False
            'End If
            'Me.dgvinput_employee.AllowUserToAddRows = True
            'Me.dgvinput_machine.AllowUserToAddRows = True

            '->cek pada setiap dgv kalo qty tidak boleh <= 0
            '1. dgv process reject
            For Each vr As DataGridViewRow In Me.dgvinput_processreject.Rows
                If String.IsNullOrEmpty(vr.Cells(Me.colinput_processreject_rejectcode.Index).Value) = False Then
                    If CDec(vr.Cells(Me.colinput_processreject_qty.Index).Value) <= 0 Then
                        MsgBox("Please enter a positive number for quantity data, process reject with reject code:" & _
                        vr.Cells(Me.colinput_processreject_rejectcode.Index).Value.ToString, MsgBoxStyle.Information, _
                                     "Information")
                        Return False
                    End If
                End If
            Next
            '2 dgv rm reject
            For Each vr As DataGridViewRow In Me.dgvinput_rmreject.Rows
                If String.IsNullOrEmpty(vr.Cells(Me.colinput_rmreject_codeno.Index).Value) = False Then
                    If String.IsNullOrEmpty(vr.Cells(Me.colinput_rmreject_rejectcode.Index).Value) = False Then
                        If CDec(vr.Cells(Me.colinput_rmreject_qty.Index).Value) <= 0 Then
                            MsgBox("Please enter a positive number for quantity data, RM reject with reject code:" & _
                                                   vr.Cells(Me.colinput_rmreject_rejectcode.Index).Value.ToString, MsgBoxStyle.Information, _
                                                                "Information")
                            Return False
                        End If
                    Else
                        MsgBox("Reject code cannot be an empty string.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            Next
            '3 dgv downtime
            For Each vr As DataGridViewRow In Me.dgvinput_downtime.Rows
                If String.IsNullOrEmpty(vr.Cells(Me.colinput_downtime_code.Index).Value) = False Then
                    If Me.ObjCD.Convert_ToMinutes(vr.Cells(Me.colinput_downtime_duration.Index).Value) <= 0 Then
                        MsgBox("Minimum down time duration is 00.01", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            Next
            '4 dgv setuptime
            For Each vr As DataGridViewRow In Me.dgvInput_SetupTime.Rows
                If String.IsNullOrEmpty(vr.Cells(Me.ColST_Code.Index).Value) = False Then
                    If Me.ObjCD.Convert_ToMinutes(vr.Cells(Me.ColST_Duration.Index).Value) <= 0 Then
                        MsgBox("Minimum setup time duration is 00.01", MsgBoxStyle.Information, _
                                                                                    "Information")
                        Return False
                    End If
                End If
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function ApakahDataItemAfterBenar() As Boolean
        Try
            '->wo tidak boleh kosong
            If Me.txtwonumber.Text = "" Then
                MsgBox("WO number cannot be an empty string.", MsgBoxStyle.Information, "Information")
                Me.txtwonumber.Focus()
                Return False
            End If

            'data item salah bila:
            '1. product code yang diisikan salah
            If Me.txtDescAfter.Text = "" Then
                MsgBox("Please enter code no.", MsgBoxStyle.Information, "Information")
                Me.TxtCodeNoAfter.Text = ""
                Me.TxtCodeNoAfter.Focus()
                Return False
            End If
            '2. nilai qty <= 0
            If Me.ObjCD.CekBilanganPositif(Me.TxtQtyperBatchAfter.Text) = False Then
                MsgBox("Please enter a positive number for quantity data.", MsgBoxStyle.Information, "Information")
                Me.TxtQtyperBatchAfter.Focus()
                Return False
            End If
            '3. Bila batch no controlled, batch no tidak diisikan
            If Me.chkBatchNoControlledAfter.Checked = True Then
                If Me.txtBatchNoAfter.Text = "" Then
                    MsgBox("Please enter batch no, this product is batch no controlled.", MsgBoxStyle.Information, "Information")
                    Me.txtBatchNoAfter.Focus()
                    Return False
                End If
            End If
            '4.a Untuk transaksi proses (in/back/out proses) no Pdo harus diisikan
            '4.b Untuk transaksi yang berkaitan dengan purchase (stock out purchase return dan stout in purchase)
            '    harus menyertakan no PO bila terhubungkan dengan software purchase
            '4.c Untuk transaksi yang berkaitan dengan sales (stock out sales dan stock in sales return) harrus
            '    menyertakan no DO bila terhubungkan dengan software sales
            '5. Bila location controlled, kode lokasi tidak diisikan 
            If Me.chkLocationControlledAfter.Checked = True Then
                If Me.txtPrimaryLocationAfter.Text = "-1" Then
                    MsgBox("Please enter location, this product is location controlled.", MsgBoxStyle.Information, "Information")
                    Me.TxtLocationAfter.Focus()
                    Return False
                End If
            End If
            '7. Bila use secondary unit, nilai qty secondary <=0
            If Me.chkUseSecondaryUnitAfter.Checked = True Then
                If Me.ObjCD.CekBilanganPositif(Me.txtSecondaryQtyAfter.Text) = False Then
                    MsgBox("Please enter a positive number for secondary quantity data.", MsgBoxStyle.Information, "Information")
                    Me.txtSecondaryQtyAfter.Focus()
                    Return False
                End If
            End If
            '8a. Bila product yang bersangkutan batch no controlled dan expiry date controlled dan expiry date
            '   yang dimasukkan pada tidak sama dengan expiry date yang ada didalam basis data maka data salah
            If Me.chkBatchNoControlledAfter.Checked = True And Me.chkExpiryDateControlledAfter.Checked = True Then
                Dim ExpiryDateLama As Date = Me.ObjProduct.GetExpiryDate(Me.TxtPrimaryCodeNoAfter.Text, Me.txtBatchNoAfter.Text)
                If ExpiryDateLama <> #1/1/1950# Then
                    If ExpiryDateLama.Date <> Me.DtpExpiredAfter.Value.Date Then
                        MsgBox("Expiry date for batch no: " + Me.txtBatchNoAfter.Text + " is " + _
                        Format(ExpiryDateLama, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                        Me.DtpExpiredAfter.Focus()
                        Return False
                    End If
                End If
                '8b. Bila product yang bersangkutan batch no controlled dan expiry date controlled dan expiry date yang dimasukkan
                '    tidak sama dengan expiry date yang ada sudah ada didalam dgv
                Dim row As DataGridViewRow
                For Each row In Me.dgvItemAfter.Rows
                    If row.Cells(7).Value = Me.txtBatchNoAfter.Text And Me.txtBatchNoAfter.Text <> "" And row.Cells(12).Value = _
                    Me.TxtPrimaryCodeNoAfter.Text Then
                        If CDate(row.Cells(8).Value) <> Me.DtpExpiredAfter.Value.Date Then
                            MsgBox("You already entered expiry date for batch no: " + Me.txtBatchNoAfter.Text + " on " + _
                            Format(CDate(row.Cells(8).Value), "dd MMM yyyy") + ". Please change your expiry date data.", _
                            MsgBoxStyle.Information, "Information")
                            Me.DtpExpiredAfter.Focus()
                            Return False
                        End If
                    End If
                Next
                ''8.c Bila product yang bersangkutan batch no controlled dan expiry date controlled dan expiry date yang dimasukkan
                '    tidak sama dengan expiry date yang ada sudah ada didalam transaksi lainnya: stock in adjust, stock in purchase, 
                '    stock in sales return, stock tag, transaksi bundling (item before) 
                '    (untuk stock out tidak dicek karena semua harus masuk dulu sebelum keluar, jadi dimasukka ke dalam pengecekan
                '    adalah semua transaksi yang menjadi pintu masuk inventory)
                '1. stock in (semua stock in)
                Dim JenisTransaksi As clsTransaksiStockInOut.EnumTipeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase
                For loopInOut As Integer = 0 To 4 Step 1
                    Select Case loopInOut
                        Case 0
                            '   a. stock in purchase
                            JenisTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase
                            Exit Select
                        Case 1
                            '   b. stock in sales return
                            JenisTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn
                            Exit Select
                        Case 2
                            '   c. stock in adjust
                            JenisTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust
                            Exit Select
                        Case 3
                            '   d. stock in back
                            JenisTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_BackProcess
                            Exit Select
                        Case 4
                            '   e. stock in proses
                            JenisTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Process
                            Exit Select
                    End Select
                    ExpiryDateLama = Me.ObjItemInOut.GetExpiryDate(Me.TxtPrimaryCodeNoAfter.Text, Me.txtBatchNoAfter.Text, _
                    JenisTransaksi)
                    If ExpiryDateLama <> #1/1/1950# Then
                        If ExpiryDateLama.Date <> Me.DtpExpiredAfter.Value Then
                            MsgBox("Expiry date for batch no: " + CStr(Me.txtBatchNoAfter.Text) + " is " + _
                            Format(ExpiryDateLama, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                            Me.DtpExpiredAfter.Focus()
                            Return False
                        End If
                    End If
                Next
                '2. transaksi stock bundling untuk item hasil bundling
                ExpiryDateLama = Me.ObjItemBundling.GetExpiryDateAfter(Me.TxtPrimaryCodeNoAfter.Text, Me.txtBatchNoAfter.Text)
                If ExpiryDateLama <> #1/1/1950# Then
                    If ExpiryDateLama.Date <> Me.DtpExpiredAfter.Value Then
                        MsgBox("Expiry date for batch no: " + CStr(Me.txtBatchNoAfter.Text) + " is " + _
                        Format(ExpiryDateLama, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                        Me.DtpExpiredAfter.Focus()
                        Return False
                    End If
                End If
                '3. proses stock taking
                'msgbox("belum")
            End If

            '6. Bila expiry date controlled, expiry date back date dibandingkan dengan tanggal transaksi
            '   hanya diberikan peringatan, karena mungkin saja mau jual barang expired dengan harga 
            '   yang sangat murah -  atau adjust untuk dikeluarkan dari data komputer
            If Me.chkExpiryDateControlledAfter.Checked = True Then
                If Me.DtpExpiredAfter.Value.Date < Me.dtpDate.Value.Date Then
                    If MsgBox("This product expiry date is smaller than transaction date. Do you want to continue?", _
                    MsgBoxStyle.YesNo, "Expired Product") = MsgBoxResult.No Then
                        Me.DtpExpiredAfter.Focus()
                        Return False
                    End If
                End If
            End If

            '7. Jika item before, sudah ada item 
            Dim rowDGV As DataGridViewRow
            For Each rowDGV In Me.dgvItemb4.Rows
                If rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value = Me.TxtPrimaryCodeNoAfter.Text Then
                    MsgBox("There product with code no: " + Me.TxtCodeNoAfter.Text + _
                    " in item before, we cannot enter this item", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Next

            'cost factor harus positif
            'If CDec(Me.txtFaktorPemberat.Text) <= 0 Then
            '    MsgBox("Cost factor must be a positif value", MsgBoxStyle.Information, "Information")
            '    Me.txtFaktorPemberat.Focus()
            '    Return False
            'End If

            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlDGV(ByVal vPrimaryCodeNo As Double, ByVal Secondary_Qty As _
    Decimal, ByVal vPrimaryLocation As Double, ByVal vBatchNo As String, _
    ByVal BatchNoControlled As Boolean, ByVal LocationControlled As Boolean, ByVal UseSecondaryUnit _
    As Boolean, ByVal vDGV As DataGridView) As Decimal
        Try
            Dim row As DataGridViewRow
            Dim Jml As Decimal = 0
            Dim OK As Boolean = False

            For Each row In vDGV.Rows
                If row.Cells(12).Value = vPrimaryCodeNo Then
                    OK = True
                    'cek secondary qty
                    If UseSecondaryUnit = True And OK = True Then
                        If row.Cells(5).Value <> Secondary_Qty Then
                            OK = False
                        End If
                    End If
                    'cek location
                    If LocationControlled = True And OK = True Then
                        If vPrimaryLocation <> row.Cells(13).Value Then
                            OK = False
                        End If
                    End If
                    'cek batch no
                    If BatchNoControlled = True Then
                        If vBatchNo <> row.Cells(7).Value Then
                            OK = False
                        End If
                    End If
                    'expiry date controlled-> untuk expired data diasumsikan user tidak bisa input data expiry
                    'date yang salah karena sudah dicek terlebih dahulu
                    If OK = True Then
                        Jml = Jml + CDbl(row.Cells(3).Value)
                    End If
                End If
            Next

            Return Jml
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function ApakahQtyRMUsageSudahBenar(ByVal vprimary_bundling As Double, _
    ByVal vprimary_product As Double, ByVal videntifier_wo As String, ByVal qtyrm As Decimal, _
    ByVal diinput_didgv As Boolean, ByVal vIdentifierJO As String) As Boolean
        Dim jml_dgv, jml_lama, jml_blc, jml_setelahedit, jml_dgvlain As Decimal
        Try

            If Me.ObjStockBundling.ApakahCodeNoDiRMSudahAda(videntifier_wo, vprimary_product, vIdentifierJO) = False Then
                MsgBox("Raw material product code is not member of output product.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            jml_blc = Me.ObjStockBundling.get_qtyrmbalance(videntifier_wo, vprimary_product, vIdentifierJO) + Me.ObjStockBundling.get_qty_wip_allocated_balance(videntifier_wo, vprimary_product, vIdentifierJO)
            jml_lama = Me.ObjStockBundling.get_qtystockbundlinglama(vprimary_bundling, vprimary_product)
            jml_dgv = Me.get_jmlproduct_indgvb4(vprimary_product)

            If diinput_didgv = False Then
                jml_setelahedit = jml_blc + jml_lama - jml_dgv - CDbl(Me.TxtQtyperBatchb4.Text)
                If jml_setelahedit < 0 Then
                    jml_dgvlain = jml_dgv - qtyrm
                    MsgBox("Maximum value you can enter for this raw material is: " & Me.ObjCD.DigitGrouping(jml_blc + jml_lama - jml_dgv, _
                    Me.ObjProduct.GetAngkaDecimal(vprimary_product)) & " " & Me.ObjProduct.GetUnitCode(vprimary_product), MsgBoxStyle.Information, "Information")
                    Me.TxtQtyperBatchb4.Focus()
                    Return False
                End If
            Else
                jml_setelahedit = jml_blc + jml_lama - jml_dgv
                If jml_setelahedit < 0 Then
                    jml_dgvlain = jml_dgv - qtyrm
                    MsgBox("Maximum value you can enter for this raw material is: " & Me.ObjCD.DigitGrouping(jml_blc + jml_lama - jml_dgvlain, _
                    Me.ObjProduct.GetAngkaDecimal(vprimary_product)) & " " & Me.ObjProduct.GetUnitCode(vprimary_product), MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function get_jmlproduct_indgvb4(ByVal vprimary_product As Double) As Decimal
        Dim retval As Decimal = 0
        Try
            For Each vr As DataGridViewRow In Me.dgvItemb4.Rows
                If vr.Cells(Me.ColPrimaryCodeNo.Index).Value = vprimary_product Then
                    retval += CDec(vr.Cells(Me.ColQty.Index).Value)
                End If
            Next
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim TanggalClosingMax As Date = #1/1/1950#
        Dim Jml As Decimal = 0
        Dim rowDGV As DataGridViewRow

        Try
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Me.dtpDate.Value Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            'cek data yang harus diisi
            KoleksiNamaData.Add("Stock bundling number")
            KoleksiTextBox.Add(Me.txtstockbundling)
            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'minimal harus punya 1 data item after dan 1 item before
            If Me.dgvItemAfter.RowCount = 0 Then
                MsgBox("Please enter at least 1 item as bundling result.", MsgBoxStyle.Information, "Information")
                Me.tc_processoutput.SelectedIndex = 0
                Return False
            End If
            If Me.dgvItemb4.RowCount = 0 Then
                MsgBox("Please enter at least 1 item as bundling raw material.", MsgBoxStyle.Information, "Information")
                Me.tc_processoutput.SelectedIndex = 2
                Return False
            End If

            'cek berkaitan dengan tanggal transaksi yang sudah dikunci
            'cek timestamp karena bisa saja sudah ada orang yang menyimpan perubahan data ketika user sedang
            'melakukan proses edit
            'msgbox("belum")

            'nilai cost factor jumlah totalnya harus sama dengan 100%

            If Me.dgvItemAfter.RowCount = 1 Then
                rowDGV = Me.dgvItemAfter.Rows(0)
                rowDGV.Cells(Me.ColCostFactor.Index).Value = 100.0
            ElseIf Me.dgvItemAfter.RowCount > 0 Then
                If MsgBox("Calculate cost factor automatically?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    If Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Stock_Bundling, _
                    ClsTransaksiSetting.EnumNamaParameter.Use_Custom_CostFaktorCalc) = "1" Then
                        'hitung menggunakan data custom
                        If Me.ObjCustom.HitungCostFactor(Me.dgvItemAfter, Me.ColQtyAfter.Index, Me.ColPrimaryCodeNoAfter.Index, _
                        Me.ColCostFactor.Index, Me.ColCodeNoAfter.Index, Me.ColAfter_PrimaryUnit.Index) = False Then
                            MsgBox("Logic error in cost factor calculation, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    Else
                        'hitung berdasarkan qty
                        Me.HitungCostFactor()
                    End If
                End If
            End If

            For Each rowDGV1 As DataGridViewRow In Me.dgvItemAfter.Rows
                Jml += CDec(rowDGV1.Cells(Me.ColCostFactor.Index).Value)
            Next
            If Jml <> 100 Then
                MsgBox("Sum of cost factor data must be 100%", MsgBoxStyle.Information, "Information")
                Me.tc_processoutput.SelectedIndex = 0
                Return False
            End If

            'berkaitan denga tanggal closing max
            TanggalClosingMax = Me.ObjCD.GetLasDateOfTheMonth(ObjCL.GetLastClosingPeriod)
            If Me.dtpDate.Value.Date <= TanggalClosingMax.Date Then
                MsgBox("The period " + Format(Me.dtpDate.Value, "MMM yyyy") + " has been closed, you cannot save this data.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            'cek berkaitan dengan perubahan tanggal
            If Me.baru = False Then
                If Me.ObjStockBundling.TanggalTransaksi.Month = Me.dtpDate.Value.Month And _
                Me.ObjStockBundling.TanggalTransaksi.Year = Me.dtpDate.Value.Year Then
                    Me.txtstockbundling.Text = Me.ObjStockBundling.NoTransaksi
                Else
                    If MsgBox("You just changed transaction date, this will change your transaction number and there's a possibility your transaction no cannot be change back. Continue and accept this change?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
                        Me.dtpDate.Value = Me.ObjStockBundling.TanggalTransaksi
                        Me.txtstockbundling.Text = Me.ObjStockBundling.NoTransaksi
                        Return False
                    End If
                End If
            End If

            'cek berkaitan perhitungan tanpa batas pada item bundling after dan before, dipindahkan ke kelas, karena
            'ada fasilitas antrian untuk transaksi di kelas.

            If Me.ApakahItemProcessSudahBenar(Me.dtpactual_start.Value, Me.dtpactual_end.Value, Me.txtprimarywu.Text) = False Then
                Me.tc_processoutput.SelectedTab = Me.tabpage_rejectdowntimeoperdata
                Return False
            End If

            If Me.ObjStockBundling.ApakahQtyRMUsageSudahBenar(Me, Me.TxtPrimaryStockBundling.Text, Me.txtidentifier_wo.Text, Me.TxtIdentifierJO.Text) = False Then
                Me.tc_processoutput.SelectedTab = Me.tabpage_rmusage
                Return False
            End If

            If Me.IsWorkShiftRequired Then
                If String.IsNullOrEmpty(Me.TxtWorkShiftCode.Text.Trim) Or Me.TxtPrimaryWorkShift.Text = "-1" Then
                    MsgBox("Work shift cannot be an empty string.", MsgBoxStyle.Information, "Information")
                    Me.TxtWorkShiftCode.Focus()
                    Return False
                End If
            End If

            Dim objLocationAccess As New Core_Module.FrmUserLocationPrivilegeSetting
            Dim listOfPrimaryLocation As New List(Of Double)
            For Each dataGridViewRow As DataGridViewRow In Me.dgvItemAfter.Rows
                If dataGridViewRow.Cells(Me.ColPrimaryLocationAfter.Index).Value <> "-1" Then
                    If listOfPrimaryLocation.Contains(dataGridViewRow.Cells(Me.ColPrimaryLocationAfter.Index).Value) = False Then listOfPrimaryLocation.Add(dataGridViewRow.Cells(Me.ColPrimaryLocationAfter.Index).Value)
                End If
            Next
            If objLocationAccess.HasAccessPrivileges(Me.UserCode, listOfPrimaryLocation, FrmUserLocationPrivilegeSetting.EnumTipeTransaksiInOut.PRODUCTION_LOG_AFTER) = False Then
                Return False
            End If
            listOfPrimaryLocation.Clear()
            For Each dataGridViewRow As DataGridViewRow In Me.dgvItemb4.Rows
                If dataGridViewRow.Cells(Me.ColPrimaryLocation.Index).Value <> "-1" Then
                    If listOfPrimaryLocation.Contains(dataGridViewRow.Cells(Me.ColPrimaryLocation.Index).Value) = False Then listOfPrimaryLocation.Add(dataGridViewRow.Cells(Me.ColPrimaryLocation.Index).Value)
                End If
            Next
            If objLocationAccess.HasAccessPrivileges(Me.UserCode, listOfPrimaryLocation, FrmUserLocationPrivilegeSetting.EnumTipeTransaksiInOut.PRODUCTION_LOG_BEFORE) = False Then
                Return False
            End If

            'cost factor tidak boleh nol
            For Each row1 As DataGridViewRow In Me.dgvItemAfter.Rows
                'row1.Cells(Me.ColCostFactor.Index).Value = 0
                If row1.Cells(Me.ColCostFactor.Index).Value <= 0 Then
                    MsgBox("Invalid cost factor on row " + (row1.Index + 1).ToString, MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Next

            'kalo sampe sini return true

            If Me.dtpactual_start.Value > Me.dtpactual_end.Value Then
                MsgBox("Actual start time must be less than actual end time.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            KoleksiNamaData = Nothing
            KoleksiTextBox = Nothing
        End Try
    End Function

    Private Sub HitungCostFactor()
        Dim rowDGV As DataGridViewRow
        Dim TotalJml As Decimal = 0
        Dim Dummy As Decimal = 0

        Try
            For Each rowDGV In Me.dgvItemAfter.Rows
                TotalJml += CDec(rowDGV.Cells(Me.ColQtyAfter.Index).Value)
            Next
            For Each rowDGV In Me.dgvItemAfter.Rows
                If rowDGV.Index = Me.dgvItemAfter.Rows.Count - 1 Then
                    'untuk baris terakhir = 100 - sub total
                    rowDGV.Cells(Me.ColCostFactor.Index).Value = Me.ObjCD.DigitGrouping(100 - Dummy, 2)
                Else
                    rowDGV.Cells(Me.ColCostFactor.Index).Value = ObjCD.DigitGrouping((CDec(rowDGV.Cells(Me.ColQtyAfter.Index).Value) * 100) / TotalJml, 2)
                    Dummy += CDec(rowDGV.Cells(Me.ColCostFactor.Index).Value)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HilangkanItem(ByVal TBLDummy As DseTrans_StockBundling.DataTable3DataTable)
        Dim Filter As String = ""
        Filter = "[Primary Transaksi] = " + Me.TxtPrimaryStockBundling.Text

        Try
            TBLDummy.DefaultView.RowFilter = Filter
            While TBLDummy.DefaultView.Count > 0
                TBLDummy.Rows.Remove(TBLDummy.DefaultView(0).Row)
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TambahkanItem(ByVal TBLDummy As DseTrans_StockBundling.DataTable3DataTable)
        Dim rowDGV_B4 As DataGridViewRow
        Dim rowDGV_After As DataGridViewRow
        Dim rowBundling As DseTrans_StockBundling.DataTable3Row

        Try
            For Each rowDGV_B4 In Me.dgvItemb4.Rows
                For Each rowDGV_After In Me.dgvItemAfter.Rows
                    rowBundling = TBLDummy.NewDataTable3Row
                    rowBundling.Primary_Product_After = rowDGV_After.Cells(Me.ColPrimaryCodeNoAfter.Index).Value
                    rowBundling.Primary_Product_Before = rowDGV_B4.Cells(Me.ColPrimaryCodeNo.Index).Value
                    rowBundling.Primary_Transaksi = Me.TxtPrimaryStockBundling.Text
                    rowBundling.Tanggal_Bundling = Me.dtpDate.Value.Date
                    TBLDummy.Rows.Add(rowBundling)
                Next
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeObjek()
        Try
            With Me.ObjStockBundling
                .TanggalTransaksi = Me.dtpDate.Value
                .NoTransaksi = Me.txtstockbundling.Text
                .Remarks = Me.txtremarks.Text
                .IdentifierWO = Me.txtidentifier_wo.Text
                .IdentifierJO = Me.TxtIdentifierJO.Text
                .TipeBundling = Me.tipe_bundling
                .code_process = Me.txtprocess_code.Text
                .desc_process = Me.txtprocess_desc.Text
                .Mono = Me.txtmo_number.Text
                .Jono = Me.txtjo_number.Text
                .code_wu = Me.txtworkunitDesc.Text
                .WoNo = Me.txtwonumber.Text
                .PrimaryProses = Me.txtprimaryprocess.Text
                .PrimaryWorkUnit = Me.txtprimarywu.Text
                .PrimaryWorkshift = Me.TxtPrimaryWorkShift.Text
                '.PrimaryWorkshift = Me.txtprimaryws.Text
                .GoodPercentage = Me.txtgood_perct.Text
                .TotalReject = Me.txttotal_reject.Text
                .TotalGood = Me.txttotal_good.Text
                .TotalProduce = Me.txttotal_produce.Text
                '-> actual
                .ProdRateActual = Me.txtactual_prodrate.Text
                .DownTimeActual = ObjCD.Convert_ToMinutes(Me.txtactual_downtime.Text)
                .WorkHourActual = ObjCD.Convert_ToMinutes(Me.txtactual_workhour.Text)
                .SetupTimeActual = ObjCD.Convert_ToMinutes(Me.txtactual_setuptime.Text)
                .EndTimeActual = Me.dtpactual_end.Value
                .StartTimeActual = Me.dtpactual_start.Value
                '-> planning
                .StartTimePlanned = Me.dtpplan_start.Value
                .EndTimePlanned = Me.dtpplan_end.Value
                .ProdRatePlanned = Me.txtplan_prodrate.Text
                .WorkHourPlanned = ObjCD.Convert_ToMinutes(Me.txtplan_workhour.Text)
                .SetupTimePlanned = ObjCD.Convert_ToMinutes(Me.txtplan_setuptime.Text)
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SalinKeField()
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Try
            With Me.ObjStockBundling
                Me.dtpDate.Value = .TanggalTransaksi
                Me.txtstockbundling.Text = .NoTransaksi
                Me.txtremarks.Text = .Remarks
                Me.TxtPrimaryStockBundling.Text = .Primary
                Me.txtidentifier_wo.Text = .IdentifierWO
                Me.TxtIdentifierJO.Text = .IdentifierJO
                Me.tipe_bundling = .TipeBundling
                Me.txtprimaryproduct_dariwo.Text = .GetPrimaryCodeByWo(Me.txtidentifier_wo.Text, Me.TxtIdentifierJO.Text)
                Me.txtprimaryprocess.Text = .PrimaryProses
                Me.txtprocess_code.Text = .code_process
                Me.txtprocess_desc.Text = .desc_process
                Me.PrimaryWorkUnit_ReffOfBundling = .PrimaryWorkUnit
                Me.txtprimarywu.Text = .PrimaryWorkUnit
                Me.txtworkunitDesc.Text = .code_wu
                Me.txtprimaryws.Text = .PrimaryWorkshift
                Me.txtmo_number.Text = .Mono
                Me.txtjo_number.Text = .Jono
                Me.txtwonumber.Text = .WoNo
                Me.txtgood_perct.Text = ObjCD.DigitGrouping(.GoodPercentage, 2)
                Me.txttotal_reject.Text = ObjCD.DigitGrouping(.TotalReject, Me.isDecimal)
                Me.txttotal_good.Text = ObjCD.DigitGrouping(.TotalGood, Me.isDecimal)
                Me.txttotal_produce.Text = ObjCD.DigitGrouping(.TotalProduce, Me.isDecimal)
                Me.TxtPrimaryWorkShift.Text = .PrimaryWorkshift
                Me.TxtWorkShiftCode.Text = Me.ObjWorkShift.GetCodeByPrimary(.PrimaryWorkshift)
                Me.TxtWorkShiftDesc.Text = Me.ObjWorkShift.GetDescription(.PrimaryWorkshift)
                '-> actual
                Me.txtactual_prodrate.Text = ObjCD.DigitGrouping(.ProdRateActual, Me.isDecimal, ClsCekData.EnumTipeRounding.Down)
                Me.txtactual_downtime.Text = ObjCD.Convert_ToString(.DownTimeActual)
                Me.txtactual_workhour.Text = ObjCD.Convert_ToString(.WorkHourActual)
                Me.txtactual_setuptime.Text = ObjCD.Convert_ToString(.SetupTimeActual)
                Me.dtpactual_end.Value = .EndTimeActual
                Me.dtpactual_start.Value = .StartTimeActual
                '-> planning
                Me.dtpplan_start.Value = .StartTimePlanned
                Me.dtpplan_end.Value = .EndTimePlanned
                Me.txtplan_prodrate.Text = ObjCD.DigitGrouping(.ProdRatePlanned, Me.isDecimal, ClsCekData.EnumTipeRounding.Down)
                Me.txtplan_workhour.Text = ObjCD.Convert_ToString(.WorkHourPlanned)
                Me.txtplan_setuptime.Text = ObjCD.Convert_ToString(.SetupTimePlanned)
                Me.GetTotalActualTime()
                '-> dummy
                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With

            Me.FormatDGV(Me.dgvItemb4)
            Me.FormatDGV(Me.dgvItemAfter)
            'digit grouping
            Dim row As DataGridViewRow
            For Each row In Me.dgvItemb4.Rows
                row.Cells(3).Value = Me.ObjCD.DigitGrouping(row.Cells(3).Value)
                If IsNumeric(row.Cells(5).Value) = True Then
                    row.Cells(5).Value = Me.ObjCD.DigitGrouping(row.Cells(5).Value)
                End If
            Next
            For Each row In Me.dgvItemAfter.Rows
                row.Cells(3).Value = Me.ObjCD.DigitGrouping(row.Cells(3).Value)
                If IsNumeric(row.Cells(5).Value) = True Then
                    row.Cells(5).Value = Me.ObjCD.DigitGrouping(row.Cells(5).Value)
                End If
            Next

            Me.AturWoNumberdanCboWorkUnit()
            Me.IsikanCboWorkUnit(Me.txtidentifier_wo.Text, Me.TxtIdentifierJO.Text)
            Windows.Forms.Cursor.Current = Cursors.Default
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.BufferWorkHour = Me.txtactual_workhour.Text
        End Try
    End Sub

    Private Sub SetTampilan()
        Try
            Me.SplitContainer3.Left = Me.SplitContainer3.Left / 2
            Me.SplitContainer3.Width = Me.SplitContainer1.Panel1.Width - Me.SplitContainer3.Left
            Me.txttotal_produce.Width = Me.SplitContainer3.Panel1.Width - Me.txttotal_produce.Left
            Me.txtgood_perct.Width = Me.txttotal_produce.Width
            Me.txttotal_good.Width = Me.SplitContainer3.Panel2.Width - Me.txttotal_good.Left
            Me.txttotal_reject.Width = Me.txttotal_good.Width

            Me.SplitContainer2.Width = Me.SplitContainer1.Panel1.Width - Me.SplitContainer2.Left
            Me.ColCostFactor.Visible = True

            'after
            Me.DtpExpiredAfter.Width = Me.SplitContainer5.Panel1.Width - Me.DtpExpiredAfter.Left - 5
            Me.txtBatchNoAfter.Width = Me.DtpExpiredAfter.Width
            Me.txtDescAfter.Width = Me.DtpExpiredAfter.Width
            Me.BtnFindCodeNoAfter.Left = Me.SplitContainer5.Panel1.Width - Me.BtnFindCodeNoAfter.Width - 5
            Me.TxtCodeNoAfter.Width = Me.BtnFindCodeNoAfter.Left - Me.TxtCodeNoAfter.Left - 5
            Me.BtnInputAfter.Left = Me.SplitContainer5.Panel2.Width - 5 - Me.BtnInputAfter.Width
            Me.lblSecondaryUnitsAfter.Left = Me.SplitContainer5.Panel2.Width - 5 - Me.lblSecondaryUnitsAfter.Width
            'Me.txtSecondaryQtyAfter.Visible = True
            'Me.lblSecondaryUnitsAfter.Visible = True
            Me.txtSecondaryQtyAfter.Left = Me.lblSecondaryUnitsAfter.Left - 5 - Me.txtSecondaryQtyAfter.Width
            Me.lblUnitsAfter.Left = Me.txtSecondaryQtyAfter.Left - 5 - Me.lblUnitsAfter.Width
            Me.TxtQtyperBatchAfter.Width = Me.lblUnitsAfter.Left - 5 - Me.TxtQtyperBatchAfter.Left
            Me.TxtRemarkItemAfter.Width = Me.BtnInputAfter.Left - 5 - Me.TxtRemarkItemAfter.Left
            Me.TxtWarehouseAfter.Width = Me.SplitContainer5.Panel2.Width - 5 - Me.TxtWarehouseAfter.Left
            Me.btnFindLocationAfter.Left = Me.SplitContainer5.Panel2.Width - 5 - Me.btnFindLocationAfter.Width
            Me.TxtLocationAfter.Width = Me.btnFindLocationAfter.Left - 5 - Me.TxtLocationAfter.Left

            'Me.lblUnitsAfter.Left = Me.SplitContainer5.Panel2.Width - 5 - Me.lblUnitsAfter.Width
            'Me.TxtQtyperBatchAfter.Width = Me.lblUnitsAfter.Left - 5 - Me.TxtQtyperBatchAfter.Left

            Me.btnNext.Visible = False
            Me.btnNext.TabStop = False
            Me.btnPrev.Visible = False
            Me.btnPrev.TabStop = False

            'Me.btnNextTab1.Top = Me.btnCancel.Top
            'Me.btnNextTab2.Top = Me.btnCancel.Top
            'Me.btnPrevTab2.Top = Me.btnCancel.Top
            'Me.btnPrevTab3.Top = Me.btnCancel.Top
            'Me.btnNextTab1.Visible = False
            'Me.btnNextTab1.TabStop = False
            'Me.btnNextTab2.Visible = False
            'Me.btnNextTab2.TabStop = False
            'Me.btnPrevTab2.Visible = False
            'Me.btnPrevTab2.TabStop = False
            'Me.btnPrevTab3.Visible = False
            'Me.btnPrevTab3.TabStop = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub LoadReadOnly()
        Try
            Me.Cursor = Cursors.WaitCursor
            If Me.ObjStockBundling.CariDetail(Me.PrimaryBundling, Me.dgvItemb4, Me.dgvItemAfter, Me, Me.tbl_workunit) = True Then
                Me.dgvInput_SetupTime.AllowUserToAddRows = False
                Me.dgvinput_downtime.AllowUserToAddRows = False
                Me.dgvinput_processreject.AllowUserToAddRows = False

                Me.ForceNavigate = True
                Me.tc_processoutput.SelectedTab = Me.tabpage_processoutput
                Me.SalinKeField()
                Me.btnNew.Enabled = False
                Me.btnEdit.Enabled = False
                Me.btnDelete.Enabled = False
                Me.btnFind.Enabled = False
                Me.btnExit.Enabled = True
                Me.btnOK.Enabled = False
                Me.btnCancel.Enabled = False
                Me.btnCancelTP1.Enabled = False
                Me.btnCancelTP2.Enabled = False
                Me.btnUpAfter.Enabled = False
                Me.btnUPb4.Enabled = False
                Me.btnDownAfter.Enabled = False
                Me.btnDownb4.Enabled = False
                Me.llb_BOM.Enabled = False
                Me.btnEdit.Focus()
            Else
                Me.btnFind.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub frmMasterTipe_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            If Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Stock_Bundling, ClsTransaksiSetting.EnumNamaParameter.ProductionLogWorkShiftRequired) = "1" Then
                IsWorkShiftRequired = True
            Else
                IsWorkShiftRequired = False
            End If
            Me.tabpage_processoutput.BackColor = Me.SplitContainer1.Panel1.BackColor
            Me.tabpage_rejectdowntimeoperdata.BackColor = Me.SplitContainer1.Panel1.BackColor
            Me.tabpage_rmusage.BackColor = Me.SplitContainer1.Panel1.BackColor
            Me.TabControl3.TabPages.Remove(TabPage4)
            Me.TabControl3.TabPages.Remove(TabPage6)
            Me.TabControl3.TabPages.Remove(TabPage8)
            Me.Text = "Production Log"
            Me.Label4.Text = "Document No"
            Me.tabpage_rejectdowntimeoperdata.Text = "Reject && Downtime data"

            Me.baru = True
            Me.NonAktifkanDetail()
            Me.KosongkanField()
            Me.ObjStockBundling.TipeBundling = ClsTrans_StockBundling.EnumTipeBundling.Inventory

            If Me.ObjStockBundling.TipeBundling = ClsTrans_StockBundling.EnumTipeBundling.Produksi1 Then
                Me.gboItemAfter.Height = txtFaktorPemberat.Top + txtFaktorPemberat.Height + 5
                Me.TxtRemarkItemAfter.Width = TxtWarehouseAfter.Width
                Me.BtnInputAfter.Top = txtFaktorPemberat.Top
                Me.BtnInputAfter.Height = Me.txtFaktorPemberat.Height
                Me.lblFaktorPemberat.Visible = True
                Me.lblFaktorPemberat.Text = "Cost Factor (%)"
                Me.txtFaktorPemberat.Visible = True
                Me.dgvItemAfter.Top = Me.gboItemAfter.Top + Me.gboItemAfter.Height + 5
                Me.dgvItemAfter.Height = Me.tabpage_processoutput.Height - Me.dgvItemAfter.Top - (Me.tabpage_processoutput.Height - Me.btnUpAfter.Top) - 5
            Else
                Me.dgvItemAfter.Columns(16).Visible = False
                Me.txtFaktorPemberat.TabStop = False
            End If

            Me.dgvItemAfter.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvItemb4.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            Me.dgvinput_processreject.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvinput_downtime.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvinput_machine.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvinput_employee.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvinput_rmreject.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvInput_SetupTime.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            If Me.PrimaryBundling <> -1 Then
                Me.LoadReadOnly()
            End If

            If Me._Command = EnumCommand.NEW_ And Me._IdentifierWO <> "" And Me._PrimaryWorkUnit <> -1 Then
                Me.CreateLogDirect1()
                Me.Timer1.Enabled = True
            End If
            If Me._Command = EnumCommand.EDIT And Me._PrimaryBundling <> -1 Then
                Me.EditLogDirect1()
            End If
            If Me._Command = EnumCommand.DELETE And Me._PrimaryBundling <> -1 Then
                Me.DeleteLogDirect1()
            End If

            If Me._Command = EnumCommand.VIEW And Me._PrimaryBundling <> -1 Then
                Me.ViewLogDirect()
            End If

            Me.SetTampilan()
            Me.ObjStockBundling.TipeBundling = ClsTrans_StockBundling.EnumTipeBundling.Produksi1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub ViewLogDirect()
        Try
            Me._Command = EnumCommand.NOTHING_
            If Me.txtremarks.Enabled Then
                If MsgBox("Are you sure to discard this data and view log?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                    Exit Sub
                End If
            End If
            ''button find
            Me.Cursor = Cursors.WaitCursor
            If Me.ObjStockBundling.CariDetail(Me._PrimaryBundling, Me.dgvItemb4, Me.dgvItemAfter, Me, Me.tbl_workunit) = True Then
                Me.ForceNavigate = True
                Me.tc_processoutput.SelectedTab = Me.tabpage_processoutput
                Me.SalinKeField()
            Else
                Me.Close()
            End If
            Me.NonAktifkanDetail()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Sub DeleteLogDirect1()
        Try
            ''button find
            Me.Cursor = Cursors.WaitCursor
            If Me.ObjStockBundling.CariDetail(Me._PrimaryBundling, Me.dgvItemb4, Me.dgvItemAfter, Me, Me.tbl_workunit) = True Then
                Me.ForceNavigate = True
                Me.tc_processoutput.SelectedTab = Me.tabpage_processoutput
                Me.SalinKeField()
                Me.Visible = True
                Me.BringToFront()
            Else
                Me.Close()
            End If

            ''button delete
            Me.Cursor = Cursors.WaitCursor
            If Me.ObjStockBundling.ApakahBolehHapus(Me.UserCode, True, True) = True Then
                If MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    Me.Cursor = Cursors.WaitCursor
                    If Me.ObjStockBundling.Hapus(Me.TxtPrimaryStockBundling.Text, Me.UserCode, True, True, Me) = True Then
                        MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                    End If
                End If
            End If
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Sub EditLogDirect1()
        Try
            ''button find
            Me.Cursor = Cursors.WaitCursor
            If Me.ObjStockBundling.CariDetail(Me._PrimaryBundling, Me.dgvItemb4, Me.dgvItemAfter, Me, Me.tbl_workunit) = True Then
                Me.ForceNavigate = True
                Me.tc_processoutput.SelectedTab = Me.tabpage_processoutput
                Me.SalinKeField()
            Else
                Me.Close()
            End If

            ''button edit 
            If Me.ObjStockBundling.ApakahBolehEdit(Me.UserCode) = True Then
                Me.baru = False
                Me.AktifkanDetail()
            Else
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Sub CreateLogDirect1()
        Dim objWorkUnit As New Core_Module.ClsMasterWorkUnit
        Try
            ''button new
            Me.baru = True
            Me.KosongkanField()
            Me.AktifkanDetail()
            Me.ObjStockBundling.ResetObject(Me.dgvItemb4, Me.dgvItemAfter)
            Me.txtstockbundling.Text = Me.ObjStockBundling.GetNoBaru(Me.dtpDate.Value, Me.tipe_bundling)
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString

            ''isi WO
            Me.txtidentifier_wo.Text = Me._IdentifierWO
            Me.TxtIdentifierJO.Text = Me._IdentifierJO
            Me.IsikanCboWorkUnit(Me._IdentifierWO, Me._IdentifierJO)
            Me.isiHeaderby_IdentifierWo()
            Me.IsiItemBy_IdentifierWo()
            Me.SetGboItemAfter()

            ''pilih work unit
            For i0 As Integer = 0 To Me.tbl_workunit.DefaultView.Count - 1
                Me.row = Me.tbl_workunit.DefaultView(i0).Row
                If Me.row.primary_work_unit = Me._PrimaryWorkUnit Then
                    Me.CboWorkUnit.SelectedIndex = i0
                End If
            Next

            ''tentukan focus user
            If Me.chkBatchNoControlledAfter.Checked Then
                Me.txtBatchNoAfter.Focus()
            ElseIf Me.chkLocationControlledAfter.Checked Then
                Me.TxtLocationAfter.Focus()
            Else
                Me.TxtQtyperBatchAfter.Focus()
            End If

            Me.TxtWorkShiftCode.Focus()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnSimpan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult

        Try
            If Me.ApakahDataBenar() = True Then
                Me.KosongkanItemAfter()
                Me.KosongkanItemB4()
                Me.SalinKeObjek()
                If baru = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjStockBundling.SimpanBaru(Me.dgvItemb4, Me.dgvItemAfter, Me.UserCode, Me) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnFind.Focus()
                        End If
                    Else
                        Exit Sub
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjStockBundling.SimpanPerubahan(Me.dgvItemb4, Me.dgvItemAfter, Me.UserCode, Me) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnEdit.Focus()
                        End If
                    Else
                        Exit Sub
                    End If
                End If

                If Me._Command <> EnumCommand.NOTHING_ And IsNothing(Me.MdiParent) Then
                    Me.Close()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub btnBatal_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.PerformCancel(sender, e)
    End Sub

    Sub PerformCancel(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim hasil As MsgBoxResult
        Try
            If Me.txtremarks.Text = "" And Me.dgvItemb4.RowCount = 0 And Me.dgvItemAfter.RowCount = 0 And _
            Me.dtpDate.Value.Date = Today And Me.TxtCodeNob4.Text = "" And Me.TxtCodeNoAfter.Text = "" And _
            Me.cboBatchNob4.Text = "" And Me.txtBatchNoAfter.Text = "" And Me.cboLocationb4.Text = "" And _
            Me.TxtLocationAfter.Text = "" And Me.TxtRemarkItemb4.Text = "" And Me.TxtRemarkItemAfter.Text = "" Then
                Me.NonAktifkanDetail()
                Me.btnNew.Focus()
                Me.KosongkanField()
            Else
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If baru = True Then
                        Me.KosongkanField()
                    Else
                        Me.Cursor = Cursors.WaitCursor
                        Me.ObjStockBundling.CariDetail(Me.TxtPrimaryStockBundling.Text, Me.dgvItemb4, Me.dgvItemAfter, Me, Me.tbl_workunit)
                        Me.SalinKeField()
                    End If
                Else
                    Exit Sub
                End If
                Me.NonAktifkanDetail()
                Me.btnNew.Focus()
            End If
            Me.Cursor = Cursors.Default
            If Me._Command <> EnumCommand.NOTHING_ And IsNothing(Me.MdiParent) Then
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Today Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            Me.baru = True
            Me.KosongkanField()
            Me.AktifkanDetail()
            Me.dtpDate.Focus()
            Me.ObjStockBundling.ResetObject(Me.dgvItemb4, Me.dgvItemAfter)
            Me.txtstockbundling.Text = Me.ObjStockBundling.GetNoBaru(Me.dtpDate.Value, Me.tipe_bundling)
            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim ij As New FrmCariTransaksiStockBundling

        Try
            ij.TipeBundling = Me.ObjStockBundling.TipeBundling
            ij.BackColorAll = Me.BackColorAll
            ij.TipeBundling = ClsTransaksiStockBundling.EnumTipeBundling.Produksi1
            ij.Text = "Find Production Log"
            ij.ShowDialog()
            If Not ij.PrimaryTransaksiStockBundling = -1 Then
                Me.Cursor = Cursors.WaitCursor
                If Me.ObjStockBundling.CariDetail(ij.PrimaryTransaksiStockBundling, _
                Me.dgvItemb4, Me.dgvItemAfter, Me, Me.tbl_workunit) = True Then
                    Me.ForceNavigate = True
                    Me.tc_processoutput.SelectedTab = Me.tabpage_processoutput
                    Me.SalinKeField()
                    Me.btnEdit.Focus()
                Else
                    Me.btnFind.Focus()
                End If
                Me.btnFind.Focus()
            End If
            Me.Cursor = Cursors.Default
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            If Me.txtstockbundling.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.txtstockbundling.Text <> "" Then
                If Me.ObjStockBundling.ApakahBolehEdit(Me.UserCode) = True Then
                    Me.baru = False
                    Me.AktifkanDetail()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.txtstockbundling.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.txtstockbundling.Text <> "" Then
                If Me.ObjStockBundling.ApakahBolehHapus(Me.UserCode, True, True) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjStockBundling.Hapus(Me.TxtPrimaryStockBundling.Text, Me.UserCode, True, True, Me) = True Then
                            Me.KosongkanField()
                            Me.NonAktifkanDetail()
                            Me.btnNew.Focus()
                            MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                        End If
                    End If
                End If

                If Me._Command <> EnumCommand.NOTHING_ Then
                    Me.Close()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub FrmMasterProductGroup_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.txtremarks.Enabled = True Then
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjStockBundling.SimpanBaru(Me.dgvItemb4, Me.dgvItemAfter, Me.UserCode, Me) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnFind.Focus()
                            Else
                                e.Cancel = True
                            End If
                        Else
                            'simpan perubahan
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjStockBundling.SimpanPerubahan(Me.dgvItemb4, Me.dgvItemAfter, Me.UserCode, Me) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnEdit.Focus()
                            Else
                                e.Cancel = True
                            End If
                        End If
                    Else
                        e.Cancel = True
                    End If
                ElseIf hasil = MsgBoxResult.Cancel Then
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiCboItem(ByVal vPrimaryProduct As Double, ByVal vPrimaryTransaksi As Double)
        Try
            'isikan dari basis data-persedian terakhir
            Me.ObjProduct.IsiCboBatchNo(vPrimaryProduct, Me.cboBatchNob4, Me.TBL)
            'tambahkan item yang sudah dicatatkan ke dalam database
            Dim row1 As DseTrans_StockBundling.tbltransaksi_stockbundling_item_before1Row
            Dim row2 As dseMasterProduct.tblmaster_productqty1Row
            Dim ketemu As Boolean
            If Me.baru = False Then
                Me.objStockBundlingItem.IsikanTBLBatchNo(Me.TBL4, vPrimaryProduct, vPrimaryTransaksi)
                'tambahkan semua baris di dalam transaksi stock out ke dalam persediaan saat ini
                While Me.TBL4.Rows.Count > 0
                    row1 = Me.TBL4.Rows(0)
                    For Each row2 In Me.TBL.Rows
                        If row1.Lot_No = row2.Lot_No Then
                            row2.Qty = row2.Qty + row1.Qty
                            ketemu = True
                            Exit For
                        End If
                    Next
                    If ketemu = False Then
                        'tambahkan baris baru
                        Me.TBL.Rows.Add(row1.Lot_No, row1.Expiry_Date, row1.Qty)
                    End If
                    Me.TBL4.Rows.Remove(row1)
                End While
            End If
            'kurangkan item yang sudah dimasukkan ke dgv semua
            Me.IsikanTBLBatchNo(vPrimaryProduct)
            While Me.TBL4.Rows.Count > 0
                row1 = Me.TBL4.Rows(0)
                For Each row2 In Me.TBL.Rows
                    If row1.Lot_No = row2.Lot_No Then
                        row2.Qty = row2.Qty - row1.Qty
                        If row2.Qty = 0 Then
                            Me.TBL.Rows.Remove(row2)
                        End If
                        ketemu = True
                        Exit For
                    End If
                Next
                If ketemu = False And TBL.Rows.Count > 0 Then
                    'tidak mungkin tidak ketemu
                    MsgBox("Logic error. Please contact your vendor if you see this message")
                    Exit Sub
                End If
                Me.TBL4.Rows.Remove(row1)
            End While
            Me.TBL.DefaultView.Sort = "[Lot No] ASC"
            Me.cboBatchNob4.DataSource = Me.TBL.DefaultView
            Me.cboBatchNob4.ValueMember = "Lot No"
            Me.cboBatchNob4.DisplayMember = "Lot No"
            Me.cboBatchNob4.Refresh()

            If Me.chkLocationControlledb4.Checked = False And _
            Me.chkUseSecondaryUnitb4.Checked = True And Me.cboBatchNob4.Text <> "" Then
                Me.IsiCboSecondary(vPrimaryProduct, vPrimaryTransaksi, _
                Me.cboBatchNob4.Text, Me.txtPrimaryLocationb4.Text)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanTBLBatchNo(ByVal vPrimaryProduct As Double)
        Try
            Dim row As DataGridViewRow
            Dim row1 As DseTrans_StockBundling.tbltransaksi_stockbundling_item_before1Row
            For Each row In Me.dgvItemb4.Rows
                If CDbl(row.Cells(12).Value) = vPrimaryProduct Then
                    Dim ketemu As Boolean = False
                    For Each row1 In Me.TBL4.Rows
                        If row1.Lot_No = CStr(row.Cells(7).Value) Then
                            'kalo ketemu ubah data qty
                            row1.Qty = row1.Qty + CDbl(row.Cells(3).Value)
                            ketemu = True
                            Exit For
                        End If
                    Next
                    'kalo ga ketemu tambahkan baris baru
                    If ketemu = False Then
                        Dim ED As Date = #1/1/1950#
                        If IsDate(row.Cells(8).Value) = True Then
                            ED = row.Cells(8).Value
                        End If
                        Me.TBL4.Rows.Add(CStr(row.Cells(7).Value), CDbl(row.Cells(3).Value), ED)
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiCboLokasi(ByVal vPrimaryProduct As Double, ByVal vPrimaryTransaksi As Double, _
    ByVal vBatchNo As String)
        Try
            'isikan tabel dari basis data
            Me.ObjProduct.IsiCboLocation(vPrimaryProduct, Me.cboLocationb4, _
            Me.TBL2, vBatchNo)
            Dim row1 As DseTrans_StockBundling.tbltransaksi_stockbundling_item_before2Row
            Dim row2 As dseMasterProduct.stock_location_viewRow
            Dim ketemu As Boolean
            If Me.baru = False Then
                Me.objStockBundlingItem.IsikanTBLLocation(Me.TBL5, vPrimaryProduct, vPrimaryTransaksi, _
                vBatchNo)
                'tambahkan semua baris di dalam transaksi stock out ke dalam persediaan saat ini
                While Me.TBL5.Rows.Count > 0
                    row1 = Me.TBL5.Rows(0)
                    For Each row2 In Me.TBL2.Rows
                        If row1.Lot_No = row2.Lot_No And row1.Location = row2.Location Then
                            row2.Qty = row2.Qty + row1.Qty
                            ketemu = True
                            Exit For
                        End If
                    Next
                    If ketemu = False Then
                        'tambahkan baris baru
                        Me.TBL2.Rows.Add(row1.Lot_No, row1.Location, row1.Qty)
                    End If
                    Me.TBL5.Rows.Remove(row1)
                End While
            End If
            'kurangkan item yang sudah dimasukkan ke dgv semua
            Me.IsikanTBLLocation(vPrimaryProduct, vBatchNo)
            While Me.TBL5.Rows.Count > 0
                row1 = Me.TBL5.Rows(0)
                For Each row2 In Me.TBL2.Rows
                    If row1.Lot_No = row2.Lot_No And row1.Location = row2.Location Then
                        row2.Qty = row2.Qty - row1.Qty
                        If row2.Qty = 0 Then
                            Me.TBL2.Rows.Remove(row2)
                        End If
                        ketemu = True
                        Exit For
                    End If
                Next
                If ketemu = False And Me.TBL2.Rows.Count > 0 Then
                    'tidak mungkin tidak ketemu
                    MsgBox("Logic error. Please contact your vendor if you see this message")
                    Exit Sub
                End If
                Me.TBL5.Rows.Remove(row1)
            End While

            Me.TBL2.DefaultView.Sort = "Location ASC"
            Me.cboLocationb4.DataSource = Me.TBL2.DefaultView
            Me.cboLocationb4.ValueMember = "Location"
            Me.cboLocationb4.DisplayMember = "Location"
            Me.cboLocationb4.Refresh()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanTBLLocation(ByVal vPrimaryProduct As Double, ByVal vBatchNo As String)
        Try
            Dim row As DataGridViewRow
            Dim row1 As DseTrans_StockBundling.tbltransaksi_stockbundling_item_before2Row
            For Each row In Me.dgvItemb4.Rows
                If CDbl(row.Cells(12).Value) = vPrimaryProduct And CStr(row.Cells(7).Value) = vBatchNo Then
                    Dim ketemu As Boolean = False
                    For Each row1 In Me.TBL5.Rows
                        If row1.Lot_No = CStr(row.Cells(7).Value) And row1.Location = row.Cells(9).Value Then
                            'kalo ketemu ubah data qty
                            row1.Qty = row1.Qty + CDbl(row.Cells(3).Value)
                            ketemu = True
                            Exit For
                        End If
                    Next
                    'kalo ga ketemu tambahkan baris baru
                    If ketemu = False Then
                        Me.TBL5.Rows.Add(CStr(row.Cells(7).Value), row.Cells(9).Value, _
                        CDbl(row.Cells(3).Value))
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiCboSecondary(ByVal vPrimaryProduct As Double, ByVal vPrimaryTransaksi As Double, _
    ByVal vBatchNo As String, ByVal vPrimaryLokasi As Double)
        Try
            'isikan tabel dari basis data
            Me.ObjProduct.IsiCboSecondaryUnit(vPrimaryProduct, Me.cboSecondaryUnitsb4, _
            Me.TBL3, vBatchNo, vPrimaryLokasi)
            Dim row1 As DseTrans_StockBundling.tbltransaksi_stockbundling_item_before3Row
            Dim row2 As dseMasterProduct.tblmaster_productqty2Row
            Dim ketemu As Boolean
            If Me.baru = False Then
                Me.objStockBundlingItem.IsikanTBLSecondaryUnit(Me.TBL6, vPrimaryProduct, _
                vPrimaryTransaksi, vBatchNo, vPrimaryLokasi)
                'tambahkan semua baris di dalam transaksi stock out ke dalam persediaan saat ini
                While Me.TBL6.Rows.Count > 0
                    row1 = Me.TBL6.Rows(0)
                    For Each row2 In Me.TBL3.Rows
                        If row1.Secondary_Qty = row2.Qty_Secondary Then
                            row2.Qty = row2.Qty + row1.Qty
                            ketemu = True
                            Exit For
                        End If
                    Next
                    If ketemu = False Then
                        'tambahkan baris baru
                        Me.TBL3.Rows.Add(row1.Secondary_Qty, row1.Qty)
                    End If
                    Me.TBL6.Rows.Remove(row1)
                End While
            End If
            'kurangkan item yang sudah dimasukkan ke dgv semua
            Me.IsikanTBLSecondary(vPrimaryProduct, vPrimaryLokasi, vBatchNo)
            While Me.TBL6.Rows.Count > 0
                row1 = Me.TBL6.Rows(0)
                For Each row2 In Me.TBL3.Rows
                    If row1.Secondary_Qty = row2.Qty_Secondary Then
                        row2.Qty = row2.Qty - row1.Qty
                        If row2.Qty = 0 Then
                            Me.TBL3.Rows.Remove(row2)
                        End If
                        ketemu = True
                        Exit For
                    End If
                Next
                If ketemu = False And Me.TBL3.Rows.Count > 0 Then
                    'tidak mungkin tidak ketemu
                    MsgBox("Logic error. Please contact your vendor if you see this message")
                    Exit Sub
                End If
                Me.TBL6.Rows.Remove(row1)
            End While

            Me.TBL3.DefaultView.Sort = "[Qty Secondary] ASC"
            Me.cboSecondaryUnitsb4.DataSource = Me.TBL3.DefaultView
            Me.cboSecondaryUnitsb4.ValueMember = "Qty Secondary"
            Me.cboSecondaryUnitsb4.DisplayMember = "Qty Secondary"
            Me.cboSecondaryUnitsb4.Refresh()
            'pilih salah satu
            If Me.cboSecondaryUnitsb4.Items.Count > 0 Then
                Me.cboSecondaryUnitsb4.SelectedIndex = -1
                Me.cboSecondaryUnitsb4.SelectedIndex = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanTBLSecondary(ByVal vPrimaryProduct As Double, ByVal vPrimaryLocation As Double, ByVal vBatchNo As String)
        Try
            Dim row As DataGridViewRow
            Dim row1 As DseTrans_StockBundling.tbltransaksi_stockbundling_item_before3Row
            For Each row In Me.dgvItemb4.Rows
                If CDbl(row.Cells(12).Value) = vPrimaryProduct And CStr(row.Cells(7).Value) = vBatchNo And _
                CDbl(row.Cells(13).Value) = vPrimaryLocation And IsNumeric(row.Cells(5).Value) = True Then
                    Dim ketemu As Boolean = False
                    For Each row1 In Me.TBL6.Rows
                        If row1.Secondary_Qty = CDbl(row.Cells(5).Value) Then
                            'kalo ketemu ubah data qty
                            row1.Qty = row1.Qty + CDbl(row.Cells(3).Value)
                            ketemu = True
                            Exit For
                        End If
                    Next
                    'kalo ga ketemu tambahkan baris baru
                    If ketemu = False Then
                        Dim ED As Date = #1/1/1950#
                        If IsDate(row.Cells(8).Value) = True Then
                            ED = row.Cells(8).Value
                        End If
                        Me.TBL6.Rows.Add(row.Cells(5).Value, row.Cells(3).Value, _
                        CStr(row.Cells(9).Value), CStr(row.Cells(7).Value))
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetGboItemAfter()
        Try
            Me.ObjProduct.IsiCheckBoxControll(Me.TxtPrimaryCodeNoAfter.Text, Me.chkLocationControlledAfter, _
            Me.chkBatchNoControlledAfter, Me.chkExpiryDateControlledAfter, Me.chkUseSecondaryUnitAfter, _
            Me.chkPdoControlledAfter, Me.txtPrimaryUnitsAfter, Me.txtPrimarySecondaryUnitsAfter)

            Me.txtPrimaryPDOAfter.Text = "-1"
            If Me.chkLocationControlledAfter.Checked = False Then
                Me.txtPrimaryLocationAfter.Text = "-1"
                Me.TxtLocationAfter.Text = ""
                Me.TxtWarehouseAfter.Clear()
                Me.TxtLocationAfter.ReadOnly = True
                Me.lblLocationAfter.Visible = True
                Me.lblWarehouseAfter.Visible = True
                Me.TxtLocationAfter.Visible = False
                Me.TxtWarehouseAfter.Visible = False
                Me.btnFindLocationAfter.Visible = False
            Else
                Me.TxtLocationAfter.ReadOnly = False
                Me.lblLocationAfter.Visible = False
                Me.lblWarehouseAfter.Visible = False
                Me.TxtLocationAfter.Visible = True
                Me.TxtWarehouseAfter.Visible = True
                Me.btnFindLocationAfter.Visible = True
            End If
            If Me.chkExpiryDateControlledAfter.Checked = True Then
                Me.lblExpiryDateAfter.Visible = False
                Me.DtpExpiredAfter.Visible = True
                Me.DtpExpiredAfter.Value = Today
            Else
                Me.lblExpiryDateAfter.Visible = True
                Me.DtpExpiredAfter.Visible = False
                Me.DtpExpiredAfter.Value = #1/1/1950#
            End If
            If Me.chkBatchNoControlledAfter.Checked = True Then
                Me.txtBatchNoAfter.ReadOnly = False
                Me.txtBatchNoAfter.TabStop = True
                Me.lblBatchNoAfter.Visible = False
                Me.txtBatchNoAfter.Visible = True
            Else
                Me.txtBatchNoAfter.ReadOnly = True
                Me.txtBatchNoAfter.TabStop = False
                Me.lblBatchNoAfter.Visible = True
                Me.txtBatchNoAfter.Visible = False
            End If
            If Me.chkUseSecondaryUnitAfter.Checked = False Then
                Me.txtSecondaryQtyAfter.Visible = False
                Me.lblSecondaryUnitsAfter.Visible = False
                Me.lblUnitsAfter.Left = Me.lblSecondaryUnitsAfter.Left
                Me.TxtQtyperBatchAfter.Width = Me.lblUnitsAfter.Left - Me.TxtQtyperBatchAfter.Left - 6
                Me.txtSecondaryQtyAfter.Text = "-1"
                Me.lblSecQtyAfter.Visible = False
            Else
                Me.txtSecondaryQtyAfter.Visible = True
                Me.lblSecondaryUnitsAfter.Visible = True
                Me.lblSecQtyAfter.Visible = True
                Me.lblUnitsAfter.Left = Me.lblSecQtyAfter.Left - 37
                Me.TxtQtyperBatchAfter.Width = Me.lblUnitsAfter.Left - Me.TxtQtyperBatchAfter.Left - 6
                Me.txtSecondaryQtyAfter.Text = "0"
            End If
            If Me.TxtPrimaryCodeNoAfter.Text <> "-1" Then
                Me.lblUnitsAfter.Text = Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNoAfter.Text)
                If Me.chkUseSecondaryUnitAfter.Checked = True Then
                    Me.lblSecondaryUnitsAfter.Text = Me.ObjProduct.GetSecondaryUnitCode(Me.TxtPrimaryCodeNoAfter.Text)
                End If
            Else
                Me.lblUnitsAfter.Text = "Units"
                Me.lblSecondaryUnitsAfter.Text = "Units"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetGboItemb4()
        Try
            Me.ObjProduct.IsiCheckBoxControll(Me.TxtPrimaryCodeNob4.Text, Me.chkLocationControlledb4, _
            Me.chkBatchNoControlledb4, Me.chkExpiryDateControlledb4, Me.chkUseSecondaryUnitb4, _
            Me.chkPdoControlledb4, Me.txtPrimaryUnitsb4, Me.txtPrimarySecondaryUnitsb4)

            Me.cboBatchNob4.DataSource = Nothing
            Me.cboLocationb4.DataSource = Nothing
            Me.cboSecondaryUnitsb4.DataSource = Nothing

            Me.txtPrimaryPDOb4.Text = "-1"
            If Me.chkBatchNoControlledb4.Checked = True Then
                Me.cboBatchNob4.TabStop = True
                Me.lblBatchNob4.Visible = False
                Me.cboBatchNob4.Visible = True
                'isikan data batch no dari basis data
                Me.IsiCboItem(Me.TxtPrimaryCodeNob4.Text, Me.TxtPrimaryStockBundling.Text)
            Else
                Me.cboBatchNob4.TabStop = False
                Me.lblBatchNob4.Visible = True
                Me.cboBatchNob4.Visible = False
                Me.cboBatchNob4.Items.Clear()
            End If
            If Me.chkLocationControlledb4.Checked = False Then
                Me.txtPrimaryLocationb4.Text = "-1"
                Me.cboLocationb4.Items.Clear()
                Me.TxtWarehouseb4.Clear()
                Me.lblLocationb4.Visible = True
                Me.lblWarehouseb4.Visible = True
                Me.cboLocationb4.Visible = False
                Me.TxtWarehouseb4.Visible = False
            Else
                Me.lblLocationb4.Visible = False
                Me.lblWarehouseb4.Visible = False
                Me.cboLocationb4.Visible = True
                Me.TxtWarehouseb4.Visible = True
                If Me.chkLocationControlledb4.Checked = True Then
                    Me.IsiCboLokasi(Me.TxtPrimaryCodeNob4.Text, Me.TxtPrimaryStockBundling.Text, Me.cboBatchNob4.Text)
                End If
            End If
            If Me.chkExpiryDateControlledb4.Checked = True Then
                Me.lblExpiryDateb4.Visible = False
                Me.DtpExpiredb4.Visible = True
            Else
                Me.lblExpiryDateb4.Visible = True
                Me.DtpExpiredb4.Visible = False
                Me.DtpExpiredb4.Value = #1/1/1950#
            End If

            If Me.chkUseSecondaryUnitb4.Checked = False Then
                Me.lblSecondaryUnitsb4.Visible = False
                Me.lblUnitsb4.Left = Me.lblSecondaryUnitsb4.Left
                Me.TxtQtyperBatchb4.Width = Me.lblUnitsb4.Left - Me.TxtQtyperBatchb4.Left - 6
                Me.txtSecondaryQtyb4.Text = "-1"
                Me.lblSecQtyb4.Visible = False
                Me.cboSecondaryUnitsb4.Visible = False
                Me.cboSecondaryUnitsb4.DataSource = Nothing
                Me.cboSecondaryUnitsb4.Items.Clear()
            Else
                Me.lblSecondaryUnitsb4.Visible = True
                Me.lblSecQtyb4.Visible = True
                Me.lblUnitsb4.Left = Me.lblSecQtyb4.Left - 37
                Me.TxtQtyperBatchb4.Width = Me.lblUnitsb4.Left - Me.TxtQtyperBatchb4.Left - 6
                Me.cboSecondaryUnitsb4.Visible = True
                If Me.chkLocationControlledb4.Checked = False Then
                    'isikan data secondary jika perlu
                    Me.IsiCboSecondary(Me.TxtPrimaryCodeNob4.Text, Me.TxtPrimaryStockBundling.Text, _
                    Me.cboBatchNob4.Text, Me.txtPrimaryLocationb4.Text)
                End If
            End If
            If Me.TxtPrimaryCodeNob4.Text <> "-1" Then
                Me.lblUnitsb4.Text = Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNob4.Text)
                If Me.chkUseSecondaryUnitb4.Checked = True Then
                    Me.lblSecondaryUnitsb4.Text = Me.ObjProduct.GetSecondaryUnitCode(Me.TxtPrimaryCodeNob4.Text)
                End If
            Else
                Me.lblUnitsb4.Text = "Units"
                Me.lblSecondaryUnitsb4.Text = "Units"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeNob4_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCodeNob4.KeyDown
        Dim i As New frmCariMasterProduct2
        Dim Jml As Long = 0
        Dim PrimaryProduct As Double = -1

        Try
            If e.KeyValue = 13 Then
                Jml = Me.ObjProduct.GetJml(Me.TxtCodeNob4.Text)
                PrimaryProduct = Me.ObjProduct.GetPrimary(Me.TxtCodeNob4.Text)

                If Jml = 1 And Me.ObjProduct.ApakahGoods(PrimaryProduct) = True Then
                    Me.TxtPrimaryCodeNob4.Text = Me.ObjProduct.GetPrimary(Me.TxtCodeNob4.Text)
                    Me.txtPrimaryCategoryb4.Text = Me.ObjProduct.GetPrimaryCategory(Me.TxtPrimaryCodeNob4.Text)
                    Me.txtDescb4.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNob4.Text)
                    Me.SetGboItemb4()
                    If Me.cboBatchNob4.Visible = True Then
                        Me.cboBatchNob4.Focus()
                    ElseIf Me.cboLocationb4.Visible = True Then
                        Me.cboLocationb4.Focus()
                    Else
                        Me.TxtQtyperBatchb4.Focus()
                    End If
                ElseIf Jml > 1 Or Jml = 0 Then
                    'ada 2 product dengan kode yang sama atau tidak ada product yang dimaksud
                    If Me.TxtCodeNob4.Text <> "" Then
                        i.txtCari.Text = Me.TxtCodeNob4.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        Me.TxtPrimaryCodeNob4.Text = i.PrimaryProduct
                        Me.TxtCodeNob4.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryCodeNob4.Text)
                        Me.txtDescb4.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNob4.Text)
                        Me.SetGboItemb4()
                        If Me.cboBatchNob4.Visible = True Then
                            Me.cboBatchNob4.Focus()
                        ElseIf Me.cboLocationb4.Visible = True Then
                            Me.cboLocationb4.Focus()
                        Else
                            Me.TxtQtyperBatchb4.Focus()
                        End If
                    Else
                        Me.TxtCodeNob4.Clear()
                        Me.TxtPrimaryCodeNob4.Text = "-1"
                        Me.txtPrimaryCategoryb4.Text = "-1"
                        Me.KosongkanItemB4()
                        Me.TxtCodeNob4.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtCodeNob4_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNob4.Leave
        Dim i As New frmCariMasterProduct2
        Dim PrimaryProduct As Double = -1

        Try
            If Me.TxtCodeNob4.Text <> "" Then
                PrimaryProduct = Me.ObjProduct.GetPrimary(Me.TxtCodeNob4.Text)

                If Me.ObjProduct.GetJml(Me.TxtCodeNob4.Text) = 1 And Me.ObjProduct.ApakahGoods(PrimaryProduct) = True Then
                    Me.TxtPrimaryCodeNob4.Text = Me.ObjProduct.GetPrimary(Me.TxtCodeNob4.Text)
                    Me.txtPrimaryCategoryb4.Text = Me.ObjProduct.GetPrimaryCategory(Me.TxtPrimaryCodeNob4.Text)
                    Me.txtDescb4.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNob4.Text)
                    Me.SetGboItemb4()
                    If Me.cboBatchNob4.Visible = True Then
                        Me.cboBatchNob4.Focus()
                    ElseIf Me.cboLocationb4.Visible = True Then
                        Me.cboLocationb4.Focus()
                    Else
                        Me.TxtQtyperBatchb4.Focus()
                    End If
                Else
                    If Me.TxtCodeNob4.Text <> "" Then
                        i.txtCari.Text = Me.TxtCodeNob4.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        Me.TxtPrimaryCodeNob4.Text = i.PrimaryProduct
                        Me.TxtCodeNob4.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryCodeNob4.Text)
                        Me.txtDescb4.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNob4.Text)
                        Me.SetGboItemb4()
                        If Me.cboBatchNob4.Visible = True Then
                            Me.cboBatchNob4.Focus()
                        ElseIf Me.cboLocationb4.Visible = True Then
                            Me.cboLocationb4.Focus()
                        Else
                            Me.TxtQtyperBatchb4.Focus()
                        End If
                    Else
                        Me.TxtCodeNob4.Clear()
                        Me.TxtPrimaryCodeNob4.Text = "-1"
                        Me.txtPrimaryCategoryb4.Text = "-1"
                        Me.KosongkanItemB4()
                        Me.TxtCodeNob4.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryCodeNob4.Text = "-1"
                Me.txtDescb4.Clear()
                Me.txtPrimaryCategoryb4.Text = "-1"
                Me.KosongkanItemB4()
            End If
            Me.SetGboItemb4()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindProductb4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindCodeNob4.Click
        Dim i As New frmCariMasterProduct2

        Try
            i.cboCriteria.SelectedIndex = 0
            If Me.TxtCodeNob4.Text <> "" Then
                i.txtCari.Text = Me.TxtCodeNob4.Text
            End If

            i.ShowDialog()
            If i.PrimaryProduct > 0 Then
                Me.TxtPrimaryCodeNob4.Text = i.PrimaryProduct
                Me.TxtCodeNob4.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryCodeNob4.Text)
                Me.txtDescb4.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNob4.Text)
                Me.SetGboItemb4()
                If Me.cboBatchNob4.Visible = True Then
                    Me.cboBatchNob4.Focus()
                ElseIf Me.cboLocationb4.Visible = True Then
                    Me.cboLocationb4.Focus()
                Else
                    Me.TxtQtyperBatchb4.Focus()
                End If
            Else
                Me.TxtCodeNob4.Clear()
                Me.TxtPrimaryCodeNob4.Text = "-1"
                Me.txtPrimaryCategoryb4.Text = "-1"
                Me.KosongkanItemB4()
                Me.TxtCodeNob4.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    ''' <summary>
    ''' mengembalikan index baris dalam dgv yang sesuai dengan parameter yang diberikan
    ''' </summary>
    ''' <param name="vPrimaryProduct"></param>
    ''' <param name="vPrimaryLocation"></param>
    ''' <param name="BatchNo"></param>
    ''' <param name="SecondaryQty"></param>
    ''' <param name="PrimaryPdO"></param>
    ''' <param name="ExpiryDate"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetIndex(ByVal vPrimaryProduct As Double, ByVal vPrimaryLocation As Double, ByVal BatchNo As String, _
    ByVal SecondaryQty As Decimal, ByVal PrimaryPdO As Double, ByVal ExpiryDate As Date, ByVal vDGV As DataGridView) As Integer
        Dim row As DataGridViewRow

        Try
            For Each row In vDGV.Rows
                If IsDate(row.Cells(8).Value) = False Then
                    If IsNumeric(row.Cells(5).Value) = False Then
                        If row.Cells(12).Value = vPrimaryProduct And row.Cells(13).Value = vPrimaryLocation And _
                        row.Cells(7).Value = BatchNo And SecondaryQty = -1 _
                        Then
                            Return row.Index
                        End If
                    Else
                        If row.Cells(12).Value = vPrimaryProduct And row.Cells(13).Value = vPrimaryLocation And _
                        row.Cells(7).Value = BatchNo And row.Cells(5).Value = SecondaryQty And _
                        ExpiryDate = #1/1/1950# Then
                            Return row.Index
                        End If
                    End If
                Else
                    If IsNumeric(row.Cells(5).Value) = False Then
                        If row.Cells(12).Value = vPrimaryProduct And row.Cells(13).Value = vPrimaryLocation And _
                        row.Cells(7).Value = BatchNo And SecondaryQty = -1 And _
                        ExpiryDate = CDate(row.Cells(8).Value) Then
                            Return row.Index
                        End If
                    Else
                        If row.Cells(12).Value = vPrimaryProduct And row.Cells(13).Value = vPrimaryLocation And _
                        row.Cells(7).Value = BatchNo And row.Cells(5).Value = SecondaryQty And _
                        ExpiryDate = CDate(row.Cells(8).Value) Then
                            Return row.Index
                        End If
                    End If
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub BtnInputb4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnInputb4.Click
        Try
            Me.InputB4(True, True, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub InputB4(ByVal vCekBalanceRM As Boolean, ByVal vFollowUpFormatDGV As Boolean, ByVal vCekApakahDataBenar As Boolean)
        Dim Lanjut As Boolean = False

        Try
            If vCekApakahDataBenar = True Then
                Lanjut = Me.ApakahDataItemB4Benar(vCekBalanceRM)
            Else
                Lanjut = True
            End If

            If Lanjut = True Then
                Dim Index As Integer = Me.GetIndex(Me.TxtPrimaryCodeNob4.Text, Me.txtPrimaryLocationb4.Text, Me.cboBatchNob4.Text, _
                 Me.txtSecondaryQtyb4.Text, Me.txtPrimaryPDOb4.Text, Me.DtpExpiredb4.Value, Me.dgvItemb4)
                'cari data yang sama persis
                If Index <> -1 Then
                    'tambahkan nilai qty
                    Dim row As DataGridViewRow = Me.dgvItemb4.Rows(Index)
                    row.Cells(3).Value = CDbl(row.Cells(3).Value) + CDbl(Me.TxtQtyperBatchb4.Text)
                Else
                    'tambahkan baris baru
                    Dim ExpiredDateStr As String
                    If Me.DtpExpiredb4.Value.Date = #1/1/1950# Then
                        ExpiredDateStr = "-"
                    Else
                        ExpiredDateStr = Format(Me.DtpExpiredb4.Value.Date, "dd MMM yyyy")
                    End If
                    Me.dgvItemb4.Rows.Add("", Me.TxtCodeNob4.Text, Me.txtDescb4.Text, Me.ObjCD.DigitGrouping(Me.TxtQtyperBatchb4.Text, Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNob4.Text)), _
                    Me.lblUnitsb4.Text, Me.txtSecondaryQtyb4.Text, Me.lblSecondaryUnitsb4.Text, Me.cboBatchNob4.Text, _
                    ExpiredDateStr, Me.cboLocationb4.Text, Me.TxtWarehouseb4.Text, Me.TxtRemarkItemb4.Text, _
                    Me.TxtPrimaryCodeNob4.Text, Me.txtPrimaryLocationb4.Text, Me.txtPrimaryUnitsb4.Text, _
                    Me.txtPrimarySecondaryUnitsb4.Text)
                    'set ulang nomor
                    If vFollowUpFormatDGV = True Then
                        Me.ObjCD.SetNumber(Me.dgvItemb4)
                    End If
                End If

                If vFollowUpFormatDGV = True Then
                    'kosongkan gbo item
                    Me.KosongkanItemB4()
                    Me.TxtCodeNob4.Focus()
                    'format dgv
                    Me.FormatDGV(Me.dgvItemb4)
                    'fokuskan ke baris terakhir
                    While Me.dgvItemb4.Rows(Me.dgvItemb4.Rows.Count - 1).Displayed = False _
                    Or Me.dgvItemb4.DisplayedRowCount(True) <> Me.dgvItemb4.DisplayedRowCount(False)
                        Me.dgvItemb4.FirstDisplayedScrollingRowIndex = Me.dgvItemb4.FirstDisplayedScrollingRowIndex + 1
                    End While
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FormatDGV(ByVal vDGV As DataGridView)
        Dim row As DataGridViewRow
        Dim TidakAdaSecondary As Boolean = True
        Dim TidakAdaLokasi As Boolean = True
        Dim TidakAdaBatchNo As Boolean = True
        Dim TidakAdaExpiryDate As Boolean = True

        Try
            'yang diformat dgv-nya:
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            For Each row In vDGV.Rows
                '1.a. untuk expiry date yang nilainya sama dengan 1/1/1950 dibuat "-"
                If IsDate(row.Cells(8).Value) = False Then
                    row.Cells(8).Value = "-"
                Else
                    If CDate(row.Cells(8).Value).Date = #1/1/1950# Then
                        row.Cells(8).Value = "-"
                    Else
                        TidakAdaExpiryDate = False
                        'format sehingga bisa dibaca lebih enak
                        row.Cells(8).Value = Format(CDate(row.Cells(8).Value), "dd MMM yyyy")
                    End If
                End If
                If IsNumeric(row.Cells(5).Value) = True Then
                    If row.Cells(5).Value <> "-1" Then
                        TidakAdaSecondary = False
                    End If
                End If
                '1.b. untuk secondary unit yang nilainya -1 maka tampilkan qty secondarynya "-" dan 
                '     secondary unitsnya "-"
                If row.Cells(5).Value = "-1" Then
                    row.Cells(5).Value = "-"
                    row.Cells(6).Value = "-"
                End If
                '1.c. untuk primary location =-1 maka code lokasi dan warehouse = "-"
                If row.Cells(13).Value = "-1" Then
                    row.Cells(9).Value = "-"
                    row.Cells(10).Value = "-"
                Else
                    TidakAdaLokasi = False
                End If
                '1.d. tandai bila ada batch no yang tidak kosong
                If row.Cells(7).Value <> "" Then
                    TidakAdaBatchNo = False
                End If
            Next
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            '2. jika semua row pada dgvitem tidak menggunakan secondary unit maka coloumn secondary unit
            '   tidak perlu ditampilkan
            If TidakAdaSecondary = True Then
                vDGV.Columns(5).Visible = False
                vDGV.Columns(6).Visible = False
            Else
                vDGV.Columns(5).Visible = True
                vDGV.Columns(6).Visible = True
            End If
            '3. jika semua data lokasi tidak ada maka tidak perlu tampilkan kolom lokasi dan warehouse
            If TidakAdaLokasi = True Then
                vDGV.Columns(9).Visible = False
                vDGV.Columns(10).Visible = False
            Else
                vDGV.Columns(9).Visible = True
                vDGV.Columns(10).Visible = True
            End If
            '4. jika semua data batch no tidak ada maka tidak perlu tampilkan kolom batch no
            If TidakAdaBatchNo = True Then
                vDGV.Columns(7).Visible = False
            Else
                vDGV.Columns(7).Visible = True
            End If
            '5. jika semua data expiry date tidak ada, tidak perlu tampilkan kolom expiry date
            If TidakAdaExpiryDate = True Then
                vDGV.Columns(8).Visible = False
            Else
                vDGV.Columns(8).Visible = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatchb4_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchb4.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtQtyperBatchb4, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatchb4_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyperBatchb4.KeyDown
        Me.ObjCD.Number_KeyDown(Me.TxtQtyperBatchb4, sender, e)
    End Sub

    Private Sub TxtQtyperBatchb4_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyperBatchb4.KeyUp
        Me.ObjCD.Number_KeyUp(Me.TxtQtyperBatchb4, sender, e)
    End Sub

    Private Sub txtSecondaryQtyb4_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSecondaryQtyb4.KeyDown
        Me.ObjCD.Number_KeyDown(Me.txtSecondaryQtyb4, sender, e)
    End Sub

    Private Sub txtSecondaryQtyb4_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSecondaryQtyb4.KeyUp
        Me.ObjCD.Number_KeyUp(Me.txtSecondaryQtyb4, sender, e)
    End Sub

    Private Sub dgvItemb4_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItemb4.CellBeginEdit
        Try
            'yang bisa diedit adalah:
            '1. Qty
            '2. Secondary Qty
            '3. Batch No
            '4. Location Code
            '5. Remarks
            'karena stock out, data expiry date tidak bisa diedit, karena data ini mengikuti data batch no
            If Me.txtremarks.Enabled = True Then
                If e.ColumnIndex = 3 Or e.ColumnIndex = 5 Or e.ColumnIndex = 7 Or e.ColumnIndex = 9 _
                Or e.ColumnIndex = 11 Then
                    'untuk batch no, secondary qty, location, expiry data -> hanya bisa diedit jika memang 
                    'masing-masing parameter yang akan diedit dikontrol
                    If e.ColumnIndex = 5 Then
                        'cek secondary qty
                        If CStr(Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value) = "-" Then
                            e.Cancel = True
                            Exit Sub
                        End If
                    ElseIf e.ColumnIndex = 7 Then
                        'cek batch no
                        If Me.ObjProduct.ApakahBatchNoControlled(Me.dgvItemb4.Item(12, e.RowIndex).Value) = False Then
                            e.Cancel = True
                            Exit Sub
                        End If
                    ElseIf e.ColumnIndex = 8 Then
                        'cek expiry date
                        If Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value = "-" Then
                            e.Cancel = True
                            Exit Sub
                        End If
                    ElseIf e.ColumnIndex = 9 Then
                        'cek location controlled
                        If Me.ObjProduct.ApakahLocationControlled(Me.dgvItemb4.Item(12, e.RowIndex).Value) = False Then
                            e.Cancel = True
                            Exit Sub
                        End If
                    End If
                    'kalo sampe sini maka boleh edit
                    Me.buffer = Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                'hanya bisa edit pada saat detail aktif
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemb4_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItemb4.CellEndEdit
        Try
            Dim SecondaryQty, JmlDGV As Decimal
            Dim BatchNo As String = ""
            Dim ExpiryDate As Date = #1/1/1950#
            Dim row As DataGridViewRow = Me.dgvItemb4.Rows(e.RowIndex)

            'pastikan ada isinya
            With row
                If IsNothing(Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColRemark.Index Then
                        If IsNothing(Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value) Then
                            Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                        Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.Replace("'", "").Replace("%", "").Replace("*", "").Trim
                    End If

                    'Bila tidak ada perubahan data maka tidak perlu dilakukan cek
                    If e.ColumnIndex <> 8 Then
                        If Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer Then
                            Exit Sub
                        End If
                    Else
                        If IsDate(Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            If CDate(Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value).Date = CDate(Me.buffer).Date Then
                                Exit Sub
                            End If
                        End If
                    End If

                    If e.ColumnIndex <> 11 Then
                        'kalo bukan data remarks maka perlu dicek kebenaran data hasil edit
                        If IsDate(.Cells(8).Value) = True Then
                            ExpiryDate = CDate(.Cells(8).Value)
                        Else
                            ExpiryDate = #1/1/1950#
                        End If
                        If IsNumeric(.Cells(5).Value) = True Then
                            SecondaryQty = .Cells(5).Value
                        Else
                            SecondaryQty = -1
                        End If
                        JmlDGV = 0

                        If e.ColumnIndex = 3 Then
                            Me.EndEditQtyb4(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                        ElseIf e.ColumnIndex = 5 Then
                            Me.EndEditSecondaryUnitb4(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                        ElseIf e.ColumnIndex = 7 Then
                            Me.EndEditBatchNob4(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                        ElseIf e.ColumnIndex = 9 Then
                            Me.EndEditLocationb4(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                        End If
                    End If
                Else
                    If e.ColumnIndex = 11 Then
                        Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Else
                        Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditQtyb4(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal SecondaryQty As Decimal, ByVal JmlDGV As Decimal, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
    ByVal row As DataGridViewRow)
        Try
            With row
                'untuk nilai qty dan secondary qty, pastikan nilainya positif
                If Me.ObjCD.CekBilanganPositif(Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("Please enter positive number for quantity data.", MsgBoxStyle.Information, _
                    "Information")
                    Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                End If

                If ApakahQtyRMUsageSudahBenar(Me.TxtPrimaryStockBundling.Text, Me.dgvItemb4.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, Me.txtidentifier_wo.Text, _
                 Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value, True, Me.TxtIdentifierJO.Text) = False Then
                    Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If

                'digit grouping
                Me.dgvItemb4(e.ColumnIndex, e.RowIndex).Value = _
                Me.ObjCD.DigitGrouping(Me.dgvItemb4(e.ColumnIndex, e.RowIndex).Value, _
                Me.ObjProduct.GetAngkaDecimal(Me.dgvItemb4.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value))
                'cek apakah jumlah digudang akan menjadi minus setelah perubahan data
                JmlDGV = .Cells(3).Value
                If Me.ApakahEditSahb4(row, .Cells(13).Value, -1, .Cells(12).Value, _
                .Cells(7).Value, ExpiryDate, SecondaryQty, JmlDGV, True) = False Then
                    Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditSecondaryUnitb4(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal SecondaryQty As Decimal, ByVal JmlDGV As Decimal, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
    ByVal row As DataGridViewRow)
        Try
            With row
                'untuk nilai qty dan secondary qty, pastikan nilainya positif
                If Me.ObjCD.CekBilanganPositif(Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("Please enter positive number for @ data.", MsgBoxStyle.Information, _
                    "Information")
                    Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
                'cek apakah jumlah digudang akan menjadi minus setelah perubahan data
                SecondaryQty = Me.buffer
                If Me.ApakahEditSahb4(row, .Cells(13).Value, -1, .Cells(12).Value, _
                .Cells(7).Value, ExpiryDate, .Cells(5).Value, JmlDGV) = False Then
                    Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                Else
                    'cek apakah ada kemungkinan merger
                    Dim AdaSama As Boolean = Me.ApakahAdaBarisYangSama(Me.dgvItemb4)
                    If AdaSama = True Then
                        If MsgBox("Your data change will merge 2 rows into 1. Do you still want to continue?", MsgBoxStyle.YesNo, _
                        "Confirm") = MsgBoxResult.No Then
                            Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                        Me.GabungkanBaris(Me.dgvItemb4)
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditBatchNob4(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal SecondaryQty As Decimal, ByVal JmlDGV As Decimal, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
    ByVal row As DataGridViewRow)
        Try
            With row
                'pastikan diisikan
                If CStr(Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value) = "" Then
                    Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
                'batch no selalu huruf besar
                Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value = CStr(Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value).ToUpper
                'batch no
                If Me.ApakahEditSahb4(row, .Cells(13).Value, -1, .Cells(12).Value, _
                .Cells(7).Value, ExpiryDate, SecondaryQty, JmlDGV) = False Then
                    Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                Else
                    'khusus bila product yang bersangkutan expiry date controlled juga maka
                    'cek ke dalam basis data:
                    Dim BufferExpiryDate As String = "-"
                    '1. data tanggal kadaluarsanya berbeda dengan yang sudah dicatatkan kedalam basis data dan datagridview
                    If IsDate(Me.dgvItemb4.Item(8, e.RowIndex).Value) = True Then
                        'basis data
                        Dim ExpiryDateLama As Date = Me.ObjProduct.GetExpiryDate(Me.dgvItemb4.Item(12, e.RowIndex).Value, _
                        Me.dgvItemb4.Item(7, e.RowIndex).Value)
                        If ExpiryDateLama <> #1/1/1950# Then
                            If ExpiryDateLama.Date <> CDate(Me.dgvItemb4(8, e.RowIndex).Value) Then
                                If MsgBox("Expiry date for batch no: " + CStr(Me.dgvItemb4.Item(7, e.RowIndex).Value) + " is " + _
                                Format(ExpiryDateLama, "dd MMM yyyy") + _
                                ". Changing your batch no will also change your expiry date data. Continue?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
                                    Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    Exit Sub
                                Else
                                    'update data expiry date juga
                                    BufferExpiryDate = CDate(Me.dgvItemb4.Item(8, e.RowIndex).Value)
                                    Me.dgvItemb4.Item(8, e.RowIndex).Value = Format(ExpiryDateLama, "dd MMM yyyy")
                                End If
                            End If
                        End If
                        'data grid view
                        Dim row1 As DataGridViewRow
                        For Each row1 In Me.dgvItemb4.Rows
                            If row1.Cells(7).Value = Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value _
                            And row1.Index <> e.RowIndex And row1.Cells(12).Value = Me.dgvItemb4.Item(12, e.RowIndex).Value Then
                                If CDate(row1.Cells(8).Value) <> CDate(Me.dgvItemb4.Item(8, e.RowIndex).Value) Then
                                    If MsgBox("Expiry date for batch no: " + CStr(Me.dgvItemb4.Item(7, e.RowIndex).Value) + " is " + _
                                        Format(CDate(row1.Cells(8).Value), "dd MMM yyyy") + _
                                        ". Changing your batch no will also change your expiry date data. Continue?", _
                                        MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                                        'update data expiry date juga
                                        BufferExpiryDate = CDate(Me.dgvItemb4.Item(8, e.RowIndex).Value)
                                        Me.dgvItemb4.Item(8, e.RowIndex).Value = row1.Cells(8).Value
                                        Exit For
                                    Else
                                        Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        Exit Sub
                                    End If
                                End If
                            End If
                        Next
                    End If
                    '3. cek apakah ada kemungkinan merger
                    Dim AdaSama As Boolean = Me.ApakahAdaBarisYangSama(Me.dgvItemb4)
                    If AdaSama = True Then
                        If MsgBox("Your data change will merge 2 rows into 1. Do you still want to continue?", MsgBoxStyle.YesNo, _
                        "Confirm") = MsgBoxResult.No Then
                            'kembalikan data batch no dan expiry date
                            Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Me.dgvItemb4.Item(8, e.RowIndex).Value = BufferExpiryDate
                            Exit Sub
                        End If
                        Me.GabungkanBaris(Me.dgvItemb4)
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditLocationb4(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal SecondaryQty As Decimal, ByVal JmlDGV As Decimal, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
    ByVal row As DataGridViewRow)
        Dim Obj As New clsMasterStorageArea
        Try
            With row
                'pastikan kode gudang adalah huruf kapital
                Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value = CStr(Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value).ToUpper
                'cek apakah kode lokasi yang dimasukkan sah
                If Obj.ApakahSudahAda(Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    Dim i As New FrmCariMasterStorageArea
                    If CStr(Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value) <> "" Then
                        i.txtCari.Text = Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value
                    End If

                    i.ShowDialog()
                    If i.PrimaryStorageArea <> -1 Then
                        Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value = CStr(Obj.GetStorageAreaCode(i.PrimaryStorageArea)).ToUpper
                    Else
                        Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        i = Nothing
                        Exit Sub
                    End If
                    i = Nothing
                Else
                    Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value = CStr(Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value).ToUpper
                End If
                'update data primary location
                .Cells(13).Value = Obj.GetPrimary(.Cells(9).Value)
                'update data gudang
                .Cells(10).Value = Obj.GetWarehouseCode(.Cells(13).Value)

                'cek jangan sampai stock digudang menjadi negatif
                If Me.ApakahEditSahb4(row, .Cells(13).Value, -1, .Cells(12).Value, _
                .Cells(7).Value, ExpiryDate, SecondaryQty, JmlDGV) = False Then
                    'kembalikan ke data semula
                    Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    .Cells(13).Value = Obj.GetPrimary(.Cells(9).Value)
                    'update data gudang
                    .Cells(10).Value = Obj.GetWarehouseCode(.Cells(13).Value)
                Else
                    'cek apakah ada kemungkinan merger
                    Dim AdaSama As Boolean = Me.ApakahAdaBarisYangSama(Me.dgvItemb4)
                    If AdaSama = True Then
                        If MsgBox("Your data change will merge 2 rows into 1. Do you still want to continue?", MsgBoxStyle.YesNo, _
                        "Confirm") = MsgBoxResult.No Then
                            'kembalikan ke data semula
                            Me.dgvItemb4.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            'update data primary location
                            .Cells(13).Value = Obj.GetPrimary(.Cells(9).Value)
                            'update data gudang
                            .Cells(10).Value = Obj.GetWarehouseCode(.Cells(13).Value)
                            Exit Sub
                        End If
                        Me.GabungkanBaris(Me.dgvItemb4)
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Obj = Nothing
        End Try
    End Sub

    Private Function ApakahAdaBarisYangSama(ByVal vDgv As DataGridView) As Boolean
        Dim row, row1 As DataGridViewRow
        Dim i, j As Integer

        Try
            While i <= vDgv.RowCount - 2
                row = vDgv.Rows(i)
                j = i + 1
                While j <= vDgv.RowCount - 1
                    row1 = vDgv.Rows(j)
                    '2 baris sama apabila semua data sama kecuali: data qty, unit, sec units, expiry date, 
                    'warehouse

                    If row.Cells(12).Value = row1.Cells(12).Value And row.Cells(2).Value = row1.Cells(2).Value And _
                    row.Cells(3).Value = row1.Cells(3).Value And row.Cells(4).Value = row1.Cells(4).Value And _
                    CStr(row.Cells(7).Value) = CStr(row1.Cells(7).Value) And row.Cells(9).Value = _
                    row1.Cells(9).Value Then
                        If IsNumeric(row.Cells(5).Value) = True And IsNumeric(row1.Cells(5).Value) = True Then
                            If row.Cells(5).Value = row1.Cells(5).Value Then
                                Return True
                            Else
                                Return False
                            End If
                        ElseIf IsNumeric(row.Cells(5).Value) = False And IsNumeric(row1.Cells(5).Value) = False Then
                            Return True
                        End If
                    End If
                    j += 1
                End While
                i += 1
            End While

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    ''' <summary>
    ''' menggabungkan 2 baris yang sama dalam dgv menjadi 1 dengan menambahkan qty-nya
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub GabungkanBaris(ByVal vdGV As DataGridView)
        Dim row, row1 As DataGridViewRow
        Dim i, j As Integer

        Try
            While i <= vdGV.RowCount - 2
                row = vdGV.Rows(i)
                j = i + 1
                While j <= vdGV.RowCount - 1
                    row1 = vdGV.Rows(j)
                    '2 baris sama apabila semua data sama kecuali: data qty, unit, sec units, expiry date, 
                    'warehouse
                    If CStr(row.Cells(1).Value) = CStr(row1.Cells(1).Value) And CStr(row.Cells(2).Value) = CStr(row1.Cells(2).Value) And _
                    CStr(row.Cells(3).Value) = CStr(row1.Cells(3).Value) And CStr(row.Cells(12).Value) = CStr(row1.Cells(12).Value) And _
                    CStr(row.Cells(7).Value) = CStr(row1.Cells(7).Value) And _
                    CStr(row.Cells(9).Value) = CStr(row1.Cells(9).Value) Then
                        If IsNumeric(row.Cells(5).Value) = True And IsNumeric(row1.Cells(5).Value) = True Then
                            If CDbl(row.Cells(5).Value) = CDbl(row1.Cells(5).Value) Then
                                Me.GabungkanBaris(row, row1, vdGV)
                                Me.ObjCD.SetNumber(vdGV)
                                Exit Sub
                            End If
                        ElseIf IsNumeric(row.Cells(5).Value) = False And IsNumeric(row1.Cells(5).Value) = False Then
                            Me.GabungkanBaris(row, row1, vdGV)
                            Me.ObjCD.SetNumber(vdGV)
                            Exit Sub
                        End If
                    Else
                        j += 1
                    End If
                End While
                i += 1
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub GabungkanBaris(ByVal row As DataGridViewRow, ByVal row1 As DataGridViewRow, _
    ByVal vDGV As DataGridView)
        Try
            'gabungkan data qty.
            row.Cells(3).Value = CDbl(row.Cells(3).Value) + CDbl(row1.Cells(3).Value)
            'gabungkan data remarks
            If CStr(row.Cells(1).Value).EndsWith(".") = False Then
                'tambahkan titik bila tidak diakhiri dengan titik
                row.Cells(11).Value = CStr(row.Cells(11).Value) + ". "
            End If
            row.Cells(11).Value = CStr(row.Cells(11).Value) + CStr(row1.Cells(11).Value)
            'hapus baris ybs
            vDGV.Rows.RemoveAt(row1.Index)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahEditSahb4(ByVal row As DataGridViewRow, ByVal PrimaryLocation As Double, _
    ByVal PrimaryPdO As Double, ByVal PrimaryProduct As Double, ByVal BatchNo As String, _
    ByVal ExpiryDate As Date, ByVal SecondaryQty As Decimal, ByVal JmlDGV As Decimal, _
    Optional ByVal QtyCheck As Boolean = False) As Boolean
        Dim JmlGudang, JmlObjLama, JmlSetelahEdit, JmlDGVTotal, JmlDGVLain As Decimal
        Dim LocationCont, BatchNoCont, ExpiryDateCont, UseSecondary As Boolean
        Dim ObjSC As New ClsTransaksiStockCard

        Try
            JmlGudang = Me.ObjProduct.GetJml3(PrimaryProduct, SecondaryQty, PrimaryLocation, BatchNo)
            'perhitungan menggunakan nilai jumlah dgv yang baru
            JmlObjLama = Me.objStockBundlingItem.GetJmlB4(Me.TxtPrimaryStockBundling.Text, PrimaryProduct, _
            SecondaryQty, PrimaryLocation, BatchNo, ExpiryDate)
            Me.ObjProduct.GetProductProperties(PrimaryProduct, LocationCont, BatchNoCont, ExpiryDateCont, UseSecondary)
            JmlDGVTotal = Me.GetJmlDGV(PrimaryProduct, SecondaryQty, PrimaryLocation, BatchNo, _
            BatchNoCont, LocationCont, UseSecondary, Me.dgvItemb4)
            JmlSetelahEdit = JmlGudang + JmlObjLama - JmlDGVTotal
            If JmlSetelahEdit < 0 Then
                JmlDGVLain = JmlDGVTotal - JmlDGV
                If QtyCheck = False Then
                    MsgBox("Your data change was cancelled, because it will make you stock quantity minus.", _
                    MsgBoxStyle.Information, "Information")
                Else
                    MsgBox("Maximum value you can enter is " + Me.ObjCD.DigitGrouping(CStr(JmlGudang + JmlObjLama - JmlDGVLain), Me.ObjProduct.GetAngkaDecimal(PrimaryProduct)) + " " + _
                    Me.ObjProduct.GetUnitCode(PrimaryProduct) + ".", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            'cek stock untuk tanggal yang bersangkutan
            JmlGudang = ObjSC.GetStock(PrimaryProduct, Me.dtpDate.Value)
            'perhitungan menggunakan nilai jumlah dgv yang baru
            Me.ObjProduct.GetProductProperties(PrimaryProduct, LocationCont, BatchNoCont, ExpiryDateCont, UseSecondary)
            JmlSetelahEdit = JmlGudang + JmlObjLama - JmlDGVTotal
            If JmlSetelahEdit < 0 Then
                JmlDGVLain = JmlDGVTotal - JmlDGV
                If QtyCheck = False Then
                    MsgBox("Your data change was cancelled, because it will make you stock quantity on " + Format(Me.dtpDate.Value, "dd MMM yyyy") + "  minus.", _
                    MsgBoxStyle.Information, "Information")
                Else
                    MsgBox("Maximum value you can enter " + " on " + Format(Me.dtpDate.Value, "dd MMM yyyy") + " is " + Me.ObjCD.DigitGrouping(CStr(JmlGudang + JmlObjLama - JmlDGVLain), Me.ObjProduct.GetAngkaDecimal(PrimaryProduct)) + " " + _
                    Me.ObjProduct.GetUnitCode(PrimaryProduct) + ".", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            'sampe sini 
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjSC = Nothing
        End Try
    End Function

    Private Sub dgvItemb4_RowsAdded(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItemb4.RowsAdded
        Me.dgvItemb4.TabStop = True
        If Me.txtremarks.Enabled = True Then
            Me.btnUPb4.Enabled = True
            Me.btnDownb4.Enabled = True
        End If
    End Sub

    Private Sub dgvItemb4_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvItemb4.RowsRemoved
        If Me.dgvItemb4.Rows.Count = 0 Then
            Me.dgvItemb4.TabStop = False
            If Me.txtremarks.Enabled = True Then
                Me.btnUPb4.Enabled = False
                Me.btnDownb4.Enabled = False
            End If
        End If
    End Sub

    Private Sub dgvItem_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItemb4.SelectionChanged
        Try
            If Me.txtremarks.Enabled = True Then
                If Me.dgvItemb4.SelectedRows.Count > 0 Then
                    Me.btnUPb4.Enabled = True
                    Me.btnDownb4.Enabled = True
                Else
                    Me.btnUPb4.Enabled = False
                    Me.btnDownb4.Enabled = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemb4_Sorted(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dgvItemb4.Sorted
        Me.ObjCD.SetNumber(Me.dgvItemb4)
    End Sub

    ''' <summary>
    ''' tampilkan row yang diberikan pada parameter
    ''' </summary>
    ''' <param name="row"></param>
    ''' <remarks></remarks>
    Private Sub ShowRow(ByVal row As DataGridViewRow, ByVal Up As Boolean, ByVal vdGV As DataGridView)
        Try
            'tentukan arah perpindahan, ke atas atau ke bawah
            If row.Index < vdGV.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke atas
                While row.Displayed = False
                    vdGV.FirstDisplayedScrollingRowIndex = vdGV.FirstDisplayedScrollingRowIndex - 1
                End While
            ElseIf row.Index > vdGV.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke bawah
                While vdGV.Rows(row.Index).Displayed = False
                    vdGV.FirstDisplayedScrollingRowIndex = vdGV.FirstDisplayedScrollingRowIndex + 1
                End While
                'kalo sampe sini pasti row yang disorot adalah row yang paling bawah, pastikan semuanya terlihat
                If vdGV.DisplayedRowCount(True) <> vdGV.DisplayedRowCount(False) Then
                    vdGV.FirstDisplayedScrollingRowIndex = vdGV.FirstDisplayedScrollingRowIndex + 1
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnUPb4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUPb4.Click
        Try
            If Me.dgvItemb4.SelectedRows.Count > 0 Then
                Me.btnUPb4.Enabled = False
                Me.btnDownb4.Enabled = False
                Dim row As DataGridViewRow = Me.dgvItemb4.SelectedRows(0)
                If row.Index <> 0 Then
                    Dim indexbaru As Integer = row.Index - 1
                    Me.dgvItemb4.Rows.Remove(row)
                    Me.dgvItemb4.Rows.InsertRange(indexbaru, row)
                    Me.dgvItemb4.Focus()
                    Me.dgvItemb4.ClearSelection()
                    row.Selected = True
                    Me.ObjCD.SetNumber(Me.dgvItemb4)
                    If row.Displayed = False Then
                        Me.ShowRow(row, True, Me.dgvItemb4)
                    End If
                End If
                Me.btnUPb4.Enabled = True
                Me.btnDownb4.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDownb4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDownb4.Click
        Try
            If Me.dgvItemb4.SelectedRows.Count > 0 Then
                Me.btnUPb4.Enabled = False
                Me.btnDownb4.Enabled = False
                Dim row As DataGridViewRow = Me.dgvItemb4.SelectedRows(0)
                If row.Index <> Me.dgvItemb4.Rows.Count - 1 Then
                    Dim indexbaru As Integer = row.Index + 1
                    Me.dgvItemb4.Rows.Remove(row)
                    Me.dgvItemb4.Rows.InsertRange(indexbaru, row)
                    Me.dgvItemb4.Focus()
                    Me.dgvItemb4.ClearSelection()
                    row.Selected = True
                    Me.ObjCD.SetNumber(Me.dgvItemb4)
                    Me.ShowRow(row, False, Me.dgvItemb4)
                End If
                Me.btnUPb4.Enabled = True
                Me.btnDownb4.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemb4_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItemb4.UserDeletedRow
        Try
            Me.ObjCD.SetNumber(Me.dgvItemb4)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemb4_UserDeletingRow(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvItemb4.UserDeletingRow
        Try
            'ketika user akan menghapus data, tidak perlu cek apakah akibatnya akan mengakibatkan ada stock yang minus, karena akan mengakibatkan
            'stock menjadi positif
            'tetapi untuk transaksi yang berkaitan dengan transaksi lainnya (DO, Invoice, GRN atau PO) harus cek validitas jumlah berkaitan dengan mereka
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtremarks.GotFocus
        Try
            If Me.ObjStockBundling.TipeBundling = ClsTrans_StockBundling.EnumTipeBundling.Inventory Then
                Me.txtremarks.BackColor = Core_Module.GetSMESSBackColorAll
            Else
                Me.txtremarks.BackColor = Core_Module.GetSMESSBackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtremarks.LostFocus
        Try
            Me.txtremarks.BackColor = Me.TxtPrimaryCodeNob4.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeNob4_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNob4.GotFocus
        Try
            If Me.ObjStockBundling.TipeBundling = ClsTrans_StockBundling.EnumTipeBundling.Inventory Then
                Me.TxtCodeNob4.BackColor = Core_Module.GetSMESSBackColorAll
            Else
                Me.TxtCodeNob4.BackColor = Core_Module.GetSMESSBackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeNob4_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNob4.LostFocus
        Try
            Me.TxtCodeNob4.BackColor = Me.TxtPrimaryCodeNob4.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatchb4_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchb4.GotFocus
        Try
            If Me.ObjStockBundling.TipeBundling = ClsTrans_StockBundling.EnumTipeBundling.Inventory Then
                Me.TxtQtyperBatchb4.BackColor = Core_Module.GetSMESSBackColorAll
            Else
                Me.TxtQtyperBatchb4.BackColor = Core_Module.GetSMESSBackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatchb4_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchb4.LostFocus
        Try
            Me.TxtQtyperBatchb4.BackColor = Me.TxtPrimaryCodeNob4.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSecondaryQtyb4_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSecondaryQtyb4.GotFocus
        Try
            If Me.ObjStockBundling.TipeBundling = ClsTrans_StockBundling.EnumTipeBundling.Inventory Then
                Me.txtSecondaryQtyb4.BackColor = Core_Module.GetSMESSBackColorAll
            Else
                Me.txtSecondaryQtyb4.BackColor = Core_Module.GetSMESSBackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSecondaryQtyb4_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSecondaryQtyb4.LostFocus
        Try
            Me.txtSecondaryQtyb4.BackColor = Me.TxtPrimaryCodeNob4.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarkItemb4_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarkItemb4.GotFocus
        Try
            If Me.ObjStockBundling.TipeBundling = ClsTrans_StockBundling.EnumTipeBundling.Inventory Then
                Me.TxtRemarkItemb4.BackColor = Core_Module.GetSMESSBackColorAll
            Else
                Me.TxtRemarkItemb4.BackColor = Core_Module.GetSMESSBackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarkItemb4_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarkItemb4.LostFocus
        Try
            Me.TxtRemarkItemb4.BackColor = Me.TxtPrimaryCodeNob4.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboBatchNob4_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboBatchNob4.SelectedIndexChanged
        Try
            If Me.SedangAddRM = True Then
                Exit Sub
            End If
            If Me.chkLocationControlledb4.Checked = True Then
                If Me.cboBatchNob4.Text <> "" Then
                    'isikan data lokasi
                    Me.IsiCboLokasi(Me.TxtPrimaryCodeNob4.Text, Me.TxtPrimaryStockBundling.Text, Me.cboBatchNob4.Text)
                End If
            Else
                'kosongkan data lokasi
                Me.cboLocationb4.DataSource = Nothing
                Me.cboLocationb4.Items.Clear()
            End If
            'isikan data expiry date 
            If Me.chkExpiryDateControlledb4.Checked = True Then
                Me.DtpExpiredb4.Value = Me.ObjProduct.GetExpiryDate(Me.TxtPrimaryCodeNob4.Text, Me.cboBatchNob4.Text)
            Else
                If Me.DtpExpiredb4.Visible = True Then
                    Me.DtpExpiredb4.Value = Today
                Else
                    Me.DtpExpiredb4.Value = #1/1/1950#
                End If
            End If
            'isikan secondary unit
            If Me.chkLocationControlledb4.Checked = False And _
            Me.chkUseSecondaryUnitb4.Checked = True And Me.cboBatchNob4.Text <> "" Then
                Me.IsiCboSecondary(Me.TxtPrimaryCodeNob4.Text, Me.TxtPrimaryStockBundling.Text, _
                Me.cboBatchNob4.Text, Me.txtPrimaryLocationb4.Text)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboLocationb4_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboLocationb4.SelectedIndexChanged
        Dim objLoc As New clsMasterStorageArea

        Try
            If Me.SedangAddRM = True Then
                Exit Sub
            End If

            If Me.cboLocationb4.Text = "" Or objLoc.ApakahSudahAda(Me.cboLocationb4.Text) = False Then
                Me.txtPrimaryLocationb4.Text = "-1"
                Me.TxtWarehouseAfter.Clear()
            Else
                Me.txtPrimaryLocationb4.Text = objLoc.GetPrimary(Me.cboLocationb4.Text)
                Me.TxtWarehouseb4.Text = objLoc.GetWarehouseCode(Me.txtPrimaryLocationb4.Text)
                Me.IsiCboSecondary(Me.TxtPrimaryCodeNob4.Text, Me.TxtPrimaryStockBundling.Text, Me.cboBatchNob4.Text, _
                Me.txtPrimaryLocationb4.Text)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            objLoc = Nothing
        End Try
    End Sub

    Private Sub cboSecondaryUnitsb4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboSecondaryUnitsb4.SelectedIndexChanged
        Try
            If Me.cboSecondaryUnitsb4.SelectedIndex <> -1 Then
                Me.txtSecondaryQtyb4.Text = Me.cboSecondaryUnitsb4.Text
            Else
                Me.txtSecondaryQtyb4.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeNoAfter_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCodeNoAfter.KeyDown
        Dim i As New frmCariMasterProduct2
        Dim Jml As Long = 0
        Dim PrimaryProduct As Double = -1

        Try
            If e.KeyValue = 13 Then
                Jml = Me.ObjProduct.GetJml(Me.TxtCodeNoAfter.Text)
                PrimaryProduct = Me.ObjProduct.GetPrimary(Me.TxtCodeNoAfter.Text)

                If Jml = 1 And Me.ObjProduct.ApakahGoods(PrimaryProduct) = True Then
                    Me.TxtPrimaryCodeNoAfter.Text = Me.ObjProduct.GetPrimary(Me.TxtCodeNoAfter.Text)
                    Me.txtPrimaryCategoryAfter.Text = Me.ObjProduct.GetPrimaryCategory(Me.TxtPrimaryCodeNoAfter.Text)
                    Me.txtDescAfter.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNoAfter.Text)
                    Me.SetGboItemAfter()
                    If Me.txtBatchNoAfter.Visible = True Then
                        Me.txtBatchNoAfter.Focus()
                    ElseIf Me.TxtLocationAfter.Visible = True Then
                        Me.TxtLocationAfter.Focus()
                    Else
                        Me.TxtQtyperBatchAfter.Focus()
                    End If
                ElseIf Jml > 1 Or Jml = 0 Then
                    'ada 2 product dengan kode yang sama atau tidak ada product yang dimaksud
                    If Me.TxtCodeNoAfter.Text <> "" Then
                        i.txtCari.Text = Me.TxtCodeNoAfter.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        Me.TxtPrimaryCodeNoAfter.Text = i.PrimaryProduct
                        Me.TxtCodeNoAfter.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryCodeNoAfter.Text)
                        Me.txtPrimaryCategoryAfter.Text = Me.ObjProduct.GetPrimaryCategory(Me.TxtPrimaryCodeNoAfter.Text)
                        Me.txtDescAfter.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNoAfter.Text)
                        Me.SetGboItemAfter()
                        If Me.txtBatchNoAfter.Visible = True Then
                            Me.txtBatchNoAfter.Focus()
                        ElseIf Me.TxtLocationAfter.Visible = True Then
                            Me.TxtLocationAfter.Focus()
                        Else
                            Me.TxtQtyperBatchAfter.Focus()
                        End If
                    Else
                        Me.TxtCodeNoAfter.Clear()
                        Me.TxtPrimaryCodeNoAfter.Text = "-1"
                        Me.txtPrimaryCategoryAfter.Text = "-1"
                        Me.KosongkanItemAfter()
                        Me.TxtCodeNoAfter.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtCodeNoAfter_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNoAfter.Leave
        Dim i As New frmCariMasterProduct2
        Dim PrimaryProduct As Double = -1

        Try
            If Me.TxtCodeNoAfter.Text <> "" Then
                PrimaryProduct = Me.ObjProduct.GetPrimary(Me.TxtCodeNoAfter.Text)
                If Me.ObjProduct.GetJml(Me.TxtCodeNoAfter.Text) = 1 And Me.ObjProduct.ApakahGoods(PrimaryProduct) Then
                    Me.TxtPrimaryCodeNoAfter.Text = Me.ObjProduct.GetPrimary(Me.TxtCodeNoAfter.Text)
                    Me.txtPrimaryCategoryAfter.Text = Me.ObjProduct.GetPrimaryCategory(Me.TxtPrimaryCodeNoAfter.Text)
                    Me.txtDescAfter.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNoAfter.Text)
                    Me.SetGboItemAfter()
                    If Me.txtBatchNoAfter.Visible = True Then
                        Me.txtBatchNoAfter.Focus()
                    ElseIf Me.TxtLocationAfter.Visible = True Then
                        Me.TxtLocationAfter.Focus()
                    Else
                        Me.TxtQtyperBatchAfter.Focus()
                    End If
                    If Me.TxtCodeNoAfter.Text = isProductCode Then
                        If Me.ApakahCodeNoDariWOSudahAda() = False Then
                            Me.TxtQtyperBatchAfter.Text = isBalanceProduce
                        End If
                    End If
                Else
                    If Me.TxtCodeNoAfter.Text <> "" Then
                        i.txtCari.Text = Me.TxtCodeNoAfter.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        Me.TxtPrimaryCodeNoAfter.Text = i.PrimaryProduct
                        Me.TxtCodeNoAfter.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryCodeNoAfter.Text)
                        Me.txtPrimaryCategoryAfter.Text = Me.ObjProduct.GetPrimaryCategory(Me.TxtPrimaryCodeNoAfter.Text)
                        Me.txtDescAfter.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNoAfter.Text)
                        Me.SetGboItemAfter()
                        If Me.txtBatchNoAfter.Visible = True Then
                            Me.txtBatchNoAfter.Focus()
                        ElseIf Me.TxtLocationAfter.Visible = True Then
                            Me.TxtLocationAfter.Focus()
                        Else
                            Me.TxtQtyperBatchAfter.Focus()
                        End If
                        If Me.TxtCodeNoAfter.Text = isProductCode Then
                            If Me.ApakahCodeNoDariWOSudahAda() = False Then
                                Me.TxtQtyperBatchAfter.Text = isBalanceProduce
                            End If
                        End If
                    Else
                        Me.TxtCodeNoAfter.Clear()
                        Me.TxtPrimaryCodeNoAfter.Text = "-1"
                        Me.txtPrimaryCategoryAfter.Text = "-1"
                        Me.KosongkanItemAfter()
                        Me.TxtCodeNoAfter.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryCodeNoAfter.Text = "-1"
                Me.txtDescAfter.Clear()
                Me.txtPrimaryCategoryAfter.Text = "-1"
                Me.KosongkanItemAfter()
            End If
            Me.SetGboItemAfter()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindProductAfter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindCodeNoAfter.Click
        Dim i As New frmCariMasterProduct2

        Try
            i.cboCriteria.SelectedIndex = 0
            If Me.TxtCodeNoAfter.Text <> "" Then
                i.txtCari.Text = Me.TxtCodeNoAfter.Text
            End If

            i.ShowDialog()
            If i.PrimaryProduct > 0 Then
                Me.TxtPrimaryCodeNoAfter.Text = i.PrimaryProduct
                Me.TxtCodeNoAfter.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryCodeNoAfter.Text)
                Me.txtDescAfter.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNoAfter.Text)
                Me.SetGboItemAfter()
                If Me.txtBatchNoAfter.Visible = True Then
                    Me.txtBatchNoAfter.Focus()
                ElseIf Me.TxtLocationAfter.Visible = True Then
                    Me.TxtLocationAfter.Focus()
                Else
                    Me.TxtQtyperBatchAfter.Focus()
                End If
            Else
                Me.TxtCodeNoAfter.Clear()
                Me.TxtPrimaryCodeNoAfter.Text = "-1"
                Me.txtPrimaryCategoryAfter.Text = "-1"
                Me.KosongkanItemAfter()
                Me.TxtCodeNoAfter.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub BtnInputAfter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnInputAfter.Click
        Dim OK As Boolean = False
        Try
            If Me.ApakahDataItemAfterBenar = True Then
                If Me.txtprimaryproduct_dariwo.Text = Me.TxtPrimaryCodeNoAfter.Text Then
                    If ApakahCodeNoDariWOSudahAda() = False Then
                        OK = True
                    End If
                End If

                Dim Index As Integer = Me.GetIndex(Me.TxtPrimaryCodeNoAfter.Text, Me.txtPrimaryLocationAfter.Text, Me.txtBatchNoAfter.Text, _
                 Me.txtSecondaryQtyAfter.Text, Me.txtPrimaryPDOAfter.Text, Me.DtpExpiredAfter.Value, Me.dgvItemAfter)
                'cari data yang sama persis
                If Index <> -1 Then
                    'tambahkan nilai qty
                    Dim row As DataGridViewRow = Me.dgvItemAfter.Rows(Index)
                    row.Cells(3).Value = CDbl(row.Cells(3).Value) + CDbl(Me.TxtQtyperBatchAfter.Text)
                Else
                    'tambahkan baris baru
                    Dim ExpiredDateStr As String
                    If Me.DtpExpiredAfter.Value.Date = #1/1/1950# Then
                        ExpiredDateStr = "-"
                    Else
                        ExpiredDateStr = Format(Me.DtpExpiredAfter.Value.Date, "dd MMM yyyy")
                    End If
                    Me.dgvItemAfter.Rows.Add("", Me.TxtCodeNoAfter.Text, Me.txtDescAfter.Text, Me.ObjCD.DigitGrouping(Me.TxtQtyperBatchAfter.Text, Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNoAfter.Text)), _
                    Me.lblUnitsAfter.Text, Me.txtSecondaryQtyAfter.Text, Me.lblSecondaryUnitsAfter.Text, Me.txtBatchNoAfter.Text, _
                    ExpiredDateStr, Me.TxtLocationAfter.Text, Me.TxtWarehouseAfter.Text, Me.TxtRemarkItemAfter.Text, _
                    Me.TxtPrimaryCodeNoAfter.Text, Me.txtPrimaryLocationAfter.Text, Me.txtPrimaryUnitsAfter.Text, _
                    Me.txtPrimarySecondaryUnitsAfter.Text, Me.txtFaktorPemberat.Text)
                    'set ulang nomor
                    Me.ObjCD.SetNumber(Me.dgvItemAfter)
                End If

                'If Me.Apakah_Worktime(Me.dtpplan_start.Value) = True And Me.Apakah_Worktime(Me.dtpplan_end.Value) = True Then
                '    Me.txtactual_workhour.Text = Me.ObjCD.Convert_ToString(GetWorkDuration())
                '    Me.BufferWorkHour = Me.txtactual_workhour.Text
                'End If

                ''isi rm dari bom
                Me.SetGboItemAfter()
                If OK = True Then
                    'Me.IsikanRawMaterial()
                End If

                'kosongkan gbo item
                Me.KosongkanItemAfter()
                Me.TxtCodeNoAfter.Focus()
                'format dgv
                Me.FormatDGV(Me.dgvItemAfter)
                'fokuskan ke baris terakhir
                If OK = False Then
                    While Me.dgvItemAfter.Rows(Me.dgvItemAfter.Rows.Count - 1).Displayed = False _
                               Or Me.dgvItemAfter.DisplayedRowCount(True) <> Me.dgvItemAfter.DisplayedRowCount(False)
                        Me.dgvItemAfter.FirstDisplayedScrollingRowIndex = Me.dgvItemAfter.FirstDisplayedScrollingRowIndex + 1
                    End While
                End If
            End If

            Me.AturWoNumberdanCboWorkUnit()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatchAfter_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchAfter.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtQtyperBatchAfter, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatchAfter_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyperBatchAfter.KeyDown
        Me.ObjCD.Number_KeyDown(Me.TxtQtyperBatchAfter, sender, e)
    End Sub

    Private Sub TxtQtyperBatchAfter_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyperBatchAfter.KeyUp
        Me.ObjCD.Number_KeyUp(Me.TxtQtyperBatchAfter, sender, e)
    End Sub

    Private Sub txtSecondaryQtyAfter_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSecondaryQtyAfter.KeyDown
        Me.ObjCD.Number_KeyDown(Me.txtSecondaryQtyAfter, sender, e)
    End Sub

    Private Sub txtSecondaryQtyAfter_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSecondaryQtyAfter.KeyUp
        Me.ObjCD.Number_KeyUp(Me.txtSecondaryQtyAfter, sender, e)
    End Sub

    Private Sub dgvItemAfter_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItemAfter.CellBeginEdit
        Try
            'yang bisa diedit adalah:
            '1. Qty
            '2. Secondary Qty
            '3. Batch No
            '4. Expiry Date
            '5. Location Code
            '6. Remarks
            '7. Cost Factor/Faktor Pemberat

            If Me.txtremarks.Enabled = True Then
                If e.ColumnIndex = 3 Or e.ColumnIndex = 5 Or e.ColumnIndex = 7 Or e.ColumnIndex = 9 _
                Or e.ColumnIndex = 11 Or e.ColumnIndex = 8 Or e.ColumnIndex = 16 Then
                    'untuk batch no, secondary qty, location, expiry data -> hanya bisa diedit jika memang 
                    'masing-masing parameter yang akan diedit dikontrol
                    If e.ColumnIndex = 5 Then
                        'cek secondary qty
                        If CStr(Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value) = "-" Then
                            e.Cancel = True
                            Exit Sub
                        End If
                    ElseIf e.ColumnIndex = 7 Then
                        'cek batch no
                        If Me.ObjProduct.ApakahBatchNoControlled(Me.dgvItemAfter.Item(12, e.RowIndex).Value) = False Then
                            e.Cancel = True
                            Exit Sub
                        End If
                    ElseIf e.ColumnIndex = 8 Then
                        'cek expiry date
                        If Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = "-" Then
                            e.Cancel = True
                            Exit Sub
                        End If
                    ElseIf e.ColumnIndex = 9 Then
                        'cek location controlled
                        If Me.ObjProduct.ApakahLocationControlled(Me.dgvItemAfter.Item(12, e.RowIndex).Value) = False Then
                            e.Cancel = True
                            Exit Sub
                        End If
                    End If
                    'kalo sampe sini maka boleh edit
                    Me.buffer = Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                'hanya bisa edit pada saat detail aktif
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemAfter_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItemAfter.CellEndEdit
        Try
            Dim SecondaryQty, JmlDGV As Decimal
            Dim BatchNo As String = ""
            Dim ExpiryDate As Date = #1/1/1950#
            Dim row As DataGridViewRow = Me.dgvItemAfter.Rows(e.RowIndex)

            'pastikan ada isinya
            With row
                If IsNothing(Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.DataGridViewTextBoxColumn12.Index Then
                        If IsNothing(Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value) Then
                            Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                        Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.Replace("'", "").Replace("%", "").Replace("*", "").Trim
                    End If

                    'Bila tidak ada perubahan data maka tidak perlu dilakukan cek
                    If e.ColumnIndex <> 8 Then
                        If Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer Then
                            Exit Sub
                        End If
                    Else
                        If IsDate(Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            If CDate(Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value).Date = CDate(Me.buffer).Date Then
                                Exit Sub
                            End If
                        End If
                    End If

                    If e.ColumnIndex <> 11 Then
                        'kalo bukan data remarks maka perlu dicek kebenaran data hasil edit
                        If IsDate(.Cells(8).Value) = True Then
                            ExpiryDate = CDate(.Cells(8).Value)
                        Else
                            ExpiryDate = #1/1/1950#
                        End If
                        If IsNumeric(.Cells(5).Value) = True Then
                            SecondaryQty = .Cells(5).Value
                        Else
                            SecondaryQty = -1
                        End If
                        JmlDGV = 0

                        If e.ColumnIndex = 3 Then
                            Me.EndEditQtyAfter(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                            'Me.HitungTotalProcess()
                        ElseIf e.ColumnIndex = 5 Then
                            Me.EndEditSecondaryUnitAfter(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                        ElseIf e.ColumnIndex = 7 Then
                            Me.EndEditBatchNoAfter(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                        ElseIf e.ColumnIndex = 9 Then
                            Me.EndEditLocationAfter(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                        ElseIf e.ColumnIndex = 8 Then
                            Me.EndEditExpiryDateAfter(sender, e, SecondaryQty, JmlDGV, BatchNo, ExpiryDate, row)
                        ElseIf e.ColumnIndex = 16 Then
                            Me.EndEditCostFactor(sender, e, row)
                        End If
                    End If
                Else
                    If e.ColumnIndex = 13 Then
                        Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Else
                        Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditExpiryDateAfter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal SecondaryQty As Decimal, ByVal JmlDGV As Decimal, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
    ByVal row As DataGridViewRow)
        Try
            With row
                'pastikan format yang dituliskan adalah format tanggal
                If IsDate(Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
                'format ulang tanggal sehingga formatnya konsisten
                Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Format(CDate(Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value), _
                "dd MMM yyyy")
                'cek apakah back date terhadap tanggal transaksi, konfirmasikan data ini
                If CDate(Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value) < Me.dtpDate.Value.Date Then
                    If MsgBox("This product expiry date is smaller than transaction date. Do you want to continue?", _
                    MsgBoxStyle.YesNo, "Expired Product") = MsgBoxResult.No Then
                        Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                End If
                'pastikan tanggal expiry date tidak konflik dengan data yang sudah ada
                Dim ExpiryDateLama As Date = Me.ObjProduct.GetExpiryDate(Me.dgvItemAfter.Item(12, e.RowIndex).Value, _
                Me.dgvItemAfter.Item(7, e.RowIndex).Value)
                If ExpiryDateLama <> #1/1/1950# Then
                    If ExpiryDateLama.Date <> CDate(Me.dgvItemAfter(e.ColumnIndex, e.RowIndex).Value) Then
                        MsgBox("Expiry date for batch no: " + CStr(Me.dgvItemAfter.Item(7, e.RowIndex).Value) + " is " + _
                        Format(ExpiryDateLama, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                        Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                End If
                'bandingkan dengan yang sudah ada didalam dgv
                Dim row1 As DataGridViewRow
                For Each row1 In Me.dgvItemAfter.Rows
                    If row1.Cells(7).Value = row.Cells(7).Value And CStr(row.Cells(7).Value) <> "" _
                    And row1.Index <> row.Index And row1.Cells(12).Value = row.Cells(12).Value Then
                        If CDate(row1.Cells(8).Value).Date <> CDate(row.Cells(8).Value).Date Then
                            MsgBox("You already entered expiry date for batch no: " + CStr(row1.Cells(7).Value) + " on " + _
                            Format(CDate(row1.Cells(8).Value), "dd MMM yyyy"), MsgBoxStyle.Information, _
                            "Information")
                            Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                    End If
                Next
                'bandingkan dengan yang sudah ada didalam transaksi lain. Detail lihat point no 8.c pada 
                'fungsi apakah data item benar
                'fungsi apakah data item benar
                '1. stock in (semua stock in)
                Dim JenisTransaksi As clsTransaksiStockInOut.EnumTipeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase
                For loopInOut As Integer = 0 To 4 Step 1
                    Select Case loopInOut
                        Case 0
                            '   a. stock in purchase
                            JenisTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase
                            Exit Select
                        Case 1
                            '   b. stock in sales return
                            JenisTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn
                            Exit Select
                        Case 2
                            '   c. stock in adjust
                            JenisTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust
                            Exit Select
                        Case 3
                            '   d. stock in back
                            JenisTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_BackProcess
                            Exit Select
                        Case 4
                            '   e. stock in proses
                            JenisTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Process
                            Exit Select
                    End Select
                    ExpiryDateLama = Me.ObjItemInOut.GetExpiryDate(Me.dgvItemAfter.Item(12, e.RowIndex).Value, _
                    Me.dgvItemAfter.Item(7, e.RowIndex).Value, JenisTransaksi)
                    If ExpiryDateLama <> #1/1/1950# Then
                        If ExpiryDateLama.Date <> CDate(Me.dgvItemAfter(e.ColumnIndex, e.RowIndex).Value) Then
                            MsgBox("Expiry date for batch no: " + CStr(Me.dgvItemAfter.Item(7, e.RowIndex).Value) + " is " + _
                            Format(ExpiryDateLama, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                            Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                    End If
                Next
                '2. transaksi stock bundling untuk item hasil bundling
                ExpiryDateLama = Me.ObjItemBundling.GetExpiryDateAfter(Me.dgvItemAfter.Item(12, e.RowIndex).Value, _
                Me.dgvItemAfter.Item(7, e.RowIndex).Value)
                If ExpiryDateLama <> #1/1/1950# Then
                    If ExpiryDateLama.Date <> CDate(Me.dgvItemAfter(e.ColumnIndex, e.RowIndex).Value) Then
                        MsgBox("Expiry date for batch no: " + CStr(Me.dgvItemAfter.Item(7, e.RowIndex).Value) + " is " + _
                        Format(ExpiryDateLama, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                        Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                End If
                '3. proses stock taking
                'msgbox("belum")
                'expiry date, tidak perlu cek kemungkinan merger, karena tidak mungkin
                'hanya expiry date saja yang berbeda sedangkan batch no-nya sama
                If Me.ApakahEditSahAfter(row, .Cells(13).Value, -1, .Cells(12).Value, _
                .Cells(7).Value, CDate(Me.buffer), SecondaryQty, JmlDGV) = False Then
                    Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditCostFactor(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal row As DataGridViewRow)
        Try
            With row
                'untuk nilai qty dan secondary qty, pastikan nilainya positif
                If Me.ObjCD.CekBilanganPositif(Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("Please enter positive number for cost factor data.", MsgBoxStyle.Information, _
                    "Information")
                    Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                End If
                'digit grouping
                Me.dgvItemAfter(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItemAfter(e.ColumnIndex, e.RowIndex).Value, 2)
                'cek nilai cost factor maximum adalah 100
                If CDbl(Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value) > 100 Then
                    MsgBox("The maximum value is 100%.", MsgBoxStyle.Information, _
                    "Information")
                    Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditQtyAfter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal SecondaryQty As Decimal, ByVal JmlDGV As Decimal, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
    ByVal row As DataGridViewRow)
        Try
            With row
                'digit grouping
                Me.dgvItemAfter(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItemAfter(e.ColumnIndex, e.RowIndex).Value, _
                Me.ObjProduct.GetAngkaDecimal(Me.dgvItemAfter.Item(Me.ColPrimaryCodeNoAfter.Index, e.RowIndex).Value))

                'untuk nilai qty dan secondary qty, pastikan nilainya positif
                If Me.ObjCD.CekBilanganPositif(Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("Please enter positive number for quantity data.", MsgBoxStyle.Information, _
                    "Information")
                    Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                End If

                'cek apakah jumlah digudang akan menjadi minus setelah perubahan data
                JmlDGV = .Cells(3).Value
                If Me.ApakahEditSahAfter(row, .Cells(13).Value, -1, .Cells(12).Value, _
                .Cells(7).Value, ExpiryDate, SecondaryQty, JmlDGV, True) = False Then
                    Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    'Else
                    '    Me.dgvItemAfter.Item(Me.ColQtyAfter.Index, e.RowIndex).Value)
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditSecondaryUnitAfter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal SecondaryQty As Decimal, ByVal JmlDGV As Decimal, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
    ByVal row As DataGridViewRow)
        Try
            With row
                'untuk nilai qty dan secondary qty, pastikan nilainya positif
                If Me.ObjCD.CekBilanganPositif(Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("Please enter positive number for @ data.", MsgBoxStyle.Information, _
                    "Information")
                    Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
                'cek apakah jumlah digudang akan menjadi minus setelah perubahan data
                SecondaryQty = Me.buffer
                If Me.ApakahEditSahAfter(row, .Cells(13).Value, -1, .Cells(12).Value, _
                .Cells(7).Value, ExpiryDate, SecondaryQty, JmlDGV) = False Then
                    Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                Else
                    'cek apakah ada kemungkinan merger
                    Dim AdaSama As Boolean = Me.ApakahAdaBarisYangSama(Me.dgvItemAfter)
                    If AdaSama = True Then
                        If MsgBox("Your data change will merge 2 rows into 1. Do you still want to continue?", MsgBoxStyle.YesNo, _
                        "Confirm") = MsgBoxResult.No Then
                            Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                        Me.GabungkanBaris(Me.dgvItemAfter)
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditBatchNoAfter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal SecondaryQty As Decimal, ByVal JmlDGV As Decimal, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
    ByVal row As DataGridViewRow)
        Try
            With row
                'pastikan diisikan
                If CStr(Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value) = "" Then
                    Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
                'batch no selalu huruf besar
                Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = CStr(Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value).ToUpper
                'batch no
                If Me.ApakahEditSahAfter(row, .Cells(13).Value, -1, .Cells(12).Value, _
                Me.buffer, ExpiryDate, SecondaryQty, JmlDGV) = False Then
                    Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                Else
                    'khusus bila product yang bersangkutan expiry date controlled juga maka
                    'cek ke dalam basis data:
                    Dim BufferExpiryDate As String = "-"
                    '1. data tanggal kadaluarsanya berbeda dengan yang sudah dicatatkan kedalam basis data dan datagridview
                    If IsDate(Me.dgvItemAfter.Item(8, e.RowIndex).Value) = True Then
                        'basis data
                        Dim ExpiryDateLama As Date = Me.ObjProduct.GetExpiryDate(Me.dgvItemAfter.Item(12, e.RowIndex).Value, _
                        Me.dgvItemAfter.Item(7, e.RowIndex).Value)
                        If ExpiryDateLama <> #1/1/1950# Then
                            If ExpiryDateLama.Date <> CDate(Me.dgvItemAfter(8, e.RowIndex).Value) Then
                                If MsgBox("Expiry date for batch no: " + CStr(Me.dgvItemAfter.Item(7, e.RowIndex).Value) + " is " + _
                                Format(ExpiryDateLama, "dd MMM yyyy") + _
                                ". Changing your batch no will also change your expiry date data. Continue?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
                                    Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    Exit Sub
                                Else
                                    'update data expiry date juga
                                    BufferExpiryDate = CDate(Me.dgvItemAfter.Item(8, e.RowIndex).Value)
                                    Me.dgvItemAfter.Item(8, e.RowIndex).Value = Format(ExpiryDateLama, "dd MMM yyyy")
                                End If
                            End If
                        End If

                        'cek juga ke transaksi:
                        '1. stock in (semua stock in)
                        Dim JenisTransaksi As clsTransaksiStockInOut.EnumTipeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase
                        For loopInOut As Integer = 0 To 4 Step 1
                            Select Case loopInOut
                                Case 0
                                    '   a. stock in purchase
                                    JenisTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Purchase
                                    Exit Select
                                Case 1
                                    '   b. stock in sales return
                                    JenisTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn
                                    Exit Select
                                Case 2
                                    '   c. stock in adjust
                                    JenisTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Adjust
                                    Exit Select
                                Case 3
                                    '   d. stock in back
                                    JenisTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_BackProcess
                                    Exit Select
                                Case 4
                                    '   e. stock in proses
                                    JenisTransaksi = clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_Process
                                    Exit Select
                            End Select
                            ExpiryDateLama = Me.ObjItemInOut.GetExpiryDate(Me.dgvItemAfter.Item(12, e.RowIndex).Value, _
                            Me.dgvItemAfter.Item(7, e.RowIndex).Value, JenisTransaksi)
                            If ExpiryDateLama <> #1/1/1950# Then
                                If ExpiryDateLama.Date <> CDate(Me.dgvItemAfter(8, e.RowIndex).Value) Then
                                    If MsgBox("Expiry date for batch no: " + CStr(Me.dgvItemAfter.Item(7, e.RowIndex).Value) + " is " + _
                                    Format(ExpiryDateLama, "dd MMM yyyy") + _
                                    ". Changing your batch no will also change your expiry date data. Continue?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
                                        Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        Exit Sub
                                    Else
                                        'update data expiry date juga
                                        BufferExpiryDate = CDate(Me.dgvItemAfter(8, e.RowIndex).Value)
                                        Me.dgvItemAfter(8, e.RowIndex).Value = Format(ExpiryDateLama, "dd MMM yyyy")
                                    End If
                                End If
                            End If
                        Next
                        '2. transaksi stock bundling untuk item hasil bundling
                        ExpiryDateLama = Me.ObjItemBundling.GetExpiryDateAfter(Me.dgvItemAfter.Item(12, e.RowIndex).Value, _
                        Me.dgvItemAfter.Item(7, e.RowIndex).Value)
                        If ExpiryDateLama <> #1/1/1950# Then
                            If ExpiryDateLama.Date <> CDate(Me.dgvItemAfter(8, e.RowIndex).Value) Then
                                If MsgBox("Expiry date for batch no: " + CStr(Me.dgvItemAfter.Item(7, e.RowIndex).Value) + " is " + _
                                Format(ExpiryDateLama, "dd MMM yyyy") + _
                                ". Changing your batch no will also change your expiry date data. Continue?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
                                    Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    Exit Sub
                                Else
                                    'update data expiry date juga
                                    BufferExpiryDate = CDate(Me.dgvItemAfter(8, e.RowIndex).Value)
                                    Me.dgvItemAfter(8, e.RowIndex).Value = Format(ExpiryDateLama, "dd MMM yyyy")
                                End If
                            End If
                        End If
                        '3. proses stock taking
                        'msgbox("belum")

                        'data grid view
                        Dim row1 As DataGridViewRow
                        For Each row1 In Me.dgvItemAfter.Rows
                            If row1.Cells(7).Value = Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value _
                            And row1.Index <> e.RowIndex And row1.Cells(12).Value = Me.dgvItemAfter.Item(12, e.RowIndex).Value Then
                                If CDate(row1.Cells(8).Value) <> CDate(Me.dgvItemAfter.Item(8, e.RowIndex).Value) Then
                                    If MsgBox("Expiry date for batch no: " + CStr(Me.dgvItemAfter.Item(7, e.RowIndex).Value) + " is " + _
                                        Format(CDate(row1.Cells(8).Value), "dd MMM yyyy") + _
                                        ". Changing your batch no will also change your expiry date data. Continue?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                                        'update data expiry date juga
                                        BufferExpiryDate = CDate(Me.dgvItemAfter.Item(8, e.RowIndex).Value)
                                        Me.dgvItemAfter.Item(8, e.RowIndex).Value = row1.Cells(8).Value
                                        Exit For
                                    Else
                                        Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                        Exit Sub
                                    End If
                                End If
                            End If
                        Next
                    End If
                    '3. cek apakah ada kemungkinan merger
                    Dim AdaSama As Boolean = Me.ApakahAdaBarisYangSama(Me.dgvItemAfter)
                    If AdaSama = True Then
                        If MsgBox("Your data change will merge 2 rows into 1. Do you still want to continue?", MsgBoxStyle.YesNo, _
                        "Confirm") = MsgBoxResult.No Then
                            'kembalikan data batch no dan expiry date
                            Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Me.dgvItemAfter.Item(8, e.RowIndex).Value = Format(CDate(BufferExpiryDate), "dd MMM yyyy")
                            Exit Sub
                        End If
                        Me.GabungkanBaris(Me.dgvItemAfter)
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditLocationAfter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal SecondaryQty As Decimal, ByVal JmlDGV As Decimal, ByVal BatchNo As String, ByVal ExpiryDate As Date, _
    ByVal row As DataGridViewRow)
        Try
            With row
                'pastikan kode gudang adalah huruf kapital
                Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = CStr(Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value).ToUpper
                'cek apakah kode lokasi yang dimasukkan sah
                Dim Obj As New clsMasterStorageArea
                If Obj.ApakahSudahAda(Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    Dim i As New FrmCariMasterStorageArea
                    If CStr(Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value) <> "" Then
                        i.txtCari.Text = Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value
                    End If

                    i.ShowDialog()
                    If i.PrimaryStorageArea <> -1 Then
                        Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = CStr(Obj.GetStorageAreaCode(i.PrimaryStorageArea)).ToUpper
                    Else
                        Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        i = Nothing
                        Exit Sub
                    End If
                    i = Nothing
                Else
                    Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = CStr(Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value).ToUpper
                End If
                Obj = Nothing
                'cek jangan sampai stock digudang menjadi negatif
                If Me.ApakahEditSahAfter(row, .Cells(13).Value, -1, .Cells(12).Value, _
                .Cells(7).Value, ExpiryDate, SecondaryQty, JmlDGV) = False Then
                    Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                Else
                    'update data primary location
                    Dim Obj1 As New clsMasterStorageArea
                    .Cells(13).Value = Obj1.GetPrimary(.Cells(9).Value)
                    'update data gudang
                    .Cells(10).Value = Obj1.GetWarehouseCode(.Cells(13).Value)

                    'cek apakah ada kemungkinan merger
                    Dim AdaSama As Boolean = Me.ApakahAdaBarisYangSama(Me.dgvItemAfter)
                    If AdaSama = True Then
                        If MsgBox("Your data change will merge 2 rows into 1. Do you still want to continue?", MsgBoxStyle.YesNo, _
                        "Confirm") = MsgBoxResult.No Then
                            Me.dgvItemAfter.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            'update primary location
                            .Cells(13).Value = Obj1.GetPrimary(.Cells(9).Value)
                            'update data gudang
                            .Cells(10).Value = Obj1.GetWarehouseCode(.Cells(13).Value)
                            Obj1 = Nothing
                            Exit Sub
                        End If
                        Me.GabungkanBaris(Me.dgvItemAfter)
                    End If
                    Obj1 = Nothing
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahEditSahAfter(ByVal row As DataGridViewRow, ByVal PrimaryLocation As Double, _
    ByVal PrimaryPdO As Double, ByVal PrimaryProduct As Double, ByVal BatchNo As String, _
    ByVal ExpiryDate As Date, ByVal SecondaryQty As Decimal, ByVal JmlDGV As Decimal, Optional ByVal QtyCheck As Boolean = False) As Boolean
        Dim JmlGudang, JmlObjLama, JmlSetelahHapus, JmlDGVTotal As Decimal
        Dim LocationCont, BatchNoCont, ExpiryDateCont, UseSecondary As Boolean

        Try
            JmlGudang = Me.ObjProduct.GetJml3(PrimaryProduct, SecondaryQty, PrimaryLocation, BatchNo)
            'perhitungan menggunakan nilai jumlah dgv yang baru
            JmlObjLama = Me.objStockBundlingItem.GetJmlAfter(Me.TxtPrimaryStockBundling.Text, PrimaryProduct, _
            SecondaryQty, PrimaryLocation, BatchNo, ExpiryDate)
            Me.ObjProduct.GetProductProperties(PrimaryProduct, LocationCont, BatchNoCont, ExpiryDateCont, UseSecondary)
            JmlDGVTotal = Me.GetJmlDGV(PrimaryProduct, SecondaryQty, PrimaryLocation, BatchNo, _
            BatchNoCont, LocationCont, UseSecondary, Me.dgvItemAfter)
            JmlSetelahHapus = JmlGudang - JmlObjLama + JmlDGV
            If JmlSetelahHapus < 0 Then
                If QtyCheck = False Then
                    MsgBox("Your data change was cancelled, because it will make you stock quantity minus.", _
                    MsgBoxStyle.Information, "Information")
                Else
                    MsgBox("Minimum value you must can enter is " + Me.ObjCD.DigitGrouping(CStr(JmlObjLama - JmlGudang), Me.ObjProduct.GetAngkaDecimal(PrimaryProduct)) + _
                    " " + Me.ObjProduct.GetUnitCode(PrimaryProduct) + ".", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub dgvItemAfter_RowsAdded(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItemAfter.RowsAdded
        Try
            Me.dgvItemAfter.TabStop = True
            If Me.txtremarks.Enabled = True Then
                Me.btnUpAfter.Enabled = True
                Me.btnDownAfter.Enabled = True
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvItemAfter.RowsRemoved
        Try
            If Me.dgvItemAfter.Rows.Count = 0 Then
                Me.dgvItemAfter.TabStop = False
                If Me.txtremarks.Enabled = True Then
                    Me.btnUpAfter.Enabled = False
                    Me.btnDownAfter.Enabled = False
                End If
            End If

            If ApakahCodeNoDariWOSudahAda() = False Then
                Me.KosongkanItemB4()
                Me.dgvItemb4.Rows.Clear()
                Me.KosongkanItemProcess()
            End If
            Me.AturWoNumberdanCboWorkUnit()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLocationAfter_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtLocationAfter.KeyDown
        Dim i As New FrmCariMasterStorageArea
        Dim Obj As New clsMasterStorageArea

        Try
            If e.KeyValue = 13 Then
                If Obj.ApakahSudahAda(Me.TxtLocationAfter.Text) = True Then
                    Me.txtPrimaryLocationAfter.Text = Obj.GetPrimary(Me.TxtLocationAfter.Text)
                    Me.TxtWarehouseAfter.Text = Obj.GetWarehouseCode(Me.txtPrimaryLocationAfter.Text)
                    Me.TxtQtyperBatchAfter.Focus()
                Else
                    If Me.TxtLocationAfter.Text <> "" Then
                        i.txtCari.Text = Me.TxtLocationAfter.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryStorageArea > 0 Then
                        Me.txtPrimaryLocationAfter.Text = i.PrimaryStorageArea
                        Me.TxtWarehouseAfter.Text = Obj.GetWarehouseCode(Me.txtPrimaryLocationAfter.Text)
                        Me.TxtLocationAfter.Text = Obj.GetStorageAreaCode(Me.txtPrimaryLocationAfter.Text)
                        Me.TxtQtyperBatchAfter.Focus()
                    Else
                        Me.TxtLocationAfter.Clear()
                        Me.txtPrimaryLocationAfter.Text = "-1"
                        Me.TxtWarehouseAfter.Clear()
                        Me.TxtLocationAfter.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            Obj = Nothing
        End Try
    End Sub

    Private Sub TxtLocationAfter_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLocationAfter.Leave
        Dim i As New FrmCariMasterStorageArea
        Dim obj As New clsMasterStorageArea

        Try
            If Me.TxtLocationAfter.Text <> "" Then
                If obj.ApakahSudahAda(Me.TxtLocationAfter.Text) = True Then
                    Me.TxtQtyperBatchAfter.Focus()
                    Me.txtPrimaryLocationAfter.Text = obj.GetPrimary(Me.TxtLocationAfter.Text)
                    Me.TxtWarehouseAfter.Text = obj.GetWarehouseCode(Me.txtPrimaryLocationAfter.Text)
                Else
                    If Me.TxtLocationAfter.Text <> "" Then
                        i.txtCari.Text = Me.TxtLocationAfter.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryStorageArea > 0 Then
                        Me.txtPrimaryLocationAfter.Text = i.PrimaryStorageArea
                        Me.TxtLocationAfter.Text = obj.GetStorageAreaCode(Me.txtPrimaryLocationAfter.Text)
                        Me.TxtWarehouseAfter.Text = obj.GetWarehouseCode(Me.txtPrimaryLocationAfter.Text)
                        Me.TxtQtyperBatchAfter.Focus()
                    Else
                        Me.TxtLocationAfter.Clear()
                        Me.TxtWarehouseAfter.Clear()
                        Me.txtPrimaryLocationAfter.Text = "-1"
                        Me.TxtLocationAfter.Focus()
                    End If
                End If
            Else
                Me.txtPrimaryLocationAfter.Text = "-1"
                Me.TxtWarehouseAfter.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub btnFindLocationAfter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindLocationAfter.Click
        Dim i As New FrmCariMasterStorageArea
        Dim obj As New clsMasterStorageArea

        Try
            i.cboCriteria.SelectedIndex = 0
            If Me.TxtLocationAfter.Text <> "" Then
                i.txtCari.Text = Me.TxtLocationAfter.Text
            End If

            i.ShowDialog()
            If i.PrimaryStorageArea > 0 Then
                Me.txtPrimaryLocationAfter.Text = i.PrimaryStorageArea
                Me.TxtLocationAfter.Text = obj.GetStorageAreaCode(Me.txtPrimaryLocationAfter.Text)
                Me.TxtWarehouseAfter.Text = obj.GetWarehouseCode(Me.txtPrimaryLocationAfter.Text)
                Me.TxtQtyperBatchAfter.Focus()
            Else
                Me.TxtWarehouseAfter.Clear()
                Me.TxtLocationAfter.Clear()
                Me.TxtLocationAfter.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
            obj = Nothing
        End Try
    End Sub

    Private Sub dgvItemAfter_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItemAfter.SelectionChanged
        Try
            If Me.txtremarks.Enabled = True Then
                If Me.dgvItemAfter.SelectedRows.Count > 0 Then
                    Me.btnUpAfter.Enabled = True
                    Me.btnDownAfter.Enabled = True
                Else
                    Me.btnUpAfter.Enabled = False
                    Me.btnDownAfter.Enabled = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_Sorted(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dgvItemAfter.Sorted
        Me.ObjCD.SetNumber(Me.dgvItemAfter)
    End Sub

    Private Sub btnUPAfter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpAfter.Click
        Try
            If Me.dgvItemAfter.SelectedRows.Count > 0 Then
                Me.btnUpAfter.Enabled = False
                Me.btnDownAfter.Enabled = False
                Dim row As DataGridViewRow = Me.dgvItemAfter.SelectedRows(0)
                If row.Index <> 0 Then
                    Dim indexbaru As Integer = row.Index - 1
                    Me.dgvItemAfter.Rows.Remove(row)
                    Me.dgvItemAfter.Rows.InsertRange(indexbaru, row)
                    Me.dgvItemAfter.Focus()
                    Me.dgvItemAfter.ClearSelection()
                    row.Selected = True
                    Me.ObjCD.SetNumber(Me.dgvItemAfter)
                    If row.Displayed = False Then
                        Me.ShowRow(row, True, Me.dgvItemAfter)
                    End If
                End If
                Me.btnUpAfter.Enabled = True
                Me.btnDownAfter.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDownAfter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDownAfter.Click
        Try
            If Me.dgvItemAfter.SelectedRows.Count > 0 Then
                Me.btnUpAfter.Enabled = False
                Me.btnDownAfter.Enabled = False
                Dim row As DataGridViewRow = Me.dgvItemAfter.SelectedRows(0)
                If row.Index <> Me.dgvItemAfter.Rows.Count - 1 Then
                    Dim indexbaru As Integer = row.Index + 1
                    Me.dgvItemAfter.Rows.Remove(row)
                    Me.dgvItemAfter.Rows.InsertRange(indexbaru, row)
                    Me.dgvItemAfter.Focus()
                    Me.dgvItemAfter.ClearSelection()
                    row.Selected = True
                    Me.ObjCD.SetNumber(Me.dgvItemAfter)
                    Me.ShowRow(row, False, Me.dgvItemAfter)
                End If
                Me.btnUpAfter.Enabled = True
                Me.btnDownAfter.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemAfter_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItemAfter.UserDeletedRow
        Try
            Me.ObjCD.SetNumber(Me.dgvItemAfter)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_UserDeletingRow(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvItemAfter.UserDeletingRow
        Try
            'ketika user akan menghapus data, cek apakah akibatnya akan mengakibatkan ada stock yang minus
            Dim PrimaryLocation, PrimaryPdO, PrimaryProduct, SecondaryQty, JmlGudang, JmlObjLama, _
            JmlDGV, JmlSetelahHapus As Decimal
            Dim BatchNo As String = ""
            Dim ExpiryDate As Date = #1/1/1950#
            Dim row As DataGridViewRow = e.Row

            With row
                PrimaryLocation = .Cells(13).Value
                PrimaryPdO = -1
                PrimaryProduct = .Cells(12).Value
                BatchNo = .Cells(7).Value
                If IsDate(.Cells(8).Value) = True Then
                    ExpiryDate = CDate(.Cells(8).Value)
                Else
                    ExpiryDate = #1/1/1950#
                End If
                If IsNumeric(.Cells(5).Value) = True Then
                    SecondaryQty = .Cells(5).Value
                Else
                    SecondaryQty = -1
                End If
                JmlGudang = Me.ObjProduct.GetJml3(PrimaryProduct, SecondaryQty, PrimaryLocation, BatchNo)
                'karena akan dihapus, perhitungan menggunakan nilai jumlah dgv =0
                JmlDGV = 0
                JmlObjLama = Me.objStockBundlingItem.GetJmlAfter(Me.TxtPrimaryStockBundling.Text, PrimaryProduct, _
                SecondaryQty, PrimaryLocation, BatchNo, ExpiryDate)
                JmlSetelahHapus = JmlGudang - JmlObjLama + JmlDGV
                If JmlSetelahHapus < 0 Then
                    MsgBox("You cannot delete this data, because this will make you stock quantity minus.", _
                    MsgBoxStyle.Information, "Information")
                    e.Cancel = True
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeNoAfter_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNoAfter.GotFocus
        Try
            If Me.ObjStockBundling.TipeBundling = ClsTrans_StockBundling.EnumTipeBundling.Inventory Then
                Me.TxtCodeNoAfter.BackColor = Core_Module.GetSMESSBackColorAll
            Else
                Me.TxtCodeNoAfter.BackColor = Core_Module.GetSMESSBackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNoAfter.LostFocus
        Try
            Me.TxtCodeNoAfter.BackColor = Me.TxtPrimaryCodeNoAfter.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtBatchNoAfter_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBatchNoAfter.GotFocus
        Try
            If Me.ObjStockBundling.TipeBundling = ClsTrans_StockBundling.EnumTipeBundling.Inventory Then
                Me.txtBatchNoAfter.BackColor = Core_Module.GetSMESSBackColorAll
            Else
                Me.txtBatchNoAfter.BackColor = Core_Module.GetSMESSBackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtBatchNoAfter_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBatchNoAfter.Leave
        Try
            If Me.txtremarks.Enabled = True Then
                Me.txtBatchNoAfter.Text = Me.txtBatchNoAfter.Text.Trim
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtBatchNoAfter_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBatchNoAfter.LostFocus
        Try
            Me.txtBatchNoAfter.BackColor = Me.TxtPrimaryCodeNoAfter.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLocationAfter_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLocationAfter.GotFocus
        Try
            If Me.ObjStockBundling.TipeBundling = ClsTrans_StockBundling.EnumTipeBundling.Inventory Then
                Me.TxtLocationAfter.BackColor = Core_Module.GetSMESSBackColorAll
            Else
                Me.TxtLocationAfter.BackColor = Core_Module.GetSMESSBackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtLocationAfter_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtLocationAfter.LostFocus
        Try
            Me.TxtLocationAfter.BackColor = Me.TxtPrimaryCodeNoAfter.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatchAfter_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchAfter.GotFocus
        Try
            If Me.ObjStockBundling.TipeBundling = ClsTrans_StockBundling.EnumTipeBundling.Inventory Then
                Me.TxtQtyperBatchAfter.BackColor = Core_Module.GetSMESSBackColorAll
            Else
                Me.TxtQtyperBatchAfter.BackColor = Core_Module.GetSMESSBackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatchAfter_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchAfter.LostFocus
        Try
            Me.TxtQtyperBatchAfter.BackColor = Me.TxtPrimaryCodeNoAfter.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSecondaryQtyAfter_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSecondaryQtyAfter.GotFocus
        Try
            If Me.ObjStockBundling.TipeBundling = ClsTrans_StockBundling.EnumTipeBundling.Inventory Then
                Me.txtSecondaryQtyAfter.BackColor = Core_Module.GetSMESSBackColorAll
            Else
                Me.txtSecondaryQtyAfter.BackColor = Core_Module.GetSMESSBackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtSecondaryQtyAfter_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSecondaryQtyAfter.LostFocus
        Try
            Me.txtSecondaryQtyAfter.BackColor = Me.TxtPrimaryCodeNoAfter.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarkItemAfter_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarkItemAfter.GotFocus
        Try
            If Me.ObjStockBundling.TipeBundling = ClsTrans_StockBundling.EnumTipeBundling.Inventory Then
                Me.TxtRemarkItemAfter.BackColor = Core_Module.GetSMESSBackColorAll
            Else
                Me.TxtRemarkItemAfter.BackColor = Core_Module.GetSMESSBackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarkItemAfter_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarkItemAfter.LostFocus
        Try
            Me.TxtRemarkItemAfter.BackColor = Me.TxtPrimaryCodeNoAfter.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtFaktorPemberat_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFaktorPemberat.Enter
        Try
            Me.ObjCD.Number_Enter(Me.txtFaktorPemberat, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtFaktorPemberat_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFaktorPemberat.GotFocus
        Try
            Me.txtFaktorPemberat.BackColor = Core_Module.GetSMESSBackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtFaktorPemberat_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFaktorPemberat.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.txtFaktorPemberat, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtFaktorPemberat_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFaktorPemberat.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.txtFaktorPemberat, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtFaktorPemberat_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFaktorPemberat.Leave
        Try
            Me.txtFaktorPemberat.Text = Me.ObjCD.DigitGrouping(Me.txtFaktorPemberat.Text, 2)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtFaktorPemberat_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFaktorPemberat.LostFocus
        Try
            Me.txtFaktorPemberat.BackColor = Me.txtPrimaryCategoryAfter.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatchAfter_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchAfter.Leave
        Try
            If Me.txtremarks.Enabled = True Then
                If Me.TxtPrimaryCodeNoAfter.Text <> "" And Me.TxtPrimaryCodeNoAfter.Text <> "" Then
                    Me.TxtQtyperBatchAfter.Text = Me.ObjCD.DigitGrouping(Me.TxtQtyperBatchAfter.Text, Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNoAfter.Text))
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpDate.ValueChanged
        Try
            If Me.txtremarks.Enabled = True Then
                If Me.baru = False Then
                    If Me.ObjStockBundling.TanggalTransaksi.Month = Me.dtpDate.Value.Month And Me.ObjStockBundling.TanggalTransaksi.Year = Me.dtpDate.Value.Year Then
                        'jika kembali ke bulan yang bersangkutan maka no kembali ke awal
                        Me.txtstockbundling.Text = Me.ObjStockBundling.NoTransaksi
                    Else
                        'jika pindah bulan maka akan ganti no
                        Me.txtstockbundling.Text = Me.ObjStockBundling.GetNoBaru(Me.dtpDate.Value, Me.tipe_bundling)
                    End If
                Else
                    Me.txtstockbundling.Text = Me.ObjStockBundling.GetNoBaru(Me.dtpDate.Value, Me.tipe_bundling)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatchb4_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchb4.Leave
        Try
            If Me.txtremarks.Enabled = True Then
                If Me.TxtPrimaryCodeNob4.Text <> "" And Me.TxtPrimaryCodeNob4.Text <> "-1" Then
                    Me.TxtQtyperBatchb4.Text = Me.ObjCD.DigitGrouping(Me.TxtQtyperBatchb4.Text, Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNob4.Text))
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tc_processoutput.SelectedIndexChanged
        Try
            If Me.txtremarks.Enabled = True Then

                ''tanya lokasi untuk isi RM saat pindah ke tab RM
                If Me.tc_processoutput.SelectedTab.Name = Me.tabpage_rmusage.Name Then
                    Me.IsikanRawMaterial()
                End If

                If Me.tc_processoutput.SelectedTab.Name = Me.tabpage_processoutput.Name Then
                    ModCekData.KendalikanControl(Me.gboItemb4.Controls, True, True, True, True, True, True, True, True)
                Else
                    ModCekData.KendalikanControl(Me.gboItemAfter.Controls, True, True, True, True, True, True, True, True)
                    ModCekData.KendalikanControl(Me.SplitContainer5.Panel1.Controls, True, True, True, True, True, True, True, True)
                    ModCekData.KendalikanControl(Me.SplitContainer5.Panel2.Controls, True, True, True, True, True, True, True, True)
                End If
            Else
                If Me.tc_processoutput.SelectedTab.Name = Me.tabpage_processoutput.Name Then
                    ModCekData.KendalikanControl(Me.gboItemb4.Controls, True, False, True, False, True, False, True, False)
                Else
                    ModCekData.KendalikanControl(Me.gboItemAfter.Controls, True, False, True, False, True, False, True, False)
                    ModCekData.KendalikanControl(Me.SplitContainer5.Panel1.Controls, True, False, True, False, True, False, True, False)
                    ModCekData.KendalikanControl(Me.SplitContainer5.Panel2.Controls, True, False, True, False, True, False, True, False)
                End If
            End If

            If Me.txtremarks.Enabled = True Then
                If Me.tc_processoutput.SelectedTab.Name = Me.tabpage_processoutput.Name Then
                    If Me.dgvItemb4.Rows.Count > 0 Then
                        Me.SetGboItemAfter()
                    End If
                End If

                If Me.tc_processoutput.SelectedTab.Name = Me.tabpage_rejectdowntimeoperdata.Name Then
                    If ApakahCodeNoDariWOSudahAda() = False Then
                        MsgBox("Define work order number first.", MsgBoxStyle.Information, "Information")
                        Me.ForceNavigate = True
                        Me.NavigateTabPage(Me.tabpage_processoutput)
                        Me.txtwonumber.Focus()
                    Else
                        Me.CalculateAll()
                    End If
                End If

                If Me.tc_processoutput.SelectedTab.Name = Me.tabpage_rmusage.Name Then
                    If ApakahCodeNoDariWOSudahAda() = False Then
                        MsgBox("Define work order number first.", MsgBoxStyle.Information, "Information")
                        Me.ForceNavigate = True
                        Me.NavigateTabPage(Me.tabpage_processoutput)
                        Me.txtwonumber.Focus()
                    Else
                        Me.SetGboItemb4()
                    End If
                End If
            End If

            If Me.tc_processoutput.SelectedTab.Name = Me.tabpage_rejectdowntimeoperdata.Name Then
                If Me.txttotal_produce.Enabled Then
                    Me.txttotal_produce.Focus()
                Else
                    Me.dgvinput_processreject.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub CalculateAll()
        Try
            Me.txttotal_good.Text = Me.ObjCD.DigitGrouping(Me.get_totalqty_indgvafter, Me.isDecimal)
            Me.txttotal_reject.Text = Me.ObjCD.DigitGrouping(Me.get_totalqty_indgvpj, Me.isDecimal)
            Me.txttotal_produce.Text = Me.ObjCD.DigitGrouping(CDec(Me.txttotal_good.Text) + CDec(Me.txttotal_reject.Text), Me.isDecimal)
            If CDec(Me.txttotal_good.Text) > 0 Then
                Me.txtgood_perct.Text = Me.ObjCD.DigitGrouping(CDec(Me.txttotal_good.Text) / CDec(Me.txttotal_produce.Text) * 100, 2)
            Else
                Me.txtgood_perct.Text = Me.ObjCD.DigitGrouping(0, 2)
            End If

            'Me.HitungDurasi(Me.dtpactual_start.Value, True, False)
            Me.GetTotalSetupTime()
            Me.GetTotalDownTime()
            Me.GetTotalWorkHour()
            Me.GetTotalProdRate()

            Me.BufferTime_Start = Me.dtpactual_start.Value
            Me.BufferTime_End = Me.dtpactual_end.Value
            Dim Buffer_ActualStartDate As Date = Me.dtpplan_start.Value
            Dim Buffer_ActualEndDate As Date = Me.dtpplan_end.Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AddByBillOfMaterial()
        Dim ijk As New FrmTransaksiStockBundling_BOM

        Try
            If Me.txtremarks.Enabled = True Then
                ijk.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ijk = Nothing
        End Try
    End Sub

    Private Sub llb_BOM_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llb_BOM.LinkClicked
        Dim ijk As New Core_Module.FrmMasterLocationPriority
        Dim Lanjut As Boolean = False
        Try
            'If Me.txtremarks.Enabled = True And Me.dgvItemAfter.Rows.Count > 0 Then
            '    If MsgBox("Add item before by bill of material?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
            '        'kosongkan item before
            '        Me.dgvItemb4.Rows.Clear()
            '        Me.KosongkanItemB4()
            '        Me.SetGboItemb4()
            '        Me.tc_processoutput.SelectedIndex = 0

            '        For Each rowDGV As DataGridViewRow In Me.dgvItemAfter.Rows
            '            Me.AddItemRawMaterial(rowDGV.Cells(Me.ColPrimaryCodeNoAfter.Index).Value, rowDGV.Cells(Me.ColQtyAfter.Index).Value)
            '        Next

            '        Me.KosongkanItemB4()
            '        Me.SetGboItemb4()
            '    End If
            'End If

            If Me.txtremarks.Enabled = True And Me.dgvItemAfter.Rows.Count > 0 Then
                If MsgBox("Add item before by bill of material?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    If Me.ApakahAda_ItemRM_yg_LocationControl = True Then
                        'isikan prioritas item before
                        ijk.BackColorAll = Me.BackColorAll
                        ijk.UserCode = Me.UserCode
                        ijk.vTipeTransaksi = ClsMasterLocationPriority.EnumTipeTransaksi.StockBundling
                        If ijk.ShowDialog() = Windows.Forms.DialogResult.OK Then
                            Lanjut = True
                            Me.ObjLocationPriority.GetLocationPrioriry(Me.TBLLocationPriority, Me.UserCode, ClsMasterLocationPriority.EnumTipeTransaksi.StockBundling)
                        Else
                            Lanjut = False
                        End If
                    Else
                        Lanjut = False
                    End If

                    If Lanjut = True Then
                        'kosongkan item before
                        Me.dgvItemb4.Rows.Clear()
                        Me.KosongkanItemB4()
                        Me.SetGboItemb4()
                        Me.tc_processoutput.SelectedTab = tabpage_rmusage

                        For Each rowDGV As DataGridViewRow In Me.dgvItemAfter.Rows
                            Me.AddItemRawMaterial(rowDGV.Cells(Me.ColPrimaryCodeNoAfter.Index).Value, Me.txtidentifier_wo.Text, _
                            rowDGV.Cells(Me.ColQtyAfter.Index).Value, Me.TxtIdentifierJO.Text, Me.TBLLocationPriority)
                        Next

                        Me.KosongkanItemB4()
                        Me.SetGboItemb4()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ijk = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' mengembalikan nilai true, jika ada salah satu rm dari item2 after yang location control
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ApakahAda_ItemRM_yg_LocationControl() As Boolean
        Dim DTA As New DseTrans_StockBundlingTableAdapters.DataTable7TableAdapter
        Dim TBL As New DseTrans_StockBundling.DataTable7DataTable
        Dim row As DseTrans_StockBundling.DataTable7Row
        Dim rowDGV As DataGridViewRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            For Each rowDGV In Me.dgvItemAfter.Rows
                DTA.Fill(TBL, rowDGV.Cells(Me.ColPrimaryCodeNoAfter.Index).Value, Me.TxtIdentifierJO.Text, Me.txtidentifier_wo.Text)
                For Each row In TBL.Rows
                    If row.Location_Controlled = 1 Then
                        Return True
                    End If
                Next
            Next

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Private Sub IsikanItemQty(ByVal vTBLProductQty As DseTrans_StockBundling.DataTable1DataTable, _
    ByVal vTBLItemBundling As DseTrans_StockBundling.tbltransaksi_stockbundling_item_before4DataTable, _
    ByVal vPrimaryProduct As Double)
        Dim rowProductQty As DseTrans_StockBundling.DataTable1Row
        Dim rowItemBundling As DseTrans_StockBundling.tbltransaksi_stockbundling_item_before4Row
        Dim DV As Data.DataRowView
        Dim ketemu As Boolean = False
        Dim CodeNo As String = ""
        Dim Desc As String = ""
        Dim DecimalPoint As Integer = 0
        Dim PrimaryUnit As Double = -1
        Dim UnitCode As String = ""

        Try
            'tambahkan dengan item db
            vTBLItemBundling.DefaultView.RowFilter = "[Primary Product] = " + vPrimaryProduct.ToString
            For Each DV In vTBLItemBundling.DefaultView
                rowItemBundling = DV.Row
                ketemu = False

                For Each rowProductQty In vTBLProductQty.Rows
                    If rowProductQty.Lot_No = rowItemBundling.Batch_No And rowProductQty.Expiry_Date = rowItemBundling.Expiry_Date And _
                    rowProductQty.Primary_Storage_Area = rowItemBundling.Primary_Storage_Area Then
                        ketemu = True
                        rowProductQty.Qty += rowItemBundling.Qty
                        Exit For
                    End If
                Next

                If ketemu = False Then
                    If CodeNo = "" Then
                        CodeNo = Me.ObjProduct.GetProductCode(rowItemBundling.Primary_Product)
                        Desc = Me.ObjProduct.GetDescription(rowItemBundling.Primary_Product)
                        DecimalPoint = Me.ObjProduct.GetAngkaDecimal(rowItemBundling.Primary_Product)
                        PrimaryUnit = Me.ObjProduct.GetPrimaryUnit(rowItemBundling.Primary_Product)
                        UnitCode = Me.ObjProduct.GetUnitCode(rowItemBundling.Primary_Product)
                    End If
                    rowProductQty = vTBLProductQty.NewDataTable1Row
                    With rowProductQty
                        .Description = Desc
                        .Product_Code = CodeNo
                        .Expiry_Date = rowItemBundling.Expiry_Date
                        .Lot_No = rowItemBundling.Batch_No
                        .Primary_Product = rowItemBundling.Primary_Product
                        .Primary_Storage_Area = rowItemBundling.Primary_Storage_Area
                        .Qty = rowItemBundling.Qty
                        ._Decimal = DecimalPoint
                        .Primary_Unit = PrimaryUnit
                        .Unit_Code = UnitCode
                        .Storage_Area_Code = rowItemBundling.Storage_Area_Code
                        .Warehouse_Code = rowItemBundling.Warehouse_Code
                    End With
                    vTBLProductQty.Rows.Add(rowProductQty)
                End If
            Next

            'kurangkan dengan dgv
            For Each rowDGV As DataGridViewRow In Me.dgvItemb4.Rows
                If rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct Then
                    ketemu = False
                    For Each rowProductQty In vTBLProductQty.Rows
                        If rowProductQty.Lot_No = rowDGV.Cells(Me.ColBatch.Index).Value And _
                        rowProductQty.Primary_Storage_Area = rowDGV.Cells(Me.ColPrimaryLocation.Index).Value Then
                            ketemu = True
                            rowProductQty.Qty -= CDec(rowDGV.Cells(Me.ColQty.Index).Value)
                            Exit For
                        End If
                    Next

                    If ketemu = False Then
                        'tambahkan baris dengan qty minus
                        If CodeNo = "" Then
                            CodeNo = rowDGV.Cells(Me.ColCodeNo.Index).Value
                            Desc = rowDGV.Cells(Me.ColDescription.Index).Value
                            DecimalPoint = Me.ObjProduct.GetAngkaDecimal(vPrimaryProduct)
                            PrimaryUnit = Me.ObjProduct.GetPrimaryUnit(vPrimaryProduct)
                            UnitCode = Me.ObjProduct.GetUnitCode(vPrimaryProduct)
                        End If
                        rowProductQty = vTBLProductQty.NewDataTable1Row
                        With rowProductQty
                            .Description = Desc
                            .Product_Code = CodeNo
                            If IsDate(rowDGV.Cells(Me.ColExpiredDate.Index).Value) = True Then
                                .Expiry_Date = rowDGV.Cells(Me.ColExpiredDate.Index).Value
                            Else
                                .Expiry_Date = #1/1/1950#
                            End If
                            .Lot_No = rowDGV.Cells(Me.ColBatch.Index).Value
                            .Primary_Product = rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value
                            .Primary_Storage_Area = rowDGV.Cells(Me.ColPrimaryLocation.Index).Value
                            .Qty = CDec(rowDGV.Cells(Me.ColQty.Index).Value) * (-1)
                            ._Decimal = DecimalPoint
                            .Primary_Unit = PrimaryUnit
                            .Unit_Code = UnitCode
                            .Storage_Area_Code = rowDGV.Cells(Me.ColLocation.Index).Value
                            .Warehouse_Code = rowDGV.Cells(Me.ColWareHouse.Index).Value
                        End With
                        vTBLProductQty.Rows.Add(rowProductQty)
                    End If
                End If
            Next

            For ij As Integer = vTBLProductQty.Rows.Count - 1 To 0 Step -1
                rowProductQty = vTBLProductQty.Rows(ij)
                If rowProductQty.Qty <= 0 Then
                    vTBLProductQty.Rows.Remove(rowProductQty)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub PrioritasStrictLocation(ByVal vTBLProductQty As DseTrans_StockBundling.DataTable1DataTable, _
    ByVal vTBLLocationPriority As Core_Module.DseMasterLocationPriority.tblmaster_location_priority1_DataTable)
        Dim TblDummy As New DseTrans_StockBundling.DataTable1DataTable
        Dim rowDummy As DseTrans_StockBundling.DataTable1Row
        Dim rowProductQty As DseTrans_StockBundling.DataTable1Row
        Dim RowLocation As Core_Module.DseMasterLocationPriority.tblmaster_location_priority1_Row
        Dim DV As Data.DataRowView

        Try
            'pindahkan ke tabel dummy
            vTBLProductQty.DefaultView.Sort = "[Expiry Date], [Lot No]"
            For Each RowLocation In vTBLLocationPriority.Rows
                vTBLProductQty.DefaultView.RowFilter = "[Primary Storage Area] = " + RowLocation.primary_location.ToString

                For Each DV In vTBLProductQty.DefaultView
                    rowProductQty = DV.Row
                    rowDummy = TblDummy.NewDataTable1Row
                    For ij As Integer = 0 To TblDummy.Columns.Count - 1
                        rowDummy.Item(ij) = rowProductQty.Item(ij)
                    Next
                    TblDummy.Rows.Add(rowDummy)
                Next

                While vTBLProductQty.DefaultView.Count > 0
                    rowProductQty = vTBLProductQty.DefaultView(0).Row
                    vTBLProductQty.Rows.Remove(rowProductQty)
                End While
            Next

            'pindahkan sisanya ke tabel dummy
            For Each rowProductQty In vTBLProductQty.Rows
                rowDummy = TblDummy.NewDataTable1Row
                For ij As Integer = 0 To TblDummy.Columns.Count - 1
                    rowDummy.Item(ij) = rowProductQty.Item(ij)
                Next
                TblDummy.Rows.Add(rowDummy)
            Next

            'kembali pindahkan ke tabel utama
            vTBLProductQty.Rows.Clear()
            For Each rowDummy In TblDummy.Rows
                rowProductQty = vTBLProductQty.NewDataTable1Row
                For ij As Integer = 0 To vTBLProductQty.Columns.Count - 1
                    rowProductQty.Item(ij) = rowDummy.Item(ij)
                Next
                vTBLProductQty.Rows.Add(rowProductQty)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TblDummy.Dispose()
        End Try
    End Sub

    Private Sub FilterStrictLocation(ByVal vTBLProductQty As DseTrans_StockBundling.DataTable1DataTable, _
    ByVal vTBLLocationPriority As Core_Module.DseMasterLocationPriority.tblmaster_location_priority1_DataTable)
        Dim TblDummy As New DseTrans_StockBundling.DataTable1DataTable
        Dim rowDummy As DseTrans_StockBundling.DataTable1Row
        Dim rowProductQty As DseTrans_StockBundling.DataTable1Row
        Dim RowLocation As Core_Module.DseMasterLocationPriority.tblmaster_location_priority1_Row
        Dim DV As Data.DataRowView

        Try
            'pindahkan ke tabel dummy
            vTBLProductQty.DefaultView.Sort = "[Expiry Date], [Lot No]"
            For Each RowLocation In vTBLLocationPriority.Rows
                vTBLProductQty.DefaultView.RowFilter = "[Primary Storage Area] = " + RowLocation.primary_location.ToString

                For Each DV In vTBLProductQty.DefaultView
                    rowProductQty = DV.Row
                    rowDummy = TblDummy.NewDataTable1Row
                    For ij As Integer = 0 To TblDummy.Columns.Count - 1
                        rowDummy.Item(ij) = rowProductQty.Item(ij)
                    Next
                    TblDummy.Rows.Add(rowDummy)
                Next
            Next

            'kembali pindahkan ke tabel utama
            vTBLProductQty.Rows.Clear()
            For Each rowDummy In TblDummy.Rows
                rowProductQty = vTBLProductQty.NewDataTable1Row
                For ij As Integer = 0 To vTBLProductQty.Columns.Count - 1
                    rowProductQty.Item(ij) = rowDummy.Item(ij)
                Next
                vTBLProductQty.Rows.Add(rowProductQty)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TblDummy.Dispose()
        End Try
    End Sub

    Public Function ModifTblRawMaterial_Replacement(ByVal videntifierWO As String, _
    ByVal tbl_BOM As DseTrans_StockBundling.tbltrans_mo_item_bomDataTable, ByVal vIdentifierJO As String) As Boolean
        Dim Row_BOM, Row_BOM1 As DseTrans_StockBundling.tbltrans_mo_item_bomRow
        Dim DTARM As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_item_rmTableAdapter
        Dim TBLRM As New DseTrans_ManufacturingOrder.tbltrans_mo_item_rmDataTable
        Dim rowRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmRow
        Dim AdaReplacement As Boolean = False
        Dim DTA_Bundling As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling_item_before7TableAdapter
        Dim TBL_Bundling As New DseTrans_StockBundling.tbltransaksi_stockbundling_item_before7DataTable
        Dim row_Bundling As DseTrans_StockBundling.tbltransaksi_stockbundling_item_before7Row
        Dim TBLRef As New DseTrans_StockBundling.TblAvailableToUseDataTable
        Dim rowRef As DseTrans_StockBundling.TblAvailableToUseRow
        Dim DefPrecReplaced As Integer = 0
        Dim IdxAcuan As Integer = -1
        Dim QtyReplaced As Decimal = 0

        Try
            DTARM.Connection.ConnectionString = Core_Module.GetConstring
            DTA_Bundling.Connection.ConnectionString = Core_Module.GetConstring

            TBLRef.Rows.Clear()

            'isikan tabel rm acuan
            DTARM.FillBy_IdentifierJO_WO(TBLRM, vIdentifierJO.Substring(0, vIdentifierJO.IndexOf("-")), vIdentifierJO, videntifierWO)
            For Each rowRM In TBLRM.Rows
                If rowRM.is_replacement = 1 Then
                    AdaReplacement = True
                    Exit For
                End If
            Next

            'jika ada pengganti, maka jalankan coding dibawah
            If AdaReplacement = True Then
                'isikan tbl used, berdasarkan bundling
                DTA_Bundling.Fill(TBL_Bundling, vIdentifierJO.Substring(0, vIdentifierJO.IndexOf("-")), videntifierWO, vIdentifierJO, Me.TxtPrimaryStockBundling.Text)
                For ij As Integer = tbl_BOM.Rows.Count - 1 To 0 Step -1
                    Row_BOM = tbl_BOM.Rows(ij)
                    AdaReplacement = False
                    IdxAcuan = -1

                    For Each rowRM In TBLRM.Rows
                        If rowRM.is_replacement = 1 And rowRM.primary_code_rm_replaced = Row_BOM.Primary_Product_RM Then
                            AdaReplacement = True
                            Exit For
                        End If
                    Next

                    'untuk setiap item rm, jika ada replacementnya, maka modifikasi row-nya
                    If AdaReplacement = True Then
                        DefPrecReplaced = Me.ObjProduct.GetAngkaDecimal(Row_BOM.Primary_Product_RM)
                        'a. kumpulkan baris replaced dan baris replacement, urutkan berdasarkan prioritas
                        TBLRM.DefaultView.RowFilter = "[primary code rm replaced] = " + Row_BOM.Primary_Product_RM.ToString + " OR [Primary Product Code RM] = " + Row_BOM.Primary_Product_RM.ToString
                        TBLRM.DefaultView.Sort = "[use priority] ASC"

                        'b. pindahkan qty awal
                        TBLRef.Rows.Clear()
                        For Each DV As Data.DataRowView In TBLRM.DefaultView
                            rowRM = DV.Row
                            rowRef = TBLRef.NewTblAvailableToUseRow
                            With rowRef
                                .Def_Prec_RM_Replaced = DefPrecReplaced
                                .Def_Prec_RM_Replacement = Me.ObjProduct.GetAngkaDecimal(rowRM.Primary_Product_Code_RM)
                                .Primary_Code_RM_Replaced = rowRM.primary_code_rm_replaced
                                .Primary_Code_RM_Replacement = rowRM.Primary_Product_Code_RM
                                .Qty_Replaced = 0
                                .Qty_Replacement = rowRM.Required_Qty - rowRM.Replace_Qty
                                .Ratio_Replaced = rowRM.ratio_replaced
                                .Ratio_Replacement = rowRM.ratio_replacement
                                .Code_No_RM_Replacement = rowRM.rm_code
                                .Desc_RM_Replacement = rowRM.rm_desc
                                .Decimal_RM_Replacement = rowRM._decimal
                                .Unit_Code_RM_Replacement = rowRM.unit_code
                            End With
                            TBLRef.Rows.Add(rowRef)
                        Next

                        'c. kurangkan dengan qty replament dengan qty used bundling
                        For Each rowRef In TBLRef.Rows
                            For Each row_Bundling In TBL_Bundling.Rows
                                If rowRef.Primary_Code_RM_Replacement = row_Bundling.Primary_Product Then
                                    rowRef.Qty_Replacement -= row_Bundling.Qty
                                    Exit For
                                End If
                            Next
                        Next

                        'd. hitung berapa yang harus dipindahkan ke tabel bom, berdasarkan prioritas peggunaan, rasio, qty yang dibutuhkan 
                        'hasil perhitungan dipindahkan ke dalam tabel bom sebagai baris baru
                        IdxAcuan = 0
                        While Row_BOM.Qty_RM > 0 And IdxAcuan <= TBLRef.Rows.Count - 1
                            rowRef = TBLRef.Rows(IdxAcuan)
                            If rowRef.Qty_Replacement > 0 Then
                                Row_BOM1 = tbl_BOM.Newtbltrans_mo_item_bomRow
                                For gh As Integer = 0 To tbl_BOM.Columns.Count - 1
                                    Row_BOM1.Item(gh) = Row_BOM.Item(gh)
                                Next
                                'ganti data sesuai rm replacement
                                Row_BOM1.Primary_Product_RM = rowRef.Primary_Code_RM_Replacement
                                Row_BOM1.code_rm = rowRef.Code_No_RM_Replacement
                                Row_BOM1.desc_rm = rowRef.Desc_RM_Replacement
                                Row_BOM1.Unit_Code = rowRef.Unit_Code_RM_Replacement
                                Row_BOM.Decimal_RM = rowRef.Decimal_RM_Replacement

                                QtyReplaced = Math.Round((rowRef.Qty_Replacement * rowRef.Ratio_Replacement) / rowRef.Ratio_Replaced, rowRef.Def_Prec_RM_Replaced, MidpointRounding.AwayFromZero)
                                If QtyReplaced > Row_BOM.Qty_RM Then
                                    'pindahkan sejumlah qty pada row bom
                                    Row_BOM1.Qty_RM = Math.Round((Row_BOM.Qty_RM * rowRef.Ratio_Replaced) / rowRef.Ratio_Replacement, rowRef.Def_Prec_RM_Replacement, MidpointRounding.AwayFromZero)
                                    'set qty bom =0
                                    Row_BOM.Qty_RM = 0
                                Else
                                    'pindahkan sejumlah qty replaced
                                    Row_BOM1.Qty_RM = rowRef.Qty_Replacement
                                    'kurangi qty pada row bom sejumlah qty replaced
                                    Row_BOM.Qty_RM -= QtyReplaced
                                End If
                                tbl_BOM.Rows.Add(Row_BOM1)
                            End If
                            IdxAcuan += 1
                        End While

                        'f. buang baris bom acuan
                        tbl_BOM.Rows.RemoveAt(ij)

                    End If
                Next


            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub AddItemRawMaterial(ByVal vPrimaryProduct As Double, ByVal videntifierWO As String, ByVal vQtyAfter As Decimal, ByVal videntifierJO As String, _
    ByVal vTBLLocationPriority As Core_Module.DseMasterLocationPriority.tblmaster_location_priority1_DataTable)
        Dim ij As Integer = 0
        Dim QtyEntry As Decimal = 0
        'Dim JmlGudang, JmlObjLama, JmlSetelahTambah, JmlDGV As Decimal
        Dim Dummy As String = ""
        Dim GoodPercentage As Decimal = 0
        Dim DummyQty As Decimal = 0
        Dim DummyQty1 As Decimal = 0
        Dim NST As Decimal = 0
        Dim ObjUSER As New Core_Module.ClsUser
        Dim LocationControl As Boolean = True
        Dim DTAProductQty As New DseTrans_StockBundlingTableAdapters.DataTable1TableAdapter
        Dim TBLProductQty As New DseTrans_StockBundling.DataTable1DataTable
        Dim rowProductQty As DseTrans_StockBundling.DataTable1Row
        Dim DTAItemBundling As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling_item_before4TableAdapter
        Dim TBLItemBundling As New DseTrans_StockBundling.tbltransaksi_stockbundling_item_before4DataTable
        Dim StrictLocation As Boolean = False
        Dim jk As Integer = 0

        Try
            Me.SedangAddRM = True

            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiStockBundling_NoStrictLocation2") = True Then
                StrictLocation = False
            Else
                StrictLocation = True
            End If

            DTAProductQty.Connection.ConnectionString = Core_Module.GetConstring
            DTAItemBundling.Connection.ConnectionString = DTAProductQty.Connection.ConnectionString

            Me.dgvItemb4.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItemb4.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            'isikan tblrm berdasarkan bill of material
            Me.ObjStockBundling.IsikanTblRawMaterial(vPrimaryProduct, videntifierWO, Me.TBL_BOM, videntifierJO)

            'kalikan qty raw berdasarkan good percentage umum, good percentage detail, qty hasil
            For Each Me.row_BOM In Me.TBL_BOM.Rows
                'kalikan good percentage item

                'DummyQty = (((((Me.row_BOM.Qty_RM * 100) / Me.row_BOM.Good_Percentage_RM) * 100) / Me.row_BOM.Good_Percentage_Process) * vQtyAfter) / Me.row_BOM.Qty_Ref_Process
                'DummyQty1 = Math.Round((((((Me.row_BOM.Qty_RM * 100) / Me.row_BOM.Good_Percentage_RM) * 100) / Me.row_BOM.Good_Percentage_Process) * vQtyAfter) / Me.row_BOM.Qty_Ref_Process, _
                'Me.row_BOM.Decimal_RM, MidpointRounding.AwayFromZero)
                DummyQty = (((Me.row_BOM.Qty_RM * 100) / Me.row_BOM.Good_Percentage_RM) * vQtyAfter) / Me.row_BOM.Qty_Ref_Process
                DummyQty1 = Math.Round((((Me.row_BOM.Qty_RM * 100) / Me.row_BOM.Good_Percentage_RM) * vQtyAfter) / Me.row_BOM.Qty_Ref_Process, _
                Me.row_BOM.Decimal_RM, MidpointRounding.AwayFromZero)

                If DummyQty1 > DummyQty Then
                    DummyQty = DummyQty1
                End If

                'Dim balance_qtyrm As Double = Me.ObjStockBundling.get_qtyrmbalance(Me.txtidentifier_wo.Text, Me.row_BOM.Primary_Product_RM)
                Me.row_BOM.Qty_RM = Math.Round(DummyQty, Me.row_BOM.Decimal_RM, MidpointRounding.AwayFromZero)
            Next

            'modifikasi table bom untuk mengakomodasi rm replacement
            Me.ModifTblRawMaterial_Replacement(videntifierWO, TBL_BOM, videntifierJO)

            'input ke dalam product, berdasarkan location priority
            TBLItemBundling.Rows.Clear()
            If Me.baru = False Then
                DTAItemBundling.Fill(TBLItemBundling, Me.TxtPrimaryStockBundling.Text)
            End If
            Me.KosongkanItemB4()

            While ij < Me.TBL_BOM.Rows.Count
                Me.row_BOM = Me.TBL_BOM.Rows(ij)
                LocationControl = Me.ObjProduct.ApakahLocationControlled(Me.row_BOM.Primary_Product_RM)

                Me.ObjProduct.IsiCheckBoxControll(Me.row_BOM.Primary_Product_RM, Me.chkLocationControlledb4, _
                Me.chkBatchNoControlledb4, Me.chkExpiryDateControlledb4, Me.chkUseSecondaryUnitb4, _
                Me.chkPdoControlledb4, Me.txtPrimaryUnitsb4, Me.txtPrimarySecondaryUnitsb4)

                'load item product qty
                DTAProductQty.Fill(TBLProductQty, Me.row_BOM.Primary_Product_RM)
                Me.IsikanItemQty(TBLProductQty, TBLItemBundling, Me.row_BOM.Primary_Product_RM)

                If LocationControl = True And StrictLocation = True Then
                    'jika location control dan strict location, maka buang yang lokasinya tidak ada dalam tabel location control
                    Me.FilterStrictLocation(TBLProductQty, vTBLLocationPriority)
                ElseIf LocationControl = True And StrictLocation = False Then
                    'jika location control dan tidak strict location, maka pindahkan lokasi yang menjadi prioritas ke baris paling atas
                    Me.PrioritasStrictLocation(TBLProductQty, vTBLLocationPriority)
                End If

                'dari baris paling atas tabel item product qty
                jk = 0
                While Me.row_BOM.Qty_RM > 0 And jk <= TBLProductQty.Rows.Count - 1
                    rowProductQty = TBLProductQty.Rows(jk)
                    If Me.row_BOM.Qty_RM > rowProductQty.Qty Then
                        'masukkan ke dgv b4
                        Me.TxtPrimaryCodeNob4.Text = rowProductQty.Primary_Product
                        Me.TxtCodeNob4.Text = rowProductQty.Product_Code
                        Me.txtDescb4.Text = rowProductQty.Description
                        Me.TxtQtyperBatchb4.Text = Me.ObjCD.DigitGrouping(rowProductQty.Qty, rowProductQty._Decimal)
                        Me.txtPrimaryUnitsb4.Text = rowProductQty.Primary_Unit
                        Me.txtPrimaryLocationb4.Text = rowProductQty.Primary_Storage_Area
                        Me.cboLocationb4.Items.Clear()
                        Me.cboLocationb4.Items.Add(rowProductQty.Storage_Area_Code)
                        Me.cboLocationb4.Text = rowProductQty.Storage_Area_Code
                        Me.TxtWarehouseb4.Text = rowProductQty.Warehouse_Code
                        Me.txtSecondaryQtyb4.Text = "-1"
                        Me.lblUnitsb4.Text = rowProductQty.Unit_Code
                        Me.cboBatchNob4.Items.Clear()
                        Me.cboBatchNob4.Items.Add(rowProductQty.Lot_No)
                        Me.cboBatchNob4.Text = rowProductQty.Lot_No
                        Me.DtpExpiredb4.Value = rowProductQty.Expiry_Date
                        Me.InputB4(False, False, False)

                        'edit qty rm item bom
                        Me.row_BOM.Qty_RM -= rowProductQty.Qty

                        'ambil semua item product qty
                        rowProductQty.Qty = 0

                        'pindah ke baris item product qty selanjutnya
                        jk += 1
                    Else
                        'masukkan ke dgv b4
                        Me.TxtPrimaryCodeNob4.Text = rowProductQty.Primary_Product
                        Me.TxtCodeNob4.Text = rowProductQty.Product_Code
                        Me.txtDescb4.Text = rowProductQty.Description
                        Me.TxtQtyperBatchb4.Text = Me.ObjCD.DigitGrouping(Me.row_BOM.Qty_RM, rowProductQty._Decimal)
                        Me.txtPrimaryUnitsb4.Text = rowProductQty.Primary_Unit
                        Me.lblUnitsb4.Text = rowProductQty.Unit_Code
                        Me.txtSecondaryQtyb4.Text = "-1"
                        Me.txtPrimaryLocationb4.Text = rowProductQty.Primary_Storage_Area
                        Me.cboLocationb4.Items.Clear()
                        Me.cboLocationb4.Items.Add(rowProductQty.Storage_Area_Code)
                        Me.cboLocationb4.Text = rowProductQty.Storage_Area_Code
                        Me.TxtWarehouseb4.Text = rowProductQty.Warehouse_Code

                        Me.cboBatchNob4.Items.Clear()
                        Me.cboBatchNob4.Items.Add(rowProductQty.Lot_No)
                        Me.cboBatchNob4.Text = rowProductQty.Lot_No
                        Me.DtpExpiredb4.Value = rowProductQty.Expiry_Date
                        Me.InputB4(False, False, False)

                        'edit qty item product qty
                        rowProductQty.Qty -= Me.row_BOM.Qty_RM

                        'ambil semua qty item bom
                        Me.row_BOM.Qty_RM = 0
                    End If
                End While

                ij += 1
            End While

            Me.dgvItemb4.Sort(Me.ColCodeNo, System.ComponentModel.ListSortDirection.Ascending)
            Me.ObjCD.SetNumber(Me.dgvItemb4, False)
            'kosongkan gbo item
            Me.KosongkanItemB4()
            Me.TxtCodeNob4.Focus()
            'format dgv
            Me.FormatDGV(Me.dgvItemb4)
            'fokuskan ke baris terakhir
            If Me.dgvItemb4.RowCount > 0 Then
                While Me.dgvItemb4.Rows(Me.dgvItemb4.Rows.Count - 1).Displayed = False _
                Or Me.dgvItemb4.DisplayedRowCount(True) <> Me.dgvItemb4.DisplayedRowCount(False)
                    Me.dgvItemb4.FirstDisplayedScrollingRowIndex = Me.dgvItemb4.FirstDisplayedScrollingRowIndex + 1
                End While
            End If

            'berikan peringatan bila qty raw material yang dimasukkan tidak cukup
            For ij = TBL_BOM.Rows.Count - 1 To 0 Step -1
                Me.row_BOM = TBL_BOM.Rows(ij)
                If Me.row_BOM.Qty_RM <= 0 Then
                    TBL_BOM.Rows.Remove(Me.row_BOM)
                End If
            Next
            If Me.TBL_BOM.Rows.Count > 0 Then
                jk = 0
                For Each Me.row_BOM In Me.TBL_BOM.Rows
                    jk += 1
                    If Dummy = "" Then
                        Dummy = jk.ToString + ". " + Me.row_BOM.code_rm + " (" + Me.ObjCD.DigitGrouping(Me.row_BOM.Qty_RM, Me.row_BOM.Decimal_RM) + " " + Me.row_BOM.Unit_Code + ")"
                    Else
                        Dummy = Dummy + vbCrLf + jk.ToString + ". " + Me.row_BOM.code_rm + " (" + Me.ObjCD.DigitGrouping(Me.row_BOM.Qty_RM, Me.row_BOM.Decimal_RM) + " " + Me.row_BOM.Unit_Code + ")"
                    End If
                Next
                MsgBox("Not enough raw material: " + vbCrLf + Dummy + ".", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjUSER = Nothing
            DTAProductQty.Dispose()
            TBLProductQty.Dispose()
            DTAItemBundling.Dispose()
            TBLItemBundling.Dispose()
            Me.SedangAddRM = False
        End Try
    End Sub

    Public Sub KosongkanItemProcess()
        Try
            Me.dgvinput_rmreject.Rows.Clear()
            Me.dgvinput_processreject.Rows.Clear()
            Me.dgvinput_machine.Rows.Clear()
            Me.dgvinput_employee.Rows.Clear()
            Me.dgvinput_downtime.Rows.Clear()
            Me.dgvInput_SetupTime.Rows.Clear()

            Me.ObjCD.KosongkanControl(Me.tabpage_rejectdowntimeoperdata.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.SplitContainer1.Panel2.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.SplitContainer2.Panel1.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.SplitContainer2.Panel2.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.SplitContainer3.Panel1.Controls, True, True, True)
            Me.ObjCD.KosongkanControl(Me.SplitContainer3.Panel2.Controls, True, True, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtactual_setuptime_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtactual_setuptime.Click
        Try
            'Me.ObjCD.Number_Click(Me.txtactual_setuptime)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtstockbundling.GotFocus, txtremarks.GotFocus, _
      txtwonumber.GotFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnOK.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.txtreadonly.BackColor
                        Else
                            Txt.BackColor = Me.BackColorAll
                        End If
                    End If
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtactual_setuptime_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtactual_setuptime.GotFocus
        Try
            If Me.ObjCD.Is_Time_String(Me.txtactual_setuptime.Text) = True Then
                buffer_st = txtactual_setuptime.Text
            End If

            If txtactual_setuptime.Enabled = True Then
                If txtactual_setuptime.ReadOnly = True Then
                    txtactual_setuptime.BackColor = Me.txtreadonly.BackColor
                Else
                    txtactual_setuptime.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

    'Private Sub txtactual_setuptime_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtactual_setuptime.Leave
    '    Try

    '        If Me.ObjCD.Is_Time_String(Me.txtactual_setuptime.Text) = False Then
    '            MsgBox("Please input hh.mm format", MsgBoxStyle.Information, "Information")
    '            Me.txtactual_setuptime.Text = buffer_st
    '        Else
    '            If Me.ObjCD.Convert_ToMinutes(Me.txtactual_setuptime.Text) = 0 Then
    '                Me.txtactual_setuptime.Text = "00.00"
    '            End If
    '        End If
    '        Me.GetTotalActualTime()
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Private Sub txtcode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtstockbundling.LostFocus, txtremarks.LostFocus, _
    txtwonumber.LostFocus, txtactual_setuptime.LostFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnOK.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.txtreadonly.BackColor
                        Else
                            Txt.BackColor = Me.TxtPrimaryStockBundling.BackColor
                        End If
                    End If
                End If
            End If
            Txt.Text = Trim(Txt.Text).ToString.Trim.Replace("%", "").Replace("*", "").Replace("'", "").Replace(" ", "")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtwonumber_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtwonumber.KeyDown
        Try
            If e.KeyValue = 13 Then
                Me.txtidentifier_wo.Text = Me.ObjStockBundling.GetIdentifierByWo(Me.txtwonumber.Text)
                Me.TxtIdentifierJO.Text = Me.ObjStockBundling.GetIdentifierJOByWoNo(Me.txtwonumber.Text)

                If ObjStockBundling.ApakahSudahAdaWO(Me.txtidentifier_wo.Text, ClsTrans_ManufacturingOrder.EnumPilihStatusWO.Scheduled, _
                ClsTrans_ManufacturingOrder.EnumPilihStatusWO.On_Progress, Me.TxtIdentifierJO.Text) = True Then
                    If Me.ApakahCodeNoDariWOSudahAda = False Then
                        Me.IsikanCboWorkUnit(Me.txtidentifier_wo.Text, Me.TxtIdentifierJO.Text)
                        Me.isiHeaderby_IdentifierWo()
                        Me.IsiItemBy_IdentifierWo()
                        Me.SetGboItemAfter()
                        Me.txtremarks.Focus()
                    Else
                        Exit Sub
                    End If
                Else
                    Me.btnfindwo.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtwonumber_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtwonumber.Leave
        Try

            If Me.txtwonumber.Text <> "" Then
                Me.txtidentifier_wo.Text = Me.ObjStockBundling.GetIdentifierByWo(Me.txtwonumber.Text)
                Me.TxtIdentifierJO.Text = Me.ObjStockBundling.GetIdentifierJOByWoNo(Me.txtwonumber.Text)

                If ObjStockBundling.ApakahSudahAdaWO(Me.txtidentifier_wo.Text, ClsTrans_ManufacturingOrder.EnumPilihStatusWO.Scheduled, _
                ClsTrans_ManufacturingOrder.EnumPilihStatusWO.On_Progress, Me.TxtIdentifierJO.Text) = True Then
                    If Me.ApakahCodeNoDariWOSudahAda = False Then
                        Me.IsikanCboWorkUnit(Me.txtidentifier_wo.Text, Me.TxtIdentifierJO.Text)
                        Me.isiHeaderby_IdentifierWo()
                        Me.IsiItemBy_IdentifierWo()
                        Me.SetGboItemAfter()
                        Me.txtremarks.Focus()
                    Else
                        Exit Sub
                    End If
                Else
                    Me.btnfindwo.PerformClick()
                End If
            Else

                Me.txtidentifier_wo.Text = ""
                Me.TxtIdentifierJO.Text = ""
                Me.txtwonumber.Text = ""
                Me.txtprimaryprocess.Text = "-1"
                Me.txtprimaryproduct_dariwo.Text = "-1"
                Me.txtprimarywu.Text = "-1"
                Me.txtprimaryws.Text = "-1"
                Me.txtwonumber.Text = ""
                Me.txtworkunitDesc.Text = ""
                Me.txtprocess_code.Text = ""
                Me.txtprocess_desc.Text = ""
                Me.txtmo_number.Text = ""
                Me.txtjo_number.Text = ""
                Me.TxtCodeNoAfter.Text = ""
                Me.txtDescAfter.Text = ""
                Me.txtremarks.Text = ""
                Me.TxtQtyperBatchAfter.Text = 0
                isProductCode = ""
                isBalanceProduce = 0
                Me.KosongkanField(True)
                Me.KosongkanField2()
            End If
            Me.SetGboItemAfter()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub isiHeaderby_IdentifierWo()
        Try
            If Me.ApakahCodeNoDariWOSudahAda = False Then
                Me.ObjStockBundling.IsikanTblWo4Header(Me.tbl_wo_4Header, Me.txtidentifier_wo.Text, Me.TxtIdentifierJO.Text)

                If tbl_wo_4Header.Rows.Count > 0 Then
                    row_wo_4Header = tbl_wo_4Header.Rows(0)
                    With row_wo_4Header
                        Me.TxtPrimaryCodeNoAfter.Text = .Primary_Product
                        Me.txtPrimaryCategoryAfter.Text = Me.ObjProduct.GetPrimaryCategory(Me.TxtPrimaryCodeNoAfter.Text)
                        Me.txtprimaryws.Text = -1
                        Me.txtwonumber.Text = .WO_Number
                        Me.txtmo_number.Text = .MO_Number
                        Me.txtjo_number.Text = .JO_Number
                        Me.isProductCode = .Product_Code
                        Me.txtprimaryproduct_dariwo.Text = .Primary_Product
                        Me.TxtCodeNoAfter.Text = .Product_Code
                        Me.txtDescAfter.Text = .Description
                        Me.isDecimal = ._Decimal
                    End With
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiItemBy_IdentifierWo()
        Try
            If Me.ApakahCodeNoDariWOSudahAda = False Then
                Me.ObjStockBundling.IsikanTblWo4Item(Me.tbl_wo_4Item, Me.txtidentifier_wo.Text, Me.txtprimarywu.Text, Me.TxtIdentifierJO.Text)
                If tbl_wo_4Item.Rows.Count > 0 Then
                    row_wo_4Item = tbl_wo_4Item.Rows(0)
                    With row_wo_4Item
                        Me.isBalanceProduce = Me.ObjCD.DigitGrouping(.balance_produce, isDecimal)
                        If .balance_produce > 0 Then
                            Me.TxtQtyperBatchAfter.Text = Me.ObjCD.DigitGrouping(.balance_produce, isDecimal)
                        Else
                            Me.TxtQtyperBatchAfter.Text = "0"
                        End If
                        Me.txtprimaryprocess.Text = .primary_process
                        Me.txtprocess_code.Text = .Process_Code
                        Me.txtprocess_desc.Text = .Process_Desc
                        '->planning
                        Me.dtpplan_start.Value = .plan_start_time
                        Me.dtpplan_end.Value = .plan_end_time
                        Me.txtplan_setuptime.Text = Me.ObjCD.Convert_ToString(.plan_setup_duration)
                        Me.txtplan_prodrate.Text = Me.ObjCD.DigitGrouping(.plan_prod_rate, Me.isDecimal, ClsCekData.EnumTipeRounding.Down)
                        Me.txtplan_workhour.Text = Me.ObjCD.Convert_ToString(.plan_work_duration)
                        Me.txtplan_total.Text = Me.ObjCD.Convert_ToString(.plan_total_duration)
                        '->actual
                        Me.dtpactual_start.Value = .plan_start_time
                        Me.dtpactual_end.Value = .plan_end_time
                        Me.txtactual_setuptime.Text = "00.00"
                        Me.txtactual_downtime.Text = "00.00"
                        'Me.txtactual_workhour.Text = Me.ObjCD.Convert_ToString(.actual_work_duration)
                        'Me.txtactual_total.Text = Me.ObjCD.Convert_ToString(.actual_total_duration)
                        'Me.txtactual_prodrate.Text = Me.ObjCD.DigitGrouping(.actual_prod_rate, 2)
                        Me.txtactual_workhour.Text = Me.ObjCD.Convert_ToString(Me.GetWorkDuration())
                        Me.BufferWorkHour = Me.txtactual_workhour.Text
                        Me.GetTotalActualTime()
                    End With
                End If
                If baru = True Then
                    objStockBundlingItem.isidgv_machineby_primaryworkunit(Me.dgvinput_machine, Me.txtprimarywu.Text, Me)
                    objStockBundlingItem.isidgv_karyawanby_primaryworkunit(Me.dgvinput_employee, Me.txtprimarywu.Text, Me)
                End If
            End If

            'If Me.TxtPrimaryStockBundling.Text <> -1 Then
            '    Me.objStockBundlingItem.isidgv_machine(Me.dgvinput_machine, Me.TxtPrimaryStockBundling.Text, Me, Me.txtprimarywu.Text)
            '    Me.objStockBundlingItem.isidgv_karyawan(Me.dgvinput_employee, Me.TxtPrimaryStockBundling.Text, Me, Me.txtprimarywu.Text)
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnfindwo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnfindwo.Click
        Dim frmcari_ As New FrmTrans_ManufacturingOrder_WO_Cari
        Try
            frmcari_.BackColorAll = Me.BackColorAll
            frmcari_.cboCriteria.SelectedIndex = 0

            frmcari_.vStatus1 = ClsTrans_ManufacturingOrder.EnumPilihStatusWO.Scheduled
            frmcari_.vStatus2 = ClsTrans_ManufacturingOrder.EnumPilihStatusWO.Scheduled
            frmcari_.vStatus3 = ClsTrans_ManufacturingOrder.EnumPilihStatusWO.On_Progress

            If Me.txtwonumber.Text <> "" Then
                frmcari_.txtCari.Text = Me.txtwonumber.Text
            End If

            frmcari_.Focus()
            frmcari_.ShowDialog()
            If frmcari_.Identifier <> "" Then
                Me.txtidentifier_wo.Text = frmcari_.Identifier
                Me.TxtIdentifierJO.Text = frmcari_.IdentifierJO

                Me.IsikanCboWorkUnit(frmcari_.Identifier, frmcari_.IdentifierJO)
                Me.isiHeaderby_IdentifierWo()
                Me.IsiItemBy_IdentifierWo()
                Me.SetGboItemAfter()
                Me.txtremarks.Focus()
            Else
                Me.txtidentifier_wo.Text = ""
                Me.TxtIdentifierJO.Text = ""

                Me.txtwonumber.Text = ""
                Me.CboWorkUnit.DataSource = Nothing
                Me.CboWorkUnit.Items.Clear()
                Me.txtprimaryprocess.Text = "-1"
                Me.txtprimaryproduct_dariwo.Text = "-1"
                Me.txtprimarywu.Text = "-1"
                Me.txtprimaryws.Text = "-1"
                Me.txtwonumber.Text = ""
                Me.txtworkunitDesc.Text = ""
                Me.txtprocess_code.Text = ""
                Me.txtprocess_desc.Text = ""
                Me.txtmo_number.Text = ""
                Me.txtjo_number.Text = ""
                Me.TxtCodeNoAfter.Text = ""
                Me.txtDescAfter.Text = ""
                Me.txtremarks.Text = ""
                Me.TxtQtyperBatchAfter.Text = 0
                isProductCode = ""
                isBalanceProduce = 0
                Me.KosongkanField(True)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            frmcari_.Dispose()
        End Try
    End Sub

    Private Sub dgvinput_processreject_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvinput_processreject.CellBeginEdit
        vr = Me.dgvinput_processreject.Rows(e.RowIndex)
        Try
            If e.ColumnIndex = Me.colinput_processreject_rejectcode.Index Or e.ColumnIndex = Me.colinput_processreject_qty.Index Or e.ColumnIndex = Me.colinput_processreject_remarks.Index Then
                If e.ColumnIndex = Me.colinput_processreject_rejectcode.Index Then
                    Me.buffer_pj = vr.Cells(Me.colinput_processreject_rejectcode.Index).Value
                End If

                If e.ColumnIndex = Me.colinput_processreject_qty.Index Then
                    If String.IsNullOrEmpty(vr.Cells(Me.colinput_processreject_codeno.Index).Value) = False Then
                        Me.buffer_pj = vr.Cells(Me.colinput_processreject_qty.Index).Value
                    Else
                        e.Cancel = True
                    End If
                End If

                If e.ColumnIndex = Me.colinput_processreject_remarks.Index Then
                    If String.IsNullOrEmpty(Me.dgvinput_processreject.Item(Me.colinput_processreject_rejectcode.Index, e.RowIndex).Value) = True Then
                        e.Cancel = True
                    End If
                End If
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvinput_processreject_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvinput_processreject.CellEndEdit
        Dim row As DataGridViewRow = Me.dgvinput_processreject.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim primaryreject As Double = -1
        Dim RowIdx As Integer = e.RowIndex
        Dim totalqty_indgvreject As Decimal = 0
        Dim totalqty_indgvafter As Decimal = 0
        Try
            With row
                If IsNothing(Me.dgvinput_processreject.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.colinput_processreject_remarks.Index Then
                        If IsNothing(Me.dgvinput_processreject.Item(e.ColumnIndex, e.RowIndex).Value) Then
                            Me.dgvinput_processreject.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                        Me.dgvinput_processreject.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvinput_processreject.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.Replace("'", "").Replace("%", "").Replace("*", "").Trim
                    End If

                    If e.ColumnIndex = Me.colinput_processreject_rejectcode.Index Then
                        Me.dgvinput_processreject.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvinput_processreject.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        If Me.objreject.ApakahSudahAda(Me.dgvinput_processreject.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            primaryreject = Me.objreject.GetPrimary(Me.dgvinput_processreject.Item(e.ColumnIndex, e.RowIndex).Value)

                            index = Me.getindexitem_pj(primaryreject, e.RowIndex)
                            If index < 0 Then
                                Me.dgvinput_processreject.Item(Me.colinput_processreject_primary.Index, e.RowIndex).Value = primaryreject
                                Me.dgvinput_processreject.Item(Me.colinput_processreject_rejectcode.Index, e.RowIndex).Value = Me.objreject.GetCode(primaryreject)
                                Me.dgvinput_processreject.Item(Me.colinput_processreject_codeno.Index, e.RowIndex).Value = Me.isProductCode
                                Me.dgvinput_processreject.Item(Me.colinput_processreject_primaryproduct.Index, e.RowIndex).Value = Me.ObjProduct.GetPrimary(Me.isProductCode)
                                Me.dgvinput_processreject.Item(Me.colinput_processreject_qty.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping("0", Me.isDecimal)
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.dgvinput_processreject.Item(Me.colinput_processreject_rejectcode.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.buffer_pj) = True Then
                                        Me.eraserowindex_pj = e.RowIndex
                                        Me.backgroundworker_processreject.RunWorkerAsync()
                                    Else
                                        Me.dgvinput_processreject.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer_pj
                                    End If
                                    Me.ObjCD.SetNumber(Me.dgvinput_processreject)
                                    Exit Sub
                                End If
                            End If
                        Else
                            Dim ij As New Core_Module.FrmMasterRejectCari
                            ij.txtCari.Text = Me.dgvinput_processreject.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryReject <> -1 Then

                                index = Me.getindexitem_pj(ij.PrimaryReject, e.RowIndex)
                                If index < 0 Then
                                    Me.dgvinput_processreject.Item(Me.colinput_processreject_primary.Index, e.RowIndex).Value = ij.PrimaryReject
                                    Me.dgvinput_processreject.Item(Me.colinput_processreject_rejectcode.Index, e.RowIndex).Value = Me.objreject.GetCode(ij.PrimaryReject)
                                    Me.dgvinput_processreject.Item(Me.colinput_processreject_codeno.Index, e.RowIndex).Value = Me.isProductCode
                                    Me.dgvinput_processreject.Item(Me.colinput_processreject_primaryproduct.Index, e.RowIndex).Value = Me.ObjProduct.GetPrimary(Me.isProductCode)
                                    Me.dgvinput_processreject.Item(Me.colinput_processreject_qty.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping("0", Me.isDecimal)
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.buffer_pj) = True Then
                                            Me.eraserowindex_pj = e.RowIndex
                                            Me.backgroundworker_processreject.RunWorkerAsync()
                                        Else
                                            Me.dgvinput_processreject.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer_pj
                                        End If
                                        Me.ObjCD.SetNumber(Me.dgvinput_processreject)
                                        Exit Sub
                                    Else
                                        Me.dgvinput_processreject.Item(Me.colinput_processreject_primary.Index, e.RowIndex).Value = ij.PrimaryReject
                                        Me.dgvinput_processreject.Item(Me.colinput_processreject_rejectcode.Index, e.RowIndex).Value = Me.objreject.GetCode(ij.PrimaryReject)
                                        Me.dgvinput_processreject.Item(Me.colinput_processreject_codeno.Index, e.RowIndex).Value = Me.isProductCode
                                        Me.dgvinput_processreject.Item(Me.colinput_processreject_primaryproduct.Index, e.RowIndex).Value = Me.ObjProduct.GetPrimary(Me.isProductCode)
                                        Me.dgvinput_processreject.Item(Me.colinput_processreject_qty.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping("0", Me.isDecimal)
                                    End If
                                End If
                            ElseIf ij.PrimaryReject = -1 Then
                                If Me.dgvinput_processreject.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.buffer_pj) = True Then
                                        Me.eraserowindex_pj = e.RowIndex
                                        Me.backgroundworker_processreject.RunWorkerAsync()
                                    Else
                                        Me.dgvinput_processreject.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer_pj
                                    End If
                                    Me.ObjCD.SetNumber(Me.dgvinput_processreject)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCD.SetNumber(Me.dgvinput_processreject)
                    End If

                    If e.ColumnIndex = Me.colinput_processreject_qty.Index Then
                        Me.dgvinput_processreject.Item(Me.colinput_processreject_qty.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvinput_processreject.Item(Me.colinput_processreject_qty.Index, e.RowIndex).Value, Me.isDecimal, ClsCekData.EnumTipeRounding.Middlepoint)
                        If Me.ObjCD.CekBilanganPositif(.Cells(Me.colinput_processreject_qty.Index).Value) = True Then
                            'totalqty_indgvafter = Me.get_totalqty_indgvafter
                            'totalqty_indgvreject = Me.get_totalqty_indgvpj
                            'If totalqty_indgvafter >= totalqty_indgvreject Then
                            '    Me.txttotal_reject.Text = Me.ObjCD.DigitGrouping(totalqty_indgvreject, Me.isDecimal)
                            '    Me.txttotal_good.Text = Me.ObjCD.DigitGrouping(Me.get_totalqty_indgvafter - CDec(Me.txttotal_reject.Text), Me.isDecimal)
                            '    If CDec(Me.txttotal_good.Text) < 0 Then Me.txttotal_good.Text = "0"
                            '    Me.txttotal_produce.Text = Me.ObjCD.DigitGrouping(CDec(Me.txttotal_good.Text) + CDec(Me.txttotal_reject.Text), Me.isDecimal)
                            '    If CDec(Me.txttotal_good.Text) > 0 Then
                            '        Me.txtgood_perct.Text = Me.ObjCD.DigitGrouping(CDec(Me.txttotal_good.Text) / CDec(Me.txttotal_produce.Text) * 100, 2)
                            '    End If
                            'End If
                        Else
                            MsgBox("Please enter positive number for quantity data.", MsgBoxStyle.Information, "Information")
                            If IsNumeric(Me.buffer_pj) = True Then
                                .Cells(Me.colinput_processreject_qty.Index).Value = buffer_pj
                            Else
                                .Cells(Me.colinput_processreject_qty.Index).Value = Me.ObjCD.DigitGrouping("0", Me.isDecimal)
                            End If
                        End If
                    End If
                Else
                    If Me.dgvinput_processreject.Rows.Count - 1 <> e.RowIndex Then
                        If e.ColumnIndex <> Me.colinput_processreject_remarks.Index Then
                            If String.IsNullOrEmpty(Me.buffer_pj) = True Then
                                Me.eraserowindex_pj = e.RowIndex
                                Me.backgroundworker_processreject.RunWorkerAsync()
                            Else
                                Me.dgvinput_processreject.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer_pj
                            End If
                            Me.ObjCD.SetNumber(Me.dgvinput_processreject)
                            Exit Sub
                        Else
                            Me.dgvinput_processreject.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    Else
                        If e.ColumnIndex <> Me.colinput_processreject_remarks.Index Then
                            Me.dgvinput_processreject.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer_pj
                        Else
                            Me.dgvinput_processreject.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    End If
                End If
            End With
            Me.CalculateAll()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            'Me.CalculateAll()
        End Try
    End Sub

    Private Sub dgvinput_processreject_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvinput_processreject.DataError

    End Sub

    Private Sub dgvinput_processreject_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvinput_processreject.RowsAdded
        Try
            vr = Me.dgvinput_processreject.Rows(Me.dgvinput_processreject.Rows.Count - 1)
            'If dgvinput_processreject.Rows.Count > 0 Then
            vr.Cells(Me.colinput_processreject_qty.Index).Value = "0"
            vr.Cells(Me.colinput_processreject_remarks.Index).Value = ""
            'End If

            Me.ObjCD.SetNumber(Me.dgvinput_processreject, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvinput_processreject_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvinput_processreject.RowsRemoved
        Try
            'Me.txttotal_reject.Text = Me.ObjCD.DigitGrouping(Me.get_totalqty_indgvpj, Me.isDecimal)
            'Me.txttotal_good.Text = Me.ObjCD.DigitGrouping(Me.get_totalqty_indgvafter - CDec(Me.txttotal_reject.Text), Me.isDecimal)
            'If CDec(Me.txttotal_good.Text) < 0 Then Me.txttotal_good.Text = "0"
            'Me.txttotal_produce.Text = Me.ObjCD.DigitGrouping(CDec(Me.txttotal_good.Text) + CDec(Me.txttotal_reject.Text), Me.isDecimal)
            'If CDec(Me.txttotal_good.Text) > 0 Then
            '    Me.txtgood_perct.Text = Me.ObjCD.DigitGrouping(CDec(Me.txttotal_good.Text) / CDec(Me.txttotal_produce.Text) * 100, 2)
            'End If

            Me.ObjCD.SetNumber(Me.dgvinput_processreject)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.CalculateAll()
        End Try
    End Sub

    Private Sub dgvinput_processreject_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvinput_processreject.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvinput_processreject)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub backgroundworker_processreject_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles backgroundworker_processreject.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub backgroundworker_processreject_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles backgroundworker_processreject.RunWorkerCompleted
        Try
            Me.dgvinput_processreject.Rows.RemoveAt(eraserowindex_pj)
            Me.ObjCD.SetNumber(Me.dgvinput_processreject)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function getindexitem_pj(ByVal vPrimary As Double, _
   ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            '''pastikan tidak ada code yang sama
            For ij As Integer = 0 To Me.dgvinput_processreject.Rows.Count - 2
                RowDGV = Me.dgvinput_processreject.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.colinput_processreject_primary.Index).Value = vPrimary Then
                        Retval = ij
                        Exit For
                    End If
                End If
            Next
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub CellEndEditReject(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim row As DataGridViewRow = Me.dgvinput_rmreject.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim primaryreject As Double = -1
        Dim RowIdx As Integer = e.RowIndex
        Dim index_1 As Integer = -1
        Dim cek_codermreject As Boolean
        Try
            With row
                If IsNothing(Me.dgvinput_rmreject.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    Me.dgvinput_rmreject.Item(Me.colinput_rmreject_rejectcode.Index, e.RowIndex).Value = (Me.dgvinput_rmreject.Item(Me.colinput_rmreject_rejectcode.Index, e.RowIndex).Value).ToString.ToUpper
                    If e.ColumnIndex = Me.colinput_rmreject_rejectcode.Index Then
                        If (Me.dgvinput_rmreject.Item(Me.colinput_rmreject_rejectcode.Index, e.RowIndex).Value).ToString.ToUpper = Me.buffer_rmreject Then
                            Exit Sub
                        End If
                        Me.dgvinput_rmreject.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvinput_rmreject.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper

                        If Me.objreject.ApakahSudahAda(Me.dgvinput_rmreject.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            primaryreject = Me.objreject.GetPrimary(Me.dgvinput_rmreject.Item(e.ColumnIndex, e.RowIndex).Value)
                            cek_codermreject = Me.getindexitem_antaradua_code(.Cells(Me.colinput_rmreject_primaryproduct.Index).Value, primaryreject)
                            If cek_codermreject = False Then
                                Me.dgvinput_rmreject.Item(Me.colinput_rmreject_primaryreject.Index, e.RowIndex).Value = primaryreject
                                Me.dgvinput_rmreject.Item(Me.colinput_rmreject_rejectcode.Index, e.RowIndex).Value = Me.objreject.GetCode(primaryreject)
                            Else
                                MsgBox("Duplication reject code, please input another code.", MsgBoxStyle.Information, "Information")
                                Me.dgvinput_rmreject.Item(Me.colinput_rmreject_rejectcode.Index, e.RowIndex).Value = Me.buffer_rmreject
                                If Me.buffer_rmreject <> "" Then
                                    Me.dgvinput_rmreject.Item(Me.colinput_rmreject_primaryreject.Index, e.RowIndex).Value = Me.objreject.GetPrimary(Me.dgvinput_rmreject.Item(e.ColumnIndex, e.RowIndex).Value)
                                Else
                                    Me.dgvinput_rmreject.Item(Me.colinput_rmreject_primaryreject.Index, e.RowIndex).Value = "-1"
                                End If
                            End If
                        Else
                            Dim ij As New Core_Module.FrmMasterRejectCari
                            ij.txtCari.Text = Me.dgvinput_rmreject.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryReject <> -1 Then
                                cek_codermreject = Me.getindexitem_antaradua_code(.Cells(Me.colinput_rmreject_primaryproduct.Index).Value, ij.PrimaryReject)
                                If cek_codermreject = False Then
                                    Me.dgvinput_rmreject.Item(Me.colinput_rmreject_primaryreject.Index, e.RowIndex).Value = ij.PrimaryReject
                                    Me.dgvinput_rmreject.Item(Me.colinput_rmreject_rejectcode.Index, e.RowIndex).Value = Me.objreject.GetCode(ij.PrimaryReject)
                                Else
                                    MsgBox("Duplication reject code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.dgvinput_rmreject.Item(Me.colinput_rmreject_rejectcode.Index, e.RowIndex).Value = Me.buffer_rmreject
                                    If Me.buffer_rmreject <> "" Then
                                        Me.dgvinput_rmreject.Item(Me.colinput_rmreject_primaryreject.Index, e.RowIndex).Value = Me.objreject.GetPrimary(Me.dgvinput_rmreject.Item(e.ColumnIndex, e.RowIndex).Value)
                                    Else
                                        Me.dgvinput_rmreject.Item(Me.colinput_rmreject_primaryreject.Index, e.RowIndex).Value = "-1"
                                    End If
                                End If
                            ElseIf ij.PrimaryReject = -1 Then
                                If Me.buffer_rmreject <> "" Then
                                    Me.dgvinput_rmreject.Item(Me.colinput_rmreject_rejectcode.Index, e.RowIndex).Value = Me.buffer_rmreject
                                    Me.dgvinput_rmreject.Item(Me.colinput_rmreject_primaryreject.Index, e.RowIndex).Value = Me.objreject.GetPrimary(Me.dgvinput_rmreject.Item(e.ColumnIndex, e.RowIndex).Value)
                                Else
                                    Me.dgvinput_rmreject.Item(Me.colinput_rmreject_rejectcode.Index, e.RowIndex).Value = ""
                                    Me.dgvinput_rmreject.Item(Me.colinput_rmreject_primaryreject.Index, e.RowIndex).Value = "-1"
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCD.SetNumber(Me.dgvinput_rmreject)
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function getindexitem_antaradua_code(ByVal primaryproduct As Double, ByVal primaryreject As Double) As Boolean
        Try
            For Each vr As DataGridViewRow In Me.dgvinput_rmreject.Rows
                If primaryproduct = vr.Cells(Me.colinput_rmreject_primaryproduct.Index).Value Then
                    For Each ir As DataGridViewRow In Me.dgvinput_rmreject.Rows
                        If primaryreject = ir.Cells(Me.colinput_rmreject_primaryreject.Index).Value Then
                            Return True
                        End If
                    Next
                End If
            Next
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function getindex_apakahcodesesuai_dgndgvbefore(ByVal vprimaryproduct As Double) As Boolean
        Try
            For Each vr As DataGridViewRow In Me.dgvItemb4.Rows
                If String.IsNullOrEmpty(vr.Cells(Me.ColCodeNo.Index).Value) = False Then
                    If vprimaryproduct = vr.Cells(Me.ColPrimaryCodeNo.Index).Value Then
                        Return True
                    End If
                End If
            Next
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function get_totalqty_indgvpj() As Decimal
        Dim retval As Decimal = 0
        Try
            For Each vr As DataGridViewRow In Me.dgvinput_processreject.Rows
                If String.IsNullOrEmpty(vr.Cells(Me.colinput_processreject_rejectcode.Index).Value) = False Then
                    retval += vr.Cells(Me.colinput_processreject_qty.Index).Value
                End If
            Next
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function get_totalqty_indgvafter() As Decimal
        Dim retval As Decimal = 0
        Try
            For Each vr As DataGridViewRow In Me.dgvItemAfter.Rows
                If String.IsNullOrEmpty(vr.Cells(Me.ColCodeNoAfter.Index).Value) = False Then
                    If vr.Cells(Me.ColCodeNoAfter.Index).Value = Me.isProductCode Then
                        retval += vr.Cells(Me.ColQtyAfter.Index).Value
                    End If
                End If
            Next
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function ApakahCodeNoDariWOSudahAda() As Boolean
        Try
            For Each vr As DataGridViewRow In Me.dgvItemAfter.Rows
                If String.IsNullOrEmpty(vr.Cells(Me.ColCodeNoAfter.Index).Value) = False Then
                    If vr.Cells(Me.ColCodeNoAfter.Index).Value = Me.isProductCode Then
                        Return True
                    End If
                End If
            Next
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub dgvinput_downtime_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvinput_downtime.CellBeginEdit
        vr = Me.dgvinput_downtime.Rows(e.RowIndex)
        Try
            With vr
                If e.ColumnIndex = Me.colinput_downtime_code.Index Or e.ColumnIndex = Me.colinput_downtime_duration.Index Or e.ColumnIndex = Me.colinput_downtime_remarks.Index Then
                    If e.ColumnIndex = Me.colinput_downtime_code.Index Then
                        buffer_downtime = .Cells(Me.colinput_downtime_code.Index).Value
                    End If

                    If e.ColumnIndex = Me.colinput_downtime_duration.Index Then
                        If String.IsNullOrEmpty(.Cells(Me.colinput_downtime_code.Index).Value) = False Then
                            If String.IsNullOrEmpty(.Cells(Me.colinput_downtime_duration.Index).Value) = False Then
                                buffer_downtime = .Cells(Me.colinput_downtime_duration.Index).Value
                            End If
                        Else
                            e.Cancel = True
                        End If
                    End If
                    If e.ColumnIndex = Me.colinput_downtime_remarks.Index Then
                        If String.IsNullOrEmpty(.Cells(Me.colinput_downtime_code.Index).Value) = True Then
                            e.Cancel = True
                        End If
                    End If
                Else
                    e.Cancel = True
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvinput_downtime_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvinput_downtime.CellEndEdit
        Dim row As DataGridViewRow = Me.dgvinput_downtime.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim primarydowntime As Double = -1
        Dim RowIdx As Integer = e.RowIndex
        Dim BuffTime As String = ""
        Try
            With row
                If IsNothing(Me.dgvinput_downtime.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.colinput_downtime_remarks.Index Then
                        If IsNothing(Me.dgvinput_downtime.Item(e.ColumnIndex, e.RowIndex).Value) Then
                            Me.dgvinput_downtime.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                        Me.dgvinput_downtime.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvinput_downtime.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.Replace("'", "").Replace("%", "").Replace("*", "").Trim
                    End If

                    If e.ColumnIndex = Me.colinput_downtime_duration.Index Then
                        If String.IsNullOrEmpty(.Cells(Me.colinput_downtime_code.Index).Value) = False Then
                            If Me.ObjCD.Is_Time_String(.Cells(Me.colinput_downtime_duration.Index).Value) = False Then
                                MsgBox("Please input hh.mm format", MsgBoxStyle.Information, "Information")
                                .Cells(Me.colinput_downtime_duration.Index).Value = Me.buffer_downtime
                                Exit Sub
                            Else
                                If Me.ObjCD.Convert_ToMinutes(.Cells(Me.ColST_Duration.Index).Value) <= 0 Then
                                    MsgBox("Please enter value for duration > 00.00", MsgBoxStyle.Information, "Information")
                                    Exit Sub
                                End If

                                If Me.ApakahEditSah_DownTime = True Then
                                    BuffTime = Me.txtactual_downtime.Text
                                    'Me.GetTotalDownTime()
                                    'Me.GetTotalWorkHour()
                                    'Me.GetTotalProdRate()
                                Else
                                    MsgBox("Down time duration must be less than total duration.", MsgBoxStyle.Information, "Information")
                                    .Cells(Me.colinput_downtime_duration.Index).Value = Me.buffer_downtime
                                    Exit Sub
                                End If
                            End If
                        Else
                            If String.IsNullOrEmpty(Me.buffer_downtime) = True Then
                                Me.eraserowindex_downtime = e.RowIndex
                                Me.BackgroundWorker_downtime.RunWorkerAsync()
                            Else
                                Me.dgvinput_downtime.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer_downtime
                            End If
                            Me.ObjCD.SetNumber(Me.dgvinput_downtime)
                            Exit Sub
                        End If
                    End If

                    If e.ColumnIndex = Me.colinput_downtime_code.Index Then
                        Me.dgvinput_downtime.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvinput_downtime.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        If Me.objdowntime.ApakahSudahAda(Me.dgvinput_downtime.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            primarydowntime = Me.objdowntime.GetPrimary(Me.dgvinput_downtime.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = Me.getindexitem_downtime(primarydowntime, e.RowIndex)
                            If index < 0 Then
                                Me.dgvinput_downtime.Item(Me.colinput_downtime_primary.Index, e.RowIndex).Value = primarydowntime
                                Me.dgvinput_downtime.Item(Me.colinput_downtime_code.Index, e.RowIndex).Value = Me.objdowntime.GetCode(primarydowntime)
                                Me.dgvinput_downtime.Item(Me.colinput_downtime_duration.Index, e.RowIndex).Value = "00.00"
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.dgvinput_downtime.Item(Me.colinput_downtime_code.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.buffer_downtime) = True Then
                                        Me.eraserowindex_downtime = e.RowIndex
                                        Me.BackgroundWorker_downtime.RunWorkerAsync()
                                    Else
                                        Me.dgvinput_downtime.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer_downtime
                                    End If
                                    Me.ObjCD.SetNumber(Me.dgvinput_downtime)
                                    Exit Sub
                                End If
                            End If
                        Else

                            Dim ij As New Core_Module.FrmMasterDownTimeCari
                            ij.txtCari.Text = Me.dgvinput_downtime.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryDownTime <> -1 Then
                                index = Me.getindexitem_downtime(ij.PrimaryDownTime, e.RowIndex)
                                If index < 0 Then
                                    Me.dgvinput_downtime.Item(Me.colinput_downtime_primary.Index, e.RowIndex).Value = ij.PrimaryDownTime
                                    Me.dgvinput_downtime.Item(Me.colinput_downtime_code.Index, e.RowIndex).Value = Me.objdowntime.GetCode(ij.PrimaryDownTime)
                                    Me.dgvinput_downtime.Item(Me.colinput_downtime_duration.Index, e.RowIndex).Value = "00.00"
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.buffer_downtime) = True Then
                                            Me.eraserowindex_downtime = e.RowIndex
                                            Me.BackgroundWorker_downtime.RunWorkerAsync()
                                        Else
                                            Me.dgvinput_downtime.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer_downtime
                                        End If
                                        Me.ObjCD.SetNumber(Me.dgvinput_downtime)
                                        Exit Sub
                                    Else
                                        Me.dgvinput_downtime.Item(Me.colinput_downtime_primary.Index, e.RowIndex).Value = ij.PrimaryDownTime
                                        Me.dgvinput_downtime.Item(Me.colinput_downtime_code.Index, e.RowIndex).Value = Me.objdowntime.GetCode(ij.PrimaryDownTime)
                                        Me.dgvinput_downtime.Item(Me.colinput_downtime_duration.Index, e.RowIndex).Value = "00.00"
                                    End If
                                End If
                            ElseIf ij.PrimaryDownTime = -1 Then
                                If Me.dgvinput_downtime.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.buffer_downtime) = True Then
                                        Me.eraserowindex_downtime = e.RowIndex
                                        Me.BackgroundWorker_downtime.RunWorkerAsync()
                                    Else
                                        Me.dgvinput_downtime.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer_downtime
                                    End If
                                    Me.ObjCD.SetNumber(Me.dgvinput_downtime)
                                    Exit Sub
                                End If
                                If Me.dgvinput_downtime.Item(Me.colinput_downtime_code.Index, e.RowIndex).Value <> "" Then
                                    If String.IsNullOrEmpty(Me.buffer_downtime) = True Then
                                        Me.eraserowindex_downtime = e.RowIndex
                                        Me.BackgroundWorker_downtime.RunWorkerAsync()
                                    Else
                                        Me.dgvinput_downtime.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer_downtime
                                    End If
                                    Me.ObjCD.SetNumber(Me.dgvinput_downtime)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCD.SetNumber(Me.dgvinput_downtime)
                    End If
                Else
                    If Me.dgvinput_downtime.Rows.Count - 1 <> e.RowIndex Then
                        If e.ColumnIndex <> Me.colinput_downtime_remarks.Index Then
                            If String.IsNullOrEmpty(Me.buffer_downtime) = True Then
                                Me.eraserowindex_downtime = e.RowIndex
                                Me.BackgroundWorker_downtime.RunWorkerAsync()
                            Else
                                Me.dgvinput_downtime.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer_downtime
                            End If
                            Me.ObjCD.SetNumber(Me.dgvinput_downtime)
                            Exit Sub
                        Else
                            Me.dgvinput_downtime.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    Else
                        If e.ColumnIndex <> Me.colinput_downtime_remarks.Index Then
                            Me.dgvinput_downtime.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer_downtime
                        Else
                            Me.dgvinput_downtime.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    End If
                End If
            End With
            Me.CalculateAll()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            'Me.CalculateAll()
        End Try
    End Sub

    Private Sub dgvinput_downtime_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvinput_downtime.DataError

    End Sub

    Private Sub dgvinput_downtime_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvinput_downtime.RowsAdded
        Try

            vr = Me.dgvinput_downtime.Rows(Me.dgvinput_downtime.Rows.Count - 1)
            vr.Cells(Me.colinput_downtime_remarks.Index).Value = ""
            'If dgvinput_downtime.Rows.Count > 0 Then
            '    vr.Cells(Me.colinput_downtime_duration.Index).Value = ""
            '    'ElseIf dgvinput_downtime.Rows.Count = 1 Then
            '    '    vr.Cells(Me.colinput_downtime_duration.Index).Value = "00.00"
            'End If
            Me.ObjCD.SetNumber(Me.dgvinput_downtime, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvinput_downtime_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvinput_downtime.RowsRemoved
        Try
            'Me.GetTotalDownTime()
            'Me.GetTotalWorkHour()
            'Me.GetTotalProdRate()
            Me.ObjCD.SetNumber(Me.dgvinput_downtime)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.CalculateAll()
        End Try
    End Sub

    Private Sub dgvinput_downtime_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvinput_downtime.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvinput_downtime)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker_downtime_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker_downtime.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker_downtime_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker_downtime.RunWorkerCompleted
        Try
            Me.dgvinput_downtime.Rows.RemoveAt(eraserowindex_downtime)
            Me.ObjCD.SetNumber(Me.dgvinput_downtime)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function getindexitem_downtime(ByVal vPrimary As Double, _
    ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            '''pastikan tidak ada code yang sama
            For ij As Integer = 0 To Me.dgvinput_downtime.Rows.Count - 2
                RowDGV = Me.dgvinput_downtime.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.colinput_downtime_primary.Index).Value = vPrimary Then
                        Retval = ij
                        Exit For
                    End If
                End If
            Next
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function getindexitem_setuptime(ByVal vPrimary As Double, _
    ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            '''pastikan tidak ada code yang sama
            For ij As Integer = 0 To Me.dgvInput_SetupTime.Rows.Count - 2
                RowDGV = Me.dgvInput_SetupTime.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColST_PrimarySetupTime.Index).Value = vPrimary Then
                        Retval = ij
                        Exit For
                    End If
                End If
            Next
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub GetTotalDownTime()
        Dim retval As Double = 0
        Dim retval1 As Double = 0
        Try
            For Each vr As DataGridViewRow In Me.dgvinput_downtime.Rows
                If String.IsNullOrEmpty(vr.Cells(Me.colinput_downtime_code.Index).Value) = False Then
                    If Me.ObjCD.Convert_ToMinutes(vr.Cells(Me.colinput_downtime_duration.Index).Value) > 0 Then
                        retval += CDbl(Me.ObjCD.Convert_ToMinutes(vr.Cells(Me.colinput_downtime_duration.Index).Value))
                    End If
                End If
            Next
            If retval > 0 Then
                Me.txtactual_downtime.Text = Me.ObjCD.Convert_ToString(retval)
            Else
                Me.txtactual_downtime.Text = "00.00"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub GetTotalSetupTime()
        Dim retval As Double = 0
        Dim retval1 As Double = 0
        Try
            For Each vr As DataGridViewRow In Me.dgvInput_SetupTime.Rows
                If String.IsNullOrEmpty(vr.Cells(Me.ColST_Code.Index).Value) = False Then
                    If Me.ObjCD.Convert_ToMinutes(vr.Cells(Me.ColST_Duration.Index).Value) > 0 Then
                        retval += CDbl(Me.ObjCD.Convert_ToMinutes(vr.Cells(Me.ColST_Duration.Index).Value))
                    End If
                End If
            Next
            If retval > 0 Then
                Me.txtactual_setuptime.Text = Me.ObjCD.Convert_ToString(retval)
            Else
                Me.txtactual_setuptime.Text = "00.00"
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub GetTotalProdRate()
        Dim retval As Double = 0
        Try
            If Me.ObjCD.Convert_ToMinutes(Me.txtactual_workhour.Text) > 0 Then
                If Me.txttotal_produce.Text = "" Then Me.txttotal_produce.Text = "0"
                retval = CDec(Me.txttotal_produce.Text) / (ObjCD.Convert_ToMinutes(Me.txtactual_workhour.Text) / 60)
                If retval > 0 Then
                    Me.txtactual_prodrate.Text = ObjCD.DigitGrouping(retval, Me.isDecimal, ClsCekData.EnumTipeRounding.Down)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvinput_rmreject_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvinput_rmreject.CellBeginEdit
        vr = Me.dgvinput_rmreject.Rows(e.RowIndex)

        Try
            With vr
                If e.ColumnIndex = Me.colinput_rmreject_codeno.Index Or e.ColumnIndex = Me.colinput_rmreject_rejectcode.Index Or e.ColumnIndex = Me.colinput_rmreject_qty.Index Or e.ColumnIndex = Me.colinput_rmreject_remarks.Index Then
                    If e.ColumnIndex = Me.colinput_rmreject_codeno.Index Then
                        buffer_rmreject = .Cells(Me.colinput_rmreject_codeno.Index).Value
                    End If
                    If e.ColumnIndex = Me.colinput_rmreject_rejectcode.Index Then
                        If String.IsNullOrEmpty(.Cells(Me.colinput_rmreject_primaryproduct.Index).Value) = False Then
                            buffer_rmreject = .Cells(Me.colinput_rmreject_rejectcode.Index).Value
                        Else
                            e.Cancel = True
                        End If
                    End If
                    If e.ColumnIndex = Me.colinput_rmreject_qty.Index Then
                        If String.IsNullOrEmpty(.Cells(Me.colinput_rmreject_primaryreject.Index).Value) = False And _
                        String.IsNullOrEmpty(.Cells(Me.colinput_rmreject_primaryproduct.Index).Value) = False Then
                            buffer_rmreject = .Cells(Me.colinput_rmreject_qty.Index).Value
                        Else
                            e.Cancel = True
                        End If
                    End If

                    If e.ColumnIndex = Me.colinput_rmreject_remarks.Index Then
                        If String.IsNullOrEmpty(.Cells(Me.colinput_rmreject_rejectcode.Index).Value) = True Then
                            e.Cancel = True
                        End If
                    End If
                Else
                    e.Cancel = True
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvinput_rmreject_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvinput_rmreject.CellEndEdit
        Dim row As DataGridViewRow = Me.dgvinput_rmreject.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim primaryreject As Double = -1
        Dim primaryproduct As Double = -1
        Dim RowIdx As Integer = e.RowIndex
        Dim index_1 As Integer = -1
        Dim cek_codesamadgn_rmusage As Boolean
        Dim qty As Decimal = 0
        Try
            With row
                If IsNothing(Me.dgvinput_rmreject.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.colinput_rmreject_codeno.Index Then
                        Me.dgvinput_rmreject.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvinput_rmreject.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        If Me.ObjProduct.ApakahSudahAda(Me.dgvinput_rmreject.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            primaryproduct = Me.ObjProduct.GetPrimary(Me.dgvinput_rmreject.Item(e.ColumnIndex, e.RowIndex).Value)
                            cek_codesamadgn_rmusage = Me.getindex_apakahcodesesuai_dgndgvbefore(primaryproduct)
                            If cek_codesamadgn_rmusage = True Then
                                Me.dgvinput_rmreject.Item(Me.colinput_rmreject_primaryproduct.Index, e.RowIndex).Value = primaryproduct
                                Me.dgvinput_rmreject.Item(Me.colinput_rmreject_codeno.Index, e.RowIndex).Value = Me.ObjProduct.GetProductCode(primaryproduct)
                                Me.dgvinput_rmreject.Item(Me.colinput_rmreject_qty.Index, e.RowIndex).Value = "0"
                            Else
                                MsgBox("Product code is not raw material this wo, please input another code.", MsgBoxStyle.Information, "Information")
                                Me.dgvinput_rmreject.Item(Me.colinput_rmreject_codeno.Index, e.RowIndex).Value = ""
                                If String.IsNullOrEmpty(Me.buffer_rmreject) = True Then
                                    Me.eraserowindex_rm = e.RowIndex
                                    Me.BackgroundWorker_rmreject.RunWorkerAsync()
                                Else
                                    Me.dgvinput_rmreject.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer_rmreject
                                End If
                                Me.ObjCD.SetNumber(Me.dgvinput_rmreject)
                                Exit Sub
                            End If
                        Else

                            Dim ij As New Core_Module.frmCariMasterProduct
                            ij.txtCari.Text = Me.dgvinput_rmreject.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryProduct <> 0 Then
                                cek_codesamadgn_rmusage = Me.getindex_apakahcodesesuai_dgndgvbefore(ij.PrimaryProduct)
                                If cek_codesamadgn_rmusage = True Then
                                    Me.dgvinput_rmreject.Item(Me.colinput_rmreject_primaryproduct.Index, e.RowIndex).Value = ij.PrimaryProduct
                                    Me.dgvinput_rmreject.Item(Me.colinput_rmreject_codeno.Index, e.RowIndex).Value = Me.ObjProduct.GetProductCode(ij.PrimaryProduct)
                                    Me.dgvinput_rmreject.Item(Me.colinput_rmreject_qty.Index, e.RowIndex).Value = "0"
                                Else
                                    MsgBox("Product code is not raw material this work order, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.dgvinput_rmreject.Item(Me.colinput_rmreject_codeno.Index, e.RowIndex).Value = "0"
                                    If String.IsNullOrEmpty(Me.buffer_rmreject) = True Then
                                        Me.eraserowindex_rm = e.RowIndex
                                        Me.BackgroundWorker_rmreject.RunWorkerAsync()
                                    Else
                                        Me.dgvinput_rmreject.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer_rmreject
                                    End If
                                    Me.ObjCD.SetNumber(Me.dgvinput_rmreject)
                                    Exit Sub
                                End If
                            ElseIf ij.PrimaryProduct = -1 Then
                                If Me.dgvinput_rmreject.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.buffer_rmreject) = True Then
                                        Me.eraserowindex_rm = e.RowIndex
                                        Me.BackgroundWorker_rmreject.RunWorkerAsync()
                                    Else
                                        Me.dgvinput_rmreject.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer_rmreject
                                    End If
                                    Me.ObjCD.SetNumber(Me.dgvinput_rmreject)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCD.SetNumber(Me.dgvinput_rmreject)
                    ElseIf e.ColumnIndex = Me.colinput_rmreject_rejectcode.Index Then
                        Me.CellEndEditReject(sender, e)
                    ElseIf e.ColumnIndex = Me.colinput_rmreject_qty.Index Then
                        If ObjCD.CekBilanganPositif(Me.dgvinput_rmreject(Me.colinput_rmreject_qty.Index, e.RowIndex).Value) = False Then
                            MsgBox("Please enter positive number for quantity data.", MsgBoxStyle.Information, _
                                             "Information")
                            Me.dgvinput_rmreject(Me.colinput_rmreject_qty.Index, e.RowIndex).Value = Me.buffer_rmreject
                        End If
                    End If
                Else
                    If Me.dgvinput_rmreject.Rows.Count - 1 <> e.RowIndex Then
                        If e.ColumnIndex <> Me.colinput_rmreject_remarks.Index And e.ColumnIndex <> Me.colinput_rmreject_rejectcode.Index Then
                            If Me.dgvinput_rmreject.Rows.Count - 1 <> e.RowIndex Then
                                If String.IsNullOrEmpty(Me.buffer_rmreject) = True Then
                                    Me.eraserowindex_rm = e.RowIndex
                                    Me.BackgroundWorker_rmreject.RunWorkerAsync()
                                Else
                                    Me.dgvinput_rmreject.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer_rmreject
                                End If
                                Me.ObjCD.SetNumber(Me.dgvinput_rmreject)
                                Exit Sub
                            End If
                        Else
                            If e.ColumnIndex = Me.colinput_rmreject_remarks.Index Then
                                Me.dgvinput_rmreject.Item(e.ColumnIndex, e.RowIndex).Value = ""
                            End If
                            If e.ColumnIndex = Me.colinput_rmreject_rejectcode.Index Then
                                Me.dgvinput_rmreject.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer_rmreject
                            End If
                        End If
                    Else
                        If e.ColumnIndex <> Me.colinput_rmreject_remarks.Index Then
                            Me.dgvinput_rmreject.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer_rmreject
                        Else
                            Me.dgvinput_rmreject.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvinput_rmreject_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvinput_rmreject.RowsAdded
        Try
            vr = Me.dgvinput_rmreject.Rows(Me.dgvinput_rmreject.Rows.Count - 1)
            'If dgvinput_rmreject.Rows.Count > 0 Then
            vr.Cells(Me.colinput_rmreject_qty.Index).Value = "0"
            vr.Cells(Me.colinput_rmreject_remarks.Index).Value = ""
            'End If
            Me.ObjCD.SetNumber(Me.dgvinput_rmreject, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvinput_rmreject_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvinput_rmreject.RowsRemoved
        Try
            Me.ObjCD.SetNumber(Me.dgvinput_rmreject)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvinput_rmreject_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvinput_rmreject.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvinput_rmreject)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function getindexitem_rm(ByVal vPrimary As Double, _
   ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.dgvinput_rmreject.Rows.Count - 2
                RowDGV = Me.dgvinput_rmreject.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.colinput_rmreject_primaryproduct.Index).Value = vPrimary Then
                        Retval = ij
                        Exit For
                    End If
                End If
            Next
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function getindexitem_rm_1(ByVal vPrimary As Double, _
   ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.dgvinput_rmreject.Rows.Count - 2
                RowDGV = Me.dgvinput_rmreject.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.colinput_rmreject_primaryreject.Index).Value = vPrimary Then
                        Retval = ij
                        Exit For
                    End If
                End If
            Next
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub BackgroundWorker_rmreject_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker_rmreject.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker_rmreject_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker_rmreject.RunWorkerCompleted
        Try
            Me.dgvinput_rmreject.Rows.RemoveAt(eraserowindex_rm)
            Me.ObjCD.SetNumber(Me.dgvinput_rmreject)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvinput_machine_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvinput_machine.CellBeginEdit
        vr = Me.dgvinput_machine.Rows(e.RowIndex)
        Try
            With vr
                If e.ColumnIndex = Me.colinput_machine_code.Index Or e.ColumnIndex = Me.colinput_machine_remarks.Index Then
                    If e.ColumnIndex = Me.colinput_machine_code.Index Then
                        Me.buffer_machine = .Cells(Me.colinput_machine_code.Index).Value
                    End If

                    If e.ColumnIndex = Me.colinput_machine_remarks.Index Then
                        If String.IsNullOrEmpty(.Cells(Me.colinput_machine_code.Index).Value) = True Then
                            e.Cancel = True
                        End If
                    End If
                Else
                    e.Cancel = True
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvinput_machine_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvinput_machine.CellEndEdit
        Dim row As DataGridViewRow = Me.dgvinput_machine.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim primarymachine As Double = -1
        Dim RowIdx As Integer = e.RowIndex
        Try
            With row
                If IsNothing(Me.dgvinput_machine.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.colinput_machine_code.Index Then
                        Me.dgvinput_machine.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvinput_machine.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        If Me.objmachine.ApakahSudahAda(Me.dgvinput_machine.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            primarymachine = Me.objmachine.GetPrimary(Me.dgvinput_machine.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = Me.getindexitem_machine(primarymachine, e.RowIndex)
                            If index < 0 Then
                                Me.dgvinput_machine.Item(Me.colinput_machine_primary.Index, e.RowIndex).Value = primarymachine
                                Me.dgvinput_machine.Item(Me.colinput_machine_code.Index, e.RowIndex).Value = Me.objmachine.GetCode(primarymachine)
                                Me.dgvinput_machine.Item(Me.colinput_machine_desc.Index, e.RowIndex).Value = Me.objmachine.GetDesc(primarymachine)
                                Me.dgvinput_machine.Item(Me.colinput_machine_prodindex.Index, e.RowIndex).Value = "0"
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.dgvinput_machine.Item(Me.colinput_machine_code.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.buffer_machine) = True Then
                                        Me.eraserowindex_machine = e.RowIndex
                                        Me.BackgroundWorker_machine.RunWorkerAsync()
                                    Else
                                        Me.dgvinput_machine.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer_machine
                                    End If
                                    Me.ObjCD.SetNumber(Me.dgvinput_machine)
                                    Exit Sub
                                End If
                            End If
                        Else

                            Dim ij As New Core_Module.FrmMasterMachineDataCari
                            ij.txtCari.Text = Me.dgvinput_machine.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryMachineData <> -1 Then
                                index = Me.getindexitem_machine(ij.PrimaryMachineData, e.RowIndex)
                                If index < 0 Then
                                    Me.dgvinput_machine.Item(Me.colinput_machine_primary.Index, e.RowIndex).Value = ij.PrimaryMachineData
                                    Me.dgvinput_machine.Item(Me.colinput_machine_code.Index, e.RowIndex).Value = Me.objmachine.GetCode(ij.PrimaryMachineData)
                                    Me.dgvinput_machine.Item(Me.colinput_machine_desc.Index, e.RowIndex).Value = Me.objmachine.GetDesc(ij.PrimaryMachineData)
                                    Me.dgvinput_machine.Item(Me.colinput_machine_prodindex.Index, e.RowIndex).Value = "0"
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.buffer_machine) = True Then
                                            Me.eraserowindex_machine = e.RowIndex
                                            Me.BackgroundWorker_machine.RunWorkerAsync()
                                        Else
                                            Me.dgvinput_machine.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer_machine
                                        End If
                                        Me.ObjCD.SetNumber(Me.dgvinput_machine)
                                        Exit Sub
                                    Else
                                        Me.dgvinput_machine.Item(Me.colinput_machine_primary.Index, e.RowIndex).Value = ij.PrimaryMachineData
                                        Me.dgvinput_machine.Item(Me.colinput_machine_code.Index, e.RowIndex).Value = Me.objmachine.GetCode(ij.PrimaryMachineData)
                                        Me.dgvinput_machine.Item(Me.colinput_machine_desc.Index, e.RowIndex).Value = Me.objmachine.GetDesc(ij.PrimaryMachineData)
                                        Me.dgvinput_machine.Item(Me.colinput_machine_prodindex.Index, e.RowIndex).Value = "0"
                                    End If
                                End If
                            ElseIf ij.PrimaryMachineData = -1 Then
                                If Me.dgvinput_machine.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.buffer_machine) = True Then
                                        Me.eraserowindex_machine = e.RowIndex
                                        Me.BackgroundWorker_machine.RunWorkerAsync()
                                    Else
                                        Me.dgvinput_machine.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer_machine
                                    End If
                                    Me.ObjCD.SetNumber(Me.dgvinput_machine)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCD.SetNumber(Me.dgvinput_machine)
                    End If
                Else
                    If Me.dgvinput_machine.Rows.Count - 1 <> e.RowIndex Then
                        If e.ColumnIndex <> Me.colinput_machine_remarks.Index Then
                            If String.IsNullOrEmpty(Me.buffer_machine) = True Then
                                Me.eraserowindex_machine = e.RowIndex
                                Me.BackgroundWorker_machine.RunWorkerAsync()
                            Else
                                Me.dgvinput_machine.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer_machine
                            End If
                            Me.ObjCD.SetNumber(Me.dgvinput_machine)
                            Exit Sub
                        Else
                            Me.dgvinput_machine.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    Else
                        If e.ColumnIndex <> Me.colinput_machine_remarks.Index Then
                            Me.dgvinput_machine.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer_machine
                        Else
                            Me.dgvinput_machine.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    End If

                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvinput_machine_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvinput_machine.RowsAdded

        Try
            vr = Me.dgvinput_machine.Rows(Me.dgvinput_machine.Rows.Count - 1)
            vr.Cells(Me.colinput_machine_remarks.Index).Value = ""
            Me.ObjCD.SetNumber(Me.dgvinput_machine, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

    Private Sub dgvinput_machine_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvinput_machine.RowsRemoved
        Try
            Me.ObjCD.SetNumber(Me.dgvinput_machine)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvinput_machine_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvinput_machine.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvinput_machine)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function getindexitem_machine(ByVal vPrimary As Double, _
  ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.dgvinput_machine.Rows.Count - 2
                RowDGV = Me.dgvinput_machine.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.colinput_machine_primary.Index).Value = vPrimary Then
                        Retval = ij
                        Exit For
                    End If
                End If
            Next
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub BackgroundWorker_machine_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker_machine.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker_machine_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker_machine.RunWorkerCompleted
        Try
            Me.dgvinput_machine.Rows.RemoveAt(eraserowindex_machine)
            Me.ObjCD.SetNumber(Me.dgvinput_machine)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvinput_employee_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvinput_employee.CellBeginEdit
        vr = Me.dgvinput_employee.Rows(e.RowIndex)
        Try
            With vr
                If e.ColumnIndex = Me.colinput_employee_code.Index Or e.ColumnIndex = Me.colinput_employee_remarks.Index Then
                    If e.ColumnIndex = Me.colinput_employee_code.Index Then
                        Me.buffer_employee = .Cells(Me.colinput_employee_code.Index).Value
                    End If

                    If e.ColumnIndex = Me.colinput_employee_name.Index Then
                        If String.IsNullOrEmpty(.Cells(Me.colinput_employee_code.Index).Value) = False Then
                            Me.buffer_employee = .Cells(Me.colinput_employee_name.Index).Value
                        End If
                    End If

                    If e.ColumnIndex = Me.colinput_employee_remarks.Index Then
                        If String.IsNullOrEmpty(.Cells(Me.colinput_employee_code.Index).Value) = True Then
                            e.Cancel = True
                        End If
                    End If
                Else
                    e.Cancel = True
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvinput_employee_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvinput_employee.CellEndEdit
        Dim row As DataGridViewRow = Me.dgvinput_employee.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim primaryemployee As Double = -1
        Dim RowIdx As Integer = e.RowIndex
        Try
            With row
                If IsNothing(Me.dgvinput_employee.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.colinput_employee_code.Index Then
                        Me.dgvinput_employee.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvinput_employee.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper

                        If Me.objemployee.ApakahSudahAda(Me.dgvinput_employee.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            primaryemployee = Me.objemployee.GetPrimaryByCode(Me.dgvinput_employee.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = Me.getindexitem_employee(primaryemployee, e.RowIndex)
                            If index < 0 Then
                                Me.dgvinput_employee.Item(Me.colinput_employee_primary.Index, e.RowIndex).Value = primaryemployee
                                Me.dgvinput_employee.Item(Me.colinput_employee_code.Index, e.RowIndex).Value = Me.objemployee.GetCodeByPrimary(primaryemployee)
                                Me.dgvinput_employee.Item(Me.colinput_employee_name.Index, e.RowIndex).Value = Me.objemployee.GetFullNameByPrimary(primaryemployee)
                                Me.dgvinput_employee.Item(Me.colinput_employee_prodindex.Index, e.RowIndex).Value = "0"
                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.dgvinput_employee.Item(Me.colinput_employee_code.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.buffer_employee) = True Then
                                        Me.eraserowindex_employee = e.RowIndex
                                        Me.BackgroundWorker_employee.RunWorkerAsync()
                                    Else
                                        Me.dgvinput_employee.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer_employee
                                    End If
                                    Me.ObjCD.SetNumber(Me.dgvinput_employee)
                                    Exit Sub
                                End If
                            End If
                        Else

                            Dim ij As New Core_Module.FrmMasterKaryawanCari
                            ij.txtCari.Text = Me.dgvinput_employee.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.PrimaryEmploye <> -1 Then
                                index = Me.getindexitem_machine(ij.PrimaryEmploye, e.RowIndex)
                                If index < 0 Then
                                    Me.dgvinput_employee.Item(Me.colinput_employee_primary.Index, e.RowIndex).Value = ij.PrimaryEmploye
                                    Me.dgvinput_employee.Item(Me.colinput_employee_code.Index, e.RowIndex).Value = Me.objemployee.GetCodeByPrimary(ij.PrimaryEmploye)
                                    Me.dgvinput_employee.Item(Me.colinput_employee_name.Index, e.RowIndex).Value = Me.objemployee.GetFullNameByPrimary(ij.PrimaryEmploye)
                                    Me.dgvinput_employee.Item(Me.colinput_employee_prodindex.Index, e.RowIndex).Value = "0"
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.buffer_employee) = True Then
                                            Me.eraserowindex_employee = e.RowIndex
                                            Me.BackgroundWorker_employee.RunWorkerAsync()
                                        Else
                                            Me.dgvinput_employee.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer_employee
                                        End If
                                        Me.ObjCD.SetNumber(Me.dgvinput_employee)
                                        Exit Sub
                                    Else
                                        Me.dgvinput_employee.Item(Me.colinput_employee_primary.Index, e.RowIndex).Value = ij.PrimaryEmploye
                                        Me.dgvinput_employee.Item(Me.colinput_employee_code.Index, e.RowIndex).Value = Me.objemployee.GetCodeByPrimary(ij.PrimaryEmploye)
                                        Me.dgvinput_employee.Item(Me.colinput_employee_name.Index, e.RowIndex).Value = Me.objemployee.GetFullNameByPrimary(ij.PrimaryEmploye)
                                        Me.dgvinput_employee.Item(Me.colinput_employee_prodindex.Index, e.RowIndex).Value = "0"
                                    End If
                                End If
                            ElseIf ij.PrimaryEmploye = -1 Then
                                If Me.dgvinput_employee.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.buffer_employee) = True Then
                                        Me.eraserowindex_employee = e.RowIndex
                                        Me.BackgroundWorker_employee.RunWorkerAsync()
                                    Else
                                        Me.dgvinput_employee.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer_employee
                                    End If
                                    Me.ObjCD.SetNumber(Me.dgvinput_employee)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCD.SetNumber(Me.dgvinput_employee)
                    End If
                Else
                    If Me.dgvinput_employee.Rows.Count - 1 <> e.RowIndex Then
                        If e.ColumnIndex <> Me.colinput_employee_remarks.Index Then
                            If String.IsNullOrEmpty(Me.buffer_employee) = True Then
                                Me.eraserowindex_employee = e.RowIndex
                                Me.BackgroundWorker_employee.RunWorkerAsync()
                            Else
                                Me.dgvinput_employee.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer_employee
                            End If
                            Me.ObjCD.SetNumber(Me.dgvinput_employee)
                            Exit Sub
                        Else
                            Me.dgvinput_employee.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    Else
                        If e.ColumnIndex <> Me.colinput_employee_remarks.Index Then
                            Me.dgvinput_employee.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer_employee
                        Else
                            Me.dgvinput_employee.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvinput_employee_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvinput_employee.RowsAdded

        Try
            vr = Me.dgvinput_employee.Rows(Me.dgvinput_employee.Rows.Count - 1)
            vr.Cells(Me.colinput_employee_remarks.Index).Value = ""
            Me.ObjCD.SetNumber(Me.dgvinput_employee, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

    Private Sub dgvinput_employee_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvinput_employee.RowsRemoved
        Try
            Me.ObjCD.SetNumber(Me.dgvinput_employee)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvinput_employee_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvinput_employee.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvinput_employee)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function getindexitem_employee(ByVal vPrimary As Double, _
    ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.dgvinput_employee.Rows.Count - 2
                RowDGV = Me.dgvinput_employee.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.colinput_employee_primary.Index).Value = vPrimary Then
                        Retval = ij
                        Exit For
                    End If
                End If
            Next
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub BackgroundWorker_employee_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker_employee.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker_employee_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker_employee.RunWorkerCompleted
        Try
            Me.dgvinput_employee.Rows.RemoveAt(eraserowindex_employee)
            Me.ObjCD.SetNumber(Me.dgvinput_employee)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanRawMaterial()
        Dim RowLocationPriority As Core_Module.DseMasterLocationPriority.tblmaster_location_priority1_Row
        Try
            If MsgBox("Do you want to input raw material base on produce qty automaticaly?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                Me.TBLLocationPriority.Clear()
                RowLocationPriority = Me.TBLLocationPriority.Newtblmaster_location_priority1_Row
                RowLocationPriority.primary = -1
                RowLocationPriority.user_name = Me.UserCode
                RowLocationPriority.primary_location = Me.PrimaryInputLocation
                RowLocationPriority.urutan = 1
                RowLocationPriority.tipe_transaksi = 2
                RowLocationPriority.Storage_Area_Code = Me.InputLocation
                RowLocationPriority.Storage_Area_Desc = Me.InputLocationDesc
                RowLocationPriority.Warehouse_Code = String.Empty
                RowLocationPriority.Warehouse_Desc = String.Empty
                RowLocationPriority.include_other_location = 0

                Me.TBLLocationPriority.Rows.Add(RowLocationPriority)
                If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiStockBundling_NoStrictLocation2") Then
                    Dim i As New FrmTrans_StockBundling_InputLocationPriority
                    i.TBLLocationPriority = Me.TBLLocationPriority
                    i.UserCode = Me.UserCode
                    i.BackColorAll = Me.BackColorAll
                    If i.ShowDialog() <> Windows.Forms.DialogResult.Yes Then
                        Exit Sub
                    End If
                End If
                Me.dgvItemb4.Rows.Clear()
                Me.KosongkanItemB4()
                Me.SetGboItemb4()
                If Me.tbl_wo_4Header.Rows.Count = 0 Then
                    Me.ObjStockBundling.IsikanTblWo4Header(Me.tbl_wo_4Header, Me.txtidentifier_wo.Text, Me.TxtIdentifierJO.Text)
                End If
                row_wo_4Header = Me.tbl_wo_4Header.Rows(0)
                Me.AddItemRawMaterial(row_wo_4Header.Primary_Product, Me.txtidentifier_wo.Text, Me.txttotal_produce.Text, Me.TxtIdentifierJO.Text, Me.TBLLocationPriority)
                Me.KosongkanItemB4()
                Me.SetGboItemb4()

                If Me.TxtWorkShiftCode.Text.Trim = "" Then
                    Me.TxtWorkShiftCode.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Private Sub IsikanRawMaterial()
    '    Dim ijk As New Core_Module.FrmMasterLocationPriority
    '    Dim Lanjut As Boolean = False

    '    Try
    '        If Me.txtremarks.Enabled = True And Me.dgvItemAfter.Rows.Count > 0 Then
    '            Lanjut = True

    '            -r5
    '            If MsgBox("Input raw material based on qty produce automatically?", MsgBoxStyle.YesNo, "Information") <> MsgBoxResult.Yes Then
    '                Lanjut = False
    '            End If

    '            If Lanjut = True Then
    '                If Me.ApakahAda_ItemRM_yg_LocationControl = True Then
    '                    'isikan prioritas item before
    '                    ijk.BackColorAll = Me.BackColorAll
    '                    ijk.UserCode = Me.UserCode
    '                    ijk.vTipeTransaksi = ClsMasterLocationPriority.EnumTipeTransaksi.StockBundling
    '                    If ijk.ShowDialog() = Windows.Forms.DialogResult.OK Then
    '                        Lanjut = True
    '                        Me.ObjLocationPriority.GetLocationPrioriry(Me.TBLLocationPriority, Me.UserCode, ClsMasterLocationPriority.EnumTipeTransaksi.StockBundling)
    '                    Else
    '                        Lanjut = False
    '                    End If

    '                Else
    '                    Lanjut = True
    '                End If
    '            End If


    '            If Lanjut = True Then
    '                'kosongkan item before
    '                Me.dgvItemb4.Rows.Clear()
    '                Me.KosongkanItemB4()
    '                Me.SetGboItemb4()

    '                Me.tc_processoutput.SelectedTab = Me.tabpage_rmusage

    '                -r5
    '                'For Each rowDGV As DataGridViewRow In Me.dgvItemAfter.Rows
    '                'Me.AddItemRawMaterial(rowDGV.Cells(Me.ColPrimaryCodeNoAfter.Index).Value, Me.txtidentifier_wo.Text, rowDGV.Cells(Me.ColQtyAfter.Index).Value)
    '                'Next
    '                If Me.tbl_wo_4Header.Rows.Count = 0 Then
    '                    Me.ObjStockBundling.IsikanTblWo4Header(Me.tbl_wo_4Header, Me.txtidentifier_wo.Text, Me.TxtIdentifierJO.Text)
    '                End If
    '                row_wo_4Header = Me.tbl_wo_4Header.Rows(0)
    '                Me.AddItemRawMaterial(row_wo_4Header.Primary_Product, Me.txtidentifier_wo.Text, Me.txttotal_produce.Text, Me.TxtIdentifierJO.Text, Me.TBLLocationPriority)

    '                Me.KosongkanItemB4()
    '                Me.SetGboItemb4()
    '            End If
    '        End If

    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        ijk = Nothing
    '    End Try
    'End Sub

    Public Function kurangkandgn_spl(ByVal vmulai As Double, ByVal vselesai As Double, ByVal vin As Double, ByVal vout As Double, ByVal veot As Double, ByVal vearly_delay As Double, ByVal vearly_break As Double, ByVal vdelayearly_beforebreak As Double, _
    ByVal vot As Double, ByVal vdelay As Double, ByVal vbreak As Double, ByVal vdelaybefore_break As Double, ByVal vbreak_umum As Double, ByVal breakin_duration As Double, ByRef vtotal_wh As Double) As Double
        Dim sch_in, sch_out As Double
        Dim _earlydelay, _earlybefore_break_out, _earlybefore_break_in As Double
        Dim _delay, _before_break_out, _before_break_in As Double
        Try
            sch_in = vin - vearly_delay - veot - vearly_break
            sch_out = vout + vdelay + vot + vbreak

            _earlydelay = vin - vearly_delay
            _earlybefore_break_out = sch_in + vdelayearly_beforebreak
            _earlybefore_break_in = _earlybefore_break_out + vearly_break

            _delay = vout + vdelay
            _before_break_out = sch_out - vdelaybefore_break
            _before_break_in = _before_break_out + vbreak

            If vmulai > 0 Then
                If vmulai >= sch_in And vmulai <= vin Then
                    If vmulai <= _earlydelay And vmulai >= _earlybefore_break_out Then
                        'perhitungan dari setelah masuk istirhat
                        vtotal_wh -= (_earlybefore_break_out - sch_in)
                    ElseIf vmulai <= _earlydelay And vmulai >= sch_in Then
                        vtotal_wh -= veot
                    End If
                End If
            Else
                If vselesai >= vout And vselesai <= sch_out Then
                    If vselesai >= _delay Then
                        vtotal_wh += vdelay - (_before_break_out - sch_out)
                    ElseIf vselesai >= _delay And vselesai <= sch_out Then

                    End If
                End If
            End If

            Return vtotal_wh
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub dtpactual_end_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpactual_end.GotFocus
        Try
            BufferTime_End = Me.dtpactual_end.Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpactual_end_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpactual_end.Leave
        Try
            If Me.txtremarks.Enabled Then
                Me.BackgroundWorker2.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpactual_start_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpactual_start.GotFocus
        Try
            BufferTime_Start = Me.dtpactual_start.Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpactual_start_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpactual_start.Leave
        Try
            If Me.txtremarks.Enabled Then
                Me.BackgroundWorker1.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Public Sub HitungTotalProcess()
    '    Try
    '        Me.txttotal_reject.Text = Me.ObjCD.DigitGrouping(Me.get_totalqty_indgvpj, Me.isDecimal)
    '        Me.txttotal_good.Text = Me.ObjCD.DigitGrouping(Me.get_totalqty_indgvafter - CDec(Me.txttotal_reject.Text), Me.isDecimal)
    '        If CDec(Me.txttotal_good.Text) < 0 Then Me.txttotal_good.Text = "0"
    '        Me.txttotal_produce.Text = Me.ObjCD.DigitGrouping(CDec(Me.txttotal_good.Text) + CDec(Me.txttotal_reject.Text), Me.isDecimal)
    '        If CDec(Me.txttotal_good.Text) > 0 Then
    '            Me.txtgood_perct.Text = Me.ObjCD.DigitGrouping(CDec(Me.txttotal_good.Text) / CDec(Me.txttotal_produce.Text) * 100, 2)
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Dim SetTab As Boolean
    '///Khusus [TAB] tambahan........
    Private Sub btnExit_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExit.GotFocus
        SetTab = True
    End Sub
    Private Sub btnFind_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.LostFocus
        SetTab = False
    End Sub
    Private Sub SplitContainerHeader_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles SplitContainer1.GotFocus
        If SetTab = False Then
            btnExit.Focus()
        Else
            btnFind.Focus()
        End If
    End Sub

    Private Function GetIndexWorkUnit() As Integer
        Try
            If tbl_workunit.Rows.Count > 0 And Me.PrimaryWorkUnit_ReffOfBundling <> -1 Then
                For ij As Integer = 0 To tbl_workunit.Rows.Count - 1
                    row = tbl_workunit.Rows(ij)
                    If Me.PrimaryWorkUnit_ReffOfBundling = row.primary_work_unit Then
                        Return ij
                    End If
                Next
            End If
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub IsikanCboWorkUnit(ByVal vIdentifierWo As String, ByVal vIdentifierJO As String)
        Dim Retval As Integer = -1
        Try
            Me.ObjStockBundling.IsikanTblWorkUnit(tbl_workunit, vIdentifierWo, vIdentifierJO)
            Me.tbl_workunit.DefaultView.Sort = "[work unit code] ASC"
            Me.CboWorkUnit.DataSource = Me.tbl_workunit.DefaultView
            Me.CboWorkUnit.ValueMember = "primary work unit"
            Me.CboWorkUnit.DisplayMember = "work unit code"
            Me.CboWorkUnit.Refresh()
            If Me.CboWorkUnit.Items.Count > 0 Then
                If Me.PrimaryWorkUnit_ReffOfBundling = -1 Then
                    Me.CboWorkUnit.SelectedIndex = -1
                    Me.CboWorkUnit.SelectedIndex = 0
                    row = tbl_workunit.Rows(CboWorkUnit.SelectedIndex)
                    Me.txtprimarywu.Text = row.primary_work_unit
                    Me.txtworkunitDesc.Text = row.description
                    Me.PrimaryWorkUnit_ReffOfBundling = -1
                Else
                    Retval = GetIndexWorkUnit()
                    Me.CboWorkUnit.SelectedIndex = -1
                    row = tbl_workunit.Rows(Retval)
                    Me.txtprimarywu.Text = row.primary_work_unit
                    Me.txtworkunitDesc.Text = row.description
                    Me.CboWorkUnit.SelectedIndex = Retval
                    Me.CboWorkUnit.Enabled = False
                End If
            End If


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Dim ReloadWorkUnit As Boolean = False
    Private Sub CboWorkUnit_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboWorkUnit.SelectedIndexChanged
        Me.CboWorkUnit.Refresh()
        Try
            If Me.ReloadWorkUnit = True Then
                If Me.CboWorkUnit.Items.Count > 0 And CboWorkUnit.SelectedIndex <> -1 Then
                    row = tbl_workunit.Rows(CboWorkUnit.SelectedIndex)
                    Me.txtprimarywu.Text = row.primary_work_unit
                    Me.txtworkunitDesc.Text = row.description
                    Me.isiHeaderby_IdentifierWo()
                    Me.IsiItemBy_IdentifierWo()
                    Me.SetGboItemAfter()
                    Me.LoadInpuOutputLocation(Me.TxtIdentifierJO.Text, Me.txtidentifier_wo.Text, Me.txtprimarywu.Text)
                Else
                    Me.txtprimarywu.Text = "-1"
                    Me.txtworkunitDesc.Text = ""
                End If
            Else
                Me.ReloadWorkUnit = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadInpuOutputLocation(ByVal vIdentifierJO As String, ByVal vIdentifierWO As String, ByVal vPrimaryWorkunit As Double)
        Try
            Me.ObjStockBundling.GetInputAndOutputLocation(Me.PrimaryInputLocation, Me.InputLocation, Me.InputLocationDesc, Me.PrimaryOutputLocation, Me.OutputLocation, vIdentifierJO, vIdentifierWO, vPrimaryWorkunit)
            Me.txtPrimaryLocationAfter.Text = Me.PrimaryOutputLocation
            Me.TxtLocationAfter.Text = Me.OutputLocation
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvInput_SetupTime_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvInput_SetupTime.CellBeginEdit
        vr = Me.dgvInput_SetupTime.Rows(e.RowIndex)
        Try
            With vr
                If e.ColumnIndex = Me.ColST_Code.Index Or e.ColumnIndex = Me.ColST_Duration.Index Or e.ColumnIndex = Me.ColST_Remarks.Index Then
                    If e.ColumnIndex = Me.ColST_Code.Index Then
                        buffer_setuptime = .Cells(Me.ColST_Code.Index).Value
                    End If

                    If e.ColumnIndex = Me.ColST_Duration.Index Then
                        If String.IsNullOrEmpty(.Cells(Me.ColST_Code.Index).Value) = False Then
                            If String.IsNullOrEmpty(.Cells(Me.ColST_Duration.Index).Value) = False Then
                                buffer_setuptime = .Cells(Me.ColST_Duration.Index).Value
                            End If
                        Else
                            e.Cancel = True
                        End If
                    End If
                    If e.ColumnIndex = Me.ColST_Remarks.Index Then
                        If String.IsNullOrEmpty(.Cells(Me.ColST_Code.Index).Value) = True Then
                            e.Cancel = True
                        End If
                    End If
                Else
                    e.Cancel = True
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvInput_SetupTime_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvInput_SetupTime.CellEndEdit
        Dim row As DataGridViewRow = Me.dgvInput_SetupTime.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim i As Integer = 0
        Dim primarySetupTime As Double = -1
        Dim RowIdx As Integer = e.RowIndex
        Dim BufTime As String = ""
        Dim Default_Time As Double = 0
        Try
            With row
                If IsNothing(Me.dgvInput_SetupTime.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColST_Remarks.Index Then
                        If IsNothing(Me.dgvInput_SetupTime.Item(e.ColumnIndex, e.RowIndex).Value) Then
                            Me.dgvInput_SetupTime.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                        Me.dgvInput_SetupTime.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvInput_SetupTime.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.Replace("'", "").Replace("%", "").Replace("*", "").Trim
                    End If

                    If e.ColumnIndex = Me.ColST_Duration.Index Then
                        If String.IsNullOrEmpty(.Cells(Me.ColST_Code.Index).Value) = False Then
                            If Me.ObjCD.Is_Time_String(.Cells(Me.ColST_Duration.Index).Value) = False Then
                                MsgBox("Please input in hh.mm format.", MsgBoxStyle.Information, "Information")
                                .Cells(Me.ColST_Duration.Index).Value = Me.buffer_setuptime
                                Exit Sub
                            Else
                                If Me.ObjCD.Convert_ToMinutes(.Cells(Me.ColST_Duration.Index).Value) <= 0 Then
                                    MsgBox("Please enter setup time duration.", MsgBoxStyle.Information, "Information")
                                    Exit Sub
                                End If

                                If ApakahEditSah_SetupTime() = True Then
                                    'Me.GetTotalSetupTime()
                                    'Me.GetTotalWorkHour()
                                    'Me.GetTotalProdRate()
                                Else
                                    MsgBox("Setup time duration must be less than total duration.", MsgBoxStyle.Information, "Information")
                                    .Cells(Me.ColST_Duration.Index).Value = Me.buffer_setuptime
                                    Exit Sub
                                End If
                            End If
                        Else
                            If String.IsNullOrEmpty(Me.buffer_setuptime) = True Then
                                Me.eraserowindex_setuptime = e.RowIndex
                                Me.BackgroundWorker_SetupTime.RunWorkerAsync()
                            Else
                                Me.dgvInput_SetupTime.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer_setuptime
                            End If
                            Me.ObjCD.SetNumber(Me.dgvInput_SetupTime)
                            Exit Sub
                        End If
                    End If

                    If e.ColumnIndex = Me.ColST_Code.Index Then
                        Me.dgvInput_SetupTime.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvInput_SetupTime.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        If Me.objsetupTime.ApakahSudahAda(Me.dgvInput_SetupTime.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            primarySetupTime = Me.objsetupTime.GetPrimary(Me.dgvInput_SetupTime.Item(e.ColumnIndex, e.RowIndex).Value)
                            index = Me.getindexitem_setuptime(primarySetupTime, e.RowIndex)
                            If index < 0 Then
                                Me.dgvInput_SetupTime.Item(Me.ColST_PrimarySetupTime.Index, e.RowIndex).Value = primarySetupTime
                                Me.dgvInput_SetupTime.Item(Me.ColST_Code.Index, e.RowIndex).Value = Me.objsetupTime.GetCode(primarySetupTime)
                                Default_Time = Me.ObjStockBundling.GetDefaultTime_SetupTime(.Cells(Me.ColST_PrimarySetupTime.Index).Value)
                                If Default_Time < Me.ObjCD.Convert_ToMinutes(Me.txtactual_workhour.Text) Then
                                    Me.dgvInput_SetupTime.Item(Me.ColST_Duration.Index, e.RowIndex).Value = Me.ObjCD.Convert_ToString(Default_Time)
                                    'Me.GetTotalSetupTime()
                                    'Me.GetTotalWorkHour()
                                    'Me.GetTotalProdRate()
                                Else
                                    Me.dgvInput_SetupTime.Item(Me.ColST_Duration.Index, e.RowIndex).Value = "00.00"
                                End If

                            Else
                                If e.RowIndex <> index Then
                                    MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                    Me.dgvInput_SetupTime.Item(Me.ColST_Code.Index, e.RowIndex).Value = ""
                                    If String.IsNullOrEmpty(Me.buffer_setuptime) = True Then
                                        Me.eraserowindex_setuptime = e.RowIndex
                                        Me.BackgroundWorker_SetupTime.RunWorkerAsync()
                                    Else
                                        Me.dgvInput_SetupTime.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer_setuptime
                                    End If
                                    Me.ObjCD.SetNumber(Me.dgvInput_SetupTime)
                                    Exit Sub
                                End If
                            End If
                        Else

                            Dim ij As New Core_Module.FrmMasterSetupTime_Cari
                            ij.txtCari.Text = Me.dgvInput_SetupTime.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            If ij.Primary <> -1 Then
                                index = Me.getindexitem_setuptime(ij.Primary, e.RowIndex)
                                If index < 0 Then
                                    Me.dgvInput_SetupTime.Item(Me.ColST_PrimarySetupTime.Index, e.RowIndex).Value = ij.Primary
                                    Me.dgvInput_SetupTime.Item(Me.ColST_Code.Index, e.RowIndex).Value = Me.objsetupTime.GetCode(ij.Primary)
                                    Default_Time = Me.ObjStockBundling.GetDefaultTime_SetupTime(.Cells(Me.ColST_PrimarySetupTime.Index).Value)
                                    If Default_Time < Me.ObjCD.Convert_ToMinutes(Me.txtactual_workhour.Text) Then
                                        Me.dgvInput_SetupTime.Item(Me.ColST_Duration.Index, e.RowIndex).Value = Me.ObjCD.Convert_ToString(Default_Time)
                                        'Me.GetTotalSetupTime()
                                        'Me.GetTotalWorkHour()
                                        'Me.GetTotalProdRate()
                                    Else
                                        Me.dgvInput_SetupTime.Item(Me.ColST_Duration.Index, e.RowIndex).Value = "00.00"
                                    End If
                                Else
                                    If e.RowIndex <> index Then
                                        MsgBox("Duplication code, please input another code.", MsgBoxStyle.Information, "Information")
                                        If String.IsNullOrEmpty(Me.buffer_setuptime) = True Then
                                            Me.eraserowindex_setuptime = e.RowIndex
                                            Me.BackgroundWorker_SetupTime.RunWorkerAsync()
                                        Else
                                            Me.dgvInput_SetupTime.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer_setuptime
                                        End If
                                        Me.ObjCD.SetNumber(Me.dgvInput_SetupTime)
                                        Exit Sub
                                    Else
                                        Me.dgvInput_SetupTime.Item(Me.ColST_PrimarySetupTime.Index, e.RowIndex).Value = ij.Primary
                                        Me.dgvInput_SetupTime.Item(Me.ColST_Code.Index, e.RowIndex).Value = Me.objsetupTime.GetCode(ij.Primary)
                                        Default_Time = Me.ObjStockBundling.GetDefaultTime_SetupTime(.Cells(Me.ColST_PrimarySetupTime.Index).Value)
                                        If Default_Time < Me.ObjCD.Convert_ToMinutes(Me.txtactual_workhour.Text) Then
                                            Me.dgvInput_SetupTime.Item(Me.ColST_Duration.Index, e.RowIndex).Value = Me.ObjCD.Convert_ToString(Default_Time)
                                            'Me.GetTotalSetupTime()
                                            'Me.GetTotalWorkHour()
                                            'Me.GetTotalProdRate()
                                        Else
                                            Me.dgvInput_SetupTime.Item(Me.ColST_Duration.Index, e.RowIndex).Value = "00.00"
                                        End If
                                    End If
                                End If
                            ElseIf ij.Primary = -1 Then
                                If Me.dgvInput_SetupTime.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.buffer_setuptime) = True Then
                                        Me.eraserowindex_setuptime = e.RowIndex
                                        Me.BackgroundWorker_SetupTime.RunWorkerAsync()
                                    Else
                                        Me.dgvInput_SetupTime.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer_setuptime
                                    End If
                                    Me.ObjCD.SetNumber(Me.dgvInput_SetupTime)
                                    Exit Sub
                                End If
                                If Me.dgvInput_SetupTime.Item(Me.ColST_Code.Index, e.RowIndex).Value <> "" Then
                                    If String.IsNullOrEmpty(Me.buffer_setuptime) = True Then
                                        Me.eraserowindex_setuptime = e.RowIndex
                                        Me.BackgroundWorker_SetupTime.RunWorkerAsync()
                                    Else
                                        Me.dgvInput_SetupTime.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer_setuptime
                                    End If
                                    Me.ObjCD.SetNumber(Me.dgvInput_SetupTime)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCD.SetNumber(Me.dgvInput_SetupTime)
                    End If
                Else
                    If Me.dgvInput_SetupTime.Rows.Count - 1 <> e.RowIndex Then
                        If e.ColumnIndex <> Me.colinput_downtime_remarks.Index Then
                            If String.IsNullOrEmpty(Me.buffer_setuptime) = True Then
                                Me.eraserowindex_setuptime = e.RowIndex
                                Me.BackgroundWorker_SetupTime.RunWorkerAsync()
                            Else
                                Me.dgvInput_SetupTime.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer_setuptime
                            End If
                            Me.ObjCD.SetNumber(Me.dgvInput_SetupTime)
                            Exit Sub
                        Else
                            Me.dgvInput_SetupTime.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    Else
                        If e.ColumnIndex <> Me.colinput_downtime_remarks.Index Then
                            Me.dgvInput_SetupTime.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer_setuptime
                        Else
                            Me.dgvInput_SetupTime.Item(e.ColumnIndex, e.RowIndex).Value = ""
                        End If
                    End If
                End If
            End With
            Me.CalculateAll()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            'Me.CalculateAll()
        End Try
    End Sub

    Private Sub dgvInput_SetupTime_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvInput_SetupTime.RowsAdded
        Try
            vr = Me.dgvInput_SetupTime.Rows(Me.dgvInput_SetupTime.Rows.Count - 1)
            vr.Cells(Me.ColST_Remarks.Index).Value = ""
            Me.ObjCD.SetNumber(Me.dgvInput_SetupTime, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvInput_SetupTime_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvInput_SetupTime.RowsRemoved
        Try
            'Me.GetTotalSetupTime()
            'Me.GetTotalWorkHour()
            'Me.GetTotalProdRate()
            Me.ObjCD.SetNumber(Me.dgvInput_SetupTime)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.CalculateAll()
        End Try
    End Sub

    Private Sub dgvInput_SetupTime_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvInput_SetupTime.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvInput_SetupTime, True)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker_SetupTime_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker_SetupTime.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker_SetupTime_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker_SetupTime.RunWorkerCompleted
        Try
            Me.dgvInput_SetupTime.Rows.RemoveAt(eraserowindex_setuptime)
            Me.ObjCD.SetNumber(Me.dgvInput_SetupTime)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AturWoNumberdanCboWorkUnit()
        Try
            If Me.ApakahCodeNoDariWOSudahAda = True Then
                Me.CboWorkUnit.Enabled = False
                Me.CboWorkUnit.TabStop = False
                Me.txtwonumber.TabStop = False
                Me.txtwonumber.ReadOnly = True
                Me.btnfindwo.Enabled = False
                Me.btnfindwo.TabStop = False
                Me.txtwonumber.BackColor = Me.txtreadonly.BackColor
            ElseIf Me.ApakahCodeNoDariWOSudahAda = False Then
                Me.CboWorkUnit.Enabled = True
                Me.txtwonumber.ReadOnly = False
                Me.CboWorkUnit.TabStop = True
                Me.txtwonumber.TabStop = True
                Me.btnfindwo.Enabled = True
                Me.btnfindwo.TabStop = True
                Me.txtwonumber.BackColor = Me.TxtPrimaryCodeNob4.BackColor
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function Apakah_Worktime(ByVal vRefDate As Date) As Boolean
        Dim ObjSCH As New ClsTrans_Schedule
        Dim TipeBar As ClsTrans_Schedule.EnumTipeSchBar = ClsTrans_Schedule.EnumTipeSchBar.free_ot
        Dim vTBLSchBar As New DseTrans_Schedule.tbl_sch_barDataTable

        Try
            ObjSCH.LoadScheduleBar_TanpaScheduledItem(Me.dtpactual_start.Value.Date - TimeSpan.FromDays(5), Me.dtpactual_end.Value.Date + TimeSpan.FromDays(5), vTBLSchBar, Me.txtprimarywu.Text)
            Return ObjSCH.ApakahFreeTime(vRefDate, Me.txtprimarywu.Text, vTBLSchBar) 
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            vTBLSchBar.Dispose()
            ObjSCH = Nothing
        End Try
    End Function

    Private Function GetWorkDuration() As Double
        Dim ObjSCH As New ClsTrans_Schedule
        Dim vTBLSchBar As New DseTrans_Schedule.tbl_sch_barDataTable
        Dim RetVal As Double = 0
        Try
            ObjSCH.LoadScheduleBar_TanpaScheduledItem(Me.dtpactual_start.Value - TimeSpan.FromDays(10), Me.dtpactual_end.Value + TimeSpan.FromDays(10), vTBLSchBar, Me.txtprimarywu.Text)
            RetVal = ObjSCH.GetWorkDuration(Me.dtpactual_start.Value, Me.dtpactual_end.Value, Me.txtprimarywu.Text, vTBLSchBar)
            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjSCH = Nothing
            vTBLSchBar.Dispose()
        End Try
    End Function

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            If Buffer_ActualStartDate <> Me.dtpactual_start.Value Then
                Me.HitungDurasi(Me.dtpactual_start.Value, True, False)
            End If
            Buffer_ActualStartDate = Me.dtpactual_start.Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.CalculateAll()
        End Try
    End Sub

    Sub HitungDurasi(ByVal vRefDuration As Date, ByVal vStartActual As Boolean, ByVal TampilkanMsgBox As Boolean)
        Try

            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            If Me.dtpactual_start.Value > Me.dtpactual_end.Value Then
                'MsgBox("Actual start time must be less than actual end time.", MsgBoxStyle.Information, "Information")
                'If vStartActual = True Then
                '    Me.dtpactual_start.Value = Me.BufferTime_Start
                'Else
                '    Me.dtpactual_end.Value = Me.BufferTime_End
                'End If
                'Me.BufferWorkHour = Me.txtactual_workhour.Text
            Else
                If Me.Apakah_Worktime(vRefDuration) = False Then
                    If TampilkanMsgBox = True Then
                        If vStartActual = True Then
                            MsgBox(Format(Me.dtpactual_start.Value, "dd MMM yyyy HH.mm") & " is not worktime please enter a valid.", MsgBoxStyle.Information, "Information")
                            Me.dtpactual_start.Value = Me.BufferTime_Start
                        Else
                            MsgBox(Format(Me.dtpactual_end.Value, "dd MMM yyyy HH.mm") & " is not worktime please enter a valid value.", MsgBoxStyle.Information, "Information")
                            Me.dtpactual_end.Value = Me.BufferTime_End
                        End If
                        Me.BufferWorkHour = Me.txtactual_workhour.Text
                    End If
                Else
                    Dim Duration As Double = GetWorkDuration()
                    Dim Duration_ST_DT As Double = (Me.ObjCD.Convert_ToMinutes(Me.txtactual_downtime.Text) + Me.ObjCD.Convert_ToMinutes(Me.txtactual_setuptime.Text))
                    If Duration >= Duration_ST_DT Then
                        Me.txtactual_total.Text = Me.ObjCD.Convert_ToString(Duration)
                        Me.BufferWorkHour = Me.txtactual_workhour.Text
                    Else
                        'MsgBox("Total actual work time must be greather than " & Me.ObjCD.Convert_ToString(Duration_ST_DT), MsgBoxStyle.Information, "Information")
                        'If vStartActual = True Then
                        '    Me.dtpactual_start.Value = Me.BufferTime_Start
                        'Else
                        '    Me.dtpactual_end.Value = Me.BufferTime_End
                        'End If
                    End If
                End If
            End If

            'Me.GetTotalWorkHour()
            'Me.GetTotalProdRate()
            Windows.Forms.Cursor.Current = Cursors.Default
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Windows.Forms.Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub GetTotalActualTime()
        Dim retval As Double = 0
        Try
            retval = ObjCD.Convert_ToMinutes(Me.txtactual_workhour.Text) + ObjCD.Convert_ToMinutes(Me.txtactual_downtime.Text) + ObjCD.Convert_ToMinutes(Me.txtactual_setuptime.Text)
            If retval > 0 Then
                Me.txtactual_total.Text = ObjCD.Convert_ToString(retval)
            Else
                Me.txtactual_total.Text = ObjCD.Convert_ToString(0)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub GetTotalWorkHour()
        Dim WH As Double = 0
        Dim DT As Double = 0
        Try
            If Me.BufferWorkHour = "0" Then
                Me.BufferWorkHour = "00.00"
            End If
            WH = Me.ObjCD.Convert_ToMinutes(Me.txtactual_total.Text) - Me.ObjCD.Convert_ToMinutes(Me.txtactual_downtime.Text) - Me.ObjCD.Convert_ToMinutes(Me.txtactual_setuptime.Text)
            If WH > 0 Then
                Me.txtactual_workhour.Text = ObjCD.Convert_ToString(WH)
            Else
                Me.txtactual_workhour.Text = ObjCD.Convert_ToString(0)
            End If
            BufferWorkHour = Me.ObjCD.Convert_ToString(WH)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Try
            If Buffer_ActualEndDate <> Me.dtpactual_end.Value Then
                Me.HitungDurasi(Me.dtpactual_end.Value, False, False)
            End If
            Buffer_ActualEndDate = Me.dtpactual_end.Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.CalculateAll()
        End Try
    End Sub

    Private Function ApakahEditSah_SetupTime() As Boolean
        Dim Retval As Double = 0
        Dim WH As Double = 0
        Try
            For Each vr As DataGridViewRow In Me.dgvInput_SetupTime.Rows
                If String.IsNullOrEmpty(vr.Cells(Me.ColST_Code.Index).Value) = False Then
                    If Me.ObjCD.Convert_ToMinutes(vr.Cells(Me.ColST_Duration.Index).Value) > 0 Then
                        Retval += CDbl(Me.ObjCD.Convert_ToMinutes(vr.Cells(Me.ColST_Duration.Index).Value))
                    End If
                End If
            Next

            If Retval > Me.ObjCD.Convert_ToMinutes(Me.txtactual_total.Text) Then
                Return False
            Else
                WH = Me.ObjCD.Convert_ToMinutes(Me.txtactual_total.Text) - Retval - Me.ObjCD.Convert_ToMinutes(Me.txtactual_downtime.Text)
                If WH <= 0 Then
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function ApakahEditSah_DownTime() As Boolean
        Dim Retval As Double = 0
        Dim WH As Double = 0
        Try
            For Each vr As DataGridViewRow In Me.dgvinput_downtime.Rows
                If String.IsNullOrEmpty(vr.Cells(Me.colinput_downtime_code.Index).Value) = False Then
                    If Me.ObjCD.Convert_ToMinutes(vr.Cells(Me.colinput_downtime_duration.Index).Value) > 0 Then
                        Retval += CDbl(Me.ObjCD.Convert_ToMinutes(vr.Cells(Me.colinput_downtime_duration.Index).Value))
                    End If
                End If
            Next

            If Retval > Me.ObjCD.Convert_ToMinutes(Me.txtactual_total.Text) Then
                Return False
            Else
                WH = Me.ObjCD.Convert_ToMinutes(Me.txtactual_total.Text) - Retval - Me.ObjCD.Convert_ToMinutes(Me.txtactual_setuptime.Text)
                If WH <= 0 Then
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Function NavigateTabPage(ByVal vTabDest As TabPage) As Boolean
        Try
            Me.ForceNavigate = True
            Me.tc_processoutput.SelectedTab = vTabDest
            Me.tc_processoutput.Refresh()
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub btnNextTab1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNextTab1.Click
        Try
            Me.NavigateTabPage(Me.tabpage_rejectdowntimeoperdata)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnPrevTab2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrevTab2.Click
        Try

            Me.NavigateTabPage(Me.tabpage_processoutput)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnNextTab2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNextTab2.Click
        Try
            If Me.Apakah_Worktime(Me.dtpactual_start.Value) = False Then
                MsgBox(Format(Me.dtpactual_start.Value, "dd MMM yyyy, HH.mm") & " is not valid worktime.", MsgBoxStyle.Information, "Information")
                Me.dtpactual_start.Value = Me.BufferTime_Start
                Me.BufferWorkHour = Me.txtactual_workhour.Text
                Exit Sub
            End If

            If Me.Apakah_Worktime(Me.dtpactual_end.Value) = False Then
                MsgBox(Format(Me.dtpactual_end.Value, "dd MMM yyyy, HH.mm") & " is not valid worktime.", MsgBoxStyle.Information, "Information")
                Me.dtpactual_end.Value = Me.BufferTime_End
                Me.BufferWorkHour = Me.txtactual_workhour.Text
                Exit Sub
            End If

            If Me.dtpactual_start.Value > Me.dtpactual_end.Value Then
                MsgBox("Actual start time must be less than actual end time.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            Me.NavigateTabPage(Me.tabpage_rmusage)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnPrevTab3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrevTab3.Click
        Try
            Me.NavigateTabPage(Me.tabpage_rejectdowntimeoperdata)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancelTP2.Click
        Me.PerformCancel(sender, e)
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancelTP1.Click
        Me.PerformCancel(sender, e)
    End Sub

    Dim ForceNavigate As Boolean = False
    Private Sub tc_processoutput_Selecting(ByVal sender As Object, ByVal e As System.Windows.Forms.TabControlCancelEventArgs) Handles tc_processoutput.Selecting
        Try
            If Me.txtremarks.Enabled = True Then
                If e.TabPage.Name <> Me.tabpage_processoutput.Name And Me.txtremarks.Enabled = True And (Me.txtprimarywu.Text = "-1" Or Me.txtprimarywu.Text = "") Then
                    MsgBox("Please select work order and work unit first.", MsgBoxStyle.Information, "Information")
                    Me.txtwonumber.Focus()
                    e.Cancel = True
                    Exit Sub
                End If
                If e.TabPage.Name <> Me.tabpage_processoutput.Name And Me.txtremarks.Enabled = True And (get_totalqty_indgvafter() <= 0) Then
                    MsgBox("Please input product output first.", MsgBoxStyle.Information, "Information")
                    Me.BtnInputAfter.Focus()
                    e.Cancel = True
                    Exit Sub
                End If
                If Me.ForceNavigate = False Then
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.ForceNavigate = False
        End Try
    End Sub

    Private Sub btnPrev_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrev.Click
        Try
            If Me.tc_processoutput.SelectedTab.Name = Me.tabpage_rmusage.Name Then
                Me.btnPrevTab3.PerformClick()
            ElseIf Me.tc_processoutput.SelectedTab.Name = Me.tabpage_rejectdowntimeoperdata.Name Then
                Me.btnPrevTab2.PerformClick()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNext.Click
        Try
            If Me.tc_processoutput.SelectedTab.Name = Me.tabpage_processoutput.Name Then
                Me.btnNextTab1.PerformClick()
            ElseIf Me.tc_processoutput.SelectedTab.Name = Me.tabpage_rejectdowntimeoperdata.Name Then
                Me.btnNextTab2.PerformClick()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_StockBundling_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        'Try
        '    If Me._Command = EnumCommand.VIEW And Me._PrimaryBundling <> -1 Then
        '        Me.ViewLogDirect()
        '    End If
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'End Try
    End Sub

    Private Sub FrmTrans_StockBundling_MdiChildActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.MdiChildActivate
        'Try
        '    If Me._Command = EnumCommand.VIEW And Me._PrimaryBundling <> -1 Then
        '        Me.ViewLogDirect()
        '    End If
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'End Try
    End Sub

    Private Sub FrmTrans_StockBundling_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        Try
            If Me._Command = EnumCommand.VIEW And Me._PrimaryBundling <> -1 Then
                Me.ViewLogDirect()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Private Sub TxtIdentifierJO_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtIdentifierJO.TextChanged
    '    MsgBox("")
    '    Me.TxtIdentifierJO.Visible = True
    '    Me.TxtIdentifierJO.BringToFront()
    'End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Try
            If Me.txtBatchNoAfter.Visible = True And Me.txtBatchNoAfter.Enabled = True Then
                Me.txtBatchNoAfter.Focus()
            ElseIf Me.TxtLocationAfter.Visible = True And Me.TxtLocationAfter.Enabled = True Then
                Me.TxtLocationAfter.Focus()
            Else
                Me.TxtQtyperBatchAfter.Focus()
            End If
            Me.TxtWorkShiftCode.Focus()
            Me.Timer1.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnFindWorkShift_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnFindWorkShift.Click
        Dim frmFindWorkShift As New Core_Module.FrmMasterWorkShift_Cari(ClsMasterWorkShift.EnumTipeWorkshift.Production)
        Try
            frmFindWorkShift.txtCari.Text = Me.TxtWorkShiftCode.Text
            frmFindWorkShift.ShowDialog()
            If frmFindWorkShift.PrimaryWorkshift > -1 Then
                Me.TxtPrimaryWorkShift.Text = frmFindWorkShift.PrimaryWorkshift
                Me.TxtWorkShiftCode.Text = Me.ObjWorkShift.GetCodeByPrimary(frmFindWorkShift.PrimaryWorkshift)
                Me.TxtWorkShiftDesc.Text = Me.ObjWorkShift.GetDescription(frmFindWorkShift.PrimaryWorkshift)
                Me.txtremarks.Focus()
            Else
                Me.TxtPrimaryWorkShift.Text = "-1"
                Me.TxtWorkShiftCode.Text = String.Empty
                Me.TxtWorkShiftDesc.Text = String.Empty
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtWorkShiftCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtWorkShiftCode.KeyDown
        Try
            If e.KeyCode = 13 Then
                Me.TxtWorkShiftCode.Text = Me.TxtWorkShiftCode.Text.Trim.Replace("'", "").Replace("%", "").Replace("*", "")
                If Me.ObjWorkShift.ApakahSudahAda(Me.TxtWorkShiftCode.Text, ClsMasterWorkShift.EnumTipeWorkshift.Production) Then
                    Me.TxtPrimaryWorkShift.Text = Me.ObjWorkShift.GetPrimary(Me.TxtWorkShiftCode.Text, ClsMasterWorkShift.EnumTipeWorkshift.Production)
                    Me.TxtWorkShiftDesc.Text = Me.ObjWorkShift.GetDescription(Me.TxtPrimaryWorkShift.Text)
                    Me.txtremarks.Focus()
                Else
                    Me.BtnFindWorkShift.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtWorkShiftCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtWorkShiftCode.Leave
        Try
            Me.TxtWorkShiftCode.Text = Me.TxtWorkShiftCode.Text.Trim.Replace("'", "").Replace("*", "")
            If Me.TxtWorkShiftCode.Text.Trim <> "" Then
                If Me.ObjWorkShift.ApakahSudahAda(Me.TxtWorkShiftCode.Text, ClsMasterWorkShift.EnumTipeWorkshift.Production) Then
                    Me.TxtPrimaryWorkShift.Text = Me.ObjWorkShift.GetPrimary(Me.TxtWorkShiftCode.Text, ClsMasterWorkShift.EnumTipeWorkshift.Production)
                    Me.TxtWorkShiftDesc.Text = Me.ObjWorkShift.GetDescription(Me.TxtPrimaryWorkShift.Text)
                Else
                    Me.BtnFindWorkShift.PerformClick()
                End If
            Else
                Me.TxtPrimaryWorkShift.Text = "-1"
                Me.TxtWorkShiftCode.Text = String.Empty
                Me.TxtWorkShiftDesc.Text = String.Empty
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtWorkShiftCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtWorkShiftCode.GotFocus
        Try
            If Me.TxtWorkShiftCode.Enabled Then
                Me.TxtWorkShiftCode.BackColor = Me.BackColorAll
            Else
                Me.TxtWorkShiftCode.BackColor = Me.txtreadonly.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtWorkShiftCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtWorkShiftCode.LostFocus
        Try
            If Me.TxtWorkShiftCode.Enabled Then
                Me.TxtWorkShiftCode.BackColor = Color.White
            Else
                Me.TxtWorkShiftCode.BackColor = Me.txtreadonly.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_StockBundling_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Try
            If Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.Stock_Bundling, ClsTransaksiSetting.EnumNamaParameter.ProductionLogWorkShiftRequired) = "1" Then
                IsWorkShiftRequired = True
                If Me.txtremarks.Enabled Then
                    Me.TxtWorkShiftCode.Enabled = True
                    Me.BtnFindWorkShift.Enabled = True
                End If
            Else
                IsWorkShiftRequired = False
                If Me.txtremarks.Enabled Then
                    Me.TxtWorkShiftCode.Text = String.Empty
                    Me.TxtWorkShiftDesc.Text = String.Empty
                    Me.TxtPrimaryWorkShift.Text = "-1"
                    Me.TxtWorkShiftCode.Enabled = False
                    Me.BtnFindWorkShift.Enabled = False
                End If
            End If


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnPrev_LocationChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPrev.LocationChanged
        'MsgBox("")
    End Sub

    Private Sub btnStop_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnStop.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.txtremarks.Text = "" And Me.dgvItemb4.RowCount = 0 And Me.dgvItemAfter.RowCount = 0 And _
            Me.dtpDate.Value.Date = Today And Me.TxtCodeNob4.Text = "" And Me.TxtCodeNoAfter.Text = "" And _
            Me.cboBatchNob4.Text = "" And Me.txtBatchNoAfter.Text = "" And Me.cboLocationb4.Text = "" And _
            Me.TxtLocationAfter.Text = "" And Me.TxtRemarkItemb4.Text = "" And Me.TxtRemarkItemAfter.Text = "" Then
                Me.NonAktifkanDetail()
                Me.btnNew.Focus()
                Me.KosongkanField()
            Else
                'konfirmasi batal atau tidak
                hasil = MsgBox("Stop data input and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    Me._Stop = True

                    If baru = True Then
                        Me.KosongkanField()
                    Else
                        Me.Cursor = Cursors.WaitCursor
                        Me.ObjStockBundling.CariDetail(Me.TxtPrimaryStockBundling.Text, Me.dgvItemb4, Me.dgvItemAfter, Me, Me.tbl_workunit)
                        Me.SalinKeField()
                    End If
                Else
                    Exit Sub
                End If
                Me.NonAktifkanDetail()
                Me.btnNew.Focus()
            End If
            Me.Cursor = Cursors.Default
            If Me._Command <> EnumCommand.NOTHING_ And IsNothing(Me.MdiParent) Then
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class