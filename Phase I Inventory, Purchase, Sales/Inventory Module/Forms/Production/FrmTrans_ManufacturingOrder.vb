Option Explicit On
Option Strict Off

Public Class FrmTrans_ManufacturingOrder
    Public UserCode As String = "USER"
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public Baru As Boolean = False
    Public FromMP As Boolean = False
    Public TBLItemProduct As New DseTrans_ManufacturingOrder.tblitem_productDataTable
    Public NewMOFromMP As Boolean = False
    Public PrimaryMOFromMP As Double = -1
    Public IdxNow, TotalIdx As Integer
    Public StopProcess As Boolean = False
    Public TBLItemDelivery As New DseTrans_ManufacturingOrder.tbltrans_mo_itemdeliveryDataTable

    Dim objSetting As New Core_Module.ClsTransaksiSettingUser
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjMO As New ClsTrans_ManufacturingOrder
    Dim ObjMO_Item As New ClsTrans_ManufacturingOrder_Item
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim ObjUnit As New Core_Module.ClsMasterUnit
    Dim objProses As New Core_Module.ClsMasterProses
    Dim objWU As New Core_Module.ClsMasterWorkUnit
    Dim objLoc As New Core_Module.clsMasterStorageArea
    Dim ObjNumbering As New Core_Module.ClsMasterTransactionNumbering
    Dim ObjSch As New ClsTrans_Schedule
    Dim vPrimaryUser As Double = -1
    Dim ObjUser As New Core_Module.ClsUser
    Dim Buffer As String = ""
    Dim EraseRowIndex As Integer = 0
    Dim vEdit As Boolean = False
    Dim ObjMonitor As New Core_Module.ClsTransaksiMonitor
    Dim Identifier As Double = -1
    Dim ObjRandom As New Random
    Dim ObjLicense As New Core_Module.ClsLicense

    Dim TblItemBom As New DseTrans_ManufacturingOrder.tbltrans_mo_item_bomDataTable
    Dim TblProsesOrder As New DseTrans_ManufacturingOrder.tbltrans_mo_processorderDataTable
    Dim TblItemWO As New DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable
    Dim TblItemRM As New DseTrans_ManufacturingOrder.tbltrans_mo_item_rmDataTable
    Dim TblItemRMAdd As New DseTrans_ManufacturingOrder.tbltrans_mo_item_rm_addDataTable
    Dim TblItemWIP As New DseTrans_ManufacturingOrder.tbltrans_mo_item_wipDataTable
    Dim TblProses_Filter As New DseTrans_ManufacturingOrder.ProsesOrder_FilterDataTable
    Dim RowItemBom As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomRow
    Dim vReadyForSchedule As ClsTrans_ManufacturingOrder.EnumPilihSchedule
    Dim ExecuteSelectionChange As Boolean = True
    Dim ExecuteChkRM As Boolean = True
    Dim ExecuteWO As Boolean = True

    Dim TblProsesOrderTemp As DseTrans_ManufacturingOrder.Proses_Order_TempDataTable
    Dim TblJO_Temp As New DseTrans_ManufacturingOrder.JO_TempDataTable
    Dim RowJO_Temp As DseTrans_ManufacturingOrder.JO_TempRow
    Dim ListPrimaryProses_Filter As New List(Of Double)

    Dim RowItemDelivery As DseTrans_ManufacturingOrder.tbltrans_mo_itemdeliveryRow
    Dim BoM_ToolsEnabled As Boolean = False
    Dim ObjSettingGlobal As New Core_Module.ClsTransaksiSetting

    Private Sub llbFilter_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbFilter.LinkClicked
        Try
            If Me.llbFilter.Text = "<< Hide Filter" Then
                Me.SplitContainer3.Panel1Collapsed = True
                Me.llbFilter.Text = ">> Show Filter"
            Else
                Me.SplitContainer3.Panel1Collapsed = False
                Me.llbFilter.Text = "<< Hide Filter"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TGV_WO_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles TGV_WO.CellBeginEdit

        Try
            If vEdit = True Then
                If e.ColumnIndex = Me.ColWO2_Remarks.Index Then
                    Me.Buffer = Me.TGV_WO.Item(e.ColumnIndex, e.RowIndex).Value
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

    Private Sub TGV_WO_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles TGV_WO.CellEndEdit
        Try
            If IsNothing(Me.TGV_WO.Item(e.ColumnIndex, e.RowIndex).Value) Then Me.TGV_WO.Item(e.ColumnIndex, e.RowIndex).Value = String.Empty

            Me.TGV_WO.Item(e.ColumnIndex, e.RowIndex).Value = Me.TGV_WO.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Replace("'", "").Replace("#", "").Replace("*", "").Replace("`", "").Replace("$", "").Replace("%", "")
            If e.ColumnIndex = Me.ColWO2_Remarks.Index Then
                'update bila ada perubahan
                Me.TblItemWO.DefaultView.RowFilter = "[Identifier] = '" & Me.TGV_WO.Item(Me.ColWO2_IdentifierItem.Index, e.RowIndex).Value & "'"
                If Me.TblItemWO.DefaultView.Count > 0 Then
                    Dim rowItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woRow
                    rowItemWO = Me.TblItemWO.DefaultView(0).Row
                    rowItemWO.Remarks = Me.TGV_WO.Item(Me.ColWO2_Remarks.Index, e.RowIndex).Value
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TGV_WO_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles TGV_WO.CellContentClick
        Try
            If e.RowIndex <> -1 Then
                If e.ColumnIndex = Me.ColWO2_ShowDetail.Index Then
                    Dim i As New FrmTrans_ManufacturingOrder_WO
                    i.vMONo = Me.TxtNoTransaksi.Text
                    i.vJONo = ObjMO.GetJONumber_InDgvJO(Me.TGV_WO.Item(Me.ColWO2_IdentifierJO.Index, e.RowIndex).Value, Me.DgvItem, Me)
                    i.vIdentifierWO = Me.TGV_WO.Item(Me.ColWO2_IdentifierItem.Index, e.RowIndex).Value
                    i.videntifierjo = Me.TGV_WO.Item(Me.ColWO2_IdentifierJO.Index, e.RowIndex).Value
                    i.vTblItemWO = Me.TblItemWO
                    i.vTblItemRM = Me.TblItemRM
                    i.UserCode = Me.UserCode
                    If Me.btnOK.Enabled = True Then
                        i.vBolehEdit = True
                    End If
                    i.ShowDialog()
                    i.Dispose()

                    'Load ulang Tgv WO / Dgv RM
                    If Me.DgvItem.SelectedCells.Count > 0 Then
                        If String.IsNullOrEmpty(Me.DgvItem.Item(Me.ColItem_PrimaryProduct.Index, Me.DgvItem.SelectedCells.Item(0).RowIndex).Value) = False Then
                            If TabControl1.SelectedTab.Name = Me.TabPage_WO.Name Then
                                LoadDgvItemWO(Me.DgvItem.Item(Me.ColItem_Identifier.Index, Me.DgvItem.SelectedCells.Item(0).RowIndex).Value, _
                                Me.TblItemWO, Me)
                            ElseIf TabControl1.SelectedTab.Name = Me.TabPage_RM.Name Then
                                LoadDgvItemRM(Me.DgvItem.Item(Me.ColItem_Identifier.Index, Me.DgvItem.SelectedCells.Item(0).RowIndex).Value, _
                                 Me.TblItemWO, Me.TblItemRM, Me)
                            ElseIf TabControl1.SelectedTab.Name = Me.TabPage_WIP.Name Then
                                Me.LoadItemWIP()
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvRM_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DGVRM.CellBeginEdit
        Try
            If Me.Baru = True Or vEdit = True Then
                If e.ColumnIndex = Me.ColRM_Remarks.Index Or e.ColumnIndex = Me.ColRM_RequiredDate.Index Then
                    Me.Buffer = Me.DGVRM.Item(e.ColumnIndex, e.RowIndex).Value
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

    Private Sub DGVRM_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGVRM.CellClick
        Dim RowItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmRow
        Dim ObjUser As New Core_Module.ClsUser

        Try
            If e.RowIndex = -1 Then
                Exit Sub
            End If
            If e.ColumnIndex = Me.ColRM_QtyRequired.Index Then
                If CDec(Me.DGVRM.Item(Me.ColRM_QtyRequired.Index, e.RowIndex).Value) > 0 Then
                    Dim ij As New FrmTrans_ManufacturingOrder_Exp_QtyReq
                    Dim rowDGV As DataGridViewRow

                    ij.BackColorAll = Me.BackColorAll
                    ij.TblItemBom = Me.TblItemBom
                    ij.TblItemRM = Me.TblItemRM
                    ij.TblItemWO = Me.TblItemWO
                    ij.IdentifierWO = Me.DGVRM.Item(Me.ColRM_IdentifierWO.Index, e.RowIndex).Value
                    ij.IdentifierJO = Me.DGVRM.Item(Me.ColRM_IdentifierItem.Index, e.RowIndex).Value
                    ij.PrimaryRM = Me.DGVRM.Item(Me.ColRM_PrimaryCodeNo.Index, e.RowIndex).Value

                    For Each rowDGV In Me.DgvItem.Rows
                        If rowDGV.Cells(Me.ColItem_Identifier.Index).Value = ij.IdentifierJO Then
                            ij.CodeNo = rowDGV.Cells(Me.ColItem_CodeNo.Index).Value
                            ij.Desc = rowDGV.Cells(Me.ColItem_Desc.Index).Value
                            ij.JONO = rowDGV.Cells(Me.ColItem_JONumber.Index).Value

                            Exit For
                        End If
                    Next
                    ij.ShowDialog()
                End If
            ElseIf e.ColumnIndex = Me.ColRM_QtyUsed.Index Then
                If CDec(Me.DGVRM.Item(Me.ColRM_QtyUsed.Index, e.RowIndex).Value) > 0 Then
                    Dim ij As New FrmTrans_ManufacturingOrder_Exp_QtyUsed
                    ij.BackColorAll = Me.BackColorAll
                    ij.CodeNo = Me.DGVRM.Item(Me.ColRM_CodeNo.Index, e.RowIndex).Value
                    ij.Desc = Me.DGVRM.Item(Me.ColRM_Desc.Index, e.RowIndex).Value
                    ij.IdentifierJO = Me.DGVRM.Item(Me.ColRM_IdentifierItem.Index, e.RowIndex).Value
                    ij.IdentifierWO = Me.DGVRM.Item(Me.ColRM_IdentifierWO.Index, e.RowIndex).Value
                    ij.PrimaryProduct = Me.DGVRM.Item(Me.ColRM_PrimaryCodeNo.Index, e.RowIndex).Value
                    ij.PrimaryUnit = Me.DGVRM.Item(Me.ColRM_PrimaryUnit.Index, e.RowIndex).Value
                    ij.ProcessCode = Me.DGVRM.Item(Me.ColRM_Process.Index, e.RowIndex).Value
                    ij.UnitCode = Me.DGVRM.Item(Me.ColRM_Unit.Index, e.RowIndex).Value
                    ij.WONO = Me.DGVRM.Item(Me.ColRM_WONumber.Index, e.RowIndex).Value

                    ij.ShowDialog()
                End If
            ElseIf e.ColumnIndex = Me.ColRM_QtyAdditional.Index Then
                If Me.TxtRemarks.Enabled = False Then
                    If CDec(Me.DGVRM.Item(e.ColumnIndex, e.RowIndex).Value) <= 0 Then
                        Exit Sub
                    End If
                End If
                Dim ij As New FrmTrans_ManufacturingOrder_Exp_QtyRMAdd
                ij.BackColorAll = Me.BackColorAll
                ij.CodeNoRM = Me.DGVRM.Item(Me.ColRM_CodeNo.Index, e.RowIndex).Value
                ij.DescRM = Me.DGVRM.Item(Me.ColRM_Desc.Index, e.RowIndex).Value
                ij.vIdentifierJO = Me.DGVRM.Item(Me.ColRM_IdentifierItem.Index, e.RowIndex).Value
                ij.vIdentifireWO = Me.DGVRM.Item(Me.ColRM_IdentifierWO.Index, e.RowIndex).Value
                ij.Primary_Product_Code_RM = Me.DGVRM.Item(Me.ColRM_PrimaryCodeNo.Index, e.RowIndex).Value
                ij.ProcessCode = Me.DGVRM.Item(Me.ColRM_Process.Index, e.RowIndex).Value
                ij.UnitCode = Me.DGVRM.Item(Me.ColRM_Unit.Index, e.RowIndex).Value
                ij.WONumber = Me.DGVRM.Item(Me.ColRM_WONumber.Index, e.RowIndex).Value
                ij.TblItemRM = Me.TblItemRM
                ij.TblItemRMAdd = Me.TblItemRMAdd
                ij.RefDate = Me.DtpTglTransaksi.Value
                ij.PrimaryMO = Me.TxtPrimary.Text
                If Me.TxtRemarks.Enabled = False Then
                    ij.ReadOnly_ = True
                Else
                    If ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_ManufacturingOrder_InputAdditional") = False Then
                        If CDec(Me.DGVRM.Item(e.ColumnIndex, e.RowIndex).Value) <> 0 Then
                            ij.ReadOnly_ = True
                        Else
                            MsgBox("You don't have access to perform this action", MsgBoxStyle.Information, "Information")
                            Exit Sub
                        End If
                    End If
                End If
                If ij.ShowDialog = Windows.Forms.DialogResult.OK Then
                    Dim RowDGV As DataGridViewRow
                    'Dim RowItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmRow

                    For Each RowItemRM In Me.TblItemRM.Rows
                        If RowItemRM.Identifier_JO = ij.vIdentifierJO And RowItemRM.Identifier_WO = ij.vIdentifireWO And _
                        RowItemRM.Primary_Product_Code_RM = ij.Primary_Product_Code_RM Then
                            RowDGV = Me.DGVRM.Rows(e.RowIndex)
                            RowDGV.Cells(Me.ColRM_QtyAdditional.Index).Value = Me.ObjCD.DigitGrouping(RowItemRM.Additional_Qty, ij.DefPrec)
                            RowDGV.Cells(Me.ColRM_QtyBalance.Index).Value = Me.ObjCD.DigitGrouping(RowItemRM.Balance_Qty, ij.DefPrec)
                            Exit For
                        End If
                    Next
                End If
            ElseIf e.ColumnIndex = Me.ColRM_QtyReplaced.Index Then
                If Me.DGVRM.Item(Me.ColRM_IsWIP.Index, e.RowIndex).Value = 0 Then
                    If Me.TxtRemarks.Enabled = False Then
                        If CDec(Me.DGVRM.Item(e.ColumnIndex, e.RowIndex).Value) <= 0 Then
                            Exit Sub
                        End If
                    End If
                    Dim ij As New FrmTrans_ManufacturingOrder_Exp_QtyRMReplace
                    ij.BackColorAll = Me.BackColorAll
                    ij.CodeNoRM = Me.DGVRM.Item(Me.ColRM_CodeNo.Index, e.RowIndex).Value
                    ij.DescRM = Me.DGVRM.Item(Me.ColRM_Desc.Index, e.RowIndex).Value
                    ij.vIdentifierJO = Me.DGVRM.Item(Me.ColRM_IdentifierItem.Index, e.RowIndex).Value
                    ij.vIdentifireWO = Me.DGVRM.Item(Me.ColRM_IdentifierWO.Index, e.RowIndex).Value
                    ij.Primary_Product_Code_RM = Me.DGVRM.Item(Me.ColRM_PrimaryCodeNo.Index, e.RowIndex).Value
                    ij.ProcessCode = Me.DGVRM.Item(Me.ColRM_Process.Index, e.RowIndex).Value
                    ij.UnitCode = Me.DGVRM.Item(Me.ColRM_Unit.Index, e.RowIndex).Value
                    ij.WONumber = Me.DGVRM.Item(Me.ColRM_WONumber.Index, e.RowIndex).Value
                    ij.TblItemRM = Me.TblItemRM
                    ij.vQtyRequired = Me.DGVRM.Item(Me.ColRM_QtyRequired.Index, e.RowIndex).Value
                    ij.vQtyAdditional = Me.DGVRM.Item(Me.ColRM_QtyAdditional.Index, e.RowIndex).Value
                    ij.PrimaryMO = Me.TxtPrimary.Text
                    For Each RowItemRM In Me.TblItemRM.Rows
                        If RowItemRM.Primary_Product_Code_RM = Me.DGVRM.Item(Me.ColRM_PrimaryCodeNo.Index, e.RowIndex).Value And _
                        RowItemRM.Identifier_JO = Me.DGVRM.Item(Me.ColRM_IdentifierItem.Index, e.RowIndex).Value And _
                        RowItemRM.Identifier_WO = Me.DGVRM.Item(Me.ColRM_IdentifierWO.Index, e.RowIndex).Value Then
                            ij.vPriorityReplaced = RowItemRM.use_priority
                            Exit For
                        End If
                    Next

                    If Me.TxtRemarks.Enabled = False Then
                        ij.ReadOnly_ = True
                    End If
                    If ij.ShowDialog = Windows.Forms.DialogResult.OK Then
                        'update rowitem rm
                        For Each RowItemRM In Me.TblItemRM.Rows
                            If RowItemRM.Primary_Product_Code_RM = Me.DGVRM.Item(Me.ColRM_PrimaryCodeNo.Index, e.RowIndex).Value And _
                            RowItemRM.Identifier_JO = Me.DGVRM.Item(Me.ColRM_IdentifierItem.Index, e.RowIndex).Value And _
                            RowItemRM.Identifier_WO = Me.DGVRM.Item(Me.ColRM_IdentifierWO.Index, e.RowIndex).Value Then
                                Dim vDecProd_Item As Integer = Me.ObjProduct.GetAngkaDecimal(RowItemRM.Primary_Product_Code_RM)

                                RowItemRM.Replace_Qty = ij.TxtQtyReplaced.Text
                                RowItemRM.Balance_Qty = Me.ObjCD.RoundUp1(RowItemRM.Allocated_Qty + RowItemRM.Additional_Qty - _
                                RowItemRM.Used_Qty - RowItemRM.WIP_Qty - RowItemRM.Replace_Qty, vDecProd_Item)
                                RowItemRM.use_priority = ij.TxtPriority.Text

                                'update rowdgv
                                Me.DGVRM.Item(Me.ColRM_QtyBalance.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(RowItemRM.Balance_Qty, vDecProd_Item)
                                Me.DGVRM.Item(Me.ColRM_QtyReplaced.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(RowItemRM.Replace_Qty, vDecProd_Item)

                                Exit For
                            End If
                        Next

                        'update rowitem rm yang replacement
                        '1. buang baris replacement yang lama
                        For ab As Integer = Me.TblItemRM.Rows.Count - 1 To 0 Step -1
                            RowItemRM = Me.TblItemRM.Rows(ab)
                            If RowItemRM.primary_code_rm_replaced = Me.DGVRM.Item(Me.ColRM_PrimaryCodeNo.Index, e.RowIndex).Value And _
                            RowItemRM.Identifier_JO = Me.DGVRM.Item(Me.ColRM_IdentifierItem.Index, e.RowIndex).Value And _
                            RowItemRM.Identifier_WO = Me.DGVRM.Item(Me.ColRM_IdentifierWO.Index, e.RowIndex).Value And _
                            RowItemRM.is_replacement = 1 Then
                                Me.TblItemRM.Rows.Remove(RowItemRM)
                            End If
                        Next

                        '2. tambahkan baris replacement yang baru
                        For Each RowDgv As DataGridViewRow In ij.DgvItemRMReplace.Rows
                            If RowDgv.Cells(ij.ColPrimaryCodeNo.Index).Value <> -1 Then
                                RowItemRM = Me.TblItemRM.Newtbltrans_mo_item_rmRow
                                With RowItemRM
                                    .Additional_Qty = 0
                                    .Allocated_Qty = 0
                                    .Available_Qty = 0
                                    .GP_Actual = 0
                                    .GP_Target = 0
                                    .Identifier_JO = Me.DGVRM.Item(Me.ColRM_IdentifierItem.Index, e.RowIndex).Value
                                    .Identifier_WO = Me.DGVRM.Item(Me.ColRM_IdentifierWO.Index, e.RowIndex).Value
                                    .is_replacement = 1
                                    .is_wip = Me.DGVRM.Item(Me.ColRM_IsWIP.Index, e.RowIndex).Value
                                    .Primary = -1
                                    .primary_code_rm_replaced = Me.DGVRM.Item(Me.ColRM_PrimaryCodeNo.Index, e.RowIndex).Value
                                    .Primary_MO = Me.TxtPrimary.Text
                                    .Primary_PickUp_Location = -1
                                    .Primary_Product_Code_RM = RowDgv.Cells(ij.ColPrimaryCodeNo.Index).Value
                                    .Primary_StandBy_Location = -1
                                    .primary_unit_replaced = RowDgv.Cells(ij.ColPrimaryUnitReplaced.Index).Value
                                    .qty_replaced = RowDgv.Cells(ij.ColQtyReplaced.Index).Value
                                    .ratio_replaced = RowDgv.Cells(ij.ColRatioReplace.Index).Value
                                    .ratio_replacement = RowDgv.Cells(ij.ColRatioReplacement.Index).Value
                                    .Remarks = ""
                                    .Replace_Qty = 0
                                    .rm_code = RowDgv.Cells(ij.ColCodeNo.Index).Value
                                    .rm_desc = RowDgv.Cells(ij.ColDesc.Index).Value
                                    If Me.DGVRM.Item(Me.ColRM_IsWIP.Index, e.RowIndex).Value = 0 Then
                                        .Status_RM = 1
                                    Else
                                        .Status_RM = 0
                                    End If
                                    .unit_code = RowDgv.Cells(ij.ColUnit.Index).Value
                                    .unit_primary = Me.ObjProduct.GetPrimaryUnit(.Primary_Product_Code_RM)
                                    .WIP_Qty = 0
                                    ._decimal = Me.ObjProduct.GetAngkaDecimal(.Primary_Product_Code_RM)

                                    'untuk replacement hanya 4 qty yang dipake, qty required=qty target, qty used, qty balance
                                    .Required_Qty = RowDgv.Cells(ij.ColQty.Index).Value
                                    .Target_Qty = .Required_Qty
                                    .Used_Qty = 0
                                    .Balance_Qty = .Required_Qty - .Used_Qty
                                    .use_priority = RowDgv.Cells(ij.ColNo.Index).Value
                                    .required_date = Me.DGVRM.Item(Me.ColRM_RequiredDate.Index, e.RowIndex).Value
                                End With
                                Me.TblItemRM.Rows.Add(RowItemRM)
                            End If
                        Next
                    End If
                Else
                    MsgBox("Replacement feature only apllies to non-WIP raw material.", MsgBoxStyle.Information, "Information")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvRM_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGVRM.CellEndEdit
        Dim vDecProd_Item As Decimal = 0

        Try
            If IsNothing(Me.DGVRM.Item(e.ColumnIndex, e.RowIndex).Value) Then Me.DGVRM.Item(e.ColumnIndex, e.RowIndex).Value = String.Empty
            vDecProd_Item = ObjProduct.GetAngkaDecimal(Me.DGVRM.Item(Me.ColRM_PrimaryCodeNo.Index, e.RowIndex).Value)
            Me.DGVRM.Item(e.ColumnIndex, e.RowIndex).Value = Me.DGVRM.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Replace("'", "").Replace("#", "").Replace("*", "").Replace("`", "").Replace("$", "").Replace("%", "")
            If e.ColumnIndex = Me.ColRM_Remarks.Index Then
                'update bila ada perubahan
                Me.TblItemRM.DefaultView.RowFilter = "[Primary Product Code RM] = " & CDbl(Me.DGVRM.Item(Me.ColRM_PrimaryCodeNo.Index, e.RowIndex).Value).ToString & _
                " AND [Identifier WO] = '" & Me.DGVRM.Item(Me.ColRM_IdentifierWO.Index, e.RowIndex).Value.ToString & "' AND [Identifier JO] = '" + Me.DGVRM.Item(Me.ColRM_IdentifierItem.Index, e.RowIndex).Value.ToString + "'"
                If Me.TblItemRM.DefaultView.Count > 0 Then
                    Dim rowItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmRow
                    rowItemRM = Me.TblItemRM.DefaultView(0).Row
                    rowItemRM.Remarks = Me.DGVRM.Item(Me.ColRM_Remarks.Index, e.RowIndex).Value
                End If
            ElseIf e.ColumnIndex = Me.ColRM_RequiredDate.Index Then
                If IsDate(Me.DGVRM.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                    If CDate(Me.DGVRM.Item(e.ColumnIndex, e.RowIndex).Value) < Me.DtpTglTransaksi.Value.Date Then
                        MsgBox("Please specify raw material required date with minimum value: " + Format(Me.DtpTglTransaksi.Value.Date, "dd MMM yyyy"), MsgBoxStyle.Information, "Information")
                        Me.DGVRM.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Exit Sub
                    End If
                    Me.TblItemRM.DefaultView.RowFilter = "[Primary Product Code RM] = " & CDbl(Me.DGVRM.Item(Me.ColRM_PrimaryCodeNo.Index, e.RowIndex).Value).ToString & _
                    " AND [Identifier WO] = '" & Me.DGVRM.Item(Me.ColRM_IdentifierWO.Index, e.RowIndex).Value.ToString & "' AND [Identifier JO] = '" + Me.DGVRM.Item(Me.ColRM_IdentifierItem.Index, e.RowIndex).Value.ToString + "'"
                    If Me.TblItemRM.DefaultView.Count > 0 Then
                        Dim rowItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmRow
                        rowItemRM = Me.TblItemRM.DefaultView(0).Row
                        rowItemRM.required_date = CDate(Me.DGVRM.Item(e.ColumnIndex, e.RowIndex).Value)
                    End If
                Else
                    Me.DGVRM.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVRM_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles DGVRM.RowsRemoved
        Try
            Me.ObjCD.SetNumber(Me.DgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvRM_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DGVRM.Sorted
        Try
            Me.ObjCD.SetNumber(Me.DGVRM)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvRMRekap_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvRMRekap.CellClick
        Dim rowDGV As DataGridViewRow

        Try
            If e.RowIndex <> -1 Then
                If e.ColumnIndex = Me.ColRMRekap_QtyRequired.Index Or _
                e.ColumnIndex = Me.ColRMRekap_QtyTarget.Index Or _
                e.ColumnIndex = Me.ColRMRekap_QtyUsed.Index Or _
                e.ColumnIndex = Me.ColRMRekap_QtyBalance.Index Then
                    Dim ij As New FrmTrans_ManufacturingOrder_DetailRM
                    ij.BackColorAll = Me.BackColorAll
                    ij.CodeNoRM = Me.DgvRMRekap.Item(Me.ColRMRekap_CodeNo.Index, e.RowIndex).Value
                    ij.DescRM = Me.DgvRMRekap.Item(Me.ColRMRekap_Desc.Index, e.RowIndex).Value
                    ij.IdentifierJO = Me.DgvRMRekap.Item(Me.ColRMRekap_IdentifierJO.Index, e.RowIndex).Value
                    ij.TblItemWO = Me.TblItemWO
                    ij.PrimaryRM = Me.DgvRMRekap.Item(Me.ColRMRekap_PrimaryCodeNo.Index, e.RowIndex).Value
                    ij.TblItemRM = Me.TblItemRM
                    ij.TblItemRMAdd = Me.TblItemRMAdd
                    ij.TblItemBom = Me.TblItemBom
                    ij.RefDate = Me.DtpTglTransaksi.Value
                    ij.PrimaryMO = Me.TxtPrimary.Text

                    For Each rowDGV In Me.DgvItem.Rows
                        If rowDGV.Cells(Me.ColItem_Identifier.Index).Value = ij.IdentifierJO Then
                            ij.CodeNoJO = rowDGV.Cells(Me.ColItem_CodeNo.Index).Value
                            ij.DescJO = rowDGV.Cells(Me.ColItem_Desc.Index).Value
                            ij.JONO = rowDGV.Cells(Me.ColItem_JONumber.Index).Value

                            Exit For
                        End If
                    Next
                    If Me.TxtRemarks.Enabled = False Then
                        ij.ReadOnly_ = True
                    End If

                    ij.ShowDialog()
                    If ij.AdaPerubahan = True Then
                        Me.LoadRMRekap(Me.DgvRMRekap.Item(Me.ColRMRekap_IdentifierJO.Index, e.RowIndex).Value, Me.TblItemWO, Me.TblItemRM, Me.DgvRMRekap, Me)
                        Me.DgvRMRekap.BringToFront()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvRMRekap_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvRMRekap.Sorted
        Try
            Me.ObjCD.SetNumber(Me.DgvRMRekap)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetRequiredDate(ByVal vIdentifierJO As String) As Date
        Dim RowItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmRow

        Try
            For Each RowItemRM In Me.TblItemRM.Rows
                If RowItemRM.Identifier_JO = vIdentifierJO Then
                    Return RowItemRM.required_date
                End If
            Next
            Return Me.DtpTglTransaksi.Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return Me.DtpTglTransaksi.Value
        End Try
    End Function

    Private Function ApakahMemilikiAllocatedWIP(ByVal vIdentifierJO As String) As Boolean
        Try
            Me.TblItemWIP.DefaultView.RowFilter = "[Identifier JO] = '" + vIdentifierJO + "'"
            If Me.TblItemWIP.DefaultView.Count > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub DgvItem_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvItem.CellContentClick
        Dim bolehEditBOM As Boolean = False
        Dim IdentifierJO As String = ""

        Try
            If e.RowIndex <> -1 Then
                If Me.DgvItem.Item(ColItem_CodeNo.Index, e.RowIndex).Value <> "" Then
                    If e.ColumnIndex = Me.ColItem_DeliveryDate.Index Then
                        Dim ReadOnly_ As Boolean = False
                        Dim DI As Integer = 0
                        Dim JmlDelivery As Integer = 0
                        Dim DelDateD As Date = Today
                        If Me.TxtRemarks.Enabled = False Then
                            ReadOnly_ = True
                        End If
                        Dim ij As New FrmTrans_ManufacturingOrder_SetDelivery(Me.BackColorAll, Me, Me.DgvItem.Item(Me.ColItem_Identifier.Index, e.RowIndex).Value, _
                                Me.DgvItem.Item(Me.ColItem_PrimaryProduct.Index, e.RowIndex).Value, Me.ObjUnit.GetPrimary(Me.DgvItem.Item(Me.ColItem_Unit.Index, e.RowIndex).Value), ReadOnly_)
                        If ReadOnly_ Then
                            ij.ShowDialog()
                            Exit Sub
                        End If
                        If ij.ShowDialog = Windows.Forms.DialogResult.OK Then
                            DI = 0
                            While DI <= Me.TBLItemDelivery.Rows.Count - 1
                                Me.RowItemDelivery = Me.TBLItemDelivery.Rows(DI)
                                If Me.RowItemDelivery.identifier_jo = ij.VIdentifierJO Then
                                    Me.TBLItemDelivery.Rows.Remove(Me.RowItemDelivery)
                                Else
                                    DI += 1
                                End If
                            End While
                            ij.dgvDetail.AllowUserToAddRows = False
                            For Each rowDGV As DataGridViewRow In ij.dgvDetail.Rows
                                RowItemDelivery = Me.TBLItemDelivery.Newtbltrans_mo_itemdeliveryRow
                                With RowItemDelivery
                                    .delivery_date = CDate(rowDGV.Cells(ij.ColDeliveryDate.Index).Value)
                                    .delivery_qty = rowDGV.Cells(ij.ColQty.Index).Value
                                    .identifier_jo = ij.VIdentifierJO
                                    .primary = -1
                                    .primary_mo = Me.TxtPrimary.Text
                                    .remarks = rowDGV.Cells(ij.ColRemarks.Index).Value
                                End With
                                Me.TBLItemDelivery.Rows.Add(RowItemDelivery)
                                JmlDelivery += 1
                                DelDateD = CDate(rowDGV.Cells(ij.ColDeliveryDate.Index).Value)
                            Next
                            If JmlDelivery > 1 Then
                                Me.DgvItem.Item(Me.ColItem_DeliveryDate.Index, e.RowIndex).Value = "Multiple"
                            ElseIf JmlDelivery = 1 Then
                                Me.DgvItem.Item(Me.ColItem_DeliveryDate.Index, e.RowIndex).Value = Format(DelDateD, "dd MMM yyyy")
                            End If
                        End If
                    End If

                    If e.ColumnIndex = Me.ColItem_Print.Index Then
                        Dim i As New CustomSMESS.FrmTrans_ManufacturingOrder_PrintJO_Option
                        i.Identifier_JO = Me.DgvItem.Item(ColItem_Identifier.Index, e.RowIndex).Value
                        i.PrimaryMO = Me.TxtPrimary.Text
                        i.MONo = Me.TxtNoTransaksi.Text
                        i.UserCode = UserCode
                        i.ShowDialog()
                    ElseIf e.ColumnIndex = Me.ColItem_SetRequiredDate.Index Then
                        If Me.TxtRemarks.Enabled = True Then
                            Dim ij As New FrmTrans_ManufacturingOrder_RMReqDate
                            ij.MODate = Me.DtpTglTransaksi.Value.Date
                            ij.dtpRequired.Value = Me.GetRequiredDate(Me.DgvItem.Item(ColItem_Identifier.Index, e.RowIndex).Value)
                            If ij.ShowDialog = Windows.Forms.DialogResult.OK Then
                                Dim RowItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmRow
                                For Each RowItemRM In Me.TblItemRM.Rows
                                    RowItemRM.required_date = ij.dtpRequired.Value.Date
                                Next
                                LoadDgvItemRM(Me.DgvItem.Item(Me.ColItem_Identifier.Index, Me.DgvItem.SelectedCells.Item(0).RowIndex).Value, _
                                Me.TblItemWO, Me.TblItemRM, Me)
                            End If
                        End If
                    ElseIf e.ColumnIndex = Me.ColItem_BoM_Tools.Index Then
                        If Me.TxtRemarks.Enabled = True Then
                            If Me.Baru = False Then
                                If Me.ObjMO.IsScheduled(Me.TxtPrimary.Text, Me.DgvItem.Item(Me.ColItem_Identifier.Index, e.RowIndex).Value) = True Then
                                    MsgBox("This JO is scheduled, you can't edit this data.", MsgBoxStyle.Information, "Information")
                                    Exit Sub
                                End If
                            End If

                            IdentifierJO = Me.DgvItem.Item(Me.ColItem_Identifier.Index, e.RowIndex).Value
                            If Me.ApakahMemilikiAllocatedWIP(IdentifierJO) = True Then
                                MsgBox("This JO is reference by wip item, you can't edit bill of material data.", MsgBoxStyle.Information, "Information")
                                Exit Sub
                            End If
                        End If
                        If Me.btnOK.Enabled = False Then Exit Sub

                        Dim frmBoMTools As New Inventory_Module.frmMasterProduct_BOM1
                        frmBoMTools.UserCode = Me.UserCode
                        frmBoMTools.BackColorAll = Me.BackColorAll
                        frmBoMTools.vPrimaryProduct = Me.DgvItem.Item(ColItem_PrimaryProduct.Index, e.RowIndex).Value
                        frmBoMTools.vPrimaryProductJO = Me.DgvItem.Item(ColItem_PrimaryProduct.Index, e.RowIndex).Value
                        frmBoMTools.vQtyTarget = CDec(Me.DgvItem.Item(Me.ColItem_QtyTarget.Index, e.RowIndex).Value)
                        frmBoMTools.vTblItemRM = Me.TblItemRM
                        frmBoMTools.TblItemBOM_ = Me.TblItemBom
                        frmBoMTools.vTblItemWO = Me.TblItemWO
                        frmBoMTools.PrimaryMO = Me.TxtPrimary.Text
                        frmBoMTools.vTblProsesOrder = Me.TblProsesOrder
                        frmBoMTools.IdentifierJO = Me.DgvItem.Item(ColItem_Identifier.Index, e.RowIndex).Value
                        frmBoMTools.vProductCode = Me.DgvItem.Item(ColItem_CodeNo.Index, e.RowIndex).Value.ToString & " - " & Me.DgvItem.Item(ColItem_Desc.Index, e.RowIndex).Value
                        frmBoMTools.vDefPrec = Me.DgvItem.Item(Me.ColItem_Decimal.Index, e.RowIndex).Value
                        frmBoMTools.ShowDialog()
                        Me.TblItemBom = frmBoMTools.TblItemBOM_
                        Me.DgvItem.Item(Me.ColItem_QtyProduce.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(ObjMO.GetQtyProduce_InTbl(Me.DgvItem.Item(Me.ColItem_PrimaryProduct.Index, e.RowIndex).Value, _
                        Me.DgvItem.Item(Me.ColItem_QtyTarget.Index, e.RowIndex).Value, Me.DgvItem.Item(Me.ColItem_Identifier.Index, e.RowIndex).Value, Me.TblItemBom), _
                        Me.DgvItem.Item(Me.ColItem_Decimal.Index, e.RowIndex).Value, Core_Module.ClsCekData.EnumTipeRounding.Up)
                        If Me.DgvItem.SelectedCells.Count > 0 Then
                            If String.IsNullOrEmpty(Me.DgvItem.Item(Me.ColItem_PrimaryProduct.Index, Me.DgvItem.SelectedCells.Item(0).RowIndex).Value) = False Then
                                If TabControl1.SelectedTab.Name = Me.TabPage_WO.Name Then
                                    LoadDgvItemWO(Me.DgvItem.Item(Me.ColItem_Identifier.Index, Me.DgvItem.SelectedCells.Item(0).RowIndex).Value, _
                                    Me.TblItemWO, Me)
                                ElseIf TabControl1.SelectedTab.Name = Me.TabPage_RM.Name Then
                                    LoadDgvItemRM(Me.DgvItem.Item(Me.ColItem_Identifier.Index, Me.DgvItem.SelectedCells.Item(0).RowIndex).Value, _
                                    Me.TblItemWO, Me.TblItemRM, Me)
                                ElseIf TabControl1.SelectedTab.Name = Me.TabPage_WIP.Name Then
                                    Me.LoadItemWIP()
                                End If
                                Me.ClbProcess.DataSource = Nothing
                                ObjMO.IsiListBoxProses(Me.ClbProcess, Me.DgvItem.Item(Me.ColItem_Identifier.Index, Me.DgvItem.SelectedCells.Item(0).RowIndex).Value, _
                                Me.TblProsesOrder, Me.TblProses_Filter)
                            End If
                        End If

                    ElseIf e.ColumnIndex = Me.ColItem_EditBOM.Index Then
                        If Me.TxtRemarks.Enabled = True Then
                            If Me.Baru = False Then
                                If Me.ObjMO.IsScheduled(Me.TxtPrimary.Text, Me.DgvItem.Item(Me.ColItem_Identifier.Index, e.RowIndex).Value) = True Then
                                    MsgBox("This JO is scheduled, you can't edit this data.", MsgBoxStyle.Information, "Information")
                                    Exit Sub
                                End If
                            End If

                            IdentifierJO = Me.DgvItem.Item(Me.ColItem_Identifier.Index, e.RowIndex).Value
                            If Me.ApakahMemilikiAllocatedWIP(IdentifierJO) = True Then
                                MsgBox("This JO is reference by wip item, you can't edit bill of material data.", MsgBoxStyle.Information, "Information")
                                Exit Sub
                            End If
                        End If

                        Dim i As New FrmTrans_ManufacturingOrder_BOM
                        i.UserCode = Me.UserCode
                        i.vPrimaryMO = CDbl(Me.TxtPrimary.Text)
                        i.vPrimaryProduct = Me.DgvItem.Item(ColItem_PrimaryProduct.Index, e.RowIndex).Value
                        i.vPrimaryProductJO = Me.DgvItem.Item(ColItem_PrimaryProduct.Index, e.RowIndex).Value
                        i.vIdentifierJO = Me.DgvItem.Item(ColItem_Identifier.Index, e.RowIndex).Value
                        i.vProductCode = Me.DgvItem.Item(ColItem_CodeNo.Index, e.RowIndex).Value.ToString & " - " & Me.DgvItem.Item(ColItem_Desc.Index, e.RowIndex).Value
                        i.vDefPrec = Me.DgvItem.Item(Me.ColItem_Decimal.Index, e.RowIndex).Value
                        i.vQtyTarget = CDec(Me.DgvItem.Item(Me.ColItem_QtyTarget.Index, e.RowIndex).Value)
                        i.vTblItemBOM = Me.TblItemBom
                        i.vTblProsesOrder = Me.TblProsesOrder
                        i.vTblItemWO = Me.TblItemWO
                        i.vTblItemRM = Me.TblItemRM
                        i.vTblItemWIP = Me.TblItemWIP

                        If Me.btnOK.Enabled = True Then
                            If ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_ManufacturingOrder_AddItem") = False Then
                                i.vBolehEdit = False
                            Else
                                i.vBolehEdit = True
                            End If
                        End If

                        'jika simpan perubahan di form bom maka update qty produce di item JO
                        If i.ShowDialog = DialogResult.OK Then
                            Me.TblItemBom = i.vTblItemBOM
                            Me.DgvItem.Item(Me.ColItem_QtyProduce.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(ObjMO.GetQtyProduce_InTbl(Me.DgvItem.Item(Me.ColItem_PrimaryProduct.Index, e.RowIndex).Value, _
                            Me.DgvItem.Item(Me.ColItem_QtyTarget.Index, e.RowIndex).Value, Me.DgvItem.Item(Me.ColItem_Identifier.Index, e.RowIndex).Value, Me.TblItemBom), _
                            Me.DgvItem.Item(Me.ColItem_Decimal.Index, e.RowIndex).Value, Core_Module.ClsCekData.EnumTipeRounding.Up)
                        End If

                        i.Dispose()

                        'Load ulang Tgv WO / Dgv RM
                        If Me.DgvItem.SelectedCells.Count > 0 Then
                            If String.IsNullOrEmpty(Me.DgvItem.Item(Me.ColItem_PrimaryProduct.Index, Me.DgvItem.SelectedCells.Item(0).RowIndex).Value) = False Then
                                If TabControl1.SelectedTab.Name = Me.TabPage_WO.Name Then
                                    LoadDgvItemWO(Me.DgvItem.Item(Me.ColItem_Identifier.Index, Me.DgvItem.SelectedCells.Item(0).RowIndex).Value, _
                                    Me.TblItemWO, Me)
                                ElseIf TabControl1.SelectedTab.Name = Me.TabPage_RM.Name Then
                                    LoadDgvItemRM(Me.DgvItem.Item(Me.ColItem_Identifier.Index, Me.DgvItem.SelectedCells.Item(0).RowIndex).Value, _
                                    Me.TblItemWO, Me.TblItemRM, Me)
                                ElseIf TabControl1.SelectedTab.Name = Me.TabPage_WIP.Name Then
                                    Me.LoadItemWIP()
                                End If

                                'isi listbox proses untuk 
                                Me.ClbProcess.DataSource = Nothing
                                ObjMO.IsiListBoxProses(Me.ClbProcess, Me.DgvItem.Item(Me.ColItem_Identifier.Index, Me.DgvItem.SelectedCells.Item(0).RowIndex).Value, _
                                Me.TblProsesOrder, Me.TblProses_Filter)
                            End If
                        End If
                    End If
                End If
                If e.ColumnIndex = Me.ColItem_Allocation.Index Then
                    If Me.TxtRemarks.Enabled = True Then
                        If Me.Baru = False Then
                            If Me.ObjMO.IsScheduled(Me.TxtPrimary.Text, Me.DgvItem.Item(Me.ColItem_Identifier.Index, e.RowIndex).Value) = True Then
                                MsgBox("This JO is scheduled, you can't edit this data.", MsgBoxStyle.Information, "Information")
                                Exit Sub
                            End If
                        End If
                        IdentifierJO = Me.DgvItem.Item(Me.ColItem_Identifier.Index, e.RowIndex).Value
                        If Me.ApakahMemilikiAllocatedWIP(IdentifierJO) = True Then
                            MsgBox("This JO is reference by wip item, you can't edit process order.", MsgBoxStyle.Information, "Information")
                            Exit Sub
                        End If
                    End If

                    Dim j As New FrmTrans_ManufacturingOrder_ProcessOrder
                    j.vPrimaryMO = CDbl(Me.TxtPrimary.Text)
                    j.vTblProsesOrder = Me.TblProsesOrder
                    j.vTblItemWO = Me.TblItemWO
                    j.vIdentifierJO = Me.DgvItem.Item(ColItem_Identifier.Index, e.RowIndex).Value
                    j.vDateTransaksi = Me.DtpTglTransaksi.Value
                    j.vFrmMO = Me

                    If Me.btnOK.Enabled = True Then
                        If ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_ManufacturingOrder_AddItem") = False Then
                            j.vBolehEdit = False
                        Else
                            j.vBolehEdit = True
                        End If
                    End If

                    j.ShowDialog()
                    j.Dispose()

                    'Load ulang Tgv WO / Dgv RM
                    If Me.DgvItem.SelectedCells.Count > 0 Then
                        If String.IsNullOrEmpty(Me.DgvItem.Item(Me.ColItem_PrimaryProduct.Index, Me.DgvItem.SelectedCells.Item(0).RowIndex).Value) = False Then
                            If TabControl1.SelectedTab.Name = Me.TabPage_WO.Name Then
                                LoadDgvItemWO(Me.DgvItem.Item(Me.ColItem_Identifier.Index, Me.DgvItem.SelectedCells.Item(0).RowIndex).Value, _
                                Me.TblItemWO, Me)
                            ElseIf TabControl1.SelectedTab.Name = Me.TabPage_RM.Name Then
                                LoadDgvItemRM(Me.DgvItem.Item(Me.ColItem_Identifier.Index, Me.DgvItem.SelectedCells.Item(0).RowIndex).Value, _
                                 Me.TblItemWO, Me.TblItemRM, Me)
                            ElseIf Me.TabControl1.SelectedTab.Name = Me.TabPage_WIP.Name Then
                                Me.LoadItemWIP()
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function ApakahDataBenar() As Boolean
        Dim RowItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmRow
        Try
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Me.DtpTglTransaksi.Value Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            If Me.Baru = True Then
                If Me.ObjMO.Apakahsudahada(Me.TxtNoTransaksi.Text) = True Then
                    Me.TxtNoTransaksi.Text = Me.ObjMO.GetNoBaru1(Me.DtpTglTransaksi.Value, ClsTrans_ManufacturingOrder.PilihFormatNumber.MO, -1, Me)
                    ' MsgBox("Duplication MO Number, please change your MO Number", _
                    'MsgBoxStyle.Information, "Information")
                    ' Return False
                End If
            Else
                Dim productGroupLama As String = Me.ObjMO.GetNoTransaksi(Me.TxtPrimary.Text)
                If productGroupLama <> Me.TxtNoTransaksi.Text Then
                    If Me.ObjMO.Apakahsudahada(Me.TxtNoTransaksi.Text) = True Then
                        Me.TxtNoTransaksi.Text = Me.ObjMO.GetNoBaru1(Me.DtpTglTransaksi.Value, ClsTrans_ManufacturingOrder.PilihFormatNumber.MO, -1, Me)
                        'MsgBox("Duplication MO Number, please change your MO Number", _
                        ' MsgBoxStyle.Information, "Information")
                        'Return False
                    End If
                End If

                'data salah jika nilai timestamp yang ada dibasis data sudah tidak valid lagi, khusus untuk 
                'proses edit
                Dim oldTimeStamp As Date = Me.ObjMO.GetTimestamp(Me.TxtPrimary.Text)
                If Me.ObjMO.TimeStamp <> oldTimeStamp Then
                    If MsgBox("Other user have changed your current data, you cannot save this data. Reload data from database?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.ObjMO.CariDetail(Me.ObjMO.Primary, Me.DgvItem, Me.TblItemBom, Me.TblProsesOrder, Me.TblItemWO, Me.TblItemRM, Me, _
                        Me.ObjCD, Me.TblItemRMAdd, Me.TblItemWIP, Me.TBLItemDelivery, Me.UserCode)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            If Me.TxtRemarks.Text.Contains("#") = True Or Me.TxtRemarks.Text.Contains("*") = True _
            Or Me.TxtRemarks.Text.Contains("%") Or Me.TxtRemarks.Text.Contains("#") = True _
            Or Me.TxtRemarks.Text.Contains("`") = True Then
                Me.TxtRemarks.Focus()
                MsgBox("Remarks data can not contain the following characters: ` , #, % , or * .", MsgBoxStyle.Information, "Error")
                Return False
            End If

            If Me.ApakahDataItemBenar = False Then
                Return False
            End If

            'pastikan qty rm > 0
            For Each RowItemRM In Me.TblItemRM.Rows
                If RowItemRM.Required_Qty < 0 Then
                    MsgBox("Please review raw material: " + Me.ObjProduct.GetProductCode(RowItemRM.Primary_Product_Code_RM) + _
                    " in JO: " + Me.GetJONumber(RowItemRM.Identifier_JO) + ". The required qty is negative.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Function GetJONumber(ByVal vIdentifierJO As String) As String
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.DgvItem.Rows
                If rowDGV.Cells(Me.ColItem_Identifier.Index).Value = vIdentifierJO Then
                    Return rowDGV.Cells(Me.ColItem_JONumber.Index).Value
                End If
            Next

            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function ApakahDataItemBenar() As Boolean

        Try
            'minimal 1 item
            If Me.DgvItem.RowCount = 1 Then
                For Each rowD As DataGridViewRow In Me.DgvItem.Rows
                    If String.IsNullOrEmpty(rowD.Cells(Me.ColItem_CodeNo.Index).Value) = True Then
                        MsgBox("Please enter at least 1 item.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                Next
            ElseIf Me.DgvItem.RowCount > 1 Then
                'For Each rowD As DataGridViewRow In Me.DgvItem.Rows
                '    If rowD.Cells(Me.ColItem_Remarks.Index).Value.Contains("#") = True Or rowD.Cells(Me.ColItem_Remarks.Index).Value.Contains("*") = True _
                '        Or rowD.Cells(Me.ColItem_Remarks.Index).Value.Contains("%") Or rowD.Cells(Me.ColItem_Remarks.Index).Value.Contains("#") = True _
                '        Or rowD.Cells(Me.ColItem_Remarks.Index).Value.Contains("`") = True Then
                '        rowD.Cells(Me.ColItem_Remarks.Index).Selected = True
                '        MsgBox("Remarks data can not contain the following characters: ` , #, % , or * .", MsgBoxStyle.Information, "Error")
                '        Return False
                '    End If
                'Next
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Sub AktifkanDetail()
        Dim OtherUser As String = ""

        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimary.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.TxtPrimary.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.TxtPrimary.BackColor, False)
            Me.DtpTglTransaksi.Value = Now
            Me.DgvItem.AllowUserToAddRows = True
            Me.DgvItem.AllowUserToDeleteRows = True
            Me.DgvItem.TabStop = True
            Me.ChkScheduling.Enabled = True
            Me.DgvItem.ReadOnly = False
            Me.dgvItemWO.ReadOnly = False
            Me.DGVWIP.ReadOnly = True
            Me.DGVWIP.AllowUserToDeleteRows = False

            Me.SplitContainer1.TabStop = False
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)
            Me.ObjCD.SetNumber(Me.DgvItem, False)

            'If Baru = False Then
            '    For ij As Integer = 0 To DgvItem.Rows.Count - 1
            '        If String.IsNullOrEmpty(Me.DgvItem.Item(Me.ColItem_CodeNo.Index, ij).Value) = False Then
            '            Me.DtpTglTransaksi.Enabled = False
            '            Exit For
            '        Else : Me.ChkScheduling.Checked = False
            '            Me.DtpTglTransaksi.Enabled = True
            '            Exit For
            '        End If
            '    Next
            'End If

            Me.btnUP.Enabled = False
            Me.btnDown.Enabled = False
            Me.ChkScheduling.Enabled = False
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            Me.TGV_WO.ReadOnly = False
            Me.DGVRM.ReadOnly = False

            If Me.TxtRemarks.Enabled = False Then
                Me.llbAddWIP.Visible = False
            Else
                If Me.TabControl1.SelectedTab.Name = Me.TabPage_WIP.Name Then
                    Me.llbAddWIP.Visible = True
                Else
                    Me.llbAddWIP.Visible = False
                End If
            End If

            Me.ColItem_Print.Visible = False
            Me.ColItem_SetRequiredDate.Visible = True
            If Me.BoM_ToolsEnabled = True Then
                Me.ColItem_BoM_Tools.Visible = True
            End If

            If Me.Baru = False Then
                Me.Identifier = Me.ObjRandom.Next(0, 10 ^ 7)
                Me.ObjMonitor.StartMonitor(Me.UserCode, My.Computer.Name, Core_Module.ClsTransaksiMonitor.EnumTipeTransaksi.MO_Edit, Me.Identifier, Me.TxtPrimary.Text)
                If Me.ObjMonitor.IsTheFirstOne(Core_Module.ClsTransaksiMonitor.EnumTipeTransaksi.MO_Edit, Me.Identifier, OtherUser, Me.TxtPrimary.Text) = False Then
                    Me.NonAktifkanDetail()
                    MsgBox("User: " + OtherUser + " currently editing this transaction. You can't edit this data.", MsgBoxStyle.Information, "Information")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.Controls, True, False, True, True, True, False, True, False)
            Me.DgvItem.AllowUserToAddRows = False
            Me.DgvItem.AllowUserToDeleteRows = False
            Me.DgvItem.TabStop = False
            Me.DgvItem.ReadOnly = True
            Me.dgvItemWO.ReadOnly = True
            Me.DGVWIP.ReadOnly = True
            Me.DGVWIP.AllowUserToDeleteRows = False
            Me.ChkScheduling.Enabled = False
            Me.SplitContainer1.TabStop = True
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.BackColor, False)
            Me.ChkScheduling.Enabled = True
            Me.CancelButton = Me.btnExit
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            Me.TGV_WO.ReadOnly = True
            Me.DGVRM.ReadOnly = True
            Me.llbAddWIP.Visible = False
            Me.ColItem_Print.Visible = True
            Me.ColItem_SetRequiredDate.Visible = False
            Me.ColItem_BoM_Tools.Visible = False

            If Me.Baru = False Then
                Me.ObjMonitor.EndMonitor(Core_Module.ClsTransaksiMonitor.EnumTipeTransaksi.MO_Edit, Me.Identifier)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_ManufacturingOrder_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.chkRM_IncludeWIP.Checked = True Then
                Me.objSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ManufacturingOrder, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ManufacturingOrder_RMIncludeWIP, "1", Me.UserCode)
            Else
                Me.objSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ManufacturingOrder, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ManufacturingOrder_RMIncludeWIP, "0", Me.UserCode)
            End If
            If Me.chkRM_ShowDetail.Checked = True Then
                Me.objSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ManufacturingOrder, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ManufacturingOrder_RMShowDetail, "1", Me.UserCode)
            Else
                Me.objSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ManufacturingOrder, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ManufacturingOrder_RMShowDetail, "0", Me.UserCode)
            End If
            If Me.rboWO_UseTreeView.Checked = True Then
                Me.objSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ManufacturingOrder, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ManufacturingOrder_WO_UseTreeView, "1", Me.UserCode)
            Else
                Me.objSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ManufacturingOrder, _
                Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ManufacturingOrder_WO_UseTreeView, "0", Me.UserCode)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try 
    End Sub

    Private Sub FrmManufacturing_Order_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim Dummy As String = ""

        Try
            Me.ColItem_SetRequiredDate.HeaderText = "Edit RM Required Date"
            Me.DGVRM.AllowUserToDeleteRows = False
            Me.ColRM_GPActual.Visible = False
            Me.ColRM_GPTarget.Visible = False
            Me.ColRM_PickUpLocation.Visible = False
            Me.ColRM_StandByLocation.Visible = False
            Me.DgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.TGV_WO.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DGVRM.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvRMRekap.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvItemWO.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DGVWIP.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.btnFind.Select()
            Me.Baru = True
            Me.NonAktifkanDetail()
            Me.KosongkanField()
            'nanti hilangin yang bawah vvv
            'Me.ColItem_Urutan.Visible = True
            'Me.ColItem_Identifier.Visible = True
            'Me.ColWO2_IdentifierJO.Visible = True
            'Me.ColWO2_IdentifierItem.Visible = True
            'Me.ColRM_IdentifierItem.Visible = True
            'Me.ColRM_IdentifierWO.Visible = True
            Me.ColItem_Allocation.HeaderText = "Process Order"
            Me.DtpTargetFinishDate.Visible = False
            Me.DtpTargetFinishDate.TabStop = False
            Me.Label4.Visible = False
            ' Me.ColItem_Urutan.Visible = True
            Me.setColumnVisibility()
            Me.SetColumnVisibilityWIP()

            Me.ExecuteChkRM = False
            Me.ExecuteWO = False
            Dummy = Me.objSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ManufacturingOrder, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ManufacturingOrder_RMIncludeWIP, Me.UserCode)
            If Dummy = "1" Then
                Me.chkRM_IncludeWIP.Checked = True
            Else
                Me.chkRM_IncludeWIP.Checked = False
            End If
            Dummy = Me.objSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ManufacturingOrder, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ManufacturingOrder_RMShowDetail, Me.UserCode)
            If Dummy = "1" Then
                Me.chkRM_ShowDetail.Checked = True
            Else
                Me.chkRM_ShowDetail.Checked = False
            End If
            Dummy = Me.objSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ManufacturingOrder, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ManufacturingOrder_WO_UseTreeView, Me.UserCode)
            If Dummy = "1" Then
                Me.rboWO_UseTreeView.Checked = True
            Else
                Me.rboWO_UseTreeView.Checked = False
            End If
            Me.ExecuteChkRM = True
            Me.ExecuteWO = True

            'Me.ObjSettingGlobal.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ManufacturingOrder, Core_Module.ClsTransaksiSetting.EnumNamaParameter.MO_BoM_ToolsEnabled, "1")
            Dim BoM_ToolsEnabledS As String = Me.ObjSettingGlobal.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ManufacturingOrder, Core_Module.ClsTransaksiSetting.EnumNamaParameter.MO_BoM_ToolsEnabled)
            If BoM_ToolsEnabledS = "" Then
                Me.ObjSettingGlobal.SetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ManufacturingOrder, Core_Module.ClsTransaksiSetting.EnumNamaParameter.MO_BoM_ToolsEnabled, "0")
            ElseIf BoM_ToolsEnabledS = "1" Then
                Me.BoM_ToolsEnabled = True
            End If

            Me.chkRM_IncludeWIP.Left = Me.chkRM_ShowDetail.Left - 5 - Me.chkRM_IncludeWIP.Width
            Me.ColItem_Allocation.Visible = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNoTransaksi.GotFocus, TxtRemarks.GotFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnOK.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.TxtReadOnly.BackColor
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

    Private Sub txtcode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNoTransaksi.LostFocus, TxtRemarks.LostFocus
        Dim Txt As TextBox
        Try
            Txt = TryCast(sender, TextBox)
            If IsNothing(Txt) = False Then
                If Me.btnOK.Enabled = True Then
                    If Txt.Enabled = True Then
                        If Txt.ReadOnly = True Then
                            Txt.BackColor = Me.TxtReadOnly.BackColor
                        Else
                            Txt.BackColor = Me.TxtPrimary.BackColor
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarks_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtRemarks.KeyPress
        Try
            If Not (Asc(e.KeyChar) = 8) Then
                Dim allowedChars As String = "'%*"
                If allowedChars.Contains(e.KeyChar.ToString.ToLower) Then
                    e.KeyChar = ChrW(0)
                    e.Handled = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemarks_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarks.Leave
        Try
            Me.TxtRemarks.Text = Me.TxtRemarks.Text.Trim.Replace("'", "").Replace("%", "").Replace("*", "")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanTabelDetail(ByVal vIdentifierJO As String)
        Dim rowItemBom As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomRow
        Dim rowProsesOrder As DseTrans_ManufacturingOrder.tbltrans_mo_processorderRow
        Dim rowItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woRow
        Dim rowItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmRow
        Dim RowItemWIP As DseTrans_ManufacturingOrder.tbltrans_mo_item_wipRow
        Dim RowItemRMAdd As DseTrans_ManufacturingOrder.tbltrans_mo_item_rm_addRow

        Try
            'hapus datatable item bom berdasarkan identifier jo
            Me.TblItemBom.DefaultView.RowFilter = "[Identifier JO] = '" & vIdentifierJO & "'"
            While Me.TblItemBom.DefaultView.Count > 0
                rowItemBom = Me.TblItemBom.DefaultView(0).Row
                Me.TblItemBom.Rows.Remove(rowItemBom)
            End While

            'hapus datatable item proses order berdasarkan identifier jo
            Me.TblProsesOrder.DefaultView.RowFilter = "[identifier jo] = '" & vIdentifierJO & "'"
            While Me.TblProsesOrder.DefaultView.Count > 0
                rowProsesOrder = Me.TblProsesOrder.DefaultView(0).Row
                Me.TblProsesOrder.Rows.Remove(rowProsesOrder)
            End While

            'hapus datatable item Work order berdasarkan identifier jo
            Me.TblItemWO.DefaultView.RowFilter = "[identifier jo] = '" & vIdentifierJO & "'"
            While TblItemWO.DefaultView.Count > 0
                rowItemWO = Me.TblItemWO.DefaultView(0).Row
                Me.TblItemWO.Rows.Remove(rowItemWO)
            End While

            'hapus datatable item Raw Material berdasarkan identifier jo
            Me.TblItemRM.DefaultView.RowFilter = "[identifier jo] = '" & vIdentifierJO & "'"
            While TblItemRM.DefaultView.Count > 0
                rowItemRM = Me.TblItemRM.DefaultView(0).Row
                Me.TblItemRM.Rows.Remove(rowItemRM)
            End While

            'hapus datatable JO item temporary
            Me.TblJO_Temp.DefaultView.RowFilter = "[identifier jo] = '" & vIdentifierJO.ToString & "'"
            While Me.TblJO_Temp.DefaultView.Count > 0
                RowJO_Temp = Me.TblJO_Temp.DefaultView(0).Row
                Me.TblJO_Temp.Rows.Remove(RowJO_Temp)
            End While

            Me.TBLItemDelivery.DefaultView.RowFilter = "[identifier jo] = '" + vIdentifierJO.ToString + "'"
            While Me.TBLItemDelivery.DefaultView.Count > 0
                Me.RowItemDelivery = Me.TBLItemDelivery.DefaultView(0).Row
                Me.TBLItemDelivery.Rows.Remove(Me.RowItemDelivery)
            End While
            Me.TBLItemDelivery.DefaultView.RowFilter = ""

            'item wip
            Me.TblItemWIP.DefaultView.RowFilter = "[Identifier JO] = '" & vIdentifierJO.ToString & "'"
            While Me.TblItemWIP.DefaultView.Count > 0
                RowItemWIP = Me.TblItemWIP.DefaultView(0).Row
                Me.TblItemWIP.Rows.Remove(RowItemWIP)
            End While
            Me.TblItemWIP.DefaultView.RowFilter = ""

            'item rm add
            Me.TblItemRMAdd.DefaultView.RowFilter = "[Identifier JO] = '" & vIdentifierJO.ToString & "'"
            While Me.TblItemRMAdd.DefaultView.Count > 0
                RowItemRMAdd = Me.TblItemRMAdd.DefaultView(0).Row
                Me.TblItemRMAdd.Rows.Remove(RowItemRMAdd)
            End While
            Me.TblItemRMAdd.DefaultView.RowFilter = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            Me.ObjCD.KosongkanControl(Me.gboDetail.Controls, True, False, True)
            Me.ObjCD.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, False, True)
            Me.ObjCD.KosongkanControl(Me.SplitContainer1.Panel2.Controls, True, False, True)
            Me.TxtNoTransaksi.Clear()
            Me.TxtRemarks.Clear()
            Me.ChkScheduling.Checked = False
            Me.TxtPrimary.Text = "-1"
            Me.TxtPrimaryUserAuth.Text = "-1"
            Me.DgvItem.Rows.Clear()
            Me.TGV_WO.Nodes.Clear()
            Me.TGV_WO.Rows.Clear()
            Me.DGVRM.Rows.Clear()
            Me.DgvRMRekap.Rows.Clear()
            Me.dgvItemWO.Rows.Clear()
            Me.DGVWIP.Rows.Clear()
            Me.TblItemBom.Rows.Clear()
            Me.TblProsesOrder.Rows.Clear()
            Me.TblItemWO.Rows.Clear()
            Me.TblItemRM.Rows.Clear()
            Me.TblItemRMAdd.Rows.Clear()
            Me.TblJO_Temp.Rows.Clear()
            Me.TblItemWIP.Rows.Clear()
            Me.TBLItemDelivery.Rows.Clear()

            Me.ClbProcess.DataSource = Nothing
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
            Me.llbAddWIP.Text = "Add WIP"
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

            If IsValidAutoNumber() = False Then
                Exit Sub
            End If

            Me.Baru = True
            Me.KosongkanField()
            Me.DtpTglTransaksi.Value = Now.Date
            Me.TxtNoTransaksi.Text = Me.ObjMO.GetNoBaru1(Me.DtpTglTransaksi.Value, ClsTrans_ManufacturingOrder.PilihFormatNumber.MO, -1, Me)
            Me.TxtNoTransaksi.Focus()
            Me.AktifkanDetail()

            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Function IsValidAutoNumber() As Boolean
        Dim ObjAutoNumber As New Core_Module.ClsMasterTransactionNumbering
        Try
            If ObjAutoNumber.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.JO) = False Then
                MsgBox("Please set auto number for job order transaction, contact your software vendor for more information.", MsgBoxStyle.Information, "Information")
                Return False
            End If
            If ObjAutoNumber.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.WO) = False Then
                MsgBox("Please set auto number for work order transaction, contact your software vendor for more information.", MsgBoxStyle.Information, "Information")
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub SalinKeObjek()
        Try
            With Me.ObjMO
                .Primary = Me.TxtPrimary.Text
                .TanggalTransaksi = Me.DtpTglTransaksi.Value
                .NoTransaksi = Me.TxtNoTransaksi.Text
                .PrimaryUser_Auth = Me.TxtPrimaryUserAuth.Text
                If Me.ChkScheduling.Checked = True Then
                    .IsReady_ForScheduling = ClsTrans_ManufacturingOrder.EnumPilihSchedule.Schedule_Aktif
                Else
                    .IsReady_ForScheduling = ClsTrans_ManufacturingOrder.EnumPilihSchedule.Schedule_tdk_Aktif
                End If
                .Remarks = Me.TxtRemarks.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SalinKeField()
        Try
            With Me.ObjMO
                Me.TxtPrimary.Text = .Primary
                Me.DtpTglTransaksi.Value = .TanggalTransaksi
                Me.TxtPrimaryUserAuth.Text = .PrimaryUser_Auth
                If .IsReady_ForScheduling = ClsTrans_ManufacturingOrder.EnumPilihSchedule.Schedule_Aktif Then
                    Me.ChkScheduling.Checked = True
                ElseIf .IsReady_ForScheduling = ClsTrans_ManufacturingOrder.EnumPilihSchedule.Schedule_tdk_Aktif Then
                    Me.ChkScheduling.Checked = False
                End If
                Me.TxtNoTransaksi.Text = .NoTransaksi
                Me.TxtRemarks.Text = .Remarks

                If Me.DgvItem.Rows.Count > 0 Then
                    Me.DgvItem.ClearSelection()
                    Dim r As DataGridViewCell = Me.DgvItem.Item(0, 0)
                    Me.ForceReload = True
                    r.Selected = True
                End If

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.DgvItem.RowCount > 0 Then
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Baru = True Then
                        Me.KosongkanField()
                    Else
                        Me.ObjMO.CariDetail(Me.TxtPrimary.Text, Me.DgvItem, Me.TblItemBom, Me.TblProsesOrder, _
                        Me.TblItemWO, Me.TblItemRM, Me, Me.ObjCD, Me.TblItemRMAdd, Me.TblItemWIP, Me.TBLItemDelivery, Me.UserCode)
                        Me.SalinKeField()
                        Me.DgvItem.AllowUserToAddRows = True
                    End If
                Else
                    Exit Sub
                End If
            Else
                Me.KosongkanField()
            End If
            Me.NonAktifkanDetail()
            Me.btnNew.Focus()

            If Me.FromMP = True Then
                Me.DialogResult = Windows.Forms.DialogResult.OK
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim hj As New FrmTrans_ManufacturingOrder_Cari
        Try
            hj.BackColorAll = Me.BackColorAll
            hj.ShowDialog()
            If Not hj.Primary = -1 Then
                If Me.ObjMO.CariDetail(hj.Primary, Me.DgvItem, Me.TblItemBom, Me.TblProsesOrder, Me.TblItemWO, _
                Me.TblItemRM, Me, Me.ObjCD, Me.TblItemRMAdd, Me.TblItemWIP, Me.TBLItemDelivery, Me.UserCode, Me.TblJO_Temp) = True Then
                    Me.SalinKeField()
                    Me.DgvItem.ClearSelection()
                    Me.DGVRM.Rows.Clear()
                    Me.TGV_WO.Nodes.Clear()

                    Me.DgvItem.Item(0, 0).Selected = True
                    Me.btnEdit.Focus()
                Else
                    Me.btnFind.Focus()
                End If
                Me.btnFind.Focus()
                Me.NonAktifkanDetail()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Try
            If Me.TxtPrimary.Text = "-1" Then
                Me.btnFind.PerformClick()
            End If
            If Me.TxtPrimary.Text <> "-1" Then
                If Me.ObjMO.ApakahBolehEdit(Me.TxtPrimary.Text, UserCode, Me.TxtNoTransaksi.Text, True) = True Then
                    Me.vEdit = True
                    Me.Baru = False
                    Me.ObjMO.CariDetail(Me.TxtPrimary.Text, Me.DgvItem, Me.TblItemBom, Me.TblProsesOrder, _
                    Me.TblItemWO, Me.TblItemRM, Me, Me.ObjCD, Me.TblItemRMAdd, Me.TblItemWIP, Me.TBLItemDelivery, Me.UserCode, Me.TblJO_Temp)
                    Me.AktifkanDetail()
                    Me.SalinKeField()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObjek()
                If Baru = True Then
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjMO.SimpanBaru(Me.DgvItem, Me.TblItemBom, Me.TblProsesOrder, Me.TblItemWO, Me.TblItemRM, Me.UserCode, Me, Me.TblItemRMAdd, Me.TblItemWIP, Me.TBLItemDelivery) Then
                            Me.DialogResult = Windows.Forms.DialogResult.OK
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnFind.Focus()
                        End If
                    End If
                Else
                    If Me.ObjMO.ApakahBolehEdit(Me.TxtPrimary.Text, UserCode, Me.TxtNoTransaksi.Text, True) = True Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                        If hasil = MsgBoxResult.Yes Then
                            If Me.ObjMO.SimpanPerubahan(Me.DgvItem, Me.TblItemBom, Me.TblProsesOrder, Me.TblItemWO, Me.TblItemRM, Me.UserCode, Me, Me.TblItemRMAdd, Me.TblItemWIP, Me.TBLItemDelivery) Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.DialogResult = Windows.Forms.DialogResult.OK
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnEdit.Focus()
                            End If
                        End If
                    End If
                End If

                If Me.FromMP = True Then
                    Me.Close()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtNoTransaksi.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtNoTransaksi.Text <> "" Then
                If Me.ObjMO.ApakahBolehHapus(Me.TxtPrimary.Text, UserCode, Me.TxtNoTransaksi.Text, Me.TblItemWO, True) = False Then
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        If Me.ObjMO.Hapus(Me.TxtPrimary.Text, Me.UserCode, Me, Me.TxtNoTransaksi.Text, Me.TblItemWO) = True Then
                            Me.KosongkanField()
                            Me.NonAktifkanDetail()
                            Me.btnNew.Focus()
                            MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        Try
            Me.llbColumnSetting.Visible = True
            If TabControl1.SelectedTab.Name = Me.TabPage_WO.Name Then
                Me.chkRM_IncludeWIP.Visible = False
                Me.chkRM_ShowDetail.Visible = False
                Me.chkRM_IncludeWIP.TabStop = False
                Me.chkRM_ShowDetail.TabStop = False

                Me.llbAddWIP.Visible = False

                Me.rboWO_UseTableView.Visible = True
                Me.rboWO_UseTableView.TabStop = True
                Me.rboWO_UseTreeView.Visible = True
                Me.rboWO_UseTreeView.TabStop = True
            ElseIf TabControl1.SelectedTab.Name = Me.TabPage_RM.Name Then
                Me.chkRM_IncludeWIP.Visible = True
                Me.chkRM_ShowDetail.Visible = True
                Me.chkRM_IncludeWIP.TabStop = True
                Me.chkRM_ShowDetail.TabStop = True

                Me.llbColumnSetting.Visible = False
                Me.llbAddWIP.Visible = False

                Me.rboWO_UseTableView.Visible = False
                Me.rboWO_UseTableView.TabStop = False
                Me.rboWO_UseTreeView.Visible = False
                Me.rboWO_UseTreeView.TabStop = False
            ElseIf Me.TabControl1.SelectedTab.Name = Me.TabPage_WIP.Name Then
                Me.llbColumnSetting.Visible = True
                If Me.TxtRemarks.Enabled = False Then
                    Me.llbAddWIP.Visible = False
                Else
                    Me.llbAddWIP.Visible = True
                End If


                Me.chkRM_IncludeWIP.Visible = False
                Me.chkRM_ShowDetail.Visible = False
                Me.chkRM_IncludeWIP.TabStop = False
                Me.chkRM_ShowDetail.TabStop = False
                Me.rboWO_UseTableView.Visible = False
                Me.rboWO_UseTableView.TabStop = False
                Me.rboWO_UseTreeView.Visible = False
                Me.rboWO_UseTreeView.TabStop = False
            End If

            If Me.DgvItem.RowCount > 0 Then
                If Me.DgvItem.SelectedCells.Count > 0 Then
                    If String.IsNullOrEmpty(Me.DgvItem.Item(Me.ColItem_PrimaryProduct.Index, Me.DgvItem.SelectedCells.Item(0).RowIndex).Value) = False Then
                        If TabControl1.SelectedTab.Name = Me.TabPage_WO.Name Then


                            If Me.TGV_WO.Rows.Count > 0 And rboWO_UseTreeView.Checked = True Then
                                If Me.DgvItem.Item(Me.ColItem_Identifier.Index, Me.DgvItem.SelectedCells.Item(0).RowIndex).Value = Me.TGV_WO.Item(Me.ColWO2_IdentifierJO.Index, 0).Value Then
                                    Exit Sub
                                End If
                            ElseIf Me.dgvItemWO.Rows.Count > 0 And rboWO_UseTableView.Checked = True Then
                                If Me.DgvItem.Item(Me.ColItem_Identifier.Index, Me.DgvItem.SelectedCells.Item(0).RowIndex).Value = Me.dgvItemWO.Item(Me.ColWOTbl_IdentifierJO.Index, 0).Value Then
                                    Exit Sub
                                End If
                            End If
                            LoadDgvItemWO(Me.DgvItem.Item(Me.ColItem_Identifier.Index, Me.DgvItem.SelectedCells.Item(0).RowIndex).Value, _
                            Me.TblItemWO, Me)


                        ElseIf TabControl1.SelectedTab.Name = Me.TabPage_RM.Name Then
                            If Me.DGVRM.Rows.Count > 0 Then
                                If ForceExit Or (ForceReload = False And (Me.DgvItem.Item(Me.ColItem_Identifier.Index, Me.DgvItem.SelectedCells.Item(0).RowIndex).Value = Me.DGVRM.Item(Me.ColRM_IdentifierItem.Index, 0).Value)) Then
                                    ForceExit = False
                                    Exit Sub
                                End If
                            End If

                            LoadDgvItemRM(Me.DgvItem.Item(Me.ColItem_Identifier.Index, Me.DgvItem.SelectedCells.Item(0).RowIndex).Value, _
                            Me.TblItemWO, Me.TblItemRM, Me)
                        ElseIf TabControl1.SelectedTab.Name = Me.TabPage_WIP.Name Then
                            If Me.DGVWIP.Rows.Count > 0 Then
                                If ForceExit Or (ForceReload = False And (Me.DgvItem.Item(Me.ColItem_Identifier.Index, Me.DgvItem.SelectedCells.Item(0).RowIndex).Value = Me.DGVWIP.Item(Me.ColWIP_IdentifierJO.Index, 0).Value)) Then
                                    ForceExit = False
                                    Exit Sub
                                End If
                            End If

                            Me.LoadItemWIP()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadItemWIP()
        Dim RowItemWIP As DseTrans_ManufacturingOrder.tbltrans_mo_item_wipRow
        Dim rowDGV As DataGridViewRow
        Dim RowItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woRow

        Try
            Me.DGVWIP.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DGVWIP.Rows.Clear()
            If Me.DgvItem.SelectedCells.Count <> 0 Then
                Me.TblItemWIP.DefaultView.RowFilter = "[Identifier JO] = '" + Me.DgvItem.Item(Me.ColItem_Identifier.Index, Me.DgvItem.SelectedCells(0).RowIndex).Value.ToString + "'"
                For Each DV As Data.DataRowView In Me.TblItemWIP.DefaultView
                    RowItemWIP = DV.Row
                    Me.DGVWIP.Rows.Add()
                    rowDGV = Me.DGVWIP.Rows(Me.DGVWIP.Rows.Count - 1)

                    rowDGV.Cells(Me.ColWIP_IdentifierJO.Index).Value = RowItemWIP.Identifier_JO
                    rowDGV.Cells(Me.ColWIP_IdentifierWO.Index).Value = RowItemWIP.Identifier_WO
                    rowDGV.Cells(Me.ColWIP_No.Index).Value = rowDGV.Index + 1
                    rowDGV.Cells(Me.ColWIP_PrimaryCodeNo.Index).Value = RowItemWIP.Primary_Product_Code_WIP
                    rowDGV.Cells(Me.ColWIP_CodeNo.Index).Value = Me.ObjProduct.GetProductCode(RowItemWIP.Primary_Product_Code_WIP)
                    rowDGV.Cells(Me.ColWIP_Desc.Index).Value = Me.ObjProduct.GetDescription(RowItemWIP.Primary_Product_Code_WIP)
                    rowDGV.Cells(Me.ColWIP_QtyWIP.Index).Value = Me.ObjCD.DigitGrouping(RowItemWIP.Qty, Me.ObjProduct.GetAngkaDecimal(RowItemWIP.Primary_Product_Code_WIP))
                    rowDGV.Cells(Me.ColWIP_Remarks.Index).Value = RowItemWIP.Remarks
                    rowDGV.Cells(Me.ColWIP_PrimaryUnit.Index).Value = Me.ObjProduct.GetPrimaryUnit(RowItemWIP.Primary_Product_Code_WIP)
                    rowDGV.Cells(Me.ColWIP_UnitCode.Index).Value = Me.ObjUnit.GetUnitCode(rowDGV.Cells(Me.ColWIP_PrimaryUnit.Index).Value)
                    rowDGV.Cells(Me.ColWIP_WONo.Index).Value = Me.GetWONumber(RowItemWIP.Identifier_WO, RowItemWIP.Identifier_JO)

                    Me.TblItemWO.DefaultView.RowFilter = "[Identifier JO] = '" & RowItemWIP.Identifier_JO & "' AND [Identifier] = '" & RowItemWIP.Identifier_WO & "'"
                    If TblItemWO.DefaultView.Count > 0 Then
                        RowItemWO = TblItemWO.DefaultView(0).Row
                        rowDGV.Cells(Me.ColWIP_PrimaryProses.Index).Value = RowItemWO.Primary_Process
                        rowDGV.Cells(Me.ColWIP_Proses.Index).Value = objProses.GetCode(RowItemWO.Primary_Process)
                    Else
                        rowDGV.Cells(Me.ColWIP_PrimaryProses.Index).Value = -1
                        rowDGV.Cells(Me.ColWIP_Proses.Index).Value = ""
                    End If
                Next
            End If

            Me.DGVWIP.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            If Me.DGVWIP.Rows.Count = 0 Then
                Me.llbAddWIP.Text = "Add WIP"
            Else
                Me.llbAddWIP.Text = "Edit WIP"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetWONumber(ByVal vIdentifierWO As String, ByVal vIdentifierJO As String) As String
        Dim vRowItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woRow

        Try
            For Each vRowItemWO In Me.TblItemWO.Rows
                If vRowItemWO.Identifier = vIdentifierWO And _
                vRowItemWO.Identifier_JO = vIdentifierJO Then
                    Return vRowItemWO.WO_Number
                End If
            Next
            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Dim ForceReload As Boolean = False
    Dim ForceExit As Boolean = False
    Private Sub DgvItem_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvItem.SelectionChanged
        If ExecuteSelectionChange = False Then
            Exit Sub
        End If
        Try
            If Me.DgvItem.SelectedRows.Count > 0 And Me.DgvItem.Rows.Count >= 1 Then
                If Me.btnOK.Enabled = False Then
                    Me.btnUP.Enabled = False
                    Me.btnDown.Enabled = False
                Else
                    If Me.DgvItem.RowCount > 2 Then
                        Me.btnUP.Enabled = True
                        Me.btnDown.Enabled = True
                    End If
                End If
                Me.DgvItem.AllowUserToAddRows = False
            Else
                If Me.btnOK.Enabled = True Then
                    Me.DgvItem.AllowUserToAddRows = True
                Else
                    Me.DgvItem.AllowUserToAddRows = False
                End If
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
            End If

            If Me.DgvItem.SelectedCells.Count > 0 Then
                If Me.DgvItem.SelectedCells(0).ColumnIndex >= Me.ColItem_EditBOM.Index Then
                    Exit Sub
                End If
                If String.IsNullOrEmpty(Me.DgvItem.Item(Me.ColItem_PrimaryProduct.Index, Me.DgvItem.SelectedCells.Item(0).RowIndex).Value) = False Then
                    'isi listbox proses untuk 
                    Me.ClbProcess.DataSource = Nothing
                    ObjMO.IsiListBoxProses(Me.ClbProcess, Me.DgvItem.Item(Me.ColItem_Identifier.Index, Me.DgvItem.SelectedCells.Item(0).RowIndex).Value, _
                    Me.TblProsesOrder, Me.TblProses_Filter)
                    'isi dgv detail
                    If TabControl1.SelectedTab.Name = Me.TabPage_WO.Name Then
                        If Me.rboWO_UseTreeView.Checked = True Then
                            If Me.TGV_WO.Rows.Count > 0 Then
                                If ForceExit Or (ForceReload = False And (Me.DgvItem.Item(Me.ColItem_Identifier.Index, Me.DgvItem.SelectedCells.Item(0).RowIndex).Value = Me.TGV_WO.Item(Me.ColWO2_IdentifierJO.Index, 0).Value)) Then
                                    ForceExit = False
                                    Exit Sub
                                End If
                            End If
                        Else
                            If Me.dgvItemWO.Rows.Count > 0 Then
                                If ForceExit Or (ForceReload = False And (Me.DgvItem.Item(Me.ColItem_Identifier.Index, Me.DgvItem.SelectedCells.Item(0).RowIndex).Value = Me.dgvItemWO.Item(Me.ColWOTbl_IdentifierJO.Index, 0).Value)) Then
                                    ForceExit = False
                                    Exit Sub
                                End If
                            End If
                        End If
                        LoadDgvItemWO(Me.DgvItem.Item(Me.ColItem_Identifier.Index, Me.DgvItem.SelectedCells.Item(0).RowIndex).Value, _
                        Me.TblItemWO, Me)
                    ElseIf TabControl1.SelectedTab.Name = Me.TabPage_RM.Name Then
                        'ini buat apa ya??
                        If Me.chkRM_ShowDetail.Checked = True Then
                            If Me.DGVRM.Rows.Count > 0 Then
                                If ForceExit Or (ForceReload = False And (Me.DgvItem.Item(Me.ColItem_Identifier.Index, Me.DgvItem.SelectedCells.Item(0).RowIndex).Value = Me.DGVRM.Item(Me.ColRM_IdentifierItem.Index, 0).Value)) Then
                                    ForceExit = False
                                    Exit Sub
                                End If
                            End If
                        Else
                            If Me.DgvRMRekap.Rows.Count > 0 Then
                                If ForceExit Or (ForceReload = False And (Me.DgvItem.Item(Me.ColItem_Identifier.Index, Me.DgvItem.SelectedCells.Item(0).RowIndex).Value = Me.DgvRMRekap.Item(Me.ColRMRekap_IdentifierJO.Index, 0).Value)) Then
                                    ForceExit = False
                                    Exit Sub
                                End If
                            End If
                        End If

                        LoadDgvItemRM(Me.DgvItem.Item(Me.ColItem_Identifier.Index, Me.DgvItem.SelectedCells.Item(0).RowIndex).Value, _
                        Me.TblItemWO, Me.TblItemRM, Me)
                    ElseIf TabControl1.SelectedTab.Name = Me.TabPage_WIP.Name Then
                        If Me.DGVWIP.Rows.Count > 0 Then
                            If ForceExit Or (ForceReload = False And (Me.DgvItem.Item(Me.ColItem_Identifier.Index, Me.DgvItem.SelectedCells.Item(0).RowIndex).Value = Me.DGVWIP.Item(Me.ColWIP_IdentifierJO.Index, 0).Value)) Then
                                ForceExit = False
                                Exit Sub
                            End If
                        End If
                        Me.LoadItemWIP()
                    End If
                    ForceReload = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub LoadDgvItemWO(ByVal vIdentifierJO As String, ByVal vDataTable_ As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable, _
    ByVal vFrm As FrmTrans_ManufacturingOrder)
        Try
            If Me.rboWO_UseTreeView.Checked = True Then
                Me.LoadDgvItemWO_TreeView(vIdentifierJO, vDataTable_, Me.TGV_WO, vFrm)
            Else
                Me.LoadDgvItemWO_TableView(vIdentifierJO, vDataTable_, Me.dgvItemWO, vFrm)
            End If
            If Me.rboWO_UseTreeView.Checked = True Then
                Me.TGV_WO.BringToFront()
            Else
                Me.dgvItemWO.BringToFront()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub LoadDgvItemWO_TableView(ByVal vIdentifierJO As String, ByVal vDataTable_ As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable, _
    ByVal vDGV As DataGridView, ByVal vFrm As FrmTrans_ManufacturingOrder)
        Dim DataTable_ As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable
        Dim RowTable_ As DseTrans_ManufacturingOrder.tbltrans_mo_item_woRow
        Dim rowDGV As DataGridViewRow
        Dim TBLDummy As New DseTrans_ManufacturingOrder.tbl_codenoDataTable
        Dim rowDummy As DseTrans_ManufacturingOrder.tbl_codenoRow
        Dim TBLDummy_Process As New DseTrans_ManufacturingOrder.tbl_processDataTable
        Dim rowDummy_Process As DseTrans_ManufacturingOrder.tbl_processRow
        Dim Status_ As String = ""

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
                        rowDGV.Cells(Me.ColWOTbl_ShowDetail.Index).Value = "Show"
                        If RowTable_.Status = ClsTrans_ManufacturingOrder.EnumPilihStatusWO.Free Then
                            Status_ = "Free"
                        ElseIf RowTable_.Status = ClsTrans_ManufacturingOrder.EnumPilihStatusWO.Planned Then
                            Status_ = "Planned"
                        ElseIf RowTable_.Status = ClsTrans_ManufacturingOrder.EnumPilihStatusWO.Scheduled Then
                            Status_ = "Scheduled"
                        ElseIf RowTable_.Status = ClsTrans_ManufacturingOrder.EnumPilihStatusWO.On_Progress Then
                            Status_ = "On Progress"
                        ElseIf RowTable_.Status = ClsTrans_ManufacturingOrder.EnumPilihStatusWO.Finish Then
                            Status_ = "Finish"
                        End If
                        rowDGV.Cells(Me.ColWOTbl_Status.Index).Value = Status_
                        rowDGV.Cells(Me.ColWOTbl_WONO.Index).Value = RowTable_.WO_Number
                        rowDGV.Cells(Me.ColWOTbl_WorkUnit.Index).Value = Me.GetWorkUnitCode(RowTable_.Identifier)
                        rowDGV.Cells(Me.ColWOTbl_StartDate.Index).Value = Format(RowTable_.Plan_Start_Time, "dd MMM yyyy")
                        rowDGV.Cells(Me.ColWOTbl_StartTime.Index).Value = Format(RowTable_.Plan_Start_Time, "HH.mm")
                        rowDGV.Cells(Me.ColWOTbl_EndDate.Index).Value = Format(RowTable_.Plan_End_Time, "dd MMM yyyy")
                        rowDGV.Cells(Me.ColWOTbl_EndTime.Index).Value = Format(RowTable_.Plan_End_Time, "HH.mm")

                    End If
                Next
            End If

            If vDGV.RowCount > 0 Then
                If rowDGV.Cells(Me.ColWOTbl_Status.Index).Value.ToString = "Scheduled" Or _
                rowDGV.Cells(Me.ColWOTbl_Status.Index).Value.ToString = "On Progress" Or _
                rowDGV.Cells(Me.ColWOTbl_Status.Index).Value.ToString = "Finish" Then
                    Me.ColWOTbl_StartDate.Visible = True
                    Me.ColWOTbl_StartTime.Visible = True
                    Me.ColWOTbl_EndDate.Visible = True
                    Me.ColWOTbl_EndTime.Visible = True
                Else
                    Me.ColWOTbl_StartDate.Visible = False
                    Me.ColWOTbl_StartTime.Visible = False
                    Me.ColWOTbl_EndDate.Visible = False
                    Me.ColWOTbl_EndTime.Visible = False
                End If
            End If
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub LoadDgvItemWO_TreeView(ByVal vIdentifierJO As String, ByVal vDataTable_ As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable, _
    ByVal vTgvWO As AdvancedDataGridView.TreeGridView, ByVal vFrm As FrmTrans_ManufacturingOrder)
        Dim DataTable_ As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable
        Dim DataRow_ As DseTrans_ManufacturingOrder.tbltrans_mo_item_woRow
        Dim ProsesCode As String = ""
        Dim ProsesDesc As String = ""
        Dim WorkUnitCode As String = ""
        Dim Status_ As String = ""
        Dim Progress_ As Decimal = 0
        Dim ListPrimary As New List(Of Double)

        Try
            vTgvWO.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vTgvWO.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            vTgvWO.AllowUserToResizeColumns = False
            vTgvWO.AllowUserToResizeRows = False

            If vTgvWO.Nodes.Count > 0 Then
                vTgvWO.Nodes.Clear()
                'vTgvWO.Rows.Clear()
            End If

            DataTable_ = vDataTable_.Copy
            vDataTable_.DefaultView.RowFilter = "[Identifier JO] = '" & vIdentifierJO & "' AND [Identifier WO Parent] = '-'"
            If vDataTable_.DefaultView.Count > 0 Then
                DataRow_ = vDataTable_.DefaultView(0).Row
                ProsesCode = objProses.GetCode(DataRow_.Primary_Process)
                ProsesDesc = objProses.GetDesc(DataRow_.Primary_Process)
                WorkUnitCode = Me.GetWorkUnitCode(DataRow_.Identifier) 'objWU.GetCode(DataRow_.Primary_Work_Unit)
                If DataRow_.Status = ClsTrans_ManufacturingOrder.EnumPilihStatusWO.Free Then
                    Status_ = "Free"
                ElseIf DataRow_.Status = ClsTrans_ManufacturingOrder.EnumPilihStatusWO.Planned Then
                    Status_ = "Planned"
                ElseIf DataRow_.Status = ClsTrans_ManufacturingOrder.EnumPilihStatusWO.Scheduled Then
                    Status_ = "Scheduled"
                ElseIf DataRow_.Status = ClsTrans_ManufacturingOrder.EnumPilihStatusWO.On_Progress Then
                    Status_ = "On Progress"
                ElseIf DataRow_.Status = ClsTrans_ManufacturingOrder.EnumPilihStatusWO.Finish Then
                    Status_ = "Finish"
                End If
                Progress_ = Me.ObjCD.DigitGrouping((DataRow_.Qty_Actual_Prod / DataRow_.Qty_Target_Prod) * 100, 2)
                If Progress_ = 0 Then
                    Progress_ = 0
                End If

                vTgvWO.Nodes.Add("", Status_.ToString, Progress_.ToString, DataRow_.WO_Number.ToString, DataRow_.Product_Code, DataRow_.Description, ProsesCode.ToString, ProsesDesc.ToString, WorkUnitCode.ToString, DataRow_.Remarks, "", DataRow_.Primary_Process, DataRow_.Identifier, DataRow_.Identifier_JO)
                FillNodes(DataTable_, DataRow_, vTgvWO.Nodes(0), True, vTgvWO, vFrm)
                'fungsi untuk expand treegridview
                vTgvWO.Nodes(0).Expand()
                Me.Update_TGV_GetExplodeItem(ListPrimary, vTgvWO.Nodes(0))
            End If

            vTgvWO.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Function GetWorkUnitCode(ByVal vIdentifierWO As String) As String
        Dim dta As New DseTrans_ManufacturingOrderTableAdapters.DataTable1TableAdapter
        Dim tbl As New DseTrans_ManufacturingOrder.DataTable1DataTable
        Dim row As DseTrans_ManufacturingOrder.DataTable1Row
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.Fill(tbl, vIdentifierWO) > 0 Then
                If tbl.Rows.Count = 1 Then
                    row = tbl.Rows(0)
                    Return row.work_unit_code
                Else
                    Return "Multiple"
                End If
            End If
            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Sub Update_TGV_GetExplodeItem(ByVal vListPrimary As List(Of Double), ByVal vParentNode As AdvancedDataGridView.TreeGridNode)
        Dim ChildNode As AdvancedDataGridView.TreeGridNode

        Try
            If vParentNode.Nodes.Count > 0 Then
                For Each ChildNode In vParentNode.Nodes
                    If ChildNode.Expand = True Then
                        vListPrimary.Add(ChildNode.Tag)
                    End If
                    If ChildNode.Nodes.Count > 0 Then
                        Me.Update_TGV_GetExplodeItem(vListPrimary, ChildNode)
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub FillNodes(ByVal vDataTable_ As Data.DataTable, ByVal vDataRow_ As DseTrans_ManufacturingOrder.tbltrans_mo_item_woRow, _
    ByVal vNode As AdvancedDataGridView.TreeGridNode, ByVal vTopNode As Boolean, ByVal vTgvWO As AdvancedDataGridView.TreeGridView, _
    ByVal vFrm As FrmTrans_ManufacturingOrder)
        Dim DataTable_ As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable
        Dim DataRow_ As DseTrans_ManufacturingOrder.tbltrans_mo_item_woRow
        Dim item As New AdvancedDataGridView.TreeGridNode()
        Dim ProsesCode As String = ""
        Dim ProsesDesc As String = ""
        Dim WorkUnitCode As String = ""
        Dim Status_ As String = ""
        Dim Progress_ As Decimal = 0

        Try
            DataTable_ = vDataTable_.Copy
            If vTopNode = False Then
                ProsesCode = objProses.GetCode(vDataRow_.Primary_Process)
                ProsesDesc = objProses.GetDesc(vDataRow_.Primary_Process)
                WorkUnitCode = Me.GetWorkUnitCode(vDataRow_.Identifier) 'objWU.GetCode(vDataRow_.Primary_Work_Unit)
                If vDataRow_.Status = ClsTrans_ManufacturingOrder.EnumPilihStatusWO.Free Then
                    Status_ = "Free"
                ElseIf vDataRow_.Status = ClsTrans_ManufacturingOrder.EnumPilihStatusWO.Planned Then
                    Status_ = "Planned"
                ElseIf vDataRow_.Status = ClsTrans_ManufacturingOrder.EnumPilihStatusWO.Scheduled Then
                    Status_ = "Scheduled"
                ElseIf vDataRow_.Status = ClsTrans_ManufacturingOrder.EnumPilihStatusWO.On_Progress Then
                    Status_ = "On Progress"
                ElseIf vDataRow_.Status = ClsTrans_ManufacturingOrder.EnumPilihStatusWO.Finish Then
                    Status_ = "Finish"
                End If
                If vDataRow_.Qty_Target_Prod = 0 Then
                    Progress_ = 0
                Else
                    Progress_ = Me.ObjCD.DigitGrouping((vDataRow_.Qty_Actual_Prod / vDataRow_.Qty_Target_Prod) * 100, 2)
                End If
                If Progress_ = 0 Then
                    Progress_ = 0
                End If

                vNode.Nodes.Add("", Status_, Progress_.ToString, vDataRow_.WO_Number.ToString, vDataRow_.Product_Code, _
                        vDataRow_.Description, ProsesCode.ToString, ProsesDesc.ToString, WorkUnitCode.ToString, vDataRow_.Remarks, "", _
                        vDataRow_.Primary_Process, vDataRow_.Identifier, vDataRow_.Identifier_JO)
            End If

            vDataTable_.DefaultView.RowFilter = "[Identifier WO Parent] = '" & vDataRow_.Identifier & "'"
            For ii As Integer = 0 To vDataTable_.DefaultView.Count - 1
                DataRow_ = vDataTable_.DefaultView(ii).Row
                If vTopNode = False Then
                    FillNodes(DataTable_, DataRow_, vNode.Nodes(vNode.Nodes.Count - 1), False, vTgvWO, vFrm)
                Else
                    FillNodes(DataTable_, DataRow_, vNode, False, vTgvWO, vFrm)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub LoadDgvItemRM(ByVal vIdentifierJO As String, ByVal vTblItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable, _
    ByVal vTblItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmDataTable, ByVal vFrm As FrmTrans_ManufacturingOrder)
        Try
            If Me.chkRM_ShowDetail.Checked = True Then
                Me.LoadRMDetail(vIdentifierJO, vTblItemWO, vTblItemRM, Me.DGVRM, vFrm)
                Me.DGVRM.BringToFront()
            Else
                Me.LoadRMRekap(vIdentifierJO, vTblItemWO, vTblItemRM, Me.DgvRMRekap, vFrm)
                Me.DgvRMRekap.BringToFront()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub LoadRMRekap(ByVal vIdentifierJO As String, ByVal vTblItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable, _
    ByVal vTblItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmDataTable, ByVal vDgvRM As DataGridView, ByVal vFrm As FrmTrans_ManufacturingOrder)
        Dim rowItemRM, rowItemRMB4 As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmRow
        Dim rowD As DataGridViewRow
        Dim ObjProses As New Core_Module.ClsMasterProses
        Dim vDecProd_Item As Integer = 0
        Dim Idx As Integer = 0

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            vDgvRM.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDgvRM.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            vDgvRM.AllowUserToResizeColumns = False
            vDgvRM.AllowUserToResizeRows = False

            If vDgvRM.Rows.Count > 0 Then
                vDgvRM.Rows.Clear()
            End If

            If Me.chkRM_IncludeWIP.Checked = False Then
                vTblItemRM.DefaultView.RowFilter = "[Identifier JO] = '" & vIdentifierJO & "' AND [is wip] = 0 AND [is replacement] = 0"
                Me.ColRMRekap_QtyWIP.Visible = False
            Else
                vTblItemRM.DefaultView.RowFilter = "[Identifier JO] = '" & vIdentifierJO & "' AND [is replacement] = 0"
                Me.ColRMRekap_QtyWIP.Visible = True
            End If

            vTblItemRM.DefaultView.Sort = "[Primary Product Code RM] ASC"
            If vTblItemRM.DefaultView.Count > 0 Then
                vDgvRM.AllowUserToAddRows = False
                For Each dv As Data.DataRowView In vTblItemRM.DefaultView
                    rowItemRM = dv.Row
                    With rowItemRM
                        If Me.ChkInsufficientRM.Checked = True Then
                            If .GP_Target > .GP_Actual Then
                                GoTo SkipRow
                            End If
                        End If

                        If Idx = 0 Then
                            vDgvRM.Rows.Add()
                            rowD = vDgvRM.Rows(vDgvRM.RowCount - 1)
                            rowD.Cells(vFrm.ColRMRekap_No.Index).Value = vDgvRM.Rows.Count.ToString
                            vDecProd_Item = ObjProduct.GetAngkaDecimal(.Primary_Product_Code_RM)
                            rowD.Cells(vFrm.ColRMRekap_PrimaryCodeNo.Index).Value = .Primary_Product_Code_RM
                            rowD.Cells(vFrm.ColRMRekap_PrimaryUnit.Index).Value = ObjProduct.GetPrimaryUnit(.Primary_Product_Code_RM)
                            rowD.Cells(vFrm.ColRMRekap_CodeNo.Index).Value = ObjProduct.GetProductCode(.Primary_Product_Code_RM)
                            rowD.Cells(vFrm.ColRMRekap_Desc.Index).Value = ObjProduct.GetDescription(.Primary_Product_Code_RM)
                            rowD.Cells(vFrm.ColRMRekap_Unit.Index).Value = ObjProduct.GetUnitCode(.Primary_Product_Code_RM)
                            rowD.Cells(vFrm.ColRMRekap_QtyAdditional.Index).Value = Me.ObjCD.DigitGrouping(.Additional_Qty, vDecProd_Item)
                            rowD.Cells(vFrm.ColRMRekap_QtyAllocated.Index).Value = Me.ObjCD.DigitGrouping(.Allocated_Qty, vDecProd_Item)
                            rowD.Cells(vFrm.ColRMRekap_QtyBalance.Index).Value = Me.ObjCD.DigitGrouping(.Balance_Qty, vDecProd_Item)
                            rowD.Cells(vFrm.ColRMRekap_QtyWIP.Index).Value = Me.ObjCD.DigitGrouping(.WIP_Qty, vDecProd_Item)
                            rowD.Cells(vFrm.ColRMRekap_QtyReplcaed.Index).Value = Me.ObjCD.DigitGrouping(.Replace_Qty, vDecProd_Item)
                            rowD.Cells(vFrm.ColRMRekap_QtyRequired.Index).Value = Me.ObjCD.DigitGrouping(.Required_Qty, vDecProd_Item)
                            rowD.Cells(vFrm.ColRMRekap_QtyTarget.Index).Value = Me.ObjCD.DigitGrouping(.Target_Qty, vDecProd_Item)
                            rowD.Cells(vFrm.ColRMRekap_QtyUsed.Index).Value = Me.ObjCD.DigitGrouping(.Used_Qty, vDecProd_Item)
                            rowD.Cells(vFrm.ColRMRekap_IdentifierJO.Index).Value = vIdentifierJO
                        Else
                            rowItemRMB4 = vTblItemRM.DefaultView(Idx - 1).Row
                            If rowItemRM.Primary_Product_Code_RM = rowItemRMB4.Primary_Product_Code_RM Then
                                rowD = vDgvRM.Rows(vDgvRM.RowCount - 1)
                                'hanya tambahkan qty
                                rowD.Cells(vFrm.ColRMRekap_QtyAdditional.Index).Value = Me.ObjCD.DigitGrouping(.Additional_Qty + CDec(rowD.Cells(vFrm.ColRMRekap_QtyAdditional.Index).Value), vDecProd_Item)
                                rowD.Cells(vFrm.ColRMRekap_QtyAllocated.Index).Value = Me.ObjCD.DigitGrouping(.Allocated_Qty + CDec(rowD.Cells(vFrm.ColRMRekap_QtyAllocated.Index).Value), vDecProd_Item)
                                rowD.Cells(vFrm.ColRMRekap_QtyBalance.Index).Value = Me.ObjCD.DigitGrouping(.Balance_Qty + CDec(rowD.Cells(vFrm.ColRMRekap_QtyBalance.Index).Value), vDecProd_Item)
                                rowD.Cells(vFrm.ColRMRekap_QtyRequired.Index).Value = Me.ObjCD.DigitGrouping(.Required_Qty + CDec(rowD.Cells(vFrm.ColRMRekap_QtyRequired.Index).Value), vDecProd_Item)
                                rowD.Cells(vFrm.ColRMRekap_QtyTarget.Index).Value = Me.ObjCD.DigitGrouping(.Target_Qty + CDec(rowD.Cells(vFrm.ColRMRekap_QtyTarget.Index).Value), vDecProd_Item)
                                rowD.Cells(vFrm.ColRMRekap_QtyUsed.Index).Value = Me.ObjCD.DigitGrouping(.Used_Qty + CDec(rowD.Cells(vFrm.ColRMRekap_QtyUsed.Index).Value), vDecProd_Item)
                                rowD.Cells(vFrm.ColRMRekap_QtyWIP.Index).Value = Me.ObjCD.DigitGrouping(.WIP_Qty + rowD.Cells(vFrm.ColRMRekap_QtyWIP.Index).Value, vDecProd_Item)
                                rowD.Cells(vFrm.ColRMRekap_QtyReplcaed.Index).Value = Me.ObjCD.DigitGrouping(.Replace_Qty + rowD.Cells(vFrm.ColRMRekap_QtyReplcaed.Index).Value, vDecProd_Item)
                            Else
                                'tambahkan baris baru
                                vDgvRM.Rows.Add()
                                rowD = vDgvRM.Rows(vDgvRM.RowCount - 1)
                                rowD.Cells(vFrm.ColRMRekap_No.Index).Value = vDgvRM.Rows.Count.ToString
                                vDecProd_Item = ObjProduct.GetAngkaDecimal(.Primary_Product_Code_RM)
                                rowD.Cells(vFrm.ColRMRekap_PrimaryCodeNo.Index).Value = .Primary_Product_Code_RM
                                rowD.Cells(vFrm.ColRMRekap_PrimaryUnit.Index).Value = ObjProduct.GetPrimaryUnit(.Primary_Product_Code_RM)
                                rowD.Cells(vFrm.ColRMRekap_CodeNo.Index).Value = ObjProduct.GetProductCode(.Primary_Product_Code_RM)
                                rowD.Cells(vFrm.ColRMRekap_Desc.Index).Value = ObjProduct.GetDescription(.Primary_Product_Code_RM)
                                rowD.Cells(vFrm.ColRMRekap_Unit.Index).Value = ObjProduct.GetUnitCode(.Primary_Product_Code_RM)
                                rowD.Cells(vFrm.ColRMRekap_QtyAdditional.Index).Value = Me.ObjCD.DigitGrouping(.Additional_Qty, vDecProd_Item)
                                rowD.Cells(vFrm.ColRMRekap_QtyAllocated.Index).Value = Me.ObjCD.DigitGrouping(.Allocated_Qty, vDecProd_Item)
                                rowD.Cells(vFrm.ColRMRekap_QtyBalance.Index).Value = Me.ObjCD.DigitGrouping(.Balance_Qty, vDecProd_Item)
                                rowD.Cells(vFrm.ColRMRekap_QtyWIP.Index).Value = Me.ObjCD.DigitGrouping(.WIP_Qty, vDecProd_Item)
                                rowD.Cells(vFrm.ColRMRekap_QtyReplcaed.Index).Value = Me.ObjCD.DigitGrouping(.Replace_Qty, vDecProd_Item)
                                rowD.Cells(vFrm.ColRMRekap_QtyRequired.Index).Value = Me.ObjCD.DigitGrouping(.Required_Qty, vDecProd_Item)
                                rowD.Cells(vFrm.ColRMRekap_QtyTarget.Index).Value = Me.ObjCD.DigitGrouping(.Target_Qty, vDecProd_Item)
                                rowD.Cells(vFrm.ColRMRekap_QtyUsed.Index).Value = Me.ObjCD.DigitGrouping(.Used_Qty, vDecProd_Item)
                                rowD.Cells(vFrm.ColRMRekap_IdentifierJO.Index).Value = vIdentifierJO
                            End If
                        End If
SkipRow:
                    End With
                    Idx += 1
                Next
            End If

            vDgvRM.Sort(vDgvRM.Columns(vFrm.ColRMRekap_CodeNo.Index), System.ComponentModel.ListSortDirection.Ascending)

            'buang baris yg qty req =0, qty add =0, dan qty wip =0
            For mn As Integer = Me.DgvRMRekap.Rows.Count - 1 To 0 Step -1
                rowD = Me.DgvRMRekap.Rows(mn)
                If rowD.Cells(Me.ColRMRekap_QtyWIP.Index).Value = 0 And _
                rowD.Cells(Me.ColRMRekap_QtyRequired.Index).Value = 0 And _
                rowD.Cells(Me.ColRMRekap_QtyAdditional.Index).Value = 0 Then
                    Me.DgvRMRekap.Rows.Remove(rowD)
                End If
            Next
            Me.ObjCD.SetNumber(vDgvRM)

            vDgvRM.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            'vDgvRM.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjProses = Nothing
        End Try
    End Sub

    Public Sub LoadRMDetail(ByVal vIdentifierJO As String, ByVal vTblItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable, _
    ByVal vTblItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmDataTable, ByVal vDgvRM As DataGridView, ByVal vFrm As FrmTrans_ManufacturingOrder)
        Dim rowItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woRow
        Dim rowItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmRow
        Dim rowD As DataGridViewRow
        Dim ObjProses As New Core_Module.ClsMasterProses
        Dim vDecProd_Item As Integer = 0
        Dim Lanjut As Boolean = False

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            vDgvRM.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDgvRM.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            vDgvRM.AllowUserToResizeColumns = False
            vDgvRM.AllowUserToResizeRows = False

            If vDgvRM.Rows.Count > 0 Then
                vDgvRM.Rows.Clear()
            End If

            If Me.chkRM_IncludeWIP.Checked = False Then
                vTblItemRM.DefaultView.RowFilter = "[Identifier JO] = '" & vIdentifierJO & "' AND [is wip] = 0 AND [is replacement] = 0"
                Me.ColRM_QtyWIP.Visible = False
            Else
                vTblItemRM.DefaultView.RowFilter = "[Identifier JO] = '" & vIdentifierJO & "' AND [is replacement] = 0"
                Me.ColRM_QtyWIP.Visible = True
            End If

            vTblItemRM.DefaultView.Sort = "[Primary Product Code RM] ASC"
            If vTblItemRM.DefaultView.Count > 0 Then
                vDgvRM.AllowUserToAddRows = False
                For Each dv As Data.DataRowView In vTblItemRM.DefaultView
                    rowItemRM = dv.Row
                    Lanjut = True

                    If rowItemRM.is_wip = 0 Then
                        If rowItemRM.Required_Qty = 0 And rowItemRM.Additional_Qty = 0 Then
                            Lanjut = False
                        End If
                    Else
                        If rowItemRM.Required_Qty = 0 And rowItemRM.Additional_Qty = 0 And rowItemRM.WIP_Qty = 0 Then
                            Lanjut = False
                        End If
                    End If
                    If Lanjut = True Then
                        With rowItemRM
                            If Me.ChkInsufficientRM.Checked = True Then
                                If .GP_Target > .GP_Actual Then
                                    GoTo SkipRow
                                End If
                            End If
                            vDgvRM.Rows.Add()
                            rowD = vDgvRM.Rows(vDgvRM.RowCount - 1)
                            rowD.Cells(vFrm.ColRM_No.Index).Value = vDgvRM.Rows.Count.ToString

                            vDecProd_Item = ._decimal ' ObjProduct.GetAngkaDecimal(.Primary_Product_Code_RM )
                            rowD.Cells(vFrm.ColRM_PrimaryCodeNo.Index).Value = .Primary_Product_Code_RM
                            vTblItemWO.DefaultView.RowFilter = "[Identifier JO] = '" & .Identifier_JO.ToString & "' AND [Identifier] = '" & .Identifier_WO & "'"
                            If vTblItemWO.DefaultView.Count > 0 Then
                                rowItemWO = vTblItemWO.DefaultView(0).Row
                                rowD.Cells(vFrm.ColRM_PrimaryProcess.Index).Value = rowItemWO.Primary_Process
                                rowD.Cells(vFrm.ColRM_Process.Index).Value = ObjProses.GetCode(rowItemWO.Primary_Process)
                            End If
                            rowD.Cells(vFrm.ColRM_PrimaryUnit.Index).Value = .unit_primary 'ObjProduct.GetPrimaryUnit(.Primary_Product_Code_RM)
                            rowD.Cells(vFrm.ColRM_WONumber.Index).Value = ObjMO.GetWONumber_InTbl(.Identifier_WO, vTblItemWO, .Identifier_JO)
                            rowD.Cells(vFrm.ColRM_CodeNo.Index).Value = .rm_code 'ObjProduct.GetProductCode(.Primary_Product_Code_RM)
                            rowD.Cells(vFrm.ColRM_Desc.Index).Value = .rm_desc 'ObjProduct.GetDescription(.Primary_Product_Code_RM)
                            rowD.Cells(vFrm.ColRM_Unit.Index).Value = .unit_code 'ObjProduct.GetUnitCode(.Primary_Product_Code_RM)
                            rowD.Cells(vFrm.ColRM_PrimaryStandByLocation.Index).Value = .Primary_StandBy_Location
                            rowD.Cells(vFrm.ColRM_PrimaryPickUpLocation.Index).Value = .Primary_PickUp_Location
                            rowD.Cells(vFrm.ColRM_RequiredDate.Index).Value = Format(.required_date, "dd MMM yyyy")

                            'rowD.Cells(vFrm.ColRM_StandByLocation.Index).Value = objLoc.GetStorageAreaCode(.Primary_StandBy_Location) -> dicancel karena location ditentukan setelah scheduling berdasarkan work unit
                            'rowD.Cells(vFrm.ColRM_PickUpLocation.Index).Value = objLoc.GetStorageAreaCode(.Primary_PickUp_Location)
                            rowD.Cells(vFrm.ColRM_StandByLocation.Index).Value = ""
                            rowD.Cells(vFrm.ColRM_PickUpLocation.Index).Value = ""
                            rowD.Cells(vFrm.ColRM_QtyAdditional.Index).Value = Me.ObjCD.DigitGrouping(.Additional_Qty, vDecProd_Item)
                            rowD.Cells(vFrm.ColRM_QtyAllocated.Index).Value = Me.ObjCD.DigitGrouping(.Allocated_Qty, vDecProd_Item)
                            rowD.Cells(vFrm.ColRM_QtyBalance.Index).Value = Me.ObjCD.DigitGrouping(.Balance_Qty, vDecProd_Item)
                            rowD.Cells(vFrm.ColRM_QtyWIP.Index).Value = Me.ObjCD.DigitGrouping(.WIP_Qty, vDecProd_Item)
                            rowD.Cells(vFrm.ColRM_QtyReplaced.Index).Value = Me.ObjCD.DigitGrouping(.Replace_Qty, vDecProd_Item)
                            rowD.Cells(vFrm.ColRM_QtyRequired.Index).Value = Me.ObjCD.DigitGrouping(.Required_Qty, vDecProd_Item)
                            rowD.Cells(vFrm.ColRM_QtyTarget.Index).Value = Me.ObjCD.DigitGrouping(.Target_Qty, vDecProd_Item)
                            rowD.Cells(vFrm.ColRM_QtyUsed.Index).Value = Me.ObjCD.DigitGrouping(.Used_Qty, vDecProd_Item)
                            rowD.Cells(vFrm.ColRM_GPTarget.Index).Value = Me.ObjCD.DigitGrouping(.GP_Target, 2)
                            rowD.Cells(vFrm.ColRM_GPActual.Index).Value = Me.ObjCD.DigitGrouping(.GP_Actual, 2)
                            rowD.Cells(vFrm.ColRM_IdentifierWO.Index).Value = .Identifier_WO
                            rowD.Cells(vFrm.ColRM_IdentifierItem.Index).Value = vIdentifierJO
                            rowD.Cells(vFrm.ColRM_Remarks.Index).Value = .Remarks
                            rowD.Cells(vFrm.ColRM_IsWIP.Index).Value = .is_wip
SkipRow:
                        End With
                    End If
                Next
            End If

            vDgvRM.Sort(vDgvRM.Columns(vFrm.ColRM_WONumber.Index), System.ComponentModel.ListSortDirection.Ascending)
            Me.ObjCD.SetNumber(vDgvRM)

            vDgvRM.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            'vDgvRM.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjProses = Nothing
        End Try
    End Sub

    Private Sub DgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvItem.CellBeginEdit
        Dim ObjUser As New Core_Module.ClsUser

        Try
            If Me.DgvItem.Item(ColItem_CodeNo.Index, e.RowIndex).Value = "" Then
                If e.ColumnIndex <> ColItem_CodeNo.Index Then
                    e.Cancel = True
                Else
                    If ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_ManufacturingOrder_AddItem") = False Then
                        MsgBox("You don't have access to perform this action.", MsgBoxStyle.Information, "Information")
                        e.Cancel = True
                        Exit Sub
                    Else
                        Me.Buffer = Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                    End If
                End If
            Else
                If e.ColumnIndex = ColItem_CodeNo.Index Or e.ColumnIndex = ColItem_QtyTarget.Index Or _
                e.ColumnIndex = ColItem_Remarks.Index Or e.ColumnIndex = Me.ColItem_QtyProduce.Index Then
                    If ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_ManufacturingOrder_AddItem") = False Then
                        MsgBox("You don't have access to perform this action.", MsgBoxStyle.Information, "Information")
                        e.Cancel = True
                        Exit Sub
                    Else
                        If Me.Baru = False Then
                            If e.ColumnIndex = ColItem_CodeNo.Index Or e.ColumnIndex = ColItem_QtyTarget.Index Or _
                            e.ColumnIndex = Me.ColItem_QtyProduce.Index Then
                                If Me.ObjMO.IsScheduled(Me.TxtPrimary.Text, Me.DgvItem.Item(Me.ColItem_Identifier.Index, e.RowIndex).Value) = True Then
                                    MsgBox("This JO is scheduled, you can't edit this data.", MsgBoxStyle.Information, "Information")
                                    e.Cancel = True
                                    Exit Sub
                                Else
                                    Me.Buffer = Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                                End If
                            ElseIf e.ColumnIndex = ColItem_Remarks.Index Then
                                Me.Buffer = Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                            End If
                        Else
                            Me.Buffer = Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                            If e.ColumnIndex = ColItem_QtyTarget.Index Or _
                            e.ColumnIndex = Me.ColItem_QtyProduce.Index Then
                                Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.BuangNolDibelakangDecimal(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                            End If
                        End If
                    End If
                ElseIf e.ColumnIndex = ColItem_TargetFinishDate.Index Then
                    If Me.DgvItem.Item(Me.ColItem_TargetFinishDateR1.Index, e.RowIndex).Value.ToString = "-" Then
                        If ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_ManufacturingOrder_EditTargetFinish") = False Then
                            MsgBox("You don't have access to perform this action.", MsgBoxStyle.Information, "Information")
                            e.Cancel = True
                            Exit Sub
                        Else
                            Me.Buffer = Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                        End If
                    Else
                        e.Cancel = True
                    End If
                ElseIf e.ColumnIndex = ColItem_TargetFinishDateR1.Index Then
                    If Me.vEdit = True Then
                        If IsDate(Me.DgvItem.Item(ColItem_TargetFinishDate.Index, e.RowIndex).Value) = True Then
                            If Me.DgvItem.Item(Me.ColItem_TargetFinishDateR2.Index, e.RowIndex).Value.ToString = "-" Then
                                If ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_ManufacturingOrder_EditTargetFinish") = False Then
                                    MsgBox("You don't have access to perform this action.", MsgBoxStyle.Information, "Information")
                                    e.Cancel = True
                                    Exit Sub
                                Else
                                    Me.Buffer = Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                                End If
                            Else
                                e.Cancel = True
                            End If
                        Else
                            e.Cancel = True
                        End If
                    Else
                        e.Cancel = True
                    End If
                ElseIf e.ColumnIndex = ColItem_TargetFinishDateR2.Index Then
                    If Me.vEdit = True Then
                        If IsDate(Me.DgvItem.Item(ColItem_TargetFinishDateR1.Index, e.RowIndex).Value) = True Then
                            If ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_ManufacturingOrder_EditTargetFinish") = False Then
                                MsgBox("You don't have access to perform this action.", MsgBoxStyle.Information, "Information")
                                e.Cancel = True
                                Exit Sub
                            Else
                                Me.Buffer = Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                            End If
                        Else
                            e.Cancel = True
                        End If
                    Else
                        e.Cancel = True
                    End If
                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvItem.CellEndEdit
        Dim row As DataGridViewRow = Me.DgvItem.Rows(e.RowIndex)
        Try
            With row
                If IsNothing(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                        Exit Sub
                    End If
                    If e.ColumnIndex = Me.ColItem_QtyProduce.Index Then
                        If IsNumeric(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            If IsNumeric(Me.Buffer) = True Then
                                Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.Buffer, Me.DgvItem.Item(Me.ColItem_Decimal.Index, e.RowIndex).Value)
                            Else
                                Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = 0
                            End If
                        Else
                            If Me.ObjCD.CekBilanganPositif(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                                MsgBox("Please enter positive number for qty produce.", MsgBoxStyle.Information, _
                                "Information")
                                Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.Buffer, Me.DgvItem.Item(Me.ColItem_Decimal.Index, e.RowIndex).Value)
                                Exit Sub
                            Else
                                Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value, Me.DgvItem.Item(Me.ColItem_Decimal.Index, e.RowIndex).Value)
                                Dim Buffer0 As Decimal = CDec(Me.DgvItem.Item(Me.ColItem_QtyTarget.Index, e.RowIndex).Value)
                                Me.DgvItem.Item(Me.ColItem_QtyTarget.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value * _
                                (Me.ObjMO.GetGoodPercentageOfFG(Me.DgvItem.Item(Me.ColItem_PrimaryProduct.Index, e.RowIndex).Value, Me.DgvItem.Item(Me.ColItem_Identifier.Index, e.RowIndex).Value, Me.TblItemBom) / 100), _
                                Me.DgvItem.Item(Me.ColItem_Decimal.Index, e.RowIndex).Value, Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)
                                Dim JmlKurang As Decimal = 0
                                If CDec(Buffer0) > CDec(Me.DgvItem.Item(Me.ColItem_QtyTarget.Index, e.RowIndex).Value) Then
                                    'untuk item delivery, jika qty turun, maka hapuskan mulai dari baris paling bawah
                                    JmlKurang = CDec(Buffer0) - CDec(Me.DgvItem.Item(Me.ColItem_QtyTarget.Index, e.RowIndex).Value)
                                    Me.TBLItemDelivery.DefaultView.RowFilter = "[identifier jo] = '" + Me.DgvItem.Item(Me.ColItem_Identifier.Index, e.RowIndex).Value + "'"
                                    Me.TBLItemDelivery.DefaultView.Sort = "[delivery date] DESC"
                                    While JmlKurang > 0
                                        If Me.TBLItemDelivery.DefaultView.Count > 0 Then
                                            Me.RowItemDelivery = Me.TBLItemDelivery.DefaultView(0).Row
                                            If Me.RowItemDelivery.delivery_qty > JmlKurang Then
                                                Me.RowItemDelivery.delivery_qty -= JmlKurang
                                                JmlKurang = 0
                                            Else
                                                JmlKurang -= Me.RowItemDelivery.delivery_qty
                                                Me.RowItemDelivery.delivery_qty = 0
                                            End If
                                            If Me.RowItemDelivery.delivery_qty = 0 Then
                                                Me.TBLItemDelivery.Rows.Remove(Me.RowItemDelivery)
                                            End If
                                        End If
                                    End While
                                ElseIf CDec(Buffer0) < CDec(Me.DgvItem.Item(Me.ColItem_QtyTarget.Index, e.RowIndex).Value) Then
                                    'untuk item delivery, jika qty naik, maka tambahkan sisanya pada baris paling bawah
                                    JmlKurang = CDec(Me.DgvItem.Item(Me.ColItem_QtyTarget.Index, e.RowIndex).Value) - CDec(Buffer0)
                                    Me.TBLItemDelivery.DefaultView.RowFilter = "[identifier jo] = '" + Me.DgvItem.Item(Me.ColItem_Identifier.Index, e.RowIndex).Value + "'"
                                    Me.TBLItemDelivery.DefaultView.Sort = "[delivery date] DESC"
                                    If Me.TBLItemDelivery.DefaultView.Count > 0 Then
                                        Me.RowItemDelivery = Me.TBLItemDelivery.DefaultView(0).Row
                                        Me.RowItemDelivery.delivery_qty += JmlKurang
                                    End If
                                End If

                                Me.IsiTblItemWORM(-1, Me.DgvItem.Item(Me.ColItem_PrimaryProduct.Index, e.RowIndex).Value, Me.DgvItem.Item(Me.ColItem_QtyTarget.Index, e.RowIndex).Value, _
                                    Me.TblItemWO, Me.TblItemRM, Me.TblItemBom, Me.DgvItem.Item(Me.ColItem_Identifier.Index, e.RowIndex).Value, "-", "-", Me.TblItemWIP, _
                                    Me.GetRequiredDate(Me.DgvItem.Item(Me.ColItem_Identifier.Index, e.RowIndex).Value), False)
                                LoadDgvItemRM(Me.DgvItem.Item(Me.ColItem_Identifier.Index, Me.DgvItem.SelectedCells.Item(0).RowIndex).Value, Me.TblItemWO, Me.TblItemRM, Me)

                                Exit Sub
                            End If
                        End If
                    ElseIf e.ColumnIndex = Me.ColItem_CodeNo.Index Then
                        Me.EndEditProductCode(sender, e)
                    ElseIf e.ColumnIndex = Me.ColItem_QtyTarget.Index Then
                        Me.EndEditQtyTarget(sender, e)

                        LoadDgvItemRM(Me.DgvItem.Item(Me.ColItem_Identifier.Index, Me.DgvItem.SelectedCells.Item(0).RowIndex).Value, _
                        Me.TblItemWO, Me.TblItemRM, Me)
                    ElseIf e.ColumnIndex = Me.ColItem_TargetFinishDate.Index Or e.ColumnIndex = Me.ColItem_TargetFinishDateR1.Index _
                    Or e.ColumnIndex = Me.ColItem_TargetFinishDateR2.Index Or e.ColumnIndex = Me.ColItem_ActualFinishDate.Index Then
                        If e.ColumnIndex = Me.ColItem_TargetFinishDateR1.Index Then
                            If Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim = "-" And _
                            Me.DgvItem.Item(Me.ColItem_TargetFinishDateR2.Index, e.RowIndex).Value.ToString.Trim = "-" Then
                                Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = "-"
                                Exit Sub
                            End If
                        ElseIf e.ColumnIndex = Me.ColItem_TargetFinishDateR2.Index Then
                            If Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim = "-" Then
                                Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = "-"
                                Exit Sub
                            End If
                        End If
                        Me.EndEditFinishDate(sender, e)
                    ElseIf e.ColumnIndex = Me.ColItem_Remarks.Index Then
                        Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Replace("'", "").Replace("#", "").Replace("*", "").Replace("`", "").Replace("$", "").Replace("%", "")
                    End If
                Else
                    If e.ColumnIndex = Me.ColItem_Remarks.Index Then
                        Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    ElseIf e.ColumnIndex = Me.ColItem_CodeNo.Index Then
                        If String.IsNullOrEmpty(Me.Buffer) = True Then
                            Me.EraseRowIndex = e.RowIndex
                            Me.BackgroundWorker2.RunWorkerAsync()
                        Else
                            Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                    ElseIf e.ColumnIndex = Me.ColItem_TargetFinishDateR1.Index Then
                        If Me.DgvItem.Item(Me.ColItem_TargetFinishDateR2.Index, e.RowIndex).Value = "-" Then
                            Me.DgvItem.Item(Me.ColItem_TargetFinishDateR1.Index, e.RowIndex).Value = "-"
                        Else
                            Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        End If
                    ElseIf e.ColumnIndex = Me.ColItem_TargetFinishDateR2.Index Then
                        Me.DgvItem.Item(Me.ColItem_TargetFinishDateR2.Index, e.RowIndex).Value = "-"
                    Else
                        Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                    End If
                End If
                'If Baru = False Then
                '    For ij As Integer = 0 To DgvItem.Rows.Count - 1
                '        If Me.DgvItem.Item(Me.ColItem_CodeNo.Index, ij).Value <> "" Then
                '            Me.DtpTglTransaksi.Enabled = False
                '            Exit For
                '        Else
                '            Me.DtpTglTransaksi.Enabled = True
                '            Exit For
                '        End If
                '    Next
                'End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function getMaxNoUrutJO() As Integer
        Dim DTA As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_moTableAdapter
        Dim urutan As Integer = 0
        Dim Max_Urutan_DB As Integer = 0

        Try
            For Each rowD As DataGridViewRow In Me.DgvItem.Rows
                If IsNumeric(rowD.Cells(Me.ColItem_Urutan.Index).Value) = True Then
                    If rowD.Cells(Me.ColItem_Urutan.Index).Value > urutan Then
                        urutan = rowD.Cells(Me.ColItem_Urutan.Index).Value
                    End If
                End If
            Next
            Return urutan
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function getMaxNoUrutWO(ByVal vTblItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable, Optional ByVal vFilter As String = "") As Integer
        Dim urutan As Integer = 0
        Dim TblItemWODumm As New DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable
        Dim rowItemWODumm As DseTrans_ManufacturingOrder.tbltrans_mo_item_woRow

        Try
            TblItemWODumm = vTblItemWO.Copy

            TblItemWODumm.DefaultView.RowFilter = Nothing
            TblItemWODumm.DefaultView.Sort = "[Urutan] Desc"
            If TblItemWODumm.DefaultView.Count > 0 Then
                rowItemWODumm = TblItemWODumm.DefaultView(0).Row
                If IsNumeric(rowItemWODumm.Urutan) = True Then
                    urutan = rowItemWODumm.Urutan
                End If
            End If

            If vFilter <> "" Then
                Me.TblItemWO.DefaultView.RowFilter = vFilter
            End If
            Return urutan
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    'Public Function getMaxUrutanSebelumnya() As Integer
    '    Dim urutan As Integer = 0
    '    Try
    '        For ij As Integer = 0 To Me.DgvItem.Rows.Count - 1
    '            If IsNumeric(Me.DgvItem.Item(Me.ColItem_Urutan.Index, ij).Value) = True Then
    '                If ij = Me.DgvItem.Rows.Count - 2 Then
    '                    urutan = CDec(Me.DgvItem.Item(Me.ColItem_Urutan.Index, ij - 1).Value)
    '                End If
    '            End If
    '        Next
    '        Return urutan
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '        Return urutan
    '    End Try
    'End Function

    Public Sub IsiAllTblItem(ByVal vPrimaryProduct As Double, ByVal vQtyTarget As Decimal, _
    ByVal vIdentifierJO As String, ByVal vUrutanJO As Integer, ByVal vShowProsesOrder As Boolean, ByVal vRmReqDate As Date, ByVal vSetRMReqDate As Boolean)
        Dim rowItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woRow
        Dim rowItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmRow

        Try
            'isi tbl item BOM
            Me.IsiTblItemBOM(vPrimaryProduct, vIdentifierJO, Me.TblItemBom, -1, "-", Me.TxtPrimary.Text)

            'isi tbl WO dan RM
            Me.TblItemWO.DefaultView.RowFilter = "[Identifier JO] = '" & vIdentifierJO & "'"
            While Me.TblItemWO.DefaultView.Count > 0
                rowItemWO = Me.TblItemWO.DefaultView(0).Row
                Me.TblItemWO.Rows.Remove(rowItemWO)
            End While
            Me.TblItemRM.DefaultView.RowFilter = "[Identifier JO] = '" & vIdentifierJO & "'"
            While Me.TblItemRM.DefaultView.Count > 0
                rowItemRM = Me.TblItemRM.DefaultView(0).Row
                Me.TblItemRM.Rows.Remove(rowItemRM)
            End While

            Me.IsiTblItemWORM(-1, vPrimaryProduct, vQtyTarget, Me.TblItemWO, Me.TblItemRM, Me.TblItemBom, vIdentifierJO, "-", "-", Me.TblItemWIP, vRmReqDate, vSetRMReqDate)

            'isi tbl proses order
            Me.IsiTblProsesOrder(vPrimaryProduct, Me.TblProsesOrder, vIdentifierJO)

            'tampilkan form proses order untuk mengatur urutan proses
            Dim Frm As New FrmTrans_ManufacturingOrder_ProcessOrder
            Frm.vTblProsesOrder = TblProsesOrder
            Frm.vUrutan = vUrutanJO
            Frm.vIdentifierJO = vIdentifierJO
            Frm.vBolehEdit = True
            Frm.vReorderHide = True
            Frm.vLgsgTutup = True
            Frm.ShowDialog()
            Frm.Dispose()

            'isi WO number di tbl item WO
            Me.IsiWONO_N_Urutan(vIdentifierJO, Me.TblProsesOrder, Me.TblItemWO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub IsiTblItemBOM(ByVal vPrimaryProduct As Double, ByVal vIdentifierJO As String, ByVal vTblItemBOM As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomDataTable, _
    ByVal vPrimaryParent As Double, ByVal vIdentifierBOMParent As String, ByVal vPrimaryMO As Double)
        Dim DTA As New DseTrans_ManufacturingOrderTableAdapters.Tbl_RekursifTableAdapter
        Dim TBL As New DseTrans_ManufacturingOrder.Tbl_RekursifDataTable
        Dim row As DseTrans_ManufacturingOrder.Tbl_RekursifRow
        Dim rowItemBom As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_PrimaryProduct(TBL, vPrimaryProduct)

            For Each row In TBL
                If row.primary_product_raw <> -1 Then
                    rowItemBom = vTblItemBOM.Newtbltrans_mo_item_bomRow
                    With rowItemBom
                        .Primary = -1
                        .Primary_MO = vPrimaryMO
                        .Identifier_JO = vIdentifierJO
                        .Identifier_WO = ""
                        .Primary_Product = row.Primary
                        .Qty_Ref_Process = row.Qty_Ref
                        .Good_Percentage_Process = row.Good_Precentage
                        .Primary_Process = row.Primary_Proses
                        .Primary_Product_RM = row.primary_product_raw
                        .Qty_RM = row.qty_raw
                        .Good_Percentage_RM = row.good_perc_raw
                        .Decimal_RM = row.Decimal_Raw
                        .Remarks = row.Remarks_Bom
                        .Primary_Parent_Product = vPrimaryParent
                        .Identifier_BOM = Me.ObjMO_Item.GetNextIdentifierBOM(vTblItemBOM, vPrimaryMO)
                        .Identifier_BOM_Parent = vIdentifierBOMParent
                        ._rm__decimal_ = row._rm__decimal_
                        .rm_unit_primary = row.rm_unit_primary
                        .rm_code = row.rm_code
                        .rm_desc = row.rm_desc
                        .rm_unit_code = row.rm_unit_code
                        .tolerance_process = row.tolerance_process
                        .production_tolerance = row.production_tolerance
                    End With
                    vTblItemBOM.Rows.Add(rowItemBom)
                    IsiTblItemBOM(row.primary_product_raw, vIdentifierJO, vTblItemBOM, row.Primary, rowItemBom.Identifier_BOM, vPrimaryMO)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsiTblProsesOrder(ByVal vPrimaryProduct As Double, ByVal vTblProsesOrder As _
    DseTrans_ManufacturingOrder.tbltrans_mo_processorderDataTable, ByVal vIdentifierJO As String)
        Dim DtaProsesOrderDB As New DseTrans_ManufacturingOrderTableAdapters.Process_Order_BeforeTableAdapter
        Dim TblProsesOrderDB As New DseTrans_ManufacturingOrder.Process_Order_BeforeDataTable
        Dim RowProsesOrderDB As DseTrans_ManufacturingOrder.Process_Order_BeforeRow
        Dim RowProsesOrder As DseTrans_ManufacturingOrder.tbltrans_mo_processorderRow
        Dim rowItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woRow
        Dim Proses_Order As Double = -1
        Dim count As Double = 0

        Try
            vTblProsesOrder.DefaultView.RowFilter = "[identifier jo] = '" & vIdentifierJO & "'"
            While vTblProsesOrder.DefaultView.Count > 0
                RowProsesOrder = vTblProsesOrder.DefaultView(0).Row
                vTblProsesOrder.Rows.Remove(RowProsesOrder)
            End While

            'cari urutan proses order atas product yang bersangkutan yang sudah ada sebelumnya di dgv item jo
            Me.TblJO_Temp.DefaultView.RowFilter = "[Primary Product] = " & vPrimaryProduct.ToString
            Me.TblJO_Temp.DefaultView.Sort = "[Urutan] Desc"
            If Me.TblJO_Temp.DefaultView.Count > 0 Then
                Me.RowJO_Temp = Me.TblJO_Temp.DefaultView(0).Row
                Me.TblProsesOrder.DefaultView.RowFilter = "[Identifier JO] = '" & Me.RowJO_Temp.Identifier_JO & "'"
                If Me.TblProsesOrder.DefaultView.Count > 0 Then
                    For Each dv As Data.DataRowView In Me.TblProsesOrder.DefaultView
                        RowProsesOrder = dv.Row
                        Proses_Order = RowProsesOrder.primary_process
                        Me.TblItemWO.DefaultView.RowFilter = "[Identifier JO] = '" & vIdentifierJO.ToString & "' AND [Primary Process] = " & RowProsesOrder.primary_process
                        If Me.TblItemWO.DefaultView.Count > 0 Then
                            RowProsesOrder = Me.TblProsesOrder.Newtbltrans_mo_processorderRow
                            With RowProsesOrder
                                .primary = -1
                                .primary_mo = CDbl(Me.TxtPrimary.Text)
                                .identifier_jo = vIdentifierJO
                                .primary_process = Proses_Order
                                .primary_product = vPrimaryProduct
                                count += 1
                                .urutan = count
                            End With
                            Me.TblProsesOrder.Rows.Add(RowProsesOrder)
                        End If
                    Next
                End If
            Else
                ''jika di dgv item jo tidak ada referesi proses order untuk product yang bersangkutan maka coba cari referensi dari database
                'DtaProsesOrderDB.FillBy_PrimaryProduct(TblProsesOrderDB, vPrimaryProduct)
                Dim Identifier_JO As String = ""
                Identifier_JO = ObjMO.GetMaxIdentifierJO_ProsesOrder(vPrimaryProduct)
                If Identifier_JO <> "" Then
                    DtaProsesOrderDB.FillBy_IdentifierJO(TblProsesOrderDB, Identifier_JO)
                    If TblProsesOrderDB.Count > 0 Then
                        For Each RowProsesOrderDB In TblProsesOrderDB.Rows
                            Proses_Order = RowProsesOrderDB.primary_process
                            Me.TblItemWO.DefaultView.RowFilter = "[Identifier JO] = '" & vIdentifierJO.ToString & "' AND [Primary Process] = " & RowProsesOrderDB.primary_process
                            If Me.TblItemWO.DefaultView.Count > 0 Then
                                RowProsesOrder = Me.TblProsesOrder.Newtbltrans_mo_processorderRow
                                With RowProsesOrder
                                    .primary = -1
                                    .primary_mo = CDbl(Me.TxtPrimary.Text)
                                    .identifier_jo = vIdentifierJO
                                    .primary_process = Proses_Order
                                    .primary_product = vPrimaryProduct
                                    count += 1
                                    .urutan = count
                                End With
                                Me.TblProsesOrder.Rows.Add(RowProsesOrder)
                            End If
                        Next
                    End If

                    DtaProsesOrderDB.Dispose()
                    TblProsesOrderDB.Dispose()
                End If
            End If

            Me.TblItemWO.DefaultView.RowFilter = "[Identifier JO] = '" & vIdentifierJO & "'"
            Me.TblItemWO.DefaultView.Sort = "[Level] ASC"
            If Me.TblItemWO.DefaultView.Count > 0 Then
                For i As Integer = Me.TblItemWO.DefaultView.Count - 1 To 0 Step -1
                    rowItemWO = Me.TblItemWO.DefaultView(i).Row
                    TblProsesOrder.DefaultView.RowFilter = "[primary process] = " & rowItemWO.Primary_Process.ToString & " And [identifier jo] = '" & vIdentifierJO & "'"
                    If vTblProsesOrder.DefaultView.Count < 1 Then
                        RowProsesOrder = vTblProsesOrder.Newtbltrans_mo_processorderRow
                        With RowProsesOrder
                            .primary = -1
                            .primary_mo = CDbl(Me.TxtPrimary.Text)
                            .identifier_jo = vIdentifierJO
                            .primary_process = rowItemWO.Primary_Process
                            .primary_product = vPrimaryProduct
                            count += 1
                            .urutan = count
                        End With
                        vTblProsesOrder.Rows.Add(RowProsesOrder)
                    End If
                Next
            End If

            Me.TblItemBom.DefaultView.RowFilter = Nothing
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub IsiTblItemWORM(ByVal vLevel As Integer, ByVal vPrimaryProduct As Double, ByVal vQtyTarget As Decimal, _
    ByVal vTblItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable, _
    ByVal vTblItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmDataTable, _
    ByVal vTblItemBOM As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomDataTable, _
    ByVal vIdentifierJO As String, ByVal vIdentifierParent As String, _
    ByVal vIdentifierBOM_Parent As String, ByVal vTblItemWIP As DseTrans_ManufacturingOrder.tbltrans_mo_item_wipDataTable, ByVal vRmReqDate As Date, ByVal vSetRMReqDate As Boolean)
        Dim vTblItemBoM_ As New DseTrans_ManufacturingOrder.tbltrans_mo_item_bomDataTable
        Dim rowItemBOM_ As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomRow
        Dim rowItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woRow
        Dim rowItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmRow
        Dim dec_Product As Integer = 0

        Try
            vTblItemBoM_ = vTblItemBOM.Copy

            vTblItemBoM_.DefaultView.RowFilter = "[primary product] = " & vPrimaryProduct.ToString & " AND [Identifier BOM Parent] = '" & _
            vIdentifierBOM_Parent.ToString & "' AND [Identifier JO] = '" & vIdentifierJO & "'"
            If vTblItemBoM_.DefaultView.Count > 0 Then
                dec_Product = Me.ObjProduct.GetAngkaDecimal(vPrimaryProduct)
                rowItemBOM_ = vTblItemBoM_.DefaultView(0).Row
                vTblItemWO.DefaultView.RowFilter = "[Primary Product] = " & vPrimaryProduct.ToString & " AND [Identifier] = '" _
                & rowItemBOM_.Identifier_WO.ToString & "' AND [Identifier JO] = '" & vIdentifierJO & "'"
                If vTblItemWO.DefaultView.Count > 0 Then
                    rowItemWO = vTblItemWO.DefaultView(0).Row
                    With rowItemWO
                        .Primary_Process = rowItemBOM_.Primary_Process
                        .Qty_Target_Result = Me.ObjCD.RoundUp1(vQtyTarget, dec_Product)
                        .Qty_Target_Prod = Me.ObjCD.RoundUp1(vQtyTarget / (rowItemBOM_.Good_Percentage_Process / 100), dec_Product)
                        .Plan_Good_Perct = rowItemBOM_.Good_Percentage_Process
                        .Qty_Projected_Result = Me.ObjCD.RoundDown1((.Qty_Target_Prod + .Qty_Target_Prod_Add - .Qty_Actual_Prod) * (.Plan_Good_Perct / 100) + .Qty_Actual_Result, dec_Product)
                        If .Qty_Target_Result = 0 Then
                            .Actual_Good_Perct = 0
                        Else
                            .Actual_Good_Perct = (.Qty_Actual_Result * 100) / .Qty_Target_Result
                        End If
                        .Balance_Produce = Me.ObjCD.RoundUp1(.Qty_Target_Prod - .Qty_Actual_Prod, dec_Product)
                        .Balance_Result = Me.ObjCD.RoundUp1(.Qty_Target_Result - .Qty_Actual_Result, dec_Product)
                        .Qty_Balance_MP = .Balance_Result
                    End With
                Else
                    rowItemWO = vTblItemWO.Newtbltrans_mo_item_woRow
                    With rowItemWO
                        .Primary = -1
                        .Primary_MO = -1
                        .Identifier_JO = vIdentifierJO
                        .WO_Number = ""
                        .Urutan = 0
                        ._Date = Now
                        If Me.ChkScheduling.Checked = False Then
                            .Status = ClsTrans_ManufacturingOrder.EnumPilihStatusWO.Free
                        Else
                            .Status = ClsTrans_ManufacturingOrder.EnumPilihStatusWO.Planned
                        End If
                        .Primary_Process = rowItemBOM_.Primary_Process
                        .Remarks = ""
                        .Identifier = ObjMO_Item.GetNextIdentifierWO(vTblItemWO, Me.TxtPrimary.Text, Me)
                        .Primary_Work_Unit = -1
                        .Plan_Start_Time = Now
                        .Plan_End_Time = Now
                        .Plan_Setup_Time = 0
                        .Plan_Work_Hour = 0
                        .Plan_Delay_Time = 0
                        .Plan_Total_Time = 0
                        .Plan_Prod_Rate = 0
                        .Actual_Start_Time = Now
                        .Actual_End_Time = Now
                        .Actual_Setup_Time = 0
                        .Actual_Work_Hour = 0
                        .Actual_Total_Time = 0
                        .Actual_Down_Time = 0
                        .Actual_Prod_Rate = 0
                        .Actual_Delay_Time = 0
                        .Qty_WIP = 0
                        .Qty_Target_Result = Me.ObjCD.RoundUp1((vQtyTarget - .Qty_WIP), dec_Product)
                        .Qty_Target_Prod = Me.ObjCD.RoundUp1((vQtyTarget - .Qty_WIP) / (rowItemBOM_.Good_Percentage_Process / 100), dec_Product)
                        .Qty_Target_Prod_Add = 0
                        .Qty_Actual_Prod = 0
                        .Qty_Actual_Result = 0
                        .Qty_Balance_MP = 0
                        .Plan_Good_Perct = rowItemBOM_.Good_Percentage_Process
                        .Qty_Projected_Result = Me.ObjCD.RoundDown1((.Qty_Target_Prod + .Qty_Target_Prod_Add - .Qty_Actual_Prod) * (.Plan_Good_Perct / 100) + .Qty_Actual_Result, dec_Product)
                        .Actual_Good_Perct = (.Qty_Actual_Result * 100) / .Qty_Target_Result
                        .Created_By = Me.UserCode
                        .Created_Date = Now
                        .Modified_By = Me.UserCode
                        .Modified_Date = Now
                        .Timestamp = Now
                        .Identifier_WO_Parent = vIdentifierParent
                        .Level = vLevel + 1
                        .Primary_Product = vPrimaryProduct
                        .Balance_Produce = Me.ObjCD.RoundUp1(.Qty_Target_Prod - .Qty_Actual_Prod, dec_Product)
                        .Balance_Result = Me.ObjCD.RoundUp1(.Qty_Target_Result - .Qty_Actual_Result, dec_Product)
                        .Qty_Balance_MP = .Balance_Result

                        .Allocated_WU = -1
                        .Product_Code = Me.ObjProduct.GetProductCode(vPrimaryProduct)
                        .Description = Me.ObjProduct.GetDescription(vPrimaryProduct)

                    End With
                    vTblItemWO.Rows.Add(rowItemWO)
                End If

                For i As Integer = 0 To vTblItemBoM_.DefaultView.Count - 1
                    rowItemBOM_ = vTblItemBoM_.DefaultView(i).Row
                    vTblItemRM.DefaultView.RowFilter = "[Primary Product Code RM] = " & rowItemBOM_.Primary_Product_RM.ToString & _
                    " AND [Identifier WO] = '" & rowItemWO.Identifier.ToString & "' AND [Identifier JO] = '" & vIdentifierJO & "'"
                    If vTblItemRM.DefaultView.Count > 0 Then
                        rowItemRM = vTblItemRM.DefaultView(0).Row
                        With rowItemRM
                            .Required_Qty = Me.ObjCD.RoundUp1((rowItemWO.Qty_Target_Prod / rowItemBOM_.Qty_Ref_Process * rowItemBOM_.Qty_RM) / (rowItemBOM_.Good_Percentage_RM / 100), rowItemBOM_.Decimal_RM)
                            .Allocated_Qty = Me.ObjCD.RoundUp1(.Required_Qty, rowItemBOM_.Decimal_RM)
                            .Balance_Qty = Me.ObjCD.RoundUp1(.Allocated_Qty + .Additional_Qty - .Used_Qty - .WIP_Qty, rowItemBOM_.Decimal_RM)
                            .Target_Qty = Me.ObjCD.RoundUp1((rowItemWO.Qty_Target_Prod / rowItemBOM_.Qty_Ref_Process) * rowItemBOM_.Qty_RM, rowItemBOM_.Decimal_RM)
                            .Primary_StandBy_Location = -1
                            .Status_RM = 0
                        End With
                    Else
                        rowItemRM = vTblItemRM.Newtbltrans_mo_item_rmRow
                        With rowItemRM
                            .Primary = -1
                            .Primary_MO = rowItemWO.Primary_MO
                            .Identifier_JO = rowItemWO.Identifier_JO
                            .Identifier_WO = rowItemWO.Identifier
                            .Primary_Product_Code_RM = rowItemBOM_.Primary_Product_RM
                            .Required_Qty = Me.ObjCD.RoundUp1((rowItemWO.Qty_Target_Prod / rowItemBOM_.Qty_Ref_Process * rowItemBOM_.Qty_RM) / (rowItemBOM_.Good_Percentage_RM / 100), rowItemBOM_.Decimal_RM)
                            .Allocated_Qty = Me.ObjCD.RoundUp1(.Required_Qty, rowItemBOM_.Decimal_RM)
                            .Additional_Qty = 0
                            .Used_Qty = 0
                            .WIP_Qty = 0
                            .Replace_Qty = 0
                            .Balance_Qty = Me.ObjCD.RoundUp1(.Allocated_Qty + .Additional_Qty - .Used_Qty - .WIP_Qty - .Replace_Qty, rowItemBOM_.Decimal_RM)
                            .Primary_StandBy_Location = -1
                            .Primary_PickUp_Location = -1
                            .Remarks = ""
                            .Target_Qty = Me.ObjCD.RoundUp1((rowItemWO.Qty_Target_Prod / rowItemBOM_.Qty_Ref_Process) * rowItemBOM_.Qty_RM, rowItemBOM_.Decimal_RM)
                            .GP_Target = rowItemBOM_.Good_Percentage_RM
                            .GP_Actual = 0
                            .Status_RM = 0
                            .use_priority = 1

                            .is_replacement = 0
                            .primary_code_rm_replaced = -1
                            .ratio_replacement = 1
                            .ratio_replaced = 1
                            .qty_replaced = 0
                            .primary_unit_replaced = -1
                            If vSetRMReqDate = True Then
                                .required_date = vRmReqDate
                            Else
                                Try
                                    If IsNothing(.required_date) = True Then
                                        .required_date = vRmReqDate
                                    Else
                                        .required_date = .required_date
                                    End If
                                Catch ex As Exception
                                    'MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
                                    .required_date = Today
                                End Try
                            End If

                            'update identifier wo di tbl item bom
                            Me.RowItemBom = vTblItemBOM.Rows(vTblItemBoM_.Rows.IndexOf(rowItemBOM_))
                            Me.RowItemBom.Identifier_WO = rowItemWO.Identifier

                            'set apakah wip
                            If Me.Is_WIP(Me.RowItemBom, vTblItemBoM_) = True Then
                                .is_wip = 1
                            Else
                                .is_wip = 0
                            End If
                            ._decimal = rowItemBOM_._rm__decimal_
                            .unit_primary = rowItemBOM_.rm_unit_primary
                            .rm_code = rowItemBOM_.rm_code
                            .rm_desc = rowItemBOM_.rm_desc
                            .unit_code = rowItemBOM_.rm_unit_code

                        End With
                        vTblItemRM.Rows.Add(rowItemRM)
                    End If
                    IsiTblItemWORM(rowItemWO.Level, rowItemRM.Primary_Product_Code_RM, rowItemRM.Balance_Qty, vTblItemWO, vTblItemRM, _
                    vTblItemBOM, vIdentifierJO, rowItemWO.Identifier, rowItemBOM_.Identifier_BOM, Me.TblItemWIP, vRmReqDate, vSetRMReqDate)
                Next

            End If
        Catch ex As Exception
            MsgBox(ex.Message + " - " + vTblItemBoM_.DefaultView.RowFilter & "," & vLevel & "," & vPrimaryProduct & "," & vQtyTarget)
        Finally
            vTblItemBoM_.Dispose()
        End Try
    End Sub

    Private Function Is_WIP(ByVal RowItemBomRef As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomRow, _
    ByVal vTblItemBoM_ As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomDataTable) As Boolean
        Dim RowItemBom As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomRow

        Try
            'cari apakah dia memiliki anak
            For Each RowItemBom In vTblItemBoM_.Rows
                If RowItemBomRef.Identifier_BOM = RowItemBom.Identifier_BOM_Parent Then
                    Return True
                End If
            Next

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Sub IsiWONO_N_Urutan(ByVal vIdentifierJO As String, ByVal vTblProsesOrder As DseTrans_ManufacturingOrder.tbltrans_mo_processorderDataTable, _
    ByVal vTblItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable)
        Dim rowProsesOrder As DseTrans_ManufacturingOrder.tbltrans_mo_processorderRow
        Dim rowItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woRow
        'Dim rowItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmRow

        Try
            vTblProsesOrder.DefaultView.RowFilter = "[Identifier JO] = '" & vIdentifierJO & "'"
            If vTblProsesOrder.DefaultView.Count > 0 Then
                For Each dv As Data.DataRowView In vTblProsesOrder.DefaultView
                    rowProsesOrder = dv.Row
                    'isi wo number dan urutan di tbl item wo
                    vTblItemWO.DefaultView.RowFilter = "[Primary Process] = " & rowProsesOrder.primary_process.ToString & _
                    " AND [Identifier JO] = '" & vIdentifierJO & "'"
                    vTblItemWO.DefaultView.Sort = "[Level] DESC"
                    If vTblItemWO.DefaultView.Count > 0 Then
                        For Each dv2 As Data.DataRowView In vTblItemWO.DefaultView
                            rowItemWO = dv2.Row
                            rowItemWO.Primary_Product = rowItemWO.Primary_Product
                            If rowItemWO.WO_Number = "" Then
                                rowItemWO.WO_Number = Me.ObjMO.GetNoBaru1(Me.DtpTglTransaksi.Value, ClsTrans_ManufacturingOrder.PilihFormatNumber.WO, _
                                Me.getMaxNoUrutWO(vTblItemWO, vTblProsesOrder.DefaultView.RowFilter), Me, rowItemWO.Urutan, vTblItemWO)
                            End If
                        Next
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Public Sub IsiDgvJO_FromSF(ByVal vPrimaryProduct As Double, ByVal vQtyTarget As Decimal, ByVal vTargetFinishDate As Date)
    '    Dim rowD As DataGridViewRow

    '    Try
    '        Me.DgvItem.AllowUserToAddRows = False

    '        Me.DgvItem.Rows.Add()
    '        Me.DgvItem.AllowUserToAddRows = False
    '        rowD = Me.DgvItem.Rows(Me.DgvItem.Rows.Count - 1)
    '        If String.IsNullOrEmpty(rowD.Cells(Me.ColItem_Urutan.Index).Value) = True Then
    '            rowD.Cells(Me.ColItem_JONumber.Index).Value = Me.ObjMO.GetNoBaru1(Me.DtpTglTransaksi.Value, _
    '            ClsTrans_ManufacturingOrder.PilihFormatNumber.JO, Me.getMaxNoUrutJO, Me, rowD.Index)
    '        End If
    '        rowD.Cells(Me.ColItem_PrimaryProduct.Index).Value = vPrimaryProduct
    '        rowD.Cells(Me.ColItem_CodeNo.Index).Value = Me.ObjProduct.GetProductCode(vPrimaryProduct)
    '        rowD.Cells(Me.ColItem_Desc.Index).Value = Me.ObjProduct.GetDescription(vPrimaryProduct)
    '        rowD.Cells(Me.ColItem_Decimal.Index).Value = Me.ObjProduct.GetAngkaDecimal(vPrimaryProduct)
    '        rowD.Cells(Me.ColItem_Unit.Index).Value = Me.ObjProduct.GetUnitCode(vPrimaryProduct)
    '        rowD.Cells(Me.ColItem_QtyTarget.Index).Value = Me.ObjCD.DigitGrouping(vQtyTarget, rowD.Cells(Me.ColItem_Decimal.Index).Value)
    '        rowD.Cells(Me.ColItem_TargetFinishDate.Index).Value = Format(vTargetFinishDate, "dd MMM yyyy")
    '        rowD.Cells(Me.ColItem_TargetFinishDateR1.Index).Value = "-"
    '        rowD.Cells(Me.ColItem_TargetFinishDateR2.Index).Value = "-"
    '        rowD.Cells(Me.ColItem_ActualFinishDate.Index).Value = "-"

    '        'fungsi untuk update table item BOM, WO, RM, Proses Order
    '        Me.IsiAllTblItem(vPrimaryProduct, rowD.Cells(Me.ColItem_QtyTarget.Index).Value, _
    '        rowD.Cells(Me.ColItem_Identifier.Index).Value, rowD.Cells(Me.ColItem_Urutan.Index).Value, False)

    '        'isi qty produce
    '        rowD.Cells(Me.ColItem_QtyProduce.Index).Value = Me.ObjCD.DigitGrouping(ObjMO.GetQtyProduce_InTbl(vPrimaryProduct, _
    '        vQtyTarget, rowD.Cells(Me.ColItem_Identifier.Index).Value, Me.TblItemBom), rowD.Cells(Me.ColItem_Decimal.Index).Value, _
    '        Core_Module.ClsCekData.EnumTipeRounding.Up)

    '        Me.JO_TempADD(vPrimaryProduct, rowD.Cells(Me.ColItem_Identifier.Index).Value)

    '        Me.DgvItem.AllowUserToAddRows = True
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Private Sub EndEditProductCode(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim row As DataGridViewRow = Me.DgvItem.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim PrimaryProduct As Double = -1
        Dim RowIdx As Integer = e.RowIndex
        Dim TypeAccProd As Integer = 0
        Dim CountItemBOM As Integer = 0
        Dim jk As New FrmTrans_ManufacturingOrder_RMReqDate

        Try
            With row
                If IsNothing(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColItem_CodeNo.Index Then
                        Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper
                        If Me.ObjProduct.ApakahSudahAda(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            If Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer Then
                                Exit Sub
                            End If
                            PrimaryProduct = Me.ObjProduct.GetPrimary(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                            TypeAccProd = ObjMO.GetTypeAccProduct(PrimaryProduct)
                            CountItemBOM = ObjMO.GetJmlItemBOM(PrimaryProduct)
                            If TypeAccProd = 0 Then
                                If CountItemBOM > 0 Then
                                    If String.IsNullOrEmpty(Me.DgvItem.Item(Me.ColItem_Urutan.Index, e.RowIndex).Value) = True Then
                                        Me.DgvItem.Item(Me.ColItem_JONumber.Index, e.RowIndex).Value = Me.ObjMO.GetNoBaru1(Me.DtpTglTransaksi.Value, _
                                        ClsTrans_ManufacturingOrder.PilihFormatNumber.JO, Me.getMaxNoUrutJO, Me, e.RowIndex)
                                    End If
                                    Me.DgvItem.Item(Me.ColItem_PrimaryProduct.Index, e.RowIndex).Value = PrimaryProduct
                                    Me.DgvItem.Item(Me.ColItem_CodeNo.Index, e.RowIndex).Value = Me.ObjProduct.GetProductCode(PrimaryProduct)
                                    Me.DgvItem.Item(Me.ColItem_Desc.Index, e.RowIndex).Value = Me.ObjProduct.GetDescription(PrimaryProduct)
                                    Me.DgvItem.Item(Me.ColItem_Decimal.Index, e.RowIndex).Value = Me.ObjProduct.GetAngkaDecimal(PrimaryProduct)
                                    Me.DgvItem.Item(Me.ColItem_Unit.Index, e.RowIndex).Value = Me.ObjProduct.GetUnitCode(PrimaryProduct)
                                    Me.DgvItem.Item(Me.ColItem_QtyTarget.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.ObjMO.GetQtyRef(PrimaryProduct), _
                                    Me.DgvItem.Item(Me.ColItem_Decimal.Index, e.RowIndex).Value)
                                    Me.DgvItem.Item(Me.ColItem_TargetFinishDate.Index, e.RowIndex).Value = Format(Me.DtpTargetFinishDate.Value, "dd MMM yyyy")
                                    Me.DgvItem.Item(Me.ColItem_TargetFinishDateR1.Index, e.RowIndex).Value = "-"
                                    Me.DgvItem.Item(Me.ColItem_TargetFinishDateR2.Index, e.RowIndex).Value = "-"
                                    Me.DgvItem.Item(Me.ColItem_ActualFinishDate.Index, e.RowIndex).Value = "-"

                                    'fungsi untuk update table item BOM, WO, RM, Proses Order
                                    jk.btnCancel.Enabled = False
                                    jk.btnCancel.Visible = False
                                    jk.Size = New Point(jk.Size.Width, jk.Size.Height - jk.btnCancel.Height)
                                    jk.ShowDialog()
                                    Me.KosongkanTabelDetail(Me.DgvItem.Item(Me.ColItem_Identifier.Index, e.RowIndex).Value)
                                    Me.IsiAllTblItem(PrimaryProduct, Me.DgvItem.Item(Me.ColItem_QtyTarget.Index, e.RowIndex).Value, _
                                    Me.DgvItem.Item(Me.ColItem_Identifier.Index, e.RowIndex).Value, Me.DgvItem.Item(Me.ColItem_Urutan.Index, e.RowIndex).Value, _
                                    True, jk.dtpRequired.Value, True)

                                    'isi qty produce
                                    Me.DgvItem.Item(Me.ColItem_QtyProduce.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(ObjMO.GetQtyProduce_InTbl(Me.DgvItem.Item(Me.ColItem_PrimaryProduct.Index, e.RowIndex).Value, _
                                    Me.DgvItem.Item(Me.ColItem_QtyTarget.Index, e.RowIndex).Value, Me.DgvItem.Item(Me.ColItem_Identifier.Index, e.RowIndex).Value, Me.TblItemBom), _
                                    Me.DgvItem.Item(Me.ColItem_Decimal.Index, e.RowIndex).Value, Core_Module.ClsCekData.EnumTipeRounding.Up)
                                    Me.JO_TempADD(Me.DgvItem.Item(Me.ColItem_PrimaryProduct.Index, e.RowIndex).Value, Me.DgvItem.Item(Me.ColItem_Identifier.Index, e.RowIndex).Value)
                                    Me.DgvItem.Item(Me.ColItem_DeliveryDate.Index, e.RowIndex).Value = Format(Me.DtpTargetFinishDate.Value, "dd MMM yyyy")
                                    Dim rowDelivery As DseTrans_ManufacturingOrder.tbltrans_mo_itemdeliveryRow = Me.TBLItemDelivery.Newtbltrans_mo_itemdeliveryRow
                                    With rowDelivery
                                        .primary = -1
                                        .primary_mo = Me.TxtPrimary.Text
                                        .identifier_jo = Me.DgvItem.Item(Me.ColItem_Identifier.Index, e.RowIndex).Value
                                        .delivery_qty = Me.DgvItem.Item(Me.ColItem_QtyTarget.Index, e.RowIndex).Value
                                        .delivery_date = Me.DtpTargetFinishDate.Value
                                        .remarks = String.Empty
                                    End With
                                    TBLItemDelivery.Rows.Add(rowDelivery)

                                    'clear semua dgv untuk triger load ulang dgv yang bersesuaian
                                    Me.TGV_WO.Nodes.Clear()
                                    Me.TGV_WO.Rows.Clear()
                                    Me.DGVRM.Rows.Clear()
                                    Me.DgvRMRekap.Rows.Clear()
                                    Me.dgvItemWO.Rows.Clear()
                                    Me.DGVWIP.Rows.Clear()
                                Else
                                    MsgBox("You can only enter a product which have BOM", MsgBoxStyle.Information, "Information")
                                    GoTo HapusBaris
                                End If
                            Else
                                MsgBox("You can only enter a product type goods", MsgBoxStyle.Information, "Information")
                                GoTo HapusBaris
                            End If
                        Else
                            Dim ij As New Core_Module.frmCariMasterProduct
                            ij.TipeProduct = Core_Module.clsMasterProduct.EnumTipeProduct.Goods
                            ij.txtCari.Text = Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                            ij.BackColorAll = Me.BackColorAll
                            ij.ShowDialog()
                            TypeAccProd = ObjMO.GetTypeAccProduct(ij.PrimaryProduct)
                            CountItemBOM = ObjMO.GetJmlItemBOM(ij.PrimaryProduct)
                            If ij.PrimaryProduct <> -1 Then
                                If TypeAccProd = 0 Then
                                    If CountItemBOM > 0 Then
                                        Me.DgvItem.Item(Me.ColItem_PrimaryProduct.Index, e.RowIndex).Value = ij.PrimaryProduct
                                        Me.DgvItem.Item(Me.ColItem_CodeNo.Index, e.RowIndex).Value = Me.ObjProduct.GetProductCode(ij.PrimaryProduct)
                                        If Me.DgvItem.Item(Me.ColItem_CodeNo.Index, e.RowIndex).Value = Me.Buffer Then
                                            Exit Sub
                                        End If
                                        Me.DgvItem.Item(Me.ColItem_Desc.Index, e.RowIndex).Value = Me.ObjProduct.GetDescription(ij.PrimaryProduct)
                                        Me.DgvItem.Item(Me.ColItem_Unit.Index, e.RowIndex).Value = Me.ObjProduct.GetUnitCode(ij.PrimaryProduct)
                                        Me.DgvItem.Item(Me.ColItem_Decimal.Index, e.RowIndex).Value = Me.ObjProduct.GetAngkaDecimal(ij.PrimaryProduct)
                                        Me.DgvItem.Item(Me.ColItem_QtyTarget.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.ObjMO.GetQtyRef(ij.PrimaryProduct), Me.DgvItem.Item(Me.ColItem_Decimal.Index, e.RowIndex).Value)
                                        If String.IsNullOrEmpty(Me.DgvItem.Item(Me.ColItem_Urutan.Index, e.RowIndex).Value) = True Then
                                            Me.DgvItem.Item(Me.ColItem_JONumber.Index, e.RowIndex).Value = Me.ObjMO.GetNoBaru1(Me.DtpTglTransaksi.Value, _
                                            ClsTrans_ManufacturingOrder.PilihFormatNumber.JO, Me.getMaxNoUrutJO, Me, e.RowIndex)
                                        End If
                                        Me.DgvItem.Item(Me.ColItem_TargetFinishDate.Index, e.RowIndex).Value = Format(Me.DtpTargetFinishDate.Value, "dd MMM yyyy")
                                        Me.DgvItem.Item(Me.ColItem_DeliveryDate.Index, e.RowIndex).Value = Format(Me.DtpTargetFinishDate.Value, "dd MMM yyyy")
                                        Me.DgvItem.Item(Me.ColItem_TargetFinishDateR1.Index, e.RowIndex).Value = "-"
                                        Me.DgvItem.Item(Me.ColItem_TargetFinishDateR2.Index, e.RowIndex).Value = "-"
                                        Me.DgvItem.Item(Me.ColItem_ActualFinishDate.Index, e.RowIndex).Value = "-"

                                        'fungsi untuk update table item BOM, WO, RM, Proses Order
                                        jk.btnCancel.Enabled = False
                                        jk.btnCancel.Visible = False
                                        jk.Size = New Point(jk.Size.Width, jk.Size.Height - jk.btnCancel.Height)
                                        jk.ShowDialog()
                                        Me.KosongkanTabelDetail(Me.DgvItem.Item(Me.ColItem_Identifier.Index, e.RowIndex).Value)
                                        Me.IsiAllTblItem(ij.PrimaryProduct, Me.DgvItem.Item(Me.ColItem_QtyTarget.Index, e.RowIndex).Value, _
                                        Me.DgvItem.Item(Me.ColItem_Identifier.Index, e.RowIndex).Value, Me.DgvItem.Item(Me.ColItem_Urutan.Index, e.RowIndex).Value, True, _
                                        jk.dtpRequired.Value.Date, True)

                                        'isi qty produce
                                        'Me.DgvItem.Item(Me.ColItem_QtyProduce.Index, e.RowIndex).Value = Me.ObjCD.RoundUp1(ObjMO.GetQtyProduce_InTbl(Me.DgvItem.Item(Me.ColItem_PrimaryProduct.Index, e.RowIndex).Value, _
                                        'Me.DgvItem.Item(Me.ColItem_QtyTarget.Index, e.RowIndex).Value, Me.DgvItem.Item(Me.ColItem_Identifier.Index, e.RowIndex).Value, Me.TblItemBom), _
                                        'Me.DgvItem.Item(Me.ColItem_Decimal.Index, e.RowIndex).Value)
                                        Me.DgvItem.Item(Me.ColItem_QtyProduce.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(ObjMO.GetQtyProduce_InTbl(Me.DgvItem.Item(Me.ColItem_PrimaryProduct.Index, e.RowIndex).Value, _
                                            Me.DgvItem.Item(Me.ColItem_QtyTarget.Index, e.RowIndex).Value, Me.DgvItem.Item(Me.ColItem_Identifier.Index, e.RowIndex).Value, Me.TblItemBom), _
                                            Me.DgvItem.Item(Me.ColItem_Decimal.Index, e.RowIndex).Value, Core_Module.ClsCekData.EnumTipeRounding.Up)

                                        Me.JO_TempADD(Me.DgvItem.Item(Me.ColItem_PrimaryProduct.Index, e.RowIndex).Value, Me.DgvItem.Item(Me.ColItem_Identifier.Index, e.RowIndex).Value)
                                        Me.DgvItem.Item(Me.ColItem_DeliveryDate.Index, e.RowIndex).Value = Format(Me.DtpTargetFinishDate.Value, "dd MMM yyyy")
                                        Dim rowDelivery As DseTrans_ManufacturingOrder.tbltrans_mo_itemdeliveryRow = Me.TBLItemDelivery.Newtbltrans_mo_itemdeliveryRow
                                        With rowDelivery
                                            .primary = -1
                                            .primary_mo = Me.TxtPrimary.Text
                                            .identifier_jo = Me.DgvItem.Item(Me.ColItem_Identifier.Index, e.RowIndex).Value
                                            .delivery_qty = Me.DgvItem.Item(Me.ColItem_QtyTarget.Index, e.RowIndex).Value
                                            .delivery_date = Me.DtpTargetFinishDate.Value
                                            .remarks = String.Empty
                                        End With
                                        TBLItemDelivery.Rows.Add(rowDelivery)

                                        'clear semua dgv untuk triger load ulang dgv yang bersesuaian
                                        Me.TGV_WO.Nodes.Clear()
                                        Me.TGV_WO.Rows.Clear()
                                        Me.DGVRM.Rows.Clear()
                                        Me.DgvRMRekap.Rows.Clear()
                                        Me.dgvItemWO.Rows.Clear()
                                        Me.DGVWIP.Rows.Clear()
                                    Else
                                        MsgBox("You can only enter a product which have BOM", MsgBoxStyle.Information, "Information")
                                        GoTo HapusBaris
                                    End If
                                Else
                                    MsgBox("You can only enter a product type goods", MsgBoxStyle.Information, "Information")
                                    GoTo HapusBaris
                                End If
                            ElseIf ij.PrimaryProduct = -1 Then
HapusBaris:
                                If Me.DgvItem.Rows.Count - 1 <> e.RowIndex Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BackgroundWorker2.RunWorkerAsync()
                                    Else
                                        Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.DgvItem)
                                    Exit Sub
                                End If
                                If Me.DgvItem.Item(Me.ColItem_CodeNo.Index, e.RowIndex).Value <> "" Then
                                    If String.IsNullOrEmpty(Me.Buffer) = True Then
                                        Me.EraseRowIndex = e.RowIndex
                                        Me.BackgroundWorker2.RunWorkerAsync()
                                    Else
                                        Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    End If
                                    Me.ObjCD.SetNumber(Me.DgvItem)
                                    Exit Sub
                                End If
                            End If
                            ij = Nothing
                        End If
                        Me.ObjCD.SetNumber(Me.DgvItem)
                    End If
                Else
                    Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub JO_TempADD(ByVal vPrimaryProduct_ As Double, ByVal vIdentifierJO_ As String)

        Try
            Me.RowJO_Temp = Me.TblJO_Temp.NewJO_TempRow
            With Me.RowJO_Temp
                .Primary_Product = vPrimaryProduct_
                .Identifier_JO = vIdentifierJO_
            End With
            Me.TblJO_Temp.Rows.Add(Me.RowJO_Temp)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditQtyTarget(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim row As DataGridViewRow = Me.DgvItem.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim PrimaryProduct As Double = -1
        Dim RowIdx As Integer = e.RowIndex
        'Dim rowItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woRow
        'Dim rowItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmRow

        Try
            With row
                If IsNothing(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColItem_QtyTarget.Index Then
                        If IsNumeric(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            If IsNumeric(Me.Buffer) = True Then
                                Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.Buffer, Me.DgvItem.Item(Me.ColItem_Decimal.Index, e.RowIndex).Value)
                            Else
                                Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = 0
                            End If
                        Else
                            If Me.ObjCD.CekBilanganPositif(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                                MsgBox("Please enter positive number for qty produce.", MsgBoxStyle.Information, _
                                "Information")
                                Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.Buffer, Me.DgvItem.Item(Me.ColItem_Decimal.Index, e.RowIndex).Value)
                                Exit Sub
                            Else
                                Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value, Me.DgvItem.Item(Me.ColItem_Decimal.Index, e.RowIndex).Value)
                            End If
                        End If

                        'saat ubah qty target di item JO maka qty produce nya akan ikut berubah
                        'Me.DgvItem.Item(Me.ColItem_QtyProduce.Index, e.RowIndex).Value = Me.ObjCD.RoundUp1(ObjMO.GetQtyProduce_InTbl(Me.DgvItem.Item(Me.ColItem_PrimaryProduct.Index, e.RowIndex).Value, _
                        'Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value, Me.DgvItem.Item(Me.ColItem_Identifier.Index, e.RowIndex).Value, Me.TblItemBom), _
                        'Me.DgvItem.Item(Me.ColItem_Decimal.Index, e.RowIndex).Value)
                        Me.DgvItem.Item(Me.ColItem_QtyProduce.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(ObjMO.GetQtyProduce_InTbl(Me.DgvItem.Item(Me.ColItem_PrimaryProduct.Index, e.RowIndex).Value, _
                                            Me.DgvItem.Item(Me.ColItem_QtyTarget.Index, e.RowIndex).Value, Me.DgvItem.Item(Me.ColItem_Identifier.Index, e.RowIndex).Value, Me.TblItemBom), _
                                            Me.DgvItem.Item(Me.ColItem_Decimal.Index, e.RowIndex).Value, Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)

                        Dim JmlKurang As Decimal = 0
                        If CDec(Me.Buffer) > CDec(Me.DgvItem.Item(Me.ColItem_QtyTarget.Index, e.RowIndex).Value) Then
                            'untuk item delivery, jika qty turun, maka hapuskan mulai dari baris paling bawah
                            JmlKurang = CDec(Me.Buffer) - CDec(Me.DgvItem.Item(Me.ColItem_QtyTarget.Index, e.RowIndex).Value)
                            Me.TBLItemDelivery.DefaultView.RowFilter = "[identifier jo] = '" + Me.DgvItem.Item(Me.ColItem_Identifier.Index, e.RowIndex).Value + "'"
                            Me.TBLItemDelivery.DefaultView.Sort = "[delivery date] DESC"
                            While JmlKurang > 0
                                If Me.TBLItemDelivery.DefaultView.Count > 0 Then
                                    Me.RowItemDelivery = Me.TBLItemDelivery.DefaultView(0).Row
                                    If Me.RowItemDelivery.delivery_qty > JmlKurang Then
                                        Me.RowItemDelivery.delivery_qty -= JmlKurang
                                        JmlKurang = 0
                                    Else
                                        JmlKurang -= Me.RowItemDelivery.delivery_qty
                                        Me.RowItemDelivery.delivery_qty = 0
                                    End If
                                    If Me.RowItemDelivery.delivery_qty = 0 Then
                                        Me.TBLItemDelivery.Rows.Remove(Me.RowItemDelivery)
                                    End If
                                End If
                            End While
                        ElseIf CDec(Me.Buffer) < CDec(Me.DgvItem.Item(Me.ColItem_QtyTarget.Index, e.RowIndex).Value) Then
                            'untuk item delivery, jika qty naik, maka tambahkan sisanya pada baris paling bawah
                            JmlKurang = CDec(Me.DgvItem.Item(Me.ColItem_QtyTarget.Index, e.RowIndex).Value) - CDec(Me.Buffer)
                            Me.TBLItemDelivery.DefaultView.RowFilter = "[identifier jo] = '" + Me.DgvItem.Item(Me.ColItem_Identifier.Index, e.RowIndex).Value + "'"
                            Me.TBLItemDelivery.DefaultView.Sort = "[delivery date] DESC"
                            If Me.TBLItemDelivery.DefaultView.Count > 0 Then
                                Me.RowItemDelivery = Me.TBLItemDelivery.DefaultView(0).Row
                                Me.RowItemDelivery.delivery_qty += JmlKurang
                            End If
                        End If

                        Me.IsiTblItemWORM(-1, Me.DgvItem.Item(Me.ColItem_PrimaryProduct.Index, e.RowIndex).Value, _
                        Me.DgvItem.Item(Me.ColItem_QtyTarget.Index, e.RowIndex).Value, _
                        Me.TblItemWO, Me.TblItemRM, Me.TblItemBom, Me.DgvItem.Item(Me.ColItem_Identifier.Index, e.RowIndex).Value, "-", "-", Me.TblItemWIP, #1/1/1950#, False)
                    End If
                Else
                    Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditFinishDate(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim row As DataGridViewRow = Me.DgvItem.Rows(e.RowIndex)
        Dim index As Integer = -1
        Dim PrimaryProduct As Double = -1
        Dim RowIdx As Integer = e.RowIndex

        Try
            With row
                If IsNothing(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If e.ColumnIndex = Me.ColItem_TargetFinishDate.Index Then
                        ''set region ke indonesia
                        'My.Application.ChangeCulture("id-ID")
                        'Thread.CurrentThread.CurrentCulture = New CultureInfo("id-ID", True)
                        If IsDate(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            If Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value <= Me.DtpTglTransaksi.Value.Date Then
                                MsgBox("Target finish date must be larger than transaction date", MsgBoxStyle.Information, "Information")
                                Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                Exit Sub
                            End If
                        End If
                        If IsDate(Me.DgvItem.Item(Me.ColItem_TargetFinishDateR1.Index, e.RowIndex).Value) = True Then
                            If IsDate(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                                If CDate(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) >= CDate(Me.DgvItem.Item(Me.ColItem_TargetFinishDateR1.Index, e.RowIndex).Value) Then
                                    MsgBox("Target finish date must be smaller than date rev 1", MsgBoxStyle.Information, "Information")
                                    Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Exit Sub
                                End If
                            End If
                        End If
                        If IsDate(Me.DgvItem.Item(Me.ColItem_TargetFinishDateR2.Index, e.RowIndex).Value) = True Then
                            If IsDate(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                                If CDate(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) >= CDate(Me.DgvItem.Item(Me.ColItem_TargetFinishDateR2.Index, e.RowIndex).Value) Then
                                    MsgBox("Target finish date must be smaller than date rev 2", MsgBoxStyle.Information, "Information")
                                    Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Exit Sub
                                End If
                            End If
                        End If

                        If IsDate(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            MsgBox("Must input date format.", MsgBoxStyle.Information, "Information")
                            Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Else
                            Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Format(CDate(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value), "dd MMM yyyy")
                        End If
                        ''kembalikan ke region defaultnya
                        'My.Application.ChangeCulture("en-US")
                        'Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US", True)
                    ElseIf e.ColumnIndex = Me.ColItem_TargetFinishDateR1.Index Then
                        ''set region ke indonesia
                        'My.Application.ChangeCulture("id-ID")
                        'Thread.CurrentThread.CurrentCulture = New CultureInfo("id-ID", True)
                        If IsDate(Me.DgvItem.Item(Me.ColItem_TargetFinishDate.Index, e.RowIndex).Value) = True Then
                            If IsDate(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                                If CDate(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) <= CDate(Me.DgvItem.Item(Me.ColItem_TargetFinishDate.Index, e.RowIndex).Value) Then
                                    MsgBox("Target finish date must be smaller than Target finish date rev 1", MsgBoxStyle.Information, "Information")
                                    Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Exit Sub
                                End If
                            End If
                        End If
                        If IsDate(Me.DgvItem.Item(Me.ColItem_TargetFinishDateR2.Index, e.RowIndex).Value) = True Then
                            If IsDate(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                                If CDate(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) >= CDate(Me.DgvItem.Item(Me.ColItem_TargetFinishDateR2.Index, e.RowIndex).Value) Then
                                    MsgBox("Target finish date rev 1 must be smaller than date rev 2", MsgBoxStyle.Information, "Information")
                                    Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Exit Sub
                                End If
                            End If
                        End If
                        If IsDate(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            MsgBox("Must input date format.", MsgBoxStyle.Information, "Information")
                            Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Else
                            Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Format(CDate(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value), "dd MMM yyyy")
                        End If
                        ''kembalikan ke region defaultnya
                        'My.Application.ChangeCulture("en-US")
                        'Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US", True)
                    ElseIf e.ColumnIndex = Me.ColItem_TargetFinishDateR2.Index Then
                        ''set region ke indonesia
                        'My.Application.ChangeCulture("id-ID")
                        'Thread.CurrentThread.CurrentCulture = New CultureInfo("id-ID", True)
                        If IsDate(Me.DgvItem.Item(Me.ColItem_TargetFinishDate.Index, e.RowIndex).Value) = True Then
                            If IsDate(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                                If CDate(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) <= CDate(Me.DgvItem.Item(Me.ColItem_TargetFinishDate.Index, e.RowIndex).Value) Then
                                    MsgBox("Target finish date must be smaller than date rev 2", MsgBoxStyle.Information, "Information")
                                    Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Exit Sub
                                End If
                            End If
                        End If
                        If IsDate(Me.DgvItem.Item(Me.ColItem_TargetFinishDateR1.Index, e.RowIndex).Value) = True Then
                            If IsDate(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                                If CDate(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) <= CDate(Me.DgvItem.Item(Me.ColItem_TargetFinishDateR1.Index, e.RowIndex).Value) Then
                                    MsgBox("Target finish date rev 1 must be smaller than date rev 2", MsgBoxStyle.Information, "Information")
                                    Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                                    Exit Sub
                                End If
                            End If
                        End If
                        If IsDate(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            MsgBox("Must input date format.", MsgBoxStyle.Information, "Information")
                            Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Else
                            Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Format(CDate(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value), "dd MMM yyyy")
                        End If
                        ''kembalikan ke region defaultnya
                        'My.Application.ChangeCulture("en-US")
                        'Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US", True)
                    ElseIf e.ColumnIndex = Me.ColItem_ActualFinishDate.Index Then
                        ''set region ke indonesia
                        'My.Application.ChangeCulture("id-ID")
                        'Thread.CurrentThread.CurrentCulture = New CultureInfo("id-ID", True)
                        If IsDate(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            MsgBox("Must input date format.", MsgBoxStyle.Information, "Information")
                            Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.Buffer
                        Else
                            Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Format(CDate(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value), "dd MMM yyyy")
                        End If
                        ''kembalikan ke region defaultnya
                        'My.Application.ChangeCulture("en-US")
                        'Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US", True)
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Dim DtpTglTransaksiValue As Date
    Private Sub DtpTglTransaksi_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles DtpTglTransaksi.Enter
        Try
            Me.DtpTglTransaksiValue = Me.DtpTglTransaksi.Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DtpTglTransaksi_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles DtpTglTransaksi.Leave
        Try
            If Me.DtpTglTransaksi.Value = Me.DtpTglTransaksiValue Then
                Exit Sub
            End If
            Dim urutan As Integer = 0
            Dim rowWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woRow

            Try
                If Me.TxtRemarks.Enabled = True Then
                    If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.MO) = False Then
                        Exit Sub
                    End If
                    If Me.Baru = False Then
                        If Me.ObjMO.TanggalTransaksi.Month = Me.DtpTglTransaksi.Value.Month And Me.ObjMO.TanggalTransaksi.Year = Me.DtpTglTransaksi.Value.Year Then
                            Me.TxtNoTransaksi.Text = Me.ObjMO.NoTransaksi
                        Else
                            Me.TxtNoTransaksi.Text = Me.ObjMO.GetNoBaru1(Me.DtpTglTransaksi.Value, ClsTrans_ManufacturingOrder.PilihFormatNumber.MO, -1, Me)
                            'For i As Integer = 1 To Me.DgvItem.RowCount - 1
                            '    Me.DgvItem.Item(Me.ColItem_JONumber.Index, i - 1).Value = Me.ObjMO.GetNoBaru(Me.DtpTglTransaksi.Value, ClsTrans_ManufacturingOrder.PilihFormatNumber.JO, Me, i - 1, Me.DgvItem.Item(Me.ColItem_Urutan.Index, i - 1).Value, Me.DgvItem.Item(Me.ColItem_Urutan.Index, i - 1).Value)
                            'Next
                        End If
                    Else
                        Me.TxtNoTransaksi.Text = Me.ObjMO.GetNoBaru1(Me.DtpTglTransaksi.Value, ClsTrans_ManufacturingOrder.PilihFormatNumber.MO, -1, Me)
                        If Me.DgvItem.RowCount > 1 Then
                            For i As Integer = 0 To Me.DgvItem.RowCount - 1
                                If Me.DgvItem.Item(Me.ColItem_PrimaryProduct.Index, i).Value <> "-1" Then
                                    Me.DgvItem.Item(Me.ColItem_JONumber.Index, i).Value = Me.ObjMO.GetNoBaru1(Me.DtpTglTransaksi.Value, ClsTrans_ManufacturingOrder.PilihFormatNumber.JO, urutan, Me)
                                    urutan = Me.ObjNumbering.GetUrutan(Me.DgvItem.Item(Me.ColItem_JONumber.Index, i).Value, Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.JO, Me.DtpTglTransaksi.Value)
                                End If
                            Next
                        End If
                        If Me.TblItemWO.Rows.Count > 0 Then
                            Dim count As Integer = 0
                            Me.TblItemWO.DefaultView.RowFilter = Nothing
                            Me.TblItemWO.DefaultView.Sort = "[Urutan] ASC"
                            For Each dv As Data.DataRowView In Me.TblItemWO.DefaultView
                                rowWO = dv.Row
                                count += 1
                                rowWO.WO_Number = Me.ObjMO.GetNoBaru1(Me.DtpTglTransaksi.Value, ClsTrans_ManufacturingOrder.PilihFormatNumber.WO, count, Me, count, Me.TblItemWO)
                            Next

                            If Me.DgvItem.SelectedCells.Count > 0 Then
                                If String.IsNullOrEmpty(Me.DgvItem.Item(Me.ColItem_PrimaryProduct.Index, Me.DgvItem.SelectedCells.Item(0).RowIndex).Value) = False Then
                                    If TabControl1.SelectedTab.Name = Me.TabPage_WO.Name Then
                                        LoadDgvItemWO(Me.DgvItem.Item(Me.ColItem_Identifier.Index, Me.DgvItem.SelectedCells.Item(0).RowIndex).Value, _
                                        Me.TblItemWO, Me)
                                    ElseIf TabControl1.SelectedTab.Name = Me.TabPage_RM.Name Then
                                        LoadDgvItemRM(Me.DgvItem.Item(Me.ColItem_Identifier.Index, Me.DgvItem.SelectedCells.Item(0).RowIndex).Value, _
                                     Me.TblItemWO, Me.TblItemRM, Me)
                                    ElseIf TabControl1.SelectedTab.Name = Me.TabPage_WIP.Name Then
                                        Me.LoadItemWIP()
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End Try
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DtpTglTransaksi_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DtpTglTransaksi.ValueChanged
        'Dim urutan As Integer = 0
        'Dim rowWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woRow

        'Try
        '    If Me.TxtRemarks.Enabled = True Then
        '        If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.MO) = False Then
        '            Exit Sub
        '        End If
        '        If Me.Baru = False Then
        '            If Me.ObjMO.TanggalTransaksi.Month = Me.DtpTglTransaksi.Value.Month And Me.ObjMO.TanggalTransaksi.Year = Me.DtpTglTransaksi.Value.Year Then
        '                Me.TxtNoTransaksi.Text = Me.ObjMO.NoTransaksi
        '            Else
        '                Me.TxtNoTransaksi.Text = Me.ObjMO.GetNoBaru1(Me.DtpTglTransaksi.Value, ClsTrans_ManufacturingOrder.PilihFormatNumber.MO, -1, Me)
        '                'For i As Integer = 1 To Me.DgvItem.RowCount - 1
        '                '    Me.DgvItem.Item(Me.ColItem_JONumber.Index, i - 1).Value = Me.ObjMO.GetNoBaru(Me.DtpTglTransaksi.Value, ClsTrans_ManufacturingOrder.PilihFormatNumber.JO, Me, i - 1, Me.DgvItem.Item(Me.ColItem_Urutan.Index, i - 1).Value, Me.DgvItem.Item(Me.ColItem_Urutan.Index, i - 1).Value)
        '                'Next
        '            End If
        '        Else
        '            Me.TxtNoTransaksi.Text = Me.ObjMO.GetNoBaru1(Me.DtpTglTransaksi.Value, ClsTrans_ManufacturingOrder.PilihFormatNumber.MO, -1, Me)
        '            If Me.DgvItem.RowCount > 1 Then
        '                For i As Integer = 0 To Me.DgvItem.RowCount - 1
        '                    If Me.DgvItem.Item(Me.ColItem_PrimaryProduct.Index, i).Value <> "-1" Then
        '                        Me.DgvItem.Item(Me.ColItem_JONumber.Index, i).Value = Me.ObjMO.GetNoBaru1(Me.DtpTglTransaksi.Value, ClsTrans_ManufacturingOrder.PilihFormatNumber.JO, urutan, Me)
        '                        urutan = Me.ObjNumbering.GetUrutan(Me.DgvItem.Item(Me.ColItem_JONumber.Index, i).Value, Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.JO, Me.DtpTglTransaksi.Value)
        '                    End If
        '                Next
        '            End If
        '            If Me.TblItemWO.Rows.Count > 0 Then
        '                Dim count As Integer = 0
        '                Me.TblItemWO.DefaultView.RowFilter = Nothing
        '                Me.TblItemWO.DefaultView.Sort = "[Urutan] ASC"
        '                For Each dv As Data.DataRowView In Me.TblItemWO.DefaultView
        '                    rowWO = dv.Row
        '                    count += 1
        '                    rowWO.WO_Number = Me.ObjMO.GetNoBaru1(Me.DtpTglTransaksi.Value, ClsTrans_ManufacturingOrder.PilihFormatNumber.WO, count, Me, count, Me.TblItemWO)
        '                Next

        '                If Me.DgvItem.SelectedCells.Count > 0 Then
        '                    If String.IsNullOrEmpty(Me.DgvItem.Item(Me.ColItem_PrimaryProduct.Index, Me.DgvItem.SelectedCells.Item(0).RowIndex).Value) = False Then
        '                        If TabControl1.SelectedTab.Name = Me.TabPage_WO.Name Then
        '                            LoadDgvItemWO(Me.DgvItem.Item(Me.ColItem_Identifier.Index, Me.DgvItem.SelectedCells.Item(0).RowIndex).Value, _
        '                            Me.TblItemWO, Me.TGV_WO, Me)
        '                        ElseIf TabControl1.SelectedTab.Name = Me.TabPage_RM.Name Then
        '                            LoadDgvItemRM(Me.DgvItem.Item(Me.ColItem_Identifier.Index, Me.DgvItem.SelectedCells.Item(0).RowIndex).Value, _
        '                         Me.TblItemWO, Me.TblItemRM, Me.DGVRM, Me)
        '                        End If
        '                    End If
        '                End If
        '            End If
        '        End If
        '    End If
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'End Try
    End Sub

    Private Sub BackgroundWorker2_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Try
            Me.DgvItem.Rows.RemoveAt(EraseRowIndex)
            Me.ObjCD.SetNumber(Me.DgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetIndexItem(ByVal vPrimary As Double, _
      ByVal vRowIdxFocus As Integer) As Double
        Dim RowDGV As DataGridViewRow
        Dim Retval As Double = -1
        Try
            For ij As Integer = 0 To Me.DgvItem.Rows.Count - 2
                RowDGV = Me.DgvItem.Rows(ij)
                If vRowIdxFocus <> ij Then
                    If RowDGV.Cells(Me.ColItem_PrimaryProduct.Index).Value = vPrimary Then
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

    Private Sub DgvItem_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvItem.Sorted
        Try
            Me.ObjCD.SetNumber(Me.DgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItem_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles DgvItem.UserDeletedRow
        Try
            Me.ObjCD.SetNumber(Me.DgvItem)

            If Baru = False Then
                If Me.DgvItem.RowCount - 1 >= 0 Then
                    For ij As Integer = 0 To DgvItem.Rows.Count - 1
                        If Me.DgvItem.Item(Me.ColRM_CodeNo.Index, ij).Value <> "" Then
                            Me.DtpTglTransaksi.Enabled = False
                            Exit For
                        Else
                            Me.DtpTglTransaksi.Enabled = True
                            Exit For
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItem_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DgvItem.RowsAdded
        Dim rowDGV As DataGridViewRow
        Try
            Me.DgvItem.TabStop = True
            rowDGV = Me.DgvItem.Rows(Me.DgvItem.RowCount - 1)
            With rowDGV
                .Cells(Me.ColItem_No.Index).Value = rowDGV.Index + 1
                .Cells(Me.ColItem_PrimaryProduct.Index).Value = "-1"
                .Cells(Me.ColItem_CodeNo.Index).Value = ""
                .Cells(Me.ColItem_Desc.Index).Value = ""
                .Cells(Me.ColItem_QtyTarget.Index).Value = 0
                .Cells(Me.ColItem_QtyProduce.Index).Value = 0
                .Cells(Me.ColItem_Unit.Index).Value = ""
                .Cells(Me.ColItem_Remarks.Index).Value = ""
                .Cells(Me.ColItem_JONumber.Index).Value = ""
                .Cells(Me.ColItem_RefDocumentPrimary.Index).Value = "-1"
                .Cells(Me.ColItem_Identifier.Index).Value = Me.ObjMO_Item.GetNextIdentifierJO(Me.DgvItem, Me.TxtPrimary.Text, Me)
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            rowDGV = Nothing
        End Try
    End Sub

    Private Sub DgvItem_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles DgvItem.UserDeletingRow
        Try
            If ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_ManufacturingOrder_AddItem") = False Then
                MsgBox("You don't have access to perform this action.", MsgBoxStyle.Information, "Information")
                e.Cancel = True
                Exit Sub
            End If

            If Me.DgvItem.RowCount <= 1 Then
                Me.DgvItem.AllowUserToAddRows = True
            End If

            'cek apakah boleh hapus JO
            If Me.Baru = False Then
                If Me.ObjMO.ApakahBolehHapusItemJO(Me.TxtPrimary.Text, Me.DgvItem.Item(Me.ColItem_Identifier.Index, e.Row.Index).Value, _
                Me.TblItemWO, True) = False Then
                    e.Cancel = True
                    MsgBox("The following job order data is scheduled" & _
                    ", you can’t delete this data.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            Me.KosongkanTabelDetail(Me.DgvItem.Item(Me.ColItem_Identifier.Index, e.Row.Index).Value)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''' <summary>
    ''' tampilkan row yang diberikan pada parameter
    ''' </summary>
    ''' <param name="row"></param>
    ''' <remarks></remarks>
    Private Sub ShowRow(ByVal row As DataGridViewRow, ByVal Up As Boolean)
        Try
            'tentukan arah perpindahan, ke atas atau ke bawah
            If row.Index < Me.DgvItem.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke atas
                While row.Displayed = False
                    Me.DgvItem.FirstDisplayedScrollingRowIndex = Me.DgvItem.FirstDisplayedScrollingRowIndex - 1
                End While
            ElseIf row.Index > Me.DgvItem.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke bawah
                While Me.DgvItem.Rows(row.Index).Displayed = False
                    Me.DgvItem.FirstDisplayedScrollingRowIndex = Me.DgvItem.FirstDisplayedScrollingRowIndex + 1
                End While
                'kalo sampe sini pasti row yang disorot adalah row yang paling bawah, pastikan semuanya terlihat
                If Me.DgvItem.DisplayedRowCount(True) <> Me.DgvItem.DisplayedRowCount(False) Then
                    Me.DgvItem.FirstDisplayedScrollingRowIndex = Me.DgvItem.FirstDisplayedScrollingRowIndex + 1
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnUP_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUP.Click
        Try
            If Me.DgvItem.SelectedRows.Count > 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
                Dim row As DataGridViewRow = Me.DgvItem.SelectedRows(0)
                If row.Index <> 0 Then
                    Dim indexbaru As Integer = row.Index - 1
                    Me.ForceExit = True
                    Me.DgvItem.Rows.Remove(row)
                    Me.DgvItem.Rows.InsertRange(indexbaru, row)
                    Me.DgvItem.Focus()
                    Me.DgvItem.ClearSelection()
                    Me.ForceExit = True
                    row.Selected = True
                    Me.ObjCD.SetNumber(Me.DgvItem)
                    If row.Displayed = False Then
                        Me.ShowRow(row, True)
                    End If
                End If
                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDown_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDown.Click
        Try
            If Me.DgvItem.SelectedRows.Count > 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
                Dim row As DataGridViewRow = Me.DgvItem.SelectedRows(0)
                If row.Index <> Me.DgvItem.Rows.Count - 1 Then
                    Dim indexbaru As Integer = row.Index + 1
                    Me.ForceExit = True
                    Me.DgvItem.Rows.Remove(row)
                    Me.DgvItem.Rows.InsertRange(indexbaru, row)
                    Me.DgvItem.Focus()
                    Me.DgvItem.ClearSelection()
                    Me.ForceExit = True
                    row.Selected = True
                    Me.ObjCD.SetNumber(Me.DgvItem)
                    Me.ShowRow(row, False)
                End If
                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub New()
        Try
            InitializeComponent()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
        ' This call is required by the Windows Form Designer.


        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub TGV_WO_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles TGV_WO.DataError

    End Sub

    Private Sub btnPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim UseCustom As String = ""

        Try
            If Me.TxtPrimary.Text = "" OrElse Me.TxtPrimary.Text = -1 Then
                Me.btnFind.PerformClick()
            End If
            UseCustom = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.ManufacturingOrder, Core_Module.ClsTransaksiSetting.EnumNamaParameter.ManufacturingOrder_CustomPrintMO)

            If UseCustom = "" Or UseCustom = "0" Then
                Dim X1 As New FrmTrans_ManufacturingOrder_Print
                If Me.TxtPrimary.Text <> -1 AndAlso Me.TxtPrimary.Text <> "" Then
                    If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_ManufacturingOrder_PrintMOBlmOtorisasi") = False And Me.ChkScheduling.Checked = False Then
                        MsgBox("Please authorize first.", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If

                    Dim i As New FrmTrans_ManufacturingOrder_PrintOption
                    If i.ShowDialog <> Windows.Forms.DialogResult.OK Then
                        Exit Sub
                    End If
                    X1.PaperSize = i.PaperSize
                    X1.PrimaryMO = Me.TxtPrimary.Text
                    X1.ShowDialog()
                End If
            Else
                Dim X2 As New CustomSMESS.FrmCustom_ManufacturingOrder_PrintDialog
                If Me.TxtPrimary.Text <> -1 AndAlso Me.TxtPrimary.Text <> "" Then
                    If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_ManufacturingOrder_PrintMOBlmOtorisasi") = False And Me.ChkScheduling.Checked = False Then
                        MsgBox("Please authorize first.", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                    X2.UserCode = Me.UserCode
                    X2.PrimaryMo = Me.TxtPrimary.Text
                    X2.ShowDialog()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadOtorisasi()
        Try
            If Me.TxtPrimaryUserAuth.Text = "-1" Then
                Me.ChkScheduling.Checked = True
                Me.ObjMO.CariDetail(Me.TxtPrimary.Text, Me.DgvItem, Me.TblItemBom, Me.TblProsesOrder, _
                Me.TblItemWO, Me.TblItemRM, Me, Me.ObjCD, Me.TblItemRMAdd, Me.TblItemWIP, Me.TBLItemDelivery, Me.UserCode)
                Me.SalinKeField()
                Me.NonAktifkanDetail()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadCancelOtorisasi()
        Try
            If Me.TxtPrimaryUserAuth.Text <> "-1" Then
                Me.ChkScheduling.Checked = False
                Me.ObjMO.CariDetail(Me.TxtPrimary.Text, Me.DgvItem, Me.TblItemBom, Me.TblProsesOrder, _
                Me.TblItemWO, Me.TblItemRM, Me, Me.ObjCD, Me.TblItemRMAdd, Me.TblItemWIP, Me.TBLItemDelivery, Me.UserCode)
                Me.SalinKeField()
                Me.NonAktifkanDetail()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkScheduling_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkScheduling.CheckedChanged
        Try
            If Me.ChkScheduling.Checked = False Then
                vReadyForSchedule = ClsTrans_ManufacturingOrder.EnumPilihSchedule.Schedule_Aktif
            Else
                vReadyForSchedule = ClsTrans_ManufacturingOrder.EnumPilihSchedule.Schedule_tdk_Aktif
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkScheduling_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkScheduling.Click

        Try
            vPrimaryUser = ObjUser.GetPrimary(Me.UserCode)

            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_ManufacturingOrder_Authorize") = False Then
                MsgBox("You don't have access to perform this action.", MsgBoxStyle.Information, "Information")
                Me.btnFind.Focus()
                If Me.TxtNoTransaksi.Text = "" Then
                    Me.ChkScheduling.Checked = False
                Else
                    If ObjMO.ApakahSudahDiOtorisasi(Me.TxtPrimary.Text) Then
                        Me.ChkScheduling.Checked = True
                    Else
                        Me.ChkScheduling.Checked = False
                    End If
                End If
                Exit Sub
            End If

            '''kalo pr number kosong tidak bisa otorisasi
            If Me.TxtNoTransaksi.Text = "" Then
                MsgBox("Please select MO number first.", MsgBoxStyle.Information, "Information")
                Me.btnFind.Focus()
                Me.ChkScheduling.Checked = False
            End If

            '''pada saat otorisasi
            If Me.TxtNoTransaksi.Text <> "" And Me.TxtRemarks.Enabled = False Then
                If ObjMO.ApakahSudahDiOtorisasi(Me.TxtPrimary.Text) = False Then
                    If MsgBox("Authorize?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Dim frm As New Core_Module.DialogPassword
                        frm.txtUserName.Text = Me.UserCode
                        If frm.ShowDialog = Windows.Forms.DialogResult.OK Then
                            If Me.ObjMO.Authorize(vPrimaryUser, Me.TxtPrimary.Text, ClsTrans_ManufacturingOrder.EnumPilihSchedule.Schedule_Aktif, Me.UserCode) = True Then
                                MsgBox("Authorization complete")
                                Me.LoadOtorisasi()
                            End If
                        Else
                            ChkScheduling.Checked = False
                            Exit Sub
                        End If
                    Else
                        ChkScheduling.Checked = False
                        Exit Sub
                    End If
                Else
                    '''pada saat cancel otorisasi
                    'If ObjMO.ApakahSudahAdaBundling(Me.TxtPrimary.Text, Me.TblItemWO) = False Then '2019-03-19-iwan: bagian ini ditutup, asumsi tidak ada bundling
                    'klo blm di jadwal. klo sudah dijadwaln, otomatis tidak bisa cancel otorisasi
                    '    MsgBox("The following manufacturing order data is reference by one or several bundling" & _
                    '    ", you can’t cancel the authorization.", MsgBoxStyle.Information, "Information")
                    '    ChkScheduling.Checked = True
                    '    Exit Sub
                    If ObjMO.IsScheduled(Me.TxtPrimary.Text) Then
                        MsgBox("This manufacturing order data is already scheduled, you can’t cancel the authorization.", MsgBoxStyle.Information, "Information")
                        ChkScheduling.Checked = True
                        Exit Sub
                    ElseIf Me.ObjMO.GetClosedJOCount(Me.TxtPrimary.Text) > 0 Then
                        MsgBox("This manufacturing order has some JO that already closed, you can’t cancel the authorization.", MsgBoxStyle.Information, "Information")
                        ChkScheduling.Checked = True
                        Exit Sub
                    ElseIf Me.ObjMO.GetScheduledMOCount(Me.TxtPrimary.Text) > 0 Then
                        MsgBox("All or some JO for this MO already scheduled, you can’t cancel the authorization.", MsgBoxStyle.Information, "Information")
                        ChkScheduling.Checked = True
                        Exit Sub

                    Else
                        If MsgBox("Cancel authorization?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                            Dim frm As New Core_Module.DialogPassword
                            frm.txtUserName.Text = Me.UserCode
                            If frm.ShowDialog = Windows.Forms.DialogResult.OK Then
                                If Me.ObjMO.CancelAuthorize(#1/1/1950#, -1, Me.TxtPrimary.Text, ClsTrans_ManufacturingOrder.EnumPilihSchedule.Schedule_tdk_Aktif, Me.UserCode) = True Then
                                    MsgBox("Cancellation complete")
                                    Me.LoadCancelOtorisasi()
                                End If
                            Else
                                ChkScheduling.Checked = True
                                Exit Sub
                            End If
                        Else
                            ChkScheduling.Checked = True
                            Exit Sub
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TGV_WO_CellEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles TGV_WO.CellEnter

        Try
            'If e.ColumnIndex = Me.ColWO2_Remarks.Index Then
            '    If Me.TGV_WO.Item(Me.ColWO2_Remarks.Index, e.RowIndex).Value.Contains("#") = True Or Me.TGV_WO.Item(Me.ColWO2_Remarks.Index, e.RowIndex).Value.Contains("*") = True _
            '     Or Me.TGV_WO.Item(Me.ColWO2_Remarks.Index, e.RowIndex).Value.Contains("%") Or Me.TGV_WO.Item(Me.ColWO2_Remarks.Index, e.RowIndex).Value.Contains("#") = True _
            '     Or Me.TGV_WO.Item(Me.ColWO2_Remarks.Index, e.RowIndex).Value.Contains("`") = True Then
            '        'MsgBox("Remarks data can not contain the following characters: ` , #, % , or * .", MsgBoxStyle.Information, "Error")
            '        Me.TGV_WO.Item(Me.ColWO2_Remarks.Index, e.RowIndex).re()
            '    End If
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ClbProcess_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ClbProcess.Click

        Try
            ListPrimaryProses_Filter.Clear()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ClbProcess_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClbProcess.SelectedIndexChanged
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkLateWO_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkLateWO.CheckedChanged
        Try
            If Me.ChkLateWO.Checked = True Then

            Else

            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkGoodPerc_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkGoodPerc.CheckedChanged
        Try
            If Me.ChkGoodPerc.Checked = True Then

            Else

            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkInsufficientRM_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkInsufficientRM.CheckedChanged
        Try
            If TabControl1.SelectedTab.Name = Me.TabPage_RM.Name Then
                'LoadDgvItemRM(Me.DgvItem.Item(Me.ColItem_Identifier.Index, Me.DgvItem.SelectedCells.Item(0).RowIndex).Value)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TGV_WO_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TGV_WO.SelectionChanged
        'Dim a As String = ""
        'Try
        '    a = Me.TGV_WO.Item(Me.ColWO2_WONumber.Index, Me.TGV_WO.SelectedCells(0).RowIndex).Value
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'End Try
    End Sub

    Private Sub llbColumnSetting_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbColumnSetting.LinkClicked
        Dim ij As New FrmTrans_ManufacturingOrder_ColumnSetting
        Dim kl As New FrmTrans_ManufacturingOrder_WIP_ColumnSettings
        Try
            If TabControl1.SelectedTab.Name = Me.TabPage_WO.Name Then
                ij.UserCode = Me.UserCode
                If ij.ShowDialog() = Windows.Forms.DialogResult.OK Then Me.setColumnVisibility()
            ElseIf TabControl1.SelectedTab.Name = Me.TabPage_WIP.Name Then
                kl.UserCode = Me.UserCode
                If kl.ShowDialog = Windows.Forms.DialogResult.OK Then Me.SetColumnVisibilityWIP()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub setColumnVisibility()
        Dim ShowCodeNo, ShowDesc, ShowProcessCode, ShowProcessDesc As String

        Try
            ShowCodeNo = Me.objSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ManufacturingOrder, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ManufacturingOrder_ShowCodeNo, Me.UserCode)
            ShowDesc = Me.objSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ManufacturingOrder, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ManufacturingOrder_ShowDescription, Me.UserCode)
            ShowProcessCode = Me.objSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ManufacturingOrder, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ManufacturingOrder_ShowProcessCode, Me.UserCode)
            ShowProcessDesc = Me.objSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ManufacturingOrder, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ManufacturingOrder_ShowProcessDesc, Me.UserCode)

            If ShowCodeNo = "" And ShowDesc = "" And ShowProcessCode = "" And ShowProcessDesc = "" Then
                ShowCodeNo = "1"
                ShowDesc = "1"
                ShowProcessCode = "1"
                ShowProcessDesc = "1"
                Me.objSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ManufacturingOrder, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ManufacturingOrder_ShowCodeNo, ShowCodeNo, Me.UserCode)
                Me.objSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ManufacturingOrder, _
                         Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ManufacturingOrder_ShowDescription, ShowDesc, Me.UserCode)
                Me.objSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ManufacturingOrder, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ManufacturingOrder_ShowProcessCode, ShowProcessCode, Me.UserCode)
                Me.objSetting.SetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ManufacturingOrder, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ManufacturingOrder_ShowProcessDesc, ShowProcessDesc, Me.UserCode)
            End If
            If ShowCodeNo = "1" Then
                Me.ColWO2_CodeNo.Visible = True
            Else
                Me.ColWO2_CodeNo.Visible = False
            End If
            If ShowDesc = "1" Then
                Me.ColWO2_ProductDesc.Visible = True
            Else
                Me.ColWO2_ProductDesc.Visible = False
            End If
            If ShowProcessCode = "1" Then
                Me.ColWO2_ProcessCode.Visible = True
            Else
                Me.ColWO2_ProcessCode.Visible = False
            End If
            If ShowProcessDesc = "1" Then
                Me.ColWO2_ProcessDesc.Visible = True
            Else
                Me.ColWO2_ProcessDesc.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub SetColumnVisibilityWIP()
        Dim ShowWO, ShowProcessCode, ShowCodeNo, ShowDesc As String
        Try
            ShowWO = Me.objSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ManufacturingOrder, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ManufacturingOrder_WIP_ShowWO, Me.UserCode)
            ShowProcessCode = Me.objSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ManufacturingOrder, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ManufacturingOrder_WIP_ShowProcessCode, Me.UserCode)
            ShowCodeNo = Me.objSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ManufacturingOrder, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ManufacturingOrder_WIP_ShowCodeNo, Me.UserCode)
            ShowDesc = Me.objSetting.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.FrmTrans_ManufacturingOrder, Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.FrmTrans_ManufacturingOrder_WIP_ShowDesc, Me.UserCode)
            If ShowWO = "" Or ShowWO = "1" Then
                Me.ColWIP_WONo.Visible = True
            Else
                Me.ColWIP_WONo.Visible = False
            End If
            If ShowProcessCode = "" Or ShowProcessCode = "1" Then
                Me.ColWIP_Proses.Visible = True
            Else
                Me.ColWIP_Proses.Visible = False
            End If
            If ShowCodeNo = "" Or ShowCodeNo = "1" Then
                Me.ColWIP_CodeNo.Visible = True
            Else
                Me.ColWIP_CodeNo.Visible = False
            End If
            If ShowDesc = "" Or ShowDesc = "1" Then
                Me.ColWIP_Desc.Visible = True
            Else
                Me.ColWIP_Desc.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_ManufacturingOrder_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Try
            If Me.ObjMO.GetAuthStatus(Me.TxtPrimary.Text) Then
                Me.ChkScheduling.Checked = True
            Else
                Me.ChkScheduling.Checked = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_ManufacturingOrder_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        Dim RowItemProduct As DseTrans_ManufacturingOrder.tblitem_productRow
        Try
            Me.llbColumnSetting.Visible = False
            If Me.TabControl1.SelectedTab.Name = Me.TabPage_WO.Name Then
                Me.llbColumnSetting.Visible = True
            End If

            If Me.FromMP = True Then
                Me.btnFind.TabStop = False
                Me.btnNew.TabStop = False
                Me.btnEdit.TabStop = False
                Me.btnPrint.TabStop = False
                Me.btnDelete.TabStop = False
                Me.btnExit.TabStop = False

                Me.gboDetail.Size = New Point(Me.gboDetail.Width, Me.btnEdit.Top + Me.btnEdit.Height - Me.gboDetail.Top)
                Me.llbColumnSetting.Left = Me.gboDetail.Left + Me.btnDown.Left + Me.btnDown.Width + 5
                Me.gboDetail.BringToFront()
                Me.llbColumnSetting.BringToFront()

                If Me.NewMOFromMP = True Then
                    Me.btnStopProcess.Visible = True
                    Me.btnStopProcess.TabStop = True
                    Me.lblProgress.Visible = True
                    Me.prgAll.Visible = True
                    Me.prgAll.Value = 0
                    Me.prgAll.Maximum = Me.TotalIdx
                    Me.prgAll.Value = Me.IdxNow
                    Me.lblProgress.Text = "Data " + Me.prgAll.Value.ToString + " of " + Me.prgAll.Maximum.ToString
                Else
                    Me.btnStopProcess.Visible = False
                    Me.btnStopProcess.TabStop = False
                    Me.lblProgress.Visible = False
                    Me.prgAll.Visible = False
                End If
            Else
                Me.btnStopProcess.Visible = False
                Me.btnStopProcess.TabStop = False
                Me.lblProgress.Visible = False
                Me.prgAll.Visible = False
            End If
            If Me.FromMP = True And Me.TBLItemProduct.Rows.Count > 0 Then
                If Me.NewMOFromMP = True Then
                    Me.ExecuteSelectionChange = False
                    Me.btnNew.PerformClick()
                    Me.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                    Me.DgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                    Me.DgvItem.AllowUserToAddRows = False
                    For Each RowItemProduct In Me.TBLItemProduct.Rows
                        Me.AddNewProduct(RowItemProduct.Primary_Product, RowItemProduct.Qty_Target)
                        If Me.TxtRemarks.Text.Trim = "" Then
                            Me.TxtRemarks.Text = RowItemProduct.remarks
                        Else
                            Me.TxtRemarks.Text = Me.TxtRemarks.Text + "; " + RowItemProduct.remarks
                        End If
                    Next
                    Me.ExecuteSelectionChange = True
                    Me.DgvItem.AllowUserToAddRows = True
                    Me.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                    Me.DgvItem.ClearSelection()
                    Me.DgvItem.Item(0, 0).Selected = True
                Else
                    'cari data mo
                    If Me.ObjMO.CariDetail(Me.PrimaryMOFromMP, Me.DgvItem, Me.TblItemBom, Me.TblProsesOrder, Me.TblItemWO, _
                    Me.TblItemRM, Me, Me.ObjCD, Me.TblItemRMAdd, Me.TblItemWIP, Me.TBLItemDelivery, Me.UserCode, Me.TblJO_Temp) = True Then
                        Me.SalinKeField()
                        Me.DgvItem.ClearSelection()
                        Me.DGVRM.Rows.Clear()
                        Me.TGV_WO.Nodes.Clear()

                        Me.DgvItem.Item(0, 0).Selected = True
                        Me.btnEdit.Focus()
                    Else
                        Me.btnFind.Focus()
                    End If
                    Me.btnFind.Focus()
                    Me.NonAktifkanDetail()

                    Me.btnEdit.PerformClick()

                    Me.ExecuteSelectionChange = False
                    Me.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                    Me.DgvItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
                    Me.DgvItem.AllowUserToAddRows = False
                    For Each RowItemProduct In Me.TBLItemProduct.Rows
                        Me.AddNewProduct(RowItemProduct.Primary_Product, RowItemProduct.Qty_Target)
                    Next
                    Me.ExecuteSelectionChange = True
                    Me.DgvItem.AllowUserToAddRows = True
                    Me.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                    Me.DgvItem.ClearSelection()
                    Me.DgvItem.Item(0, 0).Selected = True
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AddNewProduct(ByVal vPrimaryProduct As Double, ByVal vQty As Decimal)
        Dim row As DataGridViewRow
        Dim jk As New FrmTrans_ManufacturingOrder_RMReqDate

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            Me.DgvItem.Rows.Add()
            row = Me.DgvItem.Rows(Me.DgvItem.Rows.Count - 1)

            With row

                .Cells(Me.ColItem_PrimaryProduct.Index).Value = vPrimaryProduct
                .Cells(Me.ColItem_CodeNo.Index).Value = Me.ObjProduct.GetProductCode(vPrimaryProduct)
                If String.IsNullOrEmpty(Me.DgvItem.Item(Me.ColItem_Urutan.Index, row.Index).Value) = True Then
                    Me.DgvItem.Item(Me.ColItem_JONumber.Index, row.Index).Value = Me.ObjMO.GetNoBaru1(Me.DtpTglTransaksi.Value, _
                    ClsTrans_ManufacturingOrder.PilihFormatNumber.JO, Me.getMaxNoUrutJO, Me, row.Index)
                End If
                Me.DgvItem.Item(Me.ColItem_Desc.Index, row.Index).Value = Me.ObjProduct.GetDescription(vPrimaryProduct)
                Me.DgvItem.Item(Me.ColItem_Decimal.Index, row.Index).Value = Me.ObjProduct.GetAngkaDecimal(vPrimaryProduct)
                Me.DgvItem.Item(Me.ColItem_Unit.Index, row.Index).Value = Me.ObjProduct.GetUnitCode(vPrimaryProduct)
                Me.DgvItem.Item(Me.ColItem_QtyTarget.Index, row.Index).Value = Me.ObjCD.DigitGrouping(vQty, _
                Me.DgvItem.Item(Me.ColItem_Decimal.Index, row.Index).Value)
                Me.DgvItem.Item(Me.ColItem_TargetFinishDate.Index, row.Index).Value = Format(Me.DtpTargetFinishDate.Value, "dd MMM yyyy")
                Me.DgvItem.Item(Me.ColItem_TargetFinishDateR1.Index, row.Index).Value = "-"
                Me.DgvItem.Item(Me.ColItem_TargetFinishDateR2.Index, row.Index).Value = "-"
                Me.DgvItem.Item(Me.ColItem_ActualFinishDate.Index, row.Index).Value = "-"

                'fungsi untuk update table item BOM, WO, RM, Proses Order
                jk.btnCancel.Enabled = False
                jk.btnCancel.Visible = False
                jk.Size = New Point(jk.Size.Width, jk.Size.Height - jk.btnCancel.Height)
                jk.ShowDialog()
                Me.KosongkanTabelDetail(Me.DgvItem.Item(Me.ColItem_Identifier.Index, row.Index).Value)
                Me.IsiAllTblItem(vPrimaryProduct, Me.DgvItem.Item(Me.ColItem_QtyTarget.Index, row.Index).Value, _
                Me.DgvItem.Item(Me.ColItem_Identifier.Index, row.Index).Value, Me.DgvItem.Item(Me.ColItem_Urutan.Index, row.Index).Value, True, jk.dtpRequired.Value.Date, True)

                'isi qty produce
                Me.DgvItem.Item(Me.ColItem_QtyProduce.Index, row.Index).Value = Me.ObjCD.DigitGrouping(ObjMO.GetQtyProduce_InTbl(Me.DgvItem.Item(Me.ColItem_PrimaryProduct.Index, row.Index).Value, _
                Me.DgvItem.Item(Me.ColItem_QtyTarget.Index, row.Index).Value, Me.DgvItem.Item(Me.ColItem_Identifier.Index, row.Index).Value, Me.TblItemBom), _
                Me.DgvItem.Item(Me.ColItem_Decimal.Index, row.Index).Value, Core_Module.ClsCekData.EnumTipeRounding.Up)

                Me.JO_TempADD(Me.DgvItem.Item(Me.ColItem_PrimaryProduct.Index, row.Index).Value, Me.DgvItem.Item(Me.ColItem_Identifier.Index, row.Index).Value)

                Me.ObjCD.SetNumber(Me.DgvItem)

                Me.DgvItem.Item(Me.ColItem_DeliveryDate.Index, row.Index).Value = Format(Me.DtpTargetFinishDate.Value, "dd MMM yyyy")
                Dim rowDelivery As DseTrans_ManufacturingOrder.tbltrans_mo_itemdeliveryRow = Me.TBLItemDelivery.Newtbltrans_mo_itemdeliveryRow
                With rowDelivery
                    .primary = -1
                    .primary_mo = Me.TxtPrimary.Text
                    .identifier_jo = Me.DgvItem.Item(Me.ColItem_Identifier.Index, row.Index).Value
                    .delivery_qty = Me.DgvItem.Item(Me.ColItem_QtyTarget.Index, row.Index).Value
                    .delivery_date = Me.DtpTargetFinishDate.Value
                    .remarks = String.Empty
                End With
                TBLItemDelivery.Rows.Add(rowDelivery)

                'clear semua dgv untuk triger load ulang dgv yang bersesuaian
                Me.TGV_WO.Nodes.Clear()
                Me.TGV_WO.Rows.Clear()
                Me.DGVRM.Rows.Clear()
                Me.DgvRMRekap.Rows.Clear()
                Me.dgvItemWO.Rows.Clear()
                Me.DGVWIP.Rows.Clear()

                ''set qty target
                'Me.DgvItem.Item(Me.ColItem_QtyTarget.Index, row.Index).Value = Me.ObjCD.DigitGrouping(vQty, Me.DgvItem.Item(Me.ColItem_Decimal.Index, row.Index).Value)

                'Me.DgvItem.Item(Me.ColItem_QtyProduce.Index, row.Index).Value = Me.ObjCD.DigitGrouping(ObjMO.GetQtyProduce_InTbl(Me.DgvItem.Item(Me.ColItem_PrimaryProduct.Index, row.Index).Value, _
                'Me.DgvItem.Item(Me.ColItem_QtyTarget.Index, row.Index).Value, Me.DgvItem.Item(Me.ColItem_Identifier.Index, row.Index).Value, Me.TblItemBom), _
                'Me.DgvItem.Item(Me.ColItem_Decimal.Index, row.Index).Value, Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)

                'Me.IsiTblItemWORM(-1, Me.DgvItem.Item(Me.ColItem_PrimaryProduct.Index, row.Index).Value, _
                'Me.DgvItem.Item(Me.ColItem_QtyTarget.Index, row.Index).Value, _
                'Me.TblItemWO, Me.TblItemRM, Me.TblItemBom, Me.DgvItem.Item(Me.ColItem_Identifier.Index, row.Index).Value, "-", "-", Me.TblItemWIP, jk.dtpRequired.Value.Date, True)
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnStopProcess_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnStopProcess.Click
        Try
            If MsgBox("Stop creating mo?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                Me.StopProcess = True
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkRM_IncludeWIP_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkRM_IncludeWIP.CheckedChanged
        Try
            If Me.ExecuteChkRM = True Then
                If Me.DgvItem.SelectedCells.Count > 0 Then
                    LoadDgvItemRM(Me.DgvItem.Item(Me.ColItem_Identifier.Index, Me.DgvItem.SelectedCells.Item(0).RowIndex).Value, _
                                Me.TblItemWO, Me.TblItemRM, Me)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkRM_ShowDetail_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkRM_ShowDetail.CheckedChanged
        Try
            If Me.ExecuteChkRM = True Then
                If Me.DgvItem.SelectedCells.Count > 0 Then
                    LoadDgvItemRM(Me.DgvItem.Item(Me.ColItem_Identifier.Index, Me.DgvItem.SelectedCells.Item(0).RowIndex).Value, _
                                Me.TblItemWO, Me.TblItemRM, Me)
                Else
                    Me.DGVRM.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                    Me.DGVRM.Rows.Clear()
                    Me.DGVRM.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                    Me.DgvRMRekap.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                    Me.DgvRMRekap.Rows.Clear()
                    Me.DgvRMRekap.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                    If Me.chkRM_ShowDetail.Checked = True Then
                        Me.DGVRM.BringToFront()
                    Else
                        Me.DgvRMRekap.BringToFront()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub rboWO_UseTreeView_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rboWO_UseTreeView.CheckedChanged
        Try
            If Me.ExecuteWO = True Then
                If Me.DgvItem.SelectedCells.Count > 0 Then
                    Me.LoadDgvItemWO(Me.DgvItem.Item(Me.ColItem_Identifier.Index, Me.DgvItem.SelectedCells.Item(0).RowIndex).Value, _
                                Me.TblItemWO, Me)
                Else
                    Me.TGV_WO.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                    Me.TGV_WO.Nodes.Clear()
                    Me.TGV_WO.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                    Me.dgvItemWO.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                    Me.dgvItemWO.Rows.Clear()
                    Me.dgvItemWO.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                    If Me.rboWO_UseTreeView.Checked = True Then
                        Me.TGV_WO.BringToFront()
                    Else
                        Me.dgvItemWO.BringToFront()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemWO_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItemWO.CellClick
        Try
            If e.RowIndex <> -1 Then
                If e.ColumnIndex = Me.ColWOTbl_ShowDetail.Index Then
                    Dim i As New FrmTrans_ManufacturingOrder_WO
                    i.vMONo = Me.TxtNoTransaksi.Text
                    i.vJONo = ObjMO.GetJONumber_InDgvJO(Me.dgvItemWO.Item(Me.ColWOTbl_IdentifierJO.Index, e.RowIndex).Value, Me.DgvItem, Me)
                    i.vIdentifierWO = Me.dgvItemWO.Item(Me.ColWO2_IdentifierItem.Index, e.RowIndex).Value
                    i.vIdentifierJO = Me.dgvItemWO.Item(Me.ColWO2_IdentifierJO.Index, e.RowIndex).Value
                    i.vTblItemWO = Me.TblItemWO
                    i.vTblItemRM = Me.TblItemRM
                    i.UserCode = Me.UserCode
                    If Me.btnOK.Enabled = True Then
                        i.vBolehEdit = True
                    End If
                    i.ShowDialog()
                    i.Dispose()

                    'Load ulang Tgv WO / Dgv RM
                    If Me.DgvItem.SelectedCells.Count > 0 Then
                        If String.IsNullOrEmpty(Me.DgvItem.Item(Me.ColItem_PrimaryProduct.Index, Me.DgvItem.SelectedCells.Item(0).RowIndex).Value) = False Then
                            If TabControl1.SelectedTab.Name = Me.TabPage_WO.Name Then
                                LoadDgvItemWO(Me.DgvItem.Item(Me.ColItem_Identifier.Index, Me.DgvItem.SelectedCells.Item(0).RowIndex).Value, _
                                Me.TblItemWO, Me)
                            ElseIf TabControl1.SelectedTab.Name = Me.TabPage_RM.Name Then
                                LoadDgvItemRM(Me.DgvItem.Item(Me.ColItem_Identifier.Index, Me.DgvItem.SelectedCells.Item(0).RowIndex).Value, _
                                 Me.TblItemWO, Me.TblItemRM, Me)
                            ElseIf TabControl1.SelectedTab.Name = Me.TabPage_WIP.Name Then
                                Me.LoadItemWIP()
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemWO_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItemWO.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvItemWO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbAddWIP_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbAddWIP.LinkClicked
        Dim rowDGV As DataGridViewRow
        Dim frm_ij As New FrmTrans_ManufacturingOrder_AddWIP
        Dim RowItemWO, RowItemWODummy As DseTrans_ManufacturingOrder.tbltrans_mo_item_woRow
        Dim RowItemRM, RowItemRMDummy As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmRow
        Dim RowItemWIP As DseTrans_ManufacturingOrder.tbltrans_mo_item_wipRow
        Dim ObjUser As New Core_Module.ClsUser

        Try
            If Me.TxtRemarks.Enabled = True Then
                If ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_ManufacturingOrder_InputWIP") = False Then
                    MsgBox("You don't have acces to perform this action", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
                If Me.DgvItem.SelectedCells.Count > 0 Then
                    rowDGV = Me.DgvItem.Rows(Me.DgvItem.SelectedCells(0).RowIndex)
                    If Me.ObjMO.IsScheduled(Me.TxtPrimary.Text, rowDGV.Cells(Me.ColItem_Identifier.Index).Value.ToString) = True Then
                        MsgBox("This job order is already scheduled, you can't add/edit WIP allocation.", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If

                    frm_ij.vIdentifierJO = rowDGV.Cells(Me.ColItem_Identifier.Index).Value.ToString
                    frm_ij.vTblItemRM_Original = Me.TblItemRM
                    frm_ij.vTblItemWO_Original = Me.TblItemWO
                    frm_ij.BackColorAll = Me.BackColorAll
                    frm_ij.vUserCode = Me.UserCode
                    frm_ij.vJONO = rowDGV.Cells(Me.ColItem_JONumber.Index).Value
                    frm_ij.vMONo = Me.TxtNoTransaksi.Text
                    frm_ij.vMODate = Me.DtpTglTransaksi.Value.Date
                    frm_ij.vPrimaryMO = Me.TxtPrimary.Text
                    frm_ij.vTblItemBoM_Original = Me.TblItemBom

                    If frm_ij.ShowDialog() = Windows.Forms.DialogResult.OK Then
                        'coding update nilai wip
                        '1. item rm
                        For hj As Integer = Me.TblItemRM.Rows.Count - 1 To 0 Step -1
                            RowItemRM = Me.TblItemRM.Rows(hj)
                            If RowItemRM.Identifier_JO = frm_ij.vIdentifierJO Then
                                Me.TblItemRM.Rows.Remove(RowItemRM)
                            End If
                        Next
                        For Each RowItemRMDummy In frm_ij.vTblItemRM.Rows
                            RowItemRM = Me.TblItemRM.Newtbltrans_mo_item_rmRow
                            For mn As Integer = 0 To Me.TblItemRM.Columns.Count - 1
                                RowItemRM.Item(mn) = RowItemRMDummy.Item(mn)
                            Next
                            Me.TblItemRM.Rows.Add(RowItemRM)
                        Next
                        '2. item wo
                        For hj As Integer = Me.TblItemWO.Rows.Count - 1 To 0 Step -1
                            RowItemWO = Me.TblItemWO.Rows(hj)
                            If RowItemWO.Identifier_JO = frm_ij.vIdentifierJO Then
                                Me.TblItemWO.Rows.Remove(RowItemWO)
                            End If
                        Next
                        For Each RowItemWODummy In frm_ij.vTblItemWO.Rows
                            RowItemWO = Me.TblItemWO.Newtbltrans_mo_item_woRow
                            For mn As Integer = 0 To Me.TblItemWO.Columns.Count - 1
                                RowItemWO.Item(mn) = RowItemWODummy.Item(mn)
                            Next
                            Me.TblItemWO.Rows.Add(RowItemWO)
                        Next
                        '3. item wip
                        For hj As Integer = Me.TblItemWIP.Rows.Count - 1 To 0 Step -1
                            RowItemWIP = Me.TblItemWIP.Rows(hj)
                            If RowItemWIP.Identifier_JO = frm_ij.vIdentifierJO Then
                                Me.TblItemWIP.Rows.Remove(RowItemWIP)
                            End If
                        Next
                        For Each RowItemRMDummy In frm_ij.vTblItemRM.Rows
                            If RowItemRMDummy.WIP_Qty > 0 Then
                                RowItemWIP = Me.TblItemWIP.Newtbltrans_mo_item_wipRow
                                RowItemWIP.Identifier_JO = frm_ij.vIdentifierJO
                                RowItemWIP.Identifier_WO = RowItemRMDummy.Identifier_WO
                                RowItemWIP.Primary = -1
                                RowItemWIP.Primary_MO = Me.TxtPrimary.Text
                                RowItemWIP.Primary_Product_Code_WIP = RowItemRMDummy.Primary_Product_Code_RM
                                RowItemWIP.Qty = RowItemRMDummy.WIP_Qty
                                RowItemWIP.Qty_Used = 0
                                RowItemWIP.Qty_Balance = RowItemWIP.Qty - RowItemWIP.Qty_Used
                                RowItemWIP.Remarks = ""

                                Me.TblItemWIP.Rows.Add(RowItemWIP)
                            End If
                        Next

                        'load ulang dgv wip
                        Me.LoadItemWIP()
                        LoadDgvItemWO(Me.DgvItem.Item(Me.ColItem_Identifier.Index, Me.DgvItem.SelectedCells.Item(0).RowIndex).Value, _
                        Me.TblItemWO, Me)
                        LoadDgvItemRM(Me.DgvItem.Item(Me.ColItem_Identifier.Index, Me.DgvItem.SelectedCells.Item(0).RowIndex).Value, _
                        Me.TblItemWO, Me.TblItemRM, Me)
                    End If
                End If

            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVWIP_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DGVWIP.Sorted
        Try
            Me.ObjCD.SetNumber(Me.DGVWIP)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub 
End Class