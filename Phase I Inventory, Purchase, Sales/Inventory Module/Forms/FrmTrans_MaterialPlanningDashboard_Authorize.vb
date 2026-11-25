Option Explicit On
Option Strict Off

Imports Core_Module

Public Class FrmTrans_MaterialPlanningDashboard_Authorize

    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public UserCode As String = String.Empty
    Public AuthLevel As Integer = 0

    Dim ObjSettingUser As New Core_Module.ClsTransaksiSettingUser
    Dim DTAProduct_MU As New dseMasterProductTableAdapters.tblmaster_product5TableAdapter
    Dim TBLProduct_MU As New dseMasterProduct.tblmaster_product5DataTable
    Dim dtaMultiUnit As New dseMasterProductTableAdapters.tblmaster_product_item_multiunit2TableAdapter
    Dim tblMultiUnit As New dseMasterProduct.tblmaster_product_item_multiunit2DataTable
    Dim rowMultiUnit As dseMasterProduct.tblmaster_product_item_multiunit2Row
    Dim vHitungQtyMultiUnit As Boolean = True
    Dim ObjSupplier As New Core_Module.clsMasterSupplier
    Dim ObjCurrency As New Core_Module.clsMasterCurrency
    Dim ObjProd As New Core_Module.clsMasterProduct
    Dim ObjPO As New ClsTransaksiPO

    Dim vDT As New DseTransaksiMaterialPlanning2.tbltrans_material_planning_item1DataTable
    Dim vDT_ItemGIT As New DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemgitDataTable
    Dim vDT_ItemDO As New DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemdoDataTable
    Dim vDT_ItemPOB As New DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempobDataTable
    Dim vDT_ItemSOB As New DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsobDataTable
    Dim vDT_ItemSOBDelivery As New DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsob_deliveryDataTable
    Dim vDT_ItemPrOB As New DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_reqDataTable
    Dim vDT_ItemPrOBDelivery As New DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_req_deliveryDataTable
    Dim vDT_ItemSettingLocation As New DseTransaksiMaterialPlanning_WarehouseExc.tbltrans_material_planning_item_location_exclusionDataTable
    Dim vDT_ItemQty As New DseTransaksiMaterialPlanning_WarehouseExc.tbltrans_material_planning_item_qtyDataTable
    Dim vDT_ItemMOB As New DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemmobDataTable
    Dim vDT_ItemRMReq As New DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_rm_reqDataTable
    Dim vDT_ItemFollowUp As New DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_followupDataTable
    Dim TBLHistoricalPrice As New DseTransaksiPO.DataTable9DataTable

    Dim objMP As New ClsTransaksiMaterialPlanning
    Dim objCD As New Core_Module.ClsCekData
    Dim objUser As New Core_Module.ClsUser
    Dim buffer As Object
    Dim dgvRow As DataGridViewRow
    Dim rowIndexToDelete As Integer

    Enum EnumStatus As Integer
        UNAUTHORIZED_ONLY
        AUTHORIZED_ONLY
        ALL
    End Enum

    Private Sub FrmTrans_MaterialPlanningDashboard_Authorize_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Try
            Me.btnFilter.PerformClick()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_MaterialPlanningDashboard_Authorize_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.dgvAuth.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvCancel.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvAll.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvUser.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvMP.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.chkFilterDate.Checked = True
            Me.chkFilterDate.Checked = False
            Me.cboFilterStatus.SelectedIndex = EnumStatus.UNAUTHORIZED_ONLY
            Me.cboFilterStatus.Enabled = False
            Me.CancelButton = Me.btnExit

            Me.ColMP_QtyGIT.Visible = False
            Me.ColMP_QtyInquiry.Visible = False
            Me.ColMP_QtyLeadTime.Visible = False
            Me.ColMP_QtyOnHand.Visible = False
            Me.ColMP_QtyOrder.Visible = True
            Me.ColMP_QtyOrderMaximum.Visible = False
            Me.ColMP_QtyOrderOptimum.Visible = False
            Me.ColMP_QtyPendingDO.Visible = False
            Me.ColMP_QtyPOB.Visible = False
            Me.ColMP_QtyRequired.Visible = True
            Me.ColMP_QtySafetyStock.Visible = False
            Me.ColMP_QtySOB.Visible = False
            Me.ColMP_QtyOrderMinimum.Visible = False
            Me.ColMP_OrderStrategy.Visible = False
            Me.ColMP_PreferedSupplier.Visible = False
            Me.ColMP_NettPrice.Visible = False
            Me.ColMP_Disc.Visible = False
            Me.ColMP_Currency.Visible = False
            Me.ColMP_UnitPrice.Visible = False
            Me.ColMP_SubTotal.Visible = False
            Me.ColMP_UnitPrice.Visible = False
            Me.dgvMP.ReadOnly = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        Try
            If Me.TabControl1.SelectedTab.Name = Me.TabPage1.Name Then
                Me.cboFilterStatus.SelectedIndex = EnumStatus.UNAUTHORIZED_ONLY
                Me.cboFilterStatus.Enabled = False
                Me.ToolStripLabel2.Text = "Authorize"
                Me.ToolStripLabel2.Visible = True
            ElseIf Me.TabControl1.SelectedTab.Name = Me.TabPage2.Name Then
                Me.cboFilterStatus.SelectedIndex = EnumStatus.AUTHORIZED_ONLY
                Me.cboFilterStatus.Enabled = False
                Me.ToolStripLabel2.Text = "Cancel Authorize"
                Me.ToolStripLabel2.Visible = True
            ElseIf Me.TabControl1.SelectedTab.Name = Me.TabPage3.Name Then
                Me.cboFilterStatus.SelectedIndex = EnumStatus.ALL
                Me.cboFilterStatus.Enabled = True
                Me.ToolStripLabel2.Visible = False
            End If
            Me.btnFilter.PerformClick()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkFilterDate_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkFilterDate.CheckedChanged
        Try
            If Me.chkFilterDate.Checked Then
                Me.dtpFilterFrom.Enabled = True
                Me.dtpFilterTo.Enabled = True
            Else
                Me.dtpFilterFrom.Enabled = False
                Me.dtpFilterTo.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFilterFindMP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilterFindMP.Click
        Dim frmFindMP As New FrmTrans_MaterialPlanningDashboard_FindMP
        Try
            frmFindMP.txtCari.Text = Me.txtFilterDocNo.Text
            frmFindMP.ShowDialog()
            If frmFindMP.PrimaryTransaksi > -1 Then
                Me.txtFilterPrimaryMP.Text = frmFindMP.PrimaryTransaksi
                Me.txtFilterDocNo.Text = Me.objMP.GetDocumentNo(frmFindMP.PrimaryTransaksi).ToUpper
                btnFilter.Focus()
            Else
                Me.txtFilterPrimaryMP.Text = "-1"
                Me.txtFilterDocNo.Text = String.Empty
                Me.btnFilterFindMP.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtFilterDocNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFilterDocNo.GotFocus
        Try
            Me.txtFilterDocNo.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtFilterDocNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFilterDocNo.KeyDown
        Try
            If e.KeyCode = 13 Then
                Me.txtFilterDocNo.Text = Me.txtFilterDocNo.Text.Trim
                If Me.objMP.ApakahSudahAda(Me.txtFilterDocNo.Text) Then
                    Me.txtFilterPrimaryMP.Text = Me.objMP.GetPrimary(Me.txtFilterDocNo.Text)
                    Me.btnFilter.Focus()
                Else
                    Me.btnFilterFindMP.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtFilterDocNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFilterDocNo.Leave
        Try
            Me.txtFilterDocNo.Text = Me.txtFilterDocNo.Text.Trim
            If Me.txtFilterDocNo.Text.Length > 0 Then
                If Me.objMP.ApakahSudahAda(Me.txtFilterDocNo.Text) Then
                    Me.txtFilterPrimaryMP.Text = Me.objMP.GetPrimary(Me.txtFilterDocNo.Text)
                Else
                    Me.btnFilterFindMP.PerformClick()
                End If
            Else
                Me.txtFilterPrimaryMP.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtFilterDocNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFilterDocNo.LostFocus
        Try
            Me.txtFilterDocNo.BackColor = Me.txtFilterPrimaryMP.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvUser_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvUser.CellBeginEdit
        Try
            If e.ColumnIndex <> Me.col4UserCode.Index Then
                e.Cancel = True
                Exit Sub
            End If
            Me.buffer = Me.dgvUser.Item(e.ColumnIndex, e.RowIndex).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvUser_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvUser.CellEndEdit
        Try
            If e.ColumnIndex = Me.col4UserCode.Index Then
                If IsNothing(Me.dgvUser.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If String.IsNullOrEmpty(Me.dgvUser.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        Me.dgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvUser.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.ToUpper
                        If Me.buffer = Me.dgvUser.Item(e.ColumnIndex, e.RowIndex).Value Then
                            Exit Sub
                        End If
                        If Me.IsExistUserCode(Me.dgvUser.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex) = True Then
                            MsgBox("User: " & Me.dgvUser.Item(e.ColumnIndex, e.RowIndex).Value & " already exist.", MsgBoxStyle.Information, "Information")
                            If String.IsNullOrEmpty(Me.buffer) Then
                                Me.rowIndexToDelete = e.RowIndex
                                Me.BackgroundWorker1.RunWorkerAsync()
                            Else
                                Me.dgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer.ToString
                            End If
                        End If
                        If Me.objUser.ApakahSudahAda(Me.dgvUser.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            Me.dgvUser.Item(Me.col4UserName.Index, e.RowIndex).Value = Me.objUser.GetNama(Me.dgvUser.Item(e.ColumnIndex, e.RowIndex).Value)
                            Me.dgvUser.Item(Me.col4Primary.Index, e.RowIndex).Value = Me.objUser.GetPrimary(Me.dgvUser.Item(e.ColumnIndex, e.RowIndex).Value)
                        Else
                            Dim frmFindUser As New Core_Module.FrmCariMasterUser
                            frmFindUser.TxtCari.Text = Me.dgvUser.Item(e.ColumnIndex, e.RowIndex).Value
                            frmFindUser.ShowDialog()
                            If frmFindUser.Kodeuser <> "" Then
                                Me.dgvUser.Item(e.ColumnIndex, e.RowIndex).Value = frmFindUser.Kodeuser.ToUpper
                                Me.dgvUser.Item(Me.col4UserName.Index, e.RowIndex).Value = Me.objUser.GetNama(Me.dgvUser.Item(e.ColumnIndex, e.RowIndex).Value)
                                Me.dgvUser.Item(Me.col4Primary.Index, e.RowIndex).Value = Me.objUser.GetPrimary(Me.dgvUser.Item(e.ColumnIndex, e.RowIndex).Value)
                            Else
                                If String.IsNullOrEmpty(Me.buffer) Then
                                    Me.rowIndexToDelete = e.RowIndex
                                    Me.BackgroundWorker1.RunWorkerAsync()
                                Else
                                    Me.dgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer.ToString
                                End If
                            End If
                        End If
                    Else
                        If String.IsNullOrEmpty(Me.buffer) = True Then
                            Me.rowIndexToDelete = e.RowIndex
                            Me.BackgroundWorker1.RunWorkerAsync()
                        Else
                            Me.dgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer.ToString
                        End If
                    End If
                Else
                    If String.IsNullOrEmpty(Me.buffer) = True Then
                        Me.rowIndexToDelete = e.RowIndex
                        Me.BackgroundWorker1.RunWorkerAsync()
                    Else
                        Me.dgvUser.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer.ToString
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvUser_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvUser.DataError

    End Sub

    Private Sub dgvUser_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvUser.RowsAdded
        Try
            Me.dgvRow = Me.dgvUser.Rows(Me.dgvUser.Rows.Count - 1)
            Me.dgvRow.Cells(Me.col4No.Index).Value = Me.dgvUser.Rows.Count
            Me.dgvRow.Cells(Me.col4UserCode.Index).Value = String.Empty
            Me.dgvRow.Cells(Me.col4UserName.Index).Value = String.Empty
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvUser_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvUser.RowsRemoved
        Try
            Me.objCD.SetNumber(Me.dgvUser)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvUser_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvUser.Sorted
        Try
            Me.objCD.SetNumber(Me.dgvUser)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(10)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            If Me.rowIndexToDelete > -1 And Me.rowIndexToDelete <> Me.dgvUser.Rows.Count - 1 Then
                Me.dgvUser.Rows.RemoveAt(Me.rowIndexToDelete)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function IsExistUserCode(ByVal vUserCode As String, ByVal vCurrentRow As Integer) As Boolean
        Try
            Me.dgvUser.AllowUserToAddRows = False
            For Each dgvRow In Me.dgvUser.Rows
                If dgvRow.Index <> vCurrentRow Then
                    If dgvRow.Cells(Me.col4UserCode.Index).Value = vUserCode Then
                        Return True
                    End If
                End If
            Next
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            Me.dgvUser.AllowUserToAddRows = True
        End Try
    End Function

    Private Sub ToolStripLabel1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripLabel1.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter.Click
        Try
            Me.Filter()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Function GetListOfUsers() As List(Of String)
        GetListOfUsers = New List(Of String)
        Try
            Me.dgvUser.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For Each Me.dgvRow In Me.dgvUser.Rows
                If String.IsNullOrEmpty(Me.dgvRow.Cells(Me.col4UserCode.Index).Value) = False Then
                    If GetListOfUsers.Contains(Me.dgvRow.Cells(Me.col4UserCode.Index).Value) = False Then
                        GetListOfUsers.Add(Me.dgvRow.Cells(Me.col4UserCode.Index).Value)
                    End If
                End If
            Next
            Me.dgvUser.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Return GetListOfUsers
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return GetListOfUsers
        End Try
    End Function

    Private Sub Filter()
        Dim dtaFilter As New DseTrans_MaterialPlanningDashboard_AuthorizeTableAdapters.tbltrans_material_planningTableAdapter
        Dim dtaFilterAll As New DseTrans_MaterialPlanningDashboard_AuthorizeTableAdapters.tbltrans_material_planning1TableAdapter
        Dim tblFilter As New DseTrans_MaterialPlanningDashboard_Authorize.tbltrans_material_planningDataTable
        Dim tblFilterAll As New DseTrans_MaterialPlanningDashboard_Authorize.tbltrans_material_planning1DataTable
        Dim rowFilter As DseTrans_MaterialPlanningDashboard_Authorize.tbltrans_material_planningRow
        Dim rowFilterAll As DseTrans_MaterialPlanningDashboard_Authorize.tbltrans_material_planning1Row
        Dim from As Date = #1/1/1950#
        Dim to_ As Date = #1/1/2200#
        Dim listOfUser As New List(Of String)
        Try
            Me.Cursor = Cursors.WaitCursor
            listOfUser = Me.GetListOfUsers
            If Me.chkFilterDate.Checked Then
                from = Me.dtpFilterFrom.Value.Date
                to_ = Me.dtpFilterTo.Value.Date
                If from > to_ Then
                    MsgBox("Invalid filter date.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            lastRow = -1
            dtaFilter.Connection.ConnectionString = Core_Module.GetConstring
            dtaFilterAll.Connection.ConnectionString = Core_Module.GetConstring
            If Me.TabControl1.SelectedTab.Name = Me.TabPage1.Name Then
                If Me.txtFilterPrimaryMP.Text = "-1" Then
                    dtaFilter.Fill_Auth(tblFilter, Me.AuthLevel, from, to_)
                Else
                    dtaFilter.Fill_AuthWPrimary(tblFilter, Me.AuthLevel, from, to_, Me.txtFilterPrimaryMP.Text)
                End If
                Me.dgvAuth.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.dgvAuth.Rows.Clear()
                For Each rowFilter In tblFilter
                    If listOfUser.Count = 0 Or listOfUser.Contains(rowFilter.created_by) Then
                        Me.dgvRow = Me.dgvAuth.Rows(Me.dgvAuth.Rows.Add)
                        Me.dgvRow.Cells(Me.col1Auth.Index).Value = False
                        Me.dgvRow.Cells(Me.col1Date.Index).Value = Format(rowFilter._date, "dd MMM yyyy")
                        Me.dgvRow.Cells(Me.col1No.Index).Value = Me.dgvAuth.Rows.Count
                        Me.dgvRow.Cells(Me.col1DocNo.Index).Value = rowFilter.doc_no
                        Me.dgvRow.Cells(Me.col1Primary.Index).Value = rowFilter.primary
                    End If
                Next
                Me.dgvAuth.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                If Me.dgvAuth.RowCount > 0 Then
                    Me.dgvAuth.ClearSelection()
                    Me.dgvAuth.Rows(0).Cells(Me.col1DocNo.Index).Selected = True
                End If
            ElseIf Me.TabControl1.SelectedTab.Name = Me.TabPage2.Name Then
                If Me.AuthLevel = 1 Then
                    If Me.txtFilterPrimaryMP.Text = "-1" Then
                        dtaFilter.Fill_CancelLv1(tblFilter, from, to_, Me.objUser.GetPrimary(Me.UserCode))
                    Else
                        dtaFilter.Fill_CancelLv1WPrimaryMP(tblFilter, from, to_, Me.objUser.GetPrimary(Me.UserCode), Me.txtFilterPrimaryMP.Text)
                    End If
                ElseIf Me.AuthLevel = 2 Then
                    If Me.txtFilterPrimaryMP.Text = "-1" Then
                        dtaFilter.Fill_CancelLv2(tblFilter, from, to_, Me.objUser.GetPrimary(Me.UserCode))
                    Else
                        dtaFilter.Fill_CancelLv2WPrimaryMP(tblFilter, from, to_, Me.objUser.GetPrimary(Me.UserCode), Me.txtFilterPrimaryMP.Text)
                    End If
                ElseIf Me.AuthLevel = 3 Then
                    If Me.txtFilterPrimaryMP.Text = "-1" Then
                        dtaFilter.Fill_CancelLv3(tblFilter, from, to_, Me.objUser.GetPrimary(Me.UserCode))
                    Else
                        dtaFilter.Fill_CancelLv3WPrimaryMP(tblFilter, from, to_, Me.objUser.GetPrimary(Me.UserCode), Me.txtFilterPrimaryMP.Text)
                    End If
                End If
                Me.dgvCancel.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.dgvCancel.Rows.Clear()
                For Each rowFilter In tblFilter
                    If listOfUser.Count = 0 Or listOfUser.Contains(rowFilter.created_by) Then
                        Me.dgvRow = Me.dgvCancel.Rows(Me.dgvCancel.Rows.Add)
                        Me.dgvRow.Cells(Me.col2CancelAuth.Index).Value = False
                        Me.dgvRow.Cells(Me.col2Date.Index).Value = Format(rowFilter._date, "dd MMM yyyy")
                        Me.dgvRow.Cells(Me.col2No.Index).Value = Me.dgvCancel.Rows.Count
                        Me.dgvRow.Cells(Me.col2DocumentNo.Index).Value = rowFilter.doc_no
                        Me.dgvRow.Cells(Me.col2Primary.Index).Value = rowFilter.primary
                    End If
                Next
                If Me.dgvCancel.RowCount > 0 Then
                    Me.dgvCancel.ClearSelection()
                    Me.dgvCancel.Rows(0).Cells(Me.col2DocumentNo.Index).Selected = True
                End If
                Me.dgvCancel.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            ElseIf Me.TabControl1.SelectedTab.Name = Me.TabPage3.Name Then
                If Me.txtFilterPrimaryMP.Text = "-1" Then
                    dtaFilterAll.Fill_All(tblFilterAll, from, to_)
                Else
                    dtaFilterAll.Fill_AllWPrimaryMP(tblFilterAll, from, to_, Me.txtFilterPrimaryMP.Text)
                End If
                Me.dgvAll.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                Me.dgvAll.Rows.Clear()
                For Each rowFilterAll In tblFilterAll
                    Dim show As Boolean = False
                    If Me.cboFilterStatus.SelectedIndex = EnumStatus.UNAUTHORIZED_ONLY And rowFilterAll.authorize_level_needed <> rowFilterAll.authorize_level_current Then
                        show = True
                    End If
                    If Me.cboFilterStatus.SelectedIndex = EnumStatus.AUTHORIZED_ONLY And rowFilterAll.authorize_level_needed = rowFilterAll.authorize_level_current Then
                        show = True
                    End If
                    If Me.cboFilterStatus.SelectedIndex = EnumStatus.ALL Then
                        show = True
                    End If
                    If show And (listOfUser.Count = 0 Or listOfUser.Contains(rowFilterAll.created_by)) Then
                        Me.dgvRow = Me.dgvAll.Rows(Me.dgvAll.Rows.Add)
                        Me.dgvRow.Cells(Me.col3Date.Index).Value = Format(rowFilterAll._date, "dd MMM yyyy")
                        Me.dgvRow.Cells(Me.col3DocNo.Index).Value = rowFilterAll.doc_no
                        Me.dgvRow.Cells(Me.col3No.Index).Value = Me.dgvAll.Rows.Count
                        If rowFilterAll.authorize_level_needed = 0 Then
                            If rowFilterAll.is_closed = 1 Then
                                Me.dgvRow.Cells(Me.col3Status.Index).Value = "No authorize needed and closed."
                            Else
                                Me.dgvRow.Cells(Me.col3Status.Index).Value = "No authorize needed."
                            End If
                        Else
                            If rowFilterAll.is_closed = 1 Then
                                Me.dgvRow.Cells(Me.col3Status.Index).Value = "Is Closed."
                            ElseIf rowFilterAll.authorize_level_current = 3 Then
                                Me.dgvRow.Cells(Me.col3Status.Index).Value = "Already authorized lv 3 by " & rowFilterAll.user_auth_lvl_3 & " on " & Format(rowFilterAll.authorize_lv3_time, "dd MMM yyyy HH.mm") & "."
                            ElseIf rowFilterAll.authorize_level_current = 2 Then
                                Me.dgvRow.Cells(Me.col3Status.Index).Value = "Already authorized lv 2 by " & rowFilterAll.user_auth_lvl_2 & " on " & Format(rowFilterAll.authorize_lv2_time, "dd MMM yyyy HH.mm") & "."
                            ElseIf rowFilterAll.authorize_level_current = 1 Then
                                Me.dgvRow.Cells(Me.col3Status.Index).Value = "Already authorized lv 1 by " & rowFilterAll.user_auth_lvl_1 & " on " & Format(rowFilterAll.authorize_lv1_time, "dd MMM yyyy HH.mm") & "."
                            ElseIf rowFilterAll.authorize_level_current = 0 Then
                                Me.dgvRow.Cells(Me.col3Status.Index).Value = "Not yet authorized."
                            End If
                        End If
                        Me.dgvRow.Cells(Me.col3Primary.Index).Value = rowFilterAll.primary
                    End If
                Next
                If Me.dgvAll.RowCount > 0 Then
                    Me.dgvAll.ClearSelection()
                    Me.dgvAll.Rows(0).Cells(Me.col3DocNo.Index).Selected = True
                End If
                Me.dgvAll.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub ChekcUnCheckAll(ByVal vDgv As DataGridView, ByVal vCheck As Boolean, ByVal vColIndex As Integer)
        Try
            For Each dgvRow In vDgv.Rows
                If vCheck Then
                    dgvRow.Cells(vColIndex).Value = True
                Else
                    dgvRow.Cells(vColIndex).Value = False
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Try
            Me.ChekcUnCheckAll(Me.dgvAuth, True, Me.col1Auth.Index)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Try
            Me.ChekcUnCheckAll(Me.dgvAuth, False, Me.col1Auth.Index)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel4_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel4.LinkClicked
        Try
            Me.ChekcUnCheckAll(Me.dgvCancel, True, Me.col2CancelAuth.Index)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LinkLabel3_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel3.LinkClicked
        Try
            Me.ChekcUnCheckAll(Me.dgvCancel, False, Me.col2CancelAuth.Index)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Function GetListOfPrimarMP(ByVal vDgv As DataGridView, ByVal vColCheckBox As Integer, ByVal vColPrimary As Integer) As List(Of Double)
        GetListOfPrimarMP = New List(Of Double)
        Try
            For Each dgvRow In vDgv.Rows
                If dgvRow.Cells(vColCheckBox).Value = Boolean.TrueString Then
                    GetListOfPrimarMP.Add(dgvRow.Cells(vColPrimary).Value)
                End If
            Next
            Return GetListOfPrimarMP
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return GetListOfPrimarMP
        End Try
    End Function

    Private Sub ToolStripLabel2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripLabel2.Click
        Dim listOfPrimary As New List(Of Double)
        Dim pwd As New Core_Module.DialogPassword
        Try
            If Me.TabControl1.SelectedTab.Name = Me.TabPage1.Name Then
                Me.dgvAuth.EndEdit()
                listOfPrimary = Me.GetListOfPrimarMP(Me.dgvAuth, Me.col1Auth.Index, Me.col1Primary.Index)
                If listOfPrimary.Count = 0 Then
                    MsgBox("Select at least 1 material planning to authorize.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
                pwd.txtUserName.Text = Me.UserCode
                pwd.Text = "Authorize MP"
                If pwd.ShowDialog = Windows.Forms.DialogResult.OK Then
                    Me.Cursor = Cursors.WaitCursor
                    Me.Auth(listOfPrimary, True, objUser.GetPrimary(Me.UserCode))
                    If listOfPrimary.Count = 1 And Me.objMP.GetAuthLevelCurrent(listOfPrimary(0)) = 3 Then
                        If MsgBox("Authorize complete, do you want to open transcation?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                            Dim ij As Inventory_Module.FrmTrans_MaterialPlanningDashboard
                            Dim j As Windows.Forms.Form
                            For Each j In Me.ParentForm.MdiChildren
                                ij = TryCast(j, Inventory_Module.FrmTrans_MaterialPlanningDashboard)
                                If IsNothing(ij) = False Then
                                    ij.UserCode = Me.UserCode
                                    ij._PrimaryMP = listOfPrimary(0)
                                    ij._Command = Inventory_Module.FrmTrans_MaterialPlanningDashboard.EnumCommand.VIEW
                                    ij.BringToFront()
                                    ij.Focus()
                                    Exit Sub
                                End If
                            Next
                            ij = New Inventory_Module.FrmTrans_MaterialPlanningDashboard
                            ij.UserCode = Me.UserCode
                            ij.BackColorAll = Core_Module.GetSMESSBackColorAll
                            ij._Command = Inventory_Module.FrmTrans_MaterialPlanningDashboard.EnumCommand.VIEW
                            ij._PrimaryMP = listOfPrimary(0)
                            ij.MdiParent = Me.ParentForm
                            ij.ShowIcon = False
                            ij.Show()
                            Exit Sub
                        End If
                    Else
                        MsgBox("Authorize complete.", MsgBoxStyle.Information, "Information")
                    End If

                    Me.btnFilter.PerformClick()
                    'MsgBox("Authorize complete.", MsgBoxStyle.Information, "Information")
                End If
            ElseIf Me.TabControl1.SelectedTab.Name = Me.TabPage2.Name Then
                Me.dgvCancel.EndEdit()
                listOfPrimary = Me.GetListOfPrimarMP(Me.dgvCancel, Me.col2CancelAuth.Index, Me.col2Primary.Index)
                If listOfPrimary.Count = 0 Then
                    MsgBox("Select at least 1 material planning to cancel authorize.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
                pwd.txtUserName.Text = Me.UserCode
                pwd.Text = "Cancel Authorize MP"
                If pwd.ShowDialog = Windows.Forms.DialogResult.OK Then
                    Me.Cursor = Cursors.WaitCursor
                    Dim list_ As New List(Of String)
                    list_ = Me.Auth(listOfPrimary, False, objUser.GetPrimary(Me.UserCode))
                    Me.btnFilter.PerformClick()
                    If list_.Count > 0 Then
                        Dim listFailed As String = ""
                        For Each s As String In list_
                            If listFailed = "" Then
                                listFailed = s
                            Else
                                listFailed = listFailed & vbNewLine & s
                            End If
                        Next
                        MsgBox("The following material planning failed to cancel authorization, cause the MP is closed." & vbNewLine & listFailed, MsgBoxStyle.Information, "Information")
                    Else
                        MsgBox("Cancel authorize complete.", MsgBoxStyle.Information, "Information")
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Function Auth(ByVal vListOfPrimary As List(Of Double), ByVal vAuth As Boolean, ByVal vPrimaryUserCode As Double) As List(Of String)
        Dim dtaAuth As New DseTrans_MaterialPlanningDashboard_AuthorizeTableAdapters.tbltrans_material_planningTableAdapter
        Dim list As New List(Of String)
        Try
            dtaAuth.Connection.ConnectionString = Core_Module.GetConstring
            For Each primary As Double In vListOfPrimary
                If Me.AuthLevel = 1 Then
                    If vAuth Then
                        dtaAuth.UpdateQueryLv1(1, vPrimaryUserCode, Now, primary)
                        If Me.objUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_MaterialPlanningDashboard_MP_Authorize2") Then
                            dtaAuth.UpdateQueryLv2(2, vPrimaryUserCode, Now, primary)
                            If Me.objUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_MaterialPlanningDashboard_MP_Authorize3") Then
                                dtaAuth.UpdateQueryLv3(3, vPrimaryUserCode, Now, primary)
                            End If
                        End If
                    Else
                        dtaAuth.UpdateQueryLv1(0, -1, #1/1/1950#, primary)
                    End If
                ElseIf Me.AuthLevel = 2 Then
                    If vAuth Then
                        dtaAuth.UpdateQueryLv2(2, vPrimaryUserCode, Now, primary)
                        If Me.objUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_MaterialPlanningDashboard_MP_Authorize3") Then
                            dtaAuth.UpdateQueryLv3(3, vPrimaryUserCode, Now, primary)
                        End If
                    Else
                        dtaAuth.UpdateQueryLv2(1, -1, #1/1/1950#, primary)
                        If Me.objUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_MaterialPlanningDashboard_MP_Authorize1") Then
                            dtaAuth.UpdateQueryLv1(0, -1, #1/1/1950#, primary)
                        End If
                    End If
                ElseIf Me.AuthLevel = 3 Then
                    If vAuth Then
                        dtaAuth.UpdateQueryLv3(3, vPrimaryUserCode, Now, primary)
                    Else
                        If Me.objMP.IsMPClosed(primary) Then
                            list.Add(Me.objMP.GetDocumentNo(primary)) ''masukin list jika gagal cancel auth,
                        Else
                            dtaAuth.UpdateQueryLv3(2, -1, #1/1/1950#, primary)
                            If Me.objUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_MaterialPlanningDashboard_MP_Authorize2") Then
                                dtaAuth.UpdateQueryLv2(1, -1, #1/1/1950#, primary)
                                If Me.objUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_MaterialPlanningDashboard_MP_Authorize1") Then
                                    dtaAuth.UpdateQueryLv1(0, -1, #1/1/1950#, primary)
                                End If
                            End If
                        End If
                    End If
                End If
            Next
            Return list
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return list
        End Try
    End Function

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Sub LoadMP(ByVal vPrimary As Double)
        Try
            If Not vPrimary = -1 Then
                Me.Cursor = Cursors.WaitCursor
                If Me.objMP.CariDetail(vPrimary, vDT, vDT_ItemGIT, vDT_ItemDO, vDT_ItemPOB, vDT_ItemSOB, vDT_ItemPrOB, vDT_ItemSettingLocation, vDT_ItemQty, _
                vDT_ItemMOB, vDT_ItemRMReq, vDT_ItemFollowUp, Me.vDT_ItemSOBDelivery, Me.vDT_ItemPrOBDelivery) = True Then
                    Me.SalinKeField()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.ObjMP
                Me.TxtMP_DocNo.Text = .NoTransaksi
                Me.dtpMP.Value = .TanggalTransaksi
                Me.TxtMP_Remarks.Text = .Remarks
                Me.TxtMP_Primary.Text = .Primary
                Me.dtpMP_RefDate.Value = .RefDate
                If .IsClosed = 1 Then
                    Me.chkClose.Checked = True
                Else
                    Me.chkClose.Checked = False
                End If
                Me.tslMP_CreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslMP_CreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslMP_ModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslMP_ModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
            Me.chkMP_RequiredOnly.Checked = True
            Me.dtpMP_Filter.Value = Me.dtpMP_RefDate.Value
            Me.LoadDGV()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadDGV()
        Dim vRowItemFollowUp As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_followupRow
        Dim Filter As String = ""
        Dim row As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item1Row
        Dim DV As Data.DataRowView
        Dim rowDGV As DataGridViewRow
        Dim Masuk As Boolean = True
        Dim AdaOrderPoint As Boolean = False
        Dim AdaNextOrderPoint As Boolean = False
        Dim DTAPurchReq As New DseTransaksiPRTableAdapters.tbltrans_purchase_req_itemTableAdapter
        Dim TBLPurchReq As New DseTransaksiPR.tbltrans_purchase_req_itemDataTable

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DTAPurchReq.Connection.ConnectionString = Core_Module.GetConstring
            Filter = ""
            Me.vDT.DefaultView.RowFilter = Filter
            Me.vDT.DefaultView.Sort = "Product Code"
            Me.dgvMP.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvMP.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.dgvMP.Rows.Clear()
            Me.dgvMP.ClearSelection()
            For Each DV In vDT.DefaultView
                row = DV.Row
                If Me.chkMP_RequiredOnly.Checked = False Then
                    Masuk = True
                Else
                    If row.qty_required > 0 Then
                        If row.procurement_strategy = clsMasterProduct.EnumProcurementStrategy.Build_To_Stock Then
                            If row.order_point.Date <= Me.dtpMP_Filter.Value.Date Then
                                Masuk = True
                            Else
                                Masuk = False
                            End If
                        Else
                            Masuk = True
                        End If
                    Else
                        Masuk = False
                    End If
                End If

                If Masuk = True Then
                    Me.dgvMP.Rows.Add()
                    rowDGV = Me.dgvMP.Rows(Me.dgvMP.Rows.Count - 1)
                    With rowDGV
                        If row.modify = 1 Then
                            .Cells(Me.ColMP_FollowUp.Index).Value = Boolean.TrueString
                        Else
                            .Cells(Me.ColMP_FollowUp.Index).Value = Boolean.FalseString
                        End If
                        .Cells(Me.ColMP_Desc.Index).Value = row.Description
                        .Cells(Me.ColMP_CodeNo.Index).Value = row.Product_Code
                        .Cells(Me.ColMP_Currency.Index).Value = row.Currency_Code
                        .Cells(Me.ColMP_DefPrecOrderUnit.Index).Value = row.def_prec_order_unit
                        .Cells(Me.ColMP_Disc.Index).Value = objCD.DigitGrouping(row.discount, 2)
                        .Cells(Me.ColMP_NettPrice.Index).Value = objCD.DigitGrouping(row.nett_price, row.def_prec_currency)
                        .Cells(Me.ColMP_No.Index).Value = Me.dgvMP.Rows.Count
                        If row.order_point = #1/1/2200# Then
                            .Cells(Me.ColMP_OrderPoint.Index).Value = "-"
                        Else
                            .Cells(Me.ColMP_OrderPoint.Index).Value = row.order_point
                            AdaOrderPoint = True
                        End If
                        .Cells(Me.ColMP_OrderUnit.Index).Value = row.Order_Unit_Code
                        .Cells(Me.ColMP_PreferedSupplier.Index).Value = row.Supplier_Name
                        .Cells(Me.ColMP_PrimaryOrderUnit.Index).Value = row.primary_order_unit
                        If row.procurement_type = Core_Module.clsMasterProduct.EnumProcurementType.Buy Then
                            .Cells(Me.ColMP_ProcurementType.Index).Value = "Buy"
                        Else
                            .Cells(Me.ColMP_ProcurementType.Index).Value = "Produce"
                        End If
                        .Cells(Me.ColMP_QtyGIT.Index).Value = row.qty_git
                        .Cells(Me.ColMP_QtyInquiry.Index).Value = row.qty_inquiry
                        .Cells(Me.ColMP_QtyLeadTime.Index).Value = row.qty_lead_time
                        .Cells(Me.ColMP_QtyMax.Index).Value = row.qty_max
                        .Cells(Me.ColMP_QtyMin.Index).Value = row.qty_min
                        .Cells(Me.ColMP_QtyOnHand.Index).Value = row.qty_on_hand
                        .Cells(Me.ColMP_QtyOpt.Index).Value = row.qty_optimum
                        .Cells(Me.ColMP_QtyOrder.Index).Value = Me.objCD.DigitGrouping(row.qty_order, row.def_prec_order_unit)
                        .Cells(Me.ColMP_QtyOrderMaximum.Index).Value = row.qty_order_maksimum
                        .Cells(Me.ColMP_QtyOrderMinimum.Index).Value = row.qty_order_minimum
                        .Cells(Me.ColMP_QtyOrderOptimum.Index).Value = row.qty_order_optimum
                        .Cells(Me.ColMP_QtyPendingDO.Index).Value = row.qty_pending_do
                        .Cells(Me.ColMP_QtyPOB.Index).Value = row.qty_pob
                        .Cells(Me.ColMP_QtyRequired.Index).Value = Me.objCD.DigitGrouping(row.qty_required, row.default_prec_unit)
                        .Cells(Me.ColMP_QtySafetyStock.Index).Value = row.qty_safety_stock
                        .Cells(Me.ColMP_QtySOB.Index).Value = row.qty_sob
                        .Cells(Me.ColMP_QtyPrOB.Index).Value = row.qty_prob

                        .Cells(Me.ColMP_QtyOrderMulti.Index).Value = row.qty_order_multiunit
                        .Cells(Me.ColMP_PrimaryUnitMulti.Index).Value = row.primary_unit_multi
                        .Cells(Me.ColMP_DefPrecUnitMulti.Index).Value = row.def_prec_unit_multi
                        .Cells(Me.ColMP_RatioDefault.Index).Value = row.ratio_default
                        .Cells(Me.ColMP_RatioMulti.Index).Value = row.ratio_multi

                        If row.order_strategy = Core_Module.clsMasterProduct.EnumOrderStrategy.Order_By_Maximum_Order_Qty Then
                            .Cells(Me.ColMP_OrderStrategy.Index).Value = "Order By Maximum Order Qty"
                        ElseIf row.order_strategy = Core_Module.clsMasterProduct.EnumOrderStrategy.Order_By_Minimum_Order_Qty Then
                            .Cells(Me.ColMP_OrderStrategy.Index).Value = "Order By Minimum Order Qty"
                        ElseIf row.order_strategy = Core_Module.clsMasterProduct.EnumOrderStrategy.Order_By_Optimum_Order_Qty Then
                            .Cells(Me.ColMP_OrderStrategy.Index).Value = "Order By Optimum Order Qty"
                        ElseIf row.order_strategy = Core_Module.clsMasterProduct.EnumOrderStrategy.Order_To_Max_Qty Then
                            .Cells(Me.ColMP_OrderStrategy.Index).Value = "Order To Max Qty"
                        ElseIf row.order_strategy = Core_Module.clsMasterProduct.EnumOrderStrategy.Order_To_Min_Qty Then
                            .Cells(Me.ColMP_OrderStrategy.Index).Value = "Order To Min Qty"
                        ElseIf row.order_strategy = Core_Module.clsMasterProduct.EnumOrderStrategy.Order_To_Optimum_Qty Then
                            .Cells(Me.ColMP_OrderStrategy.Index).Value = "Order To Optimum Qty"
                        Else
                            .Cells(Me.ColMP_OrderStrategy.Index).Value = "Order To Safety Stock"
                        End If
                        .Cells(Me.ColMP_SubTotal.Index).Value = row.sub_total
                        .Cells(Me.ColMP_Unit.Index).Value = row.Unit_Code
                        .Cells(Me.ColMP_UnitPrice.Index).Value = row.unit_price
                        .Cells(Me.ColMP_DiscAmount.Index).Value = row.disc_amount
                        .Cells(Me.ColMP_DiscFromPerct.Index).Value = row.disc_from_perct

                        .Cells(Me.ColMP_DefPrecUnit.Index).Value = row.default_prec_unit
                        .Cells(Me.ColMP_DefPrecCurrency.Index).Value = row.def_prec_currency
                        .Cells(Me.ColMP_PrimaryCodeNo.Index).Value = row.primary_code_no
                        .Cells(Me.ColMP_PrimarySupplier.Index).Value = row.primary_prefered_supplier
                        .Cells(Me.ColMP_PrimaryCurrency.Index).Value = row.primary_currency

                        .Cells(Me.ColMP_NextOrderPointDelay.Index).Value = row.next_order_point_delay
                        If row.next_order_point = #1/1/2200# Then
                            .Cells(Me.ColMP_NextOrderPoint.Index).Value = "-"
                        Else
                            .Cells(Me.ColMP_NextOrderPoint.Index).Value = row.next_order_point
                            AdaNextOrderPoint = True
                        End If
                        .Cells(Me.ColMP_LeadTimeDays.Index).Value = row.lead_time_days
                        .Cells(Me.ColMP_OrderPointDelay.Index).Value = row.order_point_delay
                        If row.procurement_strategy = Core_Module.clsMasterProduct.EnumProcurementStrategy.Build_To_Order Then
                            .Cells(Me.ColMP_ProcurementStrategy.Index).Value = "Build To Order"
                        Else
                            .Cells(Me.ColMP_ProcurementStrategy.Index).Value = "Build To Stock"
                        End If

                        Me.vDT_ItemFollowUp.DefaultView.RowFilter = "[primary transaksi] = " & row.primary_transaksi.ToString & " and [primary product] = " & row.primary_code_no.ToString
                        If Me.vDT_ItemFollowUp.DefaultView.Count = 0 Then
                            .Cells(Me.ColMP_DetailFollowUp.Index).Value = "-"
                        ElseIf Me.vDT_ItemFollowUp.DefaultView.Count = 1 Then
                            vRowItemFollowUp = Me.vDT_ItemFollowUp.DefaultView(0).Row
                            .Cells(Me.ColMP_DetailFollowUp.Index).Value = vRowItemFollowUp.no_po_mo
                        ElseIf Me.vDT_ItemFollowUp.DefaultView.Count > 1 Then
                            .Cells(Me.ColMP_DetailFollowUp.Index).Value = "Multiple"
                        End If

                    End With
                End If
            Next

            Me.ColMP_NextOrderPoint.DefaultCellStyle.Format = "dd MMM yyyy"
            Me.ColMP_OrderPoint.DefaultCellStyle.Format = "dd MMM yyyy"
            If AdaOrderPoint = True Then
                ColMP_OrderPoint.Visible = True
            Else
                ColMP_OrderPoint.Visible = False
            End If
            If AdaNextOrderPoint = True Then
                ColMP_NextOrderPoint.Visible = True
            Else
                ColMP_NextOrderPoint.Visible = False
            End If
            UpdateKolomMP()

            Me.dgvMP.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            If Me.dgvMP.Rows.Count > 0 Then
                Me.dgvMP.Rows(0).Selected = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")

        End Try
    End Sub

    Private Sub UpdateKolomMP()
        Dim Dummy As String = ""

        Try
            Dummy = Me.ObjSettingUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.MaterialPlanning, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.MaterialPlanning_ShowProcurementStrategy, Me.UserCode)
            If Dummy = "" Then
                Me.ColMP_ProcurementStrategy.Visible = False
            Else
                If Dummy = 1 Then
                    Me.ColMP_ProcurementStrategy.Visible = True
                Else
                    Me.ColMP_ProcurementStrategy.Visible = False
                End If
            End If

            Dummy = Me.ObjSettingUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.MaterialPlanning, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.MaterialPlanning_ProcurementType, Me.UserCode)
            If Dummy = "" Then
                Me.ColMP_ProcurementType.Visible = False
            Else
                If Dummy = 1 Then
                    Me.ColMP_ProcurementType.Visible = True
                Else
                    Me.ColMP_ProcurementType.Visible = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvMP_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvMP.CellMouseClick
        Dim i As New FrmTrans_MaterialPlanningDashboard_ExplainRequired
        Dim rowDGV As DataGridViewRow
        Try
            If e.RowIndex = -1 Then
                Exit Sub
            End If
            If e.ColumnIndex = Me.ColMP_QtyRequired.Index Or e.ColumnIndex = Me.ColMP_OrderPoint.Index Or e.ColumnIndex = Me.ColMP_NextOrderPoint.Index Then
                'tampilkan penjelasan mengenai qty required
                If Me.dgvMP.SelectedCells.Count = 0 Then
                    Exit Sub
                Else
                    rowDGV = Me.dgvMP.Rows(e.RowIndex)
                    For Each row As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item1Row In Me.vDT.Rows
                        If row.primary_code_no = rowDGV.Cells(Me.ColMP_PrimaryCodeNo.Index).Value Then
                            i.ObjCD = Me.objCD
                            i.vDPM = Me.TxtMP_DaysPerMonth.Text
                            i.vRefDate = Me.dtpMP_RefDate.Value
                            i.vRow = row
                            i.ShowDialog()
                            Exit Sub
                        End If
                    Next
                End If
            End If
            If e.ColumnIndex = Me.ColMP_DetailFollowUp.Index And Me.dgvMP.Item(Me.ColMP_DetailFollowUp.Index, e.RowIndex).Value <> "-" Then
                Dim j As New FrmTrans_MaterialPlanningDashboard_FollowUpDetail
                j.vDT_ItemFollowUp = Me.vDT_ItemFollowUp
                j.PrimaryProduct = Me.dgvMP.Item(Me.ColMP_PrimaryCodeNo.Index, e.RowIndex).Value
                j.ShowDialog()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvMP_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvMP.DataError

    End Sub

    Dim lastRow As Integer = -1

    Private Sub dgvAuth_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvAuth.SelectionChanged
        Try
            If Me.dgvAuth.SelectedCells.Count > 0 Then
                If Me.dgvAuth.SelectedCells(0).RowIndex <> lastRow Then
                    If Not IsNothing(Me.dgvAuth.Item(Me.col1Primary.Index, Me.dgvAuth.SelectedCells(0).RowIndex).Value) Then
                        Me.LoadMP(Me.dgvAuth.Item(Me.col1Primary.Index, Me.dgvAuth.SelectedCells(0).RowIndex).Value)
                        lastRow = Me.dgvAuth.SelectedCells(0).RowIndex
                    End If
                End If
            Else
                Me.KosongkanField()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCancel_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvCancel.SelectionChanged
        Try
            If Me.dgvCancel.SelectedCells.Count > 0 Then
                If Me.dgvCancel.SelectedCells(0).RowIndex <> lastRow Then
                    If Not IsNothing(Me.dgvCancel.Item(Me.col2Primary.Index, Me.dgvCancel.SelectedCells(0).RowIndex).Value) Then
                        Me.LoadMP(Me.dgvCancel.Item(Me.col2Primary.Index, Me.dgvCancel.SelectedCells(0).RowIndex).Value)
                        lastRow = Me.dgvCancel.SelectedCells(0).RowIndex
                    End If
                End If
            Else
                Me.KosongkanField()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvAll_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvAll.SelectionChanged
        Try
            If Me.dgvAll.SelectedCells.Count > 0 Then
                If Me.dgvAll.SelectedCells(0).RowIndex <> lastRow Then
                    If Not IsNothing(Me.dgvAll.Item(Me.col3Primary.Index, Me.dgvAll.SelectedCells(0).RowIndex).Value) Then
                        Me.LoadMP(Me.dgvAll.Item(Me.col3Primary.Index, Me.dgvAll.SelectedCells(0).RowIndex).Value)
                        lastRow = Me.dgvAll.SelectedCells(0).RowIndex
                    End If
                End If
            Else
                Me.KosongkanField()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvMP_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvMP.SelectionChanged
        Try
            Me.SetTxtPurchaseDetail()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetTxtPurchaseDetail()
        Dim PrimaryProduct As Double = -1
        Dim rowDGV As DataGridViewRow
        Dim DummyInt As Integer = 0

        Try
            Me.vHitungQtyMultiUnit = False
            Me.cboMPOrderStrategy.Enabled = False
            If Me.dgvMP.SelectedCells.Count > 0 Then
                rowDGV = Me.dgvMP.Rows(Me.dgvMP.SelectedCells(0).RowIndex)
                For Each rowItem As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item1Row In Me.vDT.Rows
                    If rowItem.primary_code_no = rowDGV.Cells(Me.ColMP_PrimaryCodeNo.Index).Value Then
                        With rowItem
                            'detail qty
                            Me.TxtMP_QtyMOB.Text = Me.objCD.DigitGrouping(.qty_mob, .default_prec_unit)
                            Me.TxtMP_QtyRMReq.Text = Me.objCD.DigitGrouping(.qty_rm_req, .default_prec_unit)
                            Me.TxtMP_QtyGIT.Text = Me.objCD.DigitGrouping(.qty_git, .default_prec_unit)
                            'Me.TxtMP_QtyInquiry.Text = Me.objCD.DigitGrouping(.qty_inquiry, .default_prec_unit)
                            Me.TxtMP_QtyOptimumStock.Text = Me.objCD.DigitGrouping(.qty_optimum, .default_prec_unit)
                            Me.TxtMP_QtyOnHand.Text = Me.objCD.DigitGrouping(.qty_on_hand, .default_prec_unit)
                            Me.TxtMP_QtyFree.Text = Me.objCD.DigitGrouping(.qty_on_hand + .qty_pob + .qty_git + .qty_mob - .qty_sob - .qty_pending_do - .qty_prob - .qty_rm_req, .default_prec_unit) '-r8
                            'Me.TxtMP_QtyOrder.Text = Me.ObjCD.DigitGrouping(.qty_order_multiunit, .def_prec_unit_multi)
                            Me.TxtMP_QtyOrder.Text = Me.objCD.DigitGrouping(Math.Round(CDec((.qty_order * .ratio_multi) / .ratio_default), _
                            CInt(.def_prec_unit_multi), MidpointRounding.AwayFromZero), .def_prec_unit_multi)
                            Me.TxtMP_QtyPendingDO.Text = Me.objCD.DigitGrouping(.qty_pending_do, .default_prec_unit)
                            Me.TxtMP_QtyPOB.Text = Me.objCD.DigitGrouping(.qty_pob, .default_prec_unit)
                            Me.TxtMP_QtyRequired.Text = Me.objCD.DigitGrouping(.qty_required, .default_prec_unit)
                            Me.TxtMP_QtyRequired2.Text = Me.objCD.DigitGrouping(.qty_required, .default_prec_unit)
                            Me.TxtMP_QtySafetyStock.Text = Me.objCD.DigitGrouping(.qty_safety_stock, .default_prec_unit)
                            Me.TxtMP_QtySOB.Text = Me.objCD.DigitGrouping(.qty_sob, .default_prec_unit)
                            Me.TxtMP_QtyPrOB.Text = Me.objCD.DigitGrouping(.qty_prob, .default_prec_unit)
                            Me.TxtMP_AvgPerMonth.Text = Me.objCD.DigitGrouping(.qty_avg_montly_demand, .default_prec_unit)

                            If .order_strategy = Core_Module.clsMasterProduct.EnumOrderStrategy.Order_By_Maximum_Order_Qty Then
                                Me.cboMPOrderStrategy.Text = "Order By Maximum Order Qty"
                            ElseIf .order_strategy = Core_Module.clsMasterProduct.EnumOrderStrategy.Order_By_Minimum_Order_Qty Then
                                Me.cboMPOrderStrategy.Text = "Order By Minimum Order Qty"
                            ElseIf .order_strategy = Core_Module.clsMasterProduct.EnumOrderStrategy.Order_By_Optimum_Order_Qty Then
                                Me.cboMPOrderStrategy.Text = "Order By Optimum Order Qty"
                            ElseIf .order_strategy = Core_Module.clsMasterProduct.EnumOrderStrategy.Order_To_Max_Qty Then
                                Me.cboMPOrderStrategy.Text = "Order To Max Qty"
                            ElseIf .order_strategy = Core_Module.clsMasterProduct.EnumOrderStrategy.Order_To_Min_Qty Then
                                Me.cboMPOrderStrategy.Text = "Order To Min Qty"
                            ElseIf .order_strategy = Core_Module.clsMasterProduct.EnumOrderStrategy.Order_To_Optimum_Qty Then
                                Me.cboMPOrderStrategy.Text = "Order To Optimum Qty"
                            Else
                                Me.cboMPOrderStrategy.Text = "Order To Safety Stock"
                            End If

                            'detail purchase
                            Me.TxtMP_PrimarySupplier.Text = .primary_prefered_supplier
                            Me.TxtMP_PreferedSupplier.Text = Me.ObjSupplier.GetSupplierCode(.primary_prefered_supplier)
                            Me.TxtMP_PreferedSupplierDesc.Text = Me.ObjSupplier.GetSupplierName(.primary_prefered_supplier)
                            Me.TxtMP_PrimaryCurrency.Text = .primary_currency
                            Me.TxtMP_CurrencyCode.Text = Me.ObjCurrency.GetCurrencyCode(.primary_currency)
                            Me.TxtMP_UnitPrice.Text = Me.objCD.DigitGrouping(.unit_price, .def_prec_currency)
                            Me.TxtMP_Discount.Text = Me.objCD.DigitGrouping(.discount, 2)
                            Me.TxtMP_DiscFromPerct.Text = .disc_from_perct
                            Me.TxtMP_DiscAmount.Text = Me.objCD.DigitGrouping(.disc_amount, .def_prec_currency)
                            Me.TxtMP_NettPrice.Text = Me.objCD.DigitGrouping(.nett_price, .def_prec_currency)
                            Me.TxtMP_SubTotal.Text = Me.objCD.DigitGrouping(.sub_total, .def_prec_currency)

                            Me.LblMP_Unit1.Text = .Unit_Code
                            Me.llbMP_Ratio.Text = "n/a"

                            If .primary_unit_multi <> .primary_unit Then
                                Me.llbMP_Ratio.Text = Me.objCD.DigitGrouping(.ratio_multi, .def_prec_unit_multi) & " " & .Unit_Code_Multi & _
                                " = " & Me.objCD.DigitGrouping(.ratio_default, .def_prec_order_unit) & " " & .Unit_Code + _
                                " --> Req. Qty = " + Me.objCD.DigitGrouping((.qty_required * .ratio_multi) / .ratio_default, .def_prec_unit_multi) + " " + .Unit_Code_Multi
                            End If

                            Me.ObjProd.IsiCBOMultiUnit(Me.cboMP_OrderUnit, .primary_code_no, Me.DTAProduct_MU, Me.TBLProduct_MU, Me.dtaMultiUnit, Me.tblMultiUnit, .primary_unit_multi)

                            If .primary_unit_multi <> -1 Then
                                If Me.ApakahUnitMultiTepat(.primary_unit_multi) = False Then
                                    'tambahkan baris baru
                                    rowMultiUnit = tblMultiUnit.Newtblmaster_product_item_multiunit2Row
                                    rowMultiUnit.Primary = -2
                                    rowMultiUnit.Primary_Product = .primary_code_no
                                    rowMultiUnit.Primary_Unit = .primary_unit_multi
                                    rowMultiUnit.Unit_Code = .Unit_Code_Multi
                                    rowMultiUnit.Qty_Default = 1
                                    rowMultiUnit.Qty_Rasio = 1
                                    rowMultiUnit.Default_Precision = .def_prec_unit_multi

                                    tblMultiUnit.Rows.Add(rowMultiUnit)

                                    Me.cboMP_OrderUnit.SelectedIndex = -1

                                    If tblMultiUnit.DefaultView.Count > 0 Then
                                        'pilih item pertama, jika primary unit default = -1
                                        If .primary_unit_multi = -1 Then
                                            cboMP_OrderUnit.SelectedIndex = 0
                                        Else
                                            'pilih item sesuai primary unit multi yang diberikan
                                            cboMP_OrderUnit.SelectedIndex = 0
                                            For Each DV As Data.DataRowView In tblMultiUnit.DefaultView
                                                rowMultiUnit = DV.Row
                                                If rowMultiUnit.Primary_Unit = .primary_unit_multi Then
                                                    cboMP_OrderUnit.SelectedIndex = DummyInt
                                                    Exit For
                                                Else
                                                    DummyInt += 1
                                                End If
                                            Next
                                        End If
                                    End If
                                End If
                            End If
                        End With

                        Exit Sub
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.vHitungQtyMultiUnit = True
        End Try
    End Sub

    Private Function ApakahUnitMultiTepat(ByVal vPrimaryUnitMulti As Double) As Boolean
        Try
            Me.rowMultiUnit = Me.tblMultiUnit.DefaultView(Me.cboMP_OrderUnit.SelectedIndex).Row
            If Me.rowMultiUnit.Primary_Unit = vPrimaryUnitMulti Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub LlbMP_UnitPrice_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbMP_UnitPrice.LinkClicked
        Dim ij As New FrmTransaksiPO_HistoricalPrice
        Dim PrimaryCodeNo As Double = -1
        Dim rowDGV As DataGridViewRow
        Try
            If Me.dgvMP.SelectedCells.Count > 0 Then
                rowDGV = Me.dgvMP.Rows(Me.dgvMP.SelectedCells(0).RowIndex)
                PrimaryCodeNo = rowDGV.Cells(Me.ColMP_PrimaryCodeNo.Index).Value
                Me.ObjPO.ExplainHistoricalPrice(PrimaryCodeNo, Me.TBLHistoricalPrice, ij.dgvPrice, Me.objCD)
                ij.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub TabControl2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl2.SelectedIndexChanged
        Try
            If Me.TabControl2.SelectedIndex = 1 Then
                If Me.objUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_MaterialPlanningDashboard_LihatDetailPembelian") = False Then
                    Me.TabControl2.SelectedIndex = 0
                    MsgBox("You don't have authorization form this information.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
                Me.SetTxtPurchaseDetail()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            ModCekData.KosongkanControl(Me.gboMP.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.SplitContainer2.Panel1.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.SplitContainer2.Panel2.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.SplitContainer3.Panel1.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.SplitContainer3.Panel2.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.SplitContainer5.Panel1.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.SplitContainer5.Panel2.Controls, True, True, True)
            Me.TxtMP_Primary.Text = "-1"
            Me.TxtMP_PrimaryCurrency.Text = "-1"
            Me.TxtMP_PrimaryOrderUnit.Text = "-1"
            Me.TxtMP_PrimarySupplier.Text = "-1"
            Me.TxtMP_DaysPerMonth.Text = "0"
            Me.tslMP_CreatedBy.Text = "Created By: "
            Me.tslMP_CreatedDate.Text = "Created Date: "
            Me.tslMP_ModifiedBy.Text = "Modified By: "
            Me.tslMP_ModifiedDate.Text = "Modified Date: "
            Me.cboMPOrderStrategy.SelectedIndex = 0
            Me.TxtMP_QtyGIT.Text = 0
            Me.TxtMP_QtyOptimumStock.Text = 0
            Me.TxtMP_QtyOnHand.Text = 0
            Me.TxtMP_QtyFree.Text = 0
            Me.TxtMP_QtyOrder.Text = 0
            Me.TxtMP_QtyPendingDO.Text = 0
            Me.TxtMP_QtyPOB.Text = 0
            Me.TxtMP_QtyRequired.Text = 0
            Me.TxtMP_QtyRequired2.Text = 0
            Me.TxtMP_QtySafetyStock.Text = 0
            Me.TxtMP_QtySOB.Text = 0
            Me.TxtMP_AvgPerMonth.Text = 0
            Me.llbMP_Ratio.Text = "n/a"
            Me.dgvMP.Rows.Clear()
            Me.vDT.Rows.Clear()
            Me.vDT_ItemDO.Rows.Clear()
            Me.vDT_ItemGIT.Rows.Clear()
            Me.vDT_ItemPOB.Rows.Clear()
            Me.vDT_ItemSOB.Rows.Clear()
            Me.vDT_ItemSettingLocation.Rows.Clear()
            Me.vDT_ItemQty.Rows.Clear()
            Me.vDT_ItemFollowUp.Rows.Clear()
            Me.dtpMP_RefDate.Value = Today
            Me.dtpMP_Filter.Value = Today
            Me.SetTxtPurchaseDetail()
            Me.chkClose.Checked = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvAuth_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvAuth.Sorted
        Try
            Me.objCD.SetNumber(Me.dgvAuth)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvCancel_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvCancel.Sorted
        Try
            Me.objCD.SetNumber(Me.dgvCancel)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvAll_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvAll.Sorted
        Try
            Me.objCD.SetNumber(Me.dgvAll)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvMP_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvMP.Sorted
        Try
            Me.objCD.SetNumber(Me.dgvMP)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class