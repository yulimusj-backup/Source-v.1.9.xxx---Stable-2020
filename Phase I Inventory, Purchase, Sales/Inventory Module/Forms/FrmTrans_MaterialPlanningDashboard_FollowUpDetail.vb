Option Explicit On
Option Strict Off

Public Class FrmTrans_MaterialPlanningDashboard_FollowUpDetail

    Public PrimaryProduct As Double = -1
    Public vDT_ItemFollowUp As New DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_followupDataTable
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public ReadOnly_ As Boolean = False
    Dim vRowItemFollowUp As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_followupRow
    Dim rowDgv As DataGridViewRow
    Dim objCD As New Core_Module.ClsCekData
    Dim objProd As New Core_Module.clsMasterProduct
    Dim objUnit As New Core_Module.ClsMasterUnit

    Private Sub BtnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnClose.Click
        Me.Close()
    End Sub

    Private Sub FrmTrans_MaterialPlanningDashboard_FollowUpDetail_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.DgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.DgvItem.RowHeadersVisible = True
            Me.DgvItem.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect
            Me.DgvItem.MultiSelect = False
            If Me.ReadOnly_ Then
                Me.DgvItem.ReadOnly = True
            Else
                Me.DgvItem.ReadOnly = False
                Me.DgvItem.AllowUserToDeleteRows = True
            End If
            If Me.vDT_ItemFollowUp.Columns.Contains("dgv_index") = True Then
                Me.vDT_ItemFollowUp.Columns.Remove("dgv_index")
            End If
            Me.vDT_ItemFollowUp.Columns.Add("dgv_index", System.Type.GetType("System.Double"))
            Me.vDT_ItemFollowUp.Columns("dgv_index").DefaultValue = -1

            Me.LoadDgv()

            Me.FormBorderStyle = Windows.Forms.FormBorderStyle.Sizable
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadDgv()
        Try
            Me.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.DgvItem.Rows.Clear()
            Me.ColUnitMulti.Visible = False
            Me.ColMultiQty.Visible = False
            Me.ColJONo.Visible = False

            Me.vDT_ItemFollowUp.DefaultView.RowFilter = "[primary product] = " & PrimaryProduct.ToString
            Me.vDT_ItemFollowUp.DefaultView.Sort = "[followup time]"
            If Me.vDT_ItemFollowUp.DefaultView.Count > 0 Then
                For Each DV As Data.DataRowView In Me.vDT_ItemFollowUp.DefaultView
                    vRowItemFollowUp = DV.Row
                    Me.rowDgv = Me.DgvItem.Rows(Me.DgvItem.Rows.Add)
                    With Me.rowDgv
                        .Cells(Me.ColNo.Index).Value = Me.DgvItem.Rows.Count.ToString
                        .Cells(Me.ColDocumentNo.Index).Value = vRowItemFollowUp.no_po_mo
                        .Cells(Me.ColDate.Index).Value = Format(vRowItemFollowUp.tanggal_po_mo, "dd MMM yyyy")
                        If vRowItemFollowUp.tipe = ClsTransaksiMaterialPlanning.EnumActionType.BUY_CREATE_PO Then
                            .Cells(Me.ColType.Index).Value = "Create new PO"
                        ElseIf vRowItemFollowUp.tipe = ClsTransaksiMaterialPlanning.EnumActionType.BUY_ADDTO_PO Then
                            .Cells(Me.ColType.Index).Value = "Add to PO"
                        ElseIf vRowItemFollowUp.tipe = ClsTransaksiMaterialPlanning.EnumActionType.PRODUCE_CREATE_MO Then
                            .Cells(Me.ColType.Index).Value = "Create new MO"
                            Me.ColJONo.Visible = True
                        ElseIf vRowItemFollowUp.tipe = ClsTransaksiMaterialPlanning.EnumActionType.PRODUCE_ADDTO_MO Then
                            .Cells(Me.ColType.Index).Value = "Add to MO"
                            Me.ColJONo.Visible = True
                        End If
                        .Cells(Me.ColJONo.Index).Value = vRowItemFollowUp.no_jo
                        .Cells(Me.ColDefaultQty.Index).Value = Me.objCD.DigitGrouping(vRowItemFollowUp.qty_default, Me.objProd.GetAngkaDecimal(vRowItemFollowUp.primary_product))
                        .Cells(Me.ColUnitDefault.Index).Value = Me.objUnit.GetUnitCode(vRowItemFollowUp.primary_unit_default)
                        .Cells(Me.ColMultiQty.Index).Value = Me.objCD.DigitGrouping(vRowItemFollowUp.qty_multi, Me.objProd.GetAngkaDecimal(vRowItemFollowUp.primary_product, vRowItemFollowUp.primary_unit_multi))
                        .Cells(Me.ColUnitMulti.Index).Value = Me.objUnit.GetUnitCode(vRowItemFollowUp.primary_unit_multi)
                        .Cells(Me.ColUserCode.Index).Value = vRowItemFollowUp.user_code
                        .Cells(Me.ColTimestamp.Index).Value = Format(vRowItemFollowUp.followup_time, "dd MMM yyyy HH.mm")
                        .Cells(Me.ColPrimaryPO_MO.Index).Value = vRowItemFollowUp.primary_po_mo
                        .Cells(Me.ColPrimaryProduct.Index).Value = vRowItemFollowUp.primary_product
                        If Not String.IsNullOrEmpty(vRowItemFollowUp.remarks) Then
                            .Cells(Me.ColRemarks.Index).Value = vRowItemFollowUp.remarks
                        End If
                        vRowItemFollowUp.Item("dgv_index") = Me.DgvItem.Rows.IndexOf(Me.rowDgv)
                    End With
                    If vRowItemFollowUp.primary_unit_default <> vRowItemFollowUp.primary_unit_multi Then
                        Me.ColUnitMulti.Visible = True
                        Me.ColMultiQty.Visible = True
                    End If
                Next
            End If
            Me.DgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItem_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvItem.Sorted
        Try
            Me.objCD.SetNumber(Me.DgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Dim buffer As String = ""

    Private Sub DgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DgvItem.CellBeginEdit
        Try
            If Me.ReadOnly_ Then
                e.Cancel = True
                Exit Sub
            End If
            If e.ColumnIndex = Me.ColRemarks.Index Then
                buffer = Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub DgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvItem.CellEndEdit
        Try
            If e.ColumnIndex = Me.ColRemarks.Index Then
                If IsNothing(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If String.IsNullOrEmpty(Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                        Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim.Replace("'", "").Replace("%", "").Replace("*", "")
                    Else
                        Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = String.Empty
                    End If
                Else
                    Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value = String.Empty
                End If
                Me.vDT_ItemFollowUp.DefaultView.RowFilter = "[dgv_index] = " & e.RowIndex.ToString
                If Me.vDT_ItemFollowUp.DefaultView.Count > 0 Then
                    Dim row As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_followupRow = Me.vDT_ItemFollowUp.DefaultView(0).Row
                    row.remarks = Me.DgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Dim rowToRemove As Integer = -1

    Private Sub DgvItem_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles DgvItem.UserDeletingRow
        Try
            If Me.DgvItem.Item(Me.ColType.Index, e.Row.Index).Value = "Create new PO" Or Me.DgvItem.Item(Me.ColType.Index, e.Row.Index).Value = "Add to PO" Then
                If Me.IsExistInPO(Me.DgvItem.Item(Me.ColPrimaryPO_MO.Index, e.Row.Index).Value, Me.DgvItem.Item(Me.ColPrimaryProduct.Index, e.Row.Index).Value) = True Then
                    MsgBox("You can't remove this item, data item still exist in PO.", MsgBoxStyle.Information, "Information")
                    e.Cancel = True
                    Exit Sub
                End If
            ElseIf Me.DgvItem.Item(Me.ColType.Index, e.Row.Index).Value = "Create new MO" Or Me.DgvItem.Item(Me.ColType.Index, e.Row.Index).Value = "Add to MO" Then
                If Me.IsExistInMO(Me.DgvItem.Item(Me.ColPrimaryPO_MO.Index, e.Row.Index).Value, Me.DgvItem.Item(Me.ColPrimaryProduct.Index, e.Row.Index).Value) = True Then
                    MsgBox("You can't remove this item, data item still exist in MO.", MsgBoxStyle.Information, "Information")
                    e.Cancel = True
                    Exit Sub
                End If
            End If
            If MsgBox("Are you sure to remove this data?", MsgBoxStyle.YesNo, "Confirmation") <> MsgBoxResult.Yes Then
                e.Cancel = True
                Exit Sub
            End If
            Me.rowToRemove = e.Row.Index
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Function IsExistInMO(ByVal vPrimaryMO As Double, ByVal vPrimaryProduct As Double) As Boolean
        Dim dta As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planning_item_followupTableAdapter
        Dim n As New Nullable(Of Long)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            n = dta.ScalarQuery_GetCountInMO(vPrimaryMO, vPrimaryProduct)
            If n.HasValue Then
                If n.Value > 0 Then
                    Return True
                End If
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function
    Function IsExistInPO(ByVal vPrimaryPO As Double, ByVal vPrimaryProduct As Double) As Boolean
        Dim dta As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planning_item_followupTableAdapter
        Dim n As New Nullable(Of Long)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            n = dta.ScalarQuery_GetCountInPO(vPrimaryPO, vPrimaryProduct)
            If n.HasValue Then
                If n.Value > 0 Then
                    Return True
                End If
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Sub DgvItem_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles DgvItem.UserDeletedRow
        Try
            If Me.rowToRemove <> -1 Then
                Me.vDT_ItemFollowUp.DefaultView.RowFilter = "[dgv_index] = " & Me.rowToRemove.ToString
                If Me.vDT_ItemFollowUp.DefaultView.Count > 0 Then
                    Me.vRowItemFollowUp = Me.vDT_ItemFollowUp.DefaultView(0).Row
                    Me.vDT_ItemFollowUp.Rows.Remove(Me.vRowItemFollowUp)
                End If
                Me.rowToRemove = -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

End Class