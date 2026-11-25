Option Explicit On
Option Strict Off

Public Class FrmTrans_ManufacturingOrder_DetailRM
    Public PrimaryRM As Double = -1
    Public CodeNoRM, DescRM, IdentifierJO, CodeNoJO, DescJO, JONO As String
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public TblItemRM As New DseTrans_ManufacturingOrder.tbltrans_mo_item_rmDataTable
    Public TblItemRMAdd As New DseTrans_ManufacturingOrder.tbltrans_mo_item_rm_addDataTable
    Public TblItemWO As New DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable
    Public TblItemBom As New DseTrans_ManufacturingOrder.tbltrans_mo_item_bomDataTable
    Public ReadOnly_ As Boolean = False
    Public PrimaryMO As Double = -1
    Public RefDate As Date = Today
    Public AdaPerubahan As Boolean = False

    Dim ObjMO As New ClsTrans_ManufacturingOrder
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim Buffer As String = ""

    Private Sub FrmTrans_ManufacturingOrder_DetailRM_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.DGVRM.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.Text = Me.Text + " - " + Me.CodeNoRM
            Me.LoadRMDetail(Me.IdentifierJO, Me.TblItemWO, Me.TblItemRM, Me.DGVRM, Me, Me.PrimaryRM)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub LoadRMDetail(ByVal vIdentifierJO As String, ByVal vTblItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable, _
    ByVal vTblItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmDataTable, ByVal vDgvRM As DataGridView, ByVal vFrm As FrmTrans_ManufacturingOrder_DetailRM, _
    ByVal PrimaryRM As Double)
        Dim rowItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woRow
        Dim rowItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmRow
        Dim rowD As DataGridViewRow
        Dim ObjProses As New Core_Module.ClsMasterProses
        Dim vDecProd_Item As Integer = 0

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            vDgvRM.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDgvRM.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            vDgvRM.AllowUserToResizeColumns = False
            vDgvRM.AllowUserToResizeRows = False

            If vDgvRM.Rows.Count > 0 Then
                vDgvRM.Rows.Clear()
            End If

            vTblItemRM.DefaultView.RowFilter = "[Identifier JO] = '" & vIdentifierJO & "' AND [Primary Product Code RM] = " + Me.PrimaryRM.ToString

            vTblItemRM.DefaultView.Sort = "[Primary Product Code RM] ASC"
            If vTblItemRM.DefaultView.Count > 0 Then
                vDgvRM.AllowUserToAddRows = False
                For Each dv As Data.DataRowView In vTblItemRM.DefaultView
                    rowItemRM = dv.Row
                    With rowItemRM
                        vDgvRM.Rows.Add()
                        rowD = vDgvRM.Rows(vDgvRM.RowCount - 1)
                        rowD.Cells(vFrm.ColRM_No.Index).Value = vDgvRM.Rows.Count.ToString
                        vDecProd_Item = ObjProduct.GetAngkaDecimal(.Primary_Product_Code_RM)
                        rowD.Cells(vFrm.ColRM_PrimaryCodeNo.Index).Value = .Primary_Product_Code_RM
                        vTblItemWO.DefaultView.RowFilter = "[Identifier JO] = '" & .Identifier_JO.ToString & "' AND [Identifier] = '" & .Identifier_WO & "'"
                        If vTblItemWO.DefaultView.Count > 0 Then
                            rowItemWO = vTblItemWO.DefaultView(0).Row
                            rowD.Cells(vFrm.ColRM_PrimaryProcess.Index).Value = rowItemWO.Primary_Process
                            rowD.Cells(vFrm.ColRM_Process.Index).Value = ObjProses.GetCode(rowItemWO.Primary_Process)
                        End If
                        rowD.Cells(vFrm.ColRM_PrimaryUnit.Index).Value = ObjProduct.GetPrimaryUnit(.Primary_Product_Code_RM)
                        rowD.Cells(vFrm.ColRM_WONumber.Index).Value = ObjMO.GetWONumber_InTbl(.Identifier_WO, vTblItemWO, .Identifier_JO)
                        rowD.Cells(vFrm.ColRM_CodeNo.Index).Value = ObjProduct.GetProductCode(.Primary_Product_Code_RM)
                        rowD.Cells(vFrm.ColRM_Desc.Index).Value = ObjProduct.GetDescription(.Primary_Product_Code_RM)
                        rowD.Cells(vFrm.ColRM_Unit.Index).Value = ObjProduct.GetUnitCode(.Primary_Product_Code_RM)
                        rowD.Cells(vFrm.ColRM_PrimaryStandByLocation.Index).Value = .Primary_StandBy_Location
                        rowD.Cells(vFrm.ColRM_PrimaryPickUpLocation.Index).Value = .Primary_PickUp_Location
                        rowD.Cells(vFrm.ColRM_StandByLocation.Index).Value = ""
                        rowD.Cells(vFrm.ColRM_PickUpLocation.Index).Value = ""
                        rowD.Cells(vFrm.ColRM_QtyAdditional.Index).Value = Me.ObjCD.DigitGrouping(.Additional_Qty, vDecProd_Item)
                        rowD.Cells(vFrm.ColRM_QtyAllocated.Index).Value = Me.ObjCD.DigitGrouping(.Allocated_Qty, vDecProd_Item)
                        rowD.Cells(vFrm.ColRM_QtyBalance.Index).Value = Me.ObjCD.DigitGrouping(.Balance_Qty, vDecProd_Item)
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

    Private Sub DgvRM_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DGVRM.Sorted
        Try
            Me.ObjCD.SetNumber(Me.DGVRM)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvRM_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DGVRM.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColRM_Remarks.Index Then
                Me.Buffer = Me.DGVRM.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DGVRM_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGVRM.CellClick
        Try
            If e.RowIndex = -1 Then
                Exit Sub
            End If
            If e.ColumnIndex = Me.ColRM_QtyRequired.Index Then
                'If CDec(Me.DGVRM.Item(Me.ColRM_QtyRequired.Index, e.RowIndex).Value) > 0 Then
                Dim ij As New FrmTrans_ManufacturingOrder_Exp_QtyReq
                Dim rowDGV As DataGridViewRow

                ij.BackColorAll = Me.BackColorAll
                ij.TblItemBom = Me.TblItemBom
                ij.TblItemRM = Me.TblItemRM
                ij.TblItemWO = Me.TblItemWO
                ij.IdentifierWO = Me.DGVRM.Item(Me.ColRM_IdentifierWO.Index, e.RowIndex).Value
                ij.IdentifierJO = Me.DGVRM.Item(Me.ColRM_IdentifierItem.Index, e.RowIndex).Value
                ij.PrimaryRM = Me.DGVRM.Item(Me.ColRM_PrimaryCodeNo.Index, e.RowIndex).Value
                ij.CodeNo = Me.CodeNoJO
                ij.Desc = Me.DescJO
                ij.JONO = Me.JONO

                ij.ShowDialog()
                'End If
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
                If Me.ReadOnly_ = True Then
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
                ij.RefDate = Me.RefDate
                ij.PrimaryMO = Me.PrimaryMO
                ij.ReadOnly_ = Me.ReadOnly_

                If ij.ShowDialog = Windows.Forms.DialogResult.OK Then
                    Dim RowDGV As DataGridViewRow
                    Dim RowItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmRow

                    For Each RowItemRM In Me.TblItemRM.Rows
                        If RowItemRM.Identifier_JO = ij.vIdentifierJO And RowItemRM.Identifier_WO = ij.vIdentifireWO And _
                        RowItemRM.Primary_Product_Code_RM = ij.Primary_Product_Code_RM Then
                            RowDGV = Me.DGVRM.Rows(e.RowIndex)
                            RowDGV.Cells(Me.ColRM_QtyAdditional.Index).Value = Me.ObjCD.DigitGrouping(RowItemRM.Additional_Qty, ij.DefPrec)
                            RowDGV.Cells(Me.ColRM_QtyBalance.Index).Value = Me.ObjCD.DigitGrouping(RowItemRM.Balance_Qty, ij.DefPrec)
                            Exit For
                        End If
                    Next
                    Me.AdaPerubahan = True
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
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class