Option Explicit On
Option Strict Off

Public Class frmMasterProduct_BOM1
    Public TblItemBOM_ As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomDataTable
    Public vTblItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable
    Public PrimaryMO As Double = -1
    Public IdentifierJO As String = String.Empty
    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public vProductCode As String = ""
    Public vDefPrec As Integer = 0
    Public vPrimaryProduct As Double = -1
    Public vQtyTarget As Decimal = 0
    Public vTblItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmDataTable
    Public vPrimaryProductJO As Double = -1
    Public vTblProsesOrder As DseTrans_ManufacturingOrder.tbltrans_mo_processorderDataTable
    Public vTblItemWIP As New DseTrans_ManufacturingOrder.tbltrans_mo_item_wipDataTable
    Dim RowItemBOM_ As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomRow
    Dim RowlItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woRow
    Dim objProduct As New Core_Module.clsMasterProduct
    Dim objProcess As New Core_Module.ClsMasterProses
    Dim objUnit As New Core_Module.ClsMasterUnit
    Dim frmProduct As New Core_Module.frmMasterProduct
    Dim objCD As New Core_Module.ClsCekData
    Dim buffer As Object
    Dim rowToDelete As Integer
    Dim rowDgv As DataGridViewRow
    Dim DV As Data.DataRowView

#Region "Tab1"
    Private Sub btn1FindRM_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn1FindRM.Click
        Dim frmFindProduct As New Core_Module.frmCariMasterProduct
        Try
            frmFindProduct.txtCari.Text = Me.txt1RMCode.Text
            frmFindProduct.TipePencarian = Core_Module.frmCariMasterProduct.enumTipePencarian.Goods_Only
            frmFindProduct.TipeProduct = Core_Module.clsMasterProduct.EnumTipeProduct.Goods
            frmFindProduct.ShowDialog()
            If frmFindProduct.PrimaryProduct > 0 Then
                Me.txt1PrimaryRM.Text = frmFindProduct.PrimaryProduct
                Me.txt1RMCode.Text = Me.objProduct.GetProductCode(frmFindProduct.PrimaryProduct)
                Me.txt1RMDesc.Text = Me.objProduct.GetDescription(frmFindProduct.PrimaryProduct)
                Me.txt1ProcessCode.Focus()
            Else
                Me.txt1PrimaryRM.Text = "-1"
                Me.txt1RMCode.Text = String.Empty
                Me.txt1RMDesc.Text = String.Empty
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txt1RMCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txt1RMCode.KeyDown
        Try
            If e.KeyCode = 13 Then
                Me.txt1RMCode.Text = Me.txt1RMCode.Text.Trim
                If Me.objProduct.ApakahSudahAda(Me.txt1RMCode.Text) Then
                    Me.txt1PrimaryRM.Text = Me.objProduct.GetPrimary(Me.txt1RMCode.Text)
                    Me.txt1RMDesc.Text = Me.objProduct.GetDescription(Me.txt1PrimaryRM.Text)
                    Me.txt1ProcessCode.Focus()
                Else
                    Me.btn1FindRM.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txt1RMCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txt1RMCode.Leave
        Try
            Me.txt1RMCode.Text = Me.txt1RMCode.Text.Trim
            If Not String.IsNullOrEmpty(Me.txt1RMCode.Text) Then
                If Me.objProduct.ApakahSudahAda(Me.txt1RMCode.Text) Then
                    Me.txt1PrimaryRM.Text = Me.objProduct.GetPrimary(Me.txt1RMCode.Text)
                    Me.txt1RMDesc.Text = Me.objProduct.GetDescription(Me.txt1PrimaryRM.Text)
                Else
                    Me.btn1FindRM.PerformClick()
                End If
            Else
                Me.txt1PrimaryRM.Text = "-1"
                Me.txt1RMCode.Text = String.Empty
                Me.txt1RMDesc.Text = String.Empty
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind1Process_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn1FindProcess.Click
        Dim frmFindProcess As New Core_Module.FrmMasterProsesCari
        Try
            frmFindProcess.txtCari.Text = Me.txt1ProcessCode.Text
            frmFindProcess.ShowDialog()
            If frmFindProcess.mPrimary > 0 Then
                Me.txt1PrimaryProcess.Text = frmFindProcess.mPrimary
                Me.txt1ProcessCode.Text = Me.objProcess.GetCode(frmFindProcess.mPrimary)
                Me.txt1ProcessDesc.Text = Me.objProcess.GetDesc(frmFindProcess.mPrimary)
                Me.btn1Filter.Focus()
            Else
                Me.txt1PrimaryProcess.Text = "-1"
                Me.txt1ProcessCode.Text = String.Empty
                Me.txt1ProcessDesc.Text = String.Empty
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txt1ProcessCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txt1ProcessCode.KeyDown
        Try
            If e.KeyCode = 13 Then
                Me.txt1ProcessCode.Text = Me.txt1ProcessCode.Text.Trim
                If Me.objProcess.ApakahSudahAda(Me.txt1ProcessCode.Text) Then
                    Me.txt1PrimaryProcess.Text = Me.objProcess.GetPrimary(Me.txt1ProcessCode.Text)
                    Me.txt1ProcessDesc.Text = Me.objProcess.GetDesc(Me.txt1PrimaryProcess.Text)
                    Me.btn1Filter.Focus()
                Else
                    Me.btn1FindProcess.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txt1ProcessCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txt1ProcessCode.Leave
        Try
            Me.txt1ProcessCode.Text = Me.txt1ProcessCode.Text.Trim
            If Not String.IsNullOrEmpty(Me.txt1ProcessCode.Text) Then
                If Me.objProcess.ApakahSudahAda(Me.txt1ProcessCode.Text) Then
                    Me.txt1PrimaryProcess.Text = Me.objProcess.GetPrimary(Me.txt1ProcessCode.Text)
                    Me.txt1ProcessDesc.Text = Me.objProcess.GetDesc(Me.txt1PrimaryProcess.Text)
                Else
                    Me.btn1FindProcess.PerformClick()
                End If
            Else
                Me.txt1PrimaryProcess.Text = "-1"
                Me.txt1ProcessCode.Text = String.Empty
                Me.txt1ProcessDesc.Text = String.Empty
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btn1Filter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn1Filter.Click
        Dim tbl As New Core_Module.DseMasterProduct_BOM1.DataTable2DataTable
        Dim row As Core_Module.DseMasterProduct_BOM1.DataTable2Row
        Try
            If Me.txt1PrimaryRM.Text = "" Or Me.txt1PrimaryRM.Text = "-1" Then
                MsgBox("Specifiy the raw material code.", MsgBoxStyle.Information, "Information")
                Me.txt1RMCode.Focus()
                Exit Sub
            End If

            If Me.txt1PrimaryProcess.Text = "-1" Or Me.txt1PrimaryProcess.Text = "" Then
                Me.TblItemBOM_.DefaultView.RowFilter = "[Primary MO] = " & Me.PrimaryMO.ToString & " and [Identifier JO]='" & Me.IdentifierJO & "' and [Primary Product RM]='" & Me.txt1PrimaryRM.Text & "'"
            Else
                Me.TblItemBOM_.DefaultView.RowFilter = "[Primary MO] = " & Me.PrimaryMO.ToString & " and [Identifier JO]='" & Me.IdentifierJO & "' and [Primary Product RM]='" & Me.txt1PrimaryRM.Text & "' and [Primary Process] = " & Me.txt1PrimaryProcess.Text
            End If
            For Each Me.DV In TblItemBOM_.DefaultView
                RowItemBOM_ = Me.DV.Row
                row = tbl.NewDataTable2Row
                With RowItemBOM_
                    row.Identifier_WO = .Identifier_WO
                    row.WO_No = Me.GetWO_NoByIdentifier(row.Identifier_WO)
                    row.Qty_RM = .Qty_RM
                    row.__Decimal__RM = ._rm__decimal_
                    row.Primary_Unit_RM = .rm_unit_primary
                    row.Unit_Code_RM = .rm_unit_code
                    row.GP_RM = .Good_Percentage_RM
                    row.Tolerance_RM = .production_tolerance
                    row.Primary_RM = .Primary_Product_RM
                    row.Primary_FG = .Primary_Product
                    row.Product_Code_FG = .WIP_FG_Code
                    row.Product_Desc_FG = .WIP_FG_Desc
                    row.Primary_Process_FG = .Primary_Process
                    row.Process_Code_FG = Me.objProcess.GetCode(row.Primary_Process_FG)
                    row.Qty_Ref_FG = .Qty_Ref_Process
                    row.__Decimal__FG = Me.objProduct.GetAngkaDecimal(row.Primary_FG)
                    row.Primary_Unit_FG = Me.objProduct.GetPrimaryUnit(row.Primary_FG)
                    row.Unit_Code_FG = Me.objUnit.GetUnitCode(row.Primary_Unit_FG)
                    row.GP_FG = .Good_Percentage_Process
                    row.Tolerance_FG = .tolerance_process
                    row.Identifier_BoM_Parent = .Identifier_BOM_Parent
                    row.Identifier_BoM = .Identifier_BOM
                End With
                tbl.Rows.Add(row)
            Next

            Me.dgvItem1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem1.Rows.Clear()
            tbl.DefaultView.RowFilter = ""
            tbl.DefaultView.Sort = "[WO No], [Product Code FG]"
            For Each DV In tbl.DefaultView
                row = DV.Row
                Me.rowDgv = Me.dgvItem1.Rows(Me.dgvItem1.Rows.Add())
                With Me.rowDgv
                    .Cells(Me.ColCodeNo.Index).Value = row.Product_Code_FG
                    .Cells(Me.ColDesc.Index).Value = row.Product_Desc_FG
                    .Cells(Me.ColGPProcess.Index).Value = Me.objCD.DigitGrouping(row.GP_FG, 2)
                    .Cells(Me.ColGPRM.Index).Value = Me.objCD.DigitGrouping(row.GP_RM, 2)
                    .Cells(Me.ColModify.Index).Value = Boolean.FalseString
                    .Cells(Me.ColNo.Index).Value = Me.dgvItem1.Rows.Count
                    .Cells(Me.ColPrimaryCodeNo.Index).Value = row.Primary_FG
                    .Cells(Me.colPrimaryProcess.Index).Value = row.Primary_Process_FG
                    .Cells(Me.colPrimaryUnit.Index).Value = row.Primary_Unit_FG
                    .Cells(Me.colPrimaryUnitRM.Index).Value = row.Primary_Unit_RM
                    .Cells(Me.ColProcess.Index).Value = row.Process_Code_FG
                    .Cells(Me.ColQtyRef.Index).Value = Me.objCD.DigitGrouping(row.Qty_Ref_FG, row.__Decimal__FG)
                    .Cells(Me.ColQtyRM.Index).Value = Me.objCD.DigitGrouping(row.Qty_RM, row.__Decimal__RM)
                    .Cells(Me.colToleranceProcess.Index).Value = Me.objCD.DigitGrouping(row.Tolerance_FG, 2)
                    .Cells(Me.colToleranceRM.Index).Value = Me.objCD.DigitGrouping(row.Tolerance_RM, 2)
                    .Cells(Me.ColUnit.Index).Value = row.Unit_Code_FG
                    .Cells(Me.ColUnitRM.Index).Value = row.Unit_Code_RM
                    .Cells(Me.ColDecimalFG.Index).Value = row.__Decimal__FG
                    .Cells(Me.ColDecimalRM.Index).Value = row.__Decimal__RM
                    .Cells(Me.ColPrimaryCodeNoRM.Index).Value = row.Primary_RM
                    .Cells(Me.ColIdentifierWO.Index).Value = row.Identifier_WO
                    .Cells(Me.ColWO_No.Index).Value = row.WO_No
                    .Cells(Me.ColIdentifierBoM_Parent.Index).Value = row.Identifier_BoM_Parent
                    .Cells(Me.ColIdentifierBoM.Index).Value = row.Identifier_BoM
                End With
            Next
            Me.dgvItem1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem1_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem1.CellBeginEdit
        Try
            If e.ColumnIndex = Me.ColQtyRef.Index Or e.ColumnIndex = Me.ColGPProcess.Index Or e.ColumnIndex = Me.colToleranceProcess.Index Or _
            e.ColumnIndex = Me.ColQtyRM.Index Or e.ColumnIndex = Me.ColGPRM.Index Or e.ColumnIndex = Me.colToleranceRM.Index Or e.ColumnIndex = Me.ColModify.Index Then
                Me.buffer = Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem1_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem1.CellEndEdit
        Try
            If IsNothing(Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                If e.ColumnIndex <> Me.ColModify.IndeterminateValue Then
                    Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim
                End If
                If String.IsNullOrEmpty(Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If Me.buffer = Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value Then
                        Exit Sub
                    End If

                    ''qty ref FG
                    If e.ColumnIndex = Me.ColQtyRef.Index Then
                        If Me.objCD.CekBilanganPositif(Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                        Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value = Me.objCD.DigitGrouping(Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value, Me.dgvItem1.Item(Me.ColDecimalFG.Index, e.RowIndex).Value)
                        If CDec(Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value) <= 0 Then
                            Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                        Me.UpdateQtyReference(Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value, Me.dgvItem1.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, Me.dgvItem1.Item(Me.ColIdentifierWO.Index, e.RowIndex).Value)
                    End If

                    ''GP process
                    If e.ColumnIndex = Me.ColGPProcess.Index Then
                        If Me.objCD.CekBilanganPositif(Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value) = False Or CDec(Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value) > 100 Then
                            Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                        Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value = Me.objCD.DigitGrouping(Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value, 2)
                        If CDec(Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value) <= 0 Then
                            Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                        Me.dgvItem1.Item(Me.colToleranceProcess.Index, e.RowIndex).Value = Me.objCD.DigitGrouping(Me.frmProduct.GetTolerancePerc(Me.dgvItem1.Item(Me.ColGPProcess.Index, e.RowIndex).Value), 2)
                        Me.UpdateGP_AndTolerance_FG(Me.dgvItem1.Item(Me.ColGPProcess.Index, e.RowIndex).Value, Me.dgvItem1.Item(Me.colToleranceProcess.Index, e.RowIndex).Value, Me.dgvItem1.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, Me.dgvItem1.Item(Me.ColIdentifierWO.Index, e.RowIndex).Value)
                    End If

                    ''tolerance process
                    If e.ColumnIndex = Me.colToleranceProcess.Index Then
                        If Me.objCD.CekNonNegatif(Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                        Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value = Me.objCD.DigitGrouping(Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value, 2)
                        Me.dgvItem1.Item(Me.ColGPProcess.Index, e.RowIndex).Value = Me.objCD.DigitGrouping(Me.frmProduct.GetGoodPercentagePerc(Me.dgvItem1.Item(Me.colToleranceProcess.Index, e.RowIndex).Value), 2)
                        If CDec(Me.dgvItem1.Item(Me.ColGPProcess.Index, e.RowIndex).Value) <= 0 Then
                            Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Me.dgvItem1.Item(Me.ColGPProcess.Index, e.RowIndex).Value = Me.objCD.DigitGrouping(Me.frmProduct.GetGoodPercentagePerc(Me.dgvItem1.Item(Me.colToleranceProcess.Index, e.RowIndex).Value), 2)
                            Exit Sub
                        End If
                        Me.UpdateGP_AndTolerance_FG(Me.dgvItem1.Item(Me.ColGPProcess.Index, e.RowIndex).Value, Me.dgvItem1.Item(Me.colToleranceProcess.Index, e.RowIndex).Value, Me.dgvItem1.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, Me.dgvItem1.Item(Me.ColIdentifierWO.Index, e.RowIndex).Value)
                    End If

                    ''qty rm 
                    If e.ColumnIndex = Me.ColQtyRM.Index Then
                        If Me.objCD.CekBilanganPositif(Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                        Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value = Me.objCD.DigitGrouping(Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value, Me.objProduct.GetAngkaDecimal(Me.dgvItem1.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value))
                        If CDec(Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value) <= 0 Then
                            Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                        UpdateQtyRM(Me.dgvItem1.Item(Me.ColIdentifierWO.Index, e.RowIndex).Value, Me.dgvItem1.Item(Me.ColPrimaryCodeNoRM.Index, e.RowIndex).Value, Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value)
                    End If

                    ''gp rm
                    If e.ColumnIndex = Me.ColGPRM.Index Then
                        If Me.objCD.CekBilanganPositif(Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value) = False Or CDec(Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value) > 100 Then
                            Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                        Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value = Me.objCD.DigitGrouping(Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value, 2)
                        If CDec(Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value) <= 0 Then
                            Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                        Me.dgvItem1.Item(Me.colToleranceRM.Index, e.RowIndex).Value = Me.objCD.DigitGrouping(Me.frmProduct.GetTolerancePerc(Me.dgvItem1.Item(Me.ColGPRM.Index, e.RowIndex).Value), 2)
                        Me.UpdateGP_AndTolerance_RM(Me.dgvItem1.Item(Me.ColIdentifierWO.Index, e.RowIndex).Value, Me.dgvItem1.Item(Me.ColPrimaryCodeNoRM.Index, e.RowIndex).Value, Me.dgvItem1.Item(Me.ColGPRM.Index, e.RowIndex).Value.ToString, Me.dgvItem1.Item(Me.colToleranceRM.Index, e.RowIndex).Value)
                    End If

                    ''tolerance rm
                    If e.ColumnIndex = Me.colToleranceRM.Index Then
                        If Me.objCD.CekBilanganPositif(Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                            Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                        If CDec(Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value) < 0 Then
                            Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                        Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value = Me.objCD.DigitGrouping(Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value, 2)
                        Me.dgvItem1.Item(Me.ColGPRM.Index, e.RowIndex).Value = Me.objCD.DigitGrouping(Me.frmProduct.GetGoodPercentagePerc(Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value), 2)
                        Me.UpdateGP_AndTolerance_RM(Me.dgvItem1.Item(Me.ColIdentifierWO.Index, e.RowIndex).Value, Me.dgvItem1.Item(Me.ColPrimaryCodeNoRM.Index, e.RowIndex).Value, Me.dgvItem1.Item(Me.ColGPRM.Index, e.RowIndex).Value.ToString, Me.dgvItem1.Item(Me.colToleranceRM.Index, e.RowIndex).Value)
                    End If

                Else
                    Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                End If
            Else
                Me.dgvItem1.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem1_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvItem1.DataError

    End Sub

    Private Sub dgvItem1_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem1.Sorted
        Try
            Me.objCD.SetNumber(dgvItem1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Function UpdateQtyRM(ByVal vIdentifierWO As String, ByVal vPrimaryRM As Double, ByVal vQtyRm As Decimal) As Boolean
        Try
            Me.TblItemBOM_.DefaultView.RowFilter = "[Identifier WO] = '" & vIdentifierWO & "' and [Primary Product RM] = " & vPrimaryRM.ToString
            If TblItemBOM_.DefaultView.Count > 0 Then
                Me.RowItemBOM_ = Me.TblItemBOM_.DefaultView(0).Row
                Me.RowItemBOM_.Qty_RM = vQtyRm
                Return True
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Function UpdateGP_AndTolerance_RM(ByVal vIdentifierWO As String, ByVal vPrimaryRM As Double, ByVal vGP_RM As Decimal, ByVal vToleranceRM As Decimal) As Boolean
        Try
            Me.TblItemBOM_.DefaultView.RowFilter = "[Identifier WO] = '" & vIdentifierWO & "' and [Primary Product RM] = " & vPrimaryRM.ToString
            If TblItemBOM_.DefaultView.Count > 0 Then
                Me.RowItemBOM_ = Me.TblItemBOM_.DefaultView(0).Row
                Me.RowItemBOM_.Good_Percentage_RM = vGP_RM
                Me.RowItemBOM_.production_tolerance = vToleranceRM
                Return True
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Sub llbCheckAll1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbCheckAll1.LinkClicked
        Try
            Me.dgvItem1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For Each rowDgv In Me.dgvItem1.Rows
                Me.rowDgv.Cells(Me.ColModify.Index).Value = Boolean.TrueString
            Next
            Me.dgvItem1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbUncheckAll1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbUncheckAll1.LinkClicked
        Try
            Me.dgvItem1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For Each rowDgv In Me.dgvItem1.Rows
                Me.rowDgv.Cells(Me.ColModify.Index).Value = Boolean.FalseString
            Next
            Me.dgvItem1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btn1AddQtyRM_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn1AddQtyRM.Click
        Dim frmSetQtyRef As New Core_Module.frmMasterProduct_BOM1_SetQtyReference
        Try
            If Me.GetSelectedCount(Me.dgvItem1, Me.ColModify) = 0 Then
                MsgBox("Select at least one product to modify.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            frmSetQtyRef.Text = "Add Qty RM"
            frmSetQtyRef.lblQtyReference.Text = "Additional qty RM"
            frmSetQtyRef.lblUnit.Text = Me.GetFirstSelectedValue(Me.dgvItem1, Me.ColModify, Me.ColUnitRM)
            frmSetQtyRef.BackColorAll = Me.BackColorAll
            frmSetQtyRef.Decimal_ = Me.GetFirstSelectedValue(Me.dgvItem1, Me.ColModify, Me.ColDecimalRM)
            If frmSetQtyRef.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.Cursor = Cursors.WaitCursor
                For Each Me.rowDgv In Me.dgvItem1.Rows
                    If Me.rowDgv.Cells(Me.ColModify.Index).Value = Boolean.TrueString Then
                        If Me.AddRM(rowDgv.Cells(Me.ColIdentifierWO.Index).Value, rowDgv.Cells(Me.ColPrimaryCodeNoRM.Index).Value, frmSetQtyRef.QtyReference) = True Then
                            Me.rowDgv.Cells(Me.ColQtyRM.Index).Value = Me.objCD.DigitGrouping(CDec(Me.rowDgv.Cells(Me.ColQtyRM.Index).Value) + frmSetQtyRef.QtyReference, Me.rowDgv.Cells(Me.ColDecimalRM.Index).Value)
                        End If
                    End If
                Next
                MsgBox("Process complete.", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Function AddRM(ByVal vIdentifierWO As String, ByVal vPrimaryRM As Double, ByVal vAdditionalRMQty As Decimal) As Boolean
        Try
            Me.TblItemBOM_.DefaultView.RowFilter = "[Identifier WO]='" & vIdentifierWO & "' and [Primary Product RM] = " & vPrimaryRM.ToString
            If TblItemBOM_.DefaultView.Count > 0 Then
                Me.RowItemBOM_ = Me.TblItemBOM_.DefaultView(0).Row
                Me.RowItemBOM_.Qty_RM += vAdditionalRMQty
                Return True
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Sub btn1SetQtyRM_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn1SetQtyRM.Click
        Dim frmSetQtyRef As New Core_Module.frmMasterProduct_BOM1_SetQtyReference
        Try
            If Me.GetSelectedCount(Me.dgvItem1, Me.ColModify) = 0 Then
                MsgBox("Select at least one product to modify.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            frmSetQtyRef.Text = "Set Qty RM"
            frmSetQtyRef.lblQtyReference.Text = "Qty raw material"
            frmSetQtyRef.lblUnit.Text = Me.GetFirstSelectedValue(Me.dgvItem1, Me.ColModify, Me.ColUnitRM)
            frmSetQtyRef.BackColorAll = Me.BackColorAll
            frmSetQtyRef.Decimal_ = Me.GetFirstSelectedValue(Me.dgvItem1, Me.ColModify, Me.ColDecimalRM)
            If frmSetQtyRef.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.Cursor = Cursors.WaitCursor
                For Each Me.rowDgv In Me.dgvItem1.Rows
                    If Me.rowDgv.Cells(Me.ColModify.Index).Value = Boolean.TrueString Then
                        If Me.UpdateQtyRM(rowDgv.Cells(Me.ColIdentifierWO.Index).Value, rowDgv.Cells(Me.ColPrimaryCodeNoRM.Index).Value, frmSetQtyRef.QtyReference) = True Then
                            Me.rowDgv.Cells(Me.ColQtyRM.Index).Value = Me.objCD.DigitGrouping(frmSetQtyRef.QtyReference, Me.rowDgv.Cells(Me.ColDecimalRM.Index).Value)
                        End If
                    End If
                Next
                MsgBox("Process complete.", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub btn1AddQtyRMByPersen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn1AddQtyRMByPersen.Click
        Dim frmSetQtyRef As New Core_Module.frmMasterProduct_BOM1_SetQtyReference
        Try
            If Me.GetSelectedCount(Me.dgvItem1, Me.ColModify) = 0 Then
                MsgBox("Select at least one product to modify.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            frmSetQtyRef.Text = "Add Qty RM by %"
            frmSetQtyRef.lblQtyReference.Text = "Additional % RM"
            frmSetQtyRef.lblUnit.Text = "%"
            frmSetQtyRef.BackColorAll = Me.BackColorAll
            frmSetQtyRef.Decimal_ = Me.GetFirstSelectedValue(Me.dgvItem1, Me.ColModify, Me.ColDecimalRM)
            If frmSetQtyRef.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.Cursor = Cursors.WaitCursor
                Dim newVal As Decimal = 0
                For Each Me.rowDgv In Me.dgvItem1.Rows
                    If Me.rowDgv.Cells(Me.ColModify.Index).Value = Boolean.TrueString Then
                        newVal = Me.objCD.DigitGrouping(Me.rowDgv.Cells(Me.ColQtyRM.Index).Value + (Me.rowDgv.Cells(Me.ColQtyRM.Index).Value * (frmSetQtyRef.QtyReference / 100)), Me.rowDgv.Cells(Me.ColDecimalRM.Index).Value)
                        If Me.UpdateQtyRM(rowDgv.Cells(Me.ColIdentifierWO.Index).Value, rowDgv.Cells(Me.ColPrimaryCodeNoRM.Index).Value, newVal) = True Then
                            Me.rowDgv.Cells(Me.ColQtyRM.Index).Value = Me.objCD.DigitGrouping(newVal, Me.rowDgv.Cells(Me.ColDecimalRM.Index).Value)
                        End If
                    End If
                Next
                MsgBox("Process complete.", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub btn1SetGPProcess_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn1SetGPProcess.Click
        Dim frmSetQtyRef As New Core_Module.frmMasterProduct_BOM1_SetQtyReference
        Try
            If Me.GetSelectedCount(Me.dgvItem1, Me.ColModify) = 0 Then
                MsgBox("Select at least one product to modify.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            frmSetQtyRef.Text = "Set Good % Process"
            frmSetQtyRef.lblQtyReference.Text = "Good percentage"
            frmSetQtyRef.lblUnit.Text = "%"
            frmSetQtyRef.BackColorAll = Me.BackColorAll
            frmSetQtyRef.Decimal_ = 2
            frmSetQtyRef.Size = New System.Drawing.Size(frmSetQtyRef.btnOK.Left + frmSetQtyRef.btnOK.Width + 10, frmSetQtyRef.Height)
            If frmSetQtyRef.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.Cursor = Cursors.WaitCursor
                For Each Me.rowDgv In Me.dgvItem1.Rows
                    If Me.rowDgv.Cells(Me.ColModify.Index).Value = Boolean.TrueString Then
                        If Me.UpdateGP_AndTolerance_FG(frmSetQtyRef.QtyReference, frmProduct.GetTolerancePerc(frmSetQtyRef.QtyReference), Me.rowDgv.Cells(Me.ColPrimaryCodeNo.Index).Value, Me.rowDgv.Cells(Me.ColIdentifierWO.Index).Value) = True Then
                            Me.rowDgv.Cells(Me.ColGPProcess.Index).Value = Me.objCD.DigitGrouping(frmSetQtyRef.QtyReference, 2)
                            Me.rowDgv.Cells(Me.colToleranceProcess.Index).Value = Me.objCD.DigitGrouping(frmProduct.GetTolerancePerc(frmSetQtyRef.QtyReference), 2)
                        End If
                    End If
                Next
                MsgBox("Process complete.", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub btn1SetGPRM_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn1SetGPRM.Click
        Dim frmSetQtyRef As New Core_Module.frmMasterProduct_BOM1_SetQtyReference
        Try
            If Me.GetSelectedCount(Me.dgvItem1, Me.ColModify) = 0 Then
                MsgBox("Select at least one product to modify.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            frmSetQtyRef.Text = "Set Good % RM"
            frmSetQtyRef.lblQtyReference.Text = "Good % RM"
            frmSetQtyRef.lblUnit.Text = "%"
            frmSetQtyRef.BackColorAll = Me.BackColorAll
            frmSetQtyRef.Decimal_ = 2
            frmSetQtyRef.Size = New System.Drawing.Size(frmSetQtyRef.btnOK.Left + frmSetQtyRef.btnOK.Width + 10, frmSetQtyRef.Height)
            If frmSetQtyRef.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.Cursor = Cursors.WaitCursor
                For Each Me.rowDgv In Me.dgvItem1.Rows
                    If Me.rowDgv.Cells(Me.ColModify.Index).Value = Boolean.TrueString Then
                        If Me.UpdateGP_AndTolerance_RM(Me.rowDgv.Cells(Me.ColIdentifierWO.Index).Value, Me.rowDgv.Cells(Me.ColPrimaryCodeNoRM.Index).Value, frmSetQtyRef.QtyReference, frmProduct.GetTolerancePerc(frmSetQtyRef.QtyReference)) = True Then
                            Me.rowDgv.Cells(Me.ColGPRM.Index).Value = Me.objCD.DigitGrouping(frmSetQtyRef.QtyReference, 2)
                            Me.rowDgv.Cells(Me.colToleranceRM.Index).Value = Me.objCD.DigitGrouping(frmProduct.GetTolerancePerc(frmSetQtyRef.QtyReference), 2)
                        End If
                    End If
                Next
                MsgBox("Process complete.", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub btn1Remove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn1Remove.Click
        Try
            If Me.GetSelectedCount(Me.dgvItem1, Me.ColModify) = 0 Then
                MsgBox("Select at least one product to remove RM.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            If MsgBox("Are you sure to remove raw material?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                For i As Integer = Me.dgvItem1.Rows.Count - 1 To 0 Step -1  'Each Me.rowDgv In Me.dgvItem1.Rows
                    Me.rowDgv = Me.dgvItem1.Rows(i)
                    If Me.rowDgv.Cells(Me.ColModify.Index).Value = Boolean.TrueString Then
                        'If Me.DeleteRM(Me.rowDgv.Cells(Me.ColIdentifierWO.Index).Value, Me.rowDgv.Cells(Me.ColPrimaryCodeNoRM.Index).Value) = True Then
                        '    Me.dgvItem1.Rows.RemoveAt(i)
                        'End If
                        If Me.DeleteRMByIdentifierBoM(Me.rowDgv.Cells(Me.ColIdentifierBoM.Index).Value) = True Then
                            Me.dgvItem1.Rows.RemoveAt(i)
                        End If
                    End If
                Next
                Me.objCD.SetNumber(Me.dgvItem1)
                MsgBox("Process complete.", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Function DeleteRMByIdentifierBoM(ByVal vIdentifierBoM As String) As Boolean
        Try
            Me.TblItemBOM_.DefaultView.RowFilter = "[Identifier BOM] = '" & vIdentifierBoM & "' or [Identifier BOM Parent] = '" & vIdentifierBoM & "'"
            If TblItemBOM_.DefaultView.Count > 0 Then
                While TblItemBOM_.DefaultView.Count <> 0
                    Me.RowItemBOM_ = Me.TblItemBOM_.DefaultView(0).Row
                    Me.TblItemBOM_.Rows.Remove(Me.RowItemBOM_)
                End While
                Return True
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Function DeleteRM(ByVal vIdentifierWO As String, ByVal vPrimaryRM As Double) As Boolean
        Try
            Me.TblItemBOM_.DefaultView.RowFilter = "[Identifier WO] = '" & vIdentifierWO & "' and [Primary Product RM] = " & vPrimaryRM.ToString
            If TblItemBOM_.DefaultView.Count > 0 Then
                Me.RowItemBOM_ = Me.TblItemBOM_.DefaultView(0).Row
                Me.TblItemBOM_.Rows.Remove(Me.RowItemBOM_)
                Return True
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Sub btn1Replace_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn1Replace.Click
        Dim frmFindProduct As New Core_Module.frmCariMasterProduct
        Try
            If Me.GetSelectedCount(Me.dgvItem1, Me.ColModify) = 0 Then
                MsgBox("Select at least one product to replace RM.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            Me.Cursor = Cursors.WaitCursor
            frmFindProduct.ShowDialog()
            If frmFindProduct.PrimaryProduct > 0 Then
                If Me.GetFirstSelectedValue(Me.dgvItem1, Me.ColModify, Me.ColPrimaryCodeNoRM) = frmFindProduct.PrimaryProduct Then
                    MsgBox("Select another raw material.", MsgBoxStyle.Information, "Information")
                    Me.txt1RMCode.Focus()
                    Exit Sub
                End If

                ''-----------------------------------------------------------------------------------------
                Dim tblBoM As New Core_Module.DseMasterProduct_BOM1.tblmaster_product_item_bomDataTable
                Dim rowBoM As Core_Module.DseMasterProduct_BOM1.tblmaster_product_item_bomRow
                Dim alreadyExistRM As String = String.Empty
                For Each Me.rowDgv In Me.dgvItem1.Rows
                    If Me.rowDgv.Cells(Me.ColModify.Index).Value = Boolean.TrueString Then
                        Me.TblItemBOM_.DefaultView.RowFilter = "[Identifier BOM Parent] ='" & rowDgv.Cells(Me.ColIdentifierBoM_Parent.Index).Value & "' and [Primary Product RM] = " & frmFindProduct.PrimaryProduct.ToString
                        If Me.TblItemBOM_.DefaultView.Count > 0 Then
                            rowBoM = tblBoM.Newtblmaster_product_item_bomRow
                            rowBoM.primary_product = Me.rowDgv.Cells(Me.ColPrimaryCodeNo.Index).Value
                            rowBoM.primary_product_raw = frmFindProduct.PrimaryProduct
                            rowBoM.Finish_Good_Code = Me.rowDgv.Cells(Me.ColCodeNo.Index).Value
                            rowBoM.Raw_Material_Code = Me.objProduct.GetProductCode(frmFindProduct.PrimaryProduct)
                            tblBoM.Rows.Add(rowBoM)
                        End If
                    End If
                Next
                tblBoM.DefaultView.RowFilter = "[primary product raw] = " & frmFindProduct.PrimaryProduct.ToString
                If tblBoM.DefaultView.Count > 0 Then
                    For Each Me.DV In tblBoM.DefaultView
                        rowBoM = Me.DV.Row
                        If String.IsNullOrEmpty(alreadyExistRM) Then
                            alreadyExistRM = rowBoM.Finish_Good_Code
                        Else
                            alreadyExistRM = alreadyExistRM & vbNewLine & rowBoM.Finish_Good_Code
                        End If
                    Next
                End If
                If Not String.IsNullOrEmpty(alreadyExistRM) Then
                    MsgBox("The following WIP/FG already has " & Me.objProduct.GetProductCode(frmFindProduct.PrimaryProduct) & " as BoM" & vbNewLine & alreadyExistRM, MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                ''-----------------------------------------------------------------------------------------
                Dim listOfInvalid As New List(Of String)
                listOfInvalid = GetInfinitelyBoM2(frmFindProduct.PrimaryProduct)
                If listOfInvalid.Count > 0 Then
                    Dim existRM As String = String.Empty
                    For Each str As String In listOfInvalid
                        If String.IsNullOrEmpty(existRM) Then
                            existRM = str
                        Else
                            existRM = existRM & vbNewLine & str
                        End If
                    Next
                    MsgBox("Adding the following RM will cause infinite recursive, process canceled." & vbNewLine & existRM, MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                If MsgBox("Are you sure to replace RM with " & Me.objProduct.GetProductCode(frmFindProduct.PrimaryProduct) & "?", MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.Yes Then
                    For Each Me.rowDgv In Me.dgvItem1.Rows
                        If Me.rowDgv.Cells(Me.ColModify.Index).Value = Boolean.TrueString Then
                            If Me.UpdatePrimaryRM1(Me.rowDgv.Cells(Me.ColIdentifierBoM.Index).Value, frmFindProduct.PrimaryProduct, Me.rowDgv.Cells(Me.ColQtyRM.Index).Value, Me.rowDgv.Cells(Me.ColGPRM.Index).Value, Me.rowDgv.Cells(Me.ColIdentifierBoM_Parent.Index).Value, Me.rowDgv.Cells(Me.ColPrimaryCodeNo.Index).Value) Then
                            End If
                        End If
                    Next
                    Me.btn1Filter.PerformClick()
                    MsgBox("Process complete.", MsgBoxStyle.Information, "Information")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Function UpdatePrimaryRM(ByVal vIdentifierWO As String, ByVal vPrimaryRM_Old As Double, ByVal vPrimaryRM_New As Double) As Boolean
        Try
            Me.TblItemBOM_.DefaultView.RowFilter = "[Identifier WO] = '" & vIdentifierWO & "' and [Primary Product RM] = " & vPrimaryRM_Old.ToString
            If Me.TblItemBOM_.DefaultView.Count > 0 Then
                Me.RowItemBOM_ = Me.TblItemBOM_.DefaultView(0).Row
                Me.RowItemBOM_.Primary_Product_RM = vPrimaryRM_New
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Function UpdatePrimaryRM1(ByVal vIdentifierBoM As String, ByVal vNewPrimaryRM As Double, ByVal vRMQty As Decimal, ByVal vGPRM As Decimal, ByVal vIdentifierBOM_Parent As String, ByVal vPrimaryParentProduct As Double) As Boolean
        Try
            Me.RemoveBoM(Me.TblItemBOM_, vIdentifierBoM, "Identifier BoM")
            Me.AddNewBoM(vNewPrimaryRM, vRMQty, vGPRM, vIdentifierBOM_Parent, vPrimaryParentProduct, Me.TblItemBOM_)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Sub AddNewBoM(ByVal vNewPrimaryRM As Double, ByVal vRMQty As Decimal, ByVal vGPRM As Decimal, ByVal vIdentifierBOM_Parent As String, ByVal vPrimaryParentProduct As Double, ByVal vTblItemBoM As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomDataTable)
        Dim f As New FrmTrans_ManufacturingOrder
        Dim rowItemBOM As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomRow
        Dim IdentifierBOM_Parent As String = ""
        Dim ObjMO_Item As New ClsTrans_ManufacturingOrder_Item
        Dim dtaProduct As New DseTrans_ManufacturingOrder_BoMToolsTableAdapters.tblmaster_productTableAdapter
        Dim tblProduct As New DseTrans_ManufacturingOrder_BoMTools.tblmaster_productDataTable
        Dim rowProduct As DseTrans_ManufacturingOrder_BoMTools.tblmaster_productRow
        Try
            dtaProduct.Connection.ConnectionString = Core_Module.GetConstring
            dtaProduct.Fill(tblProduct, vPrimaryParentProduct)
            rowItemBOM = vTblItemBoM.Newtbltrans_mo_item_bomRow
            With rowItemBOM
                .Primary = -1
                .Primary_MO = Me.PrimaryMO
                .Identifier_JO = Me.IdentifierJO
                .Identifier_WO = "-1-1"

                .Primary_Product = vPrimaryParentProduct
                If tblProduct.Rows.Count > 0 Then
                    rowProduct = tblProduct(0)
                    .Qty_Ref_Process = rowProduct.Qty_Ref
                    .Good_Percentage_Process = rowProduct.Good_Precentage
                    .Primary_Process = rowProduct.Primary_Proses
                End If

                .Primary_Product_RM = vNewPrimaryRM
                .Qty_RM = vRMQty
                .Good_Percentage_RM = vGPRM
                .Decimal_RM = objProduct.GetAngkaDecimal(.Primary_Product_RM)
                .Remarks = String.Empty

                .Primary_Parent_Product = vPrimaryParentProduct
                .Identifier_BOM = ObjMO_Item.GetNextIdentifierBOM(vTblItemBoM, Me.PrimaryMO)
                IdentifierBOM_Parent = .Identifier_BOM
                .Identifier_BOM_Parent = vIdentifierBOM_Parent

                .Identifier_WO = String.Empty
                ._rm__decimal_ = .Decimal_RM
                .rm_unit_primary = Me.objProduct.GetPrimaryUnit(.Primary_Product_RM)
                .rm_code = Me.objProduct.GetProductCode(.Primary_Product_RM)
                .rm_desc = Me.objProduct.GetDescription(.Primary_Product_RM)
                .rm_unit_code = Me.objUnit.GetUnitCode(.rm_unit_primary)
                .production_tolerance = Math.Round(Me.GetTolerancePerc(.Good_Percentage_RM), 2)
                .tolerance_process = Math.Round(Me.GetTolerancePerc(.Good_Percentage_Process), 2)

            End With
            vTblItemBoM.Rows.Add(rowItemBOM)
            f.IsiTblItemBOM(vNewPrimaryRM, Me.IdentifierJO, vTblItemBoM, Me.vPrimaryProduct, IdentifierBOM_Parent, Me.PrimaryMO)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            f.Dispose()
        End Try
    End Sub

    Public Function GetTolerancePerc(ByVal vGoodPercentage As Decimal) As Decimal
        Try
            If vGoodPercentage > 0 Then
                Return ((100 - vGoodPercentage) / vGoodPercentage) * 100
            End If
            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        End Try
    End Function

    Public Function GetGoodPercentagePerc(ByVal vTolerance As Decimal) As Decimal
        Try
            Return (100 / (100 + vTolerance)) * 100
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        End Try
    End Function

    Function RemoveBoM(ByVal vTblItemBOM_Full As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomDataTable, ByVal vIdentifierBoM As String, ByVal vColumnName As String) As Boolean
        Dim TblItemBOM_Full As New DseTrans_ManufacturingOrder.tbltrans_mo_item_bomDataTable
        Dim RowItemBOM_ As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomRow
        Try
            TblItemBOM_Full = vTblItemBOM_Full.Copy
            TblItemBOM_Full.DefaultView.RowFilter = "[" & vColumnName & "] = '" & vIdentifierBoM & "'"
            If TblItemBOM_Full.DefaultView.Count > 0 Then
                For i As Integer = TblItemBOM_Full.DefaultView.Count - 1 To 0 Step -1
                    RowItemBOM_ = TblItemBOM_Full.DefaultView(i).Row
                    Me.RemoveBoM(vTblItemBOM_Full, RowItemBOM_.Identifier_BOM, "Identifier BOM Parent")
                    Me.RemoveBoM2(RowItemBOM_.Identifier_BOM)
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Sub RemoveBoM2(ByVal vIdentifierBoM As String)
        Try
            Me.TblItemBOM_.DefaultView.RowFilter = "[Identifier BoM] = '" & vIdentifierBoM & "'"
            While Me.TblItemBOM_.DefaultView.Count > 0
                Me.RowItemBOM_ = Me.TblItemBOM_.DefaultView(0).Row
                Me.TblItemBOM_.Rows.Remove(Me.RowItemBOM_)
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
#End Region

#Region "Tab2"

    Private Sub frmMasterProduct_BOM1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim f As New FrmTrans_ManufacturingOrder
        Try
            Me.UpdateTblProsesOrder_BOM()
            f.IsiTblItemWORM(-1, vPrimaryProduct, vQtyTarget, Me.vTblItemWO, Me.vTblItemRM, Me.TblItemBOM_, Me.IdentifierJO, "-", "-", Me.vTblItemWIP, #1/1/1950#, False)
            Me.DeleteTblItemWORM(Me.TblItemBOM_, Me.vTblItemWO, Me.vTblItemRM)
            f.IsiWONO_N_Urutan(IdentifierJO, Me.vTblProsesOrder, Me.vTblItemWO)
            f.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UpdateTblProsesOrder_BOM()
        Dim rowItemBOM As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomRow
        Dim rowProsesOrder As DseTrans_ManufacturingOrder.tbltrans_mo_processorderRow
        Dim urutanMax As Integer = 0
        Dim ListPrimaryProses As New List(Of Double)
        Dim PrimaryProses As Double = -1

        Try
            TblItemBOM_.DefaultView.RowFilter = "[identifier jo] = '" & IdentifierJO & "'"
            If TblItemBOM_.DefaultView.Count > 0 Then
                For Each dv As Data.DataRowView In TblItemBOM_.DefaultView
                    rowItemBOM = dv.Row
                    If ListPrimaryProses.Contains(rowItemBOM.Primary_Process) = False Then
                        ListPrimaryProses.Add(rowItemBOM.Primary_Process)
                    End If
                Next
            End If
            vTblProsesOrder.DefaultView.RowFilter = "[identifier jo] = '" & IdentifierJO.ToString & "'"
            vTblProsesOrder.DefaultView.Sort = "[urutan] ASC"
            If vTblProsesOrder.DefaultView.Count > 0 Then
                For Each dv As Data.DataRowView In vTblProsesOrder.DefaultView
                    rowProsesOrder = dv.Row
                    PrimaryProses = rowProsesOrder.primary_process
                    urutanMax = rowProsesOrder.urutan
                    If ListPrimaryProses.Contains(rowProsesOrder.primary_process) = False Then
                        vTblProsesOrder.Rows.Remove(rowProsesOrder)
                    End If
                    ListPrimaryProses.Remove(PrimaryProses)
                Next
                If ListPrimaryProses.Count > 0 Then
                    For Each primary_proses As Double In ListPrimaryProses
                        rowProsesOrder = vTblProsesOrder.Newtbltrans_mo_processorderRow
                        With rowProsesOrder
                            .primary = -1
                            .primary_mo = CDbl(PrimaryMO)
                            .primary_process = primary_proses
                            .primary_product = vPrimaryProductJO
                            .identifier_jo = IdentifierJO
                            urutanMax += 1
                            .urutan = urutanMax
                        End With
                        vTblProsesOrder.Rows.Add(rowProsesOrder)
                    Next
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DeleteTblItemWORM(ByVal vTblItemBom_2 As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomDataTable, _
    ByVal vTblItemWO_2 As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable, _
    ByVal vTblItemRM_2 As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmDataTable)
        Dim rowItemWO_dum As DseTrans_ManufacturingOrder.tbltrans_mo_item_woRow
        Dim rowItemRM_dum As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmRow

        Try
            vTblItemWO_2.DefaultView.RowFilter = "[Identifier JO] = '" & IdentifierJO & "'"
            If vTblItemWO_2.DefaultView.Count > 0 Then
                For Each dv As Data.DataRowView In vTblItemWO_2.DefaultView
                    rowItemWO_dum = dv.Row
                    vTblItemBom_2.DefaultView.RowFilter = "[Identifier WO] = '" & rowItemWO_dum.Identifier & "'"
                    If vTblItemBom_2.DefaultView.Count = 0 Then
                        vTblItemWO_2.Rows.Remove(rowItemWO_dum)
                    End If
                Next
            End If

            vTblItemRM_2.DefaultView.RowFilter = "[Identifier JO] = '" & IdentifierJO & "'"
            If vTblItemRM_2.DefaultView.Count > 0 Then
                For Each dv_2 As Data.DataRowView In vTblItemRM_2.DefaultView
                    rowItemRM_dum = dv_2.Row
                    vTblItemBom_2.DefaultView.RowFilter = "[Identifier WO] = '" & rowItemRM_dum.Identifier_WO.ToString _
                    & "' AND [Primary Product RM] = " & rowItemRM_dum.Primary_Product_Code_RM
                    If vTblItemBom_2.DefaultView.Count = 0 Then
                        vTblItemRM_2.Rows.Remove(rowItemRM_dum)
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub frmMasterProduct_BOM1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.dgvItem1.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvItem2.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvProcess.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            ''isi code no dan desc WIP FG
            Dim minPrimary As Double = Double.MaxValue
            Dim maxPrimary As Double = Double.MinValue
            For Each Me.RowItemBOM_ In Me.TblItemBOM_
                If minPrimary > Me.RowItemBOM_.Primary_Product Then
                    minPrimary = Me.RowItemBOM_.Primary_Product
                End If
                If maxPrimary < Me.RowItemBOM_.Primary_Product Then
                    maxPrimary = Me.RowItemBOM_.Primary_Product
                End If
            Next
            Dim dtaProduct As New DseTrans_ManufacturingOrderTableAdapters.tblmaster_productTableAdapter
            Dim tblProduct As New DseTrans_ManufacturingOrder.tblmaster_productDataTable
            Dim rowProduct As DseTrans_ManufacturingOrder.tblmaster_productRow
            dtaProduct.Connection.ConnectionString = Core_Module.GetConstring
            dtaProduct.Fill(tblProduct, minPrimary, maxPrimary)
            For Each Me.RowItemBOM_ In Me.TblItemBOM_
                tblProduct.DefaultView.RowFilter = "[Primary] = " & Me.RowItemBOM_.Primary_Product.ToString
                If tblProduct.DefaultView.Count > 0 Then
                    rowProduct = tblProduct.DefaultView(0).Row
                    Me.RowItemBOM_.WIP_FG_Code = rowProduct.Product_Code
                    Me.RowItemBOM_.WIP_FG_Desc = rowProduct.Description
                End If
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txt_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txt1ProcessCode.GotFocus, txt1ProcessDesc.GotFocus, txt1RMCode.GotFocus, txt2CodeNo.GotFocus
        Dim txt As TextBox
        Try
            txt = TryCast(sender, TextBox)
            txt.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txt_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txt1ProcessCode.LostFocus, txt1ProcessDesc.LostFocus, txt1RMCode.LostFocus, txt2CodeNo.LostFocus
        Dim txt As TextBox
        Try
            txt = TryCast(sender, TextBox)
            txt.BackColor = Color.White
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btn2FindCodeNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn2FindCodeNo.Click
        Dim frmFindCodeNo As New Core_Module.frmCariMasterProduct
        Try
            frmFindCodeNo.txtCari.Text = Me.txt2CodeNo.Text
            frmFindCodeNo.TipePencarian = Core_Module.frmCariMasterProduct.enumTipePencarian.Goods_Only
            frmFindCodeNo.TipeProduct = Core_Module.clsMasterProduct.EnumTipeProduct.Goods
            frmFindCodeNo.ShowDialog()
            If frmFindCodeNo.PrimaryProduct > -1 Then
                Me.txt2CodeNo.Text = objProduct.GetProductCode(frmFindCodeNo.PrimaryProduct)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvProcess_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvProcess.CellBeginEdit
        Try
            If e.ColumnIndex = Me.colProcess_Code.Index Then
                Me.buffer = Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvProcess_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvProcess.CellEndEdit
        Try
            If e.ColumnIndex = Me.colProcess_Code.Index Then
                If IsNothing(Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    If String.IsNullOrEmpty(Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value.ToString.Trim) = False Then
                        Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value.ToString.ToUpper.Trim
                        If Me.buffer = Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value Then
                            Exit Sub
                        End If
                        If Me.objProcess.ApakahSudahAda(Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            If Me.IsExistProcess(Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value, e.RowIndex) = True Then
                                MsgBox("Process code: " & Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value & " already exist.", MsgBoxStyle.Information, "Information")
                                Me.BackToBufferOrDelete(e)
                                Exit Sub
                            End If
                            Me.dgvProcess.Item(Me.colProcess_Primary.Index, e.RowIndex).Value = Me.objProcess.GetPrimary(Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value)
                            Me.dgvProcess.Item(Me.colProcess_Desc.Index, e.RowIndex).Value = Me.objProcess.GetDesc(Me.dgvProcess.Item(Me.colProcess_Primary.Index, e.RowIndex).Value)
                        Else
                            Dim frmFindProcess As New Core_Module.FrmMasterProsesCari
                            frmFindProcess.txtCari.Text = Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value
                            'frmFindProcess.mStatusAktif = True
                            'frmFindProcess.mSearchAll = False
                            frmFindProcess.ShowDialog()
                            If frmFindProcess.mPrimary > -1 Then
                                If Me.IsExistProcess(Me.objProcess.GetCode(frmFindProcess.mPrimary), e.RowIndex) = True Then
                                    MsgBox("Process code: " & Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value & " already exist.", MsgBoxStyle.Information, "Information")
                                    Me.BackToBufferOrDelete(e)
                                    Exit Sub
                                End If
                                Me.dgvProcess.Item(Me.colProcess_Primary.Index, e.RowIndex).Value = frmFindProcess.mPrimary
                                Me.dgvProcess.Item(Me.colProcess_Code.Index, e.RowIndex).Value = Me.objProcess.GetCode(frmFindProcess.mPrimary)
                                Me.dgvProcess.Item(Me.colProcess_Desc.Index, e.RowIndex).Value = Me.objProcess.GetDesc(frmFindProcess.mPrimary)
                            Else
                                Me.BackToBufferOrDelete(e)
                                Exit Sub
                            End If
                        End If
                    Else
                        Me.BackToBufferOrDelete(e)
                        Exit Sub
                    End If
                Else
                    Me.BackToBufferOrDelete(e)
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub BackToBufferOrDelete(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Try
            If IsNothing(Me.buffer) = False Then
                If String.IsNullOrEmpty(Me.buffer.ToString.Trim) = False Then
                    Me.dgvProcess.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                Else
                    Me.rowToDelete = e.RowIndex
                    Me.BackgroundWorker1.RunWorkerAsync()
                End If
            Else
                Me.rowToDelete = e.RowIndex
                Me.BackgroundWorker1.RunWorkerAsync()
            End If
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
            If Me.rowToDelete < Me.dgvProcess.RowCount - 1 Then
                Me.dgvProcess.Rows.RemoveAt(Me.rowToDelete)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvProcess_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvProcess.DataError

    End Sub

    Private Sub dgvProcess_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvProcess.RowsAdded
        Try
            rowDgv = Me.dgvProcess.Rows(Me.dgvProcess.Rows.Count - 1)
            rowDgv.Cells(Me.colProcess_No.Index).Value = Me.dgvProcess.Rows.Count
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvProcess_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvProcess.RowsRemoved
        Try
            Me.objCD.SetNumber(Me.dgvProcess)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvProcess_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvProcess.Sorted
        Try
            Me.objCD.SetNumber(Me.dgvProcess)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function IsExistProcess(ByVal vProcessCode As String, ByVal vCurrentRow As Integer) As Boolean
        Try
            Me.dgvProcess.AllowUserToAddRows = False
            For Each Me.rowDgv In Me.dgvProcess.Rows
                If Me.dgvProcess.Rows.IndexOf(Me.rowDgv) <> vCurrentRow Then
                    If Me.rowDgv.Cells(Me.colProcess_Code.Index).Value.ToString.Trim.ToUpper = vProcessCode.Trim.ToUpper Then
                        Return True
                    End If
                End If
            Next
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.dgvProcess.AllowUserToAddRows = True
        End Try
    End Function

    Private Sub btn2Filter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn2Filter.Click
        Dim dta As New Core_Module.DseMasterProduct_BOM1TableAdapters.DataTable1TableAdapter
        Dim tbl As New Core_Module.DseMasterProduct_BOM1.DataTable1DataTable
        Dim row As Core_Module.DseMasterProduct_BOM1.DataTable1Row
        Try
            Me.dgvItem2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItem2.Rows.Clear()

            ''isi tbl dari BoM atas JO ybs
            If Not String.IsNullOrEmpty(Me.GetProcessFilter) Then
                TblItemBOM_.DefaultView.RowFilter = "[Primary MO] = " & Me.PrimaryMO.ToString & " and [Identifier JO] = '" & Me.IdentifierJO & "' and [WIP FG Code] LIKE '" & Me.txt2CodeNo.Text.Trim & "%' and [WIP FG Desc] LIKE '" & Me.txt2Desc.Text.Trim & "%'" & " and " & Me.GetProcessFilter
            Else
                TblItemBOM_.DefaultView.RowFilter = "[Primary MO] = " & Me.PrimaryMO.ToString & " and [Identifier JO] = '" & Me.IdentifierJO & "' and [WIP FG Code] LIKE '" & Me.txt2CodeNo.Text.Trim & "%' and [WIP FG Desc] LIKE '" & Me.txt2Desc.Text.Trim & "%'"
            End If
            For Each Me.DV In TblItemBOM_.DefaultView
                Me.RowItemBOM_ = Me.DV.Row
                tbl.DefaultView.RowFilter = "[Primary Product] = " & Me.RowItemBOM_.Primary_Product.ToString & " and [Identifier WO] = '" & Me.RowItemBOM_.Identifier_WO & "'"
                If tbl.DefaultView.Count = 0 Then
                    row = tbl.NewDataTable1Row
                    With row
                        .Primary_Product = Me.RowItemBOM_.Primary_Product
                        .Product_Code = Me.RowItemBOM_.WIP_FG_Code
                        .Product_Desc = Me.RowItemBOM_.WIP_FG_Desc
                        .Primary_Process = Me.RowItemBOM_.Primary_Process
                        .Process_Code = Me.objProcess.GetCode(Me.RowItemBOM_.Primary_Process)
                        .Process_Desc = String.Empty
                        .Qty_Ref = Me.RowItemBOM_.Qty_Ref_Process
                        ._Default__Precision_ = Me.objProduct.GetAngkaDecimal(Me.RowItemBOM_.Primary_Product)
                        .Primary_Unit = Me.objProduct.GetPrimaryUnit(Me.RowItemBOM_.Primary_Product)
                        .Unit_Code = Me.objUnit.GetUnitCode(.Primary_Unit)
                        .Unit_Desc = String.Empty
                        .GP_Process = Me.RowItemBOM_.Good_Percentage_Process
                        .Tolerance_Process = Me.RowItemBOM_.tolerance_process
                        .Identifier_WO = Me.RowItemBOM_.Identifier_WO
                        .WO_No = Me.GetWO_NoByIdentifier(Me.RowItemBOM_.Identifier_WO)
                        .Primary_Parent_Product = Me.RowItemBOM_.Primary_Parent_Product
                        .IdentifierBOM_Parent = Me.RowItemBOM_.Identifier_BOM_Parent
                    End With
                    tbl.Rows.Add(row)
                End If
            Next

            If tbl.Count > 0 Then
                Me.Cursor = Cursors.WaitCursor
                tbl.DefaultView.RowFilter = ""
                tbl.DefaultView.Sort = "[WO No], [Product Code]"
                If tbl.DefaultView.Count > 0 Then
                    For Each Me.DV In tbl.DefaultView
                        row = Me.DV.Row
                        Me.rowDgv = Me.dgvItem2.Rows(Me.dgvItem2.Rows.Add)
                        With Me.rowDgv
                            .Cells(Me.col2_GPProcess.Index).Value = Me.objCD.DigitGrouping(row.GP_Process, 2)
                            .Cells(Me.col2_Tolerance.Index).Value = Me.objCD.DigitGrouping(row.Tolerance_Process, 2)
                            .Cells(Me.col2CodeNo.Index).Value = row.Product_Code
                            .Cells(Me.col2Desc.Index).Value = row.Product_Desc
                            .Cells(Me.col2Modify.Index).Value = Boolean.FalseString
                            .Cells(Me.col2No.Index).Value = Me.dgvItem2.Rows.Count.ToString
                            .Cells(Me.col2PrimaryCodeNo.Index).Value = row.Primary_Product
                            .Cells(Me.col2PrimaryProcess.Index).Value = row.Primary_Process
                            .Cells(Me.col2PrimaryUnit.Index).Value = row.Primary_Unit
                            .Cells(Me.col2Process.Index).Value = row.Process_Code
                            .Cells(Me.col2QtyRef.Index).Value = Me.objCD.DigitGrouping(row.Qty_Ref, row._Default__Precision_)
                            .Cells(Me.col2Unit.Index).Value = row.Unit_Code
                            .Cells(Me.Col2Decimal.Index).Value = row._Default__Precision_
                            .Cells(Me.Col2IdentifierWO.Index).Value = row.Identifier_WO
                            .Cells(Me.Col2WO_No.Index).Value = row.WO_No
                            .Cells(Me.Col2PrimaryParentProduct.Index).Value = row.Primary_Parent_Product
                            .Cells(Me.Col2IdentifierBoM_Parent.Index).Value = row.IdentifierBOM_Parent
                        End With
                    Next
                Else
                    MsgBox("Product not found.", MsgBoxStyle.Information, "Information")
                End If
            Else
                MsgBox("Product not found.", MsgBoxStyle.Information, "Information")
            End If
            Me.dgvItem2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Function GetProcessFilter() As String
        GetProcessFilter = String.Empty
        Try
            Me.dgvProcess.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvProcess.AllowUserToAddRows = False
            If Me.dgvProcess.RowCount > 0 Then
                For Each Me.rowDgv In Me.dgvProcess.Rows
                    If String.IsNullOrEmpty(Me.rowDgv.Cells(Me.colProcess_Primary.Index).Value) = False Then
                        If String.IsNullOrEmpty(GetProcessFilter) Then
                            GetProcessFilter = "[Primary Process] = " & Me.rowDgv.Cells(Me.colProcess_Primary.Index).Value.ToString
                        Else
                            GetProcessFilter = GetProcessFilter & " OR [Primary Process] = " & Me.rowDgv.Cells(Me.colProcess_Primary.Index).Value.ToString
                        End If
                    End If
                Next
            End If
            Me.dgvProcess.AllowUserToAddRows = True
            Me.dgvProcess.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            If String.IsNullOrEmpty(GetProcessFilter) Then
                Return String.Empty
            Else
                Return "(" & GetProcessFilter & ")"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub llbCheckAll2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbCheckAll2.LinkClicked
        Try
            Me.dgvItem2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For Each rowDgv In Me.dgvItem2.Rows
                Me.rowDgv.Cells(Me.col2Modify.Index).Value = Boolean.TrueString
            Next
            Me.dgvItem2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbUnCheckAll2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbUnCheckAll2.LinkClicked
        Try
            Me.dgvItem2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For Each rowDgv In Me.dgvItem2.Rows
                Me.rowDgv.Cells(Me.col2Modify.Index).Value = Boolean.FalseString
            Next
            Me.dgvItem2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem2_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem2.CellBeginEdit
        Try
            If e.ColumnIndex = Me.col2Modify.Index Or e.ColumnIndex = Me.col2QtyRef.Index Or e.ColumnIndex = Me.col2_GPProcess.Index Or e.ColumnIndex = Me.col2_Tolerance.Index Then
                Me.buffer = Me.dgvItem2.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem2_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem2.CellEndEdit
        Try
            If e.ColumnIndex = Me.col2QtyRef.Index Then
                If Me.objCD.CekBilanganPositif(Me.dgvItem2.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    Me.dgvItem2.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
                Me.dgvItem2.Item(e.ColumnIndex, e.RowIndex).Value = Me.objCD.DigitGrouping(Me.dgvItem2.Item(e.ColumnIndex, e.RowIndex).Value, Me.dgvItem2.Item(Me.Col2Decimal.Index, e.RowIndex).Value)
                If CDec(Me.dgvItem2.Item(e.ColumnIndex, e.RowIndex).Value) <= 0 Then
                    Me.dgvItem2.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
                Me.UpdateQtyReference(Me.dgvItem2.Item(e.ColumnIndex, e.RowIndex).Value, Me.dgvItem2.Item(Me.col2PrimaryCodeNo.Index, e.RowIndex).Value, Me.dgvItem2.Item(Me.Col2IdentifierWO.Index, e.RowIndex).Value)
            End If
            If e.ColumnIndex = Me.col2_GPProcess.Index Then
                If Me.objCD.CekBilanganPositif(Me.dgvItem2.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    Me.dgvItem2.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
                If CDec(Me.dgvItem2.Item(e.ColumnIndex, e.RowIndex).Value) > 100 Then
                    Me.dgvItem2.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
                Me.dgvItem2.Item(e.ColumnIndex, e.RowIndex).Value = Me.objCD.DigitGrouping(Me.dgvItem2.Item(e.ColumnIndex, e.RowIndex).Value, 2)
                If CDec(Me.dgvItem2.Item(e.ColumnIndex, e.RowIndex).Value) <= 0 Then
                    Me.dgvItem2.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
                Me.dgvItem2.Item(Me.col2_Tolerance.Index, e.RowIndex).Value = Me.objCD.DigitGrouping(Me.frmProduct.GetTolerancePerc(Me.dgvItem2.Item(Me.col2_GPProcess.Index, e.RowIndex).Value), 2)
                Me.UpdateGP_AndTolerance_FG(Me.dgvItem2.Item(Me.col2_GPProcess.Index, e.RowIndex).Value, Me.dgvItem2.Item(Me.col2_Tolerance.Index, e.RowIndex).Value, Me.dgvItem2.Item(Me.col2PrimaryCodeNo.Index, e.RowIndex).Value, Me.dgvItem2.Item(Me.Col2IdentifierWO.Index, e.RowIndex).Value)
            End If
            If e.ColumnIndex = Me.col2_Tolerance.Index Then
                If Me.objCD.CekNonNegatif(Me.dgvItem2.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    Me.dgvItem2.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
                Me.dgvItem2.Item(e.ColumnIndex, e.RowIndex).Value = Me.objCD.DigitGrouping(Me.dgvItem2.Item(e.ColumnIndex, e.RowIndex).Value, 2)
                Me.dgvItem2.Item(Me.col2_GPProcess.Index, e.RowIndex).Value = Me.objCD.DigitGrouping(Me.frmProduct.GetGoodPercentagePerc(Me.dgvItem2.Item(Me.col2_Tolerance.Index, e.RowIndex).Value), 2)
                If CDec(Me.dgvItem2.Item(Me.col2_GPProcess.Index, e.RowIndex).Value) <= 0 Then
                    Me.dgvItem2.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Me.dgvItem2.Item(Me.col2_GPProcess.Index, e.RowIndex).Value = Me.objCD.DigitGrouping(Me.frmProduct.GetGoodPercentagePerc(Me.dgvItem2.Item(Me.col2_Tolerance.Index, e.RowIndex).Value), 2)
                    Exit Sub
                End If
                Me.UpdateGP_AndTolerance_FG(Me.dgvItem2.Item(Me.col2_GPProcess.Index, e.RowIndex).Value, Me.dgvItem2.Item(Me.col2_Tolerance.Index, e.RowIndex).Value, Me.dgvItem2.Item(Me.col2PrimaryCodeNo.Index, e.RowIndex).Value, Me.dgvItem2.Item(Me.Col2IdentifierWO.Index, e.RowIndex).Value)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function UpdateQtyReference(ByVal vQtyRef As Decimal, ByVal vPrimaryProduct As Double, ByVal vIdentifierWO As String) As Boolean
        Try
            Me.TblItemBOM_.DefaultView.RowFilter = "[Primary MO] = " & Me.PrimaryMO.ToString & " and [Identifier JO]='" & Me.IdentifierJO & "' and [Primary Product] = " & vPrimaryProduct.ToString & " and [Identifier WO] = '" & vIdentifierWO & "'"
            If Me.TblItemBOM_.DefaultView.Count > 0 Then
                For Each Me.DV In Me.TblItemBOM_.DefaultView
                    Me.RowItemBOM_ = Me.DV.Row
                    Me.RowItemBOM_.Qty_Ref_Process = vQtyRef
                Next
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Function UpdateGP_AndTolerance_FG(ByVal vGP As Decimal, ByVal vTolerance As Decimal, ByVal vPrimaryProduct As Double, ByVal vIdentifierWO As String) As Boolean
        Try
            Me.TblItemBOM_.DefaultView.RowFilter = "[Primary MO] = " & Me.PrimaryMO.ToString & " and [Identifier JO]='" & Me.IdentifierJO & "' and [Primary Product] = " & vPrimaryProduct.ToString & " and [Identifier WO] = '" & vIdentifierWO & "'"
            If Me.TblItemBOM_.DefaultView.Count > 0 Then
                For Each Me.DV In Me.TblItemBOM_.DefaultView
                    Me.RowItemBOM_ = Me.DV.Row
                    Me.RowItemBOM_.Good_Percentage_Process = vGP
                    Me.RowItemBOM_.tolerance_process = vTolerance
                Next
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Sub dgvItem2_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvItem2.DataError

    End Sub

    Private Sub dgvItem2_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem2.Sorted
        Try
            Me.objCD.SetNumber(Me.dgvItem2)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btn2SetQtyRef_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn2SetQtyRef.Click
        Dim frmSetQtyRef As New Core_Module.frmMasterProduct_BOM1_SetQtyReference
        Try
            If Me.GetSelectedCount(Me.dgvItem2, Me.col2Modify) = 0 Then
                MsgBox("Select at least one product to modify.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            If Not Me.IsAllColumnSimilar(Me.dgvItem2, Me.col2Unit, Me.col2Modify) Then
                MsgBox("Unit code for all selected products must be same.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            frmSetQtyRef.Text = "Set Qty Ref"
            frmSetQtyRef.lblQtyReference.Text = "Qty reference"
            frmSetQtyRef.lblUnit.Text = Me.GetFirstSelectedValue(Me.dgvItem2, Me.col2Modify, Me.col2Unit)
            frmSetQtyRef.BackColorAll = Me.BackColorAll
            frmSetQtyRef.Decimal_ = Me.GetFirstSelectedValue(Me.dgvItem2, Me.col2Modify, Me.Col2Decimal)
            If frmSetQtyRef.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.Cursor = Cursors.WaitCursor
                For Each Me.rowDgv In Me.dgvItem2.Rows
                    If Me.rowDgv.Cells(Me.col2Modify.Index).Value = Boolean.TrueString Then
                        If Me.UpdateQtyReference(frmSetQtyRef.QtyReference, Me.rowDgv.Cells(Me.col2PrimaryCodeNo.Index).Value, Me.rowDgv.Cells(Me.Col2IdentifierWO.Index).Value) = True Then
                            Me.rowDgv.Cells(Me.col2QtyRef.Index).Value = Me.objCD.DigitGrouping(frmSetQtyRef.QtyReference, Me.rowDgv.Cells(Me.Col2Decimal.Index).Value)
                        End If
                    End If
                Next
                MsgBox("Process complete.", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Function GetSelectedCount(ByVal vDgv As DataGridView, ByVal vCommitColumn As DataGridViewCheckBoxColumn) As Integer
        GetSelectedCount = 0
        Try
            For Each Me.rowDgv In vDgv.Rows
                If Me.rowDgv.Cells(vCommitColumn.Index).Value = Boolean.TrueString Then
                    GetSelectedCount += 1
                End If
            Next
            Return GetSelectedCount
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return GetSelectedCount
        End Try
    End Function

    Function IsAllColumnSimilar(ByVal vDgv As DataGridView, ByVal vColumn As DataGridViewTextBoxColumn, ByVal vCommitColumn As DataGridViewCheckBoxColumn) As Boolean
        Dim lastValue As Object = Nothing
        Dim trueFound As Integer = 0
        Try
            If vDgv.Rows.Count <= 1 Then
                Return True
            End If
            For Each Me.rowDgv In vDgv.Rows
                If Me.rowDgv.Cells(vCommitColumn.Index).Value = Boolean.TrueString Then
                    trueFound += 1
                    If trueFound > 1 And Me.rowDgv.Cells(vColumn.Index).Value <> lastValue Then
                        Return False
                    End If
                    lastValue = Me.rowDgv.Cells(vColumn.Index).Value
                End If
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Function GetFirstSelectedValue(ByVal vDgv As DataGridView, ByVal vModifyColumn As DataGridViewCheckBoxColumn, ByVal vColValueToGet As DataGridViewTextBoxColumn) As String
        GetFirstSelectedValue = String.Empty
        Try
            For Each Me.rowDgv In vDgv.Rows
                If Me.rowDgv.Cells(vModifyColumn.Index).Value = Boolean.TrueString Then
                    Return Me.rowDgv.Cells(vColValueToGet.Index).Value
                End If
            Next
            Return GetFirstSelectedValue
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return GetFirstSelectedValue
        End Try
    End Function

    Private Sub btn2SetGPProcess_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn2SetGPProcess.Click
        Dim frmSetQtyRef As New Core_Module.frmMasterProduct_BOM1_SetQtyReference
        Try
            If Me.GetSelectedCount(Me.dgvItem2, Me.col2Modify) = 0 Then
                MsgBox("Select at least one product to modify.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            frmSetQtyRef.Text = "Set Good % Process"
            frmSetQtyRef.lblQtyReference.Text = "Good percentage"
            frmSetQtyRef.lblUnit.Text = "%"
            frmSetQtyRef.BackColorAll = Me.BackColorAll
            frmSetQtyRef.Decimal_ = 2
            frmSetQtyRef.Size = New System.Drawing.Size(frmSetQtyRef.btnOK.Left + frmSetQtyRef.btnOK.Width + 10, frmSetQtyRef.Height)
            If frmSetQtyRef.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.Cursor = Cursors.WaitCursor
                For Each Me.rowDgv In Me.dgvItem2.Rows
                    If Me.rowDgv.Cells(Me.col2Modify.Index).Value = Boolean.TrueString Then
                        If Me.UpdateGP_AndTolerance_FG(frmSetQtyRef.QtyReference, frmProduct.GetTolerancePerc(frmSetQtyRef.QtyReference), Me.rowDgv.Cells(Me.col2PrimaryCodeNo.Index).Value, Me.rowDgv.Cells(Me.Col2IdentifierWO.Index).Value) = True Then
                            Me.rowDgv.Cells(Me.col2_GPProcess.Index).Value = Me.objCD.DigitGrouping(frmSetQtyRef.QtyReference, 2)
                            Me.rowDgv.Cells(Me.col2_Tolerance.Index).Value = Me.objCD.DigitGrouping(frmProduct.GetTolerancePerc(frmSetQtyRef.QtyReference), 2)
                        End If
                    End If
                Next
                MsgBox("Process complete.", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub btn2AddRM_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn2AddRM.Click
        Dim frmAddRM As New Core_Module.frmMasterProduct_BOM1_AddRM
        Try
            If Me.GetSelectedCount(Me.dgvItem2, Me.col2Modify) = 0 Then
                MsgBox("Select at least one product to modify.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            If Not Me.IsAllColumnSimilar(Me.dgvItem2, Me.col2PrimaryProcess, Me.col2Modify) Then
                MsgBox("Process code for all selected products must be same.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            frmAddRM.UserCode = Me.UserCode
            frmAddRM.BackColorall = Me.BackColorAll
            If frmAddRM.ShowDialog = Windows.Forms.DialogResult.OK Then

                Dim listOfInvalid As New List(Of String)
                listOfInvalid = Me.GetAlreadyExistRM(frmAddRM)
                If listOfInvalid.Count > 0 Then
                    Dim existRM As String = String.Empty
                    For Each str As String In listOfInvalid
                        If String.IsNullOrEmpty(existRM) Then
                            existRM = str
                        Else
                            existRM = existRM & vbNewLine & str
                        End If
                    Next
                    MsgBox("The following raw material is already exist, add raw material process canceled." & vbNewLine & existRM, MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                listOfInvalid.Clear()
                listOfInvalid = Me.GetInfinitelyBoM(frmAddRM)
                If listOfInvalid.Count > 0 Then
                    Dim existRM As String = String.Empty
                    For Each str As String In listOfInvalid
                        If String.IsNullOrEmpty(existRM) Then
                            existRM = str
                        Else
                            existRM = existRM & vbNewLine & str
                        End If
                    Next
                    MsgBox("Adding the following RM will cause infinite recursive, process canceled." & vbNewLine & existRM, MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                Me.Cursor = Cursors.WaitCursor
                For Each Me.rowDgv In Me.dgvItem2.Rows
                    If Me.rowDgv.Cells(Me.col2Modify.Index).Value = Boolean.TrueString Then
                        For Each rowDgvRM As DataGridViewRow In frmAddRM.dgvItemBOM.Rows
                            If rowDgvRM.Cells(frmAddRM.ColBOM_PrimaryProduct.Index).Value <> "-1" Then
                                Dim frmFunction As New FrmTrans_ManufacturingOrder_BOM

                                frmFunction.TblItemBom_IsiTemporary2(Me.TblItemBOM_, Me.PrimaryMO, Me.IdentifierJO, Me.rowDgv.Cells(Me.col2PrimaryCodeNo.Index).Value, _
                                Me.rowDgv.Cells(Me.col2PrimaryProcess.Index).Value, Me.rowDgv.Cells(Me.col2QtyRef.Index).Value, Me.rowDgv.Cells(Me.col2_GPProcess.Index).Value, _
                                rowDgvRM.Cells(frmAddRM.ColBOM_PrimaryProduct.Index).Value, rowDgvRM.Cells(frmAddRM.ColBOM_Qty.Index).Value, rowDgvRM.Cells(frmAddRM.ColBOM_GoodPercentage.Index).Value, _
                                rowDgv.Cells(Me.Col2IdentifierWO.Index).Value, rowDgvRM.Cells(frmAddRM.ColBOM_Remarks.Index).Value, Me.rowDgv.Cells(Me.Col2PrimaryParentProduct.Index).Value, _
                                Me.rowDgv.Cells(Me.Col2IdentifierBoM_Parent.Index).Value, rowDgvRM.Cells(frmAddRM.ColBOM_CodeNo.Index).Value, rowDgvRM.Cells(frmAddRM.ColBOM_Desc.Index).Value, _
                                rowDgvRM.Cells(frmAddRM.ColBOM_Tolerance.Index).Value, Me.rowDgv.Cells(Me.col2_Tolerance.Index).Value)
                            End If
                        Next
                    End If
                Next
                MsgBox("Process complete.", MsgBoxStyle.Information, "Information")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Function GetInfinitelyBoM(ByVal frmAddRM As Core_Module.frmMasterProduct_BOM1_AddRM) As List(Of String)
        Dim listOfBoM As New List(Of Double)
        Dim listOfChild As New List(Of Double)
        GetInfinitelyBoM = New List(Of String)
        Try
            For Each Me.rowDgv In frmAddRM.dgvItemBOM.Rows
                If Me.rowDgv.Cells(frmAddRM.ColBOM_PrimaryProduct.Index).Value <> "-1" Then
                    listOfBoM.Add(Me.rowDgv.Cells(frmAddRM.ColBOM_PrimaryProduct.Index).Value)
                End If
            Next
            For Each Me.rowDgv In Me.dgvItem2.Rows
                If Me.rowDgv.Cells(Me.col2Modify.Index).Value = Boolean.TrueString Then
                    listOfChild.Clear()
                    Me.GetChild(Me.rowDgv.Cells(Me.col2PrimaryCodeNo.Index).Value, listOfChild)
                    If listOfChild.Count > 0 Then
                        Dim e As String = ""
                        For Each d As Double In listOfBoM
                            If listOfChild.Contains(d) Then
                                e = objProduct.GetProductCode(d)
                                GetInfinitelyBoM.Add("Product code " & e & "  as raw material of " & Me.rowDgv.Cells(Me.col2CodeNo.Index).Value)
                            End If
                        Next
                    End If
                End If
            Next
            Return GetInfinitelyBoM
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return GetInfinitelyBoM
        End Try
    End Function

    Function GetInfinitelyBoM2(ByVal vPrimaryNewRM As Double) As List(Of String)
        Dim listOfBoM As New List(Of Double)
        Dim listOfChild As New List(Of Double)
        GetInfinitelyBoM2 = New List(Of String)
        Try
            Me.GetChild(Me.rowDgv.Cells(Me.ColPrimaryCodeNoRM.Index).Value, listOfChild)
            If listOfChild.Count > 0 Then
                If listOfChild.Contains(vPrimaryNewRM) Then
                    GetInfinitelyBoM2.Add("Product code " & objProduct.GetProductCode(vPrimaryNewRM) & "  as raw material of " & Me.rowDgv.Cells(Me.col2CodeNo.Index).Value)
                End If
            End If
            Return GetInfinitelyBoM2
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return GetInfinitelyBoM2
        End Try
    End Function

    Private Sub GetChild(ByVal vPrimaryProduct As Double, ByVal vListChild As List(Of Double))
        Dim TblBoM As New Core_Module.dseMasterProduct_BOM.tblmaster_product_item_bomDataTable
        Dim RowBoM As Core_Module.dseMasterProduct_BOM.tblmaster_product_item_bomRow
        Try
            Me.TblItemBOM_.DefaultView.RowFilter = "[Primary Product RM] = " & vPrimaryProduct.ToString
            If TblItemBOM_.DefaultView.Count > 0 Then
                For Each DV In TblItemBOM_.DefaultView
                    RowItemBOM_ = DV.Row
                    RowBoM = TblBoM.Newtblmaster_product_item_bomRow
                    RowBoM.primary = -1
                    RowBoM.primary_product = RowItemBOM_.Primary_Product
                    RowBoM.primary_product_raw = RowItemBOM_.Primary_Product_RM
                    RowBoM.qty_raw = RowItemBOM_.Qty_RM
                    RowBoM.good_percentage = RowItemBOM_.Good_Percentage_RM
                    RowBoM.remarks = String.Empty
                    TblBoM.Rows.Add(RowBoM)
                Next
            End If
            If vListChild.Contains(vPrimaryProduct) = False Then
                vListChild.Add(vPrimaryProduct)
            End If
            For Each RowBoM In TblBoM
                If vListChild.Contains(RowBoM.primary_product) = False Then
                    vListChild.Add(RowBoM.primary_product)
                End If
                Me.GetChild(RowBoM.primary_product, vListChild)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Function GetAlreadyExistRM(ByVal frmAddRM As Core_Module.frmMasterProduct_BOM1_AddRM) As List(Of String)
        Dim tbl As New Core_Module.DseMasterProduct_BOM1.tblmaster_product_item_bomDataTable
        Dim row As Core_Module.DseMasterProduct_BOM1.tblmaster_product_item_bomRow
        GetAlreadyExistRM = New List(Of String)
        Try
            For Each Me.rowDgv In Me.dgvItem2.Rows
                If Me.rowDgv.Cells(Me.col2Modify.Index).Value = Boolean.TrueString Then
                    Me.TblItemBOM_.DefaultView.RowFilter = "[Primary Product] = " & Me.rowDgv.Cells(Me.col2PrimaryCodeNo.Index).Value.ToString
                    For Each DV As Data.DataRowView In Me.TblItemBOM_.DefaultView
                        Me.RowItemBOM_ = DV.Row
                        row = tbl.Newtblmaster_product_item_bomRow
                        row.Finish_Good_Code = Me.rowDgv.Cells(Me.col2CodeNo.Index).Value
                        row.primary_product = Me.rowDgv.Cells(Me.col2PrimaryCodeNo.Index).Value
                        row.primary_product_raw = Me.RowItemBOM_.Primary_Product_RM
                        row.Raw_Material_Code = Me.objProduct.GetProductCode(row.primary_product_raw)
                        tbl.Rows.Add(row)
                    Next
                End If
            Next
            For Each Me.rowDgv In frmAddRM.dgvItemBOM.Rows
                If Me.rowDgv.Cells(frmAddRM.ColBOM_PrimaryProduct.Index).Value <> "-1" Then
                    tbl.DefaultView.RowFilter = "[primary product raw] = " & Me.rowDgv.Cells(frmAddRM.ColBOM_PrimaryProduct.Index).Value.ToString
                    If tbl.DefaultView.Count > 0 Then
                        For Each DV As Data.DataRowView In tbl.DefaultView
                            row = DV.Row
                            GetAlreadyExistRM.Add(Me.rowDgv.Cells(frmAddRM.ColBOM_CodeNo.Index).Value & " as raw material of " & row.Finish_Good_Code)
                        Next
                    End If
                End If
            Next
            Return GetAlreadyExistRM
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return GetAlreadyExistRM
        End Try
    End Function

    Function GetWO_NoByIdentifier(ByVal vIdentifier As String) As String
        Try
            Me.vTblItemWO.DefaultView.RowFilter = "[Identifier] = '" & vIdentifier & "'"
            If vTblItemWO.DefaultView.Count > 0 Then
                Me.RowlItemWO = Me.vTblItemWO.DefaultView(0).Row
                Return Me.RowlItemWO.WO_Number
            End If
            Return String.Empty
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return String.Empty
        End Try
    End Function
#End Region

    Private Sub LlbBoMTools_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbEditBoM.LinkClicked
        Dim i As New FrmTrans_ManufacturingOrder_BOM
        Dim objUser As New Core_Module.ClsUser
        Try
            Me.dgvItem1.Rows.Clear()
            Me.dgvItem2.Rows.Clear()
            i.vPrimaryMO = PrimaryMO
            i.vIdentifierJO = IdentifierJO
            i.vTblItemBOM = Me.TblItemBOM_
            i.vTblItemWO = Me.vTblItemWO
            i.vPrimaryProduct = Me.vPrimaryProduct
            i.vPrimaryProductJO = Me.vPrimaryProductJO
            i.vProductCode = Me.vProductCode
            i.vDefPrec = Me.vDefPrec
            i.vQtyTarget = Me.vQtyTarget
            i.vTblProsesOrder = Me.vTblProsesOrder
            i.vTblItemRM = Me.vTblItemRM
            If objUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_ManufacturingOrder_AddItem") = False Then
                i.vBolehEdit = False
            Else
                i.vBolehEdit = True
            End If
            i.ShowDialog()
            Me.TblItemBOM_ = i.vTblItemBOM
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

End Class