Option Explicit On
Option Strict Off

Public Class FrmTrans_ManufacturingOrder_Exp_QtyReq
    Public IdentifierJO As String = ""
    Public IdentifierWO As String = ""
    Public TblItemBom As New DseTrans_ManufacturingOrder.tbltrans_mo_item_bomDataTable
    Public RowItemBom As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomRow
    Public TblItemWO As New DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable
    Public RowItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woRow
    Public PrimaryRM As Double = -1
    Public TblItemRM As New DseTrans_ManufacturingOrder.tbltrans_mo_item_rmDataTable
    Public RowItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmRow
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public JONO As String = ""
    Public CodeNo As String = ""
    Public Desc As String = ""

    Dim TBL As New DseTrans_ManufacturingOrder.tblexplain_qtyreqDataTable
    Dim StartExplain As Boolean = False
    Dim ObjProduct As New Core_Module.clsMasterProduct
    Dim ObjProcess As New Core_Module.ClsMasterProses
    Dim ObjCD As New Core_Module.ClsCekData
    Dim BufferIndex As Integer = -1

    Private Sub FrmTrans_ManufacturingOrder_Exp_QtyReq_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim rowDGV As DataGridViewRow
        Dim Scroll1 As Boolean = False

        Try
            Me.DgvItemBundling.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            Me.TxtJONumber.Text = Me.JONO
            Me.TxtCodeNo.Text = Me.CodeNo
            Me.TxtDesc.Text = Me.Desc

            Me.LoadTBL()
            Me.StartExplain = True
            Me.DgvItemBundling.ClearSelection()

            rowDGV = Me.DgvItemBundling.Rows(Me.DgvItemBundling.Rows.Count - 1)
            While rowDGV.Displayed = False
                Me.DgvItemBundling.FirstDisplayedScrollingRowIndex += 1
                Scroll1 = True
            End While
            If Scroll1 = True Then
                Me.DgvItemBundling.FirstDisplayedScrollingRowIndex += 1
            End If
            rowDGV.Selected = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadTBL()
        Dim TBLDummy As New DseTrans_ManufacturingOrder.tblexplain_qtyreqDataTable
        Dim RowDummy, Row As DseTrans_ManufacturingOrder.tblexplain_qtyreqRow
        Dim rowDGV As DataGridViewRow

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            TBL.Rows.Clear()

            'lacak dari item identifier jo dan identifier wo ybs, sampai ke paling atas, 
            'sambil lacak, sambil isikan data pada baris tsb
            Me.LoadTBL_RM(Me.IdentifierJO, Me.IdentifierWO, Me.PrimaryRM)

            'balik urutan barisnya
            TBLDummy = TBL.Copy
            TBL.Rows.Clear()
            For ij As Integer = TBLDummy.Rows.Count - 1 To 0 Step -1
                RowDummy = TBLDummy.Rows(ij)
                Row = TBL.Newtblexplain_qtyreqRow
                For k As Integer = 0 To TBL.Columns.Count - 1
                    Row.Item(k) = RowDummy.Item(k)
                Next
                TBL.Rows.Add(Row)
            Next

            'tampilkan ke dgv
            For Each Row In TBL.Rows
                Me.DgvItemBundling.Rows.Add()
                rowDGV = Me.DgvItemBundling.Rows(Me.DgvItemBundling.Rows.Count - 1)

                rowDGV.Cells(Me.ColCodeNo.Index).Value = Row.Code_No
                rowDGV.Cells(Me.ColDesc.Index).Value = Row.Desc
                rowDGV.Cells(Me.ColDocNo.Index).Value = Row.WO_Number
                rowDGV.Cells(Me.ColNo.Index).Value = rowDGV.Index + 1
                rowDGV.Cells(Me.ColProcess.Index).Value = Me.ObjProcess.GetCode(Row.Primary_Process)
                rowDGV.Cells(Me.ColDecimal_Product.Index).Value = Me.ObjProduct.GetAngkaDecimal(Row.Primary_Product)
                rowDGV.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(Row.Qty, rowDGV.Cells(Me.ColDecimal_Product.Index).Value)
                rowDGV.Cells(Me.ColRMCode.Index).Value = Row.Code_No_RM
                rowDGV.Cells(Me.ColRMDesc.Index).Value = Row.Desc_RM
                rowDGV.Cells(Me.ColDecimalRM.Index).Value = Me.ObjProduct.GetAngkaDecimal(Row.Primary_RM)
                rowDGV.Cells(Me.ColRMQTY.Index).Value = Me.ObjCD.DigitGrouping(Row.Qty_RM, rowDGV.Cells(Me.ColDecimalRM.Index).Value)
                rowDGV.Cells(Me.ColWipQTY.Index).Value = Me.ObjCD.DigitGrouping(Row.Qty_WIP, rowDGV.Cells(Me.ColDecimalRM.Index).Value)
                rowDGV.Cells(Me.ColRMUnit.Index).Value = Row.Unit_Code_RM
                rowDGV.Cells(Me.ColUnit.Index).Value = Row.Unit_Code
                rowDGV.Cells(Me.ColQtyRef.Index).Value = Row.Qty_Ref_BOM
                rowDGV.Cells(Me.ColQtyRefRM.Index).Value = Row.Qty_RM_BOM
                rowDGV.Cells(Me.ColGPProcess.Index).Value = Row.Good_Percentage_Process
                rowDGV.Cells(Me.ColGPRM.Index).Value = Row.Good_Percentage_RM
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadTBL_RM(ByVal vIdentifierJO As String, ByVal vIdentifierWO As String, ByVal vPrimaryRM As Double)
        Dim Row As DseTrans_ManufacturingOrder.tblexplain_qtyreqRow

        Try
            Me.TblItemBom.DefaultView.RowFilter = "[Identifier JO] = '" + vIdentifierJO + _
            "' AND [Identifier WO] = '" + vIdentifierWO + "' AND [Primary Product RM] = " + vPrimaryRM.ToString
            If Me.TblItemBom.DefaultView.Count = 1 Then
                Me.TblItemWO.DefaultView.RowFilter = "[Identifier JO] = '" + vIdentifierJO + "' AND Identifier = '" + vIdentifierWO + "'"
                If Me.TblItemWO.DefaultView.Count <> 1 Then
                    MsgBox("Logic error: no item wo specified. Please contact your endor if you see this message.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
                Me.RowItemWO = Me.TblItemWO.DefaultView(0).Row
                Me.TblItemRM.DefaultView.RowFilter = "[Identifier JO] = '" + vIdentifierJO + "' AND [Identifier WO] = '" + vIdentifierWO + _
                "' AND [Primary Product Code RM] = " + vPrimaryRM.ToString
                If Me.TblItemRM.DefaultView.Count <> 1 Then
                    MsgBox("Logic error: no item raw material specified. Please contact your endor if you see this message.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
                Me.RowItemRM = Me.TblItemRM.DefaultView(0).Row

                Me.RowItemBom = Me.TblItemBom.DefaultView(0).Row
                Row = TBL.Newtblexplain_qtyreqRow
                Row.Code_No = Me.ObjProduct.GetProductCode(Me.RowItemBom.Primary_Product)
                Row.Desc = Me.ObjProduct.GetDescription(Me.RowItemBom.Primary_Product)
                Row.Code_No_RM = Me.ObjProduct.GetProductCode(Me.RowItemBom.Primary_Product_RM)
                Row.Desc_RM = Me.ObjProduct.GetDescription(Me.RowItemBom.Primary_Product_RM)
                Row.Decimal_RM = Me.RowItemBom.Decimal_RM
                Row.Good_Percentage_Process = Me.RowItemBom.Good_Percentage_Process
                Row.Good_Percentage_RM = Me.RowItemBom.Good_Percentage_RM
                Row.Identifier_BOM = Me.RowItemBom.Identifier_BOM
                Row.Identifier_BOM_Parent = Me.RowItemBom.Identifier_BOM_Parent
                Row.Identifier_JO = Me.RowItemBom.Identifier_JO
                Row.Identifier_WO = Me.RowItemBom.Identifier_WO
                Row.Primary_Product = Me.RowItemBom.Primary_Product
                Row.Primary_RM = Me.RowItemBom.Primary_Product_RM
                Row.Qty_Ref_BOM = Me.RowItemBom.Qty_Ref_Process
                Row.Qty_RM_BOM = Me.RowItemBom.Qty_RM
                Row.Unit_Code = Me.ObjProduct.GetUnitCode(Row.Primary_Product)
                Row.Unit_Code_RM = Me.ObjProduct.GetUnitCode(Row.Primary_RM)
                Row.Qty = RowItemWO.Qty_Target_Result
                Row.Qty_RM = Me.RowItemRM.Required_Qty
                Row.Qty_WIP = Me.RowItemRM.WIP_Qty
                Row.WO_Number = RowItemWO.WO_Number
                Row.Primary_Process = RowItemWO.Primary_Process

                Me.TBL.Rows.Add(Row)

                Me.TblItemBom.DefaultView.RowFilter = "[Identifier BOM] = '" + Me.RowItemBom.Identifier_BOM_Parent + "'"
                If Me.TblItemBom.DefaultView.Count > 0 Then
                    RowItemBom = Me.TblItemBom.DefaultView(0).Row
                    Me.LoadTBL_RM(Me.RowItemBom.Identifier_JO, Me.RowItemBom.Identifier_WO, Row.Primary_Product)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DgvItemBundling_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvItemBundling.SelectionChanged
        Dim rowDGV As DataGridViewRow

        Try
            If Me.DgvItemBundling.SelectedCells.Count = 0 Then
                Me.BufferIndex = -1
                Me.TxtExplain.Text = ""
            Else
                If Me.DgvItemBundling.SelectedCells(0).RowIndex <> Me.BufferIndex Then
                    'explain qty
                    rowDGV = Me.DgvItemBundling.Rows(Me.DgvItemBundling.SelectedCells(0).RowIndex)
                    Me.TxtExplain.Text = ""
                    Me.TxtExplain1.Text = ""

                    Me.TxtExplain.Text = "WO Number : " + rowDGV.Cells(Me.ColDocNo.Index).Value
                    Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "Process :" + rowDGV.Cells(Me.ColProcess.Index).Value
                    Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "Code No : " + rowDGV.Cells(Me.ColCodeNo.Index).Value
                    Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "Description : " + rowDGV.Cells(Me.ColDesc.Index).Value
                    Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "Target Qty : " + Me.ObjCD.DigitGrouping(rowDGV.Cells(Me.ColQty.Index).Value, rowDGV.Cells(Me.ColDecimal_Product.Index).Value) + " " + rowDGV.Cells(Me.ColUnit.Index).Value + "  --> Qty Target"
                    Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + vbCrLf + "Bill of Material Data:"
                    Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "********************************"
                    Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "Good Percentage Process = " + Me.ObjCD.DigitGrouping(rowDGV.Cells(Me.ColGPProcess.Index).Value, 2) + " %  --> GP1"
                    Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "Good Percentage Raw Material = " + Me.ObjCD.DigitGrouping(rowDGV.Cells(Me.ColGPRM.Index).Value, 2) + " %  --> GP2"
                    Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "Qty Reference Target = " + Me.ObjCD.DigitGrouping(rowDGV.Cells(Me.ColQtyRef.Index).Value, rowDGV.Cells(Me.ColDecimal_Product.Index).Value) + _
                    " " + rowDGV.Cells(Me.ColUnit.Index).Value + " --> Qty Ref"
                    Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "Qty Reference Raw Material = " + Me.ObjCD.DigitGrouping(rowDGV.Cells(Me.ColQtyRefRM.Index).Value, rowDGV.Cells(Me.ColDecimalRM.Index).Value) + " " + _
                    rowDGV.Cells(Me.ColRMUnit.Index).Value + "  --> QtyRefRM"
                    Me.TxtExplain.Text = Me.TxtExplain.Text + vbCrLf + "Manufacturing " + Me.ObjCD.DigitGrouping(rowDGV.Cells(Me.ColQtyRef.Index).Value, rowDGV.Cells(Me.ColDecimal_Product.Index).Value) + _
                    " " + rowDGV.Cells(Me.ColUnit.Index).Value + " " + rowDGV.Cells(Me.ColCodeNo.Index).Value + " needs " + _
                    Me.ObjCD.DigitGrouping(rowDGV.Cells(Me.ColQtyRefRM.Index).Value, rowDGV.Cells(Me.ColDecimalRM.Index).Value) + " " + _
                    rowDGV.Cells(Me.ColRMUnit.Index).Value + " " + rowDGV.Cells(Me.ColRMCode.Index).Value

                    Me.TxtExplain1.Text = Me.TxtExplain1.Text + "Qty Required Calculation:"
                    Me.TxtExplain1.Text = Me.TxtExplain1.Text + vbCrLf + "********************************"
                    'Me.TxtExplain1.Text = Me.TxtExplain1.Text + vbCrLf + "Qty Required = ((((Qty Target * QtyRefRM)/Qty Ref)/GP1)/GP2)"
                    Me.TxtExplain1.Text = Me.TxtExplain1.Text + vbCrLf + "Qty Required = RoundUP((((Qty Target /Qty Ref)* QtyRefRM)/GP1)/GP2)"
                    Me.TxtExplain1.Text = Me.TxtExplain1.Text + vbCrLf + "Qty Required = RoundUP((((" + Me.ObjCD.DigitGrouping(rowDGV.Cells(Me.ColQty.Index).Value, rowDGV.Cells(Me.ColDecimal_Product.Index).Value) + _
                    " / " + Me.ObjCD.DigitGrouping(rowDGV.Cells(Me.ColQtyRef.Index).Value, rowDGV.Cells(Me.ColDecimal_Product.Index).Value) + ")" + _
                    " * " + Me.ObjCD.DigitGrouping(rowDGV.Cells(Me.ColQtyRefRM.Index).Value, rowDGV.Cells(Me.ColDecimalRM.Index).Value) + ")" + " / " + _
                    Me.ObjCD.DigitGrouping(rowDGV.Cells(Me.ColGPProcess.Index).Value, 2) + "%) / " + Me.ObjCD.DigitGrouping(rowDGV.Cells(Me.ColGPRM.Index).Value, 2) + "%)"
                    Me.TxtExplain1.Text = Me.TxtExplain1.Text + vbCrLf + "Qty Required = " + Me.ObjCD.DigitGrouping(rowDGV.Cells(Me.ColRMQTY.Index).Value, rowDGV.Cells(Me.ColDecimalRM.Index).Value) + " " + _
                    rowDGV.Cells(Me.ColRMUnit.Index).Value
                    'Me.TxtExplain1.Text = Me.TxtExplain1.Text + vbCrLf + "RUP = Round Up using raw material decimal point"

                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class