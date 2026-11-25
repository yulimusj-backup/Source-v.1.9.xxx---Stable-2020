Option Explicit On
Option Strict Off

Public Class FrmTrans_MaterialPlanningDashboard_Data1_Explain
    Public TBLItem_Invoice As New DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_invoiceDataTable
    Public TBLItem_SR As New DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_srDataTable
    Public TBLItem_Bundling As New DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_bundlingDataTable
    Public TBLItem_Adjustment As New DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_adjustmentDataTable
    Public BackColorAll As Color = Color.LightBlue
    Public CodeNo, Desc, Unit As String
    Public vPrimaryProduct As Double = -1

    Dim rowItem_Invoice As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_invoiceRow
    Dim rowItem_SR As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_srRow
    Dim rowItem_Bundling As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_bundlingRow
    Dim rowItem_Adjustment As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_adjustmentRow
    Dim ObjProd As New Core_Module.clsMasterProduct
    Dim DecimalProduct As Integer = 0
    Dim ObjCD As New Core_Module.ClsCekData
    Dim DecimalIDR As Integer = 0
    Dim ObjCur As New Core_Module.clsMasterCurrency

    Private Sub FrmTrans_MaterialPlanningDashboard_Data1_Explain_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim rowDGV As DataGridViewRow
        Dim DV As Data.DataRowView
        Dim DTA_TipeAdj As New DseTransaksiMaterialPlanningGroupTableAdapters.tbl_setting_adjustmentTableAdapter
        Dim TBL_TipeAdj As New DseTransaksiMaterialPlanningGroup.tbl_setting_adjustmentDataTable
        Dim row_TipeAdj As DseTransaksiMaterialPlanningGroup.tbl_setting_adjustmentRow
        Dim TotalQtySales, TotalAdj, TotalBundling, TotalSalesAmount, TotalGrossMargin As Decimal

        Try
            TotalAdj = 0
            TotalBundling = 0
            TotalGrossMargin = 0
            TotalQtySales = 0
            TotalSalesAmount = 0

            DTA_TipeAdj.Connection.ConnectionString = Core_Module.GetConstring
            DecimalProduct = Me.ObjProd.GetAngkaDecimal(Me.vPrimaryProduct)
            DecimalIDR = Me.ObjCur.GetDecimal_DefaultCurrency

            Me.dgvItemBundling.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvItemAdjustment.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvItemInvoice.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvItemSR.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            Me.Text = "Detail data: " + Me.CodeNo + " - " + Me.Desc
            Me.TxtUnitCode.Text = Me.Unit

            Me.dgvItemAdjustment.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItemBundling.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.dgvItemInvoice.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItemSR.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            Me.TBLItem_Adjustment.DefaultView.RowFilter = "[primary code no] = " + vPrimaryProduct.ToString
            Me.TBLItem_Adjustment.DefaultView.Sort = "[adj date historical] ASC, [adj no historical] ASC"
            Me.TBLItem_Bundling.DefaultView.RowFilter = "[primary code no] = " + vPrimaryProduct.ToString
            Me.TBLItem_Bundling.DefaultView.Sort = "[bundling date historical] ASC, [bundling date historical] ASC"
            Me.TBLItem_Invoice.DefaultView.RowFilter = "[primary code no] = " + vPrimaryProduct.ToString
            Me.TBLItem_Invoice.DefaultView.Sort = "[invoice date historical] ASC, [invoice no historical] ASC"
            Me.TBLItem_SR.DefaultView.RowFilter = "[primary code no] = " + vPrimaryProduct.ToString
            Me.TBLItem_SR.DefaultView.Sort = "[sr date historical] ASC, [sr no historical] ASC"

            'isikan data adjustment
            If TBLItem_Adjustment.DefaultView.Count > 0 Then
                DTA_TipeAdj.Fill(TBL_TipeAdj)
            End If
            For Each DV In Me.TBLItem_Adjustment.DefaultView
                Me.rowItem_Adjustment = DV.Row
                Me.dgvItemAdjustment.Rows.Add()
                rowDGV = Me.dgvItemAdjustment.Rows(Me.dgvItemAdjustment.Rows.Count - 1)
                With rowDGV
                    .Cells(Me.ColAdjustment_Date.Index).Value = rowItem_Adjustment.adj_date_historical
                    .Cells(Me.ColAdjustment_DocumentNo.Index).Value = rowItem_Adjustment.adj_no_historical
                    .Cells(Me.ColAdjustment_No.Index).Value = rowDGV.Index + 1
                    .Cells(Me.ColAdjustment_PrimaryTipeAdjustment.Index).Value = rowItem_Adjustment.primary_tipe_adjustment
                    .Cells(Me.ColAdjustment_Qty.Index).Value = Me.ObjCD.DigitGrouping(rowItem_Adjustment.qty, DecimalProduct)
                    If rowItem_Adjustment.primary_tipe_adjustment = -1 Then
                        .Cells(Me.ColAdjustment_TipeAdjustment.Index).Value = ""
                    Else
                        TBL_TipeAdj.DefaultView.RowFilter = "primary = " + rowItem_Adjustment.primary_tipe_adjustment.ToString
                        If TBL_TipeAdj.DefaultView.Count = 0 Then
                            .Cells(Me.ColAdjustment_TipeAdjustment.Index).Value = ""
                        Else
                            row_TipeAdj = TBL_TipeAdj.DefaultView(0).Row
                            .Cells(Me.ColAdjustment_TipeAdjustment.Index).Value = row_TipeAdj.Desc
                        End If
                    End If
                    TotalAdj += rowItem_Adjustment.qty
                End With
            Next
            Me.ColAdjustment_Date.DefaultCellStyle.Format = "dd MMM yyyy"

            'bundling
            For Each DV In Me.TBLItem_Bundling.DefaultView
                Me.rowItem_Bundling = DV.Row
                Me.dgvItemBundling.Rows.Add()
                rowDGV = Me.dgvItemBundling.Rows(Me.dgvItemBundling.Rows.Count - 1)
                With rowDGV
                    .Cells(Me.ColBundling_Date.Index).Value = rowItem_Bundling.bundling_date_historical
                    .Cells(Me.ColBundling_DocumentNo.Index).Value = rowItem_Bundling.bundling_no_historical
                    .Cells(Me.ColBundling_No.Index).Value = rowDGV.Index + 1
                    .Cells(Me.ColBundling_Qty.Index).Value = Me.ObjCD.DigitGrouping(rowItem_Bundling.qty, DecimalProduct)
                    TotalBundling += rowItem_Bundling.qty
                End With
            Next
            Me.ColBundling_Date.DefaultCellStyle.Format = "dd MMM yyyy"

            'sales invoice
            For Each DV In Me.TBLItem_Invoice.DefaultView
                Me.rowItem_Invoice = DV.Row
                Me.dgvItemInvoice.Rows.Add()
                rowDGV = Me.dgvItemInvoice.Rows(Me.dgvItemInvoice.Rows.Count - 1)
                With rowDGV
                    .Cells(Me.ColSales_Cost.Index).Value = Me.ObjCD.DigitGrouping(rowItem_Invoice.cost, Me.DecimalIDR)
                    .Cells(Me.ColSales_Customer.Index).Value = rowItem_Invoice.customer_name_historical
                    .Cells(Me.ColSales_Date.Index).Value = rowItem_Invoice.invoice_date_historical
                    .Cells(Me.ColSales_DocumentNo.Index).Value = rowItem_Invoice.invoice_no_historical
                    .Cells(Me.ColSales_GrossMargin.Index).Value = Me.ObjCD.DigitGrouping(rowItem_Invoice.sales_amount - rowItem_Invoice.cost, DecimalIDR)
                    .Cells(Me.ColSales_No.Index).Value = rowDGV.Index + 1
                    .Cells(Me.ColSales_Qty.Index).Value = Me.ObjCD.DigitGrouping(rowItem_Invoice.qty, DecimalProduct)
                    .Cells(Me.ColSales_SalesAmount.Index).Value = Me.ObjCD.DigitGrouping(rowItem_Invoice.sales_amount, DecimalIDR)
                End With
                TotalGrossMargin += (rowItem_Invoice.sales_amount - rowItem_Invoice.cost)
                TotalQtySales += rowItem_Invoice.qty
                TotalSalesAmount += rowItem_Invoice.sales_amount
            Next
            Me.ColSales_Date.DefaultCellStyle.Format = "dd MMM yyyy"

            'sales return
            For Each DV In Me.TBLItem_SR.DefaultView
                Me.rowItem_SR = DV.Row
                Me.dgvItemSR.Rows.Add()
                rowDGV = Me.dgvItemSR.Rows(Me.dgvItemSR.Rows.Count - 1)
                With rowDGV
                    .Cells(Me.ColSalesReturn_Amount.Index).Value = Me.ObjCD.DigitGrouping(rowItem_SR.sales_amount, DecimalIDR)
                    .Cells(Me.ColSalesReturn_Cost.Index).Value = Me.ObjCD.DigitGrouping(rowItem_SR.cost, DecimalIDR)
                    .Cells(Me.ColSalesReturn_Customer.Index).Value = rowItem_SR.customer_name_historical
                    .Cells(Me.ColSalesReturn_Date.Index).Value = rowItem_SR.sr_date_historical
                    .Cells(Me.ColSalesReturn_DocumentNo.Index).Value = rowItem_SR.sr_no_historical
                    .Cells(Me.ColSalesReturn_GrossMargin.Index).Value = Me.ObjCD.DigitGrouping(rowItem_SR.sales_amount - rowItem_SR.cost, DecimalIDR)
                    .Cells(Me.ColSalesReturn_No.Index).Value = rowDGV.Index + 1
                    .Cells(Me.ColSalesReturn_Qty.Index).Value = Me.ObjCD.DigitGrouping(rowItem_SR.qty, DecimalProduct)
                End With
                TotalGrossMargin -= (rowItem_SR.sales_amount - rowItem_SR.cost)
                TotalQtySales -= rowItem_SR.qty
                TotalSalesAmount -= rowItem_SR.sales_amount
            Next
            Me.ColSalesReturn_Date.DefaultCellStyle.Format = "dd MMM yyyy"

            'total
            Me.TxtUnitCode.Text = Me.Unit
            Me.TxtTotalQty.Text = Me.ObjCD.DigitGrouping(TotalAdj + TotalBundling + TotalQtySales, DecimalProduct)
            Me.TxtTotalGrossMargin.Text = Me.ObjCD.DigitGrouping(TotalGrossMargin, DecimalIDR)
            Me.TxtSalesQty.Text = Me.ObjCD.DigitGrouping(TotalQtySales, DecimalProduct)
            Me.TxtSalesAmount.Text = Me.ObjCD.DigitGrouping(TotalSalesAmount, DecimalIDR)
            Me.TxtBundlingQty.Text = Me.ObjCD.DigitGrouping(TotalBundling, DecimalProduct)
            Me.TxtAdjustmentQty.Text = Me.ObjCD.DigitGrouping(TotalAdj, DecimalProduct)

            'label set zero
            If TotalQtySales + TotalAdj + TotalBundling < 0 Then
                Me.LblCalcZero_Qty.Visible = True
            Else
                Me.LblCalcZero_Qty.Visible = False
            End If
            If TotalGrossMargin < 0 Then
                Me.LblCalcZero_GrossMargin.Visible = True
            Else
                Me.LblCalcZero_GrossMargin.Visible = False
            End If
            If TotalSalesAmount < 0 Then
                Me.LblCalcZero_SalesAmount.Visible = True
            Else
                Me.LblCalcZero_SalesAmount.Visible = False
            End If

            Me.dgvItemAdjustment.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvItemBundling.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvItemInvoice.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Me.dgvItemSR.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_TipeAdj.Dispose()
            TBL_TipeAdj.Dispose()
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
End Class