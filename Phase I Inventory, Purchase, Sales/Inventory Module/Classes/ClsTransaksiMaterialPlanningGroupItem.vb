Option Explicit On
Option Strict Off

Public Class ClsTransaksiMaterialPlanningGroupItem
    Public Function GetProductLevel(ByVal vProductGroupLevelStr As String) As ClsTransaksiMaterialPlanningGroup.ProductLevel
        Try
            If vProductGroupLevelStr = "Product Group" Then
                Return ClsTransaksiMaterialPlanningGroup.ProductLevel.ProductGroup
            ElseIf vProductGroupLevelStr = "Brand" Then
                Return ClsTransaksiMaterialPlanningGroup.ProductLevel.Brand
            ElseIf vProductGroupLevelStr = "Sub Brand" Then
                Return ClsTransaksiMaterialPlanningGroup.ProductLevel.SubBrand
            ElseIf vProductGroupLevelStr = "Category" Then
                Return ClsTransaksiMaterialPlanningGroup.ProductLevel.Category
            Else
                Return ClsTransaksiMaterialPlanningGroup.ProductLevel.Product
            End If

            Return ClsTransaksiMaterialPlanningGroup.ProductLevel.Product
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ClsTransaksiMaterialPlanningGroup.ProductLevel.Product
        End Try
    End Function

    Public Function GetProductLevelStr(ByVal vProductGroupLevel As ClsTransaksiMaterialPlanningGroup.ProductLevel) As String
        Try
            If vProductGroupLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.ProductGroup Then
                Return "Product Group"
            ElseIf vProductGroupLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Brand Then
                Return "Brand"
            ElseIf vProductGroupLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.SubBrand Then
                Return "Sub Brand"
            ElseIf vProductGroupLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Category Then
                Return "Category"
            Else
                Return "Product"
            End If

            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function GetProductLevelStr(ByVal vProductGroupLevelStr As String) As ClsTransaksiMaterialPlanningGroup.ProductLevel
        Try
            If vProductGroupLevelStr = "Product Group" Then
                Return ClsTransaksiMaterialPlanningGroup.ProductLevel.ProductGroup
            ElseIf vProductGroupLevelStr = "Brand" Then
                Return ClsTransaksiMaterialPlanningGroup.ProductLevel.Brand
            ElseIf vProductGroupLevelStr = "Sub Brand" Then
                Return ClsTransaksiMaterialPlanningGroup.ProductLevel.SubBrand
            ElseIf vProductGroupLevelStr = "Category" Then
                Return ClsTransaksiMaterialPlanningGroup.ProductLevel.Category
            Else
                Return ClsTransaksiMaterialPlanningGroup.ProductLevel.Product
            End If

            Return ClsTransaksiMaterialPlanningGroup.ProductLevel.Product
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ClsTransaksiMaterialPlanningGroup.ProductLevel.Product
        End Try
    End Function

    Public Function GetProcurementStrategyStr(ByVal vProcureMentStrategy As Core_Module.clsMasterProduct.EnumProcurementStrategy) As String
        Try
            If vProcureMentStrategy = Core_Module.clsMasterProduct.EnumProcurementStrategy.Build_To_Order Then
                Return "Build To Order"
            ElseIf vProcureMentStrategy = Core_Module.clsMasterProduct.EnumProcurementStrategy.Build_To_Stock Then
                Return "Build To Stock"
            End If

            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function GetProcurementStrategy(ByVal vProcurementStrategyStr As String) As Core_Module.clsMasterProduct.EnumProcurementStrategy
        Try
            If vProcurementStrategyStr = "Build To Order" Then
                Return Core_Module.clsMasterProduct.EnumProcurementStrategy.Build_To_Order
            ElseIf vProcurementStrategyStr = "Build To Stock" Then
                Return Core_Module.clsMasterProduct.EnumProcurementStrategy.Build_To_Stock
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetLeadTimeStrategy(ByVal vLeadStrategyStr As String) As ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy
        Try
            If vLeadStrategyStr = "Based On Usage" Then
                Return ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.Based_on_usage
            ElseIf vLeadStrategyStr = "Based On Supplier" Then
                Return ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.Based_on_prefered_supplier
            ElseIf vLeadStrategyStr = "User Defined" Then
                Return ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.User_Defined
            End If

            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetLeadTimeStrategyStr(ByVal vLeadTimeStrategy As ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy) As String
        Try
            If vLeadTimeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.Based_on_usage Then
                Return "Based On Usage"
            ElseIf vLeadTimeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.Based_on_prefered_supplier Then
                Return "Based On Supplier"
            Else
                Return "User Defined"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Sub IsiDGV_ItemGroup(ByVal vPrimary As Double, ByVal vDGVItemGroup As DataGridView, ByVal vFrm As FrmTrans_MaterialPlanningDashboard_Data1, _
    ByVal vTipeTransaksi As ClsTransaksiMaterialPlanningGroup.EnumTipe, ByVal ObjCD As Core_Module.ClsCekData)
        Dim DTAGroup As New DseTransaksiMaterialPlanningGroupTableAdapters.tbltrans_material_planning_group_item_groupTableAdapter
        Dim TBLGroup As New DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_groupDataTable
        Dim rowGroup As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_groupRow
        Dim AllowToAddRows_ItemGroup As Boolean
        Dim rowDGV As DataGridViewRow
        Dim DefPrec As Integer = 0
        Dim ObjCur As New Core_Module.clsMasterCurrency

        Try
            DTAGroup.Connection.ConnectionString = Core_Module.GetConstring
            vDGVItemGroup.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItemGroup.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            vDGVItemGroup.Rows.Clear()
            AllowToAddRows_ItemGroup = vDGVItemGroup.AllowUserToAddRows
            vDGVItemGroup.AllowUserToAddRows = False

            If vTipeTransaksi = ClsTransaksiMaterialPlanningGroup.EnumTipe.Gross_Margin_Cumulative_Percentage Or _
            vTipeTransaksi = ClsTransaksiMaterialPlanningGroup.EnumTipe.Gross_Margin_Percentage Or _
            vTipeTransaksi = ClsTransaksiMaterialPlanningGroup.EnumTipe.Nett_Sales_Cumulative_Percentage Or _
            vTipeTransaksi = ClsTransaksiMaterialPlanningGroup.EnumTipe.Nett_Sales_Cumulative_Percentage Or _
            vTipeTransaksi = ClsTransaksiMaterialPlanningGroup.EnumTipe.Nett_Sales_Percentage Or _
            vTipeTransaksi = ClsTransaksiMaterialPlanningGroup.EnumTipe.Usage_Qty_Cumulative_Percentage Or _
            vTipeTransaksi = ClsTransaksiMaterialPlanningGroup.EnumTipe.Usage_Qty_Percentage Then
                DefPrec = 2
            ElseIf vTipeTransaksi = ClsTransaksiMaterialPlanningGroup.EnumTipe.Gross_Margin_Minimum_Value Or _
            vTipeTransaksi = ClsTransaksiMaterialPlanningGroup.EnumTipe.Nett_Sales_Minimum_Value Then
                DefPrec = ObjCur.GetDecimal_DefaultCurrency
            Else
                DefPrec = -1
            End If


            DTAGroup.FillBy_PrimaryTransaksi(TBLGroup, vPrimary)
            For Each rowGroup In TBLGroup.Rows
                vDGVItemGroup.Rows.Add()
                rowDGV = vDGVItemGroup.Rows(vDGVItemGroup.Rows.Count - 1)
                With rowGroup
                    rowDGV.Cells(vFrm.ColGroup_Code.Index).Value = .group_code
                    rowDGV.Cells(vFrm.ColGroup_Desc.Index).Value = .group_desc
                    rowDGV.Cells(vFrm.ColGroup_LeadTimeDays.Index).Value = ObjCD.BuangNolDibelakangDecimal(.lead_time_days)
                    rowDGV.Cells(vFrm.ColGroup_LeadTimeQty.Index).Value = ObjCD.DigitGrouping(ObjCD.BuangNolDibelakangDecimal(.lead_time_qty))
                    rowDGV.Cells(vFrm.ColGroup_LeadTimeStrategy.Index).Value = Me.GetLeadTimeStrategyStr(.lead_time_strategy)
                    rowDGV.Cells(vFrm.ColGroup_No.Index).Value = rowDGV.Index + 1
                    rowDGV.Cells(vFrm.ColGroup_OptimumStockDays.Index).Value = ObjCD.BuangNolDibelakangDecimal(.optimum_stock_days)
                    rowDGV.Cells(vFrm.ColGroup_OptimumStockQty.Index).Value = ObjCD.DigitGrouping(ObjCD.BuangNolDibelakangDecimal(.optimum_stock_qty))
                    rowDGV.Cells(vFrm.ColGroup_OptimumStockTrategy.Index).Value = Me.GetLeadTimeStrategyStr(.optimum_stock_strategy)
                    rowDGV.Cells(vFrm.ColGroup_ProcurementStrategy.Index).Value = Me.GetProcurementStrategyStr(.procurement_strategy)
                    If DefPrec = -1 Then
                        rowDGV.Cells(vFrm.ColGroup_RefValue.Index).Value = ObjCD.BuangNolDibelakangDecimal(ObjCD.DigitGrouping(.ref_value, 6))
                    Else
                        rowDGV.Cells(vFrm.ColGroup_RefValue.Index).Value = ObjCD.DigitGrouping(.ref_value, DefPrec)
                    End If
                    rowDGV.Cells(vFrm.ColGroup_SafetyStockDays.Index).Value = ObjCD.BuangNolDibelakangDecimal(.safety_stock_days)
                    rowDGV.Cells(vFrm.ColGroup_SafetyStockQty.Index).Value = ObjCD.DigitGrouping(ObjCD.BuangNolDibelakangDecimal(.safety_stock_qty))
                    rowDGV.Cells(vFrm.ColGroup_SafetyStockStrategy.Index).Value = Me.GetLeadTimeStrategyStr(.safety_stock_strategy)
                    rowDGV.Cells(vFrm.ColGroup_Remarks.Index).Value = .remarks
                    rowDGV.Cells(vFrm.ColGroup_Identifier.Index).Value = .identifier
                End With
            Next

            vDGVItemGroup.AllowUserToAddRows = AllowToAddRows_ItemGroup
            vDGVItemGroup.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAGroup.Dispose()
            TBLGroup.Dispose()
            ObjCur = Nothing
        End Try
    End Sub

    Public Sub IsiDGV(ByVal vPrimary As Double, ByVal vDGVItemSupplier As DataGridView, ByVal vDGVItemGroup As DataGridView, ByVal vDGVItemProduct As DataGridView, _
    ByVal vDGVItemProductGroup As DataGridView, ByVal vFrm As FrmTrans_MaterialPlanningDashboard_Data1, ByVal vTipeTransaksi As ClsTransaksiMaterialPlanningGroup.EnumTipe, _
    ByVal TBLItem_Invoice As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_invoiceDataTable, _
    ByVal TBLItem_SR As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_srDataTable, _
    ByVal TBLItem_Bundling As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_bundlingDataTable, _
    ByVal TBLItem_Adjustment As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_adjustmentDataTable)
        Dim DTAProduct As New DseTransaksiMaterialPlanningGroupTableAdapters.tbltrans_material_planning_group_item_product1TableAdapter
        Dim TBLProduct As New DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_product1DataTable
        Dim rowProduct As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_product1Row
        Dim DTASupplier As New DseTransaksiMaterialPlanningGroupTableAdapters.tbltrans_material_planning_group_itemsupplier1TableAdapter
        Dim TBLSupplier As New DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_itemsupplier1DataTable
        Dim rowSupplier As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_itemsupplier1Row
        Dim DTAProductGroup As New DseTransaksiMaterialPlanningGroupTableAdapters.tbltrans_material_planning_group_item_product_groupTableAdapter
        Dim TBLProductGroup As New DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_product_groupDataTable
        Dim rowProductGroup As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_product_groupRow
        Dim DTAProduct2 As New DseTransaksiMaterialPlanningGroupTableAdapters.TblProductTableAdapter
        Dim TBLProduct2 As New DseTransaksiMaterialPlanningGroup.TblProductDataTable
        Dim rowProduct2 As DseTransaksiMaterialPlanningGroup.TblProductRow
        Dim Filter As String = ""
        Dim DTACategory As New DseTransaksiMaterialPlanningGroupTableAdapters.tblmaster_categoryTableAdapter
        Dim TBLCategory As New DseTransaksiMaterialPlanningGroup.tblmaster_categoryDataTable
        Dim rowCategory As DseTransaksiMaterialPlanningGroup.tblmaster_categoryRow
        Dim DTAProductGroup2 As New DseTransaksiMaterialPlanningGroupTableAdapters.tblmaster_product_groupTableAdapter
        Dim TblProductGroup2 As New DseTransaksiMaterialPlanningGroup.tblmaster_product_groupDataTable
        Dim rowProductGroup2 As DseTransaksiMaterialPlanningGroup.tblmaster_product_groupRow

        Dim AllowToAddRows_ItemProductGroup As Boolean
        Dim AllowToAddRows_ItemSupplier As Boolean
        'Dim AllowToAddRows_ItemGroup As Boolean

        Dim rowDGV As DataGridViewRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim LeadTimeStrategy As ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy = ClsTransaksiMaterialPlanningGroup.EnumLeadTimeStrategy.Based_on_prefered_supplier
        Dim DefPrec As Integer = 0
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim DefPrecPerc As Integer = 2
        Dim ProductLevel As ClsTransaksiMaterialPlanningGroup.ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Brand

        Dim DTAItemInvoice1 As New DseTransaksiMaterialPlanningGroupTableAdapters.tbltrans_material_planning_group_item_invoice1TableAdapter
        Dim TBLItemInvoice1 As New DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_invoice1DataTable
        Dim rowItemInvoice1 As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_invoice1Row
        Dim rowItemInvoice As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_invoiceRow
        Dim DTAItemSR1 As New DseTransaksiMaterialPlanningGroupTableAdapters.tbltrans_material_planning_group_item_sr1TableAdapter
        Dim TBLitemSR1 As New DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_sr1DataTable
        Dim rowItemSR1 As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_sr1Row
        Dim rowItemSR As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_srRow
        Dim DTACust As New DseTransaksiMaterialPlanningGroupTableAdapters.tblmaster_customerTableAdapter
        Dim TBLCust As New DseTransaksiMaterialPlanningGroup.tblmaster_customerDataTable
        Dim rowCust As DseTransaksiMaterialPlanningGroup.tblmaster_customerRow
        Dim DTAItemAdjustment1 As New DseTransaksiMaterialPlanningGroupTableAdapters.tbltrans_material_planning_group_item_adjustment1TableAdapter
        Dim TBLItemAdjustment1 As New DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_adjustment1DataTable
        Dim rowItemAdjustment1 As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_adjustment1Row
        Dim TipeStockOutAdjust As clsTransaksiStockInOut.EnumTipeStockInOut = clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust
        Dim rowItemAdjustment As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_adjustmentRow
        Dim DTAItemBundling1 As New DseTransaksiMaterialPlanningGroupTableAdapters.tbltrans_material_planning_group_item_bundling1TableAdapter
        Dim TBLItemBundling1 As New DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_bundling1DataTable
        Dim rowItemBundling1 As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_bundling1Row
        Dim rowItemBundling As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_bundlingRow
        Dim ApakahMasihAda As Boolean = True
        Dim CekIndex As Double = 0

        Try
            TBLCategory.Rows.Clear()
            TblProductGroup2.Rows.Clear()

            DTAProduct.Connection.ConnectionString = Core_Module.GetConstring
            DTAProduct2.Connection.ConnectionString = DTAProduct.Connection.ConnectionString
            DTAProductGroup.Connection.ConnectionString = DTAProduct.Connection.ConnectionString
            DTASupplier.Connection.ConnectionString = DTAProduct.Connection.ConnectionString
            DTAItemInvoice1.Connection.ConnectionString = DTAProduct.Connection.ConnectionString
            DTAItemSR1.Connection.ConnectionString = DTAProduct.Connection.ConnectionString
            DTAItemAdjustment1.Connection.ConnectionString = DTAProduct.Connection.ConnectionString
            DTAItemBundling1.Connection.ConnectionString = DTAProduct.Connection.ConnectionString
            DTACategory.Connection.ConnectionString = DTAProduct.Connection.ConnectionString
            DTAProductGroup2.Connection.ConnectionString = DTAProduct.Connection.ConnectionString
            DTACust.Connection.ConnectionString = DTAProduct.Connection.ConnectionString

            vDGVItemProduct.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItemProduct.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            vDGVItemProductGroup.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItemProductGroup.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            vDGVItemSupplier.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItemSupplier.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            vDGVItemProduct.Rows.Clear()
            vDGVItemProductGroup.Rows.Clear()
            vDGVItemSupplier.Rows.Clear()
            AllowToAddRows_ItemProductGroup = vDGVItemProductGroup.AllowUserToAddRows
            AllowToAddRows_ItemSupplier = vDGVItemSupplier.AllowUserToAddRows
            'AllowToAddRows_ItemGroup = vDGVItemGroup.AllowUserToAddRows
            vDGVItemProductGroup.AllowUserToAddRows = False
            vDGVItemSupplier.AllowUserToAddRows = False
            vDGVItemGroup.AllowUserToAddRows = False

            Me.IsiDGV_ItemGroup(vPrimary, vDGVItemGroup, vFrm, vTipeTransaksi, ObjCD)

            DefPrec = ObjCur.GetDecimal_DefaultCurrency
            DefPrecPerc = 2
            DTAProduct.Fill(TBLProduct, vPrimary)
            For Each rowProduct In TBLProduct.Rows
                vDGVItemProduct.Rows.Add()
                rowDGV = vDGVItemProduct.Rows(vDGVItemProduct.Rows.Count - 1)
                With rowProduct
                    rowDGV.Cells(vFrm.ColCodeNo_AverageGrossMargin.Index).Value = ObjCD.DigitGrouping(.gross_margin_average, DefPrec)
                    rowDGV.Cells(vFrm.ColCodeNo_AverageNettSales.Index).Value = ObjCD.DigitGrouping(.nett_sales_average, DefPrec)
                    rowDGV.Cells(vFrm.ColCodeNo_AverageQty.Index).Value = ObjCD.DigitGrouping(.total_qty_avg, ._Decimal)
                    rowDGV.Cells(vFrm.ColCodeNo_CodeNo.Index).Value = .Product_Code
                    rowDGV.Cells(vFrm.ColCodeNo_Desc.Index).Value = .Description
                    rowDGV.Cells(vFrm.ColCodeNo_GrossMargin.Index).Value = ObjCD.DigitGrouping(.gross_margin, DefPrec)
                    rowDGV.Cells(vFrm.ColCodeNo_GrossMarginPerc.Index).Value = ObjCD.DigitGrouping(.gross_margin_perc, DefPrecPerc)
                    rowDGV.Cells(vFrm.ColCodeNo_LeadTimeDays.Index).Value = ObjCD.BuangNolDibelakangDecimal(.lead_time_days)
                    rowDGV.Cells(vFrm.ColCodeNo_LeadTimeQty.Index).Value = ObjCD.DigitGrouping(.lead_time_qty, ._Decimal)
                    rowDGV.Cells(vFrm.ColCodeNo_LeadTimeStrategy.Index).Value = Me.GetLeadTimeStrategyStr(.lead_time_strategy)
                    rowDGV.Cells(vFrm.ColCodeNo_NettSales.Index).Value = ObjCD.DigitGrouping(.nett_sales, DefPrec)
                    rowDGV.Cells(vFrm.ColCodeNo_NettSalesPerc.Index).Value = ObjCD.DigitGrouping(.nett_sales_perc, DefPrecPerc)
                    rowDGV.Cells(vFrm.ColCodeNo_No.Index).Value = rowDGV.Index + 1
                    rowDGV.Cells(vFrm.ColCodeNo_OptimumStockDays.Index).Value = ObjCD.BuangNolDibelakangDecimal(.optimum_stock_days)
                    rowDGV.Cells(vFrm.ColCodeNo_OptimumStockQty.Index).Value = ObjCD.DigitGrouping(.optimum_stock_qty, DefPrec)
                    rowDGV.Cells(vFrm.ColCodeNo_OptimumStockStrategy.Index).Value = Me.GetLeadTimeStrategyStr(.optimum_stock_strategy)
                    If .override = 0 Then
                        rowDGV.Cells(vFrm.ColCodeNo_Override.Index).Value = Boolean.FalseString
                    Else
                        rowDGV.Cells(vFrm.ColCodeNo_Override.Index).Value = Boolean.TrueString
                    End If
                    rowDGV.Cells(vFrm.ColCodeNo_PrimaryUnit.Index).Value = .primary_unit
                    rowDGV.Cells(vFrm.ColCodeNo_ProcurementStrategy.Index).Value = Me.GetProcurementStrategyStr(.procurement_strategy)
                    rowDGV.Cells(vFrm.ColCodeNo_Qty.Index).Value = ObjCD.DigitGrouping(.total_qty, ._Decimal)
                    rowDGV.Cells(vFrm.ColCodeNo_QtyPerc.Index).Value = ObjCD.DigitGrouping(.total_qty_perc, DefPrecPerc)
                    rowDGV.Cells(vFrm.ColCodeNo_SafetyStockQty.Index).Value = ObjCD.DigitGrouping(.safety_stock_qty, DefPrec)
                    rowDGV.Cells(vFrm.ColCodeNo_SafetyStockStrategy.Index).Value = Me.GetLeadTimeStrategyStr(.safety_stock_strategy)
                    rowDGV.Cells(vFrm.ColCodeNo_SafetyStockStrategyDays.Index).Value = ObjCD.BuangNolDibelakangDecimal(.safety_stock_days)
                    rowDGV.Cells(vFrm.ColCodeNo_Unit.Index).Value = .Unit_Code
                    rowDGV.Cells(vFrm.ColCodeNo_Remarks.Index).Value = .remarks
                    rowDGV.Cells(vFrm.ColCodeNo_PrimaryCodeNo.Index).Value = .primary_code_no
                    rowDGV.Cells(vFrm.ColCodeNo_IdentifierGroup.Index).Value = .identifier
                    rowDGV.Cells(vFrm.ColCodeNo_Group.Index).Value = .group_code + " - " + .group_desc
                End With
            Next

            DTASupplier.Fill(TBLSupplier, vPrimary)
            For Each rowSupplier In TBLSupplier.Rows
                vDGVItemSupplier.Rows.Add()
                rowDGV = vDGVItemSupplier.Rows(vDGVItemSupplier.Rows.Count - 1)
                With rowSupplier
                    rowDGV.Cells(vFrm.ColSupplier_Code.Index).Value = .Supplier_Code
                    rowDGV.Cells(vFrm.ColSupplier_Name.Index).Value = .Supplier_Name
                    rowDGV.Cells(vFrm.ColSupplier_No.Index).Value = rowDGV.Index + 1
                    rowDGV.Cells(vFrm.ColSupplier_Primary.Index).Value = .primary_supplier
                    rowDGV.Cells(vFrm.ColSupplier_Remarks.Index).Value = .remarks
                End With
            Next

            DTAProduct2.Fill(TBLProduct2)
            DTAProductGroup.FillBy_PrimaryTransaksi(TBLProductGroup, vPrimary)
            For Each rowProductGroup In TBLProductGroup.Rows
                vDGVItemProductGroup.Rows.Add()
                rowDGV = vDGVItemProductGroup.Rows(vDGVItemProductGroup.Rows.Count - 1)
                With rowProductGroup
                    rowDGV.Cells(vFrm.ColProductGroup_No.Index).Value = rowDGV.Index + 1
                    rowDGV.Cells(vFrm.ColProductGroup_Remarks.Index).Value = .remarks
                    rowDGV.Cells(vFrm.ColProductGroup_PrimaryAcuan.Index).Value = .primary_acuan
                    ProductLevel = .level
                    rowDGV.Cells(vFrm.ColProductGroup_Level.Index).Value = Me.GetProductLevelStr(ProductLevel)
                    If ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Brand Then
                        Filter = "[Primary Brand] = " + .primary_acuan.ToString
                    ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Category Then
                        Filter = "[Primary Category] = " + .primary_acuan.ToString
                    ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Product Then
                        Filter = "[Primary] = " + .primary_acuan.ToString
                    ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.ProductGroup Then
                        Filter = "[Primary Product Group] = " + .primary_acuan.ToString
                    ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.SubBrand Then
                        Filter = "[Primary Sub Brand] = " + .primary_acuan.ToString
                    End If
                    TBLProduct2.DefaultView.RowFilter = Filter
                    If TBLProduct2.DefaultView.Count = 0 Then
                        If ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Product Then
                            MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                            Exit Sub
                        ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Category Then
                            If TBLCategory.Rows.Count = 0 Then
                                DTACategory.Fill(TBLCategory)
                            End If
                            TBLCategory.DefaultView.RowFilter = "Primary = " + .primary_acuan.ToString
                            If TBLCategory.DefaultView.Count > 0 Then
                                rowCategory = TBLCategory.DefaultView(0).Row
                                rowDGV.Cells(vFrm.ColProductGroup_CodeNo.Index).Value = rowCategory.Category_Code
                                rowDGV.Cells(vFrm.ColProductGroup_Desc.Index).Value = rowCategory.Description
                            Else
                                MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                                Exit Sub
                            End If
                        ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Brand Or _
                        ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.ProductGroup Or _
                        ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.SubBrand Then
                            If TblProductGroup2.Rows.Count = 0 Then
                                DTAProductGroup2.Fill(TblProductGroup2)
                            End If
                            TblProductGroup2.DefaultView.RowFilter = "Primary = " + .primary_acuan.ToString
                            If TblProductGroup2.DefaultView.Count > 0 Then
                                rowProductGroup2 = TblProductGroup2.DefaultView(0).Row
                                rowDGV.Cells(vFrm.ColProductGroup_CodeNo.Index).Value = rowProductGroup2.Group_Code
                                rowDGV.Cells(vFrm.ColProductGroup_Desc.Index).Value = rowProductGroup2.Description
                            Else
                                MsgBox("Logic error, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                                Exit Sub
                            End If
                        End If
                    Else
                        rowProduct2 = TBLProduct2.DefaultView(0).Row
                        If ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Brand Then
                            rowDGV.Cells(vFrm.ColProductGroup_CodeNo.Index).Value = rowProduct2.Brand_Code
                            rowDGV.Cells(vFrm.ColProductGroup_Desc.Index).Value = rowProduct2.Brand_Desc
                        ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Category Then
                            rowDGV.Cells(vFrm.ColProductGroup_CodeNo.Index).Value = rowProduct2.Category_Code
                            rowDGV.Cells(vFrm.ColProductGroup_Desc.Index).Value = rowProduct2.Category_Desc
                        ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Product Then
                            rowDGV.Cells(vFrm.ColProductGroup_CodeNo.Index).Value = rowProduct2.Product_Code
                            rowDGV.Cells(vFrm.ColProductGroup_Desc.Index).Value = rowProduct2.Product_Desc
                        ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.ProductGroup Then
                            rowDGV.Cells(vFrm.ColProductGroup_CodeNo.Index).Value = rowProduct2.Product_Group_Code
                            rowDGV.Cells(vFrm.ColProductGroup_Desc.Index).Value = rowProduct2.Product_Group_Desc
                        ElseIf ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.SubBrand Then
                            rowDGV.Cells(vFrm.ColProductGroup_CodeNo.Index).Value = rowProduct2.Sub_Brand_Code
                            rowDGV.Cells(vFrm.ColProductGroup_Desc.Index).Value = rowProduct2.Sub_Brand_Desc
                        End If
                    End If
                End With
            Next

            vDGVItemProductGroup.AllowUserToAddRows = AllowToAddRows_ItemProductGroup
            vDGVItemSupplier.AllowUserToAddRows = AllowToAddRows_ItemSupplier
            'vDGVItemGroup.AllowUserToAddRows = AllowToAddRows_ItemGroup
            vDGVItemProduct.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVItemProductGroup.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVItemSupplier.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            'adjusment
            DTAItemAdjustment1.Fill(TBLItemAdjustment1, vPrimary, CInt(TipeStockOutAdjust))
            TBLItem_Adjustment.Rows.Clear()
            For Each rowItemAdjustment1 In TBLItemAdjustment1.Rows
                rowItemAdjustment = TBLItem_Adjustment.Newtbltrans_material_planning_group_item_adjustmentRow
                With rowItemAdjustment
                    If rowItemAdjustment1.IsNo_TransaksiNull = True Then
                        .adj_no_historical = rowItemAdjustment1.adj_no_historical
                    Else
                        .adj_no_historical = rowItemAdjustment1.No_Transaksi
                    End If
                    If rowItemAdjustment1.IsTanggal_TransaksiNull = True Then
                        .adj_date_historical = rowItemAdjustment1.adj_date_historical
                    Else
                        .adj_date_historical = rowItemAdjustment1.Tanggal_Transaksi
                    End If
                    .primary = rowItemAdjustment1.primary
                    .primary_adj = rowItemAdjustment1.primary_adj
                    .primary_code_no = rowItemAdjustment1.primary_code_no
                    .primary_transaksi = rowItemAdjustment1.primary_transaksi
                    .primary_unit = rowItemAdjustment1.primary_unit
                    .qty = rowItemAdjustment1.qty
                    .primary_tipe_adjustment = rowItemAdjustment1.primary_tipe_adjustment
                End With
                TBLItem_Adjustment.Rows.Add(rowItemAdjustment)
            Next

            'bundling
            DTAItemBundling1.Fill(TBLItemBundling1, vPrimary)
            TBLItem_Bundling.Rows.Clear()
            For Each rowItemBundling1 In TBLItemBundling1.Rows
                rowItemBundling = TBLItem_Bundling.Newtbltrans_material_planning_group_item_bundlingRow
                With rowItemBundling
                    If rowItemBundling1.IsNo_TransaksiNull = True Then
                        .bundling_no_historical = rowItemBundling1.bundling_no_historical
                    Else
                        .bundling_no_historical = rowItemBundling1.No_Transaksi
                    End If
                    If rowItemBundling1.IsTanggal_TransaksiNull = True Then
                        .bundling_date_historical = rowItemBundling1.bundling_date_historical
                    Else
                        .bundling_date_historical = rowItemBundling1.Tanggal_Transaksi
                    End If
                    .primary = rowItemBundling1.primary
                    .primary_bundling = rowItemBundling1.primary_bundling
                    .primary_code_no = rowItemBundling1.primary_code_no
                    .primary_transaksi = rowItemBundling1.primary_transaksi
                    .primary_unit = rowItemBundling1.primary_unit
                    .qty = rowItemBundling1.qty
                End With
                TBLItem_Bundling.Rows.Add(rowItemBundling)
            Next

            TBLItem_Invoice.Rows.Clear()
            DTAItemInvoice1.Fill(TBLItemInvoice1, vPrimary)
            For Each rowItemInvoice1 In TBLItemInvoice1.Rows
                CekIndex += 1
                ApakahMasihAda = True
                rowItemInvoice = TBLItem_Invoice.Newtbltrans_material_planning_group_item_invoiceRow
                With rowItemInvoice
                    If rowItemInvoice1.IsNo_TransaksiNull = True Then
                        ApakahMasihAda = False
                    End If
                    If ApakahMasihAda = True Then
                        If rowItemInvoice1.primary_customer <> rowItemInvoice1.Primary_Customer_Baru Then
                            ApakahMasihAda = False
                        End If
                    End If
                    If ApakahMasihAda = False Then
                        .invoice_no_historical = rowItemInvoice1.invoice_no_historical
                        .invoice_date_historical = rowItemInvoice1.invoice_date_historical
                        .primary_customer = rowItemInvoice1.primary_customer
                        .customer_code_historical = rowItemInvoice1.customer_code_historical
                        .customer_name_historical = rowItemInvoice1.customer_name_historical
                    Else
                        .invoice_no_historical = rowItemInvoice1.No_Transaksi
                        .invoice_date_historical = rowItemInvoice1.Tanggal_Transaksi
                        .primary_customer = rowItemInvoice1.Primary_Customer_Baru
                        .customer_code_historical = rowItemInvoice1.Kode_Customer_Baru
                        .customer_name_historical = rowItemInvoice1.Nama_Customer_Baru
                    End If
                    .cost = rowItemInvoice1.cost
                    .primary = rowItemInvoice1.primary
                    .primary_code_no = rowItemInvoice1.primary_code_no
                    .primary_invoice = rowItemInvoice1.primary_invoice
                    .primary_transaksi = rowItemInvoice1.primary_transaksi
                    .primary_unit = rowItemInvoice1.primary_unit
                    .qty = rowItemInvoice1.qty
                    .sales_amount = rowItemInvoice1.sales_amount
                End With
                TBLItem_Invoice.Rows.Add(rowItemInvoice)
            Next

            TBLItem_SR.Rows.Clear()
            DTAItemSR1.Fill(TBLitemSR1, vPrimary)
            If TBLitemSR1.Rows.Count > 0 Then
                DTACust.Fill(TBLCust)
            End If
            CekIndex = 0
            For Each rowItemSR1 In TBLitemSR1.Rows
                CekIndex += 1
                rowItemSR = TBLItem_SR.Newtbltrans_material_planning_group_item_srRow
                With rowItemSR
                    .primary_customer = rowItemSR1.primary_customer
                    If rowItemSR1.IsNo_TransaksiNull = True Then
                        rowItemSR.sr_no_historical = rowItemSR1.sr_no_historical
                        rowItemSR.sr_date_historical = rowItemSR1.sr_date_historical
                        .customer_code_historical = rowItemSR1.customer_code_historical
                        .customer_name_historical = rowItemSR1.customer_name_historical
                    Else
                        rowItemSR.sr_no_historical = rowItemSR1.No_Transaksi
                        rowItemSR.sr_date_historical = rowItemSR1.Tanggal_Transaksi

                        TBLCust.DefaultView.RowFilter = "Primary = " + .primary_customer.ToString
                        If TBLCust.DefaultView.Count = 0 Then
                            .customer_code_historical = rowItemSR1.customer_code_historical
                            .customer_name_historical = rowItemSR1.customer_name_historical
                        Else
                            rowCust = TBLCust.DefaultView(0).Row
                            .customer_code_historical = rowCust.Customer_Code
                            .customer_name_historical = rowCust.Customer_Name
                        End If
                    End If
                    .cost = rowItemSR1.cost
                    .primary = rowItemSR1.primary
                    .primary_code_no = rowItemSR1.primary_code_no
                    .primary_sr = rowItemSR1.primary_sr
                    .primary_transaksi = rowItemSR1.primary_transaksi
                    .primary_unit = rowItemSR1.primary_unit
                    .qty = rowItemSR1.qty
                    .sales_amount = rowItemSR1.sales_amount
                End With
                TBLItem_SR.Rows.Add(rowItemSR)
            Next
        Catch ex As Exception
            MsgBox(ex.Message + " - cek index: " + CekIndex.ToString, MsgBoxStyle.Critical, "Error")
        Finally
            DTAProduct.Dispose()
            TBLProduct.Dispose()
            DTASupplier.Dispose()
            TBLSupplier.Dispose()
            DTAProduct2.Dispose()
            TBLProduct2.Dispose()
            DTAProductGroup.Dispose()
            TBLProductGroup.Dispose()
            ObjCD = Nothing
            ObjCur = Nothing
            DTAItemInvoice1.Dispose()
            TBLItemInvoice1.Dispose()
            DTAItemSR1.Dispose()
            TBLitemSR1.Dispose()
            DTACust.Dispose()
            TBLCust.Dispose()
            DTAItemAdjustment1.Dispose()
            TBLItemAdjustment1.Dispose()
            DTAItemBundling1.Dispose()
            TBLItemBundling1.Dispose()
            DTACategory.Dispose()
            TBLCategory.Dispose()
            DTAProductGroup2.Dispose()
            TblProductGroup2.Dispose()
        End Try
    End Sub

    Public Sub Simpan(ByVal vPrimary As Double, ByVal vDGVItemSupplier As DataGridView, ByVal vDGVItemGroup As DataGridView, ByVal vDGVItemProduct As DataGridView, _
    ByVal vDGVItemProductGroup As DataGridView, ByVal vFrm As FrmTrans_MaterialPlanningDashboard_Data1, _
    ByVal TBLItem_Invoice As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_invoiceDataTable, _
    ByVal TBLItem_SR As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_srDataTable, _
    ByVal TBLItem_Bundling As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_bundlingDataTable, _
    ByVal TBLItem_Adjustment As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_adjustmentDataTable)
        Dim DTAGroup As New DseTransaksiMaterialPlanningGroupTableAdapters.tbltrans_material_planning_group_item_groupTableAdapter
        Dim DTASupplier As New DseTransaksiMaterialPlanningGroupTableAdapters.tbltrans_material_planning_group_itemsupplierTableAdapter
        Dim DTAProductGroup As New DseTransaksiMaterialPlanningGroupTableAdapters.tbltrans_material_planning_group_item_product_groupTableAdapter
        Dim DTAProduct As New DseTransaksiMaterialPlanningGroupTableAdapters.tbltrans_material_planning_group_item_productTableAdapter
        Dim rowDGV As DataGridViewRow
        Dim AllowToAddRows_ItemGroup As Boolean
        Dim AllowToAddRows_ItemProductGroup As Boolean
        Dim AllowToAddRows_ItemSupplier As Boolean
        Dim DTAItemInvoice As New DseTransaksiMaterialPlanningGroupTableAdapters.tbltrans_material_planning_group_item_invoiceTableAdapter
        Dim DTAItemSR As New DseTransaksiMaterialPlanningGroupTableAdapters.tbltrans_material_planning_group_item_srTableAdapter
        Dim DTAItemAdjustment As New DseTransaksiMaterialPlanningGroupTableAdapters.tbltrans_material_planning_group_item_adjustmentTableAdapter
        Dim DTAItemBundling As New DseTransaksiMaterialPlanningGroupTableAdapters.tbltrans_material_planning_group_item_bundlingTableAdapter

        Try
            DTAGroup.Connection.ConnectionString = Core_Module.GetConstring
            DTAProduct.Connection.ConnectionString = DTAGroup.Connection.ConnectionString
            DTAProductGroup.Connection.ConnectionString = DTAGroup.Connection.ConnectionString
            DTASupplier.Connection.ConnectionString = DTAGroup.Connection.ConnectionString
            DTAItemInvoice.Connection.ConnectionString = DTAGroup.Connection.ConnectionString
            DTAItemSR.Connection.ConnectionString = DTAGroup.Connection.ConnectionString
            DTAItemAdjustment.Connection.ConnectionString = DTAGroup.Connection.ConnectionString
            DTAItemBundling.Connection.ConnectionString = DTAGroup.Connection.ConnectionString


            AllowToAddRows_ItemGroup = vDGVItemGroup.AllowUserToAddRows
            AllowToAddRows_ItemProductGroup = vDGVItemProductGroup.AllowUserToAddRows
            AllowToAddRows_ItemSupplier = vDGVItemSupplier.AllowUserToAddRows
            vDGVItemGroup.AllowUserToAddRows = False
            vDGVItemProductGroup.AllowUserToAddRows = False
            vDGVItemSupplier.AllowUserToAddRows = False

            'group
            For Each rowDGV In vDGVItemGroup.Rows
                DTAGroup.Insert(vPrimary, rowDGV.Cells(vFrm.ColGroup_Code.Index).Value, rowDGV.Cells(vFrm.ColGroup_Desc.Index).Value, _
                rowDGV.Cells(vFrm.ColGroup_RefValue.Index).Value, CInt(Me.GetProcurementStrategy(rowDGV.Cells(vFrm.ColGroup_ProcurementStrategy.Index).Value)), _
                CInt(Me.GetLeadTimeStrategy(rowDGV.Cells(vFrm.ColGroup_OptimumStockTrategy.Index).Value)), _
                rowDGV.Cells(vFrm.ColGroup_OptimumStockDays.Index).Value, rowDGV.Cells(vFrm.ColGroup_OptimumStockQty.Index).Value, _
                CInt(Me.GetLeadTimeStrategy(rowDGV.Cells(vFrm.ColGroup_SafetyStockStrategy.Index).Value)), _
                rowDGV.Cells(vFrm.ColGroup_SafetyStockDays.Index).Value, rowDGV.Cells(vFrm.ColGroup_SafetyStockQty.Index).Value, _
                CInt(Me.GetLeadTimeStrategy(rowDGV.Cells(vFrm.ColGroup_LeadTimeStrategy.Index).Value)), _
                rowDGV.Cells(vFrm.ColGroup_LeadTimeDays.Index).Value, rowDGV.Cells(vFrm.ColGroup_LeadTimeQty.Index).Value, _
                rowDGV.Cells(vFrm.ColGroup_Remarks.Index).Value, rowDGV.Cells(vFrm.ColGroup_Identifier.Index).Value)
            Next

            'supplier
            For Each rowDGV In vDGVItemSupplier.Rows
                DTASupplier.Insert(vPrimary, rowDGV.Cells(vFrm.ColSupplier_Primary.Index).Value, rowDGV.Cells(vFrm.ColSupplier_Remarks.Index).Value)
            Next

            'product group
            For Each rowDGV In vDGVItemProductGroup.Rows
                DTAProductGroup.Insert(vPrimary, Me.GetProductLevel(rowDGV.Cells(vFrm.ColProductGroup_Level.Index).Value), _
                rowDGV.Cells(vFrm.ColProductGroup_PrimaryAcuan.Index).Value, rowDGV.Cells(vFrm.ColProductGroup_Remarks.Index).Value)
            Next

            'product
            Me.Simpan_Item_UsingScript(vDGVItemProduct, vPrimary, vFrm)

            'item invoice
            Me.Simpan_ItemInvoice_UsingScript(TBLItem_Invoice, vPrimary)
            'item sr
            Me.Simpan_ItemSR_UsingScript(TBLItem_SR, vPrimary)
            'item bundling
            Me.Simpan_ItemBubdling_UsingScript(TBLItem_Bundling, vPrimary)
            'item adjustment
            Me.Simpan_ItemAdjustment_UsingScript(TBLItem_Adjustment, vPrimary)

            vDGVItemGroup.AllowUserToAddRows = AllowToAddRows_ItemGroup
            vDGVItemProductGroup.AllowUserToAddRows = AllowToAddRows_ItemProductGroup
            vDGVItemSupplier.AllowUserToAddRows = AllowToAddRows_ItemSupplier
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAGroup.Dispose()
            DTASupplier.Dispose()
            DTAProductGroup.Dispose()
            DTAProduct.Dispose()
            DTAItemInvoice.Dispose()
            DTAItemSR.Dispose()
            DTAItemAdjustment.Dispose()
            DTAItemBundling.Dispose()
        End Try
    End Sub

    Private Function Simpan_Item_UsingScript(ByVal vDGV As DataGridView, ByVal vPrimaryTrans As Double, ByVal vFrm As FrmTrans_MaterialPlanningDashboard_Data1) As Boolean
        Dim conStr As String = Core_Module.GetConstring
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < vDGV.Rows.Count
                sql = Me.GetInsertString_Item(StartRowIndex, BatchQty, vDGV, vPrimaryTrans, vFrm)
                script.Query = sql
                'execute script
                script.Execute()
                StartRowIndex += BatchQty
            End While

            Conn.Close()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Private Function GetInsertString_Item(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, _
    ByVal vDGV As DataGridView, ByVal vPrimaryTrans As Double, ByVal vFrm As FrmTrans_MaterialPlanningDashboard_Data1) As String
        Dim retval As String = ""
        Dim Delimiter As String = ","
        Dim ObjCD As New Core_Module.ClsCekData
        Dim row As DataGridViewRow
        Dim OverrideInt As Integer = 0
        Dim ProcurementStrategyInt As Integer = 0

        Try
            retval = "INSERT INTO `tbltrans_material_planning_group_item_product` (`primary transaksi`,`primary code no`,`override`," + _
            "`procurement strategy`,`optimum stock strategy`,`optimum stock days`,`optimum stock qty`,`safety stock strategy`,`safety stock days`," + _
            "`safety stock qty`,`lead time strategy`,`lead time days`,`lead time qty`,`gross margin`,`gross margin perc`,`gross margin average`," + _
            "`nett sales`,`nett sales perc`,`nett sales average`,`primary unit`,`total qty`,`total qty perc`,`total qty avg`,`remarks`, `identifier`) VALUES "
            For i As Double = 0 To BatchQty - 1 Step 1
                If vDGV.Rows.Count > vStartRowIndex + i Then
                    row = vDGV.Rows(i + vStartRowIndex)
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vDGV.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    With row
                        If row.Cells(vFrm.ColCodeNo_Override.Index).Value = Boolean.TrueString Then
                            OverrideInt = 1
                        Else
                            OverrideInt = 0
                        End If

                        retval = retval + vbCrLf + "(" + CDbl(vPrimaryTrans).ToString + "," + CDbl(row.Cells(vFrm.ColCodeNo_PrimaryCodeNo.Index).Value).ToString + "," _
                        + OverrideInt.ToString + "," + CInt(Me.GetProcurementStrategy(row.Cells(vFrm.ColCodeNo_ProcurementStrategy.Index).Value)).ToString + "," _
                        + CInt(Me.GetLeadTimeStrategy(row.Cells(vFrm.ColCodeNo_OptimumStockStrategy.Index).Value)).ToString + "," + CInt(row.Cells(vFrm.ColCodeNo_OptimumStockDays.Index).Value).ToString + "," _
                        + CDec(row.Cells(vFrm.ColCodeNo_OptimumStockQty.Index).Value).ToString + "," + CInt(Me.GetLeadTimeStrategy(row.Cells(vFrm.ColCodeNo_SafetyStockStrategy.Index).Value)).ToString + "," _
                        + CInt(row.Cells(vFrm.ColCodeNo_SafetyStockStrategyDays.Index).Value).ToString + "," + CDec(row.Cells(vFrm.ColCodeNo_SafetyStockQty.Index).Value).ToString + "," _
                        + CInt(Me.GetLeadTimeStrategy(row.Cells(vFrm.ColCodeNo_LeadTimeStrategy.Index).Value)).ToString + "," + CInt(row.Cells(vFrm.ColCodeNo_LeadTimeDays.Index).Value).ToString + "," _
                        + CDec(row.Cells(vFrm.ColCodeNo_LeadTimeQty.Index).Value).ToString + "," + CDec(row.Cells(vFrm.ColCodeNo_GrossMargin.Index).Value).ToString + "," _
                        + CDec(row.Cells(vFrm.ColCodeNo_GrossMarginPerc.Index).Value).ToString + "," + CDec(row.Cells(vFrm.ColCodeNo_AverageGrossMargin.Index).Value).ToString + "," _
                        + CDec(row.Cells(vFrm.ColCodeNo_NettSales.Index).Value).ToString + "," + CDec(row.Cells(vFrm.ColCodeNo_NettSalesPerc.Index).Value).ToString + "," _
                        + CDec(row.Cells(vFrm.ColCodeNo_AverageNettSales.Index).Value).ToString + "," + CDbl(row.Cells(vFrm.ColCodeNo_PrimaryUnit.Index).Value).ToString + "," _
                        + CDec(row.Cells(vFrm.ColCodeNo_Qty.Index).Value).ToString + "," + CDec(row.Cells(vFrm.ColCodeNo_QtyPerc.Index).Value).ToString + "," _
                        + CDec(row.Cells(vFrm.ColCodeNo_AverageQty.Index).Value).ToString + "," + "'" + row.Cells(vFrm.ColCodeNo_Remarks.Index).Value.ToString + "'" _
                        + "," + "'" + row.Cells(vFrm.ColCodeNo_IdentifierGroup.Index).Value.ToString + "'" + "" _
                        + ")" + Delimiter
                    End With
                Else
                    Exit For
                End If
            Next
            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            ObjCD = Nothing
        End Try
    End Function

    Private Function Simpan_ItemSR_UsingScript(ByVal vTBlItemSR As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_srDataTable, _
    ByVal vPrimaryTrans As Double) As Boolean
        Dim conStr As String = Core_Module.GetConstring
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < vTBlItemSR.Rows.Count
                sql = Me.GetInsertString_ItemSR(StartRowIndex, BatchQty, vTBlItemSR, vPrimaryTrans)
                script.Query = sql
                'execute script
                script.Execute()
                StartRowIndex += BatchQty
            End While

            Conn.Close()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Private Function GetInsertString_ItemSR(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, _
    ByVal vTBlItemSR As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_srDataTable, _
    ByVal vPrimaryTrans As Double) As String
        Dim retval As String = ""
        Dim Delimiter As String = ","
        Dim ObjCD As New Core_Module.ClsCekData
        Dim row As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_srRow

        Try
            retval = "INSERT INTO `tbltrans_material_planning_group_item_sr` (`primary transaksi`,`primary sr`,`sr no historical`," + _
            "`sr date historical`,`primary customer`,`customer code historical`,`customer name historical`,`primary code no`,`primary unit`," + _
            "`qty`,`sales amount`,`cost`) VALUES "

            For i As Double = 0 To BatchQty - 1 Step 1
                If vTBlItemSR.Rows.Count > vStartRowIndex + i Then
                    row = vTBlItemSR.Rows(i + vStartRowIndex)
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTBlItemSR.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    With row
                        retval = retval + vbCrLf + "(" + CDbl(vPrimaryTrans).ToString + "," + row.primary_sr.ToString + "," _
                        + "'" + ObjCD.FormatSpecialChar_MySQL(row.sr_no_historical) + "'" + "," + "'" + Format(row.sr_date_historical, "yyyy-MM-dd") + "'" + "," _
                        + row.primary_customer.ToString + "," + "'" + ObjCD.FormatSpecialChar_MySQL(row.customer_code_historical) + "'" + "," _
                        + "'" + ObjCD.FormatSpecialChar_MySQL(row.customer_name_historical) + "'" + "," + row.primary_code_no.ToString + "," _
                        + row.primary_unit.ToString + "," + row.qty.ToString + "," _
                        + row.sales_amount.ToString + "," + row.cost.ToString + "" _
                        + ")" + Delimiter
                    End With
                Else
                    Exit For
                End If
            Next
            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            ObjCD = Nothing
        End Try
    End Function

    Private Function Simpan_ItemBubdling_UsingScript(ByVal vTBlItemBund As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_bundlingDataTable, _
    ByVal vPrimaryTrans As Double) As Boolean
        Dim conStr As String = Core_Module.GetConstring
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < vTBlItemBund.Rows.Count
                sql = Me.GetInsertString_ItemBund(StartRowIndex, BatchQty, vTBlItemBund, vPrimaryTrans)
                script.Query = sql
                'execute script
                script.Execute()
                StartRowIndex += BatchQty
            End While

            Conn.Close()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Private Function GetInsertString_ItemBund(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, _
    ByVal vTBlItemBund As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_bundlingDataTable, _
    ByVal vPrimaryTrans As Double) As String
        Dim retval As String = ""
        Dim Delimiter As String = ","
        Dim ObjCD As New Core_Module.ClsCekData
        Dim row As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_bundlingRow

        Try
            retval = "INSERT INTO `tbltrans_material_planning_group_item_bundling` (`primary transaksi`,`primary bundling`," + _
            "`bundling no historical`,`bundling date historical`,`primary code no`,`primary unit`,`qty`) VALUES "

            For i As Double = 0 To BatchQty - 1 Step 1
                If vTBlItemBund.Rows.Count > vStartRowIndex + i Then
                    row = vTBlItemBund.Rows(i + vStartRowIndex)
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTBlItemBund.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    With row
                        retval = retval + vbCrLf + "(" + CDbl(vPrimaryTrans).ToString + "," + row.primary_bundling.ToString + "," _
                        + "'" + ObjCD.FormatSpecialChar_MySQL(row.bundling_no_historical) + "'" + "," + "'" + Format(row.bundling_date_historical, "yyyy-MM-dd") + "'" + "," _
                        + row.primary_code_no.ToString + "," _
                        + row.primary_unit.ToString + "," + row.qty.ToString + "" _
                        + ")" + Delimiter
                    End With
                Else
                    Exit For
                End If
            Next
            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            ObjCD = Nothing
        End Try
    End Function

    Private Function Simpan_ItemAdjustment_UsingScript(ByVal vTBlItemAdj As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_adjustmentDataTable, _
    ByVal vPrimaryTrans As Double) As Boolean
        Dim conStr As String = Core_Module.GetConstring
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < vTBlItemAdj.Rows.Count
                sql = Me.GetInsertString_ItemAdj(StartRowIndex, BatchQty, vTBlItemAdj, vPrimaryTrans)
                script.Query = sql
                'execute script
                script.Execute()
                StartRowIndex += BatchQty
            End While

            Conn.Close()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Private Function GetInsertString_ItemAdj(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, _
    ByVal vTBlItemAdj As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_adjustmentDataTable, _
    ByVal vPrimaryTrans As Double) As String
        Dim retval As String = ""
        Dim Delimiter As String = ","
        Dim ObjCD As New Core_Module.ClsCekData
        Dim row As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_adjustmentRow

        Try
            retval = "INSERT INTO `tbltrans_material_planning_group_item_adjustment` (`primary transaksi`," + _
            "`primary adj`,`adj no historical`,`adj date historical`,`primary code no`,`primary unit`,`qty`, `primary tipe adjustment`) VALUES "

            For i As Double = 0 To BatchQty - 1 Step 1
                If vTBlItemAdj.Rows.Count > vStartRowIndex + i Then
                    row = vTBlItemAdj.Rows(i + vStartRowIndex)
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTBlItemAdj.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    With row
                        retval = retval + vbCrLf + "(" + CDbl(vPrimaryTrans).ToString + "," + row.primary_adj.ToString + "," _
                        + "'" + ObjCD.FormatSpecialChar_MySQL(row.adj_no_historical) + "'" + "," + "'" + Format(row.adj_date_historical, "yyyy-MM-dd") + "'" + "," _
                        + row.primary_code_no.ToString + "," _
                        + row.primary_unit.ToString + "," + row.qty.ToString + "," _
                        + row.primary_tipe_adjustment.ToString + "" _
                        + ")" + Delimiter
                    End With
                Else
                    Exit For
                End If
            Next
            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            ObjCD = Nothing
        End Try
    End Function

    Private Function Simpan_ItemInvoice_UsingScript(ByVal vTBlItemInvoice As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_invoiceDataTable, _
    ByVal vPrimaryTrans As Double) As Boolean
        Dim conStr As String = Core_Module.GetConstring
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < vTBlItemInvoice.Rows.Count
                sql = Me.GetInsertString_ItemInvoice(StartRowIndex, BatchQty, vTBlItemInvoice, vPrimaryTrans)
                script.Query = sql
                'execute script
                script.Execute()
                StartRowIndex += BatchQty
            End While

            Conn.Close()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Private Function GetInsertString_ItemInvoice(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, _
    ByVal vTBlItemInvoice As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_invoiceDataTable, _
    ByVal vPrimaryTrans As Double) As String
        Dim retval As String = ""
        Dim Delimiter As String = ","
        Dim ObjCD As New Core_Module.ClsCekData
        Dim row As DseTransaksiMaterialPlanningGroup.tbltrans_material_planning_group_item_invoiceRow

        Try
            retval = "INSERT INTO `tbltrans_material_planning_group_item_invoice` (`primary transaksi`,`primary invoice`," + _
            "`invoice no historical`,`invoice date historical`,`primary customer`,`customer code historical`,`customer name historical`," + _
            "`primary code no`,`primary unit`,`qty`,`sales amount`,`cost`) VALUES "
            For i As Double = 0 To BatchQty - 1 Step 1
                If vTBlItemInvoice.Rows.Count > vStartRowIndex + i Then
                    row = vTBlItemInvoice.Rows(i + vStartRowIndex)
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTBlItemInvoice.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    With row
                        retval = retval + vbCrLf + "(" + CDbl(vPrimaryTrans).ToString + "," + row.primary_invoice.ToString + "," _
                        + "'" + ObjCD.FormatSpecialChar_MySQL(row.invoice_no_historical) + "'" + "," + "'" + Format(row.invoice_date_historical, "yyyy-MM-dd") + "'" + "," _
                        + row.primary_customer.ToString + "," + "'" + ObjCD.FormatSpecialChar_MySQL(row.customer_code_historical) + "'" + "," _
                        + "'" + ObjCD.FormatSpecialChar_MySQL(row.customer_name_historical) + "'" + "," + row.primary_code_no.ToString + "," _
                        + row.primary_unit.ToString + "," + row.qty.ToString + "," _
                        + row.sales_amount.ToString + "," + row.cost.ToString + "" _
                        + ")" + Delimiter
                    End With
                Else
                    Exit For
                End If
            Next
            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            ObjCD = Nothing
        End Try
    End Function

    Public Sub Hapus(ByVal vPrimary As Double)
        Dim DTAGroup As New DseTransaksiMaterialPlanningGroupTableAdapters.tbltrans_material_planning_group_item_groupTableAdapter
        Dim DTASupplier As New DseTransaksiMaterialPlanningGroupTableAdapters.tbltrans_material_planning_group_itemsupplierTableAdapter
        Dim DTAProductGroup As New DseTransaksiMaterialPlanningGroupTableAdapters.tbltrans_material_planning_group_item_product_groupTableAdapter
        Dim DTAProduct As New DseTransaksiMaterialPlanningGroupTableAdapters.tbltrans_material_planning_group_item_productTableAdapter
        Dim DTAItemInvoice As New DseTransaksiMaterialPlanningGroupTableAdapters.tbltrans_material_planning_group_item_invoiceTableAdapter
        Dim DTAItemSR As New DseTransaksiMaterialPlanningGroupTableAdapters.tbltrans_material_planning_group_item_srTableAdapter
        Dim DTAItemAdjustment As New DseTransaksiMaterialPlanningGroupTableAdapters.tbltrans_material_planning_group_item_adjustmentTableAdapter
        Dim DTAItemBundling As New DseTransaksiMaterialPlanningGroupTableAdapters.tbltrans_material_planning_group_item_bundlingTableAdapter

        Try
            DTAGroup.Connection.ConnectionString = Core_Module.GetConstring
            DTAProduct.Connection.ConnectionString = DTAGroup.Connection.ConnectionString
            DTAProductGroup.Connection.ConnectionString = DTAGroup.Connection.ConnectionString
            DTASupplier.Connection.ConnectionString = DTAGroup.Connection.ConnectionString
            DTAItemInvoice.Connection.ConnectionString = DTAGroup.Connection.ConnectionString
            DTAItemSR.Connection.ConnectionString = DTAGroup.Connection.ConnectionString
            DTAItemAdjustment.Connection.ConnectionString = DTAGroup.Connection.ConnectionString
            DTAItemBundling.Connection.ConnectionString = DTAGroup.Connection.ConnectionString

            DTAGroup.Delete(vPrimary)
            DTAProduct.Delete(vPrimary)
            DTAProductGroup.Delete(vPrimary)
            DTASupplier.Delete(vPrimary)
            DTAItemInvoice.Delete(vPrimary)
            DTAItemSR.Delete(vPrimary)
            DTAItemAdjustment.Delete(vPrimary)
            DTAItemBundling.Delete(vPrimary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAGroup.Dispose()
            DTASupplier.Dispose()
            DTAProductGroup.Dispose()
            DTAProduct.Dispose()
            DTAItemInvoice.Dispose()
            DTAItemSR.Dispose()
            DTAItemAdjustment.Dispose()
            DTAItemBundling.Dispose()
        End Try
    End Sub
End Class
