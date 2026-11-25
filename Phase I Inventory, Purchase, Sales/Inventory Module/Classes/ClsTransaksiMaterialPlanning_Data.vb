'''==by @rosadinnn=='''

Option Explicit On
Option Strict Off

Public Class ClsTransaksiMaterialPlanning_Data

    Dim objCd As New Core_Module.ClsCekData

    Public Enum EnumPencarian As Integer
        CodeDesc
        CodeDescSupp
        CodeDescCat
        CodeDescSubbrand
        CodeDescBrand
        CodeDescProductgroup
        CodeDescSuppCat
        CodeDescSuppSubbrand
        CodeDescSuppBrand
        CodeDescSuppProductGroup
    End Enum

    Public Function filterData(ByVal vTipePencarian As ClsTransaksiMaterialPlanning_Data.EnumPencarian, ByVal vFrm As FrmTrans_MaterialPlanningDashboard_Data) As Boolean
        Dim dta As New DseTrans_MaterialPlanningDashboard_DataTableAdapters.tblmaster_productTableAdapter
        Dim tbl As New DseTrans_MaterialPlanningDashboard_Data.tblmaster_productDataTable
        Dim row As DseTrans_MaterialPlanningDashboard_Data.tblmaster_productRow
        Dim rowDgv As DataGridViewRow

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            dta.Connection.ConnectionString = Core_Module.GetConstring
            Select Case vTipePencarian
                Case EnumPencarian.CodeDesc
                    dta.FillBy_1CodeDesc(tbl, vFrm.txtProductCode.Text & "%", vFrm.TxtProductDesc.Text & "%")
                    Exit Select
                Case EnumPencarian.CodeDescSupp
                    dta.FillBy_2CodeDescSupp(tbl, vFrm.txtProductCode.Text & "%", vFrm.TxtProductDesc.Text & "%", vFrm.txtPrimarySupp.Text)
                    Exit Select
                Case EnumPencarian.CodeDescCat
                    dta.FillBy_3CodeDescCat(tbl, vFrm.txtProductCode.Text & "%", vFrm.TxtProductDesc.Text & "%", vFrm.TxtPrimaryGrop.Text)
                    Exit Select
                Case EnumPencarian.CodeDescSubbrand
                    dta.FillBy_4CodeDescSubbrand(tbl, vFrm.txtProductCode.Text & "%", vFrm.TxtProductDesc.Text & "%", vFrm.TxtPrimaryGrop.Text)
                    Exit Select
                Case EnumPencarian.CodeDescBrand
                    dta.FillBy_5CodeDescBrand(tbl, vFrm.txtProductCode.Text & "%", vFrm.TxtProductDesc.Text & "%", vFrm.TxtPrimaryGrop.Text)
                    Exit Select
                Case EnumPencarian.CodeDescProductgroup
                    dta.FillBy_6CodeDescGroup(tbl, vFrm.txtProductCode.Text & "%", vFrm.TxtProductDesc.Text & "%", vFrm.TxtPrimaryGrop.Text)
                    Exit Select
                Case EnumPencarian.CodeDescSuppCat
                    dta.FillBy_7CodeDescSuppCat(tbl, vFrm.txtProductCode.Text & "%", vFrm.TxtProductDesc.Text & "%", vFrm.txtPrimarySupp.Text, vFrm.TxtPrimaryGrop.Text)
                    Exit Select
                Case EnumPencarian.CodeDescSuppSubbrand
                    dta.FillBy_8CodeDescSuppSubbrand(tbl, vFrm.txtProductCode.Text & "%", vFrm.TxtProductDesc.Text & "%", vFrm.txtPrimarySupp.Text, vFrm.TxtPrimaryGrop.Text)
                    Exit Select
                Case EnumPencarian.CodeDescSuppBrand
                    dta.FillBy_9CodeDescSuppBrand(tbl, vFrm.txtProductCode.Text & "%", vFrm.TxtProductDesc.Text & "%", vFrm.txtPrimarySupp.Text, vFrm.TxtPrimaryGrop.Text)
                    Exit Select
                Case EnumPencarian.CodeDescSuppProductGroup
                    dta.FillBy_10CodeDescSuppGroup(tbl, vFrm.txtProductCode.Text & "%", vFrm.TxtProductDesc.Text & "%", vFrm.txtPrimarySupp.Text, vFrm.TxtPrimaryGrop.Text)
                    Exit Select
            End Select

            'Taufik//Tambahan filter untuk 
            Dim filterStr As String = ""
            If vFrm.CboFilterProcurementStrategy.Text = "Build to stock" Then
                filterStr = "[Procurement Strategy] = 1"
            ElseIf vFrm.CboFilterProcurementStrategy.Text = "Build to order" Then
                filterStr = "[Procurement Strategy] = 0"
            End If

            If tbl.DefaultView.Count > 0 Then
                vFrm.dgvItem.Rows.Clear()
                'vFrm.dgvItem.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect
                vFrm.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                'For Each row In tbl
                For ix As Integer = 0 To tbl.DefaultView.Count - 1
                    row = tbl.DefaultView(ix).Row

                    vFrm.dgvItem.Rows.Add()
                    rowDgv = vFrm.dgvItem.Rows(vFrm.dgvItem.Rows.Count - 1)
                    With row
                        rowDgv.Cells(vFrm.ColNo.Index).Value = rowDgv.Index + 1
                        rowDgv.Cells(vFrm.ColCodeNo.Index).Value = .Code_No
                        rowDgv.Cells(vFrm.ColPrimaryCodeNo.Index).Value = .Primary_Product
                        rowDgv.Cells(vFrm.ColDescription.Index).Value = .Description
                        rowDgv.Cells(vFrm.ColUnit.Index).Value = .Unit
                        rowDgv.Cells(vFrm.ColPrimaryUnit.Index).Value = .Primary_Unit
                        If .Procurement_Type = Core_Module.clsMasterProduct.EnumProcurementType.Buy Then
                            rowDgv.Cells(vFrm.ColProcurementType.Index).Value = "Buy"
                        ElseIf .Procurement_Type = Core_Module.clsMasterProduct.EnumProcurementType.Produce Then
                            rowDgv.Cells(vFrm.ColProcurementType.Index).Value = "Produce"
                        End If
                        If .Procurement_Strategy = Core_Module.clsMasterProduct.EnumProcurementStrategy.Build_To_Order Then
                            rowDgv.Cells(vFrm.ColProcurementStrategy.Index).Value = "Build To Order"
                        ElseIf .Procurement_Strategy = Core_Module.clsMasterProduct.EnumProcurementStrategy.Build_To_Stock Then
                            rowDgv.Cells(vFrm.ColProcurementStrategy.Index).Value = "Build To Stock"
                        ElseIf .Procurement_Strategy = Core_Module.clsMasterProduct.EnumProcurementStrategy.Both Then
                            rowDgv.Cells(vFrm.ColProcurementStrategy.Index).Value = "Both"
                        End If
                        If .Supplier_Selection_Mode = Core_Module.clsMasterProduct.EnumSupplierSelectionMode.The_cheapest_in_the_last_3_order Then
                            rowDgv.Cells(vFrm.ColSupplierSelectionMode.Index).Value = "The cheapest in the last 3 order"
                        ElseIf .Supplier_Selection_Mode = Core_Module.clsMasterProduct.EnumSupplierSelectionMode.Prefered_Supplier Then
                            rowDgv.Cells(vFrm.ColSupplierSelectionMode.Index).Value = "Prefered supplier"
                        ElseIf .Supplier_Selection_Mode = Core_Module.clsMasterProduct.EnumSupplierSelectionMode.Most_Recent_PO Then
                            rowDgv.Cells(vFrm.ColSupplierSelectionMode.Index).Value = "Most recent PO"
                        End If
                        rowDgv.Cells(vFrm.ColPreferedSupplier.Index).Value = .Supplier_Code
                        rowDgv.Cells(vFrm.ColPrimaryPreferedSupplier.Index).Value = .Primary_Prefered_Supplier
                        If .Recomended_Order_Qty_Type = Core_Module.clsMasterProduct.EnumOrderStrategy.Order_To_Min_Qty Then
                            rowDgv.Cells(vFrm.ColRecomendedOrderQtyType.Index).Value = "Order To Optimum Qty"
                        ElseIf .Recomended_Order_Qty_Type = Core_Module.clsMasterProduct.EnumOrderStrategy.Order_To_Optimum_Qty Then
                            rowDgv.Cells(vFrm.ColRecomendedOrderQtyType.Index).Value = "Order To Optimum Qty"
                        ElseIf .Recomended_Order_Qty_Type = Core_Module.clsMasterProduct.EnumOrderStrategy.Order_To_Max_Qty Then
                            rowDgv.Cells(vFrm.ColRecomendedOrderQtyType.Index).Value = "Order To Optimum Qty"
                        ElseIf .Recomended_Order_Qty_Type = Core_Module.clsMasterProduct.EnumOrderStrategy.Order_To_Safety_Stock Then
                            rowDgv.Cells(vFrm.ColRecomendedOrderQtyType.Index).Value = "Order To Optimum Qty"
                        ElseIf .Recomended_Order_Qty_Type = Core_Module.clsMasterProduct.EnumOrderStrategy.Order_By_Minimum_Order_Qty Then
                            rowDgv.Cells(vFrm.ColRecomendedOrderQtyType.Index).Value = "Order By Optimum Order Qty"
                        ElseIf .Recomended_Order_Qty_Type = Core_Module.clsMasterProduct.EnumOrderStrategy.Order_By_Optimum_Order_Qty Then
                            rowDgv.Cells(vFrm.ColRecomendedOrderQtyType.Index).Value = "Order By Optimum Order Qty"
                        ElseIf .Recomended_Order_Qty_Type = Core_Module.clsMasterProduct.EnumOrderStrategy.Order_By_Maximum_Order_Qty Then
                            rowDgv.Cells(vFrm.ColRecomendedOrderQtyType.Index).Value = "Order By Optimum Order Qty"
                        End If
                        rowDgv.Cells(vFrm.ColMinStockQty.Index).Value = Me.objCd.DigitGrouping(.Mimimum_Stock_Qty, ._Default__Precision_)
                        rowDgv.Cells(vFrm.ColOptStockQty.Index).Value = Me.objCd.DigitGrouping(.Optimum_Stock_Qty, ._Default__Precision_)
                        rowDgv.Cells(vFrm.ColMaxStockQty.Index).Value = Me.objCd.DigitGrouping(.MAximum_Stock_Qty, ._Default__Precision_)
                        rowDgv.Cells(vFrm.ColPrimaryPreferedOrderUnit.Index).Value = .Primary_Purchased_Order_Unit
                        rowDgv.Cells(vFrm.ColMinOrderQty.Index).Value = Me.objCd.DigitGrouping(.Minimum_Purchase_Qty, ._Default__Precision_)
                        rowDgv.Cells(vFrm.ColOptOrderQty.Index).Value = Me.objCd.DigitGrouping(.Optimum_Purchase_Qty, ._Default__Precision_)
                        rowDgv.Cells(vFrm.ColMaxOrderQty.Index).Value = Me.objCd.DigitGrouping(.Maximum_Purchase_Qty, ._Default__Precision_)
                        rowDgv.Cells(vFrm.ColAverageMonthlyDemand.Index).Value = Me.objCd.DigitGrouping(.Avg_Monthly_Demand, ._Default__Precision_)
                        rowDgv.Cells(vFrm.ColStartPeriod.Index).Value = Format(.Start_Period, "MMMM yyyy")
                        rowDgv.Cells(vFrm.ColEndPeriod.Index).Value = Format(.End_Period, "MMMM yyyy")
                        If .Include_Sales = 0 Then
                            rowDgv.Cells(vFrm.ColIncludeSales.Index).Value = False
                        Else
                            rowDgv.Cells(vFrm.ColIncludeSales.Index).Value = True
                        End If
                        If .Include_Bundling = 0 Then
                            rowDgv.Cells(vFrm.ColIncludeBundling.Index).Value = False
                        Else
                            rowDgv.Cells(vFrm.ColIncludeBundling.Index).Value = True
                        End If
                        If .Include_Adjustment = 0 Then
                            rowDgv.Cells(vFrm.ColIncludeAdjustment.Index).Value = False
                        Else
                            rowDgv.Cells(vFrm.ColIncludeAdjustment.Index).Value = True
                        End If
                        If .Safety_Stock_Strategy = Core_Module.clsMasterProduct.EnumSafetyStockStrategy.Based_On_Usage Then
                            rowDgv.Cells(vFrm.ColSafetyStockStrategy.Index).Value = "Based On Usage"
                        ElseIf .Safety_Stock_Strategy = Core_Module.clsMasterProduct.EnumSafetyStockStrategy.User_Defined Then
                            rowDgv.Cells(vFrm.ColSafetyStockStrategy.Index).Value = "User Defined"
                        End If
                        rowDgv.Cells(vFrm.ColSafetyStockDays.Index).Value = .Safety_Stock_Days
                        rowDgv.Cells(vFrm.ColSafetyStockQty.Index).Value = Me.objCd.DigitGrouping(.Safety_Stock_Qty, ._Default__Precision_)
                        If .Lead_Time_Strategy = Core_Module.clsMasterProduct.EnumLeadTimeStrategy.Based_On_Usage Then
                            rowDgv.Cells(vFrm.ColLeadTimeStrategy.Index).Value = "Based On Usage"
                        ElseIf .Lead_Time_Strategy = Core_Module.clsMasterProduct.EnumLeadTimeStrategy.User_Defined Then
                            rowDgv.Cells(vFrm.ColLeadTimeStrategy.Index).Value = "User Defined"
                        End If
                        rowDgv.Cells(vFrm.ColLeadTimeDays.Index).Value = .Lead_Time_Days
                        rowDgv.Cells(vFrm.ColLeadTimeQty.Index).Value = Me.objCd.DigitGrouping(.Lead_Time_Qty, ._Default__Precision_)

                        If .optimum_stock_strategy = Core_Module.clsMasterProduct.EnumOptimumStockStrategy.Based_On_Usage Then
                            rowDgv.Cells(vFrm.ColOptStockStrategy.Index).Value = "Based On Usage"
                        ElseIf .optimum_stock_strategy = Core_Module.clsMasterProduct.EnumOptimumStockStrategy.User_Defined Then
                            rowDgv.Cells(vFrm.ColOptStockStrategy.Index).Value = "User Defined"
                        End If
                        rowDgv.Cells(vFrm.ColOptStockDays.Index).Value = .optimum_stock_days
                        rowDgv.Cells(vFrm.ColOptStockQty.Index).Value = Me.objCd.DigitGrouping(.Optimum_Stock_Qty, ._Default__Precision_)

                        If .Exclude_on_Material_Planning = 0 Then
                            rowDgv.Cells(vFrm.ColExclude.Index).Value = False
                        Else
                            rowDgv.Cells(vFrm.ColExclude.Index).Value = True
                        End If
                        rowDgv.Cells(vFrm.ColExplainAverageQty.Index).Value = .__Explain__Avg_Qty
                        rowDgv.Cells(vFrm.ColExplainSafetyStockQty.Index).Value = .__Explain__Safety_Stock
                        rowDgv.Cells(vFrm.ColExplainLeadTimeQty.Index).Value = .__Explain__Lead_Time
                        rowDgv.Cells(vFrm.ColExplainOptStockQty.Index).Value = .explain_optimum_stock_qty
                        rowDgv.Cells(vFrm.ColDecimal.Index).Value = ._Default__Precision_
                        If .IsSupplier_NameNull = True Then
                            rowDgv.Cells(vFrm.ColSupplierName.Index).Value = ""
                        Else
                            rowDgv.Cells(vFrm.ColSupplierName.Index).Value = .Supplier_Name
                        End If
                        rowDgv.Cells(vFrm.ColDef_UpperTolerancePerct.Index).Value = Me.objCd.DigitGrouping(.default_upper_tolerance, 2)
                        rowDgv.Cells(vFrm.ColDef_LowerTolerancePerct.Index).Value = Me.objCd.DigitGrouping(.default_lower_tolerance, 2)
                    End With
                Next
                'Me.objCd.SetNumber(vFrm.dgvItem)
                vFrm.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                If vFrm.dgvItem.Rows.Count > 0 Then
                    vFrm.dgvItem.ClearSelection()
                    Dim dgvCell As DataGridViewCell
                    dgvCell = vFrm.dgvItem.Item(0, 0)
                    vFrm.dgvItem.CurrentCell = dgvCell
                End If
                'vFrm.dgvItem.SelectionMode = DataGridViewSelectionMode.CellSelect
            Else
                MsgBox("Search result: 0", MsgBoxStyle.Information, "Information")
                Exit Function
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Sub UpdateParams_ProcurementStrategy(ByVal vPrimary As Double, ByVal vParam As Core_Module.clsMasterProduct.EnumProcurementStrategy)
        Dim dta As New DseTrans_MaterialPlanningDashboard_DataTableAdapters.tblmaster_product1TableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dta.UpdateQuery_ProcurementStrategy(vParam, vPrimary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Sub

    Public Sub UpdateParams_ProcurementType(ByVal vPrimary As Double, ByVal vParam As Core_Module.clsMasterProduct.EnumProcurementType)
        Dim dta As New DseTrans_MaterialPlanningDashboard_DataTableAdapters.tblmaster_product1TableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dta.UpdateQuery_SetProcurementType(vParam, vPrimary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Sub

    Public Function UpdateParams_A(ByVal vPrimaryPreferedSupp As Double, ByVal vSelectSuppBy As Integer, ByVal vOrderMode As Integer, _
    ByVal vUserCode As String, ByVal vProcurementStrategy As Integer, ByVal vPrimary As Double, _
    ByVal dta As DseTrans_MaterialPlanningDashboard_DataTableAdapters.tblmaster_product1TableAdapter, ByVal vUpperTolerance As Decimal, _
    ByVal vLowerTolerance As Decimal, ByVal vProcurementType As Integer) As Boolean
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dta.UpdateQuery_A_Supp_SuppMode_OrderType(vPrimaryPreferedSupp, vSelectSuppBy, vOrderMode, vUserCode, Now, Now, _
            vProcurementStrategy, vUpperTolerance, vLowerTolerance, vProcurementType, vPrimary)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function UpdateParams_A_DefaultUpperTolerance(ByVal vUpperTolerance As Decimal, ByVal vUserCode As String, ByVal vPrimary As Double, _
    ByVal dta As DseTrans_MaterialPlanningDashboard_DataTableAdapters.tblmaster_product1TableAdapter) As Boolean
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dta.UpdateQuery_A_UpperTolerance(vUpperTolerance, vUserCode, Now, Now, vPrimary)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function UpdateParams_A_DefaultLowerTolerance(ByVal vLowerTolerance As Decimal, ByVal vUserCode As String, ByVal vPrimary As Double, _
    ByVal dta As DseTrans_MaterialPlanningDashboard_DataTableAdapters.tblmaster_product1TableAdapter) As Boolean
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dta.UpdateQuery_A_LowerTolerance(vLowerTolerance, vUserCode, Now, Now, vPrimary)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function UpdateParams_A_PreferedSupplier(ByVal vPrimaryPreferedSupp As Double, ByVal vUserCode As String, ByVal vPrimary As Double) As Boolean
        Dim dta As New DseTrans_MaterialPlanningDashboard_DataTableAdapters.tblmaster_product1TableAdapter
        Dim objS As New Core_Module.clsMasterSupplier
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dta.UpdateQuery_A_Supp(vPrimaryPreferedSupp, vUserCode, Now, Now, vPrimary)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function UpdateParams_A_SupplierSelectionMode(ByVal vSelectSuppBySupp As Core_Module.clsMasterProduct.EnumSupplierSelectionMode, ByVal vUserCode As String, ByVal vPrimary As Double) As Boolean
        Dim dta As New DseTrans_MaterialPlanningDashboard_DataTableAdapters.tblmaster_product1TableAdapter
        Dim objS As New Core_Module.clsMasterSupplier
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dta.UpdateQuery_A_SuppMode(vSelectSuppBySupp, vUserCode, Now, Now, vPrimary)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function UpdateParams_A_RecomendOrderType(ByVal vRecomendOrderType As Core_Module.clsMasterProduct.EnumOrderStrategy, ByVal vUserCode As String, ByVal vPrimary As Double) As Boolean
        Dim dta As New DseTrans_MaterialPlanningDashboard_DataTableAdapters.tblmaster_product1TableAdapter
        Dim objS As New Core_Module.clsMasterSupplier
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dta.UpdateQuery_A_OrderType(vRecomendOrderType, vUserCode, Now, Now, vPrimary)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function Update_B_All(ByVal vUserCode As String, ByVal vStartDate As Date, ByVal vFinishDate As Date, _
            ByVal vAvgMonthly As Double, ByVal vSafetyStockDays As Integer, ByVal vSafetyStockQty As Double, _
            ByVal vLeadTimeDays As Integer, ByVal vLeadTimeQty As Double, ByVal vSafetyStockStrategy As Integer, _
            ByVal vLeadTimeStrategy As Integer, ByVal vIncludeSales As Integer, ByVal vIncludeBundling As Integer, _
            ByVal vIncludeAdjustment As Integer, ByVal vExplainAvg As String, ByVal vExplainSafetyStock As String, ByVal vExplainLeadTime As String, _
            ByVal vOptStockStrategy As Integer, ByVal vOptStockDays As Integer, ByVal vExplainOptStock As String, ByVal vOptStockQty As Double, _
            ByVal vPrimary As Double, ByVal dtaB As DseTrans_MaterialPlanningDashboard_DataTableAdapters.tblmaster_product1TableAdapter) As Boolean

        Try
            dtaB.Connection.ConnectionString = Core_Module.GetConstring
            dtaB.UpdateQuery_B_All(vUserCode, Now, Now, vStartDate, vFinishDate, vAvgMonthly, _
                    vSafetyStockDays, vSafetyStockQty, vLeadTimeDays, vLeadTimeQty, vSafetyStockStrategy, _
                    vLeadTimeStrategy, vIncludeSales, vIncludeBundling, vIncludeAdjustment, vExplainAvg, vExplainSafetyStock, vExplainLeadTime, _
                    vOptStockStrategy, vOptStockDays, vExplainOptStock, vOptStockQty, vPrimary)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function Update_C_Exlude(ByVal vExclude As Integer, ByVal vUserCode As String, ByVal vPrimary As Double) As Boolean
        Dim dta As New DseTrans_MaterialPlanningDashboard_DataTableAdapters.tblmaster_product1TableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dta.UpdateQuery_C_Exclude(vExclude, vUserCode, Now, Now, vPrimary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function Update_D_MinOptMaxOrderStock(ByVal vMinOrder As Double, ByVal vOptOrder As Double, ByVal vMaxOrder As Double, _
            ByVal vMinStock As Double, ByVal vOptStock As Double, ByVal vMaxStock As Double, _
            ByVal vUserCode As String, ByVal vPrimary As Double, ByVal dta As DseTrans_MaterialPlanningDashboard_DataTableAdapters.tblmaster_product1TableAdapter) As Boolean

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dta.UpdateQuery_D_MinOptMaxOrderStock(vMinStock, vOptStock, vMaxStock, vMinOrder, vOptOrder, vMaxOrder, vUserCode, Now, Now, vPrimary)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetHistoricalByPrimary(ByRef vTbl As DseTrans_MaterialPlanningDashboard_Data.masterHistoricalDataTable, ByVal vPrimary As Double, ByVal vStart As Date, ByVal vFinish As Date) As Boolean
        Dim dta As New DseTrans_MaterialPlanningDashboard_DataTableAdapters.getHistoricalTableAdapter
        Dim tbl As New DseTrans_MaterialPlanningDashboard_Data.getHistoricalDataTable
        Dim row As DseTrans_MaterialPlanningDashboard_Data.getHistoricalRow
        Dim rowMaster As DseTrans_MaterialPlanningDashboard_Data.masterHistoricalRow
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            Dim totalBulanPerPeriode As Integer = 0
            Dim rataRataPerPeriode As Double = 0
            Dim increament As Integer = 0
            If vPrimary = -1 Then
                dta.Fill(tbl)
            Else
                dta.FillBy_PrimaryProduct(tbl, vPrimary)
            End If
            If tbl.Rows.Count > 0 Then
                For Each row In tbl
                    totalBulanPerPeriode = DateDiff(DateInterval.Month, row.Start_Period, row.End_Periode) + 1
                    rataRataPerPeriode = row.Qty / totalBulanPerPeriode
                    increament = 0
                    While row.Start_Period.AddMonths(increament) <= row.End_Periode
                        ''cek ke master hostorical
                        vTbl.DefaultView.RowFilter = "[Periode] = #" & Format(Me.objCd.GetFirstDateofTheMonth(row.Start_Period.AddMonths(increament)), "MM/dd/yyyy").ToString & "#"
                        If vTbl.DefaultView.Count = 0 Then
                            rowMaster = vTbl.NewmasterHistoricalRow
                            rowMaster.Primary_Product = vPrimary
                            rowMaster.jml = rataRataPerPeriode
                            rowMaster.Periode = Me.objCd.GetFirstDateofTheMonth(row.Start_Period.AddMonths(increament))
                        End If
                        increament += 1
                    End While
                Next
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function CalculateAverageMonthlyDemandByPrimary(ByVal vPrimaryProduct As Double, _
    ByVal vStartPeriode As Date, ByVal vFinishPeriode As Date, ByVal vDaysPerMonth As Integer, _
    ByVal vIsIncludeSales As Boolean, ByVal vIncludeBundling As Boolean, ByVal vIncludeAdjustment As Boolean, _
    ByRef vAvgMonthly As Double, ByRef vExplainAvgMonthly As String, _
    ByVal vTblHistoricalMaster As DseTrans_MaterialPlanningDashboard_Data.masterHistoricalDataTable, _
    ByVal vDec As Integer) As Boolean
        Dim valSales As Double = 0
        Dim valBundling As Double = 0
        Dim valAdjustment As Double = 0
        Dim valHistorical As Double = 0
        Dim valTotalPerMonth As Double = 0
        Dim valTotalUsage As Double = 0
        Dim dtaSales As New DseTrans_MaterialPlanningDashboard_DataTableAdapters.getSalesAndAdjustment_RTableAdapter
        Dim dtaBundling As New DseTrans_MaterialPlanningDashboard_DataTableAdapters.getBundlingTableAdapter
        Dim dtaAdjustment As New DseTrans_MaterialPlanningDashboard_DataTableAdapters.getSalesAndAdjustment_RTableAdapter
        'Dim dtaHistorical As New DseTrans_MaterialPlanningDashboard_DataTableAdapters.getHistoricalTableAdapter
        Dim tblSales As New DseTrans_MaterialPlanningDashboard_Data.getSalesAndAdjustment_RDataTable
        Dim tblBundling As New DseTrans_MaterialPlanningDashboard_Data.getBundlingDataTable
        Dim tblAdjustment As New DseTrans_MaterialPlanningDashboard_Data.getSalesAndAdjustment_RDataTable
        'Dim tblHistorical As New DseTrans_MaterialPlanningDashboard_Data.getHistoricalDataTable
        'Dim tblHistoricalMaster As DseTrans_MaterialPlanningDashboard_Data.masterHistoricalDataTable
        Dim rowSales As DseTrans_MaterialPlanningDashboard_Data.getSalesAndAdjustment_RRow
        Dim rowBundling As DseTrans_MaterialPlanningDashboard_Data.getBundlingRow
        Dim rowAdj As DseTrans_MaterialPlanningDashboard_Data.getSalesAndAdjustment_RRow
        'Dim rowHistorical As DseTrans_MaterialPlanningDashboard_Data.getHistoricalRow
        Dim rowHistoricalMaster As DseTrans_MaterialPlanningDashboard_Data.masterHistoricalRow
        Dim TBLSR As New DseTrans_MaterialPlanningDashboard_Data.getSalesAndAdjustment_RDataTable
        Dim rowSR As DseTrans_MaterialPlanningDashboard_Data.getSalesAndAdjustment_RRow

        Try
            dtaAdjustment.Connection.ConnectionString = Core_Module.GetConstring
            dtaBundling.Connection.ConnectionString = dtaAdjustment.Connection.ConnectionString
            dtaSales.Connection.ConnectionString = dtaAdjustment.Connection.ConnectionString

            vStartPeriode = Me.objCd.GetFirstDateofTheMonth(vStartPeriode)
            vFinishPeriode = Me.objCd.GetLasDateOfTheMonth(vFinishPeriode)
            vExplainAvgMonthly = Format(vStartPeriode, "MMM yyyy").ToString & vbNewLine
            vExplainAvgMonthly += Format(vFinishPeriode, "MMM yyyy").ToString & vbNewLine
            vExplainAvgMonthly += vDaysPerMonth.ToString & vbNewLine
            If vIsIncludeSales = True Then ''for sales
                vExplainAvgMonthly += "1" & vbNewLine
            Else
                vExplainAvgMonthly += "0" & vbNewLine
            End If
            If vIncludeBundling = True Then ''for bundling
                vExplainAvgMonthly += "1" & vbNewLine
            Else
                vExplainAvgMonthly += "0" & vbNewLine
            End If
            If vIncludeAdjustment = True Then ''for adjustment
                vExplainAvgMonthly += "1" & vbNewLine
            Else
                vExplainAvgMonthly += "0" & vbNewLine
            End If

            ''hitung per bulan::
            For i As Integer = 0 To DateDiff(DateInterval.Month, vStartPeriode, vFinishPeriode)
                ''get sales
                If vIsIncludeSales = True Then
                    dtaSales.FillBy_PrimaryProductEtc(tblSales, clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Sales, Me.objCd.GetLasDateOfTheMonth(vStartPeriode.AddMonths(i)), vStartPeriode.AddMonths(i), vPrimaryProduct)
                    If tblSales.Rows.Count > 0 Then
                        rowSales = tblSales.Rows(0)
                        valSales = Me.objCd.DigitGrouping(rowSales.jml, vDec)
                        'kurangkan dengan sales return 
                        dtaSales.FillBy_SR_PrimaryProduct(TBLSR, clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn, _
                        Me.objCd.GetLasDateOfTheMonth(vStartPeriode.AddMonths(i)), vStartPeriode.AddMonths(i), vPrimaryProduct)
                        If TBLSR.Rows.Count > 0 Then
                            rowSR = TBLSR.Rows(0)
                            valSales -= rowSR.jml
                            If valSales < 0 Then
                                valSales = 0
                            End If
                        End If
                    Else
                        valSales = 0
                    End If
                Else
                    valSales = 0
                End If
                ''get bundling
                If vIncludeBundling = True Then
                    dtaBundling.FillBy_PrimaryProductEtc(tblBundling, Me.objCd.GetLasDateOfTheMonth(vStartPeriode.AddMonths(i)), vStartPeriode.AddMonths(i), vPrimaryProduct)
                    If tblBundling.Rows.Count > 0 Then
                        rowBundling = tblBundling.Rows(0)
                        valBundling = Me.objCd.DigitGrouping(rowBundling.jml, vDec)
                    Else
                        valBundling = 0
                    End If
                Else
                    valBundling = 0
                End If
                ''get adjustment
                If vIncludeAdjustment = True Then
                    dtaAdjustment.FillBy_PrimaryProductEtc(tblAdjustment, clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust, Me.objCd.GetLasDateOfTheMonth(vStartPeriode.AddMonths(i)), vStartPeriode.AddMonths(i), vPrimaryProduct)
                    If tblAdjustment.Rows.Count > 0 Then
                        rowAdj = tblAdjustment.Rows(0)
                        valAdjustment = Me.objCd.DigitGrouping(rowAdj.jml, vDec)
                    Else
                        valAdjustment = 0
                    End If
                Else
                    valAdjustment = 0
                End If
                ''get historical
                If vTblHistoricalMaster.Rows.Count > 0 Then
                    vTblHistoricalMaster.DefaultView.RowFilter = "[Primary Product] = " & vPrimaryProduct.ToString & " AND [Periode] = #" & Format(vStartPeriode.AddMonths(i), "MM/dd/yyyy").ToString & "#"
                    If vTblHistoricalMaster.DefaultView.Count > 0 Then
                        rowHistoricalMaster = vTblHistoricalMaster.DefaultView(0).Row
                        valHistorical = Me.objCd.DigitGrouping(rowHistoricalMaster.jml, vDec)
                    Else
                        valHistorical = 0
                    End If
                Else
                    valHistorical = 0
                End If

                ''hitung per bulan
                valTotalPerMonth = valSales + valBundling + valAdjustment + valHistorical
                valTotalUsage += valTotalPerMonth

                ''tambahkan data bulanan ke explain
                vExplainAvgMonthly += Format(vStartPeriode.AddMonths(i), "MMM yyyy").ToString & ";" & valSales.ToString & ";" & valBundling.ToString & ";" & valAdjustment.ToString & ";" & valHistorical.ToString & ";" & valTotalPerMonth.ToString & vbNewLine

            Next
            ''hitung rata-rata
            vAvgMonthly = Me.objCd.DigitGrouping(valTotalUsage / (DateDiff(DateInterval.Month, vStartPeriode, vFinishPeriode) + 1), vDec)

            ''simpan rata-rata ke explain
            vExplainAvgMonthly += valTotalUsage.ToString & vbNewLine
            vExplainAvgMonthly += valTotalUsage.ToString & " / " & (DateDiff(DateInterval.Month, vStartPeriode, vFinishPeriode) + 1).ToString & " = " & vAvgMonthly.ToString

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dtaSales.Dispose()
            dtaBundling.Dispose()
            dtaAdjustment.Dispose()
            tblSales.Dispose()
            tblBundling.Dispose()
            tblAdjustment.Dispose()
            TBLSR.Dispose()
        End Try
    End Function

    Public Function CalculateSafetyStock(ByVal vAvgMonthly As Double, ByVal vSafetyTime As Integer, ByVal vDaysPerMonth As Integer, _
            ByRef vSafetyStock As Double, ByRef vExplainSafetyStock As String, ByVal vDec As Integer) As Boolean
        Try
            If vDaysPerMonth <= 0 Then
                vSafetyStock = 0
            Else
                vSafetyStock = Me.objCd.DigitGrouping(vAvgMonthly * vSafetyTime / vDaysPerMonth, vDec)
            End If
            vExplainSafetyStock = vSafetyTime.ToString & vbNewLine
            vExplainSafetyStock += vAvgMonthly.ToString & ";" & vSafetyTime.ToString & ";" & vDaysPerMonth.ToString & ";" & vSafetyStock.ToString
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Function

    Public Function CalculateSafetyStockTime(ByVal vAvgMonthly As Double, ByVal vSafetyStock As Double, ByVal vDaysPerMonth As Integer, _
            ByRef vSafetyStockTime As Integer, ByRef vExplainSafetyStock As String, ByVal vDec As Integer) As Boolean
        Try
            If vAvgMonthly <= 0 Then
                vSafetyStockTime = 0
            Else
                vSafetyStockTime = vSafetyStock * vDaysPerMonth / vAvgMonthly
            End If
            vExplainSafetyStock = vSafetyStock.ToString & vbNewLine
            vExplainSafetyStock += vAvgMonthly.ToString & ";" & vSafetyStockTime.ToString & ";" & vDaysPerMonth.ToString & ";" & vSafetyStock.ToString
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Function

    Public Function CalculateLeadTimeQty(ByVal vAvgMonthly As Double, ByVal vLeadTimeDays As Integer, ByVal vDaysPerMonth As Integer, _
            ByRef vLeadTimeStock As Double, ByRef vExplainLeadTime As String, ByVal vDec As Integer) As Boolean
        Try
            If vDaysPerMonth <= 0 Then
                vLeadTimeStock = 0
            Else
                vLeadTimeStock = Me.objCd.DigitGrouping(vAvgMonthly * vLeadTimeDays / vDaysPerMonth, vDec)
            End If
            vExplainLeadTime = vLeadTimeDays.ToString & vbNewLine
            vExplainLeadTime += vAvgMonthly.ToString & ";" & vLeadTimeDays.ToString & ";" & vDaysPerMonth.ToString & ";" & vLeadTimeStock.ToString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Function

    Public Function CalculatedLeadTimeDays(ByVal vAvgMonthly As Double, ByVal vLeadTimeStock As Double, ByVal vDaysPerMonth As Integer, _
            ByRef vLeadTimeDays As Integer, ByRef vExplainLeadTime As String, ByVal vDec As Integer) As Boolean
        Try
            If vAvgMonthly <= 0 Then
                vLeadTimeDays = 0
            Else
                vLeadTimeDays = vLeadTimeStock * vDaysPerMonth / vAvgMonthly
            End If
            vExplainLeadTime = vLeadTimeDays.ToString & vbNewLine
            vExplainLeadTime += vAvgMonthly.ToString & ";" & vLeadTimeDays.ToString & ";" & vDaysPerMonth.ToString & ";" & vLeadTimeStock.ToString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Function

    Public Function CalculateOptimumStockQty(ByVal vAvgMonthly As Double, ByVal vOptimumStockDays As Integer, ByVal vDaysPerMonth As Integer, _
            ByRef vOptimumStockQty As Double, ByRef vExplainOptimumStock As String, ByVal vDec As Integer) As Boolean
        Try
            If vDaysPerMonth <= 0 Then
                vOptimumStockQty = 0
            Else
                vOptimumStockQty = Me.objCd.DigitGrouping(vAvgMonthly * vOptimumStockDays / vDaysPerMonth, vDec)
            End If
            vExplainOptimumStock = vOptimumStockDays.ToString & vbNewLine
            vExplainOptimumStock += vAvgMonthly.ToString & ";" & vOptimumStockDays.ToString & ";" & vDaysPerMonth.ToString & ";" & vOptimumStockQty.ToString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Function

    Public Function CalculatedOptimumStockDays(ByVal vAvgMonthly As Double, ByVal vOptimumStockQty As Double, ByVal vDaysPerMonth As Integer, _
           ByRef vOptimumStockDays As Integer, ByRef vExplainOptimumStock As String, ByVal vDec As Integer) As Boolean
        Try
            If vAvgMonthly <= 0 Then
                vOptimumStockDays = 0
            Else
                vOptimumStockDays = vOptimumStockQty * vDaysPerMonth / vAvgMonthly
            End If
            vExplainOptimumStock = vOptimumStockDays.ToString & vbNewLine
            vExplainOptimumStock += vAvgMonthly.ToString & ";" & vOptimumStockDays.ToString & ";" & vDaysPerMonth.ToString & ";" & vOptimumStockQty.ToString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Function

    Public Function CalculateAverageMonthlyDemandAll(ByVal vPrimaryProduct As Double, _
            ByVal vStartPeriode As Date, ByVal vFinishPeriode As Date, ByVal vDaysPerMonth As Integer, _
            ByVal vIsIncludeSales As Boolean, ByVal vIncludeBundling As Boolean, ByVal vIncludeAdjustment As Boolean, _
            ByRef vAvgMonthly As Double, ByRef vExplainAvgMonthly As String, _
            ByVal vTblHistoricalMaster As DseTrans_MaterialPlanningDashboard_Data.masterHistoricalDataTable, _
            ByVal vListSales As List(Of DseTrans_MaterialPlanningDashboard_Data.getSalesAndAdjustment_RDataTable), _
            ByVal vListBundling As List(Of DseTrans_MaterialPlanningDashboard_Data.getBundlingDataTable), _
            ByVal vListAdj As List(Of DseTrans_MaterialPlanningDashboard_Data.getSalesAndAdjustment_RDataTable), _
            ByVal indexSales As DseTrans_MaterialPlanningDashboard_Data.indexSalesDataTable, _
            ByVal indexBundling As DseTrans_MaterialPlanningDashboard_Data.indexBundlingDataTable, _
            ByVal indexAdj As DseTrans_MaterialPlanningDashboard_Data.indexAdjustmentDataTable, _
            ByVal vDec As Integer) As Boolean
        Dim valSales As Double = 0
        Dim valBundling As Double = 0
        Dim valAdjustment As Double = 0
        Dim valHistorical As Double = 0
        Dim valTotalPerMonth As Double = 0
        Dim valTotalUsage As Double = 0
        Dim tblSales As DseTrans_MaterialPlanningDashboard_Data.getSalesAndAdjustment_RDataTable
        Dim tblBundling As DseTrans_MaterialPlanningDashboard_Data.getBundlingDataTable
        Dim tblAdj As DseTrans_MaterialPlanningDashboard_Data.getSalesAndAdjustment_RDataTable
        Dim rowSales As DseTrans_MaterialPlanningDashboard_Data.getSalesAndAdjustment_RRow
        Dim rowBundling As DseTrans_MaterialPlanningDashboard_Data.getBundlingRow
        Dim rowAdj As DseTrans_MaterialPlanningDashboard_Data.getSalesAndAdjustment_RRow
        Dim rowHistoricalMaster As DseTrans_MaterialPlanningDashboard_Data.masterHistoricalRow
        Dim rowIndexSales As DseTrans_MaterialPlanningDashboard_Data.indexSalesRow
        Dim rowIndexBundling As DseTrans_MaterialPlanningDashboard_Data.indexBundlingRow
        Dim rowIndexAdj As DseTrans_MaterialPlanningDashboard_Data.indexAdjustmentRow

        Try
            vStartPeriode = Me.objCd.GetFirstDateofTheMonth(vStartPeriode)
            vFinishPeriode = Me.objCd.GetLasDateOfTheMonth(vFinishPeriode)
            vExplainAvgMonthly = Format(vStartPeriode, "MMM yyyy").ToString & vbNewLine
            vExplainAvgMonthly += Format(vFinishPeriode, "MMM yyyy").ToString & vbNewLine
            vExplainAvgMonthly += vDaysPerMonth.ToString & vbNewLine
            If vIsIncludeSales = True Then ''for sales
                vExplainAvgMonthly += "1" & vbNewLine
            Else
                vExplainAvgMonthly += "0" & vbNewLine
            End If
            If vIncludeBundling = True Then ''for bundling
                vExplainAvgMonthly += "1" & vbNewLine
            Else
                vExplainAvgMonthly += "0" & vbNewLine
            End If
            If vIncludeAdjustment = True Then ''for adjustment
                vExplainAvgMonthly += "1" & vbNewLine
            Else
                vExplainAvgMonthly += "0" & vbNewLine
            End If

            ''hitung per bulan::
            For i As Integer = 0 To DateDiff(DateInterval.Month, vStartPeriode, vFinishPeriode)
                ''get sales
                If vIsIncludeSales = True Then
                    indexSales.DefaultView.RowFilter = "[Period] = #" & Format(vStartPeriode.AddMonths(i), "MM/dd/yyyy").ToString & "#"
                    If indexSales.DefaultView.Count > 0 Then
                        rowIndexSales = indexSales.DefaultView(0).Row
                        tblSales = vListSales(rowIndexSales.Index_)
                        tblSales.DefaultView.RowFilter = "[Primary Product] = " & vPrimaryProduct.ToString
                        If tblSales.DefaultView.Count > 0 Then
                            rowSales = tblSales.DefaultView(0).Row
                            valSales = Me.objCd.DigitGrouping(rowSales.jml, vDec)
                            'DISINI TIDAK DIKURANGKAN LAGI DENGAN SALES RETURN KARENA ASUMSI LIST SALES YANG DIPROVIDE
                            'KE DALAM FUNGSI INI SUDAH MEMPERHITUNGKANNYA
                        Else
                            valSales = 0
                        End If
                    Else
                        valSales = 0
                    End If
                Else
                    valSales = 0
                End If
                ''get bundling
                If vIncludeBundling = True Then
                    indexBundling.DefaultView.RowFilter = "[Period] = #" & Format(vStartPeriode.AddMonths(i), "MM/dd/yyyy").ToString & "#"
                    If indexBundling.DefaultView.Count > 0 Then
                        rowIndexBundling = indexBundling.DefaultView(0).Row
                        tblBundling = vListBundling(rowIndexBundling.Index_)
                        tblBundling.DefaultView.RowFilter = "[Primary Product] = " & vPrimaryProduct.ToString
                        If tblBundling.DefaultView.Count > 0 Then
                            rowBundling = tblBundling.DefaultView(0).Row
                            valBundling = Me.objCd.DigitGrouping(rowBundling.jml, vDec)
                        Else
                            valBundling = 0
                        End If
                    Else
                        valBundling = 0
                    End If
                Else
                    valBundling = 0
                End If
                ''get adjustment
                If vIncludeAdjustment = True Then
                    indexAdj.DefaultView.RowFilter = "[Period] = #" & Format(vStartPeriode.AddMonths(i), "MM/dd/yyyy").ToString & "#"
                    If indexAdj.DefaultView.Count > 0 Then
                        rowIndexAdj = indexAdj.DefaultView(0).Row
                        tblAdj = vListAdj(rowIndexAdj.Index_)
                        tblAdj.DefaultView.RowFilter = "[Primary Product] = " & vPrimaryProduct.ToString
                        If tblAdj.DefaultView.Count > 0 Then
                            rowAdj = tblAdj.DefaultView(0).Row
                            valAdjustment = Me.objCd.DigitGrouping(rowAdj.jml, vDec)
                        Else
                            valAdjustment = 0
                        End If
                    Else
                        valAdjustment = 0
                    End If
                Else
                    valAdjustment = 0
                End If
                ''get historical
                If vTblHistoricalMaster.Rows.Count > 0 Then
                    vTblHistoricalMaster.DefaultView.RowFilter = "[Primary Product] = " & vPrimaryProduct.ToString & " AND [Periode] = #" & Format(vStartPeriode.AddMonths(i), "MM/dd/yyyy").ToString & "#"
                    If vTblHistoricalMaster.DefaultView.Count > 0 Then
                        rowHistoricalMaster = vTblHistoricalMaster.DefaultView(0).Row
                        valHistorical = Me.objCd.DigitGrouping(rowHistoricalMaster.jml, vDec)
                    Else
                        valHistorical = 0
                    End If
                Else
                    valHistorical = 0
                End If

                ''hitung per bulan
                valTotalPerMonth = valSales + valBundling + valAdjustment + valHistorical
                valTotalUsage += valTotalPerMonth

                ''tambahkan data bulanan ke explain
                vExplainAvgMonthly += Format(vStartPeriode.AddMonths(i), "MMM yyyy").ToString & ";" & valSales.ToString & ";" & valBundling.ToString & ";" & valAdjustment.ToString & ";" & valHistorical.ToString & ";" & valTotalPerMonth.ToString & vbNewLine

            Next
            ''hitung rata-rata
            vAvgMonthly = Me.objCd.DigitGrouping(valTotalUsage / (DateDiff(DateInterval.Month, vStartPeriode, vFinishPeriode) + 1), vDec)

            ''simpan rata-rata ke explain
            vExplainAvgMonthly += valTotalUsage.ToString & vbNewLine
            vExplainAvgMonthly += valTotalUsage.ToString & " / " & (DateDiff(DateInterval.Month, vStartPeriode, vFinishPeriode) + 1).ToString & " = " & vAvgMonthly.ToString

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Function

    Public Function GetListSales(ByRef vLSales As List(Of DseTrans_MaterialPlanningDashboard_Data.getSalesAndAdjustment_RDataTable), _
            ByVal vStart As Date, ByVal vFinish As Date, ByRef vTblIndexSales As DseTrans_MaterialPlanningDashboard_Data.indexSalesDataTable) As Boolean
        Dim ObjCD As New Core_Module.ClsCekData
        Dim dummyStartPeriode As Date
        Dim dta As New DseTrans_MaterialPlanningDashboard_DataTableAdapters.getSalesAndAdjustment_RTableAdapter
        Dim tblSales As New DseTrans_MaterialPlanningDashboard_Data.getSalesAndAdjustment_RDataTable
        Dim rowIndexSales As DseTrans_MaterialPlanningDashboard_Data.indexSalesRow
        Dim TblSalesReturnRevisiInv As New DseTrans_MaterialPlanningDashboard_Data.getSalesAndAdjustment_RDataTable
        Dim rowSales, rowSalesReturn As DseTrans_MaterialPlanningDashboard_Data.getSalesAndAdjustment_RRow

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dummyStartPeriode = vStart
            While dummyStartPeriode <= vFinish
                tblSales = New DseTrans_MaterialPlanningDashboard_Data.getSalesAndAdjustment_RDataTable
                dta.Fill(tblSales, clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Sales, Me.objCd.GetLasDateOfTheMonth(dummyStartPeriode), dummyStartPeriode)

                'kurangkan jumlahnya dengan revisi invoice
                dta.Fill(TblSalesReturnRevisiInv, clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn, _
                Me.objCd.GetLasDateOfTheMonth(dummyStartPeriode), dummyStartPeriode)
                For Each rowSalesReturn In TblSalesReturnRevisiInv.Rows
                    tblSales.DefaultView.RowFilter = "[Primary Product] = " + rowSalesReturn.Primary_Product.ToString
                    If tblSales.DefaultView.Count > 0 Then
                        rowSales = tblSales.DefaultView(0).Row
                        rowSales.jml -= rowSalesReturn.jml
                        If rowSales.jml <= 0 Then
                            tblSales.Rows.Remove(rowSales)
                        End If
                    End If
                    tblSales.DefaultView.RowFilter = ""
                Next

                vLSales.Add(tblSales)
                vTblIndexSales.DefaultView.RowFilter = "[Period] = #" & Format(dummyStartPeriode, "MM/dd/yyyy") & "#"
                If vTblIndexSales.DefaultView.Count = 0 Then
                    rowIndexSales = vTblIndexSales.NewindexSalesRow
                    rowIndexSales.Period = dummyStartPeriode
                    rowIndexSales.Index_ = vTblIndexSales.Rows.Count
                    vTblIndexSales.Rows.Add(rowIndexSales)
                End If
                dummyStartPeriode = dummyStartPeriode.AddMonths(1)
            End While
            'Dim rowSales As DseTrans_MaterialPlanningDashboard_Data.getSalesAndAdjustmentRow
            'tblSales = vLSales(0)
            'rowSales = tblSales.Rows(0)
            'Console.WriteLine(rowSales.jml)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
            TblSalesReturnRevisiInv.Dispose()
        End Try
    End Function

    Public Function GetListBundling(ByRef vLBundling As List(Of DseTrans_MaterialPlanningDashboard_Data.getBundlingDataTable), _
            ByVal vStart As Date, ByVal vFinish As Date, ByRef vTblIndexBundling As DseTrans_MaterialPlanningDashboard_Data.indexBundlingDataTable) As Boolean
        Dim ObjCD As New Core_Module.ClsCekData
        Dim dummyStartPeriode As Date
        Dim dta As New DseTrans_MaterialPlanningDashboard_DataTableAdapters.getBundlingTableAdapter
        Dim tblBundling As New DseTrans_MaterialPlanningDashboard_Data.getBundlingDataTable
        Dim rowIndexBundling As DseTrans_MaterialPlanningDashboard_Data.indexBundlingRow
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dummyStartPeriode = vStart
            While dummyStartPeriode <= vFinish
                tblBundling = New DseTrans_MaterialPlanningDashboard_Data.getBundlingDataTable
                dta.Fill(tblBundling, Me.objCd.GetLasDateOfTheMonth(dummyStartPeriode), dummyStartPeriode)
                vLBundling.Add(tblBundling)
                vTblIndexBundling.DefaultView.RowFilter = "[Period] = #" & Format(dummyStartPeriode, "MM/dd/yyyy") & "#"
                If vTblIndexBundling.DefaultView.Count = 0 Then
                    rowIndexBundling = vTblIndexBundling.NewindexBundlingRow
                    rowIndexBundling.Period = dummyStartPeriode
                    rowIndexBundling.Index_ = vTblIndexBundling.Rows.Count
                    vTblIndexBundling.Rows.Add(rowIndexBundling)
                End If
                dummyStartPeriode = dummyStartPeriode.AddMonths(1)
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetListAdj(ByRef vLAdj As List(Of DseTrans_MaterialPlanningDashboard_Data.getSalesAndAdjustment_RDataTable), _
                ByVal vStart As Date, ByVal vFinish As Date, ByRef vTblIndexAdj As DseTrans_MaterialPlanningDashboard_Data.indexAdjustmentDataTable) As Boolean
        Dim ObjCD As New Core_Module.ClsCekData
        Dim dummyStartPeriode As Date
        Dim dta As New DseTrans_MaterialPlanningDashboard_DataTableAdapters.getSalesAndAdjustment_RTableAdapter
        Dim tblAdj As New DseTrans_MaterialPlanningDashboard_Data.getSalesAndAdjustment_RDataTable
        Dim rowIndexAdj As DseTrans_MaterialPlanningDashboard_Data.indexAdjustmentRow
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            dummyStartPeriode = vStart
            While dummyStartPeriode <= vFinish
                tblAdj = New DseTrans_MaterialPlanningDashboard_Data.getSalesAndAdjustment_RDataTable
                dta.Fill(tblAdj, clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_Adjust, Me.objCd.GetLasDateOfTheMonth(dummyStartPeriode), dummyStartPeriode)
                vLAdj.Add(tblAdj)
                vTblIndexAdj.DefaultView.RowFilter = "[Period] = #" & Format(dummyStartPeriode, "MM/dd/yyyy") & "#"
                If vTblIndexAdj.DefaultView.Count = 0 Then
                    rowIndexAdj = vTblIndexAdj.NewindexAdjustmentRow
                    rowIndexAdj.Period = dummyStartPeriode
                    rowIndexAdj.Index_ = vTblIndexAdj.Rows.Count
                    vTblIndexAdj.Rows.Add(rowIndexAdj)
                End If
                dummyStartPeriode = dummyStartPeriode.AddMonths(1)
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function
End Class
