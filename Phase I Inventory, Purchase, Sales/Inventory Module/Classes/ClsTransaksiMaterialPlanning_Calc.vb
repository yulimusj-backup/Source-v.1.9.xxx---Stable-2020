Option Explicit On
Option Strict Off

Imports Core_Module

Public Class ClsTransaksiMaterialPlanning_Calc
    Private Sub IsikanRequiredDate(ByVal TBLWIP As DseTransaksiMaterialPlanning2.tbltrans_mo_item_wipDataTable)
        Dim rowWIP As DseTransaksiMaterialPlanning2.tbltrans_mo_item_wipRow
        Dim DTA1 As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_mo_item_rm1TableAdapter
        Dim TBL1 As New DseTransaksiMaterialPlanning2.tbltrans_mo_item_rm1DataTable
        Dim row1 As DseTransaksiMaterialPlanning2.tbltrans_mo_item_rm1Row
        Dim PrimaryMOMax, PrimaryMOMin As Double
        Dim StartIdx, Idx As Integer
        Dim DV As Data.DataRowView

        Try
            DTA1.Connection.ConnectionString = Core_Module.GetConstring
            PrimaryMOMax = Double.MinValue
            PrimaryMOMin = Double.MaxValue

            For Each rowWIP In TBLWIP.Rows
                If PrimaryMOMax < rowWIP.Primary_MO Then
                    PrimaryMOMax = rowWIP.Primary_MO
                End If
                If PrimaryMOMin > rowWIP.Primary_MO Then
                    PrimaryMOMin = rowWIP.Primary_MO
                End If
            Next

            DTA1.Fill(TBL1, PrimaryMOMax, PrimaryMOMin)

            StartIdx = 0
            TBLWIP.DefaultView.RowFilter = ""
            TBLWIP.DefaultView.Sort = "[Primary MO]"
            For Each DV In TBLWIP.DefaultView
                rowWIP = DV.Row
                For Idx = StartIdx To TBL1.Rows.Count - 1
                    row1 = TBL1.Rows(Idx)

                    If row1.Primary_MO < rowWIP.Primary_MO Then
                        StartIdx += 1
                    ElseIf row1.Primary_MO = rowWIP.Primary_MO Then
                        If row1.Identifier_JO = rowWIP.Identifier_JO And row1.Identifier_WO = rowWIP.Identifier_WO Then
                            rowWIP.Required_Date = row1.required_date
                            Exit For
                        End If
                    Else
                        Exit For
                    End If
                Next
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA1.Dispose()
            TBL1.Rows.Clear()
            TBL1.Dispose()
        End Try
    End Sub

    Public Sub HitungMaterialPlanning(ByVal vDT As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item1DataTable, _
    ByVal vDT_ItemGIT As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemgitDataTable, _
    ByVal vDT_ItemDO As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemdoDataTable, _
    ByVal vDT_ItemPOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempobDataTable, _
    ByVal vDT_ItemSOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsobDataTable, _
    ByVal vDT_ItemPrOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_reqDataTable, _
    ByVal vDT_ItemSettingLocation As DseTransaksiMaterialPlanning_WarehouseExc.tbltrans_material_planning_item_location_exclusionDataTable, _
    ByVal vDT_ItemQty As DseTransaksiMaterialPlanning_WarehouseExc.tbltrans_material_planning_item_qtyDataTable, _
    ByVal vDT_ItemMOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemmobDataTable, _
    ByVal vDT_ItemRMReq As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_rm_reqDataTable, _
    ByVal vObjSetting As Core_Module.ClsTransaksiSetting, ByRef vTxtMP_DaysPerMonth As Windows.Forms.TextBox, _
    ByVal vBackColorAll As Color, ByVal vFrmCek As FrmTrans_MaterialPlanningDashboard_DetailPO, ByVal vUserCode As String, _
    ByVal vObjCD As Core_Module.ClsCekData, ByVal vTxtMP_Primary As Windows.Forms.TextBox, _
    ByVal vObjSettingUser As Core_Module.ClsTransaksiSettingUser, ByVal vDtpMP_RefDate As Windows.Forms.DateTimePicker, _
    ByVal vDT_ItemSOBDelivery As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsob_deliveryDataTable, _
    ByVal vDT_ItemPrOBDelivery As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_req_deliveryDataTable)
        Dim ProductType As Core_Module.clsMasterProduct.EnumTipeProduct = Core_Module.clsMasterProduct.EnumTipeProduct.Goods
        Dim DTAProduct As New DseTransaksiMaterialPlanning2TableAdapters.DataTable1TableAdapter
        Dim TBLProduct As New DseTransaksiMaterialPlanning2.DataTable1DataTable
        Dim rowProduct As DseTransaksiMaterialPlanning2.DataTable1Row
        Dim rowItem As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item1Row
        Dim DTAGIT As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_packinglist_itemTableAdapter
        Dim TBLGIT As New DseTransaksiMaterialPlanning2.tbltrans_packinglist_itemDataTable
        Dim rowGIT As DseTransaksiMaterialPlanning2.tbltrans_packinglist_itemRow
        Dim rowGIT2 As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemgitRow
        Dim JmlDummy As Decimal = 0
        Dim JmlDummy1 As Decimal = 0
        Dim Filter As String = ""
        Dim DV As Data.DataRowView
        Dim DTAProductQty As New DseTransaksiMaterialPlanning2TableAdapters.tblmaster_productqtyTableAdapter
        Dim TBLProductQty As New DseTransaksiMaterialPlanning2.tblmaster_productqtyDataTable
        Dim rowProductQty As DseTransaksiMaterialPlanning2.tblmaster_productqtyRow
        Dim DTAPendingDO As New DseTransaksiMaterialPlanning2TableAdapters.DataTable2TableAdapter
        Dim TBLPendingDO As New DseTransaksiMaterialPlanning2.DataTable2DataTable
        Dim rowPendingDO As DseTransaksiMaterialPlanning2.DataTable2Row
        Dim rowPendingDO2 As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemdoRow
        Dim DTAPOB As New DseTransaksiMaterialPlanning2TableAdapters.DataTable3TableAdapter
        Dim TBLPOB As New DseTransaksiMaterialPlanning2.DataTable3DataTable
        Dim rowPOB As DseTransaksiMaterialPlanning2.DataTable3Row
        Dim rowPOB2 As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempobRow
        Dim DTAPOItem As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_po_itemTableAdapter
        Dim TBLPOItem As New DseTransaksiMaterialPlanning2.tbltrans_po_itemDataTable
        Dim rowPOItem As DseTransaksiMaterialPlanning2.tbltrans_po_itemRow
        Dim SupplierSelectionMode As Core_Module.clsMasterProduct.EnumSupplierSelectionMode = Core_Module.clsMasterProduct.EnumSupplierSelectionMode.Prefered_Supplier
        Dim CekOtorisasiPO As Boolean = True
        Dim DTASOB As New DseTransaksiMaterialPlanning2TableAdapters.DataTable4TableAdapter
        Dim TBLSOB As New DseTransaksiMaterialPlanning2.DataTable4DataTable
        Dim rowSOB As DseTransaksiMaterialPlanning2.DataTable4Row
        Dim rowSOB2 As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsobRow
        Dim PrimarySOMin, PrimarySOMax As Double
        Dim AdaDetailDelivery As Boolean = False
        Dim PrimaryPrOBMin, PrimaryPrOBMax As Double
        Dim AdaDetailDeliveryPrOB As Boolean = False
        Dim OrderStrategy As Core_Module.clsMasterProduct.EnumOrderStrategy = Core_Module.clsMasterProduct.EnumOrderStrategy.Order_By_Maximum_Order_Qty
        Dim JmlDaysPerMonth As Integer = 0
        Dim rowDGV As DataGridViewRow
        Dim FreeStock As Decimal = 0
        Dim IndexDummy As Integer = 0
        Dim ProcurementStrategy As Core_Module.clsMasterProduct.EnumProcurementStrategy = Core_Module.clsMasterProduct.EnumProcurementStrategy.Build_To_Stock
        Dim DTAPrOB As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planning_itempurch_req1TableAdapter
        Dim rowPrOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_req1Row
        Dim rowPrOB2 As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_reqRow
        Dim TBLPrOB As New DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_req1DataTable
        Dim DTAPrOBDeliveryDummy As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_purchase_req_itemdeliveryTableAdapter
        Dim TBLPrOBDeliveryDummy As New DseTransaksiMaterialPlanning2.tbltrans_purchase_req_itemdeliveryDataTable
        Dim RowPrOBDeliveryDummy As DseTransaksiMaterialPlanning2.tbltrans_purchase_req_itemdeliveryRow
        Dim DTAMultiUnit As New DseTransaksiMaterialPlanning2TableAdapters.tblmaster_product_item_multiunitTableAdapter
        Dim TBLMultiUnit As New DseTransaksiMaterialPlanning2.tblmaster_product_item_multiunitDataTable
        Dim rowUnitMulti As DseTransaksiMaterialPlanning2.tblmaster_product_item_multiunitRow
        Dim ijkl As New FrmTrans_MaterialPlanningDashboard_Adv_Exclusion
        Dim vModifySupplierSelectionMode As Boolean
        Dim vPreferedSupplierSelectionMode As Core_Module.clsMasterProduct.EnumSupplierSelectionMode = clsMasterProduct.EnumSupplierSelectionMode.The_cheapest_in_the_last_3_order
        Dim rowProductQty1 As DseTransaksiMaterialPlanning_WarehouseExc.tbltrans_material_planning_item_qtyRow
        Dim JmlTelat As Double = 0
        Dim JmlKompensasi As Decimal = 0
        Dim SSDate As Date = Today
        Dim DTAMOB As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_mo_itemTableAdapter
        Dim TBLMOB As New DseTransaksiMaterialPlanning2.tbltrans_mo_itemDataTable
        Dim rowMOB As DseTransaksiMaterialPlanning2.tbltrans_mo_itemRow
        Dim rowMOB2 As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemmobRow
        Dim DTARMReq As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_mo_item_rmTableAdapter
        Dim TBLRMReq As New DseTransaksiMaterialPlanning2.tbltrans_mo_item_rmDataTable
        Dim rowRMReq As DseTransaksiMaterialPlanning2.tbltrans_mo_item_rmRow
        Dim rowRMReq2 As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_rm_reqRow
        Dim DTAWIP As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_mo_item_wipTableAdapter
        Dim TBLWIP As New DseTransaksiMaterialPlanning2.tbltrans_mo_item_wipDataTable
        Dim rowWIP As DseTransaksiMaterialPlanning2.tbltrans_mo_item_wipRow 

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor 

            PrimarySOMin = Double.MaxValue
            PrimarySOMax = Double.MinValue
            PrimaryPrOBMax = Double.MinValue
            PrimaryPrOBMin = Double.MaxValue

            DTAProduct.Connection.ConnectionString = Core_Module.GetConstring
            DTAGIT.Connection.ConnectionString = DTAProduct.Connection.ConnectionString
            DTAProductQty.Connection.ConnectionString = DTAProduct.Connection.ConnectionString
            DTAPendingDO.Connection.ConnectionString = DTAProduct.Connection.ConnectionString
            DTAPOB.Connection.ConnectionString = DTAProduct.Connection.ConnectionString
            DTASOB.Connection.ConnectionString = DTAProduct.Connection.ConnectionString
            DTAPOItem.Connection.ConnectionString = DTAProduct.Connection.ConnectionString
            DTAPrOB.Connection.ConnectionString = DTAProduct.Connection.ConnectionString
            DTAPrOBDeliveryDummy.Connection.ConnectionString = DTAProduct.Connection.ConnectionString
            DTAMultiUnit.Connection.ConnectionString = DTAProduct.Connection.ConnectionString
            DTAMOB.Connection.ConnectionString = DTAProduct.Connection.ConnectionString
            DTARMReq.Connection.ConnectionString = DTAProduct.Connection.ConnectionString
            DTAWIP.Connection.ConnectionString = DTAProduct.Connection.ConnectionString

            vDT.Rows.Clear()
            vDT_ItemGIT.Rows.Clear()
            vDT_ItemDO.Rows.Clear()
            vDT_ItemPOB.Rows.Clear()
            vDT_ItemSOB.Rows.Clear()
            vDT_ItemSOBDelivery.Rows.Clear()
            vDT_ItemPrOB.Rows.Clear()
            vDT_ItemPrOBDelivery.Rows.Clear()
            vDT_ItemQty.Rows.Clear()
            vDT_ItemSettingLocation.Rows.Clear()
            vDT_ItemRMReq.Rows.Clear()
            vDT_ItemMOB.Rows.Clear()
            vDT_ItemRMReq.Rows.Clear()

            TBLMultiUnit.Rows.Clear()

            If vObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.MaterialPlanning, _
                ClsTransaksiSetting.EnumNamaParameter.MaterialPlanning_JmlHariKerjaPerBulan) = "" Then
                MsgBox("Please specify days quantity per month for material planning calculation", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            JmlDaysPerMonth = vObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.MaterialPlanning, _
            ClsTransaksiSetting.EnumNamaParameter.MaterialPlanning_JmlHariKerjaPerBulan)
            vTxtMP_DaysPerMonth.Text = JmlDaysPerMonth

            'isikan semua data dasar
            'DTAProduct.Fill(TBLProduct, ProductType)-> di ubah ketika memasukan perhitungan purchase request, jasa skrg diperhitungkan
            DTAProduct.FillBy_all(TBLProduct)
            'cek untuk memastikan semua product, optimum stock qty > safety stock qty
            'untuk product yang build to stock
            Filter = ""
            JmlDummy = 0
            For Each rowProduct In TBLProduct.Rows
                ProcurementStrategy = rowProduct.procurement_strategy
                If ProcurementStrategy = clsMasterProduct.EnumProcurementStrategy.Build_To_Order Then
                    If rowProduct.optimum_stock < rowProduct.safety_stock_qty Then
                        JmlDummy += 1
                        If Filter = "" Then
                            Filter = JmlDummy.ToString + ". " + rowProduct.Product_Code
                        Else
                            Filter = Filter + vbCrLf + JmlDummy.ToString + ". " + rowProduct.Product_Code
                        End If
                    End If
                End If
            Next
            If Filter <> "" Then
                MsgBox("The following product is using 'Build To Stock' strategy but optimum stock qty < safety stock qty, please fix this." + _
                vbCrLf + Filter, MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            JmlDummy = 0

            DTAGIT.Fill(TBLGIT)
            DTAProductQty.Fill(TBLProductQty)
            DTAPendingDO.Fill(TBLPendingDO)
            DTAPOB.Fill(TBLPOB)
            DTASOB.Fill(TBLSOB)
            DTAPOItem.Fill(TBLPOItem)
            DTAPrOB.Fill(TBLPrOB)
            DTAMultiUnit.Fill(TBLMultiUnit)
            DTAMOB.Fill(TBLMOB)
            DTARMReq.Fill(TBLRMReq)
            'tambahkan qty wip ke dalam tblrm req
            DTAWIP.Fill(TBLWIP)
            If TBLWIP.Rows.Count > 0 Then
                Me.IsikanRequiredDate(TBLWIP)
            End If
            For Each rowWIP In TBLWIP.Rows
                rowRMReq = TBLRMReq.Newtbltrans_mo_item_rmRow
                With rowRMReq
                    .Balance_Qty = rowWIP.Qty_Balance
                    .Identifier_JO = rowWIP.Identifier_JO
                    .Identifier_WO = rowWIP.Identifier_WO
                    .JO_Number = rowWIP.JO_Number
                    .MO_Date = rowWIP._Date
                    .MO_Number = rowWIP.MO_Number
                    .Primary_MO = rowWIP.Primary_MO
                    .Primary_Process = rowWIP.Primary_Process
                    .Primary_Product = rowWIP.Primary_Product_Code_WIP
                    .WO_Number = rowWIP.WO_Number
                    .is_wip = 1
                    .MP_Spare_Qty = 0
                    .required_date = rowWIP.Required_Date
                End With
                TBLRMReq.Rows.Add(rowRMReq)
            Next

            'untuk pob, tampilkan form peringatan bila ada po yang otorisasinya tidak sempurna
            CekOtorisasiPO = False
            For Each rowPOB In TBLPOB.Rows
                If rowPOB.Authorization_Level <> rowPOB.Authorization_Level_Needed Then
                    CekOtorisasiPO = True
                    Exit For
                End If
            Next
            If CekOtorisasiPO = True Then
                vFrmCek.vDT_ItemPOB = TBLPOB
                vFrmCek.BackColorAll = vBackColorAll
                vFrmCek.ShowDialog()

                'buang item po yang berasal dari transaksi yang tidak includekan
                For Each rowDGV In vFrmCek.dgvItem.Rows
                    If rowDGV.Cells(vFrmCek.ColInclude.Index).Value = Boolean.FalseString Then
                        TBLPOB.DefaultView.RowFilter = "[Primary PO] = " + rowDGV.Cells(vFrmCek.ColPrimaryPO.Index).Value.ToString
                        While TBLPOB.DefaultView.Count > 0
                            rowPOB = TBLPOB.DefaultView(0).Row
                            TBLPOB.RemoveDataTable3Row(rowPOB)
                        End While
                    End If
                Next
            End If

            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'product exclusion
            ijkl.UserCode = vUserCode
            ijkl.TipeTransaksiExclusion = ClsTransaksiMaterialPlanning.EnumTipeTransaksiExclusion.MaterialPlanning
            ijkl.ShowDialog()
            Me.BuangProductExclude(TBLProduct, vUserCode, TBLRMReq, TBLMOB, TBLGIT, TBLPendingDO, TBLPOB, TBLSOB, TBLPrOB, TBLPOItem, TBLProductQty)
            If ijkl.ModifySupplierSelectionMode = True Then
                vModifySupplierSelectionMode = True
                vPreferedSupplierSelectionMode = ijkl.SupplierSelectionMode
            End If

            'untuk baris product yang di exclude tidak dimasukkan ke dalam perhitungan
            IndexDummy = 0
            While IndexDummy <= TBLProduct.Rows.Count - 1
                rowProduct = TBLProduct.Rows(IndexDummy)
                If rowProduct.exclude_on_material_planning = 1 Then
                    'yang di exclude tapi masih memiliki sob, pending do, rm req, pr-ob tetap tidak di exclude
                    TBLSOB.DefaultView.RowFilter = "[Primary Product] = " + rowProduct.Primary.ToString
                    If TBLSOB.DefaultView.Count > 0 Then
                        IndexDummy += 1
                    Else
                        TBLPendingDO.DefaultView.RowFilter = "[Primary Product] = " + rowProduct.Primary.ToString
                        If TBLPendingDO.DefaultView.Count > 0 Then
                            IndexDummy += 1
                        Else
                            TBLRMReq.DefaultView.RowFilter = "[Primary Product] = " + rowProduct.Primary.ToString
                            If TBLRMReq.DefaultView.Count > 0 Then
                                IndexDummy += 1
                            Else
                                TBLPrOB.DefaultView.RowFilter = "[primary code no] = " + rowProduct.Primary.ToString
                                If TBLPrOB.DefaultView.Count > 0 Then
                                    IndexDummy += 1
                                Else
                                    TBLProduct.Rows.Remove(rowProduct)
                                End If
                            End If
                        End If
                    End If
                Else
                    IndexDummy += 1
                End If
            End While
            TBLSOB.DefaultView.RowFilter = ""
            TBLPendingDO.DefaultView.RowFilter = ""
            TBLPrOB.DefaultView.RowFilter = ""
            TBLRMReq.DefaultView.RowFilter = ""

            'location exception
            If vObjSettingUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.MaterialPlanning, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.MaterialPlanning_UseWarehouseExclusion, vUserCode) = "1" Then
                Me.IsikanSetting_ProductQty(TBLProductQty, vUserCode, vDT_ItemSettingLocation, vTxtMP_Primary)
            Else
                vDT_ItemSettingLocation.Rows.Clear()
            End If

            TBLRMReq.DefaultView.Sort = "[MO Date], [MO Number], [JO Number], [WO Number]"
            TBLGIT.DefaultView.Sort = "[Tanggal Transaksi], [No Transaksi]"
            TBLMOB.DefaultView.Sort = "[MO Date], [MO Number], [JO Number]"
            TBLPendingDO.DefaultView.Sort = "[Tanggal Transaksi], [No Transaksi]"
            TBLPOB.DefaultView.Sort = "[Tanggal Transaksi], [No Transaksi]"
            TBLSOB.DefaultView.Sort = "[Tanggal Transaksi], [No Transaksi]"
            TBLPrOB.DefaultView.Sort = "[tanggal], [no transaksi]"
            TBLPOItem.DefaultView.Sort = "[Tanggal Transaksi] DESC, [Primary Transaksi] DESC"

            For Each rowProduct In TBLProduct.Rows
                rowItem = vDT.Newtbltrans_material_planning_item1Row
                With rowItem
                    .modify = 0
                    .default_prec_unit = rowProduct._Decimal
                    .Description = rowProduct.Description
                    .order_strategy = rowProduct.recomended_order_qty_type
                    .primary = vDT.Rows.Count + 1
                    .primary_code_no = rowProduct.Primary
                    .primary_transaksi = vTxtMP_Primary.Text
                    .primary_unit = rowProduct.Primary_Unit
                    .procurement_type = rowProduct.procurement_type
                    .Product_Code = rowProduct.Product_Code
                    .Order_Unit_Code = rowProduct.Unit_Code
                    .primary_order_unit = rowProduct.Primary_Unit
                    .Primary_Category = rowProduct.Primary_Category
                    .Primary_Sub_Brand = rowProduct.Primary_Sub_Brand
                    .Primary_Brand = rowProduct.Primary_Brand
                    .Primary_Product_Goup = rowProduct.Primary_Product_Group

                    .qty_avg_montly_demand = rowProduct.avg_montly_demand
                    .qty_lead_time = rowProduct.lead_time_qty
                    .lead_time_days = rowProduct.lead_time_days
                    .qty_max = rowProduct.maximum_stock
                    .qty_min = rowProduct.Minimum_Stock
                    .qty_optimum = rowProduct.optimum_stock
                    .qty_order_maksimum = rowProduct.maximum_order_qty
                    .qty_order_minimum = rowProduct.minimum_order_qty
                    .qty_order_optimum = rowProduct.optimum_order_qty
                    .qty_safety_stock = rowProduct.safety_stock_qty
                    .Unit_Code = rowProduct.Unit_Code
                    .supplier_selection_mode = rowProduct.supplier_selection_mode
                    .Explain_Lead_Time = rowProduct.explain_lead_time_qty
                    .Explain_Safety_Stock = rowProduct.explain_safety_stock_qty
                    .ExplainAvgQty = rowProduct.explain_avg_qty
                    .explain_optimum_stock = rowProduct.explain_optimum_stock_qty
                    .SafetyStockStrategy = rowProduct.safety_stock_strategy
                    .Lead_Time_Strategy = rowProduct.lead_time_strategy
                    .procurement_strategy = rowProduct.procurement_strategy
                    .optimum_stock_strategy = rowProduct.optimum_stock_strategy
                    .product_type = rowProduct.Type

                    Filter = "[Primary Product] = " + rowProduct.Primary.ToString
                    TBLGIT.DefaultView.RowFilter = Filter
                    If TBLGIT.DefaultView.Count = 0 Then
                        .qty_git = 0
                    Else
                        JmlDummy = 0
                        For Each DV In TBLGIT.DefaultView
                            rowGIT = DV.Row
                            rowGIT2 = vDT_ItemGIT.Newtbltrans_material_planning_itemgitRow
                            With rowGIT2
                                .No_Packing_List = rowGIT.No_Transaksi
                                .packing_list_date_rec = rowGIT.Tanggal_Transaksi
                                .packing_list_no_rec = rowGIT.No_Transaksi
                                .primary = Me.GetPrimaryMax(vDT_ItemGIT) + 1
                                .primary_packing_list_rec = rowGIT.Primary_Packing_List
                                .primary_product = rowGIT.Primary_Product
                                .primary_suppier_rec = rowGIT.Primary_Supplier
                                .primary_transaksi = vTxtMP_Primary.Text
                                .qty = rowGIT.Total_Qty
                                .Supplier_Packing_List_No = rowGIT.Supplier_Packing_List_No
                                .supplier_packing_list_no_rec = rowGIT.Supplier_Packing_List_No
                                .Tgl_Packing_List = rowGIT.Tanggal_Transaksi
                            End With
                            vDT_ItemGIT.Rows.Add(rowGIT2)
                            JmlDummy += rowGIT.Total_Qty
                        Next
                        While TBLGIT.DefaultView.Count > 0
                            TBLGIT.Rows.Remove(TBLGIT.DefaultView(0).Row)
                        End While
                        .qty_git = JmlDummy
                    End If

                    If rowProduct.Type = Core_Module.clsMasterProduct.EnumTipeProduct.Services Then
                        .qty_on_hand = 0
                    Else
                        Filter = "[Primary Product] = " + rowProduct.Primary.ToString
                        TBLProductQty.DefaultView.RowFilter = Filter

                        If TBLProductQty.DefaultView.Count = 0 Then
                            .qty_on_hand = 0
                        Else
                            JmlDummy = 0
                            For Each DV In TBLProductQty.DefaultView
                                rowProductQty = DV.Row
                                JmlDummy += rowProductQty.Qty

                                'tambahkan datanya ke item qty
                                If rowProductQty.Qty > 0 Then
                                    rowProductQty1 = vDT_ItemQty.Newtbltrans_material_planning_item_qtyRow
                                    With rowProductQty1
                                        .Expiry_Date = rowProductQty.Expiry_Date
                                        .Lot_No = rowProductQty.Lot_No
                                        .Primary = vDT_ItemQty.Rows.Count + 1
                                        .Primary_Product = rowProductQty.Primary_Product
                                        .Primary_Storage_Area = rowProductQty.Primary_Storage_Area
                                        .primary_transaksi = vTxtMP_Primary.Text
                                        .Qty = rowProductQty.Qty
                                    End With
                                    vDT_ItemQty.Rows.Add(rowProductQty1)
                                End If
                            Next
                            While TBLProductQty.DefaultView.Count > 0
                                TBLProductQty.Rows.Remove(TBLProductQty.DefaultView(0).Row)
                            End While
                            .qty_on_hand = JmlDummy
                        End If
                    End If

                    .qty_mob = 0
                    Filter = "[Primary Product] = " + rowProduct.Primary.ToString
                    TBLMOB.DefaultView.RowFilter = Filter
                    If TBLMOB.DefaultView.Count = 0 Then
                        .qty_mob = 0
                    Else
                        JmlDummy = 0
                        For Each DV In TBLMOB.DefaultView
                            rowMOB2 = vDT_ItemMOB.Newtbltrans_material_planning_itemmobRow
                            rowMOB = DV.Row

                            With rowMOB2
                                .identifier_jo_rec = rowMOB.Identifier_JO
                                .jo_no_rec = rowMOB.JO_Number
                                .mo_date_rec = rowMOB.MO_Date
                                .mo_no_rec = rowMOB.MO_Number
                                .primary = -1
                                .primary_mo_rec = rowMOB.Primary_MO
                                .primary_product = rowMOB.Primary_Product
                                .primary_transaksi = vTxtMP_Primary.Text
                                .qty = rowMOB.Balance_Qty
                                .primary_proses = rowMOB.Primary_Process
                                .identifier_wo_rec = rowMOB.Identifier_WO
                                .wo_no_rec = rowMOB.WO_Number

                                JmlDummy += .qty
                            End With
                            vDT_ItemMOB.Rows.Add(rowMOB2)
                        Next
                        While TBLMOB.DefaultView.Count > 0
                            TBLMOB.Rows.Remove(TBLMOB.DefaultView(0).Row)
                        End While
                        .qty_mob = JmlDummy
                    End If

                    .qty_rm_req = 0
                    .MP_Spare_Qty_Rm_Req = 0
                    Filter = "[Primary Product] = " + rowProduct.Primary.ToString
                    TBLRMReq.DefaultView.RowFilter = Filter
                    If TBLRMReq.DefaultView.Count = 0 Then
                        .qty_rm_req = 0
                        .MP_Spare_Qty_Rm_Req = 0
                    Else
                        JmlDummy = 0
                        JmlDummy1 = 0
                        For Each DV In TBLRMReq.DefaultView
                            rowRMReq2 = vDT_ItemRMReq.Newtbltrans_material_planning_item_rm_reqRow
                            rowRMReq = DV.Row

                            With rowRMReq2
                                .identifier_wo_rec = rowRMReq.Identifier_WO
                                .wo_no_rec = rowRMReq.WO_Number
                                .identifier_jo_rec = rowRMReq.Identifier_JO
                                .jo_no_rec = rowRMReq.JO_Number
                                .mo_date_rec = rowRMReq.MO_Date
                                .mo_no_rec = rowRMReq.MO_Number
                                .primary = -1
                                .primary_mo_rec = rowRMReq.Primary_MO
                                .primary_product = rowRMReq.Primary_Product
                                .primary_transaksi = vTxtMP_Primary.Text
                                .primary_proses = rowRMReq.Primary_Process
                                .qty = rowRMReq.Balance_Qty
                                .is_wip = rowRMReq.is_wip
                                .MP_Spare_Qty = rowRMReq.MP_Spare_Qty
                                .required_date = rowRMReq.required_date

                                JmlDummy += .qty
                                JmlDummy1 += .MP_Spare_Qty
                            End With
                            vDT_ItemRMReq.Rows.Add(rowRMReq2)
                        Next
                        While TBLRMReq.DefaultView.Count > 0
                            TBLRMReq.Rows.Remove(TBLRMReq.DefaultView(0).Row)
                        End While
                        .qty_rm_req = JmlDummy
                        .MP_Spare_Qty_Rm_Req = JmlDummy1
                    End If

                    .qty_pending_do = 0
                    Filter = "[Primary Product] = " + rowProduct.Primary.ToString
                    TBLPendingDO.DefaultView.RowFilter = Filter
                    If TBLPendingDO.DefaultView.Count = 0 Then
                        .qty_pending_do = 0
                    Else
                        JmlDummy = 0
                        For Each DV In TBLPendingDO.DefaultView
                            rowPendingDO = DV.Row
                            rowPendingDO2 = vDT_ItemDO.Newtbltrans_material_planning_itemdoRow
                            With rowPendingDO2
                                .DO_Date = rowPendingDO.Tanggal_Transaksi
                                .do_date_rec = rowPendingDO.Tanggal_Transaksi
                                .do_no_rec = rowPendingDO.No_Transaksi
                                .DO_Number = rowPendingDO.No_Transaksi
                                .primary = Me.GetPrimaryMax_PendingDO(vDT_ItemDO) + 1
                                .primary_customer_rec = rowPendingDO.Primary_Customer
                                .primary_do_rec = rowPendingDO.Primary_DO
                                .primary_product = rowPendingDO.Primary_Product
                                .primary_transaksi = vTxtMP_Primary.Text
                                .qty = rowPendingDO.Total_Qty
                            End With
                            vDT_ItemDO.Rows.Add(rowPendingDO2)
                            JmlDummy += rowPendingDO.Total_Qty
                        Next
                        While TBLPendingDO.DefaultView.Count > 0
                            TBLPendingDO.Rows.Remove(TBLPendingDO.DefaultView(0).Row)
                        End While
                        .qty_pending_do = JmlDummy
                    End If

                    Filter = "[Primary Product] = " + rowProduct.Primary.ToString
                    TBLPOB.DefaultView.RowFilter = Filter
                    If TBLPOB.DefaultView.Count = 0 Then
                        .qty_pob = 0
                    Else
                        JmlDummy = 0
                        For Each DV In TBLPOB.DefaultView
                            rowPOB = DV.Row
                            rowPOB2 = vDT_ItemPOB.Newtbltrans_material_planning_itempobRow
                            With rowPOB2
                                .No_Transaksi = rowPOB.No_Transaksi
                                .po_date_rec = rowPOB.Tanggal_Transaksi
                                .po_no_rec = rowPOB.No_Transaksi
                                .primary = Me.GetPrimaryMax_POB(vDT_ItemPOB) + 1
                                .primary_po_rec = rowPOB.Primary_PO
                                .primary_product = rowPOB.Primary_Product
                                .primary_supplier_rec = rowPOB.Primary_Supplier
                                .primary_transaksi = vTxtMP_Primary.Text
                                .qty = rowPOB.POB
                                .Tanggal_Transaksi = rowPOB.Tanggal_Transaksi
                            End With
                            vDT_ItemPOB.Rows.Add(rowPOB2)
                            JmlDummy += rowPOB.POB
                        Next
                        While TBLPOB.DefaultView.Count > 0
                            TBLPOB.Rows.Remove(TBLPOB.DefaultView(0).Row)
                        End While
                        .qty_pob = JmlDummy
                    End If

                    .qty_sob = 0
                    .MP_Spare_Qty_SOB = 0
                    Filter = "[Primary Product] = " + rowProduct.Primary.ToString
                    TBLSOB.DefaultView.RowFilter = Filter
                    If TBLSOB.DefaultView.Count = 0 Then
                        .qty_sob = 0
                        .MP_Spare_Qty_SOB = 0
                    Else
                        JmlDummy = 0
                        JmlDummy1 = 0
                        For Each DV In TBLSOB.DefaultView
                            rowSOB = DV.Row
                            rowSOB2 = vDT_ItemSOB.Newtbltrans_material_planning_itemsobRow
                            With rowSOB2
                                .No_Transaksi = rowSOB.No_Transaksi
                                .primary = Me.GetPrimaryMax_SOB(vDT_ItemSOB) + 1
                                .primary_customer_rec = rowSOB.Primary_Customer
                                .primary_product = rowSOB.Primary_Product
                                .primary_so_rec = rowSOB.Primary_Transaksi
                                .primary_transaksi = rowSOB.Primary_Transaksi
                                .qty = rowSOB.SOB
                                .so_date_rec = rowSOB.Tanggal_Transaksi
                                .so_no_rec = rowSOB.No_Transaksi
                                .Tanggal_Transaksi = rowSOB.Tanggal_Transaksi
                                .MP_Spare_Qty = rowSOB.MP_Spare_Qty
                                If rowSOB.Delivery_Date_Cont = True Then
                                    .delivery_date_cont = 1
                                Else
                                    .delivery_date_cont = 0
                                End If
                                .qty_sob_multi = rowSOB.SOB_Multi
                                .ratio_default = rowSOB.Ratio_Default
                                .ratio_reference = rowSOB.Ratio_Reference
                                .primary_unit_multi = rowSOB.Primary_Units_Multi
                                .identifier_so = rowSOB.identifier
                                .primary_unit = rowSOB.Primary_Units
                            End With
                            vDT_ItemSOB.Rows.Add(rowSOB2)
                            JmlDummy += rowSOB.SOB
                            JmlDummy1 += rowSOB.MP_Spare_Qty
                            If rowSOB2.delivery_date_cont = 1 Then
                                AdaDetailDelivery = True
                                If PrimarySOMax < rowSOB.Primary_Transaksi Then
                                    PrimarySOMax = rowSOB.Primary_Transaksi
                                End If
                                If PrimarySOMin > rowSOB.Primary_Transaksi Then
                                    PrimarySOMin = rowSOB.Primary_Transaksi
                                End If
                            End If
                        Next
                        While TBLSOB.DefaultView.Count > 0
                            TBLSOB.Rows.Remove(TBLSOB.DefaultView(0).Row)
                        End While
                        .qty_sob = JmlDummy
                        .MP_Spare_Qty_SOB = JmlDummy1
                    End If

                    .qty_prob = 0
                    .MP_Spare_Qty_PrOB = 0
                    Filter = "[primary code no] = " + rowProduct.Primary.ToString
                    TBLPrOB.DefaultView.RowFilter = Filter
                    If TBLPrOB.DefaultView.Count = 0 Then
                        .qty_prob = 0
                        .MP_Spare_Qty_PrOB = 0
                    Else
                        JmlDummy = 0
                        JmlDummy1 = 0
                        For Each DV In TBLPrOB.DefaultView
                            rowPrOB = DV.Row
                            rowPrOB2 = vDT_ItemPrOB.Newtbltrans_material_planning_itempurch_reqRow
                            With rowPrOB2
                                .No_Transaksi = rowPrOB.no_transaksi
                                .primary = Me.GetPrimaryMax_PrOB(vDT_ItemPrOB) + 1
                                .primary_pr_rec = rowPrOB.primary_ref_party
                                .primary_product = rowPrOB.primary_code_no
                                .primary_pr_rec = rowPrOB.primary
                                .primary_transaksi = rowPrOB.primary
                                .qty = rowPrOB.OB
                                .MP_Spare_Qty = rowPrOB.MP_Spare_Qty
                                .pr_date_rec = rowPrOB.tanggal
                                .pr_no_rec = rowPrOB.no_transaksi
                                .Tanggal_Transaksi = rowPrOB.tanggal
                                .primary_ref_rec = rowPrOB.primary_ref_party
                                .remarks = rowPrOB.remarks_header
                                .remarks = .remarks.Trim
                                rowPrOB.remarks_item = rowPrOB.remarks_item.Trim
                                If rowPrOB.remarks_item <> "" Then
                                    If .remarks = "" Then
                                        .remarks = rowPrOB.remarks_item
                                    Else
                                        .remarks = .remarks + " - " + rowPrOB.remarks_item
                                    End If
                                End If

                                .delivery_date_cont = rowPrOB.delivery_date_cont
                                If .delivery_date_cont = 1 Then
                                    AdaDetailDeliveryPrOB = True
                                    If PrimaryPrOBMax < .primary_pr_rec Then
                                        PrimaryPrOBMax = .primary_pr_rec
                                    End If
                                    If PrimaryPrOBMin > .primary_pr_rec Then
                                        PrimaryPrOBMin = .primary_pr_rec
                                    End If
                                End If
                                .identifier_pr_item = rowPrOB.identifier
                            End With
                            vDT_ItemPrOB.Rows.Add(rowPrOB2)
                            JmlDummy += rowPrOB.OB
                            JmlDummy1 += rowPrOB.MP_Spare_Qty
                        Next
                        While TBLPrOB.DefaultView.Count > 0
                            TBLPrOB.Rows.Remove(TBLPrOB.DefaultView(0).Row)
                        End While
                        .qty_prob = JmlDummy
                        .MP_Spare_Qty_PrOB = JmlDummy1
                    End If

                    .qty_inquiry = 0

                    FreeStock = .qty_on_hand + .qty_pob + .qty_mob + .qty_git - .qty_inquiry - .qty_sob - .qty_prob - .qty_pending_do - .qty_rm_req - _
                    .MP_Spare_Qty_PrOB - .MP_Spare_Qty_Rm_Req - .MP_Spare_Qty_SOB


                    OrderStrategy = .order_strategy
                    ProcurementStrategy = rowProduct.procurement_strategy

                    If ProcurementStrategy = clsMasterProduct.EnumProcurementStrategy.Build_To_Stock Then
                        If OrderStrategy = clsMasterProduct.EnumOrderStrategy.Order_By_Maximum_Order_Qty Or _
                                          OrderStrategy = clsMasterProduct.EnumOrderStrategy.Order_By_Minimum_Order_Qty Or _
                                          OrderStrategy = clsMasterProduct.EnumOrderStrategy.Order_By_Optimum_Order_Qty Then
                            .qty_required = .qty_order_optimum
                        ElseIf OrderStrategy = clsMasterProduct.EnumOrderStrategy.Order_To_Max_Qty Or _
                        OrderStrategy = clsMasterProduct.EnumOrderStrategy.Order_To_Min_Qty Or _
                        OrderStrategy = clsMasterProduct.EnumOrderStrategy.Order_To_Optimum_Qty Or _
                        OrderStrategy = clsMasterProduct.EnumOrderStrategy.Order_To_Safety_Stock Then
                            If FreeStock > .qty_safety_stock Then
                                .qty_required = .qty_optimum - .qty_safety_stock
                            Else
                                .qty_required = .qty_optimum - FreeStock
                            End If
                        End If

                        .qty_order = .qty_required

                        If .qty_avg_montly_demand = 0 Then
                            .qty_avg_montly_demand = 1
                        End If

                        Me.HitungOrderPoint(rowItem, JmlDaysPerMonth, vDtpMP_RefDate, vObjCD)

                        If .order_point.Date < vDtpMP_RefDate.Value.Date Then
                            If FreeStock > 0 Then
                                JmlTelat = Math.Abs(CDbl((vDtpMP_RefDate.Value.Date - .order_point.Date).TotalDays))
                                JmlKompensasi = Math.Round((JmlTelat * .qty_avg_montly_demand) / JmlDaysPerMonth, .default_prec_unit, MidpointRounding.AwayFromZero)
                                If .qty_required + JmlKompensasi > .qty_optimum Then
                                    JmlKompensasi = .qty_optimum - .qty_required
                                    'Else
                                    '    MsgBox("")
                                End If
                            Else
                                JmlKompensasi = .qty_safety_stock
                            End If
                            .qty_required += JmlKompensasi
                            .qty_order = .qty_required

                            SSDate = vDtpMP_RefDate.Value.Date.AddDays(.lead_time_days)
                            .next_order_point_delay = vObjCD.RoundDown1((.qty_order * JmlDaysPerMonth) / .qty_avg_montly_demand, 0) - .lead_time_days
                            If .next_order_point_delay > (Date.MaxValue.Date - vDtpMP_RefDate.Value.Date).Days Then
                                .next_order_point = #1/1/2200#
                            Else
                                .next_order_point = SSDate + TimeSpan.FromDays(.next_order_point_delay)
                            End If
                        End If
                    Else
                        If FreeStock > .qty_safety_stock Then
                            .qty_required = 0
                        Else
                            .qty_required = .qty_safety_stock - FreeStock
                        End If

                        .qty_order = .qty_required
                        .qty_order_multiunit = .qty_order
                        If .qty_avg_montly_demand = 0 Then
                            .qty_avg_montly_demand = 1
                        End If

                        .order_point_delay = 9999
                        .order_point = #1/1/2200#
                        .next_order_point_delay = 9999
                        .next_order_point = #1/1/2200#
                    End If

                    Filter = "[Primary Product] = " + rowProduct.Primary.ToString
                    TBLPOItem.DefaultView.RowFilter = Filter
                    If TBLPOItem.DefaultView.Count = 0 Then
                        .Currency_Code = ""
                        .def_prec_currency = 0
                        .def_prec_order_unit = rowProduct._Decimal
                        .discount = 0
                        .disc_from_perct = 0
                        .disc_amount = 0
                        .nett_price = 0
                        .primary_currency = -1
                        .primary_prefered_supplier = rowProduct.primary_prefered_supplier
                        .sub_total = 0
                        .Supplier_Code = ""
                        .Supplier_Name = ""
                        .unit_price = 0

                        'belum ada rujukan po sebelumnya, qty order multi = qty order default
                        .qty_order_multiunit = .qty_order
                        .ratio_default = 1
                        .ratio_multi = 1
                        .primary_unit_multi = .primary_unit
                        .def_prec_unit_multi = .default_prec_unit
                        .Unit_Code_Multi = .Unit_Code
                        .Tolerance_Up = Math.Round(CDec((rowProduct.default_upper_tolerance * .qty_order_multiunit) / 100), CInt(.def_prec_unit_multi), MidpointRounding.AwayFromZero)
                        .Tolerance_Down = Math.Round(CDec((rowProduct.default_lower_tolerance * .qty_order_multiunit) / 100), CInt(.def_prec_unit_multi), MidpointRounding.AwayFromZero)
                    Else
                        SupplierSelectionMode = .supplier_selection_mode
                        If vModifySupplierSelectionMode = True Then
                            SupplierSelectionMode = vPreferedSupplierSelectionMode
                            .supplier_selection_mode = CInt(SupplierSelectionMode)
                        End If

                        If SupplierSelectionMode = clsMasterProduct.EnumSupplierSelectionMode.Prefered_Supplier Then
                            'cari data order terakhir untuk product tsb, samakan datanya
                            Filter = "[Primary Product] = " + rowProduct.Primary.ToString + " and [Primary Supplier] = " + rowProduct.primary_prefered_supplier.ToString
                            TBLPOItem.DefaultView.RowFilter = Filter

                            If TBLPOItem.DefaultView.Count = 0 Then
                                .Currency_Code = ""
                                .def_prec_currency = 0
                                .def_prec_order_unit = rowProduct._Decimal
                                .discount = 0
                                .disc_from_perct = 0
                                .disc_amount = 0
                                .nett_price = 0
                                .primary_currency = -1
                                .primary_prefered_supplier = -1
                                .sub_total = 0
                                .Supplier_Code = ""
                                .Supplier_Name = ""
                                .unit_price = 0

                                'belum ada rujukan po sebelumnya, qty order multi = qty order default
                                .qty_order_multiunit = .qty_order
                                .ratio_default = 1
                                .ratio_multi = 1
                                .primary_unit_multi = .primary_unit
                                .def_prec_unit_multi = .default_prec_unit
                                .Unit_Code_Multi = .Unit_Code
                                .Tolerance_Up = Math.Round(CDec((rowProduct.default_upper_tolerance * .qty_order_multiunit) / 100), CInt(.def_prec_unit_multi), MidpointRounding.AwayFromZero)
                                .Tolerance_Down = Math.Round(CDec((rowProduct.default_lower_tolerance * .qty_order_multiunit) / 100), CInt(.def_prec_unit_multi), MidpointRounding.AwayFromZero)
                            Else
                                rowPOItem = TBLPOItem.DefaultView(0).Row

                                'hitung qty multi unit
                                If rowPOItem.Primary_Unit_Multi = rowPOItem.Primary_Units Then
                                    .qty_order_multiunit = .qty_order
                                    .ratio_default = 1
                                    .ratio_multi = 1
                                    .primary_unit_multi = .primary_unit
                                    .def_prec_unit_multi = .default_prec_unit
                                    .Unit_Code_Multi = .Unit_Code
                                Else
                                    .ratio_default = rowPOItem.Ratio_Default
                                    .ratio_multi = rowPOItem.Ratio_Reference
                                    .primary_unit_multi = rowPOItem.Primary_Unit_Multi
                                    TBLMultiUnit.DefaultView.RowFilter = "[Primary Product] = " + rowPOItem.Primary_Product.ToString + _
                                    " AND [Primary Unit] = " + rowPOItem.Primary_Unit_Multi.ToString

                                    If TBLMultiUnit.DefaultView.Count = 0 Then
                                        .def_prec_unit_multi = .default_prec_unit
                                        .Unit_Code_Multi = .Unit_Code
                                    Else
                                        rowUnitMulti = TBLMultiUnit.DefaultView(0).Row
                                        .def_prec_unit_multi = rowUnitMulti.Default_Precision
                                        .Unit_Code_Multi = rowUnitMulti.Unit_Code
                                    End If
                                    .qty_order_multiunit = Math.Round(CDec((.qty_order * .ratio_multi) / .ratio_default), _
                                    CInt(.def_prec_unit_multi), MidpointRounding.AwayFromZero)
                                    'hitung kembali qty order berdasarkan qty multi unit
                                    .qty_order = Math.Round(.qty_order_multiunit * .ratio_default / _
                                    .ratio_multi, .default_prec_unit, MidpointRounding.AwayFromZero)
                                End If
                                .Tolerance_Up = Math.Round(CDec((rowProduct.default_upper_tolerance * .qty_order_multiunit) / 100), CInt(.def_prec_unit_multi), MidpointRounding.AwayFromZero)
                                .Tolerance_Down = Math.Round(CDec((rowProduct.default_lower_tolerance * .qty_order_multiunit) / 100), CInt(.def_prec_unit_multi), MidpointRounding.AwayFromZero)

                                .Currency_Code = rowPOItem.Currency_Code
                                .def_prec_currency = rowPOItem.Default_Precision
                                .def_prec_order_unit = rowProduct._Decimal
                                .discount = rowPOItem.Discount
                                .disc_from_perct = rowPOItem.Disc_From_Perct
                                .disc_amount = rowPOItem.Disc_Amount
                                .nett_price = rowPOItem.Nett_Price
                                .primary_currency = rowPOItem.Primary_Currency
                                .primary_prefered_supplier = rowPOItem.Primary_Supplier
                                .Supplier_Code = rowPOItem.Supplier_Code
                                .Supplier_Name = rowPOItem.Supplier_Name
                                .unit_price = rowPOItem.Unit_Price
                                .sub_total = vObjCD.DigitGrouping(.qty_order_multiunit * .nett_price, .def_prec_currency)
                            End If
                        ElseIf SupplierSelectionMode = clsMasterProduct.EnumSupplierSelectionMode.The_cheapest_in_the_last_3_order Then
                            'berdasarkan 3 order terakhir, cari data yang paling murah
                            rowPOItem = Me.GetTheCheapestRowOf3(TBLPOItem, rowProduct.Primary)
                            If IsNothing(rowItem) = False Then
                                'hitung qty multi unit
                                If rowProduct.Jml_Unit_Multi = 0 Then
                                    .qty_order_multiunit = .qty_order
                                    .ratio_default = 1
                                    .ratio_multi = 1
                                    .primary_unit_multi = .primary_unit
                                    .def_prec_unit_multi = .default_prec_unit
                                    .Unit_Code_Multi = .Unit_Code
                                Else
                                    .ratio_default = rowPOItem.Ratio_Default
                                    .ratio_multi = rowPOItem.Ratio_Reference
                                    .primary_unit_multi = rowPOItem.Primary_Unit_Multi
                                    TBLMultiUnit.DefaultView.RowFilter = "[Primary Product] = " + rowPOItem.Primary_Product.ToString + _
                                    " AND [Primary Unit] = " + rowPOItem.Primary_Unit_Multi.ToString

                                    If TBLMultiUnit.DefaultView.Count = 0 Then
                                        .def_prec_unit_multi = .default_prec_unit
                                        .Unit_Code_Multi = .Unit_Code
                                    Else
                                        rowUnitMulti = TBLMultiUnit.DefaultView(0).Row
                                        .def_prec_unit_multi = rowUnitMulti.Default_Precision
                                        .Unit_Code_Multi = rowUnitMulti.Unit_Code
                                    End If
                                    .qty_order_multiunit = Math.Round(CDec((.qty_order * .ratio_multi) / .ratio_default), _
                                    CInt(.def_prec_unit_multi), MidpointRounding.AwayFromZero)
                                End If
                                .Tolerance_Up = Math.Round(CDec((rowProduct.default_upper_tolerance * .qty_order_multiunit) / 100), CInt(.def_prec_unit_multi), MidpointRounding.AwayFromZero)
                                .Tolerance_Down = Math.Round(CDec((rowProduct.default_lower_tolerance * .qty_order_multiunit) / 100), CInt(.def_prec_unit_multi), MidpointRounding.AwayFromZero)

                                .Currency_Code = rowPOItem.Currency_Code
                                .def_prec_currency = rowPOItem.Default_Precision
                                .def_prec_order_unit = rowProduct._Decimal
                                .discount = rowPOItem.Discount
                                .disc_from_perct = rowPOItem.Disc_From_Perct
                                .disc_amount = rowPOItem.Disc_Amount
                                .nett_price = rowPOItem.Nett_Price
                                .primary_currency = rowPOItem.Primary_Currency
                                .primary_prefered_supplier = rowPOItem.Primary_Supplier
                                .Supplier_Code = rowPOItem.Supplier_Code
                                .Supplier_Name = rowPOItem.Supplier_Name
                                .unit_price = rowPOItem.Unit_Price
                                .sub_total = vObjCD.DigitGrouping(.qty_order_multiunit * .nett_price, .def_prec_currency)
                            Else
                                .Currency_Code = ""
                                .def_prec_currency = 0
                                .def_prec_order_unit = rowProduct._Decimal
                                .discount = 0
                                .disc_from_perct = 0
                                .disc_amount = 0
                                .nett_price = 0
                                .primary_currency = -1
                                .primary_prefered_supplier = -1
                                .sub_total = 0
                                .Supplier_Code = ""
                                .Supplier_Name = ""
                                .unit_price = 0

                                'belum ada rujukan po sebelumnya, qty order multi = qty order default
                                .qty_order_multiunit = .qty_order
                                .ratio_default = 1
                                .ratio_multi = 1
                                .primary_unit_multi = .primary_unit
                                .def_prec_unit_multi = .default_prec_unit
                                .Unit_Code_Multi = .Unit_Code
                                .Tolerance_Up = Math.Round(CDec((rowProduct.default_upper_tolerance * .qty_order_multiunit) / 100), CInt(.def_prec_unit_multi), MidpointRounding.AwayFromZero)
                                .Tolerance_Down = Math.Round(CDec((rowProduct.default_lower_tolerance * .qty_order_multiunit) / 100), CInt(.def_prec_unit_multi), MidpointRounding.AwayFromZero)
                            End If
                        ElseIf SupplierSelectionMode = clsMasterProduct.EnumSupplierSelectionMode.Most_Recent_PO Then
                            rowPOItem = Me.GetTheRecentRow(TBLPOItem, rowProduct.Primary)
                            If IsNothing(rowItem) = False Then
                                'hitung qty multi unit
                                If rowProduct.Jml_Unit_Multi = 0 Then
                                    .qty_order_multiunit = .qty_order
                                    .ratio_default = 1
                                    .ratio_multi = 1
                                    .primary_unit_multi = .primary_unit
                                    .def_prec_unit_multi = .default_prec_unit
                                    .Unit_Code_Multi = .Unit_Code
                                Else
                                    .ratio_default = rowPOItem.Ratio_Default
                                    .ratio_multi = rowPOItem.Ratio_Reference
                                    .primary_unit_multi = rowPOItem.Primary_Unit_Multi
                                    TBLMultiUnit.DefaultView.RowFilter = "[Primary Product] = " + rowPOItem.Primary_Product.ToString + _
                                    " AND [Primary Unit] = " + rowPOItem.Primary_Unit_Multi.ToString

                                    If TBLMultiUnit.DefaultView.Count = 0 Then
                                        .def_prec_unit_multi = .default_prec_unit
                                        .Unit_Code_Multi = .Unit_Code
                                    Else
                                        rowUnitMulti = TBLMultiUnit.DefaultView(0).Row
                                        .def_prec_unit_multi = rowUnitMulti.Default_Precision
                                        .Unit_Code_Multi = rowUnitMulti.Unit_Code
                                    End If
                                    .qty_order_multiunit = Math.Round(CDec((.qty_order * .ratio_multi) / .ratio_default), _
                                    CInt(.def_prec_unit_multi), MidpointRounding.AwayFromZero)
                                End If
                                .Tolerance_Up = Math.Round(CDec((rowProduct.default_upper_tolerance * .qty_order_multiunit) / 100), CInt(.def_prec_unit_multi), MidpointRounding.AwayFromZero)
                                .Tolerance_Down = Math.Round(CDec((rowProduct.default_lower_tolerance * .qty_order_multiunit) / 100), CInt(.def_prec_unit_multi), MidpointRounding.AwayFromZero)

                                .Currency_Code = rowPOItem.Currency_Code
                                .def_prec_currency = rowPOItem.Default_Precision
                                .def_prec_order_unit = rowProduct._Decimal
                                .discount = rowPOItem.Discount
                                .disc_from_perct = rowPOItem.Disc_From_Perct
                                .disc_amount = rowPOItem.Disc_Amount
                                .nett_price = rowPOItem.Nett_Price
                                .primary_currency = rowPOItem.Primary_Currency
                                .primary_prefered_supplier = rowPOItem.Primary_Supplier
                                .Supplier_Code = rowPOItem.Supplier_Code
                                .Supplier_Name = rowPOItem.Supplier_Name
                                .unit_price = rowPOItem.Unit_Price
                                .sub_total = vObjCD.DigitGrouping(.qty_order_multiunit * .nett_price, .def_prec_currency)
                            Else
                                .Currency_Code = ""
                                .def_prec_currency = 0
                                .def_prec_order_unit = rowProduct._Decimal
                                .discount = 0
                                .disc_from_perct = 0
                                .disc_amount = 0
                                .nett_price = 0
                                .primary_currency = -1
                                .primary_prefered_supplier = -1
                                .sub_total = 0
                                .Supplier_Code = ""
                                .Supplier_Name = ""
                                .unit_price = 0

                                'belum ada rujukan po sebelumnya, qty order multi = qty order default
                                .qty_order_multiunit = .qty_order
                                .ratio_default = 1
                                .ratio_multi = 1
                                .primary_unit_multi = .primary_unit
                                .def_prec_unit_multi = .default_prec_unit
                                .Unit_Code_Multi = .Unit_Code
                                .Tolerance_Up = Math.Round(CDec((rowProduct.default_upper_tolerance * .qty_order_multiunit) / 100), CInt(.def_prec_unit_multi), MidpointRounding.AwayFromZero)
                                .Tolerance_Down = Math.Round(CDec((rowProduct.default_lower_tolerance * .qty_order_multiunit) / 100), CInt(.def_prec_unit_multi), MidpointRounding.AwayFromZero)
                            End If
                        End If
                    End If
                    Filter = "[Primary Product] = " + rowProduct.Primary.ToString
                    If TBLPOItem.DefaultView.RowFilter <> Filter Then
                        TBLPOItem.DefaultView.RowFilter = Filter
                    End If
                    While TBLPOItem.DefaultView.Count > 0
                        TBLPOItem.Rows.Remove(TBLPOItem.DefaultView(0).Row)
                    End While
                End With
                vDT.Rows.Add(rowItem)
            Next

            If AdaDetailDelivery = True Then
                Me.IsikanDetailDeliverySOB(PrimarySOMin, PrimarySOMax, vDT_ItemSOB, vDT_ItemSOBDelivery)
            End If

            If AdaDetailDeliveryPrOB = True Then
                Me.IsikanDetailDeliveryPrOB(PrimaryPrOBMin, PrimaryPrOBMax, vDT_ItemPrOB, vDT_ItemPrOBDelivery)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAProduct.Dispose()
            TBLProduct.Dispose()
            DTAGIT.Dispose()
            TBLGIT.Dispose()
            DTAProductQty.Dispose()
            TBLProductQty.Dispose()
            DTAPendingDO.Dispose()
            TBLPendingDO.Dispose()
            DTAPOB.Dispose()
            TBLPOB.Dispose()
            DTASOB.Dispose()
            TBLSOB.Dispose()
            DTAPOItem.Dispose()
            TBLPOItem.Dispose()
            DTAPrOB.Dispose()
            TBLPrOB.Dispose()
            DTAMultiUnit.Dispose()
            TBLMultiUnit.Dispose()
            ijkl = Nothing 
        End Try
    End Sub

    Private Sub IsikanDetailDeliveryPrOB(ByVal PrimaryPrOBMin As Double, ByVal PrimaryPrOBMax As Double, _
    ByVal vDT_ItemPrOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_reqDataTable, _
    ByVal vDT_ItemPrOBDelivery As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_req_deliveryDataTable)
        Dim DTAPrOBDeliveryDummy As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_purchase_req_itemdeliveryTableAdapter
        Dim TBLPrOBDeliveryDummy As New DseTransaksiMaterialPlanning2.tbltrans_purchase_req_itemdeliveryDataTable
        Dim RowPrOBDeliveryDummy As DseTransaksiMaterialPlanning2.tbltrans_purchase_req_itemdeliveryRow
        Dim rowPrOB2 As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_reqRow
        Dim JmlAllocated As Decimal = 0
        Dim IdxDelivery As Integer = 0
        Dim rowPrOBDelivery As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_req_deliveryRow

        Try
            vDT_ItemPrOBDelivery.Rows.Clear()
            DTAPrOBDeliveryDummy.Connection.ConnectionString = Core_Module.GetConstring

            'isikan database
            DTAPrOBDeliveryDummy.Fill(TBLPrOBDeliveryDummy, PrimaryPrOBMax, PrimaryPrOBMin)
            TBLPrOBDeliveryDummy.DefaultView.Sort = "[delivery date] DESC"

            'isikan item delivery untuk item sob tsb
            For Each rowPrOB2 In vDT_ItemPrOB.Rows
                If rowPrOB2.delivery_date_cont = 1 Then
                    JmlAllocated = 0
                    IdxDelivery = 0

                    If rowPrOB2.delivery_date_cont = 1 Then
                        'pindahkan item delivery detail
                        TBLPrOBDeliveryDummy.DefaultView.RowFilter = "[identifier] = '" + rowPrOB2.identifier_pr_item + "'"
                        While JmlAllocated < rowPrOB2.qty And IdxDelivery <= TBLPrOBDeliveryDummy.DefaultView.Count - 1
                            RowPrOBDeliveryDummy = TBLPrOBDeliveryDummy.DefaultView(IdxDelivery).Row
                            If RowPrOBDeliveryDummy.delivery_qty <= 0 Then
                                IdxDelivery += 1
                            Else
                                If rowPrOB2.qty - JmlAllocated >= RowPrOBDeliveryDummy.delivery_qty Then
                                    'ambil secara keseluruhan karena qty sob - qty allocated > qty delivery
                                    rowPrOBDelivery = vDT_ItemPrOBDelivery.Newtbltrans_material_planning_itempurch_req_deliveryRow
                                    With rowPrOBDelivery
                                        .delivery_date = RowPrOBDeliveryDummy.delivery_date
                                        .delivery_priority = RowPrOBDeliveryDummy.delivery_priority
                                        .delivery_qty = RowPrOBDeliveryDummy.delivery_qty
                                        .identifier = RowPrOBDeliveryDummy.identifier
                                        .primary = -1
                                        .primary_purc_req = RowPrOBDeliveryDummy.primary_transaksi
                                        .primary_transaksi = -1
                                        .remarks = ""
                                    End With
                                    vDT_ItemPrOBDelivery.Rows.Add(rowPrOBDelivery)

                                    JmlAllocated += RowPrOBDeliveryDummy.delivery_qty
                                    RowPrOBDeliveryDummy.delivery_qty = 0
                                Else
                                    'ambil sebagian karena qty sob - qty allocated <= qty delivery
                                    rowPrOBDelivery = vDT_ItemPrOBDelivery.Newtbltrans_material_planning_itempurch_req_deliveryRow
                                    With rowPrOBDelivery
                                        .delivery_date = RowPrOBDeliveryDummy.delivery_date
                                        .delivery_priority = RowPrOBDeliveryDummy.delivery_priority
                                        .delivery_qty = rowPrOB2.qty - JmlAllocated
                                        .identifier = RowPrOBDeliveryDummy.identifier
                                        .primary = -1
                                        .primary_purc_req = RowPrOBDeliveryDummy.primary_transaksi
                                        .primary_transaksi = -1
                                        .remarks = ""
                                    End With
                                    vDT_ItemPrOBDelivery.Rows.Add(rowPrOBDelivery)

                                    RowPrOBDeliveryDummy.delivery_qty -= (rowPrOB2.qty - JmlAllocated)
                                    JmlAllocated = rowPrOB2.qty
                                End If

                                'tentukan apakah harus pindah ke baris selanjutnya pada item delivery
                                If RowPrOBDeliveryDummy.delivery_qty <= 0 Then
                                    IdxDelivery += 1
                                End If
                            End If
                        End While
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub IsikanDetailDeliverySOB(ByVal PrimarySOMin As Double, ByVal PrimarySOMax As Double, _
    ByVal vDT_ItemSOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsobDataTable, _
    ByVal vDT_ItemSOBDelivery As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsob_deliveryDataTable)
        Dim DTASOBDeliveryDummy As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_so_itemdeliveryTableAdapter
        Dim TBLSOBDeliveryDummy As New DseTransaksiMaterialPlanning2.tbltrans_so_itemdeliveryDataTable
        Dim RowSOBDeliveryDummy As DseTransaksiMaterialPlanning2.tbltrans_so_itemdeliveryRow
        Dim rowSOB2 As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsobRow
        Dim JmlAllocated As Decimal = 0
        Dim IdxDelivery As Integer = 0
        Dim rowSOBDelivery As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsob_deliveryRow
        Dim DTADefPrecMulti As New DseTransaksiMaterialPlanning2TableAdapters.tblmaster_product_item_multiunit1TableAdapter
        Dim TBLDefPrecMulti As New DseTransaksiMaterialPlanning2.tblmaster_product_item_multiunit1DataTable
        Dim TBLDefPrecMulti1 As New DseTransaksiMaterialPlanning2.tblmaster_product_item_multiunit1DataTable
        Dim rowDefPrecMulti, rowDefPrecMulti1 As DseTransaksiMaterialPlanning2.tblmaster_product_item_multiunit1Row
        Dim DefPrecMulti As Integer = 0

        Try
            DTASOBDeliveryDummy.Connection.ConnectionString = Core_Module.GetConstring
            DTADefPrecMulti.Connection.ConnectionString = Core_Module.GetConstring

            'isikan database
            DTASOBDeliveryDummy.Fill(TBLSOBDeliveryDummy, PrimarySOMin, PrimarySOMax)
            TBLSOBDeliveryDummy.DefaultView.Sort = "[delivery date] DESC"

            'isikan item delivery untuk item sob tsb
            For Each rowSOB2 In vDT_ItemSOB.Rows
                If rowSOB2.delivery_date_cont = 1 Then
                    JmlAllocated = 0
                    IdxDelivery = 0
                    If rowSOB2.primary_unit_multi <> rowSOB2.primary_unit Then
                        If TBLDefPrecMulti1.Rows.Count = 0 Then
                            DTADefPrecMulti.Fill(TBLDefPrecMulti1)
                        End If
                    End If

                    If rowSOB2.delivery_date_cont = 1 Then
                        'pindahkan item delivery detail
                        TBLSOBDeliveryDummy.DefaultView.RowFilter = "[identifier] = '" + rowSOB2.identifier_so + "'"
                        While JmlAllocated < rowSOB2.qty And IdxDelivery <= TBLSOBDeliveryDummy.DefaultView.Count - 1
                            RowSOBDeliveryDummy = TBLSOBDeliveryDummy.DefaultView(IdxDelivery).Row
                            If RowSOBDeliveryDummy.delivery_qty <= 0 Then
                                IdxDelivery += 1
                            Else
                                If rowSOB2.qty - JmlAllocated >= RowSOBDeliveryDummy.delivery_qty Then
                                    'ambil secara keseluruhan karena qty sob - qty allocated > qty delivery
                                    rowSOBDelivery = vDT_ItemSOBDelivery.Newtbltrans_material_planning_itemsob_deliveryRow
                                    With rowSOBDelivery
                                        .delivery_date = RowSOBDeliveryDummy.delivery_date
                                        .delivery_priority = RowSOBDeliveryDummy.delivery_priority
                                        .delivery_qty = RowSOBDeliveryDummy.delivery_qty
                                        .delivery_qty_multi = RowSOBDeliveryDummy.delivery_qty_multi
                                        .identifier = RowSOBDeliveryDummy.identifier
                                        .primary = -1
                                        .primary_so = RowSOBDeliveryDummy.primary_transaksi
                                        .primary_transaksi = -1
                                        .remarks = ""
                                    End With
                                    vDT_ItemSOBDelivery.Rows.Add(rowSOBDelivery)

                                    JmlAllocated += RowSOBDeliveryDummy.delivery_qty
                                    RowSOBDeliveryDummy.delivery_qty = 0
                                Else
                                    'ambil sebagian karena qty sob - qty allocated <= qty delivery
                                    rowSOBDelivery = vDT_ItemSOBDelivery.Newtbltrans_material_planning_itemsob_deliveryRow
                                    With rowSOBDelivery
                                        .delivery_date = RowSOBDeliveryDummy.delivery_date
                                        .delivery_priority = RowSOBDeliveryDummy.delivery_priority
                                        .delivery_qty = rowSOB2.qty - JmlAllocated
                                        If rowSOB2.primary_unit_multi = rowSOB2.primary_unit Then
                                            'jika bukan multi unit, qty multi = qty default
                                            .delivery_qty_multi = .delivery_qty
                                        Else
                                            TBLDefPrecMulti.DefaultView.RowFilter = "[Primary Product] = " + rowSOB2.primary_product.ToString + " AND [Primary Unit] = " + rowSOB2.primary_unit_multi.ToString
                                            If TBLDefPrecMulti.DefaultView.Count > 0 Then
                                                rowDefPrecMulti = TBLDefPrecMulti.DefaultView(0).Row
                                                DefPrecMulti = rowDefPrecMulti.Default_Precision
                                            Else
                                                TBLDefPrecMulti1.DefaultView.RowFilter = "[Primary Product] = " + rowSOB2.primary_product.ToString + " AND [Primary Unit] = " + rowSOB2.primary_unit_multi.ToString
                                                If TBLDefPrecMulti1.DefaultView.Count = 0 Then
                                                    DefPrecMulti = 2
                                                Else
                                                    rowDefPrecMulti1 = TBLDefPrecMulti1.DefaultView(0).Row
                                                    DefPrecMulti = rowDefPrecMulti1.Default_Precision
                                                    rowDefPrecMulti = TBLDefPrecMulti.NewRow
                                                    For ij As Integer = 0 To TBLDefPrecMulti.Columns.Count - 1
                                                        rowDefPrecMulti.Item(ij) = rowDefPrecMulti1.Item(ij)
                                                    Next
                                                    TBLDefPrecMulti.Rows.Add(rowDefPrecMulti)
                                                End If
                                            End If

                                            .delivery_qty_multi = Math.Round((.delivery_qty * rowSOB2.ratio_reference) / rowSOB2.ratio_default, DefPrecMulti, MidpointRounding.AwayFromZero)
                                        End If
                                        .identifier = RowSOBDeliveryDummy.identifier
                                        .primary = -1
                                        .primary_so = RowSOBDeliveryDummy.primary_transaksi
                                        .primary_transaksi = -1
                                        .remarks = ""
                                    End With
                                    vDT_ItemSOBDelivery.Rows.Add(rowSOBDelivery)

                                    RowSOBDeliveryDummy.delivery_qty -= (rowSOB2.qty - JmlAllocated)
                                    JmlAllocated = rowSOB2.qty
                                End If

                                'tentukan apakah harus pindah ke baris selanjutnya pada item delivery
                                If RowSOBDeliveryDummy.delivery_qty <= 0 Then
                                    IdxDelivery += 1
                                End If
                            End If
                        End While
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub HitungMaterialPlanning(ByVal vDT As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item1DataTable, _
   ByVal vDT_ItemGIT As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemgitDataTable, _
   ByVal vDT_ItemDO As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemdoDataTable, _
   ByVal vDT_ItemPOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempobDataTable, _
   ByVal vDT_ItemSOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsobDataTable, _
   ByVal vDT_ItemPrOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_reqDataTable, _
   ByVal vDT_ItemSettingLocation As DseTransaksiMaterialPlanning_WarehouseExc.tbltrans_material_planning_item_location_exclusionDataTable, _
   ByVal vDT_ItemQty As DseTransaksiMaterialPlanning_WarehouseExc.tbltrans_material_planning_item_qtyDataTable, _
   ByVal vDT_ItemMOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemmobDataTable, _
   ByVal vDT_ItemRMReq As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_rm_reqDataTable, _
   ByVal vObjSetting As Core_Module.ClsTransaksiSetting, ByRef vTxtMP_DaysPerMonth As Windows.Forms.TextBox, _
   ByVal vBackColorAll As Color, ByVal vFrmCek As FrmTrans_MaterialPlanningDashboard_DetailPO, ByVal vUserCode As String, _
   ByVal vObjCD As Core_Module.ClsCekData, ByVal vTxtMP_Primary As Windows.Forms.TextBox, _
   ByVal vObjSettingUser As Core_Module.ClsTransaksiSettingUser, ByVal vDtpMP_RefDate As Windows.Forms.DateTimePicker, ByVal vPrimaryProduct As Double)
        Dim rowItem As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item1Row
        Dim rowItemRMReq As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_rm_reqRow
        Dim rowItemSOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsobRow
        Dim rowItemPrOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_reqRow
        Dim JmlDummy As Decimal = 0
        Dim JmlDummy1 As Decimal = 0
        Dim Filter As String = ""
        Dim DV As Data.DataRowView
        Dim OrderStrategy As Core_Module.clsMasterProduct.EnumOrderStrategy = Core_Module.clsMasterProduct.EnumOrderStrategy.Order_By_Maximum_Order_Qty
        Dim JmlDaysPerMonth As Integer = 0
        'Dim rowDGV As DataGridViewRow
        Dim FreeStock As Decimal = 0
        Dim IndexDummy As Integer = 0
        Dim ProcurementStrategy As Core_Module.clsMasterProduct.EnumProcurementStrategy = Core_Module.clsMasterProduct.EnumProcurementStrategy.Build_To_Stock
        Dim JmlTelat As Double = 0
        Dim JmlKompensasi As Decimal = 0
        Dim SSDate As Date = Today

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If vObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.MaterialPlanning, _
                ClsTransaksiSetting.EnumNamaParameter.MaterialPlanning_JmlHariKerjaPerBulan) = "" Then
                MsgBox("Please specify days quantity per month for material planning calculation", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If
            JmlDaysPerMonth = vObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.MaterialPlanning, _
            ClsTransaksiSetting.EnumNamaParameter.MaterialPlanning_JmlHariKerjaPerBulan)
            vTxtMP_DaysPerMonth.Text = JmlDaysPerMonth

            For Each rowItem In vDT.Rows
                If rowItem.primary_code_no = vPrimaryProduct Then
                    With rowItem
                        .qty_rm_req = 0
                        .MP_Spare_Qty_Rm_Req = 0
                        Filter = "[primary product] = " + .primary_code_no.ToString
                        vDT_ItemRMReq.DefaultView.RowFilter = Filter
                        For Each DV In vDT_ItemRMReq.DefaultView
                            rowItemRMReq = DV.Row
                            .qty_rm_req += rowItemRMReq.qty
                            .MP_Spare_Qty_Rm_Req += rowItemRMReq.MP_Spare_Qty
                        Next

                        .qty_sob = 0
                        .MP_Spare_Qty_SOB = 0
                        Filter = "[primary product] = " + .primary_code_no.ToString
                        vDT_ItemSOB.DefaultView.RowFilter = Filter
                        For Each DV In vDT_ItemSOB.DefaultView
                            rowItemSOB = DV.Row
                            .qty_sob += rowItemSOB.qty
                            .MP_Spare_Qty_SOB += rowItemSOB.MP_Spare_Qty
                        Next

                        .qty_prob = 0
                        .MP_Spare_Qty_PrOB = 0
                        Filter = "[primary product] = " + .primary_code_no.ToString
                        vDT_ItemPrOB.DefaultView.RowFilter = Filter
                        For Each DV In vDT_ItemPrOB.DefaultView
                            rowItemPrOB = DV.Row
                            .qty_prob += rowItemPrOB.qty
                            .MP_Spare_Qty_PrOB += rowItemPrOB.MP_Spare_Qty
                        Next

                        .qty_inquiry = 0

                        FreeStock = .qty_on_hand + .qty_pob + .qty_mob + .qty_git - .qty_inquiry - .qty_sob - .qty_prob - .qty_pending_do - .qty_rm_req - _
                        .MP_Spare_Qty_PrOB - .MP_Spare_Qty_Rm_Req - .MP_Spare_Qty_SOB
                        ProcurementStrategy = .procurement_strategy

                        If ProcurementStrategy = clsMasterProduct.EnumProcurementStrategy.Build_To_Stock Then
                            If OrderStrategy = clsMasterProduct.EnumOrderStrategy.Order_By_Maximum_Order_Qty Or _
                                              OrderStrategy = clsMasterProduct.EnumOrderStrategy.Order_By_Minimum_Order_Qty Or _
                                              OrderStrategy = clsMasterProduct.EnumOrderStrategy.Order_By_Optimum_Order_Qty Then
                                .qty_required = .qty_order_optimum
                            ElseIf OrderStrategy = clsMasterProduct.EnumOrderStrategy.Order_To_Max_Qty Or _
                            OrderStrategy = clsMasterProduct.EnumOrderStrategy.Order_To_Min_Qty Or _
                            OrderStrategy = clsMasterProduct.EnumOrderStrategy.Order_To_Optimum_Qty Or _
                            OrderStrategy = clsMasterProduct.EnumOrderStrategy.Order_To_Safety_Stock Then
                                If FreeStock > .qty_safety_stock Then
                                    .qty_required = .qty_optimum - .qty_safety_stock
                                Else
                                    .qty_required = .qty_optimum - FreeStock
                                End If
                            End If

                            .qty_order = .qty_required

                            If .qty_avg_montly_demand = 0 Then
                                .qty_avg_montly_demand = 1
                            End If

                            Me.HitungOrderPoint(rowItem, JmlDaysPerMonth, vDtpMP_RefDate, vObjCD)

                            If .order_point.Date < vDtpMP_RefDate.Value.Date Then
                                If FreeStock > 0 Then
                                    JmlTelat = Math.Abs(CDbl((vDtpMP_RefDate.Value.Date - .order_point.Date).TotalDays))
                                    JmlKompensasi = Math.Round((JmlTelat * .qty_avg_montly_demand) / JmlDaysPerMonth, .default_prec_unit, MidpointRounding.AwayFromZero)
                                    If .qty_required + JmlKompensasi > .qty_optimum Then
                                        JmlKompensasi = .qty_optimum - .qty_required
                                        'Else
                                        '    MsgBox("")
                                    End If
                                Else
                                    JmlKompensasi = .qty_safety_stock
                                End If
                                .qty_required += JmlKompensasi
                                .qty_order = .qty_required

                                SSDate = vDtpMP_RefDate.Value.Date.AddDays(.lead_time_days)
                                .next_order_point_delay = vObjCD.RoundDown1((.qty_order * JmlDaysPerMonth) / .qty_avg_montly_demand, 0) - .lead_time_days
                                If .next_order_point_delay > (Date.MaxValue.Date - vDtpMP_RefDate.Value.Date).Days Then
                                    .next_order_point = #1/1/2200#
                                Else
                                    .next_order_point = SSDate + TimeSpan.FromDays(.next_order_point_delay)
                                End If
                            End If
                        Else
                            If FreeStock > .qty_safety_stock Then
                                .qty_required = 0
                            Else
                                .qty_required = .qty_safety_stock - FreeStock
                            End If

                            .qty_order = .qty_required
                            .qty_order_multiunit = .qty_order
                            If .qty_avg_montly_demand = 0 Then
                                .qty_avg_montly_demand = 1
                            End If

                            .order_point_delay = 9999
                            .order_point = #1/1/2200#
                            .next_order_point_delay = 9999
                            .next_order_point = #1/1/2200#
                        End If
                    End With
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub HitungQtyAvailable(ByVal vDT As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item2DataTable, _
    ByVal vDT_ItemGIT As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemgitDataTable, _
    ByVal vDT_ItemDO As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemdoDataTable, _
    ByVal vDT_ItemPOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempobDataTable, _
    ByVal vDT_ItemSOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsobDataTable, _
    ByVal vDT_ItemPrOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_reqDataTable, _
    ByVal vDT_ItemSettingLocation As DseTransaksiMaterialPlanning_WarehouseExc.tbltrans_material_planning_item_location_exclusionDataTable, _
    ByVal vDT_ItemQty As DseTransaksiMaterialPlanning_WarehouseExc.tbltrans_material_planning_item_qtyDataTable, _
    ByVal vDT_ItemMOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemmobDataTable, _
    ByVal vDT_ItemRMReq As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_rm_reqDataTable, _
    ByVal vObjSetting As Core_Module.ClsTransaksiSetting, ByVal vBackColorAll As Color, ByVal vUserCode As String, _
    ByVal vObjCD As Core_Module.ClsCekData, ByVal vObjSettingUser As Core_Module.ClsTransaksiSettingUser, ByVal vListPrimaryProduct As List(Of Double), _
    ByVal vTblItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable, ByVal vIdentifierJO As String, ByVal vJONO As String, _
    ByVal vMONo As String, ByVal vMODate As Date, ByVal vPrimaryMO As Double)
        Dim vFrmCek As New FrmTrans_MaterialPlanningDashboard_DetailPO
        Dim DTAProduct As New DseTransaksiMaterialPlanning2TableAdapters.DataTable1TableAdapter
        Dim TBLProduct As New DseTransaksiMaterialPlanning2.DataTable1DataTable
        Dim rowProduct As DseTransaksiMaterialPlanning2.DataTable1Row
        Dim rowItem As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item2Row
        Dim DTAGIT As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_packinglist_itemTableAdapter
        Dim TBLGIT As New DseTransaksiMaterialPlanning2.tbltrans_packinglist_itemDataTable
        Dim rowGIT As DseTransaksiMaterialPlanning2.tbltrans_packinglist_itemRow
        Dim rowGIT2 As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemgitRow
        Dim JmlDummy As Decimal = 0
        Dim Filter As String = ""
        Dim DV As Data.DataRowView
        Dim DTAProductQty As New DseTransaksiMaterialPlanning2TableAdapters.tblmaster_productqtyTableAdapter
        Dim TBLProductQty As New DseTransaksiMaterialPlanning2.tblmaster_productqtyDataTable
        Dim rowProductQty As DseTransaksiMaterialPlanning2.tblmaster_productqtyRow
        Dim DTAPendingDO As New DseTransaksiMaterialPlanning2TableAdapters.DataTable2TableAdapter
        Dim TBLPendingDO As New DseTransaksiMaterialPlanning2.DataTable2DataTable
        Dim rowPendingDO As DseTransaksiMaterialPlanning2.DataTable2Row
        Dim rowPendingDO2 As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemdoRow
        Dim DTAPOB As New DseTransaksiMaterialPlanning2TableAdapters.DataTable3TableAdapter
        Dim TBLPOB As New DseTransaksiMaterialPlanning2.DataTable3DataTable
        Dim rowPOB As DseTransaksiMaterialPlanning2.DataTable3Row
        Dim rowPOB2 As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempobRow
        Dim CekOtorisasiPO As Boolean = True
        Dim DTASOB As New DseTransaksiMaterialPlanning2TableAdapters.DataTable4TableAdapter
        Dim TBLSOB As New DseTransaksiMaterialPlanning2.DataTable4DataTable
        Dim rowSOB As DseTransaksiMaterialPlanning2.DataTable4Row
        Dim rowSOB2 As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsobRow
        Dim rowDGV As DataGridViewRow
        Dim FreeStock As Decimal = 0
        Dim IndexDummy As Integer = 0
        Dim DTAPrOB As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planning_itempurch_req1TableAdapter
        Dim rowPrOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_req1Row
        Dim rowPrOB2 As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_reqRow
        Dim TBLPrOB As New DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_req1DataTable
        Dim rowProductQty1 As DseTransaksiMaterialPlanning_WarehouseExc.tbltrans_material_planning_item_qtyRow
        Dim DTAMOB As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_mo_itemTableAdapter
        Dim TBLMOB As New DseTransaksiMaterialPlanning2.tbltrans_mo_itemDataTable
        Dim rowMOB As DseTransaksiMaterialPlanning2.tbltrans_mo_itemRow
        Dim rowMOB2 As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemmobRow
        Dim DTARMReq As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_mo_item_rmTableAdapter
        Dim TBLRMReq As New DseTransaksiMaterialPlanning2.tbltrans_mo_item_rmDataTable
        Dim rowRMReq As DseTransaksiMaterialPlanning2.tbltrans_mo_item_rmRow
        Dim rowRMReq2 As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_rm_reqRow
        Dim DTAWIP As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_mo_item_wipTableAdapter
        Dim TBLWIP As New DseTransaksiMaterialPlanning2.tbltrans_mo_item_wipDataTable
        Dim rowWIP As DseTransaksiMaterialPlanning2.tbltrans_mo_item_wipRow
        Dim rowItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woRow

        Try
            'product exclusion     
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DTAProduct.Connection.ConnectionString = Core_Module.GetConstring
            DTAGIT.Connection.ConnectionString = DTAProduct.Connection.ConnectionString
            DTAProductQty.Connection.ConnectionString = DTAProduct.Connection.ConnectionString
            DTAPendingDO.Connection.ConnectionString = DTAProduct.Connection.ConnectionString
            DTAPOB.Connection.ConnectionString = DTAProduct.Connection.ConnectionString
            DTASOB.Connection.ConnectionString = DTAProduct.Connection.ConnectionString
            DTAPrOB.Connection.ConnectionString = DTAProduct.Connection.ConnectionString
            DTAMOB.Connection.ConnectionString = DTAProduct.Connection.ConnectionString
            DTARMReq.Connection.ConnectionString = DTAProduct.Connection.ConnectionString
            DTAWIP.Connection.ConnectionString = DTAProduct.Connection.ConnectionString

            vDT.Rows.Clear()
            vDT_ItemGIT.Rows.Clear()
            vDT_ItemDO.Rows.Clear()
            vDT_ItemPOB.Rows.Clear()
            vDT_ItemSOB.Rows.Clear()
            vDT_ItemPrOB.Rows.Clear()
            vDT_ItemQty.Rows.Clear()
            vDT_ItemSettingLocation.Rows.Clear()
            vDT_ItemRMReq.Rows.Clear()
            vDT_ItemMOB.Rows.Clear()
            vDT_ItemRMReq.Rows.Clear()

            'isikan semua data dasar
            'DTAProduct.Fill(TBLProduct, ProductType)-> di ubah ketika memasukan perhitungan purchase request, jasa skrg diperhitungkan
            DTAProduct.FillBy_all(TBLProduct)

            DTAGIT.Fill(TBLGIT)
            DTAProductQty.Fill(TBLProductQty)
            DTAPendingDO.Fill(TBLPendingDO)
            DTAPOB.Fill(TBLPOB)
            DTASOB.Fill(TBLSOB)
            DTAPrOB.Fill(TBLPrOB)
            DTAMOB.Fill(TBLMOB)
            DTARMReq.Fill(TBLRMReq)

            'Fill item wip
            DTAWIP.Fill(TBLWIP)
            vTblItemWO.DefaultView.RowFilter = "[Identifier JO] = '" + vIdentifierJO + "'"
            'timpa qty wip untuk alokasi wip sesuai data transaksi terakhir
            '1. cari apakah sudah ada di dalam item rm req, jika sudah ada maka timpa qty-nya
            '2. jika tidak ada maka tambahkan baris baru
            For Each DV1 As Data.DataRowView In vTblItemWO.DefaultView
                rowItemWO = DV1.Row
                If rowItemWO.Qty_WIP > 0 Then
                    TBLWIP.DefaultView.RowFilter = "[Identifier JO] = '" + rowItemWO.Identifier_JO + "' AND [Identifier WO] = '" + rowItemWO.Identifier + "'"
                    If TBLWIP.DefaultView.Count = 0 Then
                        'tambahkan baris baru untuk nilai wipnya
                        rowWIP = TBLWIP.Newtbltrans_mo_item_wipRow
                        rowWIP.Identifier_JO = rowItemWO.Identifier_JO
                        rowWIP.Identifier_WO = rowItemWO.Identifier
                        rowWIP.JO_Number = vJONO
                        rowWIP.MO_Number = vMONo
                        rowWIP._Date = vMODate
                        rowWIP.Primary_MO = vPrimaryMO
                        rowWIP.Primary_Process = rowItemWO.Primary_Process
                        rowWIP.Primary_Product_Code_WIP = rowItemWO.Primary_Product
                        rowWIP.Qty_Balance = rowItemWO.Qty_WIP
                        rowWIP.WO_Number = rowItemWO.WO_Number

                        TBLWIP.Rows.Add(rowWIP)
                    Else
                        rowWIP = TBLWIP.DefaultView(0).Row
                        rowWIP.Qty_Balance = rowItemWO.Qty_WIP
                        rowWIP.WO_Number = rowItemWO.WO_Number
                        rowWIP.JO_Number = vJONO
                        rowWIP.MO_Number = vMONo
                    End If
                End If
            Next
            'tambahkan item pada item rm req dengan alokasi wip
            For Each rowWIP In TBLWIP.Rows
                rowRMReq = TBLRMReq.Newtbltrans_mo_item_rmRow
                With rowRMReq
                    .Balance_Qty = rowWIP.Qty_Balance
                    .Identifier_JO = rowWIP.Identifier_JO
                    .Identifier_WO = rowWIP.Identifier_WO
                    .JO_Number = rowWIP.JO_Number
                    .MO_Date = rowWIP._Date
                    .MO_Number = rowWIP.MO_Number
                    .Primary_MO = rowWIP.Primary_MO
                    .Primary_Process = rowWIP.Primary_Process
                    .Primary_Product = rowWIP.Primary_Product_Code_WIP
                    .WO_Number = rowWIP.WO_Number
                    .required_date = #1/1/1950#
                    .MP_Spare_Qty = 0
                    .is_wip = 1
                End With
                TBLRMReq.Rows.Add(rowRMReq)
            Next

            'untuk pob, tampilkan form peringatan bila ada po yang otorisasinya tidak sempurna
            CekOtorisasiPO = False
            For Each rowPOB In TBLPOB.Rows
                If rowPOB.Authorization_Level <> rowPOB.Authorization_Level_Needed Then
                    CekOtorisasiPO = True
                    Exit For
                End If
            Next
            If CekOtorisasiPO = True Then
                vFrmCek.vDT_ItemPOB = TBLPOB
                vFrmCek.BackColorAll = vBackColorAll
                vFrmCek.Show()
                vFrmCek.Location = New Point(vFrmCek.Width * -1, vFrmCek.Height * -1)

                'buang item po yang berasal dari transaksi yang tidak includekan
                For Each rowDGV In vFrmCek.dgvItem.Rows
                    If rowDGV.Cells(vFrmCek.ColInclude.Index).Value = Boolean.FalseString Then
                        TBLPOB.DefaultView.RowFilter = "[Primary PO] = " + rowDGV.Cells(vFrmCek.ColPrimaryPO.Index).Value.ToString
                        While TBLPOB.DefaultView.Count > 0
                            rowPOB = TBLPOB.DefaultView(0).Row
                            TBLPOB.RemoveDataTable3Row(rowPOB)
                        End While
                    End If
                Next

                vFrmCek.Close()
            End If


            Me.BuangProductExclude(TBLProduct, vListPrimaryProduct, TBLRMReq, TBLMOB, TBLGIT, TBLPendingDO, TBLPOB, TBLSOB, TBLPrOB, TBLProductQty)

            'untuk baris product yang di exclude tidak dimasukkan ke dalam perhitungan
            IndexDummy = 0
            While IndexDummy <= TBLProduct.Rows.Count - 1
                rowProduct = TBLProduct.Rows(IndexDummy)
                If rowProduct.exclude_on_material_planning = 1 Then
                    'yang di exclude tapi masih memiliki sob, pending do, rm req, pr-ob tetap tidak di exclude
                    TBLSOB.DefaultView.RowFilter = "[Primary Product] = " + rowProduct.Primary.ToString
                    If TBLSOB.DefaultView.Count > 0 Then
                        IndexDummy += 1
                    Else
                        TBLPendingDO.DefaultView.RowFilter = "[Primary Product] = " + rowProduct.Primary.ToString
                        If TBLPendingDO.DefaultView.Count > 0 Then
                            IndexDummy += 1
                        Else
                            TBLRMReq.DefaultView.RowFilter = "[Primary Product] = " + rowProduct.Primary.ToString
                            If TBLRMReq.DefaultView.Count > 0 Then
                                IndexDummy += 1
                            Else
                                TBLPrOB.DefaultView.RowFilter = "[primary code no] = " + rowProduct.Primary.ToString
                                If TBLPrOB.DefaultView.Count > 0 Then
                                    IndexDummy += 1
                                Else
                                    TBLProduct.Rows.Remove(rowProduct)
                                End If
                            End If
                        End If
                    End If
                Else
                    IndexDummy += 1
                End If
            End While
            TBLSOB.DefaultView.RowFilter = ""
            TBLPendingDO.DefaultView.RowFilter = ""
            TBLRMReq.DefaultView.RowFilter = ""
            TBLPrOB.DefaultView.RowFilter = ""

            'location exception
            If vObjSettingUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.MaterialPlanning, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.MaterialPlanning_UseWarehouseExclusion, vUserCode) = "1" Then
                Me.IsikanSetting_ProductQty(TBLProductQty, vUserCode, vDT_ItemSettingLocation)
            Else
                vDT_ItemSettingLocation.Rows.Clear()
            End If

            TBLPrOB.DefaultView.Sort = "[tanggal], [no transaksi]"
            TBLGIT.DefaultView.Sort = "[Tanggal Transaksi], [No Transaksi]"
            TBLMOB.DefaultView.Sort = "[MO Date], [MO Number], [JO Number]"
            TBLRMReq.DefaultView.Sort = "[MO Date], [MO Number], [JO Number], [WO Number]"
            TBLPendingDO.DefaultView.Sort = "[Tanggal Transaksi], [No Transaksi]"
            TBLPOB.DefaultView.Sort = "[Tanggal Transaksi], [No Transaksi]"
            TBLSOB.DefaultView.Sort = "[Tanggal Transaksi], [No Transaksi]"

            For Each rowProduct In TBLProduct.Rows
                rowItem = vDT.Newtbltrans_material_planning_item2Row
                With rowItem
                    .default_prec_unit = rowProduct._Decimal
                    .primary_code_no = rowProduct.Primary
                    .primary_unit = rowProduct.Primary_Unit
                    .qty_safety_stock = rowProduct.safety_stock_qty

                    Filter = "[Primary Product] = " + rowProduct.Primary.ToString
                    TBLGIT.DefaultView.RowFilter = Filter
                    If TBLGIT.DefaultView.Count = 0 Then
                        .qty_git = 0
                    Else
                        JmlDummy = 0
                        For Each DV In TBLGIT.DefaultView
                            rowGIT = DV.Row
                            rowGIT2 = vDT_ItemGIT.Newtbltrans_material_planning_itemgitRow
                            With rowGIT2
                                .No_Packing_List = rowGIT.No_Transaksi
                                .packing_list_date_rec = rowGIT.Tanggal_Transaksi
                                .packing_list_no_rec = rowGIT.No_Transaksi
                                .primary = Me.GetPrimaryMax(vDT_ItemGIT) + 1
                                .primary_packing_list_rec = rowGIT.Primary_Packing_List
                                .primary_product = rowGIT.Primary_Product
                                .primary_suppier_rec = rowGIT.Primary_Supplier
                                .primary_transaksi = -1
                                .qty = rowGIT.Total_Qty
                                .Supplier_Packing_List_No = rowGIT.Supplier_Packing_List_No
                                .supplier_packing_list_no_rec = rowGIT.Supplier_Packing_List_No
                                .Tgl_Packing_List = rowGIT.Tanggal_Transaksi
                            End With
                            vDT_ItemGIT.Rows.Add(rowGIT2)
                            JmlDummy += rowGIT.Total_Qty
                        Next
                        While TBLGIT.DefaultView.Count > 0
                            TBLGIT.Rows.Remove(TBLGIT.DefaultView(0).Row)
                        End While
                        .qty_git = JmlDummy
                    End If

                    If rowProduct.Type = Core_Module.clsMasterProduct.EnumTipeProduct.Services Then
                        .qty_on_hand = 0
                    Else
                        Filter = "[Primary Product] = " + rowProduct.Primary.ToString
                        TBLProductQty.DefaultView.RowFilter = Filter

                        If TBLProductQty.DefaultView.Count = 0 Then
                            .qty_on_hand = 0
                        Else
                            JmlDummy = 0
                            For Each DV In TBLProductQty.DefaultView
                                rowProductQty = DV.Row
                                JmlDummy += rowProductQty.Qty

                                'tambahkan datanya ke item qty
                                If rowProductQty.Qty > 0 Then
                                    rowProductQty1 = vDT_ItemQty.Newtbltrans_material_planning_item_qtyRow
                                    With rowProductQty1
                                        .Expiry_Date = rowProductQty.Expiry_Date
                                        .Lot_No = rowProductQty.Lot_No
                                        .Primary = vDT_ItemQty.Rows.Count + 1
                                        .Primary_Product = rowProductQty.Primary_Product
                                        .Primary_Storage_Area = rowProductQty.Primary_Storage_Area
                                        .primary_transaksi = -1
                                        .Qty = rowProductQty.Qty
                                    End With
                                    vDT_ItemQty.Rows.Add(rowProductQty1)
                                End If
                            Next
                            While TBLProductQty.DefaultView.Count > 0
                                TBLProductQty.Rows.Remove(TBLProductQty.DefaultView(0).Row)
                            End While
                            .qty_on_hand = JmlDummy
                        End If
                    End If

                    .qty_mob = 0
                    Filter = "[Primary Product] = " + rowProduct.Primary.ToString
                    TBLMOB.DefaultView.RowFilter = Filter
                    If TBLMOB.DefaultView.Count = 0 Then
                        .qty_mob = 0
                    Else
                        JmlDummy = 0
                        For Each DV In TBLMOB.DefaultView
                            rowMOB2 = vDT_ItemMOB.Newtbltrans_material_planning_itemmobRow
                            rowMOB = DV.Row

                            With rowMOB2
                                .identifier_jo_rec = rowMOB.Identifier_JO
                                .jo_no_rec = rowMOB.JO_Number
                                .mo_date_rec = rowMOB.MO_Date
                                .mo_no_rec = rowMOB.MO_Number
                                .primary = -1
                                .primary_mo_rec = rowMOB.Primary_MO
                                .primary_product = rowMOB.Primary_Product
                                .primary_transaksi = -1
                                .qty = rowMOB.Balance_Qty
                                .primary_proses = rowMOB.Primary_Process
                                .identifier_wo_rec = rowMOB.Identifier_WO
                                .wo_no_rec = rowMOB.WO_Number

                                JmlDummy += .qty
                            End With
                            vDT_ItemMOB.Rows.Add(rowMOB2)
                        Next
                        While TBLMOB.DefaultView.Count > 0
                            TBLMOB.Rows.Remove(TBLMOB.DefaultView(0).Row)
                        End While
                        .qty_mob = JmlDummy
                    End If

                    .qty_rm_req = 0
                    Filter = "[Primary Product] = " + rowProduct.Primary.ToString
                    TBLRMReq.DefaultView.RowFilter = Filter
                    If TBLRMReq.DefaultView.Count = 0 Then
                        .qty_rm_req = 0
                    Else
                        JmlDummy = 0
                        For Each DV In TBLRMReq.DefaultView
                            rowRMReq2 = vDT_ItemRMReq.Newtbltrans_material_planning_item_rm_reqRow
                            rowRMReq = DV.Row

                            With rowRMReq2
                                .identifier_wo_rec = rowRMReq.Identifier_WO
                                .wo_no_rec = rowRMReq.WO_Number
                                .identifier_jo_rec = rowRMReq.Identifier_JO
                                .jo_no_rec = rowRMReq.JO_Number
                                .mo_date_rec = rowRMReq.MO_Date
                                .mo_no_rec = rowRMReq.MO_Number
                                .primary = -1
                                .primary_mo_rec = rowRMReq.Primary_MO
                                .primary_product = rowRMReq.Primary_Product
                                .primary_transaksi = -1
                                .primary_proses = rowRMReq.Primary_Process
                                .qty = rowRMReq.Balance_Qty
                                .is_wip = rowRMReq.is_wip
                                .MP_Spare_Qty = 0
                                .required_date = #1/1/1950#

                                JmlDummy += .qty
                            End With
                            vDT_ItemRMReq.Rows.Add(rowRMReq2)
                        Next
                        While TBLRMReq.DefaultView.Count > 0
                            TBLRMReq.Rows.Remove(TBLRMReq.DefaultView(0).Row)
                        End While
                        .qty_rm_req = JmlDummy
                    End If

                    .qty_pending_do = 0
                    Filter = "[Primary Product] = " + rowProduct.Primary.ToString
                    TBLPendingDO.DefaultView.RowFilter = Filter
                    If TBLPendingDO.DefaultView.Count = 0 Then
                        .qty_pending_do = 0
                    Else
                        JmlDummy = 0
                        For Each DV In TBLPendingDO.DefaultView
                            rowPendingDO = DV.Row
                            rowPendingDO2 = vDT_ItemDO.Newtbltrans_material_planning_itemdoRow
                            With rowPendingDO2
                                .DO_Date = rowPendingDO.Tanggal_Transaksi
                                .do_date_rec = rowPendingDO.Tanggal_Transaksi
                                .do_no_rec = rowPendingDO.No_Transaksi
                                .DO_Number = rowPendingDO.No_Transaksi
                                .primary = Me.GetPrimaryMax_PendingDO(vDT_ItemDO) + 1
                                .primary_customer_rec = rowPendingDO.Primary_Customer
                                .primary_do_rec = rowPendingDO.Primary_DO
                                .primary_product = rowPendingDO.Primary_Product
                                .primary_transaksi = -1
                                .qty = rowPendingDO.Total_Qty
                            End With
                            vDT_ItemDO.Rows.Add(rowPendingDO2)
                            JmlDummy += rowPendingDO.Total_Qty
                        Next
                        While TBLPendingDO.DefaultView.Count > 0
                            TBLPendingDO.Rows.Remove(TBLPendingDO.DefaultView(0).Row)
                        End While
                        .qty_pending_do = JmlDummy
                    End If

                    Filter = "[Primary Product] = " + rowProduct.Primary.ToString
                    TBLPOB.DefaultView.RowFilter = Filter
                    If TBLPOB.DefaultView.Count = 0 Then
                        .qty_pob = 0
                    Else
                        JmlDummy = 0
                        For Each DV In TBLPOB.DefaultView
                            rowPOB = DV.Row
                            rowPOB2 = vDT_ItemPOB.Newtbltrans_material_planning_itempobRow
                            With rowPOB2
                                .No_Transaksi = rowPOB.No_Transaksi
                                .po_date_rec = rowPOB.Tanggal_Transaksi
                                .po_no_rec = rowPOB.No_Transaksi
                                .primary = Me.GetPrimaryMax_POB(vDT_ItemPOB) + 1
                                .primary_po_rec = rowPOB.Primary_PO
                                .primary_product = rowPOB.Primary_Product
                                .primary_supplier_rec = rowPOB.Primary_Supplier
                                .primary_transaksi = -1
                                .qty = rowPOB.POB
                                .Tanggal_Transaksi = rowPOB.Tanggal_Transaksi
                            End With
                            vDT_ItemPOB.Rows.Add(rowPOB2)
                            JmlDummy += rowPOB.POB
                        Next
                        While TBLPOB.DefaultView.Count > 0
                            TBLPOB.Rows.Remove(TBLPOB.DefaultView(0).Row)
                        End While
                        .qty_pob = JmlDummy
                    End If

                    .qty_sob = 0
                    Filter = "[Primary Product] = " + rowProduct.Primary.ToString
                    TBLSOB.DefaultView.RowFilter = Filter
                    If TBLSOB.DefaultView.Count = 0 Then
                        .qty_sob = 0
                    Else
                        JmlDummy = 0
                        For Each DV In TBLSOB.DefaultView
                            rowSOB = DV.Row
                            rowSOB2 = vDT_ItemSOB.Newtbltrans_material_planning_itemsobRow
                            With rowSOB2
                                .No_Transaksi = rowSOB.No_Transaksi
                                .primary = Me.GetPrimaryMax_SOB(vDT_ItemSOB) + 1
                                .primary_customer_rec = rowSOB.Primary_Customer
                                .primary_product = rowSOB.Primary_Product
                                .primary_so_rec = rowSOB.Primary_Transaksi
                                .primary_transaksi = rowSOB.Primary_Transaksi
                                .qty = rowSOB.SOB
                                .so_date_rec = rowSOB.Tanggal_Transaksi
                                .so_no_rec = rowSOB.No_Transaksi
                                .Tanggal_Transaksi = rowSOB.Tanggal_Transaksi
                                .MP_Spare_Qty = rowSOB.MP_Spare_Qty
                                If rowSOB.Delivery_Date_Cont = True Then
                                    .delivery_date_cont = 1
                                Else
                                    .delivery_date_cont = 0
                                End If
                                .qty_sob_multi = rowSOB.SOB_Multi
                                .ratio_default = rowSOB.Ratio_Default
                                .ratio_reference = rowSOB.Ratio_Reference
                                .primary_unit_multi = rowSOB.Primary_Units_Multi
                                .identifier_so = rowSOB.identifier
                                .primary_unit = rowSOB.Primary_Units
                            End With
                            vDT_ItemSOB.Rows.Add(rowSOB2)
                            JmlDummy += rowSOB.SOB
                        Next
                        While TBLSOB.DefaultView.Count > 0
                            TBLSOB.Rows.Remove(TBLSOB.DefaultView(0).Row)
                        End While
                        .qty_sob = JmlDummy
                    End If

                    .qty_prob = 0
                    Filter = "[primary code no] = " + rowProduct.Primary.ToString
                    TBLPrOB.DefaultView.RowFilter = Filter
                    If TBLPrOB.DefaultView.Count = 0 Then
                        .qty_prob = 0
                    Else
                        JmlDummy = 0
                        For Each DV In TBLPrOB.DefaultView
                            rowPrOB = DV.Row
                            rowPrOB2 = vDT_ItemPrOB.Newtbltrans_material_planning_itempurch_reqRow
                            With rowPrOB2
                                .No_Transaksi = rowPrOB.no_transaksi
                                .primary = Me.GetPrimaryMax_PrOB(vDT_ItemPrOB) + 1
                                .primary_pr_rec = rowPrOB.primary_ref_party
                                .primary_product = rowPrOB.primary_code_no
                                .primary_pr_rec = rowPrOB.primary
                                .primary_transaksi = rowPrOB.primary
                                .qty = rowPrOB.OB
                                .pr_date_rec = rowPrOB.tanggal
                                .pr_no_rec = rowPrOB.no_transaksi
                                .Tanggal_Transaksi = rowPrOB.tanggal
                                .primary_ref_rec = rowPrOB.primary_ref_party
                                .remarks = rowPrOB.remarks_header
                                .remarks = .remarks.Trim
                                .MP_Spare_Qty = 0
                                rowPrOB.remarks_item = rowPrOB.remarks_item.Trim
                                If rowPrOB.remarks_item <> "" Then
                                    If .remarks = "" Then
                                        .remarks = rowPrOB.remarks_item
                                    Else
                                        .remarks = .remarks + " - " + rowPrOB.remarks_item
                                    End If
                                End If
                                .MP_Spare_Qty = 0

                                .delivery_date_cont = rowPrOB.delivery_date_cont
                                .identifier_pr_item = rowPrOB.identifier
                            End With
                            vDT_ItemPrOB.Rows.Add(rowPrOB2)
                            JmlDummy += rowPrOB.OB
                        Next
                        While TBLPrOB.DefaultView.Count > 0
                            TBLPrOB.Rows.Remove(TBLPrOB.DefaultView(0).Row)
                        End While
                        .qty_prob = JmlDummy
                    End If

                    .qty_inquiry = 0
                    FreeStock = .qty_on_hand + .qty_pob + .qty_mob + .qty_git - .qty_inquiry - .qty_sob - .qty_prob - .qty_pending_do - .qty_rm_req
                    .qty_required = FreeStock
                    'If .qty_required < 0 Then
                    '    .qty_required = 0
                    'End If
                End With
                vDT.Rows.Add(rowItem)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub HitungOrderPoint(ByVal vRow As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item1Row, _
    ByVal JmlDaysPerMonth As Double, ByVal vDtpMP_RefDate As Windows.Forms.DateTimePicker, ByVal vObjCD As Core_Module.ClsCekData)
        Dim SSDate As Date = Today
        Try
            With vRow
                If .qty_pob = 0 And .qty_inquiry = 0 And .qty_git = 0 And .qty_sob = 0 And .qty_prob = 0 And .qty_pending_do = 0 And .qty_safety_stock = 0 Then
                    .order_point_delay = 9999
                    .order_point = #1/1/2200#
                    .next_order_point_delay = 9999
                    .next_order_point = #1/1/2200#
                Else
                    If .qty_order = 0 Then
                        .order_point_delay = 9999
                        .order_point = #1/1/2200#
                        .next_order_point_delay = 9999
                        .next_order_point = #1/1/2200#
                    Else
                        .order_point_delay = vObjCD.RoundDown1(((.qty_on_hand + .qty_pob + .qty_git - .qty_inquiry - .qty_sob - .qty_prob - .qty_pending_do - .qty_safety_stock) * JmlDaysPerMonth) / .qty_avg_montly_demand, 0) - .lead_time_days
                        If .order_point_delay > (Date.MaxValue.Date - vDtpMP_RefDate.Value.Date).Days Then
                            .order_point = #1/1/2200#
                        Else
                            .order_point = vDtpMP_RefDate.Value + TimeSpan.FromDays(.order_point_delay)
                        End If
                        SSDate = .order_point.AddDays(.lead_time_days)
                        .next_order_point_delay = vObjCD.RoundDown1((.qty_order * JmlDaysPerMonth) / .qty_avg_montly_demand, 0) - .lead_time_days
                        If .next_order_point_delay > (Date.MaxValue.Date - vDtpMP_RefDate.Value.Date).Days Then
                            .next_order_point = #1/1/2200#
                        Else
                            .next_order_point = SSDate + TimeSpan.FromDays(.next_order_point_delay)
                        End If
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetTheRecentRow(ByVal vTBLPOItem As DseTransaksiMaterialPlanning2.tbltrans_po_itemDataTable, _
  ByVal vPrimaryProduct As Double) As DseTransaksiMaterialPlanning2.tbltrans_po_itemRow
        Dim Dummy As Decimal = Decimal.MaxValue
        Dim RetVal As DseTransaksiMaterialPlanning2.tbltrans_po_itemRow
        Dim rowDummy As DseTransaksiMaterialPlanning2.tbltrans_po_itemRow

        Try
            vTBLPOItem.DefaultView.RowFilter = "[Primary Product] = " + vPrimaryProduct.ToString
            vTBLPOItem.DefaultView.Sort = "[Tanggal Transaksi] DESC, [Primary Transaksi] DESC"
            rowDummy = vTBLPOItem.DefaultView(0).Row
            RetVal = rowDummy

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return Nothing
        End Try
    End Function

    Private Function GetTheCheapestRowOf3(ByVal vTBLPOItem As DseTransaksiMaterialPlanning2.tbltrans_po_itemDataTable, _
    ByVal vPrimaryProduct As Double) As DseTransaksiMaterialPlanning2.tbltrans_po_itemRow
        Dim IndexTesting As Double = 0
        Dim Dummy As Decimal = Decimal.MaxValue
        Dim RetVal As DseTransaksiMaterialPlanning2.tbltrans_po_itemRow
        Dim rowDummy As DseTransaksiMaterialPlanning2.tbltrans_po_itemRow

        Try
            vTBLPOItem.DefaultView.RowFilter = "[Primary Product] = " + vPrimaryProduct.ToString
            vTBLPOItem.DefaultView.Sort = "[Tanggal Transaksi] DESC, [Primary Transaksi] DESC"

            While IndexTesting <= 2 And IndexTesting <= vTBLPOItem.DefaultView.Count - 1
                rowDummy = vTBLPOItem.DefaultView(IndexTesting).Row
                If rowDummy.Nett_Price * rowDummy.Default_Exchange_Rate < Dummy Then
                    Dummy = rowDummy.Nett_Price * rowDummy.Default_Exchange_Rate
                    RetVal = rowDummy
                End If

                IndexTesting += 1
            End While

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return Nothing
        End Try
    End Function

    Private Function GetPrimaryMax_SOB(ByVal vDT_ItemSOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsobDataTable) As Double
        Dim RetVal As Double = 0

        Try
            For Each rowSOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsobRow In vDT_ItemSOB.Rows
                If rowSOB.primary > RetVal Then
                    RetVal = rowSOB.primary
                End If
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetPrimaryMax_PrOB(ByVal vDT_ItemPrOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_reqDataTable) As Double
        Dim RetVal As Double = 0

        Try
            For Each RowPrOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_reqRow In vDT_ItemPrOB.Rows
                If RowPrOB.primary > RetVal Then
                    RetVal = RowPrOB.primary
                End If
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetPrimaryMax_POB(ByVal vDT_ItemPOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempobDataTable) As Double
        Dim RetVal As Double = 0

        Try
            For Each rowPOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempobRow In vDT_ItemPOB.Rows
                If rowPOB.primary > RetVal Then
                    RetVal = rowPOB.primary
                End If
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetPrimaryMax_PendingDO(ByVal vDT_ItemDO As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemdoDataTable) As Double
        Dim RetVal As Double = 0

        Try
            For Each rowDO As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemdoRow In vDT_ItemDO.Rows
                If rowDO.primary > RetVal Then
                    RetVal = rowDO.primary
                End If
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetPrimaryMax(ByVal vDT_ItemGIT As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemgitDataTable) As Double
        Dim RetVal As Double = 0

        Try
            For Each rowGIT As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemgitRow In vDT_ItemGIT.Rows
                If rowGIT.primary > RetVal Then
                    RetVal = rowGIT.primary
                End If
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub IsikanSetting_ProductQty(ByVal vTBLProductQty As DseTransaksiMaterialPlanning2.tblmaster_productqtyDataTable, ByVal vUserCode As String, _
    ByVal vDT_ItemSettingLocation As DseTransaksiMaterialPlanning_WarehouseExc.tbltrans_material_planning_item_location_exclusionDataTable, _
    ByVal vTxtMP_Primary As Windows.Forms.TextBox)
        Dim DTA_Setting_Location As New DseTransaksiMaterialPlanning_WarehouseExcTableAdapters.tbltrans_setting_warehouse_exclusionTableAdapter
        Dim TBL_Setting_Location As New DseTransaksiMaterialPlanning_WarehouseExc.tbltrans_setting_warehouse_exclusionDataTable
        Dim row_Setting_Location As DseTransaksiMaterialPlanning_WarehouseExc.tbltrans_setting_warehouse_exclusionRow
        Dim Level_R As FrmTrans_MaterialPlanningDashboard_Adv_Exclusion.EnumLocationLevel = FrmTrans_MaterialPlanningDashboard_Adv_Exclusion.EnumLocationLevel.Warehouse
        Dim DTALocation As New DseTransaksiMaterialPlanning2TableAdapters.tblmaster_storage_areaTableAdapter
        Dim TBLLocation As New DseTransaksiMaterialPlanning2.tblmaster_storage_areaDataTable
        Dim TBLLocation_D As New DseTransaksiMaterialPlanning2.tblmaster_storage_areaDataTable
        Dim rowLocation, rowLocation_D As DseTransaksiMaterialPlanning2.tblmaster_storage_areaRow
        Dim Inc_Excl As FrmTrans_MaterialPlanningDashboard_Adv_Exclusion.EnumExcludeInclude = FrmTrans_MaterialPlanningDashboard_Adv_Exclusion.EnumExcludeInclude.Exclude
        Dim TBLProductQty_D As New DseTransaksiMaterialPlanning2.tblmaster_productqtyDataTable
        Dim rowProductQty, rowProductQty_D As DseTransaksiMaterialPlanning2.tblmaster_productqtyRow
        Dim ListPrimaryLocation As New List(Of Double)
        Dim rowSettingLocation_Trans As DseTransaksiMaterialPlanning_WarehouseExc.tbltrans_material_planning_item_location_exclusionRow

        Try
            DTA_Setting_Location.Connection.ConnectionString = Core_Module.GetConstring
            DTALocation.Connection.ConnectionString = Core_Module.GetConstring

            'tarik data setting qty
            DTA_Setting_Location.FillBy_kode_user(TBL_Setting_Location, vUserCode)
            'pindahkan data setting untuk simpan per transaksi
            vDT_ItemSettingLocation.Rows.Clear()
            For Each row_Setting_Location In TBL_Setting_Location.Rows
                rowSettingLocation_Trans = vDT_ItemSettingLocation.Newtbltrans_material_planning_item_location_exclusionRow
                With rowSettingLocation_Trans
                    .include_exclude = row_Setting_Location.include_exclude
                    .level = row_Setting_Location.level
                    .primary = Me.GetPrimaryMax(vDT_ItemSettingLocation) + 1
                    .primary_acuan = row_Setting_Location.primary_acuan
                    .primary_transaksi = vTxtMP_Primary.Text
                    .remarks = row_Setting_Location.remarks
                End With
                vDT_ItemSettingLocation.Rows.Add(rowSettingLocation_Trans)
            Next

            'tandain, warehouse yang include/exclude
            Level_R = FrmTrans_MaterialPlanningDashboard_Adv_Exclusion.EnumLocationLevel.Warehouse
            For Each row_Setting_Location In TBL_Setting_Location.Rows
                If row_Setting_Location.level = CInt(Level_R) Then
                    TBLLocation_D.Rows.Clear()
                    DTALocation.FillBy_PrimaryWarehouse(TBLLocation_D, row_Setting_Location.primary_acuan)
                    For Each rowLocation_D In TBLLocation_D.Rows
                        rowLocation = TBLLocation.Newtblmaster_storage_areaRow
                        For ij As Integer = 0 To TBLLocation.Columns.Count - 1
                            rowLocation.Item(ij) = rowLocation_D.Item(ij)
                        Next
                        rowLocation.Include = row_Setting_Location.include_exclude
                        TBLLocation.Rows.Add(rowLocation)
                    Next
                End If
            Next

            'tandain, location yang include/exclude
            Level_R = FrmTrans_MaterialPlanningDashboard_Adv_Exclusion.EnumLocationLevel.Location
            TBLLocation_D.Rows.Clear()
            DTALocation.Fill(TBLLocation_D)
            For Each row_Setting_Location In TBL_Setting_Location.Rows
                If row_Setting_Location.level = CInt(Level_R) Then
                    TBLLocation_D.DefaultView.RowFilter = "Primary = " + row_Setting_Location.primary_acuan.ToString
                    For Each DV As Data.DataRowView In TBLLocation_D.DefaultView
                        rowLocation_D = DV.Row
                        TBLLocation.DefaultView.RowFilter = "Primary = " + row_Setting_Location.primary_acuan.ToString

                        If TBLLocation.DefaultView.Count = 0 Then
                            'copy barisnya
                            rowLocation = TBLLocation.Newtblmaster_storage_areaRow
                            For ij As Integer = 0 To TBLLocation.Columns.Count - 1
                                rowLocation.Item(ij) = rowLocation_D.Item(ij)
                            Next
                            rowLocation.Include = row_Setting_Location.include_exclude
                            TBLLocation.Rows.Add(rowLocation)
                        Else
                            'hanya copy settingnya saja
                            rowLocation = TBLLocation.DefaultView(0).Row
                            rowLocation.Include = row_Setting_Location.include_exclude
                        End If
                    Next
                End If
            Next

            'buang baris location yang tidak include
            Inc_Excl = FrmTrans_MaterialPlanningDashboard_Adv_Exclusion.EnumExcludeInclude.Exclude
            For ij As Double = TBLLocation.Rows.Count - 1 To 0 Step -1
                rowLocation = TBLLocation.Rows(ij)
                If rowLocation.Include = CInt(Inc_Excl) Then
                    TBLLocation.Rows.Remove(rowLocation)
                End If
            Next

            'untuk setiap baris master product, 
            For Each rowProductQty In vTBLProductQty.Rows
                If rowProductQty.Primary_Storage_Area = -1 Then
                    'jika primary location =-1, maka langsung masukkan ke dalam tabel qty
                    rowProductQty_D = TBLProductQty_D.Newtblmaster_productqtyRow
                    For ij As Integer = 0 To TBLProductQty_D.Columns.Count - 1
                        rowProductQty_D.Item(ij) = rowProductQty.Item(ij)
                    Next
                    TBLProductQty_D.Rows.Add(rowProductQty_D)
                Else
                    'jika primary location <> -1, maka uji kembali apakah include di dalam setting location.
                    'yang tidak include maka di buang
                    If ListPrimaryLocation.Contains(rowProductQty.Primary_Storage_Area) = False Then
                        TBLLocation.DefaultView.RowFilter = "Primary = " + rowProductQty.Primary_Storage_Area.ToString
                        If TBLLocation.DefaultView.Count > 0 Then
                            'tandain bahwa itu sdah di transfer
                            ListPrimaryLocation.Add(rowProductQty.Primary_Storage_Area)

                            'transfer semua product yang ada di lokasi itu
                            Me.Transfer_RowTblMasterProductQty(rowProductQty.Primary_Storage_Area, vTBLProductQty, TBLProductQty_D)
                        End If
                    End If
                End If
            Next

            'kosongkan tabel master product qty
            vTBLProductQty.Rows.Clear()

            'transfer balik ke tabel master product qty
            For Each rowProductQty_D In TBLProductQty_D.Rows
                rowProductQty = vTBLProductQty.Newtblmaster_productqtyRow
                For ij As Integer = 0 To vTBLProductQty.Columns.Count - 1
                    rowProductQty.Item(ij) = rowProductQty_D.Item(ij)
                Next
                vTBLProductQty.Rows.Add(rowProductQty)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_Setting_Location.Dispose()
            TBL_Setting_Location.Dispose()
            DTALocation.Dispose()
            TBLLocation.Dispose()
            TBLLocation_D.Dispose()
            TBLProductQty_D.Dispose()
            ListPrimaryLocation = Nothing
        End Try
    End Sub

    Private Sub IsikanSetting_ProductQty(ByVal vTBLProductQty As DseTransaksiMaterialPlanning2.tblmaster_productqtyDataTable, ByVal vUserCode As String, _
   ByVal vDT_ItemSettingLocation As DseTransaksiMaterialPlanning_WarehouseExc.tbltrans_material_planning_item_location_exclusionDataTable)
        Dim DTA_Setting_Location As New DseTransaksiMaterialPlanning_WarehouseExcTableAdapters.tbltrans_setting_warehouse_exclusionTableAdapter
        Dim TBL_Setting_Location As New DseTransaksiMaterialPlanning_WarehouseExc.tbltrans_setting_warehouse_exclusionDataTable
        Dim row_Setting_Location As DseTransaksiMaterialPlanning_WarehouseExc.tbltrans_setting_warehouse_exclusionRow
        Dim Level_R As FrmTrans_MaterialPlanningDashboard_Adv_Exclusion.EnumLocationLevel = FrmTrans_MaterialPlanningDashboard_Adv_Exclusion.EnumLocationLevel.Warehouse
        Dim DTALocation As New DseTransaksiMaterialPlanning2TableAdapters.tblmaster_storage_areaTableAdapter
        Dim TBLLocation As New DseTransaksiMaterialPlanning2.tblmaster_storage_areaDataTable
        Dim TBLLocation_D As New DseTransaksiMaterialPlanning2.tblmaster_storage_areaDataTable
        Dim rowLocation, rowLocation_D As DseTransaksiMaterialPlanning2.tblmaster_storage_areaRow
        Dim Inc_Excl As FrmTrans_MaterialPlanningDashboard_Adv_Exclusion.EnumExcludeInclude = FrmTrans_MaterialPlanningDashboard_Adv_Exclusion.EnumExcludeInclude.Exclude
        Dim TBLProductQty_D As New DseTransaksiMaterialPlanning2.tblmaster_productqtyDataTable
        Dim rowProductQty, rowProductQty_D As DseTransaksiMaterialPlanning2.tblmaster_productqtyRow
        Dim ListPrimaryLocation As New List(Of Double)
        Dim rowSettingLocation_Trans As DseTransaksiMaterialPlanning_WarehouseExc.tbltrans_material_planning_item_location_exclusionRow

        Try
            DTA_Setting_Location.Connection.ConnectionString = Core_Module.GetConstring
            DTALocation.Connection.ConnectionString = Core_Module.GetConstring

            'tarik data setting qty
            DTA_Setting_Location.FillBy_kode_user(TBL_Setting_Location, vUserCode)
            'pindahkan data setting untuk simpan per transaksi
            vDT_ItemSettingLocation.Rows.Clear()
            For Each row_Setting_Location In TBL_Setting_Location.Rows
                rowSettingLocation_Trans = vDT_ItemSettingLocation.Newtbltrans_material_planning_item_location_exclusionRow
                With rowSettingLocation_Trans
                    .include_exclude = row_Setting_Location.include_exclude
                    .level = row_Setting_Location.level
                    .primary = Me.GetPrimaryMax(vDT_ItemSettingLocation) + 1
                    .primary_acuan = row_Setting_Location.primary_acuan
                    .primary_transaksi = -1
                    .remarks = row_Setting_Location.remarks
                End With
                vDT_ItemSettingLocation.Rows.Add(rowSettingLocation_Trans)
            Next

            'tandain, warehouse yang include/exclude
            Level_R = FrmTrans_MaterialPlanningDashboard_Adv_Exclusion.EnumLocationLevel.Warehouse
            For Each row_Setting_Location In TBL_Setting_Location.Rows
                If row_Setting_Location.level = CInt(Level_R) Then
                    TBLLocation_D.Rows.Clear()
                    DTALocation.FillBy_PrimaryWarehouse(TBLLocation_D, row_Setting_Location.primary_acuan)
                    For Each rowLocation_D In TBLLocation_D.Rows
                        rowLocation = TBLLocation.Newtblmaster_storage_areaRow
                        For ij As Integer = 0 To TBLLocation.Columns.Count - 1
                            rowLocation.Item(ij) = rowLocation_D.Item(ij)
                        Next
                        rowLocation.Include = row_Setting_Location.include_exclude
                        TBLLocation.Rows.Add(rowLocation)
                    Next
                End If
            Next

            'tandain, location yang include/exclude
            Level_R = FrmTrans_MaterialPlanningDashboard_Adv_Exclusion.EnumLocationLevel.Location
            TBLLocation_D.Rows.Clear()
            DTALocation.Fill(TBLLocation_D)
            For Each row_Setting_Location In TBL_Setting_Location.Rows
                If row_Setting_Location.level = CInt(Level_R) Then
                    TBLLocation_D.DefaultView.RowFilter = "Primary = " + row_Setting_Location.primary_acuan.ToString
                    For Each DV As Data.DataRowView In TBLLocation_D.DefaultView
                        rowLocation_D = DV.Row
                        TBLLocation.DefaultView.RowFilter = "Primary = " + row_Setting_Location.primary_acuan.ToString

                        If TBLLocation.DefaultView.Count = 0 Then
                            'copy barisnya
                            rowLocation = TBLLocation.Newtblmaster_storage_areaRow
                            For ij As Integer = 0 To TBLLocation.Columns.Count - 1
                                rowLocation.Item(ij) = rowLocation_D.Item(ij)
                            Next
                            rowLocation.Include = row_Setting_Location.include_exclude
                            TBLLocation.Rows.Add(rowLocation)
                        Else
                            'hanya copy settingnya saja
                            rowLocation = TBLLocation.DefaultView(0).Row
                            rowLocation.Include = row_Setting_Location.include_exclude
                        End If
                    Next
                End If
            Next

            'buang baris location yang tidak include
            Inc_Excl = FrmTrans_MaterialPlanningDashboard_Adv_Exclusion.EnumExcludeInclude.Exclude
            For ij As Double = TBLLocation.Rows.Count - 1 To 0 Step -1
                rowLocation = TBLLocation.Rows(ij)
                If rowLocation.Include = CInt(Inc_Excl) Then
                    TBLLocation.Rows.Remove(rowLocation)
                End If
            Next

            'untuk setiap baris master product, 
            For Each rowProductQty In vTBLProductQty.Rows
                If rowProductQty.Primary_Storage_Area = -1 Then
                    'jika primary location =-1, maka langsung masukkan ke dalam tabel qty
                    rowProductQty_D = TBLProductQty_D.Newtblmaster_productqtyRow
                    For ij As Integer = 0 To TBLProductQty_D.Columns.Count - 1
                        rowProductQty_D.Item(ij) = rowProductQty.Item(ij)
                    Next
                    TBLProductQty_D.Rows.Add(rowProductQty_D)
                Else
                    'jika primary location <> -1, maka uji kembali apakah include di dalam setting location.
                    'yang tidak include maka di buang
                    If ListPrimaryLocation.Contains(rowProductQty.Primary_Storage_Area) = False Then
                        TBLLocation.DefaultView.RowFilter = "Primary = " + rowProductQty.Primary_Storage_Area.ToString
                        If TBLLocation.DefaultView.Count > 0 Then
                            'tandain bahwa itu sdah di transfer
                            ListPrimaryLocation.Add(rowProductQty.Primary_Storage_Area)

                            'transfer semua product yang ada di lokasi itu
                            Me.Transfer_RowTblMasterProductQty(rowProductQty.Primary_Storage_Area, vTBLProductQty, TBLProductQty_D)
                        End If
                    End If
                End If
            Next

            'kosongkan tabel master product qty
            vTBLProductQty.Rows.Clear()

            'transfer balik ke tabel master product qty
            For Each rowProductQty_D In TBLProductQty_D.Rows
                rowProductQty = vTBLProductQty.Newtblmaster_productqtyRow
                For ij As Integer = 0 To vTBLProductQty.Columns.Count - 1
                    rowProductQty.Item(ij) = rowProductQty_D.Item(ij)
                Next
                vTBLProductQty.Rows.Add(rowProductQty)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_Setting_Location.Dispose()
            TBL_Setting_Location.Dispose()
            DTALocation.Dispose()
            TBLLocation.Dispose()
            TBLLocation_D.Dispose()
            TBLProductQty_D.Dispose()
            ListPrimaryLocation = Nothing
        End Try
    End Sub

    Private Sub Transfer_RowTblMasterProductQty(ByVal vPrimaryLocation As Double, ByVal vTBLProductQty As DseTransaksiMaterialPlanning2.tblmaster_productqtyDataTable, _
   ByVal vTBLProductQtyTarget As DseTransaksiMaterialPlanning2.tblmaster_productqtyDataTable)
        Dim rowProductQty, rowProductQty_D As DseTransaksiMaterialPlanning2.tblmaster_productqtyRow

        Try
            For Each rowProductQty In vTBLProductQty.Rows
                If rowProductQty.Primary_Storage_Area = vPrimaryLocation Then
                    rowProductQty_D = vTBLProductQtyTarget.Newtblmaster_productqtyRow
                    For ij As Integer = 0 To vTBLProductQty.Columns.Count - 1
                        rowProductQty_D.Item(ij) = rowProductQty.Item(ij)
                    Next
                    vTBLProductQtyTarget.Rows.Add(rowProductQty_D)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetPrimaryMax(ByVal TBL_Setting_Location As DseTransaksiMaterialPlanning_WarehouseExc.tbltrans_material_planning_item_location_exclusionDataTable) As Double
        Dim row_Setting_Location As DseTransaksiMaterialPlanning_WarehouseExc.tbltrans_material_planning_item_location_exclusionRow
        Dim RetVal As Double = Double.MinValue

        Try
            For Each row_Setting_Location In TBL_Setting_Location.Rows
                If row_Setting_Location.primary > RetVal Then
                    RetVal = row_Setting_Location.primary
                End If
            Next

            If RetVal = Double.MinValue Then
                RetVal = 0
            End If
            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub BuangProductExclude(ByVal vTBLProduct As DseTransaksiMaterialPlanning2.DataTable1DataTable, ByVal vListPrimaryProduct As List(Of Double), _
    ByVal TBLRMReq As DseTransaksiMaterialPlanning2.tbltrans_mo_item_rmDataTable, ByVal TBLMOB As DseTransaksiMaterialPlanning2.tbltrans_mo_itemDataTable, _
    ByVal TBLGIT As DseTransaksiMaterialPlanning2.tbltrans_packinglist_itemDataTable, ByVal TBLPendingDO As DseTransaksiMaterialPlanning2.DataTable2DataTable, _
    ByVal TBLPOB As DseTransaksiMaterialPlanning2.DataTable3DataTable, ByVal TBLSOB As DseTransaksiMaterialPlanning2.DataTable4DataTable, _
    ByVal TBLPrOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_req1DataTable, ByVal TBLProductQty As DseTransaksiMaterialPlanning2.tblmaster_productqtyDataTable)
        Dim vrowProduct As DseTransaksiMaterialPlanning2.DataTable1Row

        Dim TBLRMReqDummy As New DseTransaksiMaterialPlanning2.tbltrans_mo_item_rmDataTable
        Dim rowRmReq As DseTransaksiMaterialPlanning2.tbltrans_mo_item_rmRow
        Dim mn, start_idx, start_idx1, start_idx2, start_idx3, start_idx4, start_idx5, start_idx6, start_idx8 As Double
        Dim DummySort As String = vTBLProduct.DefaultView.Sort

        Dim rowMOB As DseTransaksiMaterialPlanning2.tbltrans_mo_itemRow
        Dim rowGIT As DseTransaksiMaterialPlanning2.tbltrans_packinglist_itemRow
        Dim rowPendingDO As DseTransaksiMaterialPlanning2.DataTable2Row
        Dim rowPOB As DseTransaksiMaterialPlanning2.DataTable3Row
        Dim rowSOB As DseTransaksiMaterialPlanning2.DataTable4Row
        Dim rowPrOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_req1Row
        Dim rowProductQty As DseTransaksiMaterialPlanning2.tblmaster_productqtyRow

        Try
            TBLRMReq.DefaultView.RowFilter = ""
            TBLMOB.DefaultView.RowFilter = ""
            TBLGIT.DefaultView.RowFilter = ""
            TBLPendingDO.DefaultView.RowFilter = ""
            TBLPOB.DefaultView.RowFilter = ""
            TBLSOB.DefaultView.RowFilter = ""
            TBLPrOB.DefaultView.RowFilter = ""
            TBLProductQty.DefaultView.RowFilter = ""

            TBLRMReq.DefaultView.Sort = "[Primary Product]"
            TBLMOB.DefaultView.Sort = "[Primary Product]"
            TBLGIT.DefaultView.Sort = "[Primary Product]"
            TBLPendingDO.DefaultView.Sort = "[Primary Product]"
            TBLPOB.DefaultView.Sort = "[Primary Product]"
            TBLSOB.DefaultView.Sort = "[Primary Product]"
            TBLPrOB.DefaultView.Sort = "[primary code no]"
            TBLProductQty.DefaultView.Sort = "[Primary Product]"
            vTBLProduct.DefaultView.Sort = "Primary DESC"

            For ij As Double = vTBLProduct.DefaultView.Count - 1 To 0 Step -1
                vrowProduct = vTBLProduct.DefaultView(ij).Row
                If vListPrimaryProduct.Contains(vrowProduct.Primary) = False Then
                    start_idx = 0
                    start_idx1 = 0
                    start_idx2 = 0
                    start_idx3 = 0
                    start_idx4 = 0
                    start_idx5 = 0
                    start_idx6 = 0
                    start_idx8 = 0

                    mn = 0
                    While start_idx + mn <= TBLRMReq.DefaultView.Count - 1
                        rowRmReq = TBLRMReq.DefaultView(mn + start_idx).Row
                        If rowRmReq.Primary_Product < vrowProduct.Primary Then
                            mn += 1
                        ElseIf rowRmReq.Primary_Product = vrowProduct.Primary Then
                            TBLRMReq.Rows.Remove(rowRmReq)
                        ElseIf rowRmReq.Primary_Product > vrowProduct.Primary Then
                            start_idx = mn + start_idx
                            Exit While
                        End If
                    End While

                    mn = 0
                    While start_idx1 + mn <= TBLMOB.DefaultView.Count - 1
                        rowMOB = TBLMOB.DefaultView(mn + start_idx1).Row
                        If rowMOB.Primary_Product < vrowProduct.Primary Then
                            mn += 1
                        ElseIf rowMOB.Primary_Product = vrowProduct.Primary Then
                            TBLMOB.Rows.Remove(rowMOB)
                        ElseIf rowMOB.Primary_Product > vrowProduct.Primary Then
                            start_idx1 = mn + start_idx1
                            Exit While
                        End If
                    End While

                    mn = 0
                    While start_idx2 + mn <= TBLGIT.DefaultView.Count - 1
                        rowGIT = TBLGIT.DefaultView(mn + start_idx2).Row
                        If rowGIT.Primary_Product < vrowProduct.Primary Then
                            mn += 1
                        ElseIf rowGIT.Primary_Product = vrowProduct.Primary Then
                            TBLGIT.Rows.Remove(rowGIT)
                        ElseIf rowGIT.Primary_Product > vrowProduct.Primary Then
                            start_idx2 = mn + start_idx2
                            Exit While
                        End If
                    End While

                    mn = 0
                    While start_idx3 + mn <= TBLPendingDO.DefaultView.Count - 1
                        rowPendingDO = TBLPendingDO.DefaultView(mn + start_idx3).Row
                        If rowPendingDO.Primary_Product < vrowProduct.Primary Then
                            mn += 1
                        ElseIf rowPendingDO.Primary_Product = vrowProduct.Primary Then
                            TBLPendingDO.Rows.Remove(rowPendingDO)
                        ElseIf rowPendingDO.Primary_Product > vrowProduct.Primary Then
                            start_idx3 = mn + start_idx3
                            Exit While
                        End If
                    End While

                    mn = 0
                    While start_idx4 + mn <= TBLPOB.DefaultView.Count - 1
                        rowPOB = TBLPOB.DefaultView(mn + start_idx4).Row
                        If rowPOB.Primary_Product < vrowProduct.Primary Then
                            mn += 1
                        ElseIf rowPOB.Primary_Product = vrowProduct.Primary Then
                            TBLPOB.Rows.Remove(rowPOB)
                        ElseIf rowPOB.Primary_Product > vrowProduct.Primary Then
                            start_idx4 = mn + start_idx4
                            Exit While
                        End If
                    End While

                    mn = 0
                    While start_idx5 + mn <= TBLSOB.DefaultView.Count - 1
                        rowSOB = TBLSOB.DefaultView(mn + start_idx5).Row
                        If rowSOB.Primary_Product < vrowProduct.Primary Then
                            mn += 1
                        ElseIf rowSOB.Primary_Product = vrowProduct.Primary Then
                            TBLSOB.Rows.Remove(rowSOB)
                        ElseIf rowSOB.Primary_Product > vrowProduct.Primary Then
                            start_idx5 = mn + start_idx5
                            Exit While
                        End If
                    End While

                    mn = 0
                    While start_idx6 + mn <= TBLPrOB.DefaultView.Count - 1
                        rowPrOB = TBLPrOB.DefaultView(mn + start_idx6).Row
                        If rowPrOB.primary_code_no < vrowProduct.Primary Then
                            mn += 1
                        ElseIf rowPrOB.primary_code_no = vrowProduct.Primary Then
                            TBLPrOB.Rows.Remove(rowPrOB)
                        ElseIf rowPrOB.primary_code_no > vrowProduct.Primary Then
                            start_idx6 = mn + start_idx6
                            Exit While
                        End If
                    End While

                    mn = 0
                    While start_idx8 + mn <= TBLProductQty.DefaultView.Count - 1
                        rowProductQty = TBLProductQty.DefaultView(mn + start_idx8).Row
                        If rowProductQty.Primary_Product < vrowProduct.Primary Then
                            mn += 1
                        ElseIf rowProductQty.Primary_Product = vrowProduct.Primary Then
                            TBLProductQty.Rows.Remove(rowProductQty)
                        ElseIf rowProductQty.Primary_Product > vrowProduct.Primary Then
                            start_idx8 = mn + start_idx8
                            Exit While
                        End If
                    End While

                    vTBLProduct.Rows.Remove(vrowProduct)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            vTBLProduct.DefaultView.Sort = DummySort
        End Try
    End Sub

    Private Sub BuangProductExclude(ByVal vTBLProduct As DseTransaksiMaterialPlanning2.DataTable1DataTable, ByVal vUserCode As String, _
    ByVal TBLRMReq As DseTransaksiMaterialPlanning2.tbltrans_mo_item_rmDataTable, ByVal TBLMOB As DseTransaksiMaterialPlanning2.tbltrans_mo_itemDataTable, _
    ByVal TBLGIT As DseTransaksiMaterialPlanning2.tbltrans_packinglist_itemDataTable, ByVal TBLPendingDO As DseTransaksiMaterialPlanning2.DataTable2DataTable, _
    ByVal TBLPOB As DseTransaksiMaterialPlanning2.DataTable3DataTable, ByVal TBLSOB As DseTransaksiMaterialPlanning2.DataTable4DataTable, _
    ByVal TBLPrOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_req1DataTable, ByVal TBLPOItem As DseTransaksiMaterialPlanning2.tbltrans_po_itemDataTable, _
    ByVal TBLProductQty As DseTransaksiMaterialPlanning2.tblmaster_productqtyDataTable)
        Dim vrowProduct As DseTransaksiMaterialPlanning2.DataTable1Row
        Dim DTAGroup As New DseTransaksiMaterialPlanning2TableAdapters.tbltrans_material_planning_item_adv_exclusionTableAdapter
        Dim TBLGroup As New DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_adv_exclusionDataTable
        Dim rowGroup As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_adv_exclusionRow
        Dim ProductLevel As ClsTransaksiMaterialPlanningGroup.ProductLevel = ClsTransaksiMaterialPlanningGroup.ProductLevel.Brand
        Dim Filter As String = ""
        Dim TipeExclusion As ClsTransaksiMaterialPlanning.EnumTipeExclusion = ClsTransaksiMaterialPlanning.EnumTipeExclusion.Exclude
        Dim TblDummy As New DseTransaksiMaterialPlanning2.DataTable1DataTable
        Dim rowDummy As DseTransaksiMaterialPlanning2.DataTable1Row

        Dim TBLRMReqDummy As New DseTransaksiMaterialPlanning2.tbltrans_mo_item_rmDataTable
        Dim rowRmReq, rowRmReqDummy As DseTransaksiMaterialPlanning2.tbltrans_mo_item_rmRow
        Dim mn, start_idx, start_idx1, start_idx2, start_idx3, start_idx4, start_idx5, start_idx6, start_idx7, start_idx8 As Double
        Dim DummySort As String = vTBLProduct.DefaultView.Sort

        Dim TBLMOBDummy As New DseTransaksiMaterialPlanning2.tbltrans_mo_itemDataTable
        Dim rowMOB, rowMOBDummy As DseTransaksiMaterialPlanning2.tbltrans_mo_itemRow

        Dim TBLGITDummy As New DseTransaksiMaterialPlanning2.tbltrans_packinglist_itemDataTable
        Dim rowGIT, rowGITDummy As DseTransaksiMaterialPlanning2.tbltrans_packinglist_itemRow

        Dim TBLPendingDODummy As New DseTransaksiMaterialPlanning2.DataTable2DataTable
        Dim rowPendingDO, rowPendingDODummy As DseTransaksiMaterialPlanning2.DataTable2Row

        Dim TBLPOBDummy As New DseTransaksiMaterialPlanning2.DataTable3DataTable
        Dim rowPOB, rowPOBDummy As DseTransaksiMaterialPlanning2.DataTable3Row

        Dim TBLSOBDummy As New DseTransaksiMaterialPlanning2.DataTable4DataTable
        Dim rowSOB, rowSOBDummy As DseTransaksiMaterialPlanning2.DataTable4Row

        Dim TBLPrOBDummy As New DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_req1DataTable
        Dim rowPrOB, rowPrOBDummy As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_req1Row

        Dim TBLPOItemDummy As New DseTransaksiMaterialPlanning2.tbltrans_po_itemDataTable
        Dim rowPOItem, rowPOItemDummy As DseTransaksiMaterialPlanning2.tbltrans_po_itemRow

        Dim TBLProductQtyDummy As New DseTransaksiMaterialPlanning2.tblmaster_productqtyDataTable
        Dim rowProductQty, rowProductQtyDummy As DseTransaksiMaterialPlanning2.tblmaster_productqtyRow

        Try
            TBLRMReq.DefaultView.RowFilter = ""
            TBLMOB.DefaultView.RowFilter = ""
            TBLGIT.DefaultView.RowFilter = ""
            TBLPendingDO.DefaultView.RowFilter = ""
            TBLPOB.DefaultView.RowFilter = ""
            TBLSOB.DefaultView.RowFilter = ""
            TBLPrOB.DefaultView.RowFilter = ""
            TBLPOItem.DefaultView.RowFilter = ""
            TBLProductQty.DefaultView.RowFilter = ""

            TBLRMReq.DefaultView.Sort = "[Primary Product]"
            TBLMOB.DefaultView.Sort = "[Primary Product]"
            TBLGIT.DefaultView.Sort = "[Primary Product]"
            TBLPendingDO.DefaultView.Sort = "[Primary Product]"
            TBLPOB.DefaultView.Sort = "[Primary Product]"
            TBLSOB.DefaultView.Sort = "[Primary Product]"
            TBLPrOB.DefaultView.Sort = "[primary code no]"
            TBLPOItem.DefaultView.Sort = "[Primary Product]"
            TBLProductQty.DefaultView.Sort = "[Primary Product]"
            vTBLProduct.DefaultView.Sort = "Primary"

            DTAGroup.Connection.ConnectionString = Core_Module.GetConstring
            DTAGroup.Fill(TBLGroup, vUserCode, CInt(ClsTransaksiMaterialPlanning.EnumTipeTransaksiExclusion.MaterialPlanning))
            If TBLGroup.Rows.Count > 0 Then
                rowGroup = TBLGroup.Rows(0)
                TipeExclusion = rowGroup.tipe_exclusion

                If TipeExclusion = ClsTransaksiMaterialPlanning.EnumTipeExclusion.Exclude Then
                    For Each rowGroup In TBLGroup.Rows
                        ProductLevel = rowGroup.level
                        With rowGroup
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

                            vTBLProduct.DefaultView.RowFilter = Filter

                            start_idx = 0
                            start_idx1 = 0
                            start_idx2 = 0
                            start_idx3 = 0
                            start_idx4 = 0
                            start_idx5 = 0
                            start_idx6 = 0
                            start_idx7 = 0
                            start_idx8 = 0
                            While vTBLProduct.DefaultView.Count > 0
                                vrowProduct = vTBLProduct.DefaultView(0).Row

                                mn = 0
                                While start_idx + mn <= TBLRMReq.DefaultView.Count - 1
                                    rowRmReq = TBLRMReq.DefaultView(mn + start_idx).Row
                                    If rowRmReq.Primary_Product < vrowProduct.Primary Then
                                        mn += 1
                                    ElseIf rowRmReq.Primary_Product = vrowProduct.Primary Then
                                        TBLRMReq.Rows.Remove(rowRmReq)
                                    ElseIf rowRmReq.Primary_Product > vrowProduct.Primary Then
                                        start_idx = mn + start_idx
                                        Exit While
                                    End If
                                End While

                                mn = 0
                                While start_idx1 + mn <= TBLMOB.DefaultView.Count - 1
                                    rowMOB = TBLMOB.DefaultView(mn + start_idx1).Row
                                    If rowMOB.Primary_Product < vrowProduct.Primary Then
                                        mn += 1
                                    ElseIf rowMOB.Primary_Product = vrowProduct.Primary Then
                                        TBLMOB.Rows.Remove(rowMOB)
                                    ElseIf rowMOB.Primary_Product > vrowProduct.Primary Then
                                        start_idx1 = mn + start_idx1
                                        Exit While
                                    End If
                                End While

                                mn = 0
                                While start_idx2 + mn <= TBLGIT.DefaultView.Count - 1
                                    rowGIT = TBLGIT.DefaultView(mn + start_idx2).Row
                                    If rowGIT.Primary_Product < vrowProduct.Primary Then
                                        mn += 1
                                    ElseIf rowGIT.Primary_Product = vrowProduct.Primary Then
                                        TBLGIT.Rows.Remove(rowGIT)
                                    ElseIf rowGIT.Primary_Product > vrowProduct.Primary Then
                                        start_idx2 = mn + start_idx2
                                        Exit While
                                    End If
                                End While

                                mn = 0
                                While start_idx3 + mn <= TBLPendingDO.DefaultView.Count - 1
                                    rowPendingDO = TBLPendingDO.DefaultView(mn + start_idx3).Row
                                    If rowPendingDO.Primary_Product < vrowProduct.Primary Then
                                        mn += 1
                                    ElseIf rowPendingDO.Primary_Product = vrowProduct.Primary Then
                                        TBLPendingDO.Rows.Remove(rowPendingDO)
                                    ElseIf rowPendingDO.Primary_Product > vrowProduct.Primary Then
                                        start_idx3 = mn + start_idx3
                                        Exit While
                                    End If
                                End While

                                mn = 0
                                While start_idx4 + mn <= TBLPOB.DefaultView.Count - 1
                                    rowPOB = TBLPOB.DefaultView(mn + start_idx4).Row
                                    If rowPOB.Primary_Product < vrowProduct.Primary Then
                                        mn += 1
                                    ElseIf rowPOB.Primary_Product = vrowProduct.Primary Then
                                        TBLPOB.Rows.Remove(rowPOB)
                                    ElseIf rowPOB.Primary_Product > vrowProduct.Primary Then
                                        start_idx4 = mn + start_idx4
                                        Exit While
                                    End If
                                End While

                                mn = 0
                                While start_idx5 + mn <= TBLSOB.DefaultView.Count - 1
                                    rowSOB = TBLSOB.DefaultView(mn + start_idx5).Row
                                    If rowSOB.Primary_Product < vrowProduct.Primary Then
                                        mn += 1
                                    ElseIf rowSOB.Primary_Product = vrowProduct.Primary Then
                                        TBLSOB.Rows.Remove(rowSOB)
                                    ElseIf rowSOB.Primary_Product > vrowProduct.Primary Then
                                        start_idx5 = mn + start_idx5
                                        Exit While
                                    End If
                                End While

                                mn = 0
                                While start_idx6 + mn <= TBLPrOB.DefaultView.Count - 1
                                    rowPrOB = TBLPrOB.DefaultView(mn + start_idx6).Row
                                    If rowPrOB.primary_code_no < vrowProduct.Primary Then
                                        mn += 1
                                    ElseIf rowPrOB.primary_code_no = vrowProduct.Primary Then
                                        TBLPrOB.Rows.Remove(rowPrOB)
                                    ElseIf rowPrOB.primary_code_no > vrowProduct.Primary Then
                                        start_idx6 = mn + start_idx6
                                        Exit While
                                    End If
                                End While

                                mn = 0
                                While start_idx7 + mn <= TBLPOItem.DefaultView.Count - 1
                                    rowPOItem = TBLPOItem.DefaultView(mn + start_idx7).Row
                                    If rowPOItem.Primary_Product < vrowProduct.Primary Then
                                        mn += 1
                                    ElseIf rowPOItem.Primary_Product = vrowProduct.Primary Then
                                        TBLPOItem.Rows.Remove(rowPOItem)
                                    ElseIf rowPOItem.Primary_Product > vrowProduct.Primary Then
                                        start_idx7 = mn + start_idx7
                                        Exit While
                                    End If
                                End While

                                mn = 0
                                While start_idx8 + mn <= TBLProductQty.DefaultView.Count - 1
                                    rowProductQty = TBLProductQty.DefaultView(mn + start_idx8).Row
                                    If rowProductQty.Primary_Product < vrowProduct.Primary Then
                                        mn += 1
                                    ElseIf rowProductQty.Primary_Product = vrowProduct.Primary Then
                                        TBLProductQty.Rows.Remove(rowProductQty)
                                    ElseIf rowProductQty.Primary_Product > vrowProduct.Primary Then
                                        start_idx8 = mn + start_idx8
                                        Exit While
                                    End If
                                End While

                                vTBLProduct.Rows.Remove(vrowProduct)
                            End While
                        End With
                    Next
                Else
                    TblDummy.Rows.Clear()

                    'pindahkan ke tabel dummy
                    For Each rowGroup In TBLGroup.Rows
                        ProductLevel = rowGroup.level
                        With rowGroup
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

                            vTBLProduct.DefaultView.RowFilter = Filter
                            start_idx = 0
                            start_idx1 = 0
                            start_idx2 = 0
                            start_idx3 = 0
                            start_idx4 = 0
                            start_idx5 = 0
                            start_idx6 = 0
                            start_idx7 = 0
                            start_idx8 = 0
                            While vTBLProduct.DefaultView.Count > 0
                                vrowProduct = vTBLProduct.DefaultView(0).Row

                                mn = 0
                                While start_idx + mn <= TBLRMReq.DefaultView.Count - 1
                                    rowRmReq = TBLRMReq.DefaultView(mn + start_idx).Row
                                    If rowRmReq.Primary_Product < vrowProduct.Primary Then
                                        mn += 1
                                    ElseIf rowRmReq.Primary_Product = vrowProduct.Primary Then
                                        rowRmReqDummy = TBLRMReqDummy.Newtbltrans_mo_item_rmRow
                                        For jk As Integer = 0 To TBLRMReq.Columns.Count - 1
                                            rowRmReqDummy.Item(jk) = rowRmReq.Item(jk)
                                        Next
                                        TBLRMReqDummy.Rows.Add(rowRmReqDummy)
                                        TBLRMReq.Rows.Remove(rowRmReq)
                                    ElseIf rowRmReq.Primary_Product > vrowProduct.Primary Then
                                        start_idx = mn + start_idx
                                        Exit While
                                    End If
                                End While

                                mn = 0
                                While start_idx1 + mn <= TBLMOB.DefaultView.Count - 1
                                    rowMOB = TBLMOB.DefaultView(mn + start_idx1).Row
                                    If rowMOB.Primary_Product < vrowProduct.Primary Then
                                        mn += 1
                                    ElseIf rowMOB.Primary_Product = vrowProduct.Primary Then
                                        rowMOBDummy = TBLMOBDummy.Newtbltrans_mo_itemRow
                                        For jk As Integer = 0 To TBLMOB.Columns.Count - 1
                                            rowMOBDummy.Item(jk) = rowMOB.Item(jk)
                                        Next
                                        TBLMOBDummy.Rows.Add(rowMOBDummy)
                                        TBLMOB.Rows.Remove(rowMOB)
                                    ElseIf rowMOB.Primary_Product > vrowProduct.Primary Then
                                        start_idx1 = mn + start_idx1
                                        Exit While
                                    End If
                                End While

                                mn = 0
                                While start_idx2 + mn <= TBLGIT.DefaultView.Count - 1
                                    rowGIT = TBLGIT.DefaultView(mn + start_idx2).Row
                                    If rowGIT.Primary_Product < vrowProduct.Primary Then
                                        mn += 1
                                    ElseIf rowGIT.Primary_Product = vrowProduct.Primary Then
                                        rowGITDummy = TBLGITDummy.Newtbltrans_packinglist_itemRow
                                        For jk As Integer = 0 To TBLGIT.Columns.Count - 1
                                            rowGITDummy.Item(jk) = rowGIT.Item(jk)
                                        Next
                                        TBLGITDummy.Rows.Add(rowGITDummy)
                                        TBLGIT.Rows.Remove(rowGIT)
                                    ElseIf rowGIT.Primary_Product > vrowProduct.Primary Then
                                        start_idx2 = mn + start_idx2
                                        Exit While
                                    End If
                                End While

                                mn = 0
                                While start_idx3 + mn <= TBLPendingDO.DefaultView.Count - 1
                                    rowPendingDO = TBLPendingDO.DefaultView(mn + start_idx3).Row
                                    If rowPendingDO.Primary_Product < vrowProduct.Primary Then
                                        mn += 1
                                    ElseIf rowPendingDO.Primary_Product = vrowProduct.Primary Then
                                        rowPendingDODummy = TBLPendingDODummy.NewDataTable2Row
                                        For jk As Integer = 0 To TBLPendingDO.Columns.Count - 1
                                            rowPendingDODummy.Item(jk) = rowPendingDO.Item(jk)
                                        Next
                                        TBLPendingDODummy.Rows.Add(rowPendingDODummy)
                                        TBLPendingDO.Rows.Remove(rowPendingDO)
                                    ElseIf rowPendingDO.Primary_Product > vrowProduct.Primary Then
                                        start_idx3 = mn + start_idx3
                                        Exit While
                                    End If
                                End While

                                mn = 0
                                While start_idx4 + mn <= TBLPOB.DefaultView.Count - 1
                                    rowPOB = TBLPOB.DefaultView(mn + start_idx4).Row
                                    If rowPOB.Primary_Product < vrowProduct.Primary Then
                                        mn += 1
                                    ElseIf rowPOB.Primary_Product = vrowProduct.Primary Then
                                        rowPOBDummy = TBLPOBDummy.NewDataTable3Row
                                        For jk As Integer = 0 To TBLPOB.Columns.Count - 1
                                            rowPOBDummy.Item(jk) = rowPOB.Item(jk)
                                        Next
                                        TBLPOBDummy.Rows.Add(rowPOBDummy)
                                        TBLPOB.Rows.Remove(rowPOB)
                                    ElseIf rowPOB.Primary_Product > vrowProduct.Primary Then
                                        start_idx4 = mn + start_idx4
                                        Exit While
                                    End If
                                End While

                                mn = 0
                                While start_idx5 + mn <= TBLSOB.DefaultView.Count - 1
                                    rowSOB = TBLSOB.DefaultView(mn + start_idx5).Row
                                    If rowSOB.Primary_Product < vrowProduct.Primary Then
                                        mn += 1
                                    ElseIf rowSOB.Primary_Product = vrowProduct.Primary Then
                                        rowSOBDummy = TBLSOBDummy.NewDataTable4Row
                                        For jk As Integer = 0 To TBLSOB.Columns.Count - 1
                                            rowSOBDummy.Item(jk) = rowSOB.Item(jk)
                                        Next
                                        TBLSOBDummy.Rows.Add(rowSOBDummy)
                                        TBLSOB.Rows.Remove(rowSOB)
                                    ElseIf rowSOB.Primary_Product > vrowProduct.Primary Then
                                        start_idx5 = mn + start_idx5
                                        Exit While
                                    End If
                                End While

                                mn = 0
                                While start_idx6 + mn <= TBLPrOB.DefaultView.Count - 1
                                    rowPrOB = TBLPrOB.DefaultView(mn + start_idx6).Row
                                    If rowPrOB.primary_code_no < vrowProduct.Primary Then
                                        mn += 1
                                    ElseIf rowPrOB.primary_code_no = vrowProduct.Primary Then
                                        rowPrOBDummy = TBLPrOBDummy.Newtbltrans_material_planning_itempurch_req1Row
                                        For jk As Integer = 0 To TBLPrOB.Columns.Count - 1
                                            rowPrOBDummy.Item(jk) = rowPrOB.Item(jk)
                                        Next
                                        TBLPrOBDummy.Rows.Add(rowPrOBDummy)
                                        TBLPrOB.Rows.Remove(rowPrOB)
                                    ElseIf rowPrOB.primary_code_no > vrowProduct.Primary Then
                                        start_idx6 = mn + start_idx6
                                        Exit While
                                    End If
                                End While

                                mn = 0
                                While start_idx7 + mn <= TBLPOItem.DefaultView.Count - 1
                                    rowPOItem = TBLPOItem.DefaultView(mn + start_idx7).Row
                                    If rowPOItem.Primary_Product < vrowProduct.Primary Then
                                        mn += 1
                                    ElseIf rowPOItem.Primary_Product = vrowProduct.Primary Then
                                        rowPOItemDummy = TBLPOItemDummy.Newtbltrans_po_itemRow
                                        For jk As Integer = 0 To TBLPOItem.Columns.Count - 1
                                            rowPOItemDummy.Item(jk) = rowPOItem.Item(jk)
                                        Next
                                        TBLPOItemDummy.Rows.Add(rowPOItemDummy)
                                        TBLPOItem.Rows.Remove(rowPOItem)
                                    ElseIf rowPOItem.Primary_Product > vrowProduct.Primary Then
                                        start_idx7 = mn + start_idx7
                                        Exit While
                                    End If
                                End While

                                mn = 0
                                While start_idx8 + mn <= TBLProductQty.DefaultView.Count - 1
                                    rowProductQty = TBLProductQty.DefaultView(mn + start_idx8).Row
                                    If rowProductQty.Primary_Product < vrowProduct.Primary Then
                                        mn += 1
                                    ElseIf rowProductQty.Primary_Product = vrowProduct.Primary Then
                                        rowProductQtyDummy = TBLProductQtyDummy.Newtblmaster_productqtyRow
                                        For jk As Integer = 0 To TBLProductQty.Columns.Count - 1
                                            rowProductQtyDummy.Item(jk) = rowProductQty.Item(jk)
                                        Next
                                        TBLProductQtyDummy.Rows.Add(rowProductQtyDummy)
                                        TBLProductQty.Rows.Remove(rowProductQty)
                                    ElseIf rowProductQty.Primary_Product > vrowProduct.Primary Then
                                        start_idx8 = mn + start_idx8
                                        Exit While
                                    End If
                                End While

                                rowDummy = TblDummy.NewDataTable1Row
                                For hj As Integer = 0 To TblDummy.Columns.Count - 1
                                    rowDummy.Item(hj) = vrowProduct.Item(hj)
                                Next
                                TblDummy.Rows.Add(rowDummy)

                                vTBLProduct.Rows.Remove(vrowProduct)
                            End While
                        End With
                    Next
                    'kosongkan tabel asal
                    vTBLProduct.Rows.Clear()
                    TBLRMReq.Rows.Clear()
                    TBLMOB.Rows.Clear()
                    TBLGIT.Rows.Clear()
                    TBLPendingDO.Rows.Clear()
                    TBLPOB.Rows.Clear()
                    TBLSOB.Rows.Clear()
                    TBLPrOB.Rows.Clear()
                    TBLPOItem.Rows.Clear()
                    TBLProductQty.Rows.Clear()

                    'pindahkan kembali ke tabel asal
                    While TblDummy.Rows.Count > 0
                        rowDummy = TblDummy.Rows(0)
                        vrowProduct = vTBLProduct.NewDataTable1Row
                        For hj As Integer = 0 To TblDummy.Columns.Count - 1
                            vrowProduct.Item(hj) = rowDummy.Item(hj)
                        Next
                        vTBLProduct.Rows.Add(vrowProduct)
                        TblDummy.Rows.RemoveAt(0)
                    End While

                    While TBLRMReqDummy.Rows.Count > 0
                        rowRmReqDummy = TBLRMReqDummy.Rows(0)
                        rowRmReq = TBLRMReq.Newtbltrans_mo_item_rmRow
                        For ij As Integer = 0 To TBLRMReq.Columns.Count - 1
                            rowRmReq.Item(ij) = rowRmReqDummy.Item(ij)
                        Next
                        TBLRMReq.Rows.Add(rowRmReq)
                        TBLRMReqDummy.Rows.RemoveAt(0)
                    End While

                    While TBLMOBDummy.Rows.Count > 0
                        rowMOBDummy = TBLMOBDummy.Rows(0)
                        rowMOB = TBLMOB.Newtbltrans_mo_itemRow
                        For ij As Integer = 0 To TBLMOB.Columns.Count - 1
                            rowMOB.Item(ij) = rowMOBDummy.Item(ij)
                        Next
                        TBLMOB.Rows.Add(rowMOB)
                        TBLMOBDummy.Rows.RemoveAt(0)
                    End While

                    While TBLGITDummy.Rows.Count > 0
                        rowGITDummy = TBLGITDummy.Rows(0)
                        rowGIT = TBLGIT.Newtbltrans_packinglist_itemRow
                        For ij As Integer = 0 To TBLGIT.Columns.Count - 1
                            rowGIT.Item(ij) = rowGITDummy.Item(ij)
                        Next
                        TBLGIT.Rows.Add(rowGIT)
                        TBLGITDummy.Rows.RemoveAt(0)
                    End While

                    While TBLPendingDODummy.Rows.Count > 0
                        rowPendingDODummy = TBLPendingDODummy.Rows(0)
                        rowPendingDO = TBLPendingDO.NewDataTable2Row
                        For ij As Integer = 0 To TBLPendingDO.Columns.Count - 1
                            rowPendingDO.Item(ij) = rowPendingDODummy.Item(ij)
                        Next
                        TBLPendingDO.Rows.Add(rowPendingDO)
                        TBLPendingDODummy.Rows.RemoveAt(0)
                    End While

                    While TBLPOBDummy.Rows.Count > 0
                        rowPOBDummy = TBLPOBDummy.Rows(0)
                        rowPOB = TBLPOB.NewDataTable3Row
                        For ij As Integer = 0 To TBLPOB.Columns.Count - 1
                            rowPOB.Item(ij) = rowPOBDummy.Item(ij)
                        Next
                        TBLPOB.Rows.Add(rowPOB)
                        TBLPOBDummy.Rows.RemoveAt(0)
                    End While

                    While TBLSOBDummy.Rows.Count > 0
                        rowSOBDummy = TBLSOBDummy.Rows(0)
                        rowSOB = TBLSOB.NewDataTable4Row
                        For ij As Integer = 0 To TBLSOB.Columns.Count - 1
                            rowSOB.Item(ij) = rowSOBDummy.Item(ij)
                        Next
                        TBLSOB.Rows.Add(rowSOB)
                        TBLSOBDummy.Rows.RemoveAt(0)
                    End While

                    While TBLPrOBDummy.Rows.Count > 0
                        rowPrOBDummy = TBLPrOBDummy.Rows(0)
                        rowPrOB = TBLPrOB.Newtbltrans_material_planning_itempurch_req1Row
                        For ij As Integer = 0 To TBLPrOB.Columns.Count - 1
                            rowPrOB.Item(ij) = rowPrOBDummy.Item(ij)
                        Next
                        TBLPrOB.Rows.Add(rowPrOB)
                        TBLPrOBDummy.Rows.RemoveAt(0)
                    End While

                    While TBLPOItemDummy.Rows.Count > 0
                        rowPOItemDummy = TBLPOItemDummy.Rows(0)
                        rowPOItem = TBLPOItem.Newtbltrans_po_itemRow
                        For ij As Integer = 0 To TBLPOItem.Columns.Count - 1
                            rowPOItem.Item(ij) = rowPOItemDummy.Item(ij)
                        Next
                        TBLPOItem.Rows.Add(rowPOItem)
                        TBLPOItemDummy.Rows.RemoveAt(0)
                    End While

                    While TBLProductQtyDummy.Rows.Count > 0
                        rowProductQtyDummy = TBLProductQtyDummy.Rows(0)
                        rowProductQty = TBLProductQty.Newtblmaster_productqtyRow
                        For ij As Integer = 0 To TBLProductQty.Columns.Count - 1
                            rowProductQty.Item(ij) = rowProductQtyDummy.Item(ij)
                        Next
                        TBLProductQty.Rows.Add(rowProductQty)
                        TBLProductQtyDummy.Rows.RemoveAt(0)
                    End While
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAGroup.Dispose()
            TBLGroup = Nothing
            TblDummy.Dispose()
            vTBLProduct.DefaultView.Sort = DummySort
            TBLRMReqDummy.Dispose()
            TBLMOBDummy.Dispose()
            TBLGITDummy.Dispose()
            TBLPendingDODummy.Dispose()
            TBLPOBDummy.Dispose()
            TBLSOBDummy.Dispose()
            TBLPrOBDummy.Dispose()
            TBLPOItemDummy.Dispose()
            TBLProductQtyDummy.Dispose()
        End Try
    End Sub
End Class
