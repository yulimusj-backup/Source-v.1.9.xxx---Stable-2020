Option Explicit On
Option Strict Off

Imports Core_Module

Partial Public Class FrmTrans_MaterialPlanningDashboard
    Dim baru As Boolean
    Dim ObjMP As New ClsTransaksiMaterialPlanning
    Dim ObjMPItem As New ClsTransaksiMaterialPlanning_item
    Dim ObjLicense As New ClsLicense
    Dim ObjNumbering As New Core_Module.ClsMasterTransactionNumbering
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
    Dim vRowItemFollowUp As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_followupRow
    Dim objUnit As New Core_Module.ClsMasterUnit
    Dim SaveRequired As Boolean = False

    Dim ObjSetting As New Core_Module.ClsTransaksiSetting
    Dim ObjMO As New ClsTrans_ManufacturingOrder
    Dim FrmCek As New FrmTrans_MaterialPlanningDashboard_DetailPO
    Dim ObjSupplier As New Core_Module.clsMasterSupplier
    Dim ObjCurrency As New Core_Module.clsMasterCurrency
    Dim objPG As New Core_Module.ClsMasterProductGroup
    Dim objCategory As New Core_Module.ClsMasterCategory
    Dim ObjPO As New ClsTransaksiPO
    Dim ObjSO As New ClsTransaksiSalesOrder
    Dim ObjPurcReq As New ClsTransaksiPR
    Dim ObjStockInPurchase As New clsTransaksiStockInOut
    Dim ObjCls As New ClsTransaksiMaterialPlanning_Calc
    Dim TBLHistoricalPrice As New DseTransaksiPO.DataTable9DataTable

    Dim DTAProduct_MU As New dseMasterProductTableAdapters.tblmaster_product5TableAdapter
    Dim TBLProduct_MU As New dseMasterProduct.tblmaster_product5DataTable
    Dim dtaMultiUnit As New dseMasterProductTableAdapters.tblmaster_product_item_multiunit2TableAdapter
    Dim tblMultiUnit As New dseMasterProduct.tblmaster_product_item_multiunit2DataTable
    Dim rowMultiUnit As dseMasterProduct.tblmaster_product_item_multiunit2Row
    Dim vHitungQtyMultiUnit As Boolean = True
    Dim BufferCekBox As String = ""

    Public Enum EnumTipeDocReq As Integer
        so_
        do_
        rm_req
        pr_ob
        free
    End Enum

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboMP.Controls, Me.TxtMP_Primary.BackColor, False)
            ModCekData.KendalikanControl(Me.gboMP.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer5.Panel1.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer5.Panel2.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer4.Panel1.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.SplitContainer4.Panel2.Controls, True, True, True, True, True, True, True, True)
            ModCekData.KendalikanControl(Me.TabPage1.Controls, True, False, True, False, True, False, True, False)
            Me.CancelButton = Me.btnMP_Cancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            Me.dgvMP.ReadOnly = False

            Me.ObjCD.SetBackColor(Me.SplitContainer4.Panel1.Controls, Me.TxtMP_Primary.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer4.Panel2.Controls, Me.TxtMP_Primary.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer5.Panel1.Controls, Me.TxtMP_Primary.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer5.Panel2.Controls, Me.TxtMP_Primary.BackColor, False)

            If Me.ObjNumbering.Apakah_Autonumber(ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._MP) = True Then
                Me.TxtMP_DocNo.ReadOnly = True
                Me.TxtMP_DocNo.BackColor = Me.TxtReadOnly.BackColor
                Me.TxtMP_DocNo.TabStop = False
                Me.dtpMP.Focus()
            Else
                Me.TxtMP_DocNo.ReadOnly = False
                Me.TxtMP_DocNo.BackColor = Me.TxtMP_Primary.BackColor
                Me.TxtMP_DocNo.TabStop = True
                Me.TxtMP_DocNo.Focus()
            End If

            Me.llbMP_Calculate.Enabled = True
            Me.llbMP_CreatePO.Enabled = True
            Me.llbMP_Ratio.Enabled = True
            Me.LlbMP_UnitPrice.Enabled = True
            Me.llbMP_AddToPO.Enabled = True
            Me.llbMP_CekAll.Enabled = True
            Me.llbMP_CopyMP.Enabled = True

            Me.chkClose.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            ModCekData.KendalikanControl(Me.gboMP.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer4.Panel1.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer4.Panel2.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer5.Panel1.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.SplitContainer5.Panel2.Controls, True, False, True, False, True, False, True, False)
            ModCekData.KendalikanControl(Me.TabPage1.Controls, True, True, True, True, True, True, True, True)
            Me.CancelButton = Me.btnClose

            Me.dgvMP.AllowUserToDeleteRows = False
            Me.dgvMP.ReadOnly = True

            Me.ObjCD.SetBackColor(Me.gboMP.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer4.Panel1.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer4.Panel2.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer5.Panel1.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer5.Panel2.Controls, Me.BackColor, False)

            Me.llbMP_Calculate.Enabled = False
            Me.llbMP_CreatePO.Enabled = False
            Me.llbMP_Ratio.Enabled = False
            Me.LlbMP_UnitPrice.Enabled = False
            Me.llbMP_AddToPO.Enabled = False
            Me.llbMP_CekAll.Enabled = False
            Me.llbMP_CopyMP.Enabled = False
            Me.chkMP_RequiredOnly.Enabled = True
            Me.dtpMP_Filter.Enabled = True
            Me.chkClose.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetDPM()
        Try
            Dim JmlDPM As String = Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.MaterialPlanning, _
              ClsTransaksiSetting.EnumNamaParameter.MaterialPlanning_JmlHariKerjaPerBulan)
            If JmlDPM = "" Then
                Me.ObjSetting.SetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.MaterialPlanning, _
                            ClsTransaksiSetting.EnumNamaParameter.MaterialPlanning_JmlHariKerjaPerBulan, 30)
                Me.TxtMP_DaysPerMonth.Text = "30"
            Else
                Me.TxtMP_DaysPerMonth.Text = JmlDPM
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            ModCekData.KosongkanControl(Me.gboMP.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.SplitContainer4.Panel1.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.SplitContainer4.Panel2.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.SplitContainer5.Panel1.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.SplitContainer5.Panel2.Controls, True, True, True)
            Me.TxtMP_Primary.Text = "-1"
            Me.TxtMP_PrimaryCurrency.Text = "-1"
            Me.TxtMP_PrimaryOrderUnit.Text = "-1"
            Me.TxtMP_PrimarySupplier.Text = "-1"
            Me.SetDPM()
            Me.tslMP_CreatedBy.Text = "Created By: "
            Me.tslMP_CreatedDate.Text = "Created Date: "
            Me.tslMP_ModifiedBy.Text = "Modified By: "
            Me.tslMP_ModifiedDate.Text = "Modified Date: "
            Me.cboMPOrderStrategy.SelectedIndex = 0
            Me.TxtMP_QtyGIT.Text = 0
            Me.TxtMP_QtyInquiry.Text = 0
            Me.TxtMP_QtyOptimumStock.Text = 0
            Me.TxtMP_QtyOnHand.Text = 0
            Me.TxtMP_QtyOrder.Text = 0
            Me.TxtMP_QtyFree.Text = 0
            Me.TxtMP_QtyPendingDO.Text = 0
            Me.TxtMP_QtyPOB.Text = 0
            Me.TxtMP_QtyRequired.Text = 0
            Me.TxtMP_QtyRequired2.Text = 0
            Me.TxtMP_QtySafetyStock.Text = 0
            Me.TxtMP_QtySOB.Text = 0
            Me.TxtMP_QtySOBSpare.Text = 0
            Me.TxtMP_QtyPrOB.Text = 0
            Me.TxtMP_QtyPrOBSpare.Text = 0
            Me.TxtMP_QtyRMReq.Text = 0
            Me.TxtMP_QtyRMReqSpare.Text = 0
            Me.TxtMP_QtyMOB.Text = 0
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

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim TanggalClosingMax As Date = #1/1/1950#
        Dim DTA As New DseTransaksiDOTableAdapters.tbltrans_do3TableAdapter

        Try
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Me.dtpMP.Value Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            'cek data yang harus diisi
            KoleksiNamaData.Add("document no")
            KoleksiTextBox.Add(Me.TxtMP_DocNo)
            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            'minimal harus punya 1 data item
            If Me.vDT.Rows.Count = 0 Then
                MsgBox("Please enter at least 1 item.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'cek berkaitan dengan perubahan tanggal
            If Me.ObjNumbering.Apakah_Autonumber(ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._MP) = True Then
                If Me.baru = False Then
                    If Me.ObjMP.TanggalTransaksi.Month = Me.dtpMP.Value.Month And _
                                       Me.ObjMP.TanggalTransaksi.Year = Me.dtpMP.Value.Year Then
                        Me.TxtMP_DocNo.Text = Me.ObjMP.NoTransaksi
                    Else
                        If MsgBox("You just changed transaction date, this will change your transaction number and there's a possibility your transaction no cannot be change back. Continue and accept this change?", _
                        MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
                            Me.dtpMP.Value = Me.ObjMP.TanggalTransaksi
                            Me.TxtMP_DocNo.Text = Me.ObjMP.NoTransaksi
                            Return False
                        End If
                    End If
                End If
            End If

            'kalo sampe cek berkaitan dengan timestamp apakah masih benar
            If Me.baru = False Then
                If Me.ObjMP.TimeStamp <> Me.ObjMP.GetTimestamp(Me.ObjMP.Primary) Then
                    If MsgBox("Other user have modified this data, reload latest data?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.ObjMP.CariDetail(Me.ObjMP.Primary, Me.vDT, vDT_ItemGIT, vDT_ItemDO, vDT_ItemPOB, _
                        vDT_ItemSOB, vDT_ItemPrOB, vDT_ItemSettingLocation, vDT_ItemQty, vDT_ItemMOB, vDT_ItemRMReq, Me.vDT_ItemFollowUp, Me.vDT_ItemSOBDelivery, _
                        vDT_ItemPrOBDelivery)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            If Me.baru Then
                If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_MaterialPlanningDashboard_MP_Track") Then
                    If Me.ObjMP.GetOpenMPCount(Me.UserCode) > 0 Then
                        Dim docNo As String = Me.ObjMP.GetOpenMPDocNo(Me.UserCode)
                        MsgBox("There is already opened MP: " & docNo & ", you are only allowed 1 MP at a time.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            End If

            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            KoleksiNamaData = Nothing
            KoleksiTextBox = Nothing
            DTA.Dispose()
        End Try
    End Function

    Private Sub SalinKeObjek()
        Try
            'salin data header
            With Me.ObjMP
                .NoTransaksi = Me.TxtMP_DocNo.Text
                .TanggalTransaksi = Me.dtpMP.Value
                .Remarks = Me.TxtMP_Remarks.Text
                .RefDate = Me.dtpMP_RefDate.Value
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadDGV()
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
            If Me.TxtFilter_PrimaryPurchReq.Text <> "-1" And Me.TxtFilter_PrimaryPurchReq.Text <> "" Then
                DTAPurchReq.FillBy_Primary(TBLPurchReq, Me.TxtFilter_PrimaryPurchReq.Text)
            End If

            If Me.SplitContainer1.Panel1Collapsed = True Then
                Filter = ""
            Else
                Me.TxtFilter_CodeNo.Text = Me.TxtFilter_CodeNo.Text.Trim
                If Me.TxtFilter_CodeNo.Text <> "" Then
                    Filter = "[Product Code] LIKE '" + Me.ObjCD.FormatSpecialChar_DataFilter(Me.TxtFilter_CodeNo.Text, True) + "*'"
                End If
                'harus pilih salah satu, goods atau service
                If Me.chkFilter_Goods.Checked = False And Me.chkFilter_Service.Checked = False Then
                    Me.chkFilter_Goods.Checked = True
                End If
                If Me.chkFilter_Goods.Checked = False Or Me.chkFilter_Service.Checked = False Then
                    If Me.chkFilter_Goods.Checked = False Then
                        'service only
                        If Filter = "" Then
                            Filter = "[product type] = 1"
                        Else
                            Filter = Filter + " AND [product type] = 1"
                        End If
                    Else
                        'goods only
                        If Filter = "" Then
                            Filter = "[product type] = 0"
                        Else
                            Filter = Filter + " AND [product type] = 0"
                        End If
                    End If
                End If

                'If Me.chkFilter_PrOBItem.Checked = True Then
                '    If Filter = "" Then
                '        Filter = "[qty prob] > 0"
                '    Else
                '        Filter = Filter + " AND [qty prob] > 0"
                '    End If
                'End If
                'If Me.chkFilter_SOBItem.Checked = True Then
                '    If Filter = "" Then
                '        Filter = "[qty sob] > 0"
                '    Else
                '        Filter = Filter + " AND [qty sob] > 0"
                '    End If
                'End If
                If Me.chkFilter_FollowUpOnly.Checked = True Then
                    If Filter = "" Then
                        Filter = "[modify] > 0"
                    Else
                        Filter = Filter + " AND [modify] > 0"
                    End If
                End If

                Me.TxtFilter_Desc.Text = Me.TxtFilter_Desc.Text.Trim
                If Me.TxtFilter_Desc.Text <> "" Then
                    If Filter = "" Then
                        Filter = "[Description] like '" + Me.ObjCD.FormatSpecialChar_DataFilter(Me.TxtFilter_Desc.Text, True) + "*'"
                    Else
                        Filter = Filter + " AND [Description] like '" + Me.ObjCD.FormatSpecialChar_DataFilter(Me.TxtFilter_Desc.Text, True) + "*'"
                    End If
                End If

                If Me.TxtFilter_PreferedSupplier.Text <> "" Then
                    If Filter = "" Then
                        Filter = "[primary prefered supplier] = " + Me.TxtFilter_PrimaryPreferedSupplier.Text
                    Else
                        Filter = Filter + " AND [primary prefered supplier] = " + Me.TxtFilter_PrimaryPreferedSupplier.Text
                    End If
                End If

                If Me.TxtFilter_ProductGroupCode.Text <> "" Then
                    If Me.cboFilter_ProductGroupLevel.SelectedIndex = 0 Then
                        'product group
                        If Filter = "" Then
                            Filter = "[Primary Product Goup] = " + Me.TxtFilter_PrimaryProductGroup.Text
                        Else
                            Filter = Filter + " AND [Primary Product Goup] = " + Me.TxtFilter_PrimaryProductGroup.Text
                        End If
                    ElseIf Me.cboFilter_ProductGroupLevel.SelectedIndex = 1 Then
                        'brand
                        If Filter = "" Then
                            Filter = "[Primary Brand] = " + Me.TxtFilter_PrimaryProductGroup.Text
                        Else
                            Filter = Filter + " AND [Primary Brand] = " + Me.TxtFilter_PrimaryProductGroup.Text
                        End If
                    ElseIf Me.cboFilter_ProductGroupLevel.SelectedIndex = 2 Then
                        'sub band
                        If Filter = "" Then
                            Filter = "[Primary Sub Brand] = " + Me.TxtFilter_PrimaryProductGroup.Text
                        Else
                            Filter = Filter + " AND [Primary Sub Brand] = " + Me.TxtFilter_PrimaryProductGroup.Text
                        End If
                    ElseIf Me.cboFilter_ProductGroupLevel.SelectedIndex = 3 Then
                        'category
                        If Filter = "" Then
                            Filter = "[Primary Category] = " + Me.TxtFilter_PrimaryProductGroup.Text
                        Else
                            Filter = Filter + " AND [Primary Category] = " + Me.TxtFilter_PrimaryProductGroup.Text
                        End If
                    End If
                End If

                'rnew
                If Me.CboFilter_ProcurementStrategy.Text <> "Both" Then
                    If Me.CboFilter_ProcurementStrategy.Text = "Build to stock" Then
                        If Filter = "" Then
                            Filter = "[procurement strategy] = " + CStr(Core_Module.clsMasterProduct.EnumProcurementStrategy.Build_To_Stock)
                        Else
                            Filter = Filter + " AND [procurement strategy] = " + CStr(Core_Module.clsMasterProduct.EnumProcurementStrategy.Build_To_Stock)
                        End If
                    ElseIf Me.CboFilter_ProcurementStrategy.Text = "Build to order" Then
                        If Filter = "" Then
                            Filter = "[procurement strategy] = " + CStr(Core_Module.clsMasterProduct.EnumProcurementStrategy.Build_To_Order)
                        Else
                            Filter = Filter + " AND [procurement strategy] = " + CStr(Core_Module.clsMasterProduct.EnumProcurementStrategy.Build_To_Order)
                        End If
                    End If
                End If
            End If
            If Me.chkFilter_PrOBLargerThanPOB.Checked = True Then
                If Filter = "" Then
                    Filter = "[qty prob] > [qty pob]"
                Else
                    Filter = Filter + " AND [qty prob] > [qty pob]"
                End If
            End If

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
                    If Me.chkFilter_SOBItem.Checked = True And Me.chkFilter_PrOBItem.Checked = True Then
                        'tidak ada filtering pd bagian ini
                    ElseIf Me.chkFilter_SOBItem.Checked = True And Me.chkFilter_PrOBItem.Checked = False Then
                        If row.qty_sob = 0 Then
                            Masuk = False
                        End If
                    ElseIf Me.chkFilter_SOBItem.Checked = False And Me.chkFilter_PrOBItem.Checked = True Then
                        If row.qty_prob = 0 Then
                            Masuk = False
                        End If
                    ElseIf Me.chkFilter_SOBItem.Checked = False And Me.chkFilter_PrOBItem.Checked = False Then
                        Masuk = False
                    End If
                End If
                'If Masuk = True Then
                '    If Me.chkFilter_PrOBItem.Checked = True And row.qty_prob <= 0 Then
                '        Masuk = False
                '    End If
                'End If
                'If Masuk = True Then
                '    If Me.chkFilter_SOBItem.Checked = True And row.qty_sob <= 0 Then
                '        Masuk = False
                '    End If
                'End If

                'terkait dgn purchase request
                If Masuk = True Then
                    If Me.TxtFilter_PrimaryPurchReq.Text <> "" And Me.TxtFilter_PrimaryPurchReq.Text <> "-1" Then
                        TBLPurchReq.DefaultView.RowFilter = "[primary code no] = " + row.primary_code_no.ToString
                        If TBLPurchReq.DefaultView.Count = 0 Then
                            Masuk = False
                        End If
                    End If
                End If

                'terkait dengan primary mo-> kalo tidak ada item rm req dan item mob terkait mo tersebut, maka code no tersebut tidak masuk
                If Masuk = True Then
                    If Me.TxtFilter_PrimaryMO.Text <> "" And Me.TxtFilter_PrimaryMO.Text <> "-1" Then
                        vDT_ItemMOB.DefaultView.RowFilter = "[primary product] = " + row.primary_code_no.ToString + " AND [primary mo rec] = " + Me.TxtFilter_PrimaryMO.Text
                        If vDT_ItemMOB.DefaultView.Count = 0 Then
                            vDT_ItemRMReq.DefaultView.RowFilter = "[primary product] = " + row.primary_code_no.ToString + " AND [primary mo rec] = " + Me.TxtFilter_PrimaryMO.Text
                            If vDT_ItemRMReq.DefaultView.Count = 0 Then
                                Masuk = False
                                'Else
                                '    MsgBox("")
                            End If
                            'Else
                            '    MsgBox("")
                        End If
                    End If
                End If

                'terkait dengan tanggal mo
                If Masuk = True Then
                    If Me.chkMODateFilter.Checked = True Then
                        vDT_ItemMOB.DefaultView.RowFilter = "[primary product] = " + row.primary_code_no.ToString + " AND [mo date rec] <= #" + Format(Me.DtpFilter_MO_To.Value, "MM/dd/yyyy") + "#" + _
                        " AND [mo date rec] >= #" + Format(Me.DtpFilter_MO_From.Value, "MM/dd/yyyy") + "#"

                        If vDT_ItemMOB.DefaultView.Count = 0 Then
                            vDT_ItemRMReq.DefaultView.RowFilter = "[primary product] = " + row.primary_code_no.ToString + " AND [mo date rec] <= #" + Format(Me.DtpFilter_MO_To.Value, "MM/dd/yyyy") + "#" + _
                        " AND [mo date rec] >= #" + Format(Me.DtpFilter_MO_From.Value, "MM/dd/yyyy") + "#"
                            If vDT_ItemRMReq.DefaultView.Count = 0 Then
                                Masuk = False
                            End If
                        End If
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
                        .Cells(Me.ColMP_Disc.Index).Value = ObjCD.DigitGrouping(row.discount, 2)
                        .Cells(Me.ColMP_NettPrice.Index).Value = ObjCD.DigitGrouping(row.nett_price, row.def_prec_currency)
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
                        .Cells(Me.ColMP_QtyOrder.Index).Value = Me.ObjCD.DigitGrouping(row.qty_order, row.def_prec_order_unit)
                        .Cells(Me.ColMP_QtyOrderMaximum.Index).Value = row.qty_order_maksimum
                        .Cells(Me.ColMP_QtyOrderMinimum.Index).Value = row.qty_order_minimum
                        .Cells(Me.ColMP_QtyOrderOptimum.Index).Value = row.qty_order_optimum
                        .Cells(Me.ColMP_QtyPendingDO.Index).Value = row.qty_pending_do
                        .Cells(Me.ColMP_QtyPOB.Index).Value = row.qty_pob
                        .Cells(Me.ColMP_QtyRequired.Index).Value = Me.ObjCD.DigitGrouping(row.qty_required, row.default_prec_unit)
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

                        If Me.TxtMP_Remarks.Enabled = True Then
                            If Me.TxtFilter_PrimaryPurchReq.Text <> "" And Me.TxtFilter_PrimaryPurchReq.Text <> "-1" Then
                                Me.Buffer = .Cells(Me.ColMP_QtyOrder.Index).Value
                                .Cells(Me.ColMP_QtyOrder.Index).Value = Me.GetQty(TBLPurchReq, row.primary_code_no)
                                Me.EndEdit_Qty(Me, rowDGV.Index, Me.ColMP_QtyOrder.Index)
                            End If
                        End If

                        Me.vDT_ItemFollowUp.DefaultView.RowFilter = "[primary transaksi] = " & row.primary_transaksi.ToString & " and [primary product] = " & row.primary_code_no.ToString
                        If Me.vDT_ItemFollowUp.DefaultView.Count = 0 Then
                            .Cells(Me.ColMP_DetailFollowUp.Index).Value = "-"
                        ElseIf Me.vDT_ItemFollowUp.DefaultView.Count = 1 Then
                            Me.vRowItemFollowUp = Me.vDT_ItemFollowUp.DefaultView(0).Row
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
            Else
                Me.SetTxtPurchaseDetail()
            End If

            SetStatusCek()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAPurchReq.Dispose()
            TBLPurchReq.Dispose()
        End Try
    End Sub

    Private Sub llbMP_CekAll_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbMP_CekAll.LinkClicked
        Dim rowDGV As DataGridViewRow
        Dim RefVal As Boolean = False
        Dim row As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item1Row
        Dim Dummy As String = ""
        Dim ModifiInt As Integer = 0

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Dummy = Me.vDT.DefaultView.RowFilter

            If Me.llbMP_CekAll.Text = "Check All" Then
                RefVal = True
                Me.llbMP_CekAll.Text = "Uncheck All"
                ModifiInt = 1
            Else
                RefVal = False
                Me.llbMP_CekAll.Text = "Check All"
                ModifiInt = 0
            End If

            Me.dgvMP.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For Each rowDGV In Me.dgvMP.Rows
                If RefVal = True Then
                    If CDec(rowDGV.Cells(Me.ColMP_QtyOrder.Index).Value) > 0 Then
                        rowDGV.Cells(Me.ColMP_FollowUp.Index).Value = RefVal
                        Me.vDT.DefaultView.RowFilter = "[primary code no] = " + rowDGV.Cells(Me.ColMP_PrimaryCodeNo.Index).Value.ToString
                        row = Me.vDT.DefaultView(0).Row
                        row.modify = ModifiInt
                    End If
                Else
                    rowDGV.Cells(Me.ColMP_FollowUp.Index).Value = RefVal
                    Me.vDT.DefaultView.RowFilter = "[primary code no] = " + rowDGV.Cells(Me.ColMP_PrimaryCodeNo.Index).Value.ToString
                    row = Me.vDT.DefaultView(0).Row
                    row.modify = ModifiInt
                End If
            Next
            Me.dgvMP.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            Me.vDT.DefaultView.RowFilter = Dummy

            Me.llbMP_CekAll.Left = Me.llbMP_AddToPO.Left - 5 - Me.llbMP_CekAll.Width
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetStatusCek()
        Dim rowDGV As DataGridViewRow
        Dim CekSemua As Boolean = True

        Try
            If Me.dgvMP.RowCount = 0 Then
                Me.llbMP_CekAll.Visible = False
            Else
                Me.llbMP_CekAll.Visible = True
                For Each rowDGV In Me.dgvMP.Rows
                    If rowDGV.Cells(Me.ColMP_FollowUp.Index).Value = Boolean.FalseString Then
                        CekSemua = False
                    End If

                    If CekSemua = True Then
                        Me.llbMP_CekAll.Text = "Uncheck All"
                    Else
                        Me.llbMP_CekAll.Text = "Check All"
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
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

                Me.tslMP_CreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslMP_CreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslMP_ModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslMP_ModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
            Me.chkMP_RequiredOnly.Checked = True
            Me.dtpMP_Filter.Value = Me.dtpMP_RefDate.Value
            Me.LoadDGV()
            Me.RefreshFollowUpControl()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnSimpan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnMP_OK.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.ApakahDataBenar() = True Then
                Me.SalinKeObjek()
                If baru = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjMP.SimpanBaru(Me.vDT, Me.UserCode, Me, Me.vDT_ItemGIT, Me.vDT_ItemDO, vDT_ItemPOB, vDT_ItemSOB, vDT_ItemPrOB, _
                        vDT_ItemSettingLocation, vDT_ItemQty, vDT_ItemMOB, vDT_ItemRMReq, Me.vDT_ItemFollowUp, Me.vDT_ItemSOBDelivery, _
                        vDT_ItemPrOBDelivery) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SaveRequired = False
                            Me.ObjMP.CariDetail(Me.ObjMP.Primary, Me.vDT, Me.vDT_ItemGIT, Me.vDT_ItemDO, vDT_ItemPOB, vDT_ItemSOB, vDT_ItemPrOB, _
                            vDT_ItemSettingLocation, vDT_ItemQty, vDT_ItemMOB, vDT_ItemRMReq, Me.vDT_ItemFollowUp, Me.vDT_ItemSOBDelivery, _
                            vDT_ItemPrOBDelivery)

                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnMP_Find.Focus()
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjMP.SimpanPerubahan(Me.vDT, Me.UserCode, Me, Me.vDT_ItemGIT, Me.vDT_ItemDO, vDT_ItemPOB, vDT_ItemSOB, _
                        vDT_ItemPrOB, vDT_ItemSettingLocation, vDT_ItemQty, vDT_ItemMOB, vDT_ItemRMReq, Me.vDT_ItemFollowUp, Me.vDT_ItemSOBDelivery, _
                        vDT_ItemPrOBDelivery) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SaveRequired = False
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnMP_Edit.Focus()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub btnMP_Cancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnMP_Cancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.SaveRequired = True Then
                MsgBox("You've just been following up one or more item, you have to save this transaction.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If Me.TxtMP_Remarks.Text = "" And Me.vDT.Rows.Count = 0 And Me.dtpMP.Value.Date = Today Then
                Me.NonAktifkanDetail()
                Me.btnMP_New.Focus()
                Me.KosongkanField()
            Else
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If baru = True Then
                        Me.KosongkanField()
                    Else
                        Me.ObjMP.CariDetail(Me.TxtMP_Primary.Text, Me.vDT, Me.vDT_ItemGIT, Me.vDT_ItemDO, vDT_ItemPOB, _
                        vDT_ItemSOB, vDT_ItemPrOB, vDT_ItemSettingLocation, vDT_ItemQty, vDT_ItemMOB, vDT_ItemRMReq, Me.vDT_ItemFollowUp, Me.vDT_ItemSOBDelivery, _
                        vDT_ItemPrOBDelivery)
                        Me.SalinKeField()
                    End If
                Else
                    Exit Sub
                End If
                Me.NonAktifkanDetail()
                Me.btnMP_New.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnMP_New_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnMP_New.Click
        Try
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Today Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            Me.baru = True
            Me.AktifkanDetail()
            Me.KosongkanField()
            'Me.dtpMP.Focus()
            Me.ObjMP.ResetObject(Me.vDT)
            Me.TxtMP_DocNo.Text = Me.ObjMP.GetNoBaru1(Me.dtpMP.Value)
            Me.tslMP_CreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslMP_CreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslMP_ModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslMP_ModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.SetStatusCek()
            RefreshFollowUpControl()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnMP_Find_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnMP_Find.Click
        Dim hj As New FrmTrans_MaterialPlanningDashboard_FindMP
        Try
            hj.BackColorAll = Me.BackColorAll
            hj.ShowDialog()

            If Not hj.PrimaryTransaksi = -1 Then
                Windows.Forms.Cursor.Current = Cursors.WaitCursor

                If Me.ObjMP.CariDetail(hj.PrimaryTransaksi, Me.vDT, Me.vDT_ItemGIT, Me.vDT_ItemDO, vDT_ItemPOB, vDT_ItemSOB, vDT_ItemPrOB, _
                vDT_ItemSettingLocation, vDT_ItemQty, vDT_ItemMOB, vDT_ItemRMReq, Me.vDT_ItemFollowUp, Me.vDT_ItemSOBDelivery, _
                vDT_ItemPrOBDelivery) = True Then
                    Me.SalinKeField()
                    Me.btnMP_Edit.Focus()
                Else
                    Me.btnMP_Find.Focus()
                End If
                Me.btnMP_Find.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub btnMP_Edit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnMP_Edit.Click
        Try
            If Me.TxtMP_DocNo.Text = "" Then
                Me.btnMP_Find.PerformClick()
            End If

            If Me.TxtMP_DocNo.Text <> "" Then
                If Me.ObjMP.ApakahBolehEdit(Me.TxtMP_Primary.Text, Me.UserCode, True, True) = True Then
                    If Me.ObjMP.TimeStamp <> Me.ObjMP.GetTimestamp(Me.ObjMP.Primary) Then
                        Me.ObjMP.CariDetail(Me.ObjMP.Primary, Me.vDT, Me.vDT_ItemGIT, Me.vDT_ItemDO, vDT_ItemPOB, _
                        vDT_ItemSOB, vDT_ItemPrOB, vDT_ItemSettingLocation, vDT_ItemQty, vDT_ItemMOB, vDT_ItemRMReq, Me.vDT_ItemFollowUp, Me.vDT_ItemSOBDelivery, _
                        vDT_ItemPrOBDelivery)
                    End If
                    Me.baru = False
                    Me.AktifkanDetail()
                    'Me.dgvMP .Focus()
                    'nilai min juga harus diatas tanggal pengkuncian data terakhir
                    'msgbox("belum")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnMP_Delete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnMP_Delete.Click
        Dim hasil As MsgBoxResult

        Try
            If Me.TxtMP_DocNo.Text = "" Then
                Me.btnMP_Find.PerformClick()
            End If

            If Me.TxtMP_DocNo.Text <> "" Then
                If Me.ObjMP.ApakahBolehHapus(Me.TxtMP_Primary.Text, Me.UserCode, True, True) = False Then
                    Me.btnMP_Find.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.ObjMP.Hapus(Me.TxtMP_Primary.Text, Me.UserCode) = True Then
                            Me.KosongkanField()
                            Me.btnMP_New.Focus()
                            MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub

    Private Sub FrmMasterProductGroup_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.vPrimaryStockInPurchaseReference <> -1 Then
                Exit Sub
            End If

            If Me.SaveRequired = True Then
                MsgBox("You've just been following up one or more item, you have to save this transaction.", MsgBoxStyle.Information, "Information")
                e.Cancel = True
                Exit Sub
            End If

            If Me.TxtMP_Remarks.Enabled = True Then
                Me.TabControl1.SelectedIndex = 0
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjMP.SimpanBaru(Me.vDT, Me.UserCode, Me, Me.vDT_ItemGIT, Me.vDT_ItemDO, vDT_ItemPOB, vDT_ItemSOB, vDT_ItemPrOB, _
                            vDT_ItemSettingLocation, vDT_ItemQty, vDT_ItemMOB, vDT_ItemRMReq, Me.vDT_ItemFollowUp, Me.vDT_ItemSOBDelivery, _
                            vDT_ItemPrOBDelivery) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.ObjMP.CariDetail(Me.ObjMP.Primary, Me.vDT, Me.vDT_ItemGIT, Me.vDT_ItemDO, vDT_ItemPOB, vDT_ItemSOB, vDT_ItemPrOB, _
                                vDT_ItemSettingLocation, vDT_ItemQty, vDT_ItemMOB, vDT_ItemRMReq, Me.vDT_ItemFollowUp, Me.vDT_ItemSOBDelivery, _
                                vDT_ItemPrOBDelivery)
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnMP_Find.Focus()
                            Else
                                e.Cancel = True
                            End If
                        Else
                            'simpan perubahan
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjMP.SimpanPerubahan(Me.vDT, Me.UserCode, Me, Me.vDT_ItemGIT, Me.vDT_ItemDO, vDT_ItemPOB, vDT_ItemSOB, vDT_ItemPrOB, _
                            vDT_ItemSettingLocation, vDT_ItemQty, vDT_ItemMOB, vDT_ItemRMReq, Me.vDT_ItemFollowUp, Me.vDT_ItemSOBDelivery, _
                            vDT_ItemPrOBDelivery) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnMP_Edit.Focus()
                            Else
                                e.Cancel = True
                            End If
                        End If
                    Else
                        e.Cancel = True
                    End If
                ElseIf hasil = MsgBoxResult.Cancel Then
                    e.Cancel = True
                End If
            End If

            If Me.TxtSF_Remarks.Enabled = True Then
                Me.TabControl1.SelectedIndex = 1
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar_SF() = True Then
                        Me.SalinKeObjek_SF()
                        If baruSF = True Then
                            'Simpan baruSF
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjSF.SimpanBaru(Me.UserCode, Me, vTBLItemPOB, vTBLItemDelivery, Me.vTBLItemStock) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.ObjSF.CariDetail(Me.ObjSF.Primary, Me.vTBLItemStock, vTBLItemPOB, vTBLItemDelivery)
                                Me.SalinKeField_SF()
                                Me.NonAktifkanDetail_SF()
                                Me.btnSF_Find.Focus()
                            Else
                                e.Cancel = True
                            End If
                        Else
                            'Simpan perubahan
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.ObjSF.SimpanPerubahan(Me.UserCode, Me, vTBLItemPOB, vTBLItemDelivery, Me.vTBLItemStock) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField_SF()
                                Me.NonAktifkanDetail_SF()
                                Me.btnSF_Edit.Focus()
                            Else
                                e.Cancel = True
                            End If
                        End If
                    Else
                        e.Cancel = True
                    End If
                ElseIf hasil = MsgBoxResult.Cancel Then
                    e.Cancel = True
                End If
            End If

            'simpan setting PrOB > POB
            If Me.chkFilter_PrOBLargerThanPOB.Checked = False Then
                Me.ObjSettingUser.SetSettingValue(ClsTransaksiSettingUser.EnumTipeTransaksi.MaterialPlanning, _
                ClsTransaksiSettingUser.EnumNamaParameter.MaterialPlanningGrouping_PrOBLargerThanPOB, "0", Me.UserCode)
            Else
                Me.ObjSettingUser.SetSettingValue(ClsTransaksiSettingUser.EnumTipeTransaksi.MaterialPlanning, _
                ClsTransaksiSettingUser.EnumNamaParameter.MaterialPlanningGrouping_PrOBLargerThanPOB, "1", Me.UserCode)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_Remarks_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMP_Remarks.GotFocus
        Try
            Me.TxtMP_Remarks.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_Remarks_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMP_Remarks.LostFocus
        Try
            Me.TxtMP_Remarks.BackColor = Me.TxtMP_Primary.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpMP_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpMP.Leave
        Try
            If Me.TxtMP_Remarks.Enabled = True Then
                Me.BackgroundWorker1.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            If Me.TxtMP_Remarks.Enabled = True Then
                If Me.ObjNumbering.Apakah_Autonumber(ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering._MP) = True Then
                    If Me.baru = False Then
                        If Me.ObjMP.TanggalTransaksi.Month = Me.dtpMP.Value.Month And Me.ObjMP.TanggalTransaksi.Year = Me.dtpMP.Value.Year Then
                            'jika kembali ke bulan yang bersangkutan maka no kembali ke awal
                            Me.TxtMP_DocNo.Text = Me.ObjMP.NoTransaksi
                        Else
                            'jika pindah bulan maka akan ganti no
                            Me.TxtMP_DocNo.Text = Me.ObjMP.GetNoBaru1(Me.dtpMP.Value)
                        End If
                    Else
                        Me.TxtMP_DocNo.Text = Me.ObjMP.GetNoBaru1(Me.dtpMP.Value)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_DocNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMP_DocNo.GotFocus
        Try
            If Me.TxtMP_Remarks.Enabled = True And Me.TxtMP_DocNo.Enabled = True Then
                If Me.TxtMP_DocNo.ReadOnly = True Then
                    Me.TxtMP_DocNo.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtMP_DocNo.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_DocNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMP_DocNo.LostFocus
        Try
            If Me.TxtMP_Remarks.Enabled = True And Me.TxtMP_DocNo.Enabled = True Then
                If Me.TxtMP_DocNo.ReadOnly = True Then
                    Me.TxtMP_DocNo.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtMP_DocNo.BackColor = Me.TxtMP_Primary.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TabControl2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabControl2.SelectedIndexChanged
        Try
            If Me.TxtMP_Remarks.Enabled = True Then
                If Me.TabControl2.SelectedIndex = 1 Then
                    'pastikan user memiliki hak akses untuk melihat detail pembelian 
                    If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_MaterialPlanningDashboard_LihatDetailPembelian") = False Then
                        Me.TabControl2.SelectedIndex = 0
                        MsgBox("You don't have authorization form this information.", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If

                    Me.ObjCD.SetBackColor(Me.SplitContainer4.Panel1.Controls, Me.TxtMP_Primary.BackColor, False)
                    Me.ObjCD.SetBackColor(Me.SplitContainer4.Panel2.Controls, Me.TxtMP_Primary.BackColor, False)
                    Me.ObjCD.SetBackColor(Me.SplitContainer5.Panel1.Controls, Me.TxtMP_Primary.BackColor, False)
                    Me.ObjCD.SetBackColor(Me.SplitContainer5.Panel2.Controls, Me.TxtMP_Primary.BackColor, False)
                    Me.SetTxtPurchaseDetail()
                End If
            Else
                If Me.TabControl2.SelectedIndex = 1 Then
                    'pastikan user memiliki hak akses untuk melihat detail pembelian 
                    If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_MaterialPlanningDashboard_LihatDetailPembelian") = False Then
                        Me.TabControl2.SelectedIndex = 0
                        MsgBox("You don't have authorization for this information.", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If

                    Me.ObjCD.SetBackColor(Me.SplitContainer4.Panel1.Controls, Me.BackColor, False)
                    Me.ObjCD.SetBackColor(Me.SplitContainer4.Panel2.Controls, Me.BackColor, False)
                    Me.ObjCD.SetBackColor(Me.SplitContainer5.Panel1.Controls, Me.BackColor, False)
                    Me.ObjCD.SetBackColor(Me.SplitContainer5.Panel2.Controls, Me.BackColor, False)
                    Me.SetTxtPurchaseDetail()
                End If
            End If
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
            If Me.TxtMP_Remarks.Enabled = True Then
                If Me.dgvMP.SelectedCells.Count = 0 Then
                    ModCekData.KendalikanControl(Me.SplitContainer4.Panel1.Controls, True, False, True, False, True, False, True, False)
                    ModCekData.KendalikanControl(Me.SplitContainer4.Panel2.Controls, True, False, True, False, True, False, True, False)
                    Me.ObjCD.SetBackColor(Me.SplitContainer4.Panel1.Controls, Me.BackColor, False)
                    Me.ObjCD.SetBackColor(Me.SplitContainer4.Panel2.Controls, Me.BackColor, False)
                Else
                    ModCekData.KendalikanControl(Me.SplitContainer4.Panel1.Controls, True, True, True, True, True, True, True, True)
                    ModCekData.KendalikanControl(Me.SplitContainer4.Panel2.Controls, True, True, True, True, True, True, True, True)
                    Me.ObjCD.SetBackColor(Me.SplitContainer4.Panel1.Controls, Me.TxtMP_Primary.BackColor, False)
                    Me.ObjCD.SetBackColor(Me.SplitContainer4.Panel2.Controls, Me.TxtMP_Primary.BackColor, False)
                End If
            Else
                ModCekData.KendalikanControl(Me.SplitContainer4.Panel1.Controls, True, False, True, False, True, False, True, False)
                ModCekData.KendalikanControl(Me.SplitContainer4.Panel2.Controls, True, False, True, False, True, False, True, False)
                Me.ObjCD.SetBackColor(Me.SplitContainer4.Panel1.Controls, Me.BackColor, False)
                Me.ObjCD.SetBackColor(Me.SplitContainer4.Panel2.Controls, Me.BackColor, False)
            End If
            Me.cboMPOrderStrategy.Enabled = False

            If Me.dgvMP.SelectedCells.Count > 0 Then
                rowDGV = Me.dgvMP.Rows(Me.dgvMP.SelectedCells(0).RowIndex)
                For Each rowItem As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item1Row In Me.vDT.Rows
                    If rowItem.primary_code_no = rowDGV.Cells(Me.ColMP_PrimaryCodeNo.Index).Value Then
                        With rowItem
                            'detail qty
                            Me.TxtMP_QtyMOB.Text = Me.ObjCD.DigitGrouping(.qty_mob, .default_prec_unit)
                            Me.TxtMP_QtyRMReq.Text = Me.ObjCD.DigitGrouping(.qty_rm_req, .default_prec_unit)
                            Me.TxtMP_QtyRMReqSpare.Text = Me.ObjCD.DigitGrouping(.MP_Spare_Qty_Rm_Req, .default_prec_unit)
                            Me.TxtMP_QtyGIT.Text = Me.ObjCD.DigitGrouping(.qty_git, .default_prec_unit)
                            Me.TxtMP_QtyInquiry.Text = Me.ObjCD.DigitGrouping(.qty_inquiry, .default_prec_unit)
                            Me.TxtMP_QtyOptimumStock.Text = Me.ObjCD.DigitGrouping(.qty_optimum, .default_prec_unit)
                            Me.TxtMP_QtyOnHand.Text = Me.ObjCD.DigitGrouping(.qty_on_hand, .default_prec_unit)
                            Me.TxtMP_QtyFree.Text = Me.ObjCD.DigitGrouping(.qty_on_hand + .qty_pob + .qty_git + .qty_mob - .qty_sob - .qty_pending_do - .qty_prob - _
                            .qty_rm_req - .MP_Spare_Qty_PrOB - .MP_Spare_Qty_Rm_Req - .MP_Spare_Qty_SOB, .default_prec_unit) '-r8
                            Me.TxtMP_QtyOrder.Text = Me.ObjCD.DigitGrouping(Math.Round(CDec((.qty_order * .ratio_multi) / .ratio_default), _
                            CInt(.def_prec_unit_multi), MidpointRounding.AwayFromZero), .def_prec_unit_multi)
                            Me.TxtMP_QtyPendingDO.Text = Me.ObjCD.DigitGrouping(.qty_pending_do, .default_prec_unit)
                            Me.TxtMP_QtyPOB.Text = Me.ObjCD.DigitGrouping(.qty_pob, .default_prec_unit)
                            Me.TxtMP_QtyRequired.Text = Me.ObjCD.DigitGrouping(.qty_required, .default_prec_unit)
                            Me.TxtMP_QtyRequired2.Text = Me.ObjCD.DigitGrouping(.qty_required, .default_prec_unit)
                            Me.TxtMP_QtySafetyStock.Text = Me.ObjCD.DigitGrouping(.qty_safety_stock, .default_prec_unit)
                            Me.TxtMP_QtySOB.Text = Me.ObjCD.DigitGrouping(.qty_sob, .default_prec_unit)
                            Me.TxtMP_QtySOBSpare.Text = Me.ObjCD.DigitGrouping(.MP_Spare_Qty_SOB, .default_prec_unit)
                            Me.TxtMP_QtyPrOB.Text = Me.ObjCD.DigitGrouping(.qty_prob, .default_prec_unit)
                            Me.TxtMP_QtyPrOBSpare.Text = Me.ObjCD.DigitGrouping(.MP_Spare_Qty_PrOB, .default_prec_unit)
                            Me.TxtMP_AvgPerMonth.Text = Me.ObjCD.DigitGrouping(.qty_avg_montly_demand, .default_prec_unit)

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
                            Me.TxtMP_UnitPrice.Text = Me.ObjCD.DigitGrouping(.unit_price, .def_prec_currency)
                            Me.TxtMP_Discount.Text = Me.ObjCD.DigitGrouping(.discount, 2)
                            Me.TxtMP_DiscFromPerct.Text = .disc_from_perct
                            Me.TxtMP_DiscAmount.Text = Me.ObjCD.DigitGrouping(.disc_amount, .def_prec_currency)
                            Me.TxtMP_NettPrice.Text = Me.ObjCD.DigitGrouping(.nett_price, .def_prec_currency)
                            Me.TxtMP_SubTotal.Text = Me.ObjCD.DigitGrouping(.sub_total, .def_prec_currency)

                            Me.LblMP_Unit1.Text = .Unit_Code
                            Me.llbMP_Ratio.Text = "n/a"

                            If .primary_unit_multi <> .primary_unit Then
                                Me.llbMP_Ratio.Text = Me.ObjCD.DigitGrouping(.ratio_multi, .def_prec_unit_multi) & " " & .Unit_Code_Multi & _
                                " = " & Me.ObjCD.DigitGrouping(.ratio_default, .def_prec_order_unit) & " " & .Unit_Code + _
                                " --> Req. Qty = " + Me.ObjCD.DigitGrouping((.qty_required * .ratio_multi) / .ratio_default, .def_prec_unit_multi) + " " + .Unit_Code_Multi
                            End If

                            Me.ObjProd.IsiCBOMultiUnit(Me.cboMP_OrderUnit, .primary_code_no, Me.DTAProduct_MU, _
                            Me.TBLProduct_MU, Me.dtaMultiUnit, Me.tblMultiUnit, .primary_unit_multi)

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
            Else
                Me.TxtMP_QtyMOB.Text = 0
                Me.TxtMP_QtyRMReq.Text = 0
                Me.TxtMP_QtyRMReqSpare.Text = 0
                Me.TxtMP_QtyGIT.Text = 0
                Me.TxtMP_QtyInquiry.Text = 0
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
                Me.TxtMP_QtySOBSpare.Text = 0
                Me.TxtMP_QtyPrOB.Text = 0
                Me.TxtMP_QtyPrOBSpare.Text = 0
                Me.TxtMP_AvgPerMonth.Text = 0

                Me.cboMPOrderStrategy.Text = "Order To Safety Stock"

                'detail purchase
                Me.TxtMP_PrimarySupplier.Text = -1
                Me.TxtMP_PreferedSupplier.Text = ""
                Me.TxtMP_PreferedSupplierDesc.Text = ""
                Me.TxtMP_PrimaryCurrency.Text = -1
                Me.TxtMP_CurrencyCode.Text = ""
                Me.TxtMP_UnitPrice.Text = 0
                Me.TxtMP_Discount.Text = 0
                Me.TxtMP_DiscFromPerct.Text = 1
                Me.TxtMP_DiscAmount.Text = 0
                Me.TxtMP_NettPrice.Text = 0
                Me.TxtMP_SubTotal.Text = 0
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

    Private Sub dgvMP_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvMP.CellBeginEdit
        Dim ShowMsgbox As Boolean = False

        Try
            'yang boleh diedit hanya: qty order, follow up, remarks
            If e.ColumnIndex = Me.ColMP_QtyOrder.Index Or _
            e.ColumnIndex = Me.ColMP_FollowUp.Index Then 'Or _
                'e.ColumnIndex = Me.ColMP_OrderStrategy.Index Then
                If e.ColumnIndex = Me.ColMP_FollowUp.Index Then
                    If Me.BufferCekBox <> Me.dgvMP.Item(Me.ColMP_CodeNo.Index, e.RowIndex).Value Then
                        Me.BufferCekBox = Me.dgvMP.Item(Me.ColMP_CodeNo.Index, e.RowIndex).Value
                    Else
                        ShowMsgbox = True
                    End If
                    If CDec(Me.dgvMP.Item(Me.ColMP_QtyOrder.Index, e.RowIndex).Value) <= 0 Then
                        If ShowMsgbox = True Then
                            MsgBox("Please specify order qty first.", MsgBoxStyle.Information, "Information")
                        End If
                        e.Cancel = True
                    End If
                End If

                If e.ColumnIndex = Me.ColMP_QtyOrder.Index Then
                    If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_MaterialPlanningDashboard_MP_EditOrderQty") = False Then
                        MsgBox("You don't have access to modify order qty.", MsgBoxStyle.Information, "Information")
                        e.Cancel = True
                        Exit Sub
                    End If
                    If Me.IsValidToModifyData(Me.TxtMP_Primary.Text) = False Then
                        MsgBox("This MP is already authorized, you can't modify qty order.", MsgBoxStyle.Information, "Information")
                        e.Cancel = True
                        Exit Sub
                    End If
                End If

                Me.Buffer = Me.dgvMP.Item(e.ColumnIndex, e.RowIndex).Value
                If e.ColumnIndex = Me.ColMP_QtyOrder.Index Then
                    Me.dgvMP.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.BuangNolDibelakangDecimal(Me.dgvMP.Item(e.ColumnIndex, e.RowIndex).Value)
                End If
            ElseIf e.ColumnIndex = Me.ColMP_ProcurementType.Index Then
                If CDec(Me.dgvMP.Item(Me.ColMP_QtyOrder.Index, e.RowIndex).Value) <= 0 Then
                    MsgBox("Please specify order qty first.", MsgBoxStyle.Information, "Information")
                    e.Cancel = True
                End If
                Me.Buffer = Me.dgvMP.Item(e.ColumnIndex, e.RowIndex).Value
            Else
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEdit_Qty(ByVal sender As Object, ByVal vRowIndex As Integer, ByVal vColIndex As Integer)
        Dim DefPrec As Integer = 0
        Dim JmlMultiDummy As Decimal = 0
        Dim JmlDefaultDummy As Decimal = 0
        Dim Dummy As String = ""

        Try
            'qty order
            'pastikan positif
            If Me.ObjCD.CekNonNegatif(Me.dgvMP.Item(vColIndex, vRowIndex).Value) = False Then
                Me.dgvMP.Item(vColIndex, vRowIndex).Value = Buffer
                Exit Sub
            End If


            'hitung ulang next order point date
            For Each rowItem As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item1Row In Me.vDT.Rows
                If rowItem.primary_code_no = Me.dgvMP.Item(Me.ColMP_PrimaryCodeNo.Index, vRowIndex).Value Then
                    'digit grouping sebelum digunakan untuk perhitungan
                    Me.dgvMP.Item(Me.ColMP_QtyOrder.Index, vRowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvMP.Item(Me.ColMP_QtyOrder.Index, vRowIndex).Value, rowItem.default_prec_unit, ClsCekData.EnumTipeRounding.Middlepoint)

                    'jangan lupa update k vdt
                    With rowItem
                        'jika primary unit multi <> primary unit multi, maka pastikan qty unit adalah kelipatan terkecil dari satuan multi
                        If .primary_unit <> .primary_unit_multi Then
                            JmlMultiDummy = Math.Round(CDec((CDec(Me.dgvMP.Item(Me.ColMP_QtyOrder.Index, vRowIndex).Value) * .ratio_multi) / .ratio_default), CInt(.def_prec_unit_multi), MidpointRounding.AwayFromZero)
                            JmlDefaultDummy = Math.Round(CDec((JmlMultiDummy * .ratio_default) / .ratio_multi), CInt(.default_prec_unit), MidpointRounding.AwayFromZero)
                            If JmlDefaultDummy <> CDec(Me.dgvMP.Item(Me.ColMP_QtyOrder.Index, vRowIndex).Value) Then
                                Dummy = (Me.ObjCD.DigitGrouping(JmlDefaultDummy, .default_prec_unit))
                                'MsgBox("We change qty data to: " + Dummy + " " + .Unit_Code + _
                                '" to maintain multi unit ratio consistency.", MsgBoxStyle.Information, "Information")
                                Me.dgvMP.Item(Me.ColMP_QtyOrder.Index, vRowIndex).Value = Dummy
                            End If
                        End If

                        .qty_order = Me.dgvMP.Item(Me.ColMP_QtyOrder.Index, vRowIndex).Value
                        .qty_order_multiunit = Math.Round(CDec((.qty_order * .ratio_multi) / .ratio_default), CInt(.def_prec_unit_multi), MidpointRounding.AwayFromZero)

                        'digit grouping
                        Me.dgvMP.Item(Me.ColMP_QtyOrder.Index, vRowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvMP.Item(Me.ColMP_QtyOrder.Index, vRowIndex).Value, .default_prec_unit)
                        Me.dgvMP.Item(Me.ColMP_QtyOrderMulti.Index, vRowIndex).Value = .qty_order_multiunit

                        If Me.dgvMP.Item(Me.ColMP_ProcurementStrategy.Index, vRowIndex).Value <> "Build To Order" Then
                            Me.ObjCls.HitungOrderPoint(rowItem, Me.TxtMP_DaysPerMonth.Text, Me.dtpMP_RefDate, Me.ObjCD)

                            Me.dgvMP.Item(Me.ColMP_NextOrderPointDelay.Index, Me.dgvMP.SelectedCells(0).RowIndex).Value = .next_order_point_delay
                            If .order_point <> #1/1/2200# Then
                                Me.dgvMP.Item(Me.ColMP_NextOrderPoint.Index, Me.dgvMP.SelectedCells(0).RowIndex).Value = .next_order_point
                            Else
                                Me.dgvMP.Item(Me.ColMP_NextOrderPoint.Index, Me.dgvMP.SelectedCells(0).RowIndex).Value = "-"
                            End If
                        End If

                        'hitung ulang sub total
                        If .primary_currency <> -1 Then
                            .sub_total = Math.Round(.nett_price * .qty_order_multiunit, Me.ObjCurrency.GetDefaultPrecission(.primary_currency), MidpointRounding.AwayFromZero)
                        Else
                            .sub_total = Math.Round(.nett_price * .qty_order_multiunit, 0, MidpointRounding.AwayFromZero)
                        End If
                    End With

                    'hitung sub total
                    If Me.dgvMP.Item(Me.ColMP_PrimaryCurrency.Index, vRowIndex).Value <> -1 Then
                        DefPrec = Me.ObjCurrency.GetDefaultPrecission(Me.dgvMP.Item(Me.ColMP_PrimaryCurrency.Index, vRowIndex).Value)
                    Else
                        DefPrec = 0
                    End If
                    Me.dgvMP.Item(Me.ColMP_SubTotal.Index, vRowIndex).Value = Me.ObjCD.DigitGrouping(rowItem.sub_total, DefPrec)

                    'update tampilan harga
                    SetTxtPurchaseDetail()
                    Exit Sub
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvMP_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvMP.CellEndEdit
        Dim OrderStrategy As Core_Module.clsMasterProduct.EnumOrderStrategy = clsMasterProduct.EnumOrderStrategy.Order_By_Maximum_Order_Qty
        Dim JmlDaysPerMonth As Integer = 0
        Dim JmlDayPerMontStr As String = ""

        Try
            JmlDaysPerMonth = Me.TxtMP_DaysPerMonth.Text

            If IsNothing(Me.dgvMP.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                If e.ColumnIndex = Me.ColMP_QtyOrder.Index Then
                    'qty order
                    Me.dgvMP.Item(e.ColumnIndex, e.RowIndex).Value = Buffer
                ElseIf e.ColumnIndex = Me.ColMP_FollowUp.Index Then
                    'follow up
                    Me.dgvMP.Item(e.ColumnIndex, e.RowIndex).Value = Boolean.FalseString
                ElseIf e.ColumnIndex = Me.ColMP_OrderStrategy.Index Then
                    'order strategy
                    Me.dgvMP.Item(e.ColumnIndex, e.RowIndex).Value = Buffer
                Else
                    'procurement type
                    Me.dgvMP.Item(e.ColumnIndex, e.RowIndex).Value = Buffer
                End If
            Else
                If e.ColumnIndex = Me.ColMP_QtyOrder.Index Then
                    Me.EndEdit_Qty(sender, e.RowIndex, e.ColumnIndex)
                ElseIf e.ColumnIndex = Me.ColMP_ProcurementType.Index Then
                    For Each rowItem As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item1Row In Me.vDT.Rows
                        If rowItem.primary_code_no = Me.dgvMP.Item(Me.ColMP_PrimaryCodeNo.Index, e.RowIndex).Value Then
                            If Me.dgvMP.Item(Me.ColMP_ProcurementType.Index, e.RowIndex).Value = "Buy" Then
                                rowItem.procurement_type = Core_Module.clsMasterProduct.EnumProcurementType.Buy
                            Else
                                rowItem.procurement_type = Core_Module.clsMasterProduct.EnumProcurementType.Produce
                            End If
                            Exit Sub
                        End If
                    Next
                ElseIf e.ColumnIndex = Me.ColMP_FollowUp.Index Then
                    'follow up
                    For Each rowItem As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item1Row In Me.vDT.Rows
                        If rowItem.primary_code_no = Me.dgvMP.Item(Me.ColMP_PrimaryCodeNo.Index, e.RowIndex).Value Then
                            'cek apakah sudah ada isinya:
                            '1. prefered supplier
                            '2. unit price
                            '3. currency
                            '4. qty order
                            If rowItem.primary_prefered_supplier = -1 Or rowItem.primary_currency = -1 Or _
                            rowItem.unit_price = 0 Or rowItem.qty_order = 0 Then

                            End If
                            If Me.dgvMP.Item(Me.ColMP_FollowUp.Index, e.RowIndex).Value = Boolean.TrueString Then
                                rowItem.modify = 1
                            Else
                                rowItem.modify = 0
                            End If

                            SetStatusCek()
                            Exit Sub
                        End If
                    Next
                ElseIf e.ColumnIndex = Me.ColMP_OrderStrategy.Index Then
                    'order strategy
                    For Each rowItem As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item1Row In Me.vDT.Rows
                        If rowItem.primary_code_no = Me.dgvMP.Item(Me.ColMP_PrimaryCodeNo.Index, e.RowIndex).Value Then
                            'jangan lupa update k vdt
                            With rowItem
                                .qty_required = .qty_inquiry + .qty_sob + .qty_prob + .qty_pending_do + .qty_safety_stock - _
                                .qty_on_hand - .qty_pob - .qty_git

                                If Me.dgvMP.Item(e.ColumnIndex, e.RowIndex).Value = "Order To Min Qty" Then
                                    OrderStrategy = clsMasterProduct.EnumOrderStrategy.Order_To_Min_Qty
                                ElseIf Me.dgvMP.Item(e.ColumnIndex, e.RowIndex).Value = "Order To Optimum Qty" Then
                                    OrderStrategy = clsMasterProduct.EnumOrderStrategy.Order_To_Min_Qty
                                ElseIf Me.dgvMP.Item(e.ColumnIndex, e.RowIndex).Value = "Order To Max Qty" Then
                                    OrderStrategy = clsMasterProduct.EnumOrderStrategy.Order_To_Min_Qty
                                ElseIf Me.dgvMP.Item(e.ColumnIndex, e.RowIndex).Value = "Order To Safety Stock" Then
                                    OrderStrategy = clsMasterProduct.EnumOrderStrategy.Order_To_Min_Qty
                                ElseIf Me.dgvMP.Item(e.ColumnIndex, e.RowIndex).Value = "Order By Minimum Order Qty" Then
                                    OrderStrategy = clsMasterProduct.EnumOrderStrategy.Order_To_Min_Qty
                                ElseIf Me.dgvMP.Item(e.ColumnIndex, e.RowIndex).Value = "Order By Optimum Order Qty" Then
                                    OrderStrategy = clsMasterProduct.EnumOrderStrategy.Order_To_Min_Qty
                                ElseIf Me.dgvMP.Item(e.ColumnIndex, e.RowIndex).Value = "Order By Maximum Order Qty" Then
                                    OrderStrategy = clsMasterProduct.EnumOrderStrategy.Order_To_Min_Qty
                                End If

                                .order_strategy = OrderStrategy
                                If OrderStrategy = clsMasterProduct.EnumOrderStrategy.Order_By_Maximum_Order_Qty Then
                                    .qty_required = .qty_order_maksimum
                                ElseIf OrderStrategy = clsMasterProduct.EnumOrderStrategy.Order_By_Minimum_Order_Qty Then
                                    .qty_required = .qty_order_minimum
                                ElseIf OrderStrategy = clsMasterProduct.EnumOrderStrategy.Order_By_Optimum_Order_Qty Then
                                    .qty_required = .qty_order_optimum
                                ElseIf OrderStrategy = clsMasterProduct.EnumOrderStrategy.Order_To_Max_Qty Then
                                    If .qty_required > 0 Then
                                        .qty_required = .qty_required + .qty_max - .qty_safety_stock
                                    Else
                                        .qty_required = .qty_max - .qty_safety_stock
                                    End If
                                    If .qty_required < 0 Then
                                        .qty_required = 0
                                    End If
                                ElseIf OrderStrategy = clsMasterProduct.EnumOrderStrategy.Order_To_Min_Qty Then
                                    If .qty_required > 0 Then
                                        .qty_required = .qty_required + .qty_min - .qty_safety_stock
                                    Else
                                        .qty_required = .qty_min - .qty_safety_stock
                                    End If
                                    If .qty_required < 0 Then
                                        .qty_required = 0
                                    End If
                                ElseIf OrderStrategy = clsMasterProduct.EnumOrderStrategy.Order_By_Optimum_Order_Qty Then
                                    If .qty_required > 0 Then
                                        .qty_required = .qty_required + .qty_optimum - .qty_safety_stock
                                    Else
                                        .qty_required = .qty_optimum - .qty_safety_stock
                                    End If
                                    If .qty_required < 0 Then
                                        .qty_required = 0
                                    End If
                                ElseIf OrderStrategy = clsMasterProduct.EnumOrderStrategy.Order_To_Safety_Stock Then
                                    If .qty_required > 0 Then
                                        .qty_required = .qty_required
                                    End If
                                    If .qty_required < 0 Then
                                        .qty_required = 0
                                    End If
                                End If
                                .qty_order = .qty_required
                                .qty_order_multiunit = .qty_order

                                If .qty_avg_montly_demand = 0 Then
                                    .qty_avg_montly_demand = 1
                                End If
                                .order_point_delay = ObjCD.RoundDown1(((.qty_on_hand + .qty_pob + .qty_git - .qty_inquiry - .qty_sob - .qty_prob - .qty_pending_do - .qty_safety_stock) * JmlDaysPerMonth) / .qty_avg_montly_demand, 0) - .lead_time_days
                                If .order_point_delay > (Date.MaxValue.Date - Me.dtpMP_RefDate.Value.Date).Days Then
                                    .order_point = #1/1/1950#
                                Else
                                    .order_point = Me.dtpMP_RefDate.Value + TimeSpan.FromDays(.order_point_delay)
                                End If
                                .next_order_point_delay = ObjCD.RoundDown1(((.qty_order + .qty_on_hand + .qty_pob + .qty_git - .qty_inquiry - .qty_sob - .qty_prob - .qty_pending_do - .qty_safety_stock) * JmlDaysPerMonth) / .qty_avg_montly_demand, 0) - .lead_time_days
                                If .next_order_point_delay > (Date.MaxValue.Date - Me.dtpMP_RefDate.Value.Date).Days Then
                                    .next_order_point = #1/1/1950#
                                Else
                                    .next_order_point = Me.dtpMP_RefDate.Value + TimeSpan.FromDays(.next_order_point_delay)
                                End If

                                Me.dgvMP.Item(Me.ColMP_NextOrderPointDelay.Index, e.RowIndex).Value = .next_order_point_delay
                                Me.dgvMP.Item(Me.ColMP_NextOrderPoint.Index, e.RowIndex).Value = .next_order_point
                                Me.dgvMP.Item(Me.ColMP_OrderPointDelay.Index, e.RowIndex).Value = .order_point_delay
                                Me.dgvMP.Item(Me.ColMP_OrderPoint.Index, e.RowIndex).Value = .order_point
                                Me.dgvMP.Item(Me.ColMP_QtyOrder.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(.qty_order, .default_prec_unit)
                                Me.dgvMP.Item(Me.ColMP_QtyRequired.Index, e.RowIndex).Value = Me.ObjCD.DigitGrouping(.qty_required, .default_prec_unit)
                            End With

                            Exit Sub
                        End If
                    Next
                End If
                'tandai bahwa edit data sudah selesai
                Me.BufferCekBox = ""
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
            If e.ColumnIndex = Me.ColMP_QtyRequired.Index Or _
            e.ColumnIndex = Me.ColMP_OrderPoint.Index Or _
            e.ColumnIndex = Me.ColMP_NextOrderPoint.Index Then
                'tampilkan penjelasan mengenai qty required
                If Me.dgvMP.SelectedCells.Count = 0 Then
                    Exit Sub
                Else
                    rowDGV = Me.dgvMP.Rows(e.RowIndex)
                    For Each row As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item1Row In Me.vDT.Rows
                        If row.primary_code_no = rowDGV.Cells(Me.ColMP_PrimaryCodeNo.Index).Value Then
                            i.ObjCD = Me.ObjCD
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
                If Me.btnMP_OK.Enabled = False Then j.ReadOnly_ = True
                j.ShowDialog()
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

    Private Sub llbMP_Calculate_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbMP_Calculate.LinkClicked
        Try
            If Me.TxtMP_Remarks.Enabled = True Then
                If Me.TxtMP_Primary.Text <> "-1" Then
                    If Me.IsValidToModifyData(Me.TxtMP_Primary.Text) = False Then
                        MsgBox("This MP is already authorized, you can't recalculate this data.", MsgBoxStyle.Information)
                        Exit Sub
                    End If
                End If
                Me.HitungMaterialPlanning()
                Me.LoadDGV()
                'Control.CheckForIllegalCrossThreadCalls = False
                'Me.BackgroundWorker2.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub HitungMaterialPlanning()
        Try
            If MsgBox("Calculate material planning?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                Exit Sub
            Else 
                Me.vDT_ItemFollowUp.Rows.Clear()

                Dim obj As New ClsTransaksiMaterialPlanning_Calc
                obj.HitungMaterialPlanning(Me.vDT, Me.vDT_ItemGIT, Me.vDT_ItemDO, Me.vDT_ItemPOB, Me.vDT_ItemSOB, Me.vDT_ItemPrOB, Me.vDT_ItemSettingLocation, _
                Me.vDT_ItemQty, Me.vDT_ItemMOB, Me.vDT_ItemRMReq, Me.ObjSetting, Me.TxtMP_DaysPerMonth, Me.BackColorAll, Me.FrmCek, Me.UserCode, Me.ObjCD, _
                Me.TxtMP_Primary, Me.ObjSettingUser, Me.dtpMP_RefDate, Me.vDT_ItemSOBDelivery, Me.vDT_ItemPrOBDelivery)

            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Try
            Me.HitungMaterialPlanning()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker2.ProgressChanged
        Try
            Me.FrmCek.ShowDialog()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Try
            Me.LoadDGV()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpFilter_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpMP_Filter.Enter
        Try
            Me.Buffer = Me.dtpMP_Filter.Value.Date
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpFilter_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpMP_Filter.Leave
        Try
            If CDate(Me.Buffer).Date <> Me.dtpMP_Filter.Value.Date Then
                Me.BackgroundWorker3.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkMP_RequiredOnly_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkMP_RequiredOnly.CheckedChanged
        Try
            If Me.chkMP_RequiredOnly.Checked = True Then
                Me.dtpMP_Filter.Enabled = True
            Else
                Me.dtpMP_Filter.Enabled = False
            End If
            Me.LoadDGV()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker3_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker3.DoWork
        Threading.Thread.Sleep(100)
    End Sub

    Private Sub BackgroundWorker3_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker3.RunWorkerCompleted
        Try
            Me.LoadDGV()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_QtySOB_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TxtMP_QtySOB.MouseClick
        Dim i As New FrmTrans_MaterialPlanningDashboard_ExplainSOB
        Dim rowDGV As DataGridViewRow
        Dim vRowDT As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item1Row

        Try
            If Me.dgvMP.SelectedCells.Count = 0 Then
                Exit Sub
            Else
                rowDGV = Me.dgvMP.Rows(Me.dgvMP.SelectedCells(0).RowIndex)

                If CDec(Me.TxtMP_QtySOB.Text) > 0 Then
                    For Each vRowDT In Me.vDT.Rows
                        If vRowDT.primary_code_no = rowDGV.Cells(Me.ColMP_PrimaryCodeNo.Index).Value Then
                            i.vPrimaryProduct = rowDGV.Cells(Me.ColMP_PrimaryCodeNo.Index).Value
                            i.vDT_ItemSOB = Me.vDT_ItemSOB
                            i.vDT_ItemSOBDelivery = Me.vDT_ItemSOBDelivery
                            i.BackColorAll = Me.BackColorAll
                            i.CodeNo = rowDGV.Cells(Me.ColMP_CodeNo.Index).Value
                            i.UnitCode = rowDGV.Cells(Me.ColMP_Unit.Index).Value
                            i.ReadOnly1 = Me.btnMP_Find.Enabled
                            i.UserCode = Me.UserCode
                            i.rowProduct = vRowDT

                            i.ShowDialog()

                            If i.AdaPerubahan = True Then
                                Me.ObjCls.HitungMaterialPlanning(Me.vDT, Me.vDT_ItemGIT, Me.vDT_ItemDO, Me.vDT_ItemPOB, Me.vDT_ItemSOB, Me.vDT_ItemPrOB, Me.vDT_ItemSettingLocation, _
                                Me.vDT_ItemQty, Me.vDT_ItemMOB, Me.vDT_ItemRMReq, Me.ObjSetting, Me.TxtMP_DaysPerMonth, Me.BackColorAll, Me.FrmCek, Me.UserCode, Me.ObjCD, _
                                Me.TxtMP_Primary, Me.ObjSettingUser, Me.dtpMP_RefDate, i.vPrimaryProduct)
                                Me.dgvMP.ClearSelection()
                                rowDGV.Selected = True
                            End If

                            Exit Sub
                        End If
                    Next

                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_QtyPendingDO_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TxtMP_QtyPendingDO.MouseClick
        Dim i As New FrmTrans_MaterialPlanningDashboard_ExplainPendingDO
        Dim rowDGV As DataGridViewRow

        Try
            If Me.dgvMP.SelectedCells.Count = 0 Then
                Exit Sub
            Else
                rowDGV = Me.dgvMP.Rows(Me.dgvMP.SelectedCells(0).RowIndex)

                If CDec(Me.TxtMP_QtyPendingDO.Text) > 0 Then
                    i.vPrimaryProduct = rowDGV.Cells(Me.ColMP_PrimaryCodeNo.Index).Value
                    i.vUnitCode = rowDGV.Cells(Me.ColMP_Unit.Index).Value
                    i.vDT_ItemDO = Me.vDT_ItemDO
                    i.BackColorAll = Me.BackColorAll

                    i.ShowDialog()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_QtyPOB_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TxtMP_QtyPOB.MouseClick
        Dim i As New FrmTrans_MaterialPlanningDashboard_ExplainPOB
        Dim rowDGV As DataGridViewRow

        Try
            If Me.dgvMP.SelectedCells.Count = 0 Then
                Exit Sub
            Else
                rowDGV = Me.dgvMP.Rows(Me.dgvMP.SelectedCells(0).RowIndex)

                If CDec(Me.TxtMP_QtyPOB.Text) > 0 Then
                    i.vPrimaryProduct = rowDGV.Cells(Me.ColMP_PrimaryCodeNo.Index).Value
                    i.vDT_ItemPOB = Me.vDT_ItemPOB
                    i.BackColorAll = Me.BackColorAll
                    i.vUnitCode = rowDGV.Cells(Me.ColMP_Unit.Index).Value

                    i.ShowDialog()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_QtyRequired_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TxtMP_QtyRequired.MouseClick
        Dim i As New FrmTrans_MaterialPlanningDashboard_ExplainRequired
        Dim rowDGV As DataGridViewRow

        Try
            If Me.dgvMP.SelectedCells.Count = 0 Then
                Exit Sub
            Else
                rowDGV = Me.dgvMP.Rows(Me.dgvMP.SelectedCells(0).RowIndex)
                For Each row As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item1Row In Me.vDT.Rows
                    If row.primary_code_no = rowDGV.Cells(Me.ColMP_PrimaryCodeNo.Index).Value Then
                        i.ObjCD = Me.ObjCD
                        i.vDPM = Me.TxtMP_DaysPerMonth.Text
                        i.vRefDate = Me.dtpMP_RefDate.Value
                        i.vRow = row

                        i.ShowDialog()
                        Exit Sub
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_QtyGIT_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TxtMP_QtyGIT.MouseClick
        Dim i As New FrmTrans_MaterialPlanningDashboard_ExplainGIT
        Dim rowDGV As DataGridViewRow

        Try
            If Me.dgvMP.SelectedCells.Count = 0 Then
                Exit Sub
            Else
                rowDGV = Me.dgvMP.Rows(Me.dgvMP.SelectedCells(0).RowIndex)

                If CDec(Me.TxtMP_QtyGIT.Text) > 0 Then
                    i.vPrimaryProduct = rowDGV.Cells(Me.ColMP_PrimaryCodeNo.Index).Value
                    i.vUnitCode = rowDGV.Cells(Me.ColMP_Unit.Index).Value
                    i.vDT_ItemGIT = Me.vDT_ItemGIT
                    i.BackColorAll = Me.BackColorAll

                    i.ShowDialog()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ExplainQtyparameter()
        Dim frmExplain As New FrmTrans_MaterialPlanningDashboard_Explain
        Dim arrayBarisAvg As String()
        Dim rowDGV As DataGridViewRow
        Dim rowItem As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item1Row

        Try
            Me.Cursor = Cursors.WaitCursor
            rowDGV = Me.dgvMP.Rows(Me.dgvMP.SelectedCells(0).RowIndex)

            For Each rowItem In Me.vDT.Rows
                If rowDGV.Cells(Me.ColMP_PrimaryCodeNo.Index).Value = rowItem.primary_code_no Then
                    arrayBarisAvg = rowItem.ExplainAvgQty.Split(vbNewLine)
                    If arrayBarisAvg.Length < 9 Then
                        MsgBox("The Explanation of Average Monthly Demand is not correct, it will not be displayed.", MsgBoxStyle.Exclamation, "Exclamation")
                        Exit Sub
                    End If
                    arrayBarisAvg = rowItem.Explain_Safety_Stock.Split(vbNewLine)
                    If arrayBarisAvg.Length < 2 Then
                        MsgBox("The Explanation of Safety Stock is not correct, it will not be displayed.", MsgBoxStyle.Exclamation, "Exclamation")
                        Exit Sub
                    End If
                    arrayBarisAvg = rowItem.Explain_Lead_Time.Split(vbNewLine)
                    If arrayBarisAvg.Length < 2 Then
                        MsgBox("The Explanation of Lead Stock is not correct, it will not be displayed.", MsgBoxStyle.Exclamation, "Exclamation")
                        Exit Sub
                    End If

                    frmExplain.ExplainAvg = rowItem.ExplainAvgQty
                    frmExplain.ExplainSafetyStock = rowItem.Explain_Safety_Stock
                    frmExplain.ExplainLeadTime = rowItem.Explain_Lead_Time
                    frmExplain.ExplainOptimumStock = rowItem.explain_optimum_stock

                    If rowItem.SafetyStockStrategy = Core_Module.clsMasterProduct.EnumSafetyStockStrategy.Based_On_Usage Then
                        frmExplain.pSafateyStock = True
                    End If
                    If rowItem.Lead_Time_Strategy = Core_Module.clsMasterProduct.EnumLeadTimeStrategy.Based_On_Usage Then
                        frmExplain.pLeadTime = True
                    End If
                    If rowItem.optimum_stock_strategy = Core_Module.clsMasterProduct.EnumOptimumStockStrategy.Based_On_Usage Then
                        frmExplain.pOptimumStock = True
                    End If
                    frmExplain.vPrimaryProduct = rowDGV.Cells(Me.ColMP_PrimaryCodeNo.Index).Value
                    frmExplain.ShowDialog()
                    Exit For
                End If
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
            frmExplain.Dispose()
        End Try
    End Sub

    Private Sub TxtMP_QtyLeadStock_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TxtMP_QtyOptimumStock.MouseClick
        Try
            If Me.dgvMP.SelectedCells.Count = 0 Then
                Exit Sub
            Else
                If CDec(Me.TxtMP_QtyOptimumStock.Text) > 0 Then
                    Me.ExplainQtyparameter()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_AvgPerMonth_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TxtMP_AvgPerMonth.MouseClick
        Try
            If Me.dgvMP.SelectedCells.Count = 0 Then
                Exit Sub
            Else
                If CDec(Me.TxtMP_QtyOptimumStock.Text) > 0 Then
                    Me.ExplainQtyparameter()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_QtySafetyStock_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TxtMP_QtySafetyStock.MouseClick
        Try
            If Me.dgvMP.SelectedCells.Count = 0 Then
                Exit Sub
            Else
                If CDec(Me.TxtMP_QtyOptimumStock.Text) > 0 Then
                    Me.ExplainQtyparameter()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LlbMP_UnitPrice_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbMP_UnitPrice.LinkClicked
        Dim ij As New FrmTransaksiPO_HistoricalPrice
        Dim PrimaryCodeNo As Double = -1
        Dim rowDGV As DataGridViewRow

        Try
            If Me.dgvMP.SelectedCells.Count > 0 Then
                rowDGV = Me.dgvMP.Rows(Me.dgvMP.SelectedCells(0).RowIndex)
                PrimaryCodeNo = rowDGV.Cells(Me.ColMP_PrimaryCodeNo.Index).Value

                Me.ObjPO.ExplainHistoricalPrice(PrimaryCodeNo, Me.TBLHistoricalPrice, ij.dgvPrice, Me.ObjCD)
                ij.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub TxtMP_PreferedSupplier_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMP_PreferedSupplier.Enter
        Try
            Me.Buffer = Me.TxtMP_PrimarySupplier.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_PreferedSupplier_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMP_PreferedSupplier.GotFocus
        Try
            If Me.TxtMP_PreferedSupplier.Enabled = True Then
                If Me.TxtMP_PreferedSupplier.ReadOnly = True Then
                    TxtMP_PreferedSupplier.BackColor = Me.TxtReadOnly.BackColor
                Else
                    TxtMP_PreferedSupplier.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_PreferedSupplier_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMP_PreferedSupplier.LostFocus
        Try
            If Me.TxtMP_PreferedSupplier.Enabled = True Then
                If Me.TxtMP_PreferedSupplier.ReadOnly = True Then
                    TxtMP_PreferedSupplier.BackColor = Me.TxtReadOnly.BackColor
                Else
                    TxtMP_PreferedSupplier.BackColor = Me.TxtMP_PrimarySupplier.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnMP_FindSupplier_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMP_FindSupplier.Click
        Dim i As New Core_Module.FrmCariMasterSupplier
        i.BackColorAll = Me.BackColorAll

        Try
            i.cboCriteria.SelectedIndex = 0
            If Me.TxtMP_PreferedSupplier.Text <> "" Then
                i.txtCari.Text = Me.TxtMP_PreferedSupplier.Text
            End If

            i.ShowDialog()
            If i.PrimarySupplier > 0 Then
                Me.TxtMP_PrimarySupplier.Text = i.PrimarySupplier
                Me.TxtMP_PreferedSupplier.Text = Me.ObjSupplier.GetSupplierCode(Me.TxtMP_PrimarySupplier.Text)
                Me.TxtMP_PreferedSupplierDesc.Text = Me.ObjSupplier.GetSupplierName(Me.TxtMP_PrimarySupplier.Text)
            Else
                Me.TxtMP_PrimarySupplier.Text = "-1"
                Me.TxtMP_PreferedSupplier.Clear()
                Me.TxtMP_PreferedSupplierDesc.Clear()
                Me.TxtMP_PreferedSupplier.Focus()
            End If

            Me.UpDateDataSupplier(Me.dgvMP.Item(Me.ColMP_PrimaryCodeNo.Index, Me.dgvMP.SelectedCells(0).RowIndex).Value)

            If Me.TxtMP_PrimarySupplier.Text <> "-1" Then
                If Me.TxtMP_PrimaryCurrency.Text = "-1" Or Me.TxtMP_PrimaryCurrency.Text = "" Then
                    Me.TxtMP_PrimaryCurrency.Text = Me.ObjSupplier.GetPrimaryDefaultCurrency(Me.TxtMP_PrimarySupplier.Text)
                    Me.TxtMP_CurrencyCode.Text = Me.ObjCurrency.GetCurrencyCode(Me.TxtMP_PrimaryCurrency.Text)
                    Me.HitungSubTotal()
                    Me.UpDateDataHarga(Me.dgvMP.Item(Me.ColMP_PrimaryCodeNo.Index, Me.dgvMP.SelectedCells(0).RowIndex).Value)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtMP_PreferedSupplier_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtMP_PreferedSupplier.KeyDown
        Try
            If e.KeyValue = 13 Then
                Dim Primary As Double = ObjSupplier.GetPrimary(Me.TxtMP_PreferedSupplier.Text)
                If Primary <> -1 Then
                    Me.TxtMP_PrimarySupplier.Text = Primary
                    Me.TxtMP_PreferedSupplier.Text = Me.ObjSupplier.GetSupplierCode(Me.TxtMP_PrimarySupplier.Text)
                    Me.TxtMP_PreferedSupplierDesc.Text = Me.ObjSupplier.GetSupplierName(Me.TxtMP_PrimarySupplier.Text)

                    If Me.TxtMP_PrimaryCurrency.Text = "-1" Or Me.TxtMP_PrimaryCurrency.Text = "" Then
                        Me.TxtMP_PrimaryCurrency.Text = Me.ObjSupplier.GetPrimaryDefaultCurrency(Me.TxtMP_PrimarySupplier.Text)
                        Me.TxtMP_CurrencyCode.Text = Me.ObjCurrency.GetCurrencyCode(Me.TxtMP_PrimaryCurrency.Text)
                        Me.HitungSubTotal()
                        Me.UpDateDataHarga(Me.dgvMP.Item(Me.ColMP_PrimaryCodeNo.Index, Me.dgvMP.SelectedCells(0).RowIndex).Value)
                    End If
                Else
                    Me.btnMP_FindSupplier.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_PreferedSupplier_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtMP_PreferedSupplier.Leave
        Try
            If Me.TxtMP_PreferedSupplier.Text <> "" Then
                Dim Primary As Double = ObjSupplier.GetPrimary(Me.TxtMP_PreferedSupplier.Text)
                If Primary <> -1 Then
                    Me.TxtMP_PrimarySupplier.Text = Primary
                    Me.TxtMP_PreferedSupplier.Text = Me.ObjSupplier.GetSupplierCode(Me.TxtMP_PrimarySupplier.Text)
                    Me.TxtMP_PreferedSupplierDesc.Text = Me.ObjSupplier.GetSupplierName(Me.TxtMP_PrimarySupplier.Text)
                Else
                    Me.btnMP_FindSupplier.PerformClick()
                End If
            Else
                Me.TxtMP_PrimarySupplier.Text = "-1"
                Me.TxtMP_PreferedSupplier.Clear()
                Me.TxtMP_PreferedSupplierDesc.Clear()
            End If

            If Me.Buffer <> Me.TxtMP_PrimarySupplier.Text Then
                Me.UpDateDataSupplier(Me.dgvMP.Item(Me.ColMP_PrimaryCodeNo.Index, Me.dgvMP.SelectedCells(0).RowIndex).Value)

                If Me.TxtMP_PrimarySupplier.Text <> "-1" Then
                    If Me.TxtMP_PrimaryCurrency.Text = "-1" Or Me.TxtMP_PrimaryCurrency.Text = "" Then
                        Me.TxtMP_PrimaryCurrency.Text = Me.ObjSupplier.GetPrimaryDefaultCurrency(Me.TxtMP_PrimarySupplier.Text)
                        Me.TxtMP_CurrencyCode.Text = Me.ObjCurrency.GetCurrencyCode(Me.TxtMP_PrimaryCurrency.Text)
                        Me.HitungSubTotal()
                        Me.UpDateDataHarga(Me.dgvMP.Item(Me.ColMP_PrimaryCodeNo.Index, Me.dgvMP.SelectedCells(0).RowIndex).Value)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UpDateDataSupplier(ByVal vPrimaryProduct As Double)
        Dim row As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item1Row

        Try
            For Each row In Me.vDT.Rows
                If row.primary_code_no = vPrimaryProduct Then
                    row.primary_prefered_supplier = Me.TxtMP_PrimarySupplier.Text
                    row.Supplier_Code = Me.TxtMP_PreferedSupplier.Text
                    row.Supplier_Name = Me.TxtMP_PreferedSupplierDesc.Text
                    Exit For
                End If
            Next
            For Each rowdgv As DataGridViewRow In Me.dgvMP.Rows
                If rowdgv.Cells(Me.ColMP_PrimaryCodeNo.Index).Value = vPrimaryProduct Then
                    rowdgv.Cells(Me.ColMP_PrimarySupplier.Index).Value = Me.TxtMP_PrimarySupplier.Text
                    rowdgv.Cells(Me.ColMP_PreferedSupplier.Index).Value = Me.TxtMP_PreferedSupplier.Text
                    Exit For
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_CurrencyCode_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMP_CurrencyCode.Enter
        Try
            Me.Buffer = Me.TxtMP_PrimaryCurrency.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_CurrencyCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMP_CurrencyCode.GotFocus
        Try
            If Me.TxtMP_CurrencyCode.Enabled = True Then
                If Me.TxtMP_CurrencyCode.ReadOnly = True Then
                    TxtMP_CurrencyCode.BackColor = Me.TxtReadOnly.BackColor
                Else
                    TxtMP_CurrencyCode.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_CurrencyCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMP_CurrencyCode.LostFocus
        Try
            If Me.TxtMP_CurrencyCode.Enabled = True Then
                If Me.TxtMP_CurrencyCode.ReadOnly = True Then
                    TxtMP_CurrencyCode.BackColor = Me.TxtReadOnly.BackColor
                Else
                    TxtMP_CurrencyCode.BackColor = Me.TxtMP_PrimaryCurrency.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnMP_FindCurrency_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMP_FindCurrency.Click
        Dim i As New Core_Module.FrmCariMasterCurrency
        i.BackColorAll = Me.BackColorAll

        Try
            i.cboCriteria.SelectedIndex = 0
            If Me.TxtMP_CurrencyCode.Text <> "" Then
                i.txtCari.Text = Me.TxtMP_CurrencyCode.Text
            End If

            i.ShowDialog()
            If i.PrimaryCurrency > 0 Then
                Me.TxtMP_PrimaryCurrency.Text = i.PrimaryCurrency
                Me.TxtMP_CurrencyCode.Text = Me.ObjCurrency.GetCurrencyCode(Me.TxtMP_PrimaryCurrency.Text)
            Else
                If Me.TxtMP_PrimaryCurrency.Text = "" Or Me.TxtMP_PrimaryCurrency.Text = "-1" Then
                    Me.TxtMP_PrimaryCurrency.Text = Core_Module.PrimaryBaseCurrency
                    Me.TxtMP_CurrencyCode.Text = Me.ObjCurrency.GetCurrencyCode(Me.TxtMP_PrimaryCurrency.Text)
                End If

                Me.TxtMP_CurrencyCode.Focus()
            End If

            Me.HitungSubTotal()
            Me.UpDateDataHarga(Me.dgvMP.Item(Me.ColMP_PrimaryCodeNo.Index, Me.dgvMP.SelectedCells(0).RowIndex).Value)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtMP_CurrencyCode_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtMP_CurrencyCode.KeyDown
        Try
            If e.KeyValue = 13 Then
                Dim Primary As Double = Me.ObjCurrency.GetPrimary(Me.TxtMP_CurrencyCode.Text)
                If Primary <> -1 Then
                    Me.TxtMP_PrimaryCurrency.Text = Primary
                    Me.TxtMP_CurrencyCode.Text = Me.ObjCurrency.GetCurrencyCode(Me.TxtMP_PrimaryCurrency.Text)
                Else
                    Me.btnMP_FindCurrency.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_CurrencyCode_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtMP_CurrencyCode.Leave
        Try
            Me.TxtMP_CurrencyCode.Text = Me.TxtMP_CurrencyCode.Text.Trim
            If Me.TxtMP_CurrencyCode.Text <> "" Then
                Dim Primary As Double = Me.ObjCurrency.GetPrimary(Me.TxtMP_CurrencyCode.Text)
                If Primary <> -1 Then
                    Me.TxtMP_PrimaryCurrency.Text = Primary
                    Me.TxtMP_CurrencyCode.Text = Me.ObjCurrency.GetCurrencyCode(Me.TxtMP_PrimaryCurrency.Text)
                Else
                    Me.btnMP_FindCurrency.PerformClick()
                End If
            Else
                Me.TxtMP_PrimaryCurrency.Text = Core_Module.PrimaryBaseCurrency
                Me.TxtMP_CurrencyCode.Text = Me.ObjCurrency.GetCurrencyCode(Me.TxtMP_PrimaryCurrency.Text)
            End If

            If Me.Buffer <> Me.TxtMP_PrimaryCurrency.Text Then
                Me.HitungSubTotal()
                Me.UpDateDataHarga(Me.dgvMP.Item(Me.ColMP_PrimaryCodeNo.Index, Me.dgvMP.SelectedCells(0).RowIndex).Value)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub HitungSubTotal()
        Dim DefPrec As Integer = 0

        Try
            DefPrec = Me.ObjCurrency.GetDefaultPrecission(Me.TxtMP_PrimaryCurrency.Text)
            Me.TxtMP_UnitPrice.Text = Me.ObjCD.DigitGrouping(Me.TxtMP_UnitPrice.Text, DefPrec)

            If Me.TxtMP_DiscFromPerct.Text = "" Then
                Me.TxtMP_DiscFromPerct.Text = "1"
            End If
            If Me.TxtMP_DiscFromPerct.Text = "1" Then
                'hitung disc amount
                Me.TxtMP_DiscAmount.Text = Me.ObjCD.DigitGrouping((CDec(Me.TxtMP_UnitPrice.Text) * CDec(Me.TxtMP_Discount.Text)) / 100, DefPrec)
            Else
                'hitung disc %
                If CDec(Me.TxtMP_UnitPrice.Text) <> 0 Then
                    Me.TxtMP_Discount.Text = Me.ObjCD.DigitGrouping((CDec(Me.TxtMP_DiscAmount.Text) * 100) / CDec(Me.TxtMP_UnitPrice.Text), 2)
                Else
                    Me.TxtMP_Discount.Text = Me.ObjCD.DigitGrouping(0, 2)
                End If
            End If

            Me.TxtMP_NettPrice.Text = Me.ObjCD.DigitGrouping(CDec(Me.TxtMP_UnitPrice.Text) - CDec(Me.TxtMP_DiscAmount.Text), _
            DefPrec)
            Me.TxtMP_SubTotal.Text = Me.ObjCD.DigitGrouping(CDec(Me.TxtMP_NettPrice.Text) * CDec(Me.TxtMP_QtyOrder.Text), DefPrec)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UpDateDataHarga(ByVal vPrimaryProduct As Double)
        Dim row As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item1Row

        Try
            For Each row In Me.vDT.Rows
                If row.primary_code_no = vPrimaryProduct Then
                    row.primary_currency = Me.TxtMP_PrimaryCurrency.Text
                    row.Currency_Code = Me.TxtMP_CurrencyCode.Text
                    row.unit_price = Me.TxtMP_UnitPrice.Text
                    row.def_prec_currency = Me.ObjCurrency.GetDefaultPrecission(Me.TxtMP_PrimaryCurrency.Text)
                    row.unit_price = Me.TxtMP_UnitPrice.Text
                    row.discount = Me.TxtMP_Discount.Text
                    row.nett_price = Me.TxtMP_NettPrice.Text
                    row.sub_total = Me.TxtMP_SubTotal.Text
                    row.disc_from_perct = Me.TxtMP_DiscFromPerct.Text
                    row.disc_amount = Me.TxtMP_DiscAmount.Text
                    Exit For
                End If
            Next
            For Each rowdgv As DataGridViewRow In Me.dgvMP.Rows
                If vPrimaryProduct = rowdgv.Cells(Me.ColMP_PrimaryCodeNo.Index).Value Then
                    rowdgv.Cells(Me.ColMP_PrimaryCurrency.Index).Value = Me.TxtMP_PrimaryCurrency.Text
                    rowdgv.Cells(Me.ColMP_Currency.Index).Value = Me.TxtMP_CurrencyCode.Text
                    rowdgv.Cells(Me.ColMP_UnitPrice.Index).Value = Me.TxtMP_UnitPrice.Text
                    rowdgv.Cells(Me.ColMP_DefPrecCurrency.Index).Value = Me.ObjCurrency.GetDefaultPrecission(Me.TxtMP_PrimaryCurrency.Text)
                    rowdgv.Cells(Me.ColMP_UnitPrice.Index).Value = Me.TxtMP_UnitPrice.Text
                    rowdgv.Cells(Me.ColMP_Disc.Index).Value = Me.TxtMP_Discount.Text
                    rowdgv.Cells(Me.ColMP_NettPrice.Index).Value = Me.TxtMP_NettPrice.Text
                    rowdgv.Cells(Me.ColMP_SubTotal.Index).Value = Me.TxtMP_SubTotal.Text
                    rowdgv.Cells(Me.ColMP_DiscFromPerct.Index).Value = Me.TxtMP_DiscFromPerct.Text
                    rowdgv.Cells(Me.ColMP_DiscAmount.Index).Value = Me.TxtMP_DiscAmount.Text

                    Exit For
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_UnitPrice_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMP_UnitPrice.Enter
        Try
            Me.Buffer = TxtMP_UnitPrice.Text
            Me.ObjCD.Number_Enter(Me.TxtMP_UnitPrice, sender, e)
            If Me.TxtMP_PrimaryCurrency.Text = "" Or Me.TxtMP_PrimaryCurrency.Text = "-1" Or Me.TxtMP_CurrencyCode.Text = "" Then
                Me.TxtMP_PrimaryCurrency.Text = Core_Module.PrimaryBaseCurrency
                Me.TxtMP_CurrencyCode.Text = Me.ObjCurrency.GetCurrencyCode(Me.TxtMP_PrimaryCurrency.Text)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_UnitPrice_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMP_UnitPrice.GotFocus
        Try
            If Me.TxtMP_UnitPrice.Enabled = True Then
                If Me.TxtMP_UnitPrice.ReadOnly = True Then
                    TxtMP_UnitPrice.BackColor = Me.TxtReadOnly.BackColor
                Else
                    TxtMP_UnitPrice.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_UnitPrice_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtMP_UnitPrice.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtMP_UnitPrice, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_UnitPrice_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtMP_UnitPrice.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtMP_UnitPrice, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_UnitPrice_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMP_UnitPrice.Leave
        Try
            Me.ObjCD.Number_Leave(Me.TxtMP_UnitPrice, sender, e, Me.ObjCurrency.GetDefaultPrecission(Me.TxtMP_PrimaryCurrency.Text), Me.ObjCD)
            'jika menggunakan disc amount, maka unit price harus selalu lebih besar dari disc amount
            If CDec(Me.Buffer) <> CDec(Me.TxtMP_UnitPrice.Text) Then
                If Me.TxtMP_DiscFromPerct.Text = "0" Then
                    If CDec(Me.TxtMP_DiscAmount.Text) > CDec(Me.TxtMP_UnitPrice.Text) Then
                        Me.TxtMP_UnitPrice.Text = Me.Buffer
                        Exit Sub
                    End If
                End If

                Me.HitungSubTotal()
                Me.UpDateDataHarga(Me.dgvMP.Item(Me.ColMP_PrimaryCodeNo.Index, Me.dgvMP.SelectedCells(0).RowIndex).Value)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_UnitPrice_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMP_UnitPrice.LostFocus
        Try
            If Me.TxtMP_UnitPrice.Enabled = True Then
                If Me.TxtMP_UnitPrice.ReadOnly = True Then
                    TxtMP_UnitPrice.BackColor = Me.TxtReadOnly.BackColor
                Else
                    TxtMP_UnitPrice.BackColor = Me.TxtMP_PrimaryCurrency.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_Discount_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMP_Discount.Enter
        Try
            Me.Buffer = Me.TxtMP_Discount.Text
            Me.ObjCD.Number_Enter(Me.TxtMP_Discount, sender, e)
            If Me.TxtMP_PrimaryCurrency.Text = "" Or Me.TxtMP_PrimaryCurrency.Text = "-1" Or Me.TxtMP_CurrencyCode.Text = "" Then
                Me.TxtMP_PrimaryCurrency.Text = Core_Module.PrimaryBaseCurrency
                Me.TxtMP_CurrencyCode.Text = Me.ObjCurrency.GetCurrencyCode(Me.TxtMP_PrimaryCurrency.Text)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_Discount_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMP_Discount.GotFocus
        Try
            If Me.TxtMP_Discount.Enabled = True Then
                If Me.TxtMP_Discount.ReadOnly = True Then
                    TxtMP_Discount.BackColor = Me.TxtReadOnly.BackColor
                Else
                    TxtMP_Discount.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_Discount_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtMP_Discount.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtMP_Discount, sender, e, 100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_Discount_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtMP_Discount.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtMP_Discount, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_Discount_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMP_Discount.Leave
        Try
            Me.ObjCD.Number_Leave(Me.TxtMP_Discount, sender, e, 2, Me.ObjCD)
            If CDec(Me.TxtMP_Discount.Text) <> CDec(Me.Buffer) Then
                Me.TxtMP_DiscFromPerct.Text = "1"
                Me.HitungSubTotal()
                Me.UpDateDataHarga(Me.dgvMP.Item(Me.ColMP_PrimaryCodeNo.Index, Me.dgvMP.SelectedCells(0).RowIndex).Value)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_Discount_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMP_Discount.LostFocus
        Try
            If Me.TxtMP_Discount.Enabled = True Then
                If Me.TxtMP_Discount.ReadOnly = True Then
                    TxtMP_Discount.BackColor = Me.TxtReadOnly.BackColor
                Else
                    TxtMP_Discount.BackColor = Me.TxtMP_PrimaryCurrency.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_DiscAmount_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMP_DiscAmount.Enter
        Try
            Me.Buffer = Me.TxtMP_DiscAmount.Text
            Me.ObjCD.Number_Enter(Me.TxtMP_DiscAmount, sender, e)
            If Me.TxtMP_PrimaryCurrency.Text = "" Or Me.TxtMP_PrimaryCurrency.Text = "-1" Or Me.TxtMP_CurrencyCode.Text = "" Then
                Me.TxtMP_PrimaryCurrency.Text = Core_Module.PrimaryBaseCurrency
                Me.TxtMP_CurrencyCode.Text = Me.ObjCurrency.GetCurrencyCode(Me.TxtMP_PrimaryCurrency.Text)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_DiscAmount_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMP_DiscAmount.GotFocus
        Try
            If Me.TxtMP_DiscAmount.Enabled = True Then
                If Me.TxtMP_DiscAmount.ReadOnly = True Then
                    TxtMP_DiscAmount.BackColor = Me.TxtReadOnly.BackColor
                Else
                    TxtMP_DiscAmount.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_DiscAmount_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtMP_DiscAmount.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtMP_DiscAmount, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_DiscAmount_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtMP_DiscAmount.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtMP_DiscAmount, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_DiscAmount_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMP_DiscAmount.Leave
        Try
            Me.ObjCD.Number_Leave(Me.TxtMP_DiscAmount, sender, e, Me.ObjCurrency.GetDefaultPrecission(Me.TxtMP_PrimaryCurrency.Text), Me.ObjCD)
            If CDec(Me.TxtMP_DiscAmount.Text) > CDec(Me.TxtMP_UnitPrice.Text) Then
                Me.TxtMP_DiscAmount.Text = Me.Buffer
                Exit Sub
            End If
            If CDec(Me.TxtMP_DiscAmount.Text) <> CDec(Me.Buffer) Then
                Me.TxtMP_DiscFromPerct.Text = "0"
                Me.HitungSubTotal()
                Me.UpDateDataHarga(Me.dgvMP.Item(Me.ColMP_PrimaryCodeNo.Index, Me.dgvMP.SelectedCells(0).RowIndex).Value)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_DiscAmount_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMP_DiscAmount.LostFocus
        Try
            If Me.TxtMP_DiscAmount.Enabled = True Then
                If Me.TxtMP_DiscAmount.ReadOnly = True Then
                    TxtMP_DiscAmount.BackColor = Me.TxtReadOnly.BackColor
                Else
                    TxtMP_DiscAmount.BackColor = Me.TxtMP_PrimaryCurrency.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_QtyOrder_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMP_QtyOrder.Enter
        Try
            Me.Buffer = Me.TxtMP_QtyOrder.Text
            Me.ObjCD.Number_Enter(Me.TxtMP_QtyOrder, sender, e)
            If Me.TxtMP_PrimaryCurrency.Text = "" Or Me.TxtMP_PrimaryCurrency.Text = "-1" Or Me.TxtMP_CurrencyCode.Text = "" Then
                Me.TxtMP_PrimaryCurrency.Text = Core_Module.PrimaryBaseCurrency
                Me.TxtMP_CurrencyCode.Text = Me.ObjCurrency.GetCurrencyCode(Me.TxtMP_PrimaryCurrency.Text)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_QtyOrder_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMP_QtyOrder.GotFocus
        Try
            If Me.TxtMP_QtyOrder.Enabled = True Then
                If Me.TxtMP_QtyOrder.ReadOnly = True Then
                    TxtMP_QtyOrder.BackColor = Me.TxtReadOnly.BackColor
                Else
                    TxtMP_QtyOrder.BackColor = Me.BackColorAll
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_QtyOrder_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtMP_QtyOrder.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtMP_QtyOrder, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_QtyOrder_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtMP_QtyOrder.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtMP_QtyOrder, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_QtyOrder_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMP_QtyOrder.Leave
        Try
            'update data order
            If CDec(Me.TxtMP_QtyOrder.Text) <> CDec(Me.Buffer) Then
                For Each rowItem As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item1Row In Me.vDT.Rows
                    If rowItem.primary_code_no = Me.dgvMP.Item(Me.ColMP_PrimaryCodeNo.Index, Me.dgvMP.SelectedCells(0).RowIndex).Value Then
                        Me.ObjCD.Number_Leave(Me.TxtMP_QtyOrder, sender, e, rowItem.def_prec_unit_multi, Me.ObjCD)

                        'hitung ulang sub total + update k row
                        Me.HitungSubTotal()
                        Me.UpDateDataHarga(Me.dgvMP.Item(Me.ColMP_PrimaryCodeNo.Index, Me.dgvMP.SelectedCells(0).RowIndex).Value)

                        'jangan lupa update k vdt
                        With rowItem
                            .qty_order_multiunit = Me.TxtMP_QtyOrder.Text
                            .qty_order = Math.Round((.qty_order_multiunit * .ratio_default) / .ratio_multi, .default_prec_unit, MidpointRounding.AwayFromZero)

                            'digit grouping
                            Me.dgvMP.Item(Me.ColMP_QtyOrder.Index, Me.dgvMP.SelectedCells(0).RowIndex).Value = Me.ObjCD.DigitGrouping(.qty_order, .default_prec_unit)
                            Me.dgvMP.Item(Me.ColMP_QtyOrderMulti.Index, Me.dgvMP.SelectedCells(0).RowIndex).Value = Me.ObjCD.DigitGrouping(.qty_order_multiunit, .def_prec_unit_multi)

                            If rowItem.procurement_strategy = Core_Module.clsMasterProduct.EnumProcurementStrategy.Build_To_Stock Then
                                Me.ObjCls.HitungOrderPoint(rowItem, Me.TxtMP_DaysPerMonth.Text, Me.dtpMP_RefDate, Me.ObjCD)
                                Me.dgvMP.Item(Me.ColMP_NextOrderPointDelay.Index, Me.dgvMP.SelectedCells(0).RowIndex).Value = .next_order_point_delay
                                Me.dgvMP.Item(Me.ColMP_NextOrderPoint.Index, Me.dgvMP.SelectedCells(0).RowIndex).Value = .next_order_point
                            End If
                        End With

                        Exit Sub
                    End If
                Next
            Else
                Me.TxtMP_QtyOrder.Text = Me.Buffer
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_QtyOrder_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMP_QtyOrder.LostFocus
        Try
            If Me.TxtMP_QtyOrder.Enabled = True Then
                If Me.TxtMP_QtyOrder.ReadOnly = True Then
                    TxtMP_QtyOrder.BackColor = Me.TxtReadOnly.BackColor
                Else
                    TxtMP_QtyOrder.BackColor = Me.TxtMP_PrimaryCurrency.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpMP_RefDate_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpMP_RefDate.Enter
        Try
            If Me.TxtMP_Remarks.Enabled = True Then
                Me.Buffer = Me.dtpMP_RefDate.Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpMP_RefDate_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpMP_RefDate.Leave
        Try
            If CDate(Me.Buffer) <> Me.dtpMP_RefDate.Value.Date Then
                If Me.vDT.Rows.Count > 0 Then
                    MsgBox("Can't change reference date on calculated data. Please change reference data before calculation", MsgBoxStyle.Information, "Information")
                Else
                    Me.bwRefDate.RunWorkerAsync()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bwRefDate_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bwRefDate.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub bwRefDate_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bwRefDate.RunWorkerCompleted
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For Each row As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item1Row In Me.vDT.Rows
                Me.ObjCls.HitungOrderPoint(row, Me.TxtMP_DaysPerMonth.Text, Me.dtpMP_RefDate, Me.ObjCD)
            Next
            Me.LoadDGV()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LoadUlang()
        Try
            If Me.TabControl1.SelectedIndex = 0 Then
                Me.LoadDGV()
            Else
                Me.LoadDGV_SF()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbMP_Filter_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbMP_Filter.LinkClicked
        Me.LoadUlang()
    End Sub

    Private Sub btnFilter_FindPreferedSupplier_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter_FindPreferedSupplier.Click
        Dim frm As New Core_Module.FrmCariMasterSupplier

        Try
            If Me.TxtFilter_PreferedSupplier.Text <> "" Then
                frm.txtCari.Text = Me.TxtFilter_PreferedSupplier.Text
            End If
            frm.ShowDialog()
            If frm.PrimarySupplier > 0 Then
                Me.TxtFilter_PrimaryPreferedSupplier.Text = frm.PrimarySupplier
                Me.TxtFilter_PreferedSupplier.Text = Me.ObjSupplier.GetSupplierCode(frm.PrimarySupplier)
                Me.TxtFilter_PreferedSupplierDesc.Text = Me.ObjSupplier.GetSupplierName(frm.PrimarySupplier)
            Else
                Me.TxtFilter_PrimaryPreferedSupplier.Text = "-1"
                Me.TxtFilter_PreferedSupplier.Clear()
                Me.TxtFilter_PreferedSupplierDesc.Clear()
            End If

            Me.LoadUlang()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            frm.Dispose()
        End Try
    End Sub

    Private Sub TxtFilter_PreferedSupplier_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtFilter_PreferedSupplier.KeyDown
        Try
            Me.Cursor = Cursors.WaitCursor
            If e.KeyCode = 13 Then
                If Me.TxtFilter_PreferedSupplier.Text <> "" Then
                    If Me.ObjSupplier.ApakahSudahAda(Me.TxtFilter_PreferedSupplier.Text) = True Then
                        Me.TxtFilter_PrimaryPreferedSupplier.Text = Me.ObjSupplier.GetPrimary(Me.TxtFilter_PreferedSupplier.Text)
                        Me.TxtFilter_PreferedSupplier.Text = Me.ObjSupplier.GetSupplierCode(Me.TxtFilter_PrimaryPreferedSupplier.Text)
                        Me.TxtFilter_PreferedSupplierDesc.Text = Me.ObjSupplier.GetSupplierName(Me.TxtFilter_PrimaryPreferedSupplier.Text)
                    Else
                        Me.btnFilter_FindPreferedSupplier.PerformClick()
                    End If
                Else
                    Me.TxtFilter_PrimaryPreferedSupplier.Text = "-1"
                    Me.TxtFilter_PreferedSupplier.Clear()
                    Me.TxtFilter_PreferedSupplierDesc.Clear()
                End If
                Me.LoadUlang()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub TxtFilter_PreferedSupplier_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtFilter_PreferedSupplier.Leave
        Try
            Me.Cursor = Cursors.WaitCursor
            If Me.TxtFilter_PreferedSupplier.Text <> "" Then
                If Me.ObjSupplier.ApakahSudahAda(Me.TxtFilter_PreferedSupplier.Text) = True Then
                    Me.TxtFilter_PrimaryPreferedSupplier.Text = Me.ObjSupplier.GetPrimary(Me.TxtFilter_PreferedSupplier.Text)
                    Me.TxtFilter_PreferedSupplier.Text = Me.ObjSupplier.GetSupplierCode(Me.TxtFilter_PrimaryPreferedSupplier.Text)
                    Me.TxtFilter_PreferedSupplierDesc.Text = Me.ObjSupplier.GetSupplierName(Me.TxtFilter_PrimaryPreferedSupplier.Text)
                Else
                    Me.btnFilter_FindPreferedSupplier.PerformClick()
                End If
            Else
                Me.TxtFilter_PrimaryPreferedSupplier.Text = "-1"
                Me.TxtFilter_PreferedSupplier.Clear()
                Me.TxtFilter_PreferedSupplierDesc.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub TxtFilter_ProductGroupCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtFilter_ProductGroupCode.KeyDown
        Dim enumAktif As FrmTrans_MaterialPlanningDashboard_Data.EnumProductGroupLevel
        Try
            Me.Cursor = Cursors.WaitCursor
            If e.KeyCode = 13 Then
                If Me.cboFilter_ProductGroupLevel.SelectedIndex <> 3 Then
                    If Me.cboFilter_ProductGroupLevel.SelectedIndex = 0 Then
                        enumAktif = FrmTrans_MaterialPlanningDashboard_Data.EnumProductGroupLevel.ProductGroup
                    ElseIf Me.cboFilter_ProductGroupLevel.SelectedIndex = 1 Then
                        enumAktif = FrmTrans_MaterialPlanningDashboard_Data.EnumProductGroupLevel.Brand
                    ElseIf Me.cboFilter_ProductGroupLevel.SelectedIndex = 2 Then
                        enumAktif = FrmTrans_MaterialPlanningDashboard_Data.EnumProductGroupLevel.SubBrand
                    End If
                    If Me.TxtFilter_ProductGroupCode.Text <> "" Then
                        If Me.objPG.ApakahSudahAda(Me.TxtFilter_ProductGroupCode.Text, enumAktif) = True Then
                            Me.TxtFilter_PrimaryProductGroup.Text = Me.objPG.GetPrimary(Me.TxtFilter_ProductGroupCode.Text, enumAktif)
                            Me.TxtFilter_ProductGroupLevelDesc.Text = Me.objPG.GetGroupDesc(Me.TxtFilter_PrimaryProductGroup.Text)
                        Else
                            Me.btnFilter_FindProductGroupLevel.PerformClick()
                        End If
                    Else
                        Me.TxtFilter_PrimaryProductGroup.Text = "-1"
                        Me.TxtFilter_ProductGroupLevelDesc.Clear()
                    End If
                Else
                    If Me.TxtFilter_ProductGroupCode.Text <> "" Then
                        If Me.objCategory.ApakahSudahAda(Me.TxtFilter_ProductGroupCode.Text) = True Then
                            Me.TxtFilter_PrimaryProductGroup.Text = Me.objCategory.GetPrimary(Me.TxtFilter_ProductGroupCode.Text)
                            Me.TxtFilter_ProductGroupLevelDesc.Text = Me.objCategory.GetDesc(Me.TxtFilter_PrimaryProductGroup.Text)
                        Else
                            Me.btnFilter_FindProductGroupLevel.PerformClick()
                        End If
                    Else
                        Me.TxtFilter_PrimaryProductGroup.Text = "-1"
                        Me.TxtFilter_ProductGroupLevelDesc.Clear()
                    End If
                End If

                Me.LoadUlang()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub btnFilter_FindProductGroupLevel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter_FindProductGroupLevel.Click
        Dim frm As New Core_Module.FrmCariMasterProductGroup
        If Me.cboFilter_ProductGroupLevel.SelectedIndex = 0 Then
            frm.ProductGroupLevel = FrmTrans_MaterialPlanningDashboard_Data.EnumProductGroupLevel.ProductGroup
        ElseIf Me.cboFilter_ProductGroupLevel.SelectedIndex = 1 Then
            frm.ProductGroupLevel = FrmTrans_MaterialPlanningDashboard_Data.EnumProductGroupLevel.Brand
        ElseIf Me.cboFilter_ProductGroupLevel.SelectedIndex = 2 Then
            frm.ProductGroupLevel = FrmTrans_MaterialPlanningDashboard_Data.EnumProductGroupLevel.SubBrand
        ElseIf Me.cboFilter_ProductGroupLevel.SelectedIndex = 3 Then
            frm.ProductGroupLevel = FrmTrans_MaterialPlanningDashboard_Data.EnumProductGroupLevel.Category
        End If
        Dim frmCat As New Core_Module.frmCariMasterCategory

        Try
            Me.Cursor = Cursors.WaitCursor
            If Me.cboFilter_ProductGroupLevel.SelectedIndex <> 3 Then
                If Me.TxtFilter_ProductGroupCode.Text <> "" Then
                    frm.txtCari.Text = Me.TxtFilter_ProductGroupCode.Text
                End If
                frm.ShowDialog()
                If frm.PrimaryProductGroup > 0 Then
                    Me.TxtFilter_PrimaryProductGroup.Text = frm.PrimaryProductGroup
                    Me.TxtFilter_ProductGroupCode.Text = Me.objPG.GetGroupCode(frm.PrimaryProductGroup)
                    Me.TxtFilter_ProductGroupLevelDesc.Text = Me.objPG.GetGroupDesc(frm.PrimaryProductGroup)
                Else
                    Me.TxtFilter_PrimaryProductGroup.Text = "-1"
                    Me.TxtFilter_ProductGroupLevelDesc.Clear()
                    Me.TxtFilter_ProductGroupCode.Clear()
                End If
            Else
                If Me.TxtFilter_ProductGroupCode.Text <> "" Then
                    frmCat.txtCari.Text = Me.TxtFilter_ProductGroupCode.Text
                End If
                frmCat.ShowDialog()
                If frmCat.PrimaryCategory > 0 Then
                    Me.TxtFilter_PrimaryProductGroup.Text = frmCat.PrimaryCategory
                    Me.TxtFilter_ProductGroupCode.Text = Me.objCategory.GetCategoryCode(frmCat.PrimaryCategory)
                    Me.TxtFilter_ProductGroupLevelDesc.Text = Me.objCategory.GetDesc(frmCat.PrimaryCategory)
                Else
                    Me.TxtFilter_PrimaryProductGroup.Text = "-1"
                    Me.TxtFilter_ProductGroupCode.Clear()
                    Me.TxtFilter_ProductGroupLevelDesc.Clear()
                End If
            End If

            Me.LoadUlang()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            frm.Dispose()
            frmCat.Dispose()
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub cboFilter_ProductGroupLevel_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboFilter_ProductGroupLevel.SelectedIndexChanged
        Try
            Me.TxtFilter_PrimaryProductGroup.Text = "-1"
            Me.TxtFilter_ProductGroupCode.Clear()
            Me.TxtFilter_ProductGroupLevelDesc.Clear()

            If Me.cboFilter_ProductGroupLevel.SelectedIndex = 0 Then
                Me.LblFilter_PoductGroupDesc.Text = "Group Code"
            ElseIf Me.cboFilter_ProductGroupLevel.SelectedIndex = 1 Then
                Me.LblFilter_PoductGroupDesc.Text = "Brand Code"
            ElseIf Me.cboFilter_ProductGroupLevel.SelectedIndex = 2 Then
                Me.LblFilter_PoductGroupDesc.Text = "Sub Brand Code"
            ElseIf Me.cboFilter_ProductGroupLevel.SelectedIndex = 3 Then
                Me.LblFilter_PoductGroupDesc.Text = "Category Code"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Sub

    Private Sub TxtFilter_ProductGroupCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtFilter_ProductGroupCode.Leave
        Dim enumAktif As FrmTrans_MaterialPlanningDashboard_Data.EnumProductGroupLevel
        Try
            Me.Cursor = Cursors.WaitCursor
            If Me.cboFilter_ProductGroupLevel.SelectedIndex <> 3 Then
                If Me.cboFilter_ProductGroupLevel.SelectedIndex = 0 Then
                    enumAktif = FrmTrans_MaterialPlanningDashboard_Data.EnumProductGroupLevel.ProductGroup
                ElseIf Me.cboFilter_ProductGroupLevel.SelectedIndex = 1 Then
                    enumAktif = FrmTrans_MaterialPlanningDashboard_Data.EnumProductGroupLevel.Brand
                ElseIf Me.cboFilter_ProductGroupLevel.SelectedIndex = 2 Then
                    enumAktif = FrmTrans_MaterialPlanningDashboard_Data.EnumProductGroupLevel.SubBrand
                End If
                If Me.TxtFilter_ProductGroupCode.Text <> "" Then
                    If Me.objPG.ApakahSudahAda(Me.TxtFilter_ProductGroupCode.Text, enumAktif) = True Then
                        Me.TxtFilter_PrimaryProductGroup.Text = Me.objPG.GetPrimary(Me.TxtFilter_ProductGroupCode.Text, enumAktif)
                        Me.TxtFilter_ProductGroupLevelDesc.Text = Me.objPG.GetGroupDesc(Me.TxtFilter_PrimaryProductGroup.Text)
                    Else
                        Me.btnFilter_FindProductGroupLevel.PerformClick()
                    End If
                Else
                    Me.TxtFilter_PrimaryProductGroup.Text = "-1"
                    Me.TxtFilter_ProductGroupLevelDesc.Clear()
                End If
            Else
                If Me.TxtFilter_ProductGroupCode.Text <> "" Then
                    If Me.objCategory.ApakahSudahAda(Me.TxtFilter_ProductGroupCode.Text) = True Then
                        Me.TxtFilter_PrimaryProductGroup.Text = Me.objCategory.GetPrimary(Me.TxtFilter_ProductGroupCode.Text)
                        Me.TxtFilter_ProductGroupLevelDesc.Text = Me.objCategory.GetDesc(Me.TxtFilter_PrimaryProductGroup.Text)
                    Else
                        Me.btnFilter_FindProductGroupLevel.PerformClick()
                    End If
                Else
                    Me.TxtFilter_PrimaryProductGroup.Text = "-1"
                    Me.TxtFilter_ProductGroupLevelDesc.Clear()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub llbMP_AddToPO_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles llbMP_AddToPO.Click
        Dim rowDGV As DataGridViewRow
        Dim AdaModify As Boolean = False
        Dim Dummy As String = ""

        Try
            For Each rowDGV In Me.dgvMP.Rows
                If rowDGV.Cells(Me.ColMP_FollowUp.Index).Value = Boolean.TrueString Then
                    If CDec(rowDGV.Cells(Me.ColMP_QtyOrder.Index).Value) > 0 Then
                        AdaModify = True
                        Exit For
                    End If
                End If
            Next

            If AdaModify = True Then

                'cek product code apa aja dalam dgv yang di follow up tetapi datanya belum lengkap
                '1. cek supplier nya belum dipilih
                Dummy = ""
                For Each rowDGV In Me.dgvMP.Rows
                    If rowDGV.Cells(Me.ColMP_FollowUp.Index).Value = Boolean.TrueString And rowDGV.Cells(Me.ColMP_ProcurementType.Index).Value = "Buy" Then
                        If rowDGV.Cells(Me.ColMP_PrimarySupplier.Index).Value = -1 Then
                            If Dummy = "" Then
                                Dummy = rowDGV.Cells(Me.ColMP_CodeNo.Index).Value
                            Else
                                Dummy = Dummy + ", " + rowDGV.Cells(Me.ColMP_CodeNo.Index).Value
                            End If
                        End If
                    End If
                Next
                If Dummy <> "" Then
                    MsgBox("The following item have unspecified supplier: " + Dummy + ". Please choose supplier.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                '2. unit price = 0 
                Dummy = ""
                For Each rowDGV In Me.dgvMP.Rows
                    If rowDGV.Cells(Me.ColMP_FollowUp.Index).Value = Boolean.TrueString And rowDGV.Cells(Me.ColMP_ProcurementType.Index).Value = "Buy" Then
                        If rowDGV.Cells(Me.ColMP_UnitPrice.Index).Value = 0 Then
                            If Dummy = "" Then
                                Dummy = rowDGV.Cells(Me.ColMP_CodeNo.Index).Value
                            Else
                                Dummy = Dummy + ", " + rowDGV.Cells(Me.ColMP_CodeNo.Index).Value
                            End If
                        End If
                    End If
                Next
                If Dummy <> "" Then
                    MsgBox("The following item have unspecified unit price: " + Dummy + ". Please enter positive number for unit price data.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                '3. currency tidak ada
                Dummy = ""
                For Each rowDGV In Me.dgvMP.Rows
                    If rowDGV.Cells(Me.ColMP_FollowUp.Index).Value = Boolean.TrueString And rowDGV.Cells(Me.ColMP_ProcurementType.Index).Value = "Buy" Then
                        If rowDGV.Cells(Me.ColMP_PrimaryCurrency.Index).Value = -1 Then
                            If Dummy = "" Then
                                Dummy = rowDGV.Cells(Me.ColMP_CodeNo.Index).Value
                            Else
                                Dummy = Dummy + ", " + rowDGV.Cells(Me.ColMP_CodeNo.Index).Value
                            End If
                        End If
                    End If
                Next
                If Dummy <> "" Then
                    MsgBox("The following item have unspecified currency: " + Dummy + ". Please choose currency.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                '4. harus ada bom, untuk yang di produce
                Dummy = ""
                For Each rowDGV In Me.dgvMP.Rows
                    If rowDGV.Cells(Me.ColMP_FollowUp.Index).Value = Boolean.TrueString And rowDGV.Cells(Me.ColMP_ProcurementType.Index).Value = "Produce" Then
                        If ObjMO.GetJmlItemBOM(rowDGV.Cells(Me.ColMP_PrimaryCodeNo.Index).Value) <= 0 Then
                            If Dummy = "" Then
                                Dummy = rowDGV.Cells(Me.ColMP_CodeNo.Index).Value
                            Else
                                Dummy = Dummy + ", " + rowDGV.Cells(Me.ColMP_CodeNo.Index).Value
                            End If
                        End If
                    End If
                Next
                If Dummy <> "" Then
                    MsgBox("The following produce item have unspecified bill of material: " + Dummy + ". Please specify bill of material.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                '4.b. harus goods, untuk yang di produce
                Dummy = ""
                For Each rowDGV In Me.dgvMP.Rows
                    If rowDGV.Cells(Me.ColMP_FollowUp.Index).Value = Boolean.TrueString And rowDGV.Cells(Me.ColMP_ProcurementType.Index).Value = "Produce" Then
                        If ObjMO.GetTypeAccProduct(rowDGV.Cells(Me.ColMP_PrimaryCodeNo.Index).Value) <> 0 Then
                            If Dummy = "" Then
                                Dummy = rowDGV.Cells(Me.ColMP_CodeNo.Index).Value
                            Else
                                Dummy = Dummy + ", " + rowDGV.Cells(Me.ColMP_CodeNo.Index).Value
                            End If
                        End If
                    End If
                Next
                If Dummy <> "" Then
                    MsgBox("The following produce item is not goods: " + Dummy + ". We can only produce goods.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                'add to po
                If Me.ApakahAdaBuy = True Then
                    If Me.IsAllowedToCreateOrAddPO(Me.TxtMP_Primary.Text) Then
                        Me.AddToPO()
                    End If
                End If
                'add to mo
                If Me.ApakahAdaProduce = True Then
                    If Me.IsAllowedToCreateOrAddMO(Me.TxtMP_Primary.Text) Then
                        Me.AddToMO()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AddToMO()
        Dim TBLItemProduct As New DseTrans_ManufacturingOrder.tblitem_productDataTable
        Dim rowItemProduct As DseTrans_ManufacturingOrder.tblitem_productRow
        Dim rowDGV As DataGridViewRow
        Dim ij As New FrmTrans_ManufacturingOrder_Cari
        Dim listOfOldIdentifierJO As New List(Of String)
        Try
            ij.Text = "Please specify MO to add"
            ij.ShowDialog()
            If ij.Primary <> -1 Then
                'pastikan mo belum ready for scheduling
                If Me.ObjMO.ApakahSudahDiOtorisasi(ij.Primary) = True Then
                    MsgBox("The manufacturing order: " + Me.ObjMO.GetNoTransaksi(ij.Primary) + _
                    " is ready for scheduling. Please choose another manufacturing order.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
                'pastikan boleh edit mo
                If Me.ObjMO.ApakahBolehEdit(ij.Primary, Me.UserCode, Me.ObjMO.GetNoTransaksi(ij.Primary), True) = False Then
                    Exit Sub
                End If

                TBLItemProduct.Rows.Clear()
                For Each rowDGV In Me.dgvMP.Rows
                    If rowDGV.Cells(Me.ColMP_FollowUp.Index).Value = Boolean.TrueString Then
                        If rowDGV.Cells(Me.ColMP_ProcurementType.Index).Value = "Produce" And CDec(rowDGV.Cells(Me.ColMP_QtyOrder.Index).Value) Then
                            rowItemProduct = TBLItemProduct.Newtblitem_productRow
                            rowItemProduct.Qty_Target = CDec(rowDGV.Cells(Me.ColMP_QtyOrder.Index).Value)
                            rowItemProduct.Primary_Product = rowDGV.Cells(Me.ColMP_PrimaryCodeNo.Index).Value
                            TBLItemProduct.Rows.Add(rowItemProduct)
                        End If
                    End If
                Next

                Dim i As New FrmTrans_ManufacturingOrder
                i.UserCode = Me.UserCode
                i.BackColorAll = Me.BackColorAll
                i.NewMOFromMP = False
                i.FromMP = True
                i.TBLItemProduct = TBLItemProduct
                i.PrimaryMOFromMP = ij.Primary
                listOfOldIdentifierJO = Me.ObjMO.GetListOfIdentifierJO(ij.Primary)
                i.ShowDialog()

                If i.DialogResult = Windows.Forms.DialogResult.OK Then
                    Dim oldDefaultQty As Decimal = 0
                    Dim oldMultiQty As Decimal = 0
                    For Each rowDgvItemMO As DataGridViewRow In i.DgvItem.Rows
                        If rowDgvItemMO.Cells(i.ColItem_PrimaryProduct.Index).Value <> -1 And Not listOfOldIdentifierJO.Contains(rowDgvItemMO.Cells(i.ColItem_Identifier.Index).Value) Then
                            For Each rowDgvItemMP As DataGridViewRow In Me.dgvMP.Rows
                                If rowDgvItemMO.Cells(i.ColItem_PrimaryProduct.Index).Value = rowDgvItemMP.Cells(Me.ColMP_PrimaryCodeNo.Index).Value And rowDgvItemMP.Cells(Me.ColMP_FollowUp.Index).Value = True And rowDgvItemMP.Cells(Me.ColMP_ProcurementType.Index).Value = "Produce" Then
                                    Me.vRowItemFollowUp = Me.vDT_ItemFollowUp.Newtbltrans_material_planning_item_followupRow
                                    Me.vRowItemFollowUp.primary_transaksi = Me.TxtMP_Primary.Text
                                    Me.vRowItemFollowUp.primary_product = rowDgvItemMO.Cells(i.ColItem_PrimaryProduct.Index).Value
                                    Me.vRowItemFollowUp.tipe = ClsTransaksiMaterialPlanning.EnumActionType.PRODUCE_ADDTO_MO
                                    Me.vRowItemFollowUp.primary_po_mo = i.TxtPrimary.Text
                                    Me.vRowItemFollowUp.no_po_mo = i.TxtNoTransaksi.Text
                                    Me.vRowItemFollowUp.tanggal_po_mo = i.DtpTglTransaksi.Value
                                    Me.vRowItemFollowUp.identifier_jo = rowDgvItemMO.Cells(i.ColItem_Identifier.Index).Value
                                    Me.vRowItemFollowUp.no_jo = rowDgvItemMO.Cells(i.ColItem_JONumber.Index).Value
                                    Me.vRowItemFollowUp.qty_default = rowDgvItemMO.Cells(i.ColItem_QtyTarget.Index).Value
                                    Me.vRowItemFollowUp.primary_unit_default = objUnit.GetPrimary(rowDgvItemMO.Cells(i.ColItem_Unit.Index).Value)
                                    Me.vRowItemFollowUp.ratio_default = 1
                                    Me.vRowItemFollowUp.qty_multi = Me.vRowItemFollowUp.qty_default
                                    Me.vRowItemFollowUp.primary_unit_multi = Me.vRowItemFollowUp.primary_unit_default
                                    Me.vRowItemFollowUp.ratio_multi = 1
                                    Me.vRowItemFollowUp.user_code = Me.UserCode
                                    Me.vRowItemFollowUp.followup_time = Now
                                    Me.vDT_ItemFollowUp.Rows.Add(Me.vRowItemFollowUp)
                                    Me.SaveRequired = True
                                    Exit For
                                End If
                            Next
                        End If
                    Next
                    For Each rowDgvItemMP As DataGridViewRow In Me.dgvMP.Rows
                        Me.vDT_ItemFollowUp.DefaultView.RowFilter = "[primary transaksi] = " & Me.TxtMP_Primary.Text & " and [primary product] = " & rowDgvItemMP.Cells(Me.ColMP_PrimaryCodeNo.Index).Value
                        If Me.vDT_ItemFollowUp.DefaultView.Count = 0 Then
                            rowDgvItemMP.Cells(Me.ColMP_DetailFollowUp.Index).Value = "-"
                        ElseIf Me.vDT_ItemFollowUp.DefaultView.Count = 1 Then
                            Me.vRowItemFollowUp = Me.vDT_ItemFollowUp.DefaultView(0).Row
                            rowDgvItemMP.Cells(Me.ColMP_DetailFollowUp.Index).Value = vRowItemFollowUp.no_po_mo
                        ElseIf Me.vDT_ItemFollowUp.DefaultView.Count > 1 Then
                            rowDgvItemMP.Cells(Me.ColMP_DetailFollowUp.Index).Value = "Multiple"
                        End If
                    Next
                End If

            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub AddToPO()
        Dim i As New FrmTransaksiPO
        Dim rowDGV, rowDGVPO As DataGridViewRow
        Dim TBLKombinasi As New DseTransaksiMaterialPlanning2.kombinasiDataTable
        Dim rowKombinasi As DseTransaksiMaterialPlanning2.kombinasiRow
        Dim ij As New FrmCariTransaksiPO
        Dim IdentifierBaru As String = ""
        Dim DefPrefCurrency As Integer = 0
        Dim Dummy As String = ""
        Dim Index1 As Integer = -1
        Dim DTAPurchReq As New DseTransaksiPRTableAdapters.tbltrans_purchase_req_itemTableAdapter
        Dim TBLPurchReq As New DseTransaksiPR.tbltrans_purchase_req_itemDataTable
        Dim vRowItemDeliveryPO As DseTransaksiPO.tbltrans_po_itemdeliveryRow
        Dim JmlTambah As Decimal = 0
        Dim TBLItemDeliveryPR As New DseTransaksiPR.tbltrans_purchase_req_itemdeliveryDataTable
        Dim rowItemDeliveryPR As DseTransaksiPR.tbltrans_purchase_req_itemdeliveryRow
        Dim JmlDefaultLain, JmlMultiLain As Decimal
        Dim Index2 As Integer = 0
        Dim tblOldData As New DseTransaksiMaterialPlanning2.DataTable7DataTable
        Dim rowOldData As DseTransaksiMaterialPlanning2.DataTable7Row

        Try
            DTAPurchReq.Connection.ConnectionString = Core_Module.GetConstring
            If Me.TxtFilter_PrimaryPurchReq.Text <> "" And Me.TxtFilter_PrimaryPurchReq.Text <> "-1" Then
                DTAPurchReq.FillBy_Primary(TBLPurchReq, Me.TxtFilter_PrimaryPurchReq.Text)
            End If

            '4. kelompokkan berdasarkan primary supplier dan primary currency
            For Each rowDGV In Me.dgvMP.Rows
                If rowDGV.Cells(Me.ColMP_FollowUp.Index).Value = Boolean.TrueString And rowDGV.Cells(Me.ColMP_ProcurementType.Index).Value = "Buy" Then
                    TBLKombinasi.DefaultView.RowFilter = "[primary supplier] = " + rowDGV.Cells(Me.ColMP_PrimarySupplier.Index).Value.ToString + _
                    " AND [primary currency] = " + rowDGV.Cells(Me.ColMP_PrimaryCurrency.Index).Value.ToString

                    If TBLKombinasi.DefaultView.Count = 0 Then
                        rowKombinasi = TBLKombinasi.NewkombinasiRow
                        rowKombinasi.primary_currency = rowDGV.Cells(Me.ColMP_PrimaryCurrency.Index).Value
                        rowKombinasi.primary_supplier = rowDGV.Cells(Me.ColMP_PrimarySupplier.Index).Value
                        TBLKombinasi.Rows.Add(rowKombinasi)
                    End If
                End If
            Next

            '4.b untuk add to po, hanya boleh memasukkan data 1 kombinasi supplier dan currency
            If TBLKombinasi.Rows.Count > 1 Then
                MsgBox("To add items to a purchase order, all item must refer to 1 supplier and uses 1 type of currency.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            For Each rowKombinasi In TBLKombinasi.Rows
                ij.BackColorAll = Me.BackColorAll
                ij.ShowDialog()
                If Not ij.PrimaryPO = -1 Then
                    'cek apakah primary supplier dan primary currencynya sama
                    If Me.ObjPO.GetPrimaryCurrency(ij.PrimaryPO) <> rowKombinasi.primary_currency Then
                        MsgBox("Please choose purchase order using " + _
                        Me.ObjCurrency.GetCurrencyCode(rowKombinasi.primary_currency) + " currency", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                    If Me.ObjPO.GetPrimarySupplier(ij.PrimaryPO) <> rowKombinasi.primary_supplier Then
                        MsgBox("Please choose purchase order using supplier: " + _
                        Me.ObjSupplier.GetSupplierCode(rowKombinasi.primary_supplier) + " - " + Me.ObjSupplier.GetSupplierName(rowKombinasi.primary_supplier), _
                        MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                    'cek apakah sudah di otorisasi
                    If Me.ObjPO.GetAuthorizationLevel(ij.PrimaryPO) <> 0 Then
                        MsgBox("The purchase order you choose have been authorized, please cancel authorize first", _
                        MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                End If
            Next

            '5. untuk setiap kombinasi primary supplier dan primary currency, buat po baru
            If Not ij.PrimaryPO = -1 Then
                If MsgBox("Add to po?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                    For Each rowKombinasi In TBLKombinasi.Rows
                        If Not ij.PrimaryPO = -1 Then
                            i = New FrmTransaksiPO
                            i.Show()
                            i.BAckColorAll = Me.BackColorAll
                            i.UserCode = Me.UserCode
                            i.Location = New Point((i.Width * -1) - 1, (i.Height * -1) - 1)
                            i.CariPO_ForAdd(ij.PrimaryPO)
                            For Each rowDgvItemPO As DataGridViewRow In i.dgvItem.Rows
                                tblOldData.DefaultView.RowFilter = "[primary product] = " & rowDgvItemPO.Cells(i.ColPrimaryCodeNo.Index).Value
                                If tblOldData.DefaultView.Count > 0 Then
                                    rowOldData = tblOldData.DefaultView(0).Row
                                    rowOldData.old_default_qty += rowDgvItemPO.Cells(i.ColQty.Index).Value
                                    rowOldData.old_multi_qty += rowDgvItemPO.Cells(i.ColQtyMulti.Index).Value
                                Else
                                    rowOldData = tblOldData.NewDataTable7Row
                                    rowOldData.primary_product = rowDgvItemPO.Cells(i.ColPrimaryCodeNo.Index).Value
                                    rowOldData.old_default_qty = rowDgvItemPO.Cells(i.ColQty.Index).Value
                                    rowOldData.old_multi_qty = rowDgvItemPO.Cells(i.ColQtyMulti.Index).Value
                                    tblOldData.Rows.Add(rowOldData)
                                End If
                            Next

                            DefPrefCurrency = Me.ObjCurrency.GetDefaultPrecission(rowKombinasi.primary_currency)
                            If Me.TxtFilter_PrimaryPurchReq.Text <> "" And Me.TxtFilter_PrimaryPurchReq.Text <> "-1" Then
                                Dummy = Me.ObjPurcReq.GetRemarks(Me.TxtFilter_PrimaryPurchReq.Text)
                                If i.TxtRemark.Text.Contains(Dummy) = False Then
                                    If i.TxtRemark.Text.Trim = "" Then
                                        i.TxtRemark.Text = Dummy
                                    Else
                                        i.TxtRemark.Text = i.TxtRemark.Text + "; " + Dummy
                                    End If
                                    If i.TxtRemark.Text.Length > 150 Then
                                        i.TxtRemark.Text = i.TxtRemark.Text.Substring(0, 150)
                                    End If
                                End If
                            End If

                            'isikan data item 
                            For Each rowDGV In Me.dgvMP.Rows
                                If rowDGV.Cells(Me.ColMP_FollowUp.Index).Value = Boolean.TrueString Then
                                    If rowDGV.Cells(Me.ColMP_PrimaryCurrency.Index).Value = rowKombinasi.primary_currency And _
                                    rowDGV.Cells(Me.ColMP_PrimarySupplier.Index).Value = rowKombinasi.primary_supplier Then
                                        Index1 = i.GetIndex(rowDGV.Cells(Me.ColMP_PrimaryCodeNo.Index).Value)

                                        'cari data yang sama persis
                                        If Index1 <> -1 Then
                                            'tambahkan nilai qty
                                            Dim row As DataGridViewRow = i.dgvItem.Rows(Index1)
                                            row.Cells(i.ColQty.Index).Value = ObjCD.DigitGrouping(CDec(row.Cells(i.ColQty.Index).Value) + _
                                            CDec(rowDGV.Cells(Me.ColMP_QtyOrder.Index).Value), rowDGV.Cells(Me.ColMP_DefPrecOrderUnit.Index).Value)
                                            'hitung selisih, untuk item delivery
                                            JmlTambah = row.Cells(i.ColQtyMulti.Index).Value
                                            JmlTambah = CDec(row.Cells(i.ColQtyMulti.Index).Value) - JmlTambah
                                            row.Cells(i.ColQtyMulti.Index).Value = ObjCD.DigitGrouping((row.Cells(i.ColQty.Index).Value * _
                                            row.Cells(i.ColRatioReference.Index).Value) / row.Cells(i.ColRatioDefault.Index).Value, _
                                            rowDGV.Cells(Me.ColMP_DefPrecUnitMulti.Index).Value)

                                            i.Hitung_Amount(row)
                                            i.Hitung_Total()

                                            'kalo menambahkan ke dalam item yang sudah ada, maka update qty item delivery jika delivery date control
                                            i.vTBLItemDelivery.DefaultView.RowFilter = "identifier = '" + row.Cells(i.ColIdentifier.Index).Value + "'"
                                            i.vTBLItemDelivery.DefaultView.Sort = "[delivery date] DESC"

                                            If i.vTBLItemDelivery.DefaultView.Count > 0 Then
                                                'hitung jml multi dan jml default pada item delivery
                                                Index2 = 0
                                                JmlDefaultLain = 0
                                                JmlMultiLain = 0

                                                For Each dv As Data.DataRowView In i.vTBLItemDelivery.DefaultView
                                                    If Index2 <> 0 Then
                                                        vRowItemDeliveryPO = dv.Row
                                                        JmlMultiLain += vRowItemDeliveryPO.delivery_qty_multi
                                                        JmlDefaultLain += vRowItemDeliveryPO.delivery_qty
                                                    End If
                                                    Index2 += 1
                                                Next

                                                vRowItemDeliveryPO = i.vTBLItemDelivery.DefaultView(0).Row
                                                ''@rosadinnn
                                                ''@MUaT
                                                vRowItemDeliveryPO.delivery_qty_multi = CDec(row.Cells(i.ColQtyMulti.Index).Value) - JmlMultiLain
                                                vRowItemDeliveryPO.delivery_qty = CDec(row.Cells(i.ColQty.Index).Value) - JmlDefaultLain

                                                'vRowItemDeliveryPO.delivery_qty = Me.ObjCD.DigitGrouping(i.GetQtyDefault(vRowItemDeliveryPO.delivery_qty_multi, _
                                                'row.Cells(i.ColRatioDefault.Index).Value, row.Cells(i.ColRatioReference.Index).Value), _
                                                'Me.ObjProd.GetAngkaDecimal(row.Cells(i.ColPrimaryCodeNo.Index).Value))
                                            End If
                                        Else
                                            'tambahkan baris baru
                                            Dim ExpiredDateStr As String
                                            If i.dtpDeliveryDate.Checked = False Then
                                                ExpiredDateStr = "-"
                                            Else
                                                ExpiredDateStr = Format(#1/1/1950#, "dd MMM yyyy")
                                            End If
                                            'cegah jangan sampe primary IPO kosong sama sekali
                                            If i.txtPrimaryIPO.Text = "" Then
                                                i.txtPrimaryIPO.Text = "-1"
                                            End If
                                            IdentifierBaru = i.GetIdentifierBaru
                                            i.dgvItem.Rows.Add()
                                            rowDGVPO = i.dgvItem.Rows(i.dgvItem.Rows.Count - 1)
                                            With rowDGVPO
                                                For Each rowItem1 As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item1Row In Me.vDT.Rows
                                                    If rowItem1.primary_code_no = rowDGV.Cells(Me.ColMP_PrimaryCodeNo.Index).Value Then
                                                        .Cells(i.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(rowDGV.Cells(Me.ColMP_SubTotal.Index).Value, DefPrefCurrency)
                                                        .Cells(i.ColCodeNo.Index).Value = rowDGV.Cells(Me.ColMP_CodeNo.Index).Value
                                                        .Cells(i.ColDeliveryCont.Index).Value = Boolean.FalseString
                                                        .Cells(i.ColDeliveryDate.Index).Value = "-"
                                                        .Cells(i.ColDescription.Index).Value = Me.ObjProd.GetDescription(rowDGV.Cells(Me.ColMP_PrimaryCodeNo.Index).Value)
                                                        .Cells(i.ColDiscount.Index).Value = Me.ObjCD.DigitGrouping(rowDGV.Cells(Me.ColMP_Disc.Index).Value, 2)
                                                        .Cells(i.ColDiscAmount.Index).Value = Me.ObjCD.DigitGrouping(rowDGV.Cells(Me.ColMP_DiscAmount.Index).Value, DefPrefCurrency)
                                                        .Cells(i.ColDiscFromPerct.Index).Value = rowDGV.Cells(Me.ColMP_DiscFromPerct.Index).Value
                                                        .Cells(i.ColIPONo.Index).Value = ""
                                                        .Cells(i.ColNettPrice.Index).Value = Me.ObjCD.DigitGrouping(rowDGV.Cells(Me.ColMP_NettPrice.Index).Value, DefPrefCurrency)
                                                        .Cells(i.ColNo.Index).Value = i.dgvItem.Rows.Count
                                                        .Cells(i.ColPrimaryCodeNo.Index).Value = rowDGV.Cells(Me.ColMP_PrimaryCodeNo.Index).Value
                                                        .Cells(i.ColPrimaryIPO.Index).Value = -1
                                                        .Cells(i.ColPrimaryUnits.Index).Value = rowDGV.Cells(Me.ColMP_PrimaryOrderUnit.Index).Value
                                                        .Cells(i.ColQty.Index).Value = Me.ObjCD.DigitGrouping(rowDGV.Cells(Me.ColMP_QtyOrder.Index).Value, rowItem1.def_prec_order_unit)
                                                        If Me.TxtFilter_PrimaryPurchReq.Text <> "" And Me.TxtFilter_PrimaryPurchReq.Text <> "-1" Then
                                                            .Cells(i.ColRemark.Index).Value = Me.GetRemarks(TBLPurchReq, rowDGV.Cells(Me.ColMP_PrimaryCodeNo.Index).Value)
                                                        Else
                                                            .Cells(i.ColRemark.Index).Value = ""
                                                        End If
                                                        .Cells(i.ColUnit.Index).Value = rowDGV.Cells(Me.ColMP_OrderUnit.Index).Value
                                                        .Cells(i.ColUnitPrice.Index).Value = Me.ObjCD.DigitGrouping(rowDGV.Cells(Me.ColMP_UnitPrice.Index).Value, DefPrefCurrency)
                                                        .Cells(i.colUserSecondaryUnits.Index).Value = Boolean.FalseString
                                                        .Cells(i.ColIdentifier.Index).Value = IdentifierBaru

                                                        ''@rosadinnn
                                                        ''@MUaT
                                                        .Cells(i.ColPrimaryUnitMulti.Index).Value = rowDGV.Cells(Me.ColMP_PrimaryUnitMulti.Index).Value
                                                        .Cells(i.ColUnitCodeMulti.Index).Value = rowItem1.Unit_Code_Multi
                                                        '---------------------------------------------------------
                                                        .Cells(i.ColQtyMulti.Index).Value = Me.ObjCD.DigitGrouping(rowDGV.Cells(Me.ColMP_QtyOrderMulti.Index).Value, rowItem1.def_prec_unit_multi)
                                                        '---------------------------------------------------------
                                                        '---------------------------------------------------------
                                                        .Cells(i.ColRatioDefault.Index).Value = Me.ObjCD.DigitGrouping(rowDGV.Cells(Me.ColMP_RatioDefault.Index).Value, rowDGV.Cells(Me.ColMP_DefPrecUnit.Index).Value)
                                                        .Cells(i.ColRatioReference.Index).Value = Me.ObjCD.DigitGrouping(rowDGV.Cells(Me.ColMP_RatioMulti.Index).Value, rowDGV.Cells(Me.ColMP_DefPrecUnitMulti.Index).Value)
                                                        '---------------------------------------------------------
                                                        i.SetRatioAndQtyDefault(rowDGVPO, True)

                                                        .Cells(i.ColToleranceDown.Index).Value = Me.ObjCD.DigitGrouping(rowItem1.Tolerance_Up, rowDGV.Cells(Me.ColMP_DefPrecUnitMulti.Index).Value)
                                                        .Cells(i.ColToleranceUp.Index).Value = Me.ObjCD.DigitGrouping(rowItem1.Tolerance_Down, rowDGV.Cells(Me.ColMP_DefPrecUnitMulti.Index).Value)
                                                        If rowItem1.qty_order_multiunit = 0 Then
                                                            .Cells(i.ColToleranceDownPerc.Index).Value = Me.ObjCD.DigitGrouping(0, 2)
                                                            .Cells(i.ColToleranceUpPerc.Index).Value = Me.ObjCD.DigitGrouping(0, 2)
                                                        Else
                                                            .Cells(i.ColToleranceUpPerc.Index).Value = Me.ObjCD.DigitGrouping((rowItem1.Tolerance_Down * 100) / rowItem1.qty_order_multiunit, 2)
                                                            .Cells(i.ColToleranceDownPerc.Index).Value = Me.ObjCD.DigitGrouping((rowItem1.Tolerance_Up * 100) / rowItem1.qty_order_multiunit, 2)
                                                        End If

                                                        If Me.TxtFilter_PrimaryPurchReq.Text <> "" And Me.TxtFilter_PrimaryPurchReq.Text <> "-1" Then
                                                            'tambahkan item delivery po item, jika item pr ybs delivery date control
                                                            'untuk amannya, setiap item po meminta item delivery yang pertama saja, walaupun di pr-nya multiplen delivery date
                                                            If Me.ObjPurcReq.ApakahDeliveryDateControl(Me.TxtFilter_PrimaryPurchReq.Text, rowDGV.Cells(Me.ColMP_PrimaryCodeNo.Index).Value) = True Then
                                                                .Cells(i.ColDeliveryCont.Index).Value = Boolean.TrueString
                                                                Me.ObjPurcReq.IsikanDeliveryDateData(Me.TxtFilter_PrimaryPurchReq.Text, _
                                                                rowDGV.Cells(Me.ColMP_PrimaryCodeNo.Index).Value, TBLItemDeliveryPR)
                                                                If TBLItemDeliveryPR.Rows.Count >= 1 Then
                                                                    rowItemDeliveryPR = TBLItemDeliveryPR.Rows(0)
                                                                    .Cells(i.ColDeliveryDate.Index).Value = Format(rowItemDeliveryPR.delivery_date, "dd MMM yyyy")
                                                                    vRowItemDeliveryPO = i.vTBLItemDelivery.Newtbltrans_po_itemdeliveryRow

                                                                    With vRowItemDeliveryPO
                                                                        .delay = 0
                                                                        .delivery_date = rowItemDeliveryPR.delivery_date
                                                                        .delivery_qty = rowDGVPO.Cells(i.ColQty.Index).Value
                                                                        .delivery_qty_multi = rowDGVPO.Cells(i.ColQtyMulti.Index).Value
                                                                        .identifier = rowDGVPO.Cells(i.ColIdentifier.Index).Value
                                                                        .primary = i.vTBLItemDelivery.Rows.Count
                                                                        .primary_transaksi = -1
                                                                        .remarks = ""
                                                                    End With
                                                                    i.vTBLItemDelivery.Rows.Add(vRowItemDeliveryPO)
                                                                Else
                                                                    MsgBox("Logic error, please contact your vendor if you this message.", MsgBoxStyle.Information, "Information")
                                                                    Exit For
                                                                End If
                                                            End If
                                                        End If

                                                        Exit For
                                                    End If
                                                Next
                                            End With
                                        End If
                                    End If
                                End If
                            Next

                            'fokuskan ke baris terakhir
                            While i.dgvItem.Rows(i.dgvItem.Rows.Count - 1).Displayed = False _
                            Or i.dgvItem.DisplayedRowCount(True) <> i.dgvItem.DisplayedRowCount(False)
                                i.dgvItem.FirstDisplayedScrollingRowIndex = i.dgvItem.FirstDisplayedScrollingRowIndex + 1
                            End While

                            'hitung total
                            i.Hitung_Total()

                            'i.Location = i.BeginingLocation
                            i.Hide()
                            i.ShowDialog()
                            If i.DialogResult = Windows.Forms.DialogResult.OK Then
                                Dim oldDefaultQty As Decimal = 0
                                Dim oldMultiQty As Decimal = 0
                                For Each rowDgvItemPO As DataGridViewRow In i.dgvItem.Rows
                                    tblOldData.DefaultView.RowFilter = "[primary product] = " & rowDgvItemPO.Cells(i.ColPrimaryCodeNo.Index).Value
                                    If tblOldData.DefaultView.Count > 0 Then
                                        rowOldData = tblOldData.DefaultView(0).Row
                                        oldDefaultQty = rowOldData.old_default_qty
                                        oldMultiQty = rowOldData.old_multi_qty
                                    End If
                                    For Each rowDgvItemMP As DataGridViewRow In Me.dgvMP.Rows
                                        If rowDgvItemPO.Cells(i.ColPrimaryCodeNo.Index).Value = rowDgvItemMP.Cells(Me.ColMP_PrimaryCodeNo.Index).Value And rowDgvItemMP.Cells(Me.ColMP_FollowUp.Index).Value = True And rowDgvItemMP.Cells(Me.ColMP_ProcurementType.Index).Value = "Buy" Then
                                            Me.vRowItemFollowUp = Me.vDT_ItemFollowUp.Newtbltrans_material_planning_item_followupRow
                                            Me.vRowItemFollowUp.primary_transaksi = Me.TxtMP_Primary.Text
                                            Me.vRowItemFollowUp.primary_product = rowDgvItemPO.Cells(i.ColPrimaryCodeNo.Index).Value
                                            Me.vRowItemFollowUp.tipe = ClsTransaksiMaterialPlanning.EnumActionType.BUY_ADDTO_PO
                                            Me.vRowItemFollowUp.primary_po_mo = ObjPO.GetPrimary(i.TxtPONo.Text)
                                            Me.vRowItemFollowUp.no_po_mo = i.TxtPONo.Text
                                            Me.vRowItemFollowUp.tanggal_po_mo = i.dtpDate.Value
                                            Me.vRowItemFollowUp.identifier_jo = String.Empty
                                            Me.vRowItemFollowUp.no_jo = String.Empty
                                            If rowDgvItemPO.Cells(i.ColQty.Index).Value - oldDefaultQty < 0 Then
                                                Me.vRowItemFollowUp.qty_default = 0
                                            Else
                                                Me.vRowItemFollowUp.qty_default = rowDgvItemPO.Cells(i.ColQty.Index).Value - oldDefaultQty
                                            End If
                                            Me.vRowItemFollowUp.primary_unit_default = rowDgvItemPO.Cells(i.ColPrimaryUnits.Index).Value
                                            Me.vRowItemFollowUp.ratio_default = rowDgvItemPO.Cells(i.ColRatioDefault.Index).Value
                                            If rowDgvItemPO.Cells(i.ColQtyMulti.Index).Value - oldMultiQty < 0 Then
                                                Me.vRowItemFollowUp.qty_multi = 0
                                            Else
                                                Me.vRowItemFollowUp.qty_multi = rowDgvItemPO.Cells(i.ColQtyMulti.Index).Value - oldMultiQty
                                            End If
                                            Me.vRowItemFollowUp.primary_unit_multi = rowDgvItemPO.Cells(i.ColPrimaryUnitMulti.Index).Value
                                            Me.vRowItemFollowUp.ratio_multi = rowDgvItemPO.Cells(i.ColRatioReference.Index).Value
                                            Me.vRowItemFollowUp.user_code = Me.UserCode
                                            Me.vRowItemFollowUp.followup_time = Now
                                            Me.vDT_ItemFollowUp.Rows.Add(Me.vRowItemFollowUp)
                                            Me.SaveRequired = True
                                            Exit For
                                        End If
                                    Next
                                Next
                                For Each rowDgvItemMP As DataGridViewRow In Me.dgvMP.Rows
                                    Me.vDT_ItemFollowUp.DefaultView.RowFilter = "[primary transaksi] = " & Me.TxtMP_Primary.Text & " and [primary product] = " & rowDgvItemMP.Cells(Me.ColMP_PrimaryCodeNo.Index).Value
                                    If Me.vDT_ItemFollowUp.DefaultView.Count = 0 Then
                                        rowDgvItemMP.Cells(Me.ColMP_DetailFollowUp.Index).Value = "-"
                                    ElseIf Me.vDT_ItemFollowUp.DefaultView.Count = 1 Then
                                        Me.vRowItemFollowUp = Me.vDT_ItemFollowUp.DefaultView(0).Row
                                        rowDgvItemMP.Cells(Me.ColMP_DetailFollowUp.Index).Value = vRowItemFollowUp.no_po_mo
                                    ElseIf Me.vDT_ItemFollowUp.DefaultView.Count > 1 Then
                                        rowDgvItemMP.Cells(Me.ColMP_DetailFollowUp.Index).Value = "Multiple"
                                    End If
                                Next
                            End If

                            i.Dispose()
                        End If
                        ij.Dispose()
                    Next
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetRemarks(ByVal vTBLPurchReq As DseTransaksiPR.tbltrans_purchase_req_itemDataTable, ByVal vPrimaryCodeNo As Double) As String
        Dim RowPurchReq As DseTransaksiPR.tbltrans_purchase_req_itemRow

        Try
            For Each RowPurchReq In vTBLPurchReq.Rows
                If RowPurchReq.primary_code_no = vPrimaryCodeNo Then
                    Return RowPurchReq.remarks
                End If
            Next

            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function GetQty(ByVal vTBLPurchReq As DseTransaksiPR.tbltrans_purchase_req_itemDataTable, ByVal vPrimaryCodeNo As Double) As Decimal
        Dim RowPurchReq As DseTransaksiPR.tbltrans_purchase_req_itemRow

        Try
            For Each RowPurchReq In vTBLPurchReq.Rows
                If RowPurchReq.primary_code_no = vPrimaryCodeNo Then
                    Return RowPurchReq.qty
                End If
            Next

            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        End Try
    End Function

    Private Sub llbMP_CreatePO_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles llbMP_CreatePO.Click
        Dim rowDGV As DataGridViewRow
        Dim AdaModify As Boolean = False
        Dim Dummy As String = ""

        Try
            For Each rowDGV In Me.dgvMP.Rows
                If rowDGV.Cells(Me.ColMP_FollowUp.Index).Value = Boolean.TrueString Then
                    If CDec(rowDGV.Cells(Me.ColMP_QtyOrder.Index).Value) > 0 Then
                        AdaModify = True
                        Exit For
                    End If
                End If
            Next

            If AdaModify = True Then
                'cek product code apa aja dalam dgv yang di follow up tetapi datanya belum lengkap
                '1. cek supplier nya belum dipilih
                Dummy = ""
                For Each rowDGV In Me.dgvMP.Rows
                    If rowDGV.Cells(Me.ColMP_FollowUp.Index).Value = Boolean.TrueString And rowDGV.Cells(Me.ColMP_ProcurementType.Index).Value = "Buy" Then
                        If rowDGV.Cells(Me.ColMP_PrimarySupplier.Index).Value = -1 Then
                            If Dummy = "" Then
                                Dummy = rowDGV.Cells(Me.ColMP_CodeNo.Index).Value
                            Else
                                Dummy = Dummy + ", " + rowDGV.Cells(Me.ColMP_CodeNo.Index).Value
                            End If
                        End If
                    End If
                Next
                If Dummy <> "" Then
                    MsgBox("The following item have unspecified supplier: " + Dummy + ". Please choose supplier.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                '2. unit price = 0 
                Dummy = ""
                For Each rowDGV In Me.dgvMP.Rows
                    If rowDGV.Cells(Me.ColMP_FollowUp.Index).Value = Boolean.TrueString And rowDGV.Cells(Me.ColMP_ProcurementType.Index).Value = "Buy" Then
                        If rowDGV.Cells(Me.ColMP_UnitPrice.Index).Value = 0 Then
                            If Dummy = "" Then
                                Dummy = rowDGV.Cells(Me.ColMP_CodeNo.Index).Value
                            Else
                                Dummy = Dummy + ", " + rowDGV.Cells(Me.ColMP_CodeNo.Index).Value
                            End If
                        End If
                    End If
                Next
                If Dummy <> "" Then
                    MsgBox("The following item have unspecified unit price: " + Dummy + ". Please enter positive number for unit price data.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                '3. currency tidak ada
                Dummy = ""
                For Each rowDGV In Me.dgvMP.Rows
                    If rowDGV.Cells(Me.ColMP_FollowUp.Index).Value = Boolean.TrueString And rowDGV.Cells(Me.ColMP_ProcurementType.Index).Value = "Buy" Then
                        If rowDGV.Cells(Me.ColMP_PrimaryCurrency.Index).Value = -1 Then
                            If Dummy = "" Then
                                Dummy = rowDGV.Cells(Me.ColMP_CodeNo.Index).Value
                            Else
                                Dummy = Dummy + ", " + rowDGV.Cells(Me.ColMP_CodeNo.Index).Value
                            End If
                        End If
                    End If
                Next
                If Dummy <> "" Then
                    MsgBox("The following item have unspecified currency: " + Dummy + ". Please choose currency.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                '4. harus ada bom, untuk yang di produce
                Dummy = ""
                For Each rowDGV In Me.dgvMP.Rows
                    If rowDGV.Cells(Me.ColMP_FollowUp.Index).Value = Boolean.TrueString And rowDGV.Cells(Me.ColMP_ProcurementType.Index).Value = "Produce" Then
                        If ObjMO.GetJmlItemBOM(rowDGV.Cells(Me.ColMP_PrimaryCodeNo.Index).Value) <= 0 Then
                            If Dummy = "" Then
                                Dummy = rowDGV.Cells(Me.ColMP_CodeNo.Index).Value
                            Else
                                Dummy = Dummy + ", " + rowDGV.Cells(Me.ColMP_CodeNo.Index).Value
                            End If
                        End If
                    End If
                Next
                If Dummy <> "" Then
                    MsgBox("The following produce item have unspecified bill of material: " + Dummy + ". Please specify bill of material.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                '4.b. harus goods, untuk yang di produce
                Dummy = ""
                For Each rowDGV In Me.dgvMP.Rows
                    If rowDGV.Cells(Me.ColMP_FollowUp.Index).Value = Boolean.TrueString And rowDGV.Cells(Me.ColMP_ProcurementType.Index).Value = "Produce" Then
                        If ObjMO.GetTypeAccProduct(rowDGV.Cells(Me.ColMP_PrimaryCodeNo.Index).Value) <> 0 Then
                            If Dummy = "" Then
                                Dummy = rowDGV.Cells(Me.ColMP_CodeNo.Index).Value
                            Else
                                Dummy = Dummy + ", " + rowDGV.Cells(Me.ColMP_CodeNo.Index).Value
                            End If
                        End If
                    End If
                Next
                If Dummy <> "" Then
                    MsgBox("The following produce item is not goods: " + Dummy + ". We can only produce goods.", _
                    MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If

                '5. kelompokkan berdasarkan primary supplier dan primary currency
                If Me.ApakahAdaBuy = True Then
                    If Me.IsAllowedToCreateOrAddPO(Me.TxtMP_Primary.Text) = True Then
                        Me.CreateNewPO()
                    End If
                End If
                If Me.ApakahAdaProduce = True Then
                    If Me.IsAllowedToCreateOrAddMO(Me.TxtMP_Primary.Text) = True Then
                        Me.CreateNewMO()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CreateNewMO_ByCodeNo()
        Dim TBLItemProduct As New DseTrans_ManufacturingOrder.tblitem_productDataTable
        Dim rowItemProduct As DseTrans_ManufacturingOrder.tblitem_productRow
        Dim rowDGV As DataGridViewRow
        Dim IdxNow, TotalIdx As Integer

        Try
            IdxNow = 0
            TotalIdx = 0

            For Each rowDGV In Me.dgvMP.Rows
                If rowDGV.Cells(Me.ColMP_FollowUp.Index).Value = Boolean.TrueString Then
                    If rowDGV.Cells(Me.ColMP_ProcurementType.Index).Value = "Produce" And CDec(rowDGV.Cells(Me.ColMP_QtyOrder.Index).Value) > 0 Then
                        TotalIdx += 1
                    End If
                End If
            Next
            For Each rowDGV In Me.dgvMP.Rows
                If rowDGV.Cells(Me.ColMP_FollowUp.Index).Value = Boolean.TrueString Then
                    If rowDGV.Cells(Me.ColMP_ProcurementType.Index).Value = "Produce" And CDec(rowDGV.Cells(Me.ColMP_QtyOrder.Index).Value) > 0 Then
                        IdxNow += 1

                        Dim i As New FrmTrans_ManufacturingOrder
                        i.UserCode = Me.UserCode
                        i.BackColorAll = Me.BackColorAll
                        i.NewMOFromMP = True

                        TBLItemProduct.Rows.Clear()
                        rowItemProduct = TBLItemProduct.Newtblitem_productRow
                        rowItemProduct.Qty_Target = CDec(rowDGV.Cells(Me.ColMP_QtyOrder.Index).Value)
                        rowItemProduct.Primary_Product = rowDGV.Cells(Me.ColMP_PrimaryCodeNo.Index).Value
                        rowitemproduct.remarks =""
                        TBLItemProduct.Rows.Add(rowItemProduct)

                        i.FromMP = True
                        i.TBLItemProduct = TBLItemProduct
                        i.TotalIdx = TotalIdx
                        i.IdxNow = IdxNow

                        i.ShowDialog()

                        If i.StopProcess = True Then
                            Exit Sub
                        End If

                        If i.DialogResult = Windows.Forms.DialogResult.OK Then
                            For Each rowDgvItemMO As DataGridViewRow In i.DgvItem.Rows
                                If rowDgvItemMO.Cells(i.ColItem_PrimaryProduct.Index).Value <> -1 Then
                                    For Each rowDgvItemMP As DataGridViewRow In Me.dgvMP.Rows
                                        If rowDgvItemMO.Cells(i.ColItem_PrimaryProduct.Index).Value = rowDgvItemMP.Cells(Me.ColMP_PrimaryCodeNo.Index).Value And rowDgvItemMP.Cells(Me.ColMP_FollowUp.Index).Value = True And rowDgvItemMP.Cells(Me.ColMP_ProcurementType.Index).Value = "Produce" Then
                                            Me.vRowItemFollowUp = Me.vDT_ItemFollowUp.Newtbltrans_material_planning_item_followupRow
                                            Me.vRowItemFollowUp.primary_transaksi = Me.TxtMP_Primary.Text
                                            Me.vRowItemFollowUp.primary_product = rowDgvItemMO.Cells(i.ColItem_PrimaryProduct.Index).Value
                                            Me.vRowItemFollowUp.tipe = ClsTransaksiMaterialPlanning.EnumActionType.PRODUCE_CREATE_MO
                                            Me.vRowItemFollowUp.primary_po_mo = ObjMO.GetPrimary(i.TxtNoTransaksi.Text)
                                            Me.vRowItemFollowUp.no_po_mo = i.TxtNoTransaksi.Text
                                            Me.vRowItemFollowUp.tanggal_po_mo = i.DtpTglTransaksi.Value
                                            Me.vRowItemFollowUp.identifier_jo = rowDgvItemMO.Cells(i.ColItem_Identifier.Index).Value
                                            Me.vRowItemFollowUp.no_jo = rowDgvItemMO.Cells(i.ColItem_JONumber.Index).Value
                                            Me.vRowItemFollowUp.qty_default = rowDgvItemMO.Cells(i.ColItem_QtyTarget.Index).Value
                                            Me.vRowItemFollowUp.primary_unit_default = Me.objUnit.GetPrimary(rowDgvItemMO.Cells(i.ColItem_Unit.Index).Value)
                                            Me.vRowItemFollowUp.ratio_default = 1
                                            Me.vRowItemFollowUp.qty_multi = Me.vRowItemFollowUp.qty_default
                                            Me.vRowItemFollowUp.primary_unit_multi = Me.vRowItemFollowUp.primary_unit_default
                                            Me.vRowItemFollowUp.ratio_multi = 1
                                            Me.vRowItemFollowUp.user_code = Me.UserCode
                                            Me.vRowItemFollowUp.followup_time = Now
                                            Me.vRowItemFollowUp.remarks = i.TxtRemarks.Text
                                            Me.vDT_ItemFollowUp.Rows.Add(Me.vRowItemFollowUp)
                                            Me.SaveRequired = True
                                            Exit For
                                        End If
                                    Next
                                End If
                            Next
                            For Each rowDgvItemMP As DataGridViewRow In Me.dgvMP.Rows
                                Me.vDT_ItemFollowUp.DefaultView.RowFilter = "[primary transaksi] = " & Me.TxtMP_Primary.Text & " and [primary product] = " & rowDgvItemMP.Cells(Me.ColMP_PrimaryCodeNo.Index).Value
                                If Me.vDT_ItemFollowUp.DefaultView.Count = 0 Then
                                    rowDgvItemMP.Cells(Me.ColMP_DetailFollowUp.Index).Value = "-"
                                ElseIf Me.vDT_ItemFollowUp.DefaultView.Count = 1 Then
                                    Me.vRowItemFollowUp = Me.vDT_ItemFollowUp.DefaultView(0).Row
                                    rowDgvItemMP.Cells(Me.ColMP_DetailFollowUp.Index).Value = vRowItemFollowUp.no_po_mo
                                ElseIf Me.vDT_ItemFollowUp.DefaultView.Count > 1 Then
                                    rowDgvItemMP.Cells(Me.ColMP_DetailFollowUp.Index).Value = "Multiple"
                                End If
                            Next
                        End If

                        i.Dispose()
                    End If
                End If
            Next 
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CreateNewMO_ByRefDoc(ByVal FrmType As FrmTrans_MaterialPlanningDashboard_MP_AddMO)
        Dim rowDGV As DataGridViewRow
        Dim DV As Data.DataRowView
        Dim rowItemMO As DseTransaksiMaterialPlanning3.DataTable2Row
        Dim IdxNow, TotalIdx As Integer
        Dim TBLItemProduct As New DseTrans_ManufacturingOrder.tblitem_productDataTable
        Dim rowItemProduct As DseTrans_ManufacturingOrder.tblitem_productRow
        Dim Cust_PONo As String = ""

        Try
            IdxNow = 0
            FrmType.TblItemMO.DefaultView.RowFilter = "[Qty MO] > 0"
            TotalIdx = FrmType.TblItemMO.DefaultView.Count

            FrmType.TblItemMO.DefaultView.Sort = "[Primary Product], [Req Date], [Tipe Doc Int]"
            For Each rowDGV In FrmType.dgvCodeNo.Rows
                FrmType.TblItemMO.DefaultView.RowFilter = "[Primary Product] = " + rowDGV.Cells(FrmType.ColHeader_PrimaryCodeNo.Index).Value.ToString + " AND " + _
                "[Qty MO] > 0"

                For Each DV In FrmType.TblItemMO.DefaultView
                    rowItemMO = DV.Row

                    IdxNow += 1

                    Dim i As New FrmTrans_ManufacturingOrder
                    i.UserCode = Me.UserCode
                    i.BackColorAll = Me.BackColorAll
                    i.NewMOFromMP = True

                    TBLItemProduct.Rows.Clear()
                    rowItemProduct = TBLItemProduct.Newtblitem_productRow
                    rowItemProduct.Qty_Target = CDec(rowItemMO.Qty_MO)
                    rowItemProduct.Primary_Product = rowItemMO.Primary_Product

                    If rowItemMO.Tipe_Doc_Int = FrmTrans_MaterialPlanningDashboard_MP_AddMO.EnumPrioritasMO.SO Then
                        Cust_PONo = Me.ObjSO.GetCustPO(Me.ObjSO.GetPrimary(rowItemMO.Doc_No))
                        If Cust_PONo.Trim <> "" Then
                            rowItemProduct.remarks = rowItemMO.Tipe_Doc_Str + ":" + rowItemMO.Doc_No + " (Cust. PO:" + Cust_PONo + ")"
                        Else
                            rowItemProduct.remarks = rowItemMO.Tipe_Doc_Str + ":" + rowItemMO.Doc_No
                        End If
                    Else
                        rowItemProduct.remarks = rowItemMO.Tipe_Doc_Str + ":" + rowItemMO.Doc_No
                    End If
                    TBLItemProduct.Rows.Add(rowItemProduct)

                    i.FromMP = True
                    i.TBLItemProduct = TBLItemProduct
                    i.TotalIdx = TotalIdx
                    i.IdxNow = IdxNow

                    i.ShowDialog()

                    If i.StopProcess = True Then
                        Exit Sub
                    End If

                    If i.DialogResult = Windows.Forms.DialogResult.OK Then
                        For Each rowDgvItemMO As DataGridViewRow In i.DgvItem.Rows
                            If rowDgvItemMO.Cells(i.ColItem_PrimaryProduct.Index).Value <> -1 Then
                                For Each rowDgvItemMP As DataGridViewRow In Me.dgvMP.Rows
                                    If rowDgvItemMO.Cells(i.ColItem_PrimaryProduct.Index).Value = rowDgvItemMP.Cells(Me.ColMP_PrimaryCodeNo.Index).Value And rowDgvItemMP.Cells(Me.ColMP_FollowUp.Index).Value = True And rowDgvItemMP.Cells(Me.ColMP_ProcurementType.Index).Value = "Produce" Then
                                        Me.vRowItemFollowUp = Me.vDT_ItemFollowUp.Newtbltrans_material_planning_item_followupRow
                                        Me.vRowItemFollowUp.primary_transaksi = Me.TxtMP_Primary.Text
                                        Me.vRowItemFollowUp.primary_product = rowDgvItemMO.Cells(i.ColItem_PrimaryProduct.Index).Value
                                        Me.vRowItemFollowUp.tipe = ClsTransaksiMaterialPlanning.EnumActionType.PRODUCE_CREATE_MO
                                        Me.vRowItemFollowUp.primary_po_mo = ObjMO.GetPrimary(i.TxtNoTransaksi.Text)
                                        Me.vRowItemFollowUp.no_po_mo = i.TxtNoTransaksi.Text
                                        Me.vRowItemFollowUp.tanggal_po_mo = i.DtpTglTransaksi.Value
                                        Me.vRowItemFollowUp.identifier_jo = rowDgvItemMO.Cells(i.ColItem_Identifier.Index).Value
                                        Me.vRowItemFollowUp.no_jo = rowDgvItemMO.Cells(i.ColItem_JONumber.Index).Value
                                        Me.vRowItemFollowUp.qty_default = rowDgvItemMO.Cells(i.ColItem_QtyTarget.Index).Value
                                        Me.vRowItemFollowUp.primary_unit_default = Me.objUnit.GetPrimary(rowDgvItemMO.Cells(i.ColItem_Unit.Index).Value)
                                        Me.vRowItemFollowUp.ratio_default = 1
                                        Me.vRowItemFollowUp.qty_multi = Me.vRowItemFollowUp.qty_default
                                        Me.vRowItemFollowUp.primary_unit_multi = Me.vRowItemFollowUp.primary_unit_default
                                        Me.vRowItemFollowUp.ratio_multi = 1
                                        Me.vRowItemFollowUp.user_code = Me.UserCode
                                        Me.vRowItemFollowUp.followup_time = Now
                                        Me.vRowItemFollowUp.remarks = i.TxtRemarks.Text
                                        Me.vDT_ItemFollowUp.Rows.Add(Me.vRowItemFollowUp)
                                        Me.SaveRequired = True
                                        Exit For
                                    End If
                                Next
                            End If
                        Next
                        For Each rowDgvItemMP As DataGridViewRow In Me.dgvMP.Rows
                            Me.vDT_ItemFollowUp.DefaultView.RowFilter = "[primary transaksi] = " & Me.TxtMP_Primary.Text & " and [primary product] = " & rowDgvItemMP.Cells(Me.ColMP_PrimaryCodeNo.Index).Value
                            If Me.vDT_ItemFollowUp.DefaultView.Count = 0 Then
                                rowDgvItemMP.Cells(Me.ColMP_DetailFollowUp.Index).Value = "-"
                            ElseIf Me.vDT_ItemFollowUp.DefaultView.Count = 1 Then
                                Me.vRowItemFollowUp = Me.vDT_ItemFollowUp.DefaultView(0).Row
                                rowDgvItemMP.Cells(Me.ColMP_DetailFollowUp.Index).Value = vRowItemFollowUp.no_po_mo
                            ElseIf Me.vDT_ItemFollowUp.DefaultView.Count > 1 Then
                                rowDgvItemMP.Cells(Me.ColMP_DetailFollowUp.Index).Value = "Multiple"
                            End If
                        Next
                    End If

                    i.Dispose()
                Next
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CreateNewMO() 
        Dim rowDGV As DataGridViewRow
        Dim FrmType As New FrmTrans_MaterialPlanningDashboard_MP_AddMO
        Dim rowCodeNo As DseTransaksiMaterialPlanning3.DataTable1Row

        Try
            If MsgBox("Create new mo?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                FrmType.TblCodeNo.Rows.Clear()
                For Each rowDGV In Me.dgvMP.Rows
                    If rowDGV.Cells(Me.ColMP_FollowUp.Index).Value = Boolean.TrueString Then
                        If rowDGV.Cells(Me.ColMP_ProcurementType.Index).Value = "Produce" And CDec(rowDGV.Cells(Me.ColMP_QtyOrder.Index).Value) > 0 Then
                            rowCodeNo = FrmType.TblCodeNo.NewDataTable1Row

                            rowCodeNo.Code_No = rowDGV.Cells(Me.ColMP_CodeNo.Index).Value
                            rowCodeNo.Desc = rowDGV.Cells(Me.ColMP_Desc.Index).Value
                            rowCodeNo.Primary_Product = rowDGV.Cells(Me.ColMP_PrimaryCodeNo.Index).Value
                            rowCodeNo.Primary_Unit = rowDGV.Cells(Me.ColMP_PrimaryOrderUnit.Index).Value
                            rowCodeNo.Qty_MO = rowDGV.Cells(Me.ColMP_QtyOrder.Index).Value
                            rowCodeNo.Qty_Order = rowCodeNo.Qty_MO
                            rowCodeNo.Unit_Code = rowDGV.Cells(Me.ColMP_Unit.Index).Value
                            rowCodeNo.DefPrec = rowDGV.Cells(Me.ColMP_DefPrecUnit.Index).Value

                            FrmType.TblCodeNo.Rows.Add(rowCodeNo)
                        End If
                    End If
                Next

                FrmType.vDT_ItemDO = Me.vDT_ItemDO
                FrmType.vDT_ItemPrOB = Me.vDT_ItemPrOB
                FrmType.vDT_ItemRMReq = Me.vDT_ItemRMReq
                FrmType.vDT_ItemSOB = Me.vDT_ItemSOB
                FrmType.vDT_ItemSOBDelivery = Me.vDT_ItemSOBDelivery
                FrmType.vDT_ItemPrOBDelivery = Me.vDT_ItemPrOBDelivery

                If FrmType.ShowDialog = Windows.Forms.DialogResult.OK Then
                    If FrmType.rboNoDeliveryDate.Checked = True Then
                        CreateNewMO_ByCodeNo()
                    Else
                        Me.CreateNewMO_ByRefDoc(FrmType)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FrmType.Dispose()
        End Try
    End Sub

    Private Function ApakahAdaBuy() As Boolean
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvMP.Rows
                If rowDGV.Cells(Me.ColMP_FollowUp.Index).Value = Boolean.TrueString Then
                    If rowDGV.Cells(Me.ColMP_ProcurementType.Index).Value = "Buy" Then
                        Return True
                    End If
                End If
            Next
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function ApakahAdaProduce() As Boolean
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvMP.Rows
                If rowDGV.Cells(Me.ColMP_FollowUp.Index).Value = Boolean.TrueString Then
                    If rowDGV.Cells(Me.ColMP_ProcurementType.Index).Value = "Produce" Then
                        Return True
                    End If
                End If
            Next
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub AddDeliveryDate(ByVal vTBLItemDeliveryPO As DseTransaksiPO.tbltrans_po_itemdeliveryDataTable, ByVal i As FrmTransaksiPO)
        Dim rowDGV As DataGridViewRow
        Dim TBLDummy As New DseTransaksiMaterialPlanning2.dummy_itemdeliveryDataTable
        Dim rowDummy, rowDummy1 As DseTransaksiMaterialPlanning2.dummy_itemdeliveryRow
        Dim QtyOrder As Decimal
        Dim QtyOnHand, QtyPOB, QtyGIT, QtyMOB, QtyReducer, QtyRequired, QtySisaOrder, QtyDetailOrder As Decimal
        Dim AdaJanjiKirim As Boolean = False
        Dim rowItemDeliveryPO As DseTransaksiPO.tbltrans_po_itemdeliveryRow
        Dim DummyRemarks As String = ""

        Try
            '       -> bagaimana kalo ada sebagian ada delivery date dan sebagian tidak ada? -> set ke delivery yang paling kecil yang ada (dihandle pada fungsi: GetTBLDummy)
            '       -> bagaimana delivery date pending do? -> ambil yang paling awal, karena ini sudah pending do, artinya harus segera (dihandle pada fungsi: GetTBLDummy)
            '       -> bagaimana jika qty required < 0, tapi qty order ada? -> tidak set detail delivery
            '       -> bagaimana jika qty order > qty required? -> sisa qty-nya ditaro dipengiriman pertama
            vTBLItemDeliveryPO.Rows.Clear()

            '1. kumpulkan primary code yang diisikan ke dalam po
            For Each rowDGV In i.dgvItem.Rows
                For Each rowItem1 As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item1Row In Me.vDT.Rows
                    If rowDGV.Cells(i.ColPrimaryCodeNo.Index).Value = rowItem1.primary_code_no Then
                        TBLDummy.Rows.Clear()
                        QtyOrder = CDec(rowDGV.Cells(i.ColQty.Index).Value)
                        QtyOnHand = 0
                        QtyPOB = 0
                        QtyGIT = 0
                        QtyMOB = 0
                        QtyRequired = 0

                        'isikan qty on hand, qty pob, qty git, qty mob
                        Me.GetQtyDetail(QtyOnHand, QtyPOB, QtyGIT, QtyMOB, rowItem1)

                        '2. kumpulkan tanggal pengiriman untuk setiap code no tersebut, caranya:
                        '   qty free = qty on hand + po + git + mob - sob(+spare) - rm req(+spare) - prOB(+spare) - pending do
                        '   a. konversi tabel sob, rm req, pending do ke dalam tabel dummy
                        '   b. urutkan qty required berdasarkan tgl delivery, prioritas, tipe dokumen, dan primary dokumen 
                        Me.GetTBLDummy(rowItem1, TBLDummy)
                        '       -> bagaimana kalo tidak ada delivery date-nya -> tidak set detail delivery
                        For Each rowDummy In TBLDummy.Rows
                            If rowDummy.delivery_date <> #1/1/1950# Then
                                AdaJanjiKirim = True
                                Exit For
                            End If
                        Next
                        QtyReducer = QtyOnHand + QtyPOB + QtyGIT + QtyMOB
                        For Each rowDummy In TBLDummy.Rows
                            QtyRequired += rowDummy.delivery_qty
                        Next
                        If AdaJanjiKirim = True And QtyReducer < QtyRequired Then
                            'kurangkan qty required dengan penjumlahan qty on hand + qty pob + qty git + qty mob -> dari yg paling atas
                            While QtyReducer > 0 And TBLDummy.Rows.Count > 0
                                rowDummy = TBLDummy.Rows(0)
                                If QtyReducer >= rowDummy.delivery_qty Then
                                    QtyReducer -= rowDummy.delivery_qty
                                    TBLDummy.Rows.Remove(rowDummy)
                                Else
                                    rowDummy.delivery_qty -= QtyReducer
                                    QtyReducer = 0
                                End If
                            End While

                            'hanya lanjut klo tbldummy ada isinya
                            If TBLDummy.Rows.Count > 0 Then
                                'tambahkan qty delivery pertama dengan sisa order
                                QtySisaOrder = rowItem1.qty_order - rowItem1.qty_required
                                If QtySisaOrder > 0 Then
                                    rowDummy = TBLDummy.Newdummy_itemdeliveryRow
                                    rowDummy1 = TBLDummy.Rows(0)
                                    rowDummy.delivery_date = rowDummy1.delivery_date
                                    rowDummy.delivery_priority = rowDummy1.delivery_priority
                                    rowDummy.delivery_qty = QtySisaOrder
                                    rowDummy.order_qty = 0
                                    rowDummy.primary_delivery_priority = rowDummy1.primary_delivery_priority
                                    rowDummy.primary_transaksi = -1
                                    rowDummy.tipe_transaksi = CInt(EnumTipeDocReq.free)
                                    TBLDummy.Rows.InsertAt(rowDummy, 0)
                                End If

                                'menambahkan row qty required ke dalam tbl detail delivery
                                QtyDetailOrder = 0
                                For Each rowDummy In TBLDummy.Rows
                                    vTBLItemDeliveryPO.DefaultView.RowFilter = "identifier = '" + rowDGV.Cells(i.ColIdentifier.Index).Value.ToString + "' AND [delivery date] = #" + Format(rowDummy.delivery_date, "MM/dd/yyyy") + "#"
                                    If QtyDetailOrder < rowItem1.qty_order Then
                                        If vTBLItemDeliveryPO.DefaultView.Count > 0 Then
                                            'tambahkan delivery qty dan remarks saja
                                            rowItemDeliveryPO = vTBLItemDeliveryPO.DefaultView(0).Row
                                            If QtyDetailOrder + rowDummy.delivery_qty <= rowItem1.qty_order Then
                                                rowItemDeliveryPO.delivery_qty += rowDummy.delivery_qty
                                                QtyDetailOrder += rowDummy.delivery_qty
                                            Else
                                                rowItemDeliveryPO.delivery_qty += rowItem1.qty_order - QtyDetailOrder
                                                QtyDetailOrder = rowItem1.qty_order
                                            End If
                                            If rowDummy.tipe_transaksi = CInt(EnumTipeDocReq.free) Then
                                                DummyRemarks = "Free order"
                                            Else
                                                DummyRemarks = rowDummy.remarks
                                            End If
                                            If rowItemDeliveryPO.remarks.Contains(DummyRemarks) = False Then
                                                If rowItemDeliveryPO.remarks.Length + DummyRemarks.Length <= 250 - 2 Then
                                                    rowItemDeliveryPO.remarks = rowItemDeliveryPO.remarks + "; " + DummyRemarks
                                                End If
                                            End If
                                        Else
                                            'tambahkan item delivery baru
                                            rowItemDeliveryPO = vTBLItemDeliveryPO.Newtbltrans_po_itemdeliveryRow
                                            rowItemDeliveryPO.delay = 1
                                            rowItemDeliveryPO.delivery_date = rowDummy.delivery_date
                                            'rowItemDeliveryPO.delivery_qty = rowDummy.delivery_qty
                                            If QtyDetailOrder + rowDummy.delivery_qty <= rowItem1.qty_order Then
                                                rowItemDeliveryPO.delivery_qty = rowDummy.delivery_qty
                                                QtyDetailOrder += rowDummy.delivery_qty
                                            Else
                                                rowItemDeliveryPO.delivery_qty = rowItem1.qty_order - QtyDetailOrder
                                                QtyDetailOrder = rowItem1.qty_order
                                            End If
                                            rowItemDeliveryPO.delivery_qty_multi = 0
                                            rowItemDeliveryPO.identifier = rowDGV.Cells(i.ColIdentifier.Index).Value
                                            rowItemDeliveryPO.primary = GetPrimaryMax(vTBLItemDeliveryPO) + 1
                                            If rowItemDeliveryPO.primary = -1 Then
                                                rowItemDeliveryPO.primary = 1
                                            End If
                                            rowItemDeliveryPO.primary_transaksi = -1
                                            If rowDummy.tipe_transaksi = CInt(EnumTipeDocReq.free) Then
                                                DummyRemarks = "Free order"
                                            Else
                                                DummyRemarks = rowDummy.remarks
                                            End If
                                            rowItemDeliveryPO.remarks = DummyRemarks

                                            vTBLItemDeliveryPO.Rows.Add(rowItemDeliveryPO)
                                        End If
                                    Else
                                        Exit For
                                    End If
                                Next
                            End If
                        Else
                            'tidak ada 1 pun item required atas code no tsb yang ada janji kirimnya, maka tidak dimasukkan item detail delivery
                        End If


                        Exit For
                    End If
                Next
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetPrimaryMax(ByVal vTBLItemDeliveryPO As DseTransaksiPO.tbltrans_po_itemdeliveryDataTable) As Double
        Dim rowItemDeliveryPO As DseTransaksiPO.tbltrans_po_itemdeliveryRow
        Dim RetVal As Double = -1

        Try
            For Each rowItemDeliveryPO In vTBLItemDeliveryPO.Rows
                If rowItemDeliveryPO.primary > RetVal Then
                    RetVal = rowItemDeliveryPO.primary
                End If
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    Private Sub GetTBLDummy(ByVal rowItem1 As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item1Row, ByVal TBLDummy As DseTransaksiMaterialPlanning2.dummy_itemdeliveryDataTable)
        Dim DV, DV1 As Data.DataRowView
        Dim rowItemSOBDelivery As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsob_deliveryRow
        Dim rowItemPrOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_reqRow
        Dim rowItemPrOBDelivery As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempurch_req_deliveryRow
        Dim rowItemRMReq As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item_rm_reqRow
        Dim rowItemSOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemsobRow
        Dim rowItemDO As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemdoRow
        Dim TipeTransaksi As EnumTipeDocReq = EnumTipeDocReq.so_
        Dim CurDlvIdx As Integer = 0
        Dim TBLDummy1 As New DseTransaksiMaterialPlanning2.dummy_itemdeliveryDataTable
        Dim rowDummy, rowDummy1 As DseTransaksiMaterialPlanning2.dummy_itemdeliveryRow
        Dim DTADeliveryPriority As New DseTransaksiMaterialPlanning2TableAdapters.tblmaster_priority_levelTableAdapter
        Dim TBLDeliveryPriority As New DseTransaksiMaterialPlanning2.tblmaster_priority_levelDataTable
        Dim rowDeliveryPriority As DseTransaksiMaterialPlanning2.tblmaster_priority_levelRow
        Dim PrimaryDeliveryDefault As Double = -1
        Dim UrutanDeliveryDefault As Integer = -1
        Dim DeliveryDateMin As Date = Date.MaxValue

        Try
            DTADeliveryPriority.Connection.ConnectionString = Core_Module.GetConstring
            DTADeliveryPriority.Fill(TBLDeliveryPriority)

            For Each rowDeliveryPriority In TBLDeliveryPriority.Rows
                If rowDeliveryPriority._default = 1 Then
                    PrimaryDeliveryDefault = rowDeliveryPriority.primary
                    UrutanDeliveryDefault = rowDeliveryPriority.urutan
                    Exit For
                End If
            Next
            If PrimaryDeliveryDefault = -1 Then
                If TBLDeliveryPriority.Rows.Count = 0 Then
                    MsgBox("Logic error: no delivery priority specified. Please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                Else
                    rowDeliveryPriority = TBLDeliveryPriority.Rows(0)
                    PrimaryDeliveryDefault = rowDeliveryPriority.primary
                    UrutanDeliveryDefault = rowDeliveryPriority.urutan
                End If
            End If

            Me.vDT_ItemSOB.DefaultView.RowFilter = "[primary product] = " + rowItem1.primary_code_no.ToString
            TipeTransaksi = EnumTipeDocReq.so_
            For Each DV In Me.vDT_ItemSOB.DefaultView
                rowItemSOB = DV.Row
                CurDlvIdx = 0
                If rowItemSOB.delivery_date_cont = 1 Then
                    Me.vDT_ItemSOBDelivery.DefaultView.RowFilter = "identifier = '" + rowItemSOB.identifier_so + "'"
                    Me.vDT_ItemSOBDelivery.DefaultView.Sort = "[delivery date]"

                    For Each DV1 In Me.vDT_ItemSOBDelivery.DefaultView
                        rowItemSOBDelivery = DV1.Row
                        rowDummy = TBLDummy.Newdummy_itemdeliveryRow
                        With rowDummy
                            .delivery_date = rowItemSOBDelivery.delivery_date.Date
                            If DeliveryDateMin.Date > .delivery_date.Date Then
                                DeliveryDateMin = .delivery_date.Date
                            End If
                            For Each rowDeliveryPriority In TBLDeliveryPriority.Rows
                                If rowDeliveryPriority.primary = rowItemSOBDelivery.delivery_priority Then
                                    .delivery_priority = rowDeliveryPriority.urutan
                                    Exit For
                                End If
                            Next
                            If CurDlvIdx = 0 Then
                                'item pengiriman pertama ditambahkan spare qty
                                .delivery_qty = rowItemSOBDelivery.delivery_qty + rowItemSOB.MP_Spare_Qty
                            Else
                                .delivery_qty = rowItemSOBDelivery.delivery_qty
                            End If
                            .order_qty = 0
                            .primary_delivery_priority = rowItemSOBDelivery.delivery_priority
                            .primary_transaksi = rowItemSOB.primary_so_rec
                            .tipe_transaksi = CInt(TipeTransaksi)
                            .remarks = rowItemSOB.No_Transaksi
                        End With
                        CurDlvIdx += 1
                        TBLDummy.Rows.Add(rowDummy)
                    Next
                Else
                    rowDummy = TBLDummy.Newdummy_itemdeliveryRow
                    With rowDummy
                        .delivery_date = #1/1/1950#
                        .delivery_priority = UrutanDeliveryDefault
                        .delivery_qty = rowItemSOB.qty + rowItemSOB.MP_Spare_Qty
                        .order_qty = 0
                        .primary_delivery_priority = PrimaryDeliveryDefault
                        .primary_transaksi = rowItemSOB.primary_so_rec
                        .tipe_transaksi = CInt(TipeTransaksi)
                        .remarks = rowItemSOB.No_Transaksi
                    End With
                    TBLDummy.Rows.Add(rowDummy)
                End If
            Next

            Me.vDT_ItemRMReq.DefaultView.RowFilter = "[primary product] = " + rowItem1.primary_code_no.ToString
            Me.vDT_ItemRMReq.DefaultView.Sort = "[required date]"
            TipeTransaksi = EnumTipeDocReq.rm_req
            For Each DV In Me.vDT_ItemRMReq.DefaultView
                rowItemRMReq = DV.Row
                TBLDummy.DefaultView.RowFilter = "[tipe transaksi] = " + CInt(TipeTransaksi).ToString + " AND [identifier jo] = '" + rowItemRMReq.identifier_jo_rec + "'"
                If TBLDummy.DefaultView.Count > 0 Then
                    rowDummy = TBLDummy.DefaultView(0).Row
                    With rowDummy
                        .delivery_qty += (rowItemRMReq.qty + rowItemRMReq.MP_Spare_Qty)
                        .identifier_jo = rowItemRMReq.identifier_jo_rec
                    End With
                Else
                    rowDummy = TBLDummy.Newdummy_itemdeliveryRow
                    With rowDummy
                        .delivery_date = rowItemRMReq.required_date.Date
                        If DeliveryDateMin.Date > .delivery_date.Date Then
                            DeliveryDateMin = .delivery_date.Date
                        End If
                        .delivery_priority = UrutanDeliveryDefault
                        .delivery_qty = rowItemRMReq.qty + rowItemRMReq.MP_Spare_Qty
                        .order_qty = 0
                        .primary_delivery_priority = PrimaryDeliveryDefault
                        .primary_transaksi = rowItemRMReq.primary_mo_rec
                        .tipe_transaksi = CInt(TipeTransaksi)
                        .remarks = rowItemRMReq.jo_no_rec
                        .identifier_jo = rowItemRMReq.identifier_jo_rec
                    End With
                    TBLDummy.Rows.Add(rowDummy)
                End If
            Next

            Me.vDT_ItemPrOB.DefaultView.RowFilter = "[primary product] = " + rowItem1.primary_code_no.ToString
            TipeTransaksi = EnumTipeDocReq.pr_ob
            For Each DV In Me.vDT_ItemPrOB.DefaultView
                rowItemPrOB = DV.Row
                CurDlvIdx = 0

                If rowItemPrOB.delivery_date_cont = 1 Then
                    Me.vDT_ItemPrOBDelivery.DefaultView.RowFilter = "identifier = '" + rowItemPrOB.identifier_pr_item + "'"
                    Me.vDT_ItemPrOBDelivery.DefaultView.Sort = "[delivery date]"
                    For Each DV1 In Me.vDT_ItemPrOBDelivery.DefaultView
                        rowItemPrOBDelivery = DV1.Row
                        rowDummy = TBLDummy.Newdummy_itemdeliveryRow
                        With rowDummy
                            .delivery_date = rowItemPrOBDelivery.delivery_date.Date
                            If DeliveryDateMin.Date > .delivery_date.Date Then
                                DeliveryDateMin = .delivery_date.Date
                            End If
                            For Each rowDeliveryPriority In TBLDeliveryPriority.Rows
                                If rowDeliveryPriority.primary = rowItemPrOBDelivery.delivery_priority Then
                                    .delivery_priority = rowDeliveryPriority.urutan
                                    Exit For
                                End If
                            Next
                            If CurDlvIdx = 0 Then
                                'item pengiriman pertama ditambahkan spare qty
                                .delivery_qty = rowItemPrOBDelivery.delivery_qty + rowItemPrOB.MP_Spare_Qty
                            Else
                                .delivery_qty = rowItemPrOBDelivery.delivery_qty
                            End If
                            .order_qty = 0
                            .primary_delivery_priority = rowItemPrOBDelivery.delivery_priority
                            .primary_transaksi = rowItemPrOB.primary_pr_rec
                            .tipe_transaksi = CInt(TipeTransaksi)
                            .remarks = rowItemPrOB.No_Transaksi
                        End With
                        TBLDummy.Rows.Add(rowDummy)
                        CurDlvIdx += 1
                    Next
                Else
                    rowDummy = TBLDummy.Newdummy_itemdeliveryRow
                    With rowDummy
                        .delivery_date = #1/1/1950#
                        .delivery_priority = UrutanDeliveryDefault
                        .delivery_qty = rowItemPrOB.qty + rowItemPrOB.MP_Spare_Qty
                        .order_qty = 0
                        .primary_delivery_priority = PrimaryDeliveryDefault
                        .primary_transaksi = rowItemPrOB.primary_pr_rec
                        .tipe_transaksi = CInt(TipeTransaksi)
                        .remarks = rowItemPrOB.No_Transaksi
                    End With
                    TBLDummy.Rows.Add(rowDummy)
                End If
            Next

            Me.vDT_ItemDO.DefaultView.RowFilter = "[primary product] = " + rowItem1.primary_code_no.ToString
            TipeTransaksi = EnumTipeDocReq.do_
            For Each DV In Me.vDT_ItemDO.DefaultView
                rowItemDO = DV.Row
                rowDummy = TBLDummy.Newdummy_itemdeliveryRow
                With rowDummy
                    .delivery_date = #1/1/1950#
                    .delivery_priority = UrutanDeliveryDefault
                    .delivery_qty = rowItemDO.qty
                    .order_qty = 0
                    .primary_delivery_priority = PrimaryDeliveryDefault
                    .primary_transaksi = rowItemDO.primary_do_rec
                    .tipe_transaksi = CInt(TipeTransaksi)
                    .remarks = rowItemDO.DO_Number
                End With
                TBLDummy.Rows.Add(rowDummy)
            Next

            'ganti item2 yang delivery datanya = 1 jan 1950 ke tanggal delivery date minimum
            If DeliveryDateMin <> Date.MaxValue Then
                For Each rowDummy In TBLDummy.Rows
                    If rowDummy.delivery_date = #1/1/1950# Then
                        rowDummy.delivery_date = DeliveryDateMin
                    End If
                Next
            End If

            TBLDummy.DefaultView.RowFilter = ""
            TBLDummy.DefaultView.Sort = "[delivery date], [delivery priority], [tipe transaksi], [primary transaksi]"
            TBLDummy1.Rows.Clear()
            For Each DV In TBLDummy.DefaultView
                rowDummy = DV.Row
                rowDummy1 = TBLDummy1.Newdummy_itemdeliveryRow
                For ij As Integer = 0 To TBLDummy1.Columns.Count - 1
                    rowDummy1.Item(ij) = rowDummy.Item(ij)
                Next
                TBLDummy1.Rows.Add(rowDummy1)
            Next

            TBLDummy.Rows.Clear()
            For mn As Integer = 0 To TBLDummy1.Rows.Count - 1
                rowDummy1 = TBLDummy1.Rows(mn)
                rowDummy = TBLDummy.Newdummy_itemdeliveryRow
                For jk As Integer = 0 To TBLDummy1.Columns.Count - 1
                    rowDummy.Item(jk) = rowDummy1.Item(jk)
                Next
                TBLDummy.Rows.Add(rowDummy)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub GetQtyDetail(ByRef QtyOnHand As Decimal, ByRef QtyPOB As Decimal, ByRef QtyGIT As Decimal, ByRef QtyMOB As Decimal, _
    ByVal rowItem1 As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item1Row)
        Dim DV As Data.DataRowView
        Dim rowItemPOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itempobRow
        Dim rowItemQty As DseTransaksiMaterialPlanning_WarehouseExc.tbltrans_material_planning_item_qtyRow
        Dim rowItemGIT As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemgitRow
        Dim rowItemMOB As DseTransaksiMaterialPlanning2.tbltrans_material_planning_itemmobRow

        Try
            Me.vDT_ItemPOB.DefaultView.RowFilter = "[primary product] = " + rowItem1.primary_code_no.ToString
            For Each DV In Me.vDT_ItemPOB.DefaultView
                rowItemPOB = DV.Row
                QtyPOB += rowItemPOB.qty
            Next
            Me.vDT_ItemQty.DefaultView.RowFilter = "[Primary Product] = " + rowItem1.primary_code_no.ToString
            For Each DV In Me.vDT_ItemQty.DefaultView
                rowItemQty = DV.Row
                QtyOnHand += rowItemQty.Qty
            Next
            Me.vDT_ItemGIT.DefaultView.RowFilter = "[Primary Product] = " + rowItem1.primary_code_no.ToString
            For Each DV In Me.vDT_ItemGIT.DefaultView
                rowItemGIT = DV.Row
                QtyGIT += rowItemGIT.qty
            Next
            Me.vDT_ItemMOB.DefaultView.RowFilter = "[Primary Product] = " + rowItem1.primary_code_no.ToString
            For Each DV In Me.vDT_ItemMOB.DefaultView
                rowItemMOB = DV.Row
                QtyMOB += rowItemMOB.qty
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CreateNewPO()
        Dim i As New FrmTransaksiPO
        Dim rowDGV, rowDGVPO As DataGridViewRow
        Dim TBLKombinasi As New DseTransaksiMaterialPlanning2.kombinasiDataTable
        Dim rowKombinasi As DseTransaksiMaterialPlanning2.kombinasiRow
        Dim ij As Integer = 0
        Dim DefPrefCurrency As Integer = 0
        Dim Index1 As Integer = -1
        Dim IdentifierBaru As String = ""

        Dim DTAPurchReq As New DseTransaksiPRTableAdapters.tbltrans_purchase_req_itemTableAdapter
        Dim TBLPurchReq As New DseTransaksiPR.tbltrans_purchase_req_itemDataTable
        Dim TBLItemDeliveryPR As New DseTransaksiPR.tbltrans_purchase_req_itemdeliveryDataTable
        Dim rowItemDeliveryPR As DseTransaksiPR.tbltrans_purchase_req_itemdeliveryRow
        Dim vrowItemDeliveryPO As DseTransaksiPO.tbltrans_po_itemdeliveryRow
        Dim FrmPOStyle As New FrmTrans_MaterialPlanningDashboard_MP_AddPO
        Dim TBLItemDeliveryDummy As New DseTransaksiPO.tbltrans_po_itemdeliveryDataTable
        Dim AdaBedaDeliveryDate As Boolean = False
        Dim DeliveryDateDummy As Date = Date.MinValue
        Dim rowItemDelivery As DseTransaksiPO.tbltrans_po_itemdeliveryRow
        Dim ListDeliveryDate As New List(Of Date)
        Dim CreateIdx As Integer = 0
        Dim RetVal As Windows.Forms.DialogResult = Windows.Forms.DialogResult.OK

        Try
            If FrmPOStyle.ShowDialog = Windows.Forms.DialogResult.OK Then
                DTAPurchReq.Connection.ConnectionString = Core_Module.GetConstring
                If Me.TxtFilter_PrimaryPurchReq.Text <> "" And Me.TxtFilter_PrimaryPurchReq.Text <> "-1" Then
                    DTAPurchReq.FillBy_Primary(TBLPurchReq, Me.TxtFilter_PrimaryPurchReq.Text)
                End If

                For Each rowDGV In Me.dgvMP.Rows
                    If rowDGV.Cells(Me.ColMP_FollowUp.Index).Value = Boolean.TrueString And rowDGV.Cells(Me.ColMP_ProcurementType.Index).Value = "Buy" Then
                        TBLKombinasi.DefaultView.RowFilter = "[primary supplier] = " + rowDGV.Cells(Me.ColMP_PrimarySupplier.Index).Value.ToString + _
                        " AND [primary currency] = " + rowDGV.Cells(Me.ColMP_PrimaryCurrency.Index).Value.ToString

                        If TBLKombinasi.DefaultView.Count = 0 Then
                            rowKombinasi = TBLKombinasi.NewkombinasiRow
                            rowKombinasi.primary_currency = rowDGV.Cells(Me.ColMP_PrimaryCurrency.Index).Value
                            rowKombinasi.primary_supplier = rowDGV.Cells(Me.ColMP_PrimarySupplier.Index).Value
                            TBLKombinasi.Rows.Add(rowKombinasi)
                        End If
                    End If
                Next

                '5. untuk setiap kombinasi primary supplier dan primary currency, buat po baru
                For Each rowKombinasi In TBLKombinasi.Rows
                    ij += 1
                    i = New FrmTransaksiPO
                    i.BAckColorAll = Me.BackColorAll
                    i.UserCode = Me.UserCode
                    i.Show()
                    i.Location = New Point((i.Width * -1) - 1, (i.Height * -1) - 1)
                    i.NewPO()
                    i.SetSupplierData(rowKombinasi.primary_supplier)
                    i.SetCurrency(rowKombinasi.primary_currency)
                    DefPrefCurrency = Me.ObjCurrency.GetDefaultPrecission(rowKombinasi.primary_currency)
                    If Me.TxtFilter_PrimaryPurchReq.Text <> "" And Me.TxtFilter_PrimaryPurchReq.Text <> "-1" Then
                        i.TxtRemark.Text = Me.ObjPurcReq.GetRemarks(Me.TxtFilter_PrimaryPurchReq.Text)
                        If i.TxtRemark.Text.Length > 150 Then
                            i.TxtRemark.Text = i.TxtRemark.Text.Substring(0, 150)
                        End If
                    End If

                    i.lblProgress.Text = "Data " + ij.ToString + " of " + TBLKombinasi.Rows.Count.ToString
                    i.prgAll.Maximum = TBLKombinasi.Rows.Count
                    i.prgAll.Value = ij
                    i.lblProgress.Visible = True
                    i.btnStopProcess.Visible = True
                    i.prgAll.Visible = True

                    'isikan data item
                    For Each rowDGV In Me.dgvMP.Rows
                        If rowDGV.Cells(Me.ColMP_FollowUp.Index).Value = Boolean.TrueString Then
                            If rowDGV.Cells(Me.ColMP_PrimaryCurrency.Index).Value = rowKombinasi.primary_currency And _
                            rowDGV.Cells(Me.ColMP_PrimarySupplier.Index).Value = rowKombinasi.primary_supplier Then
                                Index1 = i.GetIndex(rowDGV.Cells(Me.ColMP_PrimaryCodeNo.Index).Value)

                                'cari data yang sama persis
                                If Index1 <> -1 Then
                                    'tambahkan nilai qty
                                    Dim row As DataGridViewRow = i.dgvItem.Rows(Index1)
                                    row.Cells(i.ColQty.Index).Value = ObjCD.DigitGrouping(CDec(row.Cells(i.ColQty.Index).Value) + _
                                    CDec(rowDGV.Cells(Me.ColMP_QtyOrder.Index).Value), rowDGV.Cells(Me.ColMP_DefPrecOrderUnit.Index).Value)

                                    row.Cells(i.ColQtyMulti.Index).Value = ObjCD.DigitGrouping((row.Cells(i.ColQty.Index).Value * _
                                    row.Cells(i.ColRatioReference.Index).Value) / row.Cells(i.ColRatioDefault.Index).Value, _
                                    rowDGV.Cells(Me.ColMP_DefPrecUnitMulti.Index).Value)

                                    i.Hitung_Amount(row)
                                    i.Hitung_Total()

                                    'daerah sini tidak ditambahkan delivery data control, karena kalo baru dan 
                                    'dirujuk dari purchase request, pasti tidak akan ke sini
                                    'karena ada batasan dalam 1 pr, hanya 1x code no disebutkan
                                Else
                                    'tambahkan baris baru
                                    Dim ExpiredDateStr As String
                                    If i.dtpDeliveryDate.Checked = False Then
                                        ExpiredDateStr = "-"
                                    Else
                                        ExpiredDateStr = Format(#1/1/1950#, "dd MMM yyyy")
                                    End If
                                    'cegah jangan sampe primary IPO kosong sama sekali
                                    If i.txtPrimaryIPO.Text = "" Then
                                        i.txtPrimaryIPO.Text = "-1"
                                    End If
                                    IdentifierBaru = i.GetIdentifierBaru
                                    i.dgvItem.Rows.Add()
                                    rowDGVPO = i.dgvItem.Rows(i.dgvItem.Rows.Count - 1)

                                    With rowDGVPO
                                        For Each rowItem1 As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item1Row In Me.vDT.Rows
                                            If rowItem1.primary_code_no = rowDGV.Cells(Me.ColMP_PrimaryCodeNo.Index).Value Then
                                                .Cells(i.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(rowDGV.Cells(Me.ColMP_SubTotal.Index).Value, DefPrefCurrency)
                                                .Cells(i.ColCodeNo.Index).Value = rowDGV.Cells(Me.ColMP_CodeNo.Index).Value
                                                .Cells(i.ColDeliveryCont.Index).Value = Boolean.FalseString
                                                .Cells(i.ColDeliveryDate.Index).Value = "-"
                                                .Cells(i.ColDescription.Index).Value = Me.ObjProd.GetDescription(rowDGV.Cells(Me.ColMP_PrimaryCodeNo.Index).Value)
                                                .Cells(i.ColDiscount.Index).Value = Me.ObjCD.DigitGrouping(rowDGV.Cells(Me.ColMP_Disc.Index).Value, 2)
                                                .Cells(i.ColDiscAmount.Index).Value = Me.ObjCD.DigitGrouping(rowDGV.Cells(Me.ColMP_DiscAmount.Index).Value, DefPrefCurrency)
                                                .Cells(i.ColDiscFromPerct.Index).Value = rowDGV.Cells(Me.ColMP_DiscFromPerct.Index).Value
                                                .Cells(i.ColIPONo.Index).Value = ""
                                                .Cells(i.ColNettPrice.Index).Value = Me.ObjCD.DigitGrouping(rowDGV.Cells(Me.ColMP_NettPrice.Index).Value, DefPrefCurrency)
                                                .Cells(i.ColNo.Index).Value = i.dgvItem.Rows.Count
                                                .Cells(i.ColPrimaryCodeNo.Index).Value = rowDGV.Cells(Me.ColMP_PrimaryCodeNo.Index).Value
                                                .Cells(i.ColPrimaryIPO.Index).Value = -1
                                                .Cells(i.ColPrimaryUnits.Index).Value = rowDGV.Cells(Me.ColMP_PrimaryOrderUnit.Index).Value
                                                .Cells(i.ColQty.Index).Value = Me.ObjCD.DigitGrouping(rowDGV.Cells(Me.ColMP_QtyOrder.Index).Value, rowItem1.def_prec_order_unit)
                                                If Me.TxtFilter_PrimaryPurchReq.Text <> "" And Me.TxtFilter_PrimaryPurchReq.Text <> "-1" Then
                                                    .Cells(i.ColRemark.Index).Value = Me.GetRemarks(TBLPurchReq, rowDGV.Cells(Me.ColMP_PrimaryCodeNo.Index).Value)
                                                Else
                                                    .Cells(i.ColRemark.Index).Value = ""
                                                End If
                                                .Cells(i.ColUnit.Index).Value = rowDGV.Cells(Me.ColMP_OrderUnit.Index).Value
                                                .Cells(i.ColUnitPrice.Index).Value = Me.ObjCD.DigitGrouping(rowDGV.Cells(Me.ColMP_UnitPrice.Index).Value, DefPrefCurrency)
                                                .Cells(i.colUserSecondaryUnits.Index).Value = Boolean.FalseString
                                                .Cells(i.ColIdentifier.Index).Value = IdentifierBaru

                                                ''@rosadinnn
                                                ''@MUaT
                                                .Cells(i.ColPrimaryUnitMulti.Index).Value = rowDGV.Cells(Me.ColMP_PrimaryUnitMulti.Index).Value
                                                .Cells(i.ColUnitCodeMulti.Index).Value = rowItem1.Unit_Code_Multi
                                                '---------------------------------------------------------
                                                .Cells(i.ColQtyMulti.Index).Value = Me.ObjCD.DigitGrouping(rowDGV.Cells(Me.ColMP_QtyOrderMulti.Index).Value, rowItem1.def_prec_unit_multi)
                                                '---------------------------------------------------------
                                                '---------------------------------------------------------
                                                .Cells(i.ColRatioDefault.Index).Value = Me.ObjCD.DigitGrouping(rowDGV.Cells(Me.ColMP_RatioDefault.Index).Value, rowDGV.Cells(Me.ColMP_DefPrecUnit.Index).Value)
                                                .Cells(i.ColRatioReference.Index).Value = Me.ObjCD.DigitGrouping(rowDGV.Cells(Me.ColMP_RatioMulti.Index).Value, rowDGV.Cells(Me.ColMP_DefPrecUnitMulti.Index).Value)
                                                '---------------------------------------------------------
                                                i.SetRatioAndQtyDefault(rowDGVPO, True)

                                                .Cells(i.ColToleranceDown.Index).Value = Me.ObjCD.DigitGrouping(rowItem1.Tolerance_Down, rowDGV.Cells(Me.ColMP_DefPrecUnitMulti.Index).Value)
                                                .Cells(i.ColToleranceUp.Index).Value = Me.ObjCD.DigitGrouping(rowItem1.Tolerance_Up, rowDGV.Cells(Me.ColMP_DefPrecUnitMulti.Index).Value)
                                                If rowItem1.qty_order_multiunit = 0 Then
                                                    .Cells(i.ColToleranceDownPerc.Index).Value = Me.ObjCD.DigitGrouping(0, 2)
                                                    .Cells(i.ColToleranceUpPerc.Index).Value = Me.ObjCD.DigitGrouping(0, 2)
                                                Else
                                                    .Cells(i.ColToleranceUpPerc.Index).Value = Me.ObjCD.DigitGrouping((rowItem1.Tolerance_Up * 100) / rowItem1.qty_order_multiunit, 2)
                                                    .Cells(i.ColToleranceDownPerc.Index).Value = Me.ObjCD.DigitGrouping((rowItem1.Tolerance_Down * 100) / rowItem1.qty_order_multiunit, 2)
                                                End If

                                                If Me.TxtFilter_PrimaryPurchReq.Text <> "" And Me.TxtFilter_PrimaryPurchReq.Text <> "-1" Then
                                                    'tambahkan item delivery po item, jika item pr ybs delivery date control
                                                    'untuk amannya, setiap item po meminta item delivery yang pertama saja, walaupun di pr-nya multiplen delivery date
                                                    If Me.ObjPurcReq.ApakahDeliveryDateControl(Me.TxtFilter_PrimaryPurchReq.Text, rowDGV.Cells(Me.ColMP_PrimaryCodeNo.Index).Value) = True Then
                                                        .Cells(i.ColDeliveryCont.Index).Value = Boolean.TrueString
                                                        Me.ObjPurcReq.IsikanDeliveryDateData(Me.TxtFilter_PrimaryPurchReq.Text, _
                                                        rowDGV.Cells(Me.ColMP_PrimaryCodeNo.Index).Value, TBLItemDeliveryPR)
                                                        If TBLItemDeliveryPR.Rows.Count >= 1 Then
                                                            rowItemDeliveryPR = TBLItemDeliveryPR.Rows(0)
                                                            .Cells(i.ColDeliveryDate.Index).Value = Format(rowItemDeliveryPR.delivery_date, "dd MMM yyyy")
                                                            vrowItemDeliveryPO = i.vTBLItemDelivery.Newtbltrans_po_itemdeliveryRow

                                                            With vrowItemDeliveryPO
                                                                .delay = 0
                                                                .delivery_date = rowItemDeliveryPR.delivery_date
                                                                .delivery_qty = rowDGVPO.Cells(i.ColQty.Index).Value
                                                                .delivery_qty_multi = rowDGVPO.Cells(i.ColQtyMulti.Index).Value
                                                                .identifier = rowDGVPO.Cells(i.ColIdentifier.Index).Value
                                                                .primary = i.vTBLItemDelivery.Rows.Count
                                                                .primary_transaksi = -1
                                                                .remarks = ""
                                                            End With
                                                            i.vTBLItemDelivery.Rows.Add(vrowItemDeliveryPO)
                                                        Else
                                                            MsgBox("Logic error, please contact your vendor if you this message.", MsgBoxStyle.Information, "Information")
                                                            Exit For
                                                        End If
                                                    End If
                                                End If

                                                Exit For
                                            End If
                                        Next
                                    End With
                                End If
                            End If
                        End If
                    Next

                    'fokuskan ke baris terakhir
                    While i.dgvItem.Rows(i.dgvItem.Rows.Count - 1).Displayed = False _
                    Or i.dgvItem.DisplayedRowCount(True) <> i.dgvItem.DisplayedRowCount(False)
                        i.dgvItem.FirstDisplayedScrollingRowIndex = i.dgvItem.FirstDisplayedScrollingRowIndex + 1
                    End While

                    'hitung total
                    i.Hitung_Total()

                    'i.Location = i.BeginingLocation
                    i.Hide()

                    If FrmPOStyle.rboSetDeliveryDate.Checked = True Or FrmPOStyle.rboSetDeliveryDate2.Checked = True Then
                        'tentukan delivery date
                        Me.AddDeliveryDate(TBLItemDeliveryDummy, i)
                        If TBLItemDeliveryDummy.Rows.Count > 0 Then
                            AddItemDetailDelivery(i, TBLItemDeliveryDummy)
                        End If
                    End If

                    ListDeliveryDate.Clear()
                    For Each rowItemDelivery In i.vTBLItemDelivery.Rows
                        If DeliveryDateDummy.Date <> Date.MinValue.Date Then
                            If rowItemDelivery.delivery_date.Date <> DeliveryDateDummy.Date Then
                                AdaBedaDeliveryDate = True
                            End If
                            If ListDeliveryDate.Contains(rowItemDelivery.delivery_date.Date) = False Then
                                ListDeliveryDate.Add(rowItemDelivery.delivery_date.Date)
                            End If
                        Else
                            DeliveryDateDummy = rowItemDelivery.delivery_date.Date
                            ListDeliveryDate.Add(rowItemDelivery.delivery_date.Date)
                        End If
                    Next

                    If FrmPOStyle.rboSetDeliveryDate.Checked = True Or FrmPOStyle.rboNoDeliveryDate.Checked = True Or (FrmPOStyle.rboSetDeliveryDate2.Checked = True And AdaBedaDeliveryDate = False) Then
                        If i.ShowDialog() = Windows.Forms.DialogResult.Abort Then
                            i.Dispose()
                            Exit For
                        End If

                        If i.DialogResult = Windows.Forms.DialogResult.OK Then
                            For Each rowDgvItemPO As DataGridViewRow In i.dgvItem.Rows
                                For Each rowDgvItemMP As DataGridViewRow In Me.dgvMP.Rows
                                    If rowDgvItemPO.Cells(i.ColPrimaryCodeNo.Index).Value = rowDgvItemMP.Cells(Me.ColMP_PrimaryCodeNo.Index).Value And rowDgvItemMP.Cells(Me.ColMP_FollowUp.Index).Value = True And rowDgvItemMP.Cells(Me.ColMP_ProcurementType.Index).Value = "Buy" Then
                                        Me.vRowItemFollowUp = Me.vDT_ItemFollowUp.Newtbltrans_material_planning_item_followupRow
                                        Me.vRowItemFollowUp.primary_transaksi = Me.TxtMP_Primary.Text
                                        Me.vRowItemFollowUp.primary_product = rowDgvItemPO.Cells(i.ColPrimaryCodeNo.Index).Value
                                        Me.vRowItemFollowUp.tipe = ClsTransaksiMaterialPlanning.EnumActionType.BUY_CREATE_PO
                                        Me.vRowItemFollowUp.primary_po_mo = ObjPO.GetPrimary(i.TxtPONo.Text)
                                        Me.vRowItemFollowUp.no_po_mo = i.TxtPONo.Text
                                        Me.vRowItemFollowUp.tanggal_po_mo = i.dtpDate.Value
                                        Me.vRowItemFollowUp.identifier_jo = String.Empty
                                        Me.vRowItemFollowUp.no_jo = String.Empty
                                        Me.vRowItemFollowUp.qty_default = rowDgvItemPO.Cells(i.ColQty.Index).Value
                                        Me.vRowItemFollowUp.primary_unit_default = rowDgvItemPO.Cells(i.ColPrimaryUnits.Index).Value
                                        Me.vRowItemFollowUp.ratio_default = rowDgvItemPO.Cells(i.ColRatioDefault.Index).Value
                                        Me.vRowItemFollowUp.qty_multi = rowDgvItemPO.Cells(i.ColQtyMulti.Index).Value
                                        Me.vRowItemFollowUp.primary_unit_multi = rowDgvItemPO.Cells(i.ColPrimaryUnitMulti.Index).Value
                                        Me.vRowItemFollowUp.ratio_multi = rowDgvItemPO.Cells(i.ColRatioReference.Index).Value
                                        Me.vRowItemFollowUp.user_code = Me.UserCode
                                        Me.vRowItemFollowUp.followup_time = Now
                                        Me.vRowItemFollowUp.remarks = String.Empty
                                        Me.vDT_ItemFollowUp.Rows.Add(Me.vRowItemFollowUp)
                                        Me.SaveRequired = True
                                        Exit For
                                    End If
                                Next
                            Next
                            For Each rowDgvItemMP As DataGridViewRow In Me.dgvMP.Rows
                                Me.vDT_ItemFollowUp.DefaultView.RowFilter = "[primary transaksi] = " & Me.TxtMP_Primary.Text & " and [primary product] = " & rowDgvItemMP.Cells(Me.ColMP_PrimaryCodeNo.Index).Value
                                If Me.vDT_ItemFollowUp.DefaultView.Count = 0 Then
                                    rowDgvItemMP.Cells(Me.ColMP_DetailFollowUp.Index).Value = "-"
                                ElseIf Me.vDT_ItemFollowUp.DefaultView.Count = 1 Then
                                    Me.vRowItemFollowUp = Me.vDT_ItemFollowUp.DefaultView(0).Row
                                    rowDgvItemMP.Cells(Me.ColMP_DetailFollowUp.Index).Value = vRowItemFollowUp.no_po_mo
                                ElseIf Me.vDT_ItemFollowUp.DefaultView.Count > 1 Then
                                    rowDgvItemMP.Cells(Me.ColMP_DetailFollowUp.Index).Value = "Multiple"
                                End If
                            Next
                        End If

                        i.Dispose()
                    Else
                        'pecah berdasarkan delivery date
                        CreateIdx = 0
                        For Each DeliveryDateDummy In ListDeliveryDate
                            Dim aj As New FrmTransaksiPO
                            RetVal = Me.BuatPOWithDelivery(DeliveryDateDummy, i, CreateIdx, ij, TBLKombinasi, aj)
                            If RetVal = Windows.Forms.DialogResult.OK Then
                                'buat item realisasi po
                                For Each rowDgvItemPO As DataGridViewRow In aj.dgvItem.Rows
                                    For Each rowDgvItemMP As DataGridViewRow In Me.dgvMP.Rows
                                        If rowDgvItemPO.Cells(aj.ColPrimaryCodeNo.Index).Value = rowDgvItemMP.Cells(Me.ColMP_PrimaryCodeNo.Index).Value And _
                                        rowDgvItemMP.Cells(Me.ColMP_FollowUp.Index).Value = True And rowDgvItemMP.Cells(Me.ColMP_ProcurementType.Index).Value = "Buy" Then
                                            Me.vRowItemFollowUp = Me.vDT_ItemFollowUp.Newtbltrans_material_planning_item_followupRow
                                            Me.vRowItemFollowUp.primary_transaksi = Me.TxtMP_Primary.Text
                                            Me.vRowItemFollowUp.primary_product = rowDgvItemPO.Cells(aj.ColPrimaryCodeNo.Index).Value
                                            Me.vRowItemFollowUp.tipe = ClsTransaksiMaterialPlanning.EnumActionType.BUY_CREATE_PO
                                            Me.vRowItemFollowUp.primary_po_mo = ObjPO.GetPrimary(aj.TxtPONo.Text)
                                            Me.vRowItemFollowUp.no_po_mo = aj.TxtPONo.Text
                                            Me.vRowItemFollowUp.tanggal_po_mo = aj.dtpDate.Value
                                            Me.vRowItemFollowUp.identifier_jo = String.Empty
                                            Me.vRowItemFollowUp.no_jo = String.Empty
                                            Me.vRowItemFollowUp.qty_default = rowDgvItemPO.Cells(aj.ColQty.Index).Value
                                            Me.vRowItemFollowUp.primary_unit_default = rowDgvItemPO.Cells(aj.ColPrimaryUnits.Index).Value
                                            Me.vRowItemFollowUp.ratio_default = rowDgvItemPO.Cells(aj.ColRatioDefault.Index).Value
                                            Me.vRowItemFollowUp.qty_multi = rowDgvItemPO.Cells(aj.ColQtyMulti.Index).Value
                                            Me.vRowItemFollowUp.primary_unit_multi = rowDgvItemPO.Cells(aj.ColPrimaryUnitMulti.Index).Value
                                            Me.vRowItemFollowUp.ratio_multi = rowDgvItemPO.Cells(aj.ColRatioReference.Index).Value
                                            Me.vRowItemFollowUp.user_code = Me.UserCode
                                            Me.vRowItemFollowUp.followup_time = Now
                                            Me.vRowItemFollowUp.remarks = String.Empty
                                            Me.vDT_ItemFollowUp.Rows.Add(Me.vRowItemFollowUp)
                                            Me.SaveRequired = True
                                            Exit For
                                        End If
                                    Next
                                Next
                                For Each rowDgvItemMP As DataGridViewRow In Me.dgvMP.Rows
                                    Me.vDT_ItemFollowUp.DefaultView.RowFilter = "[primary transaksi] = " & Me.TxtMP_Primary.Text & " and [primary product] = " & rowDgvItemMP.Cells(Me.ColMP_PrimaryCodeNo.Index).Value
                                    If Me.vDT_ItemFollowUp.DefaultView.Count = 0 Then
                                        rowDgvItemMP.Cells(Me.ColMP_DetailFollowUp.Index).Value = "-"
                                    ElseIf Me.vDT_ItemFollowUp.DefaultView.Count = 1 Then
                                        Me.vRowItemFollowUp = Me.vDT_ItemFollowUp.DefaultView(0).Row
                                        rowDgvItemMP.Cells(Me.ColMP_DetailFollowUp.Index).Value = vRowItemFollowUp.no_po_mo
                                    ElseIf Me.vDT_ItemFollowUp.DefaultView.Count > 1 Then
                                        rowDgvItemMP.Cells(Me.ColMP_DetailFollowUp.Index).Value = "Multiple"
                                    End If
                                Next
                            ElseIf RetVal = Windows.Forms.DialogResult.Abort Then
                                Exit For
                            End If
                            CreateIdx += 1
                            aj.Dispose()
                        Next
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahAdaItemDelivery(ByVal vIdentifier As String, ByVal vDeliveryDate As Date, ByVal i As FrmTransaksiPO) As Boolean
        Dim rowItemDelivery As DseTransaksiPO.tbltrans_po_itemdeliveryRow

        Try
            For Each rowItemDelivery In i.vTBLItemDelivery.Rows
                If rowItemDelivery.identifier = vIdentifier And rowItemDelivery.delivery_date.Date = vDeliveryDate.Date Then
                    Return True
                End If
            Next
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Function BuatPOWithDelivery(ByVal vDeliveryDate As Date, ByVal i As FrmTransaksiPO, ByVal CreateIdx As Integer, ByVal ij As Integer, _
    ByVal TBLKombinasi As DseTransaksiMaterialPlanning2.kombinasiDataTable, ByVal aj As FrmTransaksiPO) As System.Windows.Forms.DialogResult
        Dim DefPrefCurrency As Integer = 0
        Dim rowDGV, rowDGVPO As DataGridViewRow
        Dim Lanjut As Boolean = False
        Dim IdentifierBaru As String = ""
        Dim QtyMulti, QtyDefault As Decimal
        Dim rowItemDelivery, rowItemDelivery1 As DseTransaksiPO.tbltrans_po_itemdeliveryRow

        Try
            aj.BAckColorAll = i.BAckColorAll
            aj.UserCode = i.UserCode
            aj.Show()
            aj.Location = New Point((aj.Width * -1) - 1, (aj.Height * -1) - 1)
            aj.NewPO()
            aj.SetSupplierData(i.txtPrimarySupplier.Text)
            aj.SetCurrency(i.txtPrimaryCurrency.Text)
            DefPrefCurrency = Me.ObjCurrency.GetDefaultPrecission(aj.txtPrimaryCurrency.Text)
            If Me.TxtFilter_PrimaryPurchReq.Text <> "" And Me.TxtFilter_PrimaryPurchReq.Text <> "-1" Then
                aj.TxtRemark.Text = Me.ObjPurcReq.GetRemarks(Me.TxtFilter_PrimaryPurchReq.Text)
                If aj.TxtRemark.Text.Length > 150 Then
                    aj.TxtRemark.Text = aj.TxtRemark.Text.Substring(0, 150)
                End If
            End If

            aj.lblProgress.Text = "Data " + ij.ToString + " of " + TBLKombinasi.Rows.Count.ToString + "(Delivery Date: " + Format(vDeliveryDate, "dd MMM yyyy") + ")"
            aj.prgAll.Maximum = TBLKombinasi.Rows.Count
            aj.prgAll.Value = ij
            aj.lblProgress.Visible = True
            aj.btnStopProcess.Visible = True
            aj.prgAll.Visible = True

            For Each rowDGV In i.dgvItem.Rows
                Lanjut = False
                If CreateIdx = 0 Then
                    If rowDGV.Cells(i.ColDeliveryCont.Index).Value = Boolean.FalseString Then
                        Lanjut = True
                    Else
                        If Me.ApakahAdaItemDelivery(rowDGV.Cells(i.ColIdentifier.Index).Value, vDeliveryDate, i) = True Then
                            Lanjut = True
                        End If
                    End If
                Else
                    If rowDGV.Cells(i.ColDeliveryCont.Index).Value = Boolean.TrueString Then
                        If Me.ApakahAdaItemDelivery(rowDGV.Cells(i.ColIdentifier.Index).Value, vDeliveryDate, i) = True Then
                            Lanjut = True
                        End If
                    End If
                End If

                If Lanjut = True Then
                    'buat baris baru pada form aj, dgn qty multi = qty default = 0
                    'tambahkan baris baru
                    Dim ExpiredDateStr As String
                    If aj.dtpDeliveryDate.Checked = False Then
                        ExpiredDateStr = "-"
                    Else
                        ExpiredDateStr = Format(#1/1/1950#, "dd MMM yyyy")
                    End If
                    'cegah jangan sampe primary IPO kosong sama sekali
                    If aj.txtPrimaryIPO.Text = "" Then
                        aj.txtPrimaryIPO.Text = "-1"
                    End If
                    IdentifierBaru = aj.GetIdentifierBaru
                    aj.dgvItem.Rows.Add()
                    rowDGVPO = aj.dgvItem.Rows(aj.dgvItem.Rows.Count - 1)

                    With rowDGVPO
                        If rowDGV.Cells(i.ColDeliveryCont.Index).Value = Boolean.TrueString Then
                            .Cells(aj.ColAmount.Index).Value = 0
                            .Cells(aj.ColQty.Index).Value = 0
                            .Cells(aj.ColQtyMulti.Index).Value = 0
                            .Cells(aj.ColToleranceDown.Index).Value = 0
                            .Cells(aj.ColToleranceUp.Index).Value = 0
                        Else
                            .Cells(aj.ColAmount.Index).Value = rowDGV.Cells(i.ColAmount.Index).Value
                            .Cells(aj.ColQty.Index).Value = rowDGV.Cells(i.ColQty.Index).Value
                            .Cells(aj.ColQtyMulti.Index).Value = rowDGV.Cells(i.ColQtyMulti.Index).Value
                            .Cells(aj.ColToleranceDown.Index).Value = rowDGV.Cells(i.ColToleranceDown.Index).Value
                            .Cells(aj.ColToleranceUp.Index).Value = rowDGV.Cells(i.ColToleranceUp.Index).Value
                        End If

                        .Cells(aj.ColCodeNo.Index).Value = rowDGV.Cells(i.ColCodeNo.Index).Value
                        .Cells(aj.ColDeliveryCont.Index).Value = rowDGV.Cells(i.ColDeliveryCont.Index).Value
                        .Cells(aj.ColDeliveryDate.Index).Value = "-"
                        .Cells(aj.ColDescription.Index).Value = rowDGV.Cells(i.ColDescription.Index).Value
                        .Cells(aj.ColDiscount.Index).Value = rowDGV.Cells(i.ColDiscount.Index).Value
                        .Cells(aj.ColDiscAmount.Index).Value = rowDGV.Cells(i.ColDiscAmount.Index).Value
                        .Cells(aj.ColDiscFromPerct.Index).Value = rowDGV.Cells(i.ColDiscFromPerct.Index).Value
                        .Cells(aj.ColIPONo.Index).Value = ""
                        .Cells(aj.ColNettPrice.Index).Value = rowDGV.Cells(i.ColNettPrice.Index).Value
                        .Cells(aj.ColNo.Index).Value = aj.dgvItem.Rows.Count
                        .Cells(aj.ColPrimaryCodeNo.Index).Value = rowDGV.Cells(i.ColPrimaryCodeNo.Index).Value
                        .Cells(aj.ColPrimaryIPO.Index).Value = -1
                        .Cells(aj.ColPrimaryUnits.Index).Value = rowDGV.Cells(i.ColPrimaryUnits.Index).Value
                        If Me.TxtFilter_PrimaryPurchReq.Text <> "" And Me.TxtFilter_PrimaryPurchReq.Text <> "-1" Then
                            .Cells(aj.ColRemark.Index).Value = rowDGV.Cells(i.ColRemark.Index).Value
                        Else
                            .Cells(aj.ColRemark.Index).Value = ""
                        End If
                        .Cells(aj.ColUnit.Index).Value = rowDGV.Cells(i.ColUnit.Index).Value
                        .Cells(aj.ColUnitPrice.Index).Value = rowDGV.Cells(i.ColUnitPrice.Index).Value
                        .Cells(aj.colUserSecondaryUnits.Index).Value = Boolean.FalseString
                        .Cells(aj.ColIdentifier.Index).Value = IdentifierBaru

                        ''@rosadinnn
                        ''@MUaT
                        .Cells(aj.ColPrimaryUnitMulti.Index).Value = rowDGV.Cells(i.ColPrimaryUnitMulti.Index).Value
                        .Cells(aj.ColUnitCodeMulti.Index).Value = rowDGV.Cells(i.ColUnitCodeMulti.Index).Value
                        '---------------------------------------------------------
                        '---------------------------------------------------------
                        '---------------------------------------------------------
                        .Cells(aj.ColRatioDefault.Index).Value = rowDGV.Cells(i.ColRatioDefault.Index).Value
                        .Cells(aj.ColRatioReference.Index).Value = rowDGV.Cells(i.ColRatioReference.Index).Value
                        '---------------------------------------------------------
                        aj.SetRatioAndQtyDefault(rowDGVPO, True)

                        .Cells(aj.ColToleranceUpPerc.Index).Value = rowDGV.Cells(i.ColToleranceUpPerc.Index).Value
                        .Cells(aj.ColToleranceDownPerc.Index).Value = rowDGV.Cells(i.ColToleranceDownPerc.Index).Value
                    End With

                    For Each rowItem1 As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item1Row In Me.vDT.Rows
                        If rowItem1.primary_code_no = rowDGVPO.Cells(aj.ColPrimaryCodeNo.Index).Value Then
                            If rowDGV.Cells(i.ColDeliveryCont.Index).Value = Boolean.TrueString Then
                                QtyMulti = 0
                                QtyDefault = 0
                                For Each rowItemDelivery In i.vTBLItemDelivery.Rows
                                    If rowItemDelivery.identifier = rowDGV.Cells(i.ColIdentifier.Index).Value And rowItemDelivery.delivery_date.Date = vDeliveryDate Then
                                        QtyMulti += rowItemDelivery.delivery_qty_multi
                                        QtyDefault += rowItemDelivery.delivery_qty
                                        rowItemDelivery1 = aj.vTBLItemDelivery.Newtbltrans_po_itemdeliveryRow
                                        For mn As Integer = 0 To aj.vTBLItemDelivery.Columns.Count - 1
                                            rowItemDelivery1.Item(mn) = rowItemDelivery.Item(mn)
                                        Next
                                        rowItemDelivery1.identifier = IdentifierBaru
                                        aj.vTBLItemDelivery.Rows.Add(rowItemDelivery1)
                                    End If
                                Next

                                'isikan nilai qty multi dan qty default
                                rowDGVPO.Cells(aj.ColQty.Index).Value = Me.ObjCD.DigitGrouping(QtyDefault, rowItem1.default_prec_unit, ClsCekData.EnumTipeRounding.Middlepoint)
                                rowDGVPO.Cells(aj.ColQtyMulti.Index).Value = Me.ObjCD.DigitGrouping(QtyMulti, rowItem1.def_prec_unit_multi, ClsCekData.EnumTipeRounding.Middlepoint)
                                'update kolom delivery date
                                aj.vTBLItemDelivery.DefaultView.RowFilter = "identifier = '" + rowDGVPO.Cells(aj.ColIdentifier.Index).Value.ToString + "'"
                                If aj.vTBLItemDelivery.DefaultView.Count = 1 Then
                                    rowItemDelivery1 = aj.vTBLItemDelivery.DefaultView(0).Row
                                    rowDGVPO.Cells(aj.ColDeliveryDate.Index).Value = Format(rowItemDelivery1.delivery_date, "dd MMM yyyy")
                                Else
                                    rowDGVPO.Cells(aj.ColDeliveryDate.Index).Value = "Multiple"
                                End If

                                'update kolom tolerance qty -> ini tergantung multi unit atau ga
                                rowDGVPO.Cells(aj.ColToleranceDown.Index).Value = Me.ObjCD.DigitGrouping((CDec(rowDGVPO.Cells(aj.ColToleranceDownPerc.Index).Value) * CDec(rowDGVPO.Cells(aj.ColQtyMulti.Index).Value)) / 100, _
                                rowItem1.def_prec_unit_multi)
                                rowDGVPO.Cells(aj.ColToleranceUp.Index).Value = Me.ObjCD.DigitGrouping((CDec(rowDGVPO.Cells(aj.ColToleranceUpPerc.Index).Value) * CDec(rowDGVPO.Cells(aj.ColQtyMulti.Index).Value)) / 100, _
                                rowItem1.def_prec_unit_multi)
                            End If


                            'hitung sub total per baris
                            aj.Hitung_Amount(rowDGVPO)

                            Exit For
                        End If
                    Next
                End If
            Next

            While aj.dgvItem.Rows(aj.dgvItem.Rows.Count - 1).Displayed = False _
                  Or aj.dgvItem.DisplayedRowCount(True) <> aj.dgvItem.DisplayedRowCount(False)
                aj.dgvItem.FirstDisplayedScrollingRowIndex = aj.dgvItem.FirstDisplayedScrollingRowIndex + 1
            End While

            'hitung total
            aj.Hitung_Total()
            aj.Hide()

            Return aj.ShowDialog()

            Return Windows.Forms.DialogResult.Abort
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return Windows.Forms.DialogResult.Abort
        End Try
    End Function

    Private Sub AddItemDetailDelivery(ByVal i As FrmTransaksiPO, ByVal TBLItemDeliveryDummy As DseTransaksiPO.tbltrans_po_itemdeliveryDataTable)
        Dim rowDGV As DataGridViewRow
        Dim DV As Data.DataRowView
        Dim DefPrec As Integer = 0
        Dim RatioDefault, RatioReference As Decimal
        Dim TotalDeliveryMulti, TotalDeliveryMulti1, TotalDelivery As Decimal
        Dim rowItemDeliveryDummy, rowItemDelivery As DseTransaksiPO.tbltrans_po_itemdeliveryRow

        Try
            '2 kondisi yang harus diperhatikan:
            'untuk multi unit, pastikan nilainya minimal 1 satuan terkecil
            'untuk multi unit pastikan total qty multi = total qty multi di tabel delivery
            For Each rowDGV In i.dgvItem.Rows
                TBLItemDeliveryDummy.DefaultView.RowFilter = "identifier = '" + rowDGV.Cells(i.ColIdentifier.Index).Value.ToString + "'"
                If TBLItemDeliveryDummy.DefaultView.Count > 0 Then
                    If rowDGV.Cells(i.ColPrimaryUnits.Index).Value = rowDGV.Cells(i.ColPrimaryUnitMulti.Index).Value Then
                        'bukan multi unit, delivery qty multi = delivery qty
                        For Each DV In TBLItemDeliveryDummy.DefaultView
                            rowItemDeliveryDummy = DV.Row
                            rowItemDeliveryDummy.delivery_qty_multi = rowItemDeliveryDummy.delivery_qty
                        Next
                    Else
                        'multi unit, harus consider 2 kondisi di atas
                        '1. hitung delivery qty multi, pastikan tidak nol
                        DefPrec = Me.ObjProd.GetAngkaDecimal(rowDGV.Cells(i.ColPrimaryCodeNo.Index).Value, rowDGV.Cells(i.ColPrimaryUnitMulti.Index).Value)
                        RatioDefault = CDec(rowDGV.Cells(i.ColRatioDefault.Index).Value)
                        RatioReference = CDec(rowDGV.Cells(i.ColRatioReference.Index).Value)
                        For Each DV In TBLItemDeliveryDummy.DefaultView
                            rowItemDeliveryDummy = DV.Row
                            With rowItemDeliveryDummy
                                .delivery_qty_multi = Math.Round((.delivery_qty * RatioReference) / RatioDefault, DefPrec, MidpointRounding.AwayFromZero)
                            End With
                        Next
                        If TBLItemDeliveryDummy.DefaultView.Count > 1 Then
                            For ab As Double = TBLItemDeliveryDummy.DefaultView.Count - 1 To 0 Step -1
                                If ab >= 1 Then
                                    rowItemDeliveryDummy = TBLItemDeliveryDummy.DefaultView(ab).Row
                                    rowItemDelivery = TBLItemDeliveryDummy.DefaultView(ab - 1).Row
                                Else
                                    rowItemDeliveryDummy = TBLItemDeliveryDummy.DefaultView(ab).Row
                                    rowItemDelivery = TBLItemDeliveryDummy.DefaultView(ab + 1).Row
                                End If
                                If rowItemDeliveryDummy.delivery_qty_multi = 0 Then
                                    rowItemDelivery.delivery_qty += rowItemDeliveryDummy.delivery_qty
                                    If rowItemDelivery.remarks.Length + rowItemDeliveryDummy.remarks.Length <= 250 - 2 Then
                                        rowItemDelivery.remarks = rowItemDelivery.remarks + "; " + rowItemDeliveryDummy.remarks
                                        rowItemDelivery.delivery_qty_multi = Math.Round((rowItemDelivery.delivery_qty * RatioReference) / RatioDefault, _
                                        DefPrec, MidpointRounding.AwayFromZero)
                                    End If
                                    TBLItemDeliveryDummy.Rows.Remove(rowItemDeliveryDummy)
                                End If
                            Next
                        End If


                        '2. pastikan total delivery qty multi = total delivery qty multi pada rowdgv, begitu juga total delivery qty defaultnya
                        TotalDelivery = 0
                        TotalDeliveryMulti = 0
                        TotalDeliveryMulti1 = 0
                        For ab As Integer = 0 To TBLItemDeliveryDummy.DefaultView.Count - 1
                            rowItemDeliveryDummy = TBLItemDeliveryDummy.DefaultView(ab).Row
                            TotalDelivery += rowItemDeliveryDummy.delivery_qty
                            TotalDeliveryMulti += rowItemDeliveryDummy.delivery_qty_multi
                            If ab <> TBLItemDeliveryDummy.Count - 1 Then
                                TotalDeliveryMulti1 += rowItemDeliveryDummy.delivery_qty_multi
                            End If
                        Next
                        If TotalDelivery <> CDec(rowDGV.Cells(i.ColQty.Index).Value) Then
                            MsgBox("Logic error: total delivery qty doesn't match detail delivery qty. Please contact your vendor if you see this message.", _
                            MsgBoxStyle.Information, "Information")
                            Exit Sub
                        End If
                        If TotalDeliveryMulti <> CDec(rowDGV.Cells(i.ColQtyMulti.Index).Value) Then
                            rowItemDeliveryDummy = TBLItemDeliveryDummy.DefaultView(TBLItemDeliveryDummy.DefaultView.Count - 1).Row
                            rowItemDeliveryDummy.delivery_qty_multi = TotalDeliveryMulti - TotalDeliveryMulti1
                        End If
                    End If
                    For Each DV In TBLItemDeliveryDummy.DefaultView
                        rowItemDeliveryDummy = DV.Row
                        rowItemDelivery = i.vTBLItemDelivery.Newtbltrans_po_itemdeliveryRow
                        With rowItemDelivery
                            .delay = rowItemDeliveryDummy.delay
                            .delivery_date = rowItemDeliveryDummy.delivery_date
                            .delivery_qty = rowItemDeliveryDummy.delivery_qty
                            .delivery_qty_multi = rowItemDeliveryDummy.delivery_qty_multi
                            .identifier = rowItemDeliveryDummy.identifier
                            .primary = rowItemDeliveryDummy.primary
                            .primary_transaksi = i.TxtPrimaryPurchaseOrder.Text
                            .remarks = rowItemDeliveryDummy.remarks
                        End With
                        i.vTBLItemDelivery.Rows.Add(rowItemDelivery)
                    Next

                    rowDGV.Cells(i.ColDeliveryCont.Index).Value = Boolean.TrueString
                    If TBLItemDeliveryDummy.DefaultView.Count > 1 Then
                        rowDGV.Cells(i.ColDeliveryDate.Index).Value = "Multiple"
                    Else
                        rowItemDeliveryDummy = TBLItemDeliveryDummy.DefaultView(0).Row
                        rowDGV.Cells(i.ColDeliveryDate.Index).Value = Format(rowItemDeliveryDummy.delivery_date, "dd MMM yyyy")
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnMP_Print_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMP_Print.Click
        Dim i As New FrmTrans_MaterialPlanningDashboard_Print

        Try
            If Me.TxtMP_DocNo.Text = "" Then
                Me.btnMP_Find.PerformClick()
            End If

            If Me.TxtMP_DocNo.Text <> "" Then
                i.vPrimaryTransaksi = Me.TxtMP_Primary.Text
                i.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UpdatePerhitungan(ByVal vPrimaryProduct As Double, ByVal rowDGVMP As DataGridViewRow)
        Dim vRowDT As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item1Row
        Dim DV As Data.DataRowView
        Dim AdaOrderPoint As Boolean = False
        Dim AdaNextOrderPoint As Boolean = False

        Try
            Me.ObjCls.HitungMaterialPlanning(Me.vDT, Me.vDT_ItemGIT, Me.vDT_ItemDO, Me.vDT_ItemPOB, Me.vDT_ItemSOB, Me.vDT_ItemPrOB, Me.vDT_ItemSettingLocation, _
                                    Me.vDT_ItemQty, Me.vDT_ItemMOB, Me.vDT_ItemRMReq, Me.ObjSetting, Me.TxtMP_DaysPerMonth, Me.BackColorAll, Me.FrmCek, Me.UserCode, Me.ObjCD, _
                                    Me.TxtMP_Primary, Me.ObjSettingUser, Me.dtpMP_RefDate, vPrimaryProduct)

            Me.vDT.DefaultView.RowFilter = ""

            For Each DV In vDT.DefaultView
                vRowDT = DV.Row
                If vRowDT.primary_code_no = rowDGVMP.Cells(Me.ColMP_PrimaryCodeNo.Index).Value Then
                    With rowDGVMP
                        If vRowDT.modify = 1 Then
                            .Cells(Me.ColMP_FollowUp.Index).Value = Boolean.TrueString
                        Else
                            .Cells(Me.ColMP_FollowUp.Index).Value = Boolean.FalseString
                        End If
                        .Cells(Me.ColMP_Desc.Index).Value = vRowDT.Description
                        .Cells(Me.ColMP_CodeNo.Index).Value = vRowDT.Product_Code
                        .Cells(Me.ColMP_Currency.Index).Value = vRowDT.Currency_Code
                        .Cells(Me.ColMP_DefPrecOrderUnit.Index).Value = vRowDT.def_prec_order_unit
                        .Cells(Me.ColMP_Disc.Index).Value = ObjCD.DigitGrouping(vRowDT.discount, 2)
                        .Cells(Me.ColMP_NettPrice.Index).Value = ObjCD.DigitGrouping(vRowDT.nett_price, vRowDT.def_prec_currency)
                        .Cells(Me.ColMP_No.Index).Value = Me.dgvMP.Rows.Count
                        If vRowDT.order_point = #1/1/2200# Then
                            .Cells(Me.ColMP_OrderPoint.Index).Value = "-"
                        Else
                            .Cells(Me.ColMP_OrderPoint.Index).Value = vRowDT.order_point
                            AdaOrderPoint = True
                        End If
                        .Cells(Me.ColMP_OrderUnit.Index).Value = vRowDT.Order_Unit_Code
                        .Cells(Me.ColMP_PreferedSupplier.Index).Value = vRowDT.Supplier_Name
                        .Cells(Me.ColMP_PrimaryOrderUnit.Index).Value = vRowDT.primary_order_unit
                        If vRowDT.procurement_type = Core_Module.clsMasterProduct.EnumProcurementType.Buy Then
                            .Cells(Me.ColMP_ProcurementType.Index).Value = "Buy"
                        Else
                            .Cells(Me.ColMP_ProcurementType.Index).Value = "Produce"
                        End If
                        .Cells(Me.ColMP_QtyGIT.Index).Value = vRowDT.qty_git
                        .Cells(Me.ColMP_QtyInquiry.Index).Value = vRowDT.qty_inquiry
                        .Cells(Me.ColMP_QtyLeadTime.Index).Value = vRowDT.qty_lead_time
                        .Cells(Me.ColMP_QtyMax.Index).Value = vRowDT.qty_max
                        .Cells(Me.ColMP_QtyMin.Index).Value = vRowDT.qty_min
                        .Cells(Me.ColMP_QtyOnHand.Index).Value = vRowDT.qty_on_hand
                        .Cells(Me.ColMP_QtyOpt.Index).Value = vRowDT.qty_optimum
                        .Cells(Me.ColMP_QtyOrder.Index).Value = Me.ObjCD.DigitGrouping(vRowDT.qty_order, vRowDT.def_prec_order_unit)
                        .Cells(Me.ColMP_QtyOrderMaximum.Index).Value = vRowDT.qty_order_maksimum
                        .Cells(Me.ColMP_QtyOrderMinimum.Index).Value = vRowDT.qty_order_minimum
                        .Cells(Me.ColMP_QtyOrderOptimum.Index).Value = vRowDT.qty_order_optimum
                        .Cells(Me.ColMP_QtyPendingDO.Index).Value = vRowDT.qty_pending_do
                        .Cells(Me.ColMP_QtyPOB.Index).Value = vRowDT.qty_pob
                        .Cells(Me.ColMP_QtyRequired.Index).Value = Me.ObjCD.DigitGrouping(vRowDT.qty_required, vRowDT.default_prec_unit)
                        .Cells(Me.ColMP_QtySafetyStock.Index).Value = vRowDT.qty_safety_stock
                        .Cells(Me.ColMP_QtySOB.Index).Value = vRowDT.qty_sob
                        .Cells(Me.ColMP_QtyPrOB.Index).Value = vRowDT.qty_prob

                        .Cells(Me.ColMP_QtyOrderMulti.Index).Value = vRowDT.qty_order_multiunit
                        .Cells(Me.ColMP_PrimaryUnitMulti.Index).Value = vRowDT.primary_unit_multi
                        .Cells(Me.ColMP_DefPrecUnitMulti.Index).Value = vRowDT.def_prec_unit_multi
                        .Cells(Me.ColMP_RatioDefault.Index).Value = vRowDT.ratio_default
                        .Cells(Me.ColMP_RatioMulti.Index).Value = vRowDT.ratio_multi

                        If vRowDT.order_strategy = Core_Module.clsMasterProduct.EnumOrderStrategy.Order_By_Maximum_Order_Qty Then
                            .Cells(Me.ColMP_OrderStrategy.Index).Value = "Order By Maximum Order Qty"
                        ElseIf vRowDT.order_strategy = Core_Module.clsMasterProduct.EnumOrderStrategy.Order_By_Minimum_Order_Qty Then
                            .Cells(Me.ColMP_OrderStrategy.Index).Value = "Order By Minimum Order Qty"
                        ElseIf vRowDT.order_strategy = Core_Module.clsMasterProduct.EnumOrderStrategy.Order_By_Optimum_Order_Qty Then
                            .Cells(Me.ColMP_OrderStrategy.Index).Value = "Order By Optimum Order Qty"
                        ElseIf vRowDT.order_strategy = Core_Module.clsMasterProduct.EnumOrderStrategy.Order_To_Max_Qty Then
                            .Cells(Me.ColMP_OrderStrategy.Index).Value = "Order To Max Qty"
                        ElseIf vRowDT.order_strategy = Core_Module.clsMasterProduct.EnumOrderStrategy.Order_To_Min_Qty Then
                            .Cells(Me.ColMP_OrderStrategy.Index).Value = "Order To Min Qty"
                        ElseIf vRowDT.order_strategy = Core_Module.clsMasterProduct.EnumOrderStrategy.Order_To_Optimum_Qty Then
                            .Cells(Me.ColMP_OrderStrategy.Index).Value = "Order To Optimum Qty"
                        Else
                            .Cells(Me.ColMP_OrderStrategy.Index).Value = "Order To Safety Stock"
                        End If
                        .Cells(Me.ColMP_SubTotal.Index).Value = vRowDT.sub_total
                        .Cells(Me.ColMP_Unit.Index).Value = vRowDT.Unit_Code
                        .Cells(Me.ColMP_UnitPrice.Index).Value = vRowDT.unit_price
                        .Cells(Me.ColMP_DiscAmount.Index).Value = vRowDT.disc_amount
                        .Cells(Me.ColMP_DiscFromPerct.Index).Value = vRowDT.disc_from_perct

                        .Cells(Me.ColMP_DefPrecUnit.Index).Value = vRowDT.default_prec_unit
                        .Cells(Me.ColMP_DefPrecCurrency.Index).Value = vRowDT.def_prec_currency
                        .Cells(Me.ColMP_PrimaryCodeNo.Index).Value = vRowDT.primary_code_no
                        .Cells(Me.ColMP_PrimarySupplier.Index).Value = vRowDT.primary_prefered_supplier
                        .Cells(Me.ColMP_PrimaryCurrency.Index).Value = vRowDT.primary_currency

                        .Cells(Me.ColMP_NextOrderPointDelay.Index).Value = vRowDT.next_order_point_delay
                        If vRowDT.next_order_point = #1/1/2200# Then
                            .Cells(Me.ColMP_NextOrderPoint.Index).Value = "-"
                        Else
                            .Cells(Me.ColMP_NextOrderPoint.Index).Value = vRowDT.next_order_point
                            AdaNextOrderPoint = True
                        End If
                        .Cells(Me.ColMP_LeadTimeDays.Index).Value = vRowDT.lead_time_days
                        .Cells(Me.ColMP_OrderPointDelay.Index).Value = vRowDT.order_point_delay

                    End With

                    Exit For
                End If
            Next

            Me.dgvMP.ClearSelection()
            rowDGVMP.Selected = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_QtyPrOB_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TxtMP_QtyPrOB.MouseClick
        Dim i As New FrmTrans_MaterialPlanningDashboard_ExplainPrOB
        Dim rowDGV As DataGridViewRow

        Try
            If Me.dgvMP.SelectedCells.Count = 0 Then
                Exit Sub
            Else
                rowDGV = Me.dgvMP.Rows(Me.dgvMP.SelectedCells(0).RowIndex)

                If CDec(Me.TxtMP_QtyPrOB.Text) > 0 Then
                    i.vPrimaryProduct = rowDGV.Cells(Me.ColMP_PrimaryCodeNo.Index).Value
                    i.CodeNo = rowDGV.Cells(Me.ColMP_CodeNo.Index).Value
                    i.UnitCode = rowDGV.Cells(Me.ColMP_Unit.Index).Value
                    i.ReadOnly1 = Me.btnMP_Find.Enabled
                    i.UserCode = Me.UserCode

                    i.vDT_ItemPrOB = Me.vDT_ItemPrOB
                    i.vDT_ItemPrOBDelivery = Me.vDT_ItemPrOBDelivery
                    i.BackColorAll = Me.BackColorAll

                    i.ShowDialog()

                    If i.AdaPerubahan = True Then
                        Me.UpdatePerhitungan(i.vPrimaryProduct, rowDGV)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboMP_OrderUnit_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboMP_OrderUnit.SelectedValueChanged
        Dim UnitPrice, Discount, NettPrice, DiscAmount, DiscFromPerc As Decimal
        Dim rowDGV As DataGridViewRow
        Dim DTAPurchReq As New DseTransaksiPRTableAdapters.tbltrans_purchase_req_itemTableAdapter
        Dim TBLPurchReq As New DseTransaksiPR.tbltrans_purchase_req_itemDataTable

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DTAPurchReq.Connection.ConnectionString = Core_Module.GetConstring
            If Me.TxtFilter_PrimaryPurchReq.Text <> "-1" And Me.TxtFilter_PrimaryPurchReq.Text <> "" Then
                DTAPurchReq.FillBy_Primary(TBLPurchReq, Me.TxtFilter_PrimaryPurchReq.Text)
            End If

            'hanya hitung kalo aktifkan detail
            If Me.TxtMP_Remarks.Enabled = True Then
                'hanya hitung kalo bukan sedang loading dgv
                If Me.vHitungQtyMultiUnit = True Then
                    'hanya hitung kalo selected cell count > 0
                    If Me.dgvMP.SelectedCells.Count > 0 Then
                        For Each rowItem As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item1Row In Me.vDT.Rows
                            If rowItem.primary_code_no = Me.dgvMP.Item(Me.ColMP_PrimaryCodeNo.Index, Me.dgvMP.SelectedCells(0).RowIndex).Value Then
                                Me.rowMultiUnit = Me.tblMultiUnit.DefaultView(Me.cboMP_OrderUnit.SelectedIndex).Row
                                'ganti primary unit dan defprec
                                rowItem.primary_unit_multi = Me.rowMultiUnit.Primary_Unit
                                rowItem.def_prec_unit_multi = Me.ObjProd.GetAngkaDecimal(rowItem.primary_code_no, rowItem.primary_unit_multi)
                                Me.TxtMP_PrimaryOrderUnit.Text = rowItem.primary_unit_multi

                                'ganti digit grouping qty
                                'Me.TxtMP_QtyOrder.Text = Me.ObjCD.DigitGrouping((rowItem.qty_order_multiunit * rowItem.ratio_multi) / rowItem.ratio_default, rowItem.def_prec_unit_multi)

                                'ganti ratio qty, reference qty, dan lbl multi unit, kembali ke nilai default
                                rowItem.ratio_default = Me.rowMultiUnit.Qty_Default
                                rowItem.ratio_multi = Me.rowMultiUnit.Qty_Rasio
                                rowItem.primary_unit_multi = Me.rowMultiUnit.Primary_Unit
                                rowItem.def_prec_unit_multi = Me.rowMultiUnit.Default_Precision
                                rowItem.Unit_Code_Multi = Me.rowMultiUnit.Unit_Code
                                If rowItem.primary_unit_multi = rowItem.primary_unit Then
                                    Me.llbMP_Ratio.Text = "n/a"
                                Else
                                    Me.llbMP_Ratio.Text = Me.ObjCD.DigitGrouping(rowItem.ratio_multi, rowItem.def_prec_unit_multi) & " " & rowItem.Unit_Code_Multi & _
                                    " = " & Me.ObjCD.DigitGrouping(rowItem.ratio_default, rowItem.def_prec_order_unit) & " " & rowItem.Unit_Code + _
                                    " --> Req. Qty = " + Me.ObjCD.DigitGrouping((rowItem.qty_required * rowItem.ratio_multi) / rowItem.ratio_default, rowItem.def_prec_unit_multi) + " " + rowItem.Unit_Code_Multi
                                End If

                                Me.TxtMP_QtyOrder.Text = Me.ObjCD.DigitGrouping(Math.Round(CDec((rowItem.qty_order * rowItem.ratio_multi) / rowItem.ratio_default), _
                                CInt(rowItem.def_prec_unit_multi), MidpointRounding.AwayFromZero), CInt(rowItem.def_prec_unit_multi))
                                rowItem.qty_order_multiunit = CDec(Me.TxtMP_QtyOrder.Text)
                                rowItem.qty_order = Me.ObjCD.DigitGrouping(rowItem.qty_order_multiunit * rowItem.ratio_default / _
                                                    rowItem.ratio_multi, rowItem.default_prec_unit)

                                'ganti unit price berdasarkan data historical terakhir, lakukan perhitungan ulang untuk unit price, disc, nett price, amount
                                Me.ObjPO.GetHistoricalPrice(rowItem.primary_unit_multi, rowItem.primary_code_no, _
                                rowItem.primary_currency, UnitPrice, Discount, NettPrice, DiscAmount, DiscFromPerc)
                                Me.TxtMP_UnitPrice.Text = Me.ObjCD.DigitGrouping(UnitPrice, rowItem.def_prec_currency)
                                Me.TxtMP_Discount.Text = Me.ObjCD.DigitGrouping(Discount, 2)
                                Me.TxtMP_NettPrice.Text = Me.ObjCD.DigitGrouping(NettPrice, rowItem.def_prec_currency)
                                Me.TxtMP_DiscAmount.Text = Me.ObjCD.DigitGrouping(DiscAmount, rowItem.def_prec_currency)
                                Me.TxtMP_DiscFromPerct.Text = DiscFromPerc
                                Me.TxtMP_SubTotal.Text = Me.ObjCD.DigitGrouping(NettPrice * Me.TxtMP_QtyOrder.Text, rowItem.def_prec_currency)
                                rowItem.unit_price = Me.TxtMP_UnitPrice.Text
                                rowItem.discount = Me.TxtMP_Discount.Text
                                rowItem.disc_amount = Me.TxtMP_DiscAmount.Text
                                rowItem.disc_from_perct = Me.TxtMP_DiscFromPerct.Text
                                rowItem.nett_price = Me.TxtMP_NettPrice.Text
                                rowItem.sub_total = Me.TxtMP_SubTotal.Text

                                'tampilkan kembali rowdgv-nya
                                rowDGV = Me.dgvMP.Rows(Me.dgvMP.SelectedCells(0).RowIndex)
                                With rowDGV
                                    If rowItem.modify = 1 Then
                                        .Cells(Me.ColMP_FollowUp.Index).Value = Boolean.TrueString
                                    Else
                                        .Cells(Me.ColMP_FollowUp.Index).Value = Boolean.FalseString
                                    End If
                                    .Cells(Me.ColMP_Desc.Index).Value = rowItem.Description
                                    .Cells(Me.ColMP_CodeNo.Index).Value = rowItem.Product_Code
                                    .Cells(Me.ColMP_Currency.Index).Value = rowItem.Currency_Code
                                    .Cells(Me.ColMP_DefPrecOrderUnit.Index).Value = rowItem.def_prec_order_unit
                                    .Cells(Me.ColMP_Disc.Index).Value = ObjCD.DigitGrouping(rowItem.discount, 2)
                                    .Cells(Me.ColMP_NettPrice.Index).Value = ObjCD.DigitGrouping(rowItem.nett_price, rowItem.def_prec_currency)
                                    .Cells(Me.ColMP_No.Index).Value = Me.dgvMP.Rows.Count
                                    If rowItem.order_point = #1/1/2200# Then
                                        .Cells(Me.ColMP_OrderPoint.Index).Value = "-"
                                    Else
                                        .Cells(Me.ColMP_OrderPoint.Index).Value = rowItem.order_point
                                    End If
                                    .Cells(Me.ColMP_OrderUnit.Index).Value = rowItem.Order_Unit_Code
                                    .Cells(Me.ColMP_PreferedSupplier.Index).Value = rowItem.Supplier_Name
                                    .Cells(Me.ColMP_PrimaryOrderUnit.Index).Value = rowItem.primary_order_unit
                                    If rowItem.procurement_type = Core_Module.clsMasterProduct.EnumProcurementType.Buy Then
                                        .Cells(Me.ColMP_ProcurementType.Index).Value = "Buy"
                                    Else
                                        .Cells(Me.ColMP_ProcurementType.Index).Value = "Produce"
                                    End If
                                    .Cells(Me.ColMP_QtyGIT.Index).Value = rowItem.qty_git
                                    .Cells(Me.ColMP_QtyInquiry.Index).Value = rowItem.qty_inquiry
                                    .Cells(Me.ColMP_QtyLeadTime.Index).Value = rowItem.qty_lead_time
                                    .Cells(Me.ColMP_QtyMax.Index).Value = rowItem.qty_max
                                    .Cells(Me.ColMP_QtyMin.Index).Value = rowItem.qty_min
                                    .Cells(Me.ColMP_QtyOnHand.Index).Value = rowItem.qty_on_hand
                                    .Cells(Me.ColMP_QtyOpt.Index).Value = rowItem.qty_optimum
                                    .Cells(Me.ColMP_QtyOrder.Index).Value = Me.ObjCD.DigitGrouping(rowItem.qty_order, rowItem.def_prec_order_unit)
                                    .Cells(Me.ColMP_QtyOrderMaximum.Index).Value = rowItem.qty_order_maksimum
                                    .Cells(Me.ColMP_QtyOrderMinimum.Index).Value = rowItem.qty_order_minimum
                                    .Cells(Me.ColMP_QtyOrderOptimum.Index).Value = rowItem.qty_order_optimum
                                    .Cells(Me.ColMP_QtyPendingDO.Index).Value = rowItem.qty_pending_do
                                    .Cells(Me.ColMP_QtyPOB.Index).Value = rowItem.qty_pob
                                    .Cells(Me.ColMP_QtyRequired.Index).Value = Me.ObjCD.DigitGrouping(rowItem.qty_required, rowItem.default_prec_unit)
                                    .Cells(Me.ColMP_QtySafetyStock.Index).Value = rowItem.qty_safety_stock
                                    .Cells(Me.ColMP_QtySOB.Index).Value = rowItem.qty_sob
                                    .Cells(Me.ColMP_QtyPrOB.Index).Value = rowItem.qty_prob

                                    .Cells(Me.ColMP_QtyOrderMulti.Index).Value = rowItem.qty_order_multiunit
                                    .Cells(Me.ColMP_PrimaryUnitMulti.Index).Value = rowItem.primary_unit_multi
                                    .Cells(Me.ColMP_DefPrecUnitMulti.Index).Value = rowItem.def_prec_unit_multi
                                    .Cells(Me.ColMP_RatioDefault.Index).Value = rowItem.ratio_default
                                    .Cells(Me.ColMP_RatioMulti.Index).Value = rowItem.ratio_multi

                                    If rowItem.order_strategy = Core_Module.clsMasterProduct.EnumOrderStrategy.Order_By_Maximum_Order_Qty Then
                                        .Cells(Me.ColMP_OrderStrategy.Index).Value = "Order By Maximum Order Qty"
                                    ElseIf rowItem.order_strategy = Core_Module.clsMasterProduct.EnumOrderStrategy.Order_By_Minimum_Order_Qty Then
                                        .Cells(Me.ColMP_OrderStrategy.Index).Value = "Order By Minimum Order Qty"
                                    ElseIf rowItem.order_strategy = Core_Module.clsMasterProduct.EnumOrderStrategy.Order_By_Optimum_Order_Qty Then
                                        .Cells(Me.ColMP_OrderStrategy.Index).Value = "Order By Optimum Order Qty"
                                    ElseIf rowItem.order_strategy = Core_Module.clsMasterProduct.EnumOrderStrategy.Order_To_Max_Qty Then
                                        .Cells(Me.ColMP_OrderStrategy.Index).Value = "Order To Max Qty"
                                    ElseIf rowItem.order_strategy = Core_Module.clsMasterProduct.EnumOrderStrategy.Order_To_Min_Qty Then
                                        .Cells(Me.ColMP_OrderStrategy.Index).Value = "Order To Min Qty"
                                    ElseIf rowItem.order_strategy = Core_Module.clsMasterProduct.EnumOrderStrategy.Order_To_Optimum_Qty Then
                                        .Cells(Me.ColMP_OrderStrategy.Index).Value = "Order To Optimum Qty"
                                    Else
                                        .Cells(Me.ColMP_OrderStrategy.Index).Value = "Order To Safety Stock"
                                    End If
                                    .Cells(Me.ColMP_SubTotal.Index).Value = rowItem.sub_total
                                    .Cells(Me.ColMP_Unit.Index).Value = rowItem.Unit_Code
                                    .Cells(Me.ColMP_UnitPrice.Index).Value = rowItem.unit_price
                                    .Cells(Me.ColMP_DiscAmount.Index).Value = rowItem.disc_amount
                                    .Cells(Me.ColMP_DiscFromPerct.Index).Value = rowItem.disc_from_perct

                                    .Cells(Me.ColMP_NextOrderPointDelay.Index).Value = rowItem.next_order_point_delay
                                    If rowItem.next_order_point = #1/1/2200# Then
                                        .Cells(Me.ColMP_NextOrderPoint.Index).Value = "-"
                                    Else
                                        .Cells(Me.ColMP_NextOrderPoint.Index).Value = rowItem.next_order_point
                                    End If
                                    .Cells(Me.ColMP_LeadTimeDays.Index).Value = rowItem.lead_time_days
                                    .Cells(Me.ColMP_OrderPointDelay.Index).Value = rowItem.order_point_delay
                                    If rowItem.procurement_strategy = Core_Module.clsMasterProduct.EnumProcurementStrategy.Build_To_Order Then
                                        .Cells(Me.ColMP_ProcurementStrategy.Index).Value = "Build To Order"
                                    Else
                                        .Cells(Me.ColMP_ProcurementStrategy.Index).Value = "Build To Stock"
                                    End If

                                    If Me.TxtMP_Remarks.Enabled = True Then
                                        If Me.TxtFilter_PrimaryPurchReq.Text <> "" And Me.TxtFilter_PrimaryPurchReq.Text <> "-1" Then
                                            Me.Buffer = .Cells(Me.ColMP_QtyOrder.Index).Value
                                            .Cells(Me.ColMP_QtyOrder.Index).Value = Me.GetQty(TBLPurchReq, rowItem.primary_code_no)
                                            Me.EndEdit_Qty(Me, rowDGV.Index, Me.ColMP_QtyOrder.Index)
                                        End If
                                    End If
                                End With

                                Exit Sub
                            End If
                        Next
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAPurchReq.Dispose()
            TBLPurchReq.Dispose()
        End Try
    End Sub

    Private Sub llbMP_Ratio_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbMP_Ratio.LinkClicked
        Dim frmRatio As New FrmTransaksiPO_SetQtyRatio
        Dim rowDGV As DataGridViewRow
        Dim BufferRatioReference As String = ""
        Dim BufferRatioDefault As String = ""

        Try
            'hanya jalan klo txt remarks enable
            If Me.TxtMP_Remarks.Enabled = True Then
                'hanya hitung kalo bukan sedang loading dgv
                If Me.vHitungQtyMultiUnit = True Then
                    'hanya hitung kalo selected cell count > 0
                    If Me.dgvMP.SelectedCells.Count > 0 Then
                        If Me.llbMP_Ratio.Text <> "n/a" Then
                            For Each rowItem As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item1Row In Me.vDT.Rows
                                If rowItem.primary_code_no = Me.dgvMP.Item(Me.ColMP_PrimaryCodeNo.Index, Me.dgvMP.SelectedCells(0).RowIndex).Value Then
                                    If rowItem.primary_unit <> rowItem.primary_unit_multi Then
                                        BufferRatioReference = Me.ObjCD.DigitGrouping(rowItem.ratio_multi, rowItem.def_prec_unit_multi)
                                        BufferRatioDefault = Me.ObjCD.DigitGrouping(rowItem.ratio_default, rowItem.def_prec_order_unit)

                                        frmRatio.vDefPrecMulti = rowItem.def_prec_unit_multi
                                        frmRatio.vDefPrecDefault = rowItem.def_prec_order_unit
                                        frmRatio.vRatioMulti = rowItem.ratio_multi
                                        frmRatio.vUnitMulti = rowItem.Unit_Code_Multi
                                        frmRatio.vRatioDefault = rowItem.ratio_default
                                        frmRatio.vUnitDefault = rowItem.Unit_Code

                                        If frmRatio.ShowDialog = Windows.Forms.DialogResult.OK Then
                                            rowItem.ratio_multi = frmRatio.TxtRatioMulti.Text
                                            rowItem.ratio_default = frmRatio.TxtRatioDefault.Text

                                            'hitung ulang llb ratio
                                            With rowItem
                                                Me.llbMP_Ratio.Text = Me.ObjCD.DigitGrouping(rowItem.ratio_multi, .def_prec_unit_multi) & " " & .Unit_Code_Multi & _
                                                " = " & Me.ObjCD.DigitGrouping(.ratio_default, .def_prec_order_unit) & " " & .Unit_Code + _
                                                " --> Req. Qty = " + Me.ObjCD.DigitGrouping((.qty_required * .ratio_multi) / .ratio_default, .def_prec_unit_multi) + " " + .Unit_Code_Multi
                                                Me.TxtMP_QtyOrder.Text = Me.ObjCD.DigitGrouping((.qty_required * .ratio_multi) / .ratio_default, .def_prec_unit_multi)
                                            End With

                                            'hitung ulang qty order
                                            rowItem.qty_order_multiunit = Me.TxtMP_QtyOrder.Text
                                            rowItem.qty_order = Math.Round((rowItem.qty_order_multiunit * rowItem.ratio_default) / rowItem.ratio_multi, _
                                            rowItem.def_prec_order_unit, MidpointRounding.AwayFromZero)

                                            'modif tampilkan qty order di dgv
                                            Me.dgvMP.Item(Me.ColMP_QtyOrder.Index, Me.dgvMP.SelectedCells(0).RowIndex).Value = _
                                            Me.ObjCD.DigitGrouping(rowItem.qty_order, rowItem.def_prec_order_unit)
                                            Me.dgvMP.Item(Me.ColMP_QtyOrderMulti.Index, Me.dgvMP.SelectedCells(0).RowIndex).Value = _
                                            Me.ObjCD.DigitGrouping(rowItem.qty_order_multiunit, rowItem.def_prec_unit_multi)

                                            'update dgv item
                                            rowDGV = Me.dgvMP.Rows(Me.dgvMP.SelectedCells(0).RowIndex)
                                            With rowDGV
                                                .Cells(Me.ColMP_RatioDefault.Index).Value = rowItem.ratio_default
                                                .Cells(Me.ColMP_RatioMulti.Index).Value = rowItem.ratio_multi
                                            End With
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            frmRatio = Nothing
        End Try
    End Sub

    Private Sub llbMP_Tolerance_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbMP_Tolerance.LinkClicked
        Dim ij As New FrmTrans_MaterialPlanningDashboard_MP_DetailTolerance

        Try
            'hanya jalan klo txt remarks enable
            If Me.TxtMP_Remarks.Enabled = True Then
                'hanya hitung kalo selected cell count > 0
                If Me.dgvMP.SelectedCells.Count > 0 Then
                    For Each rowItem As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item1Row In Me.vDT.Rows
                        If rowItem.primary_code_no = Me.dgvMP.Item(Me.ColMP_PrimaryCodeNo.Index, Me.dgvMP.SelectedCells(0).RowIndex).Value Then
                            'set value basic
                            ij.vDefPrec = rowItem.def_prec_order_unit
                            ij.vQtyLowerTolerance = rowItem.Tolerance_Down
                            ij.vQtyUpperTolerance = rowItem.Tolerance_Up
                            ij.vQtyOrder = rowItem.qty_order_multiunit
                            ij.vUnitCode = rowItem.Unit_Code_Multi

                            If ij.ShowDialog = Windows.Forms.DialogResult.OK Then
                                rowItem.Tolerance_Down = ij.TxtQty_LowerTolerance.Text
                                rowItem.Tolerance_Up = ij.TxtQty_UpperTolerance.Text
                            End If
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub LinkLabel3_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel3.LinkClicked
        Dim ij As New FrmTrans_MaterialPlanningDashboard_MP_SetVisible

        Try
            'tampilkan form untuk setting data
            ij.vUserName = Me.UserCode
            If ij.ShowDialog = Windows.Forms.DialogResult.OK Then
                'update kolom
                Me.UpdateKolomMP()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub UpdateKolomMP()
        Dim Dummy As String = ""

        Try
            Dummy = Me.ObjSettingUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.MaterialPlanning, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.MaterialPlanning_ShowProcurementStrategy, _
            Me.UserCode)
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
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.MaterialPlanning_ProcurementType, _
            Me.UserCode)
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

    Private Sub btnFilter_FindPurchReq_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter_FindPurchReq.Click
        Dim frm As New FrmTransaksiPR_Cari

        Try
            If Me.TxtFilter_PurcReqNo.Text <> "" Then
                frm.txtCari.Text = Me.TxtFilter_PurcReqNo.Text
            End If

            frm.ShowDialog()
            If frm.PrimaryPR > 0 Then
                Me.TxtFilter_PrimaryPurchReq.Text = frm.PrimaryPR
                Me.TxtFilter_PurcReqNo.Text = Me.ObjPurcReq.GetNoTransaksi(frm.PrimaryPR)
            Else
                Me.TxtFilter_PrimaryPurchReq.Text = "-1"
                Me.TxtFilter_PurcReqNo.Clear()
            End If

            If Me.TxtFilter_PrimaryPurchReq.Text <> "" And Me.TxtFilter_PrimaryPurchReq.Text <> "-1" Then
                'pastikan pr yang dipilih sudah di otorisasi
                If Me.ObjPurcReq.ApakahSudahDiOtorisasi(Me.TxtFilter_PrimaryPurchReq.Text) = False Then
                    MsgBox("Please choose authorized purchase request only.", MsgBoxStyle.Information, "Information")
                    Me.TxtFilter_PrimaryPurchReq.Text = "-1"
                    Me.TxtFilter_PurcReqNo.Clear()
                    Exit Sub
                End If
            End If

            Me.LoadUlang()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            frm.Dispose()
        End Try
    End Sub

    Private Sub TxtFilter_PurcReqNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtFilter_PurcReqNo.KeyDown
        Try
            Me.Cursor = Cursors.WaitCursor
            If e.KeyCode = 13 Then
                If Me.TxtFilter_PurcReqNo.Text <> "" Then
                    If Me.ObjPurcReq.ApakahSudahAda(Me.TxtFilter_PurcReqNo.Text) = True Then
                        Me.TxtFilter_PrimaryPurchReq.Text = Me.ObjPurcReq.GetPrimary(Me.TxtFilter_PurcReqNo.Text)
                        Me.TxtFilter_PurcReqNo.Text = Me.ObjPurcReq.GetNoTransaksi(Me.TxtFilter_PrimaryPurchReq.Text)
                    Else
                        Me.btnFilter_FindPurchReq.PerformClick()
                    End If
                Else
                    Me.TxtFilter_PrimaryPurchReq.Text = "-1"
                    Me.TxtFilter_PurcReqNo.Clear()
                End If

                If Me.TxtFilter_PrimaryPurchReq.Text <> "" And Me.TxtFilter_PrimaryPurchReq.Text <> "-1" Then
                    'pastikan pr yang dipilih sudah di otorisasi
                    If Me.ObjPurcReq.ApakahSudahDiOtorisasi(Me.TxtFilter_PrimaryPurchReq.Text) = False Then
                        MsgBox("Please choose authorized purchase request only.", MsgBoxStyle.Information, "Information")
                        Me.TxtFilter_PrimaryPurchReq.Text = "-1"
                        Me.TxtFilter_PurcReqNo.Clear()
                        Exit Sub
                    End If
                End If

                Me.LoadUlang()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub TxtFilter_PurcReqNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtFilter_PurcReqNo.Leave
        Try
            Me.Cursor = Cursors.WaitCursor
            If Me.TxtFilter_PurcReqNo.Text <> "" Then
                If Me.ObjPurcReq.ApakahSudahAda(Me.TxtFilter_PurcReqNo.Text) = True Then
                    Me.TxtFilter_PrimaryPurchReq.Text = Me.ObjPurcReq.GetPrimary(Me.TxtFilter_PurcReqNo.Text)
                    Me.TxtFilter_PurcReqNo.Text = Me.ObjPurcReq.GetNoTransaksi(Me.TxtFilter_PrimaryPurchReq.Text)
                Else
                    Me.btnFilter_FindPurchReq.PerformClick()
                End If
            Else
                Me.TxtFilter_PrimaryPurchReq.Text = "-1"
                Me.TxtFilter_PurcReqNo.Clear()
            End If

            If Me.TxtFilter_PrimaryPurchReq.Text <> "" And Me.TxtFilter_PrimaryPurchReq.Text <> "-1" Then
                'pastikan pr yang dipilih sudah di otorisasi
                If Me.ObjPurcReq.ApakahSudahDiOtorisasi(Me.TxtFilter_PrimaryPurchReq.Text) = False Then
                    MsgBox("Please choose authorized purchase request only.", MsgBoxStyle.Information, "Information")
                    Me.TxtFilter_PrimaryPurchReq.Text = "-1"
                    Me.TxtFilter_PurcReqNo.Clear()
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub btnFilter_FindMO_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter_FindMO.Click
        Dim frm As New FrmTrans_ManufacturingOrder_Cari

        Try
            If Me.TxtFilter_MONo.Text <> "" Then
                frm.txtCari.Text = Me.TxtFilter_MONo.Text
            End If

            frm.ShowDialog()
            If frm.Primary > 0 Then
                Me.TxtFilter_PrimaryMO.Text = frm.Primary
                Me.TxtFilter_MONo.Text = Me.ObjMO.GetNoTransaksi(frm.Primary)
            Else
                Me.TxtFilter_PrimaryMO.Text = "-1"
                Me.TxtFilter_MONo.Clear()
            End If

            Me.LoadUlang()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            frm.Dispose()
        End Try
    End Sub

    Private Sub TxtFilter_MONo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtFilter_MONo.KeyDown
        Try
            Me.Cursor = Cursors.WaitCursor
            If e.KeyCode = 13 Then
                If Me.TxtFilter_MONo.Text <> "" Then
                    If Me.ObjMO.Apakahsudahada(Me.TxtFilter_MONo.Text) = True Then
                        Me.TxtFilter_PrimaryMO.Text = Me.ObjMO.GetPrimary(Me.TxtFilter_MONo.Text)
                        Me.TxtFilter_MONo.Text = Me.ObjMO.GetNoTransaksi(Me.TxtFilter_PrimaryMO.Text)
                    Else
                        Me.btnFilter_FindMO.PerformClick()
                    End If
                Else
                    Me.TxtFilter_PrimaryMO.Text = "-1"
                    Me.TxtFilter_MONo.Clear()
                End If

                Me.LoadUlang()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub TxtFilter_MONo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtFilter_MONo.Leave
        Try
            Me.Cursor = Cursors.WaitCursor
            If Me.TxtFilter_MONo.Text <> "" Then
                If Me.ObjMO.Apakahsudahada(Me.TxtFilter_MONo.Text) = True Then
                    Me.TxtFilter_PrimaryMO.Text = Me.ObjMO.GetPrimary(Me.TxtFilter_MONo.Text)
                    Me.TxtFilter_MONo.Text = Me.ObjMO.GetNoTransaksi(Me.TxtFilter_PrimaryMO.Text)
                Else
                    Me.btnFilter_FindMO.PerformClick()
                End If
            Else
                Me.TxtFilter_PrimaryMO.Text = "-1"
                Me.TxtFilter_MONo.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub llbMP_CopyMP_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles llbMP_CopyMP.Click
        Dim hj As New FrmTrans_MaterialPlanningDashboard_FindMP

        Try
            hj.BackColorAll = Me.BackColorAll
            hj.ShowDialog()

            If Not hj.PrimaryTransaksi = -1 Then
                If hj.PrimaryTransaksi = Me.TxtMP_Primary.Text Then
                    MsgBox("Please choose other material planning transaction.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
                Windows.Forms.Cursor.Current = Cursors.WaitCursor

                If Me.ObjMP.CariDetail(hj.PrimaryTransaksi, Me.vDT, Me.vDT_ItemGIT, Me.vDT_ItemDO, vDT_ItemPOB, vDT_ItemSOB, _
                vDT_ItemPrOB, vDT_ItemSettingLocation, vDT_ItemQty, vDT_ItemMOB, vDT_ItemRMReq, Me.vDT_ItemFollowUp, Me.vDT_ItemSOBDelivery, _
                vDT_ItemPrOBDelivery) = True Then
                    Me.chkMP_RequiredOnly.Checked = True
                    Me.dtpMP_Filter.Value = Me.dtpMP_RefDate.Value
                    Me.LoadDGV()
                Else
                    MsgBox("Material planning transaction not found.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
                Me.TxtMP_DocNo.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub btnFilter_FindCodeNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter_FindCodeNo.Click
        Dim i As New frmCariMasterProduct

        Try
            i.cboCriteria.SelectedIndex = 0
            If Me.TxtFilter_CodeNo.Text <> "" Then
                i.txtCari.Text = Me.TxtFilter_CodeNo.Text
            End If

            i.ShowDialog()
            If i.PrimaryProduct > 0 Then
                Me.TxtFilter_CodeNo.Text = Me.ObjProd.GetProductCode(i.PrimaryProduct)
                Me.LoadUlang()
            Else
                Me.TxtFilter_CodeNo.Clear()
                Me.TxtFilter_CodeNo.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub chkFilter_PrOBLargerThanPOB_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkFilter_PrOBLargerThanPOB.CheckedChanged
        Try
            Me.LoadDGV()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llmbMP_SetQtyOrderEqualQtyPrOB_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llmbMP_SetQtyOrderEqualQtyPrOB.LinkClicked
        Try
            If Me.dgvMP.Rows.Count Then

            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_QtyOnHand_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TxtMP_QtyOnHand.MouseClick
        Dim i As New FrmTrans_MaterialPlanningDashboard_ExplainQtyOnHand
        Dim rowDGV As DataGridViewRow

        Try
            If Me.dgvMP.SelectedCells.Count = 0 Then
                Exit Sub
            Else
                rowDGV = Me.dgvMP.Rows(Me.dgvMP.SelectedCells(0).RowIndex)

                If CDec(Me.TxtMP_QtyOnHand.Text) > 0 Then
                    i.vPrimaryProduct = rowDGV.Cells(Me.ColMP_PrimaryCodeNo.Index).Value
                    i.UnitCode = rowDGV.Cells(Me.ColMP_Unit.Index).Value
                    i.DefPrec = rowDGV.Cells(Me.ColMP_DefPrecUnit.Index).Value
                    i.vDT_ItemQty = Me.vDT_ItemQty
                    i.vDT_ItemSettingLocation = Me.vDT_ItemSettingLocation
                    i.vBackColorAll = Me.BackColorAll

                    i.ShowDialog()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i = Nothing
        End Try
    End Sub

    Private Sub llbMP_LocationSetting_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbMP_LocationSetting.LinkClicked
        Dim i As New FrmTrans_MaterialPlanningDashboard_ExplainWarehouseExc

        Try
            If Me.vDT_ItemSettingLocation.Rows.Count = 0 Then
                MsgBox("No location setting for this transaction.", MsgBoxStyle.Information, "Information")
            Else
                i.TBL = Me.vDT_ItemSettingLocation
                i.ShowDialog()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i = Nothing
        End Try
    End Sub

    Private Sub TxtMP_QtyProductionBalance_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TxtMP_QtyMOB.MouseClick
        Dim i As New FrmTrans_MaterialPlanningDashboard_ExplainMOB
        Dim rowDGV As DataGridViewRow

        Try
            If Me.dgvMP.SelectedCells.Count = 0 Then
                Exit Sub
            Else
                rowDGV = Me.dgvMP.Rows(Me.dgvMP.SelectedCells(0).RowIndex)

                If CDec(Me.TxtMP_QtyMOB.Text) > 0 Then
                    i.vPrimaryProduct = rowDGV.Cells(Me.ColMP_PrimaryCodeNo.Index).Value
                    i.vDT_ItemMOB = Me.vDT_ItemMOB
                    i.BackColorAll = Me.BackColorAll

                    i.ShowDialog()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_QtyRMReq_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TxtMP_QtyRMReq.MouseClick
        Dim i As New FrmTrans_MaterialPlanningDashboard_ExplainRMReq
        Dim rowDGV As DataGridViewRow

        Try
            If Me.dgvMP.SelectedCells.Count = 0 Then
                Exit Sub
            Else
                rowDGV = Me.dgvMP.Rows(Me.dgvMP.SelectedCells(0).RowIndex)

                If CDec(Me.TxtMP_QtyRMReq.Text) > 0 Then
                    i.vPrimaryProduct = rowDGV.Cells(Me.ColMP_PrimaryCodeNo.Index).Value
                    i.vDT_ItemRMReq = Me.vDT_ItemRMReq
                    i.BackColorAll = Me.BackColorAll
                    i.UserCode = Me.UserCode
                    i.CodeNo = rowDGV.Cells(Me.ColMP_CodeNo.Index).Value
                    i.UnitCode = rowDGV.Cells(Me.ColMP_Unit.Index).Value
                    i.ReadOnly1 = Me.btnMP_Find.Enabled

                    i.ShowDialog()

                    If i.AdaPerubahan = True Then
                        Me.ObjCls.HitungMaterialPlanning(Me.vDT, Me.vDT_ItemGIT, Me.vDT_ItemDO, Me.vDT_ItemPOB, Me.vDT_ItemSOB, Me.vDT_ItemPrOB, Me.vDT_ItemSettingLocation, _
                        Me.vDT_ItemQty, Me.vDT_ItemMOB, Me.vDT_ItemRMReq, Me.ObjSetting, Me.TxtMP_DaysPerMonth, Me.BackColorAll, Me.FrmCek, Me.UserCode, Me.ObjCD, _
                        Me.TxtMP_Primary, Me.ObjSettingUser, Me.dtpMP_RefDate, i.vPrimaryProduct)
                        Me.dgvMP.ClearSelection()
                        rowDGV.Selected = True
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkMODateFilter_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkMODateFilter.CheckedChanged
        Try
            If Me.chkMODateFilter.Checked = True Then
                Me.DtpFilter_MO_From.Enabled = True
                Me.DtpFilter_MO_To.Enabled = True
                Me.Label38.Enabled = True
                Me.Label37.Enabled = True
            Else
                Me.DtpFilter_MO_From.Enabled = False
                Me.DtpFilter_MO_To.Enabled = False
                Me.Label38.Enabled = False
                Me.Label37.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkClose.Click
        Try
            If Me.TxtMP_Primary.Text = "-1" Then
                Me.chkClose.Checked = False
                Exit Sub
            End If
            If Me.TxtMP_Primary.Text <> "-1" And Me.TxtMP_Remarks.Enabled = False Then
                Dim createdBy As String = Me.ObjMP.GetCreatedBy(Me.TxtMP_Primary.Text)
                Dim frmPwd As New Core_Module.DialogPassword
                frmPwd.Text = "Close MP"
                frmPwd.txtUserName.Text = Me.UserCode
                If Me.chkClose.Checked Then
                    If Me.UserCode <> createdBy Then
                        Me.chkClose.Checked = False
                        MsgBox("Only user '" & createdBy & "' allowed to close this MP.", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                    If Me.ObjMP.GetAuthLevelNeeded(Me.TxtMP_Primary.Text) > 0 Then ''if auth needed
                        If Me.ObjMP.GetAuthLevelCurrent(Me.TxtMP_Primary.Text) < 3 Then ''if not full auth, return false
                            Me.chkClose.Checked = False
                            MsgBox("This material planning transaction need to authorize first before you close.", MsgBoxStyle.Information, "Information")
                            Exit Sub
                        End If
                    End If
                    frmPwd.ShowDialog()
                    If frmPwd.DialogResult = Windows.Forms.DialogResult.OK Then
                        If Me.ObjMP.CloseMP(Me.TxtMP_Primary.Text, True) = True Then
                            MsgBox("Material planning is closed.", MsgBoxStyle.Information, "Information")
                            Me.RefreshFollowUpControl()
                            Me.btnMP_Find.Focus()
                        Else
                            Me.chkClose.Checked = False
                            MsgBox("Failed to close material planning.", MsgBoxStyle.Information, "Information")
                            Me.btnMP_Find.Focus()
                        End If
                    Else
                        Me.chkClose.Checked = False
                    End If
                Else
                    If Me.UserCode <> createdBy Then
                        Me.chkClose.Checked = True
                        MsgBox("Only user '" & createdBy & "' allowed to open this MP.", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                    If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_MaterialPlanningDashboard_MP_Track") Then
                        If Me.ObjMP.GetOpenMPCount(Me.UserCode) > 0 Then
                            Me.chkClose.Checked = True
                            Dim docNo As String = Me.ObjMP.GetOpenMPDocNo(Me.UserCode)
                            MsgBox("There is already opened MP: " & docNo & ", you are only allowed 1 MP at a time.", MsgBoxStyle.Information, "Information")
                            Exit Sub
                        End If
                    End If
                    frmPwd.ShowDialog()
                    If frmPwd.DialogResult = Windows.Forms.DialogResult.OK Then
                        If Me.ObjMP.CloseMP(Me.TxtMP_Primary.Text, False) = True Then
                            MsgBox("Material planning is opened.", MsgBoxStyle.Information, "Information")
                            Me.RefreshFollowUpControl()
                            Me.btnMP_Find.Focus()
                        Else
                            Me.chkClose.Checked = True
                            MsgBox("Failed to open material planning.", MsgBoxStyle.Information, "Information")
                            Me.btnMP_Find.Focus()
                        End If
                    Else
                        Me.chkClose.Checked = True
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub FrmTrans_MaterialPlanningDashboard_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Try
            Me.RefreshFollowUpControl()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub RefreshFollowUpControl()
        Try
            If Me.TxtMP_Primary.Text <> "-1" Then
                Me.llbMP_CekAll.Visible = False
                Me.llbMP_AddToPO.Visible = False
                Me.llbMP_CreatePO.Visible = False
                'Me.ColMP_FollowUp.Visible = False
                If Me.ObjMP.IsMPClosed(Me.TxtMP_Primary.Text) Then
                    Me.chkClose.Checked = True
                Else
                    Me.chkClose.Checked = False
                    If Me.ObjMP.GetAuthLevelNeeded(Me.TxtMP_Primary.Text) = Me.ObjMP.GetAuthLevelCurrent(Me.TxtMP_Primary.Text) Then
                        If (Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_MaterialPlanningDashboard_MP_CreateEditPO") = True And Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiPO") = True) Or _
                        (Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_MaterialPlanningDashboard_MP_CreateEditMO") = True And Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_ManufacturingOrder") = True) = True Then
                            Me.llbMP_CreatePO.Visible = True
                            Me.llbMP_AddToPO.Visible = True
                            Me.ColMP_FollowUp.Visible = True
                            Me.llbMP_CekAll.Visible = True
                        End If
                    End If
                End If
            Else
                Me.llbMP_CreatePO.Visible = True
                Me.llbMP_AddToPO.Visible = True
                Me.ColMP_FollowUp.Visible = True
                Me.llbMP_CekAll.Visible = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Function IsAllowedToCreateOrAddPO(ByVal vPrimaryMP As Double) As Boolean
        Try
            If Me.ObjMP.IsMPClosed(vPrimaryMP) = True Then
                MsgBox("This MP is closed, you can't follow up.", MsgBoxStyle.Information, "Information")
                Return False
            End If
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_MaterialPlanningDashboard_MP_CreateEditPO") = False Then
                MsgBox("You don't have access to create or add PO.", MsgBoxStyle.Information, "Information")
                Return False
            End If
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiPO") = False Then
                MsgBox("You don't have access to purchase order transaction.", MsgBoxStyle.Information, "Information")
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function
    Private Function IsAllowedToCreateOrAddMO(ByVal vPrimaryMP As Double) As Boolean
        Try
            If Me.ObjMP.IsMPClosed(vPrimaryMP) = True Then
                MsgBox("This MP is closed, you can't follow up.", MsgBoxStyle.Information, "Information")
                Return False
            End If
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_MaterialPlanningDashboard_MP_CreateEditMO") = False Then
                MsgBox("You don't have access to create or add MO.", MsgBoxStyle.Information, "Information")
                Return False
            End If
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_ManufacturingOrder") = False Then
                MsgBox("You don't have access to manufacturing order transaction.", MsgBoxStyle.Information, "Information")
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function
    Private Function IsValidToModifyData(ByVal vPrimaryMP As Double) As Boolean
        Try
            If Me.ObjMP.GetAuthLevelNeeded(vPrimaryMP) > 0 Then
                If Me.ObjMP.GetAuthLevelCurrent(vPrimaryMP) = Me.ObjMP.GetAuthLevelNeeded(vPrimaryMP) Then
                    Return False
                End If
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Sub TxtMP_QtyPrOBSpare_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TxtMP_QtyPrOBSpare.MouseClick
        Dim i As New FrmTrans_MaterialPlanningDashboard_ExplainPrOB
        Dim rowDGV As DataGridViewRow

        Try
            If Me.dgvMP.SelectedCells.Count = 0 Then
                Exit Sub
            Else
                rowDGV = Me.dgvMP.Rows(Me.dgvMP.SelectedCells(0).RowIndex)

                If CDec(Me.TxtMP_QtyPrOB.Text) > 0 Then
                    i.vPrimaryProduct = rowDGV.Cells(Me.ColMP_PrimaryCodeNo.Index).Value
                    i.CodeNo = rowDGV.Cells(Me.ColMP_CodeNo.Index).Value
                    i.UnitCode = rowDGV.Cells(Me.ColMP_Unit.Index).Value
                    i.ReadOnly1 = Me.btnMP_Find.Enabled
                    i.UserCode = Me.UserCode

                    i.vDT_ItemPrOB = Me.vDT_ItemPrOB
                    i.vDT_ItemPrOBDelivery = Me.vDT_ItemPrOBDelivery
                    i.BackColorAll = Me.BackColorAll

                    i.ShowDialog()

                    If i.AdaPerubahan = True Then
                        Me.ObjCls.HitungMaterialPlanning(Me.vDT, Me.vDT_ItemGIT, Me.vDT_ItemDO, Me.vDT_ItemPOB, Me.vDT_ItemSOB, Me.vDT_ItemPrOB, Me.vDT_ItemSettingLocation, _
                        Me.vDT_ItemQty, Me.vDT_ItemMOB, Me.vDT_ItemRMReq, Me.ObjSetting, Me.TxtMP_DaysPerMonth, Me.BackColorAll, Me.FrmCek, Me.UserCode, Me.ObjCD, _
                        Me.TxtMP_Primary, Me.ObjSettingUser, Me.dtpMP_RefDate, i.vPrimaryProduct)
                        Me.dgvMP.ClearSelection()
                        rowDGV.Selected = True
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_QtyRMReqSpare_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TxtMP_QtyRMReqSpare.MouseClick
        Dim i As New FrmTrans_MaterialPlanningDashboard_ExplainRMReq
        Dim rowDGV As DataGridViewRow

        Try
            If Me.dgvMP.SelectedCells.Count = 0 Then
                Exit Sub
            Else
                rowDGV = Me.dgvMP.Rows(Me.dgvMP.SelectedCells(0).RowIndex)

                If CDec(Me.TxtMP_QtyRMReq.Text) > 0 Then
                    i.vPrimaryProduct = rowDGV.Cells(Me.ColMP_PrimaryCodeNo.Index).Value
                    i.vDT_ItemRMReq = Me.vDT_ItemRMReq
                    i.BackColorAll = Me.BackColorAll
                    i.UserCode = Me.UserCode
                    i.CodeNo = rowDGV.Cells(Me.ColMP_CodeNo.Index).Value
                    i.UnitCode = rowDGV.Cells(Me.ColMP_Unit.Index).Value
                    i.ReadOnly1 = Me.btnMP_Find.Enabled

                    i.ShowDialog()

                    If i.AdaPerubahan = True Then
                        Me.ObjCls.HitungMaterialPlanning(Me.vDT, Me.vDT_ItemGIT, Me.vDT_ItemDO, Me.vDT_ItemPOB, Me.vDT_ItemSOB, Me.vDT_ItemPrOB, Me.vDT_ItemSettingLocation, _
                        Me.vDT_ItemQty, Me.vDT_ItemMOB, Me.vDT_ItemRMReq, Me.ObjSetting, Me.TxtMP_DaysPerMonth, Me.BackColorAll, Me.FrmCek, Me.UserCode, Me.ObjCD, _
                        Me.TxtMP_Primary, Me.ObjSettingUser, Me.dtpMP_RefDate, i.vPrimaryProduct)
                        Me.dgvMP.ClearSelection()
                        rowDGV.Selected = True
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMP_QtySOBSpare_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TxtMP_QtySOBSpare.MouseClick
        Dim i As New FrmTrans_MaterialPlanningDashboard_ExplainSOB
        Dim rowDGV As DataGridViewRow
        Dim vRowDT As DseTransaksiMaterialPlanning2.tbltrans_material_planning_item1Row

        Try
            If Me.dgvMP.SelectedCells.Count = 0 Then
                Exit Sub
            Else
                rowDGV = Me.dgvMP.Rows(Me.dgvMP.SelectedCells(0).RowIndex)

                If CDec(Me.TxtMP_QtySOB.Text) > 0 Then
                    For Each vRowDT In Me.vDT.Rows
                        If vRowDT.primary_code_no = rowDGV.Cells(Me.ColMP_PrimaryCodeNo.Index).Value Then
                            i.vPrimaryProduct = rowDGV.Cells(Me.ColMP_PrimaryCodeNo.Index).Value
                            i.vDT_ItemSOB = Me.vDT_ItemSOB
                            i.vDT_ItemSOBDelivery = Me.vDT_ItemSOBDelivery
                            i.BackColorAll = Me.BackColorAll
                            i.CodeNo = rowDGV.Cells(Me.ColMP_CodeNo.Index).Value
                            i.UnitCode = rowDGV.Cells(Me.ColMP_Unit.Index).Value
                            i.UserCode = Me.UserCode
                            i.rowProduct = vRowDT
                            i.ReadOnly1 = Me.btnMP_Find.Enabled
                            i.ShowDialog()

                            If i.AdaPerubahan = True Then
                                Me.ObjCls.HitungMaterialPlanning(Me.vDT, Me.vDT_ItemGIT, Me.vDT_ItemDO, Me.vDT_ItemPOB, Me.vDT_ItemSOB, Me.vDT_ItemPrOB, Me.vDT_ItemSettingLocation, _
                                Me.vDT_ItemQty, Me.vDT_ItemMOB, Me.vDT_ItemRMReq, Me.ObjSetting, Me.TxtMP_DaysPerMonth, Me.BackColorAll, Me.FrmCek, Me.UserCode, Me.ObjCD, _
                                Me.TxtMP_Primary, Me.ObjSettingUser, Me.dtpMP_RefDate, i.vPrimaryProduct)
                                Me.dgvMP.ClearSelection()
                                rowDGV.Selected = True
                            End If

                            Exit Sub
                        End If
                    Next

                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
