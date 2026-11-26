Option Explicit On
Option Strict Off

Imports Core_Module
Partial Public Class FrmTrans_SalesDashboard
    Dim ObjCd As New ClsCekData
    Dim ObjProduct As New clsMasterProduct
    Dim ObjPackingListItem As New Inventory_Module.ClsTransaksiPackingListItem
    Dim ObjBookingInv As New Inventory_Module.ClsTransaksiBookingInvItem
    'Dim ObjBookingGIT As New ClsTransaksiBookingGITitem
    Dim ObjDOLossItem As New Inventory_Module.ClsTransaksiDOLossItem
    Dim ObjDOItem As New Inventory_Module.ClsTransaksiDOItem
    Dim ObjPO As New Inventory_Module.ClsTransaksiPOItem
    Dim ObjSO As New Inventory_Module.ClsTransaksiSalesOrder
    Dim ObjSO1 As New ClsTransaksiSalesOrder
    Dim ObjUser As New Core_Module.ClsUser
    Dim ObjCustomer As New Core_Module.clsMasterCustomer
    Dim ObjMP As New clsMasterProduct_PriceList
    Dim ObjCur As New Core_Module.clsMasterCurrency
    Dim vTBLItemStock As New Inventory_Module.DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemstockDataTable
    Dim vTBLItemDelivery As New Inventory_Module.DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryDataTable
    Dim vTBLItemPOB As New Inventory_Module.DseTransaksiShippingForecast.tbltrans_shipping_forecast_item_pobDataTable
    Dim ObjSF As New Inventory_Module.ClsTransaksiShippingForecast
    Dim ObjSettingUser As New Core_Module.ClsTransaksiSettingUser
    Dim Buffer As String = ""

    Dim rowDGV As DataGridViewRow
    Dim TblMasterCurrency As New DseMasterProduct_EditPriceList.tblmaster_currencyDataTable
    Dim RowMasterCurrency As DseMasterProduct_EditPriceList.tblmaster_currencyRow

    'Dim TBLBookingGIT As New DseTransaksiBookingGIT.tbltrans_bookinggit_item_aktif_viewDataTable
    Dim TBLBookingInv As New Inventory_Module.DseTransaksiBookingInv.tbltrans_bookinginv_item_view1DataTable
    Dim TBLGIT As New Inventory_Module.DseTransaksiPackingList.DataTable22DataTable
    Dim TBLPendingDO As New DseTransaksiDO.pendingdo_view0DataTable

    Dim TBLUnitPrice As New Core_Module.dseMasterProduct.DataTable3DataTable
    Public FormNo As Integer = 0

    Dim DTA As New Inventory_Module.dseLaporanStockDetailTableAdapters.stock_detail_viewTableAdapter
    Dim TBL As New Inventory_Module.dseLaporanStockDetail.stock_detail_viewDataTable
    Dim TBL_POB As New Inventory_Module.DseTransaksiPO2.tbltrans_po_item9DataTable
    Dim TBL_SOB As New Inventory_Module.DseTransaksiSO_OrderBalance.tbltrans_so_item1DataTable

    Enum EnumPOBSort As Integer
        NoSort
        Asc
        Desc
    End Enum
    Dim SortDirection As EnumPOBSort = EnumPOBSort.NoSort
    Dim SortDirection_Date As EnumPOBSort = EnumPOBSort.NoSort

    'aul
    Dim UnAuthPO As Inventory_Module.ClsTransaksiPOItem.IncludeUnAuthPO = Inventory_Module.ClsTransaksiPOItem.IncludeUnAuthPO.ExcludeUnAuthPO
    '///

    Private Sub txtCodeNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCodeNo.GotFocus
        Try
            If Me.txtCodeNo.ReadOnly = False And Me.txtCodeNo.Enabled = True Then
                Me.txtCodeNo.BackColor = Me.BackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCodeNo.KeyDown
        Dim i As New frmCariMasterProduct
        Dim Jml As Long = 0

        Try
            If e.KeyValue = 13 Then
                i.BackColorAll = Me.BackColorAll
                Jml = Me.ObjProduct.GetJml(Me.txtCodeNo.Text)
                If Jml = 1 Then
                    Me.txtPrimaryCodeNo.Text = Me.ObjProduct.GetPrimary(Me.txtCodeNo.Text)
                    Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.txtPrimaryCodeNo.Text)
                    Me.SetGBO()
                ElseIf Jml > 1 Or Jml = 0 Then
                    'ada 2 product dengan kode yang sama atau tidak ada product yang dimaksud
                    If Me.txtCodeNo.Text <> "" Then
                        i.txtCari.Text = Me.txtCodeNo.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        Me.txtPrimaryCodeNo.Text = i.PrimaryProduct
                        Me.txtCodeNo.Text = Me.ObjProduct.GetProductCode(Me.txtPrimaryCodeNo.Text)
                        Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.txtPrimaryCodeNo.Text)
                        Me.SetGBO()
                    Else
                        Me.txtCodeNo.Clear()
                        Me.txtPrimaryCodeNo.Text = "-1"
                        Me.KosongkanItem()
                        Me.txtCodeNo.Focus()
                        Me.gboDetail.Text = ""
                        Me.DataGridView1.BringToFront()
                        Me.Stock_detail_viewDataGridView.Visible = False
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub KosongkanItem()
        Try
            Me.ObjCd.KosongkanControl(Me.Controls, True, False, False)
            Me.txtBookingGIT.Text = "0"
            Me.txtBookingInv.Text = "0"
            Me.txtFreeStock.Text = "0"
            Me.txtGIT.Text = "0"
            Me.txtQtyOnHand.Text = "0"
            Me.TxtPOB.Text = "0"
            Me.TxtSOB.Text = "0"
            Me.txtPendingDO.Text = "0"
            Me.txtAvailable.Text = "0"

            Me.txtPrimaryCodeNo.Text = "-1"
            Me.gboDetail.Text = ""
            Me.Stock_detail_viewDataGridView.Visible = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetGBO()
        Dim mDecimal As Integer = 0

        Try
            mDecimal = Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text)
            'set unit
            Me.lblUnit1.Text = Me.ObjProduct.GetUnitCode(Me.txtPrimaryCodeNo.Text)
            Me.lblUnit2.Text = Me.lblUnit1.Text
            Me.lblUnit3.Text = Me.lblUnit1.Text
            Me.lblUnit4.Text = Me.lblUnit1.Text
            Me.lblUnit5.Text = Me.lblUnit1.Text
            Me.LblUnit6.Text = Me.lblUnit1.Text
            Me.LblUnit7.Text = Me.lblUnit1.Text
            Me.lblUnitPendingDO.Text = Me.lblUnit1.Text
            Me.lblUnitAvailable.Text = Me.lblUnit1.Text

            'isikan data stock
            Me.txtGIT.Text = Me.ObjCd.DigitGrouping(Me.ObjPackingListItem.GetJmlAktif(Me.txtPrimaryCodeNo.Text), mDecimal)
            Me.txtQtyOnHand.Text = Me.ObjCd.DigitGrouping(Me.ObjProduct.GetJmlByPrimary(Me.txtPrimaryCodeNo.Text), mDecimal)
            Me.txtBookingInv.Text = Me.ObjCd.DigitGrouping(Me.ObjBookingInv.GetJml(Me.txtPrimaryCodeNo.Text), mDecimal)
            Me.txtBookingGIT.Text = Me.ObjCd.DigitGrouping(0, mDecimal)
            Me.txtFreeStock.Text = Me.ObjCd.DigitGrouping(CDec(Me.txtQtyOnHand.Text) - CDec(Me.txtBookingInv.Text), mDecimal)
            Me.txtPendingDO.Text = Me.ObjCd.DigitGrouping(Me.ObjDOItem.GetJmlPendingDO(Me.txtPrimaryCodeNo.Text) + _
            Me.ObjDOLossItem.GetJmlPendingDO(Me.txtPrimaryCodeNo.Text), mDecimal)
            Me.txtAvailable.Text = Me.ObjCd.DigitGrouping(CDec(Me.txtFreeStock.Text) - CDec(Me.txtPendingDO.Text), mDecimal)
            Me.TxtPOB.Text = Me.ObjCd.DigitGrouping(Me.ObjPO.GetOrderBalance(Me.txtPrimaryCodeNo.Text, Inventory_Module.ClsTransaksiPOItem.IncludeUnAuthPO.ExcludeUnAuthPO), mDecimal)
            Me.TxtSOB.Text = Me.ObjCd.DigitGrouping(Me.ObjSO.GetSalesOrderBalance(Me.txtPrimaryCodeNo.Text), mDecimal)

            'kosongkan gbo
            Me.gboDetail.Text = ""
            Me.Stock_detail_viewDataGridView.Visible = False
            Me.DataGridView1.BringToFront()
            Me.IsiUnitPrice()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCodeNo.Leave
        Dim i As New frmCariMasterProduct

        Try
            i.BackColorAll = Me.BackColorAll
            If Me.txtCodeNo.Text <> "" Then
                If Me.ObjProduct.GetJml(Me.txtCodeNo.Text) = 1 Then
                    Me.txtPrimaryCodeNo.Text = Me.ObjProduct.GetPrimary(Me.txtCodeNo.Text)
                    Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.txtPrimaryCodeNo.Text)
                    Me.SetGBO()
                Else
                    If Me.txtCodeNo.Text <> "" Then
                        i.txtCari.Text = Me.txtCodeNo.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        Me.txtPrimaryCodeNo.Text = i.PrimaryProduct
                        Me.txtCodeNo.Text = Me.ObjProduct.GetProductCode(Me.txtPrimaryCodeNo.Text)
                        Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.txtPrimaryCodeNo.Text)
                        Me.SetGBO()
                    Else
                        Me.txtCodeNo.Clear()
                        Me.txtPrimaryCodeNo.Text = "-1"
                        Me.KosongkanItem()
                        Me.txtCodeNo.Focus()
                        Me.gboDetail.Text = ""
                        Me.DataGridView1.BringToFront()
                        Me.Stock_detail_viewDataGridView.Visible = False
                    End If
                End If
            Else
                Me.txtPrimaryCodeNo.Text = "-1"
                Me.txtDesc.Clear()
                Me.KosongkanItem()
                Me.gboDetail.Text = ""
                Me.DataGridView1.BringToFront()
                Me.Stock_detail_viewDataGridView.Visible = False
                Me.IsiUnitPrice()
            End If

            Me.TxtSF_PrimaryCodeNo.Text = Me.txtPrimaryCodeNo.Text
            Me.TxtSF_ProductCode.Text = Me.txtCodeNo.Text
            Me.TxtSF_ProductDesc.Text = Me.txtDesc.Text
            Me.TxtMA_CodeDesc.Text = Me.txtDesc.Text
            Me.TxtMA_CodeNo.Text = Me.txtCodeNo.Text
            Me.TxtMA_PrimaryCodeNo.Text = Me.txtPrimaryCodeNo.Text
            'If Me.TabControl2.SelectedIndex = 1 Then
            '    Me.btnMA_FilterSalesHistory.PerformClick()
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindProduct_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindProduct.Click
        Dim i As New frmCariMasterProduct

        Try
            i.BackColorAll = Me.BackColorAll
            i.cboCriteria.SelectedIndex = 0
            If Me.txtCodeNo.Text <> "" Then
                i.txtCari.Text = Me.txtCodeNo.Text
            End If

            i.ShowDialog()
            If i.PrimaryProduct > 0 Then
                Me.txtPrimaryCodeNo.Text = i.PrimaryProduct
                Me.txtCodeNo.Text = Me.ObjProduct.GetProductCode(Me.txtPrimaryCodeNo.Text)
                Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.txtPrimaryCodeNo.Text)
                Me.SetGBO()
            Else
                Me.txtCodeNo.Clear()
                Me.txtPrimaryCodeNo.Text = "-1"
                Me.KosongkanItem()
                Me.txtCodeNo.Focus()
                Me.gboDetail.Text = ""
                Me.DataGridView1.BringToFront()
                Me.Stock_detail_viewDataGridView.Visible = False
            End If

            Me.TxtSF_PrimaryCodeNo.Text = Me.txtPrimaryCodeNo.Text
            Me.TxtSF_ProductCode.Text = Me.txtCodeNo.Text
            Me.TxtSF_ProductDesc.Text = Me.txtDesc.Text
            Me.TxtMA_CodeDesc.Text = Me.txtDesc.Text
            Me.TxtMA_CodeNo.Text = Me.txtCodeNo.Text
            Me.TxtMA_PrimaryCodeNo.Text = Me.txtPrimaryCodeNo.Text

            If Me.TabControl2.SelectedIndex = 1 Then
                Me.btnMA_FilterSalesHistory.PerformClick()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub llbColumnSettings_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbColumnSettings.LinkClicked
        Dim ij As New Inventory_Module.FrmTrans_MaterialPlanningDashboard_SF_SetVisible

        Try
            'tampilkan form untuk setting data
            ij.vUserName = Me.UserCode
            If ij.ShowDialog = Windows.Forms.DialogResult.OK Then
                'update kolom
                Me.UpdateKolomSF()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Sub UpdateKolomSF()
        Dim Dummy As String = ""

        Try
            Dummy = Me.ObjSettingUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.ShippingForecast, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.ShippingForecast_ShowSONumber, _
            Me.UserCode)
            If Dummy = "" Then
                Me.ColSF_SONumber.Visible = False
            Else
                If Dummy = 1 Then
                    Me.ColSF_SONumber.Visible = True
                Else
                    Me.ColSF_SONumber.Visible = False
                End If
            End If

            Dummy = Me.ObjSettingUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.ShippingForecast, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.ShippingForecast_ShowProductDesc, _
            Me.UserCode)
            If Dummy = "" Then
                Me.ColSF_Desc.Visible = False
            Else
                If Dummy = 1 Then
                    Me.ColSF_Desc.Visible = True
                Else
                    Me.ColSF_Desc.Visible = False
                End If
            End If

            Dummy = Me.ObjSettingUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.ShippingForecast, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.ShippingForecast_ShowCustPO, _
            Me.UserCode)
            If Dummy = "" Then
                Me.ColSF_CustomerPONumber.Visible = False
            Else
                If Dummy = 1 Then
                    Me.ColSF_CustomerPONumber.Visible = True
                Else
                    Me.ColSF_CustomerPONumber.Visible = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetFilterSF() As String
        Dim Filter As String = ""

        Try
            Me.TxtSF_ProductCode.Text = Me.TxtSF_ProductCode.Text.Trim
            If Me.TxtSF_ProductCode.Text <> "" Then
                Filter = "[Product Code] = '" + Me.TxtSF_ProductCode.Text + "'"
            End If
            If Me.TxtSF_PrimaryCust.Text <> "-1" And Me.TxtSF_PrimaryCust.Text <> "-1" Then
                If Filter = "" Then
                    Filter = "[primary cust] = " + Me.TxtSF_PrimaryCust.Text
                Else
                    Filter = Filter + " AND [primary cust] = " + Me.TxtSF_PrimaryCust.Text
                End If
            End If
            If Me.TxtSF_UserCode.Text <> "" Then
                If Filter = "" Then
                    Filter = "[so created by] = '" + Me.TxtSF_UserCode.Text + "'"
                Else
                    Filter = Filter + " AND [so created by] = '" + Me.TxtSF_UserCode.Text + "'"
                End If
            End If
            If Me.chkSF_LateOnly.Checked = True Then
                If Filter = "" Then
                    Filter = "[estimate delivery] > [promise to deliver]"
                Else
                    Filter = Filter + " AND [estimate delivery] > [promise to deliver]"
                End If
            End If
            If Me.TxtSF_PrimarySO.Text <> "" And Me.TxtSF_PrimarySO.Text <> "-1" Then
                If Filter = "" Then
                    Filter = "[primary so] = " + Me.TxtSF_PrimarySO.Text + ""
                Else
                    Filter = Filter + " AND [primary so] = " + Me.TxtSF_PrimarySO.Text + ""
                End If
            End If

            'rnew
            If Me.chkSODate.Checked = True Then
                If Filter = "" Then
                    Filter = "[so date] >= #" + Format(Me.DtpFilterSF_From.Value, "yyyy-MM-dd") + "# AND [so date] <= #" + Format(Me.DtpFilterSF_To.Value, "yyyy-MM-dd") + "#"
                Else
                    Filter = Filter + " AND [so date] >= #" + Format(Me.DtpFilterSF_From.Value, "yyyy-MM-dd") + "# AND [so date] <= #" + Format(Me.DtpFilterSF_To.Value, "yyyy-MM-dd") + "#"
                End If
            End If

            Return Filter
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Sub LoadDGV_SF()
        Dim Filter As String = ""
        Dim DV As Data.DataRowView
        Dim rowDGV As DataGridViewRow
        Dim Masuk As Boolean = True
        Dim rowItemSOB As Inventory_Module.DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryRow
        Dim Lanjut As Boolean = False

        Try
            Filter = Me.GetFilterSF

            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            'secara default disort berdasarkan: estimate delivery, customer code, product code, so date, so number
            Me.vTBLItemDelivery.DefaultView.Sort = "[promise to deliver] ASC, [customer name] ASC, [Product Code] ASC, [so date rec] ASC, [so number rec] ASC"
            Me.vTBLItemDelivery.DefaultView.RowFilter = Filter

            Me.dgvSF.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvSF.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Me.dgvSF.Rows.Clear()

            Me.ColSF_SODate.DefaultCellStyle.Format = "dd MMM yyyy"

            For Each DV In Me.vTBLItemDelivery.DefaultView
                Lanjut = True
                rowItemSOB = DV.Row
                If Me.chkSF_LateOnly.Checked = True Then
                    If rowItemSOB.promise_to_deliver = #1/1/2200# Then
                        Lanjut = False
                    Else
                        If rowItemSOB.promise_to_deliver.Date >= rowItemSOB.estimate_delivery.Date Then
                            If rowItemSOB.qty_available < rowItemSOB.qty_sob Then
                                Lanjut = True
                            Else
                                Lanjut = False
                            End If
                        Else
                            Lanjut = True
                        End If
                    End If
                Else
                    Lanjut = True
                End If
                If Lanjut = True Then
                    If Me.cboAvailability.SelectedIndex = 1 Then
                        'hanya yg available : qty available > qty sob dan tanggal available > dtp available
                        If rowItemSOB.qty_available >= rowItemSOB.qty_sob Then
                            If Me.dtpAvailable.Checked = False Then
                                Lanjut = True
                            Else
                                If rowItemSOB.estimate_delivery.Date <= Me.dtpAvailable.Value.Date Then
                                    Lanjut = True
                                Else
                                    Lanjut = False
                                End If
                            End If
                        Else
                            Lanjut = False
                        End If
                    ElseIf Me.cboAvailability.SelectedIndex = 2 Then
                        If rowItemSOB.qty_available < rowItemSOB.qty_sob Then
                            Lanjut = True
                        Else
                            Lanjut = False
                        End If
                    End If

                    'If Me.chkSF_AvailableOnly.Checked = True Then
                    '    If rowItemSOB.qty_available > 0 Then
                    '        Lanjut = True
                    '    Else
                    '        Lanjut = False
                    '    End If
                    'Else
                    '    Lanjut = True
                    'End If
                End If

                If Lanjut = True Then
                    If Me.TxtSF_PrimarySO.Text <> "" And Me.TxtSF_PrimarySO.Text <> "-1" Then
                        If rowItemSOB.primary_so <> Me.TxtSF_PrimarySO.Text Then
                            Lanjut = False
                        End If
                    End If
                End If

                If Lanjut = True Then
                    Me.dgvSF.Rows.Add()
                    rowDGV = Me.dgvSF.Rows(Me.dgvSF.Rows.Count - 1)
                    With rowDGV
                        .Cells(Me.ColSF_Primary.Index).Value = rowItemSOB.primary
                        .Cells(Me.ColSF_CodeNo.Index).Value = rowItemSOB.Product_Code
                        If rowItemSOB.create_do = 1 Then
                            .Cells(Me.ColSF_CreateDO.Index).Value = Boolean.TrueString
                        Else
                            .Cells(Me.ColSF_CreateDO.Index).Value = Boolean.FalseString
                        End If
                        .Cells(Me.ColSF_Customer.Index).Value = rowItemSOB.customer_name
                        .Cells(Me.ColSF_CustomerPONumber.Index).Value = rowItemSOB.Cust_PO_No
                        .Cells(Me.ColSF_EstimateDelivery.Index).Value = (rowItemSOB.estimate_delivery.Date - #1/1/1950#).Days
                        .Cells(Me.ColSF_No.Index).Value = Me.dgvSF.Rows.Count
                        .Cells(Me.ColSF_PrimaryCodeNo.Index).Value = rowItemSOB.primary_code_no
                        .Cells(Me.ColSF_PrimarySO.Index).Value = rowItemSOB.primary_so
                        .Cells(Me.ColSF_PrimaryUnit.Index).Value = rowItemSOB.primary_unit
                        .Cells(Me.ColSF_PromiseToDeliver.Index).Value = (rowItemSOB.promise_to_deliver.Date - #1/1/1950#).Days
                        .Cells(Me.ColSF_QtyAvailable.Index).Value = ObjCD.DigitGrouping(rowItemSOB.qty_available, rowItemSOB._Decimal)
                        .Cells(Me.ColSF_QtySOB.Index).Value = ObjCD.DigitGrouping(rowItemSOB.qty_sob, rowItemSOB._Decimal)
                        .Cells(Me.ColSF_SODate.Index).Value = rowItemSOB.so_date
                        .Cells(Me.ColSF_Unit.Index).Value = rowItemSOB.Unit_Code
                        .Cells(Me.ColSF_Desc.Index).Value = rowItemSOB.Description
                        If rowItemSOB.estimate_delivery.Date = #1/1/1950# Then
                            .Cells(Me.ColSF_EstimateDelivery_Str.Index).Value = "Today"
                        ElseIf rowItemSOB.estimate_delivery.Date = #1/1/2200# Then
                            .Cells(Me.ColSF_EstimateDelivery_Str.Index).Value = "n/a"
                        Else
                            .Cells(Me.ColSF_EstimateDelivery_Str.Index).Value = Format(rowItemSOB.estimate_delivery.Date, "dd MMM yyyy")
                        End If
                        If rowItemSOB.promise_to_deliver.Date = #1/1/1950# Then
                            .Cells(Me.ColSF_PromiseToDeliver_Str.Index).Value = "Today"
                        ElseIf rowItemSOB.promise_to_deliver = #1/1/2200# Then
                            .Cells(Me.ColSF_PromiseToDeliver_Str.Index).Value = "n/a"
                        Else
                            .Cells(Me.ColSF_PromiseToDeliver_Str.Index).Value = Format(rowItemSOB.promise_to_deliver.Date, "dd MMM yyyy")
                        End If
                        .Cells(Me.ColSF_SONumber.Index).Value = rowItemSOB.so_number
                        .Cells(Me.ColSF_PrimaryCust.Index).Value = rowItemSOB.primary_cust

                        If rowItemSOB.qty_available > 0 Then
                            If rowItemSOB.qty_sob <= rowItemSOB.qty_available Then
                                If rowItemSOB.qty_ready_stock >= rowItemSOB.qty_available Then
                                    .Cells(Me.ColSF_DeliveryStatus.Index).Value = "Ready"
                                ElseIf rowItemSOB.qty_ready_stock > 0 Then
                                    .Cells(Me.ColSF_DeliveryStatus.Index).Value = "Partial"
                                Else
                                    .Cells(Me.ColSF_DeliveryStatus.Index).Value = "Waiting"
                                End If
                            Else
                                If rowItemSOB.qty_ready_stock >= rowItemSOB.qty_available Then
                                    .Cells(Me.ColSF_DeliveryStatus.Index).Value = "Partial"
                                ElseIf rowItemSOB.qty_ready_stock > 0 Then
                                    .Cells(Me.ColSF_DeliveryStatus.Index).Value = "Partial"
                                Else
                                    .Cells(Me.ColSF_DeliveryStatus.Index).Value = "Waiting"
                                End If
                            End If
                        Else
                            .Cells(Me.ColSF_DeliveryStatus.Index).Value = "Waiting"
                        End If
                        .Cells(Me.ColSF_QtyReadyStock.Index).Value = rowItemSOB.qty_ready_stock
                        .Cells(Me.ColSF_Identifier.Index).Value = rowItemSOB.identifier
                    End With
                    If rowItemSOB.promise_to_deliver.Date < rowItemSOB.estimate_delivery.Date Or _
                    rowItemSOB.qty_available < rowItemSOB.qty_sob Then
                        rowDGV.DefaultCellStyle.ForeColor = Color.Red
                    End If
                End If
            Next

            Me.dgvSF.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkSF_LateOnly_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkSF_LateOnly.CheckedChanged
        Try
            Me.LoadDGV_SF()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCodeNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCodeNo.LostFocus
        Try
            If Me.txtCodeNo.ReadOnly = False And Me.txtCodeNo.Enabled = True Then
                Me.txtCodeNo.BackColor = Me.txtPrimaryCodeNo.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmMasterKondisiStock_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim objCustom As New CustomSMESS.ClsCustomOtorisasi
        Try
            Me.gboOpportunity.Visible = True
            Me.gboOpportunity.Dock = DockStyle.Fill
            Me.TabControl1.TabPages.RemoveAt(0)

            Me.Stock_detail_viewDataGridView.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            If Me.txtCodeNo.Text <> "" Then
                If Me.ObjProduct.ApakahSudahAda(Me.txtCodeNo.Text) = True Then
                    Me.txtPrimaryCodeNo.Text = Me.ObjProduct.GetPrimary(Me.txtCodeNo.Text)
                    Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.txtPrimaryCodeNo.Text)
                    Me.SetGBO()
                End If
            End If

            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Core_Module.GetSMESSBackColorAll
            Me.dgvItemPendingDO.AlternatingRowsDefaultCellStyle.BackColor = Core_Module.GetSMESSBackColorAll
            Me.dgvException.AlternatingRowsDefaultCellStyle.BackColor = Core_Module.GetSMESSBackColorAll
            Me.dgvSalesBalance.AlternatingRowsDefaultCellStyle.BackColor = Core_Module.GetSMESSBackColorAll
            Me.dgvItemPOB.AlternatingRowsDefaultCellStyle.BackColor = Core_Module.GetSMESSBackColorAll
            Me.dgvItemSOB.AlternatingRowsDefaultCellStyle.BackColor = Core_Module.GetSMESSBackColorAll
            Me.dgvSF.AlternatingRowsDefaultCellStyle.BackColor = Core_Module.GetSMESSBackColorAll
            Me.dgvSH.AlternatingRowsDefaultCellStyle.BackColor = Core_Module.GetSMESSBackColorAll
            Me.DGVHarga.AlternatingRowsDefaultCellStyle.BackColor = Core_Module.GetSMESSBackColorAll

            Me.ColSH_QtyDefault.Visible = False
            Me.ColSH_SOB_Default.Visible = False
            Me.ColSH_Unit_Default.Visible = False

            'atur manajemen user
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_SalesDashboard_Override_User") = True Then
                Me.AllowToSeeAllUser = True
            Else
                Me.AllowToSeeAllUser = False
                Me.TxtAR_UserCode.Text = Me.UserCode
                Me.TxtAR_UserCodeName.Text = Me.ObjUser.GetNama(Me.UserCode)
                Me.TxtUP_UserCode.Text = Me.UserCode
                Me.TxtUP_UserName.Text = Me.TxtAR_UserCodeName.Text
                Me.TxtMA_UserCode.Text = Me.UserCode
                Me.TxtMAX_UserName.Text = Me.TxtAR_UserCodeName.Text
                Me.TxtSF_UserCode.Text = Me.UserCode
                Me.TxtSF_UserName.Text = Me.TxtAR_UserCodeName.Text
            End If
            If Me.AllowToSeeAllUser = False Then
                Me.btnAR_FindUserCode.Enabled = False
                Me.TxtAR_UserCode.Enabled = False
                Me.btnUP_FindUser.Enabled = False
                Me.TxtUP_UserCode.Enabled = False
                Me.btnMA_FindUser.Enabled = False
                Me.TxtMA_UserCode.Enabled = False
                Me.btnSF_FindUser.Enabled = False
                Me.TxtSF_UserCode.Enabled = False
            End If

            Me.dtpSH_From.Value = Today - TimeSpan.FromDays(31)
            Me.dtpSH_To.Value = Today
            Me.cboAvailability.SelectedIndex = 0

            'Update SMESS 2026
            If objCustom.ApakahPunyaHakDeleteUP(Me.UserCode) = False Then
                Me.BtnUP_Delete.Visible = False
                Me.llbUP_CheckedAllDelete.Visible = False
                Me.llbUP_UncheckedAllDelete.Visible = False
                Me.dgvUP.Columns(Me.ColUP_ChkDelete.Index).Visible = False
            End If

            Me.dgvUP.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Me.dgvDOItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpAvailable_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpAvailable.Enter
        Try
            Me.Buffer = Format(Me.dtpAvailable.Value.Date, "dd MMM yyyy") + " - " + Me.dtpAvailable.Checked.ToString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dtpAvailable_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpAvailable.Leave
        Try
            If Me.Buffer <> Format(Me.dtpAvailable.Value.Date, "dd MMM yyyy") + " - " + Me.dtpAvailable.Checked.ToString Then
                Me.BackgroundWorker1.RunWorkerAsync()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker5_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker5_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            Me.LoadDGV_SF()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboAvailability_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboAvailability.SelectedIndexChanged
        Try
            If Me.cboAvailability.SelectedIndex = 1 Then
                Me.dtpAvailable.Enabled = True
            Else
                Me.dtpAvailable.Enabled = False
            End If
            Me.LoadDGV_SF()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbQtyOnHand_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbQtyOnHand.LinkClicked
        Try
            Me.ShowQtyOnHand()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ShowQtyOnHand()
        Dim rowDGV As DataGridViewRow
        Dim SecNo As Boolean = True
        Dim LotNo As Boolean = False
        Dim Location As Boolean = True
        Dim ExpiryDate As Boolean = True
        Dim col As DataGridViewColumn
        Dim rowStock As dseLaporanStockDetail.stock_detail_viewRow
        Dim DefPrec As Integer = 0
        Dim KontrolLokasi, KontrolBatch, KontrolExpiry As Boolean

        Try
            Me.Stock_detail_viewDataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            'TODO: This line of code loads data into the 'DseLaporanStockDetail.stock_detail_view' table. You can move, or remove it, as needed.
            If Me.txtCodeNo.Text <> "" Then
                DTA.FillBy_PrimaryProduct(TBL, Me.txtPrimaryCodeNo.Text)
            Else
                MsgBox("Please specify code no.", MsgBoxStyle.Information, "Information")
                Me.txtCodeNo.Focus()
                TBL.Rows.Clear()
                Exit Sub
            End If
            KontrolBatch = Me.ObjProduct.ApakahBatchNoControlled(Me.txtPrimaryCodeNo.Text)
            KontrolExpiry = Me.ObjProduct.ApakahExpiryDateControlled(Me.txtPrimaryCodeNo.Text)
            KontrolLokasi = Me.ObjProduct.ApakahLocationControlled(Me.txtPrimaryCodeNo.Text)

            If KontrolExpiry = True Then
                TBL.DefaultView.Sort = "[Expiry Date], [Lot No], Location"
            Else
                TBL.DefaultView.Sort = "[Lot No], [Expiry Date], Location"
            End If

            Me.Stock_detail_viewDataGridView.DataSource = TBL.DefaultView
            'Me.Stock_detail_viewDataGridView.Columns(TBL.QtyColumn.ColumnName).DefaultCellStyle.Format = "N" + DefPrec.ToString

            For Each rowDGV In Me.Stock_detail_viewDataGridView.Rows
                If CStr(rowDGV.Cells(TBL.Lot_NoColumn.ColumnName).Value) <> "" Then
                    LotNo = True
                End If
                If CDate(rowDGV.Cells(TBL.Expiry_DateColumn.ColumnName).Value) <> #1/1/1950# Then
                    ExpiryDate = False
                End If
                If CStr(rowDGV.Cells(TBL.LocationColumn.ColumnName).Value) <> "No Location" Then
                    Location = False
                End If
            Next

            If KontrolBatch = True Then
                Me.Stock_detail_viewDataGridView.Columns(TBL.Lot_NoColumn.ColumnName).Visible = True
            Else
                Me.Stock_detail_viewDataGridView.Columns(TBL.Lot_NoColumn.ColumnName).Visible = False
            End If
            If KontrolExpiry = True Then
                Me.Stock_detail_viewDataGridView.Columns(TBL.Expiry_DateColumn.ColumnName).Visible = True
            Else
                Me.Stock_detail_viewDataGridView.Columns(TBL.Expiry_DateColumn.ColumnName).Visible = False
            End If
            If KontrolLokasi = True Then
                Me.Stock_detail_viewDataGridView.Columns(TBL.LocationColumn.ColumnName).Visible = True
                Me.Stock_detail_viewDataGridView.Columns(TBL.WarehouseColumn.ColumnName).Visible = True
            Else
                Me.Stock_detail_viewDataGridView.Columns(TBL.LocationColumn.ColumnName).Visible = False
                Me.Stock_detail_viewDataGridView.Columns(TBL.WarehouseColumn.ColumnName).Visible = False
            End If


            'format dgv
            Me.Stock_detail_viewDataGridView.Columns(TBL.Expiry_DateColumn.ColumnName).DefaultCellStyle.Format = "dd MMM yyyy"
            Me.Stock_detail_viewDataGridView.Columns(0).DefaultCellStyle.Format = "N" + Me.ObjProduct.GetAngkaDecimal(Me.txtPrimaryCodeNo.Text).ToString
            Me.Stock_detail_viewDataGridView.Columns(TBL.Primary_Storage_AreaColumn.ColumnName).Visible = False
            Me.Stock_detail_viewDataGridView.Columns(TBL.Primary_UnitColumn.ColumnName).Visible = False
            Me.Stock_detail_viewDataGridView.Columns(TBL.Primary_WarehouseColumn.ColumnName).Visible = False
            Me.Stock_detail_viewDataGridView.Columns(TBL.DecimalColumn.ColumnName).Visible = False
            If KontrolExpiry = False Then
                Me.Stock_detail_viewDataGridView.Columns(TBL.Expiry_DateColumn.ColumnName).Visible = False
            End If
            For Each col In Me.Stock_detail_viewDataGridView.Columns
                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            Next
            Me.Stock_detail_viewDataGridView.Columns(TBL.QtyColumn.ColumnName).DefaultCellStyle.Alignment = _
            DataGridViewContentAlignment.MiddleRight

            Me.Stock_detail_viewDataGridView.Visible = True
            Me.Stock_detail_viewDataGridView.BringToFront()
            Me.gboDetail.Text = "Qty On Hand"

            Me.Stock_detail_viewDataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Private Sub llbBookingGIT_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbBookingGIT.LinkClicked
    '    Dim col As DataGridViewColumn

    '    Try
    '        Me.ObjBookingGIT.IsiDgvCheckStock(Me.dgvItem, Me.TBLBookingGIT, Me.txtPrimaryCodeNo.Text)
    '        Me.dgvItem.BringToFront()
    '        Me.Stock_detail_viewDataGridView.Visible = False
    '        Me.gboDetail.Text = "Booking GIT"

    '        'format tampilan
    '        Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
    '        For Each col In Me.dgvItem.Columns
    '            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
    '        Next
    '        Me.dgvItem.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
    '        Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Private Sub llbBookingStockInv_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbBookingStockInv.LinkClicked
        Dim row As DataGridViewRow
        Dim col As DataGridViewColumn

        Try
            Me.TabControl2.SelectedIndex = 0
            Me.dgvItem.BringToFront()
            Me.Stock_detail_viewDataGridView.Visible = False
            Me.gboDetail.Text = "Booking Inventory"
            If Me.txtCodeNo.Text = "" Then
                MsgBox("Please specify code no.", MsgBoxStyle.Information, "Information")
                Me.txtCodeNo.Focus()
                TBL.Rows.Clear()
                Exit Sub
            Else
                Me.ObjBookingInv.IsiDgv1(Me.dgvItem, Me.TBLBookingInv, Me.txtPrimaryCodeNo.Text)
            End If

            'format tampilan
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For Each col In Me.dgvItem.Columns
                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            Next
            Me.dgvItem.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            'Me.dgvItem.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            For Each row In Me.dgvItem.Rows
                If row.Cells(9).Value <> -1 Then
                    row.Cells(9).Value = Me.ObjCd.DigitGrouping(row.Cells(9).Value)
                End If
            Next
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbGIT_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbGIT.LinkClicked
        Dim col As DataGridViewColumn

        Try
            Me.TabControl2.SelectedIndex = 0
            Me.dgvItem.BringToFront()
            Me.Stock_detail_viewDataGridView.Visible = False
            Me.gboDetail.Text = "Goods In Transit"

            If Me.txtCodeNo.Text = "" Then
                MsgBox("Please specify code no.", MsgBoxStyle.Information, "Information")
                Me.txtCodeNo.Focus()
                TBL.Rows.Clear()
                Exit Sub
            Else
                Me.ObjPackingListItem.IsiDgvCheckStock(Me.dgvItem, Me.TBLGIT, Me.txtPrimaryCodeNo.Text)
            End If

            'format tampilan
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For Each col In Me.dgvItem.Columns
                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            Next
            Me.dgvItem.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'update ory
    Private Sub lblPendingDO_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lblPendingDO.LinkClicked
        Dim DTA As New DseTransaksiDOTableAdapters.pendingdo_view0TableAdapter
        Dim DTA1 As New DseTransaksiDOLossTableAdapters.tbltrans_doloss_pending_view0TableAdapter
        Dim TBL As New DseTransaksiDOLoss.tbltrans_doloss_pending_view0DataTable
        Dim rowPending As DseTransaksiDO.pendingdo_view0Row
        Dim rowDOLoss As DseTransaksiDOLoss.tbltrans_doloss_pending_view0Row

        Try
            Me.dgvItemPendingDO.DataSource = Nothing
            Me.dgvItemPendingDO.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            Me.gboDetail.Text = "Pending DO"
            Me.dgvItemPendingDO.BringToFront()
            If Me.txtCodeNo.Text = "" Then
                MsgBox("Please specify code no.", MsgBoxStyle.Information, "Information")
                Me.txtCodeNo.Focus()
                TBL.Rows.Clear()
                Exit Sub
            End If

            If Me.txtPrimaryCodeNo.Text = "" Or Me.txtPrimaryCodeNo.Text = "-1" Then
                DTA.Fill(Me.TBLPendingDO)
                DTA1.Fill(TBL)
            Else
                DTA.FillBy_PrimaryProduct(Me.TBLPendingDO, Me.txtPrimaryCodeNo.Text)
                DTA1.FillBy_PrimaryProduct(TBL, Me.txtPrimaryCodeNo.Text)
            End If

            'tambahkan baris pending do ke dalam laporan
            For Each rowDOLoss In TBL.Rows
                rowPending = Me.TBLPendingDO.Newpendingdo_view0Row
                With rowPending
                    .DO_Number = rowDOLoss.DO_Number
                    .Product_Code = rowDOLoss.Product_Code
                    .Description = rowDOLoss.Description
                    .Sales_Order_No = rowDOLoss.Sales_Order_Number
                    .Batch_No = rowDOLoss.Batch_No
                    .Expiry_Date = rowDOLoss.Expiry_Date
                    .Total_Qty = rowDOLoss.Total_Qty
                    .Unit_Code = rowDOLoss.Unit_Code
                    .Secondary_Qty = rowDOLoss.Secondary_Qty
                    .Secondary_Unit = rowDOLoss.Secondary_Unit
                    .Remarks = rowDOLoss.Remarks
                End With
                Me.TBLPendingDO.Rows.Add(rowPending)
            Next

            For Each rowPending In Me.TBLPendingDO.Rows
                rowPending.Quantity = Me.ObjCd.DigitGrouping(rowPending.Total_Qty, rowPending._Decimal)
            Next

            ' set tanggal
            For Each rowPending In TBLPendingDO
                With rowPending
                    .Tanggal = .Tanggal_Transaksi.ToString("dd MMM yyyy")
                End With
            Next

            'format tampilan
            TBLPendingDO.DefaultView.Sort = "[Tanggal Transaksi] asc,[DO Number] asc,[Customer Code] asc,[Customer Name] asc"
            Me.dgvItemPendingDO.DataSource = Me.TBLPendingDO.DefaultView
            dgvItemPendingDO.Columns(1).Visible = False
            dgvItemPendingDO.Columns(2).HeaderText = "DO Date"
            Dim row As DataGridViewRow
            Dim AdaBatch As Boolean = False
            Dim AdaExpiry As Boolean = False
            Dim AdaSecondary As Boolean = False
            If Me.dgvItemPendingDO.Columns.Count > 0 Then
                Me.dgvItemPendingDO.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                Me.dgvItemPendingDO.Columns(13).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            End If

            For Each row In Me.dgvItemPendingDO.Rows
                If row.Cells(8).Value <> "" Then
                    AdaBatch = True
                End If
                If row.Cells(9).Value <> #1/1/1950# Then
                    AdaExpiry = True
                End If
                If row.Cells(13).Value <> -1 Then
                    AdaSecondary = True
                End If
            Next


            If AdaBatch = False Then
                Me.dgvItemPendingDO.Columns(8).Visible = False
            Else
                Me.dgvItemPendingDO.Columns(8).Visible = True
            End If
            If AdaExpiry = False Then
                Me.dgvItemPendingDO.Columns(9).Visible = False
            Else
                Me.dgvItemPendingDO.Columns(9).Visible = True
            End If
            If AdaSecondary = True Then
                Me.dgvItemPendingDO.Columns(13).Visible = True
                Me.dgvItemPendingDO.Columns(14).Visible = True
            Else
                Me.dgvItemPendingDO.Columns(13).Visible = False
                Me.dgvItemPendingDO.Columns(14).Visible = False
            End If
            Me.dgvItemPendingDO.Columns(10).Visible = False
            Me.dgvItemPendingDO.Columns(16).Visible = False
            Me.dgvItemPendingDO.Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

            Me.dgvItemPendingDO.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Private Sub IsiUnitPrice()
        Try
            If Me.txtPrimaryCodeNo.Text <> "-1" And Me.txtPrimaryCodeNo.Text <> "" Then
                Me.cboMA_Currency.Enabled = True
                Me.cboMA_Currency.DataSource = Nothing
                Me.cboMA_Currency.Items.Clear()

                Me.ObjMP.IsiTBLCurrency(Me.TblMasterCurrency)

                Me.cboMA_Currency.DataSource = Me.TblMasterCurrency.DefaultView
                Me.txtPrimaryCurrency.Text = Me.TblMasterCurrency.Rows(0).Item(0).ToString
                Me.TblMasterCurrency.DefaultView.Sort = "Primary"
                Me.cboMA_Currency.DisplayMember = "Currency Code"
            Else
                Me.TxtUnitPrice.Text = 0
                Me.cboMA_Currency.Enabled = False
                Me.cboMA_Currency.DataSource = Nothing
                Me.cboMA_Currency.Items.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbPOB_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbPOB.LinkClicked
        Try
            Me.TabControl2.SelectedIndex = 0
            Me.gboDetail.Text = "Purchase Order Balance"
            'dgv bring to front
            Me.dgvItemPOB.BringToFront()

            If Me.txtCodeNo.Text = "" Then
                MsgBox("Please specify code no.", MsgBoxStyle.Information, "Information")
                Me.txtCodeNo.Focus()
                TBL.Rows.Clear()
                Exit Sub
            Else
                'isikan dgv POB
                If Me.txtPrimaryCodeNo.Text = "" Or Me.txtPrimaryCodeNo.Text = "-1" Then
                    Me.txtPrimaryCodeNo.Text = "-1"
                End If
                Me.ObjPO.IsiDGVPOB(Me.txtPrimaryCodeNo.Text, Me.dgvItemPOB, Me.TBL_POB, Me.UnAuthPO)
                Me.SortDirection = EnumPOBSort.NoSort
                Me.SortDirection_Date = EnumPOBSort.NoSort
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemPOB_ColumnHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvItemPOB.ColumnHeaderMouseClick
        Dim IndexPOB As Integer = 7
        Dim IndexPOBStr As Integer = 11
        Dim IndexPOB_Date As Integer = 10
        Dim IndexPOB_Date_Str As Integer = 12

        Try
            Me.dgvItemPOB.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItemPOB.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            If e.ColumnIndex = IndexPOBStr Then
                If Me.SortDirection = EnumPOBSort.Desc Or Me.SortDirection = EnumPOBSort.NoSort Then
                    Me.dgvItemPOB.Sort(Me.dgvItemPOB.Columns(IndexPOB), System.ComponentModel.ListSortDirection.Ascending)
                    Me.SortDirection = EnumPOBSort.Asc
                Else
                    Me.dgvItemPOB.Sort(Me.dgvItemPOB.Columns(IndexPOB), System.ComponentModel.ListSortDirection.Descending)
                    Me.SortDirection = EnumPOBSort.Desc
                End If
                Me.SortDirection_Date = EnumPOBSort.NoSort
            ElseIf e.ColumnIndex = IndexPOB_Date_Str Then
                If Me.SortDirection_Date = EnumPOBSort.Desc Or Me.SortDirection_Date = EnumPOBSort.NoSort Then
                    Me.dgvItemPOB.Sort(Me.dgvItemPOB.Columns(IndexPOB_Date), System.ComponentModel.ListSortDirection.Ascending)
                    Me.SortDirection_Date = EnumPOBSort.Asc
                Else
                    Me.dgvItemPOB.Sort(Me.dgvItemPOB.Columns(IndexPOB_Date), System.ComponentModel.ListSortDirection.Descending)
                    Me.SortDirection_Date = EnumPOBSort.Desc
                End If
                Me.SortDirection = EnumPOBSort.NoSort
            Else
                Me.SortDirection = EnumPOBSort.NoSort
                Me.SortDirection_Date = EnumPOBSort.NoSort
            End If

            Me.dgvItemPOB.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbSOB_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbSOB.LinkClicked
        Try
            Me.TabControl2.SelectedIndex = 0
            Me.gboDetail.Text = "Sales Order Balance"
            'dgv bring to front
            Me.dgvItemSOB.BringToFront()
            Me.dgvItemSOB.Rows.Clear()

            'isikan dgv POB
            If Me.txtPrimaryCodeNo.Text = "" Or Me.txtPrimaryCodeNo.Text = "-1" Then
                Me.txtPrimaryCodeNo.Text = "-1"
                MsgBox("Please specify code no.", MsgBoxStyle.Information, "Information")
                Me.txtCodeNo.Focus()
                TBL.Rows.Clear()
                Exit Sub
            End If

            Me.ObjSO.IsiDGVSalesOrderBalance(Me.txtPrimaryCodeNo.Text, Me.TBL_SOB, Me.dgvItemSOB)
            Me.SortDirection = EnumPOBSort.NoSort
            Me.SortDirection_Date = EnumPOBSort.NoSort
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemSOB_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItemSOB.CellContentClick
        Dim IndexPOB As Integer = 9
        Dim IndexPOBStr As Integer = 11
        Dim IndexPOB_Date As Integer = 10
        Dim IndexPOB_Date_Str As Integer = 12
        Dim ij As New Inventory_Module.FrmTransaksiSO_SetDelivery

        Try
            If e.RowIndex = -1 Then
                Exit Sub
            End If
            If e.ColumnIndex = Me.ColSOB_PromiseToDeliver.Index Then
                If Me.dgvItemSOB.Item(Me.ColSOB_PromiseToDeliver.Index, e.RowIndex).Value <> "-" Then
                    ij.vPrimarySO = Me.dgvItemSOB.Item(Me.ColSOB_PrimarySO.Index, e.RowIndex).Value
                    ij.VIdentifier = Me.dgvItemSOB.Item(Me.ColSOB_IdentifierSOItem.Index, e.RowIndex).Value
                    ij.vPrimaryCodeNo = Me.txtPrimaryCodeNo.Text
                    ij.vPrimaryUnitDefault = Me.ObjProduct.GetPrimaryUnit(Me.txtPrimaryCodeNo.Text)
                    ij.vPrimaryUnitMulti = Me.dgvItemSOB.Item(Me.ColSOB_PrimaryUnitMulti.Index, e.RowIndex).Value
                    'ij.FrmRef = Me
                    ij.vReadOnly = True
                    ij.ShowDialog()
                End If
            End If

            Me.dgvItemSOB.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Me.dgvItemSOB.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            If e.ColumnIndex = IndexPOBStr Then
                If Me.SortDirection = EnumPOBSort.Desc Or Me.SortDirection = EnumPOBSort.NoSort Then
                    Me.dgvItemSOB.Sort(Me.dgvItemSOB.Columns(IndexPOB), System.ComponentModel.ListSortDirection.Ascending)
                    Me.SortDirection = EnumPOBSort.Asc
                Else
                    Me.dgvItemSOB.Sort(Me.dgvItemSOB.Columns(IndexPOB), System.ComponentModel.ListSortDirection.Descending)
                    Me.SortDirection = EnumPOBSort.Desc
                End If
                Me.SortDirection_Date = EnumPOBSort.NoSort
            ElseIf e.ColumnIndex = IndexPOB_Date_Str Then
                If Me.SortDirection_Date = EnumPOBSort.Desc Or Me.SortDirection_Date = EnumPOBSort.NoSort Then
                    Me.dgvItemSOB.Sort(Me.dgvItemSOB.Columns(IndexPOB_Date), System.ComponentModel.ListSortDirection.Ascending)
                    Me.SortDirection_Date = EnumPOBSort.Asc
                Else
                    Me.dgvItemSOB.Sort(Me.dgvItemSOB.Columns(IndexPOB_Date), System.ComponentModel.ListSortDirection.Descending)
                    Me.SortDirection_Date = EnumPOBSort.Desc
                End If
                Me.SortDirection = EnumPOBSort.NoSort
            Else
                Me.SortDirection = EnumPOBSort.NoSort
                Me.SortDirection_Date = EnumPOBSort.NoSort
            End If

            Me.dgvItemSOB.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij = Nothing
        End Try
    End Sub

    Private Sub cboUnitPrice_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboMA_Currency.SelectedValueChanged
       Dim ij As New Core_Module.frmMasterProduct_PriceList

        Try
            If Me.cboMA_Currency.SelectedIndex <> -1 Then
                Me.TxtPrimaryCurrency.Text = Me.ObjCur.GetPrimary(Me.cboMA_Currency.Text)
                If Me.cboMA_Currency.Enabled = True And Me.txtPrimaryCodeNo.Text <> "-1" Then
                    If Me.TxtMA_CustCode.Text <> "" Then
                        Me.ObjMP.IsiCustomer(Me.txtCodeNo.Text, Me.TxtPrimaryCurrency.Text, Me.TxtMA_PrimaryCust.Text, _
                        Me.DGVHarga, 2, ij, frmMasterProduct_PriceList.EnumCustomerGroupTipe.Customer_Group, Me.ObjCur.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text))
                    Else
                        Me.ObjMP.IsiByPrimaryProduct(Me.DGVHarga, 2, ij, Me.ObjCur.GetDefaultPrecission(Me.TxtPrimaryCurrency.Text), _
                        Me.txtPrimaryCodeNo.Text, Me.TxtPrimaryCurrency.Text)
                    End If
                Else
                    Me.TxtUnitPrice.Text = "0"
                    Me.DGVHarga.Rows.Clear()
                End If
            Else
                Me.TxtUnitPrice.Text = "0"
                Me.DGVHarga.Rows.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij = Nothing
        End Try
    End Sub

    Private Sub TxtMA_UserCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtMA_UserCode.KeyDown
        Dim i As New FrmCariMasterUser
        i.backColorAll = Me.BackColorAll

        Try
            If e.KeyValue = 13 Then
                If Me.ObjUser.ApakahSudahAda(Me.TxtMA_UserCode.Text) = True Then
                    Me.TxtMAX_UserName.Text = Me.ObjUser.GetNama(Me.TxtMA_UserCode.Text)
                Else
                    'ada 2 Customer dengan kode yang sama atau tidak ada Customer yang dimaksud
                    If Me.TxtMA_UserCode.Text <> "" Then
                        i.TxtCari.Text = Me.TxtMA_UserCode.Text
                    End If

                    i.ShowDialog()
                    If i.Kodeuser <> "" Then
                        Me.TxtMA_UserCode.Text = i.Kodeuser
                        Me.TxtMAX_UserName.Text = Me.ObjUser.GetNama(Me.TxtMA_UserCode.Text)
                    Else
                        Me.TxtMA_UserCode.Clear()
                        Me.TxtMAX_UserName.Clear()
                        Me.TxtMA_UserCode.Focus()
                    End If
                End If
            End If
            Me.UpdateUserCodeLain()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtMA_UserCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMA_UserCode.Leave
        Dim i As New FrmCariMasterUser
        i.backColorAll = Me.BackColorAll

        Try
            If Me.TxtMA_UserCode.Text <> "" Then
                If Me.ObjUser.ApakahSudahAda(Me.TxtMA_UserCode.Text) Then
                    Me.TxtMAX_UserName.Text = Me.ObjUser.GetNama(Me.TxtMA_UserCode.Text)
                Else
                    If Me.TxtMA_UserCode.Text <> "" Then
                        i.TxtCari.Text = Me.TxtMA_UserCode.Text
                    End If

                    i.ShowDialog()
                    If i.Kodeuser <> "" Then
                        Me.TxtMA_UserCode.Text = i.Kodeuser
                        Me.TxtMAX_UserName.Text = Me.ObjUser.GetNama(Me.TxtMA_UserCode.Text)
                    Else
                        Me.TxtMA_UserCode.Clear()
                        Me.TxtMAX_UserName.Clear()
                        Me.TxtMA_UserCode.Focus()
                    End If
                End If
            Else
                Me.TxtMA_UserCode.Clear()
                Me.TxtMAX_UserName.Clear()
            End If
            Me.UpdateUserCodeLain()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnMA_FindUser_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMA_FindUser.Click
        Dim i As New FrmCariMasterUser
        i.backColorAll = Me.BackColorAll

        Try
            If Me.TxtMA_UserCode.Text <> "" Then
                i.TxtCari.Text = Me.TxtMA_UserCode.Text
            End If

            i.ShowDialog()
            If i.Kodeuser <> "" Then
                Me.TxtMA_UserCode.Text = i.Kodeuser
                Me.TxtMAX_UserName.Text = Me.ObjUser.GetNama(Me.TxtMA_UserCode.Text)
            Else
                Me.TxtMA_UserCode.Clear()
                Me.TxtMAX_UserName.Clear()
                Me.TxtMA_UserCode.Focus()
            End If

            Me.UpdateUserCodeLain()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub UpdateUserCodeLain()
        Try
            Me.TxtAR_UserCode.Text = Me.TxtMA_UserCode.Text
            Me.TxtAR_UserCodeName.Text = Me.TxtAR_UserCodeName.Text
            Me.TxtUP_UserCode.Text = Me.TxtMA_UserCode.Text
            Me.TxtUP_UserName.Text = Me.TxtMAX_UserName.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMA_CustCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtMA_CustCode.KeyDown
        Dim i As New FrmCariMasterCustomer
        i.BackColorAll = Me.BackColorAll

        Try
            If e.KeyValue = 13 Then
                If Me.ObjCustomer.ApakahSudahAda(Me.TxtMA_CustCode.Text) = True Then
                    Me.TxtMA_PrimaryCust.Text = Me.ObjCustomer.GetPrimary(Me.TxtMA_CustCode.Text)
                    Me.TxtMA_TxtCustName.Text = Me.ObjCustomer.GetCustName(Me.TxtMA_PrimaryCust.Text)
                Else
                    'ada 2 Customer dengan kode yang sama atau tidak ada Customer yang dimaksud
                    If Me.TxtMA_CustCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtMA_CustCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCustomer > 0 Then
                        Me.TxtMA_PrimaryCust.Text = i.PrimaryCustomer
                        Me.TxtMA_CustCode.Text = Me.ObjCustomer.GetCustomerCode(Me.TxtMA_PrimaryCust.Text)
                        Me.TxtMA_TxtCustName.Text = Me.ObjCustomer.GetCustName(Me.TxtMA_PrimaryCust.Text)
                    Else
                        Me.TxtMA_CustCode.Clear()
                        Me.TxtMA_TxtCustName.Clear()
                        Me.TxtMA_PrimaryCust.Text = "-1"
                        Me.TxtMA_CustCode.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtMA_CustCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMA_CustCode.Leave
        Dim i As New FrmCariMasterCustomer
        i.BackColorAll = Me.BackColorAll

        Try
            If Me.TxtMA_CustCode.Text <> "" Then
                If Me.ObjCustomer.ApakahSudahAda(Me.TxtMA_CustCode.Text) Then
                    Dim PrimaryBaru As Decimal = Me.ObjCustomer.GetPrimary(Me.TxtMA_CustCode.Text)
                    If PrimaryBaru <> Me.TxtMA_PrimaryCust.Text Then
                        Me.TxtMA_PrimaryCust.Text = PrimaryBaru
                        Me.TxtMA_TxtCustName.Text = Me.ObjCustomer.GetCustName(Me.TxtMA_PrimaryCust.Text)
                    End If
                Else
                    If Me.TxtMA_CustCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtMA_CustCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCustomer > 0 Then
                        Me.TxtMA_PrimaryCust.Text = i.PrimaryCustomer
                        Me.TxtMA_CustCode.Text = Me.ObjCustomer.GetCustomerCode(Me.TxtMA_PrimaryCust.Text)
                        Me.TxtMA_TxtCustName.Text = Me.ObjCustomer.GetCustName(Me.TxtMA_PrimaryCust.Text)
                    Else
                        Me.TxtMA_CustCode.Clear()
                        Me.TxtMA_TxtCustName.Clear()
                        Me.TxtMA_PrimaryCust.Text = "-1"
                        Me.TxtMA_CustCode.Focus()
                    End If
                End If
            Else
                Me.TxtMA_CustCode.Clear()
                Me.TxtMA_TxtCustName.Clear()
                Me.TxtMA_PrimaryCust.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnMA_FindCust_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMA_FindCust.Click
        Dim i As New FrmCariMasterCustomer
        i.BackColorAll = Me.BackColorAll
        Try
            i.cboCriteria.SelectedIndex = 0
            If Me.TxtMA_CustCode.Text <> "" Then
                i.txtCari.Text = Me.TxtMA_CustCode.Text
            End If

            i.ShowDialog()
            If i.PrimaryCustomer > 0 Then
                Me.TxtMA_PrimaryCust.Text = i.PrimaryCustomer
                Me.TxtMA_CustCode.Text = Me.ObjCustomer.GetCustomerCode(Me.TxtMA_PrimaryCust.Text)
                Me.TxtMA_TxtCustName.Text = Me.ObjCustomer.GetCustName(Me.TxtMA_PrimaryCust.Text)
            Else
                Me.TxtMA_CustCode.Clear()
                Me.TxtMA_TxtCustName.Clear()
                Me.TxtMA_PrimaryCust.Text = "-1"
                Me.TxtMA_CustCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnMA_FilterSalesHistory_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMA_FilterSalesHistory.Click
        Dim DTA As New DseTransaksiSOTableAdapters.DataTable4TableAdapter
        Dim TBL As New DseTransaksiSO.DataTable4DataTable
        Dim row As DseTransaksiSO.DataTable4Row
        Dim ApakahAdaCustPO As Boolean = False
        Dim DecMulti As Integer = 0

        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Me.dgvSH.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            'isikan data sales history
            If Me.TxtMA_UserCode.Text.Trim = "" Then
                If Me.TxtMA_CustCode.Text.Trim = "" Then
                    DTA.FillBy_All(TBL, Me.dtpSH_To.Value, Me.dtpSH_From.Value)
                Else
                    DTA.FillBy_PrimaryCust2(TBL, Me.dtpSH_To.Value, Me.dtpSH_From.Value, Me.TxtMA_PrimaryCust.Text)
                End If
            Else
                If Me.TxtMA_CustCode.Text.Trim = "" Then
                    DTA.Fill(TBL, Me.TxtMA_UserCode.Text, Me.dtpSH_To.Value, Me.dtpSH_From.Value)
                Else
                    DTA.FillBy_PrimaryCust(TBL, Me.TxtMA_UserCode.Text, Me.dtpSH_To.Value, Me.dtpSH_From.Value, Me.TxtMA_PrimaryCust.Text)
                End If
            End If

            'buang yang code no-nya tidak masuk
            If Me.TxtMA_PrimaryCodeNo.Text <> "" And Me.TxtMA_PrimaryCodeNo.Text <> "-1" Then
                For ij As Double = TBL.Rows.Count - 1 To 0 Step -1
                    row = TBL.Rows(ij)
                    If row.Primary_Product <> Me.TxtMA_PrimaryCodeNo.Text Then
                        TBL.Rows.RemoveAt(ij)
                    End If
                Next
            End If
            'buang yang primary so-nya tidak masuk
            If Me.TxtMA_PrimarySO.Text <> "" And Me.TxtMA_PrimarySO.Text <> "-1" Then
                For ij As Double = TBL.Rows.Count - 1 To 0 Step -1
                    row = TBL.Rows(ij)
                    If row.Primary_SO <> Me.TxtMA_PrimarySO.Text Then
                        TBL.Rows.RemoveAt(ij)
                    End If
                Next
            End If

            Me.dgvSH.Rows.Clear()
            For Each row In TBL.Rows
                DecMulti = Me.ObjProduct.GetAngkaDecimal(row.Primary_Product, row.Primary_Units_Multi)

                Me.dgvSH.Rows.Add()
                rowDGV = Me.dgvSH.Rows(Me.dgvSH.Rows.Count - 1)
                With rowDGV
                    .Cells(Me.ColSH_Cur.Index).Value = row.Currency_Code
                    .Cells(Me.ColSH_Cust.Index).Value = row.Customer_Name
                    .Cells(Me.ColSH_CustPONumber.Index).Value = row.Cust_PO_No
                    .Cells(Me.ColSH_Discount.Index).Value = Me.ObjCd.DigitGrouping(row.Discount, 2)
                    .Cells(Me.ColSH_NettPrice.Index).Value = Me.ObjCd.DigitGrouping(row.Nett_Price, row.Default_Precision)
                    .Cells(Me.ColSH_No.Index).Value = Me.dgvSH.Rows.Count
                    .Cells(Me.ColSH_Qty.Index).Value = Me.ObjCd.DigitGrouping(row.Qty_Multi, DecMulti)
                    .Cells(Me.ColSH_Unit.Index).Value = row.Unit_Code_Multi
                    .Cells(Me.ColSH_SOB.Index).Value = Me.ObjCd.DigitGrouping(row.SOB_Multi, DecMulti)
                    .Cells(Me.ColSH_QtyDefault.Index).Value = Me.ObjCd.DigitGrouping(row.Qty, row.Def_Prec_Prod)
                    .Cells(Me.ColSH_Unit_Default.Index).Value = row.Unit_Code
                    .Cells(Me.ColSH_SOB_Default.Index).Value = Me.ObjCd.DigitGrouping(row.SOB, row.Def_Prec_Prod)
                    .Cells(Me.ColSH_SODate.Index).Value = row._Date
                    .Cells(Me.ColSH_SONumber.Index).Value = row.SO_Number
                    .Cells(Me.ColSH_SubTotal.Index).Value = Me.ObjCd.DigitGrouping(row.Amount, row.Default_Precision)
                    .Cells(Me.ColSH_UnitPrice.Index).Value = Me.ObjCd.DigitGrouping(row.Unit_Price, row.Default_Precision)
                    .Cells(Me.ColSH_Identifier.Index).Value = row.identifier
                    .Cells(Me.ColSH_CodeNo.Index).Value = row.Product_Code
                    .Cells(Me.ColSH_Desc.Index).Value = row.Product_Desc
                End With
                If row.Cust_PO_No.Trim <> "" Then
                    ApakahAdaCustPO = True
                End If
            Next

            Me.ColSH_SODate.DefaultCellStyle.Format = "dd MMM yyyy"

            If ApakahAdaCustPO = True Then
                Me.ColSH_CustPONumber.Visible = True
            Else
                Me.ColSH_CustPONumber.Visible = False
            End If
            If Me.TxtMA_CustCode.Text.Trim = "" Then
                Me.ColSH_Cust.Visible = True
            Else
                Me.ColSH_Cust.Visible = False
            End If

            If Me.TxtMA_PrimaryCodeNo.Text = "" Or Me.TxtMA_PrimaryCodeNo.Text = "-1" Then
                Me.ColSH_Desc.Visible = True
                Me.ColSH_CodeNo.Visible = True
            Else
                Me.ColSH_CodeNo.Visible = False
                Me.ColSH_Desc.Visible = False
            End If

            Me.dgvSH.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Private Sub dgvSH_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvSH.CellClick
        Try
            If e.ColumnIndex = -1 Or e.RowIndex = -1 Then
                Exit Sub
            Else
                If e.ColumnIndex = Me.ColSH_SOB.Index Then
                    Me.ObjSO1.ExplainDelivery(Me.dgvSH.Item(Me.ColSH_Identifier.Index, e.RowIndex).Value)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSF_UserCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtSF_UserCode.KeyDown
        Dim i As New FrmCariMasterUser
        i.backColorAll = Me.BackColorAll

        Try
            If e.KeyValue = 13 Then
                If Me.ObjUser.ApakahSudahAda(Me.TxtSF_UserCode.Text) = True Then
                    Me.TxtSF_UserName.Text = Me.ObjUser.GetNama(Me.TxtSF_UserCode.Text)
                Else
                    'ada 2 Customer dengan kode yang sama atau tidak ada Customer yang dimaksud
                    If Me.TxtSF_UserCode.Text <> "" Then
                        i.TxtCari.Text = Me.TxtSF_UserCode.Text
                    End If

                    i.ShowDialog()
                    If i.Kodeuser <> "" Then
                        Me.TxtSF_UserCode.Text = i.Kodeuser
                        Me.TxtSF_UserName.Text = Me.ObjUser.GetNama(Me.TxtSF_UserCode.Text)
                    Else
                        Me.TxtSF_UserCode.Clear()
                        Me.TxtSF_UserName.Clear()
                        Me.TxtSF_UserCode.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtSF_UserCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSF_UserCode.Leave
        Dim i As New FrmCariMasterUser
        i.backColorAll = Me.BackColorAll

        Try
            If Me.TxtSF_UserCode.Text <> "" Then
                If Me.ObjUser.ApakahSudahAda(Me.TxtSF_UserCode.Text) Then
                    Me.TxtSF_UserName.Text = Me.ObjUser.GetNama(Me.TxtSF_UserCode.Text)
                Else
                    If Me.TxtSF_UserCode.Text <> "" Then
                        i.TxtCari.Text = Me.TxtSF_UserCode.Text
                    End If

                    i.ShowDialog()
                    If i.Kodeuser <> "" Then
                        Me.TxtSF_UserCode.Text = i.Kodeuser
                        Me.TxtSF_UserName.Text = Me.ObjUser.GetNama(Me.TxtSF_UserCode.Text)
                    Else
                        Me.TxtSF_UserCode.Clear()
                        Me.TxtSF_UserName.Clear()
                        Me.TxtSF_UserCode.Focus()
                    End If
                End If
            Else
                Me.TxtSF_UserCode.Clear()
                Me.TxtSF_UserName.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnSF_FindUser_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSF_FindUser.Click
        Dim i As New FrmCariMasterUser
        i.backColorAll = Me.BackColorAll

        Try
            If Me.TxtSF_UserCode.Text <> "" Then
                i.TxtCari.Text = Me.TxtSF_UserCode.Text
            End If

            i.ShowDialog()
            If i.Kodeuser <> "" Then
                Me.TxtSF_UserCode.Text = i.Kodeuser
                Me.TxtSF_UserName.Text = Me.ObjUser.GetNama(Me.TxtSF_UserCode.Text)
            Else
                Me.TxtSF_UserCode.Clear()
                Me.TxtSF_UserName.Clear()
                Me.TxtSF_UserCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtSF_CustCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtSF_CustCode.KeyDown
        Dim i As New FrmCariMasterCustomer
        i.BackColorAll = Me.BackColorAll

        Try
            If e.KeyValue = 13 Then
                If Me.ObjCustomer.ApakahSudahAda(Me.TxtSF_CustCode.Text) = True Then
                    Me.TxtSF_PrimaryCust.Text = Me.ObjCustomer.GetPrimary(Me.TxtSF_CustCode.Text)
                    Me.TxtSF_TxtCustName.Text = Me.ObjCustomer.GetCustName(Me.TxtSF_PrimaryCust.Text)
                Else
                    'ada 2 Customer dengan kode yang sama atau tidak ada Customer yang dimaksud
                    If Me.TxtSF_CustCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtSF_CustCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCustomer > 0 Then
                        Me.TxtSF_PrimaryCust.Text = i.PrimaryCustomer
                        Me.TxtSF_CustCode.Text = Me.ObjCustomer.GetCustomerCode(Me.TxtSF_PrimaryCust.Text)
                        Me.TxtSF_TxtCustName.Text = Me.ObjCustomer.GetCustName(Me.TxtSF_PrimaryCust.Text)
                    Else
                        Me.TxtSF_CustCode.Clear()
                        Me.TxtSF_TxtCustName.Clear()
                        Me.TxtSF_PrimaryCust.Text = "-1"
                        Me.TxtSF_CustCode.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtSF_CustCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSF_CustCode.Leave
        Dim i As New FrmCariMasterCustomer
        i.BackColorAll = Me.BackColorAll

        Try
            If Me.TxtSF_CustCode.Text <> "" Then
                If Me.ObjCustomer.ApakahSudahAda(Me.TxtSF_CustCode.Text) Then
                    Dim PrimaryBaru As Decimal = Me.ObjCustomer.GetPrimary(Me.TxtSF_CustCode.Text)
                    If PrimaryBaru <> Me.TxtSF_PrimaryCust.Text Then
                        Me.TxtSF_PrimaryCust.Text = PrimaryBaru
                        Me.TxtSF_TxtCustName.Text = Me.ObjCustomer.GetCustName(Me.TxtSF_PrimaryCust.Text)
                    End If
                Else
                    If Me.TxtSF_CustCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtSF_CustCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCustomer > 0 Then
                        Me.TxtSF_PrimaryCust.Text = i.PrimaryCustomer
                        Me.TxtSF_CustCode.Text = Me.ObjCustomer.GetCustomerCode(Me.TxtSF_PrimaryCust.Text)
                        Me.TxtSF_TxtCustName.Text = Me.ObjCustomer.GetCustName(Me.TxtSF_PrimaryCust.Text)
                    Else
                        Me.TxtSF_CustCode.Clear()
                        Me.TxtSF_TxtCustName.Clear()
                        Me.TxtSF_PrimaryCust.Text = "-1"
                        Me.TxtSF_CustCode.Focus()
                    End If
                End If
            Else
                Me.TxtSF_CustCode.Clear()
                Me.TxtSF_TxtCustName.Clear()
                Me.TxtSF_PrimaryCust.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnSF_FindCust_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSF_FindCust.Click
        Dim i As New FrmCariMasterCustomer
        i.BackColorAll = Me.BackColorAll
        Try
            i.cboCriteria.SelectedIndex = 0
            If Me.TxtSF_CustCode.Text <> "" Then
                i.txtCari.Text = Me.TxtSF_CustCode.Text
            End If

            i.ShowDialog()
            If i.PrimaryCustomer > 0 Then
                Me.TxtSF_PrimaryCust.Text = i.PrimaryCustomer
                Me.TxtSF_CustCode.Text = Me.ObjCustomer.GetCustomerCode(Me.TxtSF_PrimaryCust.Text)
                Me.TxtSF_TxtCustName.Text = Me.ObjCustomer.GetCustName(Me.TxtSF_PrimaryCust.Text)
            Else
                Me.TxtSF_CustCode.Clear()
                Me.TxtSF_TxtCustName.Clear()
                Me.TxtSF_PrimaryCust.Text = "-1"
                Me.TxtSF_CustCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnSF_Filter_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSF_Filter.Click
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Me.ObjSF.HitungShippingForecast(Me.vTBLItemDelivery, Me.vTBLItemPOB, Me.vTBLItemStock, -1, Me.BackColorAll, Me.TxtSF_PrimaryCodeNo.Text)
            Me.LoadDGV_SF()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSF_ProductCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtSF_ProductCode.KeyDown
        Dim i As New frmCariMasterProduct
        Dim Jml As Long = 0

        Try
            If e.KeyValue = 13 Then
                i.BackColorAll = Me.BackColorAll
                Jml = Me.ObjProduct.GetJml(Me.TxtSF_ProductCode.Text)
                If Jml = 1 Then
                    Me.TxtSF_PrimaryCodeNo.Text = Me.ObjProduct.GetPrimary(Me.TxtSF_ProductCode.Text)
                    Me.TxtSF_ProductDesc.Text = Me.ObjProduct.GetDescription(Me.TxtSF_PrimaryCodeNo.Text)
                ElseIf Jml > 1 Or Jml = 0 Then
                    'ada 2 product dengan kode yang sama atau tidak ada product yang dimaksud
                    If Me.TxtSF_ProductCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtSF_ProductCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        Me.TxtSF_PrimaryCodeNo.Text = i.PrimaryProduct
                        Me.TxtSF_ProductCode.Text = Me.ObjProduct.GetProductCode(Me.TxtSF_PrimaryCodeNo.Text)
                        Me.TxtSF_ProductDesc.Text = Me.ObjProduct.GetDescription(Me.TxtSF_PrimaryCodeNo.Text)
                    Else
                        Me.TxtSF_ProductCode.Clear()
                        Me.TxtSF_PrimaryCodeNo.Text = "-1"
                        Me.TxtSF_ProductCode.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtSF_ProductCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSF_ProductCode.Leave
        Dim i As New frmCariMasterProduct

        Try
            i.BackColorAll = Me.BackColorAll
            If Me.TxtSF_ProductCode.Text <> "" Then
                If Me.ObjProduct.GetJml(Me.TxtSF_ProductCode.Text) = 1 Then
                    Me.TxtSF_PrimaryCodeNo.Text = Me.ObjProduct.GetPrimary(Me.TxtSF_ProductCode.Text)
                    Me.TxtSF_ProductDesc.Text = Me.ObjProduct.GetDescription(Me.TxtSF_PrimaryCodeNo.Text)
                Else
                    If Me.TxtSF_ProductCode.Text <> "" Then
                        i.txtCari.Text = Me.TxtSF_ProductCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        Me.TxtSF_PrimaryCodeNo.Text = i.PrimaryProduct
                        Me.TxtSF_ProductCode.Text = Me.ObjProduct.GetProductCode(Me.TxtSF_PrimaryCodeNo.Text)
                        Me.TxtSF_ProductDesc.Text = Me.ObjProduct.GetDescription(Me.TxtSF_PrimaryCodeNo.Text)
                    Else
                        Me.TxtSF_ProductCode.Clear()
                        Me.TxtSF_PrimaryCodeNo.Text = "-1"
                        Me.TxtSF_ProductCode.Focus()
                    End If
                End If
            Else
                Me.TxtSF_PrimaryCodeNo.Text = "-1"
                Me.TxtSF_ProductDesc.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnSF_FindProductCode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSF_FindProductCode.Click
        Dim i As New frmCariMasterProduct

        Try
            i.BackColorAll = Me.BackColorAll
            i.cboCriteria.SelectedIndex = 0
            If Me.TxtSF_ProductCode.Text <> "" Then
                i.txtCari.Text = Me.TxtSF_ProductCode.Text
            End If

            i.ShowDialog()
            If i.PrimaryProduct > 0 Then
                Me.TxtSF_PrimaryCodeNo.Text = i.PrimaryProduct
                Me.TxtSF_ProductCode.Text = Me.ObjProduct.GetProductCode(Me.TxtSF_PrimaryCodeNo.Text)
                Me.TxtSF_ProductDesc.Text = Me.ObjProduct.GetDescription(Me.TxtSF_PrimaryCodeNo.Text)
            Else
                Me.TxtSF_ProductCode.Clear()
                Me.TxtSF_PrimaryCodeNo.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub dgvSF_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSF.Sorted
        Try
            Me.ObjCD.SetNumber(Me.dgvSF, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvSF_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvSF.CellMouseClick
        Try
            If e.ColumnIndex <> -1 Then
                If e.RowIndex <> -1 Then
                    If e.ColumnIndex = Me.ColSF_EstimateDelivery_Str.Index Then
                        Dim ij As New Inventory_Module.FrmTrans_MaterialPlanningDashboard_SF_DetailDelivery
                        ij.vReadOnly = True
                        ij.vTBLItemDelivery = Me.vTBLItemDelivery
                        ij.vTBLItemPOB = Me.vTBLItemPOB
                        ij.vTBLItemStock = Me.vTBLItemStock
                        ij.vPrimaryTarget = Me.dgvSF.Item(Me.ColSF_Primary.Index, e.RowIndex).Value
                        ij.vPrimaryCodeNoTarget = Me.dgvSF.Item(Me.ColSF_PrimaryCodeNo.Index, e.RowIndex).Value
                        ij.ShowDialog()
                        ij.Dispose()
                    ElseIf e.ColumnIndex = Me.ColSF_QtySOB.Index Then
                        Me.ObjSO1.ExplainDelivery(Me.dgvSF.Item(Me.ColSF_Identifier.Index, e.RowIndex).Value)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnSF_Print_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSF_Print.Click
        Dim i As New FrmTrans_SalesDashboard_MA_Print

        Try
            i.vFilter = Me.GetFilterSF
            i.vTBLItemSOB = Me.vTBLItemDelivery
            i.vAvDate = Me.dtpAvailable.Value
            i.vAvDateChecked = Me.dtpAvailable.Checked
            i.vEnumPrint = Me.cboAvailability.SelectedIndex

            i.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMA_CodeNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtMA_CodeNo.KeyDown
        Dim i As New frmCariMasterProduct
        Dim Jml As Long = 0

        Try
            If e.KeyValue = 13 Then
                i.BackColorAll = Me.BackColorAll
                Jml = Me.ObjProduct.GetJml(Me.TxtMA_CodeNo.Text)
                If Jml = 1 Then
                    Me.TxtMA_PrimaryCodeNo.Text = Me.ObjProduct.GetPrimary(Me.TxtMA_CodeNo.Text)
                    Me.TxtMA_CodeDesc.Text = Me.ObjProduct.GetDescription(Me.TxtMA_PrimaryCodeNo.Text)
                ElseIf Jml > 1 Or Jml = 0 Then
                    'ada 2 product dengan kode yang sama atau tidak ada product yang dimaksud
                    If Me.TxtMA_CodeNo.Text <> "" Then
                        i.txtCari.Text = Me.TxtMA_CodeNo.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        Me.TxtMA_PrimaryCodeNo.Text = i.PrimaryProduct
                        Me.TxtMA_CodeNo.Text = Me.ObjProduct.GetProductCode(Me.TxtMA_PrimaryCodeNo.Text)
                        Me.TxtMA_CodeDesc.Text = Me.ObjProduct.GetDescription(Me.TxtMA_PrimaryCodeNo.Text)
                    Else
                        Me.TxtMA_CodeNo.Clear()
                        Me.TxtMA_PrimaryCodeNo.Text = "-1"
                        Me.TxtMA_CodeNo.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtMA_CodeNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMA_CodeNo.Leave
        Dim i As New frmCariMasterProduct

        Try
            i.BackColorAll = Me.BackColorAll
            If Me.TxtMA_CodeNo.Text <> "" Then
                If Me.ObjProduct.GetJml(Me.TxtMA_CodeNo.Text) = 1 Then
                    Me.TxtMA_PrimaryCodeNo.Text = Me.ObjProduct.GetPrimary(Me.TxtMA_CodeNo.Text)
                    Me.TxtMA_CodeDesc.Text = Me.ObjProduct.GetDescription(Me.TxtMA_PrimaryCodeNo.Text)
                Else
                    If Me.TxtMA_CodeNo.Text <> "" Then
                        i.txtCari.Text = Me.TxtMA_CodeNo.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        Me.TxtMA_PrimaryCodeNo.Text = i.PrimaryProduct
                        Me.TxtMA_CodeNo.Text = Me.ObjProduct.GetProductCode(Me.TxtMA_PrimaryCodeNo.Text)
                        Me.TxtMA_CodeDesc.Text = Me.ObjProduct.GetDescription(Me.TxtMA_PrimaryCodeNo.Text)
                    Else
                        Me.TxtMA_CodeNo.Clear()
                        Me.TxtMA_PrimaryCodeNo.Text = "-1"
                        Me.TxtMA_CodeNo.Focus()
                    End If
                End If
            Else
                Me.TxtMA_PrimaryCodeNo.Text = "-1"
                Me.TxtMA_CodeDesc.Clear()
            End If

            If Me.TxtMA_PrimaryCodeNo.Text = "" Or Me.TxtMA_PrimaryCodeNo.Text = "-1" Then
                Me.ColSH_Desc.Visible = True
                Me.ColSH_CodeNo.Visible = True
            Else
                Me.ColSH_CodeNo.Visible = False
                Me.ColSH_Desc.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnMA_FindProductCode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMA_FindCodeNo.Click
        Dim i As New frmCariMasterProduct

        Try
            i.BackColorAll = Me.BackColorAll
            i.cboCriteria.SelectedIndex = 0
            If Me.TxtMA_CodeNo.Text <> "" Then
                i.txtCari.Text = Me.TxtMA_CodeNo.Text
            End If

            i.ShowDialog()
            If i.PrimaryProduct > 0 Then
                Me.TxtMA_PrimaryCodeNo.Text = i.PrimaryProduct
                Me.TxtMA_CodeNo.Text = Me.ObjProduct.GetProductCode(Me.TxtMA_PrimaryCodeNo.Text)
                Me.TxtMA_CodeDesc.Text = Me.ObjProduct.GetDescription(Me.TxtMA_PrimaryCodeNo.Text)
            Else
                Me.TxtMA_CodeNo.Clear()
                Me.TxtMA_PrimaryCodeNo.Text = "-1"
            End If

            If Me.TxtMA_PrimaryCodeNo.Text = "" Or Me.TxtMA_PrimaryCodeNo.Text = "-1" Then
                Me.ColSH_Desc.Visible = True
                Me.ColSH_CodeNo.Visible = True
            Else
                Me.ColSH_CodeNo.Visible = False
                Me.ColSH_Desc.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtMA_CodeNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMA_CodeNo.GotFocus
        Try
            If Me.TxtMA_CodeNo.ReadOnly = False And Me.TxtMA_CodeNo.Enabled = True Then
                Me.TxtMA_CodeNo.BackColor = Me.BackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMA_CodeNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMA_CodeNo.LostFocus
        Try
            If Me.TxtMA_CodeNo.ReadOnly = False And Me.TxtMA_CodeNo.Enabled = True Then
                Me.TxtMA_CodeNo.BackColor = Me.txtPrimaryCodeNo.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMA_SONumber_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMA_SONumber.GotFocus
        Try
            If Me.TxtMA_SONumber.ReadOnly = False And Me.TxtMA_SONumber.Enabled = True Then
                Me.TxtMA_SONumber.BackColor = Me.BackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMA_SONumber_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMA_SONumber.LostFocus
        Try
            If Me.TxtMA_SONumber.ReadOnly = False And Me.TxtMA_SONumber.Enabled = True Then
                Me.TxtMA_SONumber.BackColor = Me.txtPrimaryCodeNo.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtMA_SONumber_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtMA_SONumber.KeyDown
        Dim i As New FrmCariTransaksiSO
        Dim Jml As Long = 0

        Try
            If e.KeyValue = 13 Then
                If Me.ObjSO.ApakahSudahAda(Me.TxtMA_SONumber.Text) = True Then
                    Me.TxtMA_PrimarySO.Text = Me.ObjSO.GetPrimary(Me.TxtMA_SONumber.Text)
                    Me.btnMA_FilterSalesHistory.Focus()
                ElseIf Jml > 1 Or Jml = 0 Then
                    'ada 2 Customer dengan kode yang sama atau tidak ada Customer yang dimaksud
                    If Me.TxtMA_SONumber.Text <> "" Then
                        i.txtCari.Text = Me.TxtMA_SONumber.Text
                    End If
                    i.BackColorAll = Me.BackColorAll
                    i.ShowDialog()

                    If i.PrimaryTransaksi > 0 Then
                        Me.TxtMA_PrimarySO.Text = i.PrimaryTransaksi
                        Me.TxtMA_SONumber.Text = Me.ObjSO.GetNoSO(Me.TxtMA_PrimarySO.Text)
                        Me.btnMA_FilterSalesHistory.Focus()
                    Else
                        Me.TxtMA_SONumber.Clear()
                        Me.TxtMA_PrimarySO.Text = "-1"
                        Me.TxtMA_SONumber.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtMA_SONumber_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtMA_SONumber.Leave
        Dim i As New FrmCariTransaksiSO

        Try
            If Me.TxtMA_SONumber.Text <> "" Then
                If Me.ObjSO.ApakahSudahAda(Me.TxtMA_SONumber.Text) = True Then
                    Me.TxtMA_PrimarySO.Text = Me.ObjSO.GetPrimary(Me.TxtMA_SONumber.Text)
                Else
                    i.BackColorAll = Me.BackColorAll
                    If Me.TxtMA_SONumber.Text <> "" Then
                        i.txtCari.Text = Me.TxtMA_SONumber.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryTransaksi > 0 Then
                        Me.TxtMA_PrimarySO.Text = i.PrimaryTransaksi
                        Me.TxtMA_SONumber.Text = Me.ObjSO.GetNoSO(Me.TxtMA_PrimarySO.Text)
                    Else
                        Me.TxtMA_SONumber.Clear()
                        Me.TxtMA_PrimarySO.Text = "-1"
                        Me.KosongkanItem()
                    End If
                End If
            Else
                Me.TxtMA_PrimarySO.Text = "-1"
            End If
            Me.UpdateTglSO()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub UpdateTglSO()
        Dim TglSO As Date
        Try
            If Me.TxtMA_PrimarySO.Text = "-1" Or Me.TxtMA_PrimarySO.Text = "" Then
            Else
                TglSO = Me.ObjSO1.GetTanggal(Me.TxtMA_PrimarySO.Text)
                If TglSO.Date < Me.dtpSH_From.Value.Date Then
                    Me.dtpSH_From.Value = TglSO
                ElseIf TglSO.Date > Me.dtpSH_To.Value.Date Then
                    Me.dtpSH_To.Value = TglSO
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnMA_FindSO_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMA_FindSO.Click
        Dim i As New FrmCariTransaksiSO

        Try
            i.BackColorAll = Me.BackColorAll

            If Me.TxtMA_SONumber.Text <> "" Then
                i.txtCari.Text = Me.TxtMA_SONumber.Text
            End If

            i.ShowDialog()
            If i.PrimaryTransaksi > 0 Then
                Me.TxtMA_PrimarySO.Text = i.PrimaryTransaksi
                Me.TxtMA_SONumber.Text = Me.ObjSO.GetNoSO(Me.TxtMA_PrimarySO.Text)
                Me.btnMA_FilterSalesHistory.Focus()
            Else
                Me.TxtMA_SONumber.Clear()
                Me.TxtMA_PrimarySO.Text = "-1"
                Me.TxtMA_SONumber.Focus()
            End If
            Me.UpdateTglSO()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtSF_SONumber_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSF_SONumber.GotFocus
        Try
            If Me.TxtSF_SONumber.ReadOnly = False And Me.TxtSF_SONumber.Enabled = True Then
                Me.TxtSF_SONumber.BackColor = Me.BackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSF_SONumber_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSF_SONumber.LostFocus
        Try
            If Me.TxtSF_SONumber.ReadOnly = False And Me.TxtSF_SONumber.Enabled = True Then
                Me.TxtSF_SONumber.BackColor = Me.txtPrimaryCodeNo.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSF_SONumber_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtSF_SONumber.KeyDown
        Dim i As New FrmCariTransaksiSO
        Dim Jml As Long = 0

        Try
            If e.KeyValue = 13 Then
                If Me.ObjSO.ApakahSudahAda(Me.TxtSF_SONumber.Text) = True Then
                    Me.TxtSF_PrimarySO.Text = Me.ObjSO.GetPrimary(Me.TxtSF_SONumber.Text)
                    Me.btnSF_Filter.Focus()
                ElseIf Jml > 1 Or Jml = 0 Then
                    'ada 2 Customer dengan kode yang sama atau tidak ada Customer yang dimaksud
                    If Me.TxtSF_SONumber.Text <> "" Then
                        i.txtCari.Text = Me.TxtSF_SONumber.Text
                    End If
                    i.BackColorAll = Me.BackColorAll
                    i.ShowDialog()

                    If i.PrimaryTransaksi > 0 Then
                        Me.TxtSF_PrimarySO.Text = i.PrimaryTransaksi
                        Me.TxtSF_SONumber.Text = Me.ObjSO.GetNoSO(Me.TxtSF_PrimarySO.Text)
                        Me.btnSF_Filter.Focus()
                    Else
                        Me.TxtSF_SONumber.Clear()
                        Me.TxtSF_PrimarySO.Text = "-1"
                        Me.TxtSF_SONumber.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtSF_SONumber_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSF_SONumber.Leave
        Dim i As New FrmCariTransaksiSO

        Try
            If Me.TxtSF_SONumber.Text <> "" Then
                If Me.ObjSO.ApakahSudahAda(Me.TxtSF_SONumber.Text) = True Then
                    Me.TxtSF_PrimarySO.Text = Me.ObjSO.GetPrimary(Me.TxtSF_SONumber.Text)
                Else
                    i.BackColorAll = Me.BackColorAll
                    If Me.TxtSF_SONumber.Text <> "" Then
                        i.txtCari.Text = Me.TxtSF_SONumber.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryTransaksi > 0 Then
                        Me.TxtSF_PrimarySO.Text = i.PrimaryTransaksi
                        Me.TxtSF_SONumber.Text = Me.ObjSO.GetNoSO(Me.TxtSF_PrimarySO.Text)
                    Else
                        Me.TxtSF_SONumber.Clear()
                        Me.TxtSF_PrimarySO.Text = "-1"
                        Me.KosongkanItem()
                    End If
                End If
            Else
                Me.TxtSF_PrimarySO.Text = "-1"
            End If
            Me.UpdateTglSO()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnSF_FindSO_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSF_FindSO.Click
        Dim i As New FrmCariTransaksiSO

        Try
            i.BackColorAll = Me.BackColorAll

            If Me.TxtSF_SONumber.Text <> "" Then
                i.txtCari.Text = Me.TxtSF_SONumber.Text
            End If

            i.ShowDialog()
            If i.PrimaryTransaksi > 0 Then
                Me.TxtSF_PrimarySO.Text = i.PrimaryTransaksi
                Me.TxtSF_SONumber.Text = Me.ObjSO.GetNoSO(Me.TxtSF_PrimarySO.Text)
                Me.btnSF_Filter.Focus()
            Else
                Me.TxtSF_SONumber.Clear()
                Me.TxtSF_PrimarySO.Text = "-1"
                Me.TxtSF_SONumber.Focus()
            End If
            Me.UpdateTglSO()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub chkSODate_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSODate.CheckedChanged
        Try
            If Me.chkSODate.Checked = True Then
                Me.DtpFilterSF_From.Enabled = True
                Me.DtpFilterSF_To.Enabled = True
                Me.Label28.Enabled = True
            Else
                Me.DtpFilterSF_From.Enabled = False
                Me.DtpFilterSF_To.Enabled = False
                Me.Label28.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DtpFilterSF_From_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DtpFilterSF_From.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then
                Me.DtpFilterSF_To.Focus()
                Me.BackgroundWorker2.RunWorkerAsync()
                Me.DtpFilterSF_From.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub DtpFilterSF_To_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DtpFilterSF_To.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then
                Me.DtpFilterSF_From.Focus()
                Me.BackgroundWorker2.RunWorkerAsync()
                Me.DtpFilterSF_To.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Try
            Threading.Thread.Sleep(100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Try
            Me.LoadDGV_SF()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
