Option Explicit On
Option Strict Off

Imports Core_Module
Public Class FrmTransaksiSO
    Dim mTblMultiUnit As New DseTransaksiSO.tblmaster_product_item_multiunitDataTable
    Dim tblCboUnit As New DseTransaksiSO.CboUnitDataTable

    Dim ObjSettingUser As New Core_Module.ClsTransaksiSettingUser

    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public vTBLItemDelivery As New DseTransaksiSO.tbltrans_so_itemdeliveryDataTable
    Public vTBLItemDelivery2 As New DseTransaksiQuotation.tbltrans_quotation_itemdeliveryDataTable

    Dim baru As Boolean
    Dim objSO As New ClsTransaksiSalesOrder
    Dim objInv As New ClsTransaksiSalesInvoice
    Dim ObjCD As New ClsCekData
    Dim ObjProduct As New clsMasterProduct
    Dim ObjCustomer_sProductCode As New clsTransaksiSettingcustomerProductCodeNo
    Dim objQuotationOutItem As New ClsTransaksiQuotationItem
    Dim ObjCurrency As New Core_Module.clsMasterCurrency
    Dim ObjUser As New Core_Module.ClsUser
    Dim ObjCustomer As New clsMasterCustomer
    Dim ObjSOItem As New ClsTransaksiSalesOrderItem
    Dim ObjDOItem As New ClsTransaksiDOItem
    Dim ObjBookingInv As New ClsTransaksiBookingInvItem
    Dim ObjLicense As New ClsLicense
    Dim ObjSalesP As New Core_Module.ClsMasterSalesPerson
    Dim ObjNumbering As New Core_Module.ClsMasterTransactionNumbering
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting
    Dim ObjTaxRate As New ClsTransaksiSalesInvoiceTaxRate
    Dim ObjDeliveryPriority As New Core_Module.ClsMasterPriorityLevel
    Dim vrowItemDelivery As DseTransaksiSO.tbltrans_so_itemdeliveryRow
    Dim vrowItemDelivery2 As DseTransaksiQuotation.tbltrans_quotation_itemdeliveryRow
    Dim ObjOtorisasi As New Core_Module.ClsMasterTipeOtorisasi

    Dim DefaultPrecision As Integer = 0
    Dim buffer As String = ""
    Dim HitungOnRowsRemoved As Boolean = True

    Dim bufferToleransiDown As String = ""
    Dim bufferToleransiUp As String = ""

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.txtPrimaryCategory.BackColor, False)
            Me.ObjCD.SetBackColor(Me.TabPage1.Controls, Me.txtPrimaryCategory.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.txtPrimaryCategory.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.txtPrimaryCategory.BackColor, False)
            Me.ObjCD.SetBackColor(Me.TabPage2.Controls, Me.txtPrimaryCategory.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel1.Controls, Me.txtPrimaryCategory.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel2.Controls, Me.txtPrimaryCategory.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gboItem.Controls, Me.txtPrimaryCategory.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel1.Controls, Me.txtPrimaryCategory.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel2.Controls, Me.txtPrimaryCategory.BackColor, False)
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.TabPage1.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer4.Panel1.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer4.Panel2.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.TabPage2.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer3.Panel1.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer3.Panel2.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.gboItem.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)
            Me.RboAddressShipTo.Enabled = True
            Me.RboCustomerShipTo.Enabled = True
            Me.CancelButton = Me.btnCancel
            Me.CancelButton.DialogResult = Windows.Forms.DialogResult.None
            If Me.dgvItem.Rows.Count = 0 Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
            End If
            Me.dgvItem.AllowUserToDeleteRows = True
            Me.dgvItem.ReadOnly = False
            'Me.llbAddress.Enabled = True
            Me.llbDetailStock.Enabled = True
            llbUpdateAllDeliveryDate.Enabled = True
            Me.llbAddQuotation.Enabled = True
            Dim ij As New System.EventArgs
            If Me.chkUseCustCodeNo.Checked = True Then
                Me.btnFindCustomerProductCode.Enabled = True
                Me.BtnFindCodeNo.Enabled = False
            Else
                Me.btnFindCustomerProductCode.Enabled = False
                Me.BtnFindCodeNo.Enabled = True
            End If

            Me.KosongkanItem()
            If Me.ChkPriceCondition.Checked = True Then
                Me.CboPriceCondition.Enabled = True
                Me.cboPriceConditionItem.Enabled = False
            Else
                Me.CboPriceCondition.Enabled = False
                Me.cboPriceConditionItem.Enabled = True
            End If

            Me.LblUnitPrice.Enabled = True

            If Me.ObjNumbering.Apakah_Autonumber(ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.SO) = True Then
                Me.TxtSalesOrderNo.ReadOnly = True
                Me.TxtSalesOrderNo.BackColor = Me.TxtReadOnly.BackColor
                Me.TxtSalesOrderNo.TabStop = False
                Me.dtpDate.Focus()
            Else
                Me.TxtSalesOrderNo.ReadOnly = False
                Me.TxtSalesOrderNo.BackColor = Me.TxtPrimarySO.BackColor
                Me.TxtSalesOrderNo.TabStop = True
                Me.TxtSalesOrderNo.Focus()
            End If

            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiSO_Override_Date") = False Then
                Me.dtpDate.Enabled = False
            Else
                Me.dtpDate.Enabled = True
            End If
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiSO_Override_Promise_To_Deliver") = False Then
                Me.dtpdeliverydate.Enabled = False
                Me.dtpdeliverydate.Checked = False
            Else
                Me.dtpdeliverydate.Enabled = True
            End If
            Me.ColDeliveryHistory.Visible = False
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTrans_SalesDashboard") = False Then
                Me.llbDetailStock.Enabled = False
            Else
                Me.llbDetailStock.Enabled = True
            End If
            Me.dtpdeliverydate.Checked = False

            Me.EndEditDeliveryDateCont2()

            If Me.RboAddressShipTo.Checked Then
                CboAddressShipTo.Enabled = True
                LlbViewAddressShipTo.Enabled = True
                TxtCustomerCodeShipTo.Enabled = False
                BtnFindCustomerShipTo.Enabled = False
                CboAddressCustomerShipTo.Enabled = False
                LlbViewAddressCustomerShipTo.Enabled = False
            Else
                CboAddressShipTo.Enabled = False
                LlbViewAddressShipTo.Enabled = False
                TxtCustomerCodeShipTo.Enabled = True
                BtnFindCustomerShipTo.Enabled = True
                CboAddressCustomerShipTo.Enabled = True
                LlbViewAddressCustomerShipTo.Enabled = True
            End If
            Me.SetDPControl()

            'jika so sudah digunakan oleh do maka so tidak bisa ganti ship to customer
            Dim ObjDOItem As New ClsTransaksiDOItem
            If ObjDOItem.GetJmlBySO(Me.TxtPrimarySO.Text) > 0 Then
                Me.RboAddressShipTo.Enabled = False
                Me.RboCustomerShipTo.Enabled = False
                Me.TxtCustomerCodeShipTo.Enabled = False
                Me.TxtCustomerDescShipTo.Enabled = False
                Me.BtnFindCustomerShipTo.Enabled = False
            End If
            ObjDOItem = Nothing
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.TabPage1.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer4.Panel1.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer4.Panel2.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.TabPage2.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer3.Panel1.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer3.Panel2.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.gboItem.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)
            Me.CancelButton = Me.btnExit
            Me.dgvItem.AllowUserToDeleteRows = False
            Me.dgvItem.ReadOnly = True
            'Me.llbAddress.Enabled = False
            Me.llbDetailStock.Enabled = False
            Me.llbUpdateAllDeliveryDate.Enabled = False
            Me.llbAddQuotation.Enabled = False
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtReadOnly.BackColor, False)
            Me.ObjCD.SetBackColor(Me.TabPage1.Controls, Me.TxtReadOnly.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.TxtReadOnly.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.TxtReadOnly.BackColor, False)
            Me.ObjCD.SetBackColor(Me.TabPage2.Controls, Me.TxtReadOnly.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel1.Controls, Me.TxtReadOnly.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel2.Controls, Me.TxtReadOnly.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gboItem.Controls, Me.TxtReadOnly.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel1.Controls, Me.TxtReadOnly.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel2.Controls, Me.TxtReadOnly.BackColor, False)
            Me.LblUnitPrice.Enabled = False
            Me.RboAddressShipTo.Enabled = False
            Me.RboCustomerShipTo.Enabled = False
            Me.ColDeliveryHistory.Visible = True

            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiSO_Baru") = False Then
                Me.btnNew.Enabled = False
            End If
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiSO_Edit") = False Then
                Me.btnEdit.Enabled = False
            End If
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiSO_Delete") = False Then
                Me.btnDelete.Enabled = False
            End If
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiSO_Print") = False Then
                Me.btnPrint.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            Me.KosongkanItem()
            ModCekData.KosongkanControl(Me.gboDetail.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.TabPage1.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.SplitContainer1.Panel2.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.SplitContainer4.Panel1.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.SplitContainer4.Panel2.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.TabPage2.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.SplitContainer3.Panel1.Controls, True, True, True)
            ModCekData.KosongkanControl(Me.SplitContainer3.Panel2.Controls, True, True, True)
            Me.TxtPrimarySO.Text = "-1"
            Me.txtPrimaryCurrency.Text = "-1"
            Me.txtPrimaryCustomer.Text = "-1"
            Me.TxtPrimarySalesP.Text = "-1"
            Me.txtPaymentDueDate.Text = "0"
            Me.txtSubTotal.Text = "0"
            Me.txtTax.Text = "0"
            Me.txtTaxAmount.Text = "0"
            Me.TxtTotal.Text = "0"
            Me.txtPaymentTerm.Text = "0"
            Me.TxtExchangeRate.Text = "1"
            Me.TxtExchangeRate.ReadOnly = True
            Me.TxtExchangeRate.TabStop = False
            Me.TxtExchangeRate.BackColor = Me.TxtReadOnly.BackColor
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "

            Me.dgvItem.Rows.Clear()
            Me.cboPaymentDueDateReference.SelectedIndex = 0
            Me.chkTax.Checked = True
            Me.cboCustomerAddress.DataSource = Nothing
            Me.cboCustomerAddress.Text = ""
            Me.cboCustomerAddress.Items.Clear()
            Me.CboAddressShipTo.Items.Clear()
            Me.CboAddressCustomerShipTo.Items.Clear()
            Me.ChkDiscountConditionSama.Checked = False
            Me.txtReferenceDiscount.Text = "0"
            Me.txtReferenceDiscount.ReadOnly = True
            Me.txtReferenceDiscount.BackColor = Me.TxtSalesOrderNo.BackColor
            Me.TxtDiscount.BackColor = Me.TxtPrimaryCodeNo.BackColor
            Me.TxtDiscount.ReadOnly = False
            Me.TxtDiscount.TabStop = True
            Me.CboPriceCondition.SelectedIndex = 0
            Me.CboPriceCondition.Enabled = False

            Me.vTBLItemDelivery.Rows.Clear()
            Me.TxtPrimaryCustomerShipTo.Text = "-1"

            Me.TxtDPFromPercent.Text = 1
            Me.chkUseDP.Checked = False
            Me.TxtDPPercent.Text = "0.00"
            Me.TxtDPAmount.Text = "0"
            Me.TxtDPPercent.ReadOnly = True
            Me.TxtDPPercent.TabStop = False
            Me.TxtDPPercent.BackColor = Me.TxtReadOnly.BackColor
            Me.TxtDPAmount.ReadOnly = True
            Me.TxtDPAmount.TabStop = False
            Me.TxtDPAmount.BackColor = Me.TxtReadOnly.BackColor

            RboAddressShipTo.Checked = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanItem()
        Dim DummyDate As Date = Today
        Dim CheckDummy As Boolean = False

        Try
            DummyDate = Me.dtpdeliverydate.Value.Date
            CheckDummy = Me.dtpdeliverydate.Checked

            ModCekData.KosongkanControl(Me.gboItem.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer2.Panel1.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer2.Panel2.Controls, True, True, False)
            Me.TxtPrimaryCodeNo.Text = "-1"
            Me.txtPrimaryUnitsDefault.Text = "-1"
            Me.txtPrimaryUnitsMulti.Text = "-1"
            Me.txtPrimaryCategory.Text = "-1"
            Me.TxtQtyperBatchDefault.Text = "0"
            Me.TxtQtyperBatchMulti.Text = "0"
            Me.TxtUnitPrice.Text = "0"
            Me.TxtDiscount.Text = "0"
            Me.TxtDiscFromPerct.Text = "1"
            Me.TxtDiscAmount.Text = "0"
            Me.txtNettPrice.Text = "0"
            Me.txtAmount.Text = "0"
            Me.lblUnitsDefault.Text = "Units"
            Me.LblUnitPrice.Text = "Unit Price"
            Me.cboPriceConditionItem.SelectedIndex = 0
            If Me.ChkDiscountConditionSama.Checked = True Then
                Me.TxtDiscount.Text = Me.txtReferenceDiscount.Text
            End If
            If Me.ChkPriceCondition.Checked = True Then
                Me.cboPriceConditionItem.SelectedIndex = Me.CboPriceCondition.SelectedIndex
            End If
            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiSO_Override_Promise_To_Deliver") = False Then
                Me.dtpdeliverydate.Checked = False
                Me.dtpdeliverydate.Value = Me.dtpDate.Value
            Else
                Me.dtpdeliverydate.Checked = CheckDummy
                If Me.dtpdeliverydate.Checked = True Then
                    Me.dtpdeliverydate.Value = DummyDate
                End If
            End If

            Me.cboPriceByMultiUnit.Visible = False
            Me.cboPriceByMultiUnit.TabStop = False
            Me.txtPriceByMultiUnit.Text = "1"
            Me.cboPriceByMultiUnit.Items.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataItemBenar() As Boolean
        Dim ObjSOCustom As New CustomSMESS.ClsTransaksiSO

        Try
            'data item salah bila:
            '1. product code yang diisikan salah
            If Me.txtDesc.Text = "" Then
                MsgBox("Please enter code no.", MsgBoxStyle.Information, "Information")
                If Me.chkUseCustCodeNo.Checked = False Then
                    Me.TxtCodeNo.Text = ""
                    Me.TxtCodeNo.Focus()
                Else
                    Me.txtCustomerProductCodeNo.Text = ""
                    Me.txtCustomerProductCodeNo.Focus()
                End If
                Return False
            End If
            '2. nilai qty <= 0
            If Me.ObjCD.CekBilanganPositif(Me.TxtQtyperBatchMulti.Text) = False Then
                MsgBox("Please enter a positive number for quantity data.", MsgBoxStyle.Information, _
                "Information")
                Me.TxtQtyperBatchDefault.Focus()
                Return False
            End If
            '2.b. nilai qty <= 0
            If Me.ObjCD.CekBilanganPositif(Me.TxtQtyperBatchDefault.Text) = False Then
                MsgBox("Qty you entered make default qty=0.", MsgBoxStyle.Information, "Information")
                Me.TxtQtyperBatchMulti.Focus()
                Return False
            End If
            '3. Bila sudah ada product pada dgv dan unit price, discount, nett price-nya, 
            'delivery datenya beda - saat ini satu product hanya bisa muncul satu kali pada 1 so
            If Me.GetIndex(Me.TxtPrimaryCodeNo.Text) <> -1 Then
                If Me.cboPriceByMultiUnit.Visible = False Then
                    If Me.GetIndex(Me.TxtPrimaryCodeNo.Text, _
                                 Me.TxtUnitPrice.Text, Me.txtNettPrice.Text, Me.TxtDiscount.Text, Me.txtPrimaryUnitsMulti.Text, "/" & Me.lblUnitsMulti.Text) = -1 Then
                        MsgBox(Me.TxtCodeNo.Text + " is already enterred, double product code no in one sales order is not supported, use different sales order instead.", _
                        MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                Else
                    If Me.GetIndex(Me.TxtPrimaryCodeNo.Text, _
                    Me.TxtUnitPrice.Text, Me.txtNettPrice.Text, Me.TxtDiscount.Text, Me.txtPrimaryUnitsMulti.Text, "/" & Me.cboPriceByMultiUnit.Text) = -1 Then
                        MsgBox(Me.TxtCodeNo.Text + " is already enterred, double product code no in one sales order is not supported, use different sales order instead.", _
                        MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
             
            End If
            '4. bila janji akan ada tanggal pengiriman dan kemudian tanggal pengirimannya dibelakang hari ini
            '   berikan peringatan
            If Me.dtpdeliverydate.Checked = True And Me.dtpdeliverydate.Value.Date < Me.dtpDate.Value.Date Then
                If MsgBox("Delivery date is smaller than transaction date, continue?", MsgBoxStyle.YesNo, _
                "Confirm") = MsgBoxResult.No Then
                    Return False
                End If
            End If
            '5. bila harga yang diberikan = 0
            If Me.ObjCD.CekBilanganPositif(Me.TxtUnitPrice.Text) = False Then
                MsgBox("Please enter a positive number for unit price data.", MsgBoxStyle.Information, _
                "Information")
                Me.TxtUnitPrice.Focus()
                Return False
            End If

            'cek batasan custom
            If ObjSOCustom.ApakahDataBenarItem(Me.TxtPrimaryCodeNo.Text, Me.txtPriceByMultiUnit.Text, _
            Me.txtPrimaryUnitsMulti.Text, Me.txtPrimaryUnitsDefault.Text) = False Then
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjSOCustom = Nothing
        End Try
    End Function

    Private Function GetJmlDGV(ByVal vPrimaryProduct As Double, _
    ByVal UnitPrice As Double, ByVal NettPrice As Double, ByVal Discount As Double) As Double
        Try
            Dim row As DataGridViewRow

            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColDeliveryCont.Index).Value = Boolean.TrueString Then
                    If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And row.Cells(Me.ColDeliveryCont.Index).Value = Boolean.TrueString And _
                    row.Cells(Me.ColUnitPrice.Index).Value = UnitPrice And _
                    row.Cells(Me.ColDiscount.Index).Value = Discount And row.Cells(Me.ColNettPrice.Index).Value = NettPrice Then
                        Return row.Cells(Me.ColQtyMulti.Index).Value
                    End If
                Else
                    If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And row.Cells(Me.ColDeliveryCont.Index).Value = Boolean.FalseString And _
                    row.Cells(Me.ColDiscount.Index).Value = Discount And row.Cells(Me.ColNettPrice.Index).Value = NettPrice Then
                        Return row.Cells(Me.ColQtyMulti.Index).Value
                    End If
                End If
            Next

            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function ApakahDataBenar() As Boolean
        Dim KoleksiNamaData As New Collection
        Dim KoleksiTextBox As New Collection
        Dim PrimaryDummy As Double = -1
        Dim Dummy As String = ""
        Dim CreditLimit As Decimal
        Dim vFrmReminder As New FrmTransaksiSO_ReminderAR
        Dim DummyCodeNo As String = ""
        Dim ListDummy As New List(Of String)
        Dim Tolerance As Decimal = 0
        Dim ListIdentifier As New List(Of String)
        Dim rowDGV1 As DataGridViewRow
        Dim Qty, QtyDelivery, JmlDPBalance, JmlDP As Decimal
        Dim CodeNo As String = ""
        Dim ObjCustom As New CustomSMESS.ClsTransaksiSO
        Dim ObjBalanceDP As New Core_Module.ClsTransaksiBalanceDP
        Dim PrimaryCust As Double

        Try
            'If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Me.dtpDate.Value Then
            '    MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
            '    MsgBoxStyle.Information, "Information")
            '    Return False
            'End If

            If ObjCustom.ApakahDataBenarCustom(Me.dgvItem, Me.ColCodeNo.Index, Me.ColPrimaryCodeNo.Index, Me.chkTax.Checked, _
            Me.ColPrimaryUnits.Index, Me.ColPrimaryUnitsMulti.Index, Me.ColPrimaryCodeNo.Index, Me.ColPriceByMultiUnit.Index) = False Then
                Return False
            End If

            'cek data yang harus diisi
            KoleksiNamaData.Add("Sales order no")
            KoleksiTextBox.Add(Me.TxtSalesOrderNo)
            KoleksiNamaData.Add("Customer code")
            KoleksiTextBox.Add(Me.txtCustomerCode)
            KoleksiNamaData.Add("Currency")
            KoleksiTextBox.Add(Me.txtCurrencyCode)
            KoleksiNamaData.Add("Sales person code")
            KoleksiTextBox.Add(Me.TxtSalesPCode)

            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            If Me.txtPaymentTerm.Text = "" Then
                Me.txtPaymentTerm.Text = "0"
            End If

            'cek data no transaksi yang salah - tidak perlu cek karena ini adalah autonumber

            'minimal harus punya 1 data item
            If Me.dgvItem.RowCount = 0 Then
                MsgBox("Please enter at least 1 item.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'kalo sampe cek berkaitan dengan timestamp apakah masih benar
            If Me.baru = False Then
                If Me.objSO.TimeStamp <> Me.objSO.GetTimestamp(Me.objSO.Primary) Then
                    If MsgBox("Other user have modified this data, reload latest data?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.objSO.CariDetail(Me.objSO.Primary, Me.dgvItem, Me.vTBLItemDelivery)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            '8. Jika Ship to Customer dipilih, maka customer ship to wajib diisi
            If Me.RboCustomerShipTo.Checked And (Me.TxtPrimaryCustomerShipTo.Text = "" Or Me.TxtPrimaryCustomerShipTo.Text = "-1" Or Me.TxtIdentifierAddressCustomerShipTo.Text = "") Then
                MsgBox("If you want to ship to other customer, you have to define a customer and address.", MsgBoxStyle.Information, "Information")
                Me.TxtCustomerCodeShipTo.Focus()
                Return False
            End If

            If Me.RboCustomerShipTo.Checked And (Me.txtPrimaryCustomer.Text = Me.TxtPrimaryCustomerShipTo.Text) Then
                MsgBox("If you want to ship to other customer, please define different customer for goods shipping.", MsgBoxStyle.Information, "Information")
                Me.TxtCustomerCodeShipTo.Focus()
                Return False
            End If

            If Me.txtPrimaryCurrency.Text = Core_Module.PrimaryBaseCurrency Then
                'jika IDR, maka exchange rate =1
                If CDec(Me.TxtExchangeRate.Text) <> 1 Then
                    Me.TxtExchangeRate.Text = "1"
                End If
            Else
                'jika non-IDR, maka exchange rate <> 1
                If CDec(Me.TxtExchangeRate.Text) = 1 Then
                    MsgBox("Please specify exchange rate to " + Me.ObjCurrency.GetCurrencyCode(Core_Module.PrimaryBaseCurrency) + ".", _
                    MsgBoxStyle.Information, "Information")
                    Me.TabControl1.SelectedIndex = 1
                    Me.TxtExchangeRate.Focus()

                    Return False
                End If
            End If

            'cek berkaitan dengan perubahan tanggal
            Me.TxtSalesOrderNo.Text = Me.TxtSalesOrderNo.Text.Trim
            If Me.ObjNumbering.Apakah_Autonumber(ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.SO) = True Then
                If Me.baru = False Then
                    If Me.objSO.TanggalTransaksi.Month = Me.dtpDate.Value.Month And _
                    Me.objSO.TanggalTransaksi.Year = Me.dtpDate.Value.Year Then
                        Me.TxtSalesOrderNo.Text = Me.objSO.NoTransaksi
                    Else
                        If MsgBox("You just changed transaction date, this will change your transaction number and there's a possibility your transaction no cannot be change back. Continue and accept this change?", _
                        MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
                            Me.dtpDate.Value = Me.objSO.TanggalTransaksi
                            Me.TxtSalesOrderNo.Text = Me.objSO.NoTransaksi
                            Return False
                        End If
                    End If
                End If
            Else
                'cek data double
                If Me.baru = True Then
                    If Me.objSO.ApakahSudahAda(Me.TxtSalesOrderNo.Text) = True Then
                        MsgBox("Sales order number duplication, please specify new sales order number.", MsgBoxStyle.Information, "Information")
                        Me.TxtSalesOrderNo.Focus()
                        Return False
                    End If
                Else
                    If Me.objSO.ApakahSudahAda(Me.TxtSalesOrderNo.Text) = True Then
                        If Me.objSO.GetPrimary(Me.TxtSalesOrderNo.Text) <> Me.objSO.Primary Then
                            MsgBox("Sales order number duplication, please specify new sales order number.", MsgBoxStyle.Information, "Information")
                            Me.TxtSalesOrderNo.Focus()
                            Return False
                        End If
                    End If
                End If
            End If

            'kalo sampe sini return true
            If Me.txtPaymentDueDate.Text = "0" Then
                If MsgBox("The payment term cash, continue?", MsgBoxStyle.YesNo, "Confirmation") <> MsgBoxResult.Yes Then
                    Me.TabControl1.SelectedIndex = 1
                    Me.txtPaymentTerm.Focus()
                    Return False
                End If
            End If

            'nilai maksimal total sales adalah: 999,999,999,999
            If CDec(Me.TxtTotal.Text) > 999999999999 Then
                MsgBox("Maximum sales value is :" + Me.ObjCD.DigitGrouping(999999999999, Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text)) + ". You cannot save this data.", MsgBoxStyle.Information, "Information")
                Return False
            End If
            'nilai pajak maksimal adalah 999,999,999,999
            If CDec(Me.txtTaxAmount.Text) > 999999999999 Then
                MsgBox("Maximum VAT value is :" + Me.ObjCD.DigitGrouping(999999999999, Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text)) + ". You cannot save this data.", MsgBoxStyle.Information, "Information")
                Return False
            End If



            Me.txtCustomerPONo.Text = Me.txtCustomerPONo.Text.Trim
            If Me.cboOrderBy.Text = "By PO" Or Me.cboOrderBy.SelectedIndex = 0 Then
                If Me.txtCustomerPONo.Text = "" Then
                    'jika po customer tidak diisikan maka confirm
                    If MsgBox("You didn't enter customer's PO number, continue?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                        Me.TabControl1.SelectedIndex = 0
                        Me.txtCustomerPONo.Focus()
                        Return False
                    End If
                Else
                    'jika ada po customer yang sama yang dimasukkan pada so, maka tanyakan ke user apakah mau input 2 kali.
                    PrimaryDummy = Me.objSO.GetPrimaryBy_CustPONumber(Me.txtCustomerPONo.Text)
                    If PrimaryDummy <> CDbl(Me.TxtPrimarySO.Text) Then
                        If MsgBox("The sales order:" + Me.objSO.GetNoSO(PrimaryDummy) + " also refer to customer PO number :" _
                        + Me.txtCustomerPONo.Text + ", do you want to continue?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                            Me.TabControl1.SelectedIndex = 0
                            Me.txtCustomerPONo.Focus()
                            Return False
                        End If
                    End If
                End If
            End If

            'cek exceed credit limit
            If Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.SO, _
            ClsTransaksiSetting.EnumNamaParameter.Reminder_AccountReceivable_ExceedsCreditLimit) = "1" Then
                CreditLimit = Me.ObjCustomer.GetCreditLimit(Me.txtPrimaryCustomer.Text)
                If CreditLimit <> 0 Then
                    If CreditLimit < Me.objSO.GetAR(Me.txtPrimaryCustomer.Text) Then
                        vFrmReminder.dgvItem.Rows.Clear()
                        vFrmReminder.vPrimaryCustomer = Me.txtPrimaryCustomer.Text
                        vFrmReminder.vTipeReminder = FrmTransaksiSO_ReminderAR.EnumTipeReminder.AR_Exceeds_Limit
                        vFrmReminder.vDate = Me.dtpDate.Value
                        vFrmReminder.vCreditLimit = CreditLimit

                        If vFrmReminder.ShowDialog <> Windows.Forms.DialogResult.Yes Then
                            Return False
                        End If
                    End If
                End If
            End If

            'over due ar
            If Me.ObjSetting.GetSettingValue(ClsTransaksiSetting.EnumTipeTransaksi.SO, _
            ClsTransaksiSetting.EnumNamaParameter.Reminder_AccountReceivable_Overdue) = "1" Then
                Dummy = Me.objSO.GetAR_Detail_OverDue(Me.txtPrimaryCustomer.Text, Me.dtpDate.Value)
                If Dummy <> "" Then
                    vFrmReminder.dgvItem.Rows.Clear()
                    vFrmReminder.vPrimaryCustomer = Me.txtPrimaryCustomer.Text
                    vFrmReminder.vTipeReminder = FrmTransaksiSO_ReminderAR.EnumTipeReminder.AR_Overdue
                    vFrmReminder.vDate = Me.dtpDate.Value

                    If vFrmReminder.ShowDialog <> Windows.Forms.DialogResult.Yes Then
                        Return False
                    End If
                End If
            End If

            For Each Me.vrowItemDelivery In Me.vTBLItemDelivery.Rows
                If Me.vrowItemDelivery.delivery_date < Me.dtpDate.Value Then
                    If ListDummy.Contains(Me.vrowItemDelivery.identifier) = False Then
                        ListDummy.Add(Me.vrowItemDelivery.identifier)
                    End If
                End If
            Next

            If ListDummy.Count > 0 Then
                For Each Iden_Dummy As String In ListDummy
                    For Each rowdgv As DataGridViewRow In Me.dgvItem.Rows
                        If rowdgv.Cells(Me.ColIdentifier.Index).Value = Iden_Dummy Then
                            If DummyCodeNo = "" Then
                                DummyCodeNo = rowdgv.Cells(Me.ColCodeNo.Index).Value
                            Else
                                DummyCodeNo = DummyCodeNo + ", " + rowdgv.Cells(Me.ColCodeNo.Index).Value
                            End If
                            Exit For
                        End If
                    Next
                Next

                MsgBox("The following code no have delivery date < transaction date. Please fix first." + _
                vbCrLf + DummyCodeNo, MsgBoxStyle.Information, "Information")

                Return False
            End If

            ''cek bisi ada qty default=0
            Dim ListError As String = ""
            For Each RowDgv As DataGridViewRow In Me.dgvItem.Rows
                With RowDgv
                    If CDec(.Cells(Me.ColQty.Index).Value) <= 0 Then
                        ListError += .Cells(Me.ColCodeNo.Index).Value & ", "
                    End If
                End With
            Next

            If ListError <> "" Then
                ListError = ListError.Substring(0, ListError.Length - 3) & "."
                MsgBox("The following product has zero default qty: " & vbNewLine & ListError & vbNewLine & "Please change qty or set ratio.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'pastikan setiap item delivery, qty-nya (multi dan default) konsisten dengan item po
            ListIdentifier.Clear()
            For Each rowDGV1 In Me.dgvItem.Rows
                If rowDGV1.Cells(Me.ColDeliveryCont.Index).Value = Boolean.TrueString Then
                    If ListIdentifier.Contains(rowDGV1.Cells(Me.ColIdentifier.Index).Value) = False Then
                        ListIdentifier.Add(rowDGV1.Cells(Me.ColIdentifier.Index).Value)
                    End If
                End If
            Next
            For Each Me.vrowItemDelivery In Me.vTBLItemDelivery.Rows
                If ListIdentifier.Contains(Me.vrowItemDelivery.identifier) = False Then
                    ListIdentifier.Contains(Me.vrowItemDelivery.identifier)
                End If
            Next
            For Each Identifier As String In ListIdentifier
                'cek qty default
                Qty = Me.GetJmlDGV_Default(Identifier)
                QtyDelivery = Me.GetJmlItemDelivery_Default(Identifier)
                If Qty <> QtyDelivery Then
                    CodeNo = Me.GetCodeNo(Identifier)
                    If CodeNo <> "" Then
                        MsgBox("Qty delivery for: " + CodeNo + " is not consistent, please check.", MsgBoxStyle.Information, "Information")
                    Else
                        MsgBox("Logic error: inconsistent delivery qty, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                    End If

                    Return False
                End If

                'cek qty multi
                Qty = Me.GetJmlDGV_Multi(Identifier)
                QtyDelivery = Me.GetJmlItemDelivery_Multi(Identifier)
                If Qty <> QtyDelivery Then
                    CodeNo = Me.GetCodeNo(Identifier)
                    If CodeNo <> "" Then
                        MsgBox("Qty delivery for: " + CodeNo + " is not consistent, please check.", MsgBoxStyle.Information, "Information")
                    Else
                        MsgBox("Logic error: inconsistent delivery qty, please contact your vendor if you see this message.", MsgBoxStyle.Information, "Information")
                    End If

                    Return False
                End If
            Next

            'cek ke balance dp
            If Me.baru = False Then
                JmlDPBalance = ObjBalanceDP.GetJml(ClsTransaksiBalanceDP.EnumTipeDP.DP_Invoice, Me.TxtPrimarySO.Text)
                JmlDP = Me.objInv.GetDPAmount(Me.TxtPrimarySO.Text)

                If JmlDPBalance > 0 Then
                    'primary cust harus cocok
                    PrimaryCust = Me.objSO.GetPrimaryCustomer(Me.TxtPrimarySO.Text)
                    If PrimaryCust <> Me.txtPrimaryCustomer.Text Then
                        Dummy = ObjCustomer.GetCustomerCode(PrimaryCust)
                        Dummy = Dummy + " - " + ObjCustomer.GetCustName(PrimaryCust)

                        MsgBox("This data is refered by down payment begining balance data, you can't change customer code. Please use customer: " + Dummy, _
                        MsgBoxStyle.Information, "Information")
                        Me.txtCustomerCode.Focus()
                        Return False
                    End If

                    'primary currency harus cocok
                    PrimaryDummy = Me.objSO.GetPrimaryCurrency(Me.TxtPrimarySO.Text)
                    If Me.txtPrimaryCurrency.Text <> PrimaryDummy Then
                        Dummy = ObjCurrency.GetCurrencyCode(PrimaryDummy)
                        MsgBox("This data is refered by down payment begining balance data, you can't change currency data. Please use currency: " + Dummy, _
                        MsgBoxStyle.Information, "Information")
                        Me.txtCustomerCode.Focus()
                        Return False
                    End If
                End If
                If JmlDP > 0 Then
                    'primary cust harus cocok
                    PrimaryCust = Me.objSO.GetPrimaryCustomer(Me.TxtPrimarySO.Text)
                    If PrimaryCust <> Me.txtPrimaryCustomer.Text Then
                        Dummy = ObjCustomer.GetCustomerCode(PrimaryCust)
                        Dummy = Dummy + " - " + ObjCustomer.GetCustName(PrimaryCust)

                        MsgBox("This data is refered by down payment invoice data, you can't change customer code. Please use customer: " + Dummy, _
                        MsgBoxStyle.Information, "Information")
                        Me.txtCustomerCode.Focus()
                        Return False
                    End If

                    'primary currency harus cocok
                    PrimaryDummy = Me.objSO.GetPrimaryCurrency(Me.TxtPrimarySO.Text)
                    If Me.txtPrimaryCurrency.Text <> PrimaryDummy Then
                        Dummy = ObjCurrency.GetCurrencyCode(PrimaryDummy)
                        MsgBox("This data is refered by down payment invoice data, you can't change currency data. Please use currency: " + Dummy, _
                        MsgBoxStyle.Information, "Information")
                        Me.txtCustomerCode.Focus()
                        Return False
                    End If
                End If

                'minimal penjualan harus >= jml dp balance + jml dp normal
                If JmlDP + JmlDPBalance > CDec(Me.txtSubTotal.Text) Then
                    MsgBox("This sales order is refered by down payment invoice, minimum transaction amount is: " + Me.txtCurrencyCode.Text + " " + _
                    Me.ObjCD.DigitGrouping(JmlDP + JmlDPBalance, Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text)), _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            If CDec(Me.txtSubTotal.Text) < CDec(Me.TxtDPAmount.Text) Then
                MsgBox("Maximum DP amount is " + Me.txtCurrencyCode.Text + " " + Me.txtSubTotal.Text + ".", MsgBoxStyle.Information, "Information")
                Return False
            End If
            If Me.chkUseDP.Checked = True And CDec(Me.TxtDPAmount.Text) = 0 Then
                MsgBox("Please specify DP amount.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            '\\TR#
            Dim msg_ As String = ""
            For Each rowDgv_ As DataGridViewRow In Me.dgvItem.Rows
                With rowDgv_
                    If CDec(.Cells(Me.ColToleranceDown.Index).Value) > (CDec(.Cells(Me.ColQtyMulti.Index).Value) / 2) Or _
                            CDec(.Cells(Me.ColToleranceUp.Index).Value) > (CDec(.Cells(Me.ColQtyMulti.Index).Value)) Or _
                            CDec(.Cells(Me.ColToleranceDownPerc.Index).Value) > 50 Or _
                            CDec(.Cells(Me.ColToleranceUpPerc.Index).Value) > 100 Then
                        msg_ += .Cells(Me.ColCodeNo.Index).Value & ","
                    End If
                End With
            Next
            If String.IsNullOrEmpty(msg_) = False Then
                msg_ = msg_.Remove(msg_.Length - 1, 1)
                MsgBox("Tolerance value of the following product code cannot greater than (down:50% & up:100%): " & vbNewLine & msg_, MsgBoxStyle.Information, "Information")
                Return False
            End If

            'cek primary customer, jika sudah ada do, harus konsisten
            If Me.ObjDOItem.GetJmlBySO(Me.TxtPrimarySO.Text) > 0 Then
                Dim PrimaryCustRef As Double = Me.objSO.GetPrimaryCustomer(Me.TxtPrimarySO.Text)
                If Me.txtPrimaryCustomer.Text <> PrimaryCustRef Then
                    MsgBox("Please use customer: " + Me.ObjCustomer.GetCustomerCode(PrimaryCustRef) + " - " + Me.ObjCustomer.GetCustName(PrimaryCustRef) + _
                    ", this data is referenced by delivery order for the specified customer.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            If Me.chkTax.Checked = True Then
                If CDec(Me.txtTax.Text) <= 0 Then
                    Me.TabControl1.SelectedIndex = 1
                    Me.txtTax.Focus()
                    MsgBox("Please specify tax percentage.", MsgBoxStyle.Information, "Information")
                    Return False
                End If

                Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.TaxPercentage)
                If Dummy = "" Then
                    Dummy = "0"
                End If

                If CDec(Me.txtTax.Text) <> CDec(Dummy) Then
                    Me.TabControl1.SelectedIndex = 1
                    Me.txtTax.Focus()
                    If MsgBox("Tax rate <> " + Me.ObjCD.DigitGrouping(Dummy, 2) + "%, continue to save data?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                        Return False
                    End If
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            KoleksiNamaData = Nothing
            KoleksiTextBox = Nothing
            vFrmReminder.Dispose()
            ListIdentifier = Nothing
            ObjBalanceDP = Nothing
        End Try
    End Function

    Private Function GetCodeNo(ByVal vIdentifier As String) As String
        Dim rowDGV As DataGridViewRow

        Try
            For Each rowDGV In Me.dgvItem.Rows
                If rowDGV.Cells(Me.ColIdentifier.Index).Value = vIdentifier Then
                    Return rowDGV.Cells(Me.ColCodeNo.Index).Value
                End If
            Next
            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function GetJmlDGV_Default(ByVal vIdentifier As String) As Decimal
        Dim rowDGV As DataGridViewRow
        Dim RetVal As Decimal = 0

        Try
            For Each rowDGV In Me.dgvItem.Rows
                If rowDGV.Cells(Me.ColIdentifier.Index).Value = vIdentifier Then
                    RetVal += CDec(rowDGV.Cells(Me.ColQty.Index).Value)
                End If
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlDGV_Multi(ByVal vIdentifier As String) As Decimal
        Dim rowDGV As DataGridViewRow
        Dim RetVal As Decimal = 0

        Try
            For Each rowDGV In Me.dgvItem.Rows
                If rowDGV.Cells(Me.ColIdentifier.Index).Value = vIdentifier Then
                    RetVal += CDec(rowDGV.Cells(Me.ColQtyMulti.Index).Value)
                End If
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlItemDelivery_Default(ByVal vIdentifier As String) As Decimal
        Dim DV As DataRowView
        Dim RetVal As Decimal = 0

        Try
            Me.vTBLItemDelivery.DefaultView.RowFilter = "identifier = '" + vIdentifier + "'"
            For Each DV In Me.vTBLItemDelivery.DefaultView
                Me.vrowItemDelivery = DV.Row
                RetVal += Me.vrowItemDelivery.delivery_qty
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlItemDelivery_Multi(ByVal vIdentifier As String) As Decimal
        Dim DV As DataRowView
        Dim RetVal As Decimal = 0

        Try
            Me.vTBLItemDelivery.DefaultView.RowFilter = "identifier = '" + vIdentifier + "'"
            For Each DV In Me.vTBLItemDelivery.DefaultView
                Me.vrowItemDelivery = DV.Row
                RetVal += Me.vrowItemDelivery.delivery_qty_multi
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub SalinKeObjek()
        Try
            'salin data header
            With Me.objSO
                .CustContactPerson = Me.txtContactPerson.Text
                .CustPONo = Me.txtCustomerPONo.Text
                .CustPhone = Me.txtPhoneNo.Text
                .CustIdentifierAdd = Me.txtCustomerAddIdentifier.Text
                .NoTransaksi = Me.TxtSalesOrderNo.Text
                .PaymentDue = Me.txtPaymentDueDate.Text
                .PaymentTerm = Me.txtPaymentTerm.Text
                .PrimaryCurrency = Me.txtPrimaryCurrency.Text
                .PrimaryCustomer = Me.txtPrimaryCustomer.Text
                .RefPaymentDate = Me.cboPaymentDueDateReference.SelectedIndex
                .SubTotal = Me.txtSubTotal.Text
                .TaxAmount = Me.txtTaxAmount.Text
                .TaxPercentage = Me.txtTax.Text
                .Total = Me.TxtTotal.Text
                .UseTax = Me.chkTax.Checked
                .Remarks = Me.txtRemarks.Text
                .TanggalTransaksi = Me.dtpDate.Value
                .ShipVia = Me.TxtShipVia.Text
                .OrderBy = Me.cboOrderBy.SelectedIndex
                .JenisFakturPajak = Me.CboFakturPajak.SelectedIndex
                .PriceCondition = Me.CboPriceCondition.SelectedIndex
                .PriceConditionSama = Me.ChkPriceCondition.Checked
                .DiskonSama = Me.ChkDiscountConditionSama.Checked
                .ReferenceDiscount = Me.txtReferenceDiscount.Text

                .PrimarySalesPerson = Me.TxtPrimarySalesP.Text
                .ExchangeRate = Me.TxtExchangeRate.Text

                .CustPODate = Me.dtpPOdate.Value
                .IdentifierAddressShipTo = Me.TxtIdentifierAddressShipTo.Text
                .PrimaryCustomerShipTo = Me.TxtPrimaryCustomerShipTo.Text
                .IdentifierAddressCustomerShipTo = Me.TxtIdentifierAddressCustomerShipTo.Text

                If Me.chkUseDP.Checked = True Then
                    .NeedDP = 1
                Else
                    .NeedDP = 0
                End If
                .DPFromPercent = Me.TxtDPFromPercent.Text
                .DPPercent = Me.TxtDPPercent.Text
                .DPAmount = Me.TxtDPAmount.Text
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SalinKeField()
        Try
            With Me.objSO
                Me.txtCustomerPONo.Text = .CustPONo
                Me.TxtShipVia.Text = .ShipVia
                Me.cboOrderBy.SelectedIndex = .OrderBy
                Me.CboFakturPajak.SelectedIndex = .JenisFakturPajak
                Me.CboPriceCondition.SelectedIndex = .PriceCondition
                Me.ChkPriceCondition.Checked = .PriceConditionSama
                Me.ChkDiscountConditionSama.Checked = .DiskonSama
                Me.txtReferenceDiscount.Text = .ReferenceDiscount
                Me.txtContactPerson.Text = .CustContactPerson
                Me.txtPhoneNo.Text = .CustPhone
                Me.TxtPrimarySO.Text = .Primary
                Me.TxtSalesOrderNo.Text = .NoTransaksi
                Me.txtPaymentDueDate.Text = .PaymentDue
                Me.txtPaymentTerm.Text = .PaymentTerm
                Me.txtPrimaryCurrency.Text = .PrimaryCurrency
                Me.txtCurrencyCode.Text = Me.ObjCurrency.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                Me.DefaultPrecision = Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text)
                Me.txtPrimaryCustomer.Text = .PrimaryCustomer
                Me.txtCustomerAddIdentifier.Text = .CustIdentifierAdd
                Me.txtCustomerCode.Text = Me.ObjCustomer.GetCustomerCode(Me.txtPrimaryCustomer.Text)
                Me.txtCustomerName.Text = Me.ObjCustomer.GetCustName(Me.txtPrimaryCustomer.Text)
                Me.ObjCustomer.IsiCboAddress(Me.cboCustomerAddress, Me.txtPrimaryCustomer.Text)
                Me.cboCustomerAddress.Text = Me.ObjCustomer.GetKodeAddress(Me.txtCustomerAddIdentifier.Text)
                Me.cboPaymentDueDateReference.SelectedIndex = .RefPaymentDate
                Me.txtSubTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.SubTotal, Me.DefaultPrecision))
                Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.TaxAmount, Me.DefaultPrecision))
                Me.txtTax.Text = .TaxPercentage
                Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.Total, Me.DefaultPrecision))
                Me.chkTax.Checked = .UseTax
                Me.txtRemarks.Text = .Remarks
                Me.dtpDate.Value = .TanggalTransaksi

                Me.TxtExchangeRate.Text = Me.ObjCD.DigitGrouping(.ExchangeRate, Core_Module.DefPrecExchangeRate)
                Me.TxtPrimarySalesP.Text = .PrimarySalesPerson
                Me.TxtSalesPCode.Text = Me.ObjSalesP.GetCode(.PrimarySalesPerson)
                Me.TxtSalesPName.Text = Me.ObjSalesP.GetName(.PrimarySalesPerson)

                If .NeedDP = 1 Then
                    Me.chkUseDP.Checked = True
                Else
                    Me.chkUseDP.Checked = False
                End If
                Me.TxtDPFromPercent.Text = .DPFromPercent
                Me.TxtDPPercent.Text = Me.ObjCD.DigitGrouping(.DPPercent, 2)
                Me.TxtDPAmount.Text = Me.ObjCD.DigitGrouping(.DPAmount, Me.DefaultPrecision)

                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString

                Me.dtpPOdate.Value = .CustPODate

                If .PrimaryCustomerShipTo = -1 Then
                    Me.RboAddressShipTo.Checked = True
                    Me.TxtIdentifierAddressShipTo.Text = .IdentifierAddressShipTo
                    Me.ObjCustomer.IsiCboAddress(Me.CboAddressShipTo, Me.txtPrimaryCustomer.Text)
                    Me.CboAddressShipTo.Text = Me.ObjCustomer.GetKodeAddress(Me.TxtIdentifierAddressShipTo.Text)
                Else
                    Me.RboCustomerShipTo.Checked = True
                    Me.TxtPrimaryCustomerShipTo.Text = .PrimaryCustomerShipTo
                    Me.TxtCustomerCodeShipTo.Text = Me.ObjCustomer.GetCustomerCode(Me.TxtPrimaryCustomerShipTo.Text)
                    Me.TxtCustomerDescShipTo.Text = Me.ObjCustomer.GetCustName(Me.TxtPrimaryCustomerShipTo.Text)
                    Me.TxtIdentifierAddressCustomerShipTo.Text = .IdentifierAddressCustomerShipTo
                    Me.ObjCustomer.IsiCboAddress(Me.CboAddressCustomerShipTo, Me.TxtPrimaryCustomerShipTo.Text)
                    Me.CboAddressCustomerShipTo.Text = Me.ObjCustomer.GetKodeAddress(Me.TxtIdentifierAddressCustomerShipTo.Text)
                End If

                Me.FormatDGV()
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub frmMasterTipe_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.TxtSalesOrderNo.MaxLength = 100
            Me.txtRemarks.MaxLength = 150

            Me.baru = True
            Me.KosongkanItem()
            Me.NonAktifkanDetail()
            UpdateKolomPO()
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll
            Control.CheckForIllegalCrossThreadCalls = False
            Me.FormBorderStyle = Windows.Forms.FormBorderStyle.Sizable

            Me.llbDetailStock.Left = Me.SplitContainer2.Panel1.Width - 5 - Me.llbDetailStock.Width
            Me.BtnFindCodeNo.Left = Me.llbDetailStock.Left - 5 - Me.BtnFindCodeNo.Width
            Me.TxtCodeNo.Width = Me.BtnFindCodeNo.Left - 5 - Me.TxtCodeNo.Left
            Me.txtDesc.Width = Me.SplitContainer2.Panel1.Width - 5 - Me.txtDesc.Left
            Me.llbUpdateAllDeliveryDate.Left = Me.SplitContainer2.Panel1.Width - 5 - Me.llbUpdateAllDeliveryDate.Width
            Me.dtpdeliverydate.Width = Me.llbUpdateAllDeliveryDate.Left - 5 - Me.dtpdeliverydate.Left
            Me.CboUnit.Left = Me.SplitContainer2.Panel1.Width - 5 - Me.CboUnit.Width
            Me.TxtQtyperBatchMulti.Width = Me.CboUnit.Left - 5 - Me.TxtQtyperBatchMulti.Left
            Me.TxtDiscount.Left = Me.SplitContainer2.Panel2.Width - 5 - Me.TxtDiscount.Width
            Me.Label10.Left = Me.TxtDiscount.Left - 5 - Me.Label10.Width
            Me.TxtUnitPrice.Width = Me.Label10.Left - 5 - Me.TxtUnitPrice.Left
            Me.txtNettPrice.Left = Me.TxtDiscount.Left
            Me.Label6.Left = Me.Label10.Left
            Me.TxtDiscAmount.Width = Me.TxtUnitPrice.Width
            Me.cboPriceConditionItem.Left = Me.TxtDiscount.Left
            Me.txtAmount.Width = Me.cboPriceConditionItem.Left - 5 - Me.txtAmount.Left
            Me.BtnInput.Left = Me.SplitContainer2.Panel1.Width - 5 - Me.BtnInput.Width
            Me.TxtRemarkItem.Width = Me.BtnInput.Left - 5 - Me.TxtRemarkItem.Left

            'Me.txtCustomerAddIdentifier.Visible = True
            'Me.TxtIdentifierAddressShipTo.Visible = True

            Me.CboUnit.Left = Me.llbUpdateAllDeliveryDate.Left
            Me.CboUnit.Width = Me.llbUpdateAllDeliveryDate.Width
            Me.TxtQtyperBatchMulti.Width = Me.dtpdeliverydate.Width
            'Me.txtPriceByMultiUnit.Visible = True
            'Me.ColPriceByMultiUnit.Visible = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ResetObjOtorisasi()
        Try
            Me.ObjOtorisasi.Primary = -1
            Me.ObjOtorisasi.PrimaryApprovedBy = -1
            Me.ObjOtorisasi.PrimaryDokumenAcuan = -1
            Me.ObjOtorisasi.PrimaryRequestBy = -1
            Me.ObjOtorisasi.PrimaryTipeApproval = 0
            Me.ObjOtorisasi.Reason = ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnSimpan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Dim FrmO As New Core_Module.FrmMasterUser_Otorisasi

        Try
            If Me.ApakahDataBenar() = True Then
                Me.KosongkanItem()
                Me.SalinKeObjek()
                If baru = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.objSO.SimpanBaru(Me.dgvItem, Me.UserCode, Me, Me.vTBLItemDelivery) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnFind.Focus()
                        End If
                    End If
                Else
                    If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiSO_Otorisasi_Edit") = False Then
                        If Me.ObjUser.GetJmlByNamaHakAkses("FrmTransaksiSO_Otorisasi_Edit") = 0 Then
                            MsgBox("You don't have permission to perform this action.", MsgBoxStyle.Information, "Information")
                            Exit Sub
                        Else
                            FrmO.AccessName = "FrmTransaksiSO_Otorisasi_Edit"
                            If FrmO.ShowDialog <> Windows.Forms.DialogResult.OK Then
                                MsgBox("You don't have permission to perform this action.", MsgBoxStyle.Information, "Information")
                                Exit Sub
                            End If
                        End If
                        'simpan data otorisasi untuk kemudian disimpan ke dalam database
                        With Me.ObjOtorisasi
                            .Primary = 1
                            .PrimaryApprovedBy = Me.ObjUser.GetPrimary(FrmO.cboUserCode.Text)
                            .PrimaryDokumenAcuan = Me.TxtPrimarySO.Text
                            .PrimaryRequestBy = Me.ObjUser.GetPrimary(Me.UserCode)
                            .PrimaryTipeApproval = ClsMasterTipeOtorisasi.EnumTipeOtorisasi.SO_Edit
                            .Reason = FrmO.TxtReason.Text
                        End With
                    End If

                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.objSO.SimpanPerubahan(Me.dgvItem, Me.UserCode, Me, Me.vTBLItemDelivery) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnEdit.Focus()
                            'simpan laporan otorisasi
                            Me.ObjOtorisasi.NoDocumentHist = Me.TxtSalesOrderNo.Text
                            Me.ObjOtorisasi.SimpanData()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
            FrmO.Dispose()
            Me.ResetObjOtorisasi()
        End Try
    End Sub

    Private Sub btnBatal_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.txtRemarks.Text = "" And Me.dgvItem.RowCount = 0 And _
            Me.TxtCodeNo.Text = "" And Me.TxtRemarkItem.Text = "" Then
                Me.NonAktifkanDetail()
                Me.btnNew.Focus()
                Me.KosongkanField()
            Else
                'konfirmasi batal atau tidak
                hasil = MsgBox("Cancel save and discard data changes?", MsgBoxStyle.YesNo, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    Me.KosongkanItem()
                    If baru = True Then
                        Me.KosongkanField()
                    Else
                        Me.objSO.CariDetail(Me.TxtPrimarySO.Text, Me.dgvItem, Me.vTBLItemDelivery)
                        Me.SalinKeField()
                    End If
                Else
                    Exit Sub
                End If
                Me.NonAktifkanDetail()
                Me.btnNew.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Dim FrmO As New Core_Module.FrmMasterUser_Otorisasi
        Dim UseTaxSetting As String = ""
        Dim TaxPercentSetting As String = ""

        Try
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Today Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            'pastikan ada default priority level
            If Me.ObjDeliveryPriority.GetPrimaryDefault = -1 Then
                MsgBox("Please specify default delivery priority level.", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiSO_Baru") = False Then
                FrmO.AccessName = "FrmTransaksiSO_Baru"
                If FrmO.ShowDialog <> Windows.Forms.DialogResult.OK Then
                    MsgBox("You don't have permission to perform this action.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            Me.baru = True
            Me.KosongkanField()
            Me.AktifkanDetail()
            Me.TabControl1.SelectedIndex = 0
            'Me.dtpDate.Focus()
            Me.objSO.ResetObject(Me.dgvItem)
            Me.TxtSalesOrderNo.Text = Me.objSO.GetNoBaru1(Me.dtpDate.Value)
            Me.txtPrimaryCurrency.Text = "1"
            Me.txtCurrencyCode.Text = Me.ObjCurrency.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
            Me.DefaultPrecision = Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text)

            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString

            'kalo new mata uang langsung IDR
            Me.txtPrimaryCurrency.Text = "1"
            Me.txtCurrencyCode.Text = Me.ObjCurrency.GetCurrencyCode(1)
            'Me.txtTax.Text = "10"
            'Me.chkTax.Checked = True
            Me.CboFakturPajak.SelectedIndex = 1
            Me.cboPaymentDueDateReference.SelectedIndex = 1

            UseTaxSetting = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.SO, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseTax)
            If UseTaxSetting = "" Then
                Me.chkTax.Checked = True
                Me.txtTax.Text = Me.ObjCD.DigitGrouping("10", 2)
            Else
                If UseTaxSetting = "0" Then
                    Me.chkTax.Checked = False
                Else
                    Me.chkTax.Checked = True
                    TaxPercentSetting = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.SO, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.TaxPercentage)
                    If IsNumeric(TaxPercentSetting) = True Then
                        Me.txtTax.Text = Me.ObjCD.DigitGrouping(TaxPercentSetting, 2)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FrmO = Nothing
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim hj As New FrmCariTransaksiSO
        Try
            hj.BackColorAll = Me.BackColorAll
            hj.ShowDialog()
            If Not hj.PrimaryTransaksi = -1 Then
                If Me.objSO.CariDetail(hj.PrimaryTransaksi, Me.dgvItem, Me.vTBLItemDelivery) _
                = True Then
                    Me.SalinKeField()
                    Me.btnEdit.Focus()
                Else
                    Me.btnFind.Focus()
                End If
                Me.btnFind.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Dim FrmO As New Core_Module.FrmMasterUser_Otorisasi

        Try
            If Me.TxtSalesOrderNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtSalesOrderNo.Text <> "" Then
                If Me.objSO.ApakahBolehEdit(Me.objSO.Primary, Me.UserCode, True) = True Then
                    If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiSO_Edit") = False Then
                        FrmO.AccessName = "FrmTransaksiSO_Edit"
                        If FrmO.ShowDialog <> Windows.Forms.DialogResult.OK Then
                            MsgBox("You don't have permission to perform this action.", MsgBoxStyle.Information, "Information")
                            Exit Sub
                        End If
                    End If

                    Me.baru = False
                    Me.objSO.CariDetail(Me.objSO.Primary, Me.dgvItem, Me.vTBLItemDelivery)
                    Me.SalinKeField()
                    Me.AktifkanDetail()
                    Me.dgvItem.Focus()
                    'nilai min juga harus diatas tanggal pengkuncian data terakhir
                    'msgbox("belum")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FrmO.Dispose()
        End Try
    End Sub

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Dim FrmO As New Core_Module.FrmMasterUser_Otorisasi

        Try
            If Me.TxtSalesOrderNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtSalesOrderNo.Text <> "" Then
                If Me.objSO.ApakahBolehHapus(Me.objSO.Primary, True, Me.UserCode) = False Then
                    Me.btnFind.Focus()
                Else
                    If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiSO_Otorisasi_Delete") = False Then
                        If Me.ObjUser.GetJmlByNamaHakAkses("FrmTransaksiSO_Otorisasi_Delete") = 0 Then
                            MsgBox("You don't have permission to perform this action.", MsgBoxStyle.Information, "Information")
                            Exit Sub
                        Else
                            FrmO.AccessName = "FrmTransaksiSO_Otorisasi_Delete"
                            If FrmO.ShowDialog <> Windows.Forms.DialogResult.OK Then
                                MsgBox("You don't have permission to perform this action.", MsgBoxStyle.Information, "Information")
                                Exit Sub
                            End If
                        End If
                        With Me.ObjOtorisasi
                            .Primary = 1
                            .PrimaryApprovedBy = Me.ObjUser.GetPrimary(FrmO.cboUserCode.Text)
                            .PrimaryDokumenAcuan = Me.TxtPrimarySO.Text
                            .PrimaryRequestBy = Me.ObjUser.GetPrimary(Me.UserCode)
                            .PrimaryTipeApproval = ClsMasterTipeOtorisasi.EnumTipeOtorisasi.SO_Delete
                            .Reason = FrmO.TxtReason.Text
                            .NoDocumentHist = Me.TxtSalesOrderNo.Text
                        End With
                    End If
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.objSO.Hapus(Me.TxtPrimarySO.Text, Me.UserCode, True) = True Then
                            Me.KosongkanField()
                            Me.btnNew.Focus()
                            Me.ObjOtorisasi.SimpanData()
                            MsgBox("Data deleted", MsgBoxStyle.Information, "Information")
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Windows.Forms.Cursor.Current = Cursors.Arrow
            FrmO.Dispose()
            Me.ResetObjOtorisasi()
        End Try
    End Sub

    Private Sub FrmMasterProductGroup_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If Me.txtRemarks.Enabled = True Then
                Me.BringToFront()
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.objSO.SimpanBaru(Me.dgvItem, Me.UserCode, Me, Me.vTBLItemDelivery) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnFind.Focus()
                            Else
                                e.Cancel = True
                            End If
                        Else
                            'simpan perubahan
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.objSO.SimpanPerubahan(Me.dgvItem, Me.UserCode, Me, Me.vTBLItemDelivery) = True Then
                                MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                                Me.SalinKeField()
                                Me.NonAktifkanDetail()
                                Me.btnEdit.Focus()
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
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetGbo1()
        Try
            Me.lblUnitsDefault.Text = Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNo.Text)
            Me.txtPrimaryUnitsDefault.Text = Me.ObjProduct.GetPrimaryUnit(Me.TxtPrimaryCodeNo.Text)
            Me.chkUseSecondary.Checked = Me.ObjProduct.ApakahUseSecondary(Me.TxtPrimaryCodeNo.Text)
            If Me.chkUseSecondary.Checked = True Then
                Me.LblUnitPrice.Text = "Unit Price/" + Me.ObjProduct.GetSecondaryUnitCode(Me.TxtPrimaryCodeNo.Text)
                Me.Hitung_Amount()
            Else
                Me.SetCboPriceByMultiUnit()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetCboPriceByMultiUnit()
        Dim UnitPrice, Discount As Decimal
        Dim UnitCodeDummy As String = ""

        Try
            Me.cboPriceByMultiUnit.Items.Clear()

            If Me.CboUnit.Items.Count = 1 Then
                Me.LblUnitPrice.Text = "Unit Price / " + Me.lblUnitsMulti.Text
                Me.cboPriceByMultiUnit.Visible = False
                Me.cboPriceByMultiUnit.TabStop = False
                Me.cboPriceByMultiUnit.Items.Add(Me.lblUnitsDefault.Text)
                Me.cboPriceByMultiUnit.SelectedIndex = 0
                Me.txtPriceByMultiUnit.Text = "1" 
            Else
                Me.LblUnitPrice.Text = "Unit Price"
                Me.cboPriceByMultiUnit.Visible = True
                Me.cboPriceByMultiUnit.TabStop = True
                Me.cboPriceByMultiUnit.Items.Add(Me.lblUnitsDefault.Text)
                UnitCodeDummy = Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNo.Text)
                If Me.CboUnit.Text <> UnitCodeDummy Then
                    Me.cboPriceByMultiUnit.Items.Add(Me.lblUnitsMulti.Text)
                    Me.cboPriceByMultiUnit.SelectedIndex = Me.ObjProduct.GetPriceByMultiUnit(Me.TxtPrimaryCodeNo.Text)
                Else
                    Me.cboPriceByMultiUnit.SelectedIndex = 0
                End If
                If Me.cboPriceByMultiUnit.Text <> UnitCodeDummy Then
                    Me.txtPriceByMultiUnit.Text = "1"
                Else
                    Me.txtPriceByMultiUnit.Text = "0"
                End If
                Me.cboPriceByMultiUnit.Left = Me.LblUnitPrice.Left + Me.LblUnitPrice.Width
            End If

            If Me.txtPriceByMultiUnit.Text = "1" Then
                Me.ObjProduct.GetUnitPriceNDiskon(Me.txtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtQtyperBatchMulti.Text, _
                Me.txtPrimaryCurrency.Text, Me.txtPrimaryUnitsMulti.Text, UnitPrice, Discount)
            Else
                Me.ObjProduct.GetUnitPriceNDiskon(Me.txtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtQtyperBatchDefault.Text, _
                Me.txtPrimaryCurrency.Text, Me.txtPrimaryUnitsDefault.Text, UnitPrice, Discount)
            End If
            If UnitPrice = 0 And Discount = 0 Then
                If Me.txtPriceByMultiUnit.Text = "1" Then
                    Me.objSO.GetHistoricalPrice(Me.txtPrimaryUnitsMulti.Text, Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryCurrency.Text, _
                    UnitPrice, Discount, Me.txtPriceByMultiUnit.Text, Me.txtPrimaryCustomer.Text)
                Else
                    Me.objSO.GetHistoricalPrice(Me.txtPrimaryUnitsDefault.Text, Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryCurrency.Text, _
                    UnitPrice, Discount, Me.txtPriceByMultiUnit.Text, Me.txtPrimaryCustomer.Text)
                End If
            End If
            Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(UnitPrice, Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text))
            Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Discount, 2)

            If Me.ChkDiscountConditionSama.Checked = True Then
                If IsNumeric(Me.txtReferenceDiscount.Text) = False Then
                    Me.txtReferenceDiscount.Text = Me.ObjCD.DigitGrouping(0, 2)
                End If
                Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Me.txtReferenceDiscount.Text, 2)
            End If

            Me.Hitung_Amount()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCodeNo.KeyDown
        Dim i As New frmCariMasterProduct
        Dim Jml As Long = 0
        i.BackColorAll = Me.BackColorAll
        Dim UnitPrice, Diskon As Decimal

        Try
            If e.KeyValue = 13 And Me.chkUseCustCodeNo.Checked = False Then
                Jml = Me.ObjProduct.GetJml(Me.TxtCodeNo.Text)
                If Jml = 1 Then
                    Me.TxtPrimaryCodeNo.Text = Me.ObjProduct.GetPrimary(Me.TxtCodeNo.Text)
                    Me.txtPrimaryCategory.Text = Me.ObjProduct.GetPrimaryCategory(Me.TxtPrimaryCodeNo.Text)
                    Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                    Me.ObjCustomer_sProductCode.IsiCustomerProductCodeNDesc(Me.txtCustomerProductCodeNo, _
                    Me.txtCustomerProductDesc, Me.txtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text)
                    Me.dtpdeliverydate.Focus()
                    Me.TxtDiscFromPerct.Text = "1"
                    Me.SetGbo1()
                    'Me.ObjProduct.GetUnitPriceNDiskon(Me.txtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtQtyperBatchMulti.Text, _
                    'Me.txtPrimaryCurrency.Text, Me.txtPrimaryUnitsMulti.Text, UnitPrice, Diskon)
                    'Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(UnitPrice, Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text))
                    'If Me.ChkDiscountConditionSama.Checked = True Then
                    '    If IsNumeric(Me.txtReferenceDiscount.Text) = False Then
                    '        Me.txtReferenceDiscount.Text = Me.ObjCD.DigitGrouping(0, 2)
                    '    End If
                    '    Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Me.txtReferenceDiscount.Text, 2)
                    'Else
                    '    Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Diskon, 2)
                    'End If
                    'Me.Hitung_Amount()
                ElseIf Jml > 1 Or Jml = 0 Then
                    'ada 2 product dengan kode yang sama atau tidak ada product yang dimaksud
                    If Me.TxtCodeNo.Text <> "" Then
                        i.txtCari.Text = Me.TxtCodeNo.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        Me.TxtPrimaryCodeNo.Text = i.PrimaryProduct
                        Me.TxtCodeNo.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryCodeNo.Text)
                        Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                        Me.ObjCustomer_sProductCode.IsiCustomerProductCodeNDesc(Me.txtCustomerProductCodeNo, _
                        Me.txtCustomerProductDesc, Me.txtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text)
                        Me.dtpdeliverydate.Focus()
                        Me.TxtDiscFromPerct.Text = "1"
                        Me.SetGbo1()
                        'Me.ObjProduct.GetUnitPriceNDiskon(Me.txtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtQtyperBatchMulti.Text, _
                        '                   Me.txtPrimaryCurrency.Text, Me.txtPrimaryUnitsMulti.Text, UnitPrice, Diskon)
                        'Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(UnitPrice, Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text))
                        'If Me.ChkDiscountConditionSama.Checked = True Then
                        '    If IsNumeric(Me.txtReferenceDiscount.Text) = False Then
                        '        Me.txtReferenceDiscount.Text = Me.ObjCD.DigitGrouping(0, 2)
                        '    End If
                        '    Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Me.txtReferenceDiscount.Text, 2)
                        'Else
                        '    Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Diskon, 2)
                        'End If
                        'Me.Hitung_Amount()
                    Else
                        Me.TxtCodeNo.Clear()
                        Me.TxtPrimaryCodeNo.Text = "-1"
                        Me.txtPrimaryCategory.Text = "-1"
                        Me.KosongkanItem()
                        Me.TxtCodeNo.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub TxtCodeNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNo.Leave
        Dim i As New frmCariMasterProduct
        i.BackColorAll = Me.BackColorAll
        Dim UnitPrice, Diskon As Decimal

        Try
            If Me.TxtCodeNo.Text <> "" Then
                If Me.ObjProduct.GetJml(Me.TxtCodeNo.Text) = 1 Then
                    Me.TxtPrimaryCodeNo.Text = Me.ObjProduct.GetPrimary(Me.TxtCodeNo.Text)
                    Me.txtPrimaryCategory.Text = Me.ObjProduct.GetPrimaryCategory(Me.TxtPrimaryCodeNo.Text)
                    Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                    Me.ObjCustomer_sProductCode.IsiCustomerProductCodeNDesc(Me.txtCustomerProductCodeNo, _
                    Me.txtCustomerProductDesc, Me.txtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text)
                    Me.dtpdeliverydate.Focus()
                    Me.TxtDiscFromPerct.Text = "1"
                    Me.SetGbo1()
                    'Me.ObjProduct.GetUnitPriceNDiskon(Me.txtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtQtyperBatchMulti.Text, _
                    'Me.txtPrimaryCurrency.Text, Me.txtPrimaryUnitsMulti.Text, UnitPrice, Diskon)
                    'Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(UnitPrice, Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text))
                    'If Me.ChkDiscountConditionSama.Checked = True Then
                    '    If IsNumeric(Me.txtReferenceDiscount.Text) = False Then
                    '        Me.txtReferenceDiscount.Text = Me.ObjCD.DigitGrouping(0, 2)
                    '    End If
                    '    Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Me.txtReferenceDiscount.Text, 2)
                    'Else
                    '    Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Diskon, 2)
                    'End If
                    'Me.Hitung_Amount()
                Else
                    If Me.TxtCodeNo.Text <> "" Then
                        i.txtCari.Text = Me.TxtCodeNo.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        Me.TxtPrimaryCodeNo.Text = i.PrimaryProduct
                        Me.TxtCodeNo.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryCodeNo.Text)
                        Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                        Me.ObjCustomer_sProductCode.IsiCustomerProductCodeNDesc(Me.txtCustomerProductCodeNo, _
                        Me.txtCustomerProductDesc, Me.txtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text)
                        Me.dtpdeliverydate.Focus()
                        Me.TxtDiscFromPerct.Text = "1"
                        Me.SetGbo1()
                        'Me.ObjProduct.GetUnitPriceNDiskon(Me.txtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtQtyperBatchMulti.Text, _
                        'Me.txtPrimaryCurrency.Text, Me.txtPrimaryUnitsMulti.Text, UnitPrice, Diskon)
                        'Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(UnitPrice, Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text))
                        'If Me.ChkDiscountConditionSama.Checked = True Then
                        '    If IsNumeric(Me.txtReferenceDiscount.Text) = False Then
                        '        Me.txtReferenceDiscount.Text = Me.ObjCD.DigitGrouping(0, 2)
                        '    End If
                        '    Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Me.txtReferenceDiscount.Text, 2)
                        'Else
                        '    Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Diskon, 2)
                        'End If
                        'Me.Hitung_Amount()
                    Else
                        Me.TxtCodeNo.Clear()
                        Me.TxtPrimaryCodeNo.Text = "-1"
                        Me.txtPrimaryCategory.Text = "-1"
                        Me.KosongkanItem()
                        Me.TxtCodeNo.Focus()
                    End If
                End If
            Else
                Me.TxtPrimaryCodeNo.Text = "-1"
                Me.txtDesc.Clear()
                Me.txtPrimaryCategory.Text = "-1"
                Me.KosongkanItem()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindProduct_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindCodeNo.Click
        Dim i As New frmCariMasterProduct
        Dim UnitPrice, Diskon As Decimal

        Try
            'cek apakah sudah isi kode currency
            i.BackColorAll = Me.BackColorAll
            If Me.txtPrimaryCustomer.Text <> "" Then
                If Me.txtPrimaryCustomer.Text = -1 And Me.txtPrimaryCustomer.Text <> "" Then
                    Me.TabControl1.SelectedIndex = 0
                    Me.txtCustomerCode.Focus()
                    MsgBox("To start entering item data, you must specify customer data first.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            If Me.txtPrimaryCurrency.Text <> "" Then
                If Me.txtPrimaryCurrency.Text = -1 And Me.txtPrimaryCurrency.Text <> "" Then
                    Me.TabControl1.SelectedIndex = 1
                    Me.txtCurrencyCode.Focus()
                    MsgBox("To start entering item data, you must specify currency data first.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            i.cboCriteria.SelectedIndex = 0
            If Me.TxtCodeNo.Text <> "" Then
                i.txtCari.Text = Me.TxtCodeNo.Text
            End If

            i.ShowDialog()
            If i.PrimaryProduct > 0 Then
                Me.TxtPrimaryCodeNo.Text = i.PrimaryProduct
                Me.TxtCodeNo.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryCodeNo.Text)
                Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                Me.ObjCustomer_sProductCode.IsiCustomerProductCodeNDesc(Me.txtCustomerProductCodeNo, _
                Me.txtCustomerProductDesc, Me.txtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text)
                Me.dtpdeliverydate.Focus()
                Me.TxtDiscFromPerct.Text = "1"
                Me.SetGbo1()
                'Me.ObjProduct.GetUnitPriceNDiskon(Me.txtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtQtyperBatchMulti.Text, _
                'Me.txtPrimaryCurrency.Text, Me.txtPrimaryUnitsMulti.Text, UnitPrice, Diskon)
                'Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(UnitPrice, Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text))
                'If Me.ChkDiscountConditionSama.Checked = True Then
                '    If IsNumeric(Me.txtReferenceDiscount.Text) = False Then
                '        Me.txtReferenceDiscount.Text = Me.ObjCD.DigitGrouping(0, 2)
                '    End If
                '    Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Me.txtReferenceDiscount.Text, 2)
                'Else
                '    Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Diskon, 2)
                'End If
                'Me.TxtDiscFromPerct.Text = "1"
                'Me.Hitung_Amount()
            Else
                Me.TxtCodeNo.Clear()
                Me.TxtPrimaryCodeNo.Text = "-1"
                Me.txtPrimaryCategory.Text = "-1"
                Me.KosongkanItem()
                Me.TxtCodeNo.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    '\\TR=Quo
    Private Function GetIndex(ByVal vPrimaryProduct As Double, _
    ByVal UnitPrice As Double, ByVal NettPrice As Double, ByVal Discount As Double, ByVal vPrimaryMultiUnit As Double, ByVal vPriceByMultiUnit As String) As Integer
        Dim row As DataGridViewRow

        Try
            For Each row In Me.dgvItem.Rows
                If CBool(row.Cells(Me.ColDeliveryCont.Index).Value) = True Then
                    If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And row.Cells(Me.ColUnitPrice.Index).Value = UnitPrice And _
                    row.Cells(Me.ColDiscount.Index).Value = Discount And row.Cells(Me.ColNettPrice.Index).Value = NettPrice And _
                    row.Cells(Me.ColPrimaryUnitsMulti.Index).Value = vPrimaryMultiUnit And row.Cells(Me.ColPriceByMultiUnit_Str.Index).Value = vPriceByMultiUnit Then
                        Return row.Index
                    End If
                Else
                    If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And _
                    row.Cells(Me.ColUnitPrice.Index).Value = UnitPrice And _
                    row.Cells(Me.ColDiscount.Index).Value = Discount And row.Cells(Me.ColNettPrice.Index).Value = NettPrice And _
                    row.Cells(Me.ColPrimaryUnitsMulti.Index).Value = vPrimaryMultiUnit And row.Cells(Me.ColPriceByMultiUnit_Str.Index).Value = vPriceByMultiUnit Then
                        Return row.Index
                    End If
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetIndex_Ratio(ByVal vPrimaryProduct As Double, _
    ByVal UnitPrice As Double, ByVal NettPrice As Double, ByVal Discount As Double, ByVal vRatioDefault As Double, ByVal vRatioReference As Double) As Integer
        Dim row As DataGridViewRow

        Try
            For Each row In Me.dgvItem.Rows
                If CBool(row.Cells(Me.ColDeliveryCont.Index).Value) = True Then
                    If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And row.Cells(Me.ColUnitPrice.Index).Value = UnitPrice And _
                    row.Cells(Me.ColDiscount.Index).Value = Discount And row.Cells(Me.ColNettPrice.Index).Value = NettPrice And _
                    CDbl(row.Cells(Me.ColRatioDefault.Index).Value) = vRatioDefault And CDbl(row.Cells(Me.ColRatioReference.Index).Value) = vRatioReference Then
                        Return row.Index
                    End If
                Else
                    If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And _
                    row.Cells(Me.ColUnitPrice.Index).Value = UnitPrice And _
                    row.Cells(Me.ColDiscount.Index).Value = Discount And row.Cells(Me.ColNettPrice.Index).Value = NettPrice And _
                    CDbl(row.Cells(Me.ColRatioDefault.Index).Value) = vRatioDefault And CDbl(row.Cells(Me.ColRatioReference.Index).Value) = vRatioReference Then
                        Return row.Index
                    End If
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetIndex(ByVal vPrimaryProduct As Double) As Integer
        Dim row As DataGridViewRow

        Try
            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct Then
                    Return row.Index
                End If
            Next

            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetIdentifierBaru() As String
        Dim IdentifierD As Integer = 0
        Dim NilaiMax As Integer = 0
        Dim Retval As String = ""

        Try
            For Each rowDGV As DataGridViewRow In Me.dgvItem.Rows
                IdentifierD = rowDGV.Cells(Me.ColIdentifier.Index).Value.ToString.IndexOf("-", 1)
                IdentifierD = rowDGV.Cells(Me.ColIdentifier.Index).Value.ToString.Substring(IdentifierD + 1)
                If NilaiMax < IdentifierD Then
                    NilaiMax = IdentifierD
                End If
            Next
            NilaiMax += 1

            Retval = Me.TxtPrimarySO.Text + "-" + NilaiMax.ToString
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function GetPrimaryMax(ByVal vTBL As DseTransaksiSO.tbltrans_so_itemdeliveryDataTable) As Double
        Dim Retval As Double = 0

        Try
            For Each Me.vrowItemDelivery In Me.vTBLItemDelivery.Rows
                If Retval < Me.vrowItemDelivery.primary Then
                    Retval = Me.vrowItemDelivery.primary
                End If
            Next

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub BtnInput_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnInput.Click
        Dim DefPrecDefault As Integer = 0
        Dim DefPrecMulti As Integer = 0
        Dim rowDGV As DataGridViewRow
        Dim rowDel As DseTransaksiSO.tbltrans_so_itemdeliveryRow
        Dim IdentifierBaru As String = ""

        Try
            If Me.ApakahDataItemBenar = True Then
                Hitung_Amount()

                Dim Index As Integer
                If Me.cboPriceByMultiUnit.Visible = True Then
                    Index = Me.GetIndex(Me.TxtPrimaryCodeNo.Text, _
                                  Me.TxtUnitPrice.Text, Me.txtNettPrice.Text, Me.TxtDiscount.Text, Me.txtPrimaryUnitsMulti.Text, "/" & Me.cboPriceByMultiUnit.Text)
                Else
                    Index = Me.GetIndex(Me.TxtPrimaryCodeNo.Text, _
              Me.TxtUnitPrice.Text, Me.txtNettPrice.Text, Me.TxtDiscount.Text, Me.txtPrimaryUnitsMulti.Text, "/" & Me.lblUnitsMulti.Text)
                End If
               

                DefPrecMulti = ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text)
                DefPrecDefault = Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text)

                'cari data yang sama persis
                If Index <> -1 Then
                    'tambahkan nilai qty
                    Dim row As DataGridViewRow = Me.dgvItem.Rows(Index)
                    'row.Cells(Me.ColQty.Index).Value = CDec(row.Cells(Me.ColQty.Index).Value) + CDec(Me.TxtQtyperBatchDefault.Text)
                    row.Cells(Me.ColQtyMulti.Index).Value = ObjCD.DigitGrouping(CDbl(row.Cells(Me.ColQtyMulti.Index).Value) + CDbl(Me.TxtQtyperBatchMulti.Text), DefPrecMulti)
                    If row.Cells(Me.ColToleranceDownFromPerc.Index).Value = "1" Then
                        row.Cells(Me.ColToleranceDown.Index).Value = ObjCD.DigitGrouping(CDec(row.Cells(Me.ColQtyMulti.Index).Value) * CDec(row.Cells(Me.ColToleranceDownPerc.Index).Value) / 100, DefPrecMulti)
                    Else
                        row.Cells(Me.ColToleranceDownPerc.Index).Value = ObjCD.DigitGrouping(CDec(row.Cells(Me.ColToleranceDown.Index).Value) / CDec(row.Cells(Me.ColQtyMulti.Index).Value) * 100, 2)
                    End If
                    If row.Cells(Me.ColToleranceUpFromPerc.Index).Value = "1" Then
                        row.Cells(Me.ColToleranceUp.Index).Value = ObjCD.DigitGrouping(CDec(row.Cells(Me.ColQtyMulti.Index).Value) * CDec(row.Cells(Me.ColToleranceUpPerc.Index).Value) / 100, DefPrecMulti)
                    Else
                        row.Cells(Me.ColToleranceUpPerc.Index).Value = ObjCD.DigitGrouping(CDec(row.Cells(Me.ColToleranceUp.Index).Value) / CDec(row.Cells(Me.ColQtyMulti.Index).Value) * 100, 2)
                    End If
                    Me.SetQtyDefaultOnDgv(row, DefPrecDefault)
                    Me.Hitung_Amount(row)
                    Me.Hitung_Total()

                    'tambahkan item delivery, pada baris terakhir
                    If row.Cells(Me.ColDeliveryCont.Index).Value = Boolean.TrueString Then
                        Me.vTBLItemDelivery.DefaultView.RowFilter = "identifier = '" + row.Cells(Me.ColIdentifier.Index).Value.ToString + "'"
                        Me.vTBLItemDelivery.DefaultView.Sort = "[delivery date] DESC"
                        Me.vrowItemDelivery = Me.vTBLItemDelivery.DefaultView(0).Row
                        'Me.vrowItemDelivery.delivery_qty += CDec(Me.TxtQtyperBatchDefault.Text)
                        Me.vrowItemDelivery.delivery_qty_multi += CDec(Me.TxtQtyperBatchMulti.Text)
                        Me.vrowItemDelivery.delivery_qty = GetQtyDefault(Me.vrowItemDelivery.delivery_qty_multi, row.Cells(Me.ColRatioDefault.Index).Value, row.Cells(Me.ColRatioReference.Index).Value, DefPrecDefault)
                    End If
                Else
                    'tambahkan baris baru
                    Dim ExpiredDateStr As String
                    If Me.dtpdeliverydate.Checked = False Then
                        ExpiredDateStr = "-"
                    Else
                        ExpiredDateStr = Format(Me.dtpdeliverydate.Value.Date, "dd MMM yyyy")
                    End If
                    IdentifierBaru = Me.GetIdentifierBaru
                    Me.dgvItem.Rows.Add()
                    rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                    rowDGV.Cells(Me.ColAmount.Index).Value = Me.txtAmount.Text
                    rowDGV.Cells(Me.ColCodeNo.Index).Value = Me.TxtCodeNo.Text
                    rowDGV.Cells(Me.ColDeliveryCont.Index).Value = Me.dtpdeliverydate.Checked
                    rowDGV.Cells(Me.ColDeliveryDate.Index).Value = ExpiredDateStr
                    rowDGV.Cells(Me.ColDeliveryHistory.Index).Value = ""
                    rowDGV.Cells(Me.ColDescription.Index).Value = Me.txtDesc.Text
                    rowDGV.Cells(Me.ColDiscAmount.Index).Value = Me.TxtDiscAmount.Text
                    rowDGV.Cells(Me.ColDiscount.Index).Value = Me.TxtDiscount.Text
                    rowDGV.Cells(Me.ColIdentifier.Index).Value = IdentifierBaru
                    rowDGV.Cells(Me.ColNettPrice.Index).Value = Me.txtNettPrice.Text
                    rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvItem.Rows.Count
                    rowDGV.Cells(Me.ColPriceCond.Index).Value = Me.cboPriceConditionItem.Text
                    rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value = Me.TxtPrimaryCodeNo.Text
                    rowDGV.Cells(Me.ColPrimaryUnits.Index).Value = Me.txtPrimaryUnitsDefault.Text
                    rowDGV.Cells(Me.ColQty.Index).Value = Me.TxtQtyperBatchDefault.Text
                    rowDGV.Cells(Me.ColRemark.Index).Value = Me.TxtRemarkItem.Text
                    rowDGV.Cells(Me.ColSuppDesc.Index).Value = ""
                    rowDGV.Cells(Me.ColSupplierCodeNo.Index).Value = ""
                    rowDGV.Cells(Me.ColUnit.Index).Value = Me.lblUnitsDefault.Text
                    rowDGV.Cells(Me.ColUnitPrice.Index).Value = Me.TxtUnitPrice.Text
                    rowDGV.Cells(Me.ColUseSecondary.Index).Value = Me.chkUseSecondary.Checked
                    rowDGV.Cells(Me.ColDiscFromPerct.Index).Value = Me.TxtDiscFromPerct.Text

                    With rowDGV
                        .Cells(Me.ColQtyMulti.Index).Value = Me.TxtQtyperBatchMulti.Text
                        .Cells(Me.ColUnitsMulti.Index).Value = Me.lblUnitsMulti.Text
                        .Cells(Me.ColPrimaryUnitsMulti.Index).Value = Me.txtPrimaryUnitsMulti.Text
                        .Cells(Me.ColRatioDefault.Index).Value = Me.ObjCD.DigitGrouping(Me.objSO.GetRatioDefault(.Cells(Me.ColPrimaryCodeNo.Index).Value, _
                        .Cells(Me.ColPrimaryUnitsMulti.Index).Value), DefPrecDefault)
                        .Cells(Me.ColRatioReference.Index).Value = Me.ObjCD.DigitGrouping(Me.objSO.GetRatioReference(.Cells(Me.ColPrimaryCodeNo.Index).Value, _
                        .Cells(Me.ColPrimaryUnitsMulti.Index).Value), DefPrecMulti)
                        Me.SetRatioOnDgv(rowDGV)
                        .Cells(Me.ColToleranceDown.Index).Value = Me.ObjCD.DigitGrouping("0", DefPrecMulti)
                        .Cells(Me.ColToleranceUp.Index).Value = Me.ObjCD.DigitGrouping("0", DefPrecMulti)
                        .Cells(Me.ColToleranceDownPerc.Index).Value = Me.ObjCD.DigitGrouping("0", 2)
                        .Cells(Me.ColToleranceUpPerc.Index).Value = Me.ObjCD.DigitGrouping("0", 2)
                        .Cells(Me.ColToleranceDownFromPerc.Index).Value = 1
                        .Cells(Me.ColToleranceUpFromPerc.Index).Value = 1

                        rowDGV.Cells(Me.ColPriceByMultiUnit.Index).Value = Me.txtPriceByMultiUnit.Text
                        If Me.txtPriceByMultiUnit.Text = "1" Then
                            rowDGV.Cells(Me.ColPriceByMultiUnit_Str.Index).Value = "/" + .Cells(Me.ColUnitsMulti.Index).Value
                        Else
                            rowDGV.Cells(Me.ColPriceByMultiUnit_Str.Index).Value = "/" + rowDGV.Cells(Me.ColUnit.Index).Value
                        End If
                    End With

                    'masukkan data item delivery
                    If Me.dtpdeliverydate.Checked = True Then
                        rowDel = Me.vTBLItemDelivery.Newtbltrans_so_itemdeliveryRow
                        With rowDel
                            .delivery_date = Me.dtpdeliverydate.Value
                            .delivery_priority = Me.ObjDeliveryPriority.GetPrimaryDefault
                            .delivery_qty = Me.TxtQtyperBatchDefault.Text
                            .identifier = Me.dgvItem.Item(Me.ColIdentifier.Index, Me.dgvItem.RowCount - 1).Value
                            .primary = Me.GetPrimaryMax(Me.vTBLItemDelivery) + 1
                            .primary_transaksi = Me.TxtPrimarySO.Text
                            .remarks = ""
                            .delivery_qty_multi = Me.TxtQtyperBatchMulti.Text
                        End With
                        vTBLItemDelivery.Rows.Add(rowDel)
                    End If
                End If

                'kosongkan gbo item
                Me.KosongkanItem()
                Me.txtNettPrice.Text = Me.ObjCD.SetDecimal(Me.txtNettPrice.Text, Me.DefaultPrecision)
                Me.txtAmount.Text = Me.ObjCD.SetDecimal(Me.txtAmount.Text, Me.DefaultPrecision)
                If Me.chkUseCustCodeNo.Checked = True Then
                    Me.txtCustomerProductCodeNo.Focus()
                Else
                    Me.TxtCodeNo.Focus()
                End If

                Me.FormatDGV()

                'fokuskan ke baris terakhir
                While Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1).Displayed = False _
                Or Me.dgvItem.DisplayedRowCount(True) <> Me.dgvItem.DisplayedRowCount(False)
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                End While

                'hitung total
                Me.Hitung_Total()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetJmlItemDelivery(ByVal vIdentifier As String) As Double
        Dim Retval As Integer = 0

        Try
            For Each Me.vrowItemDelivery In Me.vTBLItemDelivery.Rows
                If Me.vrowItemDelivery.identifier = vIdentifier Then
                    Retval += 1
                End If
            Next

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub TxtQtyperBatch_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchDefault.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtQtyperBatchDefault, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatch_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyperBatchDefault.KeyDown
        Me.ObjCD.Number_KeyDown(Me.TxtQtyperBatchDefault, sender, e)
    End Sub

    Private Sub TxtQtyperBatch_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyperBatchDefault.KeyUp
        Me.ObjCD.Number_KeyUp(Me.TxtQtyperBatchDefault, sender, e)
    End Sub

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Dim ListNoTrans As New List(Of String)
        Dim Dummy As String = ""
        Dim Counter As Integer = 0

        Try
            'yang bisa diedit adalah:
            '1. Qty
            '2. Unit Price
            '3. Diskon & Disc Amount
            '4. Delivery Date Controll
            '5. Price condition
            '6. Remarks
            'untuk delivery date, diubah melalui event mouse clicks
            If Me.txtRemarks.Enabled = True Then
                If e.ColumnIndex = Me.ColQtyMulti.Index Or e.ColumnIndex = Me.ColUnitPrice.Index Or e.ColumnIndex = Me.ColDiscAmount.Index Or e.ColumnIndex = Me.ColDiscount.Index _
                Or e.ColumnIndex = Me.ColDeliveryCont.Index Or e.ColumnIndex = Me.ColPriceCond.Index Or e.ColumnIndex = Me.ColRemark.Index Or _
                e.ColumnIndex = Me.ColToleranceDown.Index Or e.ColumnIndex = Me.ColToleranceDownPerc.Index Or e.ColumnIndex = Me.ColToleranceUp.Index Or _
                e.ColumnIndex = Me.ColToleranceUpPerc.Index Then
                    'untuk qty, unit price, diskon, nett price 
                    If e.ColumnIndex = Me.ColDeliveryDate.Index Then
                        'cek apakah expexted delivery date dikontrol
                        If CStr(Me.dgvItem.Item(Me.ColDeliveryCont.Index, e.RowIndex).Value) = Boolean.FalseString Then
                            e.Cancel = True
                            Exit Sub
                        End If
                    End If
                    'cek apakah price condition boleh diedit, kasih peringatan kalo
                    'sudah ada invoice
                    'If e.ColumnIndex = Me.ColPriceCond.Index Then
                    '    'Dim ObjItemInv As New ClsTransaksiSalesInvoiceItem
                    '    If ObjItemInv.GetJml(Me.TxtPrimarySO.Text, Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value) > 0 Then
                    '        Dim ObjInv As New ClsTransaksiSalesInvoice
                    '        Dim InvNo As String = ObjInv.GetInvoiceNoBySO_n_Product(Me.TxtPrimarySO.Text, _
                    '        Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value)
                    '        MsgBox("Warning this item is already invoiced in: " + InvNo + _
                    '        ". Make sure your are aware of this.", _
                    '        MsgBoxStyle.Information, "Information")
                    '        ObjInv = Nothing
                    '    End If
                    '    ObjItemInv = Nothing
                    'End If
                    'kalo mau edit price condition sedangkan price condition dianggap sama exclude maka tidak bisa juga
                    If e.ColumnIndex = Me.ColPriceCond.Index And Me.ChkPriceCondition.Checked = True Then
                        MsgBox("Price condition is locked, to edit this data please unchecked the price condition check box.", _
                        MsgBoxStyle.Information, "Information")
                        e.Cancel = True
                        Exit Sub
                    End If
                    'kalo mau edit diskon sedangkan diskon kondisi sama dengan reference maka tidak bisa dilakukan
                    If e.ColumnIndex = Me.ColDiscount.Index And Me.ChkDiscountConditionSama.Checked = True Then
                        MsgBox("Discount condition is locked, to edit this data please unchecked the discount condition check box.", _
                        MsgBoxStyle.Information, "Information")
                        e.Cancel = True
                        Exit Sub
                    End If
                    'kalo mau edit delivery date, maka harus memiliki hak akses-nya
                    If e.ColumnIndex = Me.ColDeliveryCont.Index Then
                        If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiSO_Override_Promise_To_Deliver") = False Then
                            e.Cancel = True
                            Exit Sub
                        End If
                    End If

                    'cek hak akses jika mau mengubah harga atas item yang sudah diinvoice 
                    If Me.baru = False Then
                        If e.ColumnIndex = Me.ColQuo_UnitPrice.Index Or _
                        e.ColumnIndex = Me.ColDiscount.Index Or _
                        e.ColumnIndex = Me.ColDiscAmount.Index Then

                            If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiSO_EditPriceStlhInv") = False Then
                                Me.objInv.GetNoInvoice(Me.TxtPrimarySO.Text, Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, ListNoTrans)
                                If ListNoTrans.Count > 0 Then

                                    For Each NoTrans As String In ListNoTrans
                                        Counter += 1

                                        If Dummy = "" Then
                                            Dummy = Counter.ToString + ". " + NoTrans
                                        Else
                                            Dummy = Counter.ToString + vbCrLf + ". " + NoTrans
                                        End If
                                    Next

                                    MsgBox("The following invoice refers to this data:" + vbCrLf + Dummy + vbCrLf + vbCrLf + _
                                    "You don't have acces to do this change, please contact your administrator.", MsgBoxStyle.Information, "Information")

                                    e.Cancel = True
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If

                    'kalo sampe sini maka boleh edit
                    Me.buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                    '\\TR#
                    Me.bufferToleransiDown = Me.dgvItem.Item(Me.ColToleranceDown.Index, e.RowIndex).Value
                    Me.bufferToleransiUp = Me.dgvItem.Item(Me.ColToleranceUp.Index, e.RowIndex).Value
                    'buang nol di belakang decimal jika edit kolom qty
                    If e.ColumnIndex = Me.ColQtyMulti.Index Or _
                    e.ColumnIndex = Me.ColDiscAmount.Index Or _
                    e.ColumnIndex = Me.ColDiscount.Index Or _
                    e.ColumnIndex = Me.ColUnitPrice.Index Or _
                    e.ColumnIndex = Me.ColToleranceDown.Index Or _
                    e.ColumnIndex = Me.ColToleranceDownPerc.Index Or _
                    e.ColumnIndex = Me.ColToleranceUp.Index Or _
                    e.ColumnIndex = Me.ColToleranceUpPerc.Index Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.BuangNolDibelakangDecimal(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value)
                    End If
                Else
                    e.Cancel = True
                End If
            Else
                'hanya bisa edit pada saat detail aktif
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListNoTrans = Nothing
        End Try
    End Sub

    Private Sub dgvItem_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellContentClick
        Dim ij As New FrmTransaksiSO_SetDelivery
        Dim DI As Integer = 0
        Dim rowD As DseTransaksiSO.tbltrans_so_itemdeliveryRow
        Dim JmlDelivery As Integer = 0
        Dim DelDateD As Date = Today
        Dim JmlDefault As Decimal = 0

        Try
            If e.RowIndex = -1 Then
                Exit Sub
            End If
            If e.ColumnIndex = Me.ColDeliveryDate.Index Then
                If Me.dgvItem.Item(Me.ColDeliveryDate.Index, e.RowIndex).Value <> "-" And _
                Me.dgvItem.Item(Me.ColDeliveryCont.Index, e.RowIndex).Value = Boolean.TrueString Then

                    ij.vPrimaryUnitMulti = Me.dgvItem.Item(Me.ColPrimaryUnitsMulti.Index, e.RowIndex).Value
                    ij.vPrimaryUnitDefault = Me.dgvItem.Item(Me.ColPrimaryUnits.Index, e.RowIndex).Value
                    ij.vRatioReference = Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value
                    ij.vRatioDefault = Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value
                    ij.vDefPrecMulti = Me.ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, _
                    Me.dgvItem.Item(Me.ColPrimaryUnitsMulti.Index, e.RowIndex).Value)
                    ij.vDefPrecDefault = Me.ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value)

                    ij.vPrimaryCodeNo = Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value
                    ij.VIdentifier = Me.dgvItem.Item(Me.ColIdentifier.Index, e.RowIndex).Value
                    ij.FrmRef = Me
                    If Me.txtRemarks.Enabled = False Then
                        ij.vReadOnly = True
                        ij.ShowDialog()
                        Exit Sub
                    End If
                    If ij.ShowDialog = Windows.Forms.DialogResult.OK Then
                        'hapuskan semua item delivery yang identifiernya di edit
                        DI = 0
                        While DI <= Me.vTBLItemDelivery.Rows.Count - 1
                            Me.vrowItemDelivery = Me.vTBLItemDelivery.Rows(DI)
                            If Me.vrowItemDelivery.identifier = ij.VIdentifier Then
                                Me.vTBLItemDelivery.Rows.Remove(Me.vrowItemDelivery)
                            Else
                                DI += 1
                            End If
                        End While
                        'masukkan data item delivery baru
                        ij.dgvDetail.AllowUserToAddRows = False
                        For Each rowDGV As DataGridViewRow In ij.dgvDetail.Rows
                            rowD = Me.vTBLItemDelivery.Newtbltrans_so_itemdeliveryRow
                            With rowD
                                .delivery_date = CDate(rowDGV.Cells(ij.ColDeliveryDate.Index).Value)
                                .delivery_priority = rowDGV.Cells(ij.ColPrimaryPriority.Index).Value
                                .delivery_qty = rowDGV.Cells(ij.ColQty.Index).Value
                                .identifier = ij.VIdentifier
                                .primary = Me.GetPrimaryMax(Me.vTBLItemDelivery) + 1
                                .primary_transaksi = Me.TxtPrimarySO.Text
                                .remarks = rowDGV.Cells(ij.ColRemarks.Index).Value
                                .delivery_qty_multi = rowDGV.Cells(ij.ColQtyMulti.Index).Value
                            End With
                            Me.vTBLItemDelivery.Rows.Add(rowD)
                            JmlDelivery += 1
                            DelDateD = CDate(rowDGV.Cells(ij.ColDeliveryDate.Index).Value)
                        Next

                        If JmlDelivery > 1 Then
                            Me.dgvItem.Item(Me.ColDeliveryDate.Index, e.RowIndex).Value = "Multiple"
                        ElseIf JmlDelivery = 1 Then
                            Me.dgvItem.Item(Me.ColDeliveryDate.Index, e.RowIndex).Value = Format(DelDateD, "dd MMM yyyy")
                        End If
                    End If
                End If

                Me.EndEditDeliveryDateCont2()

            ElseIf e.ColumnIndex = Me.ColDeliveryHistory.Index Then
                Me.objSO.ExplainDelivery(Me.dgvItem.Item(Me.ColIdentifier.Index, e.RowIndex).Value)
            End If

            If e.ColumnIndex = Me.ColRatioDesc.Index Then
                Dim FrmRatio As New FrmTransaksiQuotation_SetQtyRatio
                Dim row As DataGridViewRow

                If Me.dgvItem.Item(Me.ColPrimaryUnits.Index, e.RowIndex).Value = Me.dgvItem.Item(Me.ColPrimaryUnitsMulti.Index, e.RowIndex).Value Or _
                    Me.txtRemarks.Enabled = False Then
                    Exit Sub
                End If
                Dim BufferRatioReference As String = Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value.ToString
                Dim BufferRatioDefault As String = Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value.ToString
                FrmRatio.vDefPrecMulti = ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, Me.dgvItem.Item(Me.ColPrimaryUnitsMulti.Index, e.RowIndex).Value)
                FrmRatio.vDefPrecDefault = Me.ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value)
                FrmRatio.vRatioMulti = Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value
                FrmRatio.vUnitMulti = Me.dgvItem.Item(Me.ColUnitsMulti.Index, e.RowIndex).Value
                FrmRatio.vRatioDefault = Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value
                FrmRatio.vUnitDefault = Me.dgvItem.Item(Me.ColUnit.Index, e.RowIndex).Value
                If FrmRatio.ShowDialog = Windows.Forms.DialogResult.OK Then
                    Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value = FrmRatio.TxtRatioMulti.Text
                    Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value = FrmRatio.TxtRatioDefault.Text
                    row = Me.dgvItem.Rows(e.RowIndex)
                    Me.SetRatioOnDgv(row)
                    JmlDefault = row.Cells(Me.ColQty.Index).Value
                    Me.SetQtyDefaultOnDgv(row, FrmRatio.vDefPrecDefault)

                    If CDec(Me.dgvItem.Item(Me.ColQty.Index, e.RowIndex).Value) <= 0 Then
                        MsgBox("Ratio restored to previous value, because it make default qty=0.", MsgBoxStyle.Information, "Information")
                        Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value = BufferRatioReference
                        Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value = BufferRatioDefault
                        Me.SetRatioOnDgv(row)
                        Me.SetQtyDefaultOnDgv(row, FrmRatio.vDefPrecDefault)
                        Exit Sub
                    End If

                    Me.vTBLItemDelivery.DefaultView.RowFilter = "[identifier] = '" & dgvItem.Item(Me.ColIdentifier.Index, e.RowIndex).Value & "'"
                    If Me.vTBLItemDelivery.DefaultView.Count > 0 Then
                        For i As Integer = 0 To Me.vTBLItemDelivery.DefaultView.Count - 1
                            Me.vrowItemDelivery = Me.vTBLItemDelivery.DefaultView(i).Row
                            Me.vrowItemDelivery.delivery_qty = ObjCD.DigitGrouping(vrowItemDelivery.delivery_qty_multi * CDec(FrmRatio.TxtRatioDefault.Text) / CDec(FrmRatio.TxtRatioMulti.Text), FrmRatio.vDefPrecDefault)
                        Next
                    End If

                    If row.Cells(Me.ColPriceByMultiUnit.Index).Value = 0 Then
                        If JmlDefault <> row.Cells(Me.ColQty.Index).Value Then
                            'hitung ulang row
                            Me.Hitung_Amount(row)
                            'hitung total
                            Me.Hitung_Total()
                        End If
                    End If
                End If
            ElseIf e.ColumnIndex = Me.ColPriceByMultiUnit_Str.Index And Me.txtRemarks.Enabled = True Then
                If Me.dgvItem.Item(Me.ColUnitsMulti.Index, e.RowIndex).Value <> Me.dgvItem.Item(Me.ColUnit.Index, e.RowIndex).Value Then
                    Dim FrmRef As New Core_Module.FrmTrans_PriceByMultiUnit
                    FrmRef.UnitMulti = Me.dgvItem.Item(Me.ColUnitsMulti.Index, e.RowIndex).Value
                    FrmRef.UnitDefault = Me.dgvItem.Item(Me.ColUnit.Index, e.RowIndex).Value
                    FrmRef.Buffer_PriceByMultiUnit = Me.dgvItem.Item(Me.ColPriceByMultiUnit.Index, e.RowIndex).Value

                    If FrmRef.ShowDialog = Windows.Forms.DialogResult.OK Then
                        Me.dgvItem.Item(Me.ColPriceByMultiUnit.Index, e.RowIndex).Value = FrmRef.cboPriceByMultiUnit.SelectedIndex
                        If Me.dgvItem.Item(Me.ColPriceByMultiUnit.Index, e.RowIndex).Value = "1" Then
                            Me.dgvItem.Item(Me.ColPriceByMultiUnit_Str.Index, e.RowIndex).Value = "/" + Me.dgvItem.Item(Me.ColUnitsMulti.Index, e.RowIndex).Value.ToString
                        Else
                            Me.dgvItem.Item(Me.ColPriceByMultiUnit_Str.Index, e.RowIndex).Value = "/" + Me.dgvItem.Item(Me.ColUnit.Index, e.RowIndex).Value.ToString
                        End If
                        'hitung ulang row
                        Me.Hitung_Amount(Me.dgvItem.Rows(e.RowIndex))
                        'hitung total
                        Me.Hitung_Total()
                    End If

                    FrmRef = Nothing
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
        Try
            Dim row As DataGridViewRow = Me.dgvItem.Rows(e.RowIndex)

            'pastikan ada isinya
            With row
                If IsNothing(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    'Bila tidak ada perubahan data maka tidak perlu dilakukan cek
                    If e.ColumnIndex <> Me.ColDeliveryDate.Index Then
                        If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer Then
                            Exit Sub
                        Else
                            If IsNumeric(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True And _
                            IsNumeric(Me.buffer) = True Then
                                If CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = CDec(Me.buffer) Then
                                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                                    Exit Sub
                                End If
                            End If
                        End If
                    Else
                        If IsDate(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = True Then
                            If CDate(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value).Date = CDate(Me.buffer).Date Then
                                Exit Sub
                            End If
                        End If
                    End If

                    If e.ColumnIndex <> Me.ColRemark.Index Then
                        'kalo bukan data remarks maka perlu dicek kebenaran data hasil edit
                        If e.ColumnIndex = Me.ColQtyMulti.Index Then
                            Me.EndEditQty(sender, e, row)
                            If CekBatasanToleransi(e.RowIndex) = False Then
                                Me.dgvItem.Item(Me.ColQtyMulti.Index, e.RowIndex).Value = Me.buffer
                                Me.dgvItem.Item(Me.ColToleranceDown.Index, e.RowIndex).Value = Me.bufferToleransiDown
                                Me.dgvItem.Item(Me.ColToleranceUp.Index, e.RowIndex).Value = Me.bufferToleransiUp
                                Me.EndEditQty(sender, e, row)
                                Exit Sub
                            End If
                        ElseIf e.ColumnIndex = Me.ColUnitPrice.Index Then
                            'unit price
                            Me.EndEditUnitPrice(sender, e, row)
                        ElseIf e.ColumnIndex = Me.ColDiscount.Index Then
                            'discount
                            Me.EndEditDiscount(sender, e, row)
                        ElseIf e.ColumnIndex = Me.ColDeliveryCont.Index Then
                            'kontrol delivery date
                            Me.EndEditDeliveryDateCont(sender, e, row)
                            'ElseIf e.ColumnIndex = 13 Then
                            '    'delivery date
                            '    Me.EndEditDeliveryDate(sender, e, row)
                        ElseIf e.ColumnIndex = Me.ColPriceCond.Index Then
                            'price condition
                            Me.Hitung_Total()
                        ElseIf e.ColumnIndex = Me.ColDiscAmount.Index Then
                            Me.EndEditDiskonAmount(sender, e, row)
                        ElseIf e.ColumnIndex = Me.ColToleranceDown.Index Then
                            Me.EndEditToleranceDown(sender, e, row)
                        ElseIf e.ColumnIndex = Me.ColToleranceDownPerc.Index Then
                            Me.EndEditToleranceDownPerc(sender, e, row)
                        ElseIf e.ColumnIndex = Me.ColToleranceUp.Index Then
                            Me.EndEditToleranceUp(sender, e, row)
                        ElseIf e.ColumnIndex = Me.ColToleranceUpPerc.Index Then
                            Me.EndEditToleranceUpPerc(sender, e, row)
                        End If
                    End If
                Else
                    If e.ColumnIndex = Me.ColRemark.Index Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = ""
                    Else
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetPreviousRowDeliveryDate(ByVal vPreviousRowIndex As Integer) As Date
        Try
            If vPreviousRowIndex > -1 Then
                If Me.dgvItem.Item(Me.ColDeliveryCont.Index, vPreviousRowIndex).Value = True Then
                    If IsDate(Me.dgvItem.Item(Me.ColDeliveryDate.Index, vPreviousRowIndex).Value) = True Then
                        Return CDate(Me.dgvItem.Item(Me.ColDeliveryDate.Index, vPreviousRowIndex).Value)
                    Else
                        Me.vTBLItemDelivery.DefaultView.RowFilter = "Identifier = '" & Me.dgvItem.Item(Me.ColIdentifier.Index, vPreviousRowIndex).Value & "'"
                        If Me.vTBLItemDelivery.DefaultView.Count > 0 Then
                            Me.vTBLItemDelivery.DefaultView.Sort = "[delivery date] ASC"
                            Me.vrowItemDelivery = Me.vTBLItemDelivery.DefaultView(0).Row
                            Return Me.vrowItemDelivery.delivery_date
                        End If
                    End If
                End If
            End If
            Return Me.dtpDate.Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return Me.dtpDate.Value
        End Try
    End Function

    Private Sub EndEditDeliveryDateCont(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal row As DataGridViewRow)
        Dim ij As Double = 0
        Dim rowD As DseTransaksiSO.tbltrans_so_itemdeliveryRow

        Try
            With row
                If row.Cells(Me.ColDeliveryCont.Index).Value = Boolean.TrueString Then
                    'jika jadi true maka delivery date jadi tanggal transaksi
                    row.Cells(Me.ColDeliveryDate.Index).Value = Format(Me.GetPreviousRowDeliveryDate(e.RowIndex - 1), "dd MMM yyyy")
                    'tambahkan ke dalam item delivery date
                    rowD = Me.vTBLItemDelivery.Newtbltrans_so_itemdeliveryRow
                    With rowD
                        .delivery_date = Me.GetPreviousRowDeliveryDate(e.RowIndex - 1)
                        .delivery_priority = Me.ObjDeliveryPriority.GetPrimaryDefault
                        .delivery_qty = Me.dgvItem.Item(Me.ColQty.Index, e.RowIndex).Value
                        .identifier = Me.dgvItem.Item(Me.ColIdentifier.Index, e.RowIndex).Value
                        .primary = Me.GetPrimaryMax(Me.vTBLItemDelivery) + 1
                        .primary_transaksi = Me.TxtPrimarySO.Text
                        .remarks = ""
                        .delivery_qty_multi = Me.dgvItem.Item(Me.ColQtyMulti.Index, e.RowIndex).Value
                    End With
                    Me.vTBLItemDelivery.Rows.Add(rowD)
                Else
                    'jika jadi false maka delivery date jadi "-"
                    row.Cells(Me.ColDeliveryDate.Index).Value = "-"
                    'buang item delivery
                    While ij < Me.vTBLItemDelivery.Rows.Count
                        Me.vrowItemDelivery = Me.vTBLItemDelivery.Rows(ij)
                        If Me.vrowItemDelivery.identifier = Me.dgvItem.Item(Me.ColIdentifier.Index, e.RowIndex).Value Then
                            Me.vTBLItemDelivery.Rows.Remove(Me.vrowItemDelivery)
                        Else
                            ij += 1
                        End If
                    End While
                End If

                Me.EndEditDeliveryDateCont2()
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditQty(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal row As DataGridViewRow)
        Dim JmlKurang As Decimal = 0
        Dim DefPrecMulti As Integer = 0
        Try
            With row
                'untuk nilai qty dan secondary qty, pastikan nilainya positif
                If Me.ObjCD.CekBilanganPositif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("Please enter positive number for quantity data.", MsgBoxStyle.Information, _
                    "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
                'nilai maksimal yang bisa dimasukkan adalah: 
                If CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) > 999999999999 Then
                    MsgBox("Maximum Qty value is :" + Me.ObjCD.DigitGrouping(999999999999, _
                    Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text)) + ".", _
                    MsgBoxStyle.Information, "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If

                DefPrecMulti = ObjProduct.GetAngkaDecimal(row.Cells(Me.ColPrimaryCodeNo.Index).Value, row.Cells(Me.ColPrimaryUnitsMulti.Index).Value)
                If row.Cells(Me.ColToleranceDownFromPerc.Index).Value = "1" Then
                    '\\TR#
                    If .Cells(Me.ColToleranceDownPerc.Index).Value = 50 Then
                        row.Cells(Me.ColToleranceDown.Index).Value = ObjCD.DigitGrouping(CDec(row.Cells(Me.ColQtyMulti.Index).Value) * CDec(row.Cells(Me.ColToleranceDownPerc.Index).Value) / 100, DefPrecMulti, ClsCekData.EnumTipeRounding.Down)
                    Else
                        row.Cells(Me.ColToleranceDown.Index).Value = ObjCD.DigitGrouping(CDec(row.Cells(Me.ColQtyMulti.Index).Value) * CDec(row.Cells(Me.ColToleranceDownPerc.Index).Value) / 100, DefPrecMulti)
                    End If
                Else
                    If CDec(row.Cells(Me.ColToleranceDown.Index).Value) > (CDec(row.Cells(e.ColumnIndex).Value) / 2) Then
                        MsgBox("Qty you entered less than 50% of tolerance down. Change tolerance down first.", MsgBoxStyle.Information, "Information")
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                    row.Cells(Me.ColToleranceDownPerc.Index).Value = ObjCD.DigitGrouping(CDec(row.Cells(Me.ColToleranceDown.Index).Value) / CDec(row.Cells(Me.ColQtyMulti.Index).Value) * 100, 2)
                End If
                If row.Cells(Me.ColToleranceUpFromPerc.Index).Value = "1" Then
                    '\\TR#
                    If .Cells(Me.ColToleranceUpPerc.Index).Value = 100 Then
                        row.Cells(Me.ColToleranceUp.Index).Value = ObjCD.DigitGrouping(CDec(row.Cells(Me.ColQtyMulti.Index).Value) * CDec(row.Cells(Me.ColToleranceUpPerc.Index).Value) / 100, DefPrecMulti, ClsCekData.EnumTipeRounding.Down)
                    Else
                        row.Cells(Me.ColToleranceUp.Index).Value = ObjCD.DigitGrouping(CDec(row.Cells(Me.ColQtyMulti.Index).Value) * CDec(row.Cells(Me.ColToleranceUpPerc.Index).Value) / 100, DefPrecMulti)
                    End If
                Else
                    If CDec(row.Cells(Me.ColToleranceUp.Index).Value) > (CDec(row.Cells(e.ColumnIndex).Value)) Then
                        MsgBox("Qty you entered less than 100% of tolerance up. Change tolerance up first", MsgBoxStyle.Information, "Information")
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                    row.Cells(Me.ColToleranceUpPerc.Index).Value = ObjCD.DigitGrouping(CDec(row.Cells(Me.ColToleranceUp.Index).Value) / CDec(row.Cells(Me.ColQtyMulti.Index).Value) * 100, 2)
                End If

                'digit grouping
                Me.dgvItem(e.ColumnIndex, e.RowIndex).Value = Me.ObjCD.DigitGrouping(Me.dgvItem(e.ColumnIndex, e.RowIndex).Value, DefPrecMulti)
                Me.SetQtyDefaultOnDgv(row, ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value))
                If CDec(row.Cells(Me.ColQty.Index).Value) <= 0 Then
                    MsgBox("Qty restored to previous value, because it make default qty=0.", MsgBoxStyle.Information, "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    SetQtyDefaultOnDgv(row, ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value))
                    Exit Sub
                End If

                'cek apakah edit sah dilihat dari sisi order balance
                If Me.ApakahEditSah(row, row.Cells(Me.ColPrimaryCodeNo.Index).Value) = False Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Me.SetQtyDefaultOnDgv(row, ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value))
                    Exit Sub
                End If
                'cek apakah edit sah dilihat dari sisi booking inv + git
                If Me.ApakahEditSahBooking(row, row.Cells(Me.ColPrimaryCodeNo.Index).Value) = False Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Me.SetQtyDefaultOnDgv(row, ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value))
                    Exit Sub
                End If
                'hitung ulang row
                Me.Hitung_Amount(row)
                'hitung total
                Me.Hitung_Total()

                If Me.dgvItem.Item(Me.ColDeliveryCont.Index, e.RowIndex).Value = Boolean.TrueString Then
                    If CDec(Me.buffer) > CDec(Me.dgvItem.Item(Me.ColQtyMulti.Index, e.RowIndex).Value) Then
                        'untuk item delivery, jika qty turun, maka hapuskan mulai dari baris paling bawah
                        JmlKurang = CDec(Me.buffer) - CDec(Me.dgvItem.Item(Me.ColQtyMulti.Index, e.RowIndex).Value)
                        Me.vTBLItemDelivery.DefaultView.RowFilter = "identifier = '" + Me.dgvItem.Item(Me.ColIdentifier.Index, e.RowIndex).Value + "'"
                        Me.vTBLItemDelivery.DefaultView.Sort = "[delivery date] DESC"
                        While JmlKurang > 0
                            If Me.vTBLItemDelivery.DefaultView.Count > 0 Then
                                Me.vrowItemDelivery = Me.vTBLItemDelivery.DefaultView(0).Row
                                If Me.vrowItemDelivery.delivery_qty_multi > JmlKurang Then
                                    Me.vrowItemDelivery.delivery_qty_multi -= JmlKurang
                                    Me.vrowItemDelivery.delivery_qty = Me.GetQtyDefault(Me.vrowItemDelivery.delivery_qty_multi, _
                                    row.Cells(Me.ColRatioDefault.Index).Value, row.Cells(Me.ColRatioReference.Index).Value, _
                                    Me.ObjProduct.GetAngkaDecimal(row.Cells(Me.ColPrimaryCodeNo.Index).Value))
                                    JmlKurang = 0
                                Else
                                    JmlKurang -= Me.vrowItemDelivery.delivery_qty_multi
                                    Me.vrowItemDelivery.delivery_qty_multi = 0
                                End If
                                If Me.vrowItemDelivery.delivery_qty_multi = 0 Then
                                    Me.vTBLItemDelivery.Rows.Remove(Me.vrowItemDelivery)
                                End If
                            End If
                        End While
                    ElseIf CDec(Me.buffer) < CDec(Me.dgvItem.Item(Me.ColQtyMulti.Index, e.RowIndex).Value) Then
                        'untuk item delivery, jika qty naik, maka tambahkan sisanya pada baris paling bawah
                        JmlKurang = CDec(Me.dgvItem.Item(Me.ColQtyMulti.Index, e.RowIndex).Value) - CDec(Me.buffer)
                        Me.vTBLItemDelivery.DefaultView.RowFilter = "identifier = '" + Me.dgvItem.Item(Me.ColIdentifier.Index, e.RowIndex).Value + "'"
                        Me.vTBLItemDelivery.DefaultView.Sort = "[delivery date] DESC"
                        If Me.vTBLItemDelivery.DefaultView.Count > 0 Then
                            Me.vrowItemDelivery = Me.vTBLItemDelivery.DefaultView(0).Row
                            Me.vrowItemDelivery.delivery_qty_multi += JmlKurang
                            Me.vrowItemDelivery.delivery_qty = Me.GetQtyDefault(Me.vrowItemDelivery.delivery_qty_multi, _
                            row.Cells(Me.ColRatioDefault.Index).Value, row.Cells(Me.ColRatioReference.Index).Value, _
                            Me.ObjProduct.GetAngkaDecimal(row.Cells(Me.ColPrimaryCodeNo.Index).Value))
                        End If
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahEditSahBooking(ByVal row As DataGridViewRow, _
    ByVal PrimaryProduct As Decimal) As Boolean
        Dim ObjItemBookingGIT As New ClsTransaksiBookingGITitem
        Dim ObjItemBookingInv As New ClsTransaksiBookingInvItem
        Dim mdec As Integer = 0

        Try
            With row
                Dim JmlDGVTotal As Decimal = Me.GetJmlDGV(PrimaryProduct, .Cells(Me.ColUnitPrice.Index).Value, .Cells(Me.ColNettPrice.Index).Value, .Cells(Me.ColDiscount.Index).Value)
                Dim JmlBooking As Decimal = ObjItemBookingGIT.GetQtyByPrimarySO_n_Product_Multi(Me.TxtPrimarySO.Text, PrimaryProduct) + _
                ObjItemBookingInv.GetQtyByPrimarySO_n_Product_Multi(PrimaryProduct, Me.TxtPrimarySO.Text)

                mdec = Me.ObjProduct.GetAngkaDecimal(row.Cells(Me.ColPrimaryCodeNo.Index).Value, row.Cells(Me.ColPrimaryUnitsMulti.Index).Value)

                If JmlDGVTotal < JmlBooking Then
                    If MsgBox("Booking quantity refering to this item is " + ObjCD.DigitGrouping(CStr(JmlBooking), mdec) + " " + _
                    CStr(.Cells(Me.ColUnit.Index).Value) + " editting to " + ObjCD.DigitGrouping(CStr(JmlDGVTotal), mdec) + " " + CStr(.Cells(Me.ColUnitsMulti.Index).Value) + _
                    " will cancel the booking, continue?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                        Return False
                    Else
                        Return True
                    End If
                Else
                    Return True
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjItemBookingGIT = Nothing
            ObjItemBookingInv = Nothing
        End Try
    End Function

    Private Function ApakahEditSah(ByVal row As DataGridViewRow, ByVal PrimaryProduct As Double) As Boolean
        Try
            Dim JmlOrderBalance, JmlObjLama, JmlTerkirim, JmlDGVTotal, JmlPendingDO, JmlBookingInv As Decimal
            Dim vPrimaryProduct, vPrimarySO As Double
            vPrimaryProduct = row.Cells(Me.ColPrimaryCodeNo.Index).Value
            vPrimarySO = Me.TxtPrimarySO.Text

            With row
                JmlDGVTotal = Me.GetJmlDGV(PrimaryProduct, .Cells(Me.ColUnitPrice.Index).Value, .Cells(Me.ColNettPrice.Index).Value, .Cells(Me.ColDiscount.Index).Value)
                JmlOrderBalance = Me.ObjSOItem.GetSalesOrderBalance_Multi(Me.TxtPrimarySO.Text, PrimaryProduct)
                JmlObjLama = Me.ObjSOItem.GetJml_Multi(PrimaryProduct, Me.TxtPrimarySO.Text)
                JmlPendingDO = Me.ObjDOItem.GetJmlPendingDO_Multi(vPrimaryProduct, vPrimarySO)
                JmlBookingInv = Me.ObjBookingInv.GetJml6_Multi(vPrimaryProduct, vPrimarySO)
                'JmlTerkirim = JmlPendingDO + JmlBookingInv
                JmlTerkirim = JmlObjLama - JmlOrderBalance + JmlBookingInv

                If JmlTerkirim > JmlDGVTotal Then
                    MsgBox("Minimum value you must can enter is " + Me.ObjCD.DigitGrouping(CStr(JmlTerkirim), Me.ObjProduct.GetAngkaDecimal(vPrimaryProduct, row.Cells(Me.ColPrimaryUnitsMulti.Index).Value)) + _
                    " " + row.Cells(Me.ColUnitsMulti.Index).Value + ".", MsgBoxStyle.Information, "Information")
                    Return False
                Else
                    Return True
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub EndEditUnitPrice(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal row As DataGridViewRow)

        Try
            With row
                'untuk nilai unit price, pastikan nilainya positif
                If Me.ObjCD.CekBilanganPositif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("Please enter positive number for unit price.", MsgBoxStyle.Information, _
                    "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
                'digit grouping
                Me.dgvItem(e.ColumnIndex, e.RowIndex).Value = _
                Me.ObjCD.DigitGrouping(Me.dgvItem(e.ColumnIndex, e.RowIndex).Value, DefaultPrecision)

                If CDec(Me.buffer) <> CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) Then
                    'jika unit price < nett price maka kembalikan ke nilai buffer
                    'hanya berlaku klo menentukan diskon berdasarkan amount
                    If Me.dgvItem.Item(Me.ColDiscFromPerct.Index, e.RowIndex).Value = "0" Then
                        If CDec(Me.dgvItem.Item(Me.ColDiscAmount.Index, e.RowIndex).Value) > CDec(Me.dgvItem.Item(Me.ColUnitPrice.Index, e.RowIndex).Value) Then
                            Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                            Exit Sub
                        End If
                    End If
                    'hitung ulang row
                    Me.Hitung_Amount(row)
                    'hitung total
                    Me.Hitung_Total()
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditDiskonAmount(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
ByVal row As DataGridViewRow)
        Try
            With row
                'untuk nilai disc amount, pastikan nilainya positif
                If Me.ObjCD.CekNonNegatif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("Please enter positive number for discount amount.", MsgBoxStyle.Information, _
                    "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
                'digit grouping
                Me.dgvItem(e.ColumnIndex, e.RowIndex).Value = _
                Me.ObjCD.DigitGrouping(Me.dgvItem(e.ColumnIndex, e.RowIndex).Value, DefaultPrecision)

                If CDec(Me.buffer) <> CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) Then
                    'jika unit price < nett price maka kembalikan ke nilai buffer
                    'hanya berlaku klo menentukan diskon berdasarkan amount
                    If CDec(Me.dgvItem.Item(Me.ColDiscAmount.Index, e.RowIndex).Value) > CDec(Me.dgvItem.Item(Me.ColUnitPrice.Index, e.RowIndex).Value) Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                    'tandai diskon dihitung berdasarkan amount
                    row.Cells(Me.ColDiscFromPerct.Index).Value = "0"
                    'hitung ulang row
                    Me.Hitung_Amount(row)
                    'hitung total
                    Me.Hitung_Total()
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditDiscount(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal row As DataGridViewRow)
        Try
            With row
                'untuk nilai diskon, pastikan nilainya tidak negatif
                If Me.ObjCD.CekNonNegatif(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    MsgBox("Please enter positive number for discount.", MsgBoxStyle.Information, _
                    "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
                'cek nili maksimum untuk diskon adalah 100%
                If CDbl(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) > 100 Then
                    MsgBox("Maximum value is 100%.", MsgBoxStyle.Information, _
                    "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
                'digit grouping
                Me.dgvItem(e.ColumnIndex, e.RowIndex).Value = _
                Me.ObjCD.DigitGrouping(Me.dgvItem(e.ColumnIndex, e.RowIndex).Value, 2)
                'hanya hitung ulang kalo ada perubahan data
                If CDec(Me.buffer) <> CDec(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) Then
                    'tandai di ubah dari diskon
                    row.Cells(Me.ColDiscFromPerct.Index).Value = "1"
                    'hitung ulang row
                    Me.Hitung_Amount(row)
                    'hitung total
                    Me.Hitung_Total()
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvItem.DataError
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_RowsAdded(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItem.RowsAdded
        Me.dgvItem.TabStop = True
        If Me.txtRemarks.Enabled = True Then
            Me.btnUP.Enabled = True
            Me.btnDown.Enabled = True
        End If
    End Sub

    Private Sub dgvItem_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvItem.RowsRemoved
        Try
            If Me.dgvItem.Rows.Count = 0 Then
                Me.dgvItem.TabStop = False
                If Me.txtRemarks.Enabled = True Then
                    Me.btnUP.Enabled = False
                    Me.btnDown.Enabled = False
                End If
            End If
            If Me.txtRemarks.Enabled = True Then
                If HitungOnRowsRemoved = True Then
                    Me.Hitung_Total()
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvItem.SelectionChanged
        Try
            If Me.txtRemarks.Enabled = True Then
                If Me.dgvItem.SelectedRows.Count > 0 Then
                    Me.btnUP.Enabled = True
                    Me.btnDown.Enabled = True
                Else
                    Me.btnUP.Enabled = False
                    Me.btnDown.Enabled = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_Sorted(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dgvItem.Sorted
        Me.ObjCD.SetNumber(Me.dgvItem)
    End Sub

    ''' <summary>
    ''' tampilkan row yang diberikan pada parameter
    ''' </summary>
    ''' <param name="row"></param>
    ''' <remarks></remarks>
    Private Sub ShowRow(ByVal row As DataGridViewRow, ByVal Up As Boolean)
        Try
            'tentukan arah perpindahan, ke atas atau ke bawah
            If row.Index < Me.dgvItem.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke atas
                While row.Displayed = False
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex - 1
                End While
            ElseIf row.Index > Me.dgvItem.FirstDisplayedScrollingRowIndex Then
                'perpindahan ke bawah
                While Me.dgvItem.Rows(row.Index).Displayed = False
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                End While
                'kalo sampe sini pasti row yang disorot adalah row yang paling bawah, pastikan semuanya terlihat
                If Me.dgvItem.DisplayedRowCount(True) <> Me.dgvItem.DisplayedRowCount(False) Then
                    Me.dgvItem.FirstDisplayedScrollingRowIndex = Me.dgvItem.FirstDisplayedScrollingRowIndex + 1
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnUP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUP.Click
        Try
            If Me.dgvItem.SelectedRows.Count > 0 Then
                Me.HitungOnRowsRemoved = False
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
                Dim row As DataGridViewRow = Me.dgvItem.SelectedRows(0)
                If row.Index <> 0 Then
                    Dim indexbaru As Integer = row.Index - 1
                    Me.dgvItem.Rows.Remove(row)
                    Me.dgvItem.Rows.InsertRange(indexbaru, row)
                    Me.dgvItem.Focus()
                    Me.dgvItem.ClearSelection()
                    row.Selected = True
                    Me.ObjCD.SetNumber(Me.dgvItem)
                    If row.Displayed = False Then
                        Me.ShowRow(row, True)
                    End If
                End If
                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            End If
            Me.HitungOnRowsRemoved = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.HitungOnRowsRemoved = True
        End Try
    End Sub

    Private Sub btnDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDown.Click
        Try
            If Me.dgvItem.SelectedRows.Count > 0 Then
                Me.HitungOnRowsRemoved = False
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
                Dim row As DataGridViewRow = Me.dgvItem.SelectedRows(0)
                If row.Index <> Me.dgvItem.Rows.Count - 1 Then
                    Dim indexbaru As Integer = row.Index + 1
                    Me.dgvItem.Rows.Remove(row)
                    Me.dgvItem.Rows.InsertRange(indexbaru, row)
                    Me.dgvItem.Focus()
                    Me.dgvItem.ClearSelection()
                    row.Selected = True
                    Me.ObjCD.SetNumber(Me.dgvItem)
                    Me.ShowRow(row, False)
                End If
                Me.btnUP.Enabled = True
                Me.btnDown.Enabled = True
            End If
            Me.HitungOnRowsRemoved = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.HitungOnRowsRemoved = True
        End Try
    End Sub

    Private Sub dgvItem_UserDeletedRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvItem.UserDeletedRow
        Try
            Me.ObjCD.SetNumber(Me.dgvItem)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRemarks.GotFocus
        Try
            Me.txtRemarks.BackColor = BackColorAll
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtRemark_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRemarks.LostFocus
        Try
            Me.txtRemarks.BackColor = Me.TxtPrimaryCodeNo.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSalesOrderNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSalesOrderNo.GotFocus
        If Me.TxtSalesOrderNo.ReadOnly = False And Me.TxtSalesOrderNo.Enabled = True Then
            TxtSalesOrderNo.BackColor = BackColorAll
        End If
    End Sub

    Private Sub TxtSalesOrderNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSalesOrderNo.LostFocus
        If Me.TxtSalesOrderNo.ReadOnly = False And Me.TxtSalesOrderNo.Enabled = True Then
            TxtSalesOrderNo.BackColor = Me.txtPrimaryCategory.BackColor
        End If
    End Sub

    Private Sub txtSupplierCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustomerCode.GotFocus
        txtCustomerCode.BackColor = BackColorAll
    End Sub

    Private Sub txtSupplierCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustomerCode.LostFocus
        txtCustomerCode.BackColor = Me.txtPrimaryCategory.BackColor
    End Sub

    Private Sub txtContactPerson_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtContactPerson.GotFocus
        txtContactPerson.BackColor = BackColorAll
    End Sub

    Private Sub txtContactPerson_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtContactPerson.LostFocus
        txtContactPerson.BackColor = Me.txtPrimaryCategory.BackColor
    End Sub

    Private Sub txtPhoneNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPhoneNo.GotFocus
        txtPhoneNo.BackColor = BackColorAll
    End Sub

    Private Sub txtPhoneNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPhoneNo.LostFocus
        txtPhoneNo.BackColor = Me.txtPrimaryCategory.BackColor
    End Sub

    Private Sub TxtShipVia_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtShipVia.GotFocus
        If Me.TxtShipVia.Enabled = True And Me.TxtShipVia.ReadOnly = False Then
            TxtShipVia.BackColor = BackColorAll
        End If
    End Sub

    Private Sub TxtShipVia_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtShipVia.LostFocus
        If Me.TxtShipVia.Enabled = True And Me.TxtShipVia.ReadOnly = False Then
            TxtShipVia.BackColor = Me.txtPrimaryCategory.BackColor
        End If
    End Sub

    Private Sub txtCurrencyCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCurrencyCode.GotFocus
        txtCurrencyCode.BackColor = BackColorAll
    End Sub

    Private Sub txtCurrencyCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCurrencyCode.LostFocus
        txtCurrencyCode.BackColor = Me.txtPrimaryCategory.BackColor
    End Sub

    Private Sub txtPaymentTerm_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPaymentTerm.GotFocus
        txtPaymentTerm.BackColor = BackColorAll
    End Sub

    Private Sub txtPaymentTerm_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPaymentTerm.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.txtPaymentTerm, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtPaymentTerm_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPaymentTerm.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.txtPaymentTerm, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtPaymentTerm_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPaymentTerm.LostFocus
        txtPaymentTerm.BackColor = Me.txtPrimaryCategory.BackColor
    End Sub

    Private Sub txtPaymentDueDate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPaymentDueDate.GotFocus
        txtPaymentDueDate.BackColor = BackColorAll
    End Sub

    Private Sub txtPaymentDueDate_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPaymentDueDate.LostFocus
        txtPaymentDueDate.BackColor = Me.txtPrimaryCategory.BackColor
    End Sub

    Private Sub TxtCodeNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNo.GotFocus
        Try
            If Me.TxtCodeNo.ReadOnly = False Then
                TxtCodeNo.BackColor = BackColorAll
            Else
                TxtCodeNo.BackColor = Me.TxtSalesOrderNo.BackColor
            End If
            If Me.txtPrimaryCustomer.Text <> "" Then
                If Me.txtPrimaryCustomer.Text = -1 And Me.txtPrimaryCustomer.Text <> "" Then
                    Me.TabControl1.SelectedIndex = 0
                    Me.txtCustomerCode.Focus()
                    MsgBox("To start entering item data, you must specify customer data first.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            If Me.chkUseCustCodeNo.Checked = False Then
                If Me.txtPrimaryCurrency.Text <> "" Then
                    If Me.txtPrimaryCurrency.Text = -1 And Me.txtPrimaryCurrency.Text <> "" Then
                        Me.TabControl1.SelectedIndex = 1
                        Me.txtCurrencyCode.Focus()
                        MsgBox("To start entering item data, you must specify currency data first.", MsgBoxStyle.Information, "Information")
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCodeNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCodeNo.LostFocus
        If Me.TxtCodeNo.ReadOnly = False Then
            TxtCodeNo.BackColor = Me.txtPrimaryCategory.BackColor
        Else
            Me.TxtCodeNo.BackColor = Me.TxtSalesOrderNo.BackColor
        End If
    End Sub

    Private Sub TxtQtyperBatch_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchDefault.GotFocus
        TxtQtyperBatchDefault.BackColor = BackColorAll
    End Sub

    Private Sub TxtQtyperBatch_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchDefault.LostFocus
        TxtQtyperBatchDefault.BackColor = Me.txtPrimaryCategory.BackColor
    End Sub

    Private Sub TxtUnitPrice_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitPrice.Enter
        Try
            Me.buffer = Me.TxtUnitPrice.Text
            Me.ObjCD.Number_Enter(Me.TxtUnitPrice, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUnitPrice_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitPrice.GotFocus
        TxtUnitPrice.BackColor = BackColorAll
    End Sub

    Private Sub TxtUnitPrice_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitPrice.LostFocus
        TxtUnitPrice.BackColor = Me.txtPrimaryCategory.BackColor
    End Sub

    Private Sub TxtDiscount_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscount.Enter
        Try
            Me.buffer = Me.TxtDiscount.Text
            Me.ObjCD.Number_Enter(Me.TxtDiscount, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscount_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscount.GotFocus
        If Me.TxtDiscount.ReadOnly = False And Me.TxtDiscount.Enabled = True Then
            TxtDiscount.BackColor = BackColorAll
        End If
    End Sub

    Private Sub TxtDiscount_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscount.LostFocus
        If Me.TxtDiscount.ReadOnly = False And Me.TxtDiscount.Enabled = True Then
            TxtDiscount.BackColor = Me.txtPrimaryCategory.BackColor
        End If
    End Sub

    Private Sub TxtRemarkItem_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarkItem.GotFocus
        TxtRemarkItem.BackColor = BackColorAll
    End Sub

    Private Sub TxtRemarkItem_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtRemarkItem.LostFocus
        TxtRemarkItem.BackColor = Me.txtPrimaryCategory.BackColor
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbDetailStock.LinkClicked
        Dim hj As New FrmTrans_SalesDashboard
        Try
            If Me.TxtCodeNo.Text <> "" Then
                hj.txtCodeNo.Text = Me.TxtCodeNo.Text
            End If
            hj.UserCode = Me.UserCode
            hj.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            hj.Dispose()
        End Try
    End Sub

    Private Sub txtCustomerProductCodeNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustomerProductCodeNo.GotFocus
        Try
            If Me.txtCustomerProductCodeNo.ReadOnly = False And Me.txtCustomerProductCodeNo.Enabled = True Then
                Me.txtCustomerProductCodeNo.BackColor = BackColorAll
            End If
            If Me.txtPrimaryCustomer.Text <> "" Then
                If Me.txtPrimaryCustomer.Text = -1 And Me.txtPrimaryCustomer.Text <> "" Then
                    Me.TabControl1.SelectedIndex = 0
                    Me.txtCustomerCode.Focus()
                    MsgBox("To start entering item data, you must specify customer data first.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            If Me.chkUseCustCodeNo.Checked = True Then
                If Me.txtPrimaryCurrency.Text <> "" Then
                    If Me.txtPrimaryCurrency.Text = -1 And Me.txtPrimaryCurrency.Text <> "" Then
                        Me.TabControl1.SelectedIndex = 1
                        Me.txtCurrencyCode.Focus()
                        MsgBox("To start entering item data, you must specify currency data first.", MsgBoxStyle.Information, "Information")
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCustomerProductCodeNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCustomerProductCodeNo.KeyDown
        Dim i As New FrmCariCustomerProductCodeNo
        i.BackColorAll = Me.BackColorAll
        Dim Jml As Long = 0

        Try
            If e.KeyValue = 13 And Me.chkUseCustCodeNo.Checked = True Then
                Jml = Me.ObjCustomer_sProductCode.GetJml(Me.txtCustomerProductCodeNo.Text, _
                Me.txtPrimaryCustomer.Text)
                If Jml = 1 Then
                    Me.TxtPrimaryCodeNo.Text = Me.ObjCustomer_sProductCode.GetPrimary(Me.txtCustomerProductCodeNo.Text, _
                    Me.txtPrimaryCustomer.Text)
                    Me.txtPrimaryCategory.Text = Me.ObjProduct.GetPrimaryCategory(Me.TxtPrimaryCodeNo.Text)
                    Me.TxtCodeNo.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryCodeNo.Text)
                    Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                    Me.ObjCustomer_sProductCode.IsiCustomerProductCodeNDesc(Me.txtCustomerProductCodeNo, _
                    Me.txtCustomerProductDesc, Me.txtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text)
                    Me.dtpdeliverydate.Focus()
                    Me.TxtDiscFromPerct.Text = "1"
                    Me.SetGbo1()
                ElseIf Jml > 1 Or Jml = 0 Then
                    'ada 2 product dengan kode yang sama atau tidak ada product yang dimaksud
                    If Me.txtCustomerProductCodeNo.Text <> "" Then
                        i.txtCari.Text = Me.txtCustomerProductCodeNo.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryProduct > 0 Then
                        Me.TxtPrimaryCodeNo.Text = i.PrimaryProduct
                        Me.txtPrimaryCategory.Text = Me.ObjProduct.GetPrimaryCategory(Me.TxtPrimaryCodeNo.Text)
                        Me.TxtCodeNo.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryCodeNo.Text)
                        Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                        Me.ObjCustomer_sProductCode.IsiCustomerProductCodeNDesc(Me.txtCustomerProductCodeNo, _
                        Me.txtCustomerProductDesc, Me.txtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text)
                        Me.dtpdeliverydate.Focus()
                        Me.TxtDiscFromPerct.Text = "1"
                        Me.SetGbo1()
                    Else
                        Me.txtCustomerProductCodeNo.Clear()
                        Me.TxtPrimaryCodeNo.Text = "-1"
                        Me.txtPrimaryCategory.Text = "-1"
                        Me.KosongkanItem()
                        Me.txtCustomerProductCodeNo.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtCustomerProductCodeNo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustomerProductCodeNo.Leave
        Dim i As New FrmCariCustomerProductCodeNo
        i.BackColorAll = Me.BackColorAll

        Try
            If Me.txtPrimaryCustomer.Text <> "" Then
                i.PrimaryCust = Me.txtPrimaryCustomer.Text
                If Me.txtCustomerProductCodeNo.Text <> "" Then
                    If Me.ObjCustomer_sProductCode.GetJml(Me.txtCustomerProductCodeNo.Text, Me.txtPrimaryCustomer.Text) = 1 Then
                        Me.TxtPrimaryCodeNo.Text = Me.ObjCustomer_sProductCode.GetPrimary(Me.txtCustomerProductCodeNo.Text, _
                        Me.txtPrimaryCustomer.Text)
                        Me.txtPrimaryCategory.Text = Me.ObjProduct.GetPrimaryCategory(Me.TxtPrimaryCodeNo.Text)
                        Me.TxtCodeNo.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryCodeNo.Text)
                        Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                        Me.ObjCustomer_sProductCode.IsiCustomerProductCodeNDesc(Me.txtCustomerProductCodeNo, _
                        Me.txtCustomerProductDesc, Me.txtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text)
                        Me.dtpdeliverydate.Focus()
                        Me.TxtDiscFromPerct.Text = "1"
                        Me.SetGbo1()
                    Else
                        If Me.txtCustomerProductCodeNo.Text <> "" Then
                            i.txtCari.Text = Me.txtCustomerProductCodeNo.Text
                        End If

                        i.ShowDialog()
                        If i.PrimaryProduct > 0 Then
                            Me.TxtPrimaryCodeNo.Text = i.PrimaryProduct
                            Me.txtPrimaryCategory.Text = Me.ObjProduct.GetPrimaryCategory(Me.TxtPrimaryCodeNo.Text)
                            Me.TxtCodeNo.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryCodeNo.Text)
                            Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                            Me.ObjCustomer_sProductCode.IsiCustomerProductCodeNDesc(Me.txtCustomerProductCodeNo, _
                            Me.txtCustomerProductDesc, Me.txtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text)
                            Me.dtpdeliverydate.Focus()
                            Me.TxtDiscFromPerct.Text = "1"
                            Me.SetGbo1()
                        Else
                            Me.txtCustomerProductCodeNo.Clear()
                            Me.TxtPrimaryCodeNo.Text = "-1"
                            Me.txtPrimaryCategory.Text = "-1"
                            Me.KosongkanItem()
                            Me.txtCustomerProductCodeNo.Focus()
                        End If
                    End If
                Else
                    Me.TxtPrimaryCodeNo.Text = "-1"
                    Me.txtDesc.Clear()
                    Me.txtPrimaryCategory.Text = "-1"
                    Me.KosongkanItem()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindCustomerProductCode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindCustomerProductCode.Click
        Dim i As New FrmCariCustomerProductCodeNo
        i.BackColorAll = Me.BackColorAll

        Try
            'cek apakah sudah isi kode customer dan kode currency
            If Me.txtPrimaryCustomer.Text <> "" Then
                If Me.txtPrimaryCustomer.Text = -1 And Me.txtPrimaryCustomer.Text <> "" Then
                    Me.TabControl1.SelectedIndex = 0
                    Me.txtCustomerCode.Focus()
                    MsgBox("To start entering item data, you must specify customer data first.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            If Me.txtPrimaryCurrency.Text <> "" Then
                If Me.txtPrimaryCurrency.Text = -1 And Me.txtPrimaryCurrency.Text <> "" Then
                    Me.TabControl1.SelectedIndex = 1
                    Me.txtCurrencyCode.Focus()
                    MsgBox("To start entering item data, you must specify currency data first.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If

            i.cboCriteria.SelectedIndex = 0
            If Me.txtCustomerProductCodeNo.Text <> "" Then
                i.txtCari.Text = Me.txtCustomerProductCodeNo.Text
            End If
            i.PrimaryCust = Me.txtPrimaryCustomer.Text
            i.ShowDialog()
            If i.PrimaryProduct > 0 Then
                Me.TxtPrimaryCodeNo.Text = i.PrimaryProduct
                Me.txtPrimaryCategory.Text = Me.ObjProduct.GetPrimaryCategory(Me.TxtPrimaryCodeNo.Text)
                Me.TxtCodeNo.Text = Me.ObjProduct.GetProductCode(Me.TxtPrimaryCodeNo.Text)
                Me.txtDesc.Text = Me.ObjProduct.GetDescription(Me.TxtPrimaryCodeNo.Text)
                Me.ObjCustomer_sProductCode.IsiCustomerProductCodeNDesc(Me.txtCustomerProductCodeNo, _
                Me.txtCustomerProductDesc, Me.txtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text)
                Me.dtpdeliverydate.Focus()
                Me.TxtDiscFromPerct.Text = "1"
                Me.SetGbo1()
            Else
                Me.txtCustomerProductCodeNo.Clear()
                Me.TxtPrimaryCodeNo.Text = "-1"
                Me.txtPrimaryCategory.Text = "-1"
                Me.KosongkanItem()
                Me.txtCustomerProductCodeNo.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub Hitung_Amount(ByVal row As DataGridViewRow)
        Dim DiscPerct As Integer = 0

        Try
            'hitung nett price
            DiscPerct = row.Cells(Me.ColDiscFromPerct.Index).Value
            If DiscPerct = 1 Then
                row.Cells(Me.ColDiscAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal( _
                ((CDec(row.Cells(Me.ColUnitPrice.Index).Value) * CDec(row.Cells(Me.ColDiscount.Index).Value))) / 100, DefaultPrecision))
            Else
                row.Cells(Me.ColDiscount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal( _
                ((CDec(row.Cells(Me.ColDiscAmount.Index).Value) / CDec(row.Cells(Me.ColUnitPrice.Index).Value))) * 100, 2))
            End If
            row.Cells(Me.ColNettPrice.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal( _
            CDec(row.Cells(Me.ColUnitPrice.Index).Value) - CDec(row.Cells(Me.ColDiscAmount.Index).Value) _
            , DefaultPrecision))

            'hitung amount
            If row.Cells(Me.ColUseSecondary.Index).Value = Boolean.FalseString Then
                If row.Cells(Me.ColPriceByMultiUnit.Index).Value = 1 Then
                    row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(row.Cells(Me.ColQtyMulti.Index).Value * _
                    row.Cells(Me.ColNettPrice.Index).Value, DefaultPrecision))
                Else
                    row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(row.Cells(Me.ColQty.Index).Value * _
                    row.Cells(Me.ColNettPrice.Index).Value, DefaultPrecision))
                End If
            Else
                row.Cells(Me.ColAmount.Index).Value = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Hitung_Amount()
        Try
            If Me.ChkDiscountConditionSama.Checked = True Then
                If IsNumeric(Me.txtReferenceDiscount.Text) = False Then
                    Me.txtReferenceDiscount.Text = Me.ObjCD.DigitGrouping(0, 2)
                End If
                Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Me.txtReferenceDiscount.Text, 2) 
            End If

            If Me.TxtDiscFromPerct.Text = "" Then
                Me.TxtDiscFromPerct.Text = "1"
            End If
            If Me.TxtDiscFromPerct.Text = "1" Then
                'hitung disc amount
                Me.TxtDiscAmount.Text = Me.ObjCD.DigitGrouping((CDec(Me.TxtUnitPrice.Text) * CDec(Me.TxtDiscount.Text)) / 100, DefaultPrecision)
            Else
                'hitung disc %
                Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping((CDec(Me.TxtDiscAmount.Text) * 100) / CDec(Me.TxtUnitPrice.Text), 2)
            End If
            'hitung nett price
            Me.txtNettPrice.Text = Me.ObjCD.DigitGrouping(CDec(Me.TxtUnitPrice.Text) - CDec(Me.TxtDiscAmount.Text), _
            DefaultPrecision)

            'hitung amount
            If Me.TxtQtyperBatchMulti.Text = "" Then
                Me.TxtQtyperBatchMulti.Text = "0"
            End If
            If Me.chkUseSecondary.Checked = False Then
                If Me.txtPriceByMultiUnit.Text = "" Then
                    Me.txtPriceByMultiUnit.Text = "1"
                End If
                If Me.txtPriceByMultiUnit.Text = "1" Then
                    Me.txtAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDbl(Me.txtNettPrice.Text) * _
                    CDbl(Me.TxtQtyperBatchMulti.Text), DefaultPrecision))
                Else
                    Me.txtAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDbl(Me.txtNettPrice.Text) * _
                    CDbl(Me.TxtQtyperBatchDefault.Text), DefaultPrecision))
                End If
            Else
                Me.txtAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(0, DefaultPrecision))
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Hitung_Total()
        Dim row As DataGridViewRow
        Dim SubTotal As Decimal = 0

        Try
            '2025 Core Tax Tahap 2.1
            Dim ObjSetNPWP As New Core_Module.ClsTrans_SettingNPWPNo
            If Me.dtpDate.Value >= #1/1/2025# Then
                Me.Hitung_Total_MiddlePoint()
                Exit Sub
            End If
            'hitung sub total
            For Each row In Me.dgvItem.Rows
                'yang diperhitungkan hanya yang tidak pake secondary unit
                If row.Cells(Me.ColUseSecondary.Index).Value = Boolean.FalseString Then
                    If row.Cells(Me.ColPriceCond.Index).Value = "Include" Then
                        SubTotal += Math.Round((row.Cells(Me.ColAmount.Index).Value * (100 / (CDbl(Me.txtTax.Text) + 100))), _
                        Me.DefaultPrecision, MidpointRounding.AwayFromZero)
                    Else
                        SubTotal += CDbl(row.Cells(Me.ColAmount.Index).Value)
                    End If
                End If
            Next
            Me.txtSubTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal, DefaultPrecision))
            'hitung pajak
            If Me.chkTax.Checked = True Then
                If Me.txtPrimaryCurrency.Text = "1" Then
                    Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(Me.txtTax.Text * SubTotal / 100, 0, ClsCekData.EnumTipeRounding.Down), _
                    DefaultPrecision)
                Else
                    Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.txtTax.Text * SubTotal / 100, DefaultPrecision, ClsCekData.EnumTipeRounding.Down)
                End If
            Else
                Me.txtTax.Text = "0.00"
                Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", Me.DefaultPrecision)
            End If
            'hitung total = sub total + tax
            Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDbl(Me.txtSubTotal.Text) + _
            CDbl(Me.txtTaxAmount.Text), DefaultPrecision))

            Me.Hitung_DP()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Hitung_Total_MiddlePoint()
        Dim row As DataGridViewRow
        Dim SubTotal As Decimal = 0

        Try
            'hitung sub total
            For Each row In Me.dgvItem.Rows
                'yang diperhitungkan hanya yang tidak pake secondary unit
                If row.Cells(Me.ColUseSecondary.Index).Value = Boolean.FalseString Then
                    If row.Cells(Me.ColPriceCond.Index).Value = "Include" Then
                        SubTotal += Math.Round((row.Cells(Me.ColAmount.Index).Value * (100 / (CDbl(Me.txtTax.Text) + 100))), _
                        12, MidpointRounding.AwayFromZero)
                    Else
                        SubTotal += CDbl(row.Cells(Me.ColAmount.Index).Value)
                    End If
                End If
            Next
            Me.txtSubTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal, DefaultPrecision))
            'hitung pajak
            If Me.chkTax.Checked = True Then
                If Me.txtPrimaryCurrency.Text = "1" Then
                    Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(Me.txtTax.Text * SubTotal / 100, 0, ClsCekData.EnumTipeRounding.Middlepoint), _
                    DefaultPrecision)
                Else
                    Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.txtTax.Text * SubTotal / 100, DefaultPrecision, ClsCekData.EnumTipeRounding.Middlepoint)
                End If
            Else
                Me.txtTax.Text = "0.00"
                Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping("0", Me.DefaultPrecision)
            End If
            'hitung total = sub total + tax
            Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDbl(Me.txtSubTotal.Text) + _
            CDbl(Me.txtTaxAmount.Text), DefaultPrecision))

            Me.Hitung_DP()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


    Private Sub Hitung_DP()
        Try
            If Me.txtRemarks.Enabled = True Then
                If Me.TxtDPFromPercent.Text = "" Then
                    Me.TxtDPFromPercent.Text = "1"
                End If
                If Me.TxtDPPercent.Text = "" Then
                    Me.TxtDPPercent.Text = "0"
                End If
                If Me.TxtDPAmount.Text = "" Then
                    Me.TxtDPAmount.Text = "0"
                End If
                If Me.TxtDPFromPercent.Text = "1" Then
                    'hitung disc amount
                    Me.TxtDPAmount.Text = Me.ObjCD.DigitGrouping((CDec(Me.txtSubTotal.Text) * CDec(Me.TxtDPPercent.Text)) / 100, DefaultPrecision)
                Else
                    'hitung disc %
                    If CDec(Me.txtSubTotal.Text) = 0 Then
                        Me.TxtDPPercent.Text = Me.ObjCD.DigitGrouping(0, 2)
                    Else
                        Me.TxtDPPercent.Text = Me.ObjCD.DigitGrouping((CDec(Me.TxtDPAmount.Text) * 100) / CDec(Me.txtSubTotal.Text), 2)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCurrencyCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCurrencyCode.KeyDown
        Dim i As New FrmCariMasterCurrency
        i.BackColorAll = Me.BackColorAll
        Try
            If e.KeyValue = 13 Then
                If ObjCurrency.ApakahSudahAda(Me.txtCurrencyCode.Text) = True Then
                    Me.txtPrimaryCurrency.Text = ObjCurrency.GetPrimary(Me.txtCurrencyCode.Text)
                    Me.SetDefaultPrecision1()
                    Me.chkTax.Focus()
                Else
                    If Me.txtCurrencyCode.Text <> "" Then
                        i.txtCari.Text = Me.txtCurrencyCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCurrency > 0 Then
                        Me.txtPrimaryCurrency.Text = i.PrimaryCurrency
                        Me.txtCurrencyCode.Text = ObjCurrency.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                        Me.SetDefaultPrecision1()
                        Me.chkTax.Focus()
                    Else
                        Me.txtCurrencyCode.Clear()
                        Me.txtPrimaryCurrency.Text = "-1"
                        Me.txtCurrencyCode.Focus()
                        Me.SetDefaultPrecision1()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtCurrencyCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCurrencyCode.Leave
        Dim i As New FrmCariMasterCurrency
        i.BackColorAll = Me.BackColorAll
        Dim Rate As Decimal = 0

        Try
            If Me.txtCurrencyCode.Text <> "" Then
                If ObjCurrency.ApakahSudahAda(Me.txtCurrencyCode.Text) = True Then
                    'Me.chkTax.Focus()
                    Me.txtPrimaryCurrency.Text = ObjCurrency.GetPrimary(Me.txtCurrencyCode.Text)
                    Me.SetDefaultPrecision1()
                Else
                    If Me.txtCurrencyCode.Text <> "" Then
                        i.txtCari.Text = Me.txtCurrencyCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCurrency > 0 Then
                        Me.txtPrimaryCurrency.Text = i.PrimaryCurrency
                        Me.txtCurrencyCode.Text = ObjCurrency.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                        Me.SetDefaultPrecision1()
                        ' Me.chkTax.Focus()
                    Else
                        Me.txtCurrencyCode.Clear()
                        Me.txtPrimaryCurrency.Text = "-1"
                        Me.txtCurrencyCode.Focus()
                        Me.SetDefaultPrecision1()
                    End If
                End If
            Else
                Me.txtPrimaryCurrency.Text = "-1"
            End If

            If Me.txtCurrencyCode.Enabled = True Then
                If Me.txtPrimaryCurrency.Text = "-1" Or Me.txtPrimaryCurrency.Text = "" Then
                    Me.TxtExchangeRate.ReadOnly = True
                    Me.TxtExchangeRate.TabStop = False
                    Me.TxtExchangeRate.BackColor = Me.TxtReadOnly.BackColor
                    Me.TxtExchangeRate.Text = "1"
                Else
                    If Me.ObjCurrency.ApakahSudahAda(Me.txtCurrencyCode.Text) = False Then
                        Me.TxtExchangeRate.ReadOnly = True
                        Me.TxtExchangeRate.TabStop = False
                        Me.TxtExchangeRate.BackColor = Me.TxtReadOnly.BackColor
                        Me.TxtExchangeRate.Text = "1"
                    Else
                        If Me.txtPrimaryCurrency.Text = Core_Module.PrimaryBaseCurrency Then
                            Me.TxtExchangeRate.ReadOnly = True
                            Me.TxtExchangeRate.TabStop = False
                            Me.TxtExchangeRate.BackColor = Me.TxtReadOnly.BackColor
                            Me.TxtExchangeRate.Text = "1"
                        Else
                            Me.TxtExchangeRate.ReadOnly = False
                            Me.TxtExchangeRate.TabStop = True
                            Me.TxtExchangeRate.BackColor = Me.txtPrimaryCurrency.BackColor

                            Rate = Me.ObjTaxRate.GetTaxRate(Me.dtpDate.Value.Date, Me.txtPrimaryCurrency.Text)
                            If Rate = 0 Then
                                Rate = Me.ObjCurrency.GetDefaultExhangeRate(Me.txtPrimaryCurrency.Text)
                                If Rate = 0 Then
                                    Rate = 1
                                End If
                            End If
                            If CDec(Me.TxtExchangeRate.Text) = 1 Then
                                Me.TxtExchangeRate.Text = Me.ObjCD.DigitGrouping(Rate, Core_Module.DefPrecExchangeRate)
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindCurrency_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindCurrency.Click
        Dim i As New FrmCariMasterCurrency
        i.BackColorAll = Me.BackColorAll
        Dim Rate As Decimal = 1

        Try
            i.ShowDialog()
            If i.PrimaryCurrency > 0 Then
                Me.txtPrimaryCurrency.Text = i.PrimaryCurrency
                Me.txtCurrencyCode.Text = ObjCurrency.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                Me.SetDefaultPrecision1()
                Me.chkTax.Focus()
            Else
                Me.txtCurrencyCode.Clear()
                Me.txtCurrencyCode.Focus()
                Me.SetDefaultPrecision1()
                Me.txtCurrencyCode.Focus()
            End If

            If Me.txtCurrencyCode.Enabled = True Then
                If Me.txtPrimaryCurrency.Text = "-1" Or Me.txtPrimaryCurrency.Text = "" Then
                    Me.TxtExchangeRate.ReadOnly = True
                    Me.TxtExchangeRate.TabStop = False
                    Me.TxtExchangeRate.BackColor = Me.TxtReadOnly.BackColor
                    Me.TxtExchangeRate.Text = "1"
                    Me.chkTax.Focus()
                Else
                    If Me.ObjCurrency.ApakahSudahAda(Me.txtCurrencyCode.Text) = False Then
                        Me.TxtExchangeRate.ReadOnly = True
                        Me.TxtExchangeRate.TabStop = False
                        Me.TxtExchangeRate.BackColor = Me.TxtReadOnly.BackColor
                        Me.TxtExchangeRate.Text = "1"
                        Me.chkTax.Focus()
                    Else
                        If Me.txtPrimaryCurrency.Text = Core_Module.PrimaryBaseCurrency Then
                            Me.TxtExchangeRate.ReadOnly = True
                            Me.TxtExchangeRate.TabStop = False
                            Me.TxtExchangeRate.BackColor = Me.TxtReadOnly.BackColor
                            Me.TxtExchangeRate.Text = "1"
                            Me.chkTax.Focus()
                        Else
                            Me.TxtExchangeRate.ReadOnly = False
                            Me.TxtExchangeRate.TabStop = True
                            Me.TxtExchangeRate.BackColor = Me.txtPrimaryCurrency.BackColor
                            Me.TxtExchangeRate.Focus()
                            Rate = Me.ObjTaxRate.GetTaxRate(Me.dtpDate.Value.Date, Me.txtPrimaryCurrency.Text)
                            If Rate = 0 Then
                                Rate = Me.ObjCurrency.GetDefaultExhangeRate(Me.txtPrimaryCurrency.Text)
                                If Rate = 0 Then
                                    Rate = 1
                                End If
                            End If
                            If CDec(Me.TxtExchangeRate.Text) = 1 Then
                                Me.TxtExchangeRate.Text = Me.ObjCD.DigitGrouping(Rate, Core_Module.DefPrecExchangeRate)
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub SetDefaultPrecision1()
        Try
            Me.DefaultPrecision = Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text)
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            For Each rowDGV As DataGridViewRow In Me.dgvItem.Rows
                rowDGV.Cells(Me.ColUnitPrice.Index).Value = Me.ObjCD.DigitGrouping(rowDGV.Cells(Me.ColUnitPrice.Index).Value, Me.DefaultPrecision)
                If rowDGV.Cells(Me.ColDiscFromPerct.Index).Value = "1" Then
                    rowDGV.Cells(Me.ColNettPrice.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(rowDGV.Cells(Me.ColUnitPrice.Index).Value * _
                    (100 - rowDGV.Cells(Me.ColDiscount.Index).Value) / 100, _
                    DefaultPrecision))
                    rowDGV.Cells(Me.ColDiscAmount.Index).Value = Me.ObjCD.DigitGrouping(CDec(rowDGV.Cells(Me.ColUnitPrice.Index).Value) - _
                    CDec(rowDGV.Cells(Me.ColNettPrice.Index).Value), DefaultPrecision)
                Else
                    'hitung nett price
                    rowDGV.Cells(Me.ColNettPrice.Index).Value = Me.ObjCD.DigitGrouping(CDec(rowDGV.Cells(Me.ColUnitPrice.Index).Value) - CDec(rowDGV.Cells(Me.ColDiscAmount.Index).Value), _
                    DefaultPrecision)
                    'hitung disc %
                    rowDGV.Cells(Me.ColDiscount.Index).Value = Me.ObjCD.DigitGrouping((CDec(rowDGV.Cells(Me.ColDiscAmount.Index).Value) * 100) / CDec(rowDGV.Cells(Me.ColUnitPrice.Index).Value), 2)
                End If
                '\\TR=Quo
                If rowDGV.Cells(Me.ColPriceByMultiUnit.Index).Value = 1 Then
                    rowDGV.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(rowDGV.Cells(Me.ColQtyMulti.Index).Value * _
                    rowDGV.Cells(Me.ColNettPrice.Index).Value, DefaultPrecision))
                Else
                    rowDGV.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(rowDGV.Cells(Me.ColQty.Index).Value * _
                    rowDGV.Cells(Me.ColNettPrice.Index).Value, DefaultPrecision))
                End If
                'rowDGV.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(rowDGV.Cells(Me.ColQtyMulti.Index).Value * _
                'rowDGV.Cells(Me.ColNettPrice.Index).Value, Me.DefaultPrecision))
            Next
            Me.dgvItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            'hitung yang sudah ada didalam dgv
            Me.Hitung_Total()
            'hitung item yang ada di gbo item
            Me.Hitung_Amount()
            Me.txtSubTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Me.txtSubTotal.Text, DefaultPrecision))
            Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Me.txtTaxAmount.Text, DefaultPrecision))
            Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Me.TxtTotal.Text, DefaultPrecision))
            Me.txtNettPrice.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Me.txtNettPrice.Text, DefaultPrecision))
            Me.txtAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Me.txtAmount.Text, DefaultPrecision))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkTax_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkTax.CheckedChanged
        Dim Dummy As String = ""

        Try
            If Me.chkTax.Checked = True Then
                Me.txtTax.ReadOnly = False
                Me.txtTax.BackColor = Me.TxtPrimarySO.BackColor
                Me.txtTax.TabStop = True

                Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.TaxPercentage)
                If Dummy = "" Then
                    Dummy = "0"
                End If
                Me.txtTax.Text = Me.ObjCD.DigitGrouping(Dummy, 2)
            Else
                Me.txtTax.ReadOnly = True
                Me.txtTax.BackColor = Me.TxtReadOnly.BackColor
                Me.txtTax.TabStop = False
                Me.txtTax.Text = 0
                Me.txtTaxAmount.Text = 0
            End If

            If Me.txtTax.Text = "" Then
                Me.txtTax.Text = "0"
            End If

            If Me.txtRemarks.Enabled = True And IsNumeric(Me.txtTax.Text) = True And _
            IsNumeric(Me.txtSubTotal.Text) Then
                Me.Hitung_Total()
                If CDbl(Me.txtTax.Text) <= 100 Then
                    'hitung pajak
                    If Me.chkTax.Checked = True Then
                        If Me.txtPrimaryCurrency.Text = "" Then
                            Me.txtPrimaryCurrency.Text = "-1"
                        End If
                        If Me.txtPrimaryCurrency.Text = 1 Then
                            Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(Me.txtTax.Text * CDbl(Me.txtSubTotal.Text) / 100, _
                            0, ClsCekData.EnumTipeRounding.Down) _
                            , DefaultPrecision)
                        Else
                            Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.txtTax.Text * CDbl(Me.txtSubTotal.Text) / 100, _
                            DefaultPrecision, ClsCekData.EnumTipeRounding.Down)
                        End If
                    Else
                        Me.txtTax.Text = "0"
                        Me.txtTaxAmount.Text = "0"
                    End If
                    'hitung total = sub total + tax
                    Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDbl(Me.txtSubTotal.Text) _
                    + CDbl(Me.txtTaxAmount.Text), DefaultPrecision))
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCustomerCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCustomerCode.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.ObjCustomer.ApakahSudahAda(Me.txtCustomerCode.Text) = True Then
                    Me.txtPrimaryCustomer.Text = Me.ObjCustomer.GetPrimary(Me.txtCustomerCode.Text)
                    Me.txtCustomerName.Text = Me.ObjCustomer.GetCustName(Me.txtPrimaryCustomer.Text)
                    Me.ObjCustomer.IsiCboAddress(Me.cboCustomerAddress, Me.txtPrimaryCustomer.Text)
                    Me.ObjCustomer.IsiCboAddress(Me.CboAddressShipTo, Me.txtPrimaryCustomer.Text)
                    Me.txtContactPerson.Text = Me.ObjCustomer.GetCustContactPerson(Me.txtPrimaryCustomer.Text)
                    Me.cboCustomerAddress.SelectedIndex = 0
                    Me.CboAddressShipTo.SelectedIndex = 0
                    Me.TxtSalesPCode.Focus()
                Else
                    'ada 2 Customer dengan kode yang sama atau tidak ada Customer yang dimaksud
                    Me.btnFindCustomer.PerformClick()
                End If
            End If

            Me.UpdatePaymentTerm()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCustomerCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustomerCode.Leave
        Dim i As New FrmCariMasterCustomer
        i.BackColorAll = Me.BackColorAll

        Try
            If Me.txtCustomerCode.Text <> "" Then
                If Me.ObjCustomer.ApakahSudahAda(Me.txtCustomerCode.Text) = True Then
                    Me.txtPrimaryCustomer.Text = Me.ObjCustomer.GetPrimary(Me.txtCustomerCode.Text)
                    Me.txtCustomerName.Text = Me.ObjCustomer.GetCustName(Me.txtPrimaryCustomer.Text)
                    Me.ObjCustomer.IsiCboAddress(Me.cboCustomerAddress, Me.txtPrimaryCustomer.Text)
                    Me.ObjCustomer.IsiCboAddress(Me.CboAddressShipTo, Me.txtPrimaryCustomer.Text)
                    Me.txtContactPerson.Text = Me.ObjCustomer.GetCustContactPerson(Me.txtPrimaryCustomer.Text)
                    Me.cboCustomerAddress.SelectedIndex = 0
                    Me.CboAddressShipTo.SelectedIndex = 0
                Else
                    If Me.txtCustomerCode.Text <> "" Then
                        i.txtCari.Text = Me.txtCustomerCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCustomer > 0 Then
                        Me.txtPrimaryCustomer.Text = i.PrimaryCustomer
                        Me.txtCustomerCode.Text = Me.ObjCustomer.GetCustomerCode(Me.txtPrimaryCustomer.Text)
                        Me.txtCustomerName.Text = Me.ObjCustomer.GetCustName(Me.txtPrimaryCustomer.Text)
                        Me.ObjCustomer.IsiCboAddress(Me.cboCustomerAddress, Me.txtPrimaryCustomer.Text)
                        Me.ObjCustomer.IsiCboAddress(Me.CboAddressShipTo, Me.txtPrimaryCustomer.Text)
                        Me.cboCustomerAddress.SelectedIndex = 0
                        Me.CboAddressShipTo.SelectedIndex = 0
                        Me.txtContactPerson.Text = Me.ObjCustomer.GetCustContactPerson(Me.txtPrimaryCustomer.Text)
                    Else
                        Me.txtCustomerCode.Clear()
                        Me.txtCustomerName.Clear()
                        Me.cboCustomerAddress.Items.Clear()
                        Me.txtCustomerAddIdentifier.Text = ""
                        Me.txtPrimaryCustomer.Text = "-1"
                        Me.txtContactPerson.Clear()
                        Me.txtPhoneNo.Clear()
                    End If
                End If
            Else
                Me.txtCustomerCode.Clear()
                Me.txtCustomerName.Clear()
                Me.cboCustomerAddress.Items.Clear()
                Me.txtCustomerAddIdentifier.Text = ""
                Me.txtPrimaryCustomer.Text = "-1"
                Me.txtContactPerson.Clear()
                Me.txtPhoneNo.Clear()
            End If
            Me.UpdatePaymentTerm()
            Me.UpdateSalesPerson()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub UpdatePaymentTerm()
        Dim PT As Decimal = 0

        Try
            If Me.txtRemarks.Enabled = True Then
                If Me.txtPrimaryCustomer.Text = "-1" Then
                    Me.txtPaymentDueDate.Text = "0"
                Else
                    PT = Me.ObjCustomer.GetPaymentTerm(Me.txtPrimaryCustomer.Text)

                    If CDbl(Me.txtPaymentDueDate.Text) = 0 Then
                        If Me.baru = True Then
                            Me.txtPaymentDueDate.Text = PT
                        Else
                            If MsgBox("Update payment term?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                                Me.txtPaymentDueDate.Text = PT
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindCustomer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindCustomer.Click
        Dim i As New FrmCariMasterCustomer
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.txtCustomerCode.Text <> "" Then
                i.txtCari.Text = Me.txtCustomerCode.Text
            End If

            i.ShowDialog()
            If i.PrimaryCustomer > 0 Then
                Me.txtPrimaryCustomer.Text = i.PrimaryCustomer
                Me.txtCustomerCode.Text = Me.ObjCustomer.GetCustomerCode(Me.txtPrimaryCustomer.Text)
                Me.txtCustomerName.Text = Me.ObjCustomer.GetCustName(Me.txtPrimaryCustomer.Text)
                Me.ObjCustomer.IsiCboAddress(Me.cboCustomerAddress, Me.txtPrimaryCustomer.Text)
                Me.ObjCustomer.IsiCboAddress(Me.CboAddressShipTo, Me.txtPrimaryCustomer.Text)
                Me.cboCustomerAddress.SelectedIndex = 0
                Me.CboAddressShipTo.SelectedIndex = 0
                Me.txtContactPerson.Text = Me.ObjCustomer.GetCustContactPerson(Me.txtPrimaryCustomer.Text)
                Me.TxtSalesPCode.Focus()
            Else
                Me.txtCustomerCode.Clear()
                Me.txtCustomerName.Clear()
                Me.cboCustomerAddress.Items.Clear()
                Me.CboAddressShipTo.Items.Clear()
                Me.txtCustomerAddIdentifier.Text = ""
                Me.txtPrimaryCustomer.Text = "-1"
                Me.txtContactPerson.Clear()
                Me.txtPhoneNo.Clear()
                Me.txtCustomerCode.Focus()
            End If

            Me.UpdatePaymentTerm()
            Me.UpdateSalesPerson()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub UpdateSalesPerson()
        Try
            If Me.txtRemarks.Enabled = True Then
                If Me.txtPrimaryCustomer.Text <> "-1" Then
                    If Me.TxtPrimarySalesP.Text = "-1" Or Me.TxtPrimarySalesP.Text = "" Then
                        Me.TxtPrimarySalesP.Text = Me.ObjCustomer.GetPrimary_SalesPerson(Me.txtPrimaryCustomer.Text)
                        Me.TxtSalesPCode.Text = Me.ObjSalesP.GetCode(Me.TxtPrimarySalesP.Text)
                        Me.TxtSalesPName.Text = Me.ObjSalesP.GetName(Me.TxtPrimarySalesP.Text)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboCustomerAddress_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboCustomerAddress.SelectedIndexChanged
        Try
            If Me.txtRemarks.Enabled = True And IsNumeric(Me.txtPrimaryCustomer.Text) = True Then
                If Me.cboCustomerAddress.SelectedIndex <> -1 Then
                    Me.txtCustomerAddIdentifier.Text = Me.ObjCustomer.GetIdentifierAdd(Me.txtPrimaryCustomer.Text, _
                    Me.cboCustomerAddress.Text)
                    'isikan data contact person dan no telepon
                    Me.txtPhoneNo.Text = Me.ObjCustomer.GetPhoneNo(Me.txtPrimaryCustomer.Text, _
                    Me.txtCustomerAddIdentifier.Text)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub chkUseCustCodeNo_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkUseCustCodeNo.CheckedChanged
        Try
            If Me.txtRemarks.Enabled = True Then
                If Me.chkUseCustCodeNo.Checked = True Then
                    Me.txtCustomerProductCodeNo.ReadOnly = False
                    Me.txtCustomerProductCodeNo.BackColor = Me.txtPrimaryCategory.BackColor
                    Me.txtCustomerProductCodeNo.TabStop = True
                    Me.TxtCodeNo.ReadOnly = True
                    Me.TxtCodeNo.BackColor = Me.TxtSalesOrderNo.BackColor
                    Me.TxtCodeNo.TabStop = False
                    Me.BtnFindCodeNo.Enabled = False
                    Me.btnFindCustomerProductCode.Enabled = True
                    Me.ColSuppDesc.Visible = True
                    Me.ColSupplierCodeNo.Visible = True
                    Me.txtCustomerProductCodeNo.Focus()
                Else
                    Me.txtCustomerProductCodeNo.ReadOnly = True
                    Me.txtCustomerProductCodeNo.BackColor = Me.TxtSalesOrderNo.BackColor
                    Me.txtCustomerProductCodeNo.TabStop = False
                    Me.TxtCodeNo.ReadOnly = False
                    Me.TxtCodeNo.BackColor = Me.txtPrimaryCategory.BackColor
                    Me.TxtCodeNo.TabStop = True
                    Me.BtnFindCodeNo.Enabled = True
                    Me.btnFindCustomerProductCode.Enabled = False
                    Me.ColSuppDesc.Visible = False
                    Me.ColSupplierCodeNo.Visible = False
                    Me.TxtCodeNo.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtPaymentDueDate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPaymentDueDate.KeyDown
        Try
            'harus angka bulat
            If e.KeyCode = Keys.OemPeriod Or e.KeyCode = Keys.Decimal Then
                e.SuppressKeyPress = True
                Exit Sub
            End If
            Me.ObjCD.Number_KeyDown(Me.txtPaymentDueDate, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtPaymentDueDate_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPaymentDueDate.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.txtPaymentDueDate, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTax_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTax.KeyDown
        Try
            Me.buffer = Me.txtTax.Text
            If Me.txtTax.ReadOnly = False Then
                Me.ObjCD.Number_KeyDown(Me.txtTax, sender, e)
            End If
            If IsNumeric(Me.txtTax.Text) = True Then
                If CDbl(Me.txtTax.Text) > 100 Then
                    Me.txtTax.Text = Me.buffer
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTax_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTax.KeyUp
        Try
            If Me.txtTax.ReadOnly = False Then
                Me.ObjCD.Number_KeyUp(Me.txtTax, sender, e)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTax_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTax.Leave
        Try
            If Me.txtTax.Text = "" Then
                Me.txtTax.Text = "0"
            End If
            If Me.txtRemarks.Enabled = True And IsNumeric(Me.txtTax.Text) = True And _
            IsNumeric(Me.txtSubTotal.Text) Then
                Me.Hitung_Total()
                If CDbl(Me.txtTax.Text) <= 100 Then
                    'hitung pajak
                    If Me.chkTax.Checked = True Then
                        If Me.txtPrimaryCurrency.Text = 1 Then
                            Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.DigitGrouping(Me.txtTax.Text * CDbl(Me.txtSubTotal.Text) / 100, _
                            0, ClsCekData.EnumTipeRounding.Down) _
                            , DefaultPrecision)
                        Else
                            Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.txtTax.Text * CDbl(Me.txtSubTotal.Text) / 100, _
                            DefaultPrecision, ClsCekData.EnumTipeRounding.Down)
                        End If
                    Else
                        Me.txtTax.Text = "0"
                        Me.txtTaxAmount.Text = "0"
                    End If
                    'hitung total = sub total + tax
                    Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDbl(Me.txtSubTotal.Text) _
                    + CDbl(Me.txtTaxAmount.Text), DefaultPrecision))
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCustomerProductCodeNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustomerProductCodeNo.LostFocus
        If Me.txtCustomerProductCodeNo.ReadOnly = False And Me.txtCustomerProductCodeNo.Enabled = True Then
            Me.txtCustomerProductCodeNo.BackColor = Me.txtPrimaryCustomer.BackColor
        End If
    End Sub

    Private Sub TxtQtyperBatch_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchDefault.TextChanged
        Try
            If IsNumeric(Me.TxtQtyperBatchDefault.Text) = True And IsNumeric(Me.TxtDiscount.Text) = _
            True And IsNumeric(Me.TxtUnitPrice.Text) = True Then
                Me.Hitung_Amount()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUnitPrice_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtUnitPrice.KeyDown
        'Me.buffer = Me.TxtUnitPrice.Text
        Me.ObjCD.Number_KeyDown(Me.TxtUnitPrice, sender, e)
        Me.TxtUnitPrice.MaxLength = 14
        'If Me.TxtUnitPrice.Text.Length > Me.TxtUnitPrice.MaxLength Then
        '    Me.TxtUnitPrice.Text = Me.buffer
        'End If
    End Sub

    Private Sub TxtUnitPrice_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtUnitPrice.KeyUp
        Me.ObjCD.Number_KeyUp(Me.TxtUnitPrice, sender, e)
    End Sub

    'Private Sub TxtUnitPrice_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtUnitPrice.TextChanged
    '    Try
    '        If IsNumeric(Me.TxtQtyperBatch.Text) = True And IsNumeric(Me.TxtDiscount.Text) = _
    '        True And IsNumeric(Me.TxtUnitPrice.Text) = True Then
    '            Me.Hitung_Amount()
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Private Sub TxtDiscount_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDiscount.KeyDown
        Try
            Me.buffer = Me.TxtDiscount.Text
            Me.ObjCD.Number_KeyDown(Me.TxtDiscount, sender, e)
            If IsNumeric(Me.TxtDiscount.Text) = True Then
                If CDbl(Me.TxtDiscount.Text) > 100 Then
                    Me.TxtDiscount.Text = Me.buffer
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscount_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDiscount.KeyUp
        Me.ObjCD.Number_KeyUp(Me.TxtDiscount, sender, e)
    End Sub

    'Private Sub TxtDiscount_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtDiscount.TextChanged
    '    Try
    '        If IsNumeric(Me.TxtQtyperBatch.Text) = True And IsNumeric(Me.TxtDiscount.Text) = _
    '        True And IsNumeric(Me.TxtUnitPrice.Text) = True Then
    '            If CDbl(Me.TxtDiscount.Text) <= 100 Then
    '                Me.Hitung_Amount()
    '            End If
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Private Sub TxtCustomerName_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustomerName.TextChanged
        Try
            If Me.txtPrimaryCustomer.Text <> "-1" And Me.txtPrimaryCustomer.Text <> "" Then
                If Me.txtCustomerName.Text <> "" Then
                    Me.Text = "Sales Order: " + Me.txtCustomerCode.Text + " - " + Me.txtCustomerName.Text
                Else
                    Me.Text = "Sales Order"
                End If
            Else
                Me.Text = "Sales Order"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbAddress_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbAddress.LinkClicked
        Try
            If Me.txtRemarks.Enabled = True OrElse (IsNumeric(Me.TxtPrimarySO.Text) AndAlso CDec(Me.TxtPrimarySO.Text) > 0) Then
                If Me.txtPrimaryCustomer.Text = "" Or Me.txtPrimaryCustomer.Text = "-1" Then
                    MsgBox("Please specify customer first.", MsgBoxStyle.Information, "Information")
                    Me.txtCustomerCode.Focus()
                    Exit Sub
                End If
                Dim ij As New Core_Module.FrmDetailAddress
                ij.vAddIdentifier = Me.txtCustomerAddIdentifier.Text
                ij.PrimaryCust = Me.txtPrimaryCustomer.Text
                ij.btnOK.BringToFront()
                ij.AcceptButton = ij.btnOK
                ij.btnExit.TabStop = False
                If Me.txtRemarks.Enabled = False Then
                    ij.cboCustAddCode.Enabled = False
                End If
                ij.ShowDialog()
                If ij.DialogResult = Windows.Forms.DialogResult.OK Then
                    Me.ObjCustomer.IsiCboAddress(Me.cboCustomerAddress, Me.txtPrimaryCustomer.Text)
                    Me.cboCustomerAddress.Text = Me.ObjCustomer.GetKodeAddress(ij.vAddIdentifierResult)
                End If
                ij.Dispose()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboOrderBy_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboOrderBy.SelectedIndexChanged
        Try
            If Me.txtRemarks.Enabled = True Then
                If Me.cboOrderBy.SelectedIndex <> 0 Then
                    If Me.baru = True Then
                        Me.txtCustomerPONo.Text = Me.objSO.GetNOCustPOBaru(Me.dtpDate.Value)
                    Else
                        Me.txtCustomerPONo.Text = Me.objSO.GetCustPONo(Me.objSO.Primary)
                    End If

                    Me.txtCustomerPONo.ReadOnly = True
                    Me.txtCustomerPONo.BackColor = Me.TxtSalesOrderNo.BackColor
                    Me.txtCustomerPONo.TabStop = False
                Else
                    Me.txtCustomerPONo.ReadOnly = False
                    Me.txtCustomerPONo.BackColor = Me.TxtPrimaryCodeNo.BackColor
                    Me.txtCustomerPONo.TabStop = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCustomerPONo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustomerPONo.GotFocus
        Try
            If Me.txtCustomerPONo.Enabled = True And Me.txtCustomerPONo.ReadOnly = False Then
                Me.txtCustomerPONo.BackColor = Me.BackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''' <summary>
    ''' cek jika ada no customer po number yang sama, maka berikan peringatan
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub txtCustomerPONo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustomerPONo.Leave
        Dim Dummy As String = ""

        Try
            If Me.txtCustomerPONo.Text <> "" Then
                Dummy = Me.objSO.GetNoSO_ByCustomerPONO(Me.txtCustomerPONo.Text)
                If Me.baru = True Then
                    If Dummy <> "" Then
                        MsgBox("The sales order " + Dummy + " also refer to customer po: " + Me.txtCustomerPONo.Text + _
                        ". This may indicate invalid purchase order.", MsgBoxStyle.Information, "Information")
                    End If
                Else
                    If Me.objSO.NoTransaksi <> Dummy And Dummy <> "" Then
                        MsgBox("The sales order " + Dummy + " also refer to customer po: " + Me.txtCustomerPONo.Text + _
                        ". This may indicate invalid purchase order.", MsgBoxStyle.Information, "Information")
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCustomerPONo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustomerPONo.LostFocus
        Try
            If Me.txtCustomerPONo.Enabled = True And Me.txtCustomerPONo.ReadOnly = False Then
                Me.txtCustomerPONo.BackColor = Me.TxtPrimaryCodeNo.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ChkPriceCondition_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkPriceCondition.CheckedChanged
        Dim row As DataGridViewRow
        Dim PriceCondition As String = ""

        Try
            If Me.ChkPriceCondition.Checked = True Then
                If Me.txtRemarks.Enabled = True Then
                    Me.CboPriceCondition.Enabled = True
                    Me.cboPriceConditionItem.Enabled = False
                    Me.cboPriceConditionItem.TabStop = False
                    'ubah semuanya menjadi yang diinginkan
                    If Me.CboPriceCondition.SelectedIndex = 0 Then
                        PriceCondition = "Exclude"
                    Else
                        PriceCondition = "Include"
                    End If
                    Me.cboPriceConditionItem.SelectedIndex = Me.CboPriceCondition.SelectedIndex
                    If Me.txtRemarks.Enabled = True Then
                        For Each row In Me.dgvItem.Rows
                            row.Cells(Me.ColPriceCond.Index).Value = PriceCondition
                        Next
                        Me.Hitung_Total()
                    End If
                End If
            Else
                Me.CboPriceCondition.Enabled = False
                If Me.txtRemarks.Enabled = True Then
                    Me.cboPriceConditionItem.Enabled = True
                End If
                Me.cboPriceConditionItem.TabStop = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboPriceCondition_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboPriceCondition.SelectedIndexChanged
        Try
            If Me.ChkPriceCondition.Checked = True Then
                Dim row As DataGridViewRow
                Dim PriceCondition As String = ""
                If Me.CboPriceCondition.SelectedIndex = 0 Then
                    PriceCondition = "Exclude"
                Else
                    PriceCondition = "Include"
                End If
                For Each row In Me.dgvItem.Rows
                    row.Cells(Me.ColPriceCond.Index).Value = PriceCondition
                Next
                Me.cboPriceConditionItem.SelectedIndex = Me.CboPriceCondition.SelectedIndex

                If Me.txtRemarks.Enabled = True Then
                    Me.Hitung_Total()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''' <summary>
    ''' mengembalikan jumlah barisnya yang nilai diskonnya, sama dengan yang diberikan.
    ''' </summary>
    ''' <param name="vDiskon"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetJmlByDiskon(ByVal vDiskon As Double) As Integer
        Try
            Dim row As DataGridViewRow
            Dim retval As Integer = 0

            For Each row In Me.dgvItem.Rows
                If CDbl(row.Cells(Me.ColDiscount.Index).Value) = vDiskon Then
                    retval += 1
                End If
            Next

            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub ChkDiscountConditionSama_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkDiscountConditionSama.CheckedChanged
        Try
            If Me.txtRemarks.Enabled = True Then
                If Me.ChkDiscountConditionSama.Checked = True Then
                    Me.txtReferenceDiscount.ReadOnly = False
                    Me.txtReferenceDiscount.BackColor = Me.TxtPrimaryCodeNo.BackColor
                    Me.TxtDiscount.ReadOnly = True
                    Me.TxtDiscount.TabStop = False
                    Me.TxtDiscount.BackColor = Me.TxtSalesOrderNo.BackColor
                    Me.TxtDiscAmount.ReadOnly = True
                    Me.TxtDiscAmount.TabStop = False
                    Me.TxtDiscAmount.BackColor = Me.TxtReadOnly.BackColor

                    If CDbl(Me.txtReferenceDiscount.Text) <> CDbl(Me.TxtDiscount.Text) Or _
                    Me.GetJmlByDiskon(Me.txtReferenceDiscount.Text) <> Me.dgvItem.Rows.Count Then
                        Dim row As DataGridViewRow
                        For Each row In Me.dgvItem.Rows
                            row.Cells(Me.ColDiscount.Index).Value = Me.txtReferenceDiscount.Text
                            Me.Hitung_Amount(row)
                        Next
                        Me.Hitung_Total()
                    End If
                    Me.TxtDiscount.Text = Me.txtReferenceDiscount.Text
                    Me.Hitung_Amount()
                Else
                    Me.txtReferenceDiscount.Text = Me.ObjCD.DigitGrouping("0", 2)
                    Me.txtReferenceDiscount.ReadOnly = True
                    Me.txtReferenceDiscount.BackColor = Me.TxtSalesOrderNo.BackColor
                    Me.TxtDiscount.BackColor = Me.TxtPrimaryCodeNo.BackColor
                    Me.TxtDiscount.ReadOnly = False
                    Me.TxtDiscount.TabStop = True
                    Me.TxtDiscAmount.ReadOnly = False
                    Me.TxtDiscAmount.TabStop = True
                    Me.TxtDiscAmount.BackColor = Me.TxtPrimaryCodeNo.BackColor
                    Me.Hitung_Amount()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtReferenceDiscount_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtReferenceDiscount.Enter
        Try
            If Me.txtRemarks.Enabled = True Then
                If Me.txtReferenceDiscount.Enabled = True Then
                    If Me.txtReferenceDiscount.ReadOnly = False Then
                        Me.buffer = Me.txtReferenceDiscount.Text
                        Me.ObjCD.Number_Enter(Me.txtReferenceDiscount, sender, e)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtReferenceDiscount_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtReferenceDiscount.GotFocus
        Try
            If Me.txtRemarks.Enabled = True Then
                If Me.txtReferenceDiscount.Enabled = True Then
                    If Me.txtReferenceDiscount.ReadOnly = True Then
                        Me.txtReferenceDiscount.BackColor = Me.TxtReadOnly.BackColor
                    Else
                        Me.txtReferenceDiscount.BackColor = Me.BackColorAll
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtReferenceDiscount_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtReferenceDiscount.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.txtReferenceDiscount, sender, e)
            If IsNumeric(Me.txtReferenceDiscount.Text) = True Then
                If CDbl(Me.txtReferenceDiscount.Text) > 100 Then
                    Me.txtReferenceDiscount.Text = Me.buffer
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtReferenceDiscount_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtReferenceDiscount.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.txtReferenceDiscount, sender, e)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItem_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvItem.UserDeletingRow
        Try
            'ketika user akan melakukan hapus item ada 2 hal yang perlu dicek:
            '1. jika sudah ada do yang merujuk pada item yang akan dihapus, maka item tidak bisa dihapus
            Dim ObjDOItem As New ClsTransaksiDOItem
            If ObjDOItem.GetJmlBySO(Me.objSO.Primary, _
                Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.Row.Index).Value) > 0 Then
                Dim ObjDO As New ClsTransaksiDO
                Dim DONo_s As String = ObjDO.GetDONO(Me.objSO.Primary, _
                Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.Row.Index).Value)
                If DONo_s.LastIndexOf(",") <> -1 Then
                    MsgBox("This item referred by several delivery orders, they are: " + DONo_s + ".")
                Else
                    MsgBox("This item referred by one delivery orders: " + DONo_s + ".")
                End If
                ObjDO = Nothing
                e.Cancel = True
                Exit Sub
            End If
            ObjDOItem = Nothing

            '2. jika ada booking inventory yang berkaitan dengan item yang akan dihapus maka diberikan
            '   peringatan
            Dim ObjItemBookingGIT As New ClsTransaksiBookingGITitem
            If ObjItemBookingGIT.GetQtyByPrimarySO_n_Product_Multi(Me.objSO.Primary, _
                Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.Row.Index).Value) > 0 Then
                Dim BookNo As String = ObjItemBookingGIT.GetBookingNoByPrimarySO_n_Product(Me.objSO.Primary, _
                Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.Row.Index).Value)
                If BookNo.LastIndexOf(",") <> -1 Then
                    If MsgBox("This item referred by several GIT booking, they are: " + BookNo + _
                    ". Deleting them will also delete the booking data, will you continue?", _
                    MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                        e.Cancel = True
                        Exit Sub
                    End If
                Else
                    If MsgBox("This item referred by GIT booking: " + BookNo + _
                    ". Deleting them will also delete the booking data, will you continue?", _
                    MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                        e.Cancel = True
                        Exit Sub
                    End If
                End If
            End If
            ObjItemBookingGIT = Nothing

            '3. jika ada booking GIT yang berkaitan dengan item yang akan dihapus maka akan diberikan 
            '   peringatan
            Dim ObjItemBookingINV As New ClsTransaksiBookingInvItem
            If ObjItemBookingINV.GetJmlByPrimarySO_n_PrimaryProduct(Me.objSO.Primary, _
                Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.Row.Index).Value) > 0 Then
                Dim BookNo As String = ObjItemBookingINV.GetBookingNoByPrimarySO(Me.objSO.Primary, _
                Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.Row.Index).Value)
                If BookNo.LastIndexOf(",") <> -1 Then
                    If MsgBox("This item referred by several inventory booking, they are: " + BookNo + _
                    ". Deleting them will also delete the booking data, will you continue?", _
                    MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                        e.Cancel = True
                        Exit Sub
                    End If
                Else
                    If MsgBox("This item referred by inventory booking: " + BookNo + _
                    ". Deleting them will also delete the booking data, will you continue?", _
                    MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                        e.Cancel = True
                        Exit Sub
                    End If
                End If
            End If
            ObjItemBookingINV = Nothing

            '4. hapuskan item delivery
            Me.vTBLItemDelivery.DefaultView.RowFilter = "identifier = '" + Me.dgvItem.Item(Me.ColIdentifier.Index, e.Row.Index).Value.ToString + "'"
            While Me.vTBLItemDelivery.DefaultView.Count > 0
                Me.vrowItemDelivery = Me.vTBLItemDelivery.DefaultView(0).Row
                Me.vTBLItemDelivery.Rows.Remove(Me.vrowItemDelivery)
            End While
            Me.vTBLItemDelivery.DefaultView.RowFilter = ""

            If e.Row.Index = Me.dgvItem.RowCount - 1 Then
                Me.EndEditDeliveryDateCont2(True)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbAddQuotation_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbAddQuotation.LinkClicked
        Dim rowDGV As DataGridViewRow
        Dim rowDelivery As DseTransaksiSO.tbltrans_so_itemdeliveryRow
        Dim IdentifierBaru As String = ""
        Dim index As Integer = -1
        Dim index2 As Integer = -1
        Dim hasil As MsgBoxResult
        Dim listOfDelDate As New List(Of String)
        Dim count As Integer = 0
        Dim Dv, Dv2, Dv3 As DataRowView
        Dim replace_ As Boolean = False
        Try
            'input data dari quotation
            Dim ij As New FrmTransaksiSO_ImporQuo
            Dim ik As New FrmCariTransaksiQuotation
            ik.BackColorAll = Me.BackColorAll

            ik.AuthorizedOnly = True
            ik.UserCode = Me.UserCode
            ik.ShowDialog()

            If ik.PrimaryTransaksi > 0 Then
                Dim ObjQuo As New ClsTransaksiQuotation
                ObjQuo.CariDetail(ik.PrimaryTransaksi, Me.dgvItemQuo, Me.vTBLItemDelivery2)

                If ij.ShowDialog = Windows.Forms.DialogResult.OK Then
                    'copy data header
                    If ij.rboHeaderOnly.Checked = True Or ij.rboAll.Checked = True Then
                        With ObjQuo
                            Me.txtPrimaryCustomer.Text = .PrimaryCustomer
                            Me.txtCustomerCode.Text = Me.ObjCustomer.GetCustomerCode(Me.txtPrimaryCustomer.Text)
                            Me.txtCustomerName.Text = Me.ObjCustomer.GetCustName(Me.txtPrimaryCustomer.Text)
                            Me.ObjCustomer.IsiCboAddress(Me.cboCustomerAddress, Me.txtPrimaryCustomer.Text)
                            Me.cboCustomerAddress.SelectedIndex = 0
                            While Me.txtCustomerAddIdentifier.Text <> .CustIdentifierAdd And Me.cboCustomerAddress.SelectedIndex + 1 <= Me.cboCustomerAddress.Items.Count
                                Me.cboCustomerAddress.SelectedIndex += 1
                            End While
                            If Me.RboAddressShipTo.Checked = True Then
                                Me.ObjCustomer.IsiCboAddress(Me.CboAddressShipTo, Me.txtPrimaryCustomer.Text)
                                Me.CboAddressShipTo.SelectedIndex = 0
                            Else
                                Me.RboAddressShipTo.Checked = True
                            End If
                            Me.CboAddressShipTo.SelectedIndex = Me.cboCustomerAddress.SelectedIndex

                            Me.txtContactPerson.Text = .CustContactPerson
                            Me.txtPhoneNo.Text = .CustPhone
                            Me.txtPrimaryCurrency.Text = .PrimaryCurrency
                            Me.txtCurrencyCode.Text = ObjCurrency.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                            Me.SetDefaultPrecision1()
                            Me.chkTax.Checked = .UseTax
                            Me.txtTax.Text = .TaxPercentage
                            Me.txtPaymentTerm.Text = .PaymentTerm
                            Me.txtPaymentDueDate.Text = .PaymentDue
                            Me.cboPaymentDueDateReference.SelectedIndex = .RefPaymentDate
                            Me.ChkDiscountConditionSama.Checked = False
                            Me.ChkPriceCondition.Checked = True
                            Me.CboPriceCondition.SelectedIndex = 0
                            '\\TR=Quo
                            If .NeedDP = 1 Then
                                Me.chkUseDP.Checked = True
                            Else
                                Me.chkUseDP.Checked = False
                            End If
                            Me.TxtDPFromPercent.Text = .DPFromPercent
                            Me.TxtDPPercent.Text = Me.ObjCD.DigitGrouping(.DPPercent, 2)
                            Me.TxtDPAmount.Text = Me.ObjCD.DigitGrouping(.DPAmount, Me.DefaultPrecision)
                            If .PrimaryCustomerShipTo = -1 Then
                                Me.RboAddressShipTo.Checked = True
                                Me.TxtIdentifierAddressShipTo.Text = .IdentifierAddressShipTo
                                Me.ObjCustomer.IsiCboAddress(Me.CboAddressShipTo, Me.txtPrimaryCustomer.Text)
                                Me.CboAddressShipTo.Text = Me.ObjCustomer.GetKodeAddress(Me.TxtIdentifierAddressShipTo.Text)
                            Else
                                Me.RboCustomerShipTo.Checked = True
                                Me.TxtPrimaryCustomerShipTo.Text = .PrimaryCustomerShipTo
                                Me.TxtCustomerCodeShipTo.Text = Me.ObjCustomer.GetCustomerCode(Me.TxtPrimaryCustomerShipTo.Text)
                                Me.TxtCustomerDescShipTo.Text = Me.ObjCustomer.GetCustName(Me.TxtPrimaryCustomerShipTo.Text)
                                Me.TxtIdentifierAddressCustomerShipTo.Text = .IdentifierAddressCustomerShipTo
                                Me.ObjCustomer.IsiCboAddress(Me.CboAddressCustomerShipTo, Me.TxtPrimaryCustomerShipTo.Text)
                                Me.CboAddressCustomerShipTo.Text = Me.ObjCustomer.GetKodeAddress(Me.TxtIdentifierAddressCustomerShipTo.Text)
                            End If
                        End With
                        '\\TR=Quo
                        Me.UpdateSalesPerson()
                    End If

                    'data item
                    If ij.rboItemOnly.Checked = True Or ij.rboAll.Checked = True Then
                        'hapus data item jika perlu
                        If ij.rboReplaceData.Checked = True Then
                            Me.dgvItem.Rows.Clear()
                            Me.vTBLItemDelivery.Rows.Clear()
                        Else '\\TR=Quo
                            'tambahan jika di Append Data, Takutnya ada product yg udah ada 
                            Me.CekProduct()
                        End If
                        'input data item
                        Dim row As DataGridViewRow
                        For Each row In Me.dgvItemQuo.Rows
                            If Me.GetIndex(row.Cells(Me.ColQuo_PrimaryCodeNo.Index).Value) = -1 Then
                                'tambahkan baris
                                replace_ = False
                                IdentifierBaru = Me.GetIdentifierBaru
                                Me.dgvItem.Rows.Add()
replaceFromQuotation:
                                rowDGV = Me.dgvItem.Rows(Me.dgvItem.Rows.Count - 1)
                                rowDGV.Cells(Me.ColNo.Index).Value = Me.dgvItem.RowCount
                                If replace_ = True Then
                                    index2 = Me.GetIndex(row.Cells(Me.ColQuo_PrimaryCodeNo.Index).Value)
                                    rowDGV = Me.dgvItem.Rows(index2)
                                End If
                                With row
                                    rowDGV.Cells(Me.ColCodeNo.Index).Value = row.Cells(Me.ColQuo_CodeNo.Index).Value
                                    rowDGV.Cells(Me.ColDescription.Index).Value = row.Cells(Me.ColDescription.Index).Value
                                    rowDGV.Cells(Me.ColSuppDesc.Index).Value = ""
                                    rowDGV.Cells(Me.ColSupplierCodeNo.Index).Value = ""
                                    rowDGV.Cells(Me.ColQty.Index).Value = row.Cells(Me.ColQuo_Qty.Index).Value
                                    rowDGV.Cells(Me.ColUnit.Index).Value = row.Cells(Me.ColQuo_Unit.Index).Value
                                    rowDGV.Cells(Me.ColUnitPrice.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColQuo_UnitPrice.Index).Value, Me.DefaultPrecision)
                                    rowDGV.Cells(Me.ColDiscount.Index).Value = row.Cells(Me.ColQuo_Discount.Index).Value
                                    rowDGV.Cells(Me.ColDiscAmount.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColQuo_DiscAmount.Index).Value, Me.DefaultPrecision)
                                    rowDGV.Cells(Me.ColDeliveryHistory.Index).Value = ""
                                    rowDGV.Cells(Me.ColNettPrice.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColQuo_NettPrice.Index).Value, Me.DefaultPrecision)
                                    rowDGV.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColQuo_Amount.Index).Value, Me.DefaultPrecision)
                                    rowDGV.Cells(Me.ColPriceCond.Index).Value = "Exclude"
                                    rowDGV.Cells(Me.ColDeliveryCont.Index).Value = row.Cells(Me.ColQuo_DeliveryCont.Index).Value
                                    rowDGV.Cells(Me.ColDeliveryDate.Index).Value = row.Cells(Me.ColQuo_DeliveryDate.Index).Value
                                    rowDGV.Cells(Me.ColRemark.Index).Value = row.Cells(Me.ColQuo_Remark.Index).Value
                                    rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value = row.Cells(Me.ColQuo_PrimaryCodeNo.Index).Value
                                    rowDGV.Cells(Me.ColPrimaryUnits.Index).Value = row.Cells(Me.ColQuo_PrimaryUnits.Index).Value
                                    rowDGV.Cells(Me.ColUseSecondary.Index).Value = row.Cells(Me.ColQuo_UseSecondary.Index).Value
                                    rowDGV.Cells(Me.ColDiscFromPerct.Index).Value = row.Cells(Me.ColQuo_DiscFromPerct.Index).Value
                                    If replace_ = False Then
                                        rowDGV.Cells(Me.ColIdentifier.Index).Value = IdentifierBaru
                                    End If

                                    rowDGV.Cells(Me.ColQtyMulti.Index).Value = row.Cells(Me.ColQuo_QtyMulti.Index).Value
                                    rowDGV.Cells(Me.ColUnitsMulti.Index).Value = row.Cells(Me.ColQuo_UnitsMulti.Index).Value
                                    rowDGV.Cells(Me.ColPrimaryUnitsMulti.Index).Value = row.Cells(Me.ColQuo_PrimaryUnitsMulti.Index).Value
                                    rowDGV.Cells(Me.ColRatioReference.Index).Value = row.Cells(Me.ColQuo_RatioReference.Index).Value
                                    rowDGV.Cells(Me.ColRatioDefault.Index).Value = row.Cells(Me.ColQuo_RatioDefault.Index).Value
                                    rowDGV.Cells(Me.ColRatioDesc.Index).Value = row.Cells(Me.ColQuo_RatioDesc.Index).Value
                                    rowDGV.Cells(Me.ColToleranceDown.Index).Value = row.Cells(Me.ColQuo_ToleranceDown.Index).Value
                                    rowDGV.Cells(Me.ColToleranceDownPerc.Index).Value = row.Cells(Me.ColQuo_ToleranceDownPerc.Index).Value
                                    rowDGV.Cells(Me.ColToleranceDownFromPerc.Index).Value = row.Cells(Me.ColQuo_ToleranceDownFromPerc.Index).Value
                                    rowDGV.Cells(Me.ColToleranceUp.Index).Value = row.Cells(Me.ColQuo_ToleranceUp.Index).Value
                                    rowDGV.Cells(Me.ColToleranceUpPerc.Index).Value = row.Cells(Me.ColQuo_ToleranceUpPerc.Index).Value
                                    rowDGV.Cells(Me.ColToleranceUpFromPerc.Index).Value = row.Cells(Me.ColQuo_ToleranceUpFromPerc.Index).Value
                                    '\\TR=Quo
                                    rowDGV.Cells(Me.ColPriceByMultiUnit.Index).Value = row.Cells(Me.ColQuoPriceByMultiUnit.Index).Value
                                    If row.Cells(Me.ColQuoPriceByMultiUnit.Index).Value = 1 Then
                                        rowDGV.Cells(Me.ColPriceByMultiUnit_Str.Index).Value = "/" + row.Cells(Me.ColUnitsMulti.Index).Value
                                    Else
                                        rowDGV.Cells(Me.ColPriceByMultiUnit_Str.Index).Value = "/" + row.Cells(Me.ColUnit.Index).Value
                                    End If
                                    'tambahkan item delivery bila ada janji kirim
                                    If rowDGV.Cells(Me.ColDeliveryCont.Index).Value = Boolean.TrueString Then
                                        If replace_ = True Then
                                            Me.vTBLItemDelivery.DefaultView.RowFilter = "identifier = '" + rowDGV.Cells(Me.ColIdentifier.Index).Value.ToString + "'"
                                            While Me.vTBLItemDelivery.DefaultView.Count > 0
                                                Me.vrowItemDelivery = Me.vTBLItemDelivery.DefaultView(0).Row
                                                Me.vTBLItemDelivery.Rows.Remove(Me.vrowItemDelivery)
                                            End While
                                            Me.vTBLItemDelivery.DefaultView.RowFilter = Nothing
                                        End If
                                        If rowDGV.Cells(Me.ColDeliveryDate.Index).Value = "Multiple" Then
                                            Me.vTBLItemDelivery2.DefaultView.RowFilter = "identifier = '" + row.Cells(Me.ColQuo_Identifier.Index).Value.ToString + "'"
                                            For Each Dv In vTBLItemDelivery2.DefaultView
                                                vrowItemDelivery2 = Dv.Row
                                                rowDelivery = Me.vTBLItemDelivery.Newtbltrans_so_itemdeliveryRow
                                                With rowDelivery
                                                    .delivery_date = vrowItemDelivery2.delivery_date
                                                    .delivery_qty = vrowItemDelivery2.delivery_qty
                                                    .delivery_priority = Me.ObjDeliveryPriority.GetPrimaryDefault
                                                    .identifier = rowDGV.Cells(Me.ColIdentifier.Index).Value
                                                    .primary = Me.GetPrimaryMax(Me.vTBLItemDelivery) + 1
                                                    .primary_transaksi = Me.TxtPrimarySO.Text
                                                    .remarks = ""
                                                    .delivery_qty_multi = vrowItemDelivery2.delivery_qty_multi
                                                End With
                                                Me.vTBLItemDelivery.Rows.Add(rowDelivery)
                                            Next
                                        Else
                                            rowDelivery = Me.vTBLItemDelivery.Newtbltrans_so_itemdeliveryRow
                                            With rowDelivery
                                                .delivery_date = CDate(rowDGV.Cells(Me.ColDeliveryDate.Index).Value)
                                                .delivery_qty = CDec(rowDGV.Cells(Me.ColQty.Index).Value)
                                                .delivery_priority = Me.ObjDeliveryPriority.GetPrimaryDefault
                                                .identifier = rowDGV.Cells(Me.ColIdentifier.Index).Value
                                                .primary = Me.GetPrimaryMax(Me.vTBLItemDelivery) + 1
                                                .primary_transaksi = Me.TxtPrimarySO.Text
                                                .remarks = ""
                                                .delivery_qty_multi = CDec(rowDGV.Cells(Me.ColQtyMulti.Index).Value)
                                            End With
                                            Me.vTBLItemDelivery.Rows.Add(rowDelivery)
                                        End If
                                    End If
                                End With
                            Else
                                index = Me.GetIndex_Ratio(row.Cells(Me.ColQuo_PrimaryCodeNo.Index).Value, row.Cells(Me.ColQuo_UnitPrice.Index).Value, _
                                        row.Cells(Me.ColQuo_NettPrice.Index).Value, row.Cells(Me.ColQuo_Discount.Index).Value, row.Cells(Me.ColQuo_RatioDefault.Index).Value, row.Cells(Me.ColQuo_RatioReference.Index).Value)
                                If index <> -1 Then
                                    rowDGV = Me.dgvItem.Rows(index)
                                    rowDGV.Cells(Me.ColQty.Index).Value = CDec(rowDGV.Cells(Me.ColQty.Index).Value) + CDec(row.Cells(Me.ColQuo_Qty.Index).Value)
                                    rowDGV.Cells(Me.ColQtyMulti.Index).Value = CDec(rowDGV.Cells(Me.ColQtyMulti.Index).Value) + CDec(row.Cells(Me.ColQuo_QtyMulti.Index).Value)
                                    Me.Hitung_Amount(rowDGV)
                                    'If rowDGV.Cells(Me.ColDeliveryCont.Index).Value = Boolean.TrueString Then
                                    Me.vTBLItemDelivery.DefaultView.RowFilter = "identifier = '" + rowDGV.Cells(Me.ColIdentifier.Index).Value.ToString + "'"
                                    Me.vTBLItemDelivery2.DefaultView.RowFilter = "identifier = '" + row.Cells(Me.ColQuo_Identifier.Index).Value.ToString + "'"
                                    For Each Dv In vTBLItemDelivery.DefaultView
                                        vrowItemDelivery = Dv.Row
                                        listOfDelDate.Add(vrowItemDelivery.delivery_date)
                                    Next

                                    For Each Dv2 In vTBLItemDelivery2.DefaultView
                                        vrowItemDelivery2 = Dv2.Row
                                        If listOfDelDate.Contains(vrowItemDelivery2.delivery_date) = True Then
                                            Dim a As String = ""
                                            Me.vTBLItemDelivery.DefaultView.RowFilter = Nothing
                                            Me.vTBLItemDelivery.DefaultView.RowFilter = "identifier = '" & rowDGV.Cells(Me.ColIdentifier.Index).Value & "' AND [delivery date] = #" & vrowItemDelivery2.delivery_date & "#"
                                            For Each Dv3 In vTBLItemDelivery.DefaultView
                                                vrowItemDelivery = Dv3.Row
                                                If vrowItemDelivery.delivery_date = vrowItemDelivery2.delivery_date Then
                                                    vrowItemDelivery.delivery_qty = CDec(vrowItemDelivery.delivery_qty) + CDec(vrowItemDelivery2.delivery_qty)
                                                    vrowItemDelivery.delivery_qty_multi = CDec(vrowItemDelivery.delivery_qty_multi) + CDec(vrowItemDelivery2.delivery_qty_multi)
                                                End If
                                            Next
                                        Else
                                            If vrowItemDelivery2.delivery_qty > 0 Then
                                                rowDelivery = Me.vTBLItemDelivery.Newtbltrans_so_itemdeliveryRow
                                                With rowDelivery
                                                    .delivery_date = vrowItemDelivery2.delivery_date
                                                    .delivery_qty = vrowItemDelivery2.delivery_qty
                                                    .delivery_priority = Me.ObjDeliveryPriority.GetPrimaryDefault
                                                    .identifier = rowDGV.Cells(Me.ColIdentifier.Index).Value
                                                    .primary = Me.GetPrimaryMax(Me.vTBLItemDelivery) + 1
                                                    .primary_transaksi = Me.TxtPrimarySO.Text
                                                    .delivery_qty_multi = vrowItemDelivery2.delivery_qty_multi
                                                    .remarks = ""
                                                End With
                                                Me.vTBLItemDelivery.Rows.Add(rowDelivery)
                                            End If
                                        End If
                                    Next
                                    listOfDelDate.Clear()
                                    'End If
                                Else
                                    count += 1
                                    If count = 1 Then
                                        hasil = MsgBox("Because of the unit price and discount between product quotation with SO different" & vbCrLf & "Do you want to replace data from quotation?", MsgBoxStyle.YesNo, "Confirm")
                                    End If
                                    If hasil = MsgBoxResult.Yes Then
                                        replace_ = True
                                        GoTo replaceFromQuotation
                                    End If
                                End If
                            End If
                        Next
                        'cek kembali bila ada product yang item delivery date nya lebih dari 1 maka ubah col delivery date nya menjadi 'Multiple'
                        For Each rowDGV In Me.dgvItem.Rows
                            Me.vTBLItemDelivery.DefaultView.RowFilter = Nothing
                            Me.vTBLItemDelivery.DefaultView.RowFilter = "identifier = '" + rowDGV.Cells(Me.ColIdentifier.Index).Value.ToString + "'"
                            If Me.vTBLItemDelivery.DefaultView.Count > 1 Then
                                rowDGV.Cells(Me.ColDeliveryDate.Index).Value = "Multiple"
                            End If
                        Next
                        Me.ObjCD.SetNumber(Me.dgvItem)
                        Me.Hitung_Total()
                    End If
                End If
            End If
            ik.Dispose()
            ij.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim FrmO As New Core_Module.FrmMasterUser_Otorisasi

        Try
            If Me.TxtSalesOrderNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtSalesOrderNo.Text <> "" Then
                If Me.ObjUser.ApakahMemilikiHakAkses(Me.UserCode, "FrmTransaksiSO_Print") = False Then
                    FrmO.AccessName = "FrmTransaksiSO_Print"
                    If FrmO.ShowDialog <> Windows.Forms.DialogResult.OK Then
                        MsgBox("You don't have permission to perform this action", MsgBoxStyle.Information, "Information")
                        Exit Sub
                    End If
                End If

                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                Dim obj As New ClsTransaksiSalesOrder
                Dim FrmCetak As New frmCetakSO

                If obj.CetakSO(FrmCetak.crv1, Me.TxtPrimarySO.Text) = True Then
                    FrmCetak.ShowDialog()

                End If
                obj = Nothing
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FrmO.Dispose()
        End Try
    End Sub

    Private Sub dtpDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpDate.ValueChanged
        Try
            Me.BackgroundWorker1.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTax_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTax.LostFocus
        Try
            If Me.txtRemarks.Enabled = True Then
                If Me.txtTax.ReadOnly = False Then
                    Me.txtTax.BackColor = Me.txtPrimaryCategory.BackColor
                Else
                    Me.txtTax.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTax_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTax.GotFocus
        Try
            If Me.txtRemarks.Enabled = True Then
                If Me.txtTax.ReadOnly = False Then
                    Me.txtTax.BackColor = Me.BackColorAll
                Else
                    Me.txtTax.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Private Sub TxtQtyperBatch_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchDefault.Leave
    '    Dim DIskon, UnitPrice As Decimal

    '    Try
    '        If Me.txtRemarks.Enabled = True Then
    '            If Me.TxtPrimaryCodeNo.Text <> "" And Me.TxtPrimaryCodeNo.Text <> "-1" Then
    '                'Me.TxtQtyperBatch.Text = Me.ObjCD.DigitGrouping(Me.TxtQtyperBatch.Text, Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text))
    '                Me.ObjCD.Number_Leave(Me.TxtQtyperBatchDefault, sender, e, Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text), Me.ObjCD)

    '                If IsNumeric(Me.TxtQtyperBatchDefault.Text) = True And IsNumeric(Me.TxtDiscount.Text) = _
    '                True And IsNumeric(Me.TxtUnitPrice.Text) = True Then
    '                    Me.ObjProduct.GetUnitPriceNDiskon(Me.txtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtQtyperBatchDefault.Text, _
    '                                Me.txtPrimaryCurrency.Text, Me.txtPrimaryUnitsDefault.Text, UnitPrice, DIskon)

    '                    Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(UnitPrice, _
    '                    Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text))
    '                    If Me.ChkDiscountConditionSama.Checked = True Then
    '                        If IsNumeric(Me.txtReferenceDiscount.Text) = False Then
    '                            Me.txtReferenceDiscount.Text = Me.ObjCD.DigitGrouping(0, 2)
    '                        End If
    '                        Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Me.txtReferenceDiscount.Text, 2)
    '                    Else
    '                        Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(DIskon, 2)
    '                    End If
    '                    Me.Hitung_Amount()
    '                End If
    '            End If
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Private Sub TxtDiscount_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscount.Leave
        Try
            'Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Me.TxtDiscount.Text, 2)
            Me.ObjCD.Number_Leave(Me.TxtDiscount, sender, e, 2, Me.ObjCD)
            If CDec(Me.TxtUnitPrice.Text) = 0 Then
                Me.TxtDiscount.Text = 0
                Me.ObjCD.Number_Leave(Me.TxtDiscount, sender, e, 2, Me.ObjCD)
            End If
            If CDec(Me.TxtDiscount.Text) <> CDec(Me.buffer) Then
                Me.TxtDiscFromPerct.Text = "1"
                Me.Hitung_Amount()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtUnitPrice_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitPrice.Leave
        Try
            'Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(Me.TxtUnitPrice.Text, Me.DefaultPrecision)
            Me.ObjCD.Number_Leave(Me.TxtUnitPrice, sender, e, Me.DefaultPrecision, Me.ObjCD)
            'unit price harus selalu lebih besar dari disc amount
            'hanya berlaku kalo ngitung disc berdasarkan disc amount, klo berdasarkan perc, langsung hitung ulang disc amountnya
            If Me.TxtDiscFromPerct.Text = "0" Then
                If CDec(Me.TxtDiscAmount.Text) > CDec(Me.TxtUnitPrice.Text) Then
                    Me.TxtUnitPrice.Text = Me.buffer
                    Exit Sub
                End If
            End If
            Me.Hitung_Amount()
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
            If Me.txtRemarks.Enabled = True Then
                If Me.ObjNumbering.Apakah_Autonumber(ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.SO) = True Then
                    If Me.baru = False Then
                        If Me.objSO.TanggalTransaksi.Month = Me.dtpDate.Value.Month And Me.objSO.TanggalTransaksi.Year = Me.dtpDate.Value.Year Then
                            'jika kembali ke bulan yang bersangkutan maka no kembali ke awal
                            Me.TxtSalesOrderNo.Text = Me.objSO.NoTransaksi
                        Else
                            'jika pindah bulan maka akan ganti no
                            Me.TxtSalesOrderNo.Text = Me.objSO.GetNoBaru1(Me.dtpDate.Value)
                        End If
                    Else
                        Me.TxtSalesOrderNo.Text = Me.objSO.GetNoBaru1(Me.dtpDate.Value)
                    End If
                End If

                Me.dtpdeliverydate.Value = Me.dtpDate.Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LblUnitPrice_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblUnitPrice.LinkClicked
        Try
            If Me.TxtPrimaryCodeNo.Text <> "" And Me.TxtPrimaryCodeNo.Text <> "-1" Then
                Dim ijk As New FrmTransaksiSO_HistorycalPrice
                ijk.BackColorAll = Me.BackColorAll

                ijk.mPrimaryProduct = Me.TxtPrimaryCodeNo.Text
                ijk.txtPrimaryCustomer.Text = Me.txtPrimaryCustomer.Text
                ijk.txtCustomerCode.Text = Me.txtCustomerCode.Text
                ijk.TxtCustomerName.Text = Me.txtCustomerName.Text
                ijk.ShowDialog()

                ijk = Nothing
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtExchangeRate_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtExchangeRate.Enter
        Try
            If Me.TxtExchangeRate.Enabled = True And Me.TxtExchangeRate.ReadOnly = False Then
                Me.ObjCD.Number_Enter(Me.TxtExchangeRate, sender, e)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtExchangeRate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtExchangeRate.GotFocus
        Try
            Dim txt As TextBox = TryCast(sender, TextBox)
            If IsNothing(txt) = False Then
                If txt.Enabled = True Then
                    If txt.ReadOnly = True Then
                        txt.BackColor = Me.TxtReadOnly.BackColor
                    Else
                        txt.BackColor = Me.BackColorAll
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtExchangeRate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtExchangeRate.KeyDown
        Try
            If Me.TxtExchangeRate.Enabled = True And Me.TxtExchangeRate.ReadOnly = False Then
                Me.ObjCD.Number_KeyDown(Me.TxtExchangeRate, sender, e, 0)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtExchangeRate_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtExchangeRate.KeyUp
        Try
            If Me.TxtExchangeRate.Enabled = True And Me.TxtExchangeRate.ReadOnly = False Then
                Me.ObjCD.Number_KeyUp(Me.TxtExchangeRate, sender, e, False)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtExchangeRate_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtExchangeRate.Leave
        Try
            If Me.TxtExchangeRate.Enabled = True And Me.TxtExchangeRate.ReadOnly = False Then
                If CDec(Me.TxtExchangeRate.Text) = 0 Then
                    Me.TxtExchangeRate.Text = 1
                End If
                Me.ObjCD.Number_Leave(Me.TxtExchangeRate, sender, e, Core_Module.DefPrecExchangeRate, Me.ObjCD)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtExchangeRate_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtExchangeRate.LostFocus
        Try
            Dim txt As TextBox = TryCast(sender, TextBox)
            If IsNothing(txt) = False Then
                If txt.Enabled = True Then
                    If txt.ReadOnly = True Then
                        txt.BackColor = Me.TxtReadOnly.BackColor
                    Else
                        txt.BackColor = Me.TxtPrimarySalesP.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSalesPCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSalesPCode.GotFocus
        Try
            If Me.TxtSalesPCode.Enabled = True Then
                If Me.TxtSalesPCode.ReadOnly = False Then
                    Me.TxtSalesPCode.BackColor = Me.BackColorAll
                Else
                    Me.TxtSalesPCode.BackColor = Me.TxtSalesPName.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSalesPCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSalesPCode.LostFocus
        Try
            If Me.TxtSalesPCode.Enabled = True Then
                If Me.TxtSalesPCode.ReadOnly = False Then
                    Me.TxtSalesPCode.BackColor = Me.TxtPrimarySalesP.BackColor
                Else
                    Me.TxtSalesPCode.BackColor = Me.TxtSalesPName.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFindSalesP_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFindSalesP.Click
        Dim FrmCariGroup As New FrmMasterSalesPerson_Cari

        Try
            FrmCariGroup.BackColorAll = Me.BackColorAll
            If Me.TxtPrimarySalesP.Text = -1 And Me.TxtSalesPCode.Text = "" Then
                FrmCariGroup.ShowDialog()
            Else
                FrmCariGroup.txtCari.Text = Me.TxtSalesPCode.Text
                FrmCariGroup.ShowDialog()
            End If
            If CDbl(FrmCariGroup.PrimarySalesPerson) > 0 Then
                Me.TxtPrimarySalesP.Text = FrmCariGroup.PrimarySalesPerson
                Me.TxtSalesPCode.Text = Me.ObjSalesP.GetCode(FrmCariGroup.PrimarySalesPerson)
                Me.TxtSalesPName.Text = Me.ObjSalesP.GetName(FrmCariGroup.PrimarySalesPerson)
                If Me.txtCustomerPONo.TabStop = True Then
                    Me.txtCustomerPONo.Focus()
                Else
                    Me.cboCustomerAddress.Focus()
                End If
            Else
                Me.TxtPrimarySalesP.Text = -1
                Me.TxtSalesPCode.Text = ""
                Me.TxtSalesPName.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            FrmCariGroup.Dispose()
        End Try
    End Sub

    Private Sub TxtSalesPCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtSalesPCode.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.ObjSalesP.ApakahSudahAda(Me.TxtSalesPCode.Text) = True Then
                    Me.TxtPrimarySalesP.Text = Me.ObjSalesP.GetPrimary(Me.TxtSalesPCode.Text)
                    Me.TxtSalesPName.Text = Me.ObjSalesP.GetName(Me.TxtPrimarySalesP.Text)
                    If Me.txtCustomerPONo.TabStop = True Then
                        Me.txtCustomerPONo.Focus()
                    Else
                        Me.cboCustomerAddress.Focus()
                    End If
                Else
                    Me.btnFindSalesP.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtSalesPCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSalesPCode.Leave
        Try
            If Me.TxtSalesPCode.Text <> "" Then
                If Me.ObjSalesP.ApakahSudahAda(Me.TxtSalesPCode.Text) = True Then
                    Me.TxtPrimarySalesP.Text = Me.ObjSalesP.GetPrimary(Me.TxtSalesPCode.Text)
                    Me.TxtSalesPName.Text = Me.ObjSalesP.GetName(Me.TxtPrimarySalesP.Text)
                Else
                    Me.btnFindSalesP.PerformClick()
                End If
            Else
                Me.TxtPrimarySalesP.Text = -1
                Me.TxtSalesPCode.Text = ""
                Me.TxtSalesPName.Text = ""
            End If
            Me.TxtSalesPCode.BackColor = Me.TxtPrimarySalesP.BackColor
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        Try
            If Me.txtRemarks.Enabled = True Then
                Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimarySO.BackColor, False)
                Me.ObjCD.SetBackColor(Me.TabPage1.Controls, Me.TxtPrimarySO.BackColor, False)
                Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.TxtPrimarySO.BackColor, False)
                Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.TxtPrimarySO.BackColor, False)
                Me.ObjCD.SetBackColor(Me.TabPage2.Controls, Me.TxtPrimarySO.BackColor, False)
                Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel1.Controls, Me.TxtPrimarySO.BackColor, False)
                Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel2.Controls, Me.TxtPrimarySO.BackColor, False)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscAmount_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscAmount.Enter
        Try
            Me.buffer = Me.TxtDiscAmount.Text
            Me.ObjCD.Number_Enter(Me.TxtDiscAmount, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscAmount_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscAmount.GotFocus
        Try
            If Me.txtRemarks.Enabled = True Then
                If Me.TxtDiscAmount.ReadOnly = False Then
                    Me.TxtDiscAmount.BackColor = Me.BackColorAll
                Else
                    Me.TxtDiscAmount.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscAmount_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDiscAmount.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtDiscAmount, sender, e, CDec(Me.TxtUnitPrice.Text))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscAmount_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDiscAmount.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtDiscAmount, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscAmount_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscAmount.Leave
        Try
            Me.ObjCD.Number_Leave(Me.TxtDiscAmount, sender, e, Me.DefaultPrecision, Me.ObjCD)
            'disc amount tidak boleh lebih besar dari unit price
            If CDec(Me.TxtDiscAmount.Text) > CDec(Me.TxtUnitPrice.Text) Then
                Me.TxtDiscAmount.Text = Me.buffer
                Exit Sub
            End If
            If CDec(Me.TxtDiscAmount.Text) <> CDec(Me.buffer) Then
                Me.TxtDiscFromPerct.Text = "0"
                Me.Hitung_Amount()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscAmount_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscAmount.LostFocus
        Try
            If Me.txtRemarks.Enabled = True Then
                If Me.TxtDiscAmount.ReadOnly = False Then
                    Me.TxtDiscAmount.BackColor = Me.TxtPrimarySO.BackColor
                Else
                    Me.TxtDiscAmount.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtReferenceDiscount_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtReferenceDiscount.Leave
        Try
            If Me.txtRemarks.Enabled = True Then
                If Me.txtReferenceDiscount.Enabled = True Then
                    If Me.txtReferenceDiscount.ReadOnly = False Then
                        Me.ObjCD.Number_Leave(Me.txtReferenceDiscount, sender, e, 2, Me.ObjCD)
                        If Me.buffer <> Me.txtReferenceDiscount.Text Then
                            If CDbl(Me.txtReferenceDiscount.Text) <> CDbl(Me.TxtDiscount.Text) Or _
                                               Me.GetJmlByDiskon(Me.txtReferenceDiscount.Text) <> Me.dgvItem.Rows.Count Then
                                Dim row As DataGridViewRow
                                For Each row In Me.dgvItem.Rows
                                    row.Cells(Me.ColDiscount.Index).Value = Me.txtReferenceDiscount.Text
                                    row.Cells(Me.ColDiscFromPerct.Index).Value = 1
                                    Me.Hitung_Amount(row)
                                Next
                                Me.Hitung_Total()
                            End If

                            Me.TxtDiscount.Text = Me.txtReferenceDiscount.Text
                            Me.Hitung_Amount()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtReferenceDiscount_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtReferenceDiscount.LostFocus
        Try
            If Me.txtRemarks.Enabled = True Then
                If Me.txtReferenceDiscount.Enabled = True Then
                    If Me.txtReferenceDiscount.ReadOnly = True Then
                        Me.txtReferenceDiscount.BackColor = Me.TxtReadOnly.BackColor
                    Else
                        Me.txtReferenceDiscount.BackColor = Me.TxtPrimaryCodeNo.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub dgvItemQuo_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs)
        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatchMulti_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchMulti.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtQtyperBatchMulti, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatchMulti_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchMulti.GotFocus
        TxtQtyperBatchMulti.BackColor = BackColorAll
    End Sub

    Private Sub TxtQtyperBatchMulti_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyperBatchMulti.KeyDown
        Me.ObjCD.Number_KeyDown(Me.TxtQtyperBatchMulti, sender, e)
    End Sub

    Private Sub TxtQtyperBatchMulti_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyperBatchMulti.KeyUp
        Me.ObjCD.Number_KeyUp(Me.TxtQtyperBatchMulti, sender, e)
    End Sub

    Private Sub TxtQtyperBatchMulti_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchMulti.Leave
        Dim DIskon, UnitPrice As Decimal

        Try
            If Me.txtRemarks.Enabled = True Then
                If Me.TxtPrimaryCodeNo.Text <> "" And Me.TxtPrimaryCodeNo.Text <> "-1" Then
                    'Me.TxtQtyperBatch.Text = Me.ObjCD.DigitGrouping(Me.TxtQtyperBatch.Text, Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text))
                    Me.ObjCD.Number_Leave(Me.TxtQtyperBatchMulti, sender, e, Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text), Me.ObjCD)

                    If IsNumeric(Me.TxtQtyperBatchMulti.Text) = True And IsNumeric(Me.TxtDiscount.Text) = _
                    True And IsNumeric(Me.TxtUnitPrice.Text) = True Then
                        Me.SetQtyDefault()

                        If Me.txtPriceByMultiUnit.Text = "1" Then
                            Me.ObjProduct.GetUnitPriceNDiskon(Me.txtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtQtyperBatchMulti.Text, _
                            Me.txtPrimaryCurrency.Text, Me.txtPrimaryUnitsMulti.Text, UnitPrice, DIskon)
                        Else
                            Me.ObjProduct.GetUnitPriceNDiskon(Me.txtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtQtyperBatchDefault.Text, _
                            Me.txtPrimaryCurrency.Text, Me.txtPrimaryUnitsDefault.Text, UnitPrice, DIskon)
                        End If
                        If UnitPrice = 0 And DIskon = 0 Then
                            If Me.txtPriceByMultiUnit.Text = "1" Then
                                Me.objSO.GetHistoricalPrice(Me.txtPrimaryUnitsMulti.Text, Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryCurrency.Text, _
                                UnitPrice, DIskon, Me.txtPriceByMultiUnit.Text, Me.txtPrimaryCustomer.Text)
                            Else
                                Me.objSO.GetHistoricalPrice(Me.txtPrimaryUnitsDefault.Text, Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryCurrency.Text, _
                                UnitPrice, DIskon, Me.txtPriceByMultiUnit.Text, Me.txtPrimaryCustomer.Text)
                            End If
                        End If
                        Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(UnitPrice, _
                        Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text))

                        If Me.ChkDiscountConditionSama.Checked = True Then
                            If IsNumeric(Me.txtReferenceDiscount.Text) = False Then
                                Me.txtReferenceDiscount.Text = Me.ObjCD.DigitGrouping(0, 2)
                            End If
                            Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Me.txtReferenceDiscount.Text, 2)
                        Else
                            Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(DIskon, 2)
                        End If
                        Me.Hitung_Amount()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatchMulti_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchMulti.LostFocus
        TxtQtyperBatchMulti.BackColor = Me.txtPrimaryCategory.BackColor
    End Sub

    Private Sub TxtPrimaryCodeNo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtPrimaryCodeNo.TextChanged
        Try
            Me.PrepareMultiUnit()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub PrepareMultiUnit()
        Dim Dta As New DseTransaksiSOTableAdapters.tblmaster_product_item_multiunitTableAdapter
        Try
            Me.CboUnit.DataSource = Nothing
            If Me.TxtPrimaryCodeNo.Text <> "-1" And Me.TxtPrimaryCodeNo.Text <> "" Then
                Dta.Connection.ConnectionString = Core_Module.GetConstring
                Dta.FillBy_PrimaryProduct(Me.mTblMultiUnit, Me.TxtPrimaryCodeNo.Text)
                Me.IsiCboUnit(Me.TxtPrimaryCodeNo.Text, Me.CboUnit, Me.tblCboUnit, Me.mTblMultiUnit)
                Me.CboUnit.SelectedIndex = -1
                Me.CboUnit.SelectedIndex = 0
            Else
                txtPrimaryUnitsDefault.Text = "-1"
                txtPrimaryUnitsMulti.Text = "-1"
                lblUnitsDefault.Text = "Units"
                lblUnitsMulti.Text = "Units"
                TxtQtyperBatchDefault.Text = "0"
                TxtQtyperBatchMulti.Text = "0"
                LblUnitPrice.Text = "Unit Price"
                CboUnit.DataSource = Nothing
                CboUnit.Refresh()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function IsiCboUnit(ByVal vPrimaryProduct As Double, ByVal vCbo As ComboBox, _
            ByVal vTbl As DseTransaksiSO.CboUnitDataTable, _
            ByVal vTblMultiUnit As DseTransaksiSO.tblmaster_product_item_multiunitDataTable) As Boolean
        Dim rowTbl As DseTransaksiSO.CboUnitRow
        Dim rowMultiUnit As DseTransaksiSO.tblmaster_product_item_multiunitRow
        Dim objProduct As New Core_Module.clsMasterProduct
        Dim objUnit As New Core_Module.ClsMasterUnit
        Try
            vCbo.DataSource = Nothing
            vCbo.Items.Clear()
            vTbl.Rows.Clear()
            vTblMultiUnit.DefaultView.RowFilter = "[Primary Product] = " & vPrimaryProduct.ToString
            rowTbl = vTbl.NewCboUnitRow
            rowTbl.Primary = objProduct.GetPrimaryUnit(vPrimaryProduct)
            rowTbl.Code = objUnit.GetUnitCode(rowTbl.Primary)
            vTbl.AddCboUnitRow(rowTbl)
            If vTblMultiUnit.DefaultView.Count > 0 Then
                For i As Integer = 0 To vTblMultiUnit.DefaultView.Count - 1
                    rowMultiUnit = vTblMultiUnit.DefaultView(i).Row
                    rowTbl = vTbl.NewCboUnitRow
                    rowTbl.Primary = rowMultiUnit.Primary_Unit
                    rowTbl.Code = objUnit.GetUnitCode(rowMultiUnit.Primary_Unit)
                    vTbl.AddCboUnitRow(rowTbl)
                Next
            End If
            vCbo.DataSource = vTbl.DefaultView
            vTbl.DefaultView.Sort = "[Code] ASC"
            vCbo.DisplayMember = "Code"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Private Sub CboUnit_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboUnit.Leave
        Dim UnitPrice, Diskon As Decimal

        Try
            UnitPrice = 0
            Diskon = 0

            If Me.CboUnit.Text <> Me.cboPriceByMultiUnit.Text Then
                Me.txtPriceByMultiUnit.Text = 0
            Else
                Me.txtPriceByMultiUnit.Text = 1
            End If
            If Me.txtPriceByMultiUnit.Text = "1" Then
                Me.ObjProduct.GetUnitPriceNDiskon(Me.txtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtQtyperBatchMulti.Text, _
                Me.txtPrimaryCurrency.Text, Me.txtPrimaryUnitsMulti.Text, UnitPrice, Diskon)
            Else
                Me.ObjProduct.GetUnitPriceNDiskon(Me.txtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtQtyperBatchDefault.Text, _
                Me.txtPrimaryCurrency.Text, Me.txtPrimaryUnitsDefault.Text, UnitPrice, Diskon)
            End If
            If UnitPrice = 0 And Diskon = 0 Then
                If Me.txtPriceByMultiUnit.Text = "1" Then
                    Me.objSO.GetHistoricalPrice(Me.txtPrimaryUnitsMulti.Text, Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryCurrency.Text, _
                    UnitPrice, Diskon, Me.txtPriceByMultiUnit.Text, Me.txtPrimaryCustomer.Text)
                Else
                    Me.objSO.GetHistoricalPrice(Me.txtPrimaryUnitsDefault.Text, Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryCurrency.Text, _
                    UnitPrice, Diskon, Me.txtPriceByMultiUnit.Text, Me.txtPrimaryCustomer.Text)
                End If
            End If
            Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(UnitPrice, _
            Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text))

            If Me.ChkDiscountConditionSama.Checked = True Then
                If IsNumeric(Me.txtReferenceDiscount.Text) = False Then
                    Me.txtReferenceDiscount.Text = Me.ObjCD.DigitGrouping(0, 2)
                End If
                Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Me.txtReferenceDiscount.Text, 2)
            Else
                Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Diskon, 2)
            End If

            Me.Hitung_Amount()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboUnit_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboUnit.SelectedIndexChanged
        'Dim UnitPrice, Discount As Decimal
        Dim row As DseTransaksiSO.CboUnitRow
        Try
            If Me.CboUnit.SelectedIndex <> -1 Then
                row = Me.tblCboUnit.DefaultView(Me.CboUnit.SelectedIndex).Row
                Me.txtPrimaryUnitsMulti.Text = row.Primary
                Me.lblUnitsMulti.Text = Me.CboUnit.Text
                Me.LblUnitPrice.Text = "Unit Price/" & Me.CboUnit.Text
                Me.TxtQtyperBatchMulti.Text = ObjCD.DigitGrouping(Me.TxtQtyperBatchMulti.Text, Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text))
                'Me.ObjProduct.GetUnitPriceNDiskon(Me.txtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtQtyperBatchMulti.Text, _
                'Me.txtPrimaryCurrency.Text, Me.txtPrimaryUnitsMulti.Text, UnitPrice, Discount)
                'Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(UnitPrice, Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text))
                'Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Discount, 2)
                Me.SetQtyDefault()
                Me.SetCboPriceByMultiUnit()
                'If IsNumeric(Me.TxtUnitPrice.Text) = True Then
                '    If CDec(Me.TxtUnitPrice.Text) > 0 Then
                '        Me.Hitung_Amount()
                '    End If
                'End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetQtyDefault()
        Dim RatioDefault As Decimal = 0
        Dim RatioReference As Decimal = 0
        Try
            RatioDefault = Me.objSO.GetRatioDefault(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text)
            RatioReference = Me.objSO.GetRatioReference(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text)
            Me.TxtQtyperBatchDefault.Text = Me.ObjCD.DigitGrouping(CDec(Me.TxtQtyperBatchMulti.Text) * RatioDefault / RatioReference, Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetQtyDefault(ByVal vQtyMulti As Double, ByVal vRatioDefault As Double, ByVal vRatioReference As Double, ByVal vDefPrec As Integer) As Double
        Try
            Return CDbl(ObjCD.DigitGrouping(vQtyMulti * vRatioDefault / vRatioReference, vDefPrec))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub SetQtyDefaultOnDgv(ByVal vRow As DataGridViewRow, ByVal vDefPrecDefault As Integer)
        Try
            vRow.Cells(Me.ColQty.Index).Value = ObjCD.DigitGrouping(CDec(vRow.Cells(Me.ColQtyMulti.Index).Value) * CDec(vRow.Cells(Me.ColRatioDefault.Index).Value) / CDec(vRow.Cells(Me.ColRatioReference.Index).Value), vDefPrecDefault)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SetRatioOnDgv(ByVal vDgvRow As DataGridViewRow)
        Try
            With vDgvRow
                If .Cells(Me.ColPrimaryUnitsMulti.Index).Value <> .Cells(Me.ColPrimaryUnits.Index).Value Then
                    .Cells(Me.ColRatioDesc.Index).Value = .Cells(Me.ColRatioReference.Index).Value & " " & .Cells(Me.ColUnitsMulti.Index).Value & " = " & _
                    .Cells(Me.ColRatioDefault.Index).Value & " " & .Cells(Me.ColUnit.Index).Value
                Else
                    .Cells(Me.ColRatioDesc.Index).Value = "-"
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FormatDGV()
        Dim row As DataGridViewRow
        Dim IsMultiUnit As Boolean = False
        Dim Dummy As String = ""
        Dim PriceByUnitDefault As Boolean = False

        Try
            For Each row In Me.dgvItem.Rows
                If row.Cells(Me.ColPrimaryUnitsMulti.Index).Value <> row.Cells(Me.ColPrimaryUnits.Index).Value Then
                    IsMultiUnit = True
                    Exit For
                End If
            Next
            For Each rowdgv As DataGridViewRow In Me.dgvItem.Rows
                If rowdgv.Cells(Me.ColPriceByMultiUnit.Index).Value = 0 Then
                    PriceByUnitDefault = True
                    Exit For
                End If
            Next
            Me.ColPriceByMultiUnit_Str.Visible = PriceByUnitDefault

            If IsMultiUnit = False Then
                Me.ColQty.Visible = False
                Me.ColUnit.Visible = False
                Me.ColRatioDesc.Visible = False
            Else
                Me.ColQty.Visible = True
                Me.ColUnit.Visible = True
                Me.ColRatioDesc.Visible = True

                ''walaupun multiunit=true tapi jika visible=false user defined jgn ditampilkan
                Dummy = Me.ObjSettingUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.SalesOrder, _
                            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.SalesOrder_ShowUnitsDefault, Me.UserCode)
                If Dummy = "" Then
                    Me.ColQty.Visible = False
                    Me.ColUnit.Visible = False
                    Me.ColRatioDesc.Visible = False
                Else
                    If Dummy <> 1 Then
                        Me.ColQty.Visible = False
                        Me.ColUnit.Visible = False
                        Me.ColRatioDesc.Visible = False
                    End If
                End If

            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditToleranceDown(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal row As DataGridViewRow)
        '\\TR#
        Dim DefPrecMulti As Integer = ObjProduct.GetAngkaDecimal(row.Cells(Me.ColPrimaryCodeNo.Index).Value, row.Cells(Me.ColPrimaryUnitsMulti.Index).Value)
        Try
            With row

                If Me.ObjCD.CekNonNegatif(.Cells(e.ColumnIndex).Value) = False Then
                    MsgBox("Please enter positive number for tolerance.", MsgBoxStyle.Information, "Information")
                    .Cells(e.ColumnIndex).Value = Me.buffer
                    Exit Sub
                End If
                .Cells(e.ColumnIndex).Value = Me.ObjCD.DigitGrouping(.Cells(e.ColumnIndex).Value, DefPrecMulti)
                If .Cells(e.ColumnIndex).Value > (.Cells(Me.ColQtyMulti.Index).Value / 2) Then
                    MsgBox("Tolerance cannot greater than 50% of qty.", MsgBoxStyle.Information, "Information")
                    .Cells(e.ColumnIndex).Value = Me.buffer
                    Exit Sub
                End If

                .Cells(Me.ColToleranceDownFromPerc.Index).Value = "0"
                If CDec(.Cells(Me.ColQtyMulti.Index).Value) > 0 Then
                    .Cells(Me.ColToleranceDownPerc.Index).Value = ObjCD.DigitGrouping(.Cells(Me.ColToleranceDown.Index).Value / .Cells(Me.ColQtyMulti.Index).Value * 100, 2)
                Else
                    .Cells(Me.ColToleranceDownPerc.Index).Value = ObjCD.DigitGrouping(0, 2)
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditToleranceDownPerc(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal row As DataGridViewRow)
        '\\TR#
        Dim DefPrecMulti As Integer = ObjProduct.GetAngkaDecimal(row.Cells(ColPrimaryCodeNo.Index).Value, row.Cells(ColPrimaryUnitsMulti.Index).Value)
        Try
            With row
                If Me.ObjCD.CekNonNegatif(.Cells(e.ColumnIndex).Value) = False Then
                    MsgBox("Please enter positive number for tolerance data.", MsgBoxStyle.Information, "Information")
                    .Cells(e.ColumnIndex).Value = Me.buffer
                    Exit Sub
                End If
                If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value > 50 Then
                    MsgBox("Tolerance cannot greater than 50%.", MsgBoxStyle.Information, "Information")
                    .Cells(e.ColumnIndex).Value = Me.buffer
                    Exit Sub
                End If
                If CekBatasanToleransi(e.RowIndex) = False Then
                    'MsgBox("The tolerance value (down) from the following product code cannot be >= from value of Qty.", MsgBoxStyle.Information, "Information")
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Me.dgvItem.Item(Me.ColToleranceDown.Index, e.RowIndex).Value = Me.bufferToleransiDown
                    Exit Sub
                End If
                .Cells(e.ColumnIndex).Value = ObjCD.DigitGrouping(.Cells(e.ColumnIndex).Value, 2)
                .Cells(Me.ColToleranceDownFromPerc.Index).Value = "1"
                If .Cells(Me.ColToleranceDownPerc.Index).Value = 50 Then
                    .Cells(Me.ColToleranceDown.Index).Value = ObjCD.DigitGrouping(.Cells(Me.ColToleranceDownPerc.Index).Value / 100 * .Cells(Me.ColQtyMulti.Index).Value, DefPrecMulti, ClsCekData.EnumTipeRounding.Down)
                Else
                    .Cells(Me.ColToleranceDown.Index).Value = ObjCD.DigitGrouping(.Cells(Me.ColToleranceDownPerc.Index).Value / 100 * .Cells(Me.ColQtyMulti.Index).Value, DefPrecMulti)
                End If

            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditToleranceUp(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal row As DataGridViewRow)
        Dim DefPrecMulti As Integer = ObjProduct.GetAngkaDecimal(row.Cells(Me.ColPrimaryCodeNo.Index).Value, row.Cells(Me.ColPrimaryUnitsMulti.Index).Value)
        Try
            With row
                If Me.ObjCD.CekNonNegatif(.Cells(e.ColumnIndex).Value) = False Then
                    MsgBox("Please enter positive number for tolerance.", MsgBoxStyle.Information, "Information")
                    .Cells(e.ColumnIndex).Value = Me.buffer
                    Exit Sub
                End If
                .Cells(e.ColumnIndex).Value = Me.ObjCD.DigitGrouping(.Cells(e.ColumnIndex).Value, DefPrecMulti)
                If .Cells(e.ColumnIndex).Value > CDec(.Cells(Me.ColQtyMulti.Index).Value) Then
                    MsgBox("Tolerance cannot greater than 100% of qty.", MsgBoxStyle.Information, "Information")
                    .Cells(e.ColumnIndex).Value = Me.buffer
                    Exit Sub
                End If

                .Cells(Me.ColToleranceUpFromPerc.Index).Value = "0"
                If CDec(.Cells(Me.ColQtyMulti.Index).Value) > 0 Then
                    .Cells(Me.ColToleranceUpPerc.Index).Value = ObjCD.DigitGrouping(.Cells(Me.ColToleranceUp.Index).Value / .Cells(Me.ColQtyMulti.Index).Value * 100, 2)
                Else
                    .Cells(Me.ColToleranceUpPerc.Index).Value = ObjCD.DigitGrouping(0, 2)
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditToleranceUpPerc(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal row As DataGridViewRow)
        '\\TR#
        Dim DefPrecMulti As Integer = ObjProduct.GetAngkaDecimal(row.Cells(ColPrimaryCodeNo.Index).Value, row.Cells(ColPrimaryUnitsMulti.Index).Value)
        Try
            With row
                If Me.ObjCD.CekNonNegatif(.Cells(e.ColumnIndex).Value) = False Then
                    MsgBox("Please enter positive number for tolerance data.", MsgBoxStyle.Information, "Information")
                    .Cells(e.ColumnIndex).Value = Me.buffer
                    Exit Sub
                End If
                If Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value > 100 Then
                    MsgBox("Tolerance cannot greater than 100%.", MsgBoxStyle.Information, "Information")
                    .Cells(e.ColumnIndex).Value = Me.buffer
                    Exit Sub
                End If
                .Cells(e.ColumnIndex).Value = ObjCD.DigitGrouping(.Cells(e.ColumnIndex).Value, 2)
                .Cells(Me.ColToleranceUpFromPerc.Index).Value = "1"
                If .Cells(Me.ColToleranceUp.Index).Value = 100 Then
                    .Cells(Me.ColToleranceUp.Index).Value = ObjCD.DigitGrouping(.Cells(Me.ColToleranceUpPerc.Index).Value / 100 * .Cells(Me.ColQtyMulti.Index).Value, DefPrecMulti, ClsCekData.EnumTipeRounding.Down)
                Else
                    .Cells(Me.ColToleranceUp.Index).Value = ObjCD.DigitGrouping(.Cells(Me.ColToleranceUpPerc.Index).Value / 100 * .Cells(Me.ColQtyMulti.Index).Value, DefPrecMulti)
                End If

            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    '\\TR#
    Public Function CekBatasanToleransi(ByVal vRowIndex As Integer) As Boolean
        Dim RowDgv As DataGridViewRow
        Dim decimalAngka As Integer = 0
        Try
            RowDgv = Me.dgvItem.Rows(vRowIndex)
            With RowDgv
                decimalAngka = ObjProduct.GetAngkaDecimal(.Cells(ColPrimaryCodeNo.Index).Value, .Cells(ColPrimaryUnitsMulti.Index).Value)
                If ObjCD.DigitGrouping(CDec(.Cells(Me.ColToleranceDownPerc.Index).Value) / 100 * CDec(.Cells(Me.ColQtyMulti.Index).Value), decimalAngka) >= CDec(.Cells(Me.ColQtyMulti.Index).Value) Then
                    .Cells(Me.ColToleranceDown.Index).Value = ObjCD.DigitGrouping(.Cells(Me.ColToleranceDownPerc.Index).Value / 100 * .Cells(Me.ColQtyMulti.Index).Value, decimalAngka)
                    MsgBox("The tolerance value (down) from the following product code cannot be >= from value of Qty.", MsgBoxStyle.Information, "Information")
                    Return False
                Else
                    If ObjCD.DigitGrouping(CDec(.Cells(Me.ColToleranceUpPerc.Index).Value / 100) * CDec(.Cells(Me.ColQtyMulti.Index).Value), decimalAngka) > CDec(.Cells(Me.ColQtyMulti.Index).Value) Then
                        .Cells(Me.ColToleranceUp.Index).Value = ObjCD.DigitGrouping(.Cells(Me.ColToleranceUpPerc.Index).Value / 100 * .Cells(Me.ColQtyMulti.Index).Value, decimalAngka)
                        MsgBox("The tolerance value (up) from the following product code cannot be greather than from value of Qty.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                    Return True
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
    Private Sub TxtQtyperBatchMulti_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchMulti.TextChanged
        Try
            If IsNumeric(Me.TxtQtyperBatchMulti.Text) = True And IsNumeric(Me.TxtDiscount.Text) = _
            True And IsNumeric(Me.TxtUnitPrice.Text) = True Then
                Me.Hitung_Amount()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LlbColumnSettings_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbColumnSettings.LinkClicked
        Dim ij As New FrmTransaksiSO_SetVisible
        Try
            ij.vUserName = Me.UserCode
            ij.IsMultiUnit = Me.IsMultiUnit
            If ij.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.UpdateKolomPO()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
        End Try
    End Sub

    Private Function IsMultiUnit() As Boolean
        Try
            For Each row As DataGridViewRow In Me.dgvItem.Rows
                With row
                    If .Cells(Me.ColPrimaryUnits.Index).Value <> .Cells(Me.ColPrimaryUnitsMulti.Index).Value Then
                        Return True
                    End If
                End With
            Next
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub UpdateKolomPO()
        Dim Dummy As String = ""
        Dim IsMultiUnit_ As Boolean = Me.IsMultiUnit
        Try
            Dummy = Me.ObjSettingUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.SalesOrder, _
            Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.SalesOrder_ShowUnitsDefault, Me.UserCode)
            If Dummy = "" Then
                Me.ColQty.Visible = False
                Me.ColUnit.Visible = False
                Me.ColRatioDesc.Visible = False
            Else
                If Dummy = 1 Then
                    Me.ColQty.Visible = True
                    Me.ColUnit.Visible = True
                    Me.ColRatioDesc.Visible = True
                    If IsMultiUnit_ = False Then
                        Me.ColQty.Visible = False
                        Me.ColUnit.Visible = False
                        Me.ColRatioDesc.Visible = False
                    End If
                Else
                    Me.ColQty.Visible = False
                    Me.ColUnit.Visible = False
                    Me.ColRatioDesc.Visible = False
                End If
            End If

            Dummy = Me.ObjSettingUser.GetSettingValue(Core_Module.ClsTransaksiSettingUser.EnumTipeTransaksi.SalesOrder, _
                        Core_Module.ClsTransaksiSettingUser.EnumNamaParameter.SalesOrder_ShowTolerance, Me.UserCode)
            If Dummy = "" Then
                Me.ColToleranceDown.Visible = False
                Me.ColToleranceDownPerc.Visible = False
                Me.ColToleranceUp.Visible = False
                Me.ColToleranceUpPerc.Visible = False
            Else
                If Dummy = 1 Then
                    Me.ColToleranceDown.Visible = True
                    Me.ColToleranceDownPerc.Visible = True
                    Me.ColToleranceUp.Visible = True
                    Me.ColToleranceUpPerc.Visible = True
                Else
                    Me.ColToleranceDown.Visible = False
                    Me.ColToleranceDownPerc.Visible = False
                    Me.ColToleranceUp.Visible = False
                    Me.ColToleranceUpPerc.Visible = False
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbUpdateAllDeliveryDate_LinkClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbUpdateAllDeliveryDate.LinkClicked
        Dim f As New FrmTransaksiSO_SetDeliveryDate
        Dim rowD As DseTransaksiSO.tbltrans_so_itemdeliveryRow

        Try
            If Me.dgvItem.RowCount = 0 Or Me.btnOK.Enabled = False Then
                Exit Sub
            End If
            Me.Cursor = Cursors.WaitCursor
            If f.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.vTBLItemDelivery.Clear()
                If f.RboSetDeliveryDate.Checked = True Then
                    For Each r As DataGridViewRow In Me.dgvItem.Rows
                        r.Cells(Me.ColDeliveryCont.Index).Value = True
                        r.Cells(Me.ColDeliveryDate.Index).Value = Format(f.DtpDate.Value.Date, "dd MMM yyyy")

                        'tambahkan ke dalam item delivery date
                        rowD = Me.vTBLItemDelivery.Newtbltrans_so_itemdeliveryRow
                        With rowD
                            .delivery_date = Format(f.DtpDate.Value.Date, "dd MMM yyyy")
                            .delivery_priority = Me.ObjDeliveryPriority.GetPrimaryDefault
                            .delivery_qty = r.Cells(Me.ColQty.Index).Value
                            .identifier = r.Cells(Me.ColIdentifier.Index).Value
                            .primary = Me.GetPrimaryMax(Me.vTBLItemDelivery) + 1
                            .primary_transaksi = Me.TxtPrimarySO.Text
                            .remarks = ""
                            .delivery_qty_multi = r.Cells(Me.ColQtyMulti.Index).Value
                        End With
                        Me.vTBLItemDelivery.Rows.Add(rowD)
                    Next
                    Me.dtpdeliverydate.Checked = True
                    Me.dtpdeliverydate.Value = Format(f.DtpDate.Value.Date, "dd MMM yyyy")
                Else
                    For Each r As DataGridViewRow In Me.dgvItem.Rows
                        r.Cells(Me.ColDeliveryCont.Index).Value = False
                        r.Cells(Me.ColDeliveryDate.Index).Value = "-"
                    Next
                    Me.dtpdeliverydate.Value = Now
                    Me.dtpdeliverydate.Checked = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    'Aul
    Private Sub EndEditDeliveryDateCont2(Optional ByVal vDel As Boolean = False)
        Dim rowD As DataGridViewRow

        Try
            If Me.dgvItem.RowCount > 1 Then
                If vDel = True Then
                    'jika fungsi ini dipanggil dari fungsi dgvItem (UserDeletingRow) item SO maka gunakan RowCount - 2
                    rowD = Me.dgvItem.Rows(Me.dgvItem.RowCount - 2)
                Else
                    'jika fungsi ini dipanggil dari fungsi dgvItem (CellEndEdit/CellContentClick/AktifkanDetail) 
                    'baris item SO maka gunakan RowCount - 1
                    rowD = Me.dgvItem.Rows(Me.dgvItem.RowCount - 1)
                End If

                If rowD.Cells(Me.ColDeliveryCont.Index).Value = Boolean.TrueString Then
                    Me.dtpdeliverydate.Checked = True
                    If IsDate(rowD.Cells(Me.ColDeliveryDate.Index).Value) Then
                        Me.dtpdeliverydate.Value = rowD.Cells(Me.ColDeliveryDate.Index).Value
                    ElseIf rowD.Cells(Me.ColDeliveryDate.Index).Value = "Multiple" Then
                        Me.vTBLItemDelivery.DefaultView.RowFilter = "identifier = '" + rowD.Cells(Me.ColIdentifier.Index).Value.ToString + "'"
                        If Me.vTBLItemDelivery.DefaultView.Count > 0 Then
                            vrowItemDelivery = Me.vTBLItemDelivery.DefaultView(0).Row
                            Me.dtpdeliverydate.Value = vrowItemDelivery.delivery_date
                        End If
                    End If
                Else
                    Me.dtpdeliverydate.Checked = False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function FromCutOff_SO(ByVal vRowSo As System.Data.DataRow, ByVal vTblSo As System.Data.DataTable, ByRef vTblSoItem As System.Data.DataTable, _
     ByRef vTblSoDel As System.Data.DataTable, ByVal vTblDec As System.Data.DataTable, ByVal vUserCode As String) As Double
        Dim rowSO_Item As System.Data.DataRow
        Dim rowSO_ItemDel As System.Data.DataRow
        Dim rowDec As System.Data.DataRow
        Dim rowD As DseTransaksiSO.tbltrans_so_itemdeliveryRow
        Dim rowDGV As DataGridViewRow
        Dim _QtyDelivery_ As Decimal = 0
        Dim _ExitFor_ As Boolean = True
        Dim _IdentifierDumm_, _IdentifierDumm_2, _IdentifierDumm_3 As String
        Dim DecimalProduct As Decimal = 0
        Dim vDate As New Date
        Dim vPrimaryQuickInvoice As Double = -1

        Try
            Me.NonAktifkanDetail()
            Me.baru = True
            Me.KosongkanField()
            Me.objSO.ResetObject(Me.dgvItem)
            If vTblSo.Rows.Count > 0 Then

                'isikan header
                With vRowSo
                    Me.cboOrderBy.SelectedIndex = .Item("Order By")
                    Me.CboFakturPajak.SelectedIndex = .Item("Jenis Faktur Pajak")
                    Me.CboPriceCondition.SelectedIndex = .Item("Price Condition")
                    Me.chkTax.Checked = .Item("Use Tax")
                    Me.ChkPriceCondition.Checked = .Item("Price Condtion Sama")
                    Me.ChkDiscountConditionSama.Checked = .Item("Discount Sama")
                    Me.txtPhoneNo.Text = .Item("Cust Phone No")
                    Me.txtCustomerAddIdentifier.Text = .Item("Indentifier Alamat Customer")
                    Me.TxtSalesOrderNo.Text = .Item("No Transaksi")
                    Me.txtPaymentDueDate.Text = .Item("Payment Due")
                    Me.txtPaymentTerm.Text = .Item("Payment Term")
                    Me.txtPrimaryCurrency.Text = .Item("Primary Currency")
                    Me.txtPrimaryCustomer.Text = .Item("Primary Customer")
                    Me.cboPaymentDueDateReference.SelectedIndex = .Item("Reference Payment Due")
                    Me.txtSubTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.Item("Sub Total"), Me.DefaultPrecision))
                    Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.Item("Tax Amount"), Me.DefaultPrecision))
                    Me.txtTax.Text = .Item("Tax")
                    Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.Item("Total"), Me.DefaultPrecision))
                    Me.txtRemarks.Text = .Item("Remarks")
                    vDate = vRowSo.Item("Tanggal Transaksi")
                    vPrimaryQuickInvoice = vRowSo.Item("Primary Quick Invoice")
                    Me.objSO.TanggalTransaksi = vRowSo.Item("Tanggal Transaksi")
                    Me.TxtShipVia.Text = .Item("Ship Via")
                    Me.txtReferenceDiscount.Text = .Item("Reference Discount")
                    Me.txtContactPerson.Text = .Item("Contact Person")
                    Me.TxtPrimarySO.Text = .Item("Primary")
                    Me.TxtExchangeRate.Text = Me.ObjCD.DigitGrouping(.Item("Exchange Rate"), Core_Module.DefPrecExchangeRate)
                    Me.TxtPrimarySalesP.Text = .Item("primary sales person")
                    Me.txtCustomerPONo.Text = vRowSo.Item("Cust PO No")
                End With

                'isikan item
                Dim PriceCondition As String = ""
                Dim DeliveryDate_Str As String = ""
                vTblSoItem.DefaultView.RowFilter = "[Primary Transaksi] = " & vRowSo.Item("Primary")
                For Each dv As DataRowView In vTblSoItem.DefaultView
                    rowSO_Item = dv.Row
                    With rowSO_Item
                        _IdentifierDumm_ = .Item("Identifier")
                        _IdentifierDumm_2 = .Item("Identifier")
                        .Item("Identifier") = "-1" & _IdentifierDumm_2.Remove(0, _IdentifierDumm_2.IndexOf("-"))
                        _IdentifierDumm_3 = .Item("Identifier")

                        If .Item("Price Exclude") = True Then
                            PriceCondition = "Exclude"
                        Else
                            PriceCondition = "Include"
                        End If
                        If CBool(.Item("Delivery Date Cont")) = True Then
                            If .Item("Delivery Date") = #1/1/1950# Then
                                DeliveryDate_Str = "Multiple"
                            Else
                                DeliveryDate_Str = Format(.Item("Delivery Date"), "dd MMM yyyy")
                            End If
                        Else
                            DeliveryDate_Str = "-"
                        End If
                        Me.dgvItem.Rows.Add()
                        rowDGV = dgvItem.Rows(dgvItem.RowCount - 1)
                        rowDGV.Cells(Me.ColAmount.Index).Value = .Item("Amount")
                        rowDGV.Cells(Me.ColDeliveryCont.Index).Value = CBool(.Item("Delivery Date Cont"))
                        rowDGV.Cells(Me.ColDeliveryDate.Index).Value = DeliveryDate_Str
                        rowDGV.Cells(Me.ColDeliveryHistory.Index).Value = ""
                        rowDGV.Cells(Me.ColDiscAmount.Index).Value = .Item("Disc Amount")
                        rowDGV.Cells(Me.ColDiscount.Index).Value = ObjCD.DigitGrouping(.Item("Discount"), 2)
                        rowDGV.Cells(Me.ColIdentifier.Index).Value = .Item("identifier")
                        rowDGV.Cells(Me.ColNettPrice.Index).Value = .Item("Nett Price")
                        rowDGV.Cells(Me.ColNo.Index).Value = dgvItem.Rows.Count
                        rowDGV.Cells(Me.ColPriceCond.Index).Value = PriceCondition
                        rowDGV.Cells(Me.ColPrimaryCodeNo.Index).Value = .Item("Primary Product")
                        rowDGV.Cells(Me.ColPrimaryUnits.Index).Value = .Item("Primary Units")
                        rowDGV.Cells(Me.ColQty.Index).Value = .Item("Qty")
                        rowDGV.Cells(Me.ColRemark.Index).Value = .Item("Remarks")
                        rowDGV.Cells(Me.ColSuppDesc.Index).Value = ""
                        rowDGV.Cells(Me.ColSupplierCodeNo.Index).Value = ""
                        rowDGV.Cells(Me.ColUnit.Index).Value = ""
                        rowDGV.Cells(Me.ColUnitPrice.Index).Value = .Item("Unit Price")
                        rowDGV.Cells(Me.ColUseSecondary.Index).Value = CBool(.Item("Use Secondary"))
                        rowDGV.Cells(Me.ColDiscFromPerct.Index).Value = .Item("Disc From Perct")

                        rowDGV.Cells(Me.ColQtyMulti.Index).Value = .Item("Qty Multi")
                        rowDGV.Cells(Me.ColPrimaryUnitsMulti.Index).Value = .Item("Primary Units Multi")
                        rowDGV.Cells(Me.ColToleranceDown.Index).Value = .Item("Tolerance Down")
                        rowDGV.Cells(Me.ColToleranceDownPerc.Index).Value = ObjCD.DigitGrouping(.Item("Tolerance Down Perc"), 2)
                        rowDGV.Cells(Me.ColToleranceDownFromPerc.Index).Value = ObjCD.DigitGrouping(.Item("Tolerance Down From Perc"), 0)
                        rowDGV.Cells(Me.ColToleranceUp.Index).Value = .Item("Tolerance Up")
                        rowDGV.Cells(Me.ColToleranceUpPerc.Index).Value = ObjCD.DigitGrouping(.Item("Tolerance Up Perc"), 2)
                        rowDGV.Cells(Me.ColToleranceUpFromPerc.Index).Value = ObjCD.DigitGrouping(.Item("Tolerance Up From Perc"), 0)
                        rowDGV.Cells(Me.ColRatioReference.Index).Value = .Item("Ratio Reference")
                        rowDGV.Cells(Me.ColRatioDefault.Index).Value = .Item("Ratio Default")

                        'Isikan tbl item delivery 
                        _QtyDelivery_ = .Item("Qty Multi")
                        _ExitFor_ = False
                        vTblSoDel.DefaultView.RowFilter = Nothing
                        vTblSoDel.DefaultView.RowFilter = "[Identifier] = '" & _IdentifierDumm_ & "'"
                        If vTblSoDel.DefaultView.Count > 0 Then
                            For ij As Integer = vTblSoDel.DefaultView.Count - 1 To 0 Step -1
                                rowSO_ItemDel = vTblSoDel.DefaultView(ij).Row
                                vTblDec.DefaultView.RowFilter = Nothing
                                vTblDec.DefaultView.RowFilter = "[Primary Product] = " & .Item("Primary Product")
                                If vTblDec.DefaultView.Count > 0 Then
                                    rowDec = vTblDec.DefaultView(0).Row
                                    DecimalProduct = rowDec.Item("Decimal Default")
                                End If

                                If _QtyDelivery_ > rowSO_ItemDel.Item("delivery qty multi") Then
                                    _QtyDelivery_ -= rowSO_ItemDel.Item("delivery qty multi")
                                Else
                                    rowSO_ItemDel.Item("delivery qty multi") = _QtyDelivery_
                                    _ExitFor_ = True
                                End If

                                rowD = Me.vTBLItemDelivery.Newtbltrans_so_itemdeliveryRow
                                rowD.delivery_date = CDate(rowSO_ItemDel.Item("delivery date"))
                                rowD.delivery_qty_multi = rowSO_ItemDel.Item("delivery qty multi")
                                rowD.delivery_qty = Math.Round(rowD.delivery_qty_multi * .Item("Ratio Default") / .Item("Ratio Reference"), 0, MidpointRounding.AwayFromZero)
                                rowD.identifier = _IdentifierDumm_3
                                rowD.primary = Me.GetPrimaryMax(Me.vTBLItemDelivery) + 1
                                rowD.primary_transaksi = rowSO_ItemDel.Item("primary transaksi")
                                rowD.remarks = rowSO_ItemDel.Item("remarks")
                                rowD.delivery_priority = Me.ObjDeliveryPriority.GetPrimaryDefault
                                vTBLItemDelivery.Rows.Add(rowD)
                                If _ExitFor_ = True Then
                                    Exit For
                                End If
                            Next
                        End If
                    End With
                Next

                'hitung yang sudah ada didalam dgv
                Me.Hitung_Total()
                'hitung item yang ada di gbo item
                Me.Hitung_Amount()
                'simpan
                Me.SalinKeObjek()
                Me.objSO.TanggalTransaksi = vDate
                Me.objSO.CreatedBy = vRowSo.Item("Created By")
                Me.objSO.CreatedDate = vRowSo.Item("Created Date")
                Me.objSO.ModifiedBy = vRowSo.Item("Modified By")
                Me.objSO.ModifiedDate = vRowSo.Item("Modified Date")
                Me.objSO.TimeStamp = vRowSo.Item("Timestamp")
                Me.objSO.SimpanBaru(Me.dgvItem, vUserCode, Me, Me.vTBLItemDelivery, False, True)
            End If

            Return Me.objSO.Primary
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        End Try
    End Function

    ''#######################################################################################################################################################################################

    Sub DefaultAddressShipToCondition()
        Try
            If Me.txtRemarks.Enabled = True Then
                CboAddressShipTo.Enabled = True
            End If
            LlbViewAddressShipTo.Enabled = True
            TxtCustomerCodeShipTo.Enabled = False
            BtnFindCustomerShipTo.Enabled = False
            CboAddressCustomerShipTo.Enabled = False
            LlbViewAddressCustomerShipTo.Enabled = False
            TxtPrimaryCustomerShipTo.Text = "-1"
            TxtCustomerCodeShipTo.Text = ""
            TxtCustomerDescShipTo.Text = ""
            TxtIdentifierAddressCustomerShipTo.Text = ""
            CboAddressCustomerShipTo.Items.Clear()
            ObjCustomer.IsiCboAddress(CboAddressShipTo, Me.txtPrimaryCustomer.Text)
            If CboAddressShipTo.Items.Count > 0 Then
                CboAddressShipTo.SelectedIndex = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Sub OtherCustomerAddressShipToCondition()
        Try
            CboAddressShipTo.Enabled = False
            LlbViewAddressShipTo.Enabled = False
            If Me.txtRemarks.Enabled = True Then
                TxtCustomerCodeShipTo.Enabled = True
                BtnFindCustomerShipTo.Enabled = True
                CboAddressCustomerShipTo.Enabled = True
            End If
            LlbViewAddressCustomerShipTo.Enabled = True
            TxtIdentifierAddressShipTo.Text = ""
            CboAddressShipTo.Items.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub RboAddressShipTo_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RboAddressShipTo.CheckedChanged
        Try
            If RboAddressShipTo.Checked Then
                DefaultAddressShipToCondition()
            Else
                OtherCustomerAddressShipToCondition()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub RboCustomerShipTo_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RboCustomerShipTo.CheckedChanged
        Try
            If RboCustomerShipTo.Checked Then
                OtherCustomerAddressShipToCondition()
            Else
                DefaultAddressShipToCondition()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnFindCustomerShipTo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFindCustomerShipTo.Click
        Dim i As New FrmCariMasterCustomer
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.TxtCustomerCodeShipTo.Text <> "" Then
                i.txtCari.Text = Me.TxtCustomerCodeShipTo.Text
            End If
            i.ShowDialog()
            If i.PrimaryCustomer > 0 Then
                Me.TxtPrimaryCustomerShipTo.Text = i.PrimaryCustomer
                Me.TxtCustomerCodeShipTo.Text = Me.ObjCustomer.GetCustomerCode(Me.TxtPrimaryCustomerShipTo.Text)
                Me.TxtCustomerDescShipTo.Text = Me.ObjCustomer.GetCustName(Me.TxtPrimaryCustomerShipTo.Text)
                Me.ObjCustomer.IsiCboAddress(Me.CboAddressCustomerShipTo, Me.TxtPrimaryCustomerShipTo.Text)
                Me.CboAddressCustomerShipTo.SelectedIndex = 0
                Me.CboAddressCustomerShipTo.Focus()
            Else
                Me.TxtCustomerCodeShipTo.Clear()
                Me.TxtCustomerDescShipTo.Clear()
                Me.CboAddressCustomerShipTo.Items.Clear()
                Me.TxtIdentifierAddressCustomerShipTo.Text = ""
                Me.TxtPrimaryCustomerShipTo.Text = "-1"
                Me.TxtCustomerCodeShipTo.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub LlbViewAddressShipTo_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbViewAddressShipTo.LinkClicked
        Try
            If Me.txtRemarks.Enabled = True OrElse (IsNumeric(Me.TxtPrimarySO.Text) AndAlso CDec(Me.TxtPrimarySO.Text) > 0) Then
                If Me.txtPrimaryCustomer.Text = "" Or Me.txtPrimaryCustomer.Text = "-1" Then
                    MsgBox("Please specify customer first.", MsgBoxStyle.Information, "Information")
                    Me.txtCustomerCode.Focus()
                    Exit Sub
                End If
                Dim ij As New Core_Module.FrmDetailAddress
                ij.vAddIdentifier = Me.TxtIdentifierAddressShipTo.Text
                ij.PrimaryCust = Me.txtPrimaryCustomer.Text
                ij.btnOK.BringToFront()
                ij.AcceptButton = ij.btnOK
                ij.btnExit.TabStop = False
                If Me.txtRemarks.Enabled = False Then
                    ij.cboCustAddCode.Enabled = False
                End If
                ij.ShowDialog()
                If ij.DialogResult = Windows.Forms.DialogResult.OK Then
                    Me.ObjCustomer.IsiCboAddress(Me.CboAddressShipTo, Me.txtPrimaryCustomer.Text)
                    Me.CboAddressShipTo.Text = Me.ObjCustomer.GetKodeAddress(ij.vAddIdentifierResult)
                End If
                ij.Dispose()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LlbViewAddressCustomerShipTo_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbViewAddressCustomerShipTo.LinkClicked
        Try
            If Me.txtRemarks.Enabled = True OrElse (IsNumeric(Me.TxtPrimarySO.Text) AndAlso CDec(Me.TxtPrimarySO.Text) > 0) Then
                If Me.TxtPrimaryCustomerShipTo.Text = "" Or Me.TxtPrimaryCustomerShipTo.Text = "-1" Then
                    MsgBox("Please specify ship to customer first.", MsgBoxStyle.Information, "Information")
                    Me.TxtCustomerCodeShipTo.Focus()
                    Exit Sub
                End If
                Dim ij As New Core_Module.FrmDetailAddress
                ij.vAddIdentifier = Me.TxtIdentifierAddressCustomerShipTo.Text
                ij.PrimaryCust = Me.TxtPrimaryCustomerShipTo.Text
                ij.btnOK.BringToFront()
                ij.AcceptButton = ij.btnOK
                ij.btnExit.TabStop = False
                If Me.txtRemarks.Enabled = False Then
                    ij.cboCustAddCode.Enabled = False
                End If
                ij.ShowDialog()
                If ij.DialogResult = Windows.Forms.DialogResult.OK Then
                    Me.ObjCustomer.IsiCboAddress(Me.CboAddressCustomerShipTo, Me.TxtPrimaryCustomerShipTo.Text)
                    Me.CboAddressCustomerShipTo.Text = Me.ObjCustomer.GetKodeAddress(ij.vAddIdentifierResult)
                End If
                ij.Dispose()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCustomerCodeShipTo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCustomerCodeShipTo.GotFocus
        TxtCustomerCodeShipTo.BackColor = BackColorAll
    End Sub

    Private Sub TxtCustomerCodeShipTo_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCustomerCodeShipTo.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.ObjCustomer.ApakahSudahAda(Me.TxtCustomerCodeShipTo.Text) = True Then
                    Me.TxtPrimaryCustomerShipTo.Text = Me.ObjCustomer.GetPrimary(Me.TxtCustomerCodeShipTo.Text)
                    Me.TxtCustomerDescShipTo.Text = Me.ObjCustomer.GetCustName(Me.TxtPrimaryCustomerShipTo.Text)
                    Me.ObjCustomer.IsiCboAddress(Me.CboAddressCustomerShipTo, Me.TxtPrimaryCustomerShipTo.Text)
                    Me.CboAddressCustomerShipTo.SelectedIndex = 0
                    Me.TxtShipVia.Focus()
                Else
                    Me.BtnFindCustomerShipTo.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCustomerCodeShipTo_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtCustomerCodeShipTo.Leave
        Dim i As New FrmCariMasterCustomer
        i.BackColorAll = Me.BackColorAll

        Try
            If Me.TxtCustomerCodeShipTo.Text <> "" Then
                If Me.ObjCustomer.ApakahSudahAda(Me.TxtCustomerCodeShipTo.Text) = True Then
                    Me.TxtPrimaryCustomerShipTo.Text = Me.ObjCustomer.GetPrimary(Me.TxtCustomerCodeShipTo.Text)
                    Me.TxtCustomerDescShipTo.Text = Me.ObjCustomer.GetCustName(Me.TxtPrimaryCustomerShipTo.Text)
                    Me.ObjCustomer.IsiCboAddress(Me.CboAddressCustomerShipTo, Me.TxtPrimaryCustomerShipTo.Text)
                    Me.CboAddressCustomerShipTo.SelectedIndex = 0
                Else
                    If Me.TxtCustomerCodeShipTo.Text <> "" Then
                        i.txtCari.Text = Me.TxtCustomerCodeShipTo.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCustomer > 0 Then
                        Me.TxtPrimaryCustomerShipTo.Text = i.PrimaryCustomer
                        Me.TxtCustomerCodeShipTo.Text = Me.ObjCustomer.GetCustomerCode(Me.TxtPrimaryCustomerShipTo.Text)
                        Me.TxtCustomerDescShipTo.Text = Me.ObjCustomer.GetCustName(Me.TxtPrimaryCustomerShipTo.Text)
                        Me.ObjCustomer.IsiCboAddress(Me.CboAddressCustomerShipTo, Me.TxtPrimaryCustomerShipTo.Text)
                        Me.CboAddressCustomerShipTo.SelectedIndex = 0
                    Else
                        Me.TxtCustomerCodeShipTo.Clear()
                        Me.TxtCustomerDescShipTo.Clear()
                        Me.CboAddressCustomerShipTo.Items.Clear()
                        Me.TxtIdentifierAddressCustomerShipTo.Text = ""
                        Me.TxtPrimaryCustomerShipTo.Text = "-1"
                    End If
                End If
            Else
                Me.TxtCustomerCodeShipTo.Clear()
                Me.TxtCustomerDescShipTo.Clear()
                Me.CboAddressCustomerShipTo.Items.Clear()
                Me.TxtIdentifierAddressCustomerShipTo.Text = ""
                Me.TxtPrimaryCustomerShipTo.Text = "-1"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub CboAddressShipTo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboAddressShipTo.SelectedIndexChanged
        Try
            If Me.txtRemarks.Enabled = True And IsNumeric(Me.txtPrimaryCustomer.Text) = True Then
                If Me.CboAddressShipTo.SelectedIndex <> -1 Then
                    Me.TxtIdentifierAddressShipTo.Text = Me.ObjCustomer.GetIdentifierAdd(Me.txtPrimaryCustomer.Text, Me.CboAddressShipTo.Text)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub CboAddressCustomerShipTo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboAddressCustomerShipTo.SelectedIndexChanged
        Try
            If Me.txtRemarks.Enabled = True Then
                If Me.CboAddressCustomerShipTo.SelectedIndex <> -1 Then
                    Me.TxtIdentifierAddressCustomerShipTo.Text = Me.ObjCustomer.GetIdentifierAdd(Me.TxtPrimaryCustomerShipTo.Text, Me.CboAddressCustomerShipTo.Text)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCustomerCodeShipTo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCustomerCodeShipTo.LostFocus
        TxtCustomerCodeShipTo.BackColor = Me.txtPrimaryCategory.BackColor
    End Sub

    Private Sub chkUseDP_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkUseDP.CheckedChanged
        Try
            Me.SetDPControl()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetDPControl()
        Try
            If Me.txtRemarks.Enabled = True Then
                If Me.chkUseDP.Checked = False Then
                    Me.TxtDPFromPercent.Text = "1"
                    Me.TxtDPPercent.Text = "0.00"
                    Me.TxtDPAmount.Text = "0"
                    Me.TxtDPPercent.ReadOnly = True
                    Me.TxtDPPercent.TabStop = False
                    Me.TxtDPPercent.BackColor = Me.TxtReadOnly.BackColor
                    Me.TxtDPAmount.ReadOnly = True
                    Me.TxtDPAmount.TabStop = False
                    Me.TxtDPAmount.BackColor = Me.TxtReadOnly.BackColor
                Else
                    Me.TxtDPPercent.ReadOnly = False
                    Me.TxtDPPercent.TabStop = True
                    Me.TxtDPPercent.BackColor = Me.txtPrimaryCategory.BackColor
                    Me.TxtDPAmount.ReadOnly = False
                    Me.TxtDPAmount.TabStop = True
                    Me.TxtDPAmount.BackColor = Me.txtPrimaryCategory.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Txt_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDPAmount.GotFocus, TxtDPPercent.GotFocus
        Dim Txt As TextBox

        Try
            Txt = TryCast(sender, Windows.Forms.TextBox)
            If IsNothing(Txt) = False Then
                If Txt.Enabled = True And Me.txtRemarks.Enabled = True Then
                    If Txt.ReadOnly = False Then
                        Txt.BackColor = Me.BackColorAll
                    Else
                        Txt.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Txt_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDPAmount.LostFocus, TxtDPPercent.LostFocus
        Dim Txt As TextBox

        Try
            Txt = TryCast(sender, Windows.Forms.TextBox)
            If IsNothing(Txt) = False Then
                If Txt.Enabled = True And Me.txtRemarks.Enabled = True Then
                    If Txt.ReadOnly = False Then
                        Txt.BackColor = Me.txtPrimaryCategory.BackColor
                    Else
                        Txt.BackColor = Me.TxtReadOnly.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDPAmount_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDPAmount.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtDPAmount, sender, e)
            Me.buffer = Me.TxtDPAmount.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDPAmount_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDPAmount.Leave
        Try
            Me.ObjCD.Number_Leave(Me.TxtDPAmount, sender, e, Me.DefaultPrecision, Me.ObjCD)

            If CDec(Me.TxtDPAmount.Text) <> CDec(Me.buffer) Then
                If CDec(Me.TxtDPAmount.Text) > CDec(Me.txtSubTotal.Text) Then
                    MsgBox("Maximum DP amount is " + Me.txtCurrencyCode.Text + " " + Me.txtSubTotal.Text + ".", MsgBoxStyle.Information, "Information")
                    Exit Sub
                Else
                    Me.TxtDPFromPercent.Text = "0"
                    Me.Hitung_DP()
                End If
            Else
                'do nothing
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDPAmount_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDPAmount.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtDPAmount, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDPAmount_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDPAmount.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtDPAmount, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDPPercent_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDPPercent.Enter
        Try
            Me.ObjCD.Number_Enter(Me.TxtDPPercent, sender, e)
            Me.buffer = Me.TxtDPPercent.Text
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDPPercent_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDPPercent.Leave
        Try
            Me.ObjCD.Number_Leave(Me.TxtDPPercent, sender, e, 2, Me.ObjCD)

            If CDec(Me.buffer) <> (Me.TxtDPPercent.Text) Then
                Me.TxtDPFromPercent.Text = 1
                Me.Hitung_DP()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDPPercent_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDPPercent.KeyDown
        Try
            Me.ObjCD.Number_KeyDown(Me.TxtDPPercent, sender, e, 100)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDPPercent_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtDPPercent.KeyUp
        Try
            Me.ObjCD.Number_KeyUp(Me.TxtDPPercent, sender, e, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub cboPriceByMultiUnit_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboPriceByMultiUnit.SelectedIndexChanged
        Dim UnitPrice, Discount As Decimal

        Try
            UnitPrice = 0
            Discount = 0

            If Me.cboPriceByMultiUnit.SelectedIndex <> -1 Then
                If Me.TxtRemarkItem.Enabled = True Then
                    If Me.cboPriceByMultiUnit.Text <> Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNo.Text) Then
                        Me.txtPriceByMultiUnit.Text = "1"

                        Me.ObjProduct.GetUnitPriceNDiskon(Me.txtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtQtyperBatchMulti.Text, _
                        Me.txtPrimaryCurrency.Text, Me.txtPrimaryUnitsMulti.Text, UnitPrice, Discount)

                        If UnitPrice = 0 And Discount = 0 Then
                            Me.objSO.GetHistoricalPrice(Me.txtPrimaryUnitsMulti.Text, Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryCurrency.Text, _
                            UnitPrice, Discount, Me.txtPriceByMultiUnit.Text, Me.txtPrimaryCustomer.Text)
                        End If

                        Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(UnitPrice, Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text))
                        Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Discount, 2)
                    Else
                        Me.txtPriceByMultiUnit.Text = "0"

                        Me.ObjProduct.GetUnitPriceNDiskon(Me.txtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtQtyperBatchDefault.Text, _
                        Me.txtPrimaryCurrency.Text, Me.txtPrimaryUnitsDefault.Text, UnitPrice, Discount)

                        If UnitPrice = 0 And Discount = 0 Then
                            Me.objSO.GetHistoricalPrice(Me.txtPrimaryUnitsDefault.Text, Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryCurrency.Text, _
                            UnitPrice, Discount, Me.txtPriceByMultiUnit.Text, Me.txtPrimaryCustomer.Text)
                        End If

                        Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(UnitPrice, Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text))

                        Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Discount, 2)
                    End If


                    If IsNumeric(Me.TxtUnitPrice.Text) = True Then
                        If CDec(Me.TxtUnitPrice.Text) > 0 Then
                            Me.Hitung_Amount()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    '\\TR=Quo
    Private Sub CekProduct()
        Dim rowQuo As DataGridViewRow
        Dim rowItem As DataGridViewRow
        Dim rowDGV As DataGridViewRow
        Dim pesanText As String = ""
        Dim showException As Boolean = False
        Dim FrmException As New FrmTransaksiSO_Exception
        Dim listProduct As New List(Of Double)
        Dim counter As Integer = 0
        Try
            For Each rowQuo In Me.dgvItemQuo.Rows
                pesanText = ""
                For Each rowItem In Me.dgvItem.Rows
                    If rowItem.Cells(Me.ColPrimaryCodeNo.Index).Value = rowQuo.Cells(Me.ColQuo_PrimaryCodeNo.Index).Value Then
                        If rowItem.Cells(Me.ColUnitsMulti.Index).Value <> rowQuo.Cells(Me.ColQuo_UnitsMulti.Index).Value Then
                            pesanText = "Unit"
                        End If
                        'multi unit price
                        If rowItem.Cells(Me.ColPriceByMultiUnit_Str.Index).Value <> rowQuo.Cells(Me.ColQuoPriceByMultiUnit_Str.Index).Value Then
                            If pesanText = "" Then
                                pesanText = "Price by Unit"
                            Else
                                pesanText = pesanText & ", price by Unit"
                            End If
                        End If
                        'unit price
                        If rowItem.Cells(Me.ColUnitPrice.Index).Value <> rowQuo.Cells(Me.ColQuo_UnitPrice.Index).Value Then
                            If pesanText = "" Then
                                pesanText = "Unit price"
                            Else
                                pesanText = pesanText & ", unit price"
                            End If
                        End If
                        'discount
                        If CDec(rowItem.Cells(Me.ColDiscount.Index).Value) <> CDec(rowQuo.Cells(Me.ColQuo_Discount.Index).Value) Then
                            If pesanText = "" Then
                                pesanText = "Discount"
                            Else
                                pesanText = pesanText & ", discount"
                            End If
                        End If
                        'discount amount
                        If CDec(rowItem.Cells(Me.ColDiscAmount.Index).Value) <> CDec(rowQuo.Cells(Me.ColQuo_DiscAmount.Index).Value) Then
                            If pesanText = "" Then
                                pesanText = "Discount amount"
                            Else
                                pesanText = pesanText & ", discount amount"
                            End If
                        End If
                        'nett price
                        If CDec(rowItem.Cells(Me.ColNettPrice.Index).Value) <> CDec(rowQuo.Cells(Me.ColQuo_NettPrice.Index).Value) Then
                            If pesanText = "" Then
                                pesanText = "Nett price"
                            Else
                                pesanText = pesanText & ", nett price"
                            End If
                        End If
                    End If
                    If pesanText <> "" Then
                        If listProduct.Contains(rowQuo.Cells(Me.ColQuo_PrimaryCodeNo.Index).Value) = False Then
                            listProduct.Add(rowQuo.Cells(Me.ColQuo_PrimaryCodeNo.Index).Value)
                        End If
                        pesanText = "Different on : " & pesanText & "."
                        showException = True
                        FrmException.dgvItem.Rows.Add()
                        rowDGV = FrmException.dgvItem.Rows(FrmException.dgvItem.Rows.Count - 1)
                        rowDGV.Cells(FrmException.ColNo.Index).Value = FrmException.dgvItem.Rows.Count
                        rowDGV.Cells(FrmException.ColProductCode.Index).Value = rowQuo.Cells(Me.ColQuo_CodeNo.Index).Value
                        rowDGV.Cells(FrmException.ColProdDesc.Index).Value = rowQuo.Cells(Me.ColQuo_Description.Index).Value
                        rowDGV.Cells(FrmException.ColReason.Index).Value = pesanText
                        GoTo skip
                    End If
                Next
skip:
            Next

            If showException = True Then
                FrmException.ShowDialog()
                While counter <= Me.dgvItemQuo.Rows.Count - 1
                    rowQuo = Me.dgvItemQuo.Rows.Item(counter)
                    If listProduct.Contains(rowQuo.Cells(Me.ColQuo_PrimaryCodeNo.Index).Value) = True Then
                        Me.dgvItemQuo.Rows.Remove(rowQuo)
                    Else
                        counter += 1
                    End If
                End While
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub ShowTransaction(ByVal vPrimary As Double)
        Try
            If Me.txtRemarks.Enabled = True Then
                Me.btnCancel.PerformClick()
            End If
            If Me.txtRemarks.Enabled = False Then
                If Me.objSO.CariDetail(vPrimary, Me.dgvItem, Me.vTBLItemDelivery) _
                 = True Then
                    Me.SalinKeField()
                    Me.btnEdit.Focus()
                Else
                    Me.btnFind.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
