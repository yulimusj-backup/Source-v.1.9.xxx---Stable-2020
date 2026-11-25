Option Explicit On
Option Strict Off

Imports Core_Module

Public Class FrmTransaksiQuotation
    Dim mTblMultiUnit As New DseTransaksiQuotation.tblmaster_product_item_multiunitDataTable
    Dim tblCboUnit As New DseTransaksiQuotation.CboUnitDataTable

    Public UserCode As String = ""
    Public BackColorAll As Color = Core_Module.GetSMESSBackColorAll
    Public FormNo As Integer = 0
    Public vTBLItemDelivery As New DseTransaksiQuotation.tbltrans_quotation_itemdeliveryDataTable
    Dim vrowItemDelivery As DseTransaksiQuotation.tbltrans_quotation_itemdeliveryRow

    Dim baru As Boolean
    Dim objQuotation As New ClsTransaksiQuotation
    Dim ObjCD As New ClsCekData
    Dim ObjProduct As New clsMasterProduct
    Dim ObjCustomer_sProductCode As New clsTransaksiSettingcustomerProductCodeNo
    Dim objQuotationOutItem As New ClsTransaksiQuotationItem
    Dim ObjCurrency As New Core_Module.clsMasterCurrency
    Dim ObjUser As New Core_Module.ClsUser
    Dim ObjCustomer As New clsMasterCustomer
    Dim ObjLicense As New ClsLicense
    Dim ObjDeliveryPriority As New Core_Module.ClsMasterPriorityLevel
    Dim ObjProductCustom As New CustomSMESS.ClsMasterProduct
    Dim objCustom As New CustomSMESS.ClsTransaksiQuotation

    Dim DefaultPrecision As Integer = 0
    Dim buffer As String = ""
    Dim ObjSetting As New Core_Module.ClsTransaksiSetting

    Private Sub AktifkanDetail()
        Try
            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.TxtPrimaryQuotation.BackColor, False)
            Me.ObjCD.SetBackColor(Me.TabPage1.Controls, Me.TxtPrimaryQuotation.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel1.Controls, Me.TxtPrimaryQuotation.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel2.Controls, Me.TxtPrimaryQuotation.BackColor, False)
            Me.ObjCD.SetBackColor(Me.TabPage2.Controls, Me.TxtPrimaryQuotation.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gboTerm.Controls, Me.TxtPrimaryQuotation.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gboItem.Controls, Me.TxtPrimaryQuotation.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel1.Controls, Me.TxtPrimaryQuotation.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel2.Controls, Me.TxtPrimaryQuotation.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.TxtPrimaryQuotation.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.TxtPrimaryQuotation.BackColor, False)

            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.TabPage1.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.TabPage2.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.gboTerm.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.gboItem.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer3.Panel1.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer3.Panel2.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, True, True, True, True, True, True, True, True, True)
            Me.ObjCD.KendalikanControl(Me.Controls, False, False, True, False, False, False, False, False)
            Me.BtnFindCodeNo.Enabled = True
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
            Me.llbUpdateAllDeliveryDate.Enabled = True
            Dim ij As New System.EventArgs
            If Me.chkUseCustCodeNo.Checked = True Then
                Me.btnFindCustomerProductCode.Enabled = True
                Me.BtnFindCodeNo.Enabled = False
            Else
                Me.btnFindCustomerProductCode.Enabled = False
                Me.BtnFindCodeNo.Enabled = True
            End If
            Me.dtpdeliverydate.Checked = False

            Me.EndEditDeliveryDateCont2()
            Me.SetDPControl()
            Me.RboAddressShipTo.Enabled = True
            Me.RboCustomerShipTo.Enabled = True
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
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub NonAktifkanDetail()
        Try
            Me.ObjCD.KendalikanControl(Me.gboDetail.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.TabPage1.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer2.Panel1.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer2.Panel2.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.TabPage2.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.gboTerm.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.gboItem.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer3.Panel1.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer3.Panel2.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel1.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.SplitContainer1.Panel2.Controls, True, False, True, False, True, False, True, False, True, False)
            Me.ObjCD.KendalikanControl(Me.Controls, False, False, True, True, False, False, False, False)
            Me.CancelButton = Me.btnExit
            Me.BtnFindCodeNo.Enabled = False
            Me.dgvItem.AllowUserToDeleteRows = False
            Me.dgvItem.ReadOnly = True
            Me.RboAddressShipTo.Enabled = False
            Me.RboCustomerShipTo.Enabled = False
            'Me.llbAddress.Enabled = False
            Me.llbDetailStock.Enabled = False
            Me.llbUpdateAllDeliveryDate.Enabled = False

            Me.ObjCD.SetBackColor(Me.gboDetail.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.TabPage1.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel1.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer2.Panel2.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.TabPage2.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gboTerm.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.gboItem.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel1.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer3.Panel2.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel1.Controls, Me.BackColor, False)
            Me.ObjCD.SetBackColor(Me.SplitContainer1.Panel2.Controls, Me.BackColor, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanField()
        Try
            Me.KosongkanItem()
            ModCekData.KosongkanControl(Me.gboDetail.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.TabPage1.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer2.Panel1.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer2.Panel2.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.TabPage2.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer1.Panel1.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer1.Panel2.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.gboTerm.Controls, True, True, False)
            Me.TxtPrimaryQuotation.Text = "-1"
            Me.txtPrimaryCurrency.Text = "-1"
            Me.txtPrimaryCustomer.Text = "-1"
            Me.txtPaymentDueDate.Text = "0"
            Me.txtSubTotal.Text = "0"
            Me.txtTax.Text = "0"
            Me.txtTaxAmount.Text = "0"
            Me.TxtTotal.Text = "0"
            Me.tslCreatedBy.Text = "Created By: "
            Me.tslCreatedDate.Text = "Created Date: "
            Me.tslModifiedBy.Text = "Modified By: "
            Me.tslModifiedDate.Text = "Modified Date: "
            Me.dgvItem.Rows.Clear()
            Me.cboPaymentDueDateReference.SelectedIndex = 0
            Me.chkTax.Checked = True
            Me.cboCustomerAddress.DataSource = Nothing
            Me.cboCustomerAddress.Items.Clear()
            Me.CboAddressCustomerShipTo.DataSource = Nothing
            Me.CboAddressCustomerShipTo.Items.Clear()
            Me.cboCustomerAddress.Text = ""
            Me.CboAddressCustomerShipTo.Text = ""
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
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub KosongkanItem()
        Dim dummy As Boolean = False

        Try
            dummy = Me.dtpdeliverydate.Checked

            ModCekData.KosongkanControl(Me.gboItem.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer3.Panel1.Controls, True, True, False)
            ModCekData.KosongkanControl(Me.SplitContainer3.Panel2.Controls, True, True, False)
            Me.TxtPrimaryCodeNo.Text = "-1"
            Me.txtPrimaryUnitsDefault.Text = "-1"
            Me.txtPrimaryUnitsMulti.Text = "-1"
            Me.TxtQtyperBatchMulti.Text = "-1"
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
            Me.lblUnitsMulti.Text = "Units"
            Me.LblUnitPrice.Text = "Unit Price"

            Me.dtpdeliverydate.Checked = dummy

            Me.cboPriceByMultiUnit.Visible = False
            Me.cboPriceByMultiUnit.TabStop = False
            Me.txtPriceByMultiUnit.Text = "1"
            Me.cboPriceByMultiUnit.Items.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function ApakahDataItemBenar() As Boolean
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
                Me.TxtQtyperBatchMulti.Focus()
                Return False
            End If
            '2.b. nilai qty <= 0
            If Me.ObjCD.CekBilanganPositif(Me.TxtQtyperBatchDefault.Text) = False Then
                MsgBox("Please enter a positive number for quantity default data.", MsgBoxStyle.Information, _
                "Information")
                Me.TxtQtyperBatchMulti.Focus()
                Return False
            End If
            If Me.ObjCD.CekBilanganPositif(Me.TxtUnitPrice.Text) = False Then
                MsgBox("Please enter a positive number for unit price data.", MsgBoxStyle.Information, _
                "Information")
                Me.TxtUnitPrice.Focus()
                Return False
            End If
            '3. Bila sudah ada product pada dgv dan unit price, discount, nett price-nya, 
            'delivery datenya beda - saat ini satu product hanya bisa muncul satu kali pada 1 so
            If Me.GetIndex(Me.TxtPrimaryCodeNo.Text) <> -1 Then
                If Me.GetIndex(Me.TxtPrimaryCodeNo.Text, Me.dtpdeliverydate.Checked, Me.dtpdeliverydate.Value, _
                Me.TxtUnitPrice.Text, Me.txtNettPrice.Text, Me.TxtDiscount.Text) = -1 Then
                    MsgBox(Me.TxtCodeNo.Text + " is already enterred, double product code no in one quotation is not supported", _
                    MsgBoxStyle.Information, "Information")
                    Return False
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

            If objCustom.ApakahDataBenarItem(Me.TxtPrimaryCodeNo.Text, Me.txtPriceByMultiUnit.Text, Me.txtPrimaryUnitsMulti.Text, Me.txtPrimaryUnitsDefault.Text) = False Then
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJmlDGV(ByVal vPrimaryProduct As Double, ByVal DeliveryDateCont As Boolean, ByVal DeliveryDate As Date, _
    ByVal UnitPrice As Double, ByVal NettPrice As Double, ByVal Discount As Double) As Double
        Try
            Dim row As DataGridViewRow

            For Each row In Me.dgvItem.Rows
                If DeliveryDateCont = True Then
                    If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And row.Cells(Me.ColDeliveryCont.Index).Value = Boolean.TrueString And _
                    row.Cells(Me.ColDeliveryDate.Index).Value = DeliveryDate And row.Cells(Me.ColUnitPrice.Index).Value = UnitPrice And _
                    row.Cells(Me.ColDiscount.Index).Value = Discount And row.Cells(Me.ColNettPrice.Index).Value = NettPrice Then
                        Return row.Cells(Me.ColQty.Index).Value
                    End If
                Else
                    If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And row.Cells(Me.ColDeliveryCont.Index).Value = Boolean.FalseString And _
                    row.Cells(Me.ColDeliveryDate.Index).Value = DeliveryDate And row.Cells(Me.ColUnitPrice.Index).Value = UnitPrice And _
                    row.Cells(Me.ColDiscount.Index).Value = Discount And row.Cells(Me.ColNettPrice.Index).Value = NettPrice Then
                        Return row.Cells(Me.ColQty.Index).Value
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
        Dim Dummy As String = ""
        Try
            If Me.ObjLicense.GetExpiryDate(My.Settings.b) <= Me.dtpDate.Value Then
                MsgBox("This aplication has expire please contact your software vendor to renew your license.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            'cek data yang harus diisi
            KoleksiNamaData.Add("Quotation number")
            KoleksiTextBox.Add(Me.TxtQuoNo)
            KoleksiNamaData.Add("Customer code")
            KoleksiTextBox.Add(Me.txtCustomerCode)
            KoleksiNamaData.Add("Currency")
            KoleksiTextBox.Add(Me.txtCurrencyCode)

            If ModCekData.CekDataKosong(KoleksiTextBox, KoleksiNamaData) = False Then
                Return False
            End If

            If Me.txtPaymentTerm.Text = "" Then
                Me.TabControl1.SelectedIndex = 0
                Me.txtPaymentTerm.Focus()
                MsgBox("Payment term cannot be an empty string.", MsgBoxStyle.Information, "Information")
                Me.txtPaymentTerm.Focus()
                Return False
            End If

            'cek data no transaksi yang salah - tidak perlu cek karena ini adalah autonumber

            'minimal harus punya 1 data item
            If Me.dgvItem.RowCount = 0 Then
                MsgBox("Please enter at least 1 item.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            'kalo sampe cek berkaitan dengan timestamp apakah masih benar
            If Me.baru = False Then
                If Me.objQuotation.TimeStamp <> Me.objQuotation.GetTimestamp(Me.objQuotation.Primary) Then
                    If MsgBox("Other user have modified this data, reload latest data?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        Me.objQuotation.CariDetail(Me.objQuotation.Primary, Me.dgvItem, Me.vTBLItemDelivery)
                        Me.SalinKeField()
                    End If
                    Return False
                End If
            End If

            'cek berkaitan dengan perubahan tanggal
            If Me.baru = False Then
                If Me.objQuotation.TanggalTransaksi.Month = Me.dtpDate.Value.Month And _
                Me.objQuotation.TanggalTransaksi.Year = Me.dtpDate.Value.Year Then
                    Me.TxtQuoNo.Text = Me.objQuotation.NoTransaksi
                Else
                    If MsgBox("You just changed transaction date, this will change your transaction number and there's a possibility your transaction no cannot be change back. Continue and accept this change?", _
                    MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.No Then
                        Me.dtpDate.Value = Me.objQuotation.TanggalTransaksi
                        Me.TxtQuoNo.Text = Me.objQuotation.NoTransaksi
                        Return False
                    End If
                End If
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

            If CDec(Me.txtSubTotal.Text) < CDec(Me.TxtDPAmount.Text) Then
                MsgBox("Maximum DP amount is " + Me.txtCurrencyCode.Text + " " + Me.txtSubTotal.Text + ".", MsgBoxStyle.Information, "Information")
                Return False
            End If
            If Me.chkUseDP.Checked = True And CDec(Me.TxtDPAmount.Text) = 0 Then
                MsgBox("Please specify DP amount.", MsgBoxStyle.Information, "Information")
                Return False
            End If
            '8. Jika Ship to Customer dipilih, maka customer ship to wajib diisi
            If Me.RboCustomerShipTo.Checked And (Me.TxtPrimaryCustomerShipTo.Text = "" Or Me.TxtPrimaryCustomerShipTo.Text = "-1" Or Me.TxtIdentifierAddressCustomerShipTo.Text = "") Then
                MsgBox("If you want to ship to other customer, you have to define a customer and address.", MsgBoxStyle.Information, "Information")
                Me.TabControl1.SelectedIndex = 1
                Me.TxtCustomerCodeShipTo.Focus()
                Return False
            End If

            If Me.RboAddressShipTo.Checked And (Me.txtPrimaryCustomer.Text = Me.TxtPrimaryCustomerShipTo.Text) Then
                MsgBox("If you want to ship to other customer, please define different customer for goods shipping.", MsgBoxStyle.Information, "Information")
                Me.TxtCustomerCodeShipTo.Focus()
                Me.TabControl1.SelectedIndex = 1
                Return False
            End If
            If Me.RboCustomerShipTo.Checked And (Me.txtPrimaryCustomer.Text = Me.TxtPrimaryCustomerShipTo.Text) Then
                MsgBox("If you want to ship to other customer, please define different customer for goods shipping.", MsgBoxStyle.Information, "Information")
                Me.TxtCustomerCodeShipTo.Focus()
                Return False
            End If
            If objCustom.ApakahDataBenarCustom(Me.dgvItem, Me.ColCodeNo.Index, Me.ColPrimaryCodeNo.Index, Me.chkTax.Checked, _
                       Me.ColPrimaryUnits.Index, Me.ColPrimaryUnitsMulti.Index, Me.ColPrimaryCodeNo.Index, Me.ColPriceByMultiUnit.Index) = False Then
                Return False
            End If

            If Me.chkTax.Checked = True Then
                If CDec(Me.txtTax.Text) <= 0 Then
                    Me.TabControl1.SelectedIndex = 0
                    Me.txtTax.Focus()
                    MsgBox("Please specify tax percentage.", MsgBoxStyle.Information, "Information")
                    Return False
                End If

                Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.TaxPercentage)
                If Dummy = "" Then
                    Dummy = "0"
                End If

                If CDec(Me.txtTax.Text) <> CDec(Dummy) Then
                    Me.TabControl1.SelectedIndex = 0
                    Me.txtTax.Focus()
                    If MsgBox("Tax rate <> " + Me.ObjCD.DigitGrouping(Dummy, 2) + "%, continue to save data?", MsgBoxStyle.YesNo, "Confirm") <> MsgBoxResult.Yes Then
                        Return False
                    End If
                End If
            End If

            'Comment: Perubahan NPWP 16 Digit
            If chkTax.Checked = True Then 'hanya diperiksa ketika checkbox tax bernilai true
                Dim ObjSetNPWP As New Core_Module.ClsTrans_SettingNPWPNo
                If ObjSetNPWP.SetNPWP(Me.dtpDate.Value.Date, Me.txtPrimaryCustomer.Text, Me.txtNoNPWP.Text, ClsTrans_SettingNPWPNo.EnumFromMaster.Customer) = False Then
                    Return False
                End If
            End If
            'kalo sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            KoleksiNamaData = Nothing
            KoleksiTextBox = Nothing
        End Try
    End Function

    Private Sub SalinKeObjek()
        Try
            'salin data header
            With Me.objQuotation
                If Me.baru = True Then
                    .AuthorizationLevel = 0
                    .AuthorizationLevelNeeded = 0
                    .GeneralManagerAuthorizationDate = #1/1/1950#
                    .GeneralManagerAuthorizationPrimary = -1
                    .SalesManagerAuthorizationDate = #1/1/1950#
                    .SalesManagerAuthorizationPrimary = -1
                    .SalesStaffAuthorizationDate = #1/1/1950#
                    .SalesStaffAuthorizationPrimary = -1
                End If
                .NamaSigned = Me.TxtNamaSigned.Text
                .JabatanSigned = Me.TxtJabatanSigned.Text
                .CustContactPerson = Me.txtContactPerson.Text
                .CustInquiryNo = Me.TxtInquiryNo.Text
                .CustNPWP = Me.txtNoNPWP.Text
                .CustPhone = Me.txtPhoneNo.Text
                .CustIdentifierAdd = Me.txtCustomerAddIdentifier.Text
                .NoTransaksi = Me.TxtQuoNo.Text
                .PaymentDue = Me.txtPaymentDueDate.Text
                .PaymentTerm = Me.txtPaymentTerm.Text
                .PrimaryCurrency = Me.txtPrimaryCurrency.Text
                .PrimaryCustomer = Me.txtPrimaryCustomer.Text
                .RefPaymentDate = Me.cboPaymentDueDateReference.SelectedIndex
                .SubTotal = Me.txtSubTotal.Text
                .TaxAmount = Me.txtTaxAmount.Text
                .TaxPercentage = Me.txtTax.Text
                .TermCondition = Me.txtTermsNCondition.Text
                .Total = Me.TxtTotal.Text
                .UseTax = Me.chkTax.Checked
                .Remarks = Me.txtRemarks.Text
                .TanggalTransaksi = Me.dtpDate.Value
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
            With Me.objQuotation
                Me.TxtJabatanSigned.Text = .JabatanSigned
                Me.TxtNamaSigned.Text = .NamaSigned
                Me.txtContactPerson.Text = .CustContactPerson
                Me.TxtInquiryNo.Text = .CustInquiryNo
                Me.txtNoNPWP.Text = .CustNPWP
                Me.txtPhoneNo.Text = .CustPhone
                Me.TxtPrimaryQuotation.Text = .Primary
                Me.TxtQuoNo.Text = .NoTransaksi
                Me.txtPaymentDueDate.Text = .PaymentDue
                Me.txtPaymentTerm.Text = .PaymentTerm
                Me.txtPrimaryCurrency.Text = .PrimaryCurrency
                Me.txtCurrencyCode.Text = Me.ObjCurrency.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                Me.SetDefaultPrecision(False)
                Me.txtPrimaryCustomer.Text = .PrimaryCustomer
                Me.txtCustomerAddIdentifier.Text = .CustIdentifierAdd
                Me.txtCustomerCode.Text = Me.ObjCustomer.GetCustomerCode(Me.txtPrimaryCustomer.Text)
                Me.TxtCustomerName.Text = Me.ObjCustomer.GetCustName(Me.txtPrimaryCustomer.Text)
                Me.ObjCustomer.IsiCboAddress(Me.cboCustomerAddress, Me.txtPrimaryCustomer.Text)
                Me.cboCustomerAddress.Text = Me.ObjCustomer.GetKodeAddress(Me.txtCustomerAddIdentifier.Text)
                Me.cboPaymentDueDateReference.SelectedIndex = .RefPaymentDate
                Me.txtSubTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.SubTotal, Me.DefaultPrecision))
                Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.TaxAmount, Me.DefaultPrecision))
                Me.txtTax.Text = .TaxPercentage
                Me.txtTermsNCondition.Text = .TermCondition
                Me.TxtTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(.Total, Me.DefaultPrecision))
                Me.chkTax.Checked = .UseTax
                Me.txtRemarks.Text = .Remarks
                Me.dtpDate.Value = .TanggalTransaksi
                If Me.objQuotation.AuthorizationLevel = 0 Then
                    Me.txtAuthorizationStatus.Text = "Unauthorized"
                ElseIf Me.objQuotation.AuthorizationLevel = 1 Then
                    Me.txtAuthorizationStatus.Text = "Authorized By Sales Staff"
                ElseIf Me.objQuotation.AuthorizationLevel = 2 Then
                    Me.txtAuthorizationStatus.Text = "Authorized By Sales Manager"
                ElseIf Me.objQuotation.AuthorizationLevel = 3 Then
                    Me.txtAuthorizationStatus.Text = "Authorized By Director"
                End If

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



                Me.tslCreatedBy.Text = "Created By: " + .CreatedBy
                Me.tslCreatedDate.Text = "Created Date: " + Format(.CreatedDate, "dd-MMM-yyyy") + "-" + .CreatedDate.ToShortTimeString
                Me.tslModifiedBy.Text = "Modified By: " + .ModifiedBy
                Me.tslModifiedDate.Text = "Modified Date: " + Format(.ModifiedDate, "dd-MMM-yyyy") + "-" + .ModifiedDate.ToShortTimeString
            End With
            Me.FormatDGV()
            'digit grouping
            Dim row As DataGridViewRow
            For Each row In Me.dgvItem.Rows
                row.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColQty.Index).Value)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub frmMasterTipe_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.baru = True
            Me.NonAktifkanDetail()
            Me.dgvItem.AlternatingRowsDefaultCellStyle.BackColor = Me.BackColorAll

            'set tampilan use cust code no
            Me.txtCustomerProductCodeNo.ReadOnly = True
            Me.txtCustomerProductCodeNo.BackColor = Me.TxtQuoNo.BackColor
            Me.txtCustomerProductCodeNo.TabStop = False
            Me.TxtCodeNo.ReadOnly = False
            Me.TxtCodeNo.BackColor = Me.txtPrimaryCategory.BackColor
            Me.TxtCodeNo.TabStop = True
            Me.BtnFindCodeNo.Enabled = False
            Me.btnFindCustomerProductCode.Enabled = False
            Me.ColSuppDesc.Visible = False
            Me.ColSupplierCodeNo.Visible = False
            Me.dgvItem.ScrollBars = ScrollBars.Both
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnSimpan_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.ApakahDataBenar() = True Then
                Me.KosongkanItem()
                Me.SalinKeObjek()
                If baru = True Then
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.objQuotation.SimpanBaru(Me.dgvItem, Me.UserCode, Me, Me.vTBLItemDelivery) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnFind.Focus()
                        End If
                    End If
                Else
                    'konfirmasi disimpan atau tidak
                    hasil = MsgBox("Save?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.objQuotation.SimpanPerubahan(Me.dgvItem, Me.UserCode, Me, Me.vTBLItemDelivery) Then
                            MsgBox("Data saved", MsgBoxStyle.Information, "Information")
                            Me.SalinKeField()
                            Me.NonAktifkanDetail()
                            Me.btnEdit.Focus()
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
                        Me.objQuotation.CariDetail(Me.TxtPrimaryQuotation.Text, Me.dgvItem, Me.vTBLItemDelivery)
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
        Dim Dummy As String = ""
        Try
            Me.baru = True
            Me.KosongkanField()
            Me.AktifkanDetail()
            Me.TabControl1.SelectedIndex = 0
            Me.dtpDate.Focus()
            Me.objQuotation.ResetObject(Me.dgvItem)
            Me.TxtQuoNo.Text = Me.objQuotation.GetNoBaru(Me.dtpDate.Value)
            Me.txtAuthorizationStatus.Text = "Unautorized"

            Me.tslCreatedBy.Text = "Created By: " + Me.UserCode
            Me.tslCreatedDate.Text = "Created Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString
            Me.tslModifiedBy.Text = "Modified By: " + Me.UserCode
            Me.tslModifiedDate.Text = "Modified Date: " + Format(Now, "dd-MMM-yyyy") + "-" + Now.ToShortTimeString

            'kalo new mata uang langsung IDR
            Me.txtPrimaryCurrency.Text = "1"
            Me.txtCurrencyCode.Text = Me.ObjCurrency.GetCurrencyCode(1)

            Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseTax)
            If Dummy = "" Then
                Me.chkTax.Checked = False
                Me.txtTax.Text = Me.ObjCD.DigitGrouping("0", 2)
            Else
                Me.chkTax.Checked = True
                Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.TaxPercentage)
                If Dummy = "" Then
                    Dummy = "0"
                End If
                Me.txtTax.Text = Me.ObjCD.DigitGrouping(Dummy, 2)
            End If
            'Me.chkTax.Checked = True
            'Me.txtTax.Text = "10"
            Me.cboPaymentDueDateReference.SelectedIndex = 1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFind.Click
        Dim hj As New FrmCariTransaksiQuotation
        Try
            hj.BackColorAll = Me.BackColorAll
            hj.ShowDialog()
            If Not hj.PrimaryTransaksi = -1 Then
                If Me.objQuotation.CariDetail(hj.PrimaryTransaksi, Me.dgvItem, Me.vTBLItemDelivery) _
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
        Try
            If Me.TxtQuoNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtQuoNo.Text <> "" Then
                If Me.objQuotation.ApakahBolehEdit(Me.objQuotation.Primary, True, Me.UserCode, True) = True Then
                    Me.baru = False
                    Me.objQuotation.CariDetail(Me.objQuotation.Primary, Me.dgvItem, Me.vTBLItemDelivery)
                    Me.SalinKeField()
                    Me.AktifkanDetail()
                    Me.dgvItem.Focus()
                    'nilai min juga harus diatas tanggal pengkuncian data terakhir
                    'msgbox("belum")
                Else
                    Me.btnFind.Focus()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnHapus_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim hasil As MsgBoxResult
        Try
            If Me.TxtQuoNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtQuoNo.Text <> "" Then
                If Me.objQuotation.ApakahBolehHapus(Me.objQuotation.Primary) = False Then
                    MsgBox("You cannot delete this data, cancel transaction authorization first.", _
                    MsgBoxStyle.Information, "Information")
                    Me.btnFind.Focus()
                Else
                    hasil = MsgBox("Delete this data?", MsgBoxStyle.YesNo, "Confirm")
                    If hasil = MsgBoxResult.Yes Then
                        Windows.Forms.Cursor.Current = Cursors.WaitCursor
                        If Me.objQuotation.Hapus(Me.TxtPrimaryQuotation.Text) = True Then
                            Me.KosongkanField()
                            Me.btnNew.Focus()
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
            If Me.txtRemarks.Enabled = True Then
                Me.BringToFront()
                Dim hasil As MsgBoxResult = MsgBox("Save before exit?", MsgBoxStyle.YesNoCancel, "Confirm")
                If hasil = MsgBoxResult.Yes Then
                    If Me.ApakahDataBenar() = True Then
                        Me.SalinKeObjek()
                        If baru = True Then
                            'simpan baru
                            Windows.Forms.Cursor.Current = Cursors.WaitCursor
                            If Me.objQuotation.SimpanBaru(Me.dgvItem, Me.UserCode, Me, Me.vTBLItemDelivery) = True Then
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
                            If Me.objQuotation.SimpanPerubahan(Me.dgvItem, Me.UserCode, Me, Me.vTBLItemDelivery) = True Then
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

    Private Sub SetGbo()
        Dim UnitPrice, Diskon As Decimal

        Try
            Me.lblUnitsDefault.Text = Me.ObjProduct.GetUnitCode(Me.TxtPrimaryCodeNo.Text)
            Me.txtPrimaryUnitsDefault.Text = Me.ObjProduct.GetPrimaryUnit(Me.TxtPrimaryCodeNo.Text)
            Me.chkUseSecondary.Checked = Me.ObjProduct.ApakahUseSecondary(Me.TxtPrimaryCodeNo.Text)
            If Me.chkUseSecondary.Checked = True Then
                Me.LblUnitPrice.Text = "Unit Price/" + Me.ObjProduct.GetSecondaryUnitCode(Me.TxtPrimaryCodeNo.Text)
                Me.Hitung_Amount()
            Else
                Me.SetCboPriceByMultiUnit()
                'Me.LblUnitPrice.Text = "Unit Price/" + Me.lblUnitsMulti.Text
            End If
            Me.ObjProduct.GetUnitPriceNDiskon(Me.txtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtQtyperBatchMulti.Text, _
            Me.txtPrimaryCurrency.Text, Me.txtPrimaryUnitsMulti.Text, UnitPrice, Diskon)

            Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(UnitPrice, _
            Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text))
            Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Diskon, 2)



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
            
            Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(UnitPrice, Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text))
            Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Discount, 2)

          

            Me.Hitung_Amount()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Private Sub TxtCodeNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCodeNo.KeyDown
        Dim i As New frmCariMasterProduct
        Dim Jml As Long = 0
        i.BackColorAll = Me.BackColorAll

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
                    Me.SetGbo()
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
                        Me.SetGbo()
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
        'Dim UnitPrice, Diskon As Decimal

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
                    Me.SetGbo()
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
                        Me.SetGbo()
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

        Try
            'cek apakah sudah isi kode currency
            i.BackColorAll = Me.BackColorAll
            If Me.txtPrimaryCustomer.Text <> "" Then
                If Me.txtPrimaryCustomer.Text = -1 And Me.txtPrimaryCustomer.Text <> "" Then
                    Me.TabControl1.SelectedIndex = 1
                    Me.txtCustomerCode.Focus()
                    MsgBox("To start entering item data, you specify customer data first.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            If Me.txtPrimaryCurrency.Text <> "" Then
                If Me.txtPrimaryCurrency.Text = -1 And Me.txtPrimaryCurrency.Text <> "" Then
                    Me.TabControl1.SelectedIndex = 0
                    Me.txtCurrencyCode.Focus()
                    MsgBox("To start entering item data, you specify currency data first.", MsgBoxStyle.Information, "Information")
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
                Me.SetGbo()
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

    Private Function GetIndex(ByVal vPrimaryProduct As Double, ByVal DeliveryDateCont As Boolean, ByVal DeliveryDate As Date, _
    ByVal UnitPrice As Double, ByVal NettPrice As Double, ByVal Discount As Double) As Integer
        Dim row As DataGridViewRow

        Try
            For Each row In Me.dgvItem.Rows
                If CBool(row.Cells(Me.ColDeliveryCont.Index).Value) = True Then
                    If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And CBool(row.Cells(Me.ColDeliveryCont.Index).Value) = DeliveryDateCont And _
                            row.Cells(Me.ColUnitPrice.Index).Value = UnitPrice And row.Cells(Me.ColDiscount.Index).Value = Discount And _
                            row.Cells(Me.ColNettPrice.Index).Value = NettPrice Then
                        Return row.Index
                    End If
                Else
                    If row.Cells(Me.ColPrimaryCodeNo.Index).Value = vPrimaryProduct And CBool(row.Cells(Me.ColDeliveryCont.Index).Value) = DeliveryDateCont And _
                    row.Cells(Me.ColUnitPrice.Index).Value = UnitPrice And _
                    row.Cells(Me.ColDiscount.Index).Value = Discount And row.Cells(Me.ColNettPrice.Index).Value = NettPrice Then
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

    Private Sub BtnInput_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnInput.Click
        Dim DefPrecDefault As Integer = 0
        Dim DefPrecMulti As Integer = 0
        Dim rowDGV As DataGridViewRow
        Dim rowDel As DseTransaksiQuotation.tbltrans_quotation_itemdeliveryRow
        Dim IdentifierBaru As String = ""
        Try
            If Me.ApakahDataItemBenar = True Then
                Dim Index As Integer = Me.GetIndex(Me.TxtPrimaryCodeNo.Text, Me.dtpdeliverydate.Checked, Me.dtpdeliverydate.Value, _
                 Me.TxtUnitPrice.Text, Me.txtNettPrice.Text, Me.TxtDiscount.Text)

                DefPrecMulti = ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text)
                DefPrecDefault = Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text)

                'cari data yang sama persis
                If Index <> -1 Then
                    'tambahkan nilai qty
                    Dim row As DataGridViewRow = Me.dgvItem.Rows(Index)
                    'row.Cells(Me.ColQty.Index).Value = CDbl(row.Cells(Me.ColQty.Index).Value) + CDbl(Me.TxtQtyperBatchDefault.Text)
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
                    With rowDGV
                        .Cells(Me.ColNo.Index).Value = Me.dgvItem.Rows.Count
                        .Cells(Me.ColCodeNo.Index).Value = Me.TxtCodeNo.Text
                        .Cells(Me.ColDescription.Index).Value = Me.txtDesc.Text
                        .Cells(Me.ColSupplierCodeNo.Index).Value = Me.txtCustomerProductCodeNo.Text
                        .Cells(Me.ColSuppDesc.Index).Value = Me.txtCustomerProductDesc.Text
                        .Cells(Me.ColQty.Index).Value = Me.TxtQtyperBatchDefault.Text
                        .Cells(Me.ColUnit.Index).Value = Me.lblUnitsDefault.Text
                        .Cells(Me.ColUnitPrice.Index).Value = Me.TxtUnitPrice.Text
                        .Cells(Me.ColDiscount.Index).Value = Me.TxtDiscount.Text
                        .Cells(Me.ColDiscAmount.Index).Value = Me.TxtDiscAmount.Text
                        .Cells(Me.ColNettPrice.Index).Value = Me.txtNettPrice.Text
                        .Cells(Me.ColAmount.Index).Value = Me.txtAmount.Text
                        .Cells(Me.ColDeliveryCont.Index).Value = Me.dtpdeliverydate.Checked
                        .Cells(Me.ColDeliveryDate.Index).Value = ExpiredDateStr
                        .Cells(Me.ColRemark.Index).Value = Me.TxtRemarkItem.Text
                        .Cells(Me.ColPrimaryCodeNo.Index).Value = Me.TxtPrimaryCodeNo.Text
                        .Cells(Me.ColPrimaryUnits.Index).Value = Me.txtPrimaryUnitsDefault.Text
                        .Cells(Me.ColUseSecondary.Index).Value = Me.chkUseSecondary.Checked
                        .Cells(Me.ColDiscFromPerct.Index).Value = Me.TxtDiscFromPerct.Text
                        .Cells(Me.ColIdentifier.Index).Value = IdentifierBaru

                        .Cells(Me.ColQtyMulti.Index).Value = Me.TxtQtyperBatchMulti.Text
                        .Cells(Me.ColUnitsMulti.Index).Value = Me.lblUnitsMulti.Text
                        .Cells(Me.ColPrimaryUnitsMulti.Index).Value = Me.txtPrimaryUnitsMulti.Text
                        .Cells(Me.ColRatioDefault.Index).Value = Me.ObjCD.DigitGrouping(Me.objQuotation.GetRatioDefault(.Cells(Me.ColPrimaryCodeNo.Index).Value, _
                        .Cells(Me.ColPrimaryUnitsMulti.Index).Value), DefPrecDefault)
                        .Cells(Me.ColRatioReference.Index).Value = Me.ObjCD.DigitGrouping(Me.objQuotation.GetRatioReference(.Cells(Me.ColPrimaryCodeNo.Index).Value, _
                        .Cells(Me.ColPrimaryUnitsMulti.Index).Value), DefPrecMulti)
                        Me.SetRatioOnDgv(rowDGV)
                        .Cells(Me.ColToleranceDown.Index).Value = Me.ObjCD.DigitGrouping("0", DefPrecMulti)
                        .Cells(Me.ColToleranceUp.Index).Value = Me.ObjCD.DigitGrouping("0", DefPrecMulti)
                        .Cells(Me.ColToleranceDownPerc.Index).Value = Me.ObjCD.DigitGrouping("0", 2)
                        .Cells(Me.ColToleranceUpPerc.Index).Value = Me.ObjCD.DigitGrouping("0", 2)
                        .Cells(Me.ColToleranceDownFromPerc.Index).Value = 1
                        .Cells(Me.ColToleranceUpFromPerc.Index).Value = 1

                        .Cells(Me.ColPriceByMultiUnit.Index).Value = Me.txtPriceByMultiUnit.Text
                        If Me.txtPriceByMultiUnit.Text = "1" Then
                            .Cells(Me.ColPriceByMultiUnit_Str.Index).Value = "/" + .Cells(Me.ColUnitsMulti.Index).Value
                        Else
                            .Cells(Me.ColPriceByMultiUnit_Str.Index).Value = "/" + .Cells(Me.ColUnit.Index).Value
                        End If
                    End With

                    'masukkan data item delivery
                    If Me.dtpdeliverydate.Checked = True Then
                        rowDel = Me.vTBLItemDelivery.Newtbltrans_quotation_itemdeliveryRow
                        With rowDel
                            .delivery_date = Me.dtpdeliverydate.Value
                            .delivery_priority = Me.ObjDeliveryPriority.GetPrimaryDefault
                            .delivery_qty = Me.TxtQtyperBatchDefault.Text
                            .identifier = Me.dgvItem.Item(Me.ColIdentifier.Index, Me.dgvItem.RowCount - 1).Value
                            .primary = Me.GetPrimaryMax(Me.vTBLItemDelivery) + 1
                            .primary_transaksi = Me.TxtPrimaryQuotation.Text
                            .remarks = ""
                            .delivery_qty_multi = Me.TxtQtyperBatchMulti.Text
                        End With
                        vTBLItemDelivery.Rows.Add(rowDel)
                    End If

                    'set ulang nomor
                    Me.ObjCD.SetNumber(Me.dgvItem)
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
                'format dgv
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

    Private Sub FormatDGV()
        Dim row As DataGridViewRow
        Dim IsMultiUnit As Boolean = False
        Dim PriceByUnitDefault As Boolean = False

        Try
            'yang diformat dgv-nya:
            For Each row In Me.dgvItem.Rows
                '1.a. untuk delivery date expectation yang nilainya sama dengan 1/1/1950 dibuat "-"
                If IsDate(row.Cells(ColDeliveryDate.Index).Value) = False Then
                    If row.Cells(ColDeliveryCont.Index).Value = True Then
                        row.Cells(ColDeliveryDate.Index).Value = "Multiple"
                    Else
                        row.Cells(ColDeliveryDate.Index).Value = "-"
                    End If
                Else
                    If CDate(row.Cells(ColDeliveryDate.Index).Value).Date = #1/1/1950# Then
                        If row.Cells(ColDeliveryCont.Index).Value = True Then
                            row.Cells(ColDeliveryDate.Index).Value = "Multiple"
                        Else
                            row.Cells(ColDeliveryDate.Index).Value = "-"
                        End If
                    Else
                        'format sehingga bisa dibaca lebih enak
                        row.Cells(ColDeliveryDate.Index).Value = Format(CDate(row.Cells(ColDeliveryDate.Index).Value), "dd MMM yyyy")
                    End If
                End If
                'digit grouping: qty, unit price, nett price, amount
                If IsNumeric(row.Cells(Me.ColQty.Index).Value) = True Then
                    row.Cells(Me.ColQty.Index).Value = Me.ObjCD.DigitGrouping(row.Cells(Me.ColQty.Index).Value)
                End If
                If IsNumeric(row.Cells(Me.ColUnitPrice.Index).Value) = True Then
                    row.Cells(Me.ColUnitPrice.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(row.Cells(Me.ColUnitPrice.Index).Value, Me.DefaultPrecision))
                End If
                If IsNumeric(row.Cells(Me.ColDiscAmount.Index).Value) = True Then
                    row.Cells(Me.ColDiscAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(row.Cells(Me.ColDiscAmount.Index).Value, Me.DefaultPrecision))
                End If
                If IsNumeric(row.Cells(Me.ColNettPrice.Index).Value) = True Then
                    row.Cells(Me.ColNettPrice.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(row.Cells(Me.ColNettPrice.Index).Value, Me.DefaultPrecision))
                End If
                If IsNumeric(row.Cells(Me.ColAmount.Index).Value) = True Then
                    row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(row.Cells(Me.ColAmount.Index).Value, Me.DefaultPrecision))
                End If

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
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'Private Sub TxtQtyperBatch_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchDefault.Enter
    '    Try
    '        Me.ObjCD.Number_Enter(Me.TxtQtyperBatchDefault, sender, e)
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    'Private Sub TxtQtyperBatch_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyperBatchDefault.KeyDown
    '    Me.ObjCD.Number_KeyDown(Me.TxtQtyperBatchDefault, sender, e)
    'End Sub

    'Private Sub TxtQtyperBatch_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtQtyperBatchDefault.KeyUp
    '    Me.ObjCD.Number_KeyUp(Me.TxtQtyperBatchDefault, sender, e)
    'End Sub

    Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit
        Try
            'yang bisa diedit adalah:
            '1. Qty
            '2. Unit Price
            '3. Diskon
            '4. Delivery Date Controll
            '5. Delivery Date
            '6. Remarks
            If Me.txtRemarks.Enabled = True Then
                If e.ColumnIndex = Me.ColQtyMulti.Index Or e.ColumnIndex = Me.ColUnitPrice.Index Or e.ColumnIndex = Me.ColDiscount.Index Or e.ColumnIndex = Me.ColDiscAmount.Index _
                Or e.ColumnIndex = Me.ColDeliveryCont.Index Or e.ColumnIndex = Me.ColDeliveryDate.Index Or e.ColumnIndex = Me.ColRemark.Index Or e.ColumnIndex = Me.ColToleranceDown.Index Or _
                e.ColumnIndex = Me.ColToleranceDownPerc.Index Or e.ColumnIndex = Me.ColToleranceUp.Index Or e.ColumnIndex = Me.ColToleranceUpPerc.Index Then
                    'untuk qty, unit price, diskon, nett price 
                    If e.ColumnIndex = Me.ColDeliveryDate.Index Then
                        'cek apakah expexted delivery date dikontrol
                        If CStr(Me.dgvItem.Item(Me.ColDeliveryCont.Index, e.RowIndex).Value) = Boolean.FalseString Then
                            e.Cancel = True
                            Exit Sub
                        End If
                    End If
                    'kalo sampe sini maka boleh edit
                    Me.buffer = Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value
                Else
                    e.Cancel = True
                End If
            Else
                'hanya bisa edit pada saat detail aktif
                e.Cancel = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

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

            Retval = Me.TxtPrimaryQuotation.Text + "-" + NilaiMax.ToString
            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Function GetPrimaryMax(ByVal vTBL As DseTransaksiQuotation.tbltrans_quotation_itemdeliveryDataTable) As Double
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

    Private Sub dgvItem_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellContentClick
        Dim ij As New FrmTransaksiQuotation_SetDelivery
        Dim DI As Integer = 0
        Dim rowD As DseTransaksiQuotation.tbltrans_quotation_itemdeliveryRow
        Dim JmlDelivery As Integer = 0
        Dim DelDateD As Date = Today

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
                            rowD = Me.vTBLItemDelivery.Newtbltrans_quotation_itemdeliveryRow
                            With rowD
                                .delivery_date = CDate(rowDGV.Cells(ij.ColDeliveryDate.Index).Value)
                                .delivery_priority = rowDGV.Cells(ij.ColPrimaryPriority.Index).Value
                                .delivery_qty = rowDGV.Cells(ij.ColQty.Index).Value
                                .identifier = ij.VIdentifier
                                .primary = Me.GetPrimaryMax(Me.vTBLItemDelivery) + 1
                                .primary_transaksi = Me.TxtPrimaryQuotation.Text
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

                'ElseIf e.ColumnIndex = Me.ColDeliveryHistory.Index Then
                '    Me.objSO.ExplainDelivery(Me.dgvItem.Item(Me.ColIdentifier.Index, e.RowIndex).Value)
            End If

            If e.ColumnIndex = Me.ColRatioDesc.Index Then
                Dim FrmRatio As New FrmTransaksiQuotation_SetQtyRatio
                Dim row As DataGridViewRow

                If Me.dgvItem.Item(Me.ColPrimaryUnits.Index, e.RowIndex).Value = Me.dgvItem.Item(Me.ColPrimaryUnitsMulti.Index, e.RowIndex).Value Or _
                    Me.txtTermsNCondition.Enabled = False Then
                    Exit Sub
                End If
                Dim BufferRatioReference As String = Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value.ToString
                Dim BufferRatioDefault As String = Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value.ToString
                FrmRatio.vDefPrecMulti = ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value, Me.dgvItem.Item(Me.ColPrimaryUnitsMulti.Index, e.RowIndex).Value)
                frmRatio.vDefPrecDefault = Me.ObjProduct.GetAngkaDecimal(Me.dgvItem.Item(Me.ColPrimaryCodeNo.Index, e.RowIndex).Value)
                frmRatio.vRatioMulti = Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value
                FrmRatio.vUnitMulti = Me.dgvItem.Item(Me.ColUnitsMulti.Index, e.RowIndex).Value
                frmRatio.vRatioDefault = Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value
                frmRatio.vUnitDefault = Me.dgvItem.Item(Me.ColUnit.Index, e.RowIndex).Value
                If frmRatio.ShowDialog = Windows.Forms.DialogResult.OK Then
                    Me.dgvItem.Item(Me.ColRatioReference.Index, e.RowIndex).Value = frmRatio.TxtRatioMulti.Text
                    Me.dgvItem.Item(Me.ColRatioDefault.Index, e.RowIndex).Value = frmRatio.TxtRatioDefault.Text
                    row = Me.dgvItem.Rows(e.RowIndex)
                    Me.SetRatioOnDgv(row)
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
                        ElseIf e.ColumnIndex = Me.ColUnitPrice.Index Then
                            'unit price
                            Me.EndEditUnitPrice(sender, e, row)
                        ElseIf e.ColumnIndex = Me.ColDiscount.Index Then
                            'discount
                            Me.EndEditDiscount(sender, e, row)
                        ElseIf e.ColumnIndex = Me.ColDeliveryCont.Index Then
                            'kontrol delivery date
                            Me.EndEditDeliveryDateCont(sender, e, row)
                        ElseIf e.ColumnIndex = Me.ColDeliveryDate.Index Then
                            'delivery date
                            Me.EndEditDeliveryDate(sender, e, row)
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
        Dim rowD As DseTransaksiQuotation.tbltrans_quotation_itemdeliveryRow
        Try
            With row
                If row.Cells(Me.ColDeliveryCont.Index).Value = Boolean.TrueString Then
                    'jika jadi true maka delivery date jadi tanggal transaksi
                    row.Cells(Me.ColDeliveryDate.Index).Value = Format(Me.GetPreviousRowDeliveryDate(e.RowIndex - 1), "dd MMM yyyy")
                    'tambahkan ke dalam item delivery date
                    rowD = Me.vTBLItemDelivery.Newtbltrans_quotation_itemdeliveryRow
                    With rowD
                        .delivery_date = Me.GetPreviousRowDeliveryDate(e.RowIndex - 1)
                        .delivery_priority = Me.ObjDeliveryPriority.GetPrimaryDefault
                        .delivery_qty = Me.dgvItem.Item(Me.ColQty.Index, e.RowIndex).Value
                        .identifier = Me.dgvItem.Item(Me.ColIdentifier.Index, e.RowIndex).Value
                        .primary = Me.GetPrimaryMax(Me.vTBLItemDelivery) + 1
                        .primary_transaksi = Me.TxtPrimaryQuotation.Text
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

                'Me.EndEditDeliveryDateCont2()
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub EndEditDeliveryDate(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal row As DataGridViewRow)
        Try
            With row
                'pastikan format yang dituliskan adalah format tanggal
                If IsDate(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) = False Then
                    Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                    Exit Sub
                End If
                'format ulang tanggal sehingga formatnya konsisten
                Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Format(CDate(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value), _
                "dd MMM yyyy")
                'cek apakah back date terhadap tanggal transaksi, konfirmasikan data ini
                If CDate(Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value) < Me.dtpDate.Value.Date Then
                    If MsgBox("This product delivery date is smaller than transaction date. Do you want to continue?", _
                    MsgBoxStyle.YesNo, "Confimation") = MsgBoxResult.No Then
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                End If
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

                DefPrecMulti = ObjProduct.GetAngkaDecimal(row.Cells(Me.ColPrimaryCodeNo.Index).Value, row.Cells(Me.ColPrimaryUnitsMulti.Index).Value)
                If row.Cells(Me.ColToleranceDownFromPerc.Index).Value = "1" Then
                    row.Cells(Me.ColToleranceDown.Index).Value = ObjCD.DigitGrouping(CDec(row.Cells(Me.ColQtyMulti.Index).Value) * CDec(row.Cells(Me.ColToleranceDownPerc.Index).Value) / 100, DefPrecMulti)
                Else
                    If CDec(row.Cells(Me.ColToleranceDown.Index).Value) > (CDec(row.Cells(e.ColumnIndex).Value) / 2) Then
                        MsgBox("Qty you entered less than 50% of tolerance down. Change tolerance down first.", MsgBoxStyle.Information, "Information")
                        Me.dgvItem.Item(e.ColumnIndex, e.RowIndex).Value = Me.buffer
                        Exit Sub
                    End If
                    row.Cells(Me.ColToleranceDownPerc.Index).Value = ObjCD.DigitGrouping(CDec(row.Cells(Me.ColToleranceDown.Index).Value) / CDec(row.Cells(Me.ColQtyMulti.Index).Value) * 100, 2)
                End If
                If row.Cells(Me.ColToleranceUpFromPerc.Index).Value = "1" Then
                    row.Cells(Me.ColToleranceUp.Index).Value = ObjCD.DigitGrouping(CDec(row.Cells(Me.ColQtyMulti.Index).Value) * CDec(row.Cells(Me.ColToleranceUpPerc.Index).Value) / 100, DefPrecMulti)
                Else
                    If CDec(row.Cells(Me.ColToleranceUp.Index).Value) > (CDec(row.Cells(e.ColumnIndex).Value) / 2) Then
                        MsgBox("Qty you entered less than 50% of tolerance up. Change tolerance up first", MsgBoxStyle.Information, "Information")
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
                Me.ObjCD.DigitGrouping(Me.dgvItem(e.ColumnIndex, e.RowIndex).Value, Me.DefaultPrecision)
                'hitung ulang row
                Me.Hitung_Amount(row)
                'hitung total
                Me.Hitung_Total()
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
                'tandai di ubah dari diskon
                row.Cells(Me.ColDiscFromPerct.Index).Value = "1"
                'hitung ulang row
                Me.Hitung_Amount(row)
                'hitung total
                Me.Hitung_Total()
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
        If Me.dgvItem.Rows.Count = 0 Then
            Me.dgvItem.TabStop = False
            If Me.txtRemarks.Enabled = True Then
                Me.btnUP.Enabled = False
                Me.btnDown.Enabled = False
            End If
        End If
        Me.Hitung_Total()
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
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDown.Click
        Try
            If Me.dgvItem.SelectedRows.Count > 0 Then
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
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
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

    Private Sub TxtQuoNo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQuoNo.GotFocus
        TxtQuoNo.BackColor = BackColorAll
    End Sub

    Private Sub TxtQuoNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQuoNo.LostFocus
        TxtQuoNo.BackColor = Me.txtPrimaryCategory.BackColor
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

    Private Sub TxtShipVia_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtInquiryNo.GotFocus
        TxtInquiryNo.BackColor = BackColorAll
    End Sub

    Private Sub TxtShipVia_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtInquiryNo.LostFocus
        TxtInquiryNo.BackColor = Me.txtPrimaryCategory.BackColor
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
                TxtCodeNo.BackColor = Me.TxtQuoNo.BackColor
            End If
            If Me.txtPrimaryCustomer.Text <> "" Then
                If Me.txtPrimaryCustomer.Text = -1 And Me.txtPrimaryCustomer.Text <> "" Then
                    Me.TabControl1.SelectedIndex = 1
                    Me.txtCustomerCode.Focus()
                    MsgBox("To start entering item data, you specify customer data first.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            If Me.chkUseCustCodeNo.Checked = False Then
                If Me.txtPrimaryCurrency.Text <> "" Then
                    If Me.txtPrimaryCurrency.Text = -1 And Me.txtPrimaryCurrency.Text <> "" Then
                        Me.TabControl1.SelectedIndex = 0
                        Me.txtCurrencyCode.Focus()
                        MsgBox("To start entering item data, you specify currency data first.", MsgBoxStyle.Information, "Information")
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
            Me.TxtCodeNo.BackColor = Me.TxtQuoNo.BackColor
        End If
    End Sub

    'Private Sub TxtQtyperBatch_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchDefault.GotFocus
    '    TxtQtyperBatchDefault.BackColor = BackColorAll
    'End Sub

    'Private Sub TxtQtyperBatch_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchDefault.LostFocus
    '    TxtQtyperBatchDefault.BackColor = Me.txtPrimaryCategory.BackColor
    'End Sub

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
            Me.ObjCD.Number_Enter(Me.TxtDiscount, sender, e)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscount_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscount.GotFocus
        TxtDiscount.BackColor = BackColorAll
    End Sub

    Private Sub TxtDiscount_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscount.LostFocus
        TxtDiscount.BackColor = Me.txtPrimaryCategory.BackColor
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
        Dim hj As New Inventory_Module.FrmMasterKondisiStock
        Try
            If Me.TxtCodeNo.Text <> "" Then
                hj.txtCodeNo.Text = Me.TxtCodeNo.Text
            End If
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
                    Me.TabControl1.SelectedIndex = 1
                    Me.txtCustomerCode.Focus()
                    MsgBox("To start entering item data, you specify customer data first.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            If Me.chkUseCustCodeNo.Checked = True Then
                If Me.txtPrimaryCurrency.Text <> "" Then
                    If Me.txtPrimaryCurrency.Text = -1 And Me.txtPrimaryCurrency.Text <> "" Then
                        Me.TabControl1.SelectedIndex = 0
                        Me.txtCurrencyCode.Focus()
                        MsgBox("To start entering item data, you specify currency data first.", MsgBoxStyle.Information, "Information")
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
                    Me.SetGbo()
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
                        Me.SetGbo()
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
                        Me.SetGbo()
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
                            Me.SetGbo()
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
                    Me.TabControl1.SelectedIndex = 1
                    Me.txtCustomerCode.Focus()
                    MsgBox("To start entering item data, you specify customer data first.", MsgBoxStyle.Information, "Information")
                    Exit Sub
                End If
            End If
            If Me.txtPrimaryCurrency.Text <> "" Then
                If Me.txtPrimaryCurrency.Text = -1 And Me.txtPrimaryCurrency.Text <> "" Then
                    Me.TabControl1.SelectedIndex = 0
                    Me.txtCurrencyCode.Focus()
                    MsgBox("To start entering item data, you specify currency data first.", MsgBoxStyle.Information, "Information")
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
                Me.SetGbo()
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
            'row.Cells(Me.ColNettPrice.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(row.Cells(Me.ColUnitPrice.Index).Value * _
            '(100 - row.Cells(Me.ColDiscount.Index).Value) / 100, _
            'DefaultPrecision))
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
            'row.Cells(Me.ColAmount.Index).Value = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(row.Cells(Me.ColQtyMulti.Index).Value * _
            'row.Cells(Me.ColNettPrice.Index).Value, DefaultPrecision))
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
            Me.txtNettPrice.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDbl(Me.TxtUnitPrice.Text) * (100 - CDbl(Me.TxtDiscount.Text)) / 100, _
            DefaultPrecision))
            'hitung amount
            If Me.TxtQtyperBatchMulti.Text = "" Then
                Me.TxtQtyperBatchMulti.Text = "0"
            End If
            If Me.chkUseSecondary.Checked = False Then
                If Me.txtPriceByMultiUnit.Text = "" Then
                    Me.txtPriceByMultiUnit.Text = "1"
                End If
                'Me.txtAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(CDbl(Me.txtNettPrice.Text) * _
                'CDbl(Me.TxtQtyperBatchMulti.Text), DefaultPrecision))
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
        Dim SubTotal As Double = 0

        Try
            'hitung sub total
            For Each row In Me.dgvItem.Rows
                'yang diperhitungkan hanya yang tidak pake secondary unit
                If row.Cells(Me.ColUseSecondary.Index).Value = Boolean.FalseString Then
                    SubTotal += CDbl(row.Cells(Me.ColAmount.Index).Value)
                End If
            Next
            Me.txtSubTotal.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(SubTotal, DefaultPrecision))
            'hitung pajak
            If Me.chkTax.Checked = True Then
                Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Me.txtTax.Text * SubTotal / 100, _
                DefaultPrecision))
            Else
                Me.txtTax.Text = "0"
                Me.txtTaxAmount.Text = "0"
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
                If Me.TxtDPPercent.Text.Trim = "" Then
                    Me.TxtDPPercent.Text = "0"
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

    Private Sub llbAuthorizationDetail_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbAuthorizationDetail.LinkClicked
        Dim i As New FrmTransaksiQuotationAuthorizationDetail
        Dim jmlbaris As Integer = 0
        Dim row As DataGridViewRow
        Try
            i.dgvDetail.Rows.Clear()
            'tambahkan 3 atau 2 baris authorization
            If Me.objQuotation.AuthorizationLevelNeeded = 0 Then
                'untuk yang baru, anggap 2 level
                jmlbaris = 2
            Else
                jmlbaris = Me.objQuotation.AuthorizationLevelNeeded
            End If
            'masukkan baris dan detail data otorisasi
            For j As Integer = 1 To jmlbaris
                i.dgvDetail.Rows.Add()
                row = i.dgvDetail.Rows(i.dgvDetail.Rows.Count - 1)
                If row.Index = 0 Then
                    row.Cells(Me.ColCodeNo.Index).Value = "Sales Staff Authorization"
                    If Me.objQuotation.SalesStaffAuthorizationPrimary <> -1 Then
                        row.Cells(Me.ColDescription.Index).Value = Me.ObjUser.GetNamaByPrimary(Me.objQuotation.SalesStaffAuthorizationPrimary)
                        row.Cells(Me.ColSupplierCodeNo.Index).Value = Format(Me.objQuotation.SalesStaffAuthorizationDate, "dd MMM yyyy - HH:mm")
                    Else
                        row.Cells(Me.ColDescription.Index).Value = "-"
                        row.Cells(Me.ColSupplierCodeNo.Index).Value = "-"
                    End If
                ElseIf row.Index = 1 Then
                    row.Cells(Me.ColCodeNo.Index).Value = "Sales Manager Authorization"
                    If Me.objQuotation.SalesManagerAuthorizationPrimary <> -1 Then
                        row.Cells(Me.ColDescription.Index).Value = Me.ObjUser.GetNamaByPrimary(Me.objQuotation.SalesManagerAuthorizationPrimary)
                        row.Cells(Me.ColSupplierCodeNo.Index).Value = Format(Me.objQuotation.SalesManagerAuthorizationDate, "dd MMM yyyy - HH:mm")
                    Else
                        row.Cells(Me.ColDescription.Index).Value = "-"
                        row.Cells(Me.ColSupplierCodeNo.Index).Value = "-"
                    End If
                ElseIf row.Index = 2 Then
                    row.Cells(Me.ColCodeNo.Index).Value = "General Manager Authorization"
                    If Me.objQuotation.GeneralManagerAuthorizationPrimary <> -1 Then
                        row.Cells(Me.ColDescription.Index).Value = Me.ObjUser.GetNamaByPrimary(Me.objQuotation.GeneralManagerAuthorizationPrimary)
                        row.Cells(Me.ColSupplierCodeNo.Index).Value = Format(Me.objQuotation.GeneralManagerAuthorizationDate, "dd MMM yyyy - HH:mm")
                    Else
                        row.Cells(Me.ColDescription.Index).Value = "-"
                        row.Cells(Me.ColSupplierCodeNo.Index).Value = "-"
                    End If
                End If
            Next
            'tunjukkan form
            Me.ObjCD.SetNumber(i.dgvDetail)
            i.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i = Nothing
        End Try
    End Sub

    Private Sub txtCurrencyCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCurrencyCode.KeyDown
        Dim i As New FrmCariMasterCurrency
        i.BackColorAll = Me.BackColorAll

        Try
            If e.KeyValue = 13 Then
                If ObjCurrency.ApakahSudahAda(Me.txtCurrencyCode.Text) = True Then
                    Me.txtPrimaryCurrency.Text = ObjCurrency.GetPrimary(Me.txtCurrencyCode.Text)
                    Me.SetDefaultPrecision(True)
                    Me.chkTax.Focus()
                Else
                    If Me.txtCurrencyCode.Text <> "" Then
                        i.txtCari.Text = Me.txtCurrencyCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCurrency > 0 Then
                        Me.txtPrimaryCurrency.Text = i.PrimaryCurrency
                        Me.txtCurrencyCode.Text = ObjCurrency.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                        Me.SetDefaultPrecision(True)
                        Me.chkTax.Focus()
                    Else
                        Me.txtCurrencyCode.Clear()
                        Me.txtPrimaryCurrency.Text = "-1"
                        Me.txtCurrencyCode.Focus()
                        Me.SetDefaultPrecision(True)
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

        Try
            If Me.txtCurrencyCode.Text <> "" Then
                If ObjCurrency.ApakahSudahAda(Me.txtCurrencyCode.Text) = True Then
                    Me.txtPrimaryCurrency.Text = ObjCurrency.GetPrimary(Me.txtCurrencyCode.Text)
                    Me.SetDefaultPrecision(True)
                Else
                    If Me.txtCurrencyCode.Text <> "" Then
                        i.txtCari.Text = Me.txtCurrencyCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCurrency > 0 Then
                        Me.txtPrimaryCurrency.Text = i.PrimaryCurrency
                        Me.txtCurrencyCode.Text = ObjCurrency.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                        Me.SetDefaultPrecision(True)
                        Me.chkTax.Focus()
                    Else
                        Me.txtCurrencyCode.Clear()
                        Me.txtPrimaryCurrency.Text = "-1"
                        Me.txtCurrencyCode.Focus()
                        Me.SetDefaultPrecision(True)
                    End If
                End If
            Else
                Me.txtPrimaryCurrency.Text = "-1"
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

        Try
            If Me.ObjCurrency.ApakahSudahAda(Me.txtCurrencyCode.Text) = True Then
                Me.chkTax.Focus()
                Me.txtPrimaryCurrency.Text = ObjCurrency.GetPrimary(Me.txtCurrencyCode.Text)
                Me.SetDefaultPrecision(True)
            Else
                i.cboCriteria.SelectedIndex = 0
                If Me.txtCurrencyCode.Text <> "" Then
                    i.txtCari.Text = Me.txtCurrencyCode.Text
                End If

                i.ShowDialog()
                If i.PrimaryCurrency > 0 Then
                    Me.txtPrimaryCurrency.Text = i.PrimaryCurrency
                    Me.txtCurrencyCode.Text = ObjCurrency.GetCurrencyCode(Me.txtPrimaryCurrency.Text)
                    Me.SetDefaultPrecision(True)
                    Me.chkTax.Focus()
                Else
                    Me.txtCurrencyCode.Clear()
                    Me.txtCurrencyCode.Focus()
                    Me.SetDefaultPrecision(True)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub SetDefaultPrecision(ByVal vRecalculateItemAmount As Boolean)
        Try
            Me.DefaultPrecision = Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text)
            Me.FormatDGV()
            If vRecalculateItemAmount = True Then
                For Each rowDGV As DataGridViewRow In Me.dgvItem.Rows
                    Me.Hitung_Amount(rowDGV)
                Next
            End If
            Me.Hitung_Total()
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
                Me.txtTax.TabStop = True

                Dummy = Me.ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Invoice, Core_Module.ClsTransaksiSetting.EnumNamaParameter.TaxPercentage)
                If Dummy = "" Then
                    Dummy = "0"
                End If
                Me.txtTax.Text = Me.ObjCD.DigitGrouping(Dummy, 2)
            Else
                Me.txtTax.ReadOnly = True
                Me.txtTax.TabStop = False
                Me.txtTax.Text = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtCustomerCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCustomerCode.KeyDown
        Dim i As New FrmCariMasterCustomer
        i.BackColorAll = Me.BackColorAll
        Try
            If e.KeyValue = 13 Then
                If Me.ObjCustomer.ApakahSudahAda(Me.txtCustomerCode.Text) Then
                    Me.txtPrimaryCustomer.Text = Me.ObjCustomer.GetPrimary(Me.txtCustomerCode.Text)
                    Me.TxtCustomerName.Text = Me.ObjCustomer.GetCustName(Me.txtPrimaryCustomer.Text)
                    Me.ObjCustomer.IsiCboAddress(Me.cboCustomerAddress, Me.txtPrimaryCustomer.Text)
                    Me.ObjCustomer.IsiCboAddress(Me.CboAddressShipTo, Me.txtPrimaryCustomer.Text)

                    Me.cboCustomerAddress.SelectedIndex = 0
                    Me.CboAddressShipTo.SelectedIndex = 0
                    Me.cboCustomerAddress.Focus()
                    Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWP(Me.txtPrimaryCustomer.Text)
                    Me.txtContactPerson.Text = Me.ObjCustomer.GetCustContactPerson(Me.txtPrimaryCustomer.Text)
                Else
                    'ada 2 Customer dengan kode yang sama atau tidak ada Customer yang dimaksud
                    If Me.txtCustomerCode.Text <> "" Then
                        i.txtCari.Text = Me.txtCustomerCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCustomer > 0 Then
                        Me.txtPrimaryCustomer.Text = i.PrimaryCustomer
                        Me.txtCustomerCode.Text = Me.ObjCustomer.GetCustomerCode(Me.txtPrimaryCustomer.Text)
                        Me.TxtCustomerName.Text = Me.ObjCustomer.GetCustName(Me.txtPrimaryCustomer.Text)
                        Me.ObjCustomer.IsiCboAddress(Me.cboCustomerAddress, Me.txtPrimaryCustomer.Text)
                        Me.ObjCustomer.IsiCboAddress(Me.CboAddressShipTo, Me.txtPrimaryCustomer.Text)
                        Me.cboCustomerAddress.SelectedIndex = 0
                        Me.CboAddressShipTo.SelectedIndex = 0
                        Me.cboCustomerAddress.Focus()
                        Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWP(Me.txtPrimaryCustomer.Text)
                        Me.txtContactPerson.Text = Me.ObjCustomer.GetCustContactPerson(Me.txtPrimaryCustomer.Text)
                    Else
                        Me.txtCustomerCode.Clear()
                        Me.TxtCustomerName.Clear()
                        Me.cboCustomerAddress.Items.Clear()
                        Me.CboAddressShipTo.Items.Clear()

                        Me.txtCustomerAddIdentifier.Text = ""
                        Me.TxtIdentifierAddressShipTo.Text = ""

                        Me.txtPrimaryCustomer.Text = "-1"
                        Me.txtNoNPWP.Clear()
                        Me.txtContactPerson.Clear()
                        Me.txtPhoneNo.Clear()
                        Me.txtCustomerCode.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub txtCustomerCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustomerCode.Leave
        Dim i As New FrmCariMasterCustomer
        i.BackColorAll = Me.BackColorAll
        Try
            If Me.txtCustomerCode.Text <> "" Then
                If Me.ObjCustomer.ApakahSudahAda(Me.txtCustomerCode.Text) = True Then
                    Me.txtPrimaryCustomer.Text = Me.ObjCustomer.GetPrimary(Me.txtCustomerCode.Text)
                    Me.TxtCustomerName.Text = Me.ObjCustomer.GetCustName(Me.txtPrimaryCustomer.Text)
                    Me.ObjCustomer.IsiCboAddress(Me.cboCustomerAddress, Me.txtPrimaryCustomer.Text)
                    Me.cboCustomerAddress.SelectedIndex = 0
                    If Me.RboAddressShipTo.Checked = True Then
                        Me.ObjCustomer.IsiCboAddress(Me.CboAddressShipTo, Me.txtPrimaryCustomer.Text)
                        Me.CboAddressShipTo.SelectedIndex = 0
                    End If

                    Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWP(Me.txtPrimaryCustomer.Text)
                    Me.txtContactPerson.Text = Me.ObjCustomer.GetCustContactPerson(Me.txtPrimaryCustomer.Text)
                Else
                    If Me.txtCustomerCode.Text <> "" Then
                        i.txtCari.Text = Me.txtCustomerCode.Text
                    End If

                    i.ShowDialog()
                    If i.PrimaryCustomer > 0 Then
                        Me.txtPrimaryCustomer.Text = i.PrimaryCustomer
                        Me.txtCustomerCode.Text = Me.ObjCustomer.GetCustomerCode(Me.txtPrimaryCustomer.Text)
                        Me.TxtCustomerName.Text = Me.ObjCustomer.GetCustName(Me.txtPrimaryCustomer.Text)
                        Me.ObjCustomer.IsiCboAddress(Me.cboCustomerAddress, Me.txtPrimaryCustomer.Text)
                        Me.cboCustomerAddress.SelectedIndex = 0
                        If Me.RboAddressShipTo.Checked = True Then
                            Me.ObjCustomer.IsiCboAddress(Me.CboAddressShipTo, Me.txtPrimaryCustomer.Text)
                            Me.CboAddressShipTo.SelectedIndex = 0
                        End If
                        
                        Me.cboCustomerAddress.Focus()
                        Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWP(Me.txtPrimaryCustomer.Text)
                        Me.txtContactPerson.Text = Me.ObjCustomer.GetCustContactPerson(Me.txtPrimaryCustomer.Text)
                    Else
                        Me.txtCustomerCode.Clear()
                        Me.TxtCustomerName.Clear()
                        Me.cboCustomerAddress.Items.Clear()
                        Me.CboAddressShipTo.Items.Clear()
                        Me.txtCustomerAddIdentifier.Text = ""
                        Me.TxtIdentifierAddressShipTo.Text = ""
                        Me.txtPrimaryCustomer.Text = "-1"
                        Me.txtNoNPWP.Clear()
                        Me.txtContactPerson.Clear()
                        Me.txtPhoneNo.Clear()
                        Me.txtCustomerCode.Focus()
                    End If
                End If
            Else
                Me.txtCustomerCode.Clear()
                Me.TxtCustomerName.Clear()
                Me.cboCustomerAddress.Items.Clear()
                Me.CboAddressShipTo.Items.Clear()
                Me.txtCustomerAddIdentifier.Text = ""
                Me.TxtIdentifierAddressShipTo.Text = ""
                Me.txtPrimaryCustomer.Text = "-1"
                Me.txtNoNPWP.Clear()
                Me.txtContactPerson.Clear()
                Me.txtPhoneNo.Clear()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub btnFindCustomer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindCustomer.Click
        Dim i As New FrmCariMasterCustomer
        i.BackColorAll = Me.BackColorAll

        Try
            i.cboCriteria.SelectedIndex = 0
            If Me.txtCustomerCode.Text <> "" Then
                i.txtCari.Text = Me.txtCustomerCode.Text
            End If

            i.ShowDialog()
            If i.PrimaryCustomer > 0 Then
                Me.txtPrimaryCustomer.Text = i.PrimaryCustomer
                Me.txtCustomerCode.Text = Me.ObjCustomer.GetCustomerCode(Me.txtPrimaryCustomer.Text)
                Me.TxtCustomerName.Text = Me.ObjCustomer.GetCustName(Me.txtPrimaryCustomer.Text)
                Me.ObjCustomer.IsiCboAddress(Me.cboCustomerAddress, Me.txtPrimaryCustomer.Text)
                Me.ObjCustomer.IsiCboAddress(Me.CboAddressShipTo, Me.txtPrimaryCustomer.Text)
                Me.cboCustomerAddress.SelectedIndex = 0
                Me.CboAddressShipTo.SelectedIndex = 0
                Me.cboCustomerAddress.Focus()
                Me.txtNoNPWP.Text = Me.ObjCustomer.GetNPWP(Me.txtPrimaryCustomer.Text)
                Me.txtContactPerson.Text = Me.ObjCustomer.GetCustContactPerson(Me.txtPrimaryCustomer.Text)
            Else
                Me.txtCustomerCode.Clear()
                Me.TxtCustomerName.Clear()
                Me.cboCustomerAddress.Items.Clear()
                Me.CboAddressShipTo.Items.Clear()
                Me.txtCustomerAddIdentifier.Text = ""
                Me.TxtIdentifierAddressShipTo.Text = ""
                Me.txtPrimaryCustomer.Text = "-1"
                Me.txtNoNPWP.Clear()
                Me.txtContactPerson.Clear()
                Me.txtPhoneNo.Clear()
                Me.txtCustomerCode.Focus()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            i.Dispose()
        End Try
    End Sub

    Private Sub cboCustomerAddress_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboCustomerAddress.SelectedIndexChanged
        Try
            If Me.txtRemarks.Enabled = True Then
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
                    Me.TxtCodeNo.BackColor = Me.TxtQuoNo.BackColor
                    Me.TxtCodeNo.TabStop = False
                    Me.BtnFindCodeNo.Enabled = False
                    Me.btnFindCustomerProductCode.Enabled = True
                    Me.ColSuppDesc.Visible = True
                    Me.ColSupplierCodeNo.Visible = True
                    Me.txtCustomerProductCodeNo.Focus()
                Else
                    Me.txtCustomerProductCodeNo.ReadOnly = True
                    Me.txtCustomerProductCodeNo.BackColor = Me.TxtQuoNo.BackColor
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

    Private Sub txtTax_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTax.TextChanged
        Try
            If Me.txtRemarks.Enabled = True And IsNumeric(Me.txtTax.Text) = True And _
            IsNumeric(Me.txtSubTotal.Text) Then
                If CDbl(Me.txtTax.Text) <= 100 Then
                    'hitung pajak
                    If Me.chkTax.Checked = True Then
                        Me.txtTaxAmount.Text = Me.ObjCD.DigitGrouping(Me.ObjCD.SetDecimal(Me.txtTax.Text * CDbl(Me.txtSubTotal.Text) / 100, _
                        DefaultPrecision))
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

    Private Sub TxtUnitPrice_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtUnitPrice.TextChanged
        Try
            If IsNumeric(Me.TxtQtyperBatchMulti.Text) = True And IsNumeric(Me.TxtDiscount.Text) = _
            True And IsNumeric(Me.TxtUnitPrice.Text) = True Then
                If Me.TxtUnitPrice.Text <> 0 Then
                    Me.Hitung_Amount()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

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

    Private Sub TxtDiscount_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TxtDiscount.TextChanged
        'Try
        '    If IsNumeric(Me.TxtQtyperBatchMulti.Text) = True And IsNumeric(Me.TxtDiscount.Text) = _
        '    True And IsNumeric(Me.TxtUnitPrice.Text) = True Then
        '        If CDbl(Me.TxtDiscount.Text) <= 100 Then
        '            Me.Hitung_Amount()
        '        End If
        '    End If
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'End Try
    End Sub

    Private Sub TxtCustomerName_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCustomerName.TextChanged
        Try
            If Me.txtPrimaryCustomer.Text <> -1 And Me.txtPrimaryCustomer.Text <> "" Then
                If Me.TxtCustomerName.Text <> "" Then
                    Me.Text = "Quotation: " + Me.txtCustomerCode.Text + " - " + Me.TxtCustomerName.Text
                Else
                    Me.Text = "Quotation "
                End If
            Else
                Me.Text = "Quotation"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTermsNCondition_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTermsNCondition.GotFocus
        Try
            If Me.txtTermsNCondition.Enabled = True And Me.txtTermsNCondition.ReadOnly = False Then
                Me.txtTermsNCondition.BackColor = Me.BackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtTermsNCondition_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTermsNCondition.LostFocus
        Try
            If Me.txtTermsNCondition.Enabled = True And Me.txtTermsNCondition.ReadOnly = False Then
                Me.txtTermsNCondition.BackColor = Me.TxtPrimaryCodeNo.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtNoNPWP_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNoNPWP.GotFocus
        Try
            If Me.txtNoNPWP.Enabled = True And Me.txtNoNPWP.ReadOnly = False Then
                Me.txtNoNPWP.BackColor = Me.BackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub txtNoNPWP_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNoNPWP.LostFocus
        Try
            If Me.txtNoNPWP.Enabled = True And Me.txtNoNPWP.ReadOnly = False Then
                Me.txtNoNPWP.BackColor = Me.TxtPrimaryCodeNo.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbAddress_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbAddress.LinkClicked
        Try
            If Me.txtRemarks.Enabled = True OrElse (IsNumeric(Me.TxtPrimaryQuotation.Text) AndAlso CDec(Me.TxtPrimaryQuotation.Text) > 0) Then
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

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim ij As New FrmCetakQuotation
        Dim obj As New ClsTransaksiQuotation
        Try
            If Me.TxtQuoNo.Text = "" Then
                Me.btnFind.PerformClick()
            End If

            If Me.TxtPrimaryQuotation.Text <> "" Then
                'cek apakah quotation yang akan diprint sudah do otorisaisi
                If Me.objQuotation.ApakahSudahOtorisasi(Me.TxtPrimaryQuotation.Text) = True Then
                    obj.CetakQuotation(ij.crv1, Me.TxtPrimaryQuotation.Text)
                Else
                    MsgBox("Please authorize this quotation before printing.", _
                    MsgBoxStyle.Information, "Information")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            obj = Nothing
            ij.Dispose()
        End Try
    End Sub

    Private Sub dtpDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpDate.ValueChanged
        Try
            If Me.txtRemarks.Enabled = True Then
                If Me.baru = False Then
                    If Me.objQuotation.TanggalTransaksi.Month = Me.dtpDate.Value.Month And Me.objQuotation.TanggalTransaksi.Year = Me.dtpDate.Value.Year Then
                        'jika kembali ke bulan yang bersangkutan maka no kembali ke awal
                        Me.TxtQuoNo.Text = Me.objQuotation.NoTransaksi
                    Else
                        'jika pindah bulan maka akan ganti no
                        Me.TxtQuoNo.Text = Me.objQuotation.GetNoBaru(Me.dtpDate.Value)
                    End If
                Else
                    Me.TxtQuoNo.Text = Me.objQuotation.GetNoBaru(Me.dtpDate.Value)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtNamaSigned_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNamaSigned.GotFocus
        Try
            If Me.txtRemarks.Enabled = True And Me.TxtNamaSigned.Enabled = True Then
                Me.TxtNamaSigned.BackColor = Me.BackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtJabatanSigned_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtJabatanSigned.LostFocus
        Try
            If Me.TxtJabatanSigned.Enabled = True Or Me.txtRemarks.Enabled = True Then
                Me.TxtJabatanSigned.BackColor = Me.TxtPrimaryQuotation.BackColor
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtJabatanSigned_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtJabatanSigned.GotFocus
        Try
            If Me.txtRemarks.Enabled = True And Me.TxtJabatanSigned.Enabled = True Then
                Me.TxtJabatanSigned.BackColor = Me.BackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtNamaSigned_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtNamaSigned.LostFocus
        Try
            If Me.TxtNamaSigned.Enabled = True Or Me.txtRemarks.Enabled = True Then
                Me.TxtNamaSigned.BackColor = Me.TxtPrimaryQuotation.BackColor
            End If
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
    '                Me.TxtQtyperBatchDefault.Text = Me.ObjCD.DigitGrouping(Me.TxtQtyperBatchDefault.Text, Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text))
    '            End If
    '            If IsNumeric(Me.TxtQtyperBatchDefault.Text) = True And IsNumeric(Me.TxtDiscount.Text) = _
    '            True And IsNumeric(Me.TxtUnitPrice.Text) = True Then
    '                Me.ObjProduct.GetUnitPriceNDiskon(Me.txtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtQtyperBatchDefault.Text, _
    '                            Me.txtPrimaryCurrency.Text, Me.txtPrimaryUnitsDefault.Text, UnitPrice, DIskon)

    '                Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(UnitPrice, _
    '                Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text))
    '                Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Diskon, 2)

    '                Me.Hitung_Amount()
    '            End If
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Private Sub TxtUnitPrice_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtUnitPrice.Leave
        Try
            Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(Me.TxtUnitPrice.Text, Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text))
            'unit price harus selalu lebih besar dari disc amount
            'hanya berlaku kalo ngitung disc berdasarkan disc amount, klo berdasarkan perc, langsung hitung ulang disc amountnya
            If Me.TxtDiscFromPerct.Text = "0" Then
                If CDec(Me.TxtDiscAmount.Text) > CDec(Me.TxtUnitPrice.Text) Then
                    Me.TxtUnitPrice.Text = Me.buffer
                    Exit Sub
                End If
            End If
            If IsNumeric(Me.TxtQtyperBatchMulti.Text) = True And IsNumeric(Me.TxtDiscount.Text) = _
            True And IsNumeric(Me.TxtUnitPrice.Text) = True Then
                Me.Hitung_Amount()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtDiscount_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscount.Leave
        Try
            'Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Me.TxtDiscount.Text, 2)
            Me.ObjCD.Number_Leave(Me.TxtDiscount, sender, e, 2, Me.ObjCD)
            If CDec(Me.TxtUnitPrice.Text) = 0 Then
                Me.TxtDiscount.Text = 0
                Me.ObjCD.Number_Leave(Me.TxtDiscount, sender, e, 2, Me.ObjCD)
            End If
            If IsNumeric(Me.TxtQtyperBatchMulti.Text) = True And IsNumeric(Me.TxtDiscount.Text) = _
            True And IsNumeric(Me.TxtUnitPrice.Text) = True Then
                If CDbl(Me.TxtDiscount.Text) <= 100 Then
                    Me.TxtDiscFromPerct.Text = "1"
                    Me.Hitung_Amount()
                End If
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
                ijk.TxtCustomerName.Text = Me.TxtCustomerName.Text
                ijk.ShowDialog()

                ijk = Nothing
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'aul 
    Private Sub TxtDiscAmount_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtDiscAmount.Enter
        Try
            Me.buffer = Me.TxtDiscAmount.Text
            Me.ObjCD.Number_Enter(Me.TxtDiscAmount, sender, e)
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
                    Me.TxtQtyperBatchMulti.Text = Me.ObjCD.DigitGrouping(Me.TxtQtyperBatchMulti.Text, Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text))
                End If
                If IsNumeric(Me.TxtQtyperBatchMulti.Text) = True And IsNumeric(Me.TxtDiscount.Text) = True And IsNumeric(Me.TxtUnitPrice.Text) = True Then
                    Me.ObjProduct.GetUnitPriceNDiskon(Me.txtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtQtyperBatchMulti.Text, _
                    Me.txtPrimaryCurrency.Text, Me.txtPrimaryUnitsMulti.Text, UnitPrice, DIskon)
                    Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(UnitPrice, _
                    Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text))
                    Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(DIskon, 2)
                    Me.SetQtyDefault()
                    Me.Hitung_Amount()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtQtyperBatchMulti_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtQtyperBatchMulti.LostFocus
        TxtQtyperBatchMulti.BackColor = Me.txtPrimaryCategory.BackColor
    End Sub

    Private Sub TxtPrimaryCodeNo_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPrimaryCodeNo.TextChanged
        Try
            Me.PrepareMultiUnit()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub PrepareMultiUnit()
        Dim Dta As New DseTransaksiQuotationTableAdapters.tblmaster_product_item_multiunitTableAdapter
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
            ByVal vTbl As DseTransaksiQuotation.CboUnitDataTable, _
            ByVal vTblMultiUnit As DseTransaksiQuotation.tblmaster_product_item_multiunitDataTable) As Boolean
        Dim rowTbl As DseTransaksiQuotation.CboUnitRow
        Dim rowMultiUnit As DseTransaksiQuotation.tblmaster_product_item_multiunitRow
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

    Private Sub CboUnit_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboUnit.SelectedIndexChanged
        Dim UnitPrice, Discount As Decimal
        Dim row As DseTransaksiQuotation.CboUnitRow
        Try
            If Me.CboUnit.SelectedIndex <> -1 Then
                row = Me.tblCboUnit.DefaultView(Me.CboUnit.SelectedIndex).Row
                Me.txtPrimaryUnitsMulti.Text = row.Primary
                Me.lblUnitsMulti.Text = Me.CboUnit.Text
                Me.LblUnitPrice.Text = "Unit Price/" & Me.CboUnit.Text
                Me.TxtQtyperBatchMulti.Text = ObjCD.DigitGrouping(Me.TxtQtyperBatchMulti.Text, Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text))
                Me.ObjProduct.GetUnitPriceNDiskon(Me.txtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtQtyperBatchMulti.Text, _
                Me.txtPrimaryCurrency.Text, Me.txtPrimaryUnitsMulti.Text, UnitPrice, Discount)
                Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(UnitPrice, Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text))
                Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Discount, 2)
                Me.SetQtyDefault()
                Me.SetCboPriceByMultiUnit()
                Me.Hitung_Amount()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetQtyDefault()
        Dim RatioDefault As Decimal = 0
        Dim RatioReference As Decimal = 0
        Try
            RatioDefault = Me.objQuotation.GetRatioDefault(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text)
            RatioReference = Me.objQuotation.GetRatioReference(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text)
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

    Private Sub EndEditToleranceDown(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, _
    ByVal row As DataGridViewRow)
        Dim DefPrecMulti As Integer = ObjProduct.GetAngkaDecimal(row.Cells(Me.ColPrimaryCodeNo.Index).Value, row.Cells(Me.ColPrimaryUnitsMulti.Index).Value)
        Try
            With row
                If Me.ObjCD.CekNonNegatif(.Cells(e.ColumnIndex).Value) = False Then
                    MsgBox("Please enter positive number for tolerance.", MsgBoxStyle.Information, "Information")
                    .Cells(e.ColumnIndex).Value = Me.buffer
                    Exit Sub
                End If
                If .Cells(e.ColumnIndex).Value > (.Cells(Me.ColQtyMulti.Index).Value / 2) Then
                    MsgBox("Tolerance cannot greater than 50% of qty.", MsgBoxStyle.Information, "Information")
                    .Cells(e.ColumnIndex).Value = Me.buffer
                    Exit Sub
                End If
                .Cells(e.ColumnIndex).Value = Me.ObjCD.DigitGrouping(.Cells(e.ColumnIndex).Value, DefPrecMulti)
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
                .Cells(e.ColumnIndex).Value = ObjCD.DigitGrouping(.Cells(e.ColumnIndex).Value, 2)
                .Cells(Me.ColToleranceDownFromPerc.Index).Value = "1"
                .Cells(Me.ColToleranceDown.Index).Value = ObjCD.DigitGrouping(.Cells(Me.ColToleranceDownPerc.Index).Value / 100 * .Cells(Me.ColQtyMulti.Index).Value, DefPrecMulti)
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
                If .Cells(e.ColumnIndex).Value > (.Cells(Me.ColQtyMulti.Index).Value / 2) Then
                    MsgBox("Tolerance cannot greater than 50% of qty.", MsgBoxStyle.Information, "Information")
                    .Cells(e.ColumnIndex).Value = Me.buffer
                    Exit Sub
                End If
                .Cells(e.ColumnIndex).Value = Me.ObjCD.DigitGrouping(.Cells(e.ColumnIndex).Value, DefPrecMulti)
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
                .Cells(e.ColumnIndex).Value = ObjCD.DigitGrouping(.Cells(e.ColumnIndex).Value, 2)
                .Cells(Me.ColToleranceUpFromPerc.Index).Value = "1"
                .Cells(Me.ColToleranceUp.Index).Value = ObjCD.DigitGrouping(.Cells(Me.ColToleranceUpPerc.Index).Value / 100 * .Cells(Me.ColQtyMulti.Index).Value, DefPrecMulti)
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub llbUpdateAllDeliveryDate_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llbUpdateAllDeliveryDate.LinkClicked
        Dim f As New FrmTransaksiSO_SetDeliveryDate
        Dim rowD As DseTransaksiQuotation.tbltrans_quotation_itemdeliveryRow

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
                        rowD = Me.vTBLItemDelivery.Newtbltrans_quotation_itemdeliveryRow
                        With rowD
                            .delivery_date = Format(f.DtpDate.Value.Date, "dd MMM yyyy")
                            .delivery_priority = Me.ObjDeliveryPriority.GetPrimaryDefault
                            .delivery_qty = r.Cells(Me.ColQty.Index).Value
                            .identifier = r.Cells(Me.ColIdentifier.Index).Value
                            .primary = Me.GetPrimaryMax(Me.vTBLItemDelivery) + 1
                            .primary_transaksi = Me.TxtPrimaryQuotation.Text
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

    Private Sub dgvItem_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvItem.UserDeletingRow
        Try
            If e.Row.Index = Me.dgvItem.RowCount - 1 Then
                Me.EndEditDeliveryDateCont2(True)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
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

                        Me.TxtUnitPrice.Text = Me.ObjCD.DigitGrouping(UnitPrice, Me.ObjCurrency.GetDefaultPrecission(Me.txtPrimaryCurrency.Text))
                        Me.TxtDiscount.Text = Me.ObjCD.DigitGrouping(Discount, 2)
                    Else
                        Me.txtPriceByMultiUnit.Text = "0"

                        Me.ObjProduct.GetUnitPriceNDiskon(Me.txtPrimaryCustomer.Text, Me.TxtPrimaryCodeNo.Text, Me.TxtQtyperBatchDefault.Text, _
                        Me.txtPrimaryCurrency.Text, Me.txtPrimaryUnitsDefault.Text, UnitPrice, Discount)

                     
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

    Private Sub RboAddressShipTo_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RboAddressShipTo.CheckedChanged
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

    Private Sub RboCustomerShipTo_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RboCustomerShipTo.CheckedChanged
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

    Private Sub CboAddressShipTo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboAddressShipTo.SelectedIndexChanged
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

    Private Sub LlbViewAddressShipTo_LinkClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbViewAddressShipTo.LinkClicked
        Try
            If Me.txtRemarks.Enabled = True OrElse (IsNumeric(Me.TxtPrimaryQuotation.Text) AndAlso CDec(Me.TxtPrimaryQuotation.Text) > 0) Then
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

    Private Sub TxtCustomerCodeShipTo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCustomerCodeShipTo.GotFocus
        TxtCustomerCodeShipTo.BackColor = BackColorAll
    End Sub

    Private Sub TxtCustomerCodeShipTo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCustomerCodeShipTo.KeyDown
        Try
            If e.KeyValue = 13 Then
                If Me.ObjCustomer.ApakahSudahAda(Me.TxtCustomerCodeShipTo.Text) = True Then
                    Me.TxtPrimaryCustomerShipTo.Text = Me.ObjCustomer.GetPrimary(Me.TxtCustomerCodeShipTo.Text)
                    Me.TxtCustomerDescShipTo.Text = Me.ObjCustomer.GetCustName(Me.TxtPrimaryCustomerShipTo.Text)
                    Me.ObjCustomer.IsiCboAddress(Me.CboAddressCustomerShipTo, Me.TxtPrimaryCustomerShipTo.Text)
                    Me.CboAddressCustomerShipTo.SelectedIndex = 0
                    Me.txtTermsNCondition.Focus()
                Else
                    Me.BtnFindCustomerShipTo.PerformClick()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtCustomerCodeShipTo_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCustomerCodeShipTo.Leave
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

    Private Sub TxtCustomerCodeShipTo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCustomerCodeShipTo.LostFocus
        TxtCustomerCodeShipTo.BackColor = Me.txtPrimaryCategory.BackColor
    End Sub

    Private Sub BtnFindCustomerShipTo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnFindCustomerShipTo.Click
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

    Private Sub CboAddressCustomerShipTo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CboAddressCustomerShipTo.SelectedIndexChanged
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

    Private Sub LlbViewAddressCustomerShipTo_LinkClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LlbViewAddressCustomerShipTo.LinkClicked
        Try
            If Me.txtRemarks.Enabled = True OrElse (IsNumeric(Me.TxtPrimaryQuotation.Text) AndAlso CDec(Me.TxtPrimaryQuotation.Text) > 0) Then
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
End Class